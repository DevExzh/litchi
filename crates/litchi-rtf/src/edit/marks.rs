//! Field instructions, bookmarks, paragraph tab stops and section page
//! setup by source splicing (record 0328).
//!
//! Each staged mark names an entry of the source's native collections
//! (`fields()`, `bookmarks()`, a body paragraph, `sections()`). The commit
//! tokenizes the exact source once, finds the control words or destination
//! groups that spell the entry, and splices only those bytes: a field's
//! `\fldinst` content, a bookmark's `\bkmkstart`/`\bkmkend` groups, a
//! paragraph's `\tx`/`\tb` controls with their alignment and leader
//! controls, a section's `\pgwsxn`-family controls. Every other byte keeps
//! its source spelling.
//!
//! RTF state is positional, so the candidate is reparsed with the source's
//! open policy and must read back exactly as expected: the edited entries
//! changed as staged, and the body text, every paragraph's properties and
//! runs, every table, field, bookmark and section, and the opaque content
//! unchanged otherwise. Anything else is a typed refusal.

use std::ops::Range;

use bumpalo::Bump;

use super::{Change, Commit, Error, Package};
use crate::lexer::{Lexer, Token};
use crate::{
    PageOrientation, SectionBreakType, SectionProperties, TabAlignment, TabLeader, TabStop,
    VerticalAlignment,
};

/// The most marks one edit stages.
pub(super) const MAX_MARKS: usize = 4_096;

/// One page-setup value of a section ([`super::Edit::set_section_values`]).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum SectionValue {
    /// The section break kind (`\sbknone`, `\sbkcol`, `\sbkpage`,
    /// `\sbkeven`, `\sbkodd`).
    BreakType(SectionBreakType),
    /// Page width in twips (`\pgwsxnN`).
    PageWidth(i32),
    /// Page height in twips (`\pghsxnN`).
    PageHeight(i32),
    /// Orientation (`\lndscpsxn` present for landscape).
    Orientation(PageOrientation),
    /// Left margin in twips (`\marglsxnN`).
    MarginLeft(i32),
    /// Right margin in twips (`\margrsxnN`).
    MarginRight(i32),
    /// Top margin in twips (`\margtsxnN`).
    MarginTop(i32),
    /// Bottom margin in twips (`\margbsxnN`).
    MarginBottom(i32),
    /// Gutter in twips (`\guttersxnN`).
    MarginGutter(i32),
    /// Header distance from the top edge in twips (`\headeryN`).
    HeaderDistance(i32),
    /// Footer distance from the bottom edge in twips (`\footeryN`).
    FooterDistance(i32),
    /// Number of equal-width columns (`\colsN`).
    Columns(u16),
    /// Space between equal-width columns in twips (`\colsxN`).
    ColumnSpacing(i32),
    /// A distinct first page (`\titlepg`).
    TitlePage(bool),
    /// Vertical alignment (`\vertalt`, `\vertalc`, `\vertalj`, `\vertalb`).
    VerticalAlignment(VerticalAlignment),
}

impl SectionValue {
    fn validate(self) -> Result<(), Error> {
        let twips = |value: i32, nonnegative: bool| {
            if (nonnegative && value < 0) || !(-31_680 * 5..=31_680 * 5).contains(&value) {
                Err(Error::unsupported_source(
                    "an RTF section length is outside the page-size range",
                ))
            } else {
                Ok(())
            }
        };
        match self {
            Self::PageWidth(value) | Self::PageHeight(value) => {
                if value <= 0 || value > 31_680 * 5 {
                    return Err(Error::unsupported_source(
                        "an RTF page extent must be positive and within the page-size range",
                    ));
                }
                Ok(())
            },
            Self::MarginTop(value) | Self::MarginBottom(value) => twips(value, false),
            Self::MarginLeft(value)
            | Self::MarginRight(value)
            | Self::MarginGutter(value)
            | Self::HeaderDistance(value)
            | Self::FooterDistance(value)
            | Self::ColumnSpacing(value) => twips(value, true),
            Self::Columns(count) => {
                if count == 0 || count > crate::MAX_SECTION_COLUMNS {
                    return Err(Error::unsupported_source(
                        "an RTF section has between 1 and 64 columns",
                    ));
                }
                Ok(())
            },
            _ => Ok(()),
        }
    }

    /// Apply the value to `properties`, as the reader would read it back.
    fn apply(self, properties: &mut SectionProperties) {
        match self {
            Self::BreakType(kind) => properties.break_type = kind,
            Self::PageWidth(value) => properties.page_width = value,
            Self::PageHeight(value) => properties.page_height = value,
            Self::Orientation(orientation) => properties.orientation = orientation,
            Self::MarginLeft(value) => properties.margin_left = value,
            Self::MarginRight(value) => properties.margin_right = value,
            Self::MarginTop(value) => properties.margin_top = value,
            Self::MarginBottom(value) => properties.margin_bottom = value,
            Self::MarginGutter(value) => properties.margin_gutter = value,
            Self::HeaderDistance(value) => properties.header_distance = value,
            Self::FooterDistance(value) => properties.footer_distance = value,
            Self::Columns(count) => properties.columns.count = count,
            Self::ColumnSpacing(value) => properties.columns.default_spacing = value,
            Self::TitlePage(on) => properties.title_page = on,
            Self::VerticalAlignment(alignment) => properties.vertical_alignment = alignment,
        }
    }
}

impl super::Edit {
    fn stage_mark(&mut self, mark: Mark) -> Result<&mut Self, Error> {
        if !self.operations.is_empty() {
            return Err(Error::unsupported_operation(
                "body splices and retained-destination edits cannot compose safely",
            ));
        }
        admit(&self.source, &self.marks, &mark)?;
        self.charge_replacement(mark.replacement_bytes())?;
        self.marks.push(mark);
        Ok(self)
    }

    /// Stages replacement of the instruction of field `field` (an index
    /// into [`Package::fields`]): the `\fldinst` group's text becomes
    /// `instruction`, and the field's result and every other byte keep
    /// their source spelling.
    ///
    /// # Errors
    /// Returns an error for an index past the fields, an empty or
    /// control-character instruction, a second edit of the same field,
    /// staged body work, or finite limits.
    pub fn set_field_instruction(
        &mut self,
        field: usize,
        instruction: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        self.stage_mark(Mark::FieldInstruction {
            field,
            instruction: instruction.into(),
        })
    }

    /// Stages renaming bookmark `bookmark` (an index into the document's
    /// bookmark table) in both its `\bkmkstart` and `\bkmkend` groups.
    ///
    /// # Errors
    /// Returns an error for an index past the bookmarks, an invalid or
    /// duplicate name, a second edit of the same bookmark, staged body
    /// work, or finite limits.
    pub fn rename_bookmark(
        &mut self,
        bookmark: usize,
        name: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        self.stage_mark(Mark::BookmarkName {
            bookmark,
            name: name.into(),
        })
    }

    /// Stages removal of bookmark `bookmark`'s `\bkmkstart` and
    /// `\bkmkend` groups; the text it covered stays.
    ///
    /// # Errors
    /// Returns an error for an index past the bookmarks, a second edit of
    /// the same bookmark, staged body work, or finite limits.
    pub fn remove_bookmark(&mut self, bookmark: usize) -> Result<&mut Self, Error> {
        self.stage_mark(Mark::RemoveBookmark { bookmark })
    }

    /// Stages a point bookmark named `name` where body paragraph
    /// `paragraph`'s text starts.
    ///
    /// # Errors
    /// Returns an error for a paragraph past the body, an invalid or
    /// existing name, a second bookmark at the same paragraph, staged body
    /// work, or finite limits.
    pub fn insert_bookmark(
        &mut self,
        paragraph: usize,
        name: impl Into<String>,
    ) -> Result<&mut Self, Error> {
        self.stage_mark(Mark::InsertBookmark {
            paragraph,
            name: name.into(),
        })
    }

    /// Stages replacement of body paragraph `paragraph`'s custom tab stops
    /// with `stops`, in order; an empty list removes them all.
    ///
    /// # Errors
    /// Returns an error for a paragraph past the body, more than
    /// [`crate::MAX_PARAGRAPH_TAB_STOPS`] stops or a position outside 0 to
    /// 31680 twips, a second edit of the same paragraph, staged body work,
    /// or finite limits.
    pub fn set_paragraph_tab_stops(
        &mut self,
        paragraph: usize,
        stops: &[TabStop],
    ) -> Result<&mut Self, Error> {
        self.stage_mark(Mark::ParagraphTabs {
            paragraph,
            stops: stops.to_vec(),
        })
    }

    /// Stages page-setup values of section `section` (an index into
    /// [`Package::sections`]).
    ///
    /// # Errors
    /// Returns an error for an index past the sections, a value outside
    /// its range, a second edit of the same section, staged body work, or
    /// finite limits.
    pub fn set_section_values(
        &mut self,
        section: usize,
        values: &[SectionValue],
    ) -> Result<&mut Self, Error> {
        self.stage_mark(Mark::SectionValues {
            section,
            values: values.to_vec(),
        })
    }

    pub(super) fn commit_marks(self) -> Result<Commit, Error> {
        if !self.operations.is_empty() {
            return Err(Error::unsupported_operation(
                "body splices and retained-destination edits cannot compose safely",
            ));
        }
        let count = self.marks.len();
        commit(self.source, &self.marks, count)
    }
}

/// One staged mark edit.
#[derive(Debug, Clone)]
pub(super) enum Mark {
    FieldInstruction {
        field: usize,
        instruction: String,
    },
    BookmarkName {
        bookmark: usize,
        name: String,
    },
    RemoveBookmark {
        bookmark: usize,
    },
    InsertBookmark {
        paragraph: usize,
        name: String,
    },
    ParagraphTabs {
        paragraph: usize,
        stops: Vec<TabStop>,
    },
    SectionValues {
        section: usize,
        values: Vec<SectionValue>,
    },
}

impl Mark {
    /// The entry the mark edits, so two marks of one entry conflict.
    fn key(&self) -> (u8, usize) {
        match self {
            Self::FieldInstruction { field, .. } => (0, *field),
            Self::BookmarkName { bookmark, .. } | Self::RemoveBookmark { bookmark } => {
                (1, *bookmark)
            },
            Self::InsertBookmark { paragraph, .. } => (2, *paragraph),
            Self::ParagraphTabs { paragraph, .. } => (3, *paragraph),
            Self::SectionValues { section, .. } => (4, *section),
        }
    }

    /// The bytes the mark authors, charged against the replacement limit.
    pub(super) fn replacement_bytes(&self) -> usize {
        match self {
            Self::FieldInstruction { instruction, .. } => instruction.len(),
            Self::BookmarkName { name, .. } | Self::InsertBookmark { name, .. } => {
                name.len().saturating_mul(2)
            },
            Self::RemoveBookmark { .. } => 0,
            Self::ParagraphTabs { stops, .. } => stops.len().saturating_mul(24),
            Self::SectionValues { values, .. } => values.len().saturating_mul(16),
        }
    }
}

/// Validate one mark against its source and the marks already staged.
pub(super) fn admit(source: &Package, staged: &[Mark], mark: &Mark) -> Result<(), Error> {
    if staged.len() >= MAX_MARKS {
        return Err(Error::Limit {
            kind: super::SemanticLimitKind::Operations,
            observed: staged.len().saturating_add(1),
            maximum: MAX_MARKS,
        });
    }
    let key = mark.key();
    if let Some(existing) = staged.iter().position(|held| held.key() == key) {
        return Err(Error::Conflict {
            existing,
            incoming: staged.len(),
        });
    }
    let model = source.model();
    match mark {
        Mark::FieldInstruction { field, instruction } => {
            if source.fields().get(*field).is_none() {
                return Err(Error::DestinationOutOfRange("field"));
            }
            check_text(instruction, "field instruction")?;
            if instruction.trim().is_empty() {
                return Err(Error::unsupported_source(
                    "an RTF field instruction cannot be empty",
                ));
            }
        },
        Mark::BookmarkName { bookmark, name } => {
            let bookmarks = model.bookmarks().bookmarks();
            if bookmarks.get(*bookmark).is_none() {
                return Err(Error::DestinationOutOfRange("bookmark"));
            }
            check_name(name)?;
            if bookmarks
                .iter()
                .enumerate()
                .any(|(index, held)| index != *bookmark && held.name == name.as_str())
            {
                return Err(Error::unsupported_source(
                    "another RTF bookmark already has this name",
                ));
            }
        },
        Mark::RemoveBookmark { bookmark } => {
            if model.bookmarks().bookmarks().get(*bookmark).is_none() {
                return Err(Error::DestinationOutOfRange("bookmark"));
            }
        },
        Mark::InsertBookmark { paragraph, name } => {
            if *paragraph >= source.paragraph_count() {
                return Err(Error::ParagraphOutOfRange {
                    position: *paragraph,
                    count: source.paragraph_count(),
                });
            }
            check_name(name)?;
            if model.bookmarks().contains(name) {
                return Err(Error::unsupported_source(
                    "another RTF bookmark already has this name",
                ));
            }
        },
        Mark::ParagraphTabs { paragraph, stops } => {
            if *paragraph >= source.paragraph_count() {
                return Err(Error::ParagraphOutOfRange {
                    position: *paragraph,
                    count: source.paragraph_count(),
                });
            }
            if stops.len() > crate::MAX_PARAGRAPH_TAB_STOPS {
                return Err(Error::unsupported_source(
                    "an RTF paragraph holds at most 64 tab stops",
                ));
            }
            for stop in stops {
                if !(0..=31_680).contains(&stop.position) {
                    return Err(Error::unsupported_source(
                        "an RTF tab stop position runs from 0 to 31680 twips",
                    ));
                }
            }
        },
        Mark::SectionValues { section, values } => {
            if source.sections().get(*section).is_none() {
                return Err(Error::DestinationOutOfRange("section"));
            }
            for value in values {
                value.validate()?;
            }
        },
    }
    Ok(())
}

fn check_text(text: &str, _field: &'static str) -> Result<(), Error> {
    if text.chars().any(|character| character.is_control()) {
        return Err(Error::unsupported_source(
            "RTF field instructions cannot hold control characters",
        ));
    }
    Ok(())
}

fn check_name(name: &str) -> Result<(), Error> {
    if name.is_empty()
        || name.len() > 255
        || name
            .chars()
            .any(|character| character.is_control() || matches!(character, '{' | '}' | '\\'))
    {
        return Err(Error::unsupported_source(
            "an RTF bookmark name is 1 to 255 characters without braces, backslashes or controls",
        ));
    }
    Ok(())
}

/// The letters of the control word whose source is `raw` (`\fldinst` →
/// `fldinst`).
fn control_name(raw: &[u8]) -> &[u8] {
    let Some(rest) = raw.strip_prefix(b"\\") else {
        return &[];
    };
    let letters = rest
        .iter()
        .take_while(|byte| byte.is_ascii_alphabetic())
        .count();
    rest.get(..letters).unwrap_or_default()
}

/// One group of the source: its braces and the first control word in it.
#[derive(Clone)]
struct Group {
    range: Range<usize>,
    /// The token index of the opening brace.
    open: usize,
    /// The token index of the closing brace.
    close: usize,
    /// The first control word after `{` (and after `\*`).
    word: Vec<u8>,
    /// The span of that control word.
    word_span: Range<usize>,
    depth: usize,
}

/// The tokenized source.
struct Source<'s> {
    input: &'s str,
    tokens: Vec<Token<'s>>,
    spans: Vec<Range<usize>>,
    groups: Vec<Group>,
}

impl<'s> Source<'s> {
    fn new(package: &Package, input: &'s str, arena: &'s Bump) -> Result<Self, Error> {
        let mut lexer = Lexer::new_with_limits(input, arena, package.limits())
            .for_document(package.options().leniency_value());
        let (tokens, spans) = lexer.tokenize_with_spans()?;
        let bytes = input.as_bytes();
        let mut stack: Vec<usize> = Vec::new();
        let mut groups: Vec<Group> = Vec::new();
        for (index, (token, span)) in tokens.iter().zip(&spans).enumerate() {
            match token {
                Token::OpenBrace => {
                    groups.push(Group {
                        range: span.start..span.end,
                        open: index,
                        close: index,
                        word: Vec::new(),
                        word_span: span.end..span.end,
                        depth: stack.len(),
                    });
                    stack.push(groups.len() - 1);
                },
                Token::CloseBrace => {
                    if let Some(group) = stack.pop().and_then(|open| groups.get_mut(open)) {
                        group.range.end = span.end;
                        group.close = index;
                    }
                },
                Token::Control(_) => {
                    if let Some(group) = stack.last().and_then(|&open| groups.get_mut(open)) {
                        if group.word.is_empty()
                            && index == group.next_word_index(&tokens, bytes, &spans)
                        {
                            let raw = bytes.get(span.clone()).unwrap_or_default();
                            let name = control_name(raw);
                            if !name.is_empty() {
                                group.word = name.to_vec();
                                group.word_span = span.clone();
                            }
                        }
                    }
                },
                _ => {},
            }
        }
        Ok(Self {
            input,
            tokens,
            spans,
            groups,
        })
    }

    fn bytes(&self) -> &'s [u8] {
        self.input.as_bytes()
    }

    /// The groups whose first control word is `word`, in source order.
    fn groups_named<'a>(&'a self, word: &'a [u8]) -> impl Iterator<Item = &'a Group> + 'a {
        self.groups.iter().filter(move |group| group.word == word)
    }

    /// The text tokens of `group`, concatenated, nested groups included.
    fn group_text(&self, group: &Group) -> String {
        let mut out = String::new();
        for token in self
            .tokens
            .get(group.open.saturating_add(1)..group.close)
            .unwrap_or_default()
        {
            if let Token::Text(text) = token {
                out.push_str(text);
            }
        }
        out
    }

    /// The first child group of `group` named `word`.
    fn child_named(&self, group: &Group, word: &[u8]) -> Option<&Group> {
        self.groups.iter().find(|child| {
            child.depth == group.depth + 1
                && child.range.start > group.range.start
                && child.range.end < group.range.end
                && child.word == word
        })
    }
}

impl Group {
    /// The token index of the first control word of the group, skipping
    /// an ignorable-destination marker `\*`.
    fn next_word_index(&self, tokens: &[Token<'_>], bytes: &[u8], spans: &[Range<usize>]) -> usize {
        let mut index = self.open.saturating_add(1);
        while let Some(token) = tokens.get(index) {
            match token {
                Token::Text(text) if text.trim().is_empty() => index += 1,
                Token::Control(_) => {
                    let raw = spans
                        .get(index)
                        .and_then(|span| bytes.get(span.clone()))
                        .unwrap_or_default();
                    if raw.starts_with(b"\\*") {
                        index += 1;
                        continue;
                    }
                    return index;
                },
                _ => return usize::MAX,
            }
        }
        usize::MAX
    }
}

/// Collapse white space and trim, for instruction and name comparison.
fn normalized(text: &str) -> String {
    text.split_whitespace().collect::<Vec<_>>().join(" ")
}

/// RTF-escape `text` for character data: backslash and braces, non-ASCII
/// as `\uN?`.
fn escape(text: &str) -> String {
    let mut out = String::with_capacity(text.len() + 8);
    for character in text.chars() {
        match character {
            '\\' => out.push_str("\\\\"),
            '{' => out.push_str("\\{"),
            '}' => out.push_str("\\}"),
            character if character.is_ascii() => out.push(character),
            character => {
                let mut buffer = [0_u16; 2];
                for unit in character.encode_utf16(&mut buffer) {
                    #[allow(
                        clippy::cast_possible_wrap,
                        reason = "RTF \\u takes the signed 16-bit spelling of the UTF-16 unit"
                    )]
                    let signed = *unit as i16;
                    out.push_str(&format!("\\u{signed}?"));
                }
            },
        }
    }
    out
}

/// The control words of one tab stop: alignment, leader, position.
fn tab_controls(stop: &TabStop) -> String {
    let mut out = String::new();
    match stop.alignment {
        TabAlignment::Center => out.push_str("\\tqc"),
        TabAlignment::Right => out.push_str("\\tqr"),
        TabAlignment::Decimal => out.push_str("\\tqdec"),
        TabAlignment::Left | TabAlignment::Bar => {},
    }
    match stop.leader {
        TabLeader::Dot => out.push_str("\\tldot"),
        TabLeader::MiddleDot => out.push_str("\\tlmdot"),
        TabLeader::Hyphen => out.push_str("\\tlhyph"),
        TabLeader::Underscore => out.push_str("\\tlul"),
        TabLeader::ThickLine => out.push_str("\\tlth"),
        TabLeader::Equal => out.push_str("\\tleq"),
        TabLeader::None => {},
    }
    if stop.alignment == TabAlignment::Bar {
        out.push_str(&format!("\\tb{}", stop.position));
    } else {
        out.push_str(&format!("\\tx{}", stop.position));
    }
    out
}

/// Whether a control word name is one of a tab stop's controls.
fn is_tab_control(name: &[u8]) -> bool {
    matches!(
        name,
        b"tx"
            | b"tb"
            | b"tqr"
            | b"tqc"
            | b"tqdec"
            | b"tldot"
            | b"tlmdot"
            | b"tlhyph"
            | b"tlul"
            | b"tlth"
            | b"tleq"
    )
}

/// One splice of the source: its range and replacement.
struct Splice {
    range: Range<usize>,
    bytes: String,
}

/// The source groups of the fields, in the order the native model lists
/// the fields: the `k`th field whose instruction reads as `I` is the `k`th
/// source `\field` group whose `\fldinst` reads as `I`.
fn field_group<'a>(
    source: &'a Source<'_>,
    package: &Package,
    field: usize,
) -> Result<&'a Group, Error> {
    let fields = package.fields();
    let target = normalized(
        &fields
            .get(field)
            .ok_or(Error::DestinationOutOfRange("field"))?
            .instruction,
    );
    let occurrence = fields
        .iter()
        .take(field)
        .filter(|held| normalized(&held.instruction) == target)
        .count();
    source
        .groups_named(b"field")
        .filter_map(|group| source.child_named(group, b"fldinst"))
        .filter(|instruction| normalized(&source.group_text(instruction)) == target)
        .nth(occurrence)
        .ok_or(Error::unsupported_source(
            "the RTF field's \\fldinst group is not found in the source",
        ))
}

/// The `\bkmkstart` and `\bkmkend` groups of bookmark `bookmark`.
fn bookmark_groups<'a>(
    source: &'a Source<'_>,
    package: &Package,
    bookmark: usize,
) -> Result<(&'a Group, Option<&'a Group>), Error> {
    let bookmarks = package.model().bookmarks().bookmarks();
    let name = bookmarks
        .get(bookmark)
        .ok_or(Error::DestinationOutOfRange("bookmark"))?
        .name
        .to_string();
    let occurrence = bookmarks
        .iter()
        .take(bookmark)
        .filter(|held| held.name == name.as_str())
        .count();
    let start = source
        .groups_named(b"bkmkstart")
        .filter(|group| source.group_text(group).trim() == name)
        .nth(occurrence)
        .ok_or(Error::unsupported_source(
            "the RTF bookmark's \\bkmkstart group is not found in the source",
        ))?;
    let end = source.groups_named(b"bkmkend").find(|group| {
        group.range.start > start.range.start && source.group_text(group).trim() == name
    });
    Ok((start, end))
}

/// The splice that replaces a destination group's text after its control
/// word with `text`.
fn rename(source: &Source<'_>, group: &Group, text: &str) -> Splice {
    let content_start = group.word_span.end;
    let content_end = group.range.end.saturating_sub(1);
    let separator = if source.bytes().get(content_start.wrapping_sub(1)) == Some(&b' ')
        || source.bytes().get(content_start) == Some(&b' ')
    {
        ""
    } else {
        " "
    };
    Splice {
        range: content_start..content_end,
        bytes: format!("{separator}{}", escape(text)),
    }
}

/// The source range of each body paragraph.
fn paragraph_ranges(package: &Package, input: &str) -> Result<Vec<Range<usize>>, Error> {
    let cut = super::producer_paragraphs::paragraph_ranges(package, input)?;
    if cut.len() != package.paragraph_count() {
        return Err(Error::unsupported_source(
            "the body's paragraphs are not delimited by root-level \\par controls",
        ));
    }
    Ok(cut)
}

/// The splices of a paragraph's tab-stop edit: every tab control at the
/// root level of the paragraph's segment is removed, and the new controls
/// take the place of the first one (or follow the segment's last
/// `\pard`, or open the segment).
fn tab_splices(
    source: &Source<'_>,
    range: &Range<usize>,
    stops: &[TabStop],
    out: &mut Vec<Splice>,
) {
    let bytes = source.bytes();
    let mut depth = 0usize;
    let mut first: Option<usize> = None;
    let mut last_pard: Option<usize> = None;
    let mut removed = Vec::new();
    for (token, span) in source.tokens.iter().zip(&source.spans) {
        if span.start < range.start {
            continue;
        }
        if span.start >= range.end {
            break;
        }
        match token {
            Token::OpenBrace => depth += 1,
            Token::CloseBrace => depth = depth.saturating_sub(1),
            Token::Control(_) if depth == 0 => {
                let raw = bytes.get(span.clone()).unwrap_or_default();
                let name = control_name(raw);
                if is_tab_control(name) {
                    first.get_or_insert(span.start);
                    // A control's delimiting space belongs to it.
                    let end = if bytes.get(span.end) == Some(&b' ') && !raw.ends_with(b" ") {
                        span.end + 1
                    } else {
                        span.end
                    };
                    removed.push((span.start..end, end > span.end || raw.ends_with(b" ")));
                } else if name == b"pard" {
                    last_pard = Some(span.end);
                }
            },
            _ => {},
        }
    }
    let mut controls: String = stops.iter().map(tab_controls).collect();
    let anchor = first.or(last_pard).unwrap_or(range.start);
    if first.is_some() {
        // The controls take the first removed control's place; the removed
        // run's delimiter space, if any, follows them.
        if removed.iter().any(|(_, space)| *space) && !controls.is_empty() {
            controls.push(' ');
        }
        for (index, (removal, _)) in removed.into_iter().enumerate() {
            out.push(Splice {
                range: removal,
                bytes: if index == 0 {
                    std::mem::take(&mut controls)
                } else {
                    String::new()
                },
            });
        }
    } else if !controls.is_empty() {
        // A control inserted before text needs a delimiter.
        let next = bytes.get(anchor).copied();
        let delimiter = match next {
            Some(byte) if byte.is_ascii_alphanumeric() || byte == b' ' => " ",
            _ => "",
        };
        out.push(Splice {
            range: anchor..anchor,
            bytes: format!("{controls}{delimiter}"),
        });
    }
}

/// Visible body section breaks in source order, including formatting groups.
fn section_breaks(source: &Source<'_>) -> Vec<Range<usize>> {
    let bytes = source.bytes();
    let mut destinations = Vec::new();
    let mut out = Vec::new();
    for (token, span) in source.tokens.iter().zip(&source.spans) {
        match token {
            Token::OpenBrace => destinations.push(
                destinations.last().copied().unwrap_or(false)
                    || super::producer_paragraphs::opens_destination(
                        bytes.get(span.end..).unwrap_or_default(),
                    ),
            ),
            Token::CloseBrace => {
                destinations.pop();
            },
            Token::Control(crate::lexer::ControlWord::Section)
                if !destinations.last().copied().unwrap_or(false) =>
            {
                out.push(span.clone())
            },
            _ => {},
        }
    }
    out
}

/// The control word and parameter that spell one section value, and the
/// control names of its family (replaced by it).
fn section_control(value: SectionValue) -> (String, &'static [&'static [u8]]) {
    match value {
        SectionValue::BreakType(kind) => (
            match kind {
                SectionBreakType::Continuous => "\\sbknone".to_owned(),
                SectionBreakType::Column => "\\sbkcol".to_owned(),
                SectionBreakType::Page => "\\sbkpage".to_owned(),
                SectionBreakType::EvenPage => "\\sbkeven".to_owned(),
                SectionBreakType::OddPage => "\\sbkodd".to_owned(),
            },
            &[b"sbknone", b"sbkcol", b"sbkpage", b"sbkeven", b"sbkodd"],
        ),
        SectionValue::PageWidth(value) => (format!("\\pgwsxn{value}"), &[b"pgwsxn"]),
        SectionValue::PageHeight(value) => (format!("\\pghsxn{value}"), &[b"pghsxn"]),
        SectionValue::Orientation(orientation) => (
            match orientation {
                PageOrientation::Landscape => "\\lndscpsxn".to_owned(),
                _ => String::new(),
            },
            &[b"lndscpsxn"],
        ),
        SectionValue::MarginLeft(value) => (format!("\\marglsxn{value}"), &[b"marglsxn"]),
        SectionValue::MarginRight(value) => (format!("\\margrsxn{value}"), &[b"margrsxn"]),
        SectionValue::MarginTop(value) => (format!("\\margtsxn{value}"), &[b"margtsxn"]),
        SectionValue::MarginBottom(value) => (format!("\\margbsxn{value}"), &[b"margbsxn"]),
        SectionValue::MarginGutter(value) => (format!("\\guttersxn{value}"), &[b"guttersxn"]),
        SectionValue::HeaderDistance(value) => (format!("\\headery{value}"), &[b"headery"]),
        SectionValue::FooterDistance(value) => (format!("\\footery{value}"), &[b"footery"]),
        SectionValue::Columns(count) => (format!("\\cols{count}"), &[b"cols"]),
        SectionValue::ColumnSpacing(value) => (format!("\\colsx{value}"), &[b"colsx"]),
        SectionValue::TitlePage(on) => (
            if on {
                "\\titlepg".to_owned()
            } else {
                String::new()
            },
            &[b"titlepg"],
        ),
        SectionValue::VerticalAlignment(alignment) => (
            match alignment {
                VerticalAlignment::Top => "\\vertalt".to_owned(),
                VerticalAlignment::Center => "\\vertalc".to_owned(),
                VerticalAlignment::Justify => "\\vertalj".to_owned(),
                VerticalAlignment::Bottom => "\\vertalb".to_owned(),
            },
            &[b"vertalt", b"vertalc", b"vertalj", b"vertalb"],
        ),
    }
}

/// The splices of a section's page-setup edit, inside the section's source
/// region: an existing control of each value's family is replaced (the
/// others of the family removed); a missing one follows the region's
/// `\sectd`, else opens the region's first paragraph.
fn section_splices(
    source: &Source<'_>,
    region: &Range<usize>,
    values: &[SectionValue],
    out: &mut Vec<Splice>,
) -> Result<(), Error> {
    let bytes = source.bytes();
    let mut depth = 0usize;
    let mut controls: Vec<(Vec<u8>, Range<usize>)> = Vec::new();
    let mut anchor: Option<usize> = None;
    let mut first_pard: Option<usize> = None;
    for (token, span) in source.tokens.iter().zip(&source.spans) {
        match token {
            Token::OpenBrace => depth += 1,
            Token::CloseBrace => depth = depth.saturating_sub(1),
            Token::Control(_)
                if depth == 1 && span.start >= region.start && span.start < region.end =>
            {
                let raw = bytes.get(span.clone()).unwrap_or_default();
                let name = control_name(raw).to_vec();
                if name == b"sectd" {
                    anchor.get_or_insert(span.end);
                } else if name == b"pard" {
                    first_pard.get_or_insert(span.start);
                }
                let end = if bytes.get(span.end) == Some(&b' ') && !raw.ends_with(b" ") {
                    span.end + 1
                } else {
                    span.end
                };
                controls.push((name, span.start..end));
            },
            _ => {},
        }
    }
    let anchor = anchor.or(first_pard).ok_or(Error::unsupported_source(
        "the RTF section has no \\sectd or paragraph to place page setup at",
    ))?;
    let mut inserts = String::new();
    for value in values {
        let (control, family) = section_control(*value);
        let mut placed = false;
        for (name, range) in &controls {
            if family.contains(&name.as_slice()) {
                let spaced = bytes.get(range.end.wrapping_sub(1)) == Some(&b' ');
                let replacement = if placed || control.is_empty() {
                    String::new()
                } else if spaced {
                    format!("{control} ")
                } else {
                    control.clone()
                };
                placed = true;
                out.push(Splice {
                    range: range.clone(),
                    bytes: replacement,
                });
            }
        }
        if !placed {
            inserts.push_str(&control);
        }
    }
    if !inserts.is_empty() {
        let next = bytes.get(anchor).copied();
        let delimiter = match next {
            Some(byte) if byte.is_ascii_alphanumeric() || byte == b' ' => " ",
            _ => "",
        };
        out.push(Splice {
            range: anchor..anchor,
            bytes: format!("{inserts}{delimiter}"),
        });
    }
    Ok(())
}

/// Resolve a retained section definition through the parser's actual body
/// boundaries. Inherited sections need not have separate definitions.
fn section_region(
    package: &Package,
    source: &Source<'_>,
    section: usize,
) -> Result<Range<usize>, Error> {
    let breaks = section_breaks(source);
    let boundaries = package.raw().section_breaks().collect::<Vec<_>>();
    if breaks.len() != boundaries.len() {
        return Err(Error::unsupported_source(
            "section source boundaries do not match the main story",
        ));
    }
    let previous = if section == 0 {
        None
    } else {
        Some(
            boundaries
                .iter()
                .position(|b| b.next_section == Some(section))
                .ok_or(Error::unsupported_source(
                    "section definition has no matching body boundary",
                ))?,
        )
    };
    let start = match previous {
        Some(index) => {
            breaks
                .get(index)
                .ok_or(Error::unsupported_source(
                    "section source boundary is missing",
                ))?
                .end
        },
        None => 0,
    };
    let next = previous.map_or(0, |i| i + 1);
    let end = breaks
        .get(next)
        .map_or(source.input.len(), |span| span.start);
    Ok(start..end)
}

/// Apply `splices` to `input`; they must not overlap.
fn assemble(input: &str, mut splices: Vec<Splice>) -> Result<Vec<u8>, Error> {
    splices.sort_by_key(|splice| (splice.range.start, splice.range.end));
    for pair in splices.windows(2) {
        if let [left, right] = pair
            && right.range.start < left.range.end
        {
            return Err(Error::unsupported_source(
                "two RTF mark edits touch the same source bytes",
            ));
        }
    }
    let bytes = input.as_bytes();
    let outside = || Error::unsupported_source("an RTF mark splice falls outside the source");
    let mut out = Vec::with_capacity(input.len() + 64);
    let mut cursor = 0;
    for splice in splices {
        out.extend_from_slice(bytes.get(cursor..splice.range.start).ok_or_else(outside)?);
        out.extend_from_slice(splice.bytes.as_bytes());
        cursor = splice.range.end;
    }
    out.extend_from_slice(bytes.get(cursor..).ok_or_else(outside)?);
    Ok(out)
}

/// The candidate bytes of `marks` on `package`, before readback.
fn splice(package: &Package, input: &str, marks: &[Mark]) -> Result<Vec<u8>, Error> {
    let arena = Bump::new();
    let source = Source::new(package, input, &arena)?;
    let mut splices = Vec::new();
    let mut paragraphs: Option<Vec<Range<usize>>> = None;
    for mark in marks {
        match mark {
            Mark::FieldInstruction { field, instruction } => {
                let group = field_group(&source, package, *field)?;
                splices.push(Splice {
                    range: group.word_span.end..group.range.end.saturating_sub(1),
                    bytes: format!("{{ {} }}", escape(instruction.trim())),
                });
            },
            Mark::BookmarkName { bookmark, name } => {
                let (start, end) = bookmark_groups(&source, package, *bookmark)?;
                splices.push(rename(&source, start, name));
                if let Some(end) = end {
                    splices.push(rename(&source, end, name));
                }
            },
            Mark::RemoveBookmark { bookmark } => {
                let (start, end) = bookmark_groups(&source, package, *bookmark)?;
                splices.push(Splice {
                    range: start.range.clone(),
                    bytes: String::new(),
                });
                if let Some(end) = end {
                    splices.push(Splice {
                        range: end.range.clone(),
                        bytes: String::new(),
                    });
                }
            },
            Mark::InsertBookmark { paragraph, name } => {
                let ranges = match &paragraphs {
                    Some(ranges) => ranges,
                    None => paragraphs.insert(paragraph_ranges(package, input)?),
                };
                let range = ranges
                    .get(*paragraph)
                    .ok_or(Error::DestinationOutOfRange("paragraph"))?;
                // A point bookmark where the paragraph's text starts: after
                // its last root-level paragraph reset.
                let at = paragraph_text_start(&source, range);
                let escaped = escape(name);
                splices.push(Splice {
                    range: at..at,
                    bytes: format!("{{\\*\\bkmkstart {escaped}}}{{\\*\\bkmkend {escaped}}}"),
                });
            },
            Mark::ParagraphTabs { paragraph, stops } => {
                let ranges = match &paragraphs {
                    Some(ranges) => ranges,
                    None => paragraphs.insert(paragraph_ranges(package, input)?),
                };
                let range = ranges
                    .get(*paragraph)
                    .ok_or(Error::DestinationOutOfRange("paragraph"))?
                    .clone();
                tab_splices(&source, &range, stops, &mut splices);
            },
            Mark::SectionValues { section, values } => {
                let region = section_region(package, &source, *section)?;
                section_splices(&source, &region, values, &mut splices)?;
            },
        }
    }
    assemble(input, splices)
}

/// Where a paragraph's text starts in its segment: after its root-level
/// control words.
fn paragraph_text_start(source: &Source<'_>, range: &Range<usize>) -> usize {
    let bytes = source.bytes();
    let mut at = range.start;
    for (token, span) in source.tokens.iter().zip(&source.spans) {
        if span.start < range.start {
            continue;
        }
        if span.start >= range.end {
            break;
        }
        match token {
            Token::Control(_) => {
                let raw = bytes.get(span.clone()).unwrap_or_default();
                at = if bytes.get(span.end) == Some(&b' ') && !raw.ends_with(b" ") {
                    span.end + 1
                } else {
                    span.end
                };
            },
            Token::Text(text) if text.trim().is_empty() => at = span.end,
            _ => break,
        }
    }
    at
}

/// The runs of a paragraph, adjacent runs of equal formatting joined.
fn coalesced<'a>(
    paragraph: crate::text::Paragraph<'a>,
) -> Vec<(String, &'a crate::types::Formatting)> {
    let mut out: Vec<(String, &crate::types::Formatting)> = Vec::new();
    for run in paragraph.runs() {
        let format = run.format().raw();
        match out.last_mut() {
            Some((text, held)) if *held == format => text.push_str(run.text()),
            _ => out.push((run.text().to_owned(), format)),
        }
    }
    out
}

/// The readback expectations: the candidate reads as the source with the
/// marks applied, and nothing else changed.
fn verify(source: &Package, candidate: &Package, marks: &[Mark]) -> Result<(), Error> {
    let refuse = |what: &'static str| Err(Error::unsupported_source(what));
    if source.text() != candidate.text() {
        return refuse("the RTF mark edit changed the body text");
    }
    let before: Vec<_> = source.body().paragraphs().collect();
    let after: Vec<_> = candidate.body().paragraphs().collect();
    if before.len() != after.len() {
        return refuse("the RTF mark edit changed the paragraph count");
    }
    for (index, (left, right)) in before.iter().zip(&after).enumerate() {
        let mut expected = *left.format().raw();
        for mark in marks {
            if let Mark::ParagraphTabs { paragraph, stops } = mark
                && *paragraph == index
            {
                // A fresh table: `TabStops` equality also compares unused slots.
                expected.tab_stops = crate::TabStops::new();
                for stop in stops {
                    expected
                        .tab_stops
                        .push(*stop)
                        .map_err(|_error| Error::unsupported_source("too many RTF tab stops"))?;
                }
            }
        }
        if *right.format().raw() != expected {
            return refuse("an RTF paragraph's properties do not read back as staged");
        }
        // A bookmark boundary may split a run; runs compare coalesced.
        if coalesced(*left) != coalesced(*right) {
            return refuse("an RTF run does not read back unchanged");
        }
    }
    // Fields: the edited instructions, every other field unchanged.
    let fields_before = source.fields();
    let fields_after = candidate.fields();
    if fields_before.len() != fields_after.len() {
        return refuse("the RTF mark edit changed the field count");
    }
    for (index, (left, right)) in fields_before.iter().zip(fields_after).enumerate() {
        let staged = marks.iter().find_map(|mark| match mark {
            Mark::FieldInstruction { field, instruction } if *field == index => Some(instruction),
            _ => None,
        });
        match staged {
            Some(instruction) => {
                if normalized(&right.instruction) != normalized(instruction)
                    || left.result != right.result
                    || left.position != right.position
                {
                    return refuse("the RTF field instruction does not read back as staged");
                }
            },
            None => {
                if left.instruction != right.instruction
                    || left.result != right.result
                    || left.field_type != right.field_type
                {
                    return refuse("an RTF field does not read back unchanged");
                }
            },
        }
    }
    // Bookmarks: renamed, removed and inserted as staged.
    let mut expected: Vec<(String, usize)> = source
        .model()
        .bookmarks()
        .bookmarks()
        .iter()
        .map(|bookmark| (bookmark.name.to_string(), bookmark.position))
        .collect();
    let mut removed: Vec<usize> = Vec::new();
    let mut inserted: Vec<&str> = Vec::new();
    for mark in marks {
        match mark {
            Mark::BookmarkName { bookmark, name } => {
                if let Some(entry) = expected.get_mut(*bookmark) {
                    entry.0.clone_from(name);
                }
            },
            Mark::RemoveBookmark { bookmark } => removed.push(*bookmark),
            Mark::InsertBookmark { name, .. } => inserted.push(name),
            _ => {},
        }
    }
    let kept: Vec<(String, usize)> = expected
        .into_iter()
        .enumerate()
        .filter(|(index, _)| !removed.contains(index))
        .map(|(_, entry)| entry)
        .collect();
    let observed: Vec<(String, usize)> = candidate
        .model()
        .bookmarks()
        .bookmarks()
        .iter()
        .filter(|bookmark| !inserted.contains(&bookmark.name.as_ref()))
        .map(|bookmark| (bookmark.name.to_string(), bookmark.position))
        .collect();
    if kept != observed {
        return refuse("the RTF bookmarks do not read back as staged");
    }
    for name in inserted {
        if !candidate.model().bookmarks().contains(name) {
            return refuse("an inserted RTF bookmark does not read back");
        }
    }
    // Sections: the edited page setup, every other section unchanged.
    let sections_before = source.sections();
    let sections_after = candidate.sections();
    if sections_before.len() != sections_after.len() {
        return refuse("the RTF mark edit changed the section count");
    }
    for (index, (left, right)) in sections_before.iter().zip(sections_after).enumerate() {
        let mut expected = left.properties.clone();
        for mark in marks {
            if let Mark::SectionValues { section, values } = mark
                && *section == index
            {
                for value in values {
                    value.apply(&mut expected);
                }
            }
        }
        if format!("{expected:?}") != format!("{:?}", right.properties)
            || format!("{:?}", left.headers_footers) != format!("{:?}", right.headers_footers)
        {
            return refuse("an RTF section does not read back as staged");
        }
    }
    let same_tables = source.tables().len() == candidate.tables().len()
        && source
            .tables()
            .iter()
            .zip(candidate.tables())
            .all(|(left, right)| format!("{left:?}") == format!("{right:?}"));
    let same = same_tables
        && source.pictures().len() == candidate.pictures().len()
        && source.shapes().len() == candidate.shapes().len()
        && source.notes().len() == candidate.notes().len()
        && source.annotations().len() == candidate.annotations().len()
        && source.revisions().len() == candidate.revisions().len()
        && source.objects().len() == candidate.objects().len()
        && source.opaque().len() == candidate.opaque().len()
        && source.model().unknown_syntax_markers() == candidate.model().unknown_syntax_markers();
    if !same {
        return refuse("the RTF mark edit changed tables, drawings, notes or other content");
    }
    Ok(())
}

/// Publish `marks` on `source`.
pub(super) fn commit(
    source: Package,
    marks: &[Mark],
    operation_count: usize,
) -> Result<Commit, Error> {
    let source_bytes = source.source_bytes().ok_or(Error::unsupported_operation(
        "snapshot has no exact RTF source",
    ))?;
    if crate::compressed::is_compressed_rtf(source_bytes) {
        return Err(Error::unsupported_source(
            "RTF mark edits refuse compressed RTF",
        ));
    }
    let input = std::str::from_utf8(source_bytes).map_err(|_error| {
        Error::unsupported_source("RTF mark edits require an ASCII or UTF-8 transport")
    })?;
    let bytes = splice(&source, input, marks)?;
    if bytes == source_bytes {
        return Ok(Commit::new(
            source.clone(),
            source,
            operation_count,
            Vec::new(),
        ));
    }
    super::ensure_changed_publication_allowed(&source)?;
    let candidate = Package::from_bytes_with_options(&bytes, source.options())
        .map_err(|_error| Error::unsupported_source("the RTF mark edit does not parse as RTF"))?;
    verify(&source, &candidate, marks)?;
    let delta = vec![Change::RootTransfer {
        vocabulary: "mark.transfer",
        effect: "body:marks".to_owned(),
        before: source_bytes.to_vec(),
        after: bytes,
    }];
    Ok(Commit::new(source, candidate, operation_count, delta))
}
