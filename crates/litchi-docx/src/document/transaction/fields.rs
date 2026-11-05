//! Field results of the main story, read and rewritten for field refresh
//! (record 0279).
//!
//! [`Fields`] indexes the main document: every paragraph (its style, direct
//! outline level, text and bookmarks), every simple (`w:fldSimple`) and
//! complex (`w:fldChar`) field with its instruction, cached result and
//! nesting, and every bookmark. Nothing is evaluated here: the caller
//! computes new results and hands them back as [`Update`]s.
//!
//! [`Update::Text`] replaces one field's cached result with text, keeping
//! the instruction, the markers and the first result run's formatting.
//! [`Update::Toc`] replaces a table of contents' result with one paragraph
//! per [`TocEntry`] (styles `TOC1`..`TOC9`), keeping the field code: the
//! paragraph holding the begin marker keeps its content up to the separator,
//! the entries follow, and the paragraph holding the end marker keeps its
//! content from the end marker on. [`Update::Bookmark`] wraps a paragraph's
//! content in a new bookmark, the anchor of a table-of-contents entry.
//!
//! All updates of one call are spliced into the document once and recorded
//! as exact element operations, so the patch replays and inverts exactly.

use std::ops::Range;
use std::sync::Arc;

use litchi_core::Position;
use litchi_core::xml::ReaderOrigin;
use litchi_ooxml_common::xml::attributes::BytesStartExt as _;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::ResolveResult;
use quick_xml::reader::NsReader;

use super::formatting::{ElementPath, ElementStep};
use super::substitute::FieldKind;
use super::{
    Edit, Error, MAX_DOCUMENT_DEPTH, MAX_OPERATIONS, MAX_REPLACEMENT_TEXT_BYTES, Operation,
    Refusal, Result, decode_text_fragment, replace_ranges, rewrite_run_region, scan_run_region,
};
use crate::namespace::is_wordprocessing_namespace;

#[cfg(test)]
mod tests;

/// The most fields one index lists.
const MAX_FIELDS: usize = 262_144;
/// The most text a bookmark keeps.
const MAX_BOOKMARK_TEXT: usize = 64 * 1024;
/// The deepest field nesting the index follows.
const MAX_FIELD_NESTING: usize = 64;
/// The markup-compatibility namespace: its `Fallback` content duplicates
/// the chosen content and is not indexed.
const MARKUP_COMPATIBILITY: &[u8] = b"http://schemas.openxmlformats.org/markup-compatibility/2006";

/// One paragraph of the main story.
#[derive(Clone)]
pub struct Paragraph {
    style: Option<String>,
    outline: Option<u8>,
    text: String,
    anchor: Option<String>,
    nested: bool,
    /// The paragraph element.
    range: Range<usize>,
    start_tag_end: usize,
    properties: Option<Range<usize>>,
    close_start: usize,
    parent: usize,
    /// The body child holding the paragraph.
    block: Option<usize>,
    /// The qualified name prefix with its colon (`w:`), or empty.
    prefix: String,
    body_position: Option<usize>,
}

impl Paragraph {
    /// The paragraph style id (`w:pStyle`).
    #[must_use]
    pub fn style(&self) -> Option<&str> {
        self.style.as_deref()
    }

    /// The direct outline level (`w:outlineLvl`), `0` for level 1 through
    /// `8` for level 9; body text (`9`) and invalid values are `None`.
    #[must_use]
    pub const fn outline_level(&self) -> Option<u8> {
        self.outline
    }

    /// The visible text: `w:t` text, including field results and excluding
    /// field instructions; a tab is `\t` and a break a space.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The anchor of the paragraph's first internal hyperlink.
    #[must_use]
    pub fn hyperlink_anchor(&self) -> Option<&str> {
        self.anchor.as_deref()
    }

    /// The zero-based position among the body's direct `w:p` children,
    /// for a direct body paragraph.
    #[must_use]
    pub const fn body_position(&self) -> Option<usize> {
        self.body_position
    }

    /// Whether the paragraph is inside another paragraph's run (a text
    /// box) or inside markup-compatibility content.
    #[must_use]
    pub const fn is_nested(&self) -> bool {
        self.nested
    }
}

impl std::fmt::Debug for Paragraph {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Paragraph")
            .field("style", &self.style)
            .field("outline", &self.outline)
            .field("bytes", &self.text.len())
            .finish_non_exhaustive()
    }
}

/// A field marker run.
#[derive(Clone, Default)]
struct Marker {
    run: Range<usize>,
    parent: usize,
    properties: Option<Range<usize>>,
    /// The `w:fldChar` element, when it is an empty element.
    character: Option<Range<usize>>,
}

/// The markup of a simple field.
#[derive(Clone)]
struct Simple {
    element: Range<usize>,
    start_tag_end: usize,
    /// `None` for an empty element.
    close_start: Option<usize>,
    parent: usize,
}

/// One field of the main story.
#[derive(Clone)]
pub struct Field {
    instruction: String,
    kind: FieldKind,
    parent: Option<usize>,
    begin_paragraph: usize,
    end_paragraph: usize,
    result: String,
    separated: bool,
    closed: bool,
    locked: bool,
    dirty: bool,
    nested_paragraph: bool,
    begin: Marker,
    separate: Option<Marker>,
    end: Option<Marker>,
    simple: Option<Simple>,
}

impl Field {
    /// The instruction (`TOC \o "1-3" \h`), decoded, spaces kept.
    #[must_use]
    pub fn instruction(&self) -> &str {
        &self.instruction
    }

    /// How the field is written.
    #[must_use]
    pub const fn kind(&self) -> FieldKind {
        self.kind
    }

    /// The index of the field whose instruction or result holds this one.
    #[must_use]
    pub const fn parent(&self) -> Option<usize> {
        self.parent
    }

    /// The paragraphs holding the begin and the end marker (equal for a
    /// one-paragraph field).
    #[must_use]
    pub const fn paragraphs(&self) -> (usize, usize) {
        (self.begin_paragraph, self.end_paragraph)
    }

    /// The cached result text; paragraphs are separated by `\n`, and
    /// leading and trailing paragraph breaks are dropped.
    #[must_use]
    pub fn result(&self) -> &str {
        &self.result
    }

    /// Whether the field has a result separator (simple fields always do).
    #[must_use]
    pub const fn has_result(&self) -> bool {
        self.separated
    }

    /// Whether the field is locked against updates (`w:fldLock`).
    #[must_use]
    pub const fn is_locked(&self) -> bool {
        self.locked
    }

    /// Whether the field is marked for update (`w:dirty`).
    #[must_use]
    pub const fn is_dirty(&self) -> bool {
        self.dirty
    }

    /// Whether the field begins in a nested paragraph (see
    /// [`Paragraph::is_nested`]).
    #[must_use]
    pub const fn is_in_nested_paragraph(&self) -> bool {
        self.nested_paragraph
    }
}

impl std::fmt::Debug for Field {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Field")
            .field("kind", &self.kind)
            .field("paragraphs", &(self.begin_paragraph, self.end_paragraph))
            .field("parent", &self.parent)
            .field("locked", &self.locked)
            .finish_non_exhaustive()
    }
}

/// One bookmark start.
#[derive(Clone, Debug)]
pub struct Bookmark {
    name: String,
    id: Option<i64>,
    paragraph: Option<usize>,
    text: String,
    /// The `w:bookmarkStart` tag and whether it is an empty element.
    tag: Range<usize>,
    empty: bool,
    /// The matching empty `w:bookmarkEnd` element.
    end: Option<Range<usize>>,
}

impl Bookmark {
    /// The bookmark name.
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// The bookmark id, when numeric.
    #[must_use]
    pub const fn id(&self) -> Option<i64> {
        self.id
    }

    /// The paragraph the bookmark starts in.
    #[must_use]
    pub const fn paragraph(&self) -> Option<usize> {
        self.paragraph
    }

    /// The visible text between the bookmark's start and its end, as
    /// [`Paragraph::text`] reads it; at most 64 KiB are kept.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }
}

/// A direct child of `w:body`.
#[derive(Clone)]
struct BodyChild {
    step: Option<ElementStep>,
    range: Range<usize>,
}

/// The paragraphs, fields and bookmarks of the main story.
#[derive(Clone)]
pub struct Fields {
    xml: Arc<Vec<u8>>,
    paragraphs: Vec<Paragraph>,
    fields: Vec<Field>,
    bookmarks: Vec<Bookmark>,
    children: Vec<BodyChild>,
    text_width: Option<u32>,
}

impl Fields {
    /// The paragraphs, in document order of their start tags.
    #[must_use]
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }

    /// The fields, in document order of their begin markers.
    #[must_use]
    pub fn fields(&self) -> &[Field] {
        &self.fields
    }

    /// The bookmark starts, in document order.
    #[must_use]
    pub fn bookmarks(&self) -> &[Bookmark] {
        &self.bookmarks
    }

    /// The direct body child holding paragraph `paragraph` (the paragraph
    /// itself, a table or a block content control), for a paragraph that
    /// is not nested (record 0325).
    #[must_use]
    pub fn body_block(&self, paragraph: usize) -> Option<ElementStep> {
        let block = self.paragraphs.get(paragraph)?.block?;
        self.children.get(block)?.step
    }

    /// The text width of the body's final section in twips (page width
    /// less the left and right margins), when it declares both.
    #[must_use]
    pub const fn text_width(&self) -> Option<u32> {
        self.text_width
    }

    /// Whether [`Update::Text`] can rewrite `field`: its markers lie in one
    /// paragraph under one parent, and its result holds only plain runs.
    #[must_use]
    pub fn accepts_text(&self, field: usize) -> bool {
        self.text_region(field).is_ok()
    }

    /// Whether [`Update::Toc`] can rewrite `field`: its begin and separate
    /// markers (or its simple field) are direct children of its first
    /// paragraph, its end marker of its last, and both paragraphs share a
    /// parent.
    #[must_use]
    pub fn accepts_paragraphs(&self, field: usize) -> bool {
        self.toc_shape(field).is_ok()
    }

    /// The bytes and run properties a text update rewrites.
    fn text_region(&self, index: usize) -> std::result::Result<TextRegion, Refusal> {
        let field = self.fields.get(index).ok_or(Refusal::FieldNotFound)?;
        if field.begin_paragraph != field.end_paragraph || field.nested_paragraph {
            return Err(Refusal::FieldInterior);
        }
        if let Some(simple) = &field.simple {
            let Some(close_start) = simple.close_start else {
                return Ok(TextRegion::Expand(simple.clone()));
            };
            let region = simple.start_tag_end..close_start;
            self.check_plain(&region)?;
            return Ok(TextRegion::Replace {
                region,
                properties: None,
            });
        }
        let end = field.end.as_ref().ok_or(Refusal::FieldInterior)?;
        if !field.closed || end.parent != field.begin.parent {
            return Err(Refusal::FieldInterior);
        }
        match &field.separate {
            Some(separate) => {
                if separate.parent != field.begin.parent {
                    return Err(Refusal::FieldInterior);
                }
                let region = separate.run.end..end.run.start;
                self.check_plain(&region)?;
                Ok(TextRegion::Replace {
                    region,
                    properties: field.begin.properties.clone(),
                })
            },
            None => Ok(TextRegion::Insert {
                at: end.run.start,
                properties: field.begin.properties.clone(),
            }),
        }
    }

    fn check_plain(&self, region: &Range<usize>) -> std::result::Result<(), Refusal> {
        let bytes = self.xml.get(region.clone()).ok_or(Refusal::FieldInterior)?;
        if bytes.iter().all(u8::is_ascii_whitespace) {
            return Ok(());
        }
        scan_run_region(bytes).map(|_| ())
    }

    /// The paragraphs and cut points a paragraph update rewrites.
    fn toc_shape(&self, index: usize) -> std::result::Result<TocShape, Refusal> {
        let field = self.fields.get(index).ok_or(Refusal::FieldNotFound)?;
        if field.nested_paragraph || !field.closed {
            return Err(Refusal::FieldInterior);
        }
        let first = self
            .paragraphs
            .get(field.begin_paragraph)
            .ok_or(Refusal::FieldInterior)?;
        let last = self
            .paragraphs
            .get(field.end_paragraph)
            .ok_or(Refusal::FieldInterior)?;
        if first.parent != last.parent || first.nested {
            return Err(Refusal::FieldInterior);
        }
        if let Some(simple) = &field.simple {
            if simple.parent != first.range.start {
                return Err(Refusal::FieldInterior);
            }
            return Ok(TocShape {
                cut: simple.element.start,
                resume: simple.element.end,
                simple: true,
            });
        }
        let end = field.end.as_ref().ok_or(Refusal::FieldInterior)?;
        if field.begin.parent != first.range.start || end.parent != last.range.start {
            return Err(Refusal::FieldInterior);
        }
        let cut = match &field.separate {
            Some(separate) if separate.parent == first.range.start => separate.run.end,
            Some(_) => return Err(Refusal::FieldInterior),
            None if field.begin_paragraph == field.end_paragraph => end.run.start,
            None => return Err(Refusal::FieldInterior),
        };
        Ok(TocShape {
            cut,
            resume: end.run.start,
            simple: false,
        })
    }
}

impl std::fmt::Debug for Fields {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Fields")
            .field("paragraphs", &self.paragraphs.len())
            .field("fields", &self.fields.len())
            .field("bookmarks", &self.bookmarks.len())
            .finish_non_exhaustive()
    }
}

enum TextRegion {
    Replace {
        region: Range<usize>,
        properties: Option<Range<usize>>,
    },
    Insert {
        at: usize,
        properties: Option<Range<usize>>,
    },
    Expand(Simple),
}

struct TocShape {
    /// The end of the first paragraph's kept content.
    cut: usize,
    /// The start of the last paragraph's kept content.
    resume: usize,
    simple: bool,
}

/// One entry of a table of contents.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TocEntry {
    level: u8,
    text: String,
    anchor: Option<String>,
    hyperlink: bool,
    page: Option<String>,
    separator: Option<String>,
    tab: Option<u32>,
}

impl TocEntry {
    /// An entry at `level` (1 through 9) with `text`, without a page
    /// number, link or tab stop.
    #[must_use]
    pub fn new(level: u8, text: impl Into<String>) -> Self {
        Self {
            level,
            text: text.into(),
            anchor: None,
            hyperlink: false,
            page: None,
            separator: None,
            tab: None,
        }
    }

    /// The bookmark the entry refers to: its `PAGEREF` and link target.
    #[must_use]
    pub fn anchor(mut self, anchor: impl Into<String>) -> Self {
        self.anchor = Some(anchor.into());
        self
    }

    /// Wrap the entry in an internal hyperlink to its anchor.
    #[must_use]
    pub const fn hyperlink(mut self, hyperlink: bool) -> Self {
        self.hyperlink = hyperlink;
        self
    }

    /// Give the entry a page number with this cached text: a `PAGEREF`
    /// field to the anchor, or plain text without one.
    #[must_use]
    pub fn page(mut self, page: impl Into<String>) -> Self {
        self.page = Some(page.into());
        self
    }

    /// Separate the page number by this text instead of a tab.
    #[must_use]
    pub fn separator(mut self, separator: impl Into<String>) -> Self {
        self.separator = Some(separator.into());
        self
    }

    /// A right-aligned, dot-leader tab stop at this position (twips) for
    /// the page number.
    #[must_use]
    pub const fn tab(mut self, position: u32) -> Self {
        self.tab = Some(position);
        self
    }

    /// The level.
    #[must_use]
    pub const fn level(&self) -> u8 {
        self.level
    }

    /// The entry text.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// The page number text.
    #[must_use]
    pub fn page_text(&self) -> Option<&str> {
        self.page.as_deref()
    }

    /// The anchor.
    #[must_use]
    pub fn anchor_name(&self) -> Option<&str> {
        self.anchor.as_deref()
    }
}

/// One change of [`Edit::refresh_fields`].
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Update {
    /// Replace a field's cached result with text.
    Text {
        /// The field index.
        field: usize,
        /// The new result; a tab is written as `w:tab` and a line feed as
        /// `w:br`.
        text: String,
    },
    /// Replace a table of contents' result with entry paragraphs.
    Toc {
        /// The field index.
        field: usize,
        /// The entries, in order.
        entries: Vec<TocEntry>,
        /// Mark the field dirty, so a word processor recomputes it on open.
        dirty: bool,
    },
    /// Bookmark a paragraph's content under a new name.
    Bookmark {
        /// The paragraph index.
        paragraph: usize,
        /// The bookmark name.
        name: String,
    },
    /// Replace a field's instruction and mark it for update (`w:dirty`), as
    /// OfficeCLI's field `set` does; the cached result is kept (record
    /// 0325).
    Instruction {
        /// The field index.
        field: usize,
        /// The new instruction, written as one `w:instrText` (or the simple
        /// field's `w:instr`).
        instruction: String,
    },
    /// Rename a bookmark (record 0325).
    RenameBookmark {
        /// The bookmark index.
        bookmark: usize,
        /// The new name.
        name: String,
    },
    /// Remove a bookmark's start and end markers, keeping what lies between
    /// them (record 0325).
    RemoveBookmark {
        /// The bookmark index.
        bookmark: usize,
    },
}

/// Content-free counts of one refresh.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct Outcome {
    fields: usize,
    bookmarks: usize,
    blocks: usize,
}

impl Outcome {
    /// Fields whose result changed.
    #[must_use]
    pub const fn fields(self) -> usize {
        self.fields
    }

    /// Bookmarks added, renamed or removed.
    #[must_use]
    pub const fn bookmarks(self) -> usize {
        self.bookmarks
    }

    /// Body children replaced, removed or inserted.
    #[must_use]
    pub const fn blocks(self) -> usize {
        self.blocks
    }

    /// Whether nothing changed.
    #[must_use]
    pub const fn is_empty(self) -> bool {
        self.fields == 0 && self.bookmarks == 0
    }
}

// ---------------------------------------------------------------------------
// Scanning
// ---------------------------------------------------------------------------

fn xml_error(message: impl Into<String>) -> Error {
    Error::Document(crate::Error::Xml(message.into()))
}

fn refused(position: usize, reason: Refusal) -> Error {
    Error::Refused { position, reason }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum State {
    Instruction,
    Result,
}

struct OpenParagraph {
    index: usize,
    depth: usize,
}

struct OpenRun {
    start: usize,
    depth: usize,
    parent: usize,
    properties: Option<Range<usize>>,
    properties_start: Option<usize>,
    /// `(field, marker kind, fldChar element)`: 0 begin, 1 separate, 2 end.
    markers: Vec<(usize, u8, Option<Range<usize>>)>,
}

struct Scanner<'x> {
    xml: &'x [u8],
    depth: usize,
    body_depth: Option<usize>,
    counts: [usize; 3],
    stack: Vec<usize>,
    paragraphs: Vec<OpenParagraph>,
    runs: Vec<OpenRun>,
    text: Option<usize>,
    instruction: Option<usize>,
    complex: Vec<(usize, State)>,
    simple: Vec<(usize, usize)>,
    skip: Option<usize>,
    text_boxes: usize,
    /// Open bookmarks: their `w:id` and index.
    open_bookmarks: Vec<(String, usize)>,
    section: Option<usize>,
    page: (Option<u32>, Option<u32>, Option<u32>),
    out: Fields,
}

/// Index the main-document XML.
pub(crate) fn scan(xml: Arc<Vec<u8>>) -> Result<Fields> {
    std::str::from_utf8(&xml).map_err(|_| xml_error("document XML is not UTF-8"))?;
    let bytes: &[u8] = &xml;
    let mut scanner = Scanner {
        xml: bytes,
        depth: 0,
        body_depth: None,
        counts: [0; 3],
        stack: Vec::new(),
        paragraphs: Vec::new(),
        runs: Vec::new(),
        text: None,
        instruction: None,
        complex: Vec::new(),
        simple: Vec::new(),
        skip: None,
        text_boxes: 0,
        open_bookmarks: Vec::new(),
        section: None,
        page: (None, None, None),
        out: Fields {
            xml: Arc::clone(&xml),
            paragraphs: Vec::new(),
            fields: Vec::new(),
            bookmarks: Vec::new(),
            children: Vec::new(),
            text_width: None,
        },
    };
    let mut reader = NsReader::from_reader(bytes);
    let origin = ReaderOrigin::of(bytes);
    loop {
        let start = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| xml_error("document offset overflow"))?;
        let event = reader
            .read_event()
            .map_err(|error| xml_error(format!("invalid document XML: {error}")))?;
        let end = origin
            .offset(reader.buffer_position())
            .ok_or_else(|| xml_error("document offset overflow"))?;
        let (namespace, event) = reader.resolver().resolve_event(event);
        let word = is_wordprocessing_namespace(&namespace);
        let compatibility =
            matches!(&namespace, ResolveResult::Bound(ns) if ns.as_ref() == MARKUP_COMPATIBILITY);
        match event {
            Event::Start(element) => {
                scanner.depth += 1;
                if scanner.depth > MAX_DOCUMENT_DEPTH {
                    return Err(xml_error("document XML is too deep"));
                }
                scanner.stack.push(start);
                if scanner.skip.is_none() {
                    if compatibility && element.local_name().as_ref() == b"Fallback" {
                        scanner.skip = Some(scanner.depth);
                    } else {
                        scanner.open(&reader, &element, word, start, end, false)?;
                    }
                }
            },
            Event::Empty(element) => {
                scanner.depth += 1;
                scanner.stack.push(start);
                if scanner.skip.is_none() {
                    scanner.open(&reader, &element, word, start, end, true)?;
                    scanner.close(element.local_name().as_ref(), word, start, end)?;
                }
                scanner.stack.pop();
                scanner.depth -= 1;
            },
            Event::End(element) => {
                match scanner.skip {
                    Some(depth) if depth == scanner.depth => scanner.skip = None,
                    Some(_) => {},
                    None => scanner.close(element.local_name().as_ref(), word, start, end)?,
                }
                scanner.stack.pop();
                scanner.depth = scanner
                    .depth
                    .checked_sub(1)
                    .ok_or_else(|| xml_error("unbalanced document XML"))?;
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if scanner.depth != 0 || !scanner.paragraphs.is_empty() {
        return Err(xml_error("unbalanced document XML"));
    }
    if let (Some(width), Some(left), Some(right)) = scanner.page {
        scanner.out.text_width = width.checked_sub(left).and_then(|w| w.checked_sub(right));
    }
    let mut out = scanner.out;
    for field in &mut out.fields {
        let trimmed = field.result.trim_matches('\n');
        if trimmed.len() != field.result.len() {
            field.result = trimmed.to_owned();
        }
    }
    Ok(out)
}

fn attribute(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    name: &[u8],
) -> Result<Option<String>> {
    for attribute in element.checked_attributes() {
        let attribute =
            attribute.map_err(|error| xml_error(format!("invalid attribute: {error}")))?;
        if attribute.key.local_name().as_ref() != name {
            continue;
        }
        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
        if !is_wordprocessing_namespace(&namespace) && !matches!(namespace, ResolveResult::Unbound)
        {
            continue;
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, reader.decoder())
            .map_err(|error| xml_error(format!("invalid attribute: {error}")))?;
        return Ok(Some(value.into_owned()));
    }
    Ok(None)
}

fn on(value: Option<&str>) -> bool {
    matches!(value, Some("1" | "true" | "on"))
}

fn twips(value: Option<String>) -> Option<u32> {
    value.and_then(|value| value.trim().parse::<u32>().ok())
}

impl Scanner<'_> {
    fn parent_start(&self) -> usize {
        self.stack
            .len()
            .checked_sub(2)
            .and_then(|index| self.stack.get(index))
            .copied()
            .unwrap_or_default()
    }

    fn current_paragraph(&self) -> Option<usize> {
        self.paragraphs.last().map(|paragraph| paragraph.index)
    }

    fn run_child(&self) -> bool {
        self.runs
            .last()
            .is_some_and(|run| run.depth + 1 == self.depth)
            && self.paragraphs.last().is_some_and(|paragraph| {
                self.runs
                    .last()
                    .is_some_and(|run| run.depth > paragraph.depth)
            })
    }

    /// Text visible in the current paragraph: not inside an instruction.
    fn push_text(&mut self, value: &str) {
        if self
            .complex
            .iter()
            .any(|(_, state)| *state == State::Instruction)
        {
            return;
        }
        if let Some(index) = self.current_paragraph() {
            self.out.paragraphs[index].text.push_str(value);
        }
        for (field, state) in &self.complex {
            if *state == State::Result {
                self.out.fields[*field].result.push_str(value);
            }
        }
        for (field, _) in &self.simple {
            self.out.fields[*field].result.push_str(value);
        }
        for (_, bookmark) in &self.open_bookmarks {
            let text = &mut self.out.bookmarks[*bookmark].text;
            if text.len() < MAX_BOOKMARK_TEXT {
                text.push_str(value);
            }
        }
    }

    fn new_field(&mut self, field: Field) -> Result<usize> {
        if self.out.fields.len() >= MAX_FIELDS {
            return Err(Error::Limit {
                kind: crate::document::LimitKind::Operations,
                maximum: MAX_FIELDS,
                observed: self.out.fields.len() + 1,
            });
        }
        if self.complex.len() + self.simple.len() >= MAX_FIELD_NESTING {
            return Err(xml_error("fields nest too deeply"));
        }
        self.out.fields.push(field);
        Ok(self.out.fields.len() - 1)
    }

    /// The innermost open field.
    fn open_field(&self) -> Option<usize> {
        let complex = self.complex.last().map(|(field, _)| *field);
        let simple = self.simple.last().map(|(field, _)| *field);
        complex.max(simple)
    }

    fn open(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        word: bool,
        start: usize,
        end: usize,
        empty: bool,
    ) -> Result<()> {
        let local = element.local_name();
        let local = local.as_ref();
        let depth = self.depth;
        if word && self.body_depth.is_none() && local == b"body" && depth == 2 {
            self.body_depth = Some(depth);
            return Ok(());
        }
        if self.body_depth.is_some_and(|body| depth == body + 1) {
            let step = match (word, local) {
                (true, b"p") => Some(ElementStep::Paragraph(next(&mut self.counts[0]))),
                (true, b"tbl") => Some(ElementStep::Table(next(&mut self.counts[1]))),
                (true, b"sdt") => Some(ElementStep::ContentControl(next(&mut self.counts[2]))),
                _ => None,
            };
            self.out.children.push(BodyChild {
                step,
                range: start..start,
            });
            if word && local == b"sectPr" {
                self.section = Some(depth);
            }
        }
        if self.body_depth.is_none() || !word {
            return Ok(());
        }
        if let Some(section) = self.section
            && depth == section + 1
        {
            match local {
                b"pgSz" => self.page.0 = twips(attribute(reader, element, b"w")?),
                b"pgMar" => {
                    self.page.1 = twips(attribute(reader, element, b"left")?);
                    self.page.2 = twips(attribute(reader, element, b"right")?);
                },
                _ => {},
            }
            return Ok(());
        }
        match local {
            b"p" => {
                let index = self.out.paragraphs.len();
                let nested = !self.paragraphs.is_empty() || self.text_boxes > 0;
                let prefix = match element.name().prefix() {
                    Some(prefix) => format!("{}:", String::from_utf8_lossy(prefix.as_ref())),
                    None => String::new(),
                };
                let block = (!nested)
                    .then(|| self.out.children.len().checked_sub(1))
                    .flatten();
                let body_position = match block.map(|block| &self.out.children[block]) {
                    Some(BodyChild {
                        step: Some(ElementStep::Paragraph(position)),
                        range,
                    }) if range.start == start => Some(position.get()),
                    _ => None,
                };
                self.out.paragraphs.push(Paragraph {
                    style: None,
                    outline: None,
                    text: String::new(),
                    anchor: None,
                    nested,
                    range: start..end,
                    start_tag_end: end,
                    properties: None,
                    close_start: start,
                    parent: self.parent_start(),
                    block,
                    prefix,
                    body_position,
                });
                self.paragraphs.push(OpenParagraph { index, depth });
            },
            b"txbxContent" => self.text_boxes += 1,
            b"pPr" => {
                if let Some(paragraph) = self.paragraphs.last()
                    && paragraph.depth + 1 == depth
                {
                    self.out.paragraphs[paragraph.index].properties = Some(start..end);
                }
            },
            b"pStyle" | b"outlineLvl" => {
                if let Some(paragraph) = self.paragraphs.last()
                    && paragraph.depth + 2 == depth
                {
                    let index = paragraph.index;
                    let value = attribute(reader, element, b"val")?;
                    let target = &mut self.out.paragraphs[index];
                    if local == b"pStyle" {
                        target.style = value;
                    } else {
                        target.outline = value
                            .and_then(|value| value.trim().parse::<u8>().ok())
                            .filter(|level| *level <= 8);
                    }
                }
            },
            b"hyperlink" => {
                if let Some(index) = self.current_paragraph()
                    && self.out.paragraphs[index].anchor.is_none()
                {
                    self.out.paragraphs[index].anchor = attribute(reader, element, b"anchor")?;
                }
            },
            b"bookmarkStart" => {
                if let Some(name) = attribute(reader, element, b"name")? {
                    let raw = attribute(reader, element, b"id")?;
                    let id = raw.as_deref().and_then(|id| id.trim().parse::<i64>().ok());
                    let paragraph = self.current_paragraph();
                    if let Some(raw) = raw
                        && self.open_bookmarks.len() < MAX_FIELD_NESTING
                    {
                        self.open_bookmarks.push((raw, self.out.bookmarks.len()));
                    }
                    self.out.bookmarks.push(Bookmark {
                        name,
                        id,
                        paragraph,
                        text: String::new(),
                        tag: start..end,
                        empty,
                        end: None,
                    });
                }
            },
            b"bookmarkEnd" => {
                if let Some(raw) = attribute(reader, element, b"id")?
                    && let Some(open) = self.open_bookmarks.iter().position(|(id, _)| *id == raw)
                {
                    let (_, index) = self.open_bookmarks.remove(open);
                    if empty && let Some(bookmark) = self.out.bookmarks.get_mut(index) {
                        bookmark.end = Some(start..end);
                    }
                }
            },
            b"r" if !self.paragraphs.is_empty() && self.text.is_none() => {
                let parent = self.parent_start();
                self.runs.push(OpenRun {
                    start,
                    depth,
                    parent,
                    properties: None,
                    properties_start: None,
                    markers: Vec::new(),
                });
            },
            b"fldSimple" if !self.paragraphs.is_empty() => {
                let instruction = attribute(reader, element, b"instr")?.unwrap_or_default();
                let paragraph = self.current_paragraph().unwrap_or_default();
                let field = Field {
                    instruction,
                    kind: FieldKind::Simple,
                    parent: self.open_field(),
                    begin_paragraph: paragraph,
                    end_paragraph: paragraph,
                    result: String::new(),
                    separated: true,
                    closed: false,
                    locked: on(attribute(reader, element, b"fldLock")?.as_deref()),
                    dirty: on(attribute(reader, element, b"dirty")?.as_deref()),
                    nested_paragraph: self.out.paragraphs[paragraph].nested,
                    begin: Marker::default(),
                    separate: None,
                    end: None,
                    simple: Some(Simple {
                        element: start..end,
                        start_tag_end: end,
                        close_start: (!empty).then_some(start),
                        parent: self.parent_start(),
                    }),
                };
                let index = self.new_field(field)?;
                if empty {
                    self.out.fields[index].closed = true;
                } else {
                    self.simple.push((index, depth));
                }
            },
            _ if self.run_child() => match local {
                b"rPr" => {
                    if let Some(run) = self.runs.last_mut() {
                        run.properties_start = Some(start);
                        if empty {
                            run.properties = Some(start..end);
                        }
                    }
                },
                b"t" if !empty => self.text = Some(start),
                b"instrText" if !empty => self.instruction = Some(start),
                b"tab" | b"ptab" => self.push_text("\t"),
                b"br" | b"cr" => self.push_text(" "),
                b"noBreakHyphen" => self.push_text("\u{2011}"),
                b"fldChar" => {
                    let kind = attribute(reader, element, b"fldCharType")?;
                    let character = empty.then_some(start..end);
                    self.field_char(reader, element, kind.as_deref(), character)?;
                },
                _ => {},
            },
            _ => {},
        }
        Ok(())
    }

    fn field_char(
        &mut self,
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        kind: Option<&str>,
        character: Option<Range<usize>>,
    ) -> Result<()> {
        let Some(paragraph) = self.current_paragraph() else {
            return Ok(());
        };
        match kind {
            Some("begin") => {
                let field = Field {
                    instruction: String::new(),
                    kind: FieldKind::Complex,
                    parent: self.open_field(),
                    begin_paragraph: paragraph,
                    end_paragraph: paragraph,
                    result: String::new(),
                    separated: false,
                    closed: false,
                    locked: on(attribute(reader, element, b"fldLock")?.as_deref()),
                    dirty: on(attribute(reader, element, b"dirty")?.as_deref()),
                    nested_paragraph: self.out.paragraphs[paragraph].nested,
                    begin: Marker::default(),
                    separate: None,
                    end: None,
                    simple: None,
                };
                let index = self.new_field(field)?;
                self.complex.push((index, State::Instruction));
                if let Some(run) = self.runs.last_mut() {
                    run.markers.push((index, 0, character));
                }
            },
            Some("separate") => {
                if let Some((index, state)) = self.complex.last_mut() {
                    *state = State::Result;
                    let index = *index;
                    self.out.fields[index].separated = true;
                    if let Some(run) = self.runs.last_mut() {
                        run.markers.push((index, 1, character));
                    }
                }
            },
            Some("end") => {
                if let Some((index, _)) = self.complex.pop() {
                    let field = &mut self.out.fields[index];
                    field.end_paragraph = paragraph;
                    field.closed = true;
                    if let Some(run) = self.runs.last_mut() {
                        run.markers.push((index, 2, character));
                    }
                }
            },
            _ => {},
        }
        Ok(())
    }

    fn close(&mut self, local: &[u8], word: bool, start: usize, end: usize) -> Result<()> {
        if word {
            self.close_word(local, start, end)?;
        }
        if self.body_depth.is_some_and(|body| self.depth == body + 1)
            && let Some(child) = self.out.children.last_mut()
        {
            child.range.end = end;
            if self.section == Some(self.depth) {
                self.section = None;
            }
        }
        Ok(())
    }

    fn close_word(&mut self, local: &[u8], start: usize, end: usize) -> Result<()> {
        let depth = self.depth;
        match local {
            b"t" if self.text.is_some() => {
                let Some(open) = self.text.take() else {
                    return Ok(());
                };
                let value = decode_text_fragment(
                    self.xml
                        .get(open..end)
                        .ok_or_else(|| xml_error("text offset is out of range"))?,
                )
                .map_err(|_| xml_error("invalid text"))?;
                self.push_text(&value);
            },
            b"instrText" if self.instruction.is_some() => {
                let Some(open) = self.instruction.take() else {
                    return Ok(());
                };
                let value = decode_text_fragment(
                    self.xml
                        .get(open..end)
                        .ok_or_else(|| xml_error("instruction offset is out of range"))?,
                )
                .map_err(|_| xml_error("invalid field instruction"))?;
                if let Some((index, State::Instruction)) = self.complex.last() {
                    self.out.fields[*index].instruction.push_str(&value);
                }
            },
            b"rPr" if self.runs.last().is_some_and(|run| run.depth + 1 == depth) => {
                if let Some(run) = self.runs.last_mut()
                    && let Some(properties_start) = run.properties_start.take()
                {
                    run.properties = Some(properties_start..end);
                }
            },
            b"r" if self.runs.last().is_some_and(|run| run.depth == depth) => {
                let Some(run) = self.runs.pop() else {
                    return Ok(());
                };
                for (field, kind, character) in run.markers {
                    let marker = Marker {
                        run: run.start..end,
                        parent: run.parent,
                        properties: run.properties.clone(),
                        character,
                    };
                    let field = &mut self.out.fields[field];
                    match kind {
                        0 => field.begin = marker,
                        1 => field.separate = Some(marker),
                        _ => field.end = Some(marker),
                    }
                }
            },
            b"fldSimple" if self.simple.last().is_some_and(|(_, d)| *d == depth) => {
                if let Some((index, _)) = self.simple.pop() {
                    let field = &mut self.out.fields[index];
                    field.closed = true;
                    if let Some(simple) = &mut field.simple {
                        simple.element.end = end;
                        simple.close_start = Some(start);
                    }
                }
            },
            b"pPr" => {
                if let Some(paragraph) = self.paragraphs.last()
                    && paragraph.depth + 1 == depth
                    && let Some(properties) = &mut self.out.paragraphs[paragraph.index].properties
                {
                    properties.end = end;
                }
            },
            b"txbxContent" => self.text_boxes = self.text_boxes.saturating_sub(1),
            b"p" if self
                .paragraphs
                .last()
                .is_some_and(|paragraph| paragraph.depth == depth) =>
            {
                let Some(open) = self.paragraphs.pop() else {
                    return Ok(());
                };
                let paragraph = &mut self.out.paragraphs[open.index];
                paragraph.range.end = end;
                paragraph.close_start = start;
                for (field, state) in &self.complex {
                    if *state == State::Result {
                        self.out.fields[*field].result.push('\n');
                    }
                }
            },
            _ => {},
        }
        Ok(())
    }
}

fn next(counter: &mut usize) -> Position {
    let position = Position::new(*counter);
    *counter += 1;
    position
}

// ---------------------------------------------------------------------------
// Writing
// ---------------------------------------------------------------------------

fn escape(output: &mut String, text: &str) {
    for character in text.chars() {
        match character {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            '"' => output.push_str("&quot;"),
            _ => output.push(character),
        }
    }
}

fn check_text(text: &str) -> std::result::Result<(), Refusal> {
    if text.chars().all(|character| {
        matches!(character, '\t' | '\n' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}')
    }) {
        Ok(())
    } else {
        Err(Refusal::StructuralText)
    }
}

/// Run content for `text`: `w:t` pieces, `w:tab` and `w:br`.
fn run_content(output: &mut String, prefix: &str, text: &str) {
    let mut piece = String::new();
    let flush = |output: &mut String, piece: &mut String| {
        if !piece.is_empty() {
            output.push_str(&format!("<{prefix}t xml:space=\"preserve\">"));
            escape(output, piece);
            output.push_str(&format!("</{prefix}t>"));
            piece.clear();
        }
    };
    for character in text.chars() {
        match character {
            '\t' => {
                flush(output, &mut piece);
                output.push_str(&format!("<{prefix}tab/>"));
            },
            '\n' => {
                flush(output, &mut piece);
                output.push_str(&format!("<{prefix}br/>"));
            },
            _ => piece.push(character),
        }
    }
    flush(output, &mut piece);
}

fn text_run(prefix: &str, properties: &[u8], text: &str) -> String {
    let mut output = format!("<{prefix}r>");
    output.push_str(&String::from_utf8_lossy(properties));
    run_content(&mut output, prefix, text);
    output.push_str(&format!("</{prefix}r>"));
    output
}

fn marker_run(prefix: &str, kind: &str, dirty: bool) -> String {
    let dirty = if dirty {
        format!(" {prefix}dirty=\"true\"")
    } else {
        String::new()
    };
    format!("<{prefix}r><{prefix}fldChar {prefix}fldCharType=\"{kind}\"{dirty}/></{prefix}r>")
}

fn instruction_run(prefix: &str, instruction: &str) -> String {
    let mut output = format!("<{prefix}r><{prefix}instrText xml:space=\"preserve\">");
    escape(&mut output, instruction);
    output.push_str(&format!("</{prefix}instrText></{prefix}r>"));
    output
}

fn entry_paragraph(prefix: &str, entry: &TocEntry) -> std::result::Result<String, Refusal> {
    check_text(&entry.text)?;
    let level = entry.level.clamp(1, 9);
    let mut output =
        format!("<{prefix}p><{prefix}pPr><{prefix}pStyle {prefix}val=\"TOC{level}\"/>");
    let tab = entry.page.is_some() && entry.separator.is_none();
    if let (true, Some(position)) = (tab, entry.tab) {
        output.push_str(&format!(
            "<{prefix}tabs><{prefix}tab {prefix}val=\"right\" {prefix}leader=\"dot\" {prefix}pos=\"{position}\"/></{prefix}tabs>"
        ));
    }
    output.push_str(&format!("</{prefix}pPr>"));
    let link = entry.hyperlink && entry.anchor.is_some();
    if let (true, Some(anchor)) = (link, &entry.anchor) {
        output.push_str(&format!("<{prefix}hyperlink {prefix}anchor=\""));
        escape(&mut output, anchor);
        output.push_str(&format!("\" {prefix}history=\"1\">"));
    }
    output.push_str(&text_run(prefix, &[], &entry.text));
    if let Some(page) = &entry.page {
        check_text(page)?;
        match &entry.separator {
            Some(separator) => {
                check_text(separator)?;
                output.push_str(&text_run(prefix, &[], separator));
            },
            None => output.push_str(&format!("<{prefix}r><{prefix}tab/></{prefix}r>")),
        }
        match &entry.anchor {
            Some(anchor) => {
                output.push_str(&marker_run(prefix, "begin", false));
                output.push_str(&instruction_run(prefix, &format!(" PAGEREF {anchor} \\h ")));
                output.push_str(&marker_run(prefix, "separate", false));
                if !page.is_empty() {
                    output.push_str(&text_run(prefix, &[], page));
                }
                output.push_str(&marker_run(prefix, "end", false));
            },
            None => output.push_str(&text_run(prefix, &[], page)),
        }
    }
    if link {
        output.push_str(&format!("</{prefix}hyperlink>"));
    }
    output.push_str(&format!("</{prefix}p>"));
    Ok(output)
}

/// The markup of a new bookmark around optional `text`: its start marker,
/// a plain run, and its end marker (record 0325).
fn bookmark_markup(prefix: &str, id: i64, name: &str, text: Option<&str>) -> Result<Vec<u8>> {
    check_text(name).map_err(|reason| refused(0, reason))?;
    if name.is_empty() || name.chars().count() > 40 {
        return Err(refused(0, Refusal::InvalidPropertyValue));
    }
    let mut out = format!("<{prefix}bookmarkStart {prefix}id=\"{id}\" {prefix}name=\"");
    escape(&mut out, name);
    out.push_str("\"/>");
    if let Some(text) = text.filter(|text| !text.is_empty()) {
        check_text(text).map_err(|reason| refused(0, reason))?;
        out.push_str(&text_run(prefix, &[], text));
    }
    out.push_str(&format!("<{prefix}bookmarkEnd {prefix}id=\"{id}\"/>"));
    Ok(out.into_bytes())
}

/// The markup of a new complex field: its begin, instruction, separate,
/// result and end runs, each carrying `properties` (record 0325).
fn field_markup(
    prefix: &str,
    properties: &[u8],
    instruction: &str,
    result: Option<&str>,
) -> Result<Vec<u8>> {
    check_text(instruction).map_err(|reason| refused(0, reason))?;
    if instruction.trim().is_empty() {
        return Err(refused(0, Refusal::InvalidPropertyValue));
    }
    let properties = String::from_utf8_lossy(properties);
    let marker = |kind: &str| {
        format!(
            "<{prefix}r>{properties}<{prefix}fldChar {prefix}fldCharType=\"{kind}\"/></{prefix}r>"
        )
    };
    let mut out = marker("begin");
    out.push_str(&format!(
        "<{prefix}r>{properties}<{prefix}instrText xml:space=\"preserve\">"
    ));
    escape(&mut out, instruction);
    out.push_str(&format!("</{prefix}instrText></{prefix}r>"));
    out.push_str(&marker("separate"));
    if let Some(result) = result.filter(|result| !result.is_empty()) {
        check_text(result).map_err(|reason| refused(0, reason))?;
        out.push_str(&text_run(prefix, properties.as_bytes(), result));
    }
    out.push_str(&marker("end"));
    Ok(out.into_bytes())
}

/// A namespace prefix as the markup builders take it: `w:`, or empty.
fn qualifier(prefix: &[u8]) -> String {
    if prefix.is_empty() {
        String::new()
    } else {
        format!("{}:", String::from_utf8_lossy(prefix))
    }
}

/// A start tag (or empty element) with attribute `local` (in the tag's
/// own prefix, as WordprocessingML writes its attributes) set to `value`;
/// `None` for a tag this cannot rewrite.
fn with_attribute(tag: &[u8], local: &[u8], value: &str) -> Option<Vec<u8>> {
    let text = std::str::from_utf8(tag).ok()?;
    let body = text.strip_prefix('<')?;
    let name_end = body.find(|c: char| c.is_whitespace() || c == '/' || c == '>')?;
    let name = &body[..name_end];
    let prefix = name.split_once(':').map_or("", |(prefix, _)| prefix);
    let local = std::str::from_utf8(local).ok()?;
    let qualified = if prefix.is_empty() {
        local.to_owned()
    } else {
        format!("{prefix}:{local}")
    };
    let mut escaped = String::new();
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\t' => escaped.push_str("&#9;"),
            '\n' => escaped.push_str("&#10;"),
            '\r' => escaped.push_str("&#13;"),
            other => escaped.push(other),
        }
    }
    for quote in ['"', '\''] {
        let needle = format!(" {qualified}={quote}");
        if let Some(at) = text.find(&needle) {
            let value_start = at + needle.len();
            let value_end = value_start + text[value_start..].find(quote)?;
            let mut out = String::with_capacity(text.len() + escaped.len());
            out.push_str(&text[..value_start - 1]);
            out.push('"');
            out.push_str(&escaped);
            out.push('"');
            out.push_str(&text[value_end + 1..]);
            return Some(out.into_bytes());
        }
    }
    let close = if text.ends_with("/>") {
        text.len() - 2
    } else {
        text.len().checked_sub(1)?
    };
    let head = text[..close].trim_end();
    Some(format!("{head} {qualified}=\"{escaped}\"{}", &text[close..]).into_bytes())
}

/// Add `dirty="true"` to an empty `fldChar` element.
fn dirty_marker(element: &[u8], prefix: &str) -> Vec<u8> {
    let text = String::from_utf8_lossy(element);
    if text.contains("dirty=") {
        return element.to_vec();
    }
    let body = text.trim_end_matches("/>").trim_end();
    format!("{body} {prefix}dirty=\"true\"/>").into_bytes()
}

/// A body-level replacement of children `first..=last`.
struct Region {
    first: usize,
    last: usize,
    blocks: Vec<Vec<u8>>,
}

impl Edit {
    /// Index the paragraphs, fields and bookmarks of the staged main
    /// document.
    ///
    /// # Errors
    ///
    /// Returns an XML error for malformed markup and a limit error past
    /// 262,144 fields.
    pub fn fields(&self) -> Result<Fields> {
        scan(Arc::new(self.projected.xml_bytes().to_vec()))
    }

    /// Insert a bookmark into the paragraph at `paragraph`, before its
    /// direct run `at` (after its last child for `None`): its start marker,
    /// a plain run of `text` when given, and its end marker (record 0325).
    /// The bookmark takes the identifier after the largest this document
    /// uses, which it returns.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an empty name or one longer than 40
    /// characters, text a document cannot hold, a missing paragraph or run,
    /// or a position inside a field, and an XML error for malformed markup.
    /// The edit is unchanged on failure.
    pub fn insert_bookmark_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        name: &str,
        text: Option<&str>,
    ) -> Result<i64> {
        let id = self
            .fields()?
            .bookmarks
            .iter()
            .filter_map(Bookmark::id)
            .max()
            .map_or(0, |id| id.saturating_add(1).max(0));
        self.insert_run_markup_at(paragraph, at, |prefix| {
            bookmark_markup(&qualifier(prefix), id, name, text)
        })?;
        Ok(id)
    }

    /// Insert a complex field (`begin`, `instrText`, `separate`, the result
    /// and `end` runs) into the paragraph at `paragraph`, before its direct
    /// run `at` (after its last child for `None`), and return the position
    /// of its begin run among the paragraph's direct runs (record 0325).
    /// `result` is the cached result; the instruction is written as given.
    ///
    /// # Errors
    ///
    /// Returns a typed refusal for an empty instruction, text a document
    /// cannot hold, a missing paragraph or run, or a position inside a
    /// field. The edit is unchanged on failure.
    pub fn insert_field_at(
        &mut self,
        paragraph: &ElementPath,
        at: Option<Position>,
        instruction: &str,
        result: Option<&str>,
    ) -> Result<Position> {
        self.insert_run_markup_at(paragraph, at, |prefix| {
            field_markup(&qualifier(prefix), &[], instruction, result)
        })
    }

    /// Apply `updates` planned over [`Self::fields`] of this edit. See
    /// [`crate::document::fields`] for the rewriting rules. All updates are
    /// spliced once; changed body children are recorded as exact element
    /// replacements, and a table of contents whose paragraphs are direct
    /// body children as exact removals and insertions.
    ///
    /// # Errors
    ///
    /// Returns [`Error::PatchConflict`] for an index of other bytes, a refusal
    /// for an update the field's markup does not admit (see
    /// [`Fields::accepts_text`] and [`Fields::accepts_paragraphs`]), text a
    /// document cannot hold or overlapping updates, and a limit error. The
    /// edit is unchanged on failure.
    pub fn refresh_fields(&mut self, fields: &Fields, updates: &[Update]) -> Result<Outcome> {
        self.ensure_unmanaged("refresh_fields")?;
        if self.projected.xml_bytes() != fields.xml.as_slice() {
            return Err(Error::PatchConflict);
        }
        let xml = fields.xml.as_slice();
        let mut outcome = Outcome::default();
        let mut splices: Vec<(usize, usize, Vec<u8>)> = Vec::new();
        let mut regions: Vec<Region> = Vec::new();
        let mut next_id = fields
            .bookmarks
            .iter()
            .filter_map(Bookmark::id)
            .max()
            .unwrap_or(-1)
            .max(-1)
            + 1;
        for update in updates {
            match update {
                Update::Text { field, text } => {
                    check_text(text).map_err(|reason| refused(*field, reason))?;
                    let paragraph = &fields.paragraphs[fields.fields[*field].begin_paragraph];
                    let prefix = paragraph.prefix.as_str();
                    let region = fields
                        .text_region(*field)
                        .map_err(|reason| refused(*field, reason))?;
                    let properties = |range: &Option<Range<usize>>| {
                        range.as_ref().map_or(&[][..], |range| &xml[range.clone()])
                    };
                    match region {
                        TextRegion::Replace {
                            region,
                            properties: begin,
                        } => {
                            let bytes = &xml[region.clone()];
                            let replacement = if bytes.iter().all(u8::is_ascii_whitespace) {
                                if text.is_empty() {
                                    continue;
                                }
                                text_run(prefix, properties(&begin), text).into_bytes()
                            } else if text.contains(['\t', '\n']) {
                                text_run(prefix, properties(&begin), text).into_bytes()
                            } else {
                                let (before, replacement) = rewrite_run_region(bytes, text)
                                    .map_err(|reason| refused(*field, reason))?;
                                if before == *text {
                                    continue;
                                }
                                replacement
                            };
                            splices.push((region.start, region.end, replacement));
                        },
                        TextRegion::Insert {
                            at,
                            properties: begin,
                        } => {
                            let mut inserted = marker_run(prefix, "separate", false);
                            if !text.is_empty() {
                                inserted.push_str(&text_run(prefix, properties(&begin), text));
                            }
                            splices.push((at, at, inserted.into_bytes()));
                        },
                        TextRegion::Expand(simple) => {
                            if text.is_empty() {
                                continue;
                            }
                            let tag = &xml[simple.element.clone()];
                            let tag = String::from_utf8_lossy(tag);
                            let open = tag.trim_end_matches("/>").trim_end();
                            let name_end = open[1..]
                                .find(|c: char| c.is_whitespace())
                                .map_or(open.len(), |offset| offset + 1);
                            let name = &open[1..name_end];
                            let element =
                                format!("{open}>{}</{name}>", text_run(prefix, &[], text));
                            splices.push((
                                simple.element.start,
                                simple.element.end,
                                element.into_bytes(),
                            ));
                        },
                    }
                    outcome.fields += 1;
                },
                Update::Toc {
                    field: index,
                    entries,
                    dirty,
                } => {
                    let shape = fields
                        .toc_shape(*index)
                        .map_err(|reason| refused(*index, reason))?;
                    let field = &fields.fields[*index];
                    let first = &fields.paragraphs[field.begin_paragraph];
                    let last = &fields.paragraphs[field.end_paragraph];
                    let prefix = first.prefix.as_str();
                    // The first paragraph, through the separator.
                    let mut head = Vec::new();
                    if shape.simple {
                        head.extend_from_slice(&xml[first.range.start..shape.cut]);
                        head.extend_from_slice(marker_run(prefix, "begin", *dirty).as_bytes());
                        head.extend_from_slice(
                            instruction_run(prefix, &field.instruction).as_bytes(),
                        );
                        head.extend_from_slice(marker_run(prefix, "separate", false).as_bytes());
                    } else {
                        match (&field.begin.character, dirty) {
                            (Some(character), true) => {
                                head.extend_from_slice(&xml[first.range.start..character.start]);
                                head.extend_from_slice(&dirty_marker(
                                    &xml[character.clone()],
                                    prefix,
                                ));
                                head.extend_from_slice(&xml[character.end..shape.cut]);
                            },
                            _ => head.extend_from_slice(&xml[first.range.start..shape.cut]),
                        }
                        if field.separate.is_none() {
                            head.extend_from_slice(
                                marker_run(prefix, "separate", false).as_bytes(),
                            );
                        }
                    }
                    head.extend_from_slice(format!("</{prefix}p>").as_bytes());
                    // The last paragraph, from the end marker.
                    let mut tail = Vec::new();
                    if field.begin_paragraph == field.end_paragraph {
                        tail.extend_from_slice(format!("<{prefix}p>").as_bytes());
                    } else {
                        tail.extend_from_slice(&xml[last.range.start..last.start_tag_end]);
                    }
                    if let Some(properties) = &last.properties {
                        tail.extend_from_slice(&xml[properties.clone()]);
                    }
                    if shape.simple {
                        tail.extend_from_slice(marker_run(prefix, "end", false).as_bytes());
                    }
                    tail.extend_from_slice(&xml[shape.resume..last.range.end]);
                    let mut blocks = vec![head];
                    for entry in entries {
                        blocks.push(
                            entry_paragraph(prefix, entry)
                                .map_err(|reason| refused(*index, reason))?
                                .into_bytes(),
                        );
                    }
                    blocks.push(tail);
                    let old = &xml[first.range.start..last.range.end];
                    if blocks.concat() == old {
                        continue;
                    }
                    outcome.fields += 1;
                    let body_level = first.block.is_some()
                        && fields
                            .children
                            .get(first.block.unwrap_or_default())
                            .is_some_and(|child| child.range.start == first.range.start);
                    if body_level {
                        let (Some(first_child), Some(last_child)) = (first.block, last.block)
                        else {
                            return Err(refused(*index, Refusal::FieldInterior));
                        };
                        regions.push(Region {
                            first: first_child,
                            last: last_child,
                            blocks,
                        });
                    } else {
                        splices.push((first.range.start, last.range.end, blocks.concat()));
                    }
                },
                Update::Bookmark { paragraph, name } => {
                    check_text(name).map_err(|reason| refused(*paragraph, reason))?;
                    let target = fields
                        .paragraphs
                        .get(*paragraph)
                        .ok_or_else(|| refused(*paragraph, Refusal::BlockNotFound))?;
                    if name.is_empty() || name.len() > 40 {
                        return Err(refused(*paragraph, Refusal::InvalidPropertyValue));
                    }
                    let prefix = target.prefix.as_str();
                    let id = next_id;
                    next_id += 1;
                    let at = target
                        .properties
                        .as_ref()
                        .map_or(target.start_tag_end, |properties| properties.end);
                    let mut start =
                        format!("<{prefix}bookmarkStart {prefix}id=\"{id}\" {prefix}name=\"");
                    escape(&mut start, name);
                    start.push_str("\"/>");
                    splices.push((at, at, start.into_bytes()));
                    splices.push((
                        target.close_start,
                        target.close_start,
                        format!("<{prefix}bookmarkEnd {prefix}id=\"{id}\"/>").into_bytes(),
                    ));
                    outcome.bookmarks += 1;
                },
                Update::Instruction {
                    field: index,
                    instruction,
                } => {
                    check_text(instruction).map_err(|reason| refused(*index, reason))?;
                    let field = fields
                        .fields
                        .get(*index)
                        .ok_or_else(|| refused(*index, Refusal::FieldNotFound))?;
                    if field.instruction == *instruction {
                        continue;
                    }
                    let prefix = fields
                        .paragraphs
                        .get(field.begin_paragraph)
                        .map_or("", |paragraph| paragraph.prefix.as_str());
                    if let Some(simple) = &field.simple {
                        let tag = &xml[simple.element.start..simple.start_tag_end];
                        let rewritten = with_attribute(tag, b"instr", instruction)
                            .and_then(|tag| with_attribute(&tag, b"dirty", "true"))
                            .ok_or_else(|| refused(*index, Refusal::FieldInterior))?;
                        splices.push((simple.element.start, simple.start_tag_end, rewritten));
                    } else {
                        let next = field
                            .separate
                            .as_ref()
                            .or(field.end.as_ref())
                            .ok_or_else(|| refused(*index, Refusal::FieldInterior))?;
                        // The markers must be separate runs of one parent.
                        if next.parent != field.begin.parent || field.begin.run.end > next.run.start
                        {
                            return Err(refused(*index, Refusal::FieldInterior));
                        }
                        let region = field.begin.run.end..next.run.start;
                        // A field nested in the instruction is not rewritten.
                        if xml[region.clone()]
                            .windows(b"fldChar".len())
                            .any(|window| window == b"fldChar")
                        {
                            return Err(refused(*index, Refusal::FieldInterior));
                        }
                        let properties = field
                            .begin
                            .properties
                            .as_ref()
                            .map_or(&[][..], |range| &xml[range.clone()]);
                        let mut run = format!("<{prefix}r>");
                        run.push_str(&String::from_utf8_lossy(properties));
                        run.push_str(&format!("<{prefix}instrText xml:space=\"preserve\">"));
                        escape(&mut run, instruction);
                        run.push_str(&format!("</{prefix}instrText></{prefix}r>"));
                        splices.push((region.start, region.end, run.into_bytes()));
                        if let Some(character) = &field.begin.character {
                            splices.push((
                                character.start,
                                character.end,
                                dirty_marker(&xml[character.clone()], prefix),
                            ));
                        }
                    }
                    outcome.fields += 1;
                },
                Update::RenameBookmark { bookmark, name } => {
                    check_text(name).map_err(|reason| refused(*bookmark, reason))?;
                    if name.is_empty() || name.chars().count() > 40 {
                        return Err(refused(*bookmark, Refusal::InvalidPropertyValue));
                    }
                    let target = fields
                        .bookmarks
                        .get(*bookmark)
                        .ok_or_else(|| refused(*bookmark, Refusal::BlockNotFound))?;
                    if target.name == *name {
                        continue;
                    }
                    let tag = &xml[target.tag.clone()];
                    let rewritten = with_attribute(tag, b"name", name)
                        .ok_or_else(|| refused(*bookmark, Refusal::InvalidPropertyValue))?;
                    splices.push((target.tag.start, target.tag.end, rewritten));
                    outcome.bookmarks += 1;
                },
                Update::RemoveBookmark { bookmark } => {
                    let target = fields
                        .bookmarks
                        .get(*bookmark)
                        .ok_or_else(|| refused(*bookmark, Refusal::BlockNotFound))?;
                    if !target.empty {
                        return Err(refused(*bookmark, Refusal::ComplexContent));
                    }
                    splices.push((target.tag.start, target.tag.end, Vec::new()));
                    if let Some(end) = &target.end {
                        splices.push((end.start, end.end, Vec::new()));
                    }
                    outcome.bookmarks += 1;
                },
            }
        }
        if splices.is_empty() && regions.is_empty() {
            return Ok(outcome);
        }
        splices.sort_by_key(|splice| (splice.0, splice.1));
        regions.sort_by_key(|region| region.first);
        // Group the in-place splices by body child.
        let mut replaced: Vec<(usize, Vec<(usize, usize, Vec<u8>)>)> = Vec::new();
        for (start, end, bytes) in splices {
            let child = fields
                .children
                .partition_point(|child| child.range.end <= start);
            let Some(found) = fields.children.get(child) else {
                return Err(xml_error("refresh splice lies outside the body"));
            };
            if start < found.range.start || end > found.range.end || found.step.is_none() {
                return Err(refused(0, Refusal::BlockNotFound));
            }
            if regions
                .iter()
                .any(|region| (region.first..=region.last).contains(&child))
            {
                return Err(refused(0, Refusal::OpenRange));
            }
            let local = (start - found.range.start, end - found.range.start, bytes);
            match replaced.last_mut() {
                Some((last, splices)) if *last == child => splices.push(local),
                _ => replaced.push((child, vec![local])),
            }
        }
        for pair in regions.windows(2) {
            if pair[0].last >= pair[1].first {
                return Err(refused(0, Refusal::OpenRange));
            }
        }
        let operations = replaced.len()
            + regions
                .iter()
                .map(|region| region.last - region.first + 1 + region.blocks.len())
                .sum::<usize>();
        if self.operations.len().saturating_add(operations) > MAX_OPERATIONS {
            return Err(Error::Limit {
                kind: crate::document::LimitKind::Operations,
                maximum: MAX_OPERATIONS,
                observed: self.operations.len().saturating_add(operations),
            });
        }
        // The body's children as the element operations count them.
        let body = ElementPath::new(Vec::<ElementStep>::new());
        let (children, close_start) = super::formatting::body_children(&self.projected)?;
        let mut new_operations: Vec<Operation> = Vec::with_capacity(operations);
        let mut ranges: Vec<(usize, usize, Vec<u8>)> = Vec::new();
        let mut text_bytes = self.replacement_text_bytes;
        let mut charge = |bytes: usize| -> Result<()> {
            text_bytes = text_bytes.saturating_add(bytes);
            if text_bytes > MAX_REPLACEMENT_TEXT_BYTES {
                return Err(Error::Limit {
                    kind: crate::document::LimitKind::ReplacementTextBytes,
                    maximum: MAX_REPLACEMENT_TEXT_BYTES,
                    observed: text_bytes,
                });
            }
            Ok(())
        };
        for (child, local) in &replaced {
            let found = &fields.children[*child];
            let Some(step) = found.step else {
                return Err(refused(0, Refusal::BlockNotFound));
            };
            let before = &xml[found.range.clone()];
            let after = replace_ranges(before, local)?;
            charge(after.len())?;
            ranges.push((found.range.start, found.range.end, after.clone()));
            new_operations.push(Operation::ReplaceElement {
                path: ElementPath::new([step]),
                before: Arc::new(before.to_vec()),
                after: Arc::new(after),
            });
        }
        let index_of = |start: usize| children.iter().position(|child| child.0 == start);
        let mut region_operations: Vec<Vec<Operation>> = Vec::new();
        for region in &regions {
            let first_range = &fields.children[region.first].range;
            let last_range = &fields.children[region.last].range;
            let (Some(first), Some(last)) =
                (index_of(first_range.start), index_of(last_range.start))
            else {
                return Err(Error::PatchConflict);
            };
            if last < first || last - first != region.last - region.first {
                return Err(refused(0, Refusal::BlockNotFound));
            }
            let mut list = Vec::new();
            for child in &children[first..=last] {
                let bytes = xml[child.0..child.1].to_vec();
                ranges.push((child.0, child.1, Vec::new()));
                list.push(Operation::RemoveElement {
                    container: body.clone(),
                    child: Position::new(first),
                    xml: Arc::new(bytes),
                });
            }
            let at = children.get(last + 1).map_or(close_start, |child| child.0);
            let mut inserted = Vec::new();
            for (offset, block) in region.blocks.iter().enumerate() {
                charge(block.len())?;
                inserted.extend_from_slice(block);
                list.push(Operation::InsertElement {
                    container: body.clone(),
                    child: Position::new(first + offset),
                    xml: Arc::new(block.clone()),
                });
            }
            ranges.push((at, at, inserted));
            region_operations.push(list);
        }
        ranges.sort_by_key(|range| (range.0, range.1));
        let rewritten = replace_ranges(xml, &ranges)?;
        let candidate = self.projected_with_xml(rewritten)?;
        // Regions apply last to first, so earlier child indices hold.
        for list in region_operations.into_iter().rev() {
            new_operations.extend(list);
        }
        outcome.blocks = new_operations.len();
        self.operations
            .try_reserve(new_operations.len())
            .map_err(|source| crate::Error::Allocation {
                resource: "document operation metadata",
                source,
            })?;
        self.operations.extend(new_operations);
        self.replacement_text_bytes = text_bytes;
        self.projected = candidate;
        Ok(outcome)
    }
}
