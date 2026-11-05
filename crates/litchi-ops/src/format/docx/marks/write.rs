//! Writes of the DOCX marks (record 0319): paragraph tab stops are set,
//! added and removed through the paragraph's typed properties; breaks are
//! added as a run holding one `w:br` and removed with that run.
//!
//! Bookmarks, fields and tables of contents (record 0325) are written
//! through the leaf's field index (`document::fields`): a bookmark is added
//! around a run of text (`insert_bookmark_at`), renamed and removed (its
//! markers only, keeping its content); a field is added as a complex field
//! from OfficeCLI's `fieldType` or a raw `instruction`, and its instruction
//! or cached result set; a table of contents is added as one paragraph
//! holding a `TOC` field whose result is the "Update field" placeholder
//! (the `refresh` engine computes it), its levels and switches set, and it
//! is removed with its paragraphs and title. An instruction change marks
//! the field for update (`w:dirty`), as OfficeCLI does.
//!
//! Both owners take these edits: they change only the main document.

use litchi_core::selector::{Path, Predicate, Step as PathStep};
use litchi_core::{FileFormat, Position};
use litchi_docx::RunBreakType;
use litchi_docx::document::ElementStep;
use litchi_docx::properties::{TabKind, TabStop, Twips};

use super::super::element::{element_path, path_with, refusal, run_locations};
use super::super::node::{Node, Value};
use super::super::spec::Kind;
use super::super::write::{Commit, WorkEdit};
use super::super::{Docx, DocxEditor};
use super::{
    BOOKMARK_REGISTRY, BREAK_REGISTRY, FIELD_REGISTRY, MarkSetter, MarkValue, TAB_REGISTRY,
    TOC_REGISTRY, tab_kind, tab_leader, with_stop,
};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::props::Props;
use crate::registry::Operation;
use crate::value::Scalar;
use litchi_docx::document::ElementPath;
use litchi_docx::document::fields::{Fields, Update};

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Docx)
}

/// A new mark element, by the type name `add` was given.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(in super::super) enum NewMark {
    Tab,
    Break,
    /// A bookmark (record 0325).
    Bookmark,
    /// A field: `field`, or OfficeCLI's field-type shortcut (`page`,
    /// `date`, ...), whose name is the field type (record 0325).
    Field(Option<&'static str>),
    /// A table of contents (record 0325).
    Toc,
}

/// OfficeCLI's field types and the instruction word each writes.
const FIELD_TYPES: &[(&str, &str)] = &[
    ("page", "PAGE"),
    ("pagenum", "PAGE"),
    ("pagenumber", "PAGE"),
    ("numpages", "NUMPAGES"),
    ("section", "SECTION"),
    ("sectionpages", "SECTIONPAGES"),
    ("date", "DATE"),
    ("createdate", "CREATEDATE"),
    ("savedate", "SAVEDATE"),
    ("printdate", "PRINTDATE"),
    ("time", "TIME"),
    ("edittime", "EDITTIME"),
    ("author", "AUTHOR"),
    ("lastsavedby", "LASTSAVEDBY"),
    ("title", "TITLE"),
    ("subject", "SUBJECT"),
    ("filename", "FILENAME"),
    ("numwords", "NUMWORDS"),
    ("numchars", "NUMCHARS"),
    ("revnum", "REVNUM"),
    ("template", "TEMPLATE"),
    ("keywords", "KEYWORDS"),
    ("comments", "COMMENTS"),
    ("doccomments", "COMMENTS"),
    ("mergefield", "MERGEFIELD"),
    ("ref", "REF"),
    ("pageref", "PAGEREF"),
    ("noteref", "NOTEREF"),
    ("seq", "SEQ"),
    ("styleref", "STYLEREF"),
    ("docproperty", "DOCPROPERTY"),
    ("if", "IF"),
];

/// The instruction word of field type `kind`.
fn field_word(kind: &str) -> Option<&'static str> {
    FIELD_TYPES
        .iter()
        .find(|(name, _)| name.eq_ignore_ascii_case(kind.trim()))
        .map(|(_, word)| *word)
}

/// The instruction and cached result OfficeCLI writes for a typed field;
/// a field of the date family gets no result (Litchi never reads the
/// clock), which `refresh` computes.
fn typed_field(values: &MarkValues, kind: &str) -> Result<(String, Option<String>)> {
    let name = values.field_name.as_deref();
    let format = values.format.as_deref();
    let word = field_word(kind).ok_or_else(|| {
        error(Code::InvalidValue, "unknown field type")
            .with_property("fieldType")
            .with_suggestion("give the field's instruction instead")
    })?;
    let named = |quoted: bool| -> Result<String> {
        let name = name.filter(|name| !name.trim().is_empty()).ok_or_else(|| {
            error(Code::MissingProperty, "this field type needs a name").with_property("name")
        })?;
        Ok(if quoted || name.contains(char::is_whitespace) {
            format!("\"{name}\"")
        } else {
            name.to_owned()
        })
    };
    Ok(match word {
        "DATE" | "CREATEDATE" | "SAVEDATE" | "PRINTDATE" | "TIME" => (
            match format.filter(|format| !format.is_empty()) {
                Some(format) => format!(" {word} \\@ \"{format}\" "),
                None => format!(" {word} "),
            },
            None,
        ),
        "PAGE" | "NUMPAGES" | "SECTION" | "SECTIONPAGES" => {
            (format!(" {word} "), Some("1".to_owned()))
        },
        "MERGEFIELD" => {
            let name = named(false)?;
            let shown = format!("\u{ab}{}\u{bb}", name.trim_matches('"'));
            (format!(" {word} {name} "), Some(shown))
        },
        "REF" | "NOTEREF" | "PAGEREF" => {
            let name = named(false)?;
            let shown = if word == "PAGEREF" {
                "1".to_owned()
            } else {
                format!("\u{ab}{}\u{bb}", name.trim_matches('"'))
            };
            let switch = if values.hyperlink == Some(true) {
                " \\h"
            } else {
                ""
            };
            (format!(" {word} {name}{switch} "), Some(shown))
        },
        "IF" => {
            let expression = values
                .expression
                .as_deref()
                .filter(|expression| !expression.trim().is_empty())
                .ok_or_else(|| {
                    error(Code::MissingProperty, "an IF field needs its expression")
                        .with_property("expression")
                })?;
            let quote =
                |text: Option<&str>| format!("\"{}\"", text.unwrap_or("").replace('"', "'"));
            (
                format!(
                    " IF {expression} {} {} ",
                    quote(values.true_text.as_deref()),
                    quote(values.false_text.as_deref())
                ),
                None,
            )
        },
        "SEQ" => (format!(" {word} {} ", named(false)?), None),
        "STYLEREF" | "DOCPROPERTY" => {
            let name = named(true)?;
            let shown = format!("\u{ab}{}\u{bb}", name.trim_matches('"'));
            (format!(" {word} {name} "), Some(shown))
        },
        _ => (format!(" {word} "), None),
    })
}

/// An instruction as OfficeCLI writes a raw one: with a leading and a
/// trailing space.
fn spaced(instruction: &str) -> String {
    let mut out = String::with_capacity(instruction.len() + 2);
    if !instruction.starts_with(' ') {
        out.push(' ');
    }
    out.push_str(instruction);
    if !instruction.ends_with(' ') {
        out.push(' ');
    }
    out
}

/// The levels and switches of a `TOC` instruction.
struct TocSwitches {
    levels: String,
    hyperlinks: bool,
    page_numbers: bool,
}

impl TocSwitches {
    /// OfficeCLI's `TOC \o "levels" [\h] [\n] \u` (with Word's `\n` for
    /// "no page numbers", which the read side reports).
    fn instruction(&self) -> String {
        format!(
            " TOC \\o \"{}\"{}{} \\u ",
            self.levels,
            if self.hyperlinks { " \\h" } else { "" },
            if self.page_numbers { "" } else { " \\n" },
        )
    }
}

/// `instruction` with switch `\name` removed (and its argument, for `o`).
fn without_switch(instruction: &str, name: char) -> String {
    let mut out = String::with_capacity(instruction.len());
    let mut chars = instruction.char_indices().peekable();
    let mut skip_argument = false;
    while let Some((_, character)) = chars.next() {
        if character == '\\'
            && let Some((_, switch)) = chars.peek().copied()
            && switch.eq_ignore_ascii_case(&name)
        {
            chars.next();
            skip_argument = name == 'o';
            continue;
        }
        if skip_argument {
            if character == ' ' && !out.ends_with(' ') {
                out.push(' ');
                continue;
            }
            if character == '"' {
                for (_, inner) in chars.by_ref() {
                    if inner == '"' {
                        break;
                    }
                }
                skip_argument = false;
                continue;
            }
            if character != ' ' {
                skip_argument = false;
            } else {
                continue;
            }
        }
        if character == ' ' && out.ends_with(' ') {
            continue;
        }
        out.push(character);
    }
    out
}

/// A `TOC` instruction with its levels or switches changed, keeping the
/// rest of it.
fn toc_instruction(
    instruction: &str,
    levels: Option<&str>,
    hyperlinks: Option<bool>,
    page_numbers: Option<bool>,
) -> String {
    let mut out = instruction.to_owned();
    if let Some(levels) = levels {
        out = without_switch(&out, 'o');
        out = format!("{} \\o \"{levels}\" ", out.trim_end());
    }
    if let Some(on) = hyperlinks {
        out = without_switch(&out, 'h');
        if on {
            out = format!("{} \\h ", out.trim_end());
        }
    }
    if let Some(on) = page_numbers {
        out = without_switch(&out, 'n');
        if !on {
            out = format!("{} \\n ", out.trim_end());
        }
    }
    spaced(out.trim())
}

/// The position among the transaction index's fields of the document field
/// at `wanted` (the order the read side lists them), matched by
/// instruction in document order.
fn field_index(document: &[litchi_docx::Field], index: &Fields, wanted: usize) -> Option<usize> {
    let fields = index.fields();
    let mut next = 0;
    for (position, field) in document.iter().enumerate() {
        let instruction = field.instruction().trim();
        while fields
            .get(next)
            .is_some_and(|candidate| candidate.instruction().trim() != instruction)
        {
            next += 1;
        }
        fields.get(next)?;
        if position == wanted {
            return Some(next);
        }
        next += 1;
    }
    None
}

/// The position among the transaction index's bookmarks of a read bookmark.
fn bookmark_index(index: &Fields, bookmark: &litchi_docx::bookmark::Bookmark) -> Option<usize> {
    index.bookmarks().iter().position(|candidate| {
        candidate.name() == bookmark.name() && candidate.id() == Some(i64::from(bookmark.id()))
    })
}

fn mark_path(step: &str, position: usize) -> Path<'static> {
    Path::from_steps([
        PathStep::new(step.to_owned()).with(Predicate::Index(Position::new(position)))
    ])
}

fn text_of(value: Option<Scalar<'_>>) -> Option<String> {
    match value {
        Some(Scalar::Text(text) | Scalar::Token(text)) => Some(text.to_owned()),
        Some(other) => Some(other.to_string()),
        None => None,
    }
}

fn bool_of(value: Option<Scalar<'_>>) -> Option<bool> {
    match value {
        Some(Scalar::Bool(value)) => Some(value),
        _ => None,
    }
}

/// The admitted values of one bookmark, field or table-of-contents write.
#[derive(Default)]
struct MarkValues {
    name: Option<String>,
    text: Option<String>,
    field_type: Option<String>,
    instruction: Option<String>,
    field_name: Option<String>,
    format: Option<String>,
    levels: Option<String>,
    hyperlinks: Option<bool>,
    page_numbers: Option<bool>,
    title: Option<String>,
    hyperlink: Option<bool>,
    expression: Option<String>,
    true_text: Option<String>,
    false_text: Option<String>,
}

/// Admit `props` through `table` for `operation`.
fn admit_mark<const N: usize, const A: usize>(
    table: &crate::registry::Table<MarkSetter, N, A>,
    props: &Props<'_>,
    operation: Operation,
) -> Result<MarkValues> {
    let mut values = MarkValues::default();
    for (key, value) in props.iter() {
        let (_, entry) = table.resolve(key.as_str()).map_err(|error| {
            error
                .with_property(key.as_str().to_owned())
                .with_format(FileFormat::Docx)
        })?;
        entry
            .prepare(operation, value, |setter, value| {
                match setter {
                    MarkSetter::BookmarkName => values.name = text_of(value),
                    MarkSetter::BookmarkText | MarkSetter::FieldText => {
                        values.text = text_of(value)
                    },
                    MarkSetter::FieldType => values.field_type = text_of(value),
                    MarkSetter::FieldInstruction => values.instruction = text_of(value),
                    MarkSetter::FieldName => values.field_name = text_of(value),
                    MarkSetter::FieldFormat => values.format = text_of(value),
                    MarkSetter::TocLevels => values.levels = text_of(value),
                    MarkSetter::TocHyperlinks => values.hyperlinks = bool_of(value),
                    MarkSetter::TocPageNumbers => values.page_numbers = bool_of(value),
                    MarkSetter::TocTitle => values.title = text_of(value),
                    MarkSetter::FieldHyperlink => values.hyperlink = bool_of(value),
                    MarkSetter::FieldExpression => values.expression = text_of(value),
                    MarkSetter::FieldTrueText => values.true_text = text_of(value),
                    MarkSetter::FieldFalseText => values.false_text = text_of(value),
                    MarkSetter::TabPosition
                    | MarkSetter::TabKind
                    | MarkSetter::TabLeader
                    | MarkSetter::BreakType => {},
                }
                Ok(())
            })
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
    }
    Ok(values)
}

/// A bookmark name OfficeCLI's Word accepts: at most 40 characters, and not
/// one of Word's hidden `_` names.
fn check_bookmark_name(name: &str) -> Result<()> {
    if name.is_empty() || name.chars().count() > 40 || name.starts_with('_') {
        return Err(error(
            Code::InvalidValue,
            "a bookmark name has 1 to 40 characters and does not start with _",
        )
        .with_property("name"));
    }
    Ok(())
}

impl NewMark {
    /// The new mark element `name` names, if it is one.
    pub(in super::super) fn of(name: &str) -> Option<Self> {
        if name.eq_ignore_ascii_case("tab") {
            Some(Self::Tab)
        } else if ["break", "br", "pagebreak", "columnbreak"]
            .iter()
            .any(|known| name.eq_ignore_ascii_case(known))
        {
            Some(Self::Break)
        } else if name.eq_ignore_ascii_case("bookmark") {
            Some(Self::Bookmark)
        } else if name.eq_ignore_ascii_case("field") {
            Some(Self::Field(None))
        } else if name.eq_ignore_ascii_case("toc") || name.eq_ignore_ascii_case("tableofcontents") {
            Some(Self::Toc)
        } else {
            // `section` adds a section break (OfficeCLI), not a field.
            FIELD_TYPES
                .iter()
                .filter(|(known, _)| *known != "section")
                .find(|(known, _)| name.eq_ignore_ascii_case(known))
                .map(|(known, _)| Self::Field(Some(*known)))
        }
    }
}

/// The paragraph node a mark belongs to.
fn paragraph_node<'f>(
    docx: &'f Docx,
    paragraph: &Value,
    address: &super::super::node::Address,
) -> Node<'f> {
    Node::new(docx, Kind::Paragraph, paragraph.clone(), address.clone())
}

/// Apply `props` to `stop` through the tab registry.
fn apply_tab(stop: &mut TabStop, props: &Props<'_>, operation: Operation) -> Result<()> {
    for (key, value) in props.iter() {
        let (_, entry) = TAB_REGISTRY.resolve(key.as_str()).map_err(|error| {
            error
                .with_property(key.as_str().to_owned())
                .with_format(FileFormat::Docx)
        })?;
        entry
            .prepare(operation, value, |setter, value| {
                match (setter, value) {
                    (MarkSetter::TabPosition, Some(Scalar::Int(position))) => {
                        stop.position = Twips(position)
                    },
                    (MarkSetter::TabKind, Some(Scalar::Token(token) | Scalar::Text(token))) => {
                        stop.kind = tab_kind(token).ok_or_else(|| {
                            error(Code::InvalidValue, "unknown tab stop alignment")
                        })?;
                    },
                    (MarkSetter::TabLeader, Some(Scalar::Token(token) | Scalar::Text(token))) => {
                        stop.leader = Some(
                            tab_leader(token)
                                .ok_or_else(|| error(Code::InvalidValue, "unknown tab leader"))?,
                        );
                    },
                    _ => {},
                }
                Ok(())
            })
            .map_err(|error| {
                error
                    .with_property(key.as_str().to_owned())
                    .with_format(FileFormat::Docx)
            })?;
    }
    Ok(())
}

/// Format the paragraph at `path` with `change`, on either owner.
fn format_paragraph(
    edit: &mut WorkEdit,
    path: &ElementPath,
    change: impl FnOnce(&mut litchi_docx::properties::Paragraph) -> Result<()>,
) -> Result<()> {
    let mut failure = None;
    let mut change = Some(change);
    let mut apply = |paragraph: &mut litchi_docx::properties::Paragraph| {
        if let Some(change) = change.take()
            && let Err(error) = change(paragraph)
        {
            failure = Some(error);
        }
    };
    match edit {
        WorkEdit::Eager(edit) => {
            edit.format_paragraph_at(path, &mut apply)
                .map_err(Error::from)?;
        },
        WorkEdit::Source(edit) => {
            edit.format_paragraph_at(path, &mut apply)
                .map_err(Error::from)?;
        },
    }
    failure.map_or(Ok(()), Err)
}

impl Docx {
    /// `set` on marks: `Err(editor)` hands the editor back when the targets
    /// are not marks.
    pub(in super::super) fn try_set_marks(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
        props: &Props<'_>,
    ) -> Result<std::result::Result<Commit, DocxEditor>> {
        if nodes.is_empty()
            || nodes
                .iter()
                .any(|node| !super::is_mark_kind(node.element_kind()))
        {
            return Ok(Err(editor));
        }
        if nodes.iter().all(|node| {
            matches!(
                node.element_kind(),
                Kind::Bookmark | Kind::Field | Kind::Toc
            )
        }) {
            return self.set_field_marks(editor, nodes, props).map(Ok);
        }
        if nodes.iter().any(|node| node.element_kind() != Kind::Tab) {
            return Err(error(
                Code::UnsupportedCapability,
                "tab stops, bookmarks, fields and tables of contents are set; breaks, watermarks and text boxes are read-only here",
            ));
        }
        // Admit on a scratch stop first, so a refused key stages nothing.
        let mut probe = TabStop::new(TabKind::Left, Twips(0));
        apply_tab(&mut probe, props, Operation::Set)?;
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(editor)?;
        for node in nodes {
            let Value::Mark(mark) = &node.value else {
                return Err(error(Code::InternalError, "a mark node holds no mark"));
            };
            let MarkValue::Tab {
                anchor,
                index,
                stop,
            } = mark.as_ref()
            else {
                return Err(error(Code::InternalError, "a tab node holds no tab stop"));
            };
            let path = element_path(&paragraph_node(self, &anchor.paragraph, &anchor.address))?;
            let mut next = *stop;
            apply_tab(&mut next, props, Operation::Set)?;
            let index = *index;
            format_paragraph(&mut edit, &path, |properties| {
                with_stop(properties, index, Some(next))
            })?;
        }
        Ok(Ok(edit.commit(self)?))
    }

    /// `add` of a mark under a paragraph.
    pub(in super::super) fn add_mark(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        kind: NewMark,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        match kind {
            NewMark::Bookmark => return self.add_bookmark(editor, parent, node),
            NewMark::Field(shortcut) => return self.add_field(editor, parent, shortcut, node),
            NewMark::Toc => return self.add_toc(editor, parent, node),
            NewMark::Tab | NewMark::Break => {},
        }
        if parent.element_kind() != Kind::Paragraph {
            return Err(error(
                Code::UnsupportedCapability,
                "tab stops and breaks are added under a paragraph",
            ));
        }
        let path = element_path(parent)?;
        match kind {
            NewMark::Tab => {
                let mut stop = TabStop::new(TabKind::Left, Twips(0));
                apply_tab(&mut stop, node.props(), Operation::Add)?;
                let count = super::tab_stops(&parent.value)?.len();
                self.owner().check_source()?;
                let mut edit = WorkEdit::from(editor)?;
                format_paragraph(&mut edit, &path, |properties| {
                    properties.tabs.get_or_insert_with(Vec::new).push(stop);
                    Ok(())
                })?;
                let commit = edit.commit(self)?;
                Ok((
                    commit,
                    path_with(parent, Some((Kind::Tab, Position::new(count)))),
                ))
            },
            NewMark::Bookmark | NewMark::Field(_) | NewMark::Toc => Err(error(
                Code::InternalError,
                "a field mark reached the paragraph route",
            )),
            NewMark::Break => {
                let mut kind = RunBreakType::Page;
                for (key, value) in node.props().iter() {
                    let (_, entry) = BREAK_REGISTRY.resolve(key.as_str()).map_err(|error| {
                        error
                            .with_property(key.as_str().to_owned())
                            .with_format(FileFormat::Docx)
                    })?;
                    entry
                        .prepare(Operation::Add, value, |_, value| {
                            kind = match value {
                                Some(Scalar::Token("column")) => RunBreakType::Column,
                                Some(Scalar::Token("textWrapping" | "line")) => {
                                    RunBreakType::TextWrapping
                                },
                                _ => RunBreakType::Page,
                            };
                            Ok(())
                        })
                        .map_err(|error| {
                            error
                                .with_property(key.as_str().to_owned())
                                .with_format(FileFormat::Docx)
                        })?;
                }
                let Value::Paragraph { paragraph, .. } = &parent.value else {
                    return Err(refusal("the break parent holds no paragraph"));
                };
                let locations = run_locations(self, paragraph)?;
                let at = super::super::nested::sibling_position(
                    self,
                    parent,
                    Kind::Run,
                    locations.len(),
                    node.placement(),
                )?;
                let direct = match at {
                    None => None,
                    Some(index) if index >= locations.len() => None,
                    Some(index) => match locations.get(index).map(Vec::as_slice) {
                        Some([ElementStep::Run(run)]) => Some(*run),
                        _ => {
                            return Err(refusal(
                                "a break goes beside a direct run of the paragraph",
                            ));
                        },
                    },
                };
                // The new break's ordinal: the breaks of the runs before it.
                let before = super::breaks(&parent.value)?
                    .iter()
                    .filter(|(_, run)| at.is_none_or(|at| *run < at))
                    .count();
                self.owner().check_source()?;
                let mut edit = WorkEdit::from(editor)?;
                match &mut edit {
                    WorkEdit::Eager(edit) => edit.insert_break_at(&path, direct, kind),
                    WorkEdit::Source(edit) => edit.insert_break_at(&path, direct, kind),
                }
                .map_err(Error::from)?;
                let commit = edit.commit(self)?;
                Ok((
                    commit,
                    path_with(parent, Some((Kind::Break, Position::new(before)))),
                ))
            },
        }
    }

    /// `remove` of marks: tab stops, and breaks whose run holds nothing
    /// else.
    pub(in super::super) fn remove_marks(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
    ) -> Result<Commit> {
        enum Removal {
            Tab(ElementPath, usize),
            Run(ElementPath),
        }
        let mut removals = Vec::with_capacity(nodes.len());
        for node in nodes {
            let Value::Mark(mark) = &node.value else {
                return Err(error(Code::InternalError, "a mark node holds no mark"));
            };
            match mark.as_ref() {
                MarkValue::Tab { anchor, index, .. } => removals.push(Removal::Tab(
                    element_path(&paragraph_node(self, &anchor.paragraph, &anchor.address))?,
                    *index,
                )),
                MarkValue::Break { anchor, run, .. } => {
                    let Value::Paragraph {
                        paragraph: value, ..
                    } = &anchor.paragraph
                    else {
                        return Err(error(
                            Code::InternalError,
                            "a break node holds no paragraph",
                        ));
                    };
                    let runs = value.addressable_runs().map_err(Error::from)?;
                    let holds_only_breaks = runs.get(*run).is_some_and(|candidate| {
                        candidate.contents().is_ok_and(|contents| {
                            contents
                                .iter()
                                .all(|content| matches!(content, litchi_docx::RunContent::Break(_)))
                        })
                    });
                    if !holds_only_breaks {
                        return Err(error(
                            Code::UnsupportedCapability,
                            "a break is removed with its run, which here holds other content",
                        ));
                    }
                    let parent =
                        element_path(&paragraph_node(self, &anchor.paragraph, &anchor.address))?;
                    let locations = run_locations(self, value)?;
                    let location = locations
                        .get(*run)
                        .ok_or_else(|| error(Code::PathNotFound, "the break's run is gone"))?;
                    let mut path = parent;
                    for step in location {
                        path = path.child(*step);
                    }
                    removals.push(Removal::Run(path));
                },
                MarkValue::Bookmark(_) | MarkValue::Toc(_) if nodes.len() == 1 => {
                    return self.remove_field_mark(editor, node);
                },
                _ => {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "tab stops, breaks, bookmarks and tables of contents are removed (one bookmark or table of contents a call); fields, watermarks and text boxes are not",
                    ));
                },
            }
        }
        // Later positions first, so no removal shifts another.
        removals.sort_by(|left, right| {
            let key = |removal: &Removal| match removal {
                Removal::Tab(path, index) => (super::super::nested::order_key(path), *index),
                Removal::Run(path) => (super::super::nested::order_key(path), 0),
            };
            key(right).cmp(&key(left))
        });
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(editor)?;
        for removal in removals {
            match removal {
                Removal::Tab(path, index) => {
                    format_paragraph(&mut edit, &path, |properties| {
                        with_stop(properties, index, None)
                    })?;
                },
                Removal::Run(path) => edit.remove_element_at(&path)?,
            }
        }
        edit.commit(self)
    }
}

impl Docx {
    /// The direct run of `parent` (a paragraph) a new inline mark goes
    /// before, from the add's placement: `None` appends.
    fn mark_run_position(&self, parent: &Node<'_>, node: &NewNode<'_>) -> Result<Option<Position>> {
        Ok(self.mark_run_slot(parent, node)?.0)
    }

    /// [`Self::mark_run_position`], with the addressable-run ordinal the
    /// new mark's first run takes.
    fn mark_run_slot(
        &self,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Option<Position>, usize)> {
        let Value::Paragraph { paragraph, .. } = &parent.value else {
            return Err(refusal("the mark parent holds no paragraph"));
        };
        let locations = run_locations(self, paragraph)?;
        let at = super::super::nested::sibling_position(
            self,
            parent,
            Kind::Run,
            locations.len(),
            node.placement(),
        )?;
        match at {
            None => Ok((None, locations.len())),
            Some(index) if index >= locations.len() => Ok((None, locations.len())),
            Some(index) => match locations.get(index).map(Vec::as_slice) {
                Some([ElementStep::Run(run)]) => Ok((Some(*run), index)),
                _ => Err(refusal(
                    "a bookmark or field goes beside a direct run of the paragraph",
                )),
            },
        }
    }

    /// The paragraph a body-level mark goes into: a new empty paragraph at
    /// the add's placement, inserted into `edit`.
    fn new_mark_paragraph(
        &self,
        edit: &mut WorkEdit,
        node: &NewNode<'_>,
        paragraph: &litchi_docx::document::NewParagraph,
    ) -> Result<ElementPath> {
        let anchor = super::super::structure::body_anchor(self, node.placement(), false)?;
        let position = edit.insert_paragraph(anchor, paragraph)?;
        Ok(ElementPath::new([ElementStep::Paragraph(position)]))
    }

    fn eager_work(&self, editor: DocxEditor) -> Result<WorkEdit> {
        self.owner().check_source()?;
        WorkEdit::from(self.eager_editor(editor)?)
    }

    /// `add` of a bookmark: under a paragraph, around `text` (when given)
    /// before a run; under the body, a new paragraph holding it and `text`.
    fn add_bookmark(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        let values = admit_mark(&BOOKMARK_REGISTRY, node.props(), Operation::Add)?;
        let name = values.name.ok_or_else(|| {
            error(Code::MissingProperty, "a bookmark needs its name").with_property("name")
        })?;
        check_bookmark_name(&name)?;
        if super::bookmarks(self)?
            .iter()
            .any(|mark| mark.name() == name)
        {
            return Err(
                error(Code::DuplicateName, "a bookmark with this name exists")
                    .with_property("name"),
            );
        }
        let text = values.text.filter(|text| !text.is_empty());
        // A bookmark in a header, footer or note paragraph (record 0325).
        if parent.element_kind() == Kind::Paragraph
            && let Some((key, path)) = super::super::element::story_path(parent)?
        {
            let at = self.mark_run_position(parent, node)?;
            let mut edit = self.eager_work(editor)?;
            let WorkEdit::Eager(eager) = &mut edit else {
                return Err(refusal("bookmarks need the owning package editor"));
            };
            eager.edit_story(super::super::story::target(&key), |story| -> Result<i64> {
                Ok(story.insert_bookmark_at(&path, at, &name, text.as_deref())?)
            })?;
            let commit = edit.commit(self)?;
            return Ok((commit, path_with(parent, None)));
        }
        let mut edit;
        match parent.element_kind() {
            Kind::Paragraph => {
                let path = element_path(parent)?;
                let at = self.mark_run_position(parent, node)?;
                edit = self.eager_work(editor)?;
                let WorkEdit::Eager(eager) = &mut edit else {
                    return Err(refusal("bookmarks need the owning package editor"));
                };
                eager
                    .insert_bookmark_at(&path, at, &name, text.as_deref())
                    .map_err(Error::from)?;
            },
            Kind::Body => {
                let Some(text) = text else {
                    return Err(error(
                        Code::UnsupportedCapability,
                        "a bookmark under the body needs text; add an empty one under a paragraph",
                    )
                    .with_property("text"));
                };
                edit = self.eager_work(editor)?;
                let path = self.new_mark_paragraph(
                    &mut edit,
                    node,
                    &litchi_docx::document::NewParagraph::new(),
                )?;
                let WorkEdit::Eager(eager) = &mut edit else {
                    return Err(refusal("bookmarks need the owning package editor"));
                };
                eager
                    .insert_bookmark_at(&path, None, &name, Some(&text))
                    .map_err(Error::from)?;
            },
            _ => {
                return Err(error(
                    Code::UnsupportedCapability,
                    "a bookmark is added under a paragraph or the body",
                ));
            },
        }
        let commit = edit.commit(self)?;
        let position = super::bookmarks(commit.file())?
            .iter()
            .position(|mark| mark.name() == name)
            .ok_or_else(|| refusal("the new bookmark did not read back"))?;
        Ok((commit, mark_path("bookmark", position)))
    }

    /// `add` of a field: under a paragraph (of the body, a header, a footer
    /// or a note) before a run, or under the body in a new paragraph. The
    /// path returned is the field's result run, as OfficeCLI returns it.
    fn add_field(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        shortcut: Option<&'static str>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        let values = admit_mark(&FIELD_REGISTRY, node.props(), Operation::Add)?;
        let kind = values.field_type.as_deref().or(shortcut);
        let (instruction, result) = match (&values.instruction, kind) {
            // A raw instruction overrides the typed one (OfficeCLI).
            (Some(instruction), _) => (spaced(instruction), None),
            (None, Some(kind)) => typed_field(&values, kind)?,
            (None, None) => {
                return Err(error(
                    Code::MissingProperty,
                    "a field needs its fieldType or instruction",
                )
                .with_property("fieldType"));
            },
        };
        let result = values.text.clone().or(result);
        // The result run follows the begin, instruction and separate runs;
        // without a result, the separate run is the last before the end.
        let offset = if result.as_deref().is_some_and(|result| !result.is_empty()) {
            3
        } else {
            2
        };
        match parent.element_kind() {
            Kind::Paragraph => {
                let (at, ordinal) = self.mark_run_slot(parent, node)?;
                let story = super::super::element::story_path(parent)?;
                let path = match &story {
                    Some((_, path)) => path.clone(),
                    None => element_path(parent)?,
                };
                let mut edit = self.eager_work(editor)?;
                let WorkEdit::Eager(eager) = &mut edit else {
                    return Err(refusal("fields need the owning package editor"));
                };
                match &story {
                    // A header, footer or note paragraph (record 0325).
                    Some((key, _)) => {
                        eager.edit_story(
                            super::super::story::target(key),
                            |story| -> Result<Position> {
                                Ok(story.insert_field_at(
                                    &path,
                                    at,
                                    &instruction,
                                    result.as_deref(),
                                )?)
                            },
                        )?;
                    },
                    None => {
                        eager
                            .insert_field_at(&path, at, &instruction, result.as_deref())
                            .map_err(Error::from)?;
                    },
                }
                let commit = edit.commit(self)?;
                Ok((
                    commit,
                    path_with(parent, Some((Kind::Run, Position::new(ordinal + offset)))),
                ))
            },
            Kind::Body => {
                let mut edit = self.eager_work(editor)?;
                let anchor = super::super::structure::body_anchor(self, node.placement(), false)?;
                let position =
                    edit.insert_paragraph(anchor, &litchi_docx::document::NewParagraph::new())?;
                let WorkEdit::Eager(eager) = &mut edit else {
                    return Err(refusal("fields need the owning package editor"));
                };
                eager
                    .insert_field_at(
                        &ElementPath::new([ElementStep::Paragraph(position)]),
                        None,
                        &instruction,
                        result.as_deref(),
                    )
                    .map_err(Error::from)?;
                let commit = edit.commit(self)?;
                let paragraph = path_with(parent, Some((Kind::Paragraph, position)));
                let mut steps = paragraph.steps().to_vec();
                steps.push(
                    PathStep::new(Kind::Run.step_name())
                        .with(Predicate::Index(Position::new(offset))),
                );
                Ok((commit, Path::from_steps(steps)))
            },
            _ => Err(error(
                Code::UnsupportedCapability,
                "a field is added under a paragraph or the body",
            )),
        }
    }

    /// `add` of a table of contents under the body: one paragraph holding a
    /// `TOC` field with the "Update field" placeholder, after a `TOCHeading`
    /// title paragraph when `title` is given.
    fn add_toc(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        if parent.element_kind() != Kind::Body {
            return Err(error(
                Code::UnsupportedCapability,
                "a table of contents is added under the body",
            ));
        }
        let values = admit_mark(&TOC_REGISTRY, node.props(), Operation::Add)?;
        let switches = TocSwitches {
            levels: values.levels.unwrap_or_else(|| "1-3".to_owned()),
            hyperlinks: values.hyperlinks.unwrap_or(true),
            page_numbers: values.page_numbers.unwrap_or(true),
        };
        if !switches
            .levels
            .chars()
            .all(|character| character.is_ascii_digit() || character == '-')
        {
            return Err(
                error(Code::InvalidValue, "levels are a range such as 1-3").with_property("levels")
            );
        }
        let before = super::fields(self)?
            .iter()
            .filter(|field| super::is_toc(field))
            .count();
        let anchor = super::super::structure::body_anchor(self, node.placement(), false)?;
        let mut edit = self.eager_work(editor)?;
        let position =
            edit.insert_paragraph(anchor, &litchi_docx::document::NewParagraph::new())?;
        let path = ElementPath::new([ElementStep::Paragraph(position)]);
        {
            let WorkEdit::Eager(eager) = &mut edit else {
                return Err(refusal(
                    "a table of contents needs the owning package editor",
                ));
            };
            eager
                .insert_field_at(
                    &path,
                    None,
                    &switches.instruction(),
                    Some("Update field to see table of contents"),
                )
                .map_err(Error::from)?;
        }
        if let Some(title) = values.title.filter(|title| !title.is_empty()) {
            let mut heading = litchi_docx::document::NewParagraph::new()
                .with_run(litchi_docx::document::NewRun::new(title));
            heading.properties_mut().style = Some("TOCHeading".to_owned());
            edit.insert_paragraph(
                litchi_docx::document::BodyAnchor::Before(
                    litchi_docx::document::BodyBlock::Paragraph(position),
                ),
                &heading,
            )?;
        }
        let commit = edit.commit(self)?;
        let count = super::fields(commit.file())?
            .iter()
            .filter(|field| super::is_toc(field))
            .count();
        if count != before + 1 {
            return Err(refusal("the new table of contents did not read back"));
        }
        // Tables of contents before the new one keep their positions.
        let position = super::fields(commit.file())?
            .iter()
            .filter(|field| super::is_toc(field))
            .zip(
                super::fields(self)?
                    .iter()
                    .filter(|field| super::is_toc(field)),
            )
            .position(|(new, old)| new.instruction() != old.instruction())
            .unwrap_or(before);
        Ok((commit, mark_path("toc", position)))
    }
}

impl Docx {
    /// The position of a read field (or table of contents) node among the
    /// read side's fields.
    fn field_position(node: &Node<'_>, document: &[litchi_docx::Field]) -> Result<usize> {
        let ordinal = node
            .address
            .steps()
            .last()
            .map(|step| step.ordinal as usize)
            .ok_or_else(|| error(Code::InternalError, "a field node has no address"))?;
        if node.element_kind() == Kind::Toc {
            return document
                .iter()
                .enumerate()
                .filter(|(_, field)| super::is_toc(field))
                .nth(ordinal)
                .map(|(position, _)| position)
                .ok_or_else(|| error(Code::PathNotFound, "the table of contents is gone"));
        }
        Ok(ordinal)
    }

    /// The index position of a field node when the index lists the fields
    /// as the read side does: the node's ordinal (among tables of contents,
    /// for a TOC) names a field of the same instruction.
    fn direct_field_index(
        node: &Node<'_>,
        field: &litchi_docx::Field,
        index: &Fields,
    ) -> Option<usize> {
        let ordinal = node.address.steps().last()?.ordinal as usize;
        let position = if node.element_kind() == Kind::Toc {
            index
                .fields()
                .iter()
                .enumerate()
                .filter(|(_, candidate)| {
                    candidate
                        .instruction()
                        .split_whitespace()
                        .next()
                        .is_some_and(|word| word.eq_ignore_ascii_case("TOC"))
                })
                .nth(ordinal)?
                .0
        } else {
            ordinal
        };
        (index.fields().get(position)?.instruction().trim() == field.instruction().trim())
            .then_some(position)
    }

    /// `set` of bookmarks (`name`), fields (`instruction`, `text`) and
    /// tables of contents (`levels`, `hyperlinks`, `pageNumbers`), planned
    /// over one field index and spliced in one leaf call.
    fn set_field_marks(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
        props: &Props<'_>,
    ) -> Result<Commit> {
        let kind = nodes[0].element_kind();
        if nodes.iter().any(|node| node.element_kind() != kind) {
            return Err(error(
                Code::UnsupportedCapability,
                "one set writes bookmarks, fields or tables of contents, not a mix",
            ));
        }
        let values = match kind {
            Kind::Bookmark => admit_mark(&BOOKMARK_REGISTRY, props, Operation::Set)?,
            Kind::Field => admit_mark(&FIELD_REGISTRY, props, Operation::Set)?,
            _ => admit_mark(&TOC_REGISTRY, props, Operation::Set)?,
        };
        if let Some(name) = &values.name {
            check_bookmark_name(name)?;
            if nodes.len() > 1 {
                return Err(error(
                    Code::InvalidArgument,
                    "bookmark names are unique; rename one bookmark at a time",
                )
                .with_property("name"));
            }
        }
        let mut edit = self.eager_work(editor)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal(
                "bookmark and field writes need the owning package editor",
            ));
        };
        let index = eager.fields().map_err(Error::from)?;
        // The read side's field list, only when a field's position in the
        // index is not proven from the node alone.
        let mut document: Option<&[litchi_docx::Field]> = None;
        let mut updates = Vec::with_capacity(nodes.len() * 2);
        for node in nodes {
            let Value::Mark(mark) = &node.value else {
                return Err(error(Code::InternalError, "a mark node holds no mark"));
            };
            match mark.as_ref() {
                MarkValue::Bookmark(bookmark) => {
                    let Some(name) = &values.name else {
                        continue;
                    };
                    if name != bookmark.name()
                        && index.bookmarks().iter().any(|other| other.name() == name)
                    {
                        return Err(
                            error(Code::DuplicateName, "a bookmark with this name exists")
                                .with_property("name"),
                        );
                    }
                    let position = bookmark_index(&index, bookmark)
                        .ok_or_else(|| error(Code::PathNotFound, "the bookmark is gone"))?;
                    updates.push(Update::RenameBookmark {
                        bookmark: position,
                        name: name.clone(),
                    });
                },
                MarkValue::Field(field) | MarkValue::Toc(field) => {
                    let position = match Self::direct_field_index(node, field, &index) {
                        Some(position) => position,
                        None => {
                            let document = match &mut document {
                                Some(document) => document,
                                slot => slot.insert(super::fields(self)?),
                            };
                            let wanted = Self::field_position(node, document)?;
                            field_index(document, &index, wanted).ok_or_else(|| {
                                error(
                                    Code::UnsupportedCapability,
                                    "this field is not one the field index can rewrite",
                                )
                            })?
                        },
                    };
                    let instruction = match kind {
                        Kind::Toc => (values.levels.is_some()
                            || values.hyperlinks.is_some()
                            || values.page_numbers.is_some())
                        .then(|| {
                            toc_instruction(
                                field.instruction(),
                                values.levels.as_deref(),
                                values.hyperlinks,
                                values.page_numbers,
                            )
                        }),
                        _ => values.instruction.as_deref().map(spaced),
                    };
                    if let Some(levels) = &values.levels
                        && !levels
                            .chars()
                            .all(|character| character.is_ascii_digit() || character == '-')
                    {
                        return Err(error(Code::InvalidValue, "levels are a range such as 1-3")
                            .with_property("levels"));
                    }
                    if let Some(instruction) = instruction {
                        updates.push(Update::Instruction {
                            field: position,
                            instruction,
                        });
                    }
                    if let Some(text) = &values.text {
                        if !index.accepts_text(position) {
                            return Err(error(
                                Code::UnsupportedCapability,
                                "this field's result spans paragraphs or holds more than plain runs",
                            )
                            .with_property("text"));
                        }
                        updates.push(Update::Text {
                            field: position,
                            text: text.clone(),
                        });
                    }
                },
                _ => return Err(error(Code::InternalError, "a field mark is expected")),
            }
        }
        if !updates.is_empty() {
            eager
                .refresh_fields(&index, &updates)
                .map_err(Error::from)?;
        }
        edit.commit(self)
    }

    /// `remove` of a bookmark (its markers; what lies between stays) or a
    /// table of contents (its paragraphs, and a `TOCHeading` paragraph just
    /// before them).
    fn remove_field_mark(&self, editor: DocxEditor, node: &Node<'_>) -> Result<Commit> {
        let Value::Mark(mark) = &node.value else {
            return Err(error(Code::InternalError, "a mark node holds no mark"));
        };
        let document = super::fields(self)?;
        let mut edit = self.eager_work(editor)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal(
                "bookmark and field writes need the owning package editor",
            ));
        };
        let index = eager.fields().map_err(Error::from)?;
        match mark.as_ref() {
            MarkValue::Bookmark(bookmark) => {
                let position = bookmark_index(&index, bookmark)
                    .ok_or_else(|| error(Code::PathNotFound, "the bookmark is gone"))?;
                eager
                    .refresh_fields(&index, &[Update::RemoveBookmark { bookmark: position }])
                    .map_err(Error::from)?;
            },
            MarkValue::Toc(_) => {
                let wanted = Self::field_position(node, document)?;
                let position = field_index(document, &index, wanted)
                    .ok_or_else(|| error(Code::PathNotFound, "the table of contents is gone"))?;
                let (first, last) = index.fields()[position].paragraphs();
                let paragraphs = index.paragraphs();
                // Word's own TOC lives in one block content control
                // (`docPartGallery` "Table of Contents"): it goes whole.
                let control = index.body_block(first);
                if let Some(step @ ElementStep::ContentControl(_)) = control
                    && (first..=last).all(|paragraph| index.body_block(paragraph) == control)
                {
                    eager
                        .remove_element_at(&ElementPath::new([step]))
                        .map_err(Error::from)?;
                    return edit.commit(self);
                }
                let mut body: Vec<Position> = Vec::new();
                for paragraph in &paragraphs[first..=last] {
                    let Some(at) = paragraph.body_position() else {
                        return Err(error(
                            Code::UnsupportedCapability,
                            "a table of contents outside the body's own paragraphs and content controls is not removed",
                        ));
                    };
                    body.push(Position::new(at));
                }
                if let Some(heading) = first.checked_sub(1).and_then(|at| paragraphs.get(at))
                    && heading.style() == Some("TOCHeading")
                    && let Some(at) = heading.body_position()
                    && body.first().is_some_and(|first| first.get() == at + 1)
                {
                    body.insert(0, Position::new(at));
                }
                for at in body.into_iter().rev() {
                    eager
                        .remove_element_at(&ElementPath::new([ElementStep::Paragraph(at)]))
                        .map_err(Error::from)?;
                }
            },
            _ => {
                return Err(error(
                    Code::InternalError,
                    "a bookmark or table of contents is expected",
                ));
            },
        }
        edit.commit(self)
    }
}
