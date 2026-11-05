//! The DOCX marks of record 0319: paragraph tab stops and breaks,
//! bookmarks, fields, tables of contents, the watermark and text boxes.
//!
//! | Path | Node | Leaf read |
//! |---|---|---|
//! | `/body/p[N]/tab[K]` | the K-th tab stop of the paragraph's own `w:tabs` (OfficeCLI's `tab`) | `Paragraph::direct_properties` |
//! | `/body/p[N]/break[M]` | the M-th `w:br` of the paragraph's addressable runs | `Run::contents` |
//! | `/bookmark[N]`, `/bookmark[@name=X]` | a bookmark of the main story (Word's hidden `_` bookmarks excluded) | `Document::bookmarks` |
//! | `/field[N]` | a field of the main story | `Document::fields` |
//! | `/toc[N]` | a `TOC` field | `Document::fields` |
//! | `/watermark` | the first text watermark of the headers | `Story::watermarks` |
//! | `/body/textbox[N]` | a text box of the main story | `Document::text_boxes` |
//!
//! Marks are reached by path and query; the node view of the root, the body
//! and a paragraph does not list them, as OfficeCLI's does not.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_docx::bookmark::Bookmark;
use litchi_docx::properties::{self as direct, TabKind, TabLeader, TabStop};
use litchi_docx::textbox::TextBox;
use litchi_docx::{Field, RunBreakType, RunContent, Watermark, WatermarkLayout};

use super::node::{Address, Node, Step, Value};
use super::spec::Kind;
use super::{Docx, DocxOwner};
use crate::error::{Code, Error, Result};
use crate::node::Value as NodeValue;
use crate::node::spec::{GapSpec, PropSpec};
use crate::path::navigate::{KindSet, Lookup, Resolved, Select};
use crate::registry::{Access, Constraint, Entry};
use crate::value::{Tokens, ValueType};

pub(super) mod write;

/// The paragraph a tab stop or break belongs to: its value and address.
pub(crate) struct Anchor {
    pub(super) paragraph: Value,
    pub(super) address: Address,
}

/// What a mark node holds.
pub(crate) enum MarkValue {
    /// A paragraph tab stop: its paragraph, the paragraph's address, its
    /// index among the paragraph's own tab stops, and its value.
    Tab {
        anchor: Anchor,
        index: usize,
        stop: TabStop,
    },
    /// A run break: its type, its paragraph, the paragraph's address and
    /// the index of its run among the paragraph's addressable runs.
    Break {
        kind: RunBreakType,
        anchor: Anchor,
        run: usize,
    },
    Bookmark(Bookmark),
    Field(Field),
    Toc(Field),
    Watermark(Watermark),
    TextBox(TextBox),
}

/// Whether `kind` is one of this module's elements.
pub(super) const fn is_mark_kind(kind: Kind) -> bool {
    matches!(
        kind,
        Kind::Tab
            | Kind::Break
            | Kind::Bookmark
            | Kind::Field
            | Kind::Toc
            | Kind::Watermark
            | Kind::TextBox
    )
}

fn ordinal(index: usize) -> u32 {
    u32::try_from(index).unwrap_or(u32::MAX)
}

fn node<'f>(docx: &'f Docx, kind: Kind, value: MarkValue, address: Address) -> Node<'f> {
    Node::new(docx, kind, Value::Mark(Arc::new(value)), address)
}

fn source_or<T>(
    docx: &Docx,
    eager: impl FnOnce(&litchi_docx::Package) -> litchi_docx::Result<T>,
    source: impl FnOnce(&litchi_docx::source_backed::Package) -> litchi_docx::Result<T>,
) -> Result<T> {
    let result = match docx.owner() {
        DocxOwner::Eager(package) => eager(package).map_err(Error::from),
        DocxOwner::Source(package) => source(package).map_err(Error::from),
    };
    docx.owner().finish(result)
}

/// The main story's bookmarks, read once per handle.
pub(super) fn bookmarks(docx: &Docx) -> Result<&[Bookmark]> {
    let cache = &docx.parts().bookmarks;
    if let Some(bookmarks) = cache.get() {
        docx.owner().check_source()?;
        return Ok(bookmarks);
    }
    let loaded = source_or(
        docx,
        |package| package.document()?.bookmarks(),
        |package| package.document()?.bookmarks(),
    )?;
    Ok(cache.get_or_init(|| loaded))
}

/// The main story's fields, read once per handle.
pub(super) fn fields(docx: &Docx) -> Result<&[Field]> {
    let cache = &docx.parts().fields;
    if let Some(fields) = cache.get() {
        docx.owner().check_source()?;
        return Ok(fields);
    }
    let loaded = source_or(
        docx,
        |package| package.document()?.fields(),
        |package| package.document()?.fields(),
    )?;
    Ok(cache.get_or_init(|| loaded))
}

pub(super) fn is_toc(field: &Field) -> bool {
    field.field_type().eq_ignore_ascii_case("TOC")
}

fn text_boxes(docx: &Docx) -> Result<Vec<TextBox>> {
    source_or(
        docx,
        |package| package.document()?.text_boxes(),
        |package| package.document()?.text_boxes(),
    )
}

/// The first text watermark of the headers, in header order.
fn watermark(docx: &Docx) -> Result<Option<Watermark>> {
    for index in 0..super::parts::stories(docx)?.len() {
        let xml = super::parts::story_xml(docx, index)?;
        let story = litchi_docx::header_footer::Story::from_shared_bytes(
            xml,
            litchi_docx::header_footer::Kind::Primary,
        )
        .map_err(Error::from)?;
        if let Some(found) = story.watermarks().map_err(Error::from)?.into_iter().next() {
            return Ok(Some(found));
        }
    }
    Ok(None)
}

/// The tab stops of a paragraph value, in its `w:tabs` order.
fn tab_stops(paragraph: &Value) -> Result<Vec<TabStop>> {
    let Value::Paragraph { paragraph, .. } = paragraph else {
        return Ok(Vec::new());
    };
    Ok(paragraph.direct_properties()?.tabs.unwrap_or_default())
}

/// The breaks of a paragraph value's addressable runs, in order: each
/// break's type and its run's index.
fn breaks(paragraph: &Value) -> Result<Vec<(RunBreakType, usize)>> {
    let Value::Paragraph { paragraph, .. } = paragraph else {
        return Ok(Vec::new());
    };
    let mut found = Vec::new();
    for (index, run) in paragraph.addressable_runs()?.iter().enumerate() {
        for content in run.contents()? {
            if let RunContent::Break(run_break) = content {
                found.push((run_break.break_type, index));
            }
        }
    }
    Ok(found)
}

fn pick(len: usize, select: &Select<'_>) -> Option<Option<usize>> {
    Some(match *select {
        Select::None => (len > 0).then_some(0),
        Select::Ordinal(position) => (position.get() < len).then(|| position.get()),
        Select::Last => len.checked_sub(1),
        _ => return None,
    })
}

fn found<'f>(picked: Option<Option<Node<'f>>>) -> Lookup<Node<'f>, GapSpec> {
    match picked {
        Some(Some(node)) => Lookup::Found(node),
        Some(None) => Lookup::Missing,
        None => Lookup::NotApplicable,
    }
}

/// The mark child of `parent` that `step` selects.
pub(super) fn child<'f>(
    docx: &'f Docx,
    parent: &Node<'f>,
    step: &Resolved<'_, Kind>,
) -> Result<Lookup<Node<'f>, GapSpec>> {
    let kind = step.kind;
    let select = &step.select;
    let address = parent.child_address();
    Ok(match (&parent.value, kind) {
        (Value::Paragraph { .. }, Kind::Tab) => {
            let stops = tab_stops(&parent.value)?;
            found(pick(stops.len(), select).map(|index| {
                index.map(|index| {
                    node(
                        docx,
                        kind,
                        MarkValue::Tab {
                            anchor: Anchor {
                                paragraph: parent.value.clone(),
                                address: address.clone(),
                            },
                            index,
                            stop: stops[index],
                        },
                        address.with(Step::new(kind, ordinal(index))),
                    )
                })
            }))
        },
        (Value::Paragraph { .. }, Kind::Break) => {
            let types = breaks(&parent.value)?;
            found(pick(types.len(), select).map(|index| {
                index.map(|index| {
                    node(
                        docx,
                        kind,
                        MarkValue::Break {
                            kind: types[index].0,
                            anchor: Anchor {
                                paragraph: parent.value.clone(),
                                address: address.clone(),
                            },
                            run: types[index].1,
                        },
                        address.with(Step::new(kind, ordinal(index))),
                    )
                })
            }))
        },
        (Value::Document, Kind::Bookmark) => {
            let all = bookmarks(docx)?;
            let picked = match select {
                Select::Identity(identities) => {
                    let mut wanted = None;
                    for (_, value) in identities.iter() {
                        wanted = Some(value);
                    }
                    Some(wanted.and_then(|name| all.iter().position(|mark| mark.name() == name)))
                },
                _ => pick(all.len(), select),
            };
            found(picked.map(|index| {
                index.map(|index| {
                    node(
                        docx,
                        kind,
                        MarkValue::Bookmark(all[index].clone()),
                        address.with(Step::new(kind, ordinal(index))),
                    )
                })
            }))
        },
        (Value::Document, Kind::Field | Kind::Toc) => {
            let toc = kind == Kind::Toc;
            let all: Vec<&Field> = fields(docx)?
                .iter()
                .filter(|field| !toc || is_toc(field))
                .collect();
            found(pick(all.len(), select).map(|index| {
                index.map(|index| {
                    let field = all[index].clone();
                    let value = if toc {
                        MarkValue::Toc(field)
                    } else {
                        MarkValue::Field(field)
                    };
                    node(
                        docx,
                        kind,
                        value,
                        address.with(Step::new(kind, ordinal(index))),
                    )
                })
            }))
        },
        (Value::Document, Kind::Watermark) => match select {
            Select::None | Select::Last | Select::Ordinal(_) => {
                let first = !matches!(select, Select::Ordinal(position) if position.get() != 0);
                found(Some(match watermark(docx)? {
                    Some(mark) if first => Some(node(
                        docx,
                        kind,
                        MarkValue::Watermark(mark),
                        address.with(Step::new(kind, 0)),
                    )),
                    _ => None,
                }))
            },
            _ => Lookup::NotApplicable,
        },
        (Value::Body(_), Kind::TextBox) => {
            let all = text_boxes(docx)?;
            found(pick(all.len(), select).map(|index| {
                index.map(|index| {
                    node(
                        docx,
                        kind,
                        MarkValue::TextBox(all[index].clone()),
                        address.with(Step::new(kind, ordinal(index))),
                    )
                })
            }))
        },
        _ => Lookup::NotApplicable,
    })
}

/// The mark children of a node whose kinds are in a set, read on first
/// use.
pub(crate) struct Children {
    nodes: std::vec::IntoIter<(Kind, MarkValue, Address)>,
    pending: Option<Pending>,
}

/// What a listing still has to read.
enum Pending {
    Paragraph {
        paragraph: Value,
        address: Address,
        tabs: bool,
        breaks: bool,
    },
    Document {
        address: Address,
        kinds: KindSet<Kind>,
    },
    Body {
        address: Address,
    },
}

impl Children {
    /// The mark children of `parent` among `kinds`, or `None` when it has
    /// none to list.
    pub(super) fn of(parent: &Node<'_>, kinds: KindSet<Kind>) -> Option<Self> {
        let address = parent.child_address();
        let pending = match &parent.value {
            Value::Paragraph { .. } if kinds.contains(Kind::Tab) || kinds.contains(Kind::Break) => {
                Pending::Paragraph {
                    paragraph: parent.value.clone(),
                    address,
                    tabs: kinds.contains(Kind::Tab),
                    breaks: kinds.contains(Kind::Break),
                }
            },
            Value::Document
                if [Kind::Bookmark, Kind::Field, Kind::Toc, Kind::Watermark]
                    .into_iter()
                    .any(|kind| kinds.contains(kind)) =>
            {
                Pending::Document { address, kinds }
            },
            Value::Body(_) if kinds.contains(Kind::TextBox) => Pending::Body { address },
            _ => return None,
        };
        Some(Self {
            nodes: Vec::new().into_iter(),
            pending: Some(pending),
        })
    }

    fn read(docx: &Docx, pending: Pending) -> Result<Vec<(Kind, MarkValue, Address)>> {
        let mut nodes = Vec::new();
        match pending {
            Pending::Paragraph {
                paragraph,
                address,
                tabs,
                breaks: wants_breaks,
            } => {
                if tabs {
                    for (index, stop) in tab_stops(&paragraph)?.into_iter().enumerate() {
                        nodes.push((
                            Kind::Tab,
                            MarkValue::Tab {
                                anchor: Anchor {
                                    paragraph: paragraph.clone(),
                                    address: address.clone(),
                                },
                                index,
                                stop,
                            },
                            address.with(Step::new(Kind::Tab, ordinal(index))),
                        ));
                    }
                }
                if wants_breaks {
                    for (index, (kind, run)) in breaks(&paragraph)?.into_iter().enumerate() {
                        nodes.push((
                            Kind::Break,
                            MarkValue::Break {
                                kind,
                                anchor: Anchor {
                                    paragraph: paragraph.clone(),
                                    address: address.clone(),
                                },
                                run,
                            },
                            address.with(Step::new(Kind::Break, ordinal(index))),
                        ));
                    }
                }
            },
            Pending::Document { address, kinds } => {
                if kinds.contains(Kind::Bookmark) {
                    for (index, mark) in bookmarks(docx)?.iter().enumerate() {
                        nodes.push((
                            Kind::Bookmark,
                            MarkValue::Bookmark(mark.clone()),
                            address.with(Step::new(Kind::Bookmark, ordinal(index))),
                        ));
                    }
                }
                if kinds.contains(Kind::Field) || kinds.contains(Kind::Toc) {
                    let all = fields(docx)?;
                    if kinds.contains(Kind::Field) {
                        for (index, field) in all.iter().enumerate() {
                            nodes.push((
                                Kind::Field,
                                MarkValue::Field(field.clone()),
                                address.with(Step::new(Kind::Field, ordinal(index))),
                            ));
                        }
                    }
                    if kinds.contains(Kind::Toc) {
                        for (index, field) in all.iter().filter(|field| is_toc(field)).enumerate() {
                            nodes.push((
                                Kind::Toc,
                                MarkValue::Toc(field.clone()),
                                address.with(Step::new(Kind::Toc, ordinal(index))),
                            ));
                        }
                    }
                }
                if kinds.contains(Kind::Watermark)
                    && let Some(mark) = watermark(docx)?
                {
                    nodes.push((
                        Kind::Watermark,
                        MarkValue::Watermark(mark),
                        address.with(Step::new(Kind::Watermark, 0)),
                    ));
                }
            },
            Pending::Body { address } => {
                for (index, text_box) in text_boxes(docx)?.into_iter().enumerate() {
                    nodes.push((
                        Kind::TextBox,
                        MarkValue::TextBox(text_box),
                        address.with(Step::new(Kind::TextBox, ordinal(index))),
                    ));
                }
            },
        }
        Ok(nodes)
    }

    /// The next child, read through `docx`.
    pub(super) fn next_node<'f>(&mut self, docx: &'f Docx) -> Option<Result<Node<'f>>> {
        if let Some(pending) = self.pending.take() {
            match Self::read(docx, pending) {
                Ok(nodes) => self.nodes = nodes.into_iter(),
                Err(error) => return Some(Err(error)),
            }
        }
        let (kind, value, address) = self.nodes.next()?;
        Some(Ok(node(docx, kind, value, address)))
    }
}

// ---------------------------------------------------------------------------
// Properties
// ---------------------------------------------------------------------------

/// A mark property's write declaration.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) enum MarkSetter {
    TabPosition,
    TabKind,
    TabLeader,
    BreakType,
    /// A bookmark's name, and its text on add (record 0325).
    BookmarkName,
    BookmarkText,
    /// A field's type on add, its instruction, its cached result text, and
    /// the name and format a typed field's instruction takes (record 0325).
    FieldType,
    FieldInstruction,
    FieldText,
    FieldName,
    FieldFormat,
    /// A `REF`-family field's `\\h` switch and an `IF` field's expression
    /// and texts (record 0325).
    FieldHyperlink,
    FieldExpression,
    FieldTrueText,
    FieldFalseText,
    /// A table of contents' levels and switches, and its title on add
    /// (record 0325).
    TocLevels,
    TocHyperlinks,
    TocPageNumbers,
    TocTitle,
}

/// A tab stop's `val` (`ST_TabJc`).
pub(crate) static TAB_VAL: Tokens = Tokens::new(
    "tab stop alignment",
    &[
        "left", "center", "right", "decimal", "bar", "clear", "num", "start", "end",
    ],
    &[],
);

/// A tab stop's `leader` (`ST_TabTlc`).
pub(crate) static TAB_LEADER: Tokens = Tokens::new(
    "tab leader",
    &["none", "dot", "heavy", "hyphen", "middleDot", "underscore"],
    &[],
);

/// A break's type (`ST_BrType`, with OfficeCLI's `line` for a text-wrapping
/// break).
pub(crate) static BREAK_TYPE: Tokens = Tokens::new(
    "break type",
    &["page", "column", "textWrapping", "line"],
    &[],
);

const fn rsa(
    name: &'static str,
    ty: ValueType,
    setter: MarkSetter,
    constraint: Constraint,
) -> Entry<MarkSetter> {
    Entry {
        spec: PropSpec::new(name, ty),
        access: Access::SetAndAdd(setter),
        constraint,
        clear: false,
    }
}

const fn read(name: &'static str, ty: ValueType) -> Entry<MarkSetter> {
    Entry::read(PropSpec::new(name, ty))
}

/// A key `get` reads and only `add` writes.
const fn read_add(name: &'static str, ty: ValueType, setter: MarkSetter) -> Entry<MarkSetter> {
    Entry {
        spec: PropSpec::new(name, ty),
        access: Access::Add(setter),
        constraint: Constraint::None,
        clear: false,
    }
}

/// A key `set` and `add` write and `get` does not read.
const fn write_only(name: &'static str, ty: ValueType, setter: MarkSetter) -> Entry<MarkSetter> {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    Entry {
        spec,
        access: Access::SetAndAdd(setter),
        constraint: Constraint::None,
        clear: false,
    }
}

const fn add_only(name: &'static str, ty: ValueType, setter: MarkSetter) -> Entry<MarkSetter> {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    Entry {
        spec,
        access: Access::Add(setter),
        constraint: Constraint::None,
        clear: false,
    }
}

// A paragraph tab stop (OfficeCLI's `tab`): position in twips, alignment
// and leader.
crate::registry::registry! {
    pub(crate) TAB_PROPS, TAB_REGISTRY: MarkSetter {
    rsa("pos", ValueType::Int, MarkSetter::TabPosition, Constraint::Int { min: -31_680, max: 31_680 }),
    rsa("val", ValueType::Token(&TAB_VAL), MarkSetter::TabKind, Constraint::None),
    rsa("leader", ValueType::Token(&TAB_LEADER), MarkSetter::TabLeader, Constraint::None),
    }
}

// A run break: `breakType` on get, `type` seeds a new one.
crate::registry::registry! {
    pub(crate) BREAK_PROPS, BREAK_REGISTRY: MarkSetter {
    read("breakType", ValueType::Token(&BREAK_TYPE)),
    add_only("type", ValueType::Token(&BREAK_TYPE), MarkSetter::BreakType),
    }
}

// A bookmark: its name and native identifier; `name` is renamed on set
// and, with `text`, seeds a new one (record 0325).
crate::registry::registry! {
    pub(crate) BOOKMARK_PROPS, BOOKMARK_REGISTRY: MarkSetter {
    rsa("name", ValueType::Text, MarkSetter::BookmarkName, Constraint::NotEmpty),
    read("id", ValueType::Text),
    add_only("text", ValueType::Text, MarkSetter::BookmarkText),
    }
}

// A field: its type, instruction, lock and whether it has a result. A new
// field takes OfficeCLI's `fieldType` (with `name` and `format`) or a raw
// `instruction`; `set` replaces the instruction or the result `text`
// (record 0325).
crate::registry::registry! {
    pub(crate) FIELD_PROPS, FIELD_REGISTRY: MarkSetter {
    read_add("fieldType", ValueType::Text, MarkSetter::FieldType),
    rsa("instruction", ValueType::Text, MarkSetter::FieldInstruction, Constraint::NotEmpty),
    read("fldLock", ValueType::Bool),
    read("evaluated", ValueType::Bool),
    write_only("text", ValueType::Text, MarkSetter::FieldText),
    add_only("name", ValueType::Text, MarkSetter::FieldName),
    add_only("bookmarkName", ValueType::Text, MarkSetter::FieldName),
    add_only("format", ValueType::Text, MarkSetter::FieldFormat),
    add_only("hyperlink", ValueType::Bool, MarkSetter::FieldHyperlink),
    add_only("expression", ValueType::Text, MarkSetter::FieldExpression),
    add_only("trueText", ValueType::Text, MarkSetter::FieldTrueText),
    add_only("falseText", ValueType::Text, MarkSetter::FieldFalseText),
    }
}

// A table of contents, from its `TOC` field's switches; written on add and
// set, with a `title` heading on add (record 0325).
crate::registry::registry! {
    pub(crate) TOC_PROPS, TOC_REGISTRY: MarkSetter {
    rsa("levels", ValueType::Text, MarkSetter::TocLevels, Constraint::NotEmpty),
    rsa("hyperlinks", ValueType::Bool, MarkSetter::TocHyperlinks, Constraint::None),
    rsa("pageNumbers", ValueType::Bool, MarkSetter::TocPageNumbers, Constraint::None),
    add_only("title", ValueType::Text, MarkSetter::TocTitle),
    }
}

// The watermark, from its VML text shape.
crate::registry::registry! {
    pub(crate) WATERMARK_PROPS, WATERMARK_REGISTRY: MarkSetter {
    read("text", ValueType::Text),
    read("color", ValueType::Text),
    read("font", ValueType::Text),
    read("rotation", ValueType::Int),
    read("opacity", ValueType::Text),
    read("size", ValueType::Text),
    }
}

// A text box: its shape name and preset geometry.
crate::registry::registry! {
    pub(crate) TEXT_BOX_PROPS, TEXT_BOX_REGISTRY: MarkSetter {
    read("name", ValueType::Text),
    read("geometry", ValueType::Text),
    }
}

/// The kind's property table.
pub(super) const fn table(kind: Kind) -> &'static [PropSpec] {
    match kind {
        Kind::Tab => TAB_PROPS,
        Kind::Break => BREAK_PROPS,
        Kind::Bookmark => BOOKMARK_PROPS,
        Kind::Field => FIELD_PROPS,
        Kind::Toc => TOC_PROPS,
        Kind::Watermark => WATERMARK_PROPS,
        Kind::TextBox => TEXT_BOX_PROPS,
        _ => &[],
    }
}

fn text_value<'a>(text: &str) -> NodeValue<'a> {
    NodeValue::Text(Cow::Owned(text.to_owned()))
}

/// The OfficeCLI spelling of a break type.
const fn break_token(kind: RunBreakType) -> &'static str {
    match kind {
        RunBreakType::Page => "page",
        RunBreakType::Column => "column",
        RunBreakType::TextWrapping => "textWrapping",
    }
}

/// The switch argument or presence of a `TOC` instruction: `(present,
/// argument)` for switch `name`.
fn toc_switch(instruction: &str, name: char) -> (bool, Option<String>) {
    let mut chars = instruction.char_indices().peekable();
    while let Some((at, character)) = chars.next() {
        if character != '\\' {
            continue;
        }
        let Some((_, switch)) = chars.next() else {
            break;
        };
        if !switch.eq_ignore_ascii_case(&name) {
            continue;
        }
        let rest = instruction[at + 2..].trim_start();
        let argument = rest
            .strip_prefix('"')
            .and_then(|quoted| quoted.find('"').map(|end| quoted[..end].to_owned()));
        return (true, argument);
    }
    (false, None)
}

/// The value of property `index` of a mark node.
pub(super) fn value<'a>(mark: &MarkValue, index: usize) -> Option<NodeValue<'a>> {
    match mark {
        MarkValue::Tab { stop, .. } => match TAB_PROPS.get(index)?.name {
            "pos" => Some(NodeValue::Int(stop.position.0)),
            "val" => Some(NodeValue::Token(stop.kind.as_str())),
            "leader" => stop.leader.map(|leader| NodeValue::Token(leader.as_str())),
            _ => None,
        },
        MarkValue::Break { kind, .. } => match BREAK_PROPS.get(index)?.name {
            "breakType" => Some(NodeValue::Token(break_token(*kind))),
            _ => None,
        },
        MarkValue::Bookmark(mark) => match BOOKMARK_PROPS.get(index)?.name {
            "name" => Some(text_value(mark.name())),
            "id" => Some(NodeValue::Text(Cow::Owned(mark.id().to_string()))),
            _ => None,
        },
        MarkValue::Field(field) => match FIELD_PROPS.get(index)?.name {
            "fieldType" => Some(NodeValue::Text(Cow::Owned(
                field.field_type().to_ascii_lowercase(),
            ))),
            "instruction" => Some(text_value(field.instruction().trim())),
            "fldLock" => field.is_locked().then_some(NodeValue::Bool(true)),
            "evaluated" => Some(NodeValue::Bool(field.result().is_some())),
            _ => None,
        },
        MarkValue::Toc(field) => match TOC_PROPS.get(index)?.name {
            "levels" => toc_switch(field.instruction(), 'o')
                .1
                .map(|levels| text_value(&levels)),
            "hyperlinks" => Some(NodeValue::Bool(toc_switch(field.instruction(), 'h').0)),
            "pageNumbers" => Some(NodeValue::Bool(!toc_switch(field.instruction(), 'n').0)),
            _ => None,
        },
        MarkValue::Watermark(mark) => match WATERMARK_PROPS.get(index)?.name {
            "text" => Some(text_value(mark.get_text())),
            "color" => (!mark.color().is_empty()).then(|| {
                let color = mark.color().trim_start_matches('#');
                NodeValue::Text(Cow::Owned(format!("#{}", color.to_ascii_uppercase())))
            }),
            "font" => (!mark.font().is_empty()).then(|| text_value(mark.font())),
            "rotation" => Some(NodeValue::Int(match mark.layout() {
                WatermarkLayout::Diagonal => 315,
                WatermarkLayout::Horizontal => 0,
            })),
            "opacity" => mark
                .semitransparent()
                .then_some(NodeValue::Text(Cow::Borrowed(".5"))),
            "size" => Some(NodeValue::Text(Cow::Owned(format!(
                "{}pt",
                crate::value::DotNetDouble(mark.font_size())
            )))),
            _ => None,
        },
        MarkValue::TextBox(text_box) => match TEXT_BOX_PROPS.get(index)?.name {
            "name" => text_box.name.as_deref().map(text_value),
            "geometry" => text_box
                .preset
                .as_ref()
                .map(|preset| NodeValue::Text(Cow::Owned(format!("{preset:?}")))),
            _ => None,
        },
    }
}

/// The text of a mark node: a field's result, a watermark's or text box's
/// text.
pub(super) fn text(mark: &MarkValue) -> Option<String> {
    match mark {
        MarkValue::Field(field) | MarkValue::Toc(field) => field.result().map(str::to_owned),
        MarkValue::Watermark(mark) => Some(mark.get_text().to_owned()),
        MarkValue::TextBox(text_box) => Some(text_box.text()),
        MarkValue::Tab { .. } | MarkValue::Break { .. } | MarkValue::Bookmark(_) => None,
    }
}

/// A tab stop's kind for an OfficeCLI `val`.
pub(super) fn tab_kind(value: &str) -> Option<TabKind> {
    TabKind::from_xml(value)
}

/// A tab stop's leader for an OfficeCLI `leader`.
pub(super) fn tab_leader(value: &str) -> Option<TabLeader> {
    TabLeader::from_xml(value)
}

/// A paragraph's tab stops with stop `index` replaced by `stop`, or removed
/// for `None`.
pub(super) fn with_stop(
    paragraph: &mut direct::Paragraph,
    index: usize,
    stop: Option<TabStop>,
) -> Result<()> {
    let mut stops = paragraph.tabs.take().unwrap_or_default();
    if index >= stops.len() {
        return Err(Error::new(Code::PathNotFound, "the tab stop is gone"));
    }
    match stop {
        Some(stop) => stops[index] = stop,
        None => {
            stops.remove(index);
        },
    }
    paragraph.tabs = (!stops.is_empty()).then_some(stops);
    Ok(())
}
