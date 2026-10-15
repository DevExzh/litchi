//! SmartArt in Word documents (record 0295, record 0284 follow-up 0284a).
//!
//! - SmartArt graphics are addressed document-wide in document order
//!   (`/smartart[N]`, also `/body/smartart[N]`), as OfficeCLI addresses
//!   charts. `get` reads the leaf inventory (`Package::smart_arts`): the
//!   layout (Litchi's token for a built-in layout, else the last segment of
//!   the layout identifier), the layout identifier, the node count and the
//!   outline (`text`: one node per line, one tab per level).
//! - `add(/body/p[N], smartart, {layout, text, name, alt, width, height})`
//!   builds an authored spec (`litchi_drawingml::diagram::authoring`, the
//!   PPTX vocabulary of record 0284) and inserts it as an inline drawing
//!   through the leaf `Edit::insert_smartart_at`: the five diagram parts,
//!   their relationships and the run, in one commit.
//! - `set` and `remove` of DOCX SmartArt are gaps (follow-up in record
//!   0295): the leaf has no part replacement or drawing removal for
//!   diagrams yet.

use std::borrow::Cow;
use std::sync::Arc;

use litchi_core::FileFormat;
use litchi_core::selector::{Path, Predicate, Step as PathStep};
use litchi_docx::document::NewSmartArt;
use litchi_docx::smartart::Diagram;
use litchi_drawingml::diagram::authoring::{Layout, Spec};

use super::element::{element_path, refusal};
use super::node::{Address, Node, Step, Value};
use super::spec::Kind;
use super::write::{Commit, WorkEdit};
use super::{Docx, DocxEditor, DocxOwner};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::node::Value as NodeValue;
use crate::node::spec::PropSpec;
use crate::registry::{Access, Alias, Constraint, Entry, Operation, Table};
use crate::value::{LengthSpec, Scalar, ValueType};

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Docx)
}

const fn write_only(name: &'static str, ty: ValueType) -> PropSpec {
    let mut spec = PropSpec::new(name, ty);
    spec.access.get = false;
    spec
}

/// SmartArt assignments (add only).
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum SmartArtSetter {
    Layout,
    Text,
    Name,
    Alt,
    Width,
    Height,
}

const EXTENT: Constraint = Constraint::Length {
    min: crate::value::Length::from_emu(1),
    max: crate::value::Length::from_emu(22 * 914_400),
};

const fn entry(
    spec: PropSpec,
    access: Access<SmartArtSetter>,
    constraint: Constraint,
) -> Entry<SmartArtSetter> {
    Entry {
        spec,
        access,
        constraint,
        clear: false,
    }
}

pub(crate) const SMARTART_WRITE_TABLE: Table<SmartArtSetter, 8, 1> = Table::new(
    [
        entry(
            PropSpec::new("layout", ValueType::Text),
            Access::Add(SmartArtSetter::Layout),
            Constraint::NotEmpty,
        ),
        Entry::read(PropSpec::new("layoutId", ValueType::Text)),
        Entry::read(PropSpec::new("nodeCount", ValueType::Int)),
        entry(
            PropSpec::new("text", ValueType::Text),
            Access::Add(SmartArtSetter::Text),
            Constraint::NotEmpty,
        ),
        entry(
            write_only("name", ValueType::Text),
            Access::Add(SmartArtSetter::Name),
            Constraint::NotEmpty,
        ),
        entry(
            write_only("alt", ValueType::Text),
            Access::Add(SmartArtSetter::Alt),
            Constraint::None,
        ),
        entry(
            write_only("width", ValueType::Length(LengthSpec::EXTENT)),
            Access::Add(SmartArtSetter::Width),
            EXTENT,
        ),
        entry(
            write_only("height", ValueType::Length(LengthSpec::EXTENT)),
            Access::Add(SmartArtSetter::Height),
            EXTENT,
        ),
    ],
    [Alias {
        name: "description",
        entry: 5,
    }],
);

/// The SmartArt element's properties.
pub(crate) const SMARTART_PROPS: &[PropSpec] = &SMARTART_WRITE_TABLE.readable();

/// One SmartArt graphic of the document, as a node carries it.
pub(crate) struct SmartArtValue {
    pub(super) diagram: Diagram,
}

/// The SmartArt graphics of the main document, in document order.
pub(super) fn document_smartarts(docx: &Docx) -> Result<Vec<Diagram>> {
    let result = match docx.owner() {
        DocxOwner::Eager(package) => package.smart_arts().map_err(Error::from),
        DocxOwner::Source(package) => docx
            .promoted(package)
            .and_then(|package| package.smart_arts().map_err(Error::from)),
    };
    docx.owner().finish(result)
}

/// The node of the SmartArt graphic at `ordinal`.
pub(super) fn smartart_node(docx: &Docx, ordinal: u32, diagram: Diagram) -> Node<'_> {
    Node::new(
        docx,
        Kind::SmartArt,
        Value::SmartArt(Arc::new(SmartArtValue { diagram })),
        Address::root().with(Step::new(Kind::SmartArt, ordinal)),
    )
}

/// The document-wide SmartArt a `/smartart[N]` (or `/body/smartart[N]`)
/// step picks.
pub(super) fn pick_document_smartart<'f>(
    docx: &'f Docx,
    select: &crate::path::navigate::Select<'_>,
) -> Result<crate::path::navigate::Lookup<Node<'f>, crate::node::spec::GapSpec>> {
    use crate::path::navigate::{Lookup, Select};
    let diagrams = document_smartarts(docx)?;
    let picked = match *select {
        Select::None => Some(0),
        Select::Ordinal(position) => Some(position.get()),
        Select::Last => diagrams.len().checked_sub(1),
        _ => return Ok(Lookup::NotApplicable),
    };
    Ok(
        match picked.and_then(|index| Some((index, diagrams.into_iter().nth(index)?))) {
            Some((index, diagram)) => Lookup::Found(smartart_node(
                docx,
                u32::try_from(index).unwrap_or(u32::MAX),
                diagram,
            )),
            None => Lookup::Missing,
        },
    )
}

/// The layout token of a graphic: Litchi's name of a built-in layout, else
/// the last segment of its layout identifier.
fn layout_id(diagram: &Diagram) -> Option<&str> {
    diagram
        .data
        .document_point()
        .and_then(|point| point.layout_type_id.as_deref())
        .or_else(|| {
            diagram
                .layout
                .as_ref()
                .and_then(|layout| layout.unique_id.as_deref())
        })
}

/// The outline of a graphic's content nodes: one per line, one tab per
/// level below the first.
fn outline_text(diagram: &Diagram) -> String {
    fn walk(nodes: &[litchi_drawingml::diagram::DiagramNode], depth: usize, out: &mut String) {
        for node in nodes {
            if !out.is_empty() {
                out.push('\n');
            }
            for _ in 0..depth {
                out.push('\t');
            }
            out.push_str(&node.text);
            walk(&node.children, depth + 1, out);
        }
    }
    let mut out = String::new();
    walk(&diagram.node_tree(), 0, &mut out);
    out
}

fn count(nodes: &[litchi_drawingml::diagram::DiagramNode]) -> usize {
    nodes.iter().map(|node| 1 + count(&node.children)).sum()
}

/// The value of the SmartArt table's entry at `index`.
pub(super) fn smartart_value(value: &SmartArtValue, index: usize) -> Option<NodeValue<'static>> {
    let diagram = &value.diagram;
    let text = |text: &str| NodeValue::Text(Cow::Owned(text.to_owned()));
    match index {
        0 => layout_id(diagram).map(|id| match Layout::from_unique_id(id) {
            Some(layout) => NodeValue::Text(Cow::Borrowed(layout.name())),
            None => text(id.rsplit('/').next().unwrap_or(id)),
        }),
        1 => layout_id(diagram).map(text),
        2 => Some(NodeValue::Int(
            i64::try_from(count(&diagram.node_tree())).unwrap_or(i64::MAX),
        )),
        3 => Some(text(&outline_text(diagram))),
        _ => None,
    }
}

/// The node text of a SmartArt graphic: its outline.
pub(super) fn smartart_text(value: &SmartArtValue) -> String {
    outline_text(&value.diagram)
}

fn text(value: Option<Scalar<'_>>) -> Option<String> {
    match value {
        Some(Scalar::Text(value) | Scalar::Token(value)) => Some(value.to_owned()),
        _ => None,
    }
}

fn emu(value: Option<Scalar<'_>>) -> Option<i64> {
    match value {
        Some(Scalar::Length(length)) => Some(length.emu()),
        _ => None,
    }
}

impl Docx {
    /// Add a SmartArt graphic at the end of a paragraph: `/body/p[N]` (or a
    /// paragraph in a cell or a content control).
    pub(super) fn add_smartart(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        let paragraph = element_path(parent)?;
        let mut layout = Layout::List;
        let mut outline = None;
        let mut name = None;
        let mut alt = None;
        let mut size = (None, None);
        let mut assigned = [false; 8];
        for (key, value) in node.props().iter() {
            let (index, entry) = SMARTART_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            if std::mem::replace(&mut assigned[index], true) {
                return Err(error(
                    Code::InvalidArgument,
                    "a SmartArt property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            entry.prepare(Operation::Add, value, |setter, scalar| {
                match setter {
                    SmartArtSetter::Layout => {
                        let token = text(scalar).unwrap_or_default();
                        layout = Layout::parse(&token).ok_or_else(|| {
                            error(Code::InvalidValue, "unknown SmartArt layout")
                                .with_property("layout")
                                .with_suggestion(
                                    "use list, process, cycle, hierarchy, pyramid, matrix or venn",
                                )
                        })?;
                    },
                    SmartArtSetter::Text => outline = text(scalar),
                    SmartArtSetter::Name => name = text(scalar),
                    SmartArtSetter::Alt => alt = text(scalar),
                    SmartArtSetter::Width => size.0 = emu(scalar),
                    SmartArtSetter::Height => size.1 = emu(scalar),
                }
                Ok(())
            })?;
        }
        let outline = outline.ok_or_else(|| {
            error(Code::MissingProperty, "a SmartArt graphic needs `text`")
                .with_property("text")
                .with_suggestion("give the outline: one node per line, one tab per level")
        })?;
        let spec = Spec::parse_outline(layout, &outline).map_err(|source| {
            error(Code::InvalidValue, "the SmartArt outline is not valid")
                .with_property("text")
                .with_suggestion(source.to_string())
        })?;
        for value in [&name, &alt].into_iter().flatten() {
            if value.contains('\0') {
                return Err(error(
                    Code::InvalidValue,
                    "SmartArt text cannot contain NUL",
                ));
            }
        }
        let mut smartart = NewSmartArt::new(spec).with_size(size.0, size.1);
        if let Some(name) = name {
            smartart = smartart.with_name(name);
        }
        if let Some(alt) = alt {
            smartart = smartart.with_description(alt);
        }
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("SmartArt needs the owning package editor"));
        };
        let relationship = eager
            .insert_smartart_at(&paragraph, None, &smartart)
            .map_err(Error::from)?;
        let commit = edit.commit(self)?;
        let ordinal = document_smartarts(commit.file())?
            .iter()
            .position(|diagram| diagram.data_relationship_id == relationship)
            .ok_or_else(|| error(Code::InternalError, "the added SmartArt did not read back"))?;
        let path = Path::from_steps(vec![
            PathStep::new("smartart").with(Predicate::Index(litchi_core::Position::new(ordinal))),
        ]);
        Ok((commit, path))
    }
}
