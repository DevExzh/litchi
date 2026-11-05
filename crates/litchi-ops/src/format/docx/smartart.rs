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
//! - `set` of a graphic's `layout` and `text` (the whole outline), and
//!   `remove` of the graphic with its five parts and their relationships
//!   (record 0335).
//! - `/smartart[N]/node[K]` (nested: `.../node[K]/node[J]`) reads one
//!   content node (`text`, `level`); `add` under a graphic or a node with
//!   `text`, `set text` and `remove` (with its descendants), as on PPTX
//!   (record 0284).
//!
//! The cached drawing stays consistent with the data model, by the rules
//! of record 0284 §3: a graphic Litchi authored is regenerated whole; a
//! node text edit of any other graphic is made in place in the data model
//! and in the drawing shapes that show exactly that node (else the drawing
//! is removed so consumers lay the graphic out again); adding or removing
//! nodes of any other graphic is refused (`unsupported_capability`), and
//! setting `layout` or `text` rebuilds it as a Litchi graphic.

use std::borrow::Cow;
use std::sync::Arc;

use std::sync::OnceLock;

use litchi_core::FileFormat;
use litchi_core::selector::{Path, Predicate, Step as PathStep};
use litchi_docx::document::{NewSmartArt, SmartArtParts};
use litchi_docx::smartart::Diagram;
use litchi_drawingml::diagram::Limits;
use litchi_drawingml::diagram::authoring::{
    Layout, Node as SpecNode, Spec, colors_xml, quick_style_xml,
};
use litchi_drawingml::diagram::edit::{self as diagram_edit, Outline};

use super::element::{element_path, refusal};
use super::node::{Address, Node, Step, Value};
use super::spec::Kind;
use super::write::{Commit, WorkEdit};
use super::{Docx, DocxEditor, DocxOwner};
use crate::error::{Code, Error, Result};
use crate::new_node::NewNode;
use crate::node::Value as NodeValue;
use crate::node::spec::PropSpec;
use crate::placement::Placement;
use crate::props::Props;
use crate::registry::{Access, Alias, Constraint, Entry, Operation, Table};
use crate::value::{LengthSpec, Scalar, ValueType};

fn error(code: Code, message: &'static str) -> Error {
    Error::new(code, message).with_format(FileFormat::Docx)
}

fn drawing_code(error: &litchi_drawingml::Error, fallback: Code) -> Code {
    match error {
        litchi_drawingml::Error::Limit { .. } => Code::LimitExceeded,
        litchi_drawingml::Error::InvalidLimit { .. } => Code::InvalidArgument,
        litchi_drawingml::Error::PatchConflict => Code::InvalidInput,
        _ => fallback,
    }
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

const fn entry<S: Copy>(spec: PropSpec, access: Access<S>, constraint: Constraint) -> Entry<S> {
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
            Access::SetAndAdd(SmartArtSetter::Layout),
            Constraint::NotEmpty,
        ),
        Entry::read(PropSpec::new("layoutId", ValueType::Text)),
        Entry::read(PropSpec::new("nodeCount", ValueType::Int)),
        entry(
            PropSpec::new("text", ValueType::Text),
            Access::SetAndAdd(SmartArtSetter::Text),
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

/// SmartArt node assignments.
#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum NodeSetter {
    Text,
}

/// A content node of a graphic (record 0335): its level is read; its text
/// is the node's text, written by `add` and `set`, as on PPTX.
pub(crate) const NODE_WRITE_TABLE: Table<NodeSetter, 2, 0> = Table::new(
    [
        Entry::read(PropSpec::new("level", ValueType::Int)),
        entry(
            write_only("text", ValueType::Text),
            Access::SetAndAdd(NodeSetter::Text),
            Constraint::None,
        ),
    ],
    [],
);

/// The node element's properties.
pub(crate) const NODE_PROPS: &[PropSpec] = &NODE_WRITE_TABLE.readable();

/// One SmartArt graphic of the document, as a node carries it.
pub(crate) struct SmartArtValue {
    pub(super) diagram: Diagram,
    /// The graphic's position among the document's graphics.
    pub(super) ordinal: u32,
    /// The data model's content nodes, read on first use.
    outline: OnceLock<Arc<Outline>>,
}

impl SmartArtValue {
    /// The data model's content nodes, read once.
    pub(super) fn outline(&self, docx: &Docx) -> Result<Arc<Outline>> {
        if let Some(outline) = self.outline.get() {
            return Ok(Arc::clone(outline));
        }
        let xml = data_xml(docx, &self.diagram)?;
        let outline = Arc::new(Outline::read(&xml, &Limits::default()).map_err(|source| {
            error(
                drawing_code(&source, Code::CorruptFile),
                "the SmartArt data model does not read",
            )
            .with_suggestion(source.to_string())
        })?);
        Ok(Arc::clone(self.outline.get_or_init(|| outline)))
    }
}

/// One content node of a graphic.
pub(crate) struct DiagramNodeValue {
    pub(super) smartart: Arc<SmartArtValue>,
    pub(super) outline: Arc<Outline>,
    /// The node's entry in the outline.
    pub(super) index: u32,
}

/// The data-model part of a graphic, read through the owner.
fn data_xml(docx: &Docx, diagram: &Diagram) -> Result<Vec<u8>> {
    let read = |package: &litchi_docx::Package| -> Result<Vec<u8>> {
        let name = litchi_opc::PackURI::new(diagram.data_part_name.clone())
            .map_err(|_| error(Code::CorruptFile, "invalid SmartArt data part name"))?;
        Ok(package
            .opc_package()
            .get_part(&name)
            .map_err(|_| error(Code::CorruptFile, "the SmartArt data part is missing"))?
            .blob()
            .to_vec())
    };
    let result = match docx.owner() {
        DocxOwner::Eager(package) => read(package),
        DocxOwner::Source(package) => docx.promoted(package).and_then(|package| read(&package)),
    };
    docx.owner().finish(result)
}

/// The SmartArt graphics of the main document, in document order, read
/// once per handle.
pub(super) fn document_smartarts(docx: &Docx) -> Result<Arc<[Arc<SmartArtValue>]>> {
    if let Some(graphics) = docx.parts().smartarts.get() {
        return Ok(Arc::clone(graphics));
    }
    let result = match docx.owner() {
        DocxOwner::Eager(package) => package.smart_arts().map_err(Error::from),
        DocxOwner::Source(package) => docx
            .promoted(package)
            .and_then(|package| package.smart_arts().map_err(Error::from)),
    };
    let diagrams = docx.owner().finish(result)?;
    let graphics: Arc<[Arc<SmartArtValue>]> = diagrams
        .into_iter()
        .enumerate()
        .map(|(index, diagram)| {
            Arc::new(SmartArtValue {
                diagram,
                ordinal: u32::try_from(index).unwrap_or(u32::MAX),
                outline: OnceLock::new(),
            })
        })
        .collect();
    Ok(Arc::clone(docx.parts().smartarts.get_or_init(|| graphics)))
}

/// The node of a SmartArt graphic.
pub(super) fn smartart_node(docx: &Docx, smartart: Arc<SmartArtValue>) -> Node<'_> {
    let ordinal = smartart.ordinal;
    Node::new(
        docx,
        Kind::SmartArt,
        Value::SmartArt(smartart),
        Address::root().with(Step::new(Kind::SmartArt, ordinal)),
    )
}

/// The graphic, its outline and the outline entry (`None`: the graphic
/// itself) a SmartArt or node value names.
fn graphic_of(
    docx: &Docx,
    value: &Value,
) -> Result<Option<(Arc<SmartArtValue>, Arc<Outline>, Option<usize>)>> {
    Ok(match value {
        Value::SmartArt(smartart) => Some((Arc::clone(smartart), smartart.outline(docx)?, None)),
        Value::DiagramNode(node) => Some((
            Arc::clone(&node.smartart),
            Arc::clone(&node.outline),
            Some(node.index as usize),
        )),
        _ => None,
    })
}

/// The child nodes of a graphic or node, in order.
pub(super) fn node_children(docx: &Docx, parent: &Node<'_>) -> Result<Vec<(Value, Step)>> {
    let Some((smartart, outline, at)) = graphic_of(docx, &parent.value)? else {
        return Ok(Vec::new());
    };
    Ok(outline
        .children(at)
        .enumerate()
        .map(|(position, index)| {
            (
                Value::DiagramNode(Arc::new(DiagramNodeValue {
                    smartart: Arc::clone(&smartart),
                    outline: Arc::clone(&outline),
                    index: u32::try_from(index).unwrap_or(u32::MAX),
                })),
                Step::new(Kind::Node, u32::try_from(position).unwrap_or(u32::MAX)),
            )
        })
        .collect())
}

/// The child node a `node[K]` step picks under a graphic or node.
pub(super) fn pick_node<'f>(
    docx: &'f Docx,
    parent: &Node<'f>,
    select: &crate::path::navigate::Select<'_>,
) -> Result<crate::path::navigate::Lookup<Node<'f>, crate::node::spec::GapSpec>> {
    use crate::path::navigate::{Lookup, Select};
    let children = node_children(docx, parent)?;
    let picked = match *select {
        Select::None => Some(0),
        Select::Ordinal(position) => Some(position.get()),
        Select::Last => children.len().checked_sub(1),
        _ => return Ok(Lookup::NotApplicable),
    };
    Ok(
        match picked.and_then(|index| children.into_iter().nth(index)) {
            Some((value, step)) => Lookup::Found(Node::new(
                docx,
                Kind::Node,
                value,
                parent.child_address().with(step),
            )),
            None => Lookup::Missing,
        },
    )
}

/// The value of the node table's entry at `index`: the 1-based level.
pub(super) fn node_value(node: &DiagramNodeValue, index: usize) -> Option<NodeValue<'static>> {
    let entry = node.outline.entries.get(node.index as usize)?;
    (index == 0).then(|| NodeValue::Int(i64::from(entry.depth) + 1))
}

/// The node's text.
pub(super) fn node_text(node: &DiagramNodeValue) -> Option<String> {
    node.outline
        .entries
        .get(node.index as usize)
        .map(|entry| entry.text.clone())
}

/// The number of a graphic's or node's child nodes.
pub(super) fn node_count(docx: &Docx, value: &Value) -> Result<usize> {
    Ok(match graphic_of(docx, value)? {
        Some((_, outline, at)) => outline.children(at).count(),
        None => 0,
    })
}

/// The document-wide SmartArt a `/smartart[N]` (or `/body/smartart[N]`)
/// step picks.
pub(super) fn pick_document_smartart<'f>(
    docx: &'f Docx,
    select: &crate::path::navigate::Select<'_>,
) -> Result<crate::path::navigate::Lookup<Node<'f>, crate::node::spec::GapSpec>> {
    use crate::path::navigate::{Lookup, Select};
    let graphics = document_smartarts(docx)?;
    let picked = match *select {
        Select::None => Some(0),
        Select::Ordinal(position) => Some(position.get()),
        Select::Last => graphics.len().checked_sub(1),
        _ => return Ok(Lookup::NotApplicable),
    };
    Ok(match picked.and_then(|index| graphics.get(index)) {
        Some(smartart) => Lookup::Found(smartart_node(docx, Arc::clone(smartart))),
        None => Lookup::Missing,
    })
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
        let spec = Spec::parse_outline(layout, &outline, &Limits::default()).map_err(|source| {
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
            .position(|graphic| graphic.diagram.data_relationship_id == relationship)
            .ok_or_else(|| error(Code::InternalError, "the added SmartArt did not read back"))?;
        let path = Path::from_steps(vec![
            PathStep::new("smartart").with(Predicate::Index(litchi_core::Position::new(ordinal))),
        ]);
        Ok((commit, path))
    }
}

/// The default extent of a graphic whose drawing gives none: 16 cm × 9 cm,
/// the insertion default.
const DEFAULT_EXTENT: (i64, i64) = (5_760_000, 3_240_000);

fn leaf_error(code: Code, source: impl std::fmt::Display) -> Error {
    Error::new(code, source.to_string()).with_format(FileFormat::Docx)
}

fn parse_layout(value: &str) -> Result<Layout> {
    Layout::parse(value).ok_or_else(|| {
        error(Code::InvalidValue, "unknown SmartArt layout")
            .with_property("layout")
            .with_suggestion("use list, process, cycle, hierarchy, pyramid, matrix or venn")
    })
}

fn parse_outline(layout: Layout, value: &str) -> Result<Spec> {
    let spec = Spec::parse_outline(layout, value, &Limits::default())
        .map_err(|source| leaf_error(Code::InvalidValue, source).with_property("text"))?;
    if spec.nodes().is_empty() {
        return Err(
            error(Code::InvalidValue, "the SmartArt outline has no nodes")
                .with_property("text")
                .with_suggestion("give one node per line, children indented by a tab"),
        );
    }
    Ok(spec)
}

/// The child-index path of an outline entry.
fn spec_path(outline: &Outline, index: usize) -> Vec<usize> {
    let mut out = Vec::new();
    let mut current = Some(index);
    while let Some(index) = current {
        let parent = outline.entries[index].parent;
        out.push(
            outline.entries[..index]
                .iter()
                .filter(|entry| entry.parent == parent)
                .count(),
        );
        current = parent;
    }
    out.reverse();
    out
}

/// The sibling list at `path` (the children of the node it names; the top
/// level for an empty path).
fn siblings_at<'s>(spec: &'s mut Spec, path: &[usize]) -> Option<&'s mut Vec<SpecNode>> {
    let mut list = spec.nodes_mut();
    for &index in path {
        list = &mut list.get_mut(index)?.children;
    }
    Some(list)
}

/// One structural or text change of a graphic.
enum Change {
    /// Replace a node's text.
    Text { index: usize, text: String },
    /// Insert a node under `parent` (an outline entry; `None`: top level)
    /// at a sibling position (`None`: last).
    Insert {
        parent: Option<usize>,
        position: Option<usize>,
        text: String,
    },
    /// Remove a node and its descendants.
    Remove { index: usize },
    /// Rebuild with a layout and/or a new outline.
    Rebuild {
        layout: Option<Layout>,
        outline: Option<String>,
    },
}

fn missing_node() -> Error {
    error(Code::PathNotFound, "the SmartArt node is missing")
}

fn foreign_structure() -> Error {
    error(
        Code::UnsupportedCapability,
        "nodes can be added to or removed from SmartArt Litchi authored only",
    )
    .with_suggestion(
        "set layout on the smartart to rebuild it as a Litchi graphic from its outline (formatting of the nodes is not kept), or edit node text only",
    )
}

/// The parts after `change`, for a drawing of `extent` (record 0284 §3).
fn apply(parts: &SmartArtParts, change: Change, extent: (i64, i64)) -> Result<SmartArtParts> {
    let limits = Limits::default();
    let corrupt = |source: litchi_drawingml::Error| {
        leaf_error(drawing_code(&source, Code::CorruptFile), source)
    };
    let invalid = |source: litchi_drawingml::Error| {
        leaf_error(drawing_code(&source, Code::InvalidValue), source)
    };
    let outline = Outline::read(&parts.data, &limits).map_err(corrupt)?;
    let authored =
        Spec::recognize(&outline, &parts.data, &parts.layout, &limits).map_err(corrupt)?;
    let regenerate = |spec: &Spec, all: bool| -> Result<SmartArtParts> {
        spec.validate().map_err(invalid)?;
        let relationship = outline.drawing_relationship.as_deref().unwrap_or("rId0");
        Ok(SmartArtParts {
            data: spec
                .data_xml(Some(relationship), &limits)
                .map_err(invalid)?
                .into_bytes(),
            layout: spec.layout_xml(&limits).map_err(invalid)?.into_bytes(),
            quick_style: if all {
                quick_style_xml(&limits).map_err(invalid)?.into_bytes()
            } else {
                parts.quick_style.clone()
            },
            colors: if all {
                colors_xml(&limits).map_err(invalid)?.into_bytes()
            } else {
                parts.colors.clone()
            },
            drawing: Some(
                spec.drawing_xml(extent.0, extent.1, &limits)
                    .map_err(invalid)?
                    .into_bytes(),
            ),
        })
    };
    match change {
        Change::Rebuild {
            layout,
            outline: text,
        } => {
            let base = authored
                .clone()
                .unwrap_or_else(|| Spec::from_outline(layout.unwrap_or(Layout::List), &outline));
            let layout = layout.unwrap_or(base.layout());
            let spec = match text {
                Some(text) => parse_outline(layout, &text)?,
                None => base.with_layout(layout),
            };
            // A Litchi graphic keeps its style parts; a rebuilt one takes
            // Litchi's.
            regenerate(&spec, authored.is_none())
        },
        Change::Text { index, text } => {
            if let Some(mut spec) = authored {
                let path = spec_path(&outline, index);
                let (last, parents) = path.split_last().ok_or_else(missing_node)?;
                let list = siblings_at(&mut spec, parents).ok_or_else(missing_node)?;
                list[*last].text = text;
                return regenerate(&spec, false);
            }
            let id = &outline.entries.get(index).ok_or_else(missing_node)?.id;
            let data = diagram_edit::set_text(&parts.data, id, &text, &limits).map_err(invalid)?;
            if data == parts.data {
                return Ok(parts.clone());
            }
            let drawing = match &parts.drawing {
                Some(drawing) => {
                    diagram_edit::drawing_text(drawing, &parts.data, id, &text, &limits)
                        .map_err(invalid)?
                },
                None => None,
            };
            Ok(SmartArtParts {
                data,
                drawing,
                ..parts.clone()
            })
        },
        Change::Insert {
            parent,
            position,
            text,
        } => {
            let mut spec = authored.ok_or_else(foreign_structure)?;
            let path = parent
                .map(|index| spec_path(&outline, index))
                .unwrap_or_default();
            let list = siblings_at(&mut spec, &path).ok_or_else(missing_node)?;
            let position = position.unwrap_or(list.len());
            if position > list.len() {
                return Err(error(
                    Code::OutOfRange,
                    "the node position is past the last sibling",
                ));
            }
            list.insert(position, SpecNode::new(text));
            regenerate(&spec, false)
        },
        Change::Remove { index } => {
            let mut spec = authored.ok_or_else(foreign_structure)?;
            let path = spec_path(&outline, index);
            let (last, parents) = path.split_last().ok_or_else(missing_node)?;
            let list = siblings_at(&mut spec, parents).ok_or_else(missing_node)?;
            list.remove(*last);
            regenerate(&spec, false)
        },
    }
}

fn scalar_text(value: Option<Scalar<'_>>) -> String {
    text(value).unwrap_or_default()
}

/// The canonical path of a graphic's node: `/smartart[N]` and the node's
/// child-index path, then `last` when given.
fn node_path(ordinal: u32, steps: &[usize]) -> Path<'static> {
    let mut path =
        vec![
            PathStep::new("smartart").with(Predicate::Index(litchi_core::Position::new(
                ordinal as usize,
            ))),
        ];
    path.extend(steps.iter().map(|position| {
        PathStep::new("node").with(Predicate::Index(litchi_core::Position::new(*position)))
    }));
    Path::from_steps(path)
}

impl Docx {
    /// Stage `change` on `smartart` in one edit and commit it.
    fn change_smartart(
        &self,
        editor: DocxEditor,
        smartart: &SmartArtValue,
        change: Change,
    ) -> Result<Commit> {
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("SmartArt needs the owning package editor"));
        };
        let relationship = &smartart.diagram.data_relationship_id;
        let parts = eager
            .smartart_parts(relationship)
            .map_err(Error::from)?
            .ok_or_else(|| error(Code::PathNotFound, "the SmartArt graphic is missing"))?;
        let extent = smartart.diagram.extent.unwrap_or(DEFAULT_EXTENT);
        let next = apply(&parts, change, extent)?;
        if next != parts {
            eager
                .set_smartart(relationship, next)
                .map_err(Error::from)?;
        }
        edit.commit(self)
    }

    /// A `set` whose target is one graphic or one of its nodes; `Err` hands
    /// the editor back for other targets.
    pub(super) fn try_set_smartart(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
        props: &Props<'_>,
    ) -> Result<std::result::Result<Commit, DocxEditor>> {
        if nodes.is_empty()
            || nodes
                .iter()
                .any(|node| !matches!(node.kind, Kind::SmartArt | Kind::Node))
        {
            return Ok(Err(editor));
        }
        if nodes.len() != 1 {
            return Err(error(
                Code::UnsupportedCapability,
                "a SmartArt set targets one graphic or node",
            ));
        }
        let node = &nodes[0];
        let Some((smartart, _, at)) = graphic_of(self, &node.value)? else {
            return Err(error(
                Code::InternalError,
                "the SmartArt node holds no graphic",
            ));
        };
        let change = match at {
            Some(index) => {
                let mut value = None;
                for (key, scalar) in props.iter() {
                    let (_, entry) = NODE_WRITE_TABLE
                        .resolve(key.as_str())
                        .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                    entry.prepare(Operation::Set, scalar, |setter, scalar| {
                        match setter {
                            NodeSetter::Text => value = Some(scalar_text(scalar)),
                        }
                        Ok(())
                    })?;
                }
                value.map(|text| Change::Text { index, text })
            },
            None => {
                let mut layout = None;
                let mut outline = None;
                for (key, scalar) in props.iter() {
                    let (_, entry) = SMARTART_WRITE_TABLE
                        .resolve(key.as_str())
                        .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                    entry.prepare(Operation::Set, scalar, |setter, scalar| {
                        match setter {
                            SmartArtSetter::Layout => {
                                layout = Some(parse_layout(&scalar_text(scalar))?);
                            },
                            SmartArtSetter::Text => outline = Some(scalar_text(scalar)),
                            _ => unreachable!("SmartArt set admission"),
                        }
                        Ok(())
                    })?;
                }
                (layout.is_some() || outline.is_some())
                    .then_some(Change::Rebuild { layout, outline })
            },
        };
        match change {
            Some(change) => self.change_smartart(editor, &smartart, change).map(Ok),
            None => {
                self.owner().check_source()?;
                WorkEdit::from(editor)?.commit(self).map(Ok)
            },
        }
    }

    /// `add(<smartart or node>, node, {text})` (record 0335).
    pub(super) fn add_smartart_node(
        &self,
        editor: DocxEditor,
        parent: &Node<'_>,
        node: &NewNode<'_>,
    ) -> Result<(Commit, Path<'static>)> {
        let Some((smartart, outline, at)) = graphic_of(self, &parent.value)? else {
            return Err(error(
                Code::UnsupportedType,
                "SmartArt nodes require a smartart or node parent",
            ));
        };
        let count = outline.children(at).count();
        let position = match node.placement() {
            None | Some(Placement::End) => None,
            Some(Placement::Start) => Some(0),
            Some(Placement::Position(position)) => Some(position.get()),
            Some(_) => {
                return Err(error(
                    Code::UnsupportedCapability,
                    "SmartArt node insertion supports start, end and position placement",
                ));
            },
        };
        if position.is_some_and(|position| position > count) {
            return Err(error(
                Code::OutOfRange,
                "the node position is past the last sibling",
            ));
        }
        let mut value = None;
        for (key, scalar) in node.props().iter() {
            let (_, entry) = NODE_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            entry.prepare(Operation::Add, scalar, |setter, scalar| {
                match setter {
                    NodeSetter::Text => value = Some(scalar_text(scalar)),
                }
                Ok(())
            })?;
        }
        let value = value.ok_or_else(|| {
            error(Code::MissingProperty, "a SmartArt node needs its text").with_property("text")
        })?;
        let commit = self.change_smartart(
            editor,
            &smartart,
            Change::Insert {
                parent: at,
                position,
                text: value,
            },
        )?;
        let mut steps = at
            .map(|index| spec_path(&outline, index))
            .unwrap_or_default();
        steps.push(position.unwrap_or(count));
        Ok((commit, node_path(smartart.ordinal, &steps)))
    }

    /// Remove graphics (each with its parts) or one node (record 0335).
    pub(super) fn remove_smartarts(
        &self,
        editor: DocxEditor,
        nodes: &[Node<'_>],
    ) -> Result<Commit> {
        if nodes.iter().any(|node| node.kind == Kind::Node) {
            if nodes.len() != 1 {
                return Err(error(
                    Code::UnsupportedCapability,
                    "a SmartArt node removal targets one node",
                ));
            }
            let Value::DiagramNode(node) = &nodes[0].value else {
                return Err(error(
                    Code::InternalError,
                    "the node holds no SmartArt node",
                ));
            };
            return self.change_smartart(
                editor,
                &node.smartart,
                Change::Remove {
                    index: node.index as usize,
                },
            );
        }
        let mut relationships = Vec::with_capacity(nodes.len());
        for node in nodes {
            let Value::SmartArt(smartart) = &node.value else {
                return Err(error(Code::InternalError, "the node holds no SmartArt"));
            };
            relationships.push(smartart.diagram.data_relationship_id.clone());
        }
        relationships.sort();
        relationships.dedup();
        self.owner().check_source()?;
        let mut edit = WorkEdit::from(self.eager_editor(editor)?)?;
        let WorkEdit::Eager(eager) = &mut edit else {
            return Err(refusal("SmartArt needs the owning package editor"));
        };
        for relationship in &relationships {
            eager.remove_smartart(relationship).map_err(Error::from)?;
        }
        edit.commit(self)
    }
}
