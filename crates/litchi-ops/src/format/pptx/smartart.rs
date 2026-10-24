//! Typed SmartArt CRUD on slides (record 0284).
//!
//! - `get`/`query`: `/slide[N]/smartart[M]` reads its frame and its data
//!   model (`layout`, `layoutId`, `nodeCount`, the outline as `text`);
//!   `/slide[N]/smartart[M]/node[K]` (nested: `.../node[K]/node[J]`) reads
//!   one content node (`text`, `level`).
//! - `add(/slide[N], smartart, {layout, text, name, x, y, width, height})`
//!   authors a graphic from a text outline
//!   ([`litchi_drawingml::diagram::authoring`]); `add(.../smartart[M] or
//!   .../node[K], node, {text})` adds a node.
//! - `set` of a node's `text`; of a SmartArt's `layout` and `text` (the
//!   whole outline).
//! - `remove` of a node and its descendants; the SmartArt frame itself is
//!   the ordinary shape removal, which drops every diagram part.
//!
//! The cached drawing stays consistent with the data model:
//!
//! - A graphic Litchi authored (its data model and layout are exactly what
//!   [`Spec`] writes) is regenerated whole after every edit, drawing
//!   included, deterministically.
//! - A node text edit of any other graphic rewrites that node's text in the
//!   data model in place and in the drawing shapes that show exactly that
//!   node; when no shape maps unambiguously, the drawing is removed so
//!   consumers lay the graphic out again from the data model.
//! - Adding or removing nodes of any other graphic is refused
//!   (`unsupported_capability`): its layout's presentation points cannot be
//!   regenerated faithfully. Setting `layout` (or `text`) rebuilds it as a
//!   Litchi graphic from its outline, after which nodes can be added and
//!   removed.
//!
//! Source owners are promoted to an eager package for these edits, as for
//! charts, since the source-backed slide edit does not stage new parts.

use std::sync::Arc;

use litchi_core::Position;
use litchi_core::selector::{Path, Predicate, Step};
use litchi_drawingml::diagram::Limits;
use litchi_drawingml::diagram::authoring::{
    Layout, Node as SpecNode, Spec, colors_xml, quick_style_xml,
};
use litchi_drawingml::diagram::edit::{self as diagram_edit, Outline};
use litchi_pptx::opened::SmartArtParts;

use super::node::{DiagramItem, Item, TreeItem};
use super::spec::{Kind, NODE_WRITE_TABLE, NodeSetter, SMARTART_WRITE_TABLE, SmartArtSetter};
use super::structure::{bound, error};
use super::write::{Commit, Proof, native_snapshot_budget, opened_edit_options};
use super::{Editable, Nav, Pptx};
use crate::path::PathArg;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{LimitKind, SemanticLimits};
use crate::value::Scalar;
use crate::{Code, NewNode, Placement, Props, Result};

/// The default SmartArt frame: 1in from the top-left corner, 6in × 4in.
const DEFAULT_BOUNDS: (i64, i64, i64, i64) = (914_400, 914_400, 5_486_400, 3_657_600);

fn text(value: Option<Scalar<'_>>) -> String {
    match value {
        Some(Scalar::Text(value) | Scalar::Token(value)) => value.to_owned(),
        _ => String::new(),
    }
}

fn leaf_error(code: Code, error: impl std::fmt::Display) -> crate::Error {
    crate::Error::new(code, error.to_string()).with_format(litchi_core::FileFormat::Pptx)
}

fn drawing_code(error: &litchi_drawingml::Error, fallback: Code) -> Code {
    match error {
        litchi_drawingml::Error::Limit { .. } => Code::LimitExceeded,
        litchi_drawingml::Error::InvalidLimit { .. } => Code::InvalidArgument,
        litchi_drawingml::Error::PatchConflict => Code::InvalidInput,
        _ => fallback,
    }
}

/// The data-model outline a tree item's SmartArt frame names.
pub(super) fn outline(file: &Pptx, tree: &TreeItem) -> Result<Option<Arc<Outline>>> {
    let Some(slide) = tree.container.slide() else {
        return Ok(None);
    };
    let Some(payload) = litchi_pptx::shape::raw::payload(tree.handle.view())
        .map_err(|error| leaf_error(Code::CorruptFile, error))?
    else {
        return Ok(None);
    };
    let relationship_id = payload
        .relationship_id()
        .map_err(|error| leaf_error(Code::CorruptFile, error))?;
    let Some(xml) = file.slide_diagram_xml(slide as usize, &relationship_id)? else {
        return Ok(None);
    };
    Outline::read(&xml, &Limits::default())
        .map(|outline| Some(Arc::new(outline)))
        .map_err(|error| leaf_error(drawing_code(&error, Code::CorruptFile), error))
}

/// The outline text of a data model: one line per node, one tab per level.
pub(super) fn outline_text(outline: &Outline) -> String {
    Spec::from_outline(Layout::List, outline)
        .outline(&Limits::default())
        .unwrap_or_default()
}

/// The `layout` value of a data model: Litchi's token for a layout it
/// authors, else the last segment of the layout's unique identifier.
pub(super) fn layout_name(outline: &Outline) -> Option<String> {
    let id = outline.layout_type_id.as_deref()?;
    Some(match Layout::from_unique_id(id) {
        Some(layout) => layout.name().to_owned(),
        None => id.rsplit('/').next().unwrap_or(id).to_owned(),
    })
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
        .map_err(|error| leaf_error(Code::InvalidValue, error).with_property("text"))?;
    if spec.nodes().is_empty() {
        return Err(
            error(Code::InvalidValue, "the SmartArt outline has no nodes")
                .with_property("text")
                .with_suggestion("give one node per line, children indented by a tab"),
        );
    }
    Ok(spec)
}

/// The eager package of an editable owner, promoting a source owner.
pub(super) fn eager(
    source: &Editable,
    limits: litchi_opc::SourceLimits,
) -> Result<Arc<litchi_pptx::Package>> {
    Ok(match source {
        Editable::Eager(package) => Arc::clone(package),
        Editable::Source(package) => super::eager_package_from_source(package.to_bytes()?, limits)?,
    })
}

/// The frame's extents in EMU.
fn extents(tree: &TreeItem) -> (i64, i64) {
    let view = tree.handle.view();
    if let Some(bounds) = view.bounds() {
        return (bounds.width(), bounds.height());
    }
    view.transform()
        .ok()
        .flatten()
        .and_then(|transform| {
            let extent = transform.extent()?;
            Some((extent.width().as_emu(), extent.height().as_emu()))
        })
        .filter(|(width, height)| *width > 0 && *height > 0)
        .unwrap_or((DEFAULT_BOUNDS.2, DEFAULT_BOUNDS.3))
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

/// One structural or text change of a SmartArt.
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

/// The new parts after `change`, for a frame of `extents`.
fn apply(parts: &SmartArtParts, change: Change, extents: (i64, i64)) -> Result<SmartArtParts> {
    let limits = Limits::default();
    let corrupt =
        |error: litchi_drawingml::Error| leaf_error(drawing_code(&error, Code::CorruptFile), error);
    let outline = Outline::read(&parts.data, &limits).map_err(corrupt)?;
    let authored =
        Spec::recognize(&outline, &parts.data, &parts.layout, &limits).map_err(corrupt)?;
    let regenerate = |spec: &Spec, all: bool| -> Result<SmartArtParts> {
        spec.validate()
            .map_err(|error| leaf_error(drawing_code(&error, Code::InvalidValue), error))?;
        let relationship = outline.drawing_relationship.as_deref().unwrap_or("rId0");
        let invalid = |error: litchi_drawingml::Error| {
            leaf_error(drawing_code(&error, Code::InvalidValue), error)
        };
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
                spec.drawing_xml(extents.0, extents.1, &limits)
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
                let (last, parents) = path
                    .split_last()
                    .ok_or_else(|| error(Code::PathNotFound, "the SmartArt node is missing"))?;
                let list = siblings_at(&mut spec, parents)
                    .ok_or_else(|| error(Code::PathNotFound, "the SmartArt node is missing"))?;
                list[*last].text = text;
                return regenerate(&spec, false);
            }
            let id = &outline.entries[index].id;
            let invalid = |error: litchi_drawingml::Error| {
                leaf_error(drawing_code(&error, Code::InvalidValue), error)
            };
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
            let list = siblings_at(&mut spec, &path)
                .ok_or_else(|| error(Code::PathNotFound, "the SmartArt node is missing"))?;
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
            let (last, parents) = path
                .split_last()
                .ok_or_else(|| error(Code::PathNotFound, "the SmartArt node is missing"))?;
            let list = siblings_at(&mut spec, parents)
                .ok_or_else(|| error(Code::PathNotFound, "the SmartArt node is missing"))?;
            list.remove(*last);
            regenerate(&spec, false)
        },
    }
}

fn foreign_structure() -> crate::Error {
    error(
        Code::UnsupportedCapability,
        "nodes can be added to or removed from SmartArt Litchi authored only",
    )
    .with_suggestion(
        "set layout on the smartart to rebuild it as a Litchi graphic from its outline (formatting of the nodes is not kept), or edit node text only",
    )
}

/// The SmartArt frame, and the node a target names, of a resolved node.
fn target_of(item: &Item) -> Option<(TreeItem, Option<(Arc<Outline>, usize)>)> {
    match item {
        Item::Tree(tree) if tree.kind == Kind::SmartArt => Some((tree.clone(), None)),
        Item::DiagramNode(node) => Some((
            node.smartart.clone(),
            Some((Arc::clone(&node.outline), node.index as usize)),
        )),
        _ => None,
    }
}

impl Pptx {
    /// Stage `change` on the SmartArt `tree` in one native edit.
    fn change_smartart(
        &self,
        tree: &TreeItem,
        change: Change,
        limits: &SemanticLimits,
    ) -> Result<Commit> {
        let Some(slide) = tree.container.slide() else {
            return Err(error(
                Code::UnsupportedCapability,
                "master and layout SmartArt is not editable",
            ));
        };
        if tree.handle.parent().is_some() {
            return Err(error(
                Code::UnsupportedCapability,
                "SmartArt writes require a top-level slide SmartArt",
            ));
        }
        let position = tree.handle.position().get();
        let source = if matches!(&change, Change::Text { .. } | Change::Rebuild { .. }) {
            self.editable_owner_for_write()?
        } else {
            self.editable_owner()?
        };
        let base = eager(&source, self.0.ooxml_limits)?;
        let mut native = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        let parts = native.smartart(slide as usize, position)?;
        let next = apply(&parts, change, extents(tree))?;
        if next != parts {
            native.set_smartart(slide as usize, position, next)?;
        }
        let (target, patch) = native.commit()?.into_parts();
        let target = Arc::new(target);
        let after = if patch.is_noop() {
            self.clone()
        } else {
            self.adopt_eager(Arc::clone(&target))
        };
        Ok(Commit::new(
            after,
            source,
            Editable::Eager(target),
            Proof::Eager(patch),
            1,
        ))
    }

    /// `add(/slide[N], smartart, ...)`.
    pub(super) fn add_smartart(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let resolved = crate::path::resolve::get(&Nav::new(self), &parent)?;
        let Item::Slide(slide) = resolved.item() else {
            return Err(error(
                Code::UnsupportedType,
                "SmartArt requires a slide parent",
            ));
        };
        if !matches!(node.placement(), None | Some(Placement::End)) {
            return Err(error(
                Code::UnsupportedCapability,
                "SmartArt insertion currently supports end placement",
            )
            .with_suggestion("omit placement or use Placement::End"));
        }
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let mut layout = Layout::List;
        let mut outline = None;
        let mut name = None;
        let mut geometry = [None; 4];
        let mut assigned = 0_u32;
        let mut staged_bytes = 0_usize;
        for (key, value) in node.props().iter() {
            let (index, entry) = SMARTART_WRITE_TABLE
                .resolve(key.as_str())
                .map_err(|e| e.with_property(key.as_str().to_owned()))?;
            let bit = 1_u32 << index;
            if assigned & bit != 0 {
                return Err(error(
                    Code::InvalidArgument,
                    "a SmartArt property was assigned more than once",
                )
                .with_property(entry.spec.name));
            }
            assigned |= bit;
            entry.prepare(Operation::Add, value, |setter, scalar| {
                if let Some(Scalar::Text(value)) = scalar {
                    staged_bytes = staged_bytes.saturating_add(value.len());
                    bound(
                        LimitKind::StagedBytes,
                        staged_bytes,
                        limits.max_staged_bytes(),
                    )?;
                }
                match setter {
                    SmartArtSetter::X
                    | SmartArtSetter::Y
                    | SmartArtSetter::Width
                    | SmartArtSetter::Height => {
                        let Some(Scalar::Length(value)) = scalar else {
                            unreachable!("SmartArt geometry coercion")
                        };
                        let slot = match setter {
                            SmartArtSetter::X => 0,
                            SmartArtSetter::Y => 1,
                            SmartArtSetter::Width => 2,
                            _ => 3,
                        };
                        geometry[slot] = Some(value.emu());
                    },
                    SmartArtSetter::Name => {
                        let value = text(scalar);
                        super::write::validate_text(&value)?;
                        name = Some(value);
                    },
                    SmartArtSetter::Layout => layout = parse_layout(&text(scalar))?,
                    SmartArtSetter::Text => outline = Some(text(scalar)),
                }
                Ok(())
            })?;
        }
        let outline = outline.ok_or_else(|| {
            error(Code::MissingProperty, "a SmartArt needs its outline")
                .with_property("text")
                .with_suggestion("text=\"Plan\\nBuild\\n\\tTest\\nShip\"")
        })?;
        let spec = parse_outline(layout, &outline)?;
        let bounds = (
            geometry[0].unwrap_or(DEFAULT_BOUNDS.0),
            geometry[1].unwrap_or(DEFAULT_BOUNDS.1),
            geometry[2].unwrap_or(DEFAULT_BOUNDS.2),
            geometry[3].unwrap_or(DEFAULT_BOUNDS.3),
        );
        let slide = slide.position as usize;
        let scene = self
            .slide_scene(slide)?
            .ok_or_else(|| error(Code::PathNotFound, "slide parent is missing"))?;
        let ordinal = scene
            .roots()
            .filter(|shape| super::steps::tree_kind(shape) == Some(Kind::SmartArt))
            .count();
        let name = name.unwrap_or_else(|| format!("Diagram {}", ordinal + 1));
        let source = self.editable_owner()?;
        let base = eager(&source, self.0.ooxml_limits)?;
        let mut edit = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?;
        edit.add_smartart(slide, &spec, &name, bounds)?;
        let (target, patch) = edit.commit()?.into_parts();
        let target = Arc::new(target);
        let after = self.adopt_eager(Arc::clone(&target));
        let path = Path::from_steps([
            Step::new("slide").with(Predicate::Index(Position::new(slide))),
            Step::new("smartart").with(Predicate::Index(Position::new(ordinal))),
        ]);
        Ok((
            Commit::new(
                after,
                source,
                Editable::Eager(target),
                Proof::Eager(patch),
                1,
            ),
            path,
        ))
    }

    /// `add(.../smartart[M] or .../node[K], node, {text})`.
    pub(super) fn add_smartart_node(
        &self,
        parent: PathArg<'_>,
        node: NewNode<'_>,
        limits: &SemanticLimits,
    ) -> Result<(Commit, Path<'static>)> {
        let resolved = crate::path::resolve::get(&Nav::new(self), &parent)?;
        let Some((tree, at)) = target_of(resolved.item()) else {
            return Err(error(
                Code::UnsupportedType,
                "SmartArt nodes require a smartart or node parent",
            ));
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let outline = match &at {
            Some((outline, _)) => Arc::clone(outline),
            None => tree.diagram(self)?.ok_or_else(|| {
                error(Code::CorruptFile, "the SmartArt frame names no data model")
            })?,
        };
        let parent_index = at.as_ref().map(|(_, index)| *index);
        let count = outline.children(parent_index).count();
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
                    NodeSetter::Text => {
                        let scalar = text(scalar);
                        bound(
                            LimitKind::StagedBytes,
                            scalar.len(),
                            limits.max_staged_bytes(),
                        )?;
                        value = Some(scalar);
                    },
                }
                Ok(())
            })?;
        }
        let value = value.ok_or_else(|| {
            error(Code::MissingProperty, "a SmartArt node needs its text").with_property("text")
        })?;
        let commit = self.change_smartart(
            &tree,
            Change::Insert {
                parent: parent_index,
                position,
                text: value,
            },
            limits,
        )?;
        // The new node's path: the parent's, then its sibling position.
        let mut steps: Vec<Step<'static>> = Vec::new();
        let slide = tree.container.slide().unwrap_or(0) as usize;
        steps.push(Step::new("slide").with(Predicate::Index(Position::new(slide))));
        steps.push(
            Step::new("smartart").with(Predicate::Index(Position::new(tree.ordinal as usize))),
        );
        if let Some(index) = parent_index {
            for position in spec_path(&outline, index) {
                steps.push(Step::new("node").with(Predicate::Index(Position::new(position))));
            }
        }
        steps.push(
            Step::new("node").with(Predicate::Index(Position::new(position.unwrap_or(count)))),
        );
        Ok((commit, Path::from_steps(steps)))
    }

    /// A `set` on a SmartArt or SmartArt-node path; `None` for other targets.
    pub(super) fn try_set_smartart(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        let Some(last) = path.path().steps().last() else {
            return Ok(None);
        };
        if !["smartart", "node"]
            .iter()
            .any(|name| last.name().eq_ignore_ascii_case(name))
        {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let Some((tree, at)) = target_of(node.item()) else {
            return Ok(None);
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        let mut staged_bytes = 0_usize;
        let mut charge = |value: &Option<Scalar<'_>>| -> Result<()> {
            if let Some(Scalar::Text(value)) = value {
                staged_bytes = staged_bytes.saturating_add(value.len());
                bound(
                    LimitKind::StagedBytes,
                    staged_bytes,
                    limits.max_staged_bytes(),
                )?;
            }
            Ok(())
        };
        let change = match at {
            Some((_, index)) => {
                let mut value = None;
                for (key, scalar) in props.iter() {
                    let (_, entry) = NODE_WRITE_TABLE
                        .resolve(key.as_str())
                        .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                    entry.prepare(Operation::Set, scalar, |setter, scalar| {
                        charge(&scalar)?;
                        match setter {
                            NodeSetter::Text => value = Some(text(scalar)),
                        }
                        Ok(())
                    })?;
                }
                let Some(value) = value else {
                    return Ok(Some(self.unchanged(limits)?));
                };
                Change::Text { index, text: value }
            },
            None => {
                let mut layout = None;
                let mut outline = None;
                for (key, scalar) in props.iter() {
                    let (_, entry) = SMARTART_WRITE_TABLE
                        .resolve(key.as_str())
                        .map_err(|e| e.with_property(key.as_str().to_owned()))?;
                    entry.prepare(Operation::Set, scalar, |setter, scalar| {
                        charge(&scalar)?;
                        match setter {
                            SmartArtSetter::Layout => layout = Some(parse_layout(&text(scalar))?),
                            SmartArtSetter::Text => outline = Some(text(scalar)),
                            _ => unreachable!("SmartArt set admission"),
                        }
                        Ok(())
                    })?;
                }
                if layout.is_none() && outline.is_none() {
                    return Ok(Some(self.unchanged(limits)?));
                }
                Change::Rebuild { layout, outline }
            },
        };
        self.change_smartart(&tree, change, limits).map(Some)
    }

    /// A `remove` of one SmartArt-node path; `None` for other targets.
    pub(super) fn try_remove_smartart_node(
        &self,
        target: &TargetArg<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Commit>> {
        let Some(path) = target.path() else {
            return Ok(None);
        };
        if !path
            .path()
            .steps()
            .last()
            .is_some_and(|step| step.name().eq_ignore_ascii_case("node"))
        {
            return Ok(None);
        }
        let node = crate::path::resolve::get(&Nav::new(self), path)?;
        let Item::DiagramNode(DiagramItem {
            smartart, index, ..
        }) = node.item()
        else {
            return Ok(None);
        };
        bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
        self.change_smartart(
            smartart,
            Change::Remove {
                index: *index as usize,
            },
            limits,
        )
        .map(Some)
    }

    /// A commit that changes nothing.
    fn unchanged(&self, limits: &SemanticLimits) -> Result<Commit> {
        // A genuine signed no-op must still stage through the private
        // write-capable owner.  The outer set fence re-roots its empty proof
        // on the original signed owner; refusing here would reject a no-op.
        let source = self.editable_owner_for_write()?;
        let base = eager(&source, self.0.ooxml_limits)?;
        let (target, patch) = base
            .edit_with_options(opened_edit_options(native_snapshot_budget(limits)))
            .map_err(super::write::native_error)?
            .commit()
            .map_err(super::write::native_error)?
            .into_parts();
        Ok(Commit::new(
            self.clone(),
            source,
            Editable::Eager(Arc::new(target)),
            Proof::Eager(patch),
            1,
        ))
    }
}

/// The children a SmartArt or node lists: its child nodes.
pub(super) fn children(file: &Pptx, item: &Item) -> Result<Vec<Item>> {
    let (tree, outline, parent) = match item {
        Item::Tree(tree) if tree.kind == Kind::SmartArt => match tree.diagram(file)? {
            Some(outline) => (tree.clone(), outline, None),
            None => return Ok(Vec::new()),
        },
        Item::DiagramNode(node) => (
            node.smartart.clone(),
            Arc::clone(&node.outline),
            Some(node.index as usize),
        ),
        _ => return Ok(Vec::new()),
    };
    Ok(outline
        .children(parent)
        .map(|index| {
            Item::DiagramNode(DiagramItem {
                smartart: tree.clone(),
                outline: Arc::clone(&outline),
                index: u32::try_from(index).unwrap_or(u32::MAX),
            })
        })
        .collect())
}
