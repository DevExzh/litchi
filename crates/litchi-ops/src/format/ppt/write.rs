//! Unified native-backed writes for legacy PowerPoint presentations.
//!
//! The adapter resolves paths and scoped queries against the existing read
//! navigation, validates properties through the shared registry, and submits
//! the complete call to one native [`litchi_ppt::Package::edit`] transaction.
//! The native owner remains responsible for dependency checks and publication;
//! this module does not maintain a second presentation model or rewrite raw
//! streams itself.

pub(crate) mod coalesce;
mod elements;
pub(crate) mod find;

use std::collections::HashSet;
use std::mem::size_of;

use litchi_core::{FileFormat, Position};
use litchi_ppt::Anchor;
use litchi_ppt::text_edit::Target as NativeTarget;

use super::node::Node;
use super::spec::{
    Kind, NOTES_WRITE_TABLE, PARAGRAPH_WRITE_TABLE, PLACEHOLDER_WRITE_TABLE, SHAPE_WRITE_TABLE,
    SLIDE_WRITE_TABLE, ShapeSetter, SlideSetter, TextSetter,
};
use super::{Nav, Ppt};
use crate::error::{Code, Error, Result};
use crate::path::PathArg;
use crate::path::navigate::ElementKind as _;
use crate::placement::Placement;
use crate::props::{IntoProps, Props};
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

/// The native proof and immutable endpoints of one unified PPT write.
/// The patch a PPT write step retains: a primary edit or a metadata edit.
pub(crate) type StepPatch =
    crate::format::ole_metadata::StepPatch<litchi_ppt::Patch, litchi_ppt::package::metadata::Patch>;

pub(crate) struct WriteCommit {
    pub(crate) before: Ppt,
    pub(crate) after: Ppt,
    pub(crate) patch: StepPatch,
    pub(crate) operations: usize,
    pub(crate) snapshot_bytes: usize,
}

#[derive(Debug, Clone)]
struct ShapeValues {
    text: Option<String>,
    geometry: [Option<i64>; 4],
}

impl ShapeValues {
    fn has_geometry(&self) -> bool {
        self.geometry.iter().any(Option::is_some)
    }

    fn charge(&self) -> usize {
        self.text
            .as_ref()
            .map_or(0, String::len)
            .saturating_add(if self.has_geometry() {
                4 * size_of::<i64>()
            } else {
                0
            })
            .saturating_add(2 * size_of::<usize>())
    }
}

#[derive(Debug, Clone, Copy)]
struct SlideValues {
    hidden: bool,
}

#[derive(Debug, Clone)]
struct ShapePlan {
    target: NativeTarget,
    text: Option<String>,
    anchor: Option<Anchor>,
}

#[derive(Debug, Clone, Copy)]
struct SlidePlan {
    position: Position,
    hidden: bool,
}

#[derive(Debug, Clone)]
struct NotesPlan {
    position: Position,
    text: String,
}

#[derive(Debug, Clone)]
enum Plan {
    Shape(ShapePlan),
    Slide(SlidePlan),
    Notes(NotesPlan),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum TargetIdentity {
    Slide(usize),
    Shape {
        slide: usize,
        shape: usize,
    },
    Notes(usize),
    Comment {
        slide: usize,
        index: usize,
    },
    Cell {
        slide: usize,
        shape: usize,
        row: usize,
        column: usize,
    },
    Hyperlink {
        slide: usize,
        shape: usize,
        run: Option<(usize, usize)>,
    },
    Run {
        slide: usize,
        shape: usize,
        paragraph: usize,
        run: usize,
    },
    Paragraph {
        slide: usize,
        shape: usize,
        paragraph: usize,
    },
}

/// The kind of element one PPT write targets; a write's targets share one.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum TargetClass {
    Slide,
    Shape,
    Notes,
    Paragraph,
}

impl TargetClass {
    fn of(node: &Node<'_>) -> Self {
        match node.element() {
            Kind::Slide => Self::Slide,
            Kind::Notes => Self::Notes,
            Kind::Paragraph => Self::Paragraph,
            _ => Self::Shape,
        }
    }
}

/// The bounded staging charge for one selected slide position.  The native
/// edit retains the complete before/after artifacts separately; this charge
/// covers only the resolved structural target list.
const STRUCTURAL_TARGET_BYTES: usize = size_of::<usize>();

/// Set one PPT target with the snapshot budget left by an outer transaction.
/// The budget is threaded to the native editor preflight so a previous leaf
/// cannot consume the full per-transaction limit before this editor captures
/// its before image.
pub(crate) fn set_with_snapshot_budget<'p, 'v>(
    file: &Ppt,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
    snapshot_budget: usize,
) -> Result<WriteCommit> {
    file.admit_edit_for_write()?;
    let props = Props::from_props(props)?;
    if crate::format::ole_metadata::is_target(&target) {
        return file.set_metadata(&props, limits);
    }
    let nodes = resolve_nodes(file, &target, limits)?;
    if let Some(commit) = elements::set(file, &nodes, &props, limits) {
        return commit;
    }
    let plans = build_plans(nodes, &props, limits)?;
    let native = stage_plans(file, plans, limits, snapshot_budget)?;
    file.check_source_for_write()?;
    finish_commit(file, native, limits)
}

/// Remove one slide or every slide selected by a scoped query.
///
/// Resolution is completed against the immutable source before the native edit
/// is created.  Positions are then removed in descending order so each native
/// position still refers to the source catalog selected by the caller.  The
/// native transaction remains failure-atomic and owns all slide-list and
/// outline-record dependency handling.
pub(crate) fn remove_with_limits(
    file: &Ppt,
    target: TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.admit_edit_for_write()?;
    let nodes = resolve_remove_nodes(file, &target, limits)?;
    if nodes
        .first()
        .is_some_and(|node| node.element() == Kind::Comment)
    {
        if nodes.iter().any(|node| node.element() != Kind::Comment) {
            return Err(unsupported_target(
                "one PPT removal cannot mix comments with other targets",
            ));
        }
        return elements::remove_comments(file, &nodes, limits);
    }
    if nodes
        .first()
        .is_some_and(|node| node.element() == Kind::Hyperlink)
    {
        if nodes.iter().any(|node| node.element() != Kind::Hyperlink) {
            return Err(unsupported_target(
                "one PPT removal cannot mix hyperlinks with other targets",
            ));
        }
        return elements::remove_hyperlinks(file, &nodes, limits);
    }
    if nodes
        .iter()
        .any(|node| matches!(node.element(), Kind::Row | Kind::Col | Kind::Cell))
    {
        return Err(elements::table_layout_refusal("remove"));
    }
    let mut positions = slide_targets(nodes, limits)?;
    admit_structural_targets(positions.len(), limits)?;
    positions.sort_unstable_by(|left, right| right.cmp(left));
    let native = stage_structural(file, positions.len(), limits, move |edit| {
        for position in positions {
            edit.remove_slide(Position::new(position))?;
        }
        Ok(())
    })?;
    file.check_source_for_write()?;
    finish_commit(file, native, limits)
}

/// Insert one blank slide under the presentation root.
///
/// The only node kind is `slide`; its one optional property is
/// `layout = "blank"`, the layout the binary format can create without
/// authoring placeholder shapes and slide-list text. Placement follows the
/// common contract (end by default; start, a final position, before or after
/// a slide, or into the root). The native transaction allocates the slide's
/// persist, slide, drawing and shape identifiers.
pub(crate) fn add_with_limits(
    file: &Ppt,
    parent: PathArg<'_>,
    node: crate::NewNode<'_>,
    limits: &SemanticLimits,
) -> Result<(WriteCommit, crate::path::Path<'static>)> {
    file.admit_edit_for_write()?;
    let parent = crate::path::resolve::get(&Nav::new(file), &parent)?;
    if node.kind().eq_ignore_ascii_case("comment") {
        return elements::add_comment(file, &parent, &node, limits);
    }
    if matches!(
        Kind::from_name(&node.kind().to_ascii_lowercase()),
        Some(Kind::Hyperlink)
    ) {
        return elements::add_hyperlink(file, &parent, &node, limits);
    }
    if matches!(
        Kind::from_name(&node.kind().to_ascii_lowercase()),
        Some(Kind::Row | Kind::Col | Kind::Cell)
    ) {
        return Err(elements::table_layout_refusal("add"));
    }
    if parent.element() != Kind::Presentation {
        return Err(Error::new(
            Code::UnsupportedType,
            "PPT slides can only be added under the presentation root",
        )
        .with_format(FileFormat::Ppt));
    }
    if !node.kind().eq_ignore_ascii_case("slide") {
        return Err(Error::new(
            Code::UnsupportedElement,
            "the PPT add verb supports slide nodes only",
        )
        .with_format(FileFormat::Ppt)
        .with_suggestion("add a slide under \"/\""));
    }
    for (key, value) in node.props().iter() {
        if !key.as_str().eq_ignore_ascii_case("layout") {
            return Err(Error::new(
                Code::UnsupportedProperty,
                "a new PPT slide accepts only the `layout` property",
            )
            .with_format(FileFormat::Ppt)
            .with_property(key.as_str().to_owned()));
        }
        if !value
            .as_text()
            .is_some_and(|layout| layout.eq_ignore_ascii_case("blank"))
        {
            return Err(Error::new(
                Code::UnsupportedCapability,
                "PPT slide insertion creates blank-layout slides only",
            )
            .with_format(FileFormat::Ppt)
            .with_property("layout")
            .with_suggestion("use layout \"blank\" and add content with later edits"));
        }
    }
    let count = file.slide_count()?;
    let position = match node.placement() {
        None | Some(Placement::End) => count,
        Some(Placement::Start) => 0,
        Some(Placement::Position(position)) => position.get(),
        Some(Placement::Before(path)) => resolve_slide_path(file, path)?,
        Some(Placement::After(path)) => resolve_slide_path(file, path)?
            .checked_add(1)
            .ok_or_else(|| out_of_range("PPT slide position overflow"))?,
        Some(Placement::Into(path, position)) => {
            let parent = crate::path::resolve::get(&Nav::new(file), path)?;
            if parent.element() != Kind::Presentation {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "PPT slides can only be added within the presentation root",
                )
                .with_format(FileFormat::Ppt));
            }
            position.map_or(count, Position::get)
        },
    };
    if position > count {
        return Err(out_of_range(
            "the PPT slide insertion position is outside the presentation",
        ));
    }
    admit_structural_targets(1, limits)?;
    let native = stage_structural(file, 1, limits, move |edit| {
        edit.insert_slide(Position::new(position))
    })?;
    file.check_source_for_write()?;
    let commit = finish_commit(file, native, limits)?;
    Ok((
        commit,
        crate::path::Path::from_steps([crate::path::Step::new("slide")
            .with(crate::path::Predicate::Index(Position::new(position)))]),
    ))
}

/// Move one slide to a final sibling position or relative to another slide.
///
/// Positional destinations are interpreted after detaching the moving slide,
/// matching the common placement contract.  Relative destinations are adjusted
/// against the source position before the native transaction is staged.
pub(crate) fn move_to_with_limits<'p, 'q>(
    file: &Ppt,
    path: PathArg<'p>,
    placement: Placement<'q>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.admit_edit_for_write()?;
    let from = resolve_slide_path(file, &path)?;
    let count = file.slide_count()?;
    let destination = move_destination(file, from, count, &placement)?;
    admit_structural_targets(1, limits)?;
    let native = stage_structural(file, 1, limits, move |edit| {
        if from != destination {
            edit.move_slide(Position::new(from), Position::new(destination))?;
        }
        Ok(())
    })?;
    file.check_source_for_write()?;
    finish_commit(file, native, limits)
}

/// Exchange two slides in one native edit.
///
/// The legacy native owner exposes slide movement rather than a dedicated
/// exchange operation.  A non-adjacent exchange therefore stages two moves;
/// adjacent slides need only the first move.  Both positions are resolved
/// before staging, and selecting the same slide is an exact no-op.
pub(crate) fn swap_with_limits<'p, 'q>(
    file: &Ppt,
    left: PathArg<'p>,
    right: PathArg<'q>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.admit_edit_for_write()?;
    let left = resolve_slide_path(file, &left)?;
    let right = resolve_slide_path(file, &right)?;
    let operations = if left == right {
        0
    } else if left.abs_diff(right) == 1 {
        1
    } else {
        2
    };
    admit_structural_targets(usize::from(left != right).saturating_add(1), limits)?;
    let native = stage_structural(file, operations, limits, move |edit| {
        if left == right {
            return Ok(());
        }
        edit.move_slide(Position::new(left), Position::new(right))?;
        if left < right && left + 1 != right {
            edit.move_slide(Position::new(right - 1), Position::new(left))?;
        } else if left > right && right + 1 != left {
            edit.move_slide(Position::new(right + 1), Position::new(left))?;
        }
        Ok(())
    })?;
    file.check_source_for_write()?;
    finish_commit(file, native, limits)
}

fn resolve_nodes<'f>(
    file: &'f Ppt,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<Node<'f>>> {
    let mut nodes = Vec::new();
    let mut seen = HashSet::new();
    let nav = Nav::new(file);
    if let Some(path) = target.path() {
        append_node(
            &mut nodes,
            &mut seen,
            crate::path::resolve::get(&nav, path)?,
            limits,
        )?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(nav, query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(invalid_query());
        }
        for node in walk {
            append_node(&mut nodes, &mut seen, node?, limits)?;
        }
    } else {
        return Err(empty_target());
    }
    file.check_source_for_write()?;
    if nodes.is_empty() {
        return Err(
            Error::new(Code::ZeroMatches, "mutation selector matched no element")
                .with_format(FileFormat::Ppt),
        );
    }
    Ok(nodes)
}

/// Resolve a removal target to its nodes, in selection order. Duplicates
/// are kept; the kind-specific removal deduplicates them.
fn resolve_remove_nodes<'f>(
    file: &'f Ppt,
    target: &TargetArg<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<Node<'f>>> {
    let mut nodes = Vec::new();
    let nav = Nav::new(file);
    let mut push = |node: Node<'f>| -> Result<()> {
        let count = nodes
            .len()
            .checked_add(1)
            .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
        // A selection may name one element several times; bound the raw
        // selection at twice the expansion limit, then deduplicate.
        if count > limits.max_expanded_targets().saturating_mul(2) {
            return Err(expanded_limit(count, limits.max_expanded_targets()));
        }
        nodes
            .try_reserve(1)
            .map_err(|source| allocation(source, "PPT removal targets"))?;
        nodes.push(node);
        Ok(())
    };
    if let Some(path) = target.path() {
        push(crate::path::resolve::get(&nav, path)?)?;
    } else if let Some(query) = target.query() {
        let walk = crate::query::run(nav, query)?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(invalid_query());
        }
        for node in walk {
            push(node?)?;
        }
    } else {
        return Err(empty_target());
    }
    file.check_source_for_write()?;
    if nodes.is_empty() {
        return Err(
            Error::new(Code::ZeroMatches, "mutation selector matched no slide")
                .with_format(FileFormat::Ppt),
        );
    }
    if nodes
        .first()
        .is_some_and(|node| node.element() == Kind::Comment)
    {
        let mut seen = HashSet::new();
        nodes.retain(|node| match node.element_item() {
            Some(super::elements::Element::Comment { slide, index, .. }) => {
                seen.insert((*slide, *index))
            },
            _ => true,
        });
        if nodes.len() > limits.max_expanded_targets() {
            return Err(expanded_limit(nodes.len(), limits.max_expanded_targets()));
        }
    }
    Ok(nodes)
}

/// The distinct slide positions of resolved removal nodes.
fn slide_targets(nodes: Vec<Node<'_>>, limits: &SemanticLimits) -> Result<Vec<usize>> {
    let mut positions = Vec::new();
    let mut seen = HashSet::new();
    for node in nodes {
        append_slide_target(&mut positions, &mut seen, node, limits)?;
    }
    Ok(positions)
}

fn append_slide_target(
    positions: &mut Vec<usize>,
    seen: &mut HashSet<usize>,
    node: Node<'_>,
    limits: &SemanticLimits,
) -> Result<()> {
    let position = resolve_slide_node(&node)?;
    if seen.contains(&position) {
        return Ok(());
    }
    let count = positions
        .len()
        .checked_add(1)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if count > limits.max_expanded_targets() {
        return Err(expanded_limit(count, limits.max_expanded_targets()));
    }
    positions
        .try_reserve(1)
        .map_err(|source| allocation(source, "PPT structural targets"))?;
    seen.try_reserve(1)
        .map_err(|source| allocation(source, "PPT structural target identities"))?;
    seen.insert(position);
    positions.push(position);
    Ok(())
}

fn resolve_slide_path(file: &Ppt, path: &PathArg<'_>) -> Result<usize> {
    resolve_slide_node(&crate::path::resolve::get(&Nav::new(file), path)?)
}

fn resolve_slide_node(node: &Node<'_>) -> Result<usize> {
    if node.element() != Kind::Slide {
        return Err(Error::new(
            Code::UnsupportedType,
            "PPT structural ordering requires slide targets",
        )
        .with_format(FileFormat::Ppt));
    }
    node.slide_position().ok_or_else(|| {
        Error::new(Code::InternalError, "PPT slide target has no position")
            .with_format(FileFormat::Ppt)
    })
}

fn move_destination(
    file: &Ppt,
    from: usize,
    count: usize,
    placement: &Placement<'_>,
) -> Result<usize> {
    if count == 0 {
        return Err(
            Error::new(Code::NotFound, "the PPT presentation has no slides")
                .with_format(FileFormat::Ppt),
        );
    }
    let last = count - 1;
    let destination = match placement {
        Placement::Start => 0,
        Placement::End => last,
        Placement::Position(position) => position.get(),
        Placement::Before(path) => {
            let anchor = resolve_slide_path(file, path)?;
            if from == anchor {
                from
            } else if from < anchor {
                anchor - 1
            } else {
                anchor
            }
        },
        Placement::After(path) => {
            let anchor = resolve_slide_path(file, path)?;
            if from == anchor {
                from
            } else if from < anchor {
                anchor
            } else {
                anchor
                    .checked_add(1)
                    .ok_or_else(|| out_of_range("PPT slide position overflow"))?
            }
        },
        Placement::Into(path, position) => {
            let parent = crate::path::resolve::get(&Nav::new(file), path)?;
            if parent.element() != Kind::Presentation {
                return Err(Error::new(
                    Code::UnsupportedType,
                    "PPT slides can only be reordered within the presentation root",
                )
                .with_format(FileFormat::Ppt));
            }
            position.map_or(last, Position::get)
        },
    };
    if destination >= count {
        return Err(out_of_range(
            "the PPT slide move position is outside the presentation",
        ));
    }
    Ok(destination)
}

fn admit_structural_targets(count: usize, limits: &SemanticLimits) -> Result<()> {
    bound(
        LimitKind::ExpandedTargets,
        count,
        limits.max_expanded_targets(),
    )?;
    let staged = count.saturating_mul(STRUCTURAL_TARGET_BYTES);
    bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())
}

fn append_node<'f>(
    nodes: &mut Vec<Node<'f>>,
    seen: &mut HashSet<TargetIdentity>,
    node: Node<'f>,
    limits: &SemanticLimits,
) -> Result<()> {
    let identity = match node.element() {
        Kind::Slide => TargetIdentity::Slide(
            node.slide_position()
                .ok_or_else(|| unsupported_target("slide target has no position"))?,
        ),
        Kind::Notes => TargetIdentity::Notes(
            node.slide_position()
                .ok_or_else(|| unsupported_target("notes target has no slide"))?,
        ),
        Kind::Comment => match node.element_item() {
            Some(super::elements::Element::Comment { slide, index, .. }) => {
                TargetIdentity::Comment {
                    slide: *slide,
                    index: *index,
                }
            },
            _ => return Err(unsupported_target("comment target has no comment")),
        },
        Kind::Hyperlink => match (node.native_shape_target(), node.element_item()) {
            (Some(target), Some(super::elements::Element::Hyperlink { run, .. })) => {
                TargetIdentity::Hyperlink {
                    slide: target.slide().get(),
                    shape: target.shape().get(),
                    run: *run,
                }
            },
            _ => {
                return Err(unsupported_target(
                    "nested PPT group hyperlink writes are not integrated",
                ));
            },
        },
        Kind::Run => match (node.native_shape_target(), node.element_item()) {
            (Some(target), Some(super::elements::Element::Run { paragraph, run, .. })) => {
                TargetIdentity::Run {
                    slide: target.slide().get(),
                    shape: target.shape().get(),
                    paragraph: *paragraph,
                    run: *run,
                }
            },
            _ => {
                return Err(unsupported_target(
                    "nested PPT group run writes are not integrated",
                ));
            },
        },
        Kind::Cell => match (node.native_shape_target(), node.element_item()) {
            (Some(target), Some(super::elements::Element::Cell { row, column, .. })) => {
                TargetIdentity::Cell {
                    slide: target.slide().get(),
                    shape: target.shape().get(),
                    row: *row,
                    column: *column,
                }
            },
            _ => {
                return Err(unsupported_target(
                    "nested PPT group table writes are not integrated",
                ));
            },
        },
        Kind::Paragraph => {
            let (Some(target), Some(paragraph)) =
                (node.native_shape_target(), node.paragraph_index())
            else {
                return Err(unsupported_target(
                    "nested PPT group shape writes are not integrated",
                ));
            };
            TargetIdentity::Paragraph {
                slide: target.slide().get(),
                shape: target.shape().get(),
                paragraph,
            }
        },
        _ => {
            let Some(target) = node.native_shape_target() else {
                return Err(unsupported_target(
                    "nested PPT group shape writes are not integrated",
                ));
            };
            TargetIdentity::Shape {
                slide: target.slide().get(),
                shape: target.shape().get(),
            }
        },
    };
    if seen.contains(&identity) {
        return Ok(());
    }
    let count = nodes
        .len()
        .checked_add(1)
        .ok_or_else(|| expanded_limit(usize::MAX, limits.max_expanded_targets()))?;
    if count > limits.max_expanded_targets() {
        return Err(expanded_limit(count, limits.max_expanded_targets()));
    }
    nodes
        .try_reserve(1)
        .map_err(|source| allocation(source, "PPT mutation targets"))?;
    seen.try_reserve(1)
        .map_err(|source| allocation(source, "PPT mutation target identities"))?;
    seen.insert(identity);
    nodes.push(node);
    Ok(())
}

fn build_plans(
    nodes: Vec<Node<'_>>,
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<Plan>> {
    let class = nodes.first().map_or(TargetClass::Shape, TargetClass::of);
    if nodes.iter().any(|node| TargetClass::of(node) != class) {
        return Err(unsupported_target(
            "one PPT write cannot mix slide, shape, notes and paragraph targets",
        ));
    }
    match class {
        TargetClass::Notes => return build_notes_plans(&nodes, props, limits),
        TargetClass::Paragraph => return build_paragraph_plans(&nodes, props, limits),
        TargetClass::Slide | TargetClass::Shape => {},
    }
    let shape = class == TargetClass::Shape;
    if shape {
        let mut kinds = Vec::new();
        for node in &nodes {
            let kind = node.property_kind();
            if !matches!(kind, Kind::Shape | Kind::Placeholder) {
                return Err(unsupported_target(
                    "PPT writes support only shape and placeholder elements",
                ));
            }
            if !kinds.contains(&kind) {
                kinds.push(kind);
            }
        }
        if props.is_empty() {
            return Ok(Vec::new());
        }
        let first_kind = kinds[0];
        let values = admit_shape_props(props, first_kind, limits)?;
        // A query can intentionally select both physical shapes and semantic
        // placeholders. Validate the same payload against every selected
        // registry before allocating native plans.
        for kind in kinds.into_iter().skip(1) {
            admit_shape_props(props, kind, limits)?;
        }
        let per_target = values.charge();
        let staged = nodes
            .len()
            .checked_mul(per_target)
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        if staged > limits.max_staged_bytes() {
            return Err(staged_limit(staged, limits.max_staged_bytes()));
        }
        let mut plans = Vec::new();
        plans
            .try_reserve_exact(nodes.len())
            .map_err(|source| allocation(source, "PPT shape plans"))?;
        for node in nodes {
            let target = node.native_shape_target().ok_or_else(|| {
                unsupported_target("nested PPT group shape writes are not integrated")
            })?;
            let anchor = if values.has_geometry() {
                build_anchor(&node, &values.geometry)?
            } else {
                None
            };
            plans.push(Plan::Shape(ShapePlan {
                target,
                text: values.text.clone(),
                anchor,
            }));
        }
        plans.sort_by_key(|plan| match plan {
            Plan::Shape(plan) => (plan.target.slide().get(), plan.target.shape().get()),
            Plan::Slide(_) | Plan::Notes(_) => (usize::MAX, usize::MAX),
        });
        Ok(plans)
    } else {
        if nodes
            .iter()
            .any(|node| !matches!(node.element(), Kind::Slide))
        {
            return Err(unsupported_target(
                "one PPT write cannot mix slide and shape targets",
            ));
        }
        if props.is_empty() {
            return Ok(Vec::new());
        }
        let values = admit_slide_props(props, limits)?;
        let per_target = size_of::<SlideValues>() + 2 * size_of::<usize>();
        let staged = nodes
            .len()
            .checked_mul(per_target)
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        if staged > limits.max_staged_bytes() {
            return Err(staged_limit(staged, limits.max_staged_bytes()));
        }
        let mut plans = Vec::new();
        plans
            .try_reserve_exact(nodes.len())
            .map_err(|source| allocation(source, "PPT slide plans"))?;
        for node in nodes {
            plans.push(Plan::Slide(SlidePlan {
                position: node
                    .slide_position()
                    .map(Position::new)
                    .ok_or_else(|| unsupported_target("slide target has no position"))?,
                hidden: values.hidden,
            }));
        }
        plans.sort_by_key(|plan| match plan {
            Plan::Slide(plan) => plan.position.get(),
            Plan::Shape(_) | Plan::Notes(_) => usize::MAX,
        });
        Ok(plans)
    }
}

/// Admit the one `text` property of a notes or paragraph write.
pub(super) fn admit_text_prop<const N: usize, const A: usize>(
    props: &Props<'_>,
    table: &crate::registry::Table<TextSetter, N, A>,
    limits: &SemanticLimits,
) -> Result<String> {
    let mut text = None;
    for (key, value) in props.iter() {
        let (_, entry) = table.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Ppt)
                .with_property(key.as_str().to_owned())
        })?;
        if text.is_some() {
            return Err(Error::new(
                Code::InvalidArgument,
                "a PPT text property was assigned more than once",
            )
            .with_format(FileFormat::Ppt)
            .with_property(entry.spec.name));
        }
        entry
            .prepare(Operation::Set, value, |setter, value| match setter {
                TextSetter::Text => {
                    let value = match value {
                        Some(Scalar::Text(value)) => value,
                        None => "",
                        _ => unreachable!("PPT text coercion returned a non-text scalar"),
                    };
                    bound(
                        LimitKind::StagedBytes,
                        value.len(),
                        limits.max_staged_bytes(),
                    )?;
                    let mut owned = String::new();
                    owned
                        .try_reserve_exact(value.len())
                        .map_err(|source| allocation(source, "PPT text"))?;
                    owned.push_str(value);
                    text = Some(owned);
                    Ok(())
                },
            })
            .map_err(|error| error.with_format(FileFormat::Ppt))?;
    }
    text.ok_or_else(|| {
        Error::new(
            Code::InvalidArgument,
            "the PPT text write requires the `text` property",
        )
        .with_format(FileFormat::Ppt)
    })
}

/// One native notes replacement per selected slide.
fn build_notes_plans(
    nodes: &[Node<'_>],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<Plan>> {
    if props.is_empty() {
        return Ok(Vec::new());
    }
    let text = admit_text_prop(props, &NOTES_WRITE_TABLE, limits)?;
    let staged = nodes
        .len()
        .checked_mul(text.len().saturating_add(2 * size_of::<usize>()))
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())?;
    let mut plans = Vec::new();
    plans
        .try_reserve_exact(nodes.len())
        .map_err(|source| allocation(source, "PPT notes plans"))?;
    for node in nodes {
        let position = node
            .slide_position()
            .ok_or_else(|| unsupported_target("notes target has no slide"))?;
        plans.push(Plan::Notes(NotesPlan {
            position: Position::new(position),
            text: text.clone(),
        }));
    }
    plans.sort_by_key(|plan| match plan {
        Plan::Notes(plan) => plan.position.get(),
        Plan::Shape(_) | Plan::Slide(_) => usize::MAX,
    });
    Ok(plans)
}

/// One native shape-text replacement per shape that holds selected
/// paragraphs: the shape's text with each selected paragraph replaced.
///
/// Paragraphs are the text between the binary format's paragraph marks
/// (`\r`). Replacing one therefore rewrites its shape's whole text body under
/// the native closure rules, which keep the formatting of a single-run body
/// and refuse a changed length where per-character ranges would shift.
fn build_paragraph_plans(
    nodes: &[Node<'_>],
    props: &Props<'_>,
    limits: &SemanticLimits,
) -> Result<Vec<Plan>> {
    if props.is_empty() {
        return Ok(Vec::new());
    }
    let text = admit_text_prop(props, &PARAGRAPH_WRITE_TABLE, limits)?;
    let mut by_shape =
        std::collections::BTreeMap::<(usize, usize), (NativeTarget, Vec<String>)>::new();
    let mut staged = 0usize;
    for node in nodes {
        let (Some(target), Some(index), Some(shape)) = (
            node.native_shape_target(),
            node.paragraph_index(),
            node.shape(),
        ) else {
            return Err(unsupported_target(
                "nested PPT group shape writes are not integrated",
            ));
        };
        let key = (target.slide().get(), target.shape().get());
        let entry = match by_shape.entry(key) {
            std::collections::btree_map::Entry::Occupied(entry) => entry.into_mut(),
            std::collections::btree_map::Entry::Vacant(entry) => {
                let current = shape.text().map_err(Error::from)?;
                staged = staged
                    .saturating_add(current.len())
                    .saturating_add(text.len());
                bound(LimitKind::StagedBytes, staged, limits.max_staged_bytes())?;
                let paragraphs = super::node::paragraphs(&current)
                    .map(str::to_owned)
                    .collect::<Vec<_>>();
                entry.insert((target, paragraphs))
            },
        };
        let paragraph = entry.1.get_mut(index).ok_or_else(|| {
            Error::new(Code::InternalError, "PPT paragraph vanished from its shape")
                .with_format(FileFormat::Ppt)
        })?;
        paragraph.clear();
        paragraph.push_str(&text);
    }
    let mut plans = Vec::new();
    plans
        .try_reserve_exact(by_shape.len())
        .map_err(|source| allocation(source, "PPT paragraph plans"))?;
    for (target, paragraphs) in by_shape.into_values() {
        plans.push(Plan::Shape(ShapePlan {
            target,
            text: Some(paragraphs.join("\r")),
            anchor: None,
        }));
    }
    Ok(plans)
}

fn admit_shape_props(
    props: &Props<'_>,
    kind: Kind,
    limits: &SemanticLimits,
) -> Result<ShapeValues> {
    match kind {
        Kind::Placeholder => admit_shape_props_from_table(props, &PLACEHOLDER_WRITE_TABLE, limits),
        Kind::Shape => admit_shape_props_from_table(props, &SHAPE_WRITE_TABLE, limits),
        _ => unreachable!("shape registry admission received an unsupported kind"),
    }
}

fn admit_shape_props_from_table<const N: usize, const A: usize>(
    props: &Props<'_>,
    table: &crate::registry::Table<ShapeSetter, N, A>,
    limits: &SemanticLimits,
) -> Result<ShapeValues> {
    let mut values = ShapeValues {
        text: None,
        geometry: [None; 4],
    };
    let mut seen_properties = 0_u128;
    for (key, value) in props.iter() {
        let (index, entry) = table.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Ppt)
                .with_property(key.as_str().to_owned())
        })?;
        let bit = 1_u128.checked_shl(index as u32).ok_or_else(|| {
            Error::new(
                Code::InternalError,
                "PPT shape property registry has too many entries",
            )
            .with_format(FileFormat::Ppt)
        })?;
        if seen_properties & bit != 0 {
            return Err(Error::new(
                Code::InvalidArgument,
                "a PPT shape property was assigned more than once",
            )
            .with_format(FileFormat::Ppt)
            .with_property(entry.spec.name));
        }
        seen_properties |= bit;
        entry
            .prepare(Operation::Set, value, |setter, value| match setter {
                ShapeSetter::Text => {
                    let text = match value {
                        Some(Scalar::Text(value)) => value,
                        None => "",
                        _ => unreachable!("PPT text coercion returned a non-text scalar"),
                    };
                    bound(
                        LimitKind::StagedBytes,
                        text.len(),
                        limits.max_staged_bytes(),
                    )?;
                    let mut owned = String::new();
                    owned
                        .try_reserve_exact(text.len())
                        .map_err(|source| allocation(source, "PPT shape text"))?;
                    owned.push_str(text);
                    values.text = Some(owned);
                    Ok(())
                },
                ShapeSetter::X | ShapeSetter::Y | ShapeSetter::Width | ShapeSetter::Height => {
                    let Some(Scalar::Length(value)) = value else {
                        unreachable!("PPT geometry coercion returned a non-length scalar");
                    };
                    let index = match setter {
                        ShapeSetter::X => 0,
                        ShapeSetter::Y => 1,
                        ShapeSetter::Width => 2,
                        ShapeSetter::Height => 3,
                        ShapeSetter::Text => unreachable!("PPT text setter in geometry branch"),
                    };
                    values.geometry[index] = Some(value.emu());
                    Ok(())
                },
            })
            .map_err(|error| error.with_format(FileFormat::Ppt))?;
    }
    if values.text.is_none() && !values.has_geometry() {
        return Err(Error::new(
            Code::InvalidArgument,
            "the PPT shape write has no writable property",
        )
        .with_format(FileFormat::Ppt));
    }
    Ok(values)
}

fn admit_slide_props(props: &Props<'_>, _limits: &SemanticLimits) -> Result<SlideValues> {
    let mut hidden = None;
    let mut seen_properties = 0_u128;
    for (key, value) in props.iter() {
        let (index, entry) = SLIDE_WRITE_TABLE.resolve(key.as_str()).map_err(|error| {
            error
                .with_format(FileFormat::Ppt)
                .with_property(key.as_str().to_owned())
        })?;
        let bit = 1_u128.checked_shl(index as u32).ok_or_else(|| {
            Error::new(
                Code::InternalError,
                "PPT slide property registry has too many entries",
            )
            .with_format(FileFormat::Ppt)
        })?;
        if seen_properties & bit != 0 {
            return Err(Error::new(
                Code::InvalidArgument,
                "a PPT slide property was assigned more than once",
            )
            .with_format(FileFormat::Ppt)
            .with_property(entry.spec.name));
        }
        seen_properties |= bit;
        entry
            .prepare(Operation::Set, value, |setter, value| match setter {
                SlideSetter::Hidden => {
                    let Some(Scalar::Bool(value)) = value else {
                        unreachable!("PPT hidden coercion returned a non-boolean scalar");
                    };
                    hidden = Some(value);
                    Ok(())
                },
            })
            .map_err(|error| error.with_format(FileFormat::Ppt))?;
    }
    hidden.map(|hidden| SlideValues { hidden }).ok_or_else(|| {
        Error::new(
            Code::InvalidArgument,
            "the PPT slide write requires the `hidden` property",
        )
        .with_format(FileFormat::Ppt)
    })
}

fn build_anchor(node: &Node<'_>, geometry: &[Option<i64>; 4]) -> Result<Option<Anchor>> {
    let (current_left, current_top, current_width, current_height) =
        node.anchor_bounds().ok_or_else(|| {
            Error::new(
                Code::UnsupportedCapability,
                "PPT geometry requires a shape with one editable client anchor",
            )
            .with_format(FileFormat::Ppt)
        })?;
    let left = geometry[0]
        .map(|value| master_from_emu(value, "x"))
        .transpose()?
        .unwrap_or(i64::from(current_left));
    let top = geometry[1]
        .map(|value| master_from_emu(value, "y"))
        .transpose()?
        .unwrap_or(i64::from(current_top));
    let width = geometry[2]
        .map(|value| master_from_emu(value, "width"))
        .transpose()?
        .unwrap_or(i64::from(current_width));
    let height = geometry[3]
        .map(|value| master_from_emu(value, "height"))
        .transpose()?
        .unwrap_or(i64::from(current_height));
    let right = left.checked_add(width).ok_or_else(|| {
        invalid_geometry("PPT shape width overflows the native anchor coordinate")
            .with_property("width")
    })?;
    let bottom = top.checked_add(height).ok_or_else(|| {
        invalid_geometry("PPT shape height overflows the native anchor coordinate")
            .with_property("height")
    })?;
    if left == i64::from(current_left)
        && top == i64::from(current_top)
        && width == i64::from(current_width)
        && height == i64::from(current_height)
    {
        // The unified Length may quantize back to the current native master
        // coordinate. Skipping the native call preserves exact bytes,
        // including the original Small versus Full anchor encoding.
        return Ok(None);
    }
    let left = i32::try_from(left)
        .map_err(|_| invalid_geometry("PPT x is outside native range").with_property("x"))?;
    let top = i32::try_from(top)
        .map_err(|_| invalid_geometry("PPT y is outside native range").with_property("y"))?;
    let right = i32::try_from(right).map_err(|_| {
        invalid_geometry("PPT width is outside native range").with_property("width")
    })?;
    let bottom = i32::try_from(bottom).map_err(|_| {
        invalid_geometry("PPT height is outside native range").with_property("height")
    })?;

    if let (Ok(left), Ok(top), Ok(right), Ok(bottom)) = (
        i16::try_from(left),
        i16::try_from(top),
        i16::try_from(right),
        i16::try_from(bottom),
    ) {
        return Anchor::small(left, top, right, bottom)
            .map(Some)
            .map_err(native_anchor_error);
    }
    Anchor::full(left, top, right, bottom)
        .map(Some)
        .map_err(native_anchor_error)
}

pub(super) fn master_from_emu(value: i64, property: &'static str) -> Result<i64> {
    // 576 master units per inch and 914,400 EMUs per inch. Round to the
    // nearest native unit, with exact half-way values away from zero (the
    // same policy as Rust's f64::round, without an intermediate float).
    let numerator = i128::from(value)
        .checked_mul(576)
        .ok_or_else(|| invalid_geometry("PPT length conversion overflow"))?;
    let denominator = 914_400_i128;
    let rounded = if numerator >= 0 {
        (numerator + denominator / 2) / denominator
    } else {
        -((-numerator + denominator / 2) / denominator)
    };
    i64::try_from(rounded)
        .map_err(|_| invalid_geometry("PPT length conversion overflow").with_property(property))
}

fn stage_structural<F>(
    file: &Ppt,
    operations: usize,
    limits: &SemanticLimits,
    stage: F,
) -> Result<litchi_ppt::Commit>
where
    F: FnOnce(
        &mut litchi_ppt::package::edit::Edit,
    ) -> std::result::Result<(), litchi_ppt::package::edit::Error>,
{
    let maximum = litchi_ppt::package::edit::MAX_OPERATIONS;
    if operations > maximum {
        return Err(native_edit_error(litchi_ppt::package::edit::Error::Limit {
            kind: litchi_ppt::package::edit::LimitKind::Operations,
            observed: operations,
            maximum,
        }));
    }
    // The before image alone is a lower bound; reject before the native edit
    // captures it. The complete before/after charge is checked after commit.
    let before_bytes = file.retained_source_len().unwrap_or(0);
    let minimum = before_bytes
        .saturating_mul(2)
        .saturating_add(size_of::<litchi_ppt::Patch>());
    if minimum > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(minimum, limits.max_snapshot_bytes()));
    }
    let package =
        file.editable_package_for_write_with_snapshot_limit(limits.max_snapshot_bytes())?;
    let mut edit = package.edit().map_err(native_edit_error)?;
    stage(&mut edit).map_err(native_edit_error)?;
    edit.commit().map_err(native_edit_error)
}

fn stage_plans(
    file: &Ppt,
    plans: Vec<Plan>,
    _limits: &SemanticLimits,
    snapshot_budget: usize,
) -> Result<litchi_ppt::Commit> {
    let operations = plans
        .iter()
        .map(|plan| match plan {
            Plan::Shape(plan) => {
                usize::from(plan.text.is_some()) + usize::from(plan.anchor.is_some())
            },
            Plan::Slide(_) | Plan::Notes(_) => 1,
        })
        .sum::<usize>();
    let maximum = litchi_ppt::package::edit::MAX_OPERATIONS;
    if operations > maximum {
        return Err(native_edit_error(litchi_ppt::package::edit::Error::Limit {
            kind: litchi_ppt::package::edit::LimitKind::Operations,
            observed: operations,
            maximum,
        }));
    }
    // The before image alone is a lower bound; reject before the native edit
    // captures it. The complete before/after charge is checked after commit.
    let before_bytes = file.retained_source_len().unwrap_or(0);
    let minimum = before_bytes
        .saturating_mul(2)
        .saturating_add(size_of::<litchi_ppt::Patch>());
    if minimum > snapshot_budget {
        return Err(snapshot_limit(minimum, snapshot_budget));
    }
    let package = file.editable_package_for_write_with_snapshot_limit(snapshot_budget)?;
    let mut edit = package.edit().map_err(native_edit_error)?;
    if plans.len() > 1
        && plans.iter().all(|plan| {
            matches!(
                plan,
                Plan::Shape(ShapePlan {
                    text: Some(_),
                    anchor: None,
                    ..
                })
            )
        })
    {
        let mut replacements = Vec::new();
        replacements
            .try_reserve_exact(plans.len())
            .map_err(|source| allocation(source, "PPT text batch"))?;
        for plan in plans {
            if let Plan::Shape(ShapePlan {
                target,
                text: Some(text),
                ..
            }) = plan
            {
                replacements.push(litchi_ppt::slide_order::ShapeTextReplacement::new(
                    target, text,
                ));
            }
        }
        let chunks = replacements
            .chunks(litchi_ppt::slide_order::MAX_SHAPE_TEXT_REPLACEMENTS)
            .collect::<Vec<_>>();
        edit.set_shape_text_chunks(&chunks)
            .map_err(native_edit_error)?;
        return edit.commit().map_err(native_edit_error);
    }
    for plan in plans {
        match plan {
            Plan::Shape(plan) => {
                if let Some(text) = plan.text {
                    edit.set_shape_text(plan.target, text)
                        .map_err(native_edit_error)?;
                }
                if let Some(anchor) = plan.anchor {
                    edit.set_shape_anchor(plan.target, anchor)
                        .map_err(native_edit_error)?;
                }
            },
            Plan::Slide(plan) => edit
                .set_slide_hidden(plan.position, plan.hidden)
                .map_err(native_edit_error)?,
            Plan::Notes(plan) => edit
                .set_notes_text(plan.position, plan.text)
                .map_err(native_edit_error)?,
        }
    }
    edit.commit().map_err(native_edit_error)
}

fn finish_commit(
    file: &Ppt,
    native: litchi_ppt::Commit,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.check_source_for_write()?;
    let changed = !native.patch().is_noop();
    let operations = native.patch().operation_count().max(usize::from(changed));
    let native_snapshot_bytes = native
        .patch()
        .before()
        .len()
        .saturating_add(native.patch().after().len())
        .saturating_add(size_of::<litchi_ppt::Patch>());
    if native_snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(
            native_snapshot_bytes,
            limits.max_snapshot_bytes(),
        ));
    }
    if 1 > limits.max_operations() {
        return Err(operation_limit(1, limits.max_operations()));
    }
    let (package, patch) =
        if !changed && file.signature() == crate::format::signature::Signature::Signed {
            // The ordinary writer edits an unsigned private package so a changed
            // publication can never carry signature streams forward.  A true
            // no-op must retain a native patch whose before image is the original
            // signed artifact; otherwise leaf replay would compare against the
            // private unsigned image and the facade would lose content identity.
            let original = file.original_package_for_noop()?;
            file.preflight_snapshot_bytes(limits.max_snapshot_bytes(), 0)?;
            original
                .edit()
                .map_err(native_edit_error)?
                .commit()
                .map_err(native_edit_error)?
                .into_parts()
        } else {
            native.into_parts()
        };
    let snapshot_bytes = native_snapshot_bytes.max(
        patch
            .before()
            .len()
            .saturating_add(patch.after().len())
            .saturating_add(size_of::<litchi_ppt::Patch>()),
    );
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let after = if changed {
        file.with_package(package)?
    } else {
        file.clone()
    };
    file.check_source_for_write()?;
    Ok(WriteCommit {
        before: file.clone(),
        after,
        patch: StepPatch::Primary(patch),
        operations,
        snapshot_bytes,
    })
}

fn bound(kind: LimitKind, observed: usize, maximum: usize) -> Result<()> {
    if observed > maximum {
        return Err(TransactionError::Limit {
            kind,
            observed,
            maximum,
        }
        .into());
    }
    Ok(())
}

fn allocation(source: std::collections::TryReserveError, resource: &'static str) -> Error {
    Error::from(TransactionError::Allocation { source })
        .with_format(FileFormat::Ppt)
        .with_suggestion(resource)
}

fn expanded_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::ExpandedTargets,
        observed,
        maximum,
    })
    .with_format(FileFormat::Ppt)
}

fn staged_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::StagedBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Ppt)
}

fn snapshot_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::SnapshotBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Ppt)
}

fn operation_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::Operations,
        observed,
        maximum,
    })
    .with_format(FileFormat::Ppt)
}

fn native_edit_error(error: litchi_ppt::package::edit::Error) -> Error {
    Error::from(error).with_format(FileFormat::Ppt)
}

fn native_anchor_error(error: litchi_ppt::Error) -> Error {
    Error::from(error).with_format(FileFormat::Ppt)
}

fn empty_target() -> Error {
    Error::new(Code::InvalidArgument, "the mutation target is empty").with_format(FileFormat::Ppt)
}

fn invalid_query() -> Error {
    Error::new(
        Code::InvalidSelector,
        "the mutation selector contains an unsupported predicate",
    )
    .with_format(FileFormat::Ppt)
}

pub(super) fn unsupported_target(message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message).with_format(FileFormat::Ppt)
}

fn out_of_range(message: &'static str) -> Error {
    Error::new(Code::OutOfRange, message).with_format(FileFormat::Ppt)
}

fn invalid_geometry(message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message).with_format(FileFormat::Ppt)
}
