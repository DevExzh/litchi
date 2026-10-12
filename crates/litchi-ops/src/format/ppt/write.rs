//! Unified native-backed writes for legacy PowerPoint presentations.
//!
//! The adapter resolves paths and scoped queries against the existing read
//! navigation, validates properties through the shared registry, and submits
//! the complete call to one native [`litchi_ppt::Package::edit`] transaction.
//! The native owner remains responsible for dependency checks and publication;
//! this module does not maintain a second presentation model or rewrite raw
//! streams itself.

use std::collections::HashSet;

use litchi_core::{FileFormat, Position};
use litchi_ppt::Anchor;
use litchi_ppt::text_edit::Target as NativeTarget;

use super::node::Node;
use super::spec::{
    Kind, PLACEHOLDER_WRITE_TABLE, SHAPE_WRITE_TABLE, SLIDE_WRITE_TABLE, ShapeSetter, SlideSetter,
};
use super::{Nav, Ppt};
use crate::error::{Code, Error, Result};
use crate::props::{IntoProps, Props};
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::value::Scalar;

/// The native proof and immutable endpoints of one unified PPT write.
pub(crate) struct WriteCommit {
    pub(crate) before: Ppt,
    pub(crate) after: Ppt,
    pub(crate) patch: litchi_ppt::Patch,
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
enum Plan {
    Shape(ShapePlan),
    Slide(SlidePlan),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
enum TargetIdentity {
    Slide(usize),
    Shape { slide: usize, shape: usize },
}

/// Admit one PPT shape or slide property call and publish it through the
/// native immutable package owner.
pub(crate) fn set_with_limits<'p, 'v>(
    file: &Ppt,
    target: TargetArg<'p>,
    props: impl IntoProps<'v>,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.admit_edit()?;
    let props = Props::from_props(props)?;
    let nodes = resolve_nodes(file, &target, limits)?;
    let plans = build_plans(nodes, &props, limits)?;
    let native = stage_plans(file, &plans)?;
    file.check_source()?;
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
        let walk = crate::query::run(nav, query.clone())?;
        if !walk.plan().diagnostics().is_empty() {
            return Err(invalid_query());
        }
        for node in walk {
            append_node(&mut nodes, &mut seen, node?, limits)?;
        }
    } else {
        return Err(empty_target());
    }
    file.check_source()?;
    if nodes.is_empty() {
        return Err(
            Error::new(Code::ZeroMatches, "mutation selector matched no element")
                .with_format(FileFormat::Ppt),
        );
    }
    Ok(nodes)
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
    let shape = nodes
        .iter()
        .all(|node| !matches!(node.element(), Kind::Slide));
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
            Plan::Slide(_) => (usize::MAX, usize::MAX),
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
            Plan::Shape(_) => usize::MAX,
        });
        Ok(plans)
    }
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
                    validate_text(text)?;
                    let mut owned = String::new();
                    owned
                        .try_reserve_exact(text.len())
                        .map_err(|source| allocation(source, "PPT shape text"))?;
                    owned.push_str(text);
                    values.text = Some(owned);
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

fn master_from_emu(value: i64, property: &'static str) -> Result<i64> {
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

fn stage_plans(file: &Ppt, plans: &[Plan]) -> Result<litchi_ppt::Commit> {
    let mut edit = file.editable_package()?.edit().map_err(native_edit_error)?;
    for plan in plans {
        match plan {
            Plan::Shape(plan) => {
                if let Some(text) = &plan.text {
                    edit.set_shape_text(plan.target, text.clone())
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
        }
    }
    edit.commit().map_err(native_edit_error)
}

fn finish_commit(
    file: &Ppt,
    native: litchi_ppt::Commit,
    limits: &SemanticLimits,
) -> Result<WriteCommit> {
    file.check_source()?;
    let changed = !native.patch().is_noop();
    let operations = native.patch().operation_count().max(usize::from(changed));
    if 1 > limits.max_operations() {
        return Err(operation_limit(1, limits.max_operations()));
    }
    let (package, patch) = native.into_parts();
    let snapshot_bytes = patch
        .before()
        .len()
        .saturating_add(patch.after().len())
        .saturating_add(size_of::<litchi_ppt::Patch>());
    if snapshot_bytes > limits.max_snapshot_bytes() {
        return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
    }
    let after = if changed {
        file.with_package(package)?
    } else {
        file.clone()
    };
    file.check_source()?;
    Ok(WriteCommit {
        before: file.clone(),
        after,
        patch,
        operations,
        snapshot_bytes,
    })
}

fn validate_text(value: &str) -> Result<()> {
    if !value.chars().all(|character| {
        matches!(
            character,
            '\t' | '\n' | '\r'
                | '\u{20}'..='\u{d7ff}'
                | '\u{e000}'..='\u{fffd}'
                | '\u{10000}'..='\u{10ffff}'
        )
    }) {
        return Err(Error::new(
            Code::InvalidValue,
            "text contains a character that PPT cannot represent",
        )
        .with_format(FileFormat::Ppt)
        .with_property("text"));
    }
    Ok(())
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

fn unsupported_target(message: &'static str) -> Error {
    Error::new(Code::UnsupportedCapability, message).with_format(FileFormat::Ppt)
}

fn invalid_geometry(message: &'static str) -> Error {
    Error::new(Code::InvalidValue, message).with_format(FileFormat::Ppt)
}
