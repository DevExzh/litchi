//! Consecutive direct PPT shape writes over one native primary edit.
//!
//! This adapter is intentionally narrower than the ordinary PPT writer.  It
//! admits only absolute ordinal paths to top-level shapes and placeholders;
//! queries, slide targets, raw paths and structural operations are barriers
//! owned by the transaction coordinator.  A run keeps the existing native
//! presentation projection as its only read model and publishes one native
//! package edit.

use std::borrow::Cow;
use std::collections::HashSet;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::FileFormat;
use litchi_core::selector::{Path, Predicate};
use litchi_ppt::shapes::ShapeEnum;
use litchi_ppt::text_edit::Target as NativeTarget;

use super::{ShapeValues, admit_shape_props, allocation, native_edit_error};
use crate::format::ppt::Ppt;
use crate::format::ppt::children::physical_kind;
use crate::format::ppt::node::Node;
use crate::format::ppt::spec::Kind;
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::{Code, Error, Result};

/// One slide's already parsed top-level shape array and ordinal indexes.
///
/// The array is the same source-backed projection used by PPT navigation.  We
/// retain only an `Arc` to that projection plus two small index arrays, so a
/// 1k/10k homogeneous run does not resolve each ordinal by rescanning the
/// slide tree.
#[derive(Clone)]
pub(crate) struct SceneCache {
    slide: usize,
    source: Arc<[ShapeEnum<'static>]>,
    shapes: Box<[usize]>,
    placeholders: Box<[usize]>,
}

/// One direct shape or placeholder assignment admitted for a coalesced run.
pub(crate) struct Plan {
    // Only the first plan carries a cache; staging moves it into the run.
    // A box keeps this one-time transfer small in subsequent retained plans.
    scene: Option<Box<SceneCache>>,
    target: NativeTarget,
    text: Option<String>,
    anchor: Option<litchi_ppt::Anchor>,
    changed: bool,
    staged_bytes: usize,
    /// Original shape text retained for a signed cancellation proof.
    original_text: Option<String>,
}

impl Plan {
    pub(crate) fn with_scene(mut self, scene: Option<SceneCache>) -> Self {
        self.scene = scene.map(Box::new);
        self
    }

    fn key(&self) -> (usize, usize) {
        (self.target.slide().get(), self.target.shape().get())
    }

    fn native_operations(&self) -> usize {
        usize::from(self.text.is_some()) + usize::from(self.anchor.is_some())
    }

    fn charge(&self) -> usize {
        self.staged_bytes
    }
}

struct Item {
    plan: Plan,
    index: usize,
}

/// One native primary edit shared by consecutive direct shape paths.
pub(crate) struct Run {
    base: Ppt,
    source: litchi_ppt::Package,
    edit: litchi_ppt::package::edit::Edit,
    items: Vec<Item>,
    seen: HashSet<(usize, usize)>,
    staged_bytes: usize,
    target_count: usize,
    native_operations: usize,
    max_staged_bytes: usize,
    max_expanded_targets: usize,
    limits: SemanticLimits,
    max_snapshot_bytes: usize,
    scene: Option<SceneCache>,
    allow_repeated: bool,
}

const NATIVE_MAX_OPERATIONS: usize = litchi_ppt::package::edit::MAX_OPERATIONS;

impl Run {
    /// Start a source-checked native edit and preflight the retained snapshot
    /// lower bound before the native editor captures its before image.
    pub(crate) fn new(
        base: &Ppt,
        limits: &SemanticLimits,
        remaining_snapshot: usize,
    ) -> Result<Self> {
        base.check_source_for_write()?;
        let before_bytes = base.retained_source_len().unwrap_or(0);
        let minimum = before_bytes
            .saturating_mul(2)
            .saturating_add(size_of::<litchi_ppt::Patch>());
        if minimum > remaining_snapshot {
            return Err(snapshot_limit(minimum, remaining_snapshot));
        }
        let package = if base.signature() == crate::format::signature::Signature::Signed {
            base.editable_package_for_write_with_snapshot_limit(remaining_snapshot)?
        } else {
            base.editable_package_with_snapshot_limit(remaining_snapshot)?
        };
        let edit = package.edit().map_err(native_edit_error)?;
        base.check_source_for_write()?;
        Ok(Self {
            base: base.clone(),
            source: package,
            edit,
            items: Vec::new(),
            seen: HashSet::new(),
            staged_bytes: 0,
            target_count: 0,
            native_operations: 0,
            max_staged_bytes: limits.max_staged_bytes(),
            max_expanded_targets: limits.max_expanded_targets(),
            limits: *limits,
            max_snapshot_bytes: remaining_snapshot,
            scene: None,
            allow_repeated: base.signature() == crate::format::signature::Signature::Signed,
        })
    }

    /// Resolve one direct path against the run's retained scene cache.
    pub(crate) fn shape_plan(
        &mut self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Plan>> {
        self.base
            .shape_plan_cached(target, props, limits, &mut self.scene)
    }

    /// The first wire index in this run, used to attribute an unrecoverable
    /// native failure.
    pub(crate) fn first_index(&self) -> Option<usize> {
        self.items.first().map(|item| item.index)
    }

    /// Close the run before a repeated native target or any aggregate limit.
    pub(crate) fn requires_barrier(&self, plan: &Plan, _limits: &SemanticLimits) -> bool {
        self.seen.contains(&plan.key())
            && !(self.allow_repeated && plan.text.is_some() && plan.anchor.is_none())
            || self
                .target_count
                .checked_add(1)
                .is_none_or(|count| count > self.max_expanded_targets)
            || self
                .staged_bytes
                .checked_add(plan.charge())
                .is_none_or(|bytes| bytes > self.max_staged_bytes)
            || self
                .native_operations
                .checked_add(plan.native_operations())
                .is_none_or(|count| count > NATIVE_MAX_OPERATIONS)
    }

    /// Retain a plan and its exact per-item changed/no-op result.
    pub(crate) fn stage(&mut self, mut plan: Plan, index: usize) -> Result<bool> {
        if self.requires_barrier(&plan, &self.limits) {
            return Err(Error::new(
                Code::InternalError,
                "PPT shape coalescing barrier was not flushed",
            )
            .with_format(FileFormat::Ppt));
        }
        self.base.check_source_for_write()?;
        let authored = plan
            .text
            .as_ref()
            .map_or(0, |text| text.len())
            .saturating_add(plan.anchor.map_or(0, |_| size_of::<litchi_ppt::Anchor>()));
        let observed = self
            .base
            .retained_source_len()
            .unwrap_or(0)
            .saturating_mul(2)
            .saturating_add(size_of::<litchi_ppt::Patch>())
            .saturating_add(self.staged_bytes.saturating_add(plan.charge()))
            .saturating_add(authored)
            .saturating_add(size_of::<Item>());
        if observed > self.max_snapshot_bytes {
            return Err(snapshot_limit(observed, self.max_snapshot_bytes));
        }
        self.items
            .try_reserve(1)
            .map_err(|source| TransactionError::Allocation { source })?;
        self.seen
            .try_reserve(1)
            .map_err(|source| TransactionError::Allocation { source })?;
        let changed = plan.changed;
        if let Some(scene) = plan.scene.take() {
            self.scene = Some(*scene);
        }
        self.seen.insert(plan.key());
        self.staged_bytes = self.staged_bytes.saturating_add(plan.charge());
        self.target_count = self.target_count.saturating_add(1);
        self.native_operations = self
            .native_operations
            .saturating_add(plan.native_operations());
        self.items.push(Item { plan, index });
        self.base.check_source_for_write()?;
        Ok(changed)
    }

    /// Publish all retained plans through one native primary edit. Text-only
    /// runs publish bounded native text chunks together; mixed text/geometry runs
    /// keep wire order so a late refusal can be attributed exactly.
    pub(crate) fn commit(self) -> Result<(Ppt, super::WriteCommit)> {
        let Self {
            base,
            source,
            mut edit,
            items,
            limits,
            max_snapshot_bytes,
            allow_repeated,
            ..
        } = self;
        let snapshot_limits = limits.with_snapshot_budget(max_snapshot_bytes);
        let logical_noop = allow_repeated && logical_noop(&items);
        // Native text/anchor setters may expand one authored command into
        // several physical operations.  The unified leaf diagnostic is the
        // authored command count, just like the XLS/XLSX/XLSB adapters; the
        // native count remains enforced by `requires_barrier` above.
        let authored_operations = items.len();
        let original = matches!(
            base.signature(),
            crate::format::signature::Signature::Signed
        )
        .then(|| base.original_package_for_noop())
        .transpose()?;
        base.check_source_for_write()?;
        let result = (|| {
            stage_all(&mut edit, &items)?;
            let native = edit.commit().map_err(native_edit_error)?;
            let mut commit = super::finish_commit(&base, native, &snapshot_limits)?;
            commit.operations = authored_operations;
            let staged_snapshot_bytes = commit.snapshot_bytes;
            if logical_noop {
                // Keep the staged native commit for validation, then replace
                // its changed patch with an identity patch over the original
                // signed package.  The leaf remains a real no-op, while the
                // authored operation count stays in transaction diagnostics.
                let original = original
                    .as_ref()
                    .ok_or_else(|| Error::signed(FileFormat::Ppt))?;
                base.preflight_snapshot_bytes(max_snapshot_bytes, 0)?;
                let empty = original
                    .edit()
                    .map_err(native_edit_error)?
                    .commit()
                    .map_err(native_edit_error)?;
                commit = super::finish_commit(&base, empty, &snapshot_limits)?;
                commit.operations = authored_operations;
                commit.snapshot_bytes = commit.snapshot_bytes.max(staged_snapshot_bytes);
            }
            Ok(commit)
        })();
        let freshness = base.check_source_for_write();
        match freshness.and(result) {
            Ok(commit) => Ok((commit.before.clone(), commit)),
            Err(error) => Err(locate_failure(&base, &source, &items, error)),
        }
    }
}

/// Compare the final direct text assignment for each shape with the text read
/// from the immutable run base.  Geometry and mixed plans intentionally fail
/// this proof until their complete native state is retained in the plan.
fn logical_noop(items: &[Item]) -> bool {
    let mut originals = std::collections::HashMap::<(usize, usize), String>::new();
    let mut finals = std::collections::HashMap::<(usize, usize), String>::new();
    for item in items {
        let plan = &item.plan;
        let Some(original) = plan.original_text.as_ref() else {
            return false;
        };
        if plan.anchor.is_some() {
            return false;
        }
        let key = plan.key();
        originals
            .entry((key.0, key.1))
            .or_insert_with(|| original.clone());
        let Some(text) = plan.text.as_ref() else {
            return false;
        };
        finals.insert((key.0, key.1), text.clone());
    }
    !finals.is_empty()
        && finals
            .into_iter()
            .all(|(key, value)| originals.get(&key) == Some(&value))
}

/// Admit one direct shape path. `None` is a deliberate barrier signal: the
/// ordinary writer owns queries, identities, groups, slides and all other
/// target forms.
pub(crate) fn shape_plan(
    file: &Ppt,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
    scene: &mut Option<SceneCache>,
) -> Result<Option<Plan>> {
    let Some(path) = target.path() else {
        return Ok(None);
    };
    // Raw paths are an explicit barrier. Their native identity semantics must
    // be handled by the ordinary writer, even when the visible path happens
    // to have the same direct slide/shape ordinal form.
    if path.raw {
        return Ok(None);
    }
    let Some((slide, kind, ordinal)) = direct_path(path.path()) else {
        return Ok(None);
    };

    if scene.as_ref().is_none_or(|cached| cached.slide != slide) {
        let Some(selected) = file.reader().slide_at(slide)? else {
            return Ok(None);
        };
        let source = selected.shared_shapes()?;
        let mut shapes = Vec::new();
        let mut placeholders = Vec::new();
        shapes
            .try_reserve(source.len())
            .map_err(|source| allocation(source, "PPT coalesced shape indexes"))?;
        placeholders
            .try_reserve(source.len())
            .map_err(|source| allocation(source, "PPT coalesced placeholder indexes"))?;
        for (position, shape) in source.iter().enumerate() {
            if physical_kind(shape) == Some(Kind::Shape) {
                shapes.push(position);
                if matches!(shape, ShapeEnum::Placeholder(_)) {
                    placeholders.push(position);
                }
            }
        }
        *scene = Some(SceneCache {
            slide,
            source,
            shapes: shapes.into_boxed_slice(),
            placeholders: placeholders.into_boxed_slice(),
        });
    }
    let Some(cached) = scene.as_ref() else {
        return Err(Error::new(
            Code::InternalError,
            "PPT shape scene cache was lost during admission",
        )
        .with_format(FileFormat::Ppt));
    };
    let position = match kind {
        DirectKind::Shape => cached.shapes.get(ordinal).copied(),
        DirectKind::Placeholder => cached.placeholders.get(ordinal).copied(),
    };
    let Some(position) = position else {
        return Ok(None);
    };

    let shape = &cached.source[position];
    let property_kind = if matches!(shape, ShapeEnum::Placeholder(_)) {
        Kind::Placeholder
    } else {
        Kind::Shape
    };
    let values = if props.is_empty() {
        ShapeValues {
            text: None,
            geometry: [None; 4],
        }
    } else {
        admit_shape_props(props, property_kind, limits)?
    };
    super::bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
    let staged_bytes = values.charge();
    super::bound(
        LimitKind::StagedBytes,
        staged_bytes,
        limits.max_staged_bytes(),
    )?;

    let has_geometry = values.has_geometry();
    let geometry = values.geometry;
    // Geometry-only and empty payloads follow the one-shot writer's admission
    // path without decoding a text body. A requested text value reads the
    // native body once. Unsigned plans borrow the body for comparison, while
    // signed plans retain the same read as their cancellation proof.
    let current_text = if values.text.is_some() {
        Some(match shape {
            ShapeEnum::TextBox(textbox) => Cow::Borrowed(textbox.text()),
            ShapeEnum::AutoShape(autoshape) => Cow::Borrowed(autoshape.text()),
            _ => Cow::Owned(shape.text().map_err(Error::from)?),
        })
    } else {
        None
    };
    let text_changed = values
        .text
        .as_deref()
        .is_some_and(|text| current_text.as_deref() != Some(text));
    let original_text = (file.signature() == crate::format::signature::Signature::Signed)
        .then(|| current_text.map(|text| text.into_owned()))
        .flatten();
    let text = values.text;
    let anchor = if has_geometry {
        let source = Arc::clone(&cached.source);
        let path = Arc::from(vec![position].into_boxed_slice());
        let node = Node::shape_node(file, &[], slide, Kind::Shape, ordinal, source, path);
        super::build_anchor(&node, &geometry)?
    } else {
        None
    };
    // direct_path admits only top-level ordinal shapes. The cache maps the
    // kind-specific ordinal to its physical index in the native array.
    let target = NativeTarget::new(
        litchi_core::Position::new(slide),
        litchi_core::Position::new(position),
    );
    let changed = text_changed || anchor.is_some();
    Ok(Some(Plan {
        scene: None,
        target,
        text,
        anchor,
        changed,
        staged_bytes,
        original_text,
    }))
}

/// One direct `File::set` of a shape's text through a one-item run, so a
/// single text assignment publishes through the same native chunked text
/// edit a coalesced batch uses rather than the general one-shot writer
/// (record 0377). `None` keeps the ordinary writer: any target or property
/// other than one changed `text` on an unsigned handle, and any admission or
/// publication error, which the ordinary writer then reports with its own
/// diagnostics.
pub(crate) fn single_text_set(
    file: &Ppt,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
    snapshot_budget: usize,
) -> Option<super::WriteCommit> {
    if file.signature() == crate::format::signature::Signature::Signed
        || props.len() != 1
        || props.get("text").is_none()
        || file.admit_edit_for_write().is_err()
    {
        return None;
    }
    let mut scene = None;
    let Ok(Some(plan)) = shape_plan(file, target, props, limits, &mut scene) else {
        return None;
    };
    if !plan.changed || plan.anchor.is_some() || plan.text.is_none() {
        return None;
    }
    let mut run = Run::new(file, limits, snapshot_budget).ok()?;
    if run.requires_barrier(&plan, limits) || run.stage(plan.with_scene(scene), 0).is_err() {
        return None;
    }
    // A refusal is reported by the ordinary writer below, with the
    // diagnostics a single `set` has always given.
    run.commit().ok().map(|(_, commit)| commit)
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum DirectKind {
    Shape,
    Placeholder,
}

fn direct_path(path: &Path<'_>) -> Option<(usize, DirectKind, usize)> {
    if !path.is_absolute() {
        return None;
    }
    let steps = path.steps();
    if steps.len() != 2 {
        return None;
    }
    let [slide_step, leaf] = steps else {
        return None;
    };
    let [Predicate::Index(slide)] = slide_step.predicates() else {
        return None;
    };
    let [Predicate::Index(ordinal)] = leaf.predicates() else {
        return None;
    };
    if !slide_step.name().eq_ignore_ascii_case("slide") {
        return None;
    }
    let kind = if leaf.name().eq_ignore_ascii_case("placeholder")
        || leaf.name().eq_ignore_ascii_case("ph")
    {
        DirectKind::Placeholder
    } else if leaf.name().eq_ignore_ascii_case("shape")
        || leaf.name().eq_ignore_ascii_case("textbox")
        || leaf.name().eq_ignore_ascii_case("title")
    {
        DirectKind::Shape
    } else {
        return None;
    };
    Some((slide.get(), kind, ordinal.get()))
}

fn stage_all(edit: &mut litchi_ppt::package::edit::Edit, items: &[Item]) -> Result<()> {
    let text_only = items.iter().all(|item| item.plan.anchor.is_none())
        && (!items.iter().any(|item| item.plan.original_text.is_some()) || {
            let mut targets = HashSet::new();
            items.iter().all(|item| targets.insert(item.plan.key()))
        });
    if text_only {
        let mut replacements = Vec::new();
        replacements
            .try_reserve(items.len())
            .map_err(|source| allocation(source, "PPT coalesced text replacements"))?;
        for item in items {
            if let Some(text) = item.plan.text.as_deref() {
                replacements.push(litchi_ppt::slide_order::ShapeTextReplacement::new(
                    item.plan.target,
                    text.to_owned(),
                ));
            }
        }
        let chunks = replacements
            .chunks(litchi_ppt::slide_order::MAX_SHAPE_TEXT_REPLACEMENTS)
            .collect::<Vec<_>>();
        edit.set_shape_text_chunks(&chunks)
            .map_err(native_edit_error)?;
        return Ok(());
    }

    for item in items {
        if let Some(text) = item.plan.text.as_deref() {
            edit.set_shape_text(item.plan.target, text.to_owned())
                .map_err(native_edit_error)?;
        }
        if let Some(anchor) = item.plan.anchor {
            edit.set_shape_anchor(item.plan.target, anchor)
                .map_err(native_edit_error)?;
        }
    }
    Ok(())
}

fn stage_one(edit: &mut litchi_ppt::package::edit::Edit, plan: &Plan) -> Result<()> {
    if let Some(text) = plan.text.as_deref() {
        edit.set_shape_text(plan.target, text.to_owned())
            .map_err(native_edit_error)?;
    }
    if let Some(anchor) = plan.anchor {
        edit.set_shape_anchor(plan.target, anchor)
            .map_err(native_edit_error)?;
    }
    Ok(())
}

fn replay_prefix(source: &litchi_ppt::Package, items: &[Item]) -> Result<()> {
    let mut edit = source.edit().map_err(native_edit_error)?;
    for item in items {
        stage_one(&mut edit, &item.plan).map_err(|error| error.with_item(item.index))?;
    }
    edit.commit().map_err(native_edit_error).map(|_| ())
}

fn locate_failure(base: &Ppt, source: &litchi_ppt::Package, items: &[Item], error: Error) -> Error {
    let Some(first) = items.first() else {
        return error;
    };
    if matches!(
        error.code(),
        Code::IoError | Code::Cancelled | Code::DocumentProtected
    ) {
        return error.with_item(first.index);
    }
    if let Err(failure) = replay_prefix(source, items) {
        if failure.item().is_some() {
            return failure;
        }
    } else {
        return error.with_item(first.index);
    }

    let mut low = 0_usize;
    let mut high = items.len();
    let mut failure = error;
    while low + 1 < high {
        let middle = low + (high - low) / 2;
        match replay_prefix(source, &items[..middle]) {
            Ok(()) => low = middle,
            Err(candidate) if candidate.item().is_some() => return candidate,
            Err(candidate) => {
                failure = candidate;
                high = middle;
            },
        }
    }
    base.check_source_for_write()
        .map(|_| failure.with_item(items[high - 1].index))
        .unwrap_or_else(|freshness| freshness.with_item(first.index))
}

fn snapshot_limit(observed: usize, maximum: usize) -> Error {
    Error::from(TransactionError::Limit {
        kind: LimitKind::SnapshotBytes,
        observed,
        maximum,
    })
    .with_format(FileFormat::Ppt)
}
