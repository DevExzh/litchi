//! Consecutive direct PPTX shape writes over one native primary edit.
//!
//! The unified transaction owns the run boundary.  This module owns the
//! narrow admission rule and keeps the native package editor as the only
//! projected state.  Plans contain only a stable slide/shape position and
//! the already-coerced payload; no second presentation model is retained.

use super::{Commit, Editable, Pptx, Target, admit_shape_props, allocation, error};
use crate::format::pptx::spec::Kind;
use crate::format::pptx::steps;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::{Code, Error, Props, Result};
use litchi_core::FileFormat;
use litchi_core::Position;
use litchi_core::selector::{Path, Predicate};
use litchi_pptx::shape::owned::{Handle, Scene};
use std::collections::HashSet;

/// One direct shape or placeholder assignment admitted for a coalesced run.
pub(crate) struct Plan {
    target: Target,
    text: Option<String>,
    changed: bool,
    staged_bytes: usize,
}

impl Plan {
    fn key(&self) -> (usize, usize) {
        (self.target.slide, self.target.position)
    }

    fn has_geometry(&self) -> bool {
        self.target.bounds.is_some()
    }

    fn charge(&self) -> usize {
        self.staged_bytes
    }
}

struct Item {
    plan: Plan,
    index: usize,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum BatchKind {
    Text,
    Properties,
}

#[derive(Clone)]
struct CachedScene {
    scene: Scene,
    shapes: Box<[usize]>,
    placeholders: Box<[usize]>,
}

#[allow(
    clippy::large_enum_variant,
    reason = "one inline native editor per run avoids adding a heap owner solely to balance native representations"
)]
enum NativeEdit {
    Source(litchi_pptx::source::Edit),
    Eager(litchi_pptx::opened::Edit),
}

/// A native primary edit shared by one consecutive direct-path run.
pub(crate) struct Run {
    base: Pptx,
    source: Editable,
    edit: NativeEdit,
    items: Vec<Item>,
    seen: HashSet<(usize, usize)>,
    staged_bytes: usize,
    target_count: usize,
    max_staged_bytes: usize,
    max_expanded_targets: usize,
    /// A direct-path run normally walks one slide. Keeping that one detached
    /// scene alive makes each additional ordinal path a cheap indexed lookup;
    /// replacing it when the slide changes bounds derived-cache retention.
    scene: Option<(usize, CachedScene)>,
}

impl Pptx {
    /// Admit only an absolute, two-step, ordinal path naming a direct shape
    /// or placeholder. Queries, identities, `[last()]`, groups, and content
    /// paths return `None` so the ordinary selector writer remains the source
    /// of their resolution and diagnostics.
    pub(crate) fn shape_plan(
        &self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Plan>> {
        self.check_source()?;
        let mut scene = None;
        let result = shape_plan(self, target, props, limits, &mut scene);
        self.check_source().and(result)
    }
}

impl Run {
    /// Start one native edit and its bounded direct-path resolution cache.
    pub(crate) fn new(base: &Pptx, limits: &SemanticLimits) -> Result<Self> {
        base.check_source()?;
        let source = base.editable_owner()?;
        let edit = new_native_edit(&source)?;
        base.check_source()?;
        Ok(Self {
            base: base.clone(),
            source,
            edit,
            items: Vec::new(),
            seen: HashSet::new(),
            staged_bytes: 0,
            target_count: 0,
            max_staged_bytes: limits.max_staged_bytes(),
            max_expanded_targets: limits.max_expanded_targets(),
            scene: None,
        })
    }

    /// Resolve against the run's immutable source owner. A repeated target
    /// is rejected by `requires_barrier` before staging; the coordinator
    /// re-plans it after publishing the preceding run.
    pub(crate) fn shape_plan(
        &mut self,
        target: &TargetArg<'_>,
        props: &Props<'_>,
        limits: &SemanticLimits,
    ) -> Result<Option<Plan>> {
        self.base.check_source()?;
        let result = shape_plan(&self.base, target, props, limits, &mut self.scene);
        self.base.check_source().and(result)
    }

    pub(crate) fn first_index(&self) -> Option<usize> {
        self.items.first().map(|item| item.index)
    }

    /// A run closes before a repeated target, an expanded-target bound, or a
    /// cumulative staged-payload boundary. The coordinator then resolves the
    /// next plan against the newly published PPTX revision.
    pub(crate) fn requires_barrier(&self, plan: &Plan, _limits: &SemanticLimits) -> bool {
        self.seen.contains(&plan.key())
            || self
                .target_count
                .checked_add(1)
                .is_none_or(|count| count > self.max_expanded_targets)
            || self
                .staged_bytes
                .checked_add(plan.charge())
                .is_none_or(|bytes| bytes > self.max_staged_bytes)
    }

    /// Retain one plan and its exact no-op/change result. Native parsing is
    /// deliberately deferred until commit so one run can issue bounded
    /// aggregate setters instead of reparsing the same slide per command.
    pub(crate) fn stage(&mut self, plan: Plan, index: usize) -> Result<bool> {
        if self.seen.contains(&plan.key())
            || self
                .target_count
                .checked_add(1)
                .is_none_or(|count| count > self.max_expanded_targets)
            || self
                .staged_bytes
                .checked_add(plan.charge())
                .is_none_or(|bytes| bytes > self.max_staged_bytes)
        {
            return Err(Error::new(
                Code::InternalError,
                "PPTX shape coalescing barrier was not flushed",
            )
            .with_format(FileFormat::Pptx));
        }
        self.base.check_source()?;
        self.items
            .try_reserve(1)
            .map_err(|source| TransactionError::Allocation { source })?;
        self.seen
            .try_reserve(1)
            .map_err(|source| TransactionError::Allocation { source })?;
        let changed = plan.changed;
        let charge = plan.charge();
        let key = plan.key();
        self.seen.insert(key);
        self.staged_bytes = self.staged_bytes.saturating_add(charge);
        self.target_count = self.target_count.saturating_add(1);
        self.items.push(Item { plan, index });
        self.base.check_source()?;
        Ok(changed)
    }

    /// Publish all retained plans through one native primary edit. Each
    /// chunk retains the leaf's 256-replacement capacity, while disjoint chunks
    /// share one slide publication. A refusal drops the entire edit and is
    /// located by wire-order replay.
    pub(crate) fn commit(self) -> Result<(Pptx, Commit)> {
        let Self {
            base,
            source,
            mut edit,
            items,
            ..
        } = self;
        base.check_source()?;
        let result = (|| {
            let operations = stage_all(&mut edit, &items)?;
            commit_native(&base, source, edit, operations)
        })();
        let freshness = base.check_source();
        match freshness.and(result) {
            Ok(value) => Ok(value),
            Err(error) => Err(locate_failure(&base, &items, error)),
        }
    }
}

/// The path and payload admission shared by the one-shot and run planners.
fn shape_plan(
    file: &Pptx,
    target: &TargetArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
    scene_cache: &mut Option<(usize, CachedScene)>,
) -> Result<Option<Plan>> {
    let Some(path) = target.path() else {
        return Ok(None);
    };
    let Some((slide, kind, ordinal)) = direct_path(path.path()) else {
        return Ok(None);
    };
    let values = admit_shape_props(props, limits)?;
    if values.click || !values.format.is_empty() || values.text.is_none() && !values.has_geometry()
    {
        return Ok(None);
    }
    bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
    let staged_bytes = values.charge();
    bound(
        LimitKind::StagedBytes,
        staged_bytes,
        limits.max_staged_bytes(),
    )?;

    if scene_cache
        .as_ref()
        .is_none_or(|(cached_slide, _)| *cached_slide != slide)
    {
        let Some(scene) = file.slide_scene(slide)? else {
            return Ok(None);
        };
        *scene_cache = Some((slide, cache_scene(scene)?));
    }
    let Some((_, cached)) = scene_cache.as_ref() else {
        return Err(error(
            Code::InternalError,
            "PPTX shape scene cache was lost during admission",
        ));
    };
    let Some(handle) = direct_handle(cached, kind, ordinal)? else {
        return Ok(None);
    };
    let common = handle.view().common();
    // A body without a run slot (an empty placeholder) takes its text
    // through the ordinary body edit, not the coalesced slot writer.
    if values.text.is_some() && common.text().is_none_or(str::is_empty) {
        return Ok(None);
    }
    if values.has_geometry() && kind == DirectKind::Placeholder {
        return Err(error(
            Code::UnsupportedCapability,
            "shape geometry requires a direct p:sp shape",
        ));
    }
    let current_bounds = common
        .bounds()
        .map(|bounds| (bounds.x(), bounds.y(), bounds.width(), bounds.height()));
    let bounds =
        if values.has_geometry() {
            let current = current_bounds.ok_or_else(|| {
                error(
                    Code::MissingProperty,
                    "shape geometry requires a complete local transform",
                )
                .with_property("geometry")
            })?;
            let complete = (
                values.geometry[0].unwrap_or(current.0),
                values.geometry[1].unwrap_or(current.1),
                values.geometry[2].unwrap_or(current.2),
                values.geometry[3].unwrap_or(current.3),
            );
            if complete.2 <= 0 {
                return Err(error(Code::InvalidValue, "shape width must be positive")
                    .with_property("width"));
            }
            if complete.3 <= 0 {
                return Err(error(Code::InvalidValue, "shape height must be positive")
                    .with_property("height"));
            }
            Some(complete)
        } else {
            None
        };
    let mut changed = values
        .text
        .as_deref()
        .is_some_and(|text| common.text() != Some(text));
    if let Some(requested) = values.geometry[0]
        && current_bounds.is_none_or(|bounds| bounds.0 != requested)
    {
        changed = true;
    }
    if let Some(requested) = values.geometry[1]
        && current_bounds.is_none_or(|bounds| bounds.1 != requested)
    {
        changed = true;
    }
    if let Some(requested) = values.geometry[2]
        && current_bounds.is_none_or(|bounds| bounds.2 != requested)
    {
        changed = true;
    }
    if let Some(requested) = values.geometry[3]
        && current_bounds.is_none_or(|bounds| bounds.3 != requested)
    {
        changed = true;
    }
    Ok(Some(Plan {
        target: Target {
            slot: true,
            slide,
            position: handle.position().get(),
            bounds,
        },
        text: values.text,
        changed,
        staged_bytes,
    }))
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

fn cache_scene(scene: Scene) -> Result<CachedScene> {
    let mut shapes = Vec::new();
    let mut placeholders = Vec::new();
    shapes.try_reserve(scene.len()).map_err(allocation)?;
    placeholders.try_reserve(scene.len()).map_err(allocation)?;
    for handle in scene.roots() {
        if steps::tree_kind(&handle) != Some(Kind::Shape) {
            continue;
        }
        shapes.push(handle.position().get());
        if handle.view().common().placeholder().is_some() {
            placeholders.push(handle.position().get());
        }
    }
    Ok(CachedScene {
        scene,
        shapes: shapes.into_boxed_slice(),
        placeholders: placeholders.into_boxed_slice(),
    })
}

fn direct_handle(cached: &CachedScene, kind: DirectKind, ordinal: usize) -> Result<Option<Handle>> {
    let position = match kind {
        DirectKind::Shape => cached.shapes.get(ordinal).copied(),
        DirectKind::Placeholder => cached.placeholders.get(ordinal).copied(),
    };
    let Some(position) = position else {
        return Ok(None);
    };
    cached
        .scene
        .shape(Position::new(position))
        .map_err(|_| error(Code::InternalError, "cached PPTX shape position is invalid"))
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

fn new_native_edit(source: &Editable) -> Result<NativeEdit> {
    match source {
        Editable::Source(package) => package.edit().map(NativeEdit::Source).map_err(native_error),
        Editable::Eager(package) => package.edit().map(NativeEdit::Eager).map_err(native_error),
    }
}

fn native_error(error: litchi_pptx::Error) -> Error {
    Error::from(error).with_format(FileFormat::Pptx)
}

fn stage_all(edit: &mut NativeEdit, items: &[Item]) -> Result<usize> {
    let mut operations = 0_usize;
    for run in items.chunk_by(|left, right| left.plan.target.slide == right.plan.target.slide) {
        let kind = if run.iter().any(|item| item.plan.target.bounds.is_some()) {
            BatchKind::Properties
        } else {
            BatchKind::Text
        };
        operations =
            operations.saturating_add(stage_batch(edit, run[0].plan.target.slide, kind, run)?);
    }
    Ok(operations)
}

// Retain each native chunk's target and text-byte admission while submitting
// all disjoint chunks together for one whole-slide publication.
fn chunk_ranges(items: &[Item]) -> Result<Vec<std::ops::Range<usize>>> {
    let mut chunks = Vec::new();
    let text_limit = litchi_pptx::opened::Limits::default().max_text_bytes();
    let mut start = 0;
    while start < items.len() {
        let maximum = items
            .len()
            .min(start.saturating_add(litchi_pptx::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS));
        let mut end = start;
        let mut text_bytes = 0usize;
        while end < maximum {
            let next = items[end].plan.text.as_ref().map_or(0, String::len);
            if end > start && next > text_limit.saturating_sub(text_bytes) {
                break;
            }
            // Individually oversized values retain native typed admission.
            text_bytes = text_bytes.saturating_add(next);
            end += 1;
        }
        chunks.try_reserve(1).map_err(allocation)?;
        chunks.push(start..end);
        start = end;
    }
    Ok(chunks)
}

fn stage_batch(
    edit: &mut NativeEdit,
    slide: usize,
    kind: BatchKind,
    items: &[Item],
) -> Result<usize> {
    let ranges = chunk_ranges(items)?;
    match kind {
        BatchKind::Text => {
            let mut texts = Vec::new();
            texts.try_reserve_exact(items.len()).map_err(allocation)?;
            for item in items {
                let text =
                    item.plan.text.as_deref().ok_or_else(|| {
                        Error::new(Code::InternalError, "PPTX text batch lost text")
                    })?;
                texts.push(litchi_pptx::opened::ShapeTextReplacement::new(
                    item.plan.target.position.into(),
                    text,
                ));
            }
            let mut chunks = Vec::new();
            chunks.try_reserve_exact(ranges.len()).map_err(allocation)?;
            chunks.extend(ranges.into_iter().map(|range| &texts[range]));
            match edit {
                NativeEdit::Source(edit) => edit
                    .set_shape_text_chunks(slide, &chunks)
                    .map_err(native_error),
                NativeEdit::Eager(edit) => edit
                    .set_shape_text_chunks(slide, &chunks)
                    .map_err(native_error),
            }
        },
        BatchKind::Properties => {
            let mut properties = Vec::new();
            properties
                .try_reserve_exact(items.len())
                .map_err(allocation)?;
            for item in items {
                let mut replacement = litchi_pptx::opened::ShapePropertiesReplacement::new(
                    item.plan.target.position.into(),
                );
                if let Some(text) = item.plan.text.as_deref() {
                    replacement = replacement.with_text(text);
                }
                if let Some(bounds) = item.plan.target.bounds {
                    replacement = replacement.with_bounds(bounds);
                }
                properties.push(replacement);
            }
            let mut chunks = Vec::new();
            chunks.try_reserve_exact(ranges.len()).map_err(allocation)?;
            chunks.extend(ranges.into_iter().map(|range| &properties[range]));
            match edit {
                NativeEdit::Source(edit) => edit
                    .set_shape_property_chunks(slide, &chunks)
                    .map_err(native_error),
                NativeEdit::Eager(edit) => edit
                    .set_shape_property_chunks(slide, &chunks)
                    .map_err(native_error),
            }
        },
    }
}

fn stage_one(edit: &mut NativeEdit, plan: &Plan) -> Result<()> {
    if plan.has_geometry() {
        let mut replacement =
            litchi_pptx::opened::ShapePropertiesReplacement::new(plan.target.position.into());
        if let Some(text) = plan.text.as_deref() {
            replacement = replacement.with_text(text);
        }
        if let Some(bounds) = plan.target.bounds {
            replacement = replacement.with_bounds(bounds);
        }
        match edit {
            NativeEdit::Source(edit) => edit
                .set_shape_properties(plan.target.slide, std::slice::from_ref(&replacement))
                .map(|_| ())
                .map_err(native_error),
            NativeEdit::Eager(edit) => edit
                .set_shape_properties(plan.target.slide, std::slice::from_ref(&replacement))
                .map(|_| ())
                .map_err(native_error),
        }
    } else {
        let text = plan
            .text
            .as_deref()
            .ok_or_else(|| Error::new(Code::InternalError, "PPTX text plan lost text"))?;
        let replacement =
            litchi_pptx::opened::ShapeTextReplacement::new(plan.target.position.into(), text);
        match edit {
            NativeEdit::Source(edit) => edit
                .set_shape_texts(plan.target.slide, std::slice::from_ref(&replacement))
                .map(|_| ())
                .map_err(native_error),
            NativeEdit::Eager(edit) => edit
                .set_shape_texts(plan.target.slide, std::slice::from_ref(&replacement))
                .map(|_| ())
                .map_err(native_error),
        }
    }
}

fn commit_native(
    base: &Pptx,
    source: Editable,
    edit: NativeEdit,
    operations: usize,
) -> Result<(Pptx, Commit)> {
    match edit {
        NativeEdit::Source(edit) => {
            let (target, patch) = edit.commit().map_err(native_error)?.into_parts();
            let changed = !patch.is_noop();
            let target = Editable::Source(target);
            let after = if changed {
                match &target {
                    Editable::Source(package) => base.adopt_source(package.clone())?,
                    Editable::Eager(_) => unreachable!("source edit produced eager target"),
                }
            } else {
                base.clone()
            };
            let commit_after = after.clone();
            Ok((
                after,
                Commit {
                    after: commit_after,
                    source,
                    target,
                    proof: super::Proof::Source(patch),
                    operations,
                },
            ))
        },
        NativeEdit::Eager(edit) => {
            let (target, patch) = edit.commit().map_err(native_error)?.into_parts();
            let changed = !patch.is_noop();
            let target = Editable::Eager(std::sync::Arc::new(target));
            let after = if changed {
                match &target {
                    Editable::Eager(package) => base.adopt_eager(package.clone()),
                    Editable::Source(_) => unreachable!("eager edit produced source target"),
                }
            } else {
                base.clone()
            };
            let commit_after = after.clone();
            Ok((
                after,
                Commit {
                    after: commit_after,
                    source,
                    target,
                    proof: super::Proof::Eager(patch),
                    operations,
                },
            ))
        },
    }
}

fn replay_prefix(source: &Editable, items: &[Item]) -> Result<()> {
    let mut edit = new_native_edit(source)?;
    for item in items {
        stage_one(&mut edit, &item.plan).map_err(|error| error.with_item(item.index))?;
    }
    match edit {
        NativeEdit::Source(edit) => edit.commit().map_err(native_error).map(|_| ()),
        NativeEdit::Eager(edit) => edit.commit().map_err(native_error).map(|_| ()),
    }
}

fn locate_failure(base: &Pptx, items: &[Item], error: Error) -> Error {
    let Some(first) = items.first() else {
        return error;
    };
    if matches!(
        error.code(),
        Code::IoError | Code::Cancelled | Code::InternalError | Code::DocumentProtected
    ) {
        return error.with_item(first.index);
    }
    let source = match base.editable_owner() {
        Ok(source) => source,
        Err(error) => return error.with_item(first.index),
    };
    // Usually failure occurred while staging a particular native item.
    // Replay once in wire order to recover that exact index, rather than
    // replaying every prefix (quadratic native calls on a late failure).
    let mut failure = match replay_prefix(&source, items) {
        Ok(()) => return error.with_item(first.index),
        Err(failure) if failure.item().is_some() => return failure,
        Err(failure) => failure,
    };
    // Only commit-level failures need prefix isolation. Distinct targets
    // preserve shape topology; staging errors above already carry an index.
    let mut low = 1;
    let mut high = items.len();
    while low < high {
        let middle = low + (high - low) / 2;
        match replay_prefix(&source, &items[..middle]) {
            Ok(()) => low = middle + 1,
            Err(error) if error.item().is_some() => return error,
            Err(error) => {
                failure = error;
                high = middle;
            },
        }
    }
    failure.with_item(items[low - 1].index)
}
