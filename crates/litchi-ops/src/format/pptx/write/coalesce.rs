//! Consecutive direct PPTX shape writes over one native primary edit.
//!
//! The unified transaction owns the run boundary.  This module owns the
//! narrow admission rule and keeps the native package editor as the only
//! projected state.  Plans contain only a stable slide/shape position and
//! the already-coerced payload; no second presentation model is retained.
//!
//! Besides shape text and geometry, a run takes character and paragraph
//! formatting of a direct shape, paragraph or run path (record 0333):
//! consecutive formatting plans of a slide publish as one native
//! `set_text_properties` batch, so a formatting batch is one leaf commit.
//! A plan overlapping an earlier plan of the run (the same shape's text,
//! the same paragraph or run) closes the run first, so each plan's change
//! is judged against the text it formats.

use super::text::{self, TextValues};
use super::{Commit, Editable, Pptx, Target, admit_shape_props, allocation, error};
use crate::format::pptx::spec::{Kind, PARAGRAPH_WRITE_TABLE, RUN_WRITE_TABLE};
use crate::format::pptx::steps;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::{Code, Error, Props, Result};
use litchi_core::FileFormat;
use litchi_core::Position;
use litchi_core::selector::{Path, Predicate};
use litchi_pptx::opened::{TextEdit, TextTarget};
use litchi_pptx::shape::owned::{Handle, Scene};
use std::collections::HashMap;

/// One direct shape or placeholder assignment admitted for a coalesced run.
pub(crate) struct Plan {
    target: Target,
    text: Option<String>,
    /// Rich formatting of the shape's body, a paragraph or a run (record
    /// 0333); `None` for a text or geometry plan.
    format: Option<Box<Format>>,
    changed: bool,
    staged_bytes: usize,
}

/// A formatting plan's native target and admitted assignments.
struct Format {
    target: TextTarget,
    scope: Scope,
    values: TextValues,
}

/// What part of a shape's text a plan touches.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Scope {
    /// The whole shape: its text, geometry or every paragraph.
    Shape,
    /// One paragraph, by its index.
    Paragraph(usize),
    /// One run child, by paragraph and child index.
    Run(usize, usize),
}

impl Scope {
    /// Whether two plans of one shape touch the same text.
    fn overlaps(self, other: Self) -> bool {
        match (self, other) {
            (Self::Shape, _) | (_, Self::Shape) => true,
            (Self::Paragraph(left), Self::Paragraph(right))
            | (Self::Paragraph(left), Self::Run(right, _))
            | (Self::Run(left, _), Self::Paragraph(right)) => left == right,
            (Self::Run(left, a), Self::Run(right, b)) => left == right && a == b,
        }
    }
}

impl Plan {
    fn key(&self) -> (usize, usize) {
        (self.target.slide, self.target.position)
    }

    fn scope(&self) -> Scope {
        self.format
            .as_ref()
            .map_or(Scope::Shape, |format| format.scope)
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
    /// Rich formatting through `set_text_properties` (record 0333).
    Format,
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
    /// The scopes each shape's plans touch.
    seen: HashMap<(usize, usize), Vec<Scope>>,
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
            seen: HashMap::new(),
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
        self.overlaps(plan)
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
    /// Whether `plan` touches text an earlier plan of the run touches.
    fn overlaps(&self, plan: &Plan) -> bool {
        let scope = plan.scope();
        self.seen
            .get(&plan.key())
            .is_some_and(|scopes| scopes.iter().any(|seen| seen.overlaps(scope)))
    }

    pub(crate) fn stage(&mut self, plan: Plan, index: usize) -> Result<bool> {
        if self.overlaps(&plan)
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
        let scope = plan.scope();
        let scopes = self.seen.entry(key).or_default();
        scopes
            .try_reserve(1)
            .map_err(|source| TransactionError::Allocation { source })?;
        scopes.push(scope);
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
    let Some((slide, kind, ordinal, text)) = direct_path(path.path()) else {
        return Ok(None);
    };
    if let Some(text) = text {
        return text_plan(file, slide, kind, ordinal, text, props, limits, scene_cache);
    }
    let values = admit_shape_props(props, limits)?;
    if values.click {
        return Ok(None);
    }
    if !values.format.is_empty() {
        if values.text.is_some() || values.has_geometry() || !values.format.coalescible() {
            return Ok(None);
        }
        return format_plan(
            file,
            slide,
            kind,
            ordinal,
            Scope::Shape,
            values.format,
            limits,
            scene_cache,
        );
    }
    if values.text.is_none() && !values.has_geometry() {
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
        format: None,
        changed,
        staged_bytes,
    }))
}

/// The paragraph or run a direct text path names (record 0333).
#[derive(Clone, Copy)]
enum TextStep {
    Paragraph(usize),
    Run(usize, usize),
}

/// A formatting plan for a direct paragraph or run path.
#[allow(
    clippy::too_many_arguments,
    reason = "the shared admission inputs of every plan kind"
)]
fn text_plan(
    file: &Pptx,
    slide: usize,
    kind: DirectKind,
    ordinal: usize,
    step: TextStep,
    props: &Props<'_>,
    limits: &SemanticLimits,
    scene_cache: &mut Option<(usize, CachedScene)>,
) -> Result<Option<Plan>> {
    let table = match step {
        TextStep::Paragraph(_) => {
            text::admit_for(&PARAGRAPH_WRITE_TABLE, props, Operation::Set, limits)
        },
        TextStep::Run(..) => text::admit_for(&RUN_WRITE_TABLE, props, Operation::Set, limits),
    };
    // A click action or a refused key leaves the plan to the ordinary
    // writer, which reports it.
    let Ok(values) = table else {
        return Ok(None);
    };
    if !values.coalescible() {
        return Ok(None);
    }
    match (step, values.text()) {
        // Paragraph text rebuilds the paragraph's runs; a line break in run
        // text is refused by the ordinary writer.
        (TextStep::Paragraph(_), Some(_)) => return Ok(None),
        (TextStep::Run(..), Some(text)) if text.contains(['\n', '\r']) => return Ok(None),
        _ => {},
    }
    let scope = match step {
        TextStep::Paragraph(paragraph) => Scope::Paragraph(paragraph),
        // The run's child index is resolved against the body below.
        TextStep::Run(paragraph, run) => Scope::Run(paragraph, run),
    };
    format_plan(
        file,
        slide,
        kind,
        ordinal,
        scope,
        values,
        limits,
        scene_cache,
    )
}

/// A formatting plan of a direct shape, paragraph or run, resolved against
/// the run's base and judged against the reader's model of its text.
#[allow(
    clippy::too_many_arguments,
    reason = "the shared admission inputs of every plan kind"
)]
fn format_plan(
    file: &Pptx,
    slide: usize,
    kind: DirectKind,
    ordinal: usize,
    scope: Scope,
    values: TextValues,
    limits: &SemanticLimits,
    scene_cache: &mut Option<(usize, CachedScene)>,
) -> Result<Option<Plan>> {
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
    let Ok(Some(body)) = handle.view().text_body() else {
        return Ok(None);
    };
    let shape = handle.position();
    let (target, scope, changed) = match scope {
        Scope::Shape => (
            TextTarget::Shape { shape },
            Scope::Shape,
            body.paragraphs
                .iter()
                .any(|paragraph| values.paragraph_changed(paragraph)),
        ),
        Scope::Paragraph(index) => {
            let Some(paragraph) = body.paragraphs.get(index) else {
                return Ok(None);
            };
            (
                TextTarget::Paragraph {
                    shape,
                    paragraph: Position::new(index),
                },
                Scope::Paragraph(index),
                values.paragraph_changed(paragraph),
            )
        },
        Scope::Run(paragraph_index, ordinal) => {
            let Some(paragraph) = body.paragraphs.get(paragraph_index) else {
                return Ok(None);
            };
            let Some((index, run)) = paragraph
                .runs
                .iter()
                .enumerate()
                .filter(|(_, run)| super::super::text::listed_run(run) == Some(Kind::Run))
                .nth(ordinal)
            else {
                return Ok(None);
            };
            (
                TextTarget::Run {
                    shape,
                    paragraph: Position::new(paragraph_index),
                    run: Position::new(index),
                },
                Scope::Run(paragraph_index, index),
                values.run_changed(run),
            )
        },
    };
    Ok(Some(Plan {
        target: Target {
            slot: true,
            slide,
            position: shape.get(),
            bounds: None,
        },
        text: None,
        format: Some(Box::new(Format {
            target,
            scope,
            values,
        })),
        changed,
        staged_bytes,
    }))
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum DirectKind {
    Shape,
    Placeholder,
}

fn direct_path(path: &Path<'_>) -> Option<(usize, DirectKind, usize, Option<TextStep>)> {
    if !path.is_absolute() {
        return None;
    }
    let steps = path.steps();
    let (slide_step, leaf, text) = match steps {
        [slide_step, leaf] => (slide_step, leaf, None),
        [slide_step, leaf, paragraph] => (
            slide_step,
            leaf,
            Some(TextStep::Paragraph(index_of(
                paragraph,
                &["paragraph", "p", "para"],
            )?)),
        ),
        [slide_step, leaf, paragraph, run] => (
            slide_step,
            leaf,
            Some(TextStep::Run(
                index_of(paragraph, &["paragraph", "p", "para"])?,
                index_of(run, &["run", "r"])?,
            )),
        ),
        _ => return None,
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
    Some((slide.get(), kind, ordinal.get(), text))
}

/// The ordinal of a step named one of `names` with one index predicate.
fn index_of(step: &litchi_core::selector::Step<'_>, names: &[&str]) -> Option<usize> {
    if !names
        .iter()
        .any(|name| step.name().eq_ignore_ascii_case(name))
    {
        return None;
    }
    match step.predicates() {
        [Predicate::Index(index)] => Some(index.get()),
        _ => None,
    }
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
    for run in items.chunk_by(|left, right| {
        left.plan.target.slide == right.plan.target.slide
            && left.plan.format.is_some() == right.plan.format.is_some()
    }) {
        let kind = if run[0].plan.format.is_some() {
            BatchKind::Format
        } else if run.iter().any(|item| item.plan.target.bounds.is_some()) {
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
    if kind == BatchKind::Format {
        let mut edits = Vec::new();
        edits.try_reserve_exact(items.len()).map_err(allocation)?;
        for item in items {
            let format =
                item.plan.format.as_ref().ok_or_else(|| {
                    Error::new(Code::InternalError, "PPTX format batch lost a plan")
                })?;
            edits.push(TextEdit::new(format.target, format.values.properties()));
        }
        // One native call per slide: the chunks keep the per-call limits
        // and share one scene read and publication.
        let mut chunks = Vec::new();
        chunks
            .try_reserve_exact(
                edits
                    .len()
                    .div_ceil(litchi_pptx::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS),
            )
            .map_err(allocation)?;
        chunks.extend(edits.chunks(litchi_pptx::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS));
        return match edit {
            NativeEdit::Source(edit) => edit
                .set_text_property_chunks(Position::new(slide), &chunks)
                .map_err(native_error),
            NativeEdit::Eager(edit) => edit
                .set_text_property_chunks(slide, &chunks)
                .map_err(native_error),
        };
    }
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
        BatchKind::Format => unreachable!("formatting batches return above"),
    }
}

fn stage_one(edit: &mut NativeEdit, plan: &Plan) -> Result<()> {
    if let Some(format) = &plan.format {
        let edits = [TextEdit::new(format.target, format.values.properties())];
        return match edit {
            NativeEdit::Source(edit) => edit
                .set_text_properties(Position::new(plan.target.slide), &edits)
                .map(|_| ())
                .map_err(native_error),
            NativeEdit::Eager(edit) => edit
                .set_text_properties(plan.target.slide, &edits)
                .map(|_| ())
                .map_err(native_error),
        };
    }
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
