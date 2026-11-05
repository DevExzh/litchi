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
//!
//! Record 0346 (record 0333 F3) adds paragraph `text`, table cells (their
//! text, formatting, fill, anchor and insets), the paragraphs and runs of
//! cell bodies, and the paragraphs and runs of the speaker notes. Notes
//! pages are separate parts the source-backed slide edit does not stage, so
//! notes plans join a run on eager owners only; a source owner's notes
//! writes keep the ordinary writer, which promotes it.

use super::text::{self, TextValues};
use super::{Commit, Editable, Pptx, Proof, Target, admit_shape_props, allocation, error};
use crate::format::pptx::spec::{Kind, PARAGRAPH_WRITE_TABLE, RUN_WRITE_TABLE};
use crate::format::pptx::steps;
use crate::registry::Operation;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use crate::{Code, Error, Props, Result};
use litchi_core::FileFormat;
use litchi_core::Position;
use litchi_core::selector::{Path, Predicate};
use litchi_pptx::opened::{TextAssignment, TextTarget};
use litchi_pptx::shape::owned::{Handle, Scene};
use std::collections::{HashMap, HashSet};
use std::mem::{size_of, size_of_val};
use std::sync::Arc;

/// One direct shape or placeholder assignment admitted for a coalesced run.
pub(crate) struct Plan {
    target: Target,
    text: Option<String>,
    /// Rich formatting of the shape's body, a paragraph or a run (record
    /// 0333); `None` for a text or geometry plan.
    format: Option<Box<Format>>,
    changed: bool,
    /// Whether the leaf reports the plan's change when the run stages it
    /// (record 0348): a table-cell plan, whose target is not read before
    /// staging; `changed` is then unknown until the commit.
    reported: bool,
    staged_bytes: usize,
    /// Original target text used to prove a signed cancellation.
    original_text: Option<String>,
    /// Whether the target is a text body of the slide's notes page (record
    /// 0346), which a separate native call edits.
    notes: bool,
}

/// A formatting plan's native target and admitted assignments.
struct Format {
    target: TextTarget,
    scope: Scope,
    values: TextValues,
}

/// What part of a shape's text a plan touches.
#[derive(Clone, Copy, PartialEq, Eq, Debug, Hash)]
enum Scope {
    /// The whole shape: its text, geometry or every paragraph.
    Shape,
    /// One paragraph, by its index.
    Paragraph(usize),
    /// One run child, by paragraph and child index.
    Run(usize, usize),
    /// One table cell, by row and column (record 0346).
    Cell(usize, usize),
    /// One paragraph of a table cell's body.
    CellParagraph(usize, usize, usize),
    /// One run child of a table cell's paragraph.
    CellRun(usize, usize, usize, usize),
}

/// The scopes a run's plans of one shape touch, indexed so an overlap test
/// costs O(1) however many cells or paragraphs of the shape a batch
/// formats (record 0346).
#[derive(Default)]
struct Touched {
    shape: bool,
    paragraphs: HashSet<usize>,
    run_paragraphs: HashSet<usize>,
    runs: HashSet<(usize, usize)>,
    cells: HashSet<(usize, usize)>,
    cell_texts: HashSet<(usize, usize)>,
    cell_paragraphs: HashSet<(usize, usize, usize)>,
    cell_run_paragraphs: HashSet<(usize, usize, usize)>,
    cell_runs: HashSet<(usize, usize, usize, usize)>,
}

impl Touched {
    /// Whether `scope` overlaps a scope already touched, as
    /// [`Scope::overlaps`] judges a pair.
    fn overlaps(&self, scope: Scope) -> bool {
        if self.shape {
            return true;
        }
        match scope {
            Scope::Shape => {
                !(self.paragraphs.is_empty()
                    && self.runs.is_empty()
                    && self.cells.is_empty()
                    && self.cell_texts.is_empty())
            },
            Scope::Paragraph(p) => self.paragraphs.contains(&p) || self.run_paragraphs.contains(&p),
            Scope::Run(p, r) => self.paragraphs.contains(&p) || self.runs.contains(&(p, r)),
            Scope::Cell(row, column) => {
                self.cells.contains(&(row, column)) || self.cell_texts.contains(&(row, column))
            },
            Scope::CellParagraph(row, column, p) => {
                self.cells.contains(&(row, column))
                    || self.cell_paragraphs.contains(&(row, column, p))
                    || self.cell_run_paragraphs.contains(&(row, column, p))
            },
            Scope::CellRun(row, column, p, r) => {
                self.cells.contains(&(row, column))
                    || self.cell_paragraphs.contains(&(row, column, p))
                    || self.cell_runs.contains(&(row, column, p, r))
            },
        }
    }

    fn insert(&mut self, scope: Scope) -> Result<()> {
        fn add<T: std::hash::Hash + Eq>(set: &mut HashSet<T>, value: T) -> Result<()> {
            set.try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;
            set.insert(value);
            Ok(())
        }
        match scope {
            Scope::Shape => self.shape = true,
            Scope::Paragraph(p) => add(&mut self.paragraphs, p)?,
            Scope::Run(p, r) => {
                add(&mut self.run_paragraphs, p)?;
                add(&mut self.runs, (p, r))?;
            },
            Scope::Cell(row, column) => add(&mut self.cells, (row, column))?,
            Scope::CellParagraph(row, column, p) => {
                add(&mut self.cell_texts, (row, column))?;
                add(&mut self.cell_paragraphs, (row, column, p))?;
            },
            Scope::CellRun(row, column, p, r) => {
                add(&mut self.cell_texts, (row, column))?;
                add(&mut self.cell_run_paragraphs, (row, column, p))?;
                add(&mut self.cell_runs, (row, column, p, r))?;
            },
        }
        Ok(())
    }
}

impl Plan {
    fn key(&self) -> (usize, usize, bool) {
        (self.target.slide, self.target.position, self.notes)
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
    // Notes have their own scene. Planning them must not read unrelated
    // slide shapes (which can legitimately lie outside the edit profile).
    scene: Option<Scene>,
    shapes: Box<[usize]>,
    placeholders: Box<[usize]>,
    /// The scene positions of the slide's root tables, in `table[M]` order
    /// (record 0346).
    tables: Box<[usize]>,
    /// Each table indexed so far, by scene position: its cells'
    /// paragraphs and text runs, indexed once per run (records 0346, 0348).
    cells: HashMap<usize, Arc<Cells>>,
    /// The notes page's body placeholder: its notes-scene position and
    /// body; `Some(None)` when the slide has none (record 0346).
    notes: Option<NotesBody>,
}

/// A notes page's body placeholder: its notes-scene position and the
/// paragraphs and text runs of its body; `None` when the slide has none
/// (records 0346, 0348).
pub(crate) type NotesBody = Option<(
    usize,
    Arc<litchi_pptx::shape::BodyTextIndex>,
    Option<Arc<litchi_drawingml::text::body::Body>>,
)>;

type CapturedBody = Arc<litchi_drawingml::text::body::Body>;
type CapturedRow = Vec<Option<CapturedBody>>;

/// One table's cells, by row and position in the row: the paragraphs and
/// text runs plans address (record 0348). Their properties and text are not
/// read: the leaf reports each cell plan's change when the run stages it.
pub(crate) struct Cells {
    index: litchi_pptx::table::CellTextIndex,
    /// Full cell bodies are retained only for signed text plans, where they
    /// provide the native value used by the cancellation proof. Ordinary
    /// unsigned runs keep the historical index-only path.
    bodies: Option<Vec<CapturedRow>>,
}

impl Cells {
    fn capture_charge(&self) -> usize {
        let Some(rows) = &self.bodies else {
            return 0;
        };
        rows.iter()
            .fold(size_of_val(rows.as_slice()), |total, row| {
                total
                    .saturating_add(size_of_val(row.as_slice()))
                    .saturating_add(
                        row.iter()
                            .flatten()
                            .map(|body| body_capture_charge(body))
                            .sum::<usize>(),
                    )
            })
    }
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
    seen: HashMap<(usize, usize, bool), Touched>,
    staged_bytes: usize,
    max_snapshot_bytes: usize,
    target_count: usize,
    max_staged_bytes: usize,
    max_expanded_targets: usize,
    /// A direct-path run normally walks one slide. Keeping that one detached
    /// scene alive makes each additional ordinal path a cheap indexed lookup;
    /// replacing it when the slide changes bounds derived-cache retention.
    scene: Option<(usize, CachedScene)>,
    allow_repeated: bool,
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
        self.check_source_for_write()?;
        let mut scene = None;
        let result = shape_plan(self, target, props, limits, &mut scene);
        self.check_source_for_write().and(result)
    }
}

impl Run {
    /// Start one native edit and its bounded direct-path resolution cache.
    pub(crate) fn new(
        base: &Pptx,
        limits: &SemanticLimits,
        remaining_snapshot: usize,
    ) -> Result<Self> {
        base.check_source_for_write()?;
        let source = if base.signature() == crate::format::signature::Signature::Signed {
            base.editable_owner_for_write()?
        } else {
            base.editable_owner()?
        };
        let edit = new_native_edit_with_snapshot_budget(base, &source, remaining_snapshot)?;
        base.check_source_for_write()?;
        Ok(Self {
            base: base.clone(),
            source,
            edit,
            items: Vec::new(),
            seen: HashMap::new(),
            staged_bytes: 0,
            max_snapshot_bytes: remaining_snapshot,
            target_count: 0,
            max_staged_bytes: limits.max_staged_bytes(),
            max_expanded_targets: limits.max_expanded_targets(),
            scene: None,
            allow_repeated: base.signature() == crate::format::signature::Signature::Signed,
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
        self.base.check_source_for_write()?;
        let result = shape_plan(&self.base, target, props, limits, &mut self.scene);
        self.base.check_source_for_write().and(result)
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
        if self.allow_repeated && plan.text.is_some() && plan.format.is_none() && !plan.notes {
            return false;
        }
        let scope = plan.scope();
        self.seen
            .get(&plan.key())
            .is_some_and(|touched| touched.overlaps(scope))
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
        self.base.check_source_for_write()?;
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
        self.seen.entry(key).or_default().insert(scope)?;
        self.staged_bytes = self.staged_bytes.saturating_add(charge);
        self.target_count = self.target_count.saturating_add(1);
        self.items.push(Item { plan, index });
        self.base.check_source_for_write()?;
        Ok(changed)
    }

    /// Publish all retained plans through one native primary edit. Each
    /// chunk retains the leaf's 256-replacement capacity, while disjoint chunks
    /// share one slide publication. A refusal drops the entire edit and is
    /// located by wire-order replay.
    ///
    /// The third value holds, for each plan whose change the leaf reports
    /// (record 0348), its wire index and change; [`Self::stage`] returned
    /// `false` for it.
    pub(crate) fn commit(self) -> Result<(Pptx, Commit, Vec<(usize, bool)>)> {
        let Self {
            base,
            source,
            mut edit,
            items,
            allow_repeated,
            max_snapshot_bytes,
            ..
        } = self;
        let logical_noop = allow_repeated && logical_noop(&items);
        let authored_operations = items.len();
        base.check_source_for_write()?;
        let result = (|| {
            let mut changes = Vec::new();
            let operations = stage_all(&mut edit, &items, &mut changes)?;
            let (mut after, mut commit) = commit_native(&base, source, edit, operations)?;
            commit.operations = authored_operations;
            // The native proof retains endpoints for both signed and
            // unsigned packages.  Measure it once through the commit cache
            // so unsigned runs are charged and cannot bypass the snapshot
            // limit.
            let staged_snapshot_bytes = commit.measure_snapshot_bytes();
            bound(
                LimitKind::SnapshotBytes,
                staged_snapshot_bytes,
                max_snapshot_bytes,
            )?;
            let canonical_noop = base.signature() == crate::format::signature::Signature::Signed
                && (logical_noop || proof_is_noop(&commit.proof));
            if canonical_noop {
                // The staged commit validates all intermediate operations. Its
                // changed patch is then discarded in favor of a native empty
                // edit over the original signed owner.
                let original = base.original_owner_for_noop()?;
                // Replacement is a peak-accounted phase, not an overlap: the
                // staged target and proof must be released before the
                // canonical signed owner is captured.  Keep only the
                // measured peak and authored operation count below.
                drop(commit);
                drop(after);
                (after, commit) = commit_native(
                    &base,
                    original.clone(),
                    new_native_edit_with_snapshot_budget(&base, &original, max_snapshot_bytes)?,
                    operations,
                )?;
                commit.operations = authored_operations;
                commit.snapshot_bytes = commit.snapshot_bytes().max(staged_snapshot_bytes);
                bound(
                    LimitKind::SnapshotBytes,
                    commit.snapshot_bytes(),
                    max_snapshot_bytes,
                )?;
            }
            Ok((after, commit, changes))
        })();
        let freshness = base.check_source_for_write();
        match freshness.and(result) {
            Ok(value) => Ok(value),
            Err(error) => Err(locate_failure(&base, &items, error, max_snapshot_bytes)),
        }
    }
}

fn proof_is_noop(proof: &Proof) -> bool {
    match proof {
        Proof::Source(patch) => patch.is_noop(),
        Proof::Eager(patch) => patch.is_noop(),
        Proof::Metadata(patch) => patch.is_noop(),
        Proof::Combined(proofs) => proofs.iter().all(proof_is_noop),
    }
}

/// Compare each text plan's final assignment with the immutable base. Geometry
/// and formatting-only plans retain no complete original state and therefore
/// remain a real commit unless the native patch itself proves an exact no-op.
fn logical_noop(items: &[Item]) -> bool {
    let mut originals = HashMap::<(usize, usize, bool, Scope), String>::new();
    let mut finals = HashMap::<(usize, usize, bool, Scope), String>::new();
    for item in items {
        let plan = &item.plan;
        let Some(original) = plan.original_text.as_ref() else {
            return false;
        };
        if plan.has_geometry()
            || plan
                .format
                .as_ref()
                .is_some_and(|format| !format.values.is_text_only())
        {
            return false;
        }
        let key = (
            plan.target.slide,
            plan.target.position,
            plan.notes,
            plan.scope(),
        );
        originals.entry(key).or_insert_with(|| original.clone());
        let text = match plan.text.as_ref() {
            Some(text) => text.clone(),
            None => match plan.format.as_ref().and_then(|format| format.values.text()) {
                Some(text) => text.to_owned(),
                None => return false,
            },
        };
        finals.insert(key, text);
    }
    !finals.is_empty()
        && finals
            .into_iter()
            .all(|(key, value)| originals.get(&key) == Some(&value))
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
        return story_plan(file, path, props, limits, scene_cache);
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
        .is_none_or(|(cached_slide, cached)| *cached_slide != slide || cached.scene.is_none())
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
    let current_text = values.text.as_ref().and_then(|_| {
        common.text().map(|text| {
            if file.signature() == crate::format::signature::Signature::Signed {
                std::borrow::Cow::Owned(text.to_owned())
            } else {
                std::borrow::Cow::Borrowed(text)
            }
        })
    });
    let mut changed = values
        .text
        .as_deref()
        .is_some_and(|text| current_text.as_deref() != Some(text));
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
    let original_text = (file.signature() == crate::format::signature::Signature::Signed)
        .then(|| current_text.map(|text| text.into_owned()))
        .flatten();
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
        reported: false,
        staged_bytes,
        original_text,
        notes: false,
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
    // A line break in run text is refused by the ordinary writer.
    // Paragraph text rebuilds the paragraph's runs (record 0346).
    if matches!(step, TextStep::Run(..))
        && values
            .text()
            .is_some_and(|text| text.contains(['\n', '\r']))
    {
        return Ok(None);
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
        .is_none_or(|(cached_slide, cached)| *cached_slide != slide || cached.scene.is_none())
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
    let (target, scope, changed, original_text) = match scope {
        Scope::Shape => (
            TextTarget::Shape { shape },
            Scope::Shape,
            body.paragraphs
                .iter()
                .any(|paragraph| values.paragraph_changed(paragraph)),
            None,
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
                paragraph_changed(&values, paragraph),
                (file.signature() == crate::format::signature::Signature::Signed
                    && values.text().is_some())
                .then(|| paragraph.text()),
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
                (file.signature() == crate::format::signature::Signature::Signed
                    && values.text().is_some())
                .then(|| run.text.clone()),
            )
        },
        // Cell scopes are planned by `story_plan`.
        Scope::Cell(..) | Scope::CellParagraph(..) | Scope::CellRun(..) => return Ok(None),
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
        reported: false,
        staged_bytes,
        original_text,
        notes: false,
    }))
}

/// Whether `values` change a paragraph, its text included (record 0346).
fn paragraph_changed(
    values: &TextValues,
    paragraph: &litchi_drawingml::text::body::Paragraph,
) -> bool {
    values.paragraph_changed(paragraph)
        || values.text().is_some_and(|text| text != paragraph.text())
}

/// A cell or notes path (record 0346): zero-based ordinals.
#[derive(Clone, Copy)]
struct StoryPath {
    slide: usize,
    /// `(table, row, cell)` of a table path; `None` for the notes page.
    cell: Option<(usize, usize, usize)>,
    paragraph: Option<usize>,
    run: Option<usize>,
}

impl StoryPath {
    const fn notes(&self) -> bool {
        self.cell.is_none()
    }
}

/// An absolute, all-ordinal table-cell path (`/slide[N]/table[M]/tr[R]/
/// tc[C]`, optionally `/paragraph[K]` and `/run[J]`) or notes path
/// (`/slide[N]/notes/paragraph[K]`, optionally `/run[J]`).
fn story_path(path: &Path<'_>) -> Option<StoryPath> {
    if !path.is_absolute() {
        return None;
    }
    let steps = path.steps();
    let [slide, rest @ ..] = steps else {
        return None;
    };
    let slide = index_of(slide, &["slide"])?;
    let (cell, text) = match rest {
        [notes, text @ ..]
            if notes.name().eq_ignore_ascii_case("notes") && notes.predicates().is_empty() =>
        {
            (None, text)
        },
        [table, row, cell, text @ ..] => (
            Some((
                index_of(table, &["table"])?,
                index_of(row, &["tr"])?,
                index_of(cell, &["tc"])?,
            )),
            text,
        ),
        _ => return None,
    };
    let (paragraph, run) = match text {
        [] if cell.is_some() => (None, None),
        [paragraph] => (
            Some(index_of(paragraph, &["paragraph", "p", "para"])?),
            None,
        ),
        [paragraph, run] => (
            Some(index_of(paragraph, &["paragraph", "p", "para"])?),
            Some(index_of(run, &["run", "r"])?),
        ),
        _ => return None,
    };
    Some(StoryPath {
        slide,
        cell,
        paragraph,
        run,
    })
}

/// A formatting plan of a table cell, a cell paragraph or run, or a notes
/// paragraph or run (record 0346). Tables and the notes body are read once
/// per run and slide (the run's base is immutable); anything the plan
/// cannot judge falls back to the ordinary writer.
#[allow(
    clippy::too_many_lines,
    reason = "one admission, resolution and judgement per story leaf"
)]
fn story_plan(
    file: &Pptx,
    path: &crate::path::PathArg<'_>,
    props: &Props<'_>,
    limits: &SemanticLimits,
    scene_cache: &mut Option<(usize, CachedScene)>,
) -> Result<Option<Plan>> {
    use crate::format::pptx::spec::CELL_WRITE_TABLE;
    let Some(story) = story_path(path.path()) else {
        return Ok(None);
    };
    let notes = story.notes();
    if notes && !file.edits_eagerly() {
        return Ok(None);
    }
    let table = match (story.paragraph, story.run) {
        (None, _) => text::admit_for(&CELL_WRITE_TABLE, props, Operation::Set, limits),
        (Some(_), None) => text::admit_for(&PARAGRAPH_WRITE_TABLE, props, Operation::Set, limits),
        (Some(_), Some(_)) => text::admit_for(&RUN_WRITE_TABLE, props, Operation::Set, limits),
    };
    let Ok(values) = table else {
        return Ok(None);
    };
    let admitted = match (story.paragraph, story.run) {
        (None, _) => values.cell_coalescible(),
        (Some(_), None) => values.coalescible(),
        (Some(_), Some(_)) => {
            values.coalescible()
                && !values
                    .text()
                    .is_some_and(|text| text.contains(['\n', '\r']))
        },
    };
    if !admitted {
        return Ok(None);
    }
    bound(LimitKind::ExpandedTargets, 1, limits.max_expanded_targets())?;
    let mut staged_bytes = values.charge();
    bound(
        LimitKind::StagedBytes,
        staged_bytes,
        limits.max_staged_bytes(),
    )?;
    if scene_cache.as_ref().is_none_or(|(cached_slide, cached)| {
        *cached_slide != story.slide || (!notes && cached.scene.is_none())
    }) {
        let cached = if notes {
            CachedScene {
                scene: None,
                shapes: Box::default(),
                placeholders: Box::default(),
                tables: Box::default(),
                cells: HashMap::new(),
                notes: None,
            }
        } else {
            let Some(scene) = file.slide_scene(story.slide)? else {
                return Ok(None);
            };
            cache_scene(scene)?
        };
        *scene_cache = Some((story.slide, cached));
    }
    let Some((_, cached)) = scene_cache.as_mut() else {
        return Err(error(
            Code::InternalError,
            "PPTX shape scene cache was lost during admission",
        ));
    };
    let (shape, target, scope, changed, original_text) = match story.cell {
        None => {
            let notes_was_cached = cached.notes.is_some();
            if !notes_was_cached {
                let capture_text = file.signature() == crate::format::signature::Signature::Signed;
                let remaining = limits.max_staged_bytes().saturating_sub(staged_bytes);
                let body = file.notes_body(story.slide, || {
                    read_notes(file, story.slide, capture_text, remaining)
                })?;
                cached.notes = Some(body);
            }
            if !notes_was_cached {
                let charge = cached
                    .notes
                    .as_ref()
                    .and_then(|notes| notes.as_ref())
                    .and_then(|(_, _, body)| body.as_ref())
                    .map(|body| body_capture_charge(body.as_ref()))
                    .unwrap_or(0);
                staged_bytes = staged_bytes.saturating_add(charge);
                bound(
                    LimitKind::StagedBytes,
                    staged_bytes,
                    limits.max_staged_bytes(),
                )?;
            }
            let Some(Some((position, index, body))) = &cached.notes else {
                return Ok(None);
            };
            // The paragraph and listed run the text step names, and the
            // run's child index; the leaf reports the change (record 0348).
            let Some(paragraph) = story.paragraph.filter(|p| *p < index.paragraph_count()) else {
                return Ok(None);
            };
            let run = match story.run {
                None => None,
                Some(ordinal) => match index.text_run(paragraph, ordinal) {
                    Some(run) => Some(run),
                    None => return Ok(None),
                },
            };
            let shape = Position::new(*position);
            let changed = false;
            let original_text = body
                .as_ref()
                .and_then(|body| story_text(body, story.paragraph, story.run));
            let (target, scope) = match run {
                None => (
                    TextTarget::Paragraph {
                        shape,
                        paragraph: Position::new(paragraph),
                    },
                    Scope::Paragraph(paragraph),
                ),
                Some(run) => (
                    TextTarget::Run {
                        shape,
                        paragraph: Position::new(paragraph),
                        run: Position::new(run),
                    },
                    Scope::Run(paragraph, run),
                ),
            };
            (shape, target, scope, changed, original_text)
        },
        Some((table, row, column)) => {
            let Some(&frame) = cached.tables.get(table) else {
                return Ok(None);
            };
            let cells = match cached.cells.get(&frame) {
                Some(cells) => Arc::clone(cells),
                None => {
                    let capture_text =
                        file.signature() == crate::format::signature::Signature::Signed;
                    let remaining = limits.max_staged_bytes().saturating_sub(staged_bytes);
                    let Some(cells) = file.table_cells(story.slide, frame, || {
                        read_cells(cached, frame, capture_text, remaining)
                    })?
                    else {
                        return Ok(None);
                    };
                    cached.cells.try_reserve(1).map_err(allocation)?;
                    let capture_charge = cells.capture_charge();
                    cached.cells.insert(frame, Arc::clone(&cells));
                    staged_bytes = staged_bytes.saturating_add(capture_charge);
                    bound(
                        LimitKind::StagedBytes,
                        staged_bytes,
                        limits.max_staged_bytes(),
                    )?;
                    cells
                },
            };
            let index = &cells.index;
            if !index.has_cell(row, column) {
                return Ok(None);
            }
            let shape = Position::new(frame);
            let original_text = cells.bodies.as_ref().and_then(|bodies| {
                let body = bodies.get(row)?.get(column)?.as_ref()?;
                story_text(body, story.paragraph, story.run)
            });
            let (row_at, column_at) = (Position::new(row), Position::new(column));
            // The leaf reports the change when the run stages the plan
            // (record 0348).
            let unknown = false;
            match (story.paragraph, story.run) {
                (None, _) => (
                    shape,
                    TextTarget::Cell {
                        shape,
                        row: row_at,
                        column: column_at,
                    },
                    Scope::Cell(row, column),
                    unknown,
                    original_text,
                ),
                (Some(paragraph), run) => {
                    if index
                        .paragraph_count(row, column)
                        .is_none_or(|count| paragraph >= count)
                    {
                        return Ok(None);
                    }
                    match run {
                        None => (
                            shape,
                            TextTarget::CellParagraph {
                                shape,
                                row: row_at,
                                column: column_at,
                                paragraph: Position::new(paragraph),
                            },
                            Scope::CellParagraph(row, column, paragraph),
                            unknown,
                            original_text,
                        ),
                        Some(ordinal) => {
                            let Some(run) = index.text_run(row, column, paragraph, ordinal) else {
                                return Ok(None);
                            };
                            (
                                shape,
                                TextTarget::CellRun {
                                    shape,
                                    row: row_at,
                                    column: column_at,
                                    paragraph: Position::new(paragraph),
                                    run: Position::new(run),
                                },
                                Scope::CellRun(row, column, paragraph, run),
                                unknown,
                                original_text,
                            )
                        },
                    }
                },
            }
        },
    };
    Ok(Some(Plan {
        target: Target {
            slot: true,
            slide: story.slide,
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
        reported: true,
        staged_bytes,
        original_text,
        notes,
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
    let mut tables = Vec::new();
    for handle in scene.roots() {
        if steps::tree_kind(&handle) == Some(Kind::Table) {
            tables.try_reserve(1).map_err(allocation)?;
            tables.push(handle.position().get());
        }
    }
    Ok(CachedScene {
        scene: Some(scene),
        shapes: shapes.into_boxed_slice(),
        placeholders: placeholders.into_boxed_slice(),
        tables: tables.into_boxed_slice(),
        cells: HashMap::new(),
        notes: None,
    })
}

/// Read slide `slide`'s notes body placeholder (record 0346).
fn read_notes(
    file: &Pptx,
    slide: usize,
    capture_text: bool,
    max_staged_bytes: usize,
) -> Result<NotesBody> {
    // As the navigator resolves `/slide[N]/notes/paragraph[K]`: the notes
    // page must have notes text, and its first body placeholder holds the
    // paragraphs (record 0348).
    let Some((true, scene)) = file.slide_notes_view(slide)? else {
        return Ok(None);
    };
    let found = scene.roots().find(|handle| {
        handle
            .view()
            .common()
            .placeholder()
            .is_some_and(|placeholder| placeholder.kind() == Some("body"))
    });
    Ok(match found {
        // `Slide::notes` has read the body (its text): its structure is
        // indexed without reading it again (record 0348).
        Some(handle) => match handle.view().common().read_text_index() {
            Ok(Some(index)) => {
                let body = if capture_text {
                    let body = handle.view().common().text_body()?;
                    if let Some(body) = &body {
                        bound(
                            LimitKind::StagedBytes,
                            body_capture_charge(body),
                            max_staged_bytes,
                        )?;
                    }
                    body.map(Arc::new)
                } else {
                    None
                };
                Some((handle.position().get(), Arc::new(index), body))
            },
            _ => None,
        },
        None => None,
    })
}

/// Read every cell of the table at scene position `position` once.
fn read_cells(
    cached: &CachedScene,
    position: usize,
    capture_text: bool,
    max_staged_bytes: usize,
) -> Result<Option<Arc<Cells>>> {
    let Some(handle) = cached
        .scene
        .as_ref()
        .ok_or_else(|| error(Code::InternalError, "PPTX table plan has no main scene"))?
        .shape(Position::new(position))
        .map_err(|_| error(Code::InternalError, "cached PPTX table position is invalid"))?
    else {
        return Ok(None);
    };
    let index = handle.cell_text_index()?;
    let bodies = if capture_text {
        let rows = handle.table()?.cell_contents()?;
        // Account for the complete retained shape before allocating any
        // `Arc<Body>`. The caller passes the remaining run budget, so this
        // preflight keeps a large signed table from bypassing the limit while
        // a partially filled cache is still private.
        let mut retained = size_of::<Vec<CapturedRow>>()
            .saturating_add(rows.len().saturating_mul(size_of::<CapturedRow>()));
        for row in &rows {
            retained = retained
                .saturating_add(row.len().saturating_mul(size_of::<Option<CapturedBody>>()));
            for (_, body) in row {
                if let Some(body) = body {
                    retained = retained.saturating_add(body_capture_charge(body));
                }
            }
        }
        if retained > max_staged_bytes {
            return Err(TransactionError::Limit {
                kind: LimitKind::StagedBytes,
                observed: retained,
                maximum: max_staged_bytes,
            }
            .into());
        }
        let mut output = Vec::new();
        output.try_reserve_exact(rows.len()).map_err(allocation)?;
        for row in rows {
            let mut cells = Vec::new();
            cells.try_reserve_exact(row.len()).map_err(allocation)?;
            for (_, body) in row {
                let Some(body) = body else {
                    cells.push(None);
                    continue;
                };
                cells.push(Some(Arc::new(body)));
            }
            output.push(cells);
        }
        Some(output)
    } else {
        None
    };
    Ok(Some(Arc::new(Cells { index, bodies })))
}

fn body_capture_charge(body: &litchi_drawingml::text::body::Body) -> usize {
    let mut bytes = size_of_val(body);
    for paragraph in &body.paragraphs {
        bytes = bytes.saturating_add(size_of_val(paragraph)).saturating_add(
            paragraph
                .runs
                .iter()
                .map(|run| {
                    size_of_val(run)
                        .saturating_add(run.text.len())
                        .saturating_add(size_of_val(&run.properties))
                })
                .sum::<usize>(),
        );
    }
    bytes
}

fn story_text(
    body: &litchi_drawingml::text::body::Body,
    paragraph: Option<usize>,
    run: Option<usize>,
) -> Option<String> {
    let Some(paragraph) = paragraph else {
        return Some(body.text());
    };
    let paragraph = body.paragraphs.get(paragraph)?;
    match run {
        Some(run) => paragraph.runs.get(run).map(|run| run.text.clone()),
        None => Some(paragraph.text()),
    }
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
        .as_ref()
        .ok_or_else(|| error(Code::InternalError, "PPTX shape plan has no main scene"))?
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

/// Build a native editor whose own encoded-patch limit is the transaction's
/// remaining snapshot budget.  The native packages perform the exact
/// before/after wire-size preflight at commit; the facade only contributes the
/// remaining budget and never substitutes a Rust struct size for that proof.
fn new_native_edit_with_snapshot_budget(
    _base: &Pptx,
    source: &Editable,
    remaining_snapshot: usize,
) -> Result<NativeEdit> {
    let native_limit = remaining_snapshot;
    match source {
        Editable::Source(package) => package
            .edit_with_options(
                litchi_pptx::source::EditOptions::default().max_patch_bytes(native_limit),
            )
            .map(NativeEdit::Source)
            .map_err(native_error),
        Editable::Eager(package) => {
            let options = litchi_pptx::opened::EditOptions::default()
                .limits(litchi_pptx::opened::Limits::default().with_max_patch_bytes(native_limit));
            package
                .edit_with_options(options)
                .map(NativeEdit::Eager)
                .map_err(native_error)
        },
    }
}

fn native_error(error: litchi_pptx::Error) -> Error {
    Error::from(error).with_format(FileFormat::Pptx)
}

fn stage_all(
    edit: &mut NativeEdit,
    items: &[Item],
    changes: &mut Vec<(usize, bool)>,
) -> Result<usize> {
    let mut operations = 0_usize;
    for run in items.chunk_by(|left, right| {
        left.plan.target.slide == right.plan.target.slide
            && left.plan.format.is_some() == right.plan.format.is_some()
            && left.plan.notes == right.plan.notes
    }) {
        let duplicate = if run.iter().any(|item| item.plan.original_text.is_some()) {
            let mut targets = HashSet::new();
            run.iter().any(|item| !targets.insert(item.plan.key()))
        } else {
            false
        };
        if duplicate {
            for item in run {
                stage_one(edit, &item.plan)?;
                operations = operations.saturating_add(1);
            }
            continue;
        }
        if run[0].plan.notes {
            operations = operations.saturating_add(stage_notes(
                edit,
                run[0].plan.target.slide,
                run,
                changes,
            )?);
            continue;
        }
        let kind = if run[0].plan.format.is_some() {
            BatchKind::Format
        } else if run.iter().any(|item| item.plan.target.bounds.is_some()) {
            BatchKind::Properties
        } else {
            BatchKind::Text
        };
        operations = operations.saturating_add(stage_batch(
            edit,
            run[0].plan.target.slide,
            kind,
            run,
            changes,
        )?);
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
    changes: &mut Vec<(usize, bool)>,
) -> Result<usize> {
    // A one-item run must use the same native singular operation as the
    // ordinary selector writer.  The chunked source route performs a second
    // reservation/admission pass and can retain a different patch envelope
    // even though it describes the same shape write.  That made a coalesced
    // edit reject a budget obtained from the direct proof.  Multi-item runs
    // still use the bounded chunk API so the coalescer keeps its one-pass
    // publication cost; reported story plans stay on their per-edit path
    // below because their native result is needed for diagnostics.
    if items.len() == 1 && !items[0].plan.reported {
        stage_one(edit, &items[0].plan)?;
        return Ok(usize::from(items[0].plan.changed));
    }
    if kind == BatchKind::Format {
        let mut edits = Vec::new();
        edits.try_reserve_exact(items.len()).map_err(allocation)?;
        for item in items {
            let format =
                item.plan.format.as_ref().ok_or_else(|| {
                    Error::new(Code::InternalError, "PPTX format batch lost a plan")
                })?;
            edits.push(TextAssignment::new(
                format.target,
                format.values.properties(),
            ));
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
        if !items.iter().any(|item| item.plan.reported) {
            return match edit {
                NativeEdit::Source(edit) => edit
                    .set_text_property_chunks(Position::new(slide), &chunks)
                    .map_err(native_error),
                NativeEdit::Eager(edit) => edit
                    .set_text_property_chunks(slide, &chunks)
                    .map_err(native_error),
            };
        }
        // The leaf reports each edit's change (record 0348).
        let flags = match edit {
            NativeEdit::Source(edit) => edit
                .set_text_property_chunks_per_edit(Position::new(slide), &chunks)
                .map_err(native_error)?,
            NativeEdit::Eager(edit) => edit
                .set_text_property_chunks_per_edit(slide, &chunks)
                .map_err(native_error)?,
        };
        if flags.len() != items.len() {
            return Err(Error::new(
                Code::InternalError,
                "PPTX format batch lost an edit's change",
            )
            .with_format(FileFormat::Pptx));
        }
        changes
            .try_reserve(items.iter().filter(|item| item.plan.reported).count())
            .map_err(allocation)?;
        changes.extend(
            items
                .iter()
                .zip(&flags)
                .filter(|(item, _)| item.plan.reported)
                .map(|(item, changed)| (item.index, *changed)),
        );
        return Ok(flags.iter().filter(|changed| **changed).count());
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

/// One notes page's formatting plans, through the eager editor's notes
/// text verb in 256-edit chunks (record 0346), staged in one pass over the
/// notes part (record 0348).
fn stage_notes(
    edit: &mut NativeEdit,
    slide: usize,
    items: &[Item],
    changes: &mut Vec<(usize, bool)>,
) -> Result<usize> {
    let NativeEdit::Eager(edit) = edit else {
        return Err(Error::new(
            Code::InternalError,
            "PPTX notes formatting joined a source-backed run",
        )
        .with_format(FileFormat::Pptx));
    };
    let mut edits = Vec::new();
    edits.try_reserve_exact(items.len()).map_err(allocation)?;
    for item in items {
        let format = item
            .plan
            .format
            .as_ref()
            .ok_or_else(|| Error::new(Code::InternalError, "PPTX notes batch lost a plan"))?;
        edits.push(TextAssignment::new(
            format.target,
            format.values.properties(),
        ));
    }
    let mut chunks = Vec::new();
    chunks
        .try_reserve_exact(
            edits
                .len()
                .div_ceil(litchi_pptx::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS),
        )
        .map_err(allocation)?;
    chunks.extend(edits.chunks(litchi_pptx::opened::MAX_SHAPE_PROPERTIES_REPLACEMENTS));
    // The leaf reports each edit's change (record 0348).
    let flags = edit
        .set_notes_text_property_chunks_per_edit(slide, &chunks)
        .map_err(native_error)?;
    if flags.len() != items.len() {
        return Err(Error::new(
            Code::InternalError,
            "PPTX notes batch lost an edit's change",
        )
        .with_format(FileFormat::Pptx));
    }
    changes.try_reserve(items.len()).map_err(allocation)?;
    changes.extend(
        items
            .iter()
            .zip(&flags)
            .filter(|(item, _)| item.plan.reported)
            .map(|(item, changed)| (item.index, *changed)),
    );
    Ok(flags.iter().filter(|changed| **changed).count())
}

fn stage_one(edit: &mut NativeEdit, plan: &Plan) -> Result<()> {
    if plan.notes {
        let Some(format) = &plan.format else {
            return Err(Error::new(
                Code::InternalError,
                "PPTX notes plan lost its format",
            ));
        };
        let NativeEdit::Eager(edit) = edit else {
            return Err(Error::new(
                Code::InternalError,
                "PPTX notes formatting joined a source-backed run",
            ));
        };
        let edits = [TextAssignment::new(
            format.target,
            format.values.properties(),
        )];
        return edit
            .set_notes_text_properties(plan.target.slide, &edits)
            .map(|_| ())
            .map_err(native_error);
    }
    if let Some(format) = &plan.format {
        let edits = [TextAssignment::new(
            format.target,
            format.values.properties(),
        )];
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
                Commit::new(
                    commit_after,
                    source,
                    target,
                    Proof::Source(patch),
                    operations,
                ),
            ))
        },
        NativeEdit::Eager(edit) => {
            let (target, patch) = edit.commit().map_err(native_error)?.into_parts();
            let changed = !patch.is_noop();
            let target = Editable::Eager(Arc::new(target));
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
                Commit::new(
                    commit_after,
                    source,
                    target,
                    Proof::Eager(patch),
                    operations,
                ),
            ))
        },
    }
}

fn replay_prefix(
    base: &Pptx,
    source: &Editable,
    items: &[Item],
    remaining_snapshot: usize,
) -> Result<()> {
    let mut edit = new_native_edit_with_snapshot_budget(base, source, remaining_snapshot)?;
    for item in items {
        stage_one(&mut edit, &item.plan).map_err(|error| error.with_item(item.index))?;
    }
    match edit {
        NativeEdit::Source(edit) => edit.commit().map_err(native_error).map(|_| ()),
        NativeEdit::Eager(edit) => edit.commit().map_err(native_error).map(|_| ()),
    }
}

fn locate_failure(base: &Pptx, items: &[Item], error: Error, remaining_snapshot: usize) -> Error {
    let Some(first) = items.first() else {
        return error;
    };
    if matches!(
        error.code(),
        Code::IoError | Code::Cancelled | Code::InternalError | Code::DocumentProtected
    ) {
        return error.with_item(first.index);
    }
    let source = match if base.signature() == crate::format::signature::Signature::Signed {
        base.editable_owner_for_write()
    } else {
        base.editable_owner()
    } {
        Ok(source) => source,
        Err(error) => return error.with_item(first.index),
    };
    // Usually failure occurred while staging a particular native item.
    // Replay once in wire order to recover that exact index, rather than
    // replaying every prefix (quadratic native calls on a late failure).
    let mut failure = match replay_prefix(base, &source, items, remaining_snapshot) {
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
        match replay_prefix(base, &source, &items[..middle], remaining_snapshot) {
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
