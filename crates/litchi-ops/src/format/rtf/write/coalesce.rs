//! Bounded direct-set coalescing for the RTF body.
//!
//! The native RTF editor accepts ordered batches.  This module admits only
//! exact body paragraph/run paths and keeps a source-backed positional index
//! for the lifetime of one pending transaction.  Formatting does not change
//! either ordinal axis, while text replacement closes a run for the affected
//! paragraph because native encoding may split or remove run fragments.

use std::mem::size_of;
use std::ops::Range;

use super::{
    Commit, Rtf, paragraph_properties, patch_snapshot_bytes, run_format_patch, run_text,
    staged_limit,
};
use crate::error::{Code, Error, Result};
use crate::props::Props;
use crate::target::TargetArg;
use crate::transaction::{Error as TransactionError, LimitKind, SemanticLimits};
use litchi_core::selector::{Path, Predicate, Step};
use litchi_core::{FileFormat, Position};
use litchi_rtf::edit::{
    EditOptions, ParagraphAlignmentUpdate, ParagraphLayout, ParagraphLayoutPatch,
    ParagraphLayoutUpdate, ParagraphTextReplacement, RunFormatPatch, RunFormatUpdate,
    RunTextReplacement, SemanticLimits as NativeLimits,
};
use litchi_rtf::text::{Break, Inline};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Family {
    Paragraph,
    RunText,
    RunFormat,
    /// Text of retained destination stories (cells, header/footer
    /// paragraphs, notes, comments). These never shift a positional axis.
    Destination,
    Noop,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum Target {
    Paragraph(Position),
    Run {
        paragraph: Position,
        run: Position,
    },
    /// An exact destination path, before native resolution.
    DestinationPath(super::destination::DirectPath),
    /// A resolved native destination story.
    Destination(super::destination::Target),
}

impl Target {
    fn key(self) -> (usize, usize) {
        match self {
            Self::Paragraph(paragraph) => (paragraph.get(), usize::MAX),
            Self::Run { paragraph, run } => (paragraph.get(), run.get()),
            // Destination stories do not share the body's ordinal axes.
            Self::DestinationPath(_) | Self::Destination(_) => (usize::MAX, usize::MAX),
        }
    }

    fn paragraph(self) -> usize {
        self.key().0
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
struct LayoutIntent {
    space_before: Option<i32>,
    space_after: Option<i32>,
    left_indent: Option<i32>,
    right_indent: Option<i32>,
    first_line_indent: Option<i32>,
    keep_together: Option<bool>,
    keep_with_next: Option<bool>,
    page_break_before: Option<bool>,
}

impl LayoutIntent {
    fn changed(self, current: ParagraphLayout) -> bool {
        self.space_before
            .is_some_and(|value| value != current.space_before())
            || self
                .space_after
                .is_some_and(|value| value != current.space_after())
            || self
                .left_indent
                .is_some_and(|value| value != current.left_indent())
            || self
                .right_indent
                .is_some_and(|value| value != current.right_indent())
            || self
                .first_line_indent
                .is_some_and(|value| value != current.first_line_indent())
            || self
                .keep_together
                .is_some_and(|value| value != current.keep_together())
            || self
                .keep_with_next
                .is_some_and(|value| value != current.keep_with_next())
            || self
                .page_break_before
                .is_some_and(|value| value != current.page_break_before())
    }
}

#[derive(Debug, Clone)]
pub(crate) struct Plan {
    target: Target,
    family: Family,
    text: Option<String>,
    alignment: Option<litchi_rtf::Alignment>,
    layout: Option<ParagraphLayoutPatch>,
    format: RunFormatPatch,
    staged_bytes: usize,
    changed: bool,
    structural_text: bool,
    stable_run_text: bool,
    native_operations: usize,
}

impl Plan {
    fn key(&self) -> (usize, usize) {
        self.target.key()
    }

    fn paragraph(&self) -> usize {
        self.target.paragraph()
    }
}

#[derive(Debug, Clone)]
struct ParagraphMeta {
    text: Range<usize>,
    alignment: litchi_rtf::Alignment,
    layout: ParagraphLayout,
    runs: Range<usize>,
}

#[derive(Debug, Clone)]
struct RunMeta {
    text: Range<usize>,
    format: RunFormatPatch,
}

#[derive(Debug)]
struct SourceIndex {
    paragraphs: Box<[ParagraphMeta]>,
    runs: Box<[RunMeta]>,
    /// Native note indices of the footnotes and of the endnotes, in order,
    /// so `footnote[N]`/`endnote[N]` resolve in constant time.
    notes: [Box<[usize]>; 2],
}

impl SourceIndex {
    /// Admits the index `build` would make under `limits`, without building
    /// it: the same charges, checked in the same order (record 0287 builds
    /// the index only when a plan first reads it).
    fn admit(file: &Rtf, limits: &SemanticLimits) -> Result<()> {
        let paragraph_charge = file
            .0
            .package
            .paragraph_count()
            .checked_mul(size_of::<ParagraphMeta>())
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        if paragraph_charge > limits.max_staged_bytes() {
            return Err(staged_limit(paragraph_charge, limits.max_staged_bytes()));
        }
        let run_charge = file
            .0
            .package
            .body()
            .runs()
            .count()
            .checked_mul(size_of::<RunMeta>())
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        let cache_charge = paragraph_charge
            .checked_add(run_charge)
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        if cache_charge > limits.max_staged_bytes() {
            return Err(staged_limit(cache_charge, limits.max_staged_bytes()));
        }
        let note_charge = cache_charge
            .checked_add(
                file.0
                    .package
                    .notes()
                    .len()
                    .saturating_mul(size_of::<usize>()),
            )
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        if note_charge > limits.max_staged_bytes() {
            return Err(staged_limit(note_charge, limits.max_staged_bytes()));
        }
        Ok(())
    }

    fn build(file: &Rtf, limits: &SemanticLimits) -> Result<Self> {
        let paragraph_count = file.0.package.paragraph_count();
        let paragraph_charge = paragraph_count
            .checked_mul(size_of::<ParagraphMeta>())
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        if paragraph_charge > limits.max_staged_bytes() {
            return Err(staged_limit(paragraph_charge, limits.max_staged_bytes()));
        }

        let mut paragraphs = Vec::new();
        paragraphs
            .try_reserve_exact(paragraph_count)
            .map_err(|source| TransactionError::Allocation { source })?;
        let run_count = file.0.package.body().runs().count();
        let run_charge = run_count
            .checked_mul(size_of::<RunMeta>())
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        let cache_charge = paragraph_charge
            .checked_add(run_charge)
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        if cache_charge > limits.max_staged_bytes() {
            return Err(staged_limit(cache_charge, limits.max_staged_bytes()));
        }
        let mut runs = Vec::new();
        runs.try_reserve_exact(run_count)
            .map_err(|source| TransactionError::Allocation { source })?;
        let mut body_offset = 0usize;
        for paragraph in file.0.package.body().paragraphs() {
            let text_start = body_offset;
            let text_end = text_start.saturating_add(paragraph.len());
            let run_start = runs.len();
            for inline in paragraph.inlines() {
                match inline {
                    Inline::Text(run) => {
                        let run_start_offset = body_offset;
                        let run_end_offset = run_start_offset.saturating_add(run.text().len());
                        let run_format = run.format();
                        runs.push(RunMeta {
                            text: run_start_offset..run_end_offset,
                            format: RunFormatPatch::new()
                                .with_bold(run_format.bold())
                                .with_italic(run_format.italic())
                                .with_underline(run_format.underline())
                                .with_size(run_format.size())
                                .with_strike(run_format.strike())
                                .with_double_strike(run_format.double_strike())
                                .with_hidden(run_format.hidden())
                                .with_all_caps(run_format.all_caps())
                                .with_small_caps(run_format.small_caps())
                                .with_outline(run_format.outline()),
                        });
                        body_offset = run_end_offset;
                    },
                    Inline::Break(Break::Line) => {
                        body_offset = body_offset.saturating_add(1);
                    },
                    Inline::Break(_) => {},
                    _ => {},
                }
            }
            let run_end = runs.len();
            paragraphs.push(ParagraphMeta {
                text: text_start..text_end,
                alignment: paragraph.format().alignment(),
                layout: paragraph.format().layout(),
                runs: run_start..run_end,
            });
            body_offset = text_end.saturating_add(1);
        }
        let note_count = file.0.package.notes().len();
        let note_charge = cache_charge
            .checked_add(note_count.saturating_mul(size_of::<usize>()))
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
        if note_charge > limits.max_staged_bytes() {
            return Err(staged_limit(note_charge, limits.max_staged_bytes()));
        }
        let mut footnotes = Vec::new();
        let mut endnotes = Vec::new();
        for (index, note) in file.0.package.notes().iter().enumerate() {
            let list = if note.is_footnote {
                &mut footnotes
            } else {
                &mut endnotes
            };
            list.try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;
            list.push(index);
        }
        Ok(Self {
            paragraphs: paragraphs.into_boxed_slice(),
            runs: runs.into_boxed_slice(),
            notes: [footnotes.into_boxed_slice(), endnotes.into_boxed_slice()],
        })
    }

    /// The native index of the `ordinal`-th footnote or endnote.
    fn note(&self, footnotes: bool, ordinal: usize) -> Option<usize> {
        self.notes[usize::from(!footnotes)].get(ordinal).copied()
    }

    fn paragraph(&self, position: Position) -> Result<&ParagraphMeta> {
        self.paragraphs.get(position.get()).ok_or_else(|| {
            Error::new(Code::PathNotFound, "the RTF paragraph path does not exist")
                .with_format(FileFormat::Rtf)
        })
    }

    fn run(&self, paragraph: Position, run: Position) -> Result<&RunMeta> {
        let paragraph_meta = self.paragraph(paragraph)?;
        let absolute = paragraph_meta
            .runs
            .start
            .checked_add(run.get())
            .filter(|position| *position < paragraph_meta.runs.end)
            .ok_or_else(|| {
                Error::new(Code::PathNotFound, "the RTF run path does not exist")
                    .with_format(FileFormat::Rtf)
            })?;
        self.runs.get(absolute).ok_or_else(|| {
            Error::new(Code::PathNotFound, "the RTF run path does not exist")
                .with_format(FileFormat::Rtf)
        })
    }
}

/// A run's source index, built on first use.
struct SourceRef<'a> {
    file: &'a Rtf,
    limits: &'a SemanticLimits,
    index: &'a std::sync::OnceLock<SourceIndex>,
}

impl SourceRef<'_> {
    fn get(&self) -> Result<&SourceIndex> {
        if let Some(index) = self.index.get() {
            return Ok(index);
        }
        let built = SourceIndex::build(self.file, self.limits)?;
        Ok(self.index.get_or_init(|| built))
    }
}

struct Item {
    plan: Plan,
    index: usize,
}

pub(crate) struct Run {
    base: Rtf,
    edit: litchi_rtf::edit::Edit,
    /// Built on the first plan that reads it; admitted when the run opens.
    source: std::sync::OnceLock<SourceIndex>,
    limits: SemanticLimits,
    items: Vec<Item>,
    family: Option<Family>,
    last_key: Option<(usize, usize)>,
    unsafe_run_paragraph: Option<usize>,
    structural_paragraph_text: bool,
    paragraph_text: bool,
    paragraph_alignment: bool,
    paragraph_layout: bool,
    staged_bytes: usize,
    native_operations: usize,
    /// Destination stories staged in this window, sorted for duplicate checks.
    destinations: Vec<super::destination::Target>,
}

/// Cheap syntactic admission used before constructing a source index.  It
/// keeps selectors, relative paths and structural aliases on the ordinary
/// resolver, where their full error and query semantics remain intact.
pub(crate) fn direct_candidate(target: &TargetArg<'_>, props: &Props<'_>) -> Result<bool> {
    let Some(path_arg) = target.path() else {
        return Ok(false);
    };
    if path_arg.raw {
        return Ok(false);
    }
    let path = path_arg.path();
    let Some(target) = direct_target(path) else {
        return Ok(false);
    };
    let _ = (target, props);
    Ok(true)
}

impl Run {
    pub(crate) fn new(base: &Rtf, limits: &SemanticLimits) -> Result<Option<Self>> {
        base.admit_edit()?;
        match SourceIndex::admit(base, limits) {
            Ok(()) => {},
            Err(error)
                if matches!(
                    error.leaf(),
                    crate::error::Source::Transaction(TransactionError::Limit {
                        kind: LimitKind::StagedBytes,
                        ..
                    })
                ) =>
            {
                // The bounded optimization index must not narrow scalar
                // mutation admission. The ordinary resolver walks the axis.
                return Ok(None);
            },
            Err(error) => return Err(error),
        }
        let native_limits = NativeLimits::builder()
            .max_operations(limits.max_operations())
            .map_err(Error::from)?
            .max_replacement_bytes(limits.max_staged_bytes())
            .map_err(Error::from)?
            .build()
            .map_err(Error::from)?;
        let edit = base
            .0
            .package
            .edit_with_options(EditOptions::new().semantic_limits(native_limits));
        Ok(Some(Self {
            base: base.clone(),
            edit,
            // Index storage has its own check against the staging bound;
            // it is not authored payload/action charge (nor a heap budget).
            staged_bytes: 0,
            source: std::sync::OnceLock::new(),
            limits: *limits,
            items: Vec::new(),
            destinations: Vec::new(),
            family: None,
            last_key: None,
            unsafe_run_paragraph: None,
            structural_paragraph_text: false,
            paragraph_text: false,
            paragraph_alignment: false,
            paragraph_layout: false,
            native_operations: 0,
        }))
    }

    pub(crate) fn plan(&self, target: &TargetArg<'_>, props: &Props<'_>) -> Result<Option<Plan>> {
        let Some(path_arg) = target.path() else {
            return Ok(None);
        };
        if path_arg.raw {
            return Ok(None);
        }
        let path = path_arg.path();
        let Some(target) = direct_target(path) else {
            return Ok(None);
        };
        if self.limits.max_expanded_targets() < 1 {
            return Err(TransactionError::Limit {
                kind: LimitKind::ExpandedTargets,
                observed: 1,
                maximum: self.limits.max_expanded_targets(),
            }
            .into());
        }
        make_plan(
            &self.base,
            &SourceRef {
                file: &self.base,
                limits: &self.limits,
                index: &self.source,
            },
            target,
            props,
            &self.limits,
            self.staged_bytes,
        )
        .map(Some)
    }

    pub(crate) fn requires_barrier(&self, plan: &Plan) -> bool {
        let Some(family) = self.family else {
            return false;
        };
        if let Target::Destination(target) = plan.target {
            // A destination window closes only on a family change, a full
            // window, or a second write to the same native story.
            return family != plan.family
                || self.items.len() >= self.limits.max_expanded_targets()
                || self
                    .native_operations
                    .saturating_add(plan.native_operations)
                    > self.limits.max_operations()
                || self.destinations.binary_search(&target).is_ok();
        }
        if self.items.len() >= self.limits.max_expanded_targets() {
            return true;
        }
        if family != plan.family {
            return true;
        }
        if self
            .native_operations
            .saturating_add(plan.native_operations)
            > self.limits.max_operations()
        {
            return true;
        }
        if self.last_key.is_some_and(|last| plan.key() <= last) {
            return true;
        }
        if self.structural_paragraph_text {
            return true;
        }
        if matches!(plan.family, Family::Paragraph)
            && (((self.paragraph_layout || self.paragraph_alignment) && plan.text.is_some())
                || (self.paragraph_text
                    && (plan.layout.is_some() || (plan.alignment.is_some() && plan.changed)))
                || (plan.structural_text
                    && self
                        .items
                        .iter()
                        .any(|item| item.plan.alignment.is_some() || item.plan.layout.is_some())))
        {
            return true;
        }
        if matches!(plan.target, Target::Run { .. })
            && self
                .unsafe_run_paragraph
                .is_some_and(|paragraph| paragraph == plan.paragraph())
        {
            return true;
        }
        matches!(plan.family, Family::RunText)
            && !plan.stable_run_text
            && self
                .unsafe_run_paragraph
                .is_some_and(|paragraph| paragraph == plan.paragraph())
    }

    pub(crate) fn first_index(&self) -> Option<usize> {
        self.items.first().map(|item| item.index)
    }

    pub(crate) fn stage(&mut self, plan: Plan, index: usize) -> Result<bool> {
        if self.items.len() >= self.limits.max_expanded_targets() {
            return Err(TransactionError::Limit {
                kind: LimitKind::ExpandedTargets,
                observed: self.items.len().checked_add(1).unwrap_or(usize::MAX),
                maximum: self.limits.max_expanded_targets(),
            }
            .into());
        }
        if self.requires_barrier(&plan) {
            return Err(Error::new(
                Code::InternalError,
                "RTF coalescer received a plan across a required barrier",
            )
            .with_format(FileFormat::Rtf));
        }
        let bytes = self
            .staged_bytes
            .checked_add(plan.staged_bytes)
            .ok_or_else(|| staged_limit(usize::MAX, self.limits.max_staged_bytes()))?;
        if bytes > self.limits.max_staged_bytes() {
            return Err(staged_limit(bytes, self.limits.max_staged_bytes()));
        }
        self.items
            .try_reserve(1)
            .map_err(|source| TransactionError::Allocation { source })?;
        let changed = plan.changed;
        if plan.structural_text {
            self.structural_paragraph_text = true;
        }
        if matches!(plan.family, Family::Paragraph) {
            self.paragraph_text |= plan.text.is_some();
            self.paragraph_alignment |= plan.alignment.is_some() && plan.changed;
            self.paragraph_layout |= plan.layout.is_some();
        }
        if matches!(plan.family, Family::RunText) && !plan.stable_run_text {
            self.unsafe_run_paragraph = Some(plan.paragraph());
        }
        if let Target::Destination(target) = plan.target {
            let slot = self
                .destinations
                .binary_search(&target)
                .unwrap_or_else(|slot| slot);
            self.destinations
                .try_reserve(1)
                .map_err(|source| TransactionError::Allocation { source })?;
            self.destinations.insert(slot, target);
        } else {
            self.last_key = Some(plan.key());
        }
        self.family = Some(plan.family);
        self.staged_bytes = bytes;
        self.native_operations = self
            .native_operations
            .saturating_add(plan.native_operations);
        self.items.push(Item { plan, index });
        Ok(changed)
    }

    pub(crate) fn commit(self) -> Result<(Rtf, Commit)> {
        let Self {
            base,
            mut edit,
            items,
            limits,
            ..
        } = self;
        let result = stage_native(&mut edit, &base, &items)
            .and_then(|()| edit.commit().map_err(Error::from));
        let native = match result {
            Ok(native) => native,
            Err(error) => return Err(locate_failure(&base, &items, error, &limits)),
        };
        let changed = native.diagnostics().changed();
        let operations = native.diagnostics().operation_count();
        let (package, patch) = native.into_parts();
        let snapshot_bytes = patch_snapshot_bytes(&patch);
        if snapshot_bytes > limits.max_snapshot_bytes() {
            return Err(snapshot_limit(snapshot_bytes, limits.max_snapshot_bytes()));
        }
        base.check_source()?;
        let before = base.clone();
        let after = if changed {
            base.with_package(package)
        } else {
            base.clone()
        };
        Ok((
            after.clone(),
            Commit {
                before,
                after,
                patch,
                operations,
                snapshot_bytes,
            },
        ))
    }
}

pub(super) fn direct_target(path: &Path<'_>) -> Option<Target> {
    if !path.is_absolute() {
        return None;
    }
    let steps = path.steps();
    if let Some(destination) = destination_path(steps) {
        return Some(Target::DestinationPath(destination));
    }
    if !(steps.len() == 2 || steps.len() == 3) {
        return None;
    }
    if !body_step(&steps[0]) {
        return None;
    }
    let paragraph = steps[1].index()?;
    if !name_is(&steps[1], "p") && !name_is(&steps[1], "paragraph") {
        return None;
    }
    if steps.len() == 2 {
        return Some(Target::Paragraph(paragraph));
    }
    if !name_is(&steps[2], "r") && !name_is(&steps[2], "run") {
        return None;
    }
    Some(Target::Run {
        paragraph,
        run: steps[2].index()?,
    })
}

/// An exact, purely positional destination-story path.
fn destination_path(steps: &[Step<'_>]) -> Option<super::destination::DirectPath> {
    use super::destination::DirectPath;
    let named = |step: &Step<'_>, names: &[&str]| names.iter().any(|name| name_is(step, name));
    match steps {
        [body, table, row, cell]
            if body_step(body)
                && named(table, &["tbl", "table"])
                && named(row, &["tr", "row"])
                && named(cell, &["tc", "cell"]) =>
        {
            Some(DirectPath::Cell(
                table.index()?.get(),
                row.index()?.get(),
                cell.index()?.get(),
            ))
        },
        [story] | [story, _] if named(story, &["header", "footer"]) => {
            let paragraph = match steps.get(1) {
                Some(step) if named(step, &["p", "paragraph"]) => Some(step.index()?.get()),
                Some(_) => return None,
                None => None,
            };
            Some(DirectPath::Story {
                headers: name_is(story, "header"),
                story: story.index()?.get(),
                paragraph,
            })
        },
        [note] if named(note, &["footnote", "endnote"]) => Some(DirectPath::Note {
            footnotes: name_is(note, "footnote"),
            ordinal: note.index()?.get(),
        }),
        [comments, comment] if bare(comments, "comments") && name_is(comment, "comment") => {
            Some(DirectPath::Comment(comment.index()?.get()))
        },
        _ => None,
    }
}

fn bare(step: &Step<'_>, name: &str) -> bool {
    name_is(step, name) && step.predicates().is_empty()
}

pub(super) fn body_step(step: &Step<'_>) -> bool {
    name_is(step, "body")
        && match step.predicates() {
            [] => true,
            [Predicate::Index(position)] => position.get() == 0,
            _ => false,
        }
}

fn name_is(step: &Step<'_>, name: &str) -> bool {
    step.name().eq_ignore_ascii_case(name)
}

fn make_plan(
    file: &Rtf,
    source: &SourceRef<'_>,
    target: Target,
    props: &Props<'_>,
    limits: &SemanticLimits,
    staged_base: usize,
) -> Result<Plan> {
    match target {
        Target::DestinationPath(path) => {
            let resolved = match path {
                super::destination::DirectPath::Note { footnotes, ordinal } => {
                    super::destination::Target::Note(
                        source.get()?.note(footnotes, ordinal).ok_or_else(|| {
                            Error::new(Code::PathNotFound, "the RTF note does not exist")
                                .with_format(FileFormat::Rtf)
                        })?,
                    )
                },
                _ => super::destination::resolve(file, path)?,
            };
            make_plan(
                file,
                source,
                Target::Destination(resolved),
                props,
                limits,
                staged_base,
            )
        },
        Target::Destination(resolved) => {
            let text = super::destination::replacement_text(props)?;
            let current = super::destination::current_text(file, resolved);
            let per_target = size_of::<super::destination::Target>()
                .checked_add(text.len())
                .and_then(|bytes| bytes.checked_add(current.len()))
                .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
            admit_staged(staged_base, per_target, limits)?;
            Ok(Plan {
                target: Target::Destination(resolved),
                family: Family::Destination,
                text: Some(text.to_owned()),
                alignment: None,
                layout: None,
                format: RunFormatPatch::new(),
                staged_bytes: per_target,
                changed: current != text,
                structural_text: false,
                stable_run_text: false,
                native_operations: 1,
            })
        },
        Target::Paragraph(position) => {
            let meta = source.get()?.paragraph(position)?;
            let properties = paragraph_properties(props)?;
            let current_text = source_text(file, &meta.text);
            let layout_intent = layout_intent(props)?;
            let text_changed = properties.text.is_some_and(|text| current_text != text);
            let alignment_changed = properties
                .alignment
                .is_some_and(|alignment| alignment != meta.alignment);
            let layout_changed = properties
                .layout
                .is_some_and(|_| layout_intent.changed(meta.layout));
            let action_count = usize::from(properties.text.is_some())
                .saturating_add(usize::from(properties.alignment.is_some()))
                .saturating_add(usize::from(properties.layout.is_some()));
            if action_count > limits.max_operations() {
                return Err(TransactionError::Limit {
                    kind: LimitKind::Operations,
                    observed: action_count,
                    maximum: limits.max_operations(),
                }
                .into());
            }
            let per_target = properties
                .text
                .map_or(0, |text| {
                    size_of::<ParagraphTextReplacement>().saturating_add(text.len())
                })
                .checked_add(
                    usize::from(properties.alignment.is_some())
                        * size_of::<ParagraphAlignmentUpdate>(),
                )
                .and_then(|bytes| {
                    bytes.checked_add(
                        usize::from(properties.layout.is_some())
                            * size_of::<ParagraphLayoutUpdate>(),
                    )
                })
                .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
            let family = Family::Paragraph;
            admit_staged(staged_base, per_target, limits)?;
            Ok(Plan {
                target,
                family,
                text: properties.text.map(str::to_owned),
                alignment: properties.alignment,
                layout: properties.layout,
                format: RunFormatPatch::new(),
                staged_bytes: per_target,
                changed: text_changed || alignment_changed || layout_changed,
                structural_text: properties
                    .text
                    .is_some_and(|text| text.contains(['\n', '\r'])),
                stable_run_text: false,
                native_operations: action_count,
            })
        },
        Target::Run { paragraph, run } => {
            let meta = source.get()?.run(paragraph, run)?;
            let text = run_text(props)?;
            let format = run_format_patch(props)?;
            let current_text = source_text(file, &meta.text);
            let text_changed = text.is_some_and(|value| value != current_text);
            let format_changed = patch_changes(meta.format, format);
            let family = if text.is_some() {
                Family::RunText
            } else if format.is_empty() {
                Family::Noop
            } else {
                Family::RunFormat
            };
            let per_target = if text.is_some() {
                size_of::<RunTextReplacement>()
                    .checked_add(text.map_or(0, str::len))
                    .and_then(|bytes| bytes.checked_add(current_text.len()))
            } else if format.is_empty() {
                Some(0)
            } else {
                Some(size_of::<RunFormatUpdate>())
            }
            .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
            admit_staged(staged_base, per_target, limits)?;
            Ok(Plan {
                target: Target::Run { paragraph, run },
                family,
                text: text.map(str::to_owned),
                alignment: None,
                layout: None,
                format,
                staged_bytes: per_target,
                changed: text_changed || format_changed,
                structural_text: false,
                stable_run_text: text.is_some_and(stable_run_text),
                native_operations: usize::from(!format.is_empty() || text.is_some()),
            })
        },
    }
}

fn source_text<'a>(file: &'a Rtf, range: &Range<usize>) -> &'a str {
    file.0.package.text().get(range.clone()).unwrap_or("")
}

fn admit_staged(base: usize, incoming: usize, limits: &SemanticLimits) -> Result<usize> {
    let observed = base
        .checked_add(incoming)
        .ok_or_else(|| staged_limit(usize::MAX, limits.max_staged_bytes()))?;
    if observed > limits.max_staged_bytes() {
        return Err(staged_limit(observed, limits.max_staged_bytes()));
    }
    Ok(observed)
}

fn stable_run_text(text: &str) -> bool {
    !text.is_empty()
        && text.bytes().all(|byte| {
            (byte.is_ascii_graphic() && !matches!(byte, b'\\' | b'{' | b'}')) || byte == b' '
        })
}

fn layout_intent(props: &Props<'_>) -> Result<LayoutIntent> {
    let mut intent = LayoutIntent::default();
    if let Some(value) = props.get("spaceBefore") {
        intent.space_before = Some(super::paragraph_twips(value, "spaceBefore", true)?);
    }
    if let Some(value) = props.get("spaceAfter") {
        intent.space_after = Some(super::paragraph_twips(value, "spaceAfter", true)?);
    }
    if let Some(value) = props.get("indent") {
        intent.left_indent = Some(super::paragraph_twips(value, "indent", false)?);
    }
    if let Some(value) = props.get("rightIndent") {
        intent.right_indent = Some(super::paragraph_twips(value, "rightIndent", false)?);
    }
    if let Some(value) = props.get("firstLineIndent") {
        intent.first_line_indent = Some(super::paragraph_twips(value, "firstLineIndent", false)?);
    }
    if let Some(value) = props.get("keepNext") {
        intent.keep_with_next = Some(super::paragraph_bool(value, "keepNext")?);
    }
    if let Some(value) = props.get("keepLines") {
        intent.keep_together = Some(super::paragraph_bool(value, "keepLines")?);
    }
    if let Some(value) = props.get("pageBreakBefore") {
        intent.page_break_before = Some(super::paragraph_bool(value, "pageBreakBefore")?);
    }
    Ok(intent)
}

fn patch_changes(current: RunFormatPatch, requested: RunFormatPatch) -> bool {
    requested
        .bold()
        .is_some_and(|value| Some(value) != current.bold())
        || requested
            .italic()
            .is_some_and(|value| Some(value) != current.italic())
        || requested
            .underline()
            .is_some_and(|value| Some(value) != current.underline())
        || requested
            .size()
            .is_some_and(|value| Some(value) != current.size())
        || requested
            .strike()
            .is_some_and(|value| Some(value) != current.strike())
        || requested
            .double_strike()
            .is_some_and(|value| Some(value) != current.double_strike())
        || requested
            .hidden()
            .is_some_and(|value| Some(value) != current.hidden())
        || requested
            .all_caps()
            .is_some_and(|value| Some(value) != current.all_caps())
        || requested
            .small_caps()
            .is_some_and(|value| Some(value) != current.small_caps())
        || requested
            .outline()
            .is_some_and(|value| Some(value) != current.outline())
}

fn stage_native(edit: &mut litchi_rtf::edit::Edit, base: &Rtf, items: &[Item]) -> Result<()> {
    let Some(first) = items.first() else {
        return Ok(());
    };
    match first.plan.family {
        Family::Destination => {
            for item in items {
                let Plan {
                    target: Target::Destination(target),
                    text: Some(text),
                    ..
                } = &item.plan
                else {
                    return Err(Error::new(
                        Code::InternalError,
                        "mixed RTF destination plans",
                    ));
                };
                super::destination::stage(edit, base, *target, text)?;
            }
        },
        Family::Paragraph => {
            let mut texts = Vec::new();
            let mut alignments = Vec::new();
            let mut layouts = Vec::new();
            let text_count = items.iter().filter(|item| item.plan.text.is_some()).count();
            let alignment_count = items
                .iter()
                .filter(|item| item.plan.alignment.is_some())
                .count();
            let layout_count = items
                .iter()
                .filter(|item| item.plan.layout.is_some())
                .count();
            if text_count != 0 {
                texts
                    .try_reserve_exact(text_count)
                    .map_err(|source| TransactionError::Allocation { source })?;
            }
            if alignment_count != 0 {
                alignments
                    .try_reserve_exact(alignment_count)
                    .map_err(|source| TransactionError::Allocation { source })?;
            }
            if layout_count != 0 {
                layouts
                    .try_reserve_exact(layout_count)
                    .map_err(|source| TransactionError::Allocation { source })?;
            }
            for item in items {
                let plan = &item.plan;
                let Target::Paragraph(position) = plan.target else {
                    return Err(Error::new(
                        Code::InternalError,
                        "mixed RTF coalescer targets",
                    ));
                };
                if let Some(text) = &plan.text {
                    texts.push(ParagraphTextReplacement::new(position, text.clone()));
                }
                if let Some(alignment) = plan.alignment {
                    alignments.push(ParagraphAlignmentUpdate::new(position, alignment));
                }
                if let Some(layout) = plan.layout {
                    layouts.push(ParagraphLayoutUpdate::new(position, layout));
                }
            }
            if !texts.is_empty() {
                edit.replace_body_paragraph_texts(&texts)
                    .map_err(Error::from)?;
            }
            if !alignments.is_empty() {
                edit.set_body_paragraph_alignments(&alignments)
                    .map_err(Error::from)?;
            }
            if !layouts.is_empty() {
                edit.patch_body_paragraph_layouts(&layouts)
                    .map_err(Error::from)?;
            }
        },
        Family::RunText => {
            let mut replacements = Vec::new();
            replacements
                .try_reserve_exact(items.len())
                .map_err(|source| TransactionError::Allocation { source })?;
            for item in items {
                let Plan {
                    target: Target::Run { paragraph, run },
                    text: Some(text),
                    format,
                    ..
                } = &item.plan
                else {
                    return Err(Error::new(Code::InternalError, "mixed RTF run-text plans"));
                };
                let replacement =
                    RunTextReplacement::new(*paragraph, *run, text.clone()).with_format(*format);
                replacements.push(replacement);
            }
            if !replacements.is_empty() {
                edit.replace_body_run_texts(&replacements)
                    .map_err(Error::from)?;
            }
        },
        Family::RunFormat => {
            let mut updates = Vec::new();
            updates
                .try_reserve_exact(items.len())
                .map_err(|source| TransactionError::Allocation { source })?;
            for item in items {
                let Plan {
                    target: Target::Run { paragraph, run },
                    format,
                    ..
                } = &item.plan
                else {
                    return Err(Error::new(
                        Code::InternalError,
                        "mixed RTF run-format plans",
                    ));
                };
                updates.push(RunFormatUpdate::new(*paragraph, *run, *format));
            }
            if !updates.is_empty() {
                edit.patch_body_run_formats(&updates).map_err(Error::from)?;
            }
        },
        Family::Noop => {},
    }
    Ok(())
}

fn replay(base: &Rtf, items: &[Item], limits: &SemanticLimits) -> Result<Rtf> {
    let native_limits = NativeLimits::builder()
        .max_operations(limits.max_operations())
        .map_err(Error::from)?
        .max_replacement_bytes(limits.max_staged_bytes())
        .map_err(Error::from)?
        .build()
        .map_err(Error::from)?;
    let mut edit = base
        .0
        .package
        .edit_with_options(EditOptions::new().semantic_limits(native_limits));
    stage_native(&mut edit, base, items)?;
    let native = edit.commit().map_err(Error::from)?;
    let changed = native.diagnostics().changed();
    let (package, _) = native.into_parts();
    Ok(if changed {
        base.with_package(package)
    } else {
        base.clone()
    })
}

fn locate_failure(base: &Rtf, items: &[Item], error: Error, limits: &SemanticLimits) -> Error {
    let Some(first) = items.first() else {
        return error;
    };
    if items.len() == 1
        || matches!(
            error.code(),
            Code::IoError
                | Code::Cancelled
                | Code::LimitExceeded
                | Code::InternalError
                | Code::DocumentProtected
        )
    {
        return error.with_item(first.index);
    }
    let middle = items.len() / 2;
    match replay(base, &items[..middle], limits) {
        Err(error) => locate_failure(base, &items[..middle], error, limits),
        Ok(prefix) => match replay(&prefix, &items[middle..], limits) {
            Err(error) => locate_failure(&prefix, &items[middle..], error, limits),
            Ok(_) => error.with_item(first.index),
        },
    }
}

fn snapshot_limit(observed: usize, maximum: usize) -> Error {
    TransactionError::Limit {
        kind: LimitKind::SnapshotBytes,
        observed,
        maximum,
    }
    .into()
}
