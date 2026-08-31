//! Exact-source transactions for persisted Keynote slide-table sort order.
//!
//! This owner deliberately edits only field 44 of one canonical
//! `TST.TableModelArchive`.  Physical row reordering, selected-row execution,
//! legacy type-6000 models, and native object identity remain compatibility
//! host concerns.  The public surface contains only semantic values and
//! checked slide/table selectors.  Field 45 (the native tracker) and every
//! other model field are intentionally opaque and are never rewritten.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "the package boundary redacts lower-layer failure details"
)]

use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{RawMessage, SnappyStream};
use litchi_iwa_protos::table_sort_order_codec as codec;
use thiserror::Error;

use super::slide_table_core as core;
use super::{Package, PayloadLimitKind, ReadError, SemanticLimitKind};
use crate::SlideSelector;
use crate::slide::table::TableSelector;
use crate::slide::table::sort::{ColumnIndex, Direction, Order, Rule, Scope};

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;

/// Finite resource categories enforced by a slide-table sort transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableSortLimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    PayloadObjects,
    PayloadMessages,
    References,
    WireFields,
    WireNesting,
    WireWork,
    Allocations,
    Retained,
    Scratch,
    Components,
}

impl fmt::Display for SlideTableSortLimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::References => "references",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::Components => "components",
        })
    }
}

/// Content-free semantic path for a slide-table sort operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableSortPath {
    Package,
    Table { slide: Position, table: Position },
}

impl fmt::Display for SlideTableSortPath {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => f.write_str("package"),
            Self::Table { slide, table } => {
                write!(f, "slide {} table {}", slide.get(), table.get())
            },
        }
    }
}

/// Failure from a Keynote slide-table sort read or transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideTableSortError {
    #[error("this Keynote source does not support persisted slide-table sort edits")]
    UnsupportedSource,
    #[error("the requested Keynote slide-table sort graph is outside the supported scope")]
    UnsupportedDependency,
    #[error("the requested Keynote slide-table sort topology is unsupported")]
    UnsupportedTopology,
    #[error("the Keynote slide-table sort selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no table at position {position:?}")]
    TablePositionNotFound { position: Position },
    #[error("the selected Keynote slide-table is locked")]
    Locked,
    #[error("the selected Keynote slide-table sort source is invalid")]
    InvalidSource,
    #[error(
        "Keynote slide-table sort {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SlideTableSortLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote slide-table sort transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote slide-table sort failed semantic verification")]
    Verification,
    #[error("the Keynote slide-table sort patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable sort value staged against an immutable package snapshot.
pub struct SlideTableSortEdit<'a> {
    source: &'a Package,
    selection: SortSelection,
    after: Option<Order>,
}

impl fmt::Debug for SlideTableSortEdit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideTableSortEdit")
            .field("path", &self.selection.path())
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableSortEdit<'_> {
    #[must_use]
    pub fn path(&self) -> SlideTableSortPath {
        self.selection.path()
    }
    #[must_use]
    pub fn order(&self) -> Option<&Order> {
        self.after.as_ref()
    }
    #[must_use]
    pub fn before(&self) -> Option<&Order> {
        self.selection.before.as_ref()
    }
    #[must_use]
    pub fn set(mut self, order: Order) -> Self {
        self.after = Some(order);
        self
    }
    #[must_use]
    pub fn clear(mut self) -> Self {
        self.after = None;
        self
    }
    #[must_use]
    pub fn reset(self) -> Self {
        self.clear()
    }
    pub fn commit(self) -> Result<SlideTableSortCommit, SlideTableSortError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible slide-table sort patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideTableSortPatch {
    artifacts: ExactArtifacts,
    selection: SortSelection,
    before: Option<Order>,
    after: Option<Order>,
}

impl fmt::Debug for SlideTableSortPatch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SlideTableSortPatch")
            .field("path", &self.selection.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableSortPatch {
    #[must_use]
    pub fn before(&self) -> Option<&Order> {
        self.before.as_ref()
    }
    #[must_use]
    pub fn after(&self) -> Option<&Order> {
        self.after.as_ref()
    }
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.selection.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Diagnostics for one persisted slide-table sort publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTableSortDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideTableSortDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }
    const fn published() -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews: 0,
            full_reparse_performed: true,
        }
    }
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one slide-table sort transaction.
#[must_use = "a Keynote slide-table sort commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTableSortCommit {
    package: Package,
    patch: SlideTableSortPatch,
    diagnostics: SlideTableSortDiagnostics,
}

impl SlideTableSortCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }
    #[must_use]
    pub const fn patch(&self) -> &SlideTableSortPatch {
        &self.patch
    }
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTableSortDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq, Eq)]
struct SortSelection {
    target: core::Target,
    before: Option<Order>,
    column_count: u32,
    budget: core::Budget,
}

impl fmt::Debug for SortSelection {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("SortSelection")
            .field("path", &self.path())
            .field("before", &self.before)
            .field("locked", &self.target.locked)
            .finish_non_exhaustive()
    }
}

impl SortSelection {
    const fn path(&self) -> SlideTableSortPath {
        SlideTableSortPath::Table {
            slide: self.target.slide_position,
            table: self.target.table_position,
        }
    }
}

impl Package {
    /// Read the persisted sort configuration of one slide table.
    pub fn slide_table_sort_order<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Option<Order>, SlideTableSortError> {
        let mut budget = core::Budget::new(self).map_err(map_graph_error)?;
        Ok(select_table(self, slide.into(), table.into(), &mut budget)?.before)
    }

    /// Begin an immutable exact edit of one persisted slide-table sort.
    pub fn edit_slide_table_sort_order<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<SlideTableSortEdit<'_>, SlideTableSortError> {
        let mut budget = core::Budget::new(self).map_err(map_graph_error)?;
        let selection = select_table(self, slide.into(), table.into(), &mut budget)?;
        Ok(SlideTableSortEdit {
            source: self,
            after: selection.before.clone(),
            selection,
        })
    }

    /// Apply an exact-source checked reversible slide-table sort patch.
    pub fn apply_slide_table_sort_order(
        &self,
        patch: &SlideTableSortPatch,
    ) -> Result<SlideTableSortCommit, SlideTableSortError> {
        let catalog = core::physical_catalog(self).map_err(map_graph_error)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(SlideTableSortError::PatchConflict);
        }
        let mut budget = core::Budget::new(self).map_err(map_graph_error)?;
        let current = select_table(
            self,
            SlideSelector::position(patch.selection.target.slide_position),
            TableSelector::position(patch.selection.target.table_position),
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(SlideTableSortError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideTableSortCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideTableSortDiagnostics::unchanged(),
            });
        }
        reopen_patch(self, patch, budget)
    }
}

fn commit_edit(
    source: &Package,
    selection: &SortSelection,
    after: Option<Order>,
) -> Result<SlideTableSortCommit, SlideTableSortError> {
    let mut budget = selection.budget;
    if selection.before == after {
        let bytes = shared_source_artifact(source, &mut budget)?;
        return Ok(SlideTableSortCommit {
            package: source.snapshot(),
            patch: SlideTableSortPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before: selection.before.clone(),
                after,
            },
            diagnostics: SlideTableSortDiagnostics::unchanged(),
        });
    }
    if !core::physical_catalog(source)
        .map_err(map_graph_error)?
        .source_is_exact()
    {
        return Err(SlideTableSortError::UnsupportedSource);
    }
    if selection.target.locked {
        return Err(SlideTableSortError::Locked);
    }
    let candidate = rewrite_sort(source, selection, after.clone(), &mut budget)?;
    let selected = select_table(
        &candidate,
        SlideSelector::position(selection.target.slide_position),
        TableSelector::position(selection.target.table_position),
        &mut budget,
    )?;
    if !same_selection(&selected, selection) || selected.before != after {
        return Err(SlideTableSortError::Verification);
    }
    core::verify_locality(source, &candidate, &selection.target, false, &mut budget)
        .map_err(map_graph_error)?;
    let target = shared_source_artifact(&candidate, &mut budget)?;
    let source_artifact = shared_source_artifact(source, &mut budget)?;
    Ok(SlideTableSortCommit {
        package: candidate,
        patch: SlideTableSortPatch {
            artifacts: ExactArtifacts::new(source_artifact, target),
            selection: selection.clone(),
            before: selection.before.clone(),
            after,
        },
        diagnostics: SlideTableSortDiagnostics::published(),
    })
}

fn reopen_patch(
    source: &Package,
    patch: &SlideTableSortPatch,
    mut budget: core::Budget,
) -> Result<SlideTableSortCommit, SlideTableSortError> {
    let target_source = patch.artifacts.target();
    let candidate = parse_candidate(Arc::clone(&target_source), source, &mut budget)?;
    let selected = select_table(
        &candidate,
        SlideSelector::position(patch.selection.target.slide_position),
        TableSelector::position(patch.selection.target.table_position),
        &mut budget,
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(SlideTableSortError::Verification);
    }
    core::verify_locality(
        source,
        &candidate,
        &patch.selection.target,
        false,
        &mut budget,
    )
    .map_err(map_graph_error)?;
    Ok(SlideTableSortCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideTableSortDiagnostics::published(),
    })
}

fn select_table(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
    budget: &mut core::Budget,
) -> Result<SortSelection, SlideTableSortError> {
    let target = core::select_table(package, slide_selector, table_selector, budget)
        .map_err(map_graph_error)?;
    let payload = core::model_payload(package, &target).map_err(map_graph_error)?;
    let before = decode_sort(payload, package, budget)?;
    let column_count = target.columns;
    if let Some(order) = &before {
        validate_columns(order, column_count, budget)?;
    }
    Ok(SortSelection {
        target,
        before,
        column_count,
        budget: *budget,
    })
}

fn rewrite_sort(
    source: &Package,
    selection: &SortSelection,
    after: Option<Order>,
    budget: &mut core::Budget,
) -> Result<Package, SlideTableSortError> {
    if let Some(order) = &after {
        validate_columns(order, selection.column_count, budget)?;
    }
    let catalog = core::physical_catalog(source).map_err(map_graph_error)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.target.model.component.as_ref())
        .ok_or(SlideTableSortError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideTableSortError::UnsupportedSource);
    }
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut archive =
        core::component_archive(source, selection.target.model.component.as_ref(), budget)
            .map_err(map_graph_error)?;
    let object = archive
        .object(selection.target.model.identifier)
        .ok_or(SlideTableSortError::InvalidSource)?;
    let original = object
        .messages
        .get(selection.target.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(SlideTableSortError::InvalidSource)?
        .data
        .as_slice();
    let current = decode_sort(original, source, budget)?;
    if current != selection.before {
        return Err(SlideTableSortError::InvalidSource);
    }
    let residual = budget.residual(source).map_err(map_graph_error)?;
    let options = codec_options(
        original,
        residual,
        budget.remaining_allocations().map_err(map_graph_error)?,
        budget.remaining_references().map_err(map_graph_error)?,
    );
    let prepared = codec::prepare_table_model_sort_order_rewrite(
        original,
        order_to_snapshot(after.as_ref(), budget)?,
        options,
    )
    .map_err(map_codec_error)?;
    budget
        .sort_codec_report(prepared.prepare_report())
        .map_err(map_graph_error)?;
    let requirements = prepared.execution_requirements();
    budget
        .sort_rewrite_requirements(requirements)
        .map_err(map_graph_error)?;
    let execution_limits = requirements.exact_limits();
    let output = prepared
        .execute(execution_limits)
        .map_err(map_codec_error)?;
    let report = output.report();
    if report.input_bytes() != original.len()
        || report.output_bytes() != requirements.output_bytes
        || report.fields() != requirements.fields
        || report.work_bytes() != requirements.work_bytes
        || report.max_depth() != requirements.max_depth
        || report.rules() != requirements.rules
        || report.allocations() != requirements.allocations
        || report.retained_bytes() != requirements.retained_bytes
        || report.scratch_bytes() != requirements.scratch_bytes
    {
        return Err(SlideTableSortError::Verification);
    }
    let rewritten = output.into_bytes();
    let verified = decode_sort(&rewritten, source, budget)?;
    if verified != after {
        return Err(SlideTableSortError::Verification);
    }
    archive
        .object_mut(selection.target.model.identifier)
        .ok_or(SlideTableSortError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.target.model_message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_len = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_len).map_err(map_core_error)?;
    budget
        .output(
            encoded_len
                .checked_add(compressed_bound)
                .ok_or(core::Error::InvalidSource)
                .map_err(map_graph_error)?,
        )
        .map_err(map_graph_error)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    if bytes.len() != encoded_len {
        return Err(SlideTableSortError::Verification);
    }
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(SlideTableSortError::Verification);
    }
    let edits = [EntryEdit::new(
        selection.target.model.component.as_ref(),
        &compressed,
    )];
    budget
        .preflight_reassembly(catalog, compressed.len(), 0)
        .map_err(map_graph_error)?;
    let prepared_reassembly = catalog
        .prepare_reassembly(&edits, physical_limits)
        .map_err(map_archive_error)?;
    let requirements = prepared_reassembly.execution_requirements();
    budget.reassembly(requirements).map_err(map_graph_error)?;
    let output = prepared_reassembly
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    parse_candidate(output.into(), source, budget)
}

fn order_to_snapshot(
    order: Option<&Order>,
    budget: &mut core::Budget,
) -> Result<Option<codec::SortOrderSnapshot>, SlideTableSortError> {
    order
        .map(|value| {
            let scope = match value.scope() {
                Scope::EntireTable => codec::SortScope::EntireTable,
                Scope::SelectedRows => codec::SortScope::SelectedRows,
            };
            let rules = value.rules().iter().map(|rule| {
                codec::SortRule::new(
                    rule.column().native_value(),
                    match rule.direction() {
                        Direction::Ascending => codec::SortDirection::Ascending,
                        Direction::Descending => codec::SortDirection::Descending,
                    },
                )
            });
            budget
                .allocations(value.rules().len())
                .map_err(map_graph_error)?;
            budget
                .retained(
                    value
                        .rules()
                        .len()
                        .checked_mul(size_of::<codec::SortRule>())
                        .ok_or(core::Error::InvalidSource)
                        .map_err(map_graph_error)?,
                )
                .map_err(map_graph_error)?;
            codec::SortOrderSnapshot::new(scope, rules)
                .map_err(|_| SlideTableSortError::InvalidSource)
        })
        .transpose()
}

fn snapshot_to_order(
    snapshot: Option<codec::SortOrderSnapshot>,
    budget: &mut core::Budget,
) -> Result<Option<Order>, SlideTableSortError> {
    snapshot
        .map(|value| {
            let scope = match value.scope() {
                codec::SortScope::EntireTable => Scope::EntireTable,
                codec::SortScope::SelectedRows => Scope::SelectedRows,
            };
            budget
                .allocations(value.rules().len())
                .map_err(map_graph_error)?;
            budget
                .retained(
                    value
                        .rules()
                        .len()
                        .checked_mul(size_of::<Rule>())
                        .ok_or(core::Error::InvalidSource)
                        .map_err(map_graph_error)?,
                )
                .map_err(map_graph_error)?;
            let mut rules = Vec::new();
            rules.try_reserve_exact(value.rules().len()).map_err(|_| {
                SlideTableSortError::Allocation {
                    amount: value.rules().len(),
                }
            })?;
            for rule in value.rules() {
                let column = ColumnIndex::from_native(rule.column())
                    .map_err(|_| SlideTableSortError::InvalidSource)?;
                let direction = match rule.direction() {
                    codec::SortDirection::Ascending => Direction::Ascending,
                    codec::SortDirection::Descending => Direction::Descending,
                };
                rules.push(Rule::new(column, direction));
            }
            Order::with_scope(scope, rules).map_err(|_| SlideTableSortError::InvalidSource)
        })
        .transpose()
}

fn decode_sort(
    payload: &[u8],
    package: &Package,
    budget: &mut core::Budget,
) -> Result<Option<Order>, SlideTableSortError> {
    let limits = budget.residual(package).map_err(map_graph_error)?;
    let options = codec_options(
        payload,
        limits,
        budget.remaining_allocations().map_err(map_graph_error)?,
        budget.remaining_references().map_err(map_graph_error)?,
    );
    let (snapshot, report) = codec::decode_table_model_sort_order_with_report(payload, options)
        .map_err(map_codec_error)?;
    budget.sort_codec_report(report).map_err(map_graph_error)?;
    snapshot_to_order(snapshot, budget)
}

fn codec_options(
    payload: &[u8],
    limits: WireLimits,
    max_allocations: usize,
    max_rules: usize,
) -> codec::DecodeOptions {
    codec::DecodeOptions::new(
        limits.max_input_bytes().min(payload.len()),
        limits.max_output_bytes(),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        payload.len().min(max_rules),
        usize::try_from(u32::MAX).unwrap_or(usize::MAX),
    )
    .with_max_allocations(max_allocations)
}

fn validate_columns(
    order: &Order,
    count: u32,
    budget: &mut core::Budget,
) -> Result<(), SlideTableSortError> {
    budget.work(order.rules().len()).map_err(map_graph_error)?;
    budget
        .references(order.rules().len())
        .map_err(map_graph_error)?;
    order.rules().iter().try_for_each(|rule| {
        if rule.column().native_value() >= count {
            Err(SlideTableSortError::InvalidSource)
        } else {
            Ok(())
        }
    })
}

fn same_selection(a: &SortSelection, b: &SortSelection) -> bool {
    core::same_target(&a.target, &b.target)
}

fn shared_source_artifact(
    package: &Package,
    budget: &mut core::Budget,
) -> Result<Arc<[u8]>, SlideTableSortError> {
    let source = core::physical_catalog(package)
        .map_err(map_graph_error)?
        .shared_source();
    budget.artifact(source.len(), 0).map_err(map_graph_error)?;
    Ok(source)
}

fn parse_candidate(
    source: Arc<[u8]>,
    original: &Package,
    budget: &mut core::Budget,
) -> Result<Package, SlideTableSortError> {
    budget
        .preflight_candidate(original, source.len())
        .map_err(map_graph_error)?;
    budget
        .preflight_semantic_scan(original)
        .map_err(map_graph_error)?;
    let candidate = Package::from_source_with_options(source, original.state.options)
        .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    Ok(candidate)
}

fn map_read_error(error: ReadError) -> SlideTableSortError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableSortError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => SlideTableSortLimitKind::References,
                _ => SlideTableSortLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableSortError::LimitExceeded {
            kind: match kind {
                PayloadLimitKind::Bytes => SlideTableSortLimitKind::InputBytes,
                PayloadLimitKind::Fields => SlideTableSortLimitKind::WireFields,
                PayloadLimitKind::Nesting => SlideTableSortLimitKind::WireNesting,
                PayloadLimitKind::Work => SlideTableSortLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideTableSortError::Allocation { amount },
        ReadError::Archive(error) => map_archive_error(error),
        _ => SlideTableSortError::InvalidSource,
    }
}
fn map_codec_error(error: codec::DecodeError) -> SlideTableSortError {
    if let Some(amount) = error.allocation_amount() {
        return SlideTableSortError::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return SlideTableSortError::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        codec::DecodeLimit::InputBytes { observed, maximum } => {
            (SlideTableSortLimitKind::InputBytes, observed, maximum)
        },
        codec::DecodeLimit::OutputBytes { observed, maximum } => {
            (SlideTableSortLimitKind::OutputBytes, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            (SlideTableSortLimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::WorkBytes { observed, maximum } => {
            (SlideTableSortLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => (
            SlideTableSortLimitKind::WireNesting,
            observed as usize,
            maximum as usize,
        ),
        codec::DecodeLimit::Rules { observed, maximum } => {
            (SlideTableSortLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Columns { observed, maximum } => {
            (SlideTableSortLimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Allocations { observed, maximum } => {
            (SlideTableSortLimitKind::Allocations, observed, maximum)
        },
        codec::DecodeLimit::RetainedBytes { observed, maximum } => {
            (SlideTableSortLimitKind::Retained, observed, maximum)
        },
        codec::DecodeLimit::ScratchBytes { observed, maximum } => {
            (SlideTableSortLimitKind::Scratch, observed, maximum)
        },
        _ => return SlideTableSortError::InvalidSource,
    };
    SlideTableSortError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}
fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideTableSortError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableSortError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SlideTableSortLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => SlideTableSortLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => SlideTableSortLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    SlideTableSortLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => SlideTableSortLimitKind::TotalBytes,
                _ => SlideTableSortLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideTableSortError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => SlideTableSortError::InvalidSource,
    }
}

fn map_graph_error(error: core::Error) -> SlideTableSortError {
    match error {
        core::Error::UnsupportedSource => SlideTableSortError::UnsupportedSource,
        core::Error::UnsupportedDependency => SlideTableSortError::UnsupportedDependency,
        core::Error::UnsupportedTopology => SlideTableSortError::UnsupportedTopology,
        core::Error::AmbiguousSelector => SlideTableSortError::AmbiguousSelector,
        core::Error::EmptySlideName => SlideTableSortError::EmptySlideName,
        core::Error::SlideNameNotFound => SlideTableSortError::SlideNameNotFound,
        core::Error::SlidePositionNotFound(position) => {
            SlideTableSortError::SlidePositionNotFound { position }
        },
        core::Error::TablePositionNotFound(position) => {
            SlideTableSortError::TablePositionNotFound { position }
        },
        core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableSortError::LimitExceeded {
            kind: map_limit_kind(kind),
            observed,
            maximum,
        },
        core::Error::Allocation(amount) => SlideTableSortError::Allocation { amount },
        core::Error::InvalidSource
        | core::Error::Read
        | core::Error::Wire
        | core::Error::Codec
        | core::Error::Archive
        | core::Error::Verification => SlideTableSortError::InvalidSource,
    }
}

fn map_limit_kind(kind: core::LimitKind) -> SlideTableSortLimitKind {
    match kind {
        core::LimitKind::InputBytes => SlideTableSortLimitKind::InputBytes,
        core::LimitKind::OutputBytes => SlideTableSortLimitKind::OutputBytes,
        core::LimitKind::Entries => SlideTableSortLimitKind::Entries,
        core::LimitKind::EntryBytes => SlideTableSortLimitKind::EntryBytes,
        core::LimitKind::TotalBytes => SlideTableSortLimitKind::TotalBytes,
        core::LimitKind::PayloadObjects => SlideTableSortLimitKind::PayloadObjects,
        core::LimitKind::PayloadMessages => SlideTableSortLimitKind::PayloadMessages,
        core::LimitKind::References => SlideTableSortLimitKind::References,
        core::LimitKind::WireFields => SlideTableSortLimitKind::WireFields,
        core::LimitKind::WireNesting => SlideTableSortLimitKind::WireNesting,
        core::LimitKind::WireWork => SlideTableSortLimitKind::WireWork,
        core::LimitKind::Allocations => SlideTableSortLimitKind::Allocations,
        core::LimitKind::Retained => SlideTableSortLimitKind::Retained,
        core::LimitKind::Scratch => SlideTableSortLimitKind::Scratch,
        core::LimitKind::Components => SlideTableSortLimitKind::Components,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SlideTableSortError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableSortError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => SlideTableSortLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideTableSortLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => SlideTableSortLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => SlideTableSortLimitKind::EntryBytes,
                _ => SlideTableSortLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTableSortError::Allocation { amount: requested }
        },
        _ => SlideTableSortError::InvalidSource,
    }
}
