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
use thiserror::Error as ThisError;

use super::slide_table_core as core;
use super::{Package, PayloadLimitKind, ReadError, SemanticLimitKind};
use crate::SlideSelector;
use crate::slide::table::TableSelector;
use crate::slide::table::sort::{ColumnIndex, Direction, Order, Rule, Scope};

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;

/// Finite resource categories enforced by a slide-table sort transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
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

impl fmt::Display for LimitKind {
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
pub enum Path {
    Package,
    Table { slide: Position, table: Position },
}

impl fmt::Display for Path {
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
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    #[error("the selected Keynote slide-table sort source is invalid")]
    InvalidSource,
    #[error(
        "Keynote slide-table sort {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    Limit {
        kind: LimitKind,
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

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
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
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::UnsupportedSource { .. }
            | Self::UnsupportedDependency { .. }
            | Self::UnsupportedTopology { .. } => RefusalKind::UnsupportedContent,
            Self::AmbiguousSelector { .. } | Self::EmptySlideName { .. } => {
                RefusalKind::InvalidValue
            },
            Self::SlideNameNotFound { .. }
            | Self::SlidePositionNotFound { .. }
            | Self::TablePositionNotFound { .. } => RefusalKind::NotFound,
            Self::Locked { .. } => RefusalKind::Protected,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// One mutable sort value staged against an immutable package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    selection: SortSelection,
    after: Option<Order>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Edit")
            .field("path", &self.selection.path())
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    #[must_use]
    pub fn path(&self) -> Path {
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
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible slide-table sort patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: ExactArtifacts,
    selection: SortSelection,
    before: Option<Order>,
    after: Option<Order>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch")
            .field("path", &self.selection.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
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
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
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
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
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
    const fn path(&self) -> Path {
        Path::Table {
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
    ) -> Result<Option<Order>, Error> {
        let mut budget = core::Budget::new(self).map_err(map_graph_error)?;
        Ok(select_table(self, slide.into(), table.into(), &mut budget)?.before)
    }

    /// Begin an immutable exact edit of one persisted slide-table sort.
    pub fn edit_slide_table_sort_order<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Edit<'_>, Error> {
        let mut budget = core::Budget::new(self).map_err(map_graph_error)?;
        let selection = select_table(self, slide.into(), table.into(), &mut budget)?;
        Ok(Edit {
            source: self,
            after: selection.before.clone(),
            selection,
        })
    }

    /// Apply an exact-source checked reversible slide-table sort patch.
    pub fn apply_slide_table_sort_order(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = core::physical_catalog(self).map_err(map_graph_error)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(Error::PatchConflict);
        }
        let mut budget = core::Budget::new(self).map_err(map_graph_error)?;
        let current = select_table(
            self,
            SlideSelector::position(patch.selection.target.slide_position),
            TableSelector::position(patch.selection.target.table_position),
            &mut budget,
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        reopen_patch(self, patch, budget)
    }
}

fn commit_edit(
    source: &Package,
    selection: &SortSelection,
    after: Option<Order>,
) -> Result<Commit, Error> {
    let mut budget = selection.budget;
    if selection.before == after {
        let bytes = shared_source_artifact(source, &mut budget)?;
        return Ok(Commit {
            package: source.snapshot(),
            patch: Patch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before: selection.before.clone(),
                after,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    if !core::physical_catalog(source)
        .map_err(map_graph_error)?
        .source_is_exact()
    {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    if selection.target.locked {
        return Err(Error::Refused(Refusal::Locked));
    }
    let candidate = rewrite_sort(source, selection, after.clone(), &mut budget)?;
    let selected = select_table(
        &candidate,
        SlideSelector::position(selection.target.slide_position),
        TableSelector::position(selection.target.table_position),
        &mut budget,
    )?;
    if !same_selection(&selected, selection) || selected.before != after {
        return Err(Error::Verification);
    }
    core::verify_locality(source, &candidate, &selection.target, false, &mut budget)
        .map_err(map_graph_error)?;
    let target = shared_source_artifact(&candidate, &mut budget)?;
    let source_artifact = shared_source_artifact(source, &mut budget)?;
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: ExactArtifacts::new(source_artifact, target),
            selection: selection.clone(),
            before: selection.before.clone(),
            after,
        },
        diagnostics: Diagnostics::published(),
    })
}

fn reopen_patch(
    source: &Package,
    patch: &Patch,
    mut budget: core::Budget,
) -> Result<Commit, Error> {
    let target_source = patch.artifacts.target();
    let candidate = parse_candidate(Arc::clone(&target_source), source, &mut budget)?;
    let selected = select_table(
        &candidate,
        SlideSelector::position(patch.selection.target.slide_position),
        TableSelector::position(patch.selection.target.table_position),
        &mut budget,
    )?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(Error::Verification);
    }
    core::verify_locality(
        source,
        &candidate,
        &patch.selection.target,
        false,
        &mut budget,
    )
    .map_err(map_graph_error)?;
    Ok(Commit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: Diagnostics::published(),
    })
}

fn select_table(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
    budget: &mut core::Budget,
) -> Result<SortSelection, Error> {
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
) -> Result<Package, Error> {
    if let Some(order) = &after {
        validate_columns(order, selection.column_count, budget)?;
    }
    let catalog = core::physical_catalog(source).map_err(map_graph_error)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.target.model.component.as_ref())
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let physical_limits = source.state.options.archive();
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut archive =
        core::component_archive(source, selection.target.model.component.as_ref(), budget)
            .map_err(map_graph_error)?;
    let object = archive
        .object(selection.target.model.identifier)
        .ok_or(Error::InvalidSource)?;
    let original = object
        .messages
        .get(selection.target.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(Error::InvalidSource)?
        .data
        .as_slice();
    let current = decode_sort(original, source, budget)?;
    if current != selection.before {
        return Err(Error::InvalidSource);
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
        return Err(Error::Verification);
    }
    let rewritten = output.into_bytes();
    let verified = decode_sort(&rewritten, source, budget)?;
    if verified != after {
        return Err(Error::Verification);
    }
    archive
        .object_mut(selection.target.model.identifier)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            selection.target.model_message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_len = archive
        .encoded_len(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_len, snappy_limits).map_err(map_core_error)?;
    budget
        .output(
            encoded_len
                .checked_add(compressed_bound)
                .ok_or(core::Error::InvalidSource)
                .map_err(map_graph_error)?,
        )
        .map_err(map_graph_error)?;
    let bytes = archive.to_bytes(archive_limits).map_err(map_core_error)?;
    if bytes.len() != encoded_len {
        return Err(Error::Verification);
    }
    let compressed = SnappyStream::compress(&bytes, snappy_limits).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(Error::Verification);
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
) -> Result<Option<codec::SortOrderSnapshot>, Error> {
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
            codec::SortOrderSnapshot::new(scope, rules).map_err(|_| Error::InvalidSource)
        })
        .transpose()
}

fn snapshot_to_order(
    snapshot: Option<codec::SortOrderSnapshot>,
    budget: &mut core::Budget,
) -> Result<Option<Order>, Error> {
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
            rules
                .try_reserve_exact(value.rules().len())
                .map_err(|_| Error::Allocation {
                    amount: value.rules().len(),
                })?;
            for rule in value.rules() {
                let column =
                    ColumnIndex::from_native(rule.column()).map_err(|_| Error::InvalidSource)?;
                let direction = match rule.direction() {
                    codec::SortDirection::Ascending => Direction::Ascending,
                    codec::SortDirection::Descending => Direction::Descending,
                };
                rules.push(Rule::new(column, direction));
            }
            Order::with_scope(scope, rules).map_err(|_| Error::InvalidSource)
        })
        .transpose()
}

fn decode_sort(
    payload: &[u8],
    package: &Package,
    budget: &mut core::Budget,
) -> Result<Option<Order>, Error> {
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

fn validate_columns(order: &Order, count: u32, budget: &mut core::Budget) -> Result<(), Error> {
    budget.work(order.rules().len()).map_err(map_graph_error)?;
    budget
        .references(order.rules().len())
        .map_err(map_graph_error)?;
    order.rules().iter().try_for_each(|rule| {
        if rule.column().native_value() >= count {
            Err(Error::InvalidSource)
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
) -> Result<Arc<[u8]>, Error> {
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
) -> Result<Package, Error> {
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

fn map_read_error(error: ReadError) -> Error {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                SemanticLimitKind::References => LimitKind::References,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => Error::Limit {
            kind: match kind {
                PayloadLimitKind::Bytes => LimitKind::InputBytes,
                PayloadLimitKind::Fields => LimitKind::WireFields,
                PayloadLimitKind::Nesting => LimitKind::WireNesting,
                PayloadLimitKind::Work => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => Error::Allocation { amount },
        ReadError::Archive(error) => map_archive_error(error),
        _ => Error::InvalidSource,
    }
}
fn map_codec_error(error: codec::DecodeError) -> Error {
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation { amount };
    }
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        codec::DecodeLimit::InputBytes { observed, maximum } => {
            (LimitKind::InputBytes, observed, maximum)
        },
        codec::DecodeLimit::OutputBytes { observed, maximum } => {
            (LimitKind::OutputBytes, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            (LimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::WorkBytes { observed, maximum } => {
            (LimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => {
            (LimitKind::WireNesting, observed as usize, maximum as usize)
        },
        codec::DecodeLimit::Rules { observed, maximum } => (LimitKind::WireWork, observed, maximum),
        codec::DecodeLimit::Columns { observed, maximum } => {
            (LimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Allocations { observed, maximum } => {
            (LimitKind::Allocations, observed, maximum)
        },
        codec::DecodeLimit::RetainedBytes { observed, maximum } => {
            (LimitKind::Retained, observed, maximum)
        },
        codec::DecodeLimit::ScratchBytes { observed, maximum } => {
            (LimitKind::Scratch, observed, maximum)
        },
        _ => return Error::InvalidSource,
    };
    Error::Limit {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}
fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => LimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => LimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => LimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalBytes,
                _ => LimitKind::WireWork,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => Error::InvalidSource,
    }
}

fn map_graph_error(error: core::Error) -> Error {
    match error {
        core::Error::UnsupportedSource => Error::Refused(Refusal::UnsupportedSource),
        core::Error::UnsupportedDependency => Error::Refused(Refusal::UnsupportedDependency),
        core::Error::UnsupportedTopology => Error::Refused(Refusal::UnsupportedTopology),
        core::Error::AmbiguousSelector => Error::Refused(Refusal::AmbiguousSelector),
        core::Error::EmptySlideName => Error::Refused(Refusal::EmptySlideName),
        core::Error::SlideNameNotFound => Error::Refused(Refusal::SlideNameNotFound),
        core::Error::SlidePositionNotFound(position) => {
            Error::Refused(Refusal::SlidePositionNotFound { position })
        },
        core::Error::TablePositionNotFound(position) => {
            Error::Refused(Refusal::TablePositionNotFound { position })
        },
        core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: map_limit_kind(kind),
            observed,
            maximum,
        },
        core::Error::Allocation(amount) => Error::Allocation { amount },
        core::Error::InvalidSource
        | core::Error::Read
        | core::Error::Wire
        | core::Error::Codec
        | core::Error::Archive
        | core::Error::Verification => Error::InvalidSource,
    }
}

fn map_limit_kind(kind: core::LimitKind) -> LimitKind {
    match kind {
        core::LimitKind::InputBytes => LimitKind::InputBytes,
        core::LimitKind::OutputBytes => LimitKind::OutputBytes,
        core::LimitKind::Entries => LimitKind::Entries,
        core::LimitKind::EntryBytes => LimitKind::EntryBytes,
        core::LimitKind::TotalBytes => LimitKind::TotalBytes,
        core::LimitKind::PayloadObjects => LimitKind::PayloadObjects,
        core::LimitKind::PayloadMessages => LimitKind::PayloadMessages,
        core::LimitKind::References => LimitKind::References,
        core::LimitKind::WireFields => LimitKind::WireFields,
        core::LimitKind::WireNesting => LimitKind::WireNesting,
        core::LimitKind::WireWork => LimitKind::WireWork,
        core::LimitKind::Allocations => LimitKind::Allocations,
        core::LimitKind::Retained => LimitKind::Retained,
        core::LimitKind::Scratch => LimitKind::Scratch,
        core::LimitKind::Components => LimitKind::Components,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => LimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::PayloadMessages,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::EntryBytes,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            Error::Allocation { amount: requested }
        },
        _ => Error::InvalidSource,
    }
}
