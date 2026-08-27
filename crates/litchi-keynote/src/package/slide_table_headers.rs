//! Exact-source transactions for existing Keynote slide-table header settings.
//!
//! The owner is intentionally narrow: it resolves a table from the selected
//! slide's z-order, follows the rooted `TableInfo` edge to one canonical
//! type-6001 table model, and rewrites only the seven scalar header fields in
//! that model.  Native identifiers never cross the public API and all other
//! model fields, archive metadata, previews, and package members remain
//! source-authoritative.

#![allow(
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::wildcard_enum_match_arm,
    reason = "the focused package boundary redacts lower-layer failure details"
)]

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceScope, ArchiveReferenceVisitor, FieldType, RawMessage,
    SnappyStream,
};
use litchi_iwa_protos::{numbers_table_header_settings_codec as header_codec, table_info_codec};
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::SlideSelector;
use crate::slide::table::{TableSelector, headers::Settings};

const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const LEGACY_TABLE_MODEL_MESSAGE_TYPE: u32 = 6_000;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_Z_ORDER_FIELD: u32 = 42;
const TABLE_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_FIELD: u32 = 2;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const MAX_VARINT_BYTES: usize = 10;
const CATEGORY_GROUPING_FIELD: u32 = 81;
const GROUPING_FIELD: u32 = 83;
const PIVOT_FIELD: u32 = 85;
const CATEGORY_OWNER_FIELD: u32 = 86;
const CATEGORY_OWNER_REFERENCE_MESSAGE_TYPE: u32 = 6_372;
const GROUP_BY_MESSAGE_TYPE: u32 = 6_373;
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
const HEADER_NAME_MANAGER_MESSAGE_TYPE: u32 = 6_366;

/// Finite resource categories enforced by a slide-table header transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableHeaderLimitKind {
    /// Complete source package bytes inspected.
    InputBytes,
    /// Complete candidate package bytes produced.
    OutputBytes,
    /// Physical package entries.
    Entries,
    /// Bytes in one physical entry.
    EntryBytes,
    /// Aggregate physical-entry bytes.
    TotalBytes,
    /// Decoded native payload objects.
    PayloadObjects,
    /// Decoded native payload messages.
    PayloadMessages,
    /// Archive metadata items inspected.
    PayloadItems,
    /// Native object-reference edges inspected.
    References,
    /// Protobuf input bytes.
    WireBytes,
    /// Protobuf output bytes.
    WireOutputBytes,
    /// Protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Protobuf scan/rewrite work.
    WireWork,
    /// Fallible allocation events or requested allocation sizes.
    Allocations,
    /// Bytes retained by a focused transaction.
    Retained,
    /// Temporary scratch bytes.
    Scratch,
    /// Aggregate transaction work.
    TransactionWork,
}

impl fmt::Display for SlideTableHeaderLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::References => "references",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting depth",
            Self::WireWork => "wire work",
            Self::Allocations => "allocations",
            Self::Retained => "retained bytes",
            Self::Scratch => "scratch bytes",
            Self::TransactionWork => "transaction work",
        })
    }
}

/// Content-free semantic location associated with a slide-table operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableHeaderPath {
    /// The complete Keynote package.
    Package,
    /// One checked zero-based slide/table location.
    Table { slide: Position, table: Position },
}

impl fmt::Display for SlideTableHeaderPath {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("package"),
            Self::Table { slide, table } => {
                write!(formatter, "slide {} table {}", slide.get(), table.get())
            },
        }
    }
}

/// Content-free reason for rejecting requested header partitions.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableHeaderInvalidReason {
    /// Leading header rows plus trailing footer rows exceed the table rows.
    RowSectionsExceedTable {
        /// Requested leading header rows.
        header_rows: u8,
        /// Requested trailing footer rows.
        footer_rows: u8,
        /// Native table rows.
        table_rows: u32,
    },
    /// Leading header columns exceed the table columns.
    HeaderColumnsExceedTable {
        /// Requested leading header columns.
        header_columns: u8,
        /// Native table columns.
        table_columns: u32,
    },
}

impl fmt::Display for SlideTableHeaderInvalidReason {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::RowSectionsExceedTable {
                header_rows,
                footer_rows,
                table_rows,
            } => write!(
                formatter,
                "header rows {header_rows} plus footer rows {footer_rows} exceed {table_rows} table rows"
            ),
            Self::HeaderColumnsExceedTable {
                header_columns,
                table_columns,
            } => write!(
                formatter,
                "header columns {header_columns} exceed {table_columns} table columns"
            ),
        }
    }
}

/// Failure from a Keynote slide-table header read or transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum SlideTableHeaderError {
    /// The source has no exact physical package artifact.
    #[error("this Keynote source does not support physical slide-table header edits")]
    UnsupportedSource,
    /// The selected graph contains a dependency this scalar edit cannot keep fresh.
    #[error("the requested Keynote slide-table header graph has an unsupported dependency")]
    UnsupportedDependency,
    /// The rooted graph is not the one canonical topology owned by this adapter.
    #[error("the requested Keynote slide-table header topology is unsupported")]
    UnsupportedTopology,
    /// A slide name selector matched more than one slide.
    #[error("the Keynote slide-table header selector is ambiguous")]
    AmbiguousSelector,
    /// A slide name selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// No slide matched a name selector.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A positional slide selector was outside the show.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// A table selector was outside the selected slide's z-order tables.
    #[error("the selected Keynote slide has no table at position {position:?}")]
    TablePositionNotFound { position: Position },
    /// A changed edit targeted a locked table.
    #[error("the selected Keynote slide table is locked")]
    Locked,
    /// Requested settings violate native table dimensions.
    #[error("the requested Keynote slide-table header settings are invalid at {path}: {reason}")]
    InvalidSettings {
        /// Semantic table path.
        path: SlideTableHeaderPath,
        /// Content-free validation reason.
        reason: SlideTableHeaderInvalidReason,
    },
    /// The selected graph, archive metadata, or wire framing is malformed.
    #[error("the selected Keynote slide-table header source is invalid")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error(
        "Keynote slide-table headers {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category.
        kind: SlideTableHeaderLimitKind,
        /// Observed amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote slide-table header transaction")]
    Allocation { amount: usize },
    /// Candidate reopening or locality verification failed.
    #[error("the edited Keynote slide-table headers failed semantic verification")]
    Verification,
    /// The patch was created from another exact source artifact.
    #[error("the Keynote slide-table header patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct HeaderBudget {
    max_input: usize,
    max_output: usize,
    max_fields: usize,
    max_work: usize,
    max_nesting: usize,
    max_references: usize,
    max_allocations: usize,
    max_retained: usize,
    max_scratch: usize,
    input: usize,
    output: usize,
    fields: usize,
    work: usize,
    nesting: usize,
    references: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
}

impl HeaderBudget {
    fn new(package: &Package) -> Result<Self, SlideTableHeaderError> {
        let wire = package.wire_limits().map_err(map_wire_error)?;
        let source = usize::try_from(package.state.options.archive().max_input_bytes())
            .map_err(|_| SlideTableHeaderError::InvalidSource)?;
        let aggregate = source
            .checked_mul(4)
            .ok_or(SlideTableHeaderError::InvalidSource)?;
        Ok(Self {
            max_input: aggregate,
            max_output: aggregate,
            max_fields: wire.max_fields(),
            max_work: wire.max_rewrite_work(),
            max_nesting: wire.max_nesting(),
            max_references: package.semantic_limits().max_references(),
            max_allocations: aggregate,
            max_retained: aggregate,
            max_scratch: aggregate,
            input: 0,
            output: 0,
            fields: 0,
            work: 0,
            nesting: 0,
            references: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: SlideTableHeaderLimitKind,
    ) -> Result<(), SlideTableHeaderError> {
        let observed = current
            .checked_add(amount)
            .ok_or(SlideTableHeaderError::InvalidSource)?;
        if observed > maximum {
            return Err(SlideTableHeaderError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    fn input(&mut self, amount: usize) -> Result<(), SlideTableHeaderError> {
        Self::add(
            &mut self.input,
            amount,
            self.max_input,
            SlideTableHeaderLimitKind::InputBytes,
        )
    }

    fn output(&mut self, amount: usize) -> Result<(), SlideTableHeaderError> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            SlideTableHeaderLimitKind::OutputBytes,
        )
    }

    fn fields(&mut self, amount: usize) -> Result<(), SlideTableHeaderError> {
        Self::add(
            &mut self.fields,
            amount,
            self.max_fields,
            SlideTableHeaderLimitKind::WireFields,
        )
    }

    fn work(&mut self, amount: usize) -> Result<(), SlideTableHeaderError> {
        Self::add(
            &mut self.work,
            amount,
            self.max_work,
            SlideTableHeaderLimitKind::WireWork,
        )
    }

    fn references(&mut self, amount: usize) -> Result<(), SlideTableHeaderError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            SlideTableHeaderLimitKind::References,
        )
    }

    fn allocations(&mut self, amount: usize) -> Result<(), SlideTableHeaderError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            SlideTableHeaderLimitKind::Allocations,
        )
    }

    fn retained(&mut self, amount: usize) -> Result<(), SlideTableHeaderError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            SlideTableHeaderLimitKind::Retained,
        )
    }

    fn scratch(&mut self, amount: usize) -> Result<(), SlideTableHeaderError> {
        Self::add(
            &mut self.scratch,
            amount,
            self.max_scratch,
            SlideTableHeaderLimitKind::Scratch,
        )
    }

    fn physical(&mut self, amount: usize) -> Result<(), SlideTableHeaderError> {
        self.input(amount)?;
        self.work(amount)
    }

    fn codec_report(
        &mut self,
        report: header_codec::RewriteReport,
    ) -> Result<(), SlideTableHeaderError> {
        self.input(report.input_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.output(report.output_bytes())?;
        self.allocations(report.allocations())?;
        self.nesting = self.nesting.max(report.max_depth() as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideTableHeaderError::LimitExceeded {
                kind: SlideTableHeaderLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }

    fn residual(&self, package: &Package) -> Result<WireLimits, SlideTableHeaderError> {
        let base = package.wire_limits().map_err(map_wire_error)?;
        base.with_input_bytes(
            base.max_input_bytes()
                .min(self.max_input.saturating_sub(self.input).max(1)),
        )
        .and_then(|limits| {
            limits.with_fields(
                base.max_fields()
                    .min(self.max_fields.saturating_sub(self.fields).max(1)),
            )
        })
        .and_then(|limits| {
            limits.with_output_bytes(
                base.max_output_bytes()
                    .min(self.max_output.saturating_sub(self.output).max(1)),
            )
        })
        .and_then(|limits| {
            limits.with_rewrite_work(
                base.max_rewrite_work()
                    .min(self.max_work.saturating_sub(self.work).max(1)),
            )
        })
        .and_then(|limits| limits.with_nesting(base.max_nesting().min(self.max_nesting)))
        .map_err(map_wire_error)
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), SlideTableHeaderError> {
        self.output(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        self.work(requirements.output_bytes())
    }

    fn merge_usage(&mut self, other: Self) -> Result<(), SlideTableHeaderError> {
        self.input(other.input)?;
        self.output(other.output)?;
        self.fields(other.fields)?;
        self.work(other.work)?;
        self.references(other.references)?;
        self.allocations(other.allocations)?;
        self.retained(other.retained)?;
        self.scratch(other.scratch)?;
        self.nesting = self.nesting.max(other.nesting);
        if self.nesting > self.max_nesting {
            return Err(SlideTableHeaderError::LimitExceeded {
                kind: SlideTableHeaderLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        Ok(())
    }
}

/// Mutable seven-field header settings staged against one immutable package.
pub struct SlideTableHeaderEdit<'a> {
    source: &'a Package,
    selection: HeaderSelection,
    after: Settings,
}

impl fmt::Debug for SlideTableHeaderEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableHeaderEdit")
            .field("path", &self.selection.path())
            .field("before", &self.selection.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableHeaderEdit<'_> {
    /// Return the selected slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.selection.slide_position
    }

    /// Return the selected table position.
    #[must_use]
    pub const fn table_position(&self) -> Position {
        self.selection.table_position
    }

    /// Return the semantic path.
    #[must_use]
    pub const fn path(&self) -> SlideTableHeaderPath {
        self.selection.path()
    }

    /// Return the source settings.
    #[must_use]
    pub const fn before(&self) -> Settings {
        self.selection.before
    }

    /// Return the staged settings.
    #[must_use]
    pub const fn after(&self) -> Settings {
        self.after
    }

    /// Return the staged settings (an alias useful to generic edit code).
    #[must_use]
    pub const fn settings(&self) -> Settings {
        self.after
    }

    /// Replace the complete staged settings.
    #[must_use]
    pub fn set(mut self, settings: Settings) -> Self {
        self.after = settings;
        self
    }

    /// Validate and publish the staged settings atomically.
    pub fn commit(self) -> Result<SlideTableHeaderCommit, SlideTableHeaderError> {
        commit_edit(self.source, &self.selection, self.after)
    }
}

/// Exact-source checked reversible slide-table header patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideTableHeaderPatch {
    artifacts: ExactArtifacts,
    selection: HeaderSelection,
    before: Settings,
    after: Settings,
}

impl fmt::Debug for SlideTableHeaderPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableHeaderPatch")
            .field("path", &self.selection.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableHeaderPatch {
    /// Return source settings required by this patch.
    #[must_use]
    pub const fn before(&self) -> Settings {
        self.before
    }

    /// Return settings produced by this patch.
    #[must_use]
    pub const fn after(&self) -> Settings {
        self.after
    }

    /// Return the semantic path selected by this patch.
    #[must_use]
    pub const fn path(&self) -> SlideTableHeaderPath {
        self.selection.path()
    }

    /// Return the source artifact's diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target artifact's diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this is an exact byte-and-semantic no-op.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            selection: self.selection.clone(),
            before: self.after,
            after: self.before,
        }
    }
}

/// Compact publication diagnostics for one slide-table header transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTableHeaderDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideTableHeaderDiagnostics {
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

    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten IWA members.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of deleted root previews (always zero for this owner).
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the candidate was fully reopened.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one slide-table header transaction.
#[must_use = "a slide-table header commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTableHeaderCommit {
    package: Package,
    patch: SlideTableHeaderPatch,
    diagnostics: SlideTableHeaderDiagnostics,
}

impl SlideTableHeaderCommit {
    /// Borrow the fully verified package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this result and return the package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &SlideTableHeaderPatch {
        &self.patch
    }

    /// Borrow publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTableHeaderDiagnostics {
        &self.diagnostics
    }
}

#[derive(Clone, PartialEq, Eq)]
struct HeaderSelection {
    slide_position: Position,
    table_position: Position,
    slide_identifier: u64,
    table_info_identifier: u64,
    model_identifier: u64,
    slide_message_index: usize,
    table_info_message_index: usize,
    model_message_index: usize,
    component_name: Arc<str>,
    before: Settings,
    rows: u32,
    columns: u32,
    locked: bool,
    admission_budget: HeaderBudget,
}

impl fmt::Debug for HeaderSelection {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("HeaderSelection")
            .field("slide_position", &self.slide_position)
            .field("table_position", &self.table_position)
            .field("before", &self.before)
            .field("locked", &self.locked)
            .finish_non_exhaustive()
    }
}

impl HeaderSelection {
    const fn path(&self) -> SlideTableHeaderPath {
        SlideTableHeaderPath::Table {
            slide: self.slide_position,
            table: self.table_position,
        }
    }
}

impl Package {
    /// Read one rooted slide table's lossless header/footer settings.
    pub fn slide_table_header_settings<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Settings, SlideTableHeaderError> {
        Ok(select_table(self, slide.into(), table.into())?.before)
    }

    /// Start a selector-first immutable slide-table header edit.
    pub fn edit_slide_table_headers<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<SlideTableHeaderEdit<'_>, SlideTableHeaderError> {
        let selection = select_table(self, slide.into(), table.into())?;
        Ok(SlideTableHeaderEdit {
            source: self,
            after: selection.before,
            selection,
        })
    }

    /// Apply an exact-source checked reversible slide-table header patch.
    pub fn apply_slide_table_headers(
        &self,
        patch: &SlideTableHeaderPatch,
    ) -> Result<SlideTableHeaderCommit, SlideTableHeaderError> {
        let catalog = physical_catalog(self)?;
        if !patch.artifacts.authorizes_source(&catalog.shared_source()) {
            return Err(SlideTableHeaderError::PatchConflict);
        }
        let current = select_table(
            self,
            SlideSelector::position(patch.selection.slide_position),
            TableSelector::position(patch.selection.table_position),
        )?;
        if !same_selection(&current, &patch.selection) || current.before != patch.before {
            return Err(SlideTableHeaderError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(SlideTableHeaderCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: SlideTableHeaderDiagnostics::unchanged(),
            });
        }
        reopen_patch(self, patch, current.admission_budget)
    }
}

fn commit_edit(
    source: &Package,
    selection: &HeaderSelection,
    after: Settings,
) -> Result<SlideTableHeaderCommit, SlideTableHeaderError> {
    let catalog = physical_catalog(source)?;
    if selection.before == after {
        let bytes = catalog.shared_source();
        return Ok(SlideTableHeaderCommit {
            package: source.snapshot(),
            patch: SlideTableHeaderPatch {
                artifacts: ExactArtifacts::new(Arc::clone(&bytes), bytes),
                selection: selection.clone(),
                before: selection.before,
                after,
            },
            diagnostics: SlideTableHeaderDiagnostics::unchanged(),
        });
    }
    if !catalog.source_is_exact() {
        return Err(SlideTableHeaderError::UnsupportedSource);
    }
    if selection.locked {
        return Err(SlideTableHeaderError::Locked);
    }
    let mut budget = selection.admission_budget;
    validate_requested(after, selection.rows, selection.columns, selection.path())?;
    validate_dependencies(source, selection, selection.before, after, &mut budget)?;
    let candidate = rewrite_headers(source, selection, after, &mut budget)?;
    candidate.validate().map_err(map_read_error)?;
    let selected = select_table(
        &candidate,
        SlideSelector::position(selection.slide_position),
        TableSelector::position(selection.table_position),
    )?;
    budget.merge_usage(selected.admission_budget)?;
    if !same_selection(&selected, selection) || selected.before != after {
        return Err(SlideTableHeaderError::Verification);
    }
    verify_locality(source, &candidate, selection, &mut budget)?;
    let target = physical_catalog(&candidate)?.shared_source();
    Ok(SlideTableHeaderCommit {
        package: candidate,
        patch: SlideTableHeaderPatch {
            artifacts: ExactArtifacts::new(Arc::from(source.source_bytes()), target),
            selection: selection.clone(),
            before: selection.before,
            after,
        },
        diagnostics: SlideTableHeaderDiagnostics::published(),
    })
}

fn reopen_patch(
    source: &Package,
    patch: &SlideTableHeaderPatch,
    mut budget: HeaderBudget,
) -> Result<SlideTableHeaderCommit, SlideTableHeaderError> {
    budget.physical(patch.artifacts.target().len())?;
    let candidate =
        Package::from_source_with_options(patch.artifacts.target(), source.state.options)
            .map_err(map_read_error)?;
    candidate.validate().map_err(map_read_error)?;
    let selected = select_table(
        &candidate,
        SlideSelector::position(patch.selection.slide_position),
        TableSelector::position(patch.selection.table_position),
    )?;
    budget.merge_usage(selected.admission_budget)?;
    if !same_selection(&selected, &patch.selection) || selected.before != patch.after {
        return Err(SlideTableHeaderError::Verification);
    }
    verify_locality(source, &candidate, &patch.selection, &mut budget)?;
    Ok(SlideTableHeaderCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideTableHeaderDiagnostics::published(),
    })
}

fn select_table(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
) -> Result<HeaderSelection, SlideTableHeaderError> {
    let mut budget = HeaderBudget::new(package)?;
    let catalog = physical_catalog(package)?;
    budget.input(package.source_bytes().len())?;
    budget.allocations(catalog.package().len())?;
    let slide_position = resolve_slide_position(package, slide_selector)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(SlideTableHeaderError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (_slide_component_name, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let (slide_message_index, slide_payload) = unique_message(slide, SLIDE_MESSAGE_TYPE)?;
    let wire_limits = budget.residual(package)?;
    let owned = repeated_references(
        slide_payload,
        SLIDE_OWNED_DRAWABLES_FIELD,
        wire_limits,
        &mut budget,
    )?;
    let z_order =
        repeated_references(slide_payload, SLIDE_Z_ORDER_FIELD, wire_limits, &mut budget)?;
    reject_duplicates(&owned)?;
    reject_duplicates(&z_order)?;
    validate_slide_metadata(slide, slide_message_index, &owned, &z_order)?;

    let mut tables = Vec::new();
    budget.allocations(z_order.len())?;
    tables
        .try_reserve_exact(z_order.len())
        .map_err(|_| SlideTableHeaderError::Allocation {
            amount: z_order.len(),
        })?;
    for table_info_identifier in z_order {
        let Some((_owner_component, info_object)) =
            package.object_with_component(table_info_identifier)
        else {
            return Err(SlideTableHeaderError::InvalidSource);
        };
        if info_object
            .messages
            .iter()
            .all(|message| message.type_ != TABLE_INFO_MESSAGE_TYPE)
        {
            continue;
        }
        if owned
            .iter()
            .filter(|candidate| **candidate == table_info_identifier)
            .count()
            != 1
        {
            return Err(SlideTableHeaderError::InvalidSource);
        }
        let (table_info_message_index, info_payload) =
            unique_message(info_object, TABLE_INFO_MESSAGE_TYPE)?;
        let info = decode_table_info(info_payload, package, &mut budget)?;
        let parent = table_parent(info_payload, wire_limits)?;
        if parent != record.slide_identifier {
            return Err(SlideTableHeaderError::InvalidSource);
        }
        let model_identifier = info.table_model().identifier().get();
        validate_table_info_metadata(info_object, table_info_message_index, model_identifier)?;
        let (model_component, model) = package
            .object_with_component(model_identifier)
            .ok_or(SlideTableHeaderError::InvalidSource)?;
        // A model object containing a type-6000 alias is not a canonical
        // type-6001 model, even if it also contains a decodable payload.
        if model
            .messages
            .iter()
            .any(|message| message.type_ == LEGACY_TABLE_MODEL_MESSAGE_TYPE)
        {
            return Err(SlideTableHeaderError::UnsupportedTopology);
        }
        let (model_message_index, model_payload) = unique_message(model, TABLE_MODEL_MESSAGE_TYPE)?;
        let snapshot = decode_header(model_payload, package, &mut budget)?;
        let before = settings_from_snapshot(snapshot)?;
        validate_stored(before, snapshot.rows(), snapshot.columns())?;
        if snapshot.rows() == 0 || snapshot.columns() == 0 {
            return Err(SlideTableHeaderError::UnsupportedTopology);
        }
        tables.push((
            table_info_identifier,
            model_identifier,
            Arc::<str>::from(model_component),
            slide_message_index,
            table_info_message_index,
            model_message_index,
            before,
            snapshot.rows(),
            snapshot.columns(),
            info.locked().unwrap_or(false),
        ));
    }

    let table_position = table_selector.as_position();
    let (
        table_info_identifier,
        model_identifier,
        component_name,
        slide_message_index,
        table_info_message_index,
        model_message_index,
        before,
        rows,
        columns,
        locked,
    ) = tables.get(table_position.get()).cloned().ok_or(
        SlideTableHeaderError::TablePositionNotFound {
            position: table_position,
        },
    )?;
    ensure_unique_identity(package, table_info_identifier, &mut budget)?;
    ensure_unique_identity(package, model_identifier, &mut budget)?;
    ensure_unique_table_owner(
        package,
        record.slide_identifier,
        table_info_identifier,
        model_identifier,
        wire_limits,
        &mut budget,
    )?;
    validate_global_inbound_references(
        package,
        record.slide_identifier,
        slide_message_index,
        table_info_identifier,
        table_info_message_index,
        model_identifier,
        &mut budget,
    )?;
    Ok(HeaderSelection {
        slide_position,
        table_position,
        slide_identifier: record.slide_identifier,
        table_info_identifier,
        model_identifier,
        slide_message_index,
        table_info_message_index,
        model_message_index,
        component_name,
        before,
        rows,
        columns,
        locked,
        admission_budget: budget,
    })
}

fn rewrite_headers(
    source: &Package,
    selection: &HeaderSelection,
    after: Settings,
    budget: &mut HeaderBudget,
) -> Result<Package, SlideTableHeaderError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.component_name.as_ref())
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideTableHeaderError::UnsupportedSource);
    }
    let physical_limits = source.state.options.archive();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    budget.physical(entry.data().len())?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    budget.physical(stream.as_bytes().len())?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    let object = archive
        .object(selection.model_identifier)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    if object.archive_info.identifier != Some(selection.model_identifier) {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    validate_message_header(object, selection.model_message_index)?;
    let original = object
        .messages
        .get(selection.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(SlideTableHeaderError::InvalidSource)?
        .data
        .as_slice();
    let snapshot = decode_header(original, source, budget)?;
    if settings_from_snapshot(snapshot)? != selection.before {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    validate_requested(after, snapshot.rows(), snapshot.columns(), selection.path())?;
    let write = settings_to_write(after)?;
    let output_bound = original
        .len()
        .checked_add(
            7usize
                .checked_mul(MAX_VARINT_BYTES + 1)
                .ok_or(SlideTableHeaderError::InvalidSource)?,
        )
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let limits = budget
        .residual(source)?
        .with_output_bytes(
            budget
                .max_output
                .saturating_sub(budget.output)
                .min(output_bound)
                .max(1),
        )
        .map_err(map_wire_error)?;
    let options = header_codec::DecodeOptions::new(
        limits.max_input_bytes().min(original.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    )
    .with_max_output_bytes(limits.max_output_bytes())
    .with_max_allocations(
        budget
            .max_allocations
            .saturating_sub(budget.allocations)
            .max(1),
    )
    .with_max_retained_bytes(budget.max_retained.saturating_sub(budget.retained).max(1))
    .with_max_scratch_bytes(budget.max_scratch.saturating_sub(budget.scratch).max(1));
    // `decode_header` above is the mandatory strict source pass.  The
    // prepared codec deliberately preserves unknown spans while sizing and
    // executing, so never let it be the first admission check for framing.
    let prepared = header_codec::prepare_table_header_settings_rewrite(original, write, options)
        .map_err(map_header_codec_error)?;
    let prepare_report = prepared.prepare_report();
    let requirements = prepared.execution_requirements();
    if prepare_report.output_bytes() != requirements.output_bytes
        || prepare_report.fields() != requirements.fields
        || prepare_report.work_bytes() != requirements.work_bytes
        || prepare_report.max_depth() != requirements.max_depth
        || prepare_report.allocations() != requirements.allocations
        || prepare_report.retained_bytes() != requirements.retained_bytes
        || prepare_report.scratch_bytes() != requirements.scratch_bytes
    {
        return Err(SlideTableHeaderError::Verification);
    }
    budget.codec_report(prepare_report)?;
    budget.retained(requirements.retained_bytes)?;
    budget.scratch(requirements.scratch_bytes)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_header_codec_error)?;
    let report = output.report();
    if report.output_bytes() != requirements.output_bytes
        || report.fields() != requirements.fields
        || report.work_bytes() != requirements.work_bytes
        || report.max_depth() != requirements.max_depth
        || report.allocations() != requirements.allocations
        || report.retained_bytes() != requirements.retained_bytes
        || report.scratch_bytes() != requirements.scratch_bytes
    {
        return Err(SlideTableHeaderError::Verification);
    }
    let rewritten = output.into_bytes();
    let verified = decode_header(&rewritten, source, budget)?;
    if settings_from_snapshot(verified)? != after {
        return Err(SlideTableHeaderError::Verification);
    }
    archive
        .object_mut(selection.model_identifier)
        .ok_or(SlideTableHeaderError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.model_message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: rewritten,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_bound = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_bound).map_err(map_core_error)?;
    budget.output(
        encoded_bound
            .checked_add(compressed_bound)
            .ok_or(SlideTableHeaderError::InvalidSource)?,
    )?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    if bytes.len() != encoded_bound {
        return Err(SlideTableHeaderError::Verification);
    }
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(SlideTableHeaderError::Verification);
    }
    let edits = [EntryEdit::new(
        selection.component_name.as_ref(),
        &compressed,
    )];
    let prepared = catalog
        .prepare_reassembly(&edits, physical_limits)
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget.reassembly(requirements)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let candidate_source: Arc<[u8]> = output.into();
    budget.physical(candidate_source.len())?;
    Package::from_source_with_options(candidate_source, source.state.options)
        .map_err(map_read_error)
}

fn settings_to_write(
    settings: Settings,
) -> Result<header_codec::TableHeaderSettingsWrite, SlideTableHeaderError> {
    let count = |value: Option<crate::slide::table::headers::Count>| {
        value
            .map(|count| {
                u32::try_from(count.get()).map_err(|_| SlideTableHeaderError::InvalidSource)
            })
            .transpose()
    };
    Ok(header_codec::TableHeaderSettingsWrite::new(
        count(settings.header_rows)?,
        count(settings.header_columns)?,
        count(settings.footer_rows)?,
        settings.header_rows_frozen,
        settings.header_columns_frozen,
        settings.repeating_header_rows_enabled,
        settings.repeating_header_columns_enabled,
    ))
}

fn decode_header(
    payload: &[u8],
    package: &Package,
    budget: &mut HeaderBudget,
) -> Result<header_codec::TableHeaderSettingsSnapshot, SlideTableHeaderError> {
    let limits = budget.residual(package)?;
    let options = header_codec::DecodeOptions::new(
        limits.max_input_bytes().min(payload.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    );
    let snapshot = header_codec::decode_table_header_settings(payload, options)
        .map_err(map_header_codec_error)?;
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.fields(view.fields().count())?;
    budget.work(
        payload
            .len()
            .checked_mul(4)
            .ok_or(SlideTableHeaderError::InvalidSource)?,
    )?;
    Ok(snapshot)
}

fn decode_table_info(
    payload: &[u8],
    package: &Package,
    budget: &mut HeaderBudget,
) -> Result<table_info_codec::TableInfoSnapshot, SlideTableHeaderError> {
    let limits = budget.residual(package)?;
    let options = table_info_codec::DecodeOptions::new(
        limits.max_input_bytes().min(payload.len().max(1)),
        limits.max_fields(),
        limits.max_rewrite_work(),
        u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
    );
    let snapshot = table_info_codec::decode_table_info(payload, options)
        .map_err(map_table_info_codec_error)?;
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.fields(view.fields().count())?;
    budget.work(
        payload
            .len()
            .checked_mul(4)
            .ok_or(SlideTableHeaderError::InvalidSource)?,
    )?;
    Ok(snapshot)
}

fn settings_from_snapshot(
    snapshot: header_codec::TableHeaderSettingsSnapshot,
) -> Result<Settings, SlideTableHeaderError> {
    let count = |value: Option<u32>| {
        value
            .map(|value| {
                usize::try_from(value)
                    .ok()
                    .and_then(|value| crate::slide::table::headers::Count::new(value).ok())
                    .ok_or(SlideTableHeaderError::InvalidSource)
            })
            .transpose()
    };
    Ok(Settings {
        header_rows: count(snapshot.header_rows())?,
        header_columns: count(snapshot.header_columns())?,
        footer_rows: count(snapshot.footer_rows())?,
        header_rows_frozen: snapshot.header_rows_frozen(),
        header_columns_frozen: snapshot.header_columns_frozen(),
        repeating_header_rows_enabled: snapshot.repeating_header_rows_enabled(),
        repeating_header_columns_enabled: snapshot.repeating_header_columns_enabled(),
    })
}

fn validate_stored(
    settings: Settings,
    rows: u32,
    columns: u32,
) -> Result<(), SlideTableHeaderError> {
    if u64::try_from(settings.header_row_count())
        .unwrap_or(u64::MAX)
        .saturating_add(u64::try_from(settings.footer_row_count()).unwrap_or(u64::MAX))
        > u64::from(rows)
        || u64::try_from(settings.header_column_count()).unwrap_or(u64::MAX) > u64::from(columns)
    {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    Ok(())
}

fn validate_requested(
    settings: Settings,
    rows: u32,
    columns: u32,
    path: SlideTableHeaderPath,
) -> Result<(), SlideTableHeaderError> {
    let header_rows = u8::try_from(settings.header_row_count()).unwrap_or(u8::MAX);
    let footer_rows = u8::try_from(settings.footer_row_count()).unwrap_or(u8::MAX);
    if u16::from(header_rows).saturating_add(u16::from(footer_rows))
        > u16::try_from(rows).unwrap_or(u16::MAX)
    {
        return Err(SlideTableHeaderError::InvalidSettings {
            path,
            reason: SlideTableHeaderInvalidReason::RowSectionsExceedTable {
                header_rows,
                footer_rows,
                table_rows: rows,
            },
        });
    }
    let header_columns = u8::try_from(settings.header_column_count()).unwrap_or(u8::MAX);
    if u32::from(header_columns) > columns {
        return Err(SlideTableHeaderError::InvalidSettings {
            path,
            reason: SlideTableHeaderInvalidReason::HeaderColumnsExceedTable {
                header_columns,
                table_columns: columns,
            },
        });
    }
    Ok(())
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
) -> Result<Position, SlideTableHeaderError> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideTableHeaderError::EmptySlideName);
            }
            package
                .show()
                .map_err(map_read_error)?
                .select_slide(selector)
                .map_err(|_| SlideTableHeaderError::AmbiguousSelector)?
                .map(|slide| Position::new(slide.index()))
                .ok_or(SlideTableHeaderError::SlideNameNotFound)
        },
    }
}

fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<(usize, &[u8]), SlideTableHeaderError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        validate_message_header(object, index)?;
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(SlideTableHeaderError::InvalidSource);
        }
    }
    selected.ok_or(SlideTableHeaderError::InvalidSource)
}

fn validate_message_header(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), SlideTableHeaderError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    let message = object
        .messages
        .get(message_index)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    if info.type_ != message.type_
        || usize::try_from(info.length).ok() != Some(message.data.len())
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    Ok(())
}

fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut HeaderBudget,
) -> Result<Vec<u64>, SlideTableHeaderError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.fields(view.fields().count())?;
    budget.work(payload.len())?;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
    }
    let count = view
        .fields()
        .filter(|field| field.number() == field_number)
        .count();
    let mut result = Vec::new();
    budget.allocations(count)?;
    budget.references(count)?;
    result
        .try_reserve_exact(count)
        .map_err(|_| SlideTableHeaderError::Allocation { amount: count })?;
    for field in view.fields().filter(|field| field.number() == field_number) {
        if field.wire_type() != 2 {
            return Err(SlideTableHeaderError::InvalidSource);
        }
        result.push(strict_reference(field.payload(), limits)?);
    }
    Ok(result)
}

fn table_parent(payload: &[u8], limits: WireLimits) -> Result<u64, SlideTableHeaderError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    for field in fields.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
    }
    let super_fields = fields
        .fields()
        .filter(|field| field.number() == TABLE_SUPER_FIELD)
        .collect::<Vec<_>>();
    if super_fields.len() != 1 || super_fields[0].wire_type() != 2 {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    let drawable =
        WireView::parse_with_limits(super_fields[0].payload(), limits).map_err(map_wire_error)?;
    for field in drawable.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
    }
    let parents = drawable
        .fields()
        .filter(|field| field.number() == DRAWABLE_PARENT_FIELD)
        .collect::<Vec<_>>();
    if parents.len() != 1 || parents[0].wire_type() != 2 {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    strict_reference(parents[0].payload(), limits)
}

fn strict_reference(payload: &[u8], limits: WireLimits) -> Result<u64, SlideTableHeaderError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    for field in fields.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            1 => {
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(SlideTableHeaderError::InvalidSource);
                }
                let (value, width) = decode_varint_from_bytes(field.payload())
                    .map_err(|_| SlideTableHeaderError::InvalidSource)?;
                if value == 0 || width != encoded_len(value) {
                    return Err(SlideTableHeaderError::InvalidSource);
                }
                identifier = Some(value);
            },
            2 | 3 => return Err(SlideTableHeaderError::UnsupportedTopology),
            _ => {},
        }
    }
    identifier.ok_or(SlideTableHeaderError::InvalidSource)
}

fn reject_duplicates(values: &[u64]) -> Result<(), SlideTableHeaderError> {
    if values
        .iter()
        .enumerate()
        .any(|(index, value)| values[..index].contains(value))
    {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    Ok(())
}

fn validate_slide_metadata(
    object: &ArchiveObject,
    message_index: usize,
    owned: &[u64],
    z_order: &[u64],
) -> Result<(), SlideTableHeaderError> {
    validate_message_header(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    reject_duplicates(&info.object_references)?;
    for identifier in owned.iter().chain(z_order) {
        if !info.object_references.is_empty()
            && info
                .object_references
                .iter()
                .filter(|candidate| **candidate == *identifier)
                .count()
                != 1
        {
            return Err(SlideTableHeaderError::InvalidSource);
        }
    }
    for field in &info.field_infos {
        if field.path.as_slice() == [SLIDE_OWNED_DRAWABLES_FIELD] {
            if field
                .r#type
                .is_some_and(|kind| kind != FieldType::ObjectReference)
                || !field.data_references.is_empty()
                || field.object_references.as_slice() != owned
            {
                return Err(SlideTableHeaderError::InvalidSource);
            }
        }
        if field.path.as_slice() == [SLIDE_Z_ORDER_FIELD] {
            if field
                .r#type
                .is_some_and(|kind| kind != FieldType::ObjectReference)
                || !field.data_references.is_empty()
                || field.object_references.as_slice() != z_order
            {
                return Err(SlideTableHeaderError::InvalidSource);
            }
        }
    }
    Ok(())
}

fn validate_table_info_metadata(
    object: &ArchiveObject,
    message_index: usize,
    model: u64,
) -> Result<(), SlideTableHeaderError> {
    validate_message_header(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    if info.object_references.contains(&0)
        || info
            .object_references
            .iter()
            .enumerate()
            .any(|(index, identifier)| info.object_references[..index].contains(identifier))
    {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    // The drawable parent is the rooted payload route above and is not a
    // strong ArchiveInfo edge in native Keynote.  The model edge, when an
    // aggregate is present, remains source-authoritative and unique.
    if !info.object_references.is_empty()
        && info
            .object_references
            .iter()
            .filter(|identifier| **identifier == model)
            .count()
            != 1
    {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    if !info.data_references.is_empty() {
        return Err(SlideTableHeaderError::UnsupportedTopology);
    }
    let mut model_field_count = 0usize;
    for field in &info.field_infos {
        if field.path.as_slice() != [TABLE_MODEL_FIELD] {
            if !field.object_references.is_empty() || !field.data_references.is_empty() {
                return Err(SlideTableHeaderError::UnsupportedTopology);
            }
            continue;
        }
        model_field_count = model_field_count.saturating_add(1);
        if model_field_count > 1
            || field
                .r#type
                .is_some_and(|kind| kind != FieldType::ObjectReference)
            || !field.data_references.is_empty()
            || field.object_references.as_slice() != [model]
        {
            return Err(SlideTableHeaderError::InvalidSource);
        }
    }
    Ok(())
}

fn unique_message_index(
    messages: &[RawMessage],
    message_type: u32,
) -> Result<Option<(usize, &RawMessage)>, SlideTableHeaderError> {
    let mut found = None;
    for (index, message) in messages.iter().enumerate() {
        if message.type_ == message_type && found.replace((index, message)).is_some() {
            return Err(SlideTableHeaderError::InvalidSource);
        }
    }
    Ok(found)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    selection: &HeaderSelection,
    budget: &mut HeaderBudget,
) -> Result<(), SlideTableHeaderError> {
    let source_catalog = physical_catalog(source)?;
    let candidate_catalog = physical_catalog(candidate)?;
    for entry in source_catalog.package().iter() {
        let other = candidate_catalog
            .package()
            .iter()
            .find(|value| value.name() == entry.name())
            .ok_or(SlideTableHeaderError::Verification)?;
        if entry.name() != selection.component_name.as_ref()
            && (entry.data() != other.data() || entry.metadata() != other.metadata())
        {
            return Err(SlideTableHeaderError::Verification);
        }
        budget.work(entry.data().len())?;
    }

    let source_archive = component_archive(source, selection.component_name.as_ref())?;
    let candidate_archive = component_archive(candidate, selection.component_name.as_ref())?;
    if source_archive.objects.len() != candidate_archive.objects.len() {
        return Err(SlideTableHeaderError::Verification);
    }
    let limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    for source_object in &source_archive.objects {
        let identifier = source_object
            .archive_info
            .identifier
            .ok_or(SlideTableHeaderError::Verification)?;
        let other = candidate_archive
            .object(identifier)
            .ok_or(SlideTableHeaderError::Verification)?;
        if identifier != selection.model_identifier
            && !source_object.same_content_ignoring_offsets(other)
        {
            return Err(SlideTableHeaderError::Verification);
        }
        if identifier == selection.model_identifier {
            let message = other
                .messages
                .get(selection.model_message_index)
                .ok_or(SlideTableHeaderError::Verification)?
                .clone();
            let mut expected = source_object.clone();
            expected
                .replace_message_preserving_header_with_limits(
                    selection.model_message_index,
                    message,
                    limits,
                )
                .map_err(map_core_error)?;
            expected.header_length = other.header_length;
            expected.data_length = other.data_length;
            if !expected.same_content_ignoring_offsets(other) {
                return Err(SlideTableHeaderError::Verification);
            }
        }
    }
    Ok(())
}

fn component_archive(package: &Package, name: &str) -> Result<Archive, SlideTableHeaderError> {
    let catalog = physical_catalog(package)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        package
            .state
            .options
            .archive()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    Archive::parse_with_limits(
        stream.as_bytes(),
        package
            .state
            .options
            .archive()
            .effective_archive_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)
}

fn ensure_unique_identity(
    package: &Package,
    identifier: u64,
    budget: &mut HeaderBudget,
) -> Result<(), SlideTableHeaderError> {
    let mut count = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            budget.work(
                1usize
                    .checked_add(object.messages.len())
                    .ok_or(SlideTableHeaderError::InvalidSource)?,
            )?;
            if object.archive_info.identifier == Some(identifier) {
                count = count.saturating_add(1);
            }
        }
    }
    if count == 1 {
        Ok(())
    } else {
        Err(SlideTableHeaderError::UnsupportedDependency)
    }
}

fn ensure_unique_table_owner(
    package: &Package,
    slide: u64,
    table_info: u64,
    model: u64,
    limits: WireLimits,
    budget: &mut HeaderBudget,
) -> Result<(), SlideTableHeaderError> {
    let mut owned_count = 0usize;
    let mut z_count = 0usize;
    let mut selected = false;
    let mut model_owners = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                budget.work(1)?;
                if message.type_ == SLIDE_MESSAGE_TYPE {
                    let owned = repeated_references(
                        &message.data,
                        SLIDE_OWNED_DRAWABLES_FIELD,
                        limits,
                        budget,
                    )?;
                    let z =
                        repeated_references(&message.data, SLIDE_Z_ORDER_FIELD, limits, budget)?;
                    let owned_hits = owned.iter().filter(|id| **id == table_info).count();
                    let z_hits = z.iter().filter(|id| **id == table_info).count();
                    owned_count = owned_count.saturating_add(owned_hits);
                    z_count = z_count.saturating_add(z_hits);
                    if object.archive_info.identifier == Some(slide)
                        && owned_hits == 1
                        && z_hits == 1
                    {
                        selected = true;
                    }
                }
                if message.type_ == TABLE_INFO_MESSAGE_TYPE {
                    budget.allocations(1)?;
                    let info = decode_table_info(&message.data, package, budget)?;
                    if info.table_model().identifier().get() == model {
                        model_owners = model_owners.saturating_add(1);
                    }
                }
            }
        }
    }
    if owned_count == 1 && z_count == 1 && selected && model_owners == 1 {
        Ok(())
    } else {
        Err(SlideTableHeaderError::UnsupportedDependency)
    }
}

fn validate_global_inbound_references(
    package: &Package,
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
    budget: &mut HeaderBudget,
) -> Result<(), SlideTableHeaderError> {
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let Some((_component, table_info)) = package.object_with_component(table_info_identifier)
    else {
        return Err(SlideTableHeaderError::UnsupportedDependency);
    };
    let Some(message_info) = table_info
        .archive_info
        .message_infos
        .get(table_info_message_index)
    else {
        return Err(SlideTableHeaderError::UnsupportedDependency);
    };
    let expected_model_edges = message_info
        .object_references
        .iter()
        .chain(
            message_info
                .field_infos
                .iter()
                .flat_map(|field| field.object_references.iter()),
        )
        .filter(|identifier| **identifier == model_identifier)
        .count();
    if expected_model_edges == 0 {
        return Err(SlideTableHeaderError::UnsupportedDependency);
    }
    let mut census = InboundReferenceCensus {
        slide_identifier,
        slide_message_index,
        table_info_identifier,
        table_info_message_index,
        model_identifier,
        model_edges: 0,
        expected_model_edges,
        occurrences: 0,
        invalid: false,
    };
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            let metadata_fields = object
                .archive_info
                .message_infos
                .iter()
                .try_fold(0usize, |total, info| {
                    total
                        .checked_add(info.object_references.len())
                        .and_then(|value| value.checked_add(info.data_references.len()))
                        .and_then(|value| value.checked_add(info.field_infos.len()))
                })
                .ok_or(SlideTableHeaderError::InvalidSource)?;
            let message_bytes = object
                .messages
                .iter()
                .try_fold(0usize, |total, message| {
                    total.checked_add(message.data.len())
                })
                .ok_or(SlideTableHeaderError::InvalidSource)?;
            budget.fields(metadata_fields)?;
            budget.work(
                message_bytes
                    .checked_add(metadata_fields)
                    .ok_or(SlideTableHeaderError::InvalidSource)?,
            )?;
            budget.allocations(
                object
                    .messages
                    .len()
                    .checked_add(object.archive_info.message_infos.len())
                    .ok_or(SlideTableHeaderError::InvalidSource)?,
            )?;
            let before_occurrences = census.occurrences;
            object
                .inspect_references_with_policy_and_limits(
                    &mut census,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(map_core_error)?;
            budget.references(
                census
                    .occurrences
                    .checked_sub(before_occurrences)
                    .ok_or(SlideTableHeaderError::InvalidSource)?,
            )?;
        }
    }
    if census.invalid
        || census.model_edges != census.expected_model_edges
        || !validate_selected_edge_paths(
            package,
            slide_identifier,
            slide_message_index,
            table_info_identifier,
            table_info_message_index,
            model_identifier,
        )?
    {
        return Err(SlideTableHeaderError::UnsupportedDependency);
    }
    Ok(())
}

fn validate_selected_edge_paths(
    package: &Package,
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
) -> Result<bool, SlideTableHeaderError> {
    let Some((_slide_component, slide)) = package.object_with_component(slide_identifier) else {
        return Ok(false);
    };
    let Some(slide_info) = slide.archive_info.message_infos.get(slide_message_index) else {
        return Ok(false);
    };
    for field in &slide_info.field_infos {
        if field.object_references.contains(&table_info_identifier)
            && field.path.as_slice() != [SLIDE_OWNED_DRAWABLES_FIELD]
            && field.path.as_slice() != [SLIDE_Z_ORDER_FIELD]
        {
            return Ok(false);
        }
    }
    let Some((_table_component, table_info)) = package.object_with_component(table_info_identifier)
    else {
        return Ok(false);
    };
    let Some(table_info_info) = table_info
        .archive_info
        .message_infos
        .get(table_info_message_index)
    else {
        return Ok(false);
    };
    let mut model_field_count = 0usize;
    for field in &table_info_info.field_infos {
        if field.object_references.contains(&model_identifier)
            && field.path.as_slice() != [TABLE_MODEL_FIELD]
        {
            return Ok(false);
        }
        if field.path.as_slice() == [TABLE_MODEL_FIELD]
            && field.object_references.as_slice() != [model_identifier]
        {
            return Ok(false);
        }
        if field.path.as_slice() == [TABLE_MODEL_FIELD] {
            model_field_count = model_field_count.saturating_add(1);
        }
    }
    Ok(model_field_count <= 1)
}

struct InboundReferenceCensus {
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
    model_edges: usize,
    expected_model_edges: usize,
    occurrences: usize,
    invalid: bool,
}

impl ArchiveReferenceVisitor for InboundReferenceCensus {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        self.occurrences = self.occurrences.saturating_add(1);
        if occurrence.referenced_identifier == self.model_identifier {
            if occurrence.kind == ArchiveReferenceKind::Object
                && occurrence.object_identifier == self.table_info_identifier
                && occurrence.message_index == self.table_info_message_index
            {
                self.model_edges = self.model_edges.saturating_add(1);
            } else {
                self.invalid = true;
            }
        }
        if occurrence.referenced_identifier == self.table_info_identifier
            && (occurrence.kind != ArchiveReferenceKind::Object
                || occurrence.object_identifier != self.slide_identifier
                || occurrence.message_index != self.slide_message_index)
        {
            self.invalid = true;
        }
        let _known_scope = matches!(
            occurrence.scope,
            ArchiveReferenceScope::Message | ArchiveReferenceScope::Field { .. }
        );
        Ok(())
    }
}

fn same_selection(a: &HeaderSelection, b: &HeaderSelection) -> bool {
    a.slide_position == b.slide_position
        && a.table_position == b.table_position
        && a.slide_identifier == b.slide_identifier
        && a.table_info_identifier == b.table_info_identifier
        && a.model_identifier == b.model_identifier
        && a.slide_message_index == b.slide_message_index
        && a.table_info_message_index == b.table_info_message_index
        && a.model_message_index == b.model_message_index
        && a.component_name == b.component_name
}

fn physical_catalog(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideTableHeaderError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideTableHeaderError::UnsupportedSource),
    }
}

fn map_read_error(error: ReadError) -> SlideTableHeaderError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableHeaderError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => SlideTableHeaderLimitKind::References,
                _ => SlideTableHeaderLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideTableHeaderError::Allocation { amount },
        _ => SlideTableHeaderError::InvalidSource,
    }
}

fn map_header_codec_error(error: header_codec::DecodeError) -> SlideTableHeaderError {
    if let Some(amount) = error.allocation_amount() {
        return SlideTableHeaderError::Allocation { amount };
    }
    if let Some((observed, maximum)) = error.field_limit_values() {
        return header_limit(SlideTableHeaderLimitKind::WireFields, observed, maximum);
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return header_limit(SlideTableHeaderLimitKind::WireWork, observed, maximum);
    }
    if let Some((observed, maximum)) = error.output_limit_values() {
        return header_limit(
            SlideTableHeaderLimitKind::WireOutputBytes,
            observed,
            maximum,
        );
    }
    if let Some((observed, maximum)) = error.allocation_limit_values() {
        return header_limit(SlideTableHeaderLimitKind::Allocations, observed, maximum);
    }
    if let Some((observed, maximum)) = error.retained_limit_values() {
        return header_limit(SlideTableHeaderLimitKind::Retained, observed, maximum);
    }
    if let Some((observed, maximum)) = error.scratch_limit_values() {
        return header_limit(SlideTableHeaderLimitKind::Scratch, observed, maximum);
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            header_codec::WireResourceLimit::Bytes { observed, maximum } => {
                header_limit(SlideTableHeaderLimitKind::WireBytes, observed, maximum)
            },
            header_codec::WireResourceLimit::Nesting { observed, maximum } => {
                SlideTableHeaderError::LimitExceeded {
                    kind: SlideTableHeaderLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => SlideTableHeaderError::InvalidSource,
        };
    }
    SlideTableHeaderError::InvalidSource
}

fn map_table_info_codec_error(error: table_info_codec::DecodeError) -> SlideTableHeaderError {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return header_limit(SlideTableHeaderLimitKind::WireFields, observed, maximum);
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return header_limit(SlideTableHeaderLimitKind::WireWork, observed, maximum);
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            table_info_codec::WireResourceLimit::Bytes { observed, maximum } => {
                SlideTableHeaderError::LimitExceeded {
                    kind: SlideTableHeaderLimitKind::WireBytes,
                    observed: observed.unwrap_or(0) as u64,
                    maximum: maximum.unwrap_or(0) as u64,
                }
            },
            table_info_codec::WireResourceLimit::Nesting { observed, maximum } => {
                SlideTableHeaderError::LimitExceeded {
                    kind: SlideTableHeaderLimitKind::WireNesting,
                    observed: u64::from(observed.unwrap_or(0)),
                    maximum: u64::from(maximum.unwrap_or(0)),
                }
            },
            _ => SlideTableHeaderError::InvalidSource,
        };
    }
    SlideTableHeaderError::InvalidSource
}

fn header_limit(
    kind: SlideTableHeaderLimitKind,
    observed: usize,
    maximum: usize,
) -> SlideTableHeaderError {
    SlideTableHeaderError::LimitExceeded {
        kind,
        observed: observed as u64,
        maximum: maximum as u64,
    }
}

fn map_wire_error(_error: litchi_iwa_common::Error) -> SlideTableHeaderError {
    SlideTableHeaderError::InvalidSource
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideTableHeaderError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableHeaderError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => SlideTableHeaderLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SlideTableHeaderLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => SlideTableHeaderLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes => {
                    SlideTableHeaderLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => SlideTableHeaderLimitKind::TotalBytes,
                _ => SlideTableHeaderLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideTableHeaderError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => SlideTableHeaderError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SlideTableHeaderError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableHeaderError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => SlideTableHeaderLimitKind::PayloadObjects,
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideTableHeaderLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => SlideTableHeaderLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => SlideTableHeaderLimitKind::EntryBytes,
                _ => SlideTableHeaderLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTableHeaderError::Allocation { amount: requested }
        },
        _ => SlideTableHeaderError::InvalidSource,
    }
}

fn parse_all_fields<'a>(
    payload: &'a [u8],
    limits: WireLimits,
) -> Result<WireView<'a>, SlideTableHeaderError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    for field in view.fields() {
        field.validate_canonical_framing().map_err(map_wire_error)?;
    }
    Ok(view)
}

fn charge_wire_view(
    payload: &[u8],
    view: &WireView<'_>,
    budget: &mut HeaderBudget,
) -> Result<(), SlideTableHeaderError> {
    let fields = view.fields().count();
    budget.fields(fields)?;
    budget.work(payload.len())?;
    budget.allocations(fields)
}

fn repeated_length_payloads(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
) -> Result<Vec<&[u8]>, SlideTableHeaderError> {
    let view = parse_all_fields(payload, limits)?;
    let count = view
        .fields()
        .filter(|field| field.number() == field_number)
        .count();
    let mut values = Vec::new();
    values
        .try_reserve_exact(count)
        .map_err(|_| SlideTableHeaderError::Allocation { amount: count })?;
    for field in view.fields().filter(|field| field.number() == field_number) {
        if field.wire_type() != 2 {
            return Err(SlideTableHeaderError::InvalidSource);
        }
        values.push(field.payload());
    }
    Ok(values)
}

fn canonical_varint(payload: &[u8]) -> Result<u64, SlideTableHeaderError> {
    let (value, width) =
        decode_varint_from_bytes(payload).map_err(|_| SlideTableHeaderError::InvalidSource)?;
    if width != encoded_len(value) {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    Ok(value)
}

fn validate_dependencies(
    source: &Package,
    selection: &HeaderSelection,
    before: Settings,
    after: Settings,
    budget: &mut HeaderBudget,
) -> Result<(), SlideTableHeaderError> {
    let model = source
        .object_with_component(selection.model_identifier)
        .map(|(_, object)| object)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let model_message = model
        .messages
        .get(selection.model_message_index)
        .filter(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let limits = budget.residual(source)?;
    let fields = parse_all_fields(&model_message.data, limits)?;
    charge_wire_view(&model_message.data, &fields, budget)?;
    let header_counts_changed =
        before.header_rows != after.header_rows || before.header_columns != after.header_columns;
    let section_counts_changed = header_counts_changed || before.footer_rows != after.footer_rows;
    let mut active_pivot_or_group = false;
    let mut header_count_dependency = false;
    let mut seen = [false; 6];
    for field in fields.fields() {
        let slot = match field.number() {
            CATEGORY_GROUPING_FIELD => Some(0),
            GROUPING_FIELD => Some(1),
            PIVOT_FIELD => Some(2),
            CATEGORY_OWNER_FIELD => Some(3),
            84 => Some(4),
            _ => None,
        };
        let Some(slot) = slot else {
            continue;
        };
        if seen[slot] {
            return Err(SlideTableHeaderError::InvalidSource);
        }
        seen[slot] = true;
        if field.wire_type() != 2 {
            return Err(SlideTableHeaderError::InvalidSource);
        }
        match field.number() {
            GROUPING_FIELD => {
                if !field.payload().is_empty() {
                    active_pivot_or_group = true;
                    header_count_dependency = true;
                }
            },
            PIVOT_FIELD => {
                let identifier = strict_reference(field.payload(), limits)?;
                if identifier == 1
                    || identifier == selection.slide_identifier
                    || identifier == selection.table_info_identifier
                    || identifier == selection.model_identifier
                {
                    return Err(SlideTableHeaderError::InvalidSource);
                }
                require_declared_reference(
                    model,
                    selection.model_message_index,
                    identifier,
                    &[PIVOT_FIELD],
                )?;
                // Pivot ownership is an opaque graph dependency for this
                // scalar transaction.  Even freeze/repeat-only edits could
                // otherwise publish from a dangling or wrong-typed owner, so
                // reject every changed edit that carries field 85.
                return Err(SlideTableHeaderError::UnsupportedDependency);
            },
            CATEGORY_GROUPING_FIELD | CATEGORY_OWNER_FIELD | 84 => {
                let active = if field.number() == CATEGORY_GROUPING_FIELD {
                    deprecated_category_grouping_active(field.payload(), limits, budget)?
                } else if field.number() == CATEGORY_OWNER_FIELD {
                    category_owner_reference_active(
                        source,
                        selection,
                        field.payload(),
                        limits,
                        budget,
                    )?
                } else {
                    validate_haunted_owner(field.payload(), limits, budget)?;
                    // A nonzero HauntedOwner marker denotes formula-cache
                    // state keyed to table coordinates.  This owner does not
                    // expose a safe scalar transition here, so section-count
                    // changes must fail closed while freeze/repeat flags may
                    // still be edited.
                    true
                };
                active_pivot_or_group |= active;
                header_count_dependency |= active;
            },
            _ => return Err(SlideTableHeaderError::InvalidSource),
        }
    }
    if header_counts_changed && header_count_dependency {
        return Err(SlideTableHeaderError::UnsupportedDependency);
    }
    if section_counts_changed && active_pivot_or_group {
        return Err(SlideTableHeaderError::UnsupportedDependency);
    }
    let table_info_dependency = table_info_has_count_dependency(
        source,
        selection,
        header_counts_changed,
        section_counts_changed,
        budget,
    )?;
    if table_info_dependency {
        return Err(SlideTableHeaderError::UnsupportedDependency);
    }
    if header_counts_changed && has_rooted_header_name_manager(source, budget)? {
        return Err(SlideTableHeaderError::UnsupportedDependency);
    }
    Ok(())
}

fn has_rooted_header_name_manager(
    source: &Package,
    budget: &mut HeaderBudget,
) -> Result<bool, SlideTableHeaderError> {
    let mut roots = source
        .state
        .source
        .components()
        .iter()
        .filter(|component| component.name().rsplit('/').next() == Some("Document.iwa"));
    let root_component = roots.next().ok_or(SlideTableHeaderError::InvalidSource)?;
    if roots.next().is_some() {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    let root = root_component
        .archive()
        .object(1)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let (root_message_index, root_payload) = unique_message(root, super::DOCUMENT_MESSAGE_TYPE)?;
    validate_message_header(root, root_message_index)?;
    let limits = budget.residual(source)?;
    let base_document = singular_length_payload(root_payload, 3, limits, budget)?
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let Some(engine_reference) = singular_length_payload(base_document, 4, limits, budget)? else {
        return Ok(false);
    };
    let reference_fields = parse_all_fields(engine_reference, limits)?;
    charge_wire_view(engine_reference, &reference_fields, budget)?;
    let engine_identifier = strict_reference(engine_reference, limits)?;
    require_declared_reference(root, root_message_index, engine_identifier, &[3, 4])?;
    ensure_unique_identity(source, engine_identifier, budget)?;
    let engine = source
        .object_with_component(engine_identifier)
        .map(|(_, object)| object)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let (engine_message_index, engine_message) =
        unique_message_index(&engine.messages, CALCULATION_ENGINE_MESSAGE_TYPE)?
            .ok_or(SlideTableHeaderError::InvalidSource)?;
    validate_message_header(engine, engine_message_index)?;
    let Some(manager_reference) =
        singular_length_payload(&engine_message.data, 14, limits, budget)?
    else {
        return Ok(false);
    };
    let reference_fields = parse_all_fields(manager_reference, limits)?;
    charge_wire_view(manager_reference, &reference_fields, budget)?;
    let manager_identifier = strict_reference(manager_reference, limits)?;
    require_declared_reference(engine, engine_message_index, manager_identifier, &[14])?;
    ensure_unique_identity(source, manager_identifier, budget)?;
    let manager = source
        .object_with_component(manager_identifier)
        .map(|(_, object)| object)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let (manager_message_index, _) =
        unique_message_index(&manager.messages, HEADER_NAME_MANAGER_MESSAGE_TYPE)?
            .ok_or(SlideTableHeaderError::InvalidSource)?;
    validate_message_header(manager, manager_message_index)?;
    Ok(true)
}

fn singular_length_payload<'a>(
    payload: &'a [u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut HeaderBudget,
) -> Result<Option<&'a [u8]>, SlideTableHeaderError> {
    let fields = parse_all_fields(payload, limits)?;
    charge_wire_view(payload, &fields, budget)?;
    let mut selected = None;
    for field in fields
        .fields()
        .filter(|field| field.number() == field_number)
    {
        if field.wire_type() != 2 || selected.replace(field.payload()).is_some() {
            return Err(SlideTableHeaderError::InvalidSource);
        }
    }
    Ok(selected)
}

fn table_info_has_count_dependency(
    source: &Package,
    selection: &HeaderSelection,
    header_counts_changed: bool,
    section_counts_changed: bool,
    budget: &mut HeaderBudget,
) -> Result<bool, SlideTableHeaderError> {
    let object = source
        .object_with_component(selection.table_info_identifier)
        .map(|(_, object)| object)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let message = object
        .messages
        .get(selection.table_info_message_index)
        .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let limits = budget.residual(source)?;
    let fields = parse_all_fields(&message.data, limits)?;
    charge_wire_view(&message.data, &fields, budget)?;
    let mut seen = [false; 7];
    let mut pivot_active = false;
    for field in fields.fields() {
        let slot = match field.number() {
            4 => Some(0),
            5 => Some(1),
            7 => Some(2),
            8 => Some(3),
            15 => Some(4),
            16 => Some(5),
            17 => Some(6),
            _ => None,
        };
        let Some(slot) = slot else {
            continue;
        };
        if seen[slot] {
            return Err(SlideTableHeaderError::InvalidSource);
        }
        seen[slot] = true;
        field.validate_canonical_framing().map_err(map_wire_error)?;
        match field.number() {
            4 | 5 | 15 | 17 if field.wire_type() == 2 => {
                let identifier = strict_reference(field.payload(), limits)?;
                if identifier == 1
                    || identifier == selection.slide_identifier
                    || identifier == selection.table_info_identifier
                    || identifier == selection.model_identifier
                {
                    return Err(SlideTableHeaderError::InvalidSource);
                }
                require_declared_reference(
                    object,
                    selection.table_info_message_index,
                    identifier,
                    &[field.number()],
                )?;
                // Native summary/category/pivot cache routes are all keyed to
                // table coordinates.  Their mere presence blocks section
                // count edits; the explicit pivot marker must not erase an
                // earlier cache dependency when it is false.
                pivot_active = true;
            },
            7 | 8 if field.wire_type() == 2 => {
                validate_uuid_payload(field.payload(), limits, budget)?;
            },
            16 if field.wire_type() == 0 => {
                let value = canonical_varint(field.payload())?;
                if value > 1 {
                    return Err(SlideTableHeaderError::InvalidSource);
                }
                pivot_active |= value == 1;
            },
            _ => return Err(SlideTableHeaderError::InvalidSource),
        }
    }
    // Native Keynote emits summary/category cache references and UUIDs on
    // ordinary tables.  The references remain a count dependency even when
    // the explicit pivot marker is false; freeze/repeat-only edits remain
    // safe because they do not change section coordinates.
    let _ = header_counts_changed;
    Ok(section_counts_changed && pivot_active)
}

fn validate_uuid_payload(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut HeaderBudget,
) -> Result<(u64, u64), SlideTableHeaderError> {
    let fields = parse_all_fields(payload, limits)?;
    charge_wire_view(payload, &fields, budget)?;
    let mut values = [None; 2];
    for field in fields.fields() {
        let slot = match field.number() {
            1 => 0,
            2 => 1,
            _ => return Err(SlideTableHeaderError::InvalidSource),
        };
        if values[slot].is_some() || field.wire_type() != 0 {
            return Err(SlideTableHeaderError::InvalidSource);
        }
        values[slot] = Some(canonical_varint(field.payload())?);
    }
    Ok((
        values[0].ok_or(SlideTableHeaderError::InvalidSource)?,
        values[1].ok_or(SlideTableHeaderError::InvalidSource)?,
    ))
}

fn validate_haunted_owner(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut HeaderBudget,
) -> Result<(), SlideTableHeaderError> {
    let fields = parse_all_fields(payload, limits)?;
    charge_wire_view(payload, &fields, budget)?;
    let mut owner_uid = None;
    for field in fields.fields() {
        if field.number() != 1 || field.wire_type() != 2 || owner_uid.is_some() {
            return Err(SlideTableHeaderError::InvalidSource);
        }
        owner_uid = Some(field.payload());
    }
    let owner_uid = validate_uuid_payload(
        owner_uid.ok_or(SlideTableHeaderError::InvalidSource)?,
        limits,
        budget,
    )?;
    if owner_uid == (0, 0) {
        Err(SlideTableHeaderError::InvalidSource)
    } else {
        Ok(())
    }
}

fn deprecated_category_grouping_active(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut HeaderBudget,
) -> Result<bool, SlideTableHeaderError> {
    let fields = parse_all_fields(payload, limits)?;
    charge_wire_view(payload, &fields, budget)?;
    let mut active = false;
    let mut owner_uid_seen = false;
    for field in fields.fields() {
        match field.number() {
            1 if !owner_uid_seen && field.wire_type() == 2 => {
                owner_uid_seen = true;
                validate_uuid_payload(field.payload(), limits, budget)?;
            },
            2 if field.wire_type() == 2 => {
                active |= group_by_enabled(field.payload(), limits, budget)?
                    .ok_or(SlideTableHeaderError::InvalidSource)?;
            },
            _ => return Err(SlideTableHeaderError::InvalidSource),
        }
    }
    if owner_uid_seen {
        Ok(active)
    } else {
        Err(SlideTableHeaderError::InvalidSource)
    }
}

fn group_by_enabled(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut HeaderBudget,
) -> Result<Option<bool>, SlideTableHeaderError> {
    let fields = parse_all_fields(payload, limits)?;
    charge_wire_view(payload, &fields, budget)?;
    let mut enabled = None;
    let mut group_uid_seen = false;
    let mut owner_index_seen = false;
    for field in fields.fields() {
        match field.number() {
            1 if !group_uid_seen && field.wire_type() == 2 => {
                group_uid_seen = true;
                validate_uuid_payload(field.payload(), limits, budget)?;
            },
            6 if enabled.is_none() && field.wire_type() == 0 => {
                let value = canonical_varint(field.payload())?;
                if value > 1 {
                    return Err(SlideTableHeaderError::InvalidSource);
                }
                enabled = Some(value == 1);
            },
            // The default owner index is emitted by native Keynote even for
            // an otherwise dormant group.  It is scalar metadata, not an
            // active coordinate dependency.
            14 if !owner_index_seen && field.wire_type() == 0 => {
                owner_index_seen = true;
                let value = canonical_varint(field.payload())?;
                let minimum_signed = i32::MIN as i64 as u64;
                if value > i32::MAX as u64 && value < minimum_signed {
                    return Err(SlideTableHeaderError::InvalidSource);
                }
            },
            // Every other known GroupBy payload carries coordinate, formula,
            // aggregate, row-UID, or child-reference state that this scalar
            // header transaction cannot regenerate safely.
            2..=5 | 7..=13 | 15..=18 if field.wire_type() == 2 => {
                return Err(SlideTableHeaderError::UnsupportedDependency);
            },
            _ => return Err(SlideTableHeaderError::InvalidSource),
        }
    }
    if group_uid_seen {
        Ok(enabled)
    } else {
        Err(SlideTableHeaderError::InvalidSource)
    }
}

fn category_owner_reference_active(
    source: &Package,
    selection: &HeaderSelection,
    payload: &[u8],
    limits: WireLimits,
    budget: &mut HeaderBudget,
) -> Result<bool, SlideTableHeaderError> {
    let reference_fields = parse_all_fields(payload, limits)?;
    charge_wire_view(payload, &reference_fields, budget)?;
    let owner_identifier = strict_reference(payload, limits)?;
    if owner_identifier == 1
        || owner_identifier == selection.slide_identifier
        || owner_identifier == selection.table_info_identifier
        || owner_identifier == selection.model_identifier
    {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    let model_object = source
        .object_with_component(selection.model_identifier)
        .map(|(_, object)| object)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    require_declared_reference(
        model_object,
        selection.model_message_index,
        owner_identifier,
        &[CATEGORY_OWNER_FIELD],
    )?;
    let owner_object = source
        .object_with_component(owner_identifier)
        .map(|(_, object)| object)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    let (owner_message_index, owner_message) = unique_message_index(
        &owner_object.messages,
        CATEGORY_OWNER_REFERENCE_MESSAGE_TYPE,
    )?
    .ok_or(SlideTableHeaderError::InvalidSource)?;
    validate_message_header(owner_object, owner_message_index)?;
    let references = repeated_length_payloads(&owner_message.data, 1, limits)?;
    let owner_fields = parse_all_fields(&owner_message.data, limits)?;
    charge_wire_view(&owner_message.data, &owner_fields, budget)?;
    if owner_fields
        .fields()
        .any(|field| field.number() != 1 || field.wire_type() != 2)
    {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    budget.allocations(references.len())?;
    budget.references(references.len())?;
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(references.len())
        .map_err(|_| SlideTableHeaderError::Allocation {
            amount: references.len(),
        })?;
    for reference in references {
        let identifier = strict_reference(reference, limits)?;
        if identifier == 1
            || identifier == selection.slide_identifier
            || identifier == selection.table_info_identifier
            || identifier == selection.model_identifier
            || identifier == owner_identifier
            || identifiers.contains(&identifier)
        {
            return Err(SlideTableHeaderError::InvalidSource);
        }
        identifiers.push(identifier);
    }
    require_declared_group_references(owner_object, owner_message_index, &identifiers)?;
    let mut active = false;
    for identifier in identifiers {
        let group_object = source
            .object_with_component(identifier)
            .map(|(_, object)| object)
            .ok_or(SlideTableHeaderError::InvalidSource)?;
        let (message_index, message) =
            unique_message_index(&group_object.messages, GROUP_BY_MESSAGE_TYPE)?
                .ok_or(SlideTableHeaderError::InvalidSource)?;
        validate_message_header(group_object, message_index)?;
        active |= group_by_enabled(&message.data, limits, budget)?
            .ok_or(SlideTableHeaderError::InvalidSource)?;
    }
    Ok(active)
}

fn require_declared_group_references(
    object: &ArchiveObject,
    message_index: usize,
    identifiers: &[u64],
) -> Result<(), SlideTableHeaderError> {
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    for identifier in identifiers {
        if info
            .object_references
            .iter()
            .filter(|candidate| **candidate == *identifier)
            .count()
            != 1
        {
            return Err(SlideTableHeaderError::InvalidSource);
        }
    }
    for field in &info.field_infos {
        for identifier in &field.object_references {
            if identifiers.contains(identifier)
                && (field.path.as_slice() != [1]
                    || info
                        .field_infos
                        .iter()
                        .filter(|candidate| {
                            candidate.path.as_slice() == [1]
                                && candidate.object_references.contains(identifier)
                        })
                        .count()
                        != 1)
            {
                return Err(SlideTableHeaderError::InvalidSource);
            }
        }
    }
    Ok(())
}

fn require_declared_reference(
    object: &ArchiveObject,
    message_index: usize,
    identifier: u64,
    accepted_path: &[u32],
) -> Result<(), SlideTableHeaderError> {
    validate_message_header(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableHeaderError::InvalidSource)?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(SlideTableHeaderError::InvalidSource);
    }
    let mut field_occurrence = false;
    for field in &info.field_infos {
        let count = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if count != 0 {
            if count != 1
                || field_occurrence
                || field.path.as_slice() != accepted_path
                || field
                    .r#type
                    .is_some_and(|kind| kind != FieldType::ObjectReference)
                || !field.data_references.is_empty()
            {
                return Err(SlideTableHeaderError::InvalidSource);
            }
            field_occurrence = true;
        }
    }
    Ok(())
}
