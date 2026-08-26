//! Selector-first Pop-Up Menu cell-control transactions.
//!
//! This is the package boundary for the Numbers Pop-Up Menu owner. The
//! native `TST.PopUpMenuModel` graph is handled through strict, generated-free
//! codecs plus a refcount/ownership census and metadata transition. Semantic
//! callers therefore do not need native identifiers or row/column integer
//! pairs.

use std::fmt;

use litchi_iwa_archive::package::{
    EntryEdit, OwnedExactArtifacts, ReassemblyExecutionRequirements,
};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    numbers_table_cell_control_codec as control_codec,
    numbers_table_cell_pop_up_menu_codec as popup_codec,
    numbers_table_cell_storage_codec as storage_codec,
    package_metadata_codec::{
        AdditionSaveTokenBatch, Batch as MetadataBatch, RemovalBatch, RemovalSaveTokenBatch,
        RewriteOptions as MetadataRewriteOptions, SaveTokenBatch,
    },
};
use litchi_numbers_wire::BncCell;
use thiserror::Error as ThisError;

use super::{
    Package, table_cell_pop_up_menu_metadata as popup_metadata,
    table_cell_pop_up_menu_native as popup_native,
};
use crate::{
    SheetSelector, TableSelector,
    cell::data_format::control::{DisplayFormat, Range},
    cell::data_format::number::{
        CurrencyCode, DecimalPlaces, FixedDecimalPlaces, FractionAccuracy,
        NegativeStyle as NumberNegativeStyle, ThousandsSeparator,
    },
    cell::data_format::numeral_system::{
        Base, FixedPlaces, NegativeStyle as NumeralNegativeStyle, Places,
    },
    cell::data_format::{
        CellControl, Checkbox, Currency, CurrencyStyle, Fraction, Number, NumeralSystem,
        Percentage, PopUpMenu, Scientific, Slider, StarRating, Stepper,
    },
    table::CellPosition,
};

const LIST_FORMAT: i32 = 2;
const LIST_CONTROL_CELL_SPEC: i32 = 12;

/// A content-free location associated with a Pop-Up Menu operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Numbers package.
    Package,
    /// One selected rooted table cell.
    Cell {
        /// Zero-based rooted sheet position.
        sheet: usize,
        /// Zero-based table position within the sheet.
        table: usize,
        /// Zero-based semantic cell coordinate.
        position: CellPosition,
    },
}

/// A finite resource governed by a Pop-Up Menu transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete package input bytes.
    InputBytes,
    /// Complete candidate package output bytes.
    OutputBytes,
    /// Physical package entries.
    Entries,
    /// Bytes in one physical entry.
    EntryBytes,
    /// Aggregate physical entry bytes.
    TotalEntryBytes,
    /// Physical package/container metadata bytes.
    PackageBytes,
    /// Decoded native payload bytes.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages inspected.
    PayloadMessages,
    /// Native framing/items inspected.
    PayloadItems,
    /// Native object references inspected.
    PayloadReferences,
    /// Strict codec input bytes.
    WireBytes,
    /// Strict codec output bytes.
    WireOutputBytes,
    /// Strict codec reference-envelope bytes.
    WireReferenceBytes,
    /// Strict codec selected-text bytes.
    WireTextBytes,
    /// Strict codec fields inspected.
    WireFields,
    /// Strict codec nesting depth.
    WireNesting,
    /// Strict codec work.
    WireWork,
    /// Private codec/reassembly scratch bytes.
    ScratchBytes,
    /// Private candidate bytes retained during execution.
    RetainedBytes,
    /// Bounded allocation units reserved by a phase.
    Allocations,
    /// Compressed physical member bytes.
    CompressedBytes,
    /// Aggregate transaction work.
    TransactionWork,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// A content-redacted Pop-Up Menu transaction failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// No rooted sheet matched the selector.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// No table on the selected sheet matched the selector.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// The selected semantic cell was not present in the rooted table.
    #[error("the selected Numbers table cell was not found")]
    CellNotFound,
    /// A changed operation targeted an effectively locked table.
    #[error("the selected Numbers table is locked at {path:?}")]
    TableLocked { path: Path },
    /// A native dependency is not safe to rewrite in this owner.
    #[error("the selected Numbers Pop-Up Menu has an unsupported dependency at {path:?}")]
    UnsupportedDependency { path: Path },
    /// The exact native Pop-Up Menu owner is not available for this source.
    #[error("this Numbers source does not support exact Pop-Up Menu editing")]
    UnsupportedSource,
    /// Rooted ownership or wire framing is invalid.
    #[error("the Numbers Pop-Up Menu source is invalid at {path:?}")]
    InvalidSource { path: Path },
    /// A finite transaction resource ceiling was exceeded.
    #[error("Numbers Pop-Up Menu {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
        /// Content-free failure location.
        path: Path,
    },
    /// A bounded transaction allocation failed.
    #[error("could not allocate {amount} units for the Numbers Pop-Up Menu transaction")]
    Allocation { amount: usize, path: Path },
    /// Candidate reopening or semantic locality verification failed.
    #[error("the edited Numbers Pop-Up Menu failed semantic verification")]
    Verification,
    /// A patch was created for another exact package artifact.
    #[error("the Pop-Up Menu patch does not match the exact source package")]
    PatchConflict,
}

/// A selector-first Pop-Up Menu edit staged against one package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    path: Path,
    before: Option<PopUpMenu>,
    after: Option<PopUpMenu>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the selected semantic cell.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    /// Return the value staged for publication, if any.
    #[must_use]
    pub fn pop_up_menu(&self) -> Option<&PopUpMenu> {
        self.after.as_ref()
    }

    /// Return the value staged for publication, if any.
    #[must_use]
    pub fn format(&self) -> Option<&PopUpMenu> {
        self.pop_up_menu()
    }

    /// Return the value observed when this edit was created, if one existed.
    #[must_use]
    pub fn before(&self) -> Option<&PopUpMenu> {
        self.before.as_ref()
    }

    /// Stage a Pop-Up Menu value without touching package bytes.
    #[must_use]
    pub fn set(mut self, value: PopUpMenu) -> Self {
        self.after = Some(value);
        self
    }

    /// Stage the automatic/no-menu state for the selected cell.
    #[must_use]
    pub fn clear(mut self) -> Self {
        self.after = None;
        self
    }

    /// Alias for [`Self::clear`] used by cell data-format callers.
    #[must_use]
    pub fn reset(self) -> Self {
        self.clear()
    }

    /// Validate and publish the staged edit.
    pub fn commit(self) -> Result<Commit, Error> {
        if self.before == self.after {
            return no_op_commit(self.source, self.path, self.before, self.after);
        }
        rewrite_transaction(self.source, self.path, self.before, self.after)
    }
}

/// A reversible exact-source Pop-Up Menu patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    path: Path,
    before: Option<PopUpMenu>,
    after: Option<PopUpMenu>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("path", &self.path)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the selected semantic cell.
    #[must_use]
    pub const fn path(&self) -> Path {
        self.path
    }

    /// Return the exact source value.
    #[must_use]
    pub const fn before(&self) -> Option<&PopUpMenu> {
        self.before.as_ref()
    }

    /// Return the exact target value.
    #[must_use]
    pub const fn after(&self) -> Option<&PopUpMenu> {
        self.after.as_ref()
    }

    /// Return the target-to-source inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            path: self.path,
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }

    /// Return whether this patch represents an unchanged semantic value.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }
}

/// Content-free publication diagnostics.
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

    /// Whether package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Number of previews removed by the changed operation.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether a complete candidate package was reopened.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// One fully validated Pop-Up Menu publication.
#[must_use = "a Pop-Up Menu commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the validated package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the publication and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Borrow content-free diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the selected cell's Pop-Up Menu, if one is present.
    ///
    /// The strict native owner admits only rooted same-component control
    /// graphs whose table lists, model ownership, and metadata agree exactly.
    pub fn table_cell_pop_up_menu_format<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Option<PopUpMenu>, Error> {
        let target = resolve_cell(self, sheet, table, position)?;
        read_popup(self, target)
    }

    /// Start a selector-first Pop-Up Menu edit.
    pub fn edit_table_cell_pop_up_menu_format<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
        position: CellPosition,
    ) -> Result<Edit<'_>, Error> {
        let target = resolve_cell(self, sheet, table, position)?;
        let before = read_popup(self, target)?;
        Ok(Edit {
            source: self,
            path: Path::Cell {
                sheet: target.sheet_position,
                table: target.table_position,
                position,
            },
            before: before.clone(),
            after: before,
        })
    }

    /// Apply a reversible exact-source Pop-Up Menu patch.
    pub fn apply_table_cell_pop_up_menu_format(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = super::table_headers::rewrite::physical_source(self)
            .map_err(|_| Error::UnsupportedSource)?;
        let source = catalog.__source_owner();
        if !patch.artifacts.authorizes_owner(&source) {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let Path::Cell {
            sheet,
            table,
            position,
        } = patch.path
        else {
            return Err(Error::PatchConflict);
        };
        let current = resolve_cell(
            self,
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )?;
        if read_popup(self, current)? != patch.before {
            return Err(Error::PatchConflict);
        }
        let target_owner = patch.artifacts.target_owner();
        let mut budget = TransactionBudget::new(self);
        let source_catalog = super::table_headers::rewrite::physical_source(self)
            .map_err(|_| Error::UnsupportedSource)?;
        budget.charge_package_source(source_catalog, patch.path)?;
        budget.charge_output(target_owner.as_ref().len(), patch.path)?;
        budget
            .charge_transaction_work(target_owner.as_ref().len().saturating_mul(2), patch.path)?;
        budget.charge_allocations(2, patch.path)?;
        budget.charge_candidate_input_bytes(target_owner.as_ref().len(), patch.path)?;
        let candidate = Package::from_source_owner_with_options(target_owner, self.state.options)
            .map_err(|_| Error::Verification)?;
        let candidate_catalog = super::table_headers::rewrite::physical_source(&candidate)
            .map_err(|_| Error::Verification)?;
        budget.charge_candidate_reopen(candidate_catalog, patch.path)?;
        let candidate_target = resolve_cell(
            &candidate,
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )?;
        let after = read_popup_with_budget(&candidate, candidate_target, &mut budget, true)?;
        if after != patch.after {
            return Err(Error::Verification);
        }
        let touched_components = {
            let target_catalog = super::table_headers::rewrite::physical_source(&candidate)
                .map_err(|_| Error::Verification)?;
            changed_member_count(source_catalog, target_catalog)
        };
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics {
                changed: true,
                touched_components,
                deleted_previews: 0,
                full_reparse_performed: true,
            },
        })
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct CellTarget {
    pub(super) sheet_position: usize,
    pub(super) table_position: usize,
    pub(super) position: CellPosition,
    pub(super) model_identifier: u64,
    pub(super) component_index: usize,
    pub(super) object_index: usize,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
    pub(super) locked: bool,
}

/// One operation-local ledger shared by the rooted catalog, native graph,
/// metadata sidecar, ZIP reassembly, and candidate reopen.  The lower codecs
/// are deliberately given residual ceilings from this ledger; each distinct
/// scan consumes its report exactly once.
#[derive(Debug, Clone, Copy)]
pub(super) struct TransactionBudget {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_entries: usize,
    max_entry_bytes: usize,
    max_total_entry_bytes: usize,
    max_payload_bytes: usize,
    max_payload_objects: usize,
    max_payload_messages: usize,
    max_payload_items: usize,
    max_payload_references: usize,
    max_wire_bytes: usize,
    max_wire_nesting: u32,
    max_wire_reference_bytes: usize,
    max_wire_text_bytes: usize,
    max_wire_fields: usize,
    max_wire_work: usize,
    max_scratch_bytes: usize,
    max_retained_bytes: usize,
    max_compressed_bytes: usize,
    max_candidate_input_bytes: usize,
    max_transaction_work: usize,
    max_allocations: usize,
    remaining_input_bytes: usize,
    remaining_output_bytes: usize,
    remaining_entries: usize,
    remaining_entry_bytes: usize,
    remaining_total_entry_bytes: usize,
    remaining_payload_bytes: usize,
    remaining_payload_objects: usize,
    remaining_payload_messages: usize,
    remaining_payload_items: usize,
    remaining_payload_references: usize,
    remaining_wire_bytes: usize,
    remaining_wire_nesting: u32,
    remaining_wire_reference_bytes: usize,
    remaining_wire_text_bytes: usize,
    remaining_wire_fields: usize,
    remaining_wire_work: usize,
    remaining_scratch_bytes: usize,
    remaining_retained_bytes: usize,
    remaining_compressed_bytes: usize,
    remaining_candidate_input_bytes: usize,
    remaining_transaction_work: usize,
    remaining_allocations: usize,
}

impl TransactionBudget {
    pub(super) fn new(source: &Package) -> Self {
        let archive = source.state.options.archive();
        // Package reassembly uses the physical input ceiling as its output
        // ceiling as well.  Keep the aggregate ZIP-total ceiling separate:
        // a source can fit under max_input while its edited candidate must
        // still be rejected when it grows past that caller-selected bound.
        let max_input_bytes = usize::try_from(archive.max_input_bytes())
            .unwrap_or(usize::MAX)
            .max(1);
        let max_output_bytes = max_input_bytes;
        let max_entries = archive.max_entries().max(1);
        let max_entry_bytes = usize::try_from(archive.max_entry_bytes())
            .unwrap_or(usize::MAX)
            .max(1);
        let max_total_entry_bytes = usize::try_from(archive.max_total_bytes())
            .unwrap_or(usize::MAX)
            .max(1);
        let max_payload_bytes = archive.max_iwa_stream_bytes().max(1);
        let max_payload_objects = source.state.options.semantic().max_objects().max(1);
        let max_payload_references = source.state.options.semantic().max_references().max(1);
        let max_payload_messages = max_payload_objects.saturating_mul(8).max(1);
        let max_payload_items = max_payload_references.saturating_mul(8).max(1);
        let max_wire_bytes = max_payload_bytes;
        let max_wire_nesting = 64;
        let max_wire_reference_bytes = max_payload_references.saturating_mul(16).max(1);
        let max_wire_text_bytes = max_payload_bytes;
        let max_wire_fields = max_wire_bytes.saturating_mul(64).max(1);
        let max_wire_work = max_wire_bytes.saturating_mul(256).max(1);
        let max_scratch_bytes = max_wire_work;
        let max_retained_bytes = max_wire_work;
        let max_compressed_bytes = max_total_entry_bytes;
        let max_candidate_input_bytes = max_input_bytes;
        let max_transaction_work = max_input_bytes.saturating_mul(64).max(1);
        let max_allocations = source
            .state
            .components
            .catalog()
            .len()
            .saturating_mul(128)
            .saturating_add(256)
            .max(256);
        Self {
            max_input_bytes,
            max_output_bytes,
            max_entries,
            max_entry_bytes,
            max_total_entry_bytes,
            max_payload_bytes,
            max_payload_objects,
            max_payload_messages,
            max_payload_items,
            max_payload_references,
            max_wire_bytes,
            max_wire_nesting,
            max_wire_reference_bytes,
            max_wire_text_bytes,
            max_wire_fields,
            max_wire_work,
            max_scratch_bytes,
            max_retained_bytes,
            max_compressed_bytes,
            max_candidate_input_bytes,
            max_transaction_work,
            max_allocations,
            remaining_input_bytes: max_input_bytes,
            remaining_output_bytes: max_output_bytes,
            remaining_entries: max_entries,
            remaining_entry_bytes: max_entry_bytes,
            remaining_total_entry_bytes: max_total_entry_bytes,
            remaining_payload_bytes: max_payload_bytes,
            remaining_payload_objects: max_payload_objects,
            remaining_payload_messages: max_payload_messages,
            remaining_payload_items: max_payload_items,
            remaining_payload_references: max_payload_references,
            remaining_wire_bytes: max_wire_bytes,
            remaining_wire_nesting: max_wire_nesting,
            remaining_wire_reference_bytes: max_wire_reference_bytes,
            remaining_wire_text_bytes: max_wire_text_bytes,
            remaining_wire_fields: max_wire_fields,
            remaining_wire_work: max_wire_work,
            remaining_scratch_bytes: max_scratch_bytes,
            remaining_retained_bytes: max_retained_bytes,
            remaining_compressed_bytes: max_compressed_bytes,
            remaining_candidate_input_bytes: max_candidate_input_bytes,
            remaining_transaction_work: max_transaction_work,
            remaining_allocations: max_allocations,
        }
    }

    /// Construct the aggregate ledger used by the unified scalar-control
    /// route. That route scans the model, data store, tiles, both co-located
    /// lists, every mixed CellSpec, and Metadata under one transaction. The
    /// archive's IWA-stream ceiling is a per-message bound, so summing all of
    /// those distinct reports against it would incorrectly reject valid
    /// packages. The transaction-work ceiling remains the outer aggregate
    /// bound for those scans.
    pub(super) fn for_cell_control(source: &Package) -> Self {
        let mut budget = Self::new(source);
        let aggregate = budget.max_transaction_work;
        budget.max_wire_bytes = aggregate;
        budget.max_wire_fields = aggregate;
        budget.max_wire_work = aggregate;
        budget.remaining_wire_bytes = aggregate;
        budget.remaining_wire_fields = aggregate;
        budget.remaining_wire_work = aggregate;
        budget
    }

    fn charge(
        remaining: &mut usize,
        maximum: usize,
        amount: usize,
        kind: LimitKind,
        path: Path,
    ) -> Result<(), Error> {
        if amount > *remaining {
            let observed = maximum.saturating_sub(*remaining).saturating_add(amount);
            return Err(Error::LimitExceeded {
                kind,
                observed: u64::try_from(observed).unwrap_or(u64::MAX),
                maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
                path,
            });
        }
        *remaining -= amount;
        Ok(())
    }

    pub(super) fn charge_input(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_input_bytes,
            self.max_input_bytes,
            amount,
            LimitKind::InputBytes,
            path,
        )
    }

    /// Charge bytes consumed by reopening a private candidate.  Candidate
    /// ingress has its own counter so source catalog bytes and candidate bytes
    /// are independent traversals rather than an accidental double charge of
    /// the same input ceiling.
    pub(super) fn charge_candidate_input_bytes(
        &mut self,
        amount: usize,
        path: Path,
    ) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_candidate_input_bytes,
            self.max_candidate_input_bytes,
            amount,
            LimitKind::InputBytes,
            path,
        )
    }

    pub(super) fn charge_output(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_output_bytes,
            self.max_output_bytes,
            amount,
            LimitKind::OutputBytes,
            path,
        )
    }

    pub(super) fn charge_entries(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_entries,
            self.max_entries,
            amount,
            LimitKind::Entries,
            path,
        )
    }

    pub(super) fn charge_entry_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_entry_bytes,
            self.max_entry_bytes,
            amount,
            LimitKind::EntryBytes,
            path,
        )
    }

    pub(super) fn charge_total_entry_bytes(
        &mut self,
        amount: usize,
        path: Path,
    ) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_total_entry_bytes,
            self.max_total_entry_bytes,
            amount,
            LimitKind::TotalEntryBytes,
            path,
        )
    }

    pub(super) fn charge_payload_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_payload_bytes,
            self.max_payload_bytes,
            amount,
            LimitKind::PayloadBytes,
            path,
        )
    }

    pub(super) fn charge_payload_objects(
        &mut self,
        amount: usize,
        path: Path,
    ) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_payload_objects,
            self.max_payload_objects,
            amount,
            LimitKind::PayloadObjects,
            path,
        )
    }

    pub(super) fn charge_payload_messages(
        &mut self,
        amount: usize,
        path: Path,
    ) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_payload_messages,
            self.max_payload_messages,
            amount,
            LimitKind::PayloadMessages,
            path,
        )
    }

    pub(super) fn charge_payload_items(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_payload_items,
            self.max_payload_items,
            amount,
            LimitKind::PayloadItems,
            path,
        )
    }

    pub(super) fn charge_payload_references(
        &mut self,
        amount: usize,
        path: Path,
    ) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_payload_references,
            self.max_payload_references,
            amount,
            LimitKind::PayloadReferences,
            path,
        )
    }

    pub(super) fn charge_wire_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_wire_bytes,
            self.max_wire_bytes,
            amount,
            LimitKind::WireBytes,
            path,
        )
    }

    pub(super) fn charge_wire_nesting(&mut self, depth: u32, path: Path) -> Result<(), Error> {
        let observed = self
            .max_wire_nesting
            .saturating_sub(self.remaining_wire_nesting);
        if depth > self.max_wire_nesting {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: u64::from(depth),
                maximum: u64::from(self.max_wire_nesting),
                path,
            });
        }
        // Nesting is a maximum across independent traversals, not an additive
        // byte/work counter.  Keep the deepest observed level as the only
        // debit so a second scan cannot falsely exhaust the depth ceiling.
        self.remaining_wire_nesting = self.max_wire_nesting.saturating_sub(observed.max(depth));
        Ok(())
    }

    pub(super) fn charge_wire_reference_bytes(
        &mut self,
        amount: usize,
        path: Path,
    ) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_wire_reference_bytes,
            self.max_wire_reference_bytes,
            amount,
            LimitKind::WireReferenceBytes,
            path,
        )
    }

    pub(super) fn charge_wire_text_bytes(
        &mut self,
        amount: usize,
        path: Path,
    ) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_wire_text_bytes,
            self.max_wire_text_bytes,
            amount,
            LimitKind::WireTextBytes,
            path,
        )
    }

    pub(super) fn charge_wire_fields(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_wire_fields,
            self.max_wire_fields,
            amount,
            LimitKind::WireFields,
            path,
        )
    }

    pub(super) fn charge_wire_work(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_wire_work,
            self.max_wire_work,
            amount,
            LimitKind::WireWork,
            path,
        )
    }

    pub(super) fn charge_scratch_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_scratch_bytes,
            self.max_scratch_bytes,
            amount,
            LimitKind::ScratchBytes,
            path,
        )
    }

    pub(super) fn charge_retained_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_retained_bytes,
            self.max_retained_bytes,
            amount,
            LimitKind::RetainedBytes,
            path,
        )
    }

    pub(super) fn charge_compressed_bytes(
        &mut self,
        amount: usize,
        path: Path,
    ) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_compressed_bytes,
            self.max_compressed_bytes,
            amount,
            LimitKind::CompressedBytes,
            path,
        )
    }

    pub(super) fn charge_transaction_work(
        &mut self,
        amount: usize,
        path: Path,
    ) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_transaction_work,
            self.max_transaction_work,
            amount,
            LimitKind::TransactionWork,
            path,
        )
    }

    pub(super) fn charge_allocations(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_allocations,
            self.max_allocations,
            amount,
            LimitKind::Allocations,
            path,
        )
    }

    pub(super) fn preflight_reassembly(
        &mut self,
        requirements: ReassemblyExecutionRequirements,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_output(requirements.output_bytes(), path)?;
        self.charge_allocations(requirements.allocations(), path)?;
        self.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
        self.charge_retained_bytes(requirements.retained_bytes(), path)?;
        self.charge_transaction_work(requirements.output_bytes(), path)
    }

    pub(super) fn charge_package_source(
        &mut self,
        catalog: &litchi_iwa_archive::SourceCatalog,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_input(catalog.source_bytes().len(), path)?;
        let mut total = 0usize;
        for entry in catalog.package().iter() {
            self.charge_entries(1, path)?;
            self.charge_entry_bytes(entry.data().len(), path)?;
            total = total.saturating_add(entry.data().len());
        }
        self.charge_total_entry_bytes(total, path)?;
        Ok(())
    }

    /// Charge the physical and parsed work performed by reopening a private
    /// candidate.  Candidate input bytes are precharged separately, before
    /// ZIP execution; this method accounts the distinct member/component
    /// scans after the candidate catalog exists and therefore never debits
    /// that input ceiling a second time.
    pub(super) fn charge_candidate_reopen(
        &mut self,
        catalog: &litchi_iwa_archive::SourceCatalog,
        path: Path,
    ) -> Result<(), Error> {
        let mut total_entry_bytes = 0usize;
        for entry in catalog.package().iter() {
            self.charge_entries(1, path)?;
            self.charge_entry_bytes(entry.data().len(), path)?;
            self.charge_compressed_bytes(entry.data().len(), path)?;
            total_entry_bytes = total_entry_bytes.saturating_add(entry.data().len());
        }
        self.charge_total_entry_bytes(total_entry_bytes, path)?;
        for component in catalog.components().iter() {
            let archive = component.archive();
            self.charge_payload_objects(archive.objects.len(), path)?;
            let mut messages = 0usize;
            let mut references = 0usize;
            for object in &archive.objects {
                messages = messages.saturating_add(object.messages.len());
                for info in &object.archive_info.message_infos {
                    references = references.saturating_add(info.object_references.len());
                    references = references.saturating_add(
                        info.field_infos
                            .iter()
                            .map(|field| field.object_references.len())
                            .sum::<usize>(),
                    );
                }
            }
            self.charge_payload_messages(messages, path)?;
            self.charge_payload_references(references, path)?;
            self.charge_payload_items(messages.saturating_add(archive.objects.len()), path)?;
        }
        Ok(())
    }

    pub(super) fn charge_archive(
        &mut self,
        archive: &Archive,
        decoded_bytes: usize,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_payload_bytes(decoded_bytes, path)?;
        self.charge_payload_objects(archive.objects.len(), path)?;
        let mut messages = 0usize;
        let mut references = 0usize;
        for object in &archive.objects {
            messages = messages.saturating_add(object.messages.len());
            for info in &object.archive_info.message_infos {
                references = references.saturating_add(info.object_references.len());
                references = references.saturating_add(
                    info.field_infos
                        .iter()
                        .map(|field| field.object_references.len())
                        .sum::<usize>(),
                );
            }
        }
        self.charge_payload_messages(messages, path)?;
        self.charge_payload_references(references, path)?;
        self.charge_payload_items(messages.saturating_add(archive.objects.len()), path)
    }

    pub(super) fn residual_storage_options(&self, source: &[u8]) -> storage_codec::DecodeOptions {
        storage_codec::DecodeOptions::new(
            self.remaining_wire_bytes.min(source.len().max(1)),
            self.remaining_wire_work.max(1),
            self.remaining_wire_fields.max(1),
            self.max_wire_nesting,
            self.remaining_payload_references.max(1),
            self.remaining_payload_references.max(1),
        )
    }

    pub(super) fn residual_popup_options(&self, source: &[u8]) -> popup_codec::DecodeOptions {
        let bytes = source.len().max(1);
        popup_codec::DecodeOptions::new(
            self.remaining_wire_bytes.min(bytes),
            self.remaining_output_bytes
                .min(bytes.saturating_mul(2).max(1)),
            self.remaining_wire_fields.max(1),
            self.remaining_wire_work.max(1),
            self.max_wire_nesting,
            self.remaining_payload_references.max(1),
            self.remaining_payload_items.max(1),
            self.remaining_wire_text_bytes.min(bytes).max(1),
        )
    }

    pub(super) fn charge_storage_decode_report(
        &mut self,
        report: storage_codec::DecodeReport,
        include_input: bool,
        path: Path,
    ) -> Result<(), Error> {
        if include_input {
            self.charge_wire_bytes(report.source_bytes(), path)?;
        }
        self.charge_wire_fields(report.fields(), path)?;
        self.charge_wire_work(report.work_bytes(), path)?;
        self.charge_wire_nesting(report.max_depth(), path)?;
        self.charge_payload_references(report.references(), path)?;
        self.charge_wire_reference_bytes(report.reference_bytes(), path)?;
        self.charge_wire_text_bytes(report.text_bytes(), path)
    }

    pub(super) fn residual_control_options_for_len(
        &self,
        source_len: usize,
    ) -> control_codec::DecodeOptions {
        let bytes = source_len.max(256);
        control_codec::DecodeOptions::new(
            self.remaining_wire_bytes.min(bytes).max(1),
            self.remaining_output_bytes
                .min(bytes.saturating_mul(2).max(1))
                .max(1),
            self.remaining_wire_fields
                .min(bytes.saturating_mul(8).max(1)),
            self.remaining_wire_work
                .min(bytes.saturating_mul(16).max(1)),
            self.max_wire_nesting,
            self.remaining_payload_references
                .min(bytes.saturating_mul(2).max(1)),
            self.remaining_payload_items
                .min(bytes.saturating_mul(2).max(1)),
            self.remaining_wire_text_bytes.min(bytes).max(1),
        )
    }

    /// Residual storage policy for a prepared list rewrite whose candidate
    /// may be larger than its source message. The strict storage API uses one
    /// byte ceiling for both ingress and candidate verification, so the
    /// ordinary source-sized decode policy is insufficient for an append.
    pub(super) fn residual_storage_rewrite_options(
        &self,
        source: &[u8],
    ) -> storage_codec::DecodeOptions {
        let bytes = source.len().max(1).saturating_mul(8);
        storage_codec::DecodeOptions::new(
            self.remaining_wire_bytes.min(bytes).max(1),
            self.remaining_wire_work.max(1),
            self.remaining_wire_fields.max(1),
            self.max_wire_nesting,
            self.remaining_payload_references.max(1),
            self.remaining_payload_references.max(1),
        )
    }
}

fn no_op_commit(
    source: &Package,
    path: Path,
    before: Option<PopUpMenu>,
    after: Option<PopUpMenu>,
) -> Result<Commit, Error> {
    let catalog = super::table_headers::rewrite::physical_source(source)
        .map_err(|_| Error::UnsupportedSource)?;
    let owner = catalog.__source_owner();
    Ok(Commit {
        package: source.snapshot(),
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(owner.clone(), owner),
            path,
            before,
            after,
        },
        diagnostics: Diagnostics::unchanged(),
    })
}

/// First lifecycle checkpoint: rewrite a same-component document graph into a
/// private candidate, then reopen it before publication.  The metadata sidecar
/// is updated in the same ZIP reassembly; later hardening can replace the
/// generated-free list surgery below with the prepared storage transition
/// without changing the public transaction or patch contract.
fn rewrite_transaction(
    source: &Package,
    path: Path,
    before: Option<PopUpMenu>,
    after: Option<PopUpMenu>,
) -> Result<Commit, Error> {
    let Path::Cell {
        sheet,
        table,
        position,
    } = path
    else {
        return Err(Error::InvalidSource { path });
    };
    let target = resolve_cell(
        source,
        SheetSelector::index(sheet),
        TableSelector::index(table),
        position,
    )?;
    if target.locked {
        return Err(Error::TableLocked { path });
    }
    reject_cross_component_write(source, target, path)?;
    let mut budget = TransactionBudget::new(source);
    let catalog = super::table_headers::rewrite::physical_source(source)
        .map_err(|_| Error::UnsupportedSource)?;
    if !catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    budget.charge_package_source(catalog, path)?;
    let observed = read_popup_with_budget(source, target, &mut budget, false)?;
    if observed != before {
        return Err(Error::PatchConflict);
    }
    let metadata_source = popup_metadata::strict_source(source)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    budget.charge_wire_bytes(metadata_source.payload.len(), path)?;
    let metadata_options = popup_metadata_options(metadata_source.payload.len(), 1, &budget);
    let metadata_facts = popup_metadata::inspect(source, metadata_options)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let metadata_report = metadata_facts.report();
    budget.charge_wire_fields(metadata_report.fields(), path)?;
    budget.charge_wire_work(metadata_report.work_bytes(), path)?;
    budget.charge_wire_nesting(metadata_report.max_depth(), path)?;
    budget.charge_payload_items(metadata_report.components_scanned(), path)?;
    budget.charge_payload_references(metadata_report.references_scanned(), path)?;
    budget.charge_transaction_work(
        metadata_report
            .input_bytes()
            .saturating_add(metadata_report.output_bytes()),
        path,
    )?;
    if metadata_facts.has_physical_alias() {
        return Err(Error::UnsupportedDependency { path });
    }
    metadata_facts
        .require_single_current_component(core::slice::from_ref(&target.component_index))
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let fresh = metadata_facts
        .allocate_identifiers(1)
        .map_err(|error| map_popup_metadata_error(error, path))?
        .into_iter()
        .next()
        .ok_or(Error::InvalidSource { path })?;
    let native = rewrite_native_entry(
        source,
        target,
        after.as_ref(),
        fresh.identifier,
        path,
        &metadata_facts,
        &mut budget,
    )?;
    let metadata_payload = rewrite_popup_metadata(
        source,
        &metadata_facts,
        target.component_index,
        fresh,
        native.added_model,
        &native.removed_models,
        path,
        &mut budget,
    )?;
    let metadata_bytes = pack_metadata_payload(
        source,
        metadata_facts.route(),
        metadata_payload,
        path,
        &mut budget,
    )?;
    let previews = super::table_headers::rewrite::root_preview_deletions(catalog)
        .map_err(|_| Error::InvalidSource { path })?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(native.member_edits.len().saturating_add(1))
        .map_err(|_| Error::Allocation {
            amount: native.member_edits.len().saturating_add(1),
            path,
        })?;
    for edit in &native.member_edits {
        edits.push(EntryEdit::new(
            edit.member_name.as_str(),
            edit.member_bytes.as_slice(),
        ));
    }
    edits.push(EntryEdit::new(super::metadata::ENTRY_NAME, &metadata_bytes));
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &previews, catalog.limits())
        .map_err(|_| Error::Verification)?;
    let requirements = prepared.execution_requirements();
    // Candidate reassembly and the subsequent full Package reopen are both
    // precharged before the final ZIP output allocation.  This owner retains a
    // staged-private policy: native/metadata candidates are private until the
    // ZIP candidate and semantic readback succeed.
    budget.preflight_reassembly(requirements, path)?;
    budget.charge_transaction_work(
        requirements
            .output_bytes()
            .saturating_mul(2)
            .saturating_add(source.state.components.catalog().len().saturating_mul(1024)),
        path,
    )?;
    budget.charge_candidate_input_bytes(requirements.output_bytes(), path)?;
    let bytes = prepared
        .execute(requirements.exact_limits())
        .map_err(|_| Error::Verification)?;
    let candidate = Package::from_owned_bytes_with_options(bytes, source.state.options)
        .map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(&candidate)
        .map_err(|_| Error::Verification)?;
    budget.charge_candidate_reopen(candidate_catalog, path)?;
    let native_member_names = native
        .member_edits
        .iter()
        .map(|edit| edit.member_name.as_str())
        .collect::<Vec<_>>();
    verify_package_locality_for_members(source, &candidate, &native_member_names)?;
    let reread = read_popup_with_budget(
        &candidate,
        resolve_cell(
            &candidate,
            SheetSelector::index(sheet),
            TableSelector::index(table),
            position,
        )?,
        &mut budget,
        true,
    )?;
    if reread != after {
        return Err(Error::Verification);
    }
    let source_owner = catalog.__source_owner();
    let target_owner = super::table_headers::rewrite::physical_source(&candidate)
        .map_err(|_| Error::Verification)?
        .__source_owner();
    let touched_components = native
        .member_edits
        .iter()
        .map(|edit| edit.component_index)
        .chain(std::iter::once(metadata_facts.route().component_index))
        .collect::<std::collections::BTreeSet<_>>()
        .len();
    Ok(Commit {
        package: candidate,
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(source_owner, target_owner),
            path,
            before,
            after,
        },
        diagnostics: Diagnostics {
            changed: true,
            touched_components,
            deleted_previews: previews.len(),
            full_reparse_performed: true,
        },
    })
}

/// Verify package-level locality for a native transaction that rewrites one
/// or more physical members.  All non-authorized members must remain exact,
/// including opaque/versioned entries and metadata; preview deletions are
/// intentionally handled by the caller because their policy is operation
/// specific.
pub(super) fn verify_package_locality_for_members(
    source: &Package,
    candidate: &Package,
    native_members: &[&str],
) -> Result<(), Error> {
    let source_catalog =
        super::table_headers::rewrite::physical_source(source).map_err(|_| Error::Verification)?;
    let candidate_catalog = super::table_headers::rewrite::physical_source(candidate)
        .map_err(|_| Error::Verification)?;
    let allowed = |name: &str| {
        native_members.contains(&name)
            || name == super::metadata::ENTRY_NAME
            || name.starts_with("preview")
    };
    for entry in source_catalog.package().iter() {
        if allowed(entry.name()) {
            continue;
        }
        let counterpart = candidate_catalog
            .package()
            .iter()
            .find(|candidate_entry| candidate_entry.name() == entry.name())
            .ok_or(Error::Verification)?;
        if counterpart.data() != entry.data() {
            return Err(Error::Verification);
        }
    }
    for entry in candidate_catalog.package().iter() {
        if allowed(entry.name()) {
            continue;
        }
        let counterpart = source_catalog
            .package()
            .iter()
            .find(|source_entry| source_entry.name() == entry.name())
            .ok_or(Error::Verification)?;
        if counterpart.data() != entry.data() {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

struct NativeRewrite {
    member_edits: Vec<popup_native::NativeMemberEdit>,
    added_model: Option<u64>,
    removed_models: Vec<u64>,
}

fn rewrite_native_entry(
    source: &Package,
    target: CellTarget,
    menu: Option<&PopUpMenu>,
    fresh_identifier: u64,
    path: Path,
    metadata_facts: &popup_metadata::RegistryFacts<'_>,
    budget: &mut TransactionBudget,
) -> Result<NativeRewrite, Error> {
    const POPUP_FORMAT_PAYLOAD: &[u8] = &[0x08, 0x84, 0x02];

    let component = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let archive = component.archive();
    let model = archive
        .objects
        .get(target.object_index)
        .and_then(|object| object.messages.get(target.message_index))
        .filter(|message| message.type_ == target.message_type && message.type_ == 6_001)
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_wire_bytes(model.data.len(), path)?;
    let options = budget.residual_storage_options(&model.data);
    let (model_snapshot, model_report) =
        storage_codec::decode_table_model_with_report(&model.data, options)
            .map_err(|_| Error::InvalidSource { path })?;
    budget.charge_storage_decode_report(model_report, false, path)?;
    let (store, store_report) = storage_codec::decode_data_store_with_report(
        model_snapshot.base_data_store(),
        budget.residual_storage_options(model_snapshot.base_data_store()),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    budget.charge_storage_decode_report(store_report, true, path)?;
    let control_table_identifier = store
        .control_cell_spec_table()
        .map(|reference| reference.identifier())
        .ok_or(Error::InvalidSource { path })?;
    let format_table_identifier = store
        .format_table()
        .map(|reference| reference.identifier())
        .ok_or(Error::InvalidSource { path })?;
    let mut tiles = TileCollector::default();
    let (tile_storage, tile_report) = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        budget.residual_storage_options(store.tiles()),
        &mut tiles,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    budget.charge_storage_decode_report(tile_report, true, path)?;
    let tile_size = tile_storage
        .tile_size()
        .ok_or(Error::InvalidSource { path })?;
    let tile_id = target_row_tile(tile_size, target.position.row());
    let mut tile_matches = tiles
        .tiles
        .iter()
        .filter(|tile| tile.0 == tile_id)
        .map(|tile| tile.1);
    let tile_identifier = tile_matches.next().ok_or(Error::CellNotFound)?;
    if tile_matches.next().is_some() {
        return Err(Error::InvalidSource { path });
    }

    let item_texts = menu.map(|value| {
        value
            .items()
            .iter()
            .map(|item| item.as_str())
            .collect::<Vec<_>>()
    });
    let desired =
        menu.zip(item_texts.as_deref())
            .map(|(value, items)| popup_native::NativePopUpValue {
                items,
                starts_with_first: value.initial_selection()
                    == crate::cell::data_format::pop_up_menu::InitialSelection::FirstItem,
                first_item_string_identifier: None,
            });
    let physical = source
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource)?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let native_input = popup_native::NativePopUpInput {
        archive,
        component_index: target.component_index,
        model_identifier: target.model_identifier,
        tile_identifier,
        tile_row: target.position.row(),
        tile_column: target.position.column(),
        control_table_identifier,
        format_table_identifier,
        member_name: component.name(),
        format_payload: POPUP_FORMAT_PAYLOAD,
        new_popup_model_identifier: Some(fresh_identifier),
        desired,
        limits: archive_limits,
        path,
    };
    let copy_on_write_candidates =
        popup_native::existing_popup_model_identifiers(native_input, budget, path)
            .map_err(|error| map_native_error(error, path))?;
    budget.charge_payload_items(copy_on_write_candidates.len().max(1), path)?;
    budget.charge_payload_references(copy_on_write_candidates.len(), path)?;
    budget.charge_transaction_work(
        copy_on_write_candidates
            .len()
            .saturating_add(1)
            .saturating_mul(archive.objects.len().max(1)),
        path,
    )?;
    // Every existing object whose archive message is mutated or whose
    // ownership is used to decide a cull must have exactly one current UUID
    // owner in this component.  Checking only popup models leaves tile/list
    // aliases and versioned list owners outside the transaction proof.
    for identifier in [
        target.model_identifier,
        tile_identifier,
        control_table_identifier,
        format_table_identifier,
    ] {
        metadata_facts
            .require_current_uuid(target.component_index, identifier)
            .map_err(|error| map_popup_metadata_error(error, path))?;
    }
    for identifier in &copy_on_write_candidates {
        metadata_facts
            .require_current_uuid(target.component_index, *identifier)
            .map_err(|error| map_popup_metadata_error(error, path))?;
        for (component_index, component) in source.state.components.catalog().iter().enumerate() {
            let found = popup_native::archive_has_popup_reference_strict(
                component.archive(),
                *identifier,
                archive_limits,
            )
            .map_err(|error| map_native_error(error, path))?;
            if component_index == target.component_index {
                // The selected archive necessarily contains the rooted
                // control edge; only unknown metadata is actionable here.
                continue;
            }
            if found {
                // A model referenced by another package member cannot be
                // safely removed or repointed by this same-component owner;
                // reject before the native candidate is allocated.
                return Err(Error::UnsupportedDependency { path });
            }
        }
    }
    let output = popup_native::rewrite_native_popup_menu(native_input, budget)
        .map_err(|error| map_native_error(error, path))?;
    let mut member_edits = output.member_edits.edits;
    budget.charge_allocations(member_edits.len(), path)?;
    for edit in &mut member_edits {
        let maximum_compressed = SnappyStream::maximum_compressed_len(edit.member_bytes.len())
            .map_err(|_| Error::InvalidSource { path })?;
        budget.charge_compressed_bytes(maximum_compressed, path)?;
        budget.charge_transaction_work(
            edit.member_bytes.len().saturating_add(maximum_compressed),
            path,
        )?;
        edit.member_bytes = SnappyStream::compress(&edit.member_bytes)
            .map_err(|_| Error::InvalidSource { path })?;
    }
    let added_model = match output.added_object_identifiers.as_slice() {
        [] => None,
        [identifier] => Some(*identifier),
        _ => return Err(Error::UnsupportedDependency { path }),
    };
    Ok(NativeRewrite {
        member_edits,
        added_model,
        removed_models: output.removed_object_identifiers,
    })
}

fn map_native_error(error: popup_native::NativePopUpError, path: Path) -> Error {
    match error {
        popup_native::NativePopUpError::UnsupportedDependency => {
            Error::UnsupportedDependency { path }
        },
        popup_native::NativePopUpError::Allocation => Error::Allocation { amount: 1, path },
        popup_native::NativePopUpError::Limit => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: 1,
            maximum: 0,
            path,
        },
        popup_native::NativePopUpError::InvalidSource
        | popup_native::NativePopUpError::Codec
        | popup_native::NativePopUpError::Archive => Error::InvalidSource { path },
    }
}

fn popup_metadata_options(
    payload_bytes: usize,
    additions: usize,
    budget: &TransactionBudget,
) -> MetadataRewriteOptions {
    MetadataRewriteOptions::new(
        payload_bytes.min(budget.remaining_transaction_work).max(1),
        payload_bytes
            .saturating_mul(2)
            .min(budget.remaining_transaction_work)
            .max(1),
        payload_bytes
            .saturating_mul(16)
            .min(budget.remaining_wire_fields)
            .max(1),
        payload_bytes
            .saturating_mul(64)
            .min(budget.remaining_wire_work)
            .max(1),
        64,
        budget.remaining_payload_items.max(1),
        budget.remaining_payload_references.max(1),
        additions.max(1).min(budget.remaining_payload_items.max(1)),
    )
}

fn rewrite_popup_metadata(
    _source: &Package,
    facts: &popup_metadata::RegistryFacts<'_>,
    component_index: usize,
    fresh: popup_metadata::FreshIdentifier,
    added_model: Option<u64>,
    removed_models: &[u64],
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    if added_model.is_some() && !removed_models.is_empty() {
        return Err(Error::UnsupportedDependency { path });
    }

    let selectors = facts
        .selectors_for_components(core::slice::from_ref(&component_index))
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let save_tokens = SaveTokenBatch::new(&selectors);
    let options = popup_metadata_options(
        facts.payload().len(),
        usize::from(added_model.is_some()),
        budget,
    );

    if let Some(identifier) = added_model {
        if identifier != fresh.identifier {
            return Err(Error::InvalidSource { path });
        }
        let additions = [facts
            .uuid_addition(component_index, fresh)
            .map_err(|error| map_popup_metadata_error(error, path))?];
        let batch = MetadataBatch::new(facts.last_object_identifier(), identifier, &additions, &[]);
        let prepared = popup_metadata::prepare_additions(
            facts,
            AdditionSaveTokenBatch::new(batch, save_tokens),
            options,
        )
        .map_err(|error| map_popup_metadata_error(error, path))?;
        let report = prepared.prepare_report();
        budget.charge_wire_fields(report.fields(), path)?;
        budget.charge_wire_work(report.work_bytes(), path)?;
        budget.charge_wire_nesting(report.max_depth(), path)?;
        budget.charge_payload_items(report.components_scanned(), path)?;
        budget.charge_payload_references(report.references_scanned(), path)?;
        budget.charge_transaction_work(
            report.input_bytes().saturating_add(report.output_bytes()),
            path,
        )?;
        let requirements = prepared.execution_requirements();
        budget.charge_wire_fields(requirements.fields(), path)?;
        budget.charge_wire_work(requirements.work_bytes(), path)?;
        budget.charge_payload_items(requirements.components(), path)?;
        budget.charge_payload_references(requirements.references(), path)?;
        budget.charge_allocations(requirements.allocations(), path)?;
        budget.charge_retained_bytes(requirements.retained_bytes(), path)?;
        budget.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
        budget.charge_transaction_work(requirements.output_bytes(), path)?;
        let limits = requirements.exact_limits();
        return prepared
            .execute(limits)
            .map(|output| output.into_bytes())
            .map_err(|error| {
                map_popup_metadata_error(popup_metadata::map_rewrite_error(error), path)
            });
    }

    if !removed_models.is_empty() {
        let mut removals = Vec::new();
        removals
            .try_reserve_exact(removed_models.len())
            .map_err(|_| Error::Allocation {
                amount: removed_models.len(),
                path,
            })?;
        for &identifier in removed_models {
            removals.push(
                facts
                    .uuid_removal(component_index, identifier)
                    .map_err(|error| map_popup_metadata_error(error, path))?,
            );
        }
        let batch = RemovalBatch::new(facts.last_object_identifier(), &removals, &[], &[]);
        let prepared = popup_metadata::prepare_removals(
            facts,
            RemovalSaveTokenBatch::new(batch, save_tokens),
            options,
        )
        .map_err(|error| map_popup_metadata_error(error, path))?;
        let report = prepared.prepare_report();
        budget.charge_wire_fields(report.fields(), path)?;
        budget.charge_wire_work(report.work_bytes(), path)?;
        budget.charge_wire_nesting(report.max_depth(), path)?;
        budget.charge_payload_items(report.components_scanned(), path)?;
        budget.charge_payload_references(report.references_scanned(), path)?;
        budget.charge_transaction_work(
            report.input_bytes().saturating_add(report.output_bytes()),
            path,
        )?;
        let requirements = prepared.execution_requirements();
        budget.charge_wire_fields(requirements.fields(), path)?;
        budget.charge_wire_work(requirements.work_bytes(), path)?;
        budget.charge_payload_items(requirements.components(), path)?;
        budget.charge_payload_references(requirements.references(), path)?;
        budget.charge_allocations(requirements.allocations(), path)?;
        budget.charge_retained_bytes(requirements.retained_bytes(), path)?;
        budget.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
        budget.charge_transaction_work(requirements.output_bytes(), path)?;
        let limits = requirements.exact_limits();
        return prepared
            .execute(limits)
            .map(|output| output.into_bytes())
            .map_err(|error| {
                map_popup_metadata_error(popup_metadata::map_rewrite_error(error), path)
            });
    }

    let prepared = popup_metadata::prepare_save_tokens(facts, save_tokens, options)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let report = prepared.prepare_report();
    budget.charge_wire_fields(report.fields(), path)?;
    budget.charge_wire_work(report.work_bytes(), path)?;
    budget.charge_wire_nesting(report.max_depth(), path)?;
    budget.charge_payload_items(report.components_scanned(), path)?;
    budget.charge_payload_references(report.references_scanned(), path)?;
    budget.charge_transaction_work(
        report.input_bytes().saturating_add(report.output_bytes()),
        path,
    )?;
    let requirements = prepared.execution_requirements();
    budget.charge_wire_fields(requirements.fields(), path)?;
    budget.charge_wire_work(requirements.work_bytes(), path)?;
    budget.charge_payload_items(requirements.components(), path)?;
    budget.charge_payload_references(requirements.references(), path)?;
    budget.charge_allocations(requirements.allocations(), path)?;
    budget.charge_retained_bytes(requirements.retained_bytes(), path)?;
    budget.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
    budget.charge_transaction_work(requirements.output_bytes(), path)?;
    let limits = requirements.exact_limits();
    prepared
        .execute(limits)
        .map(|output| output.into_bytes())
        .map_err(|error| map_popup_metadata_error(popup_metadata::map_rewrite_error(error), path))
}

/// Rewrite only the root and save tokens for an already-validated set of
/// current native members.  Scalar cell controls do not allocate native
/// objects, but a split graph still mutates its tile and list members; their
/// current component tokens must advance atomically with the ZIP edits.
pub(super) fn rewrite_component_save_tokens(
    source: &Package,
    component_indices: &[usize],
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let metadata_source = popup_metadata::strict_source(source)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    budget.charge_wire_bytes(metadata_source.payload.len(), path)?;
    let options = popup_metadata_options(metadata_source.payload.len(), 0, budget);
    let facts = popup_metadata::inspect(source, options)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let report = facts.report();
    budget.charge_wire_fields(report.fields(), path)?;
    budget.charge_wire_work(report.work_bytes(), path)?;
    budget.charge_wire_nesting(report.max_depth(), path)?;
    budget.charge_payload_items(report.components_scanned(), path)?;
    budget.charge_payload_references(report.references_scanned(), path)?;
    budget.charge_allocations(report.allocations(), path)?;
    budget.charge_scratch_bytes(report.scratch_bytes(), path)?;
    budget.charge_retained_bytes(report.retained_bytes(), path)?;
    budget.charge_transaction_work(
        report.input_bytes().saturating_add(report.output_bytes()),
        path,
    )?;
    if facts.has_physical_alias() {
        return Err(Error::UnsupportedDependency { path });
    }
    let selectors = facts
        .selectors_for_components(component_indices)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let save_tokens = SaveTokenBatch::new(&selectors);
    let prepared = popup_metadata::prepare_save_tokens(&facts, save_tokens, options)
        .map_err(|error| map_popup_metadata_error(error, path))?;
    let prepare_report = prepared.prepare_report();
    budget.charge_wire_fields(prepare_report.fields(), path)?;
    budget.charge_wire_work(prepare_report.work_bytes(), path)?;
    budget.charge_wire_nesting(prepare_report.max_depth(), path)?;
    budget.charge_payload_items(prepare_report.components_scanned(), path)?;
    budget.charge_payload_references(prepare_report.references_scanned(), path)?;
    budget.charge_transaction_work(
        prepare_report
            .input_bytes()
            .saturating_add(prepare_report.output_bytes()),
        path,
    )?;
    let requirements = prepared.execution_requirements();
    budget.charge_wire_fields(requirements.fields(), path)?;
    budget.charge_wire_work(requirements.work_bytes(), path)?;
    budget.charge_payload_items(requirements.components(), path)?;
    budget.charge_payload_references(requirements.references(), path)?;
    budget.charge_allocations(requirements.allocations(), path)?;
    budget.charge_retained_bytes(requirements.retained_bytes(), path)?;
    budget.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
    budget.charge_transaction_work(requirements.output_bytes(), path)?;
    let payload = prepared
        .execute(requirements.exact_limits())
        .map(|output| output.into_bytes())
        .map_err(|error| {
            map_popup_metadata_error(popup_metadata::map_rewrite_error(error), path)
        })?;
    pack_metadata_payload(source, metadata_source.route, payload, path, budget)
}

fn pack_metadata_payload(
    source: &Package,
    route: super::metadata::MessageRoute,
    payload: Vec<u8>,
    path: Path,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let component = source
        .state
        .components
        .catalog()
        .get_index(route.component_index)
        .ok_or(Error::InvalidSource { path })?;
    if component.name() != super::metadata::ENTRY_NAME {
        return Err(Error::InvalidSource { path });
    }
    let physical = source
        .state
        .components
        .physical()
        .ok_or(Error::UnsupportedSource)?;
    let entry = physical
        .package()
        .iter()
        .find(|entry| entry.name() == component.name())
        .ok_or(Error::InvalidSource { path })?;
    // `charge_package_source` already charged this borrowed Metadata.iwa
    // member.  The pack phase gets its own archive/payload/work charges
    // below, but must not debit the physical per-entry ceiling twice.
    budget.charge_transaction_work(entry.data().len().saturating_mul(4), path)?;
    budget.charge_allocations(2, path)?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical
            .limits()
            .snappy_limits()
            .map_err(|_| Error::InvalidSource { path })?,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let source_payload_bytes = stream.as_bytes().len();
    let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
        .map_err(|_| Error::InvalidSource { path })?;
    budget.charge_archive(&archive, source_payload_bytes, path)?;
    let message = archive
        .objects
        .get_mut(route.object_index)
        .and_then(|object| object.messages.get_mut(route.message_index))
        .filter(|message| message.type_ == super::metadata::MESSAGE_TYPE)
        .ok_or(Error::InvalidSource { path })?;
    let old_message_bytes = message.data.len();
    let archive_upper_bound = source_payload_bytes
        .saturating_sub(old_message_bytes)
        .saturating_add(payload.len());
    let maximum_compressed = SnappyStream::maximum_compressed_len(archive_upper_bound)
        .map_err(|_| Error::InvalidSource { path })?;
    budget.charge_payload_bytes(archive_upper_bound, path)?;
    budget.charge_compressed_bytes(maximum_compressed, path)?;
    budget.charge_transaction_work(archive_upper_bound.saturating_add(maximum_compressed), path)?;
    message.data = payload;
    let archive_bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(|_| Error::InvalidSource { path })?;
    let compressed =
        SnappyStream::compress(&archive_bytes).map_err(|_| Error::InvalidSource { path })?;
    if archive_bytes.len() > archive_upper_bound || compressed.len() > maximum_compressed {
        return Err(Error::LimitExceeded {
            kind: LimitKind::PayloadBytes,
            observed: u64::try_from(archive_bytes.len().max(compressed.len())).unwrap_or(u64::MAX),
            maximum: u64::try_from(archive_upper_bound.max(maximum_compressed)).unwrap_or(u64::MAX),
            path,
        });
    }
    Ok(compressed)
}

fn map_popup_metadata_error(error: popup_metadata::MetadataError, path: Path) -> Error {
    match error.kind {
        popup_metadata::FailureKind::Unsupported => Error::UnsupportedDependency { path },
        popup_metadata::FailureKind::Limit => Error::LimitExceeded {
            kind: LimitKind::WireWork,
            observed: u64::try_from(error.observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(error.maximum).unwrap_or(u64::MAX),
            path,
        },
        popup_metadata::FailureKind::Allocation => Error::Allocation {
            amount: error.allocation,
            path,
        },
        popup_metadata::FailureKind::MissingRoute
        | popup_metadata::FailureKind::AmbiguousRoute
        | popup_metadata::FailureKind::InvalidSource
        | popup_metadata::FailureKind::Conflict
        | popup_metadata::FailureKind::VersionedOwnership => Error::InvalidSource { path },
    }
}

pub(super) fn resolve_cell<'sheet, 'table>(
    source: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    table: impl Into<TableSelector<'table>>,
    position: CellPosition,
) -> Result<CellTarget, Error> {
    let sheet_selector = sheet.into();
    let selected_sheet = source
        .state
        .document
        .sheet(sheet_selector)
        .map_err(|_| Error::InvalidSource {
            path: Path::Package,
        })?
        .ok_or(Error::SheetNotFound)?;
    let table_selector = table.into();
    let table_position = match table_selector {
        TableSelector::Index(index) => selected_sheet.tables().nth(index).map(|_| index),
        TableSelector::Name(name) => {
            let mut matches = selected_sheet
                .tables()
                .enumerate()
                .filter(|(_, table)| table.name() == name);
            let first = matches.next().map(|(index, _)| index);
            if matches.next().is_some() {
                return Err(Error::InvalidSource {
                    path: Path::Cell {
                        sheet: selected_sheet.index(),
                        table: 0,
                        position,
                    },
                });
            }
            first
        },
    }
    .ok_or(Error::TableNotFound)?;
    let native = super::table_headers::resolve::resolve_target(
        source,
        selected_sheet.index(),
        table_position,
    )
    .map_err(|_| Error::InvalidSource {
        path: Path::Cell {
            sheet: selected_sheet.index(),
            table: table_position,
            position,
        },
    })?;
    if position.row() >= native.rows || position.column() >= native.columns {
        return Err(Error::CellNotFound);
    }
    Ok(CellTarget {
        sheet_position: selected_sheet.index(),
        table_position,
        position,
        model_identifier: native.model_identifier,
        component_index: native.component_index,
        object_index: native.object_index,
        message_index: native.message_index,
        message_type: native.message_type,
        locked: native.locked == crate::table::lock::State::Locked,
    })
}

pub(super) fn read_popup(source: &Package, target: CellTarget) -> Result<Option<PopUpMenu>, Error> {
    let mut authority = None;
    read_popup_with_policy(source, target, false, &mut authority)
}

/// Read the complete archive-free control sum for one rooted cell.
///
/// Pop-Up Menu remains on the audited popup path; the other four controls use
/// the neutral control codec for their `CellSpecArchive`/format payloads.  The
/// row projection deliberately resolves packed offsets, so a control in any
/// column of a multi-cell row is treated exactly like column zero.
pub(super) fn read_cell_control(
    source: &Package,
    target: CellTarget,
) -> Result<Option<CellControl>, Error> {
    let mut authority = None;
    let scalar = read_non_popup_control(source, target, &mut authority)?;
    if let Some(control) = scalar {
        return Ok(Some(control));
    }
    Ok(read_popup_with_policy(source, target, false, &mut authority)?.map(CellControl::PopUpMenu))
}

fn read_non_popup_control<'source>(
    source: &'source Package,
    target: CellTarget,
    authority: &mut Option<popup_metadata::RegistryFacts<'source>>,
) -> Result<Option<CellControl>, Error> {
    let path = Path::Cell {
        sheet: target.sheet_position,
        table: target.table_position,
        position: target.position,
    };
    let model = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .ok_or(Error::InvalidSource { path })?;
    if model.type_ != target.message_type || model.type_ != 6_001 {
        return Err(Error::InvalidSource { path });
    }
    let options = storage_options(&model.data);
    let (model_snapshot, _) = storage_codec::decode_table_model_with_report(&model.data, options)
        .map_err(|_| Error::InvalidSource { path })?;
    let (store, _) = storage_codec::decode_data_store_with_report(
        model_snapshot.base_data_store(),
        storage_options(model_snapshot.base_data_store()),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let mut tiles = TileCollector::default();
    let (tile_storage, _) = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        storage_options(store.tiles()),
        &mut tiles,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let tile_size = tile_storage
        .tile_size()
        .ok_or(Error::InvalidSource { path })?;
    let tile_id = target_row_tile(tile_size, target.position.row());
    let tile_ref = tiles
        .tiles
        .iter()
        .find(|tile| tile.0 == tile_id)
        .map(|tile| tile.1)
        .ok_or(Error::CellNotFound)?;
    let tile_component_index = resolved_component_index(source, tile_ref, path)?;
    validate_selected_model_reference(source, target, tile_ref, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        tile_component_index,
        Some(tile_ref),
        path,
    )?;
    let tile_message = resolve_typed_message_any(source, tile_ref, 6_002, path)?;
    let mut rows = RowCollector::default();
    storage_codec::decode_tile_with_visitor(
        &tile_message.data,
        storage_options(&tile_message.data),
        &mut rows,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let row = rows
        .rows
        .iter()
        .find(|row| row.index == target.position.row())
        .ok_or(Error::CellNotFound)?;
    let cell_bytes = row
        .cell(target.position.column())
        .ok_or(Error::InvalidSource { path })?;
    let cell = BncCell::parse(cell_bytes).map_err(|_| Error::InvalidSource { path })?;
    let format_identifier = cell.format_identifier();
    let control_identifier = cell.control_cell_spec_identifier();
    if format_identifier.is_none() && control_identifier.is_none() {
        return Ok(None);
    }
    if format_identifier.is_none() {
        return Err(Error::InvalidSource { path });
    }
    if cell.cell_format_kind() == Some(5) {
        return Ok(None);
    }
    if !matches!(cell.cell_format_kind(), Some(1 | 2 | 6)) {
        return Err(Error::InvalidSource { path });
    }
    let format_identifier = format_identifier.ok_or(Error::InvalidSource { path })?;
    let format_table_identifier = store
        .format_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    let format_resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, format_table_identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    validate_selected_model_reference(source, target, format_table_identifier, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        format_resolved.component_index,
        Some(format_table_identifier),
        path,
    )?;
    let mut format_payload = None;
    for message in format_resolved
        .messages
        .iter()
        .filter(|message| message.type_ == 6_005)
    {
        let mut entries = ListCollector::default();
        let (list, _) = storage_codec::decode_table_data_list_with_visitor(
            &message.data,
            storage_options(&message.data),
            &mut entries,
        )
        .map_err(|_| Error::InvalidSource { path })?;
        if list.list_type() != LIST_FORMAT {
            continue;
        }
        if entries.segments != 0
            || duplicate_list_keys(&entries.entries)
            || format_payload.is_some()
        {
            return Err(Error::InvalidSource { path });
        }
        let entry = entries
            .entries
            .iter()
            .find(|entry| entry.key == format_identifier)
            .ok_or(Error::InvalidSource { path })?;
        if entry.ref_count == 0 {
            return Err(Error::InvalidSource { path });
        }
        format_payload = entry.format.clone();
    }
    let format_payload = format_payload.ok_or(Error::InvalidSource { path })?;
    let (format, _) = control_codec::decode_control_format_with_report(
        &format_payload,
        control_codec::DecodeOptions::for_source(&format_payload),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let Some(control_identifier) = control_identifier else {
        if matches!(format.format_type(), 263 | 267) {
            return Err(Error::InvalidSource { path });
        }
        return Ok(None);
    };

    let control_table_identifier = store
        .control_cell_spec_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    let control_resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, control_table_identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    validate_selected_model_reference(source, target, control_table_identifier, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        control_resolved.component_index,
        Some(control_table_identifier),
        path,
    )?;
    let mut selected = None;
    for (message_index, message) in control_resolved
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == 6_005)
    {
        let mut entries = ListCollector::default();
        let (list, _) = storage_codec::decode_table_data_list_with_visitor(
            &message.data,
            storage_options(&message.data),
            &mut entries,
        )
        .map_err(|_| Error::InvalidSource { path })?;
        if list.list_type() != LIST_CONTROL_CELL_SPEC {
            continue;
        }
        if entries.segments != 0 || duplicate_list_keys(&entries.entries) || selected.is_some() {
            return Err(Error::InvalidSource { path });
        }
        let object = source
            .state
            .components
            .catalog()
            .get_index(control_resolved.component_index)
            .and_then(|component| {
                component
                    .archive()
                    .objects
                    .get(control_resolved.object_index)
            })
            .ok_or(Error::InvalidSource { path })?;
        validate_control_list_metadata(object, message_index, &entries.entries, false, path)?;
        for entry in &entries.entries {
            let spec = entry
                .cell_spec
                .as_deref()
                .ok_or(Error::InvalidSource { path })?;
            let (spec, _) = control_codec::decode_any_cell_spec_with_report(
                spec,
                control_codec::DecodeOptions::for_source(spec),
            )
            .map_err(|_| Error::InvalidSource { path })?;
            let control_codec::CellSpecSnapshot::Popup(spec) = spec else {
                continue;
            };
            let popup_identifier = spec.popup_model().identifier();
            let popup_component_index = resolved_component_index(source, popup_identifier, path)?;
            prove_cross_component_reference(
                source,
                authority,
                control_resolved.component_index,
                popup_component_index,
                Some(popup_identifier),
                path,
            )?;
            let popup = resolve_typed_message_any(source, popup_identifier, 6_206, path)?;
            popup_codec::decode_popup_menu_model_with_report(
                &popup.data,
                popup_codec::DecodeOptions::for_source(&popup.data),
            )
            .map_err(|_| Error::InvalidSource { path })?;
        }
        let entry = entries
            .entries
            .iter()
            .find(|entry| entry.key == control_identifier)
            .ok_or(Error::InvalidSource { path })?;
        if entry.ref_count == 0 {
            return Err(Error::InvalidSource { path });
        }
        let spec = entry
            .cell_spec
            .as_deref()
            .ok_or(Error::InvalidSource { path })?;
        selected = Some(spec.to_owned());
    }
    let spec = selected.ok_or(Error::InvalidSource { path })?;
    let (spec, _) = control_codec::decode_control_cell_spec_with_report(
        &spec,
        control_codec::DecodeOptions::for_source(&spec),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    control_semantic_value(source, target.component_index, spec, format, path)
}

fn control_semantic_value(
    source: &Package,
    component_index: usize,
    spec: control_codec::ControlCellSpecSnapshot<'_>,
    format: control_codec::ControlFormatSnapshot<'_>,
    path: Path,
) -> Result<Option<CellControl>, Error> {
    let interaction = spec.interaction_type();
    match interaction {
        control_codec::CHECKBOX_INTERACTION_TYPE => {
            if format.format_type() != 263 {
                return Err(Error::InvalidSource { path });
            }
            Ok(Some(CellControl::Checkbox(Checkbox)))
        },
        control_codec::STAR_RATING_INTERACTION_TYPE => {
            if format.format_type() != 267
                || spec.range_control_min() != Some(0.0)
                || spec.range_control_max() != Some(5.0)
                || spec.range_control_inc() != Some(1.0)
            {
                return Err(Error::InvalidSource { path });
            }
            Ok(Some(CellControl::StarRating(StarRating)))
        },
        control_codec::SLIDER_INTERACTION_TYPE | control_codec::STEPPER_INTERACTION_TYPE => {
            let range = Range::new(
                spec.range_control_min()
                    .ok_or(Error::InvalidSource { path })?,
                spec.range_control_max()
                    .ok_or(Error::InvalidSource { path })?,
                spec.range_control_inc()
                    .ok_or(Error::InvalidSource { path })?,
            )
            .map_err(|_| Error::InvalidSource { path })?;
            let display = control_display_format(format, path)?;
            if interaction == control_codec::SLIDER_INTERACTION_TYPE {
                Ok(Some(CellControl::Slider(Slider::new(range, display))))
            } else {
                Ok(Some(CellControl::Stepper(Stepper::new(range, display))))
            }
        },
        7 => {
            // Use the strict popup route for the one graph-bearing variant;
            // resolving it from the current cell avoids accepting a stale
            // control spec that merely resembles a popup.
            let _ = (source, component_index);
            Err(Error::InvalidSource { path })
        },
        _ => Err(Error::InvalidSource { path }),
    }
}

fn control_display_format(
    format: control_codec::ControlFormatSnapshot<'_>,
    path: Path,
) -> Result<DisplayFormat, Error> {
    let format_type = format.format_type();
    match format_type {
        256 | 258 => {
            ensure_decimal_format_fields(format, false, path)?;
            let decimal_places = decimal_places_from_native(format.decimal_places(), path)?;
            let negative_style = negative_style_from_native(format.negative_style(), path)?;
            let thousands_separator =
                thousands_separator_from_native(format.show_thousands_separator(), path)?;
            if format_type == 256 {
                Ok(DisplayFormat::Number(Number::new(
                    decimal_places,
                    negative_style,
                    thousands_separator,
                )))
            } else {
                Ok(DisplayFormat::Percentage(Percentage::new(
                    decimal_places,
                    negative_style,
                    thousands_separator,
                )))
            }
        },
        257 => {
            ensure_decimal_format_fields(format, true, path)?;
            let decimal_places = decimal_places_from_native(format.decimal_places(), path)?;
            let negative_style = negative_style_from_native(format.negative_style(), path)?;
            let thousands_separator =
                thousands_separator_from_native(format.show_thousands_separator(), path)?;
            let code = match format.currency_code() {
                Some(value) => {
                    CurrencyCode::new(value).map_err(|_| Error::InvalidSource { path })?
                },
                None => CurrencyCode::USD,
            };
            let style = match format.use_accounting_style() {
                Some(false) => CurrencyStyle::Standard,
                Some(true) => CurrencyStyle::Accounting,
                None => CurrencyStyle::Standard,
            };
            Ok(DisplayFormat::Currency(Currency::new(
                code,
                decimal_places,
                negative_style,
                thousands_separator,
                style,
            )))
        },
        259 => {
            ensure_decimal_format_fields(format, false, path)?;
            if format.negative_style().is_some_and(|value| value != 0)
                || format.show_thousands_separator().is_some_and(|value| value)
            {
                return Err(Error::InvalidSource { path });
            }
            let decimal_places = match format.decimal_places() {
                None => FixedDecimalPlaces::TWO,
                Some(value) => {
                    let DecimalPlaces::Fixed(value) =
                        decimal_places_from_native(Some(value), path)?
                    else {
                        return Err(Error::InvalidSource { path });
                    };
                    value
                },
            };
            Ok(DisplayFormat::Scientific(Scientific::new(decimal_places)))
        },
        262 => {
            ensure_fraction_format_fields(format, path)?;
            let accuracy = match format.fraction_accuracy().unwrap_or(u32::MAX - 2) as i32 {
                -1 => FractionAccuracy::UpToOneDigit,
                -2 => FractionAccuracy::UpToTwoDigits,
                -3 => FractionAccuracy::UpToThreeDigits,
                2 => FractionAccuracy::Halves,
                4 => FractionAccuracy::Quarters,
                8 => FractionAccuracy::Eighths,
                16 => FractionAccuracy::Sixteenths,
                10 => FractionAccuracy::Tenths,
                100 => FractionAccuracy::Hundredths,
                _ => return Err(Error::InvalidSource { path }),
            };
            Ok(DisplayFormat::Fraction(Fraction::new(accuracy)))
        },
        269 => {
            ensure_numeral_format_fields(format, path)?;
            let base = Base::new(
                u8::try_from(format.base().unwrap_or(10))
                    .map_err(|_| Error::InvalidSource { path })?,
            )
            .map_err(|_| Error::InvalidSource { path })?;
            let places = match format.base_places().unwrap_or(0) {
                0 => Places::Minimum,
                value => Places::Fixed(
                    FixedPlaces::new(
                        u8::try_from(value).map_err(|_| Error::InvalidSource { path })?,
                    )
                    .map_err(|_| Error::InvalidSource { path })?,
                ),
            };
            let negative_style = match format.base_use_minus_sign().unwrap_or(true) {
                true => NumeralNegativeStyle::MinusSign,
                false => NumeralNegativeStyle::TwosComplement,
            };
            let value = NumeralSystem::new(base, places, negative_style)
                .map_err(|_| Error::InvalidSource { path })?;
            Ok(DisplayFormat::NumeralSystem(value))
        },
        _ => Err(Error::InvalidSource { path }),
    }
}

fn decimal_places_from_native(value: Option<u32>, path: Path) -> Result<DecimalPlaces, Error> {
    match value.unwrap_or(253) {
        253 => Ok(DecimalPlaces::Automatic),
        value => Ok(DecimalPlaces::Fixed(
            FixedDecimalPlaces::new(
                u8::try_from(value).map_err(|_| Error::InvalidSource { path })?,
            )
            .map_err(|_| Error::InvalidSource { path })?,
        )),
    }
}

fn negative_style_from_native(
    value: Option<u32>,
    path: Path,
) -> Result<NumberNegativeStyle, Error> {
    match value {
        None | Some(0) => Ok(NumberNegativeStyle::MinusSign),
        Some(1) => Ok(NumberNegativeStyle::Red),
        Some(2) => Ok(NumberNegativeStyle::Parentheses),
        Some(3) => Ok(NumberNegativeStyle::RedParentheses),
        _ => Err(Error::InvalidSource { path }),
    }
}

fn thousands_separator_from_native(
    value: Option<bool>,
    _path: Path,
) -> Result<ThousandsSeparator, Error> {
    match value {
        Some(false) => Ok(ThousandsSeparator::Hidden),
        Some(true) => Ok(ThousandsSeparator::Shown),
        None => Ok(ThousandsSeparator::Hidden),
    }
}

fn ensure_decimal_format_fields(
    format: control_codec::ControlFormatSnapshot<'_>,
    currency: bool,
    path: Path,
) -> Result<(), Error> {
    if (!currency && (format.currency_code().is_some() || format.use_accounting_style().is_some()))
        || (currency
            && (format.currency_code().is_some() != format.use_accounting_style().is_some()))
        || format.duration_style().is_some()
        || format.base().is_some()
        || format.base_places().is_some()
        || format.base_use_minus_sign().is_some()
        || format.fraction_accuracy().is_some()
        || format.suppress_date_format().is_some()
        || format.suppress_time_format().is_some()
        || format.date_time_format().is_some()
        || format.duration_unit_largest().is_some()
        || format.duration_unit_smallest().is_some()
        || format.control_minimum().is_some()
        || format.control_maximum().is_some()
        || format.control_increment().is_some()
        || format.control_format_type().is_some()
        || format.slider_orientation().is_some()
        || format.slider_position().is_some()
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn ensure_fraction_format_fields(
    format: control_codec::ControlFormatSnapshot<'_>,
    path: Path,
) -> Result<(), Error> {
    if format.decimal_places().is_some()
        || format.currency_code().is_some()
        || format.negative_style().is_some()
        || format.show_thousands_separator().is_some()
        || format.use_accounting_style().is_some()
        || format.duration_style().is_some()
        || format.base().is_some()
        || format.base_places().is_some()
        || format.base_use_minus_sign().is_some()
        || format.suppress_date_format().is_some()
        || format.suppress_time_format().is_some()
        || format.date_time_format().is_some()
        || format.duration_unit_largest().is_some()
        || format.duration_unit_smallest().is_some()
        || format.control_minimum().is_some()
        || format.control_maximum().is_some()
        || format.control_increment().is_some()
        || format.control_format_type().is_some()
        || format.slider_orientation().is_some()
        || format.slider_position().is_some()
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn ensure_numeral_format_fields(
    format: control_codec::ControlFormatSnapshot<'_>,
    path: Path,
) -> Result<(), Error> {
    if format.decimal_places().is_some()
        || format.currency_code().is_some()
        || format.negative_style().is_some()
        || format.show_thousands_separator().is_some()
        || format.use_accounting_style().is_some()
        || format.duration_style().is_some()
        || format.fraction_accuracy().is_some()
        || format.suppress_date_format().is_some()
        || format.suppress_time_format().is_some()
        || format.date_time_format().is_some()
        || format.duration_unit_largest().is_some()
        || format.duration_unit_smallest().is_some()
        || format.control_minimum().is_some()
        || format.control_maximum().is_some()
        || format.control_increment().is_some()
        || format.control_format_type().is_some()
        || format.slider_orientation().is_some()
        || format.slider_position().is_some()
    {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

fn read_popup_with_policy<'source>(
    source: &'source Package,
    target: CellTarget,
    allow_missing_control_field_infos: bool,
    authority: &mut Option<popup_metadata::RegistryFacts<'source>>,
) -> Result<Option<PopUpMenu>, Error> {
    let path = Path::Cell {
        sheet: target.sheet_position,
        table: target.table_position,
        position: target.position,
    };
    let model = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .ok_or(Error::InvalidSource { path })?;
    if model.type_ != target.message_type || model.type_ != 6_001 {
        return Err(Error::InvalidSource { path });
    }
    let options = storage_options(model.data.as_slice());
    let (model_snapshot, _) = storage_codec::decode_table_model_with_report(&model.data, options)
        .map_err(|_| Error::InvalidSource { path })?;
    let (store, _) = storage_codec::decode_data_store_with_report(
        model_snapshot.base_data_store(),
        storage_options(model_snapshot.base_data_store()),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let mut tiles = TileCollector::default();
    let (tile_storage, _) = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        storage_options(store.tiles()),
        &mut tiles,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let tile_size = tile_storage
        .tile_size()
        .ok_or(Error::InvalidSource { path })?;
    let tile_id = target_row_tile(tile_size, target.position.row());
    let tile_ref = tiles
        .tiles
        .iter()
        .find(|tile| tile.0 == tile_id)
        .map(|tile| tile.1)
        .ok_or(Error::CellNotFound)?;
    let tile_component_index = resolved_component_index(source, tile_ref, path)?;
    validate_selected_model_reference(source, target, tile_ref, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        tile_component_index,
        Some(tile_ref),
        path,
    )?;
    let tile_message = resolve_typed_message_any(source, tile_ref, 6_002, path)?;
    let mut rows = RowCollector::default();
    let (_, _) = storage_codec::decode_tile_with_visitor(
        &tile_message.data,
        storage_options(&tile_message.data),
        &mut rows,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let row = rows
        .rows
        .iter()
        .find(|row| row.index == target.position.row())
        .ok_or(Error::CellNotFound)?;
    let cell_buffer = row
        .cell(target.position.column())
        .ok_or(Error::InvalidSource { path })?;
    let cell = BncCell::parse(cell_buffer).map_err(|_| Error::InvalidSource { path })?;
    let format_id = cell.format_identifier();
    let control_id = cell.control_cell_spec_identifier();
    if format_id.is_none() && control_id.is_none() {
        return Ok(None);
    }
    if cell.cell_format_kind() != Some(5) || control_id.is_none() {
        return Ok(None);
    }
    if format_id.is_none() {
        return Err(Error::InvalidSource { path });
    }
    let format_id = format_id.ok_or(Error::InvalidSource { path })?;
    let control_key = control_id.ok_or(Error::InvalidSource { path })?;
    let format_table_identifier = store
        .format_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    let format_resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, format_table_identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    validate_selected_model_reference(source, target, format_table_identifier, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        format_resolved.component_index,
        Some(format_table_identifier),
        path,
    )?;
    let mut format_entries = None;
    for message in format_resolved
        .messages
        .iter()
        .filter(|message| message.type_ == 6_005)
    {
        let mut candidate_entries = ListCollector::default();
        let (list, _) = storage_codec::decode_table_data_list_with_visitor(
            &message.data,
            storage_options(&message.data),
            &mut candidate_entries,
        )
        .map_err(|_| Error::InvalidSource { path })?;
        if list.list_type() != 2 {
            continue;
        }
        if candidate_entries.segments != 0 || duplicate_list_keys(&candidate_entries.entries) {
            return Err(Error::UnsupportedDependency { path });
        }
        if format_entries.replace(candidate_entries).is_some() {
            return Err(Error::InvalidSource { path });
        }
    }
    let format_entries = format_entries.ok_or(Error::InvalidSource { path })?;
    let format_entry = format_entries
        .entries
        .iter()
        .find(|entry| entry.key == format_id)
        .ok_or(Error::InvalidSource { path })?;
    if format_entry.ref_count == 0 || format_entry.format.is_none() {
        return Err(Error::InvalidSource { path });
    }
    let control_table_identifier = store
        .control_cell_spec_table()
        .ok_or(Error::InvalidSource { path })?
        .identifier();
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, control_table_identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    validate_selected_model_reference(source, target, control_table_identifier, path)?;
    prove_cross_component_reference(
        source,
        authority,
        target.component_index,
        resolved.component_index,
        Some(control_table_identifier),
        path,
    )?;
    let mut selected_entries = None;
    for (message_index, message) in resolved.messages.iter().enumerate() {
        if message.type_ != 6_005 {
            continue;
        }
        let mut candidate_entries = ListCollector::default();
        let (list, _) = storage_codec::decode_table_data_list_with_visitor(
            &message.data,
            storage_options(&message.data),
            &mut candidate_entries,
        )
        .map_err(|_| Error::InvalidSource { path })?;
        if list.list_type() == 12 {
            if candidate_entries.segments != 0 {
                return Err(Error::UnsupportedDependency { path });
            }
            let object = source
                .state
                .components
                .catalog()
                .get_index(resolved.component_index)
                .and_then(|component| component.archive().objects.get(resolved.object_index))
                .ok_or(Error::InvalidSource { path })?;
            validate_control_list_metadata(
                object,
                message_index,
                &candidate_entries.entries,
                allow_missing_control_field_infos,
                path,
            )?;
            if selected_entries.replace(candidate_entries).is_some() {
                return Err(Error::InvalidSource { path });
            }
        }
    }
    let entries = selected_entries.ok_or(Error::InvalidSource { path })?;
    for (index, entry) in entries.entries.iter().enumerate() {
        if entries.entries[..index]
            .iter()
            .any(|candidate| candidate.key == entry.key)
        {
            return Err(Error::InvalidSource { path });
        }
    }
    let entry = entries
        .entries
        .iter()
        .find(|entry| entry.key == control_key)
        .ok_or(Error::InvalidSource { path })?;
    if entry.ref_count == 0 {
        return Err(Error::InvalidSource { path });
    }
    let spec = entry
        .cell_spec
        .as_deref()
        .ok_or(Error::InvalidSource { path })?;
    let (spec, _) = popup_codec::decode_cell_spec_with_report(spec, popup_options(spec))
        .map_err(|_| Error::InvalidSource { path })?;
    if spec.interaction_type() != 7
        || spec.popup_model().deprecated_type().is_some()
        || spec.popup_model().deprecated_is_external().unwrap_or(false)
    {
        return Err(Error::InvalidSource { path });
    }
    let popup_identifier = spec.popup_model().identifier();
    let popup_component_index = resolved_component_index(source, popup_identifier, path)?;
    prove_cross_component_reference(
        source,
        authority,
        resolved.component_index,
        popup_component_index,
        Some(popup_identifier),
        path,
    )?;
    let popup_message = resolve_typed_message_any(source, popup_identifier, 6_206, path)?;
    let (popup, _) = popup_codec::decode_popup_menu_model_with_report(
        &popup_message.data,
        popup_options(&popup_message.data),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    if !popup.has_nil_sentinel() || popup.item_count() == 0 {
        return Err(Error::InvalidSource { path });
    }
    let items = popup.items().map(|item| item.value()).collect::<Vec<_>>();
    let menu = PopUpMenu::new(items).map_err(|_| Error::InvalidSource { path })?;
    Ok(Some(menu.with_initial_selection(
        if spec.starts_with_first() {
            crate::cell::data_format::pop_up_menu::InitialSelection::FirstItem
        } else {
            crate::cell::data_format::pop_up_menu::InitialSelection::Blank
        },
    )))
}

fn validate_control_list_metadata(
    object: &ArchiveObject,
    message_index: usize,
    entries: &[ListEntry],
    allow_missing_field_infos: bool,
    path: Path,
) -> Result<(), Error> {
    if duplicate_list_keys(entries) {
        return Err(Error::InvalidSource { path });
    }
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(Error::InvalidSource { path })?;
    // An empty rooted control list is valid only when its control-specific
    // FieldInfo paths are empty too.  Leaving a stale [3,key] field behind is
    // precisely the metadata ghost that can keep a culled CellSpec alive.
    if entries.is_empty() {
        if info
            .field_infos
            .iter()
            .any(|field| field.path.as_slice().first() == Some(&3))
        {
            return Err(Error::InvalidSource { path });
        }
        return Ok(());
    }
    // A control list can legitimately mix the four scalar controls with
    // Pop-Up Menu entries.  Only the popup interaction owns a 6206 model;
    // scalar CellSpecs therefore contribute neither aggregate references nor
    // [3,key] FieldInfo edges.  Decode every entry through the neutral
    // control projection so scalar records are still validated strictly
    // rather than being treated as opaque bytes.
    let mut expected = Vec::with_capacity(entries.len());
    for entry in entries {
        if entry.ref_count == 0 {
            return Err(Error::InvalidSource { path });
        }
        let payload = entry
            .cell_spec
            .as_deref()
            .ok_or(Error::InvalidSource { path })?;
        let (spec, _) = control_codec::decode_cell_spec_with_report(
            payload,
            control_codec::DecodeOptions::for_source(payload),
        )
        .map_err(|_| Error::InvalidSource { path })?;
        match spec {
            control_codec::CellSpecSnapshot::Control(spec) => {
                if !matches!(
                    spec.interaction_type(),
                    control_codec::CHECKBOX_INTERACTION_TYPE
                        | control_codec::STAR_RATING_INTERACTION_TYPE
                        | control_codec::SLIDER_INTERACTION_TYPE
                        | control_codec::STEPPER_INTERACTION_TYPE
                ) {
                    return Err(Error::InvalidSource { path });
                }
            },
            control_codec::CellSpecSnapshot::Popup(spec) => {
                if spec.interaction_type() != 7
                    || spec.popup_model().identifier() == 0
                    || spec.popup_model().deprecated_type().is_some()
                    || spec.popup_model().deprecated_is_external().unwrap_or(false)
                {
                    return Err(Error::InvalidSource { path });
                }
                expected.push((entry.key, spec.popup_model().identifier()));
            },
        }
    }
    let mut expected_aggregate = expected.iter().map(|(_, id)| *id).collect::<Vec<_>>();
    expected_aggregate.sort_unstable();
    expected_aggregate.dedup();
    let mut aggregate = info.object_references.clone();
    aggregate.sort_unstable();
    aggregate.dedup();
    if info.object_references.contains(&0)
        || info.object_references.len() != aggregate.len()
        || aggregate != expected_aggregate
    {
        return Err(Error::InvalidSource { path });
    }
    if allow_missing_field_infos && info.field_infos.is_empty() {
        return Ok(());
    }
    for (key, identifier) in &expected {
        let matches = info
            .field_infos
            .iter()
            .filter(|field| field.path.as_slice() == [3, *key]);
        let matching = matches.collect::<Vec<_>>();
        if matching.len() != 1
            || matching[0]
                .r#type
                .is_some_and(|kind| kind != litchi_iwa_core::FieldType::ObjectReference)
            || matching[0].object_references.as_slice() != [*identifier]
        {
            return Err(Error::InvalidSource { path });
        }
    }
    // Every control-list FieldInfo must be an exact [3,key] singleton above;
    // an aggregate-only [3] edge or an unrelated key is not a safe ownership
    // proof and is rejected rather than broadcast during a rewrite.
    if info.field_infos.iter().any(|field| {
        field.path.as_slice().first() == Some(&3)
            && !expected
                .iter()
                .any(|(key, _)| field.path.as_slice() == [3, *key])
    }) {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

/// Read one selected cell while charging the same transaction ledger used by
/// the changed route.  The strict read helper remains reusable by the public
/// no-op/read API; changed commits call this wrapper before native candidates
/// and again for the private candidate semantic readback.
fn read_popup_with_budget(
    source: &Package,
    target: CellTarget,
    budget: &mut TransactionBudget,
    allow_missing_control_field_infos: bool,
) -> Result<Option<PopUpMenu>, Error> {
    let path = Path::Cell {
        sheet: target.sheet_position,
        table: target.table_position,
        position: target.position,
    };
    let payload_len = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .map(|message| message.data.len())
        .ok_or(Error::InvalidSource { path })?;
    budget.charge_wire_bytes(payload_len, path)?;
    budget.charge_wire_work(payload_len.saturating_mul(16), path)?;
    budget.charge_payload_items(1, path)?;
    budget.charge_transaction_work(payload_len.saturating_mul(16), path)?;
    let mut authority = None;
    read_popup_with_policy(
        source,
        target,
        allow_missing_control_field_infos,
        &mut authority,
    )
}

fn changed_member_count(
    source: &litchi_iwa_archive::SourceCatalog,
    target: &litchi_iwa_archive::SourceCatalog,
) -> usize {
    source
        .package()
        .iter()
        .filter(|entry| !entry.name().starts_with("preview"))
        .filter(|entry| {
            target
                .package()
                .iter()
                .find(|candidate| candidate.name() == entry.name())
                .is_some_and(|candidate| candidate.data() != entry.data())
        })
        .count()
}

fn storage_options(source: &[u8]) -> storage_codec::DecodeOptions {
    let bytes = source.len().max(1);
    storage_codec::DecodeOptions::new(
        bytes,
        bytes.saturating_mul(8).max(1),
        bytes.saturating_mul(64).max(1),
        64,
        bytes.saturating_mul(2).max(1),
        bytes.saturating_mul(2).max(1),
    )
}

fn popup_options(source: &[u8]) -> popup_codec::DecodeOptions {
    let bytes = source.len().max(1);
    popup_codec::DecodeOptions::for_source(source)
        .with_max_output_bytes(bytes.saturating_mul(2).max(1))
        .with_max_items(bytes.max(1))
        .with_max_text_bytes(bytes.saturating_mul(2).max(1))
}

fn target_row_tile(tile_size: u32, row: u32) -> u32 {
    row / tile_size.max(1)
}

fn resolved_component_index(source: &Package, identifier: u64, path: Path) -> Result<usize, Error> {
    source
        .state
        .index
        .resolve_ref_id(&source.state.components, identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .map(|resolved| resolved.component_index)
        .ok_or(Error::InvalidSource { path })
}

/// Require the selected TableModel message header to declare each native
/// sidecar edge exactly once. Producer-omitted FieldInfo is accepted; when a
/// FieldInfo occurrence is present, it must be unique and object-reference
/// typed. This bounded slice does not claim exact field-path authority for
/// producer-omitted metadata.
fn validate_selected_model_reference(
    source: &Package,
    target: CellTarget,
    identifier: u64,
    path: Path,
) -> Result<(), Error> {
    let info = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.archive_info.message_infos.get(target.message_index))
        .ok_or(Error::InvalidSource { path })?;
    if info
        .object_references
        .iter()
        .filter(|candidate| **candidate == identifier)
        .count()
        != 1
    {
        return Err(Error::InvalidSource { path });
    }
    let mut field_occurrences = 0usize;
    for field in &info.field_infos {
        let occurrences = field
            .object_references
            .iter()
            .filter(|candidate| **candidate == identifier)
            .count();
        if occurrences == 0 {
            continue;
        }
        if field
            .r#type
            .is_some_and(|kind| kind != litchi_iwa_core::FieldType::ObjectReference)
        {
            return Err(Error::InvalidSource { path });
        }
        field_occurrences = field_occurrences
            .checked_add(occurrences)
            .ok_or(Error::InvalidSource { path })?;
    }
    if field_occurrences > 1 {
        return Err(Error::InvalidSource { path });
    }
    Ok(())
}

/// Resolve a generated-free message through the package index without
/// imposing the model component as an implicit owner.  The caller must prove
/// every cross-component edge with `prove_cross_component_reference`; this
/// split is what prevents a sidecar member from being mistaken for an
/// unowned/ambiguous object while still allowing the native graph to span
/// CalculationEngine, Tile, and DataList members.
fn resolve_typed_message_any(
    source: &Package,
    identifier: u64,
    message_type: u32,
    path: Path,
) -> Result<&RawMessage, Error> {
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, identifier)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    let mut matches = resolved
        .messages
        .iter()
        .filter(|message| message.type_ == message_type);
    let message = matches.next().ok_or(Error::InvalidSource { path })?;
    if matches.next().is_some() {
        return Err(Error::InvalidSource { path });
    }
    Ok(message)
}

/// Pop-Up Menu changed routes remain intentionally single-member until that
/// graph's model/string/style lifecycle can clone and publish every touched
/// sidecar together with one metadata transition. Scalar controls use their
/// separate strict multi-member writer; keep this guard for popup transitions
/// so an unsupported popup graph fails before native candidate allocation.
pub(super) fn reject_cross_component_write(
    source: &Package,
    target: CellTarget,
    path: Path,
) -> Result<(), Error> {
    let model = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .and_then(|component| component.archive().objects.get(target.object_index))
        .and_then(|object| object.messages.get(target.message_index))
        .ok_or(Error::InvalidSource { path })?;
    let (model_snapshot, _) =
        storage_codec::decode_table_model_with_report(&model.data, storage_options(&model.data))
            .map_err(|_| Error::InvalidSource { path })?;
    let (store, _) = storage_codec::decode_data_store_with_report(
        model_snapshot.base_data_store(),
        storage_options(model_snapshot.base_data_store()),
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let mut tiles = TileCollector::default();
    let (tile_storage, _) = storage_codec::decode_tile_storage_with_visitor(
        store.tiles(),
        storage_options(store.tiles()),
        &mut tiles,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let tile_id = target.position.row() / tile_storage.tile_size().unwrap_or(1).max(1);
    let tile_identifier = tiles
        .tiles
        .iter()
        .find(|(id, _)| *id == tile_id)
        .map(|(_, identifier)| *identifier)
        .ok_or(Error::CellNotFound)?;
    let mut identifiers = vec![tile_identifier];
    if let Some(reference) = store.format_table() {
        identifiers.push(reference.identifier());
    }
    let control_table_identifier = store
        .control_cell_spec_table()
        .map(|reference| reference.identifier());
    if let Some(identifier) = control_table_identifier {
        identifiers.push(identifier);
    }
    for identifier in identifiers {
        let resolved = source
            .state
            .index
            .resolve_ref_id(&source.state.components, identifier)
            .map_err(|_| Error::InvalidSource { path })?
            .ok_or(Error::InvalidSource { path })?;
        if resolved.component_index != target.component_index {
            return Err(Error::UnsupportedDependency { path });
        }
    }
    if let Some(identifier) = control_table_identifier {
        let resolved = source
            .state
            .index
            .resolve_ref_id(&source.state.components, identifier)
            .map_err(|_| Error::InvalidSource { path })?
            .ok_or(Error::InvalidSource { path })?;
        let mut control_list_seen = false;
        for message in resolved
            .messages
            .iter()
            .filter(|message| message.type_ == 6_005)
        {
            let mut entries = ListCollector::default();
            let (list, _) = storage_codec::decode_table_data_list_with_visitor(
                &message.data,
                storage_options(&message.data),
                &mut entries,
            )
            .map_err(|_| Error::InvalidSource { path })?;
            if list.list_type() != LIST_CONTROL_CELL_SPEC {
                continue;
            }
            if control_list_seen || entries.segments != 0 || duplicate_list_keys(&entries.entries) {
                return Err(Error::InvalidSource { path });
            }
            control_list_seen = true;
            for entry in &entries.entries {
                let spec = entry
                    .cell_spec
                    .as_deref()
                    .ok_or(Error::InvalidSource { path })?;
                let (spec, _) = control_codec::decode_any_cell_spec_with_report(
                    spec,
                    control_codec::DecodeOptions::for_source(spec),
                )
                .map_err(|_| Error::InvalidSource { path })?;
                let control_codec::CellSpecSnapshot::Popup(spec) = spec else {
                    continue;
                };
                let popup = source
                    .state
                    .index
                    .resolve_ref_id(&source.state.components, spec.popup_model().identifier())
                    .map_err(|_| Error::InvalidSource { path })?
                    .ok_or(Error::InvalidSource { path })?;
                if popup.component_index != target.component_index {
                    return Err(Error::UnsupportedDependency { path });
                }
            }
        }
        if !control_list_seen {
            return Err(Error::InvalidSource { path });
        }
    }
    Ok(())
}

/// Validate the metadata ownership edge before a cross-component read is
/// exposed.  Sidecar components normally have no UUID object records, so an
/// exact current external component edge is the required authority.  Missing
/// Metadata, a missing/ambiguous effective locator, versioned edges, duplicate
/// edges, and weak/reference-shape conflicts all fail closed.
pub(super) fn prove_cross_component_reference<'source>(
    source: &'source Package,
    authority: &mut Option<popup_metadata::RegistryFacts<'source>>,
    source_component_index: usize,
    target_component_index: usize,
    object_identifier: Option<u64>,
    path: Path,
) -> Result<(), Error> {
    if source_component_index == target_component_index {
        return Ok(());
    }
    if authority.is_none() {
        let bytes = source.source_bytes().len().max(1);
        let options = MetadataRewriteOptions::new(
            bytes,
            bytes.saturating_mul(2),
            bytes.saturating_mul(16),
            bytes.saturating_mul(64),
            64,
            bytes.saturating_mul(2),
            bytes.saturating_mul(2),
            1,
        );
        *authority = Some(
            popup_metadata::inspect_cross_component_read(source, options)
                .map_err(|_| Error::UnsupportedDependency { path })?,
        );
    }
    let facts = authority
        .as_ref()
        .ok_or(Error::UnsupportedDependency { path })?;
    if facts.has_physical_alias() {
        return Err(Error::UnsupportedDependency { path });
    }
    facts
        .require_external_edge(
            source_component_index,
            target_component_index,
            object_identifier,
            Some(false),
        )
        .map_err(|_| Error::UnsupportedDependency { path })
}

#[derive(Default)]
struct TileCollector {
    tiles: Vec<(u32, u64)>,
}

impl storage_codec::StorageVisitor for TileCollector {
    fn visit_tile_reference(
        &mut self,
        record: storage_codec::TileReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        self.tiles
            .push((record.tile_id(), record.reference().identifier()));
        Ok(())
    }
}

#[derive(Default)]
struct RowCollector {
    rows: Vec<RowFact>,
}

struct RowFact {
    index: u32,
    cell_count: u32,
    storage: Vec<u8>,
    offsets: Vec<u8>,
    wide_offsets: bool,
}

impl RowFact {
    fn cell(&self, column: u32) -> Option<&[u8]> {
        let column = usize::try_from(column).ok()?;
        let count = usize::try_from(self.cell_count).ok()?;
        if count == 1 && self.offsets.is_empty() {
            return (column == 0).then_some(self.storage.as_slice());
        }
        if !self.offsets.len().is_multiple_of(2) {
            return None;
        }
        let slot_count = self.offsets.len() / 2;
        if column >= slot_count || slot_count < count {
            return None;
        }
        let unit = if self.wide_offsets { 4usize } else { 1usize };
        let mut starts = Vec::with_capacity(slot_count);
        let mut previous = None;
        for encoded in self.offsets.chunks_exact(2) {
            let raw = u16::from_le_bytes([encoded[0], encoded[1]]);
            if raw == u16::MAX {
                starts.push(None);
                continue;
            }
            let start = usize::from(raw).checked_mul(unit)?;
            if previous.is_some_and(|prior| prior >= start) {
                return None;
            }
            starts.push(Some(start));
            previous = Some(start);
        }
        if starts.iter().flatten().count() != count {
            return None;
        }
        let start = starts.get(column).copied().flatten()?;
        let end = starts
            .iter()
            .skip(column + 1)
            .flatten()
            .next()
            .copied()
            .unwrap_or(self.storage.len());
        (start < end && end <= self.storage.len()).then_some(&self.storage[start..end])
    }
}

impl storage_codec::StorageVisitor for RowCollector {
    fn visit_tile_row(
        &mut self,
        row: storage_codec::TileRowInfoSnapshot<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        let storage = row
            .cell_storage_buffer()
            .unwrap_or_else(|| row.cell_storage_buffer_pre_bnc())
            .to_vec();
        let offsets = row
            .cell_offsets()
            .unwrap_or_else(|| row.cell_offsets_pre_bnc())
            .to_vec();
        self.rows.push(RowFact {
            index: row.tile_row_index(),
            cell_count: row.cell_count(),
            storage,
            offsets,
            wide_offsets: row.has_wide_offsets().unwrap_or(false),
        });
        Ok(())
    }
}

#[derive(Default)]
struct ListCollector {
    entries: Vec<ListEntry>,
    segments: usize,
}

#[derive(Clone)]
struct ListEntry {
    key: u32,
    ref_count: u32,
    cell_spec: Option<Vec<u8>>,
    format: Option<Vec<u8>>,
}

fn duplicate_list_keys(entries: &[ListEntry]) -> bool {
    entries.iter().enumerate().any(|(index, entry)| {
        entries[..index]
            .iter()
            .any(|previous| previous.key == entry.key)
    })
}

impl storage_codec::StorageVisitor for ListCollector {
    fn visit_list_entry_record(
        &mut self,
        record: storage_codec::TableDataListEntryRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        let snapshot = record.snapshot();
        self.entries.push(ListEntry {
            key: snapshot.key(),
            ref_count: snapshot.ref_count(),
            cell_spec: snapshot.cell_spec().map(ToOwned::to_owned),
            format: snapshot.format().map(ToOwned::to_owned),
        });
        Ok(())
    }

    fn visit_list_segment(
        &mut self,
        _reference: storage_codec::ReferenceRecord<'_>,
    ) -> Result<(), storage_codec::DecodeError> {
        self.segments = self.segments.saturating_add(1);
        Ok(())
    }
}
