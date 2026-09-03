//! Exact-source ownership of a rooted Numbers table's persisted sort order.
//!
//! This transaction edits only `TableModelArchive.sort_order` (field 44).  It
//! deliberately does not execute a sort: moving rows changes tiles, formulas,
//! comments, and view state and remains a host operation.  The wire helpers in
//! this module operate on borrowed field payloads and copy every field that is
//! not part of the selected sort envelope verbatim.  In particular field 45,
//! the sort-rule reference tracker, is opaque and is never synthesized or
//! removed by this owner.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The focused transaction keeps its public value types beside the exact-source machinery."
)]

use std::{fmt, mem::size_of};

use litchi_iwa_archive::package::{
    EntryEdit, OwnedExactArtifacts, ReassemblyExecutionRequirements,
};
use litchi_iwa_archive::{Error as ArchiveError, LimitKind as ArchiveLimitKind};
use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{
    Archive, Error as CoreError, LimitKind as CoreLimitKind, RawMessage, SnappyStream,
};
use litchi_iwa_protos::table_sort_order_codec as codec;
use thiserror::Error as ThisError;

use super::{
    Package,
    physical_entry_index::{Error as PhysicalEntryIndexError, PhysicalEntryIndex},
    table_headers,
};
use crate::{
    selector::{SheetSelector, TableSelector},
    table::{
        lock::State as LockState,
        sort::{Direction, Order, Rule, Scope},
    },
};

const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;

#[cfg(test)]
mod phase_observer {
    use std::{cell::Cell, thread_local};

    thread_local! {
        static PREFLIGHT: Cell<usize> = const { Cell::new(0) };
    }

    pub(super) fn hit_preflight() {
        PREFLIGHT.with(|counter| counter.set(counter.get().saturating_add(1)));
    }

    pub(super) fn reset() {
        PREFLIGHT.with(|counter| counter.set(0));
    }

    pub(super) fn preflight_count() -> usize {
        PREFLIGHT.with(Cell::get)
    }
}

#[cfg(test)]
macro_rules! preflight {
    () => {
        phase_observer::hit_preflight()
    };
}

#[cfg(not(test))]
macro_rules! preflight {
    () => {};
}

/// A content-free location associated with a sort transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Path {
    /// The complete Numbers package.
    Package,
    /// One rooted table at checked zero-based positions.
    Table { sheet: usize, table: usize },
}

/// A finite resource governed by a persisted-sort transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete source artifact bytes.
    InputBytes,
    /// Complete candidate artifact bytes.
    OutputBytes,
    /// Retained package entries.
    Entries,
    /// Bytes in one package entry.
    EntryBytes,
    /// Aggregate package entry bytes.
    TotalEntryBytes,
    /// Package naming and structural metadata bytes.
    PackageBytes,
    /// One decoded native payload.
    PayloadBytes,
    /// Aggregate decoded payload bytes.
    TotalPayloadBytes,
    /// Native objects inspected.
    PayloadObjects,
    /// Native messages inspected.
    PayloadMessages,
    /// Native metadata items inspected.
    PayloadItems,
    /// Native references inspected.
    PayloadReferences,
    /// Wire bytes inspected.
    WireBytes,
    /// Wire bytes emitted by a rewrite.
    WireOutputBytes,
    /// Wire fields inspected.
    WireFields,
    /// Wire nesting depth.
    WireNesting,
    /// Wire traversal and rewrite work.
    WireWork,
    /// Aggregate transaction work.
    TransactionWork,
    /// Codec rule-count limit.
    WireRules,
    /// Codec column-count limit.
    WireColumns,
    /// Codec allocation-count limit.
    WireAllocations,
    /// Codec retained-byte limit.
    WireRetainedBytes,
    /// Codec scratch-byte limit.
    WireScratchBytes,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalEntryBytes => "total entry bytes",
            Self::PackageBytes => "package bytes",
            Self::PayloadBytes => "payload bytes",
            Self::TotalPayloadBytes => "total payload bytes",
            Self::PayloadObjects => "payload objects",
            Self::PayloadMessages => "payload messages",
            Self::PayloadItems => "payload items",
            Self::PayloadReferences => "payload references",
            Self::WireBytes => "wire bytes",
            Self::WireOutputBytes => "wire output bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::TransactionWork => "transaction work",
            Self::WireRules => "wire rules",
            Self::WireColumns => "wire columns",
            Self::WireAllocations => "wire allocations",
            Self::WireRetainedBytes => "wire retained bytes",
            Self::WireScratchBytes => "wire scratch bytes",
        })
    }
}

/// A content-redacted persisted-sort failure.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// No rooted sheet matched the selector.
    #[error("the Numbers workbook has no sheet matching the requested selector")]
    SheetNotFound,
    /// No table on the selected sheet matched the selector.
    #[error("the selected Numbers sheet has no table matching the requested selector")]
    TableNotFound,
    /// The selected table is locked for a changed transaction.
    #[error("the selected Numbers table is locked at {path:?}")]
    TableLocked { path: Path },
    /// This exact physical source cannot publish a changed model payload.
    #[error("this Numbers source does not support exact persisted-sort editing")]
    UnsupportedSource,
    /// Rooted ownership or wire framing is invalid.
    #[error("the Numbers persisted-sort source is invalid at {path:?}")]
    InvalidSource { path: Path },
    /// A finite transaction resource ceiling was exceeded.
    #[error("Numbers persisted-sort {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
        path: Path,
    },
    /// A bounded allocation failed before publication.
    #[error("could not allocate {amount} Numbers persisted-sort units")]
    Allocation { amount: usize, path: Path },
    /// Candidate reopening or exact locality verification failed.
    #[error("the edited Numbers persisted sort failed semantic verification")]
    Verification,
    /// The patch was produced from a different exact source artifact.
    #[error("the Numbers persisted-sort patch does not match the exact source package")]
    PatchConflict,
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Target {
    native: table_headers::Target,
    before: Option<Order>,
}

/// One immutable persisted-sort edit.
pub struct Edit<'a> {
    source: &'a Package,
    target: Target,
    after: Option<Order>,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field("path", &self.path())
            .field("before", &self.target.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Edit<'_> {
    /// Return the selected table path.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Table {
            sheet: self.target.native.sheet_position,
            table: self.target.native.table_position,
        }
    }

    /// Return the order staged for publication.
    #[must_use]
    pub fn order(&self) -> Option<&Order> {
        self.after.as_ref()
    }

    /// Stage a persisted sort order.  This does not reorder any table rows.
    #[must_use]
    pub fn set(mut self, order: Order) -> Self {
        self.after = Some(order);
        self
    }

    /// Stage removal of the persisted sort configuration.
    #[must_use]
    pub fn clear(mut self) -> Self {
        self.after = None;
        self
    }

    /// Alias for [`Self::clear`].
    #[must_use]
    pub fn reset(self) -> Self {
        self.clear()
    }

    /// Validate and atomically publish this exact-source edit.
    pub fn commit(self) -> Result<Commit, Error> {
        commit_edit(self)
    }
}

/// A reversible patch bound to the exact source and target bytes.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    artifacts: OwnedExactArtifacts,
    target: Target,
    before: Option<Order>,
    after: Option<Order>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("path", &self.path())
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the selected table path.
    #[must_use]
    pub const fn path(&self) -> Path {
        Path::Table {
            sheet: self.target.native.sheet_position,
            table: self.target.native.table_position,
        }
    }

    /// Return the source semantic order.
    #[must_use]
    pub fn before(&self) -> Option<&Order> {
        self.before.as_ref()
    }

    /// Return the target semantic order.
    #[must_use]
    pub fn after(&self) -> Option<&Order> {
        self.after.as_ref()
    }

    /// Return a stable diagnostic source fingerprint without exposing bytes.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return a stable diagnostic target fingerprint without exposing bytes.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Whether the patch retains the exact source artifact.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            target: self.target.clone(),
            before: self.after.clone(),
            after: self.before.clone(),
        }
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

    const fn published() -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews: 0,
            full_reparse_performed: true,
        }
    }

    /// Whether model bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Number of rewritten native components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Sort configuration never invalidates previews.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Whether the candidate was reopened and semantically checked.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Operation-local accounting for one persisted-sort read or transaction.
///
/// The codec and physical layers expose independent limits, so this owner
/// carries one residual ledger across rooted selection, codec traversal,
/// archive/Snappy work, ZIP reassembly, candidate reopen, and locality.  The
/// native rewrite still uses staged private buffers: the ledger is charged
/// before publication, while the codec's source-borrowing plan and the IWA
/// archive remain private until the candidate has been reopened and checked.
#[derive(Debug, Clone, Copy)]
struct TransactionBudget {
    maximum_input_bytes: usize,
    remaining_input_bytes: usize,
    maximum_output_bytes: usize,
    remaining_output_bytes: usize,
    maximum_entries: usize,
    remaining_entries: usize,
    maximum_entry_bytes: usize,
    remaining_entry_bytes: usize,
    maximum_total_entry_bytes: usize,
    remaining_total_entry_bytes: usize,
    maximum_payload_bytes: usize,
    remaining_payload_bytes: usize,
    maximum_total_payload_bytes: usize,
    remaining_total_payload_bytes: usize,
    maximum_wire_bytes: usize,
    remaining_wire_bytes: usize,
    maximum_rules: usize,
    remaining_rules: usize,
    maximum_references: usize,
    remaining_references: usize,
    maximum_payload_objects: usize,
    remaining_payload_objects: usize,
    maximum_payload_messages: usize,
    remaining_payload_messages: usize,
    maximum_payload_items: usize,
    remaining_payload_items: usize,
    maximum_wire_output_bytes: usize,
    remaining_wire_output_bytes: usize,
    maximum_fields: usize,
    remaining_fields: usize,
    maximum_work: usize,
    remaining_work: usize,
    maximum_allocations: usize,
    remaining_allocations: usize,
    maximum_retained_bytes: usize,
    remaining_retained_bytes: usize,
    maximum_scratch_bytes: usize,
    remaining_scratch_bytes: usize,
    maximum_transaction_work: usize,
    remaining_transaction_work: usize,
    maximum_nesting: u32,
    observed_nesting: u32,
}

impl TransactionBudget {
    fn new(source: &Package) -> Self {
        let archive = source.state.options.archive();
        let maximum_input_bytes = usize::try_from(archive.max_input_bytes())
            .unwrap_or(usize::MAX)
            .max(1);
        let maximum_output_bytes = maximum_input_bytes;
        let maximum_entries = archive.max_entries().max(1);
        let maximum_entry_bytes = usize::try_from(archive.max_entry_bytes())
            .unwrap_or(usize::MAX)
            .max(1);
        let maximum_total_entry_bytes = usize::try_from(archive.max_total_bytes())
            .unwrap_or(usize::MAX)
            .max(1);
        let maximum_payload_bytes = archive.max_iwa_stream_bytes().max(1);
        let maximum_total_payload_bytes = maximum_total_entry_bytes;
        let maximum_wire_bytes = maximum_payload_bytes;
        let maximum_rules = maximum_wire_bytes.max(1);
        let maximum_references = maximum_wire_bytes.saturating_mul(4).max(1);
        let component_count = source.state.components.catalog().len().max(1);
        let archive_profile = archive.archive_limits();
        let maximum_payload_objects = archive_profile
            .max_objects()
            .saturating_mul(component_count)
            .max(1);
        let maximum_payload_messages = archive_profile
            .max_messages()
            .saturating_mul(component_count)
            .max(1);
        let maximum_payload_items = archive_profile
            .max_metadata_items()
            .saturating_mul(component_count)
            .max(1);
        let maximum_wire_output_bytes = maximum_wire_bytes.saturating_mul(2).max(1);
        let maximum_fields = maximum_wire_bytes.saturating_mul(8).max(1);
        let maximum_work = maximum_wire_bytes.saturating_mul(32).max(1);
        let maximum_allocations = source
            .state
            .components
            .catalog()
            .len()
            .saturating_mul(64)
            .saturating_add(128)
            .max(128);
        let maximum_retained_bytes = maximum_output_bytes;
        let maximum_scratch_bytes = maximum_payload_bytes
            .saturating_add(maximum_wire_bytes)
            .max(1);
        let maximum_transaction_work = maximum_total_entry_bytes.saturating_mul(32).max(1);
        let maximum_nesting = u32::try_from(WireLimits::MAX_NESTING).unwrap_or(u32::MAX);
        Self {
            maximum_input_bytes,
            remaining_input_bytes: maximum_input_bytes,
            maximum_output_bytes,
            remaining_output_bytes: maximum_output_bytes,
            maximum_entries,
            remaining_entries: maximum_entries,
            maximum_entry_bytes,
            remaining_entry_bytes: maximum_entry_bytes,
            maximum_total_entry_bytes,
            remaining_total_entry_bytes: maximum_total_entry_bytes,
            maximum_payload_bytes,
            remaining_payload_bytes: maximum_payload_bytes,
            maximum_total_payload_bytes,
            remaining_total_payload_bytes: maximum_total_payload_bytes,
            maximum_wire_bytes,
            remaining_wire_bytes: maximum_wire_bytes,
            maximum_rules,
            remaining_rules: maximum_rules,
            maximum_references,
            remaining_references: maximum_references,
            maximum_payload_objects,
            remaining_payload_objects: maximum_payload_objects,
            maximum_payload_messages,
            remaining_payload_messages: maximum_payload_messages,
            maximum_payload_items,
            remaining_payload_items: maximum_payload_items,
            maximum_wire_output_bytes,
            remaining_wire_output_bytes: maximum_wire_output_bytes,
            maximum_fields,
            remaining_fields: maximum_fields,
            maximum_work,
            remaining_work: maximum_work,
            maximum_allocations,
            remaining_allocations: maximum_allocations,
            maximum_retained_bytes,
            remaining_retained_bytes: maximum_retained_bytes,
            maximum_scratch_bytes,
            remaining_scratch_bytes: maximum_scratch_bytes,
            maximum_transaction_work,
            remaining_transaction_work: maximum_transaction_work,
            maximum_nesting,
            observed_nesting: 0,
        }
    }

    fn codec_options(&self, source: &[u8], columns: u32) -> codec::DecodeOptions {
        codec::DecodeOptions::new(
            self.remaining_wire_bytes.min(source.len().max(1)),
            self.remaining_wire_output_bytes.max(1),
            self.remaining_fields.max(1),
            self.remaining_work.max(1),
            self.maximum_nesting,
            self.remaining_rules.max(1),
            usize::try_from(columns).unwrap_or(usize::MAX).max(1),
        )
        .with_max_allocations(self.remaining_allocations.max(1))
    }

    fn charge_input_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_input_bytes,
            self.maximum_input_bytes,
            amount,
            LimitKind::InputBytes,
            path,
        )
    }

    fn charge_output_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_output_bytes,
            self.maximum_output_bytes,
            amount,
            LimitKind::OutputBytes,
            path,
        )
    }

    fn charge_entries(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_entries,
            self.maximum_entries,
            amount,
            LimitKind::Entries,
            path,
        )
    }

    fn charge_entry_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_entry_bytes,
            self.maximum_entry_bytes,
            amount,
            LimitKind::EntryBytes,
            path,
        )
    }

    fn charge_total_entry_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_total_entry_bytes,
            self.maximum_total_entry_bytes,
            amount,
            LimitKind::TotalEntryBytes,
            path,
        )
    }

    fn charge_payload_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_payload_bytes,
            self.maximum_payload_bytes,
            amount,
            LimitKind::PayloadBytes,
            path,
        )
    }

    fn charge_total_payload_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_total_payload_bytes,
            self.maximum_total_payload_bytes,
            amount,
            LimitKind::TotalPayloadBytes,
            path,
        )
    }

    fn charge_wire_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_wire_bytes,
            self.maximum_wire_bytes,
            amount,
            LimitKind::WireBytes,
            path,
        )
    }

    fn charge_references(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_references,
            self.maximum_references,
            amount,
            LimitKind::PayloadReferences,
            path,
        )
    }

    fn charge_rules(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_rules,
            self.maximum_rules,
            amount,
            LimitKind::WireRules,
            path,
        )
    }

    fn charge_payload_objects(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_payload_objects,
            self.maximum_payload_objects,
            amount,
            LimitKind::PayloadObjects,
            path,
        )
    }

    fn charge_payload_messages(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_payload_messages,
            self.maximum_payload_messages,
            amount,
            LimitKind::PayloadMessages,
            path,
        )
    }

    fn charge_payload_items(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_payload_items,
            self.maximum_payload_items,
            amount,
            LimitKind::PayloadItems,
            path,
        )
    }

    fn charge_wire_output_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_wire_output_bytes,
            self.maximum_wire_output_bytes,
            amount,
            LimitKind::WireOutputBytes,
            path,
        )
    }

    fn charge_fields(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_fields,
            self.maximum_fields,
            amount,
            LimitKind::WireFields,
            path,
        )
    }

    fn charge_work(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_work,
            self.maximum_work,
            amount,
            LimitKind::WireWork,
            path,
        )
    }

    fn charge_allocations(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_allocations,
            self.maximum_allocations,
            amount,
            LimitKind::WireAllocations,
            path,
        )
    }

    fn charge_retained_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_retained_bytes,
            self.maximum_retained_bytes,
            amount,
            LimitKind::WireRetainedBytes,
            path,
        )
    }

    fn charge_scratch_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_scratch_bytes,
            self.maximum_scratch_bytes,
            amount,
            LimitKind::WireScratchBytes,
            path,
        )
    }

    fn charge_transaction_work(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        charge_budget(
            &mut self.remaining_transaction_work,
            self.maximum_transaction_work,
            amount,
            LimitKind::TransactionWork,
            path,
        )
    }

    fn charge_nesting(&mut self, observed: u32, path: Path) -> Result<(), Error> {
        self.observed_nesting = self.observed_nesting.max(observed);
        if self.observed_nesting > self.maximum_nesting {
            return Err(Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: u64::from(self.observed_nesting),
                maximum: u64::from(self.maximum_nesting),
                path,
            });
        }
        Ok(())
    }

    fn preflight_source(&mut self, source: &Package, path: Path) -> Result<(), Error> {
        preflight!();
        let bytes = source.source_bytes().len();
        self.charge_input_bytes(bytes, path)?;
        self.charge_transaction_work(bytes.saturating_mul(2), path)?;
        let catalog = physical_source(source)?;
        let entries = catalog.package().iter().count();
        let total_entry_bytes = catalog.package().iter().fold(0usize, |total, entry| {
            total.saturating_add(entry.data().len())
        });
        self.charge_entries(entries, path)?;
        self.charge_total_entry_bytes(total_entry_bytes, path)?;
        self.preflight_inventory(source, path)
    }

    /// Charge the immutable parsed inventory before any focused selection or
    /// candidate staging.  Object/message/item/reference counts are kept on
    /// their own ledgers so a small semantic wire ceiling cannot be bypassed
    /// by spreading the same work over many native payloads.
    fn preflight_inventory(&mut self, source: &Package, path: Path) -> Result<(), Error> {
        let mut objects = 0usize;
        let mut messages = 0usize;
        let mut items = 0usize;
        let mut references = 0usize;
        let mut payload_bytes = 0usize;
        for object in source.state.components.iter_objects() {
            objects = objects.checked_add(1).ok_or_else(invalid_source)?;
            messages = messages
                .checked_add(object.messages.len())
                .ok_or_else(invalid_source)?;
            items = items
                .checked_add(object.archive_info.message_infos.len())
                .ok_or_else(invalid_source)?;
            for message in &object.messages {
                payload_bytes = payload_bytes
                    .checked_add(message.data.len())
                    .ok_or_else(invalid_source)?;
            }
            for message_info in &object.archive_info.message_infos {
                references = references
                    .checked_add(
                        message_info
                            .object_references
                            .len()
                            .saturating_add(message_info.data_references.len()),
                    )
                    .ok_or_else(invalid_source)?;
                items = items
                    .checked_add(message_info.field_infos.len())
                    .ok_or_else(invalid_source)?;
                for field_info in &message_info.field_infos {
                    references = references
                        .checked_add(
                            field_info
                                .object_references
                                .len()
                                .saturating_add(field_info.data_references.len()),
                        )
                        .ok_or_else(invalid_source)?;
                    items = items
                        .checked_add(
                            field_info
                                .object_references
                                .len()
                                .saturating_add(field_info.data_references.len()),
                        )
                        .ok_or_else(invalid_source)?;
                }
            }
        }
        self.charge_payload_objects(objects, path)?;
        self.charge_payload_messages(messages, path)?;
        self.charge_payload_items(items, path)?;
        self.charge_references(references, path)?;
        self.charge_total_payload_bytes(payload_bytes, path)?;
        self.charge_transaction_work(
            payload_bytes
                .saturating_add(objects)
                .saturating_add(messages)
                .saturating_add(items)
                .saturating_add(references),
            path,
        )
    }

    fn consume_codec_report(
        &mut self,
        report: codec::DecodeReport,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_wire_bytes(report.input_bytes(), path)?;
        self.charge_wire_output_bytes(report.output_bytes(), path)?;
        self.charge_fields(report.fields(), path)?;
        self.charge_work(report.work_bytes(), path)?;
        self.charge_rules(report.rules(), path)?;
        self.charge_allocations(report.allocations(), path)?;
        self.charge_retained_bytes(report.retained_bytes(), path)?;
        self.charge_scratch_bytes(report.scratch_bytes(), path)?;
        self.charge_nesting(report.max_depth(), path)
    }

    fn preflight_codec_requirements(
        &mut self,
        requirements: codec::RewriteExecutionRequirements,
        source_bytes: usize,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_wire_bytes(source_bytes, path)?;
        self.charge_wire_output_bytes(requirements.output_bytes, path)?;
        self.charge_fields(requirements.fields, path)?;
        self.charge_work(requirements.work_bytes, path)?;
        self.charge_rules(requirements.rules, path)?;
        self.charge_allocations(requirements.allocations, path)?;
        self.charge_retained_bytes(requirements.retained_bytes, path)?;
        self.charge_scratch_bytes(requirements.scratch_bytes, path)?;
        self.charge_nesting(requirements.max_depth, path)
    }

    fn preflight_physical_entry(
        &mut self,
        entry_bytes: usize,
        archive_limits: litchi_iwa_core::Limits,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_entry_bytes(entry_bytes, path)?;
        let archive_bytes = archive_limits.max_archive_bytes();
        let compressed_bound =
            SnappyStream::maximum_compressed_len(archive_bytes).map_err(map_core_error)?;
        self.charge_transaction_work(
            archive_bytes
                .saturating_add(compressed_bound)
                .saturating_add(entry_bytes),
            path,
        )?;
        self.charge_allocations(4, path)?;
        self.charge_scratch_bytes(archive_bytes, path)?;
        self.charge_retained_bytes(archive_bytes, path)?;
        self.charge_total_payload_bytes(archive_bytes, path)
    }

    fn preflight_reassembly(
        &mut self,
        requirements: ReassemblyExecutionRequirements,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_output_bytes(requirements.output_bytes(), path)?;
        self.charge_retained_bytes(requirements.retained_bytes(), path)?;
        self.charge_scratch_bytes(requirements.scratch_bytes(), path)?;
        self.charge_allocations(requirements.allocations(), path)?;
        self.charge_transaction_work(
            requirements
                .output_bytes()
                .saturating_add(requirements.scratch_bytes())
                .saturating_add(requirements.retained_bytes()),
            path,
        )
    }

    fn preflight_candidate_reopen(
        &mut self,
        candidate_bytes: usize,
        components: usize,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_transaction_work(candidate_bytes.saturating_mul(2), path)?;
        self.charge_allocations(components.saturating_add(2), path)
    }

    fn preflight_preview_scan(&mut self, source: &Package, path: Path) -> Result<(), Error> {
        let catalog = physical_source(source)?;
        let preview_candidates = catalog
            .package()
            .iter()
            .filter(|entry| {
                matches!(
                    entry.name(),
                    "preview.jpg" | "preview-micro.jpg" | "preview-web.jpg"
                )
            })
            .count();
        self.charge_transaction_work(
            catalog
                .package()
                .iter()
                .count()
                .saturating_add(preview_candidates),
            path,
        )?;
        // root_preview_deletions retains only static names, but the Vec and
        // its scan workspace still belong to this operation's private stage.
        self.charge_allocations(1, path)?;
        self.charge_scratch_bytes(
            preview_candidates.max(3).saturating_mul(size_of::<&str>()),
            path,
        )
    }

    fn preflight_locality(
        &mut self,
        source: &Package,
        candidate: &Package,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_transaction_work(
            source
                .source_bytes()
                .len()
                .saturating_add(candidate.source_bytes().len()),
            path,
        )?;
        let catalog = physical_source(candidate)?;
        let entries = catalog.package().iter().count();
        let total_entry_bytes = catalog.package().iter().fold(0usize, |total, entry| {
            total.saturating_add(entry.data().len())
        });
        self.charge_entries(entries, path)?;
        self.charge_total_entry_bytes(total_entry_bytes, path)?;
        self.charge_allocations(2, path)?;
        self.preflight_inventory(candidate, path)
    }

    fn consume_ownership_report(
        &mut self,
        report: table_headers::ownership::OwnershipReport,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_work(report.work, path)?;
        self.charge_references(report.references, path)?;
        self.charge_transaction_work(report.transaction_work, path)
    }
}

/// One fully validated immutable publication.
#[must_use = "a persisted-sort commit contains the validated package snapshot"]
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

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &Diagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted table's persisted sort configuration.
    pub fn table_sort_order<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Option<Order>, Error> {
        let mut budget = TransactionBudget::new(self);
        budget.preflight_source(self, Path::Package)?;
        Ok(resolve(self, sheet, table, &mut budget)?.before)
    }

    /// Start a selector-first immutable persisted-sort edit.
    pub fn edit_table_sort_order<'sheet, 'table>(
        &self,
        sheet: impl Into<SheetSelector<'sheet>>,
        table: impl Into<TableSelector<'table>>,
    ) -> Result<Edit<'_>, Error> {
        let mut budget = TransactionBudget::new(self);
        budget.preflight_source(self, Path::Package)?;
        let selected = resolve(self, sheet, table, &mut budget)?;
        Ok(Edit {
            source: self,
            target: selected.clone(),
            after: selected.before.clone(),
        })
    }

    /// Apply a reversible exact-source persisted-sort patch.
    pub fn apply_table_sort_order(&self, patch: &Patch) -> Result<Commit, Error> {
        let catalog = physical_source(self)?;
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
        let mut budget = TransactionBudget::new(self);
        budget.preflight_source(self, Path::Package)?;
        let current = resolve_at(
            self,
            patch.target.native.sheet_position,
            patch.target.native.table_position,
            &mut budget,
        )?;
        if current.before != patch.before
            || current.native.model_identifier != patch.target.native.model_identifier
        {
            return Err(Error::PatchConflict);
        }
        let target = patch.artifacts.target_owner();
        budget.charge_output_bytes(target.as_ref().len(), Path::Package)?;
        budget.preflight_candidate_reopen(
            target.as_ref().len(),
            self.state.components.catalog().len(),
            Path::Package,
        )?;
        budget.preflight_preview_scan(self, Path::Package)?;
        let source_previews =
            table_headers::rewrite::root_preview_deletions(catalog).map_err(map_header_error)?;
        let candidate = Package::from_source_owner_with_options(target, self.state.options)
            .map_err(|_error| Error::Verification)?;
        let selected = resolve_at(
            &candidate,
            patch.target.native.sheet_position,
            patch.target.native.table_position,
            &mut budget,
        )?;
        if selected.before != patch.after
            || selected.native.model_identifier != patch.target.native.model_identifier
        {
            return Err(Error::Verification);
        }
        let expected_payload = selected_payload(&candidate, selected.native)?;
        budget.preflight_locality(self, &candidate, Path::Package)?;
        budget.preflight_preview_scan(&candidate, Path::Package)?;
        table_headers::rewrite::verify_exact_locality(
            self,
            &candidate,
            patch.target.native,
            &source_previews,
            source_previews.len(),
            expected_payload,
        )
        .map_err(map_header_error)?;
        verify_unchanged_entries(self, &candidate, patch.target.native)?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(),
        })
    }
}

fn resolve<'sheet, 'table>(
    source: &Package,
    sheet: impl Into<SheetSelector<'sheet>>,
    table: impl Into<TableSelector<'table>>,
    budget: &mut TransactionBudget,
) -> Result<Target, Error> {
    let selected_sheet = source
        .state
        .document
        .sheet(sheet)
        .map_err(|_error| invalid_source())?
        .ok_or(Error::SheetNotFound)?;
    let table_position = match table.into() {
        TableSelector::Index(index) => selected_sheet.tables().nth(index).map(|_| index),
        TableSelector::Name(name) => {
            let mut matches = selected_sheet
                .tables()
                .enumerate()
                .filter(|(_, candidate)| candidate.name() == name);
            let first = matches.next().map(|(index, _)| index);
            if matches.next().is_some() {
                return Err(invalid_source());
            }
            first
        },
    }
    .ok_or(Error::TableNotFound)?;
    resolve_at(source, selected_sheet.index(), table_position, budget)
}

fn resolve_at(
    source: &Package,
    sheet: usize,
    table: usize,
    budget: &mut TransactionBudget,
) -> Result<Target, Error> {
    let native =
        table_headers::resolve::resolve_target(source, sheet, table).map_err(map_header_error)?;
    if native.message_type != TABLE_MODEL_MESSAGE_TYPE {
        return Err(invalid_source());
    }
    let model = selected_payload(source, native)?;
    let order = decode_model_sort(model, native.columns, budget, Path::Table { sheet, table })?;
    if let Some(order) = &order {
        validate_order_columns(order, native.columns)?;
    }
    // The source model must have a complete ArchiveInfo/message alignment; a
    // later changed transition must never guess around stale metadata.
    let object = source
        .state
        .components
        .catalog()
        .get_index(native.component_index)
        .and_then(|component| component.archive().objects.get(native.object_index))
        .ok_or_else(invalid_source)?;
    table_headers::resolve::validate_message_metadata(object, native.message_index)
        .map_err(map_header_error)?;
    validate_selected_ownership(source, native, budget, Path::Table { sheet, table })?;
    Ok(Target {
        native,
        before: order,
    })
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    if edit.target.before == edit.after {
        let source = physical_source(edit.source)?.__source_owner();
        return Ok(Commit {
            package: edit.source.snapshot(),
            patch: Patch {
                artifacts: OwnedExactArtifacts::new(source.clone(), source),
                target: edit.target.clone(),
                before: edit.target.before.clone(),
                after: edit.after,
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    let mut budget = TransactionBudget::new(edit.source);
    budget.preflight_source(edit.source, edit.path())?;
    if edit.target.native.locked == LockState::Locked {
        return Err(Error::TableLocked { path: edit.path() });
    }
    if let Some(order) = &edit.after {
        validate_order_columns(order, edit.target.native.columns)?;
    }
    let catalog = physical_source(edit.source)?;
    if !catalog.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    validate_selected_ownership(edit.source, edit.target.native, &mut budget, edit.path())?;
    let target = edit.target.clone();
    budget.preflight_preview_scan(edit.source, edit.path())?;
    let source_previews =
        table_headers::rewrite::root_preview_deletions(catalog).map_err(map_header_error)?;
    let package = rewrite(edit.source, target, edit.after.clone(), &mut budget)?;
    let selected = resolve_at(
        &package,
        edit.target.native.sheet_position,
        edit.target.native.table_position,
        &mut budget,
    )?;
    if selected.before != edit.after {
        return Err(Error::Verification);
    }
    let expected_payload = selected_payload(&package, selected.native)?;
    budget.preflight_locality(edit.source, &package, edit.path())?;
    budget.preflight_preview_scan(&package, edit.path())?;
    table_headers::rewrite::verify_exact_locality(
        edit.source,
        &package,
        edit.target.native,
        &source_previews,
        source_previews.len(),
        expected_payload,
    )
    .map_err(map_header_error)?;
    verify_unchanged_entries(edit.source, &package, edit.target.native)?;
    let source = catalog.__source_owner();
    let target = physical_source(&package)?.__source_owner();
    Ok(Commit {
        package,
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(source, target),
            target: edit.target.clone(),
            before: edit.target.before.clone(),
            after: edit.after,
        },
        diagnostics: Diagnostics::published(),
    })
}

fn decode_model_sort(
    source: &[u8],
    columns: u32,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Option<Order>, Error> {
    let options = budget.codec_options(source, columns);
    let (snapshot, report) = codec::decode_table_model_sort_order_with_report(source, options)
        .map_err(map_codec_error)?;
    budget.consume_codec_report(report, path)?;
    let order = snapshot.map(order_from_snapshot).transpose()?;
    if let Some(order) = &order {
        validate_order_columns(order, columns)?;
    }
    Ok(order)
}

fn order_from_snapshot(snapshot: codec::SortOrderSnapshot) -> Result<Order, Error> {
    let scope = match snapshot.scope() {
        codec::SortScope::EntireTable => Scope::EntireTable,
        codec::SortScope::SelectedRows => Scope::SelectedRows,
    };
    let rules = snapshot
        .rules()
        .iter()
        .map(|rule| {
            let column = crate::table::sort::ColumnIndex::from_native(rule.column())
                .map_err(|_| invalid_source())?;
            let direction = match rule.direction() {
                codec::SortDirection::Ascending => Direction::Ascending,
                codec::SortDirection::Descending => Direction::Descending,
            };
            Ok(Rule::new(column, direction))
        })
        .collect::<Result<Vec<_>, Error>>()?;
    Order::with_scope(scope, rules).map_err(|_| invalid_source())
}

fn snapshot_from_order(order: &Order) -> Result<codec::SortOrderSnapshot, Error> {
    let scope = match order.scope() {
        Scope::EntireTable => codec::SortScope::EntireTable,
        Scope::SelectedRows => codec::SortScope::SelectedRows,
    };
    let rules = order
        .rules()
        .iter()
        .map(|rule| {
            let direction = match rule.direction() {
                Direction::Ascending => codec::SortDirection::Ascending,
                Direction::Descending => codec::SortDirection::Descending,
            };
            codec::SortRule::new(rule.column().native_value(), direction)
        })
        .collect::<Vec<_>>();
    codec::SortOrderSnapshot::new(scope, rules).map_err(map_codec_error)
}

fn rewrite_model_sort(
    source: &[u8],
    after: Option<&Order>,
    columns: u32,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let desired = after.map(snapshot_from_order).transpose()?;
    let options = budget.codec_options(source, columns);
    let prepared = codec::prepare_table_model_sort_order_rewrite(source, desired, options)
        .map_err(map_codec_error)?;
    let requirements = prepared.execution_requirements();
    budget.preflight_codec_requirements(requirements, source.len(), path)?;
    let output = prepared
        // Policy: codec preparation and execution are a staged-private phase.
        // Its conservative requirements are charged before the codec output
        // Vec is reserved, but the native archive/ZIP candidate remains
        // private until reassembly, reopen, and locality verification finish.
        // This is an atomic-publication guarantee, not a single-global
        // no-allocation or single-allocation guarantee.
        .execute(requirements.exact_limits())
        .map_err(map_codec_error)?;
    let report = output.report();
    if report.output_bytes() != requirements.output_bytes
        || report.fields() != requirements.fields
        || report.work_bytes() != requirements.work_bytes
        || report.rules() != requirements.rules
        || report.allocations() != requirements.allocations
        || report.retained_bytes() != requirements.retained_bytes
        || report.scratch_bytes() != requirements.scratch_bytes
        || report.max_depth() != requirements.max_depth
    {
        return Err(Error::Verification);
    }
    Ok(output.into_bytes())
}

fn validate_order_columns(order: &Order, columns: u32) -> Result<(), Error> {
    if order
        .rules()
        .iter()
        .any(|rule| u64::from(rule.column().native_value()) >= u64::from(columns))
    {
        return Err(invalid_source());
    }
    Ok(())
}

fn rewrite(
    source: &Package,
    target: Target,
    after: Option<Order>,
    budget: &mut TransactionBudget,
) -> Result<Package, Error> {
    let catalog = physical_source(source)?;
    let component = source
        .state
        .components
        .catalog()
        .get_index(target.native.component_index)
        .ok_or_else(invalid_source)?;
    let component_name = component.name();
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .ok_or_else(invalid_source)?;
    if entry.is_opaque() {
        return Err(Error::UnsupportedSource);
    }
    table_headers::rewrite::preflight_transaction_work(source, None).map_err(map_header_error)?;
    let physical_limits = catalog.limits();
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    budget.preflight_physical_entry(entry.data().len(), archive_limits, Path::Package)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical_limits.snappy_limits().map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    budget.charge_payload_bytes(stream.as_bytes().len(), Path::Package)?;
    budget.charge_total_payload_bytes(stream.as_bytes().len(), Path::Package)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    let object = archive
        .objects
        .get_mut(target.native.object_index)
        .ok_or_else(invalid_source)?;
    if object.archive_info.identifier != Some(target.native.model_identifier) {
        return Err(invalid_source());
    }
    table_headers::resolve::validate_message_metadata(object, target.native.message_index)
        .map_err(map_header_error)?;
    let original = object
        .messages
        .get(target.native.message_index)
        .ok_or_else(invalid_source)?
        .data
        .clone();
    if decode_model_sort(&original, target.native.columns, budget, Path::Package)? != target.before
    {
        return Err(invalid_source());
    }
    let replacement = rewrite_model_sort(
        &original,
        after.as_ref(),
        target.native.columns,
        budget,
        Path::Package,
    )?;
    object
        .replace_message_preserving_header_with_limits(
            target.native.message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: replacement,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_bound = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget.charge_transaction_work(encoded_bound, Path::Package)?;
    let rewritten = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(rewritten.len()).map_err(map_core_error)?;
    let snappy_limits = physical_limits.snappy_limits().map_err(map_archive_error)?;
    if compressed_bound > snappy_limits.max_compressed_stream() {
        return Err(Error::LimitExceeded {
            kind: LimitKind::EntryBytes,
            observed: u64::try_from(compressed_bound).unwrap_or(u64::MAX),
            maximum: u64::try_from(snappy_limits.max_compressed_stream()).unwrap_or(u64::MAX),
            path: Path::Package,
        });
    }
    budget.charge_transaction_work(compressed_bound, Path::Package)?;
    budget.charge_scratch_bytes(compressed_bound, Path::Package)?;
    let compressed = SnappyStream::compress(&rewritten).map_err(map_core_error)?;
    let edits = [EntryEdit::new(component_name, compressed.as_slice())];
    let prepared = catalog
        .package()
        .prepare_reassembly_with_deletions(&edits, &[], physical_limits)
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget.preflight_reassembly(requirements, Path::Package)?;
    budget.preflight_candidate_reopen(
        requirements.output_bytes(),
        source.state.components.catalog().len(),
        Path::Package,
    )?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    Package::from_owned_bytes_with_options(output, source.state.options)
        .map_err(|_error| Error::Verification)
}

fn selected_payload(source: &Package, target: table_headers::Target) -> Result<&[u8], Error> {
    table_headers::rewrite::selected_payload(source, target).map_err(map_header_error)
}

fn validate_selected_ownership(
    source: &Package,
    target: table_headers::Target,
    budget: &mut TransactionBudget,
    path: Path,
) -> Result<(), Error> {
    let report = table_headers::ownership::validate_selected_ownership_with_report(source, target)
        .map_err(map_header_error)?;
    budget.consume_ownership_report(report, path)?;
    Ok(())
}

fn physical_source(source: &Package) -> Result<&litchi_iwa_archive::SourceCatalog, Error> {
    table_headers::rewrite::physical_source(source).map_err(map_header_error)
}

fn verify_unchanged_entries(
    source: &Package,
    candidate: &Package,
    target: table_headers::Target,
) -> Result<(), Error> {
    let source_catalog = physical_source(source)?;
    let candidate_catalog = physical_source(candidate)?;
    if source_catalog.package().iter().count() != candidate_catalog.package().iter().count() {
        return Err(Error::Verification);
    }
    let source_index =
        PhysicalEntryIndex::new(source_catalog.package()).map_err(|error| match error {
            PhysicalEntryIndexError::Allocation { .. } | PhysicalEntryIndexError::Duplicate => {
                Error::Verification
            },
        })?;
    let candidate_index =
        PhysicalEntryIndex::new(candidate_catalog.package()).map_err(|error| match error {
            PhysicalEntryIndexError::Allocation { .. } | PhysicalEntryIndexError::Duplicate => {
                Error::Verification
            },
        })?;
    let selected_name = source
        .state
        .components
        .catalog()
        .get_index(target.component_index)
        .ok_or(Error::Verification)?
        .name();
    for before in source_catalog.package().iter() {
        let after = candidate_index
            .get(before.name())
            .ok_or(Error::Verification)?;
        if before.name() == selected_name {
            continue;
        }
        if before.raw_name() != after.raw_name()
            || before.is_opaque() != after.is_opaque()
            || before.data() != after.data()
            || before.raw_record().local_record() != after.raw_record().local_record()
            || !central_record_preserved_except_offset(
                before.raw_record().central_directory_record(),
                after.raw_record().central_directory_record(),
            )
        {
            return Err(Error::Verification);
        }
    }
    for after in candidate_catalog.package().iter() {
        if source_index.get(after.name()).is_none() {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

fn central_record_preserved_except_offset(source: &[u8], candidate: &[u8]) -> bool {
    const OFFSET: std::ops::Range<usize> = 42..46;
    source.len() == candidate.len()
        && source.len() >= OFFSET.end
        && source[..OFFSET.start] == candidate[..OFFSET.start]
        && source[OFFSET.end..] == candidate[OFFSET.end..]
}

fn invalid_source() -> Error {
    Error::InvalidSource {
        path: Path::Package,
    }
}

fn map_header_error(error: table_headers::Error) -> Error {
    match error {
        table_headers::Error::SheetNotFound => Error::SheetNotFound,
        table_headers::Error::TableNotFound => Error::TableNotFound,
        table_headers::Error::UnsupportedSource => Error::UnsupportedSource,
        table_headers::Error::TableLocked { path } => {
            let table_headers::Path::Table { sheet, table } = path else {
                return invalid_source();
            };
            Error::TableLocked {
                path: Path::Table { sheet, table },
            }
        },
        table_headers::Error::Allocation { amount, .. } => Error::Allocation {
            amount,
            path: Path::Package,
        },
        table_headers::Error::LimitExceeded {
            kind,
            observed,
            maximum,
            ..
        } => Error::LimitExceeded {
            kind: match kind {
                table_headers::LimitKind::InputBytes => LimitKind::InputBytes,
                table_headers::LimitKind::OutputBytes => LimitKind::OutputBytes,
                table_headers::LimitKind::Entries => LimitKind::Entries,
                table_headers::LimitKind::EntryBytes => LimitKind::EntryBytes,
                table_headers::LimitKind::TotalEntryBytes => LimitKind::TotalEntryBytes,
                table_headers::LimitKind::PackageBytes => LimitKind::PackageBytes,
                table_headers::LimitKind::PayloadBytes => LimitKind::PayloadBytes,
                table_headers::LimitKind::TotalPayloadBytes => LimitKind::TotalPayloadBytes,
                table_headers::LimitKind::PayloadObjects => LimitKind::PayloadObjects,
                table_headers::LimitKind::PayloadMessages => LimitKind::PayloadMessages,
                table_headers::LimitKind::PayloadItems => LimitKind::PayloadItems,
                table_headers::LimitKind::PayloadReferences => LimitKind::PayloadReferences,
                table_headers::LimitKind::WireBytes => LimitKind::WireBytes,
                table_headers::LimitKind::WireOutputBytes => LimitKind::WireOutputBytes,
                table_headers::LimitKind::WireFields => LimitKind::WireFields,
                table_headers::LimitKind::WireNesting => LimitKind::WireNesting,
                table_headers::LimitKind::WireWork => LimitKind::WireWork,
                table_headers::LimitKind::TransactionWork => LimitKind::TransactionWork,
            },
            observed,
            maximum,
            path: Path::Package,
        },
        table_headers::Error::PatchConflict => Error::PatchConflict,
        table_headers::Error::Verification => Error::Verification,
        _ => invalid_source(),
    }
}

fn map_codec_error(error: codec::DecodeError) -> Error {
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation {
            amount,
            path: Path::Package,
        };
    }
    let Some(limit) = error.resource_limit() else {
        return invalid_source();
    };
    let (kind, observed, maximum) = match limit {
        codec::DecodeLimit::InputBytes { observed, maximum } => {
            (LimitKind::WireBytes, observed, maximum)
        },
        codec::DecodeLimit::OutputBytes { observed, maximum } => {
            (LimitKind::WireOutputBytes, observed, maximum)
        },
        codec::DecodeLimit::Fields { observed, maximum } => {
            (LimitKind::WireFields, observed, maximum)
        },
        codec::DecodeLimit::WorkBytes { observed, maximum } => {
            (LimitKind::WireWork, observed, maximum)
        },
        codec::DecodeLimit::Nesting { observed, maximum } => {
            return Error::LimitExceeded {
                kind: LimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
                path: Path::Package,
            };
        },
        codec::DecodeLimit::Rules { observed, maximum } => {
            (LimitKind::WireRules, observed, maximum)
        },
        codec::DecodeLimit::Columns { observed, maximum } => {
            (LimitKind::WireColumns, observed, maximum)
        },
        codec::DecodeLimit::Allocations { observed, maximum } => {
            (LimitKind::WireAllocations, observed, maximum)
        },
        codec::DecodeLimit::ScratchBytes { observed, maximum } => {
            (LimitKind::WireScratchBytes, observed, maximum)
        },
        codec::DecodeLimit::RetainedBytes { observed, maximum } => {
            (LimitKind::WireRetainedBytes, observed, maximum)
        },
        _ => return invalid_source(),
    };
    Error::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        path: Path::Package,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> Error {
    match error {
        ArchiveError::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                ArchiveLimitKind::InputBytes => LimitKind::InputBytes,
                ArchiveLimitKind::OutputBytes => LimitKind::OutputBytes,
                ArchiveLimitKind::Entries => LimitKind::Entries,
                ArchiveLimitKind::MemberNameBytes | ArchiveLimitKind::MetadataBytes => {
                    LimitKind::PackageBytes
                },
                ArchiveLimitKind::CompressedEntryBytes => LimitKind::EntryBytes,
                ArchiveLimitKind::EntryBytes => LimitKind::EntryBytes,
                ArchiveLimitKind::TotalBytes => LimitKind::TotalEntryBytes,
                ArchiveLimitKind::IwaStreamBytes => LimitKind::PayloadBytes,
                ArchiveLimitKind::IwaTotalBytes => LimitKind::TotalPayloadBytes,
            },
            observed,
            maximum,
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation {
            amount,
            path: Path::Package,
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => invalid_source(),
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> Error {
    match error {
        CoreError::Limit {
            kind,
            observed,
            maximum,
        } => Error::LimitExceeded {
            kind: match kind {
                CoreLimitKind::ArchiveBytes => LimitKind::TotalPayloadBytes,
                CoreLimitKind::Objects => LimitKind::PayloadObjects,
                CoreLimitKind::Messages | CoreLimitKind::MessagesPerObject => {
                    LimitKind::PayloadMessages
                },
                CoreLimitKind::ObjectBytes
                | CoreLimitKind::MessageBytes
                | CoreLimitKind::HeaderBytes
                | CoreLimitKind::HeaderMemoryBytes => LimitKind::PayloadBytes,
                CoreLimitKind::HeaderFields => LimitKind::PayloadItems,
                CoreLimitKind::HeaderNesting => LimitKind::WireNesting,
                CoreLimitKind::MetadataItems => LimitKind::PayloadItems,
                CoreLimitKind::SnappyChunkBytes => LimitKind::PayloadBytes,
                CoreLimitKind::SnappyStreamBytes => LimitKind::TotalPayloadBytes,
                CoreLimitKind::SnappyCompressedChunkBytes => LimitKind::EntryBytes,
                CoreLimitKind::SnappyCompressedStreamBytes => LimitKind::TotalEntryBytes,
                CoreLimitKind::SnappyFrames => LimitKind::PayloadItems,
            },
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path: Path::Package,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => Error::Allocation {
            amount: requested,
            path: Path::Package,
        },
        _ => invalid_source(),
    }
}

fn charge_budget(
    remaining: &mut usize,
    maximum: usize,
    amount: usize,
    kind: LimitKind,
    path: Path,
) -> Result<(), Error> {
    let consumed = maximum.saturating_sub(*remaining);
    let observed = consumed.saturating_add(amount);
    if observed > maximum {
        return Err(Error::LimitExceeded {
            kind,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
            path,
        });
    }
    *remaining = (*remaining).saturating_sub(amount);
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixture() -> std::path::PathBuf {
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/iwork/numbers/basic.numbers")
    }

    #[test]
    fn exact_noops_skip_preflight_and_malformed_changed_paths_still_fail()
    -> Result<(), Box<dyn std::error::Error>> {
        let package = Package::open(fixture())?;
        let source = package.source_bytes().to_vec();
        let edit = package.edit_table_sort_order(0usize, 0usize)?;
        let current = edit.order().cloned();

        phase_observer::reset();
        let commit = match current.clone() {
            Some(order) => edit.set(order).commit()?,
            None => edit.clear().commit()?,
        };
        assert_eq!(phase_observer::preflight_count(), 0);
        assert!(commit.patch().is_noop());
        assert!(commit.package().shares_snapshot(&package));
        assert_eq!(commit.package().source_bytes(), source.as_slice());
        assert!(!commit.diagnostics().changed());
        assert_eq!(commit.diagnostics().touched_components(), 0);
        assert!(!commit.diagnostics().full_reparse_performed());

        let mut malformed_noop = commit.patch().clone();
        malformed_noop.target.native.sheet_position = usize::MAX;
        malformed_noop.target.native.table_position = usize::MAX;
        phase_observer::reset();
        let applied = package.apply_table_sort_order(&malformed_noop)?;
        assert_eq!(phase_observer::preflight_count(), 0);
        assert!(applied.package().shares_snapshot(&package));
        assert_eq!(applied.package().source_bytes(), source.as_slice());
        assert!(!applied.diagnostics().changed());
        assert_eq!(applied.diagnostics().touched_components(), 0);
        assert!(!applied.diagnostics().full_reparse_performed());

        let changed = match current {
            Some(_) => package
                .edit_table_sort_order(0usize, 0usize)?
                .clear()
                .commit()?,
            None => package
                .edit_table_sort_order(0usize, 0usize)?
                .set(Order::new([Rule::new(
                    crate::table::sort::ColumnIndex::new(0)?,
                    Direction::Ascending,
                )])?)
                .commit()?,
        };
        assert!(!changed.patch().is_noop());
        let mut malformed_changed = changed.patch().clone();
        malformed_changed.target.native.sheet_position = usize::MAX;
        malformed_changed.target.native.table_position = usize::MAX;
        phase_observer::reset();
        assert!(package.apply_table_sort_order(&malformed_changed).is_err());
        assert!(phase_observer::preflight_count() > 0);
        assert_eq!(package.source_bytes(), source.as_slice());
        Ok(())
    }
}
