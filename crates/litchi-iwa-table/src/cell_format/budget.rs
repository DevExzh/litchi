//! The operation-local resource ledger of a table-cell format transaction.

use litchi_iwa_archive::package::ReassemblyExecutionRequirements;
use litchi_iwa_core::Archive;
use litchi_iwa_protos::{
    numbers_table_cell_control_codec as control_codec,
    numbers_table_cell_pop_up_menu_codec as popup_codec,
    numbers_table_cell_storage_codec as storage_codec,
};

use super::{Error, LimitKind, Path};
use crate::TableSource;

/// One operation-local ledger shared by the rooted catalog, native graph,
/// metadata sidecar, ZIP reassembly, and candidate reopen.  The lower codecs
/// are deliberately given residual ceilings from this ledger; each distinct
/// scan consumes its report exactly once.
#[derive(Debug, Clone, Copy)]
pub struct TransactionBudget {
    pub(crate) max_input_bytes: usize,
    pub(crate) max_output_bytes: usize,
    pub(crate) max_entries: usize,
    pub(crate) max_entry_bytes: usize,
    pub(crate) max_total_entry_bytes: usize,
    pub(crate) max_payload_bytes: usize,
    pub(crate) max_payload_objects: usize,
    pub(crate) max_payload_messages: usize,
    pub(crate) max_payload_items: usize,
    pub(crate) max_payload_references: usize,
    pub(crate) max_wire_bytes: usize,
    pub(crate) max_wire_nesting: u32,
    pub(crate) max_wire_reference_bytes: usize,
    pub(crate) max_wire_text_bytes: usize,
    pub(crate) max_wire_fields: usize,
    pub(crate) max_wire_work: usize,
    pub(crate) max_scratch_bytes: usize,
    pub(crate) max_retained_bytes: usize,
    pub(crate) max_compressed_bytes: usize,
    pub(crate) max_candidate_input_bytes: usize,
    pub(crate) max_transaction_work: usize,
    pub(crate) max_allocations: usize,
    pub(crate) remaining_input_bytes: usize,
    pub(crate) remaining_output_bytes: usize,
    pub(crate) remaining_entries: usize,
    pub(crate) remaining_total_entry_bytes: usize,
    pub(crate) remaining_payload_bytes: usize,
    pub(crate) remaining_payload_objects: usize,
    pub(crate) remaining_payload_messages: usize,
    pub(crate) remaining_payload_items: usize,
    pub(crate) remaining_payload_references: usize,
    pub(crate) remaining_wire_bytes: usize,
    pub(crate) remaining_wire_nesting: u32,
    pub(crate) remaining_wire_reference_bytes: usize,
    pub(crate) remaining_wire_text_bytes: usize,
    pub(crate) remaining_wire_fields: usize,
    pub(crate) remaining_wire_work: usize,
    pub(crate) remaining_scratch_bytes: usize,
    pub(crate) remaining_retained_bytes: usize,
    pub(crate) remaining_compressed_bytes: usize,
    pub(crate) remaining_candidate_input_bytes: usize,
    pub(crate) remaining_transaction_work: usize,
    pub(crate) remaining_allocations: usize,
    pub(crate) list_keys: ListKeys,
}

/// How a format or control list (`TST.TableDataList`) allocates and
/// releases entry keys.
///
/// The Numbers owners reuse the lowest free key below the allocation
/// watermark (`next_list_id`) and keep the watermark on release. The Pages
/// and Keynote owners follow the shared cell writer
/// (`litchi_numbers_wire::table_write`) and the legacy host format writer:
/// a new entry takes the watermark, and a removed last key is handed back
/// (record 0379).
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ListKeys {
    /// Reuse free keys; keep the watermark on release.
    #[default]
    Reuse,
    /// Append at the watermark; hand a removed last key back.
    Append,
}

impl TransactionBudget {
    pub fn new(source: &dyn TableSource) -> Self {
        let archive = source.archive_limits();
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
        let max_payload_objects = source.max_objects().max(1);
        let max_payload_references = source.max_references().max(1);
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
            list_keys: ListKeys::Reuse,
        }
    }

    /// Select how list entry keys are allocated and released.
    #[must_use]
    pub const fn with_list_keys(mut self, policy: ListKeys) -> Self {
        self.list_keys = policy;
        self
    }

    /// Construct the aggregate ledger used by the unified scalar-control
    /// route. That route scans the model, data store, tiles, both co-located
    /// lists, every mixed CellSpec, and Metadata under one transaction. The
    /// archive's IWA-stream ceiling is a per-message bound, so summing all of
    /// those distinct reports against it would incorrectly reject valid
    /// packages. The transaction-work ceiling remains the outer aggregate
    /// bound for those scans.
    pub fn for_cell_control(source: &dyn TableSource) -> Self {
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
            return Err(Error::Limit {
                kind,
                observed: u64::try_from(observed).unwrap_or(u64::MAX),
                maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
                path,
            });
        }
        *remaining -= amount;
        Ok(())
    }

    pub fn charge_input(&mut self, amount: usize, path: Path) -> Result<(), Error> {
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
    pub fn charge_candidate_input_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_candidate_input_bytes,
            self.max_candidate_input_bytes,
            amount,
            LimitKind::InputBytes,
            path,
        )
    }

    pub fn charge_output(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_output_bytes,
            self.max_output_bytes,
            amount,
            LimitKind::OutputBytes,
            path,
        )
    }

    pub fn charge_entries(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_entries,
            self.max_entries,
            amount,
            LimitKind::Entries,
            path,
        )
    }

    pub fn charge_entry_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        if amount > self.max_entry_bytes {
            return Err(Error::Limit {
                kind: LimitKind::EntryBytes,
                observed: u64::try_from(amount).unwrap_or(u64::MAX),
                maximum: u64::try_from(self.max_entry_bytes).unwrap_or(u64::MAX),
                path,
            });
        }
        Ok(())
    }

    pub fn charge_total_entry_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_total_entry_bytes,
            self.max_total_entry_bytes,
            amount,
            LimitKind::TotalEntryBytes,
            path,
        )
    }

    pub fn charge_payload_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_payload_bytes,
            self.max_payload_bytes,
            amount,
            LimitKind::PayloadBytes,
            path,
        )
    }

    pub fn charge_payload_objects(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_payload_objects,
            self.max_payload_objects,
            amount,
            LimitKind::PayloadObjects,
            path,
        )
    }

    pub fn charge_payload_messages(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_payload_messages,
            self.max_payload_messages,
            amount,
            LimitKind::PayloadMessages,
            path,
        )
    }

    pub fn charge_payload_items(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_payload_items,
            self.max_payload_items,
            amount,
            LimitKind::PayloadItems,
            path,
        )
    }

    pub fn charge_payload_references(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_payload_references,
            self.max_payload_references,
            amount,
            LimitKind::PayloadReferences,
            path,
        )
    }

    pub fn charge_wire_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_wire_bytes,
            self.max_wire_bytes,
            amount,
            LimitKind::WireBytes,
            path,
        )
    }

    pub fn charge_wire_nesting(&mut self, depth: u32, path: Path) -> Result<(), Error> {
        let observed = self
            .max_wire_nesting
            .saturating_sub(self.remaining_wire_nesting);
        if depth > self.max_wire_nesting {
            return Err(Error::Limit {
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

    pub fn charge_wire_reference_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_wire_reference_bytes,
            self.max_wire_reference_bytes,
            amount,
            LimitKind::WireReferenceBytes,
            path,
        )
    }

    pub fn charge_wire_text_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_wire_text_bytes,
            self.max_wire_text_bytes,
            amount,
            LimitKind::WireTextBytes,
            path,
        )
    }

    pub fn charge_wire_fields(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_wire_fields,
            self.max_wire_fields,
            amount,
            LimitKind::WireFields,
            path,
        )
    }

    pub fn charge_wire_work(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_wire_work,
            self.max_wire_work,
            amount,
            LimitKind::WireWork,
            path,
        )
    }

    pub fn charge_scratch_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_scratch_bytes,
            self.max_scratch_bytes,
            amount,
            LimitKind::ScratchBytes,
            path,
        )
    }

    pub fn charge_retained_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_retained_bytes,
            self.max_retained_bytes,
            amount,
            LimitKind::RetainedBytes,
            path,
        )
    }

    pub fn charge_compressed_bytes(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_compressed_bytes,
            self.max_compressed_bytes,
            amount,
            LimitKind::CompressedBytes,
            path,
        )
    }

    pub fn charge_transaction_work(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_transaction_work,
            self.max_transaction_work,
            amount,
            LimitKind::TransactionWork,
            path,
        )
    }

    pub fn charge_allocations(&mut self, amount: usize, path: Path) -> Result<(), Error> {
        Self::charge(
            &mut self.remaining_allocations,
            self.max_allocations,
            amount,
            LimitKind::Allocations,
            path,
        )
    }

    /// Admit one allocation-free physical candidate-ingress phase before the
    /// corresponding ZIP or decoder work starts.  The archive seam reports
    /// retained storage, traversal work, and fallible allocation events as
    /// separate dimensions; keeping those dimensions distinct preserves the
    /// cell transaction's typed refusal instead of charging the same bytes as
    /// a second candidate input scan.
    pub fn charge_candidate_ingress(
        &mut self,
        charge: litchi_iwa_archive::unstable::CandidateIngressCharge,
        path: Path,
    ) -> Result<(), Error> {
        self.charge_retained_bytes(charge.retained_bytes(), path)?;
        self.charge_transaction_work(charge.traversal_bytes(), path)?;
        self.charge_allocations(charge.allocations(), path)
    }

    pub fn preflight_reassembly(
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

    pub fn charge_package_source(
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
    pub fn charge_candidate_reopen(
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

    pub fn charge_archive(
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

    pub fn residual_storage_options(&self, source: &[u8]) -> storage_codec::DecodeOptions {
        storage_codec::DecodeOptions::new(
            self.remaining_wire_bytes.min(source.len().max(1)),
            self.remaining_wire_work.max(1),
            self.remaining_wire_fields.max(1),
            self.max_wire_nesting,
            self.remaining_payload_references.max(1),
            self.remaining_payload_references.max(1),
        )
    }

    pub fn residual_popup_options(&self, source: &[u8]) -> popup_codec::DecodeOptions {
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

    pub fn charge_storage_decode_report(
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

    pub fn residual_control_options_for_len(
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
    pub fn residual_storage_rewrite_options(&self, source: &[u8]) -> storage_codec::DecodeOptions {
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
