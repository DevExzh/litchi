//! Operation-local resource accounting for focused slide-media access.
//!
//! The package ingress already bounds the ZIP and IWA source.  A replacement
//! operation still performs several bounded projections over that source and
//! can retain a replacement buffer, a metadata candidate, and a reassembled
//! package at the same time.  `MediaBudget` keeps those costs on one ledger so
//! a sequence of individually valid lazy decodes cannot bypass the aggregate
//! semantic or wire profile.

use litchi_iwa_archive::{SourceCatalog, package::ReassemblyExecutionRequirements};
use litchi_iwa_protos::{
    keynote_media_codec,
    package_metadata_media_codec::{DecodeReport, RewriteExecutionRequirements},
};

use super::{Package, SlideMediaDataError, SlideMediaDataLimitKind};

/// Fallible operation-local budget for one media read, edit, commit, or patch
/// application.  The counters are deliberately private; callers can only
/// charge named resource categories with bounded arithmetic.
#[derive(Debug, Clone, Copy)]
pub(in crate::package) struct MediaBudget {
    max_input_bytes: usize,
    max_output_bytes: usize,
    max_entries: usize,
    max_entry_bytes: usize,
    max_total_bytes: usize,
    max_slides: usize,
    max_references: usize,
    max_media_bytes: usize,
    max_wire_fields: usize,
    max_wire_nesting: usize,
    max_wire_work: usize,
    max_allocations: usize,
    input_bytes: usize,
    output_bytes: usize,
    entries: usize,
    total_bytes: usize,
    slides: usize,
    references: usize,
    media_bytes: usize,
    wire_fields: usize,
    wire_nesting: usize,
    wire_work: usize,
    allocations: usize,
}

/// The portion of one replacement ledger that a sibling focused transaction
/// can debit against its own operation budget.  ZIP-entry and aggregate
/// physical counters stay private to the replacement owner; the media-data
/// adapter only exports work that has a corresponding semantic ceiling in the
/// consuming owner.
#[derive(Debug, Clone, Copy, Default)]
pub(in crate::package) struct MediaBudgetUsage {
    pub(in crate::package) input_bytes: usize,
    pub(in crate::package) fields: usize,
    pub(in crate::package) nesting: usize,
    pub(in crate::package) work: usize,
    pub(in crate::package) references: usize,
    pub(in crate::package) allocations: usize,
    pub(in crate::package) media_bytes: usize,
}

impl MediaBudget {
    /// Build a ledger from the exact profiles retained by one package.
    pub(super) fn for_package(package: &Package) -> Result<Self, SlideMediaDataError> {
        let physical = package.limits();
        let semantic = package.semantic_limits();
        let wire = package
            .semantic_wire_limits()
            .map_err(|_| SlideMediaDataError::Read)?;
        let max_input_bytes = aggregate(as_usize(physical.max_input_bytes())?)?;
        let max_output_bytes = aggregate(as_usize(physical.max_input_bytes())?)?;
        let max_total_bytes = aggregate(as_usize(physical.max_total_bytes())?)?;
        let max_entries = physical
            .max_entries()
            .checked_add(semantic.max_objects())
            .and_then(|value| value.checked_add(semantic.max_references()))
            .ok_or(SlideMediaDataError::InvalidSource)
            .and_then(aggregate)?;
        let max_allocations = semantic
            .max_objects()
            .checked_add(semantic.max_references())
            .and_then(|value| value.checked_mul(4))
            .ok_or(SlideMediaDataError::InvalidSource)?;
        Ok(Self {
            max_input_bytes,
            max_output_bytes,
            max_entries,
            max_entry_bytes: as_usize(physical.max_entry_bytes())?,
            max_total_bytes,
            max_slides: aggregate(semantic.max_slides())?,
            max_references: aggregate(semantic.max_references())?,
            max_media_bytes: max_total_bytes,
            max_wire_fields: aggregate(wire.max_fields())?,
            max_wire_nesting: wire.max_nesting(),
            max_wire_work: aggregate(wire.max_rewrite_work())?,
            max_allocations,
            input_bytes: 0,
            output_bytes: 0,
            entries: 0,
            total_bytes: 0,
            slides: 0,
            references: 0,
            media_bytes: 0,
            wire_fields: 0,
            wire_nesting: 0,
            wire_work: 0,
            allocations: 0,
        })
    }

    /// Build the replacement ledger with ceilings reserved by a sibling
    /// focused operation.  The physical package profile remains the hard
    /// upper bound; caller caps can only tighten it.  This constructor is
    /// intentionally private to the package so no public API can expose the
    /// replacement owner's accounting vocabulary.
    pub(in crate::package) fn for_package_with_caps(
        package: &Package,
        input_bytes: usize,
        fields: usize,
        nesting: usize,
        work: usize,
        references: usize,
        allocations: usize,
        media_bytes: usize,
    ) -> Result<Self, SlideMediaDataError> {
        let mut budget = Self::for_package(package)?;
        budget.max_input_bytes = budget.max_input_bytes.min(input_bytes);
        budget.max_wire_fields = budget.max_wire_fields.min(fields);
        budget.max_wire_nesting = budget.max_wire_nesting.min(nesting);
        budget.max_wire_work = budget.max_wire_work.min(work);
        budget.max_references = budget.max_references.min(references);
        budget.max_allocations = budget.max_allocations.min(allocations);
        budget.max_media_bytes = budget.max_media_bytes.min(media_bytes);
        Ok(budget)
    }

    /// Charge the retained physical catalog before any selected payload copy.
    pub(super) fn charge_catalog(
        &mut self,
        catalog: &SourceCatalog,
    ) -> Result<(), SlideMediaDataError> {
        self.input_bytes(catalog.source_bytes().len())?;
        self.entries(catalog.package().len())?;
        for entry in catalog.package().iter() {
            self.entry_bytes(entry.data().len())?;
            self.total_bytes(entry.data().len())?;
        }
        Ok(())
    }

    /// Charge one bounded source or nested wire input.
    pub(super) fn input_bytes(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        Self::charge(
            &mut self.input_bytes,
            self.max_input_bytes,
            amount,
            SlideMediaDataLimitKind::InputBytes,
        )
    }

    /// Charge one candidate output or reopened artifact.
    pub(super) fn output_bytes(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        Self::charge(
            &mut self.output_bytes,
            self.max_output_bytes,
            amount,
            SlideMediaDataLimitKind::OutputBytes,
        )
    }

    /// Charge retained ZIP/IWA entries.
    pub(super) fn entries(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        Self::charge(
            &mut self.entries,
            self.max_entries,
            amount,
            SlideMediaDataLimitKind::Entries,
        )
    }

    /// Check one member before a decompression or replacement allocation.
    pub(super) fn entry_bytes(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        if amount > self.max_entry_bytes {
            return Err(limit(
                SlideMediaDataLimitKind::EntryBytes,
                amount,
                self.max_entry_bytes,
            ));
        }
        Ok(())
    }

    /// Charge aggregate uncompressed ZIP bytes.
    pub(super) fn total_bytes(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        Self::charge(
            &mut self.total_bytes,
            self.max_total_bytes,
            amount,
            SlideMediaDataLimitKind::TotalBytes,
        )
    }

    /// Charge semantic slides traversed by the focused operation.
    pub(super) fn slides(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        Self::charge(
            &mut self.slides,
            self.max_slides,
            amount,
            SlideMediaDataLimitKind::Slides,
        )
    }

    /// Charge rooted graph references retained by the selector/closure.
    pub(super) fn references(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        Self::charge(
            &mut self.references,
            self.max_references,
            amount,
            SlideMediaDataLimitKind::References,
        )
    }

    /// Charge materialized media bytes before copying or retaining them.
    pub(super) fn media_bytes(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        Self::charge(
            &mut self.media_bytes,
            self.max_media_bytes,
            amount,
            SlideMediaDataLimitKind::MediaBytes,
        )
    }

    /// Charge one fallible allocation before calling `try_reserve`.
    pub(super) fn allocation(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        self.media_bytes(amount)?;
        let observed = self.allocations.checked_add(1).ok_or_else(|| {
            limit(
                SlideMediaDataLimitKind::Allocations,
                usize::MAX,
                self.max_allocations,
            )
        })?;
        if observed > self.max_allocations {
            return Err(limit(
                SlideMediaDataLimitKind::Allocations,
                observed,
                self.max_allocations,
            ));
        }
        self.allocations = observed;
        Ok(())
    }

    /// Charge a strict wire report produced by the Keynote data-reference
    /// codec before any caller-owned semantic value is retained.
    pub(super) fn keynote_report(
        &mut self,
        report: keynote_media_codec::DecodeReport,
    ) -> Result<(), SlideMediaDataError> {
        self.wire_fields(report.fields())?;
        self.wire_nesting(report.max_depth())?;
        self.wire_work(report.work_bytes())
    }

    /// Charge a strict PackageMetadata scan report.
    pub(super) fn metadata_report(
        &mut self,
        report: DecodeReport,
    ) -> Result<(), SlideMediaDataError> {
        self.input_bytes(report.input_bytes())?;
        self.wire_fields(report.fields())?;
        self.wire_nesting(report.max_depth())?;
        self.wire_work(report.work_bytes())?;
        self.entries(
            report
                .components()
                .checked_add(report.data_records())
                .ok_or(SlideMediaDataError::InvalidSource)?,
        )?;
        self.references(report.data_references())?;
        self.references(report.owners())
    }

    /// Build a metadata codec profile from the operation's remaining ledger.
    /// The codec must not reset its work ceiling for each lazy pass: a scan and
    /// its preservation pass consume the same transaction budget.
    pub(super) fn metadata_options(
        &self,
        source: &[u8],
    ) -> Result<litchi_iwa_protos::package_metadata_media_codec::DecodeOptions, SlideMediaDataError>
    {
        let fields = self.remaining(
            self.wire_fields,
            self.max_wire_fields,
            SlideMediaDataLimitKind::WireFields,
        )?;
        let work = self.remaining(
            self.wire_work,
            self.max_wire_work,
            SlideMediaDataLimitKind::WireWork,
        )?;
        let entries = self.remaining(
            self.entries,
            self.max_entries,
            SlideMediaDataLimitKind::Entries,
        )?;
        let owners = self.remaining(
            self.references,
            self.max_references,
            SlideMediaDataLimitKind::References,
        )?;
        let nesting = self.remaining(
            self.wire_nesting,
            self.max_wire_nesting,
            SlideMediaDataLimitKind::WireNesting,
        )?;
        let nesting = u32::try_from(nesting).map_err(|_| SlideMediaDataError::InvalidSource)?;
        let output = self.remaining(
            self.output_bytes,
            self.max_output_bytes,
            SlideMediaDataLimitKind::OutputBytes,
        )?;
        Ok(
            litchi_iwa_protos::package_metadata_media_codec::DecodeOptions::new(
                source.len().max(1),
                fields,
                work,
                entries,
                entries,
                owners,
                20,
                4096,
                nesting,
            )
            .with_max_output_bytes(output),
        )
    }

    /// Build a strict nested data-reference profile from remaining wire work.
    pub(super) fn keynote_options(
        &self,
        source: &[u8],
    ) -> Result<keynote_media_codec::DecodeOptions, SlideMediaDataError> {
        let fields = self.remaining(
            self.wire_fields,
            self.max_wire_fields,
            SlideMediaDataLimitKind::WireFields,
        )?;
        let work = self.remaining(
            self.wire_work,
            self.max_wire_work,
            SlideMediaDataLimitKind::WireWork,
        )?;
        let nesting = self.remaining(
            self.wire_nesting,
            self.max_wire_nesting,
            SlideMediaDataLimitKind::WireNesting,
        )?;
        let nesting = u32::try_from(nesting).map_err(|_| SlideMediaDataError::InvalidSource)?;
        Ok(keynote_media_codec::DecodeOptions::new(
            source.len().max(1),
            fields,
            work,
            nesting,
        ))
    }

    /// Charge the exact allocation and wire requirements planned by the
    /// strict metadata CAS codec before its `execute` call.
    pub(super) fn metadata_requirements(
        &mut self,
        requirements: RewriteExecutionRequirements,
    ) -> Result<(), SlideMediaDataError> {
        self.output_bytes(requirements.output_bytes())?;
        self.wire_fields(requirements.fields())?;
        self.wire_work(requirements.work_bytes())?;
        self.entries(
            requirements
                .components()
                .checked_add(requirements.data_records())
                .ok_or(SlideMediaDataError::InvalidSource)?,
        )?;
        self.references(requirements.owners())?;
        self.allocations_count(requirements.allocations())?;
        self.media_bytes(requirements.retained_bytes())?;
        self.media_bytes(requirements.scratch_bytes())
    }

    /// Charge the exact ZIP output, offset workspace, retained bytes, and
    /// allocation count reported by a source-authoritative reassembly plan.
    pub(super) fn reassembly_requirements(
        &mut self,
        requirements: ReassemblyExecutionRequirements,
    ) -> Result<(), SlideMediaDataError> {
        self.output_bytes(requirements.output_bytes())?;
        self.entries(requirements.offset_count())?;
        self.media_bytes(requirements.scratch_bytes())?;
        self.media_bytes(requirements.retained_bytes())?;
        self.allocations_count(requirements.allocations())
    }

    /// Snapshot the bounded work retained by the lazy selector/metadata
    /// passes.  The focused media-properties owner uses this to debit its
    /// `GeometryBudget` without reimplementing the PackageMetadata closure.
    pub(in crate::package) fn usage(&self) -> MediaBudgetUsage {
        MediaBudgetUsage {
            input_bytes: self.input_bytes,
            fields: self.wire_fields,
            nesting: self.wire_nesting,
            work: self.wire_work,
            references: self.references,
            allocations: self.allocations,
            // `media_bytes` covers the bounded string/vector/archive staging
            // charged by the lazy reader.  The sibling adapter classifies
            // this as scratch because its returned assets borrow the source.
            media_bytes: self.media_bytes,
        }
    }

    /// Charge fields from a bounded wire view.
    pub(super) fn wire_fields(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        Self::charge(
            &mut self.wire_fields,
            self.max_wire_fields,
            amount,
            SlideMediaDataLimitKind::WireFields,
        )
    }

    /// Charge maximum nesting observed by a strict codec.
    pub(super) fn wire_nesting(&mut self, amount: u32) -> Result<(), SlideMediaDataError> {
        let amount = usize::try_from(amount).map_err(|_| SlideMediaDataError::InvalidSource)?;
        self.wire_nesting = self.wire_nesting.max(amount);
        if self.wire_nesting > self.max_wire_nesting {
            return Err(limit(
                SlideMediaDataLimitKind::WireNesting,
                self.wire_nesting,
                self.max_wire_nesting,
            ));
        }
        Ok(())
    }

    /// Charge aggregate strict-plus-Buffa wire work.
    pub(super) fn wire_work(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        Self::charge(
            &mut self.wire_work,
            self.max_wire_work,
            amount,
            SlideMediaDataLimitKind::WireWork,
        )
    }

    /// Charge an exact count of planned allocations without double-counting
    /// the retained and scratch byte reservations.
    fn allocations_count(&mut self, amount: usize) -> Result<(), SlideMediaDataError> {
        let observed = self.allocations.checked_add(amount).ok_or_else(|| {
            limit(
                SlideMediaDataLimitKind::Allocations,
                usize::MAX,
                self.max_allocations,
            )
        })?;
        if observed > self.max_allocations {
            return Err(limit(
                SlideMediaDataLimitKind::Allocations,
                observed,
                self.max_allocations,
            ));
        }
        self.allocations = observed;
        Ok(())
    }

    fn charge(
        current: &mut usize,
        maximum: usize,
        amount: usize,
        kind: SlideMediaDataLimitKind,
    ) -> Result<(), SlideMediaDataError> {
        let observed = current
            .checked_add(amount)
            .ok_or_else(|| limit(kind, usize::MAX, maximum))?;
        if observed > maximum {
            return Err(limit(kind, observed, maximum));
        }
        *current = observed;
        Ok(())
    }

    fn remaining(
        &self,
        current: usize,
        maximum: usize,
        kind: SlideMediaDataLimitKind,
    ) -> Result<usize, SlideMediaDataError> {
        maximum
            .checked_sub(current)
            .filter(|remaining| *remaining > 0)
            .ok_or_else(|| limit(kind, current.saturating_add(1), maximum))
    }
}

fn as_usize(value: u64) -> Result<usize, SlideMediaDataError> {
    usize::try_from(value).map_err(|_| SlideMediaDataError::InvalidSource)
}

fn aggregate(value: usize) -> Result<usize, SlideMediaDataError> {
    value
        .checked_mul(4)
        .ok_or(SlideMediaDataError::InvalidSource)
}

fn limit(kind: SlideMediaDataLimitKind, observed: usize, maximum: usize) -> SlideMediaDataError {
    SlideMediaDataError::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    }
}
