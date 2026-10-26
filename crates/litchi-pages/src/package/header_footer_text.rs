//! Source-preserving transactions for rooted Pages header and footer text.
//!
//! This owner intentionally handles the existing-storage case only.  It
//! resolves a logical section/template/role/slot through the rooted Pages
//! graph, proves that the selected storage is a writable text object owned by
//! that graph, then rewrites that one physical storage.  A storage may be
//! intentionally shared by several logical slots; all aliases are read back
//! before publication.

use std::{fmt, mem::size_of, ops::Range, sync::Arc};

use litchi_iwa_archive::package::EntryEdit;
use litchi_iwa_core::{RawMessage, SnappyStream};
use litchi_iwa_protos::{package_metadata_codec, pages_header_footer_codec, pages_section_codec};
use thiserror::Error as ThisError;

use super::{
    Package, PackageError, effective_text_limit, is_body_text_message_type,
    root_references_with_limits, storage_rewrite_limits,
};
use crate::header_footer::{HeaderFooter, HeaderFooterSelector, Kind, Template};
use crate::selector::SectionSelector;
use litchi_core::Position;

const SECTION_MESSAGE_TYPE: u32 = 10_011;
const SECTION_TEMPLATE_MESSAGE_TYPE: u32 = 10_143;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const PREVIEW_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];
const MAX_VARINT_BYTES: usize = 10;

/// One checked budget for a complete header/footer transaction.
///
/// Discovery and codec reports are deliberately charged into this coordinator
/// rather than each phase receiving the package's full limits independently.
#[derive(Debug, Clone, Copy)]
struct TransactionBudget {
    maximum: [u64; 17],
    observed: [u64; 17],
}

impl TransactionBudget {
    fn new(source: &Package) -> Result<Self, Error> {
        let limits = source.state.source.limits();
        let archive = limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let physical = limits.max_input_bytes();
        let total = limits.max_total_bytes();
        let stream =
            u64::try_from(limits.max_iwa_stream_bytes()).map_err(|_| Error::InvalidSource)?;
        let work = total
            .checked_mul(8)
            .and_then(|value| value.checked_add(stream))
            .ok_or(Error::InvalidSource)?;
        let maximum = [
            physical,
            physical,
            u64::try_from(limits.max_entries()).map_err(|_| Error::InvalidSource)?,
            limits.max_entry_bytes(),
            total,
            stream,
            stream,
            stream,
            u64::try_from(archive.max_header_fields()).map_err(|_| Error::InvalidSource)?,
            u64::try_from(archive.max_header_nesting()).map_err(|_| Error::InvalidSource)?,
            work,
            u64::try_from(archive.max_metadata_items()).map_err(|_| Error::InvalidSource)?,
            u64::try_from(limits.max_entries()).map_err(|_| Error::InvalidSource)?,
            u64::try_from(limits.max_entries()).map_err(|_| Error::InvalidSource)?,
            total,
            total,
            u64::try_from(limits.max_entries()).map_err(|_| Error::InvalidSource)?,
        ];
        let mut budget = Self {
            maximum,
            observed: [0; 17],
        };
        budget.charge(
            LimitKind::InputBytes,
            source.state.source.source_bytes().len(),
        )?;
        Ok(budget)
    }

    fn index(kind: LimitKind) -> usize {
        kind as usize
    }

    fn charge(&mut self, kind: LimitKind, amount: usize) -> Result<(), Error> {
        let index = Self::index(kind);
        let amount = u64::try_from(amount).map_err(|_| Error::Limit {
            kind,
            observed: u64::MAX,
            maximum: self.maximum[index],
        })?;
        let observed = self.observed[index]
            .checked_add(amount)
            .ok_or(Error::Limit {
                kind,
                observed: u64::MAX,
                maximum: self.maximum[index],
            })?;
        if observed > self.maximum[index] {
            return Err(Error::Limit {
                kind,
                observed,
                maximum: self.maximum[index],
            });
        }
        self.observed[index] = observed;
        Ok(())
    }

    fn remaining(&self, kind: LimitKind) -> usize {
        let index = Self::index(kind);
        self.maximum[index]
            .saturating_sub(self.observed[index])
            .min(usize::MAX as u64) as usize
    }

    fn maximum(&self, kind: LimitKind) -> usize {
        self.maximum[Self::index(kind)].min(usize::MAX as u64) as usize
    }

    fn observe(&mut self, kind: LimitKind, amount: usize) -> Result<(), Error> {
        let index = Self::index(kind);
        let amount = u64::try_from(amount).map_err(|_| Error::Limit {
            kind,
            observed: u64::MAX,
            maximum: self.maximum[index],
        })?;
        if amount > self.maximum[index] {
            return Err(Error::Limit {
                kind,
                observed: amount,
                maximum: self.maximum[index],
            });
        }
        self.observed[index] = self.observed[index].max(amount);
        Ok(())
    }

    fn preflight_physical(
        &mut self,
        source: &Package,
        selected_member: &str,
        metadata_member: &str,
        before: &str,
        after: &str,
    ) -> Result<(), Error> {
        let source_bytes = source.state.source.source_bytes().len();
        let selected_bytes = source
            .state
            .source
            .package()
            .iter()
            .find(|entry| entry.name() == selected_member)
            .map_or(0, |entry| entry.data().len());
        let metadata_bytes = source
            .state
            .source
            .package()
            .iter()
            .find(|entry| entry.name() == metadata_member)
            .map_or(0, |entry| entry.data().len());
        let text_growth = after.len().saturating_sub(before.len());
        let output_bound = source_bytes
            .checked_add(text_growth)
            .and_then(|value| value.checked_add(selected_bytes))
            .and_then(|value| value.checked_add(metadata_bytes))
            .ok_or(Error::InvalidSource)?;
        self.observe(LimitKind::OutputBytes, output_bound)?;
        self.observe(LimitKind::TotalBytes, output_bound)?;
        self.observe(LimitKind::EntryBytes, selected_bytes)?;
        self.observe(LimitKind::EntryBytes, metadata_bytes)?;
        self.charge(LimitKind::TextBytes, before.len())?;
        self.charge(LimitKind::TextBytes, after.len())?;
        self.charge(
            LimitKind::TextUnits,
            before
                .encode_utf16()
                .count()
                .saturating_add(after.encode_utf16().count()),
        )?;
        self.charge(LimitKind::WireWork, source_bytes.saturating_mul(4))?;
        self.observe(
            LimitKind::ScratchBytes,
            selected_bytes
                .saturating_add(metadata_bytes)
                .saturating_add(output_bound),
        )?;
        self.observe(LimitKind::RetainedBytes, output_bound)?;
        self.charge(LimitKind::Allocations, 6)
    }

    fn precharge_candidate(&mut self, output_bytes: usize) -> Result<(), Error> {
        self.charge(LimitKind::WireWork, output_bytes.saturating_mul(2))?;
        self.observe(LimitKind::RetainedBytes, output_bytes)?;
        self.observe(LimitKind::ScratchBytes, output_bytes)?;
        self.charge(LimitKind::Allocations, 2)
    }

    fn charge_reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), Error> {
        self.observe(LimitKind::OutputBytes, requirements.output_bytes())?;
        self.observe(LimitKind::TotalBytes, requirements.output_bytes())?;
        self.charge(LimitKind::Entries, requirements.offset_count())?;
        self.observe(LimitKind::ScratchBytes, requirements.scratch_bytes())?;
        self.observe(LimitKind::RetainedBytes, requirements.retained_bytes())?;
        self.charge(LimitKind::Allocations, requirements.allocations())
    }

    fn charge_metadata_report(
        &mut self,
        report: package_metadata_codec::RewriteReport,
    ) -> Result<(), Error> {
        self.charge(LimitKind::WireBytes, report.input_bytes())?;
        self.charge(LimitKind::WireFields, report.fields())?;
        self.charge(LimitKind::WireWork, report.work_bytes())?;
        self.observe(LimitKind::WireNesting, report.max_depth() as usize)?;
        self.charge(LimitKind::Components, report.components_scanned())?;
        self.charge(LimitKind::References, report.references_scanned())?;
        self.charge(LimitKind::Allocations, report.allocations())?;
        self.charge(LimitKind::RetainedBytes, report.retained_bytes())?;
        self.charge(LimitKind::ScratchBytes, report.scratch_bytes())
    }

    fn charge_metadata_requirements(
        &mut self,
        requirements: package_metadata_codec::RewriteExecutionRequirements,
    ) -> Result<(), Error> {
        self.charge(LimitKind::OutputBytes, requirements.output_bytes())?;
        self.charge(LimitKind::WireFields, requirements.fields())?;
        self.charge(LimitKind::WireWork, requirements.work_bytes())?;
        self.charge(LimitKind::Components, requirements.components())?;
        self.charge(LimitKind::References, requirements.references())?;
        self.charge(LimitKind::Allocations, requirements.allocations())?;
        self.charge(LimitKind::RetainedBytes, requirements.retained_bytes())?;
        self.charge(LimitKind::ScratchBytes, requirements.scratch_bytes())
    }

    fn charge_section_report(
        &mut self,
        payload_len: usize,
        report: pages_section_codec::DecodeReport,
    ) -> Result<(), Error> {
        self.charge(LimitKind::WireBytes, payload_len)?;
        self.charge(LimitKind::WireFields, report.fields())?;
        self.charge(LimitKind::WireWork, report.work_bytes())?;
        self.observe(LimitKind::WireNesting, report.max_depth() as usize)?;
        self.charge(LimitKind::TextBytes, report.name_bytes())
    }

    fn charge_template_report(
        &mut self,
        report: pages_header_footer_codec::DecodeReport,
    ) -> Result<(), Error> {
        self.charge(LimitKind::WireBytes, report.input_bytes())?;
        self.charge(LimitKind::WireFields, report.fields())?;
        self.charge(LimitKind::WireWork, report.work_bytes())?;
        self.observe(LimitKind::WireNesting, report.max_depth() as usize)?;
        self.charge(LimitKind::References, report.references())?;
        self.charge(LimitKind::Allocations, report.allocations())?;
        self.charge(LimitKind::RetainedBytes, report.retained_bytes())?;
        self.charge(LimitKind::ScratchBytes, report.scratch_bytes())
    }

    fn charge_text_validation(
        &mut self,
        payload_len: usize,
        validation: litchi_iwa_text_wire::StorageValidation,
    ) -> Result<(), Error> {
        self.charge(LimitKind::WireBytes, payload_len)?;
        self.charge(LimitKind::WireFields, validation.fields())?;
        self.charge(LimitKind::WireWork, validation.validation_work())?;
        self.charge(LimitKind::References, validation.reference_occurrences())?;
        self.charge(LimitKind::Entries, validation.table_entries())?;
        self.charge(LimitKind::TextBytes, validation.utf8_len())?;
        self.charge(LimitKind::TextUnits, validation.utf16_len())
    }

    fn charge_text_requirements(
        &mut self,
        requirements: litchi_iwa_text_wire::StorageRewriteExecutionRequirements,
    ) -> Result<(), Error> {
        self.charge(LimitKind::OutputBytes, requirements.output_bytes())?;
        self.charge(LimitKind::WireWork, requirements.work())?;
        self.charge(LimitKind::References, requirements.reference_occurrences())?;
        self.charge(LimitKind::Allocations, requirements.allocations())?;
        self.charge(LimitKind::RetainedBytes, requirements.retained_bytes())?;
        self.charge(LimitKind::ScratchBytes, requirements.peak_scratch_bytes())
    }

    fn charge_discovery_allocation(&mut self, amount: usize) -> Result<(), Error> {
        self.charge(LimitKind::Allocations, 1)?;
        self.charge(
            LimitKind::ScratchBytes,
            amount.saturating_mul(size_of::<SlotRecord>()),
        )
    }

    fn charge_text_allocation(
        &mut self,
        validation: litchi_iwa_text_wire::StorageValidation,
    ) -> Result<(), Error> {
        let bytes = validation.utf8_len().saturating_add(
            validation
                .fragments()
                .saturating_mul(size_of::<usize>() * 4),
        );
        self.charge(LimitKind::Allocations, 1)?;
        self.charge(LimitKind::RetainedBytes, bytes)?;
        self.charge(LimitKind::ScratchBytes, bytes)
    }
}

/// Finite resource categories enforced by one header/footer transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    TextBytes,
    TextUnits,
    WireBytes,
    WireFields,
    WireNesting,
    WireWork,
    References,
    Components,
    RegistryChanges,
    RetainedBytes,
    ScratchBytes,
    Allocations,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::TextBytes => "text bytes",
            Self::TextUnits => "text UTF-16 units",
            Self::WireBytes => "wire bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::References => "references",
            Self::Components => "metadata components",
            Self::RegistryChanges => "metadata registry changes",
            Self::RetainedBytes => "retained bytes",
            Self::ScratchBytes => "scratch bytes",
            Self::Allocations => "allocations",
        })
    }
}

/// Error returned by header/footer reads and transactions.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    #[error("the rooted Pages header/footer graph is invalid")]
    InvalidSource,
    #[error("Pages header/footer {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        kind: LimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Pages header/footer transaction")]
    Allocation { amount: usize },
    #[error("the edited Pages header/footer failed semantic verification")]
    Verification,
    #[error("the Pages header/footer patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    #[error("the Pages header/footer selector did not match a logical slot")]
    NotFound,
    #[error("the Pages header/footer selector is ambiguous")]
    AmbiguousSelector,
    #[error("this Pages source does not support exact header/footer text editing")]
    UnsupportedSource,
    #[error("the selected Pages header/footer graph has an unsupported dependency")]
    UnsupportedDependency,
    #[error("Pages header/footer text contains an invalid UTF-16 range")]
    PositionOutOfBounds,
    #[error("Pages header/footer text exceeds its semantic byte budget")]
    TextTooLarge,
    #[error("Pages header/footer text contains a native structural marker")]
    StructuralMarker,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NotFound { .. } => RefusalKind::NotFound,
            Self::AmbiguousSelector { .. }
            | Self::PositionOutOfBounds { .. }
            | Self::TextTooLarge { .. } => RefusalKind::InvalidValue,
            Self::UnsupportedSource { .. }
            | Self::UnsupportedDependency { .. }
            | Self::StructuralMarker { .. } => RefusalKind::UnsupportedContent,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// Mutable selector-first header/footer text edit.
pub struct Edit<'a> {
    source: &'a Package,
    selector: DeferredSelector,
    resolved: Option<ResolvedSelector>,
    before: Option<String>,
    after: String,
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Edit")
            .field("selector", &self.selector)
            .field("before_bytes", &self.before.as_ref().map_or(0, String::len))
            .field("after_bytes", &self.after.len())
            .finish()
    }
}

impl Edit<'_> {
    fn resolve(&mut self) -> Result<(), Error> {
        if self.resolved.is_none() {
            let mut budget = TransactionBudget::new(self.source)?;
            budget.charge(LimitKind::Allocations, self.selector.retained_allocations())?;
            let selector = self.selector.public();
            let section_index =
                resolve_section_index(self.source, selector.section(), &mut budget)?;
            let canonical = HeaderFooterSelector::index(
                section_index,
                selector.template(),
                selector.kind(),
                selector.slot_index(),
            );
            let records = discover_slots_with_budget(self.source, &mut budget)?;
            let selected = select_record(&records, canonical, &mut budget)?;
            if self.before.is_none() {
                budget.charge(LimitKind::Allocations, 2)?;
                budget.charge(
                    LimitKind::RetainedBytes,
                    selected.text.len().saturating_mul(2),
                )?;
                self.before = Some(selected.text.clone());
                self.after = selected.text.clone();
            }
            self.resolved = Some(selected.resolved);
        }
        Ok(())
    }

    /// Borrow the current text.
    pub fn text(&mut self) -> Result<&str, Error> {
        self.resolve()?;
        Ok(&self.after)
    }

    /// Borrow the source text captured when the edit was created.
    pub fn before(&mut self) -> Result<&str, Error> {
        self.resolve()?;
        self.before.as_deref().ok_or(Error::InvalidSource)
    }

    /// Replace a checked UTF-16 range in the staged text.
    pub fn replace(&mut self, range: Range<usize>, replacement: &str) -> Result<&mut Self, Error> {
        self.resolve()?;
        validate_authored(replacement)?;
        let start = utf16_boundary(&self.after, range.start)?;
        let end = utf16_boundary(&self.after, range.end)?;
        if start > end {
            return Err(Error::Refused(Refusal::PositionOutOfBounds));
        }
        self.after.replace_range(start..end, replacement);
        validate_authored(&self.after)?;
        Ok(self)
    }

    /// Replace the complete staged value.
    pub fn set(&mut self, value: &str) -> Result<&mut Self, Error> {
        self.resolve()?;
        validate_authored(value)?;
        self.after.clear();
        self.after
            .try_reserve_exact(value.len())
            .map_err(|_| Error::Allocation {
                amount: value.len(),
            })?;
        self.after.push_str(value);
        Ok(self)
    }

    /// Clear the selected logical slot's physical storage.
    pub fn clear(&mut self) -> Result<&mut Self, Error> {
        self.set("")
    }

    /// Validate and publish the staged replacement.
    pub fn commit(mut self) -> Result<Commit, Error> {
        self.resolve()?;
        commit_edit(self)
    }
}

/// Exact-source reversible header/footer patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    source_fingerprint: u64,
    target_fingerprint: u64,
    selector: ResolvedSelector,
    before: String,
    after: String,
    aliases: Arc<[ResolvedSelector]>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        f.debug_struct("Patch")
            .field("selector", &self.selector.public)
            .field("before_bytes", &self.before.len())
            .field("after_bytes", &self.after.len())
            .finish()
    }
}

impl Patch {
    #[must_use]
    pub fn before(&self) -> &str {
        &self.before
    }

    #[must_use]
    pub fn after(&self) -> &str {
        &self.after
    }

    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.source_fingerprint == self.target_fingerprint
            && self.source.as_ref() == self.target.as_ref()
            && self.before == self.after
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            selector: self.selector,
            before: self.after.clone(),
            after: self.before.clone(),
            aliases: Arc::clone(&self.aliases),
        }
    }
}

/// Diagnostics for one header/footer text publication.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    affected_slots: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            affected_slots: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn modified(touched_components: usize, affected_slots: usize, deleted: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            affected_slots,
            deleted_previews: deleted,
            full_reparse_performed: true,
        }
    }

    #[must_use]
    pub const fn changed_bytes(self) -> bool {
        self.changed
    }

    /// Whether the transaction changed the package bytes.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    #[must_use]
    pub const fn affected_slots(self) -> usize {
        self.affected_slots
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

/// Verified result of one header/footer text transaction.
#[must_use]
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

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ResolvedSelector {
    public: HeaderFooterSelector<'static>,
    section_id: u64,
    template_id: u64,
    storage_id: u64,
    component_index: usize,
}

#[derive(Clone, PartialEq, Eq)]
struct DeferredSelector {
    section_position: Option<Position>,
    section_name: Option<Box<str>>,
    template: Template,
    kind: Kind,
    slot: Position,
}

impl fmt::Debug for DeferredSelector {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("DeferredSelector")
            .field("section_by_name", &self.section_name.is_some())
            .field("section_position", &self.section_position)
            .field("template", &self.template)
            .field("kind", &self.kind)
            .field("slot", &self.slot)
            .finish()
    }
}

impl DeferredSelector {
    fn try_from_public(
        source: &Package,
        selector: HeaderFooterSelector<'_>,
    ) -> Result<Self, Error> {
        let (section_position, section_name) = match selector.section() {
            SectionSelector::Position(position) => (Some(position), None),
            SectionSelector::Name(name) => {
                let mut budget = TransactionBudget::new(source)?;
                (None, Some(own_section_name(&mut budget, name)?))
            },
        };
        Ok(Self {
            section_position,
            section_name,
            template: selector.template(),
            kind: selector.kind(),
            slot: selector.slot(),
        })
    }

    fn public(&self) -> HeaderFooterSelector<'_> {
        let section = self.section_position.map_or_else(
            || SectionSelector::Name(self.section_name.as_deref().unwrap_or("")),
            SectionSelector::Position,
        );
        HeaderFooterSelector::new(section, self.template, self.kind, self.slot)
    }

    fn retained_allocations(&self) -> usize {
        self.section_name.is_some() as usize
    }
}

fn own_section_name(budget: &mut TransactionBudget, name: &str) -> Result<Box<str>, Error> {
    budget.charge(LimitKind::TextBytes, name.len())?;
    budget.charge(LimitKind::RetainedBytes, name.len())?;
    budget.charge(LimitKind::Allocations, 1)?;
    let mut owned = String::new();
    owned
        .try_reserve_exact(name.len())
        .map_err(|_| Error::Allocation { amount: name.len() })?;
    owned.push_str(name);
    Ok(owned.into_boxed_str())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn budget_with_allocations(maximum: u64) -> TransactionBudget {
        let mut maximums = [u64::MAX; 17];
        maximums[LimitKind::Allocations as usize] = maximum;
        TransactionBudget {
            maximum: maximums,
            observed: [0; 17],
        }
    }

    #[test]
    fn named_selector_allocation_admission_is_typed_at_exact_and_one_under() {
        let mut exact = budget_with_allocations(1);
        let owned = own_section_name(&mut exact, "section").unwrap();
        assert_eq!(&*owned, "section");
        assert_eq!(exact.observed[LimitKind::Allocations as usize], 1,);

        let mut one_under = budget_with_allocations(0);
        assert_eq!(
            own_section_name(&mut one_under, "section"),
            Err(Error::Limit {
                kind: LimitKind::Allocations,
                observed: 1,
                maximum: 0,
            })
        );
    }
}

#[derive(Clone, Debug)]
struct SlotRecord {
    resolved: ResolvedSelector,
    text: String,
}

struct MetadataMatch<'a> {
    target: &'a str,
    found: Option<(u64, String)>,
    ambiguous: bool,
}

struct MetadataOwnership<'a> {
    source: &'a Package,
    records: &'a [SlotRecord],
    unsupported: bool,
}

impl MetadataOwnership<'_> {
    fn contains(&self, identifier: u64) -> bool {
        self.records
            .iter()
            .any(|record| record.resolved.storage_id == identifier)
    }
}

impl package_metadata_codec::PackageMetadataVisitor for MetadataOwnership<'_> {
    fn visit_object_uuid(
        &mut self,
        binding: package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        let Some(record) = self
            .records
            .iter()
            .find(|record| record.resolved.storage_id == binding.object_identifier())
        else {
            return Ok(());
        };
        let Some(component) = self
            .source
            .state
            .source
            .components()
            .get_index(record.resolved.component_index)
        else {
            self.unsupported = true;
            return Ok(());
        };
        // Versioned component UUID maps are historical snapshots.  They are
        // validated by the strict metadata codec but do not own current
        // objects or participate in current selector matching.
        if !binding.component().is_current() {
            return Ok(());
        }
        if normalized_metadata_locator(binding.component().effective_locator())
            != normalized_metadata_locator(component.name())
        {
            self.unsupported = true;
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if reference
            .object_identifier()
            .is_some_and(|identifier| self.contains(identifier))
            && reference.is_weak() != Some(true)
        {
            self.unsupported = true;
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if self.contains(owner.object_identifier()) {
            self.unsupported = true;
        }
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if self.contains(identifier) {
            self.unsupported = true;
        }
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if self.contains(object_identifier) {
            self.unsupported = true;
        }
        Ok(())
    }
}

impl package_metadata_codec::PackageMetadataVisitor for MetadataMatch<'_> {
    fn visit_component(
        &mut self,
        component: package_metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if component.is_current()
            && normalized_metadata_locator(component.effective_locator()) == self.target
        {
            if self.found.is_some() {
                self.ambiguous = true;
                return Ok(());
            }
            let locator = component.effective_locator();
            let mut owned = String::new();
            owned
                .try_reserve_exact(locator.len())
                .map_err(|_| package_metadata_codec::RewriteError::allocation(locator.len()))?;
            owned.push_str(locator);
            self.found = Some((component.identifier(), owned));
        }
        Ok(())
    }
}

impl Package {
    /// List every rooted header/footer slot and its text.
    pub fn header_footers(&self) -> Result<Vec<HeaderFooter>, PackageError> {
        let mut budget = TransactionBudget::new(self)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        let records = discover_slots_with_budget(self, &mut budget)
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        let mut values = Vec::new();
        budget
            .charge_discovery_allocation(records.len())
            .map_err(|error| PackageError::InvalidFormat(error.to_string()))?;
        values
            .try_reserve_exact(records.len())
            .map_err(|_| PackageError::Allocation {
                amount: records.len(),
            })?;
        for record in records {
            values.push(HeaderFooter::new(
                Position::new(
                    record
                        .resolved
                        .public
                        .section()
                        .as_position()
                        .map_or(0, |p| p.get()),
                ),
                self.sections()
                    .get(
                        record
                            .resolved
                            .public
                            .section()
                            .as_position()
                            .map_or(0, |p| p.get()),
                    )
                    .and_then(|section| section.name()),
                record.resolved.public.template(),
                record.resolved.public.kind(),
                record.resolved.public.slot(),
                record.text.into_boxed_str(),
            ));
        }
        Ok(values)
    }

    /// Stage an edit of one rooted logical header/footer slot.
    pub fn edit_header_footer_text(
        &self,
        selector: HeaderFooterSelector<'_>,
    ) -> Result<Edit<'_>, Error> {
        Ok(Edit {
            source: self,
            selector: DeferredSelector::try_from_public(self, selector)?,
            resolved: None,
            before: None,
            after: String::new(),
        })
    }

    /// Apply an exact-source-checked header/footer patch.
    pub fn apply_header_footer_text(&self, patch: &Patch) -> Result<Commit, Error> {
        if self.source_bytes() != patch.source.as_ref()
            || fingerprint(self.source_bytes()) != patch.source_fingerprint
        {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        let mut budget = TransactionBudget::new(self)?;
        let current = discover_slots_with_budget(self, &mut budget)?;
        budget.charge(LimitKind::WireWork, current.len())?;
        let selected = current
            .iter()
            .find(|record| record.resolved == patch.selector)
            .ok_or(Error::PatchConflict)?;
        if selected.text != patch.before {
            return Err(Error::PatchConflict);
        }
        if !self.state.source.source_is_exact()
            || fingerprint(patch.target.as_ref()) != patch.target_fingerprint
        {
            return Err(Error::PatchConflict);
        }
        // Patches carry the fully verified target artifact.  Applying a patch
        // must reopen that artifact, rather than recomputing a fresh forward
        // transition (which would advance save tokens again and would delete
        // previews a second time when applying an inverse).
        budget.observe(LimitKind::OutputBytes, patch.target.len())?;
        budget.precharge_candidate(patch.target.len())?;
        let candidate = Package::from_bytes_with_archive_limits(
            patch.target.as_ref(),
            self.state.source.limits(),
        )
        .map_err(map_package_error)?;
        if candidate.source_bytes() != patch.target.as_ref() {
            return Err(Error::Verification);
        }
        let candidate_records = discover_slots_with_budget(&candidate, &mut budget)?;
        for alias in patch.aliases.iter() {
            budget.charge(LimitKind::WireWork, candidate_records.len())?;
            let found = candidate_records
                .iter()
                .find(|record| record.resolved == *alias)
                .ok_or(Error::Verification)?;
            if found.text != patch.after {
                return Err(Error::Verification);
            }
        }
        let source_previews = source_preview_count(self);
        let target_previews = source_preview_count(&candidate);
        let mut touched_components = std::collections::HashSet::new();
        let selected_component =
            component_name_for_object(self, patch.selector.storage_id, &mut budget)?;
        touched_components.insert(selected_component.clone());
        touched_components.insert("Index/Metadata.iwa".to_owned());
        verify_locality(
            self,
            &candidate,
            &selected_component,
            "Index/Metadata.iwa",
            &PREVIEW_NAMES.map(str::to_owned),
            &mut budget,
        )?;
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::modified(
                touched_components.len(),
                patch.aliases.len(),
                source_previews.saturating_sub(target_previews),
            ),
        })
    }
}

fn source_preview_count(source: &Package) -> usize {
    source
        .state
        .source
        .package()
        .iter()
        .filter(|entry| PREVIEW_NAMES.contains(&entry.name()))
        .count()
}

fn commit_edit(edit: Edit<'_>) -> Result<Commit, Error> {
    let mut budget = TransactionBudget::new(edit.source)?;
    budget.charge(LimitKind::Allocations, edit.selector.retained_allocations())?;
    edit.source.validate().map_err(map_package_error)?;
    let records = discover_slots_with_budget(edit.source, &mut budget)?;
    budget.charge(LimitKind::WireWork, records.len())?;
    let selector = edit.resolved.ok_or(Error::InvalidSource)?;
    let before = edit.before.ok_or(Error::InvalidSource)?;
    let selected = records
        .iter()
        .find(|record| record.resolved == selector)
        .ok_or(Error::PatchConflict)?;
    if selected.text != before {
        return Err(Error::PatchConflict);
    }
    if before == edit.after {
        let bytes: Arc<[u8]> = edit.source.state.source.shared_source();
        let fingerprint = fingerprint(&bytes);
        let aliases = alias_selectors(&records, selector, &mut budget)?;
        return Ok(Commit {
            package: edit.source.clone(),
            patch: Patch {
                source: Arc::clone(&bytes),
                target: bytes,
                source_fingerprint: fingerprint,
                target_fingerprint: fingerprint,
                selector,
                before,
                after: edit.after,
                aliases: aliases.into(),
            },
            diagnostics: Diagnostics::unchanged(),
        });
    }
    let aliases = alias_selectors(&records, selector, &mut budget)?;
    publish_rewrite(
        edit.source,
        selector,
        &before,
        &edit.after,
        &aliases,
        &mut budget,
    )
}

fn publish_rewrite(
    source: &Package,
    selector: ResolvedSelector,
    before: &str,
    after: &str,
    aliases: &[ResolvedSelector],
    budget: &mut TransactionBudget,
) -> Result<Commit, Error> {
    validate_authored(after)?;
    let selected_component = component_name_for_object(source, selector.storage_id, budget)?;
    let metadata_member_name = "Index/Metadata.iwa";
    budget.preflight_physical(
        source,
        &selected_component,
        metadata_member_name,
        before,
        after,
    )?;
    let metadata = rewrite_metadata_tokens(source, &selected_component, budget)?;
    let rewritten_storage =
        rewrite_storage_component(source, selector.storage_id, before, after, budget)?;
    let package_output_bound = precharge_candidate_catalog(
        source,
        [
            (
                selected_component.as_str(),
                rewritten_storage.compressed.len(),
                rewritten_storage.archive_bound,
            ),
            (
                metadata.member.as_str(),
                metadata.compressed.len(),
                metadata.archive_bound,
            ),
        ],
        budget,
    )?;
    let mut edits = Vec::new();
    budget.charge_discovery_allocation(2)?;
    edits
        .try_reserve_exact(2)
        .map_err(|_| Error::Allocation { amount: 2 })?;
    edits.push(EntryEdit::new(
        &selected_component,
        &rewritten_storage.compressed,
    ));
    edits.push(EntryEdit::new(&metadata.member, &metadata.compressed));
    let mut previews = Vec::new();
    for entry in source.state.source.package().iter() {
        if PREVIEW_NAMES.contains(&entry.name()) {
            budget.charge_discovery_allocation(1)?;
            previews.try_reserve(1).map_err(|_| Error::Allocation {
                amount: previews.len().saturating_add(1),
            })?;
            previews.push(entry.name().to_owned());
        }
    }
    let mut preview_names = Vec::new();
    budget.charge_discovery_allocation(previews.len())?;
    preview_names
        .try_reserve_exact(previews.len())
        .map_err(|_| Error::Allocation {
            amount: previews.len(),
        })?;
    preview_names.extend(previews.iter().map(String::as_str));
    budget.observe(
        LimitKind::ScratchBytes,
        source
            .state
            .source
            .source_bytes()
            .len()
            .saturating_add(edits.len().saturating_mul(size_of::<EntryEdit<'_>>())),
    )?;
    budget.charge(LimitKind::Allocations, 2)?;
    let prepared = source
        .state
        .source
        .package()
        .prepare_reassembly_with_deletions(&edits, &preview_names, source.state.source.limits())
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    if requirements.output_bytes() > package_output_bound {
        return Err(Error::Verification);
    }
    budget.charge_reassembly(requirements)?;
    budget.precharge_candidate(requirements.output_bytes())?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let target_bytes: Arc<[u8]> = output.clone().into();
    let candidate = Package::from_bytes_with_archive_limits(&output, source.state.source.limits())
        .map_err(map_package_error)?;
    verify_member_bytes(
        &candidate,
        &selected_component,
        &rewritten_storage.compressed,
        budget,
    )?;
    verify_member_bytes(&candidate, &metadata.member, &metadata.compressed, budget)?;
    let candidate_records = discover_slots_with_budget(&candidate, budget)?;
    for alias in aliases {
        budget.charge(LimitKind::WireWork, candidate_records.len())?;
        let found = candidate_records
            .iter()
            .find(|record| record.resolved == *alias)
            .ok_or(Error::Verification)?;
        if found.text != after {
            return Err(Error::Verification);
        }
    }
    verify_locality(
        source,
        &candidate,
        &selected_component,
        &metadata.member,
        &previews,
        budget,
    )?;
    let source_bytes: Arc<[u8]> = source.state.source.shared_source();
    let source_fingerprint = fingerprint(&source_bytes);
    let target_fingerprint = fingerprint(&target_bytes);
    let mut changed_components = std::collections::HashSet::new();
    changed_components.insert(selected_component.as_str());
    changed_components.insert(metadata.member.as_str());
    Ok(Commit {
        package: candidate,
        patch: Patch {
            source: source_bytes,
            target: target_bytes,
            source_fingerprint,
            target_fingerprint,
            selector,
            before: before.to_owned(),
            after: after.to_owned(),
            aliases: aliases.to_vec().into(),
        },
        diagnostics: Diagnostics::modified(changed_components.len(), aliases.len(), previews.len()),
    })
}

fn discover_slots_with_budget(
    source: &Package,
    budget: &mut TransactionBudget,
) -> Result<Vec<SlotRecord>, Error> {
    discover_slots(source, budget, false)
}

/// Discover every rooted slot; `allow_markers` admits slot text holding
/// attachment markers (for the families that own those attachments).
fn discover_slots(
    source: &Package,
    budget: &mut TransactionBudget,
    allow_markers: bool,
) -> Result<Vec<SlotRecord>, Error> {
    charge_catalog_census(source, budget)?;
    let section_ids = rooted_section_ids(source, budget)?;
    let mut records = Vec::new();
    for (section_index, section_id) in section_ids.into_iter().enumerate() {
        let section_object = unique_object(source, section_id, budget)?;
        let section_payload = unique_message(section_object, SECTION_MESSAGE_TYPE)?;
        let archive_limits = source
            .state
            .source
            .limits()
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let work =
            usize::try_from(source.state.source.limits().max_total_bytes()).unwrap_or(usize::MAX);
        residual_limit(budget, LimitKind::WireBytes, section_payload.len())?;
        let section_fields = residual_limit(budget, LimitKind::WireFields, 1)?;
        let section_work = residual_limit(budget, LimitKind::WireWork, 1)?;
        let (section, section_report) = pages_section_codec::decode_section_settings_with_report(
            section_payload,
            pages_section_codec::DecodeOptions::new(
                section_payload.len(),
                u32::try_from(
                    archive_limits
                        .max_header_nesting()
                        .min(budget.maximum(LimitKind::WireNesting)),
                )
                .unwrap_or(u32::MAX),
            )
            .with_max_fields(
                section_fields
                    .min(archive_limits.max_header_fields())
                    .max(1),
            )
            .with_max_work_bytes(section_work.min(work).max(1))
            .with_max_name_bytes(effective_text_limit(source.state.source.limits())),
        )
        .map_err(map_section_error)?;
        budget.charge_section_report(section_payload.len(), section_report)?;
        for (template, section_field, template_ref) in [
            (Template::First, 23, section.first_section_template_page()),
            (Template::Even, 24, section.even_section_template_page()),
            (Template::Odd, 25, section.odd_section_template_page()),
        ] {
            let Some(template_ref) = template_ref else {
                continue;
            };
            let template_id = template_ref.identifier().get();
            validate_declared_reference(section_object, template_id, section_field, budget)?;
            let template_object = unique_object(source, template_id, budget)?;
            let template_payload = unique_message(template_object, SECTION_TEMPLATE_MESSAGE_TYPE)?;
            let template_options = template_decode_options(source, template_payload, budget)?;
            budget.observe(
                LimitKind::ScratchBytes,
                template_payload.len().saturating_mul(2),
            )?;
            budget.charge(LimitKind::Allocations, 1)?;
            let (template_snapshot, template_report) =
                pages_header_footer_codec::decode_section_template_with_report(
                    template_payload,
                    template_options,
                )
                .map_err(map_template_error)?;
            budget.charge_template_report(template_report)?;
            for (kind, references) in [
                (Kind::Header, template_snapshot.header_identifiers()),
                (Kind::Footer, template_snapshot.footer_identifiers()),
            ] {
                let template_field = match kind {
                    Kind::Header => 1,
                    Kind::Footer => 2,
                };
                for (slot, storage_id) in references.enumerate() {
                    validate_declared_reference(
                        template_object,
                        storage_id,
                        template_field,
                        budget,
                    )?;
                    let storage_object = unique_object(source, storage_id, budget)?;
                    let payload =
                        super::unique_text_payload(&storage_object.messages, nonzero(storage_id)?)
                            .map_err(map_package_error)?;
                    let limits = storage_rewrite_limits(source.state.source.limits())
                        .map_err(|_| Error::InvalidSource)?;
                    let residual = residual_storage_limits(limits, budget)?;
                    let validation = litchi_iwa_text_wire::validate_storage(payload, residual)
                        .map_err(map_text_error)?;
                    if validation.storage_kind() != 1 {
                        return Err(Error::Refused(Refusal::UnsupportedDependency));
                    }
                    budget.charge_text_validation(payload.len(), validation)?;
                    budget.charge_text_allocation(validation)?;
                    let residual = residual_storage_limits(limits, budget)?;
                    let decoded = litchi_iwa_text_wire::decode_storage(payload, residual)
                        .map_err(map_text_error)?;
                    if decoded.validation() != validation {
                        return Err(Error::Verification);
                    }
                    budget.charge_text_validation(payload.len(), decoded.validation())?;
                    let text = decoded.into_storage().text().to_owned();
                    if !allow_markers {
                        validate_authored(&text)?;
                    }
                    let public = HeaderFooterSelector::new(
                        SectionSelector::index(section_index),
                        template,
                        kind,
                        Position::new(slot),
                    );
                    budget.charge_discovery_allocation(1)?;
                    records.try_reserve(1).map_err(|_| Error::Allocation {
                        amount: records.len().saturating_add(1),
                    })?;
                    records.push(SlotRecord {
                        resolved: ResolvedSelector {
                            public,
                            section_id,
                            template_id,
                            storage_id,
                            component_index: component_index(source, storage_id, budget)?,
                        },
                        text,
                    });
                }
            }
        }
        // A storage used by a header/footer must not simultaneously be owned
        // by body, drawable, or footnote graph records.  The alias set above
        // is the only permitted owner set for each physical storage.
    }
    reject_unattributed_storage_owners(source, &records, budget)?;
    reject_metadata_storage_owners(source, &records, budget)?;
    Ok(records)
}

fn rooted_section_ids(source: &Package, budget: &mut TransactionBudget) -> Result<Vec<u64>, Error> {
    let refs = root_references_with_limits(
        source.state.source.components(),
        source.state.source.limits(),
    )
    .map_err(map_package_error)?;
    let body = refs.body.ok_or(Error::InvalidSource)?;
    let body_object = unique_object(source, body.get(), budget)?;
    let body_payload_bytes = body_object
        .messages
        .iter()
        .try_fold(0usize, |total, message| {
            total.checked_add(message.data.len())
        })
        .ok_or(Error::InvalidSource)?;
    residual_limit(budget, LimitKind::WireBytes, body_payload_bytes.max(1))?;
    budget.charge(LimitKind::WireBytes, body_payload_bytes)?;
    budget.charge(
        LimitKind::WireWork,
        body_payload_bytes
            .saturating_mul(2)
            .saturating_add(body_object.messages.len()),
    )?;
    budget.observe(
        LimitKind::ScratchBytes,
        body_payload_bytes.saturating_add(super::MAX_SECTIONS.saturating_mul(size_of::<u64>())),
    )?;
    budget.charge(LimitKind::Allocations, 2)?;
    let (_storage, native) = super::decode_body_storage(
        &body_object.messages,
        body,
        super::MAX_SECTIONS,
        effective_text_limit(source.state.source.limits()),
        source.state.source.limits(),
    )
    .map_err(map_package_error)?;
    let native =
        super::native_section_references(native, refs.initial_section, super::MAX_SECTIONS)
            .map_err(map_package_error)?;
    let mut identifiers = Vec::new();
    for reference in native {
        let identifier = reference.identifier.get();
        budget.charge(LimitKind::WireWork, identifiers.len())?;
        if identifiers.contains(&identifier) {
            return Err(Error::InvalidSource);
        }
        budget.charge_discovery_allocation(1)?;
        identifiers.try_reserve(1).map_err(|_| Error::Allocation {
            amount: identifiers.len().saturating_add(1),
        })?;
        identifiers.push(identifier);
    }
    Ok(identifiers)
}

fn resolve_section_index(
    source: &Package,
    selector: SectionSelector<'_>,
    budget: &mut TransactionBudget,
) -> Result<usize, Error> {
    match selector {
        SectionSelector::Position(position) => {
            let index = position.get();
            if source.sections().get(index).is_none() {
                return Err(Error::Refused(Refusal::NotFound));
            }
            Ok(index)
        },
        SectionSelector::Name(name) => {
            let mut found = None;
            for (index, section) in source.sections().iter().enumerate() {
                budget.charge(
                    LimitKind::WireWork,
                    name.len()
                        .saturating_add(section.name().map_or(0, str::len))
                        .saturating_add(1),
                )?;
                if section.name() == Some(name) && found.replace(index).is_some() {
                    return Err(Error::Refused(Refusal::AmbiguousSelector));
                }
            }
            found.ok_or(Error::Refused(Refusal::NotFound))
        },
    }
}

fn select_record<'a>(
    records: &'a [SlotRecord],
    selector: HeaderFooterSelector<'_>,
    budget: &mut TransactionBudget,
) -> Result<&'a SlotRecord, Error> {
    let section_index = selector
        .section()
        .as_position()
        .ok_or(Error::InvalidSource)?
        .get();
    budget.charge(LimitKind::WireWork, records.len())?;
    records
        .iter()
        .find(|record| {
            record.resolved.public.section() == SectionSelector::index(section_index)
                && record.resolved.public.template() == selector.template()
                && record.resolved.public.kind() == selector.kind()
                && record.resolved.public.slot() == selector.slot()
        })
        .ok_or(Error::Refused(Refusal::NotFound))
}

fn alias_selectors(
    records: &[SlotRecord],
    selected: ResolvedSelector,
    budget: &mut TransactionBudget,
) -> Result<Vec<ResolvedSelector>, Error> {
    let mut aliases = Vec::new();
    for record in records {
        budget.charge(LimitKind::WireWork, 1)?;
        if record.resolved.storage_id != selected.storage_id {
            continue;
        }
        budget.charge_discovery_allocation(1)?;
        aliases.try_reserve(1).map_err(|_| Error::Allocation {
            amount: aliases.len().saturating_add(1),
        })?;
        aliases.push(record.resolved);
    }
    if aliases.is_empty() {
        return Err(Error::InvalidSource);
    }
    Ok(aliases)
}

fn unique_object<'source>(
    source: &'source Package,
    identifier: u64,
    budget: &mut TransactionBudget,
) -> Result<&'source litchi_iwa_core::ArchiveObject, Error> {
    let mut found = None;
    for component in source.state.source.components().iter() {
        budget.charge(LimitKind::Components, 1)?;
        budget.charge(LimitKind::WireWork, 1)?;
        if let Some(object) = component.archive().object(identifier) {
            if found.replace(object).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    found.ok_or(Error::InvalidSource)
}

fn component_index(
    source: &Package,
    identifier: u64,
    budget: &mut TransactionBudget,
) -> Result<usize, Error> {
    let mut found = None;
    for (index, component) in source.state.source.components().iter().enumerate() {
        budget.charge(LimitKind::Components, 1)?;
        budget.charge(LimitKind::WireWork, 1)?;
        if component.archive().object(identifier).is_some() {
            if found.replace(index).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    found.ok_or(Error::InvalidSource)
}

fn component_name_for_object(
    source: &Package,
    identifier: u64,
    budget: &mut TransactionBudget,
) -> Result<String, Error> {
    let mut found = None;
    for component in source.state.source.components().iter() {
        budget.charge(LimitKind::Components, 1)?;
        budget.charge(LimitKind::WireWork, 1)?;
        if component.archive().object(identifier).is_some() {
            if found.replace(component.name()).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    found.map(str::to_owned).ok_or(Error::InvalidSource)
}

fn unique_message(
    object: &litchi_iwa_core::ArchiveObject,
    message_type: u32,
) -> Result<&[u8], Error> {
    let mut found = None;
    for message in &object.messages {
        if message.type_ == message_type {
            if found.replace(message.data.as_slice()).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    found.ok_or(Error::InvalidSource)
}

fn validate_declared_reference(
    owner: &litchi_iwa_core::ArchiveObject,
    target: u64,
    expected_root_field: u32,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    if owner.archive_info.message_infos.len() != owner.messages.len() {
        return Err(Error::InvalidSource);
    }
    let mut matches = 0usize;
    for info in &owner.archive_info.message_infos {
        budget.charge(
            LimitKind::WireWork,
            info.object_references
                .len()
                .saturating_add(info.data_references.len())
                .saturating_add(info.field_infos.len()),
        )?;
        let aggregate = info
            .object_references
            .iter()
            .filter(|identifier| **identifier == target)
            .count();
        matches = matches.saturating_add(aggregate);
        if !info.field_infos.is_empty() {
            let mut field_matches = 0usize;
            for field in &info.field_infos {
                budget.charge(
                    LimitKind::WireWork,
                    field
                        .object_references
                        .len()
                        .saturating_add(field.data_references.len()),
                )?;
                let occurrences = field
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == target)
                    .count();
                if occurrences != 0 && field.path.as_slice() != [expected_root_field] {
                    return Err(Error::Refused(Refusal::UnsupportedDependency));
                }
                if field.data_references.contains(&target) {
                    return Err(Error::Refused(Refusal::UnsupportedDependency));
                }
                field_matches = field_matches.saturating_add(occurrences);
            }
            // Native Pages producers sometimes emit unrelated FieldInfo
            // records while leaving schema-declared references aggregate-only.
            // The strict section/template codec already proved the target at
            // `expected_root_field`; when any target-specific field metadata
            // is present it must nevertheless be complete and exact.
            if field_matches != 0 && field_matches != aggregate {
                return Err(Error::InvalidSource);
            }
        }
    }
    if matches != 1 {
        return Err(Error::InvalidSource);
    }
    Ok(())
}

fn reject_unattributed_storage_owners(
    source: &Package,
    records: &[SlotRecord],
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    for component in source.state.source.components().iter() {
        for object in &component.archive().objects {
            let owner = object.archive_info.identifier.unwrap_or(0);
            for message_info in &object.archive_info.message_infos {
                for target in &message_info.object_references {
                    budget.charge(LimitKind::WireWork, records.len())?;
                    let allowed = records
                        .iter()
                        .any(|record| record.resolved.storage_id == *target);
                    let expected = records.iter().any(|record| {
                        record.resolved.template_id == owner
                            && record.resolved.storage_id == *target
                    });
                    if allowed && !expected {
                        return Err(Error::Refused(Refusal::UnsupportedDependency));
                    }
                }
                for target in &message_info.data_references {
                    budget.charge(LimitKind::WireWork, records.len())?;
                    if records
                        .iter()
                        .any(|record| record.resolved.storage_id == *target)
                    {
                        return Err(Error::Refused(Refusal::UnsupportedDependency));
                    }
                }
                for field in &message_info.field_infos {
                    for target in &field.object_references {
                        budget.charge(LimitKind::WireWork, records.len())?;
                        let allowed = records
                            .iter()
                            .any(|record| record.resolved.storage_id == *target);
                        let expected = records.iter().any(|record| {
                            record.resolved.template_id == owner
                                && record.resolved.storage_id == *target
                        });
                        if allowed && !expected {
                            return Err(Error::Refused(Refusal::UnsupportedDependency));
                        }
                    }
                    for target in &field.data_references {
                        budget.charge(LimitKind::WireWork, records.len())?;
                        if records
                            .iter()
                            .any(|record| record.resolved.storage_id == *target)
                        {
                            return Err(Error::Refused(Refusal::UnsupportedDependency));
                        }
                    }
                }
            }
        }
    }
    Ok(())
}

fn reject_metadata_storage_owners(
    source: &Package,
    records: &[SlotRecord],
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let Some(component) = source.state.source.components().get("Index/Metadata.iwa") else {
        return Ok(());
    };
    let mut payload = None;
    for object in &component.archive().objects {
        for message in &object.messages {
            if message.type_ == METADATA_MESSAGE_TYPE
                && payload.replace(message.data.as_slice()).is_some()
            {
                return Err(Error::InvalidSource);
            }
        }
    }
    let payload = payload.ok_or(Error::InvalidSource)?;
    budget.charge(
        LimitKind::WireWork,
        payload.len().saturating_mul(records.len()),
    )?;
    let mut visitor = MetadataOwnership {
        source,
        records,
        unsupported: false,
    };
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        payload,
        metadata_options(source, payload.len(), budget)?,
        &mut visitor,
    )
    .map_err(map_metadata_error)?;
    budget.charge_metadata_report(inspection.report())?;
    if visitor.unsupported {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(())
}

fn charge_catalog_census(source: &Package, budget: &mut TransactionBudget) -> Result<(), Error> {
    let source_bytes = source.state.source.source_bytes().len();
    budget.observe(LimitKind::InputBytes, source_bytes)?;
    budget.observe(LimitKind::TotalBytes, source_bytes)?;
    let package = source.state.source.package();
    budget.observe(LimitKind::Entries, package.len())?;
    let mut total_entry_bytes = 0usize;
    for entry in package.iter() {
        budget.observe(LimitKind::EntryBytes, entry.data().len())?;
        total_entry_bytes = total_entry_bytes
            .checked_add(entry.data().len())
            .ok_or(Error::InvalidSource)?;
        budget.charge(
            LimitKind::WireWork,
            entry
                .name()
                .len()
                .saturating_add(entry.raw_name().len())
                .saturating_add(1),
        )?;
    }
    budget.observe(LimitKind::TotalBytes, total_entry_bytes)?;
    for component in source.state.source.components().iter() {
        budget.charge(LimitKind::Components, 1)?;
        budget.charge(
            LimitKind::WireWork,
            component.name().len().saturating_add(1),
        )?;
        for object in &component.archive().objects {
            budget.charge(LimitKind::WireWork, object.messages.len().saturating_add(1))?;
            for info in &object.archive_info.message_infos {
                let references = info
                    .object_references
                    .len()
                    .saturating_add(info.data_references.len())
                    .saturating_add(
                        info.field_infos
                            .iter()
                            .map(|field| {
                                field
                                    .object_references
                                    .len()
                                    .saturating_add(field.data_references.len())
                            })
                            .sum::<usize>(),
                    );
                budget.charge(LimitKind::References, references)?;
                budget.charge(
                    LimitKind::WireWork,
                    references.saturating_add(info.field_infos.len()),
                )?;
            }
        }
    }
    Ok(())
}

fn precharge_archive_rewrite(
    archive: &litchi_iwa_core::Archive,
    original_message_bytes: usize,
    rewritten_message_bytes: usize,
    archive_limits: litchi_iwa_core::Limits,
    snappy_limits: litchi_iwa_core::SnappyLimits,
    budget: &mut TransactionBudget,
) -> Result<(usize, usize), Error> {
    let source_archive_bytes = archive
        .encoded_len(archive_limits)
        .map_err(map_core_error)?;
    let message_count = archive
        .objects
        .iter()
        .try_fold(0usize, |count, object| {
            count.checked_add(object.messages.len())
        })
        .ok_or(Error::InvalidSource)?;
    // Replacing one message can widen the message length, object length, and
    // outer archive prefixes.  Reserve the exact payload delta plus the
    // maximum canonical varint growth for every enclosing record.  This is a
    // source-derived upper bound and requires no archive/output allocation.
    let framing_slack = archive
        .objects
        .len()
        .checked_add(message_count)
        .and_then(|value| value.checked_add(2))
        .and_then(|value| value.checked_mul(MAX_VARINT_BYTES.saturating_mul(3)))
        .ok_or(Error::InvalidSource)?;
    let archive_bound = source_archive_bytes
        .checked_sub(original_message_bytes)
        .and_then(|value| value.checked_add(rewritten_message_bytes))
        .and_then(|value| value.checked_add(framing_slack))
        .ok_or(Error::InvalidSource)?;
    let compressed_bound = SnappyStream::maximum_compressed_len(archive_bound, snappy_limits)
        .map_err(map_core_error)?;
    budget.observe(LimitKind::OutputBytes, rewritten_message_bytes)?;
    budget.observe(LimitKind::OutputBytes, archive_bound)?;
    budget.observe(LimitKind::EntryBytes, compressed_bound)?;
    budget.charge(
        LimitKind::WireWork,
        source_archive_bytes
            .saturating_add(archive_bound)
            .saturating_add(compressed_bound),
    )?;
    budget.observe(
        LimitKind::ScratchBytes,
        source_archive_bytes
            .saturating_add(archive_bound)
            .saturating_add(compressed_bound),
    )?;
    budget.observe(
        LimitKind::RetainedBytes,
        rewritten_message_bytes
            .saturating_add(archive_bound)
            .saturating_add(compressed_bound),
    )?;
    budget.charge(LimitKind::Allocations, 3)?;
    Ok((archive_bound, compressed_bound))
}

fn precharge_candidate_catalog(
    source: &Package,
    rewrites: [(&str, usize, usize); 2],
    budget: &mut TransactionBudget,
) -> Result<usize, Error> {
    let catalog = &source.state.source;
    let mut package_output_bound = catalog.source_bytes().len();
    for (member, member_bytes, _) in rewrites {
        budget.charge(
            LimitKind::WireWork,
            catalog
                .package()
                .len()
                .saturating_mul(member.len().saturating_add(1)),
        )?;
        let entry = catalog
            .package()
            .iter()
            .find(|entry| entry.name() == member)
            .ok_or(Error::InvalidSource)?;
        let replacement_bound = match entry.metadata().central().compression_method() {
            0 => member_bytes,
            8 => super::table_lock::deflate_compressed_bound(member_bytes)
                .ok_or(Error::InvalidSource)?,
            _ => return Err(Error::Refused(Refusal::UnsupportedSource)),
        };
        let old_compressed = usize::try_from(entry.metadata().compressed_size())
            .map_err(|_| Error::InvalidSource)?;
        package_output_bound = package_output_bound
            .checked_sub(old_compressed)
            .and_then(|value| value.checked_add(replacement_bound))
            .ok_or(Error::InvalidSource)?;
    }
    budget.observe(LimitKind::OutputBytes, package_output_bound)?;
    budget.observe(LimitKind::InputBytes, package_output_bound)?;
    budget.observe(LimitKind::RetainedBytes, package_output_bound)?;
    budget.observe(LimitKind::ScratchBytes, package_output_bound)?;
    budget.observe(LimitKind::Entries, catalog.package().len())?;
    let mut total_entry_bytes = 0usize;
    for entry in catalog.package().iter() {
        budget.charge(
            LimitKind::WireWork,
            rewrites
                .len()
                .saturating_mul(entry.name().len().saturating_add(1)),
        )?;
        let entry_bytes = rewrites
            .iter()
            .find(|(member, _, _)| *member == entry.name())
            .map_or(entry.data().len(), |(_, bytes, _)| *bytes);
        budget.observe(LimitKind::EntryBytes, entry_bytes)?;
        total_entry_bytes = total_entry_bytes
            .checked_add(entry_bytes)
            .ok_or(Error::InvalidSource)?;
        budget.charge(
            LimitKind::WireWork,
            entry.name().len().saturating_add(entry_bytes),
        )?;
    }
    budget.observe(LimitKind::TotalBytes, total_entry_bytes)?;
    let archive_limits = catalog
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    for component in catalog.components().iter() {
        budget.charge(LimitKind::Components, 1)?;
        budget.charge(
            LimitKind::WireWork,
            rewrites
                .len()
                .saturating_mul(component.name().len().saturating_add(1)),
        )?;
        let archive_bytes = rewrites
            .iter()
            .find(|(member, _, _)| *member == component.name())
            .map_or_else(
                || component.archive().encoded_len(archive_limits),
                |(_, _, archive_bound)| Ok(*archive_bound),
            )
            .map_err(map_core_error)?;
        budget.charge(LimitKind::WireWork, archive_bytes)?;
        for object in &component.archive().objects {
            budget.charge(LimitKind::WireWork, object.messages.len().saturating_add(1))?;
            for info in &object.archive_info.message_infos {
                let references = info
                    .object_references
                    .len()
                    .saturating_add(info.data_references.len())
                    .saturating_add(
                        info.field_infos
                            .iter()
                            .map(|field| {
                                field
                                    .object_references
                                    .len()
                                    .saturating_add(field.data_references.len())
                            })
                            .sum::<usize>(),
                    );
                budget.charge(LimitKind::References, references)?;
                budget.charge(
                    LimitKind::WireWork,
                    references.saturating_add(info.field_infos.len()),
                )?;
            }
        }
    }
    Ok(package_output_bound)
}

struct ComponentRewrite {
    compressed: Vec<u8>,
    archive_bound: usize,
}

fn rewrite_storage_component(
    source: &Package,
    storage_id: u64,
    before: &str,
    after: &str,
    budget: &mut TransactionBudget,
) -> Result<ComponentRewrite, Error> {
    let component_name = component_name_for_object(source, storage_id, budget)?;
    let component = source
        .state
        .source
        .components()
        .get(&component_name)
        .ok_or(Error::InvalidSource)?;
    let object = component
        .archive()
        .object(storage_id)
        .ok_or(Error::InvalidSource)?;
    let mut message_index = None;
    for (index, message) in object.messages.iter().enumerate() {
        if is_body_text_message_type(message.type_) {
            if message_index.replace(index).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    let message_index = message_index.ok_or(Error::InvalidSource)?;
    let message = &object.messages[message_index];
    let limits = residual_storage_limits(
        storage_rewrite_limits(source.state.source.limits()).map_err(|_| Error::InvalidSource)?,
        budget,
    )?;
    let validation =
        litchi_iwa_text_wire::validate_storage(&message.data, limits).map_err(map_text_error)?;
    if validation.storage_kind() != 1 {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    budget.charge_text_validation(message.data.len(), validation)?;
    budget.charge_text_allocation(validation)?;
    let limits = residual_storage_limits(
        storage_rewrite_limits(source.state.source.limits()).map_err(|_| Error::InvalidSource)?,
        budget,
    )?;
    let validated =
        litchi_iwa_text_wire::decode_storage(&message.data, limits).map_err(map_text_error)?;
    if validated.validation() != validation {
        return Err(Error::Verification);
    }
    budget.charge_text_validation(message.data.len(), validated.validation())?;
    if validated.storage().text() != before {
        return Err(Error::PatchConflict);
    }
    let units = before.encode_utf16().count();
    budget.observe(
        LimitKind::ScratchBytes,
        message
            .data
            .len()
            .saturating_add(after.len())
            .saturating_add(
                validation
                    .fragments()
                    .saturating_mul(size_of::<usize>())
                    .saturating_mul(8),
            ),
    )?;
    budget.charge(LimitKind::Allocations, 3)?;
    let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior(
        &message.data,
        0..units,
        after,
        litchi_iwa_text_wire::RewriteBehavior::PreserveOnEqualText,
        limits,
    )
    .map_err(map_text_error)?;
    let requirements = prepared.execution_requirements();
    budget.charge_text_requirements(requirements)?;
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = source
        .state
        .source
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let (archive_bound, compressed_bound) = precharge_archive_rewrite(
        component.archive(),
        message.data.len(),
        requirements.output_bytes(),
        archive_limits,
        snappy_limits,
        budget,
    )?;
    let rewritten = prepared
        .execute(litchi_iwa_text_wire::StorageRewriteExecutionLimits {
            max_output_bytes: requirements.output_bytes(),
            max_retained_elements: requirements.retained_elements(),
            max_retained_bytes: requirements.retained_bytes(),
            max_peak_scratch_bytes: requirements.peak_scratch_bytes(),
            max_allocations: requirements.allocations(),
            max_work: requirements.work(),
        })
        .map_err(map_text_error)?;
    let report = rewritten.execution_report();
    if rewritten.bytes().len() != requirements.output_bytes()
        || report.retained_elements > requirements.retained_elements()
        || report.retained_bytes > requirements.retained_bytes()
        || report.peak_scratch_bytes > requirements.peak_scratch_bytes()
        || report.allocations > requirements.allocations()
        || report.work > requirements.work()
    {
        return Err(Error::Verification);
    }
    if !rewritten.removed_object_references().is_empty()
        || rewritten.object_reference_occurrences_before()
            != rewritten.object_reference_occurrences_after()
    {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    let message_type = message.type_;
    let rewritten = rewritten.into_bytes();
    let compressed_source_len = source
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == component_name)
        .map_or(0, |entry| entry.data().len());
    budget.observe(
        LimitKind::EntryBytes,
        compressed_source_len.saturating_add(rewritten.len()),
    )?;
    budget.charge(LimitKind::Allocations, 2)?;
    let mut archive = component.archive().clone();
    archive
        .object_mut(storage_id)
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: message_type,
                data: rewritten,
            },
            source
                .state
                .source
                .limits()
                .effective_archive_limits()
                .map_err(map_archive_error)?,
        )
        .map_err(map_core_error)?;
    let exact_archive_len = archive
        .encoded_len(
            source
                .state
                .source
                .limits()
                .effective_archive_limits()
                .map_err(map_archive_error)?,
        )
        .map_err(map_core_error)?;
    if exact_archive_len > archive_bound {
        return Err(Error::Verification);
    }
    let bytes = archive
        .to_bytes(
            source
                .state
                .source
                .limits()
                .effective_archive_limits()
                .map_err(map_archive_error)?,
        )
        .map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&bytes, snappy_limits).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(Error::Verification);
    }
    Ok(ComponentRewrite {
        compressed,
        archive_bound,
    })
}

struct MetadataRewrite {
    member: String,
    compressed: Vec<u8>,
    archive_bound: usize,
}

fn rewrite_metadata_tokens(
    source: &Package,
    changed_member: &str,
    budget: &mut TransactionBudget,
) -> Result<MetadataRewrite, Error> {
    let component = source
        .state
        .source
        .components()
        .get("Index/Metadata.iwa")
        .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let mut location = None;
    for (object_index, object) in component.archive().objects.iter().enumerate() {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == METADATA_MESSAGE_TYPE {
                if location.replace((object_index, message_index)).is_some() {
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    let (object_index, message_index) = location.ok_or(Error::InvalidSource)?;
    let payload = component.archive().objects[object_index].messages[message_index]
        .data
        .as_slice();
    let target = changed_member
        .strip_prefix("Index/")
        .and_then(|value| value.strip_suffix(".iwa"))
        .ok_or(Error::InvalidSource)?;
    budget.charge(LimitKind::Allocations, 1)?;
    budget.charge(LimitKind::ScratchBytes, payload.len())?;
    let mut visitor = MetadataMatch {
        target,
        found: None,
        ambiguous: false,
    };
    let options = metadata_options(source, payload.len(), budget)?;
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        payload,
        options,
        &mut visitor,
    )
    .map_err(map_metadata_error)?;
    budget.charge_metadata_report(inspection.report())?;
    if visitor.ambiguous {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    let matching = visitor
        .found
        .as_ref()
        .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let rewrite_options = metadata_options(source, payload.len(), budget)?;
    let selector = package_metadata_codec::ComponentSelector::new(matching.0, matching.1.as_str());
    let batch = package_metadata_codec::SaveTokenBatch::new(std::slice::from_ref(&selector));
    budget.observe(
        LimitKind::ScratchBytes,
        payload.len().saturating_mul(2).saturating_add(
            source
                .state
                .source
                .components()
                .len()
                .saturating_mul(size_of::<usize>() * 4),
        ),
    )?;
    budget.charge(LimitKind::Allocations, 2)?;
    let prepared = package_metadata_codec::prepare_package_metadata_save_tokens(
        payload,
        batch,
        rewrite_options,
    )
    .map_err(map_metadata_error)?;
    let execution_limits = prepared.execution_requirements().exact_limits();
    budget.charge_metadata_report(prepared.prepare_report())?;
    budget.charge_metadata_requirements(prepared.execution_requirements())?;
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = source
        .state
        .source
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let (archive_bound, compressed_bound) = precharge_archive_rewrite(
        component.archive(),
        payload.len(),
        prepared.execution_requirements().output_bytes(),
        archive_limits,
        snappy_limits,
        budget,
    )?;
    let output = prepared
        .execute(execution_limits)
        .map_err(map_metadata_error)?
        .into_bytes();
    let mut archive = component.archive().clone();
    archive
        .object_mut(
            archive.objects[object_index]
                .archive_info
                .identifier
                .ok_or(Error::InvalidSource)?,
        )
        .ok_or(Error::InvalidSource)?
        .replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: output,
            },
            source
                .state
                .source
                .limits()
                .effective_archive_limits()
                .map_err(map_archive_error)?,
        )
        .map_err(map_core_error)?;
    let exact_archive_len = archive
        .encoded_len(archive_limits)
        .map_err(map_core_error)?;
    if exact_archive_len > archive_bound {
        return Err(Error::Verification);
    }
    let bytes = archive.to_bytes(archive_limits).map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&bytes, snappy_limits).map_err(map_core_error)?;
    if compressed.len() > compressed_bound {
        return Err(Error::Verification);
    }
    Ok(MetadataRewrite {
        member: component.name().to_owned(),
        compressed,
        archive_bound,
    })
}

fn metadata_options(
    source: &Package,
    payload_len: usize,
    budget: &TransactionBudget,
) -> Result<package_metadata_codec::RewriteOptions, Error> {
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let input =
        residual_limit(budget, LimitKind::WireBytes, payload_len.max(1))?.min(payload_len.max(1));
    let maximum = residual_limit(budget, LimitKind::OutputBytes, input)?.min(i32::MAX as usize);
    let fields = residual_limit(budget, LimitKind::WireFields, 1)?;
    let work = residual_limit(budget, LimitKind::WireWork, input)?;
    let components = residual_limit(budget, LimitKind::Components, 1)?;
    let references = residual_limit(budget, LimitKind::References, 1)?;
    Ok(package_metadata_codec::RewriteOptions::new(
        input,
        maximum,
        fields.min(archive_limits.max_header_fields()),
        work,
        u32::try_from(
            archive_limits
                .max_header_nesting()
                .min(budget.maximum(LimitKind::WireNesting)),
        )
        .unwrap_or(u32::MAX),
        components,
        references,
        0,
    ))
}

fn normalized_metadata_locator(locator: &str) -> &str {
    let locator = locator.strip_prefix("Index/").unwrap_or(locator);
    locator.strip_suffix(".iwa").unwrap_or(locator)
}

fn template_decode_options(
    source: &Package,
    payload: &[u8],
    budget: &TransactionBudget,
) -> Result<pages_header_footer_codec::DecodeOptions, Error> {
    let archive = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let input = residual_limit(budget, LimitKind::WireBytes, payload.len().max(1))?
        .min(payload.len().max(1));
    Ok(pages_header_footer_codec::DecodeOptions::new(
        input,
        payload.len().saturating_mul(2).max(1),
        residual_limit(budget, LimitKind::WireFields, 1)?
            .min(archive.max_header_fields())
            .max(1),
        residual_limit(budget, LimitKind::WireWork, 1)?.max(1),
        u32::try_from(
            archive
                .max_header_nesting()
                .min(budget.maximum(LimitKind::WireNesting)),
        )
        .unwrap_or(u32::MAX),
        residual_limit(budget, LimitKind::References, 1)?.max(1),
    ))
}

fn residual_storage_limits(
    base: litchi_iwa_text_wire::RewriteLimits,
    budget: &TransactionBudget,
) -> Result<litchi_iwa_text_wire::RewriteLimits, Error> {
    litchi_iwa_text_wire::RewriteLimits::new(
        residual_limit(budget, LimitKind::WireBytes, 1)?
            .min(base.max_message_bytes())
            .max(1),
        residual_limit(budget, LimitKind::WireFields, 1)?
            .min(base.max_fields())
            .max(1),
        budget
            .maximum(LimitKind::WireNesting)
            .min(base.max_nesting())
            .max(1),
        residual_limit(budget, LimitKind::Entries, 1)?
            .min(base.max_fragments())
            .max(1),
        residual_limit(budget, LimitKind::TextBytes, 1)?
            .min(base.max_text_bytes())
            .max(1),
        residual_limit(budget, LimitKind::Entries, 1)?
            .min(base.max_table_entries())
            .max(1),
        residual_limit(budget, LimitKind::References, 1)?
            .min(base.max_object_references())
            .max(1),
        residual_limit(budget, LimitKind::OutputBytes, 1)?
            .min(base.max_output_bytes())
            .max(1),
        residual_limit(budget, LimitKind::WireWork, 1)?
            .min(base.max_rewrite_work())
            .max(1),
    )
    .map_err(map_text_error)
}

fn residual_limit(
    budget: &TransactionBudget,
    kind: LimitKind,
    required: usize,
) -> Result<usize, Error> {
    let maximum = budget.remaining(kind);
    if maximum < required {
        return Err(Error::Limit {
            kind,
            observed: required as u64,
            maximum: maximum as u64,
        });
    }
    Ok(maximum)
}

fn verify_locality(
    source: &Package,
    candidate: &Package,
    changed_component: &str,
    metadata_component: &str,
    deleted_previews: &[String],
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    for source_entry in source.state.source.package().iter() {
        budget.charge(
            LimitKind::WireWork,
            deleted_previews
                .len()
                .saturating_mul(source_entry.name().len().saturating_add(1)),
        )?;
        if deleted_previews
            .iter()
            .any(|name| name == source_entry.name())
        {
            continue;
        }
        budget.charge(
            LimitKind::WireWork,
            candidate
                .state
                .source
                .package()
                .len()
                .saturating_mul(source_entry.name().len().saturating_add(1)),
        )?;
        let target = candidate
            .state
            .source
            .package()
            .iter()
            .find(|entry| entry.name() == source_entry.name())
            .ok_or(Error::Verification)?;
        budget.charge(
            LimitKind::WireWork,
            source_entry
                .name()
                .len()
                .saturating_add(source_entry.data().len())
                .saturating_add(target.data().len()),
        )?;
        if source_entry.name() != changed_component
            && source_entry.name() != metadata_component
            && source_entry.data() != target.data()
        {
            return Err(Error::Verification);
        }
    }
    for target_entry in candidate.state.source.package().iter() {
        if PREVIEW_NAMES.contains(&target_entry.name()) {
            continue;
        }
        budget.charge(
            LimitKind::WireWork,
            source
                .state
                .source
                .package()
                .len()
                .saturating_mul(target_entry.name().len().saturating_add(1)),
        )?;
        if source
            .state
            .source
            .package()
            .iter()
            .all(|entry| entry.name() != target_entry.name())
        {
            return Err(Error::Verification);
        }
    }
    Ok(())
}

fn verify_member_bytes(
    candidate: &Package,
    member: &str,
    expected: &[u8],
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    budget.charge(
        LimitKind::WireWork,
        candidate
            .state
            .source
            .package()
            .len()
            .saturating_mul(member.len().saturating_add(1))
            .saturating_add(expected.len()),
    )?;
    let actual = candidate
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or(Error::Verification)?;
    if actual.data() != expected {
        return Err(Error::Verification);
    }
    Ok(())
}

fn validate_authored(value: &str) -> Result<(), Error> {
    if value.contains('\u{000e}') || value.contains('\u{fffc}') {
        return Err(Error::Refused(Refusal::StructuralMarker));
    }
    let bytes = value.len();
    let units = value.encode_utf16().count();
    if units > u32::MAX as usize {
        return Err(Error::Limit {
            kind: LimitKind::TextUnits,
            observed: units as u64,
            maximum: u64::from(u32::MAX),
        });
    }
    if bytes > crate::DEFAULT_MAX_TEXT_BYTES {
        return Err(Error::Refused(Refusal::TextTooLarge));
    }
    Ok(())
}

fn utf16_boundary(value: &str, index: usize) -> Result<usize, Error> {
    let mut units = 0usize;
    for (byte, character) in value.char_indices() {
        if units == index {
            return Ok(byte);
        }
        units = units
            .checked_add(character.len_utf16())
            .ok_or(Error::Refused(Refusal::PositionOutOfBounds))?;
    }
    if units == index {
        Ok(value.len())
    } else {
        Err(Error::Refused(Refusal::PositionOutOfBounds))
    }
}

fn nonzero(value: u64) -> Result<std::num::NonZeroU64, Error> {
    std::num::NonZeroU64::new(value).ok_or(Error::InvalidSource)
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    hash
}

fn map_package_error(error: PackageError) -> Error {
    match error {
        PackageError::Allocation { amount } => Error::Allocation { amount },
        PackageError::PayloadLimit { observed, limit } => Error::Limit {
            kind: LimitKind::WireBytes,
            observed: observed as u64,
            maximum: limit as u64,
        },
        PackageError::ObjectLimit { observed, limit } => Error::Limit {
            kind: LimitKind::Entries,
            observed: observed as u64,
            maximum: limit as u64,
        },
        PackageError::Semantic(_) => Error::InvalidSource,
        PackageError::Archive(error) => map_archive_error(error),
        _ => Error::InvalidSource,
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
                litchi_iwa_archive::LimitKind::EntryBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => LimitKind::TotalBytes,
                litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::WireBytes,
                litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalBytes,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => LimitKind::WireBytes,
                litchi_iwa_archive::LimitKind::CompressedEntryBytes => LimitKind::EntryBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
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
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes => LimitKind::WireBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => LimitKind::Entries,
                litchi_iwa_core::LimitKind::MetadataItems => LimitKind::References,
                _ => LimitKind::WireBytes,
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

fn map_text_error(error: litchi_iwa_text_wire::RewriteError) -> Error {
    match error {
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => Error::Limit {
            kind: match resource {
                "input bytes" | "input" => LimitKind::InputBytes,
                "output bytes" | "output" => LimitKind::OutputBytes,
                "fields" => LimitKind::WireFields,
                "nesting" => LimitKind::WireNesting,
                "work" => LimitKind::WireWork,
                _ => LimitKind::Entries,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_text_wire::RewriteError::Allocation { amount, .. } => {
            Error::Allocation { amount }
        },
        _ => Error::InvalidSource,
    }
}

fn map_metadata_error(error: package_metadata_codec::RewriteError) -> Error {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum): (LimitKind, u64, u64) = match limit {
            package_metadata_codec::RewriteLimit::InputBytes { observed, maximum } => {
                (LimitKind::InputBytes, observed as u64, maximum as u64)
            },
            package_metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => {
                (LimitKind::OutputBytes, observed as u64, maximum as u64)
            },
            package_metadata_codec::RewriteLimit::Fields { observed, maximum } => {
                (LimitKind::WireFields, observed as u64, maximum as u64)
            },
            package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (LimitKind::WireWork, observed as u64, maximum as u64)
            },
            package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => (
                LimitKind::WireNesting,
                u64::from(observed),
                u64::from(maximum),
            ),
            package_metadata_codec::RewriteLimit::Components { observed, maximum } => {
                (LimitKind::Components, observed as u64, maximum as u64)
            },
            package_metadata_codec::RewriteLimit::References { observed, maximum } => {
                (LimitKind::References, observed as u64, maximum as u64)
            },
            package_metadata_codec::RewriteLimit::Additions { observed, maximum } => {
                (LimitKind::RegistryChanges, observed as u64, maximum as u64)
            },
            _ => return Error::InvalidSource,
        };
        return Error::Limit {
            kind,
            observed,
            maximum,
        };
    }
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount };
    }
    match error.invalid_reason() {
        Some(package_metadata_codec::InvalidReason::ExistingReferenceCollision)
        | Some(package_metadata_codec::InvalidReason::CrossComponentRemoval)
        | Some(package_metadata_codec::InvalidReason::VersionedRemoval) => {
            Error::Refused(Refusal::UnsupportedDependency)
        },
        _ => Error::InvalidSource,
    }
}

fn map_template_error(error: pages_header_footer_codec::DecodeError) -> Error {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            pages_header_footer_codec::WireResourceLimit::InputBytes { observed, maximum } => {
                (LimitKind::InputBytes, observed, maximum)
            },
            pages_header_footer_codec::WireResourceLimit::OutputBytes { observed, maximum } => {
                (LimitKind::OutputBytes, observed, maximum)
            },
            pages_header_footer_codec::WireResourceLimit::Fields { observed, maximum } => {
                (LimitKind::WireFields, observed, maximum)
            },
            pages_header_footer_codec::WireResourceLimit::WorkBytes { observed, maximum } => {
                (LimitKind::WireWork, observed, maximum)
            },
            pages_header_footer_codec::WireResourceLimit::Nesting { observed, maximum } => {
                (LimitKind::WireNesting, observed as usize, maximum as usize)
            },
            pages_header_footer_codec::WireResourceLimit::References { observed, maximum } => {
                (LimitKind::References, observed, maximum)
            },
            _ => return Error::InvalidSource,
        };
        return Error::Limit {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return Error::Allocation { amount };
    }
    Error::InvalidSource
}

fn map_section_error(error: pages_section_codec::DecodeError) -> Error {
    let Some(limit) = error.resource_limit() else {
        return Error::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        pages_section_codec::DecodeLimit::Bytes { observed, maximum } => {
            (LimitKind::WireBytes, observed as u64, maximum as u64)
        },
        pages_section_codec::DecodeLimit::Fields { observed, maximum } => {
            (LimitKind::WireFields, observed as u64, maximum as u64)
        },
        pages_section_codec::DecodeLimit::Work { observed, maximum } => {
            (LimitKind::WireWork, observed as u64, maximum as u64)
        },
        pages_section_codec::DecodeLimit::Nesting { observed, maximum } => (
            LimitKind::WireNesting,
            u64::from(observed),
            u64::from(maximum),
        ),
        pages_section_codec::DecodeLimit::NameBytes { observed, maximum } => {
            (LimitKind::RetainedBytes, observed as u64, maximum as u64)
        },
        _ => return Error::InvalidSource,
    };
    Error::Limit {
        kind,
        observed,
        maximum,
    }
}

/// The native storage identifier of one rooted header/footer slot, for the
/// storage-object families (number attachments) that edit its text storage.
pub(crate) fn header_footer_storage(
    source: &Package,
    selector: HeaderFooterSelector<'_>,
) -> Result<u64, Error> {
    let mut budget = TransactionBudget::new(source)?;
    let section_index = resolve_section_index(source, selector.section(), &mut budget)?;
    let canonical = HeaderFooterSelector::index(
        section_index,
        selector.template(),
        selector.kind(),
        selector.slot_index(),
    );
    let records = discover_slots(source, &mut budget, true)?;
    Ok(select_record(&records, canonical, &mut budget)?
        .resolved
        .storage_id)
}
