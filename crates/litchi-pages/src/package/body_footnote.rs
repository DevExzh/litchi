//! Exact-source lifecycle transactions for Pages body footnotes.
//!
//! The public surface is archive-free and selector-first.  The physical
//! insertion/removal implementation is intentionally kept in this private
//! module so native identifiers, component routes, and wire values cannot
//! escape through `litchi-pages`.

use std::{fmt, mem::size_of, num::NonZeroU64, sync::Arc};

use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ReassemblyExecutionRequirements},
};
use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    package_metadata_codec::{
        self, ComponentSelector, DataReferenceOwnerRemoval, ExternalReferenceAddition,
        ExternalReferenceRemoval, ObjectUuidAddition, ObjectUuidRemoval, PackageMetadataVisitor,
        RemovalBatch, RemovalSaveTokenBatch, RewriteError as MetadataRewriteError,
        RewriteOptions as MetadataRewriteOptions, SaveTokenBatch, UuidBits,
    },
    pages_footnote_graph_codec::{
        self as graph_codec, BodyFootnoteEntryWrite, BodyFootnoteTableWrite,
        DecodeOptions as GraphDecodeOptions, FootnoteGraphWrite,
    },
};
use litchi_iwa_text_wire::{RewriteBehavior, RewriteLimits, StorageRewriteExecutionLimits};
use thiserror::Error as ThisError;

use super::Package;
use crate::footnote::body::{Footnote, MAX_CUSTOM_MARK_BYTES, MAX_TEXT_BYTES, Position, Selector};

const FOOTNOTE_REFERENCE_MESSAGE_TYPE: u32 = 2_008;
const TEXTUAL_ATTACHMENT_MESSAGE_TYPE: u32 = 2_004;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const FOOTNOTE_TABLE_FIELD: u32 = 16;
const FOOTNOTE_ANCHOR_TEXT: &str = "\u{000e}";
const PREVIEW_NAMES: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

/// Aggregate accounting for one immutable insert/remove/apply transaction.
///
/// The individual codecs deliberately expose exact phase reports, but a
/// lifecycle operation has several simultaneously live buffers (native
/// archives, metadata candidates, compressed entries, and the final ZIP).
/// This coordinator merges those reports before allowing the next
/// allocation-bearing phase to execute.  Every counter is checked, including
/// the arithmetic used to derive an observed amount.
#[derive(Debug, Clone, Copy)]
struct TransactionBudget {
    max_input_bytes: u64,
    max_output_bytes: u64,
    max_entries: u64,
    max_entry_bytes: u64,
    max_total_bytes: u64,
    max_text_bytes: u64,
    max_text_units: u64,
    max_wire_bytes: u64,
    max_wire_fields: u64,
    max_wire_nesting: u64,
    max_wire_work: u64,
    max_references: u64,
    max_components: u64,
    max_registry_changes: u64,
    max_retained_bytes: u64,
    max_scratch_bytes: u64,
    max_allocations: u64,
    input_bytes: u64,
    output_bytes: u64,
    entries: u64,
    entry_bytes: u64,
    total_bytes: u64,
    text_bytes: u64,
    text_units: u64,
    wire_bytes: u64,
    wire_fields: u64,
    wire_nesting: u64,
    wire_work: u64,
    references: u64,
    components: u64,
    registry_changes: u64,
    retained_bytes: u64,
    scratch_bytes: u64,
    allocations: u64,
}

impl TransactionBudget {
    fn new(source: &Package) -> Result<Self, Error> {
        let limits = source.state.source.limits();
        let archive_limits = limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let max_total_bytes = limits.max_total_bytes();
        let max_iwa_stream_bytes =
            u64::try_from(limits.max_iwa_stream_bytes()).map_err(|_| Error::InvalidSource)?;
        let max_wire_work = max_total_bytes
            .checked_mul(4)
            .ok_or(Error::InvalidSource)?
            .max(max_iwa_stream_bytes);
        let mut budget = Self {
            max_input_bytes: limits.max_input_bytes(),
            max_output_bytes: limits.max_input_bytes(),
            max_entries: u64::try_from(limits.max_entries()).map_err(|_| Error::InvalidSource)?,
            max_entry_bytes: limits.max_entry_bytes(),
            max_total_bytes,
            max_text_bytes: max_iwa_stream_bytes,
            max_text_units: max_iwa_stream_bytes,
            max_wire_bytes: max_iwa_stream_bytes,
            max_wire_fields: u64::try_from(archive_limits.max_header_fields())
                .map_err(|_| Error::InvalidSource)?,
            max_wire_nesting: u64::try_from(archive_limits.max_header_nesting())
                .map_err(|_| Error::InvalidSource)?,
            max_wire_work,
            max_references: u64::try_from(archive_limits.max_metadata_items())
                .map_err(|_| Error::InvalidSource)?,
            max_components: u64::try_from(limits.max_entries())
                .map_err(|_| Error::InvalidSource)?,
            max_registry_changes: u64::try_from(limits.max_entries())
                .map_err(|_| Error::InvalidSource)?,
            max_retained_bytes: max_total_bytes,
            max_scratch_bytes: max_total_bytes,
            max_allocations: u64::try_from(limits.max_entries())
                .map_err(|_| Error::InvalidSource)?,
            input_bytes: 0,
            output_bytes: 0,
            entries: 0,
            entry_bytes: 0,
            total_bytes: 0,
            text_bytes: 0,
            text_units: 0,
            wire_bytes: 0,
            wire_fields: 0,
            wire_nesting: 0,
            wire_work: 0,
            references: 0,
            components: 0,
            registry_changes: 0,
            retained_bytes: 0,
            scratch_bytes: 0,
            allocations: 0,
        };
        budget.charge(
            LimitKind::InputBytes,
            source.state.source.source_bytes().len(),
        )?;
        Ok(budget)
    }

    fn charge(&mut self, kind: LimitKind, amount: usize) -> Result<(), Error> {
        let amount = u64::try_from(amount).map_err(|_| Error::Limit {
            kind,
            observed: u64::MAX,
            maximum: self.maximum(kind),
        })?;
        self.charge_u64(kind, amount)
    }

    fn charge_u64(&mut self, kind: LimitKind, amount: u64) -> Result<(), Error> {
        let current = *self.slot_mut(kind);
        let maximum = self.maximum(kind);
        let observed = current.checked_add(amount).ok_or(Error::Limit {
            kind,
            observed: u64::MAX,
            maximum,
        })?;
        if observed > maximum {
            return Err(Error::Limit {
                kind,
                observed,
                maximum,
            });
        }
        *self.slot_mut(kind) = observed;
        Ok(())
    }

    fn observe(&mut self, kind: LimitKind, amount: usize) -> Result<(), Error> {
        let amount = u64::try_from(amount).map_err(|_| Error::Limit {
            kind,
            observed: u64::MAX,
            maximum: self.maximum(kind),
        })?;
        let current = *self.slot_mut(kind);
        if amount > current {
            let maximum = self.maximum(kind);
            if amount > maximum {
                return Err(Error::Limit {
                    kind,
                    observed: amount,
                    maximum,
                });
            }
            *self.slot_mut(kind) = amount;
        }
        Ok(())
    }

    fn maximum(self, kind: LimitKind) -> u64 {
        match kind {
            LimitKind::InputBytes => self.max_input_bytes,
            LimitKind::OutputBytes => self.max_output_bytes,
            LimitKind::Entries => self.max_entries,
            LimitKind::EntryBytes => self.max_entry_bytes,
            LimitKind::TotalBytes => self.max_total_bytes,
            LimitKind::TextBytes => self.max_text_bytes,
            LimitKind::TextUnits => self.max_text_units,
            LimitKind::WireBytes => self.max_wire_bytes,
            LimitKind::WireFields => self.max_wire_fields,
            LimitKind::WireNesting => self.max_wire_nesting,
            LimitKind::WireWork => self.max_wire_work,
            LimitKind::References => self.max_references,
            LimitKind::Components => self.max_components,
            LimitKind::RegistryChanges => self.max_registry_changes,
            LimitKind::RetainedBytes => self.max_retained_bytes,
            LimitKind::ScratchBytes => self.max_scratch_bytes,
            LimitKind::Allocations => self.max_allocations,
        }
    }

    fn slot_mut(&mut self, kind: LimitKind) -> &mut u64 {
        match kind {
            LimitKind::InputBytes => &mut self.input_bytes,
            LimitKind::OutputBytes => &mut self.output_bytes,
            LimitKind::Entries => &mut self.entries,
            LimitKind::EntryBytes => &mut self.entry_bytes,
            LimitKind::TotalBytes => &mut self.total_bytes,
            LimitKind::TextBytes => &mut self.text_bytes,
            LimitKind::TextUnits => &mut self.text_units,
            LimitKind::WireBytes => &mut self.wire_bytes,
            LimitKind::WireFields => &mut self.wire_fields,
            LimitKind::WireNesting => &mut self.wire_nesting,
            LimitKind::WireWork => &mut self.wire_work,
            LimitKind::References => &mut self.references,
            LimitKind::Components => &mut self.components,
            LimitKind::RegistryChanges => &mut self.registry_changes,
            LimitKind::RetainedBytes => &mut self.retained_bytes,
            LimitKind::ScratchBytes => &mut self.scratch_bytes,
            LimitKind::Allocations => &mut self.allocations,
        }
    }

    fn charge_text_requirements(
        &mut self,
        requirements: litchi_iwa_text_wire::StorageRewriteExecutionRequirements,
    ) -> Result<(), Error> {
        self.charge(LimitKind::WireBytes, requirements.output_bytes())?;
        self.charge(LimitKind::Entries, requirements.retained_elements())?;
        self.charge(LimitKind::RetainedBytes, requirements.retained_bytes())?;
        self.charge(LimitKind::ScratchBytes, requirements.peak_scratch_bytes())?;
        self.charge(LimitKind::Allocations, requirements.allocations())?;
        self.charge(LimitKind::WireWork, requirements.work())?;
        self.charge(LimitKind::References, requirements.reference_occurrences())
    }

    fn charge_graph_decode(&mut self, report: graph_codec::DecodeReport) -> Result<(), Error> {
        self.charge(LimitKind::WireBytes, report.input_bytes())?;
        self.charge(LimitKind::WireFields, report.fields())?;
        self.charge(LimitKind::WireWork, report.work_bytes())?;
        self.observe(
            LimitKind::WireNesting,
            usize::try_from(report.max_depth()).map_err(|_| Error::InvalidSource)?,
        )?;
        self.charge(LimitKind::Entries, report.entries())?;
        self.charge(LimitKind::RetainedBytes, report.retained_bytes())?;
        self.charge(LimitKind::ScratchBytes, report.scratch_bytes())
    }

    fn charge_graph_rewrite(
        &mut self,
        report: graph_codec::BodyFootnoteTableRewriteReport,
    ) -> Result<(), Error> {
        self.charge_graph_decode(report.source())?;
        self.charge_graph_decode(report.result())?;
        self.charge(LimitKind::WireBytes, report.output_bytes())?;
        self.charge(LimitKind::WireWork, report.rewrite_work_bytes())?;
        self.charge(LimitKind::Entries, report.entries_after())?;
        self.charge(LimitKind::RetainedBytes, report.retained_bytes())?;
        self.charge(LimitKind::ScratchBytes, report.scratch_bytes())?;
        self.charge(LimitKind::Allocations, report.allocations())
    }

    fn charge_graph_encode(&mut self, report: graph_codec::GraphEncodeReport) -> Result<(), Error> {
        self.charge(LimitKind::WireBytes, report.output_bytes())?;
        self.charge(LimitKind::WireFields, report.fields())?;
        self.charge(LimitKind::WireWork, report.work_bytes())?;
        self.charge(LimitKind::RetainedBytes, report.retained_bytes())?;
        self.charge(LimitKind::ScratchBytes, report.scratch_bytes())?;
        self.charge(LimitKind::Allocations, report.allocations())
    }

    fn charge_metadata_prepare(
        &mut self,
        report: package_metadata_codec::RewriteReport,
    ) -> Result<(), Error> {
        self.charge(LimitKind::InputBytes, report.input_bytes())?;
        self.charge(LimitKind::WireFields, report.fields())?;
        self.charge(LimitKind::WireWork, report.work_bytes())?;
        self.observe(
            LimitKind::WireNesting,
            usize::try_from(report.max_depth()).map_err(|_| Error::InvalidSource)?,
        )?;
        self.charge(LimitKind::Components, report.components_scanned())?;
        self.charge(LimitKind::References, report.references_scanned())?;
        self.charge(
            LimitKind::RegistryChanges,
            report
                .additions()
                .checked_add(report.removals())
                .ok_or(Error::InvalidSource)?,
        )
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

    fn charge_archive_output(
        &mut self,
        encoded_bytes: usize,
        maximum_compressed_bytes: usize,
    ) -> Result<(), Error> {
        self.charge(LimitKind::EntryBytes, encoded_bytes)?;
        self.charge(LimitKind::TotalBytes, encoded_bytes)?;
        self.charge(
            LimitKind::ScratchBytes,
            encoded_bytes
                .checked_add(maximum_compressed_bytes)
                .ok_or(Error::InvalidSource)?,
        )?;
        self.charge(LimitKind::RetainedBytes, maximum_compressed_bytes)?;
        self.charge(LimitKind::Allocations, 2)
    }

    fn charge_reassembly(
        &mut self,
        requirements: ReassemblyExecutionRequirements,
    ) -> Result<(), Error> {
        self.charge(LimitKind::OutputBytes, requirements.output_bytes())?;
        self.charge(LimitKind::TotalBytes, requirements.output_bytes())?;
        self.charge(LimitKind::Entries, requirements.offset_count())?;
        self.charge(LimitKind::ScratchBytes, requirements.scratch_bytes())?;
        self.charge(LimitKind::RetainedBytes, requirements.retained_bytes())?;
        self.charge(LimitKind::Allocations, requirements.allocations())
    }
}

/// A finite resource governed while one Pages body-footnote graph is changed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package bytes.
    OutputBytes,
    /// ZIP entries, archive objects, messages, or graph records.
    Entries,
    /// Bytes in one package entry, archive object, or message.
    EntryBytes,
    /// Aggregate package or IWA payload bytes.
    TotalBytes,
    /// Retained semantic footnote text bytes.
    TextBytes,
    /// UTF-16 units in one body story or footnote storage.
    TextUnits,
    /// Strict protobuf bytes.
    WireBytes,
    /// Strict protobuf fields.
    WireFields,
    /// Strict protobuf nesting.
    WireNesting,
    /// Aggregate scan, rewrite, and verification work.
    WireWork,
    /// Graph or metadata references.
    References,
    /// Package-metadata component records.
    Components,
    /// Package-metadata registry additions or removals.
    RegistryChanges,
    /// Retained transaction bytes.
    RetainedBytes,
    /// Temporary transaction bytes.
    ScratchBytes,
    /// Fallible allocation events.
    Allocations,
}

impl fmt::Display for LimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
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

/// Failure while selecting, staging, or publishing a body-footnote lifecycle
/// transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Error {
    /// A well-formed request this route refuses; [`Refusal::kind`] classifies it
    /// (record 0351 §2.4, record 0361).
    #[error(transparent)]
    Refused(Refusal),
    /// The selected graph is malformed or ambiguous.
    #[error("the Pages body-footnote graph cannot be edited safely")]
    InvalidSource,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages body-footnote {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    Limit {
        /// Resource category.
        kind: LimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed.
    #[error("could not allocate {amount} units for the Pages body-footnote transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Candidate semantic/locality verification failed.
    #[error("the edited Pages body-footnote graph failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Pages body-footnote patch does not match the exact source package")]
    PatchConflict,
}

/// A well-formed request this family's route refuses, classified by
/// [`Refusal::kind`] (record 0351 §2.4, record 0361).
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum Refusal {
    /// No footnote matched the selector.
    #[error("the Pages body-footnote selector did not match a footnote")]
    NotFound,
    /// The requested insertion position is not a valid UTF-16 boundary.
    #[error("the Pages body-footnote position is out of bounds")]
    PositionOutOfBounds,
    /// A footnote already occupies the requested position.
    #[error("a Pages body footnote already occupies the requested position")]
    PositionOccupied,
    /// Authored text contains a native structural marker.
    #[error("Pages body-footnote text cannot contain native structural markers")]
    StructuralMarker,
    /// Authored text exceeds its semantic byte ceiling.
    #[error("Pages body-footnote text exceeds its semantic byte budget")]
    TextTooLarge,
    /// A custom marker exceeds its semantic byte ceiling.
    #[error("Pages body-footnote custom marker exceeds its semantic byte budget")]
    CustomMarkTooLarge,
    /// The package has no exact physical source suitable for lifecycle edits.
    #[error("this Pages source does not support physical body-footnote edits")]
    UnsupportedSource,
    /// A valid but unsupported dependency owns part of the graph.
    #[error("the Pages body-footnote graph has an unsupported dependency")]
    UnsupportedDependency,
}

impl Refusal {
    /// The shared refusal kind litchi-ops classifies this reason by.
    #[must_use]
    pub const fn kind(&self) -> litchi_core::edit::RefusalKind {
        use litchi_core::edit::RefusalKind;
        match self {
            Self::NotFound { .. } => RefusalKind::NotFound,
            Self::PositionOutOfBounds { .. }
            | Self::PositionOccupied { .. }
            | Self::TextTooLarge { .. }
            | Self::CustomMarkTooLarge { .. } => RefusalKind::InvalidValue,
            Self::StructuralMarker { .. }
            | Self::UnsupportedSource { .. }
            | Self::UnsupportedDependency { .. } => RefusalKind::UnsupportedContent,
        }
    }
}

impl From<Refusal> for Error {
    fn from(refusal: Refusal) -> Self {
        Self::Refused(refusal)
    }
}

/// One mutable body-footnote lifecycle edit staged against an immutable
/// package snapshot.
pub struct Edit<'a> {
    source: &'a Package,
    selector: Selector,
    resolved: std::sync::OnceLock<ResolvedSource>,
    after: Option<Arc<Footnote>>,
    staged: bool,
}

struct ResolvedSource {
    position: Position,
    before: Arc<Footnote>,
}

/// Immutable semantic body-footnote values projected from one Pages package.
///
/// The snapshot owns only checked positions, user-visible text, and optional
/// custom markers.  Native attachment, storage, marker, and package-object
/// identities are deliberately not representable here.  Cloning a snapshot
/// shares its bounded values through an `Arc`; callers can therefore resolve
/// several selectors without repeatedly decoding the native graph or copying
/// every footnote.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    values: Arc<[Footnote]>,
}

impl Snapshot {
    fn from_values(values: Vec<Footnote>) -> Self {
        Self {
            values: Arc::from(values.into_boxed_slice()),
        }
    }

    /// Return the number of rooted body footnotes in source order.
    #[must_use]
    pub fn len(&self) -> usize {
        self.values.len()
    }

    /// Return whether the rooted body has no footnotes.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.values.is_empty()
    }

    /// Borrow all semantic footnotes in their native source order.
    #[must_use]
    pub fn as_slice(&self) -> &[Footnote] {
        &self.values
    }

    /// Borrow one footnote by zero-based source order.
    #[must_use]
    pub fn get(&self, index: usize) -> Option<&Footnote> {
        self.values.get(index)
    }

    /// Resolve one position- or source-order selector without exposing native
    /// object identifiers.
    #[must_use]
    pub fn select<S>(&self, selector: S) -> Option<&Footnote>
    where
        S: Into<Selector>,
    {
        match selector.into() {
            Selector::Position(position) => self.get(position.get()),
            Selector::Anchor(position) => {
                self.values.iter().find(|value| value.position == position)
            },
        }
    }

    /// Iterate over semantic footnotes in native source order.
    pub fn iter(&self) -> std::slice::Iter<'_, Footnote> {
        self.values.iter()
    }

    fn shared_values(&self) -> Arc<[Footnote]> {
        Arc::clone(&self.values)
    }
}

impl<'a> IntoIterator for &'a Snapshot {
    type Item = &'a Footnote;
    type IntoIter = std::slice::Iter<'a, Footnote>;

    fn into_iter(self) -> Self::IntoIter {
        self.iter()
    }
}

impl fmt::Debug for Edit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Edit")
            .field(
                "resolved",
                &self.resolved.get().map(|source| source.position),
            )
            .finish_non_exhaustive()
    }
}

impl<'a> Edit<'a> {
    fn resolve_source(&self) -> Result<ResolvedSource, Error> {
        let snapshot = self
            .source
            .body_footnote_snapshot()
            .map_err(|_| Error::InvalidSource)?;
        let before = snapshot
            .select(self.selector)
            .ok_or(Error::Refused(Refusal::NotFound))?;
        Ok(ResolvedSource {
            position: before.position,
            before: Arc::new(before.clone()),
        })
    }

    fn resolve(&mut self) -> Result<(), Error> {
        if self.resolved.get().is_none() {
            self.resolved
                .set(self.resolve_source()?)
                .map_err(|_| Error::InvalidSource)?;
        }
        if !self.staged && self.after.is_none() {
            let resolved = self.resolved.get().ok_or(Error::InvalidSource)?;
            self.after = Some(Arc::clone(&resolved.before));
        }
        Ok(())
    }

    fn new(source: &'a Package, selector: Selector) -> Result<Self, Error> {
        Ok(Self {
            source,
            selector,
            resolved: std::sync::OnceLock::new(),
            after: None,
            staged: false,
        })
    }

    /// Return the checked UTF-16 anchor position, resolving on first access.
    pub fn position(&mut self) -> Result<Position, Error> {
        self.resolve()?;
        self.resolved
            .get()
            .map(|resolved| resolved.position)
            .ok_or(Error::InvalidSource)
    }

    /// Borrow the original semantic footnote.
    pub fn before(&mut self) -> Result<&Footnote, Error> {
        self.resolve()?;
        self.resolved
            .get()
            .map(|resolved| resolved.before.as_ref())
            .ok_or(Error::InvalidSource)
    }

    /// Borrow the staged semantic result. `None` represents graph removal.
    pub fn after(&mut self) -> Result<Option<&Footnote>, Error> {
        self.resolve()?;
        Ok(self.after.as_deref())
    }

    /// Stage replacement of the selected footnote text.
    pub fn set(&mut self, text: &str) -> Result<&mut Self, Error> {
        self.resolve()?;
        validate_authored(text, MAX_TEXT_BYTES, Error::Refused(Refusal::TextTooLarge))?;
        let Some(after) = self.after.as_mut() else {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        };
        Arc::make_mut(after).text = try_boxed(text)?;
        self.staged = true;
        Ok(self)
    }

    /// Stage replacement/removal of the optional custom marker.
    pub fn set_custom_mark(&mut self, custom_mark: Option<&str>) -> Result<&mut Self, Error> {
        self.resolve()?;
        if let Some(value) = custom_mark {
            validate_authored(
                value,
                MAX_CUSTOM_MARK_BYTES,
                Error::Refused(Refusal::CustomMarkTooLarge),
            )?;
        }
        let Some(after) = self.after.as_mut() else {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        };
        Arc::make_mut(after).custom_mark = custom_mark.map(try_boxed).transpose()?;
        self.staged = true;
        Ok(self)
    }

    /// Stage complete removal of the selected body-footnote graph.
    pub fn clear(&mut self) -> &mut Self {
        self.after = None;
        self.staged = true;
        self
    }

    /// Validate and atomically publish the staged lifecycle change.
    pub fn commit(&self) -> Result<Commit, Error> {
        let resolved = match self.resolved.get() {
            Some(resolved) => resolved,
            None => {
                let resolved = self.resolve_source()?;
                return commit_resolved(
                    self.source,
                    &resolved,
                    if self.staged {
                        self.after.as_ref()
                    } else {
                        Some(&resolved.before)
                    },
                );
            },
        };
        let after = if self.staged {
            self.after.as_ref()
        } else {
            Some(&resolved.before)
        };
        commit_resolved(self.source, resolved, after)
    }
}

fn commit_resolved(
    source: &Package,
    resolved: &ResolvedSource,
    after: Option<&Arc<Footnote>>,
) -> Result<Commit, Error> {
    let position = resolved.position;
    let before = resolved.before.as_ref();
    source.validate().map_err(map_package_error)?;
    let current = current_footnote(source, position)?;
    if &current != before {
        return Err(Error::PatchConflict);
    }
    let Some(after) = after else {
        return publish_rewrite(source, position, None, Some(&resolved.before));
    };
    let after_arc = after;
    let after = after_arc.as_ref();
    if after == before {
        return noop_commit(source, position, Some(&resolved.before));
    }
    let mut text_edit = source
        .edit_body_footnote_text(Selector::Anchor(position))
        .map_err(map_text_error)?;
    text_edit.set(&after.text).map_err(map_text_error)?;
    text_edit
        .set_custom_mark(after.custom_mark.as_deref())
        .map_err(map_text_error)?;
    let text_commit = text_edit.commit().map_err(map_text_error)?;
    let source_bytes: Arc<[u8]> = source.state.source.shared_source();
    let target_bytes = text_commit.package().state.source.shared_source();
    let source_fingerprint = super::section_transaction::fingerprint(&source_bytes);
    let target_fingerprint = super::section_transaction::fingerprint(&target_bytes);
    let target_sequence = footnote_sequence(text_commit.package())?;
    let package = text_commit.into_package();
    Ok(Commit {
        package,
        patch: Patch {
            source_bytes,
            target_bytes,
            source_sequence: footnote_sequence(source)?,
            target_sequence,
            source_fingerprint,
            target_fingerprint,
            position,
            before: Some(Arc::clone(&resolved.before)),
            after: Some(Arc::clone(after_arc)),
        },
        diagnostics: Diagnostics::published(1),
    })
}

/// An exact-source-checked reversible body-footnote lifecycle patch.
#[derive(Clone, PartialEq, Eq)]
pub struct Patch {
    pub(super) source_bytes: Arc<[u8]>,
    pub(super) target_bytes: Arc<[u8]>,
    source_sequence: Arc<[Footnote]>,
    target_sequence: Arc<[Footnote]>,
    pub(super) source_fingerprint: u64,
    pub(super) target_fingerprint: u64,
    position: Position,
    before: Option<Arc<Footnote>>,
    after: Option<Arc<Footnote>>,
}

impl fmt::Debug for Patch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Patch")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl Patch {
    /// Return the UTF-16 anchor position associated with the operation.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Borrow the optional semantic source value.
    #[must_use]
    pub fn before(&self) -> Option<&Footnote> {
        self.before.as_deref()
    }

    /// Borrow the optional semantic target value.
    #[must_use]
    pub fn after(&self) -> Option<&Footnote> {
        self.after.as_deref()
    }

    /// Return the compact source fingerprint used for diagnostics.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.source_fingerprint
    }

    /// Return the compact target fingerprint used for diagnostics.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.target_fingerprint
    }

    /// Return whether semantic state and exact bytes are unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after
            && self.source_fingerprint == self.target_fingerprint
            && self.source_bytes.as_ref() == self.target_bytes.as_ref()
    }

    /// Return the exact inverse patch.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source_bytes: Arc::clone(&self.target_bytes),
            target_bytes: Arc::clone(&self.source_bytes),
            source_sequence: Arc::clone(&self.target_sequence),
            target_sequence: Arc::clone(&self.source_sequence),
            source_fingerprint: self.target_fingerprint,
            target_fingerprint: self.source_fingerprint,
            position: self.position,
            before: self.after.as_ref().map(Arc::clone),
            after: self.before.as_ref().map(Arc::clone),
        }
    }
}

/// Compact evidence for one body-footnote lifecycle commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Diagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl Diagnostics {
    pub(super) const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(touched_components: usize) -> Self {
        Self {
            changed: true,
            touched_components,
            full_reparse_performed: true,
        }
    }

    /// Return whether exact package bytes changed.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of semantic native components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The verified result of one immutable body-footnote lifecycle transaction.
#[must_use = "a Pages body-footnote commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct Commit {
    package: Package,
    patch: Patch,
    diagnostics: Diagnostics,
}

impl Commit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume the commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Consume the commit and return its package and patch.
    #[must_use]
    pub fn into_parts(self) -> (Package, Patch) {
        (self.package, self.patch)
    }

    /// Borrow the reversible exact-source patch.
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
    /// Read the rooted body-footnote graph into one immutable semantic
    /// snapshot.
    ///
    /// The snapshot is selector-first and archive-free: callers can resolve
    /// multiple positions or source-order indexes without retaining native
    /// object identifiers or re-decoding the graph.  Physical source and
    /// semantic text limits retained by this package are enforced before the
    /// snapshot is published.
    ///
    /// # Errors
    ///
    /// Returns [`super::PackageError::InvalidFormat`] when the rooted graph or
    /// one of its strict footnote payloads is malformed, and a bounded
    /// semantic error when the projected values exceed the package limits.
    pub fn body_footnote_snapshot(&self) -> super::PackageResult<Snapshot> {
        super::project_body_footnotes(
            self.state.source.components(),
            self.state.source.limits(),
            super::effective_text_limit(self.state.source.limits()),
        )
        .map(Snapshot::from_values)
    }

    /// Stage a selector-first edit of one existing body footnote.
    pub fn edit_body_footnote<S>(&self, selector: S) -> Result<Edit<'_>, Error>
    where
        S: Into<Selector>,
    {
        Edit::new(self, selector.into())
    }

    /// Insert one body footnote at a checked UTF-16 position.
    pub fn insert_body_footnote(
        &self,
        position: Position,
        text: impl AsRef<str>,
        custom_mark: Option<&str>,
    ) -> Result<Commit, Error> {
        validate_authored(
            text.as_ref(),
            MAX_TEXT_BYTES,
            Error::Refused(Refusal::TextTooLarge),
        )?;
        if let Some(value) = custom_mark {
            validate_authored(
                value,
                MAX_CUSTOM_MARK_BYTES,
                Error::Refused(Refusal::CustomMarkTooLarge),
            )?;
        }
        insert_lifecycle(self, position, text.as_ref(), custom_mark)
    }

    /// Apply an exact-source-checked body-footnote lifecycle patch.
    pub fn apply_body_footnote(&self, patch: &Patch) -> Result<Commit, Error> {
        let source = &self.state.source;
        if super::section_transaction::fingerprint(source.source_bytes())
            != patch.source_fingerprint
            || source.source_bytes() != patch.source_bytes.as_ref()
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
        self.validate().map_err(map_package_error)?;
        let current_sequence = footnote_sequence(self)?;
        if current_sequence.as_ref() != patch.source_sequence.as_ref() {
            return Err(Error::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(Commit {
                package: self.clone(),
                patch: patch.clone(),
                diagnostics: Diagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact()
            || super::section_transaction::fingerprint(&patch.target_bytes)
                != patch.target_fingerprint
        {
            return Err(Error::PatchConflict);
        }
        let mut budget = TransactionBudget::new(self)?;
        budget.charge(LimitKind::OutputBytes, patch.target_bytes.len())?;
        budget.charge(LimitKind::RetainedBytes, patch.target_bytes.len())?;
        budget.charge(LimitKind::Allocations, 1)?;
        let candidate_source = self
            .state
            .source
            .reopen(Arc::clone(&patch.target_bytes))
            .map_err(map_archive_error)?;
        let candidate =
            Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
        let target_sequence = footnote_sequence(&candidate)?;
        if target_sequence.as_ref() != patch.target_sequence.as_ref() {
            return Err(Error::Verification);
        }
        Ok(Commit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: Diagnostics::published(1),
        })
    }
}

fn validate_authored(value: &str, maximum: usize, overflow: Error) -> Result<(), Error> {
    if value.len() > maximum {
        return Err(overflow);
    }
    if value.contains(['\u{000e}', '\u{fffc}']) {
        return Err(Error::Refused(Refusal::StructuralMarker));
    }
    Ok(())
}

fn try_boxed(value: &str) -> Result<Box<str>, Error> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_error| Error::Allocation {
            amount: value.len(),
        })?;
    owned.push_str(value);
    Ok(owned.into_boxed_str())
}

fn current_footnote(package: &Package, position: Position) -> Result<Footnote, Error> {
    current_footnote_at_optional(package, position)?.ok_or(Error::Refused(Refusal::NotFound))
}

fn current_footnote_at_optional(
    package: &Package,
    position: Position,
) -> Result<Option<Footnote>, Error> {
    package
        .body_footnote_snapshot()
        .map_err(map_package_error)
        .map(|values| values.select(Selector::Anchor(position)).cloned())
}

fn footnote_sequence(package: &Package) -> Result<Arc<[Footnote]>, Error> {
    package
        .body_footnote_snapshot()
        .map_err(map_package_error)
        .map(|values| values.shared_values())
}

fn noop_commit(
    package: &Package,
    position: Position,
    value: Option<&Arc<Footnote>>,
) -> Result<Commit, Error> {
    let source: Arc<[u8]> = package.state.source.shared_source();
    let fingerprint = super::section_transaction::fingerprint(&source);
    Ok(Commit {
        package: package.clone(),
        patch: Patch {
            source_bytes: Arc::clone(&source),
            target_bytes: source,
            source_sequence: footnote_sequence(package)?,
            target_sequence: footnote_sequence(package)?,
            source_fingerprint: fingerprint,
            target_fingerprint: fingerprint,
            position,
            before: value.map(Arc::clone),
            after: value.map(Arc::clone),
        },
        diagnostics: Diagnostics::unchanged(),
    })
}

fn publish_rewrite(
    source: &Package,
    position: Position,
    after: Option<&Arc<Footnote>>,
    before: Option<&Arc<Footnote>>,
) -> Result<Commit, Error> {
    let target = rewrite_remove(source, position)?;
    let source_bytes: Arc<[u8]> = source.state.source.shared_source();
    let target_bytes = target.state.source.shared_source();
    let target_sequence = footnote_sequence(&target)?;
    Ok(Commit {
        package: target,
        patch: Patch {
            source_fingerprint: super::section_transaction::fingerprint(&source_bytes),
            target_fingerprint: super::section_transaction::fingerprint(&target_bytes),
            source_bytes,
            target_bytes,
            source_sequence: footnote_sequence(source)?,
            target_sequence,
            position,
            before: before.map(Arc::clone),
            after: after.map(Arc::clone),
        },
        diagnostics: Diagnostics::published(1),
    })
}

fn map_text_error(error: super::footnote_text::Error) -> Error {
    match error {
        super::footnote_text::Error::Refused(crate::footnote::text::Refusal::NotFound) => {
            Error::Refused(Refusal::NotFound)
        },
        super::footnote_text::Error::Refused(crate::footnote::text::Refusal::StructuralMarker) => {
            Error::Refused(Refusal::StructuralMarker)
        },
        super::footnote_text::Error::Refused(crate::footnote::text::Refusal::TextTooLarge) => {
            Error::Refused(Refusal::TextTooLarge)
        },
        super::footnote_text::Error::Refused(
            crate::footnote::text::Refusal::CustomMarkTooLarge,
        ) => Error::Refused(Refusal::CustomMarkTooLarge),
        super::footnote_text::Error::Refused(crate::footnote::text::Refusal::UnsupportedSource) => {
            Error::Refused(Refusal::UnsupportedSource)
        },
        super::footnote_text::Error::InvalidSource => Error::InvalidSource,
        super::footnote_text::Error::Verification => Error::Verification,
        super::footnote_text::Error::PatchConflict => Error::PatchConflict,
        super::footnote_text::Error::Allocation { amount } => Error::Allocation { amount },
        super::footnote_text::Error::Limit {
            kind,
            observed,
            maximum,
        } => Error::Limit {
            kind: match kind {
                super::footnote_text::LimitKind::InputBytes => LimitKind::InputBytes,
                super::footnote_text::LimitKind::OutputBytes => LimitKind::OutputBytes,
                super::footnote_text::LimitKind::Entries => LimitKind::Entries,
                super::footnote_text::LimitKind::EntryBytes => LimitKind::EntryBytes,
                super::footnote_text::LimitKind::TotalBytes => LimitKind::TotalBytes,
                super::footnote_text::LimitKind::TextBytes => LimitKind::TextBytes,
                super::footnote_text::LimitKind::TextUnits => LimitKind::TextUnits,
                super::footnote_text::LimitKind::WireBytes => LimitKind::WireBytes,
                super::footnote_text::LimitKind::WireFields => LimitKind::WireFields,
                super::footnote_text::LimitKind::WireNesting => LimitKind::WireNesting,
                super::footnote_text::LimitKind::WireWork => LimitKind::WireWork,
            },
            observed,
            maximum,
        },
    }
}

fn map_package_error(error: super::PackageError) -> Error {
    match error {
        super::PackageError::Archive(error) => map_archive_error(error),
        super::PackageError::Allocation { amount } => Error::Allocation { amount },
        super::PackageError::PayloadLimit { observed, limit }
        | super::PackageError::ObjectLimit { observed, limit }
        | super::PackageError::SectionNamesTooLarge { observed, limit } => Error::Limit {
            kind: LimitKind::Entries,
            observed: observed as u64,
            maximum: limit as u64,
        },
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
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => LimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => LimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => Error::Allocation { amount },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
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
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => LimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => LimitKind::EntryBytes,
                litchi_iwa_core::LimitKind::HeaderFields => LimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => LimitKind::WireNesting,
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

fn map_graph_error(error: graph_codec::DecodeError) -> Error {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            graph_codec::DecodeLimit::InputBytes { observed, maximum } => {
                (LimitKind::InputBytes, observed, maximum)
            },
            graph_codec::DecodeLimit::OutputBytes { observed, maximum } => {
                (LimitKind::OutputBytes, observed, maximum)
            },
            graph_codec::DecodeLimit::Fields { observed, maximum } => {
                (LimitKind::WireFields, observed, maximum)
            },
            graph_codec::DecodeLimit::WorkBytes { observed, maximum } => {
                (LimitKind::WireWork, observed, maximum)
            },
            graph_codec::DecodeLimit::Nesting { observed, maximum } => {
                return Error::Limit {
                    kind: LimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                };
            },
            graph_codec::DecodeLimit::Entries { observed, maximum } => {
                (LimitKind::Entries, observed, maximum)
            },
            graph_codec::DecodeLimit::TextBytes { observed, maximum } => {
                (LimitKind::TextBytes, observed, maximum)
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

fn map_text_wire_error(error: litchi_iwa_text_wire::RewriteError) -> Error {
    match error {
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => Error::Limit {
            kind: match resource {
                "output bytes" => LimitKind::OutputBytes,
                "fields" => LimitKind::WireFields,
                "rewrite work" => LimitKind::WireWork,
                "text bytes" => LimitKind::TextBytes,
                "text units" => LimitKind::TextUnits,
                _ => LimitKind::WireBytes,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_text_wire::RewriteError::Allocation { amount, .. } => {
            Error::Allocation { amount }
        },
        litchi_iwa_text_wire::RewriteError::RangeOutOfBounds { .. }
        | litchi_iwa_text_wire::RewriteError::SurrogateSplit { .. } => {
            Error::Refused(Refusal::PositionOutOfBounds)
        },
        litchi_iwa_text_wire::RewriteError::ReversedRange { .. }
        | litchi_iwa_text_wire::RewriteError::ArithmeticOverflow { .. }
        | litchi_iwa_text_wire::RewriteError::InvalidLimit { .. }
        | litchi_iwa_text_wire::RewriteError::InvalidFormat(_)
        | litchi_iwa_text_wire::RewriteError::Projection(_) => Error::InvalidSource,
        _ => Error::InvalidSource,
    }
}

fn map_common_wire_error(error: litchi_iwa_common::Error) -> Error {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => Error::Limit {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes
                | litchi_iwa_common::LimitKind::OutputBytes => LimitKind::WireBytes,
                litchi_iwa_common::LimitKind::Fields => LimitKind::WireFields,
                litchi_iwa_common::LimitKind::Nesting => LimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => LimitKind::WireWork,
                _ => LimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => Error::Allocation { amount },
        _ => Error::InvalidSource,
    }
}

fn map_storage_rewrite_limits_error(error: super::StorageWireLimitsError) -> Error {
    match error {
        super::StorageWireLimitsError::Physical(error) => map_archive_error(error),
        super::StorageWireLimitsError::Wire(litchi_iwa_text_wire::RewriteError::InvalidLimit {
            field: "nesting",
            value,
            ..
        }) => Error::Limit {
            kind: LimitKind::WireNesting,
            observed: 4,
            maximum: value as u64,
        },
        super::StorageWireLimitsError::Wire(error) => map_text_wire_error(error),
    }
}

fn storage_wire_limits(
    physical: super::Limits,
    input_bytes: usize,
    output_bytes: usize,
) -> Result<WireLimits, Error> {
    let archive = physical
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let input_bytes = input_bytes.max(1);
    let input_limit = archive
        .max_message_bytes()
        .clamp(1, WireLimits::MAX_INPUT_BYTES);
    if input_bytes > input_limit {
        return Err(Error::Limit {
            kind: LimitKind::WireBytes,
            observed: input_bytes as u64,
            maximum: input_limit as u64,
        });
    }
    let output_bytes = output_bytes.max(1);
    let output_limit = archive
        .max_message_bytes()
        .clamp(1, WireLimits::MAX_OUTPUT_BYTES);
    if output_bytes > output_limit {
        return Err(Error::Limit {
            kind: LimitKind::WireBytes,
            observed: output_bytes as u64,
            maximum: output_limit as u64,
        });
    }
    WireLimits::default()
        .with_input_bytes(input_bytes)
        .and_then(|limits| limits.with_output_bytes(output_bytes))
        .and_then(|limits| {
            limits.with_fields(archive.max_header_fields().min(WireLimits::MAX_FIELDS))
        })
        .and_then(|limits| {
            limits.with_nesting(archive.max_header_nesting().min(WireLimits::MAX_NESTING))
        })
        .and_then(|limits| limits.with_rewrite_work(WireLimits::MAX_REWRITE_WORK))
        .map_err(map_common_wire_error)
}

#[derive(Debug, Clone)]
struct MetadataComponent {
    identifier: u64,
    preferred: String,
    locator: Option<String>,
    current: bool,
}

impl MetadataComponent {
    fn effective(&self) -> &str {
        self.locator.as_deref().unwrap_or(&self.preferred)
    }
}

#[derive(Debug, Clone, Copy)]
struct MetadataUuid {
    component: u64,
    object: u64,
    uuid: UuidBits,
    current: bool,
}

#[derive(Debug, Clone, Copy)]
struct MetadataExternal {
    source: u64,
    target: u64,
    object: Option<u64>,
    weak: Option<bool>,
    current: bool,
}

#[derive(Debug, Clone, Copy)]
struct MetadataDataOwner {
    component: u64,
    data: u64,
    object: u64,
    count: u32,
    current: bool,
    unknown: bool,
}

#[derive(Default)]
struct MetadataFactsVisitor {
    components: Vec<MetadataComponent>,
    uuids: Vec<MetadataUuid>,
    externals: Vec<MetadataExternal>,
    data_owners: Vec<MetadataDataOwner>,
    ambiguous: Vec<(u64, u64)>,
    data_metadata_maps: Vec<(u64, bool)>,
}

fn metadata_try_reserve<T>(values: &mut Vec<T>) -> Result<(), MetadataRewriteError> {
    let requested = values
        .len()
        .checked_add(1)
        .ok_or_else(|| MetadataRewriteError::allocation(usize::MAX))?;
    values
        .try_reserve(1)
        .map_err(|_| MetadataRewriteError::allocation(requested))?;
    Ok(())
}

fn metadata_try_push<T>(values: &mut Vec<T>, value: T) -> Result<(), MetadataRewriteError> {
    metadata_try_reserve(values)?;
    values.push(value);
    Ok(())
}

fn metadata_try_string(value: &str) -> Result<String, MetadataRewriteError> {
    let mut output = String::new();
    output
        .try_reserve_exact(value.len())
        .map_err(|_| MetadataRewriteError::allocation(value.len()))?;
    output.push_str(value);
    Ok(output)
}

impl PackageMetadataVisitor for MetadataFactsVisitor {
    fn visit_component(
        &mut self,
        component: package_metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        metadata_try_reserve(&mut self.components)?;
        let preferred = metadata_try_string(component.preferred_locator())?;
        let locator = component.locator().map(metadata_try_string).transpose()?;
        self.components.push(MetadataComponent {
            identifier: component.identifier(),
            preferred,
            locator,
            current: component.is_current(),
        });
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        metadata_try_push(
            &mut self.uuids,
            MetadataUuid {
                component: binding.component().identifier(),
                object: binding.object_identifier(),
                uuid: binding.uuid(),
                current: binding.component().is_current(),
            },
        )
    }

    fn visit_external_reference(
        &mut self,
        reference: package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        metadata_try_push(
            &mut self.externals,
            MetadataExternal {
                source: reference.source().identifier(),
                target: reference.target_component_identifier(),
                object: reference.object_identifier(),
                weak: reference.is_weak(),
                current: reference.source().is_current() && !reference.is_versioned(),
            },
        )
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        metadata_try_push(
            &mut self.data_owners,
            MetadataDataOwner {
                component: owner.component().identifier(),
                data: owner.data_identifier(),
                object: owner.object_identifier(),
                count: owner.count(),
                current: owner.component().is_current(),
                unknown: owner.has_unknown_fields(),
            },
        )
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        component: package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), MetadataRewriteError> {
        metadata_try_push(&mut self.ambiguous, (component.identifier(), identifier))
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        has_unknown_fields: bool,
    ) -> Result<(), MetadataRewriteError> {
        metadata_try_push(
            &mut self.data_metadata_maps,
            (object_identifier, has_unknown_fields),
        )
    }
}

struct MetadataLocation<'source> {
    member: &'source str,
    archive: &'source Archive,
    object_index: usize,
    message_index: usize,
    payload: &'source [u8],
}

struct MetadataRewrite {
    member: String,
    compressed: Vec<u8>,
}

fn metadata_location<'source>(
    source: &'source Package,
) -> Result<MetadataLocation<'source>, Error> {
    let component = source
        .state
        .source
        .components()
        .get("Index/Metadata.iwa")
        .ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    let archive = component.archive();
    let mut location = None;
    for (object_index, object) in archive.objects.iter().enumerate() {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == METADATA_MESSAGE_TYPE {
                if location.replace((object_index, message_index)).is_some() {
                    return Err(Error::InvalidSource);
                }
            }
        }
    }
    let (object_index, message_index) = location.ok_or(Error::InvalidSource)?;
    Ok(MetadataLocation {
        member: component.name(),
        payload: archive.objects[object_index].messages[message_index]
            .data
            .as_slice(),
        archive,
        object_index,
        message_index,
    })
}

fn metadata_options(source: &Package, payload_len: usize) -> Result<MetadataRewriteOptions, Error> {
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let input = payload_len.max(1);
    let maximum = source
        .state
        .source
        .limits()
        .max_iwa_stream_bytes()
        .max(input);
    let work = input.checked_mul(128).ok_or(Error::InvalidSource)?.max(1);
    Ok(MetadataRewriteOptions::new(
        input,
        maximum,
        archive_limits.max_header_fields().max(input.min(1_000_000)),
        work,
        u32::try_from(archive_limits.max_header_nesting()).unwrap_or(u32::MAX),
        input.max(1),
        input.max(1),
        64,
    ))
}

fn metadata_facts(
    source: &Package,
    location: &MetadataLocation<'_>,
    budget: &mut TransactionBudget,
) -> Result<(MetadataFactsVisitor, u64), Error> {
    let options = metadata_options(source, location.payload.len())?;
    let mut visitor = MetadataFactsVisitor::default();
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        location.payload,
        options,
        &mut visitor,
    )
    .map_err(map_metadata_error)?;
    budget.charge_metadata_prepare(inspection.report())?;
    Ok((visitor, inspection.last_object_identifier()))
}

fn map_metadata_error(error: MetadataRewriteError) -> Error {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            package_metadata_codec::RewriteLimit::InputBytes { observed, maximum } => {
                (LimitKind::InputBytes, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => {
                (LimitKind::OutputBytes, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Fields { observed, maximum } => {
                (LimitKind::WireFields, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (LimitKind::WireWork, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => {
                return Error::Limit {
                    kind: LimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                };
            },
            package_metadata_codec::RewriteLimit::Components { observed, maximum } => {
                (LimitKind::Components, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::References { observed, maximum } => {
                (LimitKind::References, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Additions { observed, maximum } => {
                (LimitKind::RegistryChanges, observed, maximum)
            },
            _ => return Error::InvalidSource,
        };
        return Error::Limit {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(amount) = error.allocation_request() {
        return Error::Allocation { amount };
    }
    match error.invalid_reason() {
        Some(package_metadata_codec::InvalidReason::VersionedRemoval)
        | Some(package_metadata_codec::InvalidReason::CrossComponentRemoval)
        | Some(package_metadata_codec::InvalidReason::ExistingReferenceCollision) => {
            Error::Refused(Refusal::UnsupportedDependency)
        },
        _ => Error::InvalidSource,
    }
}

fn component_for_member<'a>(
    facts: &'a MetadataFactsVisitor,
    member: &str,
) -> Result<&'a MetadataComponent, Error> {
    let base = member
        .rsplit('/')
        .next()
        .and_then(|value| value.strip_suffix(".iwa"))
        .ok_or(Error::InvalidSource)?;
    let mut found = None;
    for component in facts
        .components
        .iter()
        .filter(|component| component.current)
    {
        if component.effective() == base || component.preferred == base {
            if found.is_some() {
                return Err(Error::InvalidSource);
            }
            found = Some(component);
        }
    }
    found.ok_or(Error::Refused(Refusal::UnsupportedDependency))
}

fn current_metadata_component(
    facts: &MetadataFactsVisitor,
    identifier: u64,
) -> Result<&MetadataComponent, Error> {
    let mut found = None;
    for component in facts
        .components
        .iter()
        .filter(|component| component.current && component.identifier == identifier)
    {
        if found.replace(component).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    found.ok_or(Error::InvalidSource)
}

fn rewrite_metadata_archive(
    source: &Package,
    location: MetadataLocation<'_>,
    payload: Vec<u8>,
    budget: &mut TransactionBudget,
) -> Result<MetadataRewrite, Error> {
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
    // The strict metadata pass above is complete and charged before this
    // copy-on-write archive clone. Read-only callers and rejected metadata
    // batches therefore retain only borrowed source data.
    let mut archive = location.archive.clone();
    let object_identifier = archive.objects[location.object_index]
        .archive_info
        .identifier
        .ok_or(Error::InvalidSource)?;
    let object = archive
        .object_mut(object_identifier)
        .ok_or(Error::InvalidSource)?;
    object
        .replace_message_preserving_header(
            location.message_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: payload,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_length = archive
        .encoded_len(archive_limits)
        .map_err(map_core_error)?;
    let maximum_compressed = SnappyStream::maximum_compressed_len(encoded_length, snappy_limits)
        .map_err(map_core_error)?;
    budget.charge_archive_output(encoded_length, maximum_compressed)?;
    let bytes = archive.to_bytes(archive_limits).map_err(map_core_error)?;
    if bytes.len() != encoded_length {
        return Err(Error::Verification);
    }
    let compressed = SnappyStream::compress(&bytes, snappy_limits).map_err(map_core_error)?;
    if compressed.len() > maximum_compressed {
        return Err(Error::Verification);
    }
    Ok(MetadataRewrite {
        member: location.member.to_owned(),
        compressed,
    })
}

fn metadata_addition(
    source: &Package,
    body_member: &str,
    storage_identifier: u64,
    marker_identifier: u64,
    budget: &mut TransactionBudget,
) -> Result<MetadataRewrite, Error> {
    let location = metadata_location(source)?;
    let (facts, last_identifier) = metadata_facts(source, &location, budget)?;
    if marker_identifier <= last_identifier {
        return Err(Error::InvalidSource);
    }
    let body = component_for_member(&facts, body_member)?;
    let body_selector = ComponentSelector::new(body.identifier, body.effective());

    let mut view_state = None;
    for component in facts
        .components
        .iter()
        .filter(|component| component.current)
    {
        if component.effective() == "ViewState" || component.preferred == "ViewState" {
            if view_state.replace(component).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    let mut uuid = UuidBits::new(
        storage_identifier ^ 0x9e37_79b9_7f4a_7c15,
        marker_identifier ^ 0xd1b5_4a32_d192_ed03,
    );
    while facts.uuids.iter().any(|binding| binding.uuid == uuid) {
        uuid = UuidBits::new(
            uuid.lower().checked_add(1).ok_or(Error::InvalidSource)?,
            uuid.upper()
                .checked_add(0x9e37_79b9_7f4a_7c15)
                .ok_or(Error::InvalidSource)?,
        );
        if facts.uuids.iter().any(|binding| binding.uuid == uuid) {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
    }
    let storage_uuid = ObjectUuidAddition::new(body_selector, storage_identifier, uuid);
    let uuid_additions = vec![storage_uuid];
    let mut reference_additions = Vec::new();
    let mut token_selectors = Vec::new();
    token_selectors.push(body_selector);
    if let Some(view) = view_state {
        let view_selector = ComponentSelector::new(view.identifier, view.effective());
        reference_additions.push(ExternalReferenceAddition::new(
            view_selector,
            body_selector,
            storage_identifier,
            Some(true),
        ));
        token_selectors.push(view_selector);
    }

    let additions = package_metadata_codec::Batch::new(
        last_identifier,
        marker_identifier,
        &uuid_additions,
        &reference_additions,
    );
    let save_tokens = SaveTokenBatch::new(&token_selectors);
    let batch = package_metadata_codec::AdditionSaveTokenBatch::new(additions, save_tokens);
    let options = metadata_options(source, location.payload.len())?;
    let prepared = package_metadata_codec::prepare_package_metadata_additions_and_save_tokens(
        location.payload,
        batch,
        options,
    )
    .map_err(map_metadata_error)?;
    let prepare_report = prepared.prepare_report();
    budget.charge_metadata_prepare(prepare_report)?;
    let requirements = prepared.execution_requirements();
    budget.charge_metadata_requirements(requirements)?;
    let execution_limits = requirements.exact_limits();
    let output = prepared
        .execute(execution_limits)
        .map_err(map_metadata_error)?;
    if output.report().output_bytes() != requirements.output_bytes()
        || output.report().fields() != requirements.fields()
        || output.report().work_bytes() != requirements.work_bytes()
        || output.report().components_scanned() != requirements.components()
        || output.report().references_scanned() != requirements.references()
        || output.report().allocations() != requirements.allocations()
        || output.report().retained_bytes() != requirements.retained_bytes()
        || output.report().scratch_bytes() > requirements.scratch_bytes()
    {
        return Err(Error::Verification);
    }
    rewrite_metadata_archive(source, location, output.into_bytes(), budget)
}

fn metadata_removal(
    source: &Package,
    storage_identifier: u64,
    budget: &mut TransactionBudget,
) -> Result<MetadataRewrite, Error> {
    let location = metadata_location(source)?;
    let (facts, last_identifier) = metadata_facts(source, &location, budget)?;
    if facts
        .ambiguous
        .iter()
        .any(|(_, object)| *object == storage_identifier)
        || facts
            .data_metadata_maps
            .iter()
            .any(|(object, _)| *object == storage_identifier)
    {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    let mut matching_uuid_count = 0usize;
    let mut current_uuid = None;
    let mut has_versioned_uuid = false;
    for binding in facts
        .uuids
        .iter()
        .filter(|binding| binding.object == storage_identifier)
    {
        matching_uuid_count = matching_uuid_count
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
        if binding.current {
            if current_uuid.is_none() {
                current_uuid = Some(*binding);
            }
        } else {
            has_versioned_uuid = true;
        }
    }
    let current_uuid = current_uuid.ok_or(Error::Refused(Refusal::UnsupportedDependency))?;
    if has_versioned_uuid {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    if matching_uuid_count != 1 {
        return Err(Error::InvalidSource);
    }
    let target = current_metadata_component(&facts, current_uuid.component)?;
    let target_selector = ComponentSelector::new(target.identifier, target.effective());
    let uuid_removals = [ObjectUuidRemoval::new(
        target_selector,
        storage_identifier,
        current_uuid.uuid,
    )];

    let mut matching_external_count = 0usize;
    let mut has_versioned_external = false;
    for reference in facts.externals.iter().filter(|reference| {
        reference.target == current_uuid.component && reference.object == Some(storage_identifier)
    }) {
        matching_external_count = matching_external_count
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
        if !reference.current {
            has_versioned_external = true;
        }
    }
    if has_versioned_external {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    let mut matching_data_count = 0usize;
    let mut has_unsupported_data = false;
    for owner in facts
        .data_owners
        .iter()
        .filter(|owner| owner.object == storage_identifier)
    {
        matching_data_count = matching_data_count
            .checked_add(1)
            .ok_or(Error::InvalidSource)?;
        if !owner.current || owner.unknown {
            has_unsupported_data = true;
        }
    }
    if has_unsupported_data {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }

    let mut external_removals = Vec::new();
    external_removals
        .try_reserve_exact(matching_external_count)
        .map_err(|_| Error::Allocation {
            amount: matching_external_count,
        })?;
    let mut data_removals = Vec::new();
    data_removals
        .try_reserve_exact(matching_data_count)
        .map_err(|_| Error::Allocation {
            amount: matching_data_count,
        })?;
    let token_capacity = matching_external_count
        .checked_add(matching_data_count)
        .and_then(|count| count.checked_add(1))
        .ok_or(Error::InvalidSource)?;
    let mut token_selectors = Vec::new();
    token_selectors
        .try_reserve_exact(token_capacity)
        .map_err(|_| Error::Allocation {
            amount: token_capacity,
        })?;
    token_selectors.push(target_selector);
    for reference in facts.externals.iter().filter(|reference| {
        reference.target == current_uuid.component && reference.object == Some(storage_identifier)
    }) {
        let source_component = current_metadata_component(&facts, reference.source)?;
        let source_selector =
            ComponentSelector::new(source_component.identifier, source_component.effective());
        external_removals.push(ExternalReferenceRemoval::new(
            source_selector,
            target_selector,
            storage_identifier,
            reference.weak,
        ));
        if !token_selectors.iter().any(|selector| {
            selector.identifier() == source_selector.identifier()
                && selector.locator() == source_selector.locator()
        }) {
            token_selectors.push(source_selector);
        }
    }
    for owner in facts
        .data_owners
        .iter()
        .filter(|owner| owner.object == storage_identifier)
    {
        let source_component = current_metadata_component(&facts, owner.component)?;
        let source_selector =
            ComponentSelector::new(source_component.identifier, source_component.effective());
        data_removals.push(DataReferenceOwnerRemoval::new(
            source_selector,
            owner.data,
            storage_identifier,
            owner.count,
        ));
        if !token_selectors.iter().any(|selector| {
            selector.identifier() == source_selector.identifier()
                && selector.locator() == source_selector.locator()
        }) {
            token_selectors.push(source_selector);
        }
    }
    let removal = RemovalBatch::new(
        last_identifier,
        &uuid_removals,
        &external_removals,
        &data_removals,
    );
    let save_tokens = SaveTokenBatch::new(&token_selectors);
    let batch = RemovalSaveTokenBatch::new(removal, save_tokens);
    let options = metadata_options(source, location.payload.len())?;
    let prepared = package_metadata_codec::prepare_package_metadata_removals_and_save_tokens(
        location.payload,
        batch,
        options,
    )
    .map_err(map_metadata_error)?;
    let prepare_report = prepared.prepare_report();
    budget.charge_metadata_prepare(prepare_report)?;
    let requirements = prepared.execution_requirements();
    budget.charge_metadata_requirements(requirements)?;
    let execution_limits = requirements.exact_limits();
    let output = prepared
        .execute(execution_limits)
        .map_err(map_metadata_error)?;
    if output.report().output_bytes() != requirements.output_bytes()
        || output.report().fields() != requirements.fields()
        || output.report().work_bytes() != requirements.work_bytes()
        || output.report().components_scanned() != requirements.components()
        || output.report().references_scanned() != requirements.references()
        || output.report().allocations() != requirements.allocations()
        || output.report().retained_bytes() != requirements.retained_bytes()
        || output.report().scratch_bytes() > requirements.scratch_bytes()
    {
        return Err(Error::Verification);
    }
    rewrite_metadata_archive(source, location, output.into_bytes(), budget)
}

fn insert_lifecycle(
    source: &Package,
    position: Position,
    text: &str,
    custom_mark: Option<&str>,
) -> Result<Commit, Error> {
    source.validate().map_err(map_package_error)?;
    let mut budget = TransactionBudget::new(source)?;
    let position_index = usize::try_from(position.utf16_index())
        .map_err(|_| Error::Refused(Refusal::PositionOutOfBounds))?;
    if current_footnote_at_optional(source, position)?.is_some() {
        return Err(Error::Refused(Refusal::PositionOccupied));
    }
    let source_catalog = &source.state.source;
    if !source_catalog.source_is_exact() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let body_identifier =
        super::root_references_with_limits(source_catalog.components(), source_catalog.limits())
            .map_err(map_package_error)?
            .body
            .ok_or(Error::InvalidSource)?;
    let body_component = component_for_object(source, body_identifier.get())?;
    // Insertion changes the rooted body graph and publishes new metadata
    // registrations.  Prove every pre-existing footnote graph first, rather
    // than validating only the graph whose body table will be edited: an
    // alias in a sibling note would otherwise make the package unsafe after
    // the new candidate is published.
    validate_existing_footnote_graphs(source, &mut budget)?;
    let (body_archive, archive_limits) = editable_archive(source, &body_component, &mut budget)?;
    let rewrite_limits = super::storage_rewrite_limits(source_catalog.limits())
        .map_err(map_storage_rewrite_limits_error)?;
    let body_object = body_archive
        .object(body_identifier.get())
        .ok_or(Error::InvalidSource)?;
    let body_message_index = unique_body_message_index(body_object, rewrite_limits)?;
    let body_message = body_object.messages[body_message_index].clone();
    let position_u32 =
        u32::try_from(position_index).map_err(|_| Error::Refused(Refusal::PositionOutOfBounds))?;
    let max_identifier = max_object_identifier(source, &mut budget)?;
    let reference_identifier =
        NonZeroU64::new(max_identifier.checked_add(1).ok_or(Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
    let storage_identifier =
        NonZeroU64::new(max_identifier.checked_add(2).ok_or(Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
    let marker_identifier =
        NonZeroU64::new(max_identifier.checked_add(3).ok_or(Error::InvalidSource)?)
            .ok_or(Error::InvalidSource)?;
    let (new_body_payload, body_template) = rewrite_body_storage(
        &body_message.data,
        position_index,
        None,
        position_u32,
        reference_identifier,
        custom_mark,
        source_catalog.limits(),
        &mut budget,
    )?;
    let graph_write = FootnoteGraphWrite::new(
        reference_identifier,
        storage_identifier,
        marker_identifier,
        position_u32,
        text,
    )
    .with_custom_mark(custom_mark)
    .with_storage_template(
        body_template.stylesheet,
        body_template.paragraph_style,
        body_template.list_style,
        body_template.language.as_deref(),
    );
    let graph_options = graph_options(source, body_message.data.len())?;
    let (payloads, graph_report) =
        graph_codec::encode_footnote_graph_with_report(graph_write, graph_options)
            .map_err(map_graph_error)?;
    budget.charge_graph_encode(graph_report)?;
    let mut archive = body_archive;
    let body_object = archive
        .object_mut(body_identifier.get())
        .ok_or(Error::InvalidSource)?;
    body_object
        .replace_message_preserving_header(
            body_message_index,
            RawMessage {
                type_: body_message.type_,
                data: new_body_payload,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    update_body_header_reference(
        &mut body_object.archive_info.message_infos[body_message_index],
        reference_identifier.get(),
        true,
        archive_limits,
    )?;
    let objects = graph_objects(graph_write, &payloads, archive_limits, &mut budget)?;
    budget.charge(LimitKind::Allocations, 1)?;
    archive
        .append_objects(objects.to_vec(), archive_limits)
        .map_err(map_core_error)?;
    let snappy_limits = source_catalog
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let body_compressed = compress_archive(archive, archive_limits, snappy_limits, &mut budget)?;

    let metadata = metadata_addition(
        source,
        body_component.as_str(),
        storage_identifier.get(),
        marker_identifier.get(),
        &mut budget,
    )?;
    let mut edits = Vec::new();
    budget.charge(LimitKind::Allocations, 1)?;
    edits
        .try_reserve_exact(2)
        .map_err(|_| Error::Allocation { amount: 2 })?;
    edits.push(EntryEdit::new(body_component.as_str(), &body_compressed));
    edits.push(EntryEdit::new(
        metadata.member.as_str(),
        &metadata.compressed,
    ));
    let previews = preview_names(source);
    let output = reassemble_candidate(source_catalog, &edits, &previews, &mut budget)?;
    budget.charge(LimitKind::Allocations, 1)?;
    let candidate_source = source_catalog
        .reopen(output.into())
        .map_err(map_archive_error)?;
    let candidate = Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
    let expected = Footnote::with_custom_mark(
        position,
        try_boxed(text)?,
        custom_mark.map(try_boxed).transpose()?,
    )
    .map_err(|_| Error::InvalidSource)?;
    verify_lifecycle_target(&candidate, position, Some(&expected))?;
    let source_bytes: Arc<[u8]> = source_catalog.shared_source();
    let target_bytes = candidate.state.source.shared_source();
    let target_sequence = footnote_sequence(&candidate)?;
    let package = candidate;
    Ok(Commit {
        package,
        patch: Patch {
            source_fingerprint: super::section_transaction::fingerprint(&source_bytes),
            target_fingerprint: super::section_transaction::fingerprint(&target_bytes),
            source_bytes,
            target_bytes,
            source_sequence: footnote_sequence(source)?,
            target_sequence,
            position,
            before: None,
            after: Some(Arc::new(expected)),
        },
        diagnostics: Diagnostics::published(1),
    })
}

fn rewrite_remove(source: &Package, position: Position) -> Result<Package, Error> {
    let mut budget = TransactionBudget::new(source)?;
    let graphs = super::footnote_text::native_footnotes(source).map_err(map_text_error)?;
    let (graph_index, graph) = graphs
        .iter()
        .enumerate()
        .find(|(_, graph)| graph.position == position)
        .map(|(index, graph)| (index, *graph))
        .ok_or(Error::Refused(Refusal::NotFound))?;
    let source_catalog = &source.state.source;
    let body_component = component_for_object(source, graph.body_identifier.get())?;
    let (mut archive, archive_limits) = editable_archive(source, &body_component, &mut budget)?;
    let rewrite_limits = super::storage_rewrite_limits(source_catalog.limits())
        .map_err(map_storage_rewrite_limits_error)?;
    let body_object = archive
        .object(graph.body_identifier.get())
        .ok_or(Error::InvalidSource)?;
    let body_message_index = unique_body_message_index(body_object, rewrite_limits)?;
    let body_message = body_object.messages[body_message_index].clone();
    let body_payload = rewrite_body_storage(
        &body_message.data,
        usize::try_from(position.utf16_index())
            .map_err(|_| Error::Refused(Refusal::PositionOutOfBounds))?,
        Some(graph.reference_identifier),
        position.utf16_index(),
        graph.reference_identifier,
        None,
        source_catalog.limits(),
        &mut budget,
    )?
    .0;
    let body_object = archive
        .object_mut(graph.body_identifier.get())
        .ok_or(Error::InvalidSource)?;
    body_object
        .replace_message_preserving_header(
            body_message_index,
            RawMessage {
                type_: body_message.type_,
                data: body_payload,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    update_body_header_reference(
        &mut body_object.archive_info.message_infos[body_message_index],
        graph.reference_identifier.get(),
        false,
        archive_limits,
    )?;
    let census = GraphCensus::new(source, &graphs, &mut budget)?;
    prove_exclusive_graph(graph_index, &graph, &census)?;
    drop(census);
    for id in [
        graph.reference_identifier,
        graph.storage_identifier,
        graph.marker_identifier,
    ] {
        archive.remove_object(id.get());
    }
    let snappy_limits = source_catalog
        .limits()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let compressed = compress_archive(archive, archive_limits, snappy_limits, &mut budget)?;
    let metadata = metadata_removal(source, graph.storage_identifier.get(), &mut budget)?;
    let mut edits = Vec::new();
    budget.charge(LimitKind::Allocations, 1)?;
    edits
        .try_reserve_exact(2)
        .map_err(|_| Error::Allocation { amount: 2 })?;
    edits.push(EntryEdit::new(body_component.as_str(), &compressed));
    edits.push(EntryEdit::new(
        metadata.member.as_str(),
        &metadata.compressed,
    ));
    let previews = preview_names(source);
    let output = reassemble_candidate(source_catalog, &edits, &previews, &mut budget)?;
    budget.charge(LimitKind::Allocations, 1)?;
    let candidate_source = source_catalog
        .reopen(output.into())
        .map_err(map_archive_error)?;
    Package::from_source_catalog(candidate_source).map_err(map_package_error)
}

fn reassemble_candidate(
    source: &SourceCatalog,
    edits: &[EntryEdit<'_>],
    deleted_names: &[&str],
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let prepared = source
        .package()
        .prepare_reassembly_with_deletions(edits, deleted_names, source.limits())
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget.charge_reassembly(requirements)?;
    prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)
}

fn verify_lifecycle_target(
    candidate: &Package,
    position: Position,
    expected: Option<&Footnote>,
) -> Result<(), Error> {
    let actual = current_footnote_at_optional(candidate, position)?;
    if actual.as_ref() != expected {
        return Err(Error::Verification);
    }
    Ok(())
}

#[derive(Clone)]
struct BodyTemplate {
    stylesheet: Option<NonZeroU64>,
    paragraph_style: Option<NonZeroU64>,
    list_style: Option<NonZeroU64>,
    language: Option<String>,
}

fn graph_options(source: &Package, payload: usize) -> Result<GraphDecodeOptions, Error> {
    let limits = source.state.source.limits();
    let archive = limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let output = payload
        .checked_mul(8)
        .and_then(|value| value.checked_add(1024))
        .ok_or(Error::InvalidSource)?
        .min(limits.max_iwa_stream_bytes())
        .max(1);
    let fields = archive
        .max_header_fields()
        .checked_mul(16)
        .ok_or(Error::InvalidSource)?
        .max(1);
    let work = payload.checked_mul(32).ok_or(Error::InvalidSource)?.max(1);
    Ok(GraphDecodeOptions::new(
        payload.max(1),
        output,
        fields,
        work,
        u32::try_from(archive.max_header_nesting()).unwrap_or(u32::MAX),
        MAX_TEXT_BYTES.max(1),
    ))
}

fn component_for_object(package: &Package, identifier: u64) -> Result<String, Error> {
    let mut found = None;
    for component in package.state.source.components().iter() {
        if component.archive().object(identifier).is_some() {
            if found.is_some() {
                return Err(Error::InvalidSource);
            }
            found = Some(component.name().to_owned());
        }
    }
    found.ok_or(Error::InvalidSource)
}

fn editable_archive(
    package: &Package,
    name: &str,
    budget: &mut TransactionBudget,
) -> Result<(Archive, litchi_iwa_core::Limits), Error> {
    let entry = package
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(Error::InvalidSource)?;
    if entry.is_opaque() {
        return Err(Error::Refused(Refusal::UnsupportedSource));
    }
    let limits = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress(
        entry.data(),
        package
            .state
            .source
            .limits()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    budget.charge(LimitKind::EntryBytes, stream.as_bytes().len())?;
    budget.charge(LimitKind::TotalBytes, stream.as_bytes().len())?;
    budget.charge(LimitKind::RetainedBytes, stream.as_bytes().len())?;
    budget.charge(LimitKind::ScratchBytes, stream.as_bytes().len())?;
    budget.charge(LimitKind::Allocations, 1)?;
    let archive = Archive::parse(stream.as_bytes(), limits).map_err(map_core_error)?;
    Ok((archive, limits))
}

fn compress_archive(
    archive: Archive,
    limits: litchi_iwa_core::Limits,
    snappy_limits: litchi_iwa_core::SnappyLimits,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, Error> {
    let encoded_length = archive.encoded_len(limits).map_err(map_core_error)?;
    let maximum_compressed = SnappyStream::maximum_compressed_len(encoded_length, snappy_limits)
        .map_err(map_core_error)?;
    budget.charge_archive_output(encoded_length, maximum_compressed)?;
    let bytes = archive.to_bytes(limits).map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&bytes, snappy_limits).map_err(map_core_error)?;
    if compressed.len() > maximum_compressed {
        return Err(Error::Verification);
    }
    Ok(compressed)
}

fn unique_body_message_index(
    object: &ArchiveObject,
    rewrite_limits: RewriteLimits,
) -> Result<usize, Error> {
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        let is_candidate = if is_body_storage_message_type(message.type_) {
            true
        } else {
            match litchi_iwa_text_wire::decode_storage(&message.data, rewrite_limits) {
                Ok(_) => true,
                Err(error @ litchi_iwa_text_wire::RewriteError::LimitExceeded { .. })
                | Err(error @ litchi_iwa_text_wire::RewriteError::Allocation { .. }) => {
                    return Err(map_text_wire_error(error));
                },
                Err(_) => false,
            }
        };
        if is_candidate {
            if found.replace(index).is_some() {
                return Err(Error::InvalidSource);
            }
        }
    }
    found.ok_or(Error::InvalidSource)
}

fn is_body_storage_message_type(type_id: u32) -> bool {
    matches!(type_id, 2_001 | 2_022)
}

fn rewrite_body_storage(
    payload: &[u8],
    position: usize,
    remove_reference: Option<NonZeroU64>,
    position_u32: u32,
    reference_identifier: NonZeroU64,
    _custom_mark: Option<&str>,
    limits: super::Limits,
    budget: &mut TransactionBudget,
) -> Result<(Vec<u8>, BodyTemplate), Error> {
    // The text-wire rewrite is allowed to canonicalize the known storage
    // projection.  Inspect the source table first so deprecated
    // `TSP.Reference` fields cannot disappear during that rewrite and make a
    // hostile graph look canonical to the strict graph codec.
    let source_wire_limits = storage_wire_limits(limits, payload.len(), payload.len())?;
    reject_deprecated_body_references(payload, source_wire_limits)?;
    let rewrite_limits =
        super::storage_rewrite_limits(limits).map_err(map_storage_rewrite_limits_error)?;
    let range = if remove_reference.is_some() {
        position
            ..position
                .checked_add(1)
                .ok_or(Error::Refused(Refusal::PositionOutOfBounds))?
    } else {
        position..position
    };
    let replacement = if remove_reference.is_some() {
        ""
    } else {
        FOOTNOTE_ANCHOR_TEXT
    };
    let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior(
        payload,
        range,
        replacement,
        RewriteBehavior::PreserveOnEqualText,
        rewrite_limits,
    )
    .map_err(map_text_wire_error)?;
    let text_requirements = prepared.execution_requirements();
    budget.charge_text_requirements(text_requirements)?;
    let rewritten_result = prepared
        .execute(StorageRewriteExecutionLimits {
            max_output_bytes: text_requirements.output_bytes(),
            max_retained_elements: text_requirements.retained_elements(),
            max_retained_bytes: text_requirements.retained_bytes(),
            max_peak_scratch_bytes: text_requirements.peak_scratch_bytes(),
            max_allocations: text_requirements.allocations(),
            max_work: text_requirements.work(),
        })
        .map_err(map_text_wire_error)?;
    let execution_report = rewritten_result.execution_report();
    if rewritten_result.bytes().len() != text_requirements.output_bytes()
        || execution_report.retained_elements > text_requirements.retained_elements()
        || execution_report.retained_bytes > text_requirements.retained_bytes()
        || execution_report.peak_scratch_bytes > text_requirements.peak_scratch_bytes()
        || execution_report.allocations > text_requirements.allocations()
        || execution_report.work > text_requirements.work()
    {
        return Err(Error::Verification);
    }
    budget.observe(LimitKind::TextUnits, rewritten_result.after_utf16_len())?;
    let rewritten = rewritten_result.into_bytes();
    budget.charge(LimitKind::TextBytes, rewritten.len())?;
    // Parsing the rewritten storage does not allocate an output buffer.  Use
    // the exact input length for this read-only wire profile; the former 2x
    // reservation could reject a valid payload before the actual replacement
    // size was known.
    let wire_output = rewritten.len().max(1);
    let wire_limits = storage_wire_limits(limits, rewritten.len(), wire_output)?;
    let view = WireView::parse(&rewritten, wire_limits).map_err(map_common_wire_error)?;
    let table_field = view
        .fields()
        .find(|field| field.number() == FOOTNOTE_TABLE_FIELD)
        .map(|field| field.payload());
    let source_table = table_field.unwrap_or(&[]);
    let archive_limits = limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let table_output = source_table
        .len()
        .checked_mul(8)
        .and_then(|value| value.checked_add(1024))
        .ok_or(Error::InvalidSource)?
        .min(limits.max_iwa_stream_bytes())
        .max(1);
    let table_fields = rewritten
        .len()
        .checked_mul(16)
        .ok_or(Error::InvalidSource)?
        .max(1);
    let table_work = rewritten
        .len()
        .checked_mul(32)
        .ok_or(Error::InvalidSource)?
        .max(1);
    let table_options = GraphDecodeOptions::new(
        source_table.len().max(1),
        table_output,
        table_fields,
        table_work,
        u32::try_from(archive_limits.max_header_nesting()).unwrap_or(u32::MAX),
        MAX_TEXT_BYTES.max(1),
    );
    let (snapshot, decode_report) =
        graph_codec::decode_body_footnote_table_with_report(source_table, table_options)
            .map_err(map_graph_error)?;
    budget.charge_graph_decode(decode_report)?;
    let mut entries = Vec::new();
    let entry_capacity = snapshot
        .len()
        .checked_add(usize::from(remove_reference.is_none()))
        .ok_or(Error::InvalidSource)?;
    budget.charge(LimitKind::Entries, entry_capacity)?;
    budget.charge(LimitKind::Allocations, 1)?;
    entries
        .try_reserve_exact(entry_capacity)
        .map_err(|_| Error::Allocation {
            amount: snapshot.len(),
        })?;
    for entry in snapshot.entries() {
        if remove_reference.is_some_and(|id| id == entry.reference_identifier()) {
            continue;
        }
        entries.push(BodyFootnoteEntryWrite::preserve(entry));
    }
    if remove_reference.is_none() {
        entries.push(BodyFootnoteEntryWrite::new(
            position_u32,
            reference_identifier,
        ));
        entries.sort_by_key(|entry| entry.character_index());
    }
    let (rewritten_table, rewrite_report) = graph_codec::rewrite_body_footnote_table_with_report(
        source_table,
        BodyFootnoteTableWrite::new(&entries),
        table_options,
    )
    .map_err(map_graph_error)?;
    budget.charge_graph_rewrite(rewrite_report)?;
    let replacement_limits = storage_wire_limits(limits, rewritten.len(), rewritten.len())?;
    let replacement_output = length_delimited_field_output_len(
        &rewritten,
        FOOTNOTE_TABLE_FIELD,
        if entries.is_empty() {
            None
        } else {
            Some(&rewritten_table)
        },
        replacement_limits,
    )?;
    budget.charge(LimitKind::WireBytes, replacement_output)?;
    budget.charge(LimitKind::RetainedBytes, replacement_output)?;
    budget.charge(LimitKind::Allocations, 1)?;
    let output = replace_length_delimited_field(
        &rewritten,
        FOOTNOTE_TABLE_FIELD,
        if entries.is_empty() {
            None
        } else {
            Some(&rewritten_table)
        },
        replacement_limits,
        archive_limits.max_message_bytes(),
    )?;
    let template = body_template_from_storage_wire(&rewritten, wire_limits)?;
    Ok((output, template))
}

fn reject_deprecated_body_references(source: &[u8], limits: WireLimits) -> Result<(), Error> {
    let storage = WireView::parse(source, limits).map_err(map_common_wire_error)?;
    let Some(table_field) = storage
        .fields()
        .find(|field| field.number() == FOOTNOTE_TABLE_FIELD)
    else {
        return Ok(());
    };
    if table_field.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    let table = WireView::parse(table_field.payload(), limits).map_err(map_common_wire_error)?;
    for entry_field in table.fields().filter(|field| field.number() == 1) {
        if entry_field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        let entry =
            WireView::parse(entry_field.payload(), limits).map_err(map_common_wire_error)?;
        let Some(reference_field) = entry.fields().find(|field| field.number() == 2) else {
            continue;
        };
        if reference_field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        let reference =
            WireView::parse(reference_field.payload(), limits).map_err(map_common_wire_error)?;
        if reference
            .fields()
            .any(|field| matches!(field.number(), 2 | 3))
        {
            return Err(Error::InvalidSource);
        }
    }
    Ok(())
}

fn body_template_from_storage_wire(
    source: &[u8],
    limits: WireLimits,
) -> Result<BodyTemplate, Error> {
    let view = WireView::parse(source, limits).map_err(map_common_wire_error)?;
    let mut stylesheet = None;
    let mut paragraph_style: Option<Option<NonZeroU64>> = None;
    let mut list_style: Option<Option<NonZeroU64>> = None;
    let mut language: Option<Option<String>> = None;
    for field in view.fields() {
        match field.number() {
            2 => {
                field
                    .validate_canonical_key()
                    .and_then(|_| field.validate_canonical_length())
                    .map_err(|_| Error::InvalidSource)?;
                let value = parse_storage_reference(field.payload(), limits)?;
                if stylesheet.replace(value).is_some() {
                    return Err(Error::InvalidSource);
                }
            },
            5 | 7 => {
                field
                    .validate_canonical_key()
                    .and_then(|_| field.validate_canonical_length())
                    .map_err(|_| Error::InvalidSource)?;
                let value = parse_style_table_reference(field.payload(), limits)?;
                let slot = if field.number() == 5 {
                    &mut paragraph_style
                } else {
                    &mut list_style
                };
                if slot.replace(value).is_some() {
                    return Err(Error::InvalidSource);
                }
            },
            19 => {
                field
                    .validate_canonical_key()
                    .and_then(|_| field.validate_canonical_length())
                    .map_err(|_| Error::InvalidSource)?;
                let value = parse_language_table(field.payload(), limits)?;
                if language.replace(value).is_some() {
                    return Err(Error::InvalidSource);
                }
            },
            _ => {},
        }
    }
    Ok(BodyTemplate {
        stylesheet,
        paragraph_style: paragraph_style.flatten(),
        list_style: list_style.flatten(),
        language: language.flatten(),
    })
}

fn parse_language_table(source: &[u8], limits: WireLimits) -> Result<Option<String>, Error> {
    let table = WireView::parse(source, limits).map_err(map_common_wire_error)?;
    let mut language_at_zero = None;
    for field in table.fields() {
        if field.number() != 1 {
            return Err(Error::InvalidSource);
        }
        field
            .validate_canonical_key()
            .and_then(|_| field.validate_canonical_length())
            .map_err(|_| Error::InvalidSource)?;
        let (index, value) = parse_language_entry(field.payload(), limits)?;
        if index == 0 {
            if language_at_zero.is_some() {
                return Err(Error::InvalidSource);
            }
            language_at_zero = value;
        }
    }
    Ok(language_at_zero)
}

fn parse_language_entry(source: &[u8], limits: WireLimits) -> Result<(u32, Option<String>), Error> {
    let entry = WireView::parse(source, limits).map_err(map_common_wire_error)?;
    let mut index = None;
    let mut value = None;
    for field in entry.fields() {
        match field.number() {
            1 => {
                field
                    .validate_canonical_key()
                    .map_err(|_| Error::InvalidSource)?;
                if field.wire_type() != 0 || index.is_some() {
                    return Err(Error::InvalidSource);
                }
                let (raw, width) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                    .map_err(|_| Error::InvalidSource)?;
                if width != field.payload().len() {
                    return Err(Error::InvalidSource);
                }
                index = Some(u32::try_from(raw).map_err(|_| Error::InvalidSource)?);
            },
            2 => {
                field
                    .validate_canonical_key()
                    .and_then(|_| field.validate_canonical_length())
                    .map_err(|_| Error::InvalidSource)?;
                if value.is_some() {
                    return Err(Error::InvalidSource);
                }
                value = Some(
                    std::str::from_utf8(field.payload())
                        .map_err(|_| Error::InvalidSource)?
                        .to_owned(),
                );
            },
            _ => {},
        }
    }
    Ok((index.ok_or(Error::InvalidSource)?, value))
}

fn parse_storage_reference(source: &[u8], limits: WireLimits) -> Result<NonZeroU64, Error> {
    let view = WireView::parse(source, limits).map_err(map_common_wire_error)?;
    let mut identifier = None;
    for field in view.fields() {
        match field.number() {
            1 => {
                field
                    .validate_canonical_key()
                    .map_err(|_| Error::InvalidSource)?;
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(Error::InvalidSource);
                }
                let (value, width) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                    .map_err(|_| Error::InvalidSource)?;
                if width != field.payload().len() {
                    return Err(Error::InvalidSource);
                }
                identifier = Some(NonZeroU64::new(value).ok_or(Error::InvalidSource)?);
            },
            2 | 3 => return Err(Error::InvalidSource),
            _ => {},
        }
    }
    identifier.ok_or(Error::InvalidSource)
}

fn parse_style_table_reference(
    source: &[u8],
    limits: WireLimits,
) -> Result<Option<NonZeroU64>, Error> {
    let table = WireView::parse(source, limits).map_err(map_common_wire_error)?;
    let mut reference_at_zero = None;
    let mut saw_entry = false;
    for field in table.fields() {
        if field.number() != 1 {
            continue;
        }
        saw_entry = true;
        field
            .validate_canonical_key()
            .and_then(|_| field.validate_canonical_length())
            .map_err(|_| Error::InvalidSource)?;
        let (index, reference) = parse_style_table_entry(field.payload(), limits)?;
        if index == 0 {
            if reference_at_zero.is_some() {
                return Err(Error::InvalidSource);
            }
            reference_at_zero = reference;
        }
    }
    if saw_entry {
        Ok(reference_at_zero)
    } else {
        Ok(None)
    }
}

fn parse_style_table_entry(
    source: &[u8],
    limits: WireLimits,
) -> Result<(u32, Option<NonZeroU64>), Error> {
    let entry = WireView::parse(source, limits).map_err(map_common_wire_error)?;
    let mut index = None;
    let mut reference = None;
    for field in entry.fields() {
        match field.number() {
            1 => {
                field
                    .validate_canonical_key()
                    .map_err(|_| Error::InvalidSource)?;
                if field.wire_type() != 0 || index.is_some() {
                    return Err(Error::InvalidSource);
                }
                let (value, width) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                    .map_err(|_| Error::InvalidSource)?;
                if width != field.payload().len() {
                    return Err(Error::InvalidSource);
                }
                index = Some(u32::try_from(value).map_err(|_| Error::InvalidSource)?);
            },
            2 => {
                field
                    .validate_canonical_key()
                    .and_then(|_| field.validate_canonical_length())
                    .map_err(|_| Error::InvalidSource)?;
                if reference.is_some() {
                    return Err(Error::InvalidSource);
                }
                reference = Some(parse_storage_reference(field.payload(), limits)?);
            },
            _ => {},
        }
    }
    Ok((index.ok_or(Error::InvalidSource)?, reference))
}

fn replace_length_delimited_field(
    source: &[u8],
    number: u32,
    replacement: Option<&[u8]>,
    limits: WireLimits,
    max_output_bytes: usize,
) -> Result<Vec<u8>, Error> {
    let view = WireView::parse(source, limits).map_err(map_common_wire_error)?;
    let output_length = length_delimited_field_output_len_from_view(&view, number, replacement)?;
    if output_length > max_output_bytes {
        return Err(Error::Limit {
            kind: LimitKind::WireBytes,
            observed: output_length as u64,
            maximum: max_output_bytes as u64,
        });
    }
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| Error::Allocation {
            amount: output_length,
        })?;
    let mut found = false;
    for field in view.fields() {
        if field.number() == number {
            if found {
                return Err(Error::InvalidSource);
            }
            found = true;
            if let Some(payload) = replacement {
                append_length_delimited(&mut output, number, payload)?;
            }
        } else {
            output.extend_from_slice(field.raw());
        }
    }
    if !found {
        if let Some(payload) = replacement {
            append_length_delimited(&mut output, number, payload)?;
        }
    }
    if output.len() != output_length {
        return Err(Error::Verification);
    }
    Ok(output)
}

fn length_delimited_field_output_len(
    source: &[u8],
    number: u32,
    replacement: Option<&[u8]>,
    limits: WireLimits,
) -> Result<usize, Error> {
    let view = WireView::parse(source, limits).map_err(map_common_wire_error)?;
    length_delimited_field_output_len_from_view(&view, number, replacement)
}

fn length_delimited_field_output_len_from_view(
    view: &WireView<'_>,
    number: u32,
    replacement: Option<&[u8]>,
) -> Result<usize, Error> {
    let replacement_length = if let Some(payload) = replacement {
        let key = u64::from(number)
            .checked_shl(3)
            .and_then(|value| value.checked_add(2))
            .ok_or(Error::InvalidSource)?;
        let payload_length = u64::try_from(payload.len()).map_err(|_| Error::InvalidSource)?;
        let framing = litchi_iwa_common::varint::encoded_len(key)
            .checked_add(litchi_iwa_common::varint::encoded_len(payload_length))
            .ok_or(Error::InvalidSource)?;
        Some(
            framing
                .checked_add(payload.len())
                .ok_or(Error::InvalidSource)?,
        )
    } else {
        None
    };
    let mut output_length = 0usize;
    let mut found = false;
    for field in view.fields() {
        if field.number() == number {
            if found {
                return Err(Error::InvalidSource);
            }
            found = true;
            if let Some(length) = replacement_length {
                output_length = output_length
                    .checked_add(length)
                    .ok_or(Error::InvalidSource)?;
            }
        } else {
            output_length = output_length
                .checked_add(field.raw().len())
                .ok_or(Error::InvalidSource)?;
        }
    }
    if !found {
        if let Some(length) = replacement_length {
            output_length = output_length
                .checked_add(length)
                .ok_or(Error::InvalidSource)?;
        }
    }
    Ok(output_length)
}

fn append_length_delimited(output: &mut Vec<u8>, number: u32, payload: &[u8]) -> Result<(), Error> {
    put_varint(output, (u64::from(number) << 3) | 2);
    put_varint(
        output,
        u64::try_from(payload.len()).map_err(|_| Error::InvalidSource)?,
    );
    output.extend_from_slice(payload);
    Ok(())
}

fn put_varint(output: &mut Vec<u8>, mut value: u64) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        output.push(byte);
        if value == 0 {
            break;
        }
    }
}

fn graph_objects(
    write: FootnoteGraphWrite<'_>,
    payloads: &graph_codec::FootnoteGraphPayloads,
    limits: litchi_iwa_core::Limits,
    budget: &mut TransactionBudget,
) -> Result<[ArchiveObject; 3], Error> {
    budget.charge(LimitKind::Entries, 3)?;
    budget.charge(LimitKind::Allocations, 3)?;
    let mut reference = ArchiveObject::new(
        write.reference_identifier().get(),
        vec![RawMessage {
            type_: FOOTNOTE_REFERENCE_MESSAGE_TYPE,
            data: payloads.reference().to_vec(),
        }],
        limits,
    )
    .map_err(map_core_error)?;
    let mut storage_references = vec![write.marker_identifier().get()];
    if let Some(identifier) = write.stylesheet_identifier() {
        storage_references.push(identifier.get());
    }
    if let Some(identifier) = write.paragraph_style_identifier() {
        storage_references.push(identifier.get());
    }
    if let Some(identifier) = write.list_style_identifier() {
        storage_references.push(identifier.get());
    }
    reference.archive_info.message_infos[0].versions = vec![1, 0, 5].into();
    reference.archive_info.message_infos[0].object_references =
        vec![write.storage_identifier().get()];
    let mut storage = ArchiveObject::new(
        write.storage_identifier().get(),
        vec![RawMessage {
            type_: 2_001,
            data: payloads.storage().to_vec(),
        }],
        limits,
    )
    .map_err(map_core_error)?;
    storage.archive_info.message_infos[0].versions = vec![1, 0, 5].into();
    storage.archive_info.message_infos[0].object_references = storage_references;
    let mut marker = ArchiveObject::new(
        write.marker_identifier().get(),
        vec![RawMessage {
            type_: TEXTUAL_ATTACHMENT_MESSAGE_TYPE,
            data: payloads.marker().to_vec(),
        }],
        limits,
    )
    .map_err(map_core_error)?;
    marker.archive_info.message_infos[0].versions = vec![1, 0, 5].into();
    Ok([reference, storage, marker])
}

fn max_object_identifier(package: &Package, budget: &mut TransactionBudget) -> Result<u64, Error> {
    let physical_count =
        package
            .state
            .source
            .components()
            .iter()
            .try_fold(0usize, |count, component| {
                count
                    .checked_add(component.archive().objects.len())
                    .ok_or(Error::InvalidSource)
            })?;
    budget.charge(LimitKind::Entries, physical_count)?;
    let physical = package
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .filter_map(|object| object.archive_info.identifier)
        .max()
        .ok_or(Error::InvalidSource)?;
    let location = metadata_location(package)?;
    let (facts, last) = metadata_facts(package, &location, budget)?;
    let mut maximum = physical.max(last);
    for component in &facts.components {
        maximum = maximum.max(component.identifier);
    }
    for binding in &facts.uuids {
        maximum = maximum.max(binding.object);
    }
    for reference in &facts.externals {
        if let Some(object) = reference.object {
            maximum = maximum.max(object);
        }
        maximum = maximum.max(reference.source).max(reference.target);
    }
    for owner in &facts.data_owners {
        maximum = maximum
            .max(owner.component)
            .max(owner.data)
            .max(owner.object);
    }
    for (component, object) in &facts.ambiguous {
        maximum = maximum.max(*component).max(*object);
    }
    for (object, _) in &facts.data_metadata_maps {
        maximum = maximum.max(*object);
    }
    Ok(maximum)
}

fn preview_names(package: &Package) -> Vec<&'static str> {
    PREVIEW_NAMES
        .iter()
        .copied()
        .filter(|name| {
            package
                .state
                .source
                .package()
                .iter()
                .any(|entry| entry.name() == *name)
        })
        .collect()
}

fn update_body_header_reference(
    info: &mut litchi_iwa_core::MessageInfo,
    reference: u64,
    adding: bool,
    limits: litchi_iwa_core::Limits,
) -> Result<(), Error> {
    const BODY_REFERENCE_PATH: &[u32] = &[16, 1, 2];
    let aggregate_count = info
        .object_references
        .iter()
        .filter(|identifier| **identifier == reference)
        .count();
    if adding {
        if aggregate_count != 0 {
            return Err(Error::InvalidSource);
        }
    } else if aggregate_count != 1 {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    let mut canonical_fields = 0usize;
    for field in &mut info.field_infos {
        if field.path.path == BODY_REFERENCE_PATH {
            canonical_fields = canonical_fields
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
            if canonical_fields > 1 {
                return Err(Error::InvalidSource);
            }
            let count = field
                .object_references
                .iter()
                .filter(|identifier| **identifier == reference)
                .count();
            if count > 1 {
                return Err(Error::InvalidSource);
            }
            if adding {
                if count != 0 {
                    return Err(Error::InvalidSource);
                }
                field.object_references.push(reference);
            } else {
                field
                    .object_references
                    .retain(|identifier| *identifier != reference);
            }
        } else if field.object_references.contains(&reference) {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
    }
    if adding {
        info.object_references.push(reference);
        if canonical_fields == 0 {
            let path = litchi_iwa_core::FieldPath::new(BODY_REFERENCE_PATH.to_vec(), limits)
                .map_err(|_| Error::InvalidSource)?;
            let mut field = FieldInfo::new(path, limits).map_err(|_| Error::InvalidSource)?;
            field.object_references.push(reference);
            info.field_infos.push(field);
        }
    } else {
        info.object_references
            .retain(|identifier| *identifier != reference);
    }
    Ok(())
}

#[derive(Debug, Clone, Copy)]
struct GraphCensusEntry {
    identifier: u64,
    graph_index: usize,
    slot: usize,
    incoming: usize,
    typed_incoming: usize,
    unsupported: bool,
    ambiguous: bool,
}

struct GraphCensus {
    entries: Vec<GraphCensusEntry>,
}

impl GraphCensus {
    fn new(
        package: &Package,
        graphs: &[super::footnote_text::NativeFootnote],
        budget: &mut TransactionBudget,
    ) -> Result<Self, Error> {
        let entry_count = graphs.len().checked_mul(3).ok_or(Error::InvalidSource)?;
        budget.charge(LimitKind::Entries, graphs.len())?;
        budget.charge(
            LimitKind::WireWork,
            package.state.source.source_bytes().len(),
        )?;
        let census_bytes = entry_count
            .checked_mul(size_of::<GraphCensusEntry>())
            .ok_or(Error::InvalidSource)?;
        budget.charge(LimitKind::ScratchBytes, census_bytes)?;
        let mut entries = Vec::new();
        if entry_count != 0 {
            budget.charge(LimitKind::Allocations, 1)?;
            entries
                .try_reserve_exact(entry_count)
                .map_err(|_| Error::Allocation {
                    amount: entry_count,
                })?;
        }
        for (graph_index, graph) in graphs.iter().enumerate() {
            for (slot, identifier) in [
                graph.reference_identifier.get(),
                graph.storage_identifier.get(),
                graph.marker_identifier.get(),
            ]
            .into_iter()
            .enumerate()
            {
                entries.push(GraphCensusEntry {
                    identifier,
                    graph_index,
                    slot,
                    incoming: 0,
                    typed_incoming: 0,
                    unsupported: false,
                    ambiguous: false,
                });
            }
        }
        entries.sort_unstable_by_key(|entry| entry.identifier);
        for index in 1..entries.len() {
            if entries[index - 1].identifier == entries[index].identifier {
                entries[index - 1].ambiguous = true;
                entries[index].ambiguous = true;
            }
        }

        let mut census = Self { entries };
        for component in package.state.source.components().iter() {
            for object in &component.archive().objects {
                for info in &object.archive_info.message_infos {
                    for target in &info.object_references {
                        if let Some(entry) = census.entry_mut(*target) {
                            entry.incoming =
                                entry.incoming.checked_add(1).ok_or(Error::InvalidSource)?;
                        }
                    }
                    for target in &info.data_references {
                        if let Some(entry) = census.entry_mut(*target) {
                            entry.unsupported = true;
                        }
                    }
                    for field in &info.field_infos {
                        for target in &field.object_references {
                            let Some(entry) = census.entry_mut(*target) else {
                                continue;
                            };
                            let expected_path = graph_edge_path(entry.slot);
                            if field.path.path != expected_path {
                                entry.unsupported = true;
                            } else {
                                entry.typed_incoming = entry
                                    .typed_incoming
                                    .checked_add(1)
                                    .ok_or(Error::InvalidSource)?;
                            }
                        }
                        for target in &field.data_references {
                            if let Some(entry) = census.entry_mut(*target) {
                                entry.unsupported = true;
                            }
                        }
                    }
                }
            }
        }
        Ok(census)
    }

    fn entry(&self, identifier: u64) -> Option<&GraphCensusEntry> {
        let index = self
            .entries
            .binary_search_by_key(&identifier, |entry| entry.identifier)
            .ok()?;
        let entry = &self.entries[index];
        (!entry.ambiguous).then_some(entry)
    }

    fn entry_mut(&mut self, identifier: u64) -> Option<&mut GraphCensusEntry> {
        let index = self
            .entries
            .binary_search_by_key(&identifier, |entry| entry.identifier)
            .ok()?;
        if self.entries[index].ambiguous {
            return None;
        }
        Some(&mut self.entries[index])
    }
}

fn graph_edge_path(slot: usize) -> &'static [u32] {
    match slot {
        0 => &[16, 1, 2],
        1 => &[2],
        _ => &[9, 1, 2],
    }
}

fn prove_exclusive_graph(
    graph_index: usize,
    graph: &super::footnote_text::NativeFootnote,
    census: &GraphCensus,
) -> Result<(), Error> {
    let selected = [
        census.entry(graph.reference_identifier.get()),
        census.entry(graph.storage_identifier.get()),
        census.entry(graph.marker_identifier.get()),
    ];
    if selected.iter().enumerate().any(|(slot, entry)| {
        entry.is_none_or(|entry| entry.graph_index != graph_index || entry.slot != slot)
    }) {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    if selected.iter().flatten().any(|entry| entry.unsupported) {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    let incoming = selected.map(|entry| entry.map_or(0, |entry| entry.incoming));
    let typed_incoming = selected.map(|entry| entry.map_or(0, |entry| entry.typed_incoming));
    // The source graph has exactly body -> reference -> storage -> marker.
    if incoming != [1, 1, 1] || typed_incoming.iter().any(|count| *count > 1) {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(())
}

fn validate_existing_footnote_graphs(
    source: &Package,
    budget: &mut TransactionBudget,
) -> Result<(), Error> {
    let graphs = super::footnote_text::native_footnotes(source).map_err(map_text_error)?;
    let location = metadata_location(source)?;
    let (facts, _) = metadata_facts(source, &location, budget)?;
    let census = GraphCensus::new(source, &graphs, budget)?;
    for (graph_index, graph) in graphs.iter().enumerate() {
        prove_exclusive_graph(graph_index, graph, &census)?;
        validate_storage_attachment_owners(
            source,
            graph.storage_identifier.get(),
            graph.marker_identifier.get(),
        )?;

        // A footnote reference or marker is archive-owned.  A UUID-map entry
        // for either object is an alias/foreign registry claim and must fail
        // closed before the insertion batch is constructed.  The storage is
        // the sole metadata-owned native object; permit at most one current
        // registration and never a versioned registration.
        if facts.uuids.iter().any(|binding| {
            binding.object == graph.reference_identifier.get()
                || binding.object == graph.marker_identifier.get()
        }) {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        let mut storage_binding_count = 0usize;
        let mut has_unsupported_storage_binding = false;
        for binding in facts
            .uuids
            .iter()
            .filter(|binding| binding.object == graph.storage_identifier.get())
        {
            storage_binding_count = storage_binding_count
                .checked_add(1)
                .ok_or(Error::InvalidSource)?;
            if !binding.current {
                has_unsupported_storage_binding = true;
            }
        }
        if storage_binding_count != 1 || has_unsupported_storage_binding {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
    }
    Ok(())
}

fn validate_storage_attachment_owners(
    package: &Package,
    storage_identifier: u64,
    marker_identifier: u64,
) -> Result<(), Error> {
    let object = package
        .state
        .source
        .components()
        .iter()
        .find_map(|component| component.archive().object(storage_identifier))
        .ok_or(Error::InvalidSource)?;
    let messages = object
        .messages
        .iter()
        .filter(|message| is_body_storage_message_type(message.type_));
    let mut message = messages;
    let storage_message = message.next().ok_or(Error::InvalidSource)?;
    if message.next().is_some() {
        return Err(Error::InvalidSource);
    }
    let limits = storage_wire_limits(
        package.state.source.limits(),
        storage_message.data.len(),
        storage_message.data.len(),
    )?;
    let storage = WireView::parse(&storage_message.data, limits).map_err(map_common_wire_error)?;
    let mut attachment_field = None;
    for field in storage.fields().filter(|field| field.number() == 9) {
        if attachment_field.replace(field).is_some() {
            return Err(Error::InvalidSource);
        }
    }
    let attachment_field = attachment_field.ok_or(Error::InvalidSource)?;
    if attachment_field.wire_type() != 2 {
        return Err(Error::InvalidSource);
    }
    let table =
        WireView::parse(attachment_field.payload(), limits).map_err(map_common_wire_error)?;
    let mut marker_count = 0usize;
    for entry_field in table.fields().filter(|field| field.number() == 1) {
        if entry_field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        let entry =
            WireView::parse(entry_field.payload(), limits).map_err(map_common_wire_error)?;
        let mut reference_field = None;
        for field in entry.fields().filter(|field| field.number() == 2) {
            if reference_field.replace(field).is_some() {
                return Err(Error::InvalidSource);
            }
        }
        let reference_field = reference_field.ok_or(Error::InvalidSource)?;
        if reference_field.wire_type() != 2 {
            return Err(Error::InvalidSource);
        }
        let reference =
            WireView::parse(reference_field.payload(), limits).map_err(map_common_wire_error)?;
        let mut identifier_field = None;
        for field in reference.fields().filter(|field| field.number() == 1) {
            if identifier_field.replace(field).is_some() {
                return Err(Error::InvalidSource);
            }
        }
        let identifier_field = identifier_field.ok_or(Error::InvalidSource)?;
        if identifier_field.wire_type() != 0 {
            return Err(Error::InvalidSource);
        }
        let (identifier, width) =
            litchi_iwa_common::decode_varint_from_bytes(identifier_field.payload())
                .map_err(|_| Error::InvalidSource)?;
        if width != identifier_field.payload().len() || identifier != marker_identifier {
            return Err(Error::Refused(Refusal::UnsupportedDependency));
        }
        marker_count = marker_count.checked_add(1).ok_or(Error::InvalidSource)?;
    }
    if marker_count != 1 {
        return Err(Error::Refused(Refusal::UnsupportedDependency));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn footnote(position: u32, text: &'static str) -> Footnote {
        Footnote::new(
            Position::from_utf16_index(position as usize)
                .unwrap_or_else(|error| panic!("valid position: {error}")),
            text.to_owned().into_boxed_str(),
        )
        .unwrap_or_else(|error| panic!("valid footnote: {error}"))
    }

    #[test]
    fn semantic_snapshot_resolves_typed_selectors_in_source_order() {
        let snapshot = Snapshot::from_values(vec![footnote(3, "first"), footnote(9, "second")]);

        assert_eq!(snapshot.len(), 2);
        assert!(!snapshot.is_empty());
        assert_eq!(
            snapshot.get(0).map(|value| value.text.as_ref()),
            Some("first")
        );
        assert_eq!(
            snapshot.select(1usize).map(|value| value.text.as_ref()),
            Some("second")
        );
        assert_eq!(
            snapshot
                .select(Position::from_utf16_index(3).unwrap())
                .map(|value| value.text.as_ref()),
            Some("first")
        );
        assert!(snapshot.select(Selector::index(4)).is_none());
        assert_eq!(snapshot.iter().count(), 2);
        assert_eq!(snapshot.into_iter().count(), 2);
    }

    #[test]
    fn semantic_snapshot_clone_shares_bounded_values() {
        let snapshot = Snapshot::from_values(vec![footnote(0, "shared")]);
        let clone = snapshot.clone();

        assert!(Arc::ptr_eq(&snapshot.values, &clone.values));
        assert_eq!(snapshot.as_slice(), clone.as_slice());
    }

    #[test]
    fn metadata_fact_vectors_stage_through_fallible_growth() {
        let mut values = Vec::new();
        metadata_try_push(&mut values, 7u64).unwrap();
        metadata_try_push(&mut values, 11u64).unwrap();

        assert_eq!(values, [7, 11]);
    }

    #[test]
    fn metadata_singleton_lookup_rejects_duplicate_current_components() {
        let facts = MetadataFactsVisitor {
            components: vec![
                MetadataComponent {
                    identifier: 7,
                    preferred: "Document".to_owned(),
                    locator: None,
                    current: true,
                },
                MetadataComponent {
                    identifier: 7,
                    preferred: "Document".to_owned(),
                    locator: None,
                    current: true,
                },
            ],
            ..MetadataFactsVisitor::default()
        };

        assert!(matches!(
            current_metadata_component(&facts, 7),
            Err(Error::InvalidSource)
        ));
    }

    #[test]
    fn graph_census_index_rejects_ambiguous_targets() {
        let census = GraphCensus {
            entries: vec![
                GraphCensusEntry {
                    identifier: 7,
                    graph_index: 0,
                    slot: 0,
                    incoming: 1,
                    typed_incoming: 1,
                    unsupported: false,
                    ambiguous: false,
                },
                GraphCensusEntry {
                    identifier: 11,
                    graph_index: 0,
                    slot: 1,
                    incoming: 1,
                    typed_incoming: 1,
                    unsupported: false,
                    ambiguous: false,
                },
            ],
        };
        assert_eq!(census.entry(7).map(|entry| entry.slot), Some(0));

        let ambiguous = GraphCensus {
            entries: vec![GraphCensusEntry {
                identifier: 7,
                graph_index: 0,
                slot: 0,
                incoming: 1,
                typed_incoming: 1,
                unsupported: false,
                ambiguous: true,
            }],
        };
        assert!(ambiguous.entry(7).is_none());
    }

    fn raw_length_delimited(field: u32, payload: &[u8]) -> Vec<u8> {
        let mut bytes = Vec::new();
        litchi_iwa_common::encode_varint_into(&mut bytes, (u64::from(field) << 3) | 2);
        litchi_iwa_common::encode_varint_into(
            &mut bytes,
            u64::try_from(payload.len()).expect("test payload length fits"),
        );
        bytes.extend_from_slice(payload);
        bytes
    }

    fn raw_varint(field: u32, value: u64) -> Vec<u8> {
        let mut bytes = Vec::new();
        litchi_iwa_common::encode_varint_into(&mut bytes, u64::from(field) << 3);
        litchi_iwa_common::encode_varint_into(&mut bytes, value);
        bytes
    }

    #[test]
    fn replacement_output_limit_uses_exact_encoded_length() {
        let mut source = raw_length_delimited(FOOTNOTE_TABLE_FIELD, b"old");
        source.extend_from_slice(&raw_varint(3, 7));
        let replacement = b"replacement";
        let limits = WireLimits::default()
            .with_input_bytes(source.len())
            .expect("source fits wire input limit");
        let output_length = length_delimited_field_output_len(
            &source,
            FOOTNOTE_TABLE_FIELD,
            Some(replacement),
            limits,
        )
        .expect("replacement length");
        let exact = replace_length_delimited_field(
            &source,
            FOOTNOTE_TABLE_FIELD,
            Some(replacement),
            limits,
            output_length,
        )
        .expect("exact output budget admits replacement");
        assert_eq!(exact.len(), output_length);
        assert!(matches!(
            replace_length_delimited_field(
                &source,
                FOOTNOTE_TABLE_FIELD,
                Some(replacement),
                limits,
                output_length - 1,
            ),
            Err(Error::Limit {
                kind: LimitKind::WireBytes,
                observed,
                maximum,
            }) if observed == output_length as u64 && maximum == (output_length - 1) as u64
        ));
    }

    #[test]
    fn retained_wire_limits_reject_tight_input_and_output_before_parse() {
        let archive = litchi_iwa_core::Limits::default()
            .with_archive_bytes(32)
            .expect("archive bytes")
            .with_message_bytes(8)
            .expect("message bytes")
            .with_header_fields(4)
            .expect("header fields");
        let physical = litchi_iwa_archive::Limits::default()
            .with_archive_limits(archive)
            .expect("physical archive limits");

        let exact = storage_wire_limits(physical, 8, 8).expect("exact profile is admitted");
        assert_eq!(exact.max_input_bytes(), 8);
        assert_eq!(exact.max_output_bytes(), 8);
        assert_eq!(exact.max_fields(), 4);

        assert!(matches!(
            storage_wire_limits(physical, 9, 8),
            Err(Error::Limit {
                kind: LimitKind::WireBytes,
                observed: 9,
                maximum: 8,
            })
        ));
        assert!(matches!(
            storage_wire_limits(physical, 8, 9),
            Err(Error::Limit {
                kind: LimitKind::WireBytes,
                observed: 9,
                maximum: 8,
            })
        ));
    }

    #[test]
    fn retained_rewrite_profile_reports_unsupported_nesting_as_a_limit() {
        let archive = litchi_iwa_core::Limits::default()
            .with_header_nesting(3)
            .expect("archive nesting");
        let physical = litchi_iwa_archive::Limits::default()
            .with_archive_limits(archive)
            .expect("physical archive limits");
        let error = crate::package::storage_rewrite_limits(physical).unwrap_err();
        assert!(matches!(
            map_storage_rewrite_limits_error(error),
            Error::Limit {
                kind: LimitKind::WireNesting,
                observed: 4,
                maximum: 3,
            }
        ));
    }

    #[test]
    fn unique_storage_candidate_preserves_unknown_wire_fields_and_reports_limits() {
        let mut payload = raw_length_delimited(3, b"abc");
        payload.extend_from_slice(&raw_varint(99, 7));
        let object = ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 9_999,
                data: payload,
            }],
            Default::default(),
        )
        .expect("test archive object");
        let admitted =
            RewriteLimits::new(1024, 8, 4, 8, 1024, 8, 8, 1024, 4096).expect("rewrite limits");
        assert_eq!(unique_body_message_index(&object, admitted).unwrap(), 0);

        let fields_limited = RewriteLimits::new(1024, 1, 4, 8, 1024, 8, 8, 1024, 4096)
            .expect("field-limited rewrite profile");
        assert!(matches!(
            unique_body_message_index(&object, fields_limited),
            Err(Error::Limit {
                kind: LimitKind::WireFields,
                ..
            })
        ));
    }
}
