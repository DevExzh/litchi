//! Exact-source lifecycle transactions for Pages body footnotes.
//!
//! The public surface is archive-free and selector-first.  The physical
//! insertion/removal implementation is intentionally kept in this private
//! module so native identifiers, component routes, and wire values cannot
//! escape through `litchi-pages`.

use std::{fmt, num::NonZeroU64, sync::Arc};

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
use thiserror::Error;

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
    fn new(source: &Package) -> Result<Self, BodyFootnoteError> {
        let limits = source.state.source.limits();
        let archive_limits = limits
            .effective_archive_limits()
            .map_err(map_archive_error)?;
        let max_total_bytes = limits.max_total_bytes();
        let max_iwa_stream_bytes = u64::try_from(limits.max_iwa_stream_bytes())
            .map_err(|_| BodyFootnoteError::InvalidSource)?;
        let max_wire_work = max_total_bytes
            .checked_mul(4)
            .ok_or(BodyFootnoteError::InvalidSource)?
            .max(max_iwa_stream_bytes);
        let mut budget = Self {
            max_input_bytes: limits.max_input_bytes(),
            max_output_bytes: limits.max_input_bytes(),
            max_entries: u64::try_from(limits.max_entries())
                .map_err(|_| BodyFootnoteError::InvalidSource)?,
            max_entry_bytes: limits.max_entry_bytes(),
            max_total_bytes,
            max_text_bytes: max_iwa_stream_bytes,
            max_text_units: max_iwa_stream_bytes,
            max_wire_bytes: max_iwa_stream_bytes,
            max_wire_fields: u64::try_from(archive_limits.max_header_fields())
                .map_err(|_| BodyFootnoteError::InvalidSource)?,
            max_wire_nesting: u64::try_from(archive_limits.max_header_nesting())
                .map_err(|_| BodyFootnoteError::InvalidSource)?,
            max_wire_work,
            max_references: u64::try_from(archive_limits.max_metadata_items())
                .map_err(|_| BodyFootnoteError::InvalidSource)?,
            max_components: u64::try_from(limits.max_entries())
                .map_err(|_| BodyFootnoteError::InvalidSource)?,
            max_registry_changes: u64::try_from(limits.max_entries())
                .map_err(|_| BodyFootnoteError::InvalidSource)?,
            max_retained_bytes: max_total_bytes,
            max_scratch_bytes: max_total_bytes,
            max_allocations: u64::try_from(limits.max_entries())
                .map_err(|_| BodyFootnoteError::InvalidSource)?,
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
            BodyFootnoteLimitKind::InputBytes,
            source.state.source.source_bytes().len(),
        )?;
        Ok(budget)
    }

    fn charge(
        &mut self,
        kind: BodyFootnoteLimitKind,
        amount: usize,
    ) -> Result<(), BodyFootnoteError> {
        let amount = u64::try_from(amount).map_err(|_| BodyFootnoteError::LimitExceeded {
            kind,
            observed: u64::MAX,
            maximum: self.maximum(kind),
        })?;
        self.charge_u64(kind, amount)
    }

    fn charge_u64(
        &mut self,
        kind: BodyFootnoteLimitKind,
        amount: u64,
    ) -> Result<(), BodyFootnoteError> {
        let current = *self.slot_mut(kind);
        let maximum = self.maximum(kind);
        let observed = current
            .checked_add(amount)
            .ok_or(BodyFootnoteError::LimitExceeded {
                kind,
                observed: u64::MAX,
                maximum,
            })?;
        if observed > maximum {
            return Err(BodyFootnoteError::LimitExceeded {
                kind,
                observed,
                maximum,
            });
        }
        *self.slot_mut(kind) = observed;
        Ok(())
    }

    fn observe(
        &mut self,
        kind: BodyFootnoteLimitKind,
        amount: usize,
    ) -> Result<(), BodyFootnoteError> {
        let amount = u64::try_from(amount).map_err(|_| BodyFootnoteError::LimitExceeded {
            kind,
            observed: u64::MAX,
            maximum: self.maximum(kind),
        })?;
        let current = *self.slot_mut(kind);
        if amount > current {
            let maximum = self.maximum(kind);
            if amount > maximum {
                return Err(BodyFootnoteError::LimitExceeded {
                    kind,
                    observed: amount,
                    maximum,
                });
            }
            *self.slot_mut(kind) = amount;
        }
        Ok(())
    }

    fn maximum(self, kind: BodyFootnoteLimitKind) -> u64 {
        match kind {
            BodyFootnoteLimitKind::InputBytes => self.max_input_bytes,
            BodyFootnoteLimitKind::OutputBytes => self.max_output_bytes,
            BodyFootnoteLimitKind::Entries => self.max_entries,
            BodyFootnoteLimitKind::EntryBytes => self.max_entry_bytes,
            BodyFootnoteLimitKind::TotalBytes => self.max_total_bytes,
            BodyFootnoteLimitKind::TextBytes => self.max_text_bytes,
            BodyFootnoteLimitKind::TextUnits => self.max_text_units,
            BodyFootnoteLimitKind::WireBytes => self.max_wire_bytes,
            BodyFootnoteLimitKind::WireFields => self.max_wire_fields,
            BodyFootnoteLimitKind::WireNesting => self.max_wire_nesting,
            BodyFootnoteLimitKind::WireWork => self.max_wire_work,
            BodyFootnoteLimitKind::References => self.max_references,
            BodyFootnoteLimitKind::Components => self.max_components,
            BodyFootnoteLimitKind::RegistryChanges => self.max_registry_changes,
            BodyFootnoteLimitKind::RetainedBytes => self.max_retained_bytes,
            BodyFootnoteLimitKind::ScratchBytes => self.max_scratch_bytes,
            BodyFootnoteLimitKind::Allocations => self.max_allocations,
        }
    }

    fn slot_mut(&mut self, kind: BodyFootnoteLimitKind) -> &mut u64 {
        match kind {
            BodyFootnoteLimitKind::InputBytes => &mut self.input_bytes,
            BodyFootnoteLimitKind::OutputBytes => &mut self.output_bytes,
            BodyFootnoteLimitKind::Entries => &mut self.entries,
            BodyFootnoteLimitKind::EntryBytes => &mut self.entry_bytes,
            BodyFootnoteLimitKind::TotalBytes => &mut self.total_bytes,
            BodyFootnoteLimitKind::TextBytes => &mut self.text_bytes,
            BodyFootnoteLimitKind::TextUnits => &mut self.text_units,
            BodyFootnoteLimitKind::WireBytes => &mut self.wire_bytes,
            BodyFootnoteLimitKind::WireFields => &mut self.wire_fields,
            BodyFootnoteLimitKind::WireNesting => &mut self.wire_nesting,
            BodyFootnoteLimitKind::WireWork => &mut self.wire_work,
            BodyFootnoteLimitKind::References => &mut self.references,
            BodyFootnoteLimitKind::Components => &mut self.components,
            BodyFootnoteLimitKind::RegistryChanges => &mut self.registry_changes,
            BodyFootnoteLimitKind::RetainedBytes => &mut self.retained_bytes,
            BodyFootnoteLimitKind::ScratchBytes => &mut self.scratch_bytes,
            BodyFootnoteLimitKind::Allocations => &mut self.allocations,
        }
    }

    fn charge_text_requirements(
        &mut self,
        requirements: litchi_iwa_text_wire::StorageRewriteExecutionRequirements,
    ) -> Result<(), BodyFootnoteError> {
        self.charge(
            BodyFootnoteLimitKind::WireBytes,
            requirements.output_bytes(),
        )?;
        self.charge(
            BodyFootnoteLimitKind::Entries,
            requirements.retained_elements(),
        )?;
        self.charge(
            BodyFootnoteLimitKind::RetainedBytes,
            requirements.retained_bytes(),
        )?;
        self.charge(
            BodyFootnoteLimitKind::ScratchBytes,
            requirements.peak_scratch_bytes(),
        )?;
        self.charge(
            BodyFootnoteLimitKind::Allocations,
            requirements.allocations(),
        )?;
        self.charge(BodyFootnoteLimitKind::WireWork, requirements.work())?;
        self.charge(
            BodyFootnoteLimitKind::References,
            requirements.reference_occurrences(),
        )
    }

    fn charge_graph_decode(
        &mut self,
        report: graph_codec::DecodeReport,
    ) -> Result<(), BodyFootnoteError> {
        self.charge(BodyFootnoteLimitKind::WireBytes, report.input_bytes())?;
        self.charge(BodyFootnoteLimitKind::WireFields, report.fields())?;
        self.charge(BodyFootnoteLimitKind::WireWork, report.work_bytes())?;
        self.observe(
            BodyFootnoteLimitKind::WireNesting,
            usize::try_from(report.max_depth()).map_err(|_| BodyFootnoteError::InvalidSource)?,
        )?;
        self.charge(BodyFootnoteLimitKind::Entries, report.entries())?;
        self.charge(
            BodyFootnoteLimitKind::RetainedBytes,
            report.retained_bytes(),
        )?;
        self.charge(BodyFootnoteLimitKind::ScratchBytes, report.scratch_bytes())
    }

    fn charge_graph_rewrite(
        &mut self,
        report: graph_codec::BodyFootnoteTableRewriteReport,
    ) -> Result<(), BodyFootnoteError> {
        self.charge_graph_decode(report.source())?;
        self.charge_graph_decode(report.result())?;
        self.charge(BodyFootnoteLimitKind::WireBytes, report.output_bytes())?;
        self.charge(BodyFootnoteLimitKind::WireWork, report.rewrite_work_bytes())?;
        self.charge(BodyFootnoteLimitKind::Entries, report.entries_after())?;
        self.charge(
            BodyFootnoteLimitKind::RetainedBytes,
            report.retained_bytes(),
        )?;
        self.charge(BodyFootnoteLimitKind::ScratchBytes, report.scratch_bytes())?;
        self.charge(BodyFootnoteLimitKind::Allocations, report.allocations())
    }

    fn charge_graph_encode(
        &mut self,
        report: graph_codec::GraphEncodeReport,
    ) -> Result<(), BodyFootnoteError> {
        self.charge(BodyFootnoteLimitKind::WireBytes, report.output_bytes())?;
        self.charge(BodyFootnoteLimitKind::WireFields, report.fields())?;
        self.charge(BodyFootnoteLimitKind::WireWork, report.work_bytes())?;
        self.charge(
            BodyFootnoteLimitKind::RetainedBytes,
            report.retained_bytes(),
        )?;
        self.charge(BodyFootnoteLimitKind::ScratchBytes, report.scratch_bytes())?;
        self.charge(BodyFootnoteLimitKind::Allocations, report.allocations())
    }

    fn charge_metadata_prepare(
        &mut self,
        report: package_metadata_codec::RewriteReport,
    ) -> Result<(), BodyFootnoteError> {
        self.charge(BodyFootnoteLimitKind::InputBytes, report.input_bytes())?;
        self.charge(BodyFootnoteLimitKind::WireFields, report.fields())?;
        self.charge(BodyFootnoteLimitKind::WireWork, report.work_bytes())?;
        self.observe(
            BodyFootnoteLimitKind::WireNesting,
            usize::try_from(report.max_depth()).map_err(|_| BodyFootnoteError::InvalidSource)?,
        )?;
        self.charge(
            BodyFootnoteLimitKind::Components,
            report.components_scanned(),
        )?;
        self.charge(
            BodyFootnoteLimitKind::References,
            report.references_scanned(),
        )?;
        self.charge(
            BodyFootnoteLimitKind::RegistryChanges,
            report
                .additions()
                .checked_add(report.removals())
                .ok_or(BodyFootnoteError::InvalidSource)?,
        )
    }

    fn charge_metadata_requirements(
        &mut self,
        requirements: package_metadata_codec::RewriteExecutionRequirements,
    ) -> Result<(), BodyFootnoteError> {
        self.charge(
            BodyFootnoteLimitKind::OutputBytes,
            requirements.output_bytes(),
        )?;
        self.charge(BodyFootnoteLimitKind::WireFields, requirements.fields())?;
        self.charge(BodyFootnoteLimitKind::WireWork, requirements.work_bytes())?;
        self.charge(BodyFootnoteLimitKind::Components, requirements.components())?;
        self.charge(BodyFootnoteLimitKind::References, requirements.references())?;
        self.charge(
            BodyFootnoteLimitKind::Allocations,
            requirements.allocations(),
        )?;
        self.charge(
            BodyFootnoteLimitKind::RetainedBytes,
            requirements.retained_bytes(),
        )?;
        self.charge(
            BodyFootnoteLimitKind::ScratchBytes,
            requirements.scratch_bytes(),
        )
    }

    fn charge_archive_output(
        &mut self,
        encoded_bytes: usize,
        maximum_compressed_bytes: usize,
    ) -> Result<(), BodyFootnoteError> {
        self.charge(BodyFootnoteLimitKind::EntryBytes, encoded_bytes)?;
        self.charge(BodyFootnoteLimitKind::TotalBytes, encoded_bytes)?;
        self.charge(
            BodyFootnoteLimitKind::ScratchBytes,
            encoded_bytes
                .checked_add(maximum_compressed_bytes)
                .ok_or(BodyFootnoteError::InvalidSource)?,
        )?;
        self.charge(
            BodyFootnoteLimitKind::RetainedBytes,
            maximum_compressed_bytes,
        )?;
        self.charge(BodyFootnoteLimitKind::Allocations, 2)
    }

    fn charge_reassembly(
        &mut self,
        requirements: ReassemblyExecutionRequirements,
    ) -> Result<(), BodyFootnoteError> {
        self.charge(
            BodyFootnoteLimitKind::OutputBytes,
            requirements.output_bytes(),
        )?;
        self.charge(
            BodyFootnoteLimitKind::TotalBytes,
            requirements.output_bytes(),
        )?;
        self.charge(BodyFootnoteLimitKind::Entries, requirements.offset_count())?;
        self.charge(
            BodyFootnoteLimitKind::ScratchBytes,
            requirements.scratch_bytes(),
        )?;
        self.charge(
            BodyFootnoteLimitKind::RetainedBytes,
            requirements.retained_bytes(),
        )?;
        self.charge(
            BodyFootnoteLimitKind::Allocations,
            requirements.allocations(),
        )
    }
}

/// A finite resource governed while one Pages body-footnote graph is changed.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BodyFootnoteLimitKind {
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

impl fmt::Display for BodyFootnoteLimitKind {
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
#[derive(Debug, Clone, Copy, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum BodyFootnoteError {
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
    /// The selected graph is malformed or ambiguous.
    #[error("the Pages body-footnote graph cannot be edited safely")]
    InvalidSource,
    /// A valid but unsupported dependency owns part of the graph.
    #[error("the Pages body-footnote graph has an unsupported dependency")]
    UnsupportedDependency,
    /// A finite transaction resource ceiling was exceeded.
    #[error("Pages body-footnote {kind} limit exceeded: observed {observed}, maximum {maximum}")]
    LimitExceeded {
        /// Resource category.
        kind: BodyFootnoteLimitKind,
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

/// One mutable body-footnote lifecycle edit staged against an immutable
/// package snapshot.
pub struct BodyFootnoteEdit<'a> {
    source: &'a Package,
    position: Position,
    before: Footnote,
    after: Option<Footnote>,
}

impl fmt::Debug for BodyFootnoteEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyFootnoteEdit")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl<'a> BodyFootnoteEdit<'a> {
    fn new(source: &'a Package, selector: Selector) -> Result<Self, BodyFootnoteError> {
        let footnotes = source
            .body_footnotes()
            .map_err(|_error| BodyFootnoteError::InvalidSource)?;
        let before = match selector {
            Selector::Index(index) => footnotes.get(index),
            Selector::At(position) => footnotes.iter().find(|value| value.position == position),
        }
        .ok_or(BodyFootnoteError::NotFound)?
        .clone();
        let position = before.position;
        Ok(Self {
            source,
            position,
            after: Some(before.clone()),
            before,
        })
    }

    /// Return the checked UTF-16 anchor position resolved at edit start.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Borrow the original semantic footnote.
    #[must_use]
    pub const fn before(&self) -> &Footnote {
        &self.before
    }

    /// Borrow the staged semantic result. `None` represents graph removal.
    #[must_use]
    pub const fn after(&self) -> Option<&Footnote> {
        self.after.as_ref()
    }

    /// Stage replacement of the selected footnote text.
    pub fn set(&mut self, text: &str) -> Result<&mut Self, BodyFootnoteError> {
        validate_authored(text, MAX_TEXT_BYTES, BodyFootnoteError::TextTooLarge)?;
        let Some(after) = self.after.as_mut() else {
            return Err(BodyFootnoteError::UnsupportedDependency);
        };
        after.text = try_boxed(text)?;
        Ok(self)
    }

    /// Stage replacement/removal of the optional custom marker.
    pub fn set_custom_mark(
        &mut self,
        custom_mark: Option<&str>,
    ) -> Result<&mut Self, BodyFootnoteError> {
        if let Some(value) = custom_mark {
            validate_authored(
                value,
                MAX_CUSTOM_MARK_BYTES,
                BodyFootnoteError::CustomMarkTooLarge,
            )?;
        }
        let Some(after) = self.after.as_mut() else {
            return Err(BodyFootnoteError::UnsupportedDependency);
        };
        after.custom_mark = custom_mark.map(try_boxed).transpose()?;
        Ok(self)
    }

    /// Stage complete removal of the selected body-footnote graph.
    pub fn clear(&mut self) -> &mut Self {
        self.after = None;
        self
    }

    /// Validate and atomically publish the staged lifecycle change.
    pub fn commit(&self) -> Result<BodyFootnoteCommit, BodyFootnoteError> {
        self.source.validate().map_err(map_package_error)?;
        let current = current_footnote(self.source, self.position)?;
        if current != self.before {
            return Err(BodyFootnoteError::PatchConflict);
        }
        let Some(after) = self.after.as_ref() else {
            return publish_rewrite(self.source, self.position, None, Some(&self.before));
        };
        if *after == self.before {
            return noop_commit(self.source, self.position, Some(&self.before));
        }
        let mut text_edit = self
            .source
            .edit_body_footnote_text(Selector::At(self.position))
            .map_err(map_text_error)?;
        text_edit.set(&after.text).map_err(map_text_error)?;
        text_edit
            .set_custom_mark(after.custom_mark.as_deref())
            .map_err(map_text_error)?;
        let text_commit = text_edit.commit().map_err(map_text_error)?;
        let source_bytes: Arc<[u8]> = self.source.state.source.shared_source();
        let target_bytes = text_commit.package().state.source.shared_source();
        let source_fingerprint = super::section_transaction::fingerprint(&source_bytes);
        let target_fingerprint = super::section_transaction::fingerprint(&target_bytes);
        let target_sequence = footnote_sequence(text_commit.package())?;
        let package = text_commit.into_package();
        Ok(BodyFootnoteCommit {
            package,
            patch: BodyFootnotePatch {
                source_bytes,
                target_bytes,
                source_sequence: footnote_sequence(self.source)?,
                target_sequence,
                source_fingerprint,
                target_fingerprint,
                position: self.position,
                before: Some(self.before.clone()),
                after: Some(after.clone()),
            },
            diagnostics: BodyFootnoteDiagnostics::published(1),
        })
    }
}

/// An exact-source-checked reversible body-footnote lifecycle patch.
#[derive(Clone, PartialEq, Eq)]
pub struct BodyFootnotePatch {
    pub(super) source_bytes: Arc<[u8]>,
    pub(super) target_bytes: Arc<[u8]>,
    source_sequence: Arc<[Footnote]>,
    target_sequence: Arc<[Footnote]>,
    pub(super) source_fingerprint: u64,
    pub(super) target_fingerprint: u64,
    position: Position,
    before: Option<Footnote>,
    after: Option<Footnote>,
}

impl fmt::Debug for BodyFootnotePatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("BodyFootnotePatch")
            .field("position", &self.position)
            .finish_non_exhaustive()
    }
}

impl BodyFootnotePatch {
    /// Return the UTF-16 anchor position associated with the operation.
    #[must_use]
    pub const fn position(&self) -> Position {
        self.position
    }

    /// Borrow the optional semantic source value.
    #[must_use]
    pub const fn before(&self) -> Option<&Footnote> {
        self.before.as_ref()
    }

    /// Borrow the optional semantic target value.
    #[must_use]
    pub const fn after(&self) -> Option<&Footnote> {
        self.after.as_ref()
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
            before: self.after.clone(),
            after: self.before.clone(),
        }
    }
}

/// Compact evidence for one body-footnote lifecycle commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct BodyFootnoteDiagnostics {
    changed: bool,
    touched_components: usize,
    full_reparse_performed: bool,
}

impl BodyFootnoteDiagnostics {
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
pub struct BodyFootnoteCommit {
    package: Package,
    patch: BodyFootnotePatch,
    diagnostics: BodyFootnoteDiagnostics,
}

impl BodyFootnoteCommit {
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

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &BodyFootnotePatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &BodyFootnoteDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Stage a selector-first edit of one existing body footnote.
    pub fn edit_body_footnote(
        &self,
        selector: Selector,
    ) -> Result<BodyFootnoteEdit<'_>, BodyFootnoteError> {
        BodyFootnoteEdit::new(self, selector)
    }

    /// Insert one body footnote at a checked UTF-16 position.
    pub fn insert_body_footnote(
        &self,
        position: Position,
        text: impl AsRef<str>,
        custom_mark: Option<&str>,
    ) -> Result<BodyFootnoteCommit, BodyFootnoteError> {
        validate_authored(
            text.as_ref(),
            MAX_TEXT_BYTES,
            BodyFootnoteError::TextTooLarge,
        )?;
        if let Some(value) = custom_mark {
            validate_authored(
                value,
                MAX_CUSTOM_MARK_BYTES,
                BodyFootnoteError::CustomMarkTooLarge,
            )?;
        }
        insert_lifecycle(self, position, text.as_ref(), custom_mark)
    }

    /// Apply an exact-source-checked body-footnote lifecycle patch.
    pub fn apply_body_footnote(
        &self,
        patch: &BodyFootnotePatch,
    ) -> Result<BodyFootnoteCommit, BodyFootnoteError> {
        let source = &self.state.source;
        if super::section_transaction::fingerprint(source.source_bytes())
            != patch.source_fingerprint
            || source.source_bytes() != patch.source_bytes.as_ref()
        {
            return Err(BodyFootnoteError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(BodyFootnoteCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyFootnoteDiagnostics::unchanged(),
            });
        }
        self.validate().map_err(map_package_error)?;
        let current_sequence = footnote_sequence(self)?;
        if current_sequence.as_ref() != patch.source_sequence.as_ref() {
            return Err(BodyFootnoteError::PatchConflict);
        }
        if patch.is_noop() {
            return Ok(BodyFootnoteCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: BodyFootnoteDiagnostics::unchanged(),
            });
        }
        if !self.state.source.source_is_exact()
            || super::section_transaction::fingerprint(&patch.target_bytes)
                != patch.target_fingerprint
        {
            return Err(BodyFootnoteError::PatchConflict);
        }
        let mut budget = TransactionBudget::new(self)?;
        budget.charge(BodyFootnoteLimitKind::OutputBytes, patch.target_bytes.len())?;
        budget.charge(
            BodyFootnoteLimitKind::RetainedBytes,
            patch.target_bytes.len(),
        )?;
        budget.charge(BodyFootnoteLimitKind::Allocations, 1)?;
        let candidate_source = SourceCatalog::from_shared_bytes_with_limits(
            Arc::clone(&patch.target_bytes),
            self.state.source.limits(),
        )
        .map_err(map_archive_error)?;
        let candidate =
            Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
        let target_sequence = footnote_sequence(&candidate)?;
        if target_sequence.as_ref() != patch.target_sequence.as_ref() {
            return Err(BodyFootnoteError::Verification);
        }
        Ok(BodyFootnoteCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: BodyFootnoteDiagnostics::published(1),
        })
    }
}

fn validate_authored(
    value: &str,
    maximum: usize,
    overflow: BodyFootnoteError,
) -> Result<(), BodyFootnoteError> {
    if value.len() > maximum {
        return Err(overflow);
    }
    if value.contains(['\u{000e}', '\u{fffc}']) {
        return Err(BodyFootnoteError::StructuralMarker);
    }
    Ok(())
}

fn try_boxed(value: &str) -> Result<Box<str>, BodyFootnoteError> {
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_error| BodyFootnoteError::Allocation {
            amount: value.len(),
        })?;
    owned.push_str(value);
    Ok(owned.into_boxed_str())
}

fn current_footnote(package: &Package, position: Position) -> Result<Footnote, BodyFootnoteError> {
    current_footnote_at_optional(package, position)?.ok_or(BodyFootnoteError::NotFound)
}

fn current_footnote_at_optional(
    package: &Package,
    position: Position,
) -> Result<Option<Footnote>, BodyFootnoteError> {
    package
        .body_footnotes()
        .map_err(map_package_error)
        .map(|values| values.into_iter().find(|value| value.position == position))
}

fn footnote_sequence(package: &Package) -> Result<Arc<[Footnote]>, BodyFootnoteError> {
    package
        .body_footnotes()
        .map_err(map_package_error)
        .map(Into::into)
}

fn noop_commit(
    package: &Package,
    position: Position,
    value: Option<&Footnote>,
) -> Result<BodyFootnoteCommit, BodyFootnoteError> {
    let source: Arc<[u8]> = package.state.source.shared_source();
    let fingerprint = super::section_transaction::fingerprint(&source);
    Ok(BodyFootnoteCommit {
        package: package.snapshot(),
        patch: BodyFootnotePatch {
            source_bytes: Arc::clone(&source),
            target_bytes: source,
            source_sequence: footnote_sequence(package)?,
            target_sequence: footnote_sequence(package)?,
            source_fingerprint: fingerprint,
            target_fingerprint: fingerprint,
            position,
            before: value.cloned(),
            after: value.cloned(),
        },
        diagnostics: BodyFootnoteDiagnostics::unchanged(),
    })
}

fn publish_rewrite(
    source: &Package,
    position: Position,
    after: Option<&Footnote>,
    before: Option<&Footnote>,
) -> Result<BodyFootnoteCommit, BodyFootnoteError> {
    let target = rewrite_remove(source, position)?;
    let source_bytes: Arc<[u8]> = source.state.source.shared_source();
    let target_bytes = target.state.source.shared_source();
    let target_sequence = footnote_sequence(&target)?;
    Ok(BodyFootnoteCommit {
        package: target,
        patch: BodyFootnotePatch {
            source_fingerprint: super::section_transaction::fingerprint(&source_bytes),
            target_fingerprint: super::section_transaction::fingerprint(&target_bytes),
            source_bytes,
            target_bytes,
            source_sequence: footnote_sequence(source)?,
            target_sequence,
            position,
            before: before.cloned(),
            after: after.cloned(),
        },
        diagnostics: BodyFootnoteDiagnostics::published(1),
    })
}

fn map_text_error(error: super::FootnoteTextError) -> BodyFootnoteError {
    match error {
        super::FootnoteTextError::NotFound => BodyFootnoteError::NotFound,
        super::FootnoteTextError::StructuralMarker => BodyFootnoteError::StructuralMarker,
        super::FootnoteTextError::TextTooLarge => BodyFootnoteError::TextTooLarge,
        super::FootnoteTextError::CustomMarkTooLarge => BodyFootnoteError::CustomMarkTooLarge,
        super::FootnoteTextError::UnsupportedSource => BodyFootnoteError::UnsupportedSource,
        super::FootnoteTextError::InvalidSource => BodyFootnoteError::InvalidSource,
        super::FootnoteTextError::Verification => BodyFootnoteError::Verification,
        super::FootnoteTextError::PatchConflict => BodyFootnoteError::PatchConflict,
        super::FootnoteTextError::Allocation { amount } => BodyFootnoteError::Allocation { amount },
        super::FootnoteTextError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => BodyFootnoteError::LimitExceeded {
            kind: match kind {
                super::FootnoteTextLimitKind::InputBytes => BodyFootnoteLimitKind::InputBytes,
                super::FootnoteTextLimitKind::OutputBytes => BodyFootnoteLimitKind::OutputBytes,
                super::FootnoteTextLimitKind::Entries => BodyFootnoteLimitKind::Entries,
                super::FootnoteTextLimitKind::EntryBytes => BodyFootnoteLimitKind::EntryBytes,
                super::FootnoteTextLimitKind::TotalBytes => BodyFootnoteLimitKind::TotalBytes,
                super::FootnoteTextLimitKind::TextBytes => BodyFootnoteLimitKind::TextBytes,
                super::FootnoteTextLimitKind::TextUnits => BodyFootnoteLimitKind::TextUnits,
                super::FootnoteTextLimitKind::WireBytes => BodyFootnoteLimitKind::WireBytes,
                super::FootnoteTextLimitKind::WireFields => BodyFootnoteLimitKind::WireFields,
                super::FootnoteTextLimitKind::WireNesting => BodyFootnoteLimitKind::WireNesting,
                super::FootnoteTextLimitKind::WireWork => BodyFootnoteLimitKind::WireWork,
            },
            observed,
            maximum,
        },
    }
}

fn map_package_error(error: super::PackageError) -> BodyFootnoteError {
    match error {
        super::PackageError::Archive(error) => map_archive_error(error),
        super::PackageError::Allocation { amount } => BodyFootnoteError::Allocation { amount },
        super::PackageError::PayloadLimit { observed, limit }
        | super::PackageError::ObjectLimit { observed, limit }
        | super::PackageError::SectionNamesTooLarge { observed, limit } => {
            BodyFootnoteError::LimitExceeded {
                kind: BodyFootnoteLimitKind::Entries,
                observed: observed as u64,
                maximum: limit as u64,
            }
        },
        _ => BodyFootnoteError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> BodyFootnoteError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyFootnoteError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => BodyFootnoteLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => BodyFootnoteLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => BodyFootnoteLimitKind::Entries,
                litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    BodyFootnoteLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes
                | litchi_iwa_archive::LimitKind::IwaTotalBytes => BodyFootnoteLimitKind::TotalBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            BodyFootnoteError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        litchi_iwa_archive::Error::Io(_)
        | litchi_iwa_archive::Error::Zip { .. }
        | litchi_iwa_archive::Error::InvalidLimits(_)
        | litchi_iwa_archive::Error::Encrypted
        | litchi_iwa_archive::Error::SourceChanged { .. }
        | litchi_iwa_archive::Error::DirectoryChanged { .. }
        | litchi_iwa_archive::Error::Reassembly(_)
        | litchi_iwa_archive::Error::InvalidBundle(_) => BodyFootnoteError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> BodyFootnoteError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => BodyFootnoteError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => BodyFootnoteLimitKind::Entries,
                litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes
                | litchi_iwa_core::LimitKind::SnappyFrames => BodyFootnoteLimitKind::EntryBytes,
                litchi_iwa_core::LimitKind::HeaderFields => BodyFootnoteLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => BodyFootnoteLimitKind::WireNesting,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            BodyFootnoteError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::InvalidArchive { .. }
        | litchi_iwa_core::Error::InvalidLimits { .. }
        | litchi_iwa_core::Error::HeaderCodec { .. }
        | litchi_iwa_core::Error::Io(_)
        | litchi_iwa_core::Error::Snappy { .. } => BodyFootnoteError::InvalidSource,
    }
}

fn map_graph_error(error: graph_codec::DecodeError) -> BodyFootnoteError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            graph_codec::DecodeLimit::InputBytes { observed, maximum } => {
                (BodyFootnoteLimitKind::InputBytes, observed, maximum)
            },
            graph_codec::DecodeLimit::OutputBytes { observed, maximum } => {
                (BodyFootnoteLimitKind::OutputBytes, observed, maximum)
            },
            graph_codec::DecodeLimit::Fields { observed, maximum } => {
                (BodyFootnoteLimitKind::WireFields, observed, maximum)
            },
            graph_codec::DecodeLimit::WorkBytes { observed, maximum } => {
                (BodyFootnoteLimitKind::WireWork, observed, maximum)
            },
            graph_codec::DecodeLimit::Nesting { observed, maximum } => {
                return BodyFootnoteError::LimitExceeded {
                    kind: BodyFootnoteLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                };
            },
            graph_codec::DecodeLimit::Entries { observed, maximum } => {
                (BodyFootnoteLimitKind::Entries, observed, maximum)
            },
            graph_codec::DecodeLimit::TextBytes { observed, maximum } => {
                (BodyFootnoteLimitKind::TextBytes, observed, maximum)
            },
            _ => return BodyFootnoteError::InvalidSource,
        };
        return BodyFootnoteError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return BodyFootnoteError::Allocation { amount };
    }
    BodyFootnoteError::InvalidSource
}

fn map_text_wire_error(error: litchi_iwa_text_wire::RewriteError) -> BodyFootnoteError {
    match error {
        litchi_iwa_text_wire::RewriteError::LimitExceeded {
            resource,
            observed,
            limit,
        } => BodyFootnoteError::LimitExceeded {
            kind: match resource {
                "output bytes" => BodyFootnoteLimitKind::OutputBytes,
                "fields" => BodyFootnoteLimitKind::WireFields,
                "rewrite work" => BodyFootnoteLimitKind::WireWork,
                "text bytes" => BodyFootnoteLimitKind::TextBytes,
                "text units" => BodyFootnoteLimitKind::TextUnits,
                _ => BodyFootnoteLimitKind::WireBytes,
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_text_wire::RewriteError::Allocation { amount, .. } => {
            BodyFootnoteError::Allocation { amount }
        },
        litchi_iwa_text_wire::RewriteError::RangeOutOfBounds { .. }
        | litchi_iwa_text_wire::RewriteError::SurrogateSplit { .. } => {
            BodyFootnoteError::PositionOutOfBounds
        },
        litchi_iwa_text_wire::RewriteError::ReversedRange { .. }
        | litchi_iwa_text_wire::RewriteError::ArithmeticOverflow { .. }
        | litchi_iwa_text_wire::RewriteError::InvalidLimit { .. }
        | litchi_iwa_text_wire::RewriteError::InvalidFormat(_)
        | litchi_iwa_text_wire::RewriteError::Projection(_) => BodyFootnoteError::InvalidSource,
        _ => BodyFootnoteError::InvalidSource,
    }
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

impl PackageMetadataVisitor for MetadataFactsVisitor {
    fn visit_component(
        &mut self,
        component: package_metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.components.push(MetadataComponent {
            identifier: component.identifier(),
            preferred: component.preferred_locator().to_owned(),
            locator: component.locator().map(str::to_owned),
            current: component.is_current(),
        });
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.uuids.push(MetadataUuid {
            component: binding.component().identifier(),
            object: binding.object_identifier(),
            uuid: binding.uuid(),
            current: binding.component().is_current(),
        });
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.externals.push(MetadataExternal {
            source: reference.source().identifier(),
            target: reference.target_component_identifier(),
            object: reference.object_identifier(),
            weak: reference.is_weak(),
            current: reference.source().is_current() && !reference.is_versioned(),
        });
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.data_owners.push(MetadataDataOwner {
            component: owner.component().identifier(),
            data: owner.data_identifier(),
            object: owner.object_identifier(),
            count: owner.count(),
            current: owner.component().is_current(),
            unknown: owner.has_unknown_fields(),
        });
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        component: package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), MetadataRewriteError> {
        self.ambiguous.push((component.identifier(), identifier));
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        has_unknown_fields: bool,
    ) -> Result<(), MetadataRewriteError> {
        self.data_metadata_maps
            .push((object_identifier, has_unknown_fields));
        Ok(())
    }
}

struct MetadataLocation {
    member: String,
    archive: Archive,
    object_index: usize,
    message_index: usize,
    payload: Vec<u8>,
}

struct MetadataRewrite {
    member: String,
    compressed: Vec<u8>,
}

fn metadata_location(source: &Package) -> Result<MetadataLocation, BodyFootnoteError> {
    let component = source
        .state
        .source
        .components()
        .get("Index/Metadata.iwa")
        .ok_or(BodyFootnoteError::UnsupportedDependency)?;
    let archive = component.archive().clone();
    let mut location = None;
    for (object_index, object) in archive.objects.iter().enumerate() {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ == METADATA_MESSAGE_TYPE {
                if location.replace((object_index, message_index)).is_some() {
                    return Err(BodyFootnoteError::InvalidSource);
                }
            }
        }
    }
    let (object_index, message_index) = location.ok_or(BodyFootnoteError::InvalidSource)?;
    Ok(MetadataLocation {
        member: component.name().to_owned(),
        payload: archive.objects[object_index].messages[message_index]
            .data
            .clone(),
        archive,
        object_index,
        message_index,
    })
}

fn metadata_options(
    source: &Package,
    payload_len: usize,
) -> Result<MetadataRewriteOptions, BodyFootnoteError> {
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
    let work = input
        .checked_mul(128)
        .ok_or(BodyFootnoteError::InvalidSource)?
        .max(1);
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
    location: &MetadataLocation,
    budget: &mut TransactionBudget,
) -> Result<(MetadataFactsVisitor, u64), BodyFootnoteError> {
    let options = metadata_options(source, location.payload.len())?;
    let mut visitor = MetadataFactsVisitor::default();
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        &location.payload,
        options,
        &mut visitor,
    )
    .map_err(map_metadata_error)?;
    budget.charge_metadata_prepare(inspection.report())?;
    Ok((visitor, inspection.last_object_identifier()))
}

fn map_metadata_error(error: MetadataRewriteError) -> BodyFootnoteError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            package_metadata_codec::RewriteLimit::InputBytes { observed, maximum } => {
                (BodyFootnoteLimitKind::InputBytes, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => {
                (BodyFootnoteLimitKind::OutputBytes, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Fields { observed, maximum } => {
                (BodyFootnoteLimitKind::WireFields, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (BodyFootnoteLimitKind::WireWork, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => {
                return BodyFootnoteError::LimitExceeded {
                    kind: BodyFootnoteLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                };
            },
            package_metadata_codec::RewriteLimit::Components { observed, maximum } => {
                (BodyFootnoteLimitKind::Components, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::References { observed, maximum } => {
                (BodyFootnoteLimitKind::References, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Additions { observed, maximum } => {
                (BodyFootnoteLimitKind::RegistryChanges, observed, maximum)
            },
            _ => return BodyFootnoteError::InvalidSource,
        };
        return BodyFootnoteError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(amount) = error.allocation_request() {
        return BodyFootnoteError::Allocation { amount };
    }
    match error.invalid_reason() {
        Some(package_metadata_codec::InvalidReason::VersionedRemoval)
        | Some(package_metadata_codec::InvalidReason::CrossComponentRemoval)
        | Some(package_metadata_codec::InvalidReason::ExistingReferenceCollision) => {
            BodyFootnoteError::UnsupportedDependency
        },
        _ => BodyFootnoteError::InvalidSource,
    }
}

fn component_for_member<'a>(
    facts: &'a MetadataFactsVisitor,
    member: &str,
) -> Result<&'a MetadataComponent, BodyFootnoteError> {
    let base = member
        .rsplit('/')
        .next()
        .and_then(|value| value.strip_suffix(".iwa"))
        .ok_or(BodyFootnoteError::InvalidSource)?;
    let mut found = None;
    for component in facts
        .components
        .iter()
        .filter(|component| component.current)
    {
        if component.effective() == base || component.preferred == base {
            if found.is_some() {
                return Err(BodyFootnoteError::InvalidSource);
            }
            found = Some(component);
        }
    }
    found.ok_or(BodyFootnoteError::UnsupportedDependency)
}

fn rewrite_metadata_archive(
    source: &Package,
    mut location: MetadataLocation,
    payload: Vec<u8>,
    budget: &mut TransactionBudget,
) -> Result<MetadataRewrite, BodyFootnoteError> {
    let archive_limits = source
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let object_identifier = location.archive.objects[location.object_index]
        .archive_info
        .identifier
        .ok_or(BodyFootnoteError::InvalidSource)?;
    let object = location
        .archive
        .object_mut(object_identifier)
        .ok_or(BodyFootnoteError::InvalidSource)?;
    object
        .replace_message_preserving_header_with_limits(
            location.message_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: payload,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_length = location
        .archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let maximum_compressed =
        SnappyStream::maximum_compressed_len(encoded_length).map_err(map_core_error)?;
    budget.charge_archive_output(encoded_length, maximum_compressed)?;
    let bytes = location
        .archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    if bytes.len() != encoded_length {
        return Err(BodyFootnoteError::Verification);
    }
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    if compressed.len() > maximum_compressed {
        return Err(BodyFootnoteError::Verification);
    }
    Ok(MetadataRewrite {
        member: location.member,
        compressed,
    })
}

fn metadata_addition(
    source: &Package,
    body_member: &str,
    storage_identifier: u64,
    marker_identifier: u64,
    budget: &mut TransactionBudget,
) -> Result<MetadataRewrite, BodyFootnoteError> {
    let location = metadata_location(source)?;
    let (facts, last_identifier) = metadata_facts(source, &location, budget)?;
    if marker_identifier <= last_identifier {
        return Err(BodyFootnoteError::InvalidSource);
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
                return Err(BodyFootnoteError::InvalidSource);
            }
        }
    }
    let mut uuid = UuidBits::new(
        storage_identifier ^ 0x9e37_79b9_7f4a_7c15,
        marker_identifier ^ 0xd1b5_4a32_d192_ed03,
    );
    while facts.uuids.iter().any(|binding| binding.uuid == uuid) {
        uuid = UuidBits::new(
            uuid.lower()
                .checked_add(1)
                .ok_or(BodyFootnoteError::InvalidSource)?,
            uuid.upper()
                .checked_add(0x9e37_79b9_7f4a_7c15)
                .ok_or(BodyFootnoteError::InvalidSource)?,
        );
        if facts.uuids.iter().any(|binding| binding.uuid == uuid) {
            return Err(BodyFootnoteError::UnsupportedDependency);
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
        &location.payload,
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
        return Err(BodyFootnoteError::Verification);
    }
    rewrite_metadata_archive(source, location, output.into_bytes(), budget)
}

fn metadata_removal(
    source: &Package,
    storage_identifier: u64,
    budget: &mut TransactionBudget,
) -> Result<MetadataRewrite, BodyFootnoteError> {
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
        return Err(BodyFootnoteError::UnsupportedDependency);
    }
    let matching_uuids: Vec<_> = facts
        .uuids
        .iter()
        .filter(|binding| binding.object == storage_identifier)
        .collect();
    let current_uuid = matching_uuids
        .iter()
        .copied()
        .find(|binding| binding.current)
        .ok_or(BodyFootnoteError::UnsupportedDependency)?;
    if matching_uuids.iter().any(|binding| !binding.current) {
        return Err(BodyFootnoteError::UnsupportedDependency);
    }
    if matching_uuids.len() != 1 {
        return Err(BodyFootnoteError::InvalidSource);
    }
    let target = facts
        .components
        .iter()
        .filter(|component| component.current && component.identifier == current_uuid.component)
        .collect::<Vec<_>>();
    if target.len() != 1 {
        return Err(BodyFootnoteError::InvalidSource);
    }
    let target_selector = ComponentSelector::new(target[0].identifier, target[0].effective());
    let uuid_removals = [ObjectUuidRemoval::new(
        target_selector,
        storage_identifier,
        current_uuid.uuid,
    )];

    let matching_external: Vec<_> = facts
        .externals
        .iter()
        .filter(|reference| {
            reference.target == current_uuid.component
                && reference.object == Some(storage_identifier)
        })
        .collect();
    if matching_external.iter().any(|reference| !reference.current) {
        return Err(BodyFootnoteError::UnsupportedDependency);
    }
    let matching_data: Vec<_> = facts
        .data_owners
        .iter()
        .filter(|owner| owner.object == storage_identifier)
        .collect();
    if matching_data
        .iter()
        .any(|owner| !owner.current || owner.unknown)
    {
        return Err(BodyFootnoteError::UnsupportedDependency);
    }

    let mut external_removals = Vec::new();
    let mut data_removals = Vec::new();
    let mut token_selectors = vec![target_selector];
    for reference in matching_external {
        let source_components: Vec<_> = facts
            .components
            .iter()
            .filter(|component| component.current && component.identifier == reference.source)
            .collect();
        if source_components.len() != 1 {
            return Err(BodyFootnoteError::InvalidSource);
        }
        let source_selector = ComponentSelector::new(
            source_components[0].identifier,
            source_components[0].effective(),
        );
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
    for owner in matching_data {
        let source_components: Vec<_> = facts
            .components
            .iter()
            .filter(|component| component.current && component.identifier == owner.component)
            .collect();
        if source_components.len() != 1 {
            return Err(BodyFootnoteError::InvalidSource);
        }
        let source_selector = ComponentSelector::new(
            source_components[0].identifier,
            source_components[0].effective(),
        );
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
        &location.payload,
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
        return Err(BodyFootnoteError::Verification);
    }
    rewrite_metadata_archive(source, location, output.into_bytes(), budget)
}

fn insert_lifecycle(
    source: &Package,
    position: Position,
    text: &str,
    custom_mark: Option<&str>,
) -> Result<BodyFootnoteCommit, BodyFootnoteError> {
    source.validate().map_err(map_package_error)?;
    let mut budget = TransactionBudget::new(source)?;
    let position_index = usize::try_from(position.utf16_index())
        .map_err(|_| BodyFootnoteError::PositionOutOfBounds)?;
    if current_footnote_at_optional(source, position)?.is_some() {
        return Err(BodyFootnoteError::PositionOccupied);
    }
    let source_catalog = &source.state.source;
    if !source_catalog.source_is_exact() {
        return Err(BodyFootnoteError::UnsupportedSource);
    }
    let body_identifier =
        super::root_references_with_limits(source_catalog.components(), source_catalog.limits())
            .map_err(map_package_error)?
            .body
            .ok_or(BodyFootnoteError::InvalidSource)?;
    let body_component = component_for_object(source, body_identifier.get())?;
    // Insertion changes the rooted body graph and publishes new metadata
    // registrations.  Prove every pre-existing footnote graph first, rather
    // than validating only the graph whose body table will be edited: an
    // alias in a sibling note would otherwise make the package unsafe after
    // the new candidate is published.
    validate_existing_footnote_graphs(source, &mut budget)?;
    let (body_archive, archive_limits) = editable_archive(source, &body_component, &mut budget)?;
    let body_object = body_archive
        .object(body_identifier.get())
        .ok_or(BodyFootnoteError::InvalidSource)?;
    let body_message_index = unique_body_message_index(body_object)?;
    let body_message = body_object.messages[body_message_index].clone();
    let position_u32 =
        u32::try_from(position_index).map_err(|_| BodyFootnoteError::PositionOutOfBounds)?;
    let max_identifier = max_object_identifier(source, &mut budget)?;
    let reference_identifier = NonZeroU64::new(
        max_identifier
            .checked_add(1)
            .ok_or(BodyFootnoteError::InvalidSource)?,
    )
    .ok_or(BodyFootnoteError::InvalidSource)?;
    let storage_identifier = NonZeroU64::new(
        max_identifier
            .checked_add(2)
            .ok_or(BodyFootnoteError::InvalidSource)?,
    )
    .ok_or(BodyFootnoteError::InvalidSource)?;
    let marker_identifier = NonZeroU64::new(
        max_identifier
            .checked_add(3)
            .ok_or(BodyFootnoteError::InvalidSource)?,
    )
    .ok_or(BodyFootnoteError::InvalidSource)?;
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
        .ok_or(BodyFootnoteError::InvalidSource)?;
    body_object
        .replace_message_preserving_header_with_limits(
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
    )?;
    let objects = graph_objects(graph_write, &payloads, archive_limits, &mut budget)?;
    budget.charge(BodyFootnoteLimitKind::Allocations, 1)?;
    archive
        .append_objects_with_limits(objects.to_vec(), archive_limits)
        .map_err(map_core_error)?;
    let body_compressed = compress_archive(archive, archive_limits, &mut budget)?;

    let metadata = metadata_addition(
        source,
        body_component.as_str(),
        storage_identifier.get(),
        marker_identifier.get(),
        &mut budget,
    )?;
    let mut edits = Vec::new();
    budget.charge(BodyFootnoteLimitKind::Allocations, 1)?;
    edits
        .try_reserve_exact(2)
        .map_err(|_| BodyFootnoteError::Allocation { amount: 2 })?;
    edits.push(EntryEdit::new(body_component.as_str(), &body_compressed));
    edits.push(EntryEdit::new(
        metadata.member.as_str(),
        &metadata.compressed,
    ));
    let previews = preview_names(source);
    let output = reassemble_candidate(source_catalog, &edits, &previews, &mut budget)?;
    budget.charge(BodyFootnoteLimitKind::Allocations, 1)?;
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), source_catalog.limits())
            .map_err(map_archive_error)?;
    let candidate = Package::from_source_catalog(candidate_source).map_err(map_package_error)?;
    let expected = Footnote::with_custom_mark(
        position,
        try_boxed(text)?,
        custom_mark.map(try_boxed).transpose()?,
    )
    .map_err(|_| BodyFootnoteError::InvalidSource)?;
    verify_lifecycle_target(&candidate, position, Some(&expected))?;
    let source_bytes: Arc<[u8]> = source_catalog.shared_source();
    let target_bytes = candidate.state.source.shared_source();
    let target_sequence = footnote_sequence(&candidate)?;
    let package = candidate;
    Ok(BodyFootnoteCommit {
        package,
        patch: BodyFootnotePatch {
            source_fingerprint: super::section_transaction::fingerprint(&source_bytes),
            target_fingerprint: super::section_transaction::fingerprint(&target_bytes),
            source_bytes,
            target_bytes,
            source_sequence: footnote_sequence(source)?,
            target_sequence,
            position,
            before: None,
            after: Some(expected),
        },
        diagnostics: BodyFootnoteDiagnostics::published(1),
    })
}

fn rewrite_remove(source: &Package, position: Position) -> Result<Package, BodyFootnoteError> {
    let mut budget = TransactionBudget::new(source)?;
    let graphs = super::footnote_text::native_footnotes(source).map_err(map_text_error)?;
    budget.charge(BodyFootnoteLimitKind::Entries, graphs.len())?;
    budget.charge(
        BodyFootnoteLimitKind::WireWork,
        source.state.source.source_bytes().len(),
    )?;
    let graph = graphs
        .into_iter()
        .find(|graph| graph.position == position)
        .ok_or(BodyFootnoteError::NotFound)?;
    let source_catalog = &source.state.source;
    let body_component = component_for_object(source, graph.body_identifier.get())?;
    let (mut archive, archive_limits) = editable_archive(source, &body_component, &mut budget)?;
    let body_object = archive
        .object(graph.body_identifier.get())
        .ok_or(BodyFootnoteError::InvalidSource)?;
    let body_message_index = unique_body_message_index(body_object)?;
    let body_message = body_object.messages[body_message_index].clone();
    let body_payload = rewrite_body_storage(
        &body_message.data,
        usize::try_from(position.utf16_index())
            .map_err(|_| BodyFootnoteError::PositionOutOfBounds)?,
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
        .ok_or(BodyFootnoteError::InvalidSource)?;
    body_object
        .replace_message_preserving_header_with_limits(
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
    )?;
    prove_exclusive_graph(
        source,
        graph.reference_identifier.get(),
        graph.storage_identifier.get(),
        graph.marker_identifier.get(),
        &mut budget,
    )?;
    for id in [
        graph.reference_identifier,
        graph.storage_identifier,
        graph.marker_identifier,
    ] {
        archive.remove_object(id.get());
    }
    let compressed = compress_archive(archive, archive_limits, &mut budget)?;
    let metadata = metadata_removal(source, graph.storage_identifier.get(), &mut budget)?;
    let mut edits = Vec::new();
    budget.charge(BodyFootnoteLimitKind::Allocations, 1)?;
    edits
        .try_reserve_exact(2)
        .map_err(|_| BodyFootnoteError::Allocation { amount: 2 })?;
    edits.push(EntryEdit::new(body_component.as_str(), &compressed));
    edits.push(EntryEdit::new(
        metadata.member.as_str(),
        &metadata.compressed,
    ));
    let previews = preview_names(source);
    let output = reassemble_candidate(source_catalog, &edits, &previews, &mut budget)?;
    budget.charge(BodyFootnoteLimitKind::Allocations, 1)?;
    let candidate_source =
        SourceCatalog::from_shared_bytes_with_limits(output.into(), source_catalog.limits())
            .map_err(map_archive_error)?;
    Package::from_source_catalog(candidate_source).map_err(map_package_error)
}

fn reassemble_candidate(
    source: &SourceCatalog,
    edits: &[EntryEdit<'_>],
    deleted_names: &[&str],
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, BodyFootnoteError> {
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
) -> Result<(), BodyFootnoteError> {
    let actual = current_footnote_at_optional(candidate, position)?;
    if actual.as_ref() != expected {
        return Err(BodyFootnoteError::Verification);
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

fn graph_options(
    source: &Package,
    payload: usize,
) -> Result<GraphDecodeOptions, BodyFootnoteError> {
    let limits = source.state.source.limits();
    let archive = limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let output = payload
        .checked_mul(8)
        .and_then(|value| value.checked_add(1024))
        .ok_or(BodyFootnoteError::InvalidSource)?
        .min(limits.max_iwa_stream_bytes())
        .max(1);
    let fields = archive
        .max_header_fields()
        .checked_mul(16)
        .ok_or(BodyFootnoteError::InvalidSource)?
        .max(1);
    let work = payload
        .checked_mul(32)
        .ok_or(BodyFootnoteError::InvalidSource)?
        .max(1);
    Ok(GraphDecodeOptions::new(
        payload.max(1),
        output,
        fields,
        work,
        u32::try_from(archive.max_header_nesting()).unwrap_or(u32::MAX),
        MAX_TEXT_BYTES.max(1),
    ))
}

fn component_for_object(package: &Package, identifier: u64) -> Result<String, BodyFootnoteError> {
    let mut found = None;
    for component in package.state.source.components().iter() {
        if component.archive().object(identifier).is_some() {
            if found.is_some() {
                return Err(BodyFootnoteError::InvalidSource);
            }
            found = Some(component.name().to_owned());
        }
    }
    found.ok_or(BodyFootnoteError::InvalidSource)
}

fn editable_archive(
    package: &Package,
    name: &str,
    budget: &mut TransactionBudget,
) -> Result<(Archive, litchi_iwa_core::Limits), BodyFootnoteError> {
    let entry = package
        .state
        .source
        .package()
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or(BodyFootnoteError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(BodyFootnoteError::UnsupportedSource);
    }
    let limits = package
        .state
        .source
        .limits()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        package
            .state
            .source
            .limits()
            .snappy_limits()
            .map_err(map_archive_error)?,
    )
    .map_err(map_core_error)?;
    budget.charge(BodyFootnoteLimitKind::EntryBytes, stream.as_bytes().len())?;
    budget.charge(BodyFootnoteLimitKind::TotalBytes, stream.as_bytes().len())?;
    budget.charge(
        BodyFootnoteLimitKind::RetainedBytes,
        stream.as_bytes().len(),
    )?;
    budget.charge(BodyFootnoteLimitKind::ScratchBytes, stream.as_bytes().len())?;
    budget.charge(BodyFootnoteLimitKind::Allocations, 1)?;
    let archive = Archive::parse_with_limits(stream.as_bytes(), limits).map_err(map_core_error)?;
    Ok((archive, limits))
}

fn compress_archive(
    archive: Archive,
    limits: litchi_iwa_core::Limits,
    budget: &mut TransactionBudget,
) -> Result<Vec<u8>, BodyFootnoteError> {
    let encoded_length = archive
        .encoded_len_with_limits(limits)
        .map_err(map_core_error)?;
    let maximum_compressed =
        SnappyStream::maximum_compressed_len(encoded_length).map_err(map_core_error)?;
    budget.charge_archive_output(encoded_length, maximum_compressed)?;
    let bytes = archive
        .to_bytes_with_limits(limits)
        .map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    if compressed.len() > maximum_compressed {
        return Err(BodyFootnoteError::Verification);
    }
    Ok(compressed)
}

fn unique_body_message_index(object: &ArchiveObject) -> Result<usize, BodyFootnoteError> {
    let mut found = None;
    for (index, message) in object.messages.iter().enumerate() {
        if is_body_storage_message_type(message.type_)
            || litchi_iwa_text_wire::decode_storage_with_limits(
                &message.data,
                RewriteLimits::default(),
            )
            .is_ok()
        {
            if found.replace(index).is_some() {
                return Err(BodyFootnoteError::InvalidSource);
            }
        }
    }
    found.ok_or(BodyFootnoteError::InvalidSource)
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
) -> Result<(Vec<u8>, BodyTemplate), BodyFootnoteError> {
    // The text-wire rewrite is allowed to canonicalize the known storage
    // projection.  Inspect the source table first so deprecated
    // `TSP.Reference` fields cannot disappear during that rewrite and make a
    // hostile graph look canonical to the strict graph codec.
    reject_deprecated_body_references(payload)?;
    let rewrite_limits =
        super::storage_rewrite_limits(limits).map_err(|_| BodyFootnoteError::InvalidSource)?;
    let range = if remove_reference.is_some() {
        position
            ..position
                .checked_add(1)
                .ok_or(BodyFootnoteError::PositionOutOfBounds)?
    } else {
        position..position
    };
    let replacement = if remove_reference.is_some() {
        ""
    } else {
        FOOTNOTE_ANCHOR_TEXT
    };
    let prepared = litchi_iwa_text_wire::prepare_storage_text_rewrite_with_behavior_and_limits(
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
        return Err(BodyFootnoteError::Verification);
    }
    budget.observe(
        BodyFootnoteLimitKind::TextUnits,
        rewritten_result.after_utf16_len(),
    )?;
    let rewritten = rewritten_result.into_bytes();
    budget.charge(BodyFootnoteLimitKind::TextBytes, rewritten.len())?;
    let wire_output = rewritten
        .len()
        .checked_mul(2)
        .ok_or(BodyFootnoteError::InvalidSource)?
        .max(1);
    let wire_limits = WireLimits::default()
        .with_input_bytes(rewritten.len().max(1))
        .and_then(|value| value.with_output_bytes(wire_output))
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let view = WireView::parse_with_limits(&rewritten, wire_limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
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
        .ok_or(BodyFootnoteError::InvalidSource)?
        .min(limits.max_iwa_stream_bytes())
        .max(1);
    let table_fields = rewritten
        .len()
        .checked_mul(16)
        .ok_or(BodyFootnoteError::InvalidSource)?
        .max(1);
    let table_work = rewritten
        .len()
        .checked_mul(32)
        .ok_or(BodyFootnoteError::InvalidSource)?
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
        .ok_or(BodyFootnoteError::InvalidSource)?;
    budget.charge(BodyFootnoteLimitKind::Entries, entry_capacity)?;
    budget.charge(BodyFootnoteLimitKind::Allocations, 1)?;
    entries
        .try_reserve_exact(entry_capacity)
        .map_err(|_| BodyFootnoteError::Allocation {
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
    let replacement_bytes = rewritten_table.len();
    let replacement_capacity = rewritten
        .len()
        .checked_add(replacement_bytes)
        .ok_or(BodyFootnoteError::InvalidSource)?;
    budget.charge(BodyFootnoteLimitKind::WireBytes, replacement_capacity)?;
    budget.charge(BodyFootnoteLimitKind::RetainedBytes, replacement_capacity)?;
    budget.charge(BodyFootnoteLimitKind::Allocations, 1)?;
    let output = replace_length_delimited_field(
        &rewritten,
        FOOTNOTE_TABLE_FIELD,
        if entries.is_empty() {
            None
        } else {
            Some(&rewritten_table)
        },
    )?;
    let template = body_template_from_storage_wire(&rewritten)?;
    Ok((output, template))
}

fn reject_deprecated_body_references(source: &[u8]) -> Result<(), BodyFootnoteError> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let storage = WireView::parse_with_limits(source, limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let Some(table_field) = storage
        .fields()
        .find(|field| field.number() == FOOTNOTE_TABLE_FIELD)
    else {
        return Ok(());
    };
    if table_field.wire_type() != 2 {
        return Err(BodyFootnoteError::InvalidSource);
    }
    let table = WireView::parse_with_limits(table_field.payload(), limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    for entry_field in table.fields().filter(|field| field.number() == 1) {
        if entry_field.wire_type() != 2 {
            return Err(BodyFootnoteError::InvalidSource);
        }
        let entry = WireView::parse_with_limits(entry_field.payload(), limits)
            .map_err(|_| BodyFootnoteError::InvalidSource)?;
        let Some(reference_field) = entry.fields().find(|field| field.number() == 2) else {
            continue;
        };
        if reference_field.wire_type() != 2 {
            return Err(BodyFootnoteError::InvalidSource);
        }
        let reference = WireView::parse_with_limits(reference_field.payload(), limits)
            .map_err(|_| BodyFootnoteError::InvalidSource)?;
        if reference
            .fields()
            .any(|field| matches!(field.number(), 2 | 3))
        {
            return Err(BodyFootnoteError::InvalidSource);
        }
    }
    Ok(())
}

fn body_template_from_storage_wire(source: &[u8]) -> Result<BodyTemplate, BodyFootnoteError> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let view = WireView::parse_with_limits(source, limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
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
                    .map_err(|_| BodyFootnoteError::InvalidSource)?;
                let value = parse_storage_reference(field.payload())?;
                if stylesheet.replace(value).is_some() {
                    return Err(BodyFootnoteError::InvalidSource);
                }
            },
            5 | 7 => {
                field
                    .validate_canonical_key()
                    .and_then(|_| field.validate_canonical_length())
                    .map_err(|_| BodyFootnoteError::InvalidSource)?;
                let value = parse_style_table_reference(field.payload())?;
                let slot = if field.number() == 5 {
                    &mut paragraph_style
                } else {
                    &mut list_style
                };
                if slot.replace(value).is_some() {
                    return Err(BodyFootnoteError::InvalidSource);
                }
            },
            19 => {
                field
                    .validate_canonical_key()
                    .and_then(|_| field.validate_canonical_length())
                    .map_err(|_| BodyFootnoteError::InvalidSource)?;
                let value = parse_language_table(field.payload())?;
                if language.replace(value).is_some() {
                    return Err(BodyFootnoteError::InvalidSource);
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

fn parse_language_table(source: &[u8]) -> Result<Option<String>, BodyFootnoteError> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let table = WireView::parse_with_limits(source, limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let mut language_at_zero = None;
    for field in table.fields() {
        if field.number() != 1 {
            return Err(BodyFootnoteError::InvalidSource);
        }
        field
            .validate_canonical_key()
            .and_then(|_| field.validate_canonical_length())
            .map_err(|_| BodyFootnoteError::InvalidSource)?;
        let (index, value) = parse_language_entry(field.payload(), limits)?;
        if index == 0 {
            if language_at_zero.is_some() {
                return Err(BodyFootnoteError::InvalidSource);
            }
            language_at_zero = value;
        }
    }
    Ok(language_at_zero)
}

fn parse_language_entry(
    source: &[u8],
    limits: WireLimits,
) -> Result<(u32, Option<String>), BodyFootnoteError> {
    let entry = WireView::parse_with_limits(source, limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let mut index = None;
    let mut value = None;
    for field in entry.fields() {
        match field.number() {
            1 => {
                field
                    .validate_canonical_key()
                    .map_err(|_| BodyFootnoteError::InvalidSource)?;
                if field.wire_type() != 0 || index.is_some() {
                    return Err(BodyFootnoteError::InvalidSource);
                }
                let (raw, width) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                    .map_err(|_| BodyFootnoteError::InvalidSource)?;
                if width != field.payload().len() {
                    return Err(BodyFootnoteError::InvalidSource);
                }
                index = Some(u32::try_from(raw).map_err(|_| BodyFootnoteError::InvalidSource)?);
            },
            2 => {
                field
                    .validate_canonical_key()
                    .and_then(|_| field.validate_canonical_length())
                    .map_err(|_| BodyFootnoteError::InvalidSource)?;
                if value.is_some() {
                    return Err(BodyFootnoteError::InvalidSource);
                }
                value = Some(
                    std::str::from_utf8(field.payload())
                        .map_err(|_| BodyFootnoteError::InvalidSource)?
                        .to_owned(),
                );
            },
            _ => {},
        }
    }
    Ok((index.ok_or(BodyFootnoteError::InvalidSource)?, value))
}

fn parse_storage_reference(source: &[u8]) -> Result<NonZeroU64, BodyFootnoteError> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let view = WireView::parse_with_limits(source, limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let mut identifier = None;
    for field in view.fields() {
        match field.number() {
            1 => {
                field
                    .validate_canonical_key()
                    .map_err(|_| BodyFootnoteError::InvalidSource)?;
                if field.wire_type() != 0 || identifier.is_some() {
                    return Err(BodyFootnoteError::InvalidSource);
                }
                let (value, width) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                    .map_err(|_| BodyFootnoteError::InvalidSource)?;
                if width != field.payload().len() {
                    return Err(BodyFootnoteError::InvalidSource);
                }
                identifier = Some(NonZeroU64::new(value).ok_or(BodyFootnoteError::InvalidSource)?);
            },
            2 | 3 => return Err(BodyFootnoteError::InvalidSource),
            _ => {},
        }
    }
    identifier.ok_or(BodyFootnoteError::InvalidSource)
}

fn parse_style_table_reference(source: &[u8]) -> Result<Option<NonZeroU64>, BodyFootnoteError> {
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let table = WireView::parse_with_limits(source, limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
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
            .map_err(|_| BodyFootnoteError::InvalidSource)?;
        let (index, reference) = parse_style_table_entry(field.payload(), limits)?;
        if index == 0 {
            if reference_at_zero.is_some() {
                return Err(BodyFootnoteError::InvalidSource);
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
) -> Result<(u32, Option<NonZeroU64>), BodyFootnoteError> {
    let entry = WireView::parse_with_limits(source, limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let mut index = None;
    let mut reference = None;
    for field in entry.fields() {
        match field.number() {
            1 => {
                field
                    .validate_canonical_key()
                    .map_err(|_| BodyFootnoteError::InvalidSource)?;
                if field.wire_type() != 0 || index.is_some() {
                    return Err(BodyFootnoteError::InvalidSource);
                }
                let (value, width) = litchi_iwa_common::decode_varint_from_bytes(field.payload())
                    .map_err(|_| BodyFootnoteError::InvalidSource)?;
                if width != field.payload().len() {
                    return Err(BodyFootnoteError::InvalidSource);
                }
                index = Some(u32::try_from(value).map_err(|_| BodyFootnoteError::InvalidSource)?);
            },
            2 => {
                field
                    .validate_canonical_key()
                    .and_then(|_| field.validate_canonical_length())
                    .map_err(|_| BodyFootnoteError::InvalidSource)?;
                if reference.is_some() {
                    return Err(BodyFootnoteError::InvalidSource);
                }
                reference = Some(parse_storage_reference(field.payload())?);
            },
            _ => {},
        }
    }
    Ok((index.ok_or(BodyFootnoteError::InvalidSource)?, reference))
}

fn replace_length_delimited_field(
    source: &[u8],
    number: u32,
    replacement: Option<&[u8]>,
) -> Result<Vec<u8>, BodyFootnoteError> {
    let replacement_bytes = replacement.map_or(0, |value| value.len());
    let output_capacity = source
        .len()
        .checked_add(replacement_bytes)
        .and_then(|value| value.checked_add(32))
        .ok_or(BodyFootnoteError::InvalidSource)?;
    let reserve_capacity = source
        .len()
        .checked_add(replacement_bytes)
        .ok_or(BodyFootnoteError::InvalidSource)?;
    let limits = WireLimits::default()
        .with_input_bytes(source.len().max(1))
        .and_then(|value| value.with_output_bytes(output_capacity))
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let view = WireView::parse_with_limits(source, limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(reserve_capacity)
        .map_err(|_| BodyFootnoteError::Allocation {
            amount: source.len(),
        })?;
    let mut found = false;
    for field in view.fields() {
        if field.number() == number {
            if found {
                return Err(BodyFootnoteError::InvalidSource);
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
    Ok(output)
}

fn append_length_delimited(
    output: &mut Vec<u8>,
    number: u32,
    payload: &[u8],
) -> Result<(), BodyFootnoteError> {
    put_varint(output, (u64::from(number) << 3) | 2);
    put_varint(
        output,
        u64::try_from(payload.len()).map_err(|_| BodyFootnoteError::InvalidSource)?,
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
) -> Result<[ArchiveObject; 3], BodyFootnoteError> {
    budget.charge(BodyFootnoteLimitKind::Entries, 3)?;
    budget.charge(BodyFootnoteLimitKind::Allocations, 3)?;
    let mut reference = ArchiveObject::new_with_limits(
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
    reference.archive_info.message_infos[0].versions = vec![1, 0, 5];
    reference.archive_info.message_infos[0].object_references =
        vec![write.storage_identifier().get()];
    let mut storage = ArchiveObject::new_with_limits(
        write.storage_identifier().get(),
        vec![RawMessage {
            type_: 2_001,
            data: payloads.storage().to_vec(),
        }],
        limits,
    )
    .map_err(map_core_error)?;
    storage.archive_info.message_infos[0].versions = vec![1, 0, 5];
    storage.archive_info.message_infos[0].object_references = storage_references;
    let mut marker = ArchiveObject::new_with_limits(
        write.marker_identifier().get(),
        vec![RawMessage {
            type_: TEXTUAL_ATTACHMENT_MESSAGE_TYPE,
            data: payloads.marker().to_vec(),
        }],
        limits,
    )
    .map_err(map_core_error)?;
    marker.archive_info.message_infos[0].versions = vec![1, 0, 5];
    Ok([reference, storage, marker])
}

fn max_object_identifier(
    package: &Package,
    budget: &mut TransactionBudget,
) -> Result<u64, BodyFootnoteError> {
    let physical_count =
        package
            .state
            .source
            .components()
            .iter()
            .try_fold(0usize, |count, component| {
                count
                    .checked_add(component.archive().objects.len())
                    .ok_or(BodyFootnoteError::InvalidSource)
            })?;
    budget.charge(BodyFootnoteLimitKind::Entries, physical_count)?;
    let physical = package
        .state
        .source
        .components()
        .iter()
        .flat_map(|component| component.archive().objects.iter())
        .filter_map(|object| object.archive_info.identifier)
        .max()
        .ok_or(BodyFootnoteError::InvalidSource)?;
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
) -> Result<(), BodyFootnoteError> {
    const BODY_REFERENCE_PATH: &[u32] = &[16, 1, 2];
    let aggregate_count = info
        .object_references
        .iter()
        .filter(|identifier| **identifier == reference)
        .count();
    if adding {
        if aggregate_count != 0 {
            return Err(BodyFootnoteError::InvalidSource);
        }
    } else if aggregate_count != 1 {
        return Err(BodyFootnoteError::UnsupportedDependency);
    }
    let mut canonical_fields = 0usize;
    for field in &mut info.field_infos {
        if field.path.path == BODY_REFERENCE_PATH {
            canonical_fields = canonical_fields
                .checked_add(1)
                .ok_or(BodyFootnoteError::InvalidSource)?;
            if canonical_fields > 1 {
                return Err(BodyFootnoteError::InvalidSource);
            }
            let count = field
                .object_references
                .iter()
                .filter(|identifier| **identifier == reference)
                .count();
            if count > 1 {
                return Err(BodyFootnoteError::InvalidSource);
            }
            if adding {
                if count != 0 {
                    return Err(BodyFootnoteError::InvalidSource);
                }
                field.object_references.push(reference);
            } else {
                field
                    .object_references
                    .retain(|identifier| *identifier != reference);
            }
        } else if field.object_references.contains(&reference) {
            return Err(BodyFootnoteError::UnsupportedDependency);
        }
    }
    if adding {
        info.object_references.push(reference);
        if canonical_fields == 0 {
            let mut field = FieldInfo::new(BODY_REFERENCE_PATH.to_vec());
            field.object_references.push(reference);
            info.field_infos.push(field);
        }
    } else {
        info.object_references
            .retain(|identifier| *identifier != reference);
    }
    Ok(())
}

fn prove_exclusive_graph(
    package: &Package,
    reference: u64,
    storage: u64,
    marker: u64,
    budget: &mut TransactionBudget,
) -> Result<(), BodyFootnoteError> {
    let mut incoming = [0usize; 3];
    let mut typed_incoming = [0usize; 3];
    let graphs = super::footnote_text::native_footnotes(package).map_err(map_text_error)?;
    budget.charge(BodyFootnoteLimitKind::Entries, graphs.len())?;
    budget.charge(
        BodyFootnoteLimitKind::WireWork,
        package.state.source.source_bytes().len(),
    )?;
    let selected = graphs.iter().find(|graph| {
        graph.reference_identifier.get() == reference
            || graph.storage_identifier.get() == storage
            || graph.marker_identifier.get() == marker
    });
    if selected.is_none_or(|graph| {
        graph.reference_identifier.get() != reference
            || graph.storage_identifier.get() != storage
            || graph.marker_identifier.get() != marker
    }) {
        return Err(BodyFootnoteError::UnsupportedDependency);
    }
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for info in &object.archive_info.message_infos {
                for target in &info.object_references {
                    match *target {
                        value if value == reference => {
                            incoming[0] = incoming[0]
                                .checked_add(1)
                                .ok_or(BodyFootnoteError::InvalidSource)?;
                        },
                        value if value == storage => {
                            incoming[1] = incoming[1]
                                .checked_add(1)
                                .ok_or(BodyFootnoteError::InvalidSource)?;
                        },
                        value if value == marker => {
                            incoming[2] = incoming[2]
                                .checked_add(1)
                                .ok_or(BodyFootnoteError::InvalidSource)?;
                        },
                        _ => {},
                    }
                }
                for target in &info.data_references {
                    if [reference, storage, marker].contains(target) {
                        return Err(BodyFootnoteError::UnsupportedDependency);
                    }
                }
                for field in &info.field_infos {
                    for target in &field.object_references {
                        let Some(slot) = graph_edge_slot(*target, reference, storage, marker)
                        else {
                            continue;
                        };
                        let expected_path = match slot {
                            0 => &[16, 1, 2][..],
                            1 => &[2][..],
                            _ => &[9, 1, 2][..],
                        };
                        if field.path.path != expected_path {
                            return Err(BodyFootnoteError::UnsupportedDependency);
                        }
                        typed_incoming[slot] = typed_incoming[slot]
                            .checked_add(1)
                            .ok_or(BodyFootnoteError::InvalidSource)?;
                    }
                    for target in &field.data_references {
                        if [reference, storage, marker].contains(target) {
                            return Err(BodyFootnoteError::UnsupportedDependency);
                        }
                    }
                }
            }
        }
    }
    // The source graph has exactly body -> reference -> storage -> marker.
    if incoming != [1, 1, 1] || typed_incoming.iter().any(|count| *count > 1) {
        return Err(BodyFootnoteError::UnsupportedDependency);
    }
    Ok(())
}

fn validate_existing_footnote_graphs(
    source: &Package,
    budget: &mut TransactionBudget,
) -> Result<(), BodyFootnoteError> {
    let graphs = super::footnote_text::native_footnotes(source).map_err(map_text_error)?;
    budget.charge(BodyFootnoteLimitKind::Entries, graphs.len())?;
    budget.charge(
        BodyFootnoteLimitKind::WireWork,
        source.state.source.source_bytes().len(),
    )?;
    let location = metadata_location(source)?;
    let (facts, _) = metadata_facts(source, &location, budget)?;
    for graph in graphs {
        prove_exclusive_graph(
            source,
            graph.reference_identifier.get(),
            graph.storage_identifier.get(),
            graph.marker_identifier.get(),
            budget,
        )?;
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
            return Err(BodyFootnoteError::UnsupportedDependency);
        }
        let storage_bindings = facts
            .uuids
            .iter()
            .filter(|binding| binding.object == graph.storage_identifier.get())
            .collect::<Vec<_>>();
        if storage_bindings.len() != 1 || storage_bindings.iter().any(|binding| !binding.current) {
            return Err(BodyFootnoteError::UnsupportedDependency);
        }
    }
    Ok(())
}

fn validate_storage_attachment_owners(
    package: &Package,
    storage_identifier: u64,
    marker_identifier: u64,
) -> Result<(), BodyFootnoteError> {
    let object = package
        .state
        .source
        .components()
        .iter()
        .find_map(|component| component.archive().object(storage_identifier))
        .ok_or(BodyFootnoteError::InvalidSource)?;
    let messages = object
        .messages
        .iter()
        .filter(|message| is_body_storage_message_type(message.type_));
    let mut message = messages;
    let storage_message = message.next().ok_or(BodyFootnoteError::InvalidSource)?;
    if message.next().is_some() {
        return Err(BodyFootnoteError::InvalidSource);
    }
    let limits = WireLimits::default()
        .with_input_bytes(storage_message.data.len().max(1))
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let storage = WireView::parse_with_limits(&storage_message.data, limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let attachment_fields = storage
        .fields()
        .filter(|field| field.number() == 9)
        .collect::<Vec<_>>();
    if attachment_fields.len() != 1 || attachment_fields[0].wire_type() != 2 {
        return Err(BodyFootnoteError::InvalidSource);
    }
    let table = WireView::parse_with_limits(attachment_fields[0].payload(), limits)
        .map_err(|_| BodyFootnoteError::InvalidSource)?;
    let mut marker_count = 0usize;
    for entry_field in table.fields().filter(|field| field.number() == 1) {
        if entry_field.wire_type() != 2 {
            return Err(BodyFootnoteError::InvalidSource);
        }
        let entry = WireView::parse_with_limits(entry_field.payload(), limits)
            .map_err(|_| BodyFootnoteError::InvalidSource)?;
        let references = entry
            .fields()
            .filter(|field| field.number() == 2)
            .collect::<Vec<_>>();
        if references.len() != 1 || references[0].wire_type() != 2 {
            return Err(BodyFootnoteError::InvalidSource);
        }
        let reference = WireView::parse_with_limits(references[0].payload(), limits)
            .map_err(|_| BodyFootnoteError::InvalidSource)?;
        let identifiers = reference
            .fields()
            .filter(|field| field.number() == 1)
            .collect::<Vec<_>>();
        if identifiers.len() != 1 || identifiers[0].wire_type() != 0 {
            return Err(BodyFootnoteError::InvalidSource);
        }
        let (identifier, width) =
            litchi_iwa_common::decode_varint_from_bytes(identifiers[0].payload())
                .map_err(|_| BodyFootnoteError::InvalidSource)?;
        if width != identifiers[0].payload().len() || identifier != marker_identifier {
            return Err(BodyFootnoteError::UnsupportedDependency);
        }
        marker_count = marker_count
            .checked_add(1)
            .ok_or(BodyFootnoteError::InvalidSource)?;
    }
    if marker_count != 1 {
        return Err(BodyFootnoteError::UnsupportedDependency);
    }
    Ok(())
}

fn graph_edge_slot(target: u64, reference: u64, storage: u64, marker: u64) -> Option<usize> {
    if target == reference {
        Some(0)
    } else if target == storage {
        Some(1)
    } else if target == marker {
        Some(2)
    } else {
        None
    }
}
