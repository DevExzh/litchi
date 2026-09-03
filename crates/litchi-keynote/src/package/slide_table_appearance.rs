//! Selector-first, source-preserving Keynote slide-table appearance transactions.
//!
//! This module owns the public transaction boundary. Native style objects,
//! package metadata, wire framing, and physical member names remain private.

use std::collections::{HashMap, HashSet};
use std::fmt;
use std::mem::size_of;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{Catalog, Entry, EntryEdit};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, varint::encoded_len, wire::WireView,
};
use litchi_iwa_core::archive::{FieldObjectReferenceTransition, ObjectReferenceTransition};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceScope, ArchiveReferenceVisitor, FieldType, RawMessage,
    SnappyStream,
};
use litchi_iwa_protos::package_metadata_codec::{
    self as metadata_codec, CombinedBatch, CombinedSaveTokenBatch, ComponentSelector,
    ExternalReferenceAddition, ExternalReferenceRemoval, ObjectUuidAddition,
    PackageMetadataVisitor, RewriteError as MetadataRewriteError,
    RewriteOptions as MetadataRewriteOptions, SaveTokenBatch, UuidBits,
    inspect_package_metadata_with_visitor,
    prepare_package_metadata_combined_additions_and_removals_and_save_tokens,
};
use litchi_iwa_protos::{table_appearance_codec as appearance_codec, table_info_codec};
use thiserror::Error as ThisError;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::selector::SlideSelector;
use crate::slide::table::{
    TableSelector,
    appearance::{Appearance, Banding, GridlineVisibility, Gridlines, RowSizing},
};

const SLIDE_MESSAGE_TYPE: u32 = 5;
const SHOW_MESSAGE_TYPE: u32 = 2;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const TABLE_STYLESHEET_MESSAGE_TYPE: u32 = 401;
const METADATA_MESSAGE_TYPE: u32 = 11_006;
const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_Z_ORDER_FIELD: u32 = 42;
const TABLE_SUPER_FIELD: u32 = 1;
const TABLE_MODEL_FIELD: u32 = 2;
const DRAWABLE_PARENT_FIELD: u32 = 2;
const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

/// A content-free location associated with a slide-table appearance operation.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableAppearancePath {
    /// The complete Keynote package.
    Package,
    /// One checked zero-based slide/table location.
    Table { slide: Position, table: Position },
}

impl fmt::Display for SlideTableAppearancePath {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Package => formatter.write_str("package"),
            Self::Table { slide, table } => {
                write!(formatter, "slide {} table {}", slide.get(), table.get())
            },
        }
    }
}

/// A finite resource governed by an appearance transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SlideTableAppearanceLimitKind {
    InputBytes,
    OutputBytes,
    Entries,
    EntryBytes,
    TotalBytes,
    PayloadObjects,
    PayloadMessages,
    PayloadItems,
    References,
    WireBytes,
    WireOutputBytes,
    WireFields,
    WireNesting,
    WireWork,
    Styles,
    Components,
    Allocations,
    Retained,
    Scratch,
    TransactionWork,
}

impl fmt::Display for SlideTableAppearanceLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{self:?}")
    }
}

/// Failure from a Keynote slide-table appearance read or transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, ThisError)]
#[non_exhaustive]
pub enum SlideTableAppearanceError {
    #[error("this Keynote source does not support exact slide-table appearance editing")]
    UnsupportedSource,
    #[error("the selected Keynote slide-table appearance has an unsupported dependency")]
    UnsupportedDependency,
    #[error("the selected Keynote slide-table appearance topology is unsupported")]
    UnsupportedTopology,
    #[error("the Keynote slide-table appearance selector is ambiguous")]
    AmbiguousSelector,
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    #[error("the selected Keynote slide has no table at position {position:?}")]
    TablePositionNotFound { position: Position },
    #[error("the selected Keynote slide table is locked")]
    Locked,
    #[error("the selected Keynote slide-table appearance source is invalid")]
    InvalidSource,
    #[error(
        "Keynote slide-table appearance {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        kind: SlideTableAppearanceLimitKind,
        observed: u64,
        maximum: u64,
    },
    #[error("could not allocate {amount} units for the Keynote appearance transaction")]
    Allocation { amount: usize },
    #[error("the edited Keynote slide-table appearance failed semantic verification")]
    Verification,
    #[error("the slide-table appearance patch does not match the exact source package")]
    PatchConflict,
}

/// One operation-local accounting ledger for appearance reads and writes.
///
/// The package parser has its own physical and semantic limits, but those
/// limits are per parse/decode.  An appearance transaction performs several
/// parses (native members, metadata, reassembly and the candidate reopen), so
/// it needs a single residual ledger as well.  The counters intentionally
/// describe logical work and retained buffers; they are not allocator/RSS
/// telemetry.
#[derive(Debug, Clone, Copy)]
struct AppearanceBudget {
    max_input: usize,
    max_output: usize,
    max_entries: usize,
    max_entry_bytes: usize,
    max_total_bytes: usize,
    max_objects: usize,
    max_messages: usize,
    max_items: usize,
    max_references: usize,
    max_wire_bytes: usize,
    max_wire_output: usize,
    max_fields: usize,
    max_nesting: usize,
    max_work: usize,
    max_styles: usize,
    max_components: usize,
    max_allocations: usize,
    max_retained: usize,
    max_scratch: usize,
    max_transaction_work: usize,
    input: usize,
    output: usize,
    entries: usize,
    entry_bytes: usize,
    total_bytes: usize,
    objects: usize,
    messages: usize,
    items: usize,
    references: usize,
    wire_bytes: usize,
    wire_output: usize,
    fields: usize,
    nesting: usize,
    work: usize,
    styles: usize,
    components: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
    transaction_work: usize,
}

impl AppearanceBudget {
    fn new(package: &Package) -> Result<Self, SlideTableAppearanceError> {
        let wire = package.wire_limits().map_err(map_wire_error)?;
        let archive = package.state.options.archive();
        let source_limit = usize::try_from(archive.max_input_bytes())
            .map_err(|_| SlideTableAppearanceError::InvalidSource)?
            .max(1);
        let total_limit = usize::try_from(archive.max_total_bytes())
            .unwrap_or(usize::MAX)
            .max(source_limit);
        let aggregate = source_limit
            .checked_mul(8)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        let components = package.state.source.components().len().max(1);
        let component_budget = components
            .checked_mul(8)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        Ok(Self {
            max_input: aggregate,
            max_output: total_limit.min(aggregate.max(1)),
            max_entries: aggregate,
            max_entry_bytes: total_limit,
            max_total_bytes: total_limit,
            max_objects: package
                .semantic_limits()
                .max_objects()
                .saturating_mul(8)
                .max(1),
            max_messages: package
                .semantic_limits()
                .max_objects()
                .saturating_mul(8)
                .max(1),
            max_items: package
                .semantic_limits()
                .max_slides()
                .saturating_mul(8)
                .max(1),
            max_references: package
                .semantic_limits()
                .max_references()
                .saturating_mul(8)
                .max(1),
            max_wire_bytes: aggregate,
            max_wire_output: aggregate,
            max_fields: wire
                .max_fields()
                .saturating_mul(8)
                .min(WireLimits::MAX_FIELDS),
            max_nesting: wire.max_nesting(),
            max_work: wire.max_rewrite_work(),
            max_styles: package
                .semantic_limits()
                .max_references()
                .saturating_mul(8)
                .max(1),
            max_components: component_budget.max(package.semantic_limits().max_objects()),
            max_allocations: aggregate,
            max_retained: total_limit.min(aggregate.max(1)),
            max_scratch: aggregate,
            max_transaction_work: aggregate
                .checked_mul(4)
                .ok_or(SlideTableAppearanceError::InvalidSource)?,
            input: 0,
            output: 0,
            entries: 0,
            entry_bytes: 0,
            total_bytes: 0,
            objects: 0,
            messages: 0,
            items: 0,
            references: 0,
            wire_bytes: 0,
            wire_output: 0,
            fields: 0,
            nesting: 0,
            work: 0,
            styles: 0,
            components: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
            transaction_work: 0,
        })
    }

    fn add(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: SlideTableAppearanceLimitKind,
    ) -> Result<(), SlideTableAppearanceError> {
        let observed = current
            .checked_add(amount)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        if observed > maximum {
            return Err(SlideTableAppearanceError::LimitExceeded {
                kind,
                observed: observed as u64,
                maximum: maximum as u64,
            });
        }
        *current = observed;
        Ok(())
    }

    fn input(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.input,
            amount,
            self.max_input,
            SlideTableAppearanceLimitKind::InputBytes,
        )
    }

    fn output(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.output,
            amount,
            self.max_output,
            SlideTableAppearanceLimitKind::OutputBytes,
        )
    }

    fn entries(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.entries,
            amount,
            self.max_entries,
            SlideTableAppearanceLimitKind::Entries,
        )
    }

    fn entry_bytes(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.entry_bytes,
            amount,
            self.max_entry_bytes,
            SlideTableAppearanceLimitKind::EntryBytes,
        )
    }

    fn total_bytes(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.total_bytes,
            amount,
            self.max_total_bytes,
            SlideTableAppearanceLimitKind::TotalBytes,
        )
    }

    fn objects(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.objects,
            amount,
            self.max_objects,
            SlideTableAppearanceLimitKind::PayloadObjects,
        )
    }

    fn messages(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.messages,
            amount,
            self.max_messages,
            SlideTableAppearanceLimitKind::PayloadMessages,
        )
    }

    fn items(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.items,
            amount,
            self.max_items,
            SlideTableAppearanceLimitKind::PayloadItems,
        )
    }

    fn references(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.references,
            amount,
            self.max_references,
            SlideTableAppearanceLimitKind::References,
        )
    }

    fn wire_bytes(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.wire_bytes,
            amount,
            self.max_wire_bytes,
            SlideTableAppearanceLimitKind::WireBytes,
        )
    }

    fn wire_output(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.wire_output,
            amount,
            self.max_wire_output,
            SlideTableAppearanceLimitKind::WireOutputBytes,
        )
    }

    fn fields(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.fields,
            amount,
            self.max_fields,
            SlideTableAppearanceLimitKind::WireFields,
        )
    }

    fn work(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.work,
            amount,
            self.max_work,
            SlideTableAppearanceLimitKind::WireWork,
        )
    }

    fn styles(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.styles,
            amount,
            self.max_styles,
            SlideTableAppearanceLimitKind::Styles,
        )
    }

    fn components(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.components,
            amount,
            self.max_components,
            SlideTableAppearanceLimitKind::Components,
        )
    }

    fn allocations(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.allocations,
            amount,
            self.max_allocations,
            SlideTableAppearanceLimitKind::Allocations,
        )
    }

    fn retained(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.retained,
            amount,
            self.max_retained,
            SlideTableAppearanceLimitKind::Retained,
        )
    }

    fn scratch(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.scratch,
            amount,
            self.max_scratch,
            SlideTableAppearanceLimitKind::Scratch,
        )
    }

    fn transaction_work(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        Self::add(
            &mut self.transaction_work,
            amount,
            self.max_transaction_work,
            SlideTableAppearanceLimitKind::TransactionWork,
        )
    }

    fn source_catalog(
        &mut self,
        package: &Package,
        catalog: &litchi_iwa_archive::SourceCatalog,
    ) -> Result<(), SlideTableAppearanceError> {
        let source_bytes = catalog.shared_source().len();
        self.input(source_bytes)?;
        self.wire_bytes(source_bytes)?;
        self.transaction_work(source_bytes)?;
        self.entries(catalog.package().len())?;
        self.components(package.state.source.components().len())?;
        let mut objects = 0usize;
        let mut messages = 0usize;
        let mut references = 0usize;
        let mut fields = 0usize;
        let mut items = 0usize;
        for component in package.state.source.components().iter() {
            objects = objects
                .checked_add(component.archive().objects.len())
                .ok_or(SlideTableAppearanceError::InvalidSource)?;
            for object in &component.archive().objects {
                messages = messages
                    .checked_add(object.messages.len())
                    .ok_or(SlideTableAppearanceError::InvalidSource)?;
                items = items
                    .checked_add(
                        object
                            .messages
                            .iter()
                            .filter(|message| message.type_ == SLIDE_MESSAGE_TYPE)
                            .count(),
                    )
                    .ok_or(SlideTableAppearanceError::InvalidSource)?;
                for info in &object.archive_info.message_infos {
                    references = references
                        .checked_add(info.object_references.len())
                        .ok_or(SlideTableAppearanceError::InvalidSource)?;
                    fields = fields
                        .checked_add(info.field_infos.len())
                        .ok_or(SlideTableAppearanceError::InvalidSource)?;
                }
            }
        }
        self.objects(objects)?;
        self.messages(messages)?;
        self.references(references)?;
        self.fields(fields)?;
        self.items(items)?;
        self.transaction_work(
            source_bytes
                .checked_add(objects)
                .and_then(|value| value.checked_add(messages))
                .and_then(|value| value.checked_add(references))
                .and_then(|value| value.checked_add(fields))
                .ok_or(SlideTableAppearanceError::InvalidSource)?,
        )
    }

    fn physical(&mut self, amount: usize) -> Result<(), SlideTableAppearanceError> {
        self.input(amount)?;
        self.wire_bytes(amount)?;
        self.work(amount)?;
        self.transaction_work(amount)
    }

    fn residual_wire(&self, package: &Package) -> Result<WireLimits, SlideTableAppearanceError> {
        let base = package.wire_limits().map_err(map_wire_error)?;
        base.with_input_bytes(
            base.max_input_bytes()
                .min(self.max_wire_bytes.saturating_sub(self.wire_bytes).max(1)),
        )
        .and_then(|limits| {
            limits.with_fields(
                base.max_fields()
                    .min(self.max_fields.saturating_sub(self.fields).max(1)),
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

    fn codec_options(
        &self,
        package: &Package,
        source: &[u8],
    ) -> Result<appearance_codec::DecodeOptions, SlideTableAppearanceError> {
        let base = appearance_codec::DecodeOptions::for_source(source);
        let limits = self.residual_wire(package)?;
        Ok(base
            .with_max_input_bytes(
                base.max_input_bytes()
                    .min(limits.max_input_bytes())
                    .min(self.max_wire_bytes.saturating_sub(self.wire_bytes).max(1)),
            )
            .with_max_output_bytes(
                base.max_output_bytes()
                    .min(limits.max_output_bytes())
                    .min(self.max_wire_output.saturating_sub(self.wire_output).max(1)),
            )
            .with_max_fields(
                base.max_fields()
                    .min(limits.max_fields())
                    .min(self.max_fields.saturating_sub(self.fields).max(1)),
            )
            .with_max_work_bytes(
                base.max_work_bytes()
                    .min(limits.max_rewrite_work())
                    .min(self.max_work.saturating_sub(self.work).max(1)),
            )
            .with_recursion_limit(
                base.recursion_limit()
                    .min(u32::try_from(self.max_nesting).unwrap_or(u32::MAX)),
            )
            .with_max_styles(self.max_styles.saturating_sub(self.styles).max(1))
            .with_max_allocations(self.max_allocations.saturating_sub(self.allocations).max(1)))
    }

    fn codec_report(
        &mut self,
        report: appearance_codec::DecodeReport,
    ) -> Result<(), SlideTableAppearanceError> {
        self.input(report.input_bytes())?;
        self.wire_bytes(report.input_bytes())?;
        self.wire_output(report.output_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.allocations(report.allocations())?;
        self.nesting = self.nesting.max(report.max_depth() as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideTableAppearanceError::LimitExceeded {
                kind: SlideTableAppearanceLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.transaction_work(
            report
                .input_bytes()
                .checked_add(report.output_bytes())
                .ok_or(SlideTableAppearanceError::InvalidSource)?,
        )
    }

    fn codec_requirements(
        &mut self,
        report: appearance_codec::DecodeReport,
        requirements: appearance_codec::RewriteExecutionRequirements,
    ) -> Result<(), SlideTableAppearanceError> {
        self.output(requirements.output_bytes())?;
        self.wire_output(requirements.output_bytes())?;
        self.fields(requirements.fields().saturating_sub(report.fields()))?;
        self.work(
            requirements
                .work_bytes()
                .saturating_sub(report.work_bytes()),
        )?;
        self.allocations(
            requirements
                .allocations()
                .saturating_sub(report.allocations()),
        )?;
        self.nesting = self.nesting.max(requirements.max_depth() as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideTableAppearanceError::LimitExceeded {
                kind: SlideTableAppearanceLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.transaction_work(
            requirements
                .input_bytes()
                .saturating_add(requirements.output_bytes())
                .saturating_sub(report.input_bytes().saturating_add(report.output_bytes())),
        )
    }

    fn metadata_options(
        &self,
        source: &Package,
        bytes: usize,
        additions: usize,
    ) -> MetadataRewriteOptions {
        let base = metadata_options(source, bytes, additions);
        MetadataRewriteOptions::new(
            base.max_input_bytes()
                .min(self.max_wire_bytes.saturating_sub(self.wire_bytes).max(1)),
            base.max_output_bytes()
                .min(self.max_wire_output.saturating_sub(self.wire_output).max(1)),
            base.max_fields()
                .min(self.max_fields.saturating_sub(self.fields).max(1)),
            base.max_work_bytes()
                .min(self.max_work.saturating_sub(self.work).max(1)),
            base.recursion_limit().min(self.max_nesting as u32),
            base.max_components()
                .min(self.max_components.saturating_sub(self.components).max(1)),
            base.max_references()
                .min(self.max_references.saturating_sub(self.references).max(1)),
            base.max_additions()
                .min(self.max_items.saturating_sub(self.items).max(1)),
        )
    }

    fn metadata_report(
        &mut self,
        report: metadata_codec::RewriteReport,
    ) -> Result<(), SlideTableAppearanceError> {
        self.input(report.input_bytes())?;
        self.wire_bytes(report.input_bytes())?;
        self.wire_output(report.output_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.components(report.components_scanned())?;
        self.references(report.references_scanned())?;
        self.items(report.additions().saturating_add(report.removals()))?;
        self.allocations(report.allocations())?;
        self.retained(report.retained_bytes())?;
        self.scratch(report.scratch_bytes())?;
        self.nesting = self.nesting.max(report.max_depth() as usize);
        if self.nesting > self.max_nesting {
            return Err(SlideTableAppearanceError::LimitExceeded {
                kind: SlideTableAppearanceLimitKind::WireNesting,
                observed: self.nesting as u64,
                maximum: self.max_nesting as u64,
            });
        }
        self.transaction_work(
            report
                .input_bytes()
                .checked_add(report.output_bytes())
                .and_then(|value| value.checked_add(report.work_bytes()))
                .ok_or(SlideTableAppearanceError::InvalidSource)?,
        )
    }

    fn metadata_requirements(
        &mut self,
        report: metadata_codec::RewriteReport,
        requirements: metadata_codec::RewriteExecutionRequirements,
    ) -> Result<(), SlideTableAppearanceError> {
        self.output(requirements.output_bytes())?;
        self.wire_output(requirements.output_bytes())?;
        self.fields(requirements.fields().saturating_sub(report.fields()))?;
        self.work(
            requirements
                .work_bytes()
                .saturating_sub(report.work_bytes()),
        )?;
        self.components(
            requirements
                .components()
                .saturating_sub(report.components_scanned()),
        )?;
        self.references(
            requirements
                .references()
                .saturating_sub(report.references_scanned()),
        )?;
        self.allocations(
            requirements
                .allocations()
                .saturating_sub(report.allocations()),
        )?;
        self.retained(
            requirements
                .retained_bytes()
                .saturating_sub(report.retained_bytes()),
        )?;
        self.scratch(
            requirements
                .scratch_bytes()
                .saturating_sub(report.scratch_bytes()),
        )?;
        self.transaction_work(
            requirements
                .output_bytes()
                .saturating_add(requirements.retained_bytes())
                .saturating_add(requirements.scratch_bytes()),
        )
    }

    fn reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<(), SlideTableAppearanceError> {
        self.output(requirements.output_bytes())?;
        self.total_bytes(requirements.output_bytes())?;
        self.allocations(requirements.allocations())?;
        self.retained(requirements.retained_bytes())?;
        self.scratch(requirements.scratch_bytes())?;
        self.transaction_work(
            requirements
                .output_bytes()
                .checked_add(requirements.scratch_bytes())
                .and_then(|value| value.checked_add(requirements.retained_bytes()))
                .ok_or(SlideTableAppearanceError::InvalidSource)?,
        )
    }

    fn candidate_reopen(
        &mut self,
        candidate_bytes: usize,
    ) -> Result<(), SlideTableAppearanceError> {
        self.input(candidate_bytes)?;
        self.wire_bytes(candidate_bytes)?;
        self.retained(candidate_bytes)?;
        self.scratch(candidate_bytes)?;
        self.allocations(1)?;
        self.transaction_work(candidate_bytes)
    }
}

/// Immutable appearance settings staged against one exact package snapshot.
pub struct SlideTableAppearanceEdit<'a> {
    source: &'a Package,
    path: SlideTableAppearancePath,
    before: Appearance,
    after: Appearance,
    #[allow(dead_code)]
    selection: AppearanceSelection,
}

impl fmt::Debug for SlideTableAppearanceEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableAppearanceEdit")
            .field("path", &self.path)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableAppearanceEdit<'_> {
    #[must_use]
    pub const fn path(&self) -> SlideTableAppearancePath {
        self.path
    }

    #[must_use]
    pub const fn before(&self) -> Appearance {
        self.before
    }

    #[must_use]
    pub const fn after(&self) -> Appearance {
        self.after
    }

    #[must_use]
    pub const fn appearance(&self) -> Appearance {
        self.after
    }

    #[must_use]
    pub const fn set(mut self, appearance: Appearance) -> Self {
        self.after = appearance;
        self
    }

    pub fn commit(self) -> Result<SlideTableAppearanceCommit, SlideTableAppearanceError> {
        commit_edit(self)
    }
}

#[derive(Clone, PartialEq, Eq)]
struct AppearanceSelection {
    slide_position: Position,
    table_position: Position,
    slide_identifier: u64,
    table_info_identifier: u64,
    model_identifier: u64,
    slide_message_index: usize,
    table_info_message_index: usize,
    model_message_index: usize,
    model_component: Arc<str>,
    table_info_component: Arc<str>,
    style_component: Option<Arc<str>>,
    metadata_component: Option<Arc<str>>,
    model_style_identifier: u64,
    style_identifier: u64,
    style_ids: [u64; MAX_STYLE_INHERITANCE_DEPTH],
    style_count: usize,
    style_preset_identifier: Option<u64>,
    stylesheet_identifier: Option<u64>,
    before: Appearance,
    locked: bool,
}

/// A reversible process-local exact-source appearance patch.
#[derive(Clone, PartialEq, Eq)]
pub struct SlideTableAppearancePatch {
    source: Arc<[u8]>,
    target: Arc<[u8]>,
    path: SlideTableAppearancePath,
    before: Appearance,
    after: Appearance,
    touched_components: usize,
    deleted_previews: usize,
    changed_members: Arc<[String]>,
    preview_names: Arc<[&'static str]>,
}

impl fmt::Debug for SlideTableAppearancePatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("SlideTableAppearancePatch")
            .field("path", &self.path)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl SlideTableAppearancePatch {
    #[must_use]
    pub const fn path(&self) -> SlideTableAppearancePath {
        self.path
    }

    #[must_use]
    pub const fn before(&self) -> Appearance {
        self.before
    }

    #[must_use]
    pub const fn after(&self) -> Appearance {
        self.after
    }

    #[must_use]
    pub fn source_fingerprint(&self) -> u64 {
        fingerprint(&self.source)
    }

    #[must_use]
    pub fn target_fingerprint(&self) -> u64 {
        fingerprint(&self.target)
    }

    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.source == self.target
    }

    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: Arc::clone(&self.target),
            target: Arc::clone(&self.source),
            path: self.path,
            before: self.after,
            after: self.before,
            touched_components: self.touched_components,
            deleted_previews: self.deleted_previews,
            changed_members: Arc::clone(&self.changed_members),
            preview_names: Arc::clone(&self.preview_names),
        }
    }
}

/// Content-free publication diagnostics.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct SlideTableAppearanceDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl SlideTableAppearanceDiagnostics {
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

/// One fully validated immutable appearance publication.
#[must_use = "an appearance commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct SlideTableAppearanceCommit {
    package: Package,
    patch: SlideTableAppearancePatch,
    diagnostics: SlideTableAppearanceDiagnostics,
}

impl SlideTableAppearanceCommit {
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    #[must_use]
    pub const fn patch(&self) -> &SlideTableAppearancePatch {
        &self.patch
    }

    #[must_use]
    pub const fn diagnostics(&self) -> &SlideTableAppearanceDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read one rooted slide table's effective appearance.
    pub fn slide_table_appearance<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<Appearance, SlideTableAppearanceError> {
        let mut budget = AppearanceBudget::new(self)?;
        let catalog = physical_source(self)?;
        budget.source_catalog(self, catalog)?;
        Ok(select_table_with_budget(self, slide.into(), table.into(), &mut budget)?.before)
    }

    /// Start a selector-first immutable appearance edit.
    pub fn edit_slide_table_appearance<'slide>(
        &self,
        slide: impl Into<SlideSelector<'slide>>,
        table: impl Into<TableSelector>,
    ) -> Result<SlideTableAppearanceEdit<'_>, SlideTableAppearanceError> {
        let mut budget = AppearanceBudget::new(self)?;
        let catalog = physical_source(self)?;
        budget.source_catalog(self, catalog)?;
        let selection = select_table_with_budget(self, slide.into(), table.into(), &mut budget)?;
        Ok(SlideTableAppearanceEdit {
            source: self,
            path: SlideTableAppearancePath::Table {
                slide: selection.slide_position,
                table: selection.table_position,
            },
            before: selection.before,
            after: selection.before,
            selection,
        })
    }

    /// Apply an exact-source reversible appearance patch.
    pub fn apply_slide_table_appearance(
        &self,
        patch: &SlideTableAppearancePatch,
    ) -> Result<SlideTableAppearanceCommit, SlideTableAppearanceError> {
        apply_patch(self, patch)
    }
}

fn commit_edit(
    edit: SlideTableAppearanceEdit<'_>,
) -> Result<SlideTableAppearanceCommit, SlideTableAppearanceError> {
    let source_catalog = physical_source(edit.source)?;
    let source_bytes = source_catalog.shared_source();
    if edit.before == edit.after {
        return Ok(SlideTableAppearanceCommit {
            package: edit.source.snapshot(),
            patch: SlideTableAppearancePatch {
                source: Arc::clone(&source_bytes),
                target: source_bytes,
                path: edit.path,
                before: edit.before,
                after: edit.after,
                touched_components: 0,
                deleted_previews: 0,
                changed_members: empty_changed_members(),
                preview_names: empty_preview_names(),
            },
            diagnostics: SlideTableAppearanceDiagnostics {
                changed: false,
                touched_components: 0,
                deleted_previews: 0,
                full_reparse_performed: false,
            },
        });
    }
    if !source_catalog.source_is_exact() {
        return Err(SlideTableAppearanceError::UnsupportedSource);
    }
    if edit.selection.locked {
        return Err(SlideTableAppearanceError::Locked);
    }
    let mut budget = AppearanceBudget::new(edit.source)?;
    budget.source_catalog(edit.source, source_catalog)?;
    let (mut edits, new_style, new_uuid) = rewrite_native_appearance(
        edit.source,
        &edit.selection,
        edit.after,
        edit.path,
        &mut budget,
    )?;
    budget.allocations(1)?;
    budget.scratch(size_of::<NativeAppearanceEdit>())?;
    edits
        .try_reserve_exact(1)
        .map_err(|_| SlideTableAppearanceError::Allocation { amount: 1 })?;
    let metadata_edit = rewrite_metadata_for_style(
        edit.source,
        &edit.selection,
        new_style,
        new_uuid,
        edit.path,
        &mut budget,
    )?;
    edits.push(metadata_edit);
    let preview_plan =
        super::rendering_invalidation::root_preview_deletions(source_catalog.package())
            .map_err(|_| SlideTableAppearanceError::InvalidSource)?;
    let deleted_previews = preview_plan.len();
    let preview_names: Arc<[&'static str]> = Arc::from(preview_plan.names());
    let changed_members = changed_member_names(&edits, &mut budget)?;
    let touched_components = changed_members.len();
    budget.entries(edits.len())?;
    budget.allocations(edits.len())?;
    budget.retained(
        edits
            .iter()
            .map(|edit| edit.data.len())
            .try_fold(0usize, |total, size| total.checked_add(size))
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    budget.allocations(1)?;
    budget.scratch(
        edits
            .len()
            .checked_mul(size_of::<EntryEdit<'_>>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut entry_edits = Vec::new();
    entry_edits.try_reserve_exact(edits.len()).map_err(|_| {
        SlideTableAppearanceError::Allocation {
            amount: edits.len(),
        }
    })?;
    for edit in &edits {
        entry_edits.push(EntryEdit::new(edit.name.as_str(), edit.data.as_slice()));
    }
    let prepared = source_catalog
        .prepare_reassembly_with_deletions(
            &entry_edits,
            preview_plan.names(),
            edit.source.state.options.archive(),
        )
        .map_err(map_archive_error)?;
    let requirements = prepared.execution_requirements();
    budget.reassembly(requirements)?;
    let target_bytes = prepared
        .execute(requirements.exact_limits())
        .map_err(map_archive_error)?;
    let target_bytes: Arc<[u8]> = target_bytes.into();
    budget.candidate_reopen(target_bytes.len())?;
    let candidate =
        Package::from_source_with_options(Arc::clone(&target_bytes), edit.source.state.options)
            .map_err(map_read_error)?;
    let candidate_catalog = physical_source(&candidate)?;
    budget.source_catalog(&candidate, candidate_catalog)?;
    let selected = select_table_with_budget(
        &candidate,
        SlideSelector::Position(edit.selection.slide_position),
        TableSelector::index(edit.selection.table_position.get()),
        &mut budget,
    )?;
    if selected.before != edit.after {
        return Err(SlideTableAppearanceError::Verification);
    }
    verify_candidate_locality(
        edit.source,
        &candidate,
        changed_members.as_ref(),
        preview_names.as_ref(),
        &mut budget,
    )?;
    Ok(SlideTableAppearanceCommit {
        package: candidate,
        patch: SlideTableAppearancePatch {
            source: source_bytes,
            target: target_bytes,
            path: edit.path,
            before: edit.before,
            after: edit.after,
            touched_components,
            deleted_previews,
            changed_members,
            preview_names,
        },
        diagnostics: SlideTableAppearanceDiagnostics {
            changed: true,
            touched_components,
            deleted_previews,
            full_reparse_performed: true,
        },
    })
}

fn apply_patch(
    package: &Package,
    patch: &SlideTableAppearancePatch,
) -> Result<SlideTableAppearanceCommit, SlideTableAppearanceError> {
    let source_catalog = physical_source(package)?;
    let current_source = source_catalog.shared_source();
    if current_source.as_ref() != patch.source.as_ref() {
        return Err(SlideTableAppearanceError::PatchConflict);
    }
    if patch.is_noop() {
        return Ok(SlideTableAppearanceCommit {
            package: package.snapshot(),
            patch: patch.clone(),
            diagnostics: SlideTableAppearanceDiagnostics {
                changed: false,
                touched_components: 0,
                deleted_previews: 0,
                full_reparse_performed: false,
            },
        });
    }
    let mut budget = AppearanceBudget::new(package)?;
    budget.source_catalog(package, source_catalog)?;
    let current = select_table_with_budget(
        package,
        SlideSelector::Position(match patch.path {
            SlideTableAppearancePath::Table { slide, .. } => slide,
            SlideTableAppearancePath::Package => {
                return Err(SlideTableAppearanceError::PatchConflict);
            },
        }),
        TableSelector::index(match patch.path {
            SlideTableAppearancePath::Table { table, .. } => table.get(),
            SlideTableAppearancePath::Package => unreachable!(),
        }),
        &mut budget,
    )?;
    if current.before != patch.before {
        return Err(SlideTableAppearanceError::PatchConflict);
    }
    let target_bytes = Arc::clone(&patch.target);
    budget.candidate_reopen(target_bytes.len())?;
    let candidate = Package::from_source_with_options(target_bytes, package.state.options)
        .map_err(map_read_error)?;
    let candidate_catalog = physical_source(&candidate)?;
    budget.source_catalog(&candidate, candidate_catalog)?;
    let selected = select_table_with_budget(
        &candidate,
        SlideSelector::Position(current.slide_position),
        TableSelector::index(current.table_position.get()),
        &mut budget,
    )?;
    if selected.before != patch.after {
        return Err(SlideTableAppearanceError::Verification);
    }
    verify_candidate_locality(
        package,
        &candidate,
        patch.changed_members.as_ref(),
        patch.preview_names.as_ref(),
        &mut budget,
    )?;
    Ok(SlideTableAppearanceCommit {
        package: candidate,
        patch: patch.clone(),
        diagnostics: SlideTableAppearanceDiagnostics {
            changed: true,
            touched_components: patch.touched_components,
            deleted_previews: patch.deleted_previews,
            full_reparse_performed: true,
        },
    })
}

#[derive(Debug)]
struct NativeAppearanceEdit {
    name: String,
    data: Vec<u8>,
}

fn rewrite_native_appearance(
    source: &Package,
    selection: &AppearanceSelection,
    appearance: Appearance,
    path: SlideTableAppearancePath,
    budget: &mut AppearanceBudget,
) -> Result<(Vec<NativeAppearanceEdit>, u64, UuidBits), SlideTableAppearanceError> {
    let old_style = selection.style_identifier;
    if old_style == 0 || selection.model_style_identifier == 0 {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
    }
    let style_component = selection
        .style_component
        .as_deref()
        .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
    let stylesheet_identifier = selection
        .stylesheet_identifier
        .filter(|identifier| *identifier != 0)
        .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
    let stylesheet_component =
        component_name_for_identifier(source, stylesheet_identifier, path, budget)?;
    let model_payload = message_data(
        source,
        selection.model_identifier,
        selection.model_message_index,
        path,
    )?;
    let style_payload = style_message_data(source, old_style, path)?;
    let (_, old_style_object) = source
        .object_with_component(old_style)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let (old_style_message_index, _) = unique_message(old_style_object, TABLE_STYLE_MESSAGE_TYPE)?;
    let old_style_message_info = old_style_object
        .archive_info
        .message_infos
        .get(old_style_message_index)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let old_style_parent = (selection.style_count > 1).then_some(selection.style_ids[1]);
    let stylesheet_payload = stylesheet_message_data(source, stylesheet_identifier, path)?;
    let new_style = next_identifier(source, path, budget)?;
    let new_uuid = UuidBits::new(
        new_style ^ 0x9e37_79b9_7f4a_7c15,
        new_style.rotate_left(29) ^ 0xd1b5_4a32_d192_ed03,
    );
    if new_uuid.lower() == 0 && new_uuid.upper() == 0 {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let codec_work = model_payload
        .len()
        .checked_add(style_payload.len())
        .and_then(|size| size.checked_add(stylesheet_payload.len()))
        .and_then(|size| size.checked_mul(16))
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    budget.transaction_work(codec_work)?;
    budget.scratch(codec_work)?;
    let options = budget.codec_options(source, model_payload)?;
    let variation = appearance_codec::canonical_table_style_variation(
        appearance_codec::TableStyleVariationWrite {
            parent_identifier: old_style,
            stylesheet_identifier,
            overrides: appearance_overrides(appearance),
        },
        budget.codec_options(source, style_payload)?,
    )
    .map_err(map_appearance_codec_error)?;
    // The canonical writer owns the exact field/work/output/allocation
    // accounting for a fresh variation.  Merge that report after the
    // writer's bounded preflight; a rough size guess here would both miss
    // fields/allocations and double-charge output.
    let variation_report = variation.report();
    budget.output(variation_report.output_bytes())?;
    budget.codec_report(variation_report)?;
    let model_plan = appearance_codec::prepare_table_model_style_rewrite(
        model_payload,
        selection.model_style_identifier,
        new_style,
        options,
    )
    .map_err(map_appearance_codec_error)?;
    let model_requirements = model_plan.execution_requirements();
    let model_prepare_report = model_plan.prepare_report();
    budget.codec_report(model_prepare_report)?;
    budget.codec_requirements(model_prepare_report, model_requirements)?;
    let (model_rewritten, model_report) = model_plan
        .execute(model_requirements.exact_limits())
        .map_err(map_appearance_codec_error)?;
    if model_report.output_bytes() != model_requirements.output_bytes()
        || model_report.fields() != model_requirements.fields()
        || model_report.work_bytes() != model_requirements.work_bytes()
        || model_report.allocations() != model_requirements.allocations()
    {
        return Err(SlideTableAppearanceError::Verification);
    }
    let stylesheet_plan = appearance_codec::prepare_stylesheet_append(
        stylesheet_payload,
        appearance_codec::StylesheetStyleAppend {
            style_identifier: new_style,
            parent_identifier: Some(old_style),
        },
        budget.codec_options(source, stylesheet_payload)?,
    )
    .map_err(map_appearance_codec_error)?;
    let stylesheet_requirements = stylesheet_plan.execution_requirements();
    let stylesheet_prepare_report = stylesheet_plan.prepare_report();
    budget.codec_report(stylesheet_prepare_report)?;
    budget.codec_requirements(stylesheet_prepare_report, stylesheet_requirements)?;
    let (stylesheet_rewritten, stylesheet_report) = stylesheet_plan
        .execute(stylesheet_requirements.exact_limits())
        .map_err(map_appearance_codec_error)?;
    if stylesheet_report.output_bytes() != stylesheet_requirements.output_bytes()
        || stylesheet_report.fields() != stylesheet_requirements.fields()
        || stylesheet_report.work_bytes() != stylesheet_requirements.work_bytes()
        || stylesheet_report.allocations() != stylesheet_requirements.allocations()
    {
        return Err(SlideTableAppearanceError::Verification);
    }
    budget.allocations(4)?;
    budget.scratch(
        3usize
            .checked_mul(size_of::<String>())
            .and_then(|size| size.checked_add(selection.model_component.len()))
            .and_then(|size| size.checked_add(style_component.len()))
            .and_then(|size| size.checked_add(stylesheet_component.len()))
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut component_names = Vec::new();
    component_names
        .try_reserve_exact(3)
        .map_err(|_| SlideTableAppearanceError::Allocation { amount: 3 })?;
    component_names.push(selection.model_component.to_string());
    component_names.push(style_component.to_owned());
    component_names.push(stylesheet_component.clone());
    component_names.sort_unstable();
    component_names.dedup();
    let physical = physical_source(source)?;
    let entries = entry_index(
        physical.package(),
        budget,
        SlideTableAppearanceError::InvalidSource,
    )?;
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = source
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(map_archive_error)?;
    budget.allocations(1)?;
    budget.scratch(
        component_names
            .len()
            .saturating_mul(size_of::<NativeAppearanceEdit>()),
    )?;
    let mut edits = Vec::new();
    edits
        .try_reserve_exact(component_names.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: component_names.len(),
        })?;
    for name in component_names {
        let entry = entries
            .get(name.as_str())
            .copied()
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        if entry.is_opaque() {
            return Err(SlideTableAppearanceError::UnsupportedSource);
        }
        budget.physical(entry.data().len())?;
        budget.scratch(entry.data().len())?;
        budget.allocations(2)?;
        let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
            .map_err(map_core_error)?;
        budget.physical(stream.as_bytes().len())?;
        budget.retained(stream.as_bytes().len())?;
        let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
            .map_err(map_core_error)?;
        budget.objects(archive.objects.len())?;
        budget.messages(
            archive
                .objects
                .iter()
                .map(|object| object.messages.len())
                .try_fold(0usize, |total, count| total.checked_add(count))
                .ok_or(SlideTableAppearanceError::InvalidSource)?,
        )?;
        archive
            .validate_canonical_object_framing(stream.as_bytes())
            .map_err(map_core_error)?;
        if name == selection.model_component.as_ref() {
            replace_model_in_archive(
                &mut archive,
                selection.model_identifier,
                selection.model_style_identifier,
                new_style,
                model_rewritten.as_slice(),
                archive_limits,
                path,
                budget,
            )?;
        }
        if name == style_component {
            append_style_object(
                &mut archive,
                new_style,
                old_style_message_info,
                old_style_parent,
                old_style,
                stylesheet_identifier,
                variation.bytes(),
                archive_limits,
                path,
                budget,
            )?;
        }
        if name == stylesheet_component {
            replace_stylesheet_in_archive(
                &mut archive,
                stylesheet_identifier,
                new_style,
                stylesheet_rewritten.as_slice(),
                archive_limits,
                path,
                budget,
            )?;
        }
        let encoded_len = archive
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_error)?;
        budget.output(encoded_len)?;
        budget.entry_bytes(encoded_len)?;
        budget.total_bytes(encoded_len)?;
        let compressed_bound =
            SnappyStream::maximum_compressed_len(encoded_len).map_err(map_core_error)?;
        budget.output(compressed_bound)?;
        budget.wire_output(compressed_bound)?;
        budget.entry_bytes(compressed_bound)?;
        budget.scratch(compressed_bound)?;
        let bytes = archive
            .to_bytes_with_limits(archive_limits)
            .map_err(map_core_error)?;
        if bytes.len() != encoded_len {
            return Err(SlideTableAppearanceError::Verification);
        }
        let data = SnappyStream::compress(&bytes).map_err(map_core_error)?;
        if data.len() > compressed_bound {
            return Err(SlideTableAppearanceError::Verification);
        }
        budget.retained(data.len())?;
        edits.push(NativeAppearanceEdit { name, data });
    }
    Ok((edits, new_style, new_uuid))
}

fn changed_member_names(
    edits: &[NativeAppearanceEdit],
    budget: &mut AppearanceBudget,
) -> Result<Arc<[String]>, SlideTableAppearanceError> {
    budget.allocations(1)?;
    budget.scratch(
        edits
            .len()
            .checked_mul(size_of::<String>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    budget.allocations(edits.len())?;
    let mut names = Vec::new();
    names
        .try_reserve_exact(edits.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: edits.len(),
        })?;
    for edit in edits {
        budget.scratch(edit.name.len())?;
        names.push(edit.name.clone());
    }
    names.sort_unstable();
    names.dedup();
    Ok(names.into())
}

fn empty_changed_members() -> Arc<[String]> {
    Vec::<String>::new().into()
}

fn empty_preview_names() -> Arc<[&'static str]> {
    Vec::<&'static str>::new().into()
}

fn appearance_overrides(appearance: Appearance) -> appearance_codec::AppearanceOverrides {
    appearance_codec::AppearanceOverrides {
        row_banding: Some(matches!(appearance.row_banding, Banding::Enabled)),
        row_sizing: Some(matches!(appearance.row_sizing, RowSizing::FitCellContents)),
        body_horizontal: Some(matches!(
            appearance.gridlines.body_horizontal,
            GridlineVisibility::Visible
        )),
        body_vertical: Some(matches!(
            appearance.gridlines.body_vertical,
            GridlineVisibility::Visible
        )),
        header_columns_horizontal: Some(matches!(
            appearance.gridlines.header_columns_horizontal,
            GridlineVisibility::Visible
        )),
        header_rows_vertical: Some(matches!(
            appearance.gridlines.header_rows_vertical,
            GridlineVisibility::Visible
        )),
        footer_rows_vertical: Some(matches!(
            appearance.gridlines.footer_rows_vertical,
            GridlineVisibility::Visible
        )),
    }
}

fn component_name_for_identifier(
    source: &Package,
    identifier: u64,
    _path: SlideTableAppearancePath,
    budget: &mut AppearanceBudget,
) -> Result<String, SlideTableAppearanceError> {
    let (name, _) = source
        .object_with_component(identifier)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    budget.allocations(1)?;
    budget.scratch(name.len())?;
    Ok(name.to_owned())
}

fn message_data(
    source: &Package,
    identifier: u64,
    message_index: usize,
    _path: SlideTableAppearancePath,
) -> Result<&[u8], SlideTableAppearanceError> {
    let (_, object) = source
        .object_with_component(identifier)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let message = object
        .messages
        .get(message_index)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    if message.type_ != TABLE_MODEL_MESSAGE_TYPE {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    Ok(message.data.as_slice())
}

fn style_message_data(
    source: &Package,
    identifier: u64,
    _path: SlideTableAppearancePath,
) -> Result<&[u8], SlideTableAppearanceError> {
    let (_, object) = source
        .object_with_component(identifier)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let mut messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == TABLE_STYLE_MESSAGE_TYPE);
    let message = messages
        .next()
        .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
    if messages.next().is_some() {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    Ok(message.data.as_slice())
}

fn stylesheet_message_data(
    source: &Package,
    identifier: u64,
    _path: SlideTableAppearancePath,
) -> Result<&[u8], SlideTableAppearanceError> {
    let (_, object) = source
        .object_with_component(identifier)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let mut messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == TABLE_STYLESHEET_MESSAGE_TYPE);
    let message = messages
        .next()
        .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
    if messages.next().is_some() {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    Ok(message.data.as_slice())
}

fn next_identifier(
    source: &Package,
    path: SlideTableAppearancePath,
    budget: &mut AppearanceBudget,
) -> Result<u64, SlideTableAppearanceError> {
    budget.transaction_work(
        source
            .state
            .source
            .components()
            .iter()
            .map(|component| component.archive().objects.len())
            .try_fold(0usize, |total, count| total.checked_add(count))
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut maximum = 0u64;
    for component in source.state.source.components().iter() {
        for object in &component.archive().objects {
            maximum = maximum.max(
                object
                    .archive_info
                    .identifier
                    .ok_or(SlideTableAppearanceError::InvalidSource)?,
            );
        }
    }
    let metadata = metadata_payload(source, path, budget)?;
    budget.physical(metadata.len())?;
    budget.allocations(4)?;
    let mut facts = MetadataFacts::default();
    budget.scratch(MetadataFacts::reservation_scratch(metadata.len())?)?;
    facts.reserve_for_source(metadata.len())?;
    let inspection = inspect_package_metadata_with_visitor(
        metadata,
        budget.metadata_options(source, metadata.len(), 4),
        &mut facts,
    )
    .map_err(map_metadata_error)?;
    budget.metadata_report(inspection.report())?;
    if facts.unknown {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
    }
    maximum = maximum
        .max(inspection.last_object_identifier())
        .max(facts.maximum);
    maximum
        .checked_add(1)
        .filter(|identifier| *identifier != 0)
        .ok_or(SlideTableAppearanceError::LimitExceeded {
            kind: SlideTableAppearanceLimitKind::PayloadObjects,
            observed: u64::MAX,
            maximum: u64::MAX - 1,
        })
}

#[derive(Debug, Clone)]
struct MetadataRoute {
    component_name: String,
    object_identifier: u64,
    message_index: usize,
}

fn metadata_route(
    source: &Package,
    _path: SlideTableAppearancePath,
    budget: &mut AppearanceBudget,
) -> Result<MetadataRoute, SlideTableAppearanceError> {
    let mut found = None;
    for component in source.state.source.components().iter() {
        for object in &component.archive().objects {
            for (index, message) in object.messages.iter().enumerate() {
                if message.type_ != METADATA_MESSAGE_TYPE {
                    continue;
                }
                let object_identifier = object
                    .archive_info
                    .identifier
                    .ok_or(SlideTableAppearanceError::InvalidSource)?;
                budget.allocations(1)?;
                budget.scratch(component.name().len())?;
                if found
                    .replace(MetadataRoute {
                        component_name: component.name().to_owned(),
                        object_identifier,
                        message_index: index,
                    })
                    .is_some()
                {
                    return Err(SlideTableAppearanceError::InvalidSource);
                }
            }
        }
    }
    found.ok_or(SlideTableAppearanceError::InvalidSource)
}

fn metadata_payload_at<'source>(
    source: &'source Package,
    route: &MetadataRoute,
) -> Result<&'source [u8], SlideTableAppearanceError> {
    let (component, object) = source
        .object_with_component(route.object_identifier)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    if component != route.component_name {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    object
        .messages
        .get(route.message_index)
        .filter(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .map(|message| message.data.as_slice())
        .ok_or(SlideTableAppearanceError::InvalidSource)
}

fn metadata_payload<'source>(
    source: &'source Package,
    path: SlideTableAppearancePath,
    budget: &mut AppearanceBudget,
) -> Result<&'source [u8], SlideTableAppearanceError> {
    let route = metadata_route(source, path, budget)?;
    metadata_payload_at(source, &route)
}

fn metadata_locator(name: &str) -> &str {
    name.strip_prefix("Index/")
        .and_then(|name| name.strip_suffix(".iwa"))
        .unwrap_or(name)
}

fn metadata_options(source: &Package, bytes: usize, additions: usize) -> MetadataRewriteOptions {
    let limits = source.state.options.archive();
    let max_input = limits.max_iwa_stream_bytes();
    MetadataRewriteOptions::new(
        bytes.max(1).min(max_input),
        bytes.saturating_add(4096).max(1).min(max_input),
        bytes.saturating_mul(64).clamp(1, WireLimits::MAX_FIELDS),
        bytes
            .saturating_mul(256)
            .saturating_add(additions.saturating_mul(bytes))
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        64,
        bytes
            .max(source.state.source.components().len())
            .saturating_mul(4)
            .max(1),
        source
            .state
            .options
            .semantic()
            .max_references()
            .max(bytes)
            .clamp(1, WireLimits::MAX_FIELDS),
        additions.max(1),
    )
}

#[derive(Default)]
struct MetadataFacts {
    maximum: u64,
    unknown: bool,
    components: Vec<MetadataComponentFact>,
    uuids: Vec<MetadataUuidFact>,
    external: Vec<MetadataExternalFact>,
    /// Every identifier-bearing metadata namespace is retained in this
    /// compact census.  A fresh native object must not merely clear the UUID
    /// table: component, external, data-owner, ambiguous, and root-map
    /// namespaces are all authoritative identifiers too.
    identifiers: Vec<u64>,
    /// Data-reference identifiers admitted by the current Metadata registry.
    data_identifiers: Vec<u64>,
}

impl MetadataFacts {
    fn reservation_slots(source_bytes: usize) -> Result<usize, SlideTableAppearanceError> {
        source_bytes
            .checked_div(2)
            .and_then(|slots| slots.checked_add(1))
            .ok_or(SlideTableAppearanceError::InvalidSource)
    }

    fn reservation_scratch(source_bytes: usize) -> Result<usize, SlideTableAppearanceError> {
        let slots = Self::reservation_slots(source_bytes)?;
        let identifier_slots = slots
            .checked_mul(5)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        let fact_bytes = size_of::<MetadataComponentFact>()
            .checked_add(size_of::<MetadataUuidFact>())
            .and_then(|size| size.checked_add(size_of::<MetadataExternalFact>()))
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        slots
            .checked_mul(fact_bytes)
            .and_then(|size| size.checked_add(identifier_slots.checked_mul(size_of::<u64>())?))
            .and_then(|size| size.checked_add(source_bytes))
            .ok_or(SlideTableAppearanceError::InvalidSource)
    }

    fn reserve_for_source(&mut self, source_bytes: usize) -> Result<(), SlideTableAppearanceError> {
        let slots = Self::reservation_slots(source_bytes)?;
        self.components
            .try_reserve_exact(slots)
            .map_err(|_| SlideTableAppearanceError::Allocation { amount: slots })?;
        self.uuids
            .try_reserve_exact(slots)
            .map_err(|_| SlideTableAppearanceError::Allocation { amount: slots })?;
        self.external
            .try_reserve_exact(slots)
            .map_err(|_| SlideTableAppearanceError::Allocation { amount: slots })?;
        let identifier_slots = slots
            .checked_mul(5)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        self.identifiers
            .try_reserve_exact(identifier_slots)
            .map_err(|_| SlideTableAppearanceError::Allocation {
                amount: identifier_slots,
            })?;
        self.data_identifiers
            .try_reserve_exact(slots)
            .map_err(|_| SlideTableAppearanceError::Allocation { amount: slots })?;
        Ok(())
    }
}

struct MetadataComponentFact {
    identifier: u64,
    locator: String,
    current: bool,
}

struct MetadataUuidFact {
    component_identifier: u64,
    current: bool,
    object_identifier: u64,
    uuid: UuidBits,
}

struct MetadataExternalFact {
    source_identifier: u64,
    source_locator: String,
    source_current: bool,
    target_identifier: u64,
    object_identifier: Option<u64>,
    weak: Option<bool>,
    versioned: bool,
}

fn unique_metadata_component<'a>(
    components: &'a [MetadataComponentFact],
    locator: &str,
) -> Option<&'a MetadataComponentFact> {
    let mut found = None;
    for component in components {
        if !component.current || component.locator != locator {
            continue;
        }
        if found.is_some() {
            return None;
        }
        found = Some(component);
    }
    found
}

fn validate_metadata_uuid_registry(
    facts: &MetadataFacts,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    budget.allocations(1)?;
    budget.scratch(
        facts
            .uuids
            .len()
            .checked_mul(size_of::<UuidBits>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    budget.transaction_work(facts.uuids.len())?;
    let mut uuids = HashSet::new();
    uuids
        .try_reserve(facts.uuids.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: facts.uuids.len(),
        })?;
    for binding in &facts.uuids {
        if !uuids.insert(binding.uuid) {
            return Err(SlideTableAppearanceError::UnsupportedDependency);
        }
    }
    Ok(())
}

fn require_current_metadata_uuid(
    facts: &MetadataFacts,
    component_identifier: u64,
    object_identifier: u64,
) -> Result<(), SlideTableAppearanceError> {
    let mut current = 0usize;
    for binding in &facts.uuids {
        if binding.object_identifier != object_identifier {
            continue;
        }
        if !binding.current || binding.component_identifier != component_identifier {
            return Err(SlideTableAppearanceError::UnsupportedDependency);
        }
        current = current.saturating_add(1);
    }
    if current == 1 {
        Ok(())
    } else {
        Err(SlideTableAppearanceError::UnsupportedDependency)
    }
}

fn require_optional_current_metadata_uuid(
    facts: &MetadataFacts,
    component_identifier: u64,
    object_identifier: u64,
) -> Result<(), SlideTableAppearanceError> {
    let mut current = 0usize;
    let mut registered = 0usize;
    for binding in &facts.uuids {
        if binding.object_identifier != object_identifier {
            continue;
        }
        registered = registered.saturating_add(1);
        if !binding.current || binding.component_identifier != component_identifier {
            return Err(SlideTableAppearanceError::UnsupportedDependency);
        }
        current = current.saturating_add(1);
    }
    if registered == 0 || (registered == 1 && current == 1) {
        Ok(())
    } else {
        Err(SlideTableAppearanceError::UnsupportedDependency)
    }
}

impl PackageMetadataVisitor for MetadataFacts {
    fn visit_unknown_field(&mut self) -> Result<(), MetadataRewriteError> {
        self.unknown = true;
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: metadata_codec::ComponentDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.maximum = self.maximum.max(component.identifier());
        self.identifiers.push(component.identifier());
        self.components.push(MetadataComponentFact {
            identifier: component.identifier(),
            locator: component.effective_locator().to_owned(),
            current: component.is_current(),
        });
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        let component = binding.component();
        self.maximum = self.maximum.max(binding.object_identifier());
        self.identifiers.push(binding.component().identifier());
        self.identifiers.push(binding.object_identifier());
        self.uuids.push(MetadataUuidFact {
            component_identifier: component.identifier(),
            current: component.is_current(),
            object_identifier: binding.object_identifier(),
            uuid: binding.uuid(),
        });
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        let source = reference.source();
        self.maximum = self.maximum.max(reference.target_component_identifier());
        self.identifiers.push(source.identifier());
        self.identifiers
            .push(reference.target_component_identifier());
        if let Some(identifier) = reference.object_identifier() {
            self.maximum = self.maximum.max(identifier);
            self.identifiers.push(identifier);
        }
        self.external.push(MetadataExternalFact {
            source_identifier: source.identifier(),
            source_locator: source.effective_locator().to_owned(),
            source_current: source.is_current(),
            target_identifier: reference.target_component_identifier(),
            object_identifier: reference.object_identifier(),
            weak: reference.is_weak(),
            versioned: reference.is_versioned(),
        });
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        reference: metadata_codec::DataReferenceDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.maximum = self.maximum.max(reference.data_identifier());
        self.identifiers.push(reference.component().identifier());
        self.identifiers.push(reference.data_identifier());
        self.data_identifiers.push(reference.data_identifier());
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), MetadataRewriteError> {
        self.maximum = self.maximum.max(owner.data_identifier());
        self.maximum = self.maximum.max(owner.object_identifier());
        self.identifiers.push(owner.component().identifier());
        self.identifiers.push(owner.data_identifier());
        self.identifiers.push(owner.object_identifier());
        self.data_identifiers.push(owner.data_identifier());
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), MetadataRewriteError> {
        self.maximum = self.maximum.max(identifier);
        self.identifiers.push(identifier);
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), MetadataRewriteError> {
        self.maximum = self.maximum.max(object_identifier);
        self.identifiers.push(object_identifier);
        Ok(())
    }
}

fn rewrite_metadata_for_style(
    source: &Package,
    selection: &AppearanceSelection,
    new_style: u64,
    uuid: UuidBits,
    path: SlideTableAppearancePath,
    budget: &mut AppearanceBudget,
) -> Result<NativeAppearanceEdit, SlideTableAppearanceError> {
    // Keep the inspected physical route as part of the rewrite authority.
    // Re-scanning by a broad `*Metadata.iwa` suffix later could otherwise
    // silently rewrite a decoy while the validated entry remains untouched.
    let metadata_route = metadata_route(source, path, budget)?;
    let metadata = metadata_payload_at(source, &metadata_route)?;
    budget.physical(metadata.len())?;
    budget.allocations(4)?;
    let mut facts = MetadataFacts::default();
    budget.scratch(MetadataFacts::reservation_scratch(metadata.len())?)?;
    facts.reserve_for_source(metadata.len())?;
    let inspection = inspect_package_metadata_with_visitor(
        metadata,
        budget.metadata_options(source, metadata.len(), 2),
        &mut facts,
    )
    .map_err(map_metadata_error)?;
    budget.metadata_report(inspection.report())?;
    if facts.unknown {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
    }
    if selection.metadata_component.as_deref() != Some(metadata_route.component_name.as_str()) {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
    }
    let model_locator = metadata_locator(selection.model_component.as_ref());
    let style_component = selection
        .style_component
        .as_deref()
        .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
    let style_locator = metadata_locator(style_component);
    let stylesheet_identifier = selection
        .stylesheet_identifier
        .filter(|identifier| *identifier != 0)
        .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
    let stylesheet_component =
        component_name_for_identifier(source, stylesheet_identifier, path, budget)?;
    let stylesheet_locator = metadata_locator(&stylesheet_component);
    let model_component = unique_metadata_component(&facts.components, model_locator)
        .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
    let style_component_fact = unique_metadata_component(&facts.components, style_locator)
        .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
    let stylesheet_component_fact =
        unique_metadata_component(&facts.components, stylesheet_locator)
            .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
    let model_selector = ComponentSelector::new(model_component.identifier, model_locator);
    let style_selector = ComponentSelector::new(style_component_fact.identifier, style_locator);
    let stylesheet_selector =
        ComponentSelector::new(stylesheet_component_fact.identifier, stylesheet_locator);
    validate_metadata_uuid_registry(&facts, budget)?;
    require_current_metadata_uuid(
        &facts,
        model_component.identifier,
        selection.model_identifier,
    )?;
    require_optional_current_metadata_uuid(
        &facts,
        stylesheet_component_fact.identifier,
        stylesheet_identifier,
    )?;
    if selection.style_count == 0
        || selection.style_ids[0] != selection.style_identifier
        || selection.style_count > MAX_STYLE_INHERITANCE_DEPTH
    {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    for style_identifier in &selection.style_ids[..selection.style_count] {
        let style_component_name =
            component_name_for_identifier(source, *style_identifier, path, budget)?;
        let style_component_locator = metadata_locator(&style_component_name);
        let style_component_fact =
            unique_metadata_component(&facts.components, style_component_locator)
                .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
        require_current_metadata_uuid(&facts, style_component_fact.identifier, *style_identifier)?;
    }
    if facts.identifiers.contains(&new_style) {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let old_style = selection.style_identifier;
    if facts
        .uuids
        .iter()
        .any(|binding| binding.object_identifier == new_style || binding.uuid == uuid)
    {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let model_style_cross_component = model_selector.identifier() != style_selector.identifier();
    budget.allocations(2)?;
    budget.scratch(facts.external.len().saturating_mul(size_of::<u64>()))?;
    let mut model_style_objects = HashSet::new();
    model_style_objects
        .try_reserve(facts.external.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: facts.external.len(),
        })?;
    let mut selected_model_style_edges = 0usize;
    let mut model_style_weak = None;
    for reference in facts.external.iter().filter(|reference| {
        reference.source_identifier == model_selector.identifier()
            && reference.source_locator == model_locator
            && reference.target_identifier == style_selector.identifier()
    }) {
        let object_identifier = reference
            .object_identifier
            .filter(|identifier| *identifier != 0)
            .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
        if !reference.source_current
            || reference.versioned
            || reference.weak == Some(true)
            || !model_style_objects.insert(object_identifier)
            || source
                .object_with_component(object_identifier)
                .is_none_or(|(component, _)| component != style_component)
        {
            return Err(SlideTableAppearanceError::UnsupportedDependency);
        }
        if object_identifier == old_style {
            selected_model_style_edges = selected_model_style_edges.saturating_add(1);
            model_style_weak = reference.weak;
        }
    }
    if selected_model_style_edges != usize::from(model_style_cross_component)
        || (!model_style_cross_component && !model_style_objects.is_empty())
    {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
    }
    let stylesheet_cross_component =
        style_selector.identifier() != stylesheet_selector.identifier();
    let mut style_stylesheet_objects = HashSet::new();
    style_stylesheet_objects
        .try_reserve(facts.external.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: facts.external.len(),
        })?;
    let mut selected_style_stylesheet_edges = 0usize;
    for reference in facts.external.iter().filter(|reference| {
        reference.source_identifier == style_selector.identifier()
            && reference.source_locator == style_locator
            && reference.target_identifier == stylesheet_selector.identifier()
    }) {
        let object_identifier = reference
            .object_identifier
            .filter(|identifier| *identifier != 0)
            .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
        if !reference.source_current
            || reference.versioned
            || reference.weak == Some(true)
            || !style_stylesheet_objects.insert(object_identifier)
            || source
                .object_with_component(object_identifier)
                .is_none_or(|(component, _)| component != stylesheet_component)
        {
            return Err(SlideTableAppearanceError::UnsupportedDependency);
        }
        if object_identifier == stylesheet_identifier {
            selected_style_stylesheet_edges = selected_style_stylesheet_edges.saturating_add(1);
        }
    }
    if selected_style_stylesheet_edges != usize::from(stylesheet_cross_component)
        || (!stylesheet_cross_component && !style_stylesheet_objects.is_empty())
    {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
    }
    budget.allocations(4)?;
    budget.scratch(
        size_of::<ObjectUuidAddition<'_>>()
            .checked_add(size_of::<ExternalReferenceAddition<'_>>())
            .and_then(|size| size.checked_add(size_of::<ExternalReferenceRemoval<'_>>()))
            .and_then(|size| {
                size.checked_add(3usize.saturating_mul(size_of::<ComponentSelector<'_>>()))
            })
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut uuid_additions = Vec::new();
    uuid_additions
        .try_reserve_exact(1)
        .map_err(|_| SlideTableAppearanceError::Allocation { amount: 1 })?;
    uuid_additions.push(ObjectUuidAddition::new(style_selector, new_style, uuid));
    let mut external_additions = Vec::new();
    if model_selector.identifier() != style_selector.identifier() {
        external_additions
            .try_reserve_exact(1)
            .map_err(|_| SlideTableAppearanceError::Allocation { amount: 1 })?;
        external_additions.push(ExternalReferenceAddition::new(
            model_selector,
            style_selector,
            new_style,
            model_style_weak,
        ));
    }
    let mut external_removals = Vec::new();
    if model_style_cross_component {
        external_removals
            .try_reserve_exact(1)
            .map_err(|_| SlideTableAppearanceError::Allocation { amount: 1 })?;
        external_removals.push(ExternalReferenceRemoval::new(
            model_selector,
            style_selector,
            old_style,
            model_style_weak,
        ));
    }
    let mut selectors = Vec::new();
    selectors
        .try_reserve_exact(3)
        .map_err(|_| SlideTableAppearanceError::Allocation { amount: 2 })?;
    selectors.push(model_selector);
    for selector in [style_selector, stylesheet_selector] {
        if !selectors
            .iter()
            .any(|existing: &ComponentSelector<'_>| existing.identifier() == selector.identifier())
        {
            selectors.push(selector);
        }
    }
    let transition = CombinedBatch::new(
        inspection.last_object_identifier(),
        new_style,
        uuid_additions.as_slice(),
        external_additions.as_slice(),
        &[],
        external_removals.as_slice(),
        &[],
    );
    let batch = CombinedSaveTokenBatch::new(transition, SaveTokenBatch::new(selectors.as_slice()));
    let prepared = prepare_package_metadata_combined_additions_and_removals_and_save_tokens(
        metadata,
        batch,
        budget.metadata_options(
            source,
            metadata.len(),
            uuid_additions
                .len()
                .saturating_add(external_additions.len())
                .saturating_add(external_removals.len()),
        ),
    )
    .map_err(map_metadata_error)?;
    let prepare_report = prepared.prepare_report();
    let requirements = prepared.execution_requirements();
    budget.metadata_report(prepare_report)?;
    budget.metadata_requirements(prepare_report, requirements)?;
    let output = prepared
        .execute(requirements.exact_limits())
        .map_err(map_metadata_error)?;
    let report = output.report();
    if report.output_bytes() != requirements.output_bytes()
        || report.fields() != requirements.fields()
        || report.work_bytes() != requirements.work_bytes()
        || report.allocations() != requirements.allocations()
        || report.retained_bytes() != requirements.retained_bytes()
        || report.scratch_bytes() > requirements.scratch_bytes()
    {
        return Err(SlideTableAppearanceError::Verification);
    }
    rewrite_metadata_entry(source, output.into_bytes(), &metadata_route, path, budget)
}

fn rewrite_metadata_entry(
    source: &Package,
    payload: Vec<u8>,
    route: &MetadataRoute,
    _path: SlideTableAppearancePath,
    budget: &mut AppearanceBudget,
) -> Result<NativeAppearanceEdit, SlideTableAppearanceError> {
    let physical = physical_source(source)?;
    let entry = physical
        .package()
        .iter()
        .find(|entry| entry.name() == route.component_name && !entry.is_opaque())
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    budget.physical(entry.data().len())?;
    budget.scratch(entry.data().len())?;
    budget.allocations(2)?;
    let archive_limits = source
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let snappy_limits = source
        .state
        .options
        .archive()
        .snappy_limits()
        .map_err(map_archive_error)?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    budget.physical(stream.as_bytes().len())?;
    budget.retained(stream.as_bytes().len())?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    budget.objects(archive.objects.len())?;
    budget.messages(
        archive
            .objects
            .iter()
            .map(|object| object.messages.len())
            .try_fold(0usize, |total, count| total.checked_add(count))
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(map_core_error)?;
    let object = archive
        .object_mut(route.object_identifier)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    if object
        .messages
        .get(route.message_index)
        .is_none_or(|message| message.type_ != METADATA_MESSAGE_TYPE)
    {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    object
        .replace_message_preserving_header_with_limits(
            route.message_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: payload,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let encoded_len = archive
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget.output(encoded_len)?;
    budget.entry_bytes(encoded_len)?;
    budget.total_bytes(encoded_len)?;
    let compressed_bound =
        SnappyStream::maximum_compressed_len(encoded_len).map_err(map_core_error)?;
    budget.output(compressed_bound)?;
    budget.wire_output(compressed_bound)?;
    budget.entry_bytes(compressed_bound)?;
    budget.scratch(compressed_bound)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    if bytes.len() != encoded_len {
        return Err(SlideTableAppearanceError::Verification);
    }
    let data = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    if data.len() > compressed_bound {
        return Err(SlideTableAppearanceError::Verification);
    }
    budget.retained(data.len())?;
    budget.allocations(1)?;
    budget.scratch(route.component_name.len())?;
    Ok(NativeAppearanceEdit {
        name: entry.name().to_owned(),
        data,
    })
}

fn replace_model_in_archive(
    archive: &mut Archive,
    identifier: u64,
    old_style: u64,
    new_style: u64,
    replacement: &[u8],
    limits: litchi_iwa_core::Limits,
    path: SlideTableAppearancePath,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    let object = archive
        .object_mut(identifier)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let mut indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == TABLE_MODEL_MESSAGE_TYPE);
    let message_index = indexes
        .next()
        .map(|(index, _)| index)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    if indexes.next().is_some() {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let before_len = info.object_references.len();
    budget.allocations(1)?;
    budget.scratch(
        before_len
            .checked_mul(size_of::<u64>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let before = info.object_references.clone();
    let frequencies = reference_frequencies(&before, budget)?;
    if frequencies.get(&old_style).copied() != Some(1) || before.contains(&new_style) {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let after_len = before_len;
    budget.allocations(1)?;
    budget.scratch(
        after_len
            .checked_mul(size_of::<u64>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut after = Vec::new();
    after
        .try_reserve_exact(after_len)
        .map_err(|_| SlideTableAppearanceError::Allocation { amount: after_len })?;
    after.extend(before.iter().copied().filter(|value| *value != old_style));
    after.push(new_style);
    if after.len() != after_len {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    budget.allocations(1)?;
    budget.scratch(replacement.len())?;
    replace_message_with_reference_transition(
        object,
        message_index,
        RawMessage {
            type_: TABLE_MODEL_MESSAGE_TYPE,
            data: replacement.to_owned(),
        },
        before.as_slice(),
        after.as_slice(),
        &[3],
        Some((
            std::slice::from_ref(&old_style),
            std::slice::from_ref(&new_style),
        )),
        false,
        limits,
        path,
        budget,
    )
}

fn append_style_object(
    archive: &mut Archive,
    identifier: u64,
    source_message_info: &litchi_iwa_core::MessageInfo,
    source_parent: Option<u64>,
    new_parent: u64,
    stylesheet: u64,
    payload: &[u8],
    limits: litchi_iwa_core::Limits,
    _path: SlideTableAppearancePath,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    if archive
        .objects
        .iter()
        .any(|object| object.archive_info.identifier == Some(identifier))
    {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    budget.allocations(2)?;
    budget.scratch(payload.len())?;
    let mut object = ArchiveObject::new_with_limits(
        identifier,
        vec![RawMessage {
            type_: TABLE_STYLE_MESSAGE_TYPE,
            data: payload.to_owned(),
        }],
        limits,
    )
    .map_err(map_core_error)?;
    budget.allocations(
        8usize
            .checked_add(source_message_info.field_infos.len().saturating_mul(6))
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    budget.scratch(message_info_clone_scratch(source_message_info)?)?;
    let mut cloned_info = source_message_info.clone();
    cloned_info.length =
        u32::try_from(payload.len()).map_err(|_| SlideTableAppearanceError::InvalidSource)?;
    if !cloned_info.object_references.is_empty() {
        cloned_info.object_references.clear();
        cloned_info.object_references.push(new_parent);
        cloned_info.object_references.push(stylesheet);
        let mut parent_field_present = false;
        for field in &mut cloned_info.field_infos {
            if field.path.as_slice() == [1, 3] {
                field.object_references.clear();
                field.object_references.push(new_parent);
                parent_field_present = true;
            }
        }
        if source_parent.is_none() && !parent_field_present && !cloned_info.field_infos.is_empty() {
            budget.allocations(3)?;
            budget.scratch(
                size_of::<litchi_iwa_core::FieldInfo>()
                    .checked_add(size_of::<u64>())
                    .and_then(|size| size.checked_add(2usize.saturating_mul(size_of::<u32>())))
                    .ok_or(SlideTableAppearanceError::InvalidSource)?,
            )?;
            let mut field = litchi_iwa_core::FieldInfo::new(vec![1, 3]);
            field.r#type = Some(FieldType::ObjectReference);
            field.object_references.push(new_parent);
            cloned_info.field_infos.insert(0, field);
        }
    }
    object.archive_info.message_infos[0] = cloned_info;
    budget.allocations(1)?;
    archive
        .objects
        .try_reserve_exact(1)
        .map_err(|_| SlideTableAppearanceError::Allocation { amount: 1 })?;
    archive.objects.push(object);
    Ok(())
}

fn message_info_clone_scratch(
    info: &litchi_iwa_core::MessageInfo,
) -> Result<usize, SlideTableAppearanceError> {
    let u32_items = info
        .versions
        .len()
        .checked_add(info.diff_merge_version.len())
        .and_then(|count| count.checked_add(info.diff_read_version.len()))
        .and_then(|count| {
            info.fields_to_remove.iter().try_fold(count, |total, path| {
                total.checked_add(path.as_slice().len())
            })
        })
        .and_then(|count| {
            info.diff_field_path
                .as_ref()
                .map_or(Some(count), |path| count.checked_add(path.as_slice().len()))
        })
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let u64_items = info
        .object_references
        .len()
        .checked_add(info.data_references.len())
        .and_then(|count| {
            info.field_infos.iter().try_fold(count, |total, field| {
                total
                    .checked_add(field.object_references.len())
                    .and_then(|value| value.checked_add(field.data_references.len()))
            })
        })
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let field_u32_items = info
        .field_infos
        .iter()
        .try_fold(0usize, |total, field| {
            total
                .checked_add(field.path.as_slice().len())
                .and_then(|value| value.checked_add(field.known_field_version.len()))
        })
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let text_bytes = info
        .field_infos
        .iter()
        .try_fold(0usize, |total, field| {
            total.checked_add(
                field
                    .known_field_feature_identifier
                    .as_ref()
                    .map_or(0, String::len),
            )
        })
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    size_of::<litchi_iwa_core::MessageInfo>()
        .checked_add(
            info.field_infos
                .len()
                .saturating_mul(size_of::<litchi_iwa_core::FieldInfo>()),
        )
        .and_then(|size| size.checked_add(u32_items.saturating_mul(size_of::<u32>())))
        .and_then(|size| size.checked_add(field_u32_items.saturating_mul(size_of::<u32>())))
        .and_then(|size| size.checked_add(u64_items.saturating_mul(size_of::<u64>())))
        .and_then(|size| size.checked_add(text_bytes))
        .ok_or(SlideTableAppearanceError::InvalidSource)
}

fn replace_stylesheet_in_archive(
    archive: &mut Archive,
    identifier: u64,
    new_style: u64,
    replacement: &[u8],
    limits: litchi_iwa_core::Limits,
    path: SlideTableAppearancePath,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    let object = archive
        .object_mut(identifier)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let mut indexes = object
        .messages
        .iter()
        .enumerate()
        .filter(|(_, message)| message.type_ == TABLE_STYLESHEET_MESSAGE_TYPE);
    let message_index = indexes
        .next()
        .map(|(index, _)| index)
        .ok_or(SlideTableAppearanceError::UnsupportedDependency)?;
    if indexes.next().is_some() {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let before_len = info.object_references.len();
    budget.allocations(1)?;
    budget.scratch(
        before_len
            .checked_mul(size_of::<u64>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let before = info.object_references.clone();
    if before.contains(&new_style) {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let after_len = before_len
        .checked_add(1)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    budget.allocations(1)?;
    budget.scratch(
        after_len
            .checked_mul(size_of::<u64>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut after = Vec::new();
    after
        .try_reserve_exact(after_len)
        .map_err(|_| SlideTableAppearanceError::Allocation { amount: after_len })?;
    after.extend(before.iter().copied());
    after.push(new_style);
    budget.allocations(1)?;
    budget.scratch(replacement.len())?;
    replace_message_with_reference_transition(
        object,
        message_index,
        RawMessage {
            type_: TABLE_STYLESHEET_MESSAGE_TYPE,
            data: replacement.to_owned(),
        },
        before.as_slice(),
        after.as_slice(),
        &[1],
        Some((before.as_slice(), after.as_slice())),
        false,
        limits,
        path,
        budget,
    )
}

fn replace_message_with_reference_transition(
    object: &mut ArchiveObject,
    message_index: usize,
    message: RawMessage,
    before: &[u64],
    after: &[u64],
    field_path: &[u32],
    field_references: Option<(&[u64], &[u64])>,
    required_field_path: bool,
    limits: litchi_iwa_core::Limits,
    _path: SlideTableAppearancePath,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    let fields = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableAppearanceError::InvalidSource)?
        .field_infos
        .as_slice();
    let expected_before = field_references.map_or(before, |(old, _)| old);
    let expected_after = field_references.map_or(after, |(_, new)| new);
    budget.allocations(4)?;
    let field_scratch = fields
        .iter()
        .map(|field| field.object_references.len())
        .try_fold(0usize, |total, count| total.checked_add(count))
        .and_then(|count| count.checked_add(expected_after.len()))
        .and_then(|count| count.checked_mul(size_of::<u64>()))
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    budget.scratch(field_scratch)?;
    let mut field_indices = Vec::new();
    field_indices.try_reserve_exact(fields.len()).map_err(|_| {
        SlideTableAppearanceError::Allocation {
            amount: fields.len(),
        }
    })?;
    let mut field_before = Vec::new();
    field_before.try_reserve_exact(fields.len()).map_err(|_| {
        SlideTableAppearanceError::Allocation {
            amount: fields.len(),
        }
    })?;
    let mut field_after = Vec::new();
    field_after.try_reserve_exact(fields.len()).map_err(|_| {
        SlideTableAppearanceError::Allocation {
            amount: fields.len(),
        }
    })?;
    for (index, field) in fields.iter().enumerate() {
        if field.path.as_slice() != field_path {
            let references_changed = field.object_references.iter().any(|identifier| {
                (before.contains(identifier) && !after.contains(identifier))
                    || (after.contains(identifier) && !before.contains(identifier))
            });
            if references_changed {
                return Err(SlideTableAppearanceError::InvalidSource);
            }
            continue;
        }
        if !field_indices.is_empty() {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        if field
            .r#type
            .is_some_and(|kind| kind != FieldType::ObjectReference)
        {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        if field.object_references.as_slice() != expected_before {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        field_indices.push(index);
        field_before.push(field.object_references.clone());
        field_after.push(expected_after.to_vec());
    }
    if required_field_path && field_indices.len() != 1 {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let mut transitions = Vec::new();
    transitions
        .try_reserve_exact(field_indices.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: field_indices.len(),
        })?;
    for index in 0..field_indices.len() {
        transitions.push(FieldObjectReferenceTransition {
            field_info_index: field_indices[index],
            expected_path: field_path,
            before: field_before[index].as_slice(),
            after: field_after[index].as_slice(),
        });
    }
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            message_index,
            message,
            ObjectReferenceTransition {
                aggregate_before: before,
                aggregate_after: after,
                fields: transitions.as_slice(),
            },
            limits,
        )
        .map_err(map_core_error)?;
    Ok(())
}

fn verify_candidate_locality(
    source: &Package,
    candidate: &Package,
    expected_members: &[String],
    expected_preview_names: &[&'static str],
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    let source_catalog = physical_source(source)?;
    let candidate_catalog = physical_source(candidate)?;
    let scan_bytes = source_catalog
        .shared_source()
        .len()
        .checked_add(candidate_catalog.shared_source().len())
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    budget.transaction_work(scan_bytes)?;
    budget.work(scan_bytes)?;
    budget.allocations(1)?;
    budget.scratch(
        expected_members
            .len()
            .checked_mul(size_of::<String>())
            .and_then(|size| size.checked_add(scan_bytes))
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let source_previews =
        super::rendering_invalidation::root_preview_deletions(source_catalog.package())
            .map_err(|_| SlideTableAppearanceError::Verification)?;
    let candidate_previews =
        super::rendering_invalidation::root_preview_deletions(candidate_catalog.package())
            .map_err(|_| SlideTableAppearanceError::Verification)?;
    let preview_transition_is_forward =
        source_previews.names() == expected_preview_names && candidate_previews.names().is_empty();
    let preview_transition_is_inverse =
        source_previews.names().is_empty() && candidate_previews.names() == expected_preview_names;
    if !preview_transition_is_forward && !preview_transition_is_inverse {
        return Err(SlideTableAppearanceError::Verification);
    }
    let source_entries = entry_index(
        source_catalog.package(),
        budget,
        SlideTableAppearanceError::Verification,
    )?;
    let candidate_entries = entry_index(
        candidate_catalog.package(),
        budget,
        SlideTableAppearanceError::Verification,
    )?;
    let mut expected_set = HashSet::new();
    expected_set
        .try_reserve(expected_members.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: expected_members.len(),
        })?;
    budget.retained(
        expected_members
            .len()
            .checked_mul(size_of::<&str>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    for name in expected_members {
        expected_set.insert(name.as_str());
    }
    let mut before_entries = source_catalog
        .package()
        .iter()
        .filter(|entry| !source_previews.names().contains(&entry.name()));
    let mut after_entries = candidate_catalog
        .package()
        .iter()
        .filter(|entry| !candidate_previews.names().contains(&entry.name()));
    loop {
        match (before_entries.next(), after_entries.next()) {
            (Some(before), Some(after)) if before.name() == after.name() => {
                let indexed_before = source_entries
                    .get(before.name())
                    .copied()
                    .ok_or(SlideTableAppearanceError::Verification)?;
                let indexed_after = candidate_entries
                    .get(after.name())
                    .copied()
                    .ok_or(SlideTableAppearanceError::Verification)?;
                let changed = indexed_before.raw_name() != indexed_after.raw_name()
                    || indexed_before.data() != indexed_after.data()
                    || indexed_before.metadata() != indexed_after.metadata()
                    || indexed_before.raw_record().local_record()
                        != indexed_after.raw_record().local_record()
                    || indexed_before.raw_record().compressed_data()
                        != indexed_after.raw_record().compressed_data();
                let expected = expected_set.contains(before.name());
                if changed != expected {
                    return Err(SlideTableAppearanceError::Verification);
                }
            },
            (None, None) => break,
            _ => return Err(SlideTableAppearanceError::Verification),
        }
    }
    Ok(())
}

#[allow(
    clippy::type_complexity,
    reason = "the private selection tuple mirrors native graph routes"
)]
fn select_table_with_budget(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    table_selector: TableSelector,
    budget: &mut AppearanceBudget,
) -> Result<AppearanceSelection, SlideTableAppearanceError> {
    physical_source(package)?;
    budget.items(1)?;
    let slide_position = resolve_slide_position(package, slide_selector, budget)?;
    let record = package
        .slide_record_at(slide_position.get())
        .map_err(map_read_error)?
        .ok_or(SlideTableAppearanceError::SlidePositionNotFound {
            position: slide_position,
        })?;
    let (slide_component, slide) = package
        .object_with_component(record.slide_identifier)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let (slide_message_index, slide_payload) = unique_message(slide, SLIDE_MESSAGE_TYPE)?;
    let limits = budget.residual_wire(package)?;
    let owned = repeated_references(slide_payload, SLIDE_OWNED_DRAWABLES_FIELD, limits, budget)?;
    let z_order = repeated_references(slide_payload, SLIDE_Z_ORDER_FIELD, limits, budget)?;
    if owned.len() > package.semantic_limits().max_references()
        || z_order.len() > package.semantic_limits().max_references()
    {
        return Err(SlideTableAppearanceError::LimitExceeded {
            kind: SlideTableAppearanceLimitKind::References,
            observed: owned.len().saturating_add(z_order.len()) as u64,
            maximum: package.semantic_limits().max_references() as u64,
        });
    }
    let owned_set = reject_duplicates(&owned, budget)?;
    let _z_order_set = reject_duplicates(&z_order, budget)?;
    validate_slide_metadata(
        package,
        slide,
        slide_message_index,
        &owned,
        &z_order,
        budget,
    )?;
    ensure_unique_table_owner(
        package,
        record.slide_identifier,
        &owned,
        &z_order,
        limits,
        budget,
    )?;

    budget.allocations(1)?;
    budget.scratch(
        z_order
            .len()
            .checked_mul(size_of::<(
                u64,
                u64,
                Arc<str>,
                Arc<str>,
                Arc<str>,
                usize,
                usize,
                u64,
                Option<u64>,
                Option<u64>,
                Option<u64>,
                [u64; MAX_STYLE_INHERITANCE_DEPTH],
                usize,
                Appearance,
                bool,
            )>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut tables = Vec::new();
    tables
        .try_reserve_exact(z_order.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: z_order.len(),
        })?;
    for table_info_identifier in z_order.iter().copied() {
        let Some((table_info_component, table_info_object)) =
            package.object_with_component(table_info_identifier)
        else {
            return Err(SlideTableAppearanceError::InvalidSource);
        };
        if !table_info_object
            .messages
            .iter()
            .any(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        {
            if table_info_object
                .messages
                .iter()
                .any(|message| is_appearance_role_message(message.type_))
            {
                return Err(SlideTableAppearanceError::UnsupportedTopology);
            }
            continue;
        }
        if !owned_set.contains(&table_info_identifier) {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        validate_role(
            table_info_object,
            TABLE_INFO_MESSAGE_TYPE,
            &[
                TABLE_MODEL_MESSAGE_TYPE,
                TABLE_STYLE_MESSAGE_TYPE,
                TABLE_STYLE_PRESET_MESSAGE_TYPE,
                TABLE_STYLE_NETWORK_MESSAGE_TYPE,
                TABLE_STYLESHEET_MESSAGE_TYPE,
            ],
        )?;
        let (table_info_message_index, info_payload) =
            unique_message(table_info_object, TABLE_INFO_MESSAGE_TYPE)?;
        let info = decode_table_info(info_payload, limits)?;
        let parent = table_parent(info_payload, limits)?;
        if parent != record.slide_identifier {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        let model_identifier = info.table_model().identifier().get();
        validate_table_info_metadata(
            package,
            table_info_object,
            table_info_message_index,
            model_identifier,
            budget,
        )?;
        ensure_unique_identity(package, table_info_identifier, budget)?;
        ensure_unique_identity(package, model_identifier, budget)?;
        let Some((model_component, model_object)) = package.object_with_component(model_identifier)
        else {
            return Err(SlideTableAppearanceError::InvalidSource);
        };
        // Type 6000 is the historical TableInfo/model role and is not a
        // canonical model when it appears on the model object itself.
        validate_role(
            model_object,
            TABLE_MODEL_MESSAGE_TYPE,
            &[
                TABLE_INFO_MESSAGE_TYPE,
                TABLE_STYLE_MESSAGE_TYPE,
                TABLE_STYLE_PRESET_MESSAGE_TYPE,
                TABLE_STYLE_NETWORK_MESSAGE_TYPE,
                TABLE_STYLESHEET_MESSAGE_TYPE,
            ],
        )?;
        let (model_message_index, model_payload) =
            unique_message(model_object, TABLE_MODEL_MESSAGE_TYPE)?;
        let model_options = budget.codec_options(package, model_payload)?;
        let (model, model_report) =
            appearance_codec::decode_table_model_with_report(model_payload, model_options)
                .map_err(map_appearance_codec_error)?;
        budget.codec_report(model_report)?;
        let style_identifier = model.style_identifier();
        let style_preset_identifier = model.style_preset_identifier();
        validate_model_metadata(
            package,
            model_object,
            model_message_index,
            style_identifier,
            style_preset_identifier,
            budget,
        )?;
        let resolved = resolve_effective_appearance(
            package,
            style_identifier,
            style_preset_identifier,
            limits,
            budget,
        )?;
        if let Some(stylesheet_identifier) = resolved.stylesheet_identifier {
            validate_stylesheet_archive_metadata(package, stylesheet_identifier, limits, budget)?;
        }
        validate_global_inbound_references(
            package,
            record.slide_identifier,
            slide_message_index,
            table_info_identifier,
            table_info_message_index,
            model_identifier,
            model_message_index,
            &resolved,
            budget,
        )?;
        tables.push((
            table_info_identifier,
            model_identifier,
            Arc::<str>::from(slide_component),
            Arc::<str>::from(table_info_component),
            Arc::<str>::from(model_component),
            table_info_message_index,
            model_message_index,
            style_identifier,
            style_preset_identifier,
            resolved.stylesheet_identifier,
            resolved.first_style_identifier,
            resolved.style_ids,
            resolved.style_count,
            resolved.appearance,
            info.locked().unwrap_or(false),
        ));
    }

    let table_position = table_selector.as_position();
    let (
        table_info_identifier,
        model_identifier,
        _slide_component,
        table_info_component,
        model_component,
        table_info_message_index,
        model_message_index,
        model_style_identifier,
        style_preset_identifier,
        stylesheet_identifier,
        effective_style_identifier,
        style_ids,
        style_count,
        before,
        locked,
    ) = tables.get(table_position.get()).cloned().ok_or(
        SlideTableAppearanceError::TablePositionNotFound {
            position: table_position,
        },
    )?;
    let metadata_component = find_metadata_component(package, budget)?;
    Ok(AppearanceSelection {
        slide_position,
        table_position,
        slide_identifier: record.slide_identifier,
        table_info_identifier,
        model_identifier,
        slide_message_index,
        table_info_message_index,
        model_message_index,
        model_component,
        table_info_component,
        style_component: effective_style_identifier.and_then(|identifier| {
            package
                .object_with_component(identifier)
                .map(|(component, _)| Arc::from(component))
        }),
        metadata_component,
        model_style_identifier,
        style_identifier: effective_style_identifier.unwrap_or(model_style_identifier),
        style_ids,
        style_count,
        style_preset_identifier,
        stylesheet_identifier,
        before,
        locked,
    })
}

fn physical_source(
    package: &Package,
) -> Result<&litchi_iwa_archive::SourceCatalog, SlideTableAppearanceError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(SlideTableAppearanceError::UnsupportedSource),
    }
}

fn resolve_slide_position(
    package: &Package,
    selector: SlideSelector<'_>,
    budget: &mut AppearanceBudget,
) -> Result<Position, SlideTableAppearanceError> {
    match selector {
        SlideSelector::Position(position) => Ok(position),
        SlideSelector::Name(name) => {
            if name.is_empty() {
                return Err(SlideTableAppearanceError::EmptySlideName);
            }
            let limits = budget.residual_wire(package)?;
            let mut selected = None;
            for index in 0..=package.semantic_limits().max_slides() {
                budget.items(1)?;
                let Some(record) = package.slide_record_at(index).map_err(map_read_error)? else {
                    break;
                };
                let (_, slide) = package
                    .object_with_component(record.slide_identifier)
                    .ok_or(SlideTableAppearanceError::InvalidSource)?;
                let (_, payload) = unique_message(slide, SLIDE_MESSAGE_TYPE)?;
                budget.work(payload.len())?;
                budget.wire_bytes(payload.len())?;
                if slide_name(payload, limits)? == Some(name) {
                    if selected.replace(Position::new(index)).is_some() {
                        return Err(SlideTableAppearanceError::AmbiguousSelector);
                    }
                }
            }
            selected.ok_or(SlideTableAppearanceError::SlideNameNotFound)
        },
    }
}

fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<(usize, &[u8]), SlideTableAppearanceError> {
    if object.messages.len() != object.archive_info.message_infos.len() {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        validate_message_header(object, index)?;
        if message.type_ == message_type
            && selected.replace((index, message.data.as_slice())).is_some()
        {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
    }
    selected.ok_or(SlideTableAppearanceError::InvalidSource)
}

fn validate_message_header(
    object: &ArchiveObject,
    index: usize,
) -> Result<(), SlideTableAppearanceError> {
    let message = object
        .messages
        .get(index)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    if message.type_ != info.type_
        || usize::try_from(info.length).ok() != Some(message.data.len())
        || object.archive_info.should_merge == Some(true)
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    Ok(())
}

fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut AppearanceBudget,
) -> Result<Vec<u64>, SlideTableAppearanceError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let field_count = fields.fields().count();
    budget.fields(field_count)?;
    budget.work(payload.len())?;
    budget.wire_bytes(payload.len())?;
    let count = fields
        .fields()
        .filter(|field| field.number() == field_number)
        .count();
    budget.allocations(1)?;
    budget.scratch(count.saturating_mul(size_of::<u64>()))?;
    let mut values = Vec::new();
    values
        .try_reserve_exact(count)
        .map_err(|_| SlideTableAppearanceError::Allocation { amount: count })?;
    budget.references(count)?;
    for field in fields
        .fields()
        .filter(|field| field.number() == field_number)
    {
        field.validate_canonical_key().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        values.push(strict_reference(field.payload(), limits)?);
    }
    Ok(values)
}

fn slide_name(
    payload: &[u8],
    limits: WireLimits,
) -> Result<Option<&str>, SlideTableAppearanceError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut name = None;
    for field in fields.fields().filter(|field| field.number() == 10) {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.wire_type() != 2 || name.replace(field.payload()).is_some() {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
    }
    name.map(str::from_utf8)
        .transpose()
        .map_err(|_| SlideTableAppearanceError::InvalidSource)
}

fn table_parent(payload: &[u8], limits: WireLimits) -> Result<u64, SlideTableAppearanceError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut super_payload = None;
    for field in fields
        .fields()
        .filter(|field| field.number() == TABLE_SUPER_FIELD)
    {
        field.validate_canonical_key().map_err(map_wire_error)?;
        if field.wire_type() != 2 || super_payload.is_some() {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        super_payload = Some(field.payload());
    }
    let super_payload = super_payload.ok_or(SlideTableAppearanceError::InvalidSource)?;
    let drawable = WireView::parse_with_limits(super_payload, limits).map_err(map_wire_error)?;
    let mut parent_payload = None;
    for field in drawable
        .fields()
        .filter(|field| field.number() == DRAWABLE_PARENT_FIELD)
    {
        field.validate_canonical_key().map_err(map_wire_error)?;
        if field.wire_type() != 2 || parent_payload.is_some() {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        parent_payload = Some(field.payload());
    }
    strict_reference(
        parent_payload.ok_or(SlideTableAppearanceError::InvalidSource)?,
        limits,
    )
}

fn strict_reference(payload: &[u8], limits: WireLimits) -> Result<u64, SlideTableAppearanceError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    let mut identifier = None;
    for field in fields.fields() {
        field.validate_canonical_key().map_err(map_wire_error)?;
        if field.number() != 1 || field.wire_type() != 0 || identifier.is_some() {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        let (value, width) = decode_varint_from_bytes(field.payload())
            .map_err(|_| SlideTableAppearanceError::InvalidSource)?;
        if value == 0 || width != encoded_len(value) {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        identifier = Some(value);
    }
    identifier.ok_or(SlideTableAppearanceError::InvalidSource)
}

fn reject_duplicates(
    values: &[u64],
    budget: &mut AppearanceBudget,
) -> Result<HashSet<u64>, SlideTableAppearanceError> {
    budget.allocations(1)?;
    budget.scratch(
        values
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut seen = HashSet::new();
    seen.try_reserve(values.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: values.len(),
        })?;
    for value in values {
        if !seen.insert(*value) {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
    }
    Ok(seen)
}

/// Build a bounded physical-member index for a transaction pass.
///
/// The caller supplies the duplicate-name error because malformed source
/// records are an invalid source during rewrite, while an ambiguous name in a
/// reopened candidate is a verification failure.
fn entry_index<'a>(
    catalog: &'a Catalog,
    budget: &mut AppearanceBudget,
    duplicate_error: SlideTableAppearanceError,
) -> Result<HashMap<&'a str, &'a Entry>, SlideTableAppearanceError> {
    let count = catalog.iter().count();
    budget.allocations(usize::from(count != 0))?;
    budget.retained(
        count
            .checked_mul(size_of::<(&str, &Entry)>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut index = HashMap::new();
    index
        .try_reserve(count)
        .map_err(|_| SlideTableAppearanceError::Allocation { amount: count })?;
    for entry in catalog.iter() {
        budget.transaction_work(
            entry
                .name()
                .len()
                .checked_add(1)
                .ok_or(SlideTableAppearanceError::InvalidSource)?,
        )?;
        if index.insert(entry.name(), entry).is_some() {
            return Err(duplicate_error);
        }
    }
    Ok(index)
}

/// Count reference occurrences in one aggregate using one fallibly reserved
/// frequency map instead of a scan for each queried identifier.
fn reference_frequencies(
    values: &[u64],
    budget: &mut AppearanceBudget,
) -> Result<HashMap<u64, usize>, SlideTableAppearanceError> {
    budget.transaction_work(values.len())?;
    budget.allocations(usize::from(!values.is_empty()))?;
    budget.retained(
        values
            .len()
            .checked_mul(size_of::<(u64, usize)>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut frequencies = HashMap::new();
    frequencies
        .try_reserve(values.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: values.len(),
        })?;
    for value in values {
        let count = frequencies.entry(*value).or_insert(0usize);
        *count = count
            .checked_add(1)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
    }
    Ok(frequencies)
}

fn is_appearance_role_message(message_type: u32) -> bool {
    matches!(
        message_type,
        TABLE_INFO_MESSAGE_TYPE
            | TABLE_MODEL_MESSAGE_TYPE
            | TABLE_STYLE_MESSAGE_TYPE
            | TABLE_STYLE_PRESET_MESSAGE_TYPE
            | TABLE_STYLE_NETWORK_MESSAGE_TYPE
            | TABLE_STYLESHEET_MESSAGE_TYPE
    )
}

fn validate_role(
    object: &ArchiveObject,
    expected_type: u32,
    aliases: &[u32],
) -> Result<(), SlideTableAppearanceError> {
    let count = object
        .messages
        .iter()
        .filter(|message| message.type_ == expected_type)
        .count();
    if count != 1
        || aliases.iter().any(|alias| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == *alias)
        })
    {
        return Err(SlideTableAppearanceError::UnsupportedTopology);
    }
    Ok(())
}

fn validate_slide_metadata(
    package: &Package,
    object: &ArchiveObject,
    index: usize,
    owned: &[u64],
    z_order: &[u64],
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    validate_message_header(object, index)?;
    let info = object
        .archive_info
        .message_infos
        .get(index)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    validate_resolvable_aggregate_references(package, info, budget)?;
    let expected_capacity = owned
        .len()
        .checked_add(z_order.len())
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    budget.allocations(2)?;
    budget.scratch(
        expected_capacity
            .checked_mul(size_of::<u64>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut expected = Vec::new();
    expected.try_reserve_exact(expected_capacity).map_err(|_| {
        SlideTableAppearanceError::Allocation {
            amount: expected_capacity,
        }
    })?;
    let mut expected_set = HashSet::new();
    expected_set.try_reserve(expected_capacity).map_err(|_| {
        SlideTableAppearanceError::Allocation {
            amount: expected_capacity,
        }
    })?;
    for identifier in owned.iter().chain(z_order) {
        if expected_set.insert(*identifier) {
            expected.push(*identifier);
        }
    }
    let frequencies = reference_frequencies(&info.object_references, budget)?;
    if expected
        .iter()
        .any(|identifier| frequencies.get(identifier).copied() != Some(1))
    {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let mut owned_fields = 0usize;
    let mut z_order_fields = 0usize;
    for field in &info.field_infos {
        match field.path.as_slice() {
            [SLIDE_OWNED_DRAWABLES_FIELD] => {
                owned_fields += 1;
                if field
                    .r#type
                    .is_some_and(|kind| kind != FieldType::ObjectReference)
                {
                    return Err(SlideTableAppearanceError::InvalidSource);
                }
                if field.object_references.as_slice() != owned {
                    return Err(SlideTableAppearanceError::InvalidSource);
                }
            },
            [SLIDE_Z_ORDER_FIELD] => {
                z_order_fields += 1;
                if field
                    .r#type
                    .is_some_and(|kind| kind != FieldType::ObjectReference)
                {
                    return Err(SlideTableAppearanceError::InvalidSource);
                }
                if field.object_references.as_slice() != z_order {
                    return Err(SlideTableAppearanceError::InvalidSource);
                }
            },
            _ if field
                .object_references
                .iter()
                .any(|identifier| expected_set.contains(identifier)) =>
            {
                return Err(SlideTableAppearanceError::InvalidSource);
            },
            _ => {},
        }
    }
    // Keynote producers may place unrelated object references in the message
    // aggregate and omit the repeated route FieldInfos.  The selected route
    // must still occur exactly once, and any emitted route FieldInfos must be
    // complete and exact.
    if owned_fields > 1 || z_order_fields > 1 {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    Ok(())
}

fn validate_table_info_metadata(
    package: &Package,
    object: &ArchiveObject,
    index: usize,
    model_identifier: u64,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    validate_message_header(object, index)?;
    let info = &object.archive_info.message_infos[index];
    validate_resolvable_aggregate_references(package, info, budget)?;
    let frequencies = reference_frequencies(&info.object_references, budget)?;
    if frequencies.get(&model_identifier).copied() != Some(1) {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let mut model_fields = 0usize;
    for field in &info.field_infos {
        if field.path.as_slice() == [TABLE_MODEL_FIELD] {
            model_fields += 1;
            if field
                .r#type
                .is_some_and(|kind| kind != FieldType::ObjectReference)
                || field.object_references.as_slice() != [model_identifier]
            {
                return Err(SlideTableAppearanceError::InvalidSource);
            }
        } else if field.object_references.contains(&model_identifier) {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
    }
    if model_fields > 1 {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    Ok(())
}

fn validate_model_metadata(
    package: &Package,
    object: &ArchiveObject,
    index: usize,
    style_identifier: u64,
    style_preset_identifier: Option<u64>,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    validate_message_header(object, index)?;
    let info = &object.archive_info.message_infos[index];
    validate_resolvable_aggregate_references(package, info, budget)?;
    let expected = if style_identifier != 0 {
        Some((style_identifier, 3))
    } else {
        style_preset_identifier.map(|identifier| (identifier, 48))
    };
    if let Some((identifier, path)) = expected {
        let frequencies = reference_frequencies(&info.object_references, budget)?;
        if frequencies.get(&identifier).copied() != Some(1) {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        let mut selected_fields = 0usize;
        for field in &info.field_infos {
            if field.path.as_slice() == [path] {
                selected_fields += 1;
                if field
                    .r#type
                    .is_some_and(|kind| kind != FieldType::ObjectReference)
                    || field.object_references.as_slice() != [identifier]
                {
                    return Err(SlideTableAppearanceError::InvalidSource);
                }
            } else if field.object_references.contains(&identifier) {
                return Err(SlideTableAppearanceError::InvalidSource);
            }
        }
        if selected_fields > 1 {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
    }
    Ok(())
}

fn validate_style_metadata(
    package: &Package,
    object: &ArchiveObject,
    index: usize,
    parent_identifier: Option<u64>,
    stylesheet_identifier: u64,
    allow_aggregate_omission: bool,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    validate_message_header(object, index)?;
    let info = &object.archive_info.message_infos[index];
    validate_resolvable_aggregate_references(package, info, budget)?;
    let mut expected = Vec::with_capacity(2);
    if let Some(parent) = parent_identifier {
        expected.push(parent);
    }
    expected.push(stylesheet_identifier);
    let aggregate_omitted = info.object_references.is_empty();
    if aggregate_omitted && !allow_aggregate_omission {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
    }
    if !aggregate_omitted && info.object_references.as_slice() != expected.as_slice() {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let mut parent_field_count = 0usize;
    let mut stylesheet_field_count = 0usize;
    for field in &info.field_infos {
        if field.path.as_slice() == [1, 3] {
            parent_field_count += 1;
            if parent_identifier != Some(field.object_references.first().copied().unwrap_or(0))
                || field.object_references.len() != 1
                || !field.data_references.is_empty()
                || field
                    .r#type
                    .is_some_and(|kind| kind != FieldType::ObjectReference)
            {
                return Err(SlideTableAppearanceError::InvalidSource);
            }
        } else if field.path.as_slice() == [1, 5] {
            stylesheet_field_count += 1;
            if field.object_references.as_slice() != [stylesheet_identifier]
                || !field.data_references.is_empty()
                || field
                    .r#type
                    .is_some_and(|kind| kind != FieldType::ObjectReference)
            {
                return Err(SlideTableAppearanceError::InvalidSource);
            }
        } else {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
    }
    // Some native producers omit style aggregate metadata entirely.  The
    // payload route, stylesheet registry, physical identity, UUID ownership,
    // and external-edge census remain authoritative.  Otherwise the complete
    // aggregate/FieldInfo shape is narrow enough to copy to the fresh COW
    // style without inventing or dropping producer metadata.
    if aggregate_omitted && (parent_field_count != 0 || stylesheet_field_count != 0) {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    if (parent_field_count != 0 || stylesheet_field_count != 0)
        && (parent_field_count != usize::from(parent_identifier.is_some())
            || stylesheet_field_count != 1)
    {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    Ok(())
}

fn validate_resolvable_aggregate_references(
    package: &Package,
    info: &litchi_iwa_core::MessageInfo,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    let max_field_references = info
        .field_infos
        .iter()
        .map(|field| {
            field
                .object_references
                .len()
                .max(field.data_references.len())
        })
        .max()
        .unwrap_or(0);
    budget.allocations(3)?;
    budget.scratch(
        info.object_references
            .len()
            .checked_add(info.data_references.len())
            .and_then(|count| count.checked_add(max_field_references))
            .and_then(|count| count.checked_mul(size_of::<u64>()))
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut object_references = HashSet::new();
    object_references
        .try_reserve(info.object_references.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: info.object_references.len(),
        })?;
    // Native producers can repeat unrelated references in aggregate and
    // FieldInfo lists. Resolve each distinct target exactly once here; the
    // role-specific validators above still require every selected route to
    // occur exactly once and reject it on every unrelated field path.
    for identifier in &info.object_references {
        if *identifier == 0 {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        if object_references.insert(*identifier) {
            ensure_unique_identity(package, *identifier, budget)?;
        }
    }
    let mut data_references = HashSet::new();
    data_references
        .try_reserve(info.data_references.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: info.data_references.len(),
        })?;
    for identifier in &info.data_references {
        if *identifier == 0 {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        data_references.insert(*identifier);
    }
    let mut field_references = HashSet::new();
    field_references
        .try_reserve(max_field_references)
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: max_field_references,
        })?;
    for field in &info.field_infos {
        field_references.clear();
        for identifier in &field.object_references {
            if *identifier == 0 {
                return Err(SlideTableAppearanceError::InvalidSource);
            }
            if field_references.insert(*identifier) {
                ensure_unique_identity(package, *identifier, budget)?;
            }
        }
        field_references.clear();
        for identifier in &field.data_references {
            if *identifier == 0 {
                return Err(SlideTableAppearanceError::InvalidSource);
            }
            field_references.insert(*identifier);
        }
    }
    Ok(())
}

/// Validate that every effective style in the inheritance chain is actually
/// registered by the selected stylesheet.  The appearance codec validates
/// the registry wire and duplicate entries; this owner additionally binds
/// the registry's semantic membership to the exact style chain it will copy.
fn validate_stylesheet_registry(
    package: &Package,
    payload: &[u8],
    style_ids: &[u64],
    limits: WireLimits,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    let (snapshot, report) = appearance_codec::decode_stylesheet_with_report(
        payload,
        budget.codec_options(package, payload)?,
    )
    .map_err(map_appearance_codec_error)?;
    budget.codec_report(report)?;
    budget.styles(snapshot.style_count())?;
    if snapshot.style_count() == 0 || snapshot.style_count() < style_ids.len() {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
    }
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.allocations(1)?;
    budget.scratch(
        snapshot
            .style_count()
            .checked_mul(size_of::<u64>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut registered = HashSet::new();
    registered
        .try_reserve(snapshot.style_count())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: snapshot.style_count(),
        })?;
    for field in fields.fields().filter(|field| field.number() == 1) {
        field.validate_canonical_framing().map_err(map_wire_error)?;
        if field.wire_type() != 2 {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        let identifier = strict_reference(field.payload(), limits)?;
        if !registered.insert(identifier) {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
    }
    if registered.len() != snapshot.style_count()
        || style_ids
            .iter()
            .any(|identifier| !registered.contains(identifier))
    {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
    }
    Ok(())
}

fn validate_stylesheet_archive_metadata(
    package: &Package,
    identifier: u64,
    _limits: WireLimits,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    let (_, object) = package
        .object_with_component(identifier)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let (message_index, _payload) = unique_message(object, TABLE_STYLESHEET_MESSAGE_TYPE)?;
    validate_message_header(object, message_index)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    validate_resolvable_aggregate_references(package, info, budget)?;
    let mut registry_fields = 0usize;
    for field in &info.field_infos {
        if field.path.as_slice() == [1] {
            registry_fields += 1;
            if field
                .r#type
                .is_some_and(|kind| kind != FieldType::ObjectReference)
                || field.object_references.as_slice() != info.object_references.as_slice()
            {
                return Err(SlideTableAppearanceError::InvalidSource);
            }
        }
    }
    if registry_fields > 1 {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    Ok(())
}

#[derive(Clone, Copy)]
struct ResolvedAppearance {
    appearance: Appearance,
    style_ids: [u64; MAX_STYLE_INHERITANCE_DEPTH],
    style_count: usize,
    preset_identifier: Option<u64>,
    network_identifier: Option<u64>,
    stylesheet_identifier: Option<u64>,
    first_style_identifier: Option<u64>,
}

fn resolve_effective_appearance(
    package: &Package,
    style_identifier: u64,
    style_preset_identifier: Option<u64>,
    limits: WireLimits,
    budget: &mut AppearanceBudget,
) -> Result<ResolvedAppearance, SlideTableAppearanceError> {
    let mut preset_identifier = None;
    let mut network_identifier = None;
    let first_style_identifier = if style_identifier != 0 {
        Some(style_identifier)
    } else if let Some(preset) = style_preset_identifier {
        ensure_unique_identity(package, preset, budget)?;
        let (_, preset_object) = package
            .object_with_component(preset)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        validate_role(
            preset_object,
            TABLE_STYLE_PRESET_MESSAGE_TYPE,
            &[
                TABLE_STYLE_MESSAGE_TYPE,
                TABLE_STYLE_NETWORK_MESSAGE_TYPE,
                TABLE_MODEL_MESSAGE_TYPE,
                TABLE_INFO_MESSAGE_TYPE,
                TABLE_STYLESHEET_MESSAGE_TYPE,
            ],
        )?;
        let (preset_index, preset_payload) =
            unique_message(preset_object, TABLE_STYLE_PRESET_MESSAGE_TYPE)?;
        let (preset_snapshot, preset_report) =
            appearance_codec::decode_table_style_preset_with_report(
                preset_payload,
                budget.codec_options(package, preset_payload)?,
            )
            .map_err(map_appearance_codec_error)?;
        budget.codec_report(preset_report)?;
        let network = preset_snapshot
            .style_network_identifier()
            .filter(|value| *value != 0)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        validate_reference_metadata(preset_object, preset_index, &[network], &[3])?;
        ensure_unique_identity(package, network, budget)?;
        let (_, network_object) = package
            .object_with_component(network)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        validate_role(
            network_object,
            TABLE_STYLE_NETWORK_MESSAGE_TYPE,
            &[
                TABLE_STYLE_MESSAGE_TYPE,
                TABLE_STYLE_PRESET_MESSAGE_TYPE,
                TABLE_MODEL_MESSAGE_TYPE,
                TABLE_INFO_MESSAGE_TYPE,
                TABLE_STYLESHEET_MESSAGE_TYPE,
            ],
        )?;
        let (network_index, network_payload) =
            unique_message(network_object, TABLE_STYLE_NETWORK_MESSAGE_TYPE)?;
        let (network_snapshot, network_report) =
            appearance_codec::decode_table_style_network_with_report(
                network_payload,
                budget.codec_options(package, network_payload)?,
            )
            .map_err(map_appearance_codec_error)?;
        budget.codec_report(network_report)?;
        let network_references = network_reference_targets(network_payload, limits, budget)?;
        let style = network_snapshot.table_style_identifier();
        if style == 0 {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        validate_reference_metadata(network_object, network_index, &network_references, &[9])?;
        preset_identifier = Some(preset);
        network_identifier = Some(network);
        Some(style)
    } else {
        None
    };

    let Some(first_style_identifier) = first_style_identifier else {
        return Ok(ResolvedAppearance {
            appearance: Appearance::default(),
            style_ids: [0; MAX_STYLE_INHERITANCE_DEPTH],
            style_count: 0,
            preset_identifier,
            network_identifier,
            stylesheet_identifier: None,
            first_style_identifier: None,
        });
    };
    let mut style_ids = [0u64; MAX_STYLE_INHERITANCE_DEPTH];
    let mut style_count = 0usize;
    let mut current = Some(first_style_identifier);
    let mut row_banding = None;
    let mut row_sizing = None;
    let mut body_horizontal = None;
    let mut body_vertical = None;
    let mut header_columns_horizontal = None;
    let mut header_rows_vertical = None;
    let mut footer_rows_vertical = None;
    let mut stylesheet_identifier = None;
    let mut style_component_name = None;
    while let Some(identifier) = current {
        if style_ids[..style_count].contains(&identifier) {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        if style_count == style_ids.len() {
            return Err(SlideTableAppearanceError::LimitExceeded {
                kind: SlideTableAppearanceLimitKind::WireNesting,
                observed: (style_count + 1) as u64,
                maximum: MAX_STYLE_INHERITANCE_DEPTH as u64,
            });
        }
        ensure_unique_identity(package, identifier, budget)?;
        let (style_component, style_object) = package
            .object_with_component(identifier)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        if style_component_name.is_some_and(|name| name != style_component) {
            return Err(SlideTableAppearanceError::UnsupportedDependency);
        }
        style_component_name = Some(style_component);
        validate_role(
            style_object,
            TABLE_STYLE_MESSAGE_TYPE,
            &[
                TABLE_STYLE_PRESET_MESSAGE_TYPE,
                TABLE_STYLE_NETWORK_MESSAGE_TYPE,
                TABLE_MODEL_MESSAGE_TYPE,
                TABLE_INFO_MESSAGE_TYPE,
                TABLE_STYLESHEET_MESSAGE_TYPE,
            ],
        )?;
        let (style_index, style_payload) = unique_message(style_object, TABLE_STYLE_MESSAGE_TYPE)?;
        let (style, style_report) = appearance_codec::decode_table_style_with_report(
            style_payload,
            budget.codec_options(package, style_payload)?,
        )
        .map_err(map_appearance_codec_error)?;
        budget.codec_report(style_report)?;
        if style.is_variation() != style.parent_identifier().is_some() {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        let sheet = style
            .stylesheet_identifier()
            .filter(|value| *value != 0)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        if stylesheet_identifier.is_some_and(|previous| previous != sheet) {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        stylesheet_identifier = Some(sheet);
        ensure_unique_identity(package, sheet, budget)?;
        let (stylesheet_component, stylesheet_object) = package
            .object_with_component(sheet)
            .ok_or(SlideTableAppearanceError::InvalidSource)?;
        validate_role(
            stylesheet_object,
            TABLE_STYLESHEET_MESSAGE_TYPE,
            &[
                TABLE_STYLE_MESSAGE_TYPE,
                TABLE_STYLE_PRESET_MESSAGE_TYPE,
                TABLE_STYLE_NETWORK_MESSAGE_TYPE,
                TABLE_MODEL_MESSAGE_TYPE,
                TABLE_INFO_MESSAGE_TYPE,
            ],
        )?;
        validate_style_metadata(
            package,
            style_object,
            style_index,
            style.parent_identifier(),
            sheet,
            style_component == stylesheet_component,
            budget,
        )?;
        let overrides = style.overrides();
        row_banding = row_banding.or(overrides.row_banding);
        row_sizing = row_sizing.or(overrides.row_sizing);
        body_horizontal = body_horizontal.or(overrides.body_horizontal);
        body_vertical = body_vertical.or(overrides.body_vertical);
        header_columns_horizontal =
            header_columns_horizontal.or(overrides.header_columns_horizontal);
        header_rows_vertical = header_rows_vertical.or(overrides.header_rows_vertical);
        footer_rows_vertical = footer_rows_vertical.or(overrides.footer_rows_vertical);
        style_ids[style_count] = identifier;
        style_count += 1;
        current = style.parent_identifier().filter(|value| *value != 0);
    }
    let stylesheet_identifier =
        stylesheet_identifier.ok_or(SlideTableAppearanceError::InvalidSource)?;
    let stylesheet_payload = stylesheet_message_data(
        package,
        stylesheet_identifier,
        SlideTableAppearancePath::Package,
    )?;
    validate_stylesheet_registry(
        package,
        stylesheet_payload,
        &style_ids[..style_count],
        limits,
        budget,
    )?;
    Ok(ResolvedAppearance {
        appearance: appearance_from_overrides(
            row_banding,
            row_sizing,
            body_horizontal,
            body_vertical,
            header_columns_horizontal,
            header_rows_vertical,
            footer_rows_vertical,
        ),
        style_ids,
        style_count,
        preset_identifier,
        network_identifier,
        stylesheet_identifier: Some(stylesheet_identifier),
        first_style_identifier: Some(first_style_identifier),
    })
}

fn appearance_from_overrides(
    row_banding: Option<bool>,
    row_sizing: Option<bool>,
    body_horizontal: Option<bool>,
    body_vertical: Option<bool>,
    header_columns_horizontal: Option<bool>,
    header_rows_vertical: Option<bool>,
    footer_rows_vertical: Option<bool>,
) -> Appearance {
    Appearance {
        row_banding: if row_banding.unwrap_or(false) {
            Banding::Enabled
        } else {
            Banding::Disabled
        },
        row_sizing: if row_sizing.unwrap_or(false) {
            RowSizing::FitCellContents
        } else {
            RowSizing::Fixed
        },
        gridlines: Gridlines {
            body_horizontal: visibility(body_horizontal.unwrap_or(true)),
            header_columns_horizontal: visibility(header_columns_horizontal.unwrap_or(true)),
            body_vertical: visibility(body_vertical.unwrap_or(true)),
            header_rows_vertical: visibility(header_rows_vertical.unwrap_or(true)),
            footer_rows_vertical: visibility(footer_rows_vertical.unwrap_or(true)),
        },
    }
}

const fn visibility(value: bool) -> GridlineVisibility {
    if value {
        GridlineVisibility::Visible
    } else {
        GridlineVisibility::Hidden
    }
}

fn validate_reference_metadata(
    object: &ArchiveObject,
    index: usize,
    expected: &[u64],
    expected_path: &[u32],
) -> Result<(), SlideTableAppearanceError> {
    validate_message_header(object, index)?;
    let info = &object.archive_info.message_infos[index];
    if info.object_references.as_slice() != expected {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let mut selected_fields = 0usize;
    for field in &info.field_infos {
        if field.path.as_slice() == expected_path {
            selected_fields = selected_fields.saturating_add(1);
            if field
                .r#type
                .is_some_and(|kind| kind != FieldType::ObjectReference)
                || field.object_references.as_slice() != expected
            {
                return Err(SlideTableAppearanceError::InvalidSource);
            }
        } else if !field.object_references.is_empty() {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
    }
    if selected_fields > 1 {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    Ok(())
}

fn network_reference_targets(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut AppearanceBudget,
) -> Result<Vec<u64>, SlideTableAppearanceError> {
    budget.allocations(1)?;
    budget.scratch(
        9usize
            .checked_mul(size_of::<u64>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut references = Vec::new();
    references
        .try_reserve_exact(9)
        .map_err(|_| SlideTableAppearanceError::Allocation { amount: 9 })?;
    for field_number in 1..=9 {
        let values = repeated_references(payload, field_number, limits, budget)?;
        if values.len() != 1 {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        references.push(values[0]);
    }
    Ok(references)
}

fn ensure_unique_identity(
    package: &Package,
    identifier: u64,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    let object_count = package
        .state
        .source
        .components()
        .iter()
        .map(|component| component.archive().objects.len())
        .try_fold(0usize, |total, count| total.checked_add(count))
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    budget.transaction_work(object_count)?;
    let mut count = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            if object.archive_info.identifier == Some(identifier) {
                count += 1;
            }
        }
    }
    if count == 1 {
        Ok(())
    } else {
        Err(SlideTableAppearanceError::UnsupportedDependency)
    }
}

fn ensure_unique_table_owner(
    package: &Package,
    slide_identifier: u64,
    owned: &[u64],
    z_order: &[u64],
    limits: WireLimits,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    let object_count = package
        .state
        .source
        .components()
        .iter()
        .map(|component| component.archive().objects.len())
        .try_fold(0usize, |total, count| total.checked_add(count))
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let scan_work = object_count
        .checked_add(z_order.len())
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    budget.transaction_work(scan_work)?;
    budget.allocations(3)?;
    budget.scratch(
        z_order
            .len()
            .checked_mul(size_of::<u64>())
            .and_then(|size| {
                z_order
                    .len()
                    .checked_mul(size_of::<usize>())
                    .and_then(|extra| size.checked_add(extra))
            })
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut z_order_set = HashSet::new();
    z_order_set
        .try_reserve(z_order.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: z_order.len(),
        })?;
    z_order_set.extend(z_order.iter().copied());
    let selected_table_count = owned.iter().fold(0usize, |count, identifier| {
        count + usize::from(z_order_set.contains(identifier))
    });
    if selected_table_count != z_order.len() {
        return Err(SlideTableAppearanceError::InvalidSource);
    }
    let mut slide_owned_count = 0usize;
    let mut slide_z_order_count = 0usize;
    budget.retained(
        z_order
            .len()
            .checked_mul(size_of::<(u64, usize)>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut model_owner_count = HashMap::new();
    model_owner_count.try_reserve(z_order.len()).map_err(|_| {
        SlideTableAppearanceError::Allocation {
            amount: z_order.len(),
        }
    })?;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ != SLIDE_MESSAGE_TYPE {
                    continue;
                }
                let current_owned = repeated_references(
                    &message.data,
                    SLIDE_OWNED_DRAWABLES_FIELD,
                    limits,
                    budget,
                )?;
                let current_z =
                    repeated_references(&message.data, SLIDE_Z_ORDER_FIELD, limits, budget)?;
                if object.archive_info.identifier == Some(slide_identifier) {
                    slide_owned_count = slide_owned_count
                        .checked_add(current_owned.iter().fold(0usize, |count, value| {
                            count + usize::from(z_order_set.contains(value))
                        }))
                        .ok_or(SlideTableAppearanceError::InvalidSource)?;
                    slide_z_order_count = slide_z_order_count
                        .checked_add(current_z.iter().fold(0usize, |count, value| {
                            count + usize::from(z_order_set.contains(value))
                        }))
                        .ok_or(SlideTableAppearanceError::InvalidSource)?;
                }
            }
            if object
                .messages
                .iter()
                .any(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            {
                if let Some(identifier) = object.archive_info.identifier {
                    if z_order_set.contains(&identifier) {
                        let count = model_owner_count.entry(identifier).or_insert(0usize);
                        *count = count
                            .checked_add(1)
                            .ok_or(SlideTableAppearanceError::InvalidSource)?;
                    }
                }
            }
        }
    }
    if slide_owned_count != z_order.len()
        || slide_z_order_count != z_order.len()
        || z_order.iter().any(|identifier| {
            package
                .object_with_component(*identifier)
                .map(|(_, object)| {
                    object
                        .messages
                        .iter()
                        .any(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
                        && model_owner_count.get(identifier).copied().unwrap_or(0) != 1
                })
                .unwrap_or(true)
        })
    {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
    }
    Ok(())
}

fn find_metadata_component(
    package: &Package,
    budget: &mut AppearanceBudget,
) -> Result<Option<Arc<str>>, SlideTableAppearanceError> {
    package
        .state
        .source
        .components()
        .iter()
        .find(|component| {
            component.archive().objects.iter().any(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == METADATA_MESSAGE_TYPE)
            })
        })
        .map(|component| {
            budget.allocations(1)?;
            budget.scratch(component.name().len())?;
            Ok(Arc::from(component.name()))
        })
        .transpose()
}

struct AppearanceInboundVisitor<'a> {
    package: &'a Package,
    data_identifiers: &'a HashSet<u64>,
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
    model_message_index: usize,
    expected_model_edges: usize,
    model_edges: usize,
    style_ids: [u64; MAX_STYLE_INHERITANCE_DEPTH],
    style_count: usize,
    preset_identifier: Option<u64>,
    network_identifier: Option<u64>,
    stylesheet_identifier: Option<u64>,
    references: usize,
    invalid: bool,
}

impl ArchiveReferenceVisitor for AppearanceInboundVisitor<'_> {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        self.references = self.references.saturating_add(1);
        if occurrence.referenced_identifier == 0 {
            self.invalid = true;
            return Ok(());
        }
        if occurrence.kind == ArchiveReferenceKind::Data {
            if !self
                .data_identifiers
                .contains(&occurrence.referenced_identifier)
            {
                self.invalid = true;
            }
            return Ok(());
        }
        if self
            .package
            .object_with_component(occurrence.referenced_identifier)
            .is_none()
        {
            self.invalid = true;
            return Ok(());
        }
        let Some((_, source_object)) = self
            .package
            .object_with_component(occurrence.object_identifier)
        else {
            self.invalid = true;
            return Ok(());
        };
        let Some(source_message_type) = source_object
            .messages
            .get(occurrence.message_index)
            .map(|message| message.type_)
        else {
            self.invalid = true;
            return Ok(());
        };
        let source_is_model = source_message_type == TABLE_MODEL_MESSAGE_TYPE;
        let source_is_style = source_message_type == TABLE_STYLE_MESSAGE_TYPE;
        let source_is_network = source_message_type == TABLE_STYLE_NETWORK_MESSAGE_TYPE;
        let source_is_stylesheet = source_message_type == TABLE_STYLESHEET_MESSAGE_TYPE;
        let source_is_show = source_message_type == SHOW_MESSAGE_TYPE;
        let source_is_show_aggregate =
            source_is_show && matches!(occurrence.scope, ArchiveReferenceScope::Message);
        let target = occurrence.referenced_identifier;
        if target == self.model_identifier {
            if occurrence.object_identifier != self.table_info_identifier
                || occurrence.message_index != self.table_info_message_index
                || !self.occurrence_matches_field(
                    occurrence,
                    self.model_identifier,
                    &[TABLE_MODEL_FIELD],
                )
            {
                self.invalid = true;
            } else {
                self.model_edges = self.model_edges.saturating_add(1);
            }
        } else if target == self.table_info_identifier {
            if occurrence.object_identifier != self.slide_identifier
                || occurrence.message_index != self.slide_message_index
                || !self.occurrence_matches_field(
                    occurrence,
                    self.table_info_identifier,
                    &[SLIDE_OWNED_DRAWABLES_FIELD, SLIDE_Z_ORDER_FIELD],
                )
            {
                self.invalid = true;
            }
        } else if self.style_ids[..self.style_count].contains(&target) {
            let source_is_selected_model = occurrence.object_identifier == self.model_identifier;
            let expected_model_style = self.style_count > 0 && self.style_ids[0] == target;
            let path_is_valid = if source_is_model {
                (!source_is_selected_model
                    || (expected_model_style
                        && occurrence.message_index == self.model_message_index))
                    && self.occurrence_matches_field_path(occurrence, target, &[3])
            } else if source_is_style {
                self.occurrence_matches_field_path(occurrence, target, &[1, 3])
            } else if source_is_network {
                self.occurrence_matches_field_path(occurrence, target, &[9])
            } else if source_is_stylesheet {
                self.occurrence_matches_field_path(occurrence, target, &[1])
            } else {
                false
            };
            if !path_is_valid {
                self.invalid = true;
            }
        } else if self.preset_identifier == Some(target) {
            if !source_is_model
                || (occurrence.object_identifier == self.model_identifier
                    && occurrence.message_index != self.model_message_index)
                || !self.occurrence_matches_field_path(occurrence, target, &[48])
            {
                self.invalid = true;
            }
        } else if self.network_identifier == Some(target) {
            let source_is_preset = source_message_type == TABLE_STYLE_PRESET_MESSAGE_TYPE;
            if !source_is_preset || !self.occurrence_matches_field_path(occurrence, target, &[3]) {
                self.invalid = true;
            }
        } else if self.stylesheet_identifier == Some(target)
            && (!source_is_style && !source_is_stylesheet && !source_is_show_aggregate
                || (source_is_style
                    && !self.occurrence_matches_field_path(occurrence, target, &[1, 5]))
                || (source_is_stylesheet
                    && !self.occurrence_matches_field_path(occurrence, target, &[1])))
        {
            self.invalid = true;
        }
        Ok(())
    }
}

impl AppearanceInboundVisitor<'_> {
    fn occurrence_matches_field_path(
        &self,
        occurrence: ArchiveReferenceOccurrence,
        target: u64,
        expected_path: &[u32],
    ) -> bool {
        match occurrence.scope {
            ArchiveReferenceScope::Message => true,
            ArchiveReferenceScope::Field { field_index } => self
                .package
                .object_with_component(occurrence.object_identifier)
                .and_then(|(_, object)| {
                    object
                        .archive_info
                        .message_infos
                        .get(occurrence.message_index)
                })
                .and_then(|message| message.field_infos.get(field_index))
                .is_some_and(|field| {
                    field.path.as_slice() == expected_path
                        && field.object_references.contains(&target)
                        && field
                            .r#type
                            .is_none_or(|kind| kind == FieldType::ObjectReference)
                }),
        }
    }

    fn occurrence_matches_field(
        &self,
        occurrence: ArchiveReferenceOccurrence,
        target: u64,
        paths: &[u32],
    ) -> bool {
        match occurrence.scope {
            ArchiveReferenceScope::Message => true,
            ArchiveReferenceScope::Field { field_index } => self
                .package
                .object_with_component(occurrence.object_identifier)
                .and_then(|(_, object)| {
                    object
                        .archive_info
                        .message_infos
                        .get(occurrence.message_index)
                })
                .and_then(|message| message.field_infos.get(field_index))
                .is_some_and(|field| {
                    field.path.as_slice().len() == 1
                        && paths.contains(&field.path.as_slice()[0])
                        && field.object_references.contains(&target)
                        && field
                            .r#type
                            .is_none_or(|kind| kind == FieldType::ObjectReference)
                }),
        }
    }
}

fn validate_global_inbound_references(
    package: &Package,
    slide_identifier: u64,
    slide_message_index: usize,
    table_info_identifier: u64,
    table_info_message_index: usize,
    model_identifier: u64,
    model_message_index: usize,
    resolved: &ResolvedAppearance,
    budget: &mut AppearanceBudget,
) -> Result<(), SlideTableAppearanceError> {
    let metadata = metadata_payload(package, SlideTableAppearancePath::Package, budget)?;
    budget.physical(metadata.len())?;
    budget.allocations(5)?;
    let mut metadata_facts = MetadataFacts::default();
    budget.scratch(MetadataFacts::reservation_scratch(metadata.len())?)?;
    metadata_facts.reserve_for_source(metadata.len())?;
    let metadata_inspection = inspect_package_metadata_with_visitor(
        metadata,
        budget.metadata_options(package, metadata.len(), 4),
        &mut metadata_facts,
    )
    .map_err(map_metadata_error)?;
    budget.metadata_report(metadata_inspection.report())?;
    if metadata_facts.unknown {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
    }
    budget.scratch(
        metadata_facts
            .data_identifiers
            .len()
            .checked_mul(size_of::<u64>())
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    let mut data_identifiers = HashSet::new();
    data_identifiers
        .try_reserve(metadata_facts.data_identifiers.len())
        .map_err(|_| SlideTableAppearanceError::Allocation {
            amount: metadata_facts.data_identifiers.len(),
        })?;
    for identifier in metadata_facts.data_identifiers {
        if identifier == 0 {
            return Err(SlideTableAppearanceError::InvalidSource);
        }
        data_identifiers.insert(identifier);
    }
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let (_, table_info_object) = package
        .object_with_component(table_info_identifier)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let table_info_message = table_info_object
        .archive_info
        .message_infos
        .get(table_info_message_index)
        .ok_or(SlideTableAppearanceError::InvalidSource)?;
    let expected_model_edges = 1usize.saturating_add(
        table_info_message
            .field_infos
            .iter()
            .filter(|field| field.path.as_slice() == [TABLE_MODEL_FIELD])
            .count(),
    );
    let mut visitor = AppearanceInboundVisitor {
        package,
        data_identifiers: &data_identifiers,
        slide_identifier,
        slide_message_index,
        table_info_identifier,
        table_info_message_index,
        model_identifier,
        model_message_index,
        expected_model_edges,
        model_edges: 0,
        style_ids: resolved.style_ids,
        style_count: resolved.style_count,
        preset_identifier: resolved.preset_identifier,
        network_identifier: resolved.network_identifier,
        stylesheet_identifier: resolved.stylesheet_identifier,
        references: 0,
        invalid: false,
    };
    let (scan_objects, scan_messages, scan_references) =
        package.state.source.components().iter().try_fold(
            (0usize, 0usize, 0usize),
            |(objects, messages, references), component| {
                let component_objects = component.archive().objects.len();
                let component_messages = component
                    .archive()
                    .objects
                    .iter()
                    .map(|object| object.messages.len())
                    .try_fold(0usize, |total, count| total.checked_add(count))
                    .ok_or(SlideTableAppearanceError::InvalidSource)?;
                let component_references = component
                    .archive()
                    .objects
                    .iter()
                    .flat_map(|object| object.archive_info.message_infos.iter())
                    .map(|message| message.object_references.len())
                    .try_fold(0usize, |total, count| total.checked_add(count))
                    .ok_or(SlideTableAppearanceError::InvalidSource)?;
                Ok::<_, SlideTableAppearanceError>((
                    objects
                        .checked_add(component_objects)
                        .ok_or(SlideTableAppearanceError::InvalidSource)?,
                    messages
                        .checked_add(component_messages)
                        .ok_or(SlideTableAppearanceError::InvalidSource)?,
                    references
                        .checked_add(component_references)
                        .ok_or(SlideTableAppearanceError::InvalidSource)?,
                ))
            },
        )?;
    budget.transaction_work(
        scan_objects
            .checked_add(scan_messages)
            .and_then(|value| value.checked_add(scan_references))
            .ok_or(SlideTableAppearanceError::InvalidSource)?,
    )?;
    budget.allocations(scan_references)?;
    budget.scratch(scan_references.saturating_mul(size_of::<u32>()))?;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            object
                .inspect_references_with_policy_and_limits(
                    &mut visitor,
                    ArchiveReferencePolicy::RejectUnknownMetadata,
                    archive_limits,
                )
                .map_err(map_core_error)?;
        }
    }
    if visitor.references > package.semantic_limits().max_references() {
        return Err(SlideTableAppearanceError::LimitExceeded {
            kind: SlideTableAppearanceLimitKind::References,
            observed: visitor.references as u64,
            maximum: package.semantic_limits().max_references() as u64,
        });
    }
    budget.references(visitor.references)?;
    let selected_paths = validate_selected_edge_paths(
        package,
        slide_identifier,
        slide_message_index,
        table_info_identifier,
        table_info_message_index,
        model_identifier,
    )?;
    if visitor.invalid || visitor.model_edges != visitor.expected_model_edges || !selected_paths {
        return Err(SlideTableAppearanceError::UnsupportedDependency);
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
) -> Result<bool, SlideTableAppearanceError> {
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
    for field in &table_info_info.field_infos {
        if field.object_references.contains(&model_identifier)
            && field.path.as_slice() != [TABLE_MODEL_FIELD]
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn map_read_error(error: ReadError) -> SlideTableAppearanceError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => SlideTableAppearanceError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::References => SlideTableAppearanceLimitKind::References,
                SemanticLimitKind::Objects => SlideTableAppearanceLimitKind::PayloadObjects,
                SemanticLimitKind::Slides => SlideTableAppearanceLimitKind::PayloadItems,
                _ => SlideTableAppearanceLimitKind::WireWork,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        ReadError::Allocation { amount, .. } => SlideTableAppearanceError::Allocation { amount },
        _ => SlideTableAppearanceError::InvalidSource,
    }
}

fn map_metadata_error(error: metadata_codec::RewriteError) -> SlideTableAppearanceError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            metadata_codec::RewriteLimit::InputBytes { observed, maximum } => {
                (SlideTableAppearanceLimitKind::InputBytes, observed, maximum)
            },
            metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => (
                SlideTableAppearanceLimitKind::OutputBytes,
                observed,
                maximum,
            ),
            metadata_codec::RewriteLimit::Fields { observed, maximum } => {
                (SlideTableAppearanceLimitKind::WireFields, observed, maximum)
            },
            metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (SlideTableAppearanceLimitKind::WireWork, observed, maximum)
            },
            metadata_codec::RewriteLimit::Nesting { observed, maximum } => (
                SlideTableAppearanceLimitKind::WireNesting,
                usize::try_from(observed).unwrap_or(usize::MAX),
                usize::try_from(maximum).unwrap_or(usize::MAX),
            ),
            metadata_codec::RewriteLimit::Components { observed, maximum } => {
                (SlideTableAppearanceLimitKind::Components, observed, maximum)
            },
            metadata_codec::RewriteLimit::References { observed, maximum } => {
                (SlideTableAppearanceLimitKind::References, observed, maximum)
            },
            metadata_codec::RewriteLimit::Additions { observed, maximum } => (
                SlideTableAppearanceLimitKind::PayloadItems,
                observed,
                maximum,
            ),
            _ => return SlideTableAppearanceError::InvalidSource,
        };
        return SlideTableAppearanceError::LimitExceeded {
            kind,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(amount) = error.allocation_request() {
        return SlideTableAppearanceError::Allocation { amount };
    }
    SlideTableAppearanceError::InvalidSource
}

fn map_appearance_codec_error(error: appearance_codec::DecodeError) -> SlideTableAppearanceError {
    if let Some(amount) = error.allocation_requested() {
        return SlideTableAppearanceError::Allocation { amount };
    }
    match error.resource_limit() {
        Some(appearance_codec::DecodeLimit::InputBytes { observed, maximum }) => {
            SlideTableAppearanceError::LimitExceeded {
                kind: SlideTableAppearanceLimitKind::WireBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(appearance_codec::DecodeLimit::OutputBytes { observed, maximum }) => {
            SlideTableAppearanceError::LimitExceeded {
                kind: SlideTableAppearanceLimitKind::WireOutputBytes,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(appearance_codec::DecodeLimit::Fields { observed, maximum }) => {
            SlideTableAppearanceError::LimitExceeded {
                kind: SlideTableAppearanceLimitKind::WireFields,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(appearance_codec::DecodeLimit::WorkBytes { observed, maximum }) => {
            SlideTableAppearanceError::LimitExceeded {
                kind: SlideTableAppearanceLimitKind::WireWork,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(appearance_codec::DecodeLimit::Nesting { observed, maximum }) => {
            SlideTableAppearanceError::LimitExceeded {
                kind: SlideTableAppearanceLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            }
        },
        Some(appearance_codec::DecodeLimit::Styles { observed, maximum }) => {
            SlideTableAppearanceError::LimitExceeded {
                kind: SlideTableAppearanceLimitKind::Styles,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(appearance_codec::DecodeLimit::Allocations { observed, maximum }) => {
            SlideTableAppearanceError::LimitExceeded {
                kind: SlideTableAppearanceLimitKind::Allocations,
                observed: observed as u64,
                maximum: maximum as u64,
            }
        },
        Some(_) => SlideTableAppearanceError::InvalidSource,
        None => SlideTableAppearanceError::InvalidSource,
    }
}

fn decode_table_info(
    payload: &[u8],
    limits: WireLimits,
) -> Result<table_info_codec::TableInfoSnapshot, SlideTableAppearanceError> {
    table_info_codec::decode_table_info(
        payload,
        table_info_codec::DecodeOptions::new(
            limits.max_input_bytes().min(payload.len().max(1)),
            limits.max_fields(),
            limits.max_rewrite_work(),
            u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX),
        ),
    )
    .map_err(map_table_info_error)
}

fn map_table_info_error(error: table_info_codec::DecodeError) -> SlideTableAppearanceError {
    if let Some((observed, maximum)) = error.field_limit_values() {
        return SlideTableAppearanceError::LimitExceeded {
            kind: SlideTableAppearanceLimitKind::WireFields,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some((observed, maximum)) = error.work_limit_values() {
        return SlideTableAppearanceError::LimitExceeded {
            kind: SlideTableAppearanceLimitKind::WireWork,
            observed: observed as u64,
            maximum: maximum as u64,
        };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            table_info_codec::WireResourceLimit::Bytes { observed, maximum } => {
                SlideTableAppearanceError::LimitExceeded {
                    kind: SlideTableAppearanceLimitKind::WireBytes,
                    observed: observed.unwrap_or(0) as u64,
                    maximum: maximum.unwrap_or(0) as u64,
                }
            },
            table_info_codec::WireResourceLimit::Nesting { observed, maximum } => {
                SlideTableAppearanceError::LimitExceeded {
                    kind: SlideTableAppearanceLimitKind::WireNesting,
                    observed: observed.unwrap_or(0) as u64,
                    maximum: maximum.unwrap_or(0) as u64,
                }
            },
            _ => SlideTableAppearanceError::InvalidSource,
        };
    }
    SlideTableAppearanceError::InvalidSource
}

fn map_wire_error(error: litchi_iwa_common::Error) -> SlideTableAppearanceError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => SlideTableAppearanceError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => {
                    SlideTableAppearanceLimitKind::WireBytes
                },
                litchi_iwa_common::LimitKind::Fields => SlideTableAppearanceLimitKind::WireFields,
                litchi_iwa_common::LimitKind::OutputBytes => {
                    SlideTableAppearanceLimitKind::WireOutputBytes
                },
                litchi_iwa_common::LimitKind::Nesting => SlideTableAppearanceLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => {
                    SlideTableAppearanceLimitKind::WireWork
                },
                litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    SlideTableAppearanceLimitKind::PayloadItems
                },
            },
            observed: observed as u64,
            maximum: limit as u64,
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            SlideTableAppearanceError::Allocation { amount }
        },
        _ => SlideTableAppearanceError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> SlideTableAppearanceError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableAppearanceError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => {
                    SlideTableAppearanceLimitKind::InputBytes
                },
                litchi_iwa_archive::LimitKind::OutputBytes => {
                    SlideTableAppearanceLimitKind::OutputBytes
                },
                litchi_iwa_archive::LimitKind::Entries => SlideTableAppearanceLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes
                | litchi_iwa_archive::LimitKind::CompressedEntryBytes
                | litchi_iwa_archive::LimitKind::MemberNameBytes
                | litchi_iwa_archive::LimitKind::MetadataBytes => {
                    SlideTableAppearanceLimitKind::EntryBytes
                },
                litchi_iwa_archive::LimitKind::TotalBytes => {
                    SlideTableAppearanceLimitKind::TotalBytes
                },
                litchi_iwa_archive::LimitKind::IwaStreamBytes => {
                    SlideTableAppearanceLimitKind::WireBytes
                },
                litchi_iwa_archive::LimitKind::IwaTotalBytes => {
                    SlideTableAppearanceLimitKind::TotalBytes
                },
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            SlideTableAppearanceError::Allocation { amount }
        },
        litchi_iwa_archive::Error::Iwa(error) => map_core_error(error),
        _ => SlideTableAppearanceError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> SlideTableAppearanceError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => SlideTableAppearanceError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::Objects => {
                    SlideTableAppearanceLimitKind::PayloadObjects
                },
                litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject => {
                    SlideTableAppearanceLimitKind::PayloadMessages
                },
                litchi_iwa_core::LimitKind::HeaderNesting => {
                    SlideTableAppearanceLimitKind::WireNesting
                },
                litchi_iwa_core::LimitKind::HeaderFields => {
                    SlideTableAppearanceLimitKind::WireFields
                },
                litchi_iwa_core::LimitKind::ArchiveBytes
                | litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes => {
                    SlideTableAppearanceLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::MetadataItems => {
                    SlideTableAppearanceLimitKind::PayloadItems
                },
                litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    SlideTableAppearanceLimitKind::EntryBytes
                },
                litchi_iwa_core::LimitKind::SnappyFrames => SlideTableAppearanceLimitKind::Entries,
            },
            observed: observed as u64,
            maximum: maximum as u64,
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideTableAppearanceError::Allocation { amount: requested }
        },
        _ => SlideTableAppearanceError::InvalidSource,
    }
}

fn fingerprint(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x0000_0100_0000_01b3);
    }
    hash
}
