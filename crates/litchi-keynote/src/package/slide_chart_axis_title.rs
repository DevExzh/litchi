//! Exact-source, selector-first Keynote chart-axis-title transactions.
//!
//! The public surface in this module is deliberately semantic: callers select
//! a slide and then a chart by checked position or exact visible title. The
//! native chart drawable, title stand-in, non-style object, component name,
//! and generated extension remain private to this adapter.

#![allow(
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The transaction redacts lower-layer failures and keeps native graph adapters private."
)]

use std::collections::HashSet;
use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::{
    WireLimits, decode_varint_from_bytes, encode_varint_into,
    varint::encoded_len,
    wire::{WireField, WireView, parse_wire_fields_with_limits},
};
use litchi_iwa_core::{
    Archive, ArchiveObject, ArchiveReferenceKind, ArchiveReferenceOccurrence,
    ArchiveReferencePolicy, ArchiveReferenceScope, ArchiveReferenceVisitor, RawMessage,
    SnappyStream,
};
use litchi_iwa_protos::keynote_chart_axis_title_codec::{
    AxisTitleKind, AxisTitleWrite, DecodeError as ChartAxisTitleDecodeError,
    DecodeLimit as ChartAxisTitleDecodeLimit, DecodeOptions, DecodeReport,
    RewriteExecutionRequirements, WireResourceLimit as ChartAxisTitleWireResourceLimit,
    decode_axis_titles_with_report, prepare_axis_title_rewrite,
};
use litchi_iwa_protos::package_metadata_codec;
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{Axis, ChartSelector, SlideSelector};

const CHART_MESSAGE_TYPE: u32 = 5_021;
const CHART_NON_STYLE_MESSAGE_TYPE: u32 = 5_023;
const CHART_AXIS_MESSAGE_TYPE: u32 = 5_027;
const STANDIN_MESSAGE_TYPE: u32 = 3_097;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const DRAWABLE_SUPER_FIELD: u32 = 1;
const DRAWABLE_TITLE_FIELD: u32 = 10;
const CHART_EXTENSION_FIELD: u32 = 10_000;
const CHART_AXIS_VALUE_FIELD: u32 = 14;
const CHART_AXIS_CATEGORY_FIELD: u32 = 16;
const DRAWABLE_LOCKED_FIELD: u32 = 5;
const GENERATED_CHART_AXIS_EXTENSION_FIELD: u32 = 10_000;
const MAX_CHART_TITLE_BYTES: usize = 64 * 1024 * 1024;
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;

/// One aggregate finite ledger shared by selection, rewrite, reopen, and
/// locality verification for a single public operation.
struct AxisTitleBudget {
    limits: WireLimits,
    maximum_input: usize,
    maximum_output: usize,
    maximum_fields: usize,
    maximum_work: usize,
    maximum_components: usize,
    maximum_references: usize,
    maximum_allocations: usize,
    maximum_retained: usize,
    maximum_scratch: usize,
    input: usize,
    output: usize,
    fields: usize,
    max_depth: u32,
    components: usize,
    references: usize,
    work: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
    candidate_reopens: usize,
}

impl AxisTitleBudget {
    fn new(package: &Package) -> Result<Self, ChartAxisTitleError> {
        let limits = package.wire_limits().map_err(map_wire_error)?;
        let physical_input = usize::try_from(package.state.options.archive().max_input_bytes())
            .map_err(|_error| ChartAxisTitleError::InvalidSource)?;
        let aggregate_bytes = physical_input
            .checked_mul(8)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        let aggregate_memory = physical_input
            .checked_mul(64)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        let aggregate_items = 16usize;
        let maximum_allocations = physical_input
            .checked_add(64)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        Ok(Self {
            limits,
            maximum_input: aggregate_bytes,
            maximum_output: aggregate_bytes,
            maximum_fields: limits
                .max_fields()
                .checked_mul(aggregate_items)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
            maximum_work: limits
                .max_rewrite_work()
                .checked_mul(aggregate_items)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
            maximum_components: package
                .semantic_limits()
                .max_objects()
                .checked_mul(aggregate_items)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
            maximum_references: package
                .semantic_limits()
                .max_references()
                .checked_mul(aggregate_items)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
            maximum_allocations,
            maximum_retained: aggregate_memory,
            maximum_scratch: aggregate_memory,
            input: 0,
            output: 0,
            fields: 0,
            max_depth: 0,
            components: 0,
            references: 0,
            work: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
            candidate_reopens: 0,
        })
    }

    fn charge_counter(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: ChartAxisTitleLimitKind,
    ) -> Result<(), ChartAxisTitleError> {
        let observed = current
            .checked_add(amount)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        if observed > maximum {
            return Err(ChartAxisTitleError::LimitExceeded {
                kind,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            });
        }
        *current = observed;
        Ok(())
    }

    fn charge_input(&mut self, amount: usize) -> Result<(), ChartAxisTitleError> {
        Self::charge_counter(
            &mut self.input,
            amount,
            self.maximum_input,
            ChartAxisTitleLimitKind::InputBytes,
        )
    }

    fn charge_output(&mut self, amount: usize) -> Result<(), ChartAxisTitleError> {
        Self::charge_counter(
            &mut self.output,
            amount,
            self.maximum_output,
            ChartAxisTitleLimitKind::OutputBytes,
        )
    }

    fn charge_fields(&mut self, amount: usize) -> Result<(), ChartAxisTitleError> {
        Self::charge_counter(
            &mut self.fields,
            amount,
            self.maximum_fields,
            ChartAxisTitleLimitKind::WireFields,
        )
    }

    fn charge(&mut self, amount: usize) -> Result<(), ChartAxisTitleError> {
        Self::charge_counter(
            &mut self.work,
            amount,
            self.maximum_work,
            ChartAxisTitleLimitKind::WireWork,
        )
    }

    fn charge_components(&mut self, amount: usize) -> Result<(), ChartAxisTitleError> {
        Self::charge_counter(
            &mut self.components,
            amount,
            self.maximum_components,
            ChartAxisTitleLimitKind::Entries,
        )
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), ChartAxisTitleError> {
        Self::charge_counter(
            &mut self.references,
            amount,
            self.maximum_references,
            ChartAxisTitleLimitKind::References,
        )
    }

    fn charge_depth(&mut self, depth: u32) -> Result<(), ChartAxisTitleError> {
        self.max_depth = self.max_depth.max(depth);
        let maximum = u32::try_from(self.limits.max_nesting())
            .map_err(|_error| ChartAxisTitleError::InvalidSource)?;
        if self.max_depth > maximum {
            return Err(ChartAxisTitleError::LimitExceeded {
                kind: ChartAxisTitleLimitKind::WireNesting,
                observed: u64::from(self.max_depth),
                maximum: u64::from(maximum),
            });
        }
        Ok(())
    }

    fn charge_allocation(
        &mut self,
        count: usize,
        retained: usize,
        scratch: usize,
    ) -> Result<(), ChartAxisTitleError> {
        Self::charge_counter(
            &mut self.allocations,
            count,
            self.maximum_allocations,
            ChartAxisTitleLimitKind::Entries,
        )?;
        Self::charge_counter(
            &mut self.retained,
            retained,
            self.maximum_retained,
            ChartAxisTitleLimitKind::TotalBytes,
        )?;
        Self::charge_counter(
            &mut self.scratch,
            scratch,
            self.maximum_scratch,
            ChartAxisTitleLimitKind::TotalBytes,
        )
    }

    fn charge_wire_vector(&mut self, payload: usize) -> Result<(), ChartAxisTitleError> {
        let capacity = payload
            .checked_mul(size_of::<WireField>())
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        self.charge_input(payload)?;
        self.charge(payload)?;
        // One byte is the smallest possible field, so this also bounds any
        // geometric growth allocations performed inside the wire scanner.
        self.charge_allocation(payload.max(1), capacity, capacity)
    }

    fn finish_wire_scan(&mut self, fields: usize) -> Result<(), ChartAxisTitleError> {
        self.charge_fields(fields)?;
        self.charge(fields)
    }

    fn charge_reference_vector(&mut self, capacity: usize) -> Result<(), ChartAxisTitleError> {
        let bytes = capacity
            .checked_mul(size_of::<u64>())
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        self.charge_allocation(1, bytes, bytes)
    }

    fn charge_owned_clone(&mut self, bytes: usize) -> Result<(), ChartAxisTitleError> {
        self.charge(bytes)?;
        self.charge_allocation(1, bytes, bytes)
    }

    fn charge_enclosing_buffer(&mut self, bytes: usize) -> Result<(), ChartAxisTitleError> {
        self.charge_output(bytes)?;
        self.charge(bytes)?;
        self.charge_allocation(1, bytes, bytes)
    }

    fn charge_decode(&mut self, report: DecodeReport) -> Result<(), ChartAxisTitleError> {
        self.charge_input(report.source_bytes())?;
        self.charge_output(report.output_bytes())?;
        self.charge_fields(report.fields())?;
        self.charge(report.work_bytes())?;
        self.charge_depth(report.max_depth())?;
        self.charge_allocation(
            report.allocations(),
            report.retained_bytes(),
            report.scratch_bytes(),
        )
    }

    fn charge_metadata_report(
        &mut self,
        report: package_metadata_codec::RewriteReport,
    ) -> Result<(), ChartAxisTitleError> {
        self.charge_input(report.input_bytes())?;
        self.charge_output(report.output_bytes())?;
        self.charge_fields(report.fields())?;
        self.charge(report.work_bytes())?;
        self.charge_depth(report.max_depth())?;
        self.charge_components(report.components_scanned())?;
        self.charge_references(report.references_scanned())?;
        self.charge_allocation(
            report.allocations(),
            report.retained_bytes(),
            report.scratch_bytes(),
        )
    }

    fn charge_prepared(
        &mut self,
        report: DecodeReport,
        requirements: RewriteExecutionRequirements,
    ) -> Result<(), ChartAxisTitleError> {
        if report.work_bytes() > requirements.work_bytes
            || report.fields() > requirements.fields
            || report.max_depth() > requirements.max_depth
        {
            return Err(ChartAxisTitleError::InvalidSource);
        }
        self.charge_input(report.source_bytes())?;
        self.charge_output(requirements.output_bytes)?;
        self.charge_fields(requirements.fields)?;
        self.charge(requirements.work_bytes)?;
        self.charge_depth(requirements.max_depth)?;
        self.charge_allocation(
            requirements.allocations,
            requirements.retained_bytes,
            requirements.scratch_bytes,
        )
    }

    fn charge_catalog_scan(&mut self, package: &Package) -> Result<(), ChartAxisTitleError> {
        let mut work = 0usize;
        let mut references = 0usize;
        for component in package.state.source.components().iter() {
            let archive_bytes = component.archive().encoded_len().map_err(map_core_error)?;
            work = work
                .checked_add(component.name().len())
                .and_then(|value| value.checked_add(archive_bytes))
                .ok_or(ChartAxisTitleError::InvalidSource)?;
            for object in &component.archive().objects {
                for (index, message) in object.messages.iter().enumerate() {
                    work = work
                        .checked_add(message.data.len())
                        .ok_or(ChartAxisTitleError::InvalidSource)?;
                    let info = object
                        .archive_info
                        .message_infos
                        .get(index)
                        .ok_or(ChartAxisTitleError::InvalidSource)?;
                    references = references
                        .checked_add(info.object_references.len())
                        .and_then(|value| value.checked_add(info.data_references.len()))
                        .ok_or(ChartAxisTitleError::InvalidSource)?;
                    for field in &info.field_infos {
                        references = references
                            .checked_add(field.object_references.len())
                            .and_then(|value| value.checked_add(field.data_references.len()))
                            .ok_or(ChartAxisTitleError::InvalidSource)?;
                    }
                }
            }
        }
        // A prepared directory source keeps only the semantic component
        // catalog and deliberately has no ZIP byte stream.  The chart-axis
        // read path is still valid for that source, so account physical bytes
        // only when they actually exist instead of routing through
        // `PhysicalSource::shared_source` (which is intentionally
        // physical-only and panics for semantic sources).
        let source_bytes = match &package.state.source {
            PhysicalSource::Package(source) => source.shared_source().len(),
            PhysicalSource::Semantic(_) => 0,
        };
        self.charge_input(source_bytes)?;
        self.charge_components(package.state.total_objects)?;
        self.charge_references(references)?;
        self.charge(work)
    }

    fn charge_scan_pass(
        &mut self,
        package: &Package,
        retained_vectors: usize,
    ) -> Result<(), ChartAxisTitleError> {
        let inventory = package_scan_inventory(package)?;
        self.charge_components(inventory.objects)?;
        self.charge_references(inventory.references)?;
        self.charge_fields(inventory.metadata_fields)?;
        self.charge(inventory.scan_work()?)?;
        self.charge_allocation(retained_vectors, inventory.logical_index_bytes()?, 0)
    }

    fn charge_selection_scans(
        &mut self,
        package: &Package,
        mutation_guards: bool,
    ) -> Result<(), ChartAxisTitleError> {
        // Selector resolution walks the semantic slide/chart graph once. A
        // mutation additionally builds the package-wide chart-owner census.
        self.charge_scan_pass(package, 2)?;
        if mutation_guards {
            self.charge_scan_pass(package, 2)?;
        }
        Ok(())
    }

    fn charge_inbound_scan(&mut self, package: &Package) -> Result<(), ChartAxisTitleError> {
        // Archive reference inspection walks every MessageInfo and FieldInfo
        // exactly once and retains only the visitor's scalar counters.
        self.charge_scan_pass(package, 0)
    }

    fn charge_locality_scan(&mut self, package: &Package) -> Result<(), ChartAxisTitleError> {
        // Locality compares the physical catalog and the selected component's
        // complete object/message inventory without constructing a new index.
        let inventory = package_scan_inventory(package)?;
        let entries = physical_catalog(package)?.package().len();
        self.charge_components(inventory.objects)?;
        self.charge_references(inventory.references)?;
        self.charge_fields(inventory.metadata_fields)?;
        self.charge(
            inventory
                .scan_work()?
                .checked_add(entries)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
        )
    }

    fn charge_native_decompress_parse(
        &mut self,
        compressed: usize,
        decompressed: usize,
        objects: usize,
    ) -> Result<(), ChartAxisTitleError> {
        self.charge_input(compressed)?;
        self.charge_output(decompressed)?;
        self.charge(
            compressed
                .checked_add(decompressed)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
        )?;
        self.charge_allocation(
            objects
                .checked_add(2)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
            decompressed
                .checked_mul(2)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
            compressed,
        )
    }

    fn charge_archive_snappy_plan(
        &mut self,
        archive: &Archive,
        archive_limits: litchi_iwa_core::Limits,
        snappy_limits: litchi_iwa_core::SnappyLimits,
    ) -> Result<(), ChartAxisTitleError> {
        let encoded = archive
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let compressed = SnappyStream::maximum_compressed_len(encoded).map_err(map_core_error)?;
        if compressed > snappy_limits.max_compressed_stream() {
            return Err(ChartAxisTitleError::LimitExceeded {
                kind: ChartAxisTitleLimitKind::EntryBytes,
                observed: usize_to_u64(compressed),
                maximum: usize_to_u64(snappy_limits.max_compressed_stream()),
            });
        }
        self.charge_output(
            encoded
                .checked_add(compressed)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
        )?;
        self.charge(
            encoded
                .checked_add(compressed)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
        )?;
        self.charge_allocation(
            2,
            encoded
                .checked_add(compressed)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
            encoded,
        )
    }

    fn charge_reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<litchi_iwa_archive::package::ReassemblyExecutionLimits, ChartAxisTitleError> {
        self.charge_output(requirements.output_bytes())?;
        self.charge(
            requirements
                .output_bytes()
                .checked_add(requirements.scratch_bytes())
                .and_then(|value| value.checked_add(requirements.offset_count()))
                .ok_or(ChartAxisTitleError::InvalidSource)?,
        )?;
        self.charge_allocation(
            requirements.allocations(),
            requirements.retained_bytes(),
            requirements.scratch_bytes(),
        )?;
        Ok(requirements.exact_limits())
    }

    fn charge_candidate_reopen(
        &mut self,
        source: &Package,
        bytes: usize,
        replacement: Option<(&str, &Archive)>,
    ) -> Result<(), ChartAxisTitleError> {
        let envelope = candidate_reopen_envelope(source, bytes, replacement)?;
        self.charge_input(bytes)?;
        self.charge_components(envelope.objects)?;
        self.charge_references(envelope.references)?;
        self.charge_fields(envelope.metadata_fields)?;
        self.charge(envelope.work)?;
        self.charge_allocation(envelope.allocations, envelope.retained, envelope.scratch)?;
        self.candidate_reopens = self
            .candidate_reopens
            .checked_add(1)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        Ok(())
    }

    fn charge_exact_artifacts(
        &mut self,
        source: usize,
        target: usize,
    ) -> Result<(), ChartAxisTitleError> {
        self.charge(
            source
                .checked_add(target)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
        )
    }

    fn codec_options(&self, source_len: usize) -> Result<DecodeOptions, ChartAxisTitleError> {
        let input = self
            .maximum_input
            .checked_sub(self.input)
            .ok_or(ChartAxisTitleError::InvalidSource)?
            .min(self.limits.max_input_bytes());
        let output = self
            .maximum_output
            .checked_sub(self.output)
            .ok_or(ChartAxisTitleError::InvalidSource)?
            .min(self.limits.max_output_bytes());
        let fields = self
            .maximum_fields
            .checked_sub(self.fields)
            .ok_or(ChartAxisTitleError::InvalidSource)?
            .min(self.limits.max_fields());
        let work = self
            .maximum_work
            .checked_sub(self.work)
            .ok_or(ChartAxisTitleError::InvalidSource)?
            .min(self.limits.max_rewrite_work());
        if input < source_len || input == 0 {
            return Err(ChartAxisTitleError::LimitExceeded {
                kind: ChartAxisTitleLimitKind::InputBytes,
                observed: usize_to_u64(source_len.max(1)),
                maximum: usize_to_u64(input),
            });
        }
        let recursion = u32::try_from(self.limits.max_nesting())
            .map_err(|_error| ChartAxisTitleError::InvalidSource)?;
        Ok(DecodeOptions::new(input, fields, work, recursion)
            .with_max_output_bytes(output)
            .with_max_title_bytes(MAX_CHART_TITLE_BYTES.min(self.limits.max_output_bytes()))
            .with_max_allocations(
                self.maximum_allocations
                    .checked_sub(self.allocations)
                    .ok_or(ChartAxisTitleError::InvalidSource)?,
            )
            .with_max_retained_bytes(
                self.maximum_retained
                    .checked_sub(self.retained)
                    .ok_or(ChartAxisTitleError::InvalidSource)?,
            )
            .with_max_scratch_bytes(
                self.maximum_scratch
                    .checked_sub(self.scratch)
                    .ok_or(ChartAxisTitleError::InvalidSource)?,
            ))
    }
}

#[derive(Clone, Copy, Default)]
struct PackageScanInventory {
    components: usize,
    objects: usize,
    messages: usize,
    message_bytes: usize,
    metadata_fields: usize,
    references: usize,
}

impl PackageScanInventory {
    fn add_archive(&mut self, archive: &Archive) -> Result<(), ChartAxisTitleError> {
        self.components = self
            .components
            .checked_add(1)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        self.objects = self
            .objects
            .checked_add(archive.objects.len())
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        for object in &archive.objects {
            self.messages = self
                .messages
                .checked_add(object.messages.len())
                .ok_or(ChartAxisTitleError::InvalidSource)?;
            self.metadata_fields = self
                .metadata_fields
                .checked_add(object.archive_info.message_infos.len())
                .ok_or(ChartAxisTitleError::InvalidSource)?;
            for message in &object.messages {
                self.message_bytes = self
                    .message_bytes
                    .checked_add(message.data.len())
                    .ok_or(ChartAxisTitleError::InvalidSource)?;
            }
            for info in &object.archive_info.message_infos {
                self.references = self
                    .references
                    .checked_add(info.object_references.len())
                    .and_then(|value| value.checked_add(info.data_references.len()))
                    .ok_or(ChartAxisTitleError::InvalidSource)?;
                self.metadata_fields = self
                    .metadata_fields
                    .checked_add(info.field_infos.len())
                    .ok_or(ChartAxisTitleError::InvalidSource)?;
                for field in &info.field_infos {
                    self.references = self
                        .references
                        .checked_add(field.object_references.len())
                        .and_then(|value| value.checked_add(field.data_references.len()))
                        .ok_or(ChartAxisTitleError::InvalidSource)?;
                }
            }
        }
        Ok(())
    }

    fn scan_work(self) -> Result<usize, ChartAxisTitleError> {
        self.components
            .checked_add(self.objects)
            .and_then(|value| value.checked_add(self.messages))
            .and_then(|value| value.checked_add(self.message_bytes))
            .and_then(|value| value.checked_add(self.metadata_fields))
            .and_then(|value| value.checked_add(self.references))
            .ok_or(ChartAxisTitleError::InvalidSource)
    }

    fn logical_index_bytes(self) -> Result<usize, ChartAxisTitleError> {
        self.objects
            .checked_mul(3 * size_of::<usize>())
            .and_then(|value| {
                self.messages
                    .checked_mul(size_of::<RawMessage>())
                    .and_then(|messages| value.checked_add(messages))
            })
            .and_then(|value| {
                self.references
                    .checked_mul(size_of::<u64>())
                    .and_then(|references| value.checked_add(references))
            })
            .ok_or(ChartAxisTitleError::InvalidSource)
    }
}

fn package_scan_inventory(package: &Package) -> Result<PackageScanInventory, ChartAxisTitleError> {
    let mut inventory = PackageScanInventory::default();
    for component in package.state.source.components().iter() {
        inventory.add_archive(component.archive())?;
    }
    Ok(inventory)
}

struct CandidateReopenEnvelope {
    objects: usize,
    references: usize,
    metadata_fields: usize,
    work: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
}

/// Conservatively accounts the logical package state retained by reopen.
/// Lower-layer allocator capacity and cache layout are private, so this is a
/// checked package/inventory envelope rather than a claim about exact RSS.
fn candidate_reopen_envelope(
    source: &Package,
    candidate_bytes: usize,
    replacement: Option<(&str, &Archive)>,
) -> Result<CandidateReopenEnvelope, ChartAxisTitleError> {
    let physical_limits = source.state.options.archive();
    let candidate_u64 =
        u64::try_from(candidate_bytes).map_err(|_error| ChartAxisTitleError::InvalidSource)?;
    if candidate_u64 > physical_limits.max_input_bytes() {
        return Err(ChartAxisTitleError::LimitExceeded {
            kind: ChartAxisTitleLimitKind::OutputBytes,
            observed: candidate_u64,
            maximum: physical_limits.max_input_bytes(),
        });
    }
    let catalog = physical_catalog(source)?.package();
    if catalog.len() > physical_limits.max_entries() {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    let archive_limits = physical_limits
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut inventory = PackageScanInventory::default();
    let mut decompressed = 0usize;
    let mut largest_component = 0usize;
    for component in source.state.source.components().iter() {
        let archive = replacement
            .filter(|(name, _archive)| *name == component.name())
            .map_or(component.archive(), |(_name, archive)| archive);
        let bytes = archive
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_error)?;
        if bytes > physical_limits.max_iwa_stream_bytes() {
            return Err(ChartAxisTitleError::InvalidSource);
        }
        decompressed = decompressed
            .checked_add(bytes)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        largest_component = largest_component.max(bytes);
        inventory.add_archive(archive)?;
    }
    if replacement.is_none() {
        // Applied patches are owner-produced exact artifacts. Their graph is
        // verified after reopen; admit title growth conservatively before the
        // opaque lower-layer package caches are constructed.
        let growth = candidate_bytes.saturating_sub(source.source_bytes().len());
        decompressed = decompressed
            .checked_add(growth)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        largest_component = largest_component
            .checked_add(growth)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
    }
    let catalog_index = catalog
        .len()
        .checked_mul(size_of::<[usize; 8]>())
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let logical = inventory.logical_index_bytes()?;
    let retained = candidate_bytes
        .checked_add(
            decompressed
                .checked_mul(2)
                .ok_or(ChartAxisTitleError::InvalidSource)?,
        )
        .and_then(|value| value.checked_add(logical))
        .and_then(|value| value.checked_add(catalog_index))
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let scratch = largest_component
        .checked_add(candidate_bytes)
        .and_then(|value| value.checked_add(catalog_index))
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let allocations = catalog
        .len()
        .checked_mul(2)
        .and_then(|value| value.checked_add(inventory.components))
        .and_then(|value| value.checked_add(inventory.objects.checked_mul(3)?))
        .and_then(|value| value.checked_add(inventory.messages.checked_mul(3)?))
        .and_then(|value| value.checked_add(inventory.metadata_fields))
        .and_then(|value| value.checked_add(4))
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let work = candidate_bytes
        .checked_add(decompressed)
        .and_then(|value| inventory.scan_work().ok()?.checked_add(value))
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    Ok(CandidateReopenEnvelope {
        objects: inventory.objects,
        references: inventory.references,
        metadata_fields: inventory.metadata_fields,
        work,
        allocations,
        retained,
        scratch,
    })
}

/// A finite resource governed while a chart-axis-title transaction is prepared or
/// published.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartAxisTitleLimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package or payload bytes.
    OutputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Bytes in one package member, IWA object, or message.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Semantic text-storage objects.
    TextStorages,
    /// Semantic text fragments.
    TextFragments,
    /// Aggregate semantic text bytes.
    TextBytes,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf rewrite work.
    WireWork,
    /// UTF-8 bytes in one visible chart title.
    TitleBytes,
}

impl fmt::Display for ChartAxisTitleLimitKind {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            Self::InputBytes => "input bytes",
            Self::OutputBytes => "output bytes",
            Self::WireBytes => "wire bytes",
            Self::Entries => "entries",
            Self::EntryBytes => "entry bytes",
            Self::TotalBytes => "total bytes",
            Self::Slides => "slides",
            Self::References => "references",
            Self::TextStorages => "text storages",
            Self::TextFragments => "text fragments",
            Self::TextBytes => "text bytes",
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
            Self::TitleBytes => "chart title bytes",
        })
    }
}

/// A content-redacted failure raised by a chart-axis-title transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ChartAxisTitleError {
    /// The source was prepared without an exact physical package artifact.
    #[error("this Keynote source does not support physical chart-axis-title edits")]
    UnsupportedSource,
    /// An exact-name slide or chart selector was ambiguous.
    #[error("the Keynote chart-axis-title selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name slide selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound {
        /// Missing checked semantic slide position.
        position: Position,
    },
    /// An exact-name chart selector did not match.
    #[error("the selected Keynote slide has no chart matching the requested name")]
    ChartNameNotFound,
    /// A checked semantic chart position does not exist.
    #[error("the selected Keynote slide has no chart at position {position:?}")]
    ChartPositionNotFound {
        /// Missing checked semantic chart position.
        position: Position,
    },
    /// An empty exact chart name was supplied.
    #[error("the Keynote chart selector name cannot be empty")]
    EmptyChartName,
    /// The source chart graph, title stand-in, or selected native payload is
    /// malformed or unsupported.
    #[error("the Keynote chart-axis-title source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Keynote chart-axis-title {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its limit.
        kind: ChartAxisTitleLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote chart-axis-title transaction")]
    Allocation {
        /// Elements or bytes requested.
        amount: usize,
    },
    /// Full candidate reopening did not reproduce the requested title.
    #[error("the edited Keynote chart title failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote chart-axis-title patch does not match the exact source package")]
    PatchConflict,
}

/// One mutable chart-axis-title value staged against an immutable package snapshot.
pub struct ChartAxisTitleEdit<'a> {
    source: &'a Package,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    axis: Axis,
    axis_identifier: u64,
    axis_component_name: String,
    axis_message_index: usize,
    slide_identifier: u64,
    before: Option<String>,
    after: Option<String>,
}

impl fmt::Debug for ChartAxisTitleEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartAxisTitleEdit")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("has_before", &self.before.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl<'a> ChartAxisTitleEdit<'a> {
    fn new<'slide, 'chart>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
        axis: Axis,
    ) -> Result<Self, ChartAxisTitleError> {
        let mut budget = AxisTitleBudget::new(source)?;
        budget.charge_catalog_scan(source)?;
        let selection = select_axis(
            source,
            slide_selector.into(),
            chart_selector.into(),
            axis,
            true,
            &mut budget,
        )?;
        let before = selection.title;
        let after = before.as_deref().map(copy_title).transpose()?;
        Ok(Self {
            source,
            slide_position: selection.slide_position,
            chart_position: selection.chart_position,
            chart_identifier: selection.chart_identifier,
            axis: selection.axis,
            axis_identifier: selection.axis_identifier,
            axis_component_name: selection.axis_component_name,
            axis_message_index: selection.axis_message_index,
            slide_identifier: selection.slide_identifier,
            before,
            after,
        })
    }

    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart_position
    }

    /// Borrow the visible title observed when this edit began.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Borrow the title staged for publication.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Stage a visible chart title.
    ///
    /// The input is copied into bounded edit-owned storage immediately, so a
    /// caller may pass either a borrowed string slice or an owned `String`
    /// without changing the transaction's source ownership or wire behavior.
    pub fn set(mut self, title: impl AsRef<str>) -> Result<Self, ChartAxisTitleError> {
        self.after = Some(copy_title(title.as_ref())?);
        Ok(self)
    }

    /// Stage removal of the visible chart title.
    pub fn clear(mut self) -> Result<Self, ChartAxisTitleError> {
        self.after = None;
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<ChartAxisTitleCommit, ChartAxisTitleError> {
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let mut budget = AxisTitleBudget::new(self.source)?;
        budget.charge_catalog_scan(self.source)?;
        let source_selection = select_axis(
            self.source,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            self.axis,
            true,
            &mut budget,
        )?;
        if source_selection.chart_identifier != self.chart_identifier
            || source_selection.axis_identifier != self.axis_identifier
            || source_selection.axis_component_name != self.axis_component_name
            || source_selection.axis_message_index != self.axis_message_index
            || source_selection.slide_identifier != self.slide_identifier
            || source_selection.title != self.before
        {
            return Err(ChartAxisTitleError::InvalidSource);
        }

        if self.before == self.after {
            self.source.validate().map_err(map_read_error)?;
            return Ok(ChartAxisTitleCommit {
                package: self.source.snapshot(),
                patch: ChartAxisTitlePatch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    slide_position: self.slide_position,
                    chart_position: self.chart_position,
                    chart_identifier: self.chart_identifier,
                    axis: self.axis,
                    axis_identifier: self.axis_identifier,
                    axis_component_name: self.axis_component_name,
                    axis_message_index: self.axis_message_index,
                    slide_identifier: self.slide_identifier,
                    before: self.before,
                    after: self.after,
                    deleted_previews: 0,
                    target_requires_invalidated_previews: false,
                },
                diagnostics: ChartAxisTitleDiagnostics::unchanged(),
            });
        }

        if !catalog.source_is_exact() {
            return Err(ChartAxisTitleError::UnsupportedSource);
        }
        self.source.validate().map_err(map_read_error)?;
        let (package, deleted_previews) = rewrite_chart_axis_title(
            self.source,
            &source_selection,
            self.after.as_deref(),
            &mut budget,
        )?;
        package.validate().map_err(map_read_error)?;
        let target = physical_catalog(&package)?.shared_source();
        budget.charge_exact_artifacts(source_bytes.len(), target.len())?;
        let candidate_selection = select_axis(
            &package,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            self.axis,
            true,
            &mut budget,
        )?;
        if candidate_selection.chart_identifier != self.chart_identifier
            || candidate_selection.axis_identifier != self.axis_identifier
            || candidate_selection.axis_component_name != self.axis_component_name
            || candidate_selection.axis_message_index != self.axis_message_index
            || candidate_selection.slide_identifier != self.slide_identifier
            || candidate_selection.title != self.after
        {
            return Err(ChartAxisTitleError::Verification);
        }
        verify_candidate_reopen_locality(
            self.source,
            &package,
            self.slide_position,
            self.chart_position,
            self.axis,
            self.after.as_deref(),
            true,
            &mut budget,
        )?;
        Ok(ChartAxisTitleCommit {
            package,
            patch: ChartAxisTitlePatch {
                artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
                slide_position: self.slide_position,
                chart_position: self.chart_position,
                chart_identifier: self.chart_identifier,
                axis: self.axis,
                axis_identifier: self.axis_identifier,
                axis_component_name: self.axis_component_name,
                axis_message_index: self.axis_message_index,
                slide_identifier: self.slide_identifier,
                before: self.before,
                after: self.after,
                deleted_previews,
                target_requires_invalidated_previews: true,
            },
            diagnostics: ChartAxisTitleDiagnostics::published(deleted_previews),
        })
    }
}

/// An exact-source-checked reversible semantic chart-axis-title patch.
#[derive(Clone, PartialEq, Eq)]
pub struct ChartAxisTitlePatch {
    artifacts: ExactArtifacts,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    axis: Axis,
    axis_identifier: u64,
    axis_component_name: String,
    axis_message_index: usize,
    slide_identifier: u64,
    before: Option<String>,
    after: Option<String>,
    deleted_previews: usize,
    target_requires_invalidated_previews: bool,
}

impl fmt::Debug for ChartAxisTitlePatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartAxisTitlePatch")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("has_before", &self.before.is_some())
            .field("has_after", &self.after.is_some())
            .finish_non_exhaustive()
    }
}

impl ChartAxisTitlePatch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the selected semantic chart position within its slide.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart_position
    }

    /// Borrow the visible title required from the source package.
    #[must_use]
    pub fn before(&self) -> Option<&str> {
        self.before.as_deref()
    }

    /// Borrow the visible title produced by the target package.
    #[must_use]
    pub fn after(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the committed package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch preserves the exact source bytes and title.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before == self.after && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from the target back to its source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            slide_position: self.slide_position,
            chart_position: self.chart_position,
            chart_identifier: self.chart_identifier,
            axis: self.axis,
            axis_identifier: self.axis_identifier,
            axis_component_name: self.axis_component_name.clone(),
            axis_message_index: self.axis_message_index,
            slide_identifier: self.slide_identifier,
            before: self.after.clone(),
            after: self.before.clone(),
            deleted_previews: 0,
            target_requires_invalidated_previews: !self.artifacts.is_byte_noop()
                && !self.target_requires_invalidated_previews,
        }
    }
}

/// Compact evidence describing one chart-axis-title commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartAxisTitleDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl ChartAxisTitleDiagnostics {
    const fn unchanged() -> Self {
        Self {
            changed: false,
            touched_components: 0,
            deleted_previews: 0,
            full_reparse_performed: false,
        }
    }

    const fn published(deleted_previews: usize) -> Self {
        Self {
            changed: true,
            touched_components: 1,
            deleted_previews,
            full_reparse_performed: true,
        }
    }

    /// Return whether the committed package differs from its source.
    #[must_use]
    pub const fn changed(self) -> bool {
        self.changed
    }

    /// Return the number of physical IWA components rewritten.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return how many stale root rendering previews were deleted.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the complete candidate was reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// The fully verified result of one immutable chart-axis-title transaction.
#[must_use = "a Keynote chart-axis-title commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct ChartAxisTitleCommit {
    package: Package,
    patch: ChartAxisTitlePatch,
    diagnostics: ChartAxisTitleDiagnostics,
}

impl ChartAxisTitleCommit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its immutable package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &ChartAxisTitlePatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &ChartAxisTitleDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the visible title of one selected chart axis.
    pub fn slide_chart_axis_title<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
        axis: Axis,
    ) -> Result<Option<String>, ChartAxisTitleError> {
        let mut budget = AxisTitleBudget::new(self)?;
        budget.charge_catalog_scan(self)?;
        Ok(select_axis(
            self,
            slide_selector.into(),
            chart_selector.into(),
            axis,
            false,
            &mut budget,
        )?
        .title)
    }

    /// Start an exact immutable edit of one selected chart-axis title.
    pub fn edit_slide_chart_axis_title<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
        axis: Axis,
    ) -> Result<ChartAxisTitleEdit<'_>, ChartAxisTitleError> {
        ChartAxisTitleEdit::new(self, slide_selector, chart_selector, axis)
    }

    /// Apply an exact-source-checked chart-axis-title patch.
    pub fn apply_slide_chart_axis_title(
        &self,
        patch: &ChartAxisTitlePatch,
    ) -> Result<ChartAxisTitleCommit, ChartAxisTitleError> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        let mut budget = AxisTitleBudget::new(self)?;
        budget.charge_catalog_scan(self)?;
        if !patch.artifacts.authorizes_source(&source) {
            return Err(ChartAxisTitleError::PatchConflict);
        }
        let source_selection = select_axis(
            self,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            patch.axis,
            true,
            &mut budget,
        )?;
        if source_selection.chart_identifier != patch.chart_identifier
            || source_selection.axis_identifier != patch.axis_identifier
            || source_selection.axis_component_name != patch.axis_component_name
            || source_selection.axis_message_index != patch.axis_message_index
            || source_selection.slide_identifier != patch.slide_identifier
            || source_selection.title != patch.before
        {
            return Err(ChartAxisTitleError::PatchConflict);
        }
        if patch.is_noop() {
            self.validate().map_err(map_read_error)?;
            return Ok(ChartAxisTitleCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: ChartAxisTitleDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(ChartAxisTitleError::PatchConflict);
        }
        budget.charge_exact_artifacts(source.len(), patch.artifacts.target().len())?;
        budget.charge_candidate_reopen(self, patch.artifacts.target().len(), None)?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        candidate.validate().map_err(map_read_error)?;
        let candidate_selection = select_axis(
            &candidate,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            patch.axis,
            true,
            &mut budget,
        )?;
        if candidate_selection.chart_identifier != patch.chart_identifier
            || candidate_selection.axis_identifier != patch.axis_identifier
            || candidate_selection.axis_component_name != patch.axis_component_name
            || candidate_selection.axis_message_index != patch.axis_message_index
            || candidate_selection.slide_identifier != patch.slide_identifier
            || candidate_selection.title != patch.after
        {
            return Err(ChartAxisTitleError::Verification);
        }
        verify_candidate_reopen_locality(
            self,
            &candidate,
            patch.slide_position,
            patch.chart_position,
            patch.axis,
            patch.after.as_deref(),
            patch.target_requires_invalidated_previews,
            &mut budget,
        )?;
        Ok(ChartAxisTitleCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: ChartAxisTitleDiagnostics::published(patch.deleted_previews),
        })
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct AxisSelection {
    slide_position: Position,
    chart_position: Position,
    slide_identifier: u64,
    chart_identifier: u64,
    axis: Axis,
    axis_identifier: u64,
    axis_component_name: String,
    axis_message_index: usize,
    title: Option<String>,
}

fn select_axis(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    chart_selector: ChartSelector<'_>,
    axis: Axis,
    mutation_guards: bool,
    budget: &mut AxisTitleBudget,
) -> Result<AxisSelection, ChartAxisTitleError> {
    budget.charge_selection_scans(package, mutation_guards)?;
    let chart = super::slide_chart_title::select_chart(
        package,
        slide_selector,
        chart_selector,
        mutation_guards,
    )
    .map_err(map_chart_title_error)?;
    let (chart_component, chart_object) = package
        .object_with_component(chart.chart_identifier)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    if chart_component != chart.slide_component_name {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    let (chart_message_index, chart_message) = unique_message(chart_object, CHART_MESSAGE_TYPE)?;
    validate_selected_message_metadata(chart_object, chart_message_index)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let outer = accounted_wire_fields(&chart_message.data, limits, budget)?;
    let super_payload =
        unique_length_delimited_field(&outer, &chart_message.data, DRAWABLE_SUPER_FIELD)?
            .ok_or(ChartAxisTitleError::InvalidSource)?;
    if mutation_guards {
        validate_unlocked_drawable(super_payload, limits, budget)?;
    }
    let title_identifier = required_reference(super_payload, DRAWABLE_TITLE_FIELD, limits, budget)?;
    validate_graph_object(package, title_identifier, STANDIN_MESSAGE_TYPE)?;
    validate_graph_object(
        package,
        chart.non_style_identifier,
        CHART_NON_STYLE_MESSAGE_TYPE,
    )?;
    let chart_payload =
        unique_length_delimited_field(&outer, &chart_message.data, CHART_EXTENSION_FIELD)?
            .ok_or(ChartAxisTitleError::InvalidSource)?;
    let category = repeated_references(chart_payload, CHART_AXIS_CATEGORY_FIELD, limits, budget)?;
    let value = repeated_references(chart_payload, CHART_AXIS_VALUE_FIELD, limits, budget)?;
    validate_axis_roles(&category, &value, budget)?;
    let selected = match axis {
        Axis::Category => &category,
        Axis::Value => &value,
    };
    let axis_identifier = selected
        .iter()
        .copied()
        .find(|identifier| *identifier != 0)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let (axis_component, axis_object) = unique_object(package, axis_identifier)?;
    if axis_object.messages.len() != 1 {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    let (axis_message_index, axis_message) = unique_message(axis_object, CHART_AXIS_MESSAGE_TYPE)?;
    validate_selected_message_metadata(axis_object, axis_message_index)?;
    let kind = axis_title_kind(axis);
    let title = read_chart_axis_title(&axis_message.data, limits, kind, budget)?;
    if mutation_guards {
        prove_unique_primary_axis(package, axis_identifier, budget)?;
        let stylesheet_component_name = validate_global_axis_references(
            package,
            chart.chart_identifier,
            chart_message_index,
            axis_identifier,
            budget,
        )?;
        validate_axis_metadata(
            package,
            chart_component,
            axis_component,
            stylesheet_component_name,
            axis_identifier,
            budget,
        )?;
    }
    if axis_identifier == chart.chart_identifier
        || axis_identifier == chart.non_style_identifier
        || axis_identifier == chart.slide_identifier
    {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    Ok(AxisSelection {
        slide_position: chart.slide_position,
        chart_position: chart.chart_position,
        slide_identifier: chart.slide_identifier,
        chart_identifier: chart.chart_identifier,
        axis,
        axis_identifier,
        axis_component_name: axis_component.to_owned(),
        axis_message_index,
        title,
    })
}

fn repeated_references(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut AxisTitleBudget,
) -> Result<Vec<u64>, ChartAxisTitleError> {
    let fields = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.charge_input(payload.len())?;
    budget.finish_wire_scan(fields.len())?;
    let mut references = Vec::new();
    budget.charge_reference_vector(fields.len())?;
    references
        .try_reserve(fields.len())
        .map_err(|_error| ChartAxisTitleError::Allocation {
            amount: fields.len(),
        })?;
    for field in fields.fields() {
        if field.number() != field_number {
            continue;
        }
        if field.wire_type() != 2 {
            return Err(ChartAxisTitleError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        references.push(
            super::validate_reference_payload(field.payload(), limits, "Keynote chart drawable")
                .map_err(map_wire_error)?,
        );
    }
    budget.charge_references(references.len())?;
    Ok(references)
}

fn accounted_wire_fields(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut AxisTitleBudget,
) -> Result<Vec<WireField>, ChartAxisTitleError> {
    budget.charge_wire_vector(payload.len())?;
    let fields = parse_wire_fields_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.finish_wire_scan(fields.len())?;
    Ok(fields)
}

fn unique_length_delimited_field<'a>(
    fields: &[WireField],
    source: &'a [u8],
    number: u32,
) -> Result<Option<&'a [u8]>, ChartAxisTitleError> {
    let mut selected = None;
    for field in fields
        .iter()
        .copied()
        .filter(|field| field.number() == number)
    {
        if selected.is_some() || field.wire_type() != 2 {
            return Err(ChartAxisTitleError::InvalidSource);
        }
        field
            .validate_canonical_framing(source)
            .map_err(map_wire_error)?;
        selected = Some(field.payload(source).map_err(map_wire_error)?);
    }
    Ok(selected)
}

fn unique_message(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<(usize, &RawMessage), ChartAxisTitleError> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace((index, message)).is_some() {
            return Err(ChartAxisTitleError::InvalidSource);
        }
    }
    selected.ok_or(ChartAxisTitleError::InvalidSource)
}

fn required_reference(
    payload: &[u8],
    field_number: u32,
    limits: WireLimits,
    budget: &mut AxisTitleBudget,
) -> Result<u64, ChartAxisTitleError> {
    let fields = accounted_wire_fields(payload, limits, budget)?;
    let reference = unique_length_delimited_field(&fields, payload, field_number)?
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let identifier = super::validate_reference_payload(reference, limits, "Keynote chart drawable")
        .map_err(map_wire_error)?;
    if identifier == 0 {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    Ok(identifier)
}

fn validate_graph_object(
    package: &Package,
    identifier: u64,
    message_type: u32,
) -> Result<(), ChartAxisTitleError> {
    let (_component, object) = unique_object(package, identifier)?;
    if object.messages.len() != 1 {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    let (message_index, _message) = unique_message(object, message_type)?;
    validate_selected_message_metadata(object, message_index)
}

fn validate_unlocked_drawable(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut AxisTitleBudget,
) -> Result<(), ChartAxisTitleError> {
    let view = WireView::parse_with_limits(payload, limits).map_err(map_wire_error)?;
    budget.charge_input(payload.len())?;
    budget.finish_wire_scan(view.len())?;
    let mut locked = None;
    for field in view
        .fields()
        .filter(|field| field.number() == DRAWABLE_LOCKED_FIELD)
    {
        if locked.is_some() || field.wire_type() != 0 {
            return Err(ChartAxisTitleError::InvalidSource);
        }
        field.validate_canonical_framing().map_err(map_wire_error)?;
        let (value, bytes) = decode_varint_from_bytes(field.payload())
            .map_err(|_error| ChartAxisTitleError::InvalidSource)?;
        if bytes != field.payload().len() || value > 1 {
            return Err(ChartAxisTitleError::InvalidSource);
        }
        locked = Some(value != 0);
    }
    if locked == Some(true) {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    Ok(())
}

fn validate_axis_roles(
    category: &[u64],
    value: &[u64],
    budget: &mut AxisTitleBudget,
) -> Result<(), ChartAxisTitleError> {
    let category_primary = category.iter().copied().find(|identifier| *identifier != 0);
    let value_primary = value.iter().copied().find(|identifier| *identifier != 0);
    if category_primary.is_none() || value_primary.is_none() {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    let mut seen = HashSet::new();
    let roles = category
        .len()
        .checked_add(value.len())
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    budget.charge_reference_vector(roles)?;
    seen.try_reserve(roles)
        .map_err(|_error| ChartAxisTitleError::Allocation { amount: roles })?;
    for identifier in category.iter().chain(value).copied() {
        if identifier != 0 && !seen.insert(identifier) {
            return Err(ChartAxisTitleError::InvalidSource);
        }
    }
    Ok(())
}

fn unique_object(
    package: &Package,
    identifier: u64,
) -> Result<(&str, &ArchiveObject), ChartAxisTitleError> {
    let mut selected = None;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            if object.archive_info.identifier == Some(identifier) {
                if selected.replace((component.name(), object)).is_some() {
                    return Err(ChartAxisTitleError::InvalidSource);
                }
            }
        }
    }
    selected.ok_or(ChartAxisTitleError::InvalidSource)
}

fn validate_selected_message_metadata(
    object: &ArchiveObject,
    message_index: usize,
) -> Result<(), ChartAxisTitleError> {
    let message = object
        .messages
        .get(message_index)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let info = object
        .archive_info
        .message_infos
        .get(message_index)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    if object.archive_info.should_merge == Some(true)
        || info.type_ != message.type_
        || usize::try_from(info.length).ok() != Some(message.data.len())
        || info.base_message_index.is_some()
        || !info.diff_merge_version.is_empty()
        || info.diff_field_path.is_some()
        || !info.fields_to_remove.is_empty()
        || !info.diff_read_version.is_empty()
    {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    Ok(())
}

fn prove_unique_primary_axis(
    package: &Package,
    axis_identifier: u64,
    budget: &mut AxisTitleBudget,
) -> Result<(), ChartAxisTitleError> {
    budget.charge_scan_pass(package, 0)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let mut primary_owners = 0usize;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            for message in &object.messages {
                if message.type_ != CHART_MESSAGE_TYPE {
                    continue;
                }
                let outer = accounted_wire_fields(&message.data, limits, budget)?;
                let chart =
                    unique_length_delimited_field(&outer, &message.data, CHART_EXTENSION_FIELD)?
                        .ok_or(ChartAxisTitleError::InvalidSource)?;
                let category =
                    repeated_references(chart, CHART_AXIS_CATEGORY_FIELD, limits, budget)?;
                let value = repeated_references(chart, CHART_AXIS_VALUE_FIELD, limits, budget)?;
                validate_axis_roles(&category, &value, budget)?;
                for roles in [&category, &value] {
                    let primary = roles.iter().copied().find(|identifier| *identifier != 0);
                    if primary == Some(axis_identifier) {
                        primary_owners = primary_owners
                            .checked_add(1)
                            .ok_or(ChartAxisTitleError::InvalidSource)?;
                    } else if roles.contains(&axis_identifier) {
                        return Err(ChartAxisTitleError::InvalidSource);
                    }
                }
            }
        }
    }
    if primary_owners == 1 {
        Ok(())
    } else {
        Err(ChartAxisTitleError::InvalidSource)
    }
}

fn validate_global_axis_references<'a>(
    package: &'a Package,
    chart_identifier: u64,
    chart_message_index: usize,
    axis_identifier: u64,
    budget: &mut AxisTitleBudget,
) -> Result<Option<&'a str>, ChartAxisTitleError> {
    budget.charge_inbound_scan(package)?;
    let chart_object = package
        .object(chart_identifier)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let chart_info = chart_object
        .archive_info
        .message_infos
        .get(chart_message_index)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let aggregate_edges = chart_info
        .object_references
        .iter()
        .filter(|identifier| **identifier == axis_identifier)
        .count();
    let aggregate_data_edges = chart_info
        .data_references
        .iter()
        .filter(|identifier| **identifier == axis_identifier)
        .count();
    let mut field_edges = 0usize;
    let mut field_data_edges = 0usize;
    for field in &chart_info.field_infos {
        field_edges = field_edges
            .checked_add(
                field
                    .object_references
                    .iter()
                    .filter(|identifier| **identifier == axis_identifier)
                    .count(),
            )
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        field_data_edges = field_data_edges
            .checked_add(
                field
                    .data_references
                    .iter()
                    .filter(|identifier| **identifier == axis_identifier)
                    .count(),
            )
            .ok_or(ChartAxisTitleError::InvalidSource)?;
    }
    if aggregate_edges != 1
        || aggregate_data_edges != 0
        || field_edges != 0
        || field_data_edges != 0
    {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    let archive_limits = package
        .state
        .options
        .archive()
        .effective_archive_limits()
        .map_err(map_archive_error)?;
    let mut visitor = AxisInboundReferenceVisitor {
        package,
        chart_identifier,
        chart_message_index,
        axis_identifier,
        selected_references: 0,
        stylesheet_references: 0,
        stylesheet_component_name: None,
        invalid: false,
    };
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
    if visitor.invalid || visitor.selected_references != 1 {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    Ok(visitor.stylesheet_component_name)
}

struct AxisInboundReferenceVisitor<'a> {
    package: &'a Package,
    chart_identifier: u64,
    chart_message_index: usize,
    axis_identifier: u64,
    selected_references: usize,
    stylesheet_references: usize,
    stylesheet_component_name: Option<&'a str>,
    invalid: bool,
}

impl ArchiveReferenceVisitor for AxisInboundReferenceVisitor<'_> {
    fn visit_reference(
        &mut self,
        occurrence: ArchiveReferenceOccurrence,
    ) -> litchi_iwa_core::Result<()> {
        // Data references occupy a separate namespace, so unrelated chart
        // datasets do not participate in object ownership.  Still reject a
        // numeric collision with the selected axis: this focused owner does
        // not rewrite or prove that ambiguous cross-namespace relationship.
        if occurrence.kind == ArchiveReferenceKind::Data
            && occurrence.referenced_identifier == self.axis_identifier
        {
            self.invalid = true;
        }
        if occurrence.referenced_identifier == self.axis_identifier {
            if occurrence.kind == ArchiveReferenceKind::Object
                && occurrence.scope == ArchiveReferenceScope::Message
                && occurrence.object_identifier == self.chart_identifier
                && occurrence.message_index == self.chart_message_index
            {
                self.selected_references = self.selected_references.saturating_add(1);
            } else if let Some(component_name) =
                stylesheet_registration_component(self.package, occurrence)
            {
                self.stylesheet_references = self.stylesheet_references.saturating_add(1);
                self.stylesheet_component_name = Some(component_name);
                if self.stylesheet_references > 1 {
                    self.invalid = true;
                }
            } else {
                self.invalid = true;
            }
        }
        Ok(())
    }
}

fn stylesheet_registration_component(
    package: &Package,
    occurrence: ArchiveReferenceOccurrence,
) -> Option<&str> {
    if occurrence.kind != ArchiveReferenceKind::Object
        || occurrence.scope != ArchiveReferenceScope::Message
    {
        return None;
    }
    let (component, object) = package.object_with_component(occurrence.object_identifier)?;
    (object.messages.len() == 1
        && object
            .messages
            .get(occurrence.message_index)
            .is_some_and(|message| message.type_ == STYLESHEET_MESSAGE_TYPE)
        && validate_selected_message_metadata(object, occurrence.message_index).is_ok())
    .then_some(component)
}

fn validate_axis_metadata(
    package: &Package,
    chart_component_name: &str,
    axis_component_name: &str,
    stylesheet_component_name: Option<&str>,
    axis_identifier: u64,
    budget: &mut AxisTitleBudget,
) -> Result<(), ChartAxisTitleError> {
    budget.charge_scan_pass(package, 0)?;
    let Some(metadata) = package_metadata_payload(package)? else {
        return Ok(());
    };
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let remaining_input = budget
        .maximum_input
        .checked_sub(budget.input)
        .ok_or(ChartAxisTitleError::InvalidSource)?
        .min(limits.max_input_bytes());
    let remaining_output = budget
        .maximum_output
        .checked_sub(budget.output)
        .ok_or(ChartAxisTitleError::InvalidSource)?
        .min(limits.max_output_bytes());
    let remaining_fields = budget
        .maximum_fields
        .checked_sub(budget.fields)
        .ok_or(ChartAxisTitleError::InvalidSource)?
        .min(limits.max_fields());
    let remaining_work = budget
        .maximum_work
        .checked_sub(budget.work)
        .ok_or(ChartAxisTitleError::InvalidSource)?
        .min(limits.max_rewrite_work());
    let remaining_components = budget
        .maximum_components
        .checked_sub(budget.components)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let remaining_references = budget
        .maximum_references
        .checked_sub(budget.references)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let recursion =
        u32::try_from(limits.max_nesting()).map_err(|_error| ChartAxisTitleError::InvalidSource)?;
    let options = package_metadata_codec::RewriteOptions::new(
        remaining_input,
        remaining_output,
        remaining_fields,
        remaining_work,
        recursion,
        remaining_components,
        remaining_references,
        budget
            .maximum_allocations
            .checked_sub(budget.allocations)
            .ok_or(ChartAxisTitleError::InvalidSource)?,
    );
    let owner_external_component_name =
        (chart_component_name != axis_component_name).then_some(chart_component_name);
    let registry_external_component_name = stylesheet_component_name.filter(|component_name| {
        *component_name != axis_component_name
            && Some(*component_name) != owner_external_component_name
    });
    let mut selected = SelectedAxisMetadataVisitor {
        axis_component_name,
        owner_external_component_name,
        registry_external_component_name,
        axis_identifier,
        selected_uuid: None,
        selected_component_identifier: None,
        external_target_component_identifier: None,
        owner_external_seen: false,
        registry_external_seen: false,
        count: 0,
        invalid: false,
    };
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        metadata,
        options,
        &mut selected,
    )
    .map_err(map_metadata_error)?;
    budget.charge_metadata_report(inspection.report())?;
    let expects_external =
        owner_external_component_name.is_some() || registry_external_component_name.is_some();
    if selected.owner_external_seen != owner_external_component_name.is_some()
        || selected.registry_external_seen != registry_external_component_name.is_some()
        || (expects_external
            && selected.external_target_component_identifier
                != selected.selected_component_identifier)
        || (!expects_external && selected.external_target_component_identifier.is_some())
    {
        selected.invalid = true;
    }
    let selected_uuid = selected
        .selected_uuid
        .filter(|_uuid| !selected.invalid && selected.count == 1)
        .ok_or(ChartAxisTitleError::InvalidSource)?;

    let options = metadata_options(package, budget)?;
    let mut authority = AxisMetadataAuthorityVisitor {
        axis_identifier,
        selected_uuid,
        selected_pair_count: 0,
        selected_object_count: 0,
        invalid: false,
    };
    let inspection = package_metadata_codec::inspect_package_metadata_with_visitor(
        metadata,
        options,
        &mut authority,
    )
    .map_err(map_metadata_error)?;
    budget.charge_metadata_report(inspection.report())?;
    if authority.invalid
        || authority.selected_pair_count != 1
        || authority.selected_object_count != 1
    {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    Ok(())
}

fn metadata_options(
    package: &Package,
    budget: &AxisTitleBudget,
) -> Result<package_metadata_codec::RewriteOptions, ChartAxisTitleError> {
    let limits = package.wire_limits().map_err(map_wire_error)?;
    Ok(package_metadata_codec::RewriteOptions::new(
        budget
            .maximum_input
            .checked_sub(budget.input)
            .ok_or(ChartAxisTitleError::InvalidSource)?
            .min(limits.max_input_bytes()),
        budget
            .maximum_output
            .checked_sub(budget.output)
            .ok_or(ChartAxisTitleError::InvalidSource)?
            .min(limits.max_output_bytes()),
        budget
            .maximum_fields
            .checked_sub(budget.fields)
            .ok_or(ChartAxisTitleError::InvalidSource)?
            .min(limits.max_fields()),
        budget
            .maximum_work
            .checked_sub(budget.work)
            .ok_or(ChartAxisTitleError::InvalidSource)?
            .min(limits.max_rewrite_work()),
        u32::try_from(limits.max_nesting()).map_err(|_error| ChartAxisTitleError::InvalidSource)?,
        budget
            .maximum_components
            .checked_sub(budget.components)
            .ok_or(ChartAxisTitleError::InvalidSource)?,
        budget
            .maximum_references
            .checked_sub(budget.references)
            .ok_or(ChartAxisTitleError::InvalidSource)?,
        budget
            .maximum_allocations
            .checked_sub(budget.allocations)
            .ok_or(ChartAxisTitleError::InvalidSource)?,
    ))
}

fn package_metadata_payload(package: &Package) -> Result<Option<&[u8]>, ChartAxisTitleError> {
    let mut payload = None;
    for component in package.state.source.components().iter() {
        for object in &component.archive().objects {
            if object.messages.len() != object.archive_info.message_infos.len() {
                return Err(ChartAxisTitleError::InvalidSource);
            }
            for (index, message) in object.messages.iter().enumerate() {
                let info = object
                    .archive_info
                    .message_infos
                    .get(index)
                    .ok_or(ChartAxisTitleError::InvalidSource)?;
                if message.type_ != info.type_
                    || usize::try_from(info.length).ok() != Some(message.data.len())
                {
                    return Err(ChartAxisTitleError::InvalidSource);
                }
                if message.type_ == PACKAGE_METADATA_MESSAGE_TYPE {
                    validate_selected_message_metadata(object, index)?;
                    if payload.replace(message.data.as_slice()).is_some() {
                        return Err(ChartAxisTitleError::InvalidSource);
                    }
                }
            }
        }
    }
    Ok(payload)
}

fn metadata_component_matches_physical(
    component: package_metadata_codec::ComponentDescriptor<'_>,
    physical: &str,
) -> bool {
    let Some(expected) = physical
        .strip_prefix("Index/")
        .and_then(|value| value.strip_suffix(".iwa"))
    else {
        return false;
    };
    component.preferred_locator() == expected
        && component
            .locator()
            .is_none_or(|locator| locator == expected)
        && component.effective_locator() == expected
}

struct SelectedAxisMetadataVisitor<'a> {
    axis_component_name: &'a str,
    owner_external_component_name: Option<&'a str>,
    registry_external_component_name: Option<&'a str>,
    axis_identifier: u64,
    selected_uuid: Option<package_metadata_codec::UuidBits>,
    selected_component_identifier: Option<u64>,
    external_target_component_identifier: Option<u64>,
    owner_external_seen: bool,
    registry_external_seen: bool,
    count: usize,
    invalid: bool,
}

impl package_metadata_codec::PackageMetadataVisitor for SelectedAxisMetadataVisitor<'_> {
    fn visit_object_uuid(
        &mut self,
        binding: package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if binding.object_identifier() == self.axis_identifier {
            self.count = self.count.saturating_add(1);
            if self.selected_uuid.replace(binding.uuid()).is_some() {
                self.invalid = true;
            }
            if self
                .selected_component_identifier
                .replace(binding.component().identifier())
                .is_some()
            {
                self.invalid = true;
            }
            if !binding.component().is_current()
                || !metadata_component_matches_physical(
                    binding.component(),
                    self.axis_component_name,
                )
            {
                self.invalid = true;
            }
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: package_metadata_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if reference.object_identifier() == Some(self.axis_identifier) {
            let target = reference.target_component_identifier();
            if self
                .external_target_component_identifier
                .is_some_and(|identifier| identifier != target)
            {
                self.invalid = true;
            } else {
                self.external_target_component_identifier = Some(target);
            }
            let source_matches = |component_name: &str| {
                reference.source().is_current()
                    && metadata_component_matches_physical(reference.source(), component_name)
            };
            if self
                .owner_external_component_name
                .is_some_and(source_matches)
            {
                if self.owner_external_seen {
                    self.invalid = true;
                }
                self.owner_external_seen = true;
            } else if self
                .registry_external_component_name
                .is_some_and(source_matches)
            {
                if self.registry_external_seen {
                    self.invalid = true;
                }
                self.registry_external_seen = true;
            } else {
                self.invalid = true;
            }
            if reference.is_weak() == Some(true) || reference.is_versioned() {
                self.invalid = true;
            }
        }
        Ok(())
    }

    fn visit_data_reference_owner(
        &mut self,
        owner: package_metadata_codec::DataReferenceOwnerDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if owner.object_identifier() == self.axis_identifier {
            self.invalid = true;
        }
        Ok(())
    }

    fn visit_ambiguous_object_identifier(
        &mut self,
        _component: package_metadata_codec::ComponentDescriptor<'_>,
        identifier: u64,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if identifier == self.axis_identifier {
            self.invalid = true;
        }
        Ok(())
    }

    fn visit_data_metadata_map(
        &mut self,
        object_identifier: u64,
        _has_unknown_fields: bool,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if object_identifier == self.axis_identifier {
            self.invalid = true;
        }
        Ok(())
    }
}

struct AxisMetadataAuthorityVisitor {
    axis_identifier: u64,
    selected_uuid: package_metadata_codec::UuidBits,
    selected_pair_count: usize,
    selected_object_count: usize,
    invalid: bool,
}

impl package_metadata_codec::PackageMetadataVisitor for AxisMetadataAuthorityVisitor {
    fn visit_object_uuid(
        &mut self,
        binding: package_metadata_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), package_metadata_codec::RewriteError> {
        if binding.uuid() == self.selected_uuid {
            self.selected_pair_count = self.selected_pair_count.saturating_add(1);
            if binding.object_identifier() != self.axis_identifier {
                self.invalid = true;
            }
        }
        if binding.object_identifier() == self.axis_identifier {
            self.selected_object_count = self.selected_object_count.saturating_add(1);
            if binding.uuid() != self.selected_uuid {
                self.invalid = true;
            }
        }
        Ok(())
    }
}

fn rewrite_chart_axis_title(
    source: &Package,
    selection: &AxisSelection,
    after: Option<&str>,
    budget: &mut AxisTitleBudget,
) -> Result<(Package, usize), ChartAxisTitleError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.axis_component_name)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(ChartAxisTitleError::InvalidSource);
    }
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
    let parsed_component = source
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selection.axis_component_name)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let decompressed_bound = parsed_component
        .archive()
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget.charge_native_decompress_parse(
        entry.data().len(),
        decompressed_bound,
        parsed_component.archive().objects.len(),
    )?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    validate_canonical_object_length_prefixes(stream.as_bytes(), &archive)?;
    let message_data = {
        let object = archive
            .object(selection.axis_identifier)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        let message = object
            .messages
            .get(selection.axis_message_index)
            .filter(|message| message.type_ == CHART_AXIS_MESSAGE_TYPE)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        budget.charge_owned_clone(message.data.len())?;
        let mut cloned = Vec::new();
        cloned
            .try_reserve_exact(message.data.len())
            .map_err(|_error| ChartAxisTitleError::Allocation {
                amount: message.data.len(),
            })?;
        cloned.extend_from_slice(&message.data);
        cloned
    };
    let limits = source.wire_limits().map_err(map_wire_error)?;
    let patched = patch_chart_axis_title(
        message_data.as_slice(),
        after,
        limits,
        axis_title_kind(selection.axis),
        budget,
    )?;
    archive
        .object_mut(selection.axis_identifier)
        .ok_or(ChartAxisTitleError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.axis_message_index,
            RawMessage {
                type_: CHART_AXIS_MESSAGE_TYPE,
                data: patched,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    budget.charge_archive_snappy_plan(&archive, archive_limits, snappy_limits)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_error| ChartAxisTitleError::InvalidSource)?;
    let edit = EntryEdit::new(
        selection.axis_component_name.as_str(),
        compressed.as_slice(),
    );
    let edits = [edit];
    let prepared_reassembly = catalog
        .prepare_reassembly_with_deletions(&edits, previews.names(), source.state.options.archive())
        .map_err(map_archive_error)?;
    let reassembly_requirements = prepared_reassembly.execution_requirements();
    budget.charge_candidate_reopen(
        source,
        reassembly_requirements.output_bytes(),
        Some((selection.axis_component_name.as_str(), &archive)),
    )?;
    let reassembly_limits = budget.charge_reassembly(reassembly_requirements)?;
    let output = litchi_iwa_archive::package::PreparedReassembly::execute(
        prepared_reassembly,
        reassembly_limits,
    )
    .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok((candidate, previews.len()))
}

fn validate_canonical_object_length_prefixes(
    source: &[u8],
    archive: &Archive,
) -> Result<(), ChartAxisTitleError> {
    for object in &archive.objects {
        let offset = usize::try_from(object.header_offset)
            .map_err(|_error| ChartAxisTitleError::InvalidSource)?;
        let remaining = source
            .get(offset..)
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        let (header_bytes, prefix_bytes) = decode_varint_from_bytes(remaining)
            .map_err(|_error| ChartAxisTitleError::InvalidSource)?;
        if prefix_bytes != encoded_len(header_bytes) {
            return Err(ChartAxisTitleError::InvalidSource);
        }
        let framed_header_bytes = header_bytes
            .checked_add(
                u64::try_from(prefix_bytes).map_err(|_error| ChartAxisTitleError::InvalidSource)?,
            )
            .ok_or(ChartAxisTitleError::InvalidSource)?;
        if framed_header_bytes != object.header_length
            || object
                .header_offset
                .checked_add(object.header_length)
                .ok_or(ChartAxisTitleError::InvalidSource)?
                != object.data_offset
        {
            return Err(ChartAxisTitleError::InvalidSource);
        }
    }
    Ok(())
}

fn read_chart_axis_title(
    data: &[u8],
    limits: WireLimits,
    kind: AxisTitleKind,
    budget: &mut AxisTitleBudget,
) -> Result<Option<String>, ChartAxisTitleError> {
    let fields = accounted_wire_fields(data, limits, budget)?;
    let mut extension_field = None;
    let mut extension_count = 0usize;
    for field in &fields {
        if field.number() == GENERATED_CHART_AXIS_EXTENSION_FIELD {
            extension_count = extension_count
                .checked_add(1)
                .ok_or(ChartAxisTitleError::InvalidSource)?;
            extension_field = Some(*field);
        }
    }
    let Some(field) = extension_field else {
        return Ok(None);
    };
    if extension_count != 1 || field.wire_type() != 2 {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    field
        .validate_canonical_framing(data)
        .map_err(map_wire_error)?;
    let extension = field.payload(data).map_err(map_wire_error)?;
    let (snapshot, report) =
        decode_axis_titles_with_report(extension, budget.codec_options(extension.len())?)
            .map_err(map_chart_axis_title_codec_error)?;
    budget.charge_decode(report)?;
    snapshot.visible_title(kind).map(copy_title).transpose()
}

fn verify_candidate_reopen_locality(
    source: &Package,
    candidate: &Package,
    slide_position: Position,
    chart_position: Position,
    axis: Axis,
    expected_title: Option<&str>,
    target_requires_invalidated_previews: bool,
    budget: &mut AxisTitleBudget,
) -> Result<(), ChartAxisTitleError> {
    if source.state.total_objects != candidate.state.total_objects {
        return Err(ChartAxisTitleError::Verification);
    }
    let source_selected = select_axis(
        source,
        SlideSelector::position(slide_position),
        ChartSelector::index(chart_position.get()),
        axis,
        true,
        budget,
    )?;
    let candidate_selected = select_axis(
        candidate,
        SlideSelector::position(slide_position),
        ChartSelector::index(chart_position.get()),
        axis,
        true,
        budget,
    )?;
    if source_selected.axis_identifier != candidate_selected.axis_identifier
        || source_selected.chart_identifier != candidate_selected.chart_identifier
        || source_selected.axis_component_name != candidate_selected.axis_component_name
        || candidate_selected.title.as_deref() != expected_title
    {
        return Err(ChartAxisTitleError::Verification);
    }
    let source_show = source.show().map_err(map_read_error)?;
    let candidate_show = candidate.show().map_err(map_read_error)?;
    if source_show != candidate_show {
        return Err(ChartAxisTitleError::Verification);
    }
    verify_package_locality(
        source,
        candidate,
        &source_selected,
        target_requires_invalidated_previews,
        budget,
    )?;
    if target_requires_invalidated_previews
        && !super::rendering_invalidation::root_previews_absent(
            physical_catalog(candidate)?.package(),
        )
        .map_err(|_error| ChartAxisTitleError::Verification)?
    {
        return Err(ChartAxisTitleError::Verification);
    }
    Ok(())
}

fn verify_package_locality(
    source: &Package,
    candidate: &Package,
    selection: &AxisSelection,
    previews_must_be_absent: bool,
    budget: &mut AxisTitleBudget,
) -> Result<(), ChartAxisTitleError> {
    budget.charge_locality_scan(source)?;
    budget.charge_locality_scan(candidate)?;
    let source_catalog = physical_catalog(source)?.package();
    let candidate_catalog = physical_catalog(candidate)?.package();
    let source_previews = super::rendering_invalidation::root_preview_deletions(source_catalog)
        .map_err(|_error| ChartAxisTitleError::Verification)?;
    let candidate_previews =
        super::rendering_invalidation::root_preview_deletions(candidate_catalog)
            .map_err(|_error| ChartAxisTitleError::Verification)?;
    budget.charge(
        source_catalog
            .len()
            .checked_add(candidate_catalog.len())
            .ok_or(ChartAxisTitleError::InvalidSource)?,
    )?;
    for entry in source_catalog.iter() {
        if source_previews.names().contains(&entry.name()) {
            continue;
        }
        let other = candidate_catalog
            .iter()
            .find(|candidate_entry| candidate_entry.name() == entry.name())
            .ok_or(ChartAxisTitleError::Verification)?;
        let selected_component = entry.name() == selection.axis_component_name;
        if (!selected_component && entry.data() != other.data())
            || entry.raw_name() != other.raw_name()
            || entry.is_opaque() != other.is_opaque()
            // Reassembly necessarily updates CRC and compressed/uncompressed
            // sizes for the selected member.  The local/central header
            // metadata itself must remain byte-equivalent, though; this
            // catches a metadata rewrite without rejecting an expected
            // payload-size update.
            || entry.metadata().local() != other.metadata().local()
            || entry.metadata().central() != other.metadata().central()
            || (!selected_component && entry.metadata() != other.metadata())
            // Central-directory offsets may legitimately move after a
            // rewritten member, but an untouched member's local record is
            // expected to remain byte-identical.
            || (!selected_component
                && entry.raw_record().local_record() != other.raw_record().local_record())
        {
            return Err(ChartAxisTitleError::Verification);
        }
    }
    for entry in candidate_catalog.iter() {
        if candidate_previews.names().contains(&entry.name()) {
            continue;
        }
        if source_catalog
            .iter()
            .all(|source_entry| source_entry.name() != entry.name())
        {
            return Err(ChartAxisTitleError::Verification);
        }
    }
    if previews_must_be_absent && !candidate_previews.names().is_empty() {
        return Err(ChartAxisTitleError::Verification);
    }

    let source_component = source
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selection.axis_component_name)
        .ok_or(ChartAxisTitleError::Verification)?;
    let candidate_component = candidate
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selection.axis_component_name)
        .ok_or(ChartAxisTitleError::Verification)?;
    let source_objects = &source_component.archive().objects;
    let candidate_objects = &candidate_component.archive().objects;
    budget.charge(
        source_objects
            .len()
            .checked_add(candidate_objects.len())
            .ok_or(ChartAxisTitleError::InvalidSource)?,
    )?;
    if source_objects.len() != candidate_objects.len() {
        return Err(ChartAxisTitleError::Verification);
    }
    for (source_object, candidate_object) in source_objects.iter().zip(candidate_objects) {
        if source_object.archive_info.identifier != candidate_object.archive_info.identifier {
            return Err(ChartAxisTitleError::Verification);
        }
        if source_object.archive_info.identifier != Some(selection.axis_identifier)
            && (source_object.archive_info != candidate_object.archive_info
                || source_object.messages != candidate_object.messages)
        {
            return Err(ChartAxisTitleError::Verification);
        }
    }
    Ok(())
}

fn patch_chart_axis_title(
    data: &[u8],
    title: Option<&str>,
    limits: WireLimits,
    kind: AxisTitleKind,
    budget: &mut AxisTitleBudget,
) -> Result<Vec<u8>, ChartAxisTitleError> {
    let fields = accounted_wire_fields(data, limits, budget)?;
    let extension_field = fields
        .iter()
        .copied()
        .find(|field| field.number() == GENERATED_CHART_AXIS_EXTENSION_FIELD);
    if fields
        .iter()
        .filter(|field| field.number() == GENERATED_CHART_AXIS_EXTENSION_FIELD)
        .count()
        > 1
    {
        return Err(ChartAxisTitleError::InvalidSource);
    }
    if let Some(field) = extension_field {
        if field.wire_type() != 2 {
            return Err(ChartAxisTitleError::InvalidSource);
        }
        field
            .validate_canonical_framing(data)
            .map_err(map_wire_error)?;
    }
    let Some(extension) = extension_field
        .map(|field| field.payload(data).map_err(map_wire_error))
        .transpose()?
    else {
        let Some(title) = title else {
            return clone_title_source(data, limits, budget);
        };
        let extension = execute_prepared_axis_title_rewrite(
            &[],
            AxisTitleWrite::preserve().with_title(kind, Some(title)),
            budget,
        )?;
        return rewrite_chart_non_style_extension(data, None, &extension, limits, budget);
    };

    let extension = execute_prepared_axis_title_rewrite(
        extension,
        AxisTitleWrite::preserve().with_title(kind, title),
        budget,
    )?;
    rewrite_chart_non_style_extension(data, extension_field, &extension, limits, budget)
}

fn execute_prepared_axis_title_rewrite(
    source: &[u8],
    write: AxisTitleWrite<'_>,
    budget: &mut AxisTitleBudget,
) -> Result<Vec<u8>, ChartAxisTitleError> {
    let prepared = prepare_axis_title_rewrite(source, write, budget.codec_options(source.len())?)
        .map_err(map_chart_axis_title_codec_error)?;
    let prepare_report = prepared.prepare_report();
    let requirements = prepared.execution_requirements();
    budget.charge_prepared(prepare_report, requirements)?;
    prepared
        .execute(requirements.exact())
        .map_err(map_chart_axis_title_codec_error)
        .map(|output| output.into_output())
}

const fn axis_title_kind(axis: Axis) -> AxisTitleKind {
    match axis {
        Axis::Category => AxisTitleKind::Category,
        Axis::Value => AxisTitleKind::Value,
    }
}

fn rewrite_chart_non_style_extension(
    data: &[u8],
    extension_field: Option<WireField>,
    replacement: &[u8],
    limits: WireLimits,
    budget: &mut AxisTitleBudget,
) -> Result<Vec<u8>, ChartAxisTitleError> {
    let replacement_length =
        u64::try_from(replacement.len()).map_err(|_error| ChartAxisTitleError::InvalidSource)?;
    let key_length = extension_field.map_or_else(
        || encoded_len((u64::from(GENERATED_CHART_AXIS_EXTENSION_FIELD) << 3) | 2),
        |field| field.key_end() - field.start(),
    );
    let replacement_field_length = key_length
        .checked_add(encoded_len(replacement_length))
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let output_length = extension_field
        .map_or_else(
            || data.len().checked_add(replacement_field_length),
            |field| {
                data.len()
                    .checked_sub(field.end() - field.start())
                    .and_then(|length| length.checked_add(replacement_field_length))
            },
        )
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    if output_length > limits.max_output_bytes() {
        return Err(ChartAxisTitleError::LimitExceeded {
            kind: ChartAxisTitleLimitKind::OutputBytes,
            observed: usize_to_u64(output_length),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
    budget.charge_enclosing_buffer(output_length)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_error| ChartAxisTitleError::Allocation {
            amount: output_length,
        })?;
    match extension_field {
        Some(field) => {
            output.extend_from_slice(&data[..field.start()]);
            output.extend_from_slice(&data[field.start()..field.key_end()]);
            encode_varint_into(&mut output, replacement_length);
            output.extend_from_slice(replacement);
            output.extend_from_slice(&data[field.end()..]);
        },
        None => {
            output.extend_from_slice(data);
            encode_varint_into(
                &mut output,
                (u64::from(GENERATED_CHART_AXIS_EXTENSION_FIELD) << 3) | 2,
            );
            encode_varint_into(&mut output, replacement_length);
            output.extend_from_slice(replacement);
        },
    }
    debug_assert_eq!(output.len(), output_length);
    Ok(output)
}

fn clone_title_source(
    data: &[u8],
    limits: WireLimits,
    budget: &mut AxisTitleBudget,
) -> Result<Vec<u8>, ChartAxisTitleError> {
    if data.len() > limits.max_output_bytes() {
        return Err(ChartAxisTitleError::LimitExceeded {
            kind: ChartAxisTitleLimitKind::OutputBytes,
            observed: usize_to_u64(data.len()),
            maximum: usize_to_u64(limits.max_output_bytes()),
        });
    }
    budget.charge_enclosing_buffer(data.len())?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(data.len())
        .map_err(|_error| ChartAxisTitleError::Allocation { amount: data.len() })?;
    output.extend_from_slice(data);
    Ok(output)
}

fn copy_title(title: &str) -> Result<String, ChartAxisTitleError> {
    if title.len() > MAX_CHART_TITLE_BYTES {
        return Err(ChartAxisTitleError::LimitExceeded {
            kind: ChartAxisTitleLimitKind::TitleBytes,
            observed: usize_to_u64(title.len()),
            maximum: usize_to_u64(MAX_CHART_TITLE_BYTES),
        });
    }
    let mut owned = String::new();
    owned
        .try_reserve_exact(title.len())
        .map_err(|_error| ChartAxisTitleError::Allocation {
            amount: title.len(),
        })?;
    owned.push_str(title);
    Ok(owned)
}

fn map_chart_axis_title_codec_error(error: ChartAxisTitleDecodeError) -> ChartAxisTitleError {
    if let Some(limit) = error.resource_limit() {
        return match limit {
            ChartAxisTitleDecodeLimit::Bytes { observed, maximum } => {
                ChartAxisTitleError::LimitExceeded {
                    kind: ChartAxisTitleLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartAxisTitleDecodeLimit::Fields { observed, maximum } => {
                ChartAxisTitleError::LimitExceeded {
                    kind: ChartAxisTitleLimitKind::WireFields,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartAxisTitleDecodeLimit::Work { observed, maximum } => {
                ChartAxisTitleError::LimitExceeded {
                    kind: ChartAxisTitleLimitKind::WireWork,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartAxisTitleDecodeLimit::Output { observed, maximum } => {
                ChartAxisTitleError::LimitExceeded {
                    kind: ChartAxisTitleLimitKind::OutputBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartAxisTitleDecodeLimit::Title { observed, maximum } => {
                ChartAxisTitleError::LimitExceeded {
                    kind: ChartAxisTitleLimitKind::TitleBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartAxisTitleDecodeLimit::Nesting { observed, maximum } => {
                ChartAxisTitleError::LimitExceeded {
                    kind: ChartAxisTitleLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            ChartAxisTitleDecodeLimit::Allocations { observed, maximum } => {
                ChartAxisTitleError::LimitExceeded {
                    kind: ChartAxisTitleLimitKind::Entries,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartAxisTitleDecodeLimit::Retained { observed, maximum }
            | ChartAxisTitleDecodeLimit::Scratch { observed, maximum } => {
                ChartAxisTitleError::LimitExceeded {
                    kind: ChartAxisTitleLimitKind::TotalBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            _ => ChartAxisTitleError::InvalidSource,
        };
    }
    if let Some(limit) = error.wire_resource_limit() {
        return match limit {
            ChartAxisTitleWireResourceLimit::Bytes { observed, maximum } => {
                ChartAxisTitleError::LimitExceeded {
                    kind: ChartAxisTitleLimitKind::WireBytes,
                    observed: usize_to_u64(observed),
                    maximum: usize_to_u64(maximum),
                }
            },
            ChartAxisTitleWireResourceLimit::Nesting { observed, maximum } => {
                ChartAxisTitleError::LimitExceeded {
                    kind: ChartAxisTitleLimitKind::WireNesting,
                    observed: u64::from(observed),
                    maximum: u64::from(maximum),
                }
            },
            _ => ChartAxisTitleError::InvalidSource,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return ChartAxisTitleError::Allocation { amount };
    }
    ChartAxisTitleError::InvalidSource
}

fn map_chart_title_error(error: super::slide_chart_title::ChartTitleError) -> ChartAxisTitleError {
    use super::slide_chart_title::{ChartTitleError, ChartTitleLimitKind};
    match error {
        ChartTitleError::UnsupportedSource => ChartAxisTitleError::UnsupportedSource,
        ChartTitleError::AmbiguousSelector => ChartAxisTitleError::AmbiguousSelector,
        ChartTitleError::EmptySlideName => ChartAxisTitleError::EmptySlideName,
        ChartTitleError::SlideNameNotFound => ChartAxisTitleError::SlideNameNotFound,
        ChartTitleError::SlidePositionNotFound { position } => {
            ChartAxisTitleError::SlidePositionNotFound { position }
        },
        ChartTitleError::ChartNameNotFound => ChartAxisTitleError::ChartNameNotFound,
        ChartTitleError::ChartPositionNotFound { position } => {
            ChartAxisTitleError::ChartPositionNotFound { position }
        },
        ChartTitleError::EmptyChartName => ChartAxisTitleError::EmptyChartName,
        ChartTitleError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => ChartAxisTitleError::LimitExceeded {
            kind: match kind {
                ChartTitleLimitKind::InputBytes => ChartAxisTitleLimitKind::InputBytes,
                ChartTitleLimitKind::OutputBytes => ChartAxisTitleLimitKind::OutputBytes,
                ChartTitleLimitKind::WireBytes => ChartAxisTitleLimitKind::WireBytes,
                ChartTitleLimitKind::Entries => ChartAxisTitleLimitKind::Entries,
                ChartTitleLimitKind::EntryBytes => ChartAxisTitleLimitKind::EntryBytes,
                ChartTitleLimitKind::TotalBytes => ChartAxisTitleLimitKind::TotalBytes,
                ChartTitleLimitKind::Slides => ChartAxisTitleLimitKind::Slides,
                ChartTitleLimitKind::References => ChartAxisTitleLimitKind::References,
                ChartTitleLimitKind::TextStorages => ChartAxisTitleLimitKind::TextStorages,
                ChartTitleLimitKind::TextFragments => ChartAxisTitleLimitKind::TextFragments,
                ChartTitleLimitKind::TextBytes => ChartAxisTitleLimitKind::TextBytes,
                ChartTitleLimitKind::WireFields => ChartAxisTitleLimitKind::WireFields,
                ChartTitleLimitKind::WireNesting => ChartAxisTitleLimitKind::WireNesting,
                ChartTitleLimitKind::WireWork => ChartAxisTitleLimitKind::WireWork,
                ChartTitleLimitKind::TitleBytes => ChartAxisTitleLimitKind::TitleBytes,
            },
            observed,
            maximum,
        },
        ChartTitleError::Allocation { amount } => ChartAxisTitleError::Allocation { amount },
        ChartTitleError::InvalidSource
        | ChartTitleError::Verification
        | ChartTitleError::PatchConflict => ChartAxisTitleError::InvalidSource,
    }
}

fn map_metadata_error(error: package_metadata_codec::RewriteError) -> ChartAxisTitleError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            package_metadata_codec::RewriteLimit::InputBytes { observed, maximum } => {
                (ChartAxisTitleLimitKind::WireBytes, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::OutputBytes { observed, maximum } => {
                (ChartAxisTitleLimitKind::OutputBytes, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Fields { observed, maximum } => {
                (ChartAxisTitleLimitKind::WireFields, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Work { observed, maximum } => {
                (ChartAxisTitleLimitKind::WireWork, observed, maximum)
            },
            package_metadata_codec::RewriteLimit::Nesting { observed, maximum } => (
                ChartAxisTitleLimitKind::WireNesting,
                observed as usize,
                maximum as usize,
            ),
            package_metadata_codec::RewriteLimit::Components { observed, maximum }
            | package_metadata_codec::RewriteLimit::References { observed, maximum }
            | package_metadata_codec::RewriteLimit::Additions { observed, maximum } => {
                (ChartAxisTitleLimitKind::References, observed, maximum)
            },
            _ => return ChartAxisTitleError::InvalidSource,
        };
        return ChartAxisTitleError::LimitExceeded {
            kind,
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        };
    }
    if let Some(amount) = error.allocation_request() {
        return ChartAxisTitleError::Allocation { amount };
    }
    ChartAxisTitleError::InvalidSource
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, ChartAxisTitleError> {
    match &package.state.source {
        PhysicalSource::Package(source) => Ok(source),
        PhysicalSource::Semantic(_) => Err(ChartAxisTitleError::UnsupportedSource),
    }
}

fn map_read_error(error: ReadError) -> ChartAxisTitleError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartAxisTitleError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => ChartAxisTitleLimitKind::Entries,
                SemanticLimitKind::Slides => ChartAxisTitleLimitKind::Slides,
                SemanticLimitKind::References => ChartAxisTitleLimitKind::References,
                SemanticLimitKind::TextStorages => ChartAxisTitleLimitKind::TextStorages,
                SemanticLimitKind::TextFragments => ChartAxisTitleLimitKind::TextFragments,
                SemanticLimitKind::TextBytes => ChartAxisTitleLimitKind::TextBytes,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartAxisTitleError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => ChartAxisTitleLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => ChartAxisTitleLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => ChartAxisTitleLimitKind::WireNesting,
                super::PayloadLimitKind::Work => ChartAxisTitleLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => ChartAxisTitleError::Allocation { amount },
        ReadError::Archive(_)
        | ReadError::Detection(_)
        | ReadError::NotKeynote
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_)
        | ReadError::Io(_) => ChartAxisTitleError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> ChartAxisTitleError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartAxisTitleError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => ChartAxisTitleLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => ChartAxisTitleLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => ChartAxisTitleLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes => ChartAxisTitleLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => ChartAxisTitleLimitKind::TotalBytes,
                _ => ChartAxisTitleLimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            ChartAxisTitleError::Allocation { amount }
        },
        _ => ChartAxisTitleError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> ChartAxisTitleError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartAxisTitleError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes => ChartAxisTitleLimitKind::TotalBytes,
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => ChartAxisTitleLimitKind::Entries,
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    ChartAxisTitleLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::HeaderFields => ChartAxisTitleLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => ChartAxisTitleLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyFrames => ChartAxisTitleLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            ChartAxisTitleError::Allocation { amount: requested }
        },
        _ => ChartAxisTitleError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> ChartAxisTitleError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => ChartAxisTitleError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => ChartAxisTitleLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => ChartAxisTitleLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    ChartAxisTitleLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => ChartAxisTitleLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => ChartAxisTitleLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            ChartAxisTitleError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => ChartAxisTitleError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}
