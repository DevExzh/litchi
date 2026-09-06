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

use std::fmt;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::package::{EntryEdit, ExactArtifacts};
use litchi_iwa_common::{WireLimits, encode_varint_into, varint::encoded_len, wire::WireField};
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::keynote_chart_axis_title_codec::{
    AxisTitleKind, AxisTitleWrite, DecodeError as ChartAxisTitleDecodeError,
    DecodeLimit as ChartAxisTitleDecodeLimit, DecodeOptions, DecodeReport,
    RewriteExecutionRequirements, WireResourceLimit as ChartAxisTitleWireResourceLimit,
    decode_axis_titles_with_report, prepare_axis_title_rewrite,
};
use litchi_iwa_protos::package_metadata_codec;
use thiserror::Error;

use super::chart_axis_support::{self, AxisSelection, AxisSupportBudget, AxisSupportError};
use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::{Axis, ChartSelector, SlideSelector};

const CHART_AXIS_MESSAGE_TYPE: u32 = 5_027;
const GENERATED_CHART_AXIS_EXTENSION_FIELD: u32 = 10_000;
const MAX_CHART_TITLE_BYTES: usize = 64 * 1024 * 1024;

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

    fn charge_locality_scan(&mut self, package: &Package) -> Result<(), ChartAxisTitleError> {
        // Locality compares the physical catalog and the selected component's
        // complete object/message inventory without constructing a new index.
        let inventory = package_scan_inventory(package)?;
        let entries = chart_axis_support::physical_catalog(package)
            .map_err(map_axis_support_error)?
            .package()
            .len();
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

impl AxisSupportBudget for AxisTitleBudget {
    fn charge_selection_scans(
        &mut self,
        package: &Package,
        mutation_guards: bool,
    ) -> Result<(), AxisSupportError> {
        AxisTitleBudget::charge_selection_scans(self, package, mutation_guards)
            .map_err(axis_support_budget_error)
    }

    fn charge_input(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        AxisTitleBudget::charge_input(self, amount).map_err(axis_support_budget_error)
    }

    fn charge_wire_vector(&mut self, payload: usize) -> Result<(), AxisSupportError> {
        AxisTitleBudget::charge_wire_vector(self, payload).map_err(axis_support_budget_error)
    }

    fn finish_wire_scan(&mut self, fields: usize) -> Result<(), AxisSupportError> {
        AxisTitleBudget::finish_wire_scan(self, fields).map_err(axis_support_budget_error)
    }

    fn charge_reference_vector(&mut self, capacity: usize) -> Result<(), AxisSupportError> {
        AxisTitleBudget::charge_reference_vector(self, capacity).map_err(axis_support_budget_error)
    }

    fn charge_references(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        AxisTitleBudget::charge_references(self, amount).map_err(axis_support_budget_error)
    }

    fn charge_scan_pass(
        &mut self,
        package: &Package,
        retained_vectors: usize,
    ) -> Result<(), AxisSupportError> {
        AxisTitleBudget::charge_scan_pass(self, package, retained_vectors)
            .map_err(axis_support_budget_error)
    }

    fn charge_locality_scan(&mut self, package: &Package) -> Result<(), AxisSupportError> {
        AxisTitleBudget::charge_locality_scan(self, package).map_err(axis_support_budget_error)
    }

    fn charge_work(&mut self, amount: usize) -> Result<(), AxisSupportError> {
        AxisTitleBudget::charge(self, amount).map_err(axis_support_budget_error)
    }

    fn metadata_options(
        &self,
        package: &Package,
    ) -> Result<package_metadata_codec::RewriteOptions, AxisSupportError> {
        metadata_options(package, self).map_err(axis_support_budget_error)
    }

    fn charge_metadata_report(
        &mut self,
        report: package_metadata_codec::RewriteReport,
    ) -> Result<(), AxisSupportError> {
        AxisTitleBudget::charge_metadata_report(self, report).map_err(axis_support_budget_error)
    }
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
        u32::try_from(limits.max_nesting()).map_err(|_| ChartAxisTitleError::InvalidSource)?,
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
    let catalog = chart_axis_support::physical_catalog(source)
        .map_err(map_axis_support_error)?
        .package();
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
        let selection = chart_axis_support::select_axis(
            source,
            slide_selector.into(),
            chart_selector.into(),
            axis,
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        let before = read_selected_axis_title(source, &selection, &mut budget)?;
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
        let catalog =
            chart_axis_support::physical_catalog(self.source).map_err(map_axis_support_error)?;
        let source_bytes = catalog.shared_source();
        let mut budget = AxisTitleBudget::new(self.source)?;
        budget.charge_catalog_scan(self.source)?;
        let source_selection = chart_axis_support::select_axis(
            self.source,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            self.axis,
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        let source_title = read_selected_axis_title(self.source, &source_selection, &mut budget)?;
        if source_selection.chart_identifier != self.chart_identifier
            || source_selection.axis_identifier != self.axis_identifier
            || source_selection.axis_component_name != self.axis_component_name
            || source_selection.axis_message_index != self.axis_message_index
            || source_selection.slide_identifier != self.slide_identifier
            || source_title != self.before
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
                    restored_previews: 0,
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
        let target = chart_axis_support::physical_catalog(&package)
            .map_err(map_axis_support_error)?
            .shared_source();
        budget.charge_exact_artifacts(source_bytes.len(), target.len())?;
        let candidate_selection = chart_axis_support::select_axis(
            &package,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            self.axis,
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        let candidate_title =
            read_selected_axis_title(&package, &candidate_selection, &mut budget)?;
        if candidate_selection.chart_identifier != self.chart_identifier
            || candidate_selection.axis_identifier != self.axis_identifier
            || candidate_selection.axis_component_name != self.axis_component_name
            || candidate_selection.axis_message_index != self.axis_message_index
            || candidate_selection.slide_identifier != self.slide_identifier
            || candidate_title != self.after
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
                restored_previews: 0,
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
    restored_previews: usize,
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
            deleted_previews: self.restored_previews,
            restored_previews: self.deleted_previews,
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
        let selection = chart_axis_support::select_axis(
            self,
            slide_selector.into(),
            chart_selector.into(),
            axis,
            false,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        read_selected_axis_title(self, &selection, &mut budget)
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
        let catalog = chart_axis_support::physical_catalog(self).map_err(map_axis_support_error)?;
        let source = catalog.shared_source();
        let mut budget = AxisTitleBudget::new(self)?;
        budget.charge_catalog_scan(self)?;
        if !patch.artifacts.authorizes_source(&source) {
            return Err(ChartAxisTitleError::PatchConflict);
        }
        let source_selection = chart_axis_support::select_axis(
            self,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            patch.axis,
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        let source_title = read_selected_axis_title(self, &source_selection, &mut budget)?;
        if source_selection.chart_identifier != patch.chart_identifier
            || source_selection.axis_identifier != patch.axis_identifier
            || source_selection.axis_component_name != patch.axis_component_name
            || source_selection.axis_message_index != patch.axis_message_index
            || source_selection.slide_identifier != patch.slide_identifier
            || source_title != patch.before
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
        let candidate_selection = chart_axis_support::select_axis(
            &candidate,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            patch.axis,
            true,
            &mut budget,
        )
        .map_err(map_axis_support_error)?;
        let candidate_title =
            read_selected_axis_title(&candidate, &candidate_selection, &mut budget)?;
        if candidate_selection.chart_identifier != patch.chart_identifier
            || candidate_selection.axis_identifier != patch.axis_identifier
            || candidate_selection.axis_component_name != patch.axis_component_name
            || candidate_selection.axis_message_index != patch.axis_message_index
            || candidate_selection.slide_identifier != patch.slide_identifier
            || candidate_title != patch.after
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

fn rewrite_chart_axis_title(
    source: &Package,
    selection: &AxisSelection,
    after: Option<&str>,
    budget: &mut AxisTitleBudget,
) -> Result<(Package, usize), ChartAxisTitleError> {
    let catalog = chart_axis_support::physical_catalog(source).map_err(map_axis_support_error)?;
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
    chart_axis_support::validate_canonical_object_length_prefixes_with_budget(
        stream.as_bytes(),
        &archive,
        budget,
    )
    .map_err(map_axis_support_error)?;
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

fn read_selected_axis_title(
    package: &Package,
    selection: &AxisSelection,
    budget: &mut AxisTitleBudget,
) -> Result<Option<String>, ChartAxisTitleError> {
    let (_component, object) = package
        .object_with_component(selection.axis_identifier)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let message = object
        .messages
        .get(selection.axis_message_index)
        .filter(|message| message.type_ == CHART_AXIS_MESSAGE_TYPE)
        .ok_or(ChartAxisTitleError::InvalidSource)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    read_chart_axis_title(
        &message.data,
        limits,
        axis_title_kind(selection.axis),
        budget,
    )
}

fn accounted_wire_fields(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut AxisTitleBudget,
) -> Result<Vec<WireField>, ChartAxisTitleError> {
    chart_axis_support::accounted_wire_fields(payload, limits, budget)
        .map_err(map_axis_support_error)
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
    let source_selected = chart_axis_support::select_axis(
        source,
        SlideSelector::position(slide_position),
        ChartSelector::index(chart_position.get()),
        axis,
        true,
        budget,
    )
    .map_err(map_axis_support_error)?;
    let candidate_selected = chart_axis_support::select_axis(
        candidate,
        SlideSelector::position(slide_position),
        ChartSelector::index(chart_position.get()),
        axis,
        true,
        budget,
    )
    .map_err(map_axis_support_error)?;
    let candidate_title = read_selected_axis_title(candidate, &candidate_selected, budget)?;
    if source_selected.axis_identifier != candidate_selected.axis_identifier
        || source_selected.chart_identifier != candidate_selected.chart_identifier
        || source_selected.axis_component_name != candidate_selected.axis_component_name
        || candidate_title.as_deref() != expected_title
    {
        return Err(ChartAxisTitleError::Verification);
    }
    let source_show = source.show().map_err(map_read_error)?;
    let candidate_show = candidate.show().map_err(map_read_error)?;
    if source_show != candidate_show {
        return Err(ChartAxisTitleError::Verification);
    }
    chart_axis_support::verify_package_locality(
        source,
        candidate,
        &source_selected,
        target_requires_invalidated_previews,
        budget,
    )
    .map_err(map_axis_support_error)?;
    if target_requires_invalidated_previews
        && !super::rendering_invalidation::root_previews_absent(
            chart_axis_support::physical_catalog(candidate)
                .map_err(map_axis_support_error)?
                .package(),
        )
        .map_err(|_error| ChartAxisTitleError::Verification)?
    {
        return Err(ChartAxisTitleError::Verification);
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

fn axis_support_budget_error(error: ChartAxisTitleError) -> AxisSupportError {
    match error {
        ChartAxisTitleError::UnsupportedSource => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::UnsupportedSource,
        ),
        ChartAxisTitleError::AmbiguousSelector => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::AmbiguousSelector,
        ),
        ChartAxisTitleError::EmptySlideName => {
            AxisSupportError::Selector(chart_axis_support::AxisSupportSelectorError::EmptySlideName)
        },
        ChartAxisTitleError::SlideNameNotFound => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::SlideNameNotFound,
        ),
        ChartAxisTitleError::SlidePositionNotFound { position } => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::SlidePositionNotFound { position },
        ),
        ChartAxisTitleError::ChartNameNotFound => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::ChartNameNotFound,
        ),
        ChartAxisTitleError::ChartPositionNotFound { position } => AxisSupportError::Selector(
            chart_axis_support::AxisSupportSelectorError::ChartPositionNotFound { position },
        ),
        ChartAxisTitleError::EmptyChartName => {
            AxisSupportError::Selector(chart_axis_support::AxisSupportSelectorError::EmptyChartName)
        },
        ChartAxisTitleError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => AxisSupportError::LimitExceeded {
            kind: match kind {
                ChartAxisTitleLimitKind::InputBytes => {
                    chart_axis_support::AxisSupportLimitKind::InputBytes
                },
                ChartAxisTitleLimitKind::OutputBytes => {
                    chart_axis_support::AxisSupportLimitKind::OutputBytes
                },
                ChartAxisTitleLimitKind::WireBytes => {
                    chart_axis_support::AxisSupportLimitKind::WireBytes
                },
                ChartAxisTitleLimitKind::Entries => {
                    chart_axis_support::AxisSupportLimitKind::Entries
                },
                ChartAxisTitleLimitKind::EntryBytes => {
                    chart_axis_support::AxisSupportLimitKind::EntryBytes
                },
                ChartAxisTitleLimitKind::TotalBytes => {
                    chart_axis_support::AxisSupportLimitKind::TotalBytes
                },
                ChartAxisTitleLimitKind::Slides => chart_axis_support::AxisSupportLimitKind::Slides,
                ChartAxisTitleLimitKind::References => {
                    chart_axis_support::AxisSupportLimitKind::References
                },
                ChartAxisTitleLimitKind::TextStorages => {
                    chart_axis_support::AxisSupportLimitKind::TextStorages
                },
                ChartAxisTitleLimitKind::TextFragments => {
                    chart_axis_support::AxisSupportLimitKind::TextFragments
                },
                ChartAxisTitleLimitKind::TextBytes => {
                    chart_axis_support::AxisSupportLimitKind::TextBytes
                },
                ChartAxisTitleLimitKind::WireFields => {
                    chart_axis_support::AxisSupportLimitKind::WireFields
                },
                ChartAxisTitleLimitKind::WireNesting => {
                    chart_axis_support::AxisSupportLimitKind::WireNesting
                },
                ChartAxisTitleLimitKind::WireWork => {
                    chart_axis_support::AxisSupportLimitKind::WireWork
                },
                ChartAxisTitleLimitKind::TitleBytes => {
                    chart_axis_support::AxisSupportLimitKind::TitleBytes
                },
            },
            observed,
            maximum,
        },
        ChartAxisTitleError::Allocation { amount } => AxisSupportError::Allocation { amount },
        ChartAxisTitleError::InvalidSource
        | ChartAxisTitleError::Verification
        | ChartAxisTitleError::PatchConflict => AxisSupportError::InvalidSource,
    }
}

fn map_axis_support_error(error: AxisSupportError) -> ChartAxisTitleError {
    match error {
        AxisSupportError::Selector(selector) => match selector {
            chart_axis_support::AxisSupportSelectorError::UnsupportedSource => {
                ChartAxisTitleError::UnsupportedSource
            },
            chart_axis_support::AxisSupportSelectorError::AmbiguousSelector => {
                ChartAxisTitleError::AmbiguousSelector
            },
            chart_axis_support::AxisSupportSelectorError::EmptySlideName => {
                ChartAxisTitleError::EmptySlideName
            },
            chart_axis_support::AxisSupportSelectorError::SlideNameNotFound => {
                ChartAxisTitleError::SlideNameNotFound
            },
            chart_axis_support::AxisSupportSelectorError::SlidePositionNotFound { position } => {
                ChartAxisTitleError::SlidePositionNotFound { position }
            },
            chart_axis_support::AxisSupportSelectorError::ChartNameNotFound => {
                ChartAxisTitleError::ChartNameNotFound
            },
            chart_axis_support::AxisSupportSelectorError::ChartPositionNotFound { position } => {
                ChartAxisTitleError::ChartPositionNotFound { position }
            },
            chart_axis_support::AxisSupportSelectorError::EmptyChartName => {
                ChartAxisTitleError::EmptyChartName
            },
        },
        AxisSupportError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => ChartAxisTitleError::LimitExceeded {
            kind: match kind {
                chart_axis_support::AxisSupportLimitKind::InputBytes => {
                    ChartAxisTitleLimitKind::InputBytes
                },
                chart_axis_support::AxisSupportLimitKind::OutputBytes => {
                    ChartAxisTitleLimitKind::OutputBytes
                },
                chart_axis_support::AxisSupportLimitKind::WireBytes => {
                    ChartAxisTitleLimitKind::WireBytes
                },
                chart_axis_support::AxisSupportLimitKind::Entries => {
                    ChartAxisTitleLimitKind::Entries
                },
                chart_axis_support::AxisSupportLimitKind::EntryBytes => {
                    ChartAxisTitleLimitKind::EntryBytes
                },
                chart_axis_support::AxisSupportLimitKind::TotalBytes => {
                    ChartAxisTitleLimitKind::TotalBytes
                },
                chart_axis_support::AxisSupportLimitKind::Slides => ChartAxisTitleLimitKind::Slides,
                chart_axis_support::AxisSupportLimitKind::References => {
                    ChartAxisTitleLimitKind::References
                },
                chart_axis_support::AxisSupportLimitKind::TextStorages => {
                    ChartAxisTitleLimitKind::TextStorages
                },
                chart_axis_support::AxisSupportLimitKind::TextFragments => {
                    ChartAxisTitleLimitKind::TextFragments
                },
                chart_axis_support::AxisSupportLimitKind::TextBytes => {
                    ChartAxisTitleLimitKind::TextBytes
                },
                chart_axis_support::AxisSupportLimitKind::WireFields => {
                    ChartAxisTitleLimitKind::WireFields
                },
                chart_axis_support::AxisSupportLimitKind::WireNesting => {
                    ChartAxisTitleLimitKind::WireNesting
                },
                chart_axis_support::AxisSupportLimitKind::WireWork => {
                    ChartAxisTitleLimitKind::WireWork
                },
                chart_axis_support::AxisSupportLimitKind::TitleBytes => {
                    ChartAxisTitleLimitKind::TitleBytes
                },
            },
            observed,
            maximum,
        },
        AxisSupportError::Allocation { amount } => ChartAxisTitleError::Allocation { amount },
        AxisSupportError::InvalidSource => ChartAxisTitleError::InvalidSource,
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
