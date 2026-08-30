//! Exact-source, selector-first transactions for Keynote value-axis settings.
//!
//! Bounds, step counts, and scale are deliberately edited as one aggregate.
//! The chart graph is proved by [`super::chart_axis_support`], while the
//! generated value-axis extension is decoded and rewritten by the private
//! strict codec in `litchi_iwa_protos`.  No native identifier, archive object,
//! or protobuf value crosses this module's public API.

#![allow(
    clippy::cast_sign_loss,
    clippy::map_err_ignore,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    reason = "The transaction redacts lower-layer failures and keeps native graph adapters private."
)]

use std::sync::Arc;
use std::{fmt, mem::size_of};

use litchi_core::Position;
use litchi_iwa_archive::{
    SourceCatalog,
    package::{EntryEdit, ExactArtifacts},
};
use litchi_iwa_common::{WireLimits, encode_varint_into, varint::encoded_len, wire::WireField};
use litchi_iwa_core::{
    Archive, ArchiveInfo, ArchiveObject, FieldInfo, FieldPath, MessageInfo, RawMessage,
    SnappyLimits, SnappyStream,
};
use litchi_iwa_protos::keynote_chart_axis_value_settings_codec::{
    AxisValueBound as CodecBound, AxisValueBounds as CodecBounds,
    AxisValueSettingsSnapshot as CodecSnapshot, AxisValueSettingsWrite as CodecWrite,
    AxisValueSteps as CodecSteps, DecodeError as CodecError, DecodeLimit as CodecLimit,
    DecodeOptions as CodecOptions, DecodeReport as CodecReport,
    RewriteExecutionRequirements as CodecRequirements, Scale as CodecScale,
    decode_axis_value_settings_with_report, prepare_axis_value_settings_rewrite,
};
use thiserror::Error;

use super::{Package, PhysicalSource, ReadError, SemanticLimitKind};
use crate::chart::axis::{
    Bound, Bounds, MajorStepCount, MinorStepCount, Scale, Steps, ValueAxisSettings,
};
use crate::{ChartSelector, SlideSelector};

const CHART_AXIS_MESSAGE_TYPE: u32 = 5_027;
const GENERATED_CHART_AXIS_EXTENSION_FIELD: u32 = 10_000;
const MAX_VALUE_AXIS_BYTES: usize = 64 * 1024 * 1024;
// A transaction must remain finite even when callers configure the general
// archive limits to very large values.  The aggregate budget below is derived
// from the exact source size and then capped by these hard ceilings.
const MAX_VALUE_AXIS_TRANSACTION_BYTES: usize = 512 * 1024 * 1024;
const MAX_VALUE_AXIS_TRANSACTION_ITEMS: usize = 16 * 1024 * 1024;
const MAX_VALUE_AXIS_TRANSACTION_MEMORY: usize = 1024 * 1024 * 1024;
const VALUE_AXIS_TRANSACTION_FIXED_OVERHEAD: usize = 256 * 1024;
const VALUE_AXIS_TRANSACTION_FIXED_ALLOCATIONS: usize = 4 * 1024;
const VALUE_AXIS_TRANSACTION_ALLOCATION_MULTIPLIER: usize = 16;
const VALUE_AXIS_TRANSACTION_MEMORY_MULTIPLIER: usize = 256;
const SNAPPY_FRAME_HEADER_BYTES: usize = 4;

/// Resource categories charged by one aggregate value-axis transaction.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum ChartValueAxisLimitKind {
    /// Complete input package bytes.
    InputBytes,
    /// Complete rewritten package or payload bytes.
    OutputBytes,
    /// Bytes in one protobuf payload.
    WireBytes,
    /// ZIP members, IWA objects, or IWA messages.
    Entries,
    /// Bytes in one package member or IWA object.
    EntryBytes,
    /// Aggregate package or IWA bytes.
    TotalBytes,
    /// Semantic slides.
    Slides,
    /// Semantic graph references.
    References,
    /// Parsed protobuf fields.
    WireFields,
    /// Protobuf nesting depth.
    WireNesting,
    /// Aggregate protobuf work.
    WireWork,
}

impl fmt::Display for ChartValueAxisLimitKind {
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
            Self::WireFields => "wire fields",
            Self::WireNesting => "wire nesting",
            Self::WireWork => "wire work",
        })
    }
}

/// A content-redacted failure raised by a value-axis transaction.
#[derive(Debug, Clone, PartialEq, Eq, Error)]
#[non_exhaustive]
pub enum ChartValueAxisError {
    /// The source was prepared without an exact physical package artifact.
    #[error("this Keynote source does not support physical chart value-axis edits")]
    UnsupportedSource,
    /// A selector matched more than one semantic item.
    #[error("the Keynote chart value-axis selector is ambiguous")]
    AmbiguousSelector,
    /// An exact-name slide selector was empty.
    #[error("the Keynote slide selector name cannot be empty")]
    EmptySlideName,
    /// An exact-name slide selector did not match.
    #[error("the Keynote show has no slide matching the requested name")]
    SlideNameNotFound,
    /// A checked semantic slide position does not exist.
    #[error("the Keynote show has no slide at position {position:?}")]
    SlidePositionNotFound { position: Position },
    /// An exact-name chart selector did not match.
    #[error("the selected Keynote slide has no chart matching the requested name")]
    ChartNameNotFound,
    /// A checked semantic chart position does not exist.
    #[error("the selected Keynote slide has no chart at position {position:?}")]
    ChartPositionNotFound { position: Position },
    /// An empty exact chart name was supplied.
    #[error("the Keynote chart selector name cannot be empty")]
    EmptyChartName,
    /// The semantic aggregate is unsuitable for native Keynote settings.
    #[error("the Keynote chart value-axis settings are invalid")]
    InvalidSettings,
    /// The source graph or selected native payload is malformed.
    #[error("the Keynote chart value-axis source cannot be edited safely")]
    InvalidSource,
    /// A finite resource ceiling was exceeded.
    #[error(
        "Keynote chart value-axis settings {kind} limit exceeded: observed {observed}, maximum {maximum}"
    )]
    LimitExceeded {
        /// Resource category that exceeded its ceiling.
        kind: ChartValueAxisLimitKind,
        /// Observed or requested amount.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A bounded destination allocation failed before publication.
    #[error("could not allocate {amount} units for the Keynote chart value-axis transaction")]
    Allocation { amount: usize },
    /// Full candidate reopening did not reproduce the requested settings.
    #[error("the edited Keynote chart value-axis settings failed semantic verification")]
    Verification,
    /// The patch does not belong to this exact immutable package artifact.
    #[error("the Keynote chart value-axis patch does not match the exact source package")]
    PatchConflict,
}

/// A checked envelope for heap-owned vectors used by the lower layers.
///
/// This intentionally describes logical capacities and allocation events. It
/// is not an allocator or RSS measurement: allocator rounding, hash-table
/// load factors, and the allocator's private bookkeeping remain outside this
/// facade. The envelope is nevertheless charged before an operation can ask
/// the lower layer to reserve any of these buffers.
#[derive(Debug, Clone, Copy, Default)]
struct HeapFootprint {
    allocations: usize,
    bytes: usize,
    scratch: usize,
}

impl HeapFootprint {
    fn add_counter(current: &mut usize, amount: usize) -> Result<(), ChartValueAxisError> {
        *current = current
            .checked_add(amount)
            .ok_or(ChartValueAxisError::InvalidSource)?;
        Ok(())
    }

    fn add_vector<T>(&mut self, capacity: usize, scratch: bool) -> Result<(), ChartValueAxisError> {
        if capacity == 0 {
            return Ok(());
        }
        let bytes = capacity
            .checked_mul(size_of::<T>())
            .ok_or(ChartValueAxisError::InvalidSource)?;
        Self::add_counter(&mut self.allocations, 1)?;
        if scratch {
            Self::add_counter(&mut self.scratch, bytes)?;
        } else {
            Self::add_counter(&mut self.bytes, bytes)?;
        }
        Ok(())
    }

    fn add_bytes(&mut self, bytes: usize, scratch: bool) -> Result<(), ChartValueAxisError> {
        if bytes == 0 {
            return Ok(());
        }
        Self::add_counter(&mut self.allocations, 1)?;
        if scratch {
            Self::add_counter(&mut self.scratch, bytes)?;
        } else {
            Self::add_counter(&mut self.bytes, bytes)?;
        }
        Ok(())
    }

    fn merge(&mut self, other: Self) -> Result<(), ChartValueAxisError> {
        Self::add_counter(&mut self.allocations, other.allocations)?;
        Self::add_counter(&mut self.bytes, other.bytes)?;
        Self::add_counter(&mut self.scratch, other.scratch)
    }
}

/// Metadata vectors cloned by `Archive::to_bytes_with_limits` and by the
/// compatibility conversion used by the IWA header encoder.
///
/// `use_capacity` is used for a source-owned clone. Conversion targets use
/// their source lengths because the lower layer calls `try_reserve_exact` for
/// those vectors. Every arithmetic edge is checked so a hostile metadata
/// topology fails before any lower-layer allocation.
fn archive_info_heap(
    info: &ArchiveInfo,
    use_capacity: bool,
    scratch: bool,
) -> Result<HeapFootprint, ChartValueAxisError> {
    let mut footprint = HeapFootprint::default();
    let message_capacity = if use_capacity {
        info.message_infos.capacity()
    } else {
        info.message_infos.len()
    };
    footprint.add_vector::<MessageInfo>(message_capacity, scratch)?;
    for message in &info.message_infos {
        footprint.merge(message_info_heap(message, use_capacity, scratch)?)?;
    }
    Ok(footprint)
}

fn message_info_heap(
    info: &MessageInfo,
    use_capacity: bool,
    scratch: bool,
) -> Result<HeapFootprint, ChartValueAxisError> {
    let mut footprint = HeapFootprint::default();
    let capacity = |length: usize, source_capacity: usize| {
        if use_capacity {
            source_capacity
        } else {
            length
        }
    };
    footprint.add_vector::<u32>(
        capacity(info.versions.len(), info.versions.capacity()),
        scratch,
    )?;
    footprint.add_vector::<FieldInfo>(
        capacity(info.field_infos.len(), info.field_infos.capacity()),
        scratch,
    )?;
    footprint.add_vector::<u64>(
        capacity(
            info.object_references.len(),
            info.object_references.capacity(),
        ),
        scratch,
    )?;
    footprint.add_vector::<u64>(
        capacity(info.data_references.len(), info.data_references.capacity()),
        scratch,
    )?;
    footprint.add_vector::<u32>(
        capacity(
            info.diff_merge_version.len(),
            info.diff_merge_version.capacity(),
        ),
        scratch,
    )?;
    if let Some(path) = &info.diff_field_path {
        footprint.add_vector::<u32>(capacity(path.path.len(), path.path.capacity()), scratch)?;
    }
    footprint.add_vector::<FieldPath>(
        capacity(
            info.fields_to_remove.len(),
            info.fields_to_remove.capacity(),
        ),
        scratch,
    )?;
    for path in &info.fields_to_remove {
        footprint.add_vector::<u32>(capacity(path.path.len(), path.path.capacity()), scratch)?;
    }
    footprint.add_vector::<u32>(
        capacity(
            info.diff_read_version.len(),
            info.diff_read_version.capacity(),
        ),
        scratch,
    )?;
    for field in &info.field_infos {
        footprint.merge(field_info_heap(field, use_capacity, scratch)?)?;
    }
    Ok(footprint)
}

fn field_info_heap(
    info: &FieldInfo,
    use_capacity: bool,
    scratch: bool,
) -> Result<HeapFootprint, ChartValueAxisError> {
    let mut footprint = HeapFootprint::default();
    let capacity = |length: usize, source_capacity: usize| {
        if use_capacity {
            source_capacity
        } else {
            length
        }
    };
    footprint.add_vector::<u32>(
        capacity(info.path.path.len(), info.path.path.capacity()),
        scratch,
    )?;
    footprint.add_vector::<u64>(
        capacity(
            info.object_references.len(),
            info.object_references.capacity(),
        ),
        scratch,
    )?;
    footprint.add_vector::<u64>(
        capacity(info.data_references.len(), info.data_references.capacity()),
        scratch,
    )?;
    footprint.add_vector::<u32>(
        capacity(
            info.known_field_version.len(),
            info.known_field_version.capacity(),
        ),
        scratch,
    )?;
    if let Some(identifier) = &info.known_field_feature_identifier {
        let identifier_capacity = if use_capacity {
            identifier.capacity()
        } else {
            identifier.len()
        };
        footprint.add_vector::<u8>(identifier_capacity, scratch)?;
    }
    Ok(footprint)
}

/// Return checked metadata work for one archive. This is deliberately based
/// on every vector length rather than only object/message counts so late
/// metadata mismatches cannot spend uncharged comparison work.
fn archive_metadata_work(archive: &Archive) -> Result<usize, ChartValueAxisError> {
    let mut work = archive.objects.len();
    for object in &archive.objects {
        work = work
            .checked_add(object.messages.len())
            .and_then(|value| value.checked_add(object.archive_info.message_infos.len()))
            .ok_or(ChartValueAxisError::InvalidSource)?;
        for info in &object.archive_info.message_infos {
            work = work
                .checked_add(info.versions.len())
                .and_then(|value| value.checked_add(info.field_infos.len()))
                .and_then(|value| value.checked_add(info.object_references.len()))
                .and_then(|value| value.checked_add(info.data_references.len()))
                .and_then(|value| value.checked_add(info.diff_merge_version.len()))
                .and_then(|value| value.checked_add(info.diff_read_version.len()))
                .and_then(|value| value.checked_add(info.fields_to_remove.len()))
                .ok_or(ChartValueAxisError::InvalidSource)?;
            if let Some(path) = &info.diff_field_path {
                work = work
                    .checked_add(path.path.len())
                    .ok_or(ChartValueAxisError::InvalidSource)?;
            }
            for path in &info.fields_to_remove {
                work = work
                    .checked_add(path.path.len())
                    .ok_or(ChartValueAxisError::InvalidSource)?;
            }
            for field in &info.field_infos {
                work = work
                    .checked_add(field.path.path.len())
                    .and_then(|value| value.checked_add(field.object_references.len()))
                    .and_then(|value| value.checked_add(field.data_references.len()))
                    .and_then(|value| value.checked_add(field.known_field_version.len()))
                    .and_then(|value| {
                        field
                            .known_field_feature_identifier
                            .as_ref()
                            .map_or(Some(value), |feature| value.checked_add(feature.len()))
                    })
                    .ok_or(ChartValueAxisError::InvalidSource)?;
            }
        }
    }
    Ok(work)
}

fn archive_metadata_work_for_object(object: &ArchiveObject) -> Result<usize, ChartValueAxisError> {
    let mut work = object
        .messages
        .len()
        .checked_add(object.archive_info.message_infos.len())
        .ok_or(ChartValueAxisError::InvalidSource)?;
    for info in &object.archive_info.message_infos {
        work = work
            .checked_add(info.versions.len())
            .and_then(|value| value.checked_add(info.field_infos.len()))
            .and_then(|value| value.checked_add(info.object_references.len()))
            .and_then(|value| value.checked_add(info.data_references.len()))
            .and_then(|value| value.checked_add(info.diff_merge_version.len()))
            .and_then(|value| value.checked_add(info.diff_read_version.len()))
            .and_then(|value| value.checked_add(info.fields_to_remove.len()))
            .ok_or(ChartValueAxisError::InvalidSource)?;
        if let Some(path) = &info.diff_field_path {
            work = work
                .checked_add(path.path.len())
                .ok_or(ChartValueAxisError::InvalidSource)?;
        }
        for path in &info.fields_to_remove {
            work = work
                .checked_add(path.path.len())
                .ok_or(ChartValueAxisError::InvalidSource)?;
        }
        for field in &info.field_infos {
            work = work
                .checked_add(field.path.path.len())
                .and_then(|value| value.checked_add(field.object_references.len()))
                .and_then(|value| value.checked_add(field.data_references.len()))
                .and_then(|value| value.checked_add(field.known_field_version.len()))
                .and_then(|value| {
                    field
                        .known_field_feature_identifier
                        .as_ref()
                        .map_or(Some(value), |feature| value.checked_add(feature.len()))
                })
                .ok_or(ChartValueAxisError::InvalidSource)?;
        }
    }
    Ok(work)
}

/// Compute the bounded output/scratch shape of the lower-layer Snappy writer.
/// The writer emits one independent frame per `WRITE_CHUNK_SIZE` bytes, so a
/// maliciously large (or merely highly fragmented) archive cannot hide a
/// frame-vector allocation behind one aggregate compressed-byte charge.
fn snappy_output_plan(
    input_len: usize,
    limits: SnappyLimits,
) -> Result<(usize, usize, usize), ChartValueAxisError> {
    if input_len > limits.max_decompressed_stream() {
        return Err(ChartValueAxisError::LimitExceeded {
            kind: ChartValueAxisLimitKind::EntryBytes,
            observed: usize_to_u64(input_len),
            maximum: usize_to_u64(limits.max_decompressed_stream()),
        });
    }
    let frame_count = input_len
        .checked_add(SnappyStream::WRITE_CHUNK_SIZE - 1)
        .map_or(0, |length| length / SnappyStream::WRITE_CHUNK_SIZE);
    if frame_count > limits.max_frames() {
        return Err(ChartValueAxisError::LimitExceeded {
            kind: ChartValueAxisLimitKind::WireWork,
            observed: usize_to_u64(frame_count),
            maximum: usize_to_u64(limits.max_frames()),
        });
    }

    let mut remaining = input_len;
    let mut compressed = 0usize;
    let mut maximum_frame = 0usize;
    while remaining != 0 {
        let chunk = remaining.min(SnappyStream::WRITE_CHUNK_SIZE);
        if chunk > limits.max_uncompressed_chunk() {
            return Err(ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::EntryBytes,
                observed: usize_to_u64(chunk),
                maximum: usize_to_u64(limits.max_uncompressed_chunk()),
            });
        }
        let frame = SnappyStream::maximum_compressed_len(chunk).map_err(map_core_error)?;
        let compressed_chunk = frame
            .checked_sub(SNAPPY_FRAME_HEADER_BYTES)
            .ok_or(ChartValueAxisError::InvalidSource)?;
        if compressed_chunk > limits.max_compressed_chunk() {
            return Err(ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::EntryBytes,
                observed: usize_to_u64(compressed_chunk),
                maximum: usize_to_u64(limits.max_compressed_chunk()),
            });
        }
        compressed = compressed
            .checked_add(frame)
            .ok_or(ChartValueAxisError::InvalidSource)?;
        maximum_frame = maximum_frame.max(compressed_chunk);
        remaining -= chunk;
    }
    if compressed > limits.max_compressed_stream() {
        return Err(ChartValueAxisError::LimitExceeded {
            kind: ChartValueAxisLimitKind::EntryBytes,
            observed: usize_to_u64(compressed),
            maximum: usize_to_u64(limits.max_compressed_stream()),
        });
    }
    Ok((frame_count, compressed, maximum_frame))
}

/// Estimate allocations made while parsing one already-known archive shape.
/// The source archive is the topology authority for a candidate reopen; the
/// candidate is later required to retain these counts by locality checks.
fn archive_reopen_heap(archive: &Archive) -> Result<HeapFootprint, ChartValueAxisError> {
    let mut footprint = HeapFootprint::default();
    footprint.add_vector::<ArchiveObject>(archive.objects.capacity(), false)?;
    for object in &archive.objects {
        footprint.add_vector::<RawMessage>(object.messages.capacity(), false)?;
        for message in &object.messages {
            footprint.add_vector::<u8>(message.data.capacity(), false)?;
        }

        // Archive::parse first materializes a bounded compatibility header,
        // then converts it into the neutral metadata below. Both the parsed
        // metadata and the temporary compatibility vectors are charged.
        footprint.merge(archive_info_heap(&object.archive_info, true, false)?)?;
        footprint.merge(archive_info_heap(&object.archive_info, false, true)?)?;
        let header = usize::try_from(object.header_length)
            .map_err(|_| ChartValueAxisError::InvalidSource)?;
        // Raw and canonical header copies are retained only for a
        // non-canonical source, but charging both is the safe envelope for a
        // candidate whose source bytes are not known until reopen.
        footprint.add_bytes(header, false)?;
        footprint.add_bytes(header, true)?;
    }
    Ok(footprint)
}

/// Estimate `Archive::to_bytes_with_limits`' per-object metadata work and the
/// multi-frame Snappy output it feeds. The output-free encoded length check is
/// intentionally excluded; this helper covers only allocation-bearing work.
fn archive_to_bytes_heap(
    archive: &Archive,
    encoded: usize,
    snappy_limits: SnappyLimits,
) -> Result<(HeapFootprint, usize), ChartValueAxisError> {
    let (frames, compressed, maximum_frame) = snappy_output_plan(encoded, snappy_limits)?;
    let mut footprint = HeapFootprint::default();

    // Archive output remains alive while Snappy consumes it.
    footprint.add_bytes(encoded, false)?;
    for object in &archive.objects {
        // `validate_object_set` performs one compatibility conversion, and
        // the per-object encoder performs another after cloning ArchiveInfo.
        footprint.merge(archive_info_heap(&object.archive_info, false, true)?)?;
        footprint.merge(archive_info_heap(&object.archive_info, true, false)?)?;
        footprint.merge(archive_info_heap(&object.archive_info, false, true)?)?;
        let header = usize::try_from(object.header_length)
            .map_err(|_| ChartValueAxisError::InvalidSource)?;
        footprint.add_bytes(header, true)?;
    }

    // SnappyStream::compress reserves the complete maximum output once and
    // `compress_vec` creates one temporary compressed vector per frame.
    footprint.add_bytes(compressed, false)?;
    for _ in 0..frames {
        footprint.add_bytes(maximum_frame, true)?;
    }
    Ok((footprint, frames))
}

/// A single aggregate finite ledger shared by graph selection, codec rewrite,
/// reassembly, candidate reopening, and locality verification.
struct ValueAxisBudget {
    limits: WireLimits,
    maximum_input: usize,
    maximum_output: usize,
    maximum_fields: usize,
    maximum_work: usize,
    maximum_references: usize,
    maximum_allocations: usize,
    maximum_retained: usize,
    maximum_scratch: usize,
    input: usize,
    output: usize,
    fields: usize,
    work: usize,
    max_depth: u32,
    references: usize,
    allocations: usize,
    retained: usize,
    scratch: usize,
}

impl ValueAxisBudget {
    fn new(package: &Package) -> Result<Self, ChartValueAxisError> {
        let limits = package.wire_limits().map_err(map_wire_error)?;
        let configured_input = limits.max_input_bytes();
        let source_bytes = match &package.state.source {
            PhysicalSource::Package(source) => source.shared_source().len(),
            PhysicalSource::Semantic(_) => 0,
        };
        if source_bytes > MAX_VALUE_AXIS_TRANSACTION_BYTES {
            return Err(ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::InputBytes,
                observed: usize_to_u64(source_bytes),
                maximum: usize_to_u64(MAX_VALUE_AXIS_TRANSACTION_BYTES),
            });
        }
        // Semantic package snapshots have no ZIP artifact to measure.  Use
        // the checked wire ceiling as their finite basis while physical
        // snapshots are sized from the exact immutable source bytes.
        let basis = if source_bytes == 0 {
            configured_input.max(1)
        } else {
            source_bytes
        };
        let baseline = basis.min(configured_input.max(1));
        let aggregate = baseline
            .checked_mul(16)
            .map(|value| value.min(MAX_VALUE_AXIS_TRANSACTION_BYTES))
            .ok_or(ChartValueAxisError::InvalidSource)?;
        if source_bytes > aggregate {
            return Err(ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::InputBytes,
                observed: usize_to_u64(source_bytes),
                maximum: usize_to_u64(aggregate),
            });
        }
        let configured_output = limits.max_output_bytes();
        let maximum_output = aggregate.min(configured_output);
        // Retained and scratch charges are cumulative logical reservations
        // across all verification phases, rather than a measurement of
        // simultaneous resident memory. Keep their source-scaled allowance
        // finite and independently capped while accommodating repeated
        // archive, ZIP, Snappy, and candidate-reopen reservations.
        let memory = baseline
            .checked_mul(VALUE_AXIS_TRANSACTION_MEMORY_MULTIPLIER)
            .and_then(|value| value.checked_add(VALUE_AXIS_TRANSACTION_FIXED_OVERHEAD))
            .map(|value| value.min(MAX_VALUE_AXIS_TRANSACTION_MEMORY))
            .ok_or(ChartValueAxisError::InvalidSource)?;
        // One physical byte can participate in several independently
        // bounded allocation-bearing passes: selector graph construction,
        // canonical wire scans, archive cloning, candidate reopening, and
        // locality verification.  Scale the logical-event envelope with the
        // same finite pass multiplier as aggregate input/work instead of
        // assuming every source byte can be charged only once.
        let maximum_allocations = baseline
            .checked_mul(VALUE_AXIS_TRANSACTION_ALLOCATION_MULTIPLIER)
            .and_then(|value| value.checked_add(VALUE_AXIS_TRANSACTION_FIXED_ALLOCATIONS))
            .map(|value| value.min(MAX_VALUE_AXIS_TRANSACTION_ITEMS))
            .ok_or(ChartValueAxisError::InvalidSource)?;
        Ok(Self {
            limits,
            maximum_input: aggregate,
            maximum_output,
            maximum_fields: limits
                .max_fields()
                .checked_mul(16)
                .map(|value| value.min(MAX_VALUE_AXIS_TRANSACTION_ITEMS))
                .ok_or(ChartValueAxisError::InvalidSource)?,
            maximum_work: limits
                .max_rewrite_work()
                .checked_mul(16)
                .map(|value| value.min(MAX_VALUE_AXIS_TRANSACTION_BYTES))
                .ok_or(ChartValueAxisError::InvalidSource)?,
            maximum_references: package
                .semantic_limits()
                .max_references()
                .checked_mul(16)
                .map(|value| value.min(MAX_VALUE_AXIS_TRANSACTION_ITEMS))
                .ok_or(ChartValueAxisError::InvalidSource)?,
            maximum_allocations,
            maximum_retained: memory,
            maximum_scratch: memory,
            input: source_bytes,
            output: 0,
            fields: 0,
            work: 0,
            max_depth: 0,
            references: 0,
            allocations: 0,
            retained: 0,
            scratch: 0,
        })
    }

    fn charge_counter(
        current: &mut usize,
        amount: usize,
        maximum: usize,
        kind: ChartValueAxisLimitKind,
    ) -> Result<(), ChartValueAxisError> {
        let observed = current
            .checked_add(amount)
            .ok_or(ChartValueAxisError::InvalidSource)?;
        if observed > maximum {
            return Err(ChartValueAxisError::LimitExceeded {
                kind,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            });
        }
        *current = observed;
        Ok(())
    }

    fn input(&mut self, amount: usize) -> Result<(), ChartValueAxisError> {
        Self::charge_counter(
            &mut self.input,
            amount,
            self.maximum_input,
            ChartValueAxisLimitKind::InputBytes,
        )
    }

    fn output(&mut self, amount: usize) -> Result<(), ChartValueAxisError> {
        Self::charge_counter(
            &mut self.output,
            amount,
            self.maximum_output,
            ChartValueAxisLimitKind::OutputBytes,
        )
    }

    fn fields(&mut self, amount: usize) -> Result<(), ChartValueAxisError> {
        Self::charge_counter(
            &mut self.fields,
            amount,
            self.maximum_fields,
            ChartValueAxisLimitKind::WireFields,
        )
    }

    fn work(&mut self, amount: usize) -> Result<(), ChartValueAxisError> {
        Self::charge_counter(
            &mut self.work,
            amount,
            self.maximum_work,
            ChartValueAxisLimitKind::WireWork,
        )
    }

    fn references(&mut self, amount: usize) -> Result<(), ChartValueAxisError> {
        Self::charge_counter(
            &mut self.references,
            amount,
            self.maximum_references,
            ChartValueAxisLimitKind::References,
        )
    }

    fn allocations(
        &mut self,
        count: usize,
        retained: usize,
        scratch: usize,
    ) -> Result<(), ChartValueAxisError> {
        Self::charge_counter(
            &mut self.allocations,
            count,
            self.maximum_allocations,
            ChartValueAxisLimitKind::Entries,
        )?;
        Self::charge_counter(
            &mut self.retained,
            retained,
            self.maximum_retained,
            ChartValueAxisLimitKind::TotalBytes,
        )?;
        Self::charge_counter(
            &mut self.scratch,
            scratch,
            self.maximum_scratch,
            ChartValueAxisLimitKind::TotalBytes,
        )
    }

    fn charge_heap(&mut self, footprint: HeapFootprint) -> Result<(), ChartValueAxisError> {
        self.allocations(footprint.allocations, footprint.bytes, footprint.scratch)
    }

    fn charge_owned_clone(&mut self, bytes: usize) -> Result<(), ChartValueAxisError> {
        self.work(bytes)?;
        self.allocations(1, bytes, bytes)
    }

    /// Charge the temporary buffers used by the archive primitive that
    /// preserves an ArchiveInfo header while replacing one message.
    ///
    /// The core operation validates and encodes the metadata several times,
    /// parses both outer and nested wire-field vectors, and may retain two
    /// rewritten header copies. The exact allocator behavior is deliberately
    /// not promised; this is a checked upper envelope for those logical
    /// buffers and allocation events.
    fn charge_replace_message_buffers(
        &mut self,
        object: &ArchiveObject,
    ) -> Result<(), ChartValueAxisError> {
        let header = usize::try_from(object.header_length)
            .map_err(|_| ChartValueAxisError::InvalidSource)?;
        let mut footprint = HeapFootprint::default();
        // validate, canonical-before, replace-before/after validation, the
        // decoded verification metadata, and canonical-after each invoke the
        // compatibility projection or its neutral conversion.
        for _ in 0..6 {
            footprint.merge(archive_info_heap(&object.archive_info, false, true)?)?;
        }
        footprint.merge(archive_info_heap(&object.archive_info, true, false)?)?;

        let wire_capacity = header.max(1);
        // One field vector for the outer ArchiveInfo and one for its nested
        // MessageInfo are grown by parse_wire_fields_with_limits.
        footprint.add_vector::<WireField>(wire_capacity, true)?;
        footprint.add_vector::<WireField>(wire_capacity, true)?;
        // rewritten_message, rewritten_header, and canonical/header copies.
        for _ in 0..4 {
            footprint.add_bytes(header, true)?;
        }
        footprint.add_bytes(header, false)?;
        footprint.add_bytes(header, false)?;

        let metadata_work = archive_metadata_work_for_object(object)?;
        self.work(
            header
                .checked_mul(8)
                .and_then(|value| value.checked_add(metadata_work))
                .ok_or(ChartValueAxisError::InvalidSource)?,
        )?;
        self.charge_heap(footprint)
    }

    fn codec_options(&self, source_len: usize) -> Result<CodecOptions, ChartValueAxisError> {
        let input = self
            .maximum_input
            .checked_sub(self.input)
            .map(|value| value.min(self.limits.max_input_bytes()))
            .ok_or(ChartValueAxisError::InvalidSource)?;
        let output = self
            .maximum_output
            .checked_sub(self.output)
            .map(|value| value.min(self.limits.max_output_bytes()))
            .ok_or(ChartValueAxisError::InvalidSource)?;
        let fields = self
            .maximum_fields
            .checked_sub(self.fields)
            .map(|value| value.min(self.limits.max_fields()))
            .ok_or(ChartValueAxisError::InvalidSource)?;
        let work = self
            .maximum_work
            .checked_sub(self.work)
            .map(|value| value.min(self.limits.max_rewrite_work()))
            .ok_or(ChartValueAxisError::InvalidSource)?;
        if input < source_len || input == 0 {
            return Err(ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::InputBytes,
                observed: usize_to_u64(source_len.max(1)),
                maximum: usize_to_u64(input),
            });
        }
        let depth = u32::try_from(self.limits.max_nesting())
            .map_err(|_| ChartValueAxisError::InvalidSource)?;
        Ok(CodecOptions::new(input, fields, work, depth)
            .with_max_output_bytes(output)
            .with_max_allocations(
                self.maximum_allocations
                    .checked_sub(self.allocations)
                    .ok_or(ChartValueAxisError::InvalidSource)?,
            )
            .with_max_retained_bytes(
                self.maximum_retained
                    .checked_sub(self.retained)
                    .ok_or(ChartValueAxisError::InvalidSource)?,
            )
            .with_max_scratch_bytes(
                self.maximum_scratch
                    .checked_sub(self.scratch)
                    .ok_or(ChartValueAxisError::InvalidSource)?,
            ))
    }

    fn charge_codec_report(&mut self, report: CodecReport) -> Result<(), ChartValueAxisError> {
        self.input(report.source_bytes())?;
        self.fields(report.fields())?;
        self.work(report.work_bytes())?;
        self.charge_depth(report.max_depth())?;
        self.allocations(
            report.allocations(),
            report.retained_bytes(),
            report.scratch_bytes(),
        )
    }

    fn charge_requirements(
        &mut self,
        requirements: CodecRequirements,
    ) -> Result<(), ChartValueAxisError> {
        self.output(requirements.output_bytes)?;
        self.fields(requirements.fields)?;
        self.work(requirements.work_bytes)?;
        self.charge_depth(requirements.max_depth)?;
        self.allocations(
            requirements.allocations,
            requirements.retained_bytes,
            requirements.scratch_bytes,
        )
    }

    fn charge_depth(&mut self, depth: u32) -> Result<(), ChartValueAxisError> {
        self.max_depth = self.max_depth.max(depth);
        let maximum = u32::try_from(self.limits.max_nesting())
            .map_err(|_| ChartValueAxisError::InvalidSource)?;
        if self.max_depth > maximum {
            return Err(ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::WireNesting,
                observed: u64::from(self.max_depth),
                maximum: u64::from(maximum),
            });
        }
        Ok(())
    }

    fn charge_native_decompress_parse(
        &mut self,
        compressed: usize,
        decompressed: usize,
        archive: &Archive,
        snappy_limits: SnappyLimits,
    ) -> Result<(), ChartValueAxisError> {
        // Route the physical member charge through the shared support
        // contract as well as the local ledger implementation.  This keeps
        // the graph budget's input accounting live for every owner.
        <Self as super::chart_axis_support::AxisSupportBudget>::charge_input(self, compressed)
            .map_err(map_support_error)?;
        if compressed > snappy_limits.max_compressed_stream() {
            return Err(ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::EntryBytes,
                observed: usize_to_u64(compressed),
                maximum: usize_to_u64(snappy_limits.max_compressed_stream()),
            });
        }
        let (frames, _, _) = snappy_output_plan(decompressed, snappy_limits)?;
        let mut footprint = archive_reopen_heap(archive)?;
        // The decompressor grows one decoded stream across independent
        // frames. Keep its final bytes in the retained envelope and charge a
        // frame event for every possible reserve; the compressed member is
        // borrowed by the lower layer but remains a conservative scratch
        // bound for the transaction.
        footprint.add_bytes(decompressed, false)?;
        footprint.add_bytes(compressed, true)?;
        HeapFootprint::add_counter(&mut footprint.allocations, frames)?;
        let metadata_work = archive_metadata_work(archive)?;
        self.output(decompressed)?;
        self.work(
            compressed
                .checked_add(decompressed)
                .and_then(|value| value.checked_add(metadata_work))
                .and_then(|value| value.checked_add(frames))
                .ok_or(ChartValueAxisError::InvalidSource)?,
        )?;
        self.charge_heap(footprint)
    }

    fn charge_archive_snappy_plan(
        &mut self,
        archive: &Archive,
        archive_limits: litchi_iwa_core::Limits,
        snappy_limits: SnappyLimits,
    ) -> Result<(usize, usize), ChartValueAxisError> {
        let encoded = archive
            .encoded_len_with_limits(archive_limits)
            .map_err(map_core_error)?;
        let (footprint, frames) = archive_to_bytes_heap(archive, encoded, snappy_limits)?;
        let compressed = snappy_output_plan(encoded, snappy_limits)?.1;
        let metadata_work = archive_metadata_work(archive)?;
        self.output(
            encoded
                .checked_add(compressed)
                .ok_or(ChartValueAxisError::InvalidSource)?,
        )?;
        self.work(
            encoded
                .checked_add(compressed)
                .and_then(|value| value.checked_add(metadata_work))
                .and_then(|value| value.checked_add(frames))
                .ok_or(ChartValueAxisError::InvalidSource)?,
        )?;
        self.charge_heap(footprint)?;
        Ok((encoded, compressed))
    }

    fn charge_reassembly(
        &mut self,
        requirements: litchi_iwa_archive::package::ReassemblyExecutionRequirements,
    ) -> Result<litchi_iwa_archive::package::ReassemblyExecutionLimits, ChartValueAxisError> {
        self.output(requirements.output_bytes())?;
        self.work(
            requirements
                .output_bytes()
                .checked_add(requirements.scratch_bytes())
                .and_then(|value| value.checked_add(requirements.offset_count()))
                .ok_or(ChartValueAxisError::InvalidSource)?,
        )?;
        self.allocations(
            requirements.allocations(),
            requirements.retained_bytes(),
            requirements.scratch_bytes(),
        )?;
        Ok(requirements.exact_limits())
    }

    fn charge_reassembly_prepare(
        &mut self,
        source_bytes: usize,
        catalog_entries: usize,
        deleted_entries: usize,
        replacement_bytes: usize,
    ) -> Result<(), ChartValueAxisError> {
        // Preparation parses the complete ZIP and constructs entry-indexed
        // maps/sets before it exposes execution requirements.  Charge a
        // conservative bound for that work up front; the exact execution
        // requirements are charged separately after preparation.
        let selected = deleted_entries
            .checked_add(1)
            .ok_or(ChartValueAxisError::InvalidSource)?;
        let mut footprint = HeapFootprint::default();
        // ZipArchive keeps physical entries and central-directory order, and
        // Catalog keeps a second entry/index view. The concrete lower-layer
        // records are private, so these exposed fixed-size shapes are a
        // conservative capacity proxy rather than an allocator claim.
        footprint.add_vector::<[usize; 8]>(catalog_entries, false)?;
        footprint.add_vector::<usize>(catalog_entries, false)?;
        footprint.add_vector::<[usize; 4]>(catalog_entries, false)?;
        // prepare_mutations builds requested edit/deletion maps, matching
        // vectors/sets, and the prepared-edit map before execution.
        footprint.add_vector::<[usize; 4]>(1, true)?;
        footprint.add_vector::<[usize; 4]>(deleted_entries, true)?;
        footprint.add_vector::<[usize; 4]>(1, true)?;
        footprint.add_vector::<[usize; 4]>(deleted_entries, true)?;
        footprint.add_vector::<[usize; 4]>(selected, true)?;
        footprint.add_vector::<[usize; 8]>(1, true)?;
        // EntryEdit data is borrowed by the caller but the prepared mutation
        // needs a bounded member-data scratch envelope.
        footprint.add_bytes(replacement_bytes, true)?;
        let index_bytes = footprint
            .bytes
            .checked_add(footprint.scratch)
            .ok_or(ChartValueAxisError::InvalidSource)?;
        self.input(source_bytes)?;
        self.work(
            source_bytes
                .checked_add(index_bytes)
                .and_then(|value| value.checked_add(replacement_bytes))
                .ok_or(ChartValueAxisError::InvalidSource)?,
        )?;
        self.charge_heap(footprint)
    }

    fn charge_candidate_reopen(
        &mut self,
        source: &Package,
        bytes: usize,
    ) -> Result<(), ChartValueAxisError> {
        let catalog = physical_catalog(source)?;
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
        let mut footprint = HeapFootprint::default();
        let mut topology_work = 0usize;
        let mut objects = 0usize;
        for component in source.state.source.components().iter() {
            let archive = component.archive();
            footprint.merge(archive_reopen_heap(archive)?)?;
            let encoded = archive
                .encoded_len_with_limits(archive_limits)
                .map_err(map_core_error)?;
            let (frames, _, _) = snappy_output_plan(encoded, snappy_limits)?;
            let metadata_work = archive_metadata_work(archive)?;
            footprint.add_bytes(encoded, true)?;
            HeapFootprint::add_counter(&mut footprint.allocations, frames)?;
            topology_work = topology_work
                .checked_add(component.name().len())
                .and_then(|value| value.checked_add(encoded))
                .and_then(|value| value.checked_add(metadata_work))
                .ok_or(ChartValueAxisError::InvalidSource)?;
            objects = objects
                .checked_add(archive.objects.len())
                .ok_or(ChartValueAxisError::InvalidSource)?;
            // Component names are owned by the reopened SourceCatalog.
            footprint.add_bytes(component.name().len(), false)?;
        }

        let entries = catalog.package().len();
        // SourceCatalog's physical ZIP and logical catalog indexes, followed
        // by Package's object locator index. Fixed-size exposed shapes are
        // conservative proxies for private lower-layer records.
        footprint.add_vector::<[usize; 8]>(entries, false)?;
        footprint.add_vector::<usize>(entries, false)?;
        footprint.add_vector::<[usize; 4]>(entries, false)?;
        footprint.add_vector::<[usize; 4]>(objects, false)?;
        // The candidate Arc and the parser's bounded source view coexist
        // until reopen and verification complete.
        footprint.add_bytes(bytes, false)?;
        footprint.add_bytes(bytes, true)?;
        let footprint_work = footprint
            .bytes
            .checked_add(footprint.scratch)
            .ok_or(ChartValueAxisError::InvalidSource)?;
        self.input(bytes)?;
        self.output(bytes)?;
        self.scan_pass(source, 0)?;
        self.work(
            bytes
                .checked_add(topology_work)
                .and_then(|value| value.checked_add(footprint_work))
                .ok_or(ChartValueAxisError::InvalidSource)?,
        )?;
        self.charge_heap(footprint)
    }
}

/// One mutable aggregate of value-axis settings staged against an immutable
/// package snapshot.
pub struct ChartValueAxisEdit<'a> {
    source: &'a Package,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    axis_identifier: u64,
    axis_component_name: String,
    axis_message_index: usize,
    slide_identifier: u64,
    before: ValueAxisSettings,
    after: ValueAxisSettings,
}

impl fmt::Debug for ChartValueAxisEdit<'_> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartValueAxisEdit")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl<'a> ChartValueAxisEdit<'a> {
    fn new<'slide, 'chart>(
        source: &'a Package,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<Self, ChartValueAxisError> {
        let mut budget = ValueAxisBudget::new(source)?;
        let selection = select_value_axis(
            source,
            slide_selector.into(),
            chart_selector.into(),
            true,
            &mut budget,
        )?;
        let before = read_value_axis_settings(source, &selection, &mut budget)?;
        Ok(Self {
            source,
            slide_position: selection.slide_position,
            chart_position: selection.chart_position,
            chart_identifier: selection.chart_identifier,
            axis_identifier: selection.axis_identifier,
            axis_component_name: selection.axis_component_name,
            axis_message_index: selection.axis_message_index,
            slide_identifier: selection.slide_identifier,
            before,
            after: before,
        })
    }

    /// Return the aggregate observed when this edit began.
    #[must_use]
    pub const fn before(&self) -> ValueAxisSettings {
        self.before
    }

    /// Return the aggregate staged for publication.
    #[must_use]
    pub const fn after(&self) -> ValueAxisSettings {
        self.after
    }

    /// Stage all value-axis settings atomically.
    pub fn set(mut self, settings: ValueAxisSettings) -> Result<Self, ChartValueAxisError> {
        self.after = validate_settings(settings)?;
        Ok(self)
    }

    /// Stage replacement bounds while retaining steps and scale.
    pub fn set_bounds(mut self, bounds: Bounds) -> Result<Self, ChartValueAxisError> {
        self.after = self.after.with_bounds(bounds);
        self.after = validate_settings(self.after)?;
        Ok(self)
    }

    /// Stage replacement major/minor steps while retaining bounds and scale.
    pub fn set_steps(mut self, steps: Steps) -> Result<Self, ChartValueAxisError> {
        self.after = self.after.with_steps(steps);
        self.after = validate_settings(self.after)?;
        Ok(self)
    }

    /// Stage replacement scale while retaining bounds and steps.
    pub fn set_scale(mut self, scale: Scale) -> Result<Self, ChartValueAxisError> {
        self.after = self.after.with_scale(scale);
        self.after = validate_settings(self.after)?;
        Ok(self)
    }

    /// Validate and atomically publish the staged immutable candidate.
    pub fn commit(self) -> Result<ChartValueAxisCommit, ChartValueAxisError> {
        let catalog = physical_catalog(self.source)?;
        let source_bytes = catalog.shared_source();
        let mut budget = ValueAxisBudget::new(self.source)?;
        let source_selection = select_value_axis(
            self.source,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            true,
            &mut budget,
        )?;
        let current = read_value_axis_settings(self.source, &source_selection, &mut budget)?;
        if !same_selection(&source_selection, &self) || !settings_equal_exact(current, self.before)
        {
            return Err(ChartValueAxisError::InvalidSource);
        }
        if settings_equal_exact(self.before, self.after) {
            budget.scan_pass(self.source, 0)?;
            self.source.validate().map_err(map_read_error)?;
            return Ok(ChartValueAxisCommit {
                package: self.source.snapshot(),
                patch: ChartValueAxisPatch {
                    artifacts: ExactArtifacts::new(Arc::clone(&source_bytes), source_bytes),
                    slide_position: self.slide_position,
                    chart_position: self.chart_position,
                    chart_identifier: self.chart_identifier,
                    axis_identifier: self.axis_identifier,
                    axis_component_name: self.axis_component_name,
                    axis_message_index: self.axis_message_index,
                    slide_identifier: self.slide_identifier,
                    before: self.before,
                    after: self.after,
                    before_message: None,
                    after_message: None,
                    deleted_previews: 0,
                    target_requires_invalidated_previews: false,
                },
                diagnostics: ChartValueAxisDiagnostics::unchanged(),
            });
        }
        if !catalog.source_is_exact() {
            return Err(ChartValueAxisError::UnsupportedSource);
        }
        budget.scan_pass(self.source, 0)?;
        self.source.validate().map_err(map_read_error)?;
        let before_message = copy_selected_message(self.source, &source_selection, &mut budget)?;
        let (package, deleted_previews, after_message) =
            rewrite_value_axis_settings(self.source, &source_selection, self.after, &mut budget)?;
        budget.scan_pass(&package, 0)?;
        package.validate().map_err(map_read_error)?;
        let target = physical_catalog(&package)?.shared_source();
        let candidate_selection = select_value_axis(
            &package,
            SlideSelector::position(self.slide_position),
            ChartSelector::index(self.chart_position.get()),
            true,
            &mut budget,
        )?;
        let candidate_settings =
            read_value_axis_settings(&package, &candidate_selection, &mut budget)?;
        if !same_selection(&candidate_selection, &self)
            || !settings_equal_exact(candidate_settings, self.after)
        {
            return Err(ChartValueAxisError::Verification);
        }
        super::chart_axis_support::verify_package_locality_with_expected_message(
            self.source,
            &package,
            &source_selection,
            true,
            Some(after_message.as_ref()),
            &mut budget,
        )
        .map_err(map_support_error)?;
        Ok(ChartValueAxisCommit {
            package,
            patch: ChartValueAxisPatch {
                artifacts: ExactArtifacts::new(source_bytes, Arc::clone(&target)),
                slide_position: self.slide_position,
                chart_position: self.chart_position,
                chart_identifier: self.chart_identifier,
                axis_identifier: self.axis_identifier,
                axis_component_name: self.axis_component_name,
                axis_message_index: self.axis_message_index,
                slide_identifier: self.slide_identifier,
                before: self.before,
                after: self.after,
                before_message: Some(before_message),
                after_message: Some(after_message),
                deleted_previews,
                target_requires_invalidated_previews: true,
            },
            diagnostics: ChartValueAxisDiagnostics::published(deleted_previews),
        })
    }
}

/// An exact-source-checked reversible aggregate value-axis patch.
#[derive(Clone, PartialEq)]
pub struct ChartValueAxisPatch {
    artifacts: ExactArtifacts,
    slide_position: Position,
    chart_position: Position,
    chart_identifier: u64,
    axis_identifier: u64,
    axis_component_name: String,
    axis_message_index: usize,
    slide_identifier: u64,
    before: ValueAxisSettings,
    after: ValueAxisSettings,
    // Private source-authoritative bytes used by locality verification.  They
    // are deliberately omitted from the semantic/debug surface: a patch can
    // authorize only the deterministic selected message it was built for.
    before_message: Option<Arc<[u8]>>,
    after_message: Option<Arc<[u8]>>,
    deleted_previews: usize,
    target_requires_invalidated_previews: bool,
}

impl fmt::Debug for ChartValueAxisPatch {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("ChartValueAxisPatch")
            .field("slide_position", &self.slide_position)
            .field("chart_position", &self.chart_position)
            .field("before", &self.before)
            .field("after", &self.after)
            .finish_non_exhaustive()
    }
}

impl ChartValueAxisPatch {
    /// Return the selected semantic slide position.
    #[must_use]
    pub const fn slide_position(&self) -> Position {
        self.slide_position
    }

    /// Return the selected semantic chart position.
    #[must_use]
    pub const fn chart_position(&self) -> Position {
        self.chart_position
    }

    /// Return the settings required from the source package.
    #[must_use]
    pub const fn before(&self) -> ValueAxisSettings {
        self.before
    }

    /// Return the settings produced by this patch.
    #[must_use]
    pub const fn after(&self) -> ValueAxisSettings {
        self.after
    }

    /// Return the base package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn source_fingerprint(&self) -> u64 {
        self.artifacts.source_fingerprint()
    }

    /// Return the target package's compact diagnostic fingerprint.
    #[must_use]
    pub const fn target_fingerprint(&self) -> u64 {
        self.artifacts.target_fingerprint()
    }

    /// Return whether this patch preserves the exact source bytes.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        settings_equal_exact(self.before, self.after) && self.artifacts.is_byte_noop()
    }

    /// Return an exact reversible patch from target back to source.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            artifacts: self.artifacts.inverse(),
            slide_position: self.slide_position,
            chart_position: self.chart_position,
            chart_identifier: self.chart_identifier,
            axis_identifier: self.axis_identifier,
            axis_component_name: self.axis_component_name.clone(),
            axis_message_index: self.axis_message_index,
            slide_identifier: self.slide_identifier,
            before: self.after,
            after: self.before,
            before_message: self.after_message.clone(),
            after_message: self.before_message.clone(),
            deleted_previews: 0,
            target_requires_invalidated_previews: !self.artifacts.is_byte_noop()
                && !self.target_requires_invalidated_previews,
        }
    }
}

/// Compact evidence describing one value-axis settings commit.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct ChartValueAxisDiagnostics {
    changed: bool,
    touched_components: usize,
    deleted_previews: usize,
    full_reparse_performed: bool,
}

impl ChartValueAxisDiagnostics {
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

    /// Return the number of rewritten physical IWA components.
    #[must_use]
    pub const fn touched_components(self) -> usize {
        self.touched_components
    }

    /// Return how many stale root previews were deleted.
    #[must_use]
    pub const fn deleted_previews(self) -> usize {
        self.deleted_previews
    }

    /// Return whether the candidate was fully reopened before publication.
    #[must_use]
    pub const fn full_reparse_performed(self) -> bool {
        self.full_reparse_performed
    }
}

/// Fully verified result of one immutable value-axis settings transaction.
#[must_use = "a Keynote chart value-axis commit contains the validated package snapshot"]
#[derive(Debug)]
pub struct ChartValueAxisCommit {
    package: Package,
    patch: ChartValueAxisPatch,
    diagnostics: ChartValueAxisDiagnostics,
}

impl ChartValueAxisCommit {
    /// Borrow the fully reopened immutable package snapshot.
    #[must_use]
    pub const fn package(&self) -> &Package {
        &self.package
    }

    /// Consume this commit and return its package snapshot.
    #[must_use]
    pub fn into_package(self) -> Package {
        self.package
    }

    /// Borrow the reversible exact-source patch.
    #[must_use]
    pub const fn patch(&self) -> &ChartValueAxisPatch {
        &self.patch
    }

    /// Borrow compact publication diagnostics.
    #[must_use]
    pub const fn diagnostics(&self) -> &ChartValueAxisDiagnostics {
        &self.diagnostics
    }
}

impl Package {
    /// Read the complete settings of one selected chart's primary value axis.
    pub fn slide_chart_value_axis_settings<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ValueAxisSettings, ChartValueAxisError> {
        let mut budget = ValueAxisBudget::new(self)?;
        let selection = select_value_axis(
            self,
            slide_selector.into(),
            chart_selector.into(),
            false,
            &mut budget,
        )?;
        read_value_axis_settings(self, &selection, &mut budget)
    }

    /// Start an exact immutable edit of one selected chart value axis.
    pub fn edit_slide_chart_value_axis_settings<'slide, 'chart>(
        &self,
        slide_selector: impl Into<SlideSelector<'slide>>,
        chart_selector: impl Into<ChartSelector<'chart>>,
    ) -> Result<ChartValueAxisEdit<'_>, ChartValueAxisError> {
        ChartValueAxisEdit::new(self, slide_selector, chart_selector)
    }

    /// Apply an exact-source-checked aggregate value-axis patch.
    pub fn apply_slide_chart_value_axis_settings(
        &self,
        patch: &ChartValueAxisPatch,
    ) -> Result<ChartValueAxisCommit, ChartValueAxisError> {
        let catalog = physical_catalog(self)?;
        let source = catalog.shared_source();
        if !patch.artifacts.authorizes_source(&source) {
            return Err(ChartValueAxisError::PatchConflict);
        }
        let mut budget = ValueAxisBudget::new(self)?;
        let selection = select_value_axis(
            self,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            true,
            &mut budget,
        )?;
        let current = read_value_axis_settings(self, &selection, &mut budget)?;
        if !selection_matches_patch(&selection, patch)
            || !settings_equal_exact(current, patch.before)
        {
            return Err(ChartValueAxisError::PatchConflict);
        }
        if patch.is_noop() {
            budget.scan_pass(self, 0)?;
            self.validate().map_err(map_read_error)?;
            return Ok(ChartValueAxisCommit {
                package: self.snapshot(),
                patch: patch.clone(),
                diagnostics: ChartValueAxisDiagnostics::unchanged(),
            });
        }
        let expected_source_message = patch
            .before_message
            .as_deref()
            .ok_or(ChartValueAxisError::PatchConflict)?;
        budget.work(expected_source_message.len())?;
        let source_message = self
            .object(patch.axis_identifier)
            .and_then(|object| object.messages.get(patch.axis_message_index))
            .filter(|message| message.type_ == CHART_AXIS_MESSAGE_TYPE)
            .ok_or(ChartValueAxisError::PatchConflict)?;
        if source_message.data.as_slice() != expected_source_message {
            return Err(ChartValueAxisError::PatchConflict);
        }
        budget.scan_pass(self, 0)?;
        if !catalog.source_is_exact() {
            return Err(ChartValueAxisError::PatchConflict);
        }
        budget.charge_candidate_reopen(self, patch.artifacts.target().len())?;
        let candidate =
            Package::from_source_with_options(patch.artifacts.target(), self.state.options)
                .map_err(map_read_error)?;
        budget.scan_pass(&candidate, 0)?;
        candidate.validate().map_err(map_read_error)?;
        let candidate_selection = select_value_axis(
            &candidate,
            SlideSelector::position(patch.slide_position),
            ChartSelector::index(patch.chart_position.get()),
            true,
            &mut budget,
        )?;
        let candidate_settings =
            read_value_axis_settings(&candidate, &candidate_selection, &mut budget)?;
        let expected_selected_message = patch
            .after_message
            .as_deref()
            .ok_or(ChartValueAxisError::InvalidSource)?;
        if !selection_matches_patch(&candidate_selection, patch)
            || !settings_equal_exact(candidate_settings, patch.after)
        {
            return Err(ChartValueAxisError::Verification);
        }
        super::chart_axis_support::verify_package_locality_with_expected_message(
            self,
            &candidate,
            &selection,
            patch.target_requires_invalidated_previews,
            Some(expected_selected_message),
            &mut budget,
        )
        .map_err(map_support_error)?;
        Ok(ChartValueAxisCommit {
            package: candidate,
            patch: patch.clone(),
            diagnostics: ChartValueAxisDiagnostics::published(patch.deleted_previews),
        })
    }
}

fn select_value_axis(
    package: &Package,
    slide_selector: SlideSelector<'_>,
    chart_selector: ChartSelector<'_>,
    mutation_guards: bool,
    budget: &mut ValueAxisBudget,
) -> Result<super::chart_axis_support::AxisSelection, ChartValueAxisError> {
    super::chart_axis_support::select_axis(
        package,
        slide_selector,
        chart_selector,
        crate::Axis::Value,
        mutation_guards,
        budget,
    )
    .map_err(map_support_error)
}

fn read_value_axis_settings(
    package: &Package,
    selection: &super::chart_axis_support::AxisSelection,
    budget: &mut ValueAxisBudget,
) -> Result<ValueAxisSettings, ChartValueAxisError> {
    let object = package
        .object(selection.axis_identifier)
        .ok_or(ChartValueAxisError::InvalidSource)?;
    let message = object
        .messages
        .get(selection.axis_message_index)
        .filter(|message| message.type_ == CHART_AXIS_MESSAGE_TYPE)
        .ok_or(ChartValueAxisError::InvalidSource)?;
    let limits = package.wire_limits().map_err(map_wire_error)?;
    let fields = super::chart_axis_support::accounted_wire_fields(&message.data, limits, budget)
        .map_err(map_support_error)?;
    let Some(field) = fields
        .iter()
        .copied()
        .find(|field| field.number() == GENERATED_CHART_AXIS_EXTENSION_FIELD)
    else {
        return Ok(ValueAxisSettings::automatic());
    };
    if fields
        .iter()
        .filter(|candidate| candidate.number() == GENERATED_CHART_AXIS_EXTENSION_FIELD)
        .count()
        != 1
        || field.wire_type() != 2
    {
        return Err(ChartValueAxisError::InvalidSource);
    }
    field
        .validate_canonical_framing(&message.data)
        .map_err(map_wire_error)?;
    let extension = field.payload(&message.data).map_err(map_wire_error)?;
    let options = budget.codec_options(extension.len())?;
    let (snapshot, report) =
        decode_axis_value_settings_with_report(extension, options).map_err(map_codec_error)?;
    budget.charge_codec_report(report)?;
    common_settings(snapshot)
}

fn same_selection(
    selection: &super::chart_axis_support::AxisSelection,
    edit: &ChartValueAxisEdit<'_>,
) -> bool {
    selection.slide_position == edit.slide_position
        && selection.chart_position == edit.chart_position
        && selection.chart_identifier == edit.chart_identifier
        && selection.axis_identifier == edit.axis_identifier
        && selection.axis_component_name == edit.axis_component_name
        && selection.axis_message_index == edit.axis_message_index
        && selection.slide_identifier == edit.slide_identifier
}

fn selection_matches_patch(
    selection: &super::chart_axis_support::AxisSelection,
    patch: &ChartValueAxisPatch,
) -> bool {
    selection.slide_position == patch.slide_position
        && selection.chart_position == patch.chart_position
        && selection.chart_identifier == patch.chart_identifier
        && selection.axis_identifier == patch.axis_identifier
        && selection.axis_component_name == patch.axis_component_name
        && selection.axis_message_index == patch.axis_message_index
        && selection.slide_identifier == patch.slide_identifier
}

fn common_settings(snapshot: CodecSnapshot<'_>) -> Result<ValueAxisSettings, ChartValueAxisError> {
    let bounds = snapshot.bounds();
    let minimum = bounds
        .minimum()
        .map(|bound| Bound::new(bound.value()))
        .transpose()
        .map_err(|_| ChartValueAxisError::InvalidSource)?;
    let maximum = bounds
        .maximum()
        .map(|bound| Bound::new(bound.value()))
        .transpose()
        .map_err(|_| ChartValueAxisError::InvalidSource)?;
    let bounds = Bounds::new(minimum, maximum).map_err(|_| ChartValueAxisError::InvalidSource)?;
    let steps = snapshot.steps();
    let major = steps
        .major()
        .map(MajorStepCount::new)
        .transpose()
        .map_err(|_| ChartValueAxisError::InvalidSource)?;
    let minor = steps
        .minor()
        .map(MinorStepCount::new)
        .transpose()
        .map_err(|_| ChartValueAxisError::InvalidSource)?;
    let settings = ValueAxisSettings::new(
        bounds,
        Steps::new(major, minor),
        Scale::from_native(snapshot.scale().native_value()),
    );
    // A malformed source must not be silently promoted to a publishable
    // value-axis state.  In particular, native logarithmic axes cannot have
    // an explicit non-positive endpoint.
    validate_settings(settings).map_err(|_| ChartValueAxisError::InvalidSource)
}

fn validate_settings(
    settings: ValueAxisSettings,
) -> Result<ValueAxisSettings, ChartValueAxisError> {
    let minimum = settings.bounds().minimum();
    let maximum = settings.bounds().maximum();
    if minimum.is_some_and(|bound| !bound.value().is_finite())
        || maximum.is_some_and(|bound| !bound.value().is_finite())
        || matches!((minimum, maximum), (Some(low), Some(high)) if low.value() > high.value())
    {
        return Err(ChartValueAxisError::InvalidSettings);
    }
    if settings.scale() == Scale::Logarithmic
        && (minimum.is_some_and(|bound| bound.value() <= 0.0)
            || maximum.is_some_and(|bound| bound.value() <= 0.0))
    {
        return Err(ChartValueAxisError::InvalidSettings);
    }
    // Values 1 and 2 are canonicalized so a caller cannot stage an
    // Unsupported variant which the native reader necessarily maps back to a
    // known scale during candidate verification.
    Ok(settings.with_scale(Scale::from_native(settings.scale().native_value())))
}

fn settings_equal_exact(left: ValueAxisSettings, right: ValueAxisSettings) -> bool {
    let left_bounds = left.bounds();
    let right_bounds = right.bounds();
    let bounds_equal = match (
        left_bounds.minimum(),
        right_bounds.minimum(),
        left_bounds.maximum(),
        right_bounds.maximum(),
    ) {
        (Some(left), Some(right), Some(left_max), Some(right_max)) => {
            left.value() == right.value() && left_max.value() == right_max.value()
        },
        (Some(left), Some(right), None, None) => left.value() == right.value(),
        (None, None, Some(left), Some(right)) => left.value() == right.value(),
        (None, None, None, None) => true,
        _ => false,
    };
    bounds_equal
        && left.steps() == right.steps()
        && Scale::from_native(left.scale().native_value())
            == Scale::from_native(right.scale().native_value())
}

fn codec_settings(
    settings: ValueAxisSettings,
) -> Result<
    litchi_iwa_protos::keynote_chart_axis_value_settings_codec::AxisValueSettings,
    ChartValueAxisError,
> {
    let bounds = settings.bounds();
    let minimum = bounds
        .minimum()
        .map(|bound| CodecBound::new(bound.value()))
        .transpose()
        .map_err(|_| ChartValueAxisError::InvalidSettings)?;
    let maximum = bounds
        .maximum()
        .map(|bound| CodecBound::new(bound.value()))
        .transpose()
        .map_err(|_| ChartValueAxisError::InvalidSettings)?;
    let bounds =
        CodecBounds::new(minimum, maximum).map_err(|_| ChartValueAxisError::InvalidSettings)?;
    let steps = settings.steps();
    let steps = CodecSteps::new(
        steps.major().map(|value| value.value()),
        steps.minor().map(|value| value.value()),
    )
    .map_err(|_| ChartValueAxisError::InvalidSettings)?;
    let scale = CodecScale::from_native(settings.scale().native_value());
    Ok(
        litchi_iwa_protos::keynote_chart_axis_value_settings_codec::AxisValueSettings::new(
            bounds, steps, scale,
        ),
    )
}

fn codec_write(
    snapshot: CodecSnapshot<'_>,
    settings: ValueAxisSettings,
) -> Result<CodecWrite, ChartValueAxisError> {
    let native = codec_settings(settings)?;
    let write = CodecWrite::preserve()
        .with_bounds(native.bounds())
        .with_steps(native.steps());
    Ok(if snapshot.scale_present() {
        write.with_explicit_scale(native.scale())
    } else {
        write.with_scale(native.scale())
    })
}

fn map_support_error(error: super::chart_axis_support::AxisSupportError) -> ChartValueAxisError {
    use super::chart_axis_support::{
        AxisSupportError, AxisSupportLimitKind, AxisSupportSelectorError,
    };
    match error {
        AxisSupportError::Selector(selector) => match selector {
            AxisSupportSelectorError::UnsupportedSource => ChartValueAxisError::UnsupportedSource,
            AxisSupportSelectorError::AmbiguousSelector => ChartValueAxisError::AmbiguousSelector,
            AxisSupportSelectorError::EmptySlideName => ChartValueAxisError::EmptySlideName,
            AxisSupportSelectorError::SlideNameNotFound => ChartValueAxisError::SlideNameNotFound,
            AxisSupportSelectorError::SlidePositionNotFound { position } => {
                ChartValueAxisError::SlidePositionNotFound { position }
            },
            AxisSupportSelectorError::ChartNameNotFound => ChartValueAxisError::ChartNameNotFound,
            AxisSupportSelectorError::ChartPositionNotFound { position } => {
                ChartValueAxisError::ChartPositionNotFound { position }
            },
            AxisSupportSelectorError::EmptyChartName => ChartValueAxisError::EmptyChartName,
        },
        AxisSupportError::InvalidSource => ChartValueAxisError::InvalidSource,
        AxisSupportError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => ChartValueAxisError::LimitExceeded {
            kind: match kind {
                AxisSupportLimitKind::InputBytes => ChartValueAxisLimitKind::InputBytes,
                AxisSupportLimitKind::OutputBytes => ChartValueAxisLimitKind::OutputBytes,
                AxisSupportLimitKind::WireBytes => ChartValueAxisLimitKind::WireBytes,
                AxisSupportLimitKind::Entries => ChartValueAxisLimitKind::Entries,
                AxisSupportLimitKind::EntryBytes => ChartValueAxisLimitKind::EntryBytes,
                AxisSupportLimitKind::TotalBytes => ChartValueAxisLimitKind::TotalBytes,
                AxisSupportLimitKind::Slides => ChartValueAxisLimitKind::Slides,
                AxisSupportLimitKind::References => ChartValueAxisLimitKind::References,
                AxisSupportLimitKind::WireFields => ChartValueAxisLimitKind::WireFields,
                AxisSupportLimitKind::WireNesting => ChartValueAxisLimitKind::WireNesting,
                AxisSupportLimitKind::WireWork => ChartValueAxisLimitKind::WireWork,
                AxisSupportLimitKind::TextStorages
                | AxisSupportLimitKind::TextFragments
                | AxisSupportLimitKind::TextBytes
                | AxisSupportLimitKind::TitleBytes => ChartValueAxisLimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        AxisSupportError::Allocation { amount } => ChartValueAxisError::Allocation { amount },
    }
}

fn map_budget_error(error: ChartValueAxisError) -> super::chart_axis_support::AxisSupportError {
    use super::chart_axis_support::{AxisSupportError, AxisSupportLimitKind};
    match error {
        ChartValueAxisError::LimitExceeded {
            kind,
            observed,
            maximum,
        } => AxisSupportError::LimitExceeded {
            kind: match kind {
                ChartValueAxisLimitKind::InputBytes => AxisSupportLimitKind::InputBytes,
                ChartValueAxisLimitKind::OutputBytes => AxisSupportLimitKind::OutputBytes,
                ChartValueAxisLimitKind::WireBytes => AxisSupportLimitKind::WireBytes,
                ChartValueAxisLimitKind::Entries => AxisSupportLimitKind::Entries,
                ChartValueAxisLimitKind::EntryBytes => AxisSupportLimitKind::EntryBytes,
                ChartValueAxisLimitKind::TotalBytes => AxisSupportLimitKind::TotalBytes,
                ChartValueAxisLimitKind::Slides => AxisSupportLimitKind::Slides,
                ChartValueAxisLimitKind::References => AxisSupportLimitKind::References,
                ChartValueAxisLimitKind::WireFields => AxisSupportLimitKind::WireFields,
                ChartValueAxisLimitKind::WireNesting => AxisSupportLimitKind::WireNesting,
                ChartValueAxisLimitKind::WireWork => AxisSupportLimitKind::WireWork,
            },
            observed,
            maximum,
        },
        ChartValueAxisError::Allocation { amount } => AxisSupportError::Allocation { amount },
        _ => AxisSupportError::InvalidSource,
    }
}

fn copy_selected_message(
    package: &Package,
    selection: &super::chart_axis_support::AxisSelection,
    budget: &mut ValueAxisBudget,
) -> Result<Arc<[u8]>, ChartValueAxisError> {
    let message = package
        .object(selection.axis_identifier)
        .and_then(|object| object.messages.get(selection.axis_message_index))
        .filter(|message| message.type_ == CHART_AXIS_MESSAGE_TYPE)
        .ok_or(ChartValueAxisError::InvalidSource)?;
    budget.charge_owned_clone(message.data.len())?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(message.data.len())
        .map_err(|_| ChartValueAxisError::Allocation {
            amount: message.data.len(),
        })?;
    bytes.extend_from_slice(&message.data);
    Ok(bytes.into())
}

fn rewrite_value_axis_settings(
    source: &Package,
    selection: &super::chart_axis_support::AxisSelection,
    after: ValueAxisSettings,
    budget: &mut ValueAxisBudget,
) -> Result<(Package, usize, Arc<[u8]>), ChartValueAxisError> {
    let catalog = physical_catalog(source)?;
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == selection.axis_component_name)
        .ok_or(ChartValueAxisError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(ChartValueAxisError::InvalidSource);
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
    let component = source
        .state
        .source
        .components()
        .iter()
        .find(|component| component.name() == selection.axis_component_name)
        .ok_or(ChartValueAxisError::InvalidSource)?;
    let decompressed_bound = component
        .archive()
        .encoded_len_with_limits(archive_limits)
        .map_err(map_core_error)?;
    budget.charge_native_decompress_parse(
        entry.data().len(),
        decompressed_bound,
        component.archive(),
        snappy_limits,
    )?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(map_core_error)?;
    let mut archive =
        Archive::parse_with_limits(stream.as_bytes(), archive_limits).map_err(map_core_error)?;
    super::chart_axis_support::validate_canonical_object_length_prefixes_with_budget(
        stream.as_bytes(),
        &archive,
        budget,
    )
    .map_err(map_support_error)?;
    let message_data = {
        let object = archive
            .object(selection.axis_identifier)
            .ok_or(ChartValueAxisError::InvalidSource)?;
        let message = object
            .messages
            .get(selection.axis_message_index)
            .filter(|message| message.type_ == CHART_AXIS_MESSAGE_TYPE)
            .ok_or(ChartValueAxisError::InvalidSource)?;
        budget.charge_owned_clone(message.data.len())?;
        let mut cloned = Vec::new();
        cloned.try_reserve_exact(message.data.len()).map_err(|_| {
            ChartValueAxisError::Allocation {
                amount: message.data.len(),
            }
        })?;
        cloned.extend_from_slice(&message.data);
        cloned
    };
    let limits = source.wire_limits().map_err(map_wire_error)?;
    let fields = super::chart_axis_support::accounted_wire_fields(&message_data, limits, budget)
        .map_err(map_support_error)?;
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
        return Err(ChartValueAxisError::InvalidSource);
    }
    if let Some(field) = extension_field
        && field.wire_type() != 2
    {
        return Err(ChartValueAxisError::InvalidSource);
    }
    let (extension, snapshot) = if let Some(field) = extension_field {
        field
            .validate_canonical_framing(&message_data)
            .map_err(map_wire_error)?;
        let extension = field.payload(&message_data).map_err(map_wire_error)?;
        let options = budget.codec_options(extension.len())?;
        let (snapshot, report) =
            decode_axis_value_settings_with_report(extension, options).map_err(map_codec_error)?;
        budget.charge_codec_report(report)?;
        (extension.to_owned(), snapshot)
    } else {
        let options = budget.codec_options(0)?;
        let (snapshot, report) =
            decode_axis_value_settings_with_report(&[], options).map_err(map_codec_error)?;
        budget.charge_codec_report(report)?;
        (Vec::new(), snapshot)
    };
    let write = codec_write(snapshot, after)?;
    let options = budget.codec_options(extension.len())?;
    let prepared =
        prepare_axis_value_settings_rewrite(&extension, write, options).map_err(map_codec_error)?;
    let preparation = prepared.prepare_report();
    budget.charge_codec_report(preparation)?;
    let requirements = prepared.execution_requirements();
    budget.charge_requirements(requirements)?;
    let replacement = prepared
        .execute(requirements.exact())
        .map_err(map_codec_error)?
        .into_output();
    let patched =
        rewrite_chart_axis_extension(&message_data, extension_field, &replacement, limits, budget)?;
    budget.allocations(1, patched.len(), patched.len())?;
    let expected_message: Arc<[u8]> = patched.clone().into();
    let object = archive
        .object(selection.axis_identifier)
        .ok_or(ChartValueAxisError::InvalidSource)?;
    budget.charge_replace_message_buffers(object)?;
    archive
        .object_mut(selection.axis_identifier)
        .ok_or(ChartValueAxisError::InvalidSource)?
        .replace_message_preserving_header_with_limits(
            selection.axis_message_index,
            RawMessage {
                type_: CHART_AXIS_MESSAGE_TYPE,
                data: patched,
            },
            archive_limits,
        )
        .map_err(map_core_error)?;
    let (_encoded_bound, _compressed_bound) =
        budget.charge_archive_snappy_plan(&archive, archive_limits, snappy_limits)?;
    let bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(map_core_error)?;
    let compressed = SnappyStream::compress(&bytes).map_err(map_core_error)?;
    let previews = super::rendering_invalidation::root_preview_deletions(catalog.package())
        .map_err(|_| ChartValueAxisError::InvalidSource)?;
    let edit = EntryEdit::new(
        selection.axis_component_name.as_str(),
        compressed.as_slice(),
    );
    let edits = [edit];
    budget.charge_reassembly_prepare(
        catalog.shared_source().len(),
        catalog.package().len(),
        previews.len(),
        compressed.len(),
    )?;
    let prepared_reassembly = catalog
        .prepare_reassembly_with_deletions(&edits, previews.names(), source.state.options.archive())
        .map_err(map_archive_error)?;
    let requirements = prepared_reassembly.execution_requirements();
    budget.charge_candidate_reopen(source, requirements.output_bytes())?;
    let reassembly_limits = budget.charge_reassembly(requirements)?;
    let output = litchi_iwa_archive::package::PreparedReassembly::execute(
        prepared_reassembly,
        reassembly_limits,
    )
    .map_err(map_archive_error)?;
    let candidate = Package::from_source_with_options(output.into(), source.state.options)
        .map_err(map_read_error)?;
    Ok((candidate, previews.len(), expected_message))
}

fn rewrite_chart_axis_extension(
    data: &[u8],
    extension_field: Option<WireField>,
    replacement: &[u8],
    limits: WireLimits,
    budget: &mut ValueAxisBudget,
) -> Result<Vec<u8>, ChartValueAxisError> {
    let replacement_length =
        u64::try_from(replacement.len()).map_err(|_| ChartValueAxisError::InvalidSource)?;
    let key_length = extension_field.map_or_else(
        || encoded_len((u64::from(GENERATED_CHART_AXIS_EXTENSION_FIELD) << 3) | 2),
        |field| field.key_end() - field.start(),
    );
    let replacement_field_length = key_length
        .checked_add(encoded_len(replacement_length))
        .and_then(|length| length.checked_add(replacement.len()))
        .ok_or(ChartValueAxisError::InvalidSource)?;
    let output_length = extension_field
        .map_or_else(
            || data.len().checked_add(replacement_field_length),
            |field| {
                data.len()
                    .checked_sub(field.end() - field.start())
                    .and_then(|length| length.checked_add(replacement_field_length))
            },
        )
        .ok_or(ChartValueAxisError::InvalidSource)?;
    if output_length > limits.max_output_bytes() || output_length > MAX_VALUE_AXIS_BYTES {
        return Err(ChartValueAxisError::LimitExceeded {
            kind: ChartValueAxisLimitKind::OutputBytes,
            observed: usize_to_u64(output_length),
            maximum: usize_to_u64(limits.max_output_bytes().min(MAX_VALUE_AXIS_BYTES)),
        });
    }
    budget.output(output_length)?;
    budget.work(output_length)?;
    budget.allocations(1, output_length, output_length)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(output_length)
        .map_err(|_| ChartValueAxisError::Allocation {
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

fn map_codec_error(error: CodecError) -> ChartValueAxisError {
    if let Some(limit) = error.resource_limit() {
        return match limit {
            CodecLimit::Bytes { observed, maximum } => ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::WireBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            CodecLimit::Fields { observed, maximum } => ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::WireFields,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            CodecLimit::Work { observed, maximum } => ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::WireWork,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            CodecLimit::Output { observed, maximum } => ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::OutputBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            CodecLimit::Nesting { observed, maximum } => ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::WireNesting,
                observed: u64::from(observed),
                maximum: u64::from(maximum),
            },
            CodecLimit::Allocations { observed, maximum } => ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::Entries,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            CodecLimit::Retained { observed, maximum }
            | CodecLimit::Scratch { observed, maximum } => ChartValueAxisError::LimitExceeded {
                kind: ChartValueAxisLimitKind::TotalBytes,
                observed: usize_to_u64(observed),
                maximum: usize_to_u64(maximum),
            },
            _ => ChartValueAxisError::InvalidSource,
        };
    }
    if let Some(amount) = error.allocation_amount() {
        return ChartValueAxisError::Allocation { amount };
    }
    ChartValueAxisError::InvalidSource
}

fn map_read_error(error: ReadError) -> ChartValueAxisError {
    match error {
        ReadError::SemanticLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartValueAxisError::LimitExceeded {
            kind: match kind {
                SemanticLimitKind::Objects => ChartValueAxisLimitKind::Entries,
                SemanticLimitKind::Slides => ChartValueAxisLimitKind::Slides,
                SemanticLimitKind::References => ChartValueAxisLimitKind::References,
                SemanticLimitKind::TextStorages
                | SemanticLimitKind::TextFragments
                | SemanticLimitKind::TextBytes => ChartValueAxisLimitKind::Entries,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::PayloadLimit {
            kind,
            observed,
            maximum,
            ..
        } => ChartValueAxisError::LimitExceeded {
            kind: match kind {
                super::PayloadLimitKind::Bytes => ChartValueAxisLimitKind::WireBytes,
                super::PayloadLimitKind::Fields => ChartValueAxisLimitKind::WireFields,
                super::PayloadLimitKind::Nesting => ChartValueAxisLimitKind::WireNesting,
                super::PayloadLimitKind::Work => ChartValueAxisLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        ReadError::Allocation { amount, .. } => ChartValueAxisError::Allocation { amount },
        ReadError::Archive(_)
        | ReadError::Detection(_)
        | ReadError::NotKeynote
        | ReadError::InvalidFormat(_)
        | ReadError::Decode(_)
        | ReadError::TextStorage { .. }
        | ReadError::Metadata(_)
        | ReadError::Io(_) => ChartValueAxisError::InvalidSource,
    }
}

fn map_archive_error(error: litchi_iwa_archive::Error) -> ChartValueAxisError {
    match error {
        litchi_iwa_archive::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartValueAxisError::LimitExceeded {
            kind: match kind {
                litchi_iwa_archive::LimitKind::InputBytes => ChartValueAxisLimitKind::InputBytes,
                litchi_iwa_archive::LimitKind::OutputBytes => ChartValueAxisLimitKind::OutputBytes,
                litchi_iwa_archive::LimitKind::Entries => ChartValueAxisLimitKind::Entries,
                litchi_iwa_archive::LimitKind::EntryBytes => ChartValueAxisLimitKind::EntryBytes,
                litchi_iwa_archive::LimitKind::TotalBytes => ChartValueAxisLimitKind::TotalBytes,
                _ => ChartValueAxisLimitKind::WireBytes,
            },
            observed,
            maximum,
        },
        litchi_iwa_archive::Error::Allocation { amount, .. } => {
            ChartValueAxisError::Allocation { amount }
        },
        _ => ChartValueAxisError::InvalidSource,
    }
}

fn map_core_error(error: litchi_iwa_core::Error) -> ChartValueAxisError {
    match error {
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => ChartValueAxisError::LimitExceeded {
            kind: match kind {
                litchi_iwa_core::LimitKind::ArchiveBytes => ChartValueAxisLimitKind::TotalBytes,
                litchi_iwa_core::LimitKind::Objects
                | litchi_iwa_core::LimitKind::Messages
                | litchi_iwa_core::LimitKind::MessagesPerObject
                | litchi_iwa_core::LimitKind::MetadataItems => ChartValueAxisLimitKind::Entries,
                litchi_iwa_core::LimitKind::ObjectBytes
                | litchi_iwa_core::LimitKind::MessageBytes
                | litchi_iwa_core::LimitKind::HeaderBytes
                | litchi_iwa_core::LimitKind::HeaderMemoryBytes
                | litchi_iwa_core::LimitKind::SnappyChunkBytes
                | litchi_iwa_core::LimitKind::SnappyStreamBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedChunkBytes
                | litchi_iwa_core::LimitKind::SnappyCompressedStreamBytes => {
                    ChartValueAxisLimitKind::WireBytes
                },
                litchi_iwa_core::LimitKind::HeaderFields => ChartValueAxisLimitKind::WireFields,
                litchi_iwa_core::LimitKind::HeaderNesting => ChartValueAxisLimitKind::WireNesting,
                litchi_iwa_core::LimitKind::SnappyFrames => ChartValueAxisLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(maximum),
        },
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            ChartValueAxisError::Allocation { amount: requested }
        },
        _ => ChartValueAxisError::InvalidSource,
    }
}

fn map_wire_error(error: litchi_iwa_common::Error) -> ChartValueAxisError {
    match error {
        litchi_iwa_common::Error::LimitExceeded {
            kind,
            observed,
            limit,
        } => ChartValueAxisError::LimitExceeded {
            kind: match kind {
                litchi_iwa_common::LimitKind::InputBytes => ChartValueAxisLimitKind::WireBytes,
                litchi_iwa_common::LimitKind::OutputBytes => ChartValueAxisLimitKind::OutputBytes,
                litchi_iwa_common::LimitKind::Fields
                | litchi_iwa_common::LimitKind::TableRows
                | litchi_iwa_common::LimitKind::TableColumns
                | litchi_iwa_common::LimitKind::TableCells
                | litchi_iwa_common::LimitKind::MaterializedCells => {
                    ChartValueAxisLimitKind::WireFields
                },
                litchi_iwa_common::LimitKind::Nesting => ChartValueAxisLimitKind::WireNesting,
                litchi_iwa_common::LimitKind::RewriteWork => ChartValueAxisLimitKind::WireWork,
            },
            observed: usize_to_u64(observed),
            maximum: usize_to_u64(limit),
        },
        litchi_iwa_common::Error::Allocation { amount, .. } => {
            ChartValueAxisError::Allocation { amount }
        },
        litchi_iwa_common::Error::InvalidFormat(_)
        | litchi_iwa_common::Error::InvalidLimit { .. } => ChartValueAxisError::InvalidSource,
    }
}

fn usize_to_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn physical_catalog(package: &Package) -> Result<&SourceCatalog, ChartValueAxisError> {
    super::chart_axis_support::physical_catalog(package).map_err(map_support_error)
}

impl super::chart_axis_support::AxisSupportBudget for ValueAxisBudget {
    fn charge_selection_scans(
        &mut self,
        package: &Package,
        mutation_guards: bool,
    ) -> Result<(), super::chart_axis_support::AxisSupportError> {
        self.scan_pass(package, 2).map_err(map_budget_error)?;
        if mutation_guards {
            self.scan_pass(package, 2).map_err(map_budget_error)?;
        }
        Ok(())
    }

    fn charge_input(
        &mut self,
        amount: usize,
    ) -> Result<(), super::chart_axis_support::AxisSupportError> {
        self.input(amount).map_err(map_budget_error)
    }

    fn charge_wire_vector(
        &mut self,
        payload: usize,
    ) -> Result<(), super::chart_axis_support::AxisSupportError> {
        use super::chart_axis_support::AxisSupportError;
        let bytes = payload
            .checked_mul(size_of::<WireField>())
            .ok_or(AxisSupportError::InvalidSource)?;
        self.input(payload).map_err(map_budget_error)?;
        self.work(payload).map_err(map_budget_error)?;
        // parse_wire_fields_with_limits grows this vector one element at a
        // time. Charge a conservative event per payload-sized capacity while
        // retaining the exposed WireField byte shape.
        self.allocations(payload.max(1), bytes, bytes)
            .map_err(map_budget_error)
    }

    fn finish_wire_scan(
        &mut self,
        fields: usize,
    ) -> Result<(), super::chart_axis_support::AxisSupportError> {
        self.fields(fields).map_err(map_budget_error)?;
        self.work(fields).map_err(map_budget_error)
    }

    fn charge_reference_vector(
        &mut self,
        capacity: usize,
    ) -> Result<(), super::chart_axis_support::AxisSupportError> {
        use super::chart_axis_support::AxisSupportError;
        let bytes = capacity
            .checked_mul(size_of::<u64>())
            .ok_or(AxisSupportError::InvalidSource)?;
        self.allocations(1, bytes, bytes).map_err(map_budget_error)
    }

    fn charge_references(
        &mut self,
        amount: usize,
    ) -> Result<(), super::chart_axis_support::AxisSupportError> {
        self.references(amount).map_err(map_budget_error)
    }

    fn charge_scan_pass(
        &mut self,
        package: &Package,
        retained_vectors: usize,
    ) -> Result<(), super::chart_axis_support::AxisSupportError> {
        self.scan_pass(package, retained_vectors)
            .map_err(map_budget_error)
    }

    fn charge_locality_scan(
        &mut self,
        package: &Package,
    ) -> Result<(), super::chart_axis_support::AxisSupportError> {
        self.scan_pass(package, 0).map_err(map_budget_error)
    }

    fn charge_work(
        &mut self,
        amount: usize,
    ) -> Result<(), super::chart_axis_support::AxisSupportError> {
        self.work(amount).map_err(map_budget_error)
    }

    fn metadata_options(
        &self,
        package: &Package,
    ) -> Result<
        litchi_iwa_protos::package_metadata_codec::RewriteOptions,
        super::chart_axis_support::AxisSupportError,
    > {
        let limits = package
            .wire_limits()
            .map_err(|_| super::chart_axis_support::AxisSupportError::InvalidSource)?;
        let remaining = |total: usize, used: usize| {
            total
                .checked_sub(used)
                .map(|value| value.min(limits.max_input_bytes()))
                .ok_or(super::chart_axis_support::AxisSupportError::InvalidSource)
        };
        Ok(
            litchi_iwa_protos::package_metadata_codec::RewriteOptions::new(
                remaining(self.maximum_input, self.input)?,
                remaining(self.maximum_output, self.output)?,
                remaining(self.maximum_fields, self.fields)?,
                remaining(self.maximum_work, self.work)?,
                u32::try_from(limits.max_nesting())
                    .map_err(|_| super::chart_axis_support::AxisSupportError::InvalidSource)?,
                self.maximum_references
                    .checked_sub(self.references)
                    .ok_or(super::chart_axis_support::AxisSupportError::InvalidSource)?,
                self.maximum_references
                    .checked_sub(self.references)
                    .ok_or(super::chart_axis_support::AxisSupportError::InvalidSource)?,
                self.maximum_allocations
                    .checked_sub(self.allocations)
                    .ok_or(super::chart_axis_support::AxisSupportError::InvalidSource)?,
            ),
        )
    }

    fn charge_metadata_report(
        &mut self,
        report: litchi_iwa_protos::package_metadata_codec::RewriteReport,
    ) -> Result<(), super::chart_axis_support::AxisSupportError> {
        self.input(report.input_bytes()).map_err(map_budget_error)?;
        self.output(report.output_bytes())
            .map_err(map_budget_error)?;
        self.fields(report.fields()).map_err(map_budget_error)?;
        self.work(report.work_bytes()).map_err(map_budget_error)?;
        self.allocations(
            report.allocations(),
            report.retained_bytes(),
            report.scratch_bytes(),
        )
        .map_err(map_budget_error)
    }
}

impl ValueAxisBudget {
    fn scan_pass(
        &mut self,
        package: &Package,
        retained_vectors: usize,
    ) -> Result<(), ChartValueAxisError> {
        let mut objects = 0usize;
        let mut messages = 0usize;
        let mut references = 0usize;
        let mut bytes = 0usize;
        for component in package.state.source.components().iter() {
            self.work(component.name().len())?;
            for object in &component.archive().objects {
                objects = objects
                    .checked_add(1)
                    .ok_or(ChartValueAxisError::InvalidSource)?;
                for (index, message) in object.messages.iter().enumerate() {
                    messages = messages
                        .checked_add(1)
                        .ok_or(ChartValueAxisError::InvalidSource)?;
                    bytes = bytes
                        .checked_add(message.data.len())
                        .ok_or(ChartValueAxisError::InvalidSource)?;
                    let info = object
                        .archive_info
                        .message_infos
                        .get(index)
                        .ok_or(ChartValueAxisError::InvalidSource)?;
                    references = references
                        .checked_add(info.object_references.len())
                        .and_then(|value| value.checked_add(info.data_references.len()))
                        .ok_or(ChartValueAxisError::InvalidSource)?;
                    for field in &info.field_infos {
                        references = references
                            .checked_add(field.object_references.len())
                            .and_then(|value| value.checked_add(field.data_references.len()))
                            .ok_or(ChartValueAxisError::InvalidSource)?;
                    }
                }
            }
        }
        self.charge_scan_counts(objects, messages, references, bytes)?;
        if retained_vectors != 0 {
            let vector_bytes = retained_vectors
                .checked_mul(objects.max(1))
                .and_then(|value| value.checked_mul(size_of::<u64>()))
                .ok_or(ChartValueAxisError::InvalidSource)?;
            self.allocations(retained_vectors, vector_bytes, vector_bytes)?;
        }
        Ok(())
    }
    fn charge_scan_counts(
        &mut self,
        objects: usize,
        messages: usize,
        references: usize,
        bytes: usize,
    ) -> Result<(), ChartValueAxisError> {
        self.work(
            objects
                .checked_add(messages)
                .and_then(|value| value.checked_add(bytes))
                .ok_or(ChartValueAxisError::InvalidSource)?,
        )?;
        self.references(references)?;
        self.allocations(
            objects
                .checked_add(messages)
                .ok_or(ChartValueAxisError::InvalidSource)?,
            0,
            0,
        )
    }
}
