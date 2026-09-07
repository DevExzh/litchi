//! Private canonical playback-build authoring for fresh slide audio and movies.
//!
//! The package owner allocates identities and owns the surrounding slide
//! transaction.  This adapter only turns those checked identities into the
//! two native playback objects and rewrites the slide-node scalar cache. All
//! payload bytes come from the neutral Buffa writer in `litchi-iwa-protos`.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The focused build adapter keeps input mapping and bounded error mapping adjacent."
)]

use std::mem::size_of;

use litchi_iwa_common::WireLimits;
use litchi_iwa_core::{ArchiveObject, LimitKind, RawMessage};
use litchi_iwa_protos::keynote_build_creation_codec as build_codec;
use litchi_iwa_protos::keynote_media_lifecycle_codec as lifecycle_codec;
use litchi_iwa_protos::keynote_media_lifecycle_codec::node_cache;

use super::{CreationBudget, CreationContext, CreationIds};
use crate::slide::audio::creation::{SlideAudioCreationError, SlideAudioCreationLimitKind};

const BUILD_MESSAGE_TYPE: u32 = 8;
const BUILD_CHUNK_MESSAGE_TYPE: u32 = 153;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum StartPlaybackKind {
    Audio,
    Movie,
}

enum StartPlaybackWrite {
    Audio(build_codec::StartAudioBuildWrite),
    Movie(build_codec::StartMovieBuildWrite),
}

/// Author the native build and chunk objects for a fresh audio-start event.
pub(super) fn prepare_start_audio_builds(
    ids: &CreationIds,
    context: &CreationContext,
    budget: &mut CreationBudget,
) -> Result<[ArchiveObject; 2], SlideAudioCreationError> {
    prepare_start_builds(ids, context, budget, StartPlaybackKind::Audio)
}

/// Author the native build and chunk objects for a fresh movie-start event.
pub(super) fn prepare_start_movie_builds(
    ids: &CreationIds,
    context: &CreationContext,
    budget: &mut CreationBudget,
) -> Result<[ArchiveObject; 2], SlideAudioCreationError> {
    prepare_start_builds(ids, context, budget, StartPlaybackKind::Movie)
}

fn prepare_start_builds(
    ids: &CreationIds,
    context: &CreationContext,
    budget: &mut CreationBudget,
    kind: StartPlaybackKind,
) -> Result<[ArchiveObject; 2], SlideAudioCreationError> {
    ids.validate()?;
    context.validate()?;

    let write = match kind {
        StartPlaybackKind::Audio => {
            StartPlaybackWrite::Audio(build_codec::StartAudioBuildWrite::new(
                ids.drawable,
                ids.build,
                ids.chunk,
                ids.build_uuid.lower(),
                ids.build_uuid.upper(),
                ids.random_number_seed,
            ))
        },
        StartPlaybackKind::Movie => {
            StartPlaybackWrite::Movie(build_codec::StartMovieBuildWrite::new(
                ids.drawable,
                ids.build,
                ids.chunk,
                ids.build_uuid.lower(),
                ids.build_uuid.upper(),
                ids.random_number_seed,
            ))
        },
    };
    let options = build_encode_options(&write, context.wire_limits, budget);
    let build = encode_start_build(&write, options).map_err(map_build_encode_error)?;
    charge_build_report(budget, build.report())?;
    // The build payload consumed part of the operation-wide ledger.  Rebuild
    // the policy before encoding its sibling so all residual byte, field,
    // work, and allocation ceilings are enforced by the second preflight.
    let options = build_encode_options(&write, context.wire_limits, budget);
    let chunk = encode_start_chunk(&write, options).map_err(map_build_encode_error)?;
    charge_build_report(budget, chunk.report())?;

    let build_payload = build.into_bytes();
    let chunk_payload = chunk.into_bytes();
    let mut build_object = ArchiveObject::new_with_limits(
        ids.build,
        vec![RawMessage {
            type_: BUILD_MESSAGE_TYPE,
            data: build_payload,
        }],
        context.archive_limits,
    )
    .map_err(map_archive_error)?;
    set_object_reference(&mut build_object, ids.drawable, budget)?;
    let mut chunk_object = ArchiveObject::new_with_limits(
        ids.chunk,
        vec![RawMessage {
            type_: BUILD_CHUNK_MESSAGE_TYPE,
            data: chunk_payload,
        }],
        context.archive_limits,
    )
    .map_err(map_archive_error)?;
    set_object_reference(&mut chunk_object, ids.build, budget)?;
    Ok([build_object, chunk_object])
}

fn encode_start_build(
    write: &StartPlaybackWrite,
    options: build_codec::EncodeOptions,
) -> Result<build_codec::EncodeOutput, build_codec::EncodeError> {
    match write {
        StartPlaybackWrite::Audio(write) => build_codec::encode_start_audio_build(write, options),
        StartPlaybackWrite::Movie(write) => build_codec::encode_start_movie_build(write, options),
    }
}

fn encode_start_chunk(
    write: &StartPlaybackWrite,
    options: build_codec::EncodeOptions,
) -> Result<build_codec::EncodeOutput, build_codec::EncodeError> {
    match write {
        StartPlaybackWrite::Audio(write) => build_codec::encode_start_audio_chunk(write, options),
        StartPlaybackWrite::Movie(write) => build_codec::encode_start_movie_chunk(write, options),
    }
}

fn set_object_reference(
    object: &mut ArchiveObject,
    reference: u64,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let info = object
        .archive_info
        .message_infos
        .get_mut(0)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let allocation_bytes = size_of::<u64>();
    budget.charge_references(1)?;
    budget.charge_allocations(allocation_bytes)?;
    budget.charge_allocation_events(1)?;
    info.object_references.try_reserve_exact(1).map_err(|_| {
        SlideAudioCreationError::Allocation {
            amount: allocation_bytes,
        }
    })?;
    info.object_references.push(reference);
    Ok(())
}

/// Rewrite the slide-node cache using Keynote's fresh-build event-count
/// behavior while retaining every unrelated source field byte-for-byte.
pub(super) fn rewrite_node_cache(
    source: &[u8],
    event_count: usize,
    limits: WireLimits,
    budget: &mut CreationBudget,
) -> Result<Vec<u8>, SlideAudioCreationError> {
    let event_count =
        u32::try_from(event_count).map_err(|_| SlideAudioCreationError::LimitExceeded {
            kind: SlideAudioCreationLimitKind::WireFields,
            observed: u64::try_from(event_count).unwrap_or(u64::MAX),
            maximum: u64::from(u32::MAX),
        })?;
    let options = node_cache_options(source, limits, budget);
    let (output, report) = node_cache::rewrite_slide_node_build_cache_for_event_count_with_report(
        source,
        event_count,
        options,
    )
    .map_err(map_node_cache_error)?;
    budget.charge_output(report.output_bytes())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocation_events(report.allocations())?;
    Ok(output)
}

fn build_encode_options(
    write: &StartPlaybackWrite,
    limits: WireLimits,
    budget: &CreationBudget,
) -> build_codec::EncodeOptions {
    let options = match write {
        StartPlaybackWrite::Audio(write) => build_codec::EncodeOptions::for_write(write),
        StartPlaybackWrite::Movie(write) => build_codec::EncodeOptions::for_movie_write(write),
    };
    options
        .with_max_output_bytes(limits.max_output_bytes().min(budget.remaining_output()))
        .with_max_fields(limits.max_fields().min(budget.remaining_wire_fields()))
        .with_max_work_bytes(limits.max_rewrite_work().min(budget.remaining_work()))
        // The Buffa writer accounts for nested message views plus the output
        // vector: five events is the chunk upper bound.  The shared ledger's
        // remainder keeps this local preflight from allocating after the
        // operation-wide allocation ceiling has been consumed.
        .with_max_allocations(5.min(budget.remaining_allocations()))
}

fn node_cache_options(
    source: &[u8],
    limits: WireLimits,
    budget: &CreationBudget,
) -> lifecycle_codec::DecodeOptions {
    let max_message_bytes = source.len().min(limits.max_input_bytes()).max(1);
    let max_output_bytes = limits.max_output_bytes().min(budget.remaining_output());
    let max_fields = limits.max_fields().min(budget.remaining_wire_fields());
    let max_work_bytes = limits.max_rewrite_work().min(budget.remaining_work());
    let max_depth = u32::try_from(limits.max_nesting()).unwrap_or(u32::MAX);
    lifecycle_codec::DecodeOptions::new(
        max_message_bytes,
        max_output_bytes,
        max_fields,
        max_work_bytes,
        1,
        max_depth,
    )
}

fn charge_build_report(
    budget: &mut CreationBudget,
    report: build_codec::EncodeReport,
) -> Result<(), SlideAudioCreationError> {
    budget.charge_output(report.output_bytes())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_allocation_events(report.allocations())
}

fn map_build_encode_error(error: build_codec::EncodeError) -> SlideAudioCreationError {
    match error {
        build_codec::EncodeError::InvalidInput(_) | build_codec::EncodeError::Verification => {
            SlideAudioCreationError::InvalidSource
        },
        build_codec::EncodeError::Allocation { amount } => {
            SlideAudioCreationError::Allocation { amount }
        },
        build_codec::EncodeError::Buffa(_) => SlideAudioCreationError::InvalidSource,
        build_codec::EncodeError::Resource(limit) => {
            let (kind, observed, maximum) = match limit {
                build_codec::EncodeLimit::OutputBytes { observed, maximum } => {
                    (SlideAudioCreationLimitKind::OutputBytes, observed, maximum)
                },
                build_codec::EncodeLimit::Fields { observed, maximum } => {
                    (SlideAudioCreationLimitKind::WireFields, observed, maximum)
                },
                build_codec::EncodeLimit::WorkBytes { observed, maximum } => {
                    (SlideAudioCreationLimitKind::WireWork, observed, maximum)
                },
                build_codec::EncodeLimit::Allocations { observed, maximum } => {
                    (SlideAudioCreationLimitKind::Allocations, observed, maximum)
                },
                _ => return SlideAudioCreationError::InvalidSource,
            };
            limit_error(kind, observed, maximum)
        },
        _ => SlideAudioCreationError::InvalidSource,
    }
}

fn map_node_cache_error(error: lifecycle_codec::DecodeError) -> SlideAudioCreationError {
    let Some(limit) = error.resource_limit() else {
        return SlideAudioCreationError::InvalidSource;
    };
    let (kind, observed, maximum) = match limit {
        lifecycle_codec::DecodeLimit::Bytes { observed, maximum }
        | lifecycle_codec::DecodeLimit::Work { observed, maximum } => {
            (SlideAudioCreationLimitKind::WireWork, observed, maximum)
        },
        lifecycle_codec::DecodeLimit::Fields { observed, maximum } => {
            (SlideAudioCreationLimitKind::WireFields, observed, maximum)
        },
        lifecycle_codec::DecodeLimit::OutputBytes { observed, maximum } => {
            (SlideAudioCreationLimitKind::OutputBytes, observed, maximum)
        },
        lifecycle_codec::DecodeLimit::References { observed, maximum } => {
            (SlideAudioCreationLimitKind::References, observed, maximum)
        },
        lifecycle_codec::DecodeLimit::Nesting { observed, maximum } => (
            SlideAudioCreationLimitKind::WireNesting,
            observed as usize,
            maximum as usize,
        ),
        _ => return SlideAudioCreationError::InvalidSource,
    };
    limit_error(kind, observed, maximum)
}

fn map_archive_error(error: litchi_iwa_core::Error) -> SlideAudioCreationError {
    match error {
        litchi_iwa_core::Error::Allocation { requested, .. } => {
            SlideAudioCreationError::Allocation { amount: requested }
        },
        litchi_iwa_core::Error::Limit {
            kind,
            observed,
            maximum,
        } => {
            let kind = match kind {
                LimitKind::Messages | LimitKind::MessagesPerObject | LimitKind::MetadataItems => {
                    SlideAudioCreationLimitKind::Entries
                },
                LimitKind::Objects => SlideAudioCreationLimitKind::Objects,
                LimitKind::ArchiveBytes
                | LimitKind::ObjectBytes
                | LimitKind::MessageBytes
                | LimitKind::SnappyChunkBytes
                | LimitKind::SnappyStreamBytes
                | LimitKind::SnappyCompressedChunkBytes
                | LimitKind::SnappyCompressedStreamBytes => SlideAudioCreationLimitKind::EntryBytes,
                LimitKind::HeaderBytes
                | LimitKind::HeaderFields
                | LimitKind::HeaderNesting
                | LimitKind::HeaderMemoryBytes
                | LimitKind::SnappyFrames => SlideAudioCreationLimitKind::WireFields,
            };
            limit_error(kind, observed, maximum)
        },
        _ => SlideAudioCreationError::InvalidSource,
    }
}

fn limit_error(
    kind: SlideAudioCreationLimitKind,
    observed: usize,
    maximum: usize,
) -> SlideAudioCreationError {
    SlideAudioCreationError::LimitExceeded {
        kind,
        observed: u64::try_from(observed).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    }
}
