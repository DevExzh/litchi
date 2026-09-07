//! Bounded post-reopen witness for a fresh slide-owned audio graph.
//!
//! The creation transaction writes several independent native graphs before
//! assembling the ZIP.  This module follows those graphs again from the
//! reopened candidate and checks the exact identities reserved by the
//! transaction.  It deliberately uses the neutral lifecycle and metadata
//! codecs plus small raw-wire projections; generated native messages are not
//! used as an ingress or verification representation.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The witness keeps each bounded graph phase adjacent to its source checks."
)]

use std::mem::size_of;

use litchi_iwa_common::{WireLimits, wire::WireFieldView, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject};
use litchi_iwa_protos::{
    keynote_build_creation_codec as build_creation_codec, keynote_media_codec,
    keynote_media_creation_codec as media_creation_codec,
    keynote_media_lifecycle_codec as lifecycle_codec, keynote_movie_caption_codec,
    package_metadata_codec as identity_codec, package_metadata_media_codec as media_codec,
};

use crate::soundtrack::items::MAX_FILENAME_BYTES;

use super::{
    CreationBudget, CreationContext, CreationIds, MEDIA_STYLE_MESSAGE_TYPE, METADATA_COMPONENT,
    MOVIE_MESSAGE_TYPE, PACKAGE_METADATA_MESSAGE_TYPE, Package, SLIDE_MESSAGE_TYPE,
    SLIDE_NODE_MESSAGE_TYPE, STANDIN_CAPTION_MESSAGE_TYPE, STYLESHEET_MESSAGE_TYPE,
};
use crate::slide::audio::creation::{SlideAudioCreationError, SlideAudioCreationLimitKind};

const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const STANDIN_CAPTION_MESSAGE_VERSION: [u32; 3] = [10, 1, 0];
const BUILD_MESSAGE_TYPE: u32 = 8;
const BUILD_CHUNK_MESSAGE_TYPE: u32 = 153;

/// The media graph expected by the bounded fresh-creation witness.
///
/// Audio keeps the historical contract used by [`verify_created_graph`]. A
/// movie carries both data edges and the scalar values that the neutral
/// Buffa media writer emits. Keeping those values here lets the verifier
/// compare the reopened payload with a typed projection instead of accepting
/// a payload merely because a few required fields happen to decode.
#[derive(Debug, Clone, Copy, PartialEq)]
pub(in crate::package) struct CreatedMediaExpectation {
    content_data_identifier: u64,
    kind: CreatedMediaKind,
}

#[derive(Debug, Clone, Copy, PartialEq)]
enum CreatedMediaKind {
    Audio,
    Movie {
        poster_data_identifier: u64,
        geometry: CreatedGeometry,
        duration_seconds: f32,
        natural_size: CreatedSize,
        poster_image_generated_with_alpha_support: bool,
    },
}

#[derive(Debug, Clone, Copy, PartialEq)]
struct CreatedGeometry {
    position: [f32; 2],
    size: [f32; 2],
    flags: Option<u32>,
    angle: Option<f32>,
}

#[derive(Debug, Clone, Copy, PartialEq)]
struct CreatedSize {
    width: f32,
    height: f32,
}

impl CreatedMediaExpectation {
    /// Preserve the original fresh-audio witness contract.
    pub(super) const fn audio(content_data_identifier: u64) -> Self {
        Self {
            content_data_identifier,
            kind: CreatedMediaKind::Audio,
        }
    }

    /// Describe one fresh file movie with its video and poster records.
    pub(in crate::package) const fn movie(
        content_data_identifier: u64,
        poster_data_identifier: u64,
        position: [f32; 2],
        size: [f32; 2],
        duration_seconds: f32,
        natural_size: [f32; 2],
        poster_image_generated_with_alpha_support: bool,
        flags: Option<u32>,
        angle: Option<f32>,
    ) -> Self {
        Self {
            content_data_identifier,
            kind: CreatedMediaKind::Movie {
                poster_data_identifier,
                geometry: CreatedGeometry {
                    position,
                    size,
                    flags,
                    angle,
                },
                duration_seconds,
                natural_size: CreatedSize {
                    width: natural_size[0],
                    height: natural_size[1],
                },
                poster_image_generated_with_alpha_support,
            },
        }
    }

    fn validate(self) -> Result<(), SlideAudioCreationError> {
        if self.content_data_identifier == 0 {
            return Err(SlideAudioCreationError::Verification);
        }
        if let CreatedMediaKind::Movie {
            poster_data_identifier,
            geometry,
            duration_seconds,
            natural_size,
            ..
        } = self.kind
        {
            if poster_data_identifier == 0
                || poster_data_identifier == self.content_data_identifier
                || !duration_seconds.is_finite()
                || duration_seconds < 0.0
                || geometry.position.iter().any(|value| !value.is_finite())
                || geometry
                    .size
                    .iter()
                    .any(|value| !value.is_finite() || *value < 0.0)
                || geometry.angle.is_some_and(|value| !value.is_finite())
                || !natural_size.width.is_finite()
                || !natural_size.height.is_finite()
                || natural_size.width < 0.0
                || natural_size.height < 0.0
            {
                return Err(SlideAudioCreationError::Verification);
            }
        }
        Ok(())
    }

    fn poster_data_identifier(self) -> Option<u64> {
        match self.kind {
            CreatedMediaKind::Audio => None,
            CreatedMediaKind::Movie {
                poster_data_identifier,
                ..
            } => Some(poster_data_identifier),
        }
    }

    fn build_effect(self) -> BuildEffect {
        match self.kind {
            CreatedMediaKind::Audio => BuildEffect::Audio,
            CreatedMediaKind::Movie { .. } => BuildEffect::Movie,
        }
    }

    fn data_references(self) -> [u64; 2] {
        [
            self.content_data_identifier,
            self.poster_data_identifier().unwrap_or(0),
        ]
    }

    fn header_data_references(self) -> [u64; 2] {
        match self.poster_data_identifier() {
            Some(poster) => [poster, self.content_data_identifier],
            None => [self.content_data_identifier, 0],
        }
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum BuildEffect {
    Audio,
    Movie,
}

/// Verify the complete graph produced by one fresh slide-audio publication.
///
/// This witness is intentionally used for the initial candidate assembled by
/// [`Package::add_slide_audio`].  A replay of an already-authorized exact
/// artifact does not need to retain these private native identifiers: the
/// artifact fingerprint and the package-level semantic checks are its source
/// of truth.
pub(super) fn verify_created_graph(
    package: &Package,
    context: &CreationContext,
    ids: &CreationIds,
    data_identifier: u64,
    expected_event_count: usize,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    verify_created_graph_with_media(
        package,
        context,
        ids,
        CreatedMediaExpectation::audio(data_identifier),
        expected_event_count,
        budget,
    )
}

/// Verify a freshly created audio or file-movie graph after it has been
/// serialized and reopened.
///
/// The caller supplies the typed media edges and movie scalars selected by
/// its creation planner.  This keeps the verifier independent from the
/// planner's data-member and build adapters while making the candidate graph
/// check exact for movies (including the poster edge and movie-start build).
pub(in crate::package) fn verify_created_graph_with_media(
    package: &Package,
    context: &CreationContext,
    ids: &CreationIds,
    media: CreatedMediaExpectation,
    expected_event_count: usize,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    context.validate()?;
    ids.validate()?;
    media.validate()?;

    verify_selected_style(package, context, budget)?;
    verify_slide_lifecycle(package, context, ids, expected_event_count, budget)?;
    verify_created_objects(package, context, ids, media, budget)?;
    verify_metadata(package, context, ids, media, budget)?;
    verify_node_cache(package, context, expected_event_count, budget)?;
    Ok(())
}

fn verify_selected_style(
    package: &Package,
    context: &CreationContext,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let style_object = verify_object(
        package,
        context.stylesheet_component.as_ref(),
        context.style_identifier,
        MEDIA_STYLE_MESSAGE_TYPE,
    )?;
    let style_payload = unique_message_payload(style_object, MEDIA_STYLE_MESSAGE_TYPE)?;
    charge_payload(style_payload, context.wire_limits, budget)?;

    let (stylesheet_component, stylesheet_object) = package
        .object_with_component(context.stylesheet_identifier)
        .ok_or(SlideAudioCreationError::Verification)?;
    if stylesheet_component != context.stylesheet_component.as_ref() {
        return Err(SlideAudioCreationError::Verification);
    }
    let stylesheet_payload =
        message_payload_with_header(stylesheet_object, STYLESHEET_MESSAGE_TYPE)?;
    let style_references = references_in_field(stylesheet_payload, 1, context.wire_limits, budget)?;
    let selected_style = style_references.iter().copied().find(|identifier| {
        package.object(*identifier).is_some_and(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == MEDIA_STYLE_MESSAGE_TYPE)
        })
    });
    if selected_style != Some(context.style_identifier) {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(())
}

fn verify_slide_lifecycle(
    package: &Package,
    context: &CreationContext,
    ids: &CreationIds,
    expected_event_count: usize,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let (component_name, slide_object) = package
        .object_with_component(context.slide_identifier)
        .ok_or(SlideAudioCreationError::Verification)?;
    if component_name != context.slide_component.as_ref() {
        return Err(SlideAudioCreationError::Verification);
    }
    let slide_payload = unique_message_payload(slide_object, SLIDE_MESSAGE_TYPE)?;
    let options = lifecycle_options(slide_payload.len(), context.wire_limits, budget)?;
    let (snapshot, report) =
        lifecycle_codec::decode_slide_lifecycle_with_report(slide_payload, options)
            .map_err(|_| SlideAudioCreationError::Verification)?;
    charge_lifecycle_report(report, budget)?;

    let owned = collect_reference_ids(snapshot.owned_drawables(), budget)?;
    let z_order = collect_reference_ids(snapshot.drawables_z_order(), budget)?;
    let builds = collect_reference_ids(snapshot.builds(), budget)?;
    let chunks = collect_reference_ids(snapshot.build_chunks(), budget)?;
    if expected_event_count == 0
        || chunks.len() != expected_event_count
        || !contains_once(&owned, ids.drawable)
        || !contains_once(&z_order, ids.drawable)
        || !contains_once(&builds, ids.build)
        || !contains_once(&chunks, ids.chunk)
    {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(())
}

fn verify_created_objects(
    package: &Package,
    context: &CreationContext,
    ids: &CreationIds,
    media: CreatedMediaExpectation,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let movie = verify_object(
        package,
        context.slide_component.as_ref(),
        ids.drawable,
        MOVIE_MESSAGE_TYPE,
    )?;
    let title = verify_object(
        package,
        context.slide_component.as_ref(),
        ids.title,
        STANDIN_CAPTION_MESSAGE_TYPE,
    )?;
    let caption = verify_object(
        package,
        context.slide_component.as_ref(),
        ids.caption,
        STANDIN_CAPTION_MESSAGE_TYPE,
    )?;
    let build = verify_object(
        package,
        context.slide_component.as_ref(),
        ids.build,
        BUILD_MESSAGE_TYPE,
    )?;
    let chunk = verify_object(
        package,
        context.slide_component.as_ref(),
        ids.chunk,
        BUILD_CHUNK_MESSAGE_TYPE,
    )?;

    let data_references = media.header_data_references();
    let data_reference_count = 1 + usize::from(media.poster_data_identifier().is_some());
    verify_message_header(
        movie,
        &STANDARD_MESSAGE_VERSION,
        &[ids.caption, ids.title, context.style_identifier],
        &data_references[..data_reference_count],
    )?;
    verify_message_header(title, &STANDIN_CAPTION_MESSAGE_VERSION, &[], &[])?;
    verify_message_header(caption, &STANDIN_CAPTION_MESSAGE_VERSION, &[], &[])?;
    verify_message_header(build, &STANDARD_MESSAGE_VERSION, &[ids.drawable], &[])?;
    verify_message_header(chunk, &STANDARD_MESSAGE_VERSION, &[ids.build], &[])?;

    let movie_payload = unique_message_payload(movie, MOVIE_MESSAGE_TYPE)?;
    verify_movie_payload(movie_payload, context, ids, media, budget)?;
    if !unique_message_payload(title, STANDIN_CAPTION_MESSAGE_TYPE)?.is_empty()
        || !unique_message_payload(caption, STANDIN_CAPTION_MESSAGE_TYPE)?.is_empty()
    {
        return Err(SlideAudioCreationError::Verification);
    }
    verify_build_payload(build, ids, context, media.build_effect(), budget)?;
    verify_chunk_payload(chunk, ids, context, media.build_effect(), budget)?;
    Ok(())
}

fn verify_movie_payload(
    payload: &[u8],
    context: &CreationContext,
    ids: &CreationIds,
    media: CreatedMediaExpectation,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let options = movie_caption_options(payload.len(), context.wire_limits, budget)?;
    let (caption_snapshot, report) =
        keynote_movie_caption_codec::decode_movie_caption_with_report(payload, options)
            .map_err(|_| SlideAudioCreationError::Verification)?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    if !caption_snapshot.has_drawable()
        || caption_snapshot.title_identifier() != Some(ids.title)
        || caption_snapshot.caption_identifier() != Some(ids.caption)
    {
        return Err(SlideAudioCreationError::Verification);
    }

    let root = wire_view(payload, context.wire_limits, budget)?;
    let drawable_payload = required_bytes(&root, 1)?;
    let parent_payload = required_nested_field(drawable_payload, 2, context.wire_limits, budget)?;
    let parent_identifier = reference_identifier(parent_payload, context.wire_limits, budget)?;
    if parent_identifier != context.slide_identifier {
        return Err(SlideAudioCreationError::Verification);
    }
    let data_payload = required_bytes(&root, 14)?;
    let data_options = keynote_media_codec::DecodeOptions::new(
        residual(
            data_payload
                .len()
                .min(context.wire_limits.max_input_bytes()),
            SlideAudioCreationLimitKind::InputBytes,
        )?,
        residual(
            context
                .wire_limits
                .max_fields()
                .min(budget.remaining_wire_fields()),
            SlideAudioCreationLimitKind::WireFields,
        )?,
        residual(
            context
                .wire_limits
                .max_rewrite_work()
                .min(budget.remaining_work()),
            SlideAudioCreationLimitKind::WireWork,
        )?,
        u32::try_from(context.wire_limits.max_nesting()).unwrap_or(u32::MAX),
    );
    let (data_snapshot, data_report) =
        keynote_media_codec::decode_data_reference_with_report(data_payload, data_options)
            .map_err(|_| SlideAudioCreationError::Verification)?;
    budget.charge_wire_fields(data_report.fields())?;
    budget.charge_work(data_report.work_bytes())?;
    budget.charge_nesting(data_report.max_depth() as usize)?;
    if data_snapshot.identifier() != media.content_data_identifier {
        return Err(SlideAudioCreationError::Verification);
    }
    match media.poster_data_identifier() {
        Some(expected_identifier) => {
            let poster_payload = required_bytes(&root, 15)?;
            let poster_options = keynote_media_codec::DecodeOptions::new(
                residual(
                    poster_payload
                        .len()
                        .min(context.wire_limits.max_input_bytes()),
                    SlideAudioCreationLimitKind::InputBytes,
                )?,
                residual(
                    context
                        .wire_limits
                        .max_fields()
                        .min(budget.remaining_wire_fields()),
                    SlideAudioCreationLimitKind::WireFields,
                )?,
                residual(
                    context
                        .wire_limits
                        .max_rewrite_work()
                        .min(budget.remaining_work()),
                    SlideAudioCreationLimitKind::WireWork,
                )?,
                u32::try_from(context.wire_limits.max_nesting()).unwrap_or(u32::MAX),
            );
            let (poster_snapshot, poster_report) =
                keynote_media_codec::decode_data_reference_with_report(
                    poster_payload,
                    poster_options,
                )
                .map_err(|_| SlideAudioCreationError::Verification)?;
            budget.charge_wire_fields(poster_report.fields())?;
            budget.charge_work(poster_report.work_bytes())?;
            budget.charge_nesting(poster_report.max_depth() as usize)?;
            if poster_snapshot.identifier() != expected_identifier {
                return Err(SlideAudioCreationError::Verification);
            }
        },
        None => {
            if root.fields().any(|field| field.number() == 15) {
                return Err(SlideAudioCreationError::Verification);
            }
        },
    }
    let style_payload = required_bytes(&root, 19)?;
    if reference_identifier(style_payload, context.wire_limits, budget)? != context.style_identifier
    {
        return Err(SlideAudioCreationError::Verification);
    }
    if let CreatedMediaKind::Movie {
        geometry,
        duration_seconds,
        natural_size,
        poster_image_generated_with_alpha_support,
        poster_data_identifier,
    } = media.kind
    {
        verify_movie_payload_projection(
            &root,
            context,
            ids,
            media.content_data_identifier,
            poster_data_identifier,
            geometry,
            duration_seconds,
            natural_size,
            poster_image_generated_with_alpha_support,
            budget,
        )?;
    }
    Ok(())
}

/// Compare every canonical field emitted by the neutral movie writer while
/// admitting Keynote's optional video descriptor edge.  The descriptor is
/// itself checked below when a source or native writer includes it; arbitrary
/// extra fields are never silently accepted.
fn verify_movie_payload_projection(
    actual: &WireView<'_>,
    context: &CreationContext,
    ids: &CreationIds,
    content_data_identifier: u64,
    poster_data_identifier: u64,
    geometry: CreatedGeometry,
    duration_seconds: f32,
    natural_size: CreatedSize,
    poster_image_generated_with_alpha_support: bool,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let write = media_creation_codec::MediaArchiveWrite::movie(
        context.slide_identifier,
        context.style_identifier,
        ids.title,
        ids.caption,
        content_data_identifier,
        Some(poster_data_identifier),
        media_creation_codec::Geometry::new(
            media_creation_codec::Point::new(geometry.position[0], geometry.position[1]),
            media_creation_codec::Size::new(geometry.size[0], geometry.size[1]),
            geometry.flags,
            geometry.angle,
        ),
        duration_seconds,
        media_creation_codec::Size::new(natural_size.width, natural_size.height),
        poster_image_generated_with_alpha_support,
    );
    let options = media_creation_codec::EncodeOptions::for_write(&write)
        .with_max_output_bytes(residual(
            budget.remaining_output(),
            SlideAudioCreationLimitKind::OutputBytes,
        )?)
        .with_max_references(residual(
            budget.remaining_references(),
            SlideAudioCreationLimitKind::References,
        )?)
        .with_max_fields(residual(
            budget.remaining_wire_fields(),
            SlideAudioCreationLimitKind::WireFields,
        )?)
        .with_max_work_bytes(residual(
            budget.remaining_work(),
            SlideAudioCreationLimitKind::WireWork,
        )?)
        .with_max_allocations(residual(
            budget.remaining_allocations(),
            SlideAudioCreationLimitKind::Allocations,
        )?);
    let expected = media_creation_codec::encode_media_archive_with_report(&write, options)
        .map_err(|_| SlideAudioCreationError::Verification)?;
    let report = expected.report();
    budget.charge_output(report.output_bytes())?;
    budget.charge_references(report.references())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_allocation_events(report.allocations())?;
    let expected_view = wire_view(expected.bytes(), context.wire_limits, budget)?;

    let mut actual_fields = actual.fields().peekable();
    let mut descriptor_seen = false;
    for expected_field in expected_view.fields() {
        while actual_fields
            .peek()
            .is_some_and(|field| field.number() == 29 && expected_field.number() != 29)
        {
            let descriptor = actual_fields
                .next()
                .ok_or(SlideAudioCreationError::Verification)?;
            if descriptor_seen {
                return Err(SlideAudioCreationError::Verification);
            }
            descriptor_seen = true;
            verify_video_descriptor(descriptor, context.wire_limits, natural_size, budget)?;
        }
        let actual_field = actual_fields
            .next()
            .ok_or(SlideAudioCreationError::Verification)?;
        if actual_field.raw() != expected_field.raw() {
            return Err(SlideAudioCreationError::Verification);
        }
    }
    for descriptor in actual_fields {
        if descriptor_seen {
            return Err(SlideAudioCreationError::Verification);
        }
        descriptor_seen = true;
        verify_video_descriptor(descriptor, context.wire_limits, natural_size, budget)?;
    }
    Ok(())
}

/// Check the optional `TSD.MovieArchive.video` descriptor without decoding it
/// through the generated native model. The fields are the stable, bounded
/// projection observed in Keynote's file-movie graph: media kind, dimensions,
/// digest/identity bytes, flags, transform scalars, and the native unit tag.
fn verify_video_descriptor(
    field: WireFieldView<'_>,
    limits: WireLimits,
    natural_size: CreatedSize,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    if field.number() != 29 || field.wire_type() != 2 {
        return Err(SlideAudioCreationError::Verification);
    }
    let descriptor = wire_view(field.payload(), limits, budget)?;
    if descriptor.len() != 2 {
        return Err(SlideAudioCreationError::Verification);
    }
    let video_payload = required_bytes(&descriptor, 1)?;
    let marker = required_bytes(&descriptor, 2)?;
    if marker != [1, 0, 0] {
        return Err(SlideAudioCreationError::Verification);
    }
    let video = wire_view(video_payload, limits, budget)?;
    if video.len() != 19 {
        return Err(SlideAudioCreationError::Verification);
    }
    if required_text(&video, 1)? != b"vide"
        || required_varint(&video, 2)? != 1
        || required_varint(&video, 3)? == 0
        || required_bytes(&video, 4)?.len() != 40
        || required_varint(&video, 5)? != 0
        || required_varint(&video, 6)? != 1
        || required_varint(&video, 7)? != 1
        || required_varint(&video, 8)? == 0
        || required_varint(&video, 9)? == 0
        || required_varint(&video, 10)? != 1
        || required_text(&video, 19)? != b"und"
    {
        return Err(SlideAudioCreationError::Verification);
    }
    let size_payload = required_bytes(&video, 11)?;
    let size = wire_view(size_payload, limits, budget)?;
    if size.len() != 2
        || f32::from_bits(required_fixed32(&size, 1)?) != natural_size.width
        || f32::from_bits(required_fixed32(&size, 2)?) != natural_size.height
    {
        return Err(SlideAudioCreationError::Verification);
    }
    for number in 12..=18 {
        let value = unique_wire_field(&video, number)?;
        if value.wire_type() != 1 || value.payload().len() != 8 {
            return Err(SlideAudioCreationError::Verification);
        }
    }
    Ok(())
}

fn verify_build_payload(
    object: &ArchiveObject,
    ids: &CreationIds,
    context: &CreationContext,
    effect: BuildEffect,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let payload = unique_message_payload(object, BUILD_MESSAGE_TYPE)?;
    let options = lifecycle_options(payload.len(), context.wire_limits, budget)?;
    let (snapshot, report) = lifecycle_codec::decode_build_with_report(payload, options)
        .map_err(|_| SlideAudioCreationError::Verification)?;
    charge_lifecycle_report(report, budget)?;
    if snapshot.drawable().map(|reference| reference.identifier()) != Some(ids.drawable) {
        return Err(SlideAudioCreationError::Verification);
    }

    let expected = match effect {
        BuildEffect::Audio => {
            let write = build_creation_codec::StartAudioBuildWrite::new(
                ids.drawable,
                ids.build,
                ids.chunk,
                ids.build_uuid.lower(),
                ids.build_uuid.upper(),
                ids.random_number_seed,
            );
            let options = build_options(
                build_creation_codec::EncodeOptions::for_write(&write),
                budget,
            )?;
            build_creation_codec::encode_start_audio_build(&write, options)
        },
        BuildEffect::Movie => {
            let write = build_creation_codec::StartMovieBuildWrite::new(
                ids.drawable,
                ids.build,
                ids.chunk,
                ids.build_uuid.lower(),
                ids.build_uuid.upper(),
                ids.random_number_seed,
            );
            let options = build_options(
                build_creation_codec::EncodeOptions::for_movie_write(&write),
                budget,
            )?;
            build_creation_codec::encode_start_movie_build(&write, options)
        },
    }
    .map_err(|_| SlideAudioCreationError::Verification)?;
    let report = expected.report();
    budget.charge_output(report.output_bytes())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_allocation_events(report.allocations())?;
    if expected.bytes() != payload {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(())
}

fn verify_chunk_payload(
    object: &ArchiveObject,
    ids: &CreationIds,
    context: &CreationContext,
    effect: BuildEffect,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let payload = unique_message_payload(object, BUILD_CHUNK_MESSAGE_TYPE)?;
    let options = lifecycle_options(payload.len(), context.wire_limits, budget)?;
    let (snapshot, report) = lifecycle_codec::decode_build_chunk_with_report(payload, options)
        .map_err(|_| SlideAudioCreationError::Verification)?;
    charge_lifecycle_report(report, budget)?;
    let uuid = lifecycle_codec::Uuid::new(ids.build_uuid.lower(), ids.build_uuid.upper());
    if snapshot.build().identifier() != ids.build
        || snapshot.chunk_identifier().map(|value| value.uuid()) != Some(uuid)
        || snapshot.build_id().map(|value| value.uuid()) != Some(uuid)
    {
        return Err(SlideAudioCreationError::Verification);
    }

    let expected = match effect {
        BuildEffect::Audio => {
            let write = build_creation_codec::StartAudioBuildWrite::new(
                ids.drawable,
                ids.build,
                ids.chunk,
                ids.build_uuid.lower(),
                ids.build_uuid.upper(),
                ids.random_number_seed,
            );
            let options = build_options(
                build_creation_codec::EncodeOptions::for_write(&write),
                budget,
            )?;
            build_creation_codec::encode_start_audio_chunk(&write, options)
        },
        BuildEffect::Movie => {
            let write = build_creation_codec::StartMovieBuildWrite::new(
                ids.drawable,
                ids.build,
                ids.chunk,
                ids.build_uuid.lower(),
                ids.build_uuid.upper(),
                ids.random_number_seed,
            );
            let options = build_options(
                build_creation_codec::EncodeOptions::for_movie_write(&write),
                budget,
            )?;
            build_creation_codec::encode_start_movie_chunk(&write, options)
        },
    }
    .map_err(|_| SlideAudioCreationError::Verification)?;
    let report = expected.report();
    budget.charge_output(report.output_bytes())?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_allocation_events(report.allocations())?;
    if expected.bytes() != payload {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(())
}

fn build_options(
    options: build_creation_codec::EncodeOptions,
    budget: &CreationBudget,
) -> Result<build_creation_codec::EncodeOptions, SlideAudioCreationError> {
    Ok(options
        .with_max_output_bytes(residual(
            budget.remaining_output(),
            SlideAudioCreationLimitKind::OutputBytes,
        )?)
        .with_max_fields(residual(
            budget.remaining_wire_fields(),
            SlideAudioCreationLimitKind::WireFields,
        )?)
        .with_max_work_bytes(residual(
            budget.remaining_work(),
            SlideAudioCreationLimitKind::WireWork,
        )?)
        .with_max_allocations(residual(
            budget.remaining_allocations(),
            SlideAudioCreationLimitKind::Allocations,
        )?))
}

fn verify_metadata(
    package: &Package,
    context: &CreationContext,
    ids: &CreationIds,
    media: CreatedMediaExpectation,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let metadata_archive = package
        .state
        .source
        .components()
        .get(METADATA_COMPONENT)
        .map(|component| component.archive())
        .ok_or(SlideAudioCreationError::Verification)?;
    let metadata_payload =
        unique_archive_message_payload(metadata_archive, PACKAGE_METADATA_MESSAGE_TYPE)?;
    let identity_options = identity_options(package, metadata_payload.len(), context, budget)?;
    let mut identity = IdentityWitness::new(context, ids);
    let inspection = identity_codec::inspect_package_metadata_with_visitor(
        metadata_payload,
        identity_options,
        &mut identity,
    )
    .map_err(|error| preserve_resource_error(super::metadata::map_identity_error(error)))?;
    let report = inspection.report();
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocations(report.retained_bytes())?;
    budget.charge_allocations(report.scratch_bytes())?;
    budget.charge_allocation_events(report.allocations())?;
    let identity_matches = identity.matches_expected(context, ids);
    if inspection.last_object_identifier() != ids.last_identifier || !identity_matches {
        return Err(SlideAudioCreationError::Verification);
    }

    let media_options = media_options(metadata_payload.len(), context, budget)?;
    let data_identifiers = media.data_references();
    let data_count = 1 + usize::from(media.poster_data_identifier().is_some());
    let mut media = MediaWitness::new(context, ids.drawable, &data_identifiers[..data_count])?;
    let report =
        media_codec::visit_package_metadata_media(metadata_payload, media_options, &mut media)
            .map_err(|error| {
                preserve_resource_error(super::metadata::map_media_error(
                    error,
                    SlideAudioCreationLimitKind::WireWork,
                ))
            })?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    let media_matches = media.matches_expected(context, identity.slide_identifier);
    if !media_matches {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(())
}

fn preserve_resource_error(error: SlideAudioCreationError) -> SlideAudioCreationError {
    match error {
        error @ (SlideAudioCreationError::LimitExceeded { .. }
        | SlideAudioCreationError::Allocation { .. }) => error,
        _ => SlideAudioCreationError::Verification,
    }
}

fn verify_node_cache(
    package: &Package,
    context: &CreationContext,
    expected_event_count: usize,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let (_, node_object) = package
        .object_with_component(context.slide_node_identifier)
        .ok_or(SlideAudioCreationError::Verification)?;
    let payload = unique_message_payload(node_object, SLIDE_NODE_MESSAGE_TYPE)?;
    let options = lifecycle_codec::DecodeOptions::new(
        residual(
            payload.len().min(context.wire_limits.max_input_bytes()),
            SlideAudioCreationLimitKind::InputBytes,
        )?,
        residual(
            context
                .wire_limits
                .max_output_bytes()
                .min(budget.remaining_output()),
            SlideAudioCreationLimitKind::OutputBytes,
        )?,
        residual(
            context
                .wire_limits
                .max_fields()
                .min(budget.remaining_wire_fields()),
            SlideAudioCreationLimitKind::WireFields,
        )?,
        residual(
            context
                .wire_limits
                .max_rewrite_work()
                .min(budget.remaining_work()),
            SlideAudioCreationLimitKind::WireWork,
        )?,
        1,
        u32::try_from(context.wire_limits.max_nesting()).unwrap_or(u32::MAX),
    );
    let (snapshot, report) =
        lifecycle_codec::node_cache::decode_slide_node_build_cache_with_report(payload, options)
            .map_err(|_| SlideAudioCreationError::Verification)?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocations(report.retained_bytes())?;
    budget.charge_allocations(report.scratch_bytes())?;
    budget.charge_allocation_events(report.allocations())?;
    let event_count =
        u32::try_from(expected_event_count).map_err(|_| SlideAudioCreationError::Verification)?;
    if snapshot.build_event_count() != (event_count != 0).then_some(event_count)
        || snapshot.build_event_count_cache_version()
            != Some(if event_count == 0 { u32::MAX } else { 2 })
        || snapshot.has_explicit_builds() != Some(event_count != 0)
        || snapshot.has_explicit_builds_cache_version() != Some(2)
    {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(())
}

#[derive(Debug)]
struct IdentityWitness<'a> {
    slide_identifier: u64,
    slide_locator: &'a str,
    stylesheet_identifier: u64,
    stylesheet_locator: &'a str,
    style_identifier: u64,
    object_ids: [u64; 4],
    uuids: [identity_codec::UuidBits; 4],
    uuid_matches: [usize; 4],
    external_matches: usize,
    external_target_identifier: Option<u64>,
    invalid: bool,
}

impl<'a> IdentityWitness<'a> {
    fn new(context: &'a CreationContext, ids: &CreationIds) -> Self {
        Self {
            slide_identifier: 0,
            slide_locator: component_locator(context.slide_component.as_ref()),
            stylesheet_identifier: 0,
            stylesheet_locator: component_locator(context.stylesheet_component.as_ref()),
            style_identifier: context.style_identifier,
            object_ids: ids.metadata_identifiers(),
            uuids: ids.metadata_uuids(),
            uuid_matches: [0; 4],
            external_matches: 0,
            external_target_identifier: None,
            invalid: false,
        }
    }

    fn matches_expected(&self, context: &CreationContext, ids: &CreationIds) -> bool {
        self.slide_identifier != 0
            && self.stylesheet_identifier != 0
            && self.slide_locator == component_locator(context.slide_component.as_ref())
            && self.stylesheet_locator == component_locator(context.stylesheet_component.as_ref())
            && self.uuid_matches == [1; 4]
            && self.external_matches == 1
            && self.external_target_identifier == Some(self.stylesheet_identifier)
            && !self.invalid
            && self.object_ids == ids.metadata_identifiers()
    }
}

impl identity_codec::PackageMetadataVisitor for IdentityWitness<'_> {
    fn visit_unknown_field(&mut self) -> Result<(), identity_codec::RewriteError> {
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: identity_codec::ComponentDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        if component.is_current() && component.effective_locator() == self.slide_locator {
            if self.slide_identifier != 0 {
                self.invalid = true;
            }
            self.slide_identifier = component.identifier();
        }
        if component.is_current() && component.effective_locator() == self.stylesheet_locator {
            if self.stylesheet_identifier != 0 {
                self.invalid = true;
            }
            self.stylesheet_identifier = component.identifier();
        }
        Ok(())
    }

    fn visit_object_uuid(
        &mut self,
        binding: identity_codec::ObjectUuidDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        let component = binding.component();
        if !component.is_current() || component.effective_locator() != self.slide_locator {
            return Ok(());
        }
        for (index, identifier) in self.object_ids.iter().copied().enumerate() {
            if binding.object_identifier() == identifier {
                if binding.uuid() != self.uuids[index] {
                    self.invalid = true;
                }
                self.uuid_matches[index] = self.uuid_matches[index].saturating_add(1);
            }
        }
        Ok(())
    }

    fn visit_external_reference(
        &mut self,
        reference: identity_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        let source = reference.source();
        if source.is_current()
            && source.effective_locator() == self.slide_locator
            && reference.object_identifier() == Some(self.style_identifier)
            && !reference.is_versioned()
            && reference.is_weak() != Some(true)
        {
            self.external_matches = self.external_matches.saturating_add(1);
            if self
                .external_target_identifier
                .replace(reference.target_component_identifier())
                .is_some_and(|identifier| identifier != reference.target_component_identifier())
            {
                self.invalid = true;
            }
        }
        Ok(())
    }
}

#[derive(Debug)]
struct MediaWitness<'a> {
    slide_identifier: u64,
    slide_locator: &'a str,
    drawable_identifier: u64,
    data_identifiers: [u64; 2],
    data_count: usize,
    data_matches: [usize; 2],
    component_matches: usize,
    data_reference_matches: [usize; 2],
    owner_matches: [usize; 2],
    owner_count: [Option<usize>; 2],
    owner_identifier: [Option<u64>; 2],
    owner_value: [Option<u32>; 2],
}

impl<'a> MediaWitness<'a> {
    fn new(
        context: &'a CreationContext,
        drawable_identifier: u64,
        data_identifiers: &[u64],
    ) -> Result<Self, SlideAudioCreationError> {
        if data_identifiers.is_empty()
            || data_identifiers.len() > 2
            || data_identifiers
                .iter()
                .enumerate()
                .any(|(index, identifier)| {
                    *identifier == 0 || data_identifiers[..index].contains(identifier)
                })
        {
            return Err(SlideAudioCreationError::Verification);
        }
        let mut selected = [0; 2];
        selected[..data_identifiers.len()].copy_from_slice(data_identifiers);
        Ok(Self {
            slide_identifier: 0,
            slide_locator: component_locator(context.slide_component.as_ref()),
            drawable_identifier,
            data_identifiers: selected,
            data_count: data_identifiers.len(),
            data_matches: [0; 2],
            component_matches: 0,
            data_reference_matches: [0; 2],
            owner_matches: [0; 2],
            owner_count: [None; 2],
            owner_identifier: [None; 2],
            owner_value: [None; 2],
        })
    }

    fn matches_expected(&self, context: &CreationContext, component_identifier: u64) -> bool {
        self.slide_identifier == component_identifier
            && self.slide_locator == component_locator(context.slide_component.as_ref())
            && self.component_matches == 1
            && (0..self.data_count).all(|index| {
                self.data_matches[index] == 1
                    && self.data_reference_matches[index] == 1
                    && self.owner_matches[index] == 1
                    // Existing data records may already have owners from
                    // earlier drawables.  Preserve the audio witness's
                    // invariant: the selected parent declares at least one
                    // owner, and this transaction contributes exactly one
                    // matching owner registration below.
                    && self.owner_count[index].is_some_and(|count| count >= 1)
                    && self.owner_identifier[index] == Some(self.drawable_identifier)
                    && self.owner_value[index] == Some(1)
            })
    }

    fn data_index(&self, identifier: u64) -> Option<usize> {
        self.data_identifiers[..self.data_count]
            .iter()
            .position(|candidate| *candidate == identifier)
    }
}

impl media_codec::PackageMetadataMediaVisitor for MediaWitness<'_> {
    fn visit_data_info(
        &mut self,
        data_info: media_codec::DataInfoSnapshot<'_>,
    ) -> Result<(), media_codec::DecodeError> {
        if let Some(index) = self.data_index(data_info.identifier()) {
            self.data_matches[index] = self.data_matches[index].saturating_add(1);
        }
        Ok(())
    }

    fn visit_component(
        &mut self,
        component: media_codec::ComponentSnapshot<'_>,
    ) -> Result<(), media_codec::DecodeError> {
        if !component.is_versioned() && component.effective_locator() == self.slide_locator {
            self.slide_identifier = component.identifier();
            self.component_matches = self.component_matches.saturating_add(1);
        }
        Ok(())
    }

    fn visit_data_reference(
        &mut self,
        component: media_codec::ComponentSnapshot<'_>,
        data_reference: media_codec::ComponentDataReferenceSnapshot<'_>,
    ) -> Result<(), media_codec::DecodeError> {
        if !component.is_versioned()
            && component.effective_locator() == self.slide_locator
            && self.data_index(data_reference.data_identifier()).is_some()
        {
            let index = self
                .data_index(data_reference.data_identifier())
                .ok_or_else(media_codec::DecodeError::invalid_for_adapter)?;
            self.data_reference_matches[index] =
                self.data_reference_matches[index].saturating_add(1);
            self.owner_count[index] = Some(data_reference.owner_count());
        }
        Ok(())
    }

    fn visit_owner(
        &mut self,
        component: media_codec::ComponentSnapshot<'_>,
        data_reference: media_codec::ComponentDataReferenceSnapshot<'_>,
        owner: media_codec::OwnerSnapshot<'_>,
    ) -> Result<(), media_codec::DecodeError> {
        if !component.is_versioned()
            && component.effective_locator() == self.slide_locator
            && owner.object_identifier() == self.drawable_identifier
        {
            let Some(index) = self.data_index(data_reference.data_identifier()) else {
                return Ok(());
            };
            self.owner_matches[index] = self.owner_matches[index].saturating_add(1);
            self.owner_identifier[index] = Some(owner.object_identifier());
            self.owner_value[index] = Some(owner.count());
        }
        Ok(())
    }
}

fn verify_object<'a>(
    package: &'a Package,
    component: &str,
    identifier: u64,
    message_type: u32,
) -> Result<&'a ArchiveObject, SlideAudioCreationError> {
    let (actual_component, object) = package
        .object_with_component(identifier)
        .ok_or(SlideAudioCreationError::Verification)?;
    if actual_component != component
        || object.archive_info.identifier != Some(identifier)
        || object.messages.len() != 1
        || object.archive_info.message_infos.len() != 1
        || object.messages[0].type_ != message_type
        || object.archive_info.message_infos[0].type_ != message_type
    {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(object)
}

fn verify_message_header(
    object: &ArchiveObject,
    versions: &[u32],
    object_references: &[u64],
    data_references: &[u64],
) -> Result<(), SlideAudioCreationError> {
    let info = object
        .archive_info
        .message_infos
        .first()
        .ok_or(SlideAudioCreationError::Verification)?;
    if info.versions != versions
        || info.object_references != object_references
        || info.data_references != data_references
    {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(())
}

fn unique_message_payload(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<&[u8], SlideAudioCreationError> {
    let mut payload = None;
    for message in object
        .messages
        .iter()
        .filter(|message| message.type_ == message_type)
    {
        if payload.replace(message.data.as_slice()).is_some() {
            return Err(SlideAudioCreationError::Verification);
        }
    }
    payload.ok_or(SlideAudioCreationError::Verification)
}

fn unique_archive_message_payload(
    archive: &Archive,
    message_type: u32,
) -> Result<&[u8], SlideAudioCreationError> {
    let mut payload = None;
    for object in &archive.objects {
        for message in object
            .messages
            .iter()
            .filter(|message| message.type_ == message_type)
        {
            if payload.replace(message.data.as_slice()).is_some() {
                return Err(SlideAudioCreationError::Verification);
            }
        }
    }
    payload.ok_or(SlideAudioCreationError::Verification)
}

fn message_payload_with_header(
    object: &ArchiveObject,
    message_type: u32,
) -> Result<&[u8], SlideAudioCreationError> {
    let mut selected = None;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != message_type {
            continue;
        }
        if selected.replace(index).is_some()
            || object
                .archive_info
                .message_infos
                .get(index)
                .map(|info| info.type_)
                != Some(message_type)
        {
            return Err(SlideAudioCreationError::Verification);
        }
    }
    selected
        .and_then(|index| object.messages.get(index))
        .map(|message| message.data.as_slice())
        .ok_or(SlideAudioCreationError::Verification)
}

fn lifecycle_options(
    source_length: usize,
    limits: WireLimits,
    budget: &CreationBudget,
) -> Result<lifecycle_codec::DecodeOptions, SlideAudioCreationError> {
    let output = source_length
        .checked_mul(2)
        .and_then(|value| value.checked_add(4096))
        .ok_or(SlideAudioCreationError::Verification)?;
    let work = source_length
        .checked_mul(256)
        .and_then(|value| value.checked_add(4096))
        .ok_or(SlideAudioCreationError::Verification)?;
    Ok(lifecycle_codec::DecodeOptions::new(
        residual(
            source_length.min(limits.max_input_bytes()),
            SlideAudioCreationLimitKind::InputBytes,
        )?,
        residual(
            output
                .min(limits.max_output_bytes())
                .min(budget.remaining_output()),
            SlideAudioCreationLimitKind::OutputBytes,
        )?,
        residual(
            limits.max_fields().min(budget.remaining_wire_fields()),
            SlideAudioCreationLimitKind::WireFields,
        )?,
        residual(
            work.min(limits.max_rewrite_work())
                .min(budget.remaining_work()),
            SlideAudioCreationLimitKind::WireWork,
        )?,
        residual(
            limits.max_fields().min(budget.remaining_references()),
            SlideAudioCreationLimitKind::References,
        )?,
        u32::try_from(limits.max_nesting()).map_err(|_| SlideAudioCreationError::Verification)?,
    ))
}

fn residual(
    value: usize,
    kind: SlideAudioCreationLimitKind,
) -> Result<usize, SlideAudioCreationError> {
    super::nonzero_residual(value, kind)
}

fn movie_caption_options(
    source_length: usize,
    limits: WireLimits,
    budget: &CreationBudget,
) -> Result<keynote_movie_caption_codec::DecodeOptions, SlideAudioCreationError> {
    Ok(keynote_movie_caption_codec::DecodeOptions::new(
        residual(
            source_length.min(limits.max_input_bytes()),
            SlideAudioCreationLimitKind::InputBytes,
        )?,
        residual(
            limits.max_fields().min(budget.remaining_wire_fields()),
            SlideAudioCreationLimitKind::WireFields,
        )?,
        residual(
            limits.max_rewrite_work().min(budget.remaining_work()),
            SlideAudioCreationLimitKind::WireWork,
        )?,
        u32::try_from(limits.max_nesting()).map_err(|_| SlideAudioCreationError::Verification)?,
    ))
}

fn identity_options(
    package: &Package,
    source_length: usize,
    context: &CreationContext,
    budget: &CreationBudget,
) -> Result<identity_codec::RewriteOptions, SlideAudioCreationError> {
    let limits = context.wire_limits;
    Ok(identity_codec::RewriteOptions::new(
        residual(
            source_length.min(limits.max_input_bytes()),
            SlideAudioCreationLimitKind::InputBytes,
        )?,
        residual(
            limits.max_output_bytes().min(budget.remaining_output()),
            SlideAudioCreationLimitKind::OutputBytes,
        )?,
        residual(
            limits.max_fields().min(budget.remaining_wire_fields()),
            SlideAudioCreationLimitKind::WireFields,
        )?,
        residual(
            limits.max_rewrite_work().min(budget.remaining_work()),
            SlideAudioCreationLimitKind::WireWork,
        )?,
        u32::try_from(limits.max_nesting()).map_err(|_| SlideAudioCreationError::Verification)?,
        package.semantic_limits().max_objects(),
        residual(
            package
                .semantic_limits()
                .max_references()
                .min(budget.remaining_references()),
            SlideAudioCreationLimitKind::References,
        )?,
        package.semantic_limits().max_objects(),
    ))
}

fn media_options(
    source_length: usize,
    context: &CreationContext,
    budget: &CreationBudget,
) -> Result<media_codec::DecodeOptions, SlideAudioCreationError> {
    let limits = context.wire_limits;
    Ok(media_codec::DecodeOptions::new(
        residual(
            source_length.min(limits.max_input_bytes()),
            SlideAudioCreationLimitKind::InputBytes,
        )?,
        residual(
            limits.max_fields().min(budget.remaining_wire_fields()),
            SlideAudioCreationLimitKind::WireFields,
        )?,
        residual(
            limits.max_rewrite_work().min(budget.remaining_work()),
            SlideAudioCreationLimitKind::WireWork,
        )?,
        package_topology_limit(budget)?,
        package_topology_limit(budget)?,
        package_topology_limit(budget)?,
        20,
        MAX_FILENAME_BYTES,
        u32::try_from(limits.max_nesting()).map_err(|_| SlideAudioCreationError::Verification)?,
    ))
}

fn package_topology_limit(budget: &CreationBudget) -> Result<usize, SlideAudioCreationError> {
    residual(
        budget.remaining_references(),
        SlideAudioCreationLimitKind::References,
    )
}

fn charge_payload(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    let _ = wire_view(payload, limits, budget)?;
    Ok(())
}

fn wire_view<'a>(
    payload: &'a [u8],
    limits: WireLimits,
    budget: &mut CreationBudget,
) -> Result<WireView<'a>, SlideAudioCreationError> {
    let view = WireView::parse_with_limits(payload, limits)
        .map_err(|_| SlideAudioCreationError::Verification)?;
    budget.charge_wire_fields(view.len())?;
    budget.charge_work(payload.len().max(1))?;
    for field in view.fields() {
        field
            .validate_canonical_framing()
            .map_err(|_| SlideAudioCreationError::Verification)?;
    }
    Ok(view)
}

fn references_in_field(
    payload: &[u8],
    number: u32,
    limits: WireLimits,
    budget: &mut CreationBudget,
) -> Result<Vec<u64>, SlideAudioCreationError> {
    let view = wire_view(payload, limits, budget)?;
    let count = view
        .fields()
        .filter(|field| field.number() == number)
        .count();
    let bytes = count
        .checked_mul(size_of::<u64>())
        .ok_or(SlideAudioCreationError::Verification)?;
    budget.charge_allocations(bytes)?;
    let mut identifiers = Vec::new();
    identifiers
        .try_reserve_exact(count)
        .map_err(|_| SlideAudioCreationError::Allocation { amount: bytes })?;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 {
            return Err(SlideAudioCreationError::Verification);
        }
        identifiers.push(reference_identifier(field.payload(), limits, budget)?);
    }
    Ok(identifiers)
}

fn required_nested_field<'a>(
    payload: &'a [u8],
    number: u32,
    limits: WireLimits,
    budget: &mut CreationBudget,
) -> Result<&'a [u8], SlideAudioCreationError> {
    let view = wire_view(payload, limits, budget)?;
    let mut result = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 || result.replace(field.payload()).is_some() {
            return Err(SlideAudioCreationError::Verification);
        }
    }
    result.ok_or(SlideAudioCreationError::Verification)
}

fn required_bytes<'a>(
    view: &WireView<'a>,
    number: u32,
) -> Result<&'a [u8], SlideAudioCreationError> {
    let mut result = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if field.wire_type() != 2 || result.replace(field.payload()).is_some() {
            return Err(SlideAudioCreationError::Verification);
        }
    }
    result.ok_or(SlideAudioCreationError::Verification)
}

fn required_text<'a>(
    view: &WireView<'a>,
    number: u32,
) -> Result<&'a [u8], SlideAudioCreationError> {
    let field = unique_wire_field(view, number)?;
    if field.wire_type() != 2 {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(field.payload())
}

fn unique_wire_field<'a>(
    view: &WireView<'a>,
    number: u32,
) -> Result<WireFieldView<'a>, SlideAudioCreationError> {
    let mut result = None;
    for field in view.fields().filter(|field| field.number() == number) {
        if result.replace(field).is_some() {
            return Err(SlideAudioCreationError::Verification);
        }
    }
    result.ok_or(SlideAudioCreationError::Verification)
}

fn required_varint(view: &WireView<'_>, number: u32) -> Result<u64, SlideAudioCreationError> {
    let field = unique_wire_field(view, number)?;
    if field.wire_type() != 0 {
        return Err(SlideAudioCreationError::Verification);
    }
    let (value, width) = litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
        .map_err(|_| SlideAudioCreationError::Verification)?;
    if width != field.payload().len() || litchi_iwa_common::varint::encoded_len(value) != width {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(value)
}

fn required_fixed32(view: &WireView<'_>, number: u32) -> Result<u32, SlideAudioCreationError> {
    let field = unique_wire_field(view, number)?;
    if field.wire_type() != 5 || field.payload().len() != 4 {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(u32::from_le_bytes(
        field
            .payload()
            .try_into()
            .map_err(|_| SlideAudioCreationError::Verification)?,
    ))
}

fn reference_identifier(
    payload: &[u8],
    limits: WireLimits,
    budget: &mut CreationBudget,
) -> Result<u64, SlideAudioCreationError> {
    let view = wire_view(payload, limits, budget)?;
    let mut result = None;
    for field in view.fields().filter(|field| field.number() == 1) {
        if field.wire_type() != 0 || result.is_some() {
            return Err(SlideAudioCreationError::Verification);
        }
        let (identifier, width) =
            litchi_iwa_common::varint::decode_varint_from_bytes(field.payload())
                .map_err(|_| SlideAudioCreationError::Verification)?;
        if width != field.payload().len()
            || litchi_iwa_common::varint::encoded_len(identifier) != width
            || identifier == 0
        {
            return Err(SlideAudioCreationError::Verification);
        }
        result = Some(identifier);
    }
    result.ok_or(SlideAudioCreationError::Verification)
}

fn collect_reference_ids<'a>(
    references: impl ExactSizeIterator<Item = lifecycle_codec::Reference<'a>>,
    budget: &mut CreationBudget,
) -> Result<Vec<u64>, SlideAudioCreationError> {
    let count = references.len();
    let bytes = count
        .checked_mul(size_of::<u64>())
        .ok_or(SlideAudioCreationError::Verification)?;
    let mut result = Vec::new();
    if count != 0 {
        budget.charge_allocations(bytes)?;
        result
            .try_reserve_exact(count)
            .map_err(|_| SlideAudioCreationError::Allocation { amount: bytes })?;
    }
    for reference in references {
        if result.len() == count {
            return Err(SlideAudioCreationError::Verification);
        }
        result.push(reference.identifier());
    }
    if result.len() != count {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(result)
}

fn contains_once(values: &[u64], identifier: u64) -> bool {
    values.iter().filter(|value| **value == identifier).count() == 1
}

fn charge_lifecycle_report(
    report: lifecycle_codec::DecodeReport,
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocations(report.retained_bytes())?;
    budget.charge_allocations(report.scratch_bytes())?;
    budget.charge_allocation_events(report.allocations())
}

fn component_locator(name: &str) -> &str {
    let without_prefix = name.strip_prefix("Index/").unwrap_or(name);
    without_prefix
        .strip_suffix(".iwa")
        .unwrap_or(without_prefix)
}
