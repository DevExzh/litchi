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

use litchi_iwa_common::{WireLimits, wire::WireView};
use litchi_iwa_core::{Archive, ArchiveObject};
use litchi_iwa_protos::{
    keynote_build_creation_codec as build_creation_codec, keynote_media_codec,
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
    context.validate()?;
    ids.validate()?;
    if data_identifier == 0 {
        return Err(SlideAudioCreationError::Verification);
    }

    verify_selected_style(package, context, budget)?;
    verify_slide_lifecycle(package, context, ids, expected_event_count, budget)?;
    verify_created_objects(package, context, ids, data_identifier, budget)?;
    verify_metadata(package, context, ids, data_identifier, budget)?;
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
    data_identifier: u64,
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

    verify_message_header(
        movie,
        &STANDARD_MESSAGE_VERSION,
        &[ids.caption, ids.title, context.style_identifier],
        &[data_identifier],
    )?;
    verify_message_header(title, &STANDIN_CAPTION_MESSAGE_VERSION, &[], &[])?;
    verify_message_header(caption, &STANDIN_CAPTION_MESSAGE_VERSION, &[], &[])?;
    verify_message_header(build, &STANDARD_MESSAGE_VERSION, &[ids.drawable], &[])?;
    verify_message_header(chunk, &STANDARD_MESSAGE_VERSION, &[ids.build], &[])?;

    let movie_payload = unique_message_payload(movie, MOVIE_MESSAGE_TYPE)?;
    verify_movie_payload(movie_payload, context, ids, data_identifier, budget)?;
    if !unique_message_payload(title, STANDIN_CAPTION_MESSAGE_TYPE)?.is_empty()
        || !unique_message_payload(caption, STANDIN_CAPTION_MESSAGE_TYPE)?.is_empty()
    {
        return Err(SlideAudioCreationError::Verification);
    }
    verify_build_payload(build, ids, context, budget)?;
    verify_chunk_payload(chunk, ids, context, budget)?;
    Ok(())
}

fn verify_movie_payload(
    payload: &[u8],
    context: &CreationContext,
    ids: &CreationIds,
    data_identifier: u64,
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
    if data_snapshot.identifier() != data_identifier {
        return Err(SlideAudioCreationError::Verification);
    }
    let style_payload = required_bytes(&root, 19)?;
    if reference_identifier(style_payload, context.wire_limits, budget)? != context.style_identifier
    {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(())
}

fn verify_build_payload(
    object: &ArchiveObject,
    ids: &CreationIds,
    context: &CreationContext,
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

    let write = build_creation_codec::StartAudioBuildWrite::new(
        ids.drawable,
        ids.build,
        ids.chunk,
        ids.build_uuid.lower(),
        ids.build_uuid.upper(),
        ids.random_number_seed,
    );
    let options = build_creation_codec::EncodeOptions::for_write(&write)
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
        )?);
    let expected = build_creation_codec::encode_start_audio_build(&write, options)
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

    let write = build_creation_codec::StartAudioBuildWrite::new(
        ids.drawable,
        ids.build,
        ids.chunk,
        ids.build_uuid.lower(),
        ids.build_uuid.upper(),
        ids.random_number_seed,
    );
    let options = build_creation_codec::EncodeOptions::for_write(&write)
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
        )?);
    let expected = build_creation_codec::encode_start_audio_chunk(&write, options)
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

fn verify_metadata(
    package: &Package,
    context: &CreationContext,
    ids: &CreationIds,
    data_identifier: u64,
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
    .map_err(|_| SlideAudioCreationError::Verification)?;
    let report = inspection.report();
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    budget.charge_allocations(report.retained_bytes())?;
    budget.charge_allocations(report.scratch_bytes())?;
    budget.charge_allocation_events(report.allocations())?;
    if inspection.last_object_identifier() != ids.last_identifier
        || !identity.matches_expected(context, ids)
    {
        return Err(SlideAudioCreationError::Verification);
    }

    let media_options = media_options(metadata_payload.len(), context, budget)?;
    let mut media = MediaWitness::new(context, ids.drawable, data_identifier);
    let report =
        media_codec::visit_package_metadata_media(metadata_payload, media_options, &mut media)
            .map_err(|_| SlideAudioCreationError::Verification)?;
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_nesting(report.max_depth() as usize)?;
    if !media.matches_expected(context, identity.slide_identifier) {
        return Err(SlideAudioCreationError::Verification);
    }
    Ok(())
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
    data_identifier: u64,
    data_matches: usize,
    component_matches: usize,
    data_reference_matches: usize,
    owner_matches: usize,
    owner_count: Option<usize>,
    owner_identifier: Option<u64>,
    owner_value: Option<u32>,
}

impl<'a> MediaWitness<'a> {
    fn new(context: &'a CreationContext, drawable_identifier: u64, data_identifier: u64) -> Self {
        Self {
            slide_identifier: 0,
            slide_locator: component_locator(context.slide_component.as_ref()),
            drawable_identifier,
            data_identifier,
            data_matches: 0,
            component_matches: 0,
            data_reference_matches: 0,
            owner_matches: 0,
            owner_count: None,
            owner_identifier: None,
            owner_value: None,
        }
    }

    fn matches_expected(&self, context: &CreationContext, component_identifier: u64) -> bool {
        self.slide_identifier == component_identifier
            && self.slide_locator == component_locator(context.slide_component.as_ref())
            && self.data_matches == 1
            && self.component_matches == 1
            && self.data_reference_matches == 1
            && self.owner_matches == 1
            && self.owner_count.is_some_and(|count| count >= 1)
            && self.owner_identifier == Some(self.drawable_identifier)
            && self.owner_value == Some(1)
    }
}

impl media_codec::PackageMetadataMediaVisitor for MediaWitness<'_> {
    fn visit_data_info(
        &mut self,
        data_info: media_codec::DataInfoSnapshot<'_>,
    ) -> Result<(), media_codec::DecodeError> {
        if data_info.identifier() == self.data_identifier {
            self.data_matches = self.data_matches.saturating_add(1);
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
            && data_reference.data_identifier() == self.data_identifier
        {
            self.data_reference_matches = self.data_reference_matches.saturating_add(1);
            self.owner_count = Some(data_reference.owner_count());
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
            && data_reference.data_identifier() == self.data_identifier
            && owner.object_identifier() == self.drawable_identifier
        {
            self.owner_matches = self.owner_matches.saturating_add(1);
            self.owner_identifier = Some(owner.object_identifier());
            self.owner_value = Some(owner.count());
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
    references: impl Iterator<Item = lifecycle_codec::Reference<'a>>,
    budget: &mut CreationBudget,
) -> Result<Vec<u64>, SlideAudioCreationError> {
    let mut result = Vec::new();
    for reference in references {
        budget.charge_allocations(size_of::<u64>())?;
        result
            .try_reserve(1)
            .map_err(|_| SlideAudioCreationError::Allocation {
                amount: size_of::<u64>(),
            })?;
        result.push(reference.identifier());
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
