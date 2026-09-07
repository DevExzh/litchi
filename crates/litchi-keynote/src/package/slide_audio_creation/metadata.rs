//! Bounded PackageMetadata planning for a fresh slide-owned audio object.
//!
//! This adapter deliberately stops at the metadata member.  Native drawable,
//! build, and ZIP edits are staged by the package owner after this function
//! returns.  All observations are borrowed from the immutable source and both
//! metadata codecs publish candidates only after their own strict verification.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The planner keeps its source admission, identity, and media phases adjacent."
)]

use std::collections::HashSet;
use std::fmt::Write as _;
use std::mem::size_of;
use std::path::{Component, Path};

use litchi_iwa_archive::SourceCatalog;
use litchi_iwa_common::media::Type as MediaType;
use litchi_iwa_core::{Archive, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    package_metadata_codec as identity_codec, package_metadata_media_codec as media_codec,
};
use sha1::{Digest, Sha1};

use crate::slide::audio::creation::{SlideAudioCreationError, SlideAudioCreationLimitKind};
use crate::soundtrack::items::MAX_FILENAME_BYTES;

use super::super::{Package, PhysicalSource};
use super::{CreationBudget, CreationContext, CreationIds};

const METADATA_COMPONENT: &str = "Index/Metadata.iwa";
const PACKAGE_METADATA_MESSAGE_TYPE: u32 = 11_006;
const SHA1_BYTES: usize = 20;
const MAX_NAME_ATTEMPTS: u64 = 1_024;

/// Metadata and data-member work staged for the package transaction.
///
/// The audio bytes are borrowed from the caller.  A new data member can thus
/// be inserted without copying the payload; a reused digest has no data
/// insertion at all.
#[derive(Debug)]
pub(super) struct MetadataPlan {
    pub(super) data_identifier: u64,
    pub(super) digest: [u8; SHA1_BYTES],
    pub(super) compressed: Vec<u8>,
    pub(super) data_entry_name: Option<Box<str>>,
    pub(super) created_data: usize,
}

#[derive(Debug, Clone, Copy)]
struct DataRecord {
    identifier: u64,
    digest: [u8; SHA1_BYTES],
    materialized_length: Option<u64>,
}

#[derive(Debug, Default)]
struct MediaFacts {
    data: Vec<DataRecord>,
    matching_name: Option<Box<str>>,
    matching_identifier: Option<u64>,
    duplicate_match: bool,
    slide_component: ComponentMatch,
    node_component: ComponentMatch,
    stylesheet_component: ComponentMatch,
}

#[derive(Debug, Default, Clone, Copy)]
struct ComponentMatch {
    identifier: Option<u64>,
    matches: usize,
    unknown_fields: bool,
}

#[derive(Debug, Default)]
struct ExistingExternalReference {
    source_identifier: u64,
    target_component_identifier: u64,
    object_identifier: u64,
    matches: usize,
}

impl identity_codec::PackageMetadataVisitor for ExistingExternalReference {
    fn visit_external_reference(
        &mut self,
        reference: identity_codec::ExternalReferenceDescriptor<'_>,
    ) -> Result<(), identity_codec::RewriteError> {
        let source = reference.source();
        if reference.is_versioned()
            || !source.is_current()
            || source.identifier() != self.source_identifier
            || reference.target_component_identifier() != self.target_component_identifier
            || reference.object_identifier() != Some(self.object_identifier)
            || reference.is_weak() == Some(true)
        {
            return Ok(());
        }
        self.matches = self.matches.saturating_add(1);
        Ok(())
    }
}

struct MediaVisitor<'expected, 'budget> {
    target_digest: [u8; SHA1_BYTES],
    slide_locator: &'expected str,
    node_locator: &'expected str,
    stylesheet_locator: &'expected str,
    budget: &'budget mut CreationBudget,
    allocation_failure: Option<SlideAudioCreationError>,
    facts: MediaFacts,
}

impl<'expected, 'budget> MediaVisitor<'expected, 'budget> {
    fn reject(&mut self, error: SlideAudioCreationError) -> media_codec::DecodeError {
        if self.allocation_failure.is_none() {
            self.allocation_failure = Some(error);
        }
        media_codec::DecodeError::invalid_for_adapter()
    }
}

impl media_codec::PackageMetadataMediaVisitor for MediaVisitor<'_, '_> {
    fn visit_component(
        &mut self,
        component: media_codec::ComponentSnapshot<'_>,
    ) -> Result<(), media_codec::DecodeError> {
        if component.is_versioned() {
            return Ok(());
        }
        if component.effective_locator() == self.slide_locator {
            record_component(&mut self.facts.slide_component, component)?;
        }
        if component.effective_locator() == self.node_locator {
            record_component(&mut self.facts.node_component, component)?;
        }
        if component.effective_locator() == self.stylesheet_locator {
            record_component(&mut self.facts.stylesheet_component, component)?;
        }
        Ok(())
    }

    fn visit_data_info(
        &mut self,
        data_info: media_codec::DataInfoSnapshot<'_>,
    ) -> Result<(), media_codec::DecodeError> {
        let digest = <[u8; SHA1_BYTES]>::try_from(data_info.digest())
            .map_err(|_| media_codec::DecodeError::invalid_for_adapter())?;
        let current_name = data_info
            .file_name()
            .filter(|name| !name.is_empty())
            .unwrap_or_else(|| data_info.preferred_file_name());
        let record_bytes = size_of::<DataRecord>();
        if let Err(error) = self.budget.charge_allocations(record_bytes) {
            return Err(self.reject(error));
        }
        if self.facts.data.try_reserve_exact(1).is_err() {
            return Err(self.reject(SlideAudioCreationError::Allocation {
                amount: record_bytes,
            }));
        }
        self.facts.data.push(DataRecord {
            identifier: data_info.identifier(),
            digest,
            materialized_length: data_info.materialized_length(),
        });
        if digest == self.target_digest {
            if self.facts.matching_identifier.is_some() {
                self.facts.duplicate_match = true;
            } else {
                self.facts.matching_identifier = Some(data_info.identifier());
                let name_bytes = current_name.len();
                if let Err(error) = self.budget.charge_allocations(name_bytes) {
                    return Err(self.reject(error));
                }
                let mut name = String::new();
                if name.try_reserve_exact(name_bytes).is_err() {
                    return Err(
                        self.reject(SlideAudioCreationError::Allocation { amount: name_bytes })
                    );
                }
                name.push_str(current_name);
                self.facts.matching_name = Some(name.into_boxed_str());
            }
        }
        Ok(())
    }
}

fn record_component(
    selected: &mut ComponentMatch,
    component: media_codec::ComponentSnapshot<'_>,
) -> Result<(), media_codec::DecodeError> {
    selected.matches = selected
        .matches
        .checked_add(1)
        .ok_or_else(media_codec::DecodeError::invalid_for_adapter)?;
    if selected.matches == 1 {
        selected.identifier = Some(component.identifier());
    }
    selected.unknown_fields |= component.has_unknown_fields();
    Ok(())
}

/// Plan one fresh slide-owned audio metadata transaction.
///
/// The returned candidate is detached from the source package and can be
/// combined with native archive edits by the caller.  No source bytes or
/// package members are changed if either metadata codec refuses the staged
/// identity/media closure.
pub(super) fn plan_and_rewrite(
    source: &Package,
    ctx: &CreationContext,
    ids: &CreationIds,
    filename: &str,
    bytes: &[u8],
    budget: &mut CreationBudget,
) -> Result<MetadataPlan, SlideAudioCreationError> {
    ctx.validate()?;
    ids.validate()?;
    validate_audio_input(filename, bytes, budget)?;

    let catalog = match &source.state.source {
        PhysicalSource::Package(catalog) if catalog.source_is_exact() => catalog,
        PhysicalSource::Package(_) | PhysicalSource::Semantic(_) => {
            return Err(SlideAudioCreationError::UnsupportedSource);
        },
    };

    let (mut archive, object_index, message_index) =
        read_metadata_archive(source, catalog, ctx, budget)?;
    let payload = archive
        .objects
        .get(object_index)
        .and_then(|object| object.messages.get(message_index))
        .map(|message| message.data.as_slice())
        .ok_or(SlideAudioCreationError::InvalidSource)?;

    let slide_locator = component_locator(ctx.slide_component.as_ref())?;
    let stylesheet_locator = component_locator(ctx.stylesheet_component.as_ref())?;
    let (node_component_name, _) = source
        .object_with_component(ctx.slide_node_identifier)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let node_locator = component_locator(node_component_name)?;
    // SHA-1 is a full input traversal. Debit that work before touching the
    // caller's bytes so a depleted operation ledger rejects the transaction
    // without beginning the media scan.
    budget.charge_work(bytes.len())?;
    let source_media_options = media_options(payload, ctx, budget)?;
    let mut visitor = MediaVisitor {
        target_digest: Sha1::digest(bytes).into(),
        slide_locator,
        node_locator,
        stylesheet_locator,
        budget,
        allocation_failure: None,
        facts: MediaFacts::default(),
    };
    let media_result =
        media_codec::visit_package_metadata_media(payload, source_media_options, &mut visitor);
    let allocation_failure = visitor.allocation_failure.take();
    let digest = visitor.target_digest;
    let facts = std::mem::take(&mut visitor.facts);
    drop(visitor);
    if let Some(error) = allocation_failure {
        return Err(error);
    }
    let media_report = media_result
        .map_err(|error| map_media_error(error, SlideAudioCreationLimitKind::WireWork))?;
    budget.charge_wire_fields(media_report.fields())?;
    budget.charge_work(media_report.work_bytes())?;
    budget.charge_nesting(media_report.max_depth() as usize)?;

    let (data_identifier, data_entry_name, data_leaf_name, reused_existing, created_data) =
        resolve_data(catalog, &facts, &digest, filename, bytes, ids, budget)?;

    let slide_component = resolve_component(&facts.slide_component, slide_locator)?;
    let node_component = resolve_component(&facts.node_component, node_locator)?;
    let stylesheet_component = resolve_component(&facts.stylesheet_component, stylesheet_locator)?;

    let source_selector =
        identity_codec::ComponentSelector::new(slide_component.identifier, slide_component.locator);
    let stylesheet_selector = identity_codec::ComponentSelector::new(
        stylesheet_component.identifier,
        stylesheet_component.locator,
    );
    let metadata_identifiers = ids.metadata_identifiers();
    let metadata_uuids = ids.metadata_uuids();
    let uuid_additions = [
        identity_codec::ObjectUuidAddition::new(
            source_selector,
            metadata_identifiers[0],
            metadata_uuids[0],
        ),
        identity_codec::ObjectUuidAddition::new(
            source_selector,
            metadata_identifiers[1],
            metadata_uuids[1],
        ),
        identity_codec::ObjectUuidAddition::new(
            source_selector,
            metadata_identifiers[2],
            metadata_uuids[2],
        ),
        identity_codec::ObjectUuidAddition::new(
            source_selector,
            metadata_identifiers[3],
            metadata_uuids[3],
        ),
    ];
    let external_addition = identity_codec::ExternalReferenceAddition::new(
        source_selector,
        stylesheet_selector,
        ctx.style_identifier,
        None,
    );
    let node_selector =
        identity_codec::ComponentSelector::new(node_component.identifier, node_component.locator);
    let save_tokens = if source_selector == node_selector {
        [source_selector, source_selector]
    } else {
        [source_selector, node_selector]
    };
    let save_token_count = if source_selector == node_selector {
        1
    } else {
        2
    };
    let mut existing_external = ExistingExternalReference {
        source_identifier: slide_component.identifier,
        target_component_identifier: stylesheet_component.identifier,
        object_identifier: ctx.style_identifier,
        ..ExistingExternalReference::default()
    };
    let identity_inspection_options = identity_options(payload, ctx, budget)?;
    let identity_inspection = identity_codec::inspect_package_metadata_with_visitor(
        payload,
        identity_inspection_options,
        &mut existing_external,
    )
    .map_err(map_identity_error)?;
    charge_identity_inspection(budget, identity_inspection.report())?;
    if existing_external.matches > 1 {
        return Err(SlideAudioCreationError::InvalidSource);
    }
    let external_additions: &[identity_codec::ExternalReferenceAddition<'_>] =
        if existing_external.matches == 0 {
            std::slice::from_ref(&external_addition)
        } else {
            &[]
        };
    let identity_batch = identity_codec::Batch::new(
        ids.expected_last_identifier,
        ids.last_identifier,
        &uuid_additions,
        external_additions,
    );
    let identity_request = identity_codec::AdditionSaveTokenBatch::new(
        identity_batch,
        identity_codec::SaveTokenBatch::new(&save_tokens[..save_token_count]),
    );
    let identity_options = identity_options(payload, ctx, budget)?;
    let prepared_identity = identity_codec::prepare_package_metadata_additions_and_save_tokens(
        payload,
        identity_request,
        identity_options,
    )
    .map_err(map_identity_error)?;
    charge_identity_requirements(budget, &prepared_identity)?;
    let identity_limits = prepared_identity.execution_requirements().exact_limits();
    let identity_output = prepared_identity
        .execute(identity_limits)
        .map_err(map_identity_error)?;
    let identity_bytes = identity_output.into_bytes();

    let owner_addition = media_codec::DataReferenceOwnerAddition::new(
        media_codec::ComponentSelector::new(slide_component.identifier, slide_component.locator),
        data_identifier,
        ids.drawable,
        1,
    );
    let mut data_addition = None;
    if !reused_existing {
        if data_entry_name.is_none() {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        let leaf = data_leaf_name
            .as_deref()
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        data_addition = Some(
            media_codec::DataInfoAddition::new(data_identifier, &digest, filename)
                .with_file_name(leaf)
                .with_materialized_length(
                    u64::try_from(bytes.len())
                        .map_err(|_| SlideAudioCreationError::InvalidSource)?,
                ),
        );
    }
    let mut data_additions = Vec::new();
    if let Some(data_addition) = data_addition {
        budget.charge_allocations(size_of::<media_codec::DataInfoAddition<'_>>())?;
        data_additions
            .try_reserve_exact(1)
            .map_err(|_| SlideAudioCreationError::Allocation { amount: 1 })?;
        data_additions.push(data_addition);
    }
    let owner_additions = [owner_addition];
    let media_batch =
        media_codec::MediaRewriteBatch::new(&data_additions, &[], &owner_additions, &[]);
    let identity_media_options = media_options(&identity_bytes, ctx, budget)?;
    let prepared_media = media_codec::prepare_package_metadata_media_rewrite(
        &identity_bytes,
        media_batch,
        identity_media_options,
    )
    .map_err(map_media_rewrite_error)?;
    charge_media_requirements(budget, &prepared_media)?;
    let media_limits = prepared_media.execution_requirements().exact_limits();
    let media_output = prepared_media
        .execute(media_limits)
        .map_err(map_media_rewrite_error)?;
    let rewritten_payload = media_output.into_bytes();

    let object = archive
        .objects
        .get_mut(object_index)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    object
        .replace_message_preserving_header_with_limits(
            message_index,
            RawMessage {
                type_: PACKAGE_METADATA_MESSAGE_TYPE,
                data: rewritten_payload,
            },
            ctx.archive_limits,
        )
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let encoded_len = archive
        .encoded_len_with_limits(ctx.archive_limits)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_output(encoded_len)?;
    budget.charge_total(encoded_len)?;
    budget.charge_allocations(encoded_len)?;
    let serialized = archive
        .to_bytes_with_limits(ctx.archive_limits)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let compressed_max = SnappyStream::maximum_compressed_len(serialized.len())
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_output(compressed_max)?;
    budget.charge_total(compressed_max)?;
    budget.charge_allocations(compressed_max)?;
    let compressed =
        SnappyStream::compress(&serialized).map_err(|_| SlideAudioCreationError::InvalidSource)?;
    budget.charge_entry_bytes(compressed.len())?;

    // The data bytes remain owned by the caller.  The package owner copies
    // them into its staged ZIP insertion only when this plan created a new
    // DataInfo record.
    Ok(MetadataPlan {
        data_identifier,
        digest,
        compressed,
        data_entry_name,
        created_data,
    })
}

#[derive(Debug, Clone, Copy)]
struct ResolvedComponent<'source> {
    identifier: u64,
    locator: &'source str,
}

fn component_locator(name: &str) -> Result<&str, SlideAudioCreationError> {
    let without_prefix = name.strip_prefix("Index/").unwrap_or(name);
    let locator = without_prefix
        .strip_suffix(".iwa")
        .unwrap_or(without_prefix);
    if locator.is_empty() || locator.contains('/') || locator.contains('\\') {
        return Err(SlideAudioCreationError::InvalidSource);
    }
    Ok(locator)
}

fn resolve_component<'source>(
    selected: &ComponentMatch,
    locator: &'source str,
) -> Result<ResolvedComponent<'source>, SlideAudioCreationError> {
    if selected.matches != 1 || selected.unknown_fields {
        return Err(SlideAudioCreationError::InvalidSource);
    }
    Ok(ResolvedComponent {
        identifier: selected
            .identifier
            .ok_or(SlideAudioCreationError::InvalidSource)?,
        locator,
    })
}

fn validate_audio_input(
    filename: &str,
    bytes: &[u8],
    budget: &mut CreationBudget,
) -> Result<(), SlideAudioCreationError> {
    if filename.is_empty()
        || filename.len() > MAX_FILENAME_BYTES
        || filename
            .bytes()
            .any(|byte| byte == 0 || byte.is_ascii_control())
        || filename.contains(['/', '\\'])
        || Path::new(filename)
            .components()
            .any(|component| !matches!(component, Component::Normal(_)))
    {
        return Err(SlideAudioCreationError::InvalidFilename);
    }
    let extension = filename
        .rsplit_once('.')
        .map(|(_, extension)| extension)
        .filter(|extension| !extension.is_empty())
        .ok_or(SlideAudioCreationError::InvalidFilename)?;
    if MediaType::from_extension(extension) != MediaType::Audio {
        return Err(SlideAudioCreationError::InvalidFilename);
    }
    if bytes.is_empty() {
        return Err(SlideAudioCreationError::UnsupportedAudio);
    }
    if MediaType::from_bytes(bytes) != MediaType::Audio {
        return Err(SlideAudioCreationError::UnsupportedAudio);
    }
    budget.charge_input(bytes.len())?;
    budget.charge_entry_bytes(bytes.len())?;
    budget.charge_total(bytes.len())?;
    Ok(())
}

fn read_metadata_archive(
    package: &Package,
    catalog: &SourceCatalog,
    ctx: &CreationContext,
    budget: &mut CreationBudget,
) -> Result<(Archive, usize, usize), SlideAudioCreationError> {
    let entry = catalog
        .package()
        .iter()
        .find(|entry| entry.name() == METADATA_COMPONENT)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    if entry.is_opaque() {
        return Err(SlideAudioCreationError::InvalidSource);
    }
    budget.charge_entry_bytes(entry.data().len())?;
    // The Snappy decoder owns a decoded buffer, and Archive parsing then
    // retains message payloads in a second allocation. Charge the borrowed
    // compressed input before decoding and the exact decoded stream before
    // handing it to the archive parser.
    budget.charge_allocations(entry.data().len())?;
    let snappy_limits = package
        .limits()
        .snappy_limits()
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let stream = SnappyStream::decompress_with_limits(entry.data(), snappy_limits)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let decoded_len = stream.as_bytes().len();
    budget.charge_allocations(decoded_len)?;
    // Archive parsing retains each selected message payload independently of
    // the Snappy buffer. Reserve the same bounded byte width for that clone
    // before the parser starts growing its object/message vectors.
    budget.charge_allocations(decoded_len)?;
    let archive = Archive::parse_with_limits(stream.as_bytes(), ctx.archive_limits)
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(|_| SlideAudioCreationError::InvalidSource)?;
    let mut selected = None;
    for (object_index, object) in archive.objects.iter().enumerate() {
        for (message_index, message) in object.messages.iter().enumerate() {
            if message.type_ != PACKAGE_METADATA_MESSAGE_TYPE {
                continue;
            }
            if selected
                .replace((object_index, message_index, message.data.as_slice()))
                .is_some()
            {
                return Err(SlideAudioCreationError::InvalidSource);
            }
            if object
                .archive_info
                .message_infos
                .get(message_index)
                .map(|info| info.type_)
                != Some(PACKAGE_METADATA_MESSAGE_TYPE)
            {
                return Err(SlideAudioCreationError::InvalidSource);
            }
        }
    }
    let (object_index, message_index, _payload) =
        selected.ok_or(SlideAudioCreationError::InvalidSource)?;
    Ok((archive, object_index, message_index))
}

fn media_options(
    payload: &[u8],
    ctx: &CreationContext,
    budget: &CreationBudget,
) -> Result<media_codec::DecodeOptions, SlideAudioCreationError> {
    let wire = ctx.wire_limits;
    let (fields, work, output) = residual_wire_allowances(wire, budget)?;
    let depth = u32::try_from(wire.max_nesting()).unwrap_or(u32::MAX);
    Ok(media_codec::DecodeOptions::new(
        payload.len().min(wire.max_input_bytes()).max(1),
        fields,
        work,
        fields,
        fields,
        fields,
        SHA1_BYTES,
        MAX_FILENAME_BYTES,
        depth,
    )
    .with_max_output_bytes(output))
}

fn identity_options(
    payload: &[u8],
    ctx: &CreationContext,
    budget: &CreationBudget,
) -> Result<identity_codec::RewriteOptions, SlideAudioCreationError> {
    let wire = ctx.wire_limits;
    let (fields, work, output) = residual_wire_allowances(wire, budget)?;
    Ok(identity_codec::RewriteOptions::new(
        payload.len().min(wire.max_input_bytes()).max(1),
        output,
        fields,
        work,
        u32::try_from(wire.max_nesting()).unwrap_or(u32::MAX),
        fields,
        fields,
        16.min(fields),
    ))
}

fn residual_wire_allowances(
    wire: litchi_iwa_common::WireLimits,
    budget: &CreationBudget,
) -> Result<(usize, usize, usize), SlideAudioCreationError> {
    let fields = wire.max_fields().min(budget.remaining_wire_fields());
    if fields == 0 {
        return Err(SlideAudioCreationError::LimitExceeded {
            kind: SlideAudioCreationLimitKind::WireFields,
            observed: 1,
            maximum: 0,
        });
    }
    let work = wire.max_rewrite_work().min(budget.remaining_work());
    if work == 0 {
        return Err(SlideAudioCreationError::LimitExceeded {
            kind: SlideAudioCreationLimitKind::WireWork,
            observed: 1,
            maximum: 0,
        });
    }
    let output = wire.max_output_bytes().min(budget.remaining_output());
    if output == 0 {
        return Err(SlideAudioCreationError::LimitExceeded {
            kind: SlideAudioCreationLimitKind::OutputBytes,
            observed: 1,
            maximum: 0,
        });
    }
    Ok((fields, work, output))
}

fn resolve_data(
    catalog: &SourceCatalog,
    facts: &MediaFacts,
    digest: &[u8; SHA1_BYTES],
    filename: &str,
    bytes: &[u8],
    ids: &CreationIds,
    budget: &mut CreationBudget,
) -> Result<(u64, Option<Box<str>>, Option<Box<str>>, bool, usize), SlideAudioCreationError> {
    if facts.duplicate_match {
        return Err(SlideAudioCreationError::InvalidSource);
    }
    if let (Some(identifier), Some(name)) =
        (facts.matching_identifier, facts.matching_name.as_deref())
    {
        let record = facts
            .data
            .iter()
            .find(|record| record.identifier == identifier && &record.digest == digest)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        let name = valid_data_name(name)?;
        let full_name = data_member_name(name, budget)?;
        let mut entries = catalog
            .package()
            .iter()
            .filter(|entry| entry.name() == full_name);
        let entry = entries
            .next()
            .ok_or(SlideAudioCreationError::InvalidSource)?;
        if entries.next().is_some()
            || entry.is_opaque()
            || record.materialized_length
                != Some(
                    u64::try_from(bytes.len())
                        .map_err(|_| SlideAudioCreationError::InvalidSource)?,
                )
            || entry.data() != bytes
            || MediaType::from_bytes(entry.data()) != MediaType::Audio
        {
            return Err(SlideAudioCreationError::InvalidSource);
        }
        return Ok((record.identifier, None, None, true, 0));
    }

    let identifier = allocate_data_identifier(catalog, facts, ids, budget)?;
    let name = generated_name(filename, identifier, catalog, budget)?;
    let full_name = data_member_name(&name, budget)?.into_boxed_str();
    budget.charge_entries(1)?;
    Ok((
        identifier,
        Some(full_name),
        Some(name.into_boxed_str()),
        false,
        1,
    ))
}

fn valid_data_name(name: &str) -> Result<&str, SlideAudioCreationError> {
    if name.is_empty()
        || name.len() > MAX_FILENAME_BYTES
        || name
            .bytes()
            .any(|byte| byte == 0 || byte.is_ascii_control())
        || name.contains(['/', '\\'])
        || Path::new(name)
            .components()
            .any(|component| !matches!(component, Component::Normal(_)))
    {
        return Err(SlideAudioCreationError::InvalidSource);
    }
    Ok(name)
}

fn data_member_name(
    name: &str,
    budget: &mut CreationBudget,
) -> Result<String, SlideAudioCreationError> {
    let length = "Data/"
        .len()
        .checked_add(name.len())
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let mut full = String::new();
    budget.charge_allocations(length)?;
    full.try_reserve_exact(length)
        .map_err(|_| SlideAudioCreationError::Allocation { amount: length })?;
    full.push_str("Data/");
    full.push_str(name);
    Ok(full)
}

fn allocate_data_identifier(
    catalog: &SourceCatalog,
    facts: &MediaFacts,
    ids: &CreationIds,
    budget: &mut CreationBudget,
) -> Result<u64, SlideAudioCreationError> {
    let mut capacity = facts
        .data
        .len()
        .checked_add(5)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    for component in catalog.components().iter() {
        for object in &component.archive().objects {
            capacity = capacity
                .checked_add(usize::from(object.archive_info.identifier.is_some()))
                .ok_or(SlideAudioCreationError::InvalidSource)?;
            for info in &object.archive_info.message_infos {
                capacity = capacity
                    .checked_add(info.data_references.len())
                    .and_then(|value| value.checked_add(info.object_references.len()))
                    .ok_or(SlideAudioCreationError::InvalidSource)?;
                for field in &info.field_infos {
                    capacity = capacity
                        .checked_add(field.data_references.len())
                        .and_then(|value| value.checked_add(field.object_references.len()))
                        .ok_or(SlideAudioCreationError::InvalidSource)?;
                }
            }
        }
    }
    budget.charge_allocations(capacity)?;
    let mut used = HashSet::new();
    used.try_reserve(capacity)
        .map_err(|_| SlideAudioCreationError::Allocation { amount: capacity })?;
    for component in catalog.components().iter() {
        for object in &component.archive().objects {
            if let Some(identifier) = object.archive_info.identifier {
                used.insert(identifier);
            }
            for info in &object.archive_info.message_infos {
                for identifier in info.data_references.iter().chain(&info.object_references) {
                    used.insert(*identifier);
                }
                for field in &info.field_infos {
                    for identifier in field.data_references.iter().chain(&field.object_references) {
                        used.insert(*identifier);
                    }
                }
            }
        }
    }
    for record in &facts.data {
        used.insert(record.identifier);
    }
    for identifier in ids.object_identifiers() {
        used.insert(identifier);
    }
    let mut candidate = facts
        .data
        .iter()
        .map(|record| record.identifier)
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    while candidate == 0 || used.contains(&candidate) {
        candidate = candidate
            .checked_add(1)
            .ok_or(SlideAudioCreationError::InvalidSource)?;
    }
    Ok(candidate)
}

fn decimal_len(mut value: u64) -> usize {
    let mut length = 1;
    while value >= 10 {
        value /= 10;
        length += 1;
    }
    length
}

fn generated_name(
    source_name: &str,
    identifier: u64,
    catalog: &SourceCatalog,
    budget: &mut CreationBudget,
) -> Result<String, SlideAudioCreationError> {
    let (stem, extension) = source_name
        .rsplit_once('.')
        .ok_or(SlideAudioCreationError::InvalidFilename)?;
    let identifier_digits = decimal_len(identifier);
    let normal_length = stem
        .len()
        .checked_add(1)
        .and_then(|length| length.checked_add(identifier_digits))
        .and_then(|length| length.checked_add(1))
        .and_then(|length| length.checked_add(extension.len()))
        .ok_or(SlideAudioCreationError::InvalidSource)?;
    let fallback = normal_length > MAX_FILENAME_BYTES;
    for attempt in 0..MAX_NAME_ATTEMPTS {
        let (candidate_length, candidate_stem) = if attempt == 0 {
            if fallback {
                (
                    "litchi"
                        .len()
                        .checked_add(1)
                        .and_then(|length| length.checked_add(identifier_digits))
                        .and_then(|length| length.checked_add(1))
                        .and_then(|length| length.checked_add(extension.len()))
                        .ok_or(SlideAudioCreationError::InvalidSource)?,
                    "litchi",
                )
            } else {
                (normal_length, stem)
            }
        } else {
            let attempt_digits = decimal_len(attempt);
            let length = stem
                .len()
                .checked_add(1)
                .and_then(|length| length.checked_add(identifier_digits))
                .and_then(|length| length.checked_add(1))
                .and_then(|length| length.checked_add(attempt_digits))
                .and_then(|length| length.checked_add(1))
                .and_then(|length| length.checked_add(extension.len()))
                .ok_or(SlideAudioCreationError::InvalidSource)?;
            (length, stem)
        };
        if candidate_length > MAX_FILENAME_BYTES {
            continue;
        }
        budget.charge_allocations(candidate_length)?;
        let mut candidate = String::new();
        candidate.try_reserve_exact(candidate_length).map_err(|_| {
            SlideAudioCreationError::Allocation {
                amount: candidate_length,
            }
        })?;
        if attempt == 0 {
            write!(&mut candidate, "{candidate_stem}-{identifier}.{extension}")
                .map_err(|_| SlideAudioCreationError::InvalidSource)?;
        } else {
            write!(
                &mut candidate,
                "{candidate_stem}-{identifier}-{attempt}.{extension}"
            )
            .map_err(|_| SlideAudioCreationError::InvalidSource)?;
        }
        let full = data_member_name(&candidate, budget)?;
        let present = catalog.package().iter().any(|entry| entry.name() == full);
        if !present {
            return Ok(candidate);
        }
    }
    Err(SlideAudioCreationError::InvalidSource)
}

fn charge_identity_requirements(
    budget: &mut CreationBudget,
    prepared: &identity_codec::PreparedPackageMetadataAdditionSaveTokenRewrite<'_, '_>,
) -> Result<(), SlideAudioCreationError> {
    let requirements = prepared.execution_requirements();
    budget.charge_output(requirements.output_bytes())?;
    budget.charge_total(requirements.retained_bytes())?;
    budget.charge_wire_fields(requirements.fields())?;
    budget.charge_work(requirements.work_bytes())?;
    budget.charge_references(requirements.references())?;
    budget.charge_allocations(requirements.allocations())?;
    budget.charge_nesting(prepared.prepare_report().max_depth() as usize)
}

fn charge_identity_inspection(
    budget: &mut CreationBudget,
    report: identity_codec::RewriteReport,
) -> Result<(), SlideAudioCreationError> {
    budget.charge_wire_fields(report.fields())?;
    budget.charge_work(report.work_bytes())?;
    budget.charge_references(report.references_scanned())?;
    budget.charge_allocations(report.allocations())?;
    budget.charge_nesting(report.max_depth() as usize)
}

fn charge_media_requirements(
    budget: &mut CreationBudget,
    prepared: &media_codec::PreparedPackageMetadataMediaRewrite<'_>,
) -> Result<(), SlideAudioCreationError> {
    let requirements = prepared.execution_requirements();
    budget.charge_output(requirements.output_bytes())?;
    budget.charge_total(requirements.retained_bytes())?;
    budget.charge_wire_fields(requirements.fields())?;
    budget.charge_work(requirements.work_bytes())?;
    budget.charge_references(requirements.owners())?;
    budget.charge_allocations(requirements.allocations())?;
    budget.charge_nesting(prepared.source_report().max_depth() as usize)
}

fn map_identity_error(error: identity_codec::RewriteError) -> SlideAudioCreationError {
    if let Some(amount) = error.allocation_request() {
        return SlideAudioCreationError::Allocation { amount };
    }
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            identity_codec::RewriteLimit::InputBytes { observed, maximum } => {
                (SlideAudioCreationLimitKind::InputBytes, observed, maximum)
            },
            identity_codec::RewriteLimit::OutputBytes { observed, maximum } => {
                (SlideAudioCreationLimitKind::OutputBytes, observed, maximum)
            },
            identity_codec::RewriteLimit::Fields { observed, maximum } => {
                (SlideAudioCreationLimitKind::WireFields, observed, maximum)
            },
            identity_codec::RewriteLimit::Work { observed, maximum } => {
                (SlideAudioCreationLimitKind::WireWork, observed, maximum)
            },
            identity_codec::RewriteLimit::Nesting { observed, maximum } => (
                SlideAudioCreationLimitKind::WireNesting,
                observed as usize,
                maximum as usize,
            ),
            identity_codec::RewriteLimit::References { observed, maximum } => {
                (SlideAudioCreationLimitKind::References, observed, maximum)
            },
            identity_codec::RewriteLimit::Components { observed, maximum }
            | identity_codec::RewriteLimit::Additions { observed, maximum } => {
                (SlideAudioCreationLimitKind::WireFields, observed, maximum)
            },
            _ => return SlideAudioCreationError::InvalidSource,
        };
        return SlideAudioCreationError::LimitExceeded {
            kind,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        };
    }
    SlideAudioCreationError::InvalidSource
}

fn map_media_error(
    error: media_codec::DecodeError,
    fallback: SlideAudioCreationLimitKind,
) -> SlideAudioCreationError {
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            media_codec::DecodeLimit::Bytes { observed, maximum } => {
                (SlideAudioCreationLimitKind::InputBytes, observed, maximum)
            },
            media_codec::DecodeLimit::Fields { observed, maximum } => {
                (SlideAudioCreationLimitKind::WireFields, observed, maximum)
            },
            media_codec::DecodeLimit::Work { observed, maximum } => {
                (SlideAudioCreationLimitKind::WireWork, observed, maximum)
            },
            media_codec::DecodeLimit::OutputBytes { observed, maximum } => {
                (SlideAudioCreationLimitKind::OutputBytes, observed, maximum)
            },
            media_codec::DecodeLimit::Nesting { observed, maximum } => (
                SlideAudioCreationLimitKind::WireNesting,
                observed as usize,
                maximum as usize,
            ),
            media_codec::DecodeLimit::Components { observed, maximum }
            | media_codec::DecodeLimit::DataRecords { observed, maximum }
            | media_codec::DecodeLimit::Owners { observed, maximum } => {
                (SlideAudioCreationLimitKind::References, observed, maximum)
            },
            media_codec::DecodeLimit::DigestBytes { observed, maximum }
            | media_codec::DecodeLimit::NameBytes { observed, maximum } => {
                (SlideAudioCreationLimitKind::EntryBytes, observed, maximum)
            },
            _ => return SlideAudioCreationError::InvalidSource,
        };
        return SlideAudioCreationError::LimitExceeded {
            kind,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        };
    }
    let _ = fallback;
    SlideAudioCreationError::InvalidSource
}

fn map_media_rewrite_error(error: media_codec::RewriteError) -> SlideAudioCreationError {
    if let Some(amount) = error.allocation_request() {
        return SlideAudioCreationError::Allocation { amount };
    }
    if let Some(limit) = error.resource_limit() {
        let (kind, observed, maximum) = match limit {
            media_codec::DecodeLimit::Bytes { observed, maximum } => {
                (SlideAudioCreationLimitKind::WireWork, observed, maximum)
            },
            media_codec::DecodeLimit::Fields { observed, maximum } => {
                (SlideAudioCreationLimitKind::WireFields, observed, maximum)
            },
            media_codec::DecodeLimit::Work { observed, maximum } => {
                (SlideAudioCreationLimitKind::WireWork, observed, maximum)
            },
            media_codec::DecodeLimit::OutputBytes { observed, maximum } => {
                (SlideAudioCreationLimitKind::OutputBytes, observed, maximum)
            },
            media_codec::DecodeLimit::Components { observed, maximum }
            | media_codec::DecodeLimit::DataRecords { observed, maximum }
            | media_codec::DecodeLimit::Owners { observed, maximum } => {
                (SlideAudioCreationLimitKind::References, observed, maximum)
            },
            media_codec::DecodeLimit::DigestBytes { observed, maximum }
            | media_codec::DecodeLimit::NameBytes { observed, maximum } => {
                (SlideAudioCreationLimitKind::EntryBytes, observed, maximum)
            },
            media_codec::DecodeLimit::Nesting { observed, maximum } => (
                SlideAudioCreationLimitKind::WireNesting,
                observed as usize,
                maximum as usize,
            ),
            _ => return SlideAudioCreationError::InvalidSource,
        };
        return SlideAudioCreationError::LimitExceeded {
            kind,
            observed: u64::try_from(observed).unwrap_or(u64::MAX),
            maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
        };
    }
    SlideAudioCreationError::InvalidSource
}
