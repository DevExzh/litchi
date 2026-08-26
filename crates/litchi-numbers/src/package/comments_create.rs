//! Strict same-member root-comment creation.

use std::sync::Arc;

use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts};
use litchi_iwa_common::{
    WireLimits,
    wire::{append_length_delimited_field, append_varint_field},
};
use litchi_iwa_core::archive::{FieldObjectReferenceTransition, ObjectReferenceTransition};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::{
    comment_storage_codec::{
        CommentStorageLeafWrite, DateSnapshot, DecodeOptions, UuidSnapshot,
        prepare_comment_storage_leaf_write,
    },
    numbers_table_cell_storage_codec,
    package_metadata_codec::{
        CombinedBatch, CombinedSaveTokenBatch, RewriteOptions as MetadataRewriteOptions,
        SaveTokenBatch,
    },
};

use crate::{SheetSelector, TableSelector};

use crate::package::{comments_metadata, metadata, table_cell_pop_up_menu_native};

use super::{
    Comment, Commit, Diagnostics, Error, Located, MessageRoute, Package, Patch, Path,
    cell_position_for_path, message_at_route, physical_source, root_preview_deletions,
    table_cell_decode_options,
};

const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const COMMENT_LIST_TYPE: i32 = 10;

pub(super) fn create_root_comment(
    source: &Package,
    located: &Located,
    before: Option<Comment>,
    after: Comment,
) -> Result<Commit, Error> {
    let path = located.target.path;
    if before.is_some()
        || located.comment.is_some()
        || located.comment_key.is_some()
        || located.storage.is_some()
    {
        return Err(Error::InvalidSource { path });
    }
    let tile = located.tile.ok_or(Error::UnsupportedDependency { path })?;
    let source_cell = located
        .cell_bytes
        .as_ref()
        .cloned()
        .ok_or(Error::UnsupportedDependency { path })?;
    let (list, author_identifier) = creation_graph(source, path)?;
    if list.component_index != tile.component_index {
        return Err(Error::UnsupportedDependency { path });
    }
    let physical = physical_source(source)?;
    if !physical.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    let metadata_source =
        comments_metadata::strict_source(source).map_err(|_| Error::InvalidSource { path })?;
    let metadata_options = metadata_options(metadata_source.payload().len(), 1);
    let registry = comments_metadata::inspect(source, metadata_options)
        .map_err(|_| Error::InvalidSource { path })?;
    comments_metadata::reject_unknown_archive_metadata(
        source,
        physical
            .limits()
            .effective_archive_limits()
            .map_err(|_| Error::InvalidSource { path })?,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    registry
        .current_uuid_if_registered(list.component_index, author_identifier)
        .map_err(|_| Error::InvalidSource { path })?;
    let fresh = registry
        .allocate_identifiers(1)
        .map_err(|_| Error::UnsupportedDependency { path })?
        .into_iter()
        .next()
        .ok_or(Error::Verification)?;
    let list_source = message_at_route(source, list, path)?.data.as_slice();
    let (key, list_data) = append_comment_entry(source, list_source, path, fresh.identifier)?;
    let storage_data = canonical_storage(source, after.text(), author_identifier, fresh, path)?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let storage_object = storage_object(
        fresh.identifier,
        storage_data,
        author_identifier,
        archive_limits,
        path,
    )?;
    let target_cell = attach_comment_cell(located, key, path)?;

    let component = source
        .state
        .components
        .catalog()
        .get_index(list.component_index)
        .ok_or(Error::InvalidSource { path })?;
    let entry = physical
        .package()
        .iter()
        .find(|entry| entry.name() == component.name() && !entry.is_opaque())
        .ok_or(Error::UnsupportedSource)?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical
            .limits()
            .snappy_limits()
            .map_err(|_| Error::InvalidSource { path })?,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let mut archive = Archive::parse_with_limits(stream.as_bytes(), archive_limits)
        .map_err(|_| Error::InvalidSource { path })?;
    archive
        .validate_canonical_object_framing(stream.as_bytes())
        .map_err(|_| Error::InvalidSource { path })?;
    replace_list_with_transition(
        &mut archive,
        list,
        list_data,
        key,
        fresh.identifier,
        archive_limits,
        path,
    )?;
    replace_tile(
        &mut archive,
        source,
        tile,
        located,
        key,
        archive_limits,
        path,
    )?;
    archive
        .objects
        .try_reserve(1)
        .map_err(|_| Error::Allocation { amount: 1, path })?;
    archive.objects.push(storage_object);
    archive
        .objects
        .sort_by_key(|object| object.archive_info.identifier.unwrap_or_default());
    let member_bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(|_| Error::InvalidSource { path })?;
    let member_compressed =
        SnappyStream::compress(&member_bytes).map_err(|_| Error::InvalidSource { path })?;

    let additions = [registry
        .uuid_addition(list.component_index, fresh)
        .map_err(|_| Error::UnsupportedDependency { path })?];
    let selectors = registry
        .touched_selectors(&[list.component_index])
        .map_err(|_| Error::UnsupportedDependency { path })?;
    let transition = CombinedBatch::new(
        registry.last_object_identifier(),
        fresh.identifier,
        &additions,
        &[],
        &[],
        &[],
        &[],
    );
    let prepared = comments_metadata::prepare_combined(
        &registry,
        CombinedSaveTokenBatch::new(transition, SaveTokenBatch::new(&selectors)),
        metadata_options,
    )
    .map_err(|_| Error::UnsupportedDependency { path })?;
    let requirements = prepared.execution_requirements();
    let metadata_payload = prepared
        .execute(requirements.exact_limits())
        .map_err(|_| Error::Verification)?
        .into_bytes();
    let metadata_compressed = pack_metadata(source, metadata_payload, path)?;
    let edits = [
        EntryEdit::new(component.name(), member_compressed.as_slice()),
        EntryEdit::new(metadata::ENTRY_NAME, metadata_compressed.as_slice()),
    ];
    let previews = root_preview_deletions(physical)?;
    let deleted = previews.iter().map(String::as_str).collect::<Vec<_>>();
    let output = physical
        .package()
        .reassemble_with_deletions_to_bytes(&edits, &deleted, physical.limits())
        .map_err(|_| Error::InvalidSource { path })?;
    let package = Package::from_shared_bytes_with_options(output.into(), source.state.options)
        .map_err(|_| Error::Verification)?;
    let position = cell_position_for_path(path)?;
    if package.table_cell_comment(
        SheetSelector::index(path.sheet()),
        TableSelector::index(path.table()),
        position,
    )? != Some(after.clone())
    {
        return Err(Error::Verification);
    }
    let target_owner = physical_source(&package)?.__source_owner();
    Ok(Commit {
        package,
        patch: Patch {
            artifacts: OwnedExactArtifacts::new(physical.__source_owner(), target_owner),
            path,
            before: None,
            after: Some(after),
            source_cell,
            target_cell: Arc::from(target_cell.into_boxed_slice()),
            source_previews: previews.len(),
            target_previews: 0,
            touched_components: 1,
        },
        diagnostics: Diagnostics::published(1, previews.len()),
    })
}

fn creation_graph(source: &Package, path: Path) -> Result<(MessageRoute, u64), Error> {
    let census = super::census_comment_ownership(source, path)?;
    if census.comment_list_ids.len() != 1 || !census.rooted_segment_ids.is_empty() {
        return Err(Error::UnsupportedDependency { path });
    }
    super::validate_comment_authors(source, &census, path)?;
    let table_id = census.comment_list_ids[0];
    let resolved = source
        .state
        .index
        .resolve_ref_id(&source.state.components, table_id)
        .map_err(|_| Error::InvalidSource { path })?
        .ok_or(Error::InvalidSource { path })?;
    let mut message_index = None;
    for (index, message) in resolved.messages.iter().enumerate() {
        let snapshot = numbers_table_cell_storage_codec::decode_table_data_list_type_with_report(
            message.data.as_slice(),
            table_cell_decode_options(
                source,
                message.data.len(),
                source.state.options.semantic().max_references(),
            ),
        )
        .map_err(|error| super::map_table_codec_error(error, path))?
        .0;
        if snapshot.list_type() != COMMENT_LIST_TYPE {
            continue;
        }
        if message_index.replace(index).is_some() {
            return Err(Error::InvalidSource { path });
        }
    }
    let message_index = message_index.ok_or(Error::InvalidSource { path })?;
    let mut authors = census.author_ids;
    authors.sort_unstable();
    authors.dedup();
    if authors.len() != 1 {
        return Err(Error::UnsupportedDependency { path });
    }
    Ok((
        MessageRoute {
            component_index: resolved.component_index,
            object_index: resolved.object_index,
            message_index,
            message_type: resolved.messages[message_index].type_,
        },
        authors[0],
    ))
}

fn append_comment_entry(
    source: &Package,
    list: &[u8],
    path: Path,
    identifier: u64,
) -> Result<(u32, Vec<u8>), Error> {
    let snapshot = numbers_table_cell_storage_codec::decode_table_data_list_with_report(
        list,
        table_cell_decode_options(
            source,
            list.len(),
            source.state.options.semantic().max_references(),
        ),
    )
    .map_err(|error| super::map_table_codec_error(error, path))?
    .0;
    if snapshot.list_type() != COMMENT_LIST_TYPE {
        return Err(Error::UnsupportedDependency { path });
    }
    let key = snapshot.next_list_id().max(2);
    let mut reference = Vec::new();
    append_varint_field(&mut reference, 1, identifier)
        .map_err(|_| Error::InvalidSource { path })?;
    let mut entry = Vec::new();
    append_varint_field(&mut entry, 1, u64::from(key))
        .map_err(|_| Error::InvalidSource { path })?;
    append_varint_field(&mut entry, 2, 1).map_err(|_| Error::InvalidSource { path })?;
    append_length_delimited_field(&mut entry, 10, &reference)
        .map_err(|_| Error::InvalidSource { path })?;
    let mut output = list.to_vec();
    append_length_delimited_field(&mut output, 3, &entry)
        .map_err(|_| Error::InvalidSource { path })?;
    litchi_iwa_common::wire::patch_varint_field(
        &output,
        2,
        true,
        Some(u64::from(
            key.checked_add(1).ok_or(Error::InvalidSource { path })?,
        )),
    )
    .map(|bytes| (key, bytes))
    .map_err(|_| Error::InvalidSource { path })
}

fn canonical_storage(
    source: &Package,
    text: &str,
    author: u64,
    fresh: comments_metadata::FreshIdentifier,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let options = DecodeOptions::new(
        WireLimits::MAX_INPUT_BYTES,
        WireLimits::MAX_FIELDS,
        WireLimits::MAX_REWRITE_WORK,
        64,
        source.state.options.semantic().max_references().max(1),
        source
            .state
            .options
            .semantic()
            .max_output_text_bytes()
            .max(1),
    );
    let prepared = prepare_comment_storage_leaf_write(
        CommentStorageLeafWrite::new(
            text,
            DateSnapshot::from_bits(0),
            author,
            UuidSnapshot::from_parts(fresh.uuid.lower(), fresh.uuid.upper()),
        ),
        options,
    )
    .map_err(|error| super::map_comment_codec_error(error, path))?;
    prepared
        .execute(prepared.execution_requirements().exact_limits())
        .map(|output| output.into_bytes())
        .map_err(|error| super::map_comment_codec_error(error, path))
}

fn storage_object(
    identifier: u64,
    data: Vec<u8>,
    author: u64,
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<ArchiveObject, Error> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: COMMENT_STORAGE_MESSAGE_TYPE,
            data,
        }],
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource { path })?;
    info.object_references = vec![author];
    let mut field = FieldInfo::new(vec![3]);
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references = vec![author];
    info.field_infos.push(field);
    object
        .validate_with_limits(limits)
        .map_err(|_| Error::InvalidSource { path })?;
    Ok(object)
}

fn replace_list_with_transition(
    archive: &mut Archive,
    route: MessageRoute,
    data: Vec<u8>,
    key: u32,
    identifier: u64,
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<(), Error> {
    let object = archive
        .objects
        .get_mut(route.object_index)
        .ok_or(Error::InvalidSource { path })?;
    let info = object
        .archive_info
        .message_infos
        .get(route.message_index)
        .ok_or(Error::InvalidSource { path })?
        .clone();
    let before = info.object_references.clone();
    if before.contains(&identifier) {
        return Err(Error::InvalidSource { path });
    }
    let mut after = before.clone();
    after.push(identifier);
    let fields = info
        .field_infos
        .iter()
        .enumerate()
        .map(|(field_info_index, field)| FieldObjectReferenceTransition {
            field_info_index,
            expected_path: field.path.as_slice(),
            before: field.object_references.as_slice(),
            after: field.object_references.as_slice(),
        })
        .collect::<Vec<_>>();
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            route.message_index,
            RawMessage {
                type_: route.message_type,
                data,
            },
            ObjectReferenceTransition {
                aggregate_before: &before,
                aggregate_after: &after,
                fields: &fields,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource { path })?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(route.message_index)
        .ok_or(Error::InvalidSource { path })?;
    let mut field = FieldInfo::new(vec![3, key]);
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references = vec![identifier];
    info.field_infos.push(field);
    Ok(())
}

fn attach_comment_cell(located: &Located, key: u32, path: Path) -> Result<Vec<u8>, Error> {
    let mut cell = litchi_numbers_wire::BncCell::parse(
        located
            .cell_bytes
            .as_deref()
            .ok_or(Error::UnsupportedDependency { path })?,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    cell.set_comment_identifier(Some(key));
    cell.try_encode_with_limit(WireLimits::MAX_OUTPUT_BYTES)
        .map_err(|_| Error::InvalidSource { path })
}

fn replace_tile(
    archive: &mut Archive,
    source: &Package,
    route: MessageRoute,
    located: &Located,
    key: u32,
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<(), Error> {
    let tile_source = message_at_route(source, route, path)?.data.as_slice();
    let target = attach_comment_cell(located, key, path)?;
    let rewritten = table_cell_pop_up_menu_native::patch_tile_cell(
        tile_source,
        u32::try_from(located.target.row % located.target.tile_size)
            .map_err(|_| Error::InvalidSource { path })?,
        u32::try_from(located.target.column).map_err(|_| Error::InvalidSource { path })?,
        &target,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    archive
        .objects
        .get_mut(route.object_index)
        .ok_or(Error::InvalidSource { path })?
        .replace_message_preserving_header_with_limits(
            route.message_index,
            RawMessage {
                type_: route.message_type,
                data: rewritten,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource { path })?;
    Ok(())
}

fn metadata_options(bytes: usize, additions: usize) -> MetadataRewriteOptions {
    MetadataRewriteOptions::new(
        bytes.max(1),
        bytes.saturating_mul(2).max(1),
        bytes.saturating_mul(16).max(1),
        bytes.saturating_mul(128).max(1),
        64,
        bytes.max(1),
        bytes.max(1),
        additions,
    )
}

fn pack_metadata(source: &Package, payload: Vec<u8>, path: Path) -> Result<Vec<u8>, Error> {
    let physical = physical_source(source)?;
    let route = metadata::unique_message_route(source).ok_or(Error::InvalidSource { path })?;
    let entry = physical
        .package()
        .iter()
        .find(|entry| entry.name() == metadata::ENTRY_NAME)
        .ok_or(Error::InvalidSource { path })?;
    let limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let stream = SnappyStream::decompress_with_limits(
        entry.data(),
        physical
            .limits()
            .snappy_limits()
            .map_err(|_| Error::InvalidSource { path })?,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let mut archive = Archive::parse_with_limits(stream.as_bytes(), limits)
        .map_err(|_| Error::InvalidSource { path })?;
    archive
        .objects
        .get_mut(route.object_index)
        .ok_or(Error::InvalidSource { path })?
        .replace_message_preserving_header_with_limits(
            route.message_index,
            RawMessage {
                type_: metadata::MESSAGE_TYPE,
                data: payload,
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource { path })?;
    let bytes = archive
        .to_bytes_with_limits(limits)
        .map_err(|_| Error::InvalidSource { path })?;
    SnappyStream::compress(&bytes).map_err(|_| Error::InvalidSource { path })
}
