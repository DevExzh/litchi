//! Strict same-member root-comment creation.

use std::sync::Arc;

use litchi_iwa_archive::package::{EntryEdit, OwnedExactArtifacts};
use litchi_iwa_common::{
    WireLimits,
    wire::{
        append_length_delimited_field, append_repeated_length_delimited_field, append_varint_field,
    },
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

use crate::package::{comments_metadata, metadata, table_cell_edit::tile};

use super::{
    Comment, Commit, Diagnostics, Error, Located, MessageRoute, Package, Patch, Path,
    cell_position_for_path, message_at_route, physical_source, root_preview_deletions,
    table_cell_decode_options,
};

const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const COMMENT_LIST_TYPE: i32 = 10;
const ANNOTATION_AUTHOR_MESSAGE_TYPE: u32 = 212;
const ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE: u32 = 213;
const GENERATED_AUTHOR_NAME: &str = "litchi-iwa";

#[derive(Clone, Copy)]
struct CreationGraph {
    list: Option<MessageRoute>,
    model: MessageRoute,
    author_identifier: u64,
    author_storage: Option<MessageRoute>,
    create_author_storage: bool,
}

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
    let source_cell = located.cell_bytes.as_ref().cloned().unwrap_or_default();
    let graph = creation_graph(source, located, path)?;
    let owner_component_index = graph
        .list
        .map_or(graph.model.component_index, |route| route.component_index);
    if owner_component_index != tile.component_index
        || graph
            .author_storage
            .is_some_and(|route| route.component_index != tile.component_index)
    {
        return Err(Error::UnsupportedDependency { path });
    }
    let physical = physical_source(source)?;
    if !physical.source_is_exact() {
        return Err(Error::UnsupportedSource);
    }
    let metadata_source =
        comments_metadata::strict_source(source).map_err(|_| Error::InvalidSource { path })?;
    let creates_author = graph.author_storage.is_some() || graph.create_author_storage;
    let allocation_count = 1
        + usize::from(creates_author)
        + usize::from(graph.create_author_storage)
        + usize::from(graph.list.is_none());
    let metadata_options = metadata_options(metadata_source.payload().len(), allocation_count);
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
    if graph.author_storage.is_none() && !graph.create_author_storage {
        registry
            .current_uuid_if_registered(owner_component_index, graph.author_identifier)
            .map_err(|_| Error::InvalidSource { path })?;
    }
    let fresh = registry
        .allocate_identifiers(allocation_count)
        .map_err(|_| Error::UnsupportedDependency { path })?;
    let list_fresh = graph.list.is_none().then(|| fresh[0]);
    let comment_index = usize::from(list_fresh.is_some());
    let comment_fresh = fresh
        .get(comment_index)
        .copied()
        .ok_or(Error::Verification)?;
    let author_fresh = creates_author
        .then(|| fresh.get(comment_index + 1).copied())
        .flatten();
    let author_storage_fresh = graph
        .create_author_storage
        .then(|| fresh.get(comment_index + 2).copied())
        .flatten();
    let author_identifier = author_fresh
        .map(|fresh| fresh.identifier)
        .unwrap_or(graph.author_identifier);
    let list_source = match graph.list {
        Some(route) => message_at_route(source, route, path)?.data.clone(),
        None => canonical_empty_comment_list(path)?,
    };
    let (key, list_data) = append_comment_entry(
        source,
        list_source.as_slice(),
        path,
        comment_fresh.identifier,
    )?;
    let storage_data =
        canonical_storage(source, after.text(), author_identifier, comment_fresh, path)?;
    let archive_limits = physical
        .limits()
        .effective_archive_limits()
        .map_err(|_| Error::InvalidSource { path })?;
    let storage_object = storage_object(
        comment_fresh.identifier,
        storage_data,
        author_identifier,
        archive_limits,
        path,
    )?;

    let component = source
        .state
        .components
        .catalog()
        .get_index(owner_component_index)
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
    if let Some(route) = graph.list {
        replace_list_with_transition(
            &mut archive,
            route,
            list_data,
            key,
            comment_fresh.identifier,
            archive_limits,
            path,
        )?;
    } else {
        let fresh = list_fresh.ok_or(Error::Verification)?;
        attach_comment_list_to_model(
            &mut archive,
            source,
            graph.model,
            fresh.identifier,
            archive_limits,
            path,
        )?;
        archive
            .objects
            .try_reserve(1)
            .map_err(|_| Error::Allocation { amount: 1, path })?;
        archive.objects.push(comment_list_object(
            fresh.identifier,
            list_data,
            key,
            comment_fresh.identifier,
            path,
        )?);
    }
    let target_cell = replace_tile(
        &mut archive,
        source,
        tile,
        located,
        key,
        archive_limits,
        path,
    )?;
    if let (Some(route), Some(fresh)) = (graph.author_storage, author_fresh) {
        append_author_storage(
            &mut archive,
            source,
            route,
            fresh.identifier,
            archive_limits,
            path,
        )?;
        archive
            .objects
            .try_reserve(2)
            .map_err(|_| Error::Allocation { amount: 2, path })?;
        archive.objects.push(author_object(fresh.identifier, path)?);
    } else if let (Some(author), Some(storage)) = (author_fresh, author_storage_fresh) {
        archive
            .objects
            .try_reserve(2)
            .map_err(|_| Error::Allocation { amount: 2, path })?;
        archive
            .objects
            .push(author_object(author.identifier, path)?);
        archive.objects.push(author_storage_object(
            storage.identifier,
            author.identifier,
            path,
        )?);
    } else {
        archive
            .objects
            .try_reserve(1)
            .map_err(|_| Error::Allocation { amount: 1, path })?;
    }
    archive.objects.push(storage_object);
    archive
        .objects
        .sort_by_key(|object| object.archive_info.identifier.unwrap_or_default());
    let member_bytes = archive
        .to_bytes_with_limits(archive_limits)
        .map_err(|_| Error::InvalidSource { path })?;
    let member_compressed =
        SnappyStream::compress(&member_bytes).map_err(|_| Error::InvalidSource { path })?;

    let mut additions = Vec::new();
    additions
        .try_reserve_exact(allocation_count)
        .map_err(|_| Error::Allocation {
            amount: allocation_count,
            path,
        })?;
    additions.push(
        registry
            .uuid_addition(owner_component_index, comment_fresh)
            .map_err(|_| Error::UnsupportedDependency { path })?,
    );
    if let Some(fresh) = author_fresh {
        additions.push(
            registry
                .uuid_addition(owner_component_index, fresh)
                .map_err(|_| Error::UnsupportedDependency { path })?,
        );
    }
    if let Some(fresh) = author_storage_fresh {
        additions.push(
            registry
                .uuid_addition(owner_component_index, fresh)
                .map_err(|_| Error::UnsupportedDependency { path })?,
        );
    }
    if let Some(fresh) = list_fresh {
        additions.push(
            registry
                .uuid_addition(owner_component_index, fresh)
                .map_err(|_| Error::UnsupportedDependency { path })?,
        );
    }
    let selectors = registry
        .touched_selectors(&[owner_component_index])
        .map_err(|_| Error::UnsupportedDependency { path })?;
    let transition = CombinedBatch::new(
        registry.last_object_identifier(),
        fresh.last().ok_or(Error::Verification)?.identifier,
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

fn creation_graph(source: &Package, located: &Located, path: Path) -> Result<CreationGraph, Error> {
    let census = super::census_comment_ownership(source, path)?;
    if census.comment_list_ids.len() > 1 || !census.rooted_segment_ids.is_empty() {
        return Err(Error::UnsupportedDependency { path });
    }
    let model = MessageRoute {
        component_index: located.target.native.component_index,
        object_index: located.target.native.object_index,
        message_index: located.target.native.message_index,
        message_type: located.target.native.message_type,
    };
    let list = if let Some(&table_id) = census.comment_list_ids.first() {
        let resolved = source
            .state
            .index
            .resolve_ref_id(&source.state.components, table_id)
            .map_err(|_| Error::InvalidSource { path })?
            .ok_or(Error::InvalidSource { path })?;
        let mut message_index = None;
        for (index, message) in resolved.messages.iter().enumerate() {
            let snapshot =
                numbers_table_cell_storage_codec::decode_table_data_list_type_with_report(
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
        Some(MessageRoute {
            component_index: resolved.component_index,
            object_index: resolved.object_index,
            message_index,
            message_type: resolved.messages[message_index].type_,
        })
    } else {
        let message = message_at_route(source, model, path)?;
        let snapshot = numbers_table_cell_storage_codec::decode_table_model_with_report(
            message.data.as_slice(),
            table_cell_decode_options(
                source,
                message.data.len(),
                source.state.options.semantic().max_references(),
            ),
        )
        .map_err(|error| super::map_table_codec_error(error, path))?
        .0;
        let store = numbers_table_cell_storage_codec::decode_data_store_with_report(
            snapshot.base_data_store(),
            table_cell_decode_options(
                source,
                snapshot.base_data_store().len(),
                source.state.options.semantic().max_references(),
            ),
        )
        .map_err(|error| super::map_table_codec_error(error, path))?
        .0;
        if store.comment_storage_table().is_some() {
            return Err(Error::InvalidSource { path });
        }
        None
    };
    let mut authors = census.author_ids.clone();
    authors.sort_unstable();
    authors.dedup();
    if authors.len() == 1 {
        super::validate_comment_authors(source, &census, path)?;
        return Ok(CreationGraph {
            list,
            model,
            author_identifier: authors[0],
            author_storage: None,
            create_author_storage: false,
        });
    }
    if !authors.is_empty() {
        return Err(Error::UnsupportedDependency { path });
    }
    let author_storage = empty_author_storage(source, path)?;
    Ok(CreationGraph {
        list,
        model,
        author_identifier: 0,
        author_storage,
        create_author_storage: author_storage.is_none(),
    })
}

fn empty_author_storage(source: &Package, path: Path) -> Result<Option<MessageRoute>, Error> {
    let mut found = None;
    for (component_index, component) in source.state.components.catalog().iter().enumerate() {
        for (object_index, object) in component.archive().objects.iter().enumerate() {
            for (message_index, message) in object.messages.iter().enumerate() {
                if message.type_ != ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE {
                    continue;
                }
                let fields = litchi_iwa_common::wire::WireView::parse(message.data.as_slice())
                    .map_err(|_| Error::InvalidSource { path })?;
                if fields.fields().any(|field| field.number() == 1) {
                    return Err(Error::UnsupportedDependency { path });
                }
                if found
                    .replace(MessageRoute {
                        component_index,
                        object_index,
                        message_index,
                        message_type: message.type_,
                    })
                    .is_some()
                {
                    return Err(Error::InvalidSource { path });
                }
            }
        }
    }
    Ok(found)
}

fn canonical_reference(identifier: u64, path: Path) -> Result<Vec<u8>, Error> {
    let mut reference = Vec::new();
    append_varint_field(&mut reference, 1, identifier)
        .map_err(|_| Error::InvalidSource { path })?;
    Ok(reference)
}

fn append_author_storage(
    archive: &mut Archive,
    source: &Package,
    route: MessageRoute,
    identifier: u64,
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<(), Error> {
    let message = message_at_route(source, route, path)?;
    let reference = canonical_reference(identifier, path)?;
    let data = append_repeated_length_delimited_field(message.data.as_slice(), 1, &reference)
        .map_err(|_| Error::InvalidSource { path })?;
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
    if !info.object_references.is_empty() || !info.field_infos.is_empty() {
        return Err(Error::InvalidSource { path });
    }
    object
        .replace_message_transitioning_object_references_preserving_header_with_limits(
            route.message_index,
            RawMessage {
                type_: route.message_type,
                data,
            },
            ObjectReferenceTransition {
                aggregate_before: &[],
                aggregate_after: &[identifier],
                fields: &[],
            },
            limits,
        )
        .map_err(|_| Error::InvalidSource { path })?;
    let info = object
        .archive_info
        .message_infos
        .get_mut(route.message_index)
        .ok_or(Error::InvalidSource { path })?;
    let mut field = FieldInfo::new(vec![1]);
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references = vec![identifier];
    info.field_infos.push(field);
    Ok(())
}

fn author_object(identifier: u64, path: Path) -> Result<ArchiveObject, Error> {
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, GENERATED_AUTHOR_NAME.as_bytes())
        .map_err(|_| Error::InvalidSource { path })?;
    append_varint_field(&mut payload, 4, 0).map_err(|_| Error::InvalidSource { path })?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: ANNOTATION_AUTHOR_MESSAGE_TYPE,
            data: payload,
        }],
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource { path })?;
    info.field_infos.push(FieldInfo::new(vec![4]));
    Ok(object)
}

fn author_storage_object(
    identifier: u64,
    author_identifier: u64,
    path: Path,
) -> Result<ArchiveObject, Error> {
    let reference = canonical_reference(author_identifier, path)?;
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &reference)
        .map_err(|_| Error::InvalidSource { path })?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: ANNOTATION_AUTHOR_STORAGE_MESSAGE_TYPE,
            data: payload,
        }],
    )
    .map_err(|_| Error::InvalidSource { path })?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource { path })?;
    info.object_references = vec![author_identifier];
    let mut field = FieldInfo::new(vec![1]);
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references = vec![author_identifier];
    info.field_infos.push(field);
    Ok(object)
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

fn canonical_empty_comment_list(path: Path) -> Result<Vec<u8>, Error> {
    let mut output = Vec::new();
    append_varint_field(
        &mut output,
        1,
        u64::try_from(COMMENT_LIST_TYPE).map_err(|_| Error::InvalidSource { path })?,
    )
    .map_err(|_| Error::InvalidSource { path })?;
    append_varint_field(&mut output, 2, 1).map_err(|_| Error::InvalidSource { path })?;
    Ok(output)
}

fn comment_list_object(
    identifier: u64,
    data: Vec<u8>,
    key: u32,
    storage_identifier: u64,
    path: Path,
) -> Result<ArchiveObject, Error> {
    let mut object = ArchiveObject::new(identifier, vec![RawMessage { type_: 6_005, data }])
        .map_err(|_| Error::InvalidSource { path })?;
    let info = object
        .archive_info
        .message_infos
        .first_mut()
        .ok_or(Error::InvalidSource { path })?;
    info.object_references = vec![storage_identifier];
    let mut field = FieldInfo::new(vec![3, key]);
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references = vec![storage_identifier];
    info.field_infos.push(field);
    Ok(object)
}

fn attach_comment_list_to_model(
    archive: &mut Archive,
    source: &Package,
    route: MessageRoute,
    identifier: u64,
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<(), Error> {
    let message = message_at_route(source, route, path)?;
    let options = table_cell_decode_options(
        source,
        message.data.len().saturating_add(64),
        source.state.options.semantic().max_references(),
    );
    let prepared =
        numbers_table_cell_storage_codec::prepare_table_model_comment_storage_table_rewrite(
            message.data.as_slice(),
            numbers_table_cell_storage_codec::CommentStorageTableReferenceEdit::insert(identifier),
            options,
        )
        .map_err(|error| super::map_table_codec_error(error, path))?;
    let requirements = prepared.requirements();
    let data = prepared
        .execute(requirements.exact_limits())
        .map_err(|error| super::map_table_codec_error(error, path))?
        .0;
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
    if info.object_references.contains(&identifier) {
        return Err(Error::InvalidSource { path });
    }
    let before = info.object_references.clone();
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
    let mut field = FieldInfo::new(vec![4, 19]);
    field.r#type = Some(FieldType::ObjectReference);
    field.object_references = vec![identifier];
    info.field_infos.push(field);
    Ok(())
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

fn replace_tile(
    archive: &mut Archive,
    source: &Package,
    route: MessageRoute,
    located: &Located,
    key: u32,
    limits: litchi_iwa_core::Limits,
    path: Path,
) -> Result<Vec<u8>, Error> {
    let tile_source = message_at_route(source, route, path)?.data.as_slice();
    let change = tile::TileChange {
        row: u32::try_from(located.target.row % located.target.tile_size)
            .map_err(|_| Error::InvalidSource { path })?,
        column: u32::try_from(located.target.column).map_err(|_| Error::InvalidSource { path })?,
        change: tile::BncChange::CommentSet { identifier: key },
    };
    let bytes = tile_source.len().max(1);
    let prepared = tile::prepare_tile(tile::TileRewriteRequest {
        source: tile_source,
        columns: located.target.native.columns,
        changes: &[change],
        limits: tile::TileLimits::new(
            bytes,
            WireLimits::MAX_OUTPUT_BYTES,
            WireLimits::MAX_FIELDS,
            u64::try_from(WireLimits::MAX_REWRITE_WORK).unwrap_or(u64::MAX),
            located.target.tile_size,
            located.target.tile_size.saturating_mul(
                usize::try_from(located.target.native.columns)
                    .map_err(|_| Error::InvalidSource { path })?,
            ),
        ),
    })
    .map_err(|error| map_tile_error(error, path))?;
    let requirements = prepared.execution_requirements();
    let outcome = prepared
        .execute(requirements.exact_limits())
        .map_err(|error| map_tile_error(error, path))?;
    if outcome.transitions.len() != 1
        || outcome.transitions[0].after_references.comment != Some(key)
    {
        return Err(Error::Verification);
    }
    let rewritten = outcome.payload.ok_or(Error::Verification)?;
    let target_cell = extract_tile_cell(rewritten.as_slice(), change.row, change.column, path)?;
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
    Ok(target_cell)
}

fn extract_tile_cell(source: &[u8], row: u32, column: u32, path: Path) -> Result<Vec<u8>, Error> {
    let view = litchi_iwa_common::wire::WireView::parse(source)
        .map_err(|_| Error::InvalidSource { path })?;
    for field in view.fields().filter(|field| field.number() == 5) {
        let row_view = litchi_iwa_common::wire::WireView::parse(field.payload())
            .map_err(|_| Error::InvalidSource { path })?;
        let row_index = row_view
            .fields()
            .find(|field| field.number() == 1)
            .and_then(|field| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
            .and_then(|(value, _)| u32::try_from(value).ok());
        if row_index != Some(row) {
            continue;
        }
        let storage = row_view
            .fields()
            .find(|field| field.number() == 6)
            .ok_or(Error::InvalidSource { path })?
            .payload();
        let offsets = row_view
            .fields()
            .find(|field| field.number() == 7)
            .ok_or(Error::InvalidSource { path })?
            .payload();
        let wide = row_view
            .fields()
            .find(|field| field.number() == 8)
            .map(|field| {
                litchi_iwa_common::decode_varint_from_bytes(field.payload())
                    .map(|(value, _)| value != 0)
                    .map_err(|_| Error::InvalidSource { path })
            })
            .transpose()?
            .unwrap_or(false);
        let ranges = super::cell_ranges(
            offsets,
            storage.len(),
            row_view
                .fields()
                .find(|field| field.number() == 2)
                .and_then(|field| litchi_iwa_common::decode_varint_from_bytes(field.payload()).ok())
                .and_then(|(value, _)| usize::try_from(value).ok())
                .ok_or(Error::InvalidSource { path })?,
            wide,
            offsets.len() / 2,
            path,
            &mut 0,
        )?;
        let range = ranges
            .get(usize::try_from(column).map_err(|_| Error::InvalidSource { path })?)
            .and_then(Clone::clone)
            .ok_or(Error::Verification)?;
        return Ok(storage[range].to_vec());
    }
    Err(Error::InvalidSource { path })
}

fn map_tile_error(error: tile::TileError, path: Path) -> Error {
    match error {
        tile::TileError::NeedSparse { .. }
        | tile::TileError::UnsupportedSource { .. }
        | tile::TileError::UnsupportedValue { .. } => Error::UnsupportedDependency { path },
        tile::TileError::LimitExceeded { observed, maximum } => Error::LimitExceeded {
            kind: super::LimitKind::WireWork,
            observed: usize::try_from(observed).unwrap_or(usize::MAX),
            maximum: usize::try_from(maximum).unwrap_or(usize::MAX),
            path,
        },
        tile::TileError::Allocation { amount } => Error::Allocation { amount, path },
        tile::TileError::InvalidSource
        | tile::TileError::DuplicateOrUnsortedChange { .. }
        | tile::TileError::OutOfBounds { .. } => Error::InvalidSource { path },
    }
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
