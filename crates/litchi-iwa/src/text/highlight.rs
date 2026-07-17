//! Transactional native plain-highlight CRUD for TSWP text storages.

use std::collections::{HashMap, HashSet};

use crate::comments::{ensure_annotation_author, remove_generated_annotation_author_if_unused};
use crate::package_metadata::{
    add_component_external_reference, component_identifier_for_entry,
    component_identifier_for_object_uuid, next_object_identifier,
    release_package_identifier_suffix, remove_component_external_references_to_object,
    remove_component_object_uuids, set_package_last_object_identifier,
};
use crate::{Error, IWorkPackage, Result};

use super::highlight_object::{
    ensure_no_metadata_reference, insert_plain_comment_storage, new_highlight_object,
    require_exclusive_reference, validate_highlight_object, validate_plain_highlight_graph,
};
use super::highlight_storage::{
    Boundary, add_range, decoded_boundaries, encode_table, ensure_range_available, locate_storage,
    patch_highlight_table, raw_boundaries, remove_range, validate_range,
};
use super::highlight_types::{TextHighlight, TextHighlightId};
use super::position::{TextPosition, TextRange};
use super::storage_wire::{StorageLocation, text_utf16_len};

const OBJECT_IDENTIFIER_INCREMENT: u64 = 1;

/// Read every native plain highlight in a storage, ordered by text position.
pub(crate) fn text_highlights(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<Vec<TextHighlight>> {
    let location = locate_storage(package, storage_id)?;
    let boundaries = decoded_boundaries(storage_id, &location)?;
    collect_highlights(package, storage_id, &location, &boundaries)
}

/// Create a plain highlight over a currently unoccupied highlight range.
pub(crate) fn add_text_highlight(
    package: &mut IWorkPackage,
    storage_id: u64,
    range: TextRange,
) -> Result<TextHighlight> {
    let location = locate_storage(package, storage_id)?;
    ensure_no_overlapping_highlight_table(storage_id, &location)?;
    validate_range(storage_id, range, &location.storage.text)?;
    let boundaries = decoded_boundaries(storage_id, &location)?;
    ensure_range_available(storage_id, range, &boundaries, None, &location.storage.text)?;

    let mut staged = package.clone();
    let (author_id, author_entry, _) = ensure_annotation_author(&mut staged)?;
    let highlight_id = next_object_identifier(&staged)?;
    let comment_storage_id = highlight_id
        .checked_add(OBJECT_IDENTIFIER_INCREMENT)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
    let id = TextHighlightId::from_native(highlight_id);
    patch_highlight_table(
        &mut staged,
        &location.archive_name,
        storage_id,
        |table, storage| {
            let mut boundaries = raw_boundaries(storage_id, table, storage)?;
            ensure_range_available(storage_id, range, &boundaries, None, &storage.text)?;
            add_range(&mut boundaries, range, highlight_id)?;
            encode_table(table, boundaries).map(|table| (Some(table), Some(highlight_id), None))
        },
    )?;
    staged.update_archive(&location.archive_name, |archive| {
        archive.insert_object(new_highlight_object(highlight_id, comment_storage_id)?)
    })?;
    insert_plain_comment_storage(
        &mut staged,
        &location.archive_name,
        comment_storage_id,
        author_id,
    )?;
    add_author_external_reference(
        &mut staged,
        &location.archive_name,
        author_entry.as_deref(),
        author_id,
    )?;
    set_package_last_object_identifier(&mut staged, comment_storage_id)?;
    let verified = roundtrip(&staged)?;
    let created = highlight_by_id(&verified, storage_id, id)?;
    if created.range != range {
        return Err(Error::InvalidFormat(
            "iWork text-highlight creation failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(created)
}

/// Atomically move a plain highlight without changing its ID or owned metadata.
pub(crate) fn update_text_highlight(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextHighlightId,
    range: TextRange,
) -> Result<TextHighlight> {
    let current = highlight_by_id(package, storage_id, id)?;
    if current.range == range {
        return Ok(current);
    }
    let location = locate_storage(package, storage_id)?;
    ensure_no_overlapping_highlight_table(storage_id, &location)?;
    validate_range(storage_id, range, &location.storage.text)?;
    let boundaries = decoded_boundaries(storage_id, &location)?;
    ensure_range_available(
        storage_id,
        range,
        &boundaries,
        Some(id.object_id()),
        &location.storage.text,
    )?;
    require_exclusive_reference(package, storage_id, id.object_id(), "highlight")?;

    let mut staged = package.clone();
    patch_highlight_table(
        &mut staged,
        &location.archive_name,
        storage_id,
        |table, storage| {
            let mut boundaries = raw_boundaries(storage_id, table, storage)?;
            remove_range(&mut boundaries, id.object_id())?;
            ensure_range_available(storage_id, range, &boundaries, None, &storage.text)?;
            add_range(&mut boundaries, range, id.object_id())?;
            encode_table(table, boundaries).map(|table| (Some(table), None, None))
        },
    )?;
    let verified = roundtrip(&staged)?;
    let updated = highlight_by_id(&verified, storage_id, id)?;
    if updated.range != range {
        return Err(Error::InvalidFormat(
            "iWork text-highlight update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(updated)
}

/// Delete one plain highlight and its owned empty comment-storage graph.
pub(crate) fn remove_text_highlight(
    package: &mut IWorkPackage,
    storage_id: u64,
    id: TextHighlightId,
) -> Result<TextHighlight> {
    let removed = highlight_by_id(package, storage_id, id)?;
    let location = locate_storage(package, storage_id)?;
    require_exclusive_reference(package, storage_id, id.object_id(), "highlight")?;

    let mut staged = package.clone();
    patch_highlight_table(
        &mut staged,
        &location.archive_name,
        storage_id,
        |table, storage| {
            let mut boundaries = raw_boundaries(storage_id, table, storage)?;
            remove_range(&mut boundaries, id.object_id())?;
            if boundaries
                .iter()
                .any(|boundary| boundary.object_id.is_some())
            {
                encode_table(table, boundaries)
                    .map(|table| (Some(table), None, Some(id.object_id())))
            } else {
                Ok((None, None, Some(id.object_id())))
            }
        },
    )?;
    remove_detached_plain_highlights(&mut staged, &location.archive_name, &[id.object_id()])?;
    let verified = roundtrip(&staged)?;
    if text_highlights(&verified, storage_id)?
        .iter()
        .any(|highlight| highlight.id == id)
    {
        return Err(Error::InvalidFormat(
            "iWork text-highlight deletion failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(removed)
}

/// Reclaim plain-highlight graphs whose table references disappeared during text replacement.
pub(crate) fn remove_unreferenced_highlight_objects(
    package: &mut IWorkPackage,
    archive_name: &str,
    candidates: &HashSet<u64>,
) -> Result<()> {
    if candidates.is_empty() {
        return Ok(());
    }
    let archive = package.archive(archive_name)?;
    let mut identifiers = candidates
        .iter()
        .filter_map(|identifier| {
            archive
                .object(*identifier)
                .map(|object| (*identifier, object))
        })
        .filter_map(
            |(identifier, object)| match validate_highlight_object(identifier, object) {
                Ok(Some(_)) => Some(Ok(identifier)),
                Ok(None) => None,
                Err(error) => Some(Err(error)),
            },
        )
        .collect::<Result<Vec<_>>>()?;
    if identifiers.is_empty() {
        return Ok(());
    }
    identifiers.sort_unstable();
    let mut staged = package.clone();
    remove_detached_plain_highlights(&mut staged, archive_name, &identifiers)?;
    roundtrip(&staged)?;
    *package = staged;
    Ok(())
}

fn ensure_no_overlapping_highlight_table(
    storage_id: u64,
    location: &StorageLocation,
) -> Result<()> {
    if location.storage.table_overlapping_highlight.is_some() {
        return Err(Error::InvalidFormat(format!(
            "text storage {storage_id} has overlapping highlights that cannot yet be rewritten safely"
        )));
    }
    Ok(())
}

fn add_author_external_reference(
    package: &mut IWorkPackage,
    source_entry: &str,
    author_entry: Option<&str>,
    author_id: Option<u64>,
) -> Result<()> {
    let (Some(author_entry), Some(author_id)) = (author_entry, author_id) else {
        return Ok(());
    };
    let source_component = component_identifier_for_entry(package, source_entry)?;
    let author_component = component_identifier_for_entry(package, author_entry)?;
    if let (Some(source_component), Some(author_component)) = (source_component, author_component)
        && source_component != author_component
    {
        add_component_external_reference(package, source_component, author_component, author_id)?;
    }
    Ok(())
}

fn remove_detached_plain_highlights(
    package: &mut IWorkPackage,
    archive_name: &str,
    identifiers: &[u64],
) -> Result<()> {
    let archive = package.archive(archive_name)?;
    let mut graphs = HashMap::new();
    for identifier in identifiers {
        let object = archive.object(*identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork highlight object {identifier} is missing"))
        })?;
        let graph = validate_plain_highlight_graph(package, archive_name, *identifier, object)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!("object {identifier} is not a text highlight"))
            })?;
        ensure_no_metadata_reference(package, *identifier)?;
        require_exclusive_reference(
            package,
            *identifier,
            graph.comment_storage_id,
            "highlight comment-storage",
        )?;
        graphs.insert(*identifier, graph);
    }

    let owning_component = component_identifier_for_entry(package, archive_name)?;
    let mut removed_ids = Vec::with_capacity(identifiers.len() * 2);
    let mut author_ids = HashSet::new();
    for identifier in identifiers {
        let graph = graphs[identifier];
        remove_registered_object(package, owning_component, *identifier)?;
        package.update_archive(archive_name, |archive| {
            archive.remove_object(*identifier).ok_or_else(|| {
                Error::InvalidFormat(format!("iWork highlight object {identifier} is missing"))
            })?;
            Ok(())
        })?;
        ensure_no_metadata_reference(package, graph.comment_storage_id)?;
        remove_registered_object(package, owning_component, graph.comment_storage_id)?;
        package.update_archive(archive_name, |archive| {
            archive
                .remove_object(graph.comment_storage_id)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "iWork highlight comment storage {} is missing",
                        graph.comment_storage_id
                    ))
                })?;
            Ok(())
        })?;
        removed_ids.extend([*identifier, graph.comment_storage_id]);
        author_ids.extend(graph.author_id);
    }
    for author_id in author_ids {
        if remove_generated_annotation_author_if_unused(package, author_id)? {
            removed_ids.push(author_id);
        }
    }
    removed_ids.sort_unstable_by(|left, right| right.cmp(left));
    release_package_identifier_suffix(package, &removed_ids)
}

fn remove_registered_object(
    package: &mut IWorkPackage,
    owning_component: Option<u64>,
    identifier: u64,
) -> Result<()> {
    if let Some(component) = owning_component {
        remove_component_external_references_to_object(package, component, identifier)?;
    }
    if let Some(component) = component_identifier_for_object_uuid(package, identifier)? {
        remove_component_object_uuids(package, component, &[identifier])?;
    }
    Ok(())
}

fn highlight_by_id(
    package: &IWorkPackage,
    storage_id: u64,
    id: TextHighlightId,
) -> Result<TextHighlight> {
    let matches = text_highlights(package, storage_id)?
        .into_iter()
        .filter(|highlight| highlight.id == id)
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [highlight] => Ok(*highlight),
        [] => Err(Error::InvalidFormat(format!(
            "text storage {storage_id} does not own highlight object {}",
            id.object_id()
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "text storage {storage_id} references highlight object {} more than once",
            id.object_id()
        ))),
    }
}

fn collect_highlights(
    package: &IWorkPackage,
    storage_id: u64,
    location: &StorageLocation,
    boundaries: &[Boundary],
) -> Result<Vec<TextHighlight>> {
    let text_len = text_utf16_len(&location.storage.text)?;
    let archive = package.archive(&location.archive_name)?;
    let mut seen = HashSet::new();
    let mut highlights = Vec::new();
    for (position, boundary) in boundaries.iter().enumerate() {
        let Some(identifier) = boundary.object_id else {
            continue;
        };
        let end = boundaries
            .get(position + 1)
            .map_or(text_len, |next| next.index);
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} references missing highlight object {identifier}"
            ))
        })?;
        let Some(_) =
            validate_plain_highlight_graph(package, &location.archive_name, identifier, object)?
        else {
            continue;
        };
        if !seen.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} references highlight object {identifier} more than once"
            )));
        }
        highlights.push(TextHighlight::new(
            TextHighlightId::from_native(identifier),
            TextRange::new(
                TextPosition::from_native(boundary.index),
                TextPosition::from_native(end),
            )?,
        ));
    }
    Ok(highlights)
}

fn roundtrip(package: &IWorkPackage) -> Result<IWorkPackage> {
    IWorkPackage::from_bytes(&package.to_bytes()?)
}

#[cfg(test)]
#[path = "highlight_internal_tests.rs"]
mod tests;
