//! Transactional native ranged-annotation graph CRUD for TSWP text storages.

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
    AnnotationGraph, ensure_no_metadata_reference, insert_annotation_comment_storage,
    new_highlight_object, require_exclusive_reference, update_annotation_comment_text,
    validate_annotation_graph, validate_highlight_object,
};
use super::highlight_storage::{
    Boundary, add_range, decoded_boundaries, encode_table, ensure_range_available, locate_storage,
    locate_storage_with_archive, patch_highlight_table, raw_boundaries, remove_range,
    validate_range,
};
use super::position::{TextPosition, TextRange};
use super::storage_wire::{StorageLocation, text_utf16_len};

const OBJECT_IDENTIFIER_INCREMENT: u64 = 1;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum AnnotationKind {
    PlainHighlight,
    Comment,
}

impl AnnotationKind {
    fn matches(self, graph: &AnnotationGraph) -> bool {
        graph.is_plain_highlight() == matches!(self, Self::PlainHighlight)
    }

    fn label(self) -> &'static str {
        match self {
            Self::PlainHighlight => "plain highlight",
            Self::Comment => "text comment",
        }
    }
}

#[derive(Debug, Clone)]
pub(super) struct AnnotationRecord {
    pub(super) object_id: u64,
    pub(super) range: TextRange,
    pub(super) graph: AnnotationGraph,
}

pub(super) fn annotations(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<Vec<AnnotationRecord>> {
    let location = locate_storage(package, storage_id)?;
    let boundaries = decoded_boundaries(storage_id, &location)?;
    collect_annotations(package, storage_id, &location, &boundaries)
}

pub(super) fn annotation_by_id(
    package: &IWorkPackage,
    storage_id: u64,
    object_id: u64,
    kind: AnnotationKind,
) -> Result<AnnotationRecord> {
    let mut matches = annotations(package, storage_id)?
        .into_iter()
        .filter(|annotation| annotation.object_id == object_id && kind.matches(&annotation.graph));
    let Some(annotation) = matches.next() else {
        return Err(Error::InvalidFormat(format!(
            "text storage {storage_id} does not own {} object {object_id}",
            kind.label()
        )));
    };
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "text storage {storage_id} references {} object {object_id} more than once",
            kind.label()
        )));
    }
    Ok(annotation)
}

pub(super) fn add_annotation(
    package: &mut IWorkPackage,
    storage_id: u64,
    range: TextRange,
    kind: AnnotationKind,
    body: String,
) -> Result<AnnotationRecord> {
    validate_kind_body(kind, &body)?;
    let mut staged = package.clone();
    let (author_id, author_entry, _) = ensure_annotation_author(&mut staged)?;
    let annotation_id = next_object_identifier(&staged)?;
    let comment_storage_id = annotation_id
        .checked_add(OBJECT_IDENTIFIER_INCREMENT)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
    let located = locate_storage_with_archive(&staged, storage_id)?;
    let location = &located.location;
    ensure_no_overlapping_highlight_table(storage_id, location)?;
    validate_range(storage_id, range, &location.storage.text)?;
    let boundaries = decoded_boundaries(storage_id, location)?;
    ensure_range_available(storage_id, range, &boundaries, None, &location.storage.text)?;
    let archive_name = location.archive_name.clone();
    patch_highlight_table(&mut staged, located, |table, storage| {
        let mut boundaries = raw_boundaries(storage_id, table, storage)?;
        ensure_range_available(storage_id, range, &boundaries, None, &storage.text)?;
        add_range(&mut boundaries, range, annotation_id)?;
        encode_table(table, boundaries).map(|table| (Some(table), Some(annotation_id), None))
    })?;
    staged.update_archive(&archive_name, |archive| {
        archive.insert_object(new_highlight_object(annotation_id, comment_storage_id)?)
    })?;
    insert_annotation_comment_storage(
        &mut staged,
        &archive_name,
        comment_storage_id,
        body,
        author_id,
    )?;
    add_author_external_reference(
        &mut staged,
        &archive_name,
        author_entry.as_deref(),
        author_id,
    )?;
    set_package_last_object_identifier(&mut staged, comment_storage_id)?;
    let verified = roundtrip(&staged)?;
    let created = annotation_by_id(&verified, storage_id, annotation_id, kind)?;
    if created.range != range {
        return Err(Error::InvalidFormat(format!(
            "iWork {} creation failed range validation",
            kind.label()
        )));
    }
    *package = staged;
    Ok(created)
}

pub(super) fn update_annotation(
    package: &mut IWorkPackage,
    storage_id: u64,
    object_id: u64,
    range: TextRange,
    kind: AnnotationKind,
    body: &str,
) -> Result<AnnotationRecord> {
    validate_kind_body(kind, body)?;
    let current = annotation_by_id(package, storage_id, object_id, kind)?;
    if current.range == range && current.graph.body == body {
        return Ok(current);
    }
    let located = locate_storage_with_archive(package, storage_id)?;
    let location = &located.location;
    ensure_no_overlapping_highlight_table(storage_id, location)?;
    validate_range(storage_id, range, &location.storage.text)?;
    let boundaries = decoded_boundaries(storage_id, location)?;
    ensure_range_available(
        storage_id,
        range,
        &boundaries,
        Some(object_id),
        &location.storage.text,
    )?;
    require_exclusive_reference(package, storage_id, object_id, kind.label())?;
    require_exclusive_reference(
        package,
        object_id,
        current.graph.comment_storage_id,
        "annotation comment-storage",
    )?;

    let archive_name = location.archive_name.clone();
    let mut staged = package.clone();
    if current.range != range {
        patch_highlight_table(&mut staged, located, |table, storage| {
            let mut boundaries = raw_boundaries(storage_id, table, storage)?;
            remove_range(&mut boundaries, object_id)?;
            ensure_range_available(storage_id, range, &boundaries, None, &storage.text)?;
            add_range(&mut boundaries, range, object_id)?;
            encode_table(table, boundaries).map(|table| (Some(table), None, None))
        })?;
    }
    if current.graph.body != body {
        update_annotation_comment_text(
            &mut staged,
            &archive_name,
            object_id,
            current.graph.comment_storage_id,
            body,
        )?;
    }
    let verified = roundtrip(&staged)?;
    let updated = annotation_by_id(&verified, storage_id, object_id, kind)?;
    if updated.range != range || updated.graph.body != body {
        return Err(Error::InvalidFormat(format!(
            "iWork {} update failed validation",
            kind.label()
        )));
    }
    *package = staged;
    Ok(updated)
}

pub(super) fn remove_annotation(
    package: &mut IWorkPackage,
    storage_id: u64,
    object_id: u64,
    kind: AnnotationKind,
) -> Result<AnnotationRecord> {
    let removed = annotation_by_id(package, storage_id, object_id, kind)?;
    let located = locate_storage_with_archive(package, storage_id)?;
    let location = &located.location;
    require_exclusive_reference(package, storage_id, object_id, kind.label())?;

    let archive_name = location.archive_name.clone();
    let mut staged = package.clone();
    patch_highlight_table(&mut staged, located, |table, storage| {
        let mut boundaries = raw_boundaries(storage_id, table, storage)?;
        remove_range(&mut boundaries, object_id)?;
        if boundaries
            .iter()
            .any(|boundary| boundary.object_id.is_some())
        {
            encode_table(table, boundaries).map(|table| (Some(table), None, Some(object_id)))
        } else {
            Ok((None, None, Some(object_id)))
        }
    })?;
    remove_detached_annotations(&mut staged, &archive_name, &[object_id], Some(kind))?;
    let verified = roundtrip(&staged)?;
    if annotations(&verified, storage_id)?
        .iter()
        .any(|annotation| annotation.object_id == object_id)
    {
        return Err(Error::InvalidFormat(format!(
            "iWork {} deletion failed validation",
            kind.label()
        )));
    }
    *package = staged;
    Ok(removed)
}

pub(super) fn remove_unreferenced_annotation_objects(
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
    remove_detached_annotations(&mut staged, archive_name, &identifiers, None)?;
    roundtrip(&staged)?;
    *package = staged;
    Ok(())
}

fn validate_kind_body(kind: AnnotationKind, body: &str) -> Result<()> {
    let valid = match kind {
        AnnotationKind::PlainHighlight => body.is_empty(),
        AnnotationKind::Comment => !body.is_empty(),
    };
    if valid {
        Ok(())
    } else {
        Err(Error::ParseError(format!(
            "{} body presence is invalid",
            kind.label()
        )))
    }
}

fn ensure_no_overlapping_highlight_table(
    storage_id: u64,
    location: &StorageLocation,
) -> Result<()> {
    if location.storage.table_overlapping_highlight.is_some() {
        return Err(Error::InvalidFormat(format!(
            "text storage {storage_id} has overlapping annotations that cannot yet be rewritten safely"
        )));
    }
    Ok(())
}

pub(super) fn add_author_external_reference(
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

fn remove_detached_annotations(
    package: &mut IWorkPackage,
    archive_name: &str,
    identifiers: &[u64],
    expected_kind: Option<AnnotationKind>,
) -> Result<()> {
    let archive = package.archive(archive_name)?;
    let mut graphs = HashMap::new();
    for identifier in identifiers {
        let object = archive.object(*identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork annotation object {identifier} is missing"))
        })?;
        let graph = validate_annotation_graph(package, archive_name, *identifier, object)?
            .ok_or_else(|| {
                Error::InvalidFormat(format!("object {identifier} is not a text annotation"))
            })?;
        if expected_kind.is_some_and(|kind| !kind.matches(&graph)) {
            return Err(Error::InvalidFormat(format!(
                "text annotation {identifier} has an unexpected kind"
            )));
        }
        ensure_no_metadata_reference(package, *identifier)?;
        require_exclusive_reference(
            package,
            *identifier,
            graph.comment_storage_id,
            "annotation comment-storage",
        )?;
        for reply in &graph.replies {
            require_exclusive_reference(
                package,
                graph.comment_storage_id,
                reply.storage_id,
                "annotation reply",
            )?;
        }
        graphs.insert(*identifier, graph);
    }

    let owning_component = component_identifier_for_entry(package, archive_name)?;
    let mut removed_ids = Vec::new();
    let mut author_ids = HashSet::new();
    for identifier in identifiers {
        let graph = &graphs[identifier];
        author_ids.extend(graph.author_ids());
        remove_registered_object(package, owning_component, *identifier)?;
        package.update_archive(archive_name, |archive| {
            archive.remove_object(*identifier).ok_or_else(|| {
                Error::InvalidFormat(format!("iWork annotation object {identifier} is missing"))
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
                        "iWork annotation comment storage {} is missing",
                        graph.comment_storage_id
                    ))
                })?;
            Ok(())
        })?;
        removed_ids.extend([*identifier, graph.comment_storage_id]);
        for reply in &graph.replies {
            ensure_no_metadata_reference(package, reply.storage_id)?;
            remove_registered_object(package, owning_component, reply.storage_id)?;
            package.update_archive(archive_name, |archive| {
                archive.remove_object(reply.storage_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "iWork annotation reply storage {} is missing",
                        reply.storage_id
                    ))
                })?;
                Ok(())
            })?;
            removed_ids.push(reply.storage_id);
        }
    }
    for author_id in author_ids {
        if remove_generated_annotation_author_if_unused(package, author_id)? {
            removed_ids.push(author_id);
        }
    }
    removed_ids.sort_unstable_by(|left, right| right.cmp(left));
    release_package_identifier_suffix(package, &removed_ids)
}

pub(super) fn remove_registered_object(
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

fn collect_annotations(
    package: &IWorkPackage,
    storage_id: u64,
    location: &StorageLocation,
    boundaries: &[Boundary],
) -> Result<Vec<AnnotationRecord>> {
    let text_len = text_utf16_len(&location.storage.text)?;
    let archive = package.archive(&location.archive_name)?;
    let mut seen = HashSet::new();
    let mut annotations = Vec::new();
    for (position, boundary) in boundaries.iter().enumerate() {
        let Some(identifier) = boundary.object_id else {
            continue;
        };
        let end = boundaries
            .get(position + 1)
            .map_or(text_len, |next| next.index);
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} references missing annotation object {identifier}"
            ))
        })?;
        let Some(graph) =
            validate_annotation_graph(package, &location.archive_name, identifier, object)?
        else {
            continue;
        };
        if !seen.insert(identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} references annotation object {identifier} more than once"
            )));
        }
        annotations.push(AnnotationRecord {
            object_id: identifier,
            range: TextRange::new(
                TextPosition::from_native(boundary.index),
                TextPosition::from_native(end),
            )?,
            graph,
        });
    }
    Ok(annotations)
}

pub(super) fn roundtrip(package: &IWorkPackage) -> Result<IWorkPackage> {
    IWorkPackage::from_bytes(&package.to_bytes()?)
}
