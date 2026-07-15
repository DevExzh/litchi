//! Minimal private object-graph selection for a live slide created from a layout.

use std::collections::VecDeque;

use super::*;

const SLIDE_MESSAGE_TYPE: u32 = 5;
const NOTE_MESSAGE_TYPE: u32 = 15;

pub(super) struct NoteSource {
    pub(super) archive_name: String,
    pub(super) note_id: u64,
    pub(super) storage_id: u64,
    pub(super) object_ids: Vec<u64>,
}

pub(in crate::keynote::editor) fn take_identifier(next: &mut u64) -> Result<u64> {
    let identifier = *next;
    *next = next
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
    Ok(identifier)
}

/// Select private objects reachable from one or more roots inside one component.
pub(in crate::keynote::editor) fn private_clone_object_ids(
    archive: &Archive,
    roots: impl IntoIterator<Item = u64>,
    context: &str,
) -> Result<Vec<u64>> {
    let internal = archive
        .objects
        .iter()
        .map(|object| {
            object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat(format!("Keynote {context} object has no identifier"))
            })
        })
        .collect::<Result<HashSet<_>>>()?;
    let mut selected = HashSet::new();
    let mut pending = roots.into_iter().collect::<VecDeque<_>>();
    while let Some(identifier) = pending.pop_front() {
        if !internal.contains(&identifier) {
            return Err(Error::InvalidFormat(format!(
                "Keynote {context} object {identifier} is outside its component"
            )));
        }
        if !selected.insert(identifier) {
            continue;
        }
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote {context} object {identifier} is missing"))
        })?;
        for reference in object.archive_info.message_infos.iter().flat_map(|info| {
            info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| &field.object_references),
            )
        }) {
            if internal.contains(reference) && !selected.contains(reference) {
                pending.push_back(*reference);
            }
        }
    }
    Ok(archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .filter(|identifier| selected.contains(identifier))
        .collect())
}

#[allow(deprecated)]
pub(super) fn template_clone_object_ids(
    archive: &Archive,
    template_slide_id: u64,
    template: &kn::SlideArchive,
) -> Result<Vec<u64>> {
    if !template.builds.is_empty()
        || !template.build_chunks.is_empty()
        || !template.build_chunk_archives.is_empty()
        || template.note.is_some()
    {
        return Err(Error::InvalidFormat(format!(
            "Keynote layout slide {template_slide_id} contains live-only build or note state"
        )));
    }
    let internal = archive
        .objects
        .iter()
        .map(|object| {
            object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat("Keynote layout object has no identifier".to_owned())
            })
        })
        .collect::<Result<HashSet<_>>>()?;
    if !internal.contains(&template_slide_id) {
        return Err(Error::InvalidFormat(format!(
            "Keynote layout component does not contain slide {template_slide_id}"
        )));
    }
    let owned = template
        .owned_drawables
        .iter()
        .map(|reference| reference.identifier)
        .collect::<HashSet<_>>();
    for (label, reference) in [
        ("title", template.title_placeholder.as_ref()),
        ("body", template.body_placeholder.as_ref()),
    ] {
        if let Some(reference) = reference
            && !owned.contains(&reference.identifier)
        {
            return Err(Error::InvalidFormat(format!(
                "Keynote layout {label} placeholder {} is not an owned drawable",
                reference.identifier
            )));
        }
    }

    let mut selected = HashSet::from([template_slide_id]);
    let mut pending = template
        .owned_drawables
        .iter()
        .chain(template.slide_number_placeholder.iter())
        .chain(template.user_defined_guide_storage.iter())
        .map(|reference| reference.identifier)
        .collect::<VecDeque<_>>();
    while let Some(identifier) = pending.pop_front() {
        if !internal.contains(&identifier) {
            return Err(Error::InvalidFormat(format!(
                "Keynote layout private object {identifier} is outside its component"
            )));
        }
        if !selected.insert(identifier) {
            continue;
        }
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote layout object {identifier} is missing"))
        })?;
        for storage in object
            .messages
            .iter()
            .filter_map(|message| match message.type_ {
                7 => kn::PlaceholderArchive::decode(message.data.as_slice())
                    .ok()
                    .and_then(|placeholder| placeholder.super_.owned_storage),
                2_011 => tswp::ShapeInfoArchive::decode(message.data.as_slice())
                    .ok()
                    .and_then(|shape| shape.owned_storage),
                _ => None,
            })
        {
            if internal.contains(&storage.identifier) && !selected.contains(&storage.identifier) {
                pending.push_back(storage.identifier);
            }
        }
        for reference in object.archive_info.message_infos.iter().flat_map(|info| {
            info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| &field.object_references),
            )
        }) {
            if internal.contains(reference) && !selected.contains(reference) {
                pending.push_back(*reference);
            }
        }
    }

    Ok(archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .filter(|identifier| selected.contains(identifier))
        .collect())
}

pub(super) fn note_clone_object_ids(
    archive: &Archive,
    note_id: u64,
    storage_id: u64,
) -> Result<Vec<u64>> {
    let internal = archive
        .objects
        .iter()
        .map(|object| {
            object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat("Keynote note object has no identifier".to_owned())
            })
        })
        .collect::<Result<HashSet<_>>>()?;
    let mut selected = HashSet::new();
    let mut pending = VecDeque::from([note_id, storage_id]);
    while let Some(identifier) = pending.pop_front() {
        if !internal.contains(&identifier) {
            return Err(Error::InvalidFormat(format!(
                "Keynote note-private object {identifier} is outside its component"
            )));
        }
        if !selected.insert(identifier) {
            continue;
        }
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Keynote note object {identifier} is missing"))
        })?;
        for reference in object.archive_info.message_infos.iter().flat_map(|info| {
            info.object_references.iter().chain(
                info.field_infos
                    .iter()
                    .flat_map(|field| &field.object_references),
            )
        }) {
            if internal.contains(reference) && !selected.contains(reference) {
                pending.push_back(*reference);
            }
        }
    }
    Ok(archive
        .objects
        .iter()
        .filter_map(|object| object.archive_info.identifier)
        .filter(|identifier| selected.contains(identifier))
        .collect())
}

pub(super) fn find_note_source(
    editor: &KeynoteEditor,
    graph: &ObjectGraph,
    slides: &[KeynoteSlideInfo],
) -> Result<NoteSource> {
    let mut best = None;
    for slide in slides {
        let archive_name = graph.archive_name(slide.slide_id)?.to_owned();
        let archive = editor.package().archive(&archive_name)?;
        let decoded: kn::SlideArchive =
            graph.decode_type(slide.slide_id, SLIDE_MESSAGE_TYPE, "KN.SlideArchive")?;
        let Some(note_id) = decoded.note.map(|reference| reference.identifier) else {
            continue;
        };
        let note: kn::NoteArchive =
            graph.decode_type(note_id, NOTE_MESSAGE_TYPE, "KN.NoteArchive")?;
        let storage_id = note.contained_storage.identifier;
        if graph.archive_name(note_id)? != archive_name
            || graph.archive_name(storage_id)? != archive_name
        {
            continue;
        }
        let object_ids = note_clone_object_ids(&archive, note_id, storage_id)?;
        if best
            .as_ref()
            .is_none_or(|source: &NoteSource| object_ids.len() < source.object_ids.len())
        {
            best = Some(NoteSource {
                archive_name,
                note_id,
                storage_id,
                object_ids,
            });
        }
    }
    best.ok_or_else(|| {
        Error::InvalidFormat("Keynote presentation has no reusable speaker-notes graph".to_owned())
    })
}
