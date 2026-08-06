//! Keynote-specific structured extraction.

use crate::Result;
use crate::bundle::Bundle;
use crate::object_index::{ObjectIndex, ResolvedObjectRef};
use litchi_keynote::Slide;
use prost::Message;

const DOCUMENT_ARCHIVE_NAME: &str = "Index/Document.iwa";
const DOCUMENT_OBJECT_ID: u64 = 1;
const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHOW_MESSAGE_TYPE: u32 = 2;
const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const PLACEHOLDER_MESSAGE_TYPE: u32 = 7;
const NOTE_MESSAGE_TYPE: u32 = 15;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const DOCUMENT_MESSAGE_TYPES: &[u32] = &[DOCUMENT_MESSAGE_TYPE];
const SHOW_MESSAGE_TYPES: &[u32] = &[SHOW_MESSAGE_TYPE];
const SLIDE_NODE_MESSAGE_TYPES: &[u32] = &[SLIDE_NODE_MESSAGE_TYPE];
const SLIDE_MESSAGE_TYPES: &[u32] = &[SLIDE_MESSAGE_TYPE];
const PLACEHOLDER_MESSAGE_TYPES: &[u32] = &[PLACEHOLDER_MESSAGE_TYPE];
const NOTE_MESSAGE_TYPES: &[u32] = &[NOTE_MESSAGE_TYPE];
const SHAPE_INFO_MESSAGE_TYPES: &[u32] = &[SHAPE_INFO_MESSAGE_TYPE];

/// Extract slides while keeping native protobuf traversal private to IWA.
pub(super) fn extract(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Slide>> {
    let document_archive = bundle.get_archive(DOCUMENT_ARCHIVE_NAME).ok_or_else(|| {
        crate::Error::InvalidFormat("Keynote structured root archive is missing".to_owned())
    })?;
    let document_object = document_archive.object(DOCUMENT_OBJECT_ID).ok_or_else(|| {
        crate::Error::InvalidFormat("Keynote structured root object 1 is missing".to_owned())
    })?;
    let document_payload = unique_payload(
        document_object.messages.as_slice(),
        DOCUMENT_MESSAGE_TYPES,
        "Keynote structured root",
    )?;
    let document =
        crate::protobuf::kn::DocumentArchive::decode(document_payload).map_err(|error| {
            crate::Error::InvalidFormat(format!(
                "Keynote structured root payload is invalid: {error}"
            ))
        })?;

    if document.show.identifier == 0 {
        // A present, decodable default DocumentArchive is a legitimate empty
        // root. It has no show graph to traverse, unlike a non-zero reference
        // whose missing target is malformed.
        return Ok(Vec::new());
    }

    let show_object = required_object(
        bundle,
        object_index,
        document.show.identifier,
        "Keynote show",
    )?;
    let show_payload = unique_payload(show_object.messages, SHOW_MESSAGE_TYPES, "Keynote show")?;
    let show = crate::protobuf::kn::ShowArchive::decode(show_payload).map_err(|error| {
        crate::Error::InvalidFormat(format!("Keynote show payload is invalid: {error}"))
    })?;

    let mut slides = Vec::with_capacity(show.slide_tree.slides.len());

    for node_reference in show.slide_tree.slides {
        let node_object = required_object(
            bundle,
            object_index,
            node_reference.identifier,
            "Keynote slide-tree node",
        )?;
        let node_payload = unique_payload(
            node_object.messages,
            SLIDE_NODE_MESSAGE_TYPES,
            "Keynote slide-tree node",
        )?;
        let node =
            crate::protobuf::kn::SlideNodeArchive::decode(node_payload).map_err(|error| {
                crate::Error::InvalidFormat(format!(
                    "Keynote slide-tree node payload is invalid: {error}"
                ))
            })?;
        let Some(slide_reference) = node.slide else {
            return Err(crate::Error::InvalidFormat(
                "Keynote slide-tree node has no slide reference".to_owned(),
            ));
        };
        let slide_object = required_object(
            bundle,
            object_index,
            slide_reference.identifier,
            "Keynote slide",
        )?;
        let slide_payload =
            unique_payload(slide_object.messages, SLIDE_MESSAGE_TYPES, "Keynote slide")?;
        let archive =
            crate::protobuf::kn::SlideArchive::decode(slide_payload).map_err(|error| {
                crate::Error::InvalidFormat(format!("Keynote slide payload is invalid: {error}"))
            })?;

        let index = slides.len();
        let mut slide = Slide::builder(index);
        slide.set_title(archive.name.filter(|name| !name.is_empty()));
        let title_placeholder = archive
            .title_placeholder
            .as_ref()
            .map(|reference| reference.identifier);
        let body_placeholder = archive
            .body_placeholder
            .as_ref()
            .map(|reference| reference.identifier);

        if let Some(identifier) = title_placeholder
            && let Some(text) = drawable_text(
                bundle,
                object_index,
                identifier,
                "Keynote title placeholder",
                true,
            )?
        {
            slide.set_title(Some(text));
        }
        if let Some(identifier) = body_placeholder
            && let Some(text) = drawable_text(
                bundle,
                object_index,
                identifier,
                "Keynote body placeholder",
                true,
            )?
        {
            slide.push_text(text);
        }
        for drawable in archive.owned_drawables {
            if Some(drawable.identifier) == title_placeholder
                || Some(drawable.identifier) == body_placeholder
            {
                continue;
            }
            if let Some(text) = drawable_text(
                bundle,
                object_index,
                drawable.identifier,
                "Keynote slide drawable",
                false,
            )? {
                slide.push_text(text);
            }
        }
        if let Some(note) = archive.note {
            let note_object =
                required_object(bundle, object_index, note.identifier, "Keynote slide note")?;
            let note_payload = unique_payload(
                note_object.messages,
                NOTE_MESSAGE_TYPES,
                "Keynote slide note",
            )?;
            let note = crate::protobuf::kn::NoteArchive::decode(note_payload).map_err(|error| {
                crate::Error::InvalidFormat(format!(
                    "Keynote slide note payload is invalid: {error}"
                ))
            })?;
            let storage = required_object(
                bundle,
                object_index,
                note.contained_storage.identifier,
                "Keynote speaker-note storage",
            )?;
            let text = decode_storage(storage.messages, "Keynote speaker-note storage")?;
            if !text.is_empty() {
                slide.set_notes(Some(text));
            }
        }
        slides.push(slide.build());
    }

    Ok(slides)
}

fn required_object<'a>(
    bundle: &'a Bundle,
    object_index: &ObjectIndex,
    identifier: u64,
    context: &str,
) -> Result<ResolvedObjectRef<'a>> {
    object_index
        .resolve_ref_id(bundle, identifier)?
        .ok_or_else(|| {
            crate::Error::InvalidFormat(format!("{context} object {identifier} is missing"))
        })
}

fn drawable_text(
    bundle: &Bundle,
    object_index: &ObjectIndex,
    identifier: u64,
    context: &str,
    required: bool,
) -> Result<Option<String>> {
    let drawable = required_object(bundle, object_index, identifier, context)?;
    let has_placeholder = drawable
        .messages
        .iter()
        .any(|message| message.type_ == PLACEHOLDER_MESSAGE_TYPE);
    let has_shape_info = drawable
        .messages
        .iter()
        .any(|message| message.type_ == SHAPE_INFO_MESSAGE_TYPE);
    if has_placeholder && has_shape_info {
        return Err(crate::Error::InvalidFormat(format!(
            "{context} object {identifier} contains ambiguous text payloads"
        )));
    }

    let storage_id = if has_placeholder {
        let payload = unique_payload(drawable.messages, PLACEHOLDER_MESSAGE_TYPES, context)?;
        let placeholder =
            crate::protobuf::kn::PlaceholderArchive::decode(payload).map_err(|error| {
                crate::Error::InvalidFormat(format!(
                    "{context} object {identifier} payload is invalid: {error}"
                ))
            })?;
        placeholder
            .super_
            .owned_storage
            .map(|reference| reference.identifier)
    } else if has_shape_info {
        let payload = unique_payload(drawable.messages, SHAPE_INFO_MESSAGE_TYPES, context)?;
        let shape = crate::protobuf::tswp::ShapeInfoArchive::decode(payload).map_err(|error| {
            crate::Error::InvalidFormat(format!(
                "{context} object {identifier} payload is invalid: {error}"
            ))
        })?;
        shape.owned_storage.map(|reference| reference.identifier)
    } else {
        None
    };

    let Some(storage_id) = storage_id else {
        if required {
            return Err(crate::Error::InvalidFormat(format!(
                "{context} object {identifier} has no text-storage reference"
            )));
        }
        return Ok(None);
    };
    let storage_object = required_object(bundle, object_index, storage_id, context)?;
    let text = decode_storage(storage_object.messages, context)?;
    Ok((!text.is_empty()).then_some(text))
}

fn decode_storage(messages: &[crate::archive::RawMessage], context: &str) -> Result<String> {
    let payload = unique_payload(messages, STORAGE_MESSAGE_TYPES, context)?;
    let storage = crate::protobuf::tswp::StorageArchive::decode(payload).map_err(|error| {
        crate::Error::InvalidFormat(format!(
            "{context} text-storage payload is invalid: {error}"
        ))
    })?;
    Ok(storage.text.concat())
}

fn unique_payload<'a>(
    messages: &'a [crate::archive::RawMessage],
    message_types: &[u32],
    context: &str,
) -> Result<&'a [u8]> {
    let mut payload = None;
    for message in messages {
        if !message_types.contains(&message.type_) {
            continue;
        }
        if payload.is_some() {
            return Err(crate::Error::InvalidFormat(format!(
                "{context} contains duplicate required payloads"
            )));
        }
        payload = Some(message.data.as_slice());
    }
    payload.ok_or_else(|| {
        let expected = message_types
            .iter()
            .map(|message_type| format!("type-{message_type}"))
            .collect::<Vec<_>>()
            .join("/");
        crate::Error::InvalidFormat(format!("{context} has no {expected} payload"))
    })
}
