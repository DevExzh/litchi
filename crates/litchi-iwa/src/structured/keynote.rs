//! Keynote-specific structured extraction.

use crate::Result;
use crate::archive::ArchiveObject;
use crate::bundle::Bundle;
use crate::object_index::ObjectIndex;
use litchi_keynote::Slide;
use prost::Message;

/// Extract slides while keeping native protobuf traversal private to IWA.
pub(super) fn extract(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Slide>> {
    let mut slides = Vec::new();
    let Some(document_object) = bundle_object(bundle, 1) else {
        return Ok(slides);
    };
    let Some(document) = document_object.messages.iter().find_map(|message| {
        crate::protobuf::kn::DocumentArchive::decode(message.data.as_slice()).ok()
    }) else {
        return Ok(slides);
    };
    let Some(show_object) = bundle_object(bundle, document.show.identifier) else {
        return Ok(slides);
    };
    let Some(show) = show_object
        .messages
        .iter()
        .find_map(|message| crate::protobuf::kn::ShowArchive::decode(message.data.as_slice()).ok())
    else {
        return Ok(slides);
    };

    for node_reference in show.slide_tree.slides {
        let Some(node_object) = bundle_object(bundle, node_reference.identifier) else {
            continue;
        };
        let Some(node) = node_object.messages.iter().find_map(|message| {
            crate::protobuf::kn::SlideNodeArchive::decode(message.data.as_slice()).ok()
        }) else {
            continue;
        };
        let Some(slide_reference) = node.slide else {
            continue;
        };
        let Some(slide_object) = bundle_object(bundle, slide_reference.identifier) else {
            continue;
        };
        let Some(archive) = slide_object.messages.iter().find_map(|message| {
            crate::protobuf::kn::SlideArchive::decode(message.data.as_slice()).ok()
        }) else {
            continue;
        };

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
            && let Some(text) = drawable_text(bundle, object_index, identifier)?
        {
            slide.set_title(Some(text));
        }
        if let Some(identifier) = body_placeholder
            && let Some(text) = drawable_text(bundle, object_index, identifier)?
        {
            slide.push_text(text);
        }
        for drawable in archive.owned_drawables {
            if Some(drawable.identifier) == title_placeholder
                || Some(drawable.identifier) == body_placeholder
            {
                continue;
            }
            if let Some(text) = drawable_text(bundle, object_index, drawable.identifier)? {
                slide.push_text(text);
            }
        }
        if let Some(note) = archive.note
            && let Some(note_object) = object_index.resolve_ref_id(bundle, note.identifier)?
        {
            for message in note_object.messages {
                let Ok(note) = crate::protobuf::kn::NoteArchive::decode(message.data.as_slice())
                else {
                    continue;
                };
                if let Some(storage) = object_index
                    .resolve_ref_id(bundle, note.contained_storage.identifier)?
                    .and_then(|object| {
                        object.messages.iter().find_map(|message| {
                            crate::protobuf::tswp::StorageArchive::decode(message.data.as_slice())
                                .ok()
                        })
                    })
                {
                    let text = storage.text.concat();
                    if !text.is_empty() {
                        slide.set_notes(Some(text));
                    }
                }
                break;
            }
        }
        slides.push(slide.build());
    }

    Ok(slides)
}

fn bundle_object(bundle: &Bundle, identifier: u64) -> Option<&ArchiveObject> {
    bundle
        .iter_archives()
        .map(|(_, archive)| archive)
        .find_map(|archive| archive.object(identifier))
}

fn drawable_text(
    bundle: &Bundle,
    object_index: &ObjectIndex,
    identifier: u64,
) -> Result<Option<String>> {
    let Some(drawable) = object_index.resolve_ref_id(bundle, identifier)? else {
        return Ok(None);
    };
    let storage_id = drawable.messages.iter().find_map(|message| {
        crate::protobuf::kn::PlaceholderArchive::decode(message.data.as_slice())
            .ok()
            .and_then(|placeholder| placeholder.super_.owned_storage)
            .or_else(|| {
                crate::protobuf::tswp::ShapeInfoArchive::decode(message.data.as_slice())
                    .ok()
                    .and_then(|shape| shape.owned_storage)
            })
            .map(|reference| reference.identifier)
    });
    let Some(storage_id) = storage_id else {
        return Ok(None);
    };
    let Some(storage_object) = object_index.resolve_ref_id(bundle, storage_id)? else {
        return Ok(None);
    };
    for message in storage_object.messages {
        if let Ok(storage) = crate::protobuf::tswp::StorageArchive::decode(message.data.as_slice())
        {
            let text = storage.text.concat();
            return Ok((!text.is_empty()).then_some(text));
        }
    }
    Ok(None)
}
