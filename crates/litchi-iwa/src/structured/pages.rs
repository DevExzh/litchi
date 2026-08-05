//! Pages-specific structured extraction.

use crate::Result;
use crate::bundle::Bundle;
use crate::object_index::ObjectIndex;
use litchi_pages::{Section, SectionType};
use prost::Message;

/// Extract the main body section while keeping Pages wire details private.
pub(super) fn extract(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Section>> {
    let Some(document_object) = bundle
        .get_archive("Index/Document.iwa")
        .and_then(|archive| archive.object(1))
    else {
        return Ok(Vec::new());
    };
    let Some(document) = document_object
        .messages
        .iter()
        .find(|message| message.type_ == 10000)
        .and_then(|message| {
            crate::protobuf::tp::DocumentArchive::decode(message.data.as_slice()).ok()
        })
    else {
        return Ok(Vec::new());
    };

    let mut section = Section::new(0, SectionType::Body);
    if let Some(reference) = document.body_storage {
        let object = object_index
            .resolve_ref_id(bundle, reference.identifier)?
            .ok_or_else(|| {
                crate::Error::InvalidFormat(format!(
                    "Pages body storage object {} is missing",
                    reference.identifier
                ))
            })?;
        let storage = object
            .messages
            .iter()
            .filter(|message| message.type_ == 2001 || message.type_ == 2022)
            .find_map(|message| {
                crate::protobuf::tswp::StorageArchive::decode(message.data.as_slice()).ok()
            })
            .ok_or_else(|| {
                crate::Error::InvalidFormat(format!(
                    "Pages body object {} has no text storage payload",
                    reference.identifier
                ))
            })?;
        let text = storage.text.concat();
        if !text.is_empty() {
            section.paragraphs.push(text);
        }
    }

    Ok(vec![section])
}
