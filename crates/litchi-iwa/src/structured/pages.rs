//! Pages-specific structured extraction.

use crate::Result;
use crate::bundle::Bundle;
use crate::object_index::ObjectIndex;
use litchi_iwa_text::storage::Storage;
use litchi_pages::{Section, SectionType};
use prost::Message;

const DOCUMENT_ARCHIVE_NAME: &str = "Index/Document.iwa";
const DOCUMENT_OBJECT_ID: u64 = 1;
const DOCUMENT_MESSAGE_TYPE: u32 = 10_000;
const DOCUMENT_MESSAGE_TYPES: &[u32] = &[DOCUMENT_MESSAGE_TYPE];
const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];

/// Extract the main body section while keeping Pages wire details private.
pub(super) fn extract(bundle: &Bundle, object_index: &ObjectIndex) -> Result<Vec<Section>> {
    let archive = bundle.get_archive(DOCUMENT_ARCHIVE_NAME).ok_or_else(|| {
        crate::Error::InvalidFormat("Pages structured root archive is missing".to_owned())
    })?;
    let document_object = archive.object(DOCUMENT_OBJECT_ID).ok_or_else(|| {
        crate::Error::InvalidFormat("Pages structured root object 1 is missing".to_owned())
    })?;
    let payload = unique_payload(
        document_object.messages.as_slice(),
        DOCUMENT_MESSAGE_TYPES,
        "Pages structured root",
    )?;
    let document = crate::protobuf::tp::DocumentArchive::decode(payload).map_err(|error| {
        crate::Error::InvalidFormat(format!("Pages structured root payload is invalid: {error}"))
    })?;

    let mut section_builder = Section::builder(0, SectionType::Body);
    if let Some(reference) = document.body_storage {
        let object = object_index
            .resolve_ref_id(bundle, reference.identifier)?
            .ok_or_else(|| {
                crate::Error::InvalidFormat(format!(
                    "Pages body storage object {} is missing",
                    reference.identifier
                ))
            })?;
        let payload = unique_payload(
            object.messages,
            STORAGE_MESSAGE_TYPES,
            &format!("Pages body object {}", reference.identifier),
        )?;
        let storage = crate::protobuf::tswp::StorageArchive::decode(payload).map_err(|error| {
            crate::Error::InvalidFormat(format!(
                "Pages body object {} text storage payload is invalid: {error}",
                reference.identifier
            ))
        })?;
        let storage = Storage::from_text(storage.text.concat());
        if !storage.is_empty() {
            section_builder.push_text_storage(storage);
        }
    }

    Ok(vec![section_builder.build()])
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
                "{context} contains duplicate text payloads"
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
