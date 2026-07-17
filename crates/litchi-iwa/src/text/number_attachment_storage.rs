//! Lossless point-table access for native textual number attachments.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::{tsp, tswp};
use crate::wire::{
    patch_length_delimited_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use crate::{Error, IWorkPackage, Result};

use super::storage_wire::{
    STORAGE_MESSAGE_TYPES, StorageLocation, locate_storage, require_text_boundary, text_utf16_len,
};

pub(super) const ATTACHMENT_TABLE_FIELD: u32 = 9;
const TABLE_ENTRIES_FIELD: u32 = 1;
const OBJECT_REPLACEMENT_UNIT: u16 = 0xfffc;

#[derive(Debug)]
pub(super) struct AttachmentEntry {
    pub(super) index: u32,
    pub(super) object_id: u64,
    raw: Vec<u8>,
}

pub(super) fn locate_attachment_storage(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<StorageLocation> {
    let location = locate_storage(package, storage_id, ATTACHMENT_TABLE_FIELD, "attachment")?;
    if location.storage.table_attachment.is_some() != location.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} attachment table wire state is inconsistent"
        )));
    }
    Ok(location)
}

pub(super) fn decoded_attachment_entries(
    storage_id: u64,
    location: &StorageLocation,
) -> Result<Vec<AttachmentEntry>> {
    match (
        location.table_present,
        location.storage.table_attachment.as_ref(),
    ) {
        (false, None) => Ok(Vec::new()),
        (true, Some(table)) if table.entries.is_empty() => Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has an empty attachment table"
        ))),
        (true, Some(table)) => {
            let entries = table
                .entries
                .iter()
                .map(|entry| {
                    let object_id = entry.object.as_ref().map_or(0, |value| value.identifier);
                    AttachmentEntry {
                        index: entry.character_index,
                        object_id,
                        raw: Vec::new(),
                    }
                })
                .collect::<Vec<_>>();
            validate_entries(storage_id, &entries, &location.storage.text)?;
            Ok(entries)
        },
        _ => Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} attachment table wire state is inconsistent"
        ))),
    }
}

fn raw_attachment_entries(
    storage_id: u64,
    table: Option<&[u8]>,
    storage: &tswp::StorageArchive,
) -> Result<Vec<AttachmentEntry>> {
    let Some(table) = table else {
        if storage.table_attachment.is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} attachment table wire state is inconsistent"
            )));
        }
        return Ok(Vec::new());
    };
    let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
    if entries.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has an empty attachment table"
        )));
    }
    let entries = entries
        .into_iter()
        .map(|raw| {
            let entry = tswp::object_attribute_table::ObjectAttribute::decode(raw)?;
            Ok(AttachmentEntry {
                index: entry.character_index,
                object_id: entry.object.map_or(0, |value| value.identifier),
                raw: raw.to_vec(),
            })
        })
        .collect::<Result<Vec<_>>>()?;
    validate_entries(storage_id, &entries, &storage.text)?;
    Ok(entries)
}

fn validate_entries(storage_id: u64, entries: &[AttachmentEntry], text: &[String]) -> Result<()> {
    let text_len = text_utf16_len(text)?;
    let mut previous = None;
    for entry in entries {
        if entry.object_id == 0 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has a zero attachment object identifier"
            )));
        }
        if previous.is_some_and(|index| index >= entry.index) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} attachment positions are not strictly increasing"
            )));
        }
        require_text_boundary(storage_id, entry.index, text)?;
        if entry.index >= text_len
            || utf16_unit_at(text, entry.index) != Some(OBJECT_REPLACEMENT_UNIT)
        {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} attachment {} is not anchored to U+FFFC at UTF-16 index {}",
                entry.object_id, entry.index
            )));
        }
        previous = Some(entry.index);
    }
    Ok(())
}

fn utf16_unit_at(text: &[String], requested: u32) -> Option<u16> {
    let mut index = 0u32;
    for fragment in text {
        for unit in fragment.encode_utf16() {
            if index == requested {
                return Some(unit);
            }
            index = index.checked_add(1)?;
        }
    }
    None
}

pub(super) fn insert_attachment_reference(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    position: u32,
    identifier: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| STORAGE_MESSAGE_TYPES.contains(&message.type_))
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} must have exactly one writable payload"
            )));
        };
        let original = &object.messages[*index];
        let storage = tswp::StorageArchive::decode(original.data.as_slice())?;
        let tables = repeated_length_delimited_payloads(
            original.data.as_slice(),
            ATTACHMENT_TABLE_FIELD,
        )?;
        if tables.len() > 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} contains {} attachment tables",
                tables.len()
            )));
        }
        let mut entries = raw_attachment_entries(storage_id, tables.first().copied(), &storage)?;
        if entries.iter().any(|entry| entry.index == position) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} already has an attachment at UTF-16 index {position}"
            )));
        }
        require_text_boundary(storage_id, position, &storage.text)?;
        if utf16_unit_at(&storage.text, position) != Some(OBJECT_REPLACEMENT_UNIT) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has no U+FFFC at UTF-16 index {position}"
            )));
        }
        let new_entry = tswp::object_attribute_table::ObjectAttribute {
            character_index: position,
            object: Some(tsp::Reference {
                identifier,
                deprecated_type: None,
                deprecated_is_external: None,
            }),
        };
        entries.push(AttachmentEntry {
            index: position,
            object_id: identifier,
            raw: new_entry.encode_to_vec(),
        });
        entries.sort_by_key(|entry| entry.index);
        let encoded_entries = entries
            .into_iter()
            .map(|entry| entry.raw)
            .collect::<Vec<_>>();
        let table = match tables.first() {
            Some(table) => rewrite_repeated_length_delimited_fields(
                table,
                TABLE_ENTRIES_FIELD,
                &encoded_entries,
            )?,
            None => tswp::ObjectAttributeTable {
                entries: encoded_entries
                    .iter()
                    .map(|entry| {
                        tswp::object_attribute_table::ObjectAttribute::decode(entry.as_slice())
                    })
                    .collect::<std::result::Result<Vec<_>, _>>()?,
            }
            .encode_to_vec(),
        };
        let data = patch_length_delimited_field(
            &original.data,
            ATTACHMENT_TABLE_FIELD,
            !tables.is_empty(),
            Some(&table),
        )?;
        object.replace_message(
            *index,
            RawMessage {
                type_: original.type_,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[*index];
        if info.object_references.contains(&identifier) {
            return Err(Error::InvalidFormat(format!(
                "iWork storage metadata already references new attachment object {identifier}"
            )));
        }
        info.object_references.push(identifier);
        Ok(())
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn replacement_unit_lookup_spans_fragments() {
        let text = vec!["A😀".to_owned(), "\u{fffc}B".to_owned()];
        assert_eq!(utf16_unit_at(&text, 0), Some(b'A'.into()));
        assert_eq!(utf16_unit_at(&text, 3), Some(OBJECT_REPLACEMENT_UNIT));
        assert_eq!(utf16_unit_at(&text, 5), None);
    }
}
