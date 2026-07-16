//! Uniform paragraph-style references in TSWP text storage objects.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::tswp;
use crate::wire::{
    parse_wire_fields, patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const PARAGRAPH_STYLE_TABLE_FIELD: u32 = 5;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_CHARACTER_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

pub(super) struct ParagraphStorageLocation {
    pub(super) archive_name: String,
    pub(super) style_id: u64,
}

pub(super) fn locate(package: &IWorkPackage, storage_id: u64) -> Result<ParagraphStorageLocation> {
    let archive_name = object_archive_name(package, storage_id)?;
    let archive = package.archive(&archive_name)?;
    let object = archive.object(storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
    })?;
    let payloads = object
        .messages
        .iter()
        .filter_map(|message| {
            STORAGE_MESSAGE_TYPES
                .contains(&message.type_)
                .then(|| {
                    tswp::StorageArchive::decode(message.data.as_slice())
                        .ok()
                        .map(|storage| (message, storage))
                })
                .flatten()
        })
        .collect::<Vec<_>>();
    let [(message, storage)] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must have exactly one writable payload"
        )));
    };
    let table_count =
        repeated_length_delimited_payloads(message.data.as_slice(), PARAGRAPH_STYLE_TABLE_FIELD)?
            .len();
    if table_count != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must contain one paragraph-style table, found {table_count}"
        )));
    }
    let entries = storage
        .table_para_style
        .as_ref()
        .map(|table| table.entries.as_slice())
        .unwrap_or_default();
    let [entry] = entries else {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must have one uniform paragraph-style boundary"
        )));
    };
    if entry.character_index != 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} paragraph style must begin at UTF-16 index zero"
        )));
    }
    let style_id = entry
        .object
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has no uniform paragraph style"
            ))
        })?;
    Ok(ParagraphStorageLocation {
        archive_name,
        style_id,
    })
}

pub(super) fn patch_style_reference(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    old_style_id: u64,
    new_style_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| {
                STORAGE_MESSAGE_TYPES.contains(&message.type_)
                    && tswp::StorageArchive::decode(message.data.as_slice()).is_ok()
            })
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} must have exactly one writable payload"
            )));
        };
        let original = &object.messages[*index];
        let data = transform_length_delimited_field(
            &original.data,
            PARAGRAPH_STYLE_TABLE_FIELD,
            |table| {
                let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
                let [entry] = entries.as_slice() else {
                    return Err(Error::InvalidFormat(format!(
                        "iWork text storage {storage_id} must have one uniform paragraph-style boundary"
                    )));
                };
                if required_varint(
                    entry,
                    ENTRY_CHARACTER_INDEX_FIELD,
                    "paragraph-style character index",
                )? != 0
                {
                    return Err(Error::InvalidFormat(format!(
                        "iWork text storage {storage_id} paragraph style must begin at index zero"
                    )));
                }
                let reference = required_payload(
                    entry,
                    ENTRY_OBJECT_FIELD,
                    "paragraph-style reference",
                )?;
                if required_varint(
                    reference,
                    REFERENCE_IDENTIFIER_FIELD,
                    "paragraph-style identifier",
                )? != old_style_id
                {
                    return Err(Error::InvalidFormat(format!(
                        "iWork text storage {storage_id} paragraph style changed unexpectedly"
                    )));
                }
                let patched = transform_length_delimited_field(
                    entry,
                    ENTRY_OBJECT_FIELD,
                    |reference| {
                        patch_varint_field(
                            reference,
                            REFERENCE_IDENTIFIER_FIELD,
                            true,
                            Some(new_style_id),
                        )
                    },
                )?;
                rewrite_repeated_length_delimited_fields(
                    table,
                    TABLE_ENTRIES_FIELD,
                    &[patched],
                )
            },
        )?;
        object.replace_message(
            *index,
            RawMessage {
                type_: original.type_,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[*index];
        let mut replaced = 0usize;
        for reference in &mut info.object_references {
            if *reference == old_style_id {
                *reference = new_style_id;
                replaced += 1;
            }
        }
        for field in &mut info.field_infos {
            for reference in &mut field.object_references {
                if *reference == old_style_id {
                    *reference = new_style_id;
                }
            }
        }
        if replaced != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} metadata contains {replaced} paragraph-style references"
            )));
        }
        Ok(())
    })
}

fn object_archive_name(package: &IWorkPackage, identifier: u64) -> Result<String> {
    let mut found = None;
    for name in package.iwa_entry_names() {
        if package.archive(name)?.object(identifier).is_some()
            && found.replace(name.to_owned()).is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork object {identifier} occurs in multiple archives"
            )));
        }
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("iWork object {identifier} is missing")))
}

fn required_payload<'a>(data: &'a [u8], field: u32, context: &str) -> Result<&'a [u8]> {
    let payloads = repeated_length_delimited_payloads(data, field)?;
    let [payload] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{context} must contain field {field} exactly once"
        )));
    };
    Ok(payload)
}

fn required_varint(data: &[u8], field_number: u32, context: &str) -> Result<u64> {
    let fields = parse_wire_fields(data)?;
    let matches = fields
        .iter()
        .filter(|field| field.number == field_number && field.wire_type == 0)
        .collect::<Vec<_>>();
    let [field] = matches.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{context} must contain varint field {field_number} exactly once"
        )));
    };
    let (value, length) = crate::varint::decode_varint_from_bytes(&data[field.key_end..field.end])
        .map_err(|error| Error::InvalidFormat(format!("invalid {context}: {error}")))?;
    if field.key_end + length != field.end {
        return Err(Error::InvalidFormat(format!(
            "{context} has trailing varint bytes"
        )));
    }
    Ok(value)
}
