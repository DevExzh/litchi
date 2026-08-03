//! Uniform paragraph-style references in TSWP text storage objects.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::tswp;
use crate::wire::{
    parse_wire_fields, patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

use crate::text::storage_wire::{StorageLocation, locate_storage as locate_native_storage};
const PARAGRAPH_STYLE_TABLE_FIELD: u32 = 5;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_CHARACTER_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

pub(in crate::text) struct ParagraphStorageLocation {
    pub(in crate::text) wire: StorageLocation,
    pub(in crate::text) style_id: u64,
}

pub(in crate::text) fn locate(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<ParagraphStorageLocation> {
    let wire = locate_native_storage(
        package,
        storage_id,
        PARAGRAPH_STYLE_TABLE_FIELD,
        "paragraph-style",
    )?;
    if wire.storage.table_para_style.is_some() != wire.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} paragraph-style table wire state is inconsistent"
        )));
    }
    let entries = wire
        .storage
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
    Ok(ParagraphStorageLocation { wire, style_id })
}

pub(in crate::text) fn patch_style_reference(
    package: &mut IWorkPackage,
    location: &ParagraphStorageLocation,
    old_style_id: u64,
    new_style_id: u64,
) -> Result<()> {
    let storage_id = location.wire.object_id;
    package.update_archive(&location.wire.archive_name, |archive| {
        let object = archive.object_mut(location.wire.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
        })?;
        if object.archive_info.identifier != Some(location.wire.object_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} object identity changed unexpectedly"
            )));
        }
        if object.messages.get(location.wire.message_index).is_none() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload index {} is missing",
                location.wire.message_index
            )));
        }
        if object.messages[location.wire.message_index].type_ != location.wire.message_type {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload changed unexpectedly"
            )));
        }
        if object
            .archive_info
            .message_infos
            .get(location.wire.message_index)
            .is_none()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload metadata index {} is missing",
                location.wire.message_index
            )));
        }
        let (message_type, data) = {
            let original = &object.messages[location.wire.message_index];
            let before = tswp::StorageArchive::decode(original.data.as_slice())?;
            let tables = repeated_length_delimited_payloads(
                original.data.as_slice(),
                PARAGRAPH_STYLE_TABLE_FIELD,
            )?;
            if tables.len() != usize::from(location.wire.table_present)
                || !location.wire.table_present
                || before.table_para_style.is_none()
            {
                return Err(Error::InvalidFormat(format!(
                    "iWork text storage {storage_id} paragraph-style table wire state changed unexpectedly"
                )));
            }
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
            tswp::StorageArchive::decode(data.as_slice())?;
            (original.type_, data)
        };
        object.replace_message(
            location.wire.message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(location.wire.message_index)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork text storage {storage_id} writable payload metadata index {} is missing",
                    location.wire.message_index
                ))
            })?;
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
