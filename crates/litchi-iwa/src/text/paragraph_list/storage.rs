//! Uniform list-style references in TSWP text storage objects.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::{tsp, tswp};
use crate::wire::{
    parse_wire_fields, patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

const STORAGE_MESSAGE_TYPES: &[u32] = &[2_001, 2_022];
const LIST_STYLE_TABLE_FIELD: u32 = 7;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_CHARACTER_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

pub(super) struct ListStorageLocation {
    pub(super) archive_name: String,
    pub(super) style_id: u64,
}

pub(super) struct ListBoundaryStorage {
    pub(super) archive_name: String,
    pub(super) boundaries: Vec<(u32, u64)>,
    pub(super) paragraph_starts: Vec<u32>,
}

pub(super) fn locate(package: &IWorkPackage, storage_id: u64) -> Result<ListStorageLocation> {
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
        repeated_length_delimited_payloads(message.data.as_slice(), LIST_STYLE_TABLE_FIELD)?.len();
    if table_count != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must contain one list-style table, found {table_count}"
        )));
    }
    let entries = storage
        .table_list_style
        .as_ref()
        .map(|table| table.entries.as_slice())
        .unwrap_or_default();
    let [entry] = entries else {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must have one uniform list-style boundary"
        )));
    };
    if entry.character_index != 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} list style must begin at UTF-16 index zero"
        )));
    }
    let style_id = entry
        .object
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has no uniform list style"
            ))
        })?;
    Ok(ListStorageLocation {
        archive_name,
        style_id,
    })
}

pub(super) fn locate_boundaries(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<ListBoundaryStorage> {
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
        repeated_length_delimited_payloads(message.data.as_slice(), LIST_STYLE_TABLE_FIELD)?.len();
    if table_count != 1 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must contain one list-style table, found {table_count}"
        )));
    }
    let paragraph_starts = super::levels::paragraph_starts(&storage.text)?;
    let entries = storage
        .table_list_style
        .as_ref()
        .map(|table| table.entries.as_slice())
        .unwrap_or_default();
    let mut boundaries = Vec::with_capacity(entries.len());
    let mut previous = None;
    for entry in entries {
        if previous.is_some_and(|index| index >= entry.character_index) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} list-style boundaries are not strictly increasing"
            )));
        }
        if paragraph_starts
            .binary_search(&entry.character_index)
            .is_err()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} list-style boundary {} is not a paragraph start",
                entry.character_index
            )));
        }
        let style_id = entry
            .object
            .as_ref()
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork text storage {storage_id} has an empty list-style reference"
                ))
            })?;
        boundaries.push((entry.character_index, style_id));
        previous = Some(entry.character_index);
    }
    if boundaries.first().map(|entry| entry.0) != Some(0) {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} list style must begin at UTF-16 index zero"
        )));
    }
    Ok(ListBoundaryStorage {
        archive_name,
        boundaries,
        paragraph_starts,
    })
}

pub(super) fn replace_boundaries(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    old_style_ids: &[u64],
    boundaries: &[(u32, u64)],
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
        let data =
            transform_length_delimited_field(&original.data, LIST_STYLE_TABLE_FIELD, |table| {
                replace_boundary_table(table, boundaries)
            })?;
        object.replace_message(
            *index,
            RawMessage {
                type_: original.type_,
                data,
            },
        )?;
        let replacements = boundaries.iter().map(|entry| entry.1).collect::<Vec<_>>();
        let info = &mut object.archive_info.message_infos[*index];
        replace_reference_sequence(
            &mut info.object_references,
            old_style_ids,
            &replacements,
            storage_id,
        )?;
        for field in &mut info.field_infos {
            if field
                .object_references
                .iter()
                .any(|reference| old_style_ids.contains(reference))
            {
                replace_reference_sequence(
                    &mut field.object_references,
                    old_style_ids,
                    &replacements,
                    storage_id,
                )?;
            }
        }
        Ok(())
    })
}

fn replace_boundary_table(table: &[u8], boundaries: &[(u32, u64)]) -> Result<Vec<u8>> {
    let existing = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
    let mut encoded = Vec::with_capacity(boundaries.len());
    for &(character_index, style_id) in boundaries {
        let raw = existing
            .iter()
            .find(|payload| {
                required_varint(
                    payload,
                    ENTRY_CHARACTER_INDEX_FIELD,
                    "list-style character index",
                )
                .ok()
                    == Some(u64::from(character_index))
            })
            .map(|payload| {
                transform_length_delimited_field(payload, ENTRY_OBJECT_FIELD, |reference| {
                    patch_varint_field(reference, REFERENCE_IDENTIFIER_FIELD, true, Some(style_id))
                })
            })
            .transpose()?
            .unwrap_or_else(|| {
                tswp::object_attribute_table::ObjectAttribute {
                    character_index,
                    object: Some(tsp::Reference {
                        identifier: style_id,
                        ..Default::default()
                    }),
                }
                .encode_to_vec()
            });
        encoded.push(raw);
    }
    rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &encoded)
}

fn replace_reference_sequence(
    references: &mut Vec<u64>,
    old_style_ids: &[u64],
    replacements: &[u64],
    storage_id: u64,
) -> Result<()> {
    let positions = references
        .iter()
        .enumerate()
        .filter_map(|(index, reference)| old_style_ids.contains(reference).then_some(index))
        .collect::<Vec<_>>();
    if positions.len() != old_style_ids.len() {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} metadata contains {} list-style references, expected {}",
            positions.len(),
            old_style_ids.len()
        )));
    }
    let insertion = positions.first().copied().unwrap_or(references.len());
    references.retain(|reference| !old_style_ids.contains(reference));
    references.splice(insertion..insertion, replacements.iter().copied());
    Ok(())
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
        let data =
            transform_length_delimited_field(&original.data, LIST_STYLE_TABLE_FIELD, |table| {
                let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
                let [entry] = entries.as_slice() else {
                    return Err(Error::InvalidFormat(format!(
                        "iWork text storage {storage_id} must have one uniform list-style boundary"
                    )));
                };
                if required_varint(
                    entry,
                    ENTRY_CHARACTER_INDEX_FIELD,
                    "list-style character index",
                )? != 0
                {
                    return Err(Error::InvalidFormat(format!(
                        "iWork text storage {storage_id} list style must begin at index zero"
                    )));
                }
                let reference =
                    required_payload(entry, ENTRY_OBJECT_FIELD, "list-style reference")?;
                if required_varint(
                    reference,
                    REFERENCE_IDENTIFIER_FIELD,
                    "list-style identifier",
                )? != old_style_id
                {
                    return Err(Error::InvalidFormat(format!(
                        "iWork text storage {storage_id} list style changed unexpectedly"
                    )));
                }
                let patched =
                    transform_length_delimited_field(entry, ENTRY_OBJECT_FIELD, |reference| {
                        patch_varint_field(
                            reference,
                            REFERENCE_IDENTIFIER_FIELD,
                            true,
                            Some(new_style_id),
                        )
                    })?;
                rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &[patched])
            })?;
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
                "iWork text storage {storage_id} metadata contains {replaced} list-style references"
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
