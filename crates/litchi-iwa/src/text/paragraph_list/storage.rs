//! Uniform list-style references in TSWP text storage objects.

use prost::Message;

use crate::archive::{Archive, RawMessage};
use crate::protobuf::{tsp, tswp};
use crate::wire::{
    parse_wire_fields, patch_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields, transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

use super::super::storage_wire::{
    LocatedStorage, StorageLocation,
    locate_storage_with_archive as locate_native_storage_with_archive, update_parsed_archive,
};

const LIST_STYLE_TABLE_FIELD: u32 = 7;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_CHARACTER_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

pub(super) struct ListStorageLocation {
    pub(super) object_id: u64,
    pub(super) archive_name: String,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
    pub(super) style_id: u64,
}

pub(super) struct ListBoundaryStorage {
    pub(super) object_id: u64,
    pub(super) archive_name: String,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
    pub(super) boundaries: Vec<(u32, u64)>,
    pub(super) paragraph_starts: Vec<u32>,
}

pub(super) struct LocatedListBoundaryStorage {
    pub(super) location: ListBoundaryStorage,
    pub(super) archive: Archive,
}

pub(super) fn locate(package: &IWorkPackage, storage_id: u64) -> Result<ListStorageLocation> {
    let location = locate_storage(package, storage_id)?;
    let Some(table) = location.storage.table_list_style.as_ref() else {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must contain one list-style table, found 0"
        )));
    };
    let entries = table.entries.as_slice();
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
        object_id: location.object_id,
        archive_name: location.archive_name,
        message_index: location.message_index,
        message_type: location.message_type,
        style_id,
    })
}

pub(super) fn locate_boundaries(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<ListBoundaryStorage> {
    locate_boundaries_with_archive(package, storage_id).map(|located| located.location)
}

pub(super) fn locate_boundaries_with_archive(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<LocatedListBoundaryStorage> {
    let LocatedStorage { location, archive } = locate_storage_with_archive(package, storage_id)?;
    let Some(table) = location.storage.table_list_style.as_ref() else {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must contain one list-style table, found 0"
        )));
    };
    let paragraph_starts = super::levels::paragraph_starts(&location.storage.text)?;
    let entries = table.entries.as_slice();
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
    Ok(LocatedListBoundaryStorage {
        location: ListBoundaryStorage {
            object_id: location.object_id,
            archive_name: location.archive_name,
            message_index: location.message_index,
            message_type: location.message_type,
            boundaries,
            paragraph_starts,
        },
        archive,
    })
}

pub(super) fn replace_boundaries(
    package: &mut IWorkPackage,
    location: &ListBoundaryStorage,
    storage_id: u64,
    old_style_ids: &[u64],
    boundaries: &[(u32, u64)],
) -> Result<()> {
    let archive_name = location.archive_name.clone();
    package.update_archive(&archive_name, |archive| {
        replace_boundaries_in_archive(archive, location, storage_id, old_style_ids, boundaries)
    })
}

pub(super) fn replace_boundaries_with_archive(
    package: &mut IWorkPackage,
    located: LocatedListBoundaryStorage,
    storage_id: u64,
    old_style_ids: &[u64],
    boundaries: &[(u32, u64)],
) -> Result<()> {
    let LocatedListBoundaryStorage { location, archive } = located;
    let archive_name = location.archive_name.clone();
    update_parsed_archive(package, &archive_name, archive, |archive| {
        replace_boundaries_in_archive(archive, &location, storage_id, old_style_ids, boundaries)
    })
}

fn replace_boundaries_in_archive(
    archive: &mut Archive,
    location: &ListBoundaryStorage,
    storage_id: u64,
    old_style_ids: &[u64],
    boundaries: &[(u32, u64)],
) -> Result<()> {
    let object = archive.object_mut(storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
    })?;
    if object.archive_info.identifier != Some(location.object_id) {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} object identity changed unexpectedly"
        )));
    }
    let (original_type, data) = {
        let original = object.messages.get(location.message_index).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload index {} is missing",
                location.message_index
            ))
        })?;
        if original.type_ != location.message_type {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload changed unexpectedly"
            )));
        }
        let data =
            transform_length_delimited_field(&original.data, LIST_STYLE_TABLE_FIELD, |table| {
                replace_boundary_table(table, boundaries)
            })?;
        (original.type_, data)
    };
    object.replace_message(
        location.message_index,
        RawMessage {
            type_: original_type,
            data,
        },
    )?;
    let replacements = boundaries.iter().map(|entry| entry.1).collect::<Vec<_>>();
    let info = object
        .archive_info
        .message_infos
        .get_mut(location.message_index)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload metadata index {} is missing",
                location.message_index
            ))
        })?;
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
}

fn replace_boundary_table(table: &[u8], boundaries: &[(u32, u64)]) -> Result<Vec<u8>> {
    let existing = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
    let mut encoded = Vec::with_capacity(boundaries.len());
    for &(character_index, style_id) in boundaries {
        let mut matching = None;
        for payload in &existing {
            let existing_index = required_varint(
                payload,
                ENTRY_CHARACTER_INDEX_FIELD,
                "list-style character index",
            )?;
            if existing_index == u64::from(character_index) && matching.replace(*payload).is_some()
            {
                return Err(Error::InvalidFormat(format!(
                    "list-style character index {character_index} occurs multiple times"
                )));
            }
        }
        let raw = match matching {
            Some(payload) => {
                transform_length_delimited_field(payload, ENTRY_OBJECT_FIELD, |reference| {
                    patch_varint_field(reference, REFERENCE_IDENTIFIER_FIELD, true, Some(style_id))
                })?
            },
            None => tswp::object_attribute_table::ObjectAttribute {
                character_index,
                object: Some(tsp::Reference {
                    identifier: style_id,
                    ..Default::default()
                }),
            }
            .encode_to_vec(),
        };
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
    location: &ListStorageLocation,
    storage_id: u64,
    old_style_id: u64,
    new_style_id: u64,
) -> Result<()> {
    package.update_archive(&location.archive_name, |archive| {
        let object = archive.object_mut(storage_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
        })?;
        if object.archive_info.identifier != Some(location.object_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} object identity changed unexpectedly"
            )));
        }
        let (original_type, data) = {
            let original = object.messages.get(location.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork text storage {storage_id} writable payload index {} is missing",
                    location.message_index
                ))
            })?;
            if original.type_ != location.message_type {
                return Err(Error::InvalidFormat(format!(
                    "iWork text storage {storage_id} writable payload changed unexpectedly"
                )));
            }
            let data = transform_length_delimited_field(&original.data, LIST_STYLE_TABLE_FIELD, |table| {
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
            (original.type_, data)
        };
        object.replace_message(
            location.message_index,
            RawMessage {
                type_: original_type,
                data,
            },
        )?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(location.message_index)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork text storage {storage_id} writable payload metadata index {} is missing",
                    location.message_index
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
                "iWork text storage {storage_id} metadata contains {replaced} list-style references"
            )));
        }
        Ok(())
    })
}

fn locate_storage(package: &IWorkPackage, storage_id: u64) -> Result<StorageLocation> {
    locate_storage_with_archive(package, storage_id).map(|located| located.location)
}

fn locate_storage_with_archive(package: &IWorkPackage, storage_id: u64) -> Result<LocatedStorage> {
    let located = locate_native_storage_with_archive(
        package,
        storage_id,
        LIST_STYLE_TABLE_FIELD,
        "list-style",
    )?;
    let location = &located.location;
    if location.storage.table_list_style.is_some() != location.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} list-style table wire state is inconsistent"
        )));
    }
    if !location.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must contain one list-style table, found 0"
        )));
    }
    Ok(located)
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject};

    #[test]
    fn list_storage_lookup_rejects_malformed_recognized_storage() {
        let object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 2_001,
                data: vec![0x80],
            }],
        )
        .unwrap();
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                "Index/Document.iwa",
                &Archive {
                    objects: vec![object],
                },
            )
            .unwrap();

        assert!(locate(&package, 42).is_err());
        assert!(locate_boundaries(&package, 42).is_err());
    }

    #[test]
    fn boundary_rewrite_rejects_malformed_existing_entry() {
        let table = vec![0x0a, 0x01, 0x80];

        assert!(replace_boundary_table(&table, &[(0, 7)]).is_err());
    }
}
