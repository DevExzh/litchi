//! Byte-preserving paragraph-start Drop Cap references in TSWP storage.

use prost::Message;

use crate::archive::{Archive, RawMessage};
use crate::protobuf::{tsp, tswp};
use crate::wire::{
    parse_wire_fields, patch_length_delimited_field, patch_varint_field,
    repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields,
};
use crate::{Error, IWorkPackage, Result};

use super::types::ParagraphStart;
use crate::text::storage_wire::{
    LocatedStorage, StorageLocation,
    locate_storage_with_archive as locate_native_storage_with_archive, update_parsed_archive,
};
const DROP_CAP_TABLE_FIELD: u32 = 28;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_CHARACTER_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct DropCapEntry {
    pub(super) paragraph_start: ParagraphStart,
    pub(super) style_id: Option<u64>,
}

pub(super) struct DropCapStorageLocation {
    pub(super) object_id: u64,
    pub(super) archive_name: String,
    pub(super) message_index: usize,
    pub(super) message_type: u32,
    pub(super) table_present: bool,
    pub(super) stylesheet_id: u64,
    pub(super) text: String,
    pub(super) entries: Vec<DropCapEntry>,
}

pub(super) struct LocatedDropCapStorage {
    pub(super) location: DropCapStorageLocation,
    pub(super) archive: Archive,
    pub(super) storage: tswp::StorageArchive,
}

pub(super) fn locate(package: &IWorkPackage, storage_id: u64) -> Result<DropCapStorageLocation> {
    locate_with_archive(package, storage_id).map(|located| located.location)
}

pub(super) fn locate_with_archive(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<LocatedDropCapStorage> {
    let LocatedStorage { location, archive } =
        locate_native_storage_with_archive(package, storage_id, DROP_CAP_TABLE_FIELD, "Drop Cap")?;
    let StorageLocation {
        object_id,
        archive_name,
        message_index,
        message_type,
        storage,
        table_present,
    } = location;
    if storage.table_drop_cap_style.is_some() != table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has a malformed Drop Cap table"
        )));
    }
    let stylesheet_id = storage
        .style_sheet
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} has no stylesheet"))
        })?;
    let text = storage.text.concat();
    let mut previous = None;
    let mut entries = Vec::new();
    for entry in storage
        .table_drop_cap_style
        .iter()
        .flat_map(|table| &table.entries)
    {
        let start = ParagraphStart::from_utf16_index(entry.character_index as usize)?;
        validate_utf16_boundary(&text, start)?;
        if previous.is_some_and(|value| value >= start) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} Drop Cap boundaries are not strictly increasing"
            )));
        }
        previous = Some(start);
        let style_id = entry
            .object
            .as_ref()
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0);
        entries.push(DropCapEntry {
            paragraph_start: start,
            style_id,
        });
    }
    Ok(LocatedDropCapStorage {
        location: DropCapStorageLocation {
            object_id,
            archive_name,
            message_index,
            message_type,
            table_present,
            stylesheet_id,
            text,
            entries,
        },
        archive,
        storage,
    })
}

pub(super) fn validate_paragraph_start(text: &str, start: ParagraphStart) -> Result<()> {
    validate_utf16_boundary(text, start)?;
    let target = usize::try_from(start.utf16_index()).map_err(|_| {
        Error::InvalidFormat("paragraph start exceeds this platform's index range".to_owned())
    })?;
    let mut units = 0usize;
    let mut previous = None;
    for character in text.chars() {
        if units == target {
            if target != 0 && previous != Some('\n') {
                return Err(Error::InvalidFormat(format!(
                    "UTF-16 index {target} is not a paragraph start"
                )));
            }
            if character == '\n' {
                return Err(Error::InvalidFormat(format!(
                    "UTF-16 index {target} begins an empty paragraph"
                )));
            }
            return Ok(());
        }
        units += character.len_utf16();
        previous = Some(character);
    }
    Err(Error::InvalidFormat(format!(
        "UTF-16 index {target} has no paragraph text"
    )))
}

pub(super) fn patch_entry(
    package: &mut IWorkPackage,
    located: LocatedDropCapStorage,
    paragraph_start: ParagraphStart,
    old_style_id: Option<u64>,
    new_style_id: Option<u64>,
) -> Result<()> {
    let LocatedDropCapStorage {
        location,
        archive,
        storage,
    } = located;
    let storage_id = location.object_id;
    update_parsed_archive(package, &location.archive_name, archive, |archive| {
        let object = archive.object_mut(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text storage {storage_id} is missing"))
        })?;
        if object.archive_info.identifier != Some(location.object_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} object identity changed unexpectedly"
            )));
        }
        if object.messages.get(location.message_index).is_none() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload index {} is missing",
                location.message_index
            )));
        }
        if object.messages[location.message_index].type_ != location.message_type {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload changed unexpectedly"
            )));
        }
        if object
            .archive_info
            .message_infos
            .get(location.message_index)
            .is_none()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} writable payload metadata index {} is missing",
                location.message_index
            )));
        }
        let (message_type, data, change) = {
            let original = &object.messages[location.message_index];
            let raw_tables =
                repeated_length_delimited_payloads(&original.data, DROP_CAP_TABLE_FIELD)?;
            if raw_tables.len() != usize::from(location.table_present) {
                return Err(Error::InvalidFormat(format!(
                    "iWork text storage {storage_id} Drop Cap table wire state changed unexpectedly"
                )));
            }
            let table = raw_tables.first().copied().unwrap_or_default();
            let mut entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?
                .into_iter()
                .map(|raw| decode_raw_entry(raw).map(|entry| (entry, raw.to_vec())))
                .collect::<Result<Vec<_>>>()?;
            let matches = entries
                .iter()
                .enumerate()
                .filter(|(_, (entry, _))| entry.paragraph_start == paragraph_start)
                .map(|(index, _)| index)
                .collect::<Vec<_>>();
            if matches.len() > 1 {
                return Err(Error::InvalidFormat(format!(
                    "iWork text storage {storage_id} has duplicate Drop Cap boundaries"
                )));
            }
            if matches.first().and_then(|index| entries[*index].0.style_id) != old_style_id
                || (matches.is_empty() && old_style_id.is_some())
            {
                return Err(Error::InvalidFormat(format!(
                    "iWork text storage {storage_id} Drop Cap reference changed unexpectedly"
                )));
            }
            if let Some(index) = matches.first().copied() {
                entries[index].1 = patch_raw_entry(&entries[index].1, old_style_id, new_style_id)?;
                entries[index].0.style_id = new_style_id;
            } else if new_style_id.is_some() {
                let entry = tswp::object_attribute_table::ObjectAttribute {
                    character_index: paragraph_start.utf16_index(),
                    object: new_style_id.map(reference),
                };
                entries.push((
                    DropCapEntry {
                        paragraph_start,
                        style_id: new_style_id,
                    },
                    entry.encode_to_vec(),
                ));
            }
            entries.sort_by_key(|(entry, _)| entry.paragraph_start);
            if entries
                .windows(2)
                .any(|pair| pair[0].0.paragraph_start == pair[1].0.paragraph_start)
            {
                return Err(Error::InvalidFormat(format!(
                    "iWork text storage {storage_id} has duplicate Drop Cap boundaries"
                )));
            }
            let encoded_entries = entries
                .iter()
                .map(|(_, raw)| raw.clone())
                .collect::<Vec<_>>();
            let table = rewrite_repeated_length_delimited_fields(
                table,
                TABLE_ENTRIES_FIELD,
                &encoded_entries,
            )?;
            let data = patch_length_delimited_field(
                &original.data,
                DROP_CAP_TABLE_FIELD,
                location.table_present,
                Some(&table),
            )?;
            let after = tswp::StorageArchive::decode(data.as_slice())?;
            let old_count =
                count_drop_cap_reference(storage.table_drop_cap_style.as_ref(), old_style_id);
            let new_before =
                count_drop_cap_reference(storage.table_drop_cap_style.as_ref(), new_style_id);
            let old_after =
                count_drop_cap_reference(after.table_drop_cap_style.as_ref(), old_style_id);
            let new_after =
                count_drop_cap_reference(after.table_drop_cap_style.as_ref(), new_style_id);
            (
                original.type_,
                data,
                ReferenceMetadataChange {
                    old_style_id,
                    new_style_id,
                    old_count,
                    old_after,
                    new_before,
                    new_after,
                },
            )
        };
        object.replace_message(
            location.message_index,
            RawMessage {
                type_: message_type,
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
        update_reference_metadata(info, change)?;
        Ok(())
    })
}

fn validate_utf16_boundary(text: &str, start: ParagraphStart) -> Result<()> {
    let target = usize::try_from(start.utf16_index()).map_err(|_| {
        Error::InvalidFormat("paragraph start exceeds this platform's index range".to_owned())
    })?;
    let mut units = 0usize;
    for character in text.chars() {
        if units == target {
            return Ok(());
        }
        units = units
            .checked_add(character.len_utf16())
            .ok_or_else(|| Error::InvalidFormat("iWork text UTF-16 length overflow".to_owned()))?;
        if units > target {
            return Err(Error::InvalidFormat(format!(
                "UTF-16 index {target} splits a surrogate pair"
            )));
        }
    }
    if units == target {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "UTF-16 index {target} exceeds text length {units}"
        )))
    }
}

fn decode_raw_entry(raw: &[u8]) -> Result<DropCapEntry> {
    let character_index = u32::try_from(required_varint(
        raw,
        ENTRY_CHARACTER_INDEX_FIELD,
        "Drop Cap character index",
    )?)
    .map_err(|_| Error::InvalidFormat("Drop Cap character index exceeds u32".to_owned()))?;
    let references = repeated_length_delimited_payloads(raw, ENTRY_OBJECT_FIELD)?;
    let style_id = match references.as_slice() {
        [] => None,
        [raw] => {
            let identifier = tsp::Reference::decode(*raw)?.identifier;
            if identifier == 0 {
                return Err(Error::InvalidFormat(
                    "Drop Cap style reference is zero".to_owned(),
                ));
            }
            Some(identifier)
        },
        _ => {
            return Err(Error::InvalidFormat(
                "Drop Cap boundary has multiple style references".to_owned(),
            ));
        },
    };
    Ok(DropCapEntry {
        paragraph_start: ParagraphStart::from_utf16_index(character_index as usize)?,
        style_id,
    })
}

fn patch_raw_entry(
    raw: &[u8],
    old_style_id: Option<u64>,
    new_style_id: Option<u64>,
) -> Result<Vec<u8>> {
    match (old_style_id, new_style_id) {
        (Some(old), Some(new)) => {
            let reference_raw = required_payload(raw, ENTRY_OBJECT_FIELD, "Drop Cap reference")?;
            if required_varint(
                reference_raw,
                REFERENCE_IDENTIFIER_FIELD,
                "Drop Cap identifier",
            )? != old
            {
                return Err(Error::InvalidFormat(
                    "Drop Cap style reference changed during mutation".to_owned(),
                ));
            }
            let reference =
                patch_varint_field(reference_raw, REFERENCE_IDENTIFIER_FIELD, true, Some(new))?;
            patch_length_delimited_field(raw, ENTRY_OBJECT_FIELD, true, Some(&reference))
        },
        (Some(old), None) => {
            let reference_raw = required_payload(raw, ENTRY_OBJECT_FIELD, "Drop Cap reference")?;
            if required_varint(
                reference_raw,
                REFERENCE_IDENTIFIER_FIELD,
                "Drop Cap identifier",
            )? != old
            {
                return Err(Error::InvalidFormat(
                    "Drop Cap style reference changed during mutation".to_owned(),
                ));
            }
            patch_length_delimited_field(raw, ENTRY_OBJECT_FIELD, true, None)
        },
        (None, Some(new)) => {
            let reference = reference(new).encode_to_vec();
            patch_length_delimited_field(raw, ENTRY_OBJECT_FIELD, false, Some(&reference))
        },
        (None, None) => Ok(raw.to_vec()),
    }
}

fn count_drop_cap_reference(
    table: Option<&tswp::ObjectAttributeTable>,
    identifier: Option<u64>,
) -> usize {
    let Some(identifier) = identifier else {
        return 0;
    };
    table
        .into_iter()
        .flat_map(|table| &table.entries)
        .filter(|entry| {
            entry
                .object
                .as_ref()
                .is_some_and(|reference| reference.identifier == identifier)
        })
        .count()
}

struct ReferenceMetadataChange {
    old_style_id: Option<u64>,
    new_style_id: Option<u64>,
    old_count: usize,
    old_after: usize,
    new_before: usize,
    new_after: usize,
}

fn update_reference_metadata(
    info: &mut crate::archive::MessageInfo,
    change: ReferenceMetadataChange,
) -> Result<()> {
    if let Some(old) = change.old_style_id
        && change.old_count > 0
        && change.old_after == 0
    {
        let count = info
            .object_references
            .iter()
            .filter(|&&identifier| identifier == old)
            .count();
        if count != 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork storage metadata references Drop Cap style {old} {count} times"
            )));
        }
        info.object_references
            .retain(|&identifier| identifier != old);
        for field in &mut info.field_infos {
            if field.path.path.first() == Some(&DROP_CAP_TABLE_FIELD) {
                field
                    .object_references
                    .retain(|&identifier| identifier != old);
            }
        }
    }
    if let Some(new) = change.new_style_id
        && change.new_before == 0
        && change.new_after > 0
    {
        if info.object_references.contains(&new) {
            return Err(Error::InvalidFormat(format!(
                "iWork storage metadata already references new Drop Cap style {new}"
            )));
        }
        info.object_references.push(new);
        for field in &mut info.field_infos {
            if field.path.path.first() == Some(&DROP_CAP_TABLE_FIELD)
                && !field.object_references.contains(&new)
            {
                field.object_references.push(new);
            }
        }
    }
    Ok(())
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
        .filter(|field| field.number() == field_number && field.wire_type() == 0)
        .collect::<Vec<_>>();
    let [field] = matches.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{context} must contain varint field {field_number} exactly once"
        )));
    };
    let (value, length) =
        litchi_iwa_common::varint::decode_varint_from_bytes(&data[field.key_end()..field.end()])
            .map_err(|error| Error::InvalidFormat(format!("invalid {context}: {error}")))?;
    if field.key_end() + length != field.end() {
        return Err(Error::InvalidFormat(format!(
            "{context} has trailing varint bytes"
        )));
    }
    Ok(value)
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
