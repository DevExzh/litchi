//! Uniform list-style references in TSWP text storage objects.

use crate::archive::{Archive, RawMessage};
use crate::wire::{
    append_length_delimited_field, append_varint_field, parse_wire_fields, patch_varint_field,
    repeated_length_delimited_payloads, rewrite_repeated_length_delimited_fields,
    transform_length_delimited_field,
};
use crate::{Error, IWorkPackage, Result};

use super::super::storage_wire::update_parsed_archive;
use litchi_iwa_text::storage::Storage;

const LIST_STYLE_TABLE_FIELD: u32 = 7;
const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_CHARACTER_INDEX_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;
const REFERENCE_IDENTIFIER_FIELD: u32 = 1;
const STORAGE_MESSAGE_TYPES: [u32; 2] = [2_001, 2_022];

struct ListTableEntry {
    character_index: u32,
    style_id: Option<u64>,
}

struct LocatedListStorage {
    location: ListStorageLocation,
    table_present: bool,
    entries: Vec<ListTableEntry>,
    text: Storage,
    archive: Archive,
}

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
    let located = locate_storage_with_archive(package, storage_id)?;
    if !located.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must contain one list-style table, found 0"
        )));
    }
    let entries = located.entries.as_slice();
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
        .style_id
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has no uniform list style"
            ))
        })?;
    Ok(ListStorageLocation {
        object_id: located.location.object_id,
        archive_name: located.location.archive_name,
        message_index: located.location.message_index,
        message_type: located.location.message_type,
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
    let LocatedListStorage {
        location,
        table_present,
        entries,
        text,
        archive,
    } = locate_storage_with_archive(package, storage_id)?;
    if !table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} must contain one list-style table, found 0"
        )));
    }
    let paragraph_starts = paragraph_starts_from_storage(&text)?;
    let entries = entries.as_slice();
    let mut boundaries = Vec::new();
    boundaries.try_reserve_exact(entries.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "iWork text storage list-style boundaries",
            amount: entries.len(),
        })
    })?;
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
            .style_id
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
    let mut encoded = Vec::new();
    encoded.try_reserve_exact(boundaries.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "iWork text storage replacement list-style entries",
            amount: boundaries.len(),
        })
    })?;
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
            None => {
                let mut reference = Vec::new();
                append_varint_field(&mut reference, REFERENCE_IDENTIFIER_FIELD, style_id)?;
                let mut entry = Vec::new();
                append_varint_field(
                    &mut entry,
                    ENTRY_CHARACTER_INDEX_FIELD,
                    u64::from(character_index),
                )?;
                append_length_delimited_field(&mut entry, ENTRY_OBJECT_FIELD, &reference)?;
                entry
            },
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

fn locate_storage_with_archive(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<LocatedListStorage> {
    let mut found = None;
    let mut object_found = false;
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        let Some(object) = archive.object(storage_id) else {
            continue;
        };
        if object_found {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} occurs in multiple archives"
            )));
        }
        object_found = true;
        let Some(payload) = resolve_storage_payload(storage_id, archive_name, &object.messages)?
        else {
            continue;
        };
        if found.is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} occurs in multiple archives"
            )));
        }
        found = Some(LocatedListStorage {
            location: ListStorageLocation {
                object_id: storage_id,
                archive_name: archive_name.to_owned(),
                message_index: payload.message_index,
                message_type: payload.message_type,
                style_id: 0,
            },
            table_present: payload.table_present,
            entries: payload.entries,
            text: payload.text,
            archive,
        });
    }
    found.ok_or_else(|| Error::InvalidFormat(format!("iWork text storage {storage_id} is missing")))
}

struct ResolvedStoragePayload {
    message_index: usize,
    message_type: u32,
    table_present: bool,
    entries: Vec<ListTableEntry>,
    text: Storage,
}

fn resolve_storage_payload(
    storage_id: u64,
    archive_name: &str,
    messages: &[RawMessage],
) -> Result<Option<ResolvedStoragePayload>> {
    let mut found = None;
    for (message_index, message) in messages.iter().enumerate() {
        if !STORAGE_MESSAGE_TYPES.contains(&message.type_) {
            continue;
        }

        // Numbers reuses the native 2022 message type for paragraph styles.
        // Match the shared resolver's exemption before treating a payload as
        // a text storage. A paragraph-style payload normally has no field 7;
        // a field-7 payload remains a candidate and is validated below.
        let paragraph_style =
            message.type_ == 2_022 && is_paragraph_style_wire(message.data.as_slice())?;
        let table_payloads = match repeated_length_delimited_payloads(
            message.data.as_slice(),
            LIST_STYLE_TABLE_FIELD,
        ) {
            Ok(payloads) => payloads,
            Err(_error) if paragraph_style => continue,
            Err(error) => return Err(error),
        };
        if paragraph_style && table_payloads.is_empty() {
            continue;
        }
        if table_payloads.len() > 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} contains {} list-style tables",
                table_payloads.len()
            )));
        }
        validate_storage_root_wire(message.data.as_slice(), storage_id)?;
        let text = project_storage_text(storage_id, archive_name, message_index, &message.data)?;
        let table_present = !table_payloads.is_empty();
        let entries = table_payloads
            .first()
            .map(|table| parse_list_table(table))
            .transpose()?;
        let payload = ResolvedStoragePayload {
            message_index,
            message_type: message.type_,
            table_present,
            entries: entries.unwrap_or_default(),
            text,
        };
        if found.replace(payload).is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} must have exactly one writable payload"
            )));
        }
    }
    Ok(found)
}

/// Identify the Numbers paragraph-style envelope without materializing a
/// generated message.  `TSA.ParagraphStyleArchive` shares type 2022 with
/// `TSWP.StorageArchive`; its required field 1 is a length-delimited
/// `TSS.StyleArchive`, while a storage archive's field 1 is a varint kind.
///
/// This is intentionally a wire-only classifier.  Unknown fields, including
/// their original bytes, are not interpreted or rewritten.  The bounded
/// parser still validates every field framing byte before the known envelope
/// fields are inspected, so a malformed candidate cannot bypass the storage
/// resolver's normal error path.
fn is_paragraph_style_wire(data: &[u8]) -> Result<bool> {
    let source_bytes = data.len().max(1);
    let limits = litchi_iwa_common::WireLimits::default()
        .with_input_bytes(source_bytes.min(litchi_iwa_common::WireLimits::MAX_INPUT_BYTES))
        .and_then(|limits| {
            limits.with_fields(source_bytes.min(litchi_iwa_common::WireLimits::MAX_FIELDS))
        })
        .map_err(|error| {
            Error::InvalidFormat(format!(
                "paragraph-style wire classifier limits are invalid: {error}"
            ))
        })?;
    let fields = litchi_iwa_common::wire::parse_wire_fields_with_limits(data, limits)
        .map_err(|error| Error::InvalidFormat(format!("invalid paragraph-style wire: {error}")))?;
    let mut super_present = false;
    for field in fields {
        let expected_wire_type = match field.number() {
            1 | 11 | 12 => Some(2),
            10 => Some(0),
            _ => None,
        };
        if let Some(expected) = expected_wire_type
            && field.wire_type() != expected
        {
            return Ok(false);
        }
        if field.number() == 1 {
            super_present = true;
        }
    }
    Ok(super_present)
}

fn validate_storage_root_wire(data: &[u8], storage_id: u64) -> Result<()> {
    for field in parse_wire_fields(data)? {
        let expected_wire_type = match field.number() {
            1 | 4 | 10 => Some(0),
            2 | 3 | 5..=9 | 11..=12 | 14..=28 => Some(2),
            _ => None,
        };
        if let Some(expected) = expected_wire_type
            && field.wire_type() != expected
        {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} field {} has wire type {}; expected {expected}",
                field.number(),
                field.wire_type()
            )));
        }
    }
    Ok(())
}

fn parse_list_table(table: &[u8]) -> Result<Vec<ListTableEntry>> {
    let payloads = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
    let mut entries = Vec::new();
    entries.try_reserve_exact(payloads.len()).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "iWork text storage list-style entries",
            amount: payloads.len(),
        })
    })?;
    for payload in payloads {
        let character_index = required_varint(
            payload,
            ENTRY_CHARACTER_INDEX_FIELD,
            "list-style character index",
        )?;
        let character_index = u32::try_from(character_index).map_err(|_| {
            Error::InvalidFormat("list-style character index exceeds u32".to_owned())
        })?;
        let style_id =
            match repeated_length_delimited_payloads(payload, ENTRY_OBJECT_FIELD)?.as_slice() {
                [] => None,
                [reference] => Some(required_varint(
                    reference,
                    REFERENCE_IDENTIFIER_FIELD,
                    "list-style identifier",
                )?),
                _ => {
                    return Err(Error::InvalidFormat(format!(
                        "list-style entry must contain field {ENTRY_OBJECT_FIELD} at most once"
                    )));
                },
            };
        entries.push(ListTableEntry {
            character_index,
            style_id,
        });
    }
    Ok(entries)
}

fn project_storage_text(
    storage_id: u64,
    archive_name: &str,
    message_index: usize,
    data: &[u8],
) -> Result<Storage> {
    let source_bytes = data.len().max(1);
    let limits = litchi_iwa_text_wire::Limits::new(
        source_bytes.min(litchi_iwa_text_wire::Limits::MAX_MESSAGE_BYTES),
        source_bytes.min(litchi_iwa_text_wire::Limits::MAX_FIELDS),
        source_bytes.min(litchi_iwa_text_wire::Limits::MAX_FRAGMENTS),
        source_bytes.min(litchi_iwa_text_wire::Limits::MAX_TEXT_BYTES),
    )
    .map_err(|error| {
        Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has invalid text projection limits: {error}"
        ))
    })?;
    litchi_iwa_text_wire::from_bytes_with_limits(data, limits).map_err(|error| {
        Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has a malformed writable payload in {archive_name} message {message_index}: {error}"
        ))
    })
}

fn paragraph_starts_from_storage(storage: &Storage) -> Result<Vec<u32>> {
    let capacity = storage.text().len().saturating_add(1);
    let mut starts = Vec::new();
    starts.try_reserve_exact(capacity).map_err(|_| {
        Error::IwaCommon(litchi_iwa_common::Error::Allocation {
            resource: "iWork text storage paragraph starts",
            amount: capacity,
        })
    })?;
    starts.push(0);
    let mut index = 0u32;
    let mut previous_was_carriage_return = false;
    for run in storage.runs() {
        let end = run.end().ok_or_else(|| {
            Error::InvalidFormat("iWork text UTF-8 run range overflow".to_owned())
        })?;
        let fragment = storage.text().get(run.start()..end).ok_or_else(|| {
            Error::InvalidFormat("iWork text UTF-8 run range is invalid".to_owned())
        })?;
        for character in fragment.chars() {
            index = index
                .checked_add(character.len_utf16() as u32)
                .ok_or_else(|| {
                    Error::InvalidFormat("iWork text UTF-16 length overflow".to_owned())
                })?;
            match character {
                '\n' => {
                    if previous_was_carriage_return {
                        starts.pop();
                    }
                    starts.push(index);
                },
                '\r' | '\u{2028}' | '\u{2029}' => starts.push(index),
                _ => {},
            }
            previous_was_carriage_return = character == '\r';
        }
    }
    starts.dedup();
    Ok(starts)
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
        .filter(|field| field.number() == field_number)
        .collect::<Vec<_>>();
    let [field] = matches.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{context} must contain varint field {field_number} exactly once"
        )));
    };
    if field.wire_type() != 0 {
        return Err(Error::InvalidFormat(format!(
            "{context} field {field_number} has wire type {}; expected 0",
            field.wire_type()
        )));
    }
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::archive::{Archive, ArchiveObject};
    use crate::protobuf::{tsp, tss, tswp};
    use crate::wire::{append_varint_field, repeated_length_delimited_payloads};
    use prost::Message;

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
    fn paragraph_style_wire_classifier_keeps_prost_fixture_test_only() {
        let style = tswp::ParagraphStyleArchive {
            super_: tss::StyleArchive::default(),
            ..Default::default()
        };
        let mut raw = style.encode_to_vec();
        append_varint_field(&mut raw, 1_901, 7).unwrap();
        assert!(is_paragraph_style_wire(&raw).unwrap());

        let storage = tswp::StorageArchive {
            text: vec!["text".to_owned()],
            ..Default::default()
        };
        assert!(!is_paragraph_style_wire(&storage.encode_to_vec()).unwrap());
    }

    #[test]
    fn wire_projection_preserves_empty_first_fragment_and_crlf_boundaries() {
        let storage = tswp::StorageArchive {
            text: vec![String::new(), "A\r".to_owned(), "\nB".to_owned()],
            table_list_style: Some(tswp::ObjectAttributeTable {
                entries: vec![
                    tswp::object_attribute_table::ObjectAttribute {
                        character_index: 0,
                        object: Some(tsp::Reference {
                            identifier: 7,
                            ..Default::default()
                        }),
                    },
                    tswp::object_attribute_table::ObjectAttribute {
                        character_index: 3,
                        object: Some(tsp::Reference {
                            identifier: 9,
                            ..Default::default()
                        }),
                    },
                ],
            }),
            ..Default::default()
        };
        let raw = storage.encode_to_vec();
        let projected = project_storage_text(42, "Index/Document.iwa", 0, &raw).unwrap();
        assert_eq!(
            projected.runs(),
            [
                litchi_iwa_text::storage::Run::new(0, 0),
                litchi_iwa_text::storage::Run::new(0, 2),
                litchi_iwa_text::storage::Run::new(2, 2),
            ]
        );

        let object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 2_001,
                data: raw,
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
        let located = locate_boundaries_with_archive(&package, 42).unwrap();
        assert_eq!(located.location.paragraph_starts, [0, 3]);
        assert_eq!(located.location.boundaries, [(0, 7), (3, 9)]);
    }

    #[test]
    fn malformed_text_projection_is_rejected_before_any_archive_edit() {
        let storage = tswp::StorageArchive {
            text: vec!["valid".to_owned()],
            table_list_style: Some(tswp::ObjectAttributeTable {
                entries: vec![tswp::object_attribute_table::ObjectAttribute {
                    character_index: 0,
                    object: Some(tsp::Reference {
                        identifier: 7,
                        ..Default::default()
                    }),
                }],
            }),
            ..Default::default()
        };
        let mut raw = storage.encode_to_vec();
        raw.extend_from_slice(&[0x1a, 0x01, 0xff]);
        let object = ArchiveObject::new(
            42,
            vec![RawMessage {
                type_: 2_001,
                data: raw,
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
        let before = package.archive("Index/Document.iwa").unwrap();
        assert!(locate_boundaries_with_archive(&package, 42).is_err());
        assert_eq!(package.archive("Index/Document.iwa").unwrap(), before);
    }

    #[test]
    fn boundary_rewrite_rejects_malformed_existing_entry() {
        let table = vec![0x0a, 0x01, 0x80];

        assert!(replace_boundary_table(&table, &[(0, 7)]).is_err());
    }

    #[test]
    fn located_boundary_rewrite_preserves_unknown_wire_fields() {
        let storage = tswp::StorageArchive {
            text: vec!["First\nSecond".to_owned()],
            table_list_style: Some(tswp::ObjectAttributeTable {
                entries: vec![tswp::object_attribute_table::ObjectAttribute {
                    character_index: 0,
                    object: Some(tsp::Reference {
                        identifier: 7,
                        ..Default::default()
                    }),
                }],
            }),
            ..Default::default()
        };
        let mut data = storage.encode_to_vec();
        data =
            crate::wire::transform_length_delimited_field(&data, LIST_STYLE_TABLE_FIELD, |table| {
                let table = crate::wire::transform_length_delimited_field(
                    table,
                    TABLE_ENTRIES_FIELD,
                    |entry| {
                        let mut entry = entry.to_vec();
                        append_varint_field(&mut entry, 1_901, 101)?;
                        Ok(entry)
                    },
                )?;
                let mut table = table;
                append_varint_field(&mut table, 1_902, 202)?;
                Ok(table)
            })
            .unwrap();
        append_varint_field(&mut data, 1_903, 303).unwrap();

        let mut object = ArchiveObject::new(42, vec![RawMessage { type_: 2_001, data }]).unwrap();
        object.archive_info.message_infos[0].object_references = vec![7];
        let mut package = IWorkPackage::new();
        package
            .replace_archive(
                "Index/Document.iwa",
                &Archive {
                    objects: vec![object],
                },
            )
            .unwrap();

        let located = locate_boundaries_with_archive(&package, 42).unwrap();
        assert_eq!(located.location.boundaries, [(0, 7)]);
        replace_boundaries_with_archive(&mut package, located, 42, &[7], &[(0, 11)]).unwrap();

        let updated = package.archive("Index/Document.iwa").unwrap();
        let raw = &updated.object(42).unwrap().messages[0].data;
        assert_eq!(
            required_varint(raw, 1_903, "root unknown field").unwrap(),
            303
        );
        let table = repeated_length_delimited_payloads(raw, LIST_STYLE_TABLE_FIELD).unwrap();
        let [table] = table.as_slice() else {
            panic!("expected one list-style table");
        };
        assert_eq!(
            required_varint(table, 1_902, "table unknown field").unwrap(),
            202
        );
        let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD).unwrap();
        let [entry] = entries.as_slice() else {
            panic!("expected one list-style entry");
        };
        assert_eq!(
            required_varint(entry, 1_901, "entry unknown field").unwrap(),
            101
        );
        let decoded = tswp::StorageArchive::decode(raw.as_slice()).unwrap();
        assert_eq!(
            decoded.table_list_style.unwrap().entries[0]
                .object
                .unwrap()
                .identifier,
            11
        );
    }
}
