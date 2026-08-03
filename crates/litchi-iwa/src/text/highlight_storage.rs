//! Lossless range-table mutation for native text highlights.

use prost::Message;

use crate::archive::RawMessage;
use crate::protobuf::{tsp, tswp};
use crate::wire::{
    patch_length_delimited_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use crate::{Error, IWorkPackage, Result};

use super::position::TextRange;
use super::storage_wire::{
    LocatedStorage, StorageLocation, locate_storage as locate_native_storage,
    locate_storage_with_archive as locate_native_storage_with_archive, text_utf16_len,
    update_parsed_archive, validate_sorted_boundaries,
};

pub(super) const HIGHLIGHT_TABLE_FIELD: u32 = 23;
pub(super) const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;

#[derive(Debug)]
pub(super) struct Boundary {
    pub(super) index: u32,
    pub(super) object_id: Option<u64>,
    raw: Vec<u8>,
}

pub(super) fn locate_storage(package: &IWorkPackage, storage_id: u64) -> Result<StorageLocation> {
    let location = locate_native_storage(package, storage_id, HIGHLIGHT_TABLE_FIELD, "highlight")?;
    if location.storage.table_highlight.is_some() != location.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} highlight-table wire state is inconsistent"
        )));
    }
    Ok(location)
}

pub(super) fn locate_storage_with_archive(
    package: &IWorkPackage,
    storage_id: u64,
) -> Result<LocatedStorage> {
    let located = locate_native_storage_with_archive(
        package,
        storage_id,
        HIGHLIGHT_TABLE_FIELD,
        "highlight",
    )?;
    if located.location.storage.table_highlight.is_some() != located.location.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} highlight-table wire state is inconsistent"
        )));
    }
    Ok(located)
}

pub(super) fn decoded_boundaries(
    storage_id: u64,
    location: &StorageLocation,
) -> Result<Vec<Boundary>> {
    match (
        location.table_present,
        location.storage.table_highlight.as_ref(),
    ) {
        (false, None) => Ok(Vec::new()),
        (true, Some(table)) if table.entries.is_empty() => Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has an empty highlight table"
        ))),
        (true, Some(table)) => {
            let boundaries = table
                .entries
                .iter()
                .map(|entry| Boundary {
                    index: entry.character_index,
                    object_id: entry.object.as_ref().map(|reference| reference.identifier),
                    raw: Vec::new(),
                })
                .collect::<Vec<_>>();
            validate_boundaries(storage_id, &boundaries, &location.storage.text)?;
            Ok(boundaries)
        },
        _ => Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} highlight-table wire state is inconsistent"
        ))),
    }
}

pub(super) fn raw_boundaries(
    storage_id: u64,
    table: Option<&[u8]>,
    storage: &tswp::StorageArchive,
) -> Result<Vec<Boundary>> {
    let Some(table) = table else {
        if storage.table_highlight.is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} highlight-table wire state is inconsistent"
            )));
        }
        return Ok(Vec::new());
    };
    let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
    if entries.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has an empty highlight table"
        )));
    }
    let boundaries = entries
        .into_iter()
        .map(|raw| {
            let entry = tswp::object_attribute_table::ObjectAttribute::decode(raw)?;
            Ok(Boundary {
                index: entry.character_index,
                object_id: entry.object.map(|reference| reference.identifier),
                raw: raw.to_vec(),
            })
        })
        .collect::<Result<Vec<_>>>()?;
    validate_boundaries(storage_id, &boundaries, &storage.text)?;
    Ok(boundaries)
}

fn validate_boundaries(storage_id: u64, boundaries: &[Boundary], text: &[String]) -> Result<()> {
    if boundaries.is_empty() {
        return Ok(());
    }
    if boundaries[0].index != 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} highlight table must begin at UTF-16 index zero"
        )));
    }
    let text_len = text_utf16_len(text)?;
    let mut previous = None;
    for boundary in boundaries {
        if previous.is_some_and(|index| index >= boundary.index) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} highlight boundaries are not strictly increasing"
            )));
        }
        if boundary.object_id == Some(0) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has a zero highlight-object identifier"
            )));
        }
        if boundary.object_id.is_some() && boundary.index == text_len {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has an empty highlight range at its end"
            )));
        }
        previous = Some(boundary.index);
    }
    validate_sorted_boundaries(
        storage_id,
        boundaries.iter().map(|boundary| boundary.index),
        text,
    )
}

pub(super) fn validate_range(storage_id: u64, range: TextRange, text: &[String]) -> Result<()> {
    validate_sorted_boundaries(
        storage_id,
        [range.start().utf16_index(), range.end().utf16_index()],
        text,
    )
}

pub(super) fn ensure_range_available(
    storage_id: u64,
    range: TextRange,
    boundaries: &[Boundary],
    ignored_id: Option<u64>,
    text: &[String],
) -> Result<()> {
    let text_len = text_utf16_len(text)?;
    let start = range.start().utf16_index();
    let end = range.end().utf16_index();
    for (position, boundary) in boundaries.iter().enumerate() {
        let Some(identifier) = boundary.object_id else {
            continue;
        };
        if Some(identifier) == ignored_id {
            continue;
        }
        let occupied_end = boundaries
            .get(position + 1)
            .map_or(text_len, |next| next.index);
        if start < occupied_end && boundary.index < end {
            return Err(Error::InvalidFormat(format!(
                "highlight range {start}..{end} overlaps highlight object {identifier} in text storage {storage_id}"
            )));
        }
    }
    Ok(())
}

pub(super) fn add_range(
    boundaries: &mut Vec<Boundary>,
    range: TextRange,
    identifier: u64,
) -> Result<()> {
    if boundaries.is_empty() {
        boundaries.push(new_boundary(0, None));
    }
    set_boundary(boundaries, range.start().utf16_index(), Some(identifier))?;
    set_boundary(boundaries, range.end().utf16_index(), None)?;
    coalesce_boundaries(boundaries);
    Ok(())
}

pub(super) fn remove_range(boundaries: &mut Vec<Boundary>, identifier: u64) -> Result<()> {
    let matches = boundaries
        .iter()
        .enumerate()
        .filter(|(_, boundary)| boundary.object_id == Some(identifier))
        .map(|(position, _)| position)
        .collect::<Vec<_>>();
    let [position] = matches.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "highlight table must reference object {identifier} exactly once"
        )));
    };
    let boundary = &mut boundaries[*position];
    boundary.raw = patch_boundary_object(&boundary.raw, boundary.object_id, None)?;
    boundary.object_id = None;
    coalesce_boundaries(boundaries);
    Ok(())
}

fn set_boundary(boundaries: &mut Vec<Boundary>, index: u32, object_id: Option<u64>) -> Result<()> {
    if let Some(boundary) = boundaries
        .iter_mut()
        .find(|boundary| boundary.index == index)
    {
        if boundary.object_id != object_id {
            boundary.raw = patch_boundary_object(&boundary.raw, boundary.object_id, object_id)?;
            boundary.object_id = object_id;
        }
        return Ok(());
    }
    boundaries.push(new_boundary(index, object_id));
    boundaries.sort_by_key(|boundary| boundary.index);
    Ok(())
}

fn new_boundary(index: u32, object_id: Option<u64>) -> Boundary {
    let entry = tswp::object_attribute_table::ObjectAttribute {
        character_index: index,
        object: object_id.map(|identifier| tsp::Reference {
            identifier,
            ..Default::default()
        }),
    };
    Boundary {
        index,
        object_id,
        raw: entry.encode_to_vec(),
    }
}

fn patch_boundary_object(
    raw: &[u8],
    current: Option<u64>,
    replacement: Option<u64>,
) -> Result<Vec<u8>> {
    let reference = replacement.map(|identifier| {
        tsp::Reference {
            identifier,
            ..Default::default()
        }
        .encode_to_vec()
    });
    patch_length_delimited_field(
        raw,
        ENTRY_OBJECT_FIELD,
        current.is_some(),
        reference.as_deref(),
    )
}

fn coalesce_boundaries(boundaries: &mut Vec<Boundary>) {
    boundaries.dedup_by(|current, previous| current.object_id == previous.object_id);
}

pub(super) fn encode_table(original: Option<&[u8]>, boundaries: Vec<Boundary>) -> Result<Vec<u8>> {
    let entries = boundaries
        .into_iter()
        .map(|boundary| boundary.raw)
        .collect::<Vec<_>>();
    match original {
        Some(table) => {
            rewrite_repeated_length_delimited_fields(table, TABLE_ENTRIES_FIELD, &entries)
        },
        None => Ok(tswp::ObjectAttributeTable {
            entries: entries
                .iter()
                .map(|entry| {
                    tswp::object_attribute_table::ObjectAttribute::decode(entry.as_slice())
                })
                .collect::<std::result::Result<Vec<_>, _>>()?,
        }
        .encode_to_vec()),
    }
}

type TablePatch = (Option<Vec<u8>>, Option<u64>, Option<u64>);

pub(super) fn patch_highlight_table<F>(
    package: &mut IWorkPackage,
    located: LocatedStorage,
    transform: F,
) -> Result<()>
where
    F: FnOnce(Option<&[u8]>, &tswp::StorageArchive) -> Result<TablePatch>,
{
    let LocatedStorage { location, archive } = located;
    update_parsed_archive(package, &location.archive_name, archive, |archive| {
        let object = archive.object_mut(location.object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork text storage {} is missing",
                location.object_id
            ))
        })?;
        if object.archive_info.identifier != Some(location.object_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {} has an invalid archive identity",
                location.object_id
            )));
        }
        if object
            .archive_info
            .message_infos
            .get(location.message_index)
            .is_none()
        {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {} is missing metadata for anchored message {}",
                location.object_id, location.message_index
            )));
        }
        let (message_type, data, added, removed) = {
            let original = object.messages.get(location.message_index).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork text storage {} is missing anchored message {}",
                    location.object_id, location.message_index
                ))
            })?;
            if original.type_ != location.message_type {
                return Err(Error::InvalidFormat(format!(
                    "iWork text storage {} anchored message {} changed type from {} to {}",
                    location.object_id,
                    location.message_index,
                    location.message_type,
                    original.type_
                )));
            }
            let tables = repeated_length_delimited_payloads(&original.data, HIGHLIGHT_TABLE_FIELD)?;
            if tables.len() > 1 {
                return Err(Error::InvalidFormat(format!(
                    "iWork text storage {} contains {} highlight tables",
                    location.object_id,
                    tables.len()
                )));
            }
            let (replacement, added, removed) =
                transform(tables.first().copied(), &location.storage)?;
            let data = patch_length_delimited_field(
                &original.data,
                HIGHLIGHT_TABLE_FIELD,
                !tables.is_empty(),
                replacement.as_deref(),
            )?;
            (original.type_, data, added, removed)
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
                    "iWork text storage {} is missing metadata for anchored message {}",
                    location.object_id, location.message_index
                ))
            })?;
        if let Some(identifier) = added {
            if info.object_references.contains(&identifier) {
                return Err(Error::InvalidFormat(format!(
                    "iWork storage metadata already references new highlight object {identifier}"
                )));
            }
            info.object_references.push(identifier);
        }
        if let Some(identifier) = removed {
            let count = info
                .object_references
                .iter()
                .filter(|candidate| **candidate == identifier)
                .count();
            if count != 1 {
                return Err(Error::InvalidFormat(format!(
                    "iWork storage metadata references highlight object {identifier} {count} times"
                )));
            }
            info.object_references
                .retain(|candidate| *candidate != identifier);
            for field in &mut info.field_infos {
                field
                    .object_references
                    .retain(|candidate| *candidate != identifier);
            }
        }
        Ok(())
    })
}
