//! Lossless ranged-object table mutation for native hyperlinks and bookmarks.

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
    STORAGE_MESSAGE_TYPES, StorageLocation, locate_storage as locate_native_storage,
    text_utf16_len, validate_sorted_boundaries,
};

pub(super) const SMART_FIELD_TABLE_FIELD: u32 = 11;
const BOOKMARK_TABLE_FIELD: u32 = 15;
pub(super) const TABLE_ENTRIES_FIELD: u32 = 1;
const ENTRY_OBJECT_FIELD: u32 = 2;

#[derive(Debug)]
pub(super) struct Boundary {
    pub(super) index: u32,
    pub(super) object_id: Option<u64>,
    raw: Vec<u8>,
}

#[derive(Debug, Clone, Copy)]
pub(super) enum RangedObjectTable {
    SmartField,
    Bookmark,
}

impl RangedObjectTable {
    const fn field(self) -> u32 {
        match self {
            Self::SmartField => SMART_FIELD_TABLE_FIELD,
            Self::Bookmark => BOOKMARK_TABLE_FIELD,
        }
    }

    const fn label(self) -> &'static str {
        match self {
            Self::SmartField => "smart-field",
            Self::Bookmark => "bookmark",
        }
    }

    fn decoded(self, storage: &tswp::StorageArchive) -> Option<&tswp::ObjectAttributeTable> {
        match self {
            Self::SmartField => storage.table_smartfield.as_ref(),
            Self::Bookmark => storage.table_bookmark.as_ref(),
        }
    }
}

pub(super) fn locate_storage(
    package: &IWorkPackage,
    storage_id: u64,
    kind: RangedObjectTable,
) -> Result<StorageLocation> {
    let location = locate_native_storage(package, storage_id, kind.field(), kind.label())?;
    if kind.decoded(&location.storage).is_some() != location.table_present {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} {} table wire state is inconsistent",
            kind.label()
        )));
    }
    Ok(location)
}

pub(super) fn decoded_boundaries(
    storage_id: u64,
    location: &StorageLocation,
    kind: RangedObjectTable,
) -> Result<Vec<Boundary>> {
    match (location.table_present, kind.decoded(&location.storage)) {
        (false, None) => Ok(Vec::new()),
        (true, Some(table)) if table.entries.is_empty() => Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has an empty {} table",
            kind.label()
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
            validate_boundaries(storage_id, &boundaries, &location.storage.text, kind)?;
            Ok(boundaries)
        },
        _ => Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} {} table wire state is inconsistent",
            kind.label()
        ))),
    }
}

pub(super) fn raw_boundaries(
    storage_id: u64,
    table: Option<&[u8]>,
    storage: &tswp::StorageArchive,
    kind: RangedObjectTable,
) -> Result<Vec<Boundary>> {
    let Some(table) = table else {
        if kind.decoded(storage).is_some() {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} {} table wire state is inconsistent",
                kind.label()
            )));
        }
        return Ok(Vec::new());
    };
    let entries = repeated_length_delimited_payloads(table, TABLE_ENTRIES_FIELD)?;
    if entries.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} has an empty {} table",
            kind.label()
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
    validate_boundaries(storage_id, &boundaries, &storage.text, kind)?;
    Ok(boundaries)
}

fn validate_boundaries(
    storage_id: u64,
    boundaries: &[Boundary],
    text: &[String],
    kind: RangedObjectTable,
) -> Result<()> {
    if boundaries.is_empty() {
        return Ok(());
    }
    if boundaries[0].index != 0 {
        return Err(Error::InvalidFormat(format!(
            "iWork text storage {storage_id} {} table must begin at UTF-16 index zero",
            kind.label()
        )));
    }
    let text_len = text_utf16_len(text)?;
    let mut previous = None;
    for boundary in boundaries {
        if previous.is_some_and(|index| index >= boundary.index) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} {} boundaries are not strictly increasing",
                kind.label()
            )));
        }
        if boundary.object_id == Some(0) {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has a zero {} object identifier",
                kind.label()
            )));
        }
        if boundary.object_id.is_some() && boundary.index == text_len {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} has an empty {} range at its end",
                kind.label()
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
    kind: RangedObjectTable,
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
                "{} range {start}..{end} overlaps object {identifier} in text storage {storage_id}",
                kind.label()
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

pub(super) fn remove_range(
    boundaries: &mut Vec<Boundary>,
    identifier: u64,
    kind: RangedObjectTable,
) -> Result<()> {
    let matches = boundaries
        .iter()
        .enumerate()
        .filter(|(_, boundary)| boundary.object_id == Some(identifier))
        .map(|(position, _)| position)
        .collect::<Vec<_>>();
    let [position] = matches.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "{} table must reference object {identifier} exactly once",
            kind.label()
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
            deprecated_type: None,
            deprecated_is_external: None,
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
            deprecated_type: None,
            deprecated_is_external: None,
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

pub(super) fn patch_ranged_object_table<F>(
    package: &mut IWorkPackage,
    archive_name: &str,
    storage_id: u64,
    kind: RangedObjectTable,
    transform: F,
) -> Result<()>
where
    F: FnOnce(Option<&[u8]>, &tswp::StorageArchive) -> Result<TablePatch>,
{
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
        let tables = repeated_length_delimited_payloads(&original.data, kind.field())?;
        if tables.len() > 1 {
            return Err(Error::InvalidFormat(format!(
                "iWork text storage {storage_id} contains {} {} tables",
                tables.len(),
                kind.label()
            )));
        }
        let (replacement, added, removed) = transform(tables.first().copied(), &storage)?;
        let data = patch_length_delimited_field(
            &original.data,
            kind.field(),
            !tables.is_empty(),
            replacement.as_deref(),
        )?;
        object.replace_message(
            *index,
            RawMessage {
                type_: original.type_,
                data,
            },
        )?;
        let info = &mut object.archive_info.message_infos[*index];
        if let Some(identifier) = added {
            if info.object_references.contains(&identifier) {
                return Err(Error::InvalidFormat(format!(
                    "iWork storage metadata already references new {} object {identifier}",
                    kind.label()
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
                    "iWork storage metadata references {} object {identifier} {count} times",
                    kind.label()
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
