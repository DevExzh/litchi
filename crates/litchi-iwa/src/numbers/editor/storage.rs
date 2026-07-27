//! Table-data storage, rich-text, dependency, and tile wire updates.

use super::*;

const TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE: u32 = 6011;

#[derive(Debug)]
pub(super) struct RichTextEntryLocation {
    table_id: u64,
    table_archive: String,
    payload_id: u64,
    payload_archive: String,
    storage_id: u64,
    storage_archive: String,
    refcount: u32,
    owner: TableDataListEntryOwner,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) enum TableDataListEntryOwner {
    Root,
    Segment { object_id: u64, archive: String },
}

#[derive(Debug, Clone)]
pub(super) struct LocatedTableDataListEntry {
    pub(super) owner: TableDataListEntryOwner,
    pub(super) entry: tst::table_data_list::ListEntry,
}

#[derive(Debug)]
pub(super) struct ResolvedTableDataList {
    pub(super) table_id: u64,
    pub(super) table_archive: String,
    pub(super) list: TableDataList,
    pub(super) entries: Vec<LocatedTableDataListEntry>,
}

pub(super) fn resolve_table_data_list(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    list_type: tst::table_data_list::ListType,
) -> Result<ResolvedTableDataList> {
    let table_archive = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table-data-list object {table_id} is missing"
        ))
    })?;
    let archive = package.archive(table_archive)?;
    let object = archive.object(table_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table-data-list object {table_id} is missing"
        ))
    })?;
    let message_index = table_data_list_message_index(object, list_type).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Object {table_id} has no Numbers {list_type:?} TableDataList payload"
        ))
    })?;
    let list = TableDataList::decode(object.messages[message_index].data.as_slice())?;
    let mut entries = list
        .entries
        .iter()
        .cloned()
        .map(|entry| LocatedTableDataListEntry {
            owner: TableDataListEntryOwner::Root,
            entry,
        })
        .collect::<Vec<_>>();
    let mut keys = entries
        .iter()
        .map(|located| located.entry.key)
        .collect::<HashSet<_>>();
    if keys.len() != entries.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers {list_type:?} table {table_id} contains duplicate root entry keys"
        )));
    }
    let mut segment_ids = HashSet::new();
    for reference in &list.segments {
        if !segment_ids.insert(reference.identifier) {
            return Err(Error::InvalidFormat(format!(
                "Numbers {list_type:?} table {table_id} repeats segment object {}",
                reference.identifier
            )));
        }
        let segment_archive = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table-data-list segment object {} is missing",
                reference.identifier
            ))
        })?;
        let archive = package.archive(segment_archive)?;
        let segment_object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table-data-list segment object {} is missing",
                reference.identifier
            ))
        })?;
        let mut segment_messages = segment_object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE);
        let Some(segment_message) = segment_messages.next() else {
            return Err(Error::InvalidFormat(format!(
                "Object {} has no Numbers TableDataListSegment payload",
                reference.identifier
            )));
        };
        if segment_messages.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "Object {} has multiple Numbers TableDataListSegment payloads",
                reference.identifier
            )));
        }
        let segment = TableDataListSegment::decode(segment_message.data.as_slice())?;
        validate_table_data_list_segment(reference.identifier, list_type, &segment)?;
        let owner = TableDataListEntryOwner::Segment {
            object_id: reference.identifier,
            archive: segment_archive.clone(),
        };
        for entry in segment.entries {
            if !keys.insert(entry.key) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers {list_type:?} table {table_id} repeats entry key {} across root and segments",
                    entry.key
                )));
            }
            entries.push(LocatedTableDataListEntry {
                owner: owner.clone(),
                entry,
            });
        }
    }
    Ok(ResolvedTableDataList {
        table_id,
        table_archive: table_archive.clone(),
        list,
        entries,
    })
}

/// Resolve only the requested plain-string entries without cloning an entire
/// potentially large table string list.
pub(super) fn resolve_table_string_values(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    requested_keys: &HashSet<u32>,
) -> Result<HashMap<u32, String>> {
    if requested_keys.is_empty() {
        return Ok(HashMap::new());
    }
    let table_archive = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table-data-list object {table_id} is missing"
        ))
    })?;
    let archive = package.archive(table_archive)?;
    let object = archive.object(table_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table-data-list object {table_id} is missing"
        ))
    })?;
    let list_type = tst::table_data_list::ListType::String;
    let message_index = table_data_list_message_index(object, list_type).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Object {table_id} has no Numbers {list_type:?} TableDataList payload"
        ))
    })?;
    let list = TableDataList::decode(object.messages[message_index].data.as_slice())?;
    let TableDataList {
        entries, segments, ..
    } = list;
    let mut values = HashMap::with_capacity(requested_keys.len());
    let mut keys = HashSet::with_capacity(entries.len());
    for entry in entries {
        let key = entry.key;
        if !keys.insert(key) {
            return Err(Error::InvalidFormat(format!(
                "Numbers String table {table_id} contains duplicate root entry keys"
            )));
        }
        if requested_keys.contains(&key)
            && let Some(value) = entry.string
        {
            values.insert(key, value);
        }
    }

    let mut segment_ids = HashSet::with_capacity(segments.len());
    for reference in segments {
        if !segment_ids.insert(reference.identifier) {
            return Err(Error::InvalidFormat(format!(
                "Numbers String table {table_id} repeats segment object {}",
                reference.identifier
            )));
        }
        let segment_archive = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table-data-list segment object {} is missing",
                reference.identifier
            ))
        })?;
        let archive = package.archive(segment_archive)?;
        let segment_object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table-data-list segment object {} is missing",
                reference.identifier
            ))
        })?;
        let mut segment_messages = segment_object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE);
        let Some(segment_message) = segment_messages.next() else {
            return Err(Error::InvalidFormat(format!(
                "Object {} has no Numbers TableDataListSegment payload",
                reference.identifier
            )));
        };
        if segment_messages.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "Object {} has multiple Numbers TableDataListSegment payloads",
                reference.identifier
            )));
        }
        let segment = TableDataListSegment::decode(segment_message.data.as_slice())?;
        validate_table_data_list_segment(reference.identifier, list_type, &segment)?;
        for entry in segment.entries {
            let key = entry.key;
            if !keys.insert(key) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers String table {table_id} repeats entry key {key} across root and segments"
                )));
            }
            if requested_keys.contains(&key)
                && let Some(value) = entry.string
            {
                values.insert(key, value);
            }
        }
    }
    Ok(values)
}

/// Return whether a typed table-data list contains any root or segmented entry.
///
/// This avoids cloning payloads when a caller only needs to distinguish an
/// empty native allocation from storage that carries semantic content.
pub(super) fn table_data_list_has_entries(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    list_type: tst::table_data_list::ListType,
) -> Result<bool> {
    let table_archive = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table-data-list object {table_id} is missing"
        ))
    })?;
    let archive = package.archive(table_archive)?;
    let object = archive.object(table_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table-data-list object {table_id} is missing"
        ))
    })?;
    let message_index = table_data_list_message_index(object, list_type).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Object {table_id} has no Numbers {list_type:?} TableDataList payload"
        ))
    })?;
    let list = TableDataList::decode(object.messages[message_index].data.as_slice())?;
    if !list.entries.is_empty() {
        return Ok(true);
    }

    let mut segment_ids = HashSet::with_capacity(list.segments.len());
    for reference in list.segments {
        if !segment_ids.insert(reference.identifier) {
            return Err(Error::InvalidFormat(format!(
                "Numbers {list_type:?} table {table_id} repeats segment object {}",
                reference.identifier
            )));
        }
        let segment_archive = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table-data-list segment object {} is missing",
                reference.identifier
            ))
        })?;
        let archive = package.archive(segment_archive)?;
        let segment_object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table-data-list segment object {} is missing",
                reference.identifier
            ))
        })?;
        let mut segment_messages = segment_object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE);
        let Some(segment_message) = segment_messages.next() else {
            return Err(Error::InvalidFormat(format!(
                "Object {} has no Numbers TableDataListSegment payload",
                reference.identifier
            )));
        };
        if segment_messages.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "Object {} has multiple Numbers TableDataListSegment payloads",
                reference.identifier
            )));
        }
        let segment = TableDataListSegment::decode(segment_message.data.as_slice())?;
        validate_table_data_list_segment(reference.identifier, list_type, &segment)?;
        if !segment.entries.is_empty() {
            return Ok(true);
        }
    }
    Ok(false)
}

pub(super) fn validate_table_data_list_segment(
    object_id: u64,
    list_type: tst::table_data_list::ListType,
    segment: &TableDataListSegment,
) -> Result<()> {
    if segment.list_type != list_type as i32 {
        return Err(Error::InvalidFormat(format!(
            "Numbers table-data-list segment {object_id} has list type {}, expected {list_type:?}",
            segment.list_type
        )));
    }
    let end = segment
        .key_range
        .location
        .checked_add(segment.key_range.length)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table-data-list segment {object_id} key range overflows"
            ))
        })?;
    if segment
        .entries
        .iter()
        .any(|entry| entry.key < segment.key_range.location || entry.key >= end)
    {
        return Err(Error::InvalidFormat(format!(
            "Numbers table-data-list segment {object_id} contains an entry outside its key range"
        )));
    }
    Ok(())
}

pub(super) fn rewrite_table_data_list_entries(
    data: &[u8],
    field_number: u32,
    previous: &[tst::table_data_list::ListEntry],
    current: &[tst::table_data_list::ListEntry],
) -> Result<Vec<u8>> {
    let raw_entries = repeated_length_delimited_payloads(data, field_number)?;
    if raw_entries.len() != previous.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers table-data-list field {field_number} has {} raw entries but {} decoded entries",
            raw_entries.len(),
            previous.len()
        )));
    }
    let mut existing = HashMap::with_capacity(previous.len());
    for (expected, raw) in previous.iter().zip(raw_entries) {
        if tst::table_data_list::ListEntry::decode(raw)? != *expected {
            return Err(Error::InvalidFormat(format!(
                "Numbers table-data-list entry {} changed during wire mutation",
                expected.key
            )));
        }
        patch_varint_field(raw, 1, true, Some(u64::from(expected.key)))?;
        patch_varint_field(raw, 2, true, Some(u64::from(expected.refcount)))?;
        if existing.insert(expected.key, (expected, raw)).is_some() {
            return Err(Error::InvalidFormat(format!(
                "Numbers table-data-list contains duplicate key {}",
                expected.key
            )));
        }
    }

    let mut seen = HashSet::with_capacity(current.len());
    let replacements = current
        .iter()
        .map(|entry| {
            if !seen.insert(entry.key) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table-data-list would contain duplicate key {}",
                    entry.key
                )));
            }
            let Some((previous, raw)) = existing.get(&entry.key) else {
                return Ok(entry.encode_to_vec());
            };
            let mut expected = entry.clone();
            expected.refcount = previous.refcount;
            if expected != **previous {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table-data-list entry {} changed outside its refcount",
                    entry.key
                )));
            }
            patch_varint_field(raw, 2, true, Some(u64::from(entry.refcount)))
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, field_number, &replacements)
}

pub(super) fn rewrite_table_data_list_wire(
    original: &[u8],
    previous: &TableDataList,
    current: &TableDataList,
) -> Result<Vec<u8>> {
    if previous.list_type != current.list_type || previous.is_new_for_bnc != current.is_new_for_bnc
    {
        return Err(Error::InvalidFormat(
            "Numbers table-data-list immutable fields changed during mutation".to_owned(),
        ));
    }
    let mut data = patch_varint_field(original, 2, true, Some(u64::from(current.next_list_id)))?;
    data = rewrite_table_data_list_entries(&data, 3, &previous.entries, &current.entries)?;
    data = rewrite_reference_list(
        &data,
        4,
        &previous
            .segments
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
        &current
            .segments
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>(),
    )?;
    if TableDataList::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers TableDataList wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn rewrite_table_data_list_segment_wire(
    original: &[u8],
    previous: &TableDataListSegment,
    current: &TableDataListSegment,
) -> Result<Vec<u8>> {
    if previous.list_type != current.list_type {
        return Err(Error::InvalidFormat(
            "Numbers table-data-list segment type changed during mutation".to_owned(),
        ));
    }
    let mut data = patch_nested_varint_field(
        original,
        &[2, 1],
        true,
        Some(u64::from(current.key_range.location)),
    )?;
    data = patch_nested_varint_field(
        &data,
        &[2, 2],
        true,
        Some(u64::from(current.key_range.length)),
    )?;
    data = rewrite_table_data_list_entries(&data, 3, &previous.entries, &current.entries)?;
    if TableDataListSegment::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers TableDataListSegment wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn rewrite_tile_row_infos_wire(
    data: &[u8],
    previous: &[TileRowInfo],
    current: &[TileRowInfo],
) -> Result<Vec<u8>> {
    let raw_rows = repeated_length_delimited_payloads(data, 5)?;
    if raw_rows.len() != previous.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers tile has {} raw rows but {} decoded rows",
            raw_rows.len(),
            previous.len()
        )));
    }
    let mut existing = HashMap::with_capacity(previous.len());
    for (expected, raw) in previous.iter().zip(raw_rows) {
        if TileRowInfo::decode(raw)? != *expected {
            return Err(Error::InvalidFormat(format!(
                "Numbers tile row {} changed during wire mutation",
                expected.tile_row_index
            )));
        }
        patch_varint_field(raw, 1, true, Some(u64::from(expected.tile_row_index)))?;
        if existing
            .insert(expected.tile_row_index, (expected, raw))
            .is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers tile contains duplicate row {}",
                expected.tile_row_index
            )));
        }
    }

    let mut seen = HashSet::with_capacity(current.len());
    let replacements = current
        .iter()
        .map(|row| {
            if !seen.insert(row.tile_row_index) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers tile would contain duplicate row {}",
                    row.tile_row_index
                )));
            }
            let Some((previous, raw)) = existing.get(&row.tile_row_index) else {
                return Ok(row.encode_to_vec());
            };
            let mut data = patch_varint_field(raw, 2, true, Some(u64::from(row.cell_count)))?;
            data = patch_length_delimited_field(
                &data,
                3,
                true,
                Some(&row.cell_storage_buffer_pre_bnc),
            )?;
            data = patch_length_delimited_field(&data, 4, true, Some(&row.cell_offsets_pre_bnc))?;
            data = patch_varint_field(
                &data,
                5,
                previous.storage_version.is_some(),
                row.storage_version.map(u64::from),
            )?;
            data = patch_length_delimited_field(
                &data,
                6,
                previous.cell_storage_buffer.is_some(),
                row.cell_storage_buffer.as_deref(),
            )?;
            data = patch_length_delimited_field(
                &data,
                7,
                previous.cell_offsets.is_some(),
                row.cell_offsets.as_deref(),
            )?;
            data = patch_varint_field(
                &data,
                8,
                previous.has_wide_offsets.is_some(),
                row.has_wide_offsets.map(u64::from),
            )?;
            Ok(data)
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, 5, &replacements)
}

pub(super) fn rewrite_tile_wire(
    original: &[u8],
    previous: &Tile,
    current: &Tile,
) -> Result<Vec<u8>> {
    if previous.max_column != current.max_column
        || previous.max_row != current.max_row
        || previous.num_cells != current.num_cells
        || previous.should_use_wide_rows != current.should_use_wide_rows
    {
        return Err(Error::InvalidFormat(
            "Numbers tile immutable fields changed during mutation".to_owned(),
        ));
    }
    let mut data = patch_varint_field(original, 4, true, Some(u64::from(current.numrows)))?;
    data = rewrite_tile_row_infos_wire(&data, &previous.row_infos, &current.row_infos)?;
    data = patch_varint_field(
        &data,
        6,
        previous.storage_version.is_some(),
        current.storage_version.map(u64::from),
    )?;
    data = patch_varint_field(
        &data,
        7,
        previous.last_saved_in_bnc.is_some(),
        current.last_saved_in_bnc.map(u64::from),
    )?;
    if Tile::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers tile wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn rewrite_header_bucket_wire(
    original: &[u8],
    previous: &tst::HeaderStorageBucket,
    current: &tst::HeaderStorageBucket,
) -> Result<Vec<u8>> {
    if previous.bucket_hash_function != current.bucket_hash_function {
        return Err(Error::InvalidFormat(
            "Numbers header-bucket hash function changed during mutation".to_owned(),
        ));
    }
    let raw_headers = repeated_length_delimited_payloads(original, 2)?;
    if raw_headers.len() != previous.headers.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers header bucket has {} raw entries but {} decoded entries",
            raw_headers.len(),
            previous.headers.len()
        )));
    }
    let mut existing = HashMap::with_capacity(previous.headers.len());
    for (expected, raw) in previous.headers.iter().zip(raw_headers) {
        if tst::header_storage_bucket::Header::decode(raw)? != *expected {
            return Err(Error::InvalidFormat(format!(
                "Numbers header {} changed during wire mutation",
                expected.index
            )));
        }
        patch_varint_field(raw, 1, true, Some(u64::from(expected.index)))?;
        if existing.insert(expected.index, (expected, raw)).is_some() {
            return Err(Error::InvalidFormat(format!(
                "Numbers header bucket contains duplicate index {}",
                expected.index
            )));
        }
    }
    let mut seen = HashSet::with_capacity(current.headers.len());
    let replacements = current
        .headers
        .iter()
        .map(|header| {
            if !seen.insert(header.index) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers header bucket would contain duplicate index {}",
                    header.index
                )));
            }
            let Some((previous, raw)) = existing.get(&header.index) else {
                return Ok(header.encode_to_vec());
            };
            let mut expected = *header;
            expected.number_of_cells = previous.number_of_cells;
            if expected != **previous {
                return Err(Error::InvalidFormat(format!(
                    "Numbers header {} changed outside its cell count",
                    header.index
                )));
            }
            patch_varint_field(raw, 4, true, Some(u64::from(header.number_of_cells)))
        })
        .collect::<Result<Vec<_>>>()?;
    let data = rewrite_repeated_length_delimited_fields(original, 2, &replacements)?;
    if tst::HeaderStorageBucket::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers header-bucket wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn rewrite_uuid_list_wire(
    data: &[u8],
    field_number: u32,
    previous: &[crate::protobuf::tsp::Uuid],
    current: &[crate::protobuf::tsp::Uuid],
) -> Result<Vec<u8>> {
    let raw_values = repeated_length_delimited_payloads(data, field_number)?;
    if raw_values.len() != previous.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers UID-map field {field_number} has {} raw UUIDs but {} decoded UUIDs",
            raw_values.len(),
            previous.len()
        )));
    }
    let mut existing = HashMap::with_capacity(previous.len());
    for (expected, raw) in previous.iter().zip(raw_values) {
        if crate::protobuf::tsp::Uuid::decode(raw)? != *expected {
            return Err(Error::InvalidFormat(format!(
                "Numbers UID-map field {field_number} changed during wire mutation"
            )));
        }
        let key = (expected.lower, expected.upper);
        if existing.insert(key, raw).is_some() {
            return Err(Error::InvalidFormat(format!(
                "Numbers UID-map field {field_number} contains duplicate UUIDs"
            )));
        }
    }
    let mut seen = HashSet::with_capacity(current.len());
    let replacements = current
        .iter()
        .map(|uuid| {
            let key = (uuid.lower, uuid.upper);
            if !seen.insert(key) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers UID-map field {field_number} would contain duplicate UUIDs"
                )));
            }
            Ok(existing
                .get(&key)
                .map_or_else(|| uuid.encode_to_vec(), |raw| raw.to_vec()))
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, field_number, &replacements)
}

pub(super) fn rewrite_uid_index_list_wire(
    data: &[u8],
    field_number: u32,
    previous: &[u32],
    current: &[u32],
) -> Result<Vec<u8>> {
    let raw = repeated_varint_values(data, field_number)?;
    if raw != previous.iter().copied().map(u64::from).collect::<Vec<_>>() {
        return Err(Error::InvalidFormat(format!(
            "Numbers UID-map field {field_number} changed during wire mutation"
        )));
    }
    rewrite_repeated_varint_fields(
        data,
        field_number,
        &current.iter().copied().map(u64::from).collect::<Vec<_>>(),
    )
}

pub(super) fn rewrite_uid_map_wire(
    original: &[u8],
    previous: &tst::ColumnRowUidMapArchive,
    current: &tst::ColumnRowUidMapArchive,
) -> Result<Vec<u8>> {
    let mut data = rewrite_uuid_list_wire(
        original,
        1,
        &previous.sorted_column_uids,
        &current.sorted_column_uids,
    )?;
    data = rewrite_uid_index_list_wire(
        &data,
        2,
        &previous.column_index_for_uid,
        &current.column_index_for_uid,
    )?;
    data = rewrite_uid_index_list_wire(
        &data,
        3,
        &previous.column_uid_for_index,
        &current.column_uid_for_index,
    )?;
    data = rewrite_uuid_list_wire(
        &data,
        4,
        &previous.sorted_row_uids,
        &current.sorted_row_uids,
    )?;
    data = rewrite_uid_index_list_wire(
        &data,
        5,
        &previous.row_index_for_uid,
        &current.row_index_for_uid,
    )?;
    data = rewrite_uid_index_list_wire(
        &data,
        6,
        &previous.row_uid_for_index,
        &current.row_uid_for_index,
    )?;
    if tst::ColumnRowUidMapArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers UID-map wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn rewrite_stroke_sidecar_wire(
    original: &[u8],
    previous: &tst::StrokeSidecarArchive,
    current: &tst::StrokeSidecarArchive,
) -> Result<Vec<u8>> {
    let mut expected = current.clone();
    expected.column_count = previous.column_count;
    expected.row_count = previous.row_count;
    if expected != *previous {
        return Err(Error::InvalidFormat(
            "Numbers stroke sidecar changed outside its dimensions".to_owned(),
        ));
    }
    let mut data = patch_varint_field(
        original,
        2,
        previous.column_count.is_some(),
        current.column_count.map(u64::from),
    )?;
    data = patch_varint_field(
        &data,
        3,
        previous.row_count.is_some(),
        current.row_count.map(u64::from),
    )?;
    if tst::StrokeSidecarArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers stroke-sidecar wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn rewrite_table_model_comment_table_wire(
    original: &[u8],
    previous: &TableModelArchive,
    current: &TableModelArchive,
) -> Result<Vec<u8>> {
    let mut expected = current.clone();
    expected.base_data_store.comment_storage_table = previous.base_data_store.comment_storage_table;
    if expected != *previous {
        return Err(Error::InvalidFormat(
            "Numbers table model changed outside its comment-table reference".to_owned(),
        ));
    }
    let replacement = current
        .base_data_store
        .comment_storage_table
        .as_ref()
        .map(Message::encode_to_vec);
    let data = patch_nested_length_delimited_field(
        original,
        &[4, 19],
        previous.base_data_store.comment_storage_table.is_some(),
        replacement.as_deref(),
    )?;
    if TableModelArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers table-model comment-table wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn rewrite_table_model_conditional_style_wire(
    original: &[u8],
    previous: &TableModelArchive,
    current: &TableModelArchive,
) -> Result<Vec<u8>> {
    let mut expected = current.clone();
    expected.base_data_store.conditionalstyletable = previous.base_data_store.conditionalstyletable;
    expected.conditional_style_formula_owner_id =
        previous.conditional_style_formula_owner_id.clone();
    if expected != *previous {
        return Err(Error::InvalidFormat(
            "Numbers table model changed outside its conditional-style references".to_owned(),
        ));
    }
    let replacement = current
        .base_data_store
        .conditionalstyletable
        .as_ref()
        .map(Message::encode_to_vec);
    let mut data = patch_nested_length_delimited_field(
        original,
        &[4, 18],
        previous.base_data_store.conditionalstyletable.is_some(),
        replacement.as_deref(),
    )?;
    let owner = current
        .conditional_style_formula_owner_id
        .as_ref()
        .map(Message::encode_to_vec);
    data = patch_length_delimited_field(
        &data,
        39,
        previous.conditional_style_formula_owner_id.is_some(),
        owner.as_deref(),
    )?;
    if TableModelArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers table-model conditional-style wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn rewrite_table_model_format_table_wire(
    original: &[u8],
    previous: &TableModelArchive,
    current: &TableModelArchive,
) -> Result<Vec<u8>> {
    let mut expected = current.clone();
    expected.base_data_store.format_table = previous.base_data_store.format_table;
    if expected != *previous {
        return Err(Error::InvalidFormat(
            "Numbers table model changed outside its format-table reference".to_owned(),
        ));
    }
    let replacement = current
        .base_data_store
        .format_table
        .as_ref()
        .map(Message::encode_to_vec);
    let data = patch_nested_length_delimited_field(
        original,
        &[4, 22],
        previous.base_data_store.format_table.is_some(),
        replacement.as_deref(),
    )?;
    if TableModelArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers table-model format-table wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn rewrite_table_model_control_cell_spec_table_wire(
    original: &[u8],
    previous: &TableModelArchive,
    current: &TableModelArchive,
) -> Result<Vec<u8>> {
    let mut expected = current.clone();
    expected.base_data_store.control_cell_spec_table =
        previous.base_data_store.control_cell_spec_table;
    if expected != *previous {
        return Err(Error::InvalidFormat(
            "Numbers table model changed outside its control-cell-spec-table reference".to_owned(),
        ));
    }
    let replacement = current
        .base_data_store
        .control_cell_spec_table
        .as_ref()
        .map(Message::encode_to_vec);
    let data = patch_nested_length_delimited_field(
        original,
        &[4, 21],
        previous.base_data_store.control_cell_spec_table.is_some(),
        replacement.as_deref(),
    )?;
    if TableModelArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers table-model control-cell-spec-table wire mutation failed validation"
                .to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn segment_key_range(
    entries: &[tst::table_data_list::ListEntry],
) -> Result<crate::protobuf::tsp::Range> {
    let first = entries.iter().map(|entry| entry.key).min().ok_or_else(|| {
        Error::InvalidFormat("Cannot compute the key range of an empty Numbers segment".to_owned())
    })?;
    let last = entries
        .iter()
        .map(|entry| entry.key)
        .max()
        .expect("nonempty");
    Ok(crate::protobuf::tsp::Range {
        location: first,
        length: last
            .checked_sub(first)
            .and_then(|length| length.checked_add(1))
            .ok_or_else(|| Error::ParseError("Numbers segment key range overflow".to_owned()))?,
    })
}

pub(super) fn entry_object_references(entry: &tst::table_data_list::ListEntry) -> HashSet<u64> {
    [
        entry.reference.as_ref(),
        entry.rich_text_payload.as_ref(),
        entry.comment_storage.as_ref(),
    ]
    .into_iter()
    .flatten()
    .map(|reference| reference.identifier)
    .chain(
        entry
            .cell_spec
            .as_ref()
            .and_then(|spec| spec.chooser_control_popup_model.as_ref())
            .map(|reference| reference.identifier),
    )
    .collect()
}

pub(super) fn rewind_table_data_list_next_id(
    package: &mut IWorkPackage,
    resolved: &ResolvedTableDataList,
    list_type: tst::table_data_list::ListType,
    removed_key: u32,
) -> Result<()> {
    package.update_archive(&resolved.table_archive, |archive| {
        let object = archive.object_mut(resolved.table_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table-data-list object {} is missing",
                resolved.table_id
            ))
        })?;
        let message_index = table_data_list_message_index(object, list_type).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {} has no Numbers {list_type:?} TableDataList payload",
                resolved.table_id
            ))
        })?;
        let original = object.messages[message_index].data.as_slice();
        let previous = TableDataList::decode(original)?;
        if removed_key.checked_add(1) != Some(previous.next_list_id) {
            return Ok(());
        }
        let mut current = previous.clone();
        current.next_list_id = removed_key;
        let data = rewrite_table_data_list_wire(original, &previous, &current)?;
        let message_type = object.messages[message_index].type_;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn mutate_table_data_list_entry<F>(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    resolved: &ResolvedTableDataList,
    owner: &TableDataListEntryOwner,
    list_type: tst::table_data_list::ListType,
    key: u32,
    mutate: F,
) -> Result<()>
where
    F: FnOnce(&mut tst::table_data_list::ListEntry) -> Result<bool>,
{
    match owner {
        TableDataListEntryOwner::Root => {
            package.update_archive(&resolved.table_archive, |archive| {
                let object = archive.object_mut(resolved.table_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers table-data-list object {} is missing",
                        resolved.table_id
                    ))
                })?;
                let message_index =
                    table_data_list_message_index(object, list_type).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Object {} has no Numbers {list_type:?} TableDataList payload",
                            resolved.table_id
                        ))
                    })?;
                let message_type = object.messages[message_index].type_;
                let original = object.messages[message_index].data.as_slice();
                let previous = TableDataList::decode(original)?;
                let mut list = previous.clone();
                let position = list
                    .entries
                    .iter()
                    .position(|entry| entry.key == key)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!("Numbers table-data-list has no entry {key}"))
                    })?;
                let old_references = entry_object_references(&list.entries[position]);
                let remove = mutate(&mut list.entries[position])?;
                if remove {
                    list.entries.remove(position);
                    if key.checked_add(1) == Some(list.next_list_id) {
                        list.next_list_id = key;
                    }
                }
                let remaining_references = list
                    .entries
                    .iter()
                    .flat_map(entry_object_references)
                    .collect::<HashSet<_>>();
                let data = rewrite_table_data_list_wire(original, &previous, &list)?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
                for reference in old_references.difference(&remaining_references) {
                    remove_message_object_reference(object, message_index, *reference);
                }
                Ok(())
            })
        },
        TableDataListEntryOwner::Segment { object_id, archive } => {
            let mut became_empty = false;
            let mut removed_entry = false;
            package.update_archive(archive, |archive_data| {
                let object = archive_data.object_mut(*object_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers table-data-list segment object {object_id} is missing"
                    ))
                })?;
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| message.type_ == TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Object {object_id} has no Numbers TableDataListSegment payload"
                        ))
                    })?;
                let original = object.messages[message_index].data.as_slice();
                let previous = TableDataListSegment::decode(original)?;
                let mut segment = previous.clone();
                if segment.list_type != list_type as i32 {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers segment {object_id} has list type {}",
                        segment.list_type
                    )));
                }
                let position = segment
                    .entries
                    .iter()
                    .position(|entry| entry.key == key)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers segment {object_id} has no entry {key}"
                        ))
                    })?;
                let old_references = entry_object_references(&segment.entries[position]);
                let remove = mutate(&mut segment.entries[position])?;
                if remove {
                    segment.entries.remove(position);
                    removed_entry = true;
                }
                became_empty = segment.entries.is_empty();
                if !became_empty {
                    segment.key_range = segment_key_range(&segment.entries)?;
                    let remaining_references = segment
                        .entries
                        .iter()
                        .flat_map(entry_object_references)
                        .collect::<HashSet<_>>();
                    let data = rewrite_table_data_list_segment_wire(original, &previous, &segment)?;
                    object.replace_message(
                        message_index,
                        RawMessage {
                            type_: TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE,
                            data,
                        },
                    )?;
                    for reference in old_references.difference(&remaining_references) {
                        remove_message_object_reference(object, message_index, *reference);
                    }
                }
                Ok(())
            })?;
            if removed_entry {
                rewind_table_data_list_next_id(package, resolved, list_type, key)?;
            }
            if became_empty {
                package.update_archive(&resolved.table_archive, |archive_data| {
                    let object = archive_data.object_mut(resolved.table_id).ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers table-data-list object {} is missing",
                            resolved.table_id
                        ))
                    })?;
                    let message_index = table_data_list_message_index(object, list_type)
                        .ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "Object {} has no Numbers {list_type:?} TableDataList payload",
                                resolved.table_id
                            ))
                        })?;
                    let message_type = object.messages[message_index].type_;
                    let original = object.messages[message_index].data.as_slice();
                    let previous = TableDataList::decode(original)?;
                    let mut list = previous.clone();
                    list.segments
                        .retain(|reference| reference.identifier != *object_id);
                    let data = rewrite_table_data_list_wire(original, &previous, &list)?;
                    object.replace_message(
                        message_index,
                        RawMessage {
                            type_: message_type,
                            data,
                        },
                    )?;
                    remove_message_object_reference(object, message_index, *object_id);
                    Ok(())
                })?;
                remove_object_or_empty_entry(package, locations, *object_id)?;
            }
            Ok(())
        },
    }
}

pub(super) fn increment_table_data_list_entry(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    resolved: &ResolvedTableDataList,
    located: &LocatedTableDataListEntry,
    list_type: tst::table_data_list::ListType,
) -> Result<()> {
    mutate_table_data_list_entry(
        package,
        locations,
        resolved,
        &located.owner,
        list_type,
        located.entry.key,
        |entry| {
            entry.refcount = entry.refcount.checked_add(1).ok_or_else(|| {
                Error::ParseError("Numbers table-data-list reference count overflow".to_owned())
            })?;
            Ok(false)
        },
    )
}

pub(super) fn decrement_table_data_list_entry(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    resolved: &ResolvedTableDataList,
    located: &LocatedTableDataListEntry,
    list_type: tst::table_data_list::ListType,
) -> Result<bool> {
    let removed = located.entry.refcount <= 1;
    mutate_table_data_list_entry(
        package,
        locations,
        resolved,
        &located.owner,
        list_type,
        located.entry.key,
        |entry| {
            if entry.refcount == 0 {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table-data-list entry {} has a zero reference count",
                    entry.key
                )));
            }
            if entry.refcount > 1 {
                entry.refcount -= 1;
                Ok(false)
            } else {
                Ok(true)
            }
        },
    )?;
    Ok(removed)
}

pub(super) fn rich_text_entry_location(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    identifier: u32,
) -> Result<RichTextEntryLocation> {
    let table_id = model
        .base_data_store
        .rich_text_table
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers cell references rich-text entry {identifier}, but table {:?} has no rich-text list",
                model.table_name
            ))
        })?;
    let resolved = resolve_table_data_list(
        package,
        locations,
        table_id,
        tst::table_data_list::ListType::RichTextPayload,
    )?;
    let located = resolved
        .entries
        .iter()
        .find(|located| located.entry.key == identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers rich-text table has no entry {identifier}"))
        })?;
    let entry = &located.entry;
    if entry.refcount == 0 {
        return Err(Error::InvalidFormat(format!(
            "Numbers rich-text entry {identifier} has a zero reference count"
        )));
    }
    let payload_id = entry
        .rich_text_payload
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers rich-text entry {identifier} has no payload reference"
            ))
        })?;
    let payload_archive = locations.get(&payload_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers rich-text payload object {payload_id} is missing"
        ))
    })?;
    let archive = package.archive(payload_archive)?;
    let payload_object = archive.object(payload_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers rich-text payload object {payload_id} is missing"
        ))
    })?;
    let storage_id = payload_object
        .messages
        .iter()
        .find_map(|message| {
            tst::RichTextPayloadArchive::decode(message.data.as_slice())
                .ok()
                .map(|payload| payload.storage.identifier)
        })
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers rich-text payload object {payload_id} has no payload archive"
            ))
        })?;
    let storage_archive = locations.get(&storage_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers rich-text storage object {storage_id} is missing"
        ))
    })?;
    Ok(RichTextEntryLocation {
        table_id,
        table_archive: resolved.table_archive.clone(),
        payload_id,
        payload_archive: payload_archive.clone(),
        storage_id,
        storage_archive: storage_archive.clone(),
        refcount: entry.refcount,
        owner: located.owner.clone(),
    })
}

pub(super) fn set_rich_text(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    identifier: u32,
    row: usize,
    column: usize,
    replacement: &str,
) -> Result<u32> {
    let entry = rich_text_entry_location(package, locations, model, identifier)?;
    if entry.refcount == 1 {
        let mut text = IWorkTextEditor::from_package(package.clone());
        text.set_text(entry.storage_id, replacement)?;
        *package = text.into_package();
        return Ok(identifier);
    }

    let mut next_identifier = locations
        .keys()
        .copied()
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
    let new_storage_id = take_identifier(&mut next_identifier)?;
    let new_payload_id = take_identifier(&mut next_identifier)?;

    let source_storage_archive = package.archive(&entry.storage_archive)?;
    let source_storage = source_storage_archive
        .object(entry.storage_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers rich-text storage object {} is missing",
                entry.storage_id
            ))
        })?;
    if source_storage.messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "Cannot safely clone multi-payload Numbers rich-text storage object {}",
            entry.storage_id
        )));
    }
    tswp::StorageArchive::decode(source_storage.messages[0].data.as_slice()).map_err(|_| {
        Error::InvalidFormat(format!(
            "Numbers rich-text storage object {} has no TSWP storage payload",
            entry.storage_id
        ))
    })?;
    let storage_references = source_storage.archive_info.message_infos[0]
        .object_references
        .clone();
    let cloned_storage = clone_single_payload_object(
        source_storage,
        new_storage_id,
        0,
        source_storage.messages[0].data.clone(),
        storage_references,
        &HashMap::new(),
        false,
    )?;

    let source_payload_archive = package.archive(&entry.payload_archive)?;
    let source_payload = source_payload_archive
        .object(entry.payload_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers rich-text payload object {} is missing",
                entry.payload_id
            ))
        })?;
    if source_payload.messages.len() != 1 {
        return Err(Error::InvalidFormat(format!(
            "Cannot safely clone multi-payload Numbers rich-text payload object {}",
            entry.payload_id
        )));
    }
    let mut payload =
        tst::RichTextPayloadArchive::decode(source_payload.messages[0].data.as_slice())?;
    payload.storage.identifier = new_storage_id;
    payload.cellid = rich_text_cell_id(row, column)?;
    let remap = HashMap::from([(entry.storage_id, new_storage_id)]);
    let payload_references = source_payload.archive_info.message_infos[0]
        .object_references
        .iter()
        .map(|reference| remap.get(reference).copied().unwrap_or(*reference))
        .collect();
    let cloned_payload = clone_single_payload_object(
        source_payload,
        new_payload_id,
        0,
        payload.encode_to_vec(),
        payload_references,
        &remap,
        false,
    )?;

    package.update_archive(&entry.storage_archive, |archive| {
        archive.insert_object(cloned_storage)
    })?;
    package.update_archive(&entry.payload_archive, |archive| {
        archive.insert_object(cloned_payload)
    })?;

    let resolved = resolve_table_data_list(
        package,
        locations,
        entry.table_id,
        tst::table_data_list::ListType::RichTextPayload,
    )?;
    let old_entry = resolved
        .entries
        .iter()
        .find(|candidate| candidate.entry.key == identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers rich-text table has no entry {identifier}"))
        })?;
    if old_entry.entry.refcount < 2 || old_entry.owner != entry.owner {
        return Err(Error::InvalidFormat(format!(
            "Numbers rich-text entry {identifier} stopped being shared during copy-on-write"
        )));
    }
    let key = next_table_data_list_key(&resolved.list, &resolved.entries)?;
    decrement_table_data_list_entry(
        package,
        locations,
        &resolved,
        old_entry,
        tst::table_data_list::ListType::RichTextPayload,
    )?;

    package.update_archive(&entry.table_archive, |archive| {
        let object = archive.object_mut(entry.table_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers rich-text table object {} is missing",
                entry.table_id
            ))
        })?;
        let message_index =
            table_data_list_message_index(object, tst::table_data_list::ListType::RichTextPayload)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Object {} has no Numbers rich-text TableDataList payload",
                        entry.table_id
                    ))
                })?;
        let message_type = object.messages[message_index].type_;
        let original = object.messages[message_index].data.as_slice();
        let previous = TableDataList::decode(original)?;
        let mut list = previous.clone();
        list.next_list_id = key
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers rich-text identifier overflow".to_owned()))?;
        list.entries.push(tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            rich_text_payload: Some(crate::protobuf::tsp::Reference {
                identifier: new_payload_id,
                ..Default::default()
            }),
            ..Default::default()
        });
        let data = rewrite_table_data_list_wire(original, &previous, &list)?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        add_message_object_reference(object, message_index, entry.payload_id, new_payload_id);
        Ok(())
    })?;

    let mut text = IWorkTextEditor::from_package(package.clone());
    text.set_text(new_storage_id, replacement)?;
    *package = text.into_package();
    Ok(key)
}

pub(super) fn rich_text_cell_id(row: usize, column: usize) -> Result<tst::CellId> {
    let row = u32::try_from(row)
        .map_err(|_| Error::ParseError("Numbers rich-text row exceeds u32".to_owned()))?;
    let column = u32::try_from(column)
        .map_err(|_| Error::ParseError("Numbers rich-text column exceeds u32".to_owned()))?;
    if row <= u16::MAX.into() && column <= u16::MAX.into() {
        Ok(tst::CellId {
            packed_data: column << 16 | row,
            expanded_coord: None,
        })
    } else {
        Ok(tst::CellId {
            packed_data: u32::MAX,
            expanded_coord: Some(tsce::CellCoordinateArchive {
                packed_data: None,
                column: Some(column),
                row: Some(row),
            }),
        })
    }
}

pub(super) fn next_table_data_list_key(
    list: &TableDataList,
    entries: &[LocatedTableDataListEntry],
) -> Result<u32> {
    let after_entries = entries
        .iter()
        .map(|located| located.entry.key)
        .max()
        .map_or(Ok(1), |key| {
            key.checked_add(1).ok_or_else(|| {
                Error::ParseError("Numbers table-data-list identifier overflow".to_owned())
            })
        })?;
    Ok(list.next_list_id.max(after_entries).max(1))
}

pub(super) fn add_message_object_reference(
    object: &mut ArchiveObject,
    message_index: usize,
    sibling: u64,
    identifier: u64,
) {
    let info = &mut object.archive_info.message_infos[message_index];
    if !info.object_references.contains(&identifier) {
        info.object_references.push(identifier);
    }
    for field in &mut info.field_infos {
        if field.object_references.contains(&sibling)
            && !field.object_references.contains(&identifier)
        {
            field.object_references.push(identifier);
        }
    }
}

pub(super) fn remove_message_object_reference(
    object: &mut ArchiveObject,
    message_index: usize,
    identifier: u64,
) {
    let info = &mut object.archive_info.message_infos[message_index];
    info.object_references
        .retain(|reference| *reference != identifier);
    for field in &mut info.field_infos {
        field
            .object_references
            .retain(|reference| *reference != identifier);
    }
}

pub(super) fn release_rich_text(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    identifier: u32,
) -> Result<()> {
    let entry = rich_text_entry_location(package, locations, model, identifier)?;
    let resolved = resolve_table_data_list(
        package,
        locations,
        entry.table_id,
        tst::table_data_list::ListType::RichTextPayload,
    )?;
    let located = resolved
        .entries
        .iter()
        .find(|candidate| candidate.entry.key == identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers rich-text table has no entry {identifier}"))
        })?;
    let entry_removed = decrement_table_data_list_entry(
        package,
        locations,
        &resolved,
        located,
        tst::table_data_list::ListType::RichTextPayload,
    )?;

    if entry_removed && !package_references_object(package, locations, entry.payload_id)? {
        remove_object_or_empty_entry(package, locations, entry.payload_id)?;
        if !package_references_object(package, locations, entry.storage_id)? {
            remove_object_or_empty_entry(package, locations, entry.storage_id)?;
        }
    }
    Ok(())
}

pub(super) fn package_references_object(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
) -> Result<bool> {
    let archive_names = locations.values().collect::<HashSet<_>>();
    for archive_name in archive_names {
        if !package.contains_entry(archive_name) {
            continue;
        }
        let archive = package.archive(archive_name)?;
        if archive.objects.iter().any(|object| {
            object.archive_info.message_infos.iter().any(|message| {
                message.object_references.contains(&identifier)
                    || message
                        .field_infos
                        .iter()
                        .any(|field| field.object_references.contains(&identifier))
            })
        }) {
            return Ok(true);
        }
    }
    Ok(false)
}

pub(super) fn decrement_formula_table(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    identifier: u32,
) -> Result<()> {
    let resolved = resolve_table_data_list(
        package,
        locations,
        table_id,
        tst::table_data_list::ListType::Formula,
    )?;
    let located = resolved
        .entries
        .iter()
        .find(|located| located.entry.key == identifier && located.entry.formula.is_some())
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers formula table has no formula entry {identifier}"
            ))
        })?;
    decrement_table_data_list_entry(
        package,
        locations,
        &resolved,
        located,
        tst::table_data_list::ListType::Formula,
    )?;
    Ok(())
}

pub(super) fn rewrite_formula_table_entry<F>(
    package: &mut IWorkPackage,
    resolved: &ResolvedTableDataList,
    located: &LocatedTableDataListEntry,
    current_formula: &tsce::FormulaArchive,
    rewrite_formula_wire: F,
) -> Result<()>
where
    F: FnOnce(&[u8]) -> Result<Vec<u8>>,
{
    fn rewrite_container<F>(
        original: &[u8],
        previous_entries: &[tst::table_data_list::ListEntry],
        key: u32,
        current_formula: &tsce::FormulaArchive,
        rewrite_formula_wire: F,
    ) -> Result<Vec<u8>>
    where
        F: FnOnce(&[u8]) -> Result<Vec<u8>>,
    {
        let raw_entries = repeated_length_delimited_payloads(original, 3)?;
        if raw_entries.len() != previous_entries.len() {
            return Err(Error::InvalidFormat(format!(
                "Numbers formula table has {} raw entries but {} decoded entries",
                raw_entries.len(),
                previous_entries.len()
            )));
        }
        let position = previous_entries
            .iter()
            .position(|entry| entry.key == key)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers formula table has no entry {key}"))
            })?;
        let previous_entry = &previous_entries[position];
        let previous_formula = previous_entry.formula.as_ref().ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers formula table entry {key} has no formula payload"
            ))
        })?;
        if tst::table_data_list::ListEntry::decode(raw_entries[position])? != *previous_entry {
            return Err(Error::InvalidFormat(format!(
                "Numbers formula table entry {key} changed during wire mutation"
            )));
        }
        let rewritten_entry =
            transform_length_delimited_field(raw_entries[position], 5, rewrite_formula_wire)?;
        let mut expected = previous_entry.clone();
        expected.formula = Some(current_formula.clone());
        if tst::table_data_list::ListEntry::decode(rewritten_entry.as_slice())? != expected {
            return Err(Error::InvalidFormat(format!(
                "Numbers formula table entry {key} wire mutation failed validation"
            )));
        }
        let raw_formulas = repeated_length_delimited_payloads(rewritten_entry.as_slice(), 5)?;
        let [raw_formula] = raw_formulas.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Numbers formula table entry {key} must contain exactly one formula payload"
            )));
        };
        if tsce::FormulaArchive::decode(*raw_formula)? != *current_formula
            || previous_formula == current_formula
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers formula table entry {key} did not receive a distinct valid formula"
            )));
        }
        let mut replacements = raw_entries
            .into_iter()
            .map(<[u8]>::to_vec)
            .collect::<Vec<_>>();
        replacements[position] = rewritten_entry;
        rewrite_repeated_length_delimited_fields(original, 3, &replacements)
    }

    let key = located.entry.key;
    match &located.owner {
        TableDataListEntryOwner::Root => {
            package.update_archive(&resolved.table_archive, |archive| {
                let object = archive.object_mut(resolved.table_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers formula table object {} is missing",
                        resolved.table_id
                    ))
                })?;
                let message_index =
                    table_data_list_message_index(object, tst::table_data_list::ListType::Formula)
                        .ok_or_else(|| {
                            Error::InvalidFormat(format!(
                                "Object {} has no Numbers formula TableDataList payload",
                                resolved.table_id
                            ))
                        })?;
                let original = object.messages[message_index].data.as_slice();
                let previous = TableDataList::decode(original)?;
                let data = rewrite_container(
                    original,
                    &previous.entries,
                    key,
                    current_formula,
                    rewrite_formula_wire,
                )?;
                let mut expected = previous;
                let entry = expected
                    .entries
                    .iter_mut()
                    .find(|entry| entry.key == key)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers formula table has no entry {key} after wire mutation"
                        ))
                    })?;
                entry.formula = Some(current_formula.clone());
                if TableDataList::decode(data.as_slice())? != expected {
                    return Err(Error::InvalidFormat(
                        "Numbers formula table wire mutation failed validation".to_owned(),
                    ));
                }
                let message_type = object.messages[message_index].type_;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
                Ok(())
            })
        },
        TableDataListEntryOwner::Segment { object_id, archive } => {
            package.update_archive(archive, |archive_data| {
                let object = archive_data.object_mut(*object_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers formula table segment object {object_id} is missing"
                    ))
                })?;
                let message_index = object
                    .messages
                    .iter()
                    .position(|message| message.type_ == TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Object {object_id} has no Numbers TableDataListSegment payload"
                        ))
                    })?;
                let original = object.messages[message_index].data.as_slice();
                let previous = TableDataListSegment::decode(original)?;
                let data = rewrite_container(
                    original,
                    &previous.entries,
                    key,
                    current_formula,
                    rewrite_formula_wire,
                )?;
                let mut expected = previous;
                let entry = expected
                    .entries
                    .iter_mut()
                    .find(|entry| entry.key == key)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers formula table segment has no entry {key} after wire mutation"
                        ))
                    })?;
                entry.formula = Some(current_formula.clone());
                if TableDataListSegment::decode(data.as_slice())? != expected {
                    return Err(Error::InvalidFormat(
                        "Numbers formula table segment wire mutation failed validation".to_owned(),
                    ));
                }
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE,
                        data,
                    },
                )?;
                Ok(())
            })
        },
    }
}

pub(super) fn decrement_formula_error_table(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    identifier: u32,
) -> Result<()> {
    let table_id = model
        .base_data_store
        .formula_error_table
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers cell references formula-error entry {identifier}, but table {:?} has no formula-error list",
                model.table_name
            ))
        })?;
    let resolved = resolve_table_data_list(
        package,
        locations,
        table_id,
        tst::table_data_list::ListType::FormulaError,
    )?;
    let located = resolved
        .entries
        .iter()
        .find(|located| located.entry.key == identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers formula-error table has no entry {identifier}"
            ))
        })?;
    decrement_table_data_list_entry(
        package,
        locations,
        &resolved,
        located,
        tst::table_data_list::ListType::FormulaError,
    )?;
    Ok(())
}

pub(super) fn table_data_list_message_index(
    object: &ArchiveObject,
    list_type: tst::table_data_list::ListType,
) -> Option<usize> {
    object.messages.iter().position(|message| {
        (message.type_ == 6005 || message.type_ == 6201)
            && TableDataList::decode(message.data.as_slice())
                .is_ok_and(|list| list.list_type == list_type as i32)
    })
}

pub(super) fn insert_formula_table(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    formula: tsce::FormulaArchive,
) -> Result<u32> {
    let resolved = resolve_table_data_list(
        package,
        locations,
        table_id,
        tst::table_data_list::ListType::Formula,
    )?;
    if let Some(located) = resolved
        .entries
        .iter()
        .find(|located| located.entry.formula.as_ref() == Some(&formula))
    {
        increment_table_data_list_entry(
            package,
            locations,
            &resolved,
            located,
            tst::table_data_list::ListType::Formula,
        )?;
        return Ok(located.entry.key);
    }
    let key = next_table_data_list_key(&resolved.list, &resolved.entries)?;
    package.update_archive(&resolved.table_archive, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::ParseError(format!(
                "Numbers formula table object {table_id} is missing"
            ))
        })?;
        let message_index =
            table_data_list_message_index(object, tst::table_data_list::ListType::Formula)
                .ok_or_else(|| {
                    Error::ParseError(format!(
                        "Object {table_id} has no Numbers formula TableDataList payload"
                    ))
                })?;
        let message_type = object.messages[message_index].type_;
        let original = object.messages[message_index].data.as_slice();
        let previous = TableDataList::decode(original)?;
        let mut list = previous.clone();
        list.next_list_id = key.checked_add(1).ok_or_else(|| {
            Error::ParseError("Numbers formula table identifier overflow".to_owned())
        })?;
        list.entries.push(tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            formula: Some(formula.clone()),
            ..Default::default()
        });
        let data = rewrite_table_data_list_wire(original, &previous, &list)?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })?;
    Ok(key)
}

pub(super) fn set_encoded_cell_value(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: EncodedValue,
) -> Result<()> {
    table_sparse_storage::ensure_attached_cell_storage(package, table_id, row, column)?;
    let location = locate_attached_cell(package, table_id, row, column)?;
    let cell_count = update_tile(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
        location.descriptor.model.number_of_columns as usize,
        value,
    )?;
    update_row_header(
        package,
        &location.object_locations,
        &location.descriptor.model,
        row,
        cell_count,
    )
}

pub(super) fn verify_formula_link(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    formula_id: u32,
    expected: &tsce::FormulaArchive,
) -> Result<()> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    let stored = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    .ok_or_else(|| Error::InvalidFormat("Numbers formula cell is missing".to_owned()))?;
    if BncCell::parse(&stored)?.stored_value() != StoredValue::Formula(formula_id) {
        return Err(Error::InvalidFormat(
            "Numbers formula cell reference failed validation".to_owned(),
        ));
    }

    let formula_table_id = location
        .descriptor
        .model
        .base_data_store
        .formula_table
        .identifier;
    let resolved = resolve_table_data_list(
        package,
        &location.object_locations,
        formula_table_id,
        tst::table_data_list::ListType::Formula,
    )?;
    if !resolved.entries.iter().any(|located| {
        located.entry.key == formula_id
            && located.entry.refcount > 0
            && located.entry.formula.as_ref() == Some(expected)
    }) {
        return Err(Error::InvalidFormat(format!(
            "Numbers formula table entry {formula_id} failed validation"
        )));
    }
    Ok(())
}

pub(super) fn rewrite_expanded_cell_records(
    data: &[u8],
    field_number: u32,
    previous: &[tsce::CellRecordExpandedArchive],
    current: &[tsce::CellRecordExpandedArchive],
) -> Result<Vec<u8>> {
    let raw_records = repeated_length_delimited_payloads(data, field_number)?;
    if raw_records.len() != previous.len() {
        return Err(Error::InvalidFormat(format!(
            "Numbers dependency field {field_number} has {} raw records but {} decoded records",
            raw_records.len(),
            previous.len()
        )));
    }
    let mut existing = HashMap::with_capacity(previous.len());
    for (expected, raw) in previous.iter().zip(raw_records) {
        if tsce::CellRecordExpandedArchive::decode(raw)? != *expected {
            return Err(Error::InvalidFormat(format!(
                "Numbers dependency record ({}, {}) changed during wire mutation",
                expected.row, expected.column
            )));
        }
        patch_varint_field(raw, 1, true, Some(u64::from(expected.column)))?;
        patch_varint_field(raw, 2, true, Some(u64::from(expected.row)))?;
        if existing
            .insert((expected.row, expected.column), raw)
            .is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers dependencies contain duplicate cell ({}, {})",
                expected.row, expected.column
            )));
        }
    }
    let mut seen = HashSet::with_capacity(current.len());
    let replacements = current
        .iter()
        .map(|record| {
            let key = (record.row, record.column);
            if !seen.insert(key) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers dependencies would contain duplicate cell ({}, {})",
                    record.row, record.column
                )));
            }
            Ok(existing
                .get(&key)
                .map_or_else(|| record.encode_to_vec(), |raw| raw.to_vec()))
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, field_number, &replacements)
}

pub(super) fn rewrite_formula_owner_dependencies_wire(
    original: &[u8],
    previous: &tsce::FormulaOwnerDependenciesArchive,
    current: &tsce::FormulaOwnerDependenciesArchive,
) -> Result<Vec<u8>> {
    let mut immutable = current.clone();
    immutable.cell_dependencies = previous.cell_dependencies.clone();
    immutable.tiled_cell_dependencies = previous.tiled_cell_dependencies.clone();
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "Numbers formula-owner immutable fields changed during mutation".to_owned(),
        ));
    }
    let mut data = original.to_vec();
    match (&previous.cell_dependencies, &current.cell_dependencies) {
        (Some(previous), Some(current)) => {
            data = transform_length_delimited_field(&data, 4, |dependencies| {
                rewrite_expanded_cell_records(
                    dependencies,
                    1,
                    &previous.cell_record,
                    &current.cell_record,
                )
            })?;
        },
        (None, None) => {},
        _ => {
            return Err(Error::InvalidFormat(
                "Numbers inline dependency storage changed representation".to_owned(),
            ));
        },
    }
    match (
        &previous.tiled_cell_dependencies,
        &current.tiled_cell_dependencies,
    ) {
        (Some(previous), Some(current)) => {
            data = transform_length_delimited_field(&data, 13, |dependencies| {
                rewrite_reference_list(
                    dependencies,
                    1,
                    &previous
                        .cell_record_tiles
                        .iter()
                        .map(|reference| reference.identifier)
                        .collect::<Vec<_>>(),
                    &current
                        .cell_record_tiles
                        .iter()
                        .map(|reference| reference.identifier)
                        .collect::<Vec<_>>(),
                )
            })?;
        },
        (None, Some(current)) => {
            data = patch_length_delimited_field(&data, 13, false, Some(&current.encode_to_vec()))?;
        },
        (Some(_), None) => {
            data = patch_length_delimited_field(&data, 13, true, None)?;
        },
        (None, None) => {},
    }
    if tsce::FormulaOwnerDependenciesArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers formula-owner wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn rewrite_dependency_tile_wire(
    original: &[u8],
    previous: &tsce::CellRecordTileArchive,
    current: &tsce::CellRecordTileArchive,
) -> Result<Vec<u8>> {
    if previous.internal_owner_id != current.internal_owner_id
        || previous.tile_column_begin != current.tile_column_begin
        || previous.tile_row_begin != current.tile_row_begin
    {
        return Err(Error::InvalidFormat(
            "Numbers dependency-tile identity changed during mutation".to_owned(),
        ));
    }
    let data =
        rewrite_expanded_cell_records(original, 4, &previous.cell_records, &current.cell_records)?;
    if tsce::CellRecordTileArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers dependency-tile wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn rewrite_calculation_engine_formula_count_wire(
    original: &[u8],
    previous: &tsce::CalculationEngineArchive,
    current: &tsce::CalculationEngineArchive,
) -> Result<Vec<u8>> {
    let mut immutable = current.clone();
    immutable.dependency_tracker.number_of_formulas =
        previous.dependency_tracker.number_of_formulas;
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "Numbers CalculationEngine fields changed outside the formula count".to_owned(),
        ));
    }
    let data = patch_nested_varint_field(
        original,
        &[2, 5],
        previous.dependency_tracker.number_of_formulas.is_some(),
        current.dependency_tracker.number_of_formulas,
    )?;
    if tsce::CalculationEngineArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers CalculationEngine formula-count wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn update_formula_dependencies(
    package: &mut IWorkPackage,
    table_info_id: u64,
    row: usize,
    column: usize,
    present: bool,
    local_precedents: &[(u32, u32)],
    external_precedents: &[(u32, u32, u32)],
) -> Result<()> {
    // Minimal synthetic packages used by embedders may omit CalculationEngine;
    // real Numbers documents always carry it.
    let Some(component) = package.calculation_engine_entry_name()?.map(str::to_owned) else {
        return Ok(());
    };
    let row = u32::try_from(row)
        .map_err(|_| Error::ParseError("Numbers formula row exceeds u32".to_owned()))?;
    let column = u32::try_from(column)
        .map_err(|_| Error::ParseError("Numbers formula column exceeds u32".to_owned()))?;
    let tile_row_begin = row / FORMULA_DEPENDENCY_TILE_ROWS * FORMULA_DEPENDENCY_TILE_ROWS;
    let tile_column_begin =
        column / FORMULA_DEPENDENCY_TILE_COLUMNS * FORMULA_DEPENDENCY_TILE_COLUMNS;
    let new_identifier = object_locations(package)?
        .keys()
        .copied()
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;

    package.update_archive(&component, |archive| {
        let engine_id = archive
            .objects
            .iter()
            .find_map(|object| {
                object
                    .messages
                    .iter()
                    .any(|message| message.type_ == 4000)
                    .then_some(object.archive_info.identifier)
                    .flatten()
            })
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers CalculationEngine root is missing".to_owned())
            })?;
        let (owner_id, owner_message_index, owner_original, mut owner) = archive
            .objects
            .iter()
            .find_map(|object| {
                object
                    .messages
                    .iter()
                    .enumerate()
                    .filter(|(_, message)| message.type_ == 4008)
                    .find_map(|(message_index, message)| {
                        let owner =
                            tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())
                                .ok()?;
                        (owner
                            .formula_owner
                            .as_ref()
                            .map(|reference| reference.identifier)
                            == Some(table_info_id))
                        .then_some((
                            object.archive_info.identifier?,
                            message_index,
                            message.data.clone(),
                            owner,
                        ))
                    })
            })
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers table info {table_info_id} has no formula dependency owner"
                ))
            })?;
        let owner_previous = owner.clone();

        let uses_inline = owner.cell_dependencies.is_some();
        let uses_tiled = owner.tiled_cell_dependencies.is_some() || !uses_inline;
        let existing_inline = owner.cell_dependencies.as_ref().and_then(|dependencies| {
            dependencies
                .cell_record
                .iter()
                .position(|record| record.row == row && record.column == column)
        });
        if uses_inline && present == existing_inline.is_some() {
            return Err(Error::InvalidFormat(format!(
                "Numbers formula dependency for cell ({row}, {column}) has an unexpected existing state"
            )));
        }

        let tiled_references = owner
            .tiled_cell_dependencies
            .as_ref()
            .map(|dependencies| dependencies.cell_record_tiles.clone())
            .unwrap_or_default();
        let existing_tile = tiled_references
            .iter()
            .find_map(|reference| {
                let object = archive.object(reference.identifier)?;
                let tile = object.messages.iter().find_map(|message| {
                    (message.type_ == 4009)
                        .then(|| tsce::CellRecordTileArchive::decode(message.data.as_slice()).ok())
                        .flatten()
                })?;
                (tile.tile_row_begin == tile_row_begin
                    && tile.tile_column_begin == tile_column_begin)
                    .then_some((reference.identifier, tile))
            });

        let expanded_edges = expanded_formula_edges(local_precedents, external_precedents)?;
        if let Some(inline_dependencies) = owner.cell_dependencies.as_mut() {
            if present {
                inline_dependencies
                    .cell_record
                    .push(tsce::CellRecordExpandedArchive {
                        column,
                        row,
                        expanded_edges: Some(expanded_edges.clone()),
                        ..Default::default()
                    });
                inline_dependencies
                    .cell_record
                    .sort_by_key(|record| (record.row, record.column));
            } else {
                let position = existing_inline.ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers inline formula dependency for cell ({row}, {column}) is missing"
                    ))
                })?;
                inline_dependencies.cell_record.remove(position);
            }
        }

        let mut removed_tile_id = None;
        let tile_id = if !uses_tiled {
            None
        } else if let Some((tile_id, mut tile)) = existing_tile {
            let tile_previous = tile.clone();
            let position = tile
                .cell_records
                .iter()
                .position(|record| record.row == row && record.column == column);
            if present == position.is_some() {
                return Err(Error::InvalidFormat(format!(
                    "Numbers tiled formula dependency for cell ({row}, {column}) has an unexpected existing state"
                )));
            }
            if present {
                tile.cell_records.push(tsce::CellRecordExpandedArchive {
                    column,
                    row,
                    expanded_edges: Some(expanded_edges.clone()),
                    ..Default::default()
                });
                tile.cell_records
                    .sort_by_key(|record| (record.row, record.column));
            } else {
                let position = position.ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers tiled formula dependency for cell ({row}, {column}) is missing"
                    ))
                })?;
                tile.cell_records.remove(position);
            }
            if !present && tile.cell_records.is_empty() {
                owner
                    .tiled_cell_dependencies
                    .as_mut()
                    .ok_or_else(|| {
                        Error::InvalidFormat(
                            "Numbers tiled dependencies disappeared during cleanup".to_owned(),
                        )
                    })?
                    .cell_record_tiles
                    .retain(|reference| reference.identifier != tile_id);
                archive.remove_object(tile_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers formula dependency tile {tile_id} is missing"
                    ))
                })?;
                removed_tile_id = Some(tile_id);
                None
            } else {
            let object = archive.object_mut(tile_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers formula dependency tile {tile_id} is missing"
                ))
            })?;
                let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == 4009)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers formula dependency tile {tile_id} has no payload"
                        ))
                    })?;
                let data = rewrite_dependency_tile_wire(
                    object.messages[message_index].data.as_slice(),
                    &tile_previous,
                    &tile,
                )?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: 4009,
                        data,
                    },
            )?;
            Some(tile_id)
            }
        } else if present {
            let tile = tsce::CellRecordTileArchive {
                internal_owner_id: owner.internal_formula_owner_id,
                tile_column_begin,
                tile_row_begin,
                cell_records: vec![tsce::CellRecordExpandedArchive {
                    column,
                    row,
                    expanded_edges: Some(expanded_edges),
                    ..Default::default()
                }],
            };
            owner
                .tiled_cell_dependencies
                .get_or_insert_default()
                .cell_record_tiles
                .push(crate::protobuf::tsp::Reference {
                    identifier: new_identifier,
                    ..Default::default()
                });
            archive.insert_object(ArchiveObject::new(
                new_identifier,
                vec![RawMessage {
                    type_: 4009,
                    data: tile.encode_to_vec(),
                }],
            )?)?;
            Some(new_identifier)
        } else {
            return Err(Error::InvalidFormat(format!(
                "Numbers formula dependency tile for cell ({row}, {column}) is missing"
            )));
        };

        let owner_object = archive.object_mut(owner_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers formula dependency owner {owner_id} is missing"))
        })?;
        let owner_message_type = owner_object.messages[owner_message_index].type_;
        let owner_data =
            rewrite_formula_owner_dependencies_wire(&owner_original, &owner_previous, &owner)?;
        owner_object.replace_message(
            owner_message_index,
            RawMessage {
                type_: owner_message_type,
                data: owner_data,
            },
        )?;
        if present && let Some(tile_id) = tile_id {
            let references = &mut owner_object.archive_info.message_infos[owner_message_index]
                .object_references;
            if !references.contains(&tile_id) {
                references.push(tile_id);
            }
        }
        if let Some(tile_id) = removed_tile_id {
            remove_message_object_reference(owner_object, owner_message_index, tile_id);
        }

        let engine_object = archive.object_mut(engine_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers CalculationEngine root {engine_id} is missing"))
        })?;
        let engine_message_index = engine_object
            .messages
            .iter()
            .position(|message| message.type_ == 4000)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers CalculationEngine payload is missing".to_owned())
            })?;
        let engine_original = engine_object.messages[engine_message_index].data.as_slice();
        let engine_previous = tsce::CalculationEngineArchive::decode(engine_original)?;
        let mut engine = engine_previous.clone();
        let current = engine.dependency_tracker.number_of_formulas.unwrap_or(0);
        engine.dependency_tracker.number_of_formulas = Some(if present {
            current.checked_add(1).ok_or_else(|| {
                Error::ParseError("Numbers formula count overflow".to_owned())
            })?
        } else {
            current.checked_sub(1).ok_or_else(|| {
                Error::InvalidFormat("Numbers formula count underflow".to_owned())
            })?
        });
        let engine_data = rewrite_calculation_engine_formula_count_wire(
            engine_original,
            &engine_previous,
            &engine,
        )?;
        engine_object.replace_message(
            engine_message_index,
            RawMessage {
                type_: 4000,
                data: engine_data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn expanded_formula_edges(
    local_precedents: &[(u32, u32)],
    external_precedents: &[(u32, u32, u32)],
) -> Result<tsce::ExpandedEdgesArchive> {
    if local_precedents.windows(2).any(|pair| pair[0] >= pair[1]) {
        return Err(Error::InvalidFormat(
            "Numbers formula precedents must be sorted and unique".to_owned(),
        ));
    }
    if external_precedents
        .windows(2)
        .any(|pair| pair[0] >= pair[1])
    {
        return Err(Error::InvalidFormat(
            "Numbers external formula precedents must be sorted and unique".to_owned(),
        ));
    }
    Ok(tsce::ExpandedEdgesArchive {
        edge_without_owner_rows: local_precedents.iter().map(|(row, _)| *row).collect(),
        edge_without_owner_columns: local_precedents.iter().map(|(_, column)| *column).collect(),
        edge_with_owner_rows: external_precedents.iter().map(|(_, row, _)| *row).collect(),
        edge_with_owner_columns: external_precedents
            .iter()
            .map(|(_, _, column)| *column)
            .collect(),
        internal_owner_id_for_edge: external_precedents
            .iter()
            .map(|(owner, _, _)| *owner)
            .collect(),
    })
}

pub(super) fn verify_formula_dependency(
    package: &IWorkPackage,
    table_info_id: u64,
    row: usize,
    column: usize,
    local_precedents: &[(u32, u32)],
    external_precedents: &[(u32, u32, u32)],
) -> Result<()> {
    let Some(component) = package.calculation_engine_entry_name()? else {
        return Ok(());
    };
    let row = u32::try_from(row)
        .map_err(|_| Error::ParseError("Numbers formula row exceeds u32".to_owned()))?;
    let column = u32::try_from(column)
        .map_err(|_| Error::ParseError("Numbers formula column exceeds u32".to_owned()))?;
    let archive = package.archive(component)?;
    let expected_edges = expanded_formula_edges(local_precedents, external_precedents)?;
    let owner = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == 4008)
        .find_map(|message| {
            let owner =
                tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice()).ok()?;
            (owner
                .formula_owner
                .as_ref()
                .map(|reference| reference.identifier)
                == Some(table_info_id))
            .then_some(owner)
        })
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table info {table_info_id} has no formula dependency owner"
            ))
        })?;
    let has_inline = owner.cell_dependencies.is_some();
    let has_tiled = owner.tiled_cell_dependencies.is_some();
    let inline_valid = owner.cell_dependencies.as_ref().is_none_or(|dependencies| {
        dependencies.cell_record.iter().any(|record| {
            record.row == row
                && record.column == column
                && record.expanded_edges.as_ref() == Some(&expected_edges)
        })
    });
    let tile_row_begin = row / FORMULA_DEPENDENCY_TILE_ROWS * FORMULA_DEPENDENCY_TILE_ROWS;
    let tile_column_begin =
        column / FORMULA_DEPENDENCY_TILE_COLUMNS * FORMULA_DEPENDENCY_TILE_COLUMNS;
    let tile_valid = owner
        .tiled_cell_dependencies
        .as_ref()
        .is_none_or(|dependencies| {
            dependencies.cell_record_tiles.iter().any(|reference| {
                archive.object(reference.identifier).is_some_and(|object| {
                    object.messages.iter().any(|message| {
                        message.type_ == 4009
                            && tsce::CellRecordTileArchive::decode(message.data.as_slice())
                                .is_ok_and(|tile| {
                                    tile.tile_row_begin == tile_row_begin
                                        && tile.tile_column_begin == tile_column_begin
                                        && tile.cell_records.iter().any(|record| {
                                            record.row == row
                                                && record.column == column
                                                && record.expanded_edges.as_ref()
                                                    == Some(&expected_edges)
                                        })
                                })
                    })
                })
            })
        });
    if (!has_inline && !has_tiled) || !inline_valid || !tile_valid {
        return Err(Error::InvalidFormat(format!(
            "Numbers formula dependency for cell ({row}, {column}) failed validation"
        )));
    }
    Ok(())
}

pub(super) fn decrement_old_string(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    old_string: Option<u32>,
) -> Result<()> {
    if let Some(identifier) = old_string {
        update_string_table(
            package,
            locations,
            model.base_data_store.string_table.identifier,
            Some(identifier),
            None,
        )?;
    }
    Ok(())
}

pub(super) fn table_models(package: &IWorkPackage) -> Result<Vec<TableDescriptor>> {
    let document = numbers_document(package)?;
    let locations = object_locations(package)?;
    let mut result = Vec::new();
    let mut seen = HashSet::new();
    for sheet_reference in document.sheets {
        let sheet_id = sheet_reference.identifier;
        let sheet_archive_name = locations.get(&sheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet object {sheet_id} is missing"))
        })?;
        let sheet_archive = package.archive(sheet_archive_name)?;
        let sheet_object = sheet_archive.object(sheet_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers sheet object {sheet_id} is missing"))
        })?;
        let (_, sheet) = decode_sheet(sheet_object)?;
        for drawable in sheet.drawable_infos {
            let drawable_id = drawable.identifier;
            let drawable_archive_name = locations.get(&drawable_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers sheet {sheet_id} drawable {drawable_id} is missing"
                ))
            })?;
            let drawable_archive = package.archive(drawable_archive_name)?;
            let drawable_object = drawable_archive.object(drawable_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers drawable object {drawable_id} is missing"))
            })?;
            // TST type numbers differ between archive generations. Resolve a
            // candidate only when its table-model reference actually lands on
            // a decodable TableModelArchive, which also avoids protobuf's
            // permissive decoding of unrelated drawable payloads.
            let mut resolved_model = None;
            for message in &drawable_object.messages {
                let Ok(table_info) = tst::TableInfoArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                let candidate_id = table_info.table_model.identifier;
                let Some(model_archive_name) = locations.get(&candidate_id) else {
                    continue;
                };
                let model_archive = package.archive(model_archive_name)?;
                let Some(model_object) = model_archive.object(candidate_id) else {
                    continue;
                };
                let Some(model) = model_object.messages.iter().find_map(|message| {
                    (message.type_ == 6000 || message.type_ == 6001)
                        .then(|| TableModelArchive::decode(message.data.as_slice()).ok())
                        .flatten()
                }) else {
                    continue;
                };
                resolved_model = Some((candidate_id, model));
                break;
            }
            let Some((object_id, model)) = resolved_model else {
                continue;
            };
            if !seen.insert(object_id) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table model {object_id} is attached more than once"
                )));
            }
            result.push(TableDescriptor {
                object_id,
                table_info_id: drawable_id,
                model,
            });
        }
    }
    Ok(result)
}

pub(super) fn object_locations(package: &IWorkPackage) -> Result<HashMap<u64, String>> {
    let mut locations = HashMap::new();
    for name in package.iwa_entry_names() {
        let archive = package.archive(name)?;
        for object in archive.objects {
            let identifier = object
                .archive_info
                .identifier
                .ok_or_else(|| Error::Archive(format!("Object in {name} has no identifier")))?;
            if let Some(previous) = locations.insert(identifier, name.to_owned()) {
                return Err(Error::Archive(format!(
                    "Object {identifier} appears in both {previous} and {name}"
                )));
            }
        }
    }
    Ok(locations)
}

pub(super) fn read_tile_cell(
    package: &IWorkPackage,
    archive_name: &str,
    tile_id: u64,
    row: u32,
    column: usize,
) -> Result<Option<Vec<u8>>> {
    let archive = package.archive(archive_name)?;
    let object = archive
        .object(tile_id)
        .ok_or_else(|| Error::ParseError(format!("Numbers tile object {tile_id} is missing")))?;
    for message in &object.messages {
        let Ok(tile) = Tile::decode(message.data.as_slice()) else {
            continue;
        };
        let Some(row_info) = tile
            .row_infos
            .iter()
            .find(|candidate| candidate.tile_row_index == row)
        else {
            return Ok(None);
        };
        return Ok(split_row(row_info)?.get(column).cloned().flatten());
    }
    Err(Error::ParseError(format!(
        "Object {tile_id} does not contain a Numbers tile payload"
    )))
}

pub(super) fn update_string_table(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    old_identifier: Option<u32>,
    new_text: Option<&str>,
) -> Result<Option<u32>> {
    let resolved = resolve_table_data_list(
        package,
        locations,
        table_id,
        tst::table_data_list::ListType::String,
    )?;
    let old_entry = old_identifier
        .map(|old| {
            resolved
                .entries
                .iter()
                .find(|located| located.entry.key == old && located.entry.string.is_some())
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers string table has no entry for identifier {old}"
                    ))
                })
        })
        .transpose()?;
    if let (Some(old), Some(text)) = (old_entry, new_text)
        && old.entry.string.as_deref() == Some(text)
    {
        return Ok(Some(old.entry.key));
    }

    let result = if let Some(text) = new_text {
        if let Some(existing) = resolved
            .entries
            .iter()
            .find(|located| located.entry.string.as_deref() == Some(text))
        {
            increment_table_data_list_entry(
                package,
                locations,
                &resolved,
                existing,
                tst::table_data_list::ListType::String,
            )?;
            Some(existing.entry.key)
        } else {
            let key = next_table_data_list_key(&resolved.list, &resolved.entries)?;
            package.update_archive(&resolved.table_archive, |archive| {
                let object = archive.object_mut(table_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers string table object {table_id} is missing"
                    ))
                })?;
                let message_index =
                    table_data_list_message_index(object, tst::table_data_list::ListType::String)
                        .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Object {table_id} has no Numbers string TableDataList payload"
                        ))
                    })?;
                let message_type = object.messages[message_index].type_;
                let original = object.messages[message_index].data.as_slice();
                let previous = TableDataList::decode(original)?;
                let mut list = previous.clone();
                list.next_list_id = key.checked_add(1).ok_or_else(|| {
                    Error::ParseError("Numbers string table identifier overflow".to_owned())
                })?;
                list.entries.push(tst::table_data_list::ListEntry {
                    key,
                    refcount: 1,
                    string: Some(text.to_owned()),
                    ..Default::default()
                });
                let data = rewrite_table_data_list_wire(original, &previous, &list)?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
                Ok(())
            })?;
            Some(key)
        }
    } else {
        None
    };

    if let Some(old) = old_entry {
        decrement_table_data_list_entry(
            package,
            locations,
            &resolved,
            old,
            tst::table_data_list::ListType::String,
        )?;
    }
    Ok(result)
}

pub(super) fn update_tile(
    package: &mut IWorkPackage,
    archive_name: &str,
    tile_id: u64,
    row: u32,
    column: usize,
    table_columns: usize,
    value: EncodedValue,
) -> Result<u32> {
    let mut updated_cell_count = None;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(tile_id).ok_or_else(|| {
            Error::ParseError(format!("Numbers tile object {tile_id} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| Tile::decode(message.data.as_slice()).is_ok())
            .ok_or_else(|| {
                Error::ParseError(format!("Object {tile_id} has no Numbers tile payload"))
            })?;
        let raw = &object.messages[message_index];
        let previous = Tile::decode(raw.data.as_slice())?;
        let mut tile = previous.clone();

        let row_position = tile
            .row_infos
            .iter()
            .position(|candidate| candidate.tile_row_index == row);
        if row_position.is_none()
            && matches!(
                &value,
                EncodedValue::Clear
                    | EncodedValue::ClearValuePreservingMetadata
                    | EncodedValue::Comment(None)
                    | EncodedValue::ConditionalStyle {
                        identifier: None,
                        ..
                    }
            )
        {
            updated_cell_count = Some(0);
            return Ok(());
        }

        let mut cells = if let Some(position) = row_position {
            split_row(&tile.row_infos[position])?
        } else {
            vec![None; row_offset_capacity(&tile, table_columns).max(column + 1)]
        };
        if column >= cells.len() {
            cells.resize(column + 1, None);
        }

        match value {
            EncodedValue::Clear => cells[column] = None,
            EncodedValue::ClearValuePreservingMetadata => {
                let Some(data) = cells[column].as_deref() else {
                    updated_cell_count = Some(
                        u32::try_from(cells.iter().filter(|cell| cell.is_some()).count()).map_err(
                            |_| Error::ParseError("Numbers row cell count exceeds u32".to_owned()),
                        )?,
                    );
                    return Ok(());
                };
                let mut cell = BncCell::parse(data)?;
                cell.clear_value_preserving_metadata();
                cells[column] = Some(cell.encode());
            },
            EncodedValue::Comment(identifier) => {
                if let Some(data) = cells[column].as_deref() {
                    let mut cell = BncCell::parse(data)?;
                    cell.set_comment_identifier(identifier);
                    cells[column] = Some(cell.encode());
                } else if identifier.is_some() {
                    let mut cell = BncCell::minimal();
                    cell.set_comment_identifier(identifier);
                    cells[column] = Some(cell.encode());
                } else {
                    updated_cell_count = Some(
                        u32::try_from(cells.iter().filter(|cell| cell.is_some()).count()).map_err(
                            |_| Error::ParseError("Numbers row cell count exceeds u32".to_owned()),
                        )?,
                    );
                    return Ok(());
                }
            },
            EncodedValue::ConditionalStyle {
                identifier,
                applied_rule,
            } => {
                if let Some(data) = cells[column].as_deref() {
                    let mut cell = BncCell::parse(data)?;
                    cell.set_conditional_style(identifier, applied_rule);
                    cells[column] = Some(cell.encode());
                } else if identifier.is_some() {
                    let mut cell = BncCell::minimal();
                    cell.set_conditional_style(identifier, applied_rule);
                    cells[column] = Some(cell.encode());
                } else {
                    updated_cell_count = Some(
                        u32::try_from(cells.iter().filter(|cell| cell.is_some()).count()).map_err(
                            |_| Error::ParseError("Numbers row cell count exceeds u32".to_owned()),
                        )?,
                    );
                    return Ok(());
                }
            },
            EncodedValue::Raw(data) => {
                BncCell::parse(&data)?;
                cells[column] = Some(data);
            },
            value => {
                let mut cell = cells[column]
                    .as_deref()
                    .map(BncCell::parse)
                    .transpose()?
                    .unwrap_or_else(BncCell::minimal);
                match value {
                    EncodedValue::Number(number) => cell.set_number(number)?,
                    EncodedValue::Boolean(boolean) => cell.set_boolean(boolean),
                    EncodedValue::Date(date) => cell.set_date(date)?,
                    EncodedValue::Duration(duration) => cell.set_duration(duration)?,
                    EncodedValue::String(identifier) => cell.set_string(identifier),
                    EncodedValue::RichText(identifier) => cell.set_rich_text(identifier),
                    EncodedValue::Formula(identifier) => cell.set_formula_reference(identifier),
                    EncodedValue::FormulaCachedNumber(number) => {
                        cell.set_formula_cached_number(number)?;
                    },
                    EncodedValue::FormulaCachedBoolean(boolean) => {
                        cell.set_formula_cached_boolean(boolean)?;
                    },
                    EncodedValue::Clear
                    | EncodedValue::ClearValuePreservingMetadata
                    | EncodedValue::Comment(_)
                    | EncodedValue::ConditionalStyle { .. }
                    | EncodedValue::Raw(_) => unreachable!(),
                }
                cells[column] = Some(cell.encode());
            },
        }

        let cell_count = u32::try_from(cells.iter().filter(|cell| cell.is_some()).count())
            .map_err(|_| Error::ParseError("Numbers row cell count exceeds u32".to_owned()))?;
        updated_cell_count = Some(cell_count);

        if cell_count == 0 {
            if let Some(position) = row_position {
                tile.row_infos.remove(position);
            }
        } else if let Some(position) = row_position {
            let previous_wide_offsets = tile.row_infos[position].has_wide_offsets;
            rebuild_row(&mut tile.row_infos[position], &cells)?;
            if previous_wide_offsets.is_none()
                && tile.row_infos[position].has_wide_offsets == Some(false)
            {
                tile.row_infos[position].has_wide_offsets = None;
            }
        } else {
            let mut row_info = TileRowInfo {
                tile_row_index: row,
                cell_count: 0,
                cell_storage_buffer_pre_bnc: Vec::new(),
                cell_offsets_pre_bnc: Vec::new(),
                storage_version: Some(5),
                cell_storage_buffer: Some(Vec::new()),
                cell_offsets: Some(Vec::new()),
                has_wide_offsets: Some(false),
            };
            rebuild_row(&mut row_info, &cells)?;
            tile.row_infos.push(row_info);
            tile.row_infos.sort_by_key(|info| info.tile_row_index);
        }
        tile.numrows = tile
            .row_infos
            .iter()
            .map(|info| info.tile_row_index + 1)
            .max()
            .unwrap_or(0);

        // Some Numbers 3-era files retain complete BNC-v5 buffers alongside
        // their legacy rows but leave the tile-level selector on pre-BNC.
        // Promote only after validating every row so a single semantic edit
        // can never make an incomplete modern representation authoritative.
        if tile.last_saved_in_bnc != Some(true) {
            for row_info in &tile.row_infos {
                if row_info.storage_version != Some(5)
                    || row_info.cell_storage_buffer.is_none()
                    || row_info.cell_offsets.is_none()
                {
                    return Err(Error::ParseError(
                        "Pre-BNC Numbers tile has no complete BNC-v5 mirror to promote".to_owned(),
                    ));
                }
                split_row(row_info)?;
            }
            for row_info in &mut tile.row_infos {
                row_info.cell_storage_buffer_pre_bnc.clear();
                row_info.cell_offsets_pre_bnc.clear();
                row_info.storage_version = Some(5);
            }
            tile.storage_version = Some(5);
            tile.last_saved_in_bnc = Some(true);
        }

        let message_type = raw.type_;
        let data = rewrite_tile_wire(&raw.data, &previous, &tile)?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })?;
    updated_cell_count.ok_or_else(|| {
        Error::InvalidFormat("Numbers tile update returned no row cell count".to_owned())
    })
}

pub(super) fn update_row_header(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    row: usize,
    cell_count: u32,
) -> Result<()> {
    let row_u32 =
        u32::try_from(row).map_err(|_| Error::ParseError("Numbers row exceeds u32".to_owned()))?;
    let bucket_index = row / HEADER_BUCKET_ROWS;
    let bucket_id = model
        .base_data_store
        .row_headers
        .buckets
        .get(bucket_index)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table {:?} has no row-header bucket for row {row}",
                model.table_name
            ))
        })?
        .identifier;
    let archive_name = locations.get(&bucket_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers row-header bucket object {bucket_id} is missing"
        ))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(bucket_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers row-header bucket object {bucket_id} is missing"
            ))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| tst::HeaderStorageBucket::decode(message.data.as_slice()).is_ok())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {bucket_id} has no row-header bucket payload"
                ))
            })?;
        let raw = &object.messages[message_index];
        let previous = tst::HeaderStorageBucket::decode(raw.data.as_slice())?;
        let mut bucket = previous.clone();
        match bucket
            .headers
            .iter()
            .position(|header| header.index == row_u32)
        {
            Some(position) if cell_count == 0 => {
                bucket.headers.remove(position);
            },
            Some(position) => bucket.headers[position].number_of_cells = cell_count,
            None if cell_count > 0 => {
                bucket.headers.push(tst::header_storage_bucket::Header {
                    index: row_u32,
                    size: 0.0,
                    hiding_state: 0,
                    number_of_cells: cell_count,
                    cell_style: None,
                    text_style: None,
                });
            },
            None => {},
        }
        bucket.headers.sort_by_key(|header| header.index);
        let message_type = raw.type_;
        let data = rewrite_header_bucket_wire(&raw.data, &previous, &bucket)?;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        Ok(())
    })
}

pub(super) fn split_row(row: &TileRowInfo) -> Result<Vec<Option<Vec<u8>>>> {
    let storage = row.cell_storage_buffer.as_deref().ok_or_else(|| {
        Error::ParseError("Pre-BNC Numbers rows are not yet writable".to_string())
    })?;
    let offsets = row.cell_offsets.as_deref().ok_or_else(|| {
        Error::ParseError("BNC Numbers row is missing its offset table".to_string())
    })?;
    if offsets.len() % 2 != 0 {
        return Err(Error::ParseError(
            "Numbers cell offset table has an odd byte length".to_string(),
        ));
    }
    let width = if row.has_wide_offsets.unwrap_or(false) {
        4usize
    } else {
        1usize
    };
    let mut starts = Vec::new();
    for (column, bytes) in offsets.chunks_exact(2).enumerate() {
        let raw = u16::from_le_bytes([bytes[0], bytes[1]]);
        if raw == u16::MAX {
            continue;
        }
        let start = usize::from(raw)
            .checked_mul(width)
            .ok_or_else(|| Error::ParseError("Numbers cell offset overflow".to_string()))?;
        if start >= storage.len() {
            return Err(Error::ParseError(format!(
                "Numbers cell offset {start} exceeds storage length {}",
                storage.len()
            )));
        }
        starts.push((column, start));
    }
    if starts.len() != row.cell_count as usize {
        return Err(Error::ParseError(format!(
            "Numbers row declares {} cells but stores {}",
            row.cell_count,
            starts.len()
        )));
    }

    let mut cells = vec![None; offsets.len() / 2];
    for (index, &(column, start)) in starts.iter().enumerate() {
        let end = starts
            .get(index + 1)
            .map_or(storage.len(), |(_, next)| *next);
        if start >= end {
            return Err(Error::ParseError(
                "Numbers cell offsets are not strictly increasing".to_string(),
            ));
        }
        cells[column] = Some(storage[start..end].to_vec());
    }
    Ok(cells)
}

pub(super) fn rebuild_row(row: &mut TileRowInfo, cells: &[Option<Vec<u8>>]) -> Result<()> {
    let prefer_wide = row.has_wide_offsets.unwrap_or(false);
    let (storage, offsets, wide) = match encode_row(cells, prefer_wide) {
        Ok(encoded) => encoded,
        Err(_) if !prefer_wide => encode_row(cells, true)?,
        Err(error) => return Err(error),
    };
    row.cell_count = u32::try_from(cells.iter().filter(|cell| cell.is_some()).count())
        .map_err(|_| Error::ParseError("Numbers row cell count exceeds u32".to_string()))?;
    row.storage_version = Some(5);
    row.cell_storage_buffer = Some(storage);
    row.cell_offsets = Some(offsets);
    row.has_wide_offsets = Some(wide);
    Ok(())
}

pub(super) fn encode_row(
    cells: &[Option<Vec<u8>>],
    wide: bool,
) -> Result<(Vec<u8>, Vec<u8>, bool)> {
    let width = if wide { 4usize } else { 1usize };
    let mut storage = Vec::new();
    let mut offsets = Vec::with_capacity(cells.len() * 2);
    for cell in cells {
        let Some(cell) = cell else {
            offsets.extend_from_slice(&u16::MAX.to_le_bytes());
            continue;
        };
        if wide {
            while storage.len() % width != 0 {
                storage.push(0);
            }
        }
        let offset = storage.len() / width;
        let offset = u16::try_from(offset).map_err(|_| {
            Error::ParseError("Numbers row storage exceeds the BNC offset limit".to_string())
        })?;
        if offset == u16::MAX {
            return Err(Error::ParseError(
                "Numbers row storage collides with the missing-cell sentinel".to_string(),
            ));
        }
        offsets.extend_from_slice(&offset.to_le_bytes());
        storage.extend_from_slice(cell);
    }
    if wide {
        while storage.len() % width != 0 {
            storage.push(0);
        }
    }
    Ok((storage, offsets, wide))
}

pub(super) fn row_offset_capacity(tile: &Tile, table_columns: usize) -> usize {
    tile.row_infos
        .iter()
        .filter_map(|row| row.cell_offsets.as_ref().map(|offsets| offsets.len() / 2))
        .max()
        .unwrap_or(table_columns)
        .max(table_columns)
}
