//! Lossless physical storage moves for an already-validated Numbers row sort.

use std::collections::{HashMap, HashSet};

use prost::Message;

use super::*;

#[derive(Debug)]
struct RelocatedTileRow {
    global_row: usize,
    raw: Vec<u8>,
}

#[derive(Debug)]
struct RelocatedHeader {
    index: usize,
    raw: Vec<u8>,
}

pub(super) fn reorder_body_table_tile_rows(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    body_start: usize,
    destinations_by_source: &[usize],
) -> Result<()> {
    let rows = model.number_of_rows as usize;
    let tile_size = model.base_data_store.tiles.tile_size.unwrap_or(256);
    if tile_size == 0 {
        return Err(Error::InvalidFormat(
            "Numbers table declares a zero tile size".to_owned(),
        ));
    }
    let mut rows_to_write = Vec::new();
    let mut tile_keys = HashMap::with_capacity(model.base_data_store.tiles.tiles.len());
    for reference in &model.base_data_store.tiles.tiles {
        if tile_keys
            .insert(reference.tileid, reference.tile.identifier)
            .is_some()
        {
            return Err(Error::InvalidFormat(format!(
                "Numbers table repeats tile key {}",
                reference.tileid
            )));
        }
        let archive_name = locations.get(&reference.tile.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers tile object {} is missing",
                reference.tile.identifier
            ))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(reference.tile.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers tile object {} is missing",
                reference.tile.identifier
            ))
        })?;
        let message = unique_tile_message(object, reference.tile.identifier)?;
        let tile = Tile::decode(message.data.as_slice())?;
        let raw_rows = repeated_length_delimited_payloads(&message.data, 5)?;
        if raw_rows.len() != tile.row_infos.len() {
            return Err(Error::InvalidFormat(
                "Numbers tile row wire count is inconsistent".to_owned(),
            ));
        }
        let mut source_rows = HashSet::with_capacity(tile.row_infos.len());
        for (row, raw) in tile.row_infos.iter().zip(raw_rows) {
            if TileRowInfo::decode(raw)? != *row
                || row.tile_row_index >= tile_size
                || !source_rows.insert(row.tile_row_index)
            {
                return Err(Error::InvalidFormat(
                    "Numbers tile contains an invalid row payload".to_owned(),
                ));
            }
            let source = reference
                .tileid
                .checked_mul(tile_size)
                .and_then(|base| base.checked_add(row.tile_row_index))
                .ok_or_else(|| Error::ParseError("Numbers tile row overflow".to_owned()))?
                as usize;
            if source >= rows {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table stores row {source} outside its {rows} rows"
                )));
            }
            let destination = relocated_row(source, body_start, destinations_by_source)?;
            rows_to_write.push(RelocatedTileRow {
                global_row: destination,
                raw: patch_varint_field(
                    raw,
                    1,
                    true,
                    Some(
                        u64::try_from(destination % tile_size as usize).map_err(|_| {
                            Error::ParseError("Numbers tile row index exceeds u64".to_owned())
                        })?,
                    ),
                )?,
            });
        }
    }
    rows_to_write.sort_by_key(|row| row.global_row);
    if rows_to_write
        .windows(2)
        .any(|rows| rows[0].global_row == rows[1].global_row)
    {
        return Err(Error::InvalidFormat(
            "Numbers sort would create duplicate stored rows".to_owned(),
        ));
    }
    for row in &rows_to_write {
        let tile_key = u32::try_from(row.global_row / tile_size as usize)
            .map_err(|_| Error::ParseError("Numbers tile key exceeds u32".to_owned()))?;
        if !tile_keys.contains_key(&tile_key) {
            return Err(Error::ParseError(format!(
                "Cannot execute Numbers sort: stored row {} would require an unallocated tile",
                row.global_row
            )));
        }
    }

    for (tile_key, tile_id) in tile_keys {
        let archive_name = locations.get(&tile_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers tile object {tile_id} is missing"))
        })?;
        let desired = rows_to_write
            .iter()
            .filter(|row| row.global_row / tile_size as usize == tile_key as usize)
            .map(|row| row.raw.clone())
            .collect::<Vec<_>>();
        package.update_archive(archive_name, |archive| {
            let object = archive.object_mut(tile_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers tile object {tile_id} is missing"))
            })?;
            let message_index = unique_tile_message_index(object, tile_id)?;
            let message_type = object.messages[message_index].type_;
            let original = object.messages[message_index].data.as_slice();
            let previous = Tile::decode(original)?;
            let mut current = previous.clone();
            current.row_infos = desired
                .iter()
                .map(|raw| TileRowInfo::decode(raw.as_slice()).map_err(Error::from))
                .collect::<Result<Vec<_>>>()?;
            current.numrows = current
                .row_infos
                .iter()
                .map(|row| row.tile_row_index + 1)
                .max()
                .unwrap_or(0);
            let mut data = patch_varint_field(original, 4, true, Some(u64::from(current.numrows)))?;
            data = rewrite_repeated_length_delimited_fields(&data, 5, &desired)?;
            if Tile::decode(data.as_slice())? != current {
                return Err(Error::InvalidFormat(
                    "Numbers tile sort failed wire validation".to_owned(),
                ));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;
    }
    Ok(())
}

pub(super) fn reorder_body_row_headers(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    body_start: usize,
    destinations_by_source: &[usize],
) -> Result<()> {
    let rows = model.number_of_rows as usize;
    let buckets = &model.base_data_store.row_headers.buckets;
    let mut headers_to_write = Vec::new();
    let mut bucket_identifiers = HashSet::with_capacity(buckets.len());
    for reference in buckets {
        if !bucket_identifiers.insert(reference.identifier) {
            return Err(Error::InvalidFormat(format!(
                "Numbers row-header bucket object {} appears more than once",
                reference.identifier
            )));
        }
        let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers row-header bucket object {} is missing",
                reference.identifier
            ))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers row-header bucket object {} is missing",
                reference.identifier
            ))
        })?;
        let message = unique_header_bucket_message(object, reference.identifier)?;
        let bucket = tst::HeaderStorageBucket::decode(message.data.as_slice())?;
        let raw_headers = repeated_length_delimited_payloads(&message.data, 2)?;
        if raw_headers.len() != bucket.headers.len() {
            return Err(Error::InvalidFormat(
                "Numbers header wire count is inconsistent".to_owned(),
            ));
        }
        let mut source_headers = HashSet::with_capacity(bucket.headers.len());
        for (header, raw) in bucket.headers.iter().zip(raw_headers) {
            if tst::header_storage_bucket::Header::decode(raw)? != *header
                || !source_headers.insert(header.index)
            {
                return Err(Error::InvalidFormat(
                    "Numbers header payload is inconsistent".to_owned(),
                ));
            }
            let source = header.index as usize;
            if source >= rows {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table stores row-header metadata outside row {source}"
                )));
            }
            let destination = relocated_row(source, body_start, destinations_by_source)?;
            headers_to_write.push(RelocatedHeader {
                index: destination,
                raw: patch_varint_field(
                    raw,
                    1,
                    true,
                    Some(u64::try_from(destination).map_err(|_| {
                        Error::ParseError("Numbers header row exceeds u64".to_owned())
                    })?),
                )?,
            });
        }
    }
    headers_to_write.sort_by_key(|header| header.index);
    if headers_to_write
        .windows(2)
        .any(|headers| headers[0].index == headers[1].index)
    {
        return Err(Error::InvalidFormat(
            "Numbers sort would create duplicate row-header metadata".to_owned(),
        ));
    }
    if headers_to_write
        .iter()
        .any(|header| header.index / HEADER_BUCKET_ROWS >= buckets.len())
    {
        return Err(Error::ParseError(
            "Cannot execute Numbers sort: relocated row metadata would require an unallocated header bucket"
                .to_owned(),
        ));
    }

    for (bucket_index, reference) in buckets.iter().enumerate() {
        let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers row-header bucket object {} is missing",
                reference.identifier
            ))
        })?;
        let desired = headers_to_write
            .iter()
            .filter(|header| header.index / HEADER_BUCKET_ROWS == bucket_index)
            .map(|header| header.raw.clone())
            .collect::<Vec<_>>();
        package.update_archive(archive_name, |archive| {
            let object = archive.object_mut(reference.identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers row-header bucket object {} is missing",
                    reference.identifier
                ))
            })?;
            let message_index = unique_header_bucket_message_index(object, reference.identifier)?;
            let message_type = object.messages[message_index].type_;
            let original = object.messages[message_index].data.as_slice();
            let previous = tst::HeaderStorageBucket::decode(original)?;
            let mut current = previous.clone();
            current.headers = desired
                .iter()
                .map(|raw| {
                    tst::header_storage_bucket::Header::decode(raw.as_slice()).map_err(Error::from)
                })
                .collect::<Result<Vec<_>>>()?;
            let data = rewrite_repeated_length_delimited_fields(original, 2, &desired)?;
            if tst::HeaderStorageBucket::decode(data.as_slice())? != current {
                return Err(Error::InvalidFormat(
                    "Numbers row-header sort failed wire validation".to_owned(),
                ));
            }
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
            Ok(())
        })?;
    }
    Ok(())
}

pub(super) fn reorder_row_uids(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    rows: usize,
    body_start: usize,
    destinations_by_source: &[usize],
) -> Result<()> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers UID map object {identifier} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers UID map object {identifier} is missing"))
        })?;
        let message_index = unique_uid_map_message_index(object, identifier)?;
        let message_type = object.messages[message_index].type_;
        let original = object.messages[message_index].data.as_slice();
        let previous = tst::ColumnRowUidMapArchive::decode(original)?;
        validate_row_uid_map(&previous, rows)?;

        let mut current = previous.clone();
        current.row_uid_for_index = previous.row_uid_for_index.clone();
        for (source_offset, destination) in destinations_by_source.iter().copied().enumerate() {
            let source = body_start
                .checked_add(source_offset)
                .ok_or_else(|| Error::ParseError("Numbers UID source row overflow".to_owned()))?;
            current.row_uid_for_index[destination] = previous.row_uid_for_index[source];
        }
        current.row_index_for_uid = vec![0; rows];
        for (index, uid) in current.row_uid_for_index.iter().copied().enumerate() {
            let uid = usize::try_from(uid)
                .map_err(|_| Error::InvalidFormat("Numbers row UID exceeds usize".to_owned()))?;
            current.row_index_for_uid[uid] = u32::try_from(index)
                .map_err(|_| Error::ParseError("Numbers row index exceeds u32".to_owned()))?;
        }
        let data = rewrite_uid_map_wire(original, &previous, &current)?;
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

fn relocated_row(
    source: usize,
    body_start: usize,
    destinations_by_source: &[usize],
) -> Result<usize> {
    let Some(source_offset) = source.checked_sub(body_start) else {
        return Ok(source);
    };
    if source_offset >= destinations_by_source.len() {
        return Ok(source);
    }
    destinations_by_source
        .get(source_offset)
        .copied()
        .ok_or_else(|| Error::InvalidFormat("Numbers sort plan has no row destination".to_owned()))
}

fn validate_row_uid_map(map: &tst::ColumnRowUidMapArchive, rows: usize) -> Result<()> {
    if map.sorted_row_uids.len() != rows
        || map.row_index_for_uid.len() != rows
        || map.row_uid_for_index.len() != rows
    {
        return Err(Error::InvalidFormat(
            "Numbers row UID map lengths do not match table dimensions".to_owned(),
        ));
    }
    let mut uuid_values = HashSet::with_capacity(rows);
    if map
        .sorted_row_uids
        .iter()
        .any(|uuid| !uuid_values.insert((uuid.lower, uuid.upper)))
    {
        return Err(Error::InvalidFormat(
            "Numbers row UID map contains duplicate UUIDs".to_owned(),
        ));
    }
    let mut index_for_uid = vec![None; rows];
    for (index, uid) in map.row_uid_for_index.iter().copied().enumerate() {
        let uid = usize::try_from(uid)
            .map_err(|_| Error::InvalidFormat("Numbers row UID exceeds usize".to_owned()))?;
        if uid >= rows || index_for_uid[uid].replace(index).is_some() {
            return Err(Error::InvalidFormat(
                "Numbers row UID map is not a one-to-one row permutation".to_owned(),
            ));
        }
    }
    for (uid, expected) in index_for_uid.into_iter().enumerate() {
        let expected = expected.ok_or_else(|| {
            Error::InvalidFormat(
                "Numbers row UID map does not cover every stable row UID".to_owned(),
            )
        })?;
        if map.row_index_for_uid[uid]
            != u32::try_from(expected)
                .map_err(|_| Error::ParseError("Numbers row index exceeds u32".to_owned()))?
        {
            return Err(Error::InvalidFormat(
                "Numbers row UID map has inconsistent inverse indices".to_owned(),
            ));
        }
    }
    Ok(())
}

fn unique_tile_message(object: &ArchiveObject, identifier: u64) -> Result<&RawMessage> {
    let index = unique_tile_message_index(object, identifier)?;
    Ok(&object.messages[index])
}

fn unique_tile_message_index(object: &ArchiveObject, identifier: u64) -> Result<usize> {
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| Tile::decode(message.data.as_slice()).ok().map(|_| index))
        .collect::<Vec<_>>();
    match indexes.as_slice() {
        [index] => Ok(*index),
        [] => Err(Error::InvalidFormat(format!(
            "Object {identifier} has no Numbers tile payload"
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "Object {identifier} has multiple Numbers tile payloads"
        ))),
    }
}

fn unique_header_bucket_message(object: &ArchiveObject, identifier: u64) -> Result<&RawMessage> {
    let index = unique_header_bucket_message_index(object, identifier)?;
    Ok(&object.messages[index])
}

fn unique_header_bucket_message_index(object: &ArchiveObject, identifier: u64) -> Result<usize> {
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| {
            tst::HeaderStorageBucket::decode(message.data.as_slice())
                .ok()
                .map(|_| index)
        })
        .collect::<Vec<_>>();
    match indexes.as_slice() {
        [index] => Ok(*index),
        [] => Err(Error::InvalidFormat(format!(
            "Object {identifier} has no Numbers header bucket payload"
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "Object {identifier} has multiple Numbers header bucket payloads"
        ))),
    }
}

fn unique_uid_map_message_index(object: &ArchiveObject, identifier: u64) -> Result<usize> {
    let indexes = object
        .messages
        .iter()
        .enumerate()
        .filter_map(|(index, message)| {
            tst::ColumnRowUidMapArchive::decode(message.data.as_slice())
                .ok()
                .map(|_| index)
        })
        .collect::<Vec<_>>();
    match indexes.as_slice() {
        [index] => Ok(*index),
        [] => Err(Error::InvalidFormat(format!(
            "Object {identifier} has no Numbers UID map payload"
        ))),
        _ => Err(Error::InvalidFormat(format!(
            "Object {identifier} has multiple Numbers UID map payloads"
        ))),
    }
}
