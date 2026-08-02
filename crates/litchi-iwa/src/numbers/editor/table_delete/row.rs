//! Physical tile-row and row-header compaction.

use super::*;

#[derive(Debug)]
struct StoredRow {
    global_row: u32,
    raw: Vec<u8>,
}

#[derive(Debug)]
struct StoredHeader {
    index: u32,
    raw: Vec<u8>,
}

fn deleted_coordinate(value: u32, deletion: u32) -> Option<u32> {
    match value.cmp(&deletion) {
        std::cmp::Ordering::Less => Some(value),
        std::cmp::Ordering::Equal => None,
        std::cmp::Ordering::Greater => Some(value - 1),
    }
}

pub(super) fn delete_table_tile_row(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    deletion: usize,
) -> Result<()> {
    let tile_size = model.base_data_store.tiles.tile_size.unwrap_or(256);
    if tile_size == 0 {
        return Err(Error::InvalidFormat(
            "Numbers table declares a zero tile size".to_owned(),
        ));
    }
    let deletion = u32::try_from(deletion)
        .map_err(|_| Error::ParseError("Numbers row exceeds u32".to_owned()))?;
    let mut rows = Vec::new();
    let mut tile_keys = HashMap::new();
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
        let message = object
            .messages
            .iter()
            .find(|message| Tile::decode(message.data.as_slice()).is_ok())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {} has no Numbers tile payload",
                    reference.tile.identifier
                ))
            })?;
        let tile = Tile::decode(message.data.as_slice())?;
        let raw_rows = repeated_length_delimited_payloads(&message.data, 5)?;
        if raw_rows.len() != tile.row_infos.len() {
            return Err(Error::InvalidFormat(
                "Numbers tile row wire count is inconsistent".to_owned(),
            ));
        }
        for (row_info, raw) in tile.row_infos.iter().zip(raw_rows) {
            if TileRowInfo::decode(raw)? != *row_info || row_info.tile_row_index >= tile_size {
                return Err(Error::InvalidFormat(
                    "Numbers tile contains an invalid row payload".to_owned(),
                ));
            }
            let global = reference
                .tileid
                .checked_mul(tile_size)
                .and_then(|base| base.checked_add(row_info.tile_row_index))
                .ok_or_else(|| Error::ParseError("Numbers tile row overflow".to_owned()))?;
            let Some(global_row) = deleted_coordinate(global, deletion) else {
                if split_row(row_info)?.iter().any(Option::is_some) {
                    return Err(Error::InvalidFormat(format!(
                        "Deleted Numbers row {deletion} still contains stored cells"
                    )));
                }
                continue;
            };
            rows.push(StoredRow {
                global_row,
                raw: patch_varint_field(raw, 1, true, Some(u64::from(global_row % tile_size)))?,
            });
        }
    }
    rows.sort_by_key(|row| row.global_row);
    if rows
        .windows(2)
        .any(|pair| pair[0].global_row == pair[1].global_row)
    {
        return Err(Error::InvalidFormat(
            "Numbers row deletion would create duplicate stored rows".to_owned(),
        ));
    }
    for row in &rows {
        let key = row.global_row / tile_size;
        if !tile_keys.contains_key(&key) {
            return Err(Error::ParseError(format!(
                "Cannot delete Numbers row: stored row {} would require an unallocated tile",
                row.global_row
            )));
        }
    }

    for (tile_key, tile_id) in tile_keys {
        let archive_name = locations.get(&tile_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers tile object {tile_id} is missing"))
        })?;
        let desired = rows
            .iter()
            .filter(|row| row.global_row / tile_size == tile_key)
            .map(|row| row.raw.clone())
            .collect::<Vec<_>>();
        package.update_archive(archive_name, |archive| {
            let object = archive.object_mut(tile_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers tile object {tile_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| Tile::decode(message.data.as_slice()).is_ok())
                .ok_or_else(|| {
                    Error::InvalidFormat(format!("Object {tile_id} has no Numbers tile payload"))
                })?;
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
                    "Numbers tile row deletion failed wire validation".to_owned(),
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

pub(super) fn delete_row_headers(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    deletion: usize,
) -> Result<()> {
    let deletion = u32::try_from(deletion)
        .map_err(|_| Error::ParseError("Numbers row exceeds u32".to_owned()))?;
    let mut headers = Vec::new();
    for reference in &model.base_data_store.row_headers.buckets {
        let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers header bucket object {} is missing",
                reference.identifier
            ))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers header bucket object {} is missing",
                reference.identifier
            ))
        })?;
        let message = object
            .messages
            .iter()
            .find(|message| tst::HeaderStorageBucket::decode(message.data.as_slice()).is_ok())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {} has no Numbers header bucket payload",
                    reference.identifier
                ))
            })?;
        let bucket = tst::HeaderStorageBucket::decode(message.data.as_slice())?;
        let raw_headers = repeated_length_delimited_payloads(&message.data, 2)?;
        if raw_headers.len() != bucket.headers.len() {
            return Err(Error::InvalidFormat(
                "Numbers header wire count is inconsistent".to_owned(),
            ));
        }
        for (header, raw) in bucket.headers.iter().zip(raw_headers) {
            if tst::header_storage_bucket::Header::decode(raw)? != *header {
                return Err(Error::InvalidFormat(
                    "Numbers header payload is inconsistent".to_owned(),
                ));
            }
            let Some(index) = deleted_coordinate(header.index, deletion) else {
                continue;
            };
            headers.push(StoredHeader {
                index,
                raw: patch_varint_field(raw, 1, true, Some(u64::from(index)))?,
            });
        }
    }
    headers.sort_by_key(|header| header.index);
    if headers
        .windows(2)
        .any(|pair| pair[0].index == pair[1].index)
    {
        return Err(Error::InvalidFormat(
            "Numbers row deletion would create duplicate headers".to_owned(),
        ));
    }

    for (bucket_index, reference) in model.base_data_store.row_headers.buckets.iter().enumerate() {
        let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers header bucket object {} is missing",
                reference.identifier
            ))
        })?;
        let desired = headers
            .iter()
            .filter(|header| header.index as usize / HEADER_BUCKET_ROWS == bucket_index)
            .map(|header| header.raw.clone())
            .collect::<Vec<_>>();
        package.update_archive(archive_name, |archive| {
            let object = archive.object_mut(reference.identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers header bucket object {} is missing",
                    reference.identifier
                ))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| {
                    tst::HeaderStorageBucket::decode(message.data.as_slice()).is_ok()
                })
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Object {} has no Numbers header bucket payload",
                        reference.identifier
                    ))
                })?;
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
                    "Numbers header row deletion failed wire validation".to_owned(),
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
