//! Physical tile-column and column-header compaction.

use super::*;

pub(super) fn delete_table_tile_column(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    deletion: usize,
) -> Result<()> {
    let old_columns = model.number_of_columns as usize;
    for reference in &model.base_data_store.tiles.tiles {
        let tile_id = reference.tile.identifier;
        let archive_name = locations.get(&tile_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers tile object {tile_id} is missing"))
        })?;
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
            let original = object.messages[message_index].data.clone();
            let previous = Tile::decode(original.as_slice())?;
            let mut current = previous.clone();
            let mut changed = false;
            for row in &mut current.row_infos {
                let mut cells = split_row(row)?;
                if cells.len() < old_columns {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers row {} has {} cell offsets for {old_columns} table columns",
                        row.tile_row_index,
                        cells.len()
                    )));
                }
                if cells.iter().skip(old_columns).any(Option::is_some) {
                    return Err(Error::InvalidFormat(
                        "Numbers row stores cells beyond the declared table columns".to_owned(),
                    ));
                }
                if cells[deletion].is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "Deleted Numbers column {deletion} still contains a stored cell"
                    )));
                }
                let needs_capacity = cells.len() == old_columns;
                let needs_shift = deletion + 1 < old_columns
                    && cells[deletion + 1..old_columns].iter().any(Option::is_some);
                if !needs_capacity && !needs_shift {
                    continue;
                }
                if needs_shift {
                    for column in deletion..old_columns - 1 {
                        cells[column] = cells[column + 1].take();
                    }
                }
                if needs_capacity {
                    cells.truncate(old_columns - 1);
                }
                rebuild_row(row, &cells)?;
                changed = true;
            }
            if changed {
                let message_type = object.messages[message_index].type_;
                let data = rewrite_tile_wire(&original, &previous, &current)?;
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
            }
            Ok(())
        })?;
    }
    Ok(())
}

pub(super) fn delete_column_headers(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    deletion: usize,
) -> Result<()> {
    if identifier == 0 {
        return Ok(());
    }
    let deletion = u32::try_from(deletion)
        .map_err(|_| Error::ParseError("Numbers column exceeds u32".to_owned()))?;
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers column header object {identifier} is missing"
        ))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers column header object {identifier} is missing"
            ))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| tst::HeaderStorageBucket::decode(message.data.as_slice()).is_ok())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Object {identifier} has no Numbers column header payload"
                ))
            })?;
        let original = object.messages[message_index].data.clone();
        let previous = tst::HeaderStorageBucket::decode(original.as_slice())?;
        let raw_headers = repeated_length_delimited_payloads(&original, 2)?;
        if raw_headers.len() != previous.headers.len() {
            return Err(Error::InvalidFormat(
                "Numbers column header wire count is inconsistent".to_owned(),
            ));
        }
        let mut replacements = Vec::with_capacity(previous.headers.len().saturating_sub(1));
        for (header, raw) in previous.headers.iter().zip(raw_headers) {
            if tst::header_storage_bucket::Header::decode(raw)? != *header {
                return Err(Error::InvalidFormat(
                    "Numbers column header payload is inconsistent".to_owned(),
                ));
            }
            let index = match header.index.cmp(&deletion) {
                std::cmp::Ordering::Less => header.index,
                std::cmp::Ordering::Equal => continue,
                std::cmp::Ordering::Greater => header.index - 1,
            };
            replacements.push((
                index,
                patch_varint_field(raw, 1, true, Some(u64::from(index)))?,
            ));
        }
        replacements.sort_by_key(|(index, _)| *index);
        if replacements.windows(2).any(|pair| pair[0].0 == pair[1].0) {
            return Err(Error::InvalidFormat(
                "Numbers column deletion would create duplicate headers".to_owned(),
            ));
        }
        let desired = replacements
            .into_iter()
            .map(|(_, raw)| raw)
            .collect::<Vec<_>>();
        let mut current = previous.clone();
        current.headers = desired
            .iter()
            .map(|raw| {
                tst::header_storage_bucket::Header::decode(raw.as_slice()).map_err(Error::from)
            })
            .collect::<Result<Vec<_>>>()?;
        let data = rewrite_repeated_length_delimited_fields(&original, 2, &desired)?;
        if tst::HeaderStorageBucket::decode(data.as_slice())? != current {
            return Err(Error::InvalidFormat(
                "Numbers column header deletion failed wire validation".to_owned(),
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
}
