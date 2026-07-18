//! Physical cell, header, UID, stroke, and dimension storage shifts.

use super::*;

fn checked_shift(value: u32, insertion: u32, what: &str) -> Result<u32> {
    if value < insertion {
        return Ok(value);
    }
    value
        .checked_add(1)
        .ok_or_else(|| Error::ParseError(format!("Numbers {what} overflow")))
}

pub(super) fn shift_table_tile_columns(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    insertion: usize,
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
                let needs_capacity = cells.len() == old_columns;
                let needs_shift = insertion < old_columns
                    && cells[insertion..old_columns].iter().any(Option::is_some);
                if !needs_capacity && !needs_shift {
                    continue;
                }
                if needs_capacity {
                    cells.resize(old_columns + 1, None);
                }
                if needs_shift {
                    for column in (insertion..old_columns).rev() {
                        cells[column + 1] = cells[column].take();
                    }
                    cells[insertion] = None;
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

pub(super) fn shift_column_headers(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    insertion: usize,
) -> Result<()> {
    if identifier == 0 {
        return Ok(());
    }
    let insertion = u32::try_from(insertion)
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
        let mut replacements = previous
            .headers
            .iter()
            .zip(raw_headers)
            .map(|(header, raw)| {
                if tst::header_storage_bucket::Header::decode(raw)? != *header {
                    return Err(Error::InvalidFormat(
                        "Numbers column header payload is inconsistent".to_owned(),
                    ));
                }
                let index = checked_shift(header.index, insertion, "column header")?;
                Ok((
                    index,
                    patch_varint_field(raw, 1, true, Some(u64::from(index)))?,
                ))
            })
            .collect::<Result<Vec<_>>>()?;
        replacements.sort_by_key(|(index, _)| *index);
        if replacements
            .windows(2)
            .any(|entries| entries[0].0 == entries[1].0)
        {
            return Err(Error::InvalidFormat(
                "Numbers column insertion would create duplicate headers".to_owned(),
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
                "Numbers column header insertion failed wire validation".to_owned(),
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

pub(super) fn insert_column_uid(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    old_columns: usize,
    insertion: usize,
) -> Result<()> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers UID map object {identifier} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers UID map object {identifier} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| {
                tst::ColumnRowUidMapArchive::decode(message.data.as_slice()).is_ok()
            })
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Object {identifier} has no UID map payload"))
            })?;
        let original = object.messages[message_index].data.clone();
        let previous = tst::ColumnRowUidMapArchive::decode(original.as_slice())?;
        let mut current = previous.clone();
        if current.sorted_column_uids.len() != old_columns
            || current.column_index_for_uid.len() != old_columns
            || current.column_uid_for_index.len() != old_columns
        {
            return Err(Error::InvalidFormat(
                "Numbers column UID map lengths do not match table dimensions".to_owned(),
            ));
        }
        let insertion_u32 = u32::try_from(insertion)
            .map_err(|_| Error::ParseError("Numbers column exceeds u32".to_owned()))?;
        for index in &mut current.column_index_for_uid {
            if *index >= insertion_u32 {
                *index = index.checked_add(1).ok_or_else(|| {
                    Error::ParseError("Numbers column UID index overflow".to_owned())
                })?;
            }
        }
        let lower = current
            .sorted_column_uids
            .iter()
            .map(|uuid| uuid.lower)
            .max()
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers column UUID overflow".to_owned()))?;
        let upper = current
            .sorted_column_uids
            .iter()
            .map(|uuid| uuid.upper)
            .max()
            .unwrap_or(0)
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers column UUID overflow".to_owned()))?;
        current.sorted_column_uids.push(tsp::Uuid { lower, upper });
        current.column_index_for_uid.push(insertion_u32);
        current.column_uid_for_index.insert(
            insertion,
            u32::try_from(old_columns)
                .map_err(|_| Error::ParseError("Numbers UID index exceeds u32".to_owned()))?,
        );
        let data = rewrite_uid_map_wire(&original, &previous, &current)?;
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

pub(super) fn insert_stroke_column(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    identifier: u64,
    columns: u32,
) -> Result<()> {
    let archive_name = locations.get(&identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers stroke sidecar {identifier} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers stroke sidecar {identifier} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| tst::StrokeSidecarArchive::decode(message.data.as_slice()).is_ok())
            .ok_or_else(|| {
                Error::InvalidFormat(format!("Object {identifier} has no stroke sidecar payload"))
            })?;
        let original = object.messages[message_index].data.clone();
        let previous = tst::StrokeSidecarArchive::decode(original.as_slice())?;
        if !previous.left_column_stroke_layers.is_empty()
            || !previous.right_column_stroke_layers.is_empty()
            || !previous.top_row_stroke_layers.is_empty()
            || !previous.bottom_row_stroke_layers.is_empty()
        {
            return Err(Error::ParseError(
                "Cannot yet insert a Numbers column into a table with explicit stroke layers"
                    .to_owned(),
            ));
        }
        let mut current = previous.clone();
        current.column_count = Some(columns);
        let data = rewrite_stroke_sidecar_wire(&original, &previous, &current)?;
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

pub(super) fn set_table_column_count(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    columns: u32,
) -> Result<()> {
    let archive_name = locations.get(&table_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers table object {table_id} is missing"))
    })?;
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers table object {table_id} is missing"))
        })?;
        let message_index = find_table_model_message(object)?;
        let original = object.messages[message_index].data.clone();
        let data = patch_varint_field(&original, 7, true, Some(u64::from(columns)))?;
        if TableModelArchive::decode(data.as_slice())?.number_of_columns != columns {
            return Err(Error::InvalidFormat(
                "Numbers table column count failed wire validation".to_owned(),
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
