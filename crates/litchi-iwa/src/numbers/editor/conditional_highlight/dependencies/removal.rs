//! Removal of one volatile conditional-style formula host.

use prost::Message;

use super::*;

pub(in crate::numbers::editor::conditional_highlight) fn remove_volatile_host(
    package: &mut IWorkPackage,
    conditional_owner_uid: tsp::Uuid,
    row: usize,
    column: usize,
) -> Result<()> {
    let Some(entry_name) = package.calculation_engine_entry_name()?.map(str::to_owned) else {
        return Ok(());
    };
    let row = i32::try_from(row).map_err(|_| {
        Error::ParseError("conditional-highlight row exceeds the native signed index".to_owned())
    })?;
    let record_row = u32::try_from(row).map_err(|_| {
        Error::ParseError("conditional-highlight row cannot be negative".to_owned())
    })?;
    let record_column = u32::try_from(column)
        .map_err(|_| Error::ParseError("conditional-highlight column exceeds u32".to_owned()))?;
    package.update_archive(&entry_name, |archive| {
        let Some((owner_id, owner_message_index, mut owner)) =
            owner_location(archive, &conditional_owner_uid)?
        else {
            return Ok(());
        };
        let Some(record_index) = owner.cell_dependencies.as_ref().and_then(|dependencies| {
            dependencies
                .cell_record
                .iter()
                .position(|record| record.row == record_row && record.column == record_column)
        }) else {
            return Ok(());
        };
        let owner_data = archive
            .object(owner_id)
            .and_then(|object| object.messages.get(owner_message_index))
            .map(|message| message.data.clone())
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style formula owner disappeared".to_owned(),
                )
            })?;
        if owner.encode_to_vec() != owner_data {
            return Err(Error::InvalidFormat(
                "cannot safely shrink a conditional-style owner with unknown wire fields"
                    .to_owned(),
            ));
        }

        owner
            .cell_dependencies
            .as_mut()
            .expect("a matching dependency record exists")
            .cell_record
            .remove(record_index);
        let volatile_time_cells = owner
            .volatile_dependencies
            .as_mut()
            .and_then(|dependencies| dependencies.volatile_time_cells.as_mut())
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style volatile-time coordinates are missing".to_owned(),
                )
            })?;
        if !remove_coordinate(volatile_time_cells, row, record_column)? {
            return Err(Error::InvalidFormat(
                "Numbers conditional-style volatile-time coordinate is missing".to_owned(),
            ));
        }
        let references = owner.uuid_references.as_mut().ok_or_else(|| {
            Error::InvalidFormat("Numbers conditional-style UUID references are missing".to_owned())
        })?;
        let table_ref_index = references
            .table_refs
            .iter()
            .position(|table_ref| {
                table_ref
                    .coord_set
                    .as_ref()
                    .is_some_and(|coordinates| contains_coordinate(coordinates, row, record_column))
            })
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style UUID host coordinate is missing".to_owned(),
                )
            })?;
        let table_ref = &mut references.table_refs[table_ref_index];
        if !remove_coordinate(
            table_ref
                .coord_set
                .as_mut()
                .expect("the matching table reference has coordinates"),
            row,
            record_column,
        )? {
            return Err(Error::InvalidFormat(
                "Numbers conditional-style UUID host coordinate is missing".to_owned(),
            ));
        }
        if table_ref
            .coord_set
            .as_ref()
            .is_none_or(|coordinates| coordinates.column_entries.is_empty())
        {
            references.table_refs.remove(table_ref_index);
        }
        let tile_id = owner
            .tiled_cell_dependencies
            .as_ref()
            .and_then(|dependencies| dependencies.cell_record_tiles.first())
            .map(|reference| reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style dependency tile is missing".to_owned(),
                )
            })?;

        let tile_object = archive.object_mut(tile_id).ok_or_else(|| {
            Error::InvalidFormat("Numbers conditional-style dependency tile is missing".to_owned())
        })?;
        let tile_message_index = tile_object
            .messages
            .iter()
            .position(|message| message.type_ == CELL_RECORD_TILE_MESSAGE_TYPE)
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style dependency tile payload is missing".to_owned(),
                )
            })?;
        let tile_data = tile_object.messages[tile_message_index].data.clone();
        let mut tile = tsce::CellRecordTileArchive::decode(tile_data.as_slice())?;
        if tile.encode_to_vec() != tile_data {
            return Err(Error::InvalidFormat(
                "cannot safely shrink a conditional-style tile with unknown wire fields".to_owned(),
            ));
        }
        let tile_record_index = tile
            .cell_records
            .iter()
            .position(|record| record.row == record_row && record.column == record_column)
            .ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers conditional-style dependency tile host is missing".to_owned(),
                )
            })?;
        tile.cell_records.remove(tile_record_index);
        tile_object.replace_message(
            tile_message_index,
            RawMessage {
                type_: CELL_RECORD_TILE_MESSAGE_TYPE,
                data: tile.encode_to_vec(),
            },
        )?;

        let owner_object = archive.object_mut(owner_id).ok_or_else(|| {
            Error::InvalidFormat("Numbers conditional-style formula owner disappeared".to_owned())
        })?;
        let owner_message_type = owner_object.messages[owner_message_index].type_;
        owner_object.replace_message(
            owner_message_index,
            RawMessage {
                type_: owner_message_type,
                data: owner.encode_to_vec(),
            },
        )?;

        let (engine_id, engine_message_index) =
            super::super::super::formula_clone::calculation_engine_location(archive)?;
        let engine = archive.object_mut(engine_id).ok_or_else(|| {
            Error::InvalidFormat("Numbers CalculationEngine root is missing".to_owned())
        })?;
        let engine_message = engine.messages[engine_message_index].clone();
        let data = super::super::super::formula_clone::decrement_formula_count_in_engine(
            &engine_message.data,
        )?;
        engine.replace_message(
            engine_message_index,
            RawMessage {
                type_: CALCULATION_ENGINE_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(in crate::numbers::editor::conditional_highlight) fn contains_coordinate(
    coordinates: &tsce::CellCoordSetArchive,
    row: i32,
    column: u32,
) -> bool {
    coordinates.column_entries.iter().any(|entry| {
        entry.column == column
            && entry.row_set.entries.iter().any(|range| {
                (range.range_begin..=range.range_end.unwrap_or(range.range_begin)).contains(&row)
            })
    })
}

fn remove_coordinate(
    coordinates: &mut tsce::CellCoordSetArchive,
    row: i32,
    column: u32,
) -> Result<bool> {
    let Some(column_index) = coordinates
        .column_entries
        .iter()
        .position(|entry| entry.column == column)
    else {
        return Ok(false);
    };
    let entries = &mut coordinates.column_entries[column_index].row_set.entries;
    let Some(range_index) = entries.iter().position(|range| {
        (range.range_begin..=range.range_end.unwrap_or(range.range_begin)).contains(&row)
    }) else {
        return Ok(false);
    };
    let range = entries[range_index];
    let end = range.range_end.unwrap_or(range.range_begin);
    match (range.range_begin == row, end == row) {
        (true, true) => {
            entries.remove(range_index);
        },
        (true, false) => {
            let begin = row.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("Numbers coordinate range overflow".to_owned())
            })?;
            entries[range_index].range_begin = begin;
            entries[range_index].range_end = (begin != end).then_some(end);
        },
        (false, true) => {
            let new_end = row.checked_sub(1).ok_or_else(|| {
                Error::InvalidFormat("Numbers coordinate range underflow".to_owned())
            })?;
            entries[range_index].range_end = (range.range_begin != new_end).then_some(new_end);
        },
        (false, false) => {
            let lower_end = row.checked_sub(1).ok_or_else(|| {
                Error::InvalidFormat("Numbers coordinate range underflow".to_owned())
            })?;
            let upper_begin = row.checked_add(1).ok_or_else(|| {
                Error::InvalidFormat("Numbers coordinate range overflow".to_owned())
            })?;
            entries[range_index].range_end = (range.range_begin != lower_end).then_some(lower_end);
            entries.insert(
                range_index + 1,
                tsce::index_set_archive::IndexSetEntry {
                    range_begin: upper_begin,
                    range_end: (upper_begin != end).then_some(end),
                },
            );
        },
    }
    if entries.is_empty() {
        coordinates.column_entries.remove(column_index);
    }
    Ok(true)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn coordinates(begin: i32, end: i32) -> tsce::CellCoordSetArchive {
        tsce::CellCoordSetArchive {
            column_entries: vec![tsce::cell_coord_set_archive::ColumnEntry {
                column: 4,
                row_set: tsce::IndexSetArchive {
                    entries: vec![tsce::index_set_archive::IndexSetEntry {
                        range_begin: begin,
                        range_end: (begin != end).then_some(end),
                    }],
                },
            }],
        }
    }

    fn ranges(coordinates: &tsce::CellCoordSetArchive) -> Vec<(i32, i32)> {
        coordinates.column_entries[0]
            .row_set
            .entries
            .iter()
            .map(|range| {
                (
                    range.range_begin,
                    range.range_end.unwrap_or(range.range_begin),
                )
            })
            .collect()
    }

    #[test]
    fn coordinate_removal_splits_and_trims_ranges() {
        let mut value = coordinates(1, 5);
        assert!(remove_coordinate(&mut value, 3, 4).unwrap());
        assert_eq!(ranges(&value), vec![(1, 2), (4, 5)]);
        assert!(remove_coordinate(&mut value, 1, 4).unwrap());
        assert_eq!(ranges(&value), vec![(2, 2), (4, 5)]);
        assert!(remove_coordinate(&mut value, 5, 4).unwrap());
        assert_eq!(ranges(&value), vec![(2, 2), (4, 4)]);
        assert!(!remove_coordinate(&mut value, 8, 4).unwrap());
    }

    #[test]
    fn coordinate_removal_drops_empty_columns() {
        let mut value = coordinates(2, 2);
        assert!(remove_coordinate(&mut value, 2, 4).unwrap());
        assert!(value.column_entries.is_empty());
    }
}
