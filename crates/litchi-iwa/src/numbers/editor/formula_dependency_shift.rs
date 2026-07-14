//! Formula dependency coordinate shifts for inserted table axes.

use super::*;

mod wire;

use wire::{rewrite_shifted_dependency_tile_wire, rewrite_shifted_formula_owner_wire};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum DependencyAxis {
    Column,
    Row,
}

impl DependencyAxis {
    fn noun(self) -> &'static str {
        match self {
            Self::Column => "column",
            Self::Row => "row",
        }
    }
}

fn checked_shift(value: u32, insertion: u32, what: &str) -> Result<u32> {
    if value < insertion {
        return Ok(value);
    }
    value
        .checked_add(1)
        .ok_or_else(|| Error::ParseError(format!("Numbers {what} overflow")))
}

pub(super) fn shift_formula_dependencies(
    package: &mut IWorkPackage,
    table_info_id: u64,
    axis: DependencyAxis,
    insertion: u32,
) -> Result<()> {
    const COMPONENT: &str = "Index/CalculationEngine.iwa";
    if !package.contains_entry(COMPONENT) {
        return Ok(());
    }
    package.update_archive(COMPONENT, |archive| {
        let Some((owner_id, message_index)) = archive.objects.iter().find_map(|object| {
            object.messages.iter().enumerate().find_map(|(index, message)| {
                if message.type_ != 4008 {
                    return None;
                }
                let owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())
                    .ok()?;
                (owner.formula_owner.as_ref().map(|reference| reference.identifier)
                    == Some(table_info_id))
                .then_some((object.archive_info.identifier?, index))
            })
        }) else {
            return Err(Error::InvalidFormat(format!(
                "Numbers table info {table_info_id} has no formula dependency owner"
            )));
        };
        let owner_object = archive.object(owner_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers formula owner {owner_id} is missing"))
        })?;
        let original = owner_object.messages[message_index].data.clone();
        let previous = tsce::FormulaOwnerDependenciesArchive::decode(original.as_slice())?;
        validate_shiftable_formula_owner(&previous, axis)?;
        let internal_owner_id = previous.internal_formula_owner_id;
        reject_incoming_dependencies(archive, owner_id, internal_owner_id, axis)?;
        let mut current = previous.clone();
        shift_spanning_ranges(&mut current.spanning_column_dependencies, axis)?;
        shift_spanning_ranges(&mut current.spanning_row_dependencies, axis)?;
        if let Some(dependencies) = &mut current.cell_dependencies {
            for record in &mut dependencies.cell_record {
                shift_dependency_record(record, axis, insertion, internal_owner_id)?;
            }
        }

        let tile_ids = previous
            .tiled_cell_dependencies
            .as_ref()
            .map(|dependencies| {
                dependencies
                    .cell_record_tiles
                    .iter()
                    .map(|reference| reference.identifier)
                    .collect::<Vec<_>>()
            })
            .unwrap_or_default();
        for tile_id in tile_ids {
            let object = archive.object_mut(tile_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers formula dependency tile {tile_id} is missing"
                ))
            })?;
            let tile_message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == 4009)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers formula dependency tile {tile_id} has no payload"
                    ))
                })?;
            let tile_original = object.messages[tile_message_index].data.clone();
            let tile_previous = tsce::CellRecordTileArchive::decode(tile_original.as_slice())?;
            if tile_previous.internal_owner_id != internal_owner_id {
                return Err(Error::InvalidFormat(format!(
                    "Numbers dependency tile {tile_id} belongs to another owner"
                )));
            }
            let mut tile_current = tile_previous.clone();
            for record in &mut tile_current.cell_records {
                shift_dependency_record(record, axis, insertion, internal_owner_id)?;
                let expected_begin = record.row / FORMULA_DEPENDENCY_TILE_ROWS
                    * FORMULA_DEPENDENCY_TILE_ROWS;
                if expected_begin != tile_current.tile_row_begin {
                    return Err(Error::ParseError(format!(
                        "Cannot insert Numbers {}: formula dependency at row {} would cross a dependency-tile boundary",
                        axis.noun(), record.row
                    )));
                }
            }
            let data = rewrite_shifted_dependency_tile_wire(
                &tile_original,
                &tile_previous,
                &tile_current,
                axis,
            )?;
            object.replace_message(
                tile_message_index,
                RawMessage {
                    type_: 4009,
                    data,
                },
            )?;
        }

        let data = rewrite_shifted_formula_owner_wire(&original, &previous, &current, axis)?;
        let object = archive.object_mut(owner_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers formula owner {owner_id} is missing"))
        })?;
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

fn reject_incoming_dependencies(
    archive: &Archive,
    target_object_id: u64,
    target_internal_owner_id: u32,
    axis: DependencyAxis,
) -> Result<()> {
    for object in &archive.objects {
        for message in &object.messages {
            if message.type_ != 4008 {
                continue;
            }
            let owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
            if object.archive_info.identifier == Some(target_object_id) {
                continue;
            }
            validate_shiftable_formula_owner(&owner, axis)?;
            if owner
                .cell_dependencies
                .as_ref()
                .is_some_and(|dependencies| {
                    dependencies
                        .cell_record
                        .iter()
                        .any(|record| record_has_external_owner(record, target_internal_owner_id))
                })
            {
                return Err(incoming_dependency_error(axis));
            }
            for reference in owner
                .tiled_cell_dependencies
                .as_ref()
                .into_iter()
                .flat_map(|dependencies| &dependencies.cell_record_tiles)
            {
                let tile_object = archive.object(reference.identifier).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers formula dependency tile {} is missing",
                        reference.identifier
                    ))
                })?;
                let tile_message = tile_object
                    .messages
                    .iter()
                    .find(|message| message.type_ == 4009)
                    .ok_or_else(|| {
                        Error::InvalidFormat(format!(
                            "Numbers formula dependency tile {} has no payload",
                            reference.identifier
                        ))
                    })?;
                let tile = tsce::CellRecordTileArchive::decode(tile_message.data.as_slice())?;
                if tile
                    .cell_records
                    .iter()
                    .any(|record| record_has_external_owner(record, target_internal_owner_id))
                {
                    return Err(incoming_dependency_error(axis));
                }
            }
        }
    }
    Ok(())
}

fn record_has_external_owner(
    record: &tsce::CellRecordExpandedArchive,
    target_internal_owner_id: u32,
) -> bool {
    record.expanded_edges.as_ref().is_some_and(|edges| {
        edges
            .internal_owner_id_for_edge
            .contains(&target_internal_owner_id)
    })
}

fn incoming_dependency_error(axis: DependencyAxis) -> Error {
    Error::ParseError(format!(
        "Cannot yet insert a Numbers {} referenced by formulas in another table",
        axis.noun()
    ))
}

fn validate_shiftable_formula_owner(
    owner: &tsce::FormulaOwnerDependenciesArchive,
    axis: DependencyAxis,
) -> Result<()> {
    let range_dependencies = owner
        .range_dependencies
        .as_ref()
        .is_some_and(|dependencies| !dependencies.back_dependency.is_empty());
    let volatile_dependencies = owner
        .volatile_dependencies
        .as_ref()
        .is_some_and(volatile_dependencies_are_populated);
    let spanning_dependencies = owner
        .spanning_column_dependencies
        .as_ref()
        .is_some_and(|dependencies| !dependencies.coord_refers_to_spans.is_empty())
        || owner
            .spanning_row_dependencies
            .as_ref()
            .is_some_and(|dependencies| !dependencies.coord_refers_to_spans.is_empty());
    let whole_owner_dependencies = owner
        .whole_owner_dependencies
        .as_ref()
        .and_then(|dependencies| dependencies.dependent_cells.as_ref())
        .is_some_and(|cells| !cells.owner_entries.is_empty());
    let cell_errors = owner
        .cell_errors
        .as_ref()
        .is_some_and(|errors| !errors.errors.is_empty() || !errors.enhanced_errors.is_empty());
    let uuid_references = owner.uuid_references.as_ref().is_some_and(|references| {
        !references.table_refs.is_empty() || !references.table_uuid_refs.is_empty()
    });
    let tiled_ranges = owner
        .tiled_range_dependencies
        .as_ref()
        .is_some_and(|dependencies| !dependencies.range_precedents_tile.is_empty());
    let spills = owner
        .spill_range_sizes
        .as_ref()
        .is_some_and(|spills| !spills.spills.is_empty());
    if range_dependencies
        || volatile_dependencies
        || spanning_dependencies
        || whole_owner_dependencies
        || cell_errors
        || uuid_references
        || tiled_ranges
        || spills
    {
        return Err(Error::ParseError(format!(
            "Cannot yet insert a {} into a Numbers table with advanced range, volatile, spanning, error, UUID, or spill dependency state",
            axis.noun()
        )));
    }
    Ok(())
}

fn volatile_dependencies_are_populated(
    dependencies: &tsce::VolatileDependenciesExpandedArchive,
) -> bool {
    let populated_cells = [
        dependencies.volatile_time_cells.as_ref(),
        dependencies.volatile_random_cells.as_ref(),
        dependencies.volatile_locale_cells.as_ref(),
        dependencies.volatile_sheet_table_name_cells.as_ref(),
        dependencies.volatile_remote_data_cells.as_ref(),
    ]
    .into_iter()
    .flatten()
    .any(|cells| !cells.column_entries.is_empty());
    let populated_refs = dependencies
        .volatile_geometry_cell_refs
        .as_ref()
        .is_some_and(|cells| !cells.owner_entries.is_empty());
    populated_cells || populated_refs
}

fn shift_spanning_ranges(
    dependencies: &mut Option<tsce::SpanningDependenciesExpandedArchive>,
    axis: DependencyAxis,
) -> Result<()> {
    let Some(dependencies) = dependencies else {
        return Ok(());
    };
    for range in [
        dependencies.total_range_for_table.as_mut(),
        dependencies.body_range_for_table.as_mut(),
    ]
    .into_iter()
    .flatten()
    {
        match axis {
            DependencyAxis::Column => {
                range.bottom_right_column =
                    range.bottom_right_column.checked_add(1).ok_or_else(|| {
                        Error::ParseError(
                            "Numbers formula dependency range column overflow".to_owned(),
                        )
                    })?;
            },
            DependencyAxis::Row => {
                range.bottom_right_row =
                    range.bottom_right_row.checked_add(1).ok_or_else(|| {
                        Error::ParseError(
                            "Numbers formula dependency range row overflow".to_owned(),
                        )
                    })?;
            },
        }
    }
    Ok(())
}

fn shift_dependency_record(
    record: &mut tsce::CellRecordExpandedArchive,
    axis: DependencyAxis,
    insertion: u32,
    internal_owner_id: u32,
) -> Result<()> {
    match axis {
        DependencyAxis::Column => {
            record.column = checked_shift(record.column, insertion, "formula column")?;
        },
        DependencyAxis::Row => {
            record.row = checked_shift(record.row, insertion, "formula row")?;
        },
    }
    let Some(edges) = &mut record.expanded_edges else {
        return Ok(());
    };
    if edges.edge_without_owner_rows.len() != edges.edge_without_owner_columns.len()
        || edges.edge_with_owner_rows.len() != edges.edge_with_owner_columns.len()
        || edges.edge_with_owner_rows.len() != edges.internal_owner_id_for_edge.len()
    {
        return Err(Error::InvalidFormat(
            "Numbers formula dependency edge arrays have inconsistent lengths".to_owned(),
        ));
    }
    let local_coordinates = match axis {
        DependencyAxis::Column => &mut edges.edge_without_owner_columns,
        DependencyAxis::Row => &mut edges.edge_without_owner_rows,
    };
    for coordinate in local_coordinates {
        *coordinate = checked_shift(
            *coordinate,
            insertion,
            match axis {
                DependencyAxis::Column => "local formula precedent column",
                DependencyAxis::Row => "local formula precedent row",
            },
        )?;
    }
    let owned_coordinates = match axis {
        DependencyAxis::Column => &mut edges.edge_with_owner_columns,
        DependencyAxis::Row => &mut edges.edge_with_owner_rows,
    };
    for (coordinate, owner) in owned_coordinates
        .iter_mut()
        .zip(&edges.internal_owner_id_for_edge)
    {
        if *owner == internal_owner_id {
            *coordinate = checked_shift(
                *coordinate,
                insertion,
                match axis {
                    DependencyAxis::Column => "formula precedent column",
                    DependencyAxis::Row => "formula precedent row",
                },
            )?;
        }
    }
    Ok(())
}
