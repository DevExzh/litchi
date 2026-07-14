//! Wire-preserving formula dependency coordinate rewrites.

use super::*;

pub(super) fn rewrite_shifted_formula_owner_wire(
    original: &[u8],
    previous: &tsce::FormulaOwnerDependenciesArchive,
    current: &tsce::FormulaOwnerDependenciesArchive,
    axis: DependencyAxis,
) -> Result<Vec<u8>> {
    let mut immutable = current.clone();
    immutable.cell_dependencies = previous.cell_dependencies.clone();
    immutable.spanning_column_dependencies = previous.spanning_column_dependencies.clone();
    immutable.spanning_row_dependencies = previous.spanning_row_dependencies.clone();
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "Numbers formula owner changed outside row dependencies".to_owned(),
        ));
    }
    let mut data = match (&previous.cell_dependencies, &current.cell_dependencies) {
        (Some(previous), Some(current)) => {
            transform_length_delimited_field(original, 4, |dependencies| {
                rewrite_shifted_dependency_records(
                    dependencies,
                    1,
                    &previous.cell_record,
                    &current.cell_record,
                    axis,
                )
            })?
        },
        (None, None) => original.to_vec(),
        _ => {
            return Err(Error::InvalidFormat(
                "Numbers inline dependency representation changed".to_owned(),
            ));
        },
    };
    data = rewrite_optional_spanning_dependencies(
        &data,
        7,
        previous.spanning_column_dependencies.as_ref(),
        current.spanning_column_dependencies.as_ref(),
        axis,
    )?;
    data = rewrite_optional_spanning_dependencies(
        &data,
        8,
        previous.spanning_row_dependencies.as_ref(),
        current.spanning_row_dependencies.as_ref(),
        axis,
    )?;
    if tsce::FormulaOwnerDependenciesArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers formula owner row shift failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}

fn rewrite_optional_spanning_dependencies(
    data: &[u8],
    field_number: u32,
    previous: Option<&tsce::SpanningDependenciesExpandedArchive>,
    current: Option<&tsce::SpanningDependenciesExpandedArchive>,
    axis: DependencyAxis,
) -> Result<Vec<u8>> {
    match (previous, current) {
        (Some(previous), Some(current)) => {
            transform_length_delimited_field(data, field_number, |spanning| {
                rewrite_spanning_dependencies(spanning, previous, current, axis)
            })
        },
        (None, None) => Ok(data.to_vec()),
        _ => Err(Error::InvalidFormat(
            "Numbers spanning dependency representation changed".to_owned(),
        )),
    }
}

fn rewrite_spanning_dependencies(
    original: &[u8],
    previous: &tsce::SpanningDependenciesExpandedArchive,
    current: &tsce::SpanningDependenciesExpandedArchive,
    axis: DependencyAxis,
) -> Result<Vec<u8>> {
    let mut immutable = current.clone();
    immutable.total_range_for_table = previous.total_range_for_table;
    immutable.body_range_for_table = previous.body_range_for_table;
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "Numbers spanning dependencies changed outside table ranges".to_owned(),
        ));
    }
    let mut data = rewrite_optional_range_coordinate(
        original,
        2,
        previous.total_range_for_table.as_ref(),
        current.total_range_for_table.as_ref(),
        axis,
    )?;
    data = rewrite_optional_range_coordinate(
        &data,
        3,
        previous.body_range_for_table.as_ref(),
        current.body_range_for_table.as_ref(),
        axis,
    )?;
    if tsce::SpanningDependenciesExpandedArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers spanning dependency row shift failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}

fn rewrite_optional_range_coordinate(
    data: &[u8],
    field_number: u32,
    previous: Option<&tsce::RangeCoordinateArchive>,
    current: Option<&tsce::RangeCoordinateArchive>,
    axis: DependencyAxis,
) -> Result<Vec<u8>> {
    match (previous, current) {
        (Some(previous), Some(current)) => {
            let mut immutable = *current;
            match axis {
                DependencyAxis::Column => {
                    immutable.bottom_right_column = previous.bottom_right_column;
                },
                DependencyAxis::Row => immutable.bottom_right_row = previous.bottom_right_row,
            }
            if immutable != *previous {
                return Err(Error::InvalidFormat(
                    "Numbers dependency range changed outside its final coordinate".to_owned(),
                ));
            }
            transform_length_delimited_field(data, field_number, |range| {
                let (field, value) = match axis {
                    DependencyAxis::Column => (3, current.bottom_right_column),
                    DependencyAxis::Row => (4, current.bottom_right_row),
                };
                let patched = patch_varint_field(range, field, true, Some(u64::from(value)))?;
                if tsce::RangeCoordinateArchive::decode(patched.as_slice())? != *current {
                    return Err(Error::InvalidFormat(
                        "Numbers dependency range row shift failed wire validation".to_owned(),
                    ));
                }
                Ok(patched)
            })
        },
        (None, None) => Ok(data.to_vec()),
        _ => Err(Error::InvalidFormat(
            "Numbers dependency range representation changed".to_owned(),
        )),
    }
}

pub(super) fn rewrite_shifted_dependency_tile_wire(
    original: &[u8],
    previous: &tsce::CellRecordTileArchive,
    current: &tsce::CellRecordTileArchive,
    axis: DependencyAxis,
) -> Result<Vec<u8>> {
    let mut immutable = current.clone();
    immutable.cell_records = previous.cell_records.clone();
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "Numbers dependency tile identity changed during row insertion".to_owned(),
        ));
    }
    let data = rewrite_shifted_dependency_records(
        original,
        4,
        &previous.cell_records,
        &current.cell_records,
        axis,
    )?;
    if tsce::CellRecordTileArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers dependency tile row shift failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}

fn rewrite_shifted_dependency_records(
    data: &[u8],
    field_number: u32,
    previous: &[tsce::CellRecordExpandedArchive],
    current: &[tsce::CellRecordExpandedArchive],
    axis: DependencyAxis,
) -> Result<Vec<u8>> {
    if previous.len() != current.len() {
        return Err(Error::InvalidFormat(
            "Numbers dependency record count changed during row insertion".to_owned(),
        ));
    }
    let raw = repeated_length_delimited_payloads(data, field_number)?;
    if raw.len() != previous.len() {
        return Err(Error::InvalidFormat(
            "Numbers dependency wire count is inconsistent".to_owned(),
        ));
    }
    let replacements = previous
        .iter()
        .zip(current)
        .zip(raw)
        .map(|((previous, current), raw)| {
            if tsce::CellRecordExpandedArchive::decode(raw)? != *previous {
                return Err(Error::InvalidFormat(
                    "Numbers dependency record changed during row insertion".to_owned(),
                ));
            }
            let mut immutable = current.clone();
            match axis {
                DependencyAxis::Column => immutable.column = previous.column,
                DependencyAxis::Row => immutable.row = previous.row,
            }
            immutable.expanded_edges = previous.expanded_edges.clone();
            if immutable != *previous {
                return Err(Error::InvalidFormat(
                    "Numbers dependency record changed outside row coordinates".to_owned(),
                ));
            }
            let (field, value) = match axis {
                DependencyAxis::Column => (1, current.column),
                DependencyAxis::Row => (2, current.row),
            };
            let mut record = patch_varint_field(raw, field, true, Some(u64::from(value)))?;
            match (&previous.expanded_edges, &current.expanded_edges) {
                (Some(previous), Some(current)) => {
                    record = transform_length_delimited_field(&record, 6, |edges| {
                        rewrite_shifted_edges(edges, previous, current, axis)
                    })?;
                },
                (None, None) => {},
                _ => {
                    return Err(Error::InvalidFormat(
                        "Numbers dependency edge representation changed".to_owned(),
                    ));
                },
            }
            Ok(record)
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, field_number, &replacements)
}

fn rewrite_shifted_edges(
    original: &[u8],
    previous: &tsce::ExpandedEdgesArchive,
    current: &tsce::ExpandedEdgesArchive,
    axis: DependencyAxis,
) -> Result<Vec<u8>> {
    let mut immutable = current.clone();
    match axis {
        DependencyAxis::Column => {
            immutable.edge_without_owner_columns = previous.edge_without_owner_columns.clone();
            immutable.edge_with_owner_columns = previous.edge_with_owner_columns.clone();
        },
        DependencyAxis::Row => {
            immutable.edge_without_owner_rows = previous.edge_without_owner_rows.clone();
            immutable.edge_with_owner_rows = previous.edge_with_owner_rows.clone();
        },
    }
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "Numbers formula edges changed outside row coordinates".to_owned(),
        ));
    }
    let (local_field, local_coordinates, owned_field, owned_coordinates) = match axis {
        DependencyAxis::Column => (
            2,
            &current.edge_without_owner_columns,
            4,
            &current.edge_with_owner_columns,
        ),
        DependencyAxis::Row => (
            1,
            &current.edge_without_owner_rows,
            3,
            &current.edge_with_owner_rows,
        ),
    };
    let mut data = rewrite_repeated_varint_fields(
        original,
        local_field,
        &local_coordinates
            .iter()
            .copied()
            .map(u64::from)
            .collect::<Vec<_>>(),
    )?;
    data = rewrite_repeated_varint_fields(
        &data,
        owned_field,
        &owned_coordinates
            .iter()
            .copied()
            .map(u64::from)
            .collect::<Vec<_>>(),
    )?;
    if tsce::ExpandedEdgesArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers formula edge row shift failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}
