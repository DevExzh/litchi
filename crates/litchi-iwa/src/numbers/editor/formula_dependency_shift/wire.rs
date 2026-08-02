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
    immutable.range_dependencies = previous.range_dependencies.clone();
    immutable.spanning_column_dependencies = previous.spanning_column_dependencies.clone();
    immutable.spanning_row_dependencies = previous.spanning_row_dependencies.clone();
    immutable.uuid_references = previous.uuid_references.clone();
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
    data = rewrite_optional_range_dependencies(
        &data,
        previous.range_dependencies.as_ref(),
        current.range_dependencies.as_ref(),
    )?;
    data = rewrite_optional_spanning_dependencies(
        &data,
        7,
        previous.spanning_column_dependencies.as_ref(),
        current.spanning_column_dependencies.as_ref(),
        axis,
    )?;
    data = rewrite_optional_uuid_references(
        &data,
        previous.uuid_references.as_ref(),
        current.uuid_references.as_ref(),
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

fn rewrite_optional_range_dependencies(
    data: &[u8],
    previous: Option<&tsce::RangeDependenciesArchive>,
    current: Option<&tsce::RangeDependenciesArchive>,
) -> Result<Vec<u8>> {
    match (previous, current) {
        (Some(previous), Some(current)) => {
            transform_length_delimited_field(data, 5, |dependencies| {
                rewrite_range_dependencies(dependencies, previous, current)
            })
        },
        (None, None) => Ok(data.to_vec()),
        _ => Err(Error::InvalidFormat(
            "Numbers range dependency representation changed".to_owned(),
        )),
    }
}

fn rewrite_range_dependencies(
    data: &[u8],
    previous: &tsce::RangeDependenciesArchive,
    current: &tsce::RangeDependenciesArchive,
) -> Result<Vec<u8>> {
    let raw = repeated_length_delimited_payloads(data, 2)?;
    if raw.len() != previous.back_dependency.len()
        || previous.back_dependency.len() != current.back_dependency.len()
    {
        return Err(Error::InvalidFormat(
            "Numbers range dependency wire count is inconsistent".to_owned(),
        ));
    }
    let replacements = previous
        .back_dependency
        .iter()
        .zip(&current.back_dependency)
        .zip(raw)
        .map(|((previous, current), raw)| rewrite_range_back_dependency(raw, previous, current))
        .collect::<Result<Vec<_>>>()?;
    let rewritten = rewrite_repeated_length_delimited_fields(data, 2, &replacements)?;
    if tsce::RangeDependenciesArchive::decode(rewritten.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers range dependency rewrite failed wire validation".to_owned(),
        ));
    }
    Ok(rewritten)
}

fn rewrite_range_back_dependency(
    data: &[u8],
    previous: &tsce::RangeBackDependencyArchive,
    current: &tsce::RangeBackDependencyArchive,
) -> Result<Vec<u8>> {
    if tsce::RangeBackDependencyArchive::decode(data)? != *previous {
        return Err(Error::InvalidFormat(
            "Numbers range dependency wire payload is inconsistent".to_owned(),
        ));
    }
    let mut immutable = current.clone();
    immutable.cell_coord_row = previous.cell_coord_row;
    immutable.cell_coord_column = previous.cell_coord_column;
    immutable.range_reference = previous.range_reference.clone();
    immutable.internal_range_reference = previous.internal_range_reference;
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "Numbers range dependency changed outside coordinates".to_owned(),
        ));
    }
    let mut rewritten = patch_varint_field(data, 1, true, Some(u64::from(current.cell_coord_row)))?;
    rewritten = patch_varint_field(
        &rewritten,
        2,
        true,
        Some(u64::from(current.cell_coord_column)),
    )?;
    rewritten = rewrite_optional_range_reference(
        &rewritten,
        previous.range_reference.as_ref(),
        current.range_reference.as_ref(),
    )?;
    rewritten = rewrite_optional_internal_range_reference(
        &rewritten,
        previous.internal_range_reference.as_ref(),
        current.internal_range_reference.as_ref(),
    )?;
    if tsce::RangeBackDependencyArchive::decode(rewritten.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers range dependency coordinate rewrite failed wire validation".to_owned(),
        ));
    }
    Ok(rewritten)
}

fn rewrite_optional_range_reference(
    data: &[u8],
    previous: Option<&tsce::RangeReferenceArchive>,
    current: Option<&tsce::RangeReferenceArchive>,
) -> Result<Vec<u8>> {
    match (previous, current) {
        (Some(previous), Some(current)) => transform_length_delimited_field(data, 3, |range| {
            let mut immutable = current.clone();
            immutable.top_left_column = previous.top_left_column;
            immutable.top_left_row = previous.top_left_row;
            immutable.bottom_right_column = previous.bottom_right_column;
            immutable.bottom_right_row = previous.bottom_right_row;
            if immutable != *previous {
                return Err(Error::InvalidFormat(
                    "Numbers external range dependency changed outside coordinates".to_owned(),
                ));
            }
            let mut rewritten = range.to_vec();
            for (field, value) in [
                (2, current.top_left_column),
                (3, current.top_left_row),
                (4, current.bottom_right_column),
                (5, current.bottom_right_row),
            ] {
                rewritten = patch_varint_field(&rewritten, field, true, Some(u64::from(value)))?;
            }
            if tsce::RangeReferenceArchive::decode(rewritten.as_slice())? != *current {
                return Err(Error::InvalidFormat(
                    "Numbers external range dependency rewrite failed wire validation".to_owned(),
                ));
            }
            Ok(rewritten)
        }),
        (None, None) => Ok(data.to_vec()),
        _ => Err(Error::InvalidFormat(
            "Numbers external range dependency representation changed".to_owned(),
        )),
    }
}

fn rewrite_optional_internal_range_reference(
    data: &[u8],
    previous: Option<&tsce::InternalRangeReferenceArchive>,
    current: Option<&tsce::InternalRangeReferenceArchive>,
) -> Result<Vec<u8>> {
    match (previous, current) {
        (Some(previous), Some(current)) => transform_length_delimited_field(data, 4, |reference| {
            let mut immutable = *current;
            immutable.range = previous.range;
            if immutable != *previous {
                return Err(Error::InvalidFormat(
                    "Numbers internal range dependency changed outside coordinates".to_owned(),
                ));
            }
            transform_length_delimited_field(reference, 2, |range| {
                rewrite_complete_range_coordinate(range, &previous.range, &current.range)
            })
        }),
        (None, None) => Ok(data.to_vec()),
        _ => Err(Error::InvalidFormat(
            "Numbers internal range dependency representation changed".to_owned(),
        )),
    }
}

fn rewrite_complete_range_coordinate(
    data: &[u8],
    previous: &tsce::RangeCoordinateArchive,
    current: &tsce::RangeCoordinateArchive,
) -> Result<Vec<u8>> {
    if tsce::RangeCoordinateArchive::decode(data)? != *previous {
        return Err(Error::InvalidFormat(
            "Numbers range coordinate wire payload is inconsistent".to_owned(),
        ));
    }
    let mut rewritten = data.to_vec();
    for (field, value) in [
        (1, current.top_left_column),
        (2, current.top_left_row),
        (3, current.bottom_right_column),
        (4, current.bottom_right_row),
    ] {
        rewritten = patch_varint_field(&rewritten, field, true, Some(u64::from(value)))?;
    }
    if tsce::RangeCoordinateArchive::decode(rewritten.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers range coordinate rewrite failed wire validation".to_owned(),
        ));
    }
    Ok(rewritten)
}

fn rewrite_optional_uuid_references(
    data: &[u8],
    previous: Option<&tsce::UuidReferencesArchive>,
    current: Option<&tsce::UuidReferencesArchive>,
) -> Result<Vec<u8>> {
    match (previous, current) {
        (Some(previous), Some(current)) => {
            transform_length_delimited_field(data, 14, |references| {
                rewrite_uuid_references(references, previous, current)
            })
        },
        (None, None) => Ok(data.to_vec()),
        _ => Err(Error::InvalidFormat(
            "Numbers UUID-reference representation changed".to_owned(),
        )),
    }
}

fn rewrite_uuid_references(
    data: &[u8],
    previous: &tsce::UuidReferencesArchive,
    current: &tsce::UuidReferencesArchive,
) -> Result<Vec<u8>> {
    let mut rewritten = rewrite_uuid_table_refs(data, previous, current)?;
    rewritten = rewrite_uuid_table_groups(&rewritten, previous, current)?;
    if tsce::UuidReferencesArchive::decode(rewritten.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers UUID-reference rewrite failed wire validation".to_owned(),
        ));
    }
    Ok(rewritten)
}

fn rewrite_uuid_table_refs(
    data: &[u8],
    previous: &tsce::UuidReferencesArchive,
    current: &tsce::UuidReferencesArchive,
) -> Result<Vec<u8>> {
    let raw = repeated_length_delimited_payloads(data, 1)?;
    if raw.len() != previous.table_refs.len()
        || previous.table_refs.len() != current.table_refs.len()
    {
        return Err(Error::InvalidFormat(
            "Numbers UUID table-reference wire count is inconsistent".to_owned(),
        ));
    }
    let replacements = previous
        .table_refs
        .iter()
        .zip(&current.table_refs)
        .zip(raw)
        .map(|((previous, current), raw)| {
            let mut immutable = current.clone();
            immutable.coord_set = previous.coord_set.clone();
            if immutable != *previous {
                return Err(Error::InvalidFormat(
                    "Numbers UUID table reference changed outside coordinates".to_owned(),
                ));
            }
            rewrite_optional_cell_coord_set(
                raw,
                2,
                previous.coord_set.as_ref(),
                current.coord_set.as_ref(),
            )
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, 1, &replacements)
}

fn rewrite_uuid_table_groups(
    data: &[u8],
    previous: &tsce::UuidReferencesArchive,
    current: &tsce::UuidReferencesArchive,
) -> Result<Vec<u8>> {
    let raw = repeated_length_delimited_payloads(data, 2)?;
    if raw.len() != previous.table_uuid_refs.len()
        || previous.table_uuid_refs.len() != current.table_uuid_refs.len()
    {
        return Err(Error::InvalidFormat(
            "Numbers UUID table-group wire count is inconsistent".to_owned(),
        ));
    }
    let replacements = previous
        .table_uuid_refs
        .iter()
        .zip(&current.table_uuid_refs)
        .zip(raw)
        .map(|((previous, current), raw)| {
            let mut immutable = current.clone();
            immutable.uuid_refs = previous.uuid_refs.clone();
            if immutable != *previous {
                return Err(Error::InvalidFormat(
                    "Numbers UUID table group changed outside references".to_owned(),
                ));
            }
            let nested = repeated_length_delimited_payloads(raw, 2)?;
            if nested.len() != previous.uuid_refs.len()
                || previous.uuid_refs.len() != current.uuid_refs.len()
            {
                return Err(Error::InvalidFormat(
                    "Numbers UUID-reference wire count is inconsistent".to_owned(),
                ));
            }
            let nested_replacements = previous
                .uuid_refs
                .iter()
                .zip(&current.uuid_refs)
                .zip(nested)
                .map(|((previous, current), raw)| {
                    let mut immutable = current.clone();
                    immutable.coord_set = previous.coord_set.clone();
                    if immutable != *previous {
                        return Err(Error::InvalidFormat(
                            "Numbers UUID reference changed outside coordinates".to_owned(),
                        ));
                    }
                    rewrite_optional_cell_coord_set(
                        raw,
                        2,
                        previous.coord_set.as_ref(),
                        current.coord_set.as_ref(),
                    )
                })
                .collect::<Result<Vec<_>>>()?;
            rewrite_repeated_length_delimited_fields(raw, 2, &nested_replacements)
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, 2, &replacements)
}

fn rewrite_optional_cell_coord_set(
    data: &[u8],
    field_number: u32,
    previous: Option<&tsce::CellCoordSetArchive>,
    current: Option<&tsce::CellCoordSetArchive>,
) -> Result<Vec<u8>> {
    match (previous, current) {
        (Some(previous), Some(current)) => {
            transform_length_delimited_field(data, field_number, |coordinates| {
                rewrite_cell_coord_set(coordinates, previous, current)
            })
        },
        (None, None) => Ok(data.to_vec()),
        _ => Err(Error::InvalidFormat(
            "Numbers UUID coordinate-set representation changed".to_owned(),
        )),
    }
}

fn rewrite_cell_coord_set(
    data: &[u8],
    previous: &tsce::CellCoordSetArchive,
    current: &tsce::CellCoordSetArchive,
) -> Result<Vec<u8>> {
    let raw = repeated_length_delimited_payloads(data, 1)?;
    if raw.len() != previous.column_entries.len()
        || previous.column_entries.len() != current.column_entries.len()
    {
        return Err(Error::InvalidFormat(
            "Numbers UUID coordinate column count is inconsistent".to_owned(),
        ));
    }
    let replacements = previous
        .column_entries
        .iter()
        .zip(&current.column_entries)
        .zip(raw)
        .map(|((previous, current), raw)| {
            let mut immutable = current.clone();
            immutable.column = previous.column;
            immutable.row_set = previous.row_set.clone();
            if immutable != *previous {
                return Err(Error::InvalidFormat(
                    "Numbers UUID coordinate column changed outside coordinates".to_owned(),
                ));
            }
            let rewritten = patch_varint_field(raw, 1, true, Some(u64::from(current.column)))?;
            transform_length_delimited_field(&rewritten, 2, |rows| {
                rewrite_index_set(rows, &previous.row_set, &current.row_set)
            })
        })
        .collect::<Result<Vec<_>>>()?;
    let rewritten = rewrite_repeated_length_delimited_fields(data, 1, &replacements)?;
    if tsce::CellCoordSetArchive::decode(rewritten.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers UUID coordinate-set rewrite failed wire validation".to_owned(),
        ));
    }
    Ok(rewritten)
}

fn rewrite_index_set(
    data: &[u8],
    previous: &tsce::IndexSetArchive,
    current: &tsce::IndexSetArchive,
) -> Result<Vec<u8>> {
    let raw = repeated_length_delimited_payloads(data, 1)?;
    if raw.len() != previous.entries.len() || previous.entries.len() != current.entries.len() {
        return Err(Error::InvalidFormat(
            "Numbers UUID row-set wire count is inconsistent".to_owned(),
        ));
    }
    let replacements = previous
        .entries
        .iter()
        .zip(&current.entries)
        .zip(raw)
        .map(|((previous, current), raw)| {
            let mut rewritten =
                patch_varint_field(raw, 1, true, Some(current.range_begin as i64 as u64))?;
            rewritten = patch_varint_field(
                &rewritten,
                2,
                previous.range_end.is_some(),
                current.range_end.map(|value| value as i64 as u64),
            )?;
            if tsce::index_set_archive::IndexSetEntry::decode(rewritten.as_slice())? != *current {
                return Err(Error::InvalidFormat(
                    "Numbers UUID row-set rewrite failed wire validation".to_owned(),
                ));
            }
            Ok(rewritten)
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, 1, &replacements)
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

pub(super) fn rewrite_shifted_range_tile_wire(
    original: &[u8],
    previous: &tsce::RangePrecedentsTileArchive,
    current: &tsce::RangePrecedentsTileArchive,
) -> Result<Vec<u8>> {
    let mut immutable = current.clone();
    immutable.from_to_range = previous.from_to_range.clone();
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "Numbers range dependency tile identity changed".to_owned(),
        ));
    }
    let raw = repeated_length_delimited_payloads(original, 2)?;
    if raw.len() != previous.from_to_range.len()
        || previous.from_to_range.len() != current.from_to_range.len()
    {
        return Err(Error::InvalidFormat(
            "Numbers range dependency tile wire count is inconsistent".to_owned(),
        ));
    }
    let replacements = previous
        .from_to_range
        .iter()
        .zip(&current.from_to_range)
        .zip(raw)
        .map(|((previous, current), raw)| {
            let mut immutable = *current;
            immutable.from_coord = previous.from_coord;
            immutable.refers_to_rect = previous.refers_to_rect;
            if immutable != *previous {
                return Err(Error::InvalidFormat(
                    "Numbers range dependency tile record changed outside coordinates".to_owned(),
                ));
            }
            let mut rewritten = transform_length_delimited_field(raw, 1, |coordinate| {
                rewrite_cell_coordinate(coordinate, &previous.from_coord, &current.from_coord)
            })?;
            rewritten = transform_length_delimited_field(&rewritten, 2, |rect| {
                rewrite_cell_rect(rect, &previous.refers_to_rect, &current.refers_to_rect)
            })?;
            if tsce::range_precedents_tile_archive::FromToRangeArchive::decode(
                rewritten.as_slice(),
            )? != *current
            {
                return Err(Error::InvalidFormat(
                    "Numbers range dependency tile record rewrite failed wire validation"
                        .to_owned(),
                ));
            }
            Ok(rewritten)
        })
        .collect::<Result<Vec<_>>>()?;
    let rewritten = rewrite_repeated_length_delimited_fields(original, 2, &replacements)?;
    if tsce::RangePrecedentsTileArchive::decode(rewritten.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers range dependency tile rewrite failed wire validation".to_owned(),
        ));
    }
    Ok(rewritten)
}

fn rewrite_cell_coordinate(
    data: &[u8],
    previous: &tsce::CellCoordinateArchive,
    current: &tsce::CellCoordinateArchive,
) -> Result<Vec<u8>> {
    if tsce::CellCoordinateArchive::decode(data)? != *previous {
        return Err(Error::InvalidFormat(
            "Numbers cell coordinate wire payload is inconsistent".to_owned(),
        ));
    }
    let mut immutable = *current;
    immutable.column = previous.column;
    immutable.row = previous.row;
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "Numbers cell coordinate changed outside explicit coordinates".to_owned(),
        ));
    }
    let mut rewritten = patch_varint_field(
        data,
        2,
        previous.column.is_some(),
        current.column.map(u64::from),
    )?;
    rewritten = patch_varint_field(
        &rewritten,
        3,
        previous.row.is_some(),
        current.row.map(u64::from),
    )?;
    if tsce::CellCoordinateArchive::decode(rewritten.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers cell coordinate rewrite failed wire validation".to_owned(),
        ));
    }
    Ok(rewritten)
}

fn rewrite_cell_rect(
    data: &[u8],
    previous: &tsce::CellRectArchive,
    current: &tsce::CellRectArchive,
) -> Result<Vec<u8>> {
    let mut immutable = *current;
    immutable.origin = previous.origin;
    immutable.size = previous.size;
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "Numbers range rectangle changed outside coordinates".to_owned(),
        ));
    }
    let mut rewritten = transform_length_delimited_field(data, 1, |origin| {
        rewrite_cell_coordinate(origin, &previous.origin, &current.origin)
    })?;
    rewritten = transform_length_delimited_field(&rewritten, 2, |size| {
        rewrite_column_row_size(size, &previous.size, &current.size)
    })?;
    if tsce::CellRectArchive::decode(rewritten.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers range rectangle rewrite failed wire validation".to_owned(),
        ));
    }
    Ok(rewritten)
}

fn rewrite_column_row_size(
    data: &[u8],
    previous: &tsce::ColumnRowSize,
    current: &tsce::ColumnRowSize,
) -> Result<Vec<u8>> {
    if tsce::ColumnRowSize::decode(data)? != *previous {
        return Err(Error::InvalidFormat(
            "Numbers range size wire payload is inconsistent".to_owned(),
        ));
    }
    let mut rewritten = patch_varint_field(
        data,
        1,
        previous.num_columns.is_some(),
        current.num_columns.map(u64::from),
    )?;
    rewritten = patch_varint_field(
        &rewritten,
        2,
        previous.num_rows.is_some(),
        current.num_rows.map(u64::from),
    )?;
    if tsce::ColumnRowSize::decode(rewritten.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers range size rewrite failed wire validation".to_owned(),
        ));
    }
    Ok(rewritten)
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
                        rewrite_shifted_edges(edges, previous, current)
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
) -> Result<Vec<u8>> {
    let mut immutable = current.clone();
    immutable.edge_without_owner_rows = previous.edge_without_owner_rows.clone();
    immutable.edge_without_owner_columns = previous.edge_without_owner_columns.clone();
    immutable.edge_with_owner_rows = previous.edge_with_owner_rows.clone();
    immutable.edge_with_owner_columns = previous.edge_with_owner_columns.clone();
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "Numbers formula edges changed outside row coordinates".to_owned(),
        ));
    }
    let mut data = original.to_vec();
    for (field, coordinates) in [
        (1, &current.edge_without_owner_rows),
        (2, &current.edge_without_owner_columns),
        (3, &current.edge_with_owner_rows),
        (4, &current.edge_with_owner_columns),
    ] {
        data = rewrite_repeated_varint_fields(
            &data,
            field,
            &coordinates
                .iter()
                .copied()
                .map(u64::from)
                .collect::<Vec<_>>(),
        )?;
    }
    if tsce::ExpandedEdgesArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "Numbers formula edge row shift failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}
