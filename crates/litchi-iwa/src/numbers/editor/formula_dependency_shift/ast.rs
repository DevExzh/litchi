//! Lossless formula AST coordinate rewrites for table-axis edits.

use super::*;

#[derive(Debug)]
struct FormulaRewrite {
    row: u32,
    column: u32,
    previous: tsce::FormulaArchive,
    current: tsce::FormulaArchive,
}

pub(super) fn rewrite_formula_asts(
    package: &mut IWorkPackage,
    component: &str,
    table_info_id: u64,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    expand_footer_ranges: bool,
) -> Result<FormulaDependencyAdjustments> {
    let descriptor = attached_table_descriptors(package)?
        .into_iter()
        .find(|table| table.table_info_id == table_info_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table info {table_info_id} has no attached table model"
            ))
        })?;
    let archive = package.archive(component)?;
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
                "iWork table info {table_info_id} has no formula dependency owner"
            ))
        })?;
    let mut formula_cells = owner
        .cell_dependencies
        .as_ref()
        .into_iter()
        .flat_map(|dependencies| &dependencies.cell_record)
        .map(|record| (record.row, record.column))
        .collect::<HashSet<_>>();
    for reference in owner
        .tiled_cell_dependencies
        .as_ref()
        .into_iter()
        .flat_map(|dependencies| &dependencies.cell_record_tiles)
    {
        let object = archive.object(reference.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork formula dependency tile {} is missing",
                reference.identifier
            ))
        })?;
        let tile = object
            .messages
            .iter()
            .find(|message| message.type_ == 4009)
            .map(|message| tsce::CellRecordTileArchive::decode(message.data.as_slice()))
            .transpose()?
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork formula dependency tile {} has no payload",
                    reference.identifier
                ))
            })?;
        formula_cells.extend(
            tile.cell_records
                .into_iter()
                .map(|record| (record.row, record.column)),
        );
    }
    if formula_cells.is_empty() {
        return Ok(FormulaDependencyAdjustments::default());
    }

    let locations = object_locations(package)?;
    let formula_table_id = descriptor.model.base_data_store.formula_table.identifier;
    let formulas = resolve_table_data_list(
        package,
        &locations,
        formula_table_id,
        tst::table_data_list::ListType::Formula,
    )?;
    let mut formula_cells = formula_cells.into_iter().collect::<Vec<_>>();
    formula_cells.sort_unstable();
    let mut rewrites = BTreeMap::<u32, Vec<FormulaRewrite>>::new();
    let mut dependency_adjustments = FormulaDependencyAdjustments::default();
    for (row, column) in formula_cells {
        let row_index = usize::try_from(row)
            .map_err(|_| Error::ParseError("iWork formula row exceeds usize".to_owned()))?;
        let column_index = usize::try_from(column)
            .map_err(|_| Error::ParseError("iWork formula column exceeds usize".to_owned()))?;
        let location =
            locate_attached_cell(package, descriptor.object_id, row_index, column_index)?;
        let stored = read_tile_cell(
            package,
            &location.tile_archive,
            location.tile_id,
            location.tile_row,
            column_index,
        )?
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork formula cell ({row}, {column}) is missing"))
        })?;
        let identifier = match BncCell::parse(&stored)?.stored_value() {
            StoredValue::Formula(identifier) => identifier,
            _ => {
                return Err(Error::InvalidFormat(format!(
                    "iWork dependency cell ({row}, {column}) does not contain a formula"
                )));
            },
        };
        let previous = formulas
            .entries
            .iter()
            .find(|entry| entry.entry.key == identifier)
            .and_then(|entry| entry.entry.formula.as_ref())
            .cloned()
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork formula table has no formula entry {identifier}"
                ))
            })?;
        let footer_rows = descriptor.model.number_of_footer_rows.unwrap_or(0);
        let footer_boundary = descriptor.model.number_of_rows.saturating_sub(footer_rows);
        let footer = (footer_rows > 0 && row >= footer_boundary).then_some(FooterFormulaContext {
            boundary: footer_boundary,
            expand_on_insert: expand_footer_ranges,
        });
        let mut current = previous.clone();
        let mut local_adjustments = LocalPrecedentAdjustments::default();
        rewrite_formula_nodes(
            &mut current.ast_node_array,
            row,
            column,
            axis,
            position,
            mutation,
            footer,
            &mut local_adjustments,
        )?;
        local_adjustments.normalize();
        if !local_adjustments.is_empty() {
            dependency_adjustments
                .local_precedents
                .insert((row, column), local_adjustments);
        }
        rewrites
            .entry(identifier)
            .or_default()
            .push(FormulaRewrite {
                row,
                column,
                previous,
                current,
            });
    }

    for (identifier, jobs) in rewrites {
        let located = formulas
            .entries
            .iter()
            .find(|entry| entry.entry.key == identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork formula table has no formula entry {identifier}"
                ))
            })?;
        let changed = jobs
            .iter()
            .filter(|job| job.previous != job.current)
            .collect::<Vec<_>>();
        if changed.is_empty() {
            continue;
        }
        let all_references_are_rewritten = usize::try_from(located.entry.refcount)
            .is_ok_and(|refcount| refcount == jobs.len())
            && changed.len() == jobs.len();
        let one_shared_result = changed.iter().all(|job| job.current == changed[0].current);
        let existing_equivalent = formulas.entries.iter().any(|entry| {
            entry.entry.key != identifier
                && entry.entry.formula.as_ref() == Some(&changed[0].current)
        });
        if !existing_equivalent
            && (located.entry.refcount == 1 || (all_references_are_rewritten && one_shared_result))
        {
            let job = changed[0];
            rewrite_formula_table_entry(package, &formulas, located, &job.current, |wire| {
                rewrite_formula_archive_wire(wire, &job.previous, &job.current)
            })?;
            continue;
        }
        for job in changed {
            let replacement =
                insert_formula_table(package, &locations, formula_table_id, job.current.clone())?;
            set_encoded_cell_value(
                package,
                descriptor.object_id,
                usize::try_from(job.row)
                    .map_err(|_| Error::ParseError("iWork formula row exceeds usize".to_owned()))?,
                usize::try_from(job.column).map_err(|_| {
                    Error::ParseError("iWork formula column exceeds usize".to_owned())
                })?,
                EncodedValue::Formula(replacement),
            )?;
            decrement_formula_table(package, &locations, formula_table_id, identifier)?;
        }
    }
    Ok(dependency_adjustments)
}

#[derive(Clone, Copy, Debug)]
struct FooterFormulaContext {
    boundary: u32,
    expand_on_insert: bool,
}

fn footer_range_overrides(
    array: &tsce::AstNodeArrayArchive,
    host_row: u32,
    host_column: u32,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    footer: Option<FooterFormulaContext>,
    dependency_adjustments: &mut LocalPrecedentAdjustments,
) -> Result<HashMap<usize, u32>> {
    use tsce::ast_node_array_archive::AstNodeType;

    let Some(footer) = footer.filter(|_| axis == DependencyAxis::Row) else {
        return Ok(HashMap::new());
    };
    let expands_footer = mutation == DependencyMutation::Insert
        && footer.expand_on_insert
        && position == footer.boundary;
    let contracts_footer =
        mutation == DependencyMutation::Delete && position.checked_add(1) == Some(footer.boundary);
    if !expands_footer && !contracts_footer {
        return Ok(HashMap::new());
    }

    let mut overrides = HashMap::new();
    for index in 0..array.ast_node.len().saturating_sub(2) {
        let start_node = &array.ast_node[index];
        let end_node = &array.ast_node[index + 1];
        let colon = &array.ast_node[index + 2];
        if colon.ast_node_type != AstNodeType::ColonNode as i32 {
            continue;
        }
        let Some(start) = local_cell_coordinate(start_node, host_row, host_column)? else {
            continue;
        };
        let Some(end) = local_cell_coordinate(end_node, host_row, host_column)? else {
            continue;
        };
        let (endpoint_index, replacement_row) =
            if expands_footer && start.0.max(end.0).checked_add(1) == Some(position) {
                (if end.0 >= start.0 { index + 1 } else { index }, position)
            } else if contracts_footer
                && start.0.max(end.0) == position
                && start.0.min(end.0) < position
            {
                (
                    if end.0 >= start.0 { index + 1 } else { index },
                    position.checked_sub(1).ok_or_else(|| {
                        Error::ParseError("iWork footer range contraction underflow".to_owned())
                    })?,
                )
            } else {
                continue;
            };
        if overrides.insert(endpoint_index, replacement_row).is_some() {
            return Err(Error::InvalidFormat(
                "iWork formula range endpoint participates in multiple footer ranges".to_owned(),
            ));
        }
        let coordinates =
            (start.1.min(end.1)..=start.1.max(end.1)).map(|column| (position, column));
        match mutation {
            DependencyMutation::Insert => dependency_adjustments.insert.extend(coordinates),
            DependencyMutation::Delete => dependency_adjustments.remove.extend(coordinates),
        }
    }
    Ok(overrides)
}

fn local_cell_coordinate(
    node: &tsce::ast_node_array_archive::AstNodeArchive,
    host_row: u32,
    host_column: u32,
) -> Result<Option<(u32, u32)>> {
    if node.ast_cross_table_reference_extra_info.is_some() {
        return Ok(None);
    }
    let (Some(row), Some(column)) = (node.ast_row.as_ref(), node.ast_column.as_ref()) else {
        return Ok(None);
    };
    let resolve = |encoded: i32, absolute: bool, host: u32| {
        let value = if absolute {
            i64::from(encoded)
        } else {
            i64::from(host) + i64::from(encoded)
        };
        u32::try_from(value).map_err(|_| {
            Error::InvalidFormat("iWork formula AST coordinate is outside u32".to_owned())
        })
    };
    Ok(Some((
        resolve(row.row, row.absolute.unwrap_or(false), host_row)?,
        resolve(column.column, column.absolute.unwrap_or(false), host_column)?,
    )))
}

fn rewrite_formula_nodes(
    array: &mut tsce::AstNodeArrayArchive,
    host_row: u32,
    host_column: u32,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    footer: Option<FooterFormulaContext>,
    dependency_adjustments: &mut LocalPrecedentAdjustments,
) -> Result<()> {
    let host = match axis {
        DependencyAxis::Row => host_row,
        DependencyAxis::Column => host_column,
    };
    let shifted_host = mutation.coordinate(host, position, "formula host coordinate")?;
    let footer_overrides = footer_range_overrides(
        array,
        host_row,
        host_column,
        axis,
        position,
        mutation,
        footer,
        dependency_adjustments,
    )?;
    for (index, node) in array.ast_node.iter_mut().enumerate() {
        if node.ast_category_ref.is_some() && shifted_host != host {
            return Err(formula_ast_rewrite_error(axis, mutation));
        }
        rewrite_colon_tract(
            node,
            host_row,
            host_column,
            axis,
            position,
            mutation,
            footer,
            dependency_adjustments,
        )?;
        if let Some(nested) = &mut node.ast_thunk_node_array {
            rewrite_formula_nodes(
                nested,
                host_row,
                host_column,
                axis,
                position,
                mutation,
                footer,
                dependency_adjustments,
            )?;
        }
        let (encoded, absolute) = match axis {
            DependencyAxis::Row => node
                .ast_row
                .as_ref()
                .map(|coordinate| (coordinate.row, coordinate.absolute.unwrap_or(false))),
            DependencyAxis::Column => node
                .ast_column
                .as_ref()
                .map(|coordinate| (coordinate.column, coordinate.absolute.unwrap_or(false))),
        }
        .unwrap_or((0, false));
        let has_coordinate = match axis {
            DependencyAxis::Row => node.ast_row.is_some(),
            DependencyAxis::Column => node.ast_column.is_some(),
        };
        if !has_coordinate {
            continue;
        }
        let shifted_encoded = if node.ast_cross_table_reference_extra_info.is_some() {
            if absolute {
                i64::from(encoded)
            } else {
                i64::from(host)
                    .checked_add(i64::from(encoded))
                    .and_then(|target| target.checked_sub(i64::from(shifted_host)))
                    .ok_or_else(|| {
                        Error::ParseError("iWork formula coordinate overflow".to_owned())
                    })?
            }
        } else {
            let target = if absolute {
                i64::from(encoded)
            } else {
                i64::from(host) + i64::from(encoded)
            };
            let target = u32::try_from(target).map_err(|_| {
                Error::InvalidFormat("iWork formula AST coordinate is outside u32".to_owned())
            })?;
            let shifted_target = footer_overrides
                .get(&index)
                .copied()
                .map(Ok)
                .unwrap_or_else(|| mutation.coordinate(target, position, "formula reference"))?;
            if absolute {
                i64::from(shifted_target)
            } else {
                i64::from(shifted_target) - i64::from(shifted_host)
            }
        };
        let shifted_encoded = i32::try_from(shifted_encoded).map_err(|_| {
            Error::ParseError("iWork formula AST coordinate exceeds i32".to_owned())
        })?;
        match axis {
            DependencyAxis::Row => {
                let coordinate = node.ast_row.as_mut().ok_or_else(|| {
                    Error::InvalidFormat(
                        "iWork formula row coordinate disappeared during mutation".to_owned(),
                    )
                })?;
                coordinate.row = shifted_encoded;
            },
            DependencyAxis::Column => {
                let coordinate = node.ast_column.as_mut().ok_or_else(|| {
                    Error::InvalidFormat(
                        "iWork formula column coordinate disappeared during mutation".to_owned(),
                    )
                })?;
                coordinate.column = shifted_encoded;
            },
        }
    }
    Ok(())
}

fn rewrite_colon_tract(
    node: &mut tsce::ast_node_array_archive::AstNodeArchive,
    host_row: u32,
    host_column: u32,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    footer: Option<FooterFormulaContext>,
    dependency_adjustments: &mut LocalPrecedentAdjustments,
) -> Result<()> {
    let cross_table = node.ast_cross_table_reference_extra_info.is_some();
    let Some(tract) = &mut node.ast_colon_tract else {
        return Ok(());
    };
    let column_ranges = tract_axis_intervals(tract, DependencyAxis::Column, host_column)?;
    let (relative, absolute, host) = match axis {
        DependencyAxis::Column => (
            &mut tract.relative_column,
            &mut tract.absolute_column,
            host_column,
        ),
        DependencyAxis::Row => (&mut tract.relative_row, &mut tract.absolute_row, host_row),
    };
    let shifted_host = mutation.coordinate(host, position, "formula host coordinate")?;
    for range in relative {
        let begin = resolve_relative_tract_coordinate(host, range.range_begin)?;
        let end = range
            .range_end
            .map(|end| resolve_relative_tract_coordinate(host, end))
            .transpose()?
            .unwrap_or(begin);
        let (shifted_begin, shifted_end) = mutate_tract_interval(
            begin,
            end,
            cross_table,
            axis,
            position,
            mutation,
            footer,
            &column_ranges,
            dependency_adjustments,
        )?;
        range.range_begin = relative_tract_coordinate(shifted_begin, shifted_host)?;
        range.range_end = (shifted_begin != shifted_end)
            .then(|| relative_tract_coordinate(shifted_end, shifted_host))
            .transpose()?;
    }
    for range in absolute {
        let begin = range.range_begin;
        let end = range.range_end.unwrap_or(begin);
        let (shifted_begin, shifted_end) = mutate_tract_interval(
            begin,
            end,
            cross_table,
            axis,
            position,
            mutation,
            footer,
            &column_ranges,
            dependency_adjustments,
        )?;
        range.range_begin = shifted_begin;
        range.range_end = (shifted_begin != shifted_end).then_some(shifted_end);
    }
    Ok(())
}

fn tract_axis_intervals(
    tract: &tsce::ast_node_array_archive::AstColonTractArchive,
    axis: DependencyAxis,
    host: u32,
) -> Result<Vec<(u32, u32)>> {
    let (relative, absolute) = match axis {
        DependencyAxis::Column => (&tract.relative_column, &tract.absolute_column),
        DependencyAxis::Row => (&tract.relative_row, &tract.absolute_row),
    };
    let mut intervals = Vec::with_capacity(relative.len() + absolute.len());
    for range in relative {
        let begin = resolve_relative_tract_coordinate(host, range.range_begin)?;
        let end = range
            .range_end
            .map(|end| resolve_relative_tract_coordinate(host, end))
            .transpose()?
            .unwrap_or(begin);
        intervals.push((begin.min(end), begin.max(end)));
    }
    intervals.extend(absolute.iter().map(|range| {
        let end = range.range_end.unwrap_or(range.range_begin);
        (range.range_begin.min(end), range.range_begin.max(end))
    }));
    Ok(intervals)
}

#[allow(clippy::too_many_arguments)]
fn mutate_tract_interval(
    begin: u32,
    end: u32,
    cross_table: bool,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
    footer: Option<FooterFormulaContext>,
    column_ranges: &[(u32, u32)],
    dependency_adjustments: &mut LocalPrecedentAdjustments,
) -> Result<(u32, u32)> {
    if cross_table {
        return Ok((begin, end));
    }
    let low = begin.min(end);
    let high = begin.max(end);
    let footer_change = axis == DependencyAxis::Row
        && footer.is_some_and(|footer| match mutation {
            DependencyMutation::Insert => {
                footer.expand_on_insert
                    && position == footer.boundary
                    && high.checked_add(1) == Some(position)
            },
            DependencyMutation::Delete => {
                position.checked_add(1) == Some(footer.boundary)
                    && high == position
                    && low < position
            },
        });
    let shifted_high = if footer_change {
        match mutation {
            DependencyMutation::Insert => position,
            DependencyMutation::Delete => position.checked_sub(1).ok_or_else(|| {
                Error::ParseError("iWork footer tract contraction underflow".to_owned())
            })?,
        }
    } else {
        mutation.coordinate(high, position, "formula tract endpoint")?
    };
    let shifted_low = mutation.coordinate(low, position, "formula tract endpoint")?;
    if footer_change {
        if column_ranges.is_empty() {
            return Err(Error::ParseError(
                "Cannot expand an iWork footer tract without explicit columns".to_owned(),
            ));
        }
        for &(start, end) in column_ranges {
            let coordinates = (start..=end).map(|column| (position, column));
            match mutation {
                DependencyMutation::Insert => dependency_adjustments.insert.extend(coordinates),
                DependencyMutation::Delete => dependency_adjustments.remove.extend(coordinates),
            }
        }
    }
    Ok(if begin <= end {
        (shifted_low, shifted_high)
    } else {
        (shifted_high, shifted_low)
    })
}

fn resolve_relative_tract_coordinate(host: u32, offset: i32) -> Result<u32> {
    let coordinate = i64::from(host) + i64::from(offset);
    u32::try_from(coordinate)
        .map_err(|_| Error::InvalidFormat("iWork formula tract is outside u32".to_owned()))
}

fn relative_tract_coordinate(coordinate: u32, host: u32) -> Result<i32> {
    let offset = i64::from(coordinate) - i64::from(host);
    i32::try_from(offset)
        .map_err(|_| Error::ParseError("iWork formula tract exceeds i32".to_owned()))
}

pub(super) fn rewrite_formula_archive_wire(
    data: &[u8],
    previous: &tsce::FormulaArchive,
    current: &tsce::FormulaArchive,
) -> Result<Vec<u8>> {
    let data = transform_length_delimited_field(data, 1, |array| {
        rewrite_ast_array_wire(array, &previous.ast_node_array, &current.ast_node_array)
    })?;
    if tsce::FormulaArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "iWork formula AST wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn rewrite_ast_array_wire(
    data: &[u8],
    previous: &tsce::AstNodeArrayArchive,
    current: &tsce::AstNodeArrayArchive,
) -> Result<Vec<u8>> {
    let raw_nodes = repeated_length_delimited_payloads(data, 1)?;
    if raw_nodes.len() != previous.ast_node.len()
        || previous.ast_node.len() != current.ast_node.len()
    {
        return Err(Error::InvalidFormat(
            "iWork formula AST node count changed during coordinate mutation".to_owned(),
        ));
    }
    let replacements = raw_nodes
        .into_iter()
        .zip(&previous.ast_node)
        .zip(&current.ast_node)
        .map(|((raw, previous), current)| rewrite_ast_node_wire(raw, previous, current))
        .collect::<Result<Vec<_>>>()?;
    let data = rewrite_repeated_length_delimited_fields(data, 1, &replacements)?;
    if tsce::AstNodeArrayArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "iWork formula AST array wire mutation failed validation".to_owned(),
        ));
    }
    Ok(data)
}

fn rewrite_ast_node_wire(
    data: &[u8],
    previous: &tsce::ast_node_array_archive::AstNodeArchive,
    current: &tsce::ast_node_array_archive::AstNodeArchive,
) -> Result<Vec<u8>> {
    let mut data = data.to_vec();
    for (field, previous, current) in [
        (
            26,
            previous.ast_column.map(|coordinate| coordinate.column),
            current.ast_column.map(|coordinate| coordinate.column),
        ),
        (
            27,
            previous.ast_row.map(|coordinate| coordinate.row),
            current.ast_row.map(|coordinate| coordinate.row),
        ),
    ] {
        if previous != current {
            let value = current.ok_or_else(|| {
                Error::InvalidFormat(
                    "iWork formula coordinate disappeared during mutation".to_owned(),
                )
            })?;
            data = patch_nested_varint_field(&data, &[field, 1], true, Some(zigzag_i32(value)))?;
        }
    }
    if previous.ast_thunk_node_array != current.ast_thunk_node_array {
        let previous_array = previous.ast_thunk_node_array.as_ref().ok_or_else(|| {
            Error::InvalidFormat("iWork formula thunk appeared during mutation".to_owned())
        })?;
        let current_array = current.ast_thunk_node_array.as_ref().ok_or_else(|| {
            Error::InvalidFormat("iWork formula thunk disappeared during mutation".to_owned())
        })?;
        data = transform_length_delimited_field(&data, 14, |array| {
            rewrite_ast_array_wire(array, previous_array, current_array)
        })?;
    }
    if previous.ast_colon_tract != current.ast_colon_tract {
        let previous_tract = previous.ast_colon_tract.as_ref().ok_or_else(|| {
            Error::InvalidFormat("iWork formula colon tract appeared during mutation".to_owned())
        })?;
        let current_tract = current.ast_colon_tract.as_ref().ok_or_else(|| {
            Error::InvalidFormat("iWork formula colon tract disappeared during mutation".to_owned())
        })?;
        data = transform_length_delimited_field(&data, 40, |tract| {
            rewrite_colon_tract_wire(tract, previous_tract, current_tract)
        })?;
    }
    if tsce::ast_node_array_archive::AstNodeArchive::decode(data.as_slice())? != *current {
        return Err(Error::InvalidFormat(
            "iWork formula AST node changed outside supported coordinates".to_owned(),
        ));
    }
    Ok(data)
}

fn rewrite_colon_tract_wire(
    data: &[u8],
    previous: &tsce::ast_node_array_archive::AstColonTractArchive,
    current: &tsce::ast_node_array_archive::AstColonTractArchive,
) -> Result<Vec<u8>> {
    let mut immutable = current.clone();
    immutable.relative_column = previous.relative_column.clone();
    immutable.relative_row = previous.relative_row.clone();
    immutable.absolute_column = previous.absolute_column.clone();
    immutable.absolute_row = previous.absolute_row.clone();
    if immutable != *previous {
        return Err(Error::InvalidFormat(
            "iWork formula colon tract changed outside coordinate ranges".to_owned(),
        ));
    }
    let mut rewritten = data.to_vec();
    for (field, previous, current) in [
        (1, &previous.relative_column, &current.relative_column),
        (2, &previous.relative_row, &current.relative_row),
    ] {
        rewritten = rewrite_relative_tract_ranges(&rewritten, field, previous, current)?;
    }
    for (field, previous, current) in [
        (3, &previous.absolute_column, &current.absolute_column),
        (4, &previous.absolute_row, &current.absolute_row),
    ] {
        rewritten = rewrite_absolute_tract_ranges(&rewritten, field, previous, current)?;
    }
    if tsce::ast_node_array_archive::AstColonTractArchive::decode(rewritten.as_slice())? != *current
    {
        return Err(Error::InvalidFormat(
            "iWork formula colon tract rewrite failed wire validation".to_owned(),
        ));
    }
    Ok(rewritten)
}

fn rewrite_relative_tract_ranges(
    data: &[u8],
    field_number: u32,
    previous: &[tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractRelativeRangeArchive],
    current: &[tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractRelativeRangeArchive],
) -> Result<Vec<u8>> {
    let raw = repeated_length_delimited_payloads(data, field_number)?;
    if raw.len() != previous.len() || previous.len() != current.len() {
        return Err(Error::InvalidFormat(
            "iWork relative tract range count changed".to_owned(),
        ));
    }
    let replacements = previous
        .iter()
        .zip(current)
        .zip(raw)
        .map(|((previous, current), raw)| {
            let mut rewritten = patch_varint_field(
                raw,
                1,
                true,
                Some(current.range_begin as i64 as u64),
            )?;
            rewritten = patch_varint_field(
                &rewritten,
                2,
                previous.range_end.is_some(),
                current.range_end.map(|value| value as i64 as u64),
            )?;
            if tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractRelativeRangeArchive::decode(
                rewritten.as_slice(),
            )? != *current
            {
                return Err(Error::InvalidFormat(
                    "iWork relative tract range rewrite failed wire validation".to_owned(),
                ));
            }
            Ok(rewritten)
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, field_number, &replacements)
}

fn rewrite_absolute_tract_ranges(
    data: &[u8],
    field_number: u32,
    previous: &[tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractAbsoluteRangeArchive],
    current: &[tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractAbsoluteRangeArchive],
) -> Result<Vec<u8>> {
    let raw = repeated_length_delimited_payloads(data, field_number)?;
    if raw.len() != previous.len() || previous.len() != current.len() {
        return Err(Error::InvalidFormat(
            "iWork absolute tract range count changed".to_owned(),
        ));
    }
    let replacements = previous
        .iter()
        .zip(current)
        .zip(raw)
        .map(|((previous, current), raw)| {
            let mut rewritten = patch_varint_field(
                raw,
                1,
                true,
                Some(u64::from(current.range_begin)),
            )?;
            rewritten = patch_varint_field(
                &rewritten,
                2,
                previous.range_end.is_some(),
                current.range_end.map(u64::from),
            )?;
            if tsce::ast_node_array_archive::ast_colon_tract_archive::AstColonTractAbsoluteRangeArchive::decode(
                rewritten.as_slice(),
            )? != *current
            {
                return Err(Error::InvalidFormat(
                    "iWork absolute tract range rewrite failed wire validation".to_owned(),
                ));
            }
            Ok(rewritten)
        })
        .collect::<Result<Vec<_>>>()?;
    rewrite_repeated_length_delimited_fields(data, field_number, &replacements)
}

const fn zigzag_i32(value: i32) -> u64 {
    ((value << 1) ^ (value >> 31)) as u32 as u64
}

fn formula_ast_rewrite_error(axis: DependencyAxis, mutation: DependencyMutation) -> Error {
    Error::ParseError(format!(
        "Cannot safely {} an iWork {} because a surviving formula uses an unsupported tract or category reference",
        mutation.verb(),
        axis.noun()
    ))
}
