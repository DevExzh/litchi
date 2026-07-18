//! Preflight validation for formula AST coordinates affected by axis edits.

use super::*;

pub(super) fn validate_formula_ast_stability(
    package: &IWorkPackage,
    table_info_id: u64,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
) -> Result<()> {
    const COMPONENT: &str = "Index/CalculationEngine.iwa";

    let descriptor = attached_table_descriptors(package)?
        .into_iter()
        .find(|table| table.table_info_id == table_info_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table info {table_info_id} has no attached table model"
            ))
        })?;
    let archive = package.archive(COMPONENT)?;
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
        return Ok(());
    }

    let locations = object_locations(package)?;
    let formulas = resolve_table_data_list(
        package,
        &locations,
        descriptor.model.base_data_store.formula_table.identifier,
        tst::table_data_list::ListType::Formula,
    )?;
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
        let formula = formulas
            .entries
            .iter()
            .find(|entry| entry.entry.key == identifier)
            .and_then(|entry| entry.entry.formula.as_ref())
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork formula table has no formula entry {identifier}"
                ))
            })?;
        validate_formula_nodes_stable(formula, row, column, axis, position, mutation)?;
    }
    Ok(())
}

fn validate_formula_nodes_stable(
    formula: &tsce::FormulaArchive,
    host_row: u32,
    host_column: u32,
    axis: DependencyAxis,
    position: u32,
    mutation: DependencyMutation,
) -> Result<()> {
    let host = match axis {
        DependencyAxis::Row => host_row,
        DependencyAxis::Column => host_column,
    };
    let shifted_host = mutation.coordinate(host, position, "formula host coordinate")?;
    for node in &formula.ast_node_array.ast_node {
        if node.ast_colon_tract.is_some()
            || (node.ast_category_ref.is_some() && shifted_host != host)
        {
            return Err(formula_ast_rewrite_error(axis, mutation));
        }
        let coordinate = match axis {
            DependencyAxis::Row => node
                .ast_row
                .as_ref()
                .map(|coordinate| (coordinate.row, coordinate.absolute.unwrap_or(false))),
            DependencyAxis::Column => node
                .ast_column
                .as_ref()
                .map(|coordinate| (coordinate.column, coordinate.absolute.unwrap_or(false))),
        };
        let Some((encoded, absolute)) = coordinate else {
            continue;
        };
        if node.ast_cross_table_reference_extra_info.is_some() {
            if !absolute && shifted_host != host {
                return Err(formula_ast_rewrite_error(axis, mutation));
            }
            continue;
        }
        let target = if absolute {
            i64::from(encoded)
        } else {
            i64::from(host) + i64::from(encoded)
        };
        let target = u32::try_from(target).map_err(|_| {
            Error::InvalidFormat("iWork formula AST coordinate is outside u32".to_owned())
        })?;
        let shifted_target = mutation.coordinate(target, position, "formula reference")?;
        let shifted_encoded = if absolute {
            i64::from(shifted_target)
        } else {
            i64::from(shifted_target) - i64::from(shifted_host)
        };
        if shifted_encoded != i64::from(encoded) {
            return Err(formula_ast_rewrite_error(axis, mutation));
        }
    }
    Ok(())
}

fn formula_ast_rewrite_error(axis: DependencyAxis, mutation: DependencyMutation) -> Error {
    Error::ParseError(format!(
        "Cannot safely {} an iWork {} because a surviving formula would require an AST coordinate rewrite",
        mutation.verb(),
        axis.noun()
    ))
}
