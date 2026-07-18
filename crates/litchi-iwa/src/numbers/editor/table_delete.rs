//! Transactional deletion of physical table rows and columns.

use super::*;

mod column;
mod dimension;
mod row;
mod uid;

use column::{delete_column_headers, delete_table_tile_column};
use dimension::{set_stroke_dimensions, set_table_dimensions};
use formula_dependency_shift::{DependencyAxis, delete_formula_dependencies};
use row::{delete_row_headers, delete_table_tile_row};
use table_headers::set_attached_table_header_settings;
use table_topology::{category_grouping_is_enabled, filter_has_row_state};
use uid::{delete_column_uid, delete_row_uid};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum TableAxis {
    Row,
    Column,
}

impl TableAxis {
    const fn noun(self) -> &'static str {
        match self {
            Self::Row => "row",
            Self::Column => "column",
        }
    }
}

impl NumbersEditor {
    /// Delete one physical table row and compact all following rows.
    ///
    /// Stored values, comments, formula records, headers, stable UIDs, and
    /// dimension sidecars are removed or shifted together. Formulas that still
    /// reference the deleted row cause the entire operation to fail unchanged.
    pub fn remove_table_row(&mut self, table_id: u64, row: usize) -> Result<()> {
        let mut staged = self.package.clone();
        let (new_rows, columns) = remove_attached_table_row(&mut staged, table_id, row)?;
        verify_numbers_dimensions(&staged, table_id, new_rows, columns)?;
        self.package = staged;
        Ok(())
    }

    /// Delete one physical table column and compact all following columns.
    ///
    /// Stored values, comments, formula records, headers, stable UIDs, and
    /// dimension sidecars are removed or shifted together. Formulas that still
    /// reference the deleted column cause the entire operation to fail unchanged.
    pub fn remove_table_column(&mut self, table_id: u64, column: usize) -> Result<()> {
        let mut staged = self.package.clone();
        let (rows, new_columns) = remove_attached_table_column(&mut staged, table_id, column)?;
        verify_numbers_dimensions(&staged, table_id, rows, new_columns)?;
        self.package = staged;
        Ok(())
    }
}

pub(super) fn remove_attached_table_row(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
) -> Result<(usize, usize)> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let old_rows = descriptor.model.number_of_rows as usize;
    if row >= old_rows {
        return Err(axis_index_error(TableAxis::Row, row, old_rows));
    }
    let new_rows = old_rows
        .checked_sub(1)
        .ok_or_else(|| Error::ParseError("iWork tables must retain at least one row".to_owned()))?;
    let (new_rows_u32, columns_u32) =
        validate_table_dimensions(new_rows, descriptor.model.number_of_columns as usize)?;
    let updated_header_settings = header_settings_after_row_deletion(&descriptor.model, row)?;
    let locations = object_locations(package)?;
    validate_deletion_features(
        package,
        &locations,
        &descriptor.model,
        TableAxis::Row,
        new_rows,
        updated_header_settings,
    )?;
    let cells = stored_cells_on_axis(package, &locations, &descriptor.model, TableAxis::Row, row)?;

    clear_stored_cells(package, table_id, &cells)?;
    delete_formula_dependencies(
        package,
        descriptor.table_info_id,
        DependencyAxis::Row,
        u32::try_from(row).map_err(|_| Error::ParseError("iWork row exceeds u32".to_owned()))?,
    )?;
    delete_table_tile_row(package, &locations, &descriptor.model, row)?;
    delete_row_headers(package, &locations, &descriptor.model, row)?;
    let uid = descriptor
        .model
        .base_column_row_uids
        .as_ref()
        .ok_or_else(|| {
            Error::ParseError(
                "Cannot safely delete an iWork row without a stable row UID map".to_owned(),
            )
        })?;
    delete_row_uid(package, &locations, uid.identifier, old_rows, row)?;
    if let Some(sidecar) = &descriptor.model.stroke_sidecar {
        set_stroke_dimensions(
            package,
            &locations,
            sidecar.identifier,
            new_rows_u32,
            columns_u32,
        )?;
    }
    set_table_dimensions(package, &locations, table_id, new_rows_u32, columns_u32)?;
    if let Some(settings) = updated_header_settings {
        set_attached_table_header_settings(package, table_id, settings)?;
    }
    verify_attached_dimensions(package, table_id, new_rows, columns_u32 as usize)?;
    Ok((new_rows, columns_u32 as usize))
}

pub(super) fn remove_attached_table_column(
    package: &mut IWorkPackage,
    table_id: u64,
    column: usize,
) -> Result<(usize, usize)> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let old_columns = descriptor.model.number_of_columns as usize;
    if column >= old_columns {
        return Err(axis_index_error(TableAxis::Column, column, old_columns));
    }
    let new_columns = old_columns.checked_sub(1).ok_or_else(|| {
        Error::ParseError("iWork tables must retain at least one column".to_owned())
    })?;
    let (rows_u32, new_columns_u32) =
        validate_table_dimensions(descriptor.model.number_of_rows as usize, new_columns)?;
    let updated_header_settings = header_settings_after_column_deletion(&descriptor.model, column)?;
    let locations = object_locations(package)?;
    validate_deletion_features(
        package,
        &locations,
        &descriptor.model,
        TableAxis::Column,
        new_columns,
        updated_header_settings,
    )?;
    let cells = stored_cells_on_axis(
        package,
        &locations,
        &descriptor.model,
        TableAxis::Column,
        column,
    )?;

    clear_stored_cells(package, table_id, &cells)?;
    delete_formula_dependencies(
        package,
        descriptor.table_info_id,
        DependencyAxis::Column,
        u32::try_from(column)
            .map_err(|_| Error::ParseError("iWork column exceeds u32".to_owned()))?,
    )?;
    delete_table_tile_column(package, &locations, &descriptor.model, column)?;
    delete_column_headers(
        package,
        &locations,
        descriptor.model.base_data_store.column_headers.identifier,
        column,
    )?;
    let uid = descriptor
        .model
        .base_column_row_uids
        .as_ref()
        .ok_or_else(|| {
            Error::ParseError(
                "Cannot safely delete an iWork column without a stable column UID map".to_owned(),
            )
        })?;
    delete_column_uid(package, &locations, uid.identifier, old_columns, column)?;
    if let Some(sidecar) = &descriptor.model.stroke_sidecar {
        set_stroke_dimensions(
            package,
            &locations,
            sidecar.identifier,
            rows_u32,
            new_columns_u32,
        )?;
    }
    set_table_dimensions(package, &locations, table_id, rows_u32, new_columns_u32)?;
    if let Some(settings) = updated_header_settings {
        set_attached_table_header_settings(package, table_id, settings)?;
    }
    verify_attached_dimensions(package, table_id, rows_u32 as usize, new_columns)?;
    Ok((rows_u32 as usize, new_columns))
}

fn validate_deletion_features(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    axis: TableAxis,
    new_length: usize,
    updated_header_settings: Option<NumbersTableHeaderSettings>,
) -> Result<()> {
    let stored_settings = NumbersTableHeaderSettings::from_model(model)?;
    let settings = updated_header_settings.unwrap_or(stored_settings);
    let fixed_regions_fit = match axis {
        TableAxis::Row => settings
            .header_row_count()
            .checked_add(settings.footer_row_count())
            .is_some_and(|fixed| fixed <= new_length),
        TableAxis::Column => settings.header_column_count() <= new_length,
    };
    if !fixed_regions_fit {
        return Err(Error::ParseError(format!(
            "Cannot delete an iWork {} without removing configured header or footer regions",
            axis.noun()
        )));
    }
    let hidden = match axis {
        TableAxis::Row => {
            model.number_of_hidden_rows.unwrap_or(0) != 0
                || model.number_of_user_hidden_rows.unwrap_or(0) != 0
        },
        TableAxis::Column => {
            model.number_of_hidden_columns.unwrap_or(0) != 0
                || model.number_of_user_hidden_columns.unwrap_or(0) != 0
        },
    };
    if hidden
        || model.number_of_filtered_rows.unwrap_or(0) != 0
        || model.pivot_owner.is_some()
        || model
            .sort_order
            .as_ref()
            .is_some_and(|sort| !sort.rules.is_empty())
        || model.merge_owner.as_ref().is_some_and(|owner| {
            owner
                .formula_store
                .as_ref()
                .is_some_and(|store| !store.formulas.is_empty())
        })
        || filter_has_row_state(package, locations, model.row_filter_set_pre_pivot.as_ref())?
        || category_grouping_is_enabled(package, locations, model.category_owner.as_ref())?
    {
        return Err(Error::ParseError(format!(
            "Cannot yet delete a {} from a sorted, filtered, hidden, merged, grouped, pivot, or spill iWork table",
            axis.noun()
        )));
    }
    if model.base_column_row_uids.is_none() {
        return Err(Error::ParseError(format!(
            "Cannot safely delete an iWork {} without a stable UID map",
            axis.noun()
        )));
    }
    Ok(())
}

fn header_settings_after_row_deletion(
    model: &TableModelArchive,
    row: usize,
) -> Result<Option<NumbersTableHeaderSettings>> {
    let mut settings = NumbersTableHeaderSettings::from_model(model)?;
    let rows = model.number_of_rows as usize;
    let header_rows = settings.header_row_count();
    let footer_rows = settings.footer_row_count();
    if header_rows
        .checked_add(footer_rows)
        .is_none_or(|fixed| fixed > rows)
    {
        return Err(Error::InvalidFormat(
            "iWork header and footer rows exceed the table row count".to_owned(),
        ));
    }
    if row < header_rows {
        settings.header_rows = decremented_header_count(header_rows)?;
        Ok(Some(settings))
    } else if footer_rows > 0 && row >= rows - footer_rows {
        settings.footer_rows = decremented_header_count(footer_rows)?;
        Ok(Some(settings))
    } else {
        Ok(None)
    }
}

fn header_settings_after_column_deletion(
    model: &TableModelArchive,
    column: usize,
) -> Result<Option<NumbersTableHeaderSettings>> {
    let mut settings = NumbersTableHeaderSettings::from_model(model)?;
    let header_columns = settings.header_column_count();
    if header_columns > model.number_of_columns as usize {
        return Err(Error::InvalidFormat(
            "iWork header columns exceed the table column count".to_owned(),
        ));
    }
    if column < header_columns {
        settings.header_columns = decremented_header_count(header_columns)?;
        Ok(Some(settings))
    } else {
        Ok(None)
    }
}

fn decremented_header_count(count: usize) -> Result<Option<NumbersTableHeaderCount>> {
    match count.checked_sub(1) {
        Some(0) => Ok(None),
        Some(count) => NumbersTableHeaderCount::new(count).map(Some),
        None => Err(Error::InvalidFormat(
            "iWork header/footer deletion underflow".to_owned(),
        )),
    }
}

fn stored_cells_on_axis(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    axis: TableAxis,
    target: usize,
) -> Result<Vec<(usize, usize)>> {
    let tile_size = model.base_data_store.tiles.tile_size.unwrap_or(256) as usize;
    if tile_size == 0 {
        return Err(Error::InvalidFormat(
            "Numbers table declares a zero tile size".to_owned(),
        ));
    }
    let rows = model.number_of_rows as usize;
    let columns = model.number_of_columns as usize;
    let mut cells = Vec::new();
    let mut seen = HashSet::new();
    for reference in &model.base_data_store.tiles.tiles {
        let tile_id = reference.tile.identifier;
        let archive_name = locations.get(&tile_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers tile object {tile_id} is missing"))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(tile_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers tile object {tile_id} is missing"))
        })?;
        let messages = object
            .messages
            .iter()
            .filter_map(|message| Tile::decode(message.data.as_slice()).ok())
            .collect::<Vec<_>>();
        let [tile] = messages.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "Numbers tile object {tile_id} must contain exactly one tile payload"
            )));
        };
        let base_row = reference.tileid as usize * tile_size;
        for row_info in &tile.row_infos {
            let row = base_row
                .checked_add(row_info.tile_row_index as usize)
                .ok_or_else(|| Error::ParseError("Numbers tile row overflow".to_owned()))?;
            if row >= rows || !seen.insert(row) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers table has invalid or duplicate stored row {row}"
                )));
            }
            if axis == TableAxis::Row && row != target {
                continue;
            }
            let stored = split_row(row_info)?;
            if stored.len() < columns || stored.iter().skip(columns).any(Option::is_some) {
                return Err(Error::InvalidFormat(format!(
                    "Numbers row {row} cell offsets do not match its table dimensions"
                )));
            }
            match axis {
                TableAxis::Row => {
                    cells.extend(
                        stored
                            .iter()
                            .take(columns)
                            .enumerate()
                            .filter_map(|(column, cell)| cell.as_ref().map(|_| (row, column))),
                    );
                },
                TableAxis::Column if stored[target].is_some() => cells.push((row, target)),
                TableAxis::Column => {},
            }
        }
    }
    cells.sort_unstable();
    Ok(cells)
}

fn clear_stored_cells(
    package: &mut IWorkPackage,
    table_id: u64,
    cells: &[(usize, usize)],
) -> Result<()> {
    for &(row, column) in cells {
        clear_cell_comment_in_package(package, table_id, row, column)?;
        set_cell_in_package(package, table_id, row, column, CellValue::Empty)?;
    }
    Ok(())
}

fn verify_attached_dimensions(
    package: &IWorkPackage,
    table_id: u64,
    rows: usize,
    columns: usize,
) -> Result<()> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    if (
        descriptor.model.number_of_rows as usize,
        descriptor.model.number_of_columns as usize,
    ) != (rows, columns)
    {
        return Err(Error::InvalidFormat(
            "iWork table axis deletion failed dimension validation".to_owned(),
        ));
    }
    Ok(())
}

fn verify_numbers_dimensions(
    package: &IWorkPackage,
    table_id: u64,
    rows: usize,
    columns: usize,
) -> Result<()> {
    let verified = NumbersEditor::from_bytes(&package.to_bytes()?)?;
    let table = verified
        .tables()?
        .into_iter()
        .find(|table| table.object_id == table_id)
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers table disappeared after axis deletion".to_owned())
        })?;
    if (table.rows, table.columns) != (rows, columns) {
        return Err(Error::InvalidFormat(
            "Numbers table axis deletion failed dimension validation".to_owned(),
        ));
    }
    Ok(())
}

fn axis_index_error(axis: TableAxis, index: usize, length: usize) -> Error {
    Error::ParseError(format!(
        "Cannot delete iWork {} {index} from a table with {length} {}s",
        axis.noun(),
        axis.noun()
    ))
}
