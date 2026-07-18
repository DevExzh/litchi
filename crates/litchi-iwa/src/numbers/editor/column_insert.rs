//! Transactional insertion of physical table columns.

use super::*;

mod storage;

use formula_dependency_shift::{DependencyAxis, shift_formula_dependencies};
use storage::{
    insert_column_uid, insert_stroke_column, set_table_column_count, shift_column_headers,
    shift_table_tile_columns,
};
use table_topology::{category_grouping_is_enabled, filter_has_row_state};

impl NumbersEditor {
    /// Insert one blank physical column before `column` in a table.
    ///
    /// `column == table.columns` appends a column. Stored cells, column
    /// metadata, stable column UIDs, and ordinary formula dependency
    /// coordinates are shifted in lockstep. Unsupported topology is rejected
    /// transactionally without changing the package.
    pub fn insert_table_column(&mut self, table_id: u64, column: usize) -> Result<()> {
        let mut staged = self.package.clone();
        let new_columns = insert_attached_table_column(&mut staged, table_id, column)?;
        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        let table = verified
            .tables()?
            .into_iter()
            .find(|table| table.object_id == table_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers table disappeared after column insertion".to_owned())
            })?;
        if table.columns != new_columns {
            return Err(Error::InvalidFormat(
                "Numbers inserted column failed dimension validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }
}

pub(super) fn insert_attached_table_column(
    package: &mut IWorkPackage,
    table_id: u64,
    column: usize,
) -> Result<usize> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let old_columns = descriptor.model.number_of_columns as usize;
    if column > old_columns {
        return Err(Error::ParseError(format!(
            "Cannot insert iWork column {column} into a table with {old_columns} columns"
        )));
    }
    let new_columns = old_columns
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork column count overflow".to_owned()))?;
    let (_, new_columns_u32) =
        validate_table_dimensions(descriptor.model.number_of_rows as usize, new_columns)?;
    let locations = object_locations(package)?;
    validate_column_insertion_features(package, &locations, &descriptor.model, column)?;

    shift_formula_dependencies(
        package,
        descriptor.table_info_id,
        DependencyAxis::Column,
        u32::try_from(column)
            .map_err(|_| Error::ParseError("iWork column exceeds u32".to_owned()))?,
    )?;
    shift_table_tile_columns(package, &locations, &descriptor.model, column)?;
    shift_column_headers(
        package,
        &locations,
        descriptor.model.base_data_store.column_headers.identifier,
        column,
    )?;
    if let Some(reference) = &descriptor.model.base_column_row_uids {
        insert_column_uid(
            package,
            &locations,
            reference.identifier,
            old_columns,
            column,
        )?;
    }
    if let Some(reference) = &descriptor.model.stroke_sidecar {
        insert_stroke_column(package, &locations, reference.identifier, new_columns_u32)?;
    }
    set_table_column_count(package, &locations, descriptor.object_id, new_columns_u32)?;
    if attached_table_descriptor(package, table_id)?
        .model
        .number_of_columns
        != new_columns_u32
    {
        return Err(Error::InvalidFormat(
            "iWork inserted column failed dimension validation".to_owned(),
        ));
    }
    Ok(new_columns)
}

fn validate_column_insertion_features(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    column: usize,
) -> Result<()> {
    let header_columns = model.number_of_header_columns.unwrap_or(0) as usize;
    if column < header_columns {
        return Err(Error::ParseError(
            "Inserting inside iWork header columns is not yet supported".to_owned(),
        ));
    }
    if model.number_of_hidden_columns.unwrap_or(0) != 0
        || model.number_of_user_hidden_columns.unwrap_or(0) != 0
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
        return Err(Error::ParseError(
            "Cannot yet insert a column into a sorted, filtered, hidden, merged, grouped, pivot, or spill iWork table"
                .to_owned(),
        ));
    }
    if model.base_column_row_uids.is_none() {
        return Err(Error::ParseError(
            "Cannot safely insert an iWork column without a stable column UID map".to_owned(),
        ));
    }
    Ok(())
}
