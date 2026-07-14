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
        let descriptor = table_models(&self.package)?
            .into_iter()
            .find(|table| table.object_id == table_id)
            .ok_or_else(|| {
                Error::ParseError(format!("Numbers table object {table_id} not found"))
            })?;
        let old_columns = descriptor.model.number_of_columns as usize;
        if column > old_columns {
            return Err(Error::ParseError(format!(
                "Cannot insert Numbers column {column} into a table with {old_columns} columns"
            )));
        }
        let new_columns = old_columns
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers column count overflow".to_owned()))?;
        let (_, new_columns_u32) =
            validate_table_dimensions(descriptor.model.number_of_rows as usize, new_columns)?;
        let locations = object_locations(&self.package)?;
        validate_column_insertion_features(&self.package, &locations, &descriptor.model, column)?;

        let mut staged = self.package.clone();
        shift_table_tile_columns(&mut staged, &locations, &descriptor.model, column)?;
        shift_column_headers(
            &mut staged,
            &locations,
            descriptor.model.base_data_store.column_headers.identifier,
            column,
        )?;
        if let Some(reference) = &descriptor.model.base_column_row_uids {
            insert_column_uid(
                &mut staged,
                &locations,
                reference.identifier,
                old_columns,
                column,
            )?;
        }
        if let Some(reference) = &descriptor.model.stroke_sidecar {
            insert_stroke_column(
                &mut staged,
                &locations,
                reference.identifier,
                new_columns_u32,
            )?;
        }
        shift_formula_dependencies(
            &mut staged,
            descriptor.table_info_id,
            DependencyAxis::Column,
            u32::try_from(column)
                .map_err(|_| Error::ParseError("Numbers column exceeds u32".to_owned()))?,
        )?;
        set_table_column_count(
            &mut staged,
            &locations,
            descriptor.object_id,
            new_columns_u32,
        )?;

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

fn validate_column_insertion_features(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    column: usize,
) -> Result<()> {
    let header_columns = model.number_of_header_columns.unwrap_or(0) as usize;
    if column < header_columns {
        return Err(Error::ParseError(
            "Inserting inside Numbers header columns is not yet supported".to_owned(),
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
            "Cannot yet insert a column into a sorted, filtered, hidden, merged, grouped, pivot, or spill Numbers table"
                .to_owned(),
        ));
    }
    if model.base_column_row_uids.is_none() {
        return Err(Error::ParseError(
            "Cannot safely insert a Numbers column without a stable column UID map".to_owned(),
        ));
    }
    Ok(())
}
