//! Transactional insertion of physical table rows.

use super::*;

mod dependencies;
mod storage;

use dependencies::shift_formula_dependencies;
use storage::{
    insert_row_uid, insert_stroke_row, set_table_row_count, shift_row_headers,
    shift_table_tile_rows,
};

impl NumbersEditor {
    /// Insert one blank physical row before `row` in a table.
    ///
    /// `row == table.rows` appends a row. Stored cells, row metadata, stable
    /// row UIDs, and ordinary formula dependency coordinates are shifted in
    /// lockstep. The operation is transactional: table features whose row
    /// topology cannot yet be rewritten safely are rejected without changing
    /// the package.
    pub fn insert_table_row(&mut self, table_id: u64, row: usize) -> Result<()> {
        let descriptor = table_models(&self.package)?
            .into_iter()
            .find(|table| table.object_id == table_id)
            .ok_or_else(|| {
                Error::ParseError(format!("Numbers table object {table_id} not found"))
            })?;
        let old_rows = descriptor.model.number_of_rows as usize;
        if row > old_rows {
            return Err(Error::ParseError(format!(
                "Cannot insert Numbers row {row} into a table with {old_rows} rows"
            )));
        }
        let new_rows = old_rows
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers row count overflow".to_owned()))?;
        let (new_rows_u32, _) =
            validate_table_dimensions(new_rows, descriptor.model.number_of_columns as usize)?;
        let locations = object_locations(&self.package)?;
        validate_row_insertion_features(
            &self.package,
            &locations,
            &descriptor.model,
            row,
            old_rows,
        )?;
        let mut staged = self.package.clone();
        shift_table_tile_rows(&mut staged, &locations, &descriptor.model, row)?;
        shift_row_headers(&mut staged, &locations, &descriptor.model, row)?;
        if let Some(reference) = &descriptor.model.base_column_row_uids {
            insert_row_uid(&mut staged, &locations, reference.identifier, old_rows, row)?;
        }
        if let Some(reference) = &descriptor.model.stroke_sidecar {
            insert_stroke_row(&mut staged, &locations, reference.identifier, new_rows_u32)?;
        }
        shift_formula_dependencies(
            &mut staged,
            descriptor.table_info_id,
            u32::try_from(row)
                .map_err(|_| Error::ParseError("Numbers row exceeds u32".to_owned()))?,
        )?;
        set_table_row_count(&mut staged, &locations, descriptor.object_id, new_rows_u32)?;

        let verified = NumbersEditor::from_bytes(&staged.to_bytes()?)?;
        let table = verified
            .tables()?
            .into_iter()
            .find(|table| table.object_id == table_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers table disappeared after row insertion".to_owned())
            })?;
        if table.rows != new_rows {
            return Err(Error::InvalidFormat(
                "Numbers inserted row failed dimension validation".to_owned(),
            ));
        }
        self.package = staged;
        Ok(())
    }
}

fn validate_row_insertion_features(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    row: usize,
    old_rows: usize,
) -> Result<()> {
    let header_rows = model.number_of_header_rows.unwrap_or(0) as usize;
    let footer_rows = model.number_of_footer_rows.unwrap_or(0) as usize;
    if row < header_rows {
        return Err(Error::ParseError(
            "Inserting inside Numbers header rows is not yet supported".to_owned(),
        ));
    }
    if footer_rows > 0 && row > old_rows.saturating_sub(footer_rows) {
        return Err(Error::ParseError(
            "Inserting inside Numbers footer rows is not yet supported".to_owned(),
        ));
    }
    if model.number_of_hidden_rows.unwrap_or(0) != 0
        || model.number_of_user_hidden_rows.unwrap_or(0) != 0
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
            "Cannot yet insert a row into a sorted, filtered, hidden, merged, grouped, pivot, or spill Numbers table"
                .to_owned(),
        ));
    }
    if model.base_column_row_uids.is_none() {
        return Err(Error::ParseError(
            "Cannot safely insert a Numbers row without a stable row UID map".to_owned(),
        ));
    }
    Ok(())
}

fn filter_has_row_state(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    reference: Option<&tsp::Reference>,
) -> Result<bool> {
    let Some(reference) = reference else {
        return Ok(false);
    };
    let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers filter object {} is missing",
            reference.identifier
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(reference.identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers filter object {} is missing",
            reference.identifier
        ))
    })?;
    let filter = object
        .messages
        .iter()
        .find_map(|message| tst::FilterSetArchive::decode(message.data.as_slice()).ok())
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Object {} has no Numbers filter-set payload",
                reference.identifier
            ))
        })?;
    Ok(filter.is_enabled.unwrap_or(true)
        && (!filter.filter_rules_prepivot.is_empty()
            || !filter.filter_rules.is_empty()
            || filter.filter_enabled.iter().any(|enabled| *enabled)))
}

fn category_grouping_is_enabled(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    reference: Option<&tsp::Reference>,
) -> Result<bool> {
    let Some(reference) = reference else {
        return Ok(false);
    };
    let archive_name = locations.get(&reference.identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers category owner {} is missing",
            reference.identifier
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(reference.identifier).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers category owner {} is missing",
            reference.identifier
        ))
    })?;
    let references = object
        .messages
        .iter()
        .find_map(|message| tst::CategoryOwnerRefArchive::decode(message.data.as_slice()).ok())
        .map(|owner| owner.group_by)
        .unwrap_or_default();
    for group in references {
        let archive_name = locations.get(&group.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers group-by object {} is missing",
                group.identifier
            ))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(group.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers group-by object {} is missing",
                group.identifier
            ))
        })?;
        let enabled = object.messages.iter().any(|message| {
            tst::GroupByArchive::decode(message.data.as_slice()).is_ok_and(|group| group.is_enabled)
        });
        if enabled {
            return Ok(true);
        }
    }
    Ok(false)
}
