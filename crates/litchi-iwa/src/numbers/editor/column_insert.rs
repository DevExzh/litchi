//! Transactional insertion of physical table columns.

use super::*;

mod storage;

use cell_merge::{MergeAxis, shift_merges_for_axis_insertion};
use formula_dependency_shift::{DependencyAxis, FooterRangeInsertion, shift_formula_dependencies};
use storage::{
    insert_column_uid, set_table_column_count, shift_column_headers, shift_table_tile_columns,
};
use stroke_layers::{StrokeAxis, insert as insert_stroke_layers};
use table_headers::set_attached_table_header_settings;
use table_sort::validate_table_sort_order_for_topology;
use table_topology::{category_grouping_is_enabled, filter_has_row_state};

impl NumbersEditor {
    /// Insert one blank column at a section-relative position in a table.
    ///
    /// Stored cells, column metadata, stable column UIDs, header counts, and
    /// ordinary formula dependency coordinates are shifted in lockstep.
    /// Configured full-table sort-rule indices remain on their physical slots,
    /// matching Numbers when cells shift through an inserted column.
    /// Unsupported topology is rejected transactionally without changing the
    /// package.
    pub fn insert_table_column(
        &mut self,
        table_id: u64,
        insertion: TableColumnInsertion,
    ) -> Result<()> {
        let mut staged = self.package.clone();
        let new_columns = insert_attached_table_column(&mut staged, table_id, insertion)?;
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
    insertion: TableColumnInsertion,
) -> Result<usize> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let old_columns = descriptor.model.number_of_columns as usize;
    let resolved = resolve_column_insertion(&descriptor.model, insertion)?;
    let column = resolved.physical_index;
    let new_columns = old_columns
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork column count overflow".to_owned()))?;
    let (_, new_columns_u32) =
        validate_table_dimensions(descriptor.model.number_of_rows as usize, new_columns)?;
    let locations = object_locations(package)?;
    validate_table_sort_order_for_topology(package, table_id)?;
    validate_column_insertion_features(package, &locations, &descriptor.model)?;

    shift_formula_dependencies(
        package,
        descriptor.table_info_id,
        DependencyAxis::Column,
        u32::try_from(column)
            .map_err(|_| Error::ParseError("iWork column exceeds u32".to_owned()))?,
        FooterRangeInsertion::FixedSection,
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
        insert_stroke_layers(
            package,
            &locations,
            reference.identifier,
            StrokeAxis::Column,
            column,
            descriptor.model.number_of_rows,
            descriptor.model.number_of_columns,
        )?;
    }
    set_table_column_count(package, &locations, descriptor.object_id, new_columns_u32)?;
    if let Some(settings) = resolved.updated_header_settings {
        set_attached_table_header_settings(package, table_id, settings)?;
    }
    shift_merges_for_axis_insertion(package, table_id, MergeAxis::Column, column)?;
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

struct ResolvedColumnInsertion {
    physical_index: usize,
    updated_header_settings: Option<NumbersTableHeaderSettings>,
}

fn resolve_column_insertion(
    model: &TableModelArchive,
    insertion: TableColumnInsertion,
) -> Result<ResolvedColumnInsertion> {
    let columns = model.number_of_columns as usize;
    let mut settings = NumbersTableHeaderSettings::from_model(model)?;
    let header_columns = settings.header_column_count();
    let body_columns = columns.checked_sub(header_columns).ok_or_else(|| {
        Error::InvalidFormat("iWork header columns exceed the table column count".to_owned())
    })?;
    match insertion {
        TableColumnInsertion::Header { index } => {
            validate_section_insertion(index, header_columns, "header column")?;
            settings.header_columns = Some(NumbersTableHeaderCount::new(header_columns + 1)?);
            Ok(ResolvedColumnInsertion {
                physical_index: index,
                updated_header_settings: Some(settings),
            })
        },
        TableColumnInsertion::Body { index } => {
            validate_section_insertion(index, body_columns, "body column")?;
            Ok(ResolvedColumnInsertion {
                physical_index: header_columns + index,
                updated_header_settings: None,
            })
        },
    }
}

fn validate_section_insertion(index: usize, length: usize, section: &str) -> Result<()> {
    if index > length {
        return Err(Error::ParseError(format!(
            "Cannot insert iWork {section} {index} into a section with {length} columns"
        )));
    }
    Ok(())
}

fn validate_column_insertion_features(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
) -> Result<()> {
    if model.number_of_filtered_rows.unwrap_or(0) != 0
        || model.pivot_owner.is_some()
        || filter_has_row_state(package, locations, model.row_filter_set_pre_pivot.as_ref())?
        || category_grouping_is_enabled(package, locations, model.category_owner.as_ref())?
    {
        return Err(Error::ParseError(
            "Cannot yet insert a column into a filtered, grouped, pivot, or spill iWork table"
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
