//! Transactional insertion of physical table rows.

use super::*;

mod storage;

use cell_merge::{MergeAxis, shift_merges_for_axis_insertion};
use formula_dependency_shift::{DependencyAxis, FooterRangeInsertion, shift_formula_dependencies};
use storage::{insert_row_uid, set_table_row_count, shift_row_headers, shift_table_tile_rows};
use stroke_layers::{StrokeAxis, insert as insert_stroke_layers};
use table_headers::set_attached_table_header_settings;
use table_sort::validate_table_sort_order_for_topology;
use table_topology::{category_grouping_is_enabled, filter_has_row_state};

impl NumbersEditor {
    /// Insert one blank row at a section-relative position in a table.
    ///
    /// Stored cells, row metadata, stable row UIDs, header/footer counts, and
    /// ordinary formula dependency coordinates are shifted in lockstep.
    /// Configured full-table sort rules remain attached to their physical
    /// column slots. The operation is transactional: table features whose row
    /// topology cannot yet be rewritten safely are rejected without changing
    /// the package.
    pub fn insert_table_row(&mut self, table_id: u64, insertion: TableRowInsertion) -> Result<()> {
        let mut staged = self.package.clone();
        let new_rows = insert_attached_table_row(&mut staged, table_id, insertion)?;
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

pub(super) fn insert_attached_table_row(
    package: &mut IWorkPackage,
    table_id: u64,
    insertion: TableRowInsertion,
) -> Result<usize> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let old_rows = descriptor.model.number_of_rows as usize;
    let resolved = resolve_row_insertion(&descriptor.model, insertion)?;
    let row = resolved.physical_index;
    let new_rows = old_rows
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork row count overflow".to_owned()))?;
    let (new_rows_u32, _) =
        validate_table_dimensions(new_rows, descriptor.model.number_of_columns as usize)?;
    let locations = object_locations(package)?;
    validate_table_sort_order_for_topology(package, table_id)?;
    validate_row_insertion_features(package, &locations, &descriptor.model)?;
    shift_formula_dependencies(
        package,
        descriptor.table_info_id,
        DependencyAxis::Row,
        u32::try_from(row).map_err(|_| Error::ParseError("iWork row exceeds u32".to_owned()))?,
        resolved.footer_range_insertion,
    )?;
    shift_table_tile_rows(package, &locations, &descriptor.model, row)?;
    shift_row_headers(package, &locations, &descriptor.model, row)?;
    if let Some(reference) = &descriptor.model.base_column_row_uids {
        insert_row_uid(package, &locations, reference.identifier, old_rows, row)?;
    }
    if let Some(reference) = &descriptor.model.stroke_sidecar {
        insert_stroke_layers(
            package,
            &locations,
            reference.identifier,
            StrokeAxis::Row,
            row,
            descriptor.model.number_of_rows,
            descriptor.model.number_of_columns,
        )?;
    }
    set_table_row_count(package, &locations, descriptor.object_id, new_rows_u32)?;
    if let Some(settings) = resolved.updated_header_settings {
        set_attached_table_header_settings(package, table_id, settings)?;
    }
    shift_merges_for_axis_insertion(package, table_id, MergeAxis::Row, row)?;
    if attached_table_descriptor(package, table_id)?
        .model
        .number_of_rows
        != new_rows_u32
    {
        return Err(Error::InvalidFormat(
            "iWork inserted row failed dimension validation".to_owned(),
        ));
    }
    Ok(new_rows)
}

struct ResolvedRowInsertion {
    physical_index: usize,
    footer_range_insertion: FooterRangeInsertion,
    updated_header_settings: Option<NumbersTableHeaderSettings>,
}

fn resolve_row_insertion(
    model: &TableModelArchive,
    insertion: TableRowInsertion,
) -> Result<ResolvedRowInsertion> {
    let rows = model.number_of_rows as usize;
    let mut settings = NumbersTableHeaderSettings::from_model(model)?;
    let header_rows = settings.header_row_count();
    let footer_rows = settings.footer_row_count();
    let body_rows = rows
        .checked_sub(header_rows)
        .and_then(|rows| rows.checked_sub(footer_rows))
        .ok_or_else(|| {
            Error::InvalidFormat(
                "iWork header and footer rows exceed the table row count".to_owned(),
            )
        })?;
    match insertion {
        TableRowInsertion::Header { index } => {
            validate_section_insertion(index, header_rows, "header row")?;
            settings.header_rows = Some(NumbersTableHeaderCount::new(header_rows + 1)?);
            Ok(ResolvedRowInsertion {
                physical_index: index,
                footer_range_insertion: FooterRangeInsertion::FixedSection,
                updated_header_settings: Some(settings),
            })
        },
        TableRowInsertion::Body { index } => {
            validate_section_insertion(index, body_rows, "body row")?;
            Ok(ResolvedRowInsertion {
                physical_index: header_rows + index,
                footer_range_insertion: FooterRangeInsertion::Body,
                updated_header_settings: None,
            })
        },
        TableRowInsertion::Footer { index } => {
            validate_section_insertion(index, footer_rows, "footer row")?;
            settings.footer_rows = Some(NumbersTableHeaderCount::new(footer_rows + 1)?);
            Ok(ResolvedRowInsertion {
                physical_index: rows - footer_rows + index,
                footer_range_insertion: FooterRangeInsertion::FixedSection,
                updated_header_settings: Some(settings),
            })
        },
    }
}

fn validate_section_insertion(index: usize, length: usize, section: &str) -> Result<()> {
    if index > length {
        return Err(Error::ParseError(format!(
            "Cannot insert iWork {section} {index} into a section with {length} rows"
        )));
    }
    Ok(())
}

fn validate_row_insertion_features(
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
            "Cannot yet insert a row into a filtered, grouped, pivot, or spill iWork table"
                .to_owned(),
        ));
    }
    if model.base_column_row_uids.is_none() {
        return Err(Error::ParseError(
            "Cannot safely insert an iWork row without a stable row UID map".to_owned(),
        ));
    }
    Ok(())
}
