//! Transactional deletion of section-relative table rows and columns.

use super::*;

mod column;
mod dimension;
mod row;
mod uid;

use crate::table_hidden_axes::remove_table_hidden_axis;
use cell_merge::{
    MergeAnchorRelocation, MergeAxis, merge_anchor_relocations_for_axis_deletion,
    regions_in_package, shift_merges_for_axis_deletion,
};
use column::{delete_column_headers, delete_table_tile_column};
use dimension::set_table_dimensions;
use formula_dependency_shift::{
    DependencyAxis, FormulaHostCoordinate, delete_formula_dependencies,
};
use litchi_iwa_common::table::axis::AxisIndex;
use row::{delete_row_headers, delete_table_tile_row};
use stroke_layers::{StrokeAxis, delete as delete_stroke_layers};
use table_headers::set_attached_table_header_settings;
use table_sort::{delete_table_sort_column, validate_table_sort_order_for_topology};
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
    /// Delete one row from a semantic table section and compact following rows.
    ///
    /// Stored values, comments, formula records, headers, stable UIDs, and
    /// dimension sidecars are removed or shifted together. Formulas that still
    /// reference the deleted row cause the entire operation to fail unchanged.
    /// Full-table sort rules are preserved because row edits do not change
    /// their physical column slots.
    pub fn remove_table_row(
        &mut self,
        selector: litchi_numbers::TableSelector<'_>,
        deletion: RowDeletion,
    ) -> Result<()> {
        let table_id = super::selectors::table_id(self, selector)?;
        let mut staged = self.package.clone();
        let (new_rows, columns) = remove_attached_table_row(&mut staged, table_id, deletion)?;
        verify_numbers_dimensions(&staged, table_id, new_rows, columns)?;
        self.package = staged;
        Ok(())
    }

    /// Delete one column from a semantic table section and compact following columns.
    ///
    /// Stored values, comments, formula records, headers, stable UIDs, and
    /// dimension sidecars are removed or shifted together. Formulas that still
    /// reference the deleted column cause the entire operation to fail
    /// unchanged. Sort rules whose physical slot disappears are removed while
    /// all other rule indices remain unchanged, matching Numbers.
    pub fn remove_table_column(
        &mut self,
        selector: litchi_numbers::TableSelector<'_>,
        deletion: ColumnDeletion,
    ) -> Result<()> {
        let table_id = super::selectors::table_id(self, selector)?;
        let mut staged = self.package.clone();
        let (rows, new_columns) = remove_attached_table_column(&mut staged, table_id, deletion)?;
        verify_numbers_dimensions(&staged, table_id, rows, new_columns)?;
        self.package = staged;
        Ok(())
    }
}

pub(super) fn remove_attached_table_row(
    package: &mut IWorkPackage,
    table_id: u64,
    deletion: RowDeletion,
) -> Result<(usize, usize)> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let old_rows = descriptor.model.number_of_rows as usize;
    let resolved = resolve_row_deletion(&descriptor.model, deletion)?;
    let row = resolved.physical_index;
    let new_rows = old_rows
        .checked_sub(1)
        .ok_or_else(|| Error::ParseError("iWork tables must retain at least one row".to_owned()))?;
    let (new_rows_u32, columns_u32) =
        validate_table_dimensions(new_rows, descriptor.model.number_of_columns as usize)?;
    let updated_header_settings = resolved.updated_header_settings;
    let locations = object_locations(package)?;
    validate_table_sort_order_for_topology(package, table_id)?;
    validate_deletion_features(
        package,
        &locations,
        &descriptor.model,
        TableAxis::Row,
        new_rows,
        updated_header_settings,
    )?;
    let cells = stored_cells_on_axis(package, &locations, &descriptor.model, TableAxis::Row, row)?;
    let relocations =
        merge_anchor_relocations_for_axis_deletion(package, table_id, MergeAxis::Row, row)?;
    let relocated_sources = validated_merge_anchor_sources(&relocations)?;
    let retained_formula_hosts = merge_anchor_formula_hosts(package, table_id, &relocations)?;

    clear_stored_cells(
        package,
        table_id,
        &without_relocated_cells(cells, &relocated_sources),
    )?;
    delete_formula_dependencies(
        package,
        descriptor.table_info_id,
        DependencyAxis::Row,
        u32::try_from(row).map_err(|_| Error::ParseError("iWork row exceeds u32".to_owned()))?,
        retained_formula_hosts,
    )?;
    relocate_merge_anchors(package, table_id, &relocations)?;
    shift_merges_for_axis_deletion(package, table_id, MergeAxis::Row, row)?;
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
    if descriptor.model.hidden_states_owner.is_some() {
        remove_table_hidden_axis(package, table_id, AxisIndex::row(row))?;
    }
    delete_row_uid(package, &locations, uid.identifier, old_rows, row)?;
    if let Some(sidecar) = &descriptor.model.stroke_sidecar {
        delete_stroke_layers(
            package,
            &locations,
            sidecar.identifier,
            StrokeAxis::Row,
            row,
            descriptor.model.number_of_rows,
            columns_u32,
        )?;
    }
    set_table_dimensions(package, &locations, table_id, new_rows_u32, columns_u32)?;
    if let Some(settings) = updated_header_settings {
        set_attached_table_header_settings(package, table_id, settings)?;
    }
    verify_attached_dimensions(package, table_id, new_rows, columns_u32 as usize)?;
    regions_in_package(package, table_id)?;
    Ok((new_rows, columns_u32 as usize))
}

pub(super) fn remove_attached_table_column(
    package: &mut IWorkPackage,
    table_id: u64,
    deletion: ColumnDeletion,
) -> Result<(usize, usize)> {
    let descriptor = attached_table_descriptor(package, table_id)?;
    let old_columns = descriptor.model.number_of_columns as usize;
    let resolved = resolve_column_deletion(&descriptor.model, deletion)?;
    let column = resolved.physical_index;
    let new_columns = old_columns.checked_sub(1).ok_or_else(|| {
        Error::ParseError("iWork tables must retain at least one column".to_owned())
    })?;
    let (rows_u32, new_columns_u32) =
        validate_table_dimensions(descriptor.model.number_of_rows as usize, new_columns)?;
    let updated_header_settings = resolved.updated_header_settings;
    let locations = object_locations(package)?;
    validate_table_sort_order_for_topology(package, table_id)?;
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
    let relocations =
        merge_anchor_relocations_for_axis_deletion(package, table_id, MergeAxis::Column, column)?;
    let relocated_sources = validated_merge_anchor_sources(&relocations)?;
    let retained_formula_hosts = merge_anchor_formula_hosts(package, table_id, &relocations)?;

    clear_stored_cells(
        package,
        table_id,
        &without_relocated_cells(cells, &relocated_sources),
    )?;
    delete_formula_dependencies(
        package,
        descriptor.table_info_id,
        DependencyAxis::Column,
        u32::try_from(column)
            .map_err(|_| Error::ParseError("iWork column exceeds u32".to_owned()))?,
        retained_formula_hosts,
    )?;
    relocate_merge_anchors(package, table_id, &relocations)?;
    shift_merges_for_axis_deletion(package, table_id, MergeAxis::Column, column)?;
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
    if descriptor.model.hidden_states_owner.is_some() {
        remove_table_hidden_axis(package, table_id, AxisIndex::column(column))?;
    }
    delete_column_uid(package, &locations, uid.identifier, old_columns, column)?;
    if let Some(sidecar) = &descriptor.model.stroke_sidecar {
        delete_stroke_layers(
            package,
            &locations,
            sidecar.identifier,
            StrokeAxis::Column,
            column,
            rows_u32,
            descriptor.model.number_of_columns,
        )?;
    }
    delete_table_sort_column(package, table_id, column, new_columns)?;
    set_table_dimensions(package, &locations, table_id, rows_u32, new_columns_u32)?;
    if let Some(settings) = updated_header_settings {
        set_attached_table_header_settings(package, table_id, settings)?;
    }
    verify_attached_dimensions(package, table_id, rows_u32 as usize, new_columns)?;
    regions_in_package(package, table_id)?;
    Ok((rows_u32 as usize, new_columns))
}

fn validate_deletion_features(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    axis: TableAxis,
    new_length: usize,
    updated_header_settings: Option<HeaderSettings>,
) -> Result<()> {
    let stored_settings = table_headers::settings_from_model(model)?;
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
    if model.number_of_filtered_rows.unwrap_or(0) != 0
        || model.pivot_owner.is_some()
        || filter_has_row_state(package, locations, model.row_filter_set_pre_pivot.as_ref())?
        || category_grouping_is_enabled(package, locations, model.category_owner.as_ref())?
    {
        return Err(Error::ParseError(format!(
            "Cannot yet delete a {} from a filtered, grouped, pivot, or spill iWork table",
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

struct ResolvedRowDeletion {
    physical_index: usize,
    updated_header_settings: Option<HeaderSettings>,
}

fn resolve_row_deletion(
    model: &TableModelArchive,
    deletion: RowDeletion,
) -> Result<ResolvedRowDeletion> {
    let mut settings = table_headers::settings_from_model(model)?;
    let rows = model.number_of_rows as usize;
    let header_rows = settings.header_row_count();
    let footer_rows = settings.footer_row_count();
    let fixed_rows = header_rows
        .checked_add(footer_rows)
        .filter(|&fixed| fixed <= rows)
        .ok_or_else(|| {
            Error::InvalidFormat(
                "iWork header and footer rows exceed the table row count".to_owned(),
            )
        })?;
    let body_rows = rows - fixed_rows;
    match deletion {
        RowDeletion::Header { index } => {
            validate_section_deletion(index, header_rows, "header row")?;
            settings.header_rows = decremented_header_count(header_rows)?;
            Ok(ResolvedRowDeletion {
                physical_index: index,
                updated_header_settings: Some(settings),
            })
        },
        RowDeletion::Body { index } => {
            validate_section_deletion(index, body_rows, "body row")?;
            Ok(ResolvedRowDeletion {
                physical_index: header_rows + index,
                updated_header_settings: None,
            })
        },
        RowDeletion::Footer { index } => {
            validate_section_deletion(index, footer_rows, "footer row")?;
            settings.footer_rows = decremented_header_count(footer_rows)?;
            Ok(ResolvedRowDeletion {
                physical_index: rows - footer_rows + index,
                updated_header_settings: Some(settings),
            })
        },
    }
}

struct ResolvedColumnDeletion {
    physical_index: usize,
    updated_header_settings: Option<HeaderSettings>,
}

fn resolve_column_deletion(
    model: &TableModelArchive,
    deletion: ColumnDeletion,
) -> Result<ResolvedColumnDeletion> {
    let mut settings = table_headers::settings_from_model(model)?;
    let columns = model.number_of_columns as usize;
    let header_columns = settings.header_column_count();
    let body_columns = columns.checked_sub(header_columns).ok_or_else(|| {
        Error::InvalidFormat("iWork header columns exceed the table column count".to_owned())
    })?;
    match deletion {
        ColumnDeletion::Header { index } => {
            validate_section_deletion(index, header_columns, "header column")?;
            settings.header_columns = decremented_header_count(header_columns)?;
            Ok(ResolvedColumnDeletion {
                physical_index: index,
                updated_header_settings: Some(settings),
            })
        },
        ColumnDeletion::Body { index } => {
            validate_section_deletion(index, body_columns, "body column")?;
            Ok(ResolvedColumnDeletion {
                physical_index: header_columns + index,
                updated_header_settings: None,
            })
        },
    }
}

fn validate_section_deletion(index: usize, length: usize, section: &str) -> Result<()> {
    if index >= length {
        return Err(Error::ParseError(format!(
            "Cannot delete iWork {section} {index} from a section with {length} entries"
        )));
    }
    Ok(())
}

fn decremented_header_count(count: usize) -> Result<Option<HeaderCount>> {
    match count.checked_sub(1) {
        Some(0) => Ok(None),
        Some(count) => Ok(Some(HeaderCount::new(count)?)),
        None => Err(Error::InvalidFormat(
            "iWork header/footer deletion underflow".to_owned(),
        )),
    }
}

/// Move native merge anchors out of a deleted leading boundary before tile
/// compaction removes their original cell storage.
fn relocate_merge_anchors(
    package: &mut IWorkPackage,
    table_id: u64,
    relocations: &[MergeAnchorRelocation],
) -> Result<()> {
    for relocation in relocations {
        model::relocate_attached_cell_in_package(
            package,
            table_id,
            relocation.source_row,
            relocation.source_column,
            relocation.destination_row,
            relocation.destination_column,
        )?;
    }
    Ok(())
}

/// Validate a non-overlapping merge-anchor relocation plan and return every
/// source cell that must not be cleared before it moves.
fn validated_merge_anchor_sources(
    relocations: &[MergeAnchorRelocation],
) -> Result<HashSet<(usize, usize)>> {
    let mut planned_sources = HashSet::with_capacity(relocations.len());
    let mut planned_destinations = HashSet::with_capacity(relocations.len());
    for relocation in relocations {
        let source = (relocation.source_row, relocation.source_column);
        let destination = (relocation.destination_row, relocation.destination_column);
        if !planned_sources.insert(source) || !planned_destinations.insert(destination) {
            return Err(Error::InvalidFormat(
                "iWork merged-cell deletion has overlapping anchor relocations".to_owned(),
            ));
        }
    }
    if planned_sources
        .iter()
        .any(|source| planned_destinations.contains(source))
    {
        return Err(Error::InvalidFormat(
            "iWork merged-cell deletion relocates an anchor onto another anchor".to_owned(),
        ));
    }
    Ok(planned_sources)
}

/// Identify formula anchors whose dependency hosts must survive the deleted
/// table axis until their exact cell payload is relocated.
fn merge_anchor_formula_hosts(
    package: &IWorkPackage,
    table_id: u64,
    relocations: &[MergeAnchorRelocation],
) -> Result<Vec<FormulaHostCoordinate>> {
    let mut hosts = Vec::new();
    for relocation in relocations {
        if model::attached_cell_is_formula(
            package,
            table_id,
            relocation.source_row,
            relocation.source_column,
        )? {
            hosts.push(FormulaHostCoordinate::from_table_coordinates(
                relocation.source_row,
                relocation.source_column,
            )?);
        }
    }
    Ok(hosts)
}

fn without_relocated_cells(
    cells: Vec<(usize, usize, bool)>,
    relocated: &HashSet<(usize, usize)>,
) -> Vec<(usize, usize, bool)> {
    cells
        .into_iter()
        .filter(|(row, column, _)| !relocated.contains(&(*row, *column)))
        .collect()
}

fn stored_cells_on_axis(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
    axis: TableAxis,
    target: usize,
) -> Result<Vec<(usize, usize, bool)>> {
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
                    for (column, cell) in stored.iter().take(columns).enumerate() {
                        if let Some(cell) = cell {
                            cells.push((
                                row,
                                column,
                                BncCell::parse(cell)?.comment_identifier().is_some(),
                            ));
                        }
                    }
                },
                TableAxis::Column => {
                    if let Some(cell) = &stored[target] {
                        cells.push((
                            row,
                            target,
                            BncCell::parse(cell)?.comment_identifier().is_some(),
                        ));
                    }
                },
            }
        }
    }
    cells.sort_unstable_by_key(|&(row, column, _)| (row, column));
    Ok(cells)
}

fn clear_stored_cells(
    package: &mut IWorkPackage,
    table_id: u64,
    cells: &[(usize, usize, bool)],
) -> Result<()> {
    for &(row, column, has_comment) in cells {
        if has_comment {
            clear_attached_cell_comment_in_package(package, table_id, row, column)?;
        }
        set_attached_cell_in_package(package, table_id, row, column, CellValue::Empty)?;
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
