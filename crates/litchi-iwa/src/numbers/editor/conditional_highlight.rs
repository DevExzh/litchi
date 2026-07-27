//! Conditional-highlight references stored by table cells.

use prost::Message;

use super::*;

pub(super) fn info_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellConditionalHighlightInfo>> {
    let location = locate_cell(package, table_id, row, column)?;
    info_at_location(package, location, table_id, row, column)
}

pub(super) fn attached_info_in_package(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellConditionalHighlightInfo>> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    info_at_location(package, location, table_id, row, column)
}

fn info_at_location(
    package: &IWorkPackage,
    location: CellLocation,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<TableCellConditionalHighlightInfo>> {
    let Some(cell) = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(None);
    };
    let Some(list_identifier) = BncCell::parse(&cell)?.conditional_style_identifier() else {
        return Ok(None);
    };
    let (_resolved, entry) = resolve_entry(package, &location, list_identifier)?;
    let style_set_object_id = entry
        .entry
        .reference
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight entry {list_identifier} has no style-set reference"
            ))
        })?;
    let archive_name = location
        .object_locations
        .get(&style_set_object_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight style set {style_set_object_id} is missing"
            ))
        })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(style_set_object_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork conditional-highlight style set {style_set_object_id} is missing"
        ))
    })?;
    let style_set = object
        .messages
        .iter()
        .find_map(|message| {
            (message.type_ == 6_010)
                .then(|| tst::ConditionalStyleSetArchive::decode(message.data.as_slice()))
        })
        .transpose()?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight object {style_set_object_id} has no style-set payload"
            ))
        })?;
    Ok(Some(TableCellConditionalHighlightInfo {
        table_id,
        row,
        column,
        list_identifier,
        style_set_object_id,
        rule_count: style_set.rule_count,
    }))
}

pub(super) fn clear_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    let location = locate_cell(package, table_id, row, column)?;
    clear_at_location(package, location, row, column)
}

pub(super) fn clear_attached_in_package(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<()> {
    let location = locate_attached_cell(package, table_id, row, column)?;
    clear_at_location(package, location, row, column)
}

fn clear_at_location(
    package: &mut IWorkPackage,
    location: CellLocation,
    row: usize,
    column: usize,
) -> Result<()> {
    let Some(cell) = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(());
    };
    let Some(list_identifier) = BncCell::parse(&cell)?.conditional_style_identifier() else {
        return Ok(());
    };
    let locations = location.object_locations.clone();
    let (resolved, entry) = resolve_entry(package, &location, list_identifier)?;
    let style_set_object_id = entry
        .entry
        .reference
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight entry {list_identifier} has no style-set reference"
            ))
        })?;
    let removed = decrement_table_data_list_entry(
        package,
        &locations,
        &resolved,
        &entry,
        tst::table_data_list::ListType::ConditionalStyle,
    )?;
    update_cell(package, &location, row, column, None)?;
    let mut modified_entries = vec![location.tile_archive.clone(), resolved.table_archive];
    if removed {
        let style_archive = locations
            .get(&style_set_object_id)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork conditional-highlight style set {style_set_object_id} is missing"
                ))
            })?
            .clone();
        if let Some(component) = component_identifier_for_entry(package, &style_archive)? {
            remove_component_external_references_to_object(
                package,
                component,
                style_set_object_id,
            )?;
            remove_component_object_uuids(package, component, &[style_set_object_id])?;
        }
        remove_object_or_empty_entry(package, &locations, style_set_object_id)?;
        if !modified_entries.contains(&style_archive) {
            modified_entries.push(style_archive);
        }
        release_package_identifier_suffix(package, &[style_set_object_id])?;
    }
    advance_save_tokens_for_entries(package, &modified_entries)
}

fn resolve_entry(
    package: &IWorkPackage,
    location: &CellLocation,
    list_identifier: u32,
) -> Result<(ResolvedTableDataList, LocatedTableDataListEntry)> {
    let table_id = location
        .descriptor
        .model
        .base_data_store
        .conditionalstyletable
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork cell references conditional-highlight entry {list_identifier}, but its table has no conditional-style list"
            ))
        })?;
    let resolved = resolve_table_data_list(
        package,
        &location.object_locations,
        table_id,
        tst::table_data_list::ListType::ConditionalStyle,
    )?;
    let entry = resolved
        .entries
        .iter()
        .find(|entry| entry.entry.key == list_identifier)
        .cloned()
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-style list has no entry {list_identifier}"
            ))
        })?;
    Ok((resolved, entry))
}

fn update_cell(
    package: &mut IWorkPackage,
    location: &CellLocation,
    row: usize,
    column: usize,
    identifier: Option<u32>,
) -> Result<()> {
    let cell_count = update_tile(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
        location.descriptor.model.number_of_columns as usize,
        EncodedValue::ConditionalStyle(identifier),
    )?;
    update_row_header(
        package,
        &location.object_locations,
        &location.descriptor.model,
        row,
        cell_count,
    )
}
