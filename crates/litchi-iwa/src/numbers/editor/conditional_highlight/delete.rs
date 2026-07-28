//! Reference-safe deletion of conditional-highlight graphs.

use prost::Message;

use super::*;

const CONDITIONAL_STYLE_SET_MESSAGE_TYPE: u32 = 6_010;

pub(super) fn clear_at_location(
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
    if let Some(owner_uid) = location
        .descriptor
        .model
        .conditional_style_formula_owner_id
        .as_ref()
        .and_then(cfuuid_as_uuid)
    {
        dependencies::remove_volatile_host(package, owner_uid, row, column)?;
    }
    let mut owned_object_ids =
        conditional_style_owned_object_ids(package, &locations, style_set_object_id)?;
    owned_object_ids.push(style_set_object_id);
    let removed = decrement_table_data_list_entry(
        package,
        &locations,
        &resolved,
        &entry,
        tst::table_data_list::ListType::ConditionalStyle,
    )?;
    update_cell(package, &location, row, column, None, None)?;
    let mut modified_entries = vec![location.tile_archive.clone(), resolved.table_archive];
    if removed {
        let style_archive = locations.get(&style_set_object_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight style set {style_set_object_id} is missing"
            ))
        })?;
        if let Some(component) = component_identifier_for_entry(package, style_archive)? {
            for identifier in &owned_object_ids {
                remove_component_external_references_to_object(package, component, *identifier)?;
            }
            remove_component_object_uuids(package, component, &owned_object_ids)?;
        }
        for identifier in &owned_object_ids {
            if let Some(archive_name) = locations.get(identifier)
                && !modified_entries.contains(archive_name)
            {
                modified_entries.push(archive_name.clone());
            }
            remove_object_or_empty_entry(package, &locations, *identifier)?;
        }
        release_package_identifier_suffix(package, &owned_object_ids)?;
    }
    advance_save_tokens_for_entries(package, &modified_entries)
}

fn conditional_style_owned_object_ids(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    style_set_id: u64,
) -> Result<Vec<u64>> {
    let archive_name = locations.get(&style_set_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork conditional-highlight style set {style_set_id} is missing"
        ))
    })?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(style_set_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "iWork conditional-highlight style set {style_set_id} is missing"
        ))
    })?;
    let set = object
        .messages
        .iter()
        .find_map(|message| {
            (message.type_ == CONDITIONAL_STYLE_SET_MESSAGE_TYPE)
                .then(|| tst::ConditionalStyleSetArchive::decode(message.data.as_slice()))
        })
        .transpose()?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight object {style_set_id} has no style-set payload"
            ))
        })?;
    let mut identifiers = Vec::with_capacity(set.rules_prepivot.len() * 2);
    for rule in set.rules_prepivot {
        for identifier in [rule.cell_style.identifier, rule.text_style.identifier] {
            if identifier != 0 && !identifiers.contains(&identifier) {
                identifiers.push(identifier);
            }
        }
    }
    if let Some(rules) = set.rules {
        for rule in rules.rule {
            for identifier in [rule.cell_style.identifier, rule.text_style.identifier] {
                if identifier != 0 && !identifiers.contains(&identifier) {
                    identifiers.push(identifier);
                }
            }
        }
    }
    Ok(identifiers)
}
