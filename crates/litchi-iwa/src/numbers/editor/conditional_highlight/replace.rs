//! Identity-preserving replacement of existing conditional-style graphs.

use std::collections::{HashMap, HashSet};

use prost::Message;

use super::*;

const CONDITIONAL_STYLE_SET_MESSAGE_TYPE: u32 = 6_010;

pub(super) fn try_at_location(
    package: &mut IWorkPackage,
    location: &CellLocation,
    row: usize,
    column: usize,
    rules: &[Rule],
) -> Result<bool> {
    let Some(cell) = read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(false);
    };
    let Some(list_identifier) = BncCell::parse(&cell)?.conditional_style_identifier() else {
        return Ok(false);
    };
    let (resolved, entry) = resolve_entry(package, location, list_identifier)?;
    if entry.entry.refcount == 0 {
        return Err(Error::InvalidFormat(format!(
            "conditional-highlight list entry {list_identifier} has a zero reference count"
        )));
    }
    if entry.entry.refcount > 1 {
        return copy_on_write(package, location, row, column, rules, &resolved, &entry);
    }
    let style_set_id = entry
        .entry
        .reference
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight entry {list_identifier} has no style-set reference"
            ))
        })?;
    let existing =
        conditional_style_rule_identifiers(package, &location.object_locations, style_set_id)?;
    if has_duplicate_identifiers(&existing) {
        return Ok(false);
    }
    let style_archive = location
        .object_locations
        .get(&style_set_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork conditional-highlight style set {style_set_id} is missing"
            ))
        })?
        .clone();
    let replacement = replacement_identifiers(package, &existing, rules.len())?;
    let removed = removed_identifiers(&existing, rules.len());
    ownership::ensure_children_are_private(package, style_set_id, &removed)?;

    let owner_uid = location
        .descriptor
        .model
        .conditional_style_formula_owner_id
        .as_ref()
        .and_then(cfuuid_as_uuid)
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers conditional-style formula owner is missing".to_owned())
        })?;
    dependencies::remove_volatile_host(package, owner_uid, row, column)?;
    let table_uuid = parse_table_uuid(&location.descriptor.model.table_id)?;
    let formula_owner_uuid = formula_owner_uuid_for_table(&table_uuid);
    write::replace_conditional_style_graph(
        package,
        &style_archive,
        &location.object_locations,
        style_set_id,
        rules,
        &formula_owner_uuid,
        &replacement.identifiers,
    )?;
    register_added_children(
        package,
        &style_archive,
        &replacement.added,
        &replacement.identifiers,
        &location.object_locations,
    )?;
    ownership::remove_owned_objects(package, &location.object_locations, &removed)?;
    if replacement.added.is_empty() {
        release_package_identifier_suffix(package, &removed)?;
    } else if let Some(last) = replacement.added.last().copied() {
        set_package_last_object_identifier(package, last)?;
    }
    if rules
        .iter()
        .any(|rule| is_volatile_date_condition(&rule.condition))
    {
        dependencies::ensure_volatile_owner(package, &table_uuid, owner_uid, row, column)?;
    }
    let applied_rule = applied_rule_for_cell(package, location, column, rules)?;
    update_cell(
        package,
        location,
        row,
        column,
        Some(list_identifier),
        Some(applied_rule),
    )?;
    advance_replacement_save_tokens(
        package,
        location,
        &resolved.table_archive,
        &style_archive,
        &existing,
        &replacement.identifiers,
    )?;
    Ok(true)
}

fn copy_on_write(
    package: &mut IWorkPackage,
    location: &CellLocation,
    row: usize,
    column: usize,
    rules: &[Rule],
    resolved: &ResolvedTableDataList,
    entry: &LocatedTableDataListEntry,
) -> Result<bool> {
    let owner_uid = location
        .descriptor
        .model
        .conditional_style_formula_owner_id
        .as_ref()
        .and_then(cfuuid_as_uuid)
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers conditional-style formula owner is missing".to_owned())
        })?;
    dependencies::remove_volatile_host(package, owner_uid, row, column)?;
    decrement_table_data_list_entry(
        package,
        &location.object_locations,
        resolved,
        entry,
        tst::table_data_list::ListType::ConditionalStyle,
    )?;
    set_at_location(package, location, row, column, rules)?;
    Ok(true)
}

#[derive(Debug)]
struct ReplacementIdentifiers {
    identifiers: Vec<write::RuleStyleIdentifiers>,
    added: Vec<u64>,
}

fn replacement_identifiers(
    package: &IWorkPackage,
    existing: &[write::RuleStyleIdentifiers],
    rule_count: usize,
) -> Result<ReplacementIdentifiers> {
    let retained = existing.len().min(rule_count);
    let mut identifiers = existing[..retained].to_vec();
    let mut added = Vec::with_capacity(rule_count.saturating_sub(retained) * 2);
    let mut next = (rule_count > retained)
        .then(|| next_object_identifier(package))
        .transpose()?;
    for _ in retained..rule_count {
        let text_style = next.expect("growth allocates a first conditional-style identifier");
        let cell_style = text_style.checked_add(1).ok_or_else(|| {
            Error::ParseError("conditional-highlight identifier overflow".to_owned())
        })?;
        next = Some(cell_style.checked_add(1).ok_or_else(|| {
            Error::ParseError("conditional-highlight identifier overflow".to_owned())
        })?);
        identifiers.push(write::RuleStyleIdentifiers {
            text_style,
            cell_style,
        });
        added.extend([text_style, cell_style]);
    }
    Ok(ReplacementIdentifiers { identifiers, added })
}

fn removed_identifiers(existing: &[write::RuleStyleIdentifiers], rule_count: usize) -> Vec<u64> {
    existing
        .get(rule_count..)
        .unwrap_or_default()
        .iter()
        .flat_map(|identifiers| [identifiers.text_style, identifiers.cell_style])
        .collect()
}

fn has_duplicate_identifiers(identifiers: &[write::RuleStyleIdentifiers]) -> bool {
    let mut unique = HashSet::with_capacity(identifiers.len() * 2);
    identifiers.iter().any(|identifiers| {
        !unique.insert(identifiers.text_style) || !unique.insert(identifiers.cell_style)
    })
}

pub(super) fn conditional_style_rule_identifiers(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    style_set_id: u64,
) -> Result<Vec<write::RuleStyleIdentifiers>> {
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
    let identifiers = if !set.rules_prepivot.is_empty() {
        set.rules_prepivot
            .iter()
            .map(|rule| rule_style_identifiers(&rule.text_style, &rule.cell_style))
            .collect::<Result<Vec<_>>>()?
    } else {
        set.rules
            .as_ref()
            .map(|rules| {
                rules
                    .rule
                    .iter()
                    .map(|rule| rule_style_identifiers(&rule.text_style, &rule.cell_style))
                    .collect()
            })
            .transpose()?
            .unwrap_or_default()
    };
    if identifiers.len() != set.rule_count as usize {
        return Err(Error::InvalidFormat(format!(
            "conditional-highlight style set {style_set_id} declares {} rules but references {} style pairs",
            set.rule_count,
            identifiers.len()
        )));
    }
    Ok(identifiers)
}

fn rule_style_identifiers(
    text_style: &tsp::Reference,
    cell_style: &tsp::Reference,
) -> Result<write::RuleStyleIdentifiers> {
    if text_style.identifier == 0 || cell_style.identifier == 0 {
        return Err(Error::InvalidFormat(
            "conditional-highlight rule has a missing style reference".to_owned(),
        ));
    }
    Ok(write::RuleStyleIdentifiers {
        text_style: text_style.identifier,
        cell_style: cell_style.identifier,
    })
}

fn register_added_children(
    package: &mut IWorkPackage,
    style_archive: &str,
    added: &[u64],
    identifiers: &[write::RuleStyleIdentifiers],
    locations: &HashMap<u64, String>,
) -> Result<()> {
    if added.is_empty() {
        return Ok(());
    }
    let text_archive = identifiers
        .iter()
        .find_map(|identifiers| locations.get(&identifiers.text_style))
        .map(String::as_str)
        .unwrap_or(style_archive);
    let cell_archive = identifiers
        .iter()
        .find_map(|identifiers| locations.get(&identifiers.cell_style))
        .map(String::as_str)
        .unwrap_or(style_archive);
    let mut by_archive = HashMap::<&str, Vec<u64>>::new();
    for pair in identifiers {
        if added.contains(&pair.text_style) {
            by_archive
                .entry(text_archive)
                .or_default()
                .push(pair.text_style);
        }
        if added.contains(&pair.cell_style) {
            by_archive
                .entry(cell_archive)
                .or_default()
                .push(pair.cell_style);
        }
    }
    let owner_component = component_identifier_for_entry(package, style_archive)?;
    for (archive_name, identifiers) in by_archive {
        let target_component = component_identifier_for_entry(package, archive_name)?;
        if let Some(target_component) = target_component {
            add_component_object_uuids(package, target_component, &identifiers)?;
            if let Some(owner_component) = owner_component
                && owner_component != target_component
            {
                for identifier in identifiers {
                    add_component_external_reference(
                        package,
                        owner_component,
                        target_component,
                        identifier,
                    )?;
                }
            }
        }
    }
    Ok(())
}

fn advance_replacement_save_tokens(
    package: &mut IWorkPackage,
    location: &CellLocation,
    list_archive: &str,
    style_archive: &str,
    existing: &[write::RuleStyleIdentifiers],
    replacement: &[write::RuleStyleIdentifiers],
) -> Result<()> {
    let model_archive = location
        .object_locations
        .get(&location.descriptor.object_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table object {} is missing",
                location.descriptor.object_id
            ))
        })?;
    let mut modified = vec![
        location.tile_archive.clone(),
        list_archive.to_owned(),
        style_archive.to_owned(),
        model_archive.clone(),
    ];
    for identifiers in existing.iter().chain(replacement) {
        for identifier in [identifiers.text_style, identifiers.cell_style] {
            if let Some(archive_name) = location.object_locations.get(&identifier)
                && !modified.contains(archive_name)
            {
                modified.push(archive_name.clone());
            }
        }
    }
    advance_save_tokens_for_entries(package, &modified)
}
