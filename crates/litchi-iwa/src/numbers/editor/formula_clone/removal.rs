//! Transactional removal of Numbers CalculationEngine table-owner families.

use super::dependency_wire::{
    prune_cell_tile_edges_wire, prune_formula_owner_cell_edges_wire, prune_internal_owner_edges,
    remove_formula_owners_from_engine,
};
use super::*;

/// Remove one table's CalculationEngine owner family.
///
/// Incoming cross-table dependencies are rejected so deletion cannot leave a
/// surviving formula with a dangling internal owner ID.
pub(in crate::numbers::editor) fn remove_table_formula_graph(
    package: &mut IWorkPackage,
    table_info_id: u64,
) -> Result<Vec<u64>> {
    remove_table_formula_graph_for_contexts(package, &[table_info_id])
}

pub(in crate::numbers::editor) fn remove_table_formula_graph_for_contexts(
    package: &mut IWorkPackage,
    table_context_ids: &[u64],
) -> Result<Vec<u64>> {
    let contexts = table_context_ids.iter().copied().collect::<HashSet<_>>();
    if contexts.is_empty() {
        return Ok(Vec::new());
    }
    let Some(calculation_engine_entry) = calculation_engine_entry_for_contexts(package, &contexts)?
    else {
        return Ok(Vec::new());
    };
    let archive = package.archive(&calculation_engine_entry)?;
    let (owners, owner_uuids) = formula_owner_context_family(&archive, &contexts)?;
    if owners.is_empty() {
        return Ok(Vec::new());
    }
    let owner_ids = owners
        .iter()
        .map(|source| {
            source.object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat("Numbers formula owner has no object identifier".to_owned())
            })
        })
        .collect::<Result<HashSet<_>>>()?;
    let internal_owner_ids = owners
        .iter()
        .map(|source| source.owner.internal_formula_owner_id)
        .collect::<HashSet<_>>();
    prune_global_cell_dependencies(
        package,
        &calculation_engine_entry,
        &archive,
        &owner_ids,
        &internal_owner_ids,
    )?;
    let archive = package.archive(&calculation_engine_entry)?;
    reject_incoming_formula_dependencies(
        &archive,
        &contexts,
        &owner_ids,
        &internal_owner_ids,
        &owner_uuids,
    )?;
    let tiles = cell_record_tiles(&archive, &owners)?;
    let range_tiles = range_precedent_tiles(&archive, &owners)?;
    let formula_count = owners.iter().try_fold(0u64, |count, source| {
        count
            .checked_add(formula_cell_count(&source.owner, &tiles)?)
            .ok_or_else(|| Error::ParseError("Numbers formula count overflow".to_owned()))
    })?;
    let tile_ids = tiles
        .iter()
        .map(|source| {
            source.object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat("Numbers dependency tile has no object identifier".to_owned())
            })
        })
        .collect::<Result<Vec<_>>>()?;
    let range_tile_ids = range_tiles
        .iter()
        .map(|source| {
            source.object.archive_info.identifier.ok_or_else(|| {
                Error::InvalidFormat(
                    "Numbers range dependency tile has no object identifier".to_owned(),
                )
            })
        })
        .collect::<Result<Vec<_>>>()?;

    package.update_archive(&calculation_engine_entry, |archive| {
        let (engine_id, engine_message_index) = calculation_engine_location(archive)?;
        let engine_object = archive.object_mut(engine_id).ok_or_else(|| {
            Error::InvalidFormat("Numbers CalculationEngine root is missing".to_owned())
        })?;
        let original = engine_object.messages[engine_message_index].clone();
        let data = remove_formula_owners_from_engine(
            original.data.as_slice(),
            &owner_ids,
            &internal_owner_ids,
            formula_count,
        )?;
        engine_object.replace_message(
            engine_message_index,
            RawMessage {
                type_: original.type_,
                data,
            },
        )?;
        let info = &mut engine_object.archive_info.message_infos[engine_message_index];
        info.object_references
            .retain(|identifier| !owner_ids.contains(identifier));
        for field in &mut info.field_infos {
            field
                .object_references
                .retain(|identifier| !owner_ids.contains(identifier));
        }
        for identifier in owner_ids
            .iter()
            .copied()
            .chain(tile_ids.iter().copied())
            .chain(range_tile_ids.iter().copied())
        {
            archive.remove_object(identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers calculation object {identifier} disappeared during removal"
                ))
            })?;
        }
        Ok(())
    })?;

    let mut removed = owner_ids.into_iter().collect::<Vec<_>>();
    removed.extend(tile_ids);
    removed.extend(range_tile_ids);
    if let Some(component) = component_identifier_for_entry(package, &calculation_engine_entry)? {
        let mapped = component_uuid_identifiers(package, component)?.unwrap_or_default();
        let registered = removed
            .iter()
            .copied()
            .filter(|identifier| mapped.contains(identifier))
            .collect::<Vec<_>>();
        remove_component_object_uuids(package, component, &registered)?;
    }
    Ok(removed)
}

fn prune_global_cell_dependencies(
    package: &mut IWorkPackage,
    calculation_engine_entry: &str,
    archive: &Archive,
    removed_owner_ids: &HashSet<u64>,
    removed_internal_ids: &HashSet<u32>,
) -> Result<()> {
    let mut owner_updates = Vec::new();
    let mut tile_ids = HashSet::new();
    for object in &archive.objects {
        let object_id = object.archive_info.identifier.ok_or_else(|| {
            Error::InvalidFormat("iWork calculation object has no identifier".to_owned())
        })?;
        if removed_owner_ids.contains(&object_id) {
            continue;
        }
        let messages = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.len() > 1 {
            return Err(Error::InvalidFormat(
                "iWork formula owner object repeats its owner payload".to_owned(),
            ));
        }
        let Some((message_index, message)) = messages.first().copied() else {
            continue;
        };
        let mut owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
        if owner.formula_owner.is_some() {
            continue;
        }
        let changed = owner
            .cell_dependencies
            .as_mut()
            .is_some_and(|dependencies| {
                prune_internal_owner_edges(&mut dependencies.cell_record, removed_internal_ids)
            });
        if let Some(dependencies) = &owner.tiled_cell_dependencies {
            tile_ids.extend(
                dependencies
                    .cell_record_tiles
                    .iter()
                    .map(|reference| reference.identifier),
            );
        }
        if changed {
            let data = prune_formula_owner_cell_edges_wire(
                message.data.as_slice(),
                &owner,
                removed_internal_ids,
            )?;
            owner_updates.push((object_id, message_index, message.type_, data));
        }
    }

    package.update_archive(calculation_engine_entry, |archive| {
        for (object_id, message_index, message_type, data) in owner_updates {
            let object = archive.object_mut(object_id).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork formula owner object {object_id} disappeared"
                ))
            })?;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
        }
        for tile_id in tile_ids {
            let object = archive.object_mut(tile_id).ok_or_else(|| {
                Error::InvalidFormat(format!("iWork dependency tile {tile_id} is missing"))
            })?;
            let message_index = object
                .messages
                .iter()
                .position(|message| message.type_ == CELL_RECORD_TILE_MESSAGE_TYPE)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "iWork dependency tile {tile_id} has no tile payload"
                    ))
                })?;
            let mut tile = tsce::CellRecordTileArchive::decode(
                object.messages[message_index].data.as_slice(),
            )?;
            if !prune_internal_owner_edges(&mut tile.cell_records, removed_internal_ids) {
                continue;
            }
            let message_type = object.messages[message_index].type_;
            let data = prune_cell_tile_edges_wire(
                object.messages[message_index].data.as_slice(),
                &tile,
                removed_internal_ids,
            )?;
            object.replace_message(
                message_index,
                RawMessage {
                    type_: message_type,
                    data,
                },
            )?;
        }
        Ok(())
    })
}

pub(super) fn calculation_engine_entry_for_contexts(
    package: &IWorkPackage,
    table_context_ids: &HashSet<u64>,
) -> Result<Option<String>> {
    let Some(entry) = package.calculation_engine_entry_name()? else {
        return Ok(None);
    };
    let archive = package.archive(entry)?;
    let mut owns_table = false;
    for message in archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
    {
        let owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
        owns_table |= owner
            .formula_owner
            .is_some_and(|reference| table_context_ids.contains(&reference.identifier));
    }
    Ok(owns_table.then(|| entry.to_owned()))
}

fn formula_owner_context_family(
    archive: &Archive,
    table_context_ids: &HashSet<u64>,
) -> Result<(Vec<FormulaOwnerSource>, Vec<tsp::Uuid>)> {
    let mut all = Vec::new();
    for object in &archive.objects {
        let messages = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if messages.len() > 1 {
            return Err(Error::InvalidFormat(
                "iWork formula owner object repeats its owner payload".to_owned(),
            ));
        }
        let Some((message_index, message)) = messages.first().copied() else {
            continue;
        };
        all.push(FormulaOwnerSource {
            object: copy_archive_object(object)?,
            message_index,
            owner: tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?,
        });
    }

    let mut selected = all
        .iter()
        .enumerate()
        .filter_map(|(index, source)| {
            source
                .owner
                .formula_owner
                .as_ref()
                .is_some_and(|reference| table_context_ids.contains(&reference.identifier))
                .then_some(index)
        })
        .collect::<HashSet<_>>();
    if selected.is_empty() {
        return Ok((Vec::new(), Vec::new()));
    }
    loop {
        let selected_uuids = selected
            .iter()
            .map(|index| all[*index].owner.formula_owner_uid)
            .collect::<HashSet<_>>();
        let previous_len = selected.len();
        for (index, source) in all.iter().enumerate() {
            if source
                .owner
                .base_owner_uid
                .as_ref()
                .is_some_and(|uuid| selected_uuids.contains(uuid))
            {
                selected.insert(index);
            }
        }
        if selected.len() == previous_len {
            break;
        }
    }
    let owner_uuids = selected
        .iter()
        .map(|index| all[*index].owner.formula_owner_uid)
        .collect::<Vec<_>>();
    let mut owners = all
        .into_iter()
        .enumerate()
        .filter_map(|(index, source)| selected.contains(&index).then_some(source))
        .collect::<Vec<_>>();
    owners.sort_by_key(|source| source.object.archive_info.identifier);
    Ok((owners, owner_uuids))
}

fn reject_incoming_formula_dependencies(
    archive: &Archive,
    table_context_ids: &HashSet<u64>,
    removed_owner_ids: &HashSet<u64>,
    removed_internal_ids: &HashSet<u32>,
    removed_owner_uuids: &[tsp::Uuid],
) -> Result<()> {
    for object in &archive.objects {
        let object_id = object.archive_info.identifier.ok_or_else(|| {
            Error::InvalidFormat("Numbers CalculationEngine object has no identifier".to_owned())
        })?;
        if removed_owner_ids.contains(&object_id) {
            continue;
        }
        for message in &object.messages {
            if message.type_ != FORMULA_OWNER_MESSAGE_TYPE {
                continue;
            }
            let owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
            if let Some(dependency_kind) = formula_owner_removed_dependency_kind(
                archive,
                &owner,
                removed_internal_ids,
                removed_owner_uuids,
            )? {
                let target = owner
                    .formula_owner
                    .as_ref()
                    .map(|reference| reference.identifier);
                return Err(Error::ParseError(format!(
                    "Cannot remove iWork table contexts {table_context_ids:?}: surviving formula owner object {object_id} (target {target:?}, internal {}) has a {dependency_kind} dependency on them [cell={}, tiled_cell={}, range={}, volatile={}, whole_owner={}, spanning_column={}, spanning_row={}, uuid={}]",
                    owner.internal_formula_owner_id,
                    owner.cell_dependencies.is_some(),
                    owner.tiled_cell_dependencies.is_some(),
                    owner.range_dependencies.is_some(),
                    owner.volatile_dependencies.is_some(),
                    owner.whole_owner_dependencies.is_some(),
                    owner.spanning_column_dependencies.is_some(),
                    owner.spanning_row_dependencies.is_some(),
                    owner.uuid_references.is_some(),
                )));
            }
        }
    }
    Ok(())
}

fn formula_owner_removed_dependency_kind(
    archive: &Archive,
    owner: &tsce::FormulaOwnerDependenciesArchive,
    removed_internal_ids: &HashSet<u32>,
    removed_owner_uuids: &[tsp::Uuid],
) -> Result<Option<&'static str>> {
    let records_reference_removed = |records: &[tsce::CellRecordExpandedArchive]| {
        records.iter().any(|record| {
            record.expanded_edges.as_ref().is_some_and(|edges| {
                edges
                    .internal_owner_id_for_edge
                    .iter()
                    .any(|identifier| removed_internal_ids.contains(identifier))
            })
        })
    };
    if owner
        .cell_dependencies
        .as_ref()
        .is_some_and(|dependencies| records_reference_removed(&dependencies.cell_record))
    {
        return Ok(Some("cell"));
    }
    if let Some(dependencies) = &owner.tiled_cell_dependencies {
        for reference in &dependencies.cell_record_tiles {
            let object = archive.object(reference.identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers dependency tile {} is missing",
                    reference.identifier
                ))
            })?;
            let message = object
                .messages
                .iter()
                .find(|message| message.type_ == CELL_RECORD_TILE_MESSAGE_TYPE)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers dependency tile {} has no cell-tile payload",
                        reference.identifier
                    ))
                })?;
            let tile = tsce::CellRecordTileArchive::decode(message.data.as_slice())?;
            if records_reference_removed(&tile.cell_records) {
                return Ok(Some("tiled-cell"));
            }
        }
    }
    if let Some(dependencies) = &owner.tiled_range_dependencies {
        for reference in &dependencies.range_precedents_tile {
            let object = archive.object(reference.identifier).ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers range dependency tile {} is missing",
                    reference.identifier
                ))
            })?;
            let message = object
                .messages
                .iter()
                .find(|message| message.type_ == RANGE_PRECEDENTS_TILE_MESSAGE_TYPE)
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers range dependency tile {} has no range-tile payload",
                        reference.identifier
                    ))
                })?;
            let tile = tsce::RangePrecedentsTileArchive::decode(message.data.as_slice())?;
            if removed_internal_ids.contains(&tile.to_owner_id) && !tile.from_to_range.is_empty() {
                return Ok(Some("tiled-range"));
            }
        }
    }
    if owner
        .range_dependencies
        .as_ref()
        .is_some_and(|dependencies| {
            dependencies.back_dependency.iter().any(|dependency| {
                dependency
                    .internal_range_reference
                    .as_ref()
                    .is_some_and(|reference| removed_internal_ids.contains(&reference.owner_id))
            })
        })
        || owner
            .volatile_dependencies
            .as_ref()
            .and_then(|dependencies| dependencies.volatile_geometry_cell_refs.as_ref())
            .is_some_and(|references| {
                references
                    .owner_entries
                    .iter()
                    .any(|entry| removed_internal_ids.contains(&entry.owner_id))
            })
        || owner
            .whole_owner_dependencies
            .as_ref()
            .and_then(|dependencies| dependencies.dependent_cells.as_ref())
            .is_some_and(|references| {
                references
                    .owner_entries
                    .iter()
                    .any(|entry| removed_internal_ids.contains(&entry.owner_id))
            })
    {
        return Ok(Some("range, volatile, or whole-owner"));
    }
    for dependencies in [
        owner.spanning_column_dependencies.as_ref(),
        owner.spanning_row_dependencies.as_ref(),
    ]
    .into_iter()
    .flatten()
    {
        if dependencies.coord_refers_to_spans.iter().any(|reference| {
            reference
                .ranges_by_table_context
                .iter()
                .any(|range| removed_internal_ids.contains(&range.owner_id))
        }) {
            return Ok(Some("spanning-row or spanning-column"));
        }
    }
    Ok(owner
        .uuid_references
        .as_ref()
        .is_some_and(|references| {
            references
                .table_refs
                .iter()
                .any(|reference| removed_owner_uuids.contains(&reference.owner_uuid))
                || references
                    .table_uuid_refs
                    .iter()
                    .any(|reference| removed_owner_uuids.contains(&reference.owner_uuid))
        })
        .then_some("UUID"))
}
