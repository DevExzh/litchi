//! Transactional removal of Numbers CalculationEngine table-owner families.

use super::dependency_wire::remove_formula_owners_from_engine;
use super::*;

/// Remove one table's CalculationEngine owner family.
///
/// Incoming cross-table dependencies are rejected so deletion cannot leave a
/// surviving formula with a dangling internal owner ID.
pub(in crate::numbers::editor) fn remove_table_formula_graph(
    package: &mut IWorkPackage,
    table_info_id: u64,
) -> Result<Vec<u64>> {
    if !package.contains_entry(CALCULATION_ENGINE_ENTRY) {
        return Ok(Vec::new());
    }
    let archive = package.archive(CALCULATION_ENGINE_ENTRY)?;
    let direct_owner_count = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
        .map(|message| tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice()))
        .collect::<std::result::Result<Vec<_>, _>>()?
        .into_iter()
        .filter(|owner| {
            owner
                .formula_owner
                .as_ref()
                .map(|reference| reference.identifier)
                == Some(table_info_id)
        })
        .count();
    if direct_owner_count == 0 {
        return Ok(Vec::new());
    }
    let (owners, direct_owner_uuid) = formula_owner_family(&archive, table_info_id)?;
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
    reject_incoming_formula_dependencies(
        &archive,
        table_info_id,
        &owner_ids,
        &internal_owner_ids,
        &direct_owner_uuid,
    )?;
    let tiles = cell_record_tiles(&archive, &owners)?;
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

    package.update_archive(CALCULATION_ENGINE_ENTRY, |archive| {
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
        for identifier in owner_ids.iter().copied().chain(tile_ids.iter().copied()) {
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
    if let Some(component) = component_identifier_for_entry(package, CALCULATION_ENGINE_ENTRY)? {
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

fn reject_incoming_formula_dependencies(
    archive: &Archive,
    table_info_id: u64,
    removed_owner_ids: &HashSet<u64>,
    removed_internal_ids: &HashSet<u32>,
    removed_owner_uuid: &tsp::Uuid,
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
            if formula_owner_references_removed_family(
                archive,
                &owner,
                removed_internal_ids,
                removed_owner_uuid,
            )? {
                return Err(Error::ParseError(format!(
                    "Cannot remove Numbers table info {table_info_id} while another formula owner depends on it"
                )));
            }
        }
    }
    Ok(())
}

fn formula_owner_references_removed_family(
    archive: &Archive,
    owner: &tsce::FormulaOwnerDependenciesArchive,
    removed_internal_ids: &HashSet<u32>,
    removed_owner_uuid: &tsp::Uuid,
) -> Result<bool> {
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
        return Ok(true);
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
                return Ok(true);
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
        return Ok(true);
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
            return Ok(true);
        }
    }
    Ok(owner.uuid_references.as_ref().is_some_and(|references| {
        references
            .table_refs
            .iter()
            .any(|reference| &reference.owner_uuid == removed_owner_uuid)
            || references
                .table_uuid_refs
                .iter()
                .any(|reference| &reference.owner_uuid == removed_owner_uuid)
    }))
}
