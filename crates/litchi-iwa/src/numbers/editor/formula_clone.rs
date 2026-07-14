//! CalculationEngine graph cloning for independently duplicated Numbers tables.

mod dependency_wire;
mod formula_storage;

use super::*;
use crate::wire::{
    append_repeated_length_delimited_field, patch_nested_varint_field, patch_varint_field,
    transform_length_delimited_field, transform_length_delimited_fields_at_path,
};
use dependency_wire::{
    append_formula_owners_to_engine, remap_cell_records, remap_cell_tile_wire, remap_formula_owner,
    remap_formula_owner_wire,
};
pub(super) use formula_storage::{
    remap_cloned_formula_owner_storage, remap_cloned_formula_storage,
};

const CALCULATION_ENGINE_ENTRY: &str = "Index/CalculationEngine.iwa";
const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
const CELL_RECORD_TILE_MESSAGE_TYPE: u32 = 4_009;

struct FormulaOwnerSource {
    object: ArchiveObject,
    message_index: usize,
    owner: tsce::FormulaOwnerDependenciesArchive,
}

struct CellTileSource {
    object: ArchiveObject,
    message_index: usize,
    tile: tsce::CellRecordTileArchive,
}

/// Clone the CalculationEngine owner family associated with one table.
///
/// This operation preserves unknown wire fields and rejects dependency forms
/// that cannot yet be remapped without ambiguity. Packages without a
/// CalculationEngine are valid minimal documents and need no graph clone.
pub(super) fn clone_table_formula_graph(
    package: &mut IWorkPackage,
    source_table_info_id: u64,
    new_table_info_id: u64,
    source_table_uuid: &str,
    new_table_uuid: &str,
) -> Result<Vec<u64>> {
    if !package.contains_entry(CALCULATION_ENGINE_ENTRY) {
        return Ok(Vec::new());
    }
    let source_uuid = parse_table_uuid(source_table_uuid)?;
    let new_uuid = parse_table_uuid(new_table_uuid)?;
    let archive = package.archive(CALCULATION_ENGINE_ENTRY)?;
    let (owners, source_owner_uuid) = formula_owner_family(&archive, source_table_info_id)?;
    if owners.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "Numbers table info {source_table_info_id} has no CalculationEngine owner family"
        )));
    }
    for source in &owners {
        require_cloneable_formula_owner(&source.owner)?;
    }
    let tiles = cell_record_tiles(&archive, &owners)?;

    let mut next_identifier = next_object_identifier(package)?;
    let mut object_remap = HashMap::with_capacity(owners.len() + tiles.len() + 1);
    object_remap.insert(source_table_info_id, new_table_info_id);
    for source in &owners {
        let identifier = source.object.archive_info.identifier.ok_or_else(|| {
            Error::InvalidFormat("Numbers formula owner has no object identifier".to_owned())
        })?;
        object_remap.insert(identifier, take_identifier(&mut next_identifier)?);
    }
    for source in &tiles {
        let identifier = source.object.archive_info.identifier.ok_or_else(|| {
            Error::InvalidFormat("Numbers dependency tile has no object identifier".to_owned())
        })?;
        object_remap.insert(identifier, take_identifier(&mut next_identifier)?);
    }

    let mut next_internal_owner_id = archive
        .objects
        .iter()
        .flat_map(|object| &object.messages)
        .filter(|message| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
        .filter_map(|message| {
            tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())
                .ok()
                .map(|owner| owner.internal_formula_owner_id)
        })
        .max()
        .unwrap_or(0)
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("Numbers formula owner ID overflow".to_owned()))?;
    let mut internal_remap = HashMap::with_capacity(owners.len());
    for source in &owners {
        let replacement = next_internal_owner_id;
        next_internal_owner_id = next_internal_owner_id
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("Numbers formula owner ID overflow".to_owned()))?;
        internal_remap.insert(source.owner.internal_formula_owner_id, replacement);
    }

    if source_owner_uuid != formula_owner_uuid_for_table(&source_uuid) {
        return Err(Error::InvalidFormat(format!(
            "Numbers table info {source_table_info_id} has an inconsistent formula-owner UUID"
        )));
    }
    let new_owner_uuid = formula_owner_uuid_for_table(&new_uuid);
    let mut uuid_remap = HashMap::with_capacity(owners.len() + 2);
    uuid_remap.insert(uuid_key(&source_uuid), new_uuid);
    uuid_remap.insert(uuid_key(&source_owner_uuid), new_owner_uuid);
    for source in &owners {
        let replacement = translate_owner_uuid(
            &source.owner.formula_owner_uid,
            &source_owner_uuid,
            &new_owner_uuid,
        );
        let key = uuid_key(&source.owner.formula_owner_uid);
        if let Some(existing) = uuid_remap.get(&key) {
            if existing != &replacement {
                return Err(Error::InvalidFormat(
                    "Numbers formula owner UUID remap is inconsistent".to_owned(),
                ));
            }
        } else {
            uuid_remap.insert(key, replacement);
        }
    }

    let mut cloned_objects = Vec::with_capacity(owners.len() + tiles.len());
    let mut cloned_owner_ids = Vec::with_capacity(owners.len());
    let mut owner_map_entries = Vec::with_capacity(owners.len());
    let mut formula_count = 0u64;
    for source in &owners {
        let source_id = source.object.archive_info.identifier.ok_or_else(|| {
            Error::InvalidFormat("Numbers formula owner has no object identifier".to_owned())
        })?;
        let new_id = object_remap[&source_id];
        let new_internal_id = internal_remap[&source.owner.internal_formula_owner_id];
        let mut expected = source.owner.clone();
        remap_formula_owner(
            &mut expected,
            new_table_info_id,
            &object_remap,
            &internal_remap,
            &uuid_remap,
        );
        formula_count = formula_count
            .checked_add(formula_cell_count(&source.owner, &tiles)?)
            .ok_or_else(|| Error::ParseError("Numbers formula count overflow".to_owned()))?;
        let original = &source.object.messages[source.message_index];
        let data = remap_formula_owner_wire(
            original.data.as_slice(),
            &source.owner,
            &expected,
            &object_remap,
            &internal_remap,
            &uuid_remap,
        )?;
        let mut messages = source.object.messages.clone();
        messages[source.message_index] = RawMessage {
            type_: original.type_,
            data,
        };
        cloned_objects.push(clone_numbers_object_metadata(
            &source.object,
            new_id,
            messages,
            &object_remap,
        )?);
        cloned_owner_ids.push(new_id);
        owner_map_entries.push(tsce::owner_id_map_archive::OwnerIdMapArchiveEntry {
            internal_owner_id: new_internal_id,
            owner_id: uuid_as_cfuuid(&expected.formula_owner_uid),
        });
    }
    for source in &tiles {
        let source_id = source.object.archive_info.identifier.ok_or_else(|| {
            Error::InvalidFormat("Numbers dependency tile has no object identifier".to_owned())
        })?;
        let original = &source.object.messages[source.message_index];
        let mut expected = source.tile.clone();
        expected.internal_owner_id = internal_remap
            .get(&source.tile.internal_owner_id)
            .copied()
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "Numbers dependency tile {source_id} belongs to an owner outside its table family"
                ))
            })?;
        remap_cell_records(&mut expected.cell_records, &internal_remap);
        let data = remap_cell_tile_wire(
            original.data.as_slice(),
            &source.tile,
            &expected,
            &internal_remap,
        )?;
        let mut messages = source.object.messages.clone();
        messages[source.message_index] = RawMessage {
            type_: original.type_,
            data,
        };
        cloned_objects.push(clone_numbers_object_metadata(
            &source.object,
            object_remap[&source_id],
            messages,
            &object_remap,
        )?);
    }

    let source_uuid_ids = component_identifier_for_entry(package, CALCULATION_ENGINE_ENTRY)?
        .map(|component| component_uuid_identifiers(package, component))
        .transpose()?
        .flatten()
        .unwrap_or_default();
    let new_uuid_ids = object_remap
        .iter()
        .filter_map(|(source, replacement)| {
            (*source != source_table_info_id && source_uuid_ids.contains(source))
                .then_some(*replacement)
        })
        .collect::<Vec<_>>();

    package.update_archive(CALCULATION_ENGINE_ENTRY, |archive| {
        let (engine_id, engine_message_index) = calculation_engine_location(archive)?;
        let engine_object = archive.object_mut(engine_id).ok_or_else(|| {
            Error::InvalidFormat("Numbers CalculationEngine root is missing".to_owned())
        })?;
        let original = engine_object.messages[engine_message_index].clone();
        let data = append_formula_owners_to_engine(
            original.data.as_slice(),
            &cloned_owner_ids,
            &owner_map_entries,
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
        let previous_owner_set = info
            .object_references
            .iter()
            .copied()
            .collect::<HashSet<_>>();
        info.object_references
            .extend(cloned_owner_ids.iter().copied());
        for field in &mut info.field_infos {
            if field
                .object_references
                .iter()
                .any(|identifier| previous_owner_set.contains(identifier))
            {
                field
                    .object_references
                    .extend(cloned_owner_ids.iter().copied());
            }
        }
        for object in cloned_objects {
            archive.insert_object(object)?;
        }
        Ok(())
    })?;

    let last_identifier = next_identifier.checked_sub(1).ok_or_else(|| {
        Error::InvalidFormat("Numbers formula clone allocated no identifiers".to_owned())
    })?;
    set_package_last_object_identifier(package, last_identifier)?;
    if let Some(component) = component_identifier_for_entry(package, CALCULATION_ENGINE_ENTRY)? {
        add_component_object_uuids(package, component, &new_uuid_ids)?;
    }
    Ok(cloned_owner_ids)
}

pub(super) fn formula_graph_owner_uuids(
    package: &IWorkPackage,
    table_info_id: u64,
    source_table_uuid: &str,
    new_table_uuid: &str,
) -> Result<Option<(tsp::Uuid, tsp::Uuid)>> {
    if !package.contains_entry(CALCULATION_ENGINE_ENTRY) {
        return Ok(None);
    }
    let source_table_uuid = parse_table_uuid(source_table_uuid)?;
    let new_table_uuid = parse_table_uuid(new_table_uuid)?;
    let archive = package.archive(CALCULATION_ENGINE_ENTRY)?;
    let (_, source_owner_uuid) = formula_owner_family(&archive, table_info_id)?;
    if source_owner_uuid != formula_owner_uuid_for_table(&source_table_uuid) {
        return Err(Error::InvalidFormat(format!(
            "Numbers table info {table_info_id} has an inconsistent formula-owner UUID"
        )));
    }
    let new_owner_uuid = formula_owner_uuid_for_table(&new_table_uuid);
    Ok(Some((source_owner_uuid, new_owner_uuid)))
}

/// Return whether every explicit dependency edge for a table stays inside its owner family.
///
/// A standalone table duplicate intentionally keeps cross-table references pointed at the
/// original targets. Sheet duplication needs a later sheet-wide remap for those edges, so it
/// rejects them until that operation can preserve their native semantics exactly.
pub(super) fn table_formula_graph_is_self_contained(
    package: &IWorkPackage,
    table_info_id: u64,
) -> Result<bool> {
    if !package.contains_entry(CALCULATION_ENGINE_ENTRY) {
        return Ok(true);
    }
    let archive = package.archive(CALCULATION_ENGINE_ENTRY)?;
    let (owners, _) = formula_owner_family(&archive, table_info_id)?;
    let family = owners
        .iter()
        .map(|source| source.owner.internal_formula_owner_id)
        .collect::<HashSet<_>>();
    let inline_is_self_contained = owners.iter().all(|source| {
        source
            .owner
            .cell_dependencies
            .as_ref()
            .is_none_or(|dependencies| {
                dependency_records_are_self_contained(&dependencies.cell_record, &family)
            })
    });
    if !inline_is_self_contained {
        return Ok(false);
    }
    Ok(cell_record_tiles(&archive, &owners)?
        .iter()
        .all(|source| dependency_records_are_self_contained(&source.tile.cell_records, &family)))
}

fn dependency_records_are_self_contained(
    records: &[tsce::CellRecordExpandedArchive],
    family: &HashSet<u32>,
) -> bool {
    records.iter().all(|record| {
        record.expanded_edges.as_ref().is_none_or(|edges| {
            edges
                .internal_owner_id_for_edge
                .iter()
                .all(|identifier| family.contains(identifier))
        })
    })
}

/// Remap formula host and explicit self-table references in cloned formula lists.
fn formula_owner_family(
    archive: &Archive,
    source_table_info_id: u64,
) -> Result<(Vec<FormulaOwnerSource>, tsp::Uuid)> {
    let mut all = Vec::new();
    let mut direct_owner_count = 0usize;
    let mut direct_owner_uuid = None;
    for object in &archive.objects {
        let matches = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == FORMULA_OWNER_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if matches.len() > 1 {
            return Err(Error::InvalidFormat(
                "Numbers formula owner object repeats its owner payload".to_owned(),
            ));
        }
        let Some((message_index, message)) = matches.first().copied() else {
            continue;
        };
        let owner = tsce::FormulaOwnerDependenciesArchive::decode(message.data.as_slice())?;
        let direct = owner
            .formula_owner
            .as_ref()
            .map(|reference| reference.identifier)
            == Some(source_table_info_id);
        if direct {
            direct_owner_uuid = Some(owner.formula_owner_uid);
            direct_owner_count += 1;
        }
        all.push(FormulaOwnerSource {
            object: copy_archive_object(object)?,
            message_index,
            owner,
        });
    }
    if direct_owner_count != 1 {
        return Err(Error::InvalidFormat(format!(
            "Numbers table info {source_table_info_id} must have exactly one direct formula owner"
        )));
    }
    let direct_owner_uuid = direct_owner_uuid.ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table info {source_table_info_id} has no direct formula owner"
        ))
    })?;
    let mut result = all
        .into_iter()
        .filter(|source| {
            source
                .owner
                .formula_owner
                .as_ref()
                .map(|reference| reference.identifier)
                == Some(source_table_info_id)
                || source.owner.base_owner_uid.as_ref() == Some(&direct_owner_uuid)
        })
        .collect::<Vec<_>>();
    result.sort_by_key(|source| source.object.archive_info.identifier);
    Ok((result, direct_owner_uuid))
}

fn cell_record_tiles(
    archive: &Archive,
    owners: &[FormulaOwnerSource],
) -> Result<Vec<CellTileSource>> {
    let mut identifiers = BTreeMap::new();
    for source in owners {
        if let Some(dependencies) = &source.owner.tiled_cell_dependencies {
            for reference in &dependencies.cell_record_tiles {
                if identifiers.insert(reference.identifier, ()).is_some() {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers formula owner family shares dependency tile {}",
                        reference.identifier
                    )));
                }
            }
        }
    }
    let mut result = Vec::with_capacity(identifiers.len());
    for identifier in identifiers.keys().copied() {
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers dependency tile {identifier} is missing"))
        })?;
        let matches = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == CELL_RECORD_TILE_MESSAGE_TYPE)
            .collect::<Vec<_>>();
        if matches.len() != 1 {
            return Err(Error::InvalidFormat(format!(
                "Numbers dependency tile {identifier} must have exactly one cell-tile payload"
            )));
        }
        let (message_index, message) = matches[0];
        result.push(CellTileSource {
            object: copy_archive_object(object)?,
            message_index,
            tile: tsce::CellRecordTileArchive::decode(message.data.as_slice())?,
        });
    }
    Ok(result)
}

fn require_cloneable_formula_owner(owner: &tsce::FormulaOwnerDependenciesArchive) -> Result<()> {
    let unsupported = owner
        .range_dependencies
        .as_ref()
        .is_some_and(|value| !value.back_dependency.is_empty())
        || owner
            .volatile_dependencies
            .as_ref()
            .and_then(|value| value.volatile_geometry_cell_refs.as_ref())
            .is_some_and(|value| !value.owner_entries.is_empty())
        || owner
            .spanning_column_dependencies
            .as_ref()
            .is_some_and(|value| !value.coord_refers_to_spans.is_empty())
        || owner
            .spanning_row_dependencies
            .as_ref()
            .is_some_and(|value| !value.coord_refers_to_spans.is_empty())
        || owner
            .whole_owner_dependencies
            .as_ref()
            .and_then(|value| value.dependent_cells.as_ref())
            .is_some_and(|value| !value.owner_entries.is_empty())
        || owner
            .cell_errors
            .as_ref()
            .is_some_and(|value| !value.errors.is_empty() || !value.enhanced_errors.is_empty())
        || owner
            .uuid_references
            .as_ref()
            .is_some_and(|value| !value.table_refs.is_empty() || !value.table_uuid_refs.is_empty())
        || owner
            .tiled_range_dependencies
            .as_ref()
            .is_some_and(|value| !value.range_precedents_tile.is_empty());
    if unsupported {
        return Err(Error::ParseError(
            "Cannot duplicate a Numbers table with advanced range, volatile, spill, or UUID dependency state"
                .to_owned(),
        ));
    }
    Ok(())
}

fn calculation_engine_location(archive: &Archive) -> Result<(u64, usize)> {
    let matches = archive
        .objects
        .iter()
        .flat_map(|object| {
            object
                .messages
                .iter()
                .enumerate()
                .filter(|(_, message)| message.type_ == CALCULATION_ENGINE_MESSAGE_TYPE)
                .filter_map(move |(index, _)| object.archive_info.identifier.map(|id| (id, index)))
        })
        .collect::<Vec<_>>();
    match matches.as_slice() {
        [location] => Ok(*location),
        _ => Err(Error::InvalidFormat(
            "Numbers CalculationEngine must contain exactly one root payload".to_owned(),
        )),
    }
}

fn formula_cell_count(
    owner: &tsce::FormulaOwnerDependenciesArchive,
    tiles: &[CellTileSource],
) -> Result<u64> {
    let mut coordinates = BTreeMap::new();
    if let Some(dependencies) = &owner.cell_dependencies {
        for record in &dependencies.cell_record {
            coordinates.insert((record.row, record.column), ());
        }
    }
    if let Some(dependencies) = &owner.tiled_cell_dependencies {
        for reference in &dependencies.cell_record_tiles {
            let tile = tiles
                .iter()
                .find(|tile| tile.object.archive_info.identifier == Some(reference.identifier))
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "Numbers dependency tile {} is missing",
                        reference.identifier
                    ))
                })?;
            for record in &tile.tile.cell_records {
                coordinates.insert((record.row, record.column), ());
            }
        }
    }
    u64::try_from(coordinates.len())
        .map_err(|_| Error::ParseError("Numbers formula count exceeds u64".to_owned()))
}

fn remap_uuid_at_path(
    data: &[u8],
    path: &[u32],
    remap: &HashMap<(u64, u64), tsp::Uuid>,
) -> Result<Vec<u8>> {
    transform_length_delimited_fields_at_path(data, path, |uuid_data| {
        let uuid = tsp::Uuid::decode(uuid_data)?;
        let Some(replacement) = remap.get(&uuid_key(&uuid)) else {
            return Ok(uuid_data.to_vec());
        };
        patch_uuid_wire(uuid_data, replacement)
    })
}

fn patch_uuid_wire(data: &[u8], uuid: &tsp::Uuid) -> Result<Vec<u8>> {
    let data = patch_varint_field(data, 1, true, Some(uuid.lower))?;
    patch_varint_field(&data, 2, true, Some(uuid.upper))
}

fn parse_table_uuid(value: &str) -> Result<tsp::Uuid> {
    let compact = value
        .chars()
        .filter(|character| *character != '-')
        .collect::<String>();
    if compact.len() != 32 || !compact.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(Error::InvalidFormat(format!(
            "Numbers table UUID {value:?} is malformed"
        )));
    }
    let raw = u128::from_str_radix(&compact, 16)
        .map_err(|_| Error::InvalidFormat(format!("Numbers table UUID {value:?} is malformed")))?;
    Ok(tsp::Uuid {
        lower: raw as u64,
        upper: (raw >> 64) as u64,
    })
}

fn translate_owner_uuid(
    owner: &tsp::Uuid,
    source: &tsp::Uuid,
    replacement: &tsp::Uuid,
) -> tsp::Uuid {
    let owner = (u128::from(owner.upper) << 64) | u128::from(owner.lower);
    let source = (u128::from(source.upper) << 64) | u128::from(source.lower);
    let replacement = (u128::from(replacement.upper) << 64) | u128::from(replacement.lower);
    let translated = replacement.wrapping_add(owner.wrapping_sub(source));
    tsp::Uuid {
        lower: translated as u64,
        upper: (translated >> 64) as u64,
    }
}

fn formula_owner_uuid_for_table(table: &tsp::Uuid) -> tsp::Uuid {
    tsp::Uuid {
        lower: table.upper.swap_bytes(),
        upper: table.lower.swap_bytes(),
    }
}

fn uuid_key(uuid: &tsp::Uuid) -> (u64, u64) {
    (uuid.lower, uuid.upper)
}

fn uuid_as_cfuuid(uuid: &tsp::Uuid) -> tsp::CfuuidArchive {
    tsp::CfuuidArchive {
        uuid_bytes: None,
        uuid_w0: Some(uuid.lower as u32),
        uuid_w1: Some((uuid.lower >> 32) as u32),
        uuid_w2: Some(uuid.upper as u32),
        uuid_w3: Some((uuid.upper >> 32) as u32),
    }
}

fn copy_archive_object(source: &ArchiveObject) -> Result<ArchiveObject> {
    let identifier = source.archive_info.identifier.ok_or_else(|| {
        Error::InvalidFormat("Numbers archive object has no identifier".to_owned())
    })?;
    let mut copied = ArchiveObject::new(identifier, source.messages.clone())?;
    copied.archive_info = source.archive_info.clone();
    copied.header_offset = source.header_offset;
    copied.header_length = source.header_length;
    copied.data_offset = source.data_offset;
    copied.data_length = source.data_length;
    Ok(copied)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn native_formula_owner_uuid_reverses_table_uuid_bytes() {
        let table = parse_table_uuid("68B384C7-9E7F-4B90-A1F5-067B81524A6A").unwrap();
        assert_eq!(
            formula_owner_uuid_for_table(&table),
            tsp::Uuid {
                lower: 0x904b_7f9e_c784_b368,
                upper: 0x6a4a_5281_7b06_f5a1,
            }
        );
    }
}
