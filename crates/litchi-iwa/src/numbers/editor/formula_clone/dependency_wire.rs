//! Wire-preserving CalculationEngine owner and dependency-tile rewrites.

use super::*;

pub(in crate::numbers::editor) fn append_formula_owners_to_engine(
    original: &[u8],
    owner_ids: &[u64],
    owner_map_entries: &[tsce::owner_id_map_archive::OwnerIdMapArchiveEntry],
    formula_count: u64,
) -> Result<Vec<u8>> {
    let previous = tsce::CalculationEngineArchive::decode(original)?;
    let mut expected = previous.clone();
    expected
        .dependency_tracker
        .formula_owner_dependencies
        .extend(owner_ids.iter().map(|identifier| tsp::Reference {
            identifier: *identifier,
            ..Default::default()
        }));
    expected
        .dependency_tracker
        .owner_id_map
        .get_or_insert_default()
        .map_entry
        .extend_from_slice(owner_map_entries);
    let previous_count = expected
        .dependency_tracker
        .number_of_formulas
        .unwrap_or_default();
    expected.dependency_tracker.number_of_formulas = Some(
        previous_count
            .checked_add(formula_count)
            .ok_or_else(|| Error::ParseError("Numbers formula count overflow".to_owned()))?,
    );
    let data = transform_length_delimited_field(original, 2, |tracker_data| {
        let tracker = tsce::DependencyTrackerArchive::decode(tracker_data)?;
        let mut data = tracker_data.to_vec();
        for identifier in owner_ids {
            data = append_repeated_length_delimited_field(
                &data,
                6,
                &tsp::Reference {
                    identifier: *identifier,
                    ..Default::default()
                }
                .encode_to_vec(),
            )?;
        }
        let owner_map = expected
            .dependency_tracker
            .owner_id_map
            .as_ref()
            .ok_or_else(|| Error::InvalidFormat("Numbers owner map disappeared".to_owned()))?;
        data = if tracker.owner_id_map.is_some() {
            transform_length_delimited_field(&data, 3, |map_data| {
                let mut map_data = map_data.to_vec();
                for entry in owner_map_entries {
                    map_data = append_repeated_length_delimited_field(
                        &map_data,
                        1,
                        &entry.encode_to_vec(),
                    )?;
                }
                Ok(map_data)
            })?
        } else {
            crate::wire::patch_length_delimited_field(
                &data,
                3,
                false,
                Some(&owner_map.encode_to_vec()),
            )?
        };
        patch_varint_field(
            &data,
            5,
            tracker.number_of_formulas.is_some(),
            expected.dependency_tracker.number_of_formulas,
        )
    })?;
    if tsce::CalculationEngineArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers CalculationEngine formula-owner clone failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(in crate::numbers::editor) fn reorder_formula_owners_in_engine(
    original: &[u8],
    owner_ids: &[u64],
) -> Result<Vec<u8>> {
    let previous = tsce::CalculationEngineArchive::decode(original)?;
    let mut expected = previous.clone();
    expected.dependency_tracker.formula_owner_dependencies = owner_ids
        .iter()
        .map(|identifier| tsp::Reference {
            identifier: *identifier,
            ..Default::default()
        })
        .collect();
    let data = transform_length_delimited_field(original, 2, |tracker_data| {
        let references = owner_ids
            .iter()
            .map(|identifier| {
                tsp::Reference {
                    identifier: *identifier,
                    ..Default::default()
                }
                .encode_to_vec()
            })
            .collect::<Vec<_>>();
        crate::wire::rewrite_repeated_length_delimited_fields(tracker_data, 6, &references)
    })?;
    if tsce::CalculationEngineArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers CalculationEngine formula-owner reorder failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(in crate::numbers::editor) fn decrement_formula_count_in_engine(
    original: &[u8],
) -> Result<Vec<u8>> {
    let previous = tsce::CalculationEngineArchive::decode(original)?;
    let mut expected = previous.clone();
    let previous_count = expected
        .dependency_tracker
        .number_of_formulas
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers CalculationEngine has no formula count".to_owned())
        })?;
    expected.dependency_tracker.number_of_formulas =
        Some(previous_count.checked_sub(1).ok_or_else(|| {
            Error::InvalidFormat(
                "Numbers formula count cannot be decremented below zero".to_owned(),
            )
        })?);
    let data = transform_length_delimited_field(original, 2, |tracker_data| {
        let tracker = tsce::DependencyTrackerArchive::decode(tracker_data)?;
        patch_varint_field(
            tracker_data,
            5,
            tracker.number_of_formulas.is_some(),
            expected.dependency_tracker.number_of_formulas,
        )
    })?;
    if tsce::CalculationEngineArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers CalculationEngine formula-count decrement failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remove_formula_owners_from_engine(
    original: &[u8],
    owner_ids: &HashSet<u64>,
    internal_owner_ids: &HashSet<u32>,
    formula_count: u64,
) -> Result<Vec<u8>> {
    let previous = tsce::CalculationEngineArchive::decode(original)?;
    let mut expected = previous.clone();
    expected
        .dependency_tracker
        .formula_owner_dependencies
        .retain(|reference| !owner_ids.contains(&reference.identifier));
    if let Some(owner_map) = &mut expected.dependency_tracker.owner_id_map {
        owner_map
            .map_entry
            .retain(|entry| !internal_owner_ids.contains(&entry.internal_owner_id));
    }
    let previous_count = previous
        .dependency_tracker
        .number_of_formulas
        .unwrap_or_default();
    expected.dependency_tracker.number_of_formulas = Some(
        previous_count
            .checked_sub(formula_count)
            .ok_or_else(|| Error::InvalidFormat("Numbers formula count underflow".to_owned()))?,
    );

    let data = transform_length_delimited_field(original, 2, |tracker_data| {
        let tracker = tsce::DependencyTrackerArchive::decode(tracker_data)?;
        let references = crate::wire::repeated_length_delimited_payloads(tracker_data, 6)?;
        let retained = references
            .iter()
            .filter_map(|payload| {
                tsp::Reference::decode(*payload)
                    .map(|reference| {
                        (!owner_ids.contains(&reference.identifier)).then(|| payload.to_vec())
                    })
                    .transpose()
            })
            .collect::<std::result::Result<Vec<_>, _>>()?;
        let mut data =
            crate::wire::rewrite_repeated_length_delimited_fields(tracker_data, 6, &retained)?;
        if tracker.owner_id_map.is_some() {
            data = transform_length_delimited_field(&data, 3, |map_data| {
                let entries = crate::wire::repeated_length_delimited_payloads(map_data, 1)?;
                let retained = entries
                    .iter()
                    .filter_map(|payload| {
                        tsce::owner_id_map_archive::OwnerIdMapArchiveEntry::decode(*payload)
                            .map(|entry| {
                                (!internal_owner_ids.contains(&entry.internal_owner_id))
                                    .then(|| payload.to_vec())
                            })
                            .transpose()
                    })
                    .collect::<std::result::Result<Vec<_>, _>>()?;
                crate::wire::rewrite_repeated_length_delimited_fields(map_data, 1, &retained)
            })?;
        }
        patch_varint_field(
            &data,
            5,
            tracker.number_of_formulas.is_some(),
            expected.dependency_tracker.number_of_formulas,
        )
    })?;
    if tsce::CalculationEngineArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers CalculationEngine formula-owner removal failed validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_formula_owner(
    owner: &mut tsce::FormulaOwnerDependenciesArchive,
    new_table_info_id: u64,
    object_remap: &HashMap<u64, u64>,
    internal_remap: &HashMap<u32, u32>,
    uuid_remap: &HashMap<(u64, u64), tsp::Uuid>,
) {
    if let Some(replacement) = uuid_remap.get(&uuid_key(&owner.formula_owner_uid)) {
        owner.formula_owner_uid = *replacement;
    }
    if let Some(replacement) = internal_remap.get(&owner.internal_formula_owner_id) {
        owner.internal_formula_owner_id = *replacement;
    }
    if let Some(reference) = &mut owner.formula_owner {
        reference.identifier = new_table_info_id;
    }
    if let Some(base) = &mut owner.base_owner_uid
        && let Some(replacement) = uuid_remap.get(&uuid_key(base))
    {
        *base = *replacement;
    }
    if let Some(dependencies) = &mut owner.cell_dependencies {
        remap_cell_records(&mut dependencies.cell_record, internal_remap);
    }
    if let Some(dependencies) = &mut owner.range_dependencies {
        for dependency in &mut dependencies.back_dependency {
            if let Some(reference) = &mut dependency.internal_range_reference
                && let Some(replacement) = internal_remap.get(&reference.owner_id)
            {
                reference.owner_id = *replacement;
            }
            if let Some(reference) = &mut dependency.range_reference {
                remap_cfuuid(&mut reference.table_id, uuid_remap);
            }
        }
    }
    if let Some(dependencies) = &mut owner.tiled_cell_dependencies {
        for reference in &mut dependencies.cell_record_tiles {
            if let Some(replacement) = object_remap.get(&reference.identifier) {
                reference.identifier = *replacement;
            }
        }
    }
    if let Some(references) = &mut owner.uuid_references {
        for reference in &mut references.table_refs {
            remap_uuid(&mut reference.owner_uuid, uuid_remap);
        }
        for table in &mut references.table_uuid_refs {
            remap_uuid(&mut table.owner_uuid, uuid_remap);
            for reference in &mut table.uuid_refs {
                remap_uuid(&mut reference.uuid, uuid_remap);
            }
        }
    }
    if let Some(dependencies) = &mut owner.tiled_range_dependencies {
        for reference in &mut dependencies.range_precedents_tile {
            if let Some(replacement) = object_remap.get(&reference.identifier) {
                reference.identifier = *replacement;
            }
        }
    }
}

fn remap_uuid(uuid: &mut tsp::Uuid, remap: &HashMap<(u64, u64), tsp::Uuid>) {
    if let Some(replacement) = remap.get(&uuid_key(uuid)) {
        *uuid = *replacement;
    }
}

fn remap_cfuuid(uuid: &mut tsp::CfuuidArchive, remap: &HashMap<(u64, u64), tsp::Uuid>) {
    let Some(key) = cfuuid_key(uuid) else {
        return;
    };
    let Some(replacement) = remap.get(&key) else {
        return;
    };
    if uuid.uuid_bytes.is_some() {
        uuid.uuid_bytes = Some(uuid_bytes(replacement).to_vec());
    }
    if uuid.uuid_w0.is_some() {
        uuid.uuid_w0 = Some(replacement.lower as u32);
    }
    if uuid.uuid_w1.is_some() {
        uuid.uuid_w1 = Some((replacement.lower >> 32) as u32);
    }
    if uuid.uuid_w2.is_some() {
        uuid.uuid_w2 = Some(replacement.upper as u32);
    }
    if uuid.uuid_w3.is_some() {
        uuid.uuid_w3 = Some((replacement.upper >> 32) as u32);
    }
}

pub(super) fn remap_cell_records(
    records: &mut [tsce::CellRecordExpandedArchive],
    internal_remap: &HashMap<u32, u32>,
) {
    for record in records {
        if let Some(edges) = &mut record.expanded_edges {
            for identifier in &mut edges.internal_owner_id_for_edge {
                if let Some(replacement) = internal_remap.get(identifier) {
                    *identifier = *replacement;
                }
            }
        }
    }
}

pub(super) fn prune_internal_owner_edges(
    records: &mut [tsce::CellRecordExpandedArchive],
    removed_internal_ids: &HashSet<u32>,
) -> bool {
    let mut changed = false;
    for record in records {
        if let Some(edges) = &mut record.expanded_edges {
            let previous_len = edges.internal_owner_id_for_edge.len();
            edges
                .internal_owner_id_for_edge
                .retain(|identifier| !removed_internal_ids.contains(identifier));
            changed |= edges.internal_owner_id_for_edge.len() != previous_len;
        }
    }
    changed
}

pub(super) fn prune_formula_owner_cell_edges_wire(
    original: &[u8],
    expected: &tsce::FormulaOwnerDependenciesArchive,
    removed_internal_ids: &HashSet<u32>,
) -> Result<Vec<u8>> {
    let data = transform_length_delimited_fields_at_path(original, &[4, 1, 6], |edges| {
        prune_repeated_internal_ids(edges, 5, removed_internal_ids)
    })?;
    if tsce::FormulaOwnerDependenciesArchive::decode(data.as_slice())? != *expected {
        return Err(Error::InvalidFormat(
            "iWork formula-owner dependency pruning failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn prune_cell_tile_edges_wire(
    original: &[u8],
    expected: &tsce::CellRecordTileArchive,
    removed_internal_ids: &HashSet<u32>,
) -> Result<Vec<u8>> {
    let data = transform_length_delimited_fields_at_path(original, &[4, 6], |edges| {
        prune_repeated_internal_ids(edges, 5, removed_internal_ids)
    })?;
    if tsce::CellRecordTileArchive::decode(data.as_slice())? != *expected {
        return Err(Error::InvalidFormat(
            "iWork dependency-tile pruning failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_formula_owner_wire(
    original: &[u8],
    previous: &tsce::FormulaOwnerDependenciesArchive,
    expected: &tsce::FormulaOwnerDependenciesArchive,
    object_remap: &HashMap<u64, u64>,
    internal_remap: &HashMap<u32, u32>,
    uuid_remap: &HashMap<(u64, u64), tsp::Uuid>,
) -> Result<Vec<u8>> {
    let mut data = remap_uuid_at_path(original, &[1], uuid_remap)?;
    data = patch_varint_field(
        &data,
        2,
        true,
        Some(u64::from(expected.internal_formula_owner_id)),
    )?;
    if previous.formula_owner.is_some() {
        data = patch_nested_varint_field(
            &data,
            &[11, 1],
            true,
            expected
                .formula_owner
                .as_ref()
                .map(|value| value.identifier),
        )?;
    }
    if previous.base_owner_uid.is_some() {
        data = remap_uuid_at_path(&data, &[12], uuid_remap)?;
    }
    data = transform_length_delimited_fields_at_path(&data, &[4, 1, 6], |edges| {
        remap_repeated_internal_ids(edges, 5, internal_remap)
    })?;
    data = transform_length_delimited_fields_at_path(&data, &[5, 2, 4], |reference| {
        let decoded = tsce::InternalRangeReferenceArchive::decode(reference)?;
        let replacement = internal_remap
            .get(&decoded.owner_id)
            .copied()
            .unwrap_or(decoded.owner_id);
        patch_varint_field(reference, 1, true, Some(u64::from(replacement)))
    })?;
    data = remap_cfuuid_at_path(&data, &[5, 2, 3, 1], uuid_remap)?;
    data = transform_length_delimited_fields_at_path(&data, &[13, 1], |reference| {
        let decoded = tsp::Reference::decode(reference)?;
        let replacement = object_remap
            .get(&decoded.identifier)
            .copied()
            .unwrap_or(decoded.identifier);
        patch_varint_field(reference, 1, true, Some(replacement))
    })?;
    data = remap_uuid_at_path(&data, &[14, 1, 1], uuid_remap)?;
    data = remap_uuid_at_path(&data, &[14, 2, 1], uuid_remap)?;
    data = remap_uuid_at_path(&data, &[14, 2, 2, 1], uuid_remap)?;
    data = transform_length_delimited_fields_at_path(&data, &[15, 1], |reference| {
        let decoded = tsp::Reference::decode(reference)?;
        let replacement = object_remap
            .get(&decoded.identifier)
            .copied()
            .unwrap_or(decoded.identifier);
        patch_varint_field(reference, 1, true, Some(replacement))
    })?;
    if tsce::FormulaOwnerDependenciesArchive::decode(data.as_slice())? != *expected {
        return Err(Error::InvalidFormat(
            "Numbers formula owner clone failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_cell_tile_wire(
    original: &[u8],
    previous: &tsce::CellRecordTileArchive,
    expected: &tsce::CellRecordTileArchive,
    internal_remap: &HashMap<u32, u32>,
) -> Result<Vec<u8>> {
    let mut data = patch_varint_field(
        original,
        1,
        true,
        Some(u64::from(expected.internal_owner_id)),
    )?;
    data = transform_length_delimited_fields_at_path(&data, &[4, 6], |edges| {
        remap_repeated_internal_ids(edges, 5, internal_remap)
    })?;
    if previous.tile_column_begin != expected.tile_column_begin
        || previous.tile_row_begin != expected.tile_row_begin
        || tsce::CellRecordTileArchive::decode(data.as_slice())? != *expected
    {
        return Err(Error::InvalidFormat(
            "Numbers dependency-tile clone failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}

pub(super) fn remap_range_tile_wire(
    original: &[u8],
    previous: &tsce::RangePrecedentsTileArchive,
    expected: &tsce::RangePrecedentsTileArchive,
) -> Result<Vec<u8>> {
    let data = patch_varint_field(original, 1, true, Some(u64::from(expected.to_owner_id)))?;
    if previous.from_to_range != expected.from_to_range
        || tsce::RangePrecedentsTileArchive::decode(data.as_slice())? != *expected
    {
        return Err(Error::InvalidFormat(
            "Numbers range dependency-tile clone failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}

fn remap_cfuuid_at_path(
    data: &[u8],
    path: &[u32],
    remap: &HashMap<(u64, u64), tsp::Uuid>,
) -> Result<Vec<u8>> {
    transform_length_delimited_fields_at_path(data, path, |uuid_data| {
        let decoded = tsp::CfuuidArchive::decode(uuid_data)?;
        let Some(key) = cfuuid_key(&decoded) else {
            return Ok(uuid_data.to_vec());
        };
        let Some(replacement) = remap.get(&key) else {
            return Ok(uuid_data.to_vec());
        };
        let mut rewritten = uuid_data.to_vec();
        if decoded.uuid_bytes.is_some() {
            rewritten = crate::wire::patch_length_delimited_field(
                &rewritten,
                1,
                true,
                Some(&uuid_bytes(replacement)),
            )?;
        }
        for (field, present, value) in [
            (2, decoded.uuid_w0.is_some(), replacement.lower as u32),
            (
                3,
                decoded.uuid_w1.is_some(),
                (replacement.lower >> 32) as u32,
            ),
            (4, decoded.uuid_w2.is_some(), replacement.upper as u32),
            (
                5,
                decoded.uuid_w3.is_some(),
                (replacement.upper >> 32) as u32,
            ),
        ] {
            if present {
                rewritten = patch_varint_field(&rewritten, field, true, Some(u64::from(value)))?;
            }
        }
        Ok(rewritten)
    })
}

fn cfuuid_key(uuid: &tsp::CfuuidArchive) -> Option<(u64, u64)> {
    let words = || {
        Some((
            u64::from(uuid.uuid_w0?) | (u64::from(uuid.uuid_w1?) << 32),
            u64::from(uuid.uuid_w2?) | (u64::from(uuid.uuid_w3?) << 32),
        ))
    };
    let bytes = || {
        let bytes: [u8; 16] = uuid.uuid_bytes.as_deref()?.try_into().ok()?;
        let value = u128::from_be_bytes(bytes);
        Some((value as u64, (value >> 64) as u64))
    };
    words().or_else(bytes)
}

fn uuid_bytes(uuid: &tsp::Uuid) -> [u8; 16] {
    ((u128::from(uuid.upper) << 64) | u128::from(uuid.lower)).to_be_bytes()
}

fn remap_repeated_internal_ids(
    data: &[u8],
    field_number: u32,
    remap: &HashMap<u32, u32>,
) -> Result<Vec<u8>> {
    let values = crate::wire::repeated_varint_values(data, field_number)?;
    let replacements = values
        .into_iter()
        .map(|identifier| {
            u32::try_from(identifier)
                .ok()
                .and_then(|identifier| remap.get(&identifier).copied())
                .map_or(identifier, u64::from)
        })
        .collect::<Vec<_>>();
    crate::wire::rewrite_repeated_varint_fields(data, field_number, &replacements)
}

fn prune_repeated_internal_ids(
    data: &[u8],
    field_number: u32,
    removed_internal_ids: &HashSet<u32>,
) -> Result<Vec<u8>> {
    let values = crate::wire::repeated_varint_values(data, field_number)?;
    let replacements = values
        .into_iter()
        .filter(|identifier| {
            u32::try_from(*identifier)
                .map(|identifier| !removed_internal_ids.contains(&identifier))
                .unwrap_or(true)
        })
        .collect::<Vec<_>>();
    crate::wire::rewrite_repeated_varint_fields(data, field_number, &replacements)
}
