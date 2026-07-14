//! Wire-preserving CalculationEngine owner and dependency-tile rewrites.

use super::*;

pub(super) fn append_formula_owners_to_engine(
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
    if let Some(dependencies) = &mut owner.tiled_cell_dependencies {
        for reference in &mut dependencies.cell_record_tiles {
            if let Some(replacement) = object_remap.get(&reference.identifier) {
                reference.identifier = *replacement;
            }
        }
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
    data = transform_length_delimited_fields_at_path(&data, &[13, 1], |reference| {
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
