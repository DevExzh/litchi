//! Wire-preserving CalculationEngine owner and dependency-tile rewrites.

use super::*;
use litchi_iwa_common::{LimitKind, WireLimits};
use litchi_iwa_protos::numbers_table_cell_dependency_codec as dependency_codec;

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
    let options = dependency_codec_options(original);
    let engine = dependency_codec::decode_calculation_engine(original, options)
        .map_err(|error| dependency_codec_error("Numbers CalculationEngine", error))?;
    let tracker = dependency_codec::decode_dependency_tracker(engine.dependency_tracker(), options)
        .map_err(|error| {
            dependency_codec_error("Numbers CalculationEngine dependency tracker", error)
        })?;
    let previous_count = tracker.number_of_formulas().ok_or_else(|| {
        Error::InvalidFormat("Numbers CalculationEngine has no formula count".to_owned())
    })?;
    previous_count.checked_sub(1).ok_or_else(|| {
        Error::InvalidFormat("Numbers formula count cannot be decremented below zero".to_owned())
    })?;
    dependency_codec::rewrite_calculation_engine_owner_removal(original, &[], &[], 1, options)
        .map(|(data, _report)| data)
        .map_err(|error| {
            dependency_codec_error("Numbers CalculationEngine formula-count decrement", error)
        })
}

pub(super) fn remove_formula_owners_from_engine(
    original: &[u8],
    owner_ids: &HashSet<u64>,
    internal_owner_ids: &HashSet<u32>,
    formula_count: u64,
) -> Result<Vec<u8>> {
    let mut owner_ids = owner_ids.iter().copied().collect::<Vec<_>>();
    owner_ids.sort_unstable();
    let mut internal_owner_ids = internal_owner_ids.iter().copied().collect::<Vec<_>>();
    internal_owner_ids.sort_unstable();
    let options = dependency_codec_options(original);
    let engine = dependency_codec::decode_calculation_engine(original, options)
        .map_err(|error| dependency_codec_error("Numbers CalculationEngine", error))?;
    let tracker = dependency_codec::decode_dependency_tracker(engine.dependency_tracker(), options)
        .map_err(|error| {
            dependency_codec_error("Numbers CalculationEngine dependency tracker", error)
        })?;
    tracker
        .number_of_formulas()
        .unwrap_or_default()
        .checked_sub(formula_count)
        .ok_or_else(|| Error::InvalidFormat("Numbers formula count underflow".to_owned()))?;
    dependency_codec::rewrite_calculation_engine_owner_removal(
        original,
        &owner_ids,
        &internal_owner_ids,
        formula_count,
        options,
    )
    .map(|(data, _report)| data)
    .map_err(|error| {
        dependency_codec_error("Numbers CalculationEngine formula-owner removal", error)
    })
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
    _expected: &tsce::FormulaOwnerDependenciesArchive,
    removed_internal_ids: &HashSet<u32>,
) -> Result<Vec<u8>> {
    let mut removed_internal_ids = removed_internal_ids.iter().copied().collect::<Vec<_>>();
    removed_internal_ids.sort_unstable();
    let options = dependency_codec_options(original);
    dependency_codec::rewrite_formula_owner_cell_edges(original, &removed_internal_ids, options)
        .map(|(data, _report)| data)
        .map_err(|error| dependency_codec_error("iWork formula-owner dependency pruning", error))
}

pub(super) fn prune_cell_tile_edges_wire(
    original: &[u8],
    _expected: &tsce::CellRecordTileArchive,
    removed_internal_ids: &HashSet<u32>,
) -> Result<Vec<u8>> {
    let mut removed_internal_ids = removed_internal_ids.iter().copied().collect::<Vec<_>>();
    removed_internal_ids.sort_unstable();
    let options = dependency_codec_options(original);
    dependency_codec::rewrite_cell_record_tile_edges(original, &removed_internal_ids, options)
        .map(|(data, _report)| data)
        .map_err(|error| dependency_codec_error("iWork dependency-tile pruning", error))
}

const DEPENDENCY_CODEC_RECURSION_LIMIT: u32 = 64;
const DEPENDENCY_CODEC_MAX_REFERENCES: usize = litchi_numbers::MAX_REFERENCES;
const DEPENDENCY_CODEC_OUTPUT_SLACK: usize = 128;
// The rewrite performs bounded source measurement, one candidate pass, and
// strict result validation in addition to the initial projection scan.
const DEPENDENCY_CODEC_REWRITE_WORK_FACTOR: usize = 64;

fn dependency_codec_options(source: &[u8]) -> dependency_codec::DecodeOptions {
    let max_message_bytes = source
        .len()
        .checked_add(DEPENDENCY_CODEC_OUTPUT_SLACK)
        .map_or(WireLimits::MAX_OUTPUT_BYTES, |value| {
            value.min(WireLimits::MAX_OUTPUT_BYTES)
        })
        .clamp(1, WireLimits::MAX_OUTPUT_BYTES);
    dependency_codec::DecodeOptions::new(
        max_message_bytes,
        source.len().clamp(1, WireLimits::MAX_FIELDS),
        source
            .len()
            .saturating_mul(DEPENDENCY_CODEC_REWRITE_WORK_FACTOR)
            .clamp(1, WireLimits::MAX_REWRITE_WORK),
        DEPENDENCY_CODEC_RECURSION_LIMIT,
        source.len().clamp(1, DEPENDENCY_CODEC_MAX_REFERENCES),
        0,
    )
}

fn dependency_codec_error(context: &str, error: dependency_codec::DecodeError) -> Error {
    use dependency_codec::DecodeLimit;

    match error.resource_limit() {
        Some(DecodeLimit::Bytes { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::InputBytes,
                observed,
                limit: maximum,
            })
        },
        Some(DecodeLimit::Fields { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Fields,
                observed,
                limit: maximum,
            })
        },
        Some(DecodeLimit::Work { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::RewriteWork,
                observed,
                limit: maximum,
            })
        },
        Some(DecodeLimit::Nesting { observed, maximum }) => {
            Error::IwaCommon(litchi_iwa_common::Error::LimitExceeded {
                kind: LimitKind::Nesting,
                observed: observed as usize,
                limit: maximum as usize,
            })
        },
        Some(DecodeLimit::Allocation { requested }) => {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Numbers formula dependency projection",
                amount: requested,
            })
        },
        Some(DecodeLimit::References { observed, maximum }) => Error::InvalidFormat(format!(
            "{context} exceeded its aggregate reference limit: observed {observed}, limit {maximum}"
        )),
        Some(DecodeLimit::Text { observed, maximum }) => Error::InvalidFormat(format!(
            "{context} exceeded its aggregate text limit: observed {observed}, limit {maximum}"
        )),
        Some(DecodeLimit::Retained { observed, maximum }) => Error::InvalidFormat(format!(
            "{context} exceeded its retained-byte limit: observed {observed}, limit {maximum}"
        )),
        Some(_) => Error::InvalidFormat(format!(
            "{context} exceeded an unsupported strict resource limit"
        )),
        None => Error::InvalidFormat(format!("{context} failed strict validation: {error}")),
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

#[cfg(test)]
#[allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "The focused host wire fixture is intentionally infallible."
)]
mod tests {
    use super::*;

    fn append_unknown_varint(data: &mut Vec<u8>, field_number: u32, value: u64) {
        data.extend(litchi_iwa_common::varint::encode_varint(
            u64::from(field_number) << 3,
        ));
        data.extend(litchi_iwa_common::varint::encode_varint(value));
    }

    fn owner_with_edges() -> tsce::FormulaOwnerDependenciesArchive {
        tsce::FormulaOwnerDependenciesArchive {
            formula_owner_uid: tsp::Uuid { lower: 1, upper: 2 },
            internal_formula_owner_id: 3,
            cell_dependencies: Some(tsce::CellDependenciesExpandedArchive {
                cell_record: vec![tsce::CellRecordExpandedArchive {
                    column: 1,
                    row: 2,
                    expanded_edges: Some(tsce::ExpandedEdgesArchive {
                        edge_without_owner_rows: vec![1, 2, 3],
                        edge_without_owner_columns: vec![11, 12, 13],
                        edge_with_owner_rows: vec![21, 22, 23],
                        edge_with_owner_columns: vec![31, 32, 33],
                        internal_owner_id_for_edge: vec![7, 8, 7],
                    }),
                    ..Default::default()
                }],
            }),
            ..Default::default()
        }
    }

    #[test]
    fn owner_edge_pruning_keeps_parallel_arrays_and_unknown_spans() {
        let baseline = owner_with_edges().encode_to_vec();
        let mut original = crate::wire::transform_length_delimited_fields_at_path(
            &baseline,
            &[4, 1, 6],
            |payload| {
                let mut payload = payload.to_vec();
                append_unknown_varint(&mut payload, 97, 970);
                Ok(payload)
            },
        )
        .unwrap();
        append_unknown_varint(&mut original, 98, 980);
        let expected = tsce::FormulaOwnerDependenciesArchive::decode(original.as_slice()).unwrap();
        let rewritten =
            prune_formula_owner_cell_edges_wire(&original, &expected, &HashSet::from([7])).unwrap();

        assert!(rewritten.ends_with(&[0x90, 0x06, 0xd4, 0x07]));
        let rewritten_owner =
            tsce::FormulaOwnerDependenciesArchive::decode(rewritten.as_slice()).unwrap();
        let edges = rewritten_owner
            .cell_dependencies
            .as_ref()
            .unwrap()
            .cell_record[0]
            .expanded_edges
            .as_ref()
            .unwrap();
        assert_eq!(edges.edge_without_owner_rows, [1, 2, 3]);
        assert_eq!(edges.edge_without_owner_columns, [11, 12, 13]);
        assert_eq!(edges.edge_with_owner_rows, [22]);
        assert_eq!(edges.edge_with_owner_columns, [32]);
        assert_eq!(edges.internal_owner_id_for_edge, [8]);
        let expanded = crate::wire::repeated_length_delimited_payloads(&rewritten, 4).unwrap();
        let record = crate::wire::repeated_length_delimited_payloads(expanded[0], 1).unwrap();
        let edge_payload = crate::wire::repeated_length_delimited_payloads(record[0], 6).unwrap();
        assert!(edge_payload[0].ends_with(&[0x88, 0x06, 0xca, 0x07]));
    }

    #[test]
    fn tile_edge_pruning_uses_the_same_parallel_array_mask() {
        let owner = owner_with_edges();
        let record = owner.cell_dependencies.unwrap().cell_record[0].clone();
        let baseline = tsce::CellRecordTileArchive {
            internal_owner_id: 3,
            tile_column_begin: 0,
            tile_row_begin: 0,
            cell_records: vec![record],
        }
        .encode_to_vec();
        let mut original =
            crate::wire::transform_length_delimited_fields_at_path(&baseline, &[4, 6], |payload| {
                let mut payload = payload.to_vec();
                append_unknown_varint(&mut payload, 96, 960);
                Ok(payload)
            })
            .unwrap();
        append_unknown_varint(&mut original, 95, 950);
        let expected = tsce::CellRecordTileArchive::decode(original.as_slice()).unwrap();
        let rewritten =
            prune_cell_tile_edges_wire(&original, &expected, &HashSet::from([7])).unwrap();
        assert!(rewritten.ends_with(&[0xf8, 0x05, 0xb6, 0x07]));
        let tile = tsce::CellRecordTileArchive::decode(rewritten.as_slice()).unwrap();
        let edges = tile.cell_records[0].expanded_edges.as_ref().unwrap();
        assert_eq!(edges.edge_without_owner_rows, [1, 2, 3]);
        assert_eq!(edges.edge_without_owner_columns, [11, 12, 13]);
        assert_eq!(edges.edge_with_owner_rows, [22]);
        assert_eq!(edges.edge_with_owner_columns, [32]);
        assert_eq!(edges.internal_owner_id_for_edge, [8]);
    }

    #[test]
    fn engine_owner_removal_keeps_unrelated_references_and_unknown_spans() {
        let baseline = tsce::CalculationEngineArchive {
            dependency_tracker: tsce::DependencyTrackerArchive {
                formula_owner_dependencies: vec![
                    tsp::Reference {
                        identifier: 102,
                        ..Default::default()
                    },
                    tsp::Reference {
                        identifier: 103,
                        ..Default::default()
                    },
                ],
                number_of_formulas: Some(2),
                ..Default::default()
            },
            ..Default::default()
        }
        .encode_to_vec();
        let mut original =
            crate::wire::transform_length_delimited_fields_at_path(&baseline, &[2], |payload| {
                let mut payload = payload.to_vec();
                append_unknown_varint(&mut payload, 97, 970);
                Ok(payload)
            })
            .unwrap();
        append_unknown_varint(&mut original, 98, 980);

        let rewritten =
            remove_formula_owners_from_engine(&original, &HashSet::from([102]), &HashSet::new(), 1)
                .unwrap();
        assert!(rewritten.ends_with(&[0x90, 0x06, 0xd4, 0x07]));
        let engine = tsce::CalculationEngineArchive::decode(rewritten.as_slice()).unwrap();
        assert_eq!(engine.dependency_tracker.number_of_formulas, Some(1));
        assert_eq!(
            engine
                .dependency_tracker
                .formula_owner_dependencies
                .iter()
                .map(|reference| reference.identifier)
                .collect::<Vec<_>>(),
            [103]
        );
        let tracker = crate::wire::repeated_length_delimited_payloads(&rewritten, 2).unwrap();
        assert!(tracker[0].ends_with(&[0x88, 0x06, 0xca, 0x07]));
    }
}
