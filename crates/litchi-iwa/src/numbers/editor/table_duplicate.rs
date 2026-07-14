//! Populated table duplication and wire-preserving object-graph cloning.

use super::*;

pub(super) fn reject_formula_table_duplication(
    package: &IWorkPackage,
    model: &TableModelArchive,
) -> Result<()> {
    let formula_table = model.base_data_store.formula_table.identifier;
    if formula_table == 0 {
        return Ok(());
    }
    let locations = object_locations(package)?;
    let formulas = resolve_table_data_list(
        package,
        &locations,
        formula_table,
        tst::table_data_list::ListType::Formula,
    )?;
    if formulas.entries.is_empty() {
        Ok(())
    } else {
        Err(Error::ParseError(
            "Cannot duplicate a Numbers table containing formulas".to_owned(),
        ))
    }
}

pub(super) fn table_owned_graph(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    model: &TableModelArchive,
) -> Result<BTreeMap<u64, TableOwnedKind>> {
    let mut graph = table_owned_objects(model)?;
    let mut pending = graph
        .iter()
        .filter_map(|(&identifier, kind)| (*kind == TableOwnedKind::Data).then_some(identifier))
        .collect::<Vec<_>>();
    let mut visited = HashSet::new();

    while let Some(identifier) = pending.pop() {
        if !visited.insert(identifier) {
            continue;
        }
        let archive_name = locations.get(&identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table storage object {identifier} is missing"
            ))
        })?;
        let archive = package.archive(archive_name)?;
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table storage object {identifier} is missing"
            ))
        })?;
        for message in &object.messages {
            let Ok(list) = TableDataList::decode(message.data.as_slice()) else {
                continue;
            };
            if tst::table_data_list::ListType::try_from(list.list_type).is_err() {
                continue;
            }
            for segment in list.segments {
                if segment.identifier == 0 {
                    continue;
                }
                match graph.insert(segment.identifier, TableOwnedKind::Data) {
                    Some(previous) if previous != TableOwnedKind::Data => {
                        return Err(Error::InvalidFormat(format!(
                            "Numbers table object {} has conflicting storage roles",
                            segment.identifier
                        )));
                    },
                    Some(_) => {},
                    None => pending.push(segment.identifier),
                }
            }
        }
    }

    for &identifier in graph.keys() {
        if !locations.contains_key(&identifier) {
            return Err(Error::InvalidFormat(format!(
                "Numbers table storage object {identifier} is missing"
            )));
        }
    }
    Ok(graph)
}

pub(super) fn duplicate_table_name(source: &str, existing: &HashSet<&str>) -> Result<String> {
    validate_name(source, "table")?;
    let base = format!("{source} copy");
    if !existing.contains(base.as_str()) {
        return Ok(base);
    }
    for suffix in 2u32..=u32::MAX {
        let candidate = format!("{base} {suffix}");
        if !existing.contains(candidate.as_str()) {
            return Ok(candidate);
        }
    }
    Err(Error::ParseError(
        "Unable to allocate a unique Numbers table name".to_owned(),
    ))
}

pub(super) fn duplicate_table_model_wire(
    original: &[u8],
    source: &TableModelArchive,
    remap: &HashMap<u64, u64>,
    table_uuid: &str,
    name: &str,
) -> Result<Vec<u8>> {
    const OWNED_REFERENCE_PATHS: &[&[u32]] = &[
        &[4, 1, 2],
        &[4, 2],
        &[4, 3, 1, 2],
        &[4, 4],
        &[4, 5],
        &[4, 6],
        &[4, 11],
        &[4, 12],
        &[4, 13],
        &[4, 15],
        &[4, 16],
        &[4, 17],
        &[4, 18],
        &[4, 19],
        &[4, 20],
        &[4, 21],
        &[4, 22],
        &[46],
        &[49],
    ];

    let decoded = TableModelArchive::decode(original)?;
    if &decoded != source {
        return Err(Error::InvalidFormat(
            "Numbers table model changed before duplication".to_owned(),
        ));
    }
    let mut expected = source.clone();
    expected.table_id = table_uuid.to_owned();
    expected.table_name = name.to_owned();
    remap_table_model_owned_references(&mut expected, remap);

    let data = patch_length_delimited_field(original, 1, true, Some(table_uuid.as_bytes()))?;
    let data = patch_length_delimited_field(&data, 8, true, Some(name.as_bytes()))?;
    let data = remap_numbers_reference_paths(&data, OWNED_REFERENCE_PATHS, remap)?;
    if TableModelArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers duplicated table model failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}

fn remap_table_model_owned_references(model: &mut TableModelArchive, remap: &HashMap<u64, u64>) {
    let store = &mut model.base_data_store;
    for reference in &mut store.row_headers.buckets {
        remap_table_reference(reference, remap);
    }
    remap_table_reference(&mut store.column_headers, remap);
    for tile in &mut store.tiles.tiles {
        remap_table_reference(&mut tile.tile, remap);
    }
    for reference in [
        &mut store.string_table,
        &mut store.style_table,
        &mut store.formula_table,
        &mut store.format_table_pre_bnc,
    ] {
        remap_table_reference(reference, remap);
    }
    for reference in [
        &mut store.formula_error_table,
        &mut store.multiple_choice_list_format_table,
        &mut store.merge_region_map,
        &mut store.deprecated_custom_format_table,
        &mut store.rich_text_table,
        &mut store.conditionalstyletable,
        &mut store.comment_storage_table,
        &mut store.import_warning_set_table,
        &mut store.control_cell_spec_table,
        &mut store.format_table,
        &mut model.base_column_row_uids,
        &mut model.stroke_sidecar,
    ] {
        remap_optional_table_reference(reference, remap);
    }
}

pub(super) fn duplicate_table_info_wire(
    original: &[u8],
    source: &tst::TableInfoArchive,
    remap: &HashMap<u64, u64>,
    offset: f32,
) -> Result<Vec<u8>> {
    const REFERENCE_PATHS: &[&[u32]] = &[
        &[1, 2],
        &[1, 6],
        &[1, 9],
        &[2],
        &[3],
        &[4],
        &[5],
        &[6],
        &[15],
        &[17],
    ];
    if !offset.is_finite() {
        return Err(Error::ParseError(
            "Numbers table duplicate offset must be finite".to_owned(),
        ));
    }
    let decoded = tst::TableInfoArchive::decode(original)?;
    if &decoded != source {
        return Err(Error::InvalidFormat(
            "Numbers table info changed before duplication".to_owned(),
        ));
    }
    let mut expected = source.clone();
    remap_table_info_references(&mut expected, remap);
    let mut data = remap_numbers_reference_paths(original, REFERENCE_PATHS, remap)?;
    if let Some(position) = expected
        .super_
        .geometry
        .as_mut()
        .and_then(|value| value.position.as_mut())
    {
        position.x += offset;
        position.y += offset;
        if !position.x.is_finite() || !position.y.is_finite() {
            return Err(Error::ParseError(
                "Numbers table duplicate position overflow".to_owned(),
            ));
        }
        data = patch_nested_fixed32_field(&data, &[1, 1, 1, 1], true, Some(position.x.to_bits()))?;
        data = patch_nested_fixed32_field(&data, &[1, 1, 1, 2], true, Some(position.y.to_bits()))?;
    }
    if tst::TableInfoArchive::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers duplicated table info failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}

#[allow(deprecated)]
fn remap_table_info_references(info: &mut tst::TableInfoArchive, remap: &HashMap<u64, u64>) {
    remap_numbers_required_reference(&mut info.table_model, remap);
    for reference in [
        &mut info.super_.parent,
        &mut info.super_.comment,
        &mut info.editing_state,
        &mut info.summary_model,
        &mut info.category_order,
        &mut info.view_column_row_uids,
        &mut info.pivot_data_model,
        &mut info.pivot_order,
    ] {
        remap_numbers_reference(reference, remap);
    }
    for reference in &mut info.super_.pencil_annotations {
        remap_table_reference(reference, remap);
    }
}

pub(super) fn clone_table_storage_object(
    source: &ArchiveObject,
    remap: &HashMap<u64, u64>,
) -> Result<ArchiveObject> {
    let source_id = source.archive_info.identifier.ok_or_else(|| {
        Error::InvalidFormat("Numbers table storage object has no identifier".to_owned())
    })?;
    let new_identifier = remap.get(&source_id).copied().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "No clone identifier allocated for Numbers table storage object {source_id}"
        ))
    })?;
    let mut messages = Vec::with_capacity(source.messages.len());
    for (message, info) in source
        .messages
        .iter()
        .zip(&source.archive_info.message_infos)
    {
        let private_references = info
            .object_references
            .iter()
            .filter(|identifier| remap.contains_key(identifier))
            .copied()
            .collect::<HashSet<_>>();
        let data = if private_references.is_empty() {
            message.data.clone()
        } else {
            remap_table_data_list_segments(&message.data, &private_references, remap)?
        };
        messages.push(RawMessage {
            type_: message.type_,
            data,
        });
    }
    clone_numbers_object_metadata(source, new_identifier, messages, remap)
}

fn remap_table_data_list_segments(
    original: &[u8],
    private_references: &HashSet<u64>,
    remap: &HashMap<u64, u64>,
) -> Result<Vec<u8>> {
    let mut expected = TableDataList::decode(original).map_err(|_| {
        Error::InvalidFormat(
            "Cannot safely clone Numbers table storage with private object references".to_owned(),
        )
    })?;
    let segments = expected
        .segments
        .iter()
        .map(|reference| reference.identifier)
        .collect::<HashSet<_>>();
    if !private_references.is_subset(&segments) {
        return Err(Error::InvalidFormat(
            "Numbers table data list has unsupported private object references".to_owned(),
        ));
    }
    for reference in &mut expected.segments {
        remap_table_reference(reference, remap);
    }
    let data = remap_numbers_reference_paths(original, &[&[4]], remap)?;
    if TableDataList::decode(data.as_slice())? != expected {
        return Err(Error::InvalidFormat(
            "Numbers table data-list clone failed wire validation".to_owned(),
        ));
    }
    Ok(data)
}
