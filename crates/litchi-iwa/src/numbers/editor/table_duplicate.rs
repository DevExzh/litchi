//! Populated table duplication and wire-preserving object-graph cloning.

use super::*;

const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];

/// Fresh object identifiers for a cloned attached table graph.
///
/// The caller owns attaching `info_object_id` to its enclosing Pages body,
/// Numbers sheet, or Keynote slide.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(crate) struct AttachedTableGraphClone {
    pub(crate) info_object_id: u64,
    pub(crate) model_object_id: u64,
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

pub(crate) fn duplicate_table_name(source: &str, existing: &HashSet<&str>) -> Result<String> {
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
        "Unable to allocate a unique iWork table name".to_owned(),
    ))
}

/// Clone one populated attached table into `staged` without attaching it.
///
/// `source` is a stable pre-mutation snapshot. The clone gets fresh object
/// identifiers, table UUID, private storage, and CalculationEngine owner
/// family; references to the existing parent stay intact. Callers add the
/// returned drawable to their native container only after this succeeds.
pub(crate) fn duplicate_attached_table_graph_in_package(
    source: &IWorkPackage,
    staged: &mut IWorkPackage,
    source_info_id: u64,
    source_model_id: u64,
    name: &str,
    position_offset: f32,
) -> Result<AttachedTableGraphClone> {
    validate_name(name, "table")?;
    if !position_offset.is_finite() {
        return Err(Error::ParseError(
            "iWork table duplicate offset must be finite".to_owned(),
        ));
    }

    let locations = object_locations(source)?;
    let info_archive_name = locations.get(&source_info_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table info {source_info_id} is missing"))
    })?;
    let info_archive = source.archive(info_archive_name)?;
    let info_object = info_archive.object(source_info_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table info {source_info_id} is missing"))
    })?;
    let (info_message_index, source_info) = decode_table_info(info_object)?;
    if source_info.table_model.identifier != source_model_id {
        return Err(Error::InvalidFormat(format!(
            "iWork table info {source_info_id} points to model {}, expected {source_model_id}",
            source_info.table_model.identifier
        )));
    }

    let model_archive_name = locations.get(&source_model_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table model {source_model_id} is missing"))
    })?;
    let model_archive = source.archive(model_archive_name)?;
    let model_object = model_archive.object(source_model_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table model {source_model_id} is missing"))
    })?;
    let model_message_index = find_table_model_message(model_object)?;
    let source_model =
        TableModelArchive::decode(model_object.messages[model_message_index].data.as_slice())?;

    let graph = table_owned_graph(source, &locations, &source_model)?;
    if graph.contains_key(&source_info_id) || graph.contains_key(&source_model_id) {
        return Err(Error::InvalidFormat(
            "iWork table graph aliases its drawable or model object".to_owned(),
        ));
    }

    let mut next_identifier = next_object_identifier(staged)?;
    let new_info_id = take_identifier(&mut next_identifier)?;
    let new_model_id = take_identifier(&mut next_identifier)?;
    let mut remap = HashMap::with_capacity(graph.len() + 2);
    remap.insert(source_info_id, new_info_id);
    remap.insert(source_model_id, new_model_id);
    for &identifier in graph.keys() {
        remap.insert(identifier, take_identifier(&mut next_identifier)?);
    }

    let existing_table_ids = table_uuids(staged)?;
    let existing_table_ids = existing_table_ids
        .iter()
        .map(String::as_str)
        .collect::<HashSet<_>>();
    let table_uuid = allocate_table_uuid(new_model_id, &existing_table_ids);

    let model_data = duplicate_table_model_wire(
        model_object.messages[model_message_index].data.as_slice(),
        &source_model,
        &remap,
        &table_uuid,
        name,
    )?;
    let mut objects = Vec::with_capacity(graph.len() + 2);
    objects.push((
        model_archive_name.clone(),
        clone_numbers_object_metadata(
            model_object,
            new_model_id,
            vec![RawMessage {
                type_: model_object.messages[model_message_index].type_,
                data: model_data,
            }],
            &remap,
        )?,
    ));

    let info_data = duplicate_table_info_wire(
        info_object.messages[info_message_index].data.as_slice(),
        &source_info,
        &remap,
        position_offset,
    )?;
    objects.push((
        info_archive_name.clone(),
        clone_numbers_object_metadata(
            info_object,
            new_info_id,
            vec![RawMessage {
                type_: info_object.messages[info_message_index].type_,
                data: info_data,
            }],
            &remap,
        )?,
    ));

    for &source_id in graph.keys() {
        let archive_name = locations.get(&source_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table storage object {source_id} is missing"))
        })?;
        let archive = source.archive(archive_name)?;
        let source_object = archive.object(source_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table storage object {source_id} is missing"))
        })?;
        let mut cloned = clone_table_storage_object(source_object, &remap)?;
        remap_cloned_formula_storage(&mut cloned, &source_model.table_id, &table_uuid)?;
        objects.push((archive_name.clone(), cloned));
    }

    for (archive_name, object) in objects {
        staged.update_archive(&archive_name, |archive| archive.insert_object(object))?;
    }
    register_cloned_numbers_objects(staged, source, &locations, &remap)?;

    if let Some(parent) = source_info.super_.parent.as_ref() {
        let parent_archive_name = locations.get(&parent.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork table parent {} is missing",
                parent.identifier
            ))
        })?;
        register_numbers_component_reference(
            staged,
            parent_archive_name,
            info_archive_name,
            new_info_id,
        )?;
    }

    if let Some((source_owner_uuid, new_owner_uuid)) =
        formula_graph_owner_uuids(staged, source_info_id, &source_model.table_id, &table_uuid)?
    {
        for &source_id in graph.keys() {
            let archive_name = locations.get(&source_id).ok_or_else(|| {
                Error::InvalidFormat(format!("iWork table storage object {source_id} is missing"))
            })?;
            let cloned_id = remap[&source_id];
            staged.update_archive(archive_name, |archive| {
                let object = archive.object_mut(cloned_id).ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "iWork cloned table storage object {cloned_id} is missing"
                    ))
                })?;
                remap_cloned_formula_owner_storage(object, &source_owner_uuid, &new_owner_uuid)
            })?;
        }
    }

    let table_last_identifier = next_identifier.checked_sub(1).ok_or_else(|| {
        Error::InvalidFormat("iWork table clone allocated no identifiers".to_owned())
    })?;
    set_package_last_object_identifier(staged, table_last_identifier)?;
    let calculation_engine_entry = staged.calculation_engine_entry_name()?.map(str::to_owned);
    clone_table_formula_graph(
        staged,
        source_info_id,
        new_info_id,
        &source_model.table_id,
        &table_uuid,
    )?;
    if let Some(calculation_engine_entry) = calculation_engine_entry {
        register_numbers_component_reference(
            staged,
            &calculation_engine_entry,
            info_archive_name,
            new_info_id,
        )?;
    }

    Ok(AttachedTableGraphClone {
        info_object_id: new_info_id,
        model_object_id: new_model_id,
    })
}

fn table_uuids(package: &IWorkPackage) -> Result<HashSet<String>> {
    let mut identifiers = HashSet::new();
    for archive_name in package.iwa_entry_names() {
        let archive = package.archive(archive_name)?;
        for object in &archive.objects {
            for message in object
                .messages
                .iter()
                .filter(|message| TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_))
            {
                if let Ok(model) = TableModelArchive::decode(message.data.as_slice()) {
                    identifiers.insert(model.table_id);
                }
            }
        }
    }
    Ok(identifiers)
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

pub(super) fn register_cloned_numbers_objects(
    staged: &mut IWorkPackage,
    source: &IWorkPackage,
    locations: &HashMap<u64, String>,
    remap: &HashMap<u64, u64>,
) -> Result<()> {
    let mut clone_entries = HashMap::with_capacity(remap.len());
    let mut uuid_additions = HashMap::<u64, Vec<u64>>::new();
    for (&source_id, &new_id) in remap {
        let entry = locations.get(&source_id).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers cloned object {source_id} has no source component"
            ))
        })?;
        clone_entries.insert(new_id, entry.clone());
        let Some(component) = component_identifier_for_entry(source, entry)? else {
            continue;
        };
        if component_uuid_identifiers(source, component)?
            .is_some_and(|identifiers| identifiers.contains(&source_id))
        {
            uuid_additions.entry(component).or_default().push(new_id);
        }
    }
    for (component, identifiers) in uuid_additions {
        add_component_object_uuids(staged, component, &identifiers)?;
    }

    let mut references = HashSet::new();
    for (&new_id, source_entry) in &clone_entries {
        let Some(source_component) = component_identifier_for_entry(staged, source_entry)? else {
            continue;
        };
        let archive = staged.archive(source_entry)?;
        let object = archive.object(new_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Numbers cloned object {new_id} is missing"))
        })?;
        for target_id in object
            .archive_info
            .message_infos
            .iter()
            .flat_map(|info| &info.object_references)
        {
            let Some(target_entry) = clone_entries.get(target_id) else {
                continue;
            };
            let Some(target_component) = component_identifier_for_entry(staged, target_entry)?
            else {
                continue;
            };
            if source_component != target_component {
                references.insert((source_component, target_component, *target_id));
            }
        }
    }
    for (source_component, target_component, object_id) in references {
        add_component_external_reference(staged, source_component, target_component, object_id)?;
    }
    Ok(())
}

pub(super) fn register_numbers_component_reference(
    package: &mut IWorkPackage,
    source_entry: &str,
    target_entry: &str,
    object_id: u64,
) -> Result<()> {
    let Some(source_component) = component_identifier_for_entry(package, source_entry)? else {
        return Ok(());
    };
    let Some(target_component) = component_identifier_for_entry(package, target_entry)? else {
        return Ok(());
    };
    if source_component != target_component {
        add_component_external_reference(package, source_component, target_component, object_id)?;
    }
    Ok(())
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
