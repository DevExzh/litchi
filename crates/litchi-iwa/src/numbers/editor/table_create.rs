//! Creation of an independent empty table graph from an existing native template.

use super::*;

const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[6_000, 6_001];

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct EmptyTableGraph {
    pub(super) info_object_id: u64,
    pub(super) model_object_id: u64,
}

/// Clone only the structural style of a table into a fresh empty graph.
///
/// The caller owns attachment of `info_object_id` to its Pages body or Numbers
/// sheet. Native storage and CalculationEngine state are registered here so
/// the returned model is immediately writable and formula-ready.
#[allow(deprecated)]
pub(super) fn create_empty_table_graph(
    package: &mut IWorkPackage,
    template_info_id: u64,
    template_model_id: u64,
    template_parent_id: u64,
    parent_id: u64,
    name: &str,
    rows: usize,
    columns: usize,
    position_offset: Option<f32>,
) -> Result<EmptyTableGraph> {
    validate_name(name, "table")?;
    let (rows_u32, columns_u32) = validate_table_dimensions(rows, columns)?;
    if position_offset.is_some_and(|offset| !offset.is_finite()) {
        return Err(Error::ParseError(
            "iWork table position offset must be finite".to_owned(),
        ));
    }

    let descriptors = attached_table_templates(package)?;
    let template = descriptors
        .iter()
        .find(|descriptor| descriptor.object_id == template_model_id)
        .ok_or_else(|| {
            Error::ParseError(format!(
                "iWork table model {template_model_id} is not available as a creation template"
            ))
        })?;
    let locations = object_locations(package)?;
    let template_info_archive = locations.get(&template_info_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table info {template_info_id} is missing"))
    })?;
    let template_info_component = package.archive(template_info_archive)?;
    let template_info_object = template_info_component
        .object(template_info_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table info {template_info_id} is missing"))
        })?;
    let (info_message_index, mut table_info) = decode_table_info(template_info_object)?;
    let source_parent_id = table_info
        .super_
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        .unwrap_or(template_parent_id);
    let template_model_archive = locations.get(&template_model_id).ok_or_else(|| {
        Error::InvalidFormat(format!("iWork table model {template_model_id} is missing"))
    })?;
    let template_model_component = package.archive(template_model_archive)?;
    let template_model_object = template_model_component
        .object(template_model_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table model {template_model_id} is missing"))
        })?;
    let model_message_index = find_table_model_message(template_model_object)?;

    let mut next_identifier = next_object_identifier(package)?;
    let new_info_id = take_identifier(&mut next_identifier)?;
    let new_model_id = take_identifier(&mut next_identifier)?;
    let owned_kinds = table_owned_objects(&template.model)?;
    let mut remap = HashMap::with_capacity(owned_kinds.len() + 2);
    remap.insert(template_info_id, new_info_id);
    remap.insert(template_model_id, new_model_id);
    for &identifier in owned_kinds.keys() {
        remap.insert(identifier, take_identifier(&mut next_identifier)?);
    }

    let existing_table_ids = descriptors
        .iter()
        .map(|descriptor| descriptor.model.table_id.as_str())
        .collect::<HashSet<_>>();
    let table_uuid = allocate_table_uuid(new_model_id, &existing_table_ids);
    let mut model = template.model.clone();
    prepare_empty_table_model(&mut model, &remap, &table_uuid, name, rows_u32, columns_u32)?;

    table_info.super_.parent = Some(tsp::Reference {
        identifier: parent_id,
        ..Default::default()
    });
    if let Some(offset) = position_offset
        && let Some(position) = table_info
            .super_
            .geometry
            .as_mut()
            .and_then(|geometry| geometry.position.as_mut())
    {
        position.x += offset;
        position.y += offset;
        if !position.x.is_finite() || !position.y.is_finite() {
            return Err(Error::ParseError(
                "iWork table position overflow".to_owned(),
            ));
        }
    }
    table_info.super_.comment = None;
    table_info.super_.pencil_annotations.clear();
    table_info.super_.title = None;
    table_info.super_.caption = None;
    table_info.table_model = tsp::Reference {
        identifier: new_model_id,
        ..Default::default()
    };
    table_info.editing_state = None;
    table_info.summary_model = None;
    table_info.category_order = None;
    table_info.view_column_row_uids = None;
    table_info.pivot_data_model = None;
    table_info.pivot_order = None;

    let mut info_remap = remap.clone();
    info_remap.insert(source_parent_id, parent_id);
    let mut objects = Vec::with_capacity(owned_kinds.len() + 2);
    objects.push((
        template_info_archive.clone(),
        clone_single_payload_object(
            template_info_object,
            new_info_id,
            info_message_index,
            table_info.encode_to_vec(),
            vec![parent_id, new_model_id],
            &info_remap,
            false,
        )?,
    ));
    objects.push((
        template_model_archive.clone(),
        clone_single_payload_object(
            template_model_object,
            new_model_id,
            model_message_index,
            model.encode_to_vec(),
            table_model_references(&model),
            &remap,
            false,
        )?,
    ));
    for (&source_id, &kind) in &owned_kinds {
        let archive_name = locations.get(&source_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table storage object {source_id} is missing"))
        })?;
        let source_archive = package.archive(archive_name)?;
        let source = source_archive.object(source_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table storage object {source_id} is missing"))
        })?;
        objects.push((
            archive_name.clone(),
            clone_empty_table_storage(
                source,
                remap[&source_id],
                kind,
                rows_u32,
                columns_u32,
                new_model_id,
            )?,
        ));
    }

    let source = package.clone();
    for (archive_name, object) in objects {
        package.update_archive(&archive_name, |archive| archive.insert_object(object))?;
    }
    register_cloned_numbers_objects(package, &source, &locations, &remap)?;
    let parent_archive = locations
        .get(&parent_id)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork parent {parent_id} is missing")))?;
    register_numbers_component_reference(
        package,
        parent_archive,
        template_info_archive,
        new_info_id,
    )?;
    let table_last_identifier = next_identifier.checked_sub(1).ok_or_else(|| {
        Error::InvalidFormat("iWork table creation allocated no identifiers".to_owned())
    })?;
    set_package_last_object_identifier(package, table_last_identifier)?;
    if let Some((calculation_engine_entry, _)) =
        create_empty_table_formula_graph(package, template_info_id, new_info_id, &table_uuid)?
    {
        register_numbers_component_reference(
            package,
            &calculation_engine_entry,
            template_info_archive,
            new_info_id,
        )?;
    }

    Ok(EmptyTableGraph {
        info_object_id: new_info_id,
        model_object_id: new_model_id,
    })
}

fn attached_table_templates(package: &IWorkPackage) -> Result<Vec<TableDescriptor>> {
    let locations = object_locations(package)?;
    let archive_names = locations.values().collect::<HashSet<_>>();
    let mut result = Vec::new();
    let mut seen_models = HashSet::new();
    for archive_name in archive_names {
        let archive = package.archive(archive_name)?;
        for info_object in &archive.objects {
            let Some(info_id) = info_object.archive_info.identifier else {
                continue;
            };
            for message in &info_object.messages {
                let Ok(info) = tst::TableInfoArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                let model_id = info.table_model.identifier;
                let Some(model_archive_name) = locations.get(&model_id) else {
                    continue;
                };
                let model_archive = package.archive(model_archive_name)?;
                let Some(model_object) = model_archive.object(model_id) else {
                    continue;
                };
                let models = model_object
                    .messages
                    .iter()
                    .filter(|message| TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_))
                    .filter_map(|message| TableModelArchive::decode(message.data.as_slice()).ok())
                    .collect::<Vec<_>>();
                let [model] = models.as_slice() else {
                    continue;
                };
                if !seen_models.insert(model_id) {
                    return Err(Error::InvalidFormat(format!(
                        "iWork table model {model_id} has multiple table-info owners"
                    )));
                }
                result.push(TableDescriptor {
                    object_id: model_id,
                    table_info_id: info_id,
                    model: model.clone(),
                });
                break;
            }
        }
    }
    Ok(result)
}
