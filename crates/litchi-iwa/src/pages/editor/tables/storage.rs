//! Native table graph discovery and package/object storage boundaries.

use super::*;

fn legacy_table_appearance_compatibility_read(
    editor: &PagesEditor,
    model_id: u64,
) -> Result<TableAppearance> {
    crate::table_appearance::table_appearance(editor.package(), model_id)
}

#[derive(Debug, Clone)]
pub(crate) struct PagesTableGraph {
    pub(super) info: PagesTableInfo,
    pub(super) attachment_object_id: u64,
    pub(super) formula_context_ids: Vec<u64>,
}

pub(crate) fn body_table_graphs(editor: &PagesEditor) -> Result<Vec<PagesTableGraph>> {
    let body: StorageArchive = decode_typed_package_object(
        editor.package(),
        editor.body_storage_id.get(),
        editor.body_storage()?.message_type,
        "TSWP.StorageArchive",
    )?;
    let attachments = body
        .table_attachment
        .as_ref()
        .map_or(&[][..], |table| table.entries.as_slice());
    let mut anchor_offsets = Vec::new();
    anchor_offsets
        .try_reserve_exact(attachments.len())
        .map_err(|_| {
            Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                resource: "Pages table attachment anchor offsets",
                amount: attachments.len(),
            })
        })?;
    anchor_offsets.extend(
        attachments
            .iter()
            .filter(|entry| entry.object.is_some())
            .map(|entry| entry.character_index),
    );
    retain_object_replacement_offsets(&editor.body_text()?, &mut anchor_offsets);
    let mut seen_drawables = HashSet::new();
    let mut seen_models = HashSet::new();
    let mut result = Vec::new();

    for entry in attachments {
        let Some(attachment_reference) = entry.object else {
            continue;
        };
        let Some(attachment) = decode_optional_typed_package_object::<DrawableAttachmentArchive>(
            editor.package(),
            attachment_reference.identifier,
            DRAWABLE_ATTACHMENT_MESSAGE_TYPE,
        )?
        else {
            continue;
        };
        let Some(drawable) = attachment.drawable else {
            continue;
        };
        let archive_name = find_object_archive(editor.package(), drawable.identifier)?;
        let archive = editor.package().archive(&archive_name)?;
        let object = archive.object(drawable.identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages table drawable {} is missing",
                drawable.identifier
            ))
        })?;
        let mut messages = object
            .messages
            .iter()
            .filter(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE);
        let Some(message) = messages.next() else {
            continue;
        };
        if messages.next().is_some() {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} repeats its table-info payload",
                drawable.identifier
            )));
        }
        let table_info = validation::decode_table_info_ownership(&message.data)?;
        if table_info.parent().map(std::num::NonZeroU64::get) != Some(editor.body_storage_id.get())
        {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} is not owned by the body",
                drawable.identifier
            )));
        }
        if anchor_offsets
            .binary_search(&entry.character_index)
            .is_err()
        {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} has no object-replacement character",
                drawable.identifier
            )));
        }
        let model_id = table_info.table_model().identifier().get();
        let model_archive_name = find_object_archive(editor.package(), model_id)?;
        let model_archive = editor.package().archive(&model_archive_name)?;
        let model_object = model_archive.object(model_id).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages table model {model_id} is missing"))
        })?;
        let model = decode_unique_table_model(
            model_object
                .messages
                .iter()
                .filter(|message| TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_)),
            model_id,
        )?;
        if !seen_drawables.insert(drawable.identifier) || !seen_models.insert(model_id) {
            return Err(Error::InvalidFormat(format!(
                "Pages table drawable {} or model {model_id} is attached more than once",
                drawable.identifier
            )));
        }
        let mut formula_context_ids = vec![drawable.identifier, model_id];
        for reference in object
            .archive_info
            .message_infos
            .iter()
            .chain(&model_object.archive_info.message_infos)
            .flat_map(|message| {
                message.object_references.iter().copied().chain(
                    message
                        .field_infos
                        .iter()
                        .flat_map(|field| field.object_references.iter().copied()),
                )
            })
        {
            if !formula_context_ids.contains(&reference) {
                formula_context_ids.push(reference);
            }
        }
        let mut name = String::new();
        name.try_reserve_exact(model.table_name().len())
            .map_err(|_| {
                Error::IwaCommon(litchi_iwa_common::Error::Allocation {
                    resource: "Pages body-table name",
                    amount: model.table_name().len(),
                })
            })?;
        name.push_str(model.table_name());
        result.push(PagesTableGraph {
            attachment_object_id: attachment_reference.identifier,
            formula_context_ids,
            info: PagesTableInfo {
                drawable_object_id: drawable.identifier,
                model_object_id: model_id,
                anchor_character_index: entry.character_index as usize,
                name,
                rows: model.rows() as usize,
                columns: model.columns() as usize,
                appearance: legacy_table_appearance_compatibility_read(editor, model_id)?,
                lock_state: table_lock_state_from_message(&message.data)?,
            },
        });
    }
    let table_roots = result
        .iter()
        .flat_map(|graph| {
            [
                graph.attachment_object_id,
                graph.info.drawable_object_id,
                graph.info.model_object_id,
            ]
        })
        .collect::<HashSet<_>>();
    for graph in &mut result {
        let mut excluded = table_roots.clone();
        excluded.remove(&graph.attachment_object_id);
        excluded.remove(&graph.info.drawable_object_id);
        excluded.remove(&graph.info.model_object_id);
        excluded.insert(editor.body_storage_id.get());
        expand_formula_contexts(editor.package(), &mut graph.formula_context_ids, &excluded)?;
    }
    result.sort_by_key(|graph| graph.info.anchor_character_index);
    Ok(result)
}

/// Check only declared anchors while streaming the text once. Memory scales
/// with attachment metadata rather than the complete UTF-16 body, including
/// when the source lists anchors out of order or repeats an offset.
fn retain_object_replacement_offsets(text: &str, offsets: &mut Vec<u32>) {
    offsets.sort_unstable();
    offsets.dedup();
    let mut units = text.encode_utf16();
    let mut consumed = 0u64;
    offsets.retain(|offset| {
        let offset = u64::from(*offset);
        let Some(skip) = offset
            .checked_sub(consumed)
            .and_then(|value| usize::try_from(value).ok())
        else {
            return false;
        };
        let unit = units.nth(skip);
        consumed = offset + 1;
        unit == Some(OBJECT_REPLACEMENT_CHARACTER)
    });
}

pub(crate) fn clone_body_table_attachment(
    source: &IWorkPackage,
    staged: &mut IWorkPackage,
    source_attachment_id: u64,
    source_drawable_id: u64,
    new_drawable_id: u64,
) -> Result<u64> {
    let new_attachment_id = next_object_identifier(staged)?;
    let attachment_archive_name = find_object_archive(source, source_attachment_id)?;
    let attachment_archive = source.archive(&attachment_archive_name)?;
    let attachment_object = attachment_archive
        .object(source_attachment_id)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages table attachment {source_attachment_id} is missing"
            ))
        })?;
    let remap = HashMap::from([
        (source_attachment_id, new_attachment_id),
        (source_drawable_id, new_drawable_id),
    ]);
    let cloned_attachment = clone_pages_drawable_graph_object(attachment_object, &remap)?;
    staged.update_archive(&attachment_archive_name, |archive| {
        archive
            .insert_object(cloned_attachment)
            .map_err(Error::from)
    })?;

    if let Some(component) = component_identifier_for_entry(source, &attachment_archive_name)? {
        if component_uuid_identifiers(source, component)?
            .is_some_and(|identifiers| identifiers.contains(&source_attachment_id))
        {
            add_component_object_uuids(staged, component, &[new_attachment_id])?;
        }
        let new_drawable_archive = find_object_archive(staged, new_drawable_id)?;
        if let Some(target_component) =
            component_identifier_for_entry(staged, &new_drawable_archive)?
            && target_component != component
        {
            add_component_external_reference(staged, component, target_component, new_drawable_id)?;
        }
    }
    Ok(new_attachment_id)
}

pub(crate) fn expand_formula_contexts(
    package: &IWorkPackage,
    contexts: &mut Vec<u64>,
    excluded: &HashSet<u64>,
) -> Result<()> {
    const CALCULATION_ENGINE_MESSAGE_TYPE: u32 = 4_000;
    const FORMULA_OWNER_MESSAGE_TYPE: u32 = 4_008;
    const CELL_RECORD_TILE_MESSAGE_TYPE: u32 = 4_009;

    contexts.retain(|identifier| !excluded.contains(identifier));
    let mut seen = contexts.iter().copied().collect::<HashSet<_>>();
    let mut cursor = 0usize;
    while cursor < contexts.len() {
        let identifier = contexts[cursor];
        cursor += 1;
        let Ok(archive_name) = find_object_archive(package, identifier) else {
            continue;
        };
        let archive = package.archive(&archive_name)?;
        let object = archive.object(identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Pages table context object {identifier} is missing"
            ))
        })?;
        if object.messages.iter().any(|message| {
            matches!(
                message.type_,
                CALCULATION_ENGINE_MESSAGE_TYPE
                    | FORMULA_OWNER_MESSAGE_TYPE
                    | CELL_RECORD_TILE_MESSAGE_TYPE
            )
        }) {
            continue;
        }
        for reference in object
            .archive_info
            .message_infos
            .iter()
            .flat_map(|message| {
                message.object_references.iter().copied().chain(
                    message
                        .field_infos
                        .iter()
                        .flat_map(|field| field.object_references.iter().copied()),
                )
            })
        {
            if !excluded.contains(&reference) && seen.insert(reference) {
                contexts.push(reference);
            }
        }
    }
    Ok(())
}

pub(crate) fn remove_table_object(
    package: &mut IWorkPackage,
    archive_name: &str,
    identifier: u64,
) -> Result<bool> {
    let mut archive = package.archive(archive_name)?;
    archive.remove_object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Pages table object {identifier} is missing"))
    })?;
    if archive.objects.is_empty() {
        package.remove_entry(archive_name).ok_or_else(|| {
            Error::InvalidFormat(format!("Pages table component {archive_name} is missing"))
        })?;
        Ok(true)
    } else {
        package.replace_archive(archive_name, &archive)?;
        Ok(false)
    }
}

#[cfg(test)]
mod tests {
    use super::retain_object_replacement_offsets;

    #[test]
    fn attachment_anchors_use_utf16_positions_and_preserve_out_of_order_queries() {
        let mut offsets = vec![4, 1, 2, 2, 0, u32::MAX, 3];
        retain_object_replacement_offsets("😀\u{fffc}A\u{fffc}", &mut offsets);
        assert_eq!(offsets, [2, 4]);
    }

    #[test]
    fn empty_and_out_of_range_attachment_anchors_are_not_validated() {
        let mut offsets = vec![0, 1, u32::MAX];
        retain_object_replacement_offsets("", &mut offsets);
        assert!(offsets.is_empty());
        retain_object_replacement_offsets("\u{fffc}", &mut offsets);
        assert!(offsets.is_empty());
    }
}
