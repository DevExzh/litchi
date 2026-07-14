//! Native-style duplication of populated Numbers sheets.

use super::*;

mod wire;

use wire::clone_empty_sheet_object;

const TABLE_POSITION_X_PATH: &[u32] = &[1, 1, 1, 1];
const TABLE_POSITION_Y_PATH: &[u32] = &[1, 1, 1, 2];
const SHEET_COPY_TEXT_BOX_OFFSET: f32 = 0.0;

#[derive(Debug)]
enum SheetDrawableClone {
    Table {
        model_id: u64,
        info_id: u64,
        name: String,
    },
    TextBox {
        drawable_id: u64,
        text: String,
    },
}

impl NumbersEditor {
    /// Duplicate a populated sheet immediately after its source.
    ///
    /// Sheet settings and unknown wire fields are retained. Populated tables,
    /// local formula dependency graphs, and ordinary text boxes receive fresh
    /// object identities and independent writable storage. Unsupported drawable
    /// kinds and cross-table formula edges are rejected transactionally.
    pub fn duplicate_sheet(&mut self, sheet_id: u64) -> Result<NumbersSheetInfo> {
        let sheets = self.sheets()?;
        let source = sheets
            .iter()
            .find(|sheet| sheet.object_id == sheet_id)
            .ok_or_else(|| Error::ParseError(format!("Numbers sheet {sheet_id} not found")))?;
        let existing_names = sheets
            .iter()
            .map(|sheet| sheet.name.as_str())
            .collect::<HashSet<_>>();
        let name = duplicate_sheet_name(&source.name, &existing_names)?;
        let (archive_name, message_index, sheet) = numbers_sheet(&self.package, sheet_id)?;
        let drawables = classify_sheet_drawables(self, sheet_id, &sheet)?;
        for drawable in &drawables {
            if let SheetDrawableClone::Table {
                model_id, info_id, ..
            } = drawable
                && !table_formula_graph_is_self_contained(self.package(), *info_id)?
            {
                return Err(Error::ParseError(format!(
                    "Cannot duplicate Numbers sheet {sheet_id}: table {model_id} has cross-table formula dependencies"
                )));
            }
        }

        let new_sheet_id = next_object_identifier(&self.package)?;
        let cloned_sheet = {
            let archive = self.package.archive(&archive_name)?;
            let source_object = archive.object(sheet_id).ok_or_else(|| {
                Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing"))
            })?;
            clone_empty_sheet_object(
                source_object,
                message_index,
                new_sheet_id,
                &name,
                &sheet.drawable_infos,
            )?
        };

        let mut staged = self.package.clone();
        staged.update_archive(&archive_name, |archive| archive.insert_object(cloned_sheet))?;
        update_numbers_document(&mut staged, |document| {
            let matches = document
                .sheets
                .iter()
                .enumerate()
                .filter(|(_, reference)| reference.identifier == sheet_id)
                .map(|(index, _)| index)
                .collect::<Vec<_>>();
            let [source_index] = matches.as_slice() else {
                return Err(Error::InvalidFormat(format!(
                    "Numbers root must reference sheet {sheet_id} exactly once"
                )));
            };
            document.sheets.insert(
                *source_index + 1,
                tsp::Reference {
                    identifier: new_sheet_id,
                    ..Default::default()
                },
            );
            Ok(())
        })?;
        set_package_last_object_identifier(&mut staged, new_sheet_id)?;
        register_sheet_uuid_if_needed(
            &mut staged,
            &self.package,
            &archive_name,
            sheet_id,
            new_sheet_id,
        )?;

        let mut working = Self::from_package(staged)?;
        let mut cloned_drawable_ids = Vec::with_capacity(drawables.len());
        for drawable in drawables {
            match drawable {
                SheetDrawableClone::Table { model_id, name, .. } => {
                    let cloned = working.duplicate_table(model_id)?;
                    working.move_table(cloned.object_id, new_sheet_id)?;
                    working.rename_table(cloned.object_id, &name)?;
                    restore_table_geometry(&mut working.package, model_id, cloned.object_id)?;
                    cloned_drawable_ids
                        .push(find_table_owner(working.package(), cloned.object_id)?.table_info_id);
                },
                SheetDrawableClone::TextBox { drawable_id, text } => {
                    let cloned = working.duplicate_text_box_to_sheet(
                        sheet_id,
                        drawable_id,
                        new_sheet_id,
                        &text,
                        SHEET_COPY_TEXT_BOX_OFFSET,
                    )?;
                    cloned_drawable_ids.push(cloned.drawable_object_id);
                },
            }
        }

        let (_, _, verified_sheet) = numbers_sheet(working.package(), new_sheet_id)?;
        let verified_drawables = verified_sheet
            .drawable_infos
            .iter()
            .map(|reference| reference.identifier)
            .collect::<Vec<_>>();
        if verified_sheet.name != name || verified_drawables != cloned_drawable_ids {
            return Err(Error::InvalidFormat(
                "Numbers sheet duplication failed structural validation".to_owned(),
            ));
        }
        let created = working
            .sheets()?
            .into_iter()
            .find(|sheet| sheet.object_id == new_sheet_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers duplicated sheet is unreachable".to_owned())
            })?;
        self.package = working.package;
        Ok(created)
    }
}

fn classify_sheet_drawables(
    editor: &NumbersEditor,
    sheet_id: u64,
    sheet: &tn::SheetArchive,
) -> Result<Vec<SheetDrawableClone>> {
    let package = editor.package();
    let mut tables = HashMap::new();
    for descriptor in table_models(package)? {
        let owner = find_table_owner(package, descriptor.object_id)?;
        if owner.sheet_id == sheet_id {
            tables.insert(
                owner.table_info_id,
                (descriptor.object_id, descriptor.model.table_name),
            );
        }
    }
    let text_boxes = editor
        .sheet_text_boxes(sheet_id)?
        .into_iter()
        .map(|text_box| (text_box.drawable_object_id, text_box.storage.text))
        .collect::<HashMap<_, _>>();

    sheet
        .drawable_infos
        .iter()
        .map(|reference| {
            if let Some((model_id, name)) = tables.remove(&reference.identifier) {
                return Ok(SheetDrawableClone::Table {
                    model_id,
                    info_id: reference.identifier,
                    name,
                });
            }
            if let Some(text) = text_boxes.get(&reference.identifier) {
                return Ok(SheetDrawableClone::TextBox {
                    drawable_id: reference.identifier,
                    text: text.clone(),
                });
            }
            Err(Error::ParseError(format!(
                "Cannot duplicate Numbers sheet {sheet_id}: drawable {} is not a supported table or ordinary text box",
                reference.identifier
            )))
        })
        .collect()
}

fn duplicate_sheet_name(source: &str, existing: &HashSet<&str>) -> Result<String> {
    validate_name(source, "sheet")?;
    for suffix in 1u32..=u32::MAX {
        let candidate = format!("{source}-{suffix}");
        if !existing.contains(candidate.as_str()) {
            return Ok(candidate);
        }
    }
    Err(Error::ParseError(
        "Unable to allocate a unique Numbers sheet name".to_owned(),
    ))
}

fn register_sheet_uuid_if_needed(
    staged: &mut IWorkPackage,
    source: &IWorkPackage,
    archive_name: &str,
    source_sheet_id: u64,
    new_sheet_id: u64,
) -> Result<()> {
    let Some(component_id) = component_identifier_for_entry(source, archive_name)? else {
        return Ok(());
    };
    if component_uuid_identifiers(source, component_id)?
        .is_some_and(|mapped| mapped.contains(&source_sheet_id))
    {
        add_component_object_uuids(staged, component_id, &[new_sheet_id])?;
    }
    Ok(())
}

fn restore_table_geometry(
    package: &mut IWorkPackage,
    source_table_id: u64,
    cloned_table_id: u64,
) -> Result<()> {
    let source_owner = find_table_owner(package, source_table_id)?;
    let cloned_owner = find_table_owner(package, cloned_table_id)?;
    let locations = object_locations(package)?;
    let source_archive_name = locations.get(&source_owner.table_info_id).ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Numbers table info {} is missing",
            source_owner.table_info_id
        ))
    })?;
    let source_archive = package.archive(source_archive_name)?;
    let source_object = source_archive
        .object(source_owner.table_info_id)
        .ok_or_else(|| Error::InvalidFormat("Numbers source table info is missing".to_owned()))?;
    let (_, source_info) = decode_table_info(source_object)?;
    let source_position = source_info
        .super_
        .geometry
        .and_then(|geometry| geometry.position);

    let clone_archive_name = locations
        .get(&cloned_owner.table_info_id)
        .ok_or_else(|| Error::InvalidFormat("Numbers cloned table info is missing".to_owned()))?
        .to_owned();
    package.update_archive(&clone_archive_name, |archive| {
        let object = archive
            .object_mut(cloned_owner.table_info_id)
            .ok_or_else(|| {
                Error::InvalidFormat("Numbers cloned table info is missing".to_owned())
            })?;
        let (message_index, cloned_info) = decode_table_info(object)?;
        let cloned_position = cloned_info
            .super_
            .geometry
            .as_ref()
            .and_then(|geometry| geometry.position.as_ref());
        match (&source_position, cloned_position) {
            (None, None) => return Ok(()),
            (Some(_), None) | (None, Some(_)) => {
                return Err(Error::InvalidFormat(
                    "Numbers table clone changed positioned-geometry presence".to_owned(),
                ));
            },
            (Some(source_position), Some(_)) => {
                let message_type = object.messages[message_index].type_;
                let data = patch_nested_fixed32_field(
                    &object.messages[message_index].data,
                    TABLE_POSITION_X_PATH,
                    true,
                    Some(source_position.x.to_bits()),
                )?;
                let data = patch_nested_fixed32_field(
                    &data,
                    TABLE_POSITION_Y_PATH,
                    true,
                    Some(source_position.y.to_bits()),
                )?;
                let verified = tst::TableInfoArchive::decode(data.as_slice())?;
                let verified_position = verified
                    .super_
                    .geometry
                    .and_then(|geometry| geometry.position);
                if verified_position.as_ref() != Some(source_position) {
                    return Err(Error::InvalidFormat(
                        "Numbers sheet clone failed to preserve table position".to_owned(),
                    ));
                }
                object.replace_message(
                    message_index,
                    RawMessage {
                        type_: message_type,
                        data,
                    },
                )?;
            },
        }
        Ok(())
    })
}
