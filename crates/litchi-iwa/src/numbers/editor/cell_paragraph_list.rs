//! Native rich-text promotion and paragraph-list CRUD for table cells.

use super::*;

const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const RICH_TEXT_PAYLOAD_MESSAGE_TYPE: u32 = 6_218;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const STANDARD_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];

pub(super) fn paragraph_list(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<ParagraphList> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        return Ok(ParagraphList::None);
    };
    crate::text::paragraph_list_in_storage(package, storage_id)
}

pub(super) fn set_paragraph_list(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    list: ParagraphList,
) -> Result<()> {
    if paragraph_list(package, table_id, row, column)? == list {
        return Ok(());
    }
    let mut staged = package.clone();
    let storage_id = ensure_storage(&mut staged, table_id, row, column)?;
    let mut text = IWorkTextEditor::from_package(staged);
    text.set_paragraph_list(storage_id, list)?;
    staged = text.into_package();
    if paragraph_list(&staged, table_id, row, column)? != list {
        return Err(Error::InvalidFormat(
            "iWork table-cell paragraph-list update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(super) fn reset_paragraph_list(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    if paragraph_list(package, table_id, row, column)? == ParagraphList::None {
        return Ok(false);
    }
    set_paragraph_list(package, table_id, row, column, ParagraphList::None)?;
    Ok(true)
}

pub(super) fn paragraph_lists(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Vec<ParagraphListPlacement>> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        return Ok(vec![ParagraphListPlacement::new(
            ParagraphStart::ZERO,
            ParagraphList::None,
        )]);
    };
    crate::text::paragraph_lists_in_storage(package, storage_id)
}

pub(super) fn set_paragraph_lists(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    placements: &[ParagraphListPlacement],
) -> Result<()> {
    let mut staged = package.clone();
    let storage_id = ensure_storage(&mut staged, table_id, row, column)?;
    let mut text = IWorkTextEditor::from_package(staged);
    text.set_paragraph_lists(storage_id, placements)?;
    staged = text.into_package();
    let expected = IWorkTextEditor::from_package(staged.clone()).paragraph_lists(storage_id)?;
    if paragraph_lists(&staged, table_id, row, column)? != expected {
        return Err(Error::InvalidFormat(
            "iWork table-cell paragraph-list placements failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(super) fn paragraph_list_levels(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Vec<ParagraphListLevelPlacement>> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        return Ok(vec![ParagraphListLevelPlacement::new(
            ParagraphStart::ZERO,
            ParagraphListLevel::ZERO,
        )]);
    };
    crate::text::paragraph_list_levels_in_storage(package, storage_id)
}

pub(super) fn paragraph_list_level(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<ParagraphListLevel> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Ok(ParagraphListLevel::ZERO);
    };
    crate::text::paragraph_list_level_in_storage(package, storage_id, paragraph)
}

pub(super) fn set_paragraph_list_level(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
    level: ParagraphListLevel,
) -> Result<()> {
    if existing_storage_id(package, table_id, row, column)?.is_none()
        && level == ParagraphListLevel::ZERO
    {
        return require_plain_cell_paragraph_start(package, table_id, row, column, paragraph);
    }
    let mut staged = package.clone();
    let storage_id = ensure_storage(&mut staged, table_id, row, column)?;
    let mut text = IWorkTextEditor::from_package(staged);
    text.set_paragraph_list_level(storage_id, paragraph, level)?;
    staged = text.into_package();
    if crate::text::paragraph_list_level_in_storage(&staged, storage_id, paragraph)? != level {
        return Err(Error::InvalidFormat(
            "iWork table-cell paragraph list-level update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(super) fn reset_paragraph_list_level(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<bool> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Ok(false);
    };
    if crate::text::paragraph_list_level_in_storage(package, storage_id, paragraph)?
        == ParagraphListLevel::ZERO
    {
        return Ok(false);
    }
    let mut staged = package.clone();
    let storage_id = ensure_storage(&mut staged, table_id, row, column)?;
    let mut text = IWorkTextEditor::from_package(staged);
    let changed = text.reset_paragraph_list_level(storage_id, paragraph)?;
    staged = text.into_package();
    let actual = crate::text::paragraph_list_level_in_storage(&staged, storage_id, paragraph)?;
    if !changed || actual != ParagraphListLevel::ZERO {
        return Err(Error::InvalidFormat(format!(
            "iWork table-cell paragraph list-level reset failed validation: changed={changed}, actual={actual:?}"
        )));
    }
    *package = staged;
    Ok(true)
}

pub(super) fn paragraph_list_numbering(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<ParagraphListNumbering> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Ok(ParagraphListNumbering::Continue);
    };
    crate::text::paragraph_list_numbering_in_storage(package, storage_id, paragraph)
}

pub(super) fn set_paragraph_list_numbering(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
    numbering: ParagraphListNumbering,
) -> Result<()> {
    if existing_storage_id(package, table_id, row, column)?.is_none()
        && numbering == ParagraphListNumbering::Continue
    {
        return require_plain_cell_paragraph_start(package, table_id, row, column, paragraph);
    }
    let mut staged = package.clone();
    let storage_id = ensure_storage(&mut staged, table_id, row, column)?;
    let mut text = IWorkTextEditor::from_package(staged);
    text.set_paragraph_list_numbering(storage_id, paragraph, numbering)?;
    staged = text.into_package();
    if crate::text::paragraph_list_numbering_in_storage(&staged, storage_id, paragraph)?
        != numbering
    {
        return Err(Error::InvalidFormat(
            "iWork table-cell paragraph list-numbering update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(super) fn paragraph_list_number_format(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<crate::text::ParagraphListNumberFormat> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Err(Error::InvalidFormat(
            "plain iWork table cells are not numbered lists".to_owned(),
        ));
    };
    crate::text::paragraph_list_number_format_in_storage(package, storage_id, paragraph)
}

pub(super) fn set_paragraph_list_number_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
    format: crate::text::ParagraphListNumberFormat,
) -> Result<()> {
    let mut staged = package.clone();
    let storage_id = ensure_storage(&mut staged, table_id, row, column)?;
    let mut text = IWorkTextEditor::from_package(staged);
    text.set_paragraph_list_number_format(storage_id, paragraph, format)?;
    staged = text.into_package();
    if crate::text::paragraph_list_number_format_in_storage(&staged, storage_id, paragraph)?
        != format
    {
        return Err(Error::InvalidFormat(
            "iWork table-cell paragraph list-number format update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(super) fn reset_paragraph_list_number_format(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<bool> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Ok(false);
    };
    let mut text = IWorkTextEditor::from_package(package.clone());
    let changed = text.reset_paragraph_list_number_format(storage_id, paragraph)?;
    if changed {
        *package = text.into_package();
    }
    Ok(changed)
}

pub(super) fn paragraph_list_bullet(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<ParagraphListBullet> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Err(Error::InvalidFormat(
            "plain iWork table cells are not text-bullet lists".to_owned(),
        ));
    };
    crate::text::paragraph_list_bullet_in_storage(package, storage_id, paragraph)
}

pub(super) fn set_paragraph_list_bullet(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
    bullet: &ParagraphListBullet,
) -> Result<()> {
    let mut staged = package.clone();
    let storage_id = ensure_storage(&mut staged, table_id, row, column)?;
    let mut text = IWorkTextEditor::from_package(staged);
    text.set_paragraph_list_bullet(storage_id, paragraph, bullet)?;
    staged = text.into_package();
    if crate::text::paragraph_list_bullet_in_storage(&staged, storage_id, paragraph)? != *bullet {
        return Err(Error::InvalidFormat(
            "iWork table-cell paragraph text-bullet update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(super) fn reset_paragraph_list_bullet(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<bool> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Ok(false);
    };
    let mut text = IWorkTextEditor::from_package(package.clone());
    let changed = text.reset_paragraph_list_bullet(storage_id, paragraph)?;
    if changed {
        *package = text.into_package();
    }
    Ok(changed)
}

pub(super) fn paragraph_list_bullet_geometry(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<crate::text::ParagraphListBulletGeometry> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Err(Error::InvalidFormat(
            "plain iWork table cells are not text-bullet lists".to_owned(),
        ));
    };
    crate::text::paragraph_list_bullet_geometry_in_storage(package, storage_id, paragraph)
}

pub(super) fn set_paragraph_list_bullet_geometry(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
    geometry: crate::text::ParagraphListBulletGeometry,
) -> Result<()> {
    let mut staged = package.clone();
    let storage_id = ensure_storage(&mut staged, table_id, row, column)?;
    let mut text = IWorkTextEditor::from_package(staged);
    text.set_paragraph_list_bullet_geometry(storage_id, paragraph, geometry)?;
    staged = text.into_package();
    if crate::text::paragraph_list_bullet_geometry_in_storage(&staged, storage_id, paragraph)?
        != geometry
    {
        return Err(Error::InvalidFormat(
            "iWork table-cell paragraph text-bullet geometry update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(super) fn reset_paragraph_list_bullet_geometry(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<bool> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Ok(false);
    };
    let mut text = IWorkTextEditor::from_package(package.clone());
    let changed = text.reset_paragraph_list_bullet_geometry(storage_id, paragraph)?;
    if changed {
        *package = text.into_package();
    }
    Ok(changed)
}

pub(super) fn paragraph_list_indentation(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<crate::text::ParagraphListIndentation> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Err(Error::InvalidFormat(
            "plain iWork table cells do not have list indentation".to_owned(),
        ));
    };
    crate::text::paragraph_list_indentation_in_storage(package, storage_id, paragraph)
}

pub(super) fn set_paragraph_list_indentation(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
    indentation: crate::text::ParagraphListIndentation,
) -> Result<()> {
    let mut staged = package.clone();
    let storage_id = ensure_storage(&mut staged, table_id, row, column)?;
    let mut text = IWorkTextEditor::from_package(staged);
    text.set_paragraph_list_indentation(storage_id, paragraph, indentation)?;
    staged = text.into_package();
    if crate::text::paragraph_list_indentation_in_storage(&staged, storage_id, paragraph)?
        != indentation
    {
        return Err(Error::InvalidFormat(
            "iWork table-cell paragraph list-indentation update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(super) fn reset_paragraph_list_indentation(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<bool> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Ok(false);
    };
    let mut text = IWorkTextEditor::from_package(package.clone());
    let changed = text.reset_paragraph_list_indentation(storage_id, paragraph)?;
    if changed {
        *package = text.into_package();
    }
    Ok(changed)
}

pub(super) fn paragraph_list_label_color(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<crate::text::ParagraphListLabelColor> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Err(Error::InvalidFormat(
            "plain iWork table cells do not have list-label colors".to_owned(),
        ));
    };
    crate::text::paragraph_list_label_color_in_storage(package, storage_id, paragraph)
}

pub(super) fn set_paragraph_list_label_color(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
    color: crate::text::ParagraphListLabelColor,
) -> Result<()> {
    let mut staged = package.clone();
    let storage_id = ensure_storage(&mut staged, table_id, row, column)?;
    let mut text = IWorkTextEditor::from_package(staged);
    text.set_paragraph_list_label_color(storage_id, paragraph, color)?;
    staged = text.into_package();
    if crate::text::paragraph_list_label_color_in_storage(&staged, storage_id, paragraph)? != color
    {
        return Err(Error::InvalidFormat(
            "iWork table-cell paragraph list-label color update failed validation".to_owned(),
        ));
    }
    *package = staged;
    Ok(())
}

pub(super) fn reset_paragraph_list_label_color(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<bool> {
    let Some(storage_id) = existing_storage_id(package, table_id, row, column)? else {
        require_plain_cell_paragraph_start(package, table_id, row, column, paragraph)?;
        return Ok(false);
    };
    let mut text = IWorkTextEditor::from_package(package.clone());
    let changed = text.reset_paragraph_list_label_color(storage_id, paragraph)?;
    if changed {
        *package = text.into_package();
    }
    Ok(changed)
}

fn existing_storage_id(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<Option<u64>> {
    let location = model::locate_attached_cell(package, table_id, row, column)?;
    let Some(data) = storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    else {
        return Ok(None);
    };
    match BncCell::parse(&data)?.stored_value() {
        StoredValue::Empty | StoredValue::Text(_) => Ok(None),
        StoredValue::RichText(key) => Ok(Some(
            storage::rich_text_entry_location(
                package,
                &location.object_locations,
                &location.descriptor.model,
                key,
            )?
            .storage_id,
        )),
        _ => Err(Error::ParseError(
            "Paragraph lists require an empty or textual iWork table cell".to_owned(),
        )),
    }
}

fn require_plain_cell_paragraph_start(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    paragraph: ParagraphStart,
) -> Result<()> {
    let location = model::locate_attached_cell(package, table_id, row, column)?;
    let data = storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?;
    let stored = data
        .as_deref()
        .map(BncCell::parse)
        .transpose()?
        .map_or(StoredValue::Empty, |cell| cell.stored_value());
    let text = plain_cell_text(package, &location, stored)?;
    if paragraph_starts(&text)?
        .binary_search(&paragraph.utf16_index())
        .is_ok()
    {
        return Ok(());
    }
    Err(Error::InvalidFormat(format!(
        "UTF-16 index {} is not a paragraph start in an iWork table cell",
        paragraph.utf16_index()
    )))
}

fn plain_cell_text(
    package: &IWorkPackage,
    location: &model::CellLocation,
    stored: StoredValue,
) -> Result<String> {
    let StoredValue::Text(key) = stored else {
        return match stored {
            StoredValue::Empty => Ok(String::new()),
            _ => Err(Error::ParseError(
                "Paragraph lists require an empty or textual iWork table cell".to_owned(),
            )),
        };
    };
    let string_table = location
        .descriptor
        .model
        .base_data_store
        .string_table
        .identifier;
    storage::resolve_table_string_values(
        package,
        &location.object_locations,
        string_table,
        &HashSet::from([key]),
    )?
    .remove(&key)
    .ok_or_else(|| Error::InvalidFormat(format!("Numbers string table has no entry {key}")))
}

fn ensure_storage(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<u64> {
    table_sparse_storage::ensure_attached_cell_storage(package, table_id, row, column)?;
    let location = model::locate_attached_cell(package, table_id, row, column)?;
    let data = storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?;
    let stored = data
        .as_deref()
        .map(BncCell::parse)
        .transpose()?
        .map_or(StoredValue::Empty, |cell| cell.stored_value());
    match stored {
        StoredValue::RichText(key) => {
            let entry = storage::rich_text_entry_location(
                package,
                &location.object_locations,
                &location.descriptor.model,
                key,
            )?;
            let text = package
                .archive(&entry.storage_archive)?
                .object(entry.storage_id)
                .and_then(|object| object.messages.first())
                .map(|message| tswp::StorageArchive::decode(message.data.as_slice()))
                .transpose()?
                .and_then(|storage| storage.text.into_iter().next())
                .ok_or_else(|| {
                    Error::InvalidFormat(format!(
                        "iWork rich-text storage {} has no text",
                        entry.storage_id
                    ))
                })?;
            let new_key = storage::set_rich_text(
                package,
                &location.object_locations,
                &location.descriptor.model,
                key,
                row,
                column,
                &text,
            )?;
            if new_key == key {
                return Ok(entry.storage_id);
            }
            storage::set_encoded_cell_value(
                package,
                table_id,
                row,
                column,
                model::EncodedValue::RichText(new_key),
            )?;
            let locations = storage::object_locations(package)?;
            let descriptor = model::attached_table_descriptor(package, table_id)?;
            Ok(
                storage::rich_text_entry_location(package, &locations, &descriptor.model, new_key)?
                    .storage_id,
            )
        },
        StoredValue::Empty | StoredValue::Text(_) => {
            promote_plain_cell(package, table_id, row, column, stored, &location)
        },
        _ => Err(Error::ParseError(
            "Paragraph lists require an empty or textual iWork table cell".to_owned(),
        )),
    }
}

fn promote_plain_cell(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    stored: StoredValue,
    location: &model::CellLocation,
) -> Result<u64> {
    let old_string_key = match stored {
        StoredValue::Text(key) => Some(key),
        StoredValue::Empty => None,
        _ => unreachable!("plain-cell promotion receives only plain stored values"),
    };
    let text = plain_cell_text(package, location, stored)?;
    let (paragraph_style_id, stylesheet_id) =
        cell_paragraph_style::style_context(package, table_id, row, column)?;
    let list_style_id = crate::text::preset_style_id(package, stylesheet_id, ParagraphList::None)?
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork stylesheet {stylesheet_id} has no canonical None list preset"
            ))
        })?;

    let archive_name = location
        .object_locations
        .get(&table_id)
        .cloned()
        .ok_or_else(|| Error::InvalidFormat(format!("iWork table model {table_id} is missing")))?;
    let mut next_id = next_object_identifier(package)?;
    let rich_text_table_id = location
        .descriptor
        .model
        .base_data_store
        .rich_text_table
        .as_ref()
        .map(|reference| reference.identifier);
    let list_id = match rich_text_table_id {
        Some(identifier) => identifier,
        None => take_id(&mut next_id)?,
    };
    let storage_id = take_id(&mut next_id)?;
    let payload_id = take_id(&mut next_id)?;

    if rich_text_table_id.is_none() {
        let list = TableDataList {
            list_type: tst::table_data_list::ListType::RichTextPayload as i32,
            next_list_id: 1,
            entries: Vec::new(),
            segments: Vec::new(),
            is_new_for_bnc: Some(true),
        };
        let mut object = ArchiveObject::new(
            list_id,
            vec![RawMessage {
                type_: TABLE_DATA_LIST_MESSAGE_TYPE,
                data: list.encode_to_vec(),
            }],
        )?;
        object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
        package.update_archive(&archive_name, |archive| archive.insert_object(object))?;
        attach_rich_text_table(package, &archive_name, table_id, list_id)?;
    }

    let storage_archive = cell_storage(&text, stylesheet_id, paragraph_style_id, list_style_id)?;
    let storage_references = crate::text::editor::storage_object_references(&storage_archive);
    let mut storage_object = ArchiveObject::new(
        storage_id,
        vec![RawMessage {
            type_: STORAGE_MESSAGE_TYPE,
            data: storage_archive.encode_to_vec(),
        }],
    )?;
    storage_object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
    storage_object.archive_info.message_infos[0].object_references = storage_references;

    let payload = tst::RichTextPayloadArchive {
        storage: tsp::Reference {
            identifier: storage_id,
            ..Default::default()
        },
        range: None,
        cellid: storage::rich_text_cell_id(row, column)?,
    };
    let mut payload_object = ArchiveObject::new(
        payload_id,
        vec![RawMessage {
            type_: RICH_TEXT_PAYLOAD_MESSAGE_TYPE,
            data: payload.encode_to_vec(),
        }],
    )?;
    payload_object.archive_info.message_infos[0].versions = STANDARD_MESSAGE_VERSION.to_vec();
    payload_object.archive_info.message_infos[0]
        .object_references
        .push(storage_id);
    package.update_archive(&archive_name, |archive| {
        archive.insert_object(storage_object)?;
        archive.insert_object(payload_object)
    })?;

    let locations = storage::object_locations(package)?;
    let resolved = storage::resolve_table_data_list(
        package,
        &locations,
        list_id,
        tst::table_data_list::ListType::RichTextPayload,
    )?;
    let key = storage::next_table_data_list_key(&resolved.list, &resolved.entries)?;
    package.update_archive(&archive_name, |archive| {
        let object = archive.object_mut(list_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork rich-text table {list_id} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_DATA_LIST_MESSAGE_TYPE)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("iWork rich-text table {list_id} has no payload"))
            })?;
        let previous = TableDataList::decode(object.messages[message_index].data.as_slice())?;
        let mut current = previous.clone();
        current.next_list_id = key
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("iWork rich-text key overflow".to_owned()))?;
        current.entries.push(tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            rich_text_payload: Some(tsp::Reference {
                identifier: payload_id,
                ..Default::default()
            }),
            ..Default::default()
        });
        let data = storage::rewrite_table_data_list_wire(
            object.messages[message_index].data.as_slice(),
            &previous,
            &current,
        )?;
        let message_type = object.messages[message_index].type_;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        object.archive_info.message_infos[message_index]
            .object_references
            .push(payload_id);
        Ok(())
    })?;
    storage::set_encoded_cell_value(
        package,
        table_id,
        row,
        column,
        model::EncodedValue::RichText(key),
    )?;
    if let Some(old_key) = old_string_key {
        let string_table = location
            .descriptor
            .model
            .base_data_store
            .string_table
            .identifier;
        storage::update_string_table(
            package,
            &location.object_locations,
            string_table,
            Some(old_key),
            None,
        )?;
    }
    set_package_last_object_identifier(package, payload_id)?;
    Ok(storage_id)
}

fn attach_rich_text_table(
    package: &mut IWorkPackage,
    archive_name: &str,
    table_id: u64,
    list_id: u64,
) -> Result<()> {
    package.update_archive(archive_name, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork table model {table_id} is missing"))
        })?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == 6_000 || message.type_ == 6_001)
            .ok_or_else(|| {
                Error::InvalidFormat(format!("iWork table model {table_id} has no payload"))
            })?;
        let previous = TableModelArchive::decode(object.messages[message_index].data.as_slice())?;
        let mut current = previous.clone();
        current.base_data_store.rich_text_table = Some(tsp::Reference {
            identifier: list_id,
            ..Default::default()
        });
        let data = storage::rewrite_table_model_rich_text_table_wire(
            object.messages[message_index].data.as_slice(),
            &previous,
            &current,
        )?;
        let message_type = object.messages[message_index].type_;
        object.replace_message(
            message_index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        object.archive_info.message_infos[message_index]
            .object_references
            .push(list_id);
        Ok(())
    })
}

fn cell_storage(
    text: &str,
    stylesheet_id: u64,
    paragraph_style_id: u64,
    list_style_id: u64,
) -> Result<tswp::StorageArchive> {
    let object_table = |identifier| tswp::ObjectAttributeTable {
        entries: vec![tswp::object_attribute_table::ObjectAttribute {
            character_index: 0,
            object: Some(tsp::Reference {
                identifier,
                ..Default::default()
            }),
        }],
    };
    let paragraph_starts = paragraph_starts(text)?;
    let para_data = || tswp::ParaDataAttributeTable {
        entries: paragraph_starts
            .iter()
            .copied()
            .map(
                |character_index| tswp::para_data_attribute_table::ParaDataAttribute {
                    character_index,
                    first: 0,
                    second: 0,
                },
            )
            .collect(),
    };
    Ok(tswp::StorageArchive {
        kind: Some(tswp::storage_archive::KindType::Cell as i32),
        style_sheet: Some(tsp::Reference {
            identifier: stylesheet_id,
            ..Default::default()
        }),
        text: vec![text.to_owned()],
        in_document: Some(true),
        table_para_style: Some(object_table(paragraph_style_id)),
        table_para_data: Some(para_data()),
        table_list_style: Some(object_table(list_style_id)),
        table_para_starts: Some(para_data()),
        table_para_bidi: Some(para_data()),
        table_drop_cap_style: Some(tswp::ObjectAttributeTable {
            entries: vec![tswp::object_attribute_table::ObjectAttribute {
                character_index: 0,
                object: None,
            }],
        }),
        ..Default::default()
    })
}

fn paragraph_starts(text: &str) -> Result<Vec<u32>> {
    let mut starts = vec![0];
    let mut utf16_index = 0_u32;
    for character in text.chars() {
        utf16_index = utf16_index
            .checked_add(character.len_utf16() as u32)
            .ok_or_else(|| Error::ParseError("iWork cell text exceeds UTF-16 limits".to_owned()))?;
        if character == '\n' {
            starts.push(utf16_index);
        }
    }
    Ok(starts)
}

fn take_id(next: &mut u64) -> Result<u64> {
    let identifier = *next;
    *next = next
        .checked_add(1)
        .ok_or_else(|| Error::ParseError("iWork object identifier overflow".to_owned()))?;
    Ok(identifier)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::numbers::NumbersDocumentBuilder;
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize};
    use crate::text::{
        ParagraphListBulletBaselineOffset, ParagraphListBulletGeometry, ParagraphListBulletScale,
        ParagraphListIndentation, ParagraphListLabelIndent, ParagraphListNumberFormat,
        ParagraphListNumberPunctuation, ParagraphListNumberSequence, ParagraphListTextGap,
        TextPointSize,
    };

    const ROW: usize = 1;
    const COLUMN: usize = 1;
    const TEXT: &str = "First paragraph\nSecond paragraph";
    const MIXED_TEXT: &str = "😀 first\nSecond\nThird";

    fn mixed_lists() -> Vec<ParagraphListPlacement> {
        vec![
            ParagraphListPlacement::new(ParagraphStart::ZERO, ParagraphList::None),
            ParagraphListPlacement::new(
                ParagraphStart::from_utf16_index(9).unwrap(),
                ParagraphList::Bullet,
            ),
            ParagraphListPlacement::new(
                ParagraphStart::from_utf16_index(16).unwrap(),
                ParagraphList::Numbered,
            ),
        ]
    }

    fn nested_second_paragraph() -> Vec<ParagraphListLevelPlacement> {
        vec![
            ParagraphListLevelPlacement::new(ParagraphStart::ZERO, ParagraphListLevel::ZERO),
            ParagraphListLevelPlacement::new(
                ParagraphStart::from_utf16_index(9).unwrap(),
                ParagraphListLevel::ONE,
            ),
            ParagraphListLevelPlacement::new(
                ParagraphStart::from_utf16_index(16).unwrap(),
                ParagraphListLevel::ZERO,
            ),
        ]
    }

    #[test]
    fn scratch_documents_promote_plain_cells_and_roundtrip_lists() {
        let mut numbers = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let numbers_table = numbers.tables().unwrap()[0].object_id;
        numbers
            .set_cell(numbers_table, ROW, COLUMN, CellValue::Text(TEXT.to_owned()))
            .unwrap();
        assert_eq!(
            numbers
                .table_cell_paragraph_list(numbers_table, ROW, COLUMN)
                .unwrap(),
            ParagraphList::None
        );
        numbers
            .set_table_cell_paragraph_list(numbers_table, ROW, COLUMN, ParagraphList::Bullet)
            .unwrap();
        let mut numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
        assert_eq!(
            numbers
                .table_cell_paragraph_list(numbers_table, ROW, COLUMN)
                .unwrap(),
            ParagraphList::Bullet
        );
        assert!(
            numbers
                .reset_table_cell_paragraph_list(numbers_table, ROW, COLUMN)
                .unwrap()
        );
        assert!(
            !numbers
                .reset_table_cell_paragraph_list(numbers_table, ROW, COLUMN)
                .unwrap()
        );

        let mut pages = PagesDocumentBuilder::new()
            .body_table("Lists", 3, 3)
            .build()
            .unwrap();
        let pages_table = pages.tables().unwrap()[0].model_object_id;
        pages
            .set_table_cell(pages_table, ROW, COLUMN, CellValue::Text(TEXT.to_owned()))
            .unwrap();
        pages
            .set_table_cell_paragraph_list(pages_table, ROW, COLUMN, ParagraphList::Numbered)
            .unwrap();
        let pages = crate::pages::PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert_eq!(
            pages
                .table_cell_paragraph_list(pages_table, ROW, COLUMN)
                .unwrap(),
            ParagraphList::Numbered
        );

        let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
        let table = keynote
            .add_slide_table(
                0,
                "Lists",
                3,
                3,
                DrawablePoint { x: 100.0, y: 100.0 },
                DrawableSize {
                    width: 600.0,
                    height: 300.0,
                },
            )
            .unwrap();
        keynote
            .set_slide_table_cell(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                CellValue::Text(TEXT.to_owned()),
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_list(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                ParagraphList::Bullet,
            )
            .unwrap();
        let keynote =
            crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert_eq!(
            keynote
                .slide_table_cell_paragraph_list(0, table.model_object_id, ROW, COLUMN,)
                .unwrap(),
            ParagraphList::Bullet
        );
    }

    #[test]
    #[allow(deprecated)]
    fn duplicated_rich_text_cells_are_list_copy_on_write() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let source = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(source, ROW, COLUMN, CellValue::Text(TEXT.to_owned()))
            .unwrap();
        editor
            .set_table_cell_paragraph_list(source, ROW, COLUMN, ParagraphList::Bullet)
            .unwrap();
        let duplicate = editor.duplicate_table(source).unwrap().object_id;
        editor
            .set_table_cell_paragraph_list(duplicate, ROW, COLUMN, ParagraphList::Numbered)
            .unwrap();
        assert_eq!(
            editor
                .table_cell_paragraph_list(source, ROW, COLUMN)
                .unwrap(),
            ParagraphList::Bullet
        );
        assert_eq!(
            editor
                .table_cell_paragraph_list(duplicate, ROW, COLUMN)
                .unwrap(),
            ParagraphList::Numbered
        );
    }

    #[test]
    fn scratch_documents_roundtrip_mixed_cell_lists_in_every_suite() {
        let expected = mixed_lists();
        let mut numbers = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let numbers_table = numbers.tables().unwrap()[0].object_id;
        numbers
            .set_cell(
                numbers_table,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        numbers
            .set_table_cell_paragraph_lists(numbers_table, ROW, COLUMN, &expected)
            .unwrap();
        let numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
        assert_eq!(
            numbers
                .table_cell_paragraph_lists(numbers_table, ROW, COLUMN)
                .unwrap(),
            expected
        );

        let mut pages = PagesDocumentBuilder::new()
            .body_table("Mixed Lists", 3, 3)
            .build()
            .unwrap();
        let pages_table = pages.tables().unwrap()[0].model_object_id;
        pages
            .set_table_cell(
                pages_table,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        pages
            .set_table_cell_paragraph_lists(pages_table, ROW, COLUMN, &expected)
            .unwrap();
        let pages = crate::pages::PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert_eq!(
            pages
                .table_cell_paragraph_lists(pages_table, ROW, COLUMN)
                .unwrap(),
            expected
        );

        let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
        let table = keynote
            .add_slide_table(
                0,
                "Mixed Lists",
                3,
                3,
                DrawablePoint { x: 100.0, y: 100.0 },
                DrawableSize {
                    width: 600.0,
                    height: 300.0,
                },
            )
            .unwrap();
        keynote
            .set_slide_table_cell(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_lists(0, table.model_object_id, ROW, COLUMN, &expected)
            .unwrap();
        let keynote =
            crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert_eq!(
            keynote
                .slide_table_cell_paragraph_lists(0, table.model_object_id, ROW, COLUMN,)
                .unwrap(),
            expected
        );
    }

    #[test]
    fn scratch_documents_roundtrip_custom_cell_bullets_in_every_suite() {
        let paragraph = ParagraphStart::from_utf16_index(9).unwrap();
        let arrow = ParagraphListBullet::new("➡").unwrap();
        let geometry = ParagraphListBulletGeometry::new(
            ParagraphListBulletScale::from_percent(175.0).unwrap(),
            ParagraphListBulletBaselineOffset::from_points(4.0).unwrap(),
        );
        let indentation = ParagraphListIndentation::new(
            ParagraphListLabelIndent::from_points(20.0).unwrap(),
            ParagraphListTextGap::from_points(18.0, TextPointSize::TWELVE).unwrap(),
        );

        let mut numbers = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let numbers_table = numbers.tables().unwrap()[0].object_id;
        numbers
            .set_cell(
                numbers_table,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        numbers
            .set_table_cell_paragraph_lists(numbers_table, ROW, COLUMN, &mixed_lists())
            .unwrap();
        numbers
            .set_table_cell_paragraph_list_bullet(numbers_table, ROW, COLUMN, paragraph, &arrow)
            .unwrap();
        numbers
            .set_table_cell_paragraph_list_bullet_geometry(
                numbers_table,
                ROW,
                COLUMN,
                paragraph,
                geometry,
            )
            .unwrap();
        numbers
            .set_table_cell_paragraph_list_indentation(
                numbers_table,
                ROW,
                COLUMN,
                paragraph,
                indentation,
            )
            .unwrap();
        let mut numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
        assert_eq!(
            numbers
                .table_cell_paragraph_list_bullet(numbers_table, ROW, COLUMN, paragraph)
                .unwrap(),
            arrow
        );
        assert_eq!(
            numbers
                .table_cell_paragraph_list_bullet_geometry(numbers_table, ROW, COLUMN, paragraph,)
                .unwrap(),
            geometry
        );
        assert_eq!(
            numbers
                .table_cell_paragraph_list_indentation(numbers_table, ROW, COLUMN, paragraph,)
                .unwrap(),
            indentation
        );
        assert!(
            numbers
                .reset_table_cell_paragraph_list_indentation(numbers_table, ROW, COLUMN, paragraph,)
                .unwrap()
        );
        assert_eq!(
            numbers
                .table_cell_paragraph_list_bullet_geometry(numbers_table, ROW, COLUMN, paragraph,)
                .unwrap(),
            geometry
        );
        assert_eq!(
            numbers
                .table_cell_paragraph_lists(numbers_table, ROW, COLUMN)
                .unwrap(),
            mixed_lists()
        );
        assert!(
            numbers
                .reset_table_cell_paragraph_list_bullet_geometry(
                    numbers_table,
                    ROW,
                    COLUMN,
                    paragraph,
                )
                .unwrap()
        );
        assert_eq!(
            numbers
                .table_cell_paragraph_list_bullet(numbers_table, ROW, COLUMN, paragraph)
                .unwrap(),
            arrow
        );
        assert!(
            !numbers
                .reset_table_cell_paragraph_list_bullet_geometry(
                    numbers_table,
                    ROW,
                    COLUMN,
                    paragraph,
                )
                .unwrap()
        );
        assert!(
            numbers
                .reset_table_cell_paragraph_list_bullet(numbers_table, ROW, COLUMN, paragraph)
                .unwrap()
        );
        assert!(
            !numbers
                .reset_table_cell_paragraph_list_bullet(numbers_table, ROW, COLUMN, paragraph)
                .unwrap()
        );

        let mut pages = PagesDocumentBuilder::new()
            .body_table("Bullets", 3, 3)
            .build()
            .unwrap();
        let pages_table = pages.tables().unwrap()[0].model_object_id;
        pages
            .set_table_cell(
                pages_table,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        pages
            .set_table_cell_paragraph_lists(pages_table, ROW, COLUMN, &mixed_lists())
            .unwrap();
        pages
            .set_table_cell_paragraph_list_bullet(pages_table, ROW, COLUMN, paragraph, &arrow)
            .unwrap();
        pages
            .set_table_cell_paragraph_list_bullet_geometry(
                pages_table,
                ROW,
                COLUMN,
                paragraph,
                geometry,
            )
            .unwrap();
        pages
            .set_table_cell_paragraph_list_indentation(
                pages_table,
                ROW,
                COLUMN,
                paragraph,
                indentation,
            )
            .unwrap();
        let pages = crate::pages::PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert_eq!(
            pages
                .table_cell_paragraph_list_bullet(pages_table, ROW, COLUMN, paragraph)
                .unwrap(),
            arrow
        );
        assert_eq!(
            pages
                .table_cell_paragraph_list_bullet_geometry(pages_table, ROW, COLUMN, paragraph,)
                .unwrap(),
            geometry
        );
        assert_eq!(
            pages
                .table_cell_paragraph_list_indentation(pages_table, ROW, COLUMN, paragraph,)
                .unwrap(),
            indentation
        );

        let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
        let table = keynote
            .add_slide_table(
                0,
                "Bullets",
                3,
                3,
                DrawablePoint { x: 100.0, y: 100.0 },
                DrawableSize {
                    width: 600.0,
                    height: 300.0,
                },
            )
            .unwrap();
        keynote
            .set_slide_table_cell(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_lists(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                &mixed_lists(),
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_list_bullet(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                paragraph,
                &arrow,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_list_bullet_geometry(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                paragraph,
                geometry,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_list_indentation(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                paragraph,
                indentation,
            )
            .unwrap();
        let keynote =
            crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert_eq!(
            keynote
                .slide_table_cell_paragraph_list_bullet(
                    0,
                    table.model_object_id,
                    ROW,
                    COLUMN,
                    paragraph,
                )
                .unwrap(),
            arrow
        );
        assert_eq!(
            keynote
                .slide_table_cell_paragraph_list_bullet_geometry(
                    0,
                    table.model_object_id,
                    ROW,
                    COLUMN,
                    paragraph,
                )
                .unwrap(),
            geometry
        );
        assert_eq!(
            keynote
                .slide_table_cell_paragraph_list_indentation(
                    0,
                    table.model_object_id,
                    ROW,
                    COLUMN,
                    paragraph,
                )
                .unwrap(),
            indentation
        );
    }

    #[test]
    #[allow(deprecated)]
    fn duplicated_rich_text_cells_are_custom_list_style_copy_on_write() {
        let paragraph = ParagraphStart::from_utf16_index(9).unwrap();
        let arrow = ParagraphListBullet::new("➡").unwrap();
        let diamond = ParagraphListBullet::new("◆").unwrap();
        let source_indentation = ParagraphListIndentation::new(
            ParagraphListLabelIndent::from_points(20.0).unwrap(),
            ParagraphListTextGap::from_em(1.5).unwrap(),
        );
        let duplicate_indentation = ParagraphListIndentation::new(
            ParagraphListLabelIndent::from_points(24.0).unwrap(),
            ParagraphListTextGap::from_em(2.0).unwrap(),
        );
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let source = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(source, ROW, COLUMN, CellValue::Text(MIXED_TEXT.to_owned()))
            .unwrap();
        editor
            .set_table_cell_paragraph_lists(source, ROW, COLUMN, &mixed_lists())
            .unwrap();
        editor
            .set_table_cell_paragraph_list_bullet(source, ROW, COLUMN, paragraph, &arrow)
            .unwrap();
        editor
            .set_table_cell_paragraph_list_indentation(
                source,
                ROW,
                COLUMN,
                paragraph,
                source_indentation,
            )
            .unwrap();
        let duplicate = editor.duplicate_table(source).unwrap().object_id;
        editor
            .set_table_cell_paragraph_list_bullet(duplicate, ROW, COLUMN, paragraph, &diamond)
            .unwrap();
        editor
            .set_table_cell_paragraph_list_indentation(
                duplicate,
                ROW,
                COLUMN,
                paragraph,
                duplicate_indentation,
            )
            .unwrap();
        assert_eq!(
            editor
                .table_cell_paragraph_list_bullet(source, ROW, COLUMN, paragraph)
                .unwrap(),
            arrow
        );
        assert_eq!(
            editor
                .table_cell_paragraph_list_bullet(duplicate, ROW, COLUMN, paragraph)
                .unwrap(),
            diamond
        );
        assert_eq!(
            editor
                .table_cell_paragraph_list_indentation(source, ROW, COLUMN, paragraph)
                .unwrap(),
            source_indentation
        );
        assert_eq!(
            editor
                .table_cell_paragraph_list_indentation(duplicate, ROW, COLUMN, paragraph)
                .unwrap(),
            duplicate_indentation
        );
    }

    #[test]
    #[allow(deprecated)]
    fn duplicated_rich_text_cells_are_mixed_list_copy_on_write() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let source = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(source, ROW, COLUMN, CellValue::Text(MIXED_TEXT.to_owned()))
            .unwrap();
        editor
            .set_table_cell_paragraph_lists(source, ROW, COLUMN, &mixed_lists())
            .unwrap();
        let duplicate = editor.duplicate_table(source).unwrap().object_id;
        let replacement = vec![ParagraphListPlacement::new(
            ParagraphStart::ZERO,
            ParagraphList::Numbered,
        )];
        editor
            .set_table_cell_paragraph_lists(duplicate, ROW, COLUMN, &replacement)
            .unwrap();
        assert_eq!(
            editor
                .table_cell_paragraph_lists(source, ROW, COLUMN)
                .unwrap(),
            mixed_lists()
        );
        assert_eq!(
            editor
                .table_cell_paragraph_lists(duplicate, ROW, COLUMN)
                .unwrap(),
            replacement
        );
    }

    #[test]
    fn invalid_mixed_list_boundary_is_transactional() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let table = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(table, ROW, COLUMN, CellValue::Text(MIXED_TEXT.to_owned()))
            .unwrap();
        let before = editor.to_bytes().unwrap();
        let invalid = [
            ParagraphListPlacement::new(ParagraphStart::ZERO, ParagraphList::None),
            ParagraphListPlacement::new(
                ParagraphStart::from_utf16_index(8).unwrap(),
                ParagraphList::Bullet,
            ),
        ];
        assert!(
            editor
                .set_table_cell_paragraph_lists(table, ROW, COLUMN, &invalid)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn scratch_documents_roundtrip_isolated_cell_list_levels_in_every_suite() {
        let paragraph = ParagraphStart::from_utf16_index(9).unwrap();
        let expected = nested_second_paragraph();
        let mut numbers = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let numbers_table = numbers.tables().unwrap()[0].object_id;
        numbers
            .set_cell(
                numbers_table,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        numbers
            .set_table_cell_paragraph_lists(numbers_table, ROW, COLUMN, &mixed_lists())
            .unwrap();
        numbers
            .set_table_cell_paragraph_list_level(
                numbers_table,
                ROW,
                COLUMN,
                paragraph,
                ParagraphListLevel::ONE,
            )
            .unwrap();
        let mut numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
        assert_eq!(
            numbers
                .table_cell_paragraph_list_levels(numbers_table, ROW, COLUMN)
                .unwrap(),
            expected
        );
        assert!(
            numbers
                .reset_table_cell_paragraph_list_level(numbers_table, ROW, COLUMN, paragraph,)
                .unwrap()
        );
        assert!(
            !numbers
                .reset_table_cell_paragraph_list_level(numbers_table, ROW, COLUMN, paragraph,)
                .unwrap()
        );

        let mut pages = PagesDocumentBuilder::new()
            .body_table("Nested Lists", 3, 3)
            .build()
            .unwrap();
        let pages_table = pages.tables().unwrap()[0].model_object_id;
        pages
            .set_table_cell(
                pages_table,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        pages
            .set_table_cell_paragraph_lists(pages_table, ROW, COLUMN, &mixed_lists())
            .unwrap();
        pages
            .set_table_cell_paragraph_list_level(
                pages_table,
                ROW,
                COLUMN,
                paragraph,
                ParagraphListLevel::ONE,
            )
            .unwrap();
        let pages = crate::pages::PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert_eq!(
            pages
                .table_cell_paragraph_list_levels(pages_table, ROW, COLUMN)
                .unwrap(),
            expected
        );

        let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
        let table = keynote
            .add_slide_table(
                0,
                "Nested Lists",
                3,
                3,
                DrawablePoint { x: 100.0, y: 100.0 },
                DrawableSize {
                    width: 600.0,
                    height: 300.0,
                },
            )
            .unwrap();
        keynote
            .set_slide_table_cell(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_lists(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                &mixed_lists(),
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_list_level(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                paragraph,
                ParagraphListLevel::ONE,
            )
            .unwrap();
        let keynote =
            crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert_eq!(
            keynote
                .slide_table_cell_paragraph_list_levels(0, table.model_object_id, ROW, COLUMN,)
                .unwrap(),
            expected
        );
    }

    #[test]
    fn scratch_documents_roundtrip_cell_list_numbering_in_every_suite() {
        let paragraph = ParagraphStart::from_utf16_index(16).unwrap();
        let restart =
            ParagraphListNumbering::StartAt(crate::text::ParagraphListStart::new(7).unwrap());

        let mut numbers = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let numbers_table = numbers.tables().unwrap()[0].object_id;
        numbers
            .set_cell(
                numbers_table,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        numbers
            .set_table_cell_paragraph_lists(numbers_table, ROW, COLUMN, &mixed_lists())
            .unwrap();
        numbers
            .set_table_cell_paragraph_list_numbering(numbers_table, ROW, COLUMN, paragraph, restart)
            .unwrap();
        let mut numbers = NumbersEditor::from_bytes(&numbers.to_bytes().unwrap()).unwrap();
        assert_eq!(
            numbers
                .table_cell_paragraph_list_numbering(numbers_table, ROW, COLUMN, paragraph)
                .unwrap(),
            restart
        );
        numbers
            .set_table_cell_paragraph_list_numbering(
                numbers_table,
                ROW,
                COLUMN,
                paragraph,
                ParagraphListNumbering::Continue,
            )
            .unwrap();
        assert_eq!(
            numbers
                .table_cell_paragraph_lists(numbers_table, ROW, COLUMN)
                .unwrap(),
            mixed_lists()
        );

        let mut pages = PagesDocumentBuilder::new()
            .body_table("Numbering", 3, 3)
            .build()
            .unwrap();
        let pages_table = pages.tables().unwrap()[0].model_object_id;
        pages
            .set_table_cell(
                pages_table,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        pages
            .set_table_cell_paragraph_lists(pages_table, ROW, COLUMN, &mixed_lists())
            .unwrap();
        pages
            .set_table_cell_paragraph_list_numbering(pages_table, ROW, COLUMN, paragraph, restart)
            .unwrap();
        let pages = crate::pages::PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert_eq!(
            pages
                .table_cell_paragraph_list_numbering(pages_table, ROW, COLUMN, paragraph)
                .unwrap(),
            restart
        );

        let mut keynote = KeynoteDocumentBuilder::new().build().unwrap();
        let table = keynote
            .add_slide_table(
                0,
                "Numbering",
                3,
                3,
                DrawablePoint { x: 100.0, y: 100.0 },
                DrawableSize {
                    width: 600.0,
                    height: 300.0,
                },
            )
            .unwrap();
        keynote
            .set_slide_table_cell(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                CellValue::Text(MIXED_TEXT.to_owned()),
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_lists(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                &mixed_lists(),
            )
            .unwrap();
        keynote
            .set_slide_table_cell_paragraph_list_numbering(
                0,
                table.model_object_id,
                ROW,
                COLUMN,
                paragraph,
                restart,
            )
            .unwrap();
        let keynote =
            crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert_eq!(
            keynote
                .slide_table_cell_paragraph_list_numbering(
                    0,
                    table.model_object_id,
                    ROW,
                    COLUMN,
                    paragraph,
                )
                .unwrap(),
            restart
        );
    }

    #[test]
    #[allow(deprecated)]
    fn duplicated_rich_text_cells_are_list_level_copy_on_write() {
        let paragraph = ParagraphStart::from_utf16_index(9).unwrap();
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let source = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(source, ROW, COLUMN, CellValue::Text(MIXED_TEXT.to_owned()))
            .unwrap();
        editor
            .set_table_cell_paragraph_lists(source, ROW, COLUMN, &mixed_lists())
            .unwrap();
        let duplicate = editor.duplicate_table(source).unwrap().object_id;
        editor
            .set_table_cell_paragraph_list_level(
                duplicate,
                ROW,
                COLUMN,
                paragraph,
                ParagraphListLevel::ONE,
            )
            .unwrap();
        assert!(
            editor
                .table_cell_paragraph_list_levels(source, ROW, COLUMN)
                .unwrap()
                .iter()
                .all(|placement| placement.level == ParagraphListLevel::ZERO)
        );
        assert_eq!(
            editor
                .table_cell_paragraph_list_levels(duplicate, ROW, COLUMN)
                .unwrap(),
            nested_second_paragraph()
        );
        assert_eq!(
            editor
                .table_cell_paragraph_lists(source, ROW, COLUMN)
                .unwrap(),
            mixed_lists()
        );
        assert_eq!(
            editor
                .table_cell_paragraph_lists(duplicate, ROW, COLUMN)
                .unwrap(),
            mixed_lists()
        );
    }

    #[test]
    #[allow(deprecated)]
    fn duplicated_rich_text_cells_are_list_numbering_copy_on_write() {
        let paragraph = ParagraphStart::from_utf16_index(16).unwrap();
        let restart =
            ParagraphListNumbering::StartAt(crate::text::ParagraphListStart::new(7).unwrap());
        let format = ParagraphListNumberFormat::affixed(
            ParagraphListNumberSequence::RomanUppercase,
            ParagraphListNumberPunctuation::RightParenthesis,
        );
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let source = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(source, ROW, COLUMN, CellValue::Text(MIXED_TEXT.to_owned()))
            .unwrap();
        editor
            .set_table_cell_paragraph_lists(source, ROW, COLUMN, &mixed_lists())
            .unwrap();
        let duplicate = editor.duplicate_table(source).unwrap().object_id;
        editor
            .set_table_cell_paragraph_list_numbering(duplicate, ROW, COLUMN, paragraph, restart)
            .unwrap();
        editor
            .set_table_cell_paragraph_list_number_format(duplicate, ROW, COLUMN, paragraph, format)
            .unwrap();
        assert_eq!(
            editor
                .table_cell_paragraph_list_numbering(source, ROW, COLUMN, paragraph)
                .unwrap(),
            ParagraphListNumbering::Continue
        );
        assert_eq!(
            editor
                .table_cell_paragraph_list_numbering(duplicate, ROW, COLUMN, paragraph)
                .unwrap(),
            restart
        );
        assert_eq!(
            editor
                .table_cell_paragraph_list_number_format(source, ROW, COLUMN, paragraph)
                .unwrap(),
            ParagraphListNumberFormat::DECIMAL
        );
        assert_eq!(
            editor
                .table_cell_paragraph_list_number_format(duplicate, ROW, COLUMN, paragraph)
                .unwrap(),
            format
        );
        assert_eq!(
            editor
                .table_cell_paragraph_lists(duplicate, ROW, COLUMN)
                .unwrap(),
            mixed_lists()
        );
    }

    #[test]
    #[allow(deprecated)]
    fn plain_cell_default_list_metadata_is_a_validated_no_op() {
        let paragraph = ParagraphStart::from_utf16_index(9).unwrap();
        let invalid = ParagraphStart::from_utf16_index(8).unwrap();
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let table = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(table, ROW, COLUMN, CellValue::Text(MIXED_TEXT.to_owned()))
            .unwrap();
        let before = editor.to_bytes().unwrap();

        editor
            .set_table_cell_paragraph_list_level(
                table,
                ROW,
                COLUMN,
                paragraph,
                ParagraphListLevel::ZERO,
            )
            .unwrap();
        assert!(
            !editor
                .reset_table_cell_paragraph_list_level(table, ROW, COLUMN, paragraph)
                .unwrap()
        );
        assert_eq!(
            editor
                .table_cell_paragraph_list_numbering(table, ROW, COLUMN, paragraph)
                .unwrap(),
            ParagraphListNumbering::Continue
        );
        editor
            .set_table_cell_paragraph_list_numbering(
                table,
                ROW,
                COLUMN,
                paragraph,
                ParagraphListNumbering::Continue,
            )
            .unwrap();
        assert_eq!(editor.to_bytes().unwrap(), before);

        assert!(
            editor
                .set_table_cell_paragraph_list_level(
                    table,
                    ROW,
                    COLUMN,
                    invalid,
                    ParagraphListLevel::ZERO,
                )
                .is_err()
        );
        assert!(
            editor
                .reset_table_cell_paragraph_list_level(table, ROW, COLUMN, invalid)
                .is_err()
        );
        assert!(
            editor
                .table_cell_paragraph_list_numbering(table, ROW, COLUMN, invalid)
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_paragraph_list_numbering(
                    table,
                    ROW,
                    COLUMN,
                    invalid,
                    ParagraphListNumbering::Continue,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn invalid_cell_list_level_boundary_is_transactional() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let table = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(table, ROW, COLUMN, CellValue::Text(MIXED_TEXT.to_owned()))
            .unwrap();
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_table_cell_paragraph_list_level(
                    table,
                    ROW,
                    COLUMN,
                    ParagraphStart::from_utf16_index(8).unwrap(),
                    ParagraphListLevel::ONE,
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }

    #[test]
    fn invalid_coordinate_is_transactional() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let table = editor.tables().unwrap()[0].object_id;
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_table_cell_paragraph_list(table, 2, 0, ParagraphList::Bullet)
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
