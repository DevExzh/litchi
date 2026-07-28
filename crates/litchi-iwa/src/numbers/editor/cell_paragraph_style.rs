//! Copy-on-write paragraph-style properties for native table cells.

mod property;

use prost::Message;

use super::*;
use crate::text::paragraph_alignment::native::{
    ParagraphStyleOverrides, locate_style, parent_style_id, replace_variation, stylesheet_id,
    variation_object,
};
use crate::text::style_registry::{
    register_private_style, register_style_reference, unregister_owner_reference_if_unused,
    unregister_private_style,
};
use property::{CellParagraphProperty, CellParagraphPropertyKind};

const PARAGRAPH_STYLE_MESSAGE_TYPE: u32 = 2_022;

pub(super) fn alignment(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextAlignment> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Alignment,
    )? {
        CellParagraphProperty::Alignment(value) => Ok(value),
        CellParagraphProperty::Color(_)
        | CellParagraphProperty::Decorations(_)
        | CellParagraphProperty::Font(_)
        | CellParagraphProperty::TextStyle(_) => Err(Error::InvalidFormat(
            "iWork table-cell alignment resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_alignment(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextAlignment,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Alignment(value),
    )
}

pub(super) fn reset_alignment(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Alignment,
    )
}

pub(super) fn text_color(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<RgbaColor> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Color,
    )? {
        CellParagraphProperty::Color(value) => Ok(value),
        CellParagraphProperty::Alignment(_)
        | CellParagraphProperty::Decorations(_)
        | CellParagraphProperty::Font(_)
        | CellParagraphProperty::TextStyle(_) => Err(Error::InvalidFormat(
            "iWork table-cell text color resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_text_color(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: RgbaColor,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Color(value),
    )
}

pub(super) fn reset_text_color(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Color,
    )
}

pub(super) fn decorations(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextDecorations> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Decorations,
    )? {
        CellParagraphProperty::Decorations(value) => Ok(value),
        CellParagraphProperty::Alignment(_)
        | CellParagraphProperty::Color(_)
        | CellParagraphProperty::Font(_)
        | CellParagraphProperty::TextStyle(_) => Err(Error::InvalidFormat(
            "iWork table-cell text decorations resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_decorations(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextDecorations,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Decorations(value),
    )
}

pub(super) fn reset_decorations(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Decorations,
    )
}

pub(super) fn font(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextFont> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Font,
    )? {
        CellParagraphProperty::Font(value) => Ok(value),
        CellParagraphProperty::Alignment(_)
        | CellParagraphProperty::Color(_)
        | CellParagraphProperty::Decorations(_)
        | CellParagraphProperty::TextStyle(_) => Err(Error::InvalidFormat(
            "iWork table-cell font resolved as another paragraph property".to_owned(),
        )),
    }
}

pub(super) fn set_font(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextFont,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::Font(value),
    )
}

pub(super) fn reset_font(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::Font,
    )
}

pub(super) fn text_style(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<TextStyle> {
    match property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::TextStyle,
    )? {
        CellParagraphProperty::TextStyle(value) => Ok(value),
        CellParagraphProperty::Alignment(_)
        | CellParagraphProperty::Color(_)
        | CellParagraphProperty::Decorations(_)
        | CellParagraphProperty::Font(_) => Err(Error::InvalidFormat(
            "iWork table-cell character formatting resolved as another paragraph property"
                .to_owned(),
        )),
    }
}

pub(super) fn set_text_style(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: TextStyle,
) -> Result<()> {
    set_property(
        package,
        table_id,
        row,
        column,
        CellParagraphProperty::TextStyle(value),
    )
}

pub(super) fn reset_text_style(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    reset_property(
        package,
        table_id,
        row,
        column,
        CellParagraphPropertyKind::TextStyle,
    )
}

fn property(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    kind: CellParagraphPropertyKind,
) -> Result<CellParagraphProperty> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    validate_coordinate(&descriptor, row, column)?;
    let locations = storage::object_locations(package)?;
    let style_id = local_style_id(package, &descriptor, &locations, row, column)?
        .unwrap_or_else(|| base_style_id(&descriptor, row, column));
    CellParagraphProperty::inherited(package, style_id, kind)
}

fn set_property(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    value: CellParagraphProperty,
) -> Result<()> {
    let kind = value.kind();
    if property(package, table_id, row, column, kind)? == value {
        return Ok(());
    }
    let mut staged = package.clone();
    table_sparse_storage::ensure_attached_cell_storage(&mut staged, table_id, row, column)?;
    let location = model::locate_attached_cell(&staged, table_id, row, column)?;
    let locations = location.object_locations.clone();
    let style_table_id = location
        .descriptor
        .model
        .base_data_store
        .style_table
        .identifier;
    let old_key = read_bnc(&staged, &location, column)?.text_style_identifier();
    let resolved = storage::resolve_table_data_list(
        &staged,
        &locations,
        style_table_id,
        tst::table_data_list::ListType::Style,
    )?;
    let old_entry = style_entry(&resolved, old_key)?;
    let current_style_id = old_entry
        .and_then(|entry| entry.entry.reference.as_ref())
        .map_or_else(
            || base_style_id(&location.descriptor, row, column),
            |reference| reference.identifier,
        );
    let current_style = locate_style(&staged, current_style_id)?;
    let current_value = CellParagraphProperty::inherited(&staged, current_style_id, kind)?;

    if let Some(entry) = old_entry
        && entry.entry.refcount == 1
        && style_is_exclusive_to_list(
            &staged,
            current_style_id,
            stylesheet_id(&current_style.style, current_style_id)?,
            entry_owner_id(&resolved, entry),
        )?
        && let Some(mut overrides) = crate::text::paragraph_alignment::native::direct_overrides(
            &current_style.style,
            &current_style.message.data,
        )?
    {
        let parent_id = parent_style_id(&current_style.style, current_style_id)?;
        let inherited = CellParagraphProperty::inherited(&staged, parent_id, kind)?;
        if value == inherited {
            drop(staged);
            if !reset_property(package, table_id, row, column, kind)? {
                return Err(Error::InvalidFormat(format!(
                    "iWork table-cell {} could not restore its inherited value",
                    kind.name()
                )));
            }
            return Ok(());
        }
        value.apply_to(&mut overrides, &inherited)?;
        let replacement = variation_object(
            current_style_id,
            parent_id,
            stylesheet_id(&current_style.style, current_style_id)?,
            overrides,
        )?;
        replace_variation(
            &mut staged,
            &current_style.archive_name,
            current_style_id,
            replacement,
        )?;
        verify_property(&staged, table_id, row, column, &value)?;
        *package = staged;
        return Ok(());
    }

    let new_style_id = next_object_identifier(&staged)?;
    let stylesheet_id = current_style
        .style
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork paragraph style {current_style_id} has no stylesheet"
            ))
        })?;
    let mut overrides = ParagraphStyleOverrides::default();
    value.apply_to(&mut overrides, &current_value)?;
    let variation = variation_object(new_style_id, current_style_id, stylesheet_id, overrides)?;
    crate::shapes::insert_style_variation(
        &mut staged,
        &current_style.archive_name,
        stylesheet_id,
        current_style_id,
        new_style_id,
        variation,
    )?;
    let new_key = insert_style_entry(&mut staged, &locations, style_table_id, new_style_id)?;
    register_private_style(
        &mut staged,
        &resolved.table_archive,
        &current_style.archive_name,
        new_style_id,
    )?;
    write_text_style_key(&mut staged, &location, row, column, Some(new_key))?;
    if let Some(entry) = old_entry {
        let removed = storage::decrement_table_data_list_entry(
            &mut staged,
            &locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )?;
        if removed {
            unregister_owner_reference_if_unused(
                &mut staged,
                &resolved.table_archive,
                &current_style.archive_name,
                current_style_id,
            )?;
        }
    }
    set_package_last_object_identifier(&mut staged, new_style_id)?;
    verify_property(&staged, table_id, row, column, &value)?;
    *package = staged;
    Ok(())
}

fn reset_property(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    kind: CellParagraphPropertyKind,
) -> Result<bool> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    validate_coordinate(&descriptor, row, column)?;
    let locations = storage::object_locations(package)?;
    let Some(key) = text_style_key(package, &descriptor, row, column)? else {
        return Ok(false);
    };
    let style_table_id = descriptor.model.base_data_store.style_table.identifier;
    let resolved = storage::resolve_table_data_list(
        package,
        &locations,
        style_table_id,
        tst::table_data_list::ListType::Style,
    )?;
    let entry = style_entry(&resolved, Some(key))?
        .ok_or_else(|| Error::InvalidFormat(format!("iWork text style table has no key {key}")))?;
    let style_id = entry
        .entry
        .reference
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork text style key {key} has no reference"))
        })?;
    let style = locate_style(package, style_id)?;
    let Some(mut overrides) = crate::text::paragraph_alignment::native::direct_overrides(
        &style.style,
        &style.message.data,
    )?
    else {
        return Ok(false);
    };
    if !kind.has_direct(&overrides) {
        return Ok(false);
    }
    let parent_id = parent_style_id(&style.style, style_id)?;
    let inherited = CellParagraphProperty::inherited(package, parent_id, kind)?;

    if entry.entry.refcount == 1
        && style_is_exclusive_to_list(
            package,
            style_id,
            stylesheet_id(&style.style, style_id)?,
            entry_owner_id(&resolved, entry),
        )?
    {
        let mut staged = package.clone();
        kind.clear(&mut overrides);
        if overrides.is_empty() {
            let location = model::locate_attached_cell(&staged, table_id, row, column)?;
            let base_id = base_style_id(&descriptor, row, column);
            if parent_id == base_id {
                write_text_style_key(&mut staged, &location, row, column, None)?;
            } else {
                let parent_key =
                    attach_style_entry(&mut staged, &locations, style_table_id, parent_id)?;
                write_text_style_key(&mut staged, &location, row, column, Some(parent_key))?;
            }
            let removed = storage::decrement_table_data_list_entry(
                &mut staged,
                &locations,
                &resolved,
                entry,
                tst::table_data_list::ListType::Style,
            )?;
            if removed && !style_has_children(&staged, style_id)? {
                crate::shapes::remove_style_variation(
                    &mut staged,
                    &style.archive_name,
                    stylesheet_id(&style.style, style_id)?,
                    parent_id,
                    style_id,
                )?;
                unregister_private_style(
                    &mut staged,
                    &resolved.table_archive,
                    &style.archive_name,
                    style_id,
                    Some(parent_id),
                )?;
                release_package_identifier_suffix(&mut staged, &[style_id])?;
            }
        } else {
            let replacement = variation_object(
                style_id,
                parent_id,
                stylesheet_id(&style.style, style_id)?,
                overrides,
            )?;
            replace_variation(&mut staged, &style.archive_name, style_id, replacement)?;
        }
        verify_property(&staged, table_id, row, column, &inherited)?;
        *package = staged;
        return Ok(true);
    }

    let mut staged = package.clone();
    let location = model::locate_attached_cell(&staged, table_id, row, column)?;
    kind.clear(&mut overrides);
    if overrides.is_empty() {
        let base_id = base_style_id(&descriptor, row, column);
        if parent_id == base_id {
            write_text_style_key(&mut staged, &location, row, column, None)?;
        } else {
            let parent_key =
                attach_style_entry(&mut staged, &locations, style_table_id, parent_id)?;
            write_text_style_key(&mut staged, &location, row, column, Some(parent_key))?;
        }
    } else {
        let new_style_id = next_object_identifier(&staged)?;
        let parent = locate_style(&staged, parent_id)?;
        let stylesheet_id = parent
            .style
            .super_
            .stylesheet
            .as_ref()
            .map(|reference| reference.identifier)
            .ok_or_else(|| {
                Error::InvalidFormat(format!(
                    "iWork paragraph style {parent_id} has no stylesheet"
                ))
            })?;
        let variation = variation_object(new_style_id, parent_id, stylesheet_id, overrides)?;
        crate::shapes::insert_style_variation(
            &mut staged,
            &parent.archive_name,
            stylesheet_id,
            parent_id,
            new_style_id,
            variation,
        )?;
        let new_key = insert_style_entry(&mut staged, &locations, style_table_id, new_style_id)?;
        register_private_style(
            &mut staged,
            &resolved.table_archive,
            &parent.archive_name,
            new_style_id,
        )?;
        write_text_style_key(&mut staged, &location, row, column, Some(new_key))?;
        set_package_last_object_identifier(&mut staged, new_style_id)?;
    }
    let removed = storage::decrement_table_data_list_entry(
        &mut staged,
        &locations,
        &resolved,
        entry,
        tst::table_data_list::ListType::Style,
    )?;
    if removed {
        unregister_owner_reference_if_unused(
            &mut staged,
            &resolved.table_archive,
            &style.archive_name,
            style_id,
        )?;
    }
    verify_property(&staged, table_id, row, column, &inherited)?;
    *package = staged;
    Ok(true)
}

fn validate_coordinate(
    descriptor: &model::TableDescriptor,
    row: usize,
    column: usize,
) -> Result<()> {
    if row >= descriptor.model.number_of_rows as usize
        || column >= descriptor.model.number_of_columns as usize
    {
        return Err(Error::ParseError(format!(
            "Cell ({row}, {column}) is outside iWork table {:?} dimensions {}x{}",
            descriptor.model.table_name,
            descriptor.model.number_of_rows,
            descriptor.model.number_of_columns
        )));
    }
    Ok(())
}

fn base_style_id(descriptor: &model::TableDescriptor, row: usize, column: usize) -> u64 {
    let model = &descriptor.model;
    if row < model.number_of_header_rows.unwrap_or(0) as usize {
        model.header_row_text_style.identifier
    } else if row
        >= model
            .number_of_rows
            .saturating_sub(model.number_of_footer_rows.unwrap_or(0)) as usize
    {
        model.footer_row_text_style.identifier
    } else if column < model.number_of_header_columns.unwrap_or(0) as usize {
        model.header_column_text_style.identifier
    } else {
        model.body_text_style.identifier
    }
}

fn local_style_id(
    package: &IWorkPackage,
    descriptor: &model::TableDescriptor,
    locations: &HashMap<u64, String>,
    row: usize,
    column: usize,
) -> Result<Option<u64>> {
    let Some(key) = text_style_key(package, descriptor, row, column)? else {
        return Ok(None);
    };
    let resolved = storage::resolve_table_data_list(
        package,
        locations,
        descriptor.model.base_data_store.style_table.identifier,
        tst::table_data_list::ListType::Style,
    )?;
    let entry = style_entry(&resolved, Some(key))?
        .ok_or_else(|| Error::InvalidFormat(format!("iWork text style table has no key {key}")))?;
    entry
        .entry
        .reference
        .as_ref()
        .map(|reference| Some(reference.identifier))
        .ok_or_else(|| Error::InvalidFormat(format!("iWork text style key {key} has no reference")))
}

fn text_style_key(
    package: &IWorkPackage,
    descriptor: &model::TableDescriptor,
    row: usize,
    column: usize,
) -> Result<Option<u32>> {
    let location = model::locate_attached_cell(package, descriptor.object_id, row, column)?;
    storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    .map(|data| BncCell::parse(&data).map(|cell| cell.text_style_identifier()))
    .transpose()
    .map(Option::flatten)
}

fn style_entry(
    resolved: &storage::ResolvedTableDataList,
    key: Option<u32>,
) -> Result<Option<&storage::LocatedTableDataListEntry>> {
    let Some(key) = key else {
        return Ok(None);
    };
    let mut matches = resolved
        .entries
        .iter()
        .filter(|entry| entry.entry.key == key);
    let entry = matches.next();
    if matches.next().is_some() {
        return Err(Error::InvalidFormat(format!(
            "iWork text style table repeats key {key}"
        )));
    }
    Ok(entry)
}

fn insert_style_entry(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    style_id: u64,
) -> Result<u32> {
    let resolved = storage::resolve_table_data_list(
        package,
        locations,
        table_id,
        tst::table_data_list::ListType::Style,
    )?;
    let key = storage::next_table_data_list_key(&resolved.list, &resolved.entries)?;
    package.update_archive(&resolved.table_archive, |archive| {
        let object = archive.object_mut(table_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork style table {table_id} is missing"))
        })?;
        let index = storage::table_data_list_message_index(
            object,
            tst::table_data_list::ListType::Style,
        )
        .ok_or_else(|| Error::InvalidFormat(format!("Object {table_id} has no style list")))?;
        let previous = TableDataList::decode(object.messages[index].data.as_slice())?;
        let mut current = previous.clone();
        current.next_list_id = key
            .checked_add(1)
            .ok_or_else(|| Error::ParseError("iWork text style key overflow".to_owned()))?;
        current.entries.push(tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            reference: Some(tsp::Reference {
                identifier: style_id,
                ..Default::default()
            }),
            ..Default::default()
        });
        let data = storage::rewrite_table_data_list_wire(
            object.messages[index].data.as_slice(),
            &previous,
            &current,
        )?;
        let message_type = object.messages[index].type_;
        object.replace_message(
            index,
            RawMessage {
                type_: message_type,
                data,
            },
        )?;
        storage::add_message_object_reference(object, index, style_id, style_id);
        Ok(())
    })?;
    Ok(key)
}

fn attach_style_entry(
    package: &mut IWorkPackage,
    locations: &HashMap<u64, String>,
    table_id: u64,
    style_id: u64,
) -> Result<u32> {
    let resolved = storage::resolve_table_data_list(
        package,
        locations,
        table_id,
        tst::table_data_list::ListType::Style,
    )?;
    if let Some(entry) = resolved.entries.iter().find(|entry| {
        entry
            .entry
            .reference
            .as_ref()
            .is_some_and(|reference| reference.identifier == style_id)
    }) {
        let key = entry.entry.key;
        storage::increment_table_data_list_entry(
            package,
            locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )?;
        return Ok(key);
    }
    let key = insert_style_entry(package, locations, table_id, style_id)?;
    let style = locate_style(package, style_id)?;
    register_style_reference(
        package,
        &resolved.table_archive,
        &style.archive_name,
        style_id,
    )?;
    Ok(key)
}

fn read_bnc(
    package: &IWorkPackage,
    location: &model::CellLocation,
    column: usize,
) -> Result<BncCell> {
    storage::read_tile_cell(
        package,
        &location.tile_archive,
        location.tile_id,
        location.tile_row,
        column,
    )?
    .map_or_else(|| Ok(BncCell::minimal()), |data| BncCell::parse(&data))
}

fn write_text_style_key(
    package: &mut IWorkPackage,
    location: &model::CellLocation,
    row: usize,
    column: usize,
    key: Option<u32>,
) -> Result<()> {
    let mut cell = read_bnc(package, location, column)?;
    cell.set_text_style_identifier(key);
    storage::set_encoded_cell_value(
        package,
        location.descriptor.object_id,
        row,
        column,
        EncodedValue::Raw(cell.encode()),
    )
}

fn entry_owner_id(
    resolved: &storage::ResolvedTableDataList,
    entry: &storage::LocatedTableDataListEntry,
) -> u64 {
    match &entry.owner {
        storage::TableDataListEntryOwner::Root => resolved.table_id,
        storage::TableDataListEntryOwner::Segment { object_id, .. } => *object_id,
    }
}

fn style_is_exclusive_to_list(
    package: &IWorkPackage,
    style_id: u64,
    stylesheet_id: u64,
    list_owner_id: u64,
) -> Result<bool> {
    for archive_name in package.iwa_entry_names() {
        for object in package.archive(archive_name)?.objects {
            let identifier = object.archive_info.identifier.unwrap_or_default();
            if matches!(identifier, id if id == stylesheet_id || id == list_owner_id || id == style_id)
            {
                continue;
            }
            if object.archive_info.message_infos.iter().any(|info| {
                info.object_references.contains(&style_id)
                    || info
                        .field_infos
                        .iter()
                        .any(|field| field.object_references.contains(&style_id))
            }) {
                return Ok(false);
            }
        }
    }
    Ok(true)
}

fn style_has_children(package: &IWorkPackage, style_id: u64) -> Result<bool> {
    for archive_name in package.iwa_entry_names() {
        if package.archive(archive_name)?.objects.iter().any(|object| {
            object.messages.iter().any(|message| {
                message.type_ == PARAGRAPH_STYLE_MESSAGE_TYPE
                    && tswp::ParagraphStyleArchive::decode(message.data.as_slice()).is_ok_and(
                        |style| {
                            style
                                .super_
                                .parent
                                .as_ref()
                                .is_some_and(|parent| parent.identifier == style_id)
                        },
                    )
            })
        }) {
            return Ok(true);
        }
    }
    Ok(false)
}

fn verify_property(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    expected: &CellParagraphProperty,
) -> Result<()> {
    if property(package, table_id, row, column, expected.kind())? != *expected {
        return Err(Error::InvalidFormat(format!(
            "iWork table-cell {} failed validation",
            expected.kind().name()
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::keynote::KeynoteDocumentBuilder;
    use crate::numbers::{CellValue, NumbersDocumentBuilder};
    use crate::pages::PagesDocumentBuilder;
    use crate::shapes::{DrawablePoint, DrawableSize, RgbColorSpace, RgbaColor};
    use crate::text::{TextDecorations, TextFont, TextPointSize, TextStrikethrough, TextUnderline};

    fn test_color() -> RgbaColor {
        const RED: f32 = 0.72;
        const GREEN: f32 = 0.10;
        const BLUE: f32 = 0.14;
        const ALPHA: f32 = 1.0;
        RgbaColor::new(RED, GREEN, BLUE, ALPHA, RgbColorSpace::Srgb).unwrap()
    }

    fn explicit_style_id(editor: &NumbersEditor, table_id: u64, row: usize, column: usize) -> u64 {
        let descriptor = model::attached_table_descriptor(&editor.package, table_id).unwrap();
        let locations = storage::object_locations(&editor.package).unwrap();
        local_style_id(&editor.package, &descriptor, &locations, row, column)
            .unwrap()
            .unwrap()
    }

    #[test]
    fn scratch_alignment_reuses_and_reclaims_private_style() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(3, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(table_id, 1, 1, CellValue::Text("Aligned".to_owned()))
            .unwrap();
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Natural
        );

        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Center)
            .unwrap();
        let style_id = explicit_style_id(&editor, table_id, 1, 1);
        let location = model::locate_attached_cell(&editor.package, table_id, 1, 1).unwrap();
        let cell = read_bnc(&editor.package, &location, 1).unwrap();
        assert!(cell.text_style_identifier().is_some());
        assert!(cell.style_identifier().is_none());

        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Right)
            .unwrap();
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), style_id);
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Right
        );

        assert!(
            editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Natural
        );
        let location = model::locate_attached_cell(&editor.package, table_id, 1, 1).unwrap();
        assert!(
            read_bnc(&editor.package, &location, 1)
                .unwrap()
                .text_style_identifier()
                .is_none()
        );
        assert!(editor.package.iwa_entry_names().all(|archive_name| {
            editor
                .package
                .archive(archive_name)
                .unwrap()
                .object(style_id)
                .is_none()
        }));
        assert!(
            !editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
    }

    #[test]
    fn shared_alignment_is_copy_on_write_and_reset_is_idempotent() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        for column in 1..=2 {
            editor
                .set_cell(
                    table_id,
                    1,
                    column,
                    CellValue::Text(format!("Column {column}")),
                )
                .unwrap();
        }
        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Center)
            .unwrap();
        let descriptor = model::attached_table_descriptor(&editor.package, table_id).unwrap();
        let locations = storage::object_locations(&editor.package).unwrap();
        let key = text_style_key(&editor.package, &descriptor, 1, 1)
            .unwrap()
            .unwrap();
        let resolved = storage::resolve_table_data_list(
            &editor.package,
            &locations,
            descriptor.model.base_data_store.style_table.identifier,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let entry = style_entry(&resolved, Some(key)).unwrap().unwrap();
        storage::increment_table_data_list_entry(
            &mut editor.package,
            &locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let target = model::locate_attached_cell(&editor.package, table_id, 1, 2).unwrap();
        write_text_style_key(&mut editor.package, &target, 1, 2, Some(key)).unwrap();
        let shared_style = explicit_style_id(&editor, table_id, 1, 1);

        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Right)
            .unwrap();

        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style);
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Right
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 2).unwrap(),
            TextAlignment::Center
        );

        assert!(
            editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Center
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 2).unwrap(),
            TextAlignment::Center
        );
        assert!(
            editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Natural
        );
        assert!(
            !editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
    }

    #[test]
    fn character_properties_compose_with_alignment_and_reclaim_independently() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        editor
            .set_cell(table_id, 1, 1, CellValue::Text("Styled".to_owned()))
            .unwrap();
        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Center)
            .unwrap();
        let style_id = explicit_style_id(&editor, table_id, 1, 1);
        let styled = TextStyle::new(TextPointSize::from_points(18.0).unwrap())
            .with_bold(true)
            .with_italic(true);
        let font = TextFont::named("CourierNewPSMT").unwrap();
        let color = test_color();
        let decorations = TextDecorations::new(TextUnderline::Double, TextStrikethrough::Single);

        editor
            .set_table_cell_text_style(table_id, 1, 1, styled)
            .unwrap();
        editor
            .set_table_cell_text_font(table_id, 1, 1, font.clone())
            .unwrap();
        editor
            .set_table_cell_text_color(table_id, 1, 1, color)
            .unwrap();
        editor
            .set_table_cell_text_decorations(table_id, 1, 1, decorations)
            .unwrap();
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), style_id);
        assert_eq!(
            editor.table_cell_text_style(table_id, 1, 1).unwrap(),
            styled
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 1).unwrap(),
            TextAlignment::Center
        );
        assert_eq!(editor.table_cell_text_font(table_id, 1, 1).unwrap(), font);
        assert_eq!(editor.table_cell_text_color(table_id, 1, 1).unwrap(), color);
        assert_eq!(
            editor.table_cell_text_decorations(table_id, 1, 1).unwrap(),
            decorations
        );

        assert!(
            editor
                .reset_table_cell_text_decorations(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(
            editor.table_cell_text_decorations(table_id, 1, 1).unwrap(),
            TextDecorations::NONE
        );
        assert_eq!(editor.table_cell_text_color(table_id, 1, 1).unwrap(), color);
        assert!(editor.reset_table_cell_text_color(table_id, 1, 1).unwrap());
        assert_eq!(
            editor.table_cell_text_color(table_id, 1, 1).unwrap(),
            RgbaColor::black()
        );
        assert_eq!(editor.table_cell_text_font(table_id, 1, 1).unwrap(), font);

        editor
            .set_table_cell_text_style(table_id, 1, 1, TextStyle::default())
            .unwrap();
        assert_eq!(
            editor.table_cell_text_style(table_id, 1, 1).unwrap(),
            TextStyle::default()
        );
        assert!(!editor.reset_table_cell_text_style(table_id, 1, 1).unwrap());
        editor
            .set_table_cell_text_font(table_id, 1, 1, TextFont::default())
            .unwrap();
        assert_eq!(
            editor.table_cell_text_font(table_id, 1, 1).unwrap(),
            TextFont::default()
        );
        assert!(!editor.reset_table_cell_text_font(table_id, 1, 1).unwrap());
        editor
            .set_table_cell_text_font(table_id, 1, 1, font)
            .unwrap();
        assert!(editor.reset_table_cell_text_font(table_id, 1, 1).unwrap());
        editor
            .set_table_cell_text_style(table_id, 1, 1, styled)
            .unwrap();
        assert!(editor.reset_table_cell_text_style(table_id, 1, 1).unwrap());
        assert_eq!(
            editor.table_cell_text_style(table_id, 1, 1).unwrap(),
            TextStyle::default()
        );
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), style_id);
        assert!(!editor.reset_table_cell_text_style(table_id, 1, 1).unwrap());
        assert!(
            editor
                .reset_table_cell_text_alignment(table_id, 1, 1)
                .unwrap()
        );
        assert!(editor.package.iwa_entry_names().all(|archive_name| {
            editor
                .package
                .archive(archive_name)
                .unwrap()
                .object(style_id)
                .is_none()
        }));
    }

    #[test]
    fn shared_character_properties_use_copy_on_write() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        for column in 1..=2 {
            editor
                .set_cell(
                    table_id,
                    1,
                    column,
                    CellValue::Text(format!("Column {column}")),
                )
                .unwrap();
        }
        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Center)
            .unwrap();
        let descriptor = model::attached_table_descriptor(&editor.package, table_id).unwrap();
        let locations = storage::object_locations(&editor.package).unwrap();
        let key = text_style_key(&editor.package, &descriptor, 1, 1)
            .unwrap()
            .unwrap();
        let resolved = storage::resolve_table_data_list(
            &editor.package,
            &locations,
            descriptor.model.base_data_store.style_table.identifier,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let entry = style_entry(&resolved, Some(key)).unwrap().unwrap();
        storage::increment_table_data_list_entry(
            &mut editor.package,
            &locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let target = model::locate_attached_cell(&editor.package, table_id, 1, 2).unwrap();
        write_text_style_key(&mut editor.package, &target, 1, 2, Some(key)).unwrap();
        let shared_style_id = explicit_style_id(&editor, table_id, 1, 1);
        let font = TextFont::named("CourierNewPSMT").unwrap();
        let color = test_color();
        let decorations = TextDecorations::new(TextUnderline::Single, TextStrikethrough::Double);

        editor
            .set_table_cell_text_decorations(table_id, 1, 1, decorations)
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(
            editor.table_cell_text_decorations(table_id, 1, 1).unwrap(),
            decorations
        );
        assert_eq!(
            editor.table_cell_text_decorations(table_id, 1, 2).unwrap(),
            TextDecorations::NONE
        );
        assert!(
            editor
                .reset_table_cell_text_decorations(table_id, 1, 1)
                .unwrap()
        );
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);

        editor
            .set_table_cell_text_color(table_id, 1, 1, color)
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(editor.table_cell_text_color(table_id, 1, 1).unwrap(), color);
        assert_eq!(
            editor.table_cell_text_color(table_id, 1, 2).unwrap(),
            RgbaColor::black()
        );
        assert!(editor.reset_table_cell_text_color(table_id, 1, 1).unwrap());
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);

        editor
            .set_table_cell_text_font(table_id, 1, 1, font.clone())
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(editor.table_cell_text_font(table_id, 1, 1).unwrap(), font);
        assert_eq!(
            editor.table_cell_text_font(table_id, 1, 2).unwrap(),
            TextFont::default()
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 2).unwrap(),
            TextAlignment::Center
        );

        assert!(editor.reset_table_cell_text_font(table_id, 1, 1).unwrap());
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert!(!editor.reset_table_cell_text_font(table_id, 1, 1).unwrap());
    }

    #[test]
    fn shared_text_style_uses_copy_on_write() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 3)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        for column in 1..=2 {
            editor
                .set_cell(
                    table_id,
                    1,
                    column,
                    CellValue::Text(format!("Column {column}")),
                )
                .unwrap();
        }
        editor
            .set_table_cell_text_alignment(table_id, 1, 1, TextAlignment::Center)
            .unwrap();
        let descriptor = model::attached_table_descriptor(&editor.package, table_id).unwrap();
        let locations = storage::object_locations(&editor.package).unwrap();
        let key = text_style_key(&editor.package, &descriptor, 1, 1)
            .unwrap()
            .unwrap();
        let resolved = storage::resolve_table_data_list(
            &editor.package,
            &locations,
            descriptor.model.base_data_store.style_table.identifier,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let entry = style_entry(&resolved, Some(key)).unwrap().unwrap();
        storage::increment_table_data_list_entry(
            &mut editor.package,
            &locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )
        .unwrap();
        let target = model::locate_attached_cell(&editor.package, table_id, 1, 2).unwrap();
        write_text_style_key(&mut editor.package, &target, 1, 2, Some(key)).unwrap();
        let shared_style_id = explicit_style_id(&editor, table_id, 1, 1);
        let styled = TextStyle::new(TextPointSize::from_points(20.0).unwrap()).with_bold(true);

        editor
            .set_table_cell_text_style(table_id, 1, 1, styled)
            .unwrap();
        assert_ne!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert_eq!(
            editor.table_cell_text_style(table_id, 1, 1).unwrap(),
            styled
        );
        assert_eq!(
            editor.table_cell_text_style(table_id, 1, 2).unwrap(),
            TextStyle::default()
        );
        assert_eq!(
            editor.table_cell_text_alignment(table_id, 1, 2).unwrap(),
            TextAlignment::Center
        );

        assert!(editor.reset_table_cell_text_style(table_id, 1, 1).unwrap());
        assert_eq!(explicit_style_id(&editor, table_id, 1, 1), shared_style_id);
        assert!(!editor.reset_table_cell_text_style(table_id, 1, 1).unwrap());
    }

    #[test]
    fn scratch_paragraph_styles_round_trip_in_pages_and_keynote() {
        let pages_style =
            TextStyle::new(TextPointSize::from_points(17.0).unwrap()).with_italic(true);
        let pages_font = TextFont::named("AvenirNext-Regular").unwrap();
        let pages_color = test_color();
        let pages_decorations =
            TextDecorations::new(TextUnderline::Single, TextStrikethrough::None);
        let mut pages = PagesDocumentBuilder::new()
            .body_table("Aligned", 2, 2)
            .build()
            .unwrap();
        let pages_table = pages.tables().unwrap()[0].model_object_id;
        pages
            .set_table_cell_text_alignment(pages_table, 1, 1, TextAlignment::Justified)
            .unwrap();
        pages
            .set_table_cell_text_style(pages_table, 1, 1, pages_style)
            .unwrap();
        pages
            .set_table_cell_text_font(pages_table, 1, 1, pages_font.clone())
            .unwrap();
        pages
            .set_table_cell_text_color(pages_table, 1, 1, pages_color)
            .unwrap();
        pages
            .set_table_cell_text_decorations(pages_table, 1, 1, pages_decorations)
            .unwrap();
        let mut pages = crate::pages::PagesEditor::from_bytes(&pages.to_bytes().unwrap()).unwrap();
        assert_eq!(
            pages.table_cell_text_alignment(pages_table, 1, 1).unwrap(),
            TextAlignment::Justified
        );
        assert_eq!(
            pages.table_cell_text_style(pages_table, 1, 1).unwrap(),
            pages_style
        );
        assert_eq!(
            pages.table_cell_text_font(pages_table, 1, 1).unwrap(),
            pages_font
        );
        assert_eq!(
            pages.table_cell_text_color(pages_table, 1, 1).unwrap(),
            pages_color
        );
        assert_eq!(
            pages
                .table_cell_text_decorations(pages_table, 1, 1)
                .unwrap(),
            pages_decorations
        );
        assert!(
            pages
                .reset_table_cell_text_decorations(pages_table, 1, 1)
                .unwrap()
        );
        assert!(
            pages
                .reset_table_cell_text_color(pages_table, 1, 1)
                .unwrap()
        );
        assert!(pages.reset_table_cell_text_font(pages_table, 1, 1).unwrap());
        assert!(
            pages
                .reset_table_cell_text_style(pages_table, 1, 1)
                .unwrap()
        );

        let keynote_style =
            TextStyle::new(TextPointSize::from_points(19.0).unwrap()).with_bold(true);
        let keynote_font = TextFont::named("Menlo-Regular").unwrap();
        let keynote_color = test_color();
        let keynote_decorations =
            TextDecorations::new(TextUnderline::None, TextStrikethrough::Single);
        let mut keynote = KeynoteDocumentBuilder::new()
            .title("Aligned")
            .build()
            .unwrap();
        let table = keynote
            .add_slide_table(
                0,
                "Aligned",
                2,
                2,
                DrawablePoint { x: 100.0, y: 100.0 },
                DrawableSize {
                    width: 400.0,
                    height: 200.0,
                },
            )
            .unwrap();
        keynote
            .set_slide_table_cell_text_alignment(
                0,
                table.model_object_id,
                1,
                1,
                TextAlignment::Left,
            )
            .unwrap();
        keynote
            .set_slide_table_cell_text_style(0, table.model_object_id, 1, 1, keynote_style)
            .unwrap();
        keynote
            .set_slide_table_cell_text_font(0, table.model_object_id, 1, 1, keynote_font.clone())
            .unwrap();
        keynote
            .set_slide_table_cell_text_color(0, table.model_object_id, 1, 1, keynote_color)
            .unwrap();
        keynote
            .set_slide_table_cell_text_decorations(
                0,
                table.model_object_id,
                1,
                1,
                keynote_decorations,
            )
            .unwrap();
        let mut keynote =
            crate::keynote::KeynoteEditor::from_bytes(&keynote.to_bytes().unwrap()).unwrap();
        assert_eq!(
            keynote
                .slide_table_cell_text_alignment(0, table.model_object_id, 1, 1)
                .unwrap(),
            TextAlignment::Left
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_style(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_style
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_font(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_font
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_color(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_color
        );
        assert_eq!(
            keynote
                .slide_table_cell_text_decorations(0, table.model_object_id, 1, 1)
                .unwrap(),
            keynote_decorations
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_decorations(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_color(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_font(0, table.model_object_id, 1, 1)
                .unwrap()
        );
        assert!(
            keynote
                .reset_slide_table_cell_text_style(0, table.model_object_id, 1, 1)
                .unwrap()
        );
    }

    #[test]
    fn invalid_paragraph_style_coordinate_is_transactional() {
        let mut editor = NumbersDocumentBuilder::new()
            .table_dimensions(2, 2)
            .build()
            .unwrap();
        let table_id = editor.tables().unwrap()[0].object_id;
        let before = editor.to_bytes().unwrap();
        assert!(
            editor
                .set_table_cell_text_alignment(table_id, 2, 1, TextAlignment::Center)
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_style(table_id, 1, 2, TextStyle::default())
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_font(
                    table_id,
                    1,
                    2,
                    TextFont::named("CourierNewPSMT").unwrap(),
                )
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_color(table_id, 1, 2, test_color())
                .is_err()
        );
        assert!(
            editor
                .set_table_cell_text_decorations(
                    table_id,
                    1,
                    2,
                    TextDecorations::new(TextUnderline::Single, TextStrikethrough::Single),
                )
                .is_err()
        );
        assert_eq!(editor.to_bytes().unwrap(), before);
    }
}
