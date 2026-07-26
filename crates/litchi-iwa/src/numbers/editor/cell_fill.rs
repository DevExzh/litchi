//! Copy-on-write table-cell fill styles shared by Numbers, Pages, and Keynote.

use prost::Message;

use super::*;
use crate::data_reference_registry::{
    add_component_data_reference, remove_component_data_reference,
};
use crate::package_metadata::{
    add_component_object_uuids, release_package_identifier_suffix,
    remove_component_external_references_to_object, remove_component_object_uuids,
};
use crate::protobuf::tss;
use crate::shapes::{
    ShapeFill, fill_from_native, fill_to_native, image_data_identifier,
    remove_orphaned_image_asset, validate_image_asset,
};

const CELL_STYLE_MESSAGE_TYPE: u32 = 6_004;
const CELL_STYLE_MESSAGE_VERSION: [u32; 3] = [1, 0, 5];
const MAX_STYLE_INHERITANCE_DEPTH: usize = 64;

struct CellStyleLocation {
    archive_name: String,
    style: tst::CellStyleArchive,
    raw: Vec<u8>,
}

pub(super) fn cell_fill(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<ShapeFill> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    validate_coordinate(&descriptor, row, column)?;
    let locations = storage::object_locations(package)?;
    let style_id = local_style_id(package, &descriptor, &locations, row, column)?
        .unwrap_or_else(|| base_style_id(&descriptor, row, column));
    effective_fill(package, &locations, style_id)
}

pub(super) fn set_cell_fill(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    fill: &ShapeFill,
) -> Result<()> {
    validate_image_asset(package, fill)?;
    if &cell_fill(package, table_id, row, column)? == fill {
        return Ok(());
    }

    let mut staged = package.clone();
    table_sparse_storage::ensure_attached_cell_storage(&mut staged, table_id, row, column)?;
    let location = model::locate_attached_cell(&staged, table_id, row, column)?;
    let style_table_id = location
        .descriptor
        .model
        .base_data_store
        .style_table
        .identifier;
    let old_key = read_bnc(&staged, &location, column)?.style_identifier();
    let resolved = storage::resolve_table_data_list(
        &staged,
        &location.object_locations,
        style_table_id,
        tst::table_data_list::ListType::Style,
    )?;
    let old_entry = style_entry(&resolved, old_key)?;
    let parent_style_id = old_entry
        .and_then(|entry| entry.entry.reference.as_ref())
        .map_or_else(
            || base_style_id(&location.descriptor, row, column),
            |reference| reference.identifier,
        );

    if let Some(entry) = old_entry
        && entry.entry.refcount == 1
        && direct_fill_only_style(&staged, &location.object_locations, parent_style_id)?.is_some()
    {
        replace_fill_only_style(&mut staged, parent_style_id, fill)?;
        verify_fill(&staged, table_id, row, column, fill)?;
        *package = staged;
        return Ok(());
    }

    let new_style_id = crate::package_metadata::next_object_identifier(&staged)?;
    let parent = cell_style_location(&staged, &location.object_locations, parent_style_id)?;
    let stylesheet_id = parent
        .style
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "iWork cell style {parent_style_id} has no stylesheet"
            ))
        })?;
    let style_object = cell_fill_variation(new_style_id, parent_style_id, stylesheet_id, fill)?;
    crate::shapes::insert_style_variation(
        &mut staged,
        &parent.archive_name,
        stylesheet_id,
        parent_style_id,
        new_style_id,
        style_object,
    )?;
    register_style_object(&mut staged, &parent.archive_name, new_style_id, fill)?;
    let new_key = insert_style_entry(
        &mut staged,
        &location.object_locations,
        style_table_id,
        new_style_id,
    )?;
    write_style_key(&mut staged, &location, row, column, Some(new_key))?;
    if let Some(entry) = old_entry {
        storage::decrement_table_data_list_entry(
            &mut staged,
            &location.object_locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )?;
    }
    crate::package_metadata::set_package_last_object_identifier(&mut staged, new_style_id)?;
    verify_fill(&staged, table_id, row, column, fill)?;
    *package = staged;
    Ok(())
}

pub(super) fn reset_cell_fill(
    package: &mut IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
) -> Result<bool> {
    let descriptor = model::attached_table_descriptor(package, table_id)?;
    validate_coordinate(&descriptor, row, column)?;
    let locations = storage::object_locations(package)?;
    let Some(key) = cell_style_key(package, &descriptor, &locations, row, column)? else {
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
        .ok_or_else(|| Error::InvalidFormat(format!("iWork cell style table has no key {key}")))?;
    let style_id = entry
        .entry
        .reference
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork cell style key {key} has no reference"))
        })?;
    let style = cell_style_location(package, &locations, style_id)?;
    let Some(direct_fill) = style
        .style
        .cell_properties
        .as_ref()
        .and_then(|properties| properties.cell_fill.as_ref())
    else {
        return Ok(false);
    };
    let old_data_identifier = image_data_identifier(&fill_from_native(direct_fill)?);
    let base_id = base_style_id(&descriptor, row, column);
    let parent_id = style
        .style
        .super_
        .parent
        .as_ref()
        .map_or(base_id, |reference| reference.identifier);
    let inherited = effective_fill(package, &locations, parent_id)?;

    if entry.entry.refcount == 1
        && direct_fill_only_style(package, &locations, style_id)?.is_some()
        && parent_id == base_id
    {
        let mut staged = package.clone();
        let location = model::locate_attached_cell(&staged, table_id, row, column)?;
        write_style_key(&mut staged, &location, row, column, None)?;
        let removed = storage::decrement_table_data_list_entry(
            &mut staged,
            &locations,
            &resolved,
            entry,
            tst::table_data_list::ListType::Style,
        )?;
        if removed && !style_has_children(&staged, &style.archive_name, style_id)? {
            remove_disposable_style(
                &mut staged,
                &style.archive_name,
                style_id,
                parent_id,
                old_data_identifier,
            )?;
        }
        verify_fill(&staged, table_id, row, column, &inherited)?;
        remove_orphaned_image_asset(&mut staged, old_data_identifier)?;
        *package = staged;
        return Ok(true);
    }

    // Shared or richer native styles cannot be mutated for one cell. A private
    // child restores the parent's effective fill while inheriting every other
    // property from the current style.
    set_cell_fill(package, table_id, row, column, &inherited)?;
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
        model.header_row_style.identifier
    } else if row
        >= model
            .number_of_rows
            .saturating_sub(model.number_of_footer_rows.unwrap_or(0)) as usize
    {
        model.footer_row_style.identifier
    } else if column < model.number_of_header_columns.unwrap_or(0) as usize {
        model.header_column_style.identifier
    } else {
        model.body_cell_style.identifier
    }
}

fn cell_style_key(
    package: &IWorkPackage,
    descriptor: &model::TableDescriptor,
    _locations: &HashMap<u64, String>,
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
    .map(|data| BncCell::parse(&data).map(|cell| cell.style_identifier()))
    .transpose()
    .map(Option::flatten)
}

fn local_style_id(
    package: &IWorkPackage,
    descriptor: &model::TableDescriptor,
    locations: &HashMap<u64, String>,
    row: usize,
    column: usize,
) -> Result<Option<u64>> {
    let key = cell_style_key(package, descriptor, locations, row, column)?;
    let Some(key) = key else {
        return Ok(None);
    };
    let resolved = storage::resolve_table_data_list(
        package,
        locations,
        descriptor.model.base_data_store.style_table.identifier,
        tst::table_data_list::ListType::Style,
    )?;
    let entry = style_entry(&resolved, Some(key))?
        .ok_or_else(|| Error::InvalidFormat(format!("iWork cell style table has no key {key}")))?;
    entry
        .entry
        .reference
        .as_ref()
        .map(|reference| Some(reference.identifier))
        .ok_or_else(|| Error::InvalidFormat(format!("iWork cell style key {key} has no reference")))
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
            "iWork cell style table repeats key {key}"
        )));
    }
    Ok(entry)
}

fn cell_style_location(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    style_id: u64,
) -> Result<CellStyleLocation> {
    let archive_name = locations
        .get(&style_id)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork cell style {style_id} is missing")))?
        .clone();
    let archive = package.archive(&archive_name)?;
    let object = archive
        .object(style_id)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork cell style {style_id} is missing")))?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == CELL_STYLE_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork cell style {style_id} must have exactly one CellStyle payload"
        )));
    };
    Ok(CellStyleLocation {
        archive_name,
        style: tst::CellStyleArchive::decode(message.data.as_slice())?,
        raw: message.data.clone(),
    })
}

fn effective_fill(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    first_style_id: u64,
) -> Result<ShapeFill> {
    let mut style_id = first_style_id;
    let mut seen = HashSet::new();
    for _ in 0..MAX_STYLE_INHERITANCE_DEPTH {
        if !seen.insert(style_id) {
            return Err(Error::InvalidFormat(format!(
                "iWork cell-style inheritance cycles at {style_id}"
            )));
        }
        let located = cell_style_location(package, locations, style_id)?;
        if let Some(fill) = located
            .style
            .cell_properties
            .as_ref()
            .and_then(|properties| properties.cell_fill.as_ref())
        {
            return fill_from_native(fill);
        }
        let Some(parent) = located.style.super_.parent else {
            return Ok(ShapeFill::None);
        };
        style_id = parent.identifier;
    }
    Err(Error::InvalidFormat(format!(
        "iWork cell style {first_style_id} exceeds {MAX_STYLE_INHERITANCE_DEPTH} inheritance levels"
    )))
}

fn direct_fill_only_style(
    package: &IWorkPackage,
    locations: &HashMap<u64, String>,
    style_id: u64,
) -> Result<Option<ShapeFill>> {
    let located = cell_style_location(package, locations, style_id)?;
    let style = &located.style;
    let Some(properties) = style.cell_properties.as_ref() else {
        return Ok(None);
    };
    let Some(fill) = properties.cell_fill.as_ref() else {
        return Ok(None);
    };
    let semantic = style.override_count == Some(1)
        && style.super_.name.is_none()
        && style.super_.style_identifier.is_none()
        && style.super_.parent.is_some()
        && style.super_.is_variation == Some(true)
        && style.super_.stylesheet.is_some()
        && properties.text_wrap.is_none()
        && properties.deprecated_top_stroke.is_none()
        && properties.deprecated_right_stroke.is_none()
        && properties.deprecated_bottom_stroke.is_none()
        && properties.deprecated_left_stroke.is_none()
        && properties.vertical_alignment.is_none()
        && properties.padding.is_none()
        && properties.top_stroke.is_none()
        && properties.right_stroke.is_none()
        && properties.bottom_stroke.is_none()
        && properties.left_stroke.is_none();
    if !semantic {
        return Ok(None);
    }
    let exact = exact_fields(&located.raw, &[1, 10, 11])?
        && exact_fields(required_payload(&located.raw, 1)?, &[3, 4, 5])?
        && exact_fields(required_payload(&located.raw, 11)?, &[1])?;
    if !exact {
        return Ok(None);
    }
    fill_from_native(fill).map(Some)
}

fn exact_fields(data: &[u8], expected: &[u32]) -> Result<bool> {
    let mut actual = crate::wire::parse_wire_fields(data)?
        .into_iter()
        .map(|field| field.number)
        .collect::<Vec<_>>();
    let mut expected = expected.to_vec();
    actual.sort_unstable();
    expected.sort_unstable();
    Ok(actual == expected)
}

fn required_payload(data: &[u8], field: u32) -> Result<&[u8]> {
    let payloads = crate::wire::repeated_length_delimited_payloads(data, field)?;
    let [payload] = payloads.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "iWork cell style field {field} must occur exactly once"
        )));
    };
    Ok(payload)
}

fn cell_fill_variation(
    identifier: u64,
    parent_style_id: u64,
    stylesheet_id: u64,
    fill: &ShapeFill,
) -> Result<ArchiveObject> {
    let native = fill_to_native(fill);
    let data = tst::CellStyleArchive {
        super_: tss::StyleArchive {
            parent: Some(reference(parent_style_id)),
            is_variation: Some(true),
            stylesheet: Some(reference(stylesheet_id)),
            ..Default::default()
        },
        override_count: Some(1),
        cell_properties: Some(tst::CellStylePropertiesArchive {
            cell_fill: Some(native),
            ..Default::default()
        }),
    }
    .encode_to_vec();
    tst::CellStyleArchive::decode(data.as_slice())?;
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: CELL_STYLE_MESSAGE_TYPE,
            data,
        }],
    )?;
    let info = &mut object.archive_info.message_infos[0];
    info.versions = CELL_STYLE_MESSAGE_VERSION.to_vec();
    info.object_references.push(parent_style_id);
    if let Some(data_identifier) = image_data_identifier(fill) {
        info.data_references.push(data_identifier);
    }
    Ok(object)
}

fn replace_fill_only_style(
    package: &mut IWorkPackage,
    style_id: u64,
    fill: &ShapeFill,
) -> Result<()> {
    let locations = storage::object_locations(package)?;
    let located = cell_style_location(package, &locations, style_id)?;
    let parent_style_id = located
        .style
        .super_
        .parent
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork cell style {style_id} has no parent"))
        })?;
    let stylesheet_id = located
        .style
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork cell style {style_id} has no stylesheet"))
        })?;
    let old_fill = direct_fill_only_style(package, &locations, style_id)?.ok_or_else(|| {
        Error::InvalidFormat(format!("iWork cell style {style_id} is not fill-only"))
    })?;
    let old_data = image_data_identifier(&old_fill);
    let mut replacement = cell_fill_variation(style_id, parent_style_id, stylesheet_id, fill)?;
    let replacement_message = replacement.messages.pop().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "replacement iWork cell style {style_id} has no payload"
        ))
    })?;
    let replacement_data_references = replacement.archive_info.message_infos[0]
        .data_references
        .clone();
    package.update_archive(&located.archive_name, |archive| {
        let object = archive.object_mut(style_id).ok_or_else(|| {
            Error::InvalidFormat(format!("iWork cell style {style_id} is missing"))
        })?;
        let indexes = object
            .messages
            .iter()
            .enumerate()
            .filter(|(_, message)| message.type_ == CELL_STYLE_MESSAGE_TYPE)
            .map(|(index, _)| index)
            .collect::<Vec<_>>();
        let [index] = indexes.as_slice() else {
            return Err(Error::InvalidFormat(format!(
                "iWork cell style {style_id} must have exactly one CellStyle payload"
            )));
        };
        object.replace_message(*index, replacement_message)?;
        object.archive_info.message_infos[*index].data_references = replacement_data_references;
        Ok(())
    })?;
    adjust_data_reference(
        package,
        &located.archive_name,
        style_id,
        old_data,
        image_data_identifier(fill),
    )?;
    remove_orphaned_image_asset(package, old_data)
}

fn register_style_object(
    package: &mut IWorkPackage,
    archive_name: &str,
    style_id: u64,
    fill: &ShapeFill,
) -> Result<()> {
    let component = crate::package_metadata::component_identifier_for_entry(package, archive_name)?;
    if let Some(component) = component {
        add_component_object_uuids(package, component, &[style_id])?;
        if let Some(data_identifier) = image_data_identifier(fill) {
            add_component_data_reference(package, component, data_identifier, style_id)?;
        }
    } else if image_data_identifier(fill).is_some() {
        return Err(Error::InvalidFormat(
            "iWork stylesheet has no component for an image-fill reference".to_owned(),
        ));
    }
    Ok(())
}

fn adjust_data_reference(
    package: &mut IWorkPackage,
    archive_name: &str,
    style_id: u64,
    old: Option<u64>,
    new: Option<u64>,
) -> Result<()> {
    if old == new {
        return Ok(());
    }
    let component = crate::package_metadata::component_identifier_for_entry(package, archive_name)?
        .ok_or_else(|| Error::InvalidFormat("iWork stylesheet has no component".to_owned()))?;
    if let Some(identifier) = old {
        remove_component_data_reference(package, component, identifier, style_id)?;
    }
    if let Some(identifier) = new {
        add_component_data_reference(package, component, identifier, style_id)?;
    }
    Ok(())
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
            .ok_or_else(|| Error::ParseError("iWork cell style key overflow".to_owned()))?;
        current.entries.push(tst::table_data_list::ListEntry {
            key,
            refcount: 1,
            reference: Some(reference(style_id)),
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
    let table_component =
        crate::package_metadata::component_identifier_for_entry(package, &resolved.table_archive)?;
    let current_locations = storage::object_locations(package)?;
    let style_archive = current_locations
        .get(&style_id)
        .ok_or_else(|| Error::InvalidFormat(format!("iWork cell style {style_id} is missing")))?;
    let style_component =
        crate::package_metadata::component_identifier_for_entry(package, style_archive)?;
    if let (Some(source), Some(target)) = (table_component, style_component)
        && source != target
    {
        crate::package_metadata::add_component_external_reference(
            package, source, target, style_id,
        )?;
    }
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

fn write_style_key(
    package: &mut IWorkPackage,
    location: &model::CellLocation,
    row: usize,
    column: usize,
    key: Option<u32>,
) -> Result<()> {
    let mut cell = read_bnc(package, location, column)?;
    cell.set_style_identifier(key);
    storage::set_encoded_cell_value(
        package,
        location.descriptor.object_id,
        row,
        column,
        EncodedValue::Raw(cell.encode()),
    )
}

fn style_has_children(package: &IWorkPackage, archive_name: &str, style_id: u64) -> Result<bool> {
    let archive = package.archive(archive_name)?;
    Ok(archive.objects.iter().any(|object| {
        object.messages.iter().any(|message| {
            message.type_ == CELL_STYLE_MESSAGE_TYPE
                && tst::CellStyleArchive::decode(message.data.as_slice()).is_ok_and(|style| {
                    style
                        .super_
                        .parent
                        .as_ref()
                        .is_some_and(|parent| parent.identifier == style_id)
                })
        })
    }))
}

fn remove_disposable_style(
    package: &mut IWorkPackage,
    archive_name: &str,
    style_id: u64,
    parent_id: u64,
    data_identifier: Option<u64>,
) -> Result<()> {
    let located = cell_style_location(package, &storage::object_locations(package)?, style_id)?;
    let stylesheet_id = located
        .style
        .super_
        .stylesheet
        .as_ref()
        .map(|reference| reference.identifier)
        .ok_or_else(|| {
            Error::InvalidFormat(format!("iWork cell style {style_id} has no stylesheet"))
        })?;
    if let Some(component) =
        crate::package_metadata::component_identifier_for_entry(package, archive_name)?
    {
        if let Some(identifier) = data_identifier {
            remove_component_data_reference(package, component, identifier, style_id)?;
        }
        remove_component_object_uuids(package, component, &[style_id])?;
        remove_component_external_references_to_object(package, component, style_id)?;
    }
    crate::shapes::remove_style_variation(
        package,
        archive_name,
        stylesheet_id,
        parent_id,
        style_id,
    )?;
    release_package_identifier_suffix(package, &[style_id])
}

fn verify_fill(
    package: &IWorkPackage,
    table_id: u64,
    row: usize,
    column: usize,
    expected: &ShapeFill,
) -> Result<()> {
    if &cell_fill(package, table_id, row, column)? != expected {
        return Err(Error::InvalidFormat(
            "iWork table-cell fill failed package validation".to_owned(),
        ));
    }
    Ok(())
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
