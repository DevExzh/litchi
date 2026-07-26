//! Native table construction for workbooks that no longer contain a template table.

use super::*;
use crate::IWorkThemeArchive;
use crate::numbers::table_uid_map::COLUMN_ROW_UID_MAP_MESSAGE_TYPE;

const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TABLE_MODEL_MESSAGE_TYPES: &[u32] = &[TABLE_INFO_MESSAGE_TYPE, TABLE_MODEL_MESSAGE_TYPE];
const TILE_MESSAGE_TYPE: u32 = 6_002;
const DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const TABLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
const THEME_MESSAGE_TYPE: u32 = 12_009;
const IWA_OBJECT_VERSION: &[u32] = &[1, 0, 5];
const DEFAULT_TABLE_WIDTH_POINTS: f32 = 490.0;
const DEFAULT_TABLE_HEIGHT_POINTS: f32 = 200.0;
const DEFAULT_ROW_HEIGHT_POINTS: f64 = 20.0;
const DEFAULT_COLUMN_WIDTH_POINTS: f64 = 98.0;
const DEFAULT_TABLE_TITLE_HEIGHT_POINTS: f64 = 0.0;
const DEFAULT_HEADER_COUNT: u32 = 1;
const DEFAULT_TILE_SIZE: u32 = 256;
const TABLE_GEOMETRY_FLAGS: u32 = 3;
const DATA_STORE_ROW_BUCKET_HASH_FUNCTION: u32 = 1;
const DATA_STORE_VERSION: u32 = 4;
const TILE_STORAGE_VERSION: u32 = 5;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct TableObjectIds {
    info: u64,
    model: u64,
    tile: u64,
    row_headers: u64,
    column_headers: u64,
    string_list: u64,
    style_list: u64,
    formula_list: u64,
    format_list: u64,
    control_cell_spec_list: u64,
    uid_map: u64,
    stroke_sidecar: u64,
}

impl TableObjectIds {
    const COUNT: usize = 12;

    fn allocate(package: &IWorkPackage) -> Result<Self> {
        let mut next = next_object_identifier(package)?;
        Ok(Self {
            info: take_identifier(&mut next)?,
            model: take_identifier(&mut next)?,
            tile: take_identifier(&mut next)?,
            row_headers: take_identifier(&mut next)?,
            column_headers: take_identifier(&mut next)?,
            string_list: take_identifier(&mut next)?,
            style_list: take_identifier(&mut next)?,
            formula_list: take_identifier(&mut next)?,
            format_list: take_identifier(&mut next)?,
            control_cell_spec_list: take_identifier(&mut next)?,
            uid_map: take_identifier(&mut next)?,
            stroke_sidecar: take_identifier(&mut next)?,
        })
    }

    fn all(self) -> [u64; Self::COUNT] {
        [
            self.info,
            self.model,
            self.tile,
            self.row_headers,
            self.column_headers,
            self.string_list,
            self.style_list,
            self.formula_list,
            self.format_list,
            self.control_cell_spec_list,
            self.uid_map,
            self.stroke_sidecar,
        ]
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct TableStyleTemplate {
    preset: u64,
    preset_index: u32,
    table: u64,
    body_text: u64,
    header_row_text: u64,
    header_column_text: u64,
    footer_row_text: u64,
    body_cell: u64,
    header_row_cell: u64,
    header_column_cell: u64,
    footer_row_cell: u64,
    table_name: Option<u64>,
    table_name_shape: Option<u64>,
}

impl TableStyleTemplate {
    fn referenced_objects(self) -> impl Iterator<Item = u64> {
        [
            Some(self.preset),
            Some(self.table),
            Some(self.body_text),
            Some(self.header_row_text),
            Some(self.header_column_text),
            Some(self.footer_row_text),
            Some(self.body_cell),
            Some(self.header_row_cell),
            Some(self.header_column_cell),
            Some(self.footer_row_cell),
            self.table_name,
            self.table_name_shape,
        ]
        .into_iter()
        .flatten()
    }
}

/// Install an independent table graph using the workbook theme's first preset.
///
/// This path is used only when no attached table remains to clone. It builds
/// native storage directly, registers the new objects in their owning sheet
/// component, and attaches an empty CalculationEngine formula owner when the
/// workbook has a calculation component.
pub(super) fn bootstrap_empty_table_graph(
    package: &mut IWorkPackage,
    sheet_id: u64,
    name: &str,
    rows: usize,
    columns: usize,
) -> Result<table_create::EmptyTableGraph> {
    validate_name(name, "table")?;
    let (rows, columns) = validate_table_dimensions(rows, columns)?;
    let style = theme_table_style(package)?;
    let locations = object_locations(package)?;
    let sheet_archive = locations
        .get(&sheet_id)
        .cloned()
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers sheet {sheet_id} is missing")))?;
    let ids = TableObjectIds::allocate(package)?;
    let existing_table_uuids = package_table_uuids(package)?;
    let existing_table_uuid_refs = existing_table_uuids
        .iter()
        .map(String::as_str)
        .collect::<HashSet<_>>();
    let table_uuid = allocate_table_uuid(ids.model, &existing_table_uuid_refs);
    let objects = table_objects(ids, sheet_id, name, rows, columns, style, &table_uuid)?;

    package.update_archive(&sheet_archive, |archive| {
        for object in objects {
            archive.insert_object(object)?;
        }
        Ok(())
    })?;
    if let Some(component) = component_identifier_for_entry(package, &sheet_archive)? {
        add_component_object_uuids(package, component, &ids.all())?;
    }
    register_style_references(package, &sheet_archive, &locations, style)?;
    set_package_last_object_identifier(package, ids.stroke_sidecar)?;

    if let Some((calculation_entry, _)) =
        create_empty_table_formula_graph(package, ids.info, &table_uuid)?
    {
        register_numbers_component_reference(
            package,
            &calculation_entry,
            &sheet_archive,
            ids.info,
        )?;
    }

    Ok(table_create::EmptyTableGraph {
        info_object_id: ids.info,
        model_object_id: ids.model,
    })
}

fn theme_table_style(package: &IWorkPackage) -> Result<TableStyleTemplate> {
    let document = numbers_document(package)?;
    let theme_id = document.theme.identifier;
    let locations = object_locations(package)?;
    let theme_archive_name = locations
        .get(&theme_id)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers theme {theme_id} is missing")))?;
    let theme_archive = package.archive(theme_archive_name)?;
    let theme_object = theme_archive.object(theme_id).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers theme object {theme_id} is missing"))
    })?;
    let theme_messages = theme_object
        .messages
        .iter()
        .filter(|message| message.type_ == THEME_MESSAGE_TYPE)
        .collect::<Vec<_>>();
    let [theme_message] = theme_messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers theme {theme_id} must contain exactly one native payload"
        )));
    };
    let theme = IWorkThemeArchive::decode(theme_message.data.as_slice())?;
    let (theme_preset_index, preset_id) = theme
        .extensions
        .table
        .as_ref()
        .and_then(|presets| {
            presets
                .table_style_presets
                .iter()
                .enumerate()
                .find(|(_, reference)| reference.identifier != 0)
        })
        .map(|(index, reference)| (index, reference.identifier))
        .ok_or_else(|| {
            Error::InvalidFormat("Numbers theme has no native table style preset".to_owned())
        })?;
    let preset = decode_object_message::<tst::TableStylePresetArchive>(
        package,
        preset_id,
        TABLE_PRESET_MESSAGE_TYPE,
        "table preset",
    )?;
    let theme_preset_index = u32::try_from(theme_preset_index)
        .map_err(|_| Error::InvalidFormat("Numbers table preset index exceeds u32".to_owned()))?;
    let network_id = preset
        .style_network
        .as_ref()
        .map(|reference| reference.identifier)
        .filter(|identifier| *identifier != 0)
        .ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table preset {preset_id} has no style network"
            ))
        })?;
    let network = decode_object_message::<tst::TableStyleNetworkArchive>(
        package,
        network_id,
        TABLE_STYLE_NETWORK_MESSAGE_TYPE,
        "table style network",
    )?;
    let preset_index = preset
        .index
        .map(u32::try_from)
        .transpose()
        .map_err(|_| {
            Error::InvalidFormat(format!(
                "Numbers table preset {preset_id} has a negative index"
            ))
        })?
        .unwrap_or(theme_preset_index);
    let style = TableStyleTemplate {
        preset: preset_id,
        preset_index,
        table: network.table_style.identifier,
        body_text: network.body_text_style.identifier,
        header_row_text: network.header_row_text_style.identifier,
        header_column_text: network.header_column_text_style.identifier,
        footer_row_text: network.footer_row_text_style.identifier,
        body_cell: network.body_cell_style.identifier,
        header_row_cell: network.header_row_style.identifier,
        header_column_cell: network.header_column_style.identifier,
        footer_row_cell: network.footer_row_style.identifier,
        table_name: network
            .table_name_style
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0),
        table_name_shape: network
            .table_name_shape_style
            .map(|reference| reference.identifier)
            .filter(|identifier| *identifier != 0),
    };
    for identifier in style.referenced_objects() {
        if identifier == 0 || !locations.contains_key(&identifier) {
            return Err(Error::InvalidFormat(format!(
                "Numbers table style references missing object {identifier}"
            )));
        }
    }
    Ok(style)
}

fn package_table_uuids(package: &IWorkPackage) -> Result<HashSet<String>> {
    let mut uuids = HashSet::new();
    for entry in package.iwa_entry_names() {
        for object in package.archive(entry)?.objects {
            for message in object.messages {
                if !TABLE_MODEL_MESSAGE_TYPES.contains(&message.type_) {
                    continue;
                }
                let Ok(model) = TableModelArchive::decode(message.data.as_slice()) else {
                    continue;
                };
                if !model.table_id.is_empty() && !uuids.insert(model.table_id.clone()) {
                    return Err(Error::InvalidFormat(format!(
                        "Numbers table UUID {:?} is duplicated",
                        model.table_id
                    )));
                }
            }
        }
    }
    Ok(uuids)
}

fn decode_object_message<M: Message + Default>(
    package: &IWorkPackage,
    identifier: u64,
    message_type: u32,
    label: &str,
) -> Result<M> {
    let locations = object_locations(package)?;
    let archive_name = locations
        .get(&identifier)
        .ok_or_else(|| Error::InvalidFormat(format!("Numbers {label} {identifier} is missing")))?;
    let archive = package.archive(archive_name)?;
    let object = archive.object(identifier).ok_or_else(|| {
        Error::InvalidFormat(format!("Numbers {label} object {identifier} is missing"))
    })?;
    let messages = object
        .messages
        .iter()
        .filter(|message| message.type_ == message_type)
        .collect::<Vec<_>>();
    let [message] = messages.as_slice() else {
        return Err(Error::InvalidFormat(format!(
            "Numbers {label} {identifier} must contain exactly one native payload"
        )));
    };
    Ok(M::decode(message.data.as_slice())?)
}

fn table_objects(
    ids: TableObjectIds,
    sheet_id: u64,
    name: &str,
    rows: u32,
    columns: u32,
    style: TableStyleTemplate,
    table_uuid: &str,
) -> Result<Vec<ArchiveObject>> {
    let table_info = tst::TableInfoArchive {
        super_: tsd::DrawableArchive {
            geometry: Some(tsd::GeometryArchive {
                position: Some(tsp::Point { x: 0.0, y: 0.0 }),
                size: Some(tsp::Size {
                    width: DEFAULT_TABLE_WIDTH_POINTS,
                    height: DEFAULT_TABLE_HEIGHT_POINTS,
                }),
                flags: Some(TABLE_GEOMETRY_FLAGS),
                angle: Some(0.0),
            }),
            parent: Some(reference(sheet_id)),
            locked: Some(false),
            aspect_ratio_locked: Some(false),
            title_hidden: Some(true),
            caption_hidden: Some(true),
            ..Default::default()
        },
        table_model: reference(ids.model),
        formula_coord_space: Some(0),
        ..Default::default()
    };
    let model = TableModelArchive {
        table_id: table_uuid.to_owned(),
        table_style: reference(style.table),
        body_text_style: reference(style.body_text),
        header_row_text_style: reference(style.header_row_text),
        header_column_text_style: reference(style.header_column_text),
        footer_row_text_style: reference(style.footer_row_text),
        body_cell_style: reference(style.body_cell),
        header_row_style: reference(style.header_row_cell),
        header_column_style: reference(style.header_column_cell),
        footer_row_style: reference(style.footer_row_cell),
        table_name_style: style.table_name.map(reference),
        table_name_shape_style: style.table_name_shape.map(reference),
        table_style_preset: Some(reference(style.preset)),
        preset_index: Some(style.preset_index),
        base_data_store: tst::DataStore {
            row_headers: tst::HeaderStorage {
                bucket_hash_function: DATA_STORE_ROW_BUCKET_HASH_FUNCTION,
                buckets: vec![reference(ids.row_headers)],
            },
            column_headers: reference(ids.column_headers),
            tiles: tst::TileStorage {
                tiles: vec![tst::tile_storage::Tile {
                    tileid: 0,
                    tile: reference(ids.tile),
                }],
                tile_size: Some(DEFAULT_TILE_SIZE),
                ..Default::default()
            },
            string_table: reference(ids.string_list),
            style_table: reference(ids.style_list),
            formula_table: reference(ids.formula_list),
            format_table_pre_bnc: reference(ids.format_list),
            format_table: Some(reference(ids.format_list)),
            control_cell_spec_table: Some(reference(ids.control_cell_spec_list)),
            next_row_strip_id: 1,
            next_column_strip_id: 0,
            row_tile_tree: tst::TableRbTree {
                nodes: vec![tst::table_rb_tree::Node { key: 0, value: 0 }],
            },
            column_tile_tree: tst::TableRbTree::default(),
            storage_version_pre_bnc: Some(DATA_STORE_VERSION),
            ..Default::default()
        },
        number_of_rows: rows,
        number_of_columns: columns,
        table_name: name.to_owned(),
        table_name_enabled: Some(false),
        table_name_height: (style.table_name.is_some() && style.table_name_shape.is_some())
            .then_some(DEFAULT_TABLE_TITLE_HEIGHT_POINTS),
        number_of_header_rows: Some(DEFAULT_HEADER_COUNT),
        number_of_header_columns: Some(DEFAULT_HEADER_COUNT),
        header_rows_frozen: Some(true),
        header_columns_frozen: Some(true),
        default_row_height: DEFAULT_ROW_HEIGHT_POINTS,
        default_column_width: DEFAULT_COLUMN_WIDTH_POINTS,
        repeating_header_rows_enabled: Some(true),
        repeating_header_columns_enabled: Some(true),
        style_apply_clears_all: Some(false),
        base_column_row_uids: Some(reference(ids.uid_map)),
        stroke_sidecar: Some(reference(ids.stroke_sidecar)),
        ..Default::default()
    };

    let mut objects = Vec::with_capacity(TableObjectIds::COUNT);
    objects.push(table_object(
        ids.info,
        TABLE_INFO_MESSAGE_TYPE,
        table_info,
        &[sheet_id, ids.model],
    )?);
    objects.push(table_object(
        ids.model,
        TABLE_MODEL_MESSAGE_TYPE,
        model.clone(),
        &table_model_references(&model),
    )?);
    objects.push(table_object(
        ids.tile,
        TILE_MESSAGE_TYPE,
        tst::Tile {
            max_column: columns - 1,
            max_row: rows - 1,
            num_cells: 0,
            numrows: 0,
            row_infos: Vec::new(),
            storage_version: Some(TILE_STORAGE_VERSION),
            last_saved_in_bnc: Some(true),
            should_use_wide_rows: None,
        },
        &[],
    )?);
    for identifier in [ids.row_headers, ids.column_headers] {
        objects.push(table_object(
            identifier,
            HEADER_BUCKET_MESSAGE_TYPE,
            tst::HeaderStorageBucket {
                bucket_hash_function: DATA_STORE_ROW_BUCKET_HASH_FUNCTION,
                headers: Vec::new(),
            },
            &[],
        )?);
    }
    for (identifier, list_type) in [
        (ids.string_list, tst::table_data_list::ListType::String),
        (ids.style_list, tst::table_data_list::ListType::Style),
        (ids.formula_list, tst::table_data_list::ListType::Formula),
        (ids.format_list, tst::table_data_list::ListType::Format),
        (
            ids.control_cell_spec_list,
            tst::table_data_list::ListType::ControlCellSpec,
        ),
    ] {
        objects.push(table_object(
            identifier,
            DATA_LIST_MESSAGE_TYPE,
            tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 1,
                entries: Vec::new(),
                segments: Vec::new(),
                is_new_for_bnc: Some(true),
            },
            &[],
        )?);
    }
    objects.push(table_object(
        ids.uid_map,
        COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
        empty_uid_map(rows, columns, ids.model)?,
        &[],
    )?);
    objects.push(table_object(
        ids.stroke_sidecar,
        STROKE_SIDECAR_MESSAGE_TYPE,
        tst::StrokeSidecarArchive {
            row_count: Some(rows),
            column_count: Some(columns),
            ..Default::default()
        },
        &[],
    )?);
    Ok(objects)
}

fn register_style_references(
    package: &mut IWorkPackage,
    source_archive: &str,
    locations: &HashMap<u64, String>,
    style: TableStyleTemplate,
) -> Result<()> {
    let mut seen = HashSet::new();
    for identifier in style.referenced_objects() {
        if !seen.insert(identifier) {
            continue;
        }
        let target_archive = locations.get(&identifier).ok_or_else(|| {
            Error::InvalidFormat(format!(
                "Numbers table style object {identifier} is missing"
            ))
        })?;
        register_numbers_component_reference(package, source_archive, target_archive, identifier)?;
    }
    Ok(())
}

fn table_object(
    identifier: u64,
    message_type: u32,
    message: impl Message,
    references: &[u64],
) -> Result<ArchiveObject> {
    let mut object = ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data: message.encode_to_vec(),
        }],
    )?;
    object.archive_info.message_infos[0].versions = IWA_OBJECT_VERSION.to_vec();
    object.archive_info.message_infos[0].object_references = references.to_vec();
    Ok(object)
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}
