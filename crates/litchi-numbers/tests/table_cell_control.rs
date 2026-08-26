//! Exact-source integration coverage for the unified Numbers cell-control owner.
//!
//! The fixture contains one rooted five-row table.  Its rows exercise a
//! checkbox, star rating, slider, stepper, and Pop-Up Menu respectively.  A
//! second fixture mode leaves the cells scalar/automatic so creation and
//! allocator paths are tested against the same rooted graph.

#![allow(deprecated)]

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsce, tsd, tsk, tsp, tst};
use litchi_numbers::CellPosition;
use litchi_numbers::cell::data_format::control::transaction::Error as ControlError;
use litchi_numbers::cell::data_format::control::{
    CellControl, DisplayFormat, Range, Slider, Stepper,
};
use litchi_numbers::cell::data_format::{Checkbox, Number, PopUpMenu, StarRating};
use litchi_numbers::{
    Package, PackageReadOptions, PackageSemanticLimits, SheetSelector, TableSelector,
};
use litchi_numbers_wire::{BncCell, CellDataFormatKind};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const CALCULATION_MEMBER: &str = "Index/CalculationEngine.iwa";
const TILE_MEMBER: &str = "Index/Tables/Tile.iwa";
const FORMAT_MEMBER: &str = "Index/Tables/DataList-904498-2.iwa";
const CONTROL_MEMBER: &str = "Index/Tables/DataList-904499-2.iwa";
const POPUP_MEMBER: &str = "Index/Tables/Popup-905753.iwa";
const VIEW_STATE_MEMBER: &str = "Index/ViewState.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_TYPE: u32 = 11_006;
const TABLE_DATA_LIST_TYPE: u32 = 6_005;
const DOCUMENT_ID: u64 = 1;
const SHEET_ID: u64 = 2;
const TABLE_INFO_ID: u64 = 3;
const TABLE_MODEL_ID: u64 = 4;
const SIDECAR_ID: u64 = 5;
const TILE_ID: u64 = 6;
const FORMAT_LIST_ID: u64 = 7;
const CONTROL_LIST_ID: u64 = 8;
const CONTROL_MODEL_ID: u64 = 50;
const VIEW_STATE_ID: u64 = 300;
const METADATA_OBJECT_ID: u64 = 900;
const CALCULATION_COMPONENT_ID: u64 = 200;
const TILE_COMPONENT_ID: u64 = 201;
// Native anonymous format-list sidecars use the list object's identifier as
// the external target component identifier.
const FORMAT_COMPONENT_ID: u64 = FORMAT_LIST_ID;
const CONTROL_COMPONENT_ID: u64 = 203;
const POPUP_COMPONENT_ID: u64 = 204;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FixtureMode {
    Mixed,
    SharedCheckbox,
    Empty,
    ScalarSeed,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Corruption {
    WrongInteraction,
    MissingCellSpec,
    ZeroRefcount,
    DuplicateControlKey,
    SegmentedControlList,
    DuplicateControlList,
    ControlAggregateMissing,
    ControlAggregateWrong,
    ControlFieldInfoMissing,
    ControlFieldInfoWrong,
    ControlFieldInfoExtra,
    ControlKeyOverflow,
    MissingPopupModel,
    WrongPopupModelType,
    DeprecatedReferenceType,
    DeprecatedExternalReference,
    WrongBncKind,
    MissingFormatEntry,
    FormatRefcountUndercount,
    ControlRefcountUndercount,
    FormatRefcountOvercount,
    ControlRefcountOvercount,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MetadataCorruption {
    MissingControlUuid,
    VersionedOnlyControlUuid,
    DuplicateControlUuid,
    AmbiguousControlIdentifier,
    DataOwnerControlIdentifier,
    RootDataMapControlIdentifier,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SplitMetadataCorruption {
    MissingCalculationExternal,
    MissingTileExternal,
    MissingControlExternal,
    MissingPopupExternal,
    DuplicateCalculationExternal,
    DuplicatePopupExternal,
    ComponentAndObjectFormatExternal,
    ComponentAndObjectTileExternal,
    ComponentAndObjectControlExternal,
    VersionedCalculationExternal,
    VersionedPopupExternal,
    MissingFormatComponent,
    DuplicateFormatComponent,
    WrongCalculationLocator,
    WrongFormatLocator,
    WrongControlLocator,
    OpaquePopupInbound,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SplitModelReferenceCorruption {
    MissingFormatAggregate,
    DuplicateFormatAggregate,
    WrongFormatFieldType,
    DuplicateFormatField,
}

trait ExactBytes {
    fn exact_bytes(&self) -> Vec<u8>;
}

impl ExactBytes for Package {
    fn exact_bytes(&self) -> Vec<u8> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)
            .expect("in-memory package serialization cannot fail");
        bytes
    }
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(identifier: u64, type_: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage { type_, data }],
    )?)
}

fn uuid_entry(identifier: u64) -> tsp::ObjectUuidMapEntry {
    tsp::ObjectUuidMapEntry {
        identifier,
        uuid: tsp::Uuid {
            lower: identifier,
            upper: identifier.saturating_add(0x1000),
        },
    }
}

fn menu() -> TestResult<PopUpMenu> {
    Ok(
        PopUpMenu::new(["Low", "Medium", "High"])?.with_initial_selection(
            litchi_numbers::cell::data_format::pop_up_menu::InitialSelection::FirstItem,
        ),
    )
}

fn range(minimum: f64, maximum: f64, increment: f64) -> Range {
    Range::new(minimum, maximum, increment).expect("fixture range is finite and representable")
}

fn controls() -> TestResult<[CellControl; 5]> {
    Ok([
        CellControl::Checkbox(Checkbox),
        CellControl::StarRating(StarRating),
        CellControl::Slider(Slider::new(
            range(0.0, 100.0, 10.0),
            DisplayFormat::Number(Number::default()),
        )),
        CellControl::Stepper(Stepper::new(
            range(1.0, 10.0, 1.0),
            DisplayFormat::Number(Number::default()),
        )),
        CellControl::PopUpMenu(menu()?),
    ])
}

fn popup_item(value: &str) -> tsce::CellValueArchive {
    tsce::CellValueArchive {
        cell_value_type: tsce::cell_value_archive::CellValueType::StringType as i32,
        string_value: Some(tsce::StringCellValueArchive {
            value: value.to_owned(),
            format: tsk::FormatStructArchive {
                format_type: Some(260),
                ..Default::default()
            },
            format_is_implicit: None,
            format_is_explicit: Some(false),
            is_regex: Some(false),
            is_case_sensitive_regex: Some(false),
        }),
        ..Default::default()
    }
}

fn cell_for_control(
    format_identifier: u32,
    control_identifier: u32,
    kind: CellDataFormatKind,
    scalar: Option<f64>,
) -> TestResult<Vec<u8>> {
    let mut cell = BncCell::minimal();
    cell.set_data_format_identifier(format_identifier, kind, Some(control_identifier))?;
    if let Some(scalar) = scalar {
        cell.set_number(scalar)?;
    }
    Ok(cell.encode())
}

fn empty_cell() -> Vec<u8> {
    BncCell::minimal().encode()
}

fn number_cell(value: f64) -> TestResult<Vec<u8>> {
    let mut cell = BncCell::minimal();
    cell.set_data_format_identifier(1, CellDataFormatKind::NumberOrPercentage, None)?;
    cell.set_number(value)?;
    Ok(cell.encode())
}

fn pack_row(cells: Vec<Vec<u8>>) -> (Vec<u8>, Vec<u8>) {
    let mut storage = Vec::new();
    let mut offsets = Vec::with_capacity(cells.len() * 2);
    for cell in cells {
        let offset = u16::try_from(storage.len()).expect("fixture row fits narrow offsets");
        offsets.extend_from_slice(&offset.to_le_bytes());
        storage.extend_from_slice(&cell);
    }
    (storage, offsets)
}

fn tile(mode: FixtureMode) -> TestResult<tst::Tile> {
    let first_column = match mode {
        FixtureMode::Mixed => vec![
            cell_for_control(1, 1, CellDataFormatKind::Checkbox, Some(1.0))?,
            cell_for_control(2, 2, CellDataFormatKind::StarRating, Some(3.0))?,
            cell_for_control(
                3,
                3,
                CellDataFormatKind::NumericControlNumberOrPercentage,
                Some(40.0),
            )?,
            cell_for_control(
                4,
                4,
                CellDataFormatKind::NumericControlNumberOrPercentage,
                Some(4.0),
            )?,
            cell_for_control(5, 5, CellDataFormatKind::PopUpMenu, None)?,
        ],
        FixtureMode::SharedCheckbox => vec![
            cell_for_control(1, 1, CellDataFormatKind::Checkbox, Some(1.0))?,
            cell_for_control(1, 1, CellDataFormatKind::Checkbox, Some(0.0))?,
            empty_cell(),
            empty_cell(),
            empty_cell(),
        ],
        FixtureMode::ScalarSeed => vec![
            number_cell(7.0)?,
            empty_cell(),
            empty_cell(),
            empty_cell(),
            empty_cell(),
        ],
        FixtureMode::Empty => vec![empty_cell(); 5],
    };
    let rows = first_column
        .into_iter()
        .map(|first| pack_row(vec![first, empty_cell(), empty_cell()]))
        .collect::<Vec<_>>();
    Ok(tst::Tile {
        max_column: 2,
        max_row: 4,
        num_cells: 15,
        numrows: 5,
        row_infos: rows
            .into_iter()
            .enumerate()
            .map(|(row, (storage, offsets))| tst::TileRowInfo {
                tile_row_index: row as u32,
                cell_count: 3,
                storage_version: Some(5),
                cell_storage_buffer: Some(storage),
                cell_offsets: Some(offsets),
                ..Default::default()
            })
            .collect(),
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        ..Default::default()
    })
}

fn format_entry(key: u32, refcount: u32, format_type: u32) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount,
        format: Some(tsk::FormatStructArchive {
            format_type: Some(format_type),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn control_entry(
    key: u32,
    refcount: u32,
    interaction_type: u32,
    model_identifier: Option<u64>,
) -> tst::table_data_list::ListEntry {
    let mut spec = tst::CellSpecArchive {
        interaction_type,
        ..Default::default()
    };
    match interaction_type {
        6 => {
            spec.range_control_min = Some(0.0);
            spec.range_control_max = Some(5.0);
            spec.range_control_inc = Some(1.0);
        },
        5 => {
            spec.range_control_min = Some(0.0);
            spec.range_control_max = Some(100.0);
            spec.range_control_inc = Some(10.0);
        },
        4 => {
            spec.range_control_min = Some(1.0);
            spec.range_control_max = Some(10.0);
            spec.range_control_inc = Some(1.0);
        },
        7 => {
            spec.chooser_control_popup_model = model_identifier.map(reference);
            spec.chooser_control_start_w_first = Some(true);
        },
        _ => {},
    }
    tst::table_data_list::ListEntry {
        key,
        refcount,
        cell_spec: Some(spec),
        ..Default::default()
    }
}

fn sidecars(mode: FixtureMode) -> TestResult<ArchiveObject> {
    let (format_entries, control_entries) = match mode {
        FixtureMode::Mixed => (
            vec![
                format_entry(1, 1, 263),
                format_entry(2, 1, 267),
                format_entry(3, 1, 256),
                format_entry(4, 1, 256),
                format_entry(5, 1, 260),
            ],
            vec![
                control_entry(1, 1, 8, None),
                control_entry(2, 1, 6, None),
                control_entry(3, 1, 5, None),
                control_entry(4, 1, 4, None),
                control_entry(5, 1, 7, Some(CONTROL_MODEL_ID)),
            ],
        ),
        FixtureMode::SharedCheckbox => (
            vec![format_entry(1, 2, 263)],
            vec![control_entry(1, 2, 8, None)],
        ),
        FixtureMode::ScalarSeed => (vec![format_entry(1, 1, 256)], Vec::new()),
        FixtureMode::Empty => (Vec::new(), Vec::new()),
    };
    let list_specs = [
        (tst::table_data_list::ListType::String, Vec::new()),
        (tst::table_data_list::ListType::Formula, Vec::new()),
        (tst::table_data_list::ListType::Format, format_entries),
        (
            tst::table_data_list::ListType::ControlCellSpec,
            control_entries,
        ),
    ];
    let mut messages = Vec::new();
    let mut control_index = None;
    for (index, (list_type, entries)) in list_specs.into_iter().enumerate() {
        if list_type == tst::table_data_list::ListType::ControlCellSpec {
            control_index = Some(index);
        }
        messages.push(RawMessage {
            type_: TABLE_DATA_LIST_TYPE,
            data: tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: entries.len() as u32 + 1,
                entries,
                is_new_for_bnc: Some(true),
                ..Default::default()
            }
            .encode_to_vec(),
        });
    }
    let mut sidecars = ArchiveObject::new(SIDECAR_ID, messages)?;
    if let Some(index) = control_index
        && matches!(mode, FixtureMode::Mixed)
    {
        let info = sidecars
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or_else(|| io::Error::other("control message metadata is missing"))?;
        info.object_references = vec![CONTROL_MODEL_ID];
        let mut field = FieldInfo::new(vec![3, 5]);
        field.r#type = Some(FieldType::ObjectReference);
        field.object_references = vec![CONTROL_MODEL_ID];
        info.field_infos.push(field);
    }
    Ok(sidecars)
}

fn table_model() -> tst::TableModelArchive {
    tst::TableModelArchive {
        table_id: "cell-control-table-id".to_owned(),
        table_name: "Cell Controls".to_owned(),
        table_style: reference(SIDECAR_ID),
        body_text_style: reference(SIDECAR_ID),
        header_row_text_style: reference(SIDECAR_ID),
        header_column_text_style: reference(SIDECAR_ID),
        footer_row_text_style: reference(SIDECAR_ID),
        body_cell_style: reference(SIDECAR_ID),
        header_row_style: reference(SIDECAR_ID),
        header_column_style: reference(SIDECAR_ID),
        footer_row_style: reference(SIDECAR_ID),
        number_of_rows: 5,
        number_of_columns: 3,
        base_data_store: tst::DataStore {
            row_headers: tst::HeaderStorage {
                bucket_hash_function: 1,
                ..Default::default()
            },
            column_headers: reference(SIDECAR_ID),
            tiles: tst::TileStorage {
                tiles: vec![tst::tile_storage::Tile {
                    tileid: 0,
                    tile: reference(TILE_ID),
                }],
                tile_size: Some(256),
                ..Default::default()
            },
            string_table: reference(SIDECAR_ID),
            style_table: reference(SIDECAR_ID),
            formula_table: reference(SIDECAR_ID),
            format_table_pre_bnc: reference(SIDECAR_ID),
            format_table: Some(reference(SIDECAR_ID)),
            control_cell_spec_table: Some(reference(SIDECAR_ID)),
            next_row_strip_id: 1,
            next_column_strip_id: 1,
            row_tile_tree: tst::TableRbTree::default(),
            column_tile_tree: tst::TableRbTree::default(),
            ..Default::default()
        },
        ..Default::default()
    }
}

fn metadata_component(
    identifier: u64,
    locator: &str,
    save_token: u64,
    object_ids: &[u64],
) -> TestResult<Vec<u8>> {
    let mut data = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(save_token),
        object_uuid_map_entries: object_ids.iter().copied().map(uuid_entry).collect(),
        ..Default::default()
    }
    .encode_to_vec();
    append_varint_field(&mut data, 90, identifier.saturating_add(10_000))?;
    Ok(data)
}

fn metadata(mode: FixtureMode) -> TestResult<Vec<u8>> {
    let mut document_ids = vec![
        DOCUMENT_ID,
        SHEET_ID,
        TABLE_INFO_ID,
        TABLE_MODEL_ID,
        SIDECAR_ID,
        TILE_ID,
    ];
    if matches!(mode, FixtureMode::Mixed) {
        document_ids.push(CONTROL_MODEL_ID);
    }
    let document = metadata_component(100, "Document", 9, &document_ids)?;
    let view = metadata_component(300, "ViewState", 7, &[VIEW_STATE_ID])?;
    let versioned = metadata_component(100, "Document", 3, &[999])?;
    let mut data = tsp::PackageMetadata {
        last_object_identifier: 1_000,
        save_token: Some(10),
        ..Default::default()
    }
    .encode_to_vec();
    append_length_delimited_field(&mut data, 3, &document)?;
    append_length_delimited_field(&mut data, 3, &view)?;
    append_length_delimited_field(&mut data, 11, &versioned)?;
    append_varint_field(&mut data, 90, 0xfeed_beef)?;
    Ok(data)
}

fn compressed(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    Ok(SnappyStream::compress(&Archive { objects }.to_bytes()?)?)
}

fn fixture(mode: FixtureMode) -> TestResult<Vec<u8>> {
    let mut document = object(
        DOCUMENT_ID,
        1,
        tn::DocumentArchive {
            sheets: vec![reference(SHEET_ID)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    document.archive_info.message_infos[0].object_references = vec![SHEET_ID];
    let mut sheet = object(
        SHEET_ID,
        2,
        tn::SheetArchive {
            name: "Cell Control Sheet".to_owned(),
            drawable_infos: vec![reference(TABLE_INFO_ID)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    sheet.archive_info.message_infos[0].object_references = vec![TABLE_INFO_ID];
    let mut info = object(
        TABLE_INFO_ID,
        6_000,
        tst::TableInfoArchive {
            super_: tsd::DrawableArchive::default(),
            table_model: reference(TABLE_MODEL_ID),
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    info.archive_info.message_infos[0].object_references = vec![TABLE_MODEL_ID];
    let mut model = object(TABLE_MODEL_ID, 6_001, table_model().encode_to_vec())?;
    model.archive_info.message_infos[0].object_references = vec![SIDECAR_ID, TILE_ID];
    let mut unknown = FieldInfo::new(vec![99, 1]);
    unknown.data_references = vec![700];
    model.archive_info.message_infos[0]
        .field_infos
        .push(unknown);
    let mut tile = object(TILE_ID, 6_002, tile(mode)?.encode_to_vec())?;
    tile.archive_info.message_infos[0].data_references = vec![701];
    let sidecars = sidecars(mode)?;
    let mut objects = vec![document, sheet, info, model, sidecars, tile];
    if matches!(mode, FixtureMode::Mixed) {
        let mut model = object(
            CONTROL_MODEL_ID,
            6_206,
            tst::PopUpMenuModel {
                item: Vec::new(),
                tsce_item: {
                    let mut values = vec![tsce::CellValueArchive {
                        cell_value_type: tsce::cell_value_archive::CellValueType::NilType as i32,
                        ..Default::default()
                    }];
                    values.extend(["Low", "Medium", "High"].into_iter().map(popup_item));
                    values
                },
            }
            .encode_to_vec(),
        )?;
        model.archive_info.message_infos[0].data_references = vec![702];
        objects.push(model);
    }
    let document_member = compressed(objects)?;
    let view_member = compressed(vec![object(
        VIEW_STATE_ID,
        210,
        b"unselected view-state bytes".to_vec(),
    )?])?;
    let metadata_member = compressed(vec![object(
        METADATA_OBJECT_ID,
        METADATA_TYPE,
        metadata(mode)?,
    )?])?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            (DOCUMENT_MEMBER, document_member.as_slice()),
            (VIEW_STATE_MEMBER, view_member.as_slice()),
            (METADATA_MEMBER, metadata_member.as_slice()),
            ("preview.jpg", b"control preview".as_slice()),
            ("preview-micro.jpg", b"control micro".as_slice()),
            ("preview-web.jpg", b"control web".as_slice()),
            ("Data/sentinel.bin", b"unrelated control data".as_slice()),
        ],
        Limits::default(),
    )?)
}

/// Rebuild the synthetic graph with the component split used by the native
/// Wave85 source: the rooted model lives in CalculationEngine, its tile is a
/// Tables member, format/control lists are separate list members, and the
/// popup model is a separate current component.  Scalar controls are fully
/// metadata-owned across those members; the popup model remains a deliberately
/// unsupported changed dependency for this bounded owner slice.
fn split_component_fixture() -> TestResult<Vec<u8>> {
    let source = fixture(FixtureMode::Mixed)?;
    let archive = member_archive(&source, DOCUMENT_MEMBER)?;
    let document = archive
        .object(DOCUMENT_ID)
        .cloned()
        .ok_or_else(|| io::Error::other("Document object is missing"))?;
    let sheet = archive
        .object(SHEET_ID)
        .cloned()
        .ok_or_else(|| io::Error::other("Sheet object is missing"))?;
    let info = archive
        .object(TABLE_INFO_ID)
        .cloned()
        .ok_or_else(|| io::Error::other("TableInfo object is missing"))?;
    let mut model = archive
        .object(TABLE_MODEL_ID)
        .cloned()
        .ok_or_else(|| io::Error::other("TableModel object is missing"))?;
    let sidecar = archive
        .object(SIDECAR_ID)
        .ok_or_else(|| io::Error::other("sidecar object is missing"))?;
    let tile = archive
        .object(TILE_ID)
        .cloned()
        .ok_or_else(|| io::Error::other("Tile object is missing"))?;
    let popup = archive
        .object(CONTROL_MODEL_ID)
        .cloned()
        .ok_or_else(|| io::Error::other("Pop-Up model is missing"))?;

    let mut model_payload = tst::TableModelArchive::decode(
        model
            .messages
            .first()
            .ok_or_else(|| io::Error::other("TableModel payload is missing"))?
            .data
            .as_slice(),
    )?;
    model_payload.base_data_store.format_table_pre_bnc = reference(FORMAT_LIST_ID);
    model_payload.base_data_store.format_table = Some(reference(FORMAT_LIST_ID));
    model_payload.base_data_store.control_cell_spec_table = Some(reference(CONTROL_LIST_ID));
    model
        .messages
        .first_mut()
        .ok_or_else(|| io::Error::other("TableModel payload is missing"))?
        .data = model_payload.encode_to_vec();
    if let Some(info) = model.archive_info.message_infos.first_mut() {
        info.object_references = vec![SIDECAR_ID, TILE_ID, FORMAT_LIST_ID, CONTROL_LIST_ID];
    }

    let scalar_sidecar = object_with_messages(sidecar, &[0, 1])?;
    let format_sidecar = object_with_messages_with_identifier(sidecar, FORMAT_LIST_ID, &[2])?;
    let control_sidecar = object_with_messages_with_identifier(sidecar, CONTROL_LIST_ID, &[3])?;

    let document_member = compressed(vec![document, sheet, info])?;
    let calculation_member = compressed(vec![model, scalar_sidecar])?;
    let tile_member = compressed(vec![tile])?;
    let format_member = compressed(vec![format_sidecar])?;
    let control_member = compressed(vec![control_sidecar])?;
    let popup_member = compressed(vec![popup])?;
    let view_member = Catalog::from_bytes(&source)?
        .iter()
        .find(|entry| entry.name() == VIEW_STATE_MEMBER)
        .ok_or_else(|| io::Error::other("ViewState member is missing"))?
        .data()
        .to_vec();
    let metadata_member = compressed(vec![object(
        METADATA_OBJECT_ID,
        METADATA_TYPE,
        split_metadata_payload()?,
    )?])?;

    Ok(litchi_iwa_archive::package::to_bytes(
        [
            (DOCUMENT_MEMBER, document_member.as_slice()),
            (CALCULATION_MEMBER, calculation_member.as_slice()),
            (TILE_MEMBER, tile_member.as_slice()),
            (FORMAT_MEMBER, format_member.as_slice()),
            (CONTROL_MEMBER, control_member.as_slice()),
            (POPUP_MEMBER, popup_member.as_slice()),
            (VIEW_STATE_MEMBER, view_member.as_slice()),
            (METADATA_MEMBER, metadata_member.as_slice()),
            ("preview.jpg", b"split control preview".as_slice()),
            ("preview-micro.jpg", b"split control micro".as_slice()),
            ("preview-web.jpg", b"split control web".as_slice()),
            (
                "Data/sentinel.bin",
                b"split unrelated control data".as_slice(),
            ),
        ],
        Limits::default(),
    )?)
}

/// Keep the rooted model, tile, and both lists co-located while moving only
/// the shared Pop-Up Menu model to its own metadata-owned member. This guards
/// the early changed-operation refusal against the smallest split graph.
fn popup_only_split_fixture() -> TestResult<Vec<u8>> {
    let source = fixture(FixtureMode::Mixed)?;
    let popup = member_archive(&source, DOCUMENT_MEMBER)?
        .object(CONTROL_MODEL_ID)
        .cloned()
        .ok_or_else(|| io::Error::other("Pop-Up model is missing"))?;
    let source = rewrite_member(&source, DOCUMENT_MEMBER, |archive| {
        archive
            .remove_object(CONTROL_MODEL_ID)
            .ok_or_else(|| io::Error::other("Pop-Up model is missing"))?;
        Ok(())
    })?;
    let source = rewrite_metadata_root(&source, |metadata| {
        let document = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .ok_or_else(|| io::Error::other("Document metadata is missing"))?;
        document
            .object_uuid_map_entries
            .retain(|entry| entry.identifier != CONTROL_MODEL_ID);
        document
            .external_references
            .push(tsp::ComponentExternalReference {
                component_identifier: POPUP_COMPONENT_ID,
                object_identifier: Some(CONTROL_MODEL_ID),
                is_weak: Some(false),
            });
        metadata.components.push(tsp::ComponentInfo {
            identifier: POPUP_COMPONENT_ID,
            preferred_locator: "Tables/Popup-905753".to_owned(),
            locator: Some("Tables/Popup-905753".to_owned()),
            save_token: Some(10),
            object_uuid_map_entries: vec![uuid_entry(CONTROL_MODEL_ID)],
            ..Default::default()
        });
        Ok(())
    })?;
    let popup_member = compressed(vec![popup])?;
    let mut entries = Catalog::from_bytes(&source)?
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<Vec<_>>();
    entries.push((POPUP_MEMBER.to_owned(), popup_member));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        Limits::default(),
    )?)
}

fn object_with_messages(source: &ArchiveObject, indices: &[usize]) -> TestResult<ArchiveObject> {
    object_with_messages_with_identifier(
        source,
        source
            .archive_info
            .identifier
            .ok_or_else(|| io::Error::other("source object identifier is missing"))?,
        indices,
    )
}

fn object_with_messages_with_identifier(
    source: &ArchiveObject,
    identifier: u64,
    indices: &[usize],
) -> TestResult<ArchiveObject> {
    let messages = indices
        .iter()
        .map(|index| {
            source
                .messages
                .get(*index)
                .cloned()
                .ok_or_else(|| io::Error::other("sidecar message is missing"))
        })
        .collect::<Result<Vec<_>, _>>()?;
    let mut object = ArchiveObject::new(identifier, messages)?;
    object.archive_info = source.archive_info.clone();
    object.archive_info.identifier = Some(identifier);
    object.archive_info.message_infos = indices
        .iter()
        .map(|index| {
            source
                .archive_info
                .message_infos
                .get(*index)
                .cloned()
                .ok_or_else(|| io::Error::other("sidecar message metadata is missing"))
        })
        .collect::<Result<Vec<_>, _>>()?;
    Ok(object)
}

fn split_metadata_component(
    identifier: u64,
    locator: &str,
    token: u64,
    object_ids: &[u64],
    external_references: &[tsp::ComponentExternalReference],
) -> TestResult<Vec<u8>> {
    let mut data = tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(token),
        object_uuid_map_entries: object_ids.iter().copied().map(uuid_entry).collect(),
        external_references: external_references.to_vec(),
        ..Default::default()
    }
    .encode_to_vec();
    append_varint_field(&mut data, 90, identifier.saturating_add(20_000))?;
    Ok(data)
}

fn split_metadata_payload() -> TestResult<Vec<u8>> {
    let external = |component_identifier, object_identifier| tsp::ComponentExternalReference {
        component_identifier,
        object_identifier: Some(object_identifier),
        is_weak: Some(false),
    };
    let external_component = |component_identifier| tsp::ComponentExternalReference {
        component_identifier,
        object_identifier: None,
        is_weak: Some(false),
    };
    let document = split_metadata_component(
        100,
        "Document",
        9,
        &[DOCUMENT_ID, SHEET_ID, TABLE_INFO_ID],
        &[external(CALCULATION_COMPONENT_ID, TABLE_MODEL_ID)],
    )?;
    let calculation = split_metadata_component(
        CALCULATION_COMPONENT_ID,
        "CalculationEngine",
        10,
        &[TABLE_MODEL_ID],
        &[
            external(100, TABLE_INFO_ID),
            external_component(TILE_COMPONENT_ID),
            external_component(FORMAT_COMPONENT_ID),
            external_component(CONTROL_COMPONENT_ID),
        ],
    )?;
    let tile = split_metadata_component(TILE_COMPONENT_ID, "Tables/Tile", 10, &[], &[])?;
    let format = split_metadata_component(
        FORMAT_COMPONENT_ID,
        "Tables/DataList-904498-2",
        10,
        &[],
        &[],
    )?;
    let control = split_metadata_component(
        CONTROL_COMPONENT_ID,
        "Tables/DataList-904499-2",
        10,
        &[],
        &[external(POPUP_COMPONENT_ID, CONTROL_MODEL_ID)],
    )?;
    let popup = split_metadata_component(
        POPUP_COMPONENT_ID,
        "Tables/Popup-905753",
        10,
        &[CONTROL_MODEL_ID],
        &[],
    )?;
    let view = split_metadata_component(300, "ViewState", 7, &[VIEW_STATE_ID], &[])?;
    let versioned = split_metadata_component(100, "Document", 3, &[999], &[])?;
    let mut data = tsp::PackageMetadata {
        last_object_identifier: 1_000,
        save_token: Some(10),
        ..Default::default()
    }
    .encode_to_vec();
    for component in [document, calculation, tile, format, control, popup, view] {
        append_length_delimited_field(&mut data, 3, &component)?;
    }
    append_length_delimited_field(&mut data, 11, &versioned)?;
    append_varint_field(&mut data, 90, 0xfeed_beef)?;
    Ok(data)
}

fn member_archive(source: &[u8], member: &str) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other(format!("member {member} is missing")))?;
    Ok(Archive::parse(
        SnappyStream::decompress(entry.data())?.as_bytes(),
    )?)
}

fn rewrite_member(
    source: &[u8],
    member: &str,
    mut rewrite: impl FnMut(&mut Archive) -> TestResult,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other(format!("member {member} is missing")))?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    rewrite(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            member,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn rewrite_sidecar(
    source: &[u8],
    list_type: tst::table_data_list::ListType,
    mut rewrite: impl FnMut(&mut tst::TableDataList) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(SIDECAR_ID)
            .ok_or_else(|| io::Error::other("sidecar object is missing"))?;
        let index = object
            .messages
            .iter()
            .position(|message| {
                message.type_ == TABLE_DATA_LIST_TYPE
                    && tst::TableDataList::decode(message.data.as_slice())
                        .map(|list| list.list_type == list_type as i32)
                        .unwrap_or(false)
            })
            .ok_or_else(|| io::Error::other("requested table-data list is missing"))?;
        let previous = tst::TableDataList::decode(object.messages[index].data.as_slice())?;
        let mut current = previous.clone();
        rewrite(&mut current)?;
        object.messages[index].data = current.encode_to_vec();
        Ok(())
    })
}

fn rewrite_control_list(
    source: &[u8],
    rewrite: impl FnMut(&mut tst::TableDataList) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_sidecar(
        source,
        tst::table_data_list::ListType::ControlCellSpec,
        rewrite,
    )
}

fn rewrite_format_list(
    source: &[u8],
    rewrite: impl FnMut(&mut tst::TableDataList) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_sidecar(source, tst::table_data_list::ListType::Format, rewrite)
}

fn rewrite_tile(
    source: &[u8],
    mut rewrite: impl FnMut(&mut tst::Tile) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(TILE_ID)
            .ok_or_else(|| io::Error::other("tile object is missing"))?;
        let message = object
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("tile payload is missing"))?;
        let mut tile = tst::Tile::decode(message.data.as_slice())?;
        rewrite(&mut tile)?;
        message.data = tile.encode_to_vec();
        Ok(())
    })
}

fn rewrite_metadata_root(
    source: &[u8],
    mut rewrite: impl FnMut(&mut tsp::PackageMetadata) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_member(source, METADATA_MEMBER, |archive| {
        let object = archive
            .object_mut(METADATA_OBJECT_ID)
            .ok_or_else(|| io::Error::other("metadata object is missing"))?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| message.type_ == METADATA_TYPE)
            .ok_or_else(|| io::Error::other("metadata payload is missing"))?;
        let mut metadata = tsp::PackageMetadata::decode(message.data.as_slice())?;
        rewrite(&mut metadata)?;
        message.data = metadata.encode_to_vec();
        Ok(())
    })
}

fn control_message_index(archive: &Archive) -> TestResult<usize> {
    let object = archive
        .object(SIDECAR_ID)
        .ok_or_else(|| io::Error::other("sidecar object is missing"))?;
    object
        .messages
        .iter()
        .position(|message| {
            message.type_ == TABLE_DATA_LIST_TYPE
                && tst::TableDataList::decode(message.data.as_slice())
                    .map(|list| {
                        list.list_type == tst::table_data_list::ListType::ControlCellSpec as i32
                    })
                    .unwrap_or(false)
        })
        .ok_or_else(|| io::Error::other("control list message is missing").into())
}

fn rewrite_control_info(
    source: &[u8],
    mut rewrite: impl FnMut(&mut litchi_iwa_core::MessageInfo) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let index = control_message_index(archive)?;
        let object = archive
            .object_mut(SIDECAR_ID)
            .ok_or_else(|| io::Error::other("sidecar object is missing"))?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(index)
            .ok_or_else(|| io::Error::other("control message metadata is missing"))?;
        rewrite(info)
    })
}

fn object_messages(
    source: &[u8],
    member: &str,
    identifier: u64,
) -> TestResult<Vec<(u32, Vec<u8>)>> {
    Ok(member_archive(source, member)?
        .object(identifier)
        .ok_or_else(|| io::Error::other("object is missing"))?
        .messages
        .iter()
        .map(|message| (message.type_, message.data.clone()))
        .collect())
}

fn control_entries(source: &[u8]) -> TestResult<Vec<tst::table_data_list::ListEntry>> {
    let archive = member_archive(source, DOCUMENT_MEMBER)?;
    let object = archive
        .object(SIDECAR_ID)
        .ok_or_else(|| io::Error::other("sidecar object is missing"))?;
    for message in &object.messages {
        if message.type_ != TABLE_DATA_LIST_TYPE {
            continue;
        }
        let list = tst::TableDataList::decode(message.data.as_slice())?;
        if list.list_type == tst::table_data_list::ListType::ControlCellSpec as i32 {
            return Ok(list.entries);
        }
    }
    Err(io::Error::other("control list is missing").into())
}

fn format_entries(source: &[u8]) -> TestResult<Vec<tst::table_data_list::ListEntry>> {
    let archive = member_archive(source, DOCUMENT_MEMBER)?;
    let object = archive
        .object(SIDECAR_ID)
        .ok_or_else(|| io::Error::other("sidecar object is missing"))?;
    for message in &object.messages {
        if message.type_ != TABLE_DATA_LIST_TYPE {
            continue;
        }
        let list = tst::TableDataList::decode(message.data.as_slice())?;
        if list.list_type == tst::table_data_list::ListType::Format as i32 {
            return Ok(list.entries);
        }
    }
    Err(io::Error::other("format list is missing").into())
}

fn control_message_info(source: &[u8]) -> TestResult<litchi_iwa_core::MessageInfo> {
    let archive = member_archive(source, DOCUMENT_MEMBER)?;
    let index = control_message_index(&archive)?;
    archive
        .object(SIDECAR_ID)
        .and_then(|object| object.archive_info.message_infos.get(index))
        .cloned()
        .ok_or_else(|| io::Error::other("control message metadata is missing").into())
}

fn assert_control_metadata_exact(source: &[u8]) -> TestResult {
    let entries = control_entries(source)?;
    let info = control_message_info(source)?;
    let mut expected = Vec::new();
    for entry in &entries {
        if let Some(spec) = &entry.cell_spec
            && let Some(reference) = &spec.chooser_control_popup_model
        {
            expected.push((entry.key, reference.identifier));
        }
    }
    let expected_ids = expected
        .iter()
        .map(|(_, identifier)| *identifier)
        .collect::<Vec<_>>();
    assert_eq!(info.object_references, expected_ids);
    for (key, identifier) in expected {
        let matches = info
            .field_infos
            .iter()
            .filter(|field| field.path.as_slice() == [3, key])
            .collect::<Vec<_>>();
        assert_eq!(matches.len(), 1, "control FieldInfo path [{key}]");
        assert_eq!(matches[0].object_references.as_slice(), [identifier]);
    }
    assert!(info.field_infos.iter().all(|field| {
        field.path.as_slice().first() != Some(&3)
            || entries
                .iter()
                .any(|entry| field.path.as_slice() == [3, entry.key])
    }));
    Ok(())
}

fn changed_members(source: &[u8], target: &[u8]) -> TestResult<Vec<String>> {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for entry in before.iter() {
        let other = after
            .iter()
            .find(|candidate| candidate.name() == entry.name());
        let Some(other) = other else {
            if entry.name().starts_with("preview") {
                continue;
            }
            return Err(io::Error::other("candidate member disappeared").into());
        };
        if entry.data() != other.data() {
            changed.push(entry.name().to_owned());
        }
    }
    Ok(changed)
}

fn metadata_payload(source: &[u8]) -> TestResult<tsp::PackageMetadata> {
    let archive = member_archive(source, METADATA_MEMBER)?;
    let object = archive
        .object(METADATA_OBJECT_ID)
        .ok_or_else(|| io::Error::other("metadata object is missing"))?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == METADATA_TYPE)
        .ok_or_else(|| io::Error::other("metadata payload is missing"))?;
    Ok(tsp::PackageMetadata::decode(message.data.as_slice())?)
}

fn metadata_external_signature(
    source: &[u8],
) -> TestResult<Vec<(u64, u64, Option<u64>, Option<bool>)>> {
    let mut edges = Vec::new();
    for component in metadata_payload(source)?.components {
        for edge in component.external_references {
            edges.push((
                component.identifier,
                edge.component_identifier,
                edge.object_identifier,
                edge.is_weak,
            ));
        }
    }
    edges.sort();
    Ok(edges)
}

fn assert_split_native_metadata_shape(source: &[u8]) -> TestResult {
    let metadata = metadata_payload(source)?;
    for identifier in [TILE_COMPONENT_ID, FORMAT_COMPONENT_ID, CONTROL_COMPONENT_ID] {
        let component = metadata
            .components
            .iter()
            .find(|component| component.identifier == identifier)
            .ok_or_else(|| io::Error::other("required sidecar metadata is missing"))?;
        assert!(
            component.object_uuid_map_entries.is_empty(),
            "native sidecar object IDs must not be treated as UUID-owned: {identifier}"
        );
    }
    let calculation = metadata
        .components
        .iter()
        .find(|component| component.identifier == CALCULATION_COMPONENT_ID)
        .ok_or_else(|| io::Error::other("CalculationEngine metadata is missing"))?;
    for identifier in [TILE_COMPONENT_ID, FORMAT_COMPONENT_ID, CONTROL_COMPONENT_ID] {
        let edges = calculation
            .external_references
            .iter()
            .filter(|edge| edge.component_identifier == identifier)
            .collect::<Vec<_>>();
        assert_eq!(edges.len(), 1, "sidecar component edge {identifier}");
        assert_eq!(edges[0].object_identifier, None);
        assert_eq!(edges[0].is_weak, Some(false));
    }
    Ok(())
}

fn member_for_metadata_component(identifier: u64) -> Option<&'static str> {
    match identifier {
        CALCULATION_COMPONENT_ID => Some(CALCULATION_MEMBER),
        TILE_COMPONENT_ID => Some(TILE_MEMBER),
        FORMAT_COMPONENT_ID => Some(FORMAT_MEMBER),
        CONTROL_COMPONENT_ID => Some(CONTROL_MEMBER),
        POPUP_COMPONENT_ID => Some(POPUP_MEMBER),
        300 => Some(VIEW_STATE_MEMBER),
        _ => None,
    }
}

fn assert_split_metadata_transition(
    source: &[u8],
    target: &[u8],
    changed: &[String],
) -> TestResult {
    let before = metadata_payload(source)?;
    let after = metadata_payload(target)?;
    assert_eq!(
        after.save_token,
        before.save_token.and_then(|token| token.checked_add(1)),
        "the root metadata save token must advance exactly once",
    );
    assert_eq!(
        metadata_external_signature(source)?,
        metadata_external_signature(target)?,
        "scalar COW must not fabricate or drop external ownership edges",
    );
    for component in &before.components {
        let candidate = after
            .components
            .iter()
            .find(|other| other.identifier == component.identifier)
            .ok_or_else(|| io::Error::other("metadata component disappeared"))?;
        let member = member_for_metadata_component(component.identifier);
        let changed_member = member.is_some_and(|member| changed.iter().any(|name| name == member));
        assert_eq!(
            candidate.save_token,
            if changed_member {
                component.save_token.and_then(|token| token.checked_add(1))
            } else {
                component.save_token
            },
            "unexpected save-token transition for metadata component {}",
            component.identifier,
        );
    }
    for identifier in [TILE_COMPONENT_ID, FORMAT_COMPONENT_ID, CONTROL_COMPONENT_ID] {
        let source_component = before
            .components
            .iter()
            .find(|component| component.identifier == identifier)
            .ok_or_else(|| io::Error::other("required sidecar metadata is missing"))?;
        let target_component = after
            .components
            .iter()
            .find(|component| component.identifier == identifier)
            .ok_or_else(|| io::Error::other("required sidecar metadata disappeared"))?;
        assert_eq!(
            target_component.save_token,
            source_component
                .save_token
                .and_then(|token| token.checked_add(1)),
            "sidecar component {identifier} must advance exactly once",
        );
    }
    let source_model = before
        .components
        .iter()
        .find(|component| component.identifier == CALCULATION_COMPONENT_ID)
        .ok_or_else(|| io::Error::other("CalculationEngine metadata is missing"))?;
    let target_model = after
        .components
        .iter()
        .find(|component| component.identifier == CALCULATION_COMPONENT_ID)
        .ok_or_else(|| io::Error::other("CalculationEngine metadata disappeared"))?;
    assert_eq!(
        target_model.save_token, source_model.save_token,
        "the model component token must remain unchanged"
    );
    Ok(())
}

fn assert_split_scalar_locality(source: &[u8], target: &[u8], changed: &[String]) -> TestResult {
    let allowed = [
        TILE_MEMBER,
        FORMAT_MEMBER,
        CONTROL_MEMBER,
        METADATA_MEMBER,
        "preview.jpg",
        "preview-micro.jpg",
        "preview-web.jpg",
    ];
    assert!(
        changed
            .iter()
            .all(|member| allowed.contains(&member.as_str())),
        "unexpected split scalar member mutation: {changed:?}",
    );
    for required in [TILE_MEMBER, FORMAT_MEMBER, CONTROL_MEMBER, METADATA_MEMBER] {
        assert!(
            changed.iter().any(|member| member == required),
            "split scalar edit did not touch required member {required}"
        );
    }
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    for member in [
        DOCUMENT_MEMBER,
        CALCULATION_MEMBER,
        POPUP_MEMBER,
        VIEW_STATE_MEMBER,
        "Data/sentinel.bin",
    ] {
        let source_entry = before
            .iter()
            .find(|entry| entry.name() == member)
            .ok_or_else(|| io::Error::other(format!("source member {member} is missing")))?;
        let target_entry = after
            .iter()
            .find(|entry| entry.name() == member)
            .ok_or_else(|| io::Error::other(format!("target member {member} is missing")))?;
        assert_eq!(
            source_entry.data(),
            target_entry.data(),
            "unselected split member {member} changed"
        );
    }
    Ok(())
}

fn split_list_entries(
    source: &[u8],
    member: &str,
    object_identifier: u64,
    list_type: tst::table_data_list::ListType,
) -> TestResult<Vec<tst::table_data_list::ListEntry>> {
    let archive = member_archive(source, member)?;
    let object = archive
        .object(object_identifier)
        .ok_or_else(|| io::Error::other("split list object is missing"))?;
    for message in &object.messages {
        if message.type_ != TABLE_DATA_LIST_TYPE {
            continue;
        }
        let list = tst::TableDataList::decode(message.data.as_slice())?;
        if list.list_type == list_type as i32 {
            return Ok(list.entries);
        }
    }
    Err(io::Error::other("split table-data list is missing").into())
}

fn rewrite_split_list(
    source: &[u8],
    member: &str,
    object_identifier: u64,
    list_type: tst::table_data_list::ListType,
    mut rewrite: impl FnMut(&mut tst::TableDataList) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_member(source, member, |archive| {
        let object = archive
            .object_mut(object_identifier)
            .ok_or_else(|| io::Error::other("split list object is missing"))?;
        let index = object
            .messages
            .iter()
            .position(|message| {
                message.type_ == TABLE_DATA_LIST_TYPE
                    && tst::TableDataList::decode(message.data.as_slice())
                        .map(|list| list.list_type == list_type as i32)
                        .unwrap_or(false)
            })
            .ok_or_else(|| io::Error::other("split table-data list is missing"))?;
        let mut list = tst::TableDataList::decode(object.messages[index].data.as_slice())?;
        rewrite(&mut list)?;
        object.messages[index].data = list.encode_to_vec();
        Ok(())
    })
}

fn assert_previews_invalidated(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    for name in ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"] {
        assert!(before.iter().any(|entry| entry.name() == name));
        assert!(after.iter().all(|entry| entry.name() != name));
    }
    Ok(())
}

fn assert_changed_edit_rejects(source: &[u8], label: &str) -> TestResult {
    let package = match Package::from_bytes(source) {
        Err(_) => return Ok(()),
        Ok(package) => package,
    };
    let before = package.exact_bytes();
    let edit = match package.edit_table_cell_control_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
    ) {
        Err(_) => return Ok(()),
        Ok(edit) => edit,
    };
    assert!(
        edit.set(CellControl::StarRating(StarRating))
            .commit()
            .is_err(),
        "hostile control graph published: {label}"
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn assert_read_rejects(source: &[u8], label: &str) -> TestResult {
    let package = match Package::from_bytes(source) {
        Err(_) => return Ok(()),
        Ok(package) => package,
    };
    match package.table_cell_control_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
    ) {
        Err(_) => Ok(()),
        Ok(value) => Err(io::Error::other(format!(
            "hostile control graph accepted: {label}: {value:?}"
        ))
        .into()),
    }
}

fn assert_split_owner_rejects(source: &[u8], label: &str) -> TestResult {
    let package = match Package::from_bytes(source) {
        Err(_) => return Ok(()),
        Ok(package) => package,
    };
    let before = package.exact_bytes();
    assert!(
        package
            .table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(0, 0),
            )
            .is_err(),
        "hostile split graph read was accepted: {label}"
    );
    let result = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )
        .and_then(|edit| edit.set(CellControl::StarRating(StarRating)).commit());
    assert!(result.is_err(), "hostile split graph published: {label}");
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn with_corruption(source: &[u8], corruption: Corruption) -> TestResult<Vec<u8>> {
    match corruption {
        Corruption::WrongInteraction => rewrite_control_list(source, |list| {
            list.entries[0]
                .cell_spec
                .as_mut()
                .ok_or_else(|| io::Error::other("control spec missing"))?
                .interaction_type = 999;
            Ok(())
        }),
        Corruption::MissingCellSpec => rewrite_control_list(source, |list| {
            list.entries[0].cell_spec = None;
            Ok(())
        }),
        Corruption::ZeroRefcount => rewrite_control_list(source, |list| {
            list.entries[0].refcount = 0;
            Ok(())
        }),
        Corruption::DuplicateControlKey => rewrite_control_list(source, |list| {
            let duplicate = list
                .entries
                .first()
                .cloned()
                .ok_or_else(|| io::Error::other("control entry missing"))?;
            list.entries.push(duplicate);
            Ok(())
        }),
        Corruption::SegmentedControlList => rewrite_control_list(source, |list| {
            list.segments.push(reference(999));
            Ok(())
        }),
        Corruption::DuplicateControlList => rewrite_member(source, DOCUMENT_MEMBER, |archive| {
            let index = control_message_index(archive)?;
            let object = archive
                .object_mut(SIDECAR_ID)
                .ok_or_else(|| io::Error::other("sidecar object missing"))?;
            let message = object.messages[index].clone();
            let info = object
                .archive_info
                .message_infos
                .get(index)
                .cloned()
                .ok_or_else(|| io::Error::other("control metadata missing"))?;
            object.messages.push(message);
            object.archive_info.message_infos.push(info);
            Ok(())
        }),
        Corruption::ControlAggregateWrong => rewrite_control_info(source, |info| {
            info.object_references = vec![999];
            Ok(())
        }),
        Corruption::ControlAggregateMissing => rewrite_control_info(source, |info| {
            info.object_references.clear();
            Ok(())
        }),
        Corruption::ControlFieldInfoMissing => rewrite_control_info(source, |info| {
            info.field_infos.clear();
            Ok(())
        }),
        Corruption::ControlFieldInfoWrong => rewrite_control_info(source, |info| {
            let field = info
                .field_infos
                .first_mut()
                .ok_or_else(|| io::Error::other("control field metadata missing"))?;
            field.object_references = vec![999];
            Ok(())
        }),
        Corruption::ControlFieldInfoExtra => rewrite_control_info(source, |info| {
            let mut field = FieldInfo::new(vec![3, 99]);
            field.r#type = Some(FieldType::ObjectReference);
            field.object_references = vec![999];
            info.field_infos.push(field);
            Ok(())
        }),
        Corruption::ControlKeyOverflow => rewrite_control_list(source, |list| {
            list.next_list_id = u32::MAX;
            if let Some(entry) = list.entries.first_mut() {
                entry.key = u32::MAX;
            }
            Ok(())
        }),
        Corruption::MissingPopupModel => rewrite_member(source, DOCUMENT_MEMBER, |archive| {
            archive
                .objects
                .retain(|object| object.archive_info.identifier != Some(CONTROL_MODEL_ID));
            Ok(())
        }),
        Corruption::WrongPopupModelType => rewrite_member(source, DOCUMENT_MEMBER, |archive| {
            let object = archive
                .object_mut(CONTROL_MODEL_ID)
                .ok_or_else(|| io::Error::other("control model missing"))?;
            object.messages[0].type_ += 1;
            Ok(())
        }),
        Corruption::DeprecatedReferenceType => rewrite_control_list(source, |list| {
            let reference = list
                .entries
                .iter_mut()
                .find_map(|entry| {
                    entry
                        .cell_spec
                        .as_mut()
                        .and_then(|spec| spec.chooser_control_popup_model.as_mut())
                })
                .ok_or_else(|| io::Error::other("popup reference missing"))?;
            reference.deprecated_type = Some(1);
            Ok(())
        }),
        Corruption::DeprecatedExternalReference => rewrite_control_list(source, |list| {
            let reference = list
                .entries
                .iter_mut()
                .find_map(|entry| {
                    entry
                        .cell_spec
                        .as_mut()
                        .and_then(|spec| spec.chooser_control_popup_model.as_mut())
                })
                .ok_or_else(|| io::Error::other("popup reference missing"))?;
            reference.deprecated_is_external = Some(true);
            Ok(())
        }),
        Corruption::WrongBncKind => rewrite_tile(source, |tile| {
            let row = tile
                .row_infos
                .first_mut()
                .ok_or_else(|| io::Error::other("tile row missing"))?;
            let (storage, offsets) = pack_row(vec![number_cell(1.0)?, empty_cell(), empty_cell()]);
            row.cell_storage_buffer = Some(storage);
            row.cell_offsets = Some(offsets);
            row.cell_count = 3;
            Ok(())
        }),
        Corruption::MissingFormatEntry => rewrite_format_list(source, |list| {
            list.entries.clear();
            Ok(())
        }),
        Corruption::FormatRefcountUndercount => rewrite_format_list(source, |list| {
            list.entries
                .first_mut()
                .ok_or_else(|| io::Error::other("format entry missing"))?
                .refcount = 0;
            Ok(())
        }),
        Corruption::ControlRefcountUndercount => rewrite_control_list(source, |list| {
            list.entries
                .first_mut()
                .ok_or_else(|| io::Error::other("control entry missing"))?
                .refcount = 0;
            Ok(())
        }),
        Corruption::FormatRefcountOvercount => rewrite_format_list(source, |list| {
            list.entries
                .first_mut()
                .ok_or_else(|| io::Error::other("format entry missing"))?
                .refcount = 99;
            Ok(())
        }),
        Corruption::ControlRefcountOvercount => rewrite_control_list(source, |list| {
            list.entries
                .first_mut()
                .ok_or_else(|| io::Error::other("control entry missing"))?
                .refcount = 99;
            Ok(())
        }),
    }
}

fn with_metadata_corruption(source: &[u8], corruption: MetadataCorruption) -> TestResult<Vec<u8>> {
    rewrite_metadata_root(source, |metadata| {
        let document = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .ok_or_else(|| io::Error::other("Document metadata is missing"))?;
        match corruption {
            MetadataCorruption::MissingControlUuid => {
                document
                    .object_uuid_map_entries
                    .retain(|entry| entry.identifier != CONTROL_MODEL_ID);
            },
            MetadataCorruption::VersionedOnlyControlUuid => {
                document
                    .object_uuid_map_entries
                    .retain(|entry| entry.identifier != CONTROL_MODEL_ID);
                let versioned = metadata
                    .versioned_components
                    .iter_mut()
                    .find(|component| component.identifier == 100)
                    .ok_or_else(|| io::Error::other("versioned Document is missing"))?;
                versioned
                    .object_uuid_map_entries
                    .push(uuid_entry(CONTROL_MODEL_ID));
            },
            MetadataCorruption::DuplicateControlUuid => {
                document
                    .object_uuid_map_entries
                    .push(uuid_entry(CONTROL_MODEL_ID));
            },
            MetadataCorruption::AmbiguousControlIdentifier => {
                document.ambiguous_object_identifiers.push(CONTROL_MODEL_ID);
            },
            MetadataCorruption::DataOwnerControlIdentifier => {
                document.data_references.push(tsp::ComponentDataReference {
                    data_identifier: 1_001,
                    object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                        object_identifier: CONTROL_MODEL_ID,
                        count: 1,
                    }],
                });
            },
            MetadataCorruption::RootDataMapControlIdentifier => {
                metadata.data_metadata_map = Some(reference(CONTROL_MODEL_ID));
            },
        }
        Ok(())
    })
}

fn with_physical_alias(source: &[u8]) -> TestResult<Vec<u8>> {
    with_physical_alias_for_member(source, DOCUMENT_MEMBER, "Index/ControlAlias.iwa")
}

fn with_physical_alias_for_member(source: &[u8], member: &str, alias: &str) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let aliased = catalog
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other(format!("{member} member missing")))?
        .data()
        .to_vec();
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<Vec<_>>();
    entries.push((alias.to_owned(), aliased));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        Limits::default(),
    )?)
}

fn with_explicit_locator(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_metadata_root(source, |metadata| {
        let document = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .ok_or_else(|| io::Error::other("Document metadata is missing"))?;
        document.preferred_locator = "Document-preferred-alias".to_owned();
        document.locator = Some("Document".to_owned());
        Ok(())
    })
}

fn with_reserved_identifier(source: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_metadata_root(source, |metadata| {
        let document = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .ok_or_else(|| io::Error::other("Document metadata is missing"))?;
        document.ambiguous_object_identifiers.push(identifier);
        Ok(())
    })
}

fn with_split_metadata_corruption(
    source: &[u8],
    corruption: SplitMetadataCorruption,
) -> TestResult<Vec<u8>> {
    if matches!(corruption, SplitMetadataCorruption::OpaquePopupInbound) {
        return rewrite_member(source, VIEW_STATE_MEMBER, |archive| {
            let object = archive
                .object_mut(VIEW_STATE_ID)
                .ok_or_else(|| io::Error::other("ViewState object is missing"))?;
            let info = object
                .archive_info
                .message_infos
                .first_mut()
                .ok_or_else(|| io::Error::other("ViewState message metadata is missing"))?;
            info.object_references.push(CONTROL_MODEL_ID);
            Ok(())
        });
    }
    rewrite_metadata_root(source, |metadata| {
        let calculation_index = metadata
            .components
            .iter()
            .position(|component| component.identifier == CALCULATION_COMPONENT_ID)
            .ok_or_else(|| io::Error::other("CalculationEngine metadata is missing"))?;
        match corruption {
            SplitMetadataCorruption::MissingCalculationExternal => {
                metadata.components[calculation_index]
                    .external_references
                    .retain(|reference| reference.component_identifier != FORMAT_COMPONENT_ID);
            },
            SplitMetadataCorruption::MissingTileExternal => {
                metadata.components[calculation_index]
                    .external_references
                    .retain(|reference| reference.component_identifier != TILE_COMPONENT_ID);
            },
            SplitMetadataCorruption::MissingControlExternal => {
                metadata.components[calculation_index]
                    .external_references
                    .retain(|reference| reference.component_identifier != CONTROL_COMPONENT_ID);
            },
            SplitMetadataCorruption::MissingPopupExternal => {
                let control = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == CONTROL_COMPONENT_ID)
                    .ok_or_else(|| io::Error::other("control metadata is missing"))?;
                control
                    .external_references
                    .retain(|reference| reference.component_identifier != POPUP_COMPONENT_ID);
            },
            SplitMetadataCorruption::DuplicateCalculationExternal => {
                let reference = metadata.components[calculation_index]
                    .external_references
                    .iter()
                    .find(|reference| reference.component_identifier == FORMAT_COMPONENT_ID)
                    .cloned()
                    .ok_or_else(|| io::Error::other("format external edge is missing"))?;
                metadata.components[calculation_index]
                    .external_references
                    .push(reference);
            },
            SplitMetadataCorruption::DuplicatePopupExternal => {
                let control = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == CONTROL_COMPONENT_ID)
                    .ok_or_else(|| io::Error::other("control metadata is missing"))?;
                let reference = control
                    .external_references
                    .iter()
                    .find(|reference| reference.component_identifier == POPUP_COMPONENT_ID)
                    .cloned()
                    .ok_or_else(|| io::Error::other("popup external edge is missing"))?;
                control.external_references.push(reference);
            },
            SplitMetadataCorruption::ComponentAndObjectFormatExternal => {
                metadata.components[calculation_index]
                    .external_references
                    .push(tsp::ComponentExternalReference {
                        component_identifier: FORMAT_COMPONENT_ID,
                        object_identifier: Some(FORMAT_LIST_ID),
                        is_weak: None,
                    });
            },
            SplitMetadataCorruption::ComponentAndObjectTileExternal => {
                metadata.components[calculation_index]
                    .external_references
                    .push(tsp::ComponentExternalReference {
                        component_identifier: TILE_COMPONENT_ID,
                        object_identifier: Some(TILE_ID),
                        is_weak: Some(false),
                    });
            },
            SplitMetadataCorruption::ComponentAndObjectControlExternal => {
                metadata.components[calculation_index]
                    .external_references
                    .push(tsp::ComponentExternalReference {
                        component_identifier: CONTROL_COMPONENT_ID,
                        object_identifier: Some(CONTROL_LIST_ID),
                        is_weak: Some(false),
                    });
            },
            SplitMetadataCorruption::VersionedCalculationExternal => {
                let reference = metadata.components[calculation_index]
                    .external_references
                    .iter()
                    .find(|reference| reference.component_identifier == FORMAT_COMPONENT_ID)
                    .cloned()
                    .ok_or_else(|| io::Error::other("format external edge is missing"))?;
                metadata.components[calculation_index]
                    .external_references
                    .retain(|candidate| candidate.component_identifier != FORMAT_COMPONENT_ID);
                metadata.components[calculation_index]
                    .versioned_external_references
                    .push(reference);
            },
            SplitMetadataCorruption::VersionedPopupExternal => {
                let control = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == CONTROL_COMPONENT_ID)
                    .ok_or_else(|| io::Error::other("control metadata is missing"))?;
                let reference = control
                    .external_references
                    .iter()
                    .find(|reference| reference.component_identifier == POPUP_COMPONENT_ID)
                    .cloned()
                    .ok_or_else(|| io::Error::other("popup external edge is missing"))?;
                control
                    .external_references
                    .retain(|candidate| candidate.component_identifier != POPUP_COMPONENT_ID);
                control.versioned_external_references.push(reference);
            },
            SplitMetadataCorruption::MissingFormatComponent => {
                metadata
                    .components
                    .retain(|component| component.identifier != FORMAT_COMPONENT_ID);
            },
            SplitMetadataCorruption::DuplicateFormatComponent => {
                let format = metadata
                    .components
                    .iter()
                    .find(|component| component.identifier == FORMAT_COMPONENT_ID)
                    .cloned()
                    .ok_or_else(|| io::Error::other("format metadata is missing"))?;
                metadata.components.push(format);
            },
            SplitMetadataCorruption::WrongCalculationLocator => {
                metadata.components[calculation_index].preferred_locator =
                    "Wrong/CalculationEngine".to_owned();
                metadata.components[calculation_index].locator =
                    Some("Wrong/CalculationEngine".to_owned());
            },
            SplitMetadataCorruption::WrongFormatLocator => {
                let format = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == FORMAT_COMPONENT_ID)
                    .ok_or_else(|| io::Error::other("format metadata is missing"))?;
                format.preferred_locator = "Wrong/Tables/Format".to_owned();
                format.locator = Some("Wrong/Tables/Format".to_owned());
            },
            SplitMetadataCorruption::WrongControlLocator => {
                let control = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == CONTROL_COMPONENT_ID)
                    .ok_or_else(|| io::Error::other("control metadata is missing"))?;
                control.preferred_locator = "Wrong/Tables/Control".to_owned();
                control.locator = Some("Wrong/Tables/Control".to_owned());
            },
            SplitMetadataCorruption::OpaquePopupInbound => unreachable!(),
        }
        Ok(())
    })
}

fn with_locked_table(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(TABLE_INFO_ID)
            .ok_or_else(|| io::Error::other("table info object is missing"))?;
        let message = object
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("table info payload is missing"))?;
        let mut info = tst::TableInfoArchive::decode(message.data.as_slice())?;
        info.super_.locked = Some(true);
        message.data = info.encode_to_vec();
        Ok(())
    })
}

fn with_split_model_reference_corruption(
    source: &[u8],
    corruption: SplitModelReferenceCorruption,
) -> TestResult<Vec<u8>> {
    rewrite_member(source, CALCULATION_MEMBER, |archive| {
        let object = archive
            .object_mut(TABLE_MODEL_ID)
            .ok_or_else(|| io::Error::other("TableModel object is missing"))?;
        let info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("TableModel metadata is missing"))?;
        match corruption {
            SplitModelReferenceCorruption::MissingFormatAggregate => {
                info.object_references
                    .retain(|identifier| *identifier != FORMAT_LIST_ID);
            },
            SplitModelReferenceCorruption::DuplicateFormatAggregate => {
                info.object_references.push(FORMAT_LIST_ID);
            },
            SplitModelReferenceCorruption::WrongFormatFieldType => {
                let mut field = FieldInfo::new(vec![99]);
                field.r#type = Some(FieldType::Value);
                field.object_references.push(FORMAT_LIST_ID);
                info.field_infos.push(field);
            },
            SplitModelReferenceCorruption::DuplicateFormatField => {
                for path in [vec![98], vec![99]] {
                    let mut field = FieldInfo::new(path);
                    field.r#type = Some(FieldType::ObjectReference);
                    field.object_references.push(FORMAT_LIST_ID);
                    info.field_infos.push(field);
                }
            },
        }
        Ok(())
    })
}

fn without_metadata(source: &[u8]) -> TestResult<Vec<u8>> {
    Ok(
        Catalog::from_bytes(source)?.reassemble_with_deletions_to_bytes(
            &[],
            &[METADATA_MEMBER],
            Limits::default(),
        )?,
    )
}

#[test]
fn mixed_controls_read_through_one_typed_facade() -> TestResult {
    let source = fixture(FixtureMode::Mixed)?;
    let package = Package::from_bytes(&source)?;
    let expected = controls()?;
    for (row, expected) in expected.into_iter().enumerate() {
        assert_eq!(
            package.table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::name("Cell Controls"),
                CellPosition::new(row as u32, 0),
            )?,
            Some(expected)
        );
    }
    assert_eq!(
        package.table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 1),
        )?,
        None
    );
    Ok(())
}

#[test]
fn mixed_control_noop_is_byte_exact_and_does_not_touch_siblings() -> TestResult {
    let source = fixture(FixtureMode::Mixed)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(CellControl::Checkbox(Checkbox))
        .commit()?;
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.package().exact_bytes(), source);
    Ok(())
}

#[test]
fn cross_kind_replacement_is_cow_reversible_and_column_aware() -> TestResult {
    let source = fixture(FixtureMode::Mixed)?;
    let package = Package::from_bytes(&source)?;
    let replacement = CellControl::Slider(Slider::new(
        range(-10.0, 30.0, 5.0),
        DisplayFormat::Number(Number::default()),
    ));
    let commit = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(replacement.clone())
        .commit()?;
    let target = commit.package().exact_bytes();
    assert_eq!(
        commit.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?,
        Some(replacement)
    );
    assert_eq!(
        commit.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
        )?,
        Some(CellControl::StarRating(StarRating))
    );
    assert_eq!(
        commit.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 1),
        )?,
        None
    );
    assert_previews_invalidated(&source, &target)?;
    assert!(changed_members(&source, &target)?.contains(&DOCUMENT_MEMBER.to_owned()));
    assert_eq!(
        object_messages(&source, VIEW_STATE_MEMBER, VIEW_STATE_ID)?,
        object_messages(&target, VIEW_STATE_MEMBER, VIEW_STATE_ID)?,
    );
    assert_eq!(
        Package::from_bytes(&source)?
            .apply_table_cell_control_format(commit.patch())?
            .package()
            .exact_bytes(),
        target
    );
    assert!(
        Package::from_bytes(&target)?
            .apply_table_cell_control_format(commit.patch())
            .is_err()
    );
    assert_eq!(
        Package::from_bytes(&target)?
            .apply_table_cell_control_format(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

#[test]
fn every_control_kind_can_replace_an_existing_kind() -> TestResult {
    let source = fixture(FixtureMode::Mixed)?;
    let positions = [
        CellPosition::new(0, 0),
        CellPosition::new(1, 0),
        CellPosition::new(2, 0),
        CellPosition::new(3, 0),
        CellPosition::new(4, 0),
    ];
    for (position, replacement) in positions.into_iter().zip(controls()?) {
        let package = Package::from_bytes(&source)?;
        let commit = package
            .edit_table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )?
            .set(replacement.clone())
            .commit()?;
        assert_eq!(
            commit.package().table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )?,
            Some(replacement)
        );
        assert_eq!(
            Package::from_bytes(&commit.package().exact_bytes())?
                .apply_table_cell_control_format(&commit.patch().inverse())?
                .package()
                .exact_bytes(),
            source
        );
    }
    Ok(())
}

#[test]
fn empty_cells_seed_all_five_control_kinds_at_arbitrary_columns() -> TestResult {
    let source = fixture(FixtureMode::Empty)?;
    let package = Package::from_bytes(&source)?;
    let values = controls()?;
    let positions = [
        CellPosition::new(0, 0),
        CellPosition::new(0, 1),
        CellPosition::new(0, 2),
        CellPosition::new(1, 1),
        CellPosition::new(1, 2),
    ];
    for (position, value) in positions.into_iter().zip(values) {
        let commit = package
            .edit_table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )?
            .set(value.clone())
            .commit()?;
        assert_eq!(
            commit.package().table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )?,
            Some(value)
        );
        assert_control_metadata_exact(&commit.package().exact_bytes())?;
    }
    Ok(())
}

#[test]
fn numeric_scalar_seeding_preserves_cell_position_and_value_route() -> TestResult {
    let source = fixture(FixtureMode::ScalarSeed)?;
    let package = Package::from_bytes(&source)?;
    let replacement = CellControl::Stepper(Stepper::new(
        range(0.0, 20.0, 2.0),
        DisplayFormat::Number(Number::default()),
    ));
    let commit = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(replacement.clone())
        .commit()?;
    assert_eq!(
        commit.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?,
        Some(replacement)
    );
    assert_eq!(
        commit.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 2),
        )?,
        None
    );
    Ok(())
}

#[test]
fn shared_control_entries_are_refcounted_and_final_reset_culls() -> TestResult {
    let source = fixture(FixtureMode::SharedCheckbox)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(control_entries(&source)?[0].refcount, 2);
    assert_eq!(format_entries(&source)?[0].refcount, 2);
    let slider = CellControl::Slider(Slider::default());
    let changed = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(slider.clone())
        .commit()?;
    assert_eq!(
        changed.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
        )?,
        Some(CellControl::Checkbox(Checkbox))
    );
    assert!(control_entries(&changed.package().exact_bytes())?.len() >= 2);
    let reset = changed
        .package()
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .reset()
        .commit()?;
    assert_eq!(
        reset.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?,
        None
    );
    let final_reset = reset
        .package()
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
        )?
        .reset()
        .commit()?;
    assert!(control_entries(&final_reset.package().exact_bytes())?.is_empty());
    assert!(format_entries(&final_reset.package().exact_bytes())?.is_empty());
    assert_control_metadata_exact(&final_reset.package().exact_bytes())?;
    assert_eq!(
        Package::from_bytes(&final_reset.package().exact_bytes())?
            .apply_table_cell_control_format(&final_reset.patch().inverse())?
            .package()
            .exact_bytes(),
        reset.package().exact_bytes()
    );
    Ok(())
}

#[test]
fn clear_and_reset_are_equivalent_automatic_transitions() -> TestResult {
    let source = fixture(FixtureMode::Mixed)?;
    let package = Package::from_bytes(&source)?;
    let cleared = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .clear()
        .commit()?;
    assert_eq!(
        cleared.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?,
        None
    );
    let restored = Package::from_bytes(&cleared.package().exact_bytes())?
        .apply_table_cell_control_format(&cleared.patch().inverse())?;
    assert_eq!(restored.package().exact_bytes(), source);

    let empty = Package::from_bytes(&fixture(FixtureMode::Empty)?)?;
    let noop = empty
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(2, 2),
        )?
        .clear()
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().exact_bytes(), empty.exact_bytes());
    Ok(())
}

#[test]
fn malformed_control_lists_cells_and_references_fail_closed_atomically() -> TestResult {
    let source = fixture(FixtureMode::Mixed)?;
    for corruption in [
        Corruption::WrongInteraction,
        Corruption::MissingCellSpec,
        Corruption::ZeroRefcount,
        Corruption::DuplicateControlKey,
        Corruption::SegmentedControlList,
        Corruption::DuplicateControlList,
        Corruption::ControlAggregateMissing,
        Corruption::ControlAggregateWrong,
        Corruption::ControlFieldInfoMissing,
        Corruption::ControlFieldInfoWrong,
        Corruption::ControlFieldInfoExtra,
        Corruption::ControlKeyOverflow,
        Corruption::MissingPopupModel,
        Corruption::WrongPopupModelType,
        Corruption::DeprecatedReferenceType,
        Corruption::DeprecatedExternalReference,
        Corruption::WrongBncKind,
        Corruption::MissingFormatEntry,
        Corruption::FormatRefcountUndercount,
    ] {
        let hostile = with_corruption(&source, corruption)?;
        assert_read_rejects(&hostile, &format!("{corruption:?}"))?;
        assert_changed_edit_rejects(&hostile, &format!("{corruption:?}"))?;
    }
    Ok(())
}

#[test]
fn format_and_control_refcount_undercounts_and_overcounts_are_rejected() -> TestResult {
    let shared = fixture(FixtureMode::SharedCheckbox)?;
    for corruption in [
        Corruption::FormatRefcountUndercount,
        Corruption::ControlRefcountUndercount,
        Corruption::FormatRefcountOvercount,
        Corruption::ControlRefcountOvercount,
    ] {
        let hostile = with_corruption(&shared, corruption)?;
        assert_changed_edit_rejects(&hostile, &format!("{corruption:?}"))?;
    }
    Ok(())
}

#[test]
fn metadata_uuid_and_registry_collisions_fail_before_control_publication() -> TestResult {
    let source = fixture(FixtureMode::Mixed)?;
    for corruption in [
        MetadataCorruption::MissingControlUuid,
        MetadataCorruption::VersionedOnlyControlUuid,
        MetadataCorruption::DuplicateControlUuid,
        MetadataCorruption::AmbiguousControlIdentifier,
        MetadataCorruption::DataOwnerControlIdentifier,
        MetadataCorruption::RootDataMapControlIdentifier,
    ] {
        let hostile = with_metadata_corruption(&source, corruption)?;
        assert_changed_edit_rejects(&hostile, &format!("{corruption:?}"))?;
    }
    Ok(())
}

#[test]
fn missing_metadata_and_unknown_root_fields_are_atomic() -> TestResult {
    let source = fixture(FixtureMode::Mixed)?;
    let missing = without_metadata(&source)?;
    assert_changed_edit_rejects(&missing, "missing Metadata.iwa")?;

    let unknown = rewrite_metadata_root(&source, |metadata| {
        metadata.data_metadata_map = Some(reference(0xfeed_beef));
        Ok(())
    })?;
    let package = match Package::from_bytes(&unknown) {
        Err(_) => return Ok(()),
        Ok(package) => package,
    };
    let before = package.exact_bytes();
    let result = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(CellControl::StarRating(StarRating))
        .commit();
    if result.is_err() {
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

#[test]
fn explicit_locator_alias_and_reserved_identifier_policies_are_atomic() -> TestResult {
    let source = fixture(FixtureMode::Mixed)?;
    let explicit = with_explicit_locator(&source)?;
    let package = Package::from_bytes(&explicit)?;
    let replacement = CellControl::Checkbox(Checkbox);
    let commit = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(replacement.clone())
        .commit()?;
    assert_eq!(
        commit.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?,
        Some(replacement)
    );
    assert_eq!(
        Package::from_bytes(&commit.package().exact_bytes())?
            .apply_table_cell_control_format(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        explicit
    );

    let aliased = with_physical_alias(&source)?;
    assert_changed_edit_rejects(&aliased, "cross-component physical alias")?;
    let reserved = with_reserved_identifier(&fixture(FixtureMode::Empty)?, 1_001)?;
    let package = Package::from_bytes(&reserved)?;
    let before = package.exact_bytes();
    let result = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(2, 2),
        )?
        .set(CellControl::Checkbox(Checkbox))
        .commit();
    if result.is_err() {
        assert_eq!(package.exact_bytes(), before);
    }
    Ok(())
}

#[test]
fn unknown_control_and_metadata_framing_survives_admitted_rewrite() -> TestResult {
    let source = fixture(FixtureMode::Mixed)?;
    let hostile = rewrite_member(&source, DOCUMENT_MEMBER, |archive| {
        let index = control_message_index(archive)?;
        let object = archive
            .object_mut(SIDECAR_ID)
            .ok_or_else(|| io::Error::other("sidecar object missing"))?;
        append_varint_field(&mut object.messages[index].data, 90, 0x1234)?;
        Ok(())
    })?;
    let package = match Package::from_bytes(&hostile) {
        Err(_) => return Ok(()),
        Ok(package) => package,
    };
    let commit = match package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(CellControl::StarRating(StarRating))
        .commit()
    {
        Err(_) => return Ok(()),
        Ok(commit) => commit,
    };
    let target = commit.package().exact_bytes();
    let payloads = object_messages(&target, DOCUMENT_MEMBER, SIDECAR_ID)?;
    assert!(payloads.iter().any(|(_, payload)| {
        WireView::parse(payload)
            .map(|view| view.fields().any(|field| field.number() == 90))
            .unwrap_or(false)
    }));
    Ok(())
}

#[test]
fn selector_conflict_position_and_public_limits_are_typed_and_atomic() -> TestResult {
    let source = fixture(FixtureMode::Mixed)?;
    let package = Package::from_bytes(&source)?;
    assert!(
        package
            .table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(99, 2),
            )
            .is_err()
    );
    let edit = package.edit_table_cell_control_format(
        SheetSelector::name("Cell Control Sheet"),
        TableSelector::name("Cell Controls"),
        CellPosition::new(1, 2),
    )?;
    let commit = edit.set(CellControl::Checkbox(Checkbox)).commit()?;
    assert!(
        Package::from_bytes(&commit.package().exact_bytes())?
            .apply_table_cell_control_format(commit.patch())
            .is_err()
    );

    let tight = Limits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        Limits::MAX_IWA_STREAM_BYTES,
    )?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(tight, PackageSemanticLimits::default()),
        )
        .is_err()
    );
    let semantic = PackageSemanticLimits::new(1, 1, 1, 1)?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(Limits::default(), semantic),
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn locked_table_refuses_control_mutation_without_source_changes() -> TestResult {
    let source = with_locked_table(&fixture(FixtureMode::Mixed)?)?;
    let package = Package::from_bytes(&source)?;
    let before = package.exact_bytes();
    let result = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(CellControl::Checkbox(Checkbox))
        .commit();
    assert!(result.is_err());
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn assert_split_read_noop(
    source: &[u8],
    position: CellPosition,
    expected_before: CellControl,
) -> TestResult {
    let package = Package::from_bytes(source)?;
    let before = package.exact_bytes();
    let observed = package.table_cell_control_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        position,
    )?;
    assert_eq!(observed, Some(expected_before.clone()));

    let no_op = package
        .edit_table_cell_control_format(SheetSelector::index(0), TableSelector::index(0), position)?
        .set(expected_before.clone())
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().exact_bytes(), source);

    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

#[test]
fn split_components_read_noop_and_popup_transition_remains_atomic() -> TestResult {
    let source = split_component_fixture()?;
    let package = Package::from_bytes(&source)?;
    let no_op = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(controls()?[0].clone())
        .commit()?;
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().exact_bytes(), source);

    let positions = [
        CellPosition::new(0, 0),
        CellPosition::new(1, 0),
        CellPosition::new(2, 0),
        CellPosition::new(3, 0),
        CellPosition::new(4, 0),
    ];
    for (position, expected) in positions.into_iter().zip(controls()?) {
        assert_split_read_noop(&source, position, expected)?;
    }
    let package = Package::from_bytes(&source)?;
    let before = package.exact_bytes();
    assert!(
        package
            .edit_table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(4, 0),
            )?
            .set(CellControl::Checkbox(Checkbox))
            .commit()
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn split_scalar_replacements() -> [CellControl; 4] {
    [
        CellControl::StarRating(StarRating),
        CellControl::Slider(Slider::new(
            range(-20.0, 40.0, 5.0),
            DisplayFormat::Number(Number::default()),
        )),
        CellControl::Stepper(Stepper::new(
            range(2.0, 30.0, 2.0),
            DisplayFormat::Number(Number::default()),
        )),
        CellControl::Checkbox(Checkbox),
    ]
}

#[test]
fn split_scalar_controls_support_cross_kind_cow_inverse_apply_and_locality() -> TestResult {
    let source = split_component_fixture()?;
    assert_split_native_metadata_shape(&source)?;
    let positions = [
        CellPosition::new(0, 0),
        CellPosition::new(1, 0),
        CellPosition::new(2, 0),
        CellPosition::new(3, 0),
    ];
    for (position, replacement) in positions.into_iter().zip(split_scalar_replacements()) {
        let package = Package::from_bytes(&source)?;
        let before = package.table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            position,
        )?;
        assert!(before.is_some());
        let commit = package
            .edit_table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )?
            .set(replacement.clone())
            .commit()?;
        assert_eq!(commit.patch().before(), before.as_ref());
        assert_eq!(commit.patch().after(), Some(&replacement));
        assert_eq!(
            commit.package().table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                position,
            )?,
            Some(replacement.clone()),
        );
        assert!(commit.diagnostics().changed());
        assert!(commit.diagnostics().full_reparse_performed());
        let target = commit.package().exact_bytes();
        let changed = changed_members(&source, &target)?;
        assert_previews_invalidated(&source, &target)?;
        assert_split_scalar_locality(&source, &target, &changed)?;
        assert_split_metadata_transition(&source, &target, &changed)?;
        assert!(
            split_list_entries(
                &target,
                FORMAT_MEMBER,
                FORMAT_LIST_ID,
                tst::table_data_list::ListType::Format
            )?
            .iter()
            .all(|entry| entry.refcount > 0)
        );
        assert!(
            split_list_entries(
                &target,
                CONTROL_MEMBER,
                CONTROL_LIST_ID,
                tst::table_data_list::ListType::ControlCellSpec,
            )?
            .iter()
            .all(|entry| entry.refcount > 0)
        );

        let applied =
            Package::from_bytes(&source)?.apply_table_cell_control_format(commit.patch())?;
        assert_eq!(applied.package().exact_bytes(), target);
        assert!(
            Package::from_bytes(&target)?
                .apply_table_cell_control_format(commit.patch())
                .is_err(),
            "the exact split scalar patch must conflict after publication"
        );
        let inverse = Package::from_bytes(&target)?
            .apply_table_cell_control_format(&commit.patch().inverse())?;
        assert_eq!(inverse.package().exact_bytes(), source);
    }
    Ok(())
}

#[test]
fn split_scalar_create_reset_culls_entries_and_preserves_inverse() -> TestResult {
    let source = split_component_fixture()?;
    assert_split_native_metadata_shape(&source)?;
    let position = CellPosition::new(0, 1);
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            position
        )?,
        None
    );
    let original_format = split_list_entries(
        &source,
        FORMAT_MEMBER,
        FORMAT_LIST_ID,
        tst::table_data_list::ListType::Format,
    )?;
    let original_control = split_list_entries(
        &source,
        CONTROL_MEMBER,
        CONTROL_LIST_ID,
        tst::table_data_list::ListType::ControlCellSpec,
    )?;
    let create = package
        .edit_table_cell_control_format(SheetSelector::index(0), TableSelector::index(0), position)?
        .set(CellControl::Checkbox(Checkbox))
        .commit()?;
    let created_bytes = create.package().exact_bytes();
    assert_eq!(
        create.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            position,
        )?,
        Some(CellControl::Checkbox(Checkbox))
    );
    let changed = changed_members(&source, &created_bytes)?;
    assert_previews_invalidated(&source, &created_bytes)?;
    assert_split_scalar_locality(&source, &created_bytes, &changed)?;
    assert_split_metadata_transition(&source, &created_bytes, &changed)?;
    assert!(
        split_list_entries(
            &created_bytes,
            FORMAT_MEMBER,
            FORMAT_LIST_ID,
            tst::table_data_list::ListType::Format,
        )?
        .iter()
        .all(|entry| entry.refcount > 0)
    );
    assert!(
        split_list_entries(
            &created_bytes,
            CONTROL_MEMBER,
            CONTROL_LIST_ID,
            tst::table_data_list::ListType::ControlCellSpec,
        )?
        .iter()
        .all(|entry| entry.refcount > 0)
    );

    let reset = create
        .package()
        .edit_table_cell_control_format(SheetSelector::index(0), TableSelector::index(0), position)?
        .clear()
        .commit()?;
    let reset_bytes = reset.package().exact_bytes();
    assert_eq!(
        reset.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            position,
        )?,
        None
    );
    assert_eq!(
        split_list_entries(
            &reset_bytes,
            FORMAT_MEMBER,
            FORMAT_LIST_ID,
            tst::table_data_list::ListType::Format,
        )?,
        original_format
    );
    assert_eq!(
        split_list_entries(
            &reset_bytes,
            CONTROL_MEMBER,
            CONTROL_LIST_ID,
            tst::table_data_list::ListType::ControlCellSpec,
        )?,
        original_control
    );
    assert!(
        Package::from_bytes(&reset_bytes)?
            .apply_table_cell_control_format(&reset.patch().inverse())?
            .package()
            .exact_bytes()
            == created_bytes
    );
    assert_eq!(
        Package::from_bytes(&created_bytes)?
            .apply_table_cell_control_format(&create.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

#[test]
fn split_scalar_refcounts_locks_and_limits_remain_atomic() -> TestResult {
    let source = split_component_fixture()?;
    for (member, identifier, list_type, label) in [
        (
            FORMAT_MEMBER,
            FORMAT_LIST_ID,
            tst::table_data_list::ListType::Format,
            "format",
        ),
        (
            CONTROL_MEMBER,
            CONTROL_LIST_ID,
            tst::table_data_list::ListType::ControlCellSpec,
            "control",
        ),
    ] {
        let hostile = rewrite_split_list(&source, member, identifier, list_type, |list| {
            list.entries
                .first_mut()
                .ok_or_else(|| io::Error::other("split list entry is missing"))?
                .refcount = 0;
            Ok(())
        })?;
        let package = Package::from_bytes(&hostile)?;
        let before = package.exact_bytes();
        let result = package
            .edit_table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(0, 0),
            )
            .and_then(|edit| edit.set(CellControl::StarRating(StarRating)).commit());
        assert!(
            result.is_err(),
            "undercounted split {label} list was accepted"
        );
        assert_eq!(package.exact_bytes(), before);
    }

    let locked = with_locked_table(&source)?;
    let package = Package::from_bytes(&locked)?;
    let before = package.exact_bytes();
    assert!(
        package
            .edit_table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(0, 0),
            )?
            .set(CellControl::StarRating(StarRating))
            .commit()
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);

    let tight = Limits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        Limits::MAX_IWA_STREAM_BYTES,
    )?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(tight, PackageSemanticLimits::default()),
        )
        .is_err()
    );
    Ok(())
}

#[test]
fn split_scalar_clear_is_reversible_and_local() -> TestResult {
    let source = split_component_fixture()?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .clear()
        .commit()?;
    let target = commit.package().exact_bytes();
    assert_eq!(
        commit.package().table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?,
        None
    );
    let changed = changed_members(&source, &target)?;
    assert_previews_invalidated(&source, &target)?;
    assert_split_scalar_locality(&source, &target, &changed)?;
    assert_split_metadata_transition(&source, &target, &changed)?;
    assert_eq!(
        Package::from_bytes(&target)?
            .apply_table_cell_control_format(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

#[test]
fn split_components_reject_bad_edges_aliases_and_opaque_inbound_refs_atomically() -> TestResult {
    let source = split_component_fixture()?;
    for corruption in [
        SplitMetadataCorruption::MissingCalculationExternal,
        SplitMetadataCorruption::MissingTileExternal,
        SplitMetadataCorruption::MissingControlExternal,
        SplitMetadataCorruption::MissingPopupExternal,
        SplitMetadataCorruption::DuplicateCalculationExternal,
        SplitMetadataCorruption::DuplicatePopupExternal,
        SplitMetadataCorruption::ComponentAndObjectFormatExternal,
        SplitMetadataCorruption::ComponentAndObjectTileExternal,
        SplitMetadataCorruption::ComponentAndObjectControlExternal,
        SplitMetadataCorruption::VersionedCalculationExternal,
        SplitMetadataCorruption::VersionedPopupExternal,
        SplitMetadataCorruption::DuplicateFormatComponent,
        SplitMetadataCorruption::WrongCalculationLocator,
        SplitMetadataCorruption::WrongFormatLocator,
        SplitMetadataCorruption::WrongControlLocator,
    ] {
        let hostile = with_split_metadata_corruption(&source, corruption)?;
        assert_split_owner_rejects(&hostile, &format!("split edge {corruption:?}"))?;
    }
    // Wave86 deliberately admits read-only compatibility for a uniquely
    // edge-owned sidecar without its own ComponentInfo. Wave88 mutation is
    // stricter because it cannot advance a missing component save token.
    let missing_format =
        with_split_metadata_corruption(&source, SplitMetadataCorruption::MissingFormatComponent)?;
    assert_changed_edit_rejects(&missing_format, "split edge MissingFormatComponent")?;
    for corruption in [
        SplitModelReferenceCorruption::MissingFormatAggregate,
        SplitModelReferenceCorruption::DuplicateFormatAggregate,
        SplitModelReferenceCorruption::WrongFormatFieldType,
        SplitModelReferenceCorruption::DuplicateFormatField,
    ] {
        let hostile = with_split_model_reference_corruption(&source, corruption)?;
        assert_split_owner_rejects(&hostile, &format!("split model edge {corruption:?}"))?;
    }
    let opaque =
        with_split_metadata_corruption(&source, SplitMetadataCorruption::OpaquePopupInbound)?;
    let package = Package::from_bytes(&opaque)?;
    assert_eq!(
        package.table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(4, 0),
        )?,
        Some(CellControl::PopUpMenu(menu()?)),
    );
    let before = package.exact_bytes();
    assert!(
        package
            .edit_table_cell_control_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(4, 0),
            )?
            .clear()
            .commit()
            .is_err()
    );
    assert_eq!(package.exact_bytes(), before);
    for (member, alias) in [
        (CALCULATION_MEMBER, "Index/CalculationAlias.iwa"),
        (TILE_MEMBER, "Index/Tables/TileAlias.iwa"),
        (FORMAT_MEMBER, "Index/Tables/FormatAlias.iwa"),
        (CONTROL_MEMBER, "Index/Tables/ControlAlias.iwa"),
        (POPUP_MEMBER, "Index/Tables/PopupAlias.iwa"),
    ] {
        let aliased = with_physical_alias_for_member(&source, member, alias)?;
        assert_split_owner_rejects(&aliased, &format!("split physical alias {member}"))?;
    }
    Ok(())
}

#[test]
fn popup_only_split_reads_but_popup_changed_route_refuses_atomically() -> TestResult {
    let source = popup_only_split_fixture()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?,
        Some(CellControl::Checkbox(Checkbox)),
    );
    assert_eq!(
        package.table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(4, 0),
        )?,
        Some(CellControl::PopUpMenu(menu()?)),
    );
    let before = package.exact_bytes();
    let result = package
        .edit_table_cell_control_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(4, 0),
        )?
        .set(CellControl::Checkbox(Checkbox))
        .commit();
    assert!(matches!(
        result,
        Err(ControlError::UnsupportedDependency { .. })
    ));
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}
