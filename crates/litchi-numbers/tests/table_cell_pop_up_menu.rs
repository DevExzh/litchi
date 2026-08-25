//! Exact-source integration coverage for Numbers Pop-Up Menu cell controls.
//!
//! The fixture deliberately models the native ownership graph rather than
//! treating the menu as a scalar cell value.  A BNC cell points to a format
//! entry and a control-cell-spec entry; the latter points to one type-6206
//! `TST.PopUpMenuModel`.  The same model is shared by two cells so that the
//! transaction tests exercise reuse, refcounts, and final culling.

#![allow(deprecated)]

use std::io;

use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldType, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsce, tsd, tsk, tsp, tst};
use litchi_numbers::CellPosition;
use litchi_numbers::cell::data_format::pop_up_menu::{InitialSelection, PopUpMenu};
use litchi_numbers::{
    Package, PackageReadOptions, PackageSemanticLimits, SheetSelector, TableSelector,
};
use litchi_numbers_wire::{BncCell, CellDataFormatKind};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const VIEW_STATE_MEMBER: &str = "Index/ViewState.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_TYPE: u32 = 11_006;
const DOCUMENT_TYPE: u32 = 1;
const SHEET_TYPE: u32 = 2;
const TABLE_INFO_TYPE: u32 = 6_000;
const TABLE_MODEL_TYPE: u32 = 6_001;
const TILE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_TYPE: u32 = 6_005;
const POPUP_MODEL_TYPE: u32 = 6_206;

const DOCUMENT_ID: u64 = 1;
const SHEET_ID: u64 = 2;
const TABLE_INFO_ID: u64 = 3;
const TABLE_MODEL_ID: u64 = 4;
const SIDECAR_ID: u64 = 5;
const TILE_ID: u64 = 6;
const POPUP_MODEL_ID: u64 = 50;
const VIEW_STATE_ID: u64 = 300;
const METADATA_OBJECT_ID: u64 = 900;
const FORMAT_KEY: u32 = 1;
const CONTROL_KEY: u32 = 1;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum FixtureMode {
    SharedPopup,
    SinglePopup,
    EmptyCells,
    NonPopupCell,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PopupCorruption {
    DeprecatedItem,
    MissingSentinel,
    WrongSentinel,
    WrongItemType,
    NonCanonicalItem,
    DuplicateModelPayload,
    WrongModelType,
    MissingModel,
    MissingControlReference,
    WrongControlInteraction,
    DeprecatedReferenceType,
    DeprecatedExternalReference,
    ZeroControlRefcount,
    DuplicateControlKey,
    SegmentedControlList,
    DuplicateControlListMessage,
    ControlAggregateMissing,
    ControlFieldInfoMissing,
    ControlAggregateWrong,
    ControlFieldInfoWrong,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum MetadataCorruption {
    MissingPopupUuid,
    VersionedOnlyPopupUuid,
    DuplicatePopupUuid,
    AmbiguousPopupIdentifier,
    DataOwnerPopupIdentifier,
    RootDataMapPopupIdentifier,
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

fn cell(format_identifier: Option<u32>, control_identifier: Option<u32>) -> TestResult<Vec<u8>> {
    let mut cell = BncCell::minimal();
    if let (Some(format_identifier), Some(control_identifier)) =
        (format_identifier, control_identifier)
    {
        cell.set_data_format_identifier(
            format_identifier,
            CellDataFormatKind::PopUpMenu,
            Some(control_identifier),
        )?;
    }
    Ok(cell.encode())
}

fn tile(mode: FixtureMode) -> TestResult<tst::Tile> {
    let popup = matches!(mode, FixtureMode::SharedPopup | FixtureMode::SinglePopup);
    let first = cell(popup.then_some(FORMAT_KEY), popup.then_some(CONTROL_KEY))?;
    let second = if matches!(mode, FixtureMode::NonPopupCell | FixtureMode::SinglePopup) {
        let mut number = BncCell::minimal();
        number.set_number(42.0)?;
        number.encode()
    } else {
        cell(popup.then_some(FORMAT_KEY), popup.then_some(CONTROL_KEY))?
    };
    Ok(tst::Tile {
        max_column: 0,
        max_row: 1,
        num_cells: 2,
        numrows: 2,
        row_infos: vec![
            tst::TileRowInfo {
                tile_row_index: 0,
                cell_count: 1,
                storage_version: Some(5),
                cell_storage_buffer: Some(first.clone()),
                cell_offsets: Some(vec![0, 0]),
                ..Default::default()
            },
            tst::TileRowInfo {
                tile_row_index: 1,
                cell_count: 1,
                storage_version: Some(5),
                cell_storage_buffer: Some(second),
                cell_offsets: Some(vec![0, 0]),
                ..Default::default()
            },
        ],
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        ..Default::default()
    })
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

fn popup_model(items: &[&str]) -> tst::PopUpMenuModel {
    let mut values = vec![tsce::CellValueArchive {
        cell_value_type: tsce::cell_value_archive::CellValueType::NilType as i32,
        ..Default::default()
    }];
    values.extend(items.iter().copied().map(popup_item));
    tst::PopUpMenuModel {
        item: Vec::new(),
        tsce_item: values,
    }
}

fn format_entry() -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key: FORMAT_KEY,
        refcount: 2,
        format: Some(tsk::FormatStructArchive {
            format_type: Some(260),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn control_entry(start_with_first_item: bool) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key: CONTROL_KEY,
        refcount: 2,
        cell_spec: Some(tst::CellSpecArchive {
            interaction_type: 7,
            chooser_control_popup_model: Some(reference(POPUP_MODEL_ID)),
            chooser_control_start_w_first: Some(start_with_first_item),
            ..Default::default()
        }),
        ..Default::default()
    }
}

fn sidecars(mode: FixtureMode) -> TestResult<ArchiveObject> {
    let mut messages = Vec::new();
    for list_type in [
        tst::table_data_list::ListType::String,
        tst::table_data_list::ListType::Formula,
        tst::table_data_list::ListType::Format,
    ] {
        let entries = if matches!(list_type, tst::table_data_list::ListType::Format)
            && matches!(mode, FixtureMode::SharedPopup | FixtureMode::SinglePopup)
        {
            let mut entry = format_entry();
            if matches!(mode, FixtureMode::SinglePopup) {
                entry.refcount = 1;
            }
            vec![entry]
        } else {
            Vec::new()
        };
        messages.push(RawMessage {
            type_: TABLE_DATA_LIST_TYPE,
            data: tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 2,
                entries,
                is_new_for_bnc: Some(true),
                ..Default::default()
            }
            .encode_to_vec(),
        });
    }
    messages.push(RawMessage {
        type_: TABLE_DATA_LIST_TYPE,
        data: tst::TableDataList {
            list_type: tst::table_data_list::ListType::ControlCellSpec as i32,
            next_list_id: 2,
            entries: if matches!(mode, FixtureMode::SharedPopup | FixtureMode::SinglePopup) {
                let mut entry = control_entry(true);
                if matches!(mode, FixtureMode::SinglePopup) {
                    entry.refcount = 1;
                }
                vec![entry]
            } else {
                Vec::new()
            },
            is_new_for_bnc: Some(true),
            ..Default::default()
        }
        .encode_to_vec(),
    });
    let mut sidecars = ArchiveObject::new(SIDECAR_ID, messages)?;
    let control_message_index = sidecars.messages.iter().position(|message| {
        message.type_ == TABLE_DATA_LIST_TYPE
            && tst::TableDataList::decode(message.data.as_slice())
                .map(|list| {
                    list.list_type == tst::table_data_list::ListType::ControlCellSpec as i32
                })
                .unwrap_or(false)
    });
    if let Some(info) =
        control_message_index.and_then(|index| sidecars.archive_info.message_infos.get_mut(index))
        && matches!(mode, FixtureMode::SharedPopup | FixtureMode::SinglePopup)
    {
        info.object_references = vec![POPUP_MODEL_ID];
        let mut field = FieldInfo::new(vec![3, CONTROL_KEY]);
        field.r#type = Some(FieldType::ObjectReference);
        field.object_references = vec![POPUP_MODEL_ID];
        info.field_infos.push(field);
    }
    Ok(sidecars)
}

fn table_model() -> tst::TableModelArchive {
    tst::TableModelArchive {
        table_id: "popup-menu-table-id".to_owned(),
        table_name: "Pop-Up Menus".to_owned(),
        table_style: reference(SIDECAR_ID),
        body_text_style: reference(SIDECAR_ID),
        header_row_text_style: reference(SIDECAR_ID),
        header_column_text_style: reference(SIDECAR_ID),
        footer_row_text_style: reference(SIDECAR_ID),
        body_cell_style: reference(SIDECAR_ID),
        header_row_style: reference(SIDECAR_ID),
        header_column_style: reference(SIDECAR_ID),
        footer_row_style: reference(SIDECAR_ID),
        number_of_rows: 2,
        number_of_columns: 1,
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
    if matches!(mode, FixtureMode::SharedPopup | FixtureMode::SinglePopup) {
        document_ids.push(POPUP_MODEL_ID);
    }
    let document = metadata_component(100, "Document", 9, &document_ids)?;
    let view = metadata_component(300, "ViewState", 7, &[VIEW_STATE_ID])?;
    let versioned = metadata_component(100, "Document", 3, &[999])?;
    let mut data = tsp::PackageMetadata {
        // Keep the watermark above every physical archive object (including
        // the metadata object) so create tests prove allocator ownership
        // rather than relying on an archive-only maximum.
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
        DOCUMENT_TYPE,
        tn::DocumentArchive {
            sheets: vec![reference(SHEET_ID)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    document.archive_info.message_infos[0].object_references = vec![SHEET_ID];

    let mut sheet = object(
        SHEET_ID,
        SHEET_TYPE,
        tn::SheetArchive {
            name: "Pop-Up Menu Sheet".to_owned(),
            drawable_infos: vec![reference(TABLE_INFO_ID)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    sheet.archive_info.message_infos[0].object_references = vec![TABLE_INFO_ID];

    let mut info = object(
        TABLE_INFO_ID,
        TABLE_INFO_TYPE,
        tst::TableInfoArchive {
            super_: tsd::DrawableArchive::default(),
            table_model: reference(TABLE_MODEL_ID),
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    info.archive_info.message_infos[0].object_references = vec![TABLE_MODEL_ID];

    let mut model = object(
        TABLE_MODEL_ID,
        TABLE_MODEL_TYPE,
        table_model().encode_to_vec(),
    )?;
    model.archive_info.message_infos[0].object_references = vec![SIDECAR_ID, TILE_ID];
    let mut model_unknown = FieldInfo::new(vec![99, 1]);
    model_unknown.data_references = vec![700];
    model.archive_info.message_infos[0]
        .field_infos
        .push(model_unknown);

    let mut tile = object(TILE_ID, TILE_TYPE, tile(mode)?.encode_to_vec())?;
    tile.archive_info.message_infos[0].data_references = vec![701];
    let sidecars = sidecars(mode)?;
    let mut objects = vec![document, sheet, info, model, sidecars, tile];
    if matches!(mode, FixtureMode::SharedPopup | FixtureMode::SinglePopup) {
        let mut popup = object(
            POPUP_MODEL_ID,
            POPUP_MODEL_TYPE,
            popup_model(&["Low", "Medium", "High"]).encode_to_vec(),
        )?;
        popup.archive_info.message_infos[0].data_references = vec![702];
        objects.push(popup);
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
            ("preview.jpg", b"popup preview".as_slice()),
            ("preview-micro.jpg", b"popup micro".as_slice()),
            ("preview-web.jpg", b"popup web".as_slice()),
            ("Data/sentinel.bin", b"unrelated popup data".as_slice()),
        ],
        Limits::default(),
    )?)
}

fn popup<T: AsRef<str>>(items: &[T], selection: InitialSelection) -> TestResult<PopUpMenu> {
    Ok(PopUpMenu::new(items)?.with_initial_selection(selection))
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

fn object_messages(
    source: &[u8],
    member: &str,
    identifier: u64,
) -> TestResult<Vec<(u32, Vec<u8>)>> {
    let archive = member_archive(source, member)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other(format!("object {identifier} is missing")))?;
    Ok(object
        .messages
        .iter()
        .map(|message| (message.type_, message.data.clone()))
        .collect())
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

fn rewrite_popup_model(
    source: &[u8],
    mut rewrite: impl FnMut(&mut tst::PopUpMenuModel) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(POPUP_MODEL_ID)
            .ok_or_else(|| io::Error::other("popup model is missing"))?;
        let message = object
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("popup model payload is missing"))?;
        let mut model = tst::PopUpMenuModel::decode(message.data.as_slice())?;
        rewrite(&mut model)?;
        message.data = model.encode_to_vec();
        Ok(())
    })
}

fn rewrite_control_list(
    source: &[u8],
    rewrite: impl FnMut(&mut tst::TableDataList) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_sidecar_list(
        source,
        tst::table_data_list::ListType::ControlCellSpec,
        rewrite,
    )
}

fn rewrite_format_list(
    source: &[u8],
    rewrite: impl FnMut(&mut tst::TableDataList) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_sidecar_list(source, tst::table_data_list::ListType::Format, rewrite)
}

fn rewrite_sidecar_list(
    source: &[u8],
    list_type: tst::table_data_list::ListType,
    mut rewrite: impl FnMut(&mut tst::TableDataList) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(SIDECAR_ID)
            .ok_or_else(|| io::Error::other("sidecar object is missing"))?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| {
                message.type_ == TABLE_DATA_LIST_TYPE
                    && tst::TableDataList::decode(message.data.as_slice())
                        .map(|list| list.list_type == list_type as i32)
                        .unwrap_or(false)
            })
            .ok_or_else(|| io::Error::other("requested sidecar list is missing"))?;
        let mut list = tst::TableDataList::decode(message.data.as_slice())?;
        rewrite(&mut list)?;
        message.data = list.encode_to_vec();
        Ok(())
    })
}

fn rewrite_control_archive(
    source: &[u8],
    mut rewrite: impl FnMut(&mut litchi_iwa_core::MessageInfo) -> TestResult,
) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(SIDECAR_ID)
            .ok_or_else(|| io::Error::other("sidecar object is missing"))?;
        let control_index = object
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
            .ok_or_else(|| io::Error::other("control list is missing"))?;
        let info = object
            .archive_info
            .message_infos
            .get_mut(control_index)
            .ok_or_else(|| io::Error::other("control list metadata is missing"))?;
        rewrite(info)
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

fn with_metadata_corruption(source: &[u8], corruption: MetadataCorruption) -> TestResult<Vec<u8>> {
    rewrite_metadata_root(source, |metadata| {
        match corruption {
            MetadataCorruption::MissingPopupUuid => {
                let document = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 100)
                    .ok_or_else(|| io::Error::other("Document component is missing"))?;
                document
                    .object_uuid_map_entries
                    .retain(|entry| entry.identifier != POPUP_MODEL_ID);
            },
            MetadataCorruption::VersionedOnlyPopupUuid => {
                let document = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 100)
                    .ok_or_else(|| io::Error::other("Document component is missing"))?;
                document
                    .object_uuid_map_entries
                    .retain(|entry| entry.identifier != POPUP_MODEL_ID);
                let versioned = metadata
                    .versioned_components
                    .iter_mut()
                    .find(|component| component.identifier == 100)
                    .ok_or_else(|| io::Error::other("versioned Document is missing"))?;
                versioned
                    .object_uuid_map_entries
                    .push(uuid_entry(POPUP_MODEL_ID));
            },
            MetadataCorruption::DuplicatePopupUuid => {
                let document = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 100)
                    .ok_or_else(|| io::Error::other("Document component is missing"))?;
                document
                    .object_uuid_map_entries
                    .push(uuid_entry(POPUP_MODEL_ID));
            },
            MetadataCorruption::AmbiguousPopupIdentifier => {
                let document = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 100)
                    .ok_or_else(|| io::Error::other("Document component is missing"))?;
                document.ambiguous_object_identifiers.push(POPUP_MODEL_ID);
            },
            MetadataCorruption::DataOwnerPopupIdentifier => {
                let document = metadata
                    .components
                    .iter_mut()
                    .find(|component| component.identifier == 100)
                    .ok_or_else(|| io::Error::other("Document component is missing"))?;
                document.data_references.push(tsp::ComponentDataReference {
                    data_identifier: 1_001,
                    object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                        object_identifier: POPUP_MODEL_ID,
                        count: 1,
                    }],
                });
            },
            MetadataCorruption::RootDataMapPopupIdentifier => {
                metadata.data_metadata_map = Some(reference(POPUP_MODEL_ID));
            },
        }
        Ok(())
    })
}

fn metadata_payload(source: &[u8]) -> TestResult<Vec<u8>> {
    Ok(member_archive(source, METADATA_MEMBER)?
        .object(METADATA_OBJECT_ID)
        .ok_or_else(|| io::Error::other("metadata object is missing"))?
        .messages
        .iter()
        .find(|message| message.type_ == METADATA_TYPE)
        .ok_or_else(|| io::Error::other("metadata payload is missing"))?
        .data
        .clone())
}

fn metadata_root(source: &[u8]) -> TestResult<tsp::PackageMetadata> {
    Ok(tsp::PackageMetadata::decode(
        metadata_payload(source)?.as_slice(),
    )?)
}

fn document_metadata_component(source: &[u8], identifier: u64) -> TestResult<tsp::ComponentInfo> {
    for field in WireView::parse(&metadata_payload(source)?)?.fields() {
        if field.number() != 3 {
            continue;
        }
        let component = tsp::ComponentInfo::decode(field.payload())?;
        if component.identifier == identifier && component.locator.as_deref() == Some("Document") {
            return Ok(component);
        }
    }
    Err(io::Error::other("current Document metadata component is missing").into())
}

fn changed_members(source: &[u8], target: &[u8]) -> TestResult<Vec<String>> {
    let source = Catalog::from_bytes(source)?;
    let target = Catalog::from_bytes(target)?;
    let mut changed = Vec::new();
    for before in source.iter() {
        let Some(after) = target.iter().find(|entry| entry.name() == before.name()) else {
            if before.name().starts_with("preview") {
                continue;
            }
            return Err(io::Error::other(format!("member {} disappeared", before.name())).into());
        };
        if before.data() != after.data() {
            changed.push(before.name().to_owned());
        }
    }
    Ok(changed)
}

fn assert_previews_invalidated(source: &[u8], target: &[u8]) -> TestResult {
    let source = Catalog::from_bytes(source)?;
    let target = Catalog::from_bytes(target)?;
    for name in ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"] {
        assert!(
            source.iter().any(|entry| entry.name() == name),
            "source preview {name} is missing"
        );
        assert!(
            target.iter().all(|entry| entry.name() != name),
            "candidate retained preview {name}"
        );
    }
    Ok(())
}

fn assert_popup_read_rejects(source: &[u8], label: &str) -> TestResult {
    match Package::from_bytes(source) {
        Err(_) => Ok(()),
        Ok(package) => match package.table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        ) {
            Err(_) => Ok(()),
            Ok(_) => Err(io::Error::other(format!(
                "hostile Pop-Up Menu graph was accepted: {label}"
            ))
            .into()),
        },
    }
}

fn assert_changed_edit_rejects_atomically(source: &[u8], label: &str) -> TestResult {
    let package = match Package::from_bytes(source) {
        Err(_) => return Ok(()),
        Ok(package) => package,
    };
    let before = package.exact_bytes();
    let menu = popup(&["Changed", "Menu"], InitialSelection::Blank)?;
    let edit = match package.edit_table_cell_pop_up_menu_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
    ) {
        Err(_) => {
            assert_eq!(package.exact_bytes(), before);
            return Ok(());
        },
        Ok(edit) => edit,
    };
    assert!(
        edit.set(menu).commit().is_err(),
        "hostile Pop-Up Menu graph unexpectedly published: {label}"
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn assert_reset_rejects_atomically(source: &[u8], label: &str) -> TestResult {
    let package = match Package::from_bytes(source) {
        Err(_) => return Ok(()),
        Ok(package) => package,
    };
    let before = package.exact_bytes();
    let edit = match package.edit_table_cell_pop_up_menu_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
    ) {
        Err(_) => {
            assert_eq!(package.exact_bytes(), before);
            return Ok(());
        },
        Ok(edit) => edit,
    };
    assert!(
        edit.reset().commit().is_err(),
        "hostile Pop-Up Menu reset unexpectedly published: {label}"
    );
    assert_eq!(package.exact_bytes(), before);
    Ok(())
}

fn with_corruption(source: &[u8], corruption: PopupCorruption) -> TestResult<Vec<u8>> {
    match corruption {
        PopupCorruption::DeprecatedItem => rewrite_popup_model(source, |model| {
            model.item.push(tst::pop_up_menu_model::CellValue {
                cell_value_type: tst::pop_up_menu_model::CellValueType::StringType as i32,
                ..Default::default()
            });
            Ok(())
        }),
        PopupCorruption::MissingSentinel => rewrite_popup_model(source, |model| {
            model.tsce_item.remove(0);
            Ok(())
        }),
        PopupCorruption::WrongSentinel => rewrite_popup_model(source, |model| {
            model.tsce_item[0].cell_value_type =
                tsce::cell_value_archive::CellValueType::StringType as i32;
            model.tsce_item[0].string_value = Some(tsce::StringCellValueArchive {
                value: "not nil".to_owned(),
                format: tsk::FormatStructArchive {
                    format_type: Some(260),
                    ..Default::default()
                },
                ..Default::default()
            });
            Ok(())
        }),
        PopupCorruption::WrongItemType => rewrite_popup_model(source, |model| {
            model.tsce_item[1].cell_value_type =
                tsce::cell_value_archive::CellValueType::BooleanType as i32;
            model.tsce_item[1].string_value = None;
            model.tsce_item[1].boolean_value = Some(tsce::BooleanCellValueArchive {
                value: true,
                ..Default::default()
            });
            Ok(())
        }),
        PopupCorruption::NonCanonicalItem => rewrite_popup_model(source, |model| {
            model.tsce_item[1]
                .string_value
                .as_mut()
                .expect("fixture item has a string")
                .is_regex = Some(true);
            Ok(())
        }),
        PopupCorruption::DuplicateModelPayload => {
            rewrite_member(source, DOCUMENT_MEMBER, |archive| {
                let object = archive
                    .object_mut(POPUP_MODEL_ID)
                    .ok_or_else(|| io::Error::other("popup model is missing"))?;
                let duplicate = object
                    .messages
                    .first()
                    .cloned()
                    .ok_or_else(|| io::Error::other("popup model payload is missing"))?;
                let duplicate_info = object
                    .archive_info
                    .message_infos
                    .first()
                    .cloned()
                    .ok_or_else(|| io::Error::other("popup model metadata is missing"))?;
                object.messages.push(duplicate);
                object.archive_info.message_infos.push(duplicate_info);
                Ok(())
            })
        },
        PopupCorruption::WrongModelType => rewrite_member(source, DOCUMENT_MEMBER, |archive| {
            let object = archive
                .object_mut(POPUP_MODEL_ID)
                .ok_or_else(|| io::Error::other("popup model is missing"))?;
            object.messages[0].type_ = POPUP_MODEL_TYPE + 1;
            Ok(())
        }),
        PopupCorruption::MissingModel => rewrite_member(source, DOCUMENT_MEMBER, |archive| {
            archive
                .objects
                .retain(|object| object.archive_info.identifier != Some(POPUP_MODEL_ID));
            Ok(())
        }),
        PopupCorruption::MissingControlReference => rewrite_control_list(source, |list| {
            list.entries[0]
                .cell_spec
                .as_mut()
                .expect("fixture control spec exists")
                .chooser_control_popup_model = None;
            Ok(())
        }),
        PopupCorruption::WrongControlInteraction => rewrite_control_list(source, |list| {
            list.entries[0]
                .cell_spec
                .as_mut()
                .expect("fixture control spec exists")
                .interaction_type = 8;
            Ok(())
        }),
        PopupCorruption::DeprecatedReferenceType => rewrite_control_list(source, |list| {
            list.entries[0]
                .cell_spec
                .as_mut()
                .expect("fixture control spec exists")
                .chooser_control_popup_model
                .as_mut()
                .expect("fixture popup reference exists")
                .deprecated_type = Some(1);
            Ok(())
        }),
        PopupCorruption::DeprecatedExternalReference => rewrite_control_list(source, |list| {
            list.entries[0]
                .cell_spec
                .as_mut()
                .expect("fixture control spec exists")
                .chooser_control_popup_model
                .as_mut()
                .expect("fixture popup reference exists")
                .deprecated_is_external = Some(true);
            Ok(())
        }),
        PopupCorruption::ZeroControlRefcount => rewrite_control_list(source, |list| {
            list.entries[0].refcount = 0;
            Ok(())
        }),
        PopupCorruption::DuplicateControlKey => rewrite_control_list(source, |list| {
            let duplicate = list
                .entries
                .first()
                .cloned()
                .ok_or_else(|| io::Error::other("control entry is missing"))?;
            list.entries.push(duplicate);
            Ok(())
        }),
        PopupCorruption::SegmentedControlList => rewrite_control_list(source, |list| {
            list.segments.push(reference(999));
            Ok(())
        }),
        PopupCorruption::DuplicateControlListMessage => {
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
                                .map(|list| {
                                    list.list_type
                                        == tst::table_data_list::ListType::ControlCellSpec as i32
                                })
                                .unwrap_or(false)
                    })
                    .ok_or_else(|| io::Error::other("control list message is missing"))?;
                let message = object
                    .messages
                    .get(index)
                    .cloned()
                    .ok_or_else(|| io::Error::other("control list payload is missing"))?;
                let info = object
                    .archive_info
                    .message_infos
                    .get(index)
                    .cloned()
                    .ok_or_else(|| io::Error::other("control list metadata is missing"))?;
                object.messages.push(message);
                object.archive_info.message_infos.push(info);
                Ok(())
            })
        },
        PopupCorruption::ControlAggregateMissing => rewrite_control_archive(source, |info| {
            info.object_references.clear();
            Ok(())
        }),
        PopupCorruption::ControlFieldInfoMissing => rewrite_control_archive(source, |info| {
            info.field_infos.clear();
            Ok(())
        }),
        PopupCorruption::ControlAggregateWrong => rewrite_control_archive(source, |info| {
            info.object_references = vec![999];
            Ok(())
        }),
        PopupCorruption::ControlFieldInfoWrong => rewrite_control_archive(source, |info| {
            let field = info
                .field_infos
                .first_mut()
                .ok_or_else(|| io::Error::other("control field info is missing"))?;
            field.object_references = vec![999];
            Ok(())
        }),
    }
}

fn append_unknown_popup_field(source: &[u8], field: u32, value: u64) -> TestResult<Vec<u8>> {
    rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(POPUP_MODEL_ID)
            .ok_or_else(|| io::Error::other("popup model is missing"))?;
        let message = object
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("popup model payload is missing"))?;
        append_varint_field(&mut message.data, field, value)?;
        Ok(())
    })
}

fn append_unknown_popup_wire(source: &[u8]) -> TestResult<(Vec<u8>, Vec<u8>)> {
    // Field 90 carries an intentionally overlong zero value.  Field 91 is a
    // balanced unknown group containing one unknown scalar.  The strict
    // popup codec may admit these bytes as opaque future data; when it does,
    // the complete framing must survive the selected rewrite.
    let raw = vec![
        0xd0, 0x05, 0x80, 0x00, // field 90, overlong varint value
        0xdb, 0x05, 0x08, 0x07, 0xdc, 0x05, // field 91 unknown group
    ];
    let hostile = rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        let object = archive
            .object_mut(POPUP_MODEL_ID)
            .ok_or_else(|| io::Error::other("popup model is missing"))?;
        let message = object
            .messages
            .first_mut()
            .ok_or_else(|| io::Error::other("popup model payload is missing"))?;
        message.data.extend_from_slice(&raw);
        Ok(())
    })?;
    Ok((hostile, raw))
}

fn with_physical_popup_alias(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let alias_data = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or_else(|| io::Error::other("Document member is missing"))?
        .data()
        .to_vec();
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<Vec<_>>();
    entries.push(("Index/PopupAlias.iwa".to_owned(), alias_data));
    Ok(litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        Limits::default(),
    )?)
}

fn with_refcount_undercount(
    source: &[u8],
    undercount_format: bool,
    undercount_control: bool,
) -> TestResult<Vec<u8>> {
    let source = rewrite_format_list(source, |list| {
        let entry = list
            .entries
            .first_mut()
            .ok_or_else(|| io::Error::other("format entry is missing"))?;
        if undercount_format {
            entry.refcount = 1;
        }
        Ok(())
    })?;
    rewrite_control_list(&source, |list| {
        let entry = list
            .entries
            .first_mut()
            .ok_or_else(|| io::Error::other("control entry is missing"))?;
        if undercount_control {
            entry.refcount = 1;
        }
        Ok(())
    })
}

fn with_orphan_popup_model(source: &[u8], versioned_uuid: bool) -> TestResult<Vec<u8>> {
    let source = rewrite_member(source, DOCUMENT_MEMBER, |archive| {
        archive.objects.push(object(
            POPUP_MODEL_ID,
            POPUP_MODEL_TYPE,
            popup_model(&["Low", "Medium", "High"]).encode_to_vec(),
        )?);
        Ok(())
    })?;
    if versioned_uuid {
        with_metadata_corruption(&source, MetadataCorruption::VersionedOnlyPopupUuid)
    } else {
        Ok(source)
    }
}

fn with_cross_component_popup_inbound(source: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(source)?;
    let mut entries = catalog
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<Vec<_>>();
    let mut inbound = object(700, 9_999, b"opaque inbound owner".to_vec())?;
    inbound.archive_info.message_infos[0].object_references = vec![POPUP_MODEL_ID];
    let mut opaque_field = FieldInfo::new(vec![99, 7]);
    opaque_field.object_references = vec![POPUP_MODEL_ID];
    inbound.archive_info.message_infos[0]
        .field_infos
        .push(opaque_field);
    let inbound_bytes = SnappyStream::compress(
        &Archive {
            objects: vec![inbound],
        }
        .to_bytes()?,
    )?;
    entries.push(("Index/PopupInbound.iwa".to_owned(), inbound_bytes));
    let source = litchi_iwa_archive::package::to_bytes(
        entries
            .iter()
            .map(|(name, data)| (name.as_str(), data.as_slice())),
        Limits::default(),
    )?;
    rewrite_metadata_root(&source, |metadata| {
        metadata.components.push(tsp::ComponentInfo {
            identifier: 400,
            preferred_locator: "PopupInbound".to_owned(),
            locator: Some("PopupInbound".to_owned()),
            save_token: Some(9),
            object_uuid_map_entries: vec![uuid_entry(700)],
            ..Default::default()
        });
        Ok(())
    })
}

fn with_component_uuid_mutation(
    source: &[u8],
    identifier: u64,
    versioned: bool,
) -> TestResult<Vec<u8>> {
    rewrite_metadata_root(source, |metadata| {
        let document = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .ok_or_else(|| io::Error::other("Document component is missing"))?;
        document
            .object_uuid_map_entries
            .retain(|entry| entry.identifier != identifier);
        if versioned {
            let versioned_document = metadata
                .versioned_components
                .iter_mut()
                .find(|component| component.identifier == 100)
                .ok_or_else(|| io::Error::other("versioned Document is missing"))?;
            versioned_document
                .object_uuid_map_entries
                .push(uuid_entry(identifier));
        }
        Ok(())
    })
}

fn with_explicit_document_locator(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_metadata_root(source, |metadata| {
        let document = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .ok_or_else(|| io::Error::other("Document component is missing"))?;
        document.preferred_locator = "Document-preferred-alias".to_owned();
        document.locator = Some("Document".to_owned());
        Ok(())
    })
}

fn with_extra_control_aggregate(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_control_archive(source, |info| {
        info.object_references.push(999);
        Ok(())
    })
}

fn with_duplicate_control_field_info(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_control_archive(source, |info| {
        let field = info
            .field_infos
            .first()
            .cloned()
            .ok_or_else(|| io::Error::other("control field info is missing"))?;
        info.field_infos.push(field);
        Ok(())
    })
}

fn with_dangling_format_route(source: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_format_list(source, |list| {
        list.entries.clear();
        Ok(())
    })
}

fn control_message_info(source: &[u8]) -> TestResult<litchi_iwa_core::MessageInfo> {
    let archive = member_archive(source, DOCUMENT_MEMBER)?;
    let object = archive
        .object(SIDECAR_ID)
        .ok_or_else(|| io::Error::other("sidecar object is missing"))?;
    for (index, message) in object.messages.iter().enumerate() {
        if message.type_ != TABLE_DATA_LIST_TYPE {
            continue;
        }
        let list = tst::TableDataList::decode(message.data.as_slice())?;
        if list.list_type == tst::table_data_list::ListType::ControlCellSpec as i32 {
            return object
                .archive_info
                .message_infos
                .get(index)
                .cloned()
                .ok_or_else(|| io::Error::other("control message metadata is missing").into());
        }
    }
    Err(io::Error::other("control-cell-spec list is missing").into())
}

fn assert_control_metadata_matches(source: &[u8]) -> TestResult {
    let entries = control_entries(source)?;
    let info = control_message_info(source)?;
    let mut expected = Vec::with_capacity(entries.len());
    for entry in &entries {
        let spec = entry
            .cell_spec
            .as_ref()
            .ok_or_else(|| io::Error::other("control entry payload is missing"))?;
        let popup = spec
            .chooser_control_popup_model
            .as_ref()
            .ok_or_else(|| io::Error::other("control popup reference is missing"))?;
        expected.push((entry.key, popup.identifier));
    }
    let mut expected_aggregate = expected
        .iter()
        .map(|(_, identifier)| *identifier)
        .collect::<Vec<_>>();
    expected_aggregate.sort_unstable();
    expected_aggregate.dedup();
    let mut aggregate = info.object_references.clone();
    aggregate.sort_unstable();
    aggregate.dedup();
    assert_eq!(aggregate, expected_aggregate);
    assert_eq!(info.object_references.len(), aggregate.len());
    for (key, identifier) in expected {
        let matches = info
            .field_infos
            .iter()
            .filter(|field| field.path.as_slice() == [3, key])
            .collect::<Vec<_>>();
        assert_eq!(matches.len(), 1, "control FieldInfo path [{key}] count");
        assert_eq!(matches[0].object_references.as_slice(), [identifier]);
    }
    assert!(!info.field_infos.iter().any(|field| {
        field.path.as_slice().first() == Some(&3)
            && !entries
                .iter()
                .any(|entry| field.path.as_slice() == [3, entry.key])
    }));
    Ok(())
}

fn assert_unchanged_document_objects(
    source: &[u8],
    target: &[u8],
    identifiers: &[u64],
) -> TestResult {
    let source_archive = member_archive(source, DOCUMENT_MEMBER)?;
    let target_archive = member_archive(target, DOCUMENT_MEMBER)?;
    for identifier in identifiers {
        assert_eq!(
            source_archive.object(*identifier),
            target_archive.object(*identifier),
            "unselected Document object {identifier} changed"
        );
    }
    Ok(())
}

fn reserve_metadata_identifier(source: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
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
        let component = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 100)
            .ok_or_else(|| io::Error::other("Document component is missing"))?;
        component.ambiguous_object_identifiers.push(identifier);
        message.data = metadata.encode_to_vec();
        Ok(())
    })
}

#[test]
fn shared_popup_read_noop_replace_inverse_apply_and_locality() -> TestResult {
    let source = fixture(FixtureMode::SharedPopup)?;
    let package = Package::from_bytes(&source)?;
    let expected = popup(&["Low", "Medium", "High"], InitialSelection::FirstItem)?;
    assert_eq!(
        package.table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?,
        Some(expected.clone())
    );
    assert_eq!(
        package.table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::name("Pop-Up Menus"),
            CellPosition::new(1, 0),
        )?,
        Some(expected.clone())
    );

    let noop = package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(expected.clone())
        .commit()?;
    assert!(noop.patch().is_noop());
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.package().exact_bytes(), source);

    let replacement = popup(&["Draft", "Published"], InitialSelection::Blank)?;
    let commit = package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(replacement.clone())
        .commit()?;
    let target = commit.package().exact_bytes();
    assert_eq!(commit.diagnostics().deleted_previews(), 3);
    assert_eq!(commit.diagnostics().touched_components(), 2);
    assert!(commit.diagnostics().full_reparse_performed());
    assert_previews_invalidated(&source, &target)?;
    assert_eq!(
        commit.package().table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?,
        Some(replacement)
    );
    // The second cell still owns the original control entry/model.  A
    // replacement must therefore copy-on-write rather than mutating the
    // shared model in place.
    assert_eq!(
        commit.package().table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
        )?,
        Some(expected)
    );
    assert_eq!(popup_model_count(&target)?, 2);
    assert_eq!(
        Package::from_bytes(&source)?
            .apply_table_cell_pop_up_menu_format(commit.patch())?
            .package()
            .exact_bytes(),
        target
    );
    assert!(
        Package::from_bytes(&target)?
            .apply_table_cell_pop_up_menu_format(commit.patch())
            .is_err()
    );
    assert_eq!(
        Package::from_bytes(&target)?
            .apply_table_cell_pop_up_menu_format(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    assert_eq!(
        object_messages(&source, VIEW_STATE_MEMBER, VIEW_STATE_ID)?,
        object_messages(&target, VIEW_STATE_MEMBER, VIEW_STATE_ID)?,
    );
    let changed = changed_members(&source, &target)?;
    assert!(changed.iter().any(|member| member == DOCUMENT_MEMBER));
    assert!(changed.iter().any(|member| member == METADATA_MEMBER));
    assert!(metadata_root(&target)?.save_token > metadata_root(&source)?.save_token);
    Ok(())
}

#[test]
fn popup_create_reuses_shared_model_and_final_reset_culls() -> TestResult {
    let source = fixture(FixtureMode::EmptyCells)?;
    let package = Package::from_bytes(&source)?;
    let menu = popup(&["Low", "Medium", "High"], InitialSelection::FirstItem)?;
    assert_eq!(
        package.table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?,
        None
    );

    let first = package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(menu.clone())
        .commit()?;
    let first_bytes = first.package().exact_bytes();
    assert_eq!(first.diagnostics().deleted_previews(), 3);
    assert_eq!(first.diagnostics().touched_components(), 2);
    assert_previews_invalidated(&source, &first_bytes)?;
    assert_eq!(popup_model_count(&first_bytes)?, 1);
    assert_eq!(control_entries(&first_bytes)?.len(), 1);
    let first_model_ids = popup_model_ids(&first_bytes)?;
    assert!(first_model_ids.iter().all(|identifier| *identifier > 1_000));
    let first_document = document_metadata_component(&first_bytes, 100)?;
    assert!(first_model_ids.iter().all(|identifier| {
        first_document
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == *identifier)
    }));

    let second = first
        .package()
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
        )?
        .set(menu.clone())
        .commit()?;
    let second_bytes = second.package().exact_bytes();
    assert_eq!(popup_model_count(&second_bytes)?, 1);
    assert_eq!(control_entries(&second_bytes)?[0].refcount, 2);
    assert_control_metadata_matches(&second_bytes)?;

    let blank = popup(&["Low", "Medium", "High"], InitialSelection::Blank)?;
    let changed_selection = second
        .package()
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(blank)
        .commit()?;
    let changed_selection_bytes = changed_selection.package().exact_bytes();
    assert_eq!(popup_model_count(&changed_selection_bytes)?, 1);
    assert_eq!(control_entries(&changed_selection_bytes)?.len(), 2);
    assert_control_metadata_matches(&changed_selection_bytes)?;

    let reset_first = changed_selection
        .package()
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .reset()
        .commit()?;
    let reset_first_bytes = reset_first.package().exact_bytes();
    assert_eq!(popup_model_count(&reset_first_bytes)?, 1);
    assert!(
        control_entries(&reset_first_bytes)?
            .iter()
            .any(|entry| entry.refcount == 1)
    );

    let reset_second = reset_first
        .package()
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
        )?
        .reset()
        .commit()?;
    let final_bytes = reset_second.package().exact_bytes();
    assert_eq!(popup_model_count(&final_bytes)?, 0);
    assert!(control_entries(&final_bytes)?.is_empty());
    assert!(popup_model_ids(&final_bytes)?.is_empty());
    let final_control_info = control_message_info(&final_bytes)?;
    assert!(
        !final_control_info
            .object_references
            .contains(&POPUP_MODEL_ID)
    );
    assert!(
        final_control_info
            .field_infos
            .iter()
            .all(|field| !field.object_references.contains(&POPUP_MODEL_ID))
    );
    assert_control_metadata_matches(&final_bytes)?;
    let final_document = document_metadata_component(&final_bytes, 100)?;
    assert!(
        !final_document
            .object_uuid_map_entries
            .iter()
            .any(|entry| entry.identifier == POPUP_MODEL_ID)
    );
    assert_eq!(
        Package::from_bytes(&final_bytes)?
            .apply_table_cell_pop_up_menu_format(&reset_second.patch().inverse())?
            .package()
            .exact_bytes(),
        reset_first_bytes
    );
    Ok(())
}

#[test]
fn popup_allocator_respects_metadata_ambiguous_identifiers() -> TestResult {
    let source = reserve_metadata_identifier(&fixture(FixtureMode::EmptyCells)?, 1_001)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(popup(&["Reserved", "Safe"], InitialSelection::FirstItem)?)
        .commit()?;
    let ids = popup_model_ids(&commit.package().exact_bytes())?;
    assert_eq!(ids.len(), 1);
    assert!(ids[0] > 1_001);
    assert_ne!(ids[0], 1_001);
    Ok(())
}

#[test]
fn popup_selector_conflict_and_non_popup_cells_are_typed_and_atomic() -> TestResult {
    let source = fixture(FixtureMode::EmptyCells)?;
    let package = Package::from_bytes(&source)?;
    let menu = popup(&["One", "Two"], InitialSelection::FirstItem)?;
    let edit = package.edit_table_cell_pop_up_menu_format(
        SheetSelector::name("Pop-Up Menu Sheet"),
        TableSelector::name("Pop-Up Menus"),
        CellPosition::new(0, 0),
    )?;
    let commit = edit.set(menu).commit()?;
    assert!(
        Package::from_bytes(&commit.package().exact_bytes())?
            .apply_table_cell_pop_up_menu_format(commit.patch())
            .is_err()
    );
    assert!(
        package
            .table_cell_pop_up_menu_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(99, 0),
            )
            .is_err()
    );

    let non_popup_source = fixture(FixtureMode::NonPopupCell)?;
    let non_popup = Package::from_bytes(&non_popup_source)?;
    assert_eq!(
        non_popup.table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(1, 0),
        )?,
        None
    );
    let before = non_popup.exact_bytes();
    assert!(
        non_popup
            .edit_table_cell_pop_up_menu_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(1, 0),
            )?
            .set(popup(&["One", "Two"], InitialSelection::FirstItem)?)
            .commit()
            .is_err()
    );
    assert_eq!(non_popup.exact_bytes(), before);
    Ok(())
}

#[test]
fn popup_hostile_model_and_control_wires_fail_closed_atomically() -> TestResult {
    let source = fixture(FixtureMode::SharedPopup)?;
    for corruption in [
        PopupCorruption::DeprecatedItem,
        PopupCorruption::MissingSentinel,
        PopupCorruption::WrongSentinel,
        PopupCorruption::WrongItemType,
        PopupCorruption::NonCanonicalItem,
        PopupCorruption::DuplicateModelPayload,
        PopupCorruption::WrongModelType,
        PopupCorruption::MissingModel,
        PopupCorruption::MissingControlReference,
        PopupCorruption::WrongControlInteraction,
        PopupCorruption::DeprecatedReferenceType,
        PopupCorruption::DeprecatedExternalReference,
        PopupCorruption::ZeroControlRefcount,
        PopupCorruption::DuplicateControlKey,
        PopupCorruption::SegmentedControlList,
        PopupCorruption::DuplicateControlListMessage,
        PopupCorruption::ControlAggregateMissing,
        PopupCorruption::ControlFieldInfoMissing,
        PopupCorruption::ControlAggregateWrong,
        PopupCorruption::ControlFieldInfoWrong,
    ] {
        let hostile = with_corruption(&source, corruption)?;
        assert_popup_read_rejects(&hostile, &format!("{corruption:?}"))?;
        assert_changed_edit_rejects_atomically(&hostile, &format!("{corruption:?}"))?;
    }
    Ok(())
}

#[test]
fn popup_old_model_metadata_ownership_is_strict_and_atomic() -> TestResult {
    let source = fixture(FixtureMode::SinglePopup)?;
    for corruption in [
        MetadataCorruption::MissingPopupUuid,
        MetadataCorruption::VersionedOnlyPopupUuid,
        MetadataCorruption::DuplicatePopupUuid,
        MetadataCorruption::AmbiguousPopupIdentifier,
        MetadataCorruption::DataOwnerPopupIdentifier,
        MetadataCorruption::RootDataMapPopupIdentifier,
    ] {
        let hostile = with_metadata_corruption(&source, corruption)?;
        assert_reset_rejects_atomically(&hostile, &format!("{corruption:?}"))?;
    }
    Ok(())
}

#[test]
fn popup_cross_component_physical_alias_is_rejected_atomically() -> TestResult {
    let source = fixture(FixtureMode::SharedPopup)?;
    let aliased = with_physical_popup_alias(&source)?;
    assert_changed_edit_rejects_atomically(&aliased, "cross-component physical popup alias")?;
    Ok(())
}

#[test]
fn popup_missing_metadata_and_tight_ingress_fail_before_publication() -> TestResult {
    let source = fixture(FixtureMode::EmptyCells)?;
    let without_metadata = Catalog::from_bytes(&source)?.reassemble_with_deletions_to_bytes(
        &[],
        &[METADATA_MEMBER],
        Limits::default(),
    )?;
    let package = Package::from_bytes(&without_metadata)?;
    let before = package.exact_bytes();
    let menu = popup(&["One", "Two"], InitialSelection::Blank)?;
    assert!(
        package
            .edit_table_cell_pop_up_menu_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(0, 0),
            )?
            .set(menu)
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
fn popup_unknown_metadata_and_payload_bytes_survive_rewrite() -> TestResult {
    let source = fixture(FixtureMode::SharedPopup)?;
    let hostile = append_unknown_popup_field(&source, 90, 0x1234)?;
    let package = Package::from_bytes(&hostile)?;
    let menu = popup(&["Next", "State"], InitialSelection::Blank)?;
    let commit = package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(menu)
        .commit()?;
    let target = commit.package().exact_bytes();
    let target_payload = object_messages(&target, DOCUMENT_MEMBER, POPUP_MODEL_ID)?
        .into_iter()
        .find(|(type_, _)| *type_ == POPUP_MODEL_TYPE)
        .ok_or_else(|| io::Error::other("target popup payload is missing"))?
        .1;
    assert!(
        !WireView::parse(&target_payload)?
            .fields()
            .filter(|field| field.number() == 90)
            .collect::<Vec<_>>()
            .is_empty()
    );
    assert!(
        !WireView::parse(&metadata_payload(&target)?)?
            .fields()
            .filter(|field| field.number() == 90)
            .collect::<Vec<_>>()
            .is_empty()
    );
    Ok(())
}

#[test]
fn popup_unknown_overlong_and_group_bytes_are_preserved_when_ingress_admits_them() -> TestResult {
    let source = fixture(FixtureMode::SharedPopup)?;
    let (hostile, raw) = append_unknown_popup_wire(&source)?;
    let package = match Package::from_bytes(&hostile) {
        Err(_) => return Ok(()),
        Ok(package) => package,
    };
    let before = package.exact_bytes();
    let result = package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(popup(&["Opaque", "Future"], InitialSelection::Blank)?)
        .commit();
    let Ok(commit) = result else {
        assert_eq!(package.exact_bytes(), before);
        return Ok(());
    };
    let target = commit.package().exact_bytes();
    let payloads = object_messages(&target, DOCUMENT_MEMBER, POPUP_MODEL_ID)?;
    assert!(
        payloads
            .iter()
            .any(|(_, payload)| payload.windows(raw.len()).any(|window| window == raw),),
        "admitted opaque popup framing was not retained"
    );
    Ok(())
}

#[test]
fn popup_public_physical_and_semantic_limits_fail_before_publication() -> TestResult {
    let source_with_previews = fixture(FixtureMode::SharedPopup)?;
    let source = Catalog::from_bytes(&source_with_previews)?.reassemble_with_deletions_to_bytes(
        &[],
        &["preview.jpg", "preview-micro.jpg", "preview-web.jpg"],
        Limits::default(),
    )?;
    // The replacement keeps the existing graph but grows the selected popup
    // payload beyond the source, allowing exact output-limit replay without
    // exceeding the codec's source-derived item budget.
    let menu = popup(
        &[
            "A deliberately longer choice 000",
            "A deliberately longer choice 001",
            "A deliberately longer choice 002",
        ],
        InitialSelection::FirstItem,
    )?;

    // Physical input/stream ceilings reject the source before a Package can
    // expose an editable snapshot.
    let source_limits = Limits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        Limits::MAX_IWA_STREAM_BYTES,
    )?;
    assert!(
        Package::from_bytes_with_options(
            &source,
            PackageReadOptions::new(source_limits, PackageSemanticLimits::default()),
        )
        .is_err()
    );

    // First establish the exact candidate size without limits, then replay
    // that size and one byte below it through the public API.
    let unrestricted = Package::from_bytes(&source)?;
    let target = unrestricted
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(menu.clone())
        .commit()?;
    let target_bytes = target.package().exact_bytes();
    assert!(target_bytes.len() > source.len());
    let exact_output_limits = Limits::new(
        u64::try_from(target_bytes.len())?,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        4 * 1024,
    )?;
    let exact_package = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(exact_output_limits, PackageSemanticLimits::default()),
    )?;
    let exact_commit = exact_package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(menu.clone())
        .commit()?;
    assert_eq!(exact_commit.package().exact_bytes(), target_bytes);

    let limited_output_limits = Limits::new(
        u64::try_from(target_bytes.len().saturating_sub(1))?,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        4 * 1024,
    )?;
    let output_limited = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(limited_output_limits, PackageSemanticLimits::default()),
    )?;
    let before = output_limited.exact_bytes();
    let output_edit = output_limited.edit_table_cell_pop_up_menu_format(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellPosition::new(0, 0),
    )?;
    assert!(output_edit.set(menu.clone()).commit().is_err());
    assert_eq!(output_limited.exact_bytes(), before);

    // A deliberately small native stream ceiling is independently typed by
    // the public archive options.  If ingress accepts it as an opaque source,
    // the changed transaction must still leave its exact snapshot untouched.
    let stream_limits = Limits::new(
        u64::try_from(source.len())?,
        Limits::MAX_ENTRIES,
        Limits::MAX_ENTRY_BYTES,
        Limits::MAX_TOTAL_BYTES,
        1,
    )?;
    if let Ok(package) = Package::from_bytes_with_options(
        &source,
        PackageReadOptions::new(stream_limits, PackageSemanticLimits::default()),
    ) {
        let before = package.exact_bytes();
        let edit = package.edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        );
        if let Ok(edit) = edit {
            assert!(edit.set(menu).commit().is_err());
            assert_eq!(package.exact_bytes(), before);
        }
    }

    // Public semantic object/reference ceilings are ingress-level controls;
    // there are no public codec fields/work knobs on PackageReadOptions.
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
fn popup_refcount_undercount_with_sibling_bnc_is_rejected_atomically() -> TestResult {
    for (undercount_format, undercount_control, label) in [
        (true, false, "format refcount undercount"),
        (false, true, "control refcount undercount"),
        (true, true, "format/control refcount undercount"),
    ] {
        let source = with_refcount_undercount(
            &fixture(FixtureMode::SharedPopup)?,
            undercount_format,
            undercount_control,
        )?;
        assert_reset_rejects_atomically(&source, label)?;
    }
    Ok(())
}

#[test]
fn popup_orphan_matching_model_missing_or_versioned_uuid_is_not_reused() -> TestResult {
    for versioned in [false, true] {
        let source = with_orphan_popup_model(&fixture(FixtureMode::EmptyCells)?, versioned)?;
        let package = match Package::from_bytes(&source) {
            Err(_) => continue,
            Ok(package) => package,
        };
        let menu = popup(&["Low", "Medium", "High"], InitialSelection::FirstItem)?;
        let result = package
            .edit_table_cell_pop_up_menu_format(
                SheetSelector::index(0),
                TableSelector::index(0),
                CellPosition::new(0, 0),
            )?
            .set(menu)
            .commit();
        let Ok(commit) = result else {
            continue;
        };
        let ids = popup_model_ids(&commit.package().exact_bytes())?;
        let new_ids = ids
            .iter()
            .filter(|identifier| **identifier != POPUP_MODEL_ID)
            .copied()
            .collect::<Vec<_>>();
        assert!(
            !new_ids.is_empty(),
            "the orphan popup identifier was reused instead of allocating a fresh model"
        );
        let document = document_metadata_component(&commit.package().exact_bytes(), 100)?;
        assert!(
            document
                .object_uuid_map_entries
                .iter()
                .all(|entry| entry.identifier != POPUP_MODEL_ID),
            "the orphan popup must not be registered as a current owner"
        );
        assert!(
            new_ids.iter().all(|identifier| {
                document
                    .object_uuid_map_entries
                    .iter()
                    .any(|entry| entry.identifier == *identifier)
            }),
            "fresh popup models must be registered in the current Document UUID map"
        );
    }
    Ok(())
}

#[test]
fn popup_cross_component_or_opaque_inbound_prevents_final_cull() -> TestResult {
    let source = with_cross_component_popup_inbound(&fixture(FixtureMode::SinglePopup)?)?;
    assert_reset_rejects_atomically(&source, "cross-component opaque popup inbound")?;
    Ok(())
}

#[test]
fn popup_existing_graph_component_uuid_ownership_is_strict_and_atomic() -> TestResult {
    let source = fixture(FixtureMode::SharedPopup)?;
    for identifier in [TABLE_MODEL_ID, TILE_ID, SIDECAR_ID] {
        for versioned in [false, true] {
            let hostile = with_component_uuid_mutation(&source, identifier, versioned)?;
            assert_changed_edit_rejects_atomically(
                &hostile,
                &format!("component {identifier} UUID versioned={versioned}"),
            )?;
        }
    }
    Ok(())
}

#[test]
fn popup_control_metadata_requires_exact_aggregate_and_field_entries() -> TestResult {
    let source = fixture(FixtureMode::SharedPopup)?;
    for (hostile, label) in [
        (
            with_extra_control_aggregate(&source)?,
            "extra control aggregate reference",
        ),
        (
            with_duplicate_control_field_info(&source)?,
            "duplicate control FieldInfo",
        ),
    ] {
        assert_popup_read_rejects(&hostile, label)?;
        assert_changed_edit_rejects_atomically(&hostile, label)?;
    }
    Ok(())
}

#[test]
fn popup_effective_locator_uses_explicit_locator_and_inverse() -> TestResult {
    let source = with_explicit_document_locator(&fixture(FixtureMode::SharedPopup)?)?;
    let package = Package::from_bytes(&source)?;
    let replacement = popup(&["Explicit", "Locator"], InitialSelection::Blank)?;
    let commit = package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(replacement.clone())
        .commit()?;
    assert_eq!(
        commit.package().table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?,
        Some(replacement)
    );
    assert_eq!(
        Package::from_bytes(&commit.package().exact_bytes())?
            .apply_table_cell_pop_up_menu_format(&commit.patch().inverse())?
            .package()
            .exact_bytes(),
        source
    );
    Ok(())
}

#[test]
fn popup_dangling_format_route_fails_closed_atomically() -> TestResult {
    let source = with_dangling_format_route(&fixture(FixtureMode::SharedPopup)?)?;
    assert_popup_read_rejects(&source, "dangling format route")?;
    assert_changed_edit_rejects_atomically(&source, "dangling format route")?;
    Ok(())
}

#[test]
fn popup_replacement_preserves_unselected_document_objects_exactly() -> TestResult {
    let source = fixture(FixtureMode::SharedPopup)?;
    let package = Package::from_bytes(&source)?;
    let replacement = popup(&["Locality", "Only"], InitialSelection::Blank)?;
    let target = package
        .edit_table_cell_pop_up_menu_format(
            SheetSelector::index(0),
            TableSelector::index(0),
            CellPosition::new(0, 0),
        )?
        .set(replacement)
        .commit()?
        .package()
        .exact_bytes();
    assert_unchanged_document_objects(
        &source,
        &target,
        &[DOCUMENT_ID, SHEET_ID, TABLE_INFO_ID, TABLE_MODEL_ID],
    )?;
    assert_eq!(
        object_messages(&source, VIEW_STATE_MEMBER, VIEW_STATE_ID)?,
        object_messages(&target, VIEW_STATE_MEMBER, VIEW_STATE_ID)?,
    );
    Ok(())
}

fn popup_model_count(source: &[u8]) -> TestResult<usize> {
    Ok(member_archive(source, DOCUMENT_MEMBER)?
        .objects
        .iter()
        .filter(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == POPUP_MODEL_TYPE)
        })
        .count())
}

fn popup_model_ids(source: &[u8]) -> TestResult<Vec<u64>> {
    Ok(member_archive(source, DOCUMENT_MEMBER)?
        .objects
        .iter()
        .filter(|object| {
            object
                .messages
                .iter()
                .any(|message| message.type_ == POPUP_MODEL_TYPE)
        })
        .filter_map(|object| object.archive_info.identifier)
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
    Err(io::Error::other("control-cell-spec list is missing").into())
}
