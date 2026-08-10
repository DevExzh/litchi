use std::io;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::append_length_delimited_field;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsd, tsp, tst};
use prost::Message as _;

pub(crate) const TABLES_MEMBER: &str = "Index/Tables.iwa";
pub(crate) const DOCUMENT: u64 = 1;
pub(crate) const SHEET: u64 = 2;
pub(crate) const TABLE_INFO: u64 = 10;
pub(crate) const TABLE_MODEL: u64 = 20;
pub(crate) const CATEGORY_OWNER: u64 = 30;
pub(crate) const GROUP_BY: u64 = 40;
pub(crate) const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
pub(crate) const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
pub(crate) const CATEGORY_OWNER_REFERENCE_MESSAGE_TYPE: u32 = 6_372;
pub(crate) const GROUP_BY_MESSAGE_TYPE: u32 = 6_373;

type FixtureResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

pub(crate) fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn object(identifier: u64, message_type: u32, data: Vec<u8>) -> FixtureResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
    )?)
}

fn table_model() -> tst::TableModelArchive {
    tst::TableModelArchive {
        table_name: "Focused headers".to_owned(),
        number_of_rows: 1,
        number_of_columns: 1,
        base_data_store: tst::DataStore {
            string_table: reference(90),
            formula_table: reference(90),
            ..Default::default()
        },
        ..Default::default()
    }
}

fn sidecars() -> FixtureResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        90,
        [
            tst::table_data_list::ListType::String,
            tst::table_data_list::ListType::Formula,
        ]
        .into_iter()
        .map(|list_type| RawMessage {
            type_: 6_005,
            data: tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 1,
                ..Default::default()
            }
            .encode_to_vec(),
        })
        .collect(),
    )?)
}

pub(crate) fn synthetic_package() -> FixtureResult<Vec<u8>> {
    let mut document = object(
        DOCUMENT,
        1,
        tn::DocumentArchive {
            sheets: vec![reference(SHEET)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    document.archive_info.message_infos[0].object_references = vec![SHEET];
    let mut sheet = object(
        SHEET,
        2,
        tn::SheetArchive {
            name: "Synthetic".to_owned(),
            drawable_infos: vec![reference(TABLE_INFO)],
            ..Default::default()
        }
        .encode_to_vec(),
    )?;
    sheet.archive_info.message_infos[0].object_references = vec![TABLE_INFO];

    let document_component = SnappyStream::compress(
        &Archive {
            objects: vec![document, sheet],
        }
        .to_bytes()?,
    )?;

    let mut table_info_payload = Vec::new();
    append_length_delimited_field(
        &mut table_info_payload,
        1,
        &tsd::DrawableArchive::default().encode_to_vec(),
    )?;
    append_length_delimited_field(
        &mut table_info_payload,
        2,
        &reference(TABLE_MODEL).encode_to_vec(),
    )?;
    let mut table_info = object(TABLE_INFO, TABLE_INFO_MESSAGE_TYPE, table_info_payload)?;
    table_info.archive_info.message_infos[0].object_references = vec![TABLE_MODEL];
    let tables_component = SnappyStream::compress(
        &Archive {
            objects: vec![
                table_info,
                object(
                    TABLE_MODEL,
                    TABLE_MODEL_MESSAGE_TYPE,
                    table_model().encode_to_vec(),
                )?,
                sidecars()?,
            ],
        }
        .to_bytes()?,
    )?;

    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("preview.jpg", b"full preview sentinel".as_slice()),
            ("preview-micro.jpg", b"micro preview sentinel".as_slice()),
            ("preview-web.jpg", b"web preview sentinel".as_slice()),
            ("Index/Document.iwa", document_component.as_slice()),
            (TABLES_MEMBER, tables_component.as_slice()),
            (
                "Data/sentinel.bin",
                b"unrelated package sentinel".as_slice(),
            ),
        ],
        Limits::default(),
    )?)
}

pub(crate) fn rewrite_tables(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> FixtureResult,
) -> FixtureResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == TABLES_MEMBER)
        .ok_or_else(|| io::Error::other("synthetic table component is missing"))?;
    let stream = SnappyStream::decompress(entry.data())?;
    let mut archive = Archive::parse(stream.as_bytes())?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(TABLES_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

pub(crate) fn append_reference_field(
    object: &mut ArchiveObject,
    message_type: u32,
    field_number: u32,
    identifier: u64,
    declare: bool,
) -> FixtureResult {
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == message_type)
        .ok_or_else(|| io::Error::other("synthetic selected message is missing"))?;
    let mut data = object.messages[message_index].data.clone();
    append_length_delimited_field(
        &mut data,
        field_number,
        &reference(identifier).encode_to_vec(),
    )?;
    object.replace_message_preserving_header(
        message_index,
        RawMessage {
            type_: message_type,
            data,
        },
    )?;
    if declare {
        object.archive_info.message_infos[message_index]
            .object_references
            .push(identifier);
    }
    Ok(())
}

pub(crate) fn append_selected_model_raw(package: &[u8], raw: &[u8]) -> FixtureResult<Vec<u8>> {
    rewrite_tables(package, |archive| {
        let model = archive
            .object_mut(TABLE_MODEL)
            .ok_or_else(|| io::Error::other("synthetic table model is missing"))?;
        let index = model
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("synthetic table-model message is missing"))?;
        let mut data = model.messages[index].data.clone();
        data.extend_from_slice(raw);
        model.replace_message_preserving_header(
            index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

pub(crate) fn category_owner_object(references: &[u64]) -> FixtureResult<ArchiveObject> {
    let mut object = object(
        CATEGORY_OWNER,
        CATEGORY_OWNER_REFERENCE_MESSAGE_TYPE,
        tst::CategoryOwnerRefArchive {
            group_by: references.iter().copied().map(reference).collect(),
        }
        .encode_to_vec(),
    )?;
    object.archive_info.message_infos[0].object_references = references.to_vec();
    Ok(object)
}

pub(crate) fn enabled_group_object() -> FixtureResult<ArchiveObject> {
    object(
        GROUP_BY,
        GROUP_BY_MESSAGE_TYPE,
        tst::GroupByArchive {
            group_by_uid: tsp::Uuid { lower: 1, upper: 2 },
            is_enabled: true,
            ..Default::default()
        }
        .encode_to_vec(),
    )
}

pub(crate) fn push_object(archive: &mut Archive, object: ArchiveObject) {
    archive.objects.push(object);
}

pub(crate) fn transaction_work_precharge(source: &[u8], target: &[u8]) -> FixtureResult<usize> {
    let catalog = Catalog::from_bytes(source)?;
    let mut observed = source.len().saturating_mul(2);
    if source != target {
        observed = observed.saturating_add(target.len().saturating_mul(2));
    }
    for entry in catalog
        .iter()
        .filter(|entry| entry.name().ends_with(".iwa"))
    {
        let stream = SnappyStream::decompress(entry.data())?;
        let archive = Archive::parse(stream.as_bytes())?;
        for object in &archive.objects {
            for message in &object.messages {
                observed = observed.saturating_add(message.data.len().saturating_mul(16));
            }
        }
    }
    Ok(observed)
}
