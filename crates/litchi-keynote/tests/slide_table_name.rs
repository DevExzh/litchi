//! Exact-source integration coverage for Keynote slide-table names.
//!
//! The fixture intentionally follows the dimension owner: slide topology and
//! table-info objects live in Document.iwa, complete table models and storage
//! live in CalculationEngine.iwa, and package authority lives in Metadata.iwa.
//! The tests exercise the field-8 name projection without exposing native
//! identifiers through the public transaction.

use std::error::Error as StdError;
use std::io;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::{WireView, append_length_delimited_field, append_varint_field};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp, tst, tswp};
use litchi_keynote::{
    Package, ReadOptions, SemanticLimits, SlideSelector, SlideTableName, SlideTableNameCommit,
    SlideTableNameDiagnostics, SlideTableNameEdit, SlideTableNameError, SlideTableNameLimitKind,
    SlideTableNamePatch, SlideTableNamePath, SlideTableNameValueError, TableSelector,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const MODEL_MEMBER: &str = "Index/CalculationEngine.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const SENTINEL_MEMBER: &str = "Data/name-sentinel.bin";
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const TABLE_INFOS: [u64; 2] = [100, 101];
const MODELS: [u64; 2] = [110, 111];
const NON_TABLE_DRAWABLE: u64 = 130;
const METADATA_OBJECT: u64 = 900;
const FOREIGN_INBOUND_OBJECT: u64 = 902;
const FOREIGN_BUCKET_INBOUND_OBJECT: u64 = 903;

const ROW_BUCKETS: [u64; 2] = [200, 210];
const COLUMN_BUCKETS: [u64; 2] = [201, 211];
const DATA_OBJECTS: [[u64; 4]; 2] = [[220, 221, 222, 223], [230, 231, 232, 233]];

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHOW_MESSAGE_TYPE: u32 = 2;
const SLIDE_NODE_MESSAGE_TYPE: u32 = 4;
const SLIDE_MESSAGE_TYPE: u32 = 5;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const LEGACY_TABLE_MODEL_MESSAGE_TYPE: u32 = 6_000;
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const METADATA_MESSAGE_TYPE: u32 = 11_006;

const SLIDE_OWNED_DRAWABLES_FIELD: u32 = 7;
const SLIDE_Z_ORDER_FIELD: u32 = 42;
const TABLE_MODEL_FIELD: u32 = 2;
const MODEL_STORAGE_FIELD: u32 = 4;
const UNKNOWN_MODEL_FIELD: u32 = 90;
const UNKNOWN_MODEL_VALUE: u64 = 0xfeed_beef;
const UNKNOWN_BUCKET_FIELD: u32 = 98;
const UNKNOWN_BUCKET_VALUE: u64 = 0xcafe_babe;
const UNKNOWN_HEADER_FIELD: u32 = 97;
const UNKNOWN_HEADER_VALUE: u64 = 0xdead_beef;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn field_reference(path: impl Into<FieldPath>, references: &[u64]) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.object_references.extend_from_slice(references);
    field
}

fn object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    references: &[u64],
) -> TestResult<ArchiveObject> {
    let mut object = ArchiveObject::new(identifier, vec![RawMessage { type_, data }])?;
    object.archive_info.message_infos[0]
        .object_references
        .extend_from_slice(references);
    Ok(object)
}

fn storage_ids(index: usize) -> [u64; 6] {
    [
        ROW_BUCKETS[index],
        COLUMN_BUCKETS[index],
        DATA_OBJECTS[index][0],
        DATA_OBJECTS[index][1],
        DATA_OBJECTS[index][2],
        DATA_OBJECTS[index][3],
    ]
}

fn data_store(index: usize) -> tst::DataStore {
    let ids = storage_ids(index);
    tst::DataStore {
        row_headers: tst::HeaderStorage {
            bucket_hash_function: 1,
            buckets: vec![reference(ids[0])],
        },
        column_headers: reference(ids[1]),
        tiles: tst::TileStorage::default(),
        string_table: reference(ids[2]),
        style_table: reference(ids[3]),
        formula_table: reference(ids[4]),
        format_table_pre_bnc: reference(ids[5]),
        next_row_strip_id: 1,
        next_column_strip_id: 1,
        row_tile_tree: tst::TableRbTree::default(),
        column_tile_tree: tst::TableRbTree::default(),
        ..tst::DataStore::default()
    }
}

fn synthetic_model_payload(index: usize, name: &str, with_unknown: bool) -> TestResult<Vec<u8>> {
    let model = tst::TableModelArchive {
        table_id: format!("table-{name}"),
        table_style: reference(90),
        body_text_style: reference(90),
        header_row_text_style: reference(90),
        header_column_text_style: reference(90),
        footer_row_text_style: reference(90),
        body_cell_style: reference(90),
        header_row_style: reference(90),
        header_column_style: reference(90),
        footer_row_style: reference(90),
        base_data_store: data_store(index),
        number_of_rows: 8,
        number_of_columns: 4,
        table_name: name.to_owned(),
        default_row_height: 20.0,
        default_column_width: 64.0,
        ..tst::TableModelArchive::default()
    };
    let mut payload = model.encode_to_vec();
    if with_unknown {
        append_varint_field(&mut payload, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)?;
    }
    Ok(payload)
}

fn shape_info_payload() -> Vec<u8> {
    tswp::ShapeInfoArchive {
        super_: tsd::ShapeArchive {
            super_: tsd::DrawableArchive::default(),
            ..tsd::ShapeArchive::default()
        },
        ..tswp::ShapeInfoArchive::default()
    }
    .encode_to_vec()
}

fn component_payload(
    identifier: u64,
    locator: &str,
    identifiers: impl IntoIterator<Item = u64>,
) -> Vec<u8> {
    tsp::ComponentInfo {
        identifier,
        preferred_locator: locator.to_owned(),
        locator: Some(locator.to_owned()),
        save_token: Some(1),
        object_uuid_map_entries: identifiers
            .into_iter()
            .map(|identifier| tsp::ObjectUuidMapEntry {
                identifier,
                uuid: tsp::Uuid {
                    lower: identifier + 10_000,
                    upper: identifier + 20_000,
                },
            })
            .collect(),
        ..tsp::ComponentInfo::default()
    }
    .encode_to_vec()
}

fn metadata_payload() -> TestResult<Vec<u8>> {
    let document_component = component_payload(
        1,
        "Document",
        [
            1,
            2,
            SLIDE_NODE,
            SLIDE,
            TABLE_INFOS[0],
            TABLE_INFOS[1],
            NON_TABLE_DRAWABLE,
        ],
    );
    let model_component = component_payload(
        2,
        "CalculationEngine",
        MODELS
            .into_iter()
            .chain(ROW_BUCKETS)
            .chain(COLUMN_BUCKETS)
            .chain(DATA_OBJECTS.into_iter().flatten()),
    );
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, METADATA_OBJECT)?;
    append_length_delimited_field(&mut payload, 3, &document_component)?;
    append_length_delimited_field(&mut payload, 3, &model_component)?;
    Ok(payload)
}

fn bucket_header(index: u32, size: f32, with_unknown: bool) -> TestResult<Vec<u8>> {
    let mut payload = tst::header_storage_bucket::Header {
        index,
        size,
        hiding_state: 0,
        number_of_cells: 0,
        cell_style: None,
        text_style: None,
    }
    .encode_to_vec();
    if with_unknown {
        append_varint_field(&mut payload, UNKNOWN_HEADER_FIELD, UNKNOWN_HEADER_VALUE)?;
    }
    Ok(payload)
}

fn bucket_payload(headers: &[(u32, f32)], with_unknowns: bool) -> TestResult<Vec<u8>> {
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, 1)?;
    for &(index, size) in headers {
        let header = bucket_header(index, size, with_unknowns)?;
        append_length_delimited_field(&mut payload, 2, &header)?;
    }
    if with_unknowns {
        append_varint_field(&mut payload, UNKNOWN_BUCKET_FIELD, UNKNOWN_BUCKET_VALUE)?;
    }
    Ok(payload)
}

fn synthetic_package(locked: bool, with_unknowns: bool) -> TestResult<Vec<u8>> {
    let table_drawables = [TABLE_INFOS[0], NON_TABLE_DRAWABLE, TABLE_INFOS[1]];
    let document = kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(2),
        ..kn::DocumentArchive::default()
    };
    let show = kn::ShowArchive {
        theme: reference(80),
        slide_tree: kn::SlideTreeArchive {
            slides: vec![reference(SLIDE_NODE)],
            ..kn::SlideTreeArchive::default()
        },
        size: tsp::Size {
            width: 1_024.0,
            height: 768.0,
        },
        stylesheet: reference(81),
        ..kn::ShowArchive::default()
    };
    #[allow(deprecated, reason = "native schema retains cache fields")]
    let node = kn::SlideNodeArchive {
        slide: Some(reference(SLIDE)),
        is_skipped: false,
        has_builds: false,
        has_transition: false,
        ..kn::SlideNodeArchive::default()
    };
    let slide = kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: table_drawables.iter().copied().map(reference).collect(),
        drawables_z_order: table_drawables.iter().copied().map(reference).collect(),
        name: Some("Tables".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    };
    let mut slide_object = object(
        SLIDE,
        SLIDE_MESSAGE_TYPE,
        slide.encode_to_vec(),
        &table_drawables,
    )?;
    slide_object.archive_info.message_infos[0]
        .field_infos
        .extend([
            field_reference(vec![SLIDE_OWNED_DRAWABLES_FIELD], &table_drawables),
            field_reference(vec![SLIDE_Z_ORDER_FIELD], &table_drawables),
        ]);

    let mut document_objects = vec![
        object(1, DOCUMENT_MESSAGE_TYPE, document.encode_to_vec(), &[2])?,
        object(
            2,
            SHOW_MESSAGE_TYPE,
            show.encode_to_vec(),
            &[SLIDE_NODE, 80, 81],
        )?,
        object(
            SLIDE_NODE,
            SLIDE_NODE_MESSAGE_TYPE,
            node.encode_to_vec(),
            &[SLIDE],
        )?,
        slide_object,
        object(
            NON_TABLE_DRAWABLE,
            SHAPE_INFO_MESSAGE_TYPE,
            shape_info_payload(),
            &[SLIDE],
        )?,
    ];

    let mut model_objects = Vec::new();
    for (index, table_info_identifier) in TABLE_INFOS.into_iter().enumerate() {
        let (width, height) = if index == 0 {
            (232.0, 165.0)
        } else {
            (256.0, 160.0)
        };
        let info = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                geometry: Some(tsd::GeometryArchive {
                    size: Some(tsp::Size { width, height }),
                    ..tsd::GeometryArchive::default()
                }),
                parent: Some(reference(SLIDE)),
                locked: Some(locked && index == 0),
                ..tsd::DrawableArchive::default()
            },
            table_model: reference(MODELS[index]),
            ..tst::TableInfoArchive::default()
        };
        let mut info_object = object(
            table_info_identifier,
            TABLE_INFO_MESSAGE_TYPE,
            info.encode_to_vec(),
            &[MODELS[index]],
        )?;
        info_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![TABLE_MODEL_FIELD], &[MODELS[index]]));
        document_objects.push(info_object);

        let ids = storage_ids(index);
        let mut model_object = object(
            MODELS[index],
            TABLE_MODEL_MESSAGE_TYPE,
            synthetic_model_payload(
                index,
                if index == 0 { "Revenue" } else { "Costs" },
                with_unknowns,
            )?,
            &ids,
        )?;
        model_object.archive_info.message_infos[0]
            .field_infos
            .extend([
                field_reference(vec![MODEL_STORAGE_FIELD, 1, 2], &[ids[0]]),
                field_reference(vec![MODEL_STORAGE_FIELD, 2], &[ids[1]]),
                field_reference(vec![MODEL_STORAGE_FIELD, 4], &[ids[2]]),
                field_reference(vec![MODEL_STORAGE_FIELD, 5], &[ids[3]]),
                field_reference(vec![MODEL_STORAGE_FIELD, 6], &[ids[4]]),
                field_reference(vec![MODEL_STORAGE_FIELD, 11], &[ids[5]]),
            ]);
        model_objects.push(model_object);
        model_objects.push(object(
            ROW_BUCKETS[index],
            HEADER_BUCKET_MESSAGE_TYPE,
            bucket_payload(if index == 0 { &[(1, 25.0)] } else { &[] }, with_unknowns)?,
            &[],
        )?);
        model_objects.push(object(
            COLUMN_BUCKETS[index],
            HEADER_BUCKET_MESSAGE_TYPE,
            bucket_payload(if index == 0 { &[(2, 40.0)] } else { &[] }, with_unknowns)?,
            &[],
        )?);
        for identifier in DATA_OBJECTS[index] {
            model_objects.push(object(identifier, 7_000, vec![0x08, 0x01], &[])?);
        }
    }

    let document_component = SnappyStream::compress(
        &Archive {
            objects: document_objects,
        }
        .to_bytes()?,
    )?;
    let model_component = SnappyStream::compress(
        &Archive {
            objects: model_objects,
        }
        .to_bytes()?,
    )?;
    let metadata_component = SnappyStream::compress(
        &Archive {
            objects: vec![object(
                METADATA_OBJECT,
                METADATA_MESSAGE_TYPE,
                metadata_payload()?,
                &[],
            )?],
        }
        .to_bytes()?,
    )?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            (SENTINEL_MEMBER, b"untouched-name-sentinel".as_slice()),
            (DOCUMENT_MEMBER, document_component.as_slice()),
            (MODEL_MEMBER, model_component.as_slice()),
            (METADATA_MEMBER, metadata_component.as_slice()),
            (PREVIEWS[0], b"large preview".as_slice()),
            (PREVIEWS[1], b"micro preview".as_slice()),
            (PREVIEWS[2], b"web preview".as_slice()),
        ],
        Limits::default(),
    )?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn member_bytes(package: &[u8], member: &str) -> TestResult<Vec<u8>> {
    Ok(Catalog::from_bytes(package)?
        .iter()
        .find(|entry| entry.name() == member)
        .ok_or_else(|| io::Error::other(format!("missing member {member}")))?
        .data()
        .to_vec())
}

fn has_member(package: &[u8], member: &str) -> TestResult<bool> {
    Ok(Catalog::from_bytes(package)?
        .iter()
        .any(|entry| entry.name() == member))
}

fn archive(package: &[u8], member: &str) -> TestResult<Archive> {
    Ok(Archive::parse(
        SnappyStream::decompress(&member_bytes(package, member)?)?.as_bytes(),
    )?)
}

fn rewrite_member_archive(
    package: &[u8],
    member: &str,
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let mut selected = archive(package, member)?;
    mutate(&mut selected)?;
    let replacement = SnappyStream::compress(&selected.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(&[EntryEdit::new(member, &replacement)], Limits::default())?)
}

fn rewrite_document_archive(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_member_archive(package, DOCUMENT_MEMBER, mutate)
}

fn rewrite_model_archive(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_member_archive(package, MODEL_MEMBER, mutate)
}

fn model_member_payload(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    object_message_payload(package, MODEL_MEMBER, identifier, TABLE_MODEL_MESSAGE_TYPE)
}

fn object_message_payload(
    package: &[u8],
    member: &str,
    identifier: u64,
    type_: u32,
) -> TestResult<Vec<u8>> {
    let archive = archive(package, member)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing archive object"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == type_)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing archive message").into())
}

fn replace_model_payload(package: &[u8], identifier: u64, payload: Vec<u8>) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("missing model object"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing model message"))?;
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })
}

fn append_selected_model_raw(package: &[u8], raw: &[u8]) -> TestResult<Vec<u8>> {
    let mut payload = model_member_payload(package, MODELS[0])?;
    payload.extend_from_slice(raw);
    replace_model_payload(package, MODELS[0], payload)
}

fn rewrite_model_storage<F>(package: &[u8], rewrite: F) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut tst::DataStore),
{
    let payload = model_member_payload(package, MODELS[0])?;
    let mut model = tst::TableModelArchive::decode(payload.as_slice())?;
    rewrite(&mut model.base_data_store);
    // replace_model_payload deliberately preserves the model ArchiveInfo
    // header, so these mutations only change the wire-level DataStore route.
    replace_model_payload(package, MODELS[0], model.encode_to_vec())
}

fn append_duplicate_storage_aggregate(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(MODELS[0])
            .ok_or_else(|| io::Error::other("missing model object"))?;
        let info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("missing model metadata"))?;
        info.object_references.push(ROW_BUCKETS[0]);
        Ok(())
    })
}

fn append_duplicate_storage_field(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(MODELS[0])
            .ok_or_else(|| io::Error::other("missing model object"))?;
        let info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("missing model metadata"))?;
        let field = info
            .field_infos
            .iter()
            .find(|field| field.path.as_slice() == [MODEL_STORAGE_FIELD, 1, 2])
            .cloned()
            .ok_or_else(|| io::Error::other("missing row storage field"))?;
        info.field_infos.push(field);
        Ok(())
    })
}

fn append_noncanonical_storage_field(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(MODELS[0])
            .ok_or_else(|| io::Error::other("missing model object"))?;
        let info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("missing model metadata"))?;
        info.field_infos.push(field_reference(
            vec![MODEL_STORAGE_FIELD, 99],
            &[ROW_BUCKETS[0]],
        ));
        Ok(())
    })
}

fn append_storage_role_alias(
    package: &[u8],
    identifier: u64,
    alias_type: u32,
) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("missing storage bucket"))?;
        let message_index = (!object.messages.is_empty())
            .then_some(0)
            .ok_or_else(|| io::Error::other("missing storage message"))?;
        let mut alias = object.messages[message_index].clone();
        alias.type_ = alias_type;
        let mut alias_info = object.archive_info.message_infos[message_index].clone();
        alias_info.type_ = alias_type;
        object.messages.push(alias);
        object.archive_info.message_infos.push(alias_info);
        Ok(())
    })
}

fn rewrite_metadata<F>(package: &[u8], rewrite: F) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut tsp::PackageMetadata) -> TestResult<()>,
{
    rewrite_member_archive(package, METADATA_MEMBER, |archive| {
        let object = archive
            .object_mut(METADATA_OBJECT)
            .ok_or_else(|| io::Error::other("missing metadata object"))?;
        let message = object
            .messages
            .iter_mut()
            .find(|message| message.type_ == METADATA_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing metadata message"))?;
        let mut metadata = tsp::PackageMetadata::decode(message.data.as_slice())?;
        rewrite(&mut metadata)?;
        message.data = metadata.encode_to_vec();
        Ok(())
    })
}

fn remove_metadata_uuid(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_metadata(package, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 2)
        {
            component
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != identifier);
        }
        Ok(())
    })
}

fn move_metadata_uuid_to_document(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_metadata(package, |metadata| {
        let entry = metadata
            .components
            .iter()
            .find(|component| component.identifier == 2)
            .and_then(|component| {
                component
                    .object_uuid_map_entries
                    .iter()
                    .find(|entry| entry.identifier == identifier)
                    .cloned()
            })
            .ok_or_else(|| io::Error::other("missing storage UUID entry"))?;
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 2)
        {
            component
                .object_uuid_map_entries
                .retain(|candidate| candidate.identifier != identifier);
        }
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 1)
        {
            component.object_uuid_map_entries.push(entry);
        }
        Ok(())
    })
}

fn unknown_model_group(package: &[u8]) -> TestResult<Vec<u8>> {
    append_selected_model_raw(package, &[0x9b, 0x06, 0x08, 0x01, 0x9c, 0x06])
}

fn duplicate_canonical_model(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(MODELS[0])
            .ok_or_else(|| io::Error::other("missing model object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing model message"))?;
        let message = object.messages[index].clone();
        let info = object.archive_info.message_infos[index].clone();
        object.messages.push(message);
        object.archive_info.message_infos.push(info);
        Ok(())
    })
}

fn model_role_alias(package: &[u8], type_: u32) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(MODELS[0])
            .ok_or_else(|| io::Error::other("missing model object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing model message"))?;
        object.messages[index].type_ = type_;
        object.archive_info.message_infos[index].type_ = type_;
        Ok(())
    })
}

fn table_info_model_role_alias(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or_else(|| io::Error::other("missing table info"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing table info message"))?;
        object.messages[index].type_ = TABLE_MODEL_MESSAGE_TYPE;
        object.archive_info.message_infos[index].type_ = TABLE_MODEL_MESSAGE_TYPE;
        Ok(())
    })
}

fn foreign_inbound(package: &[u8], target: u64, identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        archive
            .objects
            .push(object(identifier, 7_001, vec![0x08, 0x01], &[target])?);
        Ok(())
    })
}

fn missing_model_storage_route(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(MODELS[0])
            .ok_or_else(|| io::Error::other("missing model object"))?;
        object.archive_info.message_infos[0]
            .object_references
            .retain(|value| *value != ROW_BUCKETS[0]);
        Ok(())
    })
}

fn duplicate_slide_identity(package: &[u8]) -> TestResult<Vec<u8>> {
    let slide = archive(package, DOCUMENT_MEMBER)?
        .object(SLIDE)
        .ok_or_else(|| io::Error::other("missing slide object"))?
        .clone();
    rewrite_model_archive(package, |archive| {
        archive.objects.push(slide);
        Ok(())
    })
}

fn append_metadata_unknown(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_member_archive(package, METADATA_MEMBER, |archive| {
        let object = archive
            .object_mut(METADATA_OBJECT)
            .ok_or_else(|| io::Error::other("missing metadata object"))?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == METADATA_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing metadata message"))?;
        let mut payload = object.messages[index].data.clone();
        append_varint_field(&mut payload, 99, 123_456)?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })
}

fn assert_rejected_atomically(source: &[u8]) -> TestResult {
    let package = match Package::from_bytes(source) {
        Ok(package) => package,
        Err(_) => return Ok(()),
    };
    let before = exact_bytes(&package)?;
    let read = package.slide_table_name(SlideSelector::index(0), TableSelector::index(0));
    let edit = package
        .edit_slide_table_name(SlideSelector::index(0), TableSelector::index(0))
        .and_then(|edit| edit.set_name("replacement"))
        .and_then(SlideTableNameEdit::commit);
    assert!(
        matches!(
            read,
            Err(SlideTableNameError::InvalidSource)
                | Err(SlideTableNameError::UnsupportedSource)
                | Err(SlideTableNameError::UnsupportedDependency)
                | Err(SlideTableNameError::UnsupportedTopology)
                | Err(SlideTableNameError::LimitExceeded { .. })
        ),
        "malformed name read was unexpectedly accepted: {read:?}"
    );
    assert!(
        matches!(
            edit,
            Err(SlideTableNameError::InvalidSource)
                | Err(SlideTableNameError::UnsupportedSource)
                | Err(SlideTableNameError::UnsupportedDependency)
                | Err(SlideTableNameError::UnsupportedTopology)
                | Err(SlideTableNameError::LimitExceeded { .. })
        ),
        "malformed name edit was unexpectedly accepted: {edit:?}"
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn typed_names_are_send_sync_debug_and_redacted() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<SlideTableName>();
    assert_send_sync_debug::<SlideTableNameEdit<'static>>();
    assert_send_sync_debug::<SlideTableNamePatch>();
    assert_send_sync_debug::<SlideTableNameCommit>();
    assert_send_sync_debug::<SlideTableNameDiagnostics>();
    assert_send_sync_debug::<SlideTableNameError>();
    assert_send_sync_debug::<SlideTableNameLimitKind>();
    assert_send_sync_debug::<SlideTableNamePath>();

    let package = Package::from_bytes(&synthetic_package(false, false)?)?;
    let edit = package.edit_slide_table_name("Tables", TableSelector::index(0))?;
    let debug = format!("{edit:?}");
    assert!(debug.contains("before"));
    assert!(!debug.contains(DOCUMENT_MEMBER));
    assert!(!debug.contains("110"));
    Ok(())
}

#[test]
fn positional_selectors_read_names_through_interleaved_z_order() -> TestResult {
    let source = synthetic_package(false, false)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package
            .slide_table_name("Tables", TableSelector::index(0))?
            .as_str(),
        "Revenue"
    );
    assert_eq!(
        package
            .slide_table_name(SlideSelector::index(0), TableSelector::index(1))?
            .as_str(),
        "Costs"
    );
    assert!(matches!(
        package.slide_table_name("Missing", 0usize),
        Err(SlideTableNameError::SlideNameNotFound)
    ));
    assert!(matches!(
        package.slide_table_name("Tables", 2usize),
        Err(SlideTableNameError::TablePositionNotFound { .. })
    ));
    assert!(matches!(
        package.slide_table_name("", 0usize),
        Err(SlideTableNameError::EmptySlideName)
    ));
    Ok(())
}

#[test]
fn duplicate_sibling_names_are_allowed_for_positional_selector() -> TestResult {
    let source = synthetic_package(false, false)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_table_name("Tables", 0usize)?
        .set_name("Costs")?
        .commit()?;
    assert_eq!(
        commit
            .package()
            .slide_table_name("Tables", 0usize)?
            .as_str(),
        "Costs"
    );
    assert_eq!(
        commit
            .package()
            .slide_table_name("Tables", 1usize)?
            .as_str(),
        "Costs"
    );
    Ok(())
}

#[test]
fn exact_no_op_is_source_bound_and_apply_is_idempotent() -> TestResult {
    let source = synthetic_package(false, false)?;
    let package = Package::from_bytes(&source)?;
    let before = package.slide_table_name("Tables", 0usize)?;
    let commit = package
        .edit_slide_table_name("Tables", 0usize)?
        .set(before.clone())
        .commit()?;
    assert_eq!(commit.patch().before(), &before);
    assert_eq!(commit.patch().after(), &before);
    assert!(commit.patch().is_noop());
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(!commit.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(commit.package())?, source);
    assert_eq!(
        exact_bytes(package.apply_slide_table_name(commit.patch())?.package())?,
        source
    );
    Ok(())
}

fn non_name_fields(payload: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    Ok(WireView::parse(payload)?
        .fields()
        .filter(|field| field.number() != 8)
        .map(|field| field.raw().to_vec())
        .collect())
}

fn contains_raw(payload: &[u8], raw: &[u8]) -> bool {
    payload.windows(raw.len()).any(|window| window == raw)
}

#[test]
fn changed_name_reopens_invalidates_previews_and_preserves_locality() -> TestResult {
    let source = synthetic_package(false, true)?;
    let package = Package::from_bytes(&source)?;
    let before_model = model_member_payload(&source, MODELS[0])?;
    let before_other_model = model_member_payload(&source, MODELS[1])?;
    let before_document = member_bytes(&source, DOCUMENT_MEMBER)?;
    let before_metadata = member_bytes(&source, METADATA_MEMBER)?;
    let before_sentinel = member_bytes(&source, SENTINEL_MEMBER)?;
    let before_row_bucket = object_message_payload(
        &source,
        MODEL_MEMBER,
        ROW_BUCKETS[0],
        HEADER_BUCKET_MESSAGE_TYPE,
    )?;
    let replacement = SlideTableName::new("Revenué 🚀")?;
    let commit = package
        .edit_slide_table_name("Tables", 0usize)?
        .set(replacement.clone())
        .commit()?;
    assert_eq!(
        commit.package().slide_table_name("Tables", 0usize)?,
        replacement
    );
    assert_eq!(commit.patch().before().as_str(), "Revenue");
    assert_eq!(commit.patch().after(), &replacement);
    assert!(!commit.patch().is_noop());
    assert_ne!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
    assert!(commit.diagnostics().full_reparse_performed());

    let target = exact_bytes(commit.package())?;
    let target_model = model_member_payload(&target, MODELS[0])?;
    assert_ne!(target_model, before_model);
    assert_eq!(
        non_name_fields(&target_model)?,
        non_name_fields(&before_model)?
    );
    let mut unknown = Vec::new();
    append_varint_field(&mut unknown, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)?;
    assert!(contains_raw(&target_model, &unknown));
    assert_eq!(
        model_member_payload(&target, MODELS[1])?,
        before_other_model
    );
    assert_eq!(member_bytes(&target, DOCUMENT_MEMBER)?, before_document);
    assert_eq!(member_bytes(&target, METADATA_MEMBER)?, before_metadata);
    assert_eq!(member_bytes(&target, SENTINEL_MEMBER)?, before_sentinel);
    assert_eq!(
        object_message_payload(
            &target,
            MODEL_MEMBER,
            ROW_BUCKETS[0],
            HEADER_BUCKET_MESSAGE_TYPE,
        )?,
        before_row_bucket
    );
    for preview in PREVIEWS {
        assert!(!has_member(&target, preview)?);
    }
    let reopened = Package::from_bytes(&target)?;
    assert_eq!(
        reopened.slide_table_name(SlideSelector::index(0), TableSelector::index(0))?,
        replacement
    );
    Ok(())
}

#[test]
fn inverse_restores_exact_source_and_repeated_apply_conflicts() -> TestResult {
    let source = synthetic_package(false, false)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_table_name("Tables", 0usize)?
        .set_name("Candidate")?
        .commit()?;
    let target = exact_bytes(commit.package())?;
    let applied = package.apply_slide_table_name(commit.patch())?;
    assert_eq!(exact_bytes(applied.package())?, target);
    assert!(matches!(
        commit.package().apply_slide_table_name(commit.patch()),
        Err(SlideTableNameError::PatchConflict)
    ));
    let restored = commit
        .package()
        .apply_slide_table_name(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored
            .package()
            .slide_table_name("Tables", 0usize)?
            .as_str(),
        "Revenue"
    );
    assert!(matches!(
        package.apply_slide_table_name(&commit.patch().inverse()),
        Err(SlideTableNameError::PatchConflict)
    ));
    Ok(())
}

#[test]
fn stale_patch_conflict_is_source_bound() -> TestResult {
    let source = synthetic_package(false, false)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_table_name("Tables", 0usize)?
        .set_name("Candidate")?
        .commit()?;
    let catalog = Catalog::from_bytes(&source)?;
    let tampered_source = catalog.reassemble_to_bytes(
        &[EntryEdit::new(SENTINEL_MEMBER, b"tampered")],
        Limits::default(),
    )?;
    let tampered = Package::from_bytes(&tampered_source)?;
    assert!(matches!(
        tampered.apply_slide_table_name(commit.patch()),
        Err(SlideTableNameError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&tampered)?, tampered_source);
    Ok(())
}

#[test]
fn locked_table_allows_no_op_but_rejects_changed_name() -> TestResult {
    let source = synthetic_package(true, false)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let current = package.slide_table_name("Tables", 0usize)?;
    let noop = package
        .edit_slide_table_name("Tables", 0usize)?
        .set(current)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, before);
    let error = package
        .edit_slide_table_name("Tables", 0usize)?
        .set_name("locked replacement")
        .expect("valid replacement name")
        .commit()
        .expect_err("locked table must reject changed name");
    assert!(matches!(error, SlideTableNameError::Locked));
    assert_eq!(exact_bytes(&package)?, before);

    let unlocked = package
        .edit_slide_table_name("Tables", 1usize)?
        .set_name("second replacement")?
        .commit()?;
    assert_eq!(
        unlocked
            .package()
            .slide_table_name("Tables", 1usize)?
            .as_str(),
        "second replacement"
    );
    Ok(())
}

#[test]
fn malformed_field8_wire_is_rejected_without_mutation() -> TestResult {
    let source = synthetic_package(false, false)?;
    let malformed = [
        vec![0x42, 0x01, b'X'],       // duplicate field 8
        vec![0x40, 0x01],             // field 8 with varint wire
        vec![0x42, 0x81, 0x00, b'X'], // noncanonical length
        vec![0x42],                   // truncated length-delimited field
        vec![0x42, 0x02, 0xff, 0xff], // invalid UTF-8
        vec![0xa3, 0x06, 0x08, 0x01], // unterminated unknown group
    ];
    for raw in malformed {
        assert_rejected_atomically(&append_selected_model_raw(&source, &raw)?)?;
    }

    let payload = model_member_payload(&source, MODELS[0])?;
    let without_name = WireView::parse(&payload)?
        .fields()
        .filter(|field| field.number() != 8)
        .flat_map(|field| field.raw().iter().copied())
        .collect::<Vec<_>>();
    assert_rejected_atomically(&replace_model_payload(&source, MODELS[0], without_name)?)?;
    Ok(())
}

#[test]
fn unknown_model_fields_and_groups_survive_name_rewrite() -> TestResult {
    let source = unknown_model_group(&synthetic_package(false, true)?)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_table_name("Tables", 0usize)?
        .set_name("Preserved unknowns")?
        .commit()?;
    let payload = model_member_payload(&exact_bytes(commit.package())?, MODELS[0])?;
    let mut unknown = Vec::new();
    append_varint_field(&mut unknown, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)?;
    assert!(contains_raw(&payload, &unknown));
    assert!(contains_raw(
        &payload,
        &[0x9b, 0x06, 0x08, 0x01, 0x9c, 0x06]
    ));
    Ok(())
}

#[test]
fn canonical_and_legacy_model_roles_fail_closed() -> TestResult {
    let source = synthetic_package(false, false)?;
    for malformed in [
        duplicate_canonical_model(&source)?,
        model_role_alias(&source, LEGACY_TABLE_MODEL_MESSAGE_TYPE)?,
        model_role_alias(&source, TABLE_STYLE_MESSAGE_TYPE)?,
        table_info_model_role_alias(&source)?,
    ] {
        assert_rejected_atomically(&malformed)?;
    }
    Ok(())
}

fn mutate_table_info_refs(
    package: &[u8],
    mutate: impl FnOnce(&mut Vec<u64>),
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or_else(|| io::Error::other("missing table info"))?;
        let info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("missing table info metadata"))?;
        mutate(&mut info.object_references);
        Ok(())
    })
}

#[test]
fn archive_info_and_global_inbound_authority_fail_closed() -> TestResult {
    let source = synthetic_package(false, false)?;
    let table_info_foreign = rewrite_document_archive(&source, |archive| {
        let object = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or_else(|| io::Error::other("missing table info"))?;
        let info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("missing table info metadata"))?;
        info.field_infos.push(field_reference(vec![99], &[999_999]));
        Ok(())
    })?;
    let table_info_missing_model = mutate_table_info_refs(&source, |refs| {
        refs.retain(|identifier| *identifier != MODELS[0]);
    })?;
    let malformed = [
        ("table_info_foreign", table_info_foreign),
        ("table_info_missing_model", table_info_missing_model),
        (
            "missing_model_storage_route",
            missing_model_storage_route(&source)?,
        ),
        (
            "foreign_model_inbound",
            foreign_inbound(&source, MODELS[0], FOREIGN_INBOUND_OBJECT)?,
        ),
        (
            "foreign_bucket_inbound",
            foreign_inbound(&source, ROW_BUCKETS[0], FOREIGN_BUCKET_INBOUND_OBJECT)?,
        ),
        (
            "duplicate_slide_identity",
            duplicate_slide_identity(&source)?,
        ),
    ];
    for (label, hostile) in malformed {
        assert_rejected_atomically(&hostile)
            .map_err(|error| io::Error::other(format!("{label}: {error}")))?;
    }
    Ok(())
}

#[test]
fn metadata_uuid_and_component_authority_fail_closed() -> TestResult {
    let source = synthetic_package(false, false)?;
    let missing_model_uuid = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 2)
        {
            component
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != MODELS[0]);
        }
        Ok(())
    })?;
    let duplicate_model_uuid = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 2)
        {
            if let Some(entry) = component
                .object_uuid_map_entries
                .iter()
                .find(|entry| entry.identifier == MODELS[0])
                .cloned()
            {
                component.object_uuid_map_entries.push(entry);
            }
        }
        Ok(())
    })?;
    let substituted_model_component = rewrite_metadata(&source, |metadata| {
        let model_entry = metadata
            .components
            .iter()
            .find(|component| component.identifier == 2)
            .and_then(|component| {
                component
                    .object_uuid_map_entries
                    .iter()
                    .find(|entry| entry.identifier == MODELS[0])
                    .cloned()
            });
        if let Some(model_entry) = model_entry {
            if let Some(component) = metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == 2)
            {
                component
                    .object_uuid_map_entries
                    .retain(|entry| entry.identifier != MODELS[0]);
            }
            if let Some(component) = metadata
                .components
                .iter_mut()
                .find(|component| component.identifier == 1)
            {
                component.object_uuid_map_entries.push(model_entry);
            }
        }
        Ok(())
    })?;
    let duplicate_uuid_pair = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 2)
        {
            let model_uuid = component
                .object_uuid_map_entries
                .iter()
                .find(|entry| entry.identifier == MODELS[0])
                .map(|entry| entry.uuid);
            if let (Some(model_uuid), Some(other)) = (
                model_uuid,
                component
                    .object_uuid_map_entries
                    .iter_mut()
                    .find(|entry| entry.identifier == DATA_OBJECTS[0][0]),
            ) {
                other.uuid = model_uuid;
            }
        }
        Ok(())
    })?;
    let root_map = rewrite_metadata(&source, |metadata| {
        metadata.data_metadata_map = Some(reference(MODELS[0]));
        Ok(())
    })?;
    let external_reference = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 2)
        {
            component
                .versioned_external_references
                .push(tsp::ComponentExternalReference {
                    component_identifier: 1,
                    object_identifier: Some(MODELS[0]),
                    is_weak: Some(false),
                });
        }
        Ok(())
    })?;
    let data_reference = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 2)
        {
            component.data_references.push(tsp::ComponentDataReference {
                data_identifier: MODELS[0],
                object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                    object_identifier: MODELS[0],
                    count: 1,
                }],
            });
        }
        Ok(())
    })?;
    let ambiguous_identifier = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 2)
        {
            component.ambiguous_object_identifiers.push(MODELS[0]);
        }
        Ok(())
    })?;

    for (label, hostile) in [
        ("missing_model_uuid", missing_model_uuid),
        ("duplicate_model_uuid", duplicate_model_uuid),
        ("substituted_model_component", substituted_model_component),
        ("duplicate_uuid_pair", duplicate_uuid_pair),
        ("root_map", root_map),
        ("external_reference", external_reference),
        ("data_reference", data_reference),
        ("ambiguous_identifier", ambiguous_identifier),
        ("metadata_unknown", append_metadata_unknown(&source)?),
    ] {
        assert_rejected_atomically(&hostile)
            .map_err(|error| io::Error::other(format!("{label}: {error}")))?;
    }
    Ok(())
}

#[test]
fn model_storage_payload_routes_must_match_archive_metadata() -> TestResult {
    let source = synthetic_package(false, false)?;
    let hostile = [
        (
            "row-storage-physical-substitution",
            rewrite_model_storage(&source, |store| {
                store.row_headers.buckets[0] = reference(ROW_BUCKETS[1]);
            })?,
        ),
        (
            "column-storage-physical-substitution",
            rewrite_model_storage(&source, |store| {
                store.column_headers = reference(COLUMN_BUCKETS[1]);
            })?,
        ),
        (
            "data-storage-dangling-substitution",
            rewrite_model_storage(&source, |store| {
                store.string_table = reference(999_999);
            })?,
        ),
    ];
    for (label, package) in hostile {
        assert_rejected_atomically(&package)
            .map_err(|error| io::Error::other(format!("{label}: {error}")))?;
    }
    Ok(())
}

#[test]
fn model_storage_archive_routes_are_unique_and_canonical() -> TestResult {
    let source = synthetic_package(false, false)?;
    let hostile = [
        (
            "duplicate-storage-aggregate",
            append_duplicate_storage_aggregate(&source)?,
        ),
        (
            "duplicate-storage-field",
            append_duplicate_storage_field(&source)?,
        ),
        (
            "noncanonical-storage-field",
            append_noncanonical_storage_field(&source)?,
        ),
    ];
    for (label, package) in hostile {
        assert_rejected_atomically(&package)
            .map_err(|error| io::Error::other(format!("{label}: {error}")))?;
    }
    Ok(())
}

#[test]
fn storage_metadata_uuid_and_role_authority_fail_closed() -> TestResult {
    let source = synthetic_package(false, false)?;
    let hostile = [
        (
            "missing-row-storage-uuid",
            remove_metadata_uuid(&source, ROW_BUCKETS[0])?,
        ),
        (
            "missing-data-storage-uuid",
            remove_metadata_uuid(&source, DATA_OBJECTS[0][0])?,
        ),
        (
            "wrong-component-column-storage-uuid",
            move_metadata_uuid_to_document(&source, COLUMN_BUCKETS[0])?,
        ),
        (
            "storage-bucket-role-alias",
            append_storage_role_alias(&source, ROW_BUCKETS[0], TABLE_MODEL_MESSAGE_TYPE)?,
        ),
        (
            "string-table-role-alias",
            append_storage_role_alias(&source, DATA_OBJECTS[0][0], TABLE_INFO_MESSAGE_TYPE)?,
        ),
        (
            "style-table-role-alias",
            append_storage_role_alias(&source, DATA_OBJECTS[0][1], TABLE_MODEL_MESSAGE_TYPE)?,
        ),
        (
            "formula-table-role-alias",
            append_storage_role_alias(&source, DATA_OBJECTS[0][2], TABLE_STYLE_MESSAGE_TYPE)?,
        ),
        (
            "format-table-role-alias",
            append_storage_role_alias(&source, DATA_OBJECTS[0][3], STYLESHEET_MESSAGE_TYPE)?,
        ),
    ];
    for (label, package) in hostile {
        assert_rejected_atomically(&package)
            .map_err(|error| io::Error::other(format!("{label}: {error}")))?;
    }
    Ok(())
}

#[test]
fn name_value_validation_accepts_unicode_and_rejects_invalid_values() -> TestResult {
    assert!(matches!(
        SlideTableName::new(""),
        Err(SlideTableNameValueError::Empty)
    ));
    assert!(matches!(
        SlideTableName::new("bad\0name"),
        Err(SlideTableNameValueError::ContainsNul)
    ));
    let unicode = SlideTableName::new("收入表 — Café №42")?;
    assert_eq!(unicode.as_str(), "收入表 — Café №42");
    Ok(())
}

#[test]
fn finite_limits_refuse_before_publication_and_preserve_source() -> TestResult {
    let source = synthetic_package(false, false)?;
    let input_limit = Limits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        128,
        4 * 1024 * 1024,
        8 * 1024 * 1024,
        8 * 1024 * 1024,
    )?;
    assert!(Package::from_bytes_with_limits(&source, input_limit).is_err());

    let mut found_semantic_limit = false;
    for references in 1..=128 {
        let semantic = SemanticLimits::new(
            1_000_000,
            65_536,
            references,
            1_000_000,
            1_000_000,
            64 * 1024 * 1024,
        )?;
        let Ok(restricted) = Package::from_bytes_with_options(
            &source,
            ReadOptions::new(Limits::default(), semantic),
        ) else {
            continue;
        };
        let before = exact_bytes(&restricted)?;
        let read = restricted.slide_table_name("Tables", 0usize);
        if matches!(read, Err(SlideTableNameError::LimitExceeded { .. })) {
            assert!(matches!(
                restricted.edit_slide_table_name("Tables", 0usize),
                Err(SlideTableNameError::LimitExceeded { .. })
            ));
            assert_eq!(exact_bytes(&restricted)?, before);
            found_semantic_limit = true;
            break;
        }
    }
    assert!(
        found_semantic_limit,
        "no semantic reference limit was observed"
    );

    let exact_fields = (1..=4_096)
        .find_map(|fields| {
            let archive = litchi_iwa_core::Limits::default()
                .with_header_fields(fields)
                .ok()?;
            let limits = Limits::default().with_archive_limits(archive).ok()?;
            let package = Package::from_bytes_with_limits(&source, limits).ok()?;
            package
                .slide_table_name("Tables", 0usize)
                .is_ok()
                .then_some(fields)
        })
        .ok_or_else(|| io::Error::other("no finite header-field boundary found"))?;
    assert!(exact_fields > 1);
    let archive = litchi_iwa_core::Limits::default().with_header_fields(exact_fields - 1)?;
    let limits = Limits::default().with_archive_limits(archive)?;
    if let Ok(package) = Package::from_bytes_with_limits(&source, limits) {
        let before = exact_bytes(&package)?;
        assert!(matches!(
            package.slide_table_name("Tables", 0usize),
            Err(SlideTableNameError::LimitExceeded { .. })
        ));
        assert_eq!(exact_bytes(&package)?, before);
    }
    Ok(())
}
