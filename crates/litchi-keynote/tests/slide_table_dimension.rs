//! Exact-source integration coverage for Keynote slide-table dimensions.
//!
//! The fixture intentionally keeps the presentation graph in one IWA member
//! and the table models/header buckets in another.  This exercises the
//! selector-first owner without exposing native object identifiers, while the
//! tests verify that a row/column edit rewrites only the selected bucket and
//! the owning drawable geometry.

use std::error::Error as StdError;
use std::io;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::{
    decode_varint_from_bytes,
    wire::{WireView, append_length_delimited_field, append_varint_field},
};
use litchi_iwa_core::{Archive, ArchiveObject, FieldInfo, FieldPath, RawMessage, SnappyStream};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp, tst, tswp};
use litchi_keynote::slide::table::dimension::{Dimension, Points, Size};
use litchi_keynote::{
    Package, ReadOptions, SemanticLimits, SlideSelector, SlideTableDimensionCommit,
    SlideTableDimensionDiagnostics, SlideTableDimensionEdit, SlideTableDimensionError,
    SlideTableDimensionLimitKind, SlideTableDimensionPatch, SlideTableDimensionPath, TableSelector,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const MODEL_MEMBER: &str = "Index/CalculationEngine.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const SENTINEL_MEMBER: &str = "Data/dimension-sentinel.bin";
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
            // 4 columns at 64pt except column 2 at 40pt, and eight rows at
            // 20pt except row 1 at 25pt.
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
            (SENTINEL_MEMBER, b"untouched-dimension-sentinel".as_slice()),
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
    let archive = archive(package, MODEL_MEMBER)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing model object"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing model message").into())
}

fn table_info_geometry(package: &[u8], identifier: u64) -> TestResult<tsp::Size> {
    let archive = archive(package, DOCUMENT_MEMBER)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing table-info object"))?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_INFO_MESSAGE_TYPE)
        .ok_or_else(|| io::Error::other("missing table-info message"))?;
    let info = tst::TableInfoArchive::decode(message.data.as_slice())?;
    info.super_
        .geometry
        .and_then(|geometry| geometry.size)
        .ok_or_else(|| io::Error::other("missing table geometry").into())
}

fn bucket_payload_bytes(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let archive = archive(package, MODEL_MEMBER)?;
    let object = archive
        .object(identifier)
        .ok_or_else(|| io::Error::other("missing bucket object"))?;
    object
        .messages
        .iter()
        .find(|message| message.type_ == HEADER_BUCKET_MESSAGE_TYPE)
        .map(|message| message.data.clone())
        .ok_or_else(|| io::Error::other("missing bucket message").into())
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

fn replace_first_length_delimited_field(
    payload: &[u8],
    number: u32,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    let fields = WireView::parse(payload)?;
    let mut rewritten = Vec::new();
    let mut replaced = false;
    for field in fields.fields() {
        if field.number() == number && !replaced {
            if field.wire_type() != 2 {
                return Err(io::Error::other("target field is not length-delimited").into());
            }
            append_length_delimited_field(&mut rewritten, number, replacement)?;
            replaced = true;
        } else {
            rewritten.extend_from_slice(field.raw());
        }
    }
    if !replaced {
        return Err(io::Error::other("missing length-delimited field").into());
    }
    Ok(rewritten)
}

fn append_duplicate_row_bucket_route(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let model = model_member_payload(package, identifier)?;
    let store = WireView::parse(&model)?
        .fields()
        .find(|field| field.number() == MODEL_STORAGE_FIELD)
        .ok_or_else(|| io::Error::other("missing data-store field"))?;
    let row_headers = WireView::parse(store.payload())?
        .fields()
        .find(|field| field.number() == 1)
        .ok_or_else(|| io::Error::other("missing row-headers field"))?;
    let bucket = WireView::parse(row_headers.payload())?
        .fields()
        .find(|field| field.number() == 2)
        .ok_or_else(|| io::Error::other("missing row bucket route"))?;
    let mut duplicate_row_headers = row_headers.payload().to_vec();
    append_length_delimited_field(&mut duplicate_row_headers, 2, bucket.payload())?;
    let rewritten_store =
        replace_first_length_delimited_field(store.payload(), 1, &duplicate_row_headers)?;
    let rewritten_model =
        replace_first_length_delimited_field(&model, MODEL_STORAGE_FIELD, &rewritten_store)?;
    replace_model_payload(package, identifier, rewritten_model)
}

fn append_noncanonical_model_bucket_field(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(identifier)
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

fn rewrite_bucket_payload(
    package: &[u8],
    identifier: u64,
    payload: Vec<u8>,
) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("missing bucket object"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == HEADER_BUCKET_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing bucket message"))?;
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: HEADER_BUCKET_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })
}

fn append_duplicate_header(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    let original = bucket_payload_bytes(package, identifier)?;
    let header = WireView::parse(&original)?
        .fields()
        .find(|field| field.number() == 2)
        .ok_or_else(|| io::Error::other("missing header"))?
        .payload()
        .to_vec();
    let mut duplicate = original;
    append_length_delimited_field(&mut duplicate, 2, &header)?;
    rewrite_bucket_payload(package, identifier, duplicate)
}

fn replace_bucket_message_type(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("missing bucket object"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == HEADER_BUCKET_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing bucket message"))?;
        object.messages[message_index].type_ += 1;
        object.archive_info.message_infos[message_index].type_ += 1;
        Ok(())
    })
}

fn remove_bucket(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        archive
            .objects
            .retain(|object| object.archive_info.identifier != Some(identifier));
        Ok(())
    })
}

fn replace_model_role(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("missing model object"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing model message"))?;
        object.messages[message_index].type_ = 6_003;
        object.archive_info.message_infos[message_index].type_ = 6_003;
        Ok(())
    })
}

fn drop_model_aggregate_reference(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("missing model object"))?;
        object.archive_info.message_infos[0]
            .object_references
            .retain(|value| *value != ROW_BUCKETS[0]);
        Ok(())
    })
}

fn append_bucket_role_alias(
    package: &[u8],
    identifier: u64,
    alias_type: u32,
) -> TestResult<Vec<u8>> {
    rewrite_model_archive(package, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or_else(|| io::Error::other("missing bucket object"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == HEADER_BUCKET_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing bucket message"))?;
        let mut alias = object.messages[message_index].clone();
        alias.type_ = alias_type;
        let mut alias_info = object.archive_info.message_infos[message_index].clone();
        alias_info.type_ = alias_type;
        object.messages.push(alias);
        object.archive_info.message_infos.push(alias_info);
        Ok(())
    })
}

fn drop_unrelated_drawable_metadata(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(SLIDE)
            .ok_or_else(|| io::Error::other("missing slide object"))?;
        let info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or_else(|| io::Error::other("missing slide metadata"))?;
        info.object_references
            .retain(|identifier| *identifier != NON_TABLE_DRAWABLE);
        for field in &mut info.field_infos {
            if field.path.as_slice() == [SLIDE_OWNED_DRAWABLES_FIELD]
                || field.path.as_slice() == [SLIDE_Z_ORDER_FIELD]
            {
                field
                    .object_references
                    .retain(|identifier| *identifier != NON_TABLE_DRAWABLE);
            }
        }
        Ok(())
    })
}

fn append_foreign_bucket_inbound(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        archive.objects.push(object(
            FOREIGN_BUCKET_INBOUND_OBJECT,
            7_001,
            vec![0x08, 0x01],
            &[ROW_BUCKETS[0]],
        )?);
        Ok(())
    })
}

fn append_duplicate_slide_identity(package: &[u8]) -> TestResult<Vec<u8>> {
    let slide = archive(package, DOCUMENT_MEMBER)?
        .object(SLIDE)
        .ok_or_else(|| io::Error::other("missing slide object"))?
        .clone();
    rewrite_model_archive(package, |archive| {
        archive.objects.push(slide);
        Ok(())
    })
}

fn append_foreign_inbound(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        archive.objects.push(object(
            FOREIGN_INBOUND_OBJECT,
            7_001,
            vec![0x08, 0x01],
            &[MODELS[0]],
        )?);
        Ok(())
    })
}

fn append_unknown_metadata(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_member_archive(package, METADATA_MEMBER, |archive| {
        let object = archive
            .object_mut(METADATA_OBJECT)
            .ok_or_else(|| io::Error::other("missing metadata object"))?;
        let message_index = object
            .messages
            .iter()
            .position(|message| message.type_ == METADATA_MESSAGE_TYPE)
            .ok_or_else(|| io::Error::other("missing metadata message"))?;
        let mut payload = object.messages[message_index].data.clone();
        append_varint_field(&mut payload, 99, 123_456)?;
        object.replace_message_preserving_header(
            message_index,
            RawMessage {
                type_: METADATA_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })
}

fn rewrite_metadata<F>(package: &[u8], rewrite: F) -> TestResult<Vec<u8>>
where
    F: FnOnce(&mut tsp::PackageMetadata),
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
        rewrite(&mut metadata);
        message.data = metadata.encode_to_vec();
        Ok(())
    })
}

fn assert_rejected_atomically(source: &[u8], dimension: Dimension) -> TestResult {
    let package = match Package::from_bytes(source) {
        Ok(package) => package,
        Err(_) => return Ok(()),
    };
    let before = exact_bytes(&package)?;
    let read = package.slide_table_dimension_size(SlideSelector::index(0), 0usize, dimension);
    let edit = package
        .edit_slide_table_dimension_size(SlideSelector::index(0), 0usize, dimension)
        .and_then(|edit| edit.set(Size::points(31.0).unwrap()).commit());
    assert!(
        matches!(
            read,
            Err(SlideTableDimensionError::InvalidSource)
                | Err(SlideTableDimensionError::UnsupportedSource)
                | Err(SlideTableDimensionError::UnsupportedDependency)
                | Err(SlideTableDimensionError::UnsupportedTopology)
                | Err(SlideTableDimensionError::LimitExceeded { .. })
        ),
        "malformed dimension read was unexpectedly accepted: {read:?}"
    );
    assert!(
        matches!(
            edit,
            Err(SlideTableDimensionError::InvalidSource)
                | Err(SlideTableDimensionError::UnsupportedSource)
                | Err(SlideTableDimensionError::UnsupportedDependency)
                | Err(SlideTableDimensionError::UnsupportedTopology)
                | Err(SlideTableDimensionError::LimitExceeded { .. })
        ),
        "malformed dimension edit was unexpectedly accepted: {edit:?}"
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

fn points(value: f32) -> TestResult<Size> {
    Ok(Size::points(value)?)
}

#[test]
fn typed_selectors_read_rows_columns_and_redacted_transactions() -> TestResult {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}

    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Dimension>();
    assert_send_sync_debug::<Points>();
    assert_send_sync_debug::<Size>();
    assert_send_sync_debug::<SlideTableDimensionEdit<'static>>();
    assert_send_sync_debug::<SlideTableDimensionPatch>();
    assert_send_sync_debug::<SlideTableDimensionCommit>();
    assert_send_sync_debug::<SlideTableDimensionDiagnostics>();
    assert_send_sync_debug::<SlideTableDimensionError>();
    assert_send_sync_debug::<SlideTableDimensionLimitKind>();
    assert_send_sync_debug::<SlideTableDimensionPath>();

    let source = synthetic_package(false, false)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_table_dimension_size("Tables", 0usize, Dimension::Row(0))?,
        Size::Default
    );
    assert_eq!(
        package.slide_table_dimension_size(
            SlideSelector::index(0),
            TableSelector::index(0),
            Dimension::Row(1),
        )?,
        points(25.0)?
    );
    assert_eq!(
        package.slide_table_dimension_size("Tables", 0usize, Dimension::Column(2))?,
        points(40.0)?
    );
    assert_eq!(
        package.slide_table_dimension_size("Tables", 1usize, Dimension::Row(0))?,
        Size::Default
    );
    assert!(matches!(
        package.slide_table_dimension_size("Missing", 0usize, Dimension::Row(0)),
        Err(SlideTableDimensionError::SlideNameNotFound)
    ));
    assert!(matches!(
        package.slide_table_dimension_size("Tables", 2usize, Dimension::Row(0)),
        Err(SlideTableDimensionError::TablePositionNotFound { .. })
    ));
    assert!(matches!(
        package.slide_table_dimension_size("Tables", 0usize, Dimension::Row(8)),
        Err(SlideTableDimensionError::InvalidSource)
    ));

    let edit = package.edit_slide_table_dimension_size("Tables", 0usize, Dimension::Row(1))?;
    assert_eq!(
        edit.path(),
        SlideTableDimensionPath::Table {
            slide: litchi_core::Position::new(0),
            table: litchi_core::Position::new(0),
            dimension: Dimension::Row(1),
        }
    );
    assert_eq!(edit.before(), points(25.0)?);
    let rendered = format!("{edit:?}");
    assert!(!rendered.contains(DOCUMENT_MEMBER));
    assert!(!rendered.contains("110"));
    Ok(())
}

#[test]
fn no_op_is_exact_and_reset_restores_default() -> TestResult {
    let source = synthetic_package(false, false)?;
    let package = Package::from_bytes(&source)?;
    let before = package.slide_table_dimension_size("Tables", 0usize, Dimension::Row(0))?;
    let noop = package
        .edit_slide_table_dimension_size("Tables", 0usize, Dimension::Row(0))?
        .set(before)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(noop.patch().before(), Size::Default);
    assert_eq!(noop.patch().after(), Size::Default);
    assert!(!noop.diagnostics().changed());
    assert_eq!(noop.diagnostics().touched_components(), 0);
    assert_eq!(noop.diagnostics().deleted_previews(), 0);
    assert!(!noop.diagnostics().full_reparse_performed());
    assert_eq!(exact_bytes(noop.package())?, source);

    let changed = package
        .edit_slide_table_dimension_size("Tables", 0usize, Dimension::Row(0))?
        .set(points(32.0)?)
        .commit()?;
    assert_eq!(
        changed
            .package()
            .slide_table_dimension_size("Tables", 0usize, Dimension::Row(0))?,
        points(32.0)?
    );
    let reset = changed
        .package()
        .edit_slide_table_dimension_size("Tables", 0usize, Dimension::Row(0))?
        .reset()
        .commit()?;
    assert_eq!(
        reset
            .package()
            .slide_table_dimension_size("Tables", 0usize, Dimension::Row(0))?,
        Size::Default
    );
    let reset_bytes = exact_bytes(reset.package())?;
    assert_eq!(
        member_bytes(&reset_bytes, SENTINEL_MEMBER)?,
        b"untouched-dimension-sentinel"
    );
    Ok(())
}

#[test]
fn row_and_column_changes_update_geometry_and_round_trip_inverse() -> TestResult {
    let source = synthetic_package(false, false)?;
    for (dimension, after, expected_geometry) in [
        (
            Dimension::Row(0),
            points(32.0)?,
            tsp::Size {
                width: 232.0,
                height: 177.0,
            },
        ),
        (
            Dimension::Column(2),
            points(80.0)?,
            tsp::Size {
                width: 272.0,
                height: 165.0,
            },
        ),
    ] {
        let package = Package::from_bytes(&source)?;
        let source_document = member_bytes(&source, DOCUMENT_MEMBER)?;
        let source_model = member_bytes(&source, MODEL_MEMBER)?;
        let source_metadata = member_bytes(&source, METADATA_MEMBER)?;
        let source_sentinel = member_bytes(&source, SENTINEL_MEMBER)?;
        let before = package.slide_table_dimension_size("Tables", 0usize, dimension)?;
        let before_geometry = table_info_geometry(&source, TABLE_INFOS[0])?;
        let commit = package
            .edit_slide_table_dimension_size("Tables", 0usize, dimension)?
            .set(after)
            .commit()?;
        assert_eq!(commit.patch().dimension(), dimension);
        assert_eq!(commit.patch().before(), before);
        assert_eq!(commit.patch().after(), after);
        assert!(!commit.patch().is_noop());
        assert_ne!(
            commit.patch().source_fingerprint(),
            commit.patch().target_fingerprint()
        );
        assert!(commit.diagnostics().changed());
        assert!(commit.diagnostics().full_reparse_performed());
        assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
        assert_eq!(
            table_info_geometry(&exact_bytes(commit.package())?, TABLE_INFOS[0])?,
            expected_geometry
        );
        assert_ne!(before_geometry, expected_geometry);
        assert_eq!(
            commit
                .package()
                .slide_table_dimension_size("Tables", 0usize, dimension)?,
            after
        );

        let target = exact_bytes(commit.package())?;
        assert_ne!(member_bytes(&target, DOCUMENT_MEMBER)?, source_document);
        assert_ne!(member_bytes(&target, MODEL_MEMBER)?, source_model);
        assert_eq!(member_bytes(&target, METADATA_MEMBER)?, source_metadata);
        assert_eq!(member_bytes(&target, SENTINEL_MEMBER)?, source_sentinel);
        for preview in PREVIEWS {
            assert!(!has_member(&target, preview)?);
        }
        assert_eq!(
            exact_bytes(
                &package
                    .apply_slide_table_dimension_size(commit.patch())?
                    .into_package()
            )?,
            target
        );
        let reopened = Package::from_bytes(&target)?;
        assert!(matches!(
            reopened.apply_slide_table_dimension_size(commit.patch()),
            Err(SlideTableDimensionError::PatchConflict)
        ));
        let restored = reopened.apply_slide_table_dimension_size(&commit.patch().inverse())?;
        assert_eq!(exact_bytes(restored.package())?, source);
    }
    Ok(())
}

#[test]
fn unknown_model_bucket_and_header_fields_survive_rewrite() -> TestResult {
    let source = synthetic_package(false, true)?;
    let package = Package::from_bytes(&source)?;
    let before_model = model_member_payload(&source, MODELS[0])?;
    let before_bucket = bucket_payload_bytes(&source, ROW_BUCKETS[0])?;
    let commit = package
        .edit_slide_table_dimension_size("Tables", 0usize, Dimension::Row(1))?
        .set(points(31.0)?)
        .commit()?;
    let target = exact_bytes(commit.package())?;
    let target_model = model_member_payload(&target, MODELS[0])?;
    let model = WireView::parse(&target_model)?;
    let model_unknown = model
        .fields()
        .find(|field| field.number() == UNKNOWN_MODEL_FIELD)
        .ok_or_else(|| io::Error::other("model unknown field was dropped"))?;
    assert_eq!(
        decode_varint_from_bytes(model_unknown.payload())?.0,
        UNKNOWN_MODEL_VALUE
    );
    let target_bucket = bucket_payload_bytes(&target, ROW_BUCKETS[0])?;
    let bucket = WireView::parse(&target_bucket)?;
    let bucket_unknown = bucket
        .fields()
        .find(|field| field.number() == UNKNOWN_BUCKET_FIELD)
        .ok_or_else(|| io::Error::other("bucket unknown field was dropped"))?;
    assert_eq!(
        decode_varint_from_bytes(bucket_unknown.payload())?.0,
        UNKNOWN_BUCKET_VALUE
    );
    let header = bucket
        .fields()
        .find(|field| field.number() == 2)
        .ok_or_else(|| io::Error::other("header was dropped"))?;
    let header_unknown = WireView::parse(header.payload())?
        .fields()
        .find(|field| field.number() == UNKNOWN_HEADER_FIELD)
        .ok_or_else(|| io::Error::other("header unknown field was dropped"))?;
    assert_eq!(
        decode_varint_from_bytes(header_unknown.payload())?.0,
        UNKNOWN_HEADER_VALUE
    );
    assert_eq!(before_model, model_member_payload(&target, MODELS[0])?);
    assert_ne!(
        before_bucket,
        bucket_payload_bytes(&target, ROW_BUCKETS[0])?
    );
    Ok(())
}

#[test]
fn locked_table_allows_noop_but_rejects_changed_dimension_atomically() -> TestResult {
    let source = synthetic_package(true, false)?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    let noop = package
        .edit_slide_table_dimension_size("Tables", 0usize, Dimension::Row(0))?
        .set(Size::Default)
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, before);
    let error = package
        .edit_slide_table_dimension_size("Tables", 0usize, Dimension::Row(0))?
        .set(points(32.0)?)
        .commit()
        .expect_err("locked table must reject changed dimension");
    assert!(matches!(error, SlideTableDimensionError::Locked));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn malformed_duplicate_wrong_role_archive_and_metadata_routes_fail_closed() -> TestResult {
    let source = synthetic_package(false, false)?;
    let malformed = [
        append_duplicate_header(&source, ROW_BUCKETS[0])?,
        replace_bucket_message_type(&source, ROW_BUCKETS[0])?,
        remove_bucket(&source, ROW_BUCKETS[0])?,
        replace_model_role(&source, MODELS[0])?,
        drop_model_aggregate_reference(&source, MODELS[0])?,
        append_foreign_inbound(&source)?,
        append_unknown_metadata(&source)?,
    ];
    let mut model = model_member_payload(&source, MODELS[0])?;
    append_varint_field(&mut model, 6, 99)?;
    let malformed_model = replace_model_payload(&source, MODELS[0], model)?;
    for hostile in malformed.into_iter().chain([malformed_model]) {
        assert_rejected_atomically(&hostile, Dimension::Row(1))?;
        assert_rejected_atomically(&hostile, Dimension::Column(2))?;
    }
    Ok(())
}

#[test]
fn cross_axis_indices_use_their_own_dimension_limits() -> TestResult {
    let package = Package::from_bytes(&synthetic_package(false, false)?)?;
    assert_eq!(
        package.slide_table_dimension_size(SlideSelector::index(0), 0usize, Dimension::Row(6),)?,
        Size::Default
    );
    assert_eq!(
        package.slide_table_dimension_size(
            SlideSelector::index(0),
            0usize,
            Dimension::Column(3),
        )?,
        Size::Default
    );
    Ok(())
}

#[test]
fn bucket_role_alias_and_duplicate_slide_identity_fail_closed() -> TestResult {
    let source = synthetic_package(false, false)?;
    let bucket_alias = append_bucket_role_alias(&source, ROW_BUCKETS[0], TABLE_MODEL_MESSAGE_TYPE)?;
    assert_rejected_atomically(&bucket_alias, Dimension::Row(1))?;
    let duplicate_slide = append_duplicate_slide_identity(&source)?;
    assert_rejected_atomically(&duplicate_slide, Dimension::Row(1))?;
    Ok(())
}

#[test]
fn storage_and_global_archive_routes_fail_closed() -> TestResult {
    let source = synthetic_package(false, false)?;
    let duplicate_route = append_duplicate_row_bucket_route(&source, MODELS[0])?;
    assert_rejected_atomically(&duplicate_route, Dimension::Row(1))?;
    let missing_drawable_metadata = drop_unrelated_drawable_metadata(&source)?;
    assert_rejected_atomically(&missing_drawable_metadata, Dimension::Row(1))?;
    let foreign_bucket_inbound = append_foreign_bucket_inbound(&source)?;
    assert_rejected_atomically(&foreign_bucket_inbound, Dimension::Row(1))?;
    Ok(())
}

#[test]
fn noncanonical_bucket_field_and_metadata_identity_routes_fail_closed() -> TestResult {
    let source = synthetic_package(false, false)?;
    let noncanonical_field = append_noncanonical_model_bucket_field(&source, MODELS[0])?;
    assert_rejected_atomically(&noncanonical_field, Dimension::Row(1))?;

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
    })?;
    assert_rejected_atomically(&missing_model_uuid, Dimension::Row(1))?;

    let missing_bucket_uuid = rewrite_metadata(&source, |metadata| {
        if let Some(component) = metadata
            .components
            .iter_mut()
            .find(|component| component.identifier == 2)
        {
            component
                .object_uuid_map_entries
                .retain(|entry| entry.identifier != ROW_BUCKETS[0]);
        }
    })?;
    assert_rejected_atomically(&missing_bucket_uuid, Dimension::Row(1))?;

    let substituted_model_component = rewrite_metadata(&source, |metadata| {
        let model_entry = metadata
            .components
            .iter_mut()
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
    })?;
    assert_rejected_atomically(&substituted_model_component, Dimension::Row(1))?;

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
    })?;
    assert_rejected_atomically(&duplicate_uuid_pair, Dimension::Row(1))?;
    Ok(())
}

#[test]
fn two_table_z_order_selectors_are_isolated() -> TestResult {
    let source = synthetic_package(false, false)?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_table_dimension_size("Tables", 1usize, Dimension::Row(0))?,
        Size::Default
    );
    let commit = package
        .edit_slide_table_dimension_size("Tables", 0usize, Dimension::Column(2))?
        .set(points(80.0)?)
        .commit()?;
    assert_eq!(
        commit
            .package()
            .slide_table_dimension_size("Tables", 0usize, Dimension::Column(2))?,
        points(80.0)?
    );
    assert_eq!(
        commit
            .package()
            .slide_table_dimension_size("Tables", 1usize, Dimension::Column(2))?,
        Size::Default
    );
    assert_eq!(
        table_info_geometry(&exact_bytes(commit.package())?, TABLE_INFOS[1])?,
        tsp::Size {
            width: 256.0,
            height: 160.0,
        }
    );
    Ok(())
}

#[test]
fn finite_physical_and_semantic_limits_refuse_before_publication() -> TestResult {
    let source = synthetic_package(false, false)?;
    let input_limit = Limits::new(
        u64::try_from(source.len().saturating_sub(1))?,
        128,
        4 * 1024 * 1024,
        8 * 1024 * 1024,
        8 * 1024 * 1024,
    )?;
    assert!(Package::from_bytes_with_limits(&source, input_limit).is_err());

    let exact_fields = (1..=4_096)
        .find_map(|fields| {
            let archive = litchi_iwa_core::Limits::default()
                .with_header_fields(fields)
                .ok()?;
            let limits = Limits::default().with_archive_limits(archive).ok()?;
            let package = Package::from_bytes_with_limits(&source, limits).ok()?;
            package
                .slide_table_dimension_size("Tables", 0usize, Dimension::Row(1))
                .is_ok()
                .then_some(fields)
        })
        .ok_or_else(|| io::Error::other("no finite header-field boundary found"))?;
    assert!(exact_fields > 1);
    let archive = litchi_iwa_core::Limits::default().with_header_fields(exact_fields - 1)?;
    let limits = Limits::default().with_archive_limits(archive)?;
    if let Ok(package) = Package::from_bytes_with_limits(&source, limits) {
        let before = exact_bytes(&package)?;
        let result = package.slide_table_dimension_size("Tables", 0usize, Dimension::Row(1));
        assert!(matches!(
            result,
            Err(SlideTableDimensionError::LimitExceeded { .. })
        ));
        assert_eq!(exact_bytes(&package)?, before);
    }

    let semantic =
        SemanticLimits::new(1_000_000, 65_536, 1, 1_000_000, 1_000_000, 64 * 1024 * 1024)?;
    let restricted =
        Package::from_bytes_with_options(&source, ReadOptions::new(Limits::default(), semantic))?;
    let before = exact_bytes(&restricted)?;
    assert!(
        restricted
            .slide_table_dimension_size("Tables", 0usize, Dimension::Row(1))
            .is_err()
    );
    assert_eq!(exact_bytes(&restricted)?, before);
    Ok(())
}
