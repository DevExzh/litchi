//! Exact-source integration coverage for Keynote's persisted slide-table sort.
//!
//! This suite exercises only the configuration marker in
//! `TST.TableModelArchive.sort_order` (field 44).  It deliberately does not
//! invoke Keynote's physical “Sort Now” executor: row movement, formulas,
//! tiles, and view state remain compatibility-host behavior in `litchi-iwa`.

use std::error::Error as StdError;

use litchi_iwa_archive::{
    Limits,
    package::{Catalog, EntryEdit},
};
use litchi_iwa_common::wire::{
    append_length_delimited_field, append_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsd, tsk, tsp, tst, tswp};
use litchi_keynote::slide::table::sort::{ColumnIndex, Direction, Order, Rule, Scope};
use litchi_keynote::{
    Package, ReadOptions, SemanticLimits, SlideSelector, SlideTableSortCommit,
    SlideTableSortDiagnostics, SlideTableSortEdit, SlideTableSortError, SlideTableSortLimitKind,
    SlideTableSortPatch, SlideTableSortPath, TableSelector,
};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const TABLE_INFOS: [u64; 2] = [100, 101];
const MODELS: [u64; 2] = [110, 111];
const ROW_BUCKETS: [u64; 2] = [200, 210];
const COLUMN_BUCKETS: [u64; 2] = [201, 211];
const DATA_OBJECTS: [[u64; 4]; 2] = [[220, 221, 222, 223], [230, 231, 232, 233]];
const NON_TABLE_DRAWABLE: u64 = 130;
const TITLE_STYLE: u64 = 120;
const SHAPE_STYLE: u64 = 121;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const TABLE_STYLE_MESSAGE_TYPE: u32 = 6_003;
const TABLE_STYLE_PRESET_MESSAGE_TYPE: u32 = 6_008;
const TABLE_STYLE_NETWORK_MESSAGE_TYPE: u32 = 6_247;
const STYLESHEET_MESSAGE_TYPE: u32 = 401;
const TABLE_ROLE_MESSAGE_TYPES: [u32; 7] = [
    TABLE_INFO_MESSAGE_TYPE,
    TABLE_MODEL_MESSAGE_TYPE,
    HEADER_BUCKET_MESSAGE_TYPE,
    TABLE_STYLE_MESSAGE_TYPE,
    TABLE_STYLE_PRESET_MESSAGE_TYPE,
    TABLE_STYLE_NETWORK_MESSAGE_TYPE,
    STYLESHEET_MESSAGE_TYPE,
];
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const METADATA_OBJECT: u64 = 900;
const SORT_ORDER_FIELD: u32 = 44;
const SORT_TRACKER_FIELD: u32 = 45;
const SORT_TYPE_FIELD: u32 = 1;
const SORT_RULES_FIELD: u32 = 2;
const RULE_COLUMN_FIELD: u32 = 1;
const RULE_DIRECTION_FIELD: u32 = 2;
const UNKNOWN_MODEL_FIELD: u32 = 99;
const UNKNOWN_MODEL_VALUE: u64 = 0xfeed_beef;
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..tsp::Reference::default()
    }
}

fn field_reference(path: impl Into<FieldPath>, references: &[u64]) -> FieldInfo {
    let mut field = FieldInfo::new(path);
    field.r#type = Some(FieldType::ObjectReference);
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

fn data_reference_object(
    identifier: u64,
    type_: u32,
    data: Vec<u8>,
    reference: u64,
) -> TestResult<ArchiveObject> {
    let mut object = object(identifier, type_, data, &[])?;
    object.archive_info.message_infos[0]
        .data_references
        .push(reference);
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

fn sort_payload(scope: Scope, rules: &[Rule], with_unknowns: bool) -> TestResult<Vec<u8>> {
    let mut payload = tst::TableSortOrderArchive {
        r#type: scope.native_value(),
        rules: rules
            .iter()
            .map(|rule| tst::table_sort_order_archive::SortRuleArchive {
                index: rule.column().native_value(),
                direction: rule.direction().native_value(),
            })
            .collect(),
    }
    .encode_to_vec();
    if with_unknowns {
        append_overlong_varint_field(&mut payload, 90, 0);
        let rules = repeated_length_delimited_payloads(&payload, SORT_RULES_FIELD)?;
        if let Some(first) = rules.first() {
            let mut first = first.to_vec();
            append_overlong_varint_field(&mut first, 92, 0);
            append_balanced_unknown_group(&mut first, 93);
            let mut rewritten =
                rewrite_repeated_length_delimited_fields(&payload, SORT_RULES_FIELD, &[])?;
            for (index, rule) in rules.into_iter().enumerate() {
                append_length_delimited_field(
                    &mut rewritten,
                    SORT_RULES_FIELD,
                    if index == 0 { first.as_slice() } else { rule },
                )?;
            }
            payload = rewritten;
        }
    }
    Ok(payload)
}

fn table_model(
    index: usize,
    name: &str,
    sort: Option<&[u8]>,
    with_unknowns: bool,
) -> TestResult<Vec<u8>> {
    let mut payload = tst::TableModelArchive {
        table_id: format!("table-{name}"),
        table_name: name.to_owned(),
        table_style: reference(TITLE_STYLE),
        body_text_style: reference(TITLE_STYLE),
        header_row_text_style: reference(TITLE_STYLE),
        header_column_text_style: reference(TITLE_STYLE),
        footer_row_text_style: reference(TITLE_STYLE),
        body_cell_style: reference(TITLE_STYLE),
        header_row_style: reference(TITLE_STYLE),
        header_column_style: reference(TITLE_STYLE),
        footer_row_style: reference(TITLE_STYLE),
        table_name_style: Some(reference(TITLE_STYLE)),
        table_name_shape_style: Some(reference(SHAPE_STYLE)),
        base_data_store: data_store(index),
        number_of_rows: 4,
        number_of_columns: 3,
        default_row_height: 20.0,
        default_column_width: 64.0,
        ..tst::TableModelArchive::default()
    }
    .encode_to_vec();
    if with_unknowns {
        append_varint_field(&mut payload, UNKNOWN_MODEL_FIELD, UNKNOWN_MODEL_VALUE)?;
    }
    if let Some(sort) = sort {
        append_length_delimited_field(&mut payload, SORT_ORDER_FIELD, sort)?;
    }
    Ok(payload)
}

fn synthetic_package(
    sort: Option<&[u8]>,
    with_unknowns: bool,
    with_tracker: bool,
) -> TestResult<Vec<u8>> {
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
    // The non-table drawable is intentionally present between the two table
    // drawables.  TableSelector positions count table objects, not every
    // z-order drawable.
    let slide_drawables = [TABLE_INFOS[0], NON_TABLE_DRAWABLE, TABLE_INFOS[1]];
    let slide = kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: slide_drawables.iter().copied().map(reference).collect(),
        drawables_z_order: slide_drawables.iter().copied().map(reference).collect(),
        name: Some("Tables".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    };
    let mut slide_object = object(SLIDE, 5, slide.encode_to_vec(), &slide_drawables)?;
    slide_object.archive_info.message_infos[0]
        .field_infos
        .extend([
            field_reference(vec![7], &slide_drawables),
            field_reference(vec![42], &slide_drawables),
        ]);
    let mut objects = vec![
        object(1, 1, document.encode_to_vec(), &[2])?,
        object(2, 2, show.encode_to_vec(), &[SLIDE_NODE, 80, 81])?,
        object(SLIDE_NODE, 4, node.encode_to_vec(), &[SLIDE])?,
        slide_object,
        object(
            NON_TABLE_DRAWABLE,
            SHAPE_INFO_MESSAGE_TYPE,
            shape_info_payload(),
            &[SLIDE],
        )?,
    ];
    for index in 0..TABLE_INFOS.len() {
        let info = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                parent: Some(reference(SLIDE)),
                locked: Some(false),
                ..tsd::DrawableArchive::default()
            },
            table_model: reference(MODELS[index]),
            ..tst::TableInfoArchive::default()
        };
        let mut info_object = object(
            TABLE_INFOS[index],
            TABLE_INFO_MESSAGE_TYPE,
            info.encode_to_vec(),
            // Native Keynote records the drawable parent in the payload but
            // treats it as a weak/rooting route; only the TableModel is a
            // strong MessageInfo aggregate edge.
            &[MODELS[index]],
        )?;
        info_object.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![2], &[MODELS[index]]));
        objects.push(info_object);

        let storage = storage_ids(index);
        let mut model_object = object(
            MODELS[index],
            TABLE_MODEL_MESSAGE_TYPE,
            table_model(
                index,
                if index == 0 { "Revenue" } else { "Costs" },
                sort,
                with_unknowns,
            )?,
            &storage,
        )?;
        model_object.archive_info.message_infos[0]
            .field_infos
            .extend([
                field_reference(vec![4, 1, 2], &[storage[0]]),
                field_reference(vec![4, 2], &[storage[1]]),
                field_reference(vec![4, 4], &[storage[2]]),
                field_reference(vec![4, 5], &[storage[3]]),
                field_reference(vec![4, 6], &[storage[4]]),
                field_reference(vec![4, 11], &[storage[5]]),
            ]);
        if with_tracker {
            let tracker = vec![0x0a, 0x02, 0x08, 0x01];
            let data =
                append_field_bytes(&model_object.messages[0].data, SORT_TRACKER_FIELD, &tracker)?;
            model_object.messages[0].data = data;
            model_object.archive_info.message_infos[0].length =
                model_object.messages[0].data.len().try_into()?;
        }
        objects.push(model_object);
        objects.push(object(
            ROW_BUCKETS[index],
            HEADER_BUCKET_MESSAGE_TYPE,
            vec![0x08, 0x01],
            &[],
        )?);
        objects.push(object(
            COLUMN_BUCKETS[index],
            HEADER_BUCKET_MESSAGE_TYPE,
            vec![0x08, 0x01],
            &[],
        )?);
        for identifier in DATA_OBJECTS[index] {
            objects.push(object(identifier, 7_000, vec![0x08, 0x01], &[])?);
        }
    }
    objects.push(object(TITLE_STYLE, 2_022, paragraph_style_payload(), &[])?);
    objects.push(object(SHAPE_STYLE, 2_025, shape_style_payload(), &[])?);
    let compressed = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    let metadata = SnappyStream::compress(
        &Archive {
            objects: vec![object(METADATA_OBJECT, 11_006, metadata_payload()?, &[])?],
        }
        .to_bytes()?,
    )?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("Data/sentinel.bin", b"untouched ZIP sentinel".as_slice()),
            (DOCUMENT_MEMBER, compressed.as_slice()),
            (METADATA_MEMBER, metadata.as_slice()),
            (PREVIEWS[0], b"large preview".as_slice()),
            (PREVIEWS[1], b"micro preview".as_slice()),
            (PREVIEWS[2], b"web preview".as_slice()),
        ],
        Limits::default(),
    )?)
}

fn paragraph_style_payload() -> Vec<u8> {
    tswp::ParagraphStyleArchive {
        super_: tss_style("slide-table-sort"),
        ..tswp::ParagraphStyleArchive::default()
    }
    .encode_to_vec()
}

fn tss_style(identifier: &str) -> litchi_iwa_protos::tss::StyleArchive {
    litchi_iwa_protos::tss::StyleArchive {
        style_identifier: Some(identifier.to_owned()),
        ..litchi_iwa_protos::tss::StyleArchive::default()
    }
}

fn shape_style_payload() -> Vec<u8> {
    tswp::ShapeStyleArchive {
        super_: tsd::ShapeStyleArchive {
            super_: tss_style("slide-table-sort-shape"),
            ..tsd::ShapeStyleArchive::default()
        },
        ..tswp::ShapeStyleArchive::default()
    }
    .encode_to_vec()
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

fn metadata_payload() -> TestResult<Vec<u8>> {
    let identifiers = [
        1,
        2,
        SLIDE_NODE,
        SLIDE,
        TABLE_INFOS[0],
        TABLE_INFOS[1],
        MODELS[0],
        MODELS[1],
        NON_TABLE_DRAWABLE,
        TITLE_STYLE,
        SHAPE_STYLE,
        ROW_BUCKETS[0],
        ROW_BUCKETS[1],
        COLUMN_BUCKETS[0],
        COLUMN_BUCKETS[1],
        DATA_OBJECTS[0][0],
        DATA_OBJECTS[0][1],
        DATA_OBJECTS[0][2],
        DATA_OBJECTS[0][3],
        DATA_OBJECTS[1][0],
        DATA_OBJECTS[1][1],
        DATA_OBJECTS[1][2],
        DATA_OBJECTS[1][3],
    ];
    let component = tsp::ComponentInfo {
        identifier: 1,
        preferred_locator: "Document".to_owned(),
        locator: Some("Document".to_owned()),
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
    .encode_to_vec();
    let mut payload = Vec::new();
    append_varint_field(&mut payload, 1, 900)?;
    append_length_delimited_field(&mut payload, 3, &component)?;
    Ok(payload)
}

fn append_field_bytes(source: &[u8], field: u32, payload: &[u8]) -> TestResult<Vec<u8>> {
    let mut output = source.to_vec();
    append_length_delimited_field(&mut output, field, payload)?;
    Ok(output)
}

fn document_archive(package: &[u8]) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    Ok(Archive::parse(
        SnappyStream::decompress(entry.data())?.as_bytes(),
    )?)
}

fn model_payload(package: &[u8], identifier: u64) -> TestResult<Vec<u8>> {
    Ok(document_archive(package)?
        .object(identifier)
        .ok_or("missing table model")?
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or("missing table-model message")?
        .data
        .clone())
}

fn rewrite_document_archive(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document member")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(DOCUMENT_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

fn rewrite_metadata(
    package: &[u8],
    mutate: impl FnOnce(&mut tsp::PackageMetadata) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or("missing metadata member")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    let object = archive
        .object_mut(METADATA_OBJECT)
        .ok_or("missing metadata object")?;
    let message_index = object
        .messages
        .iter()
        .position(|message| message.type_ == 11_006)
        .ok_or("missing metadata message")?;
    let mut metadata =
        tsp::PackageMetadata::decode(object.messages[message_index].data.as_slice())?;
    mutate(&mut metadata)?;
    object.replace_message_preserving_header(
        message_index,
        RawMessage {
            type_: 11_006,
            data: metadata.encode_to_vec(),
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[EntryEdit::new(METADATA_MEMBER, &compressed)],
        Limits::default(),
    )?)
}

fn rewrite_model(
    package: &[u8],
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let model = archive.object_mut(MODELS[0]).ok_or("missing model")?;
        let message = model
            .messages
            .iter_mut()
            .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
            .ok_or("missing model message")?;
        let mut data = message.data.clone();
        mutate(&mut data)?;
        model.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data,
            },
        )?;
        Ok(())
    })
}

fn rewrite_slide(
    package: &[u8],
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let slide = archive.object_mut(SLIDE).ok_or("missing slide")?;
        let message = slide.messages.first().ok_or("missing slide message")?;
        let mut data = message.data.clone();
        mutate(&mut data)?;
        slide.replace_message_preserving_header(0, RawMessage { type_: 5, data })?;
        Ok(())
    })
}

fn rewrite_field_type(
    package: &[u8],
    object_identifier: u64,
    path: &[u32],
    field_type: FieldType,
) -> TestResult<Vec<u8>> {
    rewrite_document_archive(package, |archive| {
        let object = archive
            .object_mut(object_identifier)
            .ok_or("missing object")?;
        let message_info = object
            .archive_info
            .message_infos
            .first_mut()
            .ok_or("missing message info")?;
        let field = message_info
            .field_infos
            .iter_mut()
            .find(|field| field.path.as_slice() == path)
            .ok_or("missing field info")?;
        field.r#type = Some(field_type);
        Ok(())
    })
}

fn rewrite_member(package: &[u8], name: &str, data: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog.reassemble_to_bytes(&[EntryEdit::new(name, data)], Limits::default())?)
}

fn legacy_package_bytes(flat: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(flat)?;
    let inner_entries = catalog
        .iter()
        .filter(|entry| {
            std::path::Path::new(entry.name())
                .extension()
                .is_some_and(|extension| extension.eq_ignore_ascii_case("iwa"))
        })
        .map(|entry| (entry.name(), entry.data()))
        .collect::<Vec<_>>();
    let inner =
        litchi_iwa_archive::package::to_bytes(inner_entries.iter().copied(), Limits::default())?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            ("legacy.key/Index.zip", inner.as_slice()),
            (
                "legacy.key/Data/sentinel.bin",
                b"legacy Keynote sort outer sentinel".as_slice(),
            ),
        ],
        Limits::default(),
    )?)
}

fn package_with_noncanonical_document_prefix(source: &[u8]) -> TestResult<Vec<u8>> {
    let archive = document_archive(source)?;
    let canonical = archive.to_bytes()?;
    let (header_length, prefix_length) = litchi_iwa_common::decode_varint_from_bytes(&canonical)?;
    if prefix_length != 1 || header_length >= 0x80 {
        return Err("fixture expected a one-byte object prefix".into());
    }
    let mut noncanonical = Vec::new();
    noncanonical.try_reserve_exact(canonical.len() + 1)?;
    noncanonical.push(u8::try_from(header_length)? | 0x80);
    noncanonical.push(0);
    noncanonical.extend_from_slice(&canonical[prefix_length..]);
    Archive::parse(&noncanonical)?;
    rewrite_member(
        source,
        DOCUMENT_MEMBER,
        &SnappyStream::compress(&noncanonical)?,
    )
}

fn duplicate_slide_name_package() -> TestResult<Vec<u8>> {
    let source = source_without_sort()?;
    rewrite_document_archive(&source, |archive| {
        let show = archive.object_mut(2).ok_or("missing show")?;
        let message = show.messages.first_mut().ok_or("missing show message")?;
        let mut payload = kn::ShowArchive::decode(message.data.as_slice())?;
        payload.slide_tree.slides.push(reference(SLIDE_NODE));
        message.data = payload.encode_to_vec();
        let length = message.data.len().try_into()?;
        show.archive_info.message_infos[0].length = length;
        Ok(())
    })
}

fn model_with_sort(source: &[u8], sort: &[u8], tracker: Option<&[u8]>) -> TestResult<Vec<u8>> {
    rewrite_model(source, |model| {
        let mut data = rewrite_repeated_length_delimited_fields(model, SORT_ORDER_FIELD, &[])?;
        append_length_delimited_field(&mut data, SORT_ORDER_FIELD, sort)?;
        if let Some(tracker) = tracker {
            data = rewrite_repeated_length_delimited_fields(&data, SORT_TRACKER_FIELD, &[])?;
            append_length_delimited_field(&mut data, SORT_TRACKER_FIELD, tracker)?;
        }
        *model = data;
        Ok(())
    })
}

fn append_overlong_varint_field(data: &mut Vec<u8>, field: u32, value: u64) {
    push_varint(data, u64::from(field) << 3);
    if value == 0 {
        data.extend_from_slice(&[0x80, 0x00]);
    } else {
        push_varint(data, value);
    }
}

fn append_balanced_unknown_group(data: &mut Vec<u8>, field: u32) {
    push_varint(data, u64::from(field) << 3 | 3);
    push_varint(data, 1 << 3);
    push_varint(data, 1);
    push_varint(data, u64::from(field) << 3 | 4);
}

fn push_varint(data: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        data.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    data.push(value as u8);
}

fn one_rule(column: usize, direction: Direction) -> TestResult<Order> {
    Ok(Order::new([Rule::new(
        ColumnIndex::new(column)?,
        direction,
    )])?)
}

fn two_rule_order() -> TestResult<Order> {
    Ok(Order::new([
        Rule::new(ColumnIndex::new(2)?, Direction::Descending),
        Rule::new(ColumnIndex::new(1)?, Direction::Ascending),
    ])?)
}

fn exact_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn field_payloads(package: &[u8], field: u32) -> TestResult<Vec<Vec<u8>>> {
    Ok(
        repeated_length_delimited_payloads(&model_payload(package, MODELS[0])?, field)?
            .into_iter()
            .map(ToOwned::to_owned)
            .collect(),
    )
}

fn member_bytes(package: &[u8], name: &str) -> TestResult<Vec<u8>> {
    Ok(Catalog::from_bytes(package)?
        .iter()
        .find(|entry| entry.name() == name)
        .ok_or("missing package member")?
        .data()
        .to_vec())
}

fn assert_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    let changed = before
        .iter()
        .filter(|entry| {
            after
                .iter()
                .find(|candidate| candidate.name() == entry.name())
                .is_some_and(|candidate| candidate.data() != entry.data())
        })
        .map(|entry| entry.name().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(changed, vec![DOCUMENT_MEMBER.to_owned()]);
    for entry in before.iter() {
        let candidate = after
            .iter()
            .find(|value| value.name() == entry.name())
            .ok_or("candidate removed a package member")?;
        if entry.name() != DOCUMENT_MEMBER {
            assert_eq!(entry.data(), candidate.data(), "unselected member changed");
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record()
            );
        }
    }
    assert_eq!(before.len(), after.len());
    Ok(())
}

fn assert_rejected_atomically(source: &[u8]) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        return Ok(());
    };
    let before = exact_bytes(&package)?;
    assert!(
        package
            .slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0))
            .is_err()
    );
    assert!(
        package
            .edit_slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

fn assert_parsed_rejected_atomically(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source)?;
    let before = exact_bytes(&package)?;
    assert!(
        package
            .slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0))
            .is_err()
    );
    assert!(
        package
            .edit_slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0))
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

fn source_without_sort() -> TestResult<Vec<u8>> {
    synthetic_package(None, false, false)
}

fn source_with_sort(with_unknowns: bool, with_tracker: bool) -> TestResult<Vec<u8>> {
    let order = one_rule(1, Direction::Ascending)?;
    let payload = sort_payload(Scope::EntireTable, order.rules(), with_unknowns)?;
    let tracker = with_tracker.then_some(vec![0x0a, 0x02, 0x08, 0x01]);
    model_with_sort(
        &synthetic_package(None, with_unknowns, false)?,
        &payload,
        tracker.as_deref(),
    )
}

#[test]
fn transaction_values_are_typed_and_archive_free() {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}
    assert_send_sync_debug::<ColumnIndex>();
    assert_send_sync_debug::<Direction>();
    assert_send_sync_debug::<Order>();
    assert_send_sync_debug::<Rule>();
    assert_send_sync_debug::<Scope>();
    assert_send_sync_debug::<SlideTableSortEdit<'static>>();
    assert_send_sync_debug::<SlideTableSortCommit>();
    assert_send_sync_debug::<SlideTableSortPatch>();
    assert_send_sync_debug::<SlideTableSortDiagnostics>();
    assert_send_sync_debug::<SlideTableSortError>();
    assert_send_sync_debug::<SlideTableSortLimitKind>();
    assert_send_sync_debug::<SlideTableSortPath>();
}

#[test]
fn table_selector_counts_only_table_drawables_in_z_order() -> TestResult {
    let source = source_without_sort()?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(
        package.slide_table_sort_order("Tables", TableSelector::index(0))?,
        None
    );
    assert_eq!(
        package.slide_table_sort_order(SlideSelector::index(0), TableSelector::index(1))?,
        None
    );
    assert!(
        package
            .slide_table_sort_order(SlideSelector::index(0), TableSelector::index(2))
            .is_err()
    );
    assert_eq!(
        package.slide_table_sort_order("Tables", TableSelector::index(0))?,
        package.slide_table_sort_order(SlideSelector::index(0), 0usize)?
    );
    Ok(())
}

#[test]
fn aggregate_only_storage_routes_and_unrelated_model_edges_are_preserved() -> TestResult {
    let source = rewrite_document_archive(&source_without_sort()?, |archive| {
        let table = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or("missing table info")?;
        let info = &mut table.archive_info.message_infos[0];
        info.field_infos.clear();
        info.object_references.push(TITLE_STYLE);

        let model = archive.object_mut(MODELS[0]).ok_or("missing model")?;
        let info = &mut model.archive_info.message_infos[0];
        info.field_infos.clear();
        info.object_references.push(TITLE_STYLE);
        let mut unrelated = field_reference(vec![70], &[TITLE_STYLE]);
        unrelated.r#type = Some(FieldType::Message);
        info.field_infos.push(unrelated);
        archive
            .object_mut(NON_TABLE_DRAWABLE)
            .ok_or("missing unrelated drawable")?
            .archive_info
            .message_infos[0]
            .data_references
            .push(777);
        Ok(())
    })?;
    let source = rewrite_metadata(&source, |metadata| {
        let component = metadata
            .components
            .first_mut()
            .ok_or("missing metadata component")?;
        component
            .object_uuid_map_entries
            .retain(|entry| !storage_ids(0).contains(&entry.identifier));
        component.data_references.push(tsp::ComponentDataReference {
            data_identifier: 777,
            object_reference_list: vec![tsp::component_data_reference::ObjectReference {
                object_identifier: NON_TABLE_DRAWABLE,
                count: 1,
            }],
        });
        metadata.data_metadata_map = Some(reference(SHAPE_STYLE));
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;
    assert_eq!(package.slide_table_sort_order(0usize, 0usize)?, None);

    let order = one_rule(0, Direction::Ascending)?;
    let commit = package
        .edit_slide_table_sort_order(0usize, 0usize)?
        .set(order.clone())
        .commit()?;
    assert_eq!(
        commit.package().slide_table_sort_order(0usize, 0usize)?,
        Some(order)
    );
    let target = exact_bytes(commit.package())?;
    assert_locality(&source, &target)?;
    let restored = commit
        .package()
        .apply_slide_table_sort_order(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);

    let mismatched_count = rewrite_metadata(&source, |metadata| {
        metadata
            .components
            .first_mut()
            .and_then(|component| component.data_references.last_mut())
            .and_then(|reference| reference.object_reference_list.first_mut())
            .ok_or("missing unrelated data owner")?
            .count = 2;
        Ok(())
    })?;
    assert_parsed_rejected_atomically(&mismatched_count)?;

    let partial = rewrite_document_archive(&source_without_sort()?, |archive| {
        let model = archive.object_mut(MODELS[0]).ok_or("missing model")?;
        model.archive_info.message_infos[0].field_infos.pop();
        Ok(())
    })?;
    assert_parsed_rejected_atomically(&partial)
}

#[test]
fn absent_sort_clear_and_reset_are_exact_noops() -> TestResult {
    let source = source_without_sort()?;
    let package = Package::from_bytes(&source)?;
    let clear = package
        .edit_slide_table_sort_order("Tables", TableSelector::index(0))?
        .clear()
        .commit()?;
    assert!(clear.patch().is_noop());
    assert_eq!(clear.patch().before(), None);
    assert_eq!(clear.patch().after(), None);
    assert_eq!(exact_bytes(clear.package())?, source);
    assert!(!clear.diagnostics().changed());
    assert_eq!(clear.diagnostics().touched_components(), 0);
    assert_eq!(clear.diagnostics().deleted_previews(), 0);
    assert!(!clear.diagnostics().full_reparse_performed());

    let reset = package
        .edit_slide_table_sort_order(SlideSelector::index(0), TableSelector::index(0))?
        .reset()
        .commit()?;
    assert!(reset.patch().is_noop());
    assert_eq!(exact_bytes(reset.package())?, source);
    Ok(())
}

#[test]
fn legacy_source_allows_exact_noop_but_refuses_changed_sort() -> TestResult {
    let flat = source_without_sort()?;
    let legacy = legacy_package_bytes(&flat)?;
    let package = Package::from_bytes(&legacy)?;
    assert_eq!(package.slide_table_sort_order(0usize, 0usize)?, None);

    let noop = package
        .edit_slide_table_sort_order(0usize, 0usize)?
        .clear()
        .commit()?;
    assert!(noop.patch().is_noop());
    assert_eq!(exact_bytes(noop.package())?, legacy);
    let applied = package.apply_slide_table_sort_order(noop.patch())?;
    assert_eq!(exact_bytes(applied.package())?, legacy);

    let before = exact_bytes(&package)?;
    let changed = package
        .edit_slide_table_sort_order(0usize, 0usize)?
        .set(one_rule(0, Direction::Ascending)?)
        .commit();
    assert!(matches!(
        changed,
        Err(SlideTableSortError::UnsupportedSource)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn changed_sort_rejects_noncanonical_iwa_object_framing_atomically() -> TestResult {
    let malformed = package_with_noncanonical_document_prefix(&source_without_sort()?)?;
    let package = Package::from_bytes(&malformed)?;
    assert_eq!(package.slide_table_sort_order(0usize, 0usize)?, None);
    let before = exact_bytes(&package)?;
    assert!(
        package
            .edit_slide_table_sort_order(0usize, 0usize)?
            .set(one_rule(0, Direction::Ascending)?)
            .commit()
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn set_selected_rows_preserves_tracker_unknowns_metadata_previews_and_locality() -> TestResult {
    let source = source_with_sort(true, true)?;
    let package = Package::from_bytes(&source)?;
    let before = package.slide_table_sort_order("Tables", TableSelector::index(0))?;
    assert_eq!(before, Some(one_rule(1, Direction::Ascending)?));
    let after = Order::selected_rows([
        Rule::new(ColumnIndex::new(2)?, Direction::Descending),
        Rule::new(ColumnIndex::new(1)?, Direction::Ascending),
    ])?;
    let commit = package
        .edit_slide_table_sort_order("Tables", TableSelector::index(0))?
        .set(after.clone())
        .commit()?;
    assert_eq!(commit.patch().before().cloned(), before);
    assert_eq!(commit.patch().after().cloned(), Some(after.clone()));
    assert_eq!(
        commit
            .package()
            .slide_table_sort_order("Tables", TableSelector::index(0))?,
        Some(after)
    );
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert!(commit.diagnostics().full_reparse_performed());
    let target = exact_bytes(commit.package())?;
    assert_ne!(target, source);
    assert_locality(&source, &target)?;
    assert_eq!(
        model_payload(&source, MODELS[1])?,
        model_payload(&target, MODELS[1])?,
        "unselected table model changed"
    );
    assert_eq!(
        field_payloads(&source, SORT_TRACKER_FIELD)?,
        field_payloads(&target, SORT_TRACKER_FIELD)?
    );
    assert_eq!(
        member_bytes(&source, PREVIEWS[0])?,
        member_bytes(&target, PREVIEWS[0])?
    );
    assert_eq!(
        member_bytes(&source, METADATA_MEMBER)?,
        member_bytes(&target, METADATA_MEMBER)?,
        "metadata member changed"
    );
    assert!(
        field_payloads(&target, SORT_ORDER_FIELD)?[0]
            .windows(2)
            .any(|window| window == [0x80, 0x00])
    );
    assert!(
        field_payloads(&target, SORT_ORDER_FIELD)?[0]
            .windows(2)
            .any(|window| window == [0xeb, 0x05])
    );
    assert!(
        model_payload(&target, MODELS[0])?
            .windows(2)
            .any(|window| window == [0x98, 0x06])
    );
    Ok(())
}

#[test]
fn set_apply_inverse_and_conflict_are_exact_source_bound() -> TestResult {
    let source = source_without_sort()?;
    let package = Package::from_bytes(&source)?;
    let expected = two_rule_order()?;
    let commit = package
        .edit_slide_table_sort_order(0usize, 0usize)?
        .set(expected.clone())
        .commit()?;
    let target = exact_bytes(commit.package())?;
    assert_eq!(
        exact_bytes(
            &package
                .apply_slide_table_sort_order(commit.patch())?
                .into_package()
        )?,
        target
    );
    assert!(matches!(
        commit
            .package()
            .apply_slide_table_sort_order(commit.patch()),
        Err(SlideTableSortError::PatchConflict)
    ));
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.inverse(), *commit.patch());
    let restored = commit.package().apply_slide_table_sort_order(&inverse)?;
    assert_eq!(exact_bytes(restored.package())?, source);
    assert_eq!(
        restored.package().slide_table_sort_order(0usize, 0usize)?,
        None
    );
    Ok(())
}

#[test]
fn clear_existing_marker_preserves_empty_marker_unknowns_and_inverse() -> TestResult {
    let source = source_with_sort(true, true)?;
    let package = Package::from_bytes(&source)?;
    let commit = package
        .edit_slide_table_sort_order(0usize, 0usize)?
        .clear()
        .commit()?;
    assert_eq!(
        commit.package().slide_table_sort_order(0usize, 0usize)?,
        None
    );
    let sorts = field_payloads(&exact_bytes(commit.package())?, SORT_ORDER_FIELD)?;
    assert_eq!(sorts.len(), 1);
    assert!(sorts[0].windows(2).any(|window| window == [0x80, 0x00]));
    assert_eq!(
        field_payloads(&source, SORT_TRACKER_FIELD)?,
        field_payloads(&exact_bytes(commit.package())?, SORT_TRACKER_FIELD)?
    );
    let restored = commit
        .package()
        .apply_slide_table_sort_order(&commit.patch().inverse())?;
    assert_eq!(exact_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn selected_rows_scope_is_persisted_without_a_row_range() -> TestResult {
    let source = source_without_sort()?;
    let selected = Order::selected_rows([Rule::new(ColumnIndex::new(1)?, Direction::Descending)])?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_table_sort_order(0usize, 0usize)?
        .set(selected.clone())
        .commit()?;
    assert_eq!(
        commit.package().slide_table_sort_order(0usize, 0usize)?,
        Some(selected)
    );
    Ok(())
}

#[test]
fn locked_table_refuses_changed_sort_atomically() -> TestResult {
    let source = source_without_sort()?;
    let locked = rewrite_document_archive(&source, |archive| {
        let table = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or("missing table info")?;
        let info = tst::TableInfoArchive::decode(table.messages[0].data.as_slice())?;
        table.messages[0].data = tst::TableInfoArchive {
            super_: tsd::DrawableArchive {
                locked: Some(true),
                ..info.super_
            },
            ..info
        }
        .encode_to_vec();
        table.archive_info.message_infos[0].length = table.messages[0].data.len().try_into()?;
        Ok(())
    })?;
    let package = Package::from_bytes(&locked)?;
    let before = exact_bytes(&package)?;
    assert!(
        package
            .edit_slide_table_sort_order(0usize, 0usize)?
            .set(one_rule(1, Direction::Ascending)?)
            .commit()
            .is_err()
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn malformed_duplicate_wrong_wire_and_noncanonical_fields_fail_closed() -> TestResult {
    let fixture = source_without_sort()?;
    let valid = sort_payload(
        Scope::EntireTable,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        false,
    )?;
    let duplicate = rewrite_model(&fixture, |model| {
        append_length_delimited_field(model, SORT_ORDER_FIELD, &valid)?;
        append_length_delimited_field(model, SORT_ORDER_FIELD, &valid)?;
        Ok(())
    })?;
    let wrong_wire = rewrite_model(&fixture, |model| {
        append_varint_field(model, SORT_ORDER_FIELD, 1)?;
        Ok(())
    })?;
    let mut noncanonical_payload = valid.clone();
    append_overlong_varint_field(&mut noncanonical_payload, SORT_TYPE_FIELD, 0);
    let noncanonical = rewrite_model(&fixture, |model| {
        append_length_delimited_field(model, SORT_ORDER_FIELD, &noncanonical_payload)?;
        Ok(())
    })?;
    for malformed in [duplicate, wrong_wire, noncanonical] {
        assert_rejected_atomically(&malformed)?;
    }
    Ok(())
}

#[test]
fn malformed_nested_duplicate_fields_and_unknown_enum_fail_closed() -> TestResult {
    let fixture = source_without_sort()?;
    let valid = sort_payload(
        Scope::EntireTable,
        &[Rule::new(ColumnIndex::new(1)?, Direction::Ascending)],
        false,
    )?;
    let duplicate_type = rewrite_model(&fixture, |model| {
        let mut malformed = valid.clone();
        append_varint_field(&mut malformed, SORT_TYPE_FIELD, 0)?;
        append_length_delimited_field(model, SORT_ORDER_FIELD, &malformed)?;
        Ok(())
    })?;
    let duplicate_column = rewrite_model(&fixture, |model| {
        let rules = repeated_length_delimited_payloads(&valid, SORT_RULES_FIELD)?;
        let mut rule = rules[0].to_vec();
        append_varint_field(&mut rule, RULE_COLUMN_FIELD, 1)?;
        let mut malformed =
            rewrite_repeated_length_delimited_fields(&valid, SORT_RULES_FIELD, &[])?;
        append_length_delimited_field(&mut malformed, SORT_RULES_FIELD, &rule)?;
        append_length_delimited_field(model, SORT_ORDER_FIELD, &malformed)?;
        Ok(())
    })?;
    let duplicate_direction = rewrite_model(&fixture, |model| {
        let rules = repeated_length_delimited_payloads(&valid, SORT_RULES_FIELD)?;
        let mut rule = rules[0].to_vec();
        append_varint_field(&mut rule, RULE_DIRECTION_FIELD, 0)?;
        let mut malformed =
            rewrite_repeated_length_delimited_fields(&valid, SORT_RULES_FIELD, &[])?;
        append_length_delimited_field(&mut malformed, SORT_RULES_FIELD, &rule)?;
        append_length_delimited_field(model, SORT_ORDER_FIELD, &malformed)?;
        Ok(())
    })?;
    let unknown_scope = rewrite_model(&fixture, |model| {
        let malformed =
            litchi_iwa_common::wire::patch_varint_field(&valid, SORT_TYPE_FIELD, true, Some(99))?;
        append_length_delimited_field(model, SORT_ORDER_FIELD, &malformed)?;
        Ok(())
    })?;
    for malformed in [
        duplicate_type,
        duplicate_column,
        duplicate_direction,
        unknown_scope,
    ] {
        assert_rejected_atomically(&malformed)?;
    }
    Ok(())
}

#[test]
fn model_graph_alias_and_missing_routes_are_atomic() -> TestResult {
    let fixture = source_without_sort()?;
    let wrong_type = rewrite_document_archive(&fixture, |archive| {
        archive
            .object_mut(MODELS[0])
            .ok_or("missing model")?
            .messages[0]
            .type_ = TABLE_INFO_MESSAGE_TYPE;
        Ok(())
    })?;
    let duplicate_model = rewrite_document_archive(&fixture, |archive| {
        archive
            .object_mut(MODELS[1])
            .ok_or("missing model")?
            .archive_info
            .identifier = Some(MODELS[0]);
        Ok(())
    });
    let missing_model = rewrite_document_archive(&fixture, |archive| {
        let table = archive.object_mut(TABLE_INFOS[0]).ok_or("missing table")?;
        let info = tst::TableInfoArchive::decode(table.messages[0].data.as_slice())?;
        table.messages[0].data = tst::TableInfoArchive {
            table_model: reference(999_999),
            ..info
        }
        .encode_to_vec();
        table.archive_info.message_infos[0].length = table.messages[0].data.len().try_into()?;
        Ok(())
    })?;
    let duplicate_z_order = rewrite_slide(&fixture, |slide| {
        let mut duplicate = slide.clone();
        let reference_bytes = reference(TABLE_INFOS[0]).encode_to_vec();
        append_length_delimited_field(&mut duplicate, 42, &reference_bytes)?;
        *slide = duplicate;
        Ok(())
    })?;
    for malformed in [wrong_type, missing_model, duplicate_z_order] {
        assert_rejected_atomically(&malformed)?;
    }
    match duplicate_model {
        Ok(malformed) => assert_rejected_atomically(&malformed)?,
        Err(error) => assert!(
            error.to_string().contains("duplicate object identifier"),
            "unexpected duplicate-object fixture error: {error}"
        ),
    }
    Ok(())
}

#[test]
fn cross_component_model_inbound_is_rejected_without_source_mutation() -> TestResult {
    let source = source_without_sort()?;
    let catalog = Catalog::from_bytes(&source)?;
    let inbound = object(700, 7_000, vec![0x08, 0x01], &[MODELS[0]])?;
    let component = SnappyStream::compress(
        &Archive {
            objects: vec![inbound],
        }
        .to_bytes()?,
    )?;
    let mut members = catalog
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<Vec<_>>();
    members.push(("Index/Other.iwa".to_owned(), component));
    let refs = members
        .iter()
        .map(|(name, bytes)| (name.as_str(), bytes.as_slice()))
        .collect::<Vec<_>>();
    let malformed = litchi_iwa_archive::package::to_bytes(refs, Limits::default())?;
    assert_rejected_atomically(&malformed)
}

#[test]
fn cross_component_data_inbound_is_rejected_without_source_mutation() -> TestResult {
    let source = source_without_sort()?;
    let catalog = Catalog::from_bytes(&source)?;
    let inbound = data_reference_object(700, 7_000, vec![0x08, 0x01], MODELS[0])?;
    let component = SnappyStream::compress(
        &Archive {
            objects: vec![inbound],
        }
        .to_bytes()?,
    )?;
    let mut members = catalog
        .iter()
        .map(|entry| (entry.name().to_owned(), entry.data().to_vec()))
        .collect::<Vec<_>>();
    members.push(("Index/Other.iwa".to_owned(), component));
    let refs = members
        .iter()
        .map(|(name, bytes)| (name.as_str(), bytes.as_slice()))
        .collect::<Vec<_>>();
    let malformed = litchi_iwa_archive::package::to_bytes(refs, Limits::default())?;
    assert_rejected_atomically(&malformed)
}

#[test]
fn selected_archive_info_data_references_are_rejected_atomically() -> TestResult {
    let fixture = source_without_sort()?;
    let slide_aggregate = rewrite_document_archive(&fixture, |archive| {
        let slide = archive.object_mut(SLIDE).ok_or("missing slide")?;
        slide.archive_info.message_infos[0]
            .data_references
            .push(TABLE_INFOS[0]);
        Ok(())
    })?;
    let slide_field = rewrite_document_archive(&fixture, |archive| {
        let slide = archive.object_mut(SLIDE).ok_or("missing slide")?;
        let owned_field = slide.archive_info.message_infos[0]
            .field_infos
            .iter_mut()
            .find(|field| field.path.as_slice() == [7])
            .ok_or("missing slide-owned field")?;
        owned_field.data_references.push(TABLE_INFOS[0]);
        Ok(())
    })?;
    let table_aggregate = rewrite_document_archive(&fixture, |archive| {
        let table = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or("missing table info")?;
        table.archive_info.message_infos[0]
            .data_references
            .push(MODELS[0]);
        Ok(())
    })?;
    let table_field = rewrite_document_archive(&fixture, |archive| {
        let table = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or("missing table info")?;
        let model_field = table.archive_info.message_infos[0]
            .field_infos
            .iter_mut()
            .find(|field| field.path.as_slice() == [2])
            .ok_or("missing table-model field")?;
        model_field.data_references.push(MODELS[0]);
        Ok(())
    })?;
    for malformed in [slide_aggregate, slide_field, table_aggregate, table_field] {
        assert_rejected_atomically(&malformed)?;
    }
    Ok(())
}

#[test]
fn noncanonical_model_reference_path_is_rejected_atomically() -> TestResult {
    let fixture = source_without_sort()?;
    let malformed = rewrite_document_archive(&fixture, |archive| {
        let table = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or("missing table info")?;
        table.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![99], &[MODELS[0]]));
        Ok(())
    })?;
    assert_rejected_atomically(&malformed)
}

#[test]
fn all_known_table_role_aliases_are_rejected_atomically() -> TestResult {
    let fixture = source_without_sort()?;
    for role_type in TABLE_ROLE_MESSAGE_TYPES {
        let malformed_model = rewrite_document_archive(&fixture, |archive| {
            let model = archive.object_mut(MODELS[0]).ok_or("missing model")?;
            model.push_message(RawMessage {
                type_: role_type,
                data: vec![0x08, 0x01],
            })?;
            Ok(())
        })?;
        assert_rejected_atomically(&malformed_model)?;

        let malformed_info = rewrite_document_archive(&fixture, |archive| {
            let table = archive
                .object_mut(TABLE_INFOS[0])
                .ok_or("missing table info")?;
            table.push_message(RawMessage {
                type_: role_type,
                data: vec![0x08, 0x01],
            })?;
            Ok(())
        })?;
        assert_rejected_atomically(&malformed_info)?;
    }
    Ok(())
}

#[test]
fn role_alias_on_non_table_drawable_is_rejected_atomically() -> TestResult {
    let fixture = source_without_sort()?;
    let malformed = rewrite_document_archive(&fixture, |archive| {
        let drawable = archive
            .object_mut(NON_TABLE_DRAWABLE)
            .ok_or("missing non-table drawable")?;
        drawable.push_message(RawMessage {
            type_: TABLE_STYLE_MESSAGE_TYPE,
            data: vec![0x08, 0x01],
        })?;
        Ok(())
    })?;
    assert_rejected_atomically(&malformed)
}

#[test]
fn column_order_values_reject_empty_duplicates_and_out_of_range_native_indices() -> TestResult {
    let column = ColumnIndex::new(1)?;
    let rule = Rule::new(column, Direction::Ascending);
    assert!(matches!(
        Order::new([]),
        Err(litchi_keynote::slide::table::sort::Error::EmptyOrder)
    ));
    assert!(matches!(
        Order::new([rule, Rule::new(column, Direction::Descending)]),
        Err(litchi_keynote::slide::table::sort::Error::DuplicateColumn { column: 1 })
    ));
    if let Ok(index) = usize::try_from(u64::from(u32::MAX) + 1) {
        assert!(ColumnIndex::new(index).is_err());
    }
    assert!(litchi_keynote::slide::table::sort::RowRange::new(2, 2).is_err());
    Ok(())
}

#[test]
fn selector_and_patch_debug_do_not_leak_native_names_or_bytes() -> TestResult {
    let package = Package::from_bytes(&source_without_sort()?)?;
    let edit = package.edit_slide_table_sort_order("Tables", TableSelector::index(0))?;
    let debug = format!("{edit:?}");
    assert!(!debug.contains("Index/"));
    assert!(!debug.contains("Document.iwa"));
    let commit = edit.set(one_rule(1, Direction::Ascending)?).commit()?;
    let patch_debug = format!("{:?}", commit.patch());
    assert!(!patch_debug.contains("Index/"));
    assert!(!patch_debug.contains("Document.iwa"));
    Ok(())
}

#[test]
fn selected_object_reference_routes_require_object_reference_field_types() -> TestResult {
    let fixture = source_without_sort()?;
    let routes = [
        (SLIDE, [7_u32].as_slice()),
        (SLIDE, [42_u32].as_slice()),
        (TABLE_INFOS[0], [2_u32].as_slice()),
    ];
    for (object_identifier, path) in routes {
        let malformed = rewrite_field_type(&fixture, object_identifier, path, FieldType::Value)?;
        assert_parsed_rejected_atomically(&malformed)?;
    }
    Ok(())
}

#[test]
fn extra_foreign_and_malformed_reference_routes_fail_atomically() -> TestResult {
    let fixture = source_without_sort()?;
    let duplicate_slide_aggregate = rewrite_document_archive(&fixture, |archive| {
        let slide = archive.object_mut(SLIDE).ok_or("missing slide")?;
        slide.archive_info.message_infos[0]
            .object_references
            .push(TABLE_INFOS[0]);
        Ok(())
    })?;
    let duplicate_table_aggregate = rewrite_document_archive(&fixture, |archive| {
        let table = archive
            .object_mut(TABLE_INFOS[0])
            .ok_or("missing table info")?;
        table.archive_info.message_infos[0]
            .object_references
            .push(MODELS[0]);
        Ok(())
    })?;
    let foreign_aggregate = rewrite_document_archive(&fixture, |archive| {
        let style = archive
            .object_mut(TITLE_STYLE)
            .ok_or("missing title style")?;
        style.archive_info.message_infos[0]
            .object_references
            .push(MODELS[0]);
        Ok(())
    })?;
    let foreign_field = rewrite_document_archive(&fixture, |archive| {
        let style = archive
            .object_mut(TITLE_STYLE)
            .ok_or("missing title style")?;
        style.archive_info.message_infos[0]
            .field_infos
            .push(field_reference(vec![77], &[MODELS[0]]));
        Ok(())
    })?;
    let malformed_outer_key = rewrite_slide(&fixture, |slide| {
        let references = repeated_length_delimited_payloads(slide, 42)?;
        let mut rewritten = rewrite_repeated_length_delimited_fields(slide, 42, &[])?;
        let first = references.first().ok_or("missing z-order reference")?;
        // Field 42's canonical key is d2 02; this overlong key decodes to
        // the same field number but is intentionally non-canonical.
        rewritten.extend_from_slice(&[0xd2, 0x82, 0x00]);
        push_varint(&mut rewritten, first.len() as u64);
        rewritten.extend_from_slice(first);
        for reference in references.into_iter().skip(1) {
            append_length_delimited_field(&mut rewritten, 42, reference)?;
        }
        *slide = rewritten;
        Ok(())
    })?;
    let malformed_nested_reference = rewrite_slide(&fixture, |slide| {
        let references = repeated_length_delimited_payloads(slide, 42)?;
        let mut rewritten = rewrite_repeated_length_delimited_fields(slide, 42, &[])?;
        if references.is_empty() {
            return Err("missing z-order reference".into());
        }
        // The identifier 100 is encoded overlong inside an otherwise valid
        // Reference message (e4 00 instead of the canonical 64).
        append_length_delimited_field(&mut rewritten, 42, &[0x08, 0xe4, 0x00])?;
        for reference in references.into_iter().skip(1) {
            append_length_delimited_field(&mut rewritten, 42, reference)?;
        }
        *slide = rewritten;
        Ok(())
    })?;
    let wrong_reference_wire = rewrite_slide(&fixture, |slide| {
        append_varint_field(slide, 42, TABLE_INFOS[0])?;
        Ok(())
    })?;
    for malformed in [
        duplicate_slide_aggregate,
        duplicate_table_aggregate,
        foreign_aggregate,
        foreign_field,
        malformed_outer_key,
        malformed_nested_reference,
        wrong_reference_wire,
    ] {
        assert_parsed_rejected_atomically(&malformed)?;
    }
    Ok(())
}

#[test]
fn duplicate_slide_name_is_ambiguous_and_atomic() -> TestResult {
    let source = duplicate_slide_name_package()?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.slide_table_sort_order("Tables", TableSelector::index(0)),
        Err(SlideTableSortError::AmbiguousSelector)
    ));
    assert!(matches!(
        package.edit_slide_table_sort_order("Tables", TableSelector::index(0)),
        Err(SlideTableSortError::AmbiguousSelector)
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn missing_empty_and_out_of_range_selectors_are_typed_and_atomic() -> TestResult {
    let source = source_without_sort()?;
    let package = Package::from_bytes(&source)?;
    let before = exact_bytes(&package)?;
    assert!(matches!(
        package.slide_table_sort_order(SlideSelector::name(""), TableSelector::index(0)),
        Err(SlideTableSortError::EmptySlideName)
    ));
    assert!(matches!(
        package.slide_table_sort_order(SlideSelector::name("missing"), TableSelector::index(0)),
        Err(SlideTableSortError::SlideNameNotFound)
    ));
    assert!(matches!(
        package.slide_table_sort_order(SlideSelector::index(9), TableSelector::index(0)),
        Err(SlideTableSortError::SlidePositionNotFound { position }) if position.get() == 9
    ));
    assert!(matches!(
        package.slide_table_sort_order(SlideSelector::index(0), TableSelector::index(9)),
        Err(SlideTableSortError::TablePositionNotFound { position }) if position.get() == 9
    ));
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}

#[test]
fn stale_and_foreign_patches_conflict_without_mutating_the_source() -> TestResult {
    let source = source_without_sort()?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_table_sort_order(0usize, 0usize)?
        .set(two_rule_order()?)
        .commit()?;
    let patch = commit.patch();

    let stale_source = rewrite_member(&source, "Data/sentinel.bin", b"stale sentinel")?;
    let stale = Package::from_bytes(&stale_source)?;
    let stale_before = exact_bytes(&stale)?;
    assert!(matches!(
        stale.apply_slide_table_sort_order(patch),
        Err(SlideTableSortError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&stale)?, stale_before);

    let foreign_source = source_with_sort(false, false)?;
    let foreign = Package::from_bytes(&foreign_source)?;
    let foreign_before = exact_bytes(&foreign)?;
    assert!(matches!(
        foreign.apply_slide_table_sort_order(patch),
        Err(SlideTableSortError::PatchConflict)
    ));
    assert_eq!(exact_bytes(&foreign)?, foreign_before);

    let target_before = exact_bytes(commit.package())?;
    assert!(matches!(
        commit.package().apply_slide_table_sort_order(patch),
        Err(SlideTableSortError::PatchConflict)
    ));
    assert_eq!(exact_bytes(commit.package())?, target_before);
    Ok(())
}

#[test]
fn existing_sort_noop_preserves_exact_fingerprints_and_source_bytes() -> TestResult {
    let source = source_with_sort(true, true)?;
    let package = Package::from_bytes(&source)?;
    let order = package
        .slide_table_sort_order(0usize, 0usize)?
        .ok_or("fixture sort marker missing")?;
    let commit = package
        .edit_slide_table_sort_order(0usize, 0usize)?
        .set(order)
        .commit()?;
    assert!(commit.patch().is_noop());
    assert_eq!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert_eq!(exact_bytes(commit.package())?, source);
    assert!(!commit.diagnostics().changed());
    let applied = package.apply_slide_table_sort_order(commit.patch())?;
    assert!(applied.patch().is_noop());
    assert_eq!(exact_bytes(applied.package())?, source);
    assert!(!applied.diagnostics().changed());
    Ok(())
}

#[test]
fn changing_table_position_one_is_local_and_does_not_touch_table_zero() -> TestResult {
    let source = source_without_sort()?;
    let order = one_rule(0, Direction::Descending)?;
    let commit = Package::from_bytes(&source)?
        .edit_slide_table_sort_order(0usize, TableSelector::index(1))?
        .set(order.clone())
        .commit()?;
    assert_eq!(
        commit
            .package()
            .slide_table_sort_order(0usize, TableSelector::index(1))?,
        Some(order)
    );
    assert_eq!(
        commit
            .package()
            .slide_table_sort_order(0usize, TableSelector::index(0))?,
        None
    );
    let target = exact_bytes(commit.package())?;
    assert_locality(&source, &target)?;
    assert_eq!(
        model_payload(&source, MODELS[0])?,
        model_payload(&target, MODELS[0])?,
        "unselected table model changed"
    );
    assert_ne!(
        model_payload(&source, MODELS[1])?,
        model_payload(&target, MODELS[1])?,
        "selected table model did not change"
    );
    Ok(())
}

#[test]
fn tight_publication_limit_rejects_changed_sort_atomically() -> TestResult {
    let source = source_without_sort()?;
    let order = two_rule_order()?;
    let baseline = Package::from_bytes(&source)?
        .edit_slide_table_sort_order(0usize, 0usize)?
        .set(order.clone())
        .commit()?;
    let target = exact_bytes(baseline.package())?;
    assert!(target.len() > source.len());
    let defaults = Limits::default();
    let limits = Limits::new(
        u64::try_from(target.len() - 1)?,
        defaults.max_entries(),
        defaults.max_entry_bytes(),
        defaults.max_total_bytes(),
        defaults.max_iwa_stream_bytes(),
    )?;
    let package = Package::from_bytes_with_options(
        &source,
        ReadOptions::new(limits, SemanticLimits::default()),
    )?;
    let before = exact_bytes(&package)?;
    let result = package
        .edit_slide_table_sort_order(0usize, 0usize)
        .and_then(|edit| edit.set(order).commit());
    assert!(
        matches!(result, Err(SlideTableSortError::LimitExceeded { .. })),
        "unexpected tight-publication result: {result:?}"
    );
    assert_eq!(exact_bytes(&package)?, before);
    Ok(())
}
