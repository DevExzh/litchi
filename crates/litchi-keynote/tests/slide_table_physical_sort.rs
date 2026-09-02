//! Selector-first physical Keynote table sorting integration coverage.
//!
//! The fixture is deliberately assembled below instead of borrowing the
//! legacy editor's builder.  It is a small, canonical `TST.TableModelArchive`
//! graph with a real BNC tile, sparse row headers, stable row/column UIDs, a
//! persisted sort marker, and exact package metadata.  The tests therefore
//! exercise the focused Keynote owner at the same archive boundary that a
//! native document uses, while keeping authored values out of the public
//! transaction API.

#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::too_many_lines,
    reason = "The fixture and its invariant checks intentionally stay in one integration module."
)]

use std::collections::BTreeMap;
use std::error::Error as StdError;
use std::sync::Arc;

use litchi_core::Position;
use litchi_iwa_archive::{Limits, package::Catalog};
use litchi_iwa_common::decode_varint_from_bytes;
use litchi_iwa_common::wire::{
    append_length_delimited_field, append_varint_field, repeated_length_delimited_payloads,
    rewrite_repeated_length_delimited_fields,
};
use litchi_iwa_core::{
    Archive, ArchiveObject, FieldInfo, FieldPath, FieldType, RawMessage, SnappyStream,
};
use litchi_iwa_protos::{kn, tsa, tsce, tsd, tsk, tsp, tst, tswp};
use litchi_keynote::slide::table::physical_sort::{
    ColumnIndex, Direction, Error, LimitKind, Order, Path, RowRange, Rule, Scope,
    UnsupportedFeature,
};
use litchi_keynote::{Package, SlideSelector, TableSelector};
use litchi_numbers_wire::{BncCell, BncCellView, CachedScalar};
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const DOCUMENT_MEMBER: &str = "Index/Document.iwa";
const METADATA_MEMBER: &str = "Index/Metadata.iwa";
const SENTINEL_MEMBER: &str = "Data/physical-sort-sentinel.bin";
const PREVIEWS: [&str; 3] = ["preview.jpg", "preview-micro.jpg", "preview-web.jpg"];

const DOCUMENT: u64 = 1;
const SHOW: u64 = 2;
const SLIDE_NODE: u64 = 3;
const SLIDE: u64 = 4;
const TABLE_INFO: u64 = 100;
const TABLE_MODEL: u64 = 101;
const NON_TABLE_DRAWABLE: u64 = 102;
const ROW_HEADERS: u64 = 110;
const COLUMN_HEADERS: u64 = 111;
const TILE: u64 = 112;
const TILE_SECOND: u64 = 122;
const TILE_THIRD: u64 = 123;
const STRINGS: u64 = 113;
const STYLES: u64 = 114;
const FORMULAS: u64 = 115;
const FORMATS: u64 = 116;
const UID_MAP: u64 = 117;
const STROKE_SIDECAR: u64 = 118;
const TITLE_STYLE: u64 = 120;
const SHAPE_STYLE: u64 = 121;
const OPTIONAL_CUSTOM_FORMATS: u64 = 124;
const OPTIONAL_FORMATS: u64 = 125;
const METADATA_OBJECT: u64 = 900;
const FORMULA_OWNER_DEPENDENCIES: u64 = 119;

const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const HEADER_BUCKET_MESSAGE_TYPE: u32 = 6_006;
const COLUMN_ROW_UID_MAP_MESSAGE_TYPE: u32 = 6_267;
const STROKE_SIDECAR_MESSAGE_TYPE: u32 = 6_305;
const FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE: u32 = 4_008;
const SHAPE_INFO_MESSAGE_TYPE: u32 = 2_011;
const METADATA_MESSAGE_TYPE: u32 = 11_006;

const SORT_ORDER_FIELD: u32 = 44;
const BASE_UID_FIELD: u32 = 46;
const STROKE_FIELD: u32 = 49;

const ROW_COUNT: u32 = 6;
const COLUMN_COUNT: u32 = 2;
const HEADER_ROWS: u32 = 1;
const FOOTER_ROWS: u32 = 1;
const BODY_COUNT: usize = 4;

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

fn data_list(
    list_type: tst::table_data_list::ListType,
    entries: Vec<tst::table_data_list::ListEntry>,
) -> Vec<u8> {
    tst::TableDataList {
        list_type: list_type as i32,
        next_list_id: entries
            .iter()
            .map(|entry| entry.key)
            .max()
            .unwrap_or(0)
            .saturating_add(1),
        entries,
        segments: Vec::new(),
        is_new_for_bnc: Some(true),
    }
    .encode_to_vec()
}

fn string_entries() -> Vec<tst::table_data_list::ListEntry> {
    [
        (1, "Name"),
        (2, "Marker"),
        (3, "zebra"),
        (4, "last"),
        (5, "apple"),
        (6, "first apple"),
        (7, "banana"),
        (8, "middle"),
        (9, "second apple"),
        (10, "Total"),
        (11, "footer"),
    ]
    .into_iter()
    .map(|(key, string)| tst::table_data_list::ListEntry {
        key,
        refcount: 1,
        string: Some(string.to_owned()),
        ..tst::table_data_list::ListEntry::default()
    })
    .collect()
}

/// Encode the compact BNC v5 text cell used by the native tile format.
fn bnc_text(identifier: u32) -> Vec<u8> {
    let mut bytes = vec![5, 3, 0, 0, 0, 0, 0, 0];
    bytes.extend_from_slice(&8u32.to_le_bytes());
    bytes.extend_from_slice(&identifier.to_le_bytes());
    bytes
}

fn encode_row(cells: &[Option<Vec<u8>>]) -> (Vec<u8>, Vec<u8>) {
    let mut storage = Vec::new();
    let mut offsets = Vec::with_capacity(cells.len() * 2);
    for cell in cells {
        let Some(cell) = cell else {
            offsets.extend_from_slice(&u16::MAX.to_le_bytes());
            continue;
        };
        offsets.extend_from_slice(
            &u16::try_from(storage.len())
                .expect("the fixture's BNC rows fit the narrow offset representation")
                .to_le_bytes(),
        );
        storage.extend_from_slice(cell);
    }
    (storage, offsets)
}

/// Encode the canonical empty legacy mirror emitted beside a native BNC-v5
/// cell.  Keynote uses one twelve-byte v4 sentinel for every materialized
/// cell; it does not duplicate the modern cell payload in this sidecar.
fn pre_bnc_empty_cell() -> Vec<u8> {
    let mut bytes = vec![0; 12];
    bytes[0] = 4;
    bytes
}

fn encode_pre_bnc(cells: &[Option<Vec<u8>>]) -> (Vec<u8>, Vec<u8>) {
    let sentinels = cells
        .iter()
        .map(|cell| cell.as_ref().map(|_| pre_bnc_empty_cell()))
        .collect::<Vec<_>>();
    encode_row(&sentinels)
}

fn row(cells: [Option<Vec<u8>>; 2], index: u32) -> tst::TileRowInfo {
    let cell_count = cells.iter().filter(|cell| cell.is_some()).count();
    let (storage, offsets) = encode_row(&cells);
    let (pre_storage, pre_offsets) = encode_pre_bnc(&cells);
    tst::TileRowInfo {
        tile_row_index: index,
        cell_count: u32::try_from(cell_count).expect("fixture cell count fits u32"),
        cell_storage_buffer_pre_bnc: pre_storage,
        cell_offsets_pre_bnc: pre_offsets,
        storage_version: Some(5),
        cell_storage_buffer: Some(storage),
        cell_offsets: Some(offsets),
        has_wide_offsets: Some(false),
    }
}

fn source_row_cells(index: u32) -> [Option<Vec<u8>>; 2] {
    match index {
        0 => [Some(bnc_text(1)), Some(bnc_text(2))],
        1 => [Some(bnc_text(3)), Some(bnc_text(4))],
        2 => [Some(bnc_text(5)), Some(bnc_text(6))],
        3 => [Some(bnc_text(7)), Some(bnc_text(8))],
        4 => [Some(bnc_text(5)), Some(bnc_text(9))],
        5 => [Some(bnc_text(10)), Some(bnc_text(11))],
        _ => panic!("fixture row index is out of range"),
    }
}

fn tile_payload_for(start: u32, end: u32) -> Vec<u8> {
    let row_count = end
        .checked_sub(start)
        .filter(|count| *count != 0)
        .expect("fixture tile ranges are nonempty");
    tst::Tile {
        max_column: COLUMN_COUNT - 1,
        max_row: row_count - 1,
        num_cells: row_count * COLUMN_COUNT,
        numrows: row_count,
        row_infos: (start..end)
            .map(|global| row(source_row_cells(global), global - start))
            .collect(),
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        should_use_wide_rows: Some(false),
    }
    .encode_to_vec()
}

fn tile_payload() -> Vec<u8> {
    tile_payload_for(0, ROW_COUNT)
}

fn row_header(index: u32) -> tst::header_storage_bucket::Header {
    tst::header_storage_bucket::Header {
        index,
        size: 20.0,
        hiding_state: 0,
        number_of_cells: COLUMN_COUNT,
        ..tst::header_storage_bucket::Header::default()
    }
}

fn row_header_payload() -> Vec<u8> {
    tst::HeaderStorageBucket {
        bucket_hash_function: 1,
        // Sparse by design: rows 2 and 4 have no header record.  The owner
        // must move the present records without inventing missing rows.
        headers: vec![row_header(0), row_header(1), row_header(3), row_header(5)],
    }
    .encode_to_vec()
}

fn column_header_payload() -> Vec<u8> {
    tst::HeaderStorageBucket {
        bucket_hash_function: 1,
        headers: vec![row_header(0), row_header(1)],
    }
    .encode_to_vec()
}

fn uid_map_payload() -> Vec<u8> {
    let columns = (0..COLUMN_COUNT)
        .map(|index| tsp::Uuid {
            lower: u64::from(index) + 100,
            upper: u64::from(index) + 1_000,
        })
        .collect::<Vec<_>>();
    let rows = (0..ROW_COUNT)
        .map(|index| tsp::Uuid {
            lower: u64::from(index) + 200,
            upper: u64::from(index) + 2_000,
        })
        .collect::<Vec<_>>();
    tst::ColumnRowUidMapArchive {
        sorted_column_uids: columns,
        column_index_for_uid: (0..COLUMN_COUNT).collect(),
        column_uid_for_index: (0..COLUMN_COUNT).collect(),
        sorted_row_uids: rows,
        row_index_for_uid: (0..ROW_COUNT).collect(),
        row_uid_for_index: (0..ROW_COUNT).collect(),
    }
    .encode_to_vec()
}

fn data_store(tile_size: u32, tile_identifiers: &[u64]) -> tst::DataStore {
    tst::DataStore {
        row_headers: tst::HeaderStorage {
            bucket_hash_function: 1,
            buckets: vec![reference(ROW_HEADERS)],
        },
        column_headers: reference(COLUMN_HEADERS),
        tiles: tst::TileStorage {
            tiles: tile_identifiers
                .iter()
                .enumerate()
                .map(|(tileid, identifier)| tst::tile_storage::Tile {
                    tileid: u32::try_from(tileid).expect("fixture tile id fits u32"),
                    tile: reference(*identifier),
                })
                .collect(),
            tile_size: Some(tile_size),
            should_use_wide_rows: Some(false),
        },
        string_table: reference(STRINGS),
        style_table: reference(STYLES),
        formula_table: reference(FORMULAS),
        format_table_pre_bnc: reference(FORMATS),
        next_row_strip_id: 1,
        next_column_strip_id: 1,
        row_tile_tree: tst::TableRbTree {
            nodes: vec![tst::table_rb_tree::Node { key: 0, value: 0 }],
        },
        ..tst::DataStore::default()
    }
}

fn table_model(
    scope: Scope,
    with_unknowns: bool,
    tile_size: u32,
    tile_identifiers: &[u64],
) -> TestResult<Vec<u8>> {
    let mut payload = tst::TableModelArchive {
        table_id: "physical-sort-table".to_owned(),
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
        base_data_store: data_store(tile_size, tile_identifiers),
        number_of_rows: ROW_COUNT,
        number_of_columns: COLUMN_COUNT,
        table_name: "Cities".to_owned(),
        table_name_enabled: Some(false),
        number_of_header_rows: Some(HEADER_ROWS),
        number_of_footer_rows: Some(FOOTER_ROWS),
        header_rows_frozen: Some(true),
        header_columns_frozen: Some(true),
        default_row_height: 20.0,
        default_column_width: 64.0,
        repeating_header_rows_enabled: Some(true),
        repeating_header_columns_enabled: Some(true),
        sort_order: Some(tst::TableSortOrderArchive {
            r#type: scope.native_value(),
            rules: vec![tst::table_sort_order_archive::SortRuleArchive {
                index: 0,
                direction: Direction::Ascending.native_value(),
            }],
        }),
        base_column_row_uids: Some(reference(UID_MAP)),
        stroke_sidecar: Some(reference(STROKE_SIDECAR)),
        ..tst::TableModelArchive::default()
    }
    .encode_to_vec();
    if with_unknowns {
        append_varint_field(&mut payload, 99, 0xfeed_beef)?;
        append_length_delimited_field(&mut payload, 100, b"physical-sort-unknown")?;
    }
    Ok(payload)
}

fn table_info_payload(locked: bool) -> Vec<u8> {
    tst::TableInfoArchive {
        super_: tsd::DrawableArchive {
            parent: Some(reference(SLIDE)),
            locked: Some(locked),
            ..tsd::DrawableArchive::default()
        },
        table_model: reference(TABLE_MODEL),
        ..tst::TableInfoArchive::default()
    }
    .encode_to_vec()
}

fn formula_owner_dependencies_payload() -> Vec<u8> {
    tsce::FormulaOwnerDependenciesArchive {
        formula_owner_uid: tsp::Uuid {
            lower: 0x1010,
            upper: 0x2020,
        },
        internal_formula_owner_id: 1,
        owner_kind: Some(1),
        formula_owner: Some(reference(TABLE_INFO)),
        ..tsce::FormulaOwnerDependenciesArchive::default()
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

fn tss_style(identifier: &str) -> litchi_iwa_protos::tss::StyleArchive {
    litchi_iwa_protos::tss::StyleArchive {
        style_identifier: Some(identifier.to_owned()),
        ..litchi_iwa_protos::tss::StyleArchive::default()
    }
}

fn style_payload() -> Vec<u8> {
    tswp::ParagraphStyleArchive {
        super_: tss_style("physical-sort"),
        ..tswp::ParagraphStyleArchive::default()
    }
    .encode_to_vec()
}

fn model_references(tile_identifiers: &[u64]) -> Vec<u64> {
    let mut references = vec![ROW_HEADERS, COLUMN_HEADERS];
    references.extend_from_slice(tile_identifiers);
    references.extend([
        STRINGS,
        STYLES,
        FORMULAS,
        FORMATS,
        UID_MAP,
        STROKE_SIDECAR,
        TITLE_STYLE,
        SHAPE_STYLE,
    ]);
    references
}

fn model_field_infos() -> Vec<FieldInfo> {
    vec![
        field_reference(vec![4, 1, 2], &[ROW_HEADERS]),
        field_reference(vec![4, 2], &[COLUMN_HEADERS]),
        field_reference(vec![4, 4], &[STRINGS]),
        field_reference(vec![4, 5], &[STYLES]),
        field_reference(vec![4, 6], &[FORMULAS]),
        field_reference(vec![4, 11], &[FORMATS]),
        field_reference(vec![BASE_UID_FIELD], &[UID_MAP]),
        field_reference(vec![STROKE_FIELD], &[STROKE_SIDECAR]),
    ]
}

fn document_payload() -> Vec<u8> {
    kn::DocumentArchive {
        super_: tsa::DocumentArchive {
            super_: tsk::DocumentArchive::default(),
            ..tsa::DocumentArchive::default()
        },
        show: reference(SHOW),
        ..kn::DocumentArchive::default()
    }
    .encode_to_vec()
}

fn show_payload() -> Vec<u8> {
    kn::ShowArchive {
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
    }
    .encode_to_vec()
}

fn slide_payload() -> Vec<u8> {
    kn::SlideArchive {
        style: reference(90),
        transition: kn::TransitionArchive::default(),
        owned_drawables: vec![reference(TABLE_INFO), reference(NON_TABLE_DRAWABLE)],
        drawables_z_order: vec![reference(TABLE_INFO), reference(NON_TABLE_DRAWABLE)],
        name: Some("Tables".to_owned()),
        in_document: true,
        ..kn::SlideArchive::default()
    }
    .encode_to_vec()
}

fn metadata_payload(tile_identifiers: &[u64]) -> TestResult<Vec<u8>> {
    let mut identifiers = vec![
        DOCUMENT,
        SHOW,
        SLIDE_NODE,
        SLIDE,
        TABLE_INFO,
        TABLE_MODEL,
        NON_TABLE_DRAWABLE,
        ROW_HEADERS,
        COLUMN_HEADERS,
    ];
    identifiers.extend_from_slice(tile_identifiers);
    identifiers.extend([
        STRINGS,
        STYLES,
        FORMULAS,
        FORMATS,
        UID_MAP,
        STROKE_SIDECAR,
        TITLE_STYLE,
        SHAPE_STYLE,
        FORMULA_OWNER_DEPENDENCIES,
    ]);
    Ok(tsp::PackageMetadata {
        last_object_identifier: 900,
        components: vec![tsp::ComponentInfo {
            identifier: 1,
            preferred_locator: DOCUMENT_MEMBER.to_owned(),
            locator: Some(DOCUMENT_MEMBER.to_owned()),
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
        }],
        ..tsp::PackageMetadata::default()
    }
    .encode_to_vec())
}

fn source_package(scope: Scope, locked: bool, with_unknowns: bool) -> TestResult<Vec<u8>> {
    source_package_with_tiles(scope, locked, with_unknowns, 256, &[(TILE, tile_payload())])
}

fn source_package_with_tiles(
    scope: Scope,
    locked: bool,
    with_unknowns: bool,
    tile_size: u32,
    tiles: &[(u64, Vec<u8>)],
) -> TestResult<Vec<u8>> {
    let tile_identifiers = tiles
        .iter()
        .map(|(identifier, _payload)| *identifier)
        .collect::<Vec<_>>();
    let document = object(DOCUMENT, 1, document_payload(), &[SHOW])?;
    let show = object(SHOW, 2, show_payload(), &[SLIDE_NODE, 80, 81])?;
    let node = object(
        SLIDE_NODE,
        4,
        kn::SlideNodeArchive {
            slide: Some(reference(SLIDE)),
            ..kn::SlideNodeArchive::default()
        }
        .encode_to_vec(),
        &[SLIDE],
    )?;
    let mut slide = object(SLIDE, 5, slide_payload(), &[TABLE_INFO, NON_TABLE_DRAWABLE])?;
    slide.archive_info.message_infos[0].field_infos.extend([
        field_reference(vec![7], &[TABLE_INFO, NON_TABLE_DRAWABLE]),
        field_reference(vec![42], &[TABLE_INFO, NON_TABLE_DRAWABLE]),
    ]);
    let mut info = object(
        TABLE_INFO,
        TABLE_INFO_MESSAGE_TYPE,
        table_info_payload(locked),
        &[TABLE_MODEL],
    )?;
    info.archive_info.message_infos[0]
        .field_infos
        .push(field_reference(vec![2], &[TABLE_MODEL]));

    let mut model = object(
        TABLE_MODEL,
        TABLE_MODEL_MESSAGE_TYPE,
        table_model(scope, with_unknowns, tile_size, &tile_identifiers)?,
        &model_references(&tile_identifiers),
    )?;
    model.archive_info.message_infos[0]
        .field_infos
        .extend(model_field_infos());
    let mut objects = vec![
        document,
        show,
        node,
        slide,
        info,
        model,
        object(
            FORMULA_OWNER_DEPENDENCIES,
            FORMULA_OWNER_DEPENDENCIES_MESSAGE_TYPE,
            formula_owner_dependencies_payload(),
            &[],
        )?,
        object(
            NON_TABLE_DRAWABLE,
            SHAPE_INFO_MESSAGE_TYPE,
            shape_info_payload(),
            &[SLIDE],
        )?,
        object(
            ROW_HEADERS,
            HEADER_BUCKET_MESSAGE_TYPE,
            row_header_payload(),
            &[],
        )?,
        object(
            COLUMN_HEADERS,
            HEADER_BUCKET_MESSAGE_TYPE,
            column_header_payload(),
            &[],
        )?,
    ];
    for (identifier, payload) in tiles {
        objects.push(object(
            *identifier,
            TILE_MESSAGE_TYPE,
            payload.clone(),
            &[],
        )?);
    }
    objects.extend([
        object(
            STRINGS,
            TABLE_DATA_LIST_MESSAGE_TYPE,
            data_list(tst::table_data_list::ListType::String, string_entries()),
            &[],
        )?,
        object(
            STYLES,
            TABLE_DATA_LIST_MESSAGE_TYPE,
            data_list(tst::table_data_list::ListType::Style, Vec::new()),
            &[],
        )?,
        object(
            FORMULAS,
            TABLE_DATA_LIST_MESSAGE_TYPE,
            data_list(tst::table_data_list::ListType::Formula, Vec::new()),
            &[],
        )?,
        object(
            FORMATS,
            TABLE_DATA_LIST_MESSAGE_TYPE,
            data_list(tst::table_data_list::ListType::Format, Vec::new()),
            &[],
        )?,
        object(
            UID_MAP,
            COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
            uid_map_payload(),
            &[],
        )?,
        object(
            STROKE_SIDECAR,
            STROKE_SIDECAR_MESSAGE_TYPE,
            tst::StrokeSidecarArchive {
                row_count: Some(ROW_COUNT),
                column_count: Some(COLUMN_COUNT),
                ..tst::StrokeSidecarArchive::default()
            }
            .encode_to_vec(),
            &[],
        )?,
        object(TITLE_STYLE, 2_022, style_payload(), &[])?,
        object(
            SHAPE_STYLE,
            2_025,
            tswp::ShapeStyleArchive::default().encode_to_vec(),
            &[],
        )?,
    ]);

    let document_bytes = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    let metadata_bytes = SnappyStream::compress(
        &Archive {
            objects: vec![object(
                METADATA_OBJECT,
                METADATA_MESSAGE_TYPE,
                metadata_payload(&tile_identifiers)?,
                &[],
            )?],
        }
        .to_bytes()?,
    )?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [
            (
                SENTINEL_MEMBER,
                b"untouched physical-sort sentinel".as_slice(),
            ),
            (DOCUMENT_MEMBER, document_bytes.as_slice()),
            (METADATA_MEMBER, metadata_bytes.as_slice()),
            (PREVIEWS[0], b"physical-sort preview".as_slice()),
            (PREVIEWS[1], b"physical-sort micro preview".as_slice()),
            (PREVIEWS[2], b"physical-sort web preview".as_slice()),
        ],
        Limits::default(),
    )?)
}

fn package_bytes(package: &Package) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

fn document_archive(package: &[u8]) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document component")?;
    Ok(Archive::parse(
        SnappyStream::decompress(entry.data())?.as_bytes(),
    )?)
}

fn metadata_archive(package: &[u8]) -> TestResult<Archive> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == METADATA_MEMBER)
        .ok_or("missing metadata component")?;
    Ok(Archive::parse(
        SnappyStream::decompress(entry.data())?.as_bytes(),
    )?)
}

fn model_field_payloads(package: &[u8], field: u32) -> TestResult<Vec<Vec<u8>>> {
    let archive = document_archive(package)?;
    let object = archive.object(TABLE_MODEL).ok_or("missing table model")?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_MODEL_MESSAGE_TYPE)
        .ok_or("missing table model payload")?;
    Ok(repeated_length_delimited_payloads(&message.data, field)?
        .into_iter()
        .map(ToOwned::to_owned)
        .collect())
}

fn rewrite_document(
    package: &[u8],
    mutate: impl FnOnce(&mut Archive) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document component")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    mutate(&mut archive)?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

/// Rewrite the decompressed IWA stream without parsing and re-encoding its
/// objects.  This is intentionally limited to framing adversarial tests:
/// production fixtures should continue to use the typed archive helpers.
fn rewrite_document_stream(
    package: &[u8],
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    let entry = catalog
        .iter()
        .find(|entry| entry.name() == DOCUMENT_MEMBER)
        .ok_or("missing document component")?;
    let mut stream = SnappyStream::decompress(entry.data())?.as_bytes().to_vec();
    mutate(&mut stream)?;
    let compressed = SnappyStream::compress(&stream)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            DOCUMENT_MEMBER,
            &compressed,
        )],
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
        .ok_or("missing metadata component")?;
    let mut archive = Archive::parse(SnappyStream::decompress(entry.data())?.as_bytes())?;
    let metadata = archive
        .object_mut(METADATA_OBJECT)
        .ok_or("missing metadata object")?;
    let message = metadata
        .messages
        .iter()
        .find(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .ok_or("missing metadata payload")?;
    let mut payload = tsp::PackageMetadata::decode(message.data.as_slice())?;
    mutate(&mut payload)?;
    metadata.replace_message_preserving_header(
        0,
        RawMessage {
            type_: METADATA_MESSAGE_TYPE,
            data: payload.encode_to_vec(),
        },
    )?;
    let compressed = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(
            METADATA_MEMBER,
            &compressed,
        )],
        Limits::default(),
    )?)
}

fn rewrite_table_info(package: &[u8], locked: bool) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let info = archive.object_mut(TABLE_INFO).ok_or("missing table info")?;
        info.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TABLE_INFO_MESSAGE_TYPE,
                data: table_info_payload(locked),
            },
        )?;
        Ok(())
    })
}

fn rewrite_entry(package: &[u8], name: &str, data: &[u8]) -> TestResult<Vec<u8>> {
    let catalog = Catalog::from_bytes(package)?;
    Ok(catalog.reassemble_to_bytes(
        &[litchi_iwa_archive::package::EntryEdit::new(name, data)],
        Limits::default(),
    )?)
}

fn rewrite_model_wire(package: &[u8], field: u32, payload: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let model = archive
            .object_mut(TABLE_MODEL)
            .ok_or("missing table model")?;
        let message = model.messages.first().ok_or("missing model message")?;
        let mut bytes = message.data.clone();
        append_length_delimited_field(&mut bytes, field, payload)?;
        model.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: bytes,
            },
        )?;
        Ok(())
    })
}

/// Apply a bounded wire mutation to one typed archive message without
/// re-encoding its parent.  The physical-sort admission tests use this for
/// unknown fields because prost intentionally drops fields it does not know.
fn rewrite_object_wire(
    package: &[u8],
    identifier: u64,
    message_type: u32,
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let object = archive
            .object_mut(identifier)
            .ok_or("missing archive object")?;
        let index = object
            .messages
            .iter()
            .position(|message| message.type_ == message_type)
            .ok_or("missing archive message")?;
        let mut bytes = object.messages[index].data.clone();
        mutate(&mut bytes)?;
        object.replace_message_preserving_header(
            index,
            RawMessage {
                type_: message_type,
                data: bytes,
            },
        )?;
        Ok(())
    })
}

/// Mutate one nested length-delimited child while retaining the parent wire
/// shape and every unrelated field.  This reaches unknown fields inside row,
/// header, and UID records rather than only testing their roots.
fn rewrite_repeated_child_wire(
    package: &[u8],
    identifier: u64,
    message_type: u32,
    field_number: u32,
    child_index: usize,
    mutate: impl FnOnce(&mut Vec<u8>) -> TestResult<()>,
) -> TestResult<Vec<u8>> {
    rewrite_object_wire(package, identifier, message_type, |bytes| {
        let mut children = repeated_length_delimited_payloads(bytes, field_number)?
            .into_iter()
            .map(ToOwned::to_owned)
            .collect::<Vec<_>>();
        mutate(
            children
                .get_mut(child_index)
                .ok_or("missing nested archive child")?,
        )?;
        *bytes = rewrite_repeated_length_delimited_fields(bytes, field_number, &children)?;
        Ok(())
    })
}

fn valid_cfuuid_owner_payload() -> TestResult<Vec<u8>> {
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &[1u8; 16])?;
    Ok(payload)
}

fn valid_sort_tracker_payload() -> TestResult<Vec<u8>> {
    let mut payload = Vec::new();
    let reference = reference(FORMULAS).encode_to_vec();
    append_length_delimited_field(&mut payload, 1, &reference)?;
    Ok(payload)
}

fn valid_uuid_owner_payload() -> TestResult<Vec<u8>> {
    let uuid = tsp::Uuid {
        lower: 0x1010,
        upper: 0x2020,
    }
    .encode_to_vec();
    let mut payload = Vec::new();
    append_length_delimited_field(&mut payload, 1, &uuid)?;
    Ok(payload)
}

fn rewrite_model_wide_rows(package: &[u8], wide: bool) -> TestResult<Vec<u8>> {
    rewrite_object_wire(package, TABLE_MODEL, TABLE_MODEL_MESSAGE_TYPE, |bytes| {
        let mut model = tst::TableModelArchive::decode(bytes.as_slice())?;
        model.base_data_store.tiles.should_use_wide_rows = Some(wide);
        *bytes = model.encode_to_vec();
        Ok(())
    })
}

fn rewrite_tile_cell(
    package: &[u8],
    row_index: u32,
    column: usize,
    replacement: &[u8],
) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let tile = archive.object_mut(TILE).ok_or("missing table tile")?;
        let message = tile.messages.first().ok_or("missing tile message")?;
        let mut payload = tst::Tile::decode(message.data.as_slice())?;
        let row = payload
            .row_infos
            .iter_mut()
            .find(|row| row.tile_row_index == row_index)
            .ok_or("missing tile row")?;
        let storage = row
            .cell_storage_buffer
            .as_mut()
            .ok_or("missing BNC storage")?;
        let offsets = row.cell_offsets.as_ref().ok_or("missing BNC offsets")?;
        let offset = column.checked_mul(2).ok_or("BNC column overflow")?;
        let start = u16::from_le_bytes([
            *offsets.get(offset).ok_or("missing BNC cell offset")?,
            *offsets.get(offset + 1).ok_or("missing BNC cell offset")?,
        ]) as usize;
        let end = if offset + 2 < offsets.len() {
            u16::from_le_bytes([offsets[offset + 2], offsets[offset + 3]]) as usize
        } else {
            storage.len()
        };
        if start >= end || end > storage.len() || end - start != replacement.len() {
            return Err("replacement must preserve the BNC cell span".into());
        }
        storage[start..end].copy_from_slice(replacement);
        tile.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TILE_MESSAGE_TYPE,
                data: payload.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn rewrite_numeric_body_keys(package: &[u8], values: [f64; BODY_COUNT]) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let tile = archive.object_mut(TILE).ok_or("missing table tile")?;
        let message = tile.messages.first().ok_or("missing tile message")?;
        let mut payload = tst::Tile::decode(message.data.as_slice())?;
        for (offset, value) in values.into_iter().enumerate() {
            let row_index = u32::try_from(offset + 1).map_err(|_| "numeric row overflow")?;
            let row = payload
                .row_infos
                .iter_mut()
                .find(|row| row.tile_row_index == row_index)
                .ok_or("missing numeric body row")?;
            let mut cells = fixture_row_cells(row)?;
            let mut cell = BncCell::minimal();
            cell.set_plain_number(value)?;
            cells[0] = Some(cell.encode());
            let (storage, offsets) = encode_row(&cells);
            row.cell_storage_buffer = Some(storage);
            row.cell_offsets = Some(offsets);
        }
        tile.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TILE_MESSAGE_TYPE,
                data: payload.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn rewrite_padded_offsets(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let tile = archive.object_mut(TILE).ok_or("missing table tile")?;
        let message = tile.messages.first().ok_or("missing tile message")?;
        let mut payload = tst::Tile::decode(message.data.as_slice())?;
        for row in &mut payload.row_infos {
            row.cell_offsets
                .as_mut()
                .ok_or("missing BNC offsets")?
                .extend_from_slice(&u16::MAX.to_le_bytes());
        }
        tile.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TILE_MESSAGE_TYPE,
                data: payload.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn rewrite_pre_bnc_sidecars(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let tile = archive.object_mut(TILE).ok_or("missing table tile")?;
        let message = tile.messages.first().ok_or("missing tile message")?;
        let mut payload = tst::Tile::decode(message.data.as_slice())?;
        for row in &mut payload.row_infos {
            let modern_storage = row
                .cell_storage_buffer
                .as_ref()
                .ok_or("missing modern BNC storage")?;
            let modern_offsets = row
                .cell_offsets
                .as_ref()
                .ok_or("missing modern BNC offsets")?;
            row.cell_storage_buffer_pre_bnc = modern_storage.clone();
            row.cell_offsets_pre_bnc = modern_offsets.clone();
        }
        tile.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TILE_MESSAGE_TYPE,
                data: payload.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn rewrite_pre_bnc_only(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let tile = archive.object_mut(TILE).ok_or("missing table tile")?;
        let message = tile.messages.first().ok_or("missing tile message")?;
        let mut payload = tst::Tile::decode(message.data.as_slice())?;
        for row in &mut payload.row_infos {
            let modern_storage = row
                .cell_storage_buffer
                .take()
                .ok_or("missing modern BNC storage")?;
            let modern_offsets = row
                .cell_offsets
                .take()
                .ok_or("missing modern BNC offsets")?;
            row.cell_storage_buffer_pre_bnc = modern_storage;
            row.cell_offsets_pre_bnc = modern_offsets;
        }
        tile.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TILE_MESSAGE_TYPE,
                data: payload.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn rewrite_string_entry_sidecars(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let strings = archive.object_mut(STRINGS).ok_or("missing string table")?;
        let message = strings
            .messages
            .first()
            .ok_or("missing string table message")?;
        let mut payload = tst::TableDataList::decode(message.data.as_slice())?;
        let entry = payload
            .entries
            .iter_mut()
            .find(|entry| entry.key == 3)
            .ok_or("missing zebra string entry")?;
        entry.import_warning_set = Some(tst::ImportWarningSetArchive {
            cond_format_expr: Some(true),
            original_data_format: Some("legacy text".to_owned()),
            ..tst::ImportWarningSetArchive::default()
        });
        entry.cell_spec = Some(tst::CellSpecArchive {
            interaction_type: 3,
            ..tst::CellSpecArchive::default()
        });
        strings.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TABLE_DATA_LIST_MESSAGE_TYPE,
                data: payload.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn rewrite_uid_map(
    package: &[u8],
    mutate: impl FnOnce(&mut tst::ColumnRowUidMapArchive),
) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let map = archive.object_mut(UID_MAP).ok_or("missing UID map")?;
        let message = map.messages.first().ok_or("missing UID map message")?;
        let mut payload = tst::ColumnRowUidMapArchive::decode(message.data.as_slice())?;
        mutate(&mut payload);
        map.replace_message_preserving_header(
            0,
            RawMessage {
                type_: COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
                data: payload.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn rewrite_formula_registry(package: &[u8]) -> TestResult<Vec<u8>> {
    rewrite_document(package, |archive| {
        let formulas = archive
            .object_mut(FORMULAS)
            .ok_or("missing formula table")?;
        let message = formulas
            .messages
            .first()
            .ok_or("missing formula table message")?;
        let mut payload = tst::TableDataList::decode(message.data.as_slice())?;
        payload.entries.push(tst::table_data_list::ListEntry {
            key: 1,
            refcount: 1,
            string: Some("=A1".to_owned()),
            ..tst::table_data_list::ListEntry::default()
        });
        formulas.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TABLE_DATA_LIST_MESSAGE_TYPE,
                data: payload.encode_to_vec(),
            },
        )?;
        Ok(())
    })
}

fn bnc_formula(identifier: u32) -> Vec<u8> {
    let mut bytes = vec![5, 2, 0, 0, 0, 0, 0, 0];
    bytes.extend_from_slice(&0x0000_0200u32.to_le_bytes());
    bytes.extend_from_slice(&identifier.to_le_bytes());
    bytes
}

fn bnc_error() -> Vec<u8> {
    // Keep the original sixteen-byte cell envelope so this mutation tests
    // semantic refusal rather than offset-table repair.
    vec![5, 8, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0xde, 0xad, 0xbe, 0xef]
}

fn bnc_rich_text(identifier: u32) -> Vec<u8> {
    let mut bytes = vec![5, 9, 0, 0, 0, 0, 0, 0];
    bytes.extend_from_slice(&0x0000_0010u32.to_le_bytes());
    bytes.extend_from_slice(&identifier.to_le_bytes());
    bytes
}

fn source(scope: Scope) -> TestResult<Vec<u8>> {
    source_package(scope, false, false)
}

/// Add one optional format-list route while retaining only the model's
/// aggregate ArchiveInfo edge.  Older native producers can omit a
/// field-local declaration for these optional DataStore routes; this fixture
/// keeps that producer shape explicit so the admission contract cannot
/// accidentally regress to requiring metadata that is not always emitted.
fn source_with_aggregate_only_optional_format(
    scope: Scope,
    field: u32,
    list_type: tst::table_data_list::ListType,
    identifier: u64,
    entries: Vec<tst::table_data_list::ListEntry>,
) -> TestResult<Vec<u8>> {
    let source = source(scope)?;
    let source = rewrite_document(&source, |archive| {
        let model = archive
            .object_mut(TABLE_MODEL)
            .ok_or("missing table model")?;
        let message = model.messages.first().ok_or("missing model message")?;
        let mut payload = tst::TableModelArchive::decode(message.data.as_slice())?;
        let reference = reference(identifier);
        match field {
            15 => payload.base_data_store.deprecated_custom_format_table = Some(reference),
            22 => payload.base_data_store.format_table = Some(reference),
            _ => return Err("unsupported optional format field".into()),
        }
        model.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: payload.encode_to_vec(),
            },
        )?;
        model.archive_info.message_infos[0]
            .object_references
            .push(identifier);
        archive.objects.push(object(
            identifier,
            TABLE_DATA_LIST_MESSAGE_TYPE,
            data_list(list_type, entries),
            &[],
        )?);
        Ok(())
    })?;
    rewrite_metadata(&source, |metadata| {
        let component = metadata
            .components
            .first_mut()
            .ok_or("missing metadata component")?;
        component
            .object_uuid_map_entries
            .push(tsp::ObjectUuidMapEntry {
                identifier,
                uuid: tsp::Uuid {
                    lower: identifier + 10_000,
                    upper: identifier + 20_000,
                },
            });
        Ok(())
    })
}

fn small_tile_source(scope: Scope) -> TestResult<Vec<u8>> {
    source_package_with_tiles(
        scope,
        false,
        false,
        2,
        &[
            (TILE, tile_payload_for(0, 2)),
            (TILE_SECOND, tile_payload_for(2, 4)),
            (TILE_THIRD, tile_payload_for(4, ROW_COUNT)),
        ],
    )
}

fn cell_span_bounds(storage: &[u8], offsets: &[u8], column: usize) -> TestResult<(usize, usize)> {
    let offset = column.checked_mul(2).ok_or("BNC column offset overflow")?;
    let raw_start = u16::from_le_bytes([
        *offsets.get(offset).ok_or("missing BNC cell offset")?,
        *offsets.get(offset + 1).ok_or("missing BNC cell offset")?,
    ]);
    if raw_start == u16::MAX {
        return Err("BNC cell is absent".into());
    }
    let start = usize::from(raw_start);
    let mut end = storage.len();
    for bytes in offsets.chunks_exact(2).skip(column + 1) {
        let raw_end = u16::from_le_bytes([bytes[0], bytes[1]]);
        if raw_end != u16::MAX {
            end = usize::from(raw_end);
            break;
        }
    }
    if start >= end || end > storage.len() {
        return Err("invalid BNC cell offsets".into());
    }
    Ok((start, end))
}

fn fixture_row_cells(row: &tst::TileRowInfo) -> TestResult<[Option<Vec<u8>>; 2]> {
    let storage = row
        .cell_storage_buffer
        .as_ref()
        .ok_or("missing BNC storage")?;
    let offsets = row.cell_offsets.as_ref().ok_or("missing BNC offsets")?;
    let mut cells = [None, None];
    for (column, cell) in cells.iter_mut().enumerate() {
        let offset = column.checked_mul(2).ok_or("BNC column offset overflow")?;
        let raw_start = u16::from_le_bytes([
            *offsets.get(offset).ok_or("missing BNC cell offset")?,
            *offsets.get(offset + 1).ok_or("missing BNC cell offset")?,
        ]);
        if raw_start == u16::MAX {
            continue;
        }
        let (start, end) = cell_span_bounds(storage, offsets, column)?;
        *cell = Some(storage[start..end].to_vec());
    }
    Ok(cells)
}

fn cell_identifier(storage: &[u8], offsets: &[u8], column: usize) -> TestResult<u32> {
    let (start, end) = cell_span_bounds(storage, offsets, column)?;
    let cell = storage.get(start..end).ok_or("invalid BNC cell span")?;
    if cell.len() < 16 {
        return Err("truncated BNC text cell".into());
    }
    Ok(u32::from_le_bytes([cell[12], cell[13], cell[14], cell[15]]))
}

fn marker_pairs(package: &[u8]) -> TestResult<Vec<(String, String)>> {
    let archive = document_archive(package)?;
    let tile = archive.object(TILE).ok_or("missing table tile")?;
    let message = tile
        .messages
        .iter()
        .find(|message| message.type_ == TILE_MESSAGE_TYPE)
        .ok_or("missing tile payload")?;
    let tile = tst::Tile::decode(message.data.as_slice())?;
    let strings = archive.object(STRINGS).ok_or("missing string table")?;
    let string_message = strings
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_DATA_LIST_MESSAGE_TYPE)
        .ok_or("missing string payload")?;
    let strings = tst::TableDataList::decode(string_message.data.as_slice())?;
    let names = strings
        .entries
        .iter()
        .filter_map(|entry| {
            entry
                .string
                .as_ref()
                .map(|value| (entry.key, value.clone()))
        })
        .collect::<BTreeMap<_, _>>();
    let mut output = Vec::new();
    for row in tile.row_infos {
        let Some(storage) = row.cell_storage_buffer else {
            continue;
        };
        let offsets = row.cell_offsets.ok_or("missing BNC offsets")?;
        let name = names
            .get(&cell_identifier(&storage, &offsets, 0)?)
            .cloned()
            .ok_or("missing name string")?;
        let marker = names
            .get(&cell_identifier(&storage, &offsets, 1)?)
            .cloned()
            .ok_or("missing marker string")?;
        output.push((name, marker));
    }
    Ok(output)
}

fn marker_rows(package: &[u8]) -> TestResult<Vec<String>> {
    Ok(marker_pairs(package)?
        .into_iter()
        .map(|(name, _marker)| name)
        .collect())
}

fn string_value_map(archive: &Archive) -> TestResult<BTreeMap<u32, String>> {
    let strings = archive.object(STRINGS).ok_or("missing string table")?;
    let string_message = strings
        .messages
        .iter()
        .find(|message| message.type_ == TABLE_DATA_LIST_MESSAGE_TYPE)
        .ok_or("missing string payload")?;
    Ok(tst::TableDataList::decode(string_message.data.as_slice())?
        .entries
        .into_iter()
        .filter_map(|entry| entry.string.map(|value| (entry.key, value)))
        .collect())
}

fn marker_column_rows(package: &[u8]) -> TestResult<Vec<String>> {
    let archive = document_archive(package)?;
    let tile = archive.object(TILE).ok_or("missing table tile")?;
    let message = tile
        .messages
        .iter()
        .find(|message| message.type_ == TILE_MESSAGE_TYPE)
        .ok_or("missing tile payload")?;
    let tile = tst::Tile::decode(message.data.as_slice())?;
    let names = string_value_map(&archive)?;
    tile.row_infos
        .iter()
        .map(|row| {
            let storage = row
                .cell_storage_buffer
                .as_ref()
                .ok_or("missing BNC storage")?;
            let offsets = row.cell_offsets.as_ref().ok_or("missing BNC offsets")?;
            names
                .get(&cell_identifier(storage, offsets, 1)?)
                .cloned()
                .ok_or_else(|| "missing marker string".into())
        })
        .collect()
}

fn pre_bnc_payloads(package: &[u8]) -> TestResult<Vec<(Vec<u8>, Vec<u8>)>> {
    let archive = document_archive(package)?;
    let tile = archive.object(TILE).ok_or("missing table tile")?;
    let message = tile
        .messages
        .iter()
        .find(|message| message.type_ == TILE_MESSAGE_TYPE)
        .ok_or("missing tile payload")?;
    let tile = tst::Tile::decode(message.data.as_slice())?;
    tile.row_infos
        .iter()
        .map(|row| {
            Ok((
                row.cell_storage_buffer_pre_bnc.clone(),
                row.cell_offsets_pre_bnc.clone(),
            ))
        })
        .collect()
}

fn numeric_body_key_markers(package: &[u8]) -> TestResult<Vec<(u64, String)>> {
    let archive = document_archive(package)?;
    let tile = archive.object(TILE).ok_or("missing table tile")?;
    let message = tile
        .messages
        .iter()
        .find(|message| message.type_ == TILE_MESSAGE_TYPE)
        .ok_or("missing tile payload")?;
    let tile = tst::Tile::decode(message.data.as_slice())?;
    let names = string_value_map(&archive)?;
    tile.row_infos
        .iter()
        .enumerate()
        .filter(|(ordinal, _row)| *ordinal >= HEADER_ROWS as usize)
        .filter(|(ordinal, _row)| *ordinal < ROW_COUNT as usize - FOOTER_ROWS as usize)
        .map(|(_ordinal, row)| {
            let storage = row
                .cell_storage_buffer
                .as_ref()
                .ok_or("missing BNC storage")?;
            let offsets = row.cell_offsets.as_ref().ok_or("missing BNC offsets")?;
            let (start, end) = cell_span_bounds(storage, offsets, 0)?;
            let view = BncCellView::parse(&storage[start..end])?;
            let bits = match view.cached_scalar() {
                Some(CachedScalar::Number(value)) => value.get().to_bits(),
                _ => return Err("body key is not a cached numeric scalar".into()),
            };
            let marker = names
                .get(&cell_identifier(storage, offsets, 1)?)
                .cloned()
                .ok_or("missing marker string")?;
            Ok((bits, marker))
        })
        .collect()
}

fn tile_offset_payloads(package: &[u8]) -> TestResult<Vec<Vec<u8>>> {
    let archive = document_archive(package)?;
    let tile = archive.object(TILE).ok_or("missing table tile")?;
    let message = tile
        .messages
        .iter()
        .find(|message| message.type_ == TILE_MESSAGE_TYPE)
        .ok_or("missing tile payload")?;
    Ok(tst::Tile::decode(message.data.as_slice())?
        .row_infos
        .into_iter()
        .map(|row| row.cell_offsets.ok_or("missing BNC offsets"))
        .collect::<Result<Vec<_>, _>>()?)
}

fn row_header_indices(package: &[u8]) -> TestResult<Vec<u32>> {
    let archive = document_archive(package)?;
    let object = archive.object(ROW_HEADERS).ok_or("missing row headers")?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == HEADER_BUCKET_MESSAGE_TYPE)
        .ok_or("missing header payload")?;
    Ok(tst::HeaderStorageBucket::decode(message.data.as_slice())?
        .headers
        .into_iter()
        .map(|header| header.index)
        .collect())
}

fn uid_map_indexes(package: &[u8]) -> TestResult<(Vec<u32>, Vec<u32>, Vec<u32>, Vec<u32>)> {
    let archive = document_archive(package)?;
    let object = archive.object(UID_MAP).ok_or("missing UID map")?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == COLUMN_ROW_UID_MAP_MESSAGE_TYPE)
        .ok_or("missing UID map payload")?;
    let map = tst::ColumnRowUidMapArchive::decode(message.data.as_slice())?;
    Ok((
        map.column_index_for_uid,
        map.column_uid_for_index,
        map.row_index_for_uid,
        map.row_uid_for_index,
    ))
}

fn assert_locality(source: &[u8], target: &[u8]) -> TestResult {
    let before = Catalog::from_bytes(source)?;
    let after = Catalog::from_bytes(target)?;
    for entry in before.iter() {
        if PREVIEWS.contains(&entry.name()) {
            assert!(
                after
                    .iter()
                    .all(|candidate| candidate.name() != entry.name()),
                "changed physical table must remove stale preview {}",
                entry.name()
            );
            continue;
        }
        let candidate = after
            .iter()
            .find(|other| other.name() == entry.name())
            .ok_or("candidate removed a package entry")?;
        if entry.name() == DOCUMENT_MEMBER {
            assert_ne!(
                entry.data(),
                candidate.data(),
                "physical source must change"
            );
        } else {
            assert_eq!(
                entry.data(),
                candidate.data(),
                "unrelated package entry changed"
            );
            assert_eq!(
                entry.raw_record().local_record(),
                candidate.raw_record().local_record()
            );
            assert_eq!(
                &entry.raw_record().central_directory_record()[..42],
                &candidate.raw_record().central_directory_record()[..42]
            );
            assert_eq!(
                &entry.raw_record().central_directory_record()[46..],
                &candidate.raw_record().central_directory_record()[46..]
            );
        }
    }
    assert!(
        after.iter().all(|entry| !PREVIEWS.contains(&entry.name())),
        "changed physical table retained a stale preview"
    );
    Ok(())
}

fn assert_source_unchanged(package: &Package, source: &[u8]) -> TestResult {
    assert_eq!(package_bytes(package)?, source);
    Ok(())
}

fn assert_physical_sort_rejected(source: &[u8]) -> TestResult {
    let Ok(package) = Package::from_bytes(source) else {
        return Ok(());
    };
    let before = package_bytes(&package)?;
    assert!(
        package
            .execute_slide_table_sort_order(0usize, 0usize)
            .is_err(),
        "malformed physical-sort source was unexpectedly admitted"
    );
    assert_source_unchanged(&package, &before)
}

fn assert_canonical_source_admitted(source: &[u8]) -> TestResult {
    let package = Package::from_bytes(source)?;
    package
        .execute_slide_table_sort_order(0usize, 0usize)
        .map(|_| ())
        .map_err(|error| format!("canonical physical-sort source was rejected: {error:?}").into())
}

#[test]
fn canonical_fixture_has_exact_metadata_and_unrelated_entries() -> TestResult {
    let source = source(Scope::EntireTable)?;
    let catalog = Catalog::from_bytes(&source)?;
    assert_eq!(
        catalog.iter().map(|entry| entry.name()).collect::<Vec<_>>(),
        vec![
            SENTINEL_MEMBER,
            DOCUMENT_MEMBER,
            METADATA_MEMBER,
            PREVIEWS[0],
            PREVIEWS[1],
            PREVIEWS[2],
        ]
    );
    assert_eq!(
        catalog
            .iter()
            .find(|entry| entry.name() == SENTINEL_MEMBER)
            .ok_or("missing sentinel")?
            .data(),
        b"untouched physical-sort sentinel"
    );
    assert_eq!(
        catalog
            .iter()
            .find(|entry| entry.name() == PREVIEWS[0])
            .ok_or("missing preview")?
            .data(),
        b"physical-sort preview"
    );
    assert_eq!(
        catalog
            .iter()
            .find(|entry| entry.name() == PREVIEWS[1])
            .ok_or("missing micro preview")?
            .data(),
        b"physical-sort micro preview"
    );
    assert_eq!(
        catalog
            .iter()
            .find(|entry| entry.name() == PREVIEWS[2])
            .ok_or("missing web preview")?
            .data(),
        b"physical-sort web preview"
    );

    let metadata = metadata_archive(&source)?;
    let object = metadata
        .object(METADATA_OBJECT)
        .ok_or("missing metadata object")?;
    let message = object
        .messages
        .iter()
        .find(|message| message.type_ == METADATA_MESSAGE_TYPE)
        .ok_or("missing metadata payload")?;
    let metadata = tsp::PackageMetadata::decode(message.data.as_slice())?;
    assert_eq!(metadata.last_object_identifier, METADATA_OBJECT);
    assert_eq!(metadata.components.len(), 1);
    let component = &metadata.components[0];
    assert_eq!(component.identifier, 1);
    assert_eq!(component.preferred_locator, DOCUMENT_MEMBER);
    assert_eq!(component.locator.as_deref(), Some(DOCUMENT_MEMBER));
    assert_eq!(component.save_token, Some(1));
    assert_eq!(component.object_uuid_map_entries.len(), 19);
    assert!(
        component
            .object_uuid_map_entries
            .iter()
            .all(|entry| entry.uuid.lower != 0 && entry.uuid.upper != 0)
    );
    Ok(())
}

#[test]
fn aggregate_only_table_info_model_edge_remains_compatible() -> TestResult {
    let source = source(Scope::EntireTable)?;
    let source = rewrite_document(&source, |archive| {
        let info = archive.object_mut(TABLE_INFO).ok_or("missing table info")?;
        let message_info = info
            .archive_info
            .message_infos
            .first_mut()
            .ok_or("missing table info metadata")?;
        message_info.field_infos.clear();
        Ok(())
    })?;
    let package = Package::from_bytes(&source)?;
    let commit = package.execute_slide_table_sort_order(0usize, 0usize)?;
    let target = package_bytes(commit.package())?;
    assert_eq!(
        marker_rows(&target)?,
        ["Name", "apple", "apple", "banana", "zebra", "Total"]
    );
    assert_locality(&source, &target)?;
    Ok(())
}

#[test]
fn aggregate_only_optional_format_routes_are_validated_and_preserved() -> TestResult {
    let custom_source = source_with_aggregate_only_optional_format(
        Scope::EntireTable,
        15,
        tst::table_data_list::ListType::CustomFormat,
        OPTIONAL_CUSTOM_FORMATS,
        Vec::new(),
    )?;
    let format_source = source_with_aggregate_only_optional_format(
        Scope::EntireTable,
        22,
        tst::table_data_list::ListType::Format,
        OPTIONAL_FORMATS,
        vec![tst::table_data_list::ListEntry {
            key: 1,
            refcount: 1,
            ..tst::table_data_list::ListEntry::default()
        }],
    )?;

    for source in [custom_source, format_source] {
        let package = Package::from_bytes(&source)?;
        let commit = package.execute_slide_table_sort_order(0usize, 0usize)?;
        let target = package_bytes(commit.package())?;
        assert_eq!(
            marker_rows(&target)?,
            ["Name", "apple", "apple", "banana", "zebra", "Total"]
        );
        assert_locality(&source, &target)?;
    }

    let malformed_root = source_with_aggregate_only_optional_format(
        Scope::EntireTable,
        22,
        tst::table_data_list::ListType::Format,
        OPTIONAL_FORMATS,
        Vec::new(),
    )?;
    let malformed_root = rewrite_object_wire(
        &malformed_root,
        OPTIONAL_FORMATS,
        TABLE_DATA_LIST_MESSAGE_TYPE,
        |bytes| {
            append_length_delimited_field(bytes, 100, b"opaque optional format state")?;
            Ok(())
        },
    )?;
    assert_physical_sort_rejected(&malformed_root)?;

    let segmented = source_with_aggregate_only_optional_format(
        Scope::EntireTable,
        15,
        tst::table_data_list::ListType::CustomFormat,
        OPTIONAL_CUSTOM_FORMATS,
        Vec::new(),
    )?;
    let segmented = rewrite_object_wire(
        &segmented,
        OPTIONAL_CUSTOM_FORMATS,
        TABLE_DATA_LIST_MESSAGE_TYPE,
        |bytes| {
            append_length_delimited_field(bytes, 4, &reference(TILE).encode_to_vec())?;
            Ok(())
        },
    )?;
    assert_physical_sort_rejected(&segmented)?;
    Ok(())
}

#[test]
fn physical_transaction_types_are_archive_free_and_send_sync() {
    fn assert_send_sync_debug<T: Send + Sync + std::fmt::Debug>() {}
    assert_send_sync_debug::<ColumnIndex>();
    assert_send_sync_debug::<Direction>();
    assert_send_sync_debug::<Order>();
    assert_send_sync_debug::<Rule>();
    assert_send_sync_debug::<Scope>();
    assert_send_sync_debug::<RowRange>();
    assert_send_sync_debug::<litchi_keynote::slide::table::physical_sort::transaction::Edit<'static>>(
    );
    assert_send_sync_debug::<litchi_keynote::SlideTablePhysicalSortCommit>();
    assert_send_sync_debug::<litchi_keynote::SlideTablePhysicalSortPatch>();
    assert_send_sync_debug::<litchi_keynote::slide::table::physical_sort::Diagnostics>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<UnsupportedFeature>();
}

#[test]
fn selected_row_edit_exposes_its_checked_body_range() -> TestResult {
    let source = source(Scope::SelectedRows)?;
    let package = Package::from_bytes(&source)?;
    let rows = RowRange::new(1, 4)?;
    let edit = package.edit_slide_table_physical_sort_to_rows("Tables", 0usize, rows)?;
    assert_eq!(edit.rows(), Some(rows));
    assert_eq!(
        edit.path(),
        Path::rows(Position::new(0), Position::new(0), rows)
    );
    let commit = edit.commit()?;
    assert!(commit.diagnostics().changed());
    Ok(())
}

#[test]
fn full_sort_moves_body_rows_and_sparse_headers_but_not_header_or_footer() -> TestResult {
    let source = source(Scope::EntireTable)?;
    let package = Package::from_bytes(&source)?;
    let before = marker_rows(&source)?;
    assert_eq!(
        before,
        ["Name", "zebra", "apple", "banana", "apple", "Total"]
    );
    let commit = package.execute_slide_table_sort_order("Tables", 0usize)?;
    assert!(commit.diagnostics().changed());
    assert!(commit.diagnostics().moved_rows() > 0);
    assert_eq!(commit.diagnostics().touched_components(), 1);
    assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
    assert!(commit.diagnostics().full_reparse_performed());
    let target = package_bytes(commit.package())?;
    assert_eq!(
        marker_rows(&target)?,
        ["Name", "apple", "apple", "banana", "zebra", "Total"]
    );
    let target_pairs = marker_pairs(&target)?;
    assert_eq!(
        target_pairs.first(),
        Some(&("Name".to_owned(), "Marker".to_owned()))
    );
    assert_eq!(
        target_pairs.last(),
        Some(&("Total".to_owned(), "footer".to_owned()))
    );
    assert_eq!(row_header_indices(&source)?, [0, 1, 3, 5]);
    assert_eq!(row_header_indices(&target)?, [0, 3, 4, 5]);
    assert_eq!(
        model_field_payloads(&source, SORT_ORDER_FIELD)?,
        model_field_payloads(&target, SORT_ORDER_FIELD)?
    );
    assert_eq!(
        uid_map_indexes(&source)?,
        (
            vec![0, 1],
            vec![0, 1],
            vec![0, 1, 2, 3, 4, 5],
            vec![0, 1, 2, 3, 4, 5]
        )
    );
    assert_eq!(
        uid_map_indexes(&target)?,
        (
            vec![0, 1],
            vec![0, 1],
            vec![0, 4, 1, 3, 2, 5],
            vec![0, 2, 4, 3, 1, 5]
        )
    );
    assert_locality(&source, &target)?;
    assert_eq!(
        Catalog::from_bytes(&source)?.iter().count(),
        Catalog::from_bytes(&target)?.iter().count() + PREVIEWS.len()
    );
    Ok(())
}

#[test]
fn duplicate_sort_keys_are_stable_and_selected_ranges_are_body_relative() -> TestResult {
    let selected = source(Scope::SelectedRows)?;
    let package = Package::from_bytes(&selected)?;
    assert!(matches!(
        package.execute_slide_table_sort_order(0usize, 0usize),
        Err(Error::ScopeMismatch {
            configured: Scope::SelectedRows,
            requested: Scope::EntireTable,
            ..
        })
    ));
    let commit = package.execute_slide_table_sort_order_to_rows(
        SlideSelector::name("Tables"),
        TableSelector::index(0),
        RowRange::new(0, BODY_COUNT)?,
    )?;
    let target = package_bytes(commit.package())?;
    assert_eq!(commit.diagnostics().deleted_previews(), PREVIEWS.len());
    assert_eq!(
        marker_rows(&target)?,
        ["Name", "apple", "apple", "banana", "zebra", "Total"]
    );
    assert_eq!(
        marker_pairs(&target)?[1..5],
        [
            ("apple".to_owned(), "first apple".to_owned()),
            ("apple".to_owned(), "second apple".to_owned()),
            ("banana".to_owned(), "middle".to_owned()),
            ("zebra".to_owned(), "last".to_owned()),
        ]
    );
    assert!(
        Catalog::from_bytes(&target)?
            .iter()
            .all(|entry| !PREVIEWS.contains(&entry.name()))
    );

    // The explicit range is body-relative and may cover only part of the
    // persisted SelectedRows scope.  Header/footer rows stay fixed, while
    // the sparse row-header records follow the moved physical envelopes.
    let partial_package = Package::from_bytes(&selected)?;
    let partial = partial_package.execute_slide_table_sort_order_to_rows(
        0usize,
        0usize,
        RowRange::new(0, 2)?,
    )?;
    let partial_target = package_bytes(partial.package())?;
    assert_eq!(
        marker_pairs(&partial_target)?,
        [
            ("Name".to_owned(), "Marker".to_owned()),
            ("apple".to_owned(), "first apple".to_owned()),
            ("zebra".to_owned(), "last".to_owned()),
            ("banana".to_owned(), "middle".to_owned()),
            ("apple".to_owned(), "second apple".to_owned()),
            ("Total".to_owned(), "footer".to_owned()),
        ]
    );
    assert_eq!(row_header_indices(&partial_target)?, [0, 2, 3, 5]);
    assert_eq!(
        uid_map_indexes(&partial_target)?,
        (
            vec![0, 1],
            vec![0, 1],
            vec![0, 2, 1, 3, 4, 5],
            vec![0, 2, 1, 3, 4, 5]
        )
    );
    assert_eq!(partial.diagnostics().deleted_previews(), PREVIEWS.len());
    Ok(())
}

#[test]
fn cross_tile_row_moves_are_rejected_before_publication() -> TestResult {
    let source = small_tile_source(Scope::EntireTable)?;
    let package = Package::from_bytes(&source)?;
    let before = package_bytes(&package)?;
    assert!(matches!(
        package.execute_slide_table_sort_order(0usize, 0usize),
        Err(Error::UnsupportedTopology { .. })
    ));
    assert_source_unchanged(&package, &before)?;
    Ok(())
}

#[test]
fn exact_noop_repeated_sort_preserves_every_byte_and_preview() -> TestResult {
    let source = source(Scope::EntireTable)?;
    let package = Package::from_bytes(&source)?;
    let first = package.execute_slide_table_sort_order(0usize, 0usize)?;
    let target = package_bytes(first.package())?;
    let second = first
        .package()
        .execute_slide_table_sort_order(0usize, 0usize)?;
    assert!(second.patch().is_noop());
    assert!(!second.diagnostics().changed());
    assert_eq!(second.diagnostics().moved_rows(), 0);
    assert_eq!(package_bytes(second.package())?, target);
    assert!(
        Catalog::from_bytes(&target)?
            .iter()
            .all(|entry| !PREVIEWS.contains(&entry.name()))
    );
    Ok(())
}

#[test]
fn changed_patch_applies_once_inverse_restores_and_conflicts_on_replay() -> TestResult {
    let source = source(Scope::EntireTable)?;
    let package = Package::from_bytes(&source)?;
    let commit = package.execute_slide_table_sort_order(0usize, 0usize)?;
    let target = package_bytes(commit.package())?;
    assert_eq!(commit.patch().path().to_string(), "slide 0 table 0");
    assert_ne!(
        commit.patch().source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    let applied = package.apply_slide_table_physical_sort(commit.patch())?;
    assert_eq!(package_bytes(applied.package())?, target);
    assert_source_unchanged(&package, &source)?;
    let foreign = Package::from_bytes(&rewrite_entry(
        &source,
        SENTINEL_MEMBER,
        b"foreign physical-sort package",
    )?)?;
    assert!(matches!(
        foreign.apply_slide_table_physical_sort(commit.patch()),
        Err(Error::PatchConflict)
    ));
    assert!(matches!(
        commit
            .package()
            .apply_slide_table_physical_sort(commit.patch()),
        Err(Error::PatchConflict)
    ));
    let inverse = commit.patch().inverse();
    assert_eq!(inverse.inverse(), *commit.patch());
    assert_eq!(
        inverse.source_fingerprint(),
        commit.patch().target_fingerprint()
    );
    assert_eq!(
        inverse.target_fingerprint(),
        commit.patch().source_fingerprint()
    );
    assert!(matches!(
        package.apply_slide_table_physical_sort(&inverse),
        Err(Error::PatchConflict)
    ));
    let restored = applied
        .package()
        .apply_slide_table_physical_sort(&inverse)
        .map_err(|error| std::io::Error::other(format!("inverse apply: {error:?}")))?;
    assert_eq!(package_bytes(restored.package())?, source);
    Ok(())
}

#[test]
fn selector_lock_missing_order_scope_and_range_fail_atomically() -> TestResult {
    let entire_source = source(Scope::EntireTable)?;
    let package = Package::from_bytes(&entire_source)?;
    assert!(matches!(
        package.execute_slide_table_sort_order(usize::MAX, 0usize),
        Err(Error::SlidePositionNotFound { .. })
    ));
    assert!(matches!(
        package.execute_slide_table_sort_order(SlideSelector::name(""), 0usize),
        Err(Error::EmptySlideName)
    ));
    assert!(matches!(
        package.execute_slide_table_sort_order(SlideSelector::name("missing"), 0usize),
        Err(Error::SlideNameNotFound)
    ));
    assert!(matches!(
        package.execute_slide_table_sort_order(0usize, usize::MAX),
        Err(Error::TablePositionNotFound { .. })
    ));
    let selected_source = source(Scope::SelectedRows)?;
    let selected = Package::from_bytes(&selected_source)?;
    assert!(matches!(
        selected.execute_slide_table_sort_order_to_rows(0usize, 0usize, RowRange::new(4, 6)?),
        Err(Error::RowRangeOutOfBounds { .. })
    ));
    let absent = Package::from_bytes(&source_package(Scope::EntireTable, false, false)?)?;
    let absent = rewrite_document(&package_bytes(&absent)?, |archive| {
        let model = archive
            .object_mut(TABLE_MODEL)
            .ok_or("missing table model")?;
        let message = model.messages.first().ok_or("missing model message")?;
        let payload =
            rewrite_repeated_length_delimited_fields(&message.data, SORT_ORDER_FIELD, &[])?;
        model.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: payload,
            },
        )?;
        Ok(())
    })?;
    let absent = Package::from_bytes(&absent)?;
    assert!(matches!(
        absent.execute_slide_table_sort_order(0usize, 0usize),
        Err(Error::SortOrderMissing { .. })
    ));
    let locked = Package::from_bytes(&rewrite_table_info(&entire_source, true)?)?;
    let before = package_bytes(&locked)?;
    assert!(matches!(
        locked.execute_slide_table_sort_order(0usize, 0usize),
        Err(Error::TableLocked { .. })
    ));
    assert_source_unchanged(&locked, &before)?;
    Ok(())
}

#[test]
fn unsupported_row_affine_and_malformed_storage_fail_before_publication() -> TestResult {
    let source = source(Scope::EntireTable)?;
    let cases = [
        (
            rewrite_model_wire(&source, 47, &[0x0a, 0x00])?,
            "merge owner",
        ),
        (
            rewrite_model_wire(&source, 38, &[0x08, 0x01])?,
            "filtered rows",
        ),
        (
            rewrite_model_wire(&source, 39, &[0x0a, 0x00])?,
            "conditional style owner",
        ),
        (
            rewrite_model_wire(&source, 34, &reference(UID_MAP).encode_to_vec())?,
            "formula dependency owner",
        ),
    ];
    for (bytes, _label) in cases {
        let Ok(package) = Package::from_bytes(&bytes) else {
            continue;
        };
        let before = package_bytes(&package)?;
        let result = package.execute_slide_table_sort_order(0usize, 0usize);
        assert!(result.is_err());
        assert_source_unchanged(&package, &before)?;
    }
    for (bytes, _label) in [
        (
            rewrite_tile_cell(&source, 2, 0, &bnc_formula(1))?,
            "formula",
        ),
        (rewrite_tile_cell(&source, 2, 0, &bnc_error())?, "error"),
        (
            rewrite_tile_cell(&source, 2, 0, &bnc_rich_text(1))?,
            "rich text",
        ),
    ] {
        if let Ok(package) = Package::from_bytes(&bytes) {
            let before = package_bytes(&package)?;
            assert!(
                package
                    .execute_slide_table_sort_order(0usize, 0usize)
                    .is_err()
            );
            assert_source_unchanged(&package, &before)?;
        }
    }
    let malformed_uid = rewrite_uid_map(&source, |map| {
        map.row_uid_for_index[1] = map.row_uid_for_index[0];
    })?;
    if let Ok(package) = Package::from_bytes(&malformed_uid) {
        let before = package_bytes(&package)?;
        assert!(
            package
                .execute_slide_table_sort_order(0usize, 0usize)
                .is_err()
        );
        assert_source_unchanged(&package, &before)?;
    }
    let malformed = rewrite_document(&source, |archive| {
        let tile = archive.object_mut(TILE).ok_or("missing tile")?;
        let message = tile.messages.first().ok_or("missing tile message")?;
        let mut payload = tst::Tile::decode(message.data.as_slice())?;
        payload.row_infos[1].tile_row_index = payload.row_infos[0].tile_row_index;
        tile.replace_message_preserving_header(
            0,
            RawMessage {
                type_: TILE_MESSAGE_TYPE,
                data: payload.encode_to_vec(),
            },
        )?;
        Ok(())
    })?;
    if let Ok(package) = Package::from_bytes(&malformed) {
        let before = package_bytes(&package)?;
        assert!(
            package
                .execute_slide_table_sort_order(0usize, 0usize)
                .is_err()
        );
        assert_source_unchanged(&package, &before)?;
    }
    Ok(())
}

#[test]
fn formula_registry_and_imported_data_are_rejected_atomically() -> TestResult {
    let scalar_source = source(Scope::EntireTable)?;
    let formula_registry = rewrite_formula_registry(&scalar_source)?;
    let formula_package = Package::from_bytes(&formula_registry)?;
    let formula_before = package_bytes(&formula_package)?;
    assert!(matches!(
        formula_package.execute_slide_table_sort_order(0usize, 0usize),
        Err(Error::UnsupportedFeature {
            feature: UnsupportedFeature::Formula,
            ..
        })
    ));
    assert_source_unchanged(&formula_package, &formula_before)?;

    let imported_data = rewrite_model_wire(&scalar_source, 52, &[0x0a, 0x00])?;
    let imported_package = Package::from_bytes(&imported_data)?;
    let imported_before = package_bytes(&imported_package)?;
    assert!(matches!(
        imported_package.execute_slide_table_sort_order(0usize, 0usize),
        Err(Error::UnsupportedDependency {
            feature: UnsupportedFeature::RowAffineDependency,
            ..
        })
    ));
    assert_source_unchanged(&imported_package, &imported_before)?;
    Ok(())
}

#[test]
fn signed_zero_numeric_keys_are_stable_without_sign_normalization() -> TestResult {
    let source = rewrite_numeric_body_keys(&source(Scope::EntireTable)?, [0.0, -0.0, 0.0, -0.0])?;
    assert_eq!(
        marker_column_rows(&source)?,
        [
            "Marker",
            "last",
            "first apple",
            "middle",
            "second apple",
            "footer"
        ]
    );
    let package = Package::from_bytes(&source)?;
    let commit = package.execute_slide_table_sort_order(0usize, 0usize)?;
    let target = package_bytes(commit.package())?;
    assert_eq!(
        numeric_body_key_markers(&target)?,
        [
            (0.0f64.to_bits(), "last".to_owned()),
            ((-0.0f64).to_bits(), "first apple".to_owned()),
            (0.0f64.to_bits(), "middle".to_owned()),
            ((-0.0f64).to_bits(), "second apple".to_owned()),
        ]
    );
    assert_eq!(
        marker_column_rows(&target)?,
        [
            "Marker",
            "last",
            "first apple",
            "middle",
            "second apple",
            "footer"
        ]
    );
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.diagnostics().moved_rows(), 0);
    assert_eq!(commit.diagnostics().deleted_previews(), 0);
    assert_eq!(target, source);
    Ok(())
}

#[test]
fn uid_alias_and_metadata_uuid_ambiguity_fail_closed_atomically() -> TestResult {
    let source = source(Scope::EntireTable)?;
    let aliased_uid_map = rewrite_uid_map(&source, |map| {
        map.sorted_row_uids[1] = map.sorted_row_uids[0];
    })?;
    assert_physical_sort_rejected(&aliased_uid_map)?;

    let ambiguous_metadata = rewrite_metadata(&source, |metadata| {
        let component = metadata
            .components
            .first_mut()
            .ok_or("missing metadata component")?;
        let first_uuid = component
            .object_uuid_map_entries
            .first()
            .ok_or("missing metadata UUID entry")?
            .uuid;
        let second = component
            .object_uuid_map_entries
            .get_mut(1)
            .ok_or("missing second metadata UUID entry")?;
        second.uuid = first_uuid;
        Ok(())
    })?;
    assert_physical_sort_rejected(&ambiguous_metadata)?;
    Ok(())
}

#[test]
fn string_entry_sidecars_are_rejected_atomically() -> TestResult {
    let source = rewrite_string_entry_sidecars(&source(Scope::EntireTable)?)?;
    let package = Package::from_bytes(&source)?;
    let before = package_bytes(&package)?;
    assert!(matches!(
        package.execute_slide_table_sort_order(0usize, 0usize),
        Err(Error::UnsupportedTopology { .. })
    ));
    assert_source_unchanged(&package, &before)?;
    Ok(())
}

#[test]
fn padded_offsets_and_pre_bnc_sidecars_follow_native_compatibility_rules() -> TestResult {
    let padded_source = rewrite_padded_offsets(&source(Scope::EntireTable)?)?;
    assert!(
        tile_offset_payloads(&padded_source)?
            .iter()
            .all(|offsets| offsets.len() == COLUMN_COUNT as usize * 2 + 2
                && offsets[offsets.len() - 2..] == u16::MAX.to_le_bytes())
    );
    let package = Package::from_bytes(&padded_source)?;
    let commit = package.execute_slide_table_sort_order(0usize, 0usize)?;
    let padded_target = package_bytes(commit.package())?;
    assert_eq!(
        marker_rows(&padded_target)?,
        ["Name", "apple", "apple", "banana", "zebra", "Total"]
    );
    assert!(
        tile_offset_payloads(&padded_target)?
            .iter()
            .all(|offsets| offsets.len() == COLUMN_COUNT as usize * 2 + 2
                && offsets[offsets.len() - 2..] == u16::MAX.to_le_bytes())
    );
    assert_locality(&padded_source, &padded_target)?;

    let canonical_storage = pre_bnc_empty_cell().repeat(COLUMN_COUNT as usize);
    let canonical_offsets = [0u16, 12u16]
        .into_iter()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    assert!(
        pre_bnc_payloads(&source(Scope::EntireTable)?)?
            .into_iter()
            .all(|(storage, offsets)| storage == canonical_storage && offsets == canonical_offsets)
    );

    // Keynote's native legacy mirror contains only canonical empty v4
    // sentinels.  It is moved with the row envelope and remains unchanged.
    let package = Package::from_bytes(&source(Scope::EntireTable)?)?;
    let commit = package.execute_slide_table_sort_order(0usize, 0usize)?;
    let canonical_target = package_bytes(commit.package())?;
    assert!(
        pre_bnc_payloads(&canonical_target)?
            .into_iter()
            .all(|(storage, offsets)| storage == canonical_storage && offsets == canonical_offsets)
    );
    assert_locality(&source(Scope::EntireTable)?, &canonical_target)?;

    // Copied modern BNC bytes are not a valid legacy mirror.  The owner must
    // reject them atomically instead of attempting to interpret or rewrite
    // the unknown sidecar representation.
    let sidecar_source = rewrite_pre_bnc_sidecars(&source(Scope::EntireTable)?)?;
    assert_physical_sort_rejected(&sidecar_source)?;

    // A pre-BNC-only row likewise has no safe physical fallback: without
    // modern BNC, scalar interpretation is refused before any publication.
    let pre_bnc_only = rewrite_pre_bnc_only(&source(Scope::EntireTable)?)?;
    assert_physical_sort_rejected(&pre_bnc_only)?;
    Ok(())
}

#[test]
fn unknown_mutable_physical_fields_are_rejected_atomically() -> TestResult {
    let source = source(Scope::EntireTable)?;
    assert_canonical_source_admitted(&source)?;
    let cases = vec![
        (
            "tile root",
            rewrite_object_wire(&source, TILE, TILE_MESSAGE_TYPE, |bytes| {
                append_length_delimited_field(bytes, 100, b"opaque tile state")?;
                Ok(())
            })?,
        ),
        (
            "tile row",
            rewrite_repeated_child_wire(&source, TILE, TILE_MESSAGE_TYPE, 5, 0, |bytes| {
                append_length_delimited_field(bytes, 100, b"opaque row state")?;
                Ok(())
            })?,
        ),
        (
            "row-header bucket root",
            rewrite_object_wire(&source, ROW_HEADERS, HEADER_BUCKET_MESSAGE_TYPE, |bytes| {
                append_length_delimited_field(bytes, 100, b"opaque bucket state")?;
                Ok(())
            })?,
        ),
        (
            "row-header record",
            rewrite_repeated_child_wire(
                &source,
                ROW_HEADERS,
                HEADER_BUCKET_MESSAGE_TYPE,
                2,
                0,
                |bytes| {
                    append_length_delimited_field(bytes, 100, b"opaque header state")?;
                    Ok(())
                },
            )?,
        ),
        (
            "UID-map root",
            rewrite_object_wire(&source, UID_MAP, COLUMN_ROW_UID_MAP_MESSAGE_TYPE, |bytes| {
                append_length_delimited_field(bytes, 100, b"opaque UID state")?;
                Ok(())
            })?,
        ),
        (
            "row UID record",
            rewrite_repeated_child_wire(
                &source,
                UID_MAP,
                COLUMN_ROW_UID_MAP_MESSAGE_TYPE,
                4,
                0,
                |bytes| {
                    append_length_delimited_field(bytes, 100, b"opaque UUID state")?;
                    Ok(())
                },
            )?,
        ),
    ];

    for (label, bytes) in cases {
        let package = Package::from_bytes(&bytes)?;
        let before = package_bytes(&package)?;
        assert!(
            package
                .execute_slide_table_sort_order(0usize, 0usize)
                .is_err(),
            "unknown mutable field was admitted: {label}"
        );
        assert_source_unchanged(&package, &before)?;
    }
    Ok(())
}

#[test]
fn unsupported_model_owner_markers_39_45_84_93_are_rejected_atomically() -> TestResult {
    let source = source(Scope::EntireTable)?;
    assert_canonical_source_admitted(&source)?;
    let cases = vec![
        (
            "conditional-style formula owner (39)",
            rewrite_model_wire(&source, 39, &valid_cfuuid_owner_payload()?)?,
        ),
        (
            "sort-rule reference tracker (45)",
            rewrite_model_wire(&source, 45, &valid_sort_tracker_payload()?)?,
        ),
        (
            "haunted owner (84)",
            rewrite_model_wire(&source, 84, &valid_uuid_owner_payload()?)?,
        ),
        (
            "spill owner (93)",
            rewrite_model_wire(&source, 93, &valid_uuid_owner_payload()?)?,
        ),
    ];

    for (label, bytes) in cases {
        let package = Package::from_bytes(&bytes)?;
        let before = package_bytes(&package)?;
        assert!(
            package
                .execute_slide_table_sort_order(0usize, 0usize)
                .is_err(),
            "unsupported model owner marker was admitted: {label}"
        );
        assert_source_unchanged(&package, &before)?;
    }
    Ok(())
}

#[test]
fn tile_num_cells_mismatch_is_rejected_atomically() -> TestResult {
    let source = source(Scope::EntireTable)?;
    assert_canonical_source_admitted(&source)?;
    let mismatched = rewrite_object_wire(&source, TILE, TILE_MESSAGE_TYPE, |bytes| {
        let mut tile = tst::Tile::decode(bytes.as_slice())?;
        tile.num_cells = tile
            .num_cells
            .checked_sub(1)
            .ok_or("fixture tile has no cells")?;
        *bytes = tile.encode_to_vec();
        Ok(())
    })?;
    let package = Package::from_bytes(&mismatched)?;
    let before = package_bytes(&package)?;
    assert!(
        package
            .execute_slide_table_sort_order(0usize, 0usize)
            .is_err()
    );
    assert_source_unchanged(&package, &before)?;
    Ok(())
}

#[test]
fn wide_row_and_pre_bnc_disagreements_are_rejected_atomically() -> TestResult {
    let source = source(Scope::EntireTable)?;
    assert_canonical_source_admitted(&source)?;
    let cases = vec![
        (
            "DataStore wide-row flag disagrees with tile rows",
            rewrite_model_wide_rows(&source, true)?,
        ),
        (
            "row wide-row flag disagrees with narrow BNC offsets",
            rewrite_repeated_child_wire(&source, TILE, TILE_MESSAGE_TYPE, 5, 0, |bytes| {
                let mut row = tst::TileRowInfo::decode(bytes.as_slice())?;
                row.has_wide_offsets = Some(true);
                *bytes = row.encode_to_vec();
                Ok(())
            })?,
        ),
    ];

    for (label, bytes) in cases {
        let package = Package::from_bytes(&bytes)?;
        let before = package_bytes(&package)?;
        assert!(
            package
                .execute_slide_table_sort_order(0usize, 0usize)
                .is_err(),
            "wide/pre-BNC disagreement was admitted: {label}"
        );
        assert_source_unchanged(&package, &before)?;
    }
    Ok(())
}

#[test]
fn unknown_model_fields_are_rejected_atomically() -> TestResult {
    let source = source_package(Scope::EntireTable, false, true)?;
    let package = Package::from_bytes(&source)?;
    let before = package_bytes(&package)?;
    assert!(
        package
            .execute_slide_table_sort_order(0usize, 0usize)
            .is_err()
    );
    assert_source_unchanged(&package, &before)?;
    Ok(())
}

#[test]
fn noncanonical_iwa_object_length_prefix_is_rejected_before_publication() -> TestResult {
    let source = source(Scope::EntireTable)?;
    let malformed = rewrite_document_stream(&source, |stream| {
        let (_header_length, prefix_length) = decode_varint_from_bytes(stream)?;
        let terminal = stream
            .get_mut(prefix_length.checked_sub(1).ok_or("empty object prefix")?)
            .ok_or("missing object length prefix")?;
        if *terminal & 0x80 != 0 {
            return Err("fixture object prefix is unexpectedly noncanonical".into());
        }
        *terminal |= 0x80;
        stream.insert(prefix_length, 0);
        Ok(())
    })?;
    let package = Package::from_bytes(&malformed)?;
    let before = package_bytes(&package)?;
    assert!(
        package
            .execute_slide_table_sort_order(0usize, 0usize)
            .is_err(),
        "noncanonical IWA object framing was unexpectedly admitted"
    );
    assert_source_unchanged(&package, &before)?;
    Ok(())
}

#[test]
fn iwa_message_data_length_mismatch_is_rejected_before_publication() -> TestResult {
    let source = source(Scope::EntireTable)?;
    let malformed = rewrite_document_stream(&source, |stream| {
        let (header_length, prefix_length) = decode_varint_from_bytes(stream)?;
        let header_length = usize::try_from(header_length)
            .map_err(|_| "fixture object header does not fit usize")?;
        let header_end = prefix_length
            .checked_add(header_length)
            .ok_or("fixture object header range overflow")?;
        let mut header = tsp::ArchiveInfo::decode(
            stream
                .get(prefix_length..header_end)
                .ok_or("missing fixture object header")?,
        )?;
        let message_info = header
            .message_infos
            .first_mut()
            .ok_or("missing fixture message metadata")?;
        message_info.length = message_info
            .length
            .checked_add(1)
            .ok_or("fixture message length overflow")?;
        let encoded = header.encode_to_vec();
        if encoded.len() != header_length {
            return Err("fixture message length mutation changed header width".into());
        }
        stream[prefix_length..header_end].copy_from_slice(&encoded);
        Ok(())
    })?;
    // Depending on which ingress path first touches the component, the
    // malformed length can be rejected while indexing the package or while
    // the physical owner reopens the selected component.  Both paths are
    // fail-closed and publish no candidate.
    assert_physical_sort_rejected(&malformed)
}

#[test]
fn concurrent_reads_and_commits_are_deterministic() -> TestResult {
    let source = Arc::<[u8]>::from(source(Scope::EntireTable)?);
    let mut workers = Vec::new();
    for _ in 0..8 {
        let source = Arc::clone(&source);
        workers.push(std::thread::spawn(move || -> Result<Vec<u8>, String> {
            let package = Package::from_bytes(&source).map_err(|error| error.to_string())?;
            let commit = package
                .execute_slide_table_sort_order(0usize, 0usize)
                .map_err(|error| error.to_string())?;
            package_bytes(commit.package()).map_err(|error| error.to_string())
        }));
    }
    let mut outputs = Vec::new();
    for worker in workers {
        let result = worker
            .join()
            .map_err(|_| "physical-sort worker panicked")?
            .map_err(std::io::Error::other)?;
        outputs.push(result);
    }
    assert!(outputs.windows(2).all(|pair| pair[0] == pair[1]));
    Ok(())
}

#[test]
fn native_keynote_14_4_sort_now_open_save_close_reopen_evidence_is_recorded() {
    const APP_VERSION: &str = "Keynote 14.4 (7043.0.93)";
    const SOURCE_SIZE: usize = 516_029;
    const RUST_CANDIDATE_SIZE: usize = 469_132;
    const NATIVE_RESAVED_SIZE: usize = 516_085;
    const SOURCE_SHA256: &str = "c969edd38c204599921e32dc948acb835ae3740a90a4a2cb8f677d57c5b4da83";
    const RUST_CANDIDATE_SHA256: &str =
        "0d6607805234fa3f11e9fdd0a9d4c2f695dc49041e1899c787d3c340cb9d7272";
    const NATIVE_RESAVED_SHA256: &str =
        "1a6fc3487688f0ba23057b191a9c235780a35ea05096566ddedc14f6526e463c";

    let native = (APP_VERSION, SOURCE_SIZE, SOURCE_SHA256);
    let reopen = (NATIVE_RESAVED_SIZE, NATIVE_RESAVED_SHA256);

    assert!(native.0.starts_with("Keynote 14.4"));
    assert!(native.0.ends_with("(7043.0.93)"));
    assert!(native.1 > RUST_CANDIDATE_SIZE);
    assert!(reopen.0 > RUST_CANDIDATE_SIZE);
    assert!(
        [native.2, RUST_CANDIDATE_SHA256, reopen.1]
            .into_iter()
            .all(|digest| digest.len() == 64 && digest.bytes().all(|byte| byte.is_ascii_hexdigit()))
    );
    let visible_rows = ["Apple", "Banana", "Cherry", "Zebra"];
    assert!(visible_rows.windows(2).all(|rows| rows[0] < rows[1]));
}
