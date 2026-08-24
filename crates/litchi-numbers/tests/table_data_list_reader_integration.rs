//! Integration coverage for the generated-free outer Numbers table-data-list
//! projection.
//!
//! The package graph below is intentionally small and bounded.  It uses the
//! generated protobuf types only as a test oracle; production ingress must
//! retain the raw bytes and project the selected outer list through the strict
//! generated-free codec.

use std::error::Error as StdError;
use std::sync::Arc;

use litchi_iwa_archive::Limits as ArchiveLimits;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::tsce::ast_node_array_archive::{AstNodeArchive, AstNodeType};
use litchi_iwa_protos::{tn, tsce, tsd, tsp, tst, tswp};
use litchi_numbers::cell::Value;
use litchi_numbers::{
    CellPosition, Document, Package, PackageLimits, PackageReadOptions, PackageSemanticLimits,
    TableCellCommentError, TableCellCommentPath,
};
use litchi_numbers_wire::BncCell;
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHEET_MESSAGE_TYPE: u32 = 2;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;
const TABLE_DATA_LIST_ALT_MESSAGE_TYPE: u32 = 6_201;
const TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE: u32 = 6_011;
const RICH_TEXT_PAYLOAD_MESSAGE_TYPE: u32 = 6_218;
const STORAGE_MESSAGE_TYPE: u32 = 2_001;
const COMMENT_STORAGE_MESSAGE_TYPE: u32 = 3_056;
const FORMULA_ERROR_FLAG: u32 = 0x0000_0800;

const ROOT_ID: u64 = 1;
const SHEET_ID: u64 = 2;
const TABLE_INFO_ID: u64 = 3;
const TABLE_MODEL_ID: u64 = 4;
const SIDECAR_ID: u64 = 5;
const TILE_ID: u64 = 6;
const RICH_TEXT_PAYLOAD_ID: u64 = 101;
const RICH_TEXT_STORAGE_ID: u64 = 102;
const COMMENT_STORAGE_ID: u64 = 111;
const ISOLATED_COMMENT_STORAGE_ID: u64 = 114;
const SEGMENT_COMMENT_STORAGE_ID: u64 = 115;
const COMMENT_REPLY_ID: u64 = 112;
const COMMENT_REPLY_TWO_ID: u64 = 113;
const COMMENT_AUTHOR_ID: u64 = 121;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Corruption {
    None,
    MalformedLaterSegmentPayload,
    MissingSegmentPayload,
    DuplicateSegmentPayload,
    MissingSegmentReference,
    DuplicateSegmentReference,
    WrongSegmentType,
    RangeOverflow,
    EntryOutsideRange,
    DuplicateRootKey,
    DuplicateSegmentKey,
    SelectedMalformedList,
    NonCanonicalRootList,
    CommentMissingText,
    CommentMalformedNestedWire,
    CommentDuplicateCanonicalPayload,
    CommentDuplicateMalformedPayload,
}

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
}

fn bounded<T>(items: impl IntoIterator<Item = T>, maximum: usize) -> TestResult<Vec<T>> {
    let mut values = Vec::new();
    values.try_reserve(maximum)?;
    for value in items {
        if values.len() == maximum {
            return Err(std::io::Error::other("synthetic test builder bound exceeded").into());
        }
        values.push(value);
    }
    Ok(values)
}

fn object(identifier: u64, message_type: u32, data: Vec<u8>) -> TestResult<ArchiveObject> {
    Ok(ArchiveObject::new(
        identifier,
        vec![RawMessage {
            type_: message_type,
            data,
        }],
    )?)
}

fn string_entry(key: u32, value: &str) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount: 1,
        string: Some(value.to_owned()),
        ..Default::default()
    }
}

fn formula_archive() -> tsce::FormulaArchive {
    tsce::FormulaArchive {
        ast_node_array: tsce::AstNodeArrayArchive {
            ast_node: vec![AstNodeArchive {
                ast_node_type: AstNodeType::NumberNode as i32,
                ast_number_node_number: Some(203.0),
                ..Default::default()
            }],
        },
        ..Default::default()
    }
}

fn formula_entry(key: u32) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount: 1,
        formula: Some(formula_archive()),
        ..Default::default()
    }
}

fn formula_error_entry(key: u32) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount: 1,
        string: Some("#VALUE!".to_owned()),
        ..Default::default()
    }
}

fn rich_text_entry(key: u32) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount: 1,
        rich_text_payload: Some(reference(RICH_TEXT_PAYLOAD_ID)),
        ..Default::default()
    }
}

fn comment_entry(key: u32) -> tst::table_data_list::ListEntry {
    comment_entry_for_storage(key, COMMENT_STORAGE_ID)
}

fn comment_entry_for_storage(key: u32, storage_identifier: u64) -> tst::table_data_list::ListEntry {
    tst::table_data_list::ListEntry {
        key,
        refcount: 1,
        comment_storage: Some(reference(storage_identifier)),
        ..Default::default()
    }
}

fn formula_error_cell(identifier: u32) -> TestResult<Vec<u8>> {
    // BNC cells have an eight-byte prefix and a four-byte field mask before
    // fields in the canonical layout.  The test-only raw construction keeps
    // the wire fixture independent of private production constants while
    // exercising the public BncCell parser on the resulting bytes.
    let mut bytes = BncCell::minimal().encode();
    bytes[1] = 8;
    bytes[8..12].copy_from_slice(&FORMULA_ERROR_FLAG.to_le_bytes());
    bytes.extend_from_slice(&identifier.to_le_bytes());
    BncCell::parse(&bytes)?;
    Ok(bytes)
}

fn segment_entries(
    list_type: tst::table_data_list::ListType,
    key: u32,
) -> Vec<tst::table_data_list::ListEntry> {
    match list_type {
        tst::table_data_list::ListType::String => vec![string_entry(
            key,
            match key {
                4 => "String segment one",
                7 => "String segment two",
                _ => "String segment unused",
            },
        )],
        tst::table_data_list::ListType::Formula => vec![formula_entry(key)],
        tst::table_data_list::ListType::FormulaError => vec![formula_error_entry(key)],
        tst::table_data_list::ListType::RichTextPayload => vec![rich_text_entry(key)],
        tst::table_data_list::ListType::CommentStorage => vec![comment_entry_for_storage(
            key,
            if key == 12 {
                SEGMENT_COMMENT_STORAGE_ID
            } else {
                COMMENT_STORAGE_ID
            },
        )],
        _ => Vec::new(),
    }
}

fn segment_object(
    identifier: u64,
    list_type: tst::table_data_list::ListType,
    key: u32,
    corruption: Corruption,
) -> TestResult<ArchiveObject> {
    if matches!(
        corruption,
        Corruption::MalformedLaterSegmentPayload if identifier == 51
    ) {
        return object(identifier, TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE, vec![0xff]);
    }
    if matches!(corruption, Corruption::MissingSegmentPayload if identifier == 51) {
        return object(identifier, 9_999, vec![0]);
    }

    let segment_list_type = if matches!(
        corruption,
        Corruption::WrongSegmentType if identifier == 51
    ) {
        tst::table_data_list::ListType::Formula
    } else {
        list_type
    };
    let segment_key = if matches!(
        corruption,
        Corruption::EntryOutsideRange if identifier == 51
    ) {
        key.saturating_add(8)
    } else if matches!(
        corruption,
        Corruption::DuplicateSegmentKey if identifier == 51
    ) {
        7
    } else {
        key
    };
    let range_location = if matches!(
        corruption,
        Corruption::DuplicateSegmentKey if identifier == 51
    ) {
        7
    } else {
        key
    };
    let key_range = if matches!(corruption, Corruption::RangeOverflow if identifier == 51) {
        tsp::Range {
            location: u32::MAX,
            length: 1,
        }
    } else {
        tsp::Range {
            location: range_location,
            length: 1,
        }
    };
    let segment = tst::TableDataListSegment {
        list_type: segment_list_type as i32,
        key_range,
        entries: segment_entries(list_type, segment_key),
    };
    let data = segment.encode_to_vec();
    if matches!(corruption, Corruption::DuplicateSegmentPayload if identifier == 51) {
        return Ok(ArchiveObject::new(
            identifier,
            bounded(
                [
                    RawMessage {
                        type_: TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE,
                        data: data.clone(),
                    },
                    RawMessage {
                        type_: TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE,
                        data,
                    },
                ],
                2,
            )?,
        )?);
    }
    object(identifier, TABLE_DATA_LIST_SEGMENT_MESSAGE_TYPE, data)
}

fn string_segment_references(corruption: Corruption) -> Vec<tsp::Reference> {
    match corruption {
        Corruption::MissingSegmentReference => vec![reference(52), reference(9_999)],
        Corruption::DuplicateSegmentReference => vec![reference(52), reference(52)],
        _ => vec![reference(52), reference(51)],
    }
}

fn list_message(
    list_type: tst::table_data_list::ListType,
    entries: Vec<tst::table_data_list::ListEntry>,
    segments: Vec<tsp::Reference>,
) -> tst::TableDataList {
    tst::TableDataList {
        list_type: list_type as i32,
        next_list_id: 120,
        entries,
        segments,
        ..Default::default()
    }
}

fn sidecar_object_with_message_type(
    corruption: Corruption,
    table_data_list_message_type: u32,
) -> TestResult<ArchiveObject> {
    let mut string_entries = bounded(
        [
            string_entry(17, "String root unused"),
            string_entry(5, "String root"),
        ],
        2,
    )?;
    if matches!(corruption, Corruption::DuplicateRootKey) {
        string_entries[1].key = string_entries[0].key;
    }
    let string_list = list_message(
        tst::table_data_list::ListType::String,
        string_entries,
        string_segment_references(corruption),
    );

    let formula_list = list_message(
        tst::table_data_list::ListType::Formula,
        bounded([formula_entry(21), formula_entry(3)], 2)?,
        bounded([reference(62), reference(61)], 2)?,
    );
    let formula_error_list = list_message(
        tst::table_data_list::ListType::FormulaError,
        bounded([formula_error_entry(30), formula_error_entry(2)], 2)?,
        bounded([reference(72), reference(71)], 2)?,
    );
    let rich_text_list = list_message(
        tst::table_data_list::ListType::RichTextPayload,
        bounded([rich_text_entry(40), rich_text_entry(1)], 2)?,
        bounded([reference(82), reference(81)], 2)?,
    );
    let comment_list = list_message(
        tst::table_data_list::ListType::CommentStorage,
        bounded(
            [
                comment_entry_for_storage(50, ISOLATED_COMMENT_STORAGE_ID),
                comment_entry(2),
            ],
            2,
        )?,
        bounded([reference(92), reference(91)], 2)?,
    );

    let mut payloads = bounded(
        [
            RawMessage {
                type_: table_data_list_message_type,
                data: string_list.encode_to_vec(),
            },
            RawMessage {
                type_: table_data_list_message_type,
                data: formula_list.encode_to_vec(),
            },
            RawMessage {
                type_: table_data_list_message_type,
                data: formula_error_list.encode_to_vec(),
            },
            RawMessage {
                type_: table_data_list_message_type,
                data: rich_text_list.encode_to_vec(),
            },
            RawMessage {
                type_: table_data_list_message_type,
                data: comment_list.encode_to_vec(),
            },
        ],
        5,
    )?;
    if matches!(corruption, Corruption::SelectedMalformedList) {
        payloads[0].data = vec![0xff];
    }
    if matches!(corruption, Corruption::NonCanonicalRootList) {
        // Prost accepts a duplicate scalar field with last-wins semantics;
        // the production strict list codec rejects this non-canonical wire.
        payloads[0].data.extend_from_slice(&[0x08, 0x01]);
    }
    Ok(ArchiveObject::new(SIDECAR_ID, payloads)?)
}

fn rich_text_payload_object() -> TestResult<ArchiveObject> {
    let payload = tst::RichTextPayloadArchive {
        storage: reference(RICH_TEXT_STORAGE_ID),
        range: None,
        cellid: tst::CellId {
            packed_data: 1,
            expanded_coord: None,
        },
    };
    let mut result = object(
        RICH_TEXT_PAYLOAD_ID,
        RICH_TEXT_PAYLOAD_MESSAGE_TYPE,
        payload.encode_to_vec(),
    )?;
    let message = result
        .archive_info
        .message_infos
        .first_mut()
        .ok_or_else(|| std::io::Error::other("rich-text payload message missing"))?;
    message.object_references.try_reserve(1)?;
    message.object_references.push(RICH_TEXT_STORAGE_ID);
    Ok(result)
}

fn rich_text_storage_object() -> TestResult<ArchiveObject> {
    object(
        RICH_TEXT_STORAGE_ID,
        STORAGE_MESSAGE_TYPE,
        tswp::StorageArchive {
            kind: Some(tswp::storage_archive::KindType::Cell as i32),
            text: bounded(["Rich segment".to_owned()], 1)?,
            ..Default::default()
        }
        .encode_to_vec(),
    )
}

fn comment_storage_object(corruption: Corruption) -> TestResult<ArchiveObject> {
    // Keep this fixture deliberately richer than the semantic Numbers
    // document currently exposes. Package extraction must still strictly
    // validate and stage every canonical comment fact (including the reply
    // graph) before publication, while Document extraction skips this
    // format-owned sidecar entirely.
    if matches!(corruption, Corruption::CommentMalformedNestedWire) {
        // Field 2 is a length-delimited TSP.Date. The nested payload is an
        // invalid wire tag, so a strict comment-storage decoder must reject
        // the source before publishing a Package snapshot.
        return object(
            COMMENT_STORAGE_ID,
            COMMENT_STORAGE_MESSAGE_TYPE,
            [0x12, 0x01, 0xff].to_vec(),
        );
    }

    let comment = tsd::CommentStorageArchive {
        text: (!matches!(corruption, Corruption::CommentMissingText))
            .then(|| "Comment retained by Package".to_owned()),
        creation_date: Some(tsp::Date { seconds: 123.5 }),
        author: Some(reference(COMMENT_AUTHOR_ID)),
        replies: vec![reference(COMMENT_REPLY_ID), reference(COMMENT_REPLY_TWO_ID)],
        storage_uuid: Some(tsp::Uuid {
            lower: COMMENT_STORAGE_ID,
            upper: 0x6c69_7463_6869_6977,
        }),
    };
    let data = comment.encode_to_vec();
    if matches!(corruption, Corruption::CommentDuplicateCanonicalPayload) {
        return Ok(ArchiveObject::new(
            COMMENT_STORAGE_ID,
            bounded(
                [
                    RawMessage {
                        type_: COMMENT_STORAGE_MESSAGE_TYPE,
                        data: data.clone(),
                    },
                    RawMessage {
                        type_: COMMENT_STORAGE_MESSAGE_TYPE,
                        data,
                    },
                ],
                2,
            )?,
        )?);
    }
    if matches!(corruption, Corruption::CommentDuplicateMalformedPayload) {
        return Ok(ArchiveObject::new(
            COMMENT_STORAGE_ID,
            bounded(
                [
                    RawMessage {
                        type_: COMMENT_STORAGE_MESSAGE_TYPE,
                        data,
                    },
                    RawMessage {
                        type_: COMMENT_STORAGE_MESSAGE_TYPE,
                        // Keep the valid candidate first so this fixture
                        // exercises duplicate handling independently of
                        // candidate ordering.
                        data: vec![0xff],
                    },
                ],
                2,
            )?,
        )?);
    }
    object(COMMENT_STORAGE_ID, COMMENT_STORAGE_MESSAGE_TYPE, data)
}

fn isolated_comment_storage_object() -> TestResult<ArchiveObject> {
    object(
        ISOLATED_COMMENT_STORAGE_ID,
        COMMENT_STORAGE_MESSAGE_TYPE,
        tsd::CommentStorageArchive {
            text: Some("Comment retained by Package".to_owned()),
            storage_uuid: Some(tsp::Uuid {
                lower: ISOLATED_COMMENT_STORAGE_ID,
                upper: 0x6973_6f6c_6174_6564,
            }),
            ..Default::default()
        }
        .encode_to_vec(),
    )
}

fn segmented_comment_storage_object() -> TestResult<ArchiveObject> {
    object(
        SEGMENT_COMMENT_STORAGE_ID,
        COMMENT_STORAGE_MESSAGE_TYPE,
        tsd::CommentStorageArchive {
            text: Some("Comment retained by Package".to_owned()),
            storage_uuid: Some(tsp::Uuid {
                lower: SEGMENT_COMMENT_STORAGE_ID,
                upper: 0x7365_676d_656e_7464,
            }),
            ..Default::default()
        }
        .encode_to_vec(),
    )
}

fn comment_reply_storage_object(identifier: u64, text: &str) -> TestResult<ArchiveObject> {
    object(
        identifier,
        COMMENT_STORAGE_MESSAGE_TYPE,
        tsd::CommentStorageArchive {
            text: Some(text.to_owned()),
            storage_uuid: Some(tsp::Uuid {
                lower: identifier,
                upper: identifier.rotate_left(17),
            }),
            ..Default::default()
        }
        .encode_to_vec(),
    )
}

fn table_model() -> tst::TableModelArchive {
    tst::TableModelArchive {
        table_id: "table-data-list-table-id".to_owned(),
        table_name: "Mixed table-data-list table".to_owned(),
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
            formula_error_table: Some(reference(SIDECAR_ID)),
            rich_text_table: Some(reference(SIDECAR_ID)),
            comment_storage_table: Some(reference(SIDECAR_ID)),
            format_table_pre_bnc: reference(SIDECAR_ID),
            next_row_strip_id: 1,
            next_column_strip_id: 1,
            row_tile_tree: tst::TableRbTree::default(),
            column_tile_tree: tst::TableRbTree::default(),
            ..Default::default()
        },
        ..Default::default()
    }
}

fn table_tile(comment_identifier: u32) -> TestResult<tst::Tile> {
    let mut string_cell = BncCell::minimal();
    string_cell.set_string(7);
    let mut formula_cell = BncCell::minimal();
    formula_cell.set_formula_reference(7);
    let error_cell = formula_error_cell(8)?;
    let mut rich_text_cell = BncCell::minimal();
    rich_text_cell.set_rich_text(6);
    let mut number_cell = BncCell::minimal();
    number_cell.set_number(42.0)?;
    number_cell.set_comment_identifier(Some(comment_identifier));

    let rows = [
        string_cell.encode(),
        formula_cell.encode(),
        error_cell,
        rich_text_cell.encode(),
        number_cell.encode(),
    ];
    let mut row_infos = Vec::new();
    row_infos.try_reserve(rows.len())?;
    for (row, storage) in rows.into_iter().enumerate() {
        row_infos.push(tst::TileRowInfo {
            tile_row_index: u32::try_from(row)?,
            cell_count: 1,
            storage_version: Some(5),
            cell_storage_buffer_pre_bnc: storage.clone(),
            cell_offsets_pre_bnc: vec![0, 0],
            cell_storage_buffer: Some(storage),
            cell_offsets: Some(vec![0, 0]),
            ..Default::default()
        });
    }
    Ok(tst::Tile {
        max_column: 0,
        max_row: 4,
        num_cells: 5,
        numrows: 5,
        row_infos,
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        ..Default::default()
    })
}

fn synthetic_table_data_list_package(
    corruption: Corruption,
    comment_identifier: u32,
) -> TestResult<Vec<u8>> {
    synthetic_table_data_list_package_with_message_type(
        corruption,
        comment_identifier,
        TABLE_DATA_LIST_MESSAGE_TYPE,
    )
}

fn synthetic_table_data_list_package_with_message_type(
    corruption: Corruption,
    comment_identifier: u32,
    table_data_list_message_type: u32,
) -> TestResult<Vec<u8>> {
    let mut objects = Vec::new();
    objects.try_reserve(20)?;
    objects.push(object(
        ROOT_ID,
        DOCUMENT_MESSAGE_TYPE,
        tn::DocumentArchive {
            sheets: bounded([reference(SHEET_ID)], 1)?,
            ..Default::default()
        }
        .encode_to_vec(),
    )?);
    objects.push(object(
        SHEET_ID,
        SHEET_MESSAGE_TYPE,
        tn::SheetArchive {
            name: "Mixed table-data-list sheet".to_owned(),
            drawable_infos: bounded([reference(TABLE_INFO_ID)], 1)?,
            ..Default::default()
        }
        .encode_to_vec(),
    )?);
    objects.push(object(
        TABLE_INFO_ID,
        TABLE_INFO_MESSAGE_TYPE,
        tst::TableInfoArchive {
            table_model: reference(TABLE_MODEL_ID),
            ..Default::default()
        }
        .encode_to_vec(),
    )?);
    objects.push(object(
        TABLE_MODEL_ID,
        TABLE_MODEL_MESSAGE_TYPE,
        table_model().encode_to_vec(),
    )?);
    objects.push(sidecar_object_with_message_type(
        corruption,
        table_data_list_message_type,
    )?);

    let tile = table_tile(comment_identifier)?;
    objects.push(object(TILE_ID, TILE_MESSAGE_TYPE, tile.encode_to_vec())?);

    for (identifier, list_type, key) in [
        (51, tst::table_data_list::ListType::String, 4),
        (52, tst::table_data_list::ListType::String, 7),
        (61, tst::table_data_list::ListType::Formula, 4),
        (62, tst::table_data_list::ListType::Formula, 7),
        (71, tst::table_data_list::ListType::FormulaError, 4),
        (72, tst::table_data_list::ListType::FormulaError, 8),
        (81, tst::table_data_list::ListType::RichTextPayload, 3),
        (82, tst::table_data_list::ListType::RichTextPayload, 6),
        (91, tst::table_data_list::ListType::CommentStorage, 9),
        (92, tst::table_data_list::ListType::CommentStorage, 12),
    ] {
        objects.push(segment_object(identifier, list_type, key, corruption)?);
    }
    objects.push(rich_text_payload_object()?);
    objects.push(rich_text_storage_object()?);
    objects.push(comment_storage_object(corruption)?);
    objects.push(isolated_comment_storage_object()?);
    objects.push(segmented_comment_storage_object()?);
    objects.push(comment_reply_storage_object(
        COMMENT_REPLY_ID,
        "First reply",
    )?);
    objects.push(comment_reply_storage_object(
        COMMENT_REPLY_TWO_ID,
        "Second reply",
    )?);

    let archive = Archive { objects };
    let iwa = SnappyStream::compress(&archive.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [("Index/Document.iwa", iwa.as_slice())],
        ArchiveLimits::default(),
    )?)
}

fn assert_semantics(document: &Document) -> TestResult {
    document.validate()?;
    let sheet = document
        .sheet(0)?
        .ok_or_else(|| std::io::Error::other("synthetic table-data-list sheet is missing"))?;
    assert_eq!(sheet.name(), "Mixed table-data-list sheet");
    let table = sheet
        .tables()
        .next()
        .ok_or_else(|| std::io::Error::other("synthetic table-data-list table is missing"))?;
    assert_eq!(table.name(), "Mixed table-data-list table");
    assert_eq!((table.row_count(), table.column_count()), (5, 1));
    assert_eq!(table.cell_count(), 5);
    assert!(matches!(
        table.get_a1("A1")?,
        Some(Value::Text(value)) if value == "String segment two"
    ));
    assert!(matches!(
        table.get_a1("A2")?,
        Some(Value::Formula(value)) if value == "=203"
    ));
    assert!(matches!(
        table.get_a1("A3")?,
        Some(Value::Error(value)) if value == "#VALUE!"
    ));
    assert!(matches!(
        table.get_a1("A4")?,
        Some(Value::Text(value)) if value == "Rich segment"
    ));
    assert!(matches!(
        table.get_a1("A5")?,
        Some(Value::Number(value)) if value.get().to_bits() == 42.0_f64.to_bits()
    ));
    assert_eq!(
        document.plain_text()?,
        "Mixed table-data-list sheet\nMixed table-data-list table\nString segment two\n=203\nERROR: #VALUE!\nRich segment\n42"
    );
    Ok(())
}

fn assert_rejected_by_both_readers(corruption: Corruption) -> TestResult {
    let bytes = synthetic_table_data_list_package(corruption, 12)?;
    let original = bytes.clone();
    assert!(Package::from_bytes(&bytes).is_err());
    assert_eq!(
        bytes, original,
        "Package parsing mutated its borrowed source"
    );
    assert!(Document::from_bytes(&bytes).is_err());
    assert_eq!(
        bytes, original,
        "Document parsing mutated its borrowed source"
    );
    Ok(())
}

#[test]
fn mixed_root_and_segments_project_all_selected_lists_through_package_and_reopen() -> TestResult {
    let bytes = synthetic_table_data_list_package(Corruption::None, 12)?;
    let original = bytes.clone();

    let package = Package::from_bytes(&bytes)?;
    assert_semantics(package.document())?;
    assert_eq!(package.sheets(), package.document().sheets());
    assert_eq!(
        bytes, original,
        "Package parsing mutated its borrowed source"
    );

    let borrowed = Document::from_bytes(&bytes)?;
    assert_semantics(&borrowed)?;
    assert_eq!(
        bytes, original,
        "Document parsing mutated its borrowed source"
    );

    let reopened = Document::from_shared_bytes(Arc::from(bytes.clone()))?;
    assert_semantics(&reopened)?;
    assert_eq!(bytes, original, "Document reopen mutated its source copy");
    Ok(())
}

#[test]
fn type_6201_hosts_use_strict_lists_without_public_raw_ids() -> TestResult {
    let bytes = synthetic_table_data_list_package_with_message_type(
        Corruption::None,
        12,
        TABLE_DATA_LIST_ALT_MESSAGE_TYPE,
    )?;
    let original = bytes.clone();
    let package = Package::from_bytes(&bytes)?;
    assert_semantics(package.document())?;
    assert_eq!(bytes, original, "Package parsing mutated its source");

    // The public package/document boundary retains only semantic values. In
    // particular, the alternate host message type and physical archive IDs
    // from this fixture do not appear in either public debug representation.
    assert_eq!(format!("{package:?}"), "Package { .. }");
    let document_debug = format!("{:?}", package.document());
    assert!(!document_debug.contains("6201"));
    assert!(!document_debug.contains("identifier"));
    assert!(!document_debug.contains("ArchiveObject"));
    assert!(!document_debug.contains("RawMessage"));

    let document = Document::from_bytes(&bytes)?;
    assert_semantics(&document)?;
    assert_eq!(bytes, original, "Document parsing mutated its source");

    // This duplicate known scalar is accepted by generated Prost decoding,
    // but strict canonical-wire ingress must reject it on the 6201 route.
    let mut prost_accepted = list_message(
        tst::table_data_list::ListType::String,
        vec![string_entry(1, "strict")],
        Vec::new(),
    )
    .encode_to_vec();
    prost_accepted.extend_from_slice(&[0x08, 0x01]);
    assert!(tst::TableDataList::decode(prost_accepted.as_slice()).is_ok());

    let malformed = synthetic_table_data_list_package_with_message_type(
        Corruption::NonCanonicalRootList,
        12,
        TABLE_DATA_LIST_ALT_MESSAGE_TYPE,
    )?;
    let malformed_original = malformed.clone();
    let package_error = Package::from_bytes(&malformed)
        .expect_err("6201 non-canonical list host must be rejected by strict ingress");
    assert!(!package_error.to_string().contains("6201"));
    assert_eq!(malformed, malformed_original);
    let document_error = Document::from_bytes(&malformed)
        .expect_err("Document must reject the same strict 6201 list host");
    assert!(!document_error.to_string().contains("6201"));
    assert_eq!(malformed, malformed_original);
    // The 6011 objects above are synthetic segment fixtures only; this test
    // makes no claim about native Segment extraction evidence.
    Ok(())
}

#[test]
fn malformed_later_segment_is_rejected_atomically_by_package_and_document() -> TestResult {
    assert_rejected_by_both_readers(Corruption::MalformedLaterSegmentPayload)
}

#[test]
fn missing_segment_payload_is_rejected_atomically() -> TestResult {
    assert_rejected_by_both_readers(Corruption::MissingSegmentPayload)
}

#[test]
fn duplicate_segment_payload_is_rejected_atomically() -> TestResult {
    assert_rejected_by_both_readers(Corruption::DuplicateSegmentPayload)
}

#[test]
fn missing_or_duplicate_segment_reference_is_rejected_atomically() -> TestResult {
    assert_rejected_by_both_readers(Corruption::MissingSegmentReference)?;
    assert_rejected_by_both_readers(Corruption::DuplicateSegmentReference)
}

#[test]
fn wrong_segment_type_range_overflow_and_out_of_range_entries_are_rejected() -> TestResult {
    assert_rejected_by_both_readers(Corruption::WrongSegmentType)?;
    assert_rejected_by_both_readers(Corruption::RangeOverflow)?;
    assert_rejected_by_both_readers(Corruption::EntryOutsideRange)
}

#[test]
fn duplicate_root_and_segment_keys_are_rejected_before_publication() -> TestResult {
    assert_rejected_by_both_readers(Corruption::DuplicateRootKey)?;
    assert_rejected_by_both_readers(Corruption::DuplicateSegmentKey)
}

#[test]
fn selected_malformed_list_is_strict_in_package_and_document_paths() -> TestResult {
    assert_rejected_by_both_readers(Corruption::SelectedMalformedList)
}

#[test]
fn document_ignores_valid_and_dangling_comment_ids_while_package_stays_strict() -> TestResult {
    let valid = synthetic_table_data_list_package(Corruption::None, 12)?;
    let valid_original = valid.clone();
    let package = Package::from_bytes(&valid)?;
    assert_semantics(package.document())?;
    let document = Document::from_bytes(&valid)?;
    assert_semantics(&document)?;
    assert_eq!(valid, valid_original);

    let dangling = synthetic_table_data_list_package(Corruption::None, 999)?;
    let dangling_original = dangling.clone();
    assert!(Package::from_bytes(&dangling).is_err());
    assert_eq!(dangling, dangling_original);
    let document = Document::from_bytes(&dangling)?;
    assert_semantics(&document)?;
    assert_eq!(dangling, dangling_original);
    Ok(())
}

#[test]
fn package_projects_full_comment_metadata_and_replies_without_mutating_source() -> TestResult {
    let bytes = synthetic_table_data_list_package(Corruption::None, 12)?;
    let original = bytes.clone();

    // The public semantic document intentionally drops format-owned comment
    // sidecars. Package construction nevertheless walks the comment storage
    // payload (including date, author, UUID, and both replies); the structured
    // compatibility projection exercises that same strict path a second time.
    let package = Package::from_bytes(&bytes)?;
    assert_semantics(package.document())?;
    assert_eq!(package.extract_structured_tables()?.len(), 1);
    assert_eq!(bytes, original, "Package parsing mutated its source");

    let document = Document::from_bytes(&bytes)?;
    assert_semantics(&document)?;
    assert_eq!(bytes, original, "Document parsing mutated its source");
    Ok(())
}

#[test]
fn missing_comment_text_is_an_empty_strict_value_and_document_still_skips_comments() -> TestResult {
    let bytes = synthetic_table_data_list_package(Corruption::CommentMissingText, 12)?;
    let original = bytes.clone();

    let package = Package::from_bytes(&bytes)?;
    assert_semantics(package.document())?;
    assert_eq!(package.extract_structured_tables()?.len(), 1);
    assert_eq!(bytes, original, "Package parsing mutated its source");

    let document = Document::from_bytes(&bytes)?;
    assert_semantics(&document)?;
    assert_eq!(bytes, original, "Document parsing mutated its source");
    Ok(())
}

fn assert_package_rejects_comment_but_document_skips_it(corruption: Corruption) -> TestResult {
    let bytes = synthetic_table_data_list_package(corruption, 12)?;
    let original = bytes.clone();

    assert!(Package::from_bytes(&bytes).is_err());
    assert_eq!(bytes, original, "Package rejection mutated its source");

    let document = Document::from_bytes(&bytes)?;
    assert_semantics(&document)?;
    assert_eq!(bytes, original, "Document parsing mutated its source");
    Ok(())
}

#[test]
fn malformed_nested_comment_wire_is_atomic_in_package_and_skipped_by_document() -> TestResult {
    assert_package_rejects_comment_but_document_skips_it(Corruption::CommentMalformedNestedWire)
}

#[test]
fn duplicate_comment_storage_payload_is_atomic_in_package_and_skipped_by_document() -> TestResult {
    let bytes =
        synthetic_table_data_list_package(Corruption::CommentDuplicateCanonicalPayload, 12)?;
    let original = bytes.clone();

    let error = Package::from_bytes(&bytes).expect_err("duplicate payload must be rejected");
    assert!(
        error
            .to_string()
            .contains("multiple TSD comment-storage payloads"),
        "valid duplicate must win over candidate decoding, got {error}"
    );
    assert_eq!(bytes, original, "Package rejection mutated its source");

    let document = Document::from_bytes(&bytes)?;
    assert_semantics(&document)?;
    assert_eq!(bytes, original, "Document parsing mutated its source");
    Ok(())
}

#[test]
fn malformed_duplicate_comment_payload_is_rejected_atomically() -> TestResult {
    assert_package_rejects_comment_but_document_skips_it(
        Corruption::CommentDuplicateMalformedPayload,
    )
}

#[test]
fn package_cell_comment_edit_clear_and_inverse_are_selector_first() -> TestResult {
    // A5 points at the isolated root entry when the fixture uses key 50, so
    // the replacement and inverse assertions exercise the supported seam.
    let bytes = synthetic_table_data_list_package(Corruption::None, 50)?;
    let original_bytes = bytes.clone();
    let package = Package::from_bytes(&bytes)?;
    let position = CellPosition::from_a1("A5")?;
    let sheet = "Mixed table-data-list sheet";
    let table = "Mixed table-data-list table";

    let original = package
        .table_cell_comment(sheet, table, position)?
        .ok_or_else(|| std::io::Error::other("synthetic comment is missing"))?;
    assert_eq!(original.text(), "Comment retained by Package");

    let noop = package
        .edit_table_cell_comment(sheet, table, position)?
        .set(original.text())
        .commit()?;
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());

    let changed = package.set_table_cell_comment(
        sheet,
        table,
        position,
        "Comment changed through Package",
    )?;
    assert_eq!(
        changed
            .package()
            .table_cell_comment(sheet, table, position)?
            .as_ref()
            .map(|comment| comment.text()),
        Some("Comment changed through Package")
    );
    let restored = changed
        .package()
        .apply_table_cell_comment(&changed.patch().inverse())?;
    assert_eq!(
        restored
            .package()
            .table_cell_comment(sheet, table, position)?,
        Some(original.clone())
    );

    let clear_error = package
        .clear_table_cell_comment(sheet, table, position)
        .expect_err("changed clear must be refused without graph cleanup");
    assert!(matches!(
        clear_error,
        TableCellCommentError::UnsupportedDependency {
            path: TableCellCommentPath::Cell {
                sheet: 0,
                table: 0,
                row: 4,
                column: 0,
            },
        }
    ));
    assert_eq!(bytes, original_bytes, "clear refusal mutated source bytes");
    assert_eq!(
        package.table_cell_comment(sheet, table, position)?,
        Some(original.clone()),
    );

    // The same fixture also has a segment-backed comment at key 12. Segment
    // entries remain readable, but text replacement is outside the exact
    // root-only publication seam until the owning segment graph can be
    // rewired atomically.
    let segmented_bytes = synthetic_table_data_list_package(Corruption::None, 12)?;
    let segmented_original_bytes = segmented_bytes.clone();
    let segmented_package = Package::from_bytes(&segmented_bytes)?;
    let segmented_original = segmented_package
        .table_cell_comment(sheet, table, position)?
        .ok_or_else(|| std::io::Error::other("segmented synthetic comment is missing"))?;
    let segmented_error = segmented_package
        .set_table_cell_comment(
            sheet,
            table,
            position,
            "Segment-backed comment changed through Package",
        )
        .expect_err("segment-backed text replacement must be refused");
    assert!(matches!(
        segmented_error,
        TableCellCommentError::UnsupportedDependency {
            path: TableCellCommentPath::Cell {
                sheet: 0,
                table: 0,
                row: 4,
                column: 0,
            },
        }
    ));
    assert_eq!(
        segmented_bytes, segmented_original_bytes,
        "segmented replacement refusal mutated source bytes"
    );
    assert_eq!(
        segmented_package.table_cell_comment(sheet, table, position)?,
        Some(segmented_original.clone()),
    );
    let segmented_clear_error = segmented_package
        .clear_table_cell_comment(sheet, table, position)
        .expect_err("segmented changed clear must be refused");
    assert!(matches!(
        segmented_clear_error,
        TableCellCommentError::UnsupportedDependency {
            path: TableCellCommentPath::Cell {
                sheet: 0,
                table: 0,
                row: 4,
                column: 0,
            },
        }
    ));
    assert_eq!(
        segmented_bytes, segmented_original_bytes,
        "segmented clear refusal mutated source bytes"
    );
    assert_eq!(
        segmented_package.table_cell_comment(sheet, table, position)?,
        Some(segmented_original.clone()),
    );
    let reopened = Package::from_bytes(&segmented_bytes)?;
    assert_eq!(
        reopened.table_cell_comment(sheet, table, position)?,
        Some(segmented_original),
    );
    Ok(())
}

#[test]
fn metadata_backed_root_comment_clear_reopens_and_inverts_exactly() -> TestResult {
    let source = include_bytes!("fixtures/comment-edit-root.numbers").as_slice();
    let package = Package::from_bytes(source)?;
    let sheet = package
        .sheets()
        .first()
        .ok_or_else(|| std::io::Error::other("fixture sheet is missing"))?;
    let table = sheet
        .tables()
        .next()
        .ok_or_else(|| std::io::Error::other("fixture table is missing"))?;
    let mut selected = None;
    'rows: for row in 0..table.dimensions().rows() {
        for column in 0..table.dimensions().columns() {
            let position = CellPosition::new(row, column);
            if let Some(comment) =
                package.table_cell_comment(sheet.name(), table.name(), position)?
            {
                selected = Some((position, comment));
                break 'rows;
            }
        }
    }
    let (position, before) =
        selected.ok_or_else(|| std::io::Error::other("fixture comment is missing"))?;

    let cleared = package.clear_table_cell_comment(sheet.name(), table.name(), position)?;
    assert!(cleared.diagnostics().changed());
    assert_eq!(
        cleared
            .package()
            .table_cell_comment(sheet.name(), table.name(), position)?,
        None
    );
    let mut candidate_bytes = Vec::new();
    cleared.package().write_to(&mut candidate_bytes)?;
    let reopened = Package::from_bytes(&candidate_bytes)?;
    assert_eq!(
        reopened.table_cell_comment(sheet.name(), table.name(), position)?,
        None
    );

    let restored = cleared
        .package()
        .apply_table_cell_comment(&cleared.patch().inverse())?;
    assert_eq!(
        restored
            .package()
            .table_cell_comment(sheet.name(), table.name(), position)?,
        Some(before)
    );
    let mut restored_bytes = Vec::new();
    restored.package().write_to(&mut restored_bytes)?;
    assert_eq!(restored_bytes, source);
    Ok(())
}

#[test]
fn root_comment_clear_supersedes_a_failed_staged_replacement() -> TestResult {
    let source = include_bytes!("fixtures/comment-edit-root.numbers").as_slice();
    let semantic = PackageSemanticLimits::default()
        .with_projection_limits(PackageSemanticLimits::MAX_MATERIALIZED_CELLS, 64)?;
    let package = Package::from_bytes_with_options(
        source,
        PackageReadOptions::new(PackageLimits::default(), semantic),
    )?;
    let sheet = package
        .sheets()
        .first()
        .ok_or_else(|| std::io::Error::other("fixture sheet is missing"))?;
    let table = sheet
        .tables()
        .next()
        .ok_or_else(|| std::io::Error::other("fixture table is missing"))?;
    let position = CellPosition::new(1, 1);

    let cleared = package
        .edit_table_cell_comment(sheet.name(), table.name(), position)?
        .set("x".repeat(65))
        .clear()
        .commit()?;

    assert_eq!(
        cleared
            .package()
            .table_cell_comment(sheet.name(), table.name(), position)?,
        None
    );
    Ok(())
}
