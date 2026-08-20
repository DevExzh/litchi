use std::error::Error as StdError;

use litchi_iwa_archive::Limits as ArchiveLimits;
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsp, tst};
use litchi_numbers::cell::Value;
use litchi_numbers::{Document, Package};
use litchi_numbers_wire::BncCell;
use prost::Message as _;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHEET_MESSAGE_TYPE: u32 = 2;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;

fn reference(identifier: u64) -> tsp::Reference {
    tsp::Reference {
        identifier,
        ..Default::default()
    }
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

fn row(
    tile_row_index: u32,
    storage: Vec<u8>,
    offsets: Vec<u8>,
    wide_offsets: bool,
) -> tst::TileRowInfo {
    tst::TileRowInfo {
        tile_row_index,
        cell_count: 1,
        storage_version: Some(5),
        cell_storage_buffer: Some(storage),
        cell_offsets: Some(offsets),
        has_wide_offsets: Some(wide_offsets),
        ..Default::default()
    }
}

fn modern_number(value: f64) -> TestResult<Vec<u8>> {
    let mut cell = BncCell::minimal();
    cell.set_number(value)?;
    Ok(cell.encode())
}

fn modern_boolean(value: bool) -> Vec<u8> {
    let mut cell = BncCell::minimal();
    cell.set_boolean(value);
    cell.encode()
}

fn tile_payload(malformed_later_row: bool) -> TestResult<Vec<u8>> {
    let first_storage = modern_number(41.5)?;
    let second_cell = if malformed_later_row {
        vec![5]
    } else {
        modern_boolean(true)
    };

    // The wide row stores offsets in four-byte units.  Keeping the second
    // cell at byte 256 makes the wide-offset branch observable while leaving
    // the first row's regular BNC offset at byte 0.
    let mut second_storage = vec![0; 256];
    second_storage.extend_from_slice(&second_cell);

    let tile = tst::Tile {
        max_column: 1,
        max_row: 1,
        num_cells: 2,
        numrows: 2,
        row_infos: vec![
            row(0, first_storage, vec![0, 0, 0xff, 0xff], false),
            row(1, second_storage, vec![0xff, 0xff, 64, 0], true),
        ],
        storage_version: Some(5),
        last_saved_in_bnc: Some(true),
        should_use_wide_rows: Some(true),
    };
    Ok(tile.encode_to_vec())
}

fn synthetic_tile_package(malformed_later_row: bool) -> TestResult<Vec<u8>> {
    let sidecars = ArchiveObject::new(
        5,
        [
            tst::table_data_list::ListType::String,
            tst::table_data_list::ListType::Formula,
            tst::table_data_list::ListType::RichTextPayload,
        ]
        .into_iter()
        .map(|list_type| RawMessage {
            type_: TABLE_DATA_LIST_MESSAGE_TYPE,
            data: tst::TableDataList {
                list_type: list_type as i32,
                next_list_id: 1,
                ..Default::default()
            }
            .encode_to_vec(),
        })
        .collect(),
    )?;

    let model = tst::TableModelArchive {
        table_id: "tile-reader-table-id".to_owned(),
        table_name: "Tile reader table".to_owned(),
        number_of_rows: 2,
        number_of_columns: 2,
        base_data_store: tst::DataStore {
            row_headers: tst::HeaderStorage {
                bucket_hash_function: 1,
                ..Default::default()
            },
            column_headers: reference(5),
            tiles: tst::TileStorage {
                tiles: vec![tst::tile_storage::Tile {
                    tileid: 0,
                    tile: reference(6),
                }],
                tile_size: Some(256),
                should_use_wide_rows: Some(true),
            },
            string_table: reference(5),
            style_table: reference(5),
            formula_table: reference(5),
            rich_text_table: Some(reference(5)),
            format_table_pre_bnc: reference(5),
            next_row_strip_id: 1,
            next_column_strip_id: 1,
            row_tile_tree: tst::TableRbTree::default(),
            column_tile_tree: tst::TableRbTree::default(),
            ..Default::default()
        },
        ..Default::default()
    };

    let objects = vec![
        object(
            1,
            DOCUMENT_MESSAGE_TYPE,
            tn::DocumentArchive {
                sheets: vec![reference(2)],
                ..Default::default()
            }
            .encode_to_vec(),
        )?,
        object(
            2,
            SHEET_MESSAGE_TYPE,
            tn::SheetArchive {
                name: "Tile reader sheet".to_owned(),
                drawable_infos: vec![reference(3)],
                ..Default::default()
            }
            .encode_to_vec(),
        )?,
        object(
            3,
            TABLE_INFO_MESSAGE_TYPE,
            tst::TableInfoArchive {
                table_model: reference(4),
                ..Default::default()
            }
            .encode_to_vec(),
        )?,
        object(4, TABLE_MODEL_MESSAGE_TYPE, model.encode_to_vec())?,
        sidecars,
        object(6, TILE_MESSAGE_TYPE, tile_payload(malformed_later_row)?)?,
    ];
    let iwa = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [("Index/Document.iwa", iwa.as_slice())],
        ArchiveLimits::default(),
    )?)
}

fn assert_semantics(document: &Document) -> TestResult {
    document.validate()?;
    let sheet = document
        .sheet(0)?
        .ok_or_else(|| std::io::Error::other("synthetic tile sheet is missing"))?;
    assert_eq!(sheet.name(), "Tile reader sheet");
    let table = sheet
        .tables()
        .next()
        .ok_or_else(|| std::io::Error::other("synthetic tile table is missing"))?;
    assert_eq!(table.name(), "Tile reader table");
    assert_eq!((table.row_count(), table.column_count()), (2, 2));
    assert_eq!(table.cell_count(), 2);
    assert!(matches!(
        table.get_a1("A1")?,
        Some(Value::Number(value)) if value.get().to_bits() == 41.5_f64.to_bits()
    ));
    assert!(matches!(table.get_a1("B2")?, Some(Value::Boolean(true))));
    assert_eq!(
        document.plain_text()?,
        "Tile reader sheet\nTile reader table\n41.5\ntrue"
    );
    Ok(())
}

#[test]
fn synthetic_tile_reader_reopens_source_ordered_modern_and_wide_rows() -> TestResult {
    let bytes = synthetic_tile_package(false)?;
    let package = Package::from_bytes(&bytes)?;
    assert_semantics(package.document())?;
    assert_semantics(&Document::from_bytes(&bytes)?)?;
    Ok(())
}

#[test]
fn synthetic_tile_reader_refuses_a_malformed_later_row_atomically() -> TestResult {
    let bytes = synthetic_tile_package(true)?;
    assert!(Package::from_bytes(&bytes).is_err());
    assert!(Document::from_bytes(&bytes).is_err());
    Ok(())
}
