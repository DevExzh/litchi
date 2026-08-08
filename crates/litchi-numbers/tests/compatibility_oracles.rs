use std::error::Error as StdError;

use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use litchi_iwa_protos::{tn, tsp, tst};
use litchi_numbers::{
    MAX_OBJECTS, MAX_REFERENCES, MAX_SHEETS, Package, PackageError, PackageLimits,
    PackageReadOptions, PackageSemanticLimits, PackageSemanticPath, SemanticLimitKind, cell::Value,
    compatibility_tables_from_bytes,
};
use prost::Message as _;

const DOCUMENT_MESSAGE_TYPE: u32 = 1;
const SHEET_MESSAGE_TYPE: u32 = 2;
const TABLE_INFO_MESSAGE_TYPE: u32 = 6_000;
const TABLE_MODEL_MESSAGE_TYPE: u32 = 6_001;
const TILE_MESSAGE_TYPE: u32 = 6_002;
const TABLE_DATA_LIST_MESSAGE_TYPE: u32 = 6_005;

type TestResult<T = ()> = Result<T, Box<dyn StdError>>;

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

fn sidecars(identifier: u64) -> TestResult<ArchiveObject> {
    let messages = [
        tst::table_data_list::ListType::String,
        tst::table_data_list::ListType::Formula,
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
    .collect();
    Ok(ArchiveObject::new(identifier, messages)?)
}

fn table_model(name: &str, sidecar_id: u64) -> tst::TableModelArchive {
    tst::TableModelArchive {
        table_name: name.to_owned(),
        number_of_rows: 1,
        number_of_columns: 1,
        base_data_store: tst::DataStore {
            string_table: reference(sidecar_id),
            formula_table: reference(sidecar_id),
            ..Default::default()
        },
        ..Default::default()
    }
}

fn table_model_with_type_nine_cell(
    name: &str,
    sidecar_id: u64,
    tile_id: u64,
) -> TestResult<(tst::TableModelArchive, ArchiveObject)> {
    let mut cell = vec![5, 9, 0, 0, 0, 0, 0, 0];
    cell.extend_from_slice(&1_u32.to_le_bytes());
    cell.extend_from_slice(&[0x39, 0x30, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0x3e, 0xb0]);

    let tile = tst::Tile {
        numrows: 1,
        row_infos: vec![tst::TileRowInfo {
            tile_row_index: 0,
            cell_count: 1,
            storage_version: Some(5),
            cell_storage_buffer: Some(cell),
            cell_offsets: Some(vec![0xff, 0xff, 0, 0]),
            ..Default::default()
        }],
        ..Default::default()
    };
    let mut model = table_model(name, sidecar_id);
    model.number_of_columns = 2;
    model.base_data_store.tiles = tst::TileStorage {
        tiles: vec![tst::tile_storage::Tile {
            tileid: 0,
            tile: reference(tile_id),
        }],
        tile_size: Some(256),
        ..Default::default()
    };
    Ok((
        model,
        object(tile_id, TILE_MESSAGE_TYPE, tile.encode_to_vec())?,
    ))
}

fn root_object(sheet_ids: impl IntoIterator<Item = u64>) -> TestResult<ArchiveObject> {
    let root = tn::DocumentArchive {
        sheets: sheet_ids.into_iter().map(reference).collect(),
        ..Default::default()
    };
    object(1, DOCUMENT_MESSAGE_TYPE, root.encode_to_vec())
}

fn sheet_object(
    identifier: u64,
    name: &str,
    drawable_ids: impl IntoIterator<Item = u64>,
) -> TestResult<ArchiveObject> {
    let sheet = tn::SheetArchive {
        name: name.to_owned(),
        drawable_infos: drawable_ids.into_iter().map(reference).collect(),
        ..Default::default()
    };
    object(identifier, SHEET_MESSAGE_TYPE, sheet.encode_to_vec())
}

fn table_info_object(identifier: u64, model_id: u64) -> TestResult<ArchiveObject> {
    let info = tst::TableInfoArchive {
        table_model: reference(model_id),
        ..Default::default()
    };
    object(identifier, TABLE_INFO_MESSAGE_TYPE, info.encode_to_vec())
}

fn package_bytes(objects: Vec<ArchiveObject>) -> TestResult<Vec<u8>> {
    let iwa = SnappyStream::compress(&Archive { objects }.to_bytes()?)?;
    Ok(litchi_iwa_archive::package::to_bytes(
        [("Index/Document.iwa", iwa.as_slice())],
        PackageLimits::default(),
    )?)
}

fn assert_named_empty_table(table: &litchi_numbers::Table, expected_name: &str) -> TestResult {
    assert_eq!(table.name(), expected_name);
    assert_eq!((table.row_count(), table.column_count()), (1, 1));
    assert_eq!(table.cell_count(), 0);
    assert_eq!(table.get_a1("A1")?, None);
    Ok(())
}

fn compatibility_oracle_objects(reverse_physical_order: bool) -> TestResult<Vec<ArchiveObject>> {
    let (first_model, tile) = table_model_with_type_nine_cell("First canonical", 90, 30)?;
    let first_payload = first_model.encode_to_vec();
    let first_with_legacy_duplicate = ArchiveObject::new(
        10,
        vec![
            RawMessage {
                type_: TABLE_MODEL_MESSAGE_TYPE,
                data: first_payload.clone(),
            },
            RawMessage {
                type_: TABLE_INFO_MESSAGE_TYPE,
                data: first_payload,
            },
        ],
    )?;
    let mut objects = vec![
        root_object([2])?,
        sheet_object(2, "Rooted sheet", [4, 3])?,
        table_info_object(3, 10)?,
        table_info_object(4, 11)?,
        object(
            5,
            TABLE_INFO_MESSAGE_TYPE,
            table_model("Detached legacy", 90).encode_to_vec(),
        )?,
        first_with_legacy_duplicate,
        object(
            11,
            TABLE_MODEL_MESSAGE_TYPE,
            table_model("Second canonical", 90).encode_to_vec(),
        )?,
        sidecars(90)?,
        tile,
    ];
    if reverse_physical_order {
        objects.reverse();
    }
    Ok(objects)
}

fn generated_oracle_fixture(reverse_physical_order: bool) -> TestResult<Vec<u8>> {
    package_bytes(compatibility_oracle_objects(reverse_physical_order)?)
}

fn decode_hex_fixture(source: &str) -> TestResult<Vec<u8>> {
    let mut decoded = Vec::new();
    let mut high_nibble = None;
    for byte in source.bytes().filter(|byte| !byte.is_ascii_whitespace()) {
        let nibble = match byte {
            b'0'..=b'9' => byte - b'0',
            b'a'..=b'f' => byte - b'a' + 10,
            _ => {
                return Err(std::io::Error::new(
                    std::io::ErrorKind::InvalidData,
                    "Numbers compatibility fixture contains non-lowercase-hex input",
                )
                .into());
            },
        };
        if let Some(high) = high_nibble.take() {
            decoded.push((high << 4) | nibble);
        } else {
            high_nibble = Some(nibble);
        }
    }
    if high_nibble.is_some() {
        return Err(std::io::Error::new(
            std::io::ErrorKind::InvalidData,
            "Numbers compatibility fixture contains an incomplete hex byte",
        )
        .into());
    }
    Ok(decoded)
}

fn checked_oracle_fixture() -> TestResult<Vec<u8>> {
    let checked_in = decode_hex_fixture(include_str!(
        "../../../test-data/synthetic-iwork/numbers/compatibility-oracles.hex"
    ))?;
    assert_eq!(checked_in, generated_oracle_fixture(true)?);
    Ok(checked_in)
}

#[test]
fn compatibility_projection_retains_a_detached_table_model() -> TestResult {
    let bytes = checked_oracle_fixture()?;
    let package = Package::from_bytes(&bytes)?;
    assert_eq!(package.sheets().len(), 1);
    assert_eq!(package.sheets()[0].name(), "Rooted sheet");
    assert_eq!(package.sheets()[0].tables().count(), 2);

    let tables = package.extract_structured_tables()?;
    assert_eq!(tables.len(), 3);
    assert_named_empty_table(&tables[2], "Detached legacy")?;
    assert!(
        package.sheets()[0]
            .tables()
            .all(|table| table.name() != "Detached legacy")
    );
    Ok(())
}

#[test]
fn compatibility_projection_decodes_type_nine_numeric_cells() -> TestResult {
    let bytes = checked_oracle_fixture()?;
    let package = Package::from_bytes(&bytes)?;
    let tables = package.extract_structured_tables()?;
    assert_eq!(tables[0].name(), "First canonical");
    assert_eq!((tables[0].row_count(), tables[0].column_count()), (1, 2));
    assert_eq!(tables[0].cell_count(), 1);
    assert_eq!(tables[0].get_a1("A1")?, None);
    let value = tables[0].get_a1("B1")?;
    assert_eq!(value, Some(&Value::number(-1_234.5)?));
    let Some(Value::Number(number)) = value else {
        panic!("type-nine B1 was not retained as a finite number: {value:?}");
    };
    assert!(number.get().is_finite());
    assert_eq!(number.get().to_bits(), (-1_234.5_f64).to_bits());
    Ok(())
}

#[test]
fn compatibility_projection_uses_global_identity_order_not_root_or_object_order() -> TestResult {
    let bytes = checked_oracle_fixture()?;
    let package = Package::from_bytes(&bytes)?;

    let rooted = package.sheets()[0]
        .tables()
        .map(litchi_numbers::Table::name)
        .collect::<Vec<_>>();
    assert_eq!(rooted, ["Second canonical", "First canonical"]);

    let global = package.extract_structured_tables()?;
    assert_eq!(
        global
            .iter()
            .map(litchi_numbers::Table::name)
            .collect::<Vec<_>>(),
        ["First canonical", "Second canonical", "Detached legacy"]
    );
    assert_named_empty_table(&global[1], "Second canonical")?;
    assert_named_empty_table(&global[2], "Detached legacy")?;
    assert_eq!(compatibility_tables_from_bytes(&bytes)?, global);

    let reordered =
        Package::from_bytes(&generated_oracle_fixture(false)?)?.extract_structured_tables()?;
    assert_eq!(reordered, global);
    Ok(())
}

#[test]
fn canonical_models_precede_legacy_models_and_duplicate_payloads_are_emitted_once() -> TestResult {
    let bytes = checked_oracle_fixture()?;
    let package = Package::from_bytes(&bytes)?;
    let tables = package.extract_structured_tables()?;
    assert_eq!(tables.len(), 3);
    assert_eq!(tables[0].name(), "First canonical");
    assert_named_empty_table(&tables[1], "Second canonical")?;
    assert_named_empty_table(&tables[2], "Detached legacy")?;
    Ok(())
}

#[test]
fn global_table_limit_is_inclusive_and_reports_the_first_exceeded_table() -> TestResult {
    let bytes = checked_oracle_fixture()?;
    let options = |max_tables| -> TestResult<PackageReadOptions> {
        Ok(PackageReadOptions::new(
            PackageLimits::default(),
            PackageSemanticLimits::new(MAX_OBJECTS, MAX_SHEETS, max_tables, MAX_REFERENCES)?,
        ))
    };

    let exact =
        Package::from_bytes_with_options(&bytes, options(3)?)?.extract_structured_tables()?;
    assert_eq!(exact.len(), 3);
    assert_eq!(exact[0].name(), "First canonical");
    assert_named_empty_table(&exact[1], "Second canonical")?;
    assert_named_empty_table(&exact[2], "Detached legacy")?;

    assert!(matches!(
        Package::from_bytes_with_options(&bytes, options(2)?)?.extract_structured_tables(),
        Err(PackageError::SemanticLimit {
            kind: SemanticLimitKind::Tables,
            observed: 3,
            maximum: 2,
            path: PackageSemanticPath::StructuredTables,
        })
    ));
    Ok(())
}
