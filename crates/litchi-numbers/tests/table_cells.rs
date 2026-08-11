//! Public integration coverage for selector-first Numbers cell reads.

use std::{fmt::Debug, path::PathBuf};

use litchi_numbers::{
    Package, SheetSelector, TableSelector,
    cell::Value,
    table::{
        CellPosition, CellRange, Dimensions,
        cells::{Error, LimitKind, Path, State, Storage},
    },
};

type TestResult<T = ()> = Result<T, Box<dyn std::error::Error>>;

const FIXTURE_MARKER: &str = "Litchi native Numbers fixture";

fn fixture_path() -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../../test-data/iwork/numbers/basic.numbers")
}

fn assert_send_sync_debug<T: Send + Sync + Debug>() {}

#[test]
fn native_single_reads_use_checked_index_and_name_selectors() -> TestResult {
    let package = Package::open(fixture_path())?;
    let b2 = CellPosition::from_a1("B2")?;
    let b3 = CellPosition::from_a1("B3")?;
    let g22 = CellPosition::from_a1("G22")?;

    let text_by_index = package.table_cell(0usize, 0usize, b2)?;
    let text_by_name = package.table_cell(
        SheetSelector::name("Sheet 1"),
        TableSelector::name("Table 1"),
        b2,
    )?;
    assert_eq!(text_by_name, text_by_index);
    assert_eq!(text_by_index.position(), b2);
    assert!(matches!(
        text_by_index.storage(),
        Storage::Stored(Value::Text(text)) if text == FIXTURE_MARKER
    ));
    assert_eq!(
        text_by_index.storage().value().map(Value::cell_type),
        Some(litchi_numbers::cell::Type::Text)
    );

    let number = package.table_cell("Sheet 1", "Table 1", b3)?;
    assert_eq!(number.position(), b3);
    assert!(matches!(
        number.storage(),
        Storage::Stored(Value::Number(value)) if value.get() == 42.0
    ));

    let missing = package.table_cell("Sheet 1", "Table 1", g22)?;
    assert_eq!(missing.position(), g22);
    assert!(missing.storage().is_missing());
    assert_eq!(missing.storage().value(), None);
    assert!(matches!(missing.storage(), Storage::Missing));
    Ok(())
}

#[test]
fn native_dense_range_is_row_major_and_presence_preserving() -> TestResult {
    let package = Package::open(fixture_path())?;
    let states = package.table_cells(
        SheetSelector::index(0),
        TableSelector::index(0),
        CellRange::from_a1("B2:B3")?,
    )?;

    assert_eq!(states.len(), 2);
    assert_eq!(states[0].position(), CellPosition::from_a1("B2")?);
    assert_eq!(states[1].position(), CellPosition::from_a1("B3")?);
    assert!(matches!(
        states[0].storage(),
        Storage::Stored(Value::Text(text)) if text == FIXTURE_MARKER
    ));
    assert!(matches!(
        states[1].storage(),
        Storage::Stored(Value::Number(value)) if value.get() == 42.0
    ));

    let missing = package.table_cells(0usize, 0usize, CellRange::from_a1("F21:G22")?)?;
    assert_eq!(missing.len(), 4);
    assert_eq!(
        missing.iter().map(State::position).collect::<Vec<_>>(),
        ["F21", "G21", "F22", "G22"]
            .map(CellPosition::from_a1)
            .into_iter()
            .collect::<Result<Vec<_>, _>>()?
    );
    assert!(missing.iter().all(|state| state.storage().is_missing()));
    Ok(())
}

#[test]
fn missing_selectors_and_bounds_return_typed_errors() -> TestResult {
    let package = Package::open(fixture_path())?;
    let b2 = CellPosition::from_a1("B2")?;
    assert!(matches!(
        package.table_cell("missing sheet", 0usize, b2),
        Err(Error::SheetNotFound)
    ));
    assert!(matches!(
        package.table_cell(0usize, "missing table", b2),
        Err(Error::TableNotFound)
    ));

    let dimensions = Dimensions::new(22, 7);
    let outside = CellPosition::new(22, 0);
    assert!(matches!(
        package.table_cell(0usize, 0usize, outside),
        Err(Error::OutOfBounds {
            position,
            dimensions: actual,
        }) if position == outside && actual == dimensions
    ));

    let end = CellPosition::new(22, 8);
    let outside_range = CellRange::new(CellPosition::new(21, 6), end)?;
    assert!(matches!(
        package.table_cells(0usize, 0usize, outside_range),
        Err(Error::OutOfBounds {
            position,
            dimensions: actual,
        }) if position == end && actual == dimensions
    ));
    Ok(())
}

#[test]
fn public_read_values_are_send_sync_and_debug_redacted() -> TestResult {
    assert_send_sync_debug::<Package>();
    assert_send_sync_debug::<Storage>();
    assert_send_sync_debug::<State>();
    assert_send_sync_debug::<Error>();
    assert_send_sync_debug::<LimitKind>();
    assert_send_sync_debug::<Path>();

    let package = Package::open(fixture_path())?;
    let state = package.table_cell("Sheet 1", "Table 1", CellPosition::from_a1("B2")?)?;
    for rendered in [format!("{state:?}"), format!("{:?}", state.storage())] {
        assert!(rendered.contains("Text"));
        assert!(!rendered.contains(FIXTURE_MARKER));
    }

    let error = package
        .table_cell(
            "private missing sheet name",
            0usize,
            CellPosition::new(0, 0),
        )
        .expect_err("missing selector must fail");
    for rendered in [format!("{error:?}"), error.to_string()] {
        assert!(!rendered.contains("private missing sheet name"));
        assert!(!rendered.contains(FIXTURE_MARKER));
    }
    Ok(())
}
