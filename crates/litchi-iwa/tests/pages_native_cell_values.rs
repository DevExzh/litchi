//! Native Pages table cell readback coverage.
//!
//! The fixture was entered, saved, closed, and reopened in Pages 14.4.  The
//! assertions below intentionally follow the host reader's typed projection:
//! the date-looking and duration-looking entries are text in the archive, and
//! the division-by-zero entry retains its formula source even though Pages
//! displays an error for it.

use std::error::Error;

use litchi_iwa::pages::{PagesCellValue, PagesEditor};
use litchi_pages::Package as FocusedPagesPackage;

type TestResult<T = ()> = Result<T, Box<dyn Error>>;

const NATIVE_CELLS: &[u8] = include_bytes!(concat!(
    env!("CARGO_MANIFEST_DIR"),
    "/../../test-data/iwork/pages/body-table-cells-native.pages"
));

#[derive(Debug, Clone, Copy)]
struct ExpectedTable {
    name: &'static str,
    rows: usize,
    columns: usize,
}

const EXPECTED_TABLES: [ExpectedTable; 2] = [
    ExpectedTable {
        name: "Table 1",
        rows: 5,
        columns: 4,
    },
    ExpectedTable {
        name: "Table 2",
        rows: 3,
        columns: 2,
    },
];

fn number(value: f64) -> PagesCellValue {
    PagesCellValue::number(value).expect("fixture number is finite")
}

fn assert_cell(
    table: &litchi_iwa::pages::PagesTable,
    row: usize,
    column: usize,
    expected: Option<PagesCellValue>,
) {
    assert_eq!(
        table.get_cell(row, column),
        expected.as_ref(),
        "unexpected native value at ({row}, {column}) in {}",
        table.info.name
    );
}

fn assert_table_one(editor: &PagesEditor, model_object_id: u64) -> TestResult {
    let table = editor.table(model_object_id)?;
    assert_cell(&table, 0, 0, Some(PagesCellValue::Text("Label".to_owned())));
    assert_cell(&table, 0, 1, Some(PagesCellValue::Text("Value".to_owned())));
    assert_cell(&table, 0, 2, Some(PagesCellValue::Text("Flag".to_owned())));
    assert_cell(&table, 0, 3, Some(PagesCellValue::Text("Note".to_owned())));

    assert_cell(&table, 1, 0, Some(PagesCellValue::Text("Alpha".to_owned())));
    assert_cell(&table, 1, 1, Some(number(12.5)));
    assert_cell(&table, 1, 2, Some(PagesCellValue::Boolean(true)));
    assert_cell(&table, 1, 3, Some(PagesCellValue::Text("北京".to_owned())));

    assert_cell(&table, 2, 0, Some(PagesCellValue::Text("Beta".to_owned())));
    assert_cell(&table, 2, 1, Some(number(-7.0)));
    assert_cell(&table, 2, 2, Some(PagesCellValue::Boolean(false)));
    assert_cell(&table, 2, 3, Some(PagesCellValue::Text("Café".to_owned())));

    assert_cell(&table, 3, 0, Some(PagesCellValue::Text("Total".to_owned())));
    assert_cell(
        &table,
        3,
        1,
        Some(PagesCellValue::Formula("=SUM(B2:B3)".to_owned())),
    );
    assert_cell(&table, 3, 2, None);
    assert_cell(&table, 3, 3, None);

    assert_cell(&table, 4, 0, Some(PagesCellValue::Text("Tail".to_owned())));
    assert_cell(&table, 4, 1, Some(number(0.0)));
    assert_cell(&table, 4, 2, None);
    assert_cell(&table, 4, 3, None);

    assert_eq!(
        editor.table_formula(model_object_id, 3, 1)?.as_deref(),
        Some("=SUM(B2:B3)")
    );
    Ok(())
}

fn assert_table_two(editor: &PagesEditor, model_object_id: u64) -> TestResult {
    let table = editor.table(model_object_id)?;
    assert_cell(
        &table,
        0,
        0,
        Some(PagesCellValue::Text("2026-09-09".to_owned())),
    );
    assert_cell(
        &table,
        0,
        1,
        Some(PagesCellValue::Text("1h 30m".to_owned())),
    );
    assert_cell(
        &table,
        1,
        0,
        Some(PagesCellValue::Formula("=(1/0)".to_owned())),
    );
    assert_cell(&table, 1, 1, None);
    assert_cell(&table, 2, 0, None);
    assert_cell(&table, 2, 1, None);

    assert_eq!(
        editor.table_formula(model_object_id, 1, 0)?.as_deref(),
        Some("=(1/0)")
    );
    Ok(())
}

fn assert_host_tables(editor: &PagesEditor) -> TestResult {
    let tables = editor.tables()?;
    assert_eq!(tables.len(), EXPECTED_TABLES.len());
    for (index, (info, expected)) in tables.iter().zip(EXPECTED_TABLES).enumerate() {
        assert_eq!(info.name, expected.name, "table {index} name");
        assert_eq!(info.rows, expected.rows, "table {index} row count");
        assert_eq!(info.columns, expected.columns, "table {index} column count");
        assert_eq!(info, &editor.table(info.model_object_id)?.info);
        match index {
            0 => assert_table_one(editor, info.model_object_id)?,
            1 => assert_table_two(editor, info.model_object_id)?,
            _ => unreachable!("expected table list has two entries"),
        }
    }
    Ok(())
}

fn focused_bytes(package: &FocusedPagesPackage) -> TestResult<Vec<u8>> {
    let mut bytes = Vec::new();
    package.write_to(&mut bytes)?;
    Ok(bytes)
}

#[test]
fn native_pages_cell_values_and_catalog_roundtrip_exactly() -> TestResult {
    let editor = PagesEditor::from_bytes(NATIVE_CELLS)?;
    assert_eq!(
        editor.to_bytes()?,
        NATIVE_CELLS,
        "source reader changed bytes"
    );
    assert_host_tables(&editor)?;

    let focused = FocusedPagesPackage::from_bytes(NATIVE_CELLS)?;
    let catalog = focused.body_tables()?;
    assert_eq!(catalog.len(), EXPECTED_TABLES.len());
    for (index, (snapshot, expected)) in catalog.iter().zip(EXPECTED_TABLES).enumerate() {
        assert_eq!(snapshot.index(), index);
        assert_eq!(snapshot.name(), expected.name);
        assert_eq!(snapshot.rows() as usize, expected.rows);
        assert_eq!(snapshot.columns() as usize, expected.columns);
        assert_eq!(catalog.select(snapshot.selector())?, Some(snapshot));
        assert_eq!(catalog.select(snapshot.name_selector())?, Some(snapshot));
    }
    assert_eq!(
        focused_bytes(&focused)?,
        NATIVE_CELLS,
        "catalog changed bytes"
    );

    let source_bytes = editor.to_bytes()?;
    let reopened_editor = PagesEditor::from_bytes(&source_bytes)?;
    assert_eq!(reopened_editor.to_bytes()?, NATIVE_CELLS);
    assert_host_tables(&reopened_editor)?;

    let reopened_focused = FocusedPagesPackage::from_bytes(&source_bytes)?;
    assert_eq!(reopened_focused.body_tables()?, catalog);
    assert_eq!(focused_bytes(&reopened_focused)?, NATIVE_CELLS);
    Ok(())
}
