use litchi_ods::{Builder, Cell, CellValue, CellView, MutableSpreadsheet, Row, Sheet, Spreadsheet};

#[test]
fn worksheet_graph_round_trips_repeats_formula_and_style() {
    let mut sheet = Sheet::new("Data").expect("test fixture or operation should succeed");
    sheet
        .set_style_name("sheet-style")
        .expect("test fixture or operation should succeed");

    let mut row = Row::repeated(3).expect("test fixture or operation should succeed");
    let mut number = Cell::repeated(CellValue::Number(7.5), "7.5", 2)
        .expect("test fixture or operation should succeed");
    number
        .set_formula("of:=SUM([.A1:.B1])")
        .expect("test fixture or operation should succeed");
    number
        .set_style_name("number-style")
        .expect("test fixture or operation should succeed");
    row.push_cell(number)
        .expect("test fixture or operation should succeed");
    row.push_cell(
        Cell::repeated(CellValue::Text("same".to_string()), "same", 4)
            .expect("test fixture or operation should succeed"),
    )
    .expect("test fixture or operation should succeed");
    sheet
        .push_row(row)
        .expect("test fixture or operation should succeed");

    let mut builder = Builder::new();
    builder
        .add_sheet(sheet)
        .expect("test fixture or operation should succeed");
    let spreadsheet = Spreadsheet::from_bytes(
        builder
            .build()
            .expect("test fixture or operation should succeed"),
    )
    .expect("test fixture or operation should succeed");

    let data = spreadsheet
        .sheet("Data")
        .expect("test fixture or operation should succeed");
    assert_eq!(data.logical_row_count(), 3);
    assert_eq!(data.logical_column_count(), 6);
    let CellView::Stored(cell) = spreadsheet
        .cell("Data", 2, 1)
        .expect("test fixture or operation should succeed")
    else {
        panic!("expected the repeated physical cell run")
    };
    assert_eq!(cell.repeat(), 2);
    assert_eq!(cell.value, CellValue::Number(7.5));
    assert_eq!(cell.formula.as_deref(), Some("of:=SUM([.A1:.B1])"));
    assert_eq!(cell.style_name.as_deref(), Some("number-style"));
    assert!(spreadsheet.content_xml().contains("office:value=\"7.5\""));
}

#[test]
fn builder_cell_edit_splits_only_the_touched_repeat_runs() {
    let mut sheet = Sheet::new("Data").expect("test fixture or operation should succeed");
    let mut row = Row::repeated(5).expect("test fixture or operation should succeed");
    row.push_cell(
        Cell::repeated(CellValue::Empty, "", 5).expect("test fixture or operation should succeed"),
    )
    .expect("test fixture or operation should succeed");
    sheet
        .push_row(row)
        .expect("test fixture or operation should succeed");

    let mut builder = Builder::new();
    builder
        .add_sheet(sheet)
        .expect("test fixture or operation should succeed");
    builder
        .set_cell("Data", 2, 2, Cell::new(CellValue::Number(42.0), "42"))
        .expect("test fixture or operation should succeed")
        .set_formula("Data", 2, 2, "of:=42")
        .expect("test fixture or operation should succeed")
        .set_cell_style("Data", 2, 2, "cell-style")
        .expect("test fixture or operation should succeed");

    let sheets = builder
        .sheets()
        .expect("test fixture or operation should succeed");
    assert_eq!(sheets[0].rows.len(), 3);
    assert_eq!(sheets[0].rows[0].repeat(), 2);
    assert_eq!(sheets[0].rows[1].repeat(), 1);
    assert_eq!(sheets[0].rows[2].repeat(), 2);
    let cell = sheets[0]
        .cell(2, 2)
        .expect("test fixture or operation should succeed");
    assert_eq!(cell.repeat(), 1);
    assert_eq!(cell.value, CellValue::Number(42.0));
    assert_eq!(cell.formula.as_deref(), Some("of:=42"));
    assert_eq!(cell.style_name.as_deref(), Some("cell-style"));
}

#[test]
fn mutable_worksheet_edits_are_atomic_on_validation_failure() {
    let mut sheet = Sheet::new("Data").expect("test fixture or operation should succeed");
    sheet
        .push_row(Row::new())
        .expect("test fixture or operation should succeed");
    let mut builder = Builder::new();
    builder
        .add_sheet(sheet)
        .expect("test fixture or operation should succeed");
    let bytes = builder
        .build()
        .expect("test fixture or operation should succeed");

    let mut mutable =
        MutableSpreadsheet::from_bytes(bytes).expect("test fixture or operation should succeed");
    let before = mutable.sheets().to_vec();
    assert!(
        mutable
            .add_sheet(Sheet::new("Data").expect("test fixture or operation should succeed"))
            .is_err()
    );
    assert_eq!(mutable.sheets(), before.as_slice());

    assert!(mutable.set_cell_style("Data", 0, 0, "").is_err());
    assert_eq!(mutable.sheets(), before.as_slice());

    mutable
        .set_cell(
            "Data",
            0,
            0,
            Cell::new(CellValue::Text("ok".to_string()), "ok"),
        )
        .expect("test fixture or operation should succeed");
    assert!(matches!(
        mutable.cell("Data", 0, 0),
        Some(CellView::Stored(_))
    ));
}

#[test]
fn mutable_worksheet_remove_is_transactional() {
    let first = Sheet::new("First").expect("test fixture or operation should succeed");
    let second = Sheet::new("Second").expect("test fixture or operation should succeed");
    let mut builder = Builder::new();
    builder
        .add_sheet(first)
        .expect("test fixture or operation should succeed");
    builder
        .add_sheet(second)
        .expect("test fixture or operation should succeed");

    let mut mutable = MutableSpreadsheet::from_bytes(
        builder
            .build()
            .expect("test fixture or operation should succeed"),
    )
    .expect("test fixture or operation should succeed");
    let removed = mutable
        .remove_sheet("First")
        .expect("test fixture or operation should succeed");
    assert_eq!(removed.name, "First");
    assert_eq!(mutable.sheets().len(), 1);
    assert_eq!(mutable.sheets()[0].name, "Second");
    assert!(mutable.remove_sheet("Missing").is_err());
    assert_eq!(mutable.sheets().len(), 1);
}
