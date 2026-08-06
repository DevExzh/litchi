use litchi_ods::{Builder, Cell, CellValue, CellView, MutableSpreadsheet, Row, Sheet, Spreadsheet};

#[test]
fn worksheet_graph_round_trips_repeats_formula_and_style() {
    let mut sheet = Sheet::new("Data").unwrap();
    sheet.set_style_name("sheet-style").unwrap();

    let mut row = Row::repeated(3).unwrap();
    let mut number = Cell::repeated(CellValue::Number(7.5), "7.5", 2).unwrap();
    number.set_formula("of:=SUM([.A1:.B1])").unwrap();
    number.set_style_name("number-style").unwrap();
    row.push_cell(number).unwrap();
    row.push_cell(Cell::repeated(CellValue::Text("same".to_string()), "same", 4).unwrap())
        .unwrap();
    sheet.push_row(row).unwrap();

    let mut builder = Builder::new();
    builder.add_sheet(sheet).unwrap();
    let spreadsheet = Spreadsheet::from_bytes(builder.build().unwrap()).unwrap();

    let data = spreadsheet.sheet("Data").unwrap();
    assert_eq!(data.logical_row_count(), 3);
    assert_eq!(data.logical_column_count(), 6);
    let CellView::Stored(cell) = spreadsheet.cell("Data", 2, 1).unwrap() else {
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
    let mut sheet = Sheet::new("Data").unwrap();
    let mut row = Row::repeated(5).unwrap();
    row.push_cell(Cell::repeated(CellValue::Empty, "", 5).unwrap())
        .unwrap();
    sheet.push_row(row).unwrap();

    let mut builder = Builder::new();
    builder.add_sheet(sheet).unwrap();
    builder
        .set_cell("Data", 2, 2, Cell::new(CellValue::Number(42.0), "42"))
        .unwrap()
        .set_formula("Data", 2, 2, "of:=42")
        .unwrap()
        .set_cell_style("Data", 2, 2, "cell-style")
        .unwrap();

    let sheets = builder.sheets().unwrap();
    assert_eq!(sheets[0].rows.len(), 3);
    assert_eq!(sheets[0].rows[0].repeat(), 2);
    assert_eq!(sheets[0].rows[1].repeat(), 1);
    assert_eq!(sheets[0].rows[2].repeat(), 2);
    let cell = sheets[0].cell(2, 2).unwrap();
    assert_eq!(cell.repeat(), 1);
    assert_eq!(cell.value, CellValue::Number(42.0));
    assert_eq!(cell.formula.as_deref(), Some("of:=42"));
    assert_eq!(cell.style_name.as_deref(), Some("cell-style"));
}

#[test]
fn mutable_worksheet_edits_are_atomic_on_validation_failure() {
    let mut sheet = Sheet::new("Data").unwrap();
    sheet.push_row(Row::new()).unwrap();
    let mut builder = Builder::new();
    builder.add_sheet(sheet).unwrap();
    let bytes = builder.build().unwrap();

    let mut mutable = MutableSpreadsheet::from_bytes(bytes).unwrap();
    let before = mutable.sheets().to_vec();
    assert!(mutable.add_sheet(Sheet::new("Data").unwrap()).is_err());
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
        .unwrap();
    assert!(matches!(
        mutable.cell("Data", 0, 0),
        Some(CellView::Stored(_))
    ));
}

#[test]
fn mutable_worksheet_remove_is_transactional() {
    let first = Sheet::new("First").unwrap();
    let second = Sheet::new("Second").unwrap();
    let mut builder = Builder::new();
    builder.add_sheet(first).unwrap();
    builder.add_sheet(second).unwrap();

    let mut mutable = MutableSpreadsheet::from_bytes(builder.build().unwrap()).unwrap();
    let removed = mutable.remove_sheet("First").unwrap();
    assert_eq!(removed.name, "First");
    assert_eq!(mutable.sheets().len(), 1);
    assert_eq!(mutable.sheets()[0].name, "Second");
    assert!(mutable.remove_sheet("Missing").is_err());
    assert_eq!(mutable.sheets().len(), 1);
}
