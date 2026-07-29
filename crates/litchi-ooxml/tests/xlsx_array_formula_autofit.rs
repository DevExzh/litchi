use litchi_ooxml::xlsx::Workbook;
use litchi_opc::{OpcPackage, PackURI};

#[test]
fn array_formulas_and_auto_sized_columns_survive_save_and_reopen() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("array-formulas.xlsx");

    let mut workbook = Workbook::create().unwrap();
    {
        let worksheet = workbook.worksheet_mut(0).unwrap();
        worksheet.set_array_formula(1, 1, 3, 1, "ROW(A1:A3)*2");
        worksheet.set_cell_value(5, 1, "plain");
        // Single-cell array formula without a recorded range.
        worksheet.set_cell_formula(6, 1, "SUM(A1:A3)");
        worksheet.set_cell_value(1, 2, "a much longer string");
        assert_eq!(worksheet.auto_size_column(2), Some(22.0));
        // Column C stays untouched: no measurable content.
        assert_eq!(worksheet.auto_size_column(3), None);
    }
    workbook.save(&path).unwrap();

    let package = OpcPackage::open(&path).unwrap();
    let worksheet_part = package
        .get_part(&PackURI::new("/xl/worksheets/sheet1.xml").unwrap())
        .unwrap();
    let xml = std::str::from_utf8(worksheet_part.blob()).unwrap();
    assert!(xml.contains(r#"<f t="array" ref="A1:A3">ROW(A1:A3)*2</f>"#));
    assert!(xml.contains(r#"width="22""#), "auto-sized column width: {xml}");

    let reopened = Workbook::open(&path).unwrap();
    let formulas = reopened.worksheet_array_formulas(0).unwrap();
    assert_eq!(formulas.len(), 1);
    let formula = &formulas[0];
    assert_eq!(formula.formula, "ROW(A1:A3)*2");
    assert_eq!(formula.range.as_deref(), Some("A1:A3"));
    assert_eq!((formula.row, formula.column), (1, 1));
}

#[test]
fn worksheet_array_formulas_rejects_out_of_bounds_sheet() {
    let workbook = Workbook::create().unwrap();
    assert!(workbook.worksheet_array_formulas(7).is_err());
}
