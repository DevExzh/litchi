use litchi_xlsx::Workbook;

#[test]
fn edit_transaction_appends_a_validated_worksheet() {
    let workbook = Workbook::create().unwrap();
    let mut edit = workbook.edit().unwrap();
    edit.add("Inserted").unwrap();
    let committed = edit.commit().unwrap();
    let names = committed
        .workbook()
        .sheets()
        .map(|sheet| sheet.name().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(names, ["Sheet1", "Inserted"]);
}
