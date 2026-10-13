#[test]
fn edit_transaction_appends_a_validated_worksheet() {
    let workbook = litchi_xlsx::Package::new().unwrap();
    let mut edit = workbook.edit().unwrap();
    edit.add("Inserted").unwrap();
    let committed = edit.commit().unwrap();
    let names = committed
        .package()
        .workbook_ref()
        .sheets()
        .map(|sheet| sheet.name().to_owned())
        .collect::<Vec<_>>();
    assert_eq!(names, ["Sheet1", "Inserted"]);
}
