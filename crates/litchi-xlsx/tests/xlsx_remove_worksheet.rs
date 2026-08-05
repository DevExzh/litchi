use litchi_xlsx::Workbook;

#[test]
fn edit_transaction_removes_a_worksheet_from_the_new_snapshot() {
    let workbook = Workbook::create().unwrap();
    let mut add = workbook.edit().unwrap();
    add.add("Second").unwrap();
    let added = add.commit().unwrap().into_workbook();

    let mut remove = added.edit().unwrap();
    assert!(remove.remove("Sheet1").unwrap().is_some());
    let committed = remove.commit().unwrap();
    assert_eq!(committed.workbook().len(), 1);
    assert_eq!(
        committed.workbook().sheet(0).unwrap().unwrap().name(),
        "Second"
    );
}
