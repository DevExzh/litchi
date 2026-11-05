#[test]
fn edit_transaction_removes_a_worksheet_from_the_new_snapshot() {
    let workbook = litchi_xlsx::Package::new().unwrap();
    let mut add = workbook.edit().unwrap();
    add.add("Second").unwrap();
    let added = add.commit().unwrap().into_package();

    let mut remove = added.edit().unwrap();
    assert!(remove.remove("Sheet1").unwrap().is_some());
    let committed = remove.commit().unwrap();
    assert_eq!(committed.package().workbook_ref().len(), 1);
    assert_eq!(
        committed
            .package()
            .workbook_ref()
            .sheet(0)
            .unwrap()
            .unwrap()
            .name(),
        "Second"
    );
}
