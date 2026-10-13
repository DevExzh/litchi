use tempfile::NamedTempFile;

#[test]
fn opened_workbooks_edit_through_an_explicit_snapshot_transaction() {
    let source = NamedTempFile::with_suffix(".xlsx").unwrap();
    litchi_xlsx::Package::new()
        .unwrap()
        .save(source.path())
        .unwrap();

    let opened = litchi_xlsx::Package::open(source.path()).unwrap();
    assert_eq!(opened.workbook_ref().len(), 1);
    let mut edit = opened.edit().unwrap();
    edit.add("Edited").unwrap();
    let committed = edit.commit().unwrap();

    assert_eq!(opened.workbook_ref().len(), 1);
    assert_eq!(committed.package().workbook_ref().len(), 2);
    assert_eq!(
        committed
            .package()
            .workbook_ref()
            .sheet("Edited")
            .unwrap()
            .unwrap()
            .name(),
        "Edited"
    );
}
