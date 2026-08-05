use litchi_xlsx::Workbook;
use tempfile::NamedTempFile;

#[test]
fn opened_workbooks_edit_through_an_explicit_snapshot_transaction() {
    let source = NamedTempFile::with_suffix(".xlsx").unwrap();
    Workbook::create().unwrap().save(source.path()).unwrap();

    let opened = Workbook::open(source.path()).unwrap();
    assert_eq!(opened.len(), 1);
    let mut edit = opened.edit().unwrap();
    edit.add("Edited").unwrap();
    let committed = edit.commit().unwrap();

    assert_eq!(opened.len(), 1);
    assert_eq!(committed.workbook().len(), 2);
    assert_eq!(
        committed
            .workbook()
            .sheet("Edited")
            .unwrap()
            .unwrap()
            .name(),
        "Edited"
    );
}
