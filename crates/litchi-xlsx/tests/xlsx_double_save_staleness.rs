use litchi_xlsx::Workbook;

#[test]
fn repeated_snapshot_publication_is_byte_stable() {
    let workbook = Workbook::create().unwrap();
    let first = workbook.to_bytes().unwrap();
    let second = workbook.to_bytes().unwrap();
    assert_eq!(first, second);
}
