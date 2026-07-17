use std::fs::File;

use litchi_ole::xls::XlsWorkbook;

#[test]
fn parses_poi_simple_shared_string_index() {
    let path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/poi/test-data/spreadsheet/Simple.xls"
    );
    let workbook = XlsWorkbook::new(File::open(path).unwrap()).unwrap();
    let index = workbook.shared_string_index().unwrap().unwrap();
    assert_eq!(index.unique_string_count(), 1);
    assert_eq!(index.strings_per_bucket(), 8);
    assert_eq!(index.buckets().len(), 1);
    let bucket = index.bucket_for_string(0).unwrap();
    assert_eq!(bucket.stream_position(), 1_420);
    assert_eq!(bucket.record_offset(), 12);
    assert_eq!(bucket.record_position(), 1_408);
}
