use litchi_ole::xls::{
    XlsHorizontalAlignment, XlsReadingOrder, XlsTextRotation, XlsVerticalAlignment, XlsWorkbook,
};
use std::fs::File;
use std::path::PathBuf;

fn poi_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../3rdparty/poi/test-data/spreadsheet")
        .join(name)
}

#[test]
fn reads_shrink_to_fit_xf_metadata() {
    let workbook = XlsWorkbook::new(File::open(poi_fixture("ShrinkToFit.xls")).unwrap()).unwrap();
    let formats = workbook.extended_formats();

    assert!(formats
        .iter()
        .any(|format| format.alignment().shrinks_to_fit()));
    assert!(formats
        .iter()
        .any(|format| !format.alignment().shrinks_to_fit()));
}

#[test]
fn retains_default_formatting_alignment() {
    let workbook = XlsWorkbook::new(File::open(poi_fixture("Formatting.xls")).unwrap()).unwrap();
    let alignment = workbook.extended_formats()[0].alignment();

    assert_eq!(alignment.horizontal(), XlsHorizontalAlignment::General);
    assert_eq!(alignment.vertical(), XlsVerticalAlignment::Bottom);
    assert!(!alignment.wraps_text());
    assert!(!alignment.justifies_last_line());
    assert_eq!(alignment.rotation(), XlsTextRotation::None);
    assert_eq!(alignment.indent(), 0);
    assert!(!alignment.shrinks_to_fit());
    assert_eq!(alignment.reading_order(), XlsReadingOrder::Context);
}
