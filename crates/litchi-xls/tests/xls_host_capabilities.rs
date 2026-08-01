use std::fs::File;
use std::path::PathBuf;

use litchi_xls::XlsWorkbook;

#[test]
fn facade_exposes_signatures_and_standard_property_sets() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");

    let mut signed = XlsWorkbook::new(
        File::open(root.join("test-data/poi/test-data/spreadsheet/Simple.xls"))
            .expect("XLS signature fixture"),
    )
    .expect("valid XLS fixture");
    assert!(
        signed
            .signatures()
            .expect("inspect inert signatures")
            .is_empty()
    );

    let mut metadata = XlsWorkbook::new(
        File::open(root.join("test-data/ole/xls/Simple.xls")).expect("XLS metadata fixture"),
    )
    .expect("valid XLS fixture");
    let _ = metadata
        .summary_information()
        .expect("inspect property set");
}
