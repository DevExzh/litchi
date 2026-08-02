use std::path::PathBuf;

#[test]
fn ppt_facade_exposes_standard_property_sets() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mut ppt =
        litchi_ppt::Package::open(root.join("test-data/poi/test-data/slideshow/text-margins.ppt"))
            .unwrap();
    let _ = ppt.document_summary_information().unwrap();
}
