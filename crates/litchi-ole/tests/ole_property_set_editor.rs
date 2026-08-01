use std::path::PathBuf;

#[test]
fn doc_and_ppt_facades_expose_standard_property_sets() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mut doc =
        litchi_ole::doc::Package::open(root.join("test-data/ole/doc/documentProperties.doc"))
            .unwrap();
    assert!(doc.summary_information().unwrap().is_some());
    let mut ppt = litchi_ole::ppt::Package::open(
        root.join("test-data/poi/test-data/slideshow/text-margins.ppt"),
    )
    .unwrap();
    let _ = ppt.document_summary_information().unwrap();
}
