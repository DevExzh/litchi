use std::path::PathBuf;

#[test]
fn doc_facade_exposes_standard_property_sets() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mut doc =
        litchi_doc::Package::open(root.join("test-data/ole/doc/documentProperties.doc")).unwrap();
    assert!(doc.summary_information().unwrap().is_some());
}
