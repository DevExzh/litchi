use std::path::PathBuf;

#[test]
fn doc_and_ppt_facades_discover_unsigned_signature_state() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mut doc =
        litchi_ole::doc::Package::open(root.join("test-data/ole/doc/documentProperties.doc"))
            .unwrap();
    assert!(doc.signatures().unwrap().is_empty());
    let mut ppt =
        litchi_ole::ppt::Package::open(root.join("test-data/ole/ppt/text-margins.ppt")).unwrap();
    assert!(ppt.signatures().unwrap().is_empty());
}
