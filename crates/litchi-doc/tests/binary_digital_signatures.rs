use std::path::PathBuf;

#[test]
fn doc_facade_discovers_unsigned_signature_state() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mut doc =
        litchi_doc::Package::open(root.join("test-data/ole/doc/documentProperties.doc")).unwrap();
    assert!(doc.signatures().unwrap().is_empty());
}
