use litchi_docx::Package;
use litchi_docx::font::Conformance;

#[test]
fn package_delegates_fonts_to_the_canonical_docx_owner() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mut package = Package::open(root.join("test-data/poi/test-data/document/saut_page.docx"))
        .expect("open fixture");
    let fonts = package.fonts().expect("read fonts").expect("font table");

    assert_eq!(fonts.len(), 7);
    assert!(fonts.get("Aptos").expect("lookup").is_some());
    assert!(fonts.get("Arial").is_err());
    assert!(fonts.get(usize::MAX).expect("lookup").is_none());

    package
        .put_fonts(fonts, Conformance::Transitional)
        .expect("move font table");
    assert!(package.fonts().expect("read fonts").is_some());
    assert!(package.remove_fonts().expect("remove fonts"));
    assert!(package.fonts().expect("read fonts").is_none());
    assert!(!package.remove_fonts().expect("already absent"));
}
