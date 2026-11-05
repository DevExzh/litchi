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

    package = {
        let mut edit = package.edit_fonts();
        edit.put(fonts, Conformance::Transitional)
            .expect("move font table");
        edit.commit().expect("move font table").into_package()
    };
    assert!(package.fonts().expect("read fonts").is_some());
    assert!({
        let mut edit = package.edit_fonts();
        let value = edit.remove().expect("remove fonts");
        package = edit.commit().expect("remove fonts").into_package();
        value
    });
    assert!(package.fonts().expect("read fonts").is_none());
    assert!(!{
        let mut edit = package.edit_fonts();
        let value = edit.remove().expect("already absent");
        package = edit.commit().expect("already absent").into_package();
        value
    });
    assert!(package.fonts().expect("read fonts").is_none());
}
