//! A flat text document's packaged projection and the republication of an
//! edit of it (record 0292).

use litchi_odt::flat::Document;
use litchi_odt::metadata::Field;

fn fixture() -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf/odt/title-field-invalidate.fodt"),
    )
    .unwrap()
}

#[test]
fn an_edit_of_the_projection_republishes_into_the_flat_bytes() {
    let source = fixture();
    let flat = Document::from_bytes(&source).unwrap();
    let package = flat.to_package().unwrap();
    let document = package.document().unwrap();
    assert!(document.text().unwrap().contains("body text"));
    // An unchanged projection republishes the flat bytes exactly.
    assert_eq!(flat.republish(&package).unwrap().as_bytes(), source);

    let mut edit = package.edit_metadata();
    edit.set(Field::Title, Some("Projected")).unwrap();
    let edited = edit.commit().unwrap().into_package();
    let republished = flat.republish(&edited).unwrap();
    let text = std::str::from_utf8(republished.as_bytes()).unwrap();
    assert!(text.contains("Projected"));
    assert_eq!(
        republished.metadata().unwrap().title.as_deref(),
        Some("Projected")
    );
    // Only `office:meta` changed: the body and styles keep their bytes.
    let body = |bytes: &[u8]| {
        let text = std::str::from_utf8(bytes).unwrap();
        text[text.find("<office:body>").unwrap()..].to_owned()
    };
    assert_eq!(body(republished.as_bytes()), body(&source));
}
