//! Regression coverage for the small archive-free reader entry points exposed
//! by the umbrella's format semantic namespaces.

#[cfg(feature = "pages")]
#[test]
fn pages_reader_helpers_keep_semantic_values_and_typed_routing() {
    use litchi::pages::semantic::{self, ReadError};

    let bytes = include_bytes!("../../../test-data/iwork/pages/basic.pages");
    let document = semantic::from_bytes(bytes).expect("Pages fixture must decode");
    assert_eq!(document.section_count(), 1);
    assert!(
        document
            .plain_text()
            .contains("Litchi native Pages fixture")
    );

    let keynote = include_bytes!("../../../test-data/iwork/keynote/basic.key");
    assert_eq!(
        semantic::from_bytes(keynote).expect_err("Keynote must not route as Pages"),
        ReadError::NotPages
    );

    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/basic.pages");
    assert_eq!(
        semantic::open(path)
            .expect("Pages path reader must decode")
            .section_count(),
        1
    );
}

#[cfg(feature = "keynote")]
#[test]
fn keynote_reader_helpers_keep_semantic_values_and_typed_routing() {
    use litchi::keynote::semantic::{self, DocumentReadError};

    let bytes = include_bytes!("../../../test-data/iwork/keynote/basic.key");
    let document = semantic::from_bytes(bytes).expect("Keynote fixture must decode");
    assert_eq!(document.slides().len(), 1);
    assert!(
        document
            .text()
            .expect("Keynote text read must succeed")
            .contains("Litchi native Keynote fixture")
    );

    let pages = include_bytes!("../../../test-data/iwork/pages/basic.pages");
    assert_eq!(
        semantic::from_bytes(pages).expect_err("Pages must not route as Keynote"),
        DocumentReadError::NotKeynote
    );

    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/basic.key");
    assert_eq!(
        semantic::open(path)
            .expect("Keynote path reader must decode")
            .slides()
            .len(),
        1
    );
}

#[cfg(feature = "numbers")]
#[test]
fn numbers_reader_helpers_keep_semantic_values_and_typed_routing() {
    use litchi::numbers::semantic::{self, DocumentReadError, SheetSelector, TableSelector};

    let bytes = include_bytes!("../../../test-data/iwork/numbers/basic.numbers");
    let document = semantic::from_bytes(bytes).expect("Numbers fixture must decode");
    assert_eq!(document.sheet_count(), 1);
    let table = document
        .table(
            SheetSelector::name("Sheet 1"),
            TableSelector::name("Table 1"),
        )
        .expect("Numbers table lookup must retain its typed selector result")
        .expect("Numbers fixture must have its named table");
    assert_eq!(table.name(), "Table 1");
    assert!(
        document
            .plain_text()
            .expect("Numbers text read must succeed")
            .contains("Litchi native Numbers fixture")
    );

    let pages = include_bytes!("../../../test-data/iwork/pages/basic.pages");
    assert_eq!(
        semantic::from_bytes(pages).expect_err("Pages must not route as Numbers"),
        DocumentReadError::NotNumbers
    );

    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/basic.numbers");
    assert_eq!(
        semantic::open(path)
            .expect("Numbers path reader must decode")
            .sheet_count(),
        1
    );
}
