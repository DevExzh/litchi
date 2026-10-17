//! The iWork semantic namespaces hold types only (record 0334): the facade's
//! one opener is `litchi::open`, and format-specific reads go through the
//! leaf openers the format modules re-export. These tests read each fixture
//! both ways and keep the leaf readers' typed wrong-family routing.

#[cfg(feature = "pages")]
#[test]
fn pages_reads_through_the_unified_and_leaf_openers() {
    use litchi::pages::Package;
    use litchi::pages::semantic::{Document, PackageError};

    let read = |bytes: &[u8]| Package::from_bytes(bytes).map(|package| package.document());

    let bytes = include_bytes!("../../../test-data/iwork/pages/basic.pages");
    let document: Document = read(bytes).expect("Pages fixture must decode");
    assert_eq!(document.section_count(), 1);
    assert!(
        document
            .plain_text()
            .contains("Litchi native Pages fixture")
    );

    let keynote = include_bytes!("../../../test-data/iwork/keynote/basic.key");
    assert!(matches!(
        read(keynote).expect_err("Keynote must not route as Pages"),
        PackageError::UnexpectedFormat {
            detected: litchi::FileFormat::Keynote
        }
    ));

    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/pages/basic.pages");
    assert_eq!(
        Package::open(&path)
            .expect("Pages path reader must decode")
            .document()
            .section_count(),
        1
    );

    // The same file through the facade's one opener.
    let file = litchi::open(&path).expect("Pages opens through litchi::open");
    assert_eq!(file.format(), Some(litchi::FileFormat::Pages));
    assert!(
        file.text()
            .expect("Pages text")
            .contains("Litchi native Pages fixture")
    );
}

#[cfg(feature = "keynote")]
#[test]
fn keynote_reads_through_the_unified_and_leaf_openers() {
    use litchi::keynote::semantic::{Presentation, presentation};

    let bytes = include_bytes!("../../../test-data/iwork/keynote/basic.key");
    let document = Presentation::from_bytes(bytes).expect("Keynote fixture must decode");
    assert_eq!(document.slides().len(), 1);
    assert!(
        document
            .text()
            .expect("Keynote text read must succeed")
            .contains("Litchi native Keynote fixture")
    );

    let pages = include_bytes!("../../../test-data/iwork/pages/basic.pages");
    assert_eq!(
        Presentation::from_bytes(pages).expect_err("Pages must not route as Keynote"),
        presentation::Error::UnexpectedFormat {
            detected: litchi::FileFormat::Pages
        }
    );

    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/keynote/basic.key");
    assert_eq!(
        Presentation::open(&path)
            .expect("Keynote path reader must decode")
            .slides()
            .len(),
        1
    );

    // The same file through the facade's one opener.
    let file = litchi::open(&path).expect("Keynote opens through litchi::open");
    assert_eq!(file.format(), Some(litchi::FileFormat::Keynote));
    assert!(
        file.text()
            .expect("Keynote text")
            .contains("Litchi native Keynote fixture")
    );
}

#[cfg(feature = "numbers")]
#[test]
fn numbers_reads_through_the_unified_and_leaf_openers() {
    use litchi::numbers::semantic::{
        ReadError as DocumentReadError, SheetSelector, TableSelector, Workbook,
    };

    let bytes = include_bytes!("../../../test-data/iwork/numbers/basic.numbers");
    let document = Workbook::from_bytes(bytes).expect("Numbers fixture must decode");
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
        Workbook::from_bytes(pages).expect_err("Pages must not route as Numbers"),
        DocumentReadError::UnexpectedFormat {
            detected: litchi::FileFormat::Pages,
        }
    );

    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/iwork/numbers/basic.numbers");
    assert_eq!(
        Workbook::open(&path)
            .expect("Numbers path reader must decode")
            .sheet_count(),
        1
    );

    // The same file through the facade's one opener.
    let file = litchi::open(&path).expect("Numbers opens through litchi::open");
    assert_eq!(file.format(), Some(litchi::FileFormat::Numbers));
    assert!(
        file.text()
            .expect("Numbers text")
            .contains("Litchi native Numbers fixture")
    );
}
