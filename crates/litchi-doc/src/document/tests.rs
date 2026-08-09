//! Focused tests for the document owner.

#[cfg(all(test, feature = "formula"))]
mod owned_mtef_tests {
    use crate::Document;
    use std::collections::HashMap;
    use std::sync::Arc;

    #[test]
    fn malformed_multiple_formulas_are_independently_owned_and_dropped() {
        let mut inputs = HashMap::new();
        inputs.insert("equation-a".to_string(), vec![0xAA; 7]);
        inputs.insert("equation-b".to_string(), vec![0xBB; 13]);

        let rendered = Document::parse_all_mtef_data(&inputs).expect("malformed formulas render");
        assert_eq!(rendered.len(), 2);
        assert!(rendered["equation-a"].contains("Invalid MTEF format"));
        assert!(rendered["equation-b"].contains("Invalid MTEF format"));
        assert!(!Arc::ptr_eq(
            &rendered["equation-a"],
            &rendered["equation-b"]
        ));

        let retained = Arc::clone(&rendered["equation-a"]);
        let weak = Arc::downgrade(&retained);
        drop(retained);
        drop(rendered);
        assert!(weak.upgrade().is_none());
    }
}

use crate::package::{Error as PackageError, Package};
use crate::parts::fib::WORD_97_NFIB;
use crate::{Image, ImageError, Writer};
use std::io::Cursor;
use std::path::Path;

#[test]
fn test_extract_png_image_from_doc() {
    let base = Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap();
    let doc_path = base
        .join("test-data")
        .join("ole")
        .join("doc")
        .join("PngPicture.doc");

    let mut pkg = Package::open(&doc_path).expect("open doc");
    let doc = pkg.document().expect("load document");
    const PNG_SIGNATURE: [u8; 8] = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];

    let mut found_signature = doc
        .word_document
        .windows(PNG_SIGNATURE.len())
        .any(|window| window == PNG_SIGNATURE);

    if let Some(data_stream) = doc.data_stream.as_ref() {
        found_signature |= data_stream
            .windows(PNG_SIGNATURE.len())
            .any(|window| window == PNG_SIGNATURE);
    }

    assert!(
        found_signature,
        "expected PNG signature in document streams"
    );
}

#[test]
fn opened_document_exposes_versioned_document_properties() {
    let mut writer = Writer::new();
    writer.add_paragraph("Body").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut package = Package::from_reader(Cursor::new(output.into_inner())).unwrap();
    let document = package.document().expect("load document");
    let properties = document
        .document_properties()
        .expect("valid DopBase")
        .expect("document carries a Dop");

    assert_eq!(
        properties.to_bytes().unwrap().len(),
        properties.version().byte_len()
    );
    assert!(matches!(
        properties
            .versioned()
            .expect("valid versioned Dop extension"),
        crate::VersionedDocumentProperties::Word97(_)
    ));
}

#[test]
fn test_image_data_with_invalid_offset() {
    let base = Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap();
    let doc_path = base
        .join("test-data")
        .join("ole")
        .join("doc")
        .join("PngPicture.doc");

    let mut pkg = Package::open(&doc_path).expect("open doc");
    let doc = pkg.document().expect("load document");

    let img = Image::new(u32::MAX);
    let err = doc.image_data(&img).expect_err("expected invalid offset");
    assert!(matches!(err, ImageError::InvalidPicOffset(_)));
}

/// Word 6.0 and Word 95 keep the structures MS-DOC places in a table
/// stream inside `WordDocument`, so they have no `0Table`/`1Table`. The
/// reader used to report a bare "Stream not found: 0Table", which told the
/// caller nothing; it must name the format generation instead. Apache POI
/// reaches the same diagnosis at the same point.
#[test]
fn word_6_documents_report_their_version_not_a_missing_stream() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc/word6-no-table-stream.doc");
    let mut package = Package::open(&path).expect("the CFB container opens");

    match package.document() {
        Err(PackageError::UnsupportedVersion { nfib, name }) => {
            assert!(
                nfib < WORD_97_NFIB,
                "expected a pre-Word-97 nFib, got {nfib:#06x}"
            );
            assert!(name.contains("Word 6"), "unexpected version name: {name}");
        },
        Err(other) => panic!("expected an UnsupportedVersion error, got {other:?}"),
        Ok(_) => panic!("expected a Word 6.0 document to be rejected"),
    }
}
