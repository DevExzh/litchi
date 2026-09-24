//! Focused tests for the document owner.

#[cfg(test)]
mod lazy_auxiliary_tests {
    use crate::package::Package;
    use crate::parts::fib::FileInformationBlock;
    use crate::tracked_revision::Limits;
    use crate::writer::Writer;
    use litchi_ole_common::object::{Editor as PackageEditor, Targets};
    use std::io::Cursor;

    #[test]
    fn malformed_optional_auxiliary_tables_are_deferred_until_access() {
        let mut writer = Writer::new();
        writer.add_paragraph("Body").expect("fixture paragraph");
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("fixture DOC");
        let mut package =
            PackageEditor::open(output.into_inner(), Targets::default(), Limits::default())
                .expect("fixture package");
        let word_path = ["WordDocument".to_string()];
        let mut word = package
            .stream(&word_path)
            .expect("WordDocument stream")
            .to_vec();
        for index in [27usize, 28, 29, 30, 60, 99, 109] {
            let pair = 154 + index * 8;
            word[pair..pair + 4].copy_from_slice(&0u32.to_le_bytes());
            word[pair + 4..pair + 8].copy_from_slice(&1u32.to_le_bytes());
        }
        let pgp_pair = 154 + crate::parts::paragraph_groups::FIB_INDEX_PGP * 8;
        word[pgp_pair + 4..pgp_pair + 8].copy_from_slice(
            &u32::try_from(crate::parts::paragraph_groups::MAX_PGP_BYTES + 1)
                .expect("PGP limit")
                .to_le_bytes(),
        );
        let dofr_pair = 154 + crate::parts::dofr::FIB_INDEX_RG_DOFR * 8;
        word[dofr_pair..dofr_pair + 4].copy_from_slice(&u32::MAX.to_le_bytes());
        package
            .put_stream(&word_path, word)
            .expect("malformed optional pointers");
        let bytes = package.finish().expect("fixture package finish");

        let mut package = Package::from_reader(Cursor::new(bytes)).expect("package open");
        let document = package
            .document()
            .expect("unrelated document open must not parse optional tables");
        assert!(document.paragraph_groups_source.is_err());
        assert!(document.dofr_records_source.is_err());
        assert!(document.saved_selection().is_err());
        let paragraph_error = document
            .paragraph_groups()
            .expect_err("malformed PGP metadata");
        assert_eq!(
            document
                .paragraph_groups()
                .expect_err("cached malformed PGP metadata")
                .to_string(),
            paragraph_error.to_string()
        );
        let dofr_error = document
            .dofr_records()
            .expect_err("malformed RgDofr metadata");
        assert_eq!(
            document
                .dofr_records()
                .expect_err("cached malformed RgDofr metadata")
                .to_string(),
            dofr_error.to_string()
        );
        let print_error = document
            .print_environment()
            .expect_err("malformed print metadata");
        assert_eq!(
            document
                .print_environment()
                .expect_err("cached malformed print metadata")
                .to_string(),
            print_error.to_string()
        );
        assert!(document.vba_signatures().is_err());
    }

    #[test]
    fn deferred_auxiliary_sources_retain_only_selected_ranges() {
        let mut writer = Writer::new();
        writer.add_paragraph("Body").expect("fixture paragraph");
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("fixture DOC");
        let mut package =
            PackageEditor::open(output.into_inner(), Targets::default(), Limits::default())
                .expect("fixture package");
        let word_path = ["WordDocument".to_string()];
        let mut word = package
            .stream(&word_path)
            .expect("WordDocument stream")
            .to_vec();
        let fib = FileInformationBlock::parse(&word).expect("fixture FIB");
        let table_name = if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        };
        let table_path = [table_name.to_string()];
        let mut table = package.stream(&table_path).expect("table stream").to_vec();
        let pgp_offset = u32::try_from(table.len()).expect("table offset");
        let pgp = [0u8, 0u8];
        table.extend_from_slice(&pgp);
        table.extend(std::iter::repeat_n(0xA5, 1024 * 1024));
        package
            .put_stream(&table_path, table.clone())
            .expect("expanded table stream");

        for index in [
            crate::parts::print_environment::FIB_INDEX_PR_DRVR,
            crate::parts::print_environment::FIB_INDEX_PR_ENV_PORT,
            crate::parts::print_environment::FIB_INDEX_PR_ENV_LAND,
            crate::parts::dofr::FIB_INDEX_RG_DOFR,
        ] {
            let pointer = 154 + index * 8;
            word[pointer..pointer + 8].fill(0);
        }
        let pointer = 154 + crate::parts::paragraph_groups::FIB_INDEX_PGP * 8;
        word[pointer..pointer + 4].copy_from_slice(&pgp_offset.to_le_bytes());
        word[pointer + 4..pointer + 8]
            .copy_from_slice(&(u32::try_from(pgp.len()).expect("PGP length")).to_le_bytes());
        package.put_stream(&word_path, word).expect("PGP pointer");
        let bytes = package.finish().expect("fixture package finish");

        let mut package = Package::from_reader(Cursor::new(bytes)).expect("package open");
        let document = package.document().expect("document open");
        let source = document
            .paragraph_groups_source
            .as_ref()
            .expect("PGP range capture")
            .as_ref()
            .expect("PGP range present");
        assert_eq!(source, &pgp);
        assert!(source.len() < table.len());
        assert!(
            document
                .paragraph_groups()
                .expect("lazy PGP parse")
                .expect("PGP metadata")
                .is_empty()
        );
        assert!(matches!(&document.dofr_records_source, Ok(None)));
    }
}

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
        crate::VersionedDocumentProperties::Word2002(_)
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
