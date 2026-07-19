use litchi_rtf::{RevisionSaveMetadata, RtfDocument, RtfWriter};
use std::fs;

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn preserves_ordered_revision_save_table_and_root_round_trip() {
    let document = RtfDocument::parse(
        r#"{\rtf1\ansi{\*\rsidtbl \rsid7564464\rsid8398352\rsid9049968}\rsidroot9049968 Body}"#,
    )
    .unwrap();
    assert_eq!(document.text(), "Body");
    let metadata = document.revision_save_metadata().unwrap();
    assert_eq!(metadata.ids(), [7_564_464, 8_398_352, 9_049_968]);
    assert_eq!(metadata.root(), Some(9_049_968));

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.revision_save_metadata(), Some(metadata));
}

#[test]
fn preserves_valid_empty_revision_save_table() {
    let document = RtfDocument::parse(r#"{\rtf1{\*\rsidtbl}\rsidroot7 Body}"#).unwrap();
    let metadata = document.revision_save_metadata().unwrap();
    assert!(metadata.ids().is_empty());
    assert_eq!(metadata.root(), Some(7));
    assert_eq!(document.text(), "Body");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.revision_save_metadata(), Some(metadata));
    assert_eq!(reparsed.text(), "Body");
    assert!(RevisionSaveMetadata::new(Vec::new(), None).is_ok());
    assert!(RevisionSaveMetadata::new(Vec::new(), Some(7)).is_ok());
}

#[test]
fn mutation_validates_membership_uniqueness_and_preserves_body() {
    let mut metadata = RevisionSaveMetadata::new(vec![11, 22], Some(11)).unwrap();
    metadata.push_id(33).unwrap();
    metadata.set_root(Some(33)).unwrap();
    assert!(metadata.push_id(22).is_err());
    assert!(metadata.set_root(Some(44)).is_err());

    let mut document = RtfDocument::parse(r#"{\rtf1 Text}"#).unwrap();
    document
        .set_revision_save_metadata(metadata.clone())
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), "Text");
    assert_eq!(reparsed.revision_save_metadata(), Some(&metadata));

    document.clear_revision_save_metadata();
    assert!(document.revision_save_metadata().is_none());
    assert_eq!(document.text(), "Text");
}

#[test]
fn rejects_malformed_revision_save_metadata() {
    let cases = [
        r#"{\rtf1{\rsidtbl \rsid1}\rsidroot1}"#,
        r#"{\rtf1{\*\rsidtbl \rsid1\rsid1}\rsidroot1}"#,
        r#"{\rtf1{\*\rsidtbl \rsid0}}"#,
        r#"{\rtf1{\*\rsidtbl \rsid1}\rsidroot2}"#,
        r#"{\rtf1{\*\rsidtbl \rsid1}\rsidroot1\rsidroot1}"#,
        r#"{\rtf1{\*\rsidtbl \rsid1}{\b\rsidroot1}}"#,
        r#"{\rtf1\rsid1}"#,
        r#"{\rtf1{\*\rsidtbl text\rsid1}}"#,
        r#"{\rtf1{\*\rsidtbl {\rsid1}}}"#,
        r#"{\rtf1{\*\rsidtbl \bin2 xx}}"#,
        r#"{\rtf1{\*\rsidtbl \b\rsid1}}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}

#[test]
fn parses_bundled_libreoffice_revision_save_fixtures() {
    const FIXTURES: &[&str] = &[
        "sw/qa/core/data/rtf/pass/tdf116851.rtf",
        "sw/qa/extras/ooxmlexport/data/ooo39250-1-min.rtf",
        "sw/qa/extras/ooxmlexport/data/tdf141173_missingFrames.rtf",
        "sw/qa/writerfilter/filters-test/data/pass/TCI-TN65GP-DDRHDLL-partial.rtf",
    ];
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/"
    );
    for fixture in FIXTURES {
        let bytes = fs::read(format!("{root}{fixture}")).unwrap();
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        let metadata = document
            .revision_save_metadata()
            .unwrap_or_else(|| panic!("fixture exposed no revision-save table: {fixture}"));
        assert!(!metadata.ids().is_empty());
        assert!(
            metadata
                .root()
                .is_some_and(|root| metadata.ids().contains(&root))
        );
    }
}
