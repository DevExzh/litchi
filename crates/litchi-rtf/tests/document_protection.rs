#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{ProtectionLevel, ProtectionType, RtfDocument, RtfWriter};
use std::fs;

fn round_trip(document: &RtfDocument<'_>) -> RtfDocument<'static> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    RtfDocument::parse_bytes(&output).unwrap()
}

#[test]
fn parses_all_protection_controls_and_round_trips_inert_hash() {
    let source = concat!(
        r#"{\rtf1\ansi{\info{\title Protected}{\*\password aBcD0123}}"#,
        r#"\formprot\annotprot0\revprot\readprot0\allprot"#,
        r#"\enforceprot1\protlevel2 Body}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    let protection = document.protection();
    assert_eq!(protection.forms, Some(true));
    assert_eq!(protection.annotations, Some(false));
    assert_eq!(protection.revisions, Some(true));
    assert_eq!(protection.read_only, Some(false));
    assert_eq!(protection.all, Some(true));
    assert_eq!(protection.enforced, Some(true));
    assert_eq!(protection.level, Some(ProtectionLevel::Level2));
    assert_eq!(protection.password_hash.as_deref(), Some("aBcD0123"));
    assert_eq!(
        protection.protection_type(),
        ProtectionType::RevisionTracking
    );
    assert_eq!(document.text(), "Body");

    let reparsed = round_trip(&document);
    assert_eq!(reparsed.protection(), protection);
    assert_eq!(reparsed.info().title, document.info().title);
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_malformed_duplicate_or_misplaced_protection() {
    for source in [
        r"{\rtf1\formprot2}",
        r"{\rtf1\enforceprot}",
        r"{\rtf1\enforceprot2}",
        r"{\rtf1\protlevel}",
        r"{\rtf1\protlevel4}",
        r"{\rtf1\formprot\formprot}",
        r"{\rtf1{\b\readprot}}",
        r"{\rtf1 Body\revprot}",
        r"{\rtf1{\info{\password 00000000}}}",
        r"{\rtf1{\info{\*\password xyz}}}",
        r"{\rtf1{\info{\*\password 00000000}{\*\password 11111111}}}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }
}

#[test]
fn parses_real_libreoffice_protection_fixtures() {
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras"
    );

    let read_only = fs::read(format!("{root}/rtfimport/data/read-only-protect.rtf")).unwrap();
    let document = RtfDocument::parse_bytes(&read_only).unwrap();
    assert_eq!(document.protection().annotations, Some(true));
    assert_eq!(document.protection().read_only, Some(true));
    assert_eq!(document.protection().enforced, Some(true));
    assert_eq!(document.protection().level, Some(ProtectionLevel::Level3));

    let forms = fs::read(format!("{root}/rtfexport/data/4010_min.rtf")).unwrap();
    let document = RtfDocument::parse_bytes(&forms).unwrap();
    assert_eq!(document.protection().forms, Some(true));
    assert_eq!(document.protection().all, Some(true));
    assert_eq!(document.protection().level, Some(ProtectionLevel::Level2));

    let password = fs::read(format!("{root}/rtfexport/data/fdo55504-1-min.rtf")).unwrap();
    let document = RtfDocument::parse_bytes(&password).unwrap();
    assert_eq!(
        document.protection().password_hash.as_deref(),
        Some("00000000")
    );
}
