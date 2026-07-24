use litchi_rtf::{RtfDocument, RtfWriter, XmlNamespace};
use std::borrow::Cow;
use std::fs;

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_ordered_unicode_namespaces_as_inert_metadata_and_round_trips() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\*\xmlnstbl "#,
        r#"{\xmlns2 urn:example:\u20320?}{\xmlns1 http://schemas.example.test/word}}Body}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "Body");
    let namespaces = document.xml_namespaces().unwrap();
    assert_eq!(namespaces.len(), 2);
    assert_eq!(namespaces[0].id, 2);
    assert_eq!(namespaces[0].namespace, "urn:example:你");
    assert_eq!(namespaces[1].id, 1);

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.xml_namespaces(), Some(namespaces));
}

#[test]
fn mutation_and_empty_table_presence_round_trip() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_xml_namespaces(Vec::new()).unwrap();
    assert_eq!(document.xml_namespaces(), Some([].as_slice()));
    let empty = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(empty.xml_namespaces(), Some([].as_slice()));

    document
        .push_xml_namespace(
            XmlNamespace::new(4, Cow::Borrowed("https://schemas.example.test/x")).unwrap(),
        )
        .unwrap();
    assert!(
        document
            .push_xml_namespace(XmlNamespace::new(4, Cow::Borrowed("urn:duplicate")).unwrap())
            .is_err()
    );
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.xml_namespaces(), document.xml_namespaces());
    assert_eq!(reparsed.text(), "Body");

    document.clear_xml_namespaces();
    assert!(document.xml_namespaces().is_none());
}

#[test]
fn rejects_malformed_or_active_namespace_tables() {
    let cases = [
        r#"{\rtf1{\xmlnstbl {\xmlns1 urn:x}}}"#,
        r#"{\rtf1{\*\xmlnstbl}{\*\xmlnstbl}}"#,
        r#"{\rtf1{\*\xmlnstbl {\xmlns1 urn:x}{\xmlns1 urn:y}}}"#,
        r#"{\rtf1{\*\xmlnstbl {\xmlns0 urn:x}}}"#,
        r#"{\rtf1{\*\xmlnstbl {urn:x}}}"#,
        r#"{\rtf1{\*\xmlnstbl {\xmlns1 }}}"#,
        r#"{\rtf1{\*\xmlnstbl \xmlns1 urn:x}}"#,
        r#"{\rtf1{\*\xmlnstbl {\xmlns1 urn:{nested}}}}"#,
        r#"{\rtf1{\*\xmlnstbl {\xmlns1 urn:\b x}}}"#,
        r#"{\rtf1{\*\xmlnstbl {\xmlns1\bin2 xx}}}"#,
        r#"{\rtf1\xmlns1 Body}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
    assert!(XmlNamespace::new(0, Cow::Borrowed("urn:x")).is_err());
    assert!(XmlNamespace::new(1, Cow::Borrowed(" ")).is_err());
}

#[test]
fn parses_bundled_libreoffice_xml_namespace_fixtures() {
    const FIXTURES: &[&str] = &[
        "sw/qa/core/data/rtf/pass/tdf116851.rtf",
        "sw/qa/extras/ooxmlexport/data/tdf154703_framePr2.rtf",
        "sw/qa/extras/odfexport/data/tdf165315.rtf",
        "sw/qa/extras/rtfexport/data/FWDP90_min.rtf",
    ];
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core"
    );
    for fixture in FIXTURES {
        let bytes = fs::read(format!("{root}/{fixture}")).unwrap();
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        let namespaces = document
            .xml_namespaces()
            .unwrap_or_else(|| panic!("fixture exposed no XML namespace table: {fixture}"));
        assert!(namespaces.iter().any(|entry| {
            entry.id == 1
                && entry.namespace == "http://schemas.microsoft.com/office/word/2003/wordml"
        }));
    }
}
