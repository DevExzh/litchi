#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{DocumentDataStore, RtfDocument, RtfWriter};
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
fn parses_inert_data_store_bytes_and_round_trips() {
    let document =
        RtfDocument::parse(r"{\rtf1{\*\datastore 01050000 0200000018000000}Body}").unwrap();
    assert_eq!(document.text(), "Body");
    let store = document.data_store().unwrap();
    assert_eq!(store.data.as_ref(), [1, 5, 0, 0, 2, 0, 0, 0, 24, 0, 0, 0]);

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r"{\*\datastore 010500000200000018000000}"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.data_store(), Some(store));
}

#[test]
fn mutation_validates_and_clear_preserves_body() {
    let store = DocumentDataStore::new(Cow::Borrowed(b"opaque\0bytes")).unwrap();
    let mut document = RtfDocument::parse(r"{\rtf1 Text}").unwrap();
    document.set_data_store(store.clone()).unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.data_store(), Some(&store));
    assert_eq!(reparsed.text(), "Text");

    document.clear_data_store();
    assert!(document.data_store().is_none());
    assert_eq!(document.text(), "Text");
    assert!(DocumentDataStore::new(Cow::Borrowed(&[])).is_err());
}

#[test]
fn rejects_malformed_or_active_data_store_payloads() {
    let cases = [
        r"{\rtf1{\datastore 00}}",
        r"{\rtf1{\*\datastore 00}{\*\datastore 01}}",
        r"{\rtf1{\*\datastore }}",
        r"{\rtf1{\*\datastore 0}}",
        r"{\rtf1{\*\datastore 0x}}",
        r"{\rtf1{\*\datastore 00{11}}}",
        r"{\rtf1{\*\datastore 00\b 11}}",
        r"{\rtf1{\*\datastore\bin2 xx}}",
        r"{\rtf1\datastore 00}",
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}

#[test]
fn parses_bundled_libreoffice_data_store_fixtures() {
    const FIXTURES: &[&str] = &[
        "sw/qa/core/data/rtf/pass/tdf116851.rtf",
        "sw/qa/extras/ooxmlexport/data/tdf154703_framePr2.rtf",
        "sw/qa/extras/odfexport/data/tdf165315.rtf",
        "sw/qa/extras/rtfexport/data/tdf158830.rtf",
    ];
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core"
    );
    for fixture in FIXTURES {
        let bytes = fs::read(format!("{root}/{fixture}")).unwrap();
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        let store = document
            .data_store()
            .unwrap_or_else(|| panic!("fixture exposed no datastore: {fixture}"));
        assert!(store.data.starts_with(&[1, 5, 0, 0]));
        assert!(store.data.len() > 32);
    }
}
