#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use std::borrow::Cow;

use litchi_rtf::raw::Tree;
use litchi_rtf::{
    DocumentWriteReservations, LegacyWriteReservation, WriteReservationHash, write::Writer,
};

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn parses_modern_hex_and_opaque_legacy_destinations_without_using_them() {
    let document = Tree::parse(
        concat!(
            r#"{\rtf1{\*\writereservhash 00aBff}"#,
            r#"{\*\writereservation old\{weak\}\\value\u20320?}Body}"#,
        ),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(
        document
            .write_reservations()
            .hash
            .as_ref()
            .unwrap()
            .data
            .as_ref(),
        &[0x00, 0xab, 0xff]
    );
    assert_eq!(
        document.write_reservations().legacy.as_ref().unwrap().data,
        "old{weak}\\value你"
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn writer_canonicalizes_hash_and_round_trips_both_independent_values() {
    let document = Tree::parse(
        concat!(
            r#"{\rtf1{\*\writereservation legacy}"#,
            r#"{\*\writereservhash 0aBc}Body}"#,
        ),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r"{\*\writereservhash 0ABC}"));
    assert!(
        serialized.find("\\writereservhash").unwrap()
            < serialized.find("\\writereservation").unwrap()
    );
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(reparsed.write_reservations(), document.write_reservations());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn typed_api_is_passive_clearable_and_coexists_with_adjacent_metadata() {
    let mut document = Tree::parse(
        r"{\rtf1{\*\wgrffmtfilter 2002}\stylesortmethod4\readonlyrecommended Body}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    document
        .set_write_reservations(DocumentWriteReservations {
            legacy: Some(LegacyWriteReservation::new(Cow::Borrowed("opaque")).unwrap()),
            hash: Some(WriteReservationHash::new(Cow::Borrowed(&[1, 2, 3])).unwrap()),
        })
        .unwrap();
    let serialized = String::from_utf8(write(&document)).unwrap();
    for (first, second) in [
        ("\\wgrffmtfilter", "\\stylesortmethod4"),
        ("\\stylesortmethod4", "\\writereservhash"),
        ("\\writereservhash", "\\writereservation"),
        ("\\writereservation", "\\readonlyrecommended"),
    ] {
        assert!(serialized.find(first).unwrap() < serialized.find(second).unwrap());
    }
    assert_eq!(document.text(), "Body");
    document.clear_write_reservations();
    assert!(document.write_reservations().is_empty());
}

#[test]
fn rejects_invalid_hash_and_active_or_oversized_payloads() {
    for source in [
        r"{\rtf1{\*\writereservhash}Body}",
        r"{\rtf1{\*\writereservhash 0}Body}",
        r"{\rtf1{\*\writereservhash 0G}Body}",
        r"{\rtf1{\*\writereservhash 00 11}Body}",
        r"{\rtf1{\*\writereservhash \u48?0}Body}",
        r"{\rtf1{\*\writereservhash {00}}Body}",
        r"{\rtf1{\*\writereservhash \bin1 x}Body}",
        r"{\rtf1{\*\writereservation}Body}",
        r"{\rtf1{\*\writereservation \b active}Body}",
        r"{\rtf1{\*\writereservation {nested}}Body}",
        r"{\rtf1{\*\writereservation \bin1 x}Body}",
    ] {
        assert!(
            Tree::parse(source, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted malformed {source}"
        );
    }

    let oversized = "a".repeat(litchi_rtf::MAX_WRITE_RESERVATION_BYTES + 1);
    let legacy = format!(r"{{\rtf1{{\*\writereservation {oversized}}}Body}}");
    assert!(Tree::parse(&legacy, litchi_rtf::read::OpenOptions::default()).is_err());
    let oversized_hash = "aa".repeat(litchi_rtf::MAX_WRITE_RESERVATION_BYTES + 1);
    let hash = format!(r"{{\rtf1{{\*\writereservhash {oversized_hash}}}Body}}");
    assert!(Tree::parse(&hash, litchi_rtf::read::OpenOptions::default()).is_err());
}

#[test]
fn rejects_parameters_unstarred_direct_duplicate_nested_and_late_destinations() {
    for source in [
        r"{\rtf1{\*\writereservhash1 00}Body}",
        r"{\rtf1{\*\writereservation1 old}Body}",
        r"{\rtf1{\writereservhash 00}Body}",
        r"{\rtf1{\writereservation old}Body}",
        r"{\rtf1\writereservhash 00 Body}",
        r"{\rtf1\writereservation old Body}",
        r"{\rtf1{\*\writereservhash 00}{\*\writereservhash 11}Body}",
        r"{\rtf1{\*\writereservation one}{\*\writereservation two}Body}",
        r"{\rtf1{{\*\writereservhash 00}}Body}",
        r"{\rtf1 Body{\*\writereservation late}}",
    ] {
        assert!(
            Tree::parse(source, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted malformed {source}"
        );
    }
}
