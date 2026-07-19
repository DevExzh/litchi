use std::borrow::Cow;

use litchi_rtf::{
    DocumentWriteReservations, LegacyWriteReservation, RtfDocument, RtfWriter, WriteReservationHash,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_modern_hex_and_opaque_legacy_destinations_without_using_them() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\*\writereservhash 00aBff}"#,
        r#"{\*\writereservation old\{weak\}\\value\u20320?}Body}"#,
    ))
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
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\*\writereservation legacy}"#,
        r#"{\*\writereservhash 0aBc}Body}"#,
    ))
    .unwrap();
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r#"{\*\writereservhash 0ABC}"#));
    assert!(
        serialized.find("\\writereservhash").unwrap()
            < serialized.find("\\writereservation").unwrap()
    );
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.write_reservations(), document.write_reservations());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn typed_api_is_passive_clearable_and_coexists_with_adjacent_metadata() {
    let mut document = RtfDocument::parse(
        r#"{\rtf1{\*\wgrffmtfilter 2002}\stylesortmethod4\readonlyrecommended Body}"#,
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
        r#"{\rtf1{\*\writereservhash}Body}"#,
        r#"{\rtf1{\*\writereservhash 0}Body}"#,
        r#"{\rtf1{\*\writereservhash 0G}Body}"#,
        r#"{\rtf1{\*\writereservhash 00 11}Body}"#,
        r#"{\rtf1{\*\writereservhash \u48?0}Body}"#,
        r#"{\rtf1{\*\writereservhash {00}}Body}"#,
        r#"{\rtf1{\*\writereservhash \bin1 x}Body}"#,
        r#"{\rtf1{\*\writereservation}Body}"#,
        r#"{\rtf1{\*\writereservation \b active}Body}"#,
        r#"{\rtf1{\*\writereservation {nested}}Body}"#,
        r#"{\rtf1{\*\writereservation \bin1 x}Body}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }

    let oversized = "a".repeat(litchi_rtf::MAX_WRITE_RESERVATION_BYTES + 1);
    let legacy = format!(r#"{{\rtf1{{\*\writereservation {oversized}}}Body}}"#);
    assert!(RtfDocument::parse(&legacy).is_err());
    let oversized_hash = "aa".repeat(litchi_rtf::MAX_WRITE_RESERVATION_BYTES + 1);
    let hash = format!(r#"{{\rtf1{{\*\writereservhash {oversized_hash}}}Body}}"#);
    assert!(RtfDocument::parse(&hash).is_err());
}

#[test]
fn rejects_parameters_unstarred_direct_duplicate_nested_and_late_destinations() {
    for source in [
        r#"{\rtf1{\*\writereservhash1 00}Body}"#,
        r#"{\rtf1{\*\writereservation1 old}Body}"#,
        r#"{\rtf1{\writereservhash 00}Body}"#,
        r#"{\rtf1{\writereservation old}Body}"#,
        r#"{\rtf1\writereservhash 00 Body}"#,
        r#"{\rtf1\writereservation old Body}"#,
        r#"{\rtf1{\*\writereservhash 00}{\*\writereservhash 11}Body}"#,
        r#"{\rtf1{\*\writereservation one}{\*\writereservation two}Body}"#,
        r#"{\rtf1{{\*\writereservhash 00}}Body}"#,
        r#"{\rtf1 Body{\*\writereservation late}}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
