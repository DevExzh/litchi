//! Focused [MS-OSHARED] composite-value codec tests.

use super::super::super::model::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, DocParts, HeadingPair, HeadingPairs,
    PID_DOC_PARTS, PID_HEADING_PAIRS, Section, Stream, TextEncoding, Value,
};
use super::parse_typed_property_for_property;
use litchi_cfb::consts::{VT_I4, VT_LPSTR, VT_VARIANT, VT_VECTOR};

fn typed(variant_type: u16, body: &[u8]) -> Vec<u8> {
    let mut value = Vec::with_capacity(4 + body.len());
    value.extend_from_slice(&variant_type.to_le_bytes());
    value.extend_from_slice(&0u16.to_le_bytes());
    value.extend_from_slice(body);
    value
}

fn append_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn append_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

#[test]
fn round_trips_ansi_heading_pairs_and_unaligned_doc_parts() {
    let headings = HeadingPairs::new(vec![
        HeadingPair::new("Title", 1).expect("valid title heading"),
        HeadingPair::new("Headings", 2).expect("valid headings heading"),
    ])
    .expect("valid heading pairs");
    let parts = DocParts::ansi(vec![
        "Document title".into(),
        "Heading 1".into(),
        "Heading 2".into(),
    ])
    .expect("valid document parts");
    let mut section = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    section.set_page(CodePage::WINDOWS_1252);
    section
        .add(PID_HEADING_PAIRS, Value::HeadingPairs(headings.clone()))
        .expect("heading pairs should be accepted");
    section
        .add(PID_DOC_PARTS, Value::DocParts(parts.clone()))
        .expect("document parts should be accepted");

    let parsed = Stream::new(section)
        .to_bytes()
        .and_then(|bytes| Stream::parse(&bytes))
        .expect("composite values should round-trip");
    assert_eq!(
        parsed.sections[0].property(PID_HEADING_PAIRS),
        Some(&Value::HeadingPairs(headings))
    );
    assert_eq!(
        parsed.sections[0].property(PID_DOC_PARTS),
        Some(&Value::DocParts(parts))
    );
}

#[test]
fn parses_the_checked_in_heading_pair_wire_example() {
    let mut body = Vec::new();
    append_u32(&mut body, 4);
    append_u16(&mut body, VT_LPSTR);
    append_u16(&mut body, 0);
    append_u32(&mut body, 6);
    body.extend_from_slice(b"Title\0");
    append_u16(&mut body, VT_I4);
    append_u16(&mut body, 0);
    body.extend_from_slice(&1i32.to_le_bytes());
    append_u16(&mut body, VT_LPSTR);
    append_u16(&mut body, 0);
    append_u32(&mut body, 9);
    body.extend_from_slice(b"Headings\0");
    append_u16(&mut body, VT_I4);
    append_u16(&mut body, 0);
    body.extend_from_slice(&8i32.to_le_bytes());

    let value = parse_typed_property_for_property(
        &typed(VT_VECTOR | VT_VARIANT, &body),
        1252,
        0,
        PID_HEADING_PAIRS,
    )
    .expect("the MS-OSHARED heading-pair example should parse");
    let expected = HeadingPairs::new(vec![
        HeadingPair::new("Title", 1).unwrap(),
        HeadingPair::new("Headings", 8).unwrap(),
    ])
    .unwrap();
    assert_eq!(value, Value::HeadingPairs(expected));
}

#[test]
fn round_trips_unicode_composites_with_relative_string_padding() {
    let headings = HeadingPairs::new(vec![
        HeadingPair::new("章节", 1).expect("valid Unicode heading"),
    ])
    .expect("valid Unicode heading pairs");
    let parts = DocParts::new(TextEncoding::Unicode, vec!["标题".into()])
        .expect("valid Unicode document parts");
    let mut section = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    section.set_page(CodePage::Utf16Le);
    section
        .add(PID_HEADING_PAIRS, Value::HeadingPairs(headings.clone()))
        .expect("heading pairs should be accepted");
    section
        .add(PID_DOC_PARTS, Value::DocParts(parts.clone()))
        .expect("document parts should be accepted");

    let parsed = Stream::parse(&Stream::new(section).to_bytes().unwrap()).unwrap();
    assert_eq!(
        parsed.sections[0].property(PID_HEADING_PAIRS),
        Some(&Value::HeadingPairs(headings))
    );
    assert_eq!(
        parsed.sections[0].property(PID_DOC_PARTS),
        Some(&Value::DocParts(parts))
    );
}

#[test]
fn parses_unaligned_ansi_parts_without_inventing_padding() {
    let mut body = Vec::new();
    append_u32(&mut body, 2);
    append_u32(&mut body, 2);
    body.extend_from_slice(b"A\0");
    append_u32(&mut body, 4);
    body.extend_from_slice(b"BCD\0");
    let value = parse_typed_property_for_property(
        &typed(VT_VECTOR | VT_LPSTR, &body),
        1252,
        0,
        PID_DOC_PARTS,
    )
    .expect("unaligned document parts should parse");
    assert_eq!(
        value,
        Value::DocParts(DocParts::ansi(vec!["A".into(), "BCD".into()]).unwrap())
    );
}

#[test]
fn rejects_odd_heading_counts_and_bounded_count_amplification() {
    let odd = typed(VT_VECTOR | VT_VARIANT, &3u32.to_le_bytes());
    assert!(
        parse_typed_property_for_property(&odd, 1252, 0, PID_HEADING_PAIRS).is_err(),
        "heading cElements must be even"
    );

    let too_many_parts = typed(VT_VECTOR | VT_LPSTR, &(1_000_001u32).to_le_bytes());
    assert!(
        parse_typed_property_for_property(&too_many_parts, 1252, 0, PID_DOC_PARTS).is_err(),
        "document-part vectors must be bounded before allocation"
    );
}

#[test]
fn preserves_generic_variant_vectors_outside_the_well_known_pid() {
    let mut body = Vec::new();
    append_u32(&mut body, 1);
    body.extend_from_slice(&VT_I4.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&7i32.to_le_bytes());
    let value =
        parse_typed_property_for_property(&typed(VT_VECTOR | VT_VARIANT, &body), 1252, 0, 42)
            .expect("generic variant vector should remain supported");
    assert!(matches!(value, Value::Vector(_)));
}

#[test]
fn preserves_unknown_property_values_outside_the_well_known_pids() {
    let value = parse_typed_property_for_property(&typed(0x7777, &[1, 2, 3]), 1252, 0, 42)
        .expect("unknown property values should remain readable");
    assert_eq!(
        value,
        Value::Unknown {
            variant_type: 0x7777,
            data: vec![1, 2, 3],
        }
    );
}

#[test]
fn rejects_mismatched_heading_totals() {
    let mut section = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    section
        .add(
            PID_HEADING_PAIRS,
            Value::HeadingPairs(
                HeadingPairs::new(vec![HeadingPair::new("Title", 2).unwrap()]).unwrap(),
            ),
        )
        .unwrap();
    section
        .add(
            PID_DOC_PARTS,
            Value::DocParts(DocParts::ansi(vec!["one".into()]).unwrap()),
        )
        .unwrap();
    assert!(Stream::new(section).to_bytes().is_err());
}
