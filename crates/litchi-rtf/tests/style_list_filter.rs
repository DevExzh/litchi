#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{DocumentStyleListFilter, DocumentXslTransformUsage, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_specification_example_as_typed_passive_flags() {
    let document = RtfDocument::parse(r"{\rtf1{\*\wgrffmtfilter 2002}Body}").unwrap();
    let filter = document.style_list_filter().unwrap();
    assert_eq!(filter.bits(), 0x2002);
    assert!(filter.contains(DocumentStyleListFilter::CUSTOM_STYLES));
    assert!(filter.contains(DocumentStyleListFilter::TOP_LEVEL_HEADING_STYLES));
    assert!(!filter.contains(DocumentStyleListFilter::ALL_STYLES));
    assert_eq!(document.text(), "Body");
}

#[test]
fn typed_api_and_writer_canonicalize_uppercase_exact_width_hex() {
    let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    let filter = DocumentStyleListFilter::ALL_STYLES
        .union(DocumentStyleListFilter::TABLE_STYLES)
        .union(DocumentStyleListFilter::ALTERNATE_STYLE_NAMES);
    document.set_style_list_filter(filter).unwrap();
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r"{\*\wgrffmtfilter 8081}"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.style_list_filter(), Some(filter));
    assert_eq!(reparsed.text(), "Body");
    document.clear_style_list_filter();
    assert!(document.style_list_filter().is_none());
}

#[test]
fn explicit_zero_and_coexisting_transform_metadata_round_trip_independently() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\*\xform transform.xsl}\usexform"#,
        r#"{\*\wgrffmtfilter 0000}Body}"#,
    ))
    .unwrap();
    assert!(document.style_list_filter().unwrap().is_empty());
    assert_eq!(
        document.xsl_transform_usage(),
        DocumentXslTransformUsage::Requested
    );
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.find("\\xform").unwrap() < serialized.find("\\usexform").unwrap());
    assert!(serialized.find("\\usexform").unwrap() < serialized.find("\\wgrffmtfilter").unwrap());
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.style_list_filter(), document.style_list_filter());
    assert_eq!(
        reparsed.xsl_transform_usage(),
        document.xsl_transform_usage()
    );
}

#[test]
fn rejects_reserved_bit_invalid_width_text_parameters_and_active_content() {
    for source in [
        r"{\rtf1{\*\wgrffmtfilter}Body}",
        r"{\rtf1{\*\wgrffmtfilter 002}Body}",
        r"{\rtf1{\*\wgrffmtfilter 00001}Body}",
        r"{\rtf1{\*\wgrffmtfilter 00G1}Body}",
        r"{\rtf1{\*\wgrffmtfilter 0001 }Body}",
        r"{\rtf1{\*\wgrffmtfilter2002}Body}",
        r"{\rtf1{\*\wgrffmtfilter \b 0001}Body}",
        r"{\rtf1{\*\wgrffmtfilter {0001}}Body}",
        r"{\rtf1{\*\wgrffmtfilter \bin4 0001}Body}",
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
    assert!(DocumentStyleListFilter::from_bits(0x0010).is_err());
}

#[test]
fn preserves_undefined_reserved_bit_from_real_world_fixture_until_modified() {
    let source = include_str!("../../../test-data/rtf/watermark.rtf");
    let document = RtfDocument::parse(source).unwrap();
    let filter = document.style_list_filter().unwrap();
    assert_ne!(filter.bits() & 0x0010, 0);

    let mut destination = Vec::new();
    RtfWriter::new(&mut destination)
        .write_style_list_filter(Some(filter))
        .unwrap();
    assert_eq!(
        String::from_utf8(destination).unwrap(),
        format!(r"{{\*\wgrffmtfilter {:04X}}}", filter.bits())
    );

    let modified = filter.union(DocumentStyleListFilter::ALL_STYLES);
    let mut rejected = Vec::new();
    assert!(
        RtfWriter::new(&mut rejected)
            .write_style_list_filter(Some(modified))
            .is_err()
    );
    assert!(DocumentStyleListFilter::from_bits(filter.bits()).is_err());
}

#[test]
fn rejects_unstarred_direct_duplicate_nested_and_late_destinations() {
    for source in [
        r"{\rtf1{\wgrffmtfilter 0001}Body}",
        r"{\rtf1\wgrffmtfilter 0001 Body}",
        r"{\rtf1{\*\wgrffmtfilter 0001}{\*\wgrffmtfilter 0002}Body}",
        r"{\rtf1{{\*\wgrffmtfilter 0001}}Body}",
        r"{\rtf1 Body{\*\wgrffmtfilter 0001}}",
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
