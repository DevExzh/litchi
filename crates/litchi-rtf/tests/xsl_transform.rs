#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use std::borrow::Cow;

use litchi_rtf::{DocumentXslTransform, DocumentXslTransformUsage, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_passive_transform_location_and_decodes_destination_text() {
    let document =
        RtfDocument::parse(r"{\rtf1{\*\xform file:///C:\\transform\u20320?\{v1\}.xsl}Body}")
            .unwrap();
    assert_eq!(
        document.xsl_transform().unwrap().location,
        "file:///C:\\transform你{v1}.xsl"
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn writer_emits_required_starred_destination_and_round_trips() {
    let document = RtfDocument::parse(r"{\rtf1{\*\xform  relative path.xsl }Body}").unwrap();
    assert_eq!(
        document.xsl_transform().unwrap().location,
        " relative path.xsl "
    );

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r"{\*\xform  relative path.xsl }"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.xsl_transform(), document.xsl_transform());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn typed_api_stores_but_never_resolves_or_executes_location() {
    let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    document
        .set_xsl_transform(
            DocumentXslTransform::new(Cow::Borrowed("https://invalid.example/transform.xsl"))
                .unwrap(),
        )
        .unwrap();
    assert_eq!(document.text(), "Body");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.xsl_transform(), document.xsl_transform());
    assert_eq!(reparsed.text(), "Body");

    document.clear_xsl_transform();
    assert!(document.xsl_transform().is_none());
}

#[test]
fn rejects_unstarred_misplaced_duplicate_empty_active_nested_and_binary_forms() {
    for source in [
        r"{\rtf1{\xform unstarred.xsl}Body}",
        r"{\rtf1\xform direct.xsl Body}",
        r"{\rtf1{\*\xform}Body}",
        r"{\rtf1{\*\xform one.xsl}{\*\xform two.xsl}Body}",
        r"{\rtf1 Body{\*\xform late.xsl}}",
        r"{\rtf1{{\*\xform nested.xsl}}Body}",
        r"{\rtf1{\*\xform outer{inner}}Body}",
        r"{\rtf1{\*\xform \b active}Body}",
        r"{\rtf1{\*\xform \bin1 x}Body}",
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn enforces_model_and_parser_resource_bounds() {
    assert!(DocumentXslTransform::new(Cow::Borrowed("line\nbreak.xsl")).is_err());
    assert!(DocumentXslTransform::new(Cow::Borrowed("\0")).is_err());

    let oversized = "x".repeat(litchi_rtf::MAX_DOCUMENT_XSL_TRANSFORM_LOCATION_BYTES + 1);
    let source = format!(r"{{\rtf1{{\*\xform {oversized}}}Body}}");
    assert!(RtfDocument::parse(&source).is_err());
}

#[test]
fn usexform_preserves_requested_intent_and_writes_after_location() {
    let document = RtfDocument::parse(r"{\rtf1{\*\xform transform.xsl}\usexform Body}").unwrap();
    assert_eq!(
        document.xsl_transform_usage(),
        DocumentXslTransformUsage::Requested
    );
    assert_eq!(document.text(), "Body");

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.find("\\xform").unwrap() < serialized.find("\\usexform").unwrap());
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.xsl_transform(), document.xsl_transform());
    assert_eq!(
        reparsed.xsl_transform_usage(),
        document.xsl_transform_usage()
    );
}

#[test]
fn requested_intent_is_inert_and_independent_of_location() {
    let mut document = RtfDocument::parse(r"{\rtf1\usexform Body}").unwrap();
    assert!(document.xsl_transform().is_none());
    assert!(document.xsl_transform_usage().is_requested());
    assert_eq!(document.text(), "Body");

    document.set_xsl_transform_usage(DocumentXslTransformUsage::NotRequested);
    assert!(!document.xsl_transform_usage().is_requested());
    document.set_xsl_transform_usage(DocumentXslTransformUsage::Requested);
    document.clear_xsl_transform();
    assert!(document.xsl_transform_usage().is_requested());
    document.clear_xsl_transform_usage();
    assert!(!document.xsl_transform_usage().is_requested());
}

#[test]
fn rejects_usexform_parameters_duplicates_starred_grouped_and_late_forms() {
    for source in [
        r"{\rtf1\usexform0 Body}",
        r"{\rtf1\usexform1 Body}",
        r"{\rtf1\usexform-1 Body}",
        r"{\rtf1\usexform\usexform Body}",
        r"{\rtf1{\*\usexform}Body}",
        r"{\rtf1{\usexform}Body}",
        r"{\rtf1 Body\usexform}",
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
