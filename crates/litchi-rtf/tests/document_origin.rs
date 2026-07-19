use litchi_rtf::{
    DocumentAutoFormatType, DocumentOrigin, DocumentOriginMetadata, HtmlEmailVersion, RtfDocument,
    RtfWriter,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_plain_text_and_both_documented_html_origin_forms() {
    let text = RtfDocument::parse(r#"{\rtf1\fromtext\deff0 Body}"#).unwrap();
    assert_eq!(
        text.origin_metadata().origin,
        Some(DocumentOrigin::PlainTextEmail)
    );
    let html = RtfDocument::parse(r#"{\rtf1\fromhtml\deff0 Body}"#).unwrap();
    assert_eq!(
        html.origin_metadata().origin,
        Some(DocumentOrigin::HtmlEmail { version: None })
    );
    let html1 = RtfDocument::parse(r#"{\rtf1\fromhtml1\deff0 Body}"#).unwrap();
    assert_eq!(
        html1.origin_metadata().origin,
        Some(DocumentOrigin::HtmlEmail {
            version: Some(HtmlEmailVersion::Version1),
        })
    );
    assert_eq!(html1.text(), "Body");
}

#[test]
fn parses_all_document_types_and_parameterless_default() {
    for (source, expected) in [
        (r#"{\rtf1\doctype Body}"#, DocumentAutoFormatType::General),
        (r#"{\rtf1\doctype0 Body}"#, DocumentAutoFormatType::General),
        (r#"{\rtf1\doctype1 Body}"#, DocumentAutoFormatType::Letter),
        (r#"{\rtf1\doctype2 Body}"#, DocumentAutoFormatType::Email),
    ] {
        let document = RtfDocument::parse(source).unwrap();
        assert_eq!(document.origin_metadata().auto_format_type, Some(expected));
        assert_eq!(
            document.origin_metadata().effective_auto_format_type(),
            expected
        );
    }
    let omitted = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    assert_eq!(omitted.origin_metadata().auto_format_type, None);
    assert_eq!(
        omitted.origin_metadata().effective_auto_format_type(),
        DocumentAutoFormatType::General
    );
}

#[test]
fn writer_places_origin_before_default_font_and_round_trips_typed_api() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_origin_metadata(DocumentOriginMetadata {
        origin: Some(DocumentOrigin::HtmlEmail {
            version: Some(HtmlEmailVersion::Version1),
        }),
        auto_format_type: Some(DocumentAutoFormatType::Email),
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.find("\\ansicpg").unwrap() < serialized.find("\\fromhtml1").unwrap());
    assert!(serialized.find("\\fromhtml1").unwrap() < serialized.find("\\deff").unwrap());
    assert!(serialized.contains("\\doctype2"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.origin_metadata(), document.origin_metadata());
    assert_eq!(reparsed.text(), "Body");
    document.clear_origin_metadata();
    assert!(document.origin_metadata().is_empty());
}

#[test]
fn origin_and_classification_coexist_independently_with_root_metadata() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\fromtext\deff0{\*\wgrffmtfilter 2002}"#,
        r#"\stylesortmethod4\doctype2\readonlyrecommended Body}"#,
    ))
    .unwrap();
    assert_eq!(
        document.origin_metadata().origin,
        Some(DocumentOrigin::PlainTextEmail)
    );
    assert_eq!(
        document.origin_metadata().auto_format_type,
        Some(DocumentAutoFormatType::Email)
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn rejects_invalid_versions_types_conflicts_starred_grouped_and_late_forms() {
    for source in [
        r#"{\rtf1\fromtext0 Body}"#,
        r#"{\rtf1\fromhtml0 Body}"#,
        r#"{\rtf1\fromhtml2 Body}"#,
        r#"{\rtf1\fromhtml-1 Body}"#,
        r#"{\rtf1\fromtext\fromhtml1 Body}"#,
        r#"{\rtf1\fromhtml1\fromhtml1 Body}"#,
        r#"{\rtf1\doctype-1 Body}"#,
        r#"{\rtf1\doctype3 Body}"#,
        r#"{\rtf1\doctype1\doctype2 Body}"#,
        r#"{\rtf1{\*\fromtext}Body}"#,
        r#"{\rtf1{\*\fromhtml1}Body}"#,
        r#"{\rtf1{\*\doctype2}Body}"#,
        r#"{\rtf1{\fromtext}Body}"#,
        r#"{\rtf1{\doctype2}Body}"#,
        r#"{\rtf1{\fonttbl{\f0 Arial;}}\fromtext Body}"#,
        r#"{\rtf1 Body\fromtext}"#,
        r#"{\rtf1 Body\doctype2}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
