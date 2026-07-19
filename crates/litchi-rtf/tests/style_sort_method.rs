use litchi_rtf::{DocumentStyleSortMethod, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_every_specified_numeric_value() {
    for (value, expected) in [
        (0, DocumentStyleSortMethod::Name),
        (1, DocumentStyleSortMethod::HostDefault),
        (2, DocumentStyleSortMethod::Font),
        (3, DocumentStyleSortMethod::BasedOnStyle),
        (4, DocumentStyleSortMethod::StyleType),
    ] {
        let source = format!(r#"{{\rtf1\stylesortmethod{value} Body}}"#);
        let document = RtfDocument::parse(&source).unwrap();
        assert_eq!(document.style_sort_method(), Some(expected));
        assert_eq!(document.effective_style_sort_method(), expected);
        assert_eq!(document.text(), "Body");
    }
}

#[test]
fn omission_and_missing_parameter_use_host_default_but_preserve_presence() {
    let omitted = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    assert_eq!(omitted.style_sort_method(), None);
    assert_eq!(
        omitted.effective_style_sort_method(),
        DocumentStyleSortMethod::HostDefault
    );

    let present = RtfDocument::parse(r#"{\rtf1\stylesortmethod Body}"#).unwrap();
    assert_eq!(
        present.style_sort_method(),
        Some(DocumentStyleSortMethod::HostDefault)
    );
    let serialized = String::from_utf8(write(&present)).unwrap();
    assert!(serialized.contains(r#"\stylesortmethod1"#));
}

#[test]
fn typed_api_round_trips_and_clears_passive_metadata() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_style_sort_method(DocumentStyleSortMethod::BasedOnStyle);
    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.style_sort_method(), document.style_sort_method());
    assert_eq!(reparsed.text(), "Body");
    document.clear_style_sort_method();
    assert_eq!(document.style_sort_method(), None);
    assert_eq!(
        document.effective_style_sort_method(),
        DocumentStyleSortMethod::HostDefault
    );
}

#[test]
fn coexists_independently_with_filter_and_transform_metadata_in_stable_order() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\*\xform transform.xsl}\usexform"#,
        r#"{\*\wgrffmtfilter 2002}\stylesortmethod4 Body}"#,
    ))
    .unwrap();
    let serialized = String::from_utf8(write(&document)).unwrap();
    assert!(serialized.find("\\xform").unwrap() < serialized.find("\\usexform").unwrap());
    assert!(serialized.find("\\usexform").unwrap() < serialized.find("\\wgrffmtfilter").unwrap());
    assert!(
        serialized.find("\\wgrffmtfilter").unwrap()
            < serialized.find("\\stylesortmethod4").unwrap()
    );
}

#[test]
fn rejects_undefined_duplicates_starred_grouped_and_late_values() {
    for source in [
        r#"{\rtf1\stylesortmethod-1 Body}"#,
        r#"{\rtf1\stylesortmethod5 Body}"#,
        r#"{\rtf1\stylesortmethod2147483647 Body}"#,
        r#"{\rtf1\stylesortmethod1\stylesortmethod2 Body}"#,
        r#"{\rtf1{\*\stylesortmethod2}Body}"#,
        r#"{\rtf1{\stylesortmethod2}Body}"#,
        r#"{\rtf1 Body\stylesortmethod2}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
