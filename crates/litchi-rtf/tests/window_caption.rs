use std::borrow::Cow;

use litchi_rtf::{DocumentWindowCaption, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_starred_and_legacy_unstarred_string_destinations() {
    let starred =
        RtfDocument::parse(r#"{\rtf1{\*\windowcaption Release \u20320? \{A\} \\ status}Body}"#)
            .unwrap();
    assert_eq!(
        starred.window_caption().unwrap().text,
        "Release 你 {A} \\ status"
    );
    assert_eq!(starred.text(), "Body");

    let unstarred = RtfDocument::parse(r#"{\rtf1{\windowcaption Legacy title}Body}"#).unwrap();
    assert_eq!(unstarred.window_caption().unwrap().text, "Legacy title");
    assert_eq!(unstarred.text(), "Body");
}

#[test]
fn writer_canonicalizes_to_starred_destination_and_preserves_string_spacing() {
    let document = RtfDocument::parse(r#"{\rtf1{\windowcaption  padded }Body}"#).unwrap();
    assert_eq!(document.window_caption().unwrap().text, " padded ");

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r#"{\*\windowcaption  padded }"#));
    assert!(!serialized.contains(r#"{\windowcaption"#));

    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.window_caption(), document.window_caption());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn typed_api_is_inert_and_clearable() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document
        .set_window_caption(
            DocumentWindowCaption::new(Cow::Borrowed("file:///never-opened")).unwrap(),
        )
        .unwrap();
    assert_eq!(
        document.window_caption().unwrap().text,
        "file:///never-opened"
    );
    assert_eq!(document.text(), "Body");

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.window_caption(), document.window_caption());
    assert_eq!(reparsed.text(), "Body");

    document.clear_window_caption();
    assert!(document.window_caption().is_none());
}

#[test]
fn rejects_misplaced_duplicate_empty_active_nested_and_binary_destinations() {
    for source in [
        r#"{\rtf1\windowcaption direct Body}"#,
        r#"{\rtf1{\windowcaption}Body}"#,
        r#"{\rtf1{\windowcaption One}{\*\windowcaption Two}Body}"#,
        r#"{\rtf1 Body{\windowcaption Late}}"#,
        r#"{\rtf1{{\windowcaption Nested}}Body}"#,
        r#"{\rtf1{\windowcaption Outer {inner}}Body}"#,
        r#"{\rtf1{\windowcaption \b active}Body}"#,
        r#"{\rtf1{\windowcaption \bin1 x}Body}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn enforces_caption_model_and_parser_resource_bounds() {
    assert!(DocumentWindowCaption::new(Cow::Borrowed("line\nbreak")).is_err());
    assert!(DocumentWindowCaption::new(Cow::Borrowed("\0")).is_err());

    let oversized = "x".repeat(litchi_rtf::MAX_WINDOW_CAPTION_BYTES + 1);
    let source = format!(r#"{{\rtf1{{\*\windowcaption {oversized}}}Body}}"#);
    assert!(RtfDocument::parse(&source).is_err());
}
