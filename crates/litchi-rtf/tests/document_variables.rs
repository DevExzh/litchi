#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{RtfDocument, RtfWriter};

#[test]
fn parses_ordered_document_variables_without_body_leakage() {
    let document = RtfDocument::parse(
        r"{\rtf1\ansi{\*\docvar {FullName}{Jeff Smith}}{\*\docvar {Unused}{Hello World}}Body}",
    )
    .unwrap();
    assert_eq!(document.document_variables().len(), 2);
    assert_eq!(document.document_variables()[0].name, "FullName");
    assert_eq!(document.document_variables()[0].value, "Jeff Smith");
    assert_eq!(document.document_variables()[1].name, "Unused");
    assert_eq!(document.text(), "Body");
}

#[test]
fn parses_libreoffice_document_variable_fixtures_when_available() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sw/qa/extras");
    let fixtures = [
        "rtfexport/data/tdf150267.rtf",
        "rtfexport/data/tdf151370.rtf",
        "rtfexport/data/tdf158762.rtf",
        "rtfimport/data/tdf169298.rtf",
    ];
    if !root.exists() {
        return;
    }
    for fixture in fixtures {
        let source = std::fs::read_to_string(root.join(fixture)).unwrap();
        let document = RtfDocument::parse(&source).unwrap();
        assert!(!document.document_variables().is_empty(), "{fixture}");
    }
    let source = std::fs::read_to_string(root.join("rtfimport/data/tdf169298.rtf")).unwrap();
    let document = RtfDocument::parse(&source).unwrap();
    assert_eq!(document.document_variables()[0].value, "foo\r\nbar");
    assert!(document.text().starts_with("This document contains"));

    let source = std::fs::read_to_string(root.join("rtfexport/data/tdf151370.rtf")).unwrap();
    let document = RtfDocument::parse(&source).unwrap();
    assert_eq!(
        document.document_variables()[0].name,
        "LocalChars\u{c1}rv\u{ed}zturoT\u{fc}k\u{f6}rf\u{fa}r\u{f3}g\u{e9}p"
    );
    assert_eq!(
        document.document_variables()[0].value,
        "\u{e1}rv\u{ed}zturot\u{fc}k\u{f6}rf\u{fa}r\u{f3}g\u{e9}p"
    );
}

#[test]
fn rejects_malformed_or_active_document_variable_content() {
    for source in [
        r"{\rtf1{\*\docvar{}{value}}}",
        r"{\rtf1{\*\docvar{name}}}",
        r"{\rtf1{\*\docvar{name}{value}{extra}}}",
        r"{\rtf1{\*\docvar{name}{{nested}}}}",
        r"{\rtf1{\*\docvar{name}{\bin4 abcd}}}",
        r"{\rtf1{\docvar{name}{value}}}",
        r"{\rtf1{\b{\*\docvar{name}{value}}}}",
        r"{\rtf1 Body{\*\docvar{name}{value}}}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "{source}");
    }
}

#[test]
fn preserves_duplicate_names_as_ordered_inert_entries() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\ansi{\*\docvar {Same}{first}}"#,
        r#"{\*\docvar {Same}{second}}Body}"#,
    ))
    .unwrap();
    assert_eq!(document.document_variables().len(), 2);
    assert_eq!(document.document_variables()[0].name, "Same");
    assert_eq!(document.document_variables()[0].value, "first");
    assert_eq!(document.document_variables()[1].name, "Same");
    assert_eq!(document.document_variables()[1].value, "second");

    let owned = document.document_variables()[0].clone().into_owned();
    assert_eq!(owned, document.document_variables()[0]);
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.document_variables(), document.document_variables());
}

#[test]
fn writer_round_trips_destination_safe_text_and_unicode() {
    let document = RtfDocument::parse(
        "{\\rtf1\\ansi{\\*\\docvar {A\\{B\\}\\\\C}{line\\'0d\\'0a\\u-10179?\\u-8704?}}Body}",
    )
    .unwrap();
    assert_eq!(document.document_variables()[0].name, "A{B}\\C");
    assert_eq!(document.document_variables()[0].value, "line\r\n😀");

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let output = String::from_utf8(output).unwrap();
    assert!(output.contains("{\\*\\docvar"));
    assert!(output.contains("\\'0d\\'0a"));
    let reparsed = RtfDocument::parse(&output).unwrap();
    assert_eq!(reparsed.document_variables(), document.document_variables());
    assert_eq!(reparsed.text(), "Body");
}
