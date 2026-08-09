#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{IndexPageReference, NavigationEntry, RtfDocument, RtfWriter};

#[test]
fn parses_microsoft_grammar_index_controls_and_hidden_text() {
    let document = RtfDocument::parse(include_str!("fixtures/navigation-index.rtf")).unwrap();
    assert_eq!(document.text(), "Before after");
    assert_eq!(document.navigation_entries().len(), 1);
    let NavigationEntry::Index(entry) = &document.navigation_entries()[0] else {
        panic!("expected index entry");
    };
    assert_eq!(entry.position, "Before ".len());
    assert_eq!(entry.text, "Alpha");
    assert_eq!(entry.index_id, Some(b'A'));
    assert!(entry.bold_page_number);
    assert!(entry.italic_page_number);
    assert!(matches!(
        &entry.page_reference,
        IndexPageReference::ReplacementText(value) if value == "see also"
    ));
    assert_eq!(entry.yomi.as_deref(), Some("Arufa"));
}

#[test]
fn parses_toc_defaults_no_page_and_visible_source_text_once() {
    let document = RtfDocument::parse(include_str!("fixtures/navigation-toc.rtf")).unwrap();
    assert_eq!(document.text(), "ABVisibleC");
    assert_eq!(document.navigation_entries().len(), 2);
    let NavigationEntry::TableOfContents(first) = &document.navigation_entries()[0] else {
        panic!("expected TOC entry");
    };
    assert_eq!(first.position, 1);
    assert_eq!(first.text, "Chapter");
    assert_eq!(first.table_id, b'C');
    assert_eq!(first.level, 2);
    assert!(!first.suppress_page_number);
    let NavigationEntry::TableOfContents(second) = &document.navigation_entries()[1] else {
        panic!("expected TOC entry");
    };
    assert_eq!(second.position, 2);
    assert_eq!(second.text, "Visible");
    assert!(second.suppress_page_number);
}

#[test]
fn round_trips_unicode_safe_text_positions_and_coexisting_markup() {
    let source = concat!(
        r#"{\rtf1\ansi \u20320?"#,
        r#"{\*\bkmkstart B}{\xe\v A\{B\}\\C{\rxe range}}"#,
        r#"x{\*\bkmkend B}{\tc\v\tcl3 \u-10179?\u-8704?}}"#,
    );
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.text(), "你x");
    assert_eq!(document.navigation_entries()[0].position(), "你".len());

    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(&document)
        .unwrap();
    let output = String::from_utf8(output).unwrap();
    assert!(output.contains("\\xe\\v"));
    assert!(output.contains("\\tc\\v"));
    let reparsed = RtfDocument::parse(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.navigation_entries(), document.navigation_entries());
    let actual = reparsed.bookmarks().get("B").unwrap();
    let expected = document.bookmarks().get("B").unwrap();
    assert_eq!(actual.position, expected.position);
    assert_eq!(actual.content, expected.content);
}

#[test]
fn rejects_conflicts_active_content_invalid_parameters_and_bad_structure() {
    for source in [
        r"{\rtf1{\xe}}",
        r"{\rtf1{\xe\xef no-param}}",
        r"{\rtf1{\xe\xef64 bad}}",
        r"{\rtf1{\xe x{\txe one}{\rxe two}}}",
        r"{\rtf1{\xe x{\txe}}}",
        r"{\rtf1{\xe x\yxe}}",
        r"{\rtf1{\xe x{\*\pxe y}}}",
        r"{\rtf1{\xe x{\field danger}}}",
        r"{\rtf1{\xe x{\object danger}}}",
        "{\\rtf1{\\xe x\\bin4 abcd}}",
        r"{\rtf1{\tc\tcf64 bad}}",
        r"{\rtf1{\tc\tcl10 bad}}",
        r"{\rtf1{\tc\tcl2\tcl3 bad}}",
    ] {
        assert!(RtfDocument::parse(source).is_err(), "{source}");
    }
}

#[test]
fn never_evaluates_generated_navigation_fields_or_bookmark_ranges() {
    let document = RtfDocument::parse(
        r"{\rtf1{\xe\v term{\rxe file:///never-opened}}{\field{\*\fldinst INDEX}{\fldrslt stale}}Body}",
    )
    .unwrap();
    assert_eq!(document.text(), "Body");
    assert_eq!(document.navigation_entries().len(), 1);
    assert_eq!(document.fields().len(), 1);
}
