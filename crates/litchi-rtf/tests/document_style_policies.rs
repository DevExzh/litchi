#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{DocumentStylePolicies, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_all_four_passive_flags_without_applying_styles() {
    let document = RtfDocument::parse(
        r"{\rtf1\linkstyles\stylelocktheme\stylelockqfset\usenormstyforlist Body}",
    )
    .unwrap();
    assert_eq!(
        *document.style_policies(),
        DocumentStylePolicies {
            update_styles_from_template: true,
            lock_theme: true,
            lock_quick_format_set: true,
            use_normal_style_for_lists: true,
        }
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn parses_common_producer_flag_independently() {
    let document = RtfDocument::parse(r"{\rtf1\usenormstyforlist Body}").unwrap();
    assert!(!document.style_policies().update_styles_from_template);
    assert!(!document.style_policies().lock_theme);
    assert!(!document.style_policies().lock_quick_format_set);
    assert!(document.style_policies().use_normal_style_for_lists);
    assert_eq!(document.text(), "Body");
}

#[test]
fn omission_remains_empty_and_is_not_serialized() {
    let document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    assert!(document.style_policies().is_empty());
    let serialized = String::from_utf8(write(&document)).unwrap();
    assert!(!serialized.contains("stylelocktheme"));
    assert!(!serialized.contains("stylelockqfset"));
    assert!(!serialized.contains("usenormstyforlist"));
}

#[test]
fn typed_api_round_trips_in_stable_order_and_clears_without_style_changes() {
    let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    document.set_style_policies(DocumentStylePolicies {
        update_styles_from_template: true,
        lock_theme: true,
        lock_quick_format_set: true,
        use_normal_style_for_lists: true,
    });
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    let controls = [
        "\\linkstyles",
        "\\stylelocktheme",
        "\\stylelockqfset",
        "\\usenormstyforlist",
    ];
    for pair in controls.windows(2) {
        assert!(serialized.find(pair[0]).unwrap() < serialized.find(pair[1]).unwrap());
    }
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.style_policies(), document.style_policies());
    assert_eq!(reparsed.text(), "Body");

    document.clear_style_policies();
    assert!(document.style_policies().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn coexists_with_theme_data_filter_sort_and_language_metadata() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\stylelockqfset\themelang1033{\*\themedata 0102}"#,
        r#"{\*\wgrffmtfilter 0002}\stylesortmethod3"#,
        r#"\usenormstyforlist\stylelocktheme Body}"#,
    ))
    .unwrap();
    let output = write(&document);
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.style_policies(), document.style_policies());
    assert_eq!(reparsed.theme_languages(), document.theme_languages());
    assert_eq!(reparsed.style_list_filter(), document.style_list_filter());
    assert_eq!(reparsed.style_sort_method(), document.style_sort_method());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn parses_bundled_producer_after_stylesheet_without_loading_a_template() {
    let fixture = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/n751020.rtf"
    );
    let document = RtfDocument::parse_bytes(fixture).unwrap();
    assert!(document.style_policies().update_styles_from_template);
    assert!(document.text().contains("first"));
    assert!(document.text().contains("second"));

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("\\linkstyles"));
    assert!(!serialized.contains("\\template"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.style_policies(), document.style_policies());
    assert!(reparsed.text().contains("first"));
}

#[test]
fn rejects_parameters_duplicates_starred_grouped_and_late_flags() {
    for name in [
        "linkstyles",
        "stylelocktheme",
        "stylelockqfset",
        "usenormstyforlist",
    ] {
        for suffix in ["0", "1", "2147483647", "99999999999"] {
            let source = format!(r"{{\rtf1\{name}{suffix} Body}}");
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
        for source in [
            format!(r"{{\rtf1\{name}\{name} Body}}"),
            format!(r"{{\rtf1{{\*\{name}}}Body}}"),
            format!(r"{{\rtf1{{\{name}}}Body}}"),
            format!(r"{{\rtf1 Body\{name}}}"),
        ] {
            assert!(
                RtfDocument::parse(&source).is_err(),
                "accepted malformed {source}"
            );
        }
    }
}
