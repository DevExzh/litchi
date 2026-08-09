#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{DocumentLegacyLayoutCompatibility, RtfDocument, RtfWriter};

const CONTROLS: [&str; 5] = [
    "splytwnine",
    "ftnlytwnine",
    "htmautsp",
    "useltbaln",
    "oldas",
];

fn write(doc: &RtfDocument<'_>) -> String {
    let mut bytes = Vec::new();
    RtfWriter::new(&mut bytes).write_document(doc).unwrap();
    String::from_utf8(bytes).unwrap()
}

#[test]
fn parses_complete_cluster_as_passive_metadata() {
    let doc = RtfDocument::parse(concat!(
        r"{\rtf1\ansi\splytwnine\ftnlytwnine\htmautsp",
        r"\useltbaln\oldas Body}"
    ))
    .unwrap();
    let compatibility = doc.legacy_layout_compatibility();
    assert!(compatibility.do_not_use_word_97_shape_layout);
    assert!(compatibility.use_legacy_footnote_layout);
    assert!(compatibility.use_html_paragraph_auto_spacing);
    assert!(compatibility.preserve_last_tab_alignment);
    assert!(compatibility.use_word_95_auto_spacing);
    assert_eq!(doc.text(), "Body");
    assert!(doc.shapes().is_empty());
    assert!(doc.notes().is_empty());
}

#[test]
fn omission_is_empty_and_serializes_no_legacy_layout_flags() {
    let doc = RtfDocument::parse(r"{\rtf1\ansi Body}").unwrap();
    assert!(doc.legacy_layout_compatibility().is_empty());
    let serialized = write(&doc);
    for name in CONTROLS {
        assert!(!serialized.contains(&format!("\\{name}")));
    }
}

#[test]
fn typed_api_round_trips_in_specification_order_and_clears_inertly() {
    let mut doc = RtfDocument::parse(r"{\rtf1\ansi Body}").unwrap();
    doc.set_legacy_layout_compatibility(DocumentLegacyLayoutCompatibility {
        do_not_use_word_97_shape_layout: true,
        use_legacy_footnote_layout: true,
        use_html_paragraph_auto_spacing: true,
        preserve_last_tab_alignment: true,
        use_word_95_auto_spacing: true,
    });

    let serialized = write(&doc);
    let positions: Vec<_> = CONTROLS
        .iter()
        .map(|name| serialized.find(&format!("\\{name}")).unwrap())
        .collect();
    assert!(positions.windows(2).all(|pair| pair[0] < pair[1]));
    assert!(!serialized.contains("\\tx"));
    assert!(!serialized.contains("\\footnote"));
    assert!(!serialized.contains("\\shp"));
    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert_eq!(
        *reparsed.legacy_layout_compatibility(),
        *doc.legacy_layout_compatibility()
    );
    assert_eq!(reparsed.text(), "Body");

    doc.clear_legacy_layout_compatibility();
    assert!(doc.legacy_layout_compatibility().is_empty());
    assert_eq!(doc.text(), "Body");
}

#[test]
fn accepts_each_flag_independently_without_layout_side_effects() {
    for name in CONTROLS {
        let doc = RtfDocument::parse(&format!(r"{{\rtf1\ansi\{name} Body}}")).unwrap();
        assert_eq!(doc.text(), "Body");
        assert!(doc.shapes().is_empty());
        assert!(doc.notes().is_empty());
    }
}

#[test]
fn rejects_parameters_duplicates_starred_grouped_and_late_flags() {
    for name in CONTROLS {
        for input in [
            format!(r"{{\rtf1\ansi\{name}0 Body}}"),
            format!(r"{{\rtf1\ansi\{name}\{name} Body}}"),
            format!(r"{{\rtf1\ansi{{\*\{name}}}Body}}"),
            format!(r"{{\rtf1\ansi{{\{name}}}Body}}"),
            format!(r"{{\rtf1\ansi Body\{name}}}"),
        ] {
            assert!(RtfDocument::parse(&input).is_err(), "accepted {input}");
        }
    }
    assert!(RtfDocument::parse(r"{\rtf1\ansi\oldas999999999999 Body}").is_err());
}

#[test]
fn parses_bundled_libreoffice_legacy_layout_flags() {
    let common_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf161878.rtf"
    );
    let common = RtfDocument::parse_bytes(&std::fs::read(common_path).unwrap()).unwrap();
    let compatibility = common.legacy_layout_compatibility();
    assert!(compatibility.do_not_use_word_97_shape_layout);
    assert!(compatibility.use_legacy_footnote_layout);
    assert!(compatibility.preserve_last_tab_alignment);

    let html_path = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/fdo82006.rtf"
    );
    let html = RtfDocument::parse_bytes(&std::fs::read(html_path).unwrap()).unwrap();
    assert!(
        html.legacy_layout_compatibility()
            .use_html_paragraph_auto_spacing
    );
}
