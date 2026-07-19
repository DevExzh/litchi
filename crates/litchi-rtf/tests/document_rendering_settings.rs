use litchi_rtf::{
    DocumentJustificationMode, DocumentRenderingOrientation, DocumentRenderingSettings,
    RtfDocument, RtfWriter,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_producer_style_horizontal_compressed_grid_flags() {
    let document = RtfDocument::parse(r#"{\rtf1\horzdoc\jcompress\lnongrid Body}"#).unwrap();
    assert_eq!(
        *document.rendering_settings(),
        DocumentRenderingSettings {
            orientation: Some(DocumentRenderingOrientation::Horizontal),
            justification_mode: Some(DocumentJustificationMode::Compress),
            line_based_on_grid: true,
        }
    );
    assert_eq!(document.text(), "Body");
}

#[test]
fn parses_vertical_and_expanding_alternatives() {
    let document = RtfDocument::parse(r#"{\rtf1\vertdoc\jexpand Body}"#).unwrap();
    assert_eq!(
        document.rendering_settings().orientation,
        Some(DocumentRenderingOrientation::Vertical)
    );
    assert_eq!(
        document.rendering_settings().justification_mode,
        Some(DocumentJustificationMode::Expand)
    );
}

#[test]
fn omission_is_distinct_from_the_effective_compression_default() {
    let document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    assert!(document.rendering_settings().is_empty());
    assert_eq!(
        document.rendering_settings().effective_justification_mode(),
        DocumentJustificationMode::Compress
    );
    let serialized = String::from_utf8(write(&document)).unwrap();
    assert!(!serialized.contains("\\jcompress"));
}

#[test]
fn typed_api_round_trips_in_stable_order_and_clears_without_rendering_effects() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document.set_rendering_settings(DocumentRenderingSettings {
        orientation: Some(DocumentRenderingOrientation::Vertical),
        justification_mode: Some(DocumentJustificationMode::Expand),
        line_based_on_grid: true,
    });

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.find("\\vertdoc").unwrap() < serialized.find("\\jexpand").unwrap());
    assert!(serialized.find("\\jexpand").unwrap() < serialized.find("\\lnongrid").unwrap());
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.rendering_settings(), document.rendering_settings());
    assert_eq!(reparsed.text(), "Body");

    document.clear_rendering_settings();
    assert!(document.rendering_settings().is_empty());
    assert_eq!(document.text(), "Body");
}

#[test]
fn coexists_with_adjacent_output_and_file_properties_in_canonical_order() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\jexpand\psover\makebackup\lnongrid"#,
        r#"\horzdoc\muser Body}"#,
    ))
    .unwrap();
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();

    assert!(serialized.find("\\makebackup").unwrap() < serialized.find("\\muser").unwrap());
    assert!(serialized.find("\\muser").unwrap() < serialized.find("\\psover").unwrap());
    assert!(serialized.find("\\psover").unwrap() < serialized.find("\\horzdoc").unwrap());
    assert!(serialized.find("\\horzdoc").unwrap() < serialized.find("\\jexpand").unwrap());
    assert!(serialized.find("\\jexpand").unwrap() < serialized.find("\\lnongrid").unwrap());

    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.rendering_settings(), document.rendering_settings());
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn rejects_numeric_starred_grouped_late_duplicate_and_conflicting_forms() {
    for source in [
        r#"{\rtf1\horzdoc0 Body}"#,
        r#"{\rtf1\vertdoc2147483647 Body}"#,
        r#"{\rtf1\jcompress-2147483648 Body}"#,
        r#"{\rtf1\jexpand1 Body}"#,
        r#"{\rtf1\lnongrid99999999999 Body}"#,
        r#"{\rtf1\horzdoc\horzdoc Body}"#,
        r#"{\rtf1\vertdoc\vertdoc Body}"#,
        r#"{\rtf1\jcompress\jcompress Body}"#,
        r#"{\rtf1\jexpand\jexpand Body}"#,
        r#"{\rtf1\lnongrid\lnongrid Body}"#,
        r#"{\rtf1\horzdoc\vertdoc Body}"#,
        r#"{\rtf1\vertdoc\horzdoc Body}"#,
        r#"{\rtf1\jcompress\jexpand Body}"#,
        r#"{\rtf1\jexpand\jcompress Body}"#,
        r#"{\rtf1{\*\horzdoc}Body}"#,
        r#"{\rtf1{\*\vertdoc}Body}"#,
        r#"{\rtf1{\*\jcompress}Body}"#,
        r#"{\rtf1{\*\jexpand}Body}"#,
        r#"{\rtf1{\*\lnongrid}Body}"#,
        r#"{\rtf1{\horzdoc}Body}"#,
        r#"{\rtf1{\vertdoc}Body}"#,
        r#"{\rtf1{\jcompress}Body}"#,
        r#"{\rtf1{\jexpand}Body}"#,
        r#"{\rtf1{\lnongrid}Body}"#,
        r#"{\rtf1 Body\horzdoc}"#,
        r#"{\rtf1 Body\vertdoc}"#,
        r#"{\rtf1 Body\jcompress}"#,
        r#"{\rtf1 Body\jexpand}"#,
        r#"{\rtf1 Body\lnongrid}"#,
    ] {
        assert!(
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}
