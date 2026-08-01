use litchi_rtf::{RtfDocument, RtfWriter, StyleType};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

const TABLE_STYLE_SOURCE: &str = concat!(
    r#"{\rtf1\ansi{\stylesheet{\ql Normal;}"#,
    r#"{\*\ts16\tsrowd\b \tscfirstrow\tsclastrow\tscfirstcol\tsclastcol"#,
    r#"\tscbandhorzodd\tscbandverteven\tscbandsh2\tscbandsv3 Table List;}}Body}"#,
);

#[test]
fn parses_table_style_conditional_formatting_and_round_trips() {
    let document = RtfDocument::parse(TABLE_STYLE_SOURCE).unwrap();
    assert_eq!(document.text(), "Body");
    let style = document
        .stylesheet()
        .styles()
        .iter()
        .find(|style| style.style_type == StyleType::Table)
        .expect("table style");
    let conditional = style.table_conditional;
    assert!(conditional.row_defaults_marker);
    assert!(conditional.first_row);
    assert!(conditional.last_row);
    assert!(conditional.first_column);
    assert!(conditional.last_column);
    assert!(conditional.band_horizontal_odd);
    assert!(!conditional.band_horizontal_even);
    assert!(!conditional.band_vertical_odd);
    assert!(conditional.band_vertical_even);
    assert_eq!(conditional.horizontal_band_size, Some(2));
    assert_eq!(conditional.vertical_band_size, Some(3));

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    for control in [
        "\\tsrowd",
        "\\tscfirstrow",
        "\\tsclastrow",
        "\\tscfirstcol",
        "\\tsclastcol",
        "\\tscbandhorzodd",
        "\\tscbandverteven",
        "\\tscbandsh2",
        "\\tscbandsv3",
    ] {
        assert!(
            serialized.contains(control),
            "missing {control} in {serialized}"
        );
    }

    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    let reparsed_style = reparsed
        .stylesheet()
        .styles()
        .iter()
        .find(|style| style.style_type == StyleType::Table)
        .unwrap();
    assert_eq!(reparsed_style.table_conditional, conditional);
}

#[test]
fn table_style_without_conditional_metadata_stays_empty() {
    let document =
        RtfDocument::parse(r#"{\rtf1\ansi{\stylesheet{\ql Normal;}{\*\ts16\b Table List;}}Body}"#)
            .unwrap();
    let style = document
        .stylesheet()
        .styles()
        .iter()
        .find(|style| style.style_type == StyleType::Table)
        .unwrap();
    assert!(style.table_conditional.is_empty());

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert!(
        reparsed.stylesheet().styles()[1]
            .table_conditional
            .is_empty()
    );
}

#[test]
fn rejects_misplaced_or_duplicate_table_style_controls() {
    let cases = [
        // Conditional controls in a paragraph style.
        r#"{\rtf1{\stylesheet{\s1\tscfirstrow Para;}}}"#,
        r#"{\rtf1{\stylesheet{\s1\tsrowd Para;}}}"#,
        r#"{\rtf1{\stylesheet{\s1\tscbandsh2 Para;}}}"#,
        // Conditional controls in a character style.
        r#"{\rtf1{\stylesheet{\*\cs1\tscfirstcol Char;}}}"#,
        // Duplicate controls.
        r#"{\rtf1{\stylesheet{\*\ts16\tscfirstrow\tscfirstrow T;}}}"#,
        r#"{\rtf1{\stylesheet{\*\ts16\tsrowd\tsrowd T;}}}"#,
        r#"{\rtf1{\stylesheet{\*\ts16\tscbandsh2\tscbandsh3 T;}}}"#,
        // Parameterized flag controls.
        r#"{\rtf1{\stylesheet{\*\ts16\tscfirstrow1 T;}}}"#,
        r#"{\rtf1{\stylesheet{\*\ts16\tsrowd0 T;}}}"#,
        // Missing band-size parameter.
        r#"{\rtf1{\stylesheet{\*\ts16\tscbandsh T;}}}"#,
        // Out-of-range band size.
        r#"{\rtf1{\stylesheet{\*\ts16\tscbandsv65536 T;}}}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}
