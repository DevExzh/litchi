use litchi_rtf::{DocumentTheme, RtfDocument, RtfWriter};
use std::borrow::Cow;
use std::fs;

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_inert_theme_bytes_and_round_trips_without_interpretation() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\*\themedata 504B03041400}"#,
        r#"{\*\colorschememapping 3C3F786D6C3E}Body}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "Body");
    let theme = document.theme().unwrap();
    assert_eq!(theme.data.as_ref(), [0x50, 0x4b, 0x03, 0x04, 0x14, 0x00]);
    assert_eq!(
        theme.color_scheme_mapping.as_deref(),
        Some(b"<?xml>".as_slice())
    );

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r#"{\*\themedata 504b03041400}"#));
    assert!(serialized.contains(r#"{\*\colorschememapping 3c3f786d6c3e}"#));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.theme(), Some(theme));
}

#[test]
fn mutation_validates_and_clear_preserves_body() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Text}"#).unwrap();
    let theme = DocumentTheme::new(
        Cow::Borrowed(b"PK\x03\x04"),
        Some(Cow::Borrowed(b"<mapping/>")),
    )
    .unwrap();
    document.set_theme(theme.clone()).unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.theme(), Some(&theme));
    assert_eq!(reparsed.text(), "Text");

    document.clear_theme();
    assert!(document.theme().is_none());
    assert_eq!(document.text(), "Text");
    assert!(DocumentTheme::new(Cow::Borrowed(&[]), None).is_err());
    assert!(DocumentTheme::new(Cow::Borrowed(b"x"), Some(Cow::Borrowed(&[]))).is_err());
}

#[test]
fn rejects_malformed_or_active_theme_payloads() {
    let cases = [
        r#"{\rtf1{\themedata 00}}"#,
        r#"{\rtf1{\*\themedata 00}{\*\themedata 01}}"#,
        r#"{\rtf1{\*\colorschememapping 00}}"#,
        r#"{\rtf1{\*\themedata }}"#,
        r#"{\rtf1{\*\themedata 0}}"#,
        r#"{\rtf1{\*\themedata 0x}}"#,
        r#"{\rtf1{\*\themedata 00{11}}}"#,
        r#"{\rtf1{\*\themedata 00\b 11}}"#,
        r#"{\rtf1{\*\themedata\bin2 xx}}"#,
        r#"{\rtf1\themedata 00}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}

#[test]
fn parses_bundled_libreoffice_theme_fixtures() {
    const FIXTURES: &[&str] = &[
        "sw/qa/core/data/rtf/pass/tdf116851.rtf",
        "sw/qa/extras/ooxmlexport/data/tdf154703_framePr2.rtf",
        "sw/qa/extras/odfexport/data/tdf165315.rtf",
        "sw/qa/extras/rtfexport/data/tdf158830.rtf",
    ];
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core"
    );
    for fixture in FIXTURES {
        let bytes = fs::read(format!("{root}{fixture}")).unwrap();
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        let theme = document
            .theme()
            .unwrap_or_else(|| panic!("fixture exposed no theme: {fixture}"));
        assert!(theme.data.starts_with(b"PK\x03\x04"));
        assert!(
            theme
                .color_scheme_mapping
                .as_deref()
                .is_some_and(|mapping| mapping.starts_with(b"<?xml"))
        );
    }
}
