//! Round-trip tests for theme-font selectors on font-table entries.

use litchi_rtf::{FontTheme, RtfDocument, RtfWriter};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

const SELECTORS: &[(&str, FontTheme)] = &[
    (r"\flomajor", FontTheme::MajorLatin),
    (r"\fhimajor", FontTheme::MajorHighAnsi),
    (r"\fdbmajor", FontTheme::MajorDoubleByte),
    (r"\fbimajor", FontTheme::MajorBidi),
    (r"\flominor", FontTheme::MinorLatin),
    (r"\fhiminor", FontTheme::MinorHighAnsi),
    (r"\fdbminor", FontTheme::MinorDoubleByte),
    (r"\fbiminor", FontTheme::MinorBidi),
];

#[test]
fn theme_font_selectors_round_trip() {
    for (control, theme) in SELECTORS {
        let source = format!(
            r"{{\rtf1\ansi{{\fonttbl{{\f0\froman{control} Times New Roman;}}}}\f0 Text\par}}"
        );
        let document = RtfDocument::parse(&source).unwrap();
        assert_eq!(document.font_table().get(0).unwrap().theme, Some(*theme));

        let output = write(&document);
        let serialized = String::from_utf8(output).unwrap();
        assert!(serialized.contains(control), "missing {control} in {serialized}");

        let reparsed = RtfDocument::parse(&serialized).unwrap();
        assert_eq!(reparsed.font_table().get(0).unwrap().theme, Some(*theme));
    }
}

#[test]
fn fonts_without_theme_selectors_have_none() {
    let source = r"{\rtf1\ansi{\fonttbl{\f0\froman Arial;}}\f0 Text\par}";
    let document = RtfDocument::parse(source).unwrap();
    assert_eq!(document.font_table().get(0).unwrap().theme, None);
}

#[test]
fn duplicate_theme_selectors_are_rejected() {
    let source = r"{\rtf1\ansi{\fonttbl{\f0\flomajor\fhimajor Arial;}}\f0 X\par}";
    assert!(RtfDocument::parse(source).is_err());
}
