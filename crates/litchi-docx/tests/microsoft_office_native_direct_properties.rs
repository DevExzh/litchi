//! Direct properties on native Word runs and paragraphs; style inheritance stays distinct.
use litchi_docx::{
    Package,
    properties::{HalfPoints, Justification},
};
use std::path::Path;

fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(file),
    )
    .unwrap()
}

#[test]
fn native_word_direct_properties_keep_absence_distinct_from_style_inheritance() {
    let bytes = native("word-text-structure.docx");
    let p = Package::from_bytes(&bytes).unwrap();
    let doc = p.document().unwrap();
    let paragraphs = doc.paragraphs().unwrap();
    let title = paragraphs
        .iter()
        .find(|p| p.text().unwrap() == "Native Word text and structure")
        .unwrap();
    let props = title.direct_properties().unwrap();
    assert_eq!(props.style.as_deref(), Some("Title"));
    assert!(props.justification.is_none() && props.spacing.is_none() && props.mark.is_none());
    assert_eq!(props.invalid_properties().count(), 0);
    let title_runs = title.runs().unwrap();
    let local = title_runs[0].direct_properties().unwrap();
    assert!(local.bold.is_none() && local.size.is_none() && local.fonts.is_none());
    let centered = paragraphs
        .iter()
        .find(|p| p.text().unwrap() == "Paragraph beta: centered.")
        .unwrap();
    assert_eq!(
        centered.direct_properties().unwrap().justification,
        Some(Justification::Center)
    );
    let left = paragraphs
        .iter()
        .find(|p| p.text().unwrap() == "Paragraph alpha: left aligned.")
        .unwrap();
    assert!(left.direct_properties().unwrap().justification.is_none());
    let sample = paragraphs
        .iter()
        .find(|p| p.text().unwrap().starts_with("Formatting sample:"))
        .unwrap();
    let runs = sample.runs().unwrap();
    for (text, bold, italic, strike, underline) in [
        ("bold", Some(true), None, None, false),
        ("italic", None, Some(true), None, false),
        ("underline", None, None, None, true),
        ("strikethrough", None, None, Some(true), false),
    ] {
        let run = runs.iter().find(|r| r.text().unwrap() == text).unwrap();
        let local = run.direct_properties().unwrap();
        assert_eq!(local.bold, bold);
        assert_eq!(local.italic, italic);
        assert_eq!(local.strike, strike);
        assert_eq!(local.underline.is_some(), underline);
        if underline {
            assert_eq!(
                local.underline.unwrap().style,
                litchi_docx::format::UnderlineStyle::Single
            );
        }
        assert_eq!(local.invalid_properties().count(), 0);
    }
    let unicode = paragraphs
        .iter()
        .find(|p| p.text().unwrap().starts_with("Unicode:"))
        .unwrap();
    let runs = unicode.runs().unwrap();
    for (text, font, east_asia) in [
        ("中文", "MS Gothic", Some("MS Gothic")),
        ("日本語", "MS Gothic", Some("MS Gothic")),
        ("العربية", "Arial", None),
        ("😀", "Apple Color Emoji", None),
    ] {
        let local = runs
            .iter()
            .find(|r| r.text().unwrap() == text)
            .unwrap()
            .direct_properties()
            .unwrap();
        let fonts = local.fonts.unwrap();
        assert_eq!(fonts.ascii.as_deref(), Some(font));
        assert_eq!(fonts.h_ansi.as_deref(), Some(font));
        assert_eq!(fonts.complex_script.as_deref(), Some(font));
        assert_eq!(fonts.east_asia.as_deref(), east_asia);
        assert!(fonts.ascii_theme.is_none() && fonts.h_ansi_theme.is_none());
    }
    assert_eq!(p.to_bytes().unwrap(), bytes);
}

#[test]
fn native_macro_word_direct_run_and_paragraph_mark_properties_retain_explicit_zeroes() {
    for file in ["word-recorded-macro.docm", "word-recorded-macro.dotm"] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let doc = p.document().unwrap();
        let paragraphs = doc.paragraphs().unwrap();
        let paragraph = paragraphs
            .iter()
            .find(|p| p.text().unwrap() == "MACRO_FORMAT_TARGET")
            .unwrap();
        let props = paragraph.direct_properties().unwrap();
        let mark = props.mark.as_ref().unwrap();
        let runs = paragraph.runs().unwrap();
        let direct = runs[0].direct_properties().unwrap();
        for local in [&direct, mark.as_ref()] {
            assert_eq!(local.bold, Some(true));
            assert_eq!(local.bold_complex_script, Some(true));
            assert_eq!(local.kerning, Some(HalfPoints(0)));
            assert!(local.size.is_none());
            assert_eq!(
                local.fonts.as_ref().unwrap().ascii.as_deref(),
                Some("Times New Roman")
            );
            assert_eq!(local.invalid_properties().count(), 0);
        }
        assert!(props.style.is_none());
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}
