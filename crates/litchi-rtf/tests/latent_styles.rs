use litchi_rtf::{LatentStyleException, LatentStyles, RtfDocument, RtfWriter};
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
fn parses_full_unicode_latent_styles_and_round_trips_order() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1{\*\latentstyles\lsdstimax156\lsdlockeddef0"#,
        r#"\lsdsemihiddendef1\lsdunhideuseddef1\lsdqformatdef0\lsdprioritydef99"#,
        r#"{\lsdlockedexcept \lsdqformat1\lsdpriority0 Normal;"#,
        r#"\lsdsemihidden1\lsdunhideused0 Heading \u20320?;}}Body}"#,
    ))
    .unwrap();
    assert_eq!(document.text(), "Body");
    let styles = document.latent_styles().unwrap();
    assert_eq!(styles.max_style_index, 156);
    assert_eq!(styles.locked_default, Some(false));
    assert_eq!(styles.semi_hidden_default, Some(true));
    assert_eq!(styles.priority_default, Some(99));
    assert_eq!(styles.exceptions.len(), 2);
    assert_eq!(styles.exceptions[0].name, "Normal");
    assert_eq!(styles.exceptions[0].priority, Some(0));
    assert_eq!(styles.exceptions[1].name, "Heading 你");
    assert_eq!(styles.exceptions[1].semi_hidden, Some(true));

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.latent_styles(), Some(styles));
}

#[test]
fn mutation_validates_and_clear_preserves_body() {
    let styles = LatentStyles {
        max_style_index: 20,
        locked_default: Some(false),
        semi_hidden_default: None,
        unhide_when_used_default: None,
        quick_format_default: Some(false),
        priority_default: Some(99),
        exceptions: vec![LatentStyleException {
            name: Cow::Borrowed("Title"),
            locked: Some(false),
            semi_hidden: Some(false),
            unhide_when_used: Some(true),
            quick_format: Some(true),
            priority: Some(10),
        }],
    };
    let mut document = RtfDocument::parse(r#"{\rtf1 Text}"#).unwrap();
    document.set_latent_styles(styles.clone()).unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.latent_styles(), Some(&styles));
    assert_eq!(reparsed.text(), "Text");

    document.clear_latent_styles();
    assert!(document.latent_styles().is_none());
    assert_eq!(document.text(), "Text");
}

#[test]
fn rejects_malformed_or_active_latent_style_grammar() {
    let cases = [
        r#"{\rtf1{\latentstyles\lsdstimax1}}"#,
        r#"{\rtf1{\*\latentstyles\lsdlockeddef0}}"#,
        r#"{\rtf1{\*\latentstyles\lsdstimax1\lsdstimax1}}"#,
        r#"{\rtf1{\*\latentstyles\lsdstimax1\lsdlockeddef2}}"#,
        r#"{\rtf1{\*\latentstyles\lsdstimax1\lsdprioritydef100}}"#,
        r#"{\rtf1{\*\latentstyles\lsdstimax1{\lsdlockedexcept Name}}}"#,
        r#"{\rtf1{\*\latentstyles\lsdstimax1{\lsdlockedexcept ;}}}"#,
        r#"{\rtf1{\*\latentstyles\lsdstimax1{\lsdlockedexcept Normal;Normal;}}}"#,
        r#"{\rtf1{\*\latentstyles\lsdstimax1{\lsdlockedexcept \lsdqformat1\lsdqformat0 Name;}}}"#,
        r#"{\rtf1{\*\latentstyles\lsdstimax1{\lsdlockedexcept {Name;}}}}"#,
        r#"{\rtf1{\*\latentstyles\lsdstimax1{\lsdlockedexcept\bin2 xx}}}"#,
        r#"{\rtf1\lsdstimax1}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }
}

#[test]
fn parses_bundled_libreoffice_latent_style_fixtures() {
    const FIXTURES: &[(&str, bool)] = &[
        (
            "sw/qa/writerfilter/filters-test/data/pass/TCI-TN65GP-DDRHDLL-partial.rtf",
            false,
        ),
        ("sw/qa/core/data/rtf/pass/tdf116851.rtf", true),
        ("sw/qa/extras/ooxmlexport/data/tdf154703_framePr2.rtf", true),
        ("sw/qa/extras/rtfexport/data/fdo55504-1-min.rtf", false),
    ];
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core"
    );
    for (fixture, has_exceptions) in FIXTURES {
        let bytes = fs::read(format!("{root}{fixture}")).unwrap();
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        let styles = document
            .latent_styles()
            .unwrap_or_else(|| panic!("fixture exposed no latent styles: {fixture}"));
        assert!(styles.max_style_index > 0);
        assert_eq!(!styles.exceptions.is_empty(), *has_exceptions);
    }
}
