use litchi_rtf::{DocumentGenerator, RtfDocument, RtfWriter};
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
fn parses_unicode_generator_as_inert_provenance_and_round_trips() {
    let document =
        RtfDocument::parse(r#"{\rtf1\ansi{\*\generator Acme \u20320? 1.0;}Visible}"#).unwrap();
    assert_eq!(document.text(), "Visible");
    assert_eq!(document.generator().unwrap().value, "Acme 你 1.0");

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r#"{\*\generator Acme \u20320? 1.0;}"#));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.generator(), document.generator());
}

#[test]
fn mutation_validates_and_clear_preserves_body() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    document
        .set_generator(DocumentGenerator::new(Cow::Borrowed("Litchi 1.0")).unwrap())
        .unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.generator().unwrap().value, "Litchi 1.0");
    assert_eq!(reparsed.text(), "Body");

    document.clear_generator();
    assert!(document.generator().is_none());
    assert_eq!(document.text(), "Body");
}

#[test]
fn rejects_duplicate_or_active_generator_destinations() {
    let cases = [
        r#"{\rtf1{\generator Direct;}Body}"#,
        r#"{\rtf1{\*\generator One;}{\*\generator Two;}Body}"#,
        r#"{\rtf1{\*\generator ;}Body}"#,
        r#"{\rtf1{\*\generator A{nested};}Body}"#,
        r#"{\rtf1{\*\generator A\b B;}Body}"#,
        r#"{\rtf1{\*\generator\bin2 xx}Body}"#,
    ];
    for rtf in cases {
        assert!(RtfDocument::parse(rtf).is_err(), "accepted malformed {rtf}");
    }

    let oversized = format!(r#"{{\rtf1{{\*\generator {};}}}}"#, "x".repeat(65_537));
    assert!(RtfDocument::parse(&oversized).is_err());
    assert!(DocumentGenerator::new(Cow::Borrowed("\n")).is_err());
}

#[test]
fn parses_bundled_libreoffice_generator_fixtures() {
    const FIXTURES: &[(&str, &str)] = &[
        (
            "sw/qa/core/data/rtf/pass/fdo80924.rtf",
            "Microsoft Word 11.0.5604",
        ),
        (
            "sw/qa/extras/uiwriter/data/tdf86639.rtf",
            "Msftedit 5.41.21.2510",
        ),
        (
            "sw/qa/extras/rtfexport/data/fdo39001.rtf",
            "Apache XML Graphics RTF Library",
        ),
        (
            "sw/qa/extras/rtfexport/data/fdo76633.rtf",
            "LibreOfficeDev/4.4.0.0.alpha0$Linux_X86_64",
        ),
    ];
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../3rdparty/libreoffice-core/"
    );
    for (fixture, expected_prefix) in FIXTURES {
        let bytes = fs::read(format!("{root}{fixture}")).unwrap();
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        assert!(
            document
                .generator()
                .is_some_and(|generator| generator.value.starts_with(expected_prefix)),
            "unexpected generator in {fixture}"
        );
    }
}
