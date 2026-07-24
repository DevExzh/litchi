use litchi_rtf::{DocumentLanguageDefaults, LanguageId, RtfDocument, RtfWriter};
use std::fs;

fn language(value: u32) -> LanguageId {
    LanguageId::new(value).unwrap()
}

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_header_defaults_and_language_run_boundaries() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\deflang1033\deflangfe2052\adeflang1025 "#,
        r#"English {\lang1049\langfe1049 Russian} English again}"#,
    ))
    .unwrap();
    assert_eq!(document.language_defaults().primary, Some(language(1033)));
    assert_eq!(
        document.language_defaults().east_asian,
        Some(language(2052))
    );
    assert_eq!(
        document.language_defaults().complex_script,
        Some(language(1025))
    );
    let russian = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("Russian"))
        .unwrap();
    assert_eq!(russian.formatting.language, Some(language(1049)));
    assert_eq!(russian.formatting.east_asian_language, Some(language(1049)));

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.language_defaults(), document.language_defaults());
    assert_eq!(reparsed.text(), document.text());
    assert!(
        reparsed
            .blocks()
            .iter()
            .any(|block| block.formatting.language == Some(language(1049)))
    );
}

#[test]
fn plain_restores_document_languages_and_clears_no_proof() {
    let document = RtfDocument::parse(concat!(
        r#"{\rtf1\deflang1033\deflangfe2052 "#,
        r#"{\lang1024\langfe1024\langnp1060\langfenp1060\noproof unchecked}"#,
        r#"\plain checked}"#,
    ))
    .unwrap();
    let unchecked = document
        .blocks()
        .iter()
        .find(|block| block.text.contains("unchecked"))
        .unwrap();
    assert!(unchecked.formatting.no_proof);
    assert_eq!(unchecked.formatting.language, Some(LanguageId::UNDEFINED));
    assert_eq!(unchecked.formatting.language_no_proof, Some(language(1060)));

    let checked = document
        .blocks()
        .iter()
        .find(|block| block.text.trim() == "checked")
        .unwrap();
    assert!(!checked.formatting.no_proof);
    assert_eq!(checked.formatting.language, Some(language(1033)));
    assert_eq!(checked.formatting.east_asian_language, Some(language(2052)));
    assert_eq!(checked.formatting.language_no_proof, Some(language(1033)));
    assert_eq!(
        checked.formatting.east_asian_language_no_proof,
        Some(language(2052))
    );
}

#[test]
fn mutation_writer_and_language_id_validation() {
    let mut document = RtfDocument::parse(r#"{\rtf1 Body}"#).unwrap();
    let defaults = DocumentLanguageDefaults {
        primary: Some(language(1031)),
        east_asian: Some(language(1041)),
        complex_script: Some(language(1025)),
    };
    document.set_language_defaults(defaults).unwrap();
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains(r#"\deflang1031"#));
    assert!(serialized.contains(r#"\deflangfe1041"#));
    assert!(serialized.contains(r#"\adeflang1025"#));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(*reparsed.language_defaults(), defaults);
    assert_eq!(reparsed.text(), "Body");

    document.clear_language_defaults();
    assert_eq!(
        document.language_defaults(),
        &DocumentLanguageDefaults::default()
    );
    assert!(LanguageId::new(65_535).is_ok());
    assert!(LanguageId::new(65_536).is_err());
    assert!(RtfDocument::parse(r#"{\rtf1\lang-1 bad}"#).is_err());
    assert!(RtfDocument::parse(r#"{\rtf1\lang65536 bad}"#).is_err());
}

#[test]
fn parses_bundled_libreoffice_language_and_proofing_fixtures() {
    const FIXTURES: &[&str] = &[
        "sw/qa/core/data/rtf/pass/tdf116851.rtf",
        "sw/qa/extras/rtfexport/data/cjklist25.rtf",
        "sw/qa/extras/ooxmlexport/data/ooo39250-1-min.rtf",
    ];
    let root = concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core"
    );
    for fixture in FIXTURES {
        let bytes = fs::read(format!("{root}/{fixture}")).unwrap();
        let document = RtfDocument::parse_bytes(&bytes)
            .unwrap_or_else(|error| panic!("failed to parse {fixture}: {error}"));
        assert!(
            document.blocks().iter().any(|block| {
                block.formatting.language.is_some()
                    || block.formatting.east_asian_language.is_some()
                    || block.formatting.no_proof
            }),
            "fixture exposed no language/proofing metadata: {fixture}"
        );
    }
}
