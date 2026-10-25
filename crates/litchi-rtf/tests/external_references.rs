#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::raw::Tree;
use litchi_rtf::{DocumentExternalReferences, write::Writer};
use std::borrow::Cow;

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn parses_both_opaque_reference_names_and_round_trips_deterministically() {
    let document = Tree::parse(
        concat!(
            r#"{\rtf1\ansi{\*\template C:\\Templates\\\u20320?.dot}"#,
            r#"{\*\nextfile queue\\part2.rtf}Body}"#,
        ),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "Body");
    assert_eq!(
        document.external_references().template.as_deref(),
        Some("C:\\Templates\\你.dot")
    );
    assert_eq!(
        document.external_references().next_file.as_deref(),
        Some("queue\\part2.rtf")
    );

    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.find("\\nextfile").unwrap() < serialized.find("\\template").unwrap());
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(
        reparsed.external_references(),
        document.external_references()
    );
    assert_eq!(reparsed.text(), document.text());
}

#[test]
fn typed_api_validates_and_clears_without_resolving_names() {
    let mut document =
        Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
    document
        .set_external_references(DocumentExternalReferences {
            next_file: Some(Cow::Borrowed("file:///never-opened.rtf")),
            template: Some(Cow::Borrowed("https://never-resolved.invalid/template.dot")),
        })
        .unwrap();
    let reparsed =
        Tree::parse_bytes(&write(&document), litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(
        reparsed.external_references(),
        document.external_references()
    );
    assert_eq!(reparsed.text(), "Body");

    document.clear_external_references();
    assert!(document.external_references().is_empty());
}

#[test]
fn rejects_bad_placement_cardinality_active_content_and_resource_exhaustion() {
    for source in [
        r"{\rtf1{\template direct.dot}Body}",
        r"{\rtf1{\nextfile direct.rtf}Body}",
        r"{\rtf1{\*\template one.dot}{\*\template two.dot}Body}",
        r"{\rtf1{\*\nextfile one.rtf}{\*\nextfile two.rtf}Body}",
        r"{\rtf1{\*\template }Body}",
        r"{\rtf1 Body{\*\template late.dot}}",
        r"{\rtf1{{\*\template nested.dot}}Body}",
        r"{\rtf1{\*\template x{nested}.dot}Body}",
        r"{\rtf1{\*\template x\b y.dot}Body}",
        r"{\rtf1{\*\nextfile\bin2 xx}Body}",
        r"{\rtf1\template Body}",
    ] {
        assert!(
            Tree::parse(source, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted malformed {source}"
        );
    }

    let oversized = format!(r"{{\rtf1{{\*\template {}}}Body}}", "x".repeat(65_537));
    assert!(Tree::parse(&oversized, litchi_rtf::read::OpenOptions::default()).is_err());
}

#[test]
fn parses_bundled_libreoffice_template_fixture() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/libreoffice-core/sw/qa/extras/rtfimport/data/tdf163003.rtf"
    ))
    .unwrap();
    let document = Tree::parse_bytes(&bytes, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(
        document.external_references().template.as_deref(),
        Some(r"C:\Users\xk1c\AppData\Roaming\Microsoft\Templates\kis.3.0.dot")
    );
    assert!(document.external_references().next_file.is_none());
}
