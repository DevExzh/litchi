#![allow(
    clippy::needless_pass_by_value,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "integration-test assertions panic on failure by design"
)]

use litchi_odp::{Package, core::OwnedPackage};

macro_rules! fixture {
    ($name:literal) => {
        include_bytes!(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../test-data/odf/odp/",
            $name
        ))
        .to_vec()
    };
}

fn deck_text(bytes: Vec<u8>) -> String {
    Package::from_bytes(&bytes)
        .unwrap()
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(litchi_odp::Slide::all_text)
        .collect::<Vec<_>>()
        .join("\n")
}

fn assert_exact_noop(bytes: Vec<u8>) {
    let source = Package::from_bytes(&bytes).unwrap();
    let commit = source.edit().unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().as_bytes(), bytes);
}

#[test]
fn real_fixture_text_and_unknown_markup_survive_exact_noop_commits() {
    for (name, bytes) in [
        ("cellspan.odp", fixture!("cellspan.odp")),
        ("text-in-image.odp", fixture!("text-in-image.odp")),
        ("tdf102223.odp", fixture!("tdf102223.odp")),
        ("tdf105502.odp", fixture!("tdf105502.odp")),
        ("tdf169979.odp", fixture!("tdf169979.odp")),
    ] {
        let before = deck_text(bytes.clone());
        assert!(
            !before.trim().is_empty() || name == "tdf169979.odp",
            "{name}"
        );
        assert_exact_noop(bytes.clone());
        assert_eq!(deck_text(bytes), before, "{name}");
    }
}

#[test]
fn table_spans_fonts_and_automatic_styles_are_retained_byte_exactly() {
    let table = fixture!("cellspan.odp");
    let package = OwnedPackage::from_bytes(table.clone()).unwrap();
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains("<table:table"));
    assert!(content.contains("table:number-columns-spanned=\"2\""));
    assert_exact_noop(table);

    let styled = fixture!("tdf102223.odp");
    let package = OwnedPackage::from_bytes(styled.clone()).unwrap();
    let content = String::from_utf8(package.get_file("content.xml").unwrap()).unwrap();
    assert!(content.contains("<office:font-face-decls>"));
    assert!(content.contains("style:family=\"drawing-page\""));
    assert_exact_noop(styled);
}

#[test]
fn empty_style_name_reference_opens_and_is_an_exact_noop() {
    let bytes = fixture!("tdf169979.odp");
    Package::from_bytes(&bytes)
        .unwrap()
        .presentation()
        .unwrap()
        .slides()
        .unwrap();
    assert_exact_noop(bytes);
}
