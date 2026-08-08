#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::Error;
use litchi_odg::FlatDrawing;

const SHAPE_FODG: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?>"#,
    r#"<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"office:mimetype="application/vnd.oasis.opendocument.graphics"><office:body>"#,
    r#"<office:drawing><draw:page draw:name="p1"><draw:rect draw:name="box">"#,
    r#"<text:p>Old label</text:p></draw:rect></draw:page></office:drawing>"#,
    r#"</office:body></office:document>"#,
);
const HARDENING_SEED: &str = concat!(
    r#"<?xml version="1.0"?><office:document "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
    r#"office:mimetype="application/vnd.oasis.opendocument.graphics"><office:body>"#,
    r#"<office:drawing><draw:page draw:name="p"><draw:g><draw:custom-shape>"#,
    r#"<text:p>rounded</text:p></draw:custom-shape></draw:g></draw:page>"#,
    r#"</office:drawing></office:body></office:document>"#,
);

#[test]
fn flat_pages_and_noop_bytes_stay_exact() {
    let drawing = FlatDrawing::from_bytes(SHAPE_FODG.as_bytes().to_vec()).unwrap();
    assert_eq!(drawing.pages().len(), 1);
    assert_eq!(drawing.as_bytes(), SHAPE_FODG.as_bytes());
    assert_eq!(drawing.clone().into_bytes(), SHAPE_FODG.as_bytes());
}

#[test]
fn flat_shape_text_edit_is_lossless_and_transactional() {
    let source = FlatDrawing::from_bytes(SHAPE_FODG.as_bytes().to_vec()).unwrap();
    assert_eq!(source.pages()[0].shapes()[0].text(), "Old label");
    let mut edit = source.edit();
    edit.set_shape_text(0, 0, "New <label>").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(source.pages()[0].shapes()[0].text(), "Old label");
    assert_eq!(
        commit.snapshot().pages()[0].shapes()[0].text(),
        "New <label>"
    );
    let output = std::str::from_utf8(commit.snapshot().as_bytes()).unwrap();
    assert!(output.contains("<text:p>New &lt;label&gt;</text:p>"));
    assert_eq!(commit.patch().changes()[0].before(), "Old label");
    assert!(commit.patch().is_applicable_to(&source));
    let reapplied = commit.patch().apply(&source).unwrap();
    assert_eq!(reapplied.as_bytes(), commit.snapshot().as_bytes());

    let different = FlatDrawing::from_bytes(
        SHAPE_FODG
            .replace("Old label", "Different label")
            .into_bytes(),
    )
    .unwrap();
    assert!(!commit.patch().is_applicable_to(&different));
    assert!(matches!(
        commit.patch().apply(&different),
        Err(Error::InvalidFormat(_))
    ));

    let inverse = commit.patch().inverse();
    assert_eq!(inverse.changes()[0].after(), "Old label");
    let restored = inverse.apply(commit.snapshot()).unwrap();
    assert_eq!(restored.as_bytes(), source.as_bytes());
}

#[test]
fn flat_wrong_family_and_malformed_or_misplaced_structures_are_typed_errors() {
    let wrong = SHAPE_FODG.replace(
        "application/vnd.oasis.opendocument.graphics",
        "application/vnd.oasis.opendocument.text",
    );
    let cases = [
        wrong,
        SHAPE_FODG.replace("<office:drawing>", ""),
        SHAPE_FODG.replace("</office:drawing>", ""),
        SHAPE_FODG.replace(
            "<office:drawing><draw:page draw:name=\"p1\">",
            "<draw:page draw:name=\"p1\"><office:drawing>",
        ),
        SHAPE_FODG.replace(
            "</draw:page></office:drawing>",
            "</office:drawing></draw:page>",
        ),
    ];
    for case in cases {
        assert!(matches!(
            FlatDrawing::from_bytes(case.into_bytes()),
            Err(Error::InvalidFormat(_))
        ));
    }
}

#[test]
fn flat_truncation_and_mutation_sweeps_never_panic() {
    let bytes = HARDENING_SEED.as_bytes();
    for end in 0..bytes.len() {
        let _ = FlatDrawing::from_bytes(bytes[..end].to_vec());
    }
    for position in 0..bytes.len() {
        let mut mutated = bytes.to_vec();
        mutated[position] ^= 1;
        let _ = FlatDrawing::from_bytes(mutated);
    }
    assert!(FlatDrawing::from_bytes(bytes.to_vec()).is_ok());
    let mut invalid_utf8 = bytes.to_vec();
    invalid_utf8.insert(bytes.len() / 2, 0xfe);
    assert!(matches!(
        FlatDrawing::from_bytes(invalid_utf8),
        Err(Error::InvalidFormat(_))
    ));
}
