#![allow(clippy::unwrap_used, reason = "test assertions use unwrap for clarity")]

use litchi_core::Error as CoreError;
use litchi_odg::flat::{Drawing, Error};

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
    let drawing = Drawing::from_bytes(SHAPE_FODG.as_bytes()).unwrap();
    assert_eq!(drawing.pages().len(), 1);
    assert_eq!(drawing.as_bytes(), SHAPE_FODG.as_bytes());
    assert_eq!(drawing.clone().into_bytes(), SHAPE_FODG.as_bytes());
}

#[test]
fn flat_shape_text_edit_is_lossless_and_transactional() {
    let source = Drawing::from_bytes(SHAPE_FODG.as_bytes()).unwrap();
    assert_eq!(source.pages()[0].shapes()[0].text(), "Old label");
    let mut edit = source.edit();
    edit.set_shape_text(0, 0, "New <label>").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(source.pages()[0].shapes()[0].text(), "Old label");
    assert_eq!(
        commit.package().pages()[0].shapes()[0].text(),
        "New <label>"
    );
    let output = std::str::from_utf8(commit.package().as_bytes()).unwrap();
    assert!(output.contains("<text:p>New &lt;label&gt;</text:p>"));
    assert_eq!(commit.patch().changes()[0].before(), "Old label");
    assert!(commit.patch().is_applicable_to(&source));
    assert!(commit.diagnostics().changed());
    let reapplied = source.apply(commit.patch()).unwrap();
    assert!(reapplied.diagnostics().changed());
    assert_eq!(reapplied.package().as_bytes(), commit.package().as_bytes());

    let different = Drawing::from_bytes(
        &SHAPE_FODG
            .replace("Old label", "Different label")
            .into_bytes(),
    )
    .unwrap();
    assert!(!commit.patch().is_applicable_to(&different));
    assert!(matches!(
        different.apply(commit.patch()),
        Err(Error::PatchConflict { .. })
    ));

    let inverse = commit.patch().inverse();
    assert_eq!(inverse.changes()[0].after(), "Old label");
    let restored = commit.package().apply(&inverse).unwrap();
    assert_eq!(restored.package().as_bytes(), source.as_bytes());
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
            Drawing::from_bytes(case.as_bytes()),
            Err(Error::Core(CoreError::InvalidFormat(_)))
        ));
    }
}

#[test]
fn flat_truncation_and_mutation_sweeps_never_panic() {
    let bytes = HARDENING_SEED.as_bytes();
    for end in 0..bytes.len() {
        drop(Drawing::from_bytes(&bytes[..end]));
    }
    for position in 0..bytes.len() {
        let mut mutated = bytes.to_vec();
        mutated[position] ^= 1;
        drop(Drawing::from_bytes(&mutated));
    }
    assert!(Drawing::from_bytes(bytes).is_ok());
    let mut invalid_utf8 = bytes.to_vec();
    invalid_utf8.insert(bytes.len() / 2, 0xfe);
    assert!(matches!(
        Drawing::from_bytes(&invalid_utf8),
        Err(Error::Core(CoreError::InvalidFormat(_)))
    ));
}

#[test]
fn flat_drawing_takes_caller_limits_through_open_options() {
    use litchi_odg::{LimitKind, OpenOptions, SemanticLimitKind, SemanticLimits};

    let shallow = OpenOptions::new().semantic_limits(
        SemanticLimits::builder()
            .max_depth(3)
            .unwrap()
            .build()
            .unwrap(),
    );
    assert!(matches!(
        Drawing::from_bytes_with_options(SHAPE_FODG.as_bytes(), shallow),
        Err(Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::Depth),
            ..
        })
    ));
    let one_shape = OpenOptions::new().semantic_limits(
        SemanticLimits::builder()
            .max_shapes(1)
            .unwrap()
            .max_text_bytes(1024)
            .unwrap()
            .build()
            .unwrap(),
    );
    let drawing = Drawing::from_bytes_with_options(SHAPE_FODG.as_bytes(), one_shape).unwrap();
    assert_eq!(drawing.options(), one_shape);
    let mut edit = drawing.edit();
    assert!(matches!(
        edit.set_shape_text(0, 0, "x".repeat(1025)),
        Err(Error::Limit {
            kind: LimitKind::Semantic(SemanticLimitKind::TextBytes),
            ..
        })
    ));
    assert!(matches!(
        edit.set_shape_text(0, 9, "x"),
        Err(Error::Refused(_))
    ));
    edit.set_shape_text(0, 0, "short").unwrap();
    edit.reset();
    let noop = edit.commit().unwrap();
    assert!(!noop.diagnostics().changed());
    assert!(noop.patch().is_noop());
    assert_eq!(noop.package().as_bytes(), drawing.as_bytes());

    let reader = Drawing::from_reader(std::io::Cursor::new(SHAPE_FODG.as_bytes())).unwrap();
    assert_eq!(reader.as_bytes(), SHAPE_FODG.as_bytes());
    let tiny = OpenOptions::new().source_limits(
        litchi_odg::SourceLimits::builder()
            .max_input_bytes(8)
            .unwrap()
            .build()
            .unwrap(),
    );
    assert!(matches!(
        Drawing::from_reader_with_options(std::io::Cursor::new(SHAPE_FODG.as_bytes()), tiny),
        Err(Error::Limit { .. })
    ));
}
