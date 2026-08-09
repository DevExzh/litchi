#![allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]

use litchi_core::Position;
use litchi_oth::{Builder, Template, link::Link, paragraph::Paragraph};

const COMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">"#,
    r#"<office:body><office:text><text:p>compact</text:p></office:text></office:body>"#,
    r#"</office:document-content>"#,
);

const NONCOMPACT_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" office:version="1.3">"#,
    "\n<office:body><office:text/></office:body></office:document-content>",
);

const SEMANTIC_WHITESPACE_CONTENT: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
    r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.3">"#,
    "<office:body><office:text><text:p>line one\n  line two</text:p></office:text>",
    "</office:body></office:document-content>",
);

const PREFIX_ALIASED_TEXT_WEB: &str = include_str!("fixtures/odf14-text-web.xml");
const WRONG_FAMILY: &str = include_str!("fixtures/wrong-family.xml");
const DUPLICATE_BODY: &str = include_str!("fixtures/duplicate-body.xml");

#[test]
fn focused_modules_are_the_canonical_semantic_api() {
    assert_eq!(Paragraph::new("Welcome").text(), "Welcome");
    let link = Link::new("https://example.test", "Example");
    assert_eq!(link.href(), "https://example.test");
    assert_eq!(link.label(), "Example");

    let bytes = Builder::new().build().unwrap();
    let template = Template::from_bytes(bytes).unwrap();
    assert!(template.content_xml().contains("<office:text"));
}

#[test]
fn compact_authored_content_is_published_without_rewriting() {
    let template =
        Template::from_bytes(Builder::new().content_xml(COMPACT_CONTENT).build().unwrap()).unwrap();
    assert_eq!(template.content_xml(), COMPACT_CONTENT);
}

#[test]
fn noncompact_authored_content_returns_a_typed_error() {
    assert!(matches!(
        Builder::new()
            .content_xml(NONCOMPACT_CONTENT)
            .build()
            .unwrap_err(),
        litchi_core::Error::XmlCompactness {
            kind: litchi_core::xml::CompactnessKind::FormattingWhitespace,
            ..
        }
    ));
}

#[test]
fn semantic_whitespace_is_preserved_exactly() {
    let template = Template::from_bytes(
        Builder::new()
            .content_xml(SEMANTIC_WHITESPACE_CONTENT)
            .build()
            .unwrap(),
    )
    .unwrap();
    assert_eq!(template.content_xml(), SEMANTIC_WHITESPACE_CONTENT);
}

#[test]
fn loaded_template_snapshot_remains_byte_exact() {
    let bytes = Builder::new().build().unwrap();
    let template = Template::from_bytes(bytes.clone()).unwrap();
    assert_eq!(template.as_bytes(), bytes.as_slice());
    assert_eq!(template.into_bytes(), bytes);
}

#[test]
fn local_odf14_fixture_uses_namespace_aware_text_web_contract() {
    let bytes = Builder::new()
        .content_xml(PREFIX_ALIASED_TEXT_WEB)
        .build()
        .unwrap();
    let template = Template::from_bytes(bytes).unwrap();
    let body = template.text_body().unwrap();
    assert_eq!(body.paragraphs().len(), 1);
    assert_eq!(body.paragraphs()[0].text(), "prefix-safe text");
}

#[test]
fn local_invalid_family_fixtures_fail_before_package_publication() {
    for xml in [WRONG_FAMILY, DUPLICATE_BODY] {
        assert!(Builder::new().content_xml(xml).build().is_err());
    }
}

#[test]
fn exact_noop_edit_is_source_checked_and_atomic() {
    let source = Template::from_bytes(Builder::new().build().unwrap()).unwrap();
    let alternate =
        Template::from_bytes(Builder::new().content_xml(COMPACT_CONTENT).build().unwrap()).unwrap();
    let commit = source.edit().commit().unwrap();
    assert!(!commit.changed());
    assert_eq!(commit.template().as_bytes(), source.as_bytes());
    assert_eq!(
        commit.patch().apply(&source).unwrap().as_bytes(),
        source.as_bytes()
    );
    assert!(commit.patch().apply(&alternate).is_err());
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.template())
            .unwrap()
            .as_bytes(),
        source.as_bytes(),
    );
}

#[test]
fn paragraph_edit_is_bounded_reversible_and_preserves_unknown_content() {
    const CONTENT: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4">"#,
        r#"<office:body><office:text><text:p>before</text:p><foreign:keep xmlns:foreign="urn:example">opaque</foreign:keep></office:text></office:body></office:document-content>"#,
    );
    let source =
        Template::from_bytes(Builder::new().content_xml(CONTENT).build().unwrap()).unwrap();
    let mut edit = source.edit();
    edit.set_paragraph_text(Position::new(0), "after & exact")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.changed());
    assert_eq!(
        commit.template().text_body().unwrap().paragraphs()[0].text(),
        "after & exact"
    );
    assert!(
        commit
            .template()
            .content_xml()
            .contains("<foreign:keep xmlns:foreign=\"urn:example\">opaque</foreign:keep>")
    );
    assert_eq!(
        commit.patch().change().unwrap().paragraph(),
        Position::new(0)
    );
    assert_eq!(commit.patch().change().unwrap().before(), "before");
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.template())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn paragraph_edit_refuses_nested_markup_that_cannot_be_rewritten_losslessly() {
    const CONTENT: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4">"#,
        r#"<office:body><office:text><text:p>plain <text:span>nested</text:span></text:p></office:text></office:body></office:document-content>"#,
    );
    let source =
        Template::from_bytes(Builder::new().content_xml(CONTENT).build().unwrap()).unwrap();
    let mut edit = source.edit();
    assert!(
        edit.set_paragraph_text(Position::new(0), "replacement")
            .is_err()
    );
}
