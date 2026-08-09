#![allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]

use litchi_core::Position;
use litchi_odf_common::core::PackageWriter;
use litchi_oth::{Block, Builder, Template, heading::Heading, link::Link, paragraph::Paragraph};
use std::sync::Arc;

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
fn typed_builder_writes_byte_minimal_escaped_blocks() {
    let heading = Heading::styled(2, "A & B", "Heading").unwrap();
    let bytes = Builder::new()
        .heading(heading)
        .paragraph(Paragraph::styled("<body>", "Body"))
        .build()
        .unwrap();
    let template = Template::from_bytes(bytes).unwrap();
    assert!(template.content_xml().contains(
        "<text:h text:outline-level=\"2\" text:style-name=\"Heading\">A &amp; B</text:h>"
    ));
    assert!(
        template
            .content_xml()
            .contains("<text:p text:style-name=\"Body\">&lt;body&gt;</text:p>")
    );
    assert!(!template.content_xml().contains('\n'));
    assert!(!template.content_xml().contains(" />"));
    assert!(
        Builder::new()
            .content_xml(COMPACT_CONTENT)
            .paragraph(Paragraph::new("ambiguous"))
            .build()
            .is_err()
    );
}

#[test]
fn compact_authored_content_is_published_without_rewriting() {
    let template =
        Template::from_bytes(Builder::new().content_xml(COMPACT_CONTENT).build().unwrap()).unwrap();
    assert_eq!(template.content_xml(), COMPACT_CONTENT);
    assert!(!template.content_xml().contains("\n<"));
    assert!(!template.content_xml().contains(" />"));
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
fn builder_authors_compact_optional_xml_parts_and_reopens_before_publication() {
    const META: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:meta/></office:document-meta>"#;
    const SETTINGS: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-settings xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:settings/></office:document-settings>"#;
    const STYLES: &str = r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:styles/></office:document-styles>"#;
    let bytes = Builder::new()
        .meta_xml(META)
        .settings_xml(SETTINGS)
        .styles_xml(STYLES)
        .build()
        .unwrap();
    let template = Template::from_bytes(bytes).unwrap();
    assert_eq!(template.styles_xml(), Some(STYLES));
    assert!(
        template
            .files()
            .unwrap()
            .iter()
            .any(|path| path == "settings.xml")
    );
    assert!(
        Builder::new()
            .styles_xml("<office:document-styles />")
            .build()
            .is_err()
    );
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
fn shared_archive_input_opens_without_changing_the_source_bytes() {
    let bytes = Arc::new(Builder::new().build().unwrap());
    let template = Template::from_shared_bytes(Arc::clone(&bytes)).unwrap();
    assert_eq!(template.as_bytes(), bytes.as_slice());
    assert_eq!(Arc::strong_count(&bytes), 2);
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

#[test]
fn projects_headings_links_styles_and_odf_whitespace_as_inert_values() {
    const CONTENT: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4">"#,
        r#"<office:body><office:text><text:h text:outline-level="2">Welcome</text:h>"#,
        r#"<text:p text:style-name="Body">one<text:s text:c="2"/><text:a xlink:href="https://example.test/value">link</text:a><text:tab/><text:line-break/>two</text:p>"#,
        r#"</office:text></office:body></office:document-content>"#,
    );
    let template =
        Template::from_bytes(Builder::new().content_xml(CONTENT).build().unwrap()).unwrap();
    let body = template.text_body().unwrap();
    assert_eq!(body.headings()[0].level(), 2);
    assert_eq!(body.paragraphs()[0].text(), "one  link\t\ntwo");
    assert_eq!(body.paragraphs()[0].style_name(), Some("Body"));
    assert_eq!(body.paragraphs()[0].links()[0].label(), "link");
    let blocks = body.blocks().collect::<Vec<_>>();
    assert!(matches!(blocks[0], Block::Heading(_)));
    assert!(matches!(blocks[1], Block::Paragraph(_)));
}

#[test]
fn multi_paragraph_edit_is_atomic_and_can_fill_an_empty_element() {
    const CONTENT: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.4">"#,
        r#"<office:body><office:text><text:p/><text:p>second</text:p><text:p>third</text:p></office:text></office:body></office:document-content>"#,
    );
    let source =
        Template::from_bytes(Builder::new().content_xml(CONTENT).build().unwrap()).unwrap();
    let mut edit = source.edit();
    edit.set_paragraph_text(Position::new(0), "first & ready")
        .unwrap();
    edit.set_paragraph_text(Position::new(2), "last").unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(commit.patch().changes().len(), 2);
    assert!(
        commit
            .template()
            .content_xml()
            .contains("<text:p>first &amp; ready</text:p>")
    );
    let body = commit.template().text_body().unwrap();
    let texts = body
        .paragraphs()
        .iter()
        .map(Paragraph::text)
        .collect::<Vec<_>>();
    assert_eq!(texts, ["first & ready", "second", "last"]);
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
fn opens_and_edits_the_pretty_printed_libreoffice_web_template() {
    const CONTENT: &str = include_str!(
        "../../../3rdparty/libreoffice-core/extras/source/templates/wizard/desktop/html/content.xml"
    );
    const META: &str = include_str!(
        "../../../3rdparty/libreoffice-core/extras/source/templates/wizard/desktop/html/meta.xml"
    );
    const SETTINGS: &str = include_str!(
        "../../../3rdparty/libreoffice-core/extras/source/templates/wizard/desktop/html/settings.xml"
    );
    const STYLES: &str = include_str!(
        "../../../3rdparty/libreoffice-core/extras/source/templates/wizard/desktop/html/styles.xml"
    );

    let mut writer = PackageWriter::new();
    writer
        .set_mimetype("application/vnd.oasis.opendocument.text-web")
        .unwrap();
    writer.add_file("content.xml", CONTENT.as_bytes()).unwrap();
    writer.add_file("meta.xml", META.as_bytes()).unwrap();
    writer
        .add_file("settings.xml", SETTINGS.as_bytes())
        .unwrap();
    writer.add_file("styles.xml", STYLES.as_bytes()).unwrap();
    let bytes = writer.finish_to_bytes().unwrap();
    let source = Template::from_bytes(bytes.clone()).unwrap();
    assert_eq!(source.as_bytes(), bytes);
    assert_eq!(
        source.text_body().unwrap().paragraphs()[0].style_name(),
        Some("Standard")
    );

    let mut edit = source.edit();
    edit.set_paragraph_text(Position::new(0), "Web template")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit.template().text_body().unwrap().paragraphs()[0].text(),
        "Web template"
    );
    assert!(commit.template().styles_xml().is_some());
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
