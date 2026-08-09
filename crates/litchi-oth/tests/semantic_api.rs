#![allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]

use litchi_core::{HistoryLimits, Metadata, Position};
use litchi_oth::{
    Block, Builder, History, JoinFailure, Patch, SecurityPolicy, Template, TransferPolicy,
    TransferSelector, field,
    form::{Control, Form},
    heading::Heading,
    inline::{Content as Inline, Field as InlineField, Span},
    link::Link,
    list::{Item, List},
    paragraph::Paragraph,
    resource,
    style::{Slant, Style, TextProperties, Weight},
};
use soapberry_zip::office::StreamingArchiveWriter;
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

fn raw_negative_package(content: &str) -> Vec<u8> {
    const MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.text-web";
    const MANIFEST: &[u8] = br#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text-web"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>"#;
    let mut archive = StreamingArchiveWriter::new();
    archive.write_stored("mimetype", MIMETYPE).unwrap();
    archive
        .write_deflated("content.xml", content.as_bytes())
        .unwrap();
    archive
        .write_deflated("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    archive.finish_to_bytes().unwrap()
}

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
fn raw_zip_negatives_are_rejected_by_the_open_boundary() {
    for xml in [WRONG_FAMILY, DUPLICATE_BODY] {
        assert!(Template::from_bytes(raw_negative_package(xml)).is_err());
    }
    let dtd = COMPACT_CONTENT.replacen(
        "<office:document-content",
        "<!DOCTYPE office:document-content><office:document-content",
        1,
    );
    assert!(Template::from_bytes(raw_negative_package(&dtd)).is_err());
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
    let mut absent_removal_edit = source.edit();
    absent_removal_edit.remove_metadata();
    absent_removal_edit.remove_styles();
    let absent_removal_commit = absent_removal_edit.commit().unwrap();
    assert!(!absent_removal_commit.changed());
    assert_eq!(
        absent_removal_commit.template().as_bytes(),
        source.as_bytes()
    );
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
fn projects_lists_bookmarks_fields_formatting_resources_forms_and_styles() {
    const CONTENT: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" "#,
        r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
        r#"xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" "#,
        r#"xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.4">"#,
        r#"<office:automatic-styles><style:style style:name="Strong" style:family="text"/></office:automatic-styles>"#,
        r#"<office:body><office:text><text:p>pre<text:bookmark-start text:name="chapter"/><text:span text:style-name="Strong">bold <text:date office:date-value="2026-08-10" text:fixed="true">today</text:date></text:span></text:p>"#,
        r#"<text:list text:style-name="Numbered"><text:list-item text:start-value="3"><text:p>three<text:bookmark-end text:name="chapter"/></text:p></text:list-item><text:list-item><text:p>four</text:p></text:list-item></text:list>"#,
        r#"<text:p><draw:frame><draw:image xlink:href="Pictures/pic.png"/><draw:object xlink:href="./Object 1"/></draw:frame></text:p>"#,
        r#"<office:forms><form:form form:name="Search"><form:text form:id="query" form:name="q"/></form:form></office:forms>"#,
        r#"</office:text></office:body></office:document-content>"#,
    );
    const STYLES: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-styles "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"><office:styles>"#,
        r#"<style:style style:name="Body" style:family="paragraph" style:parent-style-name="Standard"/>"#,
        r#"</office:styles></office:document-styles>"#,
    );
    let template = Template::from_bytes(
        Builder::new()
            .content_xml(CONTENT)
            .styles_xml(STYLES)
            .build()
            .unwrap(),
    )
    .unwrap();
    let body = template.text_body().unwrap();
    assert_eq!(body.lists().len(), 1);
    assert_eq!(body.lists()[0].style_name(), Some("Numbered"));
    assert_eq!(body.lists()[0].items()[0].start_value(), Some(3));
    assert_eq!(body.lists()[0].items()[0].paragraphs()[0].text(), "three");
    assert_eq!(body.bookmarks()[0].name(), "chapter");
    assert_eq!(body.bookmarks()[0].start().block(), Position::new(0));
    assert_eq!(body.bookmarks()[0].end().unwrap().block(), Position::new(1));
    assert_eq!(
        body.paragraphs()[0].formatting_runs()[0].style_name(),
        "Strong"
    );
    let projected_field = &body.paragraphs()[0].fields()[0];
    assert!(matches!(projected_field.kind(), field::Kind::Date));
    assert!(projected_field.is_fixed());
    assert_eq!(projected_field.value(), Some("2026-08-10"));
    assert_eq!(body.resources().len(), 2);
    assert_eq!(body.resources()[0].kind(), resource::Kind::Image);
    assert!(body.resources()[0].is_embedded());
    assert_eq!(body.forms()[0].name(), Some("Search"));
    assert_eq!(body.forms()[0].controls()[0].kind(), "text");
    assert_eq!(body.forms()[0].controls()[0].id(), Some("query"));
    assert_eq!(template.styles().len(), 2);
    assert_eq!(template.styles()[1].parent_name(), Some("Standard"));
}

#[test]
fn typed_list_builder_append_composition_history_and_refusals_reopen() {
    let bytes = Builder::new()
        .paragraph(Paragraph::new("before"))
        .list(List::styled(
            "Numbered",
            [
                Item::new(Paragraph::new("one")),
                Item::new(Paragraph::new("two")),
            ],
        ))
        .build()
        .unwrap();
    let source = Template::from_bytes(bytes).unwrap();
    assert_eq!(source.text_body().unwrap().lists()[0].items().len(), 2);

    let mut text_edit = source.edit();
    text_edit
        .set_paragraph_text(Position::new(0), "after")
        .unwrap();
    let mut append_edit = source.edit();
    append_edit
        .append_block(List::new([Item::new(Paragraph::new("tail"))]))
        .unwrap();
    text_edit.join(append_edit).unwrap();
    let commit = text_edit.commit().unwrap();
    assert!(commit.template().content_xml().contains("<text:list>"));
    assert!(!commit.template().content_xml().contains("\n<"));
    assert_eq!(commit.template().text_body().unwrap().lists().len(), 2);
    assert_eq!(commit.patch().appended().len(), 1);
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.template())
            .unwrap()
            .as_bytes(),
        source.as_bytes(),
    );

    let target = commit.template().clone();
    let mut history = History::new(source.clone(), HistoryLimits::new(4, 1_000_000));
    history.record(commit).unwrap();
    assert_eq!(history.current().as_bytes(), target.as_bytes());
    assert!(history.undo());
    assert_eq!(history.current().as_bytes(), source.as_bytes());
    assert!(history.redo());

    let mut first = source.edit();
    first.set_paragraph_text(Position::new(0), "first").unwrap();
    let mut second = source.edit();
    second
        .set_paragraph_text(Position::new(0), "second")
        .unwrap();
    let error = first.join(second).err().unwrap();
    assert_eq!(error.failure(), JoinFailure::Paragraph(Position::new(0)));
    assert!(error.into_rejected().commit().unwrap().changed());

    let alternate = Template::from_bytes(
        Builder::new()
            .paragraph(Paragraph::new("alternate"))
            .build()
            .unwrap(),
    )
    .unwrap();
    let mut base_edit = source.edit();
    let cross_lineage = base_edit.join(alternate.edit()).err().unwrap();
    assert_eq!(cross_lineage.failure(), JoinFailure::DifferentSnapshot);

    let mut first_append = source.edit();
    first_append
        .append_block(Paragraph::new("first tail"))
        .unwrap();
    let mut second_append = source.edit();
    second_append
        .append_block(Paragraph::new("second tail"))
        .unwrap();
    assert_eq!(
        first_append.join(second_append).err().unwrap().failure(),
        JoinFailure::Append
    );
}

#[test]
fn rich_semantic_ranges_and_resources_fail_closed() {
    const PREFIX: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
        r#"xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text>"#,
    );
    const SUFFIX: &str = "</office:text></office:body></office:document-content>";
    for body in [
        r#"<text:p><text:bookmark-start text:name="open"/>text</text:p>"#,
        r#"<text:p><text:bookmark-end text:name="missing"/></text:p>"#,
        r"<text:p><draw:image/></text:p>",
        r"<text:list-item><text:p>orphan</text:p></text:list-item>",
    ] {
        let xml = format!("{PREFIX}{body}{SUFFIX}");
        assert!(Builder::new().content_xml(xml).build().is_err());
    }
}

fn structural_source() -> Template {
    Template::from_bytes(
        Builder::new()
            .heading(Heading::new(1, "Heading").unwrap())
            .paragraph(Paragraph::new("outside"))
            .list(List::new([
                Item::new(Paragraph::new("one")),
                Item::new(Paragraph::new("two")),
            ]))
            .build()
            .unwrap(),
    )
    .unwrap()
}

#[test]
fn heading_list_durable_and_merge_workflows_are_source_checked() {
    let source = structural_source();
    let mut left = source.edit();
    left.set_heading_text(Position::new(0), "Changed heading")
        .unwrap();
    left.set_paragraph_text(Position::new(0), "changed outside")
        .unwrap();
    let left_commit = left.commit().unwrap();

    let wire = left_commit.patch().to_bytes().unwrap();
    assert_eq!(wire, left_commit.patch().to_bytes().unwrap());
    let durable = Patch::from_bytes(&wire).unwrap();
    assert_eq!(durable.heading_changes()[0].before(), "Heading");
    assert_eq!(
        durable.apply(&source).unwrap().as_bytes(),
        left_commit.template().as_bytes()
    );
    let alternate = Template::from_bytes(Builder::new().build().unwrap()).unwrap();
    assert!(durable.apply(&alternate).is_err());
    let inverse_wire = durable.inverse().to_bytes().unwrap();
    let inverse = Patch::from_bytes(&inverse_wire).unwrap();
    assert_eq!(
        inverse.apply(left_commit.template()).unwrap().as_bytes(),
        source.as_bytes()
    );

    let mut right = source.edit();
    right
        .set_list(
            Position::new(0),
            List::styled("Bullets", [Item::new(Paragraph::new("replacement"))]),
        )
        .unwrap();
    let right_commit = right.commit().unwrap();
    let right_wire = right_commit.patch().to_bytes().unwrap();
    let right_durable = Patch::from_bytes(&right_wire).unwrap();
    assert_eq!(right_durable.list_changes().len(), 1);
    assert_eq!(
        right_durable.list_changes()[0]
            .after()
            .and_then(List::style_name),
        Some("Bullets")
    );
    assert_eq!(
        right_durable.apply(&source).unwrap().as_bytes(),
        right_commit.template().as_bytes()
    );
    let plan = Patch::plan_three_way(&source, left_commit.patch(), right_commit.patch()).unwrap();
    assert!(plan.conflicts().is_empty());
    let merged = plan.publish().unwrap();
    assert_eq!(
        merged.text_body().unwrap().headings()[0].text(),
        "Changed heading"
    );
    assert_eq!(
        merged.text_body().unwrap().lists()[0].style_name(),
        Some("Bullets")
    );

    let mut remove = source.edit();
    remove.remove_list(Position::new(0)).unwrap();
    let removed = remove.commit().unwrap();
    assert!(removed.template().text_body().unwrap().lists().is_empty());
    let removed_wire = removed.patch().to_bytes().unwrap();
    let durable_removed = Patch::from_bytes(&removed_wire).unwrap();
    assert!(durable_removed.list_changes()[0].after().is_none());
    let restored_wire = durable_removed.inverse().to_bytes().unwrap();
    let durable_restored = Patch::from_bytes(&restored_wire).unwrap();
    assert_eq!(
        durable_restored
            .apply(removed.template())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn rich_inline_crud_is_durable_reversible_and_reopened() {
    let source = structural_source();
    let content = [
        Inline::text("before "),
        Inline::bookmark("point"),
        Inline::Link(Link::new("https://example.test/oth", "linked")),
        Inline::Span(Span::new("Strong", " styled")),
        Inline::Field(
            InlineField::new(field::Kind::Date, " today")
                .with_value("2026-08-10")
                .fixed(),
        ),
        Inline::bookmark_start("range"),
        Inline::text(" ranged"),
        Inline::bookmark_end("range"),
    ];
    let mut edit = source.edit();
    edit.set_paragraph_inline(Position::new(0), &content)
        .unwrap();
    let commit = edit.commit().unwrap();
    let body = commit.template().text_body().unwrap();
    let paragraph = &body.paragraphs()[0];
    assert_eq!(paragraph.text(), "before linked styled today ranged");
    assert_eq!(paragraph.links()[0].href(), "https://example.test/oth");
    assert_eq!(paragraph.formatting_runs()[0].style_name(), "Strong");
    assert!(matches!(paragraph.fields()[0].kind(), field::Kind::Date));
    assert_eq!(body.bookmarks().len(), 2);
    assert!(!commit.template().content_xml().contains(">\n<"));

    let wire = commit.patch().to_bytes().unwrap();
    let durable = Patch::from_bytes(&wire).unwrap();
    assert_eq!(durable.inline_changes().len(), 1);
    assert_eq!(
        durable.apply(&source).unwrap().as_bytes(),
        commit.template().as_bytes()
    );
    let inverse_wire = durable.inverse().to_bytes().unwrap();
    let inverse = Patch::from_bytes(&inverse_wire).unwrap();
    assert_eq!(
        inverse.apply(commit.template()).unwrap().as_bytes(),
        source.as_bytes()
    );
}

#[test]
fn inert_form_catalog_crud_is_durable_and_atomic() {
    let source = structural_source();
    let search =
        Form::new("Search").with_control(Control::new("text").with_id("query").with_name("q"));
    let mut create = source.edit();
    create.set_forms(&[search]).unwrap();
    let created = create.commit().unwrap();
    assert_eq!(created.template().text_body().unwrap().forms().len(), 1);
    assert_eq!(
        created.template().text_body().unwrap().forms()[0].controls()[0].id(),
        Some("query")
    );
    let durable = Patch::from_bytes(&created.patch().to_bytes().unwrap()).unwrap();
    assert_eq!(durable.forms_change().unwrap().after().len(), 1);

    let updated_form = Form::new("Search").with_control(Control::new("button").with_name("go"));
    let mut update = created.template().edit();
    update.set_forms(&[updated_form]).unwrap();
    let updated = update.commit().unwrap();
    assert_eq!(
        updated.template().text_body().unwrap().forms()[0].controls()[0].kind(),
        "button"
    );

    let mut remove = updated.template().edit();
    remove.set_forms(&[]).unwrap();
    let removed = remove.commit().unwrap();
    assert!(removed.template().text_body().unwrap().forms().is_empty());
    assert_eq!(
        removed
            .patch()
            .inverse()
            .apply(removed.template())
            .unwrap()
            .as_bytes(),
        updated.template().as_bytes()
    );
}

#[test]
fn isolated_nested_list_replace_remove_is_durable() {
    const CONTENT: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:text>"#,
        r#"<text:list><text:list-item><text:p>outer</text:p><text:list><text:list-item><text:p>inner</text:p></text:list-item></text:list></text:list-item></text:list>"#,
        r#"</office:text></office:body></office:document-content>"#,
    );
    let source =
        Template::from_bytes(Builder::new().content_xml(CONTENT).build().unwrap()).unwrap();
    assert_eq!(source.text_body().unwrap().lists()[0].level(), 2);
    let mut replacement_edit = source.edit();
    replacement_edit
        .set_list(
            Position::new(0),
            List::new([Item::new(Paragraph::new("nested replacement"))]),
        )
        .unwrap();
    let replacement_commit = replacement_edit.commit().unwrap();
    assert_eq!(
        replacement_commit.template().text_body().unwrap().lists()[0].level(),
        2
    );
    assert_eq!(
        replacement_commit.template().text_body().unwrap().lists()[0].items()[0].paragraphs()[0]
            .text(),
        "nested replacement"
    );
    let durable = Patch::from_bytes(&replacement_commit.patch().to_bytes().unwrap()).unwrap();
    assert_eq!(durable.list_changes()[0].before().unwrap().level(), 2);

    let mut removal_edit = source.edit();
    removal_edit.remove_list(Position::new(0)).unwrap();
    let removal_commit = removal_edit.commit().unwrap();
    assert_eq!(
        removal_commit.template().text_body().unwrap().lists().len(),
        1
    );
    assert_eq!(
        removal_commit.template().text_body().unwrap().lists()[0].level(),
        1
    );
}

#[test]
fn metadata_and_styles_are_durable_and_reversible() {
    let source = structural_source();
    let metadata = Metadata {
        title: Some("Durable title".to_string()),
        author: Some("Template author".to_string()),
        page_count: Some(3),
        ..Metadata::default()
    };
    let body_style = Style::new("Body2", "paragraph")
        .unwrap()
        .with_parent("Standard")
        .unwrap()
        .with_text_properties(
            TextProperties::new()
                .with_color("#aa00cc")
                .unwrap()
                .with_background_color("#00ff00")
                .unwrap()
                .with_weight(Weight::Bold)
                .with_slant(Slant::Italic),
        );
    let mut parts = source.edit();
    parts.set_metadata(&metadata).unwrap();
    parts.set_styles(&[body_style]).unwrap();
    let parts_commit = parts.commit().unwrap();
    assert_eq!(
        parts_commit
            .template()
            .metadata()
            .and_then(|value| value.title.as_deref()),
        Some("Durable title")
    );
    assert_eq!(parts_commit.template().styles()[0].name(), "Body2");
    assert_eq!(
        parts_commit.template().styles()[0].parent_name(),
        Some("Standard")
    );
    let properties = parts_commit.template().styles()[0]
        .text_properties()
        .unwrap();
    assert_eq!(properties.color(), Some("#AA00CC"));
    assert_eq!(properties.background_color(), Some("#00FF00"));
    assert_eq!(properties.weight(), Some(Weight::Bold));
    assert_eq!(properties.slant(), Some(Slant::Italic));
    let parts_wire = parts_commit.patch().to_bytes().unwrap();
    let parts_durable = Patch::from_bytes(&parts_wire).unwrap();
    assert_eq!(
        parts_durable.apply(&source).unwrap().as_bytes(),
        parts_commit.template().as_bytes()
    );
    assert_eq!(
        parts_durable
            .inverse()
            .apply(parts_commit.template())
            .unwrap()
            .as_bytes(),
        source.as_bytes()
    );

    let mut removal = parts_commit.template().edit();
    removal.remove_metadata();
    removal.remove_styles();
    let removed = removal.commit().unwrap();
    assert!(removed.template().meta_xml().is_none());
    assert!(removed.template().styles_xml().is_none());
    assert!(removed.patch().removes_metadata());
    assert!(removed.patch().removes_styles());
    let removal_wire = removed.patch().to_bytes().unwrap();
    let durable_removal = Patch::from_bytes(&removal_wire).unwrap();
    assert!(durable_removal.removes_metadata());
    assert!(durable_removal.removes_styles());
    assert_eq!(
        durable_removal
            .inverse()
            .apply(removed.template())
            .unwrap()
            .as_bytes(),
        parts_commit.template().as_bytes()
    );
}

#[test]
fn security_policy_is_explicit_and_default_deny() {
    const ACTIVE: &str = concat!(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content "#,
        r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" "#,
        r#"xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" "#,
        r#"xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" "#,
        r#"xmlns:form="urn:oasis:names:tc:opendocument:xmlns:form:1.0" "#,
        r#"xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:text>"#,
        r#"<text:p><draw:image xlink:href="Pictures/a.png"/><draw:image xlink:href="https://example.test/a.png"/></text:p>"#,
        r#"<office:forms><form:form form:name="f"><form:text form:name="q"/></form:form></office:forms>"#,
        r#"</office:text></office:body></office:document-content>"#,
    );
    let active = Template::from_bytes(Builder::new().content_xml(ACTIVE).build().unwrap()).unwrap();
    assert!(active.check_security(SecurityPolicy::default()).is_err());
    let report = active
        .check_security(SecurityPolicy {
            allow_embedded_objects: true,
            allow_external_resources: true,
            allow_forms: true,
            allow_scripts: false,
            allow_signatures: false,
        })
        .unwrap();
    assert_eq!(report.embedded_objects, 1);
    assert_eq!(report.external_resources, 1);
    assert_eq!(report.forms, 1);
}

#[test]
fn transfer_plan_resolves_style_parents_and_refuses_collisions() {
    let source_without_styles = Template::from_bytes(
        Builder::new()
            .paragraph(Paragraph::styled("transfer me", "Body"))
            .build()
            .unwrap(),
    )
    .unwrap();
    let parent = Style::new("Base", "paragraph").unwrap();
    let body = Style::new("Body", "paragraph")
        .unwrap()
        .with_parent("Base")
        .unwrap()
        .with_text_properties(TextProperties::new().with_weight(Weight::Bold));
    let mut source_edit = source_without_styles.edit();
    source_edit
        .set_styles(&[parent.clone(), body.clone()])
        .unwrap();
    let source = source_edit.commit().unwrap().into_template();
    let destination = Template::from_bytes(Builder::new().build().unwrap()).unwrap();

    let plan = destination
        .plan_transfer_from(
            &source,
            TransferSelector::Paragraph(Position::new(0)),
            TransferPolicy::default(),
        )
        .unwrap();
    assert_eq!(plan.imported_styles().len(), 2);
    let published = plan.publish().unwrap();
    assert_eq!(
        published.template().text_body().unwrap().paragraphs()[0].text(),
        "transfer me"
    );
    assert!(
        published
            .template()
            .styles()
            .iter()
            .any(|style| style.name() == "Base")
    );

    assert!(
        destination
            .plan_transfer_from(
                &source,
                TransferSelector::Paragraph(Position::new(0)),
                TransferPolicy {
                    include_styles: false,
                    ..TransferPolicy::default()
                },
            )
            .is_err()
    );

    let colliding_source = Template::from_bytes(Builder::new().build().unwrap()).unwrap();
    let mut colliding_edit = colliding_source.edit();
    colliding_edit
        .set_styles(&[Style::new("Body", "text").unwrap()])
        .unwrap();
    let colliding = colliding_edit.commit().unwrap().into_template();
    assert!(
        colliding
            .plan_transfer_from(
                &source,
                TransferSelector::Paragraph(Position::new(0)),
                TransferPolicy::default(),
            )
            .is_err()
    );
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
    let bytes = include_bytes!("fixtures/libreoffice-desktop-html.oth").to_vec();
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
    assert_eq!(commit.template().meta_xml(), source.meta_xml());
    assert_eq!(commit.template().styles_xml(), source.styles_xml());
    assert!(commit.template().styles_xml().is_some());
    let mut rich_edit = commit.template().edit();
    rich_edit
        .set_paragraph_inline(
            Position::new(0),
            &[
                Inline::bookmark("writer-web"),
                Inline::text("Web "),
                Inline::Link(Link::new("https://example.test/writer", "template")),
            ],
        )
        .unwrap();
    let rich_commit = rich_edit.commit().unwrap();
    let rich_reopen = Template::from_bytes(rich_commit.template().as_bytes().to_vec()).unwrap();
    assert_eq!(
        rich_reopen.text_body().unwrap().paragraphs()[0].text(),
        "Web template"
    );
    assert_eq!(rich_reopen.text_body().unwrap().bookmarks().len(), 1);
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
