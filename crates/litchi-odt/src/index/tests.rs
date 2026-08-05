//! Focused regression tests for the layered text-index owner.

use super::codec::{MAX_INDEX_DEPTH, parse_text_indexes};
use super::{TextIndexContent, TextIndexKind};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

#[test]
fn parses_every_text_index_kind_and_complete_ordered_subtrees() {
    let xml = format!(
        r#"<o:document-content xmlns:o="{OFFICE}" xmlns:t="{TEXT}" xmlns:u="urn:vendor"><o:body><o:text><t:table-of-content t:name="Contents &amp; More" t:protected="1" u:future="yes"><t:table-of-content-source t:outline-level="3"><t:index-title-template t:style-name="Title">Contents</t:index-title-template><t:table-of-content-entry-template t:outline-level="1" t:style-name="Entry"><t:index-entry-text/><u:extension u:value="A&amp;B"/></t:table-of-content-entry-template></t:table-of-content-source><t:index-body><t:index-title t:name="Cached"><t:p>Title</t:p></t:index-title><t:p>Pre&amp;<t:span>Mid</t:span><![CDATA[!]]>Post</t:p><t:user-index t:name="Nested"><t:user-index-source t:index-name="N"/><t:index-body><t:p>Nested</t:p></t:index-body></t:user-index></t:index-body></t:table-of-content><t:illustration-index t:name="I"/><t:table-index t:name="T"/><t:object-index t:name="O"/><t:alphabetical-index t:name="A"/><t:bibliography t:name="B"/></o:text></o:body></o:document-content>"#
    );
    let indexes = parse_text_indexes(&xml).unwrap();
    assert_eq!(indexes.len(), 7);
    assert_eq!(indexes[0].kind(), TextIndexKind::TableOfContents);
    assert_eq!(indexes[0].name(), "Contents & More");
    assert!(indexes[0].protected());
    assert_eq!(
        indexes[0].root().attribute(Some("urn:vendor"), "future"),
        Some("yes")
    );
    let source = indexes[0].source().unwrap();
    assert_eq!(source.local_name(), "table-of-content-source");
    assert_eq!(source.attribute(Some(TEXT), "outline-level"), Some("3"));
    let template = source
        .child_elements()
        .find(|element| element.local_name() == "table-of-content-entry-template")
        .unwrap();
    let extension = template
        .child_elements()
        .find(|element| element.namespace_uri() == Some("urn:vendor"))
        .unwrap();
    assert_eq!(
        extension.attribute(Some("urn:vendor"), "value"),
        Some("A&B")
    );

    let body = indexes[0].body().unwrap();
    let paragraph = body
        .child_elements()
        .find(|element| element.local_name() == "p" && element.all_text().starts_with("Pre"))
        .unwrap();
    assert_eq!(paragraph.all_text(), "Pre&Mid!Post");
    assert!(matches!(paragraph.content()[0], TextIndexContent::Text(_)));
    assert!(matches!(
        paragraph.content()[1],
        TextIndexContent::Element(_)
    ));
    assert!(matches!(paragraph.content()[2], TextIndexContent::Text(_)));

    assert_eq!(indexes[1].kind(), TextIndexKind::User);
    assert_eq!(indexes[1].name(), "Nested");
    assert_eq!(indexes[1].body().unwrap().all_text(), "Nested");
    assert_eq!(indexes[2].kind(), TextIndexKind::Illustration);
    assert_eq!(indexes[3].kind(), TextIndexKind::Table);
    assert_eq!(indexes[4].kind(), TextIndexKind::Object);
    assert_eq!(indexes[5].kind(), TextIndexKind::Alphabetical);
    assert_eq!(indexes[6].kind(), TextIndexKind::Bibliography);
}

#[test]
fn text_indexes_reject_malformed_ambiguous_or_invalid_roots() {
    let missing_name = format!(r#"<t:table-of-content xmlns:t="{TEXT}"/>"#);
    assert!(parse_text_indexes(&missing_name).is_err());

    let invalid_boolean =
        format!(r#"<t:table-index xmlns:t="{TEXT}" t:name="T" t:protected="yes"/>"#);
    assert!(parse_text_indexes(&invalid_boolean).is_err());

    let duplicate =
        format!(r#"<t:object-index xmlns:t="{TEXT}" xmlns:u="{TEXT}" t:name="A" u:name="B"/>"#);
    assert!(parse_text_indexes(&duplicate).is_err());

    let unknown_prefix = format!(r#"<t:bibliography xmlns:t="{TEXT}" t:name="B" x:value="bad"/>"#);
    assert!(parse_text_indexes(&unknown_prefix).is_err());
    assert!(parse_text_indexes("<t:table-index>").is_err());
}

#[test]
fn text_indexes_enforce_nesting_bound() {
    let mut xml = format!(r#"<t:table-of-content xmlns:t="{TEXT}" t:name="T">"#);
    for _ in 0..MAX_INDEX_DEPTH {
        xml.push_str("<t:span>");
    }
    for _ in 0..MAX_INDEX_DEPTH {
        xml.push_str("</t:span>");
    }
    xml.push_str("</t:table-of-content>");
    assert!(parse_text_indexes(&xml).is_err());
}
