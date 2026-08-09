#![allow(
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use super::codec::{annotations, apply_edits, serialize};
use super::{Anchor, Annotation, Position};

const NS: &str = r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:foo="urn:example:foreign""#;

fn document(body: &str) -> String {
    format!(
        r"<office:document-content {NS}><office:body><office:presentation>{body}</office:presentation></office:body></office:document-content>"
    )
}

#[test]
fn inventories_page_and_nested_shape_anchors() {
    let xml = document(
        r#"<draw:page><draw:frame draw:name="title"><office:annotation office:name="shape"><text:p>shape</text:p><foo:foreign foo:value="1">opaque</foo:foreign></office:annotation></draw:frame><office:annotation office:name="page"><text:p>page</text:p></office:annotation></draw:page>"#,
    );

    let items = annotations(&xml).unwrap();
    assert_eq!(items.len(), 2);
    assert_eq!(
        items[0].anchor.position(),
        &Position::Shape {
            page_index: 0,
            name: "title".to_string()
        }
    );
    assert_eq!(items[0].annotation.text(), "shape\nopaque");
    assert_eq!(items[1].anchor.position(), &Position::Page { index: 0 });
}

#[test]
fn nested_office_annotation_is_retained_as_body_content() {
    let xml = document(
        r#"<draw:page><office:annotation office:name="outer"><text:p>before<office:annotation><text:p>nested</text:p></office:annotation>after</text:p></office:annotation></draw:page>"#,
    );
    let items = annotations(&xml).unwrap();
    assert_eq!(items.len(), 1);
    assert_eq!(items[0].annotation.text(), "beforenestedafter");
}

#[test]
fn serialization_adds_standard_namespace_bindings_and_keeps_foreign_content() {
    let mut annotation = Annotation::new("hello & goodbye");
    let mut foreign = litchi_odf_common::annotation::Element::new("foo:foreign").unwrap();
    foreign.set_attribute("foo:value", "x&y").unwrap();
    foreign.push_text("opaque");
    annotation
        .set_namespace("foo", "urn:example:foreign")
        .unwrap();
    annotation.push_element(foreign);

    let xml = serialize(&annotation).unwrap();
    assert!(xml.contains("xmlns:foo=\"urn:example:foreign\""));
    assert!(xml.contains("foo:value=\"x&amp;y\""));
    assert!(xml.contains("hello &amp; goodbye"));
}

#[test]
fn invalid_anchor_names_are_rejected_before_lookup() {
    assert!(Anchor::shape(0, "").is_err());
    assert!(Anchor::shape(0, "title\n").is_err());
}

#[test]
fn edits_preserve_unselected_xml_byte_for_byte() {
    let xml = document(
        r#"<draw:page><office:annotation office:name="first"><text:p>one</text:p></office:annotation><draw:rect draw:name="keep" foo:unknown="yes"><foo:opaque/></draw:rect><office:annotation office:name="second"><text:p>two</text:p></office:annotation></draw:page>"#,
    );
    let items = annotations(&xml).unwrap();
    let replacement = serialize(&Annotation::new("changed")).unwrap();
    let edited = apply_edits(
        &xml,
        vec![super::codec::Edit {
            start: {
                let marker = "<office:annotation office:name=\"first\">";
                xml.find(marker).unwrap()
            },
            end: {
                let marker = "</office:annotation>";
                let start = xml
                    .find("<office:annotation office:name=\"first\">")
                    .unwrap();
                xml[start..].find(marker).unwrap() + start + marker.len()
            },
            replacement,
        }],
    )
    .unwrap();
    assert_eq!(items.len(), 2);
    assert!(edited.contains(r#"draw:rect draw:name="keep" foo:unknown="yes"><foo:opaque/>"#));
    assert!(edited.contains(
        r#"<office:annotation office:name="second"><text:p>two</text:p></office:annotation>"#
    ));
}
