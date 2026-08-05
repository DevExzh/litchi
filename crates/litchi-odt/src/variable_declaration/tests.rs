//! Focused regression tests for the layered variable declaration owner.

use super::{
    Body, Declaration, Group, Kind, Part, Scope, ValueType, parse_parts, remove_xml, set_xml,
};

const OFFICE: &str = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const TEXT: &str = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";

fn content(body: &str) -> String {
    format!(
        r#"<office:document-content xmlns:office="{OFFICE}" xmlns:text="{TEXT}"><office:body><office:text>{body}</office:text></office:body></office:document-content>"#
    )
}

fn simple(name: &str) -> Group {
    Group {
        kind: Kind::Simple,
        part: Part::Content,
        scope: Scope::Body(Body::Text),
        declarations: vec![Declaration::Simple {
            name: name.to_owned(),
            value_type: ValueType::Float,
        }],
    }
}

#[test]
fn layered_codec_preserves_unknown_siblings_and_round_trips_model() {
    let original = content(
        r#"<!--keep--><custom:before xmlns:custom="urn:custom"/><text:user-field-decls/><custom:after xmlns:custom="urn:custom"/>"#,
    );
    let inserted = set_xml(&original, &simple("counter")).unwrap();
    assert!(inserted.contains("<!--keep--><custom:before"));
    assert!(inserted.contains("<custom:after xmlns:custom=\"urn:custom\"/>"));

    let parsed = parse_parts(&[(&inserted, Part::Content)]).unwrap();
    assert_eq!(parsed.groups.len(), 2);
    assert_eq!(
        parsed.find(Kind::Simple, "counter").unwrap().name(),
        "counter"
    );

    let removed = remove_xml(&inserted, &Scope::Body(Body::Text), Kind::Simple).unwrap();
    assert!(removed.contains("<text:user-field-decls/>"));
    assert!(removed.contains("<custom:before"));
    assert!(removed.contains("<custom:after"));
}

#[test]
fn model_keeps_sequence_ergonomics() {
    let declaration = Declaration::Sequence {
        name: "chapter".to_owned(),
        display_outline_level: 2,
        separation_character: None,
    };
    assert_eq!(declaration.effective_separation_character(), Some('.'));
}
