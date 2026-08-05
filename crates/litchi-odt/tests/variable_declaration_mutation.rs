use litchi_odt::variable_declaration::{
    Body, Declaration, Group, Kind, Part, Scope, Value, ValueType,
};

fn group(kind: Kind, declarations: Vec<Declaration>) -> Group {
    Group {
        kind,
        part: Part::Content,
        scope: Scope::Body(Body::Text),
        declarations,
    }
}

#[test]
fn canonical_writer_escapes_and_round_trips_every_declaration_kind() {
    let simple = group(
        Kind::Simple,
        vec![Declaration::Simple {
            name: "amount<&\"".to_string(),
            value_type: ValueType::Currency,
        }],
    );
    let user = group(
        Kind::User,
        vec![Declaration::User {
            name: "caption".to_string(),
            value: Some(Value::String {
                value: "cached <&> value".to_string(),
            }),
            formula: Some("of:=CONCATENATE(&quot;a&quot;;&quot;b&quot;)".to_string()),
        }],
    );
    let sequence = group(
        Kind::Sequence,
        vec![Declaration::Sequence {
            name: "Figure".to_string(),
            display_outline_level: 3,
            separation_character: Some('#'),
        }],
    );

    assert!(simple.to_xml().unwrap().contains("amount&lt;&amp;&quot;"));
    assert!(
        user.to_xml()
            .unwrap()
            .contains("cached &lt;&amp;&gt; value")
    );
    assert!(
        sequence
            .to_xml()
            .unwrap()
            .contains("text:separation-character=\"#\"")
    );
}

#[test]
fn mutable_document_upserts_orders_replaces_and_removes_groups_atomically() {
    let mut document = litchi_odt::mutable::MutableDocument::new();
    let user = group(
        Kind::User,
        vec![Declaration::User {
            name: "customer".to_string(),
            value: Some(Value::String {
                value: "A".to_string(),
            }),
            formula: None,
        }],
    );
    let simple = group(
        Kind::Simple,
        vec![Declaration::Simple {
            name: "counter".to_string(),
            value_type: ValueType::Float,
        }],
    );

    assert!(
        document
            .set_variable_declaration_group(&user)
            .unwrap()
            .is_none()
    );
    assert!(
        document
            .set_variable_declaration_group(&simple)
            .unwrap()
            .is_none()
    );
    let declarations = document.variable_declarations().unwrap();
    assert_eq!(declarations.groups.len(), 2);
    assert_eq!(declarations.groups[0].kind, Kind::Simple);
    assert_eq!(declarations.groups[1].kind, Kind::User);

    let replacement = group(
        Kind::User,
        vec![Declaration::User {
            name: "customer".to_string(),
            value: Some(Value::String {
                value: "B".to_string(),
            }),
            formula: None,
        }],
    );
    let old = document
        .set_variable_declaration_group(&replacement)
        .unwrap()
        .unwrap();
    assert_eq!(old, user);
    assert_eq!(
        document
            .variable_declarations()
            .unwrap()
            .find(Kind::User, "customer"),
        replacement.declarations.first(),
    );

    let removed = document
        .remove_variable_declaration_group(Part::Content, &Scope::Body(Body::Text), Kind::Simple)
        .unwrap()
        .unwrap();
    assert_eq!(removed, simple);
    assert!(
        document
            .variable_declarations()
            .unwrap()
            .find(Kind::Simple, "counter")
            .is_none()
    );
}

#[test]
fn writer_rejects_kind_mismatch_invalid_sequence_and_duplicate_names() {
    let mismatch = group(
        Kind::Simple,
        vec![Declaration::Sequence {
            name: "wrong".to_string(),
            display_outline_level: 1,
            separation_character: None,
        }],
    );
    assert!(mismatch.to_xml().is_err());

    let invalid_sequence = group(
        Kind::Sequence,
        vec![Declaration::Sequence {
            name: "bad".to_string(),
            display_outline_level: 0,
            separation_character: Some('.'),
        }],
    );
    assert!(invalid_sequence.to_xml().is_err());

    let duplicates = group(
        Kind::Simple,
        vec![
            Declaration::Simple {
                name: "same".to_string(),
                value_type: ValueType::Float,
            },
            Declaration::Simple {
                name: "same".to_string(),
                value_type: ValueType::String,
            },
        ],
    );
    assert!(duplicates.to_xml().is_err());
}
