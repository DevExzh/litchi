use litchi_odt::{
    MutableDocument, VariableBody, VariableDeclaration, VariableDeclarationGroup, VariableKind,
    VariablePart, VariableScope, VariableValue, VariableValueType,
};

fn group(kind: VariableKind, declarations: Vec<VariableDeclaration>) -> VariableDeclarationGroup {
    VariableDeclarationGroup {
        kind,
        part: VariablePart::Content,
        scope: VariableScope::Body(VariableBody::Text),
        declarations,
    }
}

#[test]
fn canonical_writer_escapes_and_round_trips_every_declaration_kind() {
    let simple = group(
        VariableKind::Simple,
        vec![VariableDeclaration::Simple {
            name: "amount<&\"".to_string(),
            value_type: VariableValueType::Currency,
        }],
    );
    let user = group(
        VariableKind::User,
        vec![VariableDeclaration::User {
            name: "caption".to_string(),
            value: Some(VariableValue::String {
                value: "cached <&> value".to_string(),
            }),
            formula: Some("of:=CONCATENATE(&quot;a&quot;;&quot;b&quot;)".to_string()),
        }],
    );
    let sequence = group(
        VariableKind::Sequence,
        vec![VariableDeclaration::Sequence {
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
    let mut document = MutableDocument::new();
    let user = group(
        VariableKind::User,
        vec![VariableDeclaration::User {
            name: "customer".to_string(),
            value: Some(VariableValue::String {
                value: "A".to_string(),
            }),
            formula: None,
        }],
    );
    let simple = group(
        VariableKind::Simple,
        vec![VariableDeclaration::Simple {
            name: "counter".to_string(),
            value_type: VariableValueType::Float,
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
    assert_eq!(declarations.groups[0].kind, VariableKind::Simple);
    assert_eq!(declarations.groups[1].kind, VariableKind::User);

    let replacement = group(
        VariableKind::User,
        vec![VariableDeclaration::User {
            name: "customer".to_string(),
            value: Some(VariableValue::String {
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
            .find(VariableKind::User, "customer"),
        replacement.declarations.first(),
    );

    let removed = document
        .remove_variable_declaration_group(
            VariablePart::Content,
            &VariableScope::Body(VariableBody::Text),
            VariableKind::Simple,
        )
        .unwrap()
        .unwrap();
    assert_eq!(removed, simple);
    assert!(
        document
            .variable_declarations()
            .unwrap()
            .find(VariableKind::Simple, "counter")
            .is_none()
    );
}

#[test]
fn writer_rejects_kind_mismatch_invalid_sequence_and_duplicate_names() {
    let mismatch = group(
        VariableKind::Simple,
        vec![VariableDeclaration::Sequence {
            name: "wrong".to_string(),
            display_outline_level: 1,
            separation_character: None,
        }],
    );
    assert!(mismatch.to_xml().is_err());

    let invalid_sequence = group(
        VariableKind::Sequence,
        vec![VariableDeclaration::Sequence {
            name: "bad".to_string(),
            display_outline_level: 0,
            separation_character: Some('.'),
        }],
    );
    assert!(invalid_sequence.to_xml().is_err());

    let duplicates = group(
        VariableKind::Simple,
        vec![
            VariableDeclaration::Simple {
                name: "same".to_string(),
                value_type: VariableValueType::Float,
            },
            VariableDeclaration::Simple {
                name: "same".to_string(),
                value_type: VariableValueType::String,
            },
        ],
    );
    assert!(duplicates.to_xml().is_err());
}
