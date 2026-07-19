use litchi_odf::{
    MutableDocument, OdfVariableBody, OdfVariableDeclaration, OdfVariableDeclarationGroup,
    OdfVariableKind, OdfVariablePart, OdfVariableScope, OdfVariableValue, OdfVariableValueType,
};

fn group(
    kind: OdfVariableKind,
    declarations: Vec<OdfVariableDeclaration>,
) -> OdfVariableDeclarationGroup {
    OdfVariableDeclarationGroup {
        kind,
        part: OdfVariablePart::Content,
        scope: OdfVariableScope::Body(OdfVariableBody::Text),
        declarations,
    }
}

#[test]
fn canonical_writer_escapes_and_round_trips_every_declaration_kind() {
    let simple = group(
        OdfVariableKind::Simple,
        vec![OdfVariableDeclaration::Simple {
            name: "amount<&\"".to_string(),
            value_type: OdfVariableValueType::Currency,
        }],
    );
    let user = group(
        OdfVariableKind::User,
        vec![OdfVariableDeclaration::User {
            name: "caption".to_string(),
            value: Some(OdfVariableValue::String {
                value: "cached <&> value".to_string(),
            }),
            formula: Some("of:=CONCATENATE(&quot;a&quot;;&quot;b&quot;)".to_string()),
        }],
    );
    let sequence = group(
        OdfVariableKind::Sequence,
        vec![OdfVariableDeclaration::Sequence {
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
        OdfVariableKind::User,
        vec![OdfVariableDeclaration::User {
            name: "customer".to_string(),
            value: Some(OdfVariableValue::String {
                value: "A".to_string(),
            }),
            formula: None,
        }],
    );
    let simple = group(
        OdfVariableKind::Simple,
        vec![OdfVariableDeclaration::Simple {
            name: "counter".to_string(),
            value_type: OdfVariableValueType::Float,
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
    assert_eq!(declarations.groups[0].kind, OdfVariableKind::Simple);
    assert_eq!(declarations.groups[1].kind, OdfVariableKind::User);

    let replacement = group(
        OdfVariableKind::User,
        vec![OdfVariableDeclaration::User {
            name: "customer".to_string(),
            value: Some(OdfVariableValue::String {
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
            .find(OdfVariableKind::User, "customer"),
        replacement.declarations.first(),
    );

    let removed = document
        .remove_variable_declaration_group(
            OdfVariablePart::Content,
            &OdfVariableScope::Body(OdfVariableBody::Text),
            OdfVariableKind::Simple,
        )
        .unwrap()
        .unwrap();
    assert_eq!(removed, simple);
    assert!(
        document
            .variable_declarations()
            .unwrap()
            .find(OdfVariableKind::Simple, "counter")
            .is_none()
    );
}

#[test]
fn writer_rejects_kind_mismatch_invalid_sequence_and_duplicate_names() {
    let mismatch = group(
        OdfVariableKind::Simple,
        vec![OdfVariableDeclaration::Sequence {
            name: "wrong".to_string(),
            display_outline_level: 1,
            separation_character: None,
        }],
    );
    assert!(mismatch.to_xml().is_err());

    let invalid_sequence = group(
        OdfVariableKind::Sequence,
        vec![OdfVariableDeclaration::Sequence {
            name: "bad".to_string(),
            display_outline_level: 0,
            separation_character: Some('.'),
        }],
    );
    assert!(invalid_sequence.to_xml().is_err());

    let duplicates = group(
        OdfVariableKind::Simple,
        vec![
            OdfVariableDeclaration::Simple {
                name: "same".to_string(),
                value_type: OdfVariableValueType::Float,
            },
            OdfVariableDeclaration::Simple {
                name: "same".to_string(),
                value_type: OdfVariableValueType::String,
            },
        ],
    );
    assert!(duplicates.to_xml().is_err());
}
