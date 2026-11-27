//! Regression tests for tracked-change declaration and marker codecs.

use crate::parser::{ChangeType, Parser, TrackChange, TrackedChanges};

fn change(id: &str, change_type: ChangeType, content: &str) -> TrackChange {
    TrackChange {
        id: id.to_string(),
        xml_id: Some(id.to_string()),
        author: Some("A & B".to_string()),
        date: Some("2026-07-17T12:00:00+08:00".to_string()),
        comment: None,
        change_type,
        style_name: (change_type == ChangeType::FormatChange).then(|| "Changed Style".to_string()),
        merge_last_paragraph: (change_type == ChangeType::Deletion).then_some(false),
        content: content.to_string(),
    }
}

#[test]
fn serializes_and_reparses_all_declaration_kinds() {
    let declarations = TrackedChanges {
        track_changes: Some(true),
        protection_key: Some("YWJj".to_string()),
        protection_key_digest_algorithm: Some("urn:example:sha256".to_string()),
        changes: vec![
            change("insert_1", ChangeType::Insertion, "live text"),
            change("delete_1", ChangeType::Deletion, "gone  text\t<&\nnext"),
            change("format_1", ChangeType::FormatChange, "formatted"),
        ],
    };
    let xml = declarations.to_xml_fragment().unwrap();
    assert!(xml.contains("<text:s text:c=\"2\"/>"));
    assert!(xml.contains("<text:tab/>"));
    assert!(xml.contains("&lt;&amp;"));
    assert!(!xml.contains("live text"));

    let parsed = Parser::parse_tracked_changes(&xml).unwrap();
    assert_eq!(parsed.track_changes, Some(true));
    assert_eq!(parsed.protection_key.as_deref(), Some("YWJj"));
    assert_eq!(parsed.changes.len(), 3);
    assert_eq!(parsed.changes[1].content, "gone  text\t<&\nnext");
    assert_eq!(
        parsed.changes[2].style_name.as_deref(),
        Some("Changed Style")
    );
}

#[test]
fn rejects_invalid_constructed_declarations() {
    let mut declarations = TrackedChanges {
        changes: vec![change("same", ChangeType::Insertion, "")],
        ..TrackedChanges::default()
    };
    declarations.changes.push(declarations.changes[0].clone());
    assert!(declarations.to_xml_fragment().is_err());

    declarations.changes.truncate(1);
    declarations.changes[0].id = "bad:id".to_string();
    assert!(declarations.to_xml_fragment().is_err());

    declarations.changes[0].id = "valid".to_string();
    declarations.changes[0].xml_id = Some("other".to_string());
    assert!(declarations.to_xml_fragment().is_err());

    declarations.changes[0].xml_id = Some("valid".to_string());
    declarations.protection_key_digest_algorithm = Some("urn:sha256".to_string());
    assert!(declarations.to_xml_fragment().is_err());
}

#[test]
fn parses_a_libreoffice_flat_document_before_serializing() {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/libreoffice-core/sw/qa/uitest/data/redline-autocorrect.fodt");
    let Ok(xml) = std::fs::read_to_string(path) else {
        return;
    };
    let parsed = Parser::parse_tracked_changes(&xml).unwrap();
    assert!(!parsed.changes.is_empty());
    let serialized = parsed.to_xml_fragment().unwrap();
    let reparsed = Parser::parse_tracked_changes(&serialized).unwrap();
    assert_eq!(reparsed.changes.len(), parsed.changes.len());
}
