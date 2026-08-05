use super::model::{Paragraph, Run};
use crate::parts::chp::CharacterProperties;
use crate::parts::revisions::RevisionAuthorTable;
use crate::revision::RevisionKind;

#[test]
fn resolves_run_revision_authors_and_timestamps() {
    let timestamp =
        30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
    let mut run = Run::new(
        "changed".to_string(),
        CharacterProperties {
            is_revision_inserted: Some(true),
            revision_author_index: Some(1),
            revision_timestamp: Some(timestamp),
            revision_id: Some(42),
            insertion_revision_save_id: Some(0x11223344),
            ..CharacterProperties::default()
        },
    );
    let authors = RevisionAuthorTable::from_authors(&["Unknown", "Alice"]);
    run.resolve_revisions(&authors).unwrap();
    let revision = run.insertion_revision().unwrap();
    assert_eq!(revision.kind, RevisionKind::Insertion);
    assert_eq!(revision.author, "Alice");
    assert_eq!(revision.revision_id, Some(42));
    assert_eq!(revision.reason.unwrap().raw(), 42);
    assert_eq!(revision.revision_save_id, Some(0x11223344));
    assert_eq!(revision.timestamp.unwrap().year, 2026);
    assert!(run.deletion_revision().is_none());
    assert!(run.formatting_revision().is_none());

    let mut formatted = Run::new(
        "formatted".to_string(),
        CharacterProperties {
            has_formatting_revision: Some(true),
            formatting_revision_author_index: Some(1),
            formatting_revision_timestamp: Some(timestamp),
            formatting_revision_save_id: Some(0x55667788),
            ..CharacterProperties::default()
        },
    );
    formatted.resolve_revisions(&authors).unwrap();
    let revision = formatted.formatting_revision().unwrap();
    assert_eq!(revision.kind, RevisionKind::Formatting);
    assert_eq!(revision.author, "Alice");
    assert_eq!(revision.timestamp.unwrap().year, 2026);
    assert_eq!(revision.revision_id, None);
    assert_eq!(revision.reason, None);
    assert_eq!(revision.revision_save_id, Some(0x55667788));

    let mut bad_author = Run::new(
        "changed".to_string(),
        CharacterProperties {
            is_revision_inserted: Some(true),
            revision_author_index: Some(2),
            ..CharacterProperties::default()
        },
    );
    assert!(bad_author.resolve_revisions(&authors).is_err());

    let mut bad_time = Run::new(
        "changed".to_string(),
        CharacterProperties {
            is_revision_inserted: Some(true),
            revision_timestamp: Some(63),
            ..CharacterProperties::default()
        },
    );
    assert!(bad_time.resolve_revisions(&authors).is_err());
}

#[test]
fn resolves_table_row_revision_authors_and_timestamps() {
    let timestamp =
        30u32 | (14u32 << 6) | (15u32 << 11) | (7u32 << 16) | (126u32 << 20) | (3u32 << 29);
    let mut paragraph = Paragraph::with_properties(
        String::new(),
        crate::parts::pap::ParagraphProperties {
            has_table_formatting_revision: Some(true),
            table_formatting_revision_author_index: Some(1),
            table_formatting_revision_timestamp: Some(timestamp),
            ..Default::default()
        },
    );
    let authors = RevisionAuthorTable::from_authors(&["Unknown", "Table Editor"]);
    paragraph.resolve_revision(&authors).unwrap();
    let revision = paragraph.table_formatting_revision().unwrap();
    assert_eq!(revision.kind, RevisionKind::Formatting);
    assert_eq!(revision.author, "Table Editor");
    assert_eq!(revision.timestamp.unwrap().year, 2026);
}

#[test]
fn test_paragraph_text() {
    let para = Paragraph::new("Hello, World!".to_string());
    assert_eq!(para.text().unwrap(), "Hello, World!");
}

#[test]
fn test_run_text() {
    let run = Run::new("Test text".to_string(), CharacterProperties::default());
    assert_eq!(run.text().unwrap(), "Test text");
    assert_eq!(run.bold(), None);
    assert_eq!(run.italic(), None);
}

#[test]
#[allow(clippy::field_reassign_with_default)]
fn test_run_with_formatting() {
    let mut props = CharacterProperties::default();
    props.is_bold = Some(true);
    props.is_italic = Some(true);
    props.font_size = Some(24); // 12pt

    let run = Run::new("Formatted text".to_string(), props);
    assert!(run.bold().unwrap_or(false));
    assert!(run.italic().unwrap_or(false));
    assert_eq!(run.font_size(), Some(24));
}
