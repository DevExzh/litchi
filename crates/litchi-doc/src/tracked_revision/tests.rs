//! Focused regression tests for the tracked-revision semantic layer.

use super::{Error, Limits, RevisionEditor, RevisionKind, RevisionMetadata, Snapshot};
use crate::writer::{CharacterFormatting, ParagraphFormatting, TextRevision, Writer};
use std::io::Cursor;

fn base_doc() -> Vec<u8> {
    let mut writer = Writer::new();
    writer
        .add_paragraph_runs(
            vec![
                ("kept ".to_string(), CharacterFormatting::default()),
                (
                    "old".to_string(),
                    CharacterFormatting {
                        deletion_revision: Some(
                            TextRevision::new("Existing").with_revision_save_id(7),
                        ),
                        ..Default::default()
                    },
                ),
                (" tail".to_string(), CharacterFormatting::default()),
            ],
            ParagraphFormatting::default(),
        )
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn metadata_builder_is_composable() {
    let metadata = RevisionMetadata::new("Alice")
        .with_reason(0x2b)
        .with_revision_save_id(42);

    assert_eq!(metadata.author, "Alice");
    assert_eq!(metadata.reason, Some(0x2b));
    assert_eq!(metadata.revision_save_id, Some(42));
    assert_eq!(metadata.timestamp, None);
}

#[test]
fn revision_kinds_are_copyable_and_distinct() {
    let insertion = RevisionKind::Insertion;
    let deletion = insertion;

    assert_eq!(insertion, deletion);
    assert_ne!(insertion, RevisionKind::Deletion);
}

#[test]
fn transaction_supports_revision_metadata_crud_and_replacement() {
    let source = Snapshot::parse(&base_doc()).unwrap();
    let original = source.revisions().unwrap();
    let index = original
        .iter()
        .position(|revision| revision.author == "Existing")
        .unwrap();

    let mut transaction = source.edit().unwrap();
    let replacement = transaction
        .replace_metadata(
            index,
            RevisionMetadata::new("Replacement")
                .with_reason(0x2b)
                .with_revision_save_id(17),
        )
        .unwrap();
    assert_eq!(replacement.author, "Replacement");
    assert_eq!(replacement.reason, Some(0x2b));
    assert_eq!(replacement.revision_save_id, Some(17));

    let added = transaction
        .add(
            0,
            4,
            RevisionKind::Insertion,
            RevisionMetadata::new("Added"),
        )
        .unwrap();
    assert_eq!((added.start_cp, added.end_cp), (0, 4));
    let added_index = transaction
        .revisions()
        .unwrap()
        .iter()
        .position(|revision| revision.author == "Added")
        .unwrap();
    assert_eq!(transaction.remove(added_index).unwrap().author, "Added");

    let committed = transaction.commit().unwrap();
    assert!(committed.changed());
    let revisions = committed.snapshot().revisions().unwrap();
    assert!(revisions.iter().any(|revision| {
        revision.author == "Replacement"
            && revision.reason == Some(0x2b)
            && revision.revision_save_id == Some(17)
    }));
    assert!(!revisions.iter().any(|revision| revision.author == "Added"));
}

#[test]
fn failed_metadata_replacement_is_atomic_and_equal_replacement_is_a_noop() {
    let source = Snapshot::parse(&base_doc()).unwrap();
    let index = source
        .revisions()
        .unwrap()
        .iter()
        .position(|revision| revision.author == "Existing")
        .unwrap();
    let mut transaction = source.edit().unwrap();
    let before = transaction.revisions().unwrap();
    assert!(
        transaction
            .replace(index, RevisionMetadata::new("Invalid").with_reason(0x2c))
            .is_err()
    );
    assert_eq!(transaction.revisions().unwrap(), before);

    let current = before[index].clone();
    let mut metadata = RevisionMetadata::new(current.author.clone());
    if let Some(timestamp) = current.timestamp {
        metadata = metadata.with_timestamp(timestamp);
    }
    if let Some(reason) = current.reason {
        metadata = metadata.with_reason(reason);
    }
    if let Some(revision_save_id) = current.revision_save_id {
        metadata = metadata.with_revision_save_id(revision_save_id);
    }
    transaction.replace(index, metadata).unwrap();
    let committed = transaction.commit().unwrap();
    assert!(!committed.changed());
    assert!(committed.patch().is_noop());
    assert_eq!(committed.snapshot().bytes(), source.bytes());
}

#[test]
fn no_op_inverse_and_stale_source_checks_preserve_exact_bytes() {
    let source_bytes = base_doc();
    let source = Snapshot::parse(&source_bytes).unwrap();
    let mut transaction = source.edit().unwrap();
    let index = transaction
        .revisions()
        .unwrap()
        .iter()
        .position(|revision| revision.author == "Existing")
        .unwrap();
    transaction
        .replace(index, RevisionMetadata::new("Changed"))
        .unwrap();
    let committed = transaction.commit().unwrap();
    let applied = committed.patch().apply(&source).unwrap();
    assert_eq!(applied, *committed.snapshot());

    let reverted = committed.patch().inverse().apply(&applied).unwrap();
    assert_eq!(reverted.bytes(), source_bytes.as_slice());

    let mut other_editor = RevisionEditor::open(source_bytes, Limits::default()).unwrap();
    other_editor
        .add_text(
            0,
            "other",
            RevisionKind::Insertion,
            RevisionMetadata::new("Other"),
        )
        .unwrap();
    let other = Snapshot::parse(&other_editor.finish().unwrap()).unwrap();
    assert!(matches!(
        committed.patch().apply(&other),
        Err(Error::Conflict)
    ));
}

#[test]
fn replacement_retains_unmodeled_sprm_bytes() {
    let unknown = [0x01, 0x20, 0xa5];
    let mut group = unknown.to_vec();
    group.extend_from_slice(
        &super::codec::encode_revision(RevisionKind::Insertion, 1, &RevisionMetadata::new("Alice"))
            .unwrap(),
    );
    let replacement = super::codec::replace_revision_sprms(
        &group,
        RevisionKind::Insertion,
        Some((2, &RevisionMetadata::new("Bob"))),
    )
    .unwrap();
    assert_eq!(&replacement[..unknown.len()], &unknown);
}
