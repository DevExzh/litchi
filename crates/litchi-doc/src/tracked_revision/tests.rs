//! Focused regression tests for the tracked-revision semantic layer.

use super::{Error, Limits, RevisionEditor, RevisionKind, RevisionMetadata, Snapshot};
use crate::package::Error as PackageError;
use crate::parts::fib::FileInformationBlock;
use crate::parts::protection::{ProtectionAuthorization, ProtectionPolicy};
use crate::writer::{CharacterFormatting, ParagraphFormatting, TextRevision, Writer};
use litchi_ole_common::object::{Editor as ObjectEditor, Targets};
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
    normalize_word2002_dop(output.into_inner())
}

fn normalize_word2002_dop(bytes: Vec<u8>) -> Vec<u8> {
    let mut package = ObjectEditor::open(bytes, Targets::default(), Limits::default()).unwrap();
    let word_path = ["WordDocument".to_string()];
    let word = package.stream(&word_path).unwrap();
    let fib = FileInformationBlock::parse(word).unwrap();
    let table_name = if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    };
    let table_path = [table_name.to_string()];
    let mut word = word.to_vec();
    let mut table = package.stream(&table_path).unwrap().to_vec();
    let (offset, length) = fib.get_table_pointer(31).unwrap();
    let offset = usize::try_from(offset).unwrap();
    let length = usize::try_from(length).unwrap();
    let dop = crate::parts::document_properties::DocumentProperties::writer_bytes(
        false, false, false, true,
    );
    if length < dop.len() {
        let insertion = offset + length;
        let extra = dop.len() - length;
        table.splice(insertion..insertion, std::iter::repeat_n(0, extra));
        let count = fib.table_pointer_count().unwrap();
        for index in 0..count {
            let pointer = 154 + index * 8;
            let current = usize::try_from(u32::from_le_bytes(
                word[pointer..pointer + 4].try_into().unwrap(),
            ))
            .unwrap();
            if current >= insertion {
                let shifted = u32::try_from(current + extra).unwrap();
                word[pointer..pointer + 4].copy_from_slice(&shifted.to_le_bytes());
            }
        }
    }
    table[offset..offset + dop.len()].copy_from_slice(&dop);
    let pointer = 154 + 31 * 8;
    word[pointer + 4..pointer + 8].copy_from_slice(&(dop.len() as u32).to_le_bytes());
    package.put_stream(&word_path, word).unwrap();
    package.put_stream(&table_path, table).unwrap();
    package.finish().unwrap()
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

#[test]
fn protected_revision_edit_requires_audited_policy_but_noop_is_readable() {
    let source_bytes = protected_doc();
    let snapshot = Snapshot::open(source_bytes.clone(), Limits::default()).unwrap();
    assert_eq!(snapshot.finish(), source_bytes);
    let mut transaction = snapshot.edit().unwrap();
    let error = transaction
        .add_text(
            0,
            "blocked",
            RevisionKind::Insertion,
            RevisionMetadata::new("Alice"),
        )
        .unwrap_err();
    assert!(matches!(
        error,
        Error::Invalid(PackageError::ProtectionDenied(_))
    ));

    let editor = RevisionEditor::open(source_bytes.clone(), Limits::default()).unwrap();
    assert_eq!(editor.finish().unwrap(), source_bytes);

    let mut editor = RevisionEditor::open(source_bytes.clone(), Limits::default()).unwrap();
    let error = editor
        .add_text(
            0,
            "blocked",
            RevisionKind::Insertion,
            RevisionMetadata::new("Alice"),
        )
        .unwrap_err();
    assert!(matches!(error, PackageError::ProtectionDenied(_)));

    let authorization =
        ProtectionAuthorization::audited("alice", "approved revision repair").unwrap();
    let policy = ProtectionPolicy::allow_protected(authorization);
    let mut editor =
        RevisionEditor::open_with_policy(source_bytes, Limits::default(), policy).unwrap();
    editor
        .add_text(
            0,
            "allowed",
            RevisionKind::Insertion,
            RevisionMetadata::new("Alice"),
        )
        .unwrap();
    assert_ne!(editor.finish().unwrap(), protected_doc());

    let authorization =
        ProtectionAuthorization::audited("alice", "approved snapshot revision repair").unwrap();
    let policy = ProtectionPolicy::allow_protected(authorization);
    let snapshot = Snapshot::open_with_policy(protected_doc(), Limits::default(), policy).unwrap();
    let mut transaction = snapshot.edit().unwrap();
    transaction
        .add_text(
            0,
            "allowed",
            RevisionKind::Insertion,
            RevisionMetadata::new("Alice"),
        )
        .unwrap();
    let committed = transaction.commit().unwrap();
    assert_ne!(committed.snapshot().bytes(), protected_doc().as_slice());
}

#[test]
fn protected_patch_cannot_cross_into_an_enforcing_destination() {
    let authorization =
        ProtectionAuthorization::audited("alice", "approved revision patch").unwrap();
    let allowed = Snapshot::open_with_policy(
        protected_doc(),
        Limits::default(),
        ProtectionPolicy::allow_protected(authorization),
    )
    .unwrap();
    let enforcing = Snapshot::open(allowed.bytes().to_vec(), Limits::default()).unwrap();

    let mut transaction = allowed.edit().unwrap();
    transaction
        .add_text(
            0,
            "allowed",
            RevisionKind::Insertion,
            RevisionMetadata::new("Alice"),
        )
        .unwrap();
    let commit = transaction.commit().unwrap();
    assert!(matches!(
        commit.patch().apply(&enforcing),
        Err(Error::Invalid(PackageError::ProtectionDenied(_)))
    ));

    let changed_enforcing =
        Snapshot::open(commit.snapshot().bytes().to_vec(), Limits::default()).unwrap();
    assert!(matches!(
        commit.patch().inverse().apply(&changed_enforcing),
        Err(Error::Invalid(PackageError::ProtectionDenied(_)))
    ));
}

fn protected_doc() -> Vec<u8> {
    let mut package =
        ObjectEditor::open(base_doc(), Targets::default(), Limits::default()).unwrap();
    let word_path = ["WordDocument".to_string()];
    let table_name = {
        let word = package.stream(&word_path).unwrap();
        let fib = FileInformationBlock::parse(word).unwrap();
        if fib.which_table_stream() {
            "1Table"
        } else {
            "0Table"
        }
        .to_owned()
    };
    let table_path = [table_name];
    let mut word = package.stream(&word_path).unwrap().to_vec();
    let mut table = package.stream(&table_path).unwrap().to_vec();
    let offset = u32::try_from(table.len()).unwrap();
    let mut dop = crate::parts::document_properties::DocumentProperties::writer_bytes(
        false, false, false, true,
    );
    dop[6] = 0x10;
    table.extend_from_slice(&dop);
    let pointer = 154 + 31 * 8;
    word[pointer..pointer + 4].copy_from_slice(&offset.to_le_bytes());
    word[pointer + 4..pointer + 8].copy_from_slice(&(dop.len() as u32).to_le_bytes());
    package.put_stream(&word_path, word).unwrap();
    package.put_stream(&table_path, table).unwrap();
    package.finish().unwrap()
}

/// Replaces a checked-in fixture's `Selsf` with a selection of the first
/// main-story character `unit`, flagged `flags`, and returns the new bytes and
/// that character's CP; `None` when the fixture has no such character or no
/// `Selsf`.
fn fixture_with_object_selection(name: &str, unit: u16, flags: u16) -> Option<(Vec<u8>, u32)> {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc")
        .join(name);
    let bytes = std::fs::read(path).unwrap();
    let editor = RevisionEditor::open(bytes.clone(), Limits::default()).unwrap();
    let text = editor.main_story_text().unwrap();
    let cp = u32::try_from(text.encode_utf16().position(|value| value == unit)?).unwrap();
    let mut package = ObjectEditor::open(bytes, Targets::default(), Limits::default()).unwrap();
    let word_path = ["WordDocument".to_string()];
    let fib = FileInformationBlock::parse(package.stream(&word_path).unwrap()).unwrap();
    let table_path = [if fib.which_table_stream() {
        "1Table"
    } else {
        "0Table"
    }
    .to_string()];
    let (offset, length) = fib.get_table_pointer(30)?;
    if length != 36 {
        return None;
    }
    let offset = usize::try_from(offset).unwrap();
    let mut table = package.stream(&table_path).unwrap().to_vec();
    let selsf = &mut table[offset..offset + 36];
    selsf.fill(0);
    selsf[0..2].copy_from_slice(&flags.to_le_bytes());
    selsf[2] = 1;
    selsf[4..8].copy_from_slice(&cp.to_le_bytes());
    selsf[8..12].copy_from_slice(&(cp + 1).to_le_bytes());
    selsf[20..24].copy_from_slice(&cp.to_le_bytes());
    selsf[24..26].copy_from_slice(&1u16.to_le_bytes());
    package.put_stream(&table_path, table).unwrap();
    Some((package.finish().unwrap(), cp))
}

/// Change 0768 review: text inserted at a selected inline picture (its
/// 0x0001 character) or floating shape (its 0x0008 anchor) is written in
/// front of the object, so the saved selection moves with the object and
/// covers none of the new text; rejecting the insertion restores the record.
#[test]
fn tracked_insertion_at_a_selected_object_moves_the_selection_with_the_object() {
    let fixtures = [
        "testPictures.doc",
        "FloatingPictures.doc",
        "picture.doc",
        "PngPicture.doc",
        "pictures_escher.doc",
        "image-comment-at-char.doc",
        "watermark.doc",
    ];
    for (unit, flags) in [(0x0001u16, 1u16 << 12), (0x0008, 1 << 8)] {
        let mut tested = 0;
        for name in fixtures {
            let Some((source, cp)) = fixture_with_object_selection(name, unit, flags) else {
                continue;
            };
            let mut editor = RevisionEditor::open(source, Limits::default()).unwrap();
            let recorded = editor.saved_selection().unwrap().unwrap();
            editor
                .add_text(
                    cp,
                    "0768",
                    RevisionKind::Insertion,
                    RevisionMetadata::new("0768"),
                )
                .unwrap_or_else(|error| panic!("{name}: {error}"));
            let moved = editor.saved_selection().unwrap().unwrap();
            assert_eq!(
                (moved.cp_first(), moved.cp_lim(), moved.cp_anchor()),
                (cp + 4, cp + 5, cp + 4),
                "{name} {unit:#06x}"
            );
            let text = editor
                .main_story_text()
                .unwrap()
                .encode_utf16()
                .collect::<Vec<_>>();
            assert_eq!(text[cp as usize + 4], unit, "{name} {unit:#06x}");

            let output = editor.clone().finish().unwrap();
            let mut reopened = RevisionEditor::open(output, Limits::default()).unwrap();
            let index = reopened
                .revisions()
                .unwrap()
                .iter()
                .position(|revision| revision.author == "0768")
                .unwrap();
            reopened.reject(index).unwrap();
            assert_eq!(
                reopened.saved_selection().unwrap().unwrap().bytes(),
                recorded.bytes(),
                "{name} {unit:#06x}"
            );
            tested += 1;
        }
        assert!(tested > 0, "no fixture has a main-story {unit:#06x}");
    }
}
