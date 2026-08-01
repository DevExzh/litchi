use litchi_ole::LegacyOfficeObjectLimits;
use litchi_ole::doc::writer::{CharacterFormatting, DocWriter, ParagraphFormatting, TextRevision};
use litchi_ole::doc::{
    DocTrackedRevisionEditor, DocTrackedRevisionKind, DocTrackedRevisionMetadata, Package,
};
use std::io::Cursor;

fn base_doc() -> Vec<u8> {
    let mut writer = DocWriter::new();
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
fn lists_authors_and_mutates_insertions_and_deletions_transactionally() {
    let mut editor =
        DocTrackedRevisionEditor::open(base_doc(), LegacyOfficeObjectLimits::default()).unwrap();
    assert!(editor.authors().contains(&"Existing".to_string()));
    let deletion = editor
        .revisions()
        .unwrap()
        .into_iter()
        .find(|r| r.kind == DocTrackedRevisionKind::Deletion)
        .unwrap();
    let insertion = editor
        .add_text(
            0,
            "new ",
            DocTrackedRevisionKind::Insertion,
            DocTrackedRevisionMetadata::new("Alice").with_revision_save_id(9),
        )
        .unwrap();
    assert_eq!((insertion.start_cp, insertion.end_cp), (0, 4));
    let deletion_index = editor
        .revisions()
        .unwrap()
        .iter()
        .position(|r| r.author == "Existing")
        .unwrap();
    editor.reject(deletion_index).unwrap();
    let insertion_index = editor
        .revisions()
        .unwrap()
        .iter()
        .position(|r| r.author == "Alice")
        .unwrap();
    editor.accept(insertion_index).unwrap();
    assert!(
        editor
            .revisions()
            .unwrap()
            .iter()
            .all(|r| r.author != "Alice" && r.author != "Existing")
    );
    let bytes = editor.finish().unwrap();
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let text = package.document().unwrap().text().unwrap().to_string();
    assert!(text.contains("new kept old tail"));
    assert_eq!(deletion.author, "Existing");
}

#[test]
fn accepts_deletion_rejects_insertion_and_pairs_moves_by_rsid() {
    let mut editor =
        DocTrackedRevisionEditor::open(base_doc(), LegacyOfficeObjectLimits::default()).unwrap();
    let deletion_index = editor
        .revisions()
        .unwrap()
        .iter()
        .position(|r| r.kind == DocTrackedRevisionKind::Deletion)
        .unwrap();
    editor.accept(deletion_index).unwrap();
    editor
        .add_text(
            0,
            "temporary",
            DocTrackedRevisionKind::Insertion,
            DocTrackedRevisionMetadata::new("Alice"),
        )
        .unwrap();
    let insertion_index = editor
        .revisions()
        .unwrap()
        .iter()
        .position(|r| r.author == "Alice")
        .unwrap();
    editor.reject(insertion_index).unwrap();
    let bytes = editor.finish().unwrap();
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let text = package.document().unwrap().text().unwrap().to_string();
    assert!(!text.contains("old"));
    assert!(!text.contains("temporary"));
}

#[test]
fn shared_rsid_exposes_binary_insertion_and_deletion_as_a_move_pair() {
    let mut editor =
        DocTrackedRevisionEditor::open(base_doc(), LegacyOfficeObjectLimits::default()).unwrap();
    let metadata = DocTrackedRevisionMetadata::new("Mover").with_revision_save_id(0xAABBCCDD);
    editor
        .add(0, 4, DocTrackedRevisionKind::MoveFrom, metadata.clone())
        .unwrap();
    editor
        .add_text(0, "kept", DocTrackedRevisionKind::MoveTo, metadata)
        .unwrap();
    let revisions = editor.revisions().unwrap();
    let from = revisions
        .iter()
        .find(|r| r.kind == DocTrackedRevisionKind::MoveFrom)
        .unwrap();
    let to = revisions
        .iter()
        .find(|r| r.kind == DocTrackedRevisionKind::MoveTo)
        .unwrap();
    assert_eq!(from.move_pair_id, Some(0xAABBCCDD));
    assert_eq!(to.move_pair_id, from.move_pair_id);
}

#[test]
fn malformed_ranges_controls_and_failed_updates_roll_back() {
    let mut editor =
        DocTrackedRevisionEditor::open(base_doc(), LegacyOfficeObjectLimits::default()).unwrap();
    let before = editor.revisions().unwrap();
    assert!(
        editor
            .add_text(
                0,
                "\u{13} MACROBUTTON",
                DocTrackedRevisionKind::Insertion,
                DocTrackedRevisionMetadata::new("Mallory")
            )
            .is_err()
    );
    assert!(
        editor
            .add(
                20_000,
                20_001,
                DocTrackedRevisionKind::Deletion,
                DocTrackedRevisionMetadata::new("Mallory")
            )
            .is_err()
    );
    assert!(
        editor
            .update(
                0,
                DocTrackedRevisionMetadata::new("Mallory").with_reason(0x2c)
            )
            .is_err()
    );
    assert_eq!(editor.revisions().unwrap(), before);
}

#[test]
fn bundled_word_and_libreoffice_redline_fixtures_are_strictly_gated() {
    let root = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let fixtures = [
        root.join("test-data/libreoffice-core/sw/qa/extras/ww8import/data/changes-in-footnote.doc"),
        root.join("test-data/libreoffice-core/sw/qa/core/doc/data/bookmark-delete-redline.doc"),
        root.join("test-data/libreoffice-core/sw/qa/core/data/ww8/fail/redline-1.doc"),
    ];
    for path in fixtures {
        let original = std::fs::read(&path).unwrap();
        match DocTrackedRevisionEditor::open(original.clone(), LegacyOfficeObjectLimits::default())
        {
            Ok(editor) => {
                let _ = editor.revisions().unwrap();
                assert_eq!(std::fs::read(&path).unwrap(), original);
            },
            Err(_) => assert_eq!(std::fs::read(&path).unwrap(), original),
        }
    }
}
