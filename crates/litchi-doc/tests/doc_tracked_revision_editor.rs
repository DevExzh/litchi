#![allow(
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_wrap,
    clippy::let_underscore_must_use,
    clippy::manual_midpoint,
    clippy::map_unwrap_or,
    clippy::needless_pass_by_value,
    clippy::shadow_reuse,
    clippy::wildcard_enum_match_arm,
    clippy::bool_assert_comparison,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::decimal_bitwise_operands,
    clippy::default_trait_access,
    clippy::doc_markdown,
    clippy::expect_used,
    clippy::field_reassign_with_default,
    clippy::float_cmp,
    clippy::implicit_clone,
    clippy::items_after_statements,
    clippy::manual_let_else,
    clippy::manual_repeat_n,
    clippy::manual_string_new,
    clippy::match_wildcard_for_single_variants,
    clippy::needless_raw_string_hashes,
    clippy::redundant_closure_for_method_calls,
    clippy::shadow_unrelated,
    clippy::similar_names,
    clippy::uninlined_format_args,
    clippy::unreadable_literal,
    clippy::unwrap_used,
    reason = "integration-test fixtures favor explicit wire values and concise panic-driven assertions over production-style ergonomics"
)]

use litchi_cfb::OleFile;
use litchi_doc::Package;
use litchi_doc::tracked_revision::{Limits, RevisionEditor, RevisionKind, RevisionMetadata};
use litchi_doc::writer::{CharacterFormatting, ParagraphFormatting, TextRevision, Writer};
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

fn root_stream_start(bytes: &[u8], name: &str) -> (u32, u64) {
    let ole = OleFile::open(Cursor::new(bytes.to_vec())).expect("DOC CFB should open");
    let entry = ole
        .list_directory_entries(&[])
        .expect("DOC root directory should open")
        .into_iter()
        .find(|entry| entry.name == name)
        .expect("DOC stream should exist");
    (entry.start_sector, entry.size)
}

#[test]
fn lists_authors_and_mutates_insertions_and_deletions_transactionally() {
    let mut editor = RevisionEditor::open(base_doc(), Limits::default()).unwrap();
    assert!(editor.authors().contains(&"Existing".to_string()));
    let deletion = editor
        .revisions()
        .unwrap()
        .into_iter()
        .find(|r| r.kind == RevisionKind::Deletion)
        .unwrap();
    let insertion = editor
        .add_text(
            0,
            "new ",
            RevisionKind::Insertion,
            RevisionMetadata::new("Alice").with_revision_save_id(9),
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
    let mut editor = RevisionEditor::open(base_doc(), Limits::default()).unwrap();
    let deletion_index = editor
        .revisions()
        .unwrap()
        .iter()
        .position(|r| r.kind == RevisionKind::Deletion)
        .unwrap();
    editor.accept(deletion_index).unwrap();
    editor
        .add_text(
            0,
            "temporary",
            RevisionKind::Insertion,
            RevisionMetadata::new("Alice"),
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
    let mut editor = RevisionEditor::open(base_doc(), Limits::default()).unwrap();
    let metadata = RevisionMetadata::new("Mover").with_revision_save_id(0xAABBCCDD);
    editor
        .add(0, 4, RevisionKind::MoveFrom, metadata.clone())
        .unwrap();
    editor
        .add_text(0, "kept", RevisionKind::MoveTo, metadata)
        .unwrap();
    let revisions = editor.revisions().unwrap();
    let from = revisions
        .iter()
        .find(|r| r.kind == RevisionKind::MoveFrom)
        .unwrap();
    let to = revisions
        .iter()
        .find(|r| r.kind == RevisionKind::MoveTo)
        .unwrap();
    assert_eq!(from.move_pair_id, Some(0xAABBCCDD));
    assert_eq!(to.move_pair_id, from.move_pair_id);
}

#[test]
fn malformed_ranges_controls_and_failed_updates_roll_back() {
    let mut editor = RevisionEditor::open(base_doc(), Limits::default()).unwrap();
    let before = editor.revisions().unwrap();
    assert!(
        editor
            .add_text(
                0,
                "\u{13} MACROBUTTON",
                RevisionKind::Insertion,
                RevisionMetadata::new("Mallory")
            )
            .is_err()
    );
    assert!(
        editor
            .add(
                20_000,
                20_001,
                RevisionKind::Deletion,
                RevisionMetadata::new("Mallory")
            )
            .is_err()
    );
    assert!(
        editor
            .update(0, RevisionMetadata::new("Mallory").with_reason(0x2c))
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
        match RevisionEditor::open(original.clone(), Limits::default()) {
            Ok(editor) => {
                let _ = editor.revisions().unwrap();
                assert_eq!(std::fs::read(&path).unwrap(), original);
            },
            Err(_) => assert_eq!(std::fs::read(&path).unwrap(), original),
        }
    }
}

#[test]
fn ordinary_tracked_revision_save_reuses_the_opened_doc_layout() {
    let root = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let source_path = root.join("test-data/ole/doc/picture.doc");
    // LibreOffice's Word 2002 FIB (cswNew 0) and 610-byte DOP carry no
    // protection, so the producer fixture is edited as checked in.
    let source = std::fs::read(source_path).expect("DOC fixture should exist");
    let table_name = {
        let ole = OleFile::open(Cursor::new(source.clone())).expect("DOC CFB should open");
        ole.list_directory_entries(&[])
            .expect("DOC root directory should open")
            .into_iter()
            .find(|entry| entry.name == "0Table" || entry.name == "1Table")
            .expect("DOC table stream should exist")
            .name
            .clone()
    };
    let before = root_stream_start(&source, &table_name);

    let mut editor = RevisionEditor::open(source.clone(), Limits::default())
        .expect("DOC fixture should enter the tracked editor");
    editor
        .add_text(
            0,
            "inserted ",
            RevisionKind::Insertion,
            RevisionMetadata::new("0663"),
        )
        .expect("tracked edit should commit");
    let output = editor.finish().expect("tracked DOC save should finish");
    let after = root_stream_start(&output, &table_name);
    assert_eq!(after.0, before.0, "DOC table allocation moved during save");
    assert_ne!(output, source, "the tracked edit should publish a change");
}

fn ole_doc_fixture(name: &str) -> Vec<u8> {
    let path = std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/doc")
        .join(name);
    std::fs::read(path).expect("DOC fixture should exist")
}

/// Every readable checked-in DOC fixture. None carries document or range
/// protection (change 0768), so each accepts a tracked insertion at CP 0 as
/// checked in, including the 24 that the strict protection grammar refused and
/// the 11 whose saved selection is a caret at CP 0.
const READABLE_FIXTURES: [&str; 35] = [
    "3endnotes.doc",
    "DiffFirstPageHeadFoot.doc",
    "FancyFoot.doc",
    "FloatingPictures.doc",
    "HeaderFooterProblematic.doc",
    "HeaderFooterUnicode.doc",
    "Lists.doc",
    "NoHeadFoot.doc",
    "PngPicture.doc",
    "ThreeColFoot.doc",
    "ThreeColHead.doc",
    "ThreeColHeadFoot.doc",
    "cfb-truncated-final-sector.doc",
    "cjklist30.doc",
    "cjklist31.doc",
    "cjklist34.doc",
    "cjklist35.doc",
    "commented-table.doc",
    "documentProperties.doc",
    "duplicate-style-names.doc",
    "empty.doc",
    "endingnote.doc",
    "equation.doc",
    "first-header-footer.doc",
    "footnote.doc",
    "hyperlink.doc",
    "image-comment-at-char.doc",
    "inline-endnote-and-footnote.doc",
    "lists-margins.doc",
    "picture.doc",
    "pictures_escher.doc",
    "table-merged-cells.doc",
    "tdf71749_with_footnote.doc",
    "testPictures.doc",
    "watermark.doc",
];

#[test]
fn every_readable_fixture_accepts_a_tracked_insertion_at_cp_zero() {
    for name in READABLE_FIXTURES {
        let source = ole_doc_fixture(name);
        let mut editor = RevisionEditor::open(source.clone(), Limits::default())
            .unwrap_or_else(|error| panic!("{name} should open: {error}"));
        let inserted = editor
            .add_text(
                0,
                "0768 ",
                RevisionKind::Insertion,
                RevisionMetadata::new("0768"),
            )
            .unwrap_or_else(|error| panic!("{name} should accept the insertion: {error}"));
        assert_eq!((inserted.start_cp, inserted.end_cp), (0, 5), "{name}");
        let output = editor
            .finish()
            .unwrap_or_else(|error| panic!("{name} should publish: {error}"));
        assert_ne!(output, source, "{name}");

        let reopened = RevisionEditor::open(output.clone(), Limits::default())
            .unwrap_or_else(|error| panic!("{name} output should reopen: {error}"));
        assert!(
            reopened.revisions().unwrap().iter().any(|revision| {
                revision.kind == RevisionKind::Insertion
                    && (revision.start_cp, revision.end_cp) == (0, 5)
                    && revision.author == "0768"
            }),
            "{name}"
        );
        assert_eq!(
            reopened.finish().unwrap(),
            output,
            "{name}: unchanged republication"
        );
        // The public reader accepts the output exactly when it accepts the
        // source; its strict stylesheet rules are independent of the edit.
        match public_text(&source) {
            Some(_) => assert!(
                public_text(&output).is_some_and(|text| text.contains("0768 ")),
                "{name}"
            ),
            None => assert!(public_text(&output).is_none(), "{name}"),
        }
    }
}

fn public_text(bytes: &[u8]) -> Option<String> {
    let mut package = Package::from_reader(Cursor::new(bytes.to_vec())).ok()?;
    let document = package.document().ok()?;
    document.text().ok().map(|text| text.to_string())
}

#[test]
fn unreadable_fixtures_remain_refused_at_open() {
    for name in [
        "PasswordProtected.doc",
        "cfb-v3-uninitialized-size-high-word.doc",
        "word6-no-table-stream.doc",
    ] {
        assert!(
            RevisionEditor::open(ole_doc_fixture(name), Limits::default()).is_err(),
            "{name}"
        );
    }
}

fn saved_selection_cps(bytes: &[u8]) -> (u32, u32, u32) {
    let mut package = Package::from_reader(Cursor::new(bytes.to_vec())).unwrap();
    let document = package.document().unwrap();
    let selection = document.saved_selection().unwrap().unwrap();
    (
        selection.cp_first(),
        selection.cp_lim(),
        selection.cp_anchor(),
    )
}

/// A tracked insertion at the saved caret leaves the caret on its recorded CP,
/// before the new text; rejecting the insertion restores the recorded Selsf.
#[test]
fn tracked_insertion_at_the_saved_caret_keeps_the_caret_before_the_new_text() {
    for (name, recorded, inserted) in [
        ("cjklist30.doc", (0, 0, 0), (0, 0, 0)),
        ("empty.doc", (0, 0, 0), (0, 0, 0)),
        ("HeaderFooterUnicode.doc", (0, 0, 407), (0, 0, 412)),
        ("NoHeadFoot.doc", (179, 179, 179), (184, 184, 184)),
    ] {
        let source = ole_doc_fixture(name);
        assert_eq!(saved_selection_cps(&source), recorded, "{name}");
        let mut editor = RevisionEditor::open(source, Limits::default()).unwrap();
        editor
            .add_text(
                0,
                "0768 ",
                RevisionKind::Insertion,
                RevisionMetadata::new("0768"),
            )
            .unwrap();
        let output = editor.clone().finish().unwrap();
        assert_eq!(saved_selection_cps(&output), inserted, "{name}");

        let index = editor
            .revisions()
            .unwrap()
            .iter()
            .position(|revision| revision.author == "0768")
            .unwrap();
        editor.reject(index).unwrap();
        let rejected = editor.finish().unwrap();
        assert_eq!(saved_selection_cps(&rejected), recorded, "{name}");
    }
}
