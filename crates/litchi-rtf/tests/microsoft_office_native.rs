//! Unmodified Microsoft Word 16.110.2 specimens. Provenance and expectations
//! are in test-data/microsoft-office/manifest.json and word/*-inspection.json.

use litchi_core::Position;
use litchi_rtf::{Package, edit::RunTextReplacement};
use sha2::{Digest, Sha256};

fn native(name: &str) -> Package {
    let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/word")
        .join(name);
    Package::open(path).unwrap_or_else(|error| panic!("{name}: {error}"))
}

fn serialize_model(source: &Package) -> Package {
    // Exercise the actual writer, rather than the unchanged package's exact
    // source-byte publication shortcut.
    let mut bytes = Vec::new();
    litchi_rtf::write::Writer::new(&mut bytes)
        .write_document(source.raw())
        .unwrap();
    Package::from_bytes(&bytes).unwrap()
}

const MINIMAL: &[u8] =
    include_bytes!("../../../test-data/microsoft-office/word/word-rtf-minimal.rtf");
const WRITE_RESERVATION: &[u8] =
    include_bytes!("../../../test-data/microsoft-office/word/word-write-reservation.rtf");
const MINIMAL_TEXT: &str = "RTF_MINIMAL_006: plain native Microsoft Word paragraph.";

#[test]
fn native_word_empty_datastore_preserves_body_and_round_trips() {
    let source = Package::from_bytes(MINIMAL).unwrap();
    assert_eq!(source.text().trim(), MINIMAL_TEXT);
    assert_eq!(source.paragraph_count(), 1);
    assert!(source.raw().data_store().unwrap().data.is_empty());

    let bytes = source.to_bytes().unwrap();
    let reopened = Package::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.text(), source.text());
    assert_eq!(reopened.raw().data_store(), source.raw().data_store());
    let written = serialize_model(&source);
    assert_eq!(written.text(), source.text());
    assert_eq!(written.raw().data_store(), source.raw().data_store());
}

#[test]
fn native_word_write_reservation_keeps_exact_hash_and_five_paragraphs() {
    let source = Package::from_bytes(WRITE_RESERVATION).unwrap();
    assert_eq!(source.paragraph_count(), 5);
    for marker in [
        "WORD_ENCRYPTED_012",
        "Fixture ID: word-encrypted-012",
        "Native password-to-open specimen.",
        "VALUE_012: 42",
        "END_ENCRYPTED_012",
    ] {
        assert!(source.text().contains(marker), "missing {marker}");
    }
    let reservations = source.raw().write_reservations();
    assert!(reservations.legacy.is_none());
    let hash = &reservations.hash.as_ref().unwrap().data;
    assert_eq!(hash.len(), 120);
    assert_eq!(
        Sha256::digest(hash.as_ref())
            .iter()
            .map(|byte| format!("{byte:02x}"))
            .collect::<String>(),
        "4ae9142f45c0b136c67498e8c12403eee167b00e4eabfb8c0b9fbfd716195f57"
    );
    let reopened = Package::from_bytes(&source.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.text(), source.text());
    assert_eq!(reopened.raw().write_reservations(), reservations);
    assert_eq!(reopened.raw().data_store(), source.raw().data_store());
    let written = serialize_model(&source);
    assert_eq!(written.text(), source.text());
    assert_eq!(written.raw().write_reservations(), reservations);
}

#[test]
fn native_word_run_edit_preserves_empty_datastore_and_source_snapshot() {
    let source = Package::from_bytes(MINIMAL).unwrap();
    let mut edit = source.edit();
    edit.replace_body_run_texts(&[RunTextReplacement::new(
        Position::new(0),
        Position::new(0),
        "Edited native Word paragraph.",
    )])
    .unwrap();
    let changed = edit.commit().unwrap();
    let reopened = Package::from_bytes(&changed.package().to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.text().trim(), "Edited native Word paragraph.");
    assert_eq!(reopened.raw().data_store(), source.raw().data_store());
    assert_eq!(source.text().trim(), MINIMAL_TEXT);
    assert_eq!(source.to_bytes().unwrap(), MINIMAL);
}

#[test]
fn native_word_datastore_documents_retain_markers_through_model_writer() {
    for (name, marker) in [
        ("word-embedded-fonts.rtf", "END_FONT_011"),
        ("word-visual-objects.rtf", "CHART_007"),
        ("word-legacy-form.rtf", "FORM_VALUE_010"),
        ("word-rtf-minimal.rtf", "RTF_MINIMAL_006"),
        ("word-write-reservation.rtf", "END_ENCRYPTED_012"),
        ("word-document-protection.rtf", "END_PROTECTION_008"),
        ("word-review-controls.rtf", "REVIEW_CONTROLS_END"),
        ("word-bidi-eastasia.rtf", "END_BIDI_EASTASIA_015"),
    ] {
        let source = native(name);
        assert!(source.text().contains(marker), "{name}: missing {marker}");
        assert!(source.raw().data_store().unwrap().data.is_empty(), "{name}");
        let written = serialize_model(&source);
        assert_eq!(written.text(), source.text(), "{name}");
        assert_eq!(
            written.raw().data_store(),
            source.raw().data_store(),
            "{name}"
        );
        assert_eq!(written.tables().len(), source.tables().len(), "{name}");
        assert_eq!(written.fields().len(), source.fields().len(), "{name}");
    }
}

#[test]
fn native_word_structural_opaque_tables_preserve_source_and_refuse_canonical_loss() {
    // These two sources contain structurally anchored unknown controls. The
    // feature matrix permits exact snapshot/source-splice output, while the
    // canonical writer must refuse to silently discard their owner scopes.
    for (name, marker) in [
        ("word-rtf-table-grid.rtf", "RTF_TABLE_GRID_006"),
        ("word-layout-extensions.rtf", "END_LAYOUT_EXTENSIONS_016"),
    ] {
        let source = native(name);
        assert!(source.text().contains(marker), "{name}");
        assert_eq!(source.tables().len(), 1, "{name}");
        assert!(source.raw().data_store().unwrap().data.is_empty());
        let saved = source.to_bytes().unwrap();
        let reopened = Package::from_bytes(&saved).unwrap();
        assert_eq!(reopened.text(), source.text());
        assert_eq!(reopened.raw().data_store(), source.raw().data_store());
        let mut output = Vec::new();
        let error = litchi_rtf::write::Writer::new(&mut output)
            .write_document(source.raw())
            .unwrap_err();
        assert_eq!(error.kind(), std::io::ErrorKind::InvalidInput);
        assert_eq!(
            error.to_string(),
            "RTF canonical rewrite cannot preserve a structurally scoped opaque node"
        );
        assert!(output.is_empty(), "refusal must precede publication");
        assert_eq!(source.to_bytes().unwrap(), saved);
    }
}

#[test]
fn native_word_table_grid_exposes_all_four_cell_values() {
    let source = native("word-rtf-table-grid.rtf");
    let reopened = Package::from_bytes(&source.to_bytes().unwrap()).unwrap();
    for package in [&source, &reopened] {
        let tables = package.raw().tables();
        assert_eq!(tables.len(), 1);
        let rows = tables[0].rows();
        assert_eq!(rows.len(), 2);
        for (row, expected) in rows.iter().zip([["Key", "Value"], ["RTF_CELL_006", "42"]]) {
            assert_eq!(row.cells().len(), 2);
            for (cell, value) in row.cells().iter().zip(expected) {
                assert_eq!(cell.text().trim(), value);
            }
        }
    }
}

#[test]
fn native_word_table_cell_edit_preserves_every_unrelated_source_byte() {
    let source = native("word-rtf-table-grid.rtf");
    let before = source.to_bytes().unwrap();
    let needle = b"RTF_CELL_006";
    let offsets = before
        .windows(needle.len())
        .enumerate()
        .filter_map(|(index, bytes)| (bytes == needle).then_some(index))
        .collect::<Vec<_>>();
    assert_eq!(offsets.len(), 1);
    let mut expected = before[..offsets[0]].to_vec();
    expected.extend_from_slice(b"CELL_EDIT_NATIVE");
    expected.extend_from_slice(&before[offsets[0] + needle.len()..]);

    let mut edit = source.edit();
    edit.set_table_cell_text(
        litchi_rtf::TableCellPath::outer(0, 1, 0),
        "CELL_EDIT_NATIVE",
    )
    .unwrap();
    let changed = edit.commit().unwrap();
    let saved = changed.package().to_bytes().unwrap();
    assert_eq!(saved, expected);
    assert_eq!(source.to_bytes().unwrap(), before);
    let reopened = Package::from_bytes(&saved).unwrap();
    assert_eq!(
        reopened.raw().tables()[0].rows()[1].cells()[0]
            .text()
            .trim(),
        "CELL_EDIT_NATIVE"
    );
    assert_eq!(
        reopened.raw().tables()[0].rows()[1].cells()[1]
            .text()
            .trim(),
        "42"
    );
    assert_eq!(reopened.raw().data_store(), source.raw().data_store());
}

#[test]
fn native_word_review_retains_insert_delete_and_comment_content() {
    let source = native("word-review-controls.rtf");
    let written = serialize_model(&source);
    for package in [&source, &written] {
        let revisions = package.raw().revisions();
        assert_eq!(revisions.len(), 2);
        assert!(revisions.iter().any(|revision| {
            revision.revision_type == litchi_rtf::RevisionType::Insertion
                && revision.content.contains("REVIEW_INSERT_003")
        }));
        assert!(revisions.iter().any(|revision| {
            revision.revision_type == litchi_rtf::RevisionType::Deletion
                && revision.content.contains("REVIEW_DELETE_003")
        }));
        let annotations = package.raw().annotations();
        assert_eq!(annotations.len(), 2);
        for marker in ["MODERN_COMMENT_003", "MODERN_REPLY_003"] {
            assert!(
                annotations
                    .iter()
                    .any(|annotation| annotation.text.contains(marker))
            );
        }
    }
}

#[test]
fn native_word_bidi_and_east_asian_text_retains_unicode_scalars() {
    let source = native("word-bidi-eastasia.rtf");
    let written = serialize_model(&source);
    for package in [&source, &written] {
        for expected in [
            "ARABIC_015: مرحبا بالعالم — اختبار النص العربي ١٢٣٤٥",
            "HEBREW_015: שלום עולם — בדיקת טקסט 12345",
            "JAPANESE_015: 日本語の禁則処理、句読点。「試験」１２３４５",
            "CHINESE_015: 中文排版测试，标点符号。《测试》一二三四五",
        ] {
            assert!(package.text().contains(expected), "missing {expected}");
        }
    }
}
