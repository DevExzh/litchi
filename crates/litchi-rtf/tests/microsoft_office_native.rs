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

#[test]
fn native_word_mail_merge_preserves_empty_file_table_and_inert_source() {
    let source = native("word-mailmerge-fields.rtf");
    let merge = source.raw().mail_merge().unwrap();
    assert!(merge.link_to_query);
    let data_source = merge.data_source.as_deref().unwrap();
    assert!(data_source.ends_with("/word/word-mailmerge-source.docx"));
    assert_eq!(
        merge.query.as_deref(),
        Some(format!("SELECT * FROM {data_source}").as_str())
    );
    assert!(source.raw().file_table().unwrap().entries().is_empty());
    assert!(
        source
            .raw()
            .fields()
            .iter()
            .any(|field| field.result == "MATCH_009")
    );
    assert_eq!(
        merge.main_type,
        Some(litchi_rtf::MailMergeMainType::Letters)
    );
    assert_eq!(merge.input_type, Some(litchi_rtf::MailMergeInputType::File));
    assert_eq!(
        merge.destination,
        Some(litchi_rtf::MailMergeDestination::NewDocument)
    );
    assert_eq!(merge.current_record, Some(1));
    assert_eq!(merge.error_mode, Some(2));
    assert_eq!(source.raw().merge_field_count(), 3);
    assert!(
        source
            .raw()
            .fields()
            .iter()
            .any(|field| field.instruction.trim_start().starts_with("IF ")
                && field.instruction.contains("FirstName")),
        "{:?}",
        source
            .raw()
            .fields()
            .iter()
            .map(|field| (&field.instruction, &field.result))
            .collect::<Vec<_>>()
    );
    assert_eq!(source.raw().mail_merge_counter_count(), 2);
    let written = serialize_model(&source);
    assert_eq!(written.raw().file_table(), source.raw().file_table());
    assert_eq!(written.raw().mail_merge(), source.raw().mail_merge());
    assert_eq!(written.text(), source.text());
    assert_eq!(written.raw().merge_field_count(), 3);
    assert_eq!(written.raw().mail_merge_counter_count(), 2);
}

#[test]
fn native_word_table_style_retains_all_seven_conditional_definitions() {
    let source = native("word-tables-layout.rtf");
    let sheet = source.raw().stylesheet();
    let variants = sheet
        .styles()
        .iter()
        .filter(|style| style.style_type == litchi_rtf::StyleType::Table && style.id == 36)
        .collect::<Vec<_>>();
    assert_eq!(variants.len(), 7);
    assert!(
        variants
            .iter()
            .all(|style| style.name == "Grid Table 2 Accent 1")
    );
    let scopes = variants
        .iter()
        .map(|style| {
            let c = style.table_conditional;
            [
                c.first_row,
                c.last_row,
                c.first_column,
                c.last_column,
                c.band_horizontal_odd,
                c.band_vertical_odd,
            ]
        })
        .collect::<std::collections::HashSet<_>>();
    let mut expected = std::collections::HashSet::from([[false; 6]]);
    for index in 0..6 {
        let mut scope = [false; 6];
        scope[index] = true;
        expected.insert(scope);
    }
    assert_eq!(scopes, expected);
    assert_eq!(
        sheet
            .get_typed(litchi_rtf::StyleType::Table, 36)
            .unwrap()
            .based_on,
        Some(11)
    );
    let mut bytes = Vec::new();
    let mut writer = litchi_rtf::write::Writer::new(&mut bytes);
    writer.write_document_header().unwrap();
    writer.write_stylesheet(sheet).unwrap();
    writer.write_str("}").unwrap();
    let written = Package::from_bytes(&bytes).unwrap();
    assert_eq!(written.raw().stylesheet(), sheet);
    for marker in [
        "TABLE_SECTION_START",
        "LANDSCAPE_SECTION_START",
        "SECTION_END_MARKER",
    ] {
        assert!(source.text().contains(marker), "missing {marker}");
    }
}

#[test]
fn native_word_vector_group_retains_unstarred_result_and_source_bytes() {
    let source = native("word-vector-diagram.rtf");
    let groups = source
        .raw()
        .fields()
        .iter()
        .flat_map(|field| &field.shape_groups)
        .collect::<Vec<_>>();
    assert_eq!(groups.len(), 1);
    let group = groups[0];
    assert_eq!(group.name, "Graphic 3");
    assert!(!group.shapes().is_empty());
    let litchi_rtf::LegacyDrawingPrimitive::Group { children, .. } =
        &group.result.as_ref().unwrap().drawing.primitive
    else {
        panic!("expected native grouped legacy fallback")
    };
    assert_eq!(children.len(), 3);
    assert!(
        matches!(&children[1], litchi_rtf::LegacyDrawingPrimitive::Polyline {closed:true, points, ..} if points.len()==20)
    );
    for marker in [
        "WORD_VECTOR_DIAGRAM_013",
        "SVG_VECTOR_013",
        "CONVERTED_VECTOR_GROUP_013",
    ] {
        assert!(source.text().contains(marker), "missing {marker}");
    }
    let bytes = source.to_bytes().unwrap();
    assert_eq!(
        bytes.as_slice(),
        include_bytes!("../../../test-data/microsoft-office/word/word-vector-diagram.rtf")
    );
    let reopened = Package::from_bytes(&bytes).unwrap();
    let reopened_groups = reopened
        .raw()
        .fields()
        .iter()
        .flat_map(|field| &field.shape_groups)
        .collect::<Vec<_>>();
    assert_eq!(reopened_groups, groups);
}

#[test]
fn native_word_tracked_move_markers_coexist_with_insert_delete_revisions() {
    let source = native("word-tracked-moves.rtf");
    let moves = source.raw().move_bookmarks();
    assert_eq!(moves.len(), 2);
    assert_eq!(moves[0].kind, litchi_rtf::MoveBookmarkKind::From);
    assert_eq!(moves[1].kind, litchi_rtf::MoveBookmarkKind::To);
    for bookmark in moves {
        assert_eq!(bookmark.tag, "move242987239");
        assert_eq!(bookmark.author, 1);
        assert_eq!(bookmark.date, 0x87ea7d07);
    }
    let moved = "SOURCE_018: Move this complete native paragraph to the destination below.";
    assert!(moves[1].content.contains(moved));
    // Deleted text is exposed through the revision rather than through the
    // visible body range used by the Move From bookmark.
    let revisions = source.raw().revisions();
    assert!(revisions.iter().any(|revision| revision.revision_type
        == litchi_rtf::RevisionType::Deletion
        && revision.content.contains(moved)));
    assert!(revisions.iter().any(|revision| revision.revision_type
        == litchi_rtf::RevisionType::Insertion
        && revision.content.contains(moved)));
    let bytes = source.to_bytes().unwrap();
    assert_eq!(
        bytes.as_slice(),
        include_bytes!("../../../test-data/microsoft-office/word/word-tracked-moves.rtf")
    );
    let reopened = Package::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.raw().move_bookmarks(), moves);
    assert_eq!(reopened.raw().revisions(), revisions);
}

#[test]
fn native_word_math_retains_inline_and_display_equation_trees() {
    use litchi_rtf::{MathObject, MathStructureChild, MathStructureKind, MathZoneKind};
    fn math_text(objects: &[MathObject<'_>]) -> String {
        let mut text = String::new();
        for object in objects {
            match object {
                MathObject::Run(run) => text.push_str(&run.text),
                MathObject::Structure(structure) => {
                    for child in &structure.children {
                        match child {
                            MathStructureChild::Element(element) => {
                                text.push_str(&math_text(&element.content))
                            },
                            MathStructureChild::MatrixRow(row) => {
                                for cell in &row.cells {
                                    text.push_str(&math_text(&cell.content));
                                }
                            },
                        }
                    }
                },
            }
        }
        text
    }

    let source = native("word-math-navigation.rtf");
    let zones = source.raw().math_zones();
    assert_eq!(zones.len(), 2);
    assert_eq!(zones[0].kind, MathZoneKind::Inline);
    assert_eq!(zones[1].kind, MathZoneKind::Display);
    assert_eq!(math_text(&zones[0].content), "x2+12+9");
    assert_eq!(math_text(&zones[1].content), "k=14k=10");
    assert_eq!(zones[0].content.len(), 5);
    assert!(
        matches!(&zones[0].content[0], MathObject::Structure(s) if s.kind == MathStructureKind::ScriptSup)
    );
    assert!(
        matches!(&zones[0].content[2], MathObject::Structure(s) if s.kind == MathStructureKind::Fraction)
    );
    let MathObject::Structure(sum) = &zones[1].content[0] else {
        panic!("expected summation");
    };
    assert_eq!(sum.kind, MathStructureKind::Nary);
    let properties = &sum.properties.as_ref().unwrap().properties;
    assert!(properties.iter().any(
        |property| property.name == litchi_rtf::MathPropertyName::Char && property.value == "∑"
    ));
    assert!(properties.iter().any(|property| property.name
        == litchi_rtf::MathPropertyName::LimitLocation
        && property.value == "undOvr"));
    for marker in [
        "OMML_INLINE_005",
        "OMML_DISPLAY_005",
        "MATH_NAVIGATION_END_005",
    ] {
        assert!(source.text().contains(marker), "missing {marker}");
    }
    let bytes = source.to_bytes().unwrap();
    assert_eq!(
        bytes.as_slice(),
        include_bytes!("../../../test-data/microsoft-office/word/word-math-navigation.rtf")
    );
    assert_eq!(
        Package::from_bytes(&bytes).unwrap().raw().math_zones(),
        zones
    );

    // Write just the typed equations, independently of source-byte replay.
    // Cosmetic character formatting and picture fallbacks are outside this
    // model's existing contract; the original package keeps their exact bytes.
    let mut expected = zones.to_vec();
    let mut bytes = Vec::new();
    let mut writer = litchi_rtf::write::Writer::new(&mut bytes);
    writer.write_document_header().unwrap();
    for zone in &mut expected {
        zone.position = 0;
        writer.write_math_zone(zone).unwrap();
    }
    writer.write_str("}").unwrap();
    assert_eq!(
        Package::from_bytes(&bytes).unwrap().raw().math_zones(),
        expected
    );
}
