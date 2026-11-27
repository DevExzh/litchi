//! Semantic checks against unmodified native Microsoft Word binary files.

use litchi_core::Position;
use litchi_doc::Package;

fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}

#[test]
fn native_word_template_exposes_attached_autotext_with_end_sentinel() {
    let package = native("word-glossary.dot");
    let document = package.document();
    let glossary = document.attached_glossary().unwrap().unwrap();
    assert!(document.fib().is_template());
    assert!(!document.fib().is_glossary_document());
    assert!(glossary.fib().is_glossary_document());
    let metadata = glossary.metadata();
    assert_eq!(metadata.items().len(), 1);
    assert_eq!(metadata.items()[0].name(), "LITCHI_AUTOTEXT_014");
    assert_eq!(
        metadata.items()[0].kind(),
        litchi_doc::parts::glossary::GlossaryItemKind::NamedAutoText
    );
    assert_eq!(
        (metadata.items()[0].start_cp(), metadata.items()[0].end_cp()),
        (0, 57)
    );
    assert_eq!(
        (
            metadata.terminal_cp(),
            metadata.ignored_cp(),
            metadata.main_text_length()
        ),
        (57, 58, 58)
    );
    assert_eq!(metadata.styles().len(), 1);
    assert_eq!(metadata.styles()[0].name(), "Normal");
    assert_eq!(
        glossary.text(),
        "AUTOTEXT_BODY_014: reusable native Word building block.\r\r\r"
    );
    assert_eq!(
        glossary.item_text(0),
        Some("AUTOTEXT_BODY_014: reusable native Word building block.\r")
    );
}

#[test]
fn native_word_move_toggles_expose_both_ranges_and_shared_identity() {
    use litchi_doc::tracked_revision::RevisionKind;
    let package = native("word-tracked-moves.doc");
    let mut edit = package.edit_tracked_revisions().unwrap();
    let revisions = edit.revisions().unwrap();
    assert_eq!(revisions.len(), 2);
    assert_eq!(package.tracked_revisions().unwrap(), revisions);
    for (revision, (kind, start, end)) in revisions.iter().zip([
        (RevisionKind::MoveFrom, 58, 132),
        (RevisionKind::MoveTo, 193, 267),
    ]) {
        assert_eq!(revision.kind, kind);
        assert_eq!((revision.start_cp, revision.end_cp), (start, end));
        assert_eq!(revision.author_index, 1);
        assert_eq!(revision.author, "Ryker Zhu");
        assert_eq!(revision.revision_save_id, Some(0x00EC3828));
        assert_eq!(revision.move_pair_id, Some(0x00EC3828));
        let date = revision.timestamp.unwrap();
        assert_eq!(
            (date.year, date.month, date.day, date.hour, date.minute),
            (2026, 10, 15, 20, 7)
        );
    }
    assert!(
        package
            .document()
            .tolerance_report()
            .defects()
            .iter()
            .all(|defect| !matches!(
                defect.defect,
                litchi_doc::leniency::Defect::Chp(litchi_doc::leniency::ChpDefect::Boolean)
            ))
    );
}

#[test]
fn native_word_binary_variants_read_all_story_catalogs_and_preserve_exact_bytes() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/word");
    for (name, paragraphs, tables, marker) in [
        ("word-text-structure.doc", 25, 0, "WORD_NATIVE_END"),
        ("word-tables-layout.doc", 62, 1, "TABLE_END"),
        ("word-review-controls.doc", 12, 0, "REVIEW_CONTROLS_END"),
        (
            "word-recorded-macro.doc",
            5,
            0,
            "Native Word recorded macro fixture",
        ),
        (
            "word-recorded-macro.dot",
            5,
            0,
            "Native Word recorded macro fixture",
        ),
        ("word-math-navigation.doc", 21, 0, "MATH_NAVIGATION_END_005"),
        ("word-visual-objects.doc", 13, 0, "WORD_VISUAL_OBJECTS_007"),
        ("word-mailmerge-fields.doc", 9, 0, "END_MAILMERGE_009"),
        ("word-mailmerge-source.doc", 46, 1, "NATIVE_MERGE_009"),
        ("word-document-protection.doc", 7, 0, "WORD_PROTECTION_008"),
        ("word-legacy-form.doc", 6, 0, "WORD_LEGACY_FORM_010"),
        ("word-embedded-fonts.doc", 5, 0, "WORD_EMBEDDED_FONTS_011"),
        ("word-vector-diagram.doc", 10, 0, "WORD_VECTOR_DIAGRAM_013"),
        ("word-glossary.dot", 4, 0, "WORD_GLOSSARY_014"),
        ("word-bidi-eastasia.doc", 7, 0, "WORD_BIDI_EASTASIA_015"),
        (
            "word-layout-extensions.doc",
            27,
            1,
            "WORD_LAYOUT_EXTENSIONS_016",
        ),
        (
            "word-layout-extensions.dot",
            27,
            1,
            "WORD_LAYOUT_EXTENSIONS_016",
        ),
        ("word-master-native.doc", 3, 0, "WORD_MASTER_LEGACY_017"),
        (
            "word-master-native/LEGACY",
            3,
            0,
            "LEGACY_CHILD_SECTION_017",
        ),
        ("word-tracked-moves.doc", 7, 0, "WORD_TRACKED_MOVES_018"),
    ] {
        let source = std::fs::read(root.join(name)).unwrap();
        let package = Package::from_bytes(&source).unwrap();
        let document = package.document();
        assert_eq!(document.paragraph_count().unwrap(), paragraphs, "{name}");
        assert_eq!(document.paragraphs().unwrap().len(), paragraphs, "{name}");
        assert_eq!(document.table_count().unwrap(), tables, "{name}");
        assert_eq!(document.tables().unwrap().len(), tables, "{name}");
        assert!(document.text().unwrap().contains(marker), "{name}");
        assert!(
            package.to_bytes().unwrap() == source,
            "{name}: no-op changed bytes"
        );
    }
}

fn streams(bytes: &[u8]) -> std::collections::BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole = litchi_cfb::OleFile::open(
        std::io::Cursor::new(bytes),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let refs: Vec<_> = path.iter().map(String::as_str).collect();
            let bytes = ole.open_stream(&refs).unwrap();
            (path, bytes)
        })
        .collect()
}

#[test]
fn native_word_binary_title_edits_preserve_unrelated_streams_and_invert() {
    for name in [
        "word-text-structure.doc",
        "word-recorded-macro.doc",
        "word-recorded-macro.dot",
        "word-math-navigation.doc",
        "word-visual-objects.doc",
        "word-embedded-fonts.doc",
        "word-vector-diagram.doc",
        "word-glossary.dot",
        "word-bidi-eastasia.doc",
        "word-layout-extensions.doc",
        "word-layout-extensions.dot",
        "word-master-native.doc",
        "word-master-native/LEGACY",
    ] {
        let package = native(name);
        let before = package.to_bytes().unwrap();
        let before_streams = streams(&before);
        let mut edit = package.edit().unwrap_or_else(|e| panic!("{name}: {e:?}"));
        edit.replace_paragraph(Position::new(0), "NATIVE_BINARY_TITLE_EDIT")
            .unwrap_or_else(|e| panic!("{name}: {e:?}"));
        let commit = edit.commit().unwrap_or_else(|e| panic!("{name}: {e:?}"));
        let after = commit.package().to_bytes().unwrap();
        let reopened = Package::from_bytes(&after).unwrap();
        assert_eq!(
            reopened
                .document()
                .body_paragraph_text(Position::new(0))
                .unwrap(),
            Some("NATIVE_BINARY_TITLE_EDIT"),
            "{name}"
        );
        let after_streams = streams(&after);
        assert_eq!(
            before_streams.keys().collect::<Vec<_>>(),
            after_streams.keys().collect::<Vec<_>>(),
            "{name}"
        );
        for (path, bytes) in &before_streams {
            if !matches!(path.as_slice(),[name] if matches!(name.as_str(),"WordDocument"|"0Table"|"1Table"))
            {
                assert!(
                    after_streams[path] == *bytes,
                    "{name}: changed {}",
                    path.join("/")
                );
            }
        }
        assert_eq!(
            reopened.document().paragraph_count().unwrap(),
            package.document().paragraph_count().unwrap(),
            "{name}"
        );
        assert_eq!(
            reopened.document().table_count().unwrap(),
            package.document().table_count().unwrap(),
            "{name}"
        );
        if name == "word-glossary.dot" {
            assert_eq!(
                reopened
                    .document()
                    .attached_glossary()
                    .unwrap()
                    .unwrap()
                    .text(),
                package
                    .document()
                    .attached_glossary()
                    .unwrap()
                    .unwrap()
                    .text()
            );
        }
        let restored = reopened
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package();
        assert!(
            restored.to_bytes().unwrap() == before,
            "{name}: inverse changed source"
        );
        assert!(
            package.to_bytes().unwrap() == before,
            "{name}: source changed"
        );
    }
}

#[test]
fn native_word_binary_plain_body_create_move_and_delete_round_trip() {
    let source = native("word-bidi-eastasia.doc");
    let before = source.to_bytes().unwrap();
    let original = source
        .document()
        .body_paragraphs()
        .collect::<Result<Vec<_>, _>>()
        .unwrap()
        .into_iter()
        .map(str::to_owned)
        .collect::<Vec<_>>();
    let mut edit = source.edit().unwrap();
    edit.insert_paragraph(Position::new(1), "INSERTED_😀_中文")
        .unwrap();
    edit.move_paragraph(Position::new(1), Position::new(original.len()))
        .unwrap();
    let commit = edit.commit().unwrap();
    let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    let mut expected = original.clone();
    expected.push("INSERTED_😀_中文".into());
    assert_eq!(
        reopened
            .document()
            .body_paragraphs()
            .collect::<Result<Vec<_>, _>>()
            .unwrap(),
        expected
    );
    let mut remove = reopened.edit().unwrap();
    remove
        .remove_paragraph(Position::new(original.len()))
        .unwrap();
    let removed = remove.commit().unwrap();
    assert_eq!(
        removed
            .package()
            .document()
            .body_paragraphs()
            .collect::<Result<Vec<_>, _>>()
            .unwrap(),
        original
    );
    assert!(
        reopened
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
    assert!(source.to_bytes().unwrap() == before);
}

#[test]
fn native_word_binary_master_directory_and_extensionless_child_stay_inert() {
    use litchi_doc::parts::subdocuments::Kind;
    let parent = native("word-master-native.doc");
    let document = parent.document();
    let directory = document.subdocuments().unwrap();
    assert_eq!(directory.referenced_files().len(), 1);
    assert_eq!(directory.subdocuments().len(), 1);
    let reference = &directory.subdocuments()[0];
    assert_eq!(reference.start(), 59);
    assert_eq!(reference.outline_level(), 2);
    assert_eq!(reference.file_name_key().kind(), Kind::Subdocument);
    assert_eq!(reference.file_name_key().identifier(), 0);
    let name = directory.file_name_of(reference);
    assert_eq!(
        name.path(),
        "/Users/ryker/CodeProjects/litchi/test-data/microsoft-office/word/word-master-native/LEGACY"
    );
    assert_eq!(name.relative_path_offset(), Some(65));
    assert_eq!(name.relative_path(), Some("word-master-native/LEGACY"));
    assert_eq!(name.file_system_flags(), 0x0B);
    assert!(
        !document
            .text()
            .unwrap()
            .contains("LEGACY_CHILD_SECTION_017")
    );
    let child = native("word-master-native/LEGACY");
    assert!(child.document().subdocuments().is_none());
    assert_eq!(
        child
            .document()
            .body_paragraphs()
            .collect::<Result<Vec<_>, _>>()
            .unwrap(),
        [
            "LEGACY_CHILD_SECTION_017",
            "Native legacy subdocument content.",
            "\u{c}",
        ]
    );
}

#[test]
fn native_word_binary_password_contract_preserves_ciphertext_and_reservation() {
    use litchi_doc::{Error, OpenOptions, Password};
    let source = include_bytes!("../../../test-data/microsoft-office/word/word-encrypted.doc");
    assert!(matches!(
        Package::from_bytes(source),
        Err(Error::PasswordRequired)
    ));
    assert!(matches!(
        Package::from_bytes_with_options(
            source,
            OpenOptions::default().password(Password::new("wrong".into()))
        ),
        Err(Error::InvalidPassword)
    ));
    let package = Package::from_bytes_with_options(
        source,
        OpenOptions::default().password(Password::new("Office42!".into())),
    )
    .unwrap();
    let document = package.document();
    assert!(document.fib().is_encrypted());
    assert_eq!(
        document
            .body_paragraphs()
            .collect::<Result<Vec<_>, _>>()
            .unwrap(),
        [
            "WORD_ENCRYPTED_012",
            "Fixture ID: word-encrypted-012",
            "Native password-to-open specimen.",
            "VALUE_012: 42",
            "END_ENCRYPTED_012",
        ]
    );
    assert_eq!(
        document
            .associated_strings()
            .unwrap()
            .write_reservation_password(),
        "Office42!"
    );
    assert!(
        package.to_bytes().unwrap() == source,
        "decrypted projection must not replace retained ciphertext"
    );
    assert!(matches!(
        package.edit(),
        Err(litchi_doc::package::edit::Error::Body(
            litchi_doc::body_text::Error::Refused(litchi_doc::body_text::Refusal::Encrypted)
        ))
    ));
    assert!(package.to_bytes().unwrap() == source);
}

#[test]
fn native_word_binary_form_data_retains_defaults_and_undefined_stored_states() {
    use litchi_doc::{CheckBoxState, LegacyFormFieldKind};
    let package = native("word-legacy-form.doc");
    let fields = package.document().legacy_form_fields().unwrap();
    assert_eq!(fields.len(), 3);
    assert_eq!(
        fields.iter().map(|field| field.kind()).collect::<Vec<_>>(),
        [
            LegacyFormFieldKind::Text,
            LegacyFormFieldKind::CheckBox,
            LegacyFormFieldKind::DropDown
        ]
    );
    let text = fields[0].form_data().unwrap();
    assert_eq!(text.name(), "FixtureText010");
    assert_eq!(text.default_text(), Some("FORM_VALUE_010"));
    assert_eq!(fields[0].cached_result(), Some("FORM_VALUE_010"));
    assert_eq!(text.max_length(), Some(24));
    assert_eq!(text.help_text(), "HELP_TEXT_010");
    let checkbox = fields[1].form_data().unwrap();
    assert_eq!(checkbox.name(), "FixtureCheck010");
    assert_eq!(checkbox.checkbox_state(), Some(CheckBoxState::Undefined));
    assert_eq!(checkbox.is_checked_by_default(), Some(true));
    assert_eq!(checkbox.checkbox_size_half_points(), Some(28));
    let dropdown = fields[2].form_data().unwrap();
    assert_eq!(dropdown.name(), "FixtureDropdown010");
    assert_eq!(dropdown.default_item_index(), Some(0));
    assert_eq!(dropdown.selected_item_index(), None);
    assert_eq!(dropdown.dropdown_items(), ["OPTION_A_010", "OPTION_B_010"]);
    for field in &fields {
        let data = field.form_data().unwrap();
        assert!(data.entry_macro().is_empty());
        assert!(data.exit_macro().is_empty());
    }
}

#[test]
fn native_word_mail_merge_retains_original_source_path_and_cached_records() {
    use litchi_doc::parts::mail_merge::{MailMergeDestination, MailMergeType, MergeDataSourceKind};
    let package = native("word-mailmerge-fields.doc");
    let document = package.document();
    let state = document.mail_merge_state().unwrap();
    assert!(state.state.main_document && state.state.data_source);
    assert!(!state.state.header_file && !state.state.is_automatic && !state.state.record_selection);
    assert!(state.state.suppress_blank_lines);
    assert_eq!(state.state.merge_type, MailMergeType::None);
    assert_eq!(state.state.destination, MailMergeDestination::None);
    assert_eq!(state.current_record, Some(1));
    let source = state.sources[usize::from(state.fetch_source_index)];
    assert_eq!(source.source_kind, MergeDataSourceKind::DataFile);
    assert!(source.link_to_file);
    assert!(!source.uses_query);
    let file = document
        .subdocuments()
        .unwrap()
        .file_name(source.file_name)
        .unwrap();
    assert_eq!(
        file.kind(),
        litchi_doc::parts::subdocuments::Kind::MailMergeDataSource
    );
    // Word saved the DOCX path even in the binary DOC. Do not silently relink it.
    assert_eq!(
        file.path(),
        "/Users/ryker/CodeProjects/litchi/test-data/microsoft-office/word/word-mailmerge-source.docx"
    );
    assert_eq!(file.relative_path(), None);
    let fields = document.merge_fields().unwrap();
    assert_eq!(
        fields
            .iter()
            .map(|f| (f.field_name(), f.cached_result()))
            .collect::<Vec<_>>(),
        [
            ("FirstName", Some("«FirstName»")),
            ("LastName", Some("«LastName»")),
            ("Amount", Some("«Amount»")),
            ("FirstName", Some("Alex"))
        ]
    );
    let counters = document.mail_merge_counters().unwrap();
    assert_eq!(
        counters
            .iter()
            .map(|f| (f.instruction().trim(), f.cached_result()))
            .collect::<Vec<_>>(),
        [
            ("MERGEREC", Some("«Merge Record #»")),
            ("MERGESEQ", Some("«Merge Sequence #»"))
        ]
    );
    let reopened = Package::from_bytes(&package.to_bytes().unwrap()).unwrap();
    assert_eq!(reopened.document().mail_merge(), document.mail_merge());
    assert_eq!(reopened.document().subdocuments(), document.subdocuments());
    assert_eq!(reopened.document().merge_fields().unwrap(), fields);
}

#[test]
fn native_word_binary_review_comments_retain_reply_parent_and_local_timestamps() {
    let package = native("word-review-controls.doc");
    let comments = package.document().comments().unwrap();
    assert_eq!(comments.len(), 2);
    for (comment, (text, parent, depth, minute)) in comments.iter().zip([
        ("MODERN_COMMENT_003: review fixture comment.\r", None, 0, 50),
        ("MODERN_REPLY_003: native reply retained.\r", Some(0), 1, 51),
    ]) {
        assert_eq!(comment.text(), text);
        assert_eq!(comment.author, "Ryker Zhu");
        let metadata = comment.extended_metadata.unwrap();
        assert_eq!((metadata.parent_index, metadata.depth), (parent, depth));
        let timestamp = metadata.modified_at.unwrap();
        assert_eq!(
            (
                timestamp.year,
                timestamp.month,
                timestamp.day,
                timestamp.hour,
                timestamp.minute
            ),
            (2026, 10, 15, 13, minute)
        );
        assert!(!metadata.is_ink);
    }
}

#[test]
fn native_word_move_accept_and_reject_preserve_the_source_and_are_reversible() {
    let source = native("word-tracked-moves.doc");
    let before = source.to_bytes().unwrap();
    let moved = "SOURCE_018: Move this complete native paragraph to the destination below.";
    for accept in [true, false] {
        let mut edit = source.edit_tracked_revisions().unwrap();
        for _ in 0..2 {
            if accept {
                edit.accept(0).unwrap();
            } else {
                edit.reject(0).unwrap();
            }
        }
        assert!(edit.revisions().unwrap().is_empty());
        let commit = edit.commit().unwrap();
        let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
        let expected = if accept {
            vec![
                "WORD_TRACKED_MOVES_018",
                "Fixture ID: word-tracked-moves-018",
                "ANCHOR_018: This paragraph remains in place.",
                "DESTINATION_018",
                moved,
                "END_TRACKED_MOVES_018",
            ]
        } else {
            vec![
                "WORD_TRACKED_MOVES_018",
                "Fixture ID: word-tracked-moves-018",
                moved,
                "ANCHOR_018: This paragraph remains in place.",
                "DESTINATION_018",
                "END_TRACKED_MOVES_018",
            ]
        };
        assert_eq!(
            reopened
                .document()
                .body_paragraphs()
                .collect::<Result<Vec<_>, _>>()
                .unwrap(),
            expected
        );
        assert!(reopened.tracked_revisions().unwrap().is_empty());
        let restored = reopened
            .apply_tracked_revisions(&commit.patch().inverse())
            .unwrap();
        assert!(restored.package().to_bytes().unwrap() == before);
        assert!(source.to_bytes().unwrap() == before);
    }
}

#[test]
fn native_word_glossary_survives_deferred_character_formatting() {
    let source = native("word-glossary.dot");
    let document = source.document();
    let original = document.attached_glossary().unwrap().unwrap();
    let mut edit = source.edit().unwrap();
    edit.set_paragraph_bold(Position::new(0), true).unwrap();
    let commit = edit.commit().unwrap();
    let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    let document = reopened.document();
    let glossary = document.attached_glossary().unwrap().unwrap();
    assert_eq!(glossary.text(), original.text());
    assert_eq!(glossary.metadata(), original.metadata());
    let old_paragraphs = original.paragraphs().unwrap();
    let new_paragraphs = glossary.paragraphs().unwrap();
    assert_eq!(old_paragraphs.len(), new_paragraphs.len());
    for (before, after) in old_paragraphs.iter().zip(&new_paragraphs) {
        assert_eq!(before.text().unwrap(), after.text().unwrap());
        assert_eq!(
            before.properties().justification,
            after.properties().justification
        );
        let before_runs = before.runs().unwrap();
        let after_runs = after.runs().unwrap();
        assert_eq!(before_runs.len(), after_runs.len());
        for (before, after) in before_runs.iter().zip(&after_runs) {
            assert_eq!(before.text().unwrap(), after.text().unwrap());
            assert_eq!(before.bold(), after.bold());
            assert_eq!(before.font_size(), after.font_size());
        }
    }
    for index in [12, 13] {
        assert_eq!(
            document.fib().get_table_pointer(index),
            glossary.fib().get_table_pointer(index)
        );
    }
    assert_eq!(
        document.fib().word_document_size(),
        glossary.fib().word_document_size()
    );
}

#[test]
fn changed_publication_refuses_preexisting_broken_glossary_header_coupling() {
    let source = native("word-glossary.dot").to_bytes().unwrap();
    for (offset, reason) in [
        (
            4096 + 154 + 12 * 8,
            "source glossary formatting pointers are not shared",
        ),
        (4096 + 64, "invalid attached glossary FIB coupling"),
    ] {
        let mut members = streams(&source);
        let word = members.get_mut(&vec!["WordDocument".into()]).unwrap();
        let value = u32::from_le_bytes(word[offset..offset + 4].try_into().unwrap());
        word[offset..offset + 4].copy_from_slice(&(value + 1).to_le_bytes());
        let mut writer = litchi_cfb::OleWriter::new();
        for (path, data) in members {
            let parts: Vec<_> = path.iter().map(String::as_str).collect();
            writer.create_stream(&parts, &data).unwrap();
        }
        let mut output = std::io::Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        let bytes = output.into_inner();
        let package = Package::from_bytes(&bytes).unwrap();
        assert!(package.document().attached_glossary().is_err());
        let mut edit = package.edit().unwrap();
        let error = edit
            .replace_paragraph(Position::new(0), "must not publish")
            .unwrap_err();
        assert!(
            matches!(error,litchi_doc::package::edit::Error::Body(litchi_doc::body_text::Error::Invalid(litchi_doc::Error::Corrupted(message))) if message==reason)
        );
        assert!(package.to_bytes().unwrap() == bytes);
    }
}

fn native_font_block_hashes(package: &Package) -> Vec<(String, usize, String)> {
    use sha2::{Digest, Sha256};
    let document = package.document();
    let all = streams(&package.to_bytes().unwrap());
    let stream = &all[&vec!["WordDocument".to_owned()]];
    document
        .embedded_fonts()
        .unwrap()
        .fonts()
        .iter()
        .map(|font| {
            let start = usize::try_from(font.font_data_offset()).unwrap();
            let size = usize::try_from(u32::from_le_bytes(
                stream[start..start + 4].try_into().unwrap(),
            ))
            .unwrap();
            let block = &stream[start..start + size];
            let hash = Sha256::digest(block)
                .iter()
                .map(|v| format!("{v:02x}"))
                .collect();
            (
                document
                    .font_name(font.font_table_index())
                    .unwrap()
                    .unwrap()
                    .to_owned(),
                size,
                hash,
            )
        })
        .collect()
}

#[test]
fn native_binary_font_descriptions_resolve_exact_legacy_blocks() {
    let package = native("word-embedded-fonts.doc");
    let document = package.document();
    assert_eq!(
        document
            .font_table()
            .unwrap()
            .unwrap()
            .iter()
            .map(|f| f.name())
            .collect::<Vec<_>>(),
        [
            "Times New Roman",
            "Symbol",
            "Arial",
            "Aptos",
            "Aptos Display",
            "Cambria Math"
        ]
    );
    let fonts = document.embedded_fonts().unwrap().fonts();
    assert_eq!(fonts.len(), 2);
    assert_eq!(
        fonts
            .iter()
            .map(|f| (
                f.font_data_offset(),
                f.font_table_index(),
                f.is_bold(),
                f.is_italic(),
                f.subset_order()
            ))
            .collect::<Vec<_>>(),
        [
            (3630, 5, false, false, None),
            (573197, 3, false, false, None)
        ]
    );
    assert!(
        document
            .document_properties()
            .unwrap()
            .unwrap()
            .base()
            .embed_true_type_fonts()
    );
    assert_eq!(
        native_font_block_hashes(&package),
        [
            (
                "Cambria Math".into(),
                569567,
                "53e5b643b6f6382a558b814c3880067b2908ab5c4a4643beb627a1b3141de23c".into()
            ),
            (
                "Aptos".into(),
                99845,
                "a2bb161962578ec61efc3bd7d887bb9fe3883bf2306c50610b8c5df1ab035191".into()
            )
        ]
    );
}

#[test]
fn native_binary_body_edit_preserves_embedded_font_descriptions_and_blocks() {
    let package = native("word-embedded-fonts.doc");
    let before = package.to_bytes().unwrap();
    let hashes = native_font_block_hashes(&package);
    let mut edit = package.edit().unwrap();
    edit.replace_paragraph(Position::new(0), "FONT_METADATA_PRESERVED")
        .unwrap();
    let commit = edit.commit().unwrap();
    let reopened = Package::from_bytes(&commit.package().to_bytes().unwrap()).unwrap();
    assert_eq!(
        reopened
            .document()
            .body_paragraph_text(Position::new(0))
            .unwrap(),
        Some("FONT_METADATA_PRESERVED")
    );
    assert_eq!(native_font_block_hashes(&reopened), hashes);
    assert_eq!(
        reopened.document().embedded_fonts(),
        package.document().embedded_fonts()
    );
    assert!(
        reopened
            .document()
            .document_properties()
            .unwrap()
            .unwrap()
            .base()
            .embed_true_type_fonts()
    );
    assert!(package.to_bytes().unwrap() == before);
    assert!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
}
