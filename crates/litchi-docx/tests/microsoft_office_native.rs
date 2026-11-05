//! Native Microsoft Word specimens, with bounded semantic and edit assertions.

use litchi_core::Position;
use litchi_docx::Package;

fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}

fn members(bytes: &[u8]) -> std::collections::BTreeMap<String, Vec<u8>> {
    let reader = litchi_opc::phys_pkg::OwnedPhysPkgReader::from_bytes(
        bytes.to_vec(),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    reader
        .member_names()
        .unwrap()
        .into_iter()
        .map(|name| {
            let data = reader.read_member(&name).unwrap();
            (name, data)
        })
        .collect()
}

fn assert_members_equal(left: &[u8], right: &[u8]) {
    let a = members(left);
    let b = members(right);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        assert!(b[&name] == bytes, "changed {name}");
    }
}

#[test]
fn native_word_variants_read_body_catalogs_and_preserve_every_package_member() {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/word");
    for (name, paragraphs, tables, marker) in [
        ("word-text-structure.docx", 12, 0, "WORD_NATIVE_END"),
        ("word-tables-layout.docx", 36, 1, "TABLE_END"),
        ("word-review-controls.docx", 9, 0, "REVIEW_CONTROLS_END"),
        (
            "word-recorded-macro.docm",
            5,
            0,
            "Native Word recorded macro fixture",
        ),
        (
            "word-recorded-macro.dotm",
            5,
            0,
            "Native Word recorded macro fixture",
        ),
        (
            "word-math-navigation.docx",
            21,
            0,
            "MATH_NAVIGATION_END_005",
        ),
        (
            "word-math-navigation.dotx",
            21,
            0,
            "MATH_NAVIGATION_END_005",
        ),
        ("word-visual-objects.docx", 9, 0, "WORD_VISUAL_OBJECTS_007"),
        ("word-mailmerge-fields.docx", 9, 0, "END_MAILMERGE_009"),
        ("word-mailmerge-source.docx", 43, 1, "NATIVE_MERGE_009"),
        ("word-document-protection.docx", 5, 0, "WORD_PROTECTION_008"),
        (
            "word-legacy-form-unprotected.docx",
            6,
            0,
            "WORD_LEGACY_FORM_010",
        ),
        ("word-legacy-form.docx", 6, 0, "WORD_LEGACY_FORM_010"),
        ("word-embedded-fonts.docx", 5, 0, "WORD_EMBEDDED_FONTS_011"),
        (
            "word-embedded-fonts-subset.docx",
            5,
            0,
            "WORD_EMBEDDED_FONTS_011",
        ),
        ("word-vector-diagram.docx", 7, 0, "WORD_VECTOR_DIAGRAM_013"),
        ("word-glossary.dotx", 4, 0, "WORD_GLOSSARY_014"),
        ("word-glossary.docx", 4, 0, "WORD_GLOSSARY_014"),
        ("word-bidi-eastasia.docx", 7, 0, "WORD_BIDI_EASTASIA_015"),
        (
            "word-layout-extensions.docx",
            14,
            1,
            "WORD_LAYOUT_EXTENSIONS_016",
        ),
        (
            "word-layout-extensions.dotx",
            14,
            1,
            "WORD_LAYOUT_EXTENSIONS_016",
        ),
        (
            "word-master-document.docx",
            4,
            0,
            "WORD_MASTER_DOCUMENT_017",
        ),
        ("word-master-document/CHILD", 4, 0, "CHILD_SECTION_017"),
        ("word-tracked-moves.docx", 7, 0, "WORD_TRACKED_MOVES_018"),
    ] {
        let source = std::fs::read(root.join(name)).unwrap();
        let package = Package::from_bytes(&source).unwrap();
        let doc = package.document().unwrap();
        assert_eq!(doc.paragraph_count().unwrap(), paragraphs, "{name}");
        assert_eq!(doc.table_count().unwrap(), tables, "{name}");
        assert!(doc.text().unwrap().contains(marker), "{name}: {marker}");
        for paragraph in doc.paragraphs().unwrap() {
            for run in paragraph.runs().unwrap() {
                let _ = run.text().unwrap();
            }
        }
        assert_members_equal(&package.to_bytes().unwrap(), &source);
    }
}

#[test]
fn native_word_inline_and_display_math_keep_distinct_semantic_boundaries() {
    for name in ["word-math-navigation.docx", "word-math-navigation.dotx"] {
        let package = native(name);
        let doc = package.document().unwrap();
        let mut inline = Vec::new();
        let mut display = Vec::new();
        for paragraph in doc.paragraphs().unwrap() {
            inline.extend(paragraph.inline_office_math().unwrap());
            display.extend(paragraph.display_office_math().unwrap());
        }
        assert_eq!(inline.len(), 1, "{name}: inline");
        assert_eq!(display.len(), 1, "{name}: display");
        for tag in ["<m:sSup>", "<m:f>", "<m:rad>"] {
            assert!(inline[0].xml().contains(tag), "{tag}");
        }
        assert!(display[0].xml().contains("<m:nary>"));
        assert!(display[0].xml().contains("m:val=\"∑\""));
        assert!(display[0].xml().contains("<m:t>k=1</m:t>"));
        assert!(display[0].xml().contains("<m:t>=10</m:t>"));
    }
}

#[test]
fn native_word_equations_reuse_the_word_paragraph_envelope() {
    let source = native("word-math-navigation.docx");
    let mut inline = Vec::new();
    let mut display = Vec::new();
    for paragraph in source.document().unwrap().paragraphs().unwrap() {
        inline.extend(paragraph.inline_office_math().unwrap());
        display.extend(paragraph.display_office_math().unwrap());
    }
    let mut builder = litchi_docx::Builder::new().unwrap();
    let paragraph = builder.document().unwrap().add_paragraph();
    paragraph.add_run_with_text("before ");
    paragraph.add_inline_office_math(inline[0].clone());
    paragraph.add_run_with_text(" after");
    paragraph.add_display_office_math(display[0].clone());
    let bytes = builder.build().unwrap().to_bytes().unwrap();
    let package = Package::from_bytes(&bytes).unwrap();
    let paragraph = package.document().unwrap().paragraphs().unwrap().remove(0);
    let inlines = paragraph.inlines().unwrap();
    assert_eq!(inlines.len(), 4);
    assert!(matches!(&inlines[0], litchi_docx::Inline::Run(_)));
    assert!(
        matches!(&inlines[1], litchi_docx::Inline::Unknown(value) if value.kind() == litchi_docx::InlineKind::Math)
    );
    assert!(matches!(&inlines[2], litchi_docx::Inline::Run(_)));
    assert!(
        matches!(&inlines[3], litchi_docx::Inline::Unknown(value) if value.kind() == litchi_docx::InlineKind::MathParagraph)
    );
    assert_eq!(paragraph.text().unwrap(), "before  after");
    assert_eq!(paragraph.omml_formulas().unwrap(), [inline[0].xml()]);
    assert_eq!(
        paragraph.paragraph_level_formulas().unwrap(),
        [display[0].xml()]
    );
}

#[test]
fn native_word_source_title_edits_preserve_other_members_and_invert_exactly() {
    for name in [
        "word-text-structure.docx",
        "word-review-controls.docx",
        "word-math-navigation.docx",
        "word-math-navigation.dotx",
        "word-recorded-macro.docm",
        "word-recorded-macro.dotm",
        "word-visual-objects.docx",
        "word-vector-diagram.docx",
        "word-embedded-fonts.docx",
        "word-embedded-fonts-subset.docx",
        "word-master-document.docx",
        "word-master-document/CHILD",
        "word-glossary.dotx",
    ] {
        let source = native(name);
        let before = source.to_bytes().unwrap();
        let before_members = members(&before);
        let mut edit = source.edit().unwrap();
        edit.replace_run_text(
            Position::new(0),
            Position::new(0),
            "LITCHI_NATIVE_TITLE_EDIT",
        )
        .unwrap();
        let commit = edit.commit().unwrap();
        let after = commit.package().to_bytes().unwrap();
        let reopened = Package::from_bytes(&after).unwrap();
        assert_eq!(
            reopened.document().unwrap().paragraphs().unwrap()[0]
                .text()
                .unwrap(),
            "LITCHI_NATIVE_TITLE_EDIT",
            "{name}"
        );
        let after_members = members(&after);
        assert_eq!(
            before_members.keys().collect::<Vec<_>>(),
            after_members.keys().collect::<Vec<_>>(),
            "{name}"
        );
        for (part, bytes) in &before_members {
            if part != "word/document.xml" {
                assert!(after_members[part] == *bytes, "{name}: changed {part}");
            }
        }
        assert!(
            source.to_bytes().unwrap() == before,
            "{name}: source mutated"
        );
        let inverse = commit.patch().inverse();
        let restored = reopened.apply(&inverse).unwrap().into_package();
        assert_members_equal(&restored.to_bytes().unwrap(), &before);
    }
}

#[test]
fn native_word_tables_sections_and_secondary_stories_have_exact_structure() {
    let package = native("word-tables-layout.docx");
    let document = package.document().unwrap();
    let table = document.table(0).unwrap().unwrap();
    assert_eq!(table.row_count().unwrap(), 5);
    assert_eq!(table.column_count().unwrap(), 4);
    assert_eq!(
        table
            .rows()
            .unwrap()
            .iter()
            .map(|row| row.cell_count().unwrap())
            .collect::<Vec<_>>(),
        [4, 4, 4, 4, 1]
    );
    for (row, expected) in [
        (0, ["Item", "Quantity", "Price", "Notes"]),
        (1, ["Apples", "3", "1.25", "Fruit"]),
        (2, ["Pears", "5", "2.50", "Unicode 中文"]),
    ] {
        for (column, text) in expected.iter().enumerate() {
            assert_eq!(
                table.cell(row, column).unwrap().unwrap().text().unwrap(),
                *text
            );
        }
    }
    let merged = table.cell(4, 0).unwrap().unwrap();
    assert_eq!(merged.grid_span().unwrap(), 4);
    assert!(merged.text().unwrap().contains("Merged summary"));
    assert!(merged.text().unwrap().contains("TABLE_END"));
    let nested = table
        .cell(3, 3)
        .unwrap()
        .unwrap()
        .blocks()
        .unwrap()
        .table(Position::new(0))
        .unwrap();
    assert_eq!(nested.row_count().unwrap(), 2);
    assert_eq!(nested.column_count().unwrap(), 2);
    assert_eq!(
        nested.cell(0, 0).unwrap().unwrap().v_merge().unwrap(),
        Some(litchi_docx::VMergeState::Restart)
    );
    assert_eq!(
        nested.cell(1, 0).unwrap().unwrap().v_merge().unwrap(),
        Some(litchi_docx::VMergeState::Continue)
    );
    assert_eq!(nested.cell(0, 1).unwrap().unwrap().text().unwrap(), "N2");
    assert_eq!(nested.cell(1, 1).unwrap().unwrap().text().unwrap(), "20");
    let mut sections = document.sections().unwrap();
    assert_eq!(sections.len(), 2);
    for (section, (width, height, orientation, columns)) in sections.iter_mut().zip([
        (12240, 15840, litchi_docx::section::Orientation::Portrait, 1),
        (
            15840,
            12240,
            litchi_docx::section::Orientation::Landscape,
            2,
        ),
    ]) {
        let page = section.page_size_checked().unwrap();
        assert_eq!(page.width.unwrap().to_twips(), width);
        assert_eq!(page.height.unwrap().to_twips(), height);
        assert_eq!(page.orientation, orientation);
        let margins = section.margins_checked().unwrap();
        for margin in [margins.top, margins.bottom, margins.left, margins.right] {
            assert_eq!(margin.unwrap().to_twips(), 1440);
        }
        let layout = section.columns().unwrap().unwrap();
        assert_eq!(layout.count, columns);
        assert_eq!(layout.space.unwrap().to_twips(), 720);
    }
    let headers = document.headers().unwrap();
    assert_eq!(headers.len(), 1);
    assert!(headers[0].text().unwrap().contains("HEADER_DEFAULT_002"));
    let footers = document.footers().unwrap();
    assert_eq!(footers.len(), 2);
    assert!(
        footers
            .iter()
            .any(|footer| footer.text().unwrap().contains("FOOTER_PAGE_002:"))
    );
    assert!(document.footnotes().unwrap().is_empty());
    let notes = document.endnotes().unwrap();
    assert_eq!(notes.len(), 1);
    assert_eq!(notes[0].id(), 1);
    assert_eq!(
        notes[0].text().unwrap(),
        " ENDNOTE_TEXT_002: native endnote value 84."
    );
}

#[test]
fn native_word_review_threads_and_revisions_keep_ids_authors_and_two_timestamps() {
    use litchi_docx::revision::RevisionType;
    let package = native("word-review-controls.docx");
    let document = package.document().unwrap();
    let comments = document.comments().unwrap();
    assert_eq!(comments.len(), 2);
    for (comment, (id, paragraph_id, text, date)) in comments.iter().zip([
        (
            5,
            0x44BDD156,
            "MODERN_COMMENT_003: review fixture comment.",
            "2026-10-15T13:50:00Z",
        ),
        (
            6,
            0x3FD135F2,
            "MODERN_REPLY_003: native reply retained.",
            "2026-10-15T13:51:00Z",
        ),
    ]) {
        assert_eq!(comment.id(), id);
        assert_eq!(comment.paragraph_id(), Some(paragraph_id));
        assert_eq!(comment.author(), "Ryker Zhu");
        assert_eq!(comment.date(), Some(date));
        assert_eq!(comment.text().unwrap(), text);
    }
    let threads = document.comments_extended().unwrap();
    assert_eq!(threads.len(), 2);
    assert_eq!(
        (
            threads[0].paragraph_id,
            threads[0].parent_paragraph_id,
            threads[0].done
        ),
        (0x44BDD156, None, false)
    );
    assert_eq!(
        (
            threads[1].paragraph_id,
            threads[1].parent_paragraph_id,
            threads[1].done
        ),
        (0x3FD135F2, Some(0x44BDD156), false)
    );
    let revisions = document.revisions().unwrap();
    assert_eq!(revisions.len(), 4);
    for (revision, (id, kind)) in revisions.iter().zip([
        ("0", RevisionType::FormatChange),
        ("1", RevisionType::FormatChange),
        ("2", RevisionType::Insert),
        ("3", RevisionType::Delete),
    ]) {
        assert_eq!(revision.id(), id);
        assert_eq!(revision.revision_type(), kind);
        assert_eq!(revision.author(), Some("Ryker Zhu"));
        assert_eq!(revision.date(), Some("2026-10-15T13:49:00Z"));
        assert_eq!(revision.date_utc(), Some("2026-10-15T05:49:00Z"));
    }
    assert_eq!(revisions[2].text(), "REVIEW_INSERT_003");
    assert_eq!(revisions[3].text(), "REVIEW_DELETE_003");
}

#[test]
fn native_word_tracked_moves_keep_paragraph_marks_and_content_separate() {
    use litchi_docx::revision::RevisionType;
    let package = native("word-tracked-moves.docx");
    let revisions = package.document().unwrap().revisions().unwrap();
    assert_eq!(revisions.len(), 4);
    let moved = "SOURCE_018: Move this complete native paragraph to the destination below.";
    for (revision, (id, kind, text)) in revisions.iter().zip([
        ("0", RevisionType::ParagraphMarkMoveFrom, ""),
        ("2", RevisionType::MoveFrom, moved),
        ("3", RevisionType::ParagraphMarkMoveTo, ""),
        ("5", RevisionType::MoveTo, moved),
    ]) {
        assert_eq!(revision.id(), id);
        assert_eq!(revision.revision_type(), kind);
        assert_eq!(revision.text(), text);
        assert_eq!(revision.author(), Some("Ryker Zhu"));
        assert_eq!(revision.date(), Some("2026-10-15T20:07:00Z"));
        assert_eq!(revision.date_utc(), Some("2026-10-15T12:07:00Z"));
    }
}

#[test]
fn native_word_protection_and_tracking_refuse_changed_publication_without_mutation() {
    use litchi_docx::settings::ProtectionType;
    for (name, protected, protection, tracked) in [
        (
            "word-document-protection.docx",
            true,
            Some(ProtectionType::Comments),
            false,
        ),
        (
            "word-legacy-form.docx",
            true,
            Some(ProtectionType::Forms),
            false,
        ),
        ("word-tracked-moves.docx", false, None, true),
    ] {
        let source = native(name);
        let before = source.to_bytes().unwrap();
        let settings = source.document().unwrap().settings().unwrap().unwrap();
        assert_eq!(settings.is_protected(), protected);
        assert_eq!(settings.protection_type(), protection);
        assert_eq!(settings.track_revisions(), tracked);
        let noop = source.edit().unwrap().commit().unwrap();
        assert_members_equal(&noop.package().to_bytes().unwrap(), &before);
        let mut edit = source.edit().unwrap();
        edit.replace_run_text(Position::new(0), Position::new(0), "refused title")
            .unwrap();
        assert!(
            matches!(edit.commit(),Err(litchi_docx::document::Error::Document(litchi_docx::Error::Refused(refusal))) if refusal.kind() == litchi_core::edit::RefusalKind::Protected),
            "{name}"
        );
        assert!(
            source.to_bytes().unwrap() == before,
            "{name}: source mutated"
        );
    }
    let unprotected = native("word-legacy-form-unprotected.docx");
    let settings = unprotected.document().unwrap().settings().unwrap().unwrap();
    assert!(!settings.is_protected());
    assert_eq!(settings.protection_type(), None);
    let mut edit = unprotected.edit().unwrap();
    edit.replace_run_text(Position::new(0), Position::new(0), "unprotected title")
        .unwrap();
    assert_eq!(
        edit.commit()
            .unwrap()
            .package()
            .document()
            .unwrap()
            .paragraphs()
            .unwrap()[0]
            .text()
            .unwrap(),
        "unprotected title"
    );
}

#[cfg(feature = "encryption")]
#[test]
fn native_word_agile_password_ingress_and_rekey_preserve_clear_members() {
    use litchi_crypto::ooxml::{Error as CryptoError, Mode};
    let source = include_bytes!("../../../test-data/microsoft-office/word/word-encrypted.docx");
    for password in ["", "wrong"] {
        assert!(matches!(
            Package::from_bytes_with_password(source, password),
            Err(litchi_docx::Error::Encryption(CryptoError::Password))
        ));
    }
    let package = Package::from_bytes_with_password(source, "Office42!").unwrap();
    assert_eq!(package.encryption(), Some(Mode::AgileAes256Sha512));
    let document = package.document().unwrap();
    assert_eq!(document.paragraph_count().unwrap(), 5);
    assert_eq!(
        document
            .paragraphs()
            .unwrap()
            .iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect::<Vec<_>>(),
        [
            "WORD_ENCRYPTED_012",
            "Fixture ID: word-encrypted-012",
            "Native password-to-open specimen.",
            "VALUE_012: 42",
            "END_ENCRYPTED_012",
        ]
    );
    assert!(matches!(
        package.to_bytes(),
        Err(litchi_docx::Error::EncryptionPolicy { .. })
    ));
    let before = package.to_plain_bytes().unwrap();
    let encrypted = package.to_reencrypted("ChangedOffice42!").unwrap();
    assert!(matches!(
        Package::from_bytes_with_password(&encrypted, "Office42!"),
        Err(litchi_docx::Error::Encryption(CryptoError::Password))
    ));
    let reopened = Package::from_bytes_with_password(&encrypted, "ChangedOffice42!").unwrap();
    assert_eq!(reopened.encryption(), Some(Mode::AgileAes256Sha512));
    assert_members_equal(&before, &reopened.to_plain_bytes().unwrap());
}

#[cfg(feature = "vba-inspection")]
#[test]
fn native_word_macro_variants_expose_inert_project_graph_and_source() {
    use sha2::{Digest as _, Sha256};
    for (name, digest) in [
        (
            "word-recorded-macro.docm",
            "3cc296e042be194b7d1ddf30777731de852780c623e383b71e8f04dbef59f89f",
        ),
        (
            "word-recorded-macro.dotm",
            "9320243694e82e3f1206dbf7c5bbef7ba8e3510718bb325b566e2f43e34a42aa",
        ),
    ] {
        let package = native(name);
        let graph = package.vba().unwrap().unwrap();
        assert_eq!(graph.source_part_name().as_str(), "/word/document.xml");
        assert_eq!(graph.project_part_name().as_str(), "/word/vbaProject.bin");
        assert_eq!(
            graph.supplemental_data_part_name().as_str(),
            "/word/vbaData.xml"
        );
        assert_eq!(graph.project_relationship_id(), "rId1");
        assert_eq!(graph.supplemental_data_relationship_id(), "rId1");
        let project = graph
            .project(package.opc_package(), &litchi_vba::Limits::default())
            .unwrap();
        assert_eq!(project.name(), "Project");
        assert_eq!(
            project
                .modules()
                .iter()
                .map(|module| module.name())
                .collect::<Vec<_>>(),
            ["ThisDocument", "NewMacros"]
        );
        assert_eq!(project.references().len(), 3);
        assert!(
            project.modules()[1]
                .source()
                .text()
                .contains("Sub FixtureBoldSelection()")
        );
        assert!(
            project.modules()[1]
                .source()
                .text()
                .contains("Selection.Font.Bold = wdToggle")
        );
        let payloads = members(&package.to_bytes().unwrap());
        let payload = &payloads["word/vbaProject.bin"];
        assert_eq!(payload.len(), 12800);
        assert_eq!(
            Sha256::digest(payload)
                .iter()
                .map(|byte| format!("{byte:02x}"))
                .collect::<String>(),
            digest
        );
    }
}

#[test]
fn native_word_master_keeps_external_child_inert_and_child_opens_independently() {
    let package = native("word-master-document.docx");
    let document = package.document().unwrap();
    let paragraphs = document.paragraphs().unwrap();
    assert_eq!(paragraphs.len(), 4);
    assert!(!document.text().unwrap().contains("CHILD_SECTION_017"));
    let inlines = paragraphs[3].inlines().unwrap();
    let subdocuments: Vec<_> = inlines
        .iter()
        .filter_map(|inline| match inline {
            litchi_docx::Inline::Unknown(value)
                if value.kind() == litchi_docx::InlineKind::SubDocument =>
            {
                Some(value)
            },
            _ => None,
        })
        .collect();
    assert_eq!(subdocuments.len(), 1);
    assert!(
        std::str::from_utf8(subdocuments[0].xml_bytes())
            .unwrap()
            .contains("r:id=\"rId4\"")
    );
    let main = package.opc_package().main_document_part().unwrap();
    let references: Vec<_> = main
        .rels()
        .iter()
        .filter(|relationship| {
            relationship.reltype()
                == "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument"
        })
        .collect();
    assert_eq!(references.len(), 1);
    assert_eq!(references[0].r_id(), "rId4");
    assert!(references[0].is_external());
    assert_eq!(references[0].target_path(), "word-master-document/CHILD");
    let child = native("word-master-document/CHILD");
    assert_eq!(
        child
            .document()
            .unwrap()
            .paragraphs()
            .unwrap()
            .iter()
            .map(|paragraph| paragraph.text().unwrap())
            .collect::<Vec<_>>(),
        [
            "CHILD_SECTION_017",
            "Native master-document child content.",
            "",
            "",
        ]
    );
}
