//! Native Microsoft Word RTF layout and auxiliary-story regressions.
use litchi_rtf::{Package, PageOrientation};
fn native(name: &str) -> Package {
    Package::open(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(name),
    )
    .unwrap()
}
#[test]
fn native_section_resets_keep_document_margins_and_landscape_columns() {
    let package = native("word-tables-layout.rtf");
    let sections = package.sections();
    assert_eq!(sections.len(), 2);
    for section in sections {
        let p = &section.properties;
        assert_eq!(
            (p.margin_left, p.margin_right, p.margin_top, p.margin_bottom),
            (1440, 1440, 1440, 1440)
        );
    }
    assert_eq!(
        (
            sections[0].properties.page_width,
            sections[0].properties.page_height,
            sections[0].properties.orientation
        ),
        (12240, 15840, PageOrientation::Portrait)
    );
    assert_eq!(
        (
            sections[1].properties.page_width,
            sections[1].properties.page_height,
            sections[1].properties.orientation
        ),
        (15840, 12240, PageOrientation::Landscape)
    );
    assert_eq!(sections[1].properties.columns.count, 2);
    assert_eq!(sections[1].properties.document_grid.line_grid, Some(360));
    assert_eq!(
        sections[1].properties.document_grid.grid_type,
        Some(litchi_rtf::SectionDocumentGridType::Default)
    );
}
#[test]
fn native_grouped_section_break_binds_the_landscape_section() {
    let package = native("word-tables-layout.rtf");
    let boundaries = package.raw().section_breaks().collect::<Vec<_>>();
    assert_eq!(boundaries.len(), 1);
    assert_eq!(boundaries[0].next_section, Some(1));
    let body = package.text();
    let prefix = "Second section begins below.\n";
    assert_eq!(
        boundaries[0].position,
        body.find(prefix).unwrap() + prefix.len()
    );
}

#[test]
fn native_header_footer_stories_retain_paragraphs_and_cached_page_fields() {
    use litchi_rtf::{FieldOwner, FieldType, HeaderFooterType as H};
    let package = native("word-tables-layout.rtf");
    let stories = &package.sections()[0].headers_footers;
    assert_eq!(
        stories
            .iter()
            .map(|s| (s.header_type, s.text()))
            .collect::<Vec<_>>(),
        [
            (H::HeaderRight, "HEADER_DEFAULT_002\n".into()),
            (H::FooterLeft, "\n".into()),
            (H::FooterRight, "\nFOOTER_PAGE_002:".into())
        ]
    );
    assert!(package.sections()[1].headers_footers.is_empty());
    let fields = package.raw().fields();
    assert_eq!(fields.len(), 2);
    assert_eq!(
        fields
            .iter()
            .map(|f| (
                f.field_type,
                f.owner,
                f.instruction.trim(),
                f.result.as_ref()
            ))
            .collect::<Vec<_>>(),
        [
            (FieldType::Page, FieldOwner::Footer, "PAGE", ""),
            (FieldType::Page, FieldOwner::Footer, "PAGE", "2")
        ]
    );
    assert_eq!(
        stories[1].story_events,
        [litchi_rtf::StoryEvent::Field(litchi_rtf::StoryField {
            field_index: 0,
            position: 0
        })]
    );
    assert_eq!(
        stories[2].story_events,
        [litchi_rtf::StoryEvent::Field(litchi_rtf::StoryField {
            field_index: 1,
            position: 0
        })]
    );
    use litchi_rtf::{NoteSeparatorElement as E, NoteSeparatorKind as K};
    assert_eq!(package.raw().note_separators().entries().len(), 4);
    for (kind, mark) in [
        (K::FootnoteSeparator, E::SeparatorMark),
        (
            K::FootnoteContinuationSeparator,
            E::ContinuationSeparatorMark,
        ),
        (K::EndnoteSeparator, E::SeparatorMark),
        (
            K::EndnoteContinuationSeparator,
            E::ContinuationSeparatorMark,
        ),
    ] {
        let separator = package.raw().note_separators().get(kind).unwrap();
        assert_eq!(separator.elements, [mark, E::ParagraphBreak]);
        assert!(separator.shapes.is_empty() && separator.shape_groups.is_empty());
    }
}

#[test]
fn native_footnote_and_endnote_stories_keep_their_exact_text_and_kind() {
    for (name, footnote, text) in [
        (
            "word-tables-layout.rtf",
            false,
            " ENDNOTE_TEXT_002: native endnote value 84.",
        ),
        (
            "word-layout-extensions.rtf",
            true,
            " FOOTNOTE_COLUMN_016: native two-column footnote-area specimen. This text exercises a footnote section layout separately from the main document columns.",
        ),
    ] {
        let package = native(name);
        let notes = package.notes();
        assert_eq!(notes.len(), 1);
        assert_eq!(notes[0].is_footnote, footnote);
        assert_eq!(
            package.raw().note_options().present_kinds,
            Some(if footnote {
                litchi_rtf::PresentNoteKinds::FootnotesOnly
            } else {
                litchi_rtf::PresentNoteKinds::EndnotesOnly
            })
        );
        assert_eq!(
            package.raw().note_options().endnote_placement,
            Some(litchi_rtf::NotePlacement::EndOfDocument)
        );

        assert_eq!(notes[0].content, text);
        assert_eq!(notes[0].reference, if footnote { "1" } else { "i" });
        assert!(notes[0].position <= package.text().len());
        assert!(!package.text().contains(text.trim()));
        assert_eq!(
            package
                .raw()
                .body_story_events()
                .iter()
                .filter(|event| matches!(event, litchi_rtf::BodyStoryEvent::Note(0)))
                .count(),
            1
        );
    }
}

#[test]
fn native_passive_layout_policies_survive_canonical_writing() {
    use litchi_rtf::{DocumentViewKind, NotePlacement, PresentNoteKinds};
    let package = native("word-rtf-minimal.rtf");
    let doc = package.raw();
    assert_eq!(doc.document_view().kind, Some(DocumentViewKind::PageLayout));
    assert_eq!(doc.document_view().scale_percent, Some(100));
    assert_eq!(doc.print_layout_settings().document_gutter_twips, Some(0));
    assert!(doc.legacy_layout_compatibility().use_legacy_footnote_layout);
    assert_eq!(doc.xml_policies().validate_custom_xml, Some(true));
    assert_eq!(
        doc.note_options().present_kinds,
        Some(PresentNoteKinds::FootnotesOnly)
    );
    assert_eq!(
        doc.note_options().footnote_placement,
        Some(NotePlacement::BottomOfPage)
    );
    let mut bytes = Vec::new();
    litchi_rtf::write::Writer::new(&mut bytes)
        .write_document(doc)
        .unwrap();
    let reopened = Package::from_bytes(&bytes).unwrap();
    assert_eq!(reopened.raw().document_view(), doc.document_view());
    assert_eq!(
        reopened.raw().print_layout_settings(),
        doc.print_layout_settings()
    );
    assert_eq!(
        reopened.raw().legacy_layout_compatibility(),
        doc.legacy_layout_compatibility()
    );
    assert_eq!(reopened.raw().xml_policies(), doc.xml_policies());
    assert_eq!(reopened.raw().note_options(), doc.note_options());
    assert_eq!(
        reopened.sections()[0].properties,
        package.sections()[0].properties
    );
    assert_eq!(reopened.text(), package.text());
}

#[test]
fn native_running_header_edit_changes_only_the_selected_text() {
    let package = native("word-tables-layout.rtf");
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit();
    edit.set_header_footer_text(
        litchi_rtf::edit::HeaderFooterParagraph::new(
            0,
            litchi_rtf::HeaderFooterType::HeaderRight,
            0,
        ),
        "HEADER_CHANGED_022",
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let expected = String::from_utf8(before.clone())
        .unwrap()
        .replace("HEADER_DEFAULT_002", "HEADER_CHANGED_022");
    assert!(after == expected.as_bytes());
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(
        reopened.sections()[0].headers_footers[0].text(),
        "HEADER_CHANGED_022\n"
    );
    assert_eq!(reopened.text(), package.text());
    assert!(
        reopened
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
}

#[test]
fn native_note_text_edits_preserve_main_story_sections_and_note_identity() {
    for name in ["word-tables-layout.rtf", "word-layout-extensions.rtf"] {
        let package = native(name);
        let before = package.to_bytes().unwrap();
        let note = &package.notes()[0];
        let mut edit = package.edit();
        edit.set_note_text(0, "NATIVE_NOTE_CHANGED_022").unwrap();
        let commit = edit.commit().unwrap();
        let directory = tempfile::tempdir().unwrap();
        let path = directory.path().join("edited.rtf");
        commit.package().save(&path).unwrap();
        let reopened = Package::open(&path).unwrap();
        let after = reopened.to_bytes().unwrap();
        let old_range = note_group(&before);
        let new_range = note_group(&after);
        assert!(before[..old_range.start] == after[..new_range.start]);
        assert!(before[old_range.end..] == after[new_range.end..]);
        assert_eq!(
            before.windows(6).filter(|w| *w == b"\\chftn").count(),
            after.windows(6).filter(|w| *w == b"\\chftn").count()
        );

        assert_eq!(reopened.notes()[0].content, "NATIVE_NOTE_CHANGED_022");
        assert_eq!(
            (
                reopened.notes()[0].is_footnote,
                reopened.notes()[0].reference.as_ref(),
                reopened.notes()[0].position
            ),
            (note.is_footnote, note.reference.as_ref(), note.position)
        );
        assert_eq!(reopened.text(), package.text());
        assert_eq!(
            reopened
                .sections()
                .iter()
                .map(|s| &s.properties)
                .collect::<Vec<_>>(),
            package
                .sections()
                .iter()
                .map(|s| &s.properties)
                .collect::<Vec<_>>()
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
        assert!(package.to_bytes().unwrap() == before);
    }
}

#[test]
fn native_landscape_section_edit_preserves_the_first_section_and_auxiliary_stories() {
    let package = native("word-tables-layout.rtf");
    let before = package.to_bytes().unwrap();
    let mut edit = package.edit();
    edit.set_section_values(
        1,
        &[
            litchi_rtf::edit::SectionValue::Columns(3),
            litchi_rtf::edit::SectionValue::MarginLeft(1000),
        ],
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let bytes = commit.package().to_bytes().unwrap();
    let expected = String::from_utf8(before.clone())
        .unwrap()
        .replace("\\cols2", "\\cols3")
        .replace(
            "\\sectd \\ltrsect\\lndscpsxn",
            "\\sectd \\marglsxn1000\\ltrsect\\lndscpsxn",
        );
    assert!(bytes == expected.as_bytes());
    let reopened = Package::from_bytes(&bytes).unwrap();
    assert_eq!(
        reopened.sections()[0].properties,
        package.sections()[0].properties
    );
    assert_eq!(reopened.sections()[1].properties.columns.count, 3);
    assert_eq!(reopened.sections()[1].properties.margin_left, 1000);
    assert_eq!(reopened.notes()[0].content, package.notes()[0].content);
    assert_eq!(reopened.text(), package.text());
    assert!(
        reopened
            .apply(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == before
    );
}

// These native note groups contain no binary payloads; escaped braces remain text.
fn note_group(bytes: &[u8]) -> std::ops::Range<usize> {
    let start = bytes
        .windows(b"{\\footnote".len())
        .position(|w| w == b"{\\footnote")
        .unwrap();
    let mut depth = 0;
    let mut at = start;
    while at < bytes.len() {
        match bytes[at] {
            b'\\' => {
                at += 2;
                continue;
            },
            b'{' => depth += 1,
            b'}' => {
                depth -= 1;
                if depth == 0 {
                    return start..at + 1;
                }
            },
            _ => {},
        }
        at += 1;
    }
    panic!("native note group is not closed")
}

#[test]
fn native_visual_page_breaks_keep_exact_body_offsets_and_event_order() {
    for (name, positions) in [
        ("word-visual-objects.rtf", [96, 115]),
        ("word-vector-diagram.rtf", [75, 97]),
    ] {
        let package = native(name);
        let before = package.to_bytes().unwrap();
        assert_eq!(
            package
                .raw()
                .page_breaks()
                .map(|p| p.position)
                .collect::<Vec<_>>(),
            positions
        );
        assert_eq!(package.raw().column_breaks().count(), 0);
        assert_eq!(package.raw().soft_breaks().count(), 0);
        let reopened = Package::from_bytes(&before).unwrap();
        assert_eq!(
            reopened.raw().body_story_events(),
            package.raw().body_story_events()
        );
        assert_eq!(reopened.text(), package.text());
        if name == "word-visual-objects.rtf" {
            let mut bytes = Vec::new();
            litchi_rtf::write::Writer::new(&mut bytes)
                .write_document(package.raw())
                .unwrap();
            let written = Package::from_bytes(&bytes).unwrap();
            assert_eq!(
                written
                    .raw()
                    .page_breaks()
                    .map(|p| p.position)
                    .collect::<Vec<_>>(),
                positions
            );
            assert_eq!(written.text(), package.text());
            assert_eq!(
                written.raw().body_story_events(),
                package.raw().body_story_events()
            );
        }
    }
}
