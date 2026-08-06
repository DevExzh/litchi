use super::*;
use crate::property_set::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, HeadingPair, HeadingPairs, Section, Stream, Value,
};

#[test]
fn typed_pidssi_round_trip_uses_contextual_accessors() {
    let snapshot = Snapshot::new(CodePage::Utf16Le).unwrap();
    let mut transaction = snapshot.transaction().unwrap();
    {
        let mut edit = transaction.edit();
        edit.set_category("Presentation").unwrap();
        edit.set_presentation_format("On-screen Show (16:9)")
            .unwrap();
        edit.set_byte_count(2048).unwrap();
        edit.set_line_count(12).unwrap();
        edit.set_paragraph_count(4).unwrap();
        edit.set_slide_count(3).unwrap();
        edit.set_note_count(1).unwrap();
        edit.set_hidden_count(0).unwrap();
        edit.set_multimedia_clip_count(2).unwrap();
        edit.set_scale(false).unwrap();
        edit.set_heading_pairs(
            HeadingPairs::new(vec![HeadingPair::new("Slides", 2).unwrap()]).unwrap(),
        )
        .unwrap();
        edit.set_document_parts(vec!["Slide 1".into(), "Slide 2".into()])
            .unwrap();
        edit.set_manager("Manager").unwrap();
        edit.set_company("Company").unwrap();
        edit.set_links_dirty(true).unwrap();
        edit.set_character_count_with_spaces(99).unwrap();
        edit.set_shared_document(false).unwrap();
        edit.set_hyperlinks_changed(true).unwrap();
        edit.set_version(Version::new(3, 7).unwrap()).unwrap();
        edit.set_content_type("presentation").unwrap();
        edit.set_content_status("Draft").unwrap();
        edit.set_language("en-US").unwrap();
        edit.set_document_version("3.7").unwrap();
    }
    let commit = transaction.commit().unwrap();
    assert!(commit.changed());
    let stream = Stream::new(commit.section().clone());
    let parsed = Stream::parse(&stream.to_bytes().unwrap()).unwrap();
    let typed = Snapshot::from_stream(&parsed).unwrap();

    assert_eq!(typed.codepage(), Some(CodePage::Utf16Le));
    assert_eq!(typed.category(), Some("Presentation"));
    assert_eq!(typed.presentation_format(), Some("On-screen Show (16:9)"));
    assert_eq!(typed.byte_count(), Some(2048));
    assert_eq!(typed.line_count(), Some(12));
    assert_eq!(typed.paragraph_count(), Some(4));
    assert_eq!(typed.slide_count(), Some(3));
    assert_eq!(typed.note_count(), Some(1));
    assert_eq!(typed.hidden_count(), Some(0));
    assert_eq!(typed.multimedia_clip_count(), Some(2));
    assert_eq!(typed.scale(), Some(false));
    assert_eq!(typed.heading_pairs().unwrap().document_part_count(), 2);
    assert_eq!(typed.document_parts().unwrap().value(1), Some("Slide 2"));
    assert_eq!(typed.manager(), Some("Manager"));
    assert_eq!(typed.company(), Some("Company"));
    assert_eq!(typed.links_dirty(), Some(true));
    assert_eq!(typed.character_count_with_spaces(), Some(99));
    assert_eq!(typed.shared_document(), Some(false));
    assert_eq!(typed.hyperlinks_changed(), Some(true));
    assert_eq!(typed.version(), Version::new(3, 7));
    assert_eq!(typed.content_type(), Some("presentation"));
    assert_eq!(typed.content_status(), Some("Draft"));
    assert_eq!(typed.language(), Some("en-US"));
    assert_eq!(typed.document_version(), Some("3.7"));
    assert!(matches!(
        typed.property(CATEGORY),
        Some(Value::Lpwstr(value)) if value == "Presentation"
    ));
}

#[test]
fn opaque_pidssi_payloads_survive_typed_edits_and_wire_round_trip() {
    let mut section = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    section.set_page(CodePage::WINDOWS_1252);
    section
        .add(
            DIGITAL_SIGNATURE,
            Value::Unknown {
                variant_type: 0x7F01,
                data: vec![0x00, 0xAA, 0x55, 0xFE],
            },
        )
        .unwrap();
    section
        .add(
            0x25,
            Value::Unknown {
                variant_type: 0x7F02,
                data: vec![1, 3, 3, 7],
            },
        )
        .unwrap();

    let snapshot = Snapshot::from_section(&section).unwrap();
    let mut transaction = snapshot.transaction().unwrap();
    transaction.edit().set_category("Changed").unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(
        commit.section().property(DIGITAL_SIGNATURE),
        section.property(DIGITAL_SIGNATURE)
    );
    assert_eq!(commit.section().property(0x25), section.property(0x25));

    let stream = Stream::new(commit.section().clone());
    let parsed = Stream::parse(&stream.to_bytes().unwrap()).unwrap();
    let typed = Snapshot::from_stream(&parsed).unwrap();
    assert_eq!(
        typed.property(DIGITAL_SIGNATURE),
        section.property(DIGITAL_SIGNATURE)
    );
    assert_eq!(typed.property(0x25), section.property(0x25));
    assert_eq!(typed.category(), Some("Changed"));
}

#[test]
fn patch_requires_the_exact_source_and_failed_edits_are_transactional() {
    let snapshot = Snapshot::new(CodePage::WINDOWS_1252).unwrap();
    let original = snapshot.section().clone();
    let mut transaction = snapshot.transaction().unwrap();
    assert!(transaction.edit().set_scale(true).is_err());
    let unchanged = transaction.commit().unwrap();
    assert!(!unchanged.changed());
    assert_eq!(unchanged.section(), &original);

    let mut transaction = snapshot.transaction().unwrap();
    transaction.edit().set_category("New").unwrap();
    let commit = transaction.commit().unwrap();
    let patch = commit.patch();
    assert_eq!(
        patch.apply(&original).unwrap().property(CATEGORY),
        commit.section().property(CATEGORY)
    );

    let mut different_source = original.clone();
    different_source
        .add(CATEGORY, Value::Lpstr("Other".into()))
        .unwrap();
    assert!(patch.apply(&different_source).is_err());
}

#[test]
fn known_pid_types_and_required_invariants_are_checked() {
    assert_eq!(Version::new(0, 1), None);
    assert!(Version::from_raw(0).is_err());

    let mut wrong_type = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    wrong_type.set_page(CodePage::WINDOWS_1252);
    wrong_type.add(CATEGORY, Value::I4(1)).unwrap();
    assert!(Snapshot::from_section(&wrong_type).is_err());

    let mut named = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    named.set_page(CodePage::WINDOWS_1252);
    named
        .add_named(2, "Category".into(), Value::Lpstr("x".into()))
        .unwrap();
    assert!(Snapshot::from_section(&named).is_err());

    let mut missing_codepage = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    missing_codepage
        .add(CATEGORY, Value::Lpstr("x".into()))
        .unwrap();
    assert!(Snapshot::from_section(&missing_codepage).is_err());
}

#[test]
fn typed_string_edits_are_bounded_and_follow_the_section_codepage() {
    let ansi_snapshot = Snapshot::new(CodePage::WINDOWS_1252).unwrap();
    let mut ansi = ansi_snapshot.transaction().unwrap();
    ansi.edit().set_category("ANSI").unwrap();
    let commit = ansi.commit().unwrap();
    assert!(matches!(
        commit.section().property(CATEGORY),
        Some(Value::Lpstr(_))
    ));

    let unicode_snapshot = Snapshot::new(CodePage::Utf16Le).unwrap();
    let mut unicode = unicode_snapshot.transaction().unwrap();
    unicode.edit().set_category("Unicode").unwrap();
    let commit = unicode.commit().unwrap();
    assert!(matches!(
        commit.section().property(CATEGORY),
        Some(Value::Lpwstr(_))
    ));

    let oversized = "x".repeat(MAX_TEXT_BYTES + 1);
    let bounded_snapshot = Snapshot::new(CodePage::WINDOWS_1252).unwrap();
    let mut bounded = bounded_snapshot.transaction().unwrap();
    assert!(bounded.edit().set_category(&oversized).is_err());
    assert!(!bounded.commit().unwrap().changed());
}
