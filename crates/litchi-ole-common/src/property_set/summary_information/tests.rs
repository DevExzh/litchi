use super::*;
use super::{ClipboardTag, ImageFormat};
use crate::property_set::{CodePage, SUMMARY_INFORMATION_FMTID, Section, Stream, Value};
use chrono::{DateTime, Duration, Utc};

fn populated(page: CodePage) -> Snapshot {
    let mut snapshot = Snapshot::new(page).unwrap();
    let mut transaction = snapshot.transaction().unwrap();
    {
        let mut edit = transaction.edit();
        edit.set_title("Title").unwrap();
        edit.set_subject("").unwrap();
        edit.set_author("Author").unwrap();
        edit.set_keywords("one;two").unwrap();
        edit.set_comments("Comments").unwrap();
        edit.set_template("Normal.dotm").unwrap();
        edit.set_last_author("Editor").unwrap();
        edit.set_revision_number("7").unwrap();
        edit.set_edit_time(FileTime::from_duration(Duration::seconds(12)).unwrap())
            .unwrap();
        edit.set_last_printed(FileTime::from_raw(11)).unwrap();
        edit.set_create_time(FileTime::from_raw(12)).unwrap();
        edit.set_last_save_time(FileTime::from_raw(13)).unwrap();
        edit.set_page_count(0).unwrap();
        edit.set_word_count(42).unwrap();
        edit.set_character_count(100).unwrap();
        edit.set_thumbnail(Thumbnail::empty()).unwrap();
        edit.set_app_name("litchi").unwrap();
        edit.set_document_security(
            DocumentSecurity::PASSWORD_PROTECTED | DocumentSecurity::READ_ONLY_RECOMMENDED,
        )
        .unwrap();
    }
    snapshot = Snapshot::from_section(&transaction.commit().unwrap().into_section()).unwrap();
    snapshot
}

#[test]
fn typed_pid_values_round_trip_without_reimplementing_wire_grammar() {
    let snapshot = populated(CodePage::WINDOWS_1252);
    let stream = Stream::new(snapshot.section().clone());
    let parsed = Stream::parse(&stream.to_bytes().unwrap()).unwrap();
    let typed = Snapshot::from_stream(&parsed).unwrap();

    assert_eq!(typed.codepage(), Some(CodePage::WINDOWS_1252));
    assert_eq!(typed.title(), Some("Title"));
    assert_eq!(typed.subject(), Some(""));
    assert_eq!(typed.author(), Some("Author"));
    assert_eq!(typed.keywords(), Some("one;two"));
    assert_eq!(typed.comments(), Some("Comments"));
    assert_eq!(typed.template(), Some("Normal.dotm"));
    assert_eq!(typed.last_author(), Some("Editor"));
    assert_eq!(typed.revision_number(), Some("7"));
    assert_eq!(
        typed.edit_time().unwrap().duration(),
        Some(Duration::seconds(12))
    );
    assert_eq!(typed.last_printed().unwrap().raw(), 11);
    assert_eq!(typed.create_time().unwrap().raw(), 12);
    assert_eq!(typed.last_save_time().unwrap().raw(), 13);
    assert_eq!(typed.page_count(), Some(0));
    assert_eq!(typed.word_count(), Some(42));
    assert_eq!(typed.character_count(), Some(100));
    let thumbnail = typed.thumbnail().unwrap();
    assert_eq!(thumbnail.tag(), ClipboardTag::Empty);
    assert_eq!(thumbnail.format(), None);
    assert!(thumbnail.is_empty());
    assert_eq!(typed.app_name(), Some("litchi"));
    assert!(
        typed
            .document_security()
            .unwrap()
            .contains(DocumentSecurity::PASSWORD_PROTECTED)
    );
    assert!(matches!(typed.property(TITLE), Some(Value::Lpstr(value)) if value == "Title"));
}

#[test]
fn unicode_codepage_uses_the_codepage_aware_lpwstr_variant() {
    let snapshot = Snapshot::new(CodePage::Utf16Le).unwrap();
    let mut transaction = snapshot.transaction().unwrap();
    transaction.edit().set_title("界").unwrap();
    let committed = transaction.commit().unwrap();
    assert!(matches!(
        committed.section().property(TITLE),
        Some(Value::Lpwstr(value)) if value == "界"
    ));

    let parsed =
        Stream::parse(&Stream::new(committed.section().clone()).to_bytes().unwrap()).unwrap();
    assert_eq!(Snapshot::from_stream(&parsed).unwrap().title(), Some("界"));
}

#[test]
fn absence_is_distinct_from_explicit_empty_zero_and_false_like_values() {
    let empty = Snapshot::new(CodePage::WINDOWS_1252).unwrap();
    assert_eq!(empty.title(), None);
    assert_eq!(empty.page_count(), None);
    assert_eq!(empty.document_security(), None);
    assert_eq!(empty.thumbnail(), None);

    let explicit = populated(CodePage::WINDOWS_1252);
    assert_eq!(explicit.subject(), Some(""));
    assert_eq!(explicit.page_count(), Some(0));
    assert_eq!(explicit.document_security().unwrap().bits(), 0x0000_0003);
    assert_eq!(explicit.thumbnail().unwrap().len(), 0);

    let mut transaction = explicit.transaction().unwrap();
    transaction.edit().remove(SUBJECT).unwrap();
    assert_eq!(
        transaction.commit().unwrap().section().property(SUBJECT),
        None
    );
}

#[test]
fn unknown_properties_survive_typed_edits_and_round_trip_in_order() {
    let mut section = Section::new(SUMMARY_INFORMATION_FMTID);
    section.set_page(CodePage::WINDOWS_1252);
    section
        .add(
            0x0000_0020,
            Value::Unknown {
                variant_type: 0x7F01,
                data: vec![0x01, 0x02, 0x03, 0x04],
            },
        )
        .unwrap();
    section.add(TITLE, Value::Lpstr("before".into())).unwrap();
    let source_order: Vec<_> = section.property_ids().collect();
    let snapshot = Snapshot::from_section(&section).unwrap();

    let mut transaction = snapshot.transaction().unwrap();
    transaction.edit().set_title("after").unwrap();
    let commit = transaction.commit().unwrap();
    assert_eq!(commit.section().property(0x20), section.property(0x20));
    assert_eq!(
        commit.section().property_ids().collect::<Vec<_>>(),
        source_order
    );

    let bytes = Stream::new(commit.section().clone()).to_bytes().unwrap();
    let parsed = Stream::parse(&bytes).unwrap();
    let typed = Snapshot::from_stream(&parsed).unwrap();
    assert_eq!(typed.property(0x20), section.property(0x20));
    assert_eq!(typed.title(), Some("after"));
}

#[test]
fn codepage_and_typed_values_are_checked_before_publication() {
    let snapshot = Snapshot::new(CodePage::WINDOWS_1252).unwrap();
    let mut transaction = snapshot.transaction().unwrap();
    assert!(transaction.edit().set_title("界").is_err());
    assert_eq!(transaction.commit().unwrap().changed(), false);

    let oversized = "x".repeat(MAX_TEXT_BYTES + 1);
    let mut transaction = snapshot.transaction().unwrap();
    assert!(transaction.edit().set_title(&oversized).is_err());
    assert!(!transaction.commit().unwrap().changed());

    let mut wrong_type = Section::new(SUMMARY_INFORMATION_FMTID);
    wrong_type.set_page(CodePage::WINDOWS_1252);
    wrong_type.add(TITLE, Value::I4(1)).unwrap();
    assert!(Snapshot::from_section(&wrong_type).is_err());

    let mut negative_count = Section::new(SUMMARY_INFORMATION_FMTID);
    negative_count.set_page(CodePage::WINDOWS_1252);
    negative_count.add(PAGE_COUNT, Value::I4(-1)).unwrap();
    assert!(Snapshot::from_section(&negative_count).is_err());

    assert_eq!(DocumentSecurity::new(0x10).unknown_bits(), 0x10);
    assert!(
        Thumbnail::new(
            ClipboardTag::Windows,
            Some(ImageFormat::Jpeg),
            vec![0; MAX_THUMBNAIL_BYTES + 1]
        )
        .is_err()
    );
}

#[test]
fn filetime_conversions_and_reversible_patches_are_source_checked() {
    let epoch = DateTime::<Utc>::from_timestamp(0, 0).unwrap();
    let filetime = FileTime::from_date_time(epoch).unwrap();
    assert_eq!(filetime.raw(), 116_444_736_000_000_000);
    assert_eq!(filetime.date_time(), Some(epoch));
    assert!(FileTime::from_duration(Duration::nanoseconds(1)).is_err());

    let snapshot = Snapshot::new(CodePage::WINDOWS_1252).unwrap();
    let source = snapshot.section().clone();
    let mut transaction = snapshot.transaction().unwrap();
    transaction.edit().set_title("changed").unwrap();
    let commit = transaction.commit().unwrap();
    let patch = commit.patch();
    let forward = patch.apply(&source).unwrap();
    assert_eq!(forward, *commit.section());
    assert_eq!(patch.revert(&forward).unwrap(), source);

    let mut different = forward.clone();
    different
        .add(SUBJECT, Value::Lpstr("wrong".into()))
        .unwrap();
    assert!(patch.revert(&different).is_err());
    assert!(patch.apply(&different).is_err());
}

#[test]
fn missing_codepage_and_wrong_identity_are_rejected() {
    let mut missing = Section::new(SUMMARY_INFORMATION_FMTID);
    missing.add(TITLE, Value::Lpstr("x".into())).unwrap();
    assert!(Snapshot::from_section(&missing).is_err());

    let mut wrong = Section::new(crate::property_set::DOCUMENT_SUMMARY_INFORMATION_FMTID);
    wrong.set_page(CodePage::WINDOWS_1252);
    assert!(Snapshot::from_section(&wrong).is_err());

    let mut versioned = Stream::new(
        Snapshot::new(CodePage::WINDOWS_1252)
            .unwrap()
            .section()
            .clone(),
    );
    versioned.version = Stream::VERSION_1;
    assert!(Snapshot::from_stream(&versioned).is_err());
}
