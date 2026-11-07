//! Explicit inert metadata on native binary PowerPoint parents.
use litchi_core::Position;
use litchi_ppt::{Package, Record, RecordParser, RecordType, SourceLimits};
use std::{collections::BTreeMap, io::Cursor, path::Path};
#[path = "../../../tools/test_support/cfb_directory_metadata.rs"]
mod cfb_metadata;

const PARENTS: [&str; 4] = [
    "powerpoint-visual-objects.ppt",
    "powerpoint-visual-objects.pps",
    "powerpoint-visual-objects.pot",
    "powerpoint-passive-macro.ppt",
];
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let bytes = ole
                .open_stream(&path.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (path, bytes)
        })
        .collect()
}
fn wire(r: &Record) -> Vec<u8> {
    let mut bytes = ((r.instance << 4) | r.version).to_le_bytes().to_vec();
    bytes.extend_from_slice(&r.record_type_raw.to_le_bytes());
    bytes.extend_from_slice(&r.data_length.to_le_bytes());
    bytes.extend_from_slice(&r.data);
    bytes
}
fn persisted(bytes: &[u8]) -> BTreeMap<u32, Vec<u8>> {
    let parts = streams(bytes);
    let data = &parts[&vec!["PowerPoint Document".into()]];
    let limits = SourceLimits::default();
    let mut parser = RecordParser::new();
    parser.parse_document(data, &limits).unwrap();
    let map = litchi_ppt::PersistMapping::build_from_records_ref(&parser.find_records_ref());
    map.get_persist_ids()
        .into_iter()
        .map(|id| {
            let (record, _) =
                Record::parse(data, map.get_offset(id).unwrap() as usize, &limits).unwrap();
            (id, wire(&record))
        })
        .collect()
}
fn document(p: &Package) -> Record {
    let id = p
        .presentation()
        .unwrap()
        .slide_directory()
        .document_persist_id();
    Record::parse(
        &persisted(&p.to_bytes().unwrap())[&id],
        0,
        &SourceLimits::default(),
    )
    .unwrap()
    .0
}

fn atom(kind: RecordType, data: Vec<u8>) -> Record {
    Record {
        record_type: kind,
        record_type_raw: kind.as_u16(),
        version: 0,
        instance: 0,
        data_length: u32::try_from(data.len()).unwrap(),
        data: data.into(),
        children: Vec::new(),
    }
}
fn rebuild(record: &mut Record) {
    record.data = record
        .children
        .iter()
        .flat_map(wire)
        .collect::<Vec<_>>()
        .into();
    record.data_length = u32::try_from(record.data.len()).unwrap();
}
fn container(kind: RecordType, children: Vec<Record>) -> Record {
    let mut result = atom(kind, Vec::new());
    result.version = 15;
    result.children = children;
    rebuild(&mut result);
    result
}
fn tagged(mut doc: Record, version: u8, payload: Vec<u8>) -> Record {
    let info = doc
        .children
        .iter_mut()
        .find(|r| r.record_type == RecordType::DocInfoList)
        .unwrap();
    let tags = info
        .children
        .iter_mut()
        .find(|r| r.record_type == RecordType::ProgTags)
        .unwrap();
    let tag_name = format!("___PPT{version}")
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect::<Vec<_>>();
    let replacement = container(
        RecordType::ProgBinaryTag,
        vec![
            atom(RecordType::CString, tag_name.clone()),
            atom(RecordType::BinaryTagData, payload),
        ],
    );
    let existing = tags.children.iter_mut().find(|r| {
        r.children
            .first()
            .is_some_and(|name| name.data.as_ref() == tag_name)
    });
    if let Some(existing) = existing {
        *existing = replacement;
    } else {
        tags.children.push(replacement);
    }
    rebuild(tags);
    rebuild(info);
    rebuild(&mut doc);
    let parsed = Record::parse(&wire(&doc), 0, &SourceLimits::default())
        .unwrap()
        .0;
    litchi_ppt::ProgTags::parse_document(&parsed, litchi_ppt::ProgTagLimits::default())
        .unwrap()
        .unwrap()
        .document_extensions()
        .unwrap();
    parsed
}

/// Seed a clearly authored document derivative through public record/CFB APIs.
/// Existing persisted objects stay at their original offsets. Only the live
/// Document reference is overridden by the appended incremental edit.
fn seed_document(p: &Package, doc: &Record) -> Package {
    let bytes = p.to_bytes().unwrap();
    let parts = streams(&bytes);
    let key = vec!["PowerPoint Document".into()];
    let user_key = vec!["Current User".into()];
    let mut data = parts[&key].clone();
    let mut user = parts[&user_key].clone();
    let old_edit_offset = u32::from_le_bytes(user[16..20].try_into().unwrap());
    let (mut user_edit, _) =
        Record::parse(&data, old_edit_offset as usize, &SourceLimits::default()).unwrap();
    assert_eq!(user_edit.record_type, RecordType::UserEditAtom);
    assert_eq!(user_edit.data.len(), 28);
    let doc_id = p
        .presentation()
        .unwrap()
        .slide_directory()
        .document_persist_id();
    assert_eq!(
        u32::from_le_bytes(user_edit.data[16..20].try_into().unwrap()),
        doc_id
    );
    let doc_offset = u32::try_from(data.len()).unwrap();
    data.extend_from_slice(&wire(doc));
    let directory_offset = u32::try_from(data.len()).unwrap();
    let mut directory = litchi_ppt::writer::persist::PersistPtrBuilder::new();
    directory.set_offset(doc_id, doc_offset);
    data.extend_from_slice(&directory.generate_incremental_record());
    let next_edit_offset = u32::try_from(data.len()).unwrap();
    user_edit.data[8..12].copy_from_slice(&old_edit_offset.to_le_bytes());
    user_edit.data[12..16].copy_from_slice(&directory_offset.to_le_bytes());
    data.extend_from_slice(&wire(&user_edit));
    user[16..20].copy_from_slice(&next_edit_offset.to_le_bytes());
    let source = litchi_ole_common::object::Package::open(bytes.clone()).unwrap();
    let mut edit = source.edit().unwrap();
    edit.put_stream(&key, data).unwrap();
    edit.put_stream(&user_key, user).unwrap();
    let c = edit.commit().unwrap();
    let result = Package::from_bytes(c.patch().after()).unwrap();
    preserved(p, &result, RecordType::Document);
    assert_eq!(document(&result), *doc);
    let target = litchi_ole_common::object::Package::open(result.to_bytes().unwrap()).unwrap();
    assert_eq!(
        c.patch().inverse().apply(&target).unwrap().patch().after(),
        bytes
    );
    reopen(&result)
}
fn preserved(before: &Package, after: &Package, owned_type: RecordType) {
    let a = before.to_bytes().unwrap();
    let b = after.to_bytes().unwrap();
    assert_eq!(cfb_metadata::directory(&a), cfb_metadata::directory(&b));
    let x = streams(&a);
    let y = streams(&b);
    assert_eq!(x.keys().collect::<Vec<_>>(), y.keys().collect::<Vec<_>>());
    for (name, data) in &x {
        if name != &["PowerPoint Document"] && name != &["Current User"] {
            assert!(y.get(name) == Some(data), "changed stream {name:?}");
        }
    }
    assert!(
        y[&vec!["PowerPoint Document".into()]].starts_with(&x[&vec!["PowerPoint Document".into()]])
    );
    let x = persisted(&a);
    let y = persisted(&b);
    assert_eq!(x.keys().collect::<Vec<_>>(), y.keys().collect::<Vec<_>>());
    let changed = x
        .iter()
        .filter(|(id, bytes)| y.get(id) != Some(bytes))
        .collect::<Vec<_>>();
    assert_eq!(changed.len(), 1);
    let old = Record::parse(changed[0].1, 0, &SourceLimits::default())
        .unwrap()
        .0;
    assert_eq!(old.record_type, owned_type);
    let p = before.presentation().unwrap();
    let q = after.presentation().unwrap();
    assert_eq!(p.slide_count(), q.slide_count());
    assert_eq!(
        p.extract_text_fast().unwrap(),
        q.extract_text_fast().unwrap()
    );
    for (left, right) in p.slides().unwrap().iter().zip(q.slides().unwrap()) {
        assert_eq!(left.slide_id(), right.slide_id());
        assert_eq!(left.text().unwrap(), right.text().unwrap());
        assert_eq!(
            left.powerpoint12_round_trip_metadata().unwrap(),
            right.powerpoint12_round_trip_metadata().unwrap()
        );
    }
}
struct TempPpt(std::path::PathBuf);
impl TempPpt {
    fn new() -> Self {
        static NEXT: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
        let count = NEXT.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        let now = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        Self(std::env::temp_dir().join(format!(
            "litchi-native-auxiliary-{}-{now}-{count}.ppt",
            std::process::id()
        )))
    }
}
impl Drop for TempPpt {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}
fn reopen(p: &Package) -> Package {
    let file = TempPpt::new();
    p.save(&file.0).unwrap();
    let q = Package::open(&file.0).unwrap();
    assert_eq!(q.to_bytes().unwrap(), p.to_bytes().unwrap());
    q
}
fn sync_publish(p: &Package, c: &litchi_ppt::slide_sync::Commit) -> Package {
    let q = reopen(c.package());
    preserved(p, &q, RecordType::Slide);
    assert_eq!(
        p.apply_slide_sync(c.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        q.to_bytes().unwrap()
    );
    assert_eq!(
        q.apply_slide_sync(&c.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        p.to_bytes().unwrap()
    );
    let before = p.presentation().unwrap();
    let after = q.presentation().unwrap();
    let a = before.slide_at(0).unwrap().unwrap();
    let b = after.slide_at(0).unwrap().unwrap();
    let unowned = |r: &Record| {
        r.children
            .iter()
            .filter(|r| r.record_type != RecordType::RoundTripSlideSyncInfo12)
            .map(wire)
            .collect::<Vec<_>>()
    };
    assert_eq!(unowned(a.record()), unowned(b.record()));
    q
}

#[test]
fn native_ppt_slide_library_metadata_lifecycle_preserves_round_trip_records_and_streams() {
    use litchi_ppt::slide_sync::{Synchronization, SystemTime};
    for file in PARENTS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        assert!(
            p.presentation()
                .unwrap()
                .slide_at(0)
                .unwrap()
                .unwrap()
                .sync_info()
                .unwrap()
                .is_none()
        );
        let value = Synchronization::new(
            "native-slide-日本語-🍎",
            "http://example.invalid/library?slide=74&version=1",
            SystemTime::new(2024, 2, 4, 29, 12, 34, 56, 789).unwrap(),
            SystemTime::new(2026, 11, 6, 7, 10, 20, 30, 125).unwrap(),
        )
        .unwrap();
        let mut e = p.edit_slide_sync(Position::new(0)).unwrap();
        e.set(value.clone()).unwrap();
        assert_eq!(e.synchronization().unwrap(), Some(value.clone()));
        let c = e.commit().unwrap();
        let added = sync_publish(&p, &c);
        assert_eq!(
            added
                .presentation()
                .unwrap()
                .slide_at(0)
                .unwrap()
                .unwrap()
                .sync_info()
                .unwrap(),
            Some(&value)
        );
        let mut e = added.edit_slide_sync(Position::new(0)).unwrap();
        e.set(value.clone()).unwrap();
        assert!(e.commit().unwrap().patch().is_noop());
        let updated_value = Synchronization::new(
            "replacement-74",
            "http://example.invalid/library?slide=74&version=2",
            value.client_inserted(),
            value.server_modified(),
        )
        .unwrap();
        let mut e = added.edit_slide_sync(Position::new(0)).unwrap();
        e.set(updated_value.clone()).unwrap();
        let c = e.commit().unwrap();
        let updated = sync_publish(&added, &c);
        assert_eq!(
            updated
                .presentation()
                .unwrap()
                .slide_at(0)
                .unwrap()
                .unwrap()
                .sync_info()
                .unwrap(),
            Some(&updated_value)
        );
        assert!(matches!(
            updated.apply_slide_sync(c.patch()),
            Err(litchi_ppt::slide_sync::Error::PatchConflict)
        ));
        let mut e = updated.edit_slide_sync(Position::new(0)).unwrap();
        assert!(e.clear().unwrap());
        assert!(!e.clear().unwrap());
        let c = e.commit().unwrap();
        let removed = sync_publish(&updated, &c);
        assert!(
            removed
                .presentation()
                .unwrap()
                .slide_at(0)
                .unwrap()
                .unwrap()
                .sync_info()
                .unwrap()
                .is_none()
        );
        assert_eq!(persisted(&removed.to_bytes().unwrap()), persisted(&bytes));
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}

#[test]
fn native_ppt_slide_sync_invalid_values_and_missing_slide_keep_source_unchanged() {
    use litchi_ppt::slide_sync::{LibraryUrl, ServerId, SystemTime};
    let bytes = native(PARENTS[0]);
    let p = Package::from_bytes(&bytes).unwrap();
    assert!(
        matches!(LibraryUrl::new("https://example.invalid/slides"), Err(litchi_ppt::Error::InvalidFormat(m)) if m == "slide-library URL must use the HTTP scheme")
    );
    assert!(
        matches!(ServerId::new("bad\0id"), Err(litchi_ppt::Error::InvalidFormat(m)) if m == "server slide identifier contains a non-printable character")
    );
    assert!(
        matches!(SystemTime::new(2023, 2, 0, 29, 0, 0, 0, 0), Err(litchi_ppt::Error::InvalidFormat(m)) if m == "SYSTEMTIME contains an out-of-range field")
    );
    let e = p
        .edit_slide_sync(Position::new(p.presentation().unwrap().slide_count()))
        .unwrap();
    assert!(matches!(
        e.synchronization(),
        Err(litchi_ppt::slide_sync::Error::Refused(refusal))
            if refusal.kind() == litchi_core::edit::RefusalKind::NotFound
                && refusal.operation() == "publish"
                && refusal.reason() == "the presentation cannot publish this scoped edit"
    ));
    assert_eq!(p.to_bytes().unwrap(), bytes);
    assert_eq!(native(PARENTS[0]), bytes);
}

#[test]
fn native_ppt_photo_album_settings_read_all_layouts_and_frames_without_changing_pictures() {
    use litchi_ppt::{
        PhotoAlbumFrameShape as Frame, PhotoAlbumLayout as Layout, PhotoAlbumSettings,
    };
    let layouts = [
        Layout::FitToSlide,
        Layout::OnePicture,
        Layout::TwoPictures,
        Layout::FourPictures,
        Layout::OnePictureAndTitle,
        Layout::TwoPicturesAndTitle,
        Layout::FourPicturesAndTitle,
    ];
    let frames = [
        Frame::Rectangle,
        Frame::RoundedRectangle,
        Frame::Beveled,
        Frame::Oval,
        Frame::Octagon,
        Frame::Cross,
        Frame::Plaque,
    ];
    for file in PARENTS {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let doc = document(&p);
        let before =
            litchi_ppt::DocumentProperties10::parse(&doc, &SourceLimits::default()).unwrap();
        assert!(before.photo_album.is_none());
        for i in 0..7u8 {
            let expected = PhotoAlbumSettings {
                use_grayscale: i % 2 != 0,
                has_captions: i % 3 != 0,
                layout: layouts[usize::from(i)],
                unused: 0xa5,
                frame_shape: frames[usize::from(6 - i)],
            };
            let record = atom(
                RecordType::PhotoAlbumInfo10Atom,
                vec![
                    u8::from(expected.use_grayscale),
                    u8::from(expected.has_captions),
                    i,
                    0xa5,
                    6 - i,
                    0,
                ],
            );
            let tags = p
                .presentation()
                .unwrap()
                .programmable_tags()
                .unwrap()
                .unwrap();
            let mut extension = tags.document_extensions().unwrap().powerpoint10.unwrap();
            extension.photo_album_info = Some(record);
            let authored = tagged(doc.clone(), 10, extension.to_payload().unwrap());
            let q = seed_document(&p, &authored);
            let actual =
                litchi_ppt::DocumentProperties10::parse(&document(&q), &SourceLimits::default())
                    .unwrap();
            let mut expected_properties = before.clone();
            expected_properties.photo_album = Some(expected);
            assert_eq!(actual, expected_properties);
            assert_eq!(
                q.presentation()
                    .unwrap()
                    .programmable_tags()
                    .unwrap()
                    .unwrap()
                    .document_extensions()
                    .unwrap()
                    .powerpoint10
                    .unwrap(),
                extension
            );
            assert_eq!(q.metadata().unwrap(), p.metadata().unwrap());
        }
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}

#[test]
fn native_ppt_invalid_photo_album_enums_report_exact_reader_errors() {
    let bytes = native(PARENTS[0]);
    let p = Package::from_bytes(&bytes).unwrap();
    let doc = document(&p);
    for (data, message) in [
        (
            vec![0, 0, 7, 0, 0, 0],
            "PhotoAlbumInfo10Atom has an invalid layout",
        ),
        (
            vec![0, 0, 0, 0, 7, 0],
            "PhotoAlbumInfo10Atom has an invalid frame shape",
        ),
    ] {
        let mut ext = p
            .presentation()
            .unwrap()
            .programmable_tags()
            .unwrap()
            .unwrap()
            .document_extensions()
            .unwrap()
            .powerpoint10
            .unwrap();
        ext.photo_album_info = Some(atom(RecordType::PhotoAlbumInfo10Atom, data));
        let q = seed_document(&p, &tagged(doc.clone(), 10, ext.to_payload().unwrap()));
        assert!(
            matches!(litchi_ppt::DocumentProperties10::parse(&document(&q), &SourceLimits::default()), Err(litchi_ppt::Error::Corrupted(m)) if m == message)
        );
    }
    assert_eq!(p.to_bytes().unwrap(), bytes);
    assert_eq!(native(PARENTS[0]), bytes);
}

fn envelope() -> litchi_ppt::EnvelopeData {
    use litchi_ppt::envelope_data::*;
    let text = |s: &str| MsoEnvelopeText::Unicode(s.encode_utf16().collect());
    EnvelopeData {
        clsid: MSO_ENVELOPE_CLSID,
        payload: EnvelopePayload::Mso(MsoEnvelope {
            version: MsoEnvelopeVersion::Office8,
            last_sent_time: 0,
            flag_status: MsoFollowUpStatus::Flagged,
            reply_time: 1,
            request: text("Reply 日本語"),
            sent_representing_entry_id: vec![1, 2, 3],
            sent_representing_name: text("Native Sender"),
            internet_account_stamp: text("inert-stamp"),
            internet_account_name: text("account"),
            expiry_time: 2,
            deferred_delivery_time: 0,
            delete_after_submit: false,
            security: MsoSecurityFlags {
                signed: true,
                encrypted: false,
            },
            delivery_report: true,
            read_receipt: false,
            categories: text("fixture"),
            sensitivity: MsoSensitivity::Private,
            importance: MsoImportance::High,
            subject: text("Stored envelope 🍎"),
            voting_options: b"yes;no".to_vec(),
            reply_recipients: MsoRecipientCollection {
                recipients: vec![MsoRecipientProperties {
                    properties: vec![MsoRecipientProperty {
                        property_id: 0x3001,
                        value: MsoPropertyValue::Unicode("Recipient".encode_utf16().collect()),
                    }],
                }],
            },
            contact_link_recipients: Some(MsoRecipientCollection::default()),
            recipients: MsoRecipientCollection::default(),
            attachments: vec![MsoAttachment {
                method: 1,
                name: "saved.txt".encode_utf16().collect(),
                data: b"inert attachment".to_vec(),
            }],
            intro_text: Some("Stored intro".encode_utf16().collect()),
        }),
    }
}

#[test]
fn native_ppt_privacy_web_routing_and_envelope_settings_stay_inert_through_metadata_edits() {
    use litchi_ppt::html_publish::{CodePage, WebFrameColors, WebOutput, WebScreenSize};
    use litchi_ppt::routing_slip::{Address, Slip, Text};
    use litchi_ppt::{
        EnvelopeSettings, HtmlDocumentSettings, HtmlPublishSettings, PrivacySettings,
    };
    let html = HtmlDocumentSettings {
        encoding: CodePage::new(65001),
        frame_colors: WebFrameColors::PresentationAccent,
        screen_size: WebScreenSize::Pixels1024x768,
        output: WebOutput::Dual,
        show_frame: true,
        resize_graphics: false,
        organize_in_folder: true,
        use_long_file_names: true,
        rely_on_vml: false,
        allow_png: true,
        show_slide_animation: true,
    };
    let publish = HtmlPublishSettings {
        file_name: "file:///never-created/native.html".into(),
        named_show: None,
        start_slide: 1,
        end_slide: 1,
        output: WebOutput::Html4,
        use_slide_range: true,
        use_named_show: false,
        load_in_browser: true,
        show_speaker_notes: true,
    };
    let flags = EnvelopeSettings::new(true, true, true).unwrap();
    let envelope = envelope();
    let privacy = PrivacySettings {
        remove_personally_identifiable_information_on_save: true,
    };
    let text = |s: &[u8]| Text::from_ansi_bytes(s.to_vec()).unwrap();
    let mut slip = Slip::new(
        Address::new(text(b"originator@example.invalid")),
        vec![Address::new(text(b"recipient@example.invalid"))],
        text(b"Saved routing slip"),
        text(b"Inert message \x80"),
    )
    .unwrap();
    slip.track_status = true;
    slip.return_when_done = true;
    slip.unused1 = 0x12345678;
    slip.unused2 = 0xabcdef01;
    slip.trailing_undefined = vec![0x74; 12];
    for file in [PARENTS[0], PARENTS[3]] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let tags = p
            .presentation()
            .unwrap()
            .programmable_tags()
            .unwrap()
            .unwrap();
        let extensions = tags.document_extensions().unwrap();
        let mut pp9 = extensions.powerpoint9.unwrap_or_default();
        assert!(
            pp9.html_doc_info.is_none()
                && pp9.html_publish_info.is_none()
                && pp9.envelope_flags.is_none()
                && pp9.envelope_data.is_none()
        );
        pp9.html_doc_info = Some(html.to_record());
        pp9.html_publish_info = Some(publish.to_record().unwrap());
        pp9.envelope_flags = Some(flags.to_record().unwrap());
        pp9.envelope_data = Some(envelope.to_record().unwrap());
        let mut pp10 = extensions.powerpoint10.unwrap();
        assert!(pp10.filter_privacy_flags.is_none());
        pp10.filter_privacy_flags = Some(privacy.to_record());
        let doc = tagged(document(&p), 9, pp9.to_payload().unwrap());
        let mut doc = tagged(doc, 10, pp10.to_payload().unwrap());
        assert!(
            !doc.children
                .iter()
                .any(|r| r.record_type == RecordType::DocRoutingSlipAtom)
        );
        let insert = doc
            .children
            .iter()
            .position(|r| {
                matches!(
                    r.record_type,
                    RecordType::PrintOptionsAtom | RecordType::EndDocument
                )
            })
            .unwrap();
        doc.children.insert(insert, slip.to_record().unwrap());
        rebuild(&mut doc);
        let seeded = seed_document(&p, &doc);
        let assert_values = |p: &Package| {
            let presentation = p.presentation().unwrap();
            assert_eq!(presentation.privacy_settings().unwrap(), Some(privacy));
            assert_eq!(presentation.html_document_settings().unwrap(), Some(html));
            assert_eq!(
                presentation.html_publish_settings().unwrap(),
                Some(publish.clone())
            );
            assert_eq!(presentation.envelope_settings().unwrap(), Some(flags));
            assert_eq!(
                presentation.envelope_data().unwrap(),
                Some(envelope.clone())
            );
            assert_eq!(presentation.routing_slip().unwrap(), Some(slip.clone()));
        };
        assert_values(&seeded);
        assert_eq!(seeded.metadata().unwrap(), p.metadata().unwrap());
        let mut edit = seeded.edit_metadata().unwrap();
        edit.merge_metadata(&litchi_core::Metadata {
            title: Some("Metadata retained with privacy preference".into()),
            ..Default::default()
        })
        .unwrap();
        let c = edit.commit().unwrap();
        let q = reopen(c.package());
        assert_values(&q);
        let before = streams(&seeded.to_bytes().unwrap());
        let after = streams(&q.to_bytes().unwrap());
        assert_eq!(
            before.keys().collect::<Vec<_>>(),
            after.keys().collect::<Vec<_>>()
        );
        for (name, data) in before {
            if name != ["\u{5}SummaryInformation"] {
                assert_eq!(after[&name], data, "metadata edit changed {name:?}");
            }
        }
        let mut expected = seeded.metadata().unwrap();
        expected.title = Some("Metadata retained with privacy preference".into());
        assert_eq!(q.metadata().unwrap(), expected);
        assert_eq!(
            q.apply_metadata(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            seeded.to_bytes().unwrap()
        );
        assert_eq!(
            seeded
                .apply_metadata(c.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            q.to_bytes().unwrap()
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}

fn broadcast() -> litchi_ppt::Broadcast {
    use litchi_ppt::{Broadcast, BroadcastProperties, slide_sync::SystemTime};
    Broadcast {
        title: Some("Stored broadcast 日本語".into()),
        description: Some("Inert session".into()),
        speaker: Some("Speaker".into()),
        contact: Some("Contact".into()),
        remote_server_name: Some("CAMERA01".into()),
        email_address: Some("feedback@example.invalid".into()),
        email_name: Some("Feedback".into()),
        chat_url: Some("http://example.invalid/chat?id=74".into()),
        archive_directory: Some("C:\\NeverOpened".into()),
        netshow_files_base_directory: Some("\\\\server\\share".into()),
        netshow_files_directory: Some("\\\\server\\share\\netshow".into()),
        netshow_server_name: Some("NETSHOW01".into()),
        ppt_files_base_directory: "\\\\server\\share".into(),
        ppt_files_directory: "\\\\server\\share\\ppt".into(),
        ppt_files_base_url: "http://example.invalid/base".into(),
        user_name: "scheduler".into(),
        broadcast_date_time: "2026-11-07T09-30".into(),
        presentation_name: "native.ppt".into(),
        asd_file_name: "\\\\server\\share\\stream.asd".into(),
        entry_id: Some("inert-calendar-74".into()),
        properties: BroadcastProperties {
            send_audio: true,
            send_video: true,
            camera_remote: true,
            use_netshow: true,
            use_other_server: false,
            can_email: true,
            can_chat: true,
            archive: true,
            speaker_notes: true,
            quarter_screen: false,
            show_tools: true,
            record_only: false,
            start_time: SystemTime::new(2026, 11, 6, 7, 9, 30, 15, 125).unwrap(),
            end_time: SystemTime::new(2026, 11, 6, 7, 10, 30, 15, 125).unwrap(),
        },
    }
}

#[test]
fn native_ppt_broadcast_metadata_edits_preserve_opaque_records_and_native_content() {
    for file in [PARENTS[0], PARENTS[3]] {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let expected = broadcast();
        let mut record = Record::parse(
            &wire(&expected.to_record().unwrap()),
            0,
            &SourceLimits::default(),
        )
        .unwrap()
        .0;
        let mut offset = 0;
        while offset < record.data.len() {
            let (child, consumed) =
                Record::parse(&record.data, offset, &SourceLimits::default()).unwrap();
            assert!(consumed >= 8);
            record.children.push(child);
            offset += consumed;
        }
        assert_eq!(offset, record.data.len());
        let mut opaque = atom(RecordType::Unknown, vec![0x74, 0, 0xff, 0xa5]);
        opaque.record_type_raw = 0x7e74;
        opaque.instance = 7;
        record.children.insert(1, opaque.clone());
        rebuild(&mut record);
        let mut pp9 = p
            .presentation()
            .unwrap()
            .programmable_tags()
            .unwrap()
            .unwrap()
            .document_extensions()
            .unwrap()
            .powerpoint9
            .unwrap_or_default();
        assert!(pp9.broadcasts.is_empty());
        pp9.broadcasts.push(record);
        let seeded = seed_document(&p, &tagged(document(&p), 9, pp9.to_payload().unwrap()));
        assert_eq!(
            seeded
                .presentation()
                .unwrap()
                .broadcasts()
                .unwrap()
                .broadcasts
                .as_slice(),
            std::slice::from_ref(&expected)
        );
        let mut e = seeded.edit_broadcast(Position::new(0)).unwrap();
        e.replace(expected.clone()).unwrap();
        assert!(e.commit().unwrap().patch().is_noop());
        let mut e = seeded.edit_broadcast(Position::new(0)).unwrap();
        assert_eq!(e.unknown_records().unwrap().len(), 1);
        assert_eq!(
            e.unknown_records().unwrap()[0].to_record_bytes().unwrap(),
            wire(&opaque)
        );
        e.set_title(Some("Updated inert broadcast 🍎".into()))
            .unwrap();
        let staged = e.broadcast().unwrap().clone();
        let mut invalid = staged.clone();
        invalid.remote_server_name = None;
        assert!(
            matches!(e.replace(invalid), Err(litchi_ppt::broadcast::Error::Package(litchi_ppt::Error::Corrupted(m))) if m == "Remote-camera broadcast is missing BCRexServerNameAtom")
        );
        assert_eq!(e.broadcast().unwrap(), &staged);
        let c = e.commit().unwrap();
        let q = reopen(c.package());
        preserved(&seeded, &q, RecordType::Document);
        assert_eq!(
            q.presentation().unwrap().broadcasts().unwrap().broadcasts,
            [staged]
        );
        assert_eq!(
            q.edit_broadcast(Position::new(0))
                .unwrap()
                .unknown_records()
                .unwrap()[0]
                .to_record_bytes()
                .unwrap(),
            wire(&opaque)
        );
        let old_doc = document(&seeded);
        let new_doc = document(&q);
        let unowned = |doc: &Record| {
            doc.children
                .iter()
                .filter(|r| r.record_type != RecordType::DocInfoList)
                .map(wire)
                .collect::<Vec<_>>()
        };
        assert_eq!(unowned(&old_doc), unowned(&new_doc));
        let before_tags = seeded
            .presentation()
            .unwrap()
            .programmable_tags()
            .unwrap()
            .unwrap();
        let after_tags = q
            .presentation()
            .unwrap()
            .programmable_tags()
            .unwrap()
            .unwrap();
        assert_eq!(before_tags.tags.len(), after_tags.tags.len());
        for (a, b) in before_tags.tags.iter().zip(&after_tags.tags) {
            if !matches!(a, litchi_ppt::ProgTag::Binary(tag) if tag.version == litchi_ppt::ProgBinaryTagVersion::PowerPoint9)
            {
                assert_eq!(a, b);
            }
        }
        assert_eq!(
            seeded
                .apply_broadcast(c.patch())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            q.to_bytes().unwrap()
        );
        assert_eq!(
            q.apply_broadcast(&c.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            seeded.to_bytes().unwrap()
        );
        assert!(matches!(
            q.apply_broadcast(c.patch()),
            Err(litchi_ppt::broadcast::Error::PatchConflict)
        ));
        assert_eq!(p.to_bytes().unwrap(), bytes);
        assert_eq!(native(file), bytes);
    }
}
