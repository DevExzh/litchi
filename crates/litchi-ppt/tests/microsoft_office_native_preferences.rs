//! Native PowerPoint preferences, programmable tags and CFB property edits.
use litchi_core::{Metadata, Position};
use litchi_ole_common::property_set::summary_properties::MetadataError;
use litchi_ole_common::property_set::{PropertySetReader, SummaryProperties, Value};
use litchi_ppt::{
    NormalViewSetPayload, Package, ProgBinaryTagVersion, ProgTag, ProgTagLimits, ProgTagScope,
    ProgTags, Record, RecordParser, RecordType,
};
use std::{collections::BTreeMap, io::Cursor};

const VISUAL: [&str; 3] = [
    "powerpoint-visual-objects.ppt",
    "powerpoint-visual-objects.pps",
    "powerpoint-visual-objects.pot",
];
const HANDOUT: &str = "powerpoint-fonts-math-handout.ppt";
const NAV: &str = "powerpoint-navigation-review.ppt";
const MACRO: &str = "powerpoint-passive-macro.ppt";
struct TempPpt(std::path::PathBuf);
impl TempPpt {
    fn new() -> Self {
        static NEXT: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
        let sequence = NEXT.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
        let nonce = std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let path = std::env::temp_dir().join(format!(
            "litchi-native-preferences-{}-{nonce}-{sequence}.ppt",
            std::process::id()
        ));
        std::fs::OpenOptions::new()
            .write(true)
            .create_new(true)
            .open(&path)
            .unwrap();
        Self(path)
    }
    fn path(&self) -> &std::path::Path {
        &self.0
    }
}
impl Drop for TempPpt {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}

fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/powerpoint")
            .join(name),
    )
    .unwrap()
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut cfb =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    cfb.list_streams()
        .into_iter()
        .map(|path| {
            let names = path.iter().map(String::as_str).collect::<Vec<_>>();
            let bytes = cfb.open_stream(&names).unwrap();
            (path, bytes)
        })
        .collect()
}
fn source_document(bytes: &[u8]) -> Record {
    let parts = streams(bytes);
    let stream = &parts[&vec!["PowerPoint Document".to_owned()]];
    let mut parser = RecordParser::new();
    parser
        .parse_document(stream, &litchi_ppt::SourceLimits::default())
        .unwrap();
    let records = parser.find_records_ref();
    let docs = records
        .iter()
        .filter(|r| r.record_type == RecordType::Document)
        .collect::<Vec<_>>();
    // These immutable native inputs each contain one Document. Edited
    // incremental histories are tested through the public live-document APIs.
    assert_eq!(docs.len(), 1);
    (**docs[0]).clone()
}
fn encoded(record: &Record) -> Vec<u8> {
    let mut out = ((record.instance << 4) | record.version)
        .to_le_bytes()
        .to_vec();
    out.extend_from_slice(&record.record_type_raw.to_le_bytes());
    out.extend_from_slice(&record.data_length.to_le_bytes());
    out.extend_from_slice(&record.data);
    out
}
fn properties(p: &Package) -> SummaryProperties {
    p.edit_metadata().unwrap().properties().unwrap().clone()
}
fn custom(p: &Package) -> BTreeMap<String, Value> {
    litchi_cfb::OleFile::open(
        Cursor::new(p.to_bytes().unwrap()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap()
    .get_metadata()
    .unwrap()
    .custom_properties
    .into_iter()
    .collect()
}
fn assert_unrelated_streams(before: &[u8], after: &[u8], mutable: &[&str]) {
    let a = streams(before);
    let b = streams(after);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (name, bytes) in a {
        if !matches!(name.as_slice(),[n] if mutable.contains(&n.as_str())) {
            assert_eq!(bytes, b[&name], "{}", name.join("/"));
        }
    }
}
fn reopen_metadata(source: &Package, commit: &litchi_ppt::package::metadata::Commit) -> Package {
    let file = TempPpt::new();
    commit.package().save(file.path()).unwrap();
    let p = Package::open(file.path()).unwrap();
    let original = source.to_bytes().unwrap();
    let after = p.to_bytes().unwrap();
    assert_unrelated_streams(
        &original,
        &after,
        &["\u{5}SummaryInformation", "\u{5}DocumentSummaryInformation"],
    );
    assert_eq!(
        properties(source).document_summary_information(),
        properties(&p).document_summary_information()
    );
    assert_eq!(
        source
            .apply_metadata(commit.patch())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        after
    );
    assert!(matches!(
        p.apply_metadata(commit.patch()),
        Err(litchi_ppt::package::metadata::Error::PatchConflict)
    ));
    assert_eq!(
        p.apply_metadata(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap(),
        original
    );
    p
}

#[test]
fn native_slide_notes_and_opaque_pane_preferences_keep_variant_differences() {
    for (name, scale, x, notes_view, notes_text, normal) in [
        (
            VISUAL[0],
            Some(104),
            1440,
            false,
            true,
            [
                149, 105, 176, 246, 0, 202, 154, 59, 171, 172, 135, 199, 0, 202, 154, 59, 1, 1, 0,
                0,
            ],
        ),
        (
            VISUAL[1],
            Some(104),
            1440,
            false,
            true,
            [
                47, 232, 177, 246, 0, 202, 154, 59, 171, 172, 135, 199, 0, 202, 154, 59, 1, 1, 0, 0,
            ],
        ),
        (
            VISUAL[2],
            None,
            0,
            false,
            false,
            [
                47, 232, 177, 246, 0, 202, 154, 59, 171, 172, 135, 199, 0, 202, 154, 59, 1, 1, 0, 0,
            ],
        ),
        (
            HANDOUT,
            Some(104),
            1440,
            true,
            true,
            [
                47, 232, 177, 246, 0, 202, 154, 59, 209, 166, 147, 199, 0, 202, 154, 59, 1, 1, 0, 0,
            ],
        ),
        (
            NAV,
            Some(105),
            1384,
            false,
            true,
            [
                149, 105, 176, 246, 0, 202, 154, 59, 255, 248, 235, 198, 0, 202, 154, 59, 1, 1, 0,
                0,
            ],
        ),
    ] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let d = p.presentation().unwrap();
        let views = d.slide_view_information().unwrap();
        let slide = views.slide().unwrap();
        assert!(!slide.preferences().snap_to_grid() && !slide.preferences().snap_to_shape());
        assert!(slide.guides().is_empty());
        assert_eq!(slide.zoom().is_some(), scale.is_some());
        if let Some(numerator) = scale {
            let z = slide.zoom().unwrap();
            assert_eq!(
                (
                    z.x_scale().numerator(),
                    z.x_scale().denominator(),
                    z.y_scale().numerator(),
                    z.y_scale().denominator()
                ),
                (numerator, 100, numerator, 100)
            );
            assert_eq!((z.origin().x(), z.origin().y()), (x, 192));
            assert!(z.uses_variable_scale() && !z.is_draft_mode());
        }
        assert_eq!(views.notes().is_some(), notes_view);
        if notes_view {
            let notes = views.notes().unwrap();
            let zoom = notes.zoom().unwrap();
            assert_eq!(
                (
                    zoom.x_scale().numerator(),
                    zoom.x_scale().denominator(),
                    zoom.y_scale().numerator(),
                    zoom.y_scale().denominator()
                ),
                (111, 100, 111, 100)
            );
            assert_eq!((zoom.origin().x(), zoom.origin().y()), (3240, 200));
            assert!(notes.guides().is_empty());
        }
        let text = d.notes_text_view_info().unwrap();
        assert_eq!(text.is_some(), notes_text);
        if let Some(text) = text {
            let v = text.view_info();
            assert_eq!(
                (
                    v.x_scale().numerator(),
                    v.x_scale().denominator(),
                    v.y_scale().numerator(),
                    v.y_scale().denominator()
                ),
                (1, 1, 1, 1)
            );
            assert_eq!((v.origin().x(), v.origin().y()), (0, 0));
            assert!(!v.draft_mode());
        }
        let pane = d.normal_view_set_info().unwrap().unwrap();
        assert_eq!(
            pane.payload(),
            &NormalViewSetPayload::Other(normal.to_vec())
        );
        assert!(pane.layout().is_none());
        let doc = source_document(&bytes);
        assert!(litchi_ppt::PrintOptions::parse(&doc).unwrap().is_none());
        let info = doc.find_child(RecordType::DocInfoList).unwrap();
        assert_eq!(
            pane.to_bytes().unwrap(),
            encoded(info.find_child(RecordType::NormalViewSetInfo9).unwrap())
        );
        for view in [views.slide(), views.notes()].into_iter().flatten() {
            let wire = info
                .children
                .iter()
                .find(|r| {
                    r.record_type == RecordType::SlideViewInfo
                        && r.instance
                            == if view.kind() == litchi_ppt::ViewKind::Slide {
                                0
                            } else {
                                1
                            }
                })
                .unwrap();
            assert_eq!(view.to_bytes().unwrap(), encoded(wire));
        }
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_views_and_tags_survive_incremental_reorder_save_reopen_and_inverse() {
    for name in [VISUAL[0], VISUAL[1], VISUAL[2], HANDOUT, NAV] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let before = p.presentation().unwrap();
        let views = before.slide_view_information().unwrap();
        let normal = before.normal_view_set_info().unwrap();
        let notes = before.notes_text_view_info().unwrap();
        let tags = before.programmable_tags().unwrap();
        let original_text = before
            .slides()
            .unwrap()
            .iter()
            .map(|s| s.text().unwrap().to_owned())
            .collect::<Vec<_>>();
        let mut edit = p.edit().unwrap();
        edit.move_slide(Position::new(0), Position::new(before.slide_count() - 1))
            .unwrap();
        let commit = edit.commit().unwrap();
        let file = TempPpt::new();
        commit.package().save(file.path()).unwrap();
        let reopened = Package::open(file.path()).unwrap();
        let after = reopened.presentation().unwrap();
        let changed = after.slide_view_information().unwrap();
        assert_eq!(changed.slide(), views.slide());
        assert_eq!(changed.notes(), views.notes());
        assert_eq!(after.normal_view_set_info().unwrap(), normal);
        assert_eq!(after.notes_text_view_info().unwrap(), notes);
        assert_eq!(after.programmable_tags().unwrap(), tags);
        let outline = after.outline_sorter_view_information().unwrap();
        assert!(outline.outline().is_none() && outline.sorter().is_none());
        let mut expected = original_text;
        expected.rotate_left(1);
        assert_eq!(
            after
                .slides()
                .unwrap()
                .iter()
                .map(|s| s.text().unwrap().to_owned())
                .collect::<Vec<_>>(),
            expected
        );
        let output = reopened.to_bytes().unwrap();
        assert_unrelated_streams(&bytes, &output, &["PowerPoint Document", "Current User"]);
        let a = streams(&bytes);
        let b = streams(&output);
        let key = vec!["PowerPoint Document".to_owned()];
        assert!(b[&key].starts_with(&a[&key]));
        assert_eq!(
            reopened
                .apply(&commit.patch().inverse())
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            bytes
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_programmable_tag_order_and_payloads_round_trip_without_interpretation() {
    for name in [VISUAL[0], VISUAL[1], VISUAL[2], HANDOUT, NAV] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let d = p.presentation().unwrap();
        let tags = d.programmable_tags().unwrap().unwrap();
        assert_eq!(tags.scope, ProgTagScope::Document);
        assert_eq!(tags.instance, 0);
        let versions = tags
            .tags
            .iter()
            .map(|t| match t {
                ProgTag::Binary(t) => t.version,
                _ => panic!("unexpected string tag"),
            })
            .collect::<Vec<_>>();
        let mut expected = vec![
            ProgBinaryTagVersion::PowerPoint12,
            ProgBinaryTagVersion::PowerPoint10,
        ];
        if name == NAV {
            expected.push(ProgBinaryTagVersion::PowerPoint9);
        }
        assert_eq!(versions, expected);
        let doc = source_document(&bytes);
        let raw = doc
            .find_child(RecordType::DocInfoList)
            .unwrap()
            .find_child(RecordType::ProgTags)
            .unwrap();
        assert_eq!(
            tags.to_bytes(ProgTagLimits::default()).unwrap(),
            encoded(raw)
        );
        assert_eq!(
            ProgTags::parse(
                &tags.to_record(ProgTagLimits::default()).unwrap(),
                ProgTagScope::Document,
                ProgTagLimits::default()
            )
            .unwrap(),
            tags
        );
        let pp12 = tags.binary_tag(ProgBinaryTagVersion::PowerPoint12).unwrap();
        assert_eq!(pp12.payload, [0, 0, 37, 4, 1, 0, 0, 0, 0]);
        let pp10 = tags.binary_tag(ProgBinaryTagVersion::PowerPoint10).unwrap();
        let mut payload = vec![0, 0, 13, 4, 8, 0, 0, 0, 0, 192, 0, 0, 0, 192, 0, 0];
        if name == HANDOUT {
            payload.extend_from_slice(&[0, 0, 200, 50, 4, 0, 0, 0, 1, 0, 0, 0]);
        }
        assert_eq!(pp10.payload, payload);
        let ext = tags.document_extensions().unwrap();
        assert_eq!(
            ext.powerpoint12.unwrap().to_payload().unwrap(),
            pp12.payload
        );
        assert_eq!(
            ext.powerpoint10.unwrap().to_payload().unwrap(),
            pp10.payload
        );
        assert!(ext.powerpoint11.is_none());
        let props =
            litchi_ppt::DocumentProperties10::parse(&doc, &litchi_ppt::SourceLimits::default())
                .unwrap();
        assert_eq!(props.grid_spacing.unwrap().grid_units, 49152);
        assert_eq!(
            d.powerpoint12_document_properties()
                .unwrap()
                .compress_pictures_on_save,
            Some(false)
        );
        if name == NAV {
            let pp9 = tags.binary_tag(ProgBinaryTagVersion::PowerPoint9).unwrap();
            let extension = ext.powerpoint9.unwrap();
            assert_eq!(extension.external_hyperlinks.len(), 8);
            assert_eq!(
                extension
                    .external_hyperlinks
                    .iter()
                    .map(|r| (r.record_type_raw, r.data_length))
                    .collect::<Vec<_>>(),
                [
                    (4068, 72),
                    (4068, 24),
                    (4068, 24),
                    (4068, 24),
                    (4068, 72),
                    (4068, 24),
                    (4068, 24),
                    (4068, 24)
                ]
            );
            assert_eq!(extension.to_payload().unwrap(), pp9.payload);
        }
        let mut duplicate = tags.clone();
        duplicate.tags.push(duplicate.tags[0].clone());
        assert!(matches!(duplicate.to_bytes(ProgTagLimits::default()),
            Err(litchi_ppt::Error::Corrupted(message)) if message == "ProgTags container contains a duplicate versioned binary tag"));
        let limits = ProgTagLimits {
            max_binary_payload_bytes: 8,
            ..ProgTagLimits::default()
        };
        assert!(
            matches!(tags.to_bytes(limits), Err(litchi_ppt::Error::Corrupted(message))
            if message == "binary programmable tag payload exceeds its configured limit")
        );
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}

#[test]
fn native_powerpoint_mac_summary_dates_and_saved_counts_keep_variant_precision() {
    for (name, created, modified, words, revision, template, last_printed) in [
        (
            VISUAL[0],
            "2026-10-15T06:15:44.602Z",
            "2026-10-15T06:23:11.830Z",
            49,
            "1",
            None,
            None,
        ),
        (
            VISUAL[1],
            "2026-10-15T06:15:44Z",
            "2026-10-15T06:59:26.712Z",
            49,
            "1",
            Some(""),
            None,
        ),
        (
            VISUAL[2],
            "2026-10-15T06:15:44Z",
            "2026-10-15T06:59:12.715Z",
            49,
            "1",
            Some(""),
            None,
        ),
        (
            HANDOUT,
            "2026-10-15T10:08:38Z",
            "2026-10-15T11:34:21.242Z",
            50,
            "3",
            None,
            Some("2026-10-15T11:33:27.083Z"),
        ),
        (
            MACRO,
            "2026-10-15T07:45:54.224Z",
            "2026-10-15T07:53:24.305Z",
            17,
            "1",
            Some(""),
            None,
        ),
    ] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let m = p.metadata().unwrap();
        assert_eq!(m.title.as_deref(), Some(""));
        assert_eq!(m.company.as_deref(), Some(""));
        assert_eq!(m.author.as_deref(), Some("Ryker Zhu"));
        assert_eq!(m.last_modified_by, m.author);
        assert_eq!(
            m.application.as_deref(),
            Some("Microsoft Macintosh PowerPoint")
        );
        assert_eq!(m.codepage, Some(10000));
        assert_eq!(m.revision.as_deref(), Some(revision));
        assert_eq!(m.word_count, Some(words));
        assert_eq!(m.template.as_deref(), template);
        assert_eq!(m.created.unwrap().to_string(), created);
        assert_eq!(m.modified.unwrap().to_string(), modified);
        assert_eq!(
            m.last_printed.map(|t| t.to_string()).as_deref(),
            last_printed
        );
        assert_eq!(
            p.edit_metadata()
                .unwrap()
                .commit()
                .unwrap()
                .package()
                .to_bytes()
                .unwrap(),
            bytes
        );
    }
}

#[test]
fn native_summary_property_crud_preserves_presentation_views_tags_and_macro_payloads() {
    for name in [VISUAL[0], VISUAL[1], VISUAL[2], HANDOUT, MACRO] {
        let bytes = native(name);
        let p = Package::from_bytes(&bytes).unwrap();
        let original = p.metadata().unwrap();
        let original_custom = custom(&p);
        let mut edit = p.edit_metadata().unwrap();
        edit.merge_metadata(&Metadata {
            title: Some("Reviewed native presentation".into()),
            ..Metadata::default()
        })
        .unwrap();
        edit.properties_mut()
            .unwrap()
            .set_user_defined("NativeReviewed", Value::Bool(true))
            .unwrap();
        let staged = edit.properties().unwrap().clone();
        assert!(matches!(
            edit.merge_metadata(&Metadata {
                title: Some("must not be staged".into()),
                identifier: Some("urn:unsupported".into()),
                ..Metadata::default()
            }),
            Err(litchi_ppt::package::metadata::Error::Metadata(
                MetadataError::Unsupported {
                    field: "identifier"
                }
            ))
        ));
        assert_eq!(edit.properties().unwrap(), &staged);
        let commit = edit.commit().unwrap();
        let after = reopen_metadata(&p, &commit);
        let mut expected = original.clone();
        expected.title = Some("Reviewed native presentation".into());
        assert_eq!(after.metadata().unwrap(), expected);
        let mut expected_custom = original_custom.clone();
        expected_custom.insert("NativeReviewed".into(), Value::Bool(true));
        assert_eq!(custom(&after), expected_custom);
        let mut edit = after.edit_metadata().unwrap();
        edit.properties_mut()
            .unwrap()
            .set_user_defined("NativeReviewed", Value::Bool(false))
            .unwrap();
        let updated = edit.commit().unwrap();
        let updated = reopen_metadata(&after, &updated);
        assert_eq!(custom(&updated)["NativeReviewed"], Value::Bool(false));
        let mut edit = updated.edit_metadata().unwrap();
        assert!(
            edit.properties_mut()
                .unwrap()
                .remove("NativeReviewed")
                .unwrap()
        );
        assert!(edit.properties_mut().unwrap().remove("title").unwrap());
        let removed = edit.commit().unwrap();
        let removed = reopen_metadata(&updated, &removed);
        let mut expected = original;
        expected.title = None;
        assert_eq!(removed.metadata().unwrap(), expected);
        assert_eq!(custom(&removed), original_custom);
        assert_eq!(p.to_bytes().unwrap(), bytes);
    }
}
