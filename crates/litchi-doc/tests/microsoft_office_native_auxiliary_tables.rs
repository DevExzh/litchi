//! Real producer caches/recognizer state and empty review-thread slots, kept inert.
use litchi_doc::parts::{
    proofing::{ProofingFeature, ProofingStateTable},
    smart_tags::{self, SmartTagRecognizerState as State},
};
use litchi_doc::{OpenOptions, Package, Password};
use litchi_ole_common::object;
use std::{collections::BTreeMap, io::Cursor, path::Path};
#[path = "../../../tools/test_support/cfb_directory_metadata.rs"]
mod cfb_metadata;
fn root() -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/microsoft-office/word")
}
fn native(name: &str) -> Vec<u8> {
    std::fs::read(root().join(name)).unwrap()
}
fn open(name: &str, bytes: &[u8]) -> Package {
    if name == "word-encrypted.doc" {
        Package::from_bytes_with_options(
            bytes,
            OpenOptions::default().password(Password::new("Office42!".into())),
        )
        .unwrap()
    } else {
        Package::from_bytes(bytes).unwrap()
    }
}
fn streams(bytes: &[u8]) -> BTreeMap<Vec<String>, Vec<u8>> {
    let mut ole =
        litchi_cfb::OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default())
            .unwrap();
    ole.list_streams()
        .into_iter()
        .map(|p| {
            let b = ole
                .open_stream(&p.iter().map(String::as_str).collect::<Vec<_>>())
                .unwrap();
            (p, b)
        })
        .collect()
}
fn table(p: &Package) -> (Vec<String>, Vec<u8>) {
    let path = vec![format!(
        "{}Table",
        u8::from(p.document().fib().which_table_stream())
    )];
    let mut all = streams(&p.to_bytes().unwrap());
    let data = all.remove(&path).unwrap();
    (path, data)
}
fn range(p: &Package, index: usize) -> std::ops::Range<usize> {
    let (at, len) = p.document().fib().get_table_pointer(index).unwrap();
    usize::try_from(at).unwrap()..usize::try_from(at + len).unwrap()
}
const CACHES: [(&str, u32); 20] = [
    ("word-text-structure.doc", 399),
    ("word-tables-layout.doc", 434),
    ("word-review-controls.doc", 257),
    ("word-recorded-macro.doc", 181),
    ("word-recorded-macro.dot", 181),
    ("word-math-navigation.doc", 653),
    ("word-visual-objects.doc", 186),
    ("word-mailmerge-fields.doc", 403),
    ("word-mailmerge-source.doc", 213),
    ("word-document-protection.doc", 192),
    ("word-legacy-form.doc", 190),
    ("word-embedded-fonts.doc", 200),
    ("word-vector-diagram.doc", 157),
    ("word-glossary.dot", 121),
    ("word-bidi-eastasia.doc", 246),
    ("word-layout-extensions.doc", 379),
    ("word-layout-extensions.dot", 379),
    ("word-master-native.doc", 81),
    ("word-tracked-moves.doc", 289),
    ("word-encrypted.doc", 116),
];
#[test]
fn native_word_table_character_caches_retain_exact_sentinels_and_flags() {
    for (file, end) in CACHES {
        let bytes = native(file);
        let p = open(file, &bytes);
        let d = p.document();
        let cache = d.table_character_cache().unwrap().unwrap();
        assert_eq!(d.fib().get_main_doc_range().1, end);
        assert_eq!(d.fib().get_table_pointer(93).unwrap().1, 20);
        assert_eq!(
            cache
                .entries()
                .iter()
                .map(|e| (e.start_cp(), e.info().is_unknown()))
                .collect::<Vec<_>>(),
            [(0, false), (end, true)]
        );
        assert_eq!(cache.terminal_cp(), end + 2);
        let expected = [0, end, end + 2, 0, 1]
            .into_iter()
            .flat_map(u32::to_le_bytes)
            .collect::<Vec<_>>();
        assert_eq!(cache.to_bytes().unwrap(), expected);
        if file != "word-encrypted.doc" {
            assert_eq!(table(&p).1[range(&p, 93)], expected);
        }
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_word_revision_thread_tables_keep_empty_message_and_style_slots() {
    for (file, _) in CACHES {
        let bytes = native(file);
        let p = open(file, &bytes);
        let d = p.document();
        let threading = d.rmd_threading().unwrap();
        let count = if matches!(file, "word-review-controls.doc" | "word-tracked-moves.doc") {
            2
        } else {
            1
        };
        assert_eq!(threading.messages().len(), count);
        assert_eq!(threading.personal_styles(), vec![""; count]);
        for m in threading.messages() {
            assert_eq!(m.message_id(), "");
            assert_eq!(m.display(), None);
        }
        assert_eq!(
            d.revision_authors(),
            if count == 2 {
                vec!["Unknown", "Ryker Zhu"]
            } else {
                vec!["Unknown"]
            }
        );
        assert_eq!(
            d.fib().get_table_pointer(94).unwrap().1,
            if count == 2 { 60 } else { 48 }
        );
        assert!(p.to_bytes().unwrap() == bytes);
    }
}
fn expected(file: &str) -> (Vec<(u32, u32, u16)>, Vec<(u32, u32, State)>) {
    match file {
        "word-text-structure.doc" => (
            vec![
                (0, 217, 3),
                (217, 218, 4),
                (218, 316, 3),
                (316, 398, 4),
                (398, 514, 7),
            ],
            vec![
                (0, 217, State::Dirty),
                (217, 218, State::Edit),
                (218, 316, State::Dirty),
                (316, 398, State::Edit),
                (398, 514, State::Clean),
            ],
        ),
        "word-bidi-eastasia.doc" => (
            vec![
                (0, 58, 4),
                (58, 110, 8),
                (110, 111, 4),
                (111, 151, 8),
                (151, 152, 4),
                (152, 188, 8),
                (188, 189, 4),
                (189, 223, 8),
                (223, 245, 4),
                (245, 248, 7),
            ],
            vec![
                (0, 58, State::Edit),
                (58, 110, State::Dirty),
                (110, 111, State::Edit),
                (111, 151, State::Dirty),
                (151, 152, State::Edit),
                (152, 188, State::Dirty),
                (188, 189, State::Edit),
                (189, 223, State::Dirty),
                (223, 245, State::Edit),
                (245, 248, State::Clean),
            ],
        ),
        "word-layout-extensions.dot" => (
            vec![
                (0, 215, 3),
                (215, 234, 4),
                (234, 299, 3),
                (299, 378, 4),
                (378, 549, 7),
            ],
            vec![
                (0, 215, State::Dirty),
                (215, 234, State::Edit),
                (234, 299, State::Dirty),
                (299, 306, State::Edit),
                (306, 331, State::Clean),
                (331, 378, State::Edit),
                (378, 549, State::Clean),
            ],
        ),
        _ => panic!("unexpected fixture"),
    }
}
fn wire(cps: &[(u32, u32, u16)]) -> Vec<u8> {
    let mut out = cps
        .iter()
        .map(|r| r.0)
        .chain([cps.last().unwrap().1])
        .flat_map(u32::to_le_bytes)
        .collect::<Vec<_>>();
    out.extend(cps.iter().flat_map(|r| r.2.to_le_bytes()));
    out
}
#[test]
fn native_word_language_and_smart_tag_states_keep_distinct_saved_ranges() {
    for file in [
        "word-text-structure.doc",
        "word-bidi-eastasia.doc",
        "word-layout-extensions.dot",
    ] {
        let bytes = native(file);
        let p = open(file, &bytes);
        let d = p.document();
        let (lad, recognizer) = expected(file);
        let (_, data) = table(&p);
        let language = d.proofing_tables().unwrap().language_detection().unwrap();
        assert_eq!(
            language
                .ranges()
                .map(|r| (
                    r.start_cp(),
                    r.end_cp(),
                    r.status()
                        .to_raw(ProofingFeature::LanguageAutoDetect)
                        .unwrap()
                ))
                .collect::<Vec<_>>(),
            lad
        );
        assert_eq!(language.to_bytes().unwrap(), wire(&lad));
        assert_eq!(data[range(&p, 98)], wire(&lad));
        let smart = d.smart_tags().unwrap();
        assert!(smart.tags.is_empty() && smart.store.is_none());
        assert_eq!(
            smart
                .recognizer_ranges
                .iter()
                .map(|r| (r.start, r.end, r.state))
                .collect::<Vec<_>>(),
            recognizer
        );
        let raw = recognizer
            .iter()
            .map(|&(start, end, state)| {
                (
                    start,
                    end,
                    match state {
                        State::Dirty => 3,
                        State::Edit => 4,
                        State::Clean => 7,
                        _ => panic!("unexpected native state"),
                    },
                )
            })
            .collect::<Vec<_>>();
        assert_eq!(data[range(&p, 132)], wire(&raw));
        for index in [114, 115, 117, 118] {
            assert_eq!(d.fib().get_table_pointer(index).unwrap().1, 0);
        }
        assert!(p.to_bytes().unwrap() == bytes);
    }
}
fn publish_table(source: &Package, data: Vec<u8>, selected: std::ops::Range<usize>) -> Package {
    let bytes = source.to_bytes().unwrap();
    let (path, before) = table(source);
    assert_eq!(data.len(), before.len());
    assert_eq!(&data[..selected.start], &before[..selected.start]);
    assert_eq!(&data[selected.end..], &before[selected.end..]);
    let object = object::Package::open(bytes.clone()).unwrap();
    let mut e = object.edit().unwrap();
    e.put_stream(&path, data.clone()).unwrap();
    let c = e.commit().unwrap();
    let output = c.patch().after();
    let a = streams(&bytes);
    let b = streams(output);
    assert_eq!(a.keys().collect::<Vec<_>>(), b.keys().collect::<Vec<_>>());
    for (n, v) in a {
        assert!(
            b[&n] == if n == path { data.clone() } else { v },
            "changed {}",
            n.join("/")
        );
    }
    assert_eq!(
        cfb_metadata::directory(&bytes),
        cfb_metadata::directory(output)
    );
    assert!(c.patch().apply(&object).unwrap().patch().after() == output);
    let after = object::Package::open(output.to_vec()).unwrap();
    assert!(c.patch().inverse().apply(&after).unwrap().patch().after() == bytes);
    let p = Package::from_bytes(output).unwrap();
    let f = tempfile::NamedTempFile::with_suffix(".doc").unwrap();
    p.save(f.path()).unwrap();
    let q = Package::open(f.path()).unwrap();
    assert!(q.to_bytes().unwrap() == output);
    assert_eq!(
        q.document().text().unwrap(),
        source.document().text().unwrap()
    );
    q
}
#[test]
fn native_word_recognizer_transaction_changes_only_its_owned_range_and_inverts() {
    for file in [
        "word-text-structure.doc",
        "word-bidi-eastasia.doc",
        "word-layout-extensions.dot",
    ] {
        let bytes = native(file);
        let p = open(file, &bytes);
        let d = p.document();
        let (_, data) = table(&p);
        let snapshot = smart_tags::Snapshot::parse(d.fib(), &data).unwrap();
        assert!(snapshot.is_present());
        assert!(snapshot.tags().is_empty());
        assert_eq!(snapshot.metadata(), d.smart_tags().unwrap());
        assert!(snapshot.edit().commit().unwrap().patch().is_noop());
        let mut e = snapshot.edit();
        e.set_recognizer_state(0, State::MaybeDirty).unwrap();
        let staged = e.metadata().clone();
        let mut invalid = staged.recognizer_ranges[0];
        invalid.end += 1;
        assert!(matches!(
            e.set_recognizer_range(0, invalid),
            Err(smart_tags::Error::Invalid(litchi_doc::Error::Corrupted(m))) if m=="Plcffactoid ranges must form one contiguous CP sequence"
        ));
        assert_eq!(e.metadata(), &staged);
        let c = e.commit().unwrap();
        assert_eq!(c.package().recognizer_ranges()[0].state, State::MaybeDirty);
        assert_eq!(c.patch().apply(&snapshot).unwrap(), *c.package());
        assert_eq!(c.patch().inverse().apply(c.package()).unwrap(), snapshot);
        assert!(matches!(
            c.patch().apply(c.package()),
            Err(smart_tags::Error::PatchConflict)
        ));
        assert_eq!(c.package().fib_bytes(), snapshot.fib_bytes());
        let q = publish_table(&p, c.package().table_bytes().to_vec(), range(&p, 132));
        assert_eq!(q.document().smart_tags(), Some(c.package().metadata()));
        assert_eq!(
            q.document().proofing_tables().unwrap(),
            d.proofing_tables().unwrap()
        );
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
#[test]
fn native_word_language_codec_edits_leave_other_proofing_and_recognizer_state_exact() {
    for file in ["word-text-structure.doc", "word-bidi-eastasia.doc"] {
        let bytes = native(file);
        let p = open(file, &bytes);
        let d = p.document();
        let source = d.proofing_tables().unwrap().language_detection().unwrap();
        let mut all = d.proofing_tables().unwrap().clone();
        let status = litchi_doc::parts::proofing::ProofingStatus::try_new(
            ProofingFeature::LanguageAutoDetect,
            litchi_doc::parts::proofing::ProofingState::Clean,
            false,
            false,
            false,
        )
        .unwrap();
        let mut entries = source.entries().to_vec();
        entries[0] = litchi_doc::parts::proofing::ProofingEntry::new(entries[0].start_cp(), status);
        let changed = ProofingStateTable::try_new(
            ProofingFeature::LanguageAutoDetect,
            entries,
            source.terminal_cp(),
        )
        .unwrap();
        all.set(changed.clone());
        assert_eq!(all.spelling(), d.proofing_tables().unwrap().spelling());
        assert_eq!(all.grammar(), d.proofing_tables().unwrap().grammar());
        let (_, mut data) = table(&p);
        let selected = range(&p, 98);
        data[selected.clone()].copy_from_slice(&changed.to_bytes().unwrap());
        let q = publish_table(&p, data, selected);
        assert_eq!(q.document().proofing_tables().unwrap(), &all);
        assert_eq!(q.document().smart_tags(), d.smart_tags());
        let mut detached = all.clone();
        assert_eq!(
            detached.remove(ProofingFeature::LanguageAutoDetect),
            Some(changed)
        );
        assert!(detached.language_detection().is_none());
        assert_eq!(detached.spelling(), all.spelling());
        assert!(p.to_bytes().unwrap() == bytes);
    }
}

#[test]
fn native_word_recognizer_catalog_accepts_absent_and_state_only_graphs() {
    let with_state = [
        ("word-text-structure.doc", 5),
        ("word-tables-layout.doc", 37),
        ("word-review-controls.doc", 6),
        ("word-math-navigation.doc", 15),
        ("word-visual-objects.doc", 3),
        ("word-mailmerge-fields.doc", 17),
        ("word-vector-diagram.doc", 3),
        ("word-bidi-eastasia.doc", 10),
        ("word-layout-extensions.dot", 7),
    ];
    let mut seen = 0;
    for (file, _) in CACHES {
        let bytes = native(file);
        let p = open(file, &bytes);
        let d = p.document();
        let expected = with_state.iter().find(|r| r.0 == file).map(|r| r.1);
        assert_eq!(d.smart_tags().map(|s| s.recognizer_ranges.len()), expected);
        if let Some(s) = d.smart_tags() {
            assert!(s.tags.is_empty() && s.store.is_none());
            seen += 1;
        }
        if file != "word-encrypted.doc" {
            let (_, data) = table(&p);
            let snapshot = smart_tags::Snapshot::parse(d.fib(), &data).unwrap();
            assert_eq!(snapshot.is_present(), expected.is_some());
            assert_eq!(snapshot.recognizer_ranges().len(), expected.unwrap_or(0));
            let c = snapshot.edit().commit().unwrap();
            assert!(c.patch().is_noop());
            assert_eq!(c.package().fib_bytes(), d.fib().raw_data());
            assert_eq!(c.package().table_bytes(), data);
        }
        assert!(p.to_bytes().unwrap() == bytes);
    }
    assert_eq!(seen, 9);
}

#[test]
fn native_word_recognizer_point_ranges_preserve_the_source_terminal() {
    let file = "word-text-structure.doc";
    let bytes = native(file);
    let p = open(file, &bytes);
    let d = p.document();
    let (_, data) = table(&p);
    let snapshot = smart_tags::Snapshot::parse(d.fib(), &data).unwrap();
    let terminal = snapshot.recognizer_ranges().last().unwrap().end;
    assert_eq!(terminal, 514);
    assert!(terminal > snapshot.topology().document_end());
    let mut e = snapshot.edit();
    e.update(|m| {
        m.recognizer_ranges[0].end = 0;
        m.recognizer_ranges[1].start = 0;
        Ok(())
    })
    .unwrap();
    let valid = e.metadata().clone();
    let last = valid.recognizer_ranges.len() - 1;
    let mut invalid = valid.recognizer_ranges[last];
    invalid.end += 1;
    assert!(
        matches!(e.set_recognizer_range(last,invalid),Err(smart_tags::Error::Invalid(litchi_doc::Error::Corrupted(m))) if m=="Plcffactoid terminal sentinel must remain unchanged")
    );
    assert_eq!(e.metadata(), &valid);
    let mut invalid = valid.recognizer_ranges[0];
    invalid.start = snapshot.topology().document_end() + 1;
    assert!(
        matches!(e.set_recognizer_range(0,invalid),Err(smart_tags::Error::Invalid(litchi_doc::Error::Corrupted(m))) if m=="Plcffactoid range is outside the document parts")
    );
    assert_eq!(e.metadata(), &valid);
    let c = e.commit().unwrap();
    assert_eq!(c.package().recognizer_ranges()[0].start, 0);
    assert_eq!(c.package().recognizer_ranges()[0].end, 0);
    assert_eq!(
        c.package().recognizer_ranges().last().unwrap().end,
        terminal
    );
    let q = publish_table(&p, c.package().table_bytes().to_vec(), range(&p, 132));
    assert_eq!(q.document().smart_tags(), Some(c.package().metadata()));
    assert_eq!(c.patch().inverse().apply(c.package()).unwrap(), snapshot);
}
