//! Native AutoSummary state; priority additions use explicit record codecs and CFB publication.
use litchi_doc::{
    AutoSummaryRange, AutoSummaryState, AutoSummaryView, DocumentAutoSummary, Dop97, Package,
};
use litchi_ole_common::object;
use std::{collections::BTreeMap, io::Cursor, path::Path};
#[path = "../../../tools/test_support/cfb_directory_metadata.rs"]
mod cfb_metadata;
const FILES: [&str; 3] = [
    "word-text-structure.doc",
    "word-recorded-macro.dot",
    "word-layout-extensions.dot",
];
fn native(file: &str) -> Vec<u8> {
    std::fs::read(
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office/word")
            .join(file),
    )
    .unwrap()
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
fn state(p: &Package) -> AutoSummaryState {
    Dop97::parse(
        &p.document()
            .document_properties()
            .unwrap()
            .unwrap()
            .to_bytes()
            .unwrap(),
    )
    .unwrap()
    .auto_summary
}
fn absent_state() -> AutoSummaryState {
    AutoSummaryState {
        valid: false,
        view_active: false,
        view: AutoSummaryView::Highlight,
        update_properties: false,
        desired_level: 0,
        highest_level: 0,
        current_level: 0,
    }
}
fn active_state() -> AutoSummaryState {
    AutoSummaryState {
        valid: true,
        view_active: true,
        view: AutoSummaryView::HideNonSummary,
        update_properties: false,
        desired_level: 75,
        highest_level: 3,
        current_level: 2,
    }
}
#[test]
fn native_doc_autosummary_state_is_inert_and_has_no_priority_table() {
    for file in FILES {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let d = p.document();
        assert_eq!(state(&p), absent_state());
        assert!(d.auto_summary().is_none());
        assert_eq!(d.fib().get_table_pointer(89).unwrap().1, 0);
        let raw = d
            .document_properties()
            .unwrap()
            .unwrap()
            .to_bytes()
            .unwrap();
        let mut output = raw.clone();
        Dop97::parse(&raw).unwrap().write_into(&mut output).unwrap();
        assert!(raw == output);
        assert!(p.to_bytes().unwrap() == bytes);
    }
}
fn publish(
    source: &Package,
    summary: Option<&DocumentAutoSummary>,
    value: AutoSummaryState,
) -> Package {
    let bytes = source.to_bytes().unwrap();
    let before = streams(&bytes);
    let d = source.document();
    let fib = d.fib();
    let word_path = vec!["WordDocument".to_owned()];
    let table_path = vec![format!("{}Table", u8::from(fib.which_table_stream()))];
    let mut word = before[&word_path].clone();
    let mut table = before[&table_path].clone();
    let (offset, len) = fib.get_table_pointer(31).unwrap();
    let start = usize::try_from(offset).unwrap();
    let end = usize::try_from(offset + len).unwrap();
    let mut dop = Dop97::parse(&table[start..end]).unwrap();
    dop.auto_summary = value;
    dop.write_into(&mut table[start..end]).unwrap();
    // The Asumyi record is 12 bytes at offset 414 in the versioned DOP.
    let mut old_prefix = before[&table_path].clone();
    let mut new_prefix = table.clone();
    old_prefix[start + 414..start + 426].fill(0);
    new_prefix[start + 414..start + 426].fill(0);
    assert!(old_prefix == new_prefix);
    let pointer = 154 + 89 * 8;
    if let Some(summary) = summary {
        let raw = summary.to_bytes().unwrap();
        word[pointer..pointer + 4]
            .copy_from_slice(&u32::try_from(table.len()).unwrap().to_le_bytes());
        word[pointer + 4..pointer + 8]
            .copy_from_slice(&u32::try_from(raw.len()).unwrap().to_le_bytes());
        table.extend(raw);
    } else {
        word[pointer..pointer + 8].fill(0);
    }
    let mut expected_word = before[&word_path].clone();
    let mut masked = word.clone();
    expected_word[pointer..pointer + 8].fill(0);
    masked[pointer..pointer + 8].fill(0);
    assert!(expected_word == masked);
    let cfb = object::Package::open(bytes.clone()).unwrap();
    let mut e = cfb.edit().unwrap();
    e.put_stream(&word_path, word.clone()).unwrap();
    e.put_stream(&table_path, table.clone()).unwrap();
    let c = e.commit().unwrap();
    let output = c.patch().after();
    assert_eq!(
        cfb_metadata::directory(&bytes),
        cfb_metadata::directory(output)
    );
    let after = streams(output);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (path, data) in before {
        let wanted = if path == word_path {
            &word
        } else if path == table_path {
            &table
        } else {
            &data
        };
        assert!(
            after.get(&path) == Some(wanted),
            "changed {}",
            path.join("/")
        );
    }
    assert!(c.patch().apply(&cfb).unwrap().patch().after() == output);
    let reopened_cfb = object::Package::open(output.to_vec()).unwrap();
    assert!(
        c.patch()
            .inverse()
            .apply(&reopened_cfb)
            .unwrap()
            .patch()
            .after()
            == bytes
    );
    let p = Package::from_bytes(output).unwrap();
    let f = tempfile::NamedTempFile::with_suffix(".doc").unwrap();
    p.save(f.path()).unwrap();
    let q = Package::open(f.path()).unwrap();
    assert!(q.to_bytes().unwrap() == output);
    assert_eq!(q.document().auto_summary(), summary);
    assert_eq!(state(&q), value);
    assert_eq!(q.document().text().unwrap(), d.text().unwrap());
    assert!(source.to_bytes().unwrap() == bytes);
    q
}
#[test]
fn native_doc_autosummary_record_edits_preserve_the_native_dop_and_other_streams() {
    for file in FILES {
        let bytes = native(file);
        let p = Package::from_bytes(&bytes).unwrap();
        let mut summary = DocumentAutoSummary::try_new(vec![
            AutoSummaryRange::new(0, 8, 1).unwrap(),
            AutoSummaryRange::new(8, 16, 3).unwrap(),
        ])
        .unwrap();
        let added = publish(&p, Some(&summary), active_state());
        let stored = added.document().auto_summary().unwrap().clone();
        assert_eq!(
            stored
                .ranges()
                .iter()
                .map(|r| (r.start(), r.end(), r.level()))
                .collect::<Vec<_>>(),
            [(0, 8, 1), (8, 16, 3)]
        );
        summary
            .replace(1, AutoSummaryRange::new(8, 16, 2).unwrap())
            .unwrap();
        let valid = summary.clone();
        let error = summary
            .push(AutoSummaryRange::new(17, 20, 4).unwrap())
            .unwrap_err();
        assert!(
            matches!(error,litchi_doc::Error::Corrupted(m) if m=="PlcfAsumy range must start at the current terminal CP")
        );
        assert_eq!(summary, valid);
        let updated = publish(&added, Some(&summary), active_state());
        assert_eq!(summary.remove(1).unwrap().level(), 2);
        let removed = publish(&updated, Some(&summary), active_state());
        let cleared = publish(&removed, None, absent_state());
        assert!(cleared.document().auto_summary().is_none());
        assert_eq!(cleared.document().fib().get_table_pointer(89), Some((0, 0)));
        assert!(p.to_bytes().unwrap() == bytes);
        assert!(native(file) == bytes);
    }
}
