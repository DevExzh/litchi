#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration tests use concise assertions for checked fixtures"
)]

use litchi_cfb::{OleError, OleFile, OleWriter};
use litchi_ole_common::object::{Editor, Limits, MAX_STREAM_REMOVALS, Targets};
use std::io::Cursor;

fn path(parts: &[&str]) -> Vec<String> {
    parts.iter().map(|part| (*part).to_string()).collect()
}

fn package() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_storage(&["Parent"]).unwrap();
    writer
        .set_storage_clsid(
            &["Parent"],
            [
                0x12, 0x34, 0x56, 0x78, 0x9a, 0xbc, 0xde, 0xf0, 0x80, 0x01, 0x02, 0x03, 0x04, 0x05,
                0x06, 0x07,
            ],
        )
        .unwrap();
    writer.create_storage(&["Parent", "Empty"]).unwrap();
    writer.create_stream(&["First"], b"first").unwrap();
    writer
        .create_stream(&["Parent", "Mini"], &vec![0x5a; 4_095])
        .unwrap();
    writer
        .create_stream(&["Parent", "Regular"], &vec![0xa5; 4_096])
        .unwrap();
    writer.create_stream(&["Last"], b"last").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn open(bytes: Vec<u8>) -> Editor {
    Editor::open(bytes, Targets::default(), Limits::default()).unwrap()
}

#[test]
fn removes_first_middle_last_and_only_stream_without_removing_storages() {
    let source = package();
    let source_ole = OleFile::open(Cursor::new(source.clone())).unwrap();
    let parent_clsid = source_ole
        .list_directory_entries(&[])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "Parent")
        .unwrap()
        .clsid
        .clone();
    let mut editor = open(source);
    assert_eq!(
        editor.remove_stream(&path(&["first"])).unwrap().as_deref(),
        Some(&b"first"[..])
    );
    assert_eq!(
        editor
            .remove_stream(&path(&["parent", "mini"]))
            .unwrap()
            .as_deref()
            .map(<[u8]>::len),
        Some(4_095)
    );
    assert_eq!(
        editor
            .remove_stream(&path(&["Parent", "Regular"]))
            .unwrap()
            .as_deref()
            .map(<[u8]>::len),
        Some(4_096)
    );
    assert_eq!(
        editor.remove_stream(&path(&["Last"])).unwrap().as_deref(),
        Some(&b"last"[..])
    );

    let output = editor.finish().unwrap();
    let ole = OleFile::open(Cursor::new(output)).unwrap();
    assert!(ole.list_streams().is_empty());
    assert_eq!(ole.list_directory_entries(&[]).unwrap().len(), 1);
    assert_eq!(ole.list_directory_entries(&["Parent"]).unwrap().len(), 1);
    assert_eq!(
        ole.list_directory_entries(&[])
            .unwrap()
            .into_iter()
            .find(|entry| entry.name == "Parent")
            .unwrap()
            .clsid,
        parent_clsid
    );
}

#[test]
fn batch_reports_absence_preserves_unrelated_bytes_and_is_reversible() {
    let source = package();
    let first = path(&["First"]);
    let missing = path(&["Missing"]);
    let regular = path(&["Parent", "Regular"]);
    let mut editor = open(source.clone());
    let shared_first = editor.stream_shared(&first).unwrap();
    let removed = editor
        .remove_streams([first.as_slice(), missing.as_slice(), regular.as_slice()])
        .unwrap();
    assert_eq!(removed[0].as_deref(), Some(&b"first"[..]));
    assert!(std::sync::Arc::ptr_eq(
        &shared_first,
        removed[0].as_ref().unwrap()
    ));
    assert!(removed[1].is_none());
    assert_eq!(removed[2].as_deref().map(<[u8]>::len), Some(4_096));
    assert_eq!(editor.stream(&path(&["Last"])), Some(&b"last"[..]));
    assert_eq!(
        editor.stream(&path(&["Parent", "Mini"])).map(<[u8]>::len),
        Some(4_095)
    );

    let commit = editor.commit().unwrap();
    assert!(!commit.patch().is_noop());
    let after = commit.patch().apply(&source).unwrap();
    assert_eq!(after, commit.patch().after());
    assert_eq!(commit.patch().inverse().apply(&after).unwrap(), source);
    assert!(commit.patch().apply(b"foreign source").is_err());
}

#[test]
fn missing_empty_and_invalid_deletions_are_exact_atomic_noops() {
    let source = package();
    let missing = path(&["missing"]);
    let mut editor = open(source.clone());
    let shared_last = editor.stream_shared(&path(&["Last"])).unwrap();
    assert!(editor.remove_stream(&missing).unwrap().is_none());
    assert!(
        editor
            .remove_streams(std::iter::empty())
            .unwrap()
            .is_empty()
    );
    assert_eq!(editor.remove_streams([missing.as_slice()]).unwrap(), [None]);
    assert!(std::sync::Arc::ptr_eq(
        &shared_last,
        &editor.stream_shared(&path(&["Last"])).unwrap()
    ));
    assert!(!editor.is_changed());
    let commit = editor.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().after(), source);

    for invalid in [vec![], path(&[""]), path(&["bad/name"]), path(&["Parent"])] {
        let mut failed = open(source.clone());
        assert!(failed.remove_stream(&invalid).is_err(), "{invalid:?}");
        assert!(!failed.is_changed());
        assert_eq!(failed.finish().unwrap(), source);
    }
}

#[test]
fn duplicates_limit_and_late_invalid_paths_leave_the_editor_unchanged() {
    let source = package();
    let first = path(&["First"]);
    let equivalent = path(&["first"]);
    let invalid = path(&["bad:name"]);

    let mut duplicate = open(source.clone());
    assert!(
        duplicate
            .remove_streams([first.as_slice(), equivalent.as_slice()])
            .is_err()
    );
    assert!(!duplicate.is_changed());
    assert_eq!(duplicate.finish().unwrap(), source);

    let mut late = open(source.clone());
    assert!(
        late.remove_streams([first.as_slice(), invalid.as_slice()])
            .is_err()
    );
    assert!(!late.is_changed());
    assert_eq!(late.finish().unwrap(), source);

    let storage = path(&["Parent"]);
    let mut late_storage = open(source.clone());
    assert!(
        late_storage
            .remove_streams([first.as_slice(), storage.as_slice()])
            .is_err()
    );
    assert!(!late_storage.is_changed());
    assert_eq!(late_storage.finish().unwrap(), source);

    let overdepth = vec!["x".to_string(); Limits::default().max_storage_depth + 2];
    let exact_depth = vec!["x".to_string(); Limits::default().max_storage_depth + 1];
    let mut exact_depth_editor = open(source.clone());
    assert!(
        exact_depth_editor
            .remove_stream(&exact_depth)
            .unwrap()
            .is_none()
    );
    assert!(!exact_depth_editor.is_changed());

    let mut single_overdepth = open(source.clone());
    assert!(single_overdepth.remove_stream(&overdepth).is_err());
    assert!(!single_overdepth.is_changed());
    assert_eq!(single_overdepth.finish().unwrap(), source);

    let mut late_overdepth = open(source.clone());
    assert!(
        late_overdepth
            .remove_streams([first.as_slice(), overdepth.as_slice()])
            .is_err()
    );
    assert!(!late_overdepth.is_changed());
    assert_eq!(late_overdepth.finish().unwrap(), source);

    let selectors = (0..MAX_STREAM_REMOVALS)
        .map(|index| path(&[&format!("Missing{index}")]))
        .collect::<Vec<_>>();
    let mut exact_limit = open(source.clone());
    let results = exact_limit
        .remove_streams(selectors.iter().map(Vec::as_slice))
        .unwrap();
    assert_eq!(results.len(), MAX_STREAM_REMOVALS);
    assert!(results.iter().all(Option::is_none));
    assert!(!exact_limit.is_changed());

    let mut selectors = selectors;
    selectors.push(path(&["OneTooMany"]));
    let mut over_limit = open(source.clone());
    assert!(
        over_limit
            .remove_streams(selectors.iter().map(Vec::as_slice))
            .is_err()
    );
    assert!(!over_limit.is_changed());
    assert_eq!(over_limit.finish().unwrap(), source);
}

#[test]
fn deletion_composes_with_add_and_replace_in_one_editor() {
    let mut editor = open(package());
    editor
        .add_stream(path(&["Parent", "Added"]), b"added".to_vec())
        .unwrap();
    editor
        .put_stream(&path(&["Last"]), b"replaced".to_vec())
        .unwrap();
    assert!(editor.remove_stream(&path(&["First"])).unwrap().is_some());

    let output = editor.finish().unwrap();
    let mut ole = OleFile::open(Cursor::new(output)).unwrap();
    assert_eq!(ole.open_stream(&["Parent", "Added"]).unwrap(), b"added");
    assert_eq!(ole.open_stream(&["Last"]).unwrap(), b"replaced");
    assert!(matches!(
        ole.open_stream(&["First"]),
        Err(OleError::StreamNotFound)
    ));
}

#[test]
fn protected_packages_are_refused_before_stream_deletion_can_begin() {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["DigitalSignature"], b"opaque signature")
        .unwrap();
    writer.create_stream(&["Payload"], b"payload").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    assert!(Editor::open(output.into_inner(), Targets::default(), Limits::default()).is_err());
}

#[test]
fn configured_capture_boundaries_remain_in_force_for_deletion() {
    let source = package();
    let exact = Limits {
        max_objects: 1,
        max_storage_depth: 2,
        max_streams_per_object: 1,
        max_streams: 4,
        max_stream_size: 4_096,
        max_object_size: 512,
        max_total_size: 8_200,
    };
    let mut editor = Editor::open(source.clone(), Targets::default(), exact).unwrap();
    assert!(
        editor
            .remove_stream(&path(&["Parent", "Regular"]))
            .unwrap()
            .is_some()
    );
    assert!(editor.finish().is_ok());

    for limits in [
        Limits {
            max_streams: 3,
            ..exact
        },
        Limits {
            max_stream_size: 4_095,
            ..exact
        },
        Limits {
            max_total_size: 8_199,
            ..exact
        },
    ] {
        assert!(Editor::open(source.clone(), Targets::default(), limits).is_err());
    }
}

#[test]
fn supplementary_plane_simple_uppercase_resolves_the_stored_stream_name() {
    let mut writer = OleWriter::new();
    writer.create_stream(&["𐐀"], b"deseret").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let source = output.into_inner();

    let mut editor = open(source);
    assert_eq!(
        editor.remove_stream(&path(&["𐐨"])).unwrap().as_deref(),
        Some(&b"deseret"[..])
    );
    let output = editor.finish().unwrap();
    assert!(
        OleFile::open(Cursor::new(output))
            .unwrap()
            .list_streams()
            .is_empty()
    );
}
