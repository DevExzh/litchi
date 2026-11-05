#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration tests use concise assertions for checked fixtures"
)]

use litchi_cfb::{OleError, OleFile, OleWriter};
use litchi_ole_common::object::{Edit, Limits, MAX_STREAM_REMOVALS, OpenOptions, Package, Targets};
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

fn open(bytes: Vec<u8>) -> Edit {
    Package::open_with_options(
        bytes,
        OpenOptions::new()
            .targets(Targets::default())
            .limits(Limits::default()),
    )
    .unwrap()
    .edit()
    .unwrap()
}

trait RenderBytes {
    fn render_bytes(&self) -> Result<Vec<u8>, litchi_ole_common::Error>;
}

impl RenderBytes for Edit {
    fn render_bytes(&self) -> Result<Vec<u8>, litchi_ole_common::Error> {
        let package = self.clone().commit()?.into_package();
        let mut bytes = Vec::new();
        package.write_to(&mut bytes)?;
        Ok(bytes)
    }
}

impl RenderBytes for Package {
    fn render_bytes(&self) -> Result<Vec<u8>, litchi_ole_common::Error> {
        let mut bytes = Vec::new();
        self.write_to(&mut bytes)?;
        Ok(bytes)
    }
}

#[test]
fn removes_first_middle_last_and_only_stream_without_removing_storages() {
    let source = package();
    let source_ole = OleFile::open(
        Cursor::new(source.clone()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
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

    let output = editor.render_bytes().unwrap();
    let ole = OleFile::open(Cursor::new(output), litchi_cfb::OleFileLimits::default()).unwrap();
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
    let source_package = Package::open_with_options(
        source.clone(),
        OpenOptions::new()
            .targets(Targets::default())
            .limits(Limits::default()),
    )
    .unwrap();
    let after = commit.patch().apply(&source_package).unwrap();
    assert_eq!(
        after.package().render_bytes().unwrap(),
        commit.patch().after()
    );
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(after.package())
            .unwrap()
            .package()
            .render_bytes()
            .unwrap(),
        source
    );
    assert!(commit.patch().apply(commit.package()).is_err());
}

#[test]
fn a_committed_package_is_the_base_for_a_follow_up_noop_edit() {
    let source = package();
    let mut edit = open(source);
    edit.remove_stream(&path(&["First"]))
        .expect("first edit should remove the stream");
    let first = edit.commit().expect("first edit should commit");

    let follow_up = first
        .package()
        .edit()
        .expect("published package should remain editable")
        .commit()
        .expect("follow-up no-op should commit");

    assert!(follow_up.patch().is_noop());
    assert_eq!(follow_up.patch().before(), first.patch().after());
    assert_eq!(follow_up.patch().after(), first.patch().after());
    assert_eq!(
        follow_up.package().render_bytes().unwrap(),
        first.patch().after()
    );
}

#[test]
fn chained_publications_rebase_indexes_and_replay_on_reopened_content() {
    let source = package();
    let unchanged_path = path(&["Parent", "Regular"]);
    let unchanged = Package::open(source.clone())
        .unwrap()
        .stream(&unchanged_path)
        .unwrap()
        .to_vec();

    let mut edit = open(source.clone());
    edit.put_stream(&path(&["First"]), b"changed first".to_vec())
        .unwrap();
    edit.put_stream(&path(&["Parent", "Mini"]), b"changed mini".to_vec())
        .unwrap();
    edit.remove_stream(&path(&["Last"])).unwrap();
    edit.add_stream(path(&["Last"]), b"recreated last".to_vec())
        .unwrap();
    let first = edit.commit().unwrap();

    // The untouched stream must still be read from the package whose CFB
    // indexes were recaptured after publication, rather than from the old
    // source/index pair.
    assert_eq!(
        first.package().stream(&unchanged_path).unwrap(),
        unchanged.as_slice()
    );

    let reopened_source = Package::open(source).unwrap();
    let forward = first.patch().apply(&reopened_source).unwrap();
    assert_eq!(
        forward.package().render_bytes().unwrap(),
        first.patch().after()
    );

    let reopened_after = Package::open(first.patch().after().to_vec()).unwrap();
    let inverse = first.patch().inverse().apply(&reopened_after).unwrap();
    assert_eq!(
        inverse.package().render_bytes().unwrap(),
        first.patch().before()
    );

    let follow_up = first.package().edit().unwrap().commit().unwrap();
    assert!(follow_up.patch().is_noop());
    assert_eq!(
        follow_up.package().render_bytes().unwrap(),
        first.patch().after()
    );
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
    assert!(!editor.clone().commit().unwrap().diagnostics().changed());
    let commit = editor.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().after(), source);

    for invalid in [vec![], path(&[""]), path(&["bad/name"]), path(&["Parent"])] {
        let mut failed = open(source.clone());
        assert!(failed.remove_stream(&invalid).is_err(), "{invalid:?}");
        assert!(!failed.clone().commit().unwrap().diagnostics().changed());
        assert_eq!(failed.render_bytes().unwrap(), source);
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
    assert!(!duplicate.clone().commit().unwrap().diagnostics().changed());
    assert_eq!(duplicate.render_bytes().unwrap(), source);

    let mut late = open(source.clone());
    assert!(
        late.remove_streams([first.as_slice(), invalid.as_slice()])
            .is_err()
    );
    assert!(!late.clone().commit().unwrap().diagnostics().changed());
    assert_eq!(late.render_bytes().unwrap(), source);

    let storage = path(&["Parent"]);
    let mut late_storage = open(source.clone());
    assert!(
        late_storage
            .remove_streams([first.as_slice(), storage.as_slice()])
            .is_err()
    );
    assert!(
        !late_storage
            .clone()
            .commit()
            .unwrap()
            .diagnostics()
            .changed()
    );
    assert_eq!(late_storage.render_bytes().unwrap(), source);

    let overdepth = vec!["x".to_string(); Limits::default().max_storage_depth + 2];
    let exact_depth = vec!["x".to_string(); Limits::default().max_storage_depth + 1];
    let mut exact_depth_editor = open(source.clone());
    assert!(
        exact_depth_editor
            .remove_stream(&exact_depth)
            .unwrap()
            .is_none()
    );
    assert!(
        !exact_depth_editor
            .clone()
            .commit()
            .unwrap()
            .diagnostics()
            .changed()
    );

    let mut single_overdepth = open(source.clone());
    assert!(single_overdepth.remove_stream(&overdepth).is_err());
    assert!(
        !single_overdepth
            .clone()
            .commit()
            .unwrap()
            .diagnostics()
            .changed()
    );
    assert_eq!(single_overdepth.render_bytes().unwrap(), source);

    let mut late_overdepth = open(source.clone());
    assert!(
        late_overdepth
            .remove_streams([first.as_slice(), overdepth.as_slice()])
            .is_err()
    );
    assert!(
        !late_overdepth
            .clone()
            .commit()
            .unwrap()
            .diagnostics()
            .changed()
    );
    assert_eq!(late_overdepth.render_bytes().unwrap(), source);

    let selectors = (0..MAX_STREAM_REMOVALS)
        .map(|index| path(&[&format!("Missing{index}")]))
        .collect::<Vec<_>>();
    let mut exact_limit = open(source.clone());
    let results = exact_limit
        .remove_streams(selectors.iter().map(Vec::as_slice))
        .unwrap();
    assert_eq!(results.len(), MAX_STREAM_REMOVALS);
    assert!(results.iter().all(Option::is_none));
    assert!(
        !exact_limit
            .clone()
            .commit()
            .unwrap()
            .diagnostics()
            .changed()
    );

    let mut selectors = selectors;
    selectors.push(path(&["OneTooMany"]));
    let mut over_limit = open(source.clone());
    assert!(
        over_limit
            .remove_streams(selectors.iter().map(Vec::as_slice))
            .is_err()
    );
    assert!(!over_limit.clone().commit().unwrap().diagnostics().changed());
    assert_eq!(over_limit.render_bytes().unwrap(), source);
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

    let output = editor.render_bytes().unwrap();
    let mut ole = OleFile::open(Cursor::new(output), litchi_cfb::OleFileLimits::default()).unwrap();
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
    assert!(
        Package::open_with_options(
            output.into_inner(),
            OpenOptions::new()
                .targets(Targets::default())
                .limits(Limits::default())
        )
        .is_err()
    );
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
    let mut editor = Package::open_with_options(
        source.clone(),
        OpenOptions::new().targets(Targets::default()).limits(exact),
    )
    .unwrap()
    .edit()
    .unwrap();
    assert!(
        editor
            .remove_stream(&path(&["Parent", "Regular"]))
            .unwrap()
            .is_some()
    );
    assert!(editor.render_bytes().is_ok());

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
        assert!(
            Package::open_with_options(
                source.clone(),
                OpenOptions::new()
                    .targets(Targets::default())
                    .limits(limits)
            )
            .is_err()
        );
    }
}

#[test]
fn supplementary_plane_case_pairs_are_distinct_stream_names() {
    // [MS-CFB] 2.6.4 compares UTF-16 code points and uppercases neither
    // surrogate of a supplementary scalar (see `litchi-cfb`'s
    // `directory_name.rs`), so "𐐨" (U+10428) does not name the stored "𐐀"
    // (U+10400). This test used to expect the opposite, from the scalar-level
    // fold both crates shared before `litchi-cfb` adopted the code-point rule;
    // the object owner now compares names with `litchi-cfb`'s key.
    let mut writer = OleWriter::new();
    writer.create_stream(&["𐐀"], b"deseret").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let source = output.into_inner();

    let mut editor = open(source);
    assert_eq!(editor.remove_stream(&path(&["𐐨"])).unwrap(), None);
    assert_eq!(
        editor.remove_stream(&path(&["𐐀"])).unwrap().as_deref(),
        Some(&b"deseret"[..])
    );
    let output = editor.render_bytes().unwrap();
    assert!(
        OleFile::open(Cursor::new(output), litchi_cfb::OleFileLimits::default())
            .unwrap()
            .list_streams()
            .is_empty()
    );
}
