#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use litchi_cfb::writer::{OleWriter, StorageMoveLimits};
use litchi_cfb::{DirectoryEntry, OleFile};
use std::io::Cursor;

fn serialize(writer: &mut OleWriter) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn child<'a>(
    file: &'a OleFile<Cursor<Vec<u8>>>,
    parent: &[&str],
    name: &str,
) -> &'a DirectoryEntry {
    file.list_directory_entries(parent)
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == name)
        .expect("named child")
}

fn bounded_writer() -> OleWriter {
    let mut writer = OleWriter::new();
    writer.create_storage(&["Source"]).unwrap();
    writer.create_storage(&["Source", "Nested"]).unwrap();
    writer
        .create_stream(&["Source", "Nested", "Payload"], b"payload")
        .unwrap();
    writer
}

#[test]
fn moves_complete_subtree_without_reallocating_or_reordering_stream_payloads() {
    let first = vec![0x11; 5_003];
    let second = vec![0x22; 6_007];
    let unrelated = vec![0x33; 7_009];

    let mut writer = OleWriter::new();
    writer.create_storage(&["Source"]).unwrap();
    writer.create_storage(&["Source", "Nested"]).unwrap();
    writer.create_storage(&["Destination"]).unwrap();
    writer.create_storage(&["Unrelated"]).unwrap();
    writer.set_storage_clsid(&["Source"], [0x11; 16]).unwrap();
    writer
        .set_storage_clsid(&["Source", "Nested"], [0x22; 16])
        .unwrap();
    writer
        .set_storage_clsid(&["Unrelated"], [0x33; 16])
        .unwrap();
    writer
        .create_stream_owned(&["Source", "First"], first.clone())
        .unwrap();
    writer
        .create_stream_owned(&["Source", "Nested", "Second"], second.clone())
        .unwrap();
    writer
        .create_stream_owned(&["Unrelated", "Third"], unrelated.clone())
        .unwrap();

    let before = OleFile::open(Cursor::new(serialize(&mut writer))).unwrap();
    let before_first_sector = child(&before, &["Source"], "First").start_sector;
    let before_second_sector = child(&before, &["Source", "Nested"], "Second").start_sector;
    let before_third = child(&before, &["Unrelated"], "Third");
    let before_third_sector = before_third.start_sector;
    let before_unrelated_clsid = child(&before, &[], "Unrelated").clsid.clone();
    let before_source_clsid = child(&before, &[], "Source").clsid.clone();
    let before_nested_clsid = child(&before, &["Source"], "Nested").clsid.clone();

    writer
        .move_storage(&["sOURCE"], &["dESTINATION", "Renamed"])
        .unwrap();

    let mut after = OleFile::open(Cursor::new(serialize(&mut writer))).unwrap();
    assert!(!after.directory_exists(&["Source"]));
    assert!(after.directory_exists(&["Destination", "Renamed"]));
    assert!(after.directory_exists(&["Destination", "Renamed", "Nested"]));
    assert_eq!(
        after
            .open_stream(&["Destination", "Renamed", "First"])
            .unwrap(),
        first
    );
    assert_eq!(
        after
            .open_stream(&["Destination", "Renamed", "Nested", "Second"])
            .unwrap(),
        second
    );
    assert_eq!(
        after.open_stream(&["Unrelated", "Third"]).unwrap(),
        unrelated
    );

    assert_eq!(
        child(&after, &["Destination", "Renamed"], "First").start_sector,
        before_first_sector
    );
    assert_eq!(
        child(&after, &["Destination", "Renamed", "Nested"], "Second").start_sector,
        before_second_sector
    );
    assert_eq!(
        child(&after, &["Unrelated"], "Third").start_sector,
        before_third_sector
    );
    assert_eq!(
        child(&after, &["Destination"], "Renamed").clsid,
        before_source_clsid
    );
    assert_eq!(
        child(&after, &["Destination", "Renamed"], "Nested").clsid,
        before_nested_clsid
    );
    assert_eq!(
        child(&after, &[], "Unrelated").clsid,
        before_unrelated_clsid
    );
}

#[test]
fn accepts_case_only_rename_but_refuses_other_case_insensitive_collisions() {
    let mut rename = OleWriter::new();
    rename.create_storage(&["Reports"]).unwrap();
    rename
        .create_stream(&["Reports", "Data"], b"retained")
        .unwrap();
    rename.move_storage(&["Reports"], &["rEPORTS"]).unwrap();
    let mut renamed = OleFile::open(Cursor::new(serialize(&mut rename))).unwrap();
    assert_eq!(child(&renamed, &[], "rEPORTS").name, "rEPORTS");
    assert_eq!(
        renamed.open_stream(&["rEPORTS", "Data"]).unwrap(),
        b"retained"
    );

    let mut collision = OleWriter::new();
    collision.create_storage(&["Source"]).unwrap();
    collision
        .create_stream(&["Source", "Payload"], b"source")
        .unwrap();
    collision.create_storage(&["Target"]).unwrap();
    collision
        .create_stream(&["Target", "tAKEN"], b"unrelated")
        .unwrap();
    let before = serialize(&mut collision);
    let error = collision
        .move_storage(&["sOURCE"], &["tARGET", "Taken"])
        .unwrap_err();
    assert!(error.to_string().contains("collides"));
    assert_eq!(serialize(&mut collision), before);
}

#[test]
fn root_self_descendant_and_missing_parent_refusals_are_atomic() {
    let mut writer = bounded_writer();
    let before = serialize(&mut writer);

    let root_error = writer.move_storage(&[], &["Moved"]).unwrap_err();
    assert!(root_error.to_string().contains("root"));
    assert_eq!(serialize(&mut writer), before);

    let root_destination_error = writer.move_storage(&["Source"], &[]).unwrap_err();
    assert!(root_destination_error.to_string().contains("root"));
    assert_eq!(serialize(&mut writer), before);

    let self_error = writer
        .move_storage(&["Source"], &["source", "Child"])
        .unwrap_err();
    assert!(self_error.to_string().contains("own subtree"));
    assert_eq!(serialize(&mut writer), before);

    let parent_error = writer
        .move_storage(&["Source"], &["Missing", "Moved"])
        .unwrap_err();
    assert!(parent_error.to_string().contains("parent"));
    assert_eq!(serialize(&mut writer), before);
}

#[test]
fn refuses_oversized_untrusted_names_before_admitting_a_move() {
    let mut writer = bounded_writer();
    let before = serialize(&mut writer);
    let oversized = "x".repeat(1_000_000);

    let error = writer
        .move_storage(&[oversized.as_str()], &["Moved"])
        .unwrap_err();
    assert!(error.to_string().contains("31 UTF-16"));
    assert_eq!(serialize(&mut writer), before);
}

#[test]
fn refuses_an_oversized_existing_path_before_canonical_allocation() {
    let mut writer = bounded_writer();
    let oversized = "y".repeat(1_000_000);
    writer
        .create_storage(&[oversized.as_str()])
        .expect("the legacy creation API defers directory-name validation");

    let error = writer.move_storage(&["Source"], &["Moved"]).unwrap_err();
    assert!(error.to_string().contains("31 UTF-16"));

    // The failed move did not consume or rename the source. Once the invalid
    // unrelated entry is removed, the same source still moves successfully.
    writer.delete_storage(&[oversized.as_str()]).unwrap();
    writer.move_storage(&["Source"], &["Moved"]).unwrap();
    let mut reopened = OleFile::open(Cursor::new(serialize(&mut writer))).unwrap();
    assert_eq!(
        reopened
            .open_stream(&["Moved", "Nested", "Payload"])
            .unwrap(),
        b"payload"
    );
}

#[test]
fn explicit_limits_accept_exact_bounds_and_refuse_one_less_atomically() {
    assert!(StorageMoveLimits::new(0, 1, 1).is_err());
    let exact = StorageMoveLimits::new(3, 3, 3).unwrap();
    assert_eq!(exact.max_entries_scanned(), 3);
    assert_eq!(exact.max_descendants(), 3);
    assert_eq!(exact.max_path_components(), 3);

    let mut accepted = bounded_writer();
    accepted
        .move_storage_with_limits(&["Source"], &["Moved"], exact)
        .unwrap();
    let mut reopened = OleFile::open(Cursor::new(serialize(&mut accepted))).unwrap();
    assert_eq!(
        reopened
            .open_stream(&["Moved", "Nested", "Payload"])
            .unwrap(),
        b"payload"
    );

    let cases = [
        StorageMoveLimits::new(2, 3, 3).unwrap(),
        StorageMoveLimits::new(3, 2, 3).unwrap(),
        StorageMoveLimits::new(3, 3, 2).unwrap(),
    ];
    for limits in cases {
        let mut refused = bounded_writer();
        let before = serialize(&mut refused);
        let error = refused
            .move_storage_with_limits(&["Source"], &["Moved"], limits)
            .unwrap_err();
        assert!(error.to_string().contains("limit"));
        assert_eq!(serialize(&mut refused), before);
    }
}
