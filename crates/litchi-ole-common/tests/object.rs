#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::cast_possible_truncation,
    clippy::drop_non_drop,
    clippy::default_trait_access,
    clippy::bool_assert_comparison,
    reason = "integration tests use concise assertions and checked fixture-sized literals"
)]

use litchi_cfb::{
    OleFile, OleWriter, OverlayLimits, SameLengthStreamOverlay, SectorLayoutPolicy, SharedOleFile,
};
use litchi_core::OwnedSource;
use litchi_ole_common::object::{Editor, EntryKind, Limits, Snapshot, Target, Targets, discover};
use litchi_ole_common::property_set::Guid;
use std::io::Cursor;
use std::sync::Arc;

fn write_cfb(build: impl FnOnce(&mut OleWriter)) -> Vec<u8> {
    let mut writer = OleWriter::new();
    build(&mut writer);
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("test CFB should write");
    output.into_inner()
}

fn first_free_sector(bytes: &[u8]) -> usize {
    let sector_size = 1usize << u16::from_le_bytes(bytes[0x1E..0x20].try_into().unwrap());
    let fat_count = u32::from_le_bytes(bytes[0x2C..0x30].try_into().unwrap()) as usize;
    let mut fat = Vec::new();
    for index in 0..fat_count {
        let sector = u32::from_le_bytes(
            bytes[0x4C + index * 4..0x50 + index * 4]
                .try_into()
                .unwrap(),
        ) as usize;
        let start = (sector + 1) * sector_size;
        for word in bytes[start..start + sector_size].chunks_exact(4) {
            fat.push(u32::from_le_bytes(word.try_into().unwrap()));
        }
    }
    let physical_sectors = bytes.len() / sector_size - 1;
    fat.iter()
        .take(physical_sectors)
        .position(|entry| *entry == 0xFFFF_FFFF)
        .expect("source should contain a free sector")
}

fn target(key: &str, path: &[&str]) -> Target {
    Target::new(key, path.iter().copied()).expect("test target should validate")
}

fn targets(key: &str, path: &[&str]) -> Targets {
    Targets::one(target(key, path))
}

fn ansi(value: &str, output: &mut Vec<u8>) {
    output.extend_from_slice(&((value.len() + 1) as u32).to_le_bytes());
    output.extend_from_slice(value.as_bytes());
    output.push(0);
}

fn comp_obj(user_type: &str, prog_id: &str) -> Vec<u8> {
    let mut output = vec![0; 28];
    output[12..28].copy_from_slice(&[
        0x12, 0x34, 0x56, 0x78, 0x9A, 0xBC, 0xDE, 0xF0, 0x80, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06,
        0x07,
    ]);
    ansi(user_type, &mut output);
    ansi("Embedded Object", &mut output);
    ansi(prog_id, &mut output);
    output
}

fn native(command: &str, payload: &[u8]) -> Vec<u8> {
    let mut body = Vec::new();
    body.extend_from_slice(&2u16.to_le_bytes());
    body.extend_from_slice(b"report.txt\0");
    body.extend_from_slice(b"report.txt\0");
    body.extend_from_slice(&[0; 4]);
    body.extend_from_slice(command.as_bytes());
    body.push(0);
    body.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    body.extend_from_slice(payload);
    let mut output = (body.len() as u32).to_le_bytes().to_vec();
    output.extend_from_slice(&body);
    output
}

fn doc_with_object(obj_info: &[u8]) -> Vec<u8> {
    let metadata = comp_obj("Package", "Package");
    let native = native("do-not-run", b"opaque native bytes");
    write_cfb(|writer| {
        writer
            .create_stream(&["WordDocument"], b"unknown-records")
            .expect("test stream should write");
        writer
            .create_storage(&["ObjectPool", "_42"])
            .expect("test storage should write");
        writer
            .create_stream(&["ObjectPool", "_42", "\u{3}ObjInfo"], obj_info)
            .expect("test metadata should write");
        writer
            .create_stream(&["ObjectPool", "_42", "\u{1}CompObj"], &metadata)
            .expect("test metadata should write");
        writer
            .create_stream(&["ObjectPool", "_42", "\u{1}Ole10Native"], &native)
            .expect("test native stream should write");
        writer
            .create_stream(&["ObjectPool", "_42", "\u{3}PRINT"], b"metafile")
            .expect("test preview stream should write");
    })
}

#[test]
fn discovers_only_host_selected_storage_and_keeps_metadata_opaque() {
    let bytes = doc_with_object(&[0x40, 0x00, 0x02, 0x00]);
    let mut ole = OleFile::open(Cursor::new(bytes)).expect("test CFB should open");
    let selected = targets("host-object", &["ObjectPool", "_42"]);
    let objects = discover(&mut ole, &selected, Limits::default()).expect("discovery should pass");
    let object = objects
        .get("host-object")
        .expect("target should be present");
    assert_eq!(object.key(), "host-object");
    assert_eq!(
        object.path(),
        ["ObjectPool".to_string(), "_42".to_string()].as_slice()
    );
    assert_eq!(
        object.stream(&["\u{3}ObjInfo"]),
        Some(&[0x40, 0x00, 0x02, 0x00][..])
    );
    assert_eq!(object.storage().directory().kind(), EntryKind::Storage);
    assert!(object.storage().directory().sid().raw() > 0);
    assert_eq!(object.streams().len(), 4);
    assert!(object.stream(&["\u{1}Ole10Native"]).is_some());
    let preview = object
        .streams()
        .iter()
        .find(|stream| stream.name() == Some("\u{3}PRINT"))
        .expect("preview stream metadata should be captured");
    let preview_directory = preview
        .directory()
        .expect("preview should retain directory metadata");
    assert_eq!(preview_directory.kind(), EntryKind::Stream);
    assert_eq!(
        preview_directory.stream_size(),
        preview.bytes().len() as u64
    );
    assert!(preview_directory.links().child().is_none());
    assert!(object.compound().starts_with(&[0xD0, 0xCF, 0x11, 0xE0]));
    assert_eq!(
        objects.at(0).map(litchi_ole_common::object::Object::key),
        Some("host-object")
    );
}

#[test]
fn target_catalog_is_explicit_and_rejects_ambiguous_paths() {
    let first = target("first", &["Pool", "A"]);
    let second = target("second", &["Pool", "B"]);
    let selected = Targets::new([first.clone(), second]).expect("targets should validate");
    assert_eq!(selected.get("first"), Some(&first));
    assert!(Targets::new([first.clone(), target("other", &["Pool", "A"])]).is_err());
    assert!(Targets::new([first, target("first", &["Pool", "C"])]).is_err());
    assert!(Targets::new([target("parent", &["Pool"]), target("child", &["Pool", "A"]),]).is_err());
}

#[test]
fn target_paths_follow_cfb_name_limits_and_simple_uppercase_identity() {
    assert!(Target::new("empty", [""]).is_err());
    assert!(Target::new("forbidden", ["Pool/Child"]).is_err());
    assert!(Target::new("nul", ["Pool\0Child"]).is_err());
    assert!(Target::new("too-long", ["😀".repeat(16)]).is_err());
    assert!(Target::new("control-is-allowed", ["\u{3}ObjInfo"]).is_ok());

    let upper = target("upper", &["Pool", "Child"]);
    let lower = target("lower", &["pool", "child"]);
    assert!(Targets::new([upper, lower]).is_err());
}

#[test]
fn discovery_resolves_case_variant_target_paths_to_stored_cfb_names() {
    let bytes = doc_with_object(&[0, 0, 0, 0]);
    let mut ole = OleFile::open(Cursor::new(bytes)).expect("test CFB should open");
    let selected = targets("object", &["objectpool", "_42"]);
    let objects = discover(&mut ole, &selected, Limits::default()).expect("target should resolve");
    assert_eq!(
        objects
            .get("object")
            .expect("object should be present")
            .path(),
        ["ObjectPool".to_string(), "_42".to_string()].as_slice()
    );

    let editor = Editor::open(doc_with_object(&[0, 0, 0, 0]), selected, Limits::default())
        .expect("editor target should resolve");
    assert_eq!(
        editor
            .targets()
            .get("object")
            .expect("resolved target should be present")
            .path(),
        ["ObjectPool".to_string(), "_42".to_string()].as_slice()
    );
}

#[test]
fn malformed_format_metadata_is_retained_without_common_classification() {
    let malformed = doc_with_object(&[0x00, 0x04, 0x00, 0x00]);
    let mut ole = OleFile::open(Cursor::new(malformed)).expect("test CFB should open");
    let selected = targets("object", &["ObjectPool", "_42"]);
    let objects = discover(&mut ole, &selected, Limits::default()).expect("opaque data is valid");
    assert_eq!(
        objects
            .get("object")
            .expect("object should be present")
            .stream(&["\u{3}ObjInfo"]),
        Some(&[0x00, 0x04, 0x00, 0x00][..])
    );

    let valid = doc_with_object(&[0, 0, 0, 0]);
    let mut ole = OleFile::open(Cursor::new(valid)).expect("test CFB should open");
    let limits = Limits {
        max_stream_size: 4,
        ..Limits::default()
    };
    assert!(discover(&mut ole, &selected, limits).is_err());
}

#[test]
fn missing_target_is_a_checked_discovery_error() {
    let bytes = doc_with_object(&[0, 0, 0, 0]);
    let mut ole = OleFile::open(Cursor::new(bytes)).expect("test CFB should open");
    let selected = targets("missing", &["ObjectPool", "_404"]);
    assert!(discover(&mut ole, &selected, Limits::default()).is_err());
}

#[test]
fn targeted_replace_preserves_unrelated_streams_and_opaque_reference() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let replacement = write_cfb(|writer| {
        writer
            .create_stream(&["\u{1}CompObj"], &comp_obj("Worksheet", "Excel.Sheet.8"))
            .expect("replacement metadata should write");
        writer
            .create_stream(&["CONTENTS"], b"new inert workbook bytes")
            .expect("replacement payload should write");
    });
    let selected = targets("object", &["ObjectPool", "_42"]);
    let mut editor =
        Editor::open(original, selected.clone(), Limits::default()).expect("editor should open");
    editor
        .replace("object", replacement)
        .expect("replacement should commit");
    assert!(editor.is_changed());
    let output = editor.finish().expect("editor should finish");
    let mut ole = OleFile::open(Cursor::new(output)).expect("output CFB should open");
    assert_eq!(
        ole.open_stream(&["WordDocument"])
            .expect("unrelated stream should remain"),
        b"unknown-records"
    );
    assert_eq!(
        ole.open_stream(&["ObjectPool", "_42", "CONTENTS"])
            .expect("replacement stream should be present"),
        b"new inert workbook bytes"
    );
    let objects = discover(&mut ole, &selected, Limits::default()).expect("reopen should pass");
    assert_eq!(
        objects
            .get("object")
            .expect("object should remain selected")
            .stream(&["\u{1}CompObj"])
            .expect("opaque metadata should remain"),
        comp_obj("Worksheet", "Excel.Sheet.8").as_slice()
    );
}

#[test]
fn no_op_editor_round_trip_is_byte_identical() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let editor = Editor::open(
        original.clone(),
        targets("object", &["ObjectPool", "_42"]),
        Limits::default(),
    )
    .expect("editor should open");
    assert!(!editor.is_changed());
    assert_eq!(editor.finish().expect("editor should finish"), original);
}

#[test]
fn commit_exposes_snapshot_and_reversible_patch() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let mut editor = Editor::open(
        original.clone(),
        targets("object", &["ObjectPool", "_42"]),
        Limits::default(),
    )
    .expect("editor should open");
    editor
        .put_stream(&["WordDocument".into()], b"changed".to_vec())
        .expect("stream edit should commit");

    let committed = editor.commit().expect("commit should validate");
    assert_eq!(committed.patch().before(), original.as_slice());
    assert_eq!(
        committed
            .snapshot()
            .finish()
            .expect("snapshot should finish"),
        committed.patch().after()
    );
    assert_eq!(
        committed
            .patch()
            .inverse()
            .apply(committed.patch().after())
            .expect("inverse should apply"),
        original
    );
}

#[test]
fn editor_edit_preserves_directory_metadata_for_unchanged_entries() {
    let original = write_cfb(|writer| {
        writer.set_root_state_bits(0x1020_3040);
        writer.set_root_creation_time_from_source(0x0102_0304_0506_0708);
        writer.set_root_modified_time(0x1112_1314_1516_1718);
        writer.create_storage(&["ObjectPool"]).unwrap();
        writer
            .set_storage_metadata(
                &["ObjectPool"],
                0xA1A2_A3A4,
                0x2122_2324_2526_2728,
                0x3132_3334_3536_3738,
            )
            .unwrap();
        writer.create_storage(&["ObjectPool", "_42"]).unwrap();
        writer
            .set_storage_metadata(
                &["ObjectPool", "_42"],
                0xB1B2_B3B4,
                0x4142_4344_4546_4748,
                0x5152_5354_5556_5758,
            )
            .unwrap();
        writer
            .create_stream(&["WordDocument"], b"unchanged host bytes")
            .unwrap();
        writer
            .create_stream(&["ObjectPool", "_42", "\u{3}PRINT"], b"preview")
            .unwrap();
        writer
            .set_stream_metadata(
                &["ObjectPool", "_42", "\u{3}PRINT"],
                0xC1C2_C3C4,
                0x6162_6364_6566_6768,
                0x7172_7374_7576_7778,
            )
            .unwrap();
    });
    let selected = targets("object", &["ObjectPool", "_42"]);
    let mut editor = Editor::open(original, selected, Limits::default()).expect("editor opens");
    editor
        .put_stream(&["WordDocument".into()], b"edited host bytes".to_vec())
        .expect("unrelated stream edit should commit");
    let output = editor.finish().expect("edited package should finish");
    let file = OleFile::open(Cursor::new(output)).expect("edited CFB should open");

    let root = file.root_entry().expect("root entry");
    assert_eq!(root.state_bits, 0x1020_3040);
    assert_eq!(root.creation_time, 0x0102_0304_0506_0708);
    assert_eq!(root.modified_time, 0x1112_1314_1516_1718);
    let object_pool = file
        .list_directory_entries(&[])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "ObjectPool")
        .unwrap();
    assert_eq!(object_pool.state_bits, 0xA1A2_A3A4);
    assert_eq!(object_pool.creation_time, 0x2122_2324_2526_2728);
    assert_eq!(object_pool.modified_time, 0x3132_3334_3536_3738);
    let object = file
        .list_directory_entries(&["ObjectPool"])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "_42")
        .unwrap();
    assert_eq!(object.state_bits, 0xB1B2_B3B4);
    assert_eq!(object.creation_time, 0x4142_4344_4546_4748);
    assert_eq!(object.modified_time, 0x5152_5354_5556_5758);
    let preview = file
        .list_directory_entries(&["ObjectPool", "_42"])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "\u{3}PRINT")
        .unwrap();
    assert_eq!(preview.state_bits, 0xC1C2_C3C4);
    assert_eq!(preview.creation_time, 0x6162_6364_6566_6768);
    assert_eq!(preview.modified_time, 0x7172_7374_7576_7778);
}

#[test]
fn selected_storage_metadata_stays_on_object_while_promoted_root_uses_defaults() {
    let source = write_cfb(|writer| {
        writer.create_storage(&["Pool", "Object"]).unwrap();
        writer
            .set_storage_metadata(
                &["Pool", "Object"],
                0xA1A2_A3A4,
                0x0102_0304_0506_0708,
                0x1112_1314_1516_1718,
            )
            .unwrap();
        writer
            .create_stream(&["Pool", "Object", "Payload"], b"payload")
            .unwrap();
    });
    let mut ole = OleFile::open(Cursor::new(source.clone())).expect("source CFB should open");
    let objects = discover(
        &mut ole,
        &targets("object", &["Pool", "Object"]),
        Limits::default(),
    )
    .expect("object discovery should pass");
    let object = objects.get("object").expect("object should be present");
    let source_storage = object.storage().directory();
    assert_eq!(source_storage.kind(), EntryKind::Storage);
    assert_eq!(source_storage.state_bits(), 0xA1A2_A3A4);
    assert_eq!(source_storage.creation_time(), 0x0102_0304_0506_0708);
    assert_eq!(source_storage.modified_time(), 0x1112_1314_1516_1718);

    let promoted = OleFile::open(Cursor::new(object.compound().to_vec()))
        .expect("promoted object CFB should open");
    let root = promoted.root_entry().expect("promoted root should exist");
    assert_eq!(root.entry_type, EntryKind::Root.raw());
    assert_eq!(root.state_bits, 0);
    assert_eq!(root.creation_time, 0);
    assert_eq!(root.modified_time, 0);

    let editor = Editor::open(
        source,
        targets("object", &["Pool", "Object"]),
        Limits::default(),
    )
    .expect("editor should open");
    let editor_object = editor.objects().get("object").expect("object should exist");
    assert_eq!(
        editor_object.storage().directory().creation_time(),
        0x0102_0304_0506_0708
    );
    let editor_promoted = OleFile::open(Cursor::new(editor_object.compound().to_vec()))
        .expect("editor-promoted object CFB should open");
    let editor_root = editor_promoted
        .root_entry()
        .expect("editor-promoted root should exist");
    assert_eq!(editor_root.state_bits, 0);
    assert_eq!(editor_root.creation_time, 0);
    assert_eq!(editor_root.modified_time, 0);
}

#[test]
fn replacement_preserves_target_storage_metadata_while_mapping_clsid() {
    const REPLACEMENT_CLSID: [u8; 16] = [
        0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00, 0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
        0x46,
    ];
    let original = write_cfb(|writer| {
        writer.create_storage(&["Pool", "Object"]).unwrap();
        writer
            .set_storage_metadata(
                &["Pool", "Object"],
                0xA1A2_A3A4,
                0x0102_0304_0506_0708,
                0x1112_1314_1516_1718,
            )
            .unwrap();
        writer
            .create_stream(&["Pool", "Object", "Payload"], b"old")
            .unwrap();
    });
    let replacement = write_cfb(|writer| {
        writer.set_root_clsid(REPLACEMENT_CLSID);
        writer.set_root_state_bits(0xB1B2_B3B4);
        writer.set_root_creation_time_from_source(0x2122_2324_2526_2728);
        writer.set_root_modified_time(0x3132_3334_3536_3738);
        writer.create_stream(&["Payload"], b"new").unwrap();
    });
    let mut editor = Editor::open(
        original,
        targets("object", &["Pool", "Object"]),
        Limits::default(),
    )
    .expect("editor should open");
    editor
        .replace("object", replacement)
        .expect("replacement should commit");
    let output = editor.finish().expect("edited package should finish");
    let mut file = OleFile::open(Cursor::new(output)).expect("edited CFB should open");
    let target = file
        .list_directory_entries(&["Pool"])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "Object")
        .unwrap();
    assert_eq!(target.state_bits, 0xA1A2_A3A4);
    assert_eq!(target.creation_time, 0x0102_0304_0506_0708);
    assert_eq!(target.modified_time, 0x1112_1314_1516_1718);
    assert_eq!(target.clsid, "00020906-0000-0000-C000-000000000046");
    let objects = discover(
        &mut file,
        &targets("object", &["Pool", "Object"]),
        Limits::default(),
    )
    .expect("replaced object should rediscover");
    assert_eq!(
        objects
            .get("object")
            .expect("replaced object should exist")
            .storage()
            .class_id(),
        Some(Guid::from_bytes(REPLACEMENT_CLSID))
    );
    assert_eq!(
        file.open_stream(&["Pool", "Object", "Payload"]).unwrap(),
        b"new"
    );
}

#[test]
fn added_storage_uses_zero_metadata_for_nonzero_source_root() {
    let mut editor = Editor::open(
        doc_with_object(&[0, 0, 0, 0]),
        targets("first", &["ObjectPool", "_42"]),
        Limits::default(),
    )
    .expect("editor should open");
    let replacement = write_cfb(|writer| {
        writer.set_root_state_bits(0xA1A2_A3A4);
        writer.set_root_creation_time_from_source(0x0102_0304_0506_0708);
        writer.set_root_modified_time(0x1112_1314_1516_1718);
        writer
            .create_stream(&["CONTENTS"], b"new object")
            .expect("nested payload should write");
    });
    editor
        .add_storage(target("second", &["ObjectPool", "_43"]), replacement)
        .expect("explicit storage should be added");

    let added = editor
        .objects()
        .get("second")
        .expect("storage should be present");
    let metadata = added.storage().directory();
    assert_eq!(metadata.kind(), EntryKind::Storage);
    assert_eq!(metadata.state_bits(), 0);
    assert_eq!(metadata.creation_time(), 0);
    assert_eq!(metadata.modified_time(), 0);
    assert_eq!(added.stream(&["CONTENTS"]), Some(&b"new object"[..]));

    let output = editor.finish().expect("edited package should finish");
    let file = OleFile::open(Cursor::new(output)).expect("edited CFB should open");
    let target = file
        .list_directory_entries(&["ObjectPool"])
        .unwrap()
        .into_iter()
        .find(|entry| entry.name == "_43")
        .expect("added storage should be serialized");
    assert_eq!(target.state_bits, 0);
    assert_eq!(target.creation_time, 0);
    assert_eq!(target.modified_time, 0);
}

#[test]
fn object_add_and_replace_preflight_limits_are_failure_atomic() {
    let replacement = write_cfb(|writer| {
        writer
            .create_storage(&["A"])
            .expect("first replacement storage should write");
        writer
            .create_storage(&["B"])
            .expect("second replacement storage should write");
    });
    let limits = Limits {
        max_objects: 2,
        max_storage_depth: 2,
        ..Limits::default()
    };

    let original = write_cfb(|writer| {
        writer
            .create_storage(&["ObjectPool", "_42"])
            .expect("selected storage should write");
        writer
            .create_storage(&["Other"])
            .expect("unrelated storage should write");
    });
    let mut editor = Editor::open(
        original.clone(),
        targets("object", &["ObjectPool", "_42"]),
        limits,
    )
    .expect("source should fit aggregate storage limits");
    let prepared = editor
        .prepare_replacement("object", replacement.clone())
        .expect("replacement should be admitted independently")
        .expect("replacement should change the selected object");
    assert!(editor.replace_prepared(prepared).is_err());
    assert!(!editor.is_changed());
    assert_eq!(
        editor.finish().expect("failed replacement stays exact"),
        original
    );

    let original = doc_with_object(&[0, 0, 0, 0]);
    let mut editor = Editor::open(
        original.clone(),
        targets("first", &["ObjectPool", "_42"]),
        limits,
    )
    .expect("source should fit aggregate storage limits");
    assert!(
        editor
            .add_storage(target("second", &["ObjectPool", "_43"]), replacement)
            .is_err()
    );
    assert!(!editor.is_changed());
    assert_eq!(editor.finish().expect("failed add stays exact"), original);
}

#[test]
fn failed_replacement_is_transactional() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let mut editor = Editor::open(
        original.clone(),
        targets("object", &["ObjectPool", "_42"]),
        Limits::default(),
    )
    .expect("editor should open");
    assert!(editor.replace("object", vec![1, 2, 3]).is_err());
    assert!(!editor.is_changed());
    assert_eq!(editor.finish().expect("editor should finish"), original);
}

#[test]
fn uniquely_owned_noop_editor_retains_the_consumed_vec_allocation() {
    let mut source = doc_with_object(&[0, 0, 0, 0]);
    source.reserve(4096);
    let expected = source.clone();
    let allocation = source.as_ptr();
    let capacity = source.capacity();
    let editor = Editor::open(
        source,
        targets("object", &["ObjectPool", "_42"]),
        Limits::default(),
    )
    .expect("editor should admit the owned source");
    let retained = editor.source_shared();
    assert_eq!(retained.as_ptr(), allocation);
    assert_eq!(retained.capacity(), capacity);
    drop(retained);
    let output = editor.finish().expect("unique no-op should finish");
    assert_eq!(output.as_ptr(), allocation);
    assert_eq!(output.capacity(), capacity);
    assert_eq!(output, expected);
}

#[test]
fn prepared_replacement_admission_is_bounded_and_source_bound() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let selected = targets("object", &["ObjectPool", "_42"]);
    let replacement = write_cfb(|writer| {
        writer
            .create_stream(&["CONTENTS"], b"prepared replacement")
            .expect("replacement stream should write");
    });

    let editor = Editor::open(original.clone(), selected.clone(), Limits::default())
        .expect("editor should open");
    let original_object_size = editor
        .objects()
        .get("object")
        .expect("object should exist")
        .compound()
        .len() as u64;
    assert!(
        editor
            .prepare_replacement(
                "object",
                editor.objects().get("object").unwrap().compound().to_vec()
            )
            .expect("exact replacement should be admitted")
            .is_none()
    );
    assert!(!editor.is_changed());
    assert_eq!(
        editor
            .clone()
            .finish()
            .expect("no-op should finish exactly"),
        original
    );

    let bounded_limits = Limits {
        max_object_size: original_object_size,
        ..Limits::default()
    };
    let bounded = Editor::open(original.clone(), selected.clone(), bounded_limits)
        .expect("bounded source should open");
    let oversized = write_cfb(|writer| {
        writer
            .create_stream(
                &["Large"],
                &vec![0u8; usize::try_from(original_object_size).unwrap()],
            )
            .expect("oversized replacement should write");
    });
    assert!(bounded.prepare_replacement("object", oversized).is_err());
    assert!(!bounded.is_changed());
    assert_eq!(
        bounded.finish().expect("failed admission stays exact"),
        original
    );

    let object_count_limits = Limits {
        max_streams_per_object: 4,
        max_storage_depth: 2,
        ..Limits::default()
    };
    let object_count_editor = Editor::open(original.clone(), selected.clone(), object_count_limits)
        .expect("source should fit per-object limits");
    let too_many_streams = write_cfb(|writer| {
        for name in ["A", "B", "C", "D", "E"] {
            writer
                .create_stream(&[name], b"replacement")
                .expect("replacement stream should write");
        }
    });
    assert!(
        object_count_editor
            .prepare_replacement("object", too_many_streams)
            .is_err()
    );

    let too_many_storages = write_cfb(|writer| {
        writer.create_storage(&["A"]).expect("storage should write");
        writer.create_storage(&["B"]).expect("storage should write");
        writer.create_storage(&["C"]).expect("storage should write");
    });
    assert!(
        object_count_editor
            .prepare_replacement("object", too_many_storages)
            .is_err()
    );
    assert!(!object_count_editor.is_changed());
    assert_eq!(
        object_count_editor
            .finish()
            .expect("per-object rejection stays exact"),
        original
    );

    let mut stale = Editor::open(original.clone(), selected.clone(), Limits::default())
        .expect("editor should open");
    let prepared = stale
        .prepare_replacement("object", replacement.clone())
        .expect("replacement should be admitted")
        .expect("replacement should change the object");
    stale
        .put_stream(
            &[
                "ObjectPool".to_owned(),
                "_42".to_owned(),
                "\u{3}PRINT".to_owned(),
            ],
            b"changed selected stream".to_vec(),
        )
        .expect("host edit should commit");
    let changed_source = stale.clone().finish().expect("host edit should finish");
    assert!(stale.replace_prepared(prepared).is_err());
    assert_eq!(
        stale.clone().finish().expect("stale admission stays exact"),
        changed_source
    );

    let mut different_limits = Limits::default();
    different_limits.max_total_size -= 1;
    let mut other = Editor::open(original.clone(), selected, different_limits)
        .expect("different limits should still admit the fixture");
    let prepared = editor
        .prepare_replacement("object", replacement.clone())
        .expect("replacement should be admitted")
        .expect("replacement should change the object");
    assert!(other.replace_prepared(prepared).is_err());
    assert!(!other.is_changed());
    assert_eq!(
        other.finish().expect("limit mismatch stays exact"),
        original
    );

    let mut same_limit_other = Editor::open(
        original.clone(),
        targets("object", &["ObjectPool", "_42"]),
        Limits::default(),
    )
    .expect("same-limit editor should open");
    let prepared = editor
        .prepare_replacement("object", replacement)
        .expect("replacement should be admitted")
        .expect("replacement should change the object");
    assert!(same_limit_other.replace_prepared(prepared).is_err());
    assert!(!same_limit_other.is_changed());
    assert_eq!(
        same_limit_other
            .finish()
            .expect("cross-editor rejection stays exact"),
        original
    );
}

#[test]
fn prepared_replacement_commit_supports_forward_and_inverse_patch() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let replacement = write_cfb(|writer| {
        writer
            .create_stream(&["CONTENTS"], b"forward replacement")
            .expect("replacement stream should write");
    });
    let mut editor = Editor::open(
        original.clone(),
        targets("object", &["ObjectPool", "_42"]),
        Limits::default(),
    )
    .expect("editor should open");
    let prepared = editor
        .prepare_replacement("object", replacement)
        .expect("replacement should be admitted")
        .expect("replacement should change the object");
    editor
        .replace_prepared(prepared)
        .expect("prepared replacement should publish");

    let committed = editor.commit().expect("replacement should commit");
    let forward = committed
        .patch()
        .apply(&original)
        .expect("forward replacement patch should apply");
    assert_eq!(forward, committed.patch().after());
    assert_eq!(
        committed
            .patch()
            .inverse()
            .apply(&forward)
            .expect("inverse replacement patch should apply"),
        original
    );
}

#[test]
fn shared_stream_replacement_reuses_validated_allocation() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let mut editor = Editor::open(
        original,
        targets("object", &["ObjectPool", "_42"]),
        Limits::default(),
    )
    .expect("editor should open");
    let path = vec![
        "ObjectPool".to_string(),
        "_42".to_string(),
        "\u{3}PRINT".to_string(),
    ];
    let replacement: Arc<[u8]> = Arc::from(&b"shared-word-stream"[..]);
    editor
        .put_stream_shared(&path, Arc::clone(&replacement))
        .expect("stream replacement should commit");
    let installed = editor
        .stream_shared(&path)
        .expect("stream should remain available");
    assert!(Arc::ptr_eq(&replacement, &installed));
    assert_eq!(editor.stream(&path), Some(&b"shared-word-stream"[..]));
    let metadata = editor
        .objects()
        .get("object")
        .and_then(|object| {
            object
                .streams()
                .iter()
                .find(|stream| stream.path() == &path[2..])
        })
        .and_then(|stream| stream.directory())
        .expect("edited stream should be reparsed before publication");
    assert_eq!(metadata.kind(), EntryKind::Stream);
    assert_eq!(metadata.stream_size(), replacement.len() as u64);
    assert!(metadata.uses_mini_stream());
}

#[test]
fn same_length_editor_edit_uses_source_backed_copy_through() {
    let base = write_cfb(|writer| {
        writer
            .create_stream(&["A"], &vec![0x11u8; 40_000])
            .expect("source stream should write");
        writer
            .create_stream(&["B"], &vec![0x22u8; 5_000])
            .expect("source stream should write");
    });
    let source = {
        let mut writer = OleWriter::new();
        assert!(writer.adopt_source_layout(&base).unwrap());
        writer
            .create_stream(&["A"], &vec![0x33u8; 6_000])
            .expect("source edit should write");
        writer
            .create_stream(&["B"], &vec![0x22u8; 5_000])
            .expect("source edit should write");
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    };
    let free_sector = first_free_sector(&source);
    let sector_size = 1usize << u16::from_le_bytes(source[0x1E..0x20].try_into().unwrap());
    let free_offset = (free_sector + 1) * sector_size;
    let mut mutated = source.clone();
    mutated[free_offset] = 0xA7;

    let mut editor = Editor::open(mutated.clone(), Targets::default(), Limits::default())
        .expect("source should open");
    editor
        .put_stream(&["A".into()], vec![0x44u8; 6_000])
        .expect("same-length edit should commit");
    let output = editor.finish().expect("same-length edit should finish");

    assert_eq!(output.len(), mutated.len());
    assert_eq!(
        output[free_offset], 0xA7,
        "copy-through retains untouched bytes"
    );
    let mut ole = OleFile::open(Cursor::new(output)).expect("copy-through output should reopen");
    assert_eq!(ole.open_stream(&["A"]).unwrap(), vec![0x44; 6_000]);
    assert_eq!(ole.open_stream(&["B"]).unwrap(), vec![0x22; 5_000]);
}

/// Publishes `overlays` over `source` through a generic positional CFB plan:
/// the exact route `render_copy_through` took before change 0748 sealed the
/// editor's original allocation.
fn generic_overlay_publication(source: &[u8], overlays: Vec<SameLengthStreamOverlay>) -> Vec<u8> {
    let shared = SharedOleFile::open(Arc::new(OwnedSource::new(source.to_vec())))
        .expect("generic source should open");
    let plan = shared
        .plan_same_length_stream_overlays(overlays, OverlayLimits::default())
        .expect("generic overlay should plan");
    let mut output = Vec::new();
    plan.write_to(&mut output)
        .expect("generic overlay should publish");
    output
}

#[test]
fn sealed_copy_through_publishes_the_generic_overlay_bytes() {
    let base = write_cfb(|writer| {
        writer
            .create_stream(&["A"], &vec![0x11u8; 40_000])
            .expect("source stream should write");
        writer
            .create_stream(&["B"], &vec![0x22u8; 5_000])
            .expect("source stream should write");
        writer
            .create_stream(&["Mini"], &vec![0x55u8; 700])
            .expect("source stream should write");
    });
    // Shrinking `A` under the adopted layout frees sectors; a byte in one of
    // them distinguishes copy-through from a re-render.
    let shrunk = {
        let mut writer = OleWriter::new();
        assert!(writer.adopt_source_layout(&base).unwrap());
        writer.create_stream(&["A"], &vec![0x33u8; 6_000]).unwrap();
        writer.create_stream(&["B"], &vec![0x22u8; 5_000]).unwrap();
        writer.create_stream(&["Mini"], &vec![0x55u8; 700]).unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    };
    let free_offset = {
        let sector_size = 1usize << u16::from_le_bytes(shrunk[0x1E..0x20].try_into().unwrap());
        (first_free_sector(&shrunk) + 1) * sector_size
    };
    let mut source = shrunk;
    source[free_offset] = 0xA7;

    let a = vec![0x44u8; 6_000];
    let mini = vec![0x66u8; 700];
    let expected_a = generic_overlay_publication(
        &source,
        vec![SameLengthStreamOverlay::new(
            vec!["A".into()],
            Arc::from(a.clone()),
        )],
    );
    let expected_both = generic_overlay_publication(
        &source,
        vec![
            SameLengthStreamOverlay::new(vec!["A".into()], Arc::from(a.clone())),
            SameLengthStreamOverlay::new(vec!["Mini".into()], Arc::from(mini.clone())),
        ],
    );
    assert_eq!(expected_a[free_offset], 0xA7);

    // One edit: the editor's commit-time render and its finish both take the
    // sealed copy-through and publish the generic plan's bytes.
    let mut editor =
        Editor::open(source.clone(), Targets::default(), Limits::default()).expect("open");
    editor
        .put_stream(&["A".into()], a.clone())
        .expect("same-length edit should commit");
    let snapshot = editor.snapshot();
    assert_eq!(snapshot.finish().expect("snapshot finish"), expected_a);
    let commit = editor.clone().commit().expect("commit");
    assert_eq!(commit.patch().before(), source.as_slice());
    assert_eq!(commit.patch().after(), expected_a.as_slice());
    assert_eq!(editor.finish().expect("finish"), expected_a);

    // Chained edits still overlay the one original allocation.
    let mut editor =
        Editor::open(source.clone(), Targets::default(), Limits::default()).expect("open");
    editor.put_stream(&["A".into()], a).expect("first edit");
    editor
        .put_stream(&["Mini".into()], mini)
        .expect("second edit");
    let chained = editor.finish().expect("chained finish");
    assert_eq!(chained, expected_both);
    let mut ole = OleFile::open(Cursor::new(chained)).expect("reopen");
    assert_eq!(ole.open_stream(&["B"]).unwrap(), vec![0x22; 5_000]);
}

#[test]
fn same_length_overlay_declines_noncanonical_v3_empty_size_word() {
    let source = write_cfb(|writer| {
        writer
            .create_stream(&["Empty"], &[])
            .expect("empty stream should write");
        writer
            .create_stream(&["A"], &[0x11u8; 128])
            .expect("edited stream should write");
    });
    let empty_sid = {
        let ole = OleFile::open(Cursor::new(source.clone())).expect("source should open");
        ole.list_directory_entries(&[])
            .expect("root entries should list")
            .into_iter()
            .find(|entry| entry.name == "Empty")
            .expect("empty stream should be present")
            .sid
    };
    let sector_size = 1usize << u16::from_le_bytes(source[0x1E..0x20].try_into().unwrap());
    let first_directory_sector =
        u32::from_le_bytes(source[0x30..0x34].try_into().unwrap()) as usize;
    let size_high = (first_directory_sector + 1) * sector_size + empty_sid as usize * 128 + 0x7C;
    let mut mutated = source;
    mutated[size_high..size_high + 4].copy_from_slice(&0xDEAD_BEEFu32.to_le_bytes());

    let mut editor = Editor::open(mutated, Targets::default(), Limits::default())
        .expect("mutated source should open");
    editor
        .put_stream(&["A".into()], vec![0x44u8; 128])
        .expect("same-length edit should commit");
    let output = editor.finish().expect("fallback layout should finish");
    let mut ole = OleFile::open(Cursor::new(output.clone())).expect("output should reopen");
    assert_eq!(ole.open_stream(&["Empty"]).unwrap(), Vec::<u8>::new());
    assert_eq!(ole.open_stream(&["A"]).unwrap(), vec![0x44; 128]);
    let output_size_high =
        (first_directory_sector + 1) * sector_size + empty_sid as usize * 128 + 0x7C;
    assert_eq!(
        &output[output_size_high..output_size_high + 4],
        &[0, 0, 0, 0]
    );
}

#[test]
fn batched_stream_replacement_is_atomic_and_reuses_allocations() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let selected = targets("object", &["ObjectPool", "_42"]);
    let word_path = vec!["WordDocument".to_string()];
    let preview_path = vec![
        "ObjectPool".to_string(),
        "_42".to_string(),
        "\u{3}PRINT".to_string(),
    ];
    let word: Arc<[u8]> = Arc::from(&b"batched word stream"[..]);
    let preview: Arc<[u8]> = Arc::from(&b"batched preview stream"[..]);
    let mut noop = Editor::open(original.clone(), selected.clone(), Limits::default())
        .expect("no-op editor should open");
    let original_word = noop
        .stream_shared(&word_path)
        .expect("original WordDocument should be available");
    noop.put_streams_shared([(word_path.as_slice(), Arc::clone(&original_word))])
        .expect("identical batch should be a no-op");
    assert!(!noop.is_changed());
    assert!(Arc::ptr_eq(
        &original_word,
        &noop
            .stream_shared(&word_path)
            .expect("no-op WordDocument should remain available")
    ));
    assert_eq!(noop.finish().expect("no-op batch stays exact"), original);

    let mut sequential = Editor::open(original.clone(), selected.clone(), Limits::default())
        .expect("sequential editor should open");
    sequential
        .put_stream_shared(&word_path, Arc::clone(&word))
        .expect("first sequential stream should commit");
    sequential
        .put_stream_shared(&preview_path, Arc::clone(&preview))
        .expect("second sequential stream should commit");
    let sequential = sequential
        .finish()
        .expect("sequential editor should finish");

    let mut editor = Editor::open(original.clone(), selected.clone(), Limits::default())
        .expect("editor should open");
    editor
        .put_streams_shared([
            (word_path.as_slice(), Arc::clone(&word)),
            (preview_path.as_slice(), Arc::clone(&preview)),
        ])
        .expect("stream batch should commit");
    assert!(Arc::ptr_eq(
        &word,
        &editor
            .stream_shared(&word_path)
            .expect("WordDocument should remain available")
    ));
    assert!(Arc::ptr_eq(
        &preview,
        &editor
            .stream_shared(&preview_path)
            .expect("preview should remain available")
    ));
    assert_eq!(
        editor.finish().expect("batched editor should finish"),
        sequential
    );

    let mut rendered_editor = Editor::open(original.clone(), selected.clone(), Limits::default())
        .expect("rendered editor should open");
    let rendered = rendered_editor
        .put_streams_shared_with_rendered([
            (word_path.as_slice(), Arc::clone(&word)),
            (preview_path.as_slice(), Arc::clone(&preview)),
        ])
        .expect("rendered batch should commit")
        .expect("effective batch should return rendered bytes");
    assert_eq!(
        rendered,
        rendered_editor
            .clone()
            .finish()
            .expect("rendered editor should finish")
    );
    assert!(Arc::ptr_eq(
        &word,
        &rendered_editor
            .stream_shared(&word_path)
            .expect("rendered WordDocument should remain available")
    ));
    assert!(Arc::ptr_eq(
        &preview,
        &rendered_editor
            .stream_shared(&preview_path)
            .expect("rendered preview should remain available")
    ));

    let current_word = rendered_editor
        .stream_shared(&word_path)
        .expect("current WordDocument should be available");
    assert!(
        rendered_editor
            .put_streams_shared_with_rendered([(word_path.as_slice(), Arc::clone(&current_word),)])
            .expect("all-no-op rendered batch should succeed")
            .is_none()
    );
    assert!(rendered_editor.is_changed());

    let missing_path = vec!["missing".to_string()];
    let mut failed = Editor::open(original.clone(), selected, Limits::default())
        .expect("second editor should open");
    let failed_word = failed
        .stream_shared(&word_path)
        .expect("failed editor WordDocument should be available");
    assert!(
        failed
            .put_streams_shared_with_rendered([
                (word_path.as_slice(), Arc::clone(&word)),
                (missing_path.as_slice(), Arc::clone(&preview)),
            ])
            .is_err()
    );
    assert!(!failed.is_changed());
    assert!(Arc::ptr_eq(
        &failed_word,
        &failed
            .stream_shared(&word_path)
            .expect("failed batch must preserve WordDocument")
    ));
    assert_eq!(failed.finish().expect("failed batch stays exact"), original);
}

#[test]
fn rendered_batch_preserves_repeated_path_last_value() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let word_path = vec!["WordDocument".to_string()];
    let first: Arc<[u8]> = Arc::from(&b"first repeated value"[..]);
    let second: Arc<[u8]> = Arc::from(&b"last repeated value"[..]);
    let mut editor =
        Editor::open(original, Targets::default(), Limits::default()).expect("editor should open");

    let rendered = editor
        .put_streams_shared_with_rendered([
            (word_path.as_slice(), Arc::clone(&first)),
            (word_path.as_slice(), Arc::clone(&second)),
        ])
        .expect("repeated-path batch should commit")
        .expect("effective repeated-path batch should render");

    assert_eq!(editor.stream(&word_path), Some(second.as_ref()));
    assert!(Arc::ptr_eq(
        &second,
        &editor
            .stream_shared(&word_path)
            .expect("last repeated value should be installed")
    ));
    assert_eq!(
        rendered,
        editor
            .clone()
            .finish()
            .expect("repeated-path editor should finish")
    );

    let mut restored = Editor::open(
        doc_with_object(&[0, 0, 0, 0]),
        Targets::default(),
        Limits::default(),
    )
    .expect("restored editor should open");
    let original_word = restored
        .stream_shared(&word_path)
        .expect("restored editor WordDocument should be available");
    let rendered = restored
        .put_streams_shared_with_rendered([
            (word_path.as_slice(), Arc::clone(&first)),
            (word_path.as_slice(), Arc::clone(&original_word)),
        ])
        .expect("restore batch should succeed");
    assert!(rendered.is_some());
    assert!(restored.is_changed());
    assert!(Arc::ptr_eq(
        &original_word,
        &restored
            .stream_shared(&word_path)
            .expect("restored value should be installed")
    ));
}

#[test]
fn rendered_batch_empty_and_equal_inputs_return_none_without_publication() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let word_path = vec!["WordDocument".to_string()];
    let mut editor = Editor::open(original.clone(), Targets::default(), Limits::default())
        .expect("editor should open");
    let original_word = editor
        .stream_shared(&word_path)
        .expect("original WordDocument should be available");

    assert!(
        editor
            .put_streams_shared_with_rendered(std::iter::empty::<(&[String], Arc<[u8]>)>())
            .expect("empty batch should succeed")
            .is_none()
    );
    assert!(!editor.is_changed());
    assert!(Arc::ptr_eq(
        &original_word,
        &editor
            .stream_shared(&word_path)
            .expect("empty batch should preserve WordDocument")
    ));

    assert!(
        editor
            .put_streams_shared_with_rendered([(word_path.as_slice(), Arc::clone(&original_word))])
            .expect("equal batch should succeed")
            .is_none()
    );
    assert!(!editor.is_changed());
    assert!(Arc::ptr_eq(
        &original_word,
        &editor
            .stream_shared(&word_path)
            .expect("equal batch should preserve WordDocument")
    ));
    assert_eq!(
        editor.finish().expect("no-op batch should stay exact"),
        original
    );
}

#[test]
fn rendered_batch_checks_max_streams_before_equal_repeated_items() {
    let original = write_cfb(|writer| {
        writer
            .create_stream(&["A"], b"source")
            .expect("source stream should write");
    });
    let limits = Limits {
        max_streams: 1,
        ..Limits::default()
    };
    let path = vec!["A".to_string()];
    let mut editor = Editor::open(original.clone(), Targets::default(), limits)
        .expect("bounded editor should open");
    let current = editor
        .stream_shared(&path)
        .expect("bounded source stream should be available");

    assert!(
        editor
            .put_streams_shared_with_rendered([
                (path.as_slice(), Arc::clone(&current)),
                (path.as_slice(), Arc::clone(&current)),
            ])
            .is_err()
    );
    assert!(!editor.is_changed());
    assert!(Arc::ptr_eq(
        &current,
        &editor
            .stream_shared(&path)
            .expect("count failure should preserve source stream")
    ));
    assert_eq!(
        editor.finish().expect("count failure should stay exact"),
        original
    );
}

#[test]
fn rendered_batch_matches_recomputed_output_under_both_layout_policies() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let path = vec!["WordDocument".to_string()];
    let replacement: Arc<[u8]> = Arc::from(&b"layout policy replacement"[..]);

    for policy in [SectorLayoutPolicy::Reuse, SectorLayoutPolicy::Rewrite] {
        let mut editor = Editor::open(original.clone(), Targets::default(), Limits::default())
            .expect("editor should open");
        editor.set_sector_layout_policy(policy);
        let rendered = editor
            .put_streams_shared_with_rendered([(path.as_slice(), Arc::clone(&replacement))])
            .expect("layout-policy batch should commit")
            .expect("effective layout-policy batch should render");
        assert_eq!(
            rendered,
            editor
                .clone()
                .finish()
                .expect("layout-policy editor should finish")
        );
    }
}

#[test]
fn add_and_remove_use_explicit_targets() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let mut editor = Editor::open(
        original,
        targets("first", &["ObjectPool", "_42"]),
        Limits::default(),
    )
    .expect("editor should open");
    let nested = write_cfb(|writer| {
        writer
            .create_stream(&["CONTENTS"], b"new object")
            .expect("nested payload should write");
    });
    editor
        .add_storage(target("second", &["ObjectPool", "_43"]), nested)
        .expect("explicit storage should be added");
    assert!(editor.objects().get("second").is_some());
    let removed = editor
        .remove_storage("second")
        .expect("explicit storage should be removed");
    assert!(removed.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]));
    assert!(editor.objects().get("second").is_none());
    assert!(editor.objects().get("first").is_some());
}

#[test]
fn snapshots_share_streams_and_edit_independently() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let selected = targets("object", &["ObjectPool", "_42"]);
    let snapshot = Snapshot::open(original.clone(), selected, Limits::default())
        .expect("snapshot should open");
    let clone = snapshot.clone();
    assert!(!snapshot.is_changed());
    let path = vec!["WordDocument".to_string()];
    let first = snapshot
        .stream_shared(&path)
        .expect("snapshot stream should exist");
    let second = clone
        .stream_shared(&path)
        .expect("cloned snapshot stream should exist");
    assert!(Arc::ptr_eq(&first, &second));

    let mut editor = snapshot.edit();
    editor
        .put_stream(&path, b"edited from snapshot".to_vec())
        .expect("snapshot edit should commit");
    assert!(!snapshot.is_changed());
    assert_eq!(snapshot.finish().expect("source should finish"), original);
    assert_eq!(editor.stream(&path), Some(&b"edited from snapshot"[..]));
}

#[test]
fn changed_snapshot_finish_keeps_the_source_layout_used_by_doc_editors() {
    let base = write_cfb(|writer| {
        writer
            .create_stream(&["A"], &vec![0x11u8; 20_000])
            .expect("source stream should write");
        writer
            .create_stream(&["B"], &vec![0x22u8; 5_000])
            .expect("source stream should write");
    });

    // Move B into A's released sectors and leave A with a shorter allocation.
    // This makes the adopted source physically different from the deterministic
    // from-scratch order while preserving the same logical directory shape.
    let source = {
        let mut writer = OleWriter::new();
        assert!(
            writer
                .adopt_source_layout(&base)
                .expect("source should adopt")
        );
        writer
            .create_stream(&["A"], &vec![0x33u8; 5_000])
            .expect("source edit should write");
        writer
            .create_stream(&["B"], &vec![0x44u8; 20_000])
            .expect("source edit should write");
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("source should write");
        assert!(
            writer
                .last_sector_layout()
                .expect("source report")
                .reused_source_layout()
        );
        output.into_inner()
    };

    let mut editor = Editor::open(source.clone(), Targets::default(), Limits::default())
        .expect("source should open");
    editor
        .put_stream(&["A".into()], vec![0x55u8; 5_000])
        .expect("DOC stream edit should commit");
    let commit = editor.commit().expect("DOC edit should commit");
    assert_eq!(
        commit.snapshot().sector_layout_policy(),
        SectorLayoutPolicy::Reuse
    );
    assert_eq!(
        commit.snapshot().finish().expect("snapshot should finish"),
        commit.patch().after(),
        "a changed source-backed snapshot must use the same layout policy as the DOC save"
    );

    let mut rewrite = OleWriter::new();
    rewrite
        .create_stream(&["A"], &vec![0x55u8; 5_000])
        .expect("rewrite stream should write");
    rewrite
        .create_stream(&["B"], &vec![0x44u8; 20_000])
        .expect("rewrite stream should write");
    let mut rewritten = Cursor::new(Vec::new());
    rewrite
        .write_to(&mut rewritten)
        .expect("rewrite should write");
    let rewritten = rewritten.into_inner();
    assert_ne!(
        commit.patch().after(),
        rewritten.as_slice(),
        "the source-backed route should retain its adopted physical layout"
    );
}
