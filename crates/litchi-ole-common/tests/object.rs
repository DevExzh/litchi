use litchi_cfb::{OleFile, OleWriter};
use litchi_ole_common::object::{Editor, Limits, Snapshot, Target, Targets, discover};
use std::io::Cursor;
use std::sync::Arc;

fn write_cfb(build: impl FnOnce(&mut OleWriter)) -> Vec<u8> {
    let mut writer = OleWriter::new();
    build(&mut writer);
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).expect("test CFB should write");
    output.into_inner()
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
    assert_eq!(object.streams().len(), 4);
    assert!(object.stream(&["\u{1}Ole10Native"]).is_some());
    assert!(object.compound().starts_with(&[0xD0, 0xCF, 0x11, 0xE0]));
    assert_eq!(objects.at(0).map(|value| value.key()), Some("host-object"));
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
fn shared_stream_replacement_reuses_validated_allocation() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let mut editor = Editor::open(
        original,
        targets("object", &["ObjectPool", "_42"]),
        Limits::default(),
    )
    .expect("editor should open");
    let path = vec!["WordDocument".to_string()];
    let replacement: Arc<[u8]> = Arc::from(&b"shared-word-stream"[..]);
    editor
        .put_stream_shared(&path, Arc::clone(&replacement))
        .expect("stream replacement should commit");
    let installed = editor
        .stream_shared(&path)
        .expect("stream should remain available");
    assert!(Arc::ptr_eq(&replacement, &installed));
    assert_eq!(editor.stream(&path), Some(&b"shared-word-stream"[..]));
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
