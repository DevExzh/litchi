use litchi_ole::{
    LegacyOfficeObjectEditor, LegacyOfficeObjectFormat, LegacyOfficeObjectKind,
    LegacyOfficeObjectLimits, OleFile, OleWriter, discover_legacy_office_objects,
};
use std::io::Cursor;

fn write_cfb(build: impl FnOnce(&mut OleWriter)) -> Vec<u8> {
    let mut writer = OleWriter::new();
    build(&mut writer);
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
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

fn native_package(command: &str, payload: &[u8]) -> Vec<u8> {
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
    let comp_obj = comp_obj("Package", "Package");
    let native = native_package("do-not-run", b"opaque native bytes");
    write_cfb(|writer| {
        writer
            .create_stream(&["WordDocument"], b"unknown-records")
            .unwrap();
        writer.create_storage(&["ObjectPool", "_42"]).unwrap();
        writer
            .create_stream(&["ObjectPool", "_42", "\u{3}ObjInfo"], obj_info)
            .unwrap();
        writer
            .create_stream(&["ObjectPool", "_42", "\u{1}CompObj"], &comp_obj)
            .unwrap();
        writer
            .create_stream(&["ObjectPool", "_42", "\u{1}Ole10Native"], &native)
            .unwrap();
        writer
            .create_stream(&["ObjectPool", "_42", "\u{3}PRINT"], b"metafile")
            .unwrap();
    })
}

#[test]
fn discovers_doc_object_metadata_and_keeps_native_content_inert() {
    let bytes = doc_with_object(&[0x40, 0x00, 0x02, 0x00]);
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let objects = discover_legacy_office_objects(
        &mut ole,
        LegacyOfficeObjectFormat::Doc,
        LegacyOfficeObjectLimits::default(),
    )
    .unwrap();
    let object = objects.find("_42").unwrap();
    assert_eq!(object.internal_storage_reference, Some(42));
    assert_eq!(object.kind, LegacyOfficeObjectKind::Embedded);
    assert_eq!(object.prog_id.as_deref(), Some("Package"));
    assert_eq!(object.display_name.as_deref(), Some("Package"));
    assert!(object.doc_descriptor.unwrap().display_as_icon);
    assert_eq!(
        object.native_package.as_ref().unwrap().command,
        "do-not-run"
    );
    assert_eq!(
        object.native_package.as_ref().unwrap().data,
        b"opaque native bytes"
    );
    assert_eq!(object.previews.len(), 1);
    assert!(object.compound_file.starts_with(&[0xD0, 0xCF, 0x11, 0xE0]));
}

#[test]
fn discovers_xls_link_without_resolving_it() {
    let comp_obj = comp_obj("Linked Worksheet", "Excel.Sheet.8");
    let bytes = write_cfb(|writer| {
        writer.create_stream(&["Workbook"], b"opaque-biff").unwrap();
        writer.create_storage(&["LNK0000002A"]).unwrap();
        writer
            .create_stream(&["LNK0000002A", "\u{1}CompObj"], &comp_obj)
            .unwrap();
    });
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let objects = discover_legacy_office_objects(
        &mut ole,
        LegacyOfficeObjectFormat::Xls,
        LegacyOfficeObjectLimits::default(),
    )
    .unwrap();
    let object = objects.find("LNK0000002A").unwrap();
    assert_eq!(object.internal_storage_reference, Some(42));
    assert_eq!(object.kind, LegacyOfficeObjectKind::Linked);
    assert_eq!(object.prog_id.as_deref(), Some("Excel.Sheet.8"));
    assert_eq!(object.link_metadata.as_deref(), Some("Linked Worksheet"));
}

#[test]
fn rejects_malformed_doc_descriptor_and_resource_exhaustion() {
    let malformed = doc_with_object(&[0x00, 0x04, 0x00, 0x00]);
    let mut ole = OleFile::open(Cursor::new(malformed)).unwrap();
    assert!(
        discover_legacy_office_objects(
            &mut ole,
            LegacyOfficeObjectFormat::Doc,
            LegacyOfficeObjectLimits::default(),
        )
        .is_err()
    );

    let valid = doc_with_object(&[0, 0, 0, 0]);
    let mut ole = OleFile::open(Cursor::new(valid)).unwrap();
    let limits = LegacyOfficeObjectLimits {
        max_stream_size: 4,
        ..LegacyOfficeObjectLimits::default()
    };
    assert!(
        discover_legacy_office_objects(&mut ole, LegacyOfficeObjectFormat::Doc, limits,).is_err()
    );
}

#[test]
fn collection_mutations_are_atomic_and_validate_ids() {
    let bytes = doc_with_object(&[0, 0, 0, 0]);
    let mut ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let mut objects = discover_legacy_office_objects(
        &mut ole,
        LegacyOfficeObjectFormat::Doc,
        LegacyOfficeObjectLimits::default(),
    )
    .unwrap();
    let duplicate = objects.find("_42").unwrap().clone();
    assert!(objects.add(duplicate).is_err());
    assert_eq!(objects.as_slice().len(), 1);
    assert!(
        objects
            .update("_42", |object| {
                object.id.clear();
                Ok(())
            })
            .is_err()
    );
    assert!(objects.find("_42").is_some());
    assert!(objects.reorder(&["missing".to_string()]).is_err());
}

#[test]
fn targeted_replace_preserves_unrelated_streams_and_reference() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let replacement_comp_obj = comp_obj("Worksheet", "Excel.Sheet.8");
    let replacement = write_cfb(|writer| {
        writer
            .create_stream(&["\u{1}CompObj"], &replacement_comp_obj)
            .unwrap();
        writer
            .create_stream(&["CONTENTS"], b"new inert workbook bytes")
            .unwrap();
    });
    let mut editor = LegacyOfficeObjectEditor::open(
        &original,
        LegacyOfficeObjectFormat::Doc,
        LegacyOfficeObjectLimits::default(),
    )
    .unwrap();
    editor.replace("_42", replacement).unwrap();
    assert!(editor.is_changed());
    let output = editor.finish().unwrap();
    let mut ole = OleFile::open(Cursor::new(output)).unwrap();
    assert_eq!(
        ole.open_stream(&["WordDocument"]).unwrap(),
        b"unknown-records"
    );
    assert_eq!(
        ole.open_stream(&["ObjectPool", "_42", "CONTENTS"]).unwrap(),
        b"new inert workbook bytes"
    );
    let objects = discover_legacy_office_objects(
        &mut ole,
        LegacyOfficeObjectFormat::Doc,
        LegacyOfficeObjectLimits::default(),
    )
    .unwrap();
    assert_eq!(
        objects.find("_42").unwrap().prog_id.as_deref(),
        Some("Excel.Sheet.8")
    );
}

#[test]
fn no_op_editor_round_trip_preserves_stream_payloads() {
    let original = doc_with_object(&[0, 0, 0, 0]);
    let editor = LegacyOfficeObjectEditor::open(
        &original,
        LegacyOfficeObjectFormat::Doc,
        LegacyOfficeObjectLimits::default(),
    )
    .unwrap();
    assert!(!editor.is_changed());
    let output = editor.finish().unwrap();
    let mut ole = OleFile::open(Cursor::new(output)).unwrap();
    assert_eq!(
        ole.open_stream(&["WordDocument"]).unwrap(),
        b"unknown-records"
    );
    assert_eq!(
        ole.open_stream(&["ObjectPool", "_42", "\u{1}Ole10Native"])
            .unwrap(),
        native_package("do-not-run", b"opaque native bytes")
    );
}
