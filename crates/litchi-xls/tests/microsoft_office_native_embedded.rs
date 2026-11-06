//! Native BIFF object identity, inert child payloads and source-preserving edits.
use litchi_ole_common::object::{OpenOptions, Package as Objects, Target, Targets};
use litchi_xls::{
    ObjSubrecord, Package,
    ole_object::{EmbeddedPayload, Limits, ObjectMetadataChange},
};
use std::{collections::BTreeMap, io::Cursor};
fn native(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/microsoft-office")
            .join(name),
    )
    .unwrap()
}
fn capture(bytes: &[u8], name: &str) -> Objects {
    Objects::open_with_options(
        bytes.to_vec(),
        OpenOptions::new().targets(Targets::new([Target::new(name, [name]).unwrap()]).unwrap()),
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
            let names = path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = ole.open_stream(&names).unwrap();
            (path, data)
        })
        .collect()
}

fn records(bytes: &[u8]) -> Vec<(u16, Vec<u8>)> {
    let mut result = Vec::new();
    let mut pos = 0;
    while pos < bytes.len() {
        let kind = u16::from_le_bytes(bytes[pos..pos + 2].try_into().unwrap());
        let len = usize::from(u16::from_le_bytes(
            bytes[pos + 2..pos + 4].try_into().unwrap(),
        ));
        result.push((kind, bytes[pos + 4..pos + 4 + len].to_vec()));
        pos += 4 + len;
    }
    result
}

fn assert_other_streams(before: &[u8], after: &[u8]) {
    let before = streams(before);
    let after = streams(after);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        after.keys().collect::<Vec<_>>()
    );
    for (path, data) in &before {
        if path != &["Workbook"] {
            assert!(after[path] == *data, "changed {path:?}");
        }
    }
}

#[test]
fn native_embedded_word_obj_resolves_the_exact_cfb_payload() {
    let bytes = native("excel/excel-links-objects.xls");
    let p = Package::from_bytes(&bytes).unwrap();
    let snapshot = p.ole_objects().unwrap();
    assert_eq!(snapshot.worksheet_count(), 1);
    assert_eq!(snapshot.objects(0).unwrap().len(), 1);
    let o = &snapshot.objects(0).unwrap()[0];
    assert_eq!(o.object_id(), 3);
    assert_eq!(o.storage_position(), Some(14497353));
    assert_eq!(o.storage_name().as_deref(), Some("MBD00DD3649"));
    assert!(snapshot.form_controls(0).unwrap().is_empty());
    assert!(
        matches!(&o.subrecords[0],ObjSubrecord::Common(c) if c.object_type==8 && c.object_id==3 && c.flags==24593)
    );
    assert!(matches!(&o.subrecords[1],ObjSubrecord::PictureFormat(c) if c.format==65535));
    assert!(
        matches!(&o.subrecords[2],ObjSubrecord::PictureFlags(c) if c.raw==1 && !c.is_dde() && !c.is_control() && !c.auto_load())
    );
    let objects = capture(&bytes, &o.storage_name().unwrap());
    let object = objects.objects().get("MBD00DD3649").unwrap();
    let payload = object.stream(&["Package"]).unwrap();
    assert_eq!(payload.len(), 13759);
    assert!(payload == native("word/word-legacy-form-unprotected.docx"));
    assert_eq!(object.stream(&["\u{1}CompObj"]).unwrap().len(), 80);
    assert!(p.to_bytes().unwrap() == bytes);
}
#[test]
fn native_obj_metadata_edit_preserves_storage_and_non_obj_records() {
    let bytes = native("excel/excel-links-objects.xls");
    let p = Package::from_bytes(&bytes).unwrap();
    let mut edit = p.edit_ole_objects().unwrap();
    let original = edit.objects(0).unwrap()[0].clone();
    edit.update_object_metadata(0, 3, ObjectMetadataChange::new().with_common_flags(24577))
        .unwrap();
    assert!(matches!(
        edit.update_object_metadata(0, 3, ObjectMetadataChange::new().with_object_id(1)),
        Err(litchi_xls::Error::InvalidRecord {
            record_type: 0x005d,
            ..
        })
    ));
    assert_eq!(edit.objects(0).unwrap()[0].object_id(), 3);
    let commit = edit.commit().unwrap();
    let after = commit.package().to_bytes().unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    let mut expected = original;
    let ObjSubrecord::Common(c) = &mut expected.subrecords[0] else {
        unreachable!()
    };
    c.flags = 24577;
    assert_eq!(
        reopened.ole_objects().unwrap().objects(0).unwrap(),
        [expected]
    );
    assert_other_streams(&bytes, &after);
    let before = streams(&bytes);
    let after = streams(&after);
    let key = vec!["Workbook".to_string()];
    let before = records(&before[&key]);
    let after = records(&after[&key]);
    assert_eq!(before.len(), after.len());
    let changed = before
        .iter()
        .zip(&after)
        .filter(|(a, b)| a != b)
        .collect::<Vec<_>>();
    assert_eq!(changed.len(), 1);
    assert_eq!(changed[0].0.0, 0x005d);
    assert!(
        reopened
            .apply_ole_objects(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
}
#[test]
fn native_embedded_payload_replace_preserves_parent_records_and_inverts() {
    let bytes = native("excel/excel-links-objects.xls");
    let p = Package::from_bytes(&bytes).unwrap();
    let captured = capture(&bytes, "MBD00DD3649");
    let original = captured.objects().get("MBD00DD3649").unwrap();
    let replacement = native("word/word-text-structure.docx");
    let nested = Objects::open(original.compound().to_vec()).unwrap();
    let mut child_edit = nested.edit().unwrap();
    child_edit
        .put_stream(&["Package".into()], replacement.clone())
        .unwrap();
    let child = child_edit
        .commit()
        .unwrap()
        .package()
        .source_shared()
        .as_ref()
        .clone();
    let mut edit = p.edit_ole_objects().unwrap();
    edit.replace_embedded_payload(
        0,
        3,
        EmbeddedPayload::new(child, Limits::default()).unwrap(),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let output = tempfile::NamedTempFile::with_suffix(".xls").unwrap();
    commit.package().save(output.path()).unwrap();
    let after = std::fs::read(output.path()).unwrap();
    let reopened = Package::from_bytes(&after).unwrap();
    assert_eq!(
        reopened.ole_objects().unwrap().objects(0).unwrap(),
        p.ole_objects().unwrap().objects(0).unwrap()
    );
    let before = streams(&bytes);
    let changed = streams(&after);
    assert_eq!(
        before.keys().collect::<Vec<_>>(),
        changed.keys().collect::<Vec<_>>()
    );
    let difference = before
        .iter()
        .filter(|(key, value)| changed[*key] != **value)
        .map(|(key, _)| key.clone())
        .collect::<Vec<_>>();
    assert_eq!(
        difference,
        [vec!["MBD00DD3649".to_string(), "Package".to_string()]]
    );
    let captured = capture(&after, "MBD00DD3649");
    let result = captured.objects().get("MBD00DD3649").unwrap();
    assert!(result.stream(&["Package"]).unwrap() == replacement);
    assert_eq!(result.storage().class_id(), original.storage().class_id());
    assert!(
        reopened
            .apply_ole_objects(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(p.to_bytes().unwrap() == bytes);
}
