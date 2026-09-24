//! Focused publication tests for the OLE-object facade.

use super::super::{
    CheckState, EmbeddedObjectDraft, EmbeddedPayload, FormControl, FtCblsData, FtCf, FtCmo,
    FtPictFmla, FtPioGrbit, Limits, ObjSubrecord, ObjectMetadataEdit, OleObjectRecord,
};
use super::super::{Snapshot, Transaction};
use litchi_cfb::{OleFile, OleWriter};
use litchi_ole_common::property_set::document_summary::DIGITAL_SIGNATURE;
use litchi_ole_common::property_set::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Section, Stream, Value,
};
use std::io::Cursor;

fn record(kind: u16, body: &[u8]) -> Vec<u8> {
    let mut value = kind.to_le_bytes().to_vec();
    value.extend_from_slice(&(body.len() as u16).to_le_bytes());
    value.extend_from_slice(body);
    value
}

fn workbook_stream(controls: &[Vec<u8>]) -> Vec<u8> {
    let bof = record(0x0809, &[0; 16]);
    let eof = record(0x000A, &[]);
    let mut bound_body = vec![0; 8];
    bound_body[6] = 1;
    bound_body[7] = b'S';
    let mut bound = record(0x0085, &bound_body);
    let globals_len = bof.len() + bound.len() + eof.len();
    bound[4..8].copy_from_slice(&(globals_len as u32).to_le_bytes());

    let mut output = bof;
    output.extend_from_slice(&bound);
    output.extend_from_slice(&eof);
    output.extend_from_slice(&record(0x0809, &[0; 16]));
    for control in controls {
        output.extend_from_slice(control);
    }
    output.extend_from_slice(&eof);
    output
}

fn workbook_cfb(controls: &[Vec<u8>]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &workbook_stream(controls))
        .expect("Workbook stream should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("test compound file should be written");
    output.into_inner()
}

fn workbook_cfb_with_unknown_root_stream() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &workbook_stream(&[]))
        .expect("Workbook stream should be created");
    writer
        .create_stream(&["UnrelatedRootStream"], b"retain me")
        .expect("unknown root stream should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("test compound file should be written");
    output.into_inner()
}

fn ole_object(id: u16, position: u32, marker: u8) -> OleObjectRecord {
    OleObjectRecord {
        subrecords: vec![
            ObjSubrecord::Common(FtCmo {
                object_type: 8,
                object_id: id,
                flags: 0x0011,
                reserved: [0xCC; 12],
            }),
            ObjSubrecord::PictureFormat(FtCf {
                format: FtCf::UNSPECIFIED,
            }),
            ObjSubrecord::PictureFlags(FtPioGrbit { raw: 0x0208 }),
            ObjSubrecord::Unknown {
                kind: 0x7777,
                data: vec![marker, 0xA5],
            },
            ObjSubrecord::PictureFormula(FtPictFmla {
                formula: vec![
                    0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00,
                    0x00,
                ],
                storage_position: Some(position),
                control_buffer_size: None,
            }),
            ObjSubrecord::End,
        ],
        text_object: None,
    }
}

fn object_workbook_cfb(object: &OleObjectRecord, payload: &[u8]) -> Vec<u8> {
    object_workbook_cfb_named(object, "MBD0000002A", payload)
}

fn object_workbook_cfb_named(
    object: &OleObjectRecord,
    storage_name: &str,
    payload: &[u8],
) -> Vec<u8> {
    let object_bytes = object.to_record_bytes().unwrap();
    let workbook = workbook_stream(std::slice::from_ref(&object_bytes));
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &workbook)
        .expect("Workbook stream should be created");
    writer
        .create_storage(&[storage_name])
        .expect("object storage should be created");
    writer
        .create_stream(&[storage_name, "Payload"], payload)
        .expect("opaque payload should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("test compound file should be written");
    output.into_inner()
}

fn malformed_opaque_object_record(id: u16, position: u32) -> Vec<u8> {
    let mut cmo = Vec::new();
    cmo.extend_from_slice(&8u16.to_le_bytes());
    cmo.extend_from_slice(&id.to_le_bytes());
    cmo.extend_from_slice(&0u16.to_le_bytes());
    cmo.extend_from_slice(&[0; 12]);
    let mut body = record(0x0015, &cmo);
    body.extend_from_slice(&record(0x0007, &[0, 0, 0])); // malformed FtCf
    body.extend_from_slice(&record(0x0008, &[0, 0]));
    let mut formula = 5u16.to_le_bytes().to_vec();
    formula.extend_from_slice(&[0x05, 0, 0, 0, 0]);
    formula.extend_from_slice(&position.to_le_bytes());
    body.extend_from_slice(&record(0x0009, &formula));
    body.extend_from_slice(&record(0, &[]));
    record(0x005D, &body)
}

fn oversized_valid_opaque_object_record(id: u16, position: u32) -> Vec<u8> {
    let mut cmo = Vec::new();
    cmo.extend_from_slice(&8u16.to_le_bytes());
    cmo.extend_from_slice(&id.to_le_bytes());
    cmo.extend_from_slice(&0u16.to_le_bytes());
    cmo.extend_from_slice(&[0; 12]);
    let mut body = record(0x0015, &cmo);
    body.extend_from_slice(&record(0x0007, &0xFFFFu16.to_le_bytes()));
    body.extend_from_slice(&record(0x0008, &[0, 0]));
    let formula = [
        0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00, 0x00,
    ];
    let mut formula_body = (formula.len() as u16).to_le_bytes().to_vec();
    formula_body.extend_from_slice(&formula);
    formula_body.extend_from_slice(&position.to_le_bytes());
    body.extend_from_slice(&record(0x0009, &formula_body));
    for marker in 0..1_025u16 {
        body.extend_from_slice(&record(0x7777, &[marker as u8]));
    }
    body.extend_from_slice(&record(0, &[]));
    record(0x005D, &body)
}

fn object_workbook_cfb_with_opaque_object(
    object: &OleObjectRecord,
    opaque: &[u8],
    storage_name: &str,
) -> Vec<u8> {
    let object_bytes = object.to_record_bytes().unwrap();
    let workbook = workbook_stream(&[object_bytes, opaque.to_vec()]);
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &workbook)
        .expect("Workbook stream should be created");
    writer
        .create_storage(&[storage_name])
        .expect("object storage should be created");
    writer
        .create_stream(&[storage_name, "Payload"], b"opaque shared payload")
        .expect("object storage payload should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("test compound file should be written");
    output.into_inner()
}

fn objects_workbook_cfb(objects: &[OleObjectRecord], payload: &[u8]) -> Vec<u8> {
    let object_bytes = objects
        .iter()
        .map(|object| object.to_record_bytes().unwrap())
        .collect::<Vec<_>>();
    let workbook = workbook_stream(&object_bytes);
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &workbook)
        .expect("Workbook stream should be created");
    writer
        .create_storage(&["MBD0000002A"])
        .expect("object storage should be created");
    writer
        .create_stream(&["MBD0000002A", "Payload"], payload)
        .expect("opaque payload should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("test compound file should be written");
    output.into_inner()
}

fn embedded_payload(marker: &[u8], unknown: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["\u{0001}Ole"], b"inert OLE metadata")
        .expect("payload metadata should be created");
    writer
        .create_storage(&["OpaqueStorage"])
        .expect("opaque storage should be created");
    writer
        .create_stream(&["OpaqueStorage", "Payload"], marker)
        .expect("opaque payload should be created");
    writer
        .create_stream(&["UnknownStream"], unknown)
        .expect("unknown payload stream should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("payload compound file should be written");
    output.into_inner()
}

fn protected_payload() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["DigitalSignature"], b"signature")
        .expect("signature stream should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("protected payload should be written");
    output.into_inner()
}

fn payload_with_root_stream(name: &str, data: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&[name], data)
        .expect("nested protected stream should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("nested payload CFB should be written");
    output.into_inner()
}

fn nested_pidssi_with_signature() -> Vec<u8> {
    let mut section = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    section.set_page(CodePage::WINDOWS_1252);
    section
        .add(
            DIGITAL_SIGNATURE,
            Value::Unknown {
                variant_type: 0x7F01,
                data: vec![0xAA, 0x55],
            },
        )
        .expect("nested signature property should be accepted");
    payload_with_root_stream(
        "\u{0005}DocumentSummaryInformation",
        &Stream::new(section)
            .to_bytes()
            .expect("nested PIDDSI should serialize"),
    )
}

fn nested_child_pidssi_with_signature() -> Vec<u8> {
    let mut section = Section::new(DOCUMENT_SUMMARY_INFORMATION_FMTID);
    section.set_page(CodePage::WINDOWS_1252);
    section
        .add(
            DIGITAL_SIGNATURE,
            Value::Unknown {
                variant_type: 0x7F01,
                data: vec![0xAA, 0x55],
            },
        )
        .expect("nested child signature property should be accepted");
    let data = Stream::new(section)
        .to_bytes()
        .expect("nested child PIDDSI should serialize");
    let mut writer = OleWriter::new();
    writer
        .create_storage(&["Child"])
        .expect("nested child storage should be created");
    writer
        .create_stream(&["Child", "\u{0005}DocumentSummaryInformation"], &data)
        .expect("nested child PIDDSI stream should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("nested child payload CFB should be written");
    output.into_inner()
}

fn deeply_nested_payload() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_storage(&["Level1"])
        .expect("first payload storage should be created");
    writer
        .create_storage(&["Level1", "Level2"])
        .expect("second payload storage should be created");
    writer
        .create_stream(&["Level1", "Level2", "Payload"], b"nested")
        .expect("nested payload stream should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("nested payload CFB should be written");
    output.into_inner()
}

fn wide_payload(siblings: usize) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for index in 0..siblings {
        let name = format!("S{index:03}");
        writer
            .create_storage(&[name.as_str()])
            .expect("sibling payload storage should be created");
        writer
            .create_stream(&[name.as_str(), "Payload"], b"x")
            .expect("sibling payload stream should be created");
    }
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("wide payload CFB should be written");
    output.into_inner()
}

fn publication_limits() -> Limits {
    Limits {
        max_objects: 8,
        max_storage_depth: 8,
        max_streams_per_object: 16,
        max_streams: 16,
        max_stream_size: 1024,
        max_object_size: 1024 * 1024,
        max_total_size: 2 * 1024 * 1024,
    }
}

fn checkbox_control(id: u16, marker: u8) -> FormControl {
    FormControl {
        subrecords: vec![
            ObjSubrecord::Common(FtCmo {
                object_type: 0x000B,
                object_id: id,
                flags: 0,
                reserved: [0xCC; 12],
            }),
            ObjSubrecord::CheckBoxData(FtCblsData {
                state: CheckState::Checked,
                accelerator: 0,
                reserved: 0,
                flags: 1,
            }),
            ObjSubrecord::Unknown {
                kind: 0x7777,
                data: vec![marker],
            },
            ObjSubrecord::End,
        ],
        text_object: None,
    }
}

#[test]
fn snapshot_and_noop_commit_preserve_exact_cfb_bytes() {
    let existing = checkbox_control(7, 0xA1).to_record_bytes().unwrap();
    let input = workbook_cfb(std::slice::from_ref(&existing));
    let snapshot = Snapshot::open(input.clone(), Limits::default()).unwrap();

    assert_eq!(snapshot.finish(), input);
    let commit = snapshot.edit().commit().unwrap();
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.snapshot().finish(), input);
    assert_eq!(commit.patch().apply(&snapshot).unwrap().finish(), input);
}

#[test]
fn invalid_control_edit_is_failure_atomic() {
    let existing = checkbox_control(7, 0xA1).to_record_bytes().unwrap();
    let snapshot = Snapshot::open(
        workbook_cfb(std::slice::from_ref(&existing)),
        Limits::default(),
    )
    .unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();

    let result = transaction.add_form_control(0, checkbox_control(7, 0xB2));
    assert!(result.is_err());
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
    assert_eq!(transaction.form_controls(0).unwrap().len(), 1);
}

#[test]
fn valid_control_edit_publishes_typed_patch_and_keeps_unknowns() {
    let existing = checkbox_control(7, 0xA1).to_record_bytes().unwrap();
    let input = workbook_cfb(std::slice::from_ref(&existing));
    let snapshot = Snapshot::open(input, Limits::default()).unwrap();
    let authored = checkbox_control(8, 0xB2);
    let authored_bytes = authored.to_record_bytes().unwrap();

    let mut transaction: Transaction = snapshot.edit();
    transaction.add_form_control(0, authored).unwrap();
    let commit = transaction.commit().unwrap();

    assert!(commit.changed());
    assert_eq!(commit.snapshot().form_controls(0).unwrap().len(), 2);
    let output = commit.snapshot().finish();
    assert!(
        output
            .windows(existing.len())
            .any(|window| window == existing)
    );
    assert!(
        output
            .windows(authored_bytes.len())
            .any(|window| window == authored_bytes)
    );
    let applied = commit.patch().apply(&snapshot).unwrap();
    assert_eq!(applied.finish(), output);
    assert!(commit.patch().apply(&applied).is_err());
    assert_eq!(
        commit.patch().inverse().apply(&applied).unwrap().finish(),
        snapshot.finish()
    );
}

#[test]
fn empty_object_metadata_edit_preserves_exact_bytes() {
    let source = ole_object(7, 0x2A, 0xB1);
    let input = object_workbook_cfb(&source, b"raw embedded payload");
    let snapshot = Snapshot::open(input.clone(), Limits::default()).unwrap();
    let mut transaction = snapshot.edit();

    transaction
        .update_object_metadata(0, 7, ObjectMetadataEdit::new())
        .unwrap();
    let commit = transaction.commit().unwrap();

    assert!(!commit.changed());
    assert_eq!(commit.snapshot().finish(), input);
}

#[test]
fn object_metadata_edit_changes_typed_fields_and_preserves_unknown_payload() {
    let source = ole_object(7, 0x2A, 0xB1);
    let input = object_workbook_cfb(&source, b"raw embedded payload");
    let snapshot = Snapshot::open(input, Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    transaction
        .update_object_metadata(
            0,
            7,
            ObjectMetadataEdit::new()
                .with_object_id(9)
                .with_common_flags(0xA55A)
                .with_picture_flags(FtPioGrbit {
                    raw: 0x0208 | 0x0001,
                }),
        )
        .unwrap();
    let commit = transaction.commit().unwrap();
    let object = &commit.snapshot().objects(0).unwrap()[0];

    assert_eq!(object.object_id(), 9);
    assert!(
        object
            .subrecords
            .iter()
            .any(|value| { matches!(value, ObjSubrecord::Common(FtCmo { flags: 0xA55A, .. })) })
    );
    assert!(object.subrecords.iter().any(|value| {
        matches!(
            value,
            ObjSubrecord::PictureFlags(FtPioGrbit { raw: 0x0209 })
        )
    }));
    assert!(object.subrecords.iter().any(|value| {
        matches!(value, ObjSubrecord::Unknown { kind: 0x7777, data } if data == &[0xB1, 0xA5])
    }));
    assert_eq!(object.storage_position(), Some(0x2A));

    let mut ole = OleFile::open(Cursor::new(commit.snapshot().finish())).unwrap();
    assert_eq!(
        ole.open_stream(&["MBD0000002A", "Payload"]).unwrap(),
        b"raw embedded payload"
    );
}

#[test]
fn metadata_edit_rejects_stale_identity_and_invalid_flags_atomically() {
    let source = ole_object(7, 0x2A, 0xB1);
    let input = object_workbook_cfb(&source, b"raw embedded payload");
    let snapshot = Snapshot::open(input, Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();

    assert!(
        transaction
            .update_object_metadata(0, 8, ObjectMetadataEdit::new().with_common_flags(0xAAAA),)
            .is_err()
    );
    assert!(
        transaction
            .update_object_metadata(
                0,
                7,
                ObjectMetadataEdit::new().with_picture_flags(FtPioGrbit { raw: 0x0012 }),
            )
            .is_err()
    );
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
}

#[test]
fn metadata_edit_rejects_storage_class_transition_atomically() {
    let source = ole_object(7, 0x2A, 0xB1);
    let input = object_workbook_cfb(&source, b"raw embedded payload");
    let snapshot = Snapshot::open(input, Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();

    assert!(
        transaction
            .update_object_metadata(
                0,
                7,
                ObjectMetadataEdit::new().with_picture_flags(FtPioGrbit { raw: 0x020A }),
            )
            .is_err()
    );
    assert!(
        transaction
            .update_object_metadata(
                0,
                7,
                ObjectMetadataEdit::new().with_picture_flags(FtPioGrbit { raw: 0x0288 }),
            )
            .is_err()
    );
    assert!(
        transaction
            .update_object_metadata(
                0,
                7,
                ObjectMetadataEdit::new().with_picture_flags(FtPioGrbit { raw: 0x0228 }),
            )
            .is_err()
    );
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
    assert_eq!(
        transaction.objects(0).unwrap()[0].storage_name().as_deref(),
        Some("MBD0000002A")
    );
}

#[test]
fn metadata_patch_rejects_stale_snapshot_without_mutation() {
    let source = ole_object(7, 0x2A, 0xB1);
    let input = object_workbook_cfb(&source, b"raw embedded payload");
    let snapshot = Snapshot::open(input.clone(), Limits::default()).unwrap();
    let stale = Snapshot::open(
        object_workbook_cfb(&ole_object(8, 0x2A, 0xB1), b"raw embedded payload"),
        Limits::default(),
    )
    .unwrap();
    let stale_before = stale.finish();
    let mut transaction = snapshot.edit();
    transaction
        .update_object_metadata(0, 7, ObjectMetadataEdit::new().with_common_flags(0xA55A))
        .unwrap();
    let commit = transaction.commit().unwrap();

    assert!(commit.patch().apply(&stale).is_err());
    assert_eq!(stale.finish(), stale_before);
}

#[test]
fn add_embedded_payload_binds_obj_and_mbd_and_preserves_opaque_streams() {
    let input = workbook_cfb_with_unknown_root_stream();
    let payload_bytes = embedded_payload(b"payload-v1", b"unknown-v1");
    let payload = EmbeddedPayload::new(payload_bytes).expect("payload CFB should validate");
    let draft = EmbeddedObjectDraft::new(9, 0x2A, payload).expect("draft should validate");
    let snapshot = Snapshot::open(input.clone(), Limits::default()).expect("workbook should open");
    let mut transaction = snapshot.edit();

    transaction
        .add_embedded_payload(0, draft)
        .expect("typed embedded payload should publish");
    let commit = transaction.commit().expect("commit should validate");
    assert!(commit.changed());
    assert_eq!(
        commit.snapshot().objects(0).unwrap()[0]
            .storage_name()
            .as_deref(),
        Some("MBD0000002A")
    );
    assert!(matches!(
        commit.snapshot().objects(0).unwrap()[0]
            .subrecords
            .iter()
            .find_map(|value| match value {
                ObjSubrecord::PictureFormula(value) => Some(value),
                _ => None,
            }),
        Some(FtPictFmla {
            storage_position: Some(0x2A),
            control_buffer_size: None,
            ..
        })
    ));
    let formula = commit.snapshot().objects(0).unwrap()[0]
        .subrecords
        .iter()
        .find_map(|value| match value {
            ObjSubrecord::PictureFormula(value) => Some(value),
            _ => None,
        })
        .expect("embedded object should have FtPictFmla");
    assert_eq!(
        formula.formula,
        [
            0x05, 0x00, // ObjectParsedFormula.cce = 5
            0x00, 0x00, 0x00, 0x00, // ObjectParsedFormula unused
            0x02, 0x00, 0x00, 0x00, 0x00, // PtgTbl
            0x03, 0x00, 0x00, // PictFmlaEmbedInfo (no class name)
        ]
    );
    assert_eq!(formula.formula.len() % 2, 0);

    let mut ole = OleFile::open(Cursor::new(commit.snapshot().finish())).expect("CFB should open");
    assert_eq!(
        ole.open_stream(&["UnrelatedRootStream"]).unwrap(),
        b"retain me"
    );
    assert_eq!(
        ole.open_stream(&["MBD0000002A", "OpaqueStorage", "Payload"])
            .unwrap(),
        b"payload-v1"
    );
    assert_eq!(
        ole.open_stream(&["MBD0000002A", "UnknownStream"]).unwrap(),
        b"unknown-v1"
    );
    let applied = commit
        .patch()
        .apply(&snapshot)
        .expect("patch source should match");
    assert_eq!(applied.finish(), commit.snapshot().finish());
    assert_eq!(
        commit.patch().inverse().apply(&applied).unwrap().finish(),
        input
    );
}

#[test]
fn replacing_with_the_original_embedded_payload_is_an_exact_noop() {
    let input = workbook_cfb(&[]);
    let payload_bytes = embedded_payload(b"payload-v1", b"unknown-v1");
    let payload = EmbeddedPayload::new(payload_bytes.clone()).expect("payload CFB should validate");
    let draft = EmbeddedObjectDraft::new(9, 0x2A, payload).expect("draft should validate");
    let snapshot = Snapshot::open(input, Limits::default()).expect("workbook should open");
    let mut add = snapshot.edit();
    add.add_embedded_payload(0, draft)
        .expect("typed embedded payload should publish");
    let added_commit = add.commit().expect("add should commit");
    let added = added_commit.snapshot();

    let mut replace = added.edit();
    replace
        .replace_embedded_payload(
            0,
            9,
            EmbeddedPayload::new(payload_bytes).expect("payload CFB should validate"),
        )
        .expect("same payload should be accepted");
    let commit = replace.commit().expect("no-op replacement should commit");
    assert!(!commit.changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.snapshot().finish(), added.finish());
}

#[test]
fn replace_embedded_payload_retains_identity_and_replaces_only_selected_storage() {
    let source_object = ole_object(7, 0x2A, 0xB1);
    let input = object_workbook_cfb(&source_object, b"source payload");
    let replacement = embedded_payload(b"payload-v2", b"unknown-v2");
    let snapshot = Snapshot::open(input, Limits::default()).expect("workbook should open");
    let mut transaction = snapshot.edit();

    transaction
        .replace_embedded_payload(
            0,
            7,
            EmbeddedPayload::new(replacement).expect("replacement CFB should validate"),
        )
        .expect("replacement should publish");
    let commit = transaction.commit().expect("commit should validate");
    let object = &commit.snapshot().objects(0).unwrap()[0];
    assert_eq!(object.object_id(), 7);
    assert_eq!(object.storage_name().as_deref(), Some("MBD0000002A"));
    assert!(object.subrecords.iter().any(|value| {
        matches!(value, ObjSubrecord::Unknown { kind: 0x7777, data } if data == &[0xB1, 0xA5])
    }));

    let mut ole = OleFile::open(Cursor::new(commit.snapshot().finish())).expect("CFB should open");
    assert_eq!(
        ole.open_stream(&["MBD0000002A", "OpaqueStorage", "Payload"])
            .unwrap(),
        b"payload-v2"
    );
    assert_eq!(
        ole.open_stream(&["MBD0000002A", "UnknownStream"]).unwrap(),
        b"unknown-v2"
    );
    assert!(ole.open_stream(&["MBD0000002A", "Payload"]).is_err());
}

#[test]
fn remove_embedded_payload_removes_unreferenced_storage_and_is_reversible() {
    let source_object = ole_object(7, 0x2A, 0xB1);
    let input = object_workbook_cfb(&source_object, b"source payload");
    let snapshot = Snapshot::open(input.clone(), Limits::default()).expect("workbook should open");
    let mut transaction = snapshot.edit();

    transaction
        .remove_embedded_payload(0, 7)
        .expect("embedded object should be removable");
    let commit = transaction.commit().expect("commit should validate");
    assert!(commit.snapshot().objects(0).unwrap().is_empty());
    let mut ole = OleFile::open(Cursor::new(commit.snapshot().finish())).expect("CFB should open");
    assert!(ole.open_stream(&["MBD0000002A"]).is_err());
    assert_eq!(
        commit
            .patch()
            .inverse()
            .apply(commit.snapshot())
            .unwrap()
            .finish(),
        input
    );
}

#[test]
fn removing_one_of_two_shared_mbd_references_keeps_storage_until_last_owner() {
    let input = objects_workbook_cfb(
        &[ole_object(7, 0x2A, 0xB1), ole_object(8, 0x2A, 0xB2)],
        b"shared payload",
    );
    let snapshot = Snapshot::open(input, Limits::default()).expect("workbook should open");
    let mut first = snapshot.edit();
    first
        .remove_embedded_payload(0, 7)
        .expect("first owner should be removable");
    let first_commit = first.commit().expect("first removal should commit");
    assert_eq!(first_commit.snapshot().objects(0).unwrap().len(), 1);
    let mut first_ole = OleFile::open(Cursor::new(first_commit.snapshot().finish()))
        .expect("first result should be a CFB");
    assert_eq!(
        first_ole.open_stream(&["MBD0000002A", "Payload"]).unwrap(),
        b"shared payload"
    );

    let mut second = first_commit.snapshot().edit();
    second
        .remove_embedded_payload(0, 8)
        .expect("last owner should be removable");
    let second_commit = second.commit().expect("second removal should commit");
    assert!(second_commit.snapshot().objects(0).unwrap().is_empty());
    let mut second_ole = OleFile::open(Cursor::new(second_commit.snapshot().finish()))
        .expect("second result should be a CFB");
    assert!(second_ole.open_stream(&["MBD0000002A"]).is_err());
}

#[test]
fn embedded_payload_api_rejects_dde_and_invalid_payloads_atomically() {
    assert!(EmbeddedPayload::new(vec![1, 2, 3]).is_err());
    assert!(
        EmbeddedObjectDraft::new(
            0,
            0x2A,
            EmbeddedPayload::new(embedded_payload(b"payload", b"unknown")).unwrap(),
        )
        .is_err()
    );

    let dde = OleObjectRecord {
        subrecords: vec![
            ObjSubrecord::Common(FtCmo {
                object_type: 8,
                object_id: 7,
                flags: 0,
                reserved: [0; 12],
            }),
            ObjSubrecord::PictureFormat(FtCf {
                format: FtCf::UNSPECIFIED,
            }),
            ObjSubrecord::PictureFlags(FtPioGrbit { raw: 0x0002 }),
            ObjSubrecord::PictureFormula(FtPictFmla {
                formula: vec![0x05, 0, 0, 0, 0],
                storage_position: Some(0x2A),
                control_buffer_size: Some(0),
            }),
            ObjSubrecord::End,
        ],
        text_object: None,
    };
    let snapshot = Snapshot::open(
        object_workbook_cfb_named(&dde, "LNK0000002A", b"link payload"),
        Limits::default(),
    )
    .expect("DDE object should remain readable");
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();
    assert!(transaction.remove_embedded_payload(0, 7).is_err());
    assert!(
        transaction
            .replace_embedded_payload(
                0,
                7,
                EmbeddedPayload::new(embedded_payload(b"new", b"unknown")).unwrap(),
            )
            .is_err()
    );
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
}

#[test]
fn add_embedded_payload_rejects_duplicate_object_or_storage_identity_atomically() {
    let source = ole_object(7, 0x2A, 0xB1);
    let snapshot = Snapshot::open(
        object_workbook_cfb(&source, b"source payload"),
        Limits::default(),
    )
    .expect("workbook should open");
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();

    let same_object = EmbeddedObjectDraft::new(
        7,
        0x2A,
        EmbeddedPayload::new(embedded_payload(b"new", b"unknown")).unwrap(),
    )
    .unwrap();
    assert!(transaction.add_embedded_payload(0, same_object).is_err());
    assert_eq!(transaction.snapshot().unwrap().finish(), before);

    let same_storage = EmbeddedObjectDraft::new(
        8,
        0x2A,
        EmbeddedPayload::new(embedded_payload(b"new", b"unknown")).unwrap(),
    )
    .unwrap();
    assert!(transaction.add_embedded_payload(0, same_storage).is_err());
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
}

#[test]
fn malformed_opaque_obj_is_lossless_but_blocks_unproven_identity_edits() {
    let source = ole_object(7, 0x2A, 0xB1);
    let opaque = malformed_opaque_object_record(9, 0x2B);
    let input = object_workbook_cfb_with_opaque_object(&source, &opaque, "MBD0000002A");
    let snapshot = Snapshot::open(input.clone(), Limits::default())
        .expect("malformed opaque Obj should remain readable");
    assert_eq!(snapshot.objects(0).unwrap().len(), 1);
    assert_eq!(snapshot.finish(), input);

    let payload = EmbeddedPayload::new(embedded_payload(b"new", b"unknown")).unwrap();
    let draft = EmbeddedObjectDraft::new(9, 0x2B, payload).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();
    assert!(transaction.add_embedded_payload(0, draft).is_err());
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
    assert!(transaction.remove_embedded_payload(0, 7).is_err());
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
}

#[test]
fn duplicate_opaque_mbd_identity_within_one_sheet_blocks_topology_edits() {
    let source = ole_object(7, 0x2B, 0xB1);
    let object_bytes = source.to_record_bytes().unwrap();
    let opaque_a = oversized_valid_opaque_object_record(9, 0x2A);
    let opaque_b = oversized_valid_opaque_object_record(10, 0x2A);
    let workbook = workbook_stream(&[object_bytes, opaque_a, opaque_b]);
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    writer.create_storage(&["MBD0000002B"]).unwrap();
    writer
        .create_stream(&["MBD0000002B", "Payload"], b"typed payload")
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let input = output.into_inner();

    let snapshot = Snapshot::open(input.clone(), Limits::default())
        .expect("oversized opaque records should remain byte-preserved");
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();
    let draft = EmbeddedObjectDraft::new(
        11,
        0x2C,
        EmbeddedPayload::new(embedded_payload(b"new", b"opaque")).unwrap(),
    )
    .unwrap();
    assert!(transaction.add_embedded_payload(0, draft).is_err());
    assert!(transaction.remove_embedded_payload(0, 7).is_err());
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
}

#[test]
fn add_embedded_payload_preflights_target_limit_before_payload_capture() {
    let source = ole_object(7, 0x2A, 0xB1);
    let limits = Limits {
        max_objects: 1,
        ..publication_limits()
    };
    let snapshot = Snapshot::open(object_workbook_cfb(&source, b"source payload"), limits)
        .expect("one existing target should fit the configured limit");
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();
    let payload = EmbeddedPayload::new(embedded_payload(b"new", b"unknown")).unwrap();
    let draft = EmbeddedObjectDraft::new(8, 0x2B, payload).unwrap();
    let error = transaction
        .add_embedded_payload(0, draft)
        .expect_err("a second target must be refused before capture");
    assert!(
        matches!(error, crate::error::Error::InvalidData(message) if message.contains("object limit"))
    );
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
}

#[test]
fn embedded_payload_api_refuses_protected_outer_and_nested_cfb() {
    let protected = {
        let mut writer = OleWriter::new();
        writer
            .create_stream(&["Workbook"], &workbook_stream(&[]))
            .unwrap();
        writer
            .create_stream(&["DigitalSignature"], b"signature")
            .unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    };
    assert!(Snapshot::open(protected, Limits::default()).is_err());

    let payload = EmbeddedPayload::new(protected_payload()).expect("payload CFB parses");
    let draft = EmbeddedObjectDraft::new(7, 0x2A, payload).unwrap();
    let snapshot = Snapshot::open(workbook_cfb(&[]), Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();
    assert!(transaction.add_embedded_payload(0, draft).is_err());
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
}

#[test]
fn embedded_payload_api_refuses_nested_pidssi_and_encryption() {
    let snapshot = Snapshot::open(workbook_cfb(&[]), Limits::default()).unwrap();
    for bytes in [
        nested_pidssi_with_signature(),
        nested_child_pidssi_with_signature(),
        payload_with_root_stream("encryption", b"opaque encrypted payload"),
    ] {
        let payload = EmbeddedPayload::new(bytes).expect("nested CFB should parse");
        let draft = EmbeddedObjectDraft::new(7, 0x2A, payload).unwrap();
        let mut transaction = snapshot.edit();
        let before = transaction.snapshot().unwrap().finish();
        let error = transaction
            .add_embedded_payload(0, draft)
            .expect_err("nested protection must refuse publication");
        assert!(matches!(error, crate::error::Error::UnsafeEdit(_)));
        assert_eq!(transaction.snapshot().unwrap().finish(), before);
    }
}

#[test]
fn embedded_payload_checks_stream_bounds_before_parsing_nested_pidssi() {
    let limits = Limits {
        max_stream_size: 1024,
        ..publication_limits()
    };
    let bytes = payload_with_root_stream("\u{0005}DocumentSummaryInformation", &vec![0xA5; 2_048]);
    let payload = EmbeddedPayload::with_limits(bytes, limits).expect("header should be admitted");
    let draft = EmbeddedObjectDraft::new(9, 0x2A, payload).unwrap();
    let snapshot = Snapshot::open(workbook_cfb(&[]), limits).unwrap();
    let mut transaction = snapshot.edit();
    let error = transaction
        .add_embedded_payload(0, draft)
        .expect_err("oversized PIDDSI must be refused before parsing");
    assert!(
        matches!(error, crate::error::Error::InvalidData(message) if message.contains("stream exceeds"))
    );
}

#[test]
fn embedded_payload_publication_applies_common_count_size_and_depth_limits() {
    let cases = [
        (
            "per-object stream count",
            Limits {
                max_streams_per_object: 2,
                ..publication_limits()
            },
            embedded_payload(b"payload", b"unknown"),
        ),
        (
            "stream size",
            Limits {
                max_stream_size: 1024,
                ..publication_limits()
            },
            embedded_payload(&vec![b'x'; 1025], b"unknown"),
        ),
        (
            "storage depth",
            Limits {
                max_storage_depth: 1,
                ..publication_limits()
            },
            deeply_nested_payload(),
        ),
    ];

    for (name, limits, bytes) in cases {
        let payload = EmbeddedPayload::with_limits(bytes, limits)
            .unwrap_or_else(|error| panic!("{name} should pass standalone admission: {error:?}"));
        let draft = EmbeddedObjectDraft::new(9, 0x2A, payload).unwrap();
        let snapshot = Snapshot::open(workbook_cfb(&[]), limits)
            .unwrap_or_else(|error| panic!("{name} workbook should open: {error:?}"));
        let mut transaction = snapshot.edit();
        let before = transaction.snapshot().unwrap().finish();
        assert!(
            transaction.add_embedded_payload(0, draft).is_err(),
            "{name} should be enforced by workbook publication"
        );
        assert_eq!(transaction.snapshot().unwrap().finish(), before);
    }
}

#[test]
fn embedded_payload_publication_bounds_wide_storage_siblings_before_capture() {
    let limits = Limits {
        max_objects: 1,
        max_storage_depth: 2,
        max_streams_per_object: 32,
        max_streams: 32,
        ..publication_limits()
    };
    let payload = EmbeddedPayload::with_limits(wide_payload(16), limits)
        .expect("standalone payload admission should only parse the CFB header");
    let draft = EmbeddedObjectDraft::new(9, 0x2A, payload).unwrap();
    let snapshot = Snapshot::open(workbook_cfb(&[]), limits).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction.snapshot().unwrap().finish();
    let error = transaction
        .add_embedded_payload(0, draft)
        .expect_err("wide sibling directories must be bounded before capture");
    assert!(
        matches!(error, crate::error::Error::InvalidData(message) if message.contains("storage count"))
    );
    assert_eq!(transaction.snapshot().unwrap().finish(), before);
}

#[test]
fn embedded_payload_uses_per_object_storage_ceiling_before_capture() {
    let limits = Limits {
        max_objects: 8,
        max_storage_depth: 1,
        max_streams_per_object: 8,
        max_streams: 8,
        ..publication_limits()
    };
    let payload = EmbeddedPayload::with_limits(wide_payload(2), limits)
        .expect("standalone payload header should be admitted");
    let draft = EmbeddedObjectDraft::new(9, 0x2A, payload).unwrap();
    let snapshot = Snapshot::open(workbook_cfb(&[]), limits).unwrap();
    let mut transaction = snapshot.edit();
    let error = transaction
        .add_embedded_payload(0, draft)
        .expect_err("one selected payload must use the per-object storage ceiling");
    assert!(
        matches!(error, crate::error::Error::InvalidData(message) if message.contains("storage count"))
    );
}
