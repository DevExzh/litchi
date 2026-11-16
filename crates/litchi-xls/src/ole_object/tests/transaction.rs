//! Focused publication tests for the OLE-object facade.

use super::super::edit::RecordCommit;
use super::super::{
    CheckState, EmbeddedObjectDraft, EmbeddedPayload, FormControl, FtCblsData, FtCf, FtCmo,
    FtPictFmla, FtPioGrbit, Limits, ObjSubrecord, ObjectMetadataChange, OleObjectRecord,
};
use super::super::{Edit, Snapshot};
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

const PNG: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x60, 0x00, 0x02, 0x00,
    0x00, 0x05, 0x00, 0x01, 0xe9, 0xfa, 0xdc, 0xd8, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44,
    0xae, 0x42, 0x60, 0x82,
];

fn anchor() -> crate::sheet_features::edit::Anchor {
    crate::sheet_features::edit::Anchor {
        first_row: 2,
        first_column: 1,
        last_row: 6,
        last_column: 3,
        behavior: crate::AnchorBehavior::MoveWithoutResize,
    }
}

fn draft(payload: EmbeddedPayload) -> EmbeddedObjectDraft {
    EmbeddedObjectDraft::new(payload, anchor(), PNG.to_vec())
}

/// The streams of a compound file, by path.
fn cfb_streams(bytes: &[u8]) -> Vec<(Vec<String>, Vec<u8>)> {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    ole.list_streams()
        .into_iter()
        .map(|path| {
            let parts: Vec<&str> = path.iter().map(String::as_str).collect();
            let data = ole.open_stream(&parts).unwrap();
            (path, data)
        })
        .collect()
}

/// A compound file of `streams`.
fn cfb_of(streams: &[(Vec<String>, Vec<u8>)]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for (path, data) in streams {
        let parts: Vec<&str> = path.iter().map(String::as_str).collect();
        for depth in 1..parts.len() {
            let _ = writer.create_storage(&parts[..depth]);
        }
        writer.create_stream(&parts, data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

/// A complete one-worksheet workbook (it has the records a drawing layer
/// is written beside), with `extra` root streams.
fn drawable_workbook_cfb(extra: &[(&str, &[u8])]) -> Vec<u8> {
    let bytes = crate::Package::new().unwrap().to_bytes().unwrap();
    let mut streams = cfb_streams(&bytes);
    for (name, data) in extra {
        streams.push((vec![(*name).to_owned()], data.to_vec()));
    }
    cfb_of(&streams)
}

/// `input` with one embedded object per payload added through the
/// OLE-object edit; returns the bytes and the new object identifiers.
fn with_embedded_objects(input: Vec<u8>, payloads: &[&[u8]]) -> (Vec<u8>, Vec<u16>) {
    let snapshot = Snapshot::open(input, Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    let ids = payloads
        .iter()
        .map(|payload| {
            transaction
                .add_embedded_payload(
                    0,
                    draft(EmbeddedPayload::new(payload.to_vec(), Limits::default()).unwrap()),
                )
                .unwrap()
        })
        .collect();
    let bytes = transaction
        .commit_record()
        .unwrap()
        .into_snapshot()
        .finish();
    (bytes, ids)
}

/// The Workbook stream of a compound file.
fn workbook_of(bytes: &[u8]) -> Vec<u8> {
    let mut ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    ole.open_stream(&["Workbook"]).unwrap()
}

/// The record types of the worksheet substreams, in order.
fn worksheet_records(workbook: &[u8]) -> Vec<u16> {
    let mut out = Vec::new();
    let mut at = 0;
    let mut substream = 0;
    while at + 4 <= workbook.len() {
        let kind = u16::from_le_bytes([workbook[at], workbook[at + 1]]);
        let length = usize::from(u16::from_le_bytes([workbook[at + 2], workbook[at + 3]]));
        if kind == 0x0809 {
            substream += 1;
        }
        if substream > 1 {
            out.push(kind);
        }
        at += 4 + length;
    }
    out
}

/// Every `Obj` of the worksheets comes right after its `MsoDrawing` shape
/// ([MS-XLS] 2.4.181 and the `OBJECTS` grammar), and every shape that is
/// followed by `Obj` has one; returns the number of `Obj` records.
fn assert_objects_follow_their_shapes(bytes: &[u8]) -> usize {
    let records = worksheet_records(&workbook_of(bytes));
    let mut count = 0;
    for (index, kind) in records.iter().enumerate() {
        if *kind == 0x005D {
            count += 1;
            assert_eq!(
                records.get(index.wrapping_sub(1)),
                Some(&0x00EC),
                "an Obj is not preceded by its MsoDrawing shape"
            );
        }
    }
    count
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
    let commit = snapshot.edit().commit_record().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.snapshot().finish(), input);
    assert_eq!(commit.patch().apply(&snapshot).unwrap().finish(), input);
}

#[test]
fn invalid_control_edit_is_failure_atomic() {
    let snapshot = Snapshot::open(drawable_workbook_cfb(&[]), Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction
        .clone()
        .commit_record()
        .map(RecordCommit::into_snapshot)
        .unwrap()
        .finish();

    // A label is not written, so a control with a text object is refused.
    let labelled = checkbox_control(7, 0xB2).with_text_object(record(0x01B6, &[0; 18]));
    assert!(transaction.add_form_control(0, labelled, anchor()).is_err());
    let mut outside = anchor();
    outside.last_column = 256;
    assert!(
        transaction
            .add_form_control(0, checkbox_control(7, 0xB2), outside)
            .is_err()
    );
    assert_eq!(
        transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish(),
        before
    );
    assert!(transaction.form_controls(0).unwrap().is_empty());
}

#[test]
fn valid_control_edit_publishes_typed_patch_and_keeps_unknowns() {
    let snapshot = Snapshot::open(drawable_workbook_cfb(&[]), Limits::default()).unwrap();
    let mut first = snapshot.edit();
    let first_id = first
        .add_form_control(0, checkbox_control(7, 0xA1), anchor())
        .unwrap();
    let source = first.commit_record().unwrap().into_snapshot();
    let existing = source.form_controls(0).unwrap()[0]
        .to_record_bytes()
        .unwrap();
    assert_eq!(source.form_controls(0).unwrap()[0].object_id(), first_id);

    let mut transaction: Edit = source.edit();
    let id = transaction
        .add_form_control(0, checkbox_control(7, 0xB2), anchor())
        .unwrap();
    assert_ne!(id, first_id);
    let commit = transaction.commit_record().unwrap();

    assert!(commit.diagnostics().changed());
    let controls = commit.snapshot().form_controls(0).unwrap();
    assert_eq!(controls.len(), 2);
    assert_eq!(controls[1].object_id(), id);
    assert!(controls[1].subrecords.iter().any(|value| {
        matches!(value, ObjSubrecord::Unknown { kind: 0x7777, data } if data == &[0xB2])
    }));
    let output = commit.snapshot().finish();
    // Each control's Obj follows its own host-control shape, and the first
    // control keeps its bytes.
    assert_eq!(assert_objects_follow_their_shapes(&output), 2);
    let workbook = workbook_of(&output);
    assert!(
        workbook
            .windows(existing.len())
            .any(|window| window == existing)
    );
    let applied = commit.patch().apply(&source).unwrap();
    assert_eq!(applied.finish(), output);
    assert!(commit.patch().apply(&applied).is_err());
    assert_eq!(
        commit.patch().inverse().apply(&applied).unwrap().finish(),
        source.finish()
    );
}

#[test]
fn empty_object_metadata_edit_preserves_exact_bytes() {
    let source = ole_object(7, 0x2A, 0xB1);
    let input = object_workbook_cfb(&source, b"raw embedded payload");
    let snapshot = Snapshot::open(input.clone(), Limits::default()).unwrap();
    let mut transaction = snapshot.edit();

    transaction
        .update_object_metadata(0, 7, ObjectMetadataChange::new())
        .unwrap();
    let commit = transaction.commit_record().unwrap();

    assert!(!commit.diagnostics().changed());
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
            ObjectMetadataChange::new()
                .with_object_id(9)
                .with_common_flags(0xA55A)
                .with_picture_flags(FtPioGrbit {
                    raw: 0x0208 | 0x0001,
                }),
        )
        .unwrap();
    let commit = transaction.commit_record().unwrap();
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

    let mut ole = OleFile::open(
        Cursor::new(commit.snapshot().finish()),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
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
    let before = transaction
        .clone()
        .commit_record()
        .map(RecordCommit::into_snapshot)
        .unwrap()
        .finish();

    assert!(
        transaction
            .update_object_metadata(0, 8, ObjectMetadataChange::new().with_common_flags(0xAAAA),)
            .is_err()
    );
    assert!(
        transaction
            .update_object_metadata(
                0,
                7,
                ObjectMetadataChange::new().with_picture_flags(FtPioGrbit { raw: 0x0012 }),
            )
            .is_err()
    );
    assert_eq!(
        transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish(),
        before
    );
}

#[test]
fn metadata_edit_rejects_storage_class_transition_atomically() {
    let source = ole_object(7, 0x2A, 0xB1);
    let input = object_workbook_cfb(&source, b"raw embedded payload");
    let snapshot = Snapshot::open(input, Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction
        .clone()
        .commit_record()
        .map(RecordCommit::into_snapshot)
        .unwrap()
        .finish();

    assert!(
        transaction
            .update_object_metadata(
                0,
                7,
                ObjectMetadataChange::new().with_picture_flags(FtPioGrbit { raw: 0x020A }),
            )
            .is_err()
    );
    assert!(
        transaction
            .update_object_metadata(
                0,
                7,
                ObjectMetadataChange::new().with_picture_flags(FtPioGrbit { raw: 0x0288 }),
            )
            .is_err()
    );
    assert!(
        transaction
            .update_object_metadata(
                0,
                7,
                ObjectMetadataChange::new().with_picture_flags(FtPioGrbit { raw: 0x0228 }),
            )
            .is_err()
    );
    assert_eq!(
        transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish(),
        before
    );
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
        .update_object_metadata(0, 7, ObjectMetadataChange::new().with_common_flags(0xA55A))
        .unwrap();
    let commit = transaction.commit_record().unwrap();

    assert!(commit.patch().apply(&stale).is_err());
    assert_eq!(stale.finish(), stale_before);
}

#[test]
fn add_embedded_payload_binds_obj_and_mbd_and_preserves_opaque_streams() {
    let input = drawable_workbook_cfb(&[("UnrelatedRootStream", b"retain me")]);
    let payload_bytes = embedded_payload(b"payload-v1", b"unknown-v1");
    let payload = EmbeddedPayload::new(payload_bytes, Limits::default())
        .expect("payload CFB should validate");
    let snapshot = Snapshot::open(input.clone(), Limits::default()).expect("workbook should open");
    let mut transaction = snapshot.edit();

    let id = transaction
        .add_embedded_payload(0, draft(payload))
        .expect("typed embedded payload should publish");
    let commit = transaction.commit_record().expect("commit should validate");
    assert!(commit.diagnostics().changed());
    let object = &commit.snapshot().objects(0).unwrap()[0];
    assert_eq!(object.object_id(), id);
    let storage = object.storage_name().expect("the object names its storage");
    assert!(storage.starts_with("MBD"));
    assert!(matches!(
        object.subrecords.iter().find_map(|value| match value {
            ObjSubrecord::PictureFormula(value) => Some(value),
            _ => None,
        }),
        Some(FtPictFmla {
            storage_position: Some(_),
            control_buffer_size: None,
            ..
        })
    ));
    let formula = object
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
    // The Obj is written right after its picture-frame shape.
    assert_eq!(
        assert_objects_follow_their_shapes(&commit.snapshot().finish()),
        1
    );

    let mut ole = OleFile::open(
        Cursor::new(commit.snapshot().finish()),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("CFB should open");
    assert_eq!(
        ole.open_stream(&["UnrelatedRootStream"]).unwrap(),
        b"retain me"
    );
    assert_eq!(
        ole.open_stream(&[storage.as_str(), "OpaqueStorage", "Payload"])
            .unwrap(),
        b"payload-v1"
    );
    assert_eq!(
        ole.open_stream(&[storage.as_str(), "UnknownStream"])
            .unwrap(),
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
fn add_embedded_payload_names_the_prog_id_and_refuses_a_bad_preview_atomically() {
    let snapshot = Snapshot::open(drawable_workbook_cfb(&[]), Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction
        .clone()
        .commit_record()
        .map(RecordCommit::into_snapshot)
        .unwrap()
        .finish();
    let payload = || {
        EmbeddedPayload::new(embedded_payload(b"payload", b"unknown"), Limits::default()).unwrap()
    };
    let not_a_picture = EmbeddedObjectDraft::new(payload(), anchor(), b"GIF89a".to_vec());
    assert!(transaction.add_embedded_payload(0, not_a_picture).is_err());
    let non_ascii = draft(payload()).with_prog_id("Wörd.Document");
    assert!(transaction.add_embedded_payload(0, non_ascii).is_err());
    assert_eq!(
        transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish(),
        before
    );

    transaction
        .add_embedded_payload(0, draft(payload()).with_prog_id("Package"))
        .unwrap();
    let commit = transaction.commit_record().unwrap();
    let object = &commit.snapshot().objects(0).unwrap()[0];
    let formula = object
        .subrecords
        .iter()
        .find_map(|value| match value {
            ObjSubrecord::PictureFormula(value) => Some(value),
            _ => None,
        })
        .unwrap();
    assert_eq!(&formula.formula[11..15], &[0x03, 7, 0x00, 0x00]);
    assert_eq!(&formula.formula[15..22], b"Package");
}

#[test]
fn replacing_with_the_original_embedded_payload_is_an_exact_noop() {
    let input = drawable_workbook_cfb(&[]);
    let payload_bytes = embedded_payload(b"payload-v1", b"unknown-v1");
    let payload = EmbeddedPayload::new(payload_bytes.clone(), Limits::default())
        .expect("payload CFB should validate");
    let snapshot = Snapshot::open(input, Limits::default()).expect("workbook should open");
    let mut add = snapshot.edit();
    let id = add
        .add_embedded_payload(0, draft(payload))
        .expect("typed embedded payload should publish");
    let added_commit = add.commit_record().expect("add should commit");
    let added = added_commit.snapshot();

    let mut replace = added.edit();
    replace
        .replace_embedded_payload(
            0,
            id,
            EmbeddedPayload::new(payload_bytes, Limits::default())
                .expect("payload CFB should validate"),
        )
        .expect("same payload should be accepted");
    let commit = replace
        .commit_record()
        .expect("no-op replacement should commit");
    assert!(!commit.diagnostics().changed());
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
            EmbeddedPayload::new(replacement, Limits::default())
                .expect("replacement CFB should validate"),
        )
        .expect("replacement should publish");
    let commit = transaction.commit_record().expect("commit should validate");
    let object = &commit.snapshot().objects(0).unwrap()[0];
    assert_eq!(object.object_id(), 7);
    assert_eq!(object.storage_name().as_deref(), Some("MBD0000002A"));
    assert!(object.subrecords.iter().any(|value| {
        matches!(value, ObjSubrecord::Unknown { kind: 0x7777, data } if data == &[0xB1, 0xA5])
    }));

    let mut ole = OleFile::open(
        Cursor::new(commit.snapshot().finish()),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("CFB should open");
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
    let (input, ids) = with_embedded_objects(
        drawable_workbook_cfb(&[]),
        &[&embedded_payload(b"source payload", b"unknown")],
    );
    let snapshot = Snapshot::open(input.clone(), Limits::default()).expect("workbook should open");
    let storage = snapshot.objects(0).unwrap()[0].storage_name().unwrap();
    let mut transaction = snapshot.edit();

    transaction
        .remove_embedded_payload(0, ids[0])
        .expect("embedded object should be removable");
    let commit = transaction.commit_record().expect("commit should validate");
    assert!(commit.snapshot().objects(0).unwrap().is_empty());
    // The shape went with its Obj: the worksheet has no drawing left.
    let records = worksheet_records(&workbook_of(&commit.snapshot().finish()));
    assert!(!records.contains(&0x005D));
    assert!(!records.contains(&0x00EC));
    let mut ole = OleFile::open(
        Cursor::new(commit.snapshot().finish()),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("CFB should open");
    assert!(!ole.exists(&[storage.as_str()]));
    assert!(ole.open_stream(&[storage.as_str()]).is_err());
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
fn removing_one_object_keeps_every_other_obj_paired_with_its_shape() {
    let (input, ids) = with_embedded_objects(
        drawable_workbook_cfb(&[]),
        &[
            &embedded_payload(b"first", b"unknown"),
            &embedded_payload(b"second", b"unknown"),
            &embedded_payload(b"third", b"unknown"),
        ],
    );
    assert_eq!(assert_objects_follow_their_shapes(&input), 3);
    let snapshot = Snapshot::open(input, Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    let removed = transaction.remove_object(0, ids[1]).unwrap();
    assert_eq!(removed.object_id(), ids[1]);
    let after = transaction.commit_record().unwrap().into_snapshot();
    assert_eq!(assert_objects_follow_their_shapes(&after.finish()), 2);
    assert_eq!(
        after
            .objects(0)
            .unwrap()
            .iter()
            .map(OleObjectRecord::object_id)
            .collect::<Vec<_>>(),
        vec![ids[0], ids[2]]
    );
    // Each remaining object keeps its own storage, and its shape keeps its
    // own anchor and preview.
    let package = crate::Package::from_bytes(&after.finish()).unwrap();
    let mut features = package.edit_sheet_features().unwrap();
    features.remove_embedded_object(0, ids[2]).unwrap();
    features.remove_embedded_object(0, ids[0]).unwrap();
    assert!(features.remove_embedded_object(0, ids[1]).is_err());
}

#[test]
fn removing_one_of_two_shared_mbd_references_keeps_storage_until_last_owner() {
    // Two objects, then the second Obj is pointed at the first storage and
    // the second storage dropped: two owners of one storage.
    let (input, ids) = with_embedded_objects(
        drawable_workbook_cfb(&[]),
        &[
            &embedded_payload(b"shared payload", b"unknown"),
            &embedded_payload(b"other payload", b"unknown"),
        ],
    );
    let snapshot = Snapshot::open(input.clone(), Limits::default()).unwrap();
    let objects = snapshot.objects(0).unwrap();
    let first = objects[0].storage_position().unwrap();
    let second = objects[1].storage_position().unwrap();
    let second_name = objects[1].storage_name().unwrap();
    let mut streams = cfb_streams(&input);
    streams.retain(|(path, _)| path[0] != second_name);
    for (path, data) in &mut streams {
        if path == &["Workbook"] {
            let needle = second.to_le_bytes();
            let at = data
                .windows(4)
                .rposition(|window| window == needle)
                .unwrap();
            data[at..at + 4].copy_from_slice(&first.to_le_bytes());
        }
    }
    let input = cfb_of(&streams);
    let shared = format!("MBD{first:08X}");

    let snapshot = Snapshot::open(input, Limits::default()).expect("workbook should open");
    let mut first_edit = snapshot.edit();
    first_edit
        .remove_embedded_payload(0, ids[0])
        .expect("first owner should be removable");
    let first_commit = first_edit
        .commit_record()
        .expect("first removal should commit");
    assert_eq!(first_commit.snapshot().objects(0).unwrap().len(), 1);
    let mut first_ole = OleFile::open(
        Cursor::new(first_commit.snapshot().finish()),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("first result should be a CFB");
    assert_eq!(
        first_ole
            .open_stream(&[shared.as_str(), "OpaqueStorage", "Payload"])
            .unwrap(),
        b"shared payload"
    );

    let mut second_edit = first_commit.snapshot().edit();
    second_edit
        .remove_embedded_payload(0, ids[1])
        .expect("last owner should be removable");
    let second_commit = second_edit
        .commit_record()
        .expect("second removal should commit");
    assert!(second_commit.snapshot().objects(0).unwrap().is_empty());
    let second_ole = OleFile::open(
        Cursor::new(second_commit.snapshot().finish()),
        litchi_cfb::OleFileLimits::default(),
    )
    .expect("second result should be a CFB");
    assert!(!second_ole.exists(&[shared.as_str()]));
}

#[test]
fn embedded_payload_api_rejects_dde_and_invalid_payloads_atomically() {
    assert!(EmbeddedPayload::new(vec![1, 2, 3], Limits::default()).is_err());

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
    let before = transaction
        .clone()
        .commit_record()
        .map(RecordCommit::into_snapshot)
        .unwrap()
        .finish();
    assert!(transaction.remove_embedded_payload(0, 7).is_err());
    assert!(
        transaction
            .replace_embedded_payload(
                0,
                7,
                EmbeddedPayload::new(embedded_payload(b"new", b"unknown"), Limits::default())
                    .unwrap(),
            )
            .is_err()
    );
    assert_eq!(
        transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish(),
        before
    );
}

#[test]
fn add_embedded_payload_allocates_fresh_object_and_storage_identities() {
    let (input, ids) = with_embedded_objects(
        drawable_workbook_cfb(&[]),
        &[&embedded_payload(b"first", b"unknown")],
    );
    let snapshot = Snapshot::open(input, Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    let id = transaction
        .add_embedded_payload(
            0,
            draft(
                EmbeddedPayload::new(embedded_payload(b"new", b"unknown"), Limits::default())
                    .unwrap(),
            ),
        )
        .unwrap();
    assert_ne!(id, ids[0]);
    let objects = transaction.objects(0).unwrap();
    assert_eq!(objects.len(), 2);
    assert_ne!(objects[0].storage_name(), objects[1].storage_name());
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

    let payload =
        EmbeddedPayload::new(embedded_payload(b"new", b"unknown"), Limits::default()).unwrap();
    let draft = draft(payload);
    let mut transaction = snapshot.edit();
    let before = transaction
        .clone()
        .commit_record()
        .map(RecordCommit::into_snapshot)
        .unwrap()
        .finish();
    assert!(transaction.add_embedded_payload(0, draft).is_err());
    assert_eq!(
        transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish(),
        before
    );
    assert!(transaction.remove_embedded_payload(0, 7).is_err());
    assert_eq!(
        transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish(),
        before
    );
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
    let before = transaction
        .clone()
        .commit_record()
        .map(RecordCommit::into_snapshot)
        .unwrap()
        .finish();
    let draft = draft(
        EmbeddedPayload::new(embedded_payload(b"new", b"opaque"), Limits::default()).unwrap(),
    );
    assert!(transaction.add_embedded_payload(0, draft).is_err());
    assert!(transaction.remove_embedded_payload(0, 7).is_err());
    assert_eq!(
        transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish(),
        before
    );
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
    let before = transaction
        .clone()
        .commit_record()
        .map(RecordCommit::into_snapshot)
        .unwrap()
        .finish();
    let payload =
        EmbeddedPayload::new(embedded_payload(b"new", b"unknown"), Limits::default()).unwrap();
    let draft = draft(payload);
    let error = transaction
        .add_embedded_payload(0, draft)
        .expect_err("a second target must be refused before capture");
    assert!(matches!(
        error,
        crate::error::Error::Limit {
            kind: crate::LimitKind::OleObject(super::super::LimitKind::Objects),
            observed: 2,
            maximum: 1
        }
    ));
    assert_eq!(
        transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish(),
        before
    );
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

    let payload =
        EmbeddedPayload::new(protected_payload(), Limits::default()).expect("payload CFB parses");
    let draft = draft(payload);
    let snapshot = Snapshot::open(workbook_cfb(&[]), Limits::default()).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction
        .clone()
        .commit_record()
        .map(RecordCommit::into_snapshot)
        .unwrap()
        .finish();
    assert!(transaction.add_embedded_payload(0, draft).is_err());
    assert_eq!(
        transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish(),
        before
    );
}

#[test]
fn embedded_payload_api_refuses_nested_pidssi_and_encryption() {
    let snapshot = Snapshot::open(workbook_cfb(&[]), Limits::default()).unwrap();
    for bytes in [
        nested_pidssi_with_signature(),
        nested_child_pidssi_with_signature(),
        payload_with_root_stream("encryption", b"opaque encrypted payload"),
    ] {
        let payload =
            EmbeddedPayload::new(bytes, Limits::default()).expect("nested CFB should parse");
        let draft = draft(payload);
        let mut transaction = snapshot.edit();
        let before = transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish();
        let error = transaction
            .add_embedded_payload(0, draft)
            .expect_err("nested protection must refuse publication");
        assert!(matches!(
            error,
            crate::error::Error::Refused(_)
                | crate::error::Error::PatchConflict { .. }
                | crate::error::Error::Internal(_)
        ));
        assert_eq!(
            transaction
                .clone()
                .commit_record()
                .map(RecordCommit::into_snapshot)
                .unwrap()
                .finish(),
            before
        );
    }
}

#[test]
fn embedded_payload_checks_stream_bounds_before_parsing_nested_pidssi() {
    let limits = Limits {
        max_stream_size: 1024,
        ..publication_limits()
    };
    let bytes = payload_with_root_stream("\u{0005}DocumentSummaryInformation", &vec![0xA5; 2_048]);
    let payload = EmbeddedPayload::new(bytes, limits).expect("header should be admitted");
    let draft = draft(payload);
    let snapshot = Snapshot::open(workbook_cfb(&[]), limits).unwrap();
    let mut transaction = snapshot.edit();
    let error = transaction
        .add_embedded_payload(0, draft)
        .expect_err("oversized PIDDSI must be refused before parsing");
    assert!(matches!(
        error,
        crate::error::Error::Limit {
            kind: crate::LimitKind::OleObject(super::super::LimitKind::StreamBytes),
            observed: 2048,
            maximum: 1024
        }
    ));
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
        let payload = EmbeddedPayload::new(bytes, limits)
            .unwrap_or_else(|error| panic!("{name} should pass standalone admission: {error:?}"));
        let draft = draft(payload);
        let snapshot = Snapshot::open(workbook_cfb(&[]), limits)
            .unwrap_or_else(|error| panic!("{name} workbook should open: {error:?}"));
        let mut transaction = snapshot.edit();
        let before = transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish();
        assert!(
            transaction.add_embedded_payload(0, draft).is_err(),
            "{name} should be enforced by workbook publication"
        );
        assert_eq!(
            transaction
                .clone()
                .commit_record()
                .map(RecordCommit::into_snapshot)
                .unwrap()
                .finish(),
            before
        );
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
    let payload = EmbeddedPayload::new(wide_payload(16), limits)
        .expect("standalone payload admission should only parse the CFB header");
    let draft = draft(payload);
    let snapshot = Snapshot::open(workbook_cfb(&[]), limits).unwrap();
    let mut transaction = snapshot.edit();
    let before = transaction
        .clone()
        .commit_record()
        .map(RecordCommit::into_snapshot)
        .unwrap()
        .finish();
    let error = transaction
        .add_embedded_payload(0, draft)
        .expect_err("wide sibling directories must be bounded before capture");
    assert!(
        matches!(error, crate::error::Error::Limit { kind: crate::LimitKind::OleObject(super::super::LimitKind::Storages), observed, maximum } if observed == maximum + 1)
    );
    assert_eq!(
        transaction
            .clone()
            .commit_record()
            .map(RecordCommit::into_snapshot)
            .unwrap()
            .finish(),
        before
    );
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
    let payload = EmbeddedPayload::new(wide_payload(2), limits)
        .expect("standalone payload header should be admitted");
    let draft = draft(payload);
    let snapshot = Snapshot::open(workbook_cfb(&[]), limits).unwrap();
    let mut transaction = snapshot.edit();
    let error = transaction
        .add_embedded_payload(0, draft)
        .expect_err("one selected payload must use the per-object storage ceiling");
    assert!(
        matches!(error, crate::error::Error::Limit { kind: crate::LimitKind::OleObject(super::super::LimitKind::Storages), observed, maximum } if observed == maximum + 1)
    );
}

#[test]
fn standalone_and_publication_byte_limits_keep_numeric_evidence() {
    let bytes = payload_with_root_stream("private-payload-name", &[0xa5; 4]);
    let maximum = bytes.len() as u64 - 1;
    let error = EmbeddedPayload::new(
        bytes.clone(),
        Limits {
            max_object_size: maximum,
            ..publication_limits()
        },
    )
    .expect_err("object byte bound is checked before retaining the payload");
    assert!(matches!(error, crate::Error::Limit {
        kind: crate::LimitKind::OleObject(super::super::LimitKind::ObjectBytes),
        observed, maximum: bound,
    } if observed == maximum + 1 && bound == maximum));

    let limits = Limits {
        max_total_size: 3,
        ..publication_limits()
    };
    let payload = EmbeddedPayload::new(bytes, limits).expect("standalone admission is unchanged");
    let error = payload
        .validate_for_publication(limits)
        .expect_err("aggregate stream bytes are checked before capture");
    assert!(matches!(
        error,
        crate::Error::Limit {
            kind: crate::LimitKind::OleObject(super::super::LimitKind::AggregateBytes),
            observed: 4,
            maximum: 3,
        }
    ));
    assert!(!error.to_string().contains("private-payload-name"));
}

#[test]
fn invalid_publication_limits_are_distinct_from_exhausted_budgets() {
    let bytes = payload_with_root_stream("Opaque", &[0xa5; 4]);
    let payload = EmbeddedPayload::new(bytes, Limits::default()).unwrap();
    for (limits, expected) in [
        (
            Limits {
                max_streams: 0,
                ..publication_limits()
            },
            super::super::LimitKind::Streams,
        ),
        (
            Limits {
                max_total_size: 0,
                ..publication_limits()
            },
            super::super::LimitKind::AggregateBytes,
        ),
    ] {
        let error = payload
            .validate_for_publication(limits)
            .expect_err("zero publication limits are invalid configurations");
        assert!(matches!(error, crate::Error::InvalidLimit {
            kind: crate::LimitKind::OleObject(kind), value: 0, ..
        } if kind == expected));
    }
}
