//! Regression tests for the XLS OLE-object owner.

use super::super::package::{read_workbook, targets_for_sheets};
use super::super::*;
use crate::error::Error;
use litchi_cfb::{OleFile, OleWriter};
use litchi_ole_common::property_set::document_summary::DIGITAL_SIGNATURE;
use litchi_ole_common::property_set::{
    CodePage, DOCUMENT_SUMMARY_INFORMATION_FMTID, Section, Stream, Value,
};
use std::io::Cursor;

fn object(id: u16, position: u32, dde: bool) -> OleObjectRecord {
    OleObjectRecord {
        subrecords: vec![
            ObjSubrecord::Common(FtCmo {
                object_type: 8,
                object_id: id,
                flags: 0,
                reserved: [0; 12],
            }),
            ObjSubrecord::PictureFormat(FtCf {
                format: FtCf::UNSPECIFIED,
            }),
            ObjSubrecord::PictureFlags(FtPioGrbit {
                raw: if dde { 0x0002 } else { 0 },
            }),
            ObjSubrecord::PictureFormula(FtPictFmla {
                formula: vec![
                    0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00,
                    0x00,
                ],
                storage_position: Some(position),
                control_buffer_size: if dde { Some(0) } else { None },
            }),
            ObjSubrecord::End,
        ],
        text_object: None,
    }
}

#[test]
fn derives_deduplicated_mbd_and_lnk_targets_from_obj_records() {
    let targets = targets_for_sheets(&[vec![
        object(1, 0x2A, false),
        object(2, 0x2A, false),
        object(3, 0x2A, true),
    ]])
    .expect("BIFF references should produce valid targets");

    assert_eq!(targets.len(), 2);
    let mbd = targets.get("MBD0000002A").expect("MBD target");
    assert_eq!(mbd.path(), &["MBD0000002A".to_owned()]);
    let lnk = targets.get("LNK0000002A").expect("LNK target");
    assert_eq!(lnk.path(), &["LNK0000002A".to_owned()]);
}

#[test]
fn storage_name_does_not_claim_camera_or_control_stream_ownership() {
    let mut camera = object(1, 0x2A, false);
    if let Some(ObjSubrecord::PictureFlags(flags)) = camera
        .subrecords
        .iter_mut()
        .find(|value| matches!(value, ObjSubrecord::PictureFlags(_)))
    {
        flags.raw |= 0x0080;
    }
    assert_eq!(camera.storage_name(), None);

    let mut controls_stream = object(2, 0x2A, false);
    if let Some(ObjSubrecord::PictureFlags(flags)) = controls_stream
        .subrecords
        .iter_mut()
        .find(|value| matches!(value, ObjSubrecord::PictureFlags(_)))
    {
        flags.raw |= 0x0020;
    }
    assert_eq!(controls_stream.storage_name(), None);
}

#[test]
fn bounds_workbook_read_before_stream_materialization() {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &[0; 128])
        .expect("test stream should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("test compound file should be written");

    let limits = Limits {
        max_stream_size: 64,
        ..Default::default()
    };
    let error = read_workbook(&output.into_inner(), limits)
        .expect_err("oversized Workbook must be rejected before reading");
    assert!(matches!(error, Error::InvalidData(message) if message.contains("read limit")));
}

fn record(kind: u16, body: &[u8]) -> Vec<u8> {
    let mut output = kind.to_le_bytes().to_vec();
    output.extend_from_slice(&(body.len() as u16).to_le_bytes());
    output.extend_from_slice(body);
    output
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

fn workbook_cfb_with_root_stream(name: &str, data: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &workbook_stream(&[]))
        .expect("Workbook stream should be created");
    writer
        .create_stream(&[name], data)
        .expect("root stream should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("test compound file should be written");
    output.into_inner()
}

fn pidssi_with_signature() -> Vec<u8> {
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
        .expect("signature property should be accepted by the generic set");
    Stream::new(section)
        .to_bytes()
        .expect("PIDDSI stream should serialize")
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

fn occurrences(haystack: &[u8], needle: &[u8]) -> usize {
    haystack
        .windows(needle.len())
        .filter(|window| *window == needle)
        .count()
}

#[test]
fn add_form_control_is_typed_atomic_and_lossless() {
    let existing = checkbox_control(7, 0xA1).to_record_bytes().unwrap();
    let mut editor = Editor::new(
        workbook_cfb(std::slice::from_ref(&existing)),
        Limits::default(),
    )
    .expect("workbook should open");
    let authored = checkbox_control(8, 0xB2);
    editor
        .add_form_control(0, authored.clone())
        .expect("typed control should be authored");
    assert_eq!(editor.form_controls(0).unwrap().len(), 2);

    let bytes = editor.finish().expect("transaction should finish");
    let mut ole = OleFile::open(Cursor::new(bytes.clone())).expect("CFB should reopen");
    let workbook = ole
        .open_stream(&["Workbook"])
        .expect("Workbook stream should remain present");
    let authored_bytes = authored.to_record_bytes().unwrap();
    assert_eq!(occurrences(&workbook, &existing), 1);
    assert_eq!(occurrences(&workbook, &authored_bytes), 1);

    let mut editor = Editor::new(bytes, Limits::default()).expect("edited workbook should open");
    let second_authored = checkbox_control(9, 0xC3);
    editor
        .add_form_control(0, second_authored.clone())
        .expect("second typed control should be authored");
    let bytes = editor.finish().expect("second transaction should finish");
    let mut ole = OleFile::open(Cursor::new(bytes)).expect("second CFB should reopen");
    let workbook = ole
        .open_stream(&["Workbook"])
        .expect("second Workbook stream should remain present");
    assert_eq!(occurrences(&workbook, &existing), 1);
    assert_eq!(occurrences(&workbook, &authored_bytes), 1);
    assert_eq!(
        occurrences(&workbook, &second_authored.to_record_bytes().unwrap()),
        1
    );
}

#[test]
fn add_form_control_rejects_duplicate_ids_without_mutation() {
    let existing = checkbox_control(7, 0xA1).to_record_bytes().unwrap();
    let mut editor = Editor::new(
        workbook_cfb(std::slice::from_ref(&existing)),
        Limits::default(),
    )
    .expect("workbook should open");
    let error = editor
        .add_form_control(0, checkbox_control(7, 0xB2))
        .expect_err("duplicate control IDs must be rejected");
    assert!(matches!(
        error,
        Error::InvalidRecord {
            record_type: OBJ,
            ..
        }
    ));
    assert_eq!(editor.form_controls(0).unwrap().len(), 1);
    assert_eq!(
        editor.form_controls(0).unwrap()[0]
            .to_record_bytes()
            .unwrap(),
        existing
    );
}

#[test]
fn workbook_reader_rejects_filepass_and_active_biff_protection_before_editing() {
    let cases = [
        ("FILEPASS", 0x002F, vec![0, 0]),
        ("unknown FILEPASS", 0x002F, vec![0xFF; 8]),
        ("malformed FILEPASS", 0x002F, vec![0]),
        ("workbook PROTECT", 0x0012, vec![1, 0]),
        ("worksheet OBJECTPROTECT", 0x0063, vec![1, 0]),
        ("invalid protection Boolean", 0x0012, vec![2, 0]),
        (
            "FILESHARING write reservation",
            0x005B,
            vec![0, 0, 1, 0, 0, 0, 0],
        ),
    ];
    for (name, kind, body) in cases {
        let error = read_workbook(&workbook_cfb(&[record(kind, &body)]), Limits::default())
            .expect_err(name);
        assert!(
            matches!(
                error,
                Error::PasswordProtected
                    | Error::UnsafeEdit(_)
                    | Error::InvalidRecord { .. }
                    | Error::InvalidLength { .. }
            ),
            "{name}: {error:?}"
        );
    }
}

#[test]
fn file_sharing_read_only_recommendation_is_not_protection() {
    let input = workbook_cfb(&[record(0x005B, &[1, 0, 0, 0, 0, 0])]);
    let editor = Editor::new(input.clone(), Limits::default())
        .expect("read-only recommendation without a write password is editable");
    assert_eq!(editor.finish().expect("no-op should finish"), input);
}

#[test]
fn workbook_reader_rejects_root_encryption_stream() {
    let error = read_workbook(
        &workbook_cfb_with_root_stream("encryption", b"opaque"),
        Limits::default(),
    )
    .expect_err("root encryption stream must refuse edits");
    assert!(matches!(error, Error::PasswordProtected));
}

#[test]
fn workbook_reader_refuses_pidssi_digital_signature() {
    let error = read_workbook(
        &workbook_cfb_with_root_stream(
            "\u{0005}DocumentSummaryInformation",
            &pidssi_with_signature(),
        ),
        Limits::default(),
    )
    .expect_err("a stale PIDDSI signature must not be rewritten");
    assert!(matches!(error, Error::UnsafeEdit(message) if message.contains("DigitalSignature")));
}

#[test]
fn disabled_protection_markers_remain_readable_and_exact_noop() {
    let input = workbook_cfb(&[
        record(0x0012, &[0, 0]),
        record(0x0063, &[0, 0]),
        record(0x0013, &[0, 0]),
    ]);
    let editor = Editor::new(input.clone(), Limits::default())
        .expect("disabled protection metadata should remain readable");
    let _ = editor
        .objects(0)
        .expect("worksheet object view should remain available");
    assert_eq!(editor.finish().expect("no-op should finish"), input);
}
