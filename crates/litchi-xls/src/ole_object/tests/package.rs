//! Regression tests for the XLS OLE-object owner.

use super::super::package::{read_workbook, targets_for_sheets};
use super::super::*;
use crate::error::Error;
use litchi_cfb::{OleFile, OleWriter};
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
            ObjSubrecord::PictureFlags(FtPioGrbit {
                raw: if dde { 0x0002 } else { 0 },
            }),
            ObjSubrecord::PictureFormula(FtPictFmla {
                formula: vec![0x05, 0, 0, 0, 0],
                storage_position: Some(position),
                control_buffer_size: Some(0),
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
fn bounds_workbook_read_before_stream_materialization() {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], &[0; 128])
        .expect("test stream should be created");
    let mut output = Cursor::new(Vec::new());
    writer
        .write_to(&mut output)
        .expect("test compound file should be written");

    let mut limits = Limits::default();
    limits.max_stream_size = 64;
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
