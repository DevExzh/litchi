//! Focused publication tests for the OLE-object facade.

use super::super::{CheckState, FormControl, FtCblsData, FtCmo, Limits, ObjSubrecord};
use super::super::{Snapshot, Transaction};
use litchi_cfb::OleWriter;
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
