//! Regression tests for the XLS OLE-object owner.

use super::super::package::{read_workbook, targets_for_sheets};
use super::super::*;
use crate::error::Error;
use litchi_cfb::OleWriter;
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
