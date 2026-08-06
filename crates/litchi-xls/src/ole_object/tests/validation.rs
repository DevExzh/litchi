//! Focused semantic validation coverage.

use super::super::*;

#[test]
fn ole_obj_rejects_conflicting_link_and_control_flags() {
    let value = OleObjectRecord {
        subrecords: vec![
            ObjSubrecord::Common(FtCmo {
                object_type: 8,
                object_id: 1,
                flags: 0,
                reserved: [0; 12],
            }),
            ObjSubrecord::PictureFlags(FtPioGrbit { raw: 0x0012 }),
            ObjSubrecord::PictureFormula(FtPictFmla {
                formula: vec![0],
                storage_position: Some(1),
                control_buffer_size: Some(0),
            }),
            ObjSubrecord::End,
        ],
        text_object: None,
    };
    assert!(value.validate().is_err());
}

#[test]
fn scrollbar_validation_rejects_out_of_range_values() {
    let value = FtSbs {
        reserved: [0; 4],
        value: 11,
        minimum: 0,
        maximum: 10,
        increment: 1,
        page_increment: 1,
        horizontal: false,
        scroll_width: 1,
        flags: 0,
    };
    assert!(value.validate().is_err());
}
