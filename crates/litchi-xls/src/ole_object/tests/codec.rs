//! Focused wire-level round-trip coverage.

use super::super::*;

fn subrecord(kind: u16, body: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(body.len() + 4);
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(&(body.len() as u16).to_le_bytes());
    output.extend_from_slice(body);
    output
}

#[test]
fn ole_obj_round_trip_retains_unknown_and_reserved_bytes() {
    let value = OleObjectRecord {
        subrecords: vec![
            ObjSubrecord::Common(FtCmo {
                object_type: 8,
                object_id: 7,
                flags: 0xA5A5,
                reserved: [0xCC; 12],
            }),
            ObjSubrecord::PictureFlags(FtPioGrbit { raw: 0x8202 }),
            ObjSubrecord::Unknown {
                kind: 0x7777,
                data: vec![0x10, 0x20, 0x30],
            },
            ObjSubrecord::PictureFormula(FtPictFmla {
                formula: vec![1, 2, 3],
                storage_position: Some(0x2A),
                control_buffer_size: Some(0),
            }),
            ObjSubrecord::End,
        ],
        text_object: Some(vec![0xB6, 0, 0, 0]),
    };
    let bytes = value.to_record_bytes().expect("valid Obj");
    let parsed = OleObjectRecord::parse(&bytes[4..], value.text_object.clone())
        .expect("serialized Obj should parse");
    assert_eq!(parsed, value);
    assert_eq!(parsed.to_record_bytes().expect("round trip"), bytes);
}

#[test]
fn malformed_control_payload_stays_inert_and_lossless() {
    let mut body = Vec::new();
    let mut cmo = Vec::new();
    cmo.extend_from_slice(&0x000Bu16.to_le_bytes());
    cmo.extend_from_slice(&7u16.to_le_bytes());
    cmo.extend_from_slice(&0u16.to_le_bytes());
    cmo.extend_from_slice(&[0xDD; 12]);
    body.extend_from_slice(&subrecord(0x0015, &cmo));
    body.extend_from_slice(&subrecord(0x0012, &[0xFE]));
    body.extend_from_slice(&subrecord(0, &[]));

    let control = FormControl::parse(&body, None).expect("checkbox Obj");
    assert!(matches!(
        control.subrecords[1],
        ObjSubrecord::Unknown { kind: 0x0012, ref data } if data == &[0xFE]
    ));
    let mut expected = 0x005Du16.to_le_bytes().to_vec();
    expected.extend_from_slice(&(body.len() as u16).to_le_bytes());
    expected.extend_from_slice(&body);
    assert_eq!(
        control.to_record_bytes().expect("control round trip"),
        expected
    );
}
