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
            ObjSubrecord::PictureFormat(FtCf {
                format: FtCf::UNSPECIFIED,
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
fn storage_obj_fmla_without_controls_size_remains_lossless() {
    let value = OleObjectRecord {
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
            ObjSubrecord::PictureFlags(FtPioGrbit { raw: 0 }),
            ObjSubrecord::PictureFormula(FtPictFmla {
                formula: vec![
                    0x05, 0x00, 0x00, 0x00, 0x00, 0x00, 0x02, 0x00, 0x00, 0x00, 0x00, 0x03, 0x00,
                    0x00,
                ],
                storage_position: Some(0x2A),
                control_buffer_size: None,
            }),
            ObjSubrecord::End,
        ],
        text_object: None,
    };
    let bytes = value.to_record_bytes().expect("storage Obj should encode");
    let parsed = OleObjectRecord::parse(&bytes[4..], None).expect("storage Obj should parse");
    assert_eq!(parsed, value);
}

#[test]
fn storage_obj_fmla_authoring_wire_has_even_length_ptgtbl_and_embed_info() {
    let formula = vec![
        0x05, 0x00, // ObjectParsedFormula.cce = 5
        0x00, 0x00, 0x00, 0x00, // ObjectParsedFormula unused
        0x02, 0x00, 0x00, 0x00, 0x00, // PtgTbl
        0x03, 0x00, 0x00, // PictFmlaEmbedInfo (no class name)
    ];
    assert_eq!(formula.len(), 14);
    let value = OleObjectRecord {
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
            ObjSubrecord::PictureFlags(FtPioGrbit { raw: 0 }),
            ObjSubrecord::PictureFormula(FtPictFmla {
                formula: formula.clone(),
                storage_position: Some(0x2A),
                control_buffer_size: None,
            }),
            ObjSubrecord::End,
        ],
        text_object: None,
    };

    let mut obj_body = Vec::new();
    obj_body.extend_from_slice(&subrecord(
        FT_CMO,
        &[
            0x08, 0x00, // cmo.ot
            0x07, 0x00, // cmo.id
            0x00, 0x00, // cmo.flags
            0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, // reserved
        ],
    ));
    obj_body.extend_from_slice(&subrecord(
        FT_CF,
        &[0xFF, 0xFF], // cf = unspecified; ft/cb are the Obj subrecord header
    ));
    obj_body.extend_from_slice(&subrecord(FT_PIO, &[0x00, 0x00]));
    let mut formula_body = (formula.len() as u16).to_le_bytes().to_vec();
    formula_body.extend_from_slice(&formula);
    formula_body.extend_from_slice(&0x2Au32.to_le_bytes());
    obj_body.extend_from_slice(&subrecord(FT_PICT_FMLA, &formula_body));
    obj_body.extend_from_slice(&subrecord(FT_END, &[]));

    let mut expected = 0x005Du16.to_le_bytes().to_vec();
    expected.extend_from_slice(&(obj_body.len() as u16).to_le_bytes());
    expected.extend_from_slice(&obj_body);
    assert_eq!(
        value.to_record_bytes().expect("Obj should encode"),
        expected
    );

    let parsed = OleObjectRecord::parse(&expected[4..], None).expect("wire should parse");
    assert_eq!(parsed, value);
}

#[test]
fn picture_format_rejects_unknown_clipboard_selector() {
    let value = OleObjectRecord {
        subrecords: vec![
            ObjSubrecord::Common(FtCmo {
                object_type: 8,
                object_id: 7,
                flags: 0,
                reserved: [0; 12],
            }),
            ObjSubrecord::PictureFormat(FtCf { format: 0x1234 }),
            ObjSubrecord::PictureFlags(FtPioGrbit { raw: 0 }),
            ObjSubrecord::PictureFormula(FtPictFmla {
                formula: vec![0x05, 0, 0, 0, 0],
                storage_position: Some(0x2A),
                control_buffer_size: None,
            }),
            ObjSubrecord::End,
        ],
        text_object: None,
    };
    assert!(value.to_record_bytes().is_err());
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
