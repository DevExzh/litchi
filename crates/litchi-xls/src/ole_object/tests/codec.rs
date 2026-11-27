//! Focused wire-level round-trip coverage.

use super::super::*;

fn subrecord(kind: u16, body: &[u8]) -> Vec<u8> {
    let mut output = Vec::with_capacity(body.len() + 4);
    output.extend_from_slice(&kind.to_le_bytes());
    output.extend_from_slice(&(body.len() as u16).to_le_bytes());
    output.extend_from_slice(body);
    output
}

fn terminal_list(header: u16, body: &[u8]) -> Vec<u8> {
    let mut common = [0; 18];
    common[..2].copy_from_slice(&ObjectType::List.code().to_le_bytes());
    common[2..4].copy_from_slice(&7u16.to_le_bytes());
    let mut data = subrecord(FT_CMO, &common);
    data.extend_from_slice(&FT_LBS_DATA.to_le_bytes());
    data.extend_from_slice(&header.to_le_bytes());
    data.extend_from_slice(body);
    data
}

#[test]
fn lbs_continuation_header_can_exceed_the_terminal_payload_length() {
    let data = terminal_list(0x1fde, &[0, 0, 3, 0, 2, 0, 0, 0, 0, 0]);
    let control = FormControl::parse(&data, None).unwrap();
    let list = control.list_box_data().unwrap();
    assert_eq!((list.entry_count, list.selected_index), (3, 2));
    assert_eq!(list.cb_f_continued(), Some(0x1fde));
    assert!(matches!(
        control.subrecords.last(),
        Some(ObjSubrecord::ListBoxData(_))
    ));
    assert_eq!(&control.to_record_bytes().unwrap()[4..], data);
}

#[test]
fn lbs_partial_framing_and_unmodeled_payloads_keep_their_original_headers() {
    let mut body = vec![0, 0, 2, 0, 1, 0, 2, 0, 0, 0];
    body.extend_from_slice(&[1, 0, 0, b'A']);
    let header = u16::try_from(body.len() - 1).unwrap();
    let data = terminal_list(header, &body);
    let control = FormControl::parse(&data, None).unwrap();
    assert!(control.list_box_data().unwrap().partial);
    assert_eq!(
        control.list_box_data().unwrap().cb_f_continued(),
        Some(header)
    );
    assert_eq!(&control.to_record_bytes().unwrap()[4..], data);
    let raw = terminal_list(0x1fde, &[0xff]);
    let control = FormControl::parse(&raw, None).unwrap();
    assert!(control.list_box_data().is_none());
    assert!(
        matches!(control.subrecords.last(),Some(ObjSubrecord::RawListBoxData { cb_f_continued:0x1fde,data }) if data==&[0xff])
    );
    assert_eq!(&control.to_record_bytes().unwrap()[4..], raw);
    assert!(control.validate().is_err());
}

#[test]
fn empty_lbs_payload_and_present_zero_fields_remain_distinct() {
    for data in [terminal_list(0, &[]), terminal_list(10, &[0; 10])] {
        let control = FormControl::parse(&data, None).unwrap();
        assert_eq!(&control.to_record_bytes().unwrap()[4..], data);
    }
}

#[test]
fn legacy_length_framed_lbs_with_explicit_end_stays_readable() {
    let mut data = terminal_list(10, &[0, 0, 3, 0, 2, 0, 0, 0, 0, 0]);
    data.extend(subrecord(FT_END, &[]));
    let control = FormControl::parse(&data, None).unwrap();
    assert_eq!(control.list_box_data().unwrap().selected_index, 2);
    assert!(matches!(control.subrecords.last(), Some(ObjSubrecord::End)));
    assert_eq!(&control.to_record_bytes().unwrap()[4..], data);
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
