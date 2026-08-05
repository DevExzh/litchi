use litchi_cfb::{OleFile, OleWriter};
use litchi_xls::ole_object::Limits;
use litchi_xls::{
    CheckState, DropDownStyle, EditBoxValidation, Editor, FormControl, FtCmo, FtPictFmla,
    FtPioGrbit, LbsItem, ListBehaviorClass, ListSelectionType, ObjSubrecord, ObjectType,
    OleObjectRecord,
};
use std::io::Cursor;

const OBJ: u16 = 0x005D;
const FT_CMO: u16 = 0x0015;
const FT_SBS: u16 = 0x000C;
const FT_GBO_DATA: u16 = 0x000F;
const FT_EDO_DATA: u16 = 0x0010;
const FT_RBO_DATA: u16 = 0x0011;
const FT_CBLS_DATA: u16 = 0x0012;
const FT_LBS_DATA: u16 = 0x0013;
const FT_END: u16 = 0;

fn subrecord(kind: u16, body: &[u8]) -> Vec<u8> {
    let mut value = kind.to_le_bytes().to_vec();
    value.extend_from_slice(&(body.len() as u16).to_le_bytes());
    value.extend_from_slice(body);
    value
}

fn cmo(object_type: u16, object_id: u16) -> Vec<u8> {
    let mut body = object_type.to_le_bytes().to_vec();
    body.extend_from_slice(&object_id.to_le_bytes());
    body.extend_from_slice(&0u16.to_le_bytes());
    body.extend_from_slice(&[0; 12]);
    subrecord(FT_CMO, &body)
}

/// A complete OBJ record (with record header) from subrecord bodies.
fn obj_record(parts: &[Vec<u8>]) -> Vec<u8> {
    let body = parts.concat();
    subrecord(OBJ, &body)
}

/// A compressed XLUnicodeString fixture.
fn xl_string(text: &str) -> Vec<u8> {
    let mut value = (text.len() as u16).to_le_bytes().to_vec();
    value.push(0);
    value.extend_from_slice(text.as_bytes());
    value
}

fn form_control(record: &[u8]) -> FormControl {
    let body_len = u16::from_le_bytes([record[2], record[3]]) as usize;
    FormControl::parse(&record[4..4 + body_len], None).expect("form control parses")
}

fn round_trip(record: &[u8]) {
    let control = form_control(record);
    assert_eq!(control.to_record_bytes().unwrap(), record);
}

fn checkbox_obj(id: u16, state: u16, flags: u16) -> Vec<u8> {
    let mut cbls = state.to_le_bytes().to_vec();
    cbls.extend_from_slice(&0u16.to_le_bytes()); // accel
    cbls.extend_from_slice(&0u16.to_le_bytes()); // reserved
    cbls.extend_from_slice(&flags.to_le_bytes());
    obj_record(&[
        cmo(0x000B, id),
        subrecord(FT_CBLS_DATA, &cbls),
        subrecord(FT_END, &[]),
    ])
}

#[test]
fn parses_checkbox_data() {
    let record = checkbox_obj(5, 1, 1);
    let control = form_control(&record);
    assert_eq!(control.object_id(), 5);
    assert_eq!(control.control_type(), Some(ObjectType::CheckBox));
    let data = control.check_box_data().unwrap();
    assert_eq!(data.state, CheckState::Checked);
    assert_eq!(data.accelerator, 0);
    assert!(data.no_3d());
    round_trip(&record);
}

#[test]
fn parses_mixed_and_unchecked_states() {
    for (code, state) in [(0u16, CheckState::Unchecked), (2, CheckState::Mixed)] {
        let control = form_control(&checkbox_obj(1, code, 0));
        assert_eq!(control.check_box_data().unwrap().state, state);
    }
}

#[test]
fn parses_radio_button_data() {
    let mut rbo = 9u16.to_le_bytes().to_vec();
    rbo.extend_from_slice(&1u16.to_le_bytes());
    let record = obj_record(&[
        cmo(0x000C, 3),
        subrecord(FT_RBO_DATA, &rbo),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    assert_eq!(control.control_type(), Some(ObjectType::RadioButton));
    let data = control.radio_button_data().unwrap();
    assert_eq!(data.next_radio_button_id, 9);
    assert!(data.first_in_group);
    round_trip(&record);
}

#[test]
fn parses_edit_box_data() {
    let mut edo = 2u16.to_le_bytes().to_vec(); // ivtEdit: number
    edo.extend_from_slice(&1u16.to_le_bytes()); // fMultiLine
    edo.extend_from_slice(&0u16.to_le_bytes()); // fVScroll
    edo.extend_from_slice(&7u16.to_le_bytes()); // id
    let record = obj_record(&[
        cmo(0x000D, 4),
        subrecord(FT_EDO_DATA, &edo),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    assert_eq!(control.control_type(), Some(ObjectType::EditBox));
    let data = control.edit_box_data().unwrap();
    assert_eq!(data.validation, EditBoxValidation::Number);
    assert!(data.multi_line);
    assert!(!data.vertical_scroll_bar);
    assert_eq!(data.list_control_id, 7);
    round_trip(&record);
}

#[test]
fn parses_group_box_data() {
    let mut gbo = 0x41u16.to_le_bytes().to_vec(); // accel 'A'
    gbo.extend_from_slice(&0u16.to_le_bytes());
    gbo.extend_from_slice(&1u16.to_le_bytes()); // fNo3d
    let record = obj_record(&[
        cmo(0x0013, 8),
        subrecord(FT_GBO_DATA, &gbo),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    assert_eq!(control.control_type(), Some(ObjectType::GroupBox));
    let data = control.group_box_data().unwrap();
    assert_eq!(data.accelerator, 0x41);
    assert!(data.no_3d());
    round_trip(&record);
}

#[test]
fn parses_scroll_bar_data() {
    let mut sbs = vec![0xEE; 4]; // unused1, preserved verbatim
    for value in [-5i16, -10, 10, 1, 5] {
        sbs.extend_from_slice(&value.to_le_bytes());
    }
    sbs.extend_from_slice(&1u16.to_le_bytes()); // fHoriz
    sbs.extend_from_slice(&16i16.to_le_bytes()); // dxScroll
    sbs.extend_from_slice(&0x000Fu16.to_le_bytes()); // all four flags
    let record = obj_record(&[
        cmo(0x0011, 2),
        subrecord(FT_SBS, &sbs),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    assert_eq!(control.control_type(), Some(ObjectType::ScrollBar));
    let data = control.scroll_bar_data().unwrap();
    assert_eq!(data.reserved, [0xEE; 4]);
    assert_eq!((data.value, data.minimum, data.maximum), (-5, -10, 10));
    assert_eq!((data.increment, data.page_increment), (1, 5));
    assert!(data.horizontal);
    assert_eq!(data.scroll_width, 16);
    assert!(data.draw() && data.draw_slider_only() && data.track_elevator() && data.no_3d());
    round_trip(&record);
}

#[test]
fn preserves_invalid_scroll_bar_ranges_as_unknown() {
    for values in [
        [6i16, 0, 5, 1, 1, 1],
        [0, 5, 4, 1, 1, 1],
        [0, 0, 5, -1, 1, 1],
        [0, 0, 5, 1, -1, 1],
        [0, 0, 5, 1, 1, -1],
    ] {
        let mut sbs = vec![0xEE; 4];
        for value in values[..5].iter().copied() {
            sbs.extend_from_slice(&value.to_le_bytes());
        }
        sbs.extend_from_slice(&0u16.to_le_bytes());
        sbs.extend_from_slice(&values[5].to_le_bytes());
        sbs.extend_from_slice(&0u16.to_le_bytes());
        let record = obj_record(&[
            cmo(0x0011, 2),
            subrecord(FT_SBS, &sbs),
            subrecord(FT_END, &[]),
        ]);
        let control = form_control(&record);
        assert!(control.scroll_bar_data().is_none());
        assert!(matches!(
            &control.subrecords[1],
            ObjSubrecord::Unknown { kind: FT_SBS, data } if data == &sbs
        ));
        round_trip(&record);
    }
}

#[test]
fn parses_list_box_data_with_items_and_selection() {
    let mut lbs = 4u16.to_le_bytes().to_vec(); // cbFmla
    lbs.extend_from_slice(&[0xAA, 0xBB, 0xCC, 0xDD]); // fmla
    lbs.extend_from_slice(&3u16.to_le_bytes()); // cLines
    lbs.extend_from_slice(&2u16.to_le_bytes()); // iSel
    // fValidPlex | wListSelType = Multi | lct = Regular
    lbs.extend_from_slice(&0x0012u16.to_le_bytes());
    lbs.extend_from_slice(&0u16.to_le_bytes()); // idEdit
    for item in ["one", "two", "three"] {
        lbs.extend_from_slice(&xl_string(item));
    }
    lbs.extend_from_slice(&[1, 0, 1]); // bsels
    let record = obj_record(&[
        cmo(0x0012, 6),
        subrecord(FT_LBS_DATA, &lbs),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    assert_eq!(control.control_type(), Some(ObjectType::List));
    let data = control.list_box_data().unwrap();
    assert_eq!(data.formula, [0xAA, 0xBB, 0xCC, 0xDD]);
    assert_eq!(data.entry_count, 3);
    assert_eq!(data.selected_index, 2);
    assert!(data.has_item_strings());
    assert!(!data.has_behavior_class());
    assert!(!data.no_3d());
    assert_eq!(data.selection_type(), ListSelectionType::Multi);
    assert_eq!(data.behavior_class(), ListBehaviorClass::Regular);
    assert!(data.drop_down.is_none());
    assert_eq!(
        data.items()
            .iter()
            .map(|item| item.text())
            .collect::<Vec<_>>(),
        vec!["one", "two", "three"]
    );
    assert_eq!(data.multi_selection, vec![true, false, true]);
    assert!(data.trailing.is_empty());
    round_trip(&record);
}

#[test]
fn parses_dropdown_data() {
    let mut lbs = 2u16.to_le_bytes().to_vec(); // cbFmla
    lbs.extend_from_slice(&[0x01, 0x02]);
    lbs.extend_from_slice(&2u16.to_le_bytes()); // cLines
    lbs.extend_from_slice(&1u16.to_le_bytes()); // iSel
    // fValidPlex | fValidIds | wListSelType = CtrlMulti | lct = DataValidation
    lbs.extend_from_slice(&0x0626u16.to_le_bytes());
    lbs.extend_from_slice(&5u16.to_le_bytes()); // idEdit
    // LbsDropData: wStyle = ComboEdit | fFiltered, cLine = 8, dxMin = 100
    lbs.extend_from_slice(&0x0005u16.to_le_bytes());
    lbs.extend_from_slice(&8u16.to_le_bytes());
    lbs.extend_from_slice(&100u16.to_le_bytes());
    lbs.extend_from_slice(&xl_string("abcd")); // odd size: padding follows
    lbs.push(0x00); // unused3
    lbs.extend_from_slice(&xl_string("x"));
    lbs.extend_from_slice(&xl_string("yz"));
    lbs.extend_from_slice(&[1, 0]); // bsels
    let record = obj_record(&[
        cmo(0x0014, 7),
        subrecord(FT_LBS_DATA, &lbs),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    assert_eq!(control.control_type(), Some(ObjectType::DropDown));
    let data = control.list_box_data().unwrap();
    assert_eq!(data.selection_type(), ListSelectionType::CtrlMulti);
    assert_eq!(data.behavior_class(), ListBehaviorClass::DataValidation);
    assert_eq!(data.edit_box_id, 5);
    assert!(data.has_edit_box());
    let drop_down = data.drop_down.as_ref().unwrap();
    assert_eq!(drop_down.style(), DropDownStyle::ComboEdit);
    assert!(drop_down.filtered());
    assert_eq!((drop_down.line_count, drop_down.min_width), (8, 100));
    assert_eq!(drop_down.text(), "abcd");
    assert_eq!(drop_down.padding, Some(0));
    assert_eq!(
        data.items()
            .iter()
            .map(|item| item.text())
            .collect::<Vec<_>>(),
        vec!["x", "yz"]
    );
    assert_eq!(data.multi_selection, vec![true, false]);
    round_trip(&record);
}

#[test]
fn vacant_lbs_data_round_trips_empty() {
    // cbFContinued == 0: no further fields exist (MS-XLS 2.5.147).
    let record = obj_record(&[
        cmo(0x0012, 6),
        subrecord(FT_LBS_DATA, &[]),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    let data = control.list_box_data().unwrap();
    assert_eq!(data.entry_count, 0);
    assert!(data.items().is_empty());
    round_trip(&record);
}

#[test]
fn preserves_invalid_list_header_as_unknown() {
    for (entry_count, selected_index) in [(2u16, 3u16), (0x8000, 0)] {
        let mut lbs = 0u16.to_le_bytes().to_vec();
        lbs.extend_from_slice(&entry_count.to_le_bytes());
        lbs.extend_from_slice(&selected_index.to_le_bytes());
        lbs.extend_from_slice(&0u16.to_le_bytes());
        lbs.extend_from_slice(&0u16.to_le_bytes());
        let record = obj_record(&[
            cmo(0x0012, 6),
            subrecord(FT_LBS_DATA, &lbs),
            subrecord(FT_END, &[]),
        ]);
        let control = form_control(&record);
        assert!(control.list_box_data().is_none());
        assert!(matches!(
            &control.subrecords[1],
            ObjSubrecord::Unknown { kind: FT_LBS_DATA, data } if data == &lbs
        ));
        round_trip(&record);
    }
}

#[test]
fn list_item_authoring_respects_ms_xls_character_limit() {
    assert!(LbsItem::new(&"x".repeat(255)).is_some());
    assert!(LbsItem::new(&"x".repeat(256)).is_none());
}

#[test]
fn malformed_subrecords_fall_back_to_unknown() {
    // Wrong length.
    let record = obj_record(&[
        cmo(0x000B, 1),
        subrecord(FT_CBLS_DATA, &[0; 7]),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    assert!(control.check_box_data().is_none());
    assert!(matches!(
        &control.subrecords[1],
        ObjSubrecord::Unknown { kind: FT_CBLS_DATA, data } if data == &[0; 7]
    ));
    round_trip(&record);

    // Out-of-range checked state.
    let record = checkbox_obj(1, 9, 0);
    let control = form_control(&record);
    assert!(control.check_box_data().is_none());
    assert!(matches!(
        &control.subrecords[1],
        ObjSubrecord::Unknown { .. }
    ));
    round_trip(&record);

    // Truncated scroll bar data.
    let record = obj_record(&[
        cmo(0x0011, 2),
        subrecord(FT_SBS, &[0; 19]),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    assert!(control.scroll_bar_data().is_none());
    round_trip(&record);

    // Out-of-range Boolean in radio button data.
    let mut rbo = 9u16.to_le_bytes().to_vec();
    rbo.extend_from_slice(&2u16.to_le_bytes());
    let record = obj_record(&[
        cmo(0x000C, 3),
        subrecord(FT_RBO_DATA, &rbo),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    assert!(control.radio_button_data().is_none());
    round_trip(&record);
}

#[test]
fn lbs_data_with_continued_items_preserves_tail() {
    let mut lbs = 0u16.to_le_bytes().to_vec(); // cbFmla
    lbs.extend_from_slice(&2u16.to_le_bytes()); // cLines
    lbs.extend_from_slice(&0u16.to_le_bytes()); // iSel
    lbs.extend_from_slice(&0x0002u16.to_le_bytes()); // fValidPlex, single selection
    lbs.extend_from_slice(&0u16.to_le_bytes()); // idEdit
    lbs.extend_from_slice(&xl_string("ab"));
    // The second string spills into a Continue record this parser never sees.
    let record = obj_record(&[
        cmo(0x0012, 6),
        subrecord(FT_LBS_DATA, &lbs),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    let data = control.list_box_data().unwrap();
    assert_eq!(
        data.items()
            .iter()
            .map(|item| item.text())
            .collect::<Vec<_>>(),
        vec!["ab"]
    );
    assert!(data.trailing.is_empty());
    round_trip(&record);
}

#[test]
fn lbs_data_with_defective_item_preserves_tail_bytes() {
    let mut lbs = 0u16.to_le_bytes().to_vec(); // cbFmla
    lbs.extend_from_slice(&2u16.to_le_bytes()); // cLines
    lbs.extend_from_slice(&0u16.to_le_bytes()); // iSel
    lbs.extend_from_slice(&0x0002u16.to_le_bytes()); // fValidPlex
    lbs.extend_from_slice(&0u16.to_le_bytes()); // idEdit
    lbs.extend_from_slice(&xl_string("ab"));
    // A string header claiming 10 characters but supplying only 2 bytes.
    lbs.extend_from_slice(&[10, 0, 0, b'x', b'y']);
    let record = obj_record(&[
        cmo(0x0012, 6),
        subrecord(FT_LBS_DATA, &lbs),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    let data = control.list_box_data().unwrap();
    assert_eq!(
        data.items()
            .iter()
            .map(|item| item.text())
            .collect::<Vec<_>>(),
        vec!["ab"]
    );
    assert_eq!(data.trailing, [10, 0, 0, b'x', b'y']);
    round_trip(&record);
}

#[test]
fn utf16_items_round_trip() {
    let mut lbs = 0u16.to_le_bytes().to_vec();
    lbs.extend_from_slice(&1u16.to_le_bytes()); // cLines
    lbs.extend_from_slice(&0u16.to_le_bytes()); // iSel
    lbs.extend_from_slice(&0x0002u16.to_le_bytes()); // fValidPlex
    lbs.extend_from_slice(&0u16.to_le_bytes()); // idEdit
    // "漢字" uncompressed.
    lbs.extend_from_slice(&2u16.to_le_bytes());
    lbs.push(1); // fHighByte
    lbs.extend_from_slice(&[0x22, 0x6F, 0x57, 0x5B]);
    let record = obj_record(&[
        cmo(0x0012, 6),
        subrecord(FT_LBS_DATA, &lbs),
        subrecord(FT_END, &[]),
    ]);
    let control = form_control(&record);
    let data = control.list_box_data().unwrap();
    assert_eq!(data.items()[0].text(), "漢字");
    round_trip(&record);
}

#[test]
fn non_control_objects_are_not_form_controls() {
    // A picture Obj (ot = 0x08) without the OLE-specific subrecords.
    let picture = obj_record(&[cmo(0x0008, 1), subrecord(FT_END, &[])]);
    let body = &picture[4..];
    assert!(FormControl::parse(body, None).is_none());
    // A note Obj (ot = 0x19).
    let note = obj_record(&[cmo(0x0019, 1), subrecord(FT_END, &[])]);
    assert!(FormControl::parse(&note[4..], None).is_none());
    // 0x000A is reserved between Polygon and Checkbox in cmo.ot.
    let reserved = obj_record(&[cmo(0x000A, 1), subrecord(FT_END, &[])]);
    assert!(FormControl::parse(&reserved[4..], None).is_none());
}

fn ole_object(id: u16, storage: u32) -> OleObjectRecord {
    OleObjectRecord {
        subrecords: vec![
            ObjSubrecord::Common(FtCmo {
                object_type: 8,
                object_id: id,
                flags: 0,
                reserved: [0; 12],
            }),
            ObjSubrecord::PictureFlags(FtPioGrbit { raw: 0 }),
            ObjSubrecord::PictureFormula(FtPictFmla {
                formula: vec![1, 2, 3],
                storage_position: Some(storage),
                control_buffer_size: Some(0),
            }),
            ObjSubrecord::End,
        ],
        text_object: None,
    }
}

fn nested_cfb(marker: &[u8]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["CONTENTS"], marker).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn workbook_stream(sheet_records: &[Vec<u8>]) -> Vec<u8> {
    let bof = subrecord(0x0809, &[0; 16]);
    let eof = subrecord(0x000A, &[]);
    let mut bound_body = vec![0; 8];
    bound_body[6] = 1;
    bound_body[7] = b'S';
    let bound = subrecord(0x0085, &bound_body);
    let globals_len = bof.len() + bound.len() + eof.len();
    let mut output = bof;
    let mut bound = bound;
    bound[4..8].copy_from_slice(&(globals_len as u32).to_le_bytes());
    output.extend_from_slice(&bound);
    output.extend_from_slice(&eof);
    output.extend_from_slice(&subrecord(0x0809, &[0; 16]));
    for record in sheet_records {
        output.extend_from_slice(record);
    }
    output.extend_from_slice(&eof);
    output
}

fn xls(sheet_records: &[Vec<u8>], storages: &[u32]) -> Vec<u8> {
    let workbook = workbook_stream(sheet_records);
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], &workbook).unwrap();
    for id in storages {
        let name = format!("MBD{id:08X}");
        writer.create_storage(&[&name]).unwrap();
        let mut nested = OleFile::open(Cursor::new(nested_cfb(&id.to_le_bytes()))).unwrap();
        let contents = nested.open_stream(&["CONTENTS"]).unwrap();
        writer
            .create_stream(&[&name, "CONTENTS"], &contents)
            .unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn editor_enumerates_form_controls_and_preserves_them_on_rewrite() {
    let checkbox = checkbox_obj(5, 1, 1);
    let mut rbo = 0u16.to_le_bytes().to_vec();
    rbo.extend_from_slice(&1u16.to_le_bytes());
    let radio = obj_record(&[
        cmo(0x000C, 6),
        subrecord(FT_RBO_DATA, &rbo),
        subrecord(FT_END, &[]),
    ]);
    let ole = ole_object(1, 42).to_record_bytes().unwrap();
    let bytes = xls(&[ole, checkbox.clone(), radio], &[42]);

    let editor = Editor::new(bytes.clone(), Limits::default()).unwrap();
    assert_eq!(editor.objects(0).unwrap().len(), 1);
    let controls = editor.form_controls(0).unwrap();
    assert_eq!(controls.len(), 2);
    assert_eq!(controls[0].control_type(), Some(ObjectType::CheckBox));
    assert_eq!(
        controls[0].check_box_data().unwrap().state,
        CheckState::Checked
    );
    assert_eq!(controls[1].control_type(), Some(ObjectType::RadioButton));
    assert!(controls[1].radio_button_data().unwrap().first_in_group);
    assert!(editor.form_controls(9).is_err());

    // Rewriting the workbook for an unrelated OLE edit keeps the form
    // control Obj records byte-identical.
    let mut editor = Editor::new(bytes, Limits::default()).unwrap();
    editor.remove(0, 1).unwrap();
    let bytes = editor.finish().unwrap();
    let mut ole_file = OleFile::open(Cursor::new(bytes.clone())).unwrap();
    let workbook = ole_file.open_stream(&["Workbook"]).unwrap();
    assert!(
        workbook
            .windows(checkbox.len())
            .any(|window| window == checkbox)
    );

    let editor = Editor::new(bytes, Limits::default()).unwrap();
    let controls = editor.form_controls(0).unwrap();
    assert_eq!(controls.len(), 2);
    assert_eq!(
        controls[0].check_box_data().unwrap().state,
        CheckState::Checked
    );
    assert_eq!(controls[0].to_record_bytes().unwrap(), checkbox);
}
