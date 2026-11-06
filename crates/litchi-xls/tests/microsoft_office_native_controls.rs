//! Native Excel BIFF8 controls keep their metadata and original record framing.
use litchi_xls::ole_object::{
    CheckState, DropDownStyle, FormControl, FtCblsData, ObjSubrecord, ObjectType,
};

fn path() -> std::path::PathBuf {
    std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/microsoft-office/excel/excel-form-controls.xls")
}
fn obj_records(bytes: &[u8]) -> std::collections::BTreeMap<u16, Vec<u8>> {
    let mut ole = litchi_cfb::OleFile::open(
        std::io::Cursor::new(bytes),
        litchi_cfb::OleFileLimits::default(),
    )
    .unwrap();
    let workbook = ole.open_stream(&["Workbook"]).unwrap();
    let mut result = std::collections::BTreeMap::new();
    let mut offset = 0;
    while offset + 4 <= workbook.len() {
        let kind = u16::from_le_bytes(workbook[offset..offset + 2].try_into().unwrap());
        let size = usize::from(u16::from_le_bytes(
            workbook[offset + 2..offset + 4].try_into().unwrap(),
        ));
        let end = offset + 4 + size;
        let body = &workbook[offset + 4..end];
        if kind == 0x005D {
            assert_eq!(&body[..4], &[0x15, 0, 18, 0]);
            let id = u16::from_le_bytes(body[6..8].try_into().unwrap());
            assert!(result.insert(id, workbook[offset..end].to_vec()).is_none());
        }
        offset = end;
    }
    result
}

#[test]
fn native_excel_form_controls_include_all_eleven_obj_records() {
    let bytes = std::fs::read(path()).unwrap();
    let package = litchi_xls::Package::from_bytes(&bytes).unwrap();
    let owner = package.ole_objects().unwrap();
    let controls = owner.form_controls(0).unwrap();
    assert_eq!(
        controls.iter().map(|c| c.object_id()).collect::<Vec<_>>(),
        [1, 2, 3, 4, 5, 6, 7, 9, 10, 11, 12]
    );
    let records = obj_records(&bytes);
    assert_eq!(records.len(), 11);
    for control in controls {
        control.validate().unwrap();
        assert!(control.to_record_bytes().unwrap() == records[&control.object_id()]);
    }
    assert!(owner.objects(0).unwrap().is_empty());
    assert_eq!(controls[8].control_type(), Some(ObjectType::Button));
    assert_eq!(
        controls[0].check_box_data().unwrap().state,
        CheckState::Checked
    );
    assert_eq!(
        controls[9].check_box_data().unwrap().state,
        CheckState::Mixed
    );
    assert_eq!(
        controls[6].check_box_data().unwrap().state,
        CheckState::Unchecked
    );
    assert_eq!(
        controls[10].check_box_data().unwrap().state,
        CheckState::Checked
    );
    assert_eq!(
        controls[6]
            .radio_button_data()
            .unwrap()
            .next_radio_button_id,
        12
    );
    assert!(controls[6].radio_button_data().unwrap().first_in_group);
    assert_eq!(
        controls[10]
            .radio_button_data()
            .unwrap()
            .next_radio_button_id,
        7
    );
    assert!(!controls[10].radio_button_data().unwrap().first_in_group);
    for (index, selected) in [(1, 2), (2, 3)] {
        let list = controls[index].list_box_data().unwrap();
        assert_eq!(list.entry_count, 3);
        assert_eq!(list.selected_index, selected);
        assert_eq!(list.cb_f_continued(), Some(0x1fde));
        assert!(!list.has_item_strings());
        assert!(list.items().is_empty());
        assert_eq!(
            list.formula,
            [9, 0, 0, 0, 0, 0, 0x25, 8, 0, 10, 0, 3, 0, 3, 0, 3]
        );
    }
    let drop = controls[1]
        .list_box_data()
        .unwrap()
        .drop_down
        .as_ref()
        .unwrap();
    assert_eq!(drop.style(), DropDownStyle::Combo);
    assert_eq!(drop.line_count, 8);
    assert_eq!(drop.padding, Some(10));
    assert_eq!(drop.text(), "");
    for (index, values) in [(3, (5, 1, 20, 2, 10, false)), (4, (10, 0, 50, 2, 5, true))] {
        let data = controls[index].scroll_bar_data().unwrap();
        assert_eq!(
            (
                data.value,
                data.minimum,
                data.maximum,
                data.increment,
                data.page_increment,
                data.horizontal
            ),
            values
        );
    }
    assert!(package.to_bytes().unwrap() == bytes);
}

#[test]
fn native_control_metadata_append_preserves_all_original_obj_bytes() {
    let bytes = std::fs::read(path()).unwrap();
    let original = obj_records(&bytes);
    let package = litchi_xls::Package::from_bytes(&bytes).unwrap();
    let mut edit = package.edit_ole_objects().unwrap();
    edit.add_form_control(
        0,
        FormControl::new(
            ObjectType::CheckBox,
            13,
            [ObjSubrecord::CheckBoxData(FtCblsData {
                state: CheckState::Unchecked,
                accelerator: 0,
                reserved: 0,
                flags: 0,
            })],
        ),
    )
    .unwrap();
    let commit = edit.commit().unwrap();
    let directory = tempfile::tempdir().unwrap();
    let destination = directory.path().join("control-metadata.xls");
    commit.package().save(&destination).unwrap();
    let saved = std::fs::read(&destination).unwrap();
    let reopened = litchi_xls::Package::from_bytes(&saved).unwrap();
    let owner = reopened.ole_objects().unwrap();
    assert_eq!(owner.form_controls(0).unwrap().len(), 12);
    assert_eq!(owner.form_controls(0).unwrap()[11].object_id(), 13);
    let after = obj_records(&saved);
    for (id, record) in original {
        assert!(after[&id] == record, "changed native control {id}");
    }
    assert!(
        reopened
            .apply_ole_objects(&commit.patch().inverse())
            .unwrap()
            .package()
            .to_bytes()
            .unwrap()
            == bytes
    );
    assert!(package.to_bytes().unwrap() == bytes);
}
