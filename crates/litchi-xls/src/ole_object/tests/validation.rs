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

fn valid_sbs() -> FtSbs {
    FtSbs {
        reserved: [0; 4],
        value: 0,
        minimum: 0,
        maximum: 10,
        increment: 1,
        page_increment: 2,
        horizontal: false,
        scroll_width: 1,
        flags: 0,
    }
}

#[test]
fn form_control_constructor_authors_typed_checkbox() {
    let control = FormControl::new(
        ObjectType::CheckBox,
        41,
        [ObjSubrecord::CheckBoxData(FtCblsData {
            state: CheckState::Checked,
            accelerator: 0,
            reserved: 0,
            flags: 0,
        })],
    );
    control
        .validate()
        .expect("constructed checkbox should validate");
    let bytes = control
        .to_record_bytes()
        .expect("constructed checkbox should encode");
    assert_eq!(
        FormControl::parse(&bytes[4..], None)
            .expect("encoded checkbox should parse")
            .object_id(),
        41
    );
}

#[test]
fn list_authoring_synchronizes_items_and_requires_sbs() {
    let mut list = FtLbsData::from_texts(["one", "二"]).expect("list strings should encode");
    assert_eq!(list.entry_count, 2);
    assert!(list.has_item_strings());
    list.set_items_checked(vec![LbsItem::new("only").expect("short item")])
        .expect("item replacement should be checked");
    assert_eq!(list.entry_count, 1);
    assert_eq!(list.items()[0].text(), "only");
    list.validate().expect("synchronized list should validate");

    let missing_sbs = FormControl::new(
        ObjectType::List,
        42,
        [ObjSubrecord::ListBoxData(list.clone())],
    );
    assert!(missing_sbs.validate().is_err());

    let valid = FormControl::new(
        ObjectType::List,
        42,
        [
            ObjSubrecord::ScrollBarData(valid_sbs()),
            ObjSubrecord::ListBoxData(list),
        ],
    );
    valid.validate().expect("complete list should validate");
}

#[test]
fn dropdown_authoring_requires_matching_drop_data() {
    let drop_down = LbsDropData::try_new(DropDownStyle::ComboEdit, 8, 100, "choose")
        .expect("dropdown metadata should encode");
    assert!(LbsDropData::new(DropDownStyle::Reserved, 8, 100, "choose").is_none());
    let mut data = FtLbsData::from_texts(["one", "two"]).expect("dropdown items should encode");
    data.drop_down = Some(drop_down);

    let dropdown = FormControl::new(
        ObjectType::DropDown,
        43,
        [
            ObjSubrecord::ScrollBarData(valid_sbs()),
            ObjSubrecord::ListBoxData(data.clone()),
        ],
    );
    dropdown
        .validate()
        .expect("matching dropdown metadata should validate");

    data.drop_down = None;
    let missing_drop_data = FormControl::new(
        ObjectType::DropDown,
        43,
        [
            ObjSubrecord::ScrollBarData(valid_sbs()),
            ObjSubrecord::ListBoxData(data),
        ],
    );
    assert!(missing_drop_data.validate().is_err());
}

#[test]
fn list_validation_rejects_incomplete_typed_arrays() {
    let mut list = FtLbsData::from_texts(["one"]).expect("list should encode");
    list.entry_count = 2;
    assert!(list.validate().is_err());

    list.entry_count = 1;
    list.flags |= 1 << LBS_SELECTION_TYPE_SHIFT;
    assert!(list.validate().is_err());
}
