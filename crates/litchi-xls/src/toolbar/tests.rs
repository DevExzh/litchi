use super::{Command, Control, Toolbar, ToolbarSet, VisualData, Wrapper, parse, to_bytes};
use litchi_ole_common::toolbar::{
    ControlFlags, ControlHeader, ControlType, Data, Flags, GeneralFlags, GeneralInfo, Header,
    Restrictions, SpecificFlags, Type, WString,
};

fn header(control_count: i16) -> Header<'static> {
    Header::new(
        control_count,
        Restrictions::new(Type::Basic).expect("basic restrictions"),
        1,
        Flags::from_raw(0),
        WString::new("Custom toolbar").expect("toolbar name"),
    )
    .expect("toolbar header")
}

fn toolbar(with_visual_data: bool) -> Toolbar<'static> {
    let toolbar = Toolbar::new(header(0), super::APPLICATION_TOOLBAR_ID).expect("toolbar");
    if with_visual_data {
        toolbar.with_visual_data(VisualData::new([0xA5; 60]))
    } else {
        toolbar
    }
}

fn active_x_toolbar() -> Toolbar<'static> {
    let control_header = ControlHeader::new(
        ControlType::ActiveX,
        0x1234,
        ControlFlags::from_raw(0),
        SpecificFlags::from_raw(0),
        0,
        None,
    )
    .expect("ActiveX header");
    let control = Control::new(control_header).expect("ActiveX control");
    Toolbar::from_parts(
        header(1),
        None,
        super::APPLICATION_TOOLBAR_ID,
        vec![control],
    )
    .expect("ActiveX toolbar")
}

fn data(specific: &[u8]) -> Data<'static> {
    let general =
        GeneralInfo::new(GeneralFlags::default(), None, None, None, None).expect("general info");
    Data::new(general, specific.to_vec()).expect("TBCData")
}

fn data_toolbar(with_command: bool) -> Toolbar<'static> {
    let control_header = ControlHeader::new(
        ControlType::Button,
        0x1234,
        ControlFlags::from_raw(0),
        SpecificFlags::from_raw(0),
        0,
        None,
    )
    .expect("button header");
    let command = with_command.then(|| Command::new(0x1234, 0x01, false).expect("TBCCmd"));
    let control = Control::from_parts(control_header, command, Some(data(&[0xA5, 0x03, 0x01])))
        .expect("button control");
    Toolbar::from_parts(
        header(1),
        None,
        super::APPLICATION_TOOLBAR_ID,
        vec![control],
    )
    .expect("button toolbar")
}

#[test]
fn empty_toolbar_roundtrips_without_visual_data() {
    let value = Wrapper::new(vec![toolbar(false)]).expect("wrapper");
    let bytes = to_bytes(&value).expect("encode");
    let parsed = parse(&bytes).expect("decode");

    assert_eq!(parsed, value);
    assert_eq!(parsed.toolbar_set().toolbar_count(), 1);
    assert_eq!(
        parsed.toolbars()[0].header().name().text(),
        "Custom toolbar"
    );
}

#[test]
fn optional_visual_data_is_lossless_and_views_are_bounded() {
    let value = Wrapper::new(vec![toolbar(true)]).expect("wrapper");
    let bytes = to_bytes(&value).expect("encode");
    let parsed = parse(&bytes).expect("decode");
    let visual = parsed.toolbars()[0].visual_data().expect("visual data");

    assert_eq!(visual.bytes(), &[0xA5; 60]);
    assert_eq!(visual.view(0).expect("normal view"), &[0xA5; 20]);
    assert!(visual.view(3).is_none());
}

#[test]
fn parsed_toolbar_can_be_owned_for_workbook_lifetime() {
    let value = Wrapper::new(vec![toolbar(true)]).expect("wrapper");
    let bytes = to_bytes(&value).expect("encode");
    let owned = parse(&bytes).expect("decode").into_owned();

    assert_eq!(owned.toolbars()[0].header().name().text(), "Custom toolbar");
    assert_eq!(
        owned.toolbars()[0].visual_data().unwrap().bytes(),
        &[0xA5; 60]
    );
    assert_eq!(to_bytes(&owned).expect("re-encode"), bytes);
}

#[test]
fn multiple_toolbars_keep_optional_records_separate() {
    let value = Wrapper::new(vec![toolbar(false), toolbar(true)]).expect("wrapper");
    let bytes = to_bytes(&value).expect("encode");
    let parsed = parse(&bytes).expect("decode");

    assert_eq!(parsed.toolbars().len(), 2);
    assert!(parsed.toolbars()[0].visual_data().is_none());
    assert!(parsed.toolbars()[1].visual_data().is_some());
}

#[test]
fn fixed_header_activex_controls_are_typed_without_executing_them() {
    let value = Wrapper::new(vec![active_x_toolbar()]).expect("wrapper");
    let bytes = to_bytes(&value).expect("encode");
    let parsed = parse(&bytes).expect("decode");

    assert_eq!(parsed.toolbars()[0].controls().len(), 1);
    assert!(parsed.toolbars()[0].controls()[0].is_active_x());
    assert_eq!(
        parsed.toolbars()[0].controls()[0].header().control_id(),
        0x1234
    );
}

#[test]
fn reserved_set_fields_and_visual_bytes_are_preserved() {
    let set = ToolbarSet::from_parts(1, 1, 0x1122, 0x3344, 0x5566, 1, 3, 1).expect("toolbar set");
    let value = Wrapper::from_parts(set, vec![toolbar(true)]).expect("wrapper");
    let bytes = to_bytes(&value).expect("encode");
    let parsed = parse(&bytes).expect("decode");

    assert_eq!(parsed.toolbar_set().reserved1(), 0x1122);
    assert_eq!(parsed.toolbar_set().reserved2(), 0x3344);
    assert_eq!(parsed.toolbar_set().reserved3(), 0x5566);
    assert_eq!(parsed.toolbar_set().active_view(), 1);
    assert_eq!(
        parsed.toolbars()[0].visual_data().unwrap().bytes(),
        &[0xA5; 60]
    );
}

#[test]
fn malformed_signatures_counts_and_lengths_are_rejected() {
    let value = Wrapper::new(vec![toolbar(false)]).expect("wrapper");
    let bytes = to_bytes(&value).expect("encode");

    let mut bad_signature = bytes.clone();
    bad_signature[0] = 0;
    assert!(parse(&bad_signature).is_err());

    let mut bad_view = bytes.clone();
    bad_view[12..14].copy_from_slice(&2u16.to_le_bytes());
    assert!(parse(&bad_view).is_err());

    let mut bad_application_id = bytes.clone();
    let application_offset = 14 + header(0).to_bytes().len();
    bad_application_id[application_offset..application_offset + 4]
        .copy_from_slice(&0x1000i32.to_le_bytes());
    assert!(parse(&bad_application_id).is_err());

    assert!(parse(&bytes[..bytes.len() - 1]).is_err());
}

#[test]
fn variable_tbc_data_and_tbccmd_roundtrip_without_execution() {
    let value = Wrapper::new(vec![data_toolbar(true)]).expect("wrapper");
    let bytes = to_bytes(&value).expect("encode");
    let parsed = parse(&bytes).expect("decode");
    let control = &parsed.toolbars()[0].controls()[0];

    assert!(!control.is_active_x());
    assert_eq!(control.command().expect("TBCCmd").command_id(), 0x1234);
    assert_eq!(control.command().expect("TBCCmd").command_type(), 0x01);
    assert_eq!(control.general().expect("general info").flags().raw(), 0);
    assert_eq!(
        control.data().expect("TBCData").specific(),
        &[0xA5, 0x03, 0x01]
    );
    assert_eq!(to_bytes(&parsed).expect("re-encode"), bytes);
}

#[test]
fn shared_general_and_extra_info_are_decoded_losslessly() {
    let set = ToolbarSet::new(1, 0).expect("toolbar set");
    let control_header = ControlHeader::new(
        ControlType::Button,
        1,
        ControlFlags::from_raw(0),
        SpecificFlags::from_raw(0),
        0,
        None,
    )
    .expect("button header");
    let mut bytes = Vec::new();
    bytes.push(set.signature());
    bytes.push(set.version());
    bytes.extend_from_slice(&set.reserved1().to_le_bytes());
    bytes.extend_from_slice(&set.reserved2().to_le_bytes());
    bytes.extend_from_slice(&set.reserved3().to_le_bytes());
    bytes.extend_from_slice(&set.toolbar_count().to_le_bytes());
    bytes.extend_from_slice(&set.view_count().to_le_bytes());
    bytes.extend_from_slice(&set.active_view().to_le_bytes());
    bytes.extend_from_slice(&header(1).to_bytes());
    bytes.extend_from_slice(&super::APPLICATION_TOOLBAR_ID.to_le_bytes());
    bytes.extend_from_slice(&control_header.to_bytes());

    let mut data = vec![GeneralFlags::default().with_save_misc_custom(true).raw()];
    data.extend_from_slice(&WString::new("").expect("help file").to_bytes());
    data.extend_from_slice(&0i32.to_le_bytes());
    data.extend_from_slice(&WString::new("Tag").expect("tag").to_bytes());
    data.extend_from_slice(&WString::new("Macro").expect("action").to_bytes());
    data.extend_from_slice(&WString::new("argument").expect("parameter").to_bytes());
    data.extend_from_slice(&[0, 0]);
    bytes.extend_from_slice(&data);

    let parsed = parse(&bytes).expect("decode");
    let control = &parsed.toolbars()[0].controls()[0];
    assert_eq!(control.general().expect("general info").flags().raw(), 0x04);
    assert_eq!(control.extra().expect("extra info").tag().text(), "Tag");
    assert_eq!(
        control.extra().expect("extra info").on_action().text(),
        "Macro"
    );
    assert_eq!(to_bytes(&parsed).expect("re-encode"), bytes);
}

#[test]
fn multiple_variable_controls_use_the_next_tbc_header_boundary() {
    let first_header = ControlHeader::new(
        ControlType::Button,
        1,
        ControlFlags::from_raw(0),
        SpecificFlags::from_raw(0),
        0,
        None,
    )
    .expect("first label header");
    let second_header = ControlHeader::new(
        ControlType::Button,
        1,
        ControlFlags::from_raw(0),
        SpecificFlags::from_raw(0),
        0,
        None,
    )
    .expect("second button header");
    let second =
        Control::from_parts(second_header, None, Some(data(&[0xCC]))).expect("second control");
    let first =
        Control::from_parts(first_header, None, Some(data(&[0xAA]))).expect("first control");
    let value = Toolbar::from_parts(
        header(2),
        None,
        super::APPLICATION_TOOLBAR_ID,
        vec![first, second],
    )
    .expect("toolbar");
    let wrapper = Wrapper::new(vec![value]).expect("wrapper");
    let bytes = to_bytes(&wrapper).expect("encode");
    let parsed = parse(&bytes).expect("decode");

    assert_eq!(parsed.toolbars()[0].controls().len(), 2);
    assert_eq!(
        parsed.toolbars()[0].controls()[0]
            .data()
            .expect("first data")
            .specific(),
        &[0xAA]
    );
    assert_eq!(
        parsed.toolbars()[0].controls()[1]
            .data()
            .expect("second data")
            .specific(),
        &[0xCC]
    );
    assert_eq!(to_bytes(&parsed).expect("re-encode"), bytes);
}

#[test]
fn ambiguous_tbcdata_boundaries_are_rejected() {
    let set = ToolbarSet::new(1, 0).expect("toolbar set");
    let mut bytes = Vec::with_capacity(14);
    bytes.push(set.signature());
    bytes.push(set.version());
    bytes.extend_from_slice(&set.reserved1().to_le_bytes());
    bytes.extend_from_slice(&set.reserved2().to_le_bytes());
    bytes.extend_from_slice(&set.reserved3().to_le_bytes());
    bytes.extend_from_slice(&set.toolbar_count().to_le_bytes());
    bytes.extend_from_slice(&set.view_count().to_le_bytes());
    bytes.extend_from_slice(&set.active_view().to_le_bytes());
    bytes.extend_from_slice(&header(2).to_bytes());
    let first_header = ControlHeader::new(
        ControlType::Button,
        1,
        ControlFlags::from_raw(0),
        SpecificFlags::from_raw(0),
        0,
        None,
    )
    .expect("first header");
    let fake_header = ControlHeader::new(
        ControlType::Button,
        1,
        ControlFlags::from_raw(0),
        SpecificFlags::from_raw(0),
        0,
        None,
    )
    .expect("fake header");
    let second_header = ControlHeader::new(
        ControlType::ActiveX,
        2,
        ControlFlags::from_raw(0),
        SpecificFlags::from_raw(0),
        0,
        None,
    )
    .expect("second header");
    bytes.extend_from_slice(&super::APPLICATION_TOOLBAR_ID.to_le_bytes());
    bytes.extend_from_slice(&first_header.to_bytes());
    bytes.push(0);
    bytes.extend_from_slice(&fake_header.to_bytes());
    bytes.push(0);
    bytes.extend_from_slice(&second_header.to_bytes());

    let error = parse(&bytes).expect_err("ambiguous TBCData must be refused");
    assert!(
        matches!(error, crate::Error::UnsupportedFeature(message) if message.contains("ambiguous"))
    );
}
