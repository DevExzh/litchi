use super::{Control, Toolbar, ToolbarSet, VisualData, Wrapper, parse, to_bytes};
use litchi_ole_common::toolbar::{
    ControlFlags, ControlHeader, ControlType, Flags, Header, Restrictions, SpecificFlags, Type,
    WString,
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
fn variable_tbc_data_is_rejected_before_boundary_guessing() {
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
    bytes.extend_from_slice(&header(1).to_bytes());
    let control_header = ControlHeader::new(
        ControlType::Button,
        1,
        ControlFlags::from_raw(0),
        SpecificFlags::from_raw(0),
        0,
        None,
    )
    .expect("button header");
    bytes.extend_from_slice(&super::APPLICATION_TOOLBAR_ID.to_le_bytes());
    bytes.extend_from_slice(&control_header.to_bytes());

    let error = parse(&bytes).expect_err("variable TBCData must be refused");
    assert!(
        matches!(error, crate::Error::UnsupportedFeature(message) if message.contains("TBCData"))
    );
}
