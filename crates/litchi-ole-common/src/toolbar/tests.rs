use super::*;

fn ok<T>(result: Result<T, Error>) -> T {
    match result {
        Ok(value) => value,
        Err(error) => panic!("{error}"),
    }
}

fn text(value: &str) -> WString<'static> {
    ok(WString::new(value))
}

#[test]
fn wstring_decodes_borrowed_utf16_and_serializes_exactly() {
    let original = text("A😀");
    let wire = original.to_bytes();
    assert_eq!(wire, vec![3, 0x41, 0x00, 0x3D, 0xD8, 0x00, 0xDE]);

    let (decoded, consumed) = ok(WString::parse_prefix(&wire));
    assert_eq!(consumed, wire.len());
    assert_eq!(decoded.text(), "A😀");
    assert_eq!(decoded.len(), 3);
    assert_eq!(decoded.encoded_bytes().as_ptr(), wire[1..].as_ptr());
    assert_eq!(decoded.to_bytes(), wire);
}

#[test]
fn wstring_rejects_invalid_lengths_nuls_and_surrogates() {
    assert!(WString::parse(&[]).is_err());
    assert!(WString::parse(&[1, 0x41]).is_err());
    assert!(WString::parse(&[1, 0x00, 0x00]).is_err());
    assert!(WString::parse(&[1, 0x00, 0xD8]).is_err());
    assert!(WString::parse(&[1, 0x00, 0xDC]).is_err());
    assert!(WString::new("bad\0text").is_err());

    let too_long = "a".repeat(256);
    assert!(WString::new(&too_long).is_err());
}

#[test]
fn flag_models_preserve_unknown_and_reserved_bits() {
    let control = ControlFlags::from_raw(0xA5);
    assert_eq!(control.reserved_bits(), 0xA0);
    assert_eq!(control.to_bytes(), [0xA5]);
    assert!(ControlFlags::try_from_raw(0xA5).is_err());

    let specific = SpecificFlags::from_raw(0xF900_F1FF);
    assert_eq!(specific.to_bytes(), 0xF900_F1FFu32.to_le_bytes());
    assert_eq!(specific.reserved_bits(), 0x8000_0000);
    assert_ne!(specific.unused_bits(), 0);
    assert!(SpecificFlags::try_from_raw(specific.raw()).is_err());

    let general = GeneralFlags::from_raw(0xF9);
    assert_eq!(general.unused_bits(), 0xF0);
    assert_eq!(general.to_bytes(), [0xF9]);

    let button = ButtonFlags::from_raw(0xE2);
    assert!(matches!(button.state(), ButtonState::Unknown(2)));
    assert!(matches!(button.hyperlink(), HyperlinkType::Unknown(3)));
    assert_eq!(button.to_bytes(), [0xE2]);
    assert!(ButtonFlags::try_from_raw(button.raw()).is_err());
}

#[test]
fn toolbar_header_roundtrips_and_keeps_prefix_parsing() {
    let restrictions = ok(Restrictions::new(Type::Menu));
    let flags = Flags::default()
        .with_controls_modified(true)
        .with_needs_positioning(true);
    let info = ok(Header::new(2, restrictions, 3, flags, text("Custom")));
    let wire = info.to_bytes();
    assert_eq!(&wire[..2], &[0x02, 0x01]);
    assert_eq!(&wire[2..4], &2i16.to_le_bytes());
    assert_eq!(wire[16], 6);

    let mut framed = wire.clone();
    framed.push(0xEE);
    let (decoded, consumed) = ok(Header::parse_prefix(&framed));
    assert_eq!(consumed, wire.len());
    assert_eq!(decoded.name().text(), "Custom");
    assert_eq!(decoded.to_bytes(), wire);
    assert!(Header::parse(&framed).is_err());
}

#[test]
fn toolbar_header_validates_fixed_fields_and_menu_restrictions() {
    let name = text("Toolbar");
    let bad_menu = ok(Restrictions::new(Type::Menu)).with_no_resize(false);
    assert!(Header::new(0, bad_menu, 0, Flags::default(), name.clone()).is_err());
    assert!(
        Header::new(
            0,
            Restrictions::default(),
            256,
            Flags::default(),
            name.clone()
        )
        .is_err()
    );
    assert!(Header::new(0, Restrictions::default(), 0, Flags::from_raw(0x20), name).is_err());

    let mut wire = ok(Header::new(
        0,
        Restrictions::default(),
        0,
        Flags::default(),
        text("Toolbar"),
    ))
    .to_bytes();
    wire[0] = 0;
    assert!(Header::parse(&wire).is_err());
}

#[test]
fn control_header_roundtrips_optional_dimensions() {
    let flags = ControlFlags::default()
        .with_hidden(true)
        .with_save_dimensions(true);
    let specifics = SpecificFlags::default()
        .with_text_icon(TextIcon::TextAndIcon)
        .with_save_ui_strings(true)
        .with_text_below(true);
    let header = ok(ControlHeader::new(
        ControlType::Button,
        0x0071,
        flags,
        specifics,
        3,
        Some(Dimensions::new(10, 20)),
    ));
    let wire = header.to_bytes();
    assert_eq!(wire.len(), 15);
    let (decoded, consumed) = ok(ControlHeader::parse_prefix(&wire));
    assert_eq!(consumed, wire.len());
    assert_eq!(decoded.control_type(), ControlType::Button);
    assert_eq!(decoded.control_id(), 0x0071);
    assert_eq!(decoded.dimensions(), Some(Dimensions::new(10, 20)));
    assert_eq!(decoded.to_bytes(), wire);

    let no_dimensions = ok(ControlHeader::new(
        ControlType::Label,
        1,
        ControlFlags::default(),
        SpecificFlags::default(),
        0,
        None,
    ));
    assert_eq!(no_dimensions.to_bytes().len(), 11);
}

#[test]
fn control_header_matches_ms_oshared_example_bytes() {
    let header = ok(ControlHeader::new(
        ControlType::Button,
        0x0071,
        ControlFlags::default(),
        SpecificFlags::from_raw(0x00AA_0020),
        3,
        None,
    ));
    assert_eq!(
        header.to_bytes(),
        vec![
            0x03, 0x01, 0x00, 0x01, 0x71, 0x00, 0x20, 0x00, 0xAA, 0x00, 0x03
        ]
    );
}

#[test]
fn control_header_preserves_unknown_type_and_rejects_bad_construction() {
    let mut wire = vec![
        0x03, 0x01, 0x20, 0x05, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00,
    ];
    let (decoded, consumed) = ok(ControlHeader::parse_prefix(&wire));
    assert_eq!(consumed, wire.len());
    assert!(matches!(decoded.control_type(), ControlType::Unknown(5)));
    assert_eq!(decoded.to_bytes(), wire);

    assert!(
        ControlHeader::new(
            ControlType::Unknown(5),
            1,
            ControlFlags::default(),
            SpecificFlags::default(),
            0,
            None,
        )
        .is_err()
    );
    assert!(
        ControlHeader::new(
            ControlType::Button,
            1,
            ControlFlags::default(),
            SpecificFlags::default(),
            8,
            None,
        )
        .is_err()
    );
    assert!(
        ControlHeader::new(
            ControlType::Button,
            1,
            ControlFlags::default(),
            SpecificFlags::default(),
            0,
            Some(Dimensions::new(1, 1)),
        )
        .is_err()
    );

    wire[0] = 0x02;
    assert!(ControlHeader::parse(&wire).is_err());
}

#[test]
fn general_info_and_extra_data_roundtrip_without_allocating_decoded_strings() {
    let extra = ok(ExtraInfo::new(
        text(""),
        0,
        text("tag"),
        text("Macro"),
        text("argument"),
        MergeMode::Server,
        MenuMerge::None,
    ));
    let flags = GeneralFlags::default()
        .with_save_text(true)
        .with_save_misc_ui_strings(true)
        .with_save_misc_custom(true)
        .with_disabled(true);
    let general = ok(GeneralInfo::new(
        flags,
        Some(text("Caption")),
        Some(text("Description")),
        Some(text("Tooltip")),
        Some(extra),
    ));
    let value = ok(Data::new(general, vec![0xA5, 0x03, 0x01]));
    let bytes = value.to_bytes();
    let (decoded, consumed) = ok(Data::parse_prefix(&bytes, 3));

    assert_eq!(consumed, bytes.len());
    assert_eq!(decoded.general().flags().raw(), flags.raw());
    assert_eq!(decoded.general().custom_text().unwrap().text(), "Caption");
    assert_eq!(decoded.general().tooltip().unwrap().text(), "Tooltip");
    assert_eq!(
        decoded.general().extra().unwrap().on_action().text(),
        "Macro"
    );
    assert_eq!(
        decoded.general().extra().unwrap().merge(),
        MergeMode::Server
    );
    assert_eq!(
        decoded.general().extra().unwrap().menu_merge(),
        MenuMerge::None
    );
    assert_eq!(decoded.specific(), &[0xA5, 0x03, 0x01]);
    assert_eq!(decoded.to_bytes(), bytes);
    assert_eq!(decoded.into_owned().to_bytes(), bytes);
}

#[test]
fn general_info_rejects_fields_without_their_presence_flags() {
    assert!(
        GeneralInfo::new(
            GeneralFlags::default(),
            Some(text("unexpected")),
            None,
            None,
            None,
        )
        .is_err()
    );
    assert!(GeneralInfo::parse(&[GeneralFlags::default().with_save_text(true).raw()]).is_err());
}
