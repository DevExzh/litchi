//! Differential-format semantic and wire regression tests.

use super::codec;
use super::*;

fn color() -> XfColor {
    XfColor::try_new(
        XfColorSource::Theme(ThemeColor::Accent2),
        100,
        [1, 2, 3, 255],
    )
    .unwrap()
}

#[test]
fn typed_dxf_round_trips_representative_property_families() {
    let dxf = DifferentialFormat::try_new(
        true,
        vec![
            XfProperty::Gradient(XfGradient::linear(45.0).unwrap()),
            XfProperty::GradientStop(XfGradientStop::try_new(0.5, color()).unwrap()),
            XfProperty::TopBorder(XfBorder::new(color(), BorderStyle::Thin)),
            XfProperty::VerticalBorder(XfBorder::new(color(), BorderStyle::Dashed)),
            XfProperty::HorizontalAlignment(Some(HorizontalAlignment::Distributed)),
            XfProperty::JustifyDistributed(true),
            XfProperty::TextRotation(TextRotation::Clockwise(30)),
            XfProperty::FontName("Aptos".to_string()),
            XfProperty::FontWeight(XfFontWeight::Bold),
            XfProperty::FontUnderline(FontUnderline::Double),
            XfProperty::FontSizeTwips(220),
            XfProperty::NumberFormatCode("0.00".to_string()),
            XfProperty::RelativeIndent(Some(-2)),
            XfProperty::Locked(true),
        ],
    )
    .unwrap();
    let payload = dxf.to_payload().unwrap();
    assert_eq!(DifferentialFormat::parse_payload(&payload).unwrap(), dxf);
    let record = dxf.to_record_bytes().unwrap();
    assert_eq!(&record[..2], &[0x8D, 0x08]);
}

#[test]
fn xfprop_color_uses_low_flag_bit_and_high_seven_type_bits() {
    // Apache POI producer forms: bit 7 is clear, while fValidRGBA in bit 0 is set.
    let rgb = [0x05, 0xFF, 0x00, 0x00, 0xFF, 0xC7, 0xCE, 0xFF];
    let parsed = XfColor::parse(&rgb).unwrap();
    assert_eq!(parsed.source(), XfColorSource::Rgb);
    let mut encoded = Vec::new();
    parsed.write_to(&mut encoded);
    assert_eq!(encoded, rgb);

    let theme = [0x07, 0x04, 0x65, 0x66, 0xDC, 0xE6, 0xF1, 0xFF];
    let parsed = XfColor::parse(&theme).unwrap();
    assert_eq!(parsed.source(), XfColorSource::Theme(ThemeColor::Accent1));
    let mut encoded = Vec::new();
    parsed.write_to(&mut encoded);
    assert_eq!(encoded, theme);

    let indexed = [0x03, 0x40, 0, 0, 1, 2, 3, 4];
    assert_eq!(
        XfColor::parse(&indexed).unwrap().source(),
        XfColorSource::Indexed(0x40)
    );
    let automatic = [0x01, 0xAA, 0, 0, 1, 2, 3, 4];
    assert_eq!(
        XfColor::parse(&automatic).unwrap().source(),
        XfColorSource::Automatic
    );
    let not_set = [0x09, 0xAA, 0, 0, 1, 2, 3, 4];
    assert_eq!(
        XfColor::parse(&not_set).unwrap().source(),
        XfColorSource::NotSet
    );
}

#[test]
fn xfprop_color_rejects_clear_valid_flag_and_invalid_type_data() {
    assert!(XfColor::parse(&[0x04, 0, 0, 0, 0, 0, 0, 0]).is_err());
    assert!(XfColor::parse(&[0x0B, 0, 0, 0, 0, 0, 0, 0]).is_err());
    assert!(XfColor::parse(&[0x03, 66, 0, 0, 0, 0, 0, 0]).is_err());
    assert!(XfColor::parse(&[0x07, 12, 0, 0, 0, 0, 0, 0]).is_err());
    assert!(XfColor::parse(&[0x05, 0, 0x00, 0x80, 0, 0, 0, 0]).is_err());
}

#[test]
fn rejects_hostile_headers_sizes_flags_and_property_relationships() {
    let empty = DifferentialFormat::try_new(false, vec![])
        .unwrap()
        .to_payload()
        .unwrap();
    assert!(DifferentialFormat::parse_payload(&empty[..17]).is_err());
    let mut bad = empty.clone();
    bad[0] = 0;
    assert!(DifferentialFormat::parse_payload(&bad).is_err());
    let mut bad = empty.clone();
    bad[14] = 1;
    assert!(DifferentialFormat::parse_payload(&bad).is_err());
    let mut bad = empty;
    bad[16..18].copy_from_slice(&1u16.to_le_bytes());
    assert!(DifferentialFormat::parse_payload(&bad).is_err());

    assert!(
        DifferentialFormat::try_new(
            false,
            vec![XfProperty::VerticalBorder(XfBorder::new(
                color(),
                BorderStyle::Thin,
            ))],
        )
        .is_err()
    );
    assert!(
        DifferentialFormat::try_new(
            false,
            vec![
                XfProperty::FillPattern(FillPattern::Solid),
                XfProperty::Gradient(XfGradient::linear(0.0).unwrap()),
            ],
        )
        .is_err()
    );
    assert!(
        DifferentialFormat::try_new(false, vec![XfProperty::JustifyDistributed(true)],).is_err()
    );
}

#[test]
fn fixed_width_reads_reject_offset_overflow_without_panicking() {
    assert!(matches!(
        std::panic::catch_unwind(|| read_u16(&[], usize::MAX, "u16")),
        Ok(Err(_))
    ));
    assert!(matches!(
        std::panic::catch_unwind(|| read_u32(&[], usize::MAX, "u32")),
        Ok(Err(_))
    ));
    assert!(matches!(
        std::panic::catch_unwind(|| read_f64(&[], usize::MAX, "f64")),
        Ok(Err(_))
    ));
}

#[test]
fn malformed_fixed_width_properties_return_errors_without_panicking() {
    let empty = DifferentialFormat::try_new(false, vec![])
        .unwrap()
        .to_payload()
        .unwrap();
    for property_type in [
        0x0000u16, 0x0001, 0x0003, 0x0004, 0x0006, 0x000D, 0x000F, 0x0010, 0x0011, 0x0012, 0x0013,
        0x0018, 0x0019, 0x001A, 0x001B, 0x0022, 0x0023, 0x0024, 0x0025, 0x0029, 0x002A,
    ] {
        let mut payload = empty.clone();
        payload[16..18].copy_from_slice(&1u16.to_le_bytes());
        payload.extend_from_slice(&property_type.to_le_bytes());
        payload.extend_from_slice(&4u16.to_le_bytes());
        let parsed = std::panic::catch_unwind(|| DifferentialFormat::parse_payload(&payload));
        assert!(
            matches!(parsed, Ok(Err(_))),
            "property type 0x{property_type:04X} did not reject empty data"
        );
    }
}

#[test]
fn oversized_dxf_writes_are_rejected_without_truncating_record_length() {
    let dxf =
        DifferentialFormat::try_new(false, vec![XfProperty::WrapText(false); MAX_XF_PROPERTIES])
            .unwrap();
    assert!(dxf.to_payload().is_err());
    assert!(dxf.to_record_bytes().is_err());
}

#[test]
fn enforces_resource_caps() {
    assert!(
        XfProperties::try_new(vec![XfProperty::WrapText(false); MAX_XF_PROPERTIES + 1]).is_err()
    );
    let huge = "x".repeat(256);
    assert!(DifferentialFormat::try_new(false, vec![XfProperty::NumberFormatCode(huge)]).is_err());
}

#[test]
fn number_format_code_matches_producer_wide_string_bytes() {
    let producer = [
        0x05, 0x00, 0x22, 0x00, 0x24, 0x00, 0x22, 0x00, 0x30, 0x00, 0x30, 0x00,
    ];
    assert_eq!(
        codec::parse_number_format_code(&producer).unwrap(),
        "\"$\"00"
    );

    let mut encoded = Vec::new();
    codec::write_number_format_code("\"$\"00", &mut encoded).unwrap();
    assert_eq!(encoded, producer);
}

#[test]
fn number_format_code_rejects_malformed_wide_strings() {
    assert!(codec::parse_number_format_code(&[0x00, 0x00]).is_err());
    assert!(codec::parse_number_format_code(&[0x00, 0x01]).is_err());
    assert!(codec::parse_number_format_code(&[0x02, 0x00, 0x30, 0x00]).is_err());
    assert!(codec::parse_number_format_code(&[0x01, 0x00, 0x30, 0x00, 0x00]).is_err());
    assert!(codec::parse_number_format_code(&[0x01, 0x00, 0x00, 0xD8]).is_err());

    // A BIFF XLUnicodeString flags byte is not part of this XFProp payload.
    assert!(codec::parse_number_format_code(&[0x01, 0x00, 0x01, 0x30, 0x00]).is_err());
}
