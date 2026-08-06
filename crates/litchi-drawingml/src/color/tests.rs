//! Focused conformance tests for typed DrawingML colors.

use std::mem::size_of;

use super::{
    Angle, FixedPercentage, Hsl, Percentage, PositiveFixedPercentage, PositivePercentage, Preset,
    Rgb, ScRgb, Scheme, System, Transform, Value, codec,
};

#[test]
fn rgb_values_are_checked_compact_and_canonical() {
    let value = Value::rgb(Rgb::new(0x12, 0xAB, 0xF0));
    assert_eq!(value.as_rgb().unwrap().channels(), [0x12, 0xAB, 0xF0]);
    assert_eq!(value.as_rgb().unwrap().to_hex(), "12ABF0");
    assert_eq!(size_of::<Rgb>(), 3);

    let xml = codec::write(&value).unwrap();
    assert_eq!(xml, br#"<a:srgbClr val="12ABF0"/>"#);
    assert_eq!(codec::read(&xml).unwrap(), value);
    assert_eq!(
        codec::read(br#" <a:srgbClr val="12abf0"/> "#).unwrap(),
        value
    );
}

#[test]
fn scheme_values_cover_the_shared_closed_vocabulary() {
    for scheme in [
        Scheme::Background,
        Scheme::Text,
        Scheme::Background2,
        Scheme::Text2,
        Scheme::Accent1,
        Scheme::Accent2,
        Scheme::Accent3,
        Scheme::Accent4,
        Scheme::Accent5,
        Scheme::Accent6,
        Scheme::Hyperlink,
        Scheme::FollowedHyperlink,
        Scheme::Dark1,
        Scheme::Light1,
        Scheme::Dark2,
        Scheme::Light2,
        Scheme::Placeholder,
    ] {
        let value = Value::scheme(scheme);
        let xml = codec::write(&value).unwrap();
        assert_eq!(codec::read(&xml).unwrap(), value);
        assert_eq!(Scheme::from_token(scheme.token()), Some(scheme));
    }
}

#[test]
fn typed_color_choices_and_transforms_round_trip() {
    let choices = [
        Value::scrgb(
            ScRgb::new(100_000, 50_000, 0).expect("scRGB channels within the schema bounds"),
        ),
        Value::hsl(Hsl::new(21_600_000, 50_000, 25_000).unwrap()),
        Value::system(System::new("windowText", Some(Rgb::new(0, 0, 0))).unwrap()),
        Value::preset(Preset::new("transparent").unwrap()),
    ];
    for value in choices {
        let xml = codec::write(&value).unwrap();
        assert_eq!(codec::read(&xml).unwrap(), value);
    }

    let xml = br#"<a:srgbClr val="112233">
        <a:alpha val="50%"/><a:tint val="25000"/>
        <a:shade val="10000"/><a:hueOff val="-60000"/>
    </a:srgbClr>"#;
    let value = codec::read(xml).unwrap();
    assert_eq!(
        value.as_transformed().unwrap().transforms(),
        &[
            Transform::Alpha(PositiveFixedPercentage::new(50_000).unwrap()),
            Transform::Tint(PositiveFixedPercentage::new(25_000).unwrap()),
            Transform::Shade(PositiveFixedPercentage::new(10_000).unwrap()),
            Transform::HueOff(Angle::new(-60_000)),
        ]
    );
    assert_eq!(
        codec::write(&value).unwrap(),
        br#"<a:srgbClr val="112233"><a:alpha val="50000"/><a:tint val="25000"/><a:shade val="10000"/><a:hueOff val="-60000"/></a:srgbClr>"#
    );
}

#[test]
fn every_typed_transform_round_trips_in_source_order() {
    let percentage = Percentage::new(10_000).unwrap();
    let positive = PositivePercentage::new(10_000).unwrap();
    let fixed = FixedPercentage::new(-10_000).unwrap();
    let positive_fixed = PositiveFixedPercentage::new(10_000).unwrap();
    let angle = Angle::new(-60_000);
    let positive_angle = super::PositiveAngle::new(60_000).unwrap();
    let transforms = [
        Transform::Alpha(positive_fixed),
        Transform::AlphaMod(positive),
        Transform::AlphaOff(fixed),
        Transform::Blue(percentage),
        Transform::BlueMod(percentage),
        Transform::BlueOff(percentage),
        Transform::Complement,
        Transform::Gamma,
        Transform::Gray,
        Transform::Green(percentage),
        Transform::GreenMod(percentage),
        Transform::GreenOff(percentage),
        Transform::Hue(positive_angle),
        Transform::HueMod(positive),
        Transform::HueOff(angle),
        Transform::Inverse,
        Transform::InverseGamma,
        Transform::Lum(percentage),
        Transform::LumMod(percentage),
        Transform::LumOff(percentage),
        Transform::Red(percentage),
        Transform::RedMod(percentage),
        Transform::RedOff(percentage),
        Transform::Sat(percentage),
        Transform::SatMod(percentage),
        Transform::SatOff(percentage),
        Transform::Shade(positive_fixed),
        Transform::Tint(positive_fixed),
    ];
    let value = Value::transformed(Rgb::new(0x12, 0x34, 0x56), transforms).unwrap();
    let xml = codec::write(&value).unwrap();
    assert_eq!(codec::read(&xml).unwrap(), value);
}

#[test]
fn unsupported_choices_and_attributes_are_retained_exactly() {
    for xml in [
        br#"<a:schemeClr val="futureAccent"/>"#.as_slice(),
        br#"<a:srgbClr extra="future" val="112233"/>"#.as_slice(),
        br#"<a:srgbClr val="112233"><a:futureTransform data="x"/></a:srgbClr>"#.as_slice(),
        br#"<x:futureClr xmlns:x="urn:future" second="2" first="1"><x:item/></x:futureClr>"#
            .as_slice(),
    ] {
        let value = codec::read(xml).unwrap();
        assert!(value.is_unknown());
        assert_eq!(codec::write(&value).unwrap(), xml);
    }

    assert!(codec::read(br#"<a:srgbClr val="GG0000"/>"#).is_err());
}

#[test]
fn scalar_and_transform_bounds_are_rejected() {
    assert!(ScRgb::new(100_001, 0, 0).is_err());
    assert!(Hsl::new(21_600_001, 0, 0).is_err());
    assert!(Percentage::new(100_001).is_err());
    assert!(PositivePercentage::new(100_001).is_err());
    assert!(codec::read(br#"<a:scrgbClr r="100001" g="0" b="0"/>"#).is_err());
    assert!(codec::read(br#"<a:hslClr hue="21600001" sat="0" lum="0"/>"#).is_err());
    assert!(
        codec::read(br#"<a:srgbClr val="000000"><a:alpha val="100001"/></a:srgbClr>"#).is_err()
    );
}

#[test]
fn malformed_fragments_and_resource_exhaustion_are_rejected() {
    for xml in [
        br#""#.as_slice(),
        br#"<a:srgbClr val="000000"><a:alpha/></a:srgbClr"#.as_slice(),
        br#"<a:srgbClr val="000000"/><a:schemeClr val="accent1"/>"#.as_slice(),
    ] {
        assert!(codec::read(xml).is_err(), "accepted {xml:?}");
    }
    let oversized = vec![b'x'; codec::MAX_XML_BYTES + 1];
    assert!(matches!(
        codec::read(&oversized),
        Err(crate::Error::Limit { .. })
    ));

    let mut transforms = String::from(r#"<a:srgbClr val="000000">"#);
    for _ in 0..=codec::MAX_TRANSFORMS {
        transforms.push_str(r#"<a:alpha val="0"/>"#);
    }
    transforms.push_str("</a:srgbClr>");
    assert!(matches!(
        codec::read(transforms.as_bytes()),
        Err(crate::Error::Limit {
            resource: "DrawingML color transforms",
            limit: codec::MAX_TRANSFORMS,
        })
    ));
}
