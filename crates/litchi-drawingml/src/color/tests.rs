//! Focused conformance tests for typed DrawingML colors.

use std::mem::size_of;

use super::{Rgb, Scheme, Value, codec};

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
fn unsupported_choices_and_transforms_are_retained_without_becoming_unchecked() {
    for xml in [
        br#"<a:sysClr val="windowText" lastClr="000000"/>"#.as_slice(),
        br#"<a:srgbClr val="112233"><a:alpha val="50000"/></a:srgbClr>"#.as_slice(),
        br#"<a:schemeClr val="futureAccent"/>"#.as_slice(),
    ] {
        let value = codec::read(xml).unwrap();
        assert!(value.is_unknown());
        assert_eq!(codec::write(&value).unwrap(), xml);
    }

    assert!(codec::read(br#"<a:srgbClr val="GG0000"/>"#).is_err());
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
}
