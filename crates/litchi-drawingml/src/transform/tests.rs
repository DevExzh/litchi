//! Focused conformance tests for typed DrawingML transforms.

use super::{Angle, Package, Point, Size, Transform, codec};
use crate::Error;

const TRANSFORM_XML: &[u8] = br#"<x:xfrm xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/main" flipV="true" rot=" 60000 "><x:off y="2" x="1"/><x:ext cx="3" cy="4"/><x:chOff x="-5" y="6"/><x:chExt cx="7" cy="8"/></x:xfrm>"#;

#[test]
fn typed_transform_reads_all_shared_fields_and_round_trips() {
    let value =
        codec::read(TRANSFORM_XML, &codec::Limits::DEFAULT).expect("transform fixture must parse");
    assert_eq!(value.rotation(), Angle::new(60_000));
    assert_eq!(value.authored_rotation(), Some(Angle::new(60_000)));
    assert!(!value.flip_horizontal());
    assert!(value.flip_vertical());
    assert_eq!(value.authored_flip_horizontal(), None);
    assert_eq!(value.authored_flip_vertical(), Some(true));
    assert_eq!(value.offset().unwrap().x().as_emu(), Some(1));
    assert_eq!(value.offset().unwrap().y().as_emu(), Some(2));
    assert_eq!(value.extent().unwrap().width().as_emu(), 3);
    assert_eq!(value.extent().unwrap().height().as_emu(), 4);
    assert_eq!(value.child_offset().unwrap().x().as_emu(), Some(-5));
    assert_eq!(value.child_offset().unwrap().y().as_emu(), Some(6));
    assert_eq!(value.child_extent().unwrap().width().as_emu(), 7);
    assert_eq!(value.child_extent().unwrap().height().as_emu(), 8);

    let encoded =
        codec::write(&value, &codec::Limits::DEFAULT).expect("typed transform must serialize");
    assert_eq!(
        encoded,
        br#"<a:xfrm rot="60000" flipV="1"><a:off x="1" y="2"/><a:ext cx="3" cy="4"/><a:chOff x="-5" y="6"/><a:chExt cx="7" cy="8"/></a:xfrm>"#
    );
    assert_eq!(
        codec::read(&encoded, &codec::Limits::DEFAULT).unwrap(),
        value
    );
}

#[test]
fn detached_builder_keeps_authored_defaults_distinct() {
    let value = Transform::new()
        .with_offset(Point::emu(10, -20).unwrap())
        .with_extent(Size::emu(300, 400).unwrap())
        .with_rotation(Angle::ZERO)
        .with_flip_horizontal(false);
    assert_eq!(value.rotation(), Angle::ZERO);
    assert_eq!(value.authored_rotation(), Some(Angle::ZERO));
    assert_eq!(value.authored_flip_horizontal(), Some(false));
    assert_eq!(
        codec::write(&value, &codec::Limits::DEFAULT).unwrap(),
        br#"<a:xfrm rot="0" flipH="0"><a:off x="10" y="-20"/><a:ext cx="300" cy="400"/></a:xfrm>"#
    );
}

#[test]
fn snapshots_replay_source_exactly_and_edits_are_atomic() {
    let package = Package::from_xml(TRANSFORM_XML, &codec::Limits::DEFAULT)
        .expect("source package must parse");
    assert_eq!(package.xml_bytes(), TRANSFORM_XML);

    let unchanged = package.edit().commit().unwrap();
    assert_eq!(unchanged.package().xml_bytes(), TRANSFORM_XML);
    assert!(!unchanged.diagnostics().changed());

    let mut edit = package.edit();
    edit.set_rotation(Some(Angle::new(-120_000)))
        .set_flip_horizontal(Some(true));
    let committed = edit.commit().expect("valid transform edit must commit");
    assert_eq!(committed.package().value().rotation(), Angle::new(-120_000));
    assert!(committed.package().value().flip_horizontal());
    assert!(committed.diagnostics().changed());
    assert_eq!(package.value().rotation(), Angle::new(60_000));
    let replayed = package
        .apply(committed.patch())
        .expect("package replay must publish a commit");
    assert!(replayed.diagnostics().changed());
    assert_eq!(replayed.package().value(), committed.package().value());

    let inverse = committed.patch().clone().inverse();
    let restored = committed
        .package()
        .apply(&inverse)
        .expect("inverse patch must apply to its commit");
    assert_eq!(restored.package().value(), package.value());

    let other_xml = std::str::from_utf8(TRANSFORM_XML)
        .expect("transform fixture is UTF-8")
        .replace("60000", "60001");
    let other =
        Package::from_xml(other_xml, &codec::Limits::DEFAULT).expect("other transform source");
    assert!(matches!(
        other.apply(committed.patch()),
        Err(Error::PatchConflict)
    ));
}

#[test]
fn scalar_and_structure_bounds_are_strict() {
    assert!(codec::read(br#"<a:xfrm rot="2147483648"/>"#, &codec::Limits::DEFAULT).is_err());
    assert!(codec::read(br#"<a:xfrm flipH="maybe"/>"#, &codec::Limits::DEFAULT).is_err());
    assert!(
        codec::read(
            br#"<a:xfrm><a:off x="0"/></a:xfrm>"#,
            &codec::Limits::DEFAULT
        )
        .is_err()
    );
    assert!(
        codec::read(
            br#"<a:xfrm><a:ext cx="-1" cy="0"/></a:xfrm>"#,
            &codec::Limits::DEFAULT
        )
        .is_err()
    );
    assert!(
        codec::read(
            br#"<a:xfrm><a:ext cy="0" cx="1"/><a:off x="0" y="0"/></a:xfrm>"#,
            &codec::Limits::DEFAULT
        )
        .is_err()
    );
    assert!(
        codec::read(
            br#"<a:xfrm><a:off x="0" y="0"/><a:off x="1" y="1"/></a:xfrm>"#,
            &codec::Limits::DEFAULT
        )
        .is_err()
    );
    assert!(codec::read(br#"<a:xfrm future="1"/>"#, &codec::Limits::DEFAULT).is_err());
    assert!(codec::read(br#"<a:xfrm><a:future/></a:xfrm>"#, &codec::Limits::DEFAULT).is_err());
    assert!(
        codec::read(
            br#"<a:xfrm><a:off x="0" y="0"><a:nested/></a:off></a:xfrm>"#,
            &codec::Limits::DEFAULT
        )
        .is_err()
    );
}

#[test]
fn resource_limits_and_fragment_shape_are_checked() {
    let oversized = vec![b'x'; codec::MAX_XML_BYTES + 1];
    assert!(matches!(
        codec::read(&oversized, &codec::Limits::DEFAULT),
        Err(Error::Limit {
            kind: crate::LimitKind::XmlBytes,
            maximum: codec::MAX_XML_BYTES,
            ..
        })
    ));
    assert!(codec::read(br#"<a:xfrm/><a:xfrm/>"#, &codec::Limits::DEFAULT).is_err());
    assert!(codec::read(br#"<a:xfrm"#, &codec::Limits::DEFAULT).is_err());
    assert!(codec::read(br#""#, &codec::Limits::DEFAULT).is_err());
}

#[test]
fn overlong_coordinate_lexicals_keep_the_typed_budget_dimension() {
    let lexical = "1".repeat(crate::coordinate::MAX_BYTES + 1);
    let xml = format!(r#"<a:xfrm><a:off x="{lexical}" y="0"/></a:xfrm>"#);
    assert!(matches!(
        codec::read(xml.as_bytes(), &codec::Limits::DEFAULT),
        Err(Error::Limit {
            kind: crate::LimitKind::LexicalValueBytes,
            observed,
            maximum,
        }) if observed == crate::coordinate::MAX_BYTES + 1 && maximum == crate::coordinate::MAX_BYTES
    ));
}

#[test]
fn transform_errors_do_not_echo_authored_xml() {
    for xml in [
        r#"<a:xfrm rot="private-value"/>"#,
        r#"<a:xfrm flipV="private-value"/>"#,
        r#"<a:xfrm private-value="1"/>"#,
        r#"<a:xfrm><a:private-value/></a:xfrm>"#,
        r#"<a:xfrm rot="&private-value;"/>"#,
        r#"<a:xfrm><a:off x="private-value" y="0"/></a:xfrm>"#,
        r#"<a:xfrm><a:ext cx="private-value" cy="0"/></a:xfrm>"#,
        r#"<a:xfrm></private-value>"#,
    ] {
        let error =
            codec::read(xml.as_bytes(), &codec::Limits::DEFAULT).expect_err("invalid transform");
        assert!(!format!("{error:?} {error}").contains("private-value"));
    }
    let error = Angle::parse("private-value").unwrap_err();
    assert!(!format!("{error:?} {error}").contains("private-value"));
}
