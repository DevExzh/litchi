//! Focused conformance tests for typed DrawingML transforms.

use super::{Angle, Point, Size, Snapshot, Transform, codec};

const TRANSFORM_XML: &[u8] = br#"<x:xfrm xmlns:x="http://schemas.openxmlformats.org/drawingml/2006/main" flipV="true" rot=" 60000 "><x:off y="2" x="1"/><x:ext cx="3" cy="4"/><x:chOff x="-5" y="6"/><x:chExt cx="7" cy="8"/></x:xfrm>"#;

#[test]
fn typed_transform_reads_all_shared_fields_and_round_trips() {
    let value = codec::read(TRANSFORM_XML).expect("transform fixture must parse");
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

    let encoded = codec::write(&value).expect("typed transform must serialize");
    assert_eq!(
        encoded,
        br#"<a:xfrm rot="60000" flipV="1"><a:off x="1" y="2"/><a:ext cx="3" cy="4"/><a:chOff x="-5" y="6"/><a:chExt cx="7" cy="8"/></a:xfrm>"#
    );
    assert_eq!(codec::read(&encoded).unwrap(), value);
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
        codec::write(&value).unwrap(),
        br#"<a:xfrm rot="0" flipH="0"><a:off x="10" y="-20"/><a:ext cx="300" cy="400"/></a:xfrm>"#
    );
}

#[test]
fn snapshots_replay_source_exactly_and_edits_are_atomic() {
    let snapshot = Snapshot::from_xml(TRANSFORM_XML).expect("source snapshot must parse");
    assert_eq!(snapshot.xml_bytes(), TRANSFORM_XML);

    let unchanged = snapshot.edit().commit().unwrap();
    assert_eq!(unchanged.snapshot().xml_bytes(), TRANSFORM_XML);
    assert!(!snapshot.edit().is_changed());

    let mut edit = snapshot.edit();
    edit.set_rotation(Some(Angle::new(-120_000)))
        .set_flip_horizontal(Some(true));
    assert!(edit.is_changed());
    let committed = edit.commit().expect("valid transform edit must commit");
    assert_eq!(
        committed.snapshot().value().rotation(),
        Angle::new(-120_000)
    );
    assert!(committed.snapshot().value().flip_horizontal());
    assert_eq!(snapshot.value().rotation(), Angle::new(60_000));

    let restored = committed
        .patch()
        .clone()
        .inverse()
        .apply(committed.snapshot())
        .expect("inverse patch must apply to its commit");
    assert_eq!(restored.value(), snapshot.value());
}

#[test]
fn scalar_and_structure_bounds_are_strict() {
    assert!(codec::read(br#"<a:xfrm rot="2147483648"/>"#).is_err());
    assert!(codec::read(br#"<a:xfrm flipH="maybe"/>"#).is_err());
    assert!(codec::read(br#"<a:xfrm><a:off x="0"/></a:xfrm>"#).is_err());
    assert!(codec::read(br#"<a:xfrm><a:ext cx="-1" cy="0"/></a:xfrm>"#).is_err());
    assert!(
        codec::read(br#"<a:xfrm><a:ext cy="0" cx="1"/><a:off x="0" y="0"/></a:xfrm>"#).is_err()
    );
    assert!(codec::read(br#"<a:xfrm><a:off x="0" y="0"/><a:off x="1" y="1"/></a:xfrm>"#).is_err());
    assert!(codec::read(br#"<a:xfrm future="1"/>"#).is_err());
    assert!(codec::read(br#"<a:xfrm><a:future/></a:xfrm>"#).is_err());
    assert!(codec::read(br#"<a:xfrm><a:off x="0" y="0"><a:nested/></a:off></a:xfrm>"#).is_err());
}

#[test]
fn resource_limits_and_fragment_shape_are_checked() {
    let oversized = vec![b'x'; codec::MAX_XML_BYTES + 1];
    assert!(matches!(
        codec::read(&oversized),
        Err(crate::Error::Limit {
            resource: "DrawingML transform XML",
            limit: codec::MAX_XML_BYTES,
        })
    ));
    assert!(codec::read(br#"<a:xfrm/><a:xfrm/>"#).is_err());
    assert!(codec::read(br#"<a:xfrm"#).is_err());
    assert!(codec::read(br#""#).is_err());
}
