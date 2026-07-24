use std::borrow::Cow;

use litchi_rtf::{
    RtfDocument, RtfWriter, Shape, ShapeGeometry, ShapeGroup, ShapeGroupInfo,
    ShapeHorizontalAnchor, ShapeProperty, ShapeRotationDegrees, ShapeTwips, ShapeType,
    ShapeVerticalAnchor, ShapeWrapSide, ShapeWrapStyle, ShapeZOrder, StoryDrawing,
};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output).write_document(document).unwrap();
    output
}

fn named(name: &'static str, id: i32) -> Shape<'static> {
    let mut shape = Shape::new(ShapeType::Rectangle);
    shape.name = Cow::Borrowed(name);
    shape.info.push(ShapeGroupInfo::ShapeId(id));
    shape.properties.push(ShapeProperty::new(Cow::Borrowed("wzName"), Cow::Borrowed(name)));
    shape
}

#[test]
fn document_shape_find_replace_remove_and_reorder_are_atomic() {
    let mut document = RtfDocument::parse(r#"{\rtf1\ansi\ansicpg1252 Caf\'e9}"#).unwrap();
    let mut first = named("first", 41);
    first.position = document.text().len();
    let mut second = named("second", 42);
    second.position = document.text().len();
    assert_eq!(document.add_shape(first).unwrap(), 0);
    assert_eq!(document.add_shape(second).unwrap(), 1);
    assert_eq!(document.find_shape_by_name("second").unwrap().shape_id(), Some(42));
    assert_eq!(document.find_shape_by_id(41).unwrap().name, "first");

    document.move_drawing(1, 0).unwrap();
    assert_eq!(document.drawing_order(), &[StoryDrawing::Shape(1), StoryDrawing::Shape(0)]);
    let mut replacement = named("updated", 41);
    replacement.position = document.text().len();
    replacement.geometry.set_z_order(ShapeZOrder::new(9));
    document.replace_shape(0, replacement).unwrap();
    assert_eq!(document.find_shape_by_name("updated").unwrap().geometry.z_order_value(), ShapeZOrder::new(9));

    let before: Vec<_> = document.shapes().iter().cloned().map(Shape::into_owned).collect();
    let mut relocated = named("bad", 9);
    relocated.position = 0;
    assert!(document.replace_shape(0, relocated).is_err());
    assert_eq!(document.shapes(), before);
    document.remove_shape(1).unwrap();
    assert_eq!(document.drawing_order(), &[StoryDrawing::Shape(0)]);

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), "Café");
    assert!(reparsed.find_shape_by_name("updated").is_some());
}

#[test]
fn typed_units_anchors_and_nested_group_mutation_round_trip() {
    let mut root = ShapeGroup::new();
    root.geometry = ShapeGeometry::from_twips(ShapeTwips::new(10), ShapeTwips::new(20), ShapeTwips::new(300), ShapeTwips::new(400));
    root.geometry.set_rotation_degrees(ShapeRotationDegrees::new(15));
    root.info.extend([
        ShapeGroupInfo::horizontal_anchor(ShapeHorizontalAnchor::Margin),
        ShapeGroupInfo::vertical_anchor(ShapeVerticalAnchor::Paragraph),
        ShapeGroupInfo::wrap(ShapeWrapStyle::Tight),
        ShapeGroupInfo::wrap_side(ShapeWrapSide::Right),
    ]);
    root.add_shape(named("one", 1));
    root.add_shape(named("two", 2));
    root.move_child(1, 0).unwrap();
    root.replace_shape(0, named("replacement", 3)).unwrap();
    root.remove_shape(1).unwrap();
    assert_eq!(root.find_shape_by_id(3).unwrap().name, "replacement");
    assert_eq!(root.geometry.x_twips(), ShapeTwips::new(10));
    assert_eq!(root.info[0].as_horizontal_anchor(), Some(ShapeHorizontalAnchor::Margin));
    assert_eq!(root.info[1].as_vertical_anchor(), Some(ShapeVerticalAnchor::Paragraph));
    assert_eq!(root.info[2].as_wrap(), Some(ShapeWrapStyle::Tight));
    assert_eq!(root.info[3].as_wrap_side(), Some(ShapeWrapSide::Right));

    let mut document = RtfDocument::parse(r#"{\rtf1 body}"#).unwrap();
    root.position = document.text().len();
    document.add_shape_group(root).unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.shape_groups()[0].find_shape_by_name("replacement").unwrap().shape_id(), Some(3));
}

#[test]
fn libreoffice_fixture_mutates_without_interpreting_picture_or_ole_payloads() {
    let source = include_bytes!("../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/shpz-dhgt.rtf");
    let mut document = RtfDocument::parse_bytes(source).unwrap();
    let index = document.shapes().iter().position(|shape| shape.shape_id().is_some()).unwrap();
    let mut replacement = document.shapes()[index].clone().into_owned();
    replacement.set_property(ShapeProperty::new(Cow::Borrowed("litchiUnknownProperty"), Cow::Borrowed("opaque"))).unwrap();
    document.replace_shape(index, replacement).unwrap();
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.shapes()[index].property("litchiUnknownProperty"), Some("opaque"));
}

#[test]
fn reorder_across_body_anchors_and_invalid_group_indices_roll_back() {
    let mut document = RtfDocument::parse(r#"{\rtf1 AB}"#).unwrap();
    let mut first = named("a", 1);
    first.position = 0;
    let mut second = named("b", 2);
    second.position = 1;
    document.add_shape(first).unwrap();
    document.add_shape(second).unwrap();
    let before = document.drawing_order().to_vec();
    assert!(document.move_drawing(0, 1).is_err());
    assert_eq!(document.drawing_order(), before);

    let mut group = ShapeGroup::new();
    group.add_shape(named("only", 7));
    let before = group.clone();
    assert!(group.remove_shape(9).is_err());
    assert!(group.move_child(0, 9).is_err());
    assert_eq!(group, before);
}
