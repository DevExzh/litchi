#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use std::borrow::Cow;

use litchi_rtf::{
    RtfDocument, RtfWriter, Shape, ShapeGroup, ShapeGroupChild, ShapeProperty, ShapeType,
};

fn write_document(document: &RtfDocument<'_>) -> String {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    String::from_utf8(output).unwrap()
}

#[test]
fn parses_libreoffice_group_producers_and_round_trips_order() {
    let fdo = RtfDocument::parse(include_str!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/fdo89496.rtf"
    ))
    .unwrap();
    assert_eq!(fdo.shape_groups().len(), 1);
    assert!(!fdo.shape_groups()[0].shapes().is_empty());

    let source = include_str!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf127806.rtf"
    );
    let document = RtfDocument::parse(source).unwrap();
    let group = &document.shape_groups()[0];
    assert_eq!(
        group.child_order().len(),
        group.shapes().len() + group.groups().len()
    );
    let serialized = write_document(&document);
    assert!(serialized.contains("{\\shpgrp{\\*\\shpinst"));
    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert_eq!(reparsed.shape_groups().len(), 1);
    assert_eq!(
        reparsed.shape_groups()[0].child_order(),
        group.child_order()
    );
}

#[test]
fn typed_group_writes_nested_children() {
    let mut root = ShapeGroup::new();
    root.properties.push(ShapeProperty::new(
        Cow::Borrowed("wzName"),
        Cow::Borrowed("root"),
    ));
    root.add_shape(Shape::new(ShapeType::Rectangle)).unwrap();
    let mut nested = ShapeGroup::new();
    nested.properties.push(ShapeProperty::new(
        Cow::Borrowed("wzName"),
        Cow::Borrowed("nested"),
    ));
    nested.add_shape(Shape::new(ShapeType::Ellipse)).unwrap();
    root.add_group(nested).unwrap();
    assert_eq!(
        root.child_order(),
        &[ShapeGroupChild::Shape(0), ShapeGroupChild::Group(0)]
    );

    let mut document = RtfDocument::parse(r"{\rtf1}").unwrap();
    document.push_shape_group(root).unwrap();
    let serialized = write_document(&document);
    let reparsed = RtfDocument::parse(&serialized).unwrap();
    let group = &reparsed.shape_groups()[0];
    assert_eq!(
        group.child_order(),
        &[ShapeGroupChild::Shape(0), ShapeGroupChild::Group(0)]
    );
    assert_eq!(group.groups()[0].shapes()[0].shape_type, ShapeType::Ellipse);
}

#[test]
fn rejects_hostile_group_grammar_and_typed_order_abuse() {
    let malformed = [
        r"{\rtf1{\shpgrp1{\*\shpinst}}}",
        r"{\rtf1{\*\shpgrp{\*\shpinst}}}",
        r"{\rtf1{\shpgrp}}",
        r"{\rtf1{\shpgrp{\shpinst}}}",
        r"{\rtf1{\shpgrp{\*\shpinst}{\*\shpinst}}}",
        r"{\rtf1{\shpgrp{\sp{\sn x}{\sv 1}}{\*\shpinst}}}",
        r"{\rtf1{\shpgrp{\*\shpinst{\shpgrp{\*\shpinst}{\*\shprslt{\*\do}}}}}}",
        "{\\rtf1{\\shpgrp{\\*\\shpinst\\bin1 x}}}",
    ];
    for source in malformed {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    let mut group = ShapeGroup::new();
    group.add_shape(Shape::new(ShapeType::Rectangle)).unwrap();
    group.child_order.push(ShapeGroupChild::Shape(0));
    let mut document = RtfDocument::parse(r"{\rtf1}").unwrap();
    assert!(document.push_shape_group(group).is_err());
}
