#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use std::borrow::Cow;

use litchi_rtf::raw::Tree;
use litchi_rtf::{
    Formatting, HeaderFooter, HeaderFooterParagraph, HeaderFooterType, Paragraph, Shape,
    ShapeGroup, ShapeProperty, ShapeType, write::Writer,
};

fn write_document(document: &Tree<'_>) -> String {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    String::from_utf8(output).unwrap()
}

#[test]
fn libreoffice_root_groups_keep_body_position_and_round_trip() {
    for source in [
        include_str!(
            "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/fdo89496.rtf"
        ),
        include_str!(
            "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/tdf127806.rtf"
        ),
    ] {
        let document = Tree::parse(source, litchi_rtf::read::OpenOptions::default()).unwrap();
        assert_eq!(document.shape_groups().len(), 1);
        assert_eq!(document.shape_groups()[0].position, 0);
        let reparsed = Tree::parse(
            &write_document(&document),
            litchi_rtf::read::OpenOptions::default(),
        )
        .unwrap();
        assert_eq!(reparsed.shape_groups().len(), 1);
        assert_eq!(reparsed.shape_groups()[0].position, 0);
    }
}

#[test]
fn typed_body_group_round_trips_at_unicode_boundary() {
    let mut document = Tree::parse(
        r"{\rtf1 A\u20320?B}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let mut group = ShapeGroup::new();
    group.position = "A你".len();
    group.properties.push(ShapeProperty::new(
        Cow::Borrowed("wzName"),
        Cow::Borrowed("body group"),
    ));
    group.add_shape(Shape::new(ShapeType::Rectangle)).unwrap();
    document.push_shape_group(group).unwrap();
    let reparsed = Tree::parse(
        &write_document(&document),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(reparsed.text(), "A你B");
    assert_eq!(reparsed.shape_groups()[0].position, "A你".len());
}

#[test]
fn parser_and_writer_keep_header_drawings_in_the_header_story() {
    let source = r"{\rtf1{\header A{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 1}}}}{\shpgrp{\*\shpinst{\sp{\sn wzName}{\sv header group}}}}B}Body}";
    let document = Tree::parse(source, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert!(document.shapes().is_empty());
    assert!(document.shape_groups().is_empty());
    let header = &document.sections()[0].headers_footers[0];
    assert_eq!(header.shapes.len(), 1);
    assert_eq!(header.shape_groups.len(), 1);
    assert_eq!(header.shapes[0].position, 1);
    assert_eq!(header.shape_groups[0].position, 1);

    let reparsed = Tree::parse(
        &write_document(&document),
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    let header = &reparsed.sections()[0].headers_footers[0];
    assert_eq!(header.shapes.len(), 1);
    assert_eq!(header.shape_groups.len(), 1);
    assert!(reparsed.shapes().is_empty());
    assert!(reparsed.shape_groups().is_empty());
}

#[test]
fn header_footer_api_validates_utf8_positions_and_nested_position_abuse() {
    let mut header = HeaderFooter::new(HeaderFooterType::Header);
    header.add_paragraph(HeaderFooterParagraph::new(
        Cow::Borrowed("A你B"),
        Formatting::default(),
        Paragraph::default(),
    ));
    let mut split = ShapeGroup::new();
    split.position = 2;
    assert!(header.push_shape_group(split).is_err());

    let mut nested = ShapeGroup::new();
    nested.position = 1;
    let mut root = ShapeGroup::new();
    root.groups.push(nested);
    root.child_order.push(litchi_rtf::ShapeGroupChild::Group(0));
    assert!(header.push_shape_group(root).is_err());

    let mut valid = ShapeGroup::new();
    valid.position = "A你".len();
    header.push_shape_group(valid).unwrap();
    let mut output = Vec::new();
    Writer::new(&mut output)
        .write_header_footer(&header)
        .unwrap();
    let wrapped = format!("{{\\rtf1{}Body}}", String::from_utf8(output).unwrap());
    let reparsed = Tree::parse(&wrapped, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(
        reparsed.sections()[0].headers_footers[0].shape_groups[0].position,
        "A你".len()
    );
}
