#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::raw::Tree;
use litchi_rtf::{LegacyDrawingPrimitive, Shape, ShapeType, write::Writer};

fn write(document: &Tree<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    Writer::new(&mut output).write_document(document).unwrap();
    output
}

#[test]
fn parses_real_libreoffice_legacy_shape_result_and_round_trips_canonically() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/fdo55504-1-min.rtf"
    );
    let producer = Tree::parse_bytes(source, litchi_rtf::read::OpenOptions::default()).unwrap();
    let result = producer
        .shapes()
        .iter()
        .find_map(|shape| shape.result.as_ref())
        .unwrap();
    assert_eq!(result.drawing.position, 0);
    assert!(matches!(
        result.drawing.primitive,
        LegacyDrawingPrimitive::Line { .. }
    ));

    let mut document =
        Tree::parse(r"{\rtf1 Body}", litchi_rtf::read::OpenOptions::default()).unwrap();
    let mut shape = Shape::new(ShapeType::Line);
    shape.result = Some(result.clone());
    document.set_background_shape(shape).unwrap();
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("{\\*\\shprslt{\\*\\do"));
    let reparsed = Tree::parse_bytes(&output, litchi_rtf::read::OpenOptions::default()).unwrap();
    assert_eq!(
        reparsed.background_shape().unwrap().result,
        document.background_shape().unwrap().result
    );
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn accepts_text_only_producer_result_without_leaking_or_misrepresenting_it() {
    let document = Tree::parse(
        r"{\rtf1 A{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 1}}}{\shprslt\par\pard fallback}}B}",
        litchi_rtf::read::OpenOptions::default(),
    )
    .unwrap();
    assert_eq!(document.text(), "AB");
    assert!(document.shapes()[0].result.is_none());
}

#[test]
fn rejects_hostile_shape_result_grammar() {
    for source in [
        r"{\rtf1{\*\shprslt{\*\do\dobxpage\dobypage\dodhgt0\dpline\dpptx0\dppty0\dpptx1\dppty1\dpx0\dpy0\dpxsize1\dpysize1}}}",
        r"{\rtf1{\shprslt text}}",
        r"{\rtf1{\shp{\*\shprslt1{\*\do\dobxpage\dobypage\dodhgt0\dpline\dpptx0\dppty0\dpptx1\dppty1\dpx0\dpy0\dpxsize1\dpysize1}}}}",
        r"{\rtf1{\shp\shprslt text}}",
        r"{\rtf1{\shpgrp{\shp{\*\shprslt{\*\do\dobxpage\dobypage\dodhgt0\dpline\dpptx0\dppty0\dpptx1\dppty1\dpx0\dpy0\dpxsize1\dpysize1}}}}}",
        r"{\rtf1{\shp{\*\shprslt{\*\do\dobxpage\dobypage\dodhgt0\dpline\dpptx0\dppty0\dpptx1\dppty1\dpx0\dpy0\dpxsize1\dpysize1}}{\*\shprslt text}}}",
        r"{\rtf1{\shp{\*\shprslt{\do\dobxpage\dobypage\dodhgt0\dpline\dpptx0\dppty0\dpptx1\dppty1\dpx0\dpy0\dpxsize1\dpysize1}}}}",
        r"{\rtf1{\shp{\*\shprslt{\*\do\dobxpage\dobypage\dodhgt0\dpline\dpptx0\dppty0\dpptx1\dppty1\dpx0\dpy0\dpxsize1\dpysize1} trailing}}}",
        r"{\rtf1{\shp{\*\shprslt{\*\do\dobxpage\dobypage\dodhgt0\dpline\dpptx0\dppty0\dpptx1\dppty1\dpx0\dpy0\dpxsize1\dpysize1}}{\sp{\sn x}{\sv 1}}}}",
    ] {
        assert!(
            Tree::parse(source, litchi_rtf::read::OpenOptions::default()).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn enforces_typed_nested_position_invariant() {
    let source = r"{\rtf1{\shp{\*\shprslt{\*\do\dobxpage\dobypage\dodhgt0\dpline\dpptx0\dppty0\dpptx1\dppty1\dpx0\dpy0\dpxsize1\dpysize1}}}}";
    let document = Tree::parse(source, litchi_rtf::read::OpenOptions::default()).unwrap();
    let mut result = document.shapes()[0].result.clone().unwrap();
    result.drawing.position = 1;
    assert!(result.validate().is_err());

    let mut target = Tree::parse(r"{\rtf1}", litchi_rtf::read::OpenOptions::default()).unwrap();
    let mut shape = Shape::new(ShapeType::Line);
    shape.result = Some(result);
    assert!(target.set_background_shape(shape).is_err());
}
