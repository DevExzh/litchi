#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{LegacyDrawingPrimitive, RtfDocument, RtfWriter, Shape, ShapeType};

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

#[test]
fn parses_real_libreoffice_legacy_shape_result_and_round_trips_canonically() {
    let source = include_bytes!(
        "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/fdo55504-1-min.rtf"
    );
    let producer = RtfDocument::parse_bytes(source).unwrap();
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

    let mut document = RtfDocument::parse(r"{\rtf1 Body}").unwrap();
    let mut shape = Shape::new(ShapeType::Line);
    shape.result = Some(result.clone());
    document.set_background_shape(shape).unwrap();
    let output = write(&document);
    let serialized = String::from_utf8(output.clone()).unwrap();
    assert!(serialized.contains("{\\*\\shprslt{\\*\\do"));
    let reparsed = RtfDocument::parse_bytes(&output).unwrap();
    assert_eq!(
        reparsed.background_shape().unwrap().result,
        document.background_shape().unwrap().result
    );
    assert_eq!(reparsed.text(), "Body");
}

#[test]
fn accepts_text_only_producer_result_without_leaking_or_misrepresenting_it() {
    let document = RtfDocument::parse(
        r"{\rtf1 A{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 1}}}{\shprslt\par\pard fallback}}B}",
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
            RtfDocument::parse(source).is_err(),
            "accepted malformed {source}"
        );
    }
}

#[test]
fn enforces_typed_nested_position_invariant() {
    let source = r"{\rtf1{\shp{\*\shprslt{\*\do\dobxpage\dobypage\dodhgt0\dpline\dpptx0\dppty0\dpptx1\dppty1\dpx0\dpy0\dpxsize1\dpysize1}}}}";
    let document = RtfDocument::parse(source).unwrap();
    let mut result = document.shapes()[0].result.clone().unwrap();
    result.drawing.position = 1;
    assert!(result.validate().is_err());

    let mut target = RtfDocument::parse(r"{\rtf1}").unwrap();
    let mut shape = Shape::new(ShapeType::Line);
    shape.result = Some(result);
    assert!(target.set_background_shape(shape).is_err());
}
