#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use std::borrow::Cow;

use litchi_rtf::{RtfDocument, RtfWriter, Shape, ShapeProperty, ShapeType};

fn write_document(document: &RtfDocument<'_>) -> String {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    String::from_utf8(output).unwrap()
}

#[test]
fn parses_root_shape_producers_at_start_and_after_visible_body() {
    let producers = [
        include_str!("../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/relsize.rtf"),
        include_str!(
            "../../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/fdo55504-1-min.rtf"
        ),
        include_str!(
            "../../../test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf141173_missingFrames.rtf"
        ),
    ];
    let documents: Vec<_> = producers
        .iter()
        .map(|source| RtfDocument::parse(source).unwrap())
        .collect();
    assert!(
        documents
            .iter()
            .all(|document| !document.shapes().is_empty())
    );
    assert!(documents.iter().any(|document| {
        document
            .shapes()
            .iter()
            .any(|shape| !shape.instruction_present)
    }));
    assert!(
        documents
            .iter()
            .all(|document| document.shapes().iter().all(|shape| document
                .text()
                .get(shape.position..shape.position)
                .is_some()))
    );
}

#[test]
fn typed_api_inserts_at_utf8_boundary_and_clears_without_touching_background() {
    let mut document = RtfDocument::parse(r"{\rtf1 A\u20320?B}").unwrap();
    let mut shape = Shape::new(ShapeType::Rectangle);
    shape.position = "A你".len();
    shape.properties.push(ShapeProperty::new(
        Cow::Borrowed("wzName"),
        Cow::Borrowed("typed"),
    ));
    document.push_shape(shape).unwrap();
    let serialized = write_document(&document);
    assert!(serialized.contains("{\\shp{\\*\\shpinst"));
    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert_eq!(reparsed.shapes()[0].position, "A你".len());
    assert_eq!(reparsed.text(), "A你B");
    document.clear_shapes();
    assert!(document.shapes().is_empty());
}

#[test]
fn rejects_hostile_root_shape_grammar_and_resource_abuse() {
    let malformed = [
        r"{\rtf1{\shp1{\*\shpinst}}}",
        r"{\rtf1{\*\shp{\*\shpinst}}}",
        r"{\rtf1{\shp}}",
        r"{\rtf1{\shp{\shpinst}}}",
        r"{\rtf1{\shp{\*\shpinst}{\*\shpinst}}}",
        r"{\rtf1{\shp{\sp{\sn x}{\sv 1}}{\*\shpinst}}}",
        r"{\rtf1{\shp{\*\shpinst{\shptxt x}{\sp{\sn x}{\sv 1}}}}}",
        r"{\rtf1{\shp{\*\shpinst}direct}}",
        "{\\rtf1{\\shp{\\*\\shpinst\\bin1 x}}}",
    ];
    for source in malformed {
        assert!(RtfDocument::parse(source).is_err(), "accepted {source}");
    }

    let mut document = RtfDocument::parse(r"{\rtf1 x}").unwrap();
    let mut split = Shape::new(ShapeType::Rectangle);
    split.position = 2;
    assert!(document.push_shape(split).is_err());
    let mut huge = Shape::new(ShapeType::TextBox);
    huge.text = Cow::Owned("x".repeat(16 * 1_048_576 + 1));
    assert!(document.push_shape(huge).is_err());
}
