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
    Annotation, Note, RtfDocument, RtfWriter, Shape, ShapeGroup, ShapeType, StoryDrawing,
};

const ROOT_SHAPE: &str = r"{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 1}}}}";
const ROOT_GROUP: &str = r"{\shpgrp{\*\shpinst{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 1}}}}}}";

fn write(document: &RtfDocument<'_>) -> Vec<u8> {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    output
}

fn unicode_drawing_story() -> String {
    format!(r"a{ROOT_SHAPE}\u20320?{ROOT_GROUP}b")
}

fn equal_offset_group_then_shape_story() -> String {
    format!(r"a{ROOT_GROUP}{ROOT_SHAPE}b")
}

#[test]
fn preserves_equal_offset_cross_kind_order_in_every_story() {
    let story = equal_offset_group_then_shape_story();
    let source = [
        r"{\rtf1{\header H{\b ",
        &story,
        r"}Z}{\footer F{\i ",
        &story,
        r"}Z}",
        ROOT_GROUP,
        ROOT_SHAPE,
        r"X{\footnote ",
        &story,
        r"}Y{\endnote ",
        &story,
        r"}Z{\*\atnid I}{\*\atnauthor Ada}\chatn{\*\annotation ",
        &story,
        r"}{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}{\shptxt ",
        &story,
        r"}}}{\*\do\dptxbx{\dptxbxtext ",
        &story,
        r"}}Q}",
    ]
    .concat();
    let document = RtfDocument::parse(&source).unwrap();
    assert_eq!(
        document.drawing_order(),
        &[
            StoryDrawing::ShapeGroup(0),
            StoryDrawing::Shape(0),
            StoryDrawing::Shape(1)
        ]
    );
    for header_footer in &document.sections()[0].headers_footers {
        assert!(matches!(header_footer.text().as_str(), "HabZ" | "FabZ"));
        assert_eq!(
            header_footer.drawing_order,
            [StoryDrawing::ShapeGroup(0), StoryDrawing::Shape(0)]
        );
        assert_eq!(header_footer.shape_groups[0].position, 2);
        assert_eq!(header_footer.shapes[0].position, 2);
    }
    for note in document.notes() {
        assert_eq!(
            note.drawing_order,
            [StoryDrawing::ShapeGroup(0), StoryDrawing::Shape(0)]
        );
    }
    assert_eq!(
        document.annotations()[0].drawing_order,
        [StoryDrawing::ShapeGroup(0), StoryDrawing::Shape(0)]
    );
    assert_eq!(
        document.shapes()[1].text_drawing_order,
        [StoryDrawing::ShapeGroup(0), StoryDrawing::Shape(0)]
    );
    assert_eq!(
        document.legacy_text_boxes()[0].drawing_order,
        [StoryDrawing::ShapeGroup(0), StoryDrawing::Shape(0)]
    );

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.drawing_order(), document.drawing_order());
    assert_eq!(
        reparsed.sections()[0].headers_footers[0].drawing_order,
        document.sections()[0].headers_footers[0].drawing_order
    );
    assert_eq!(
        reparsed.notes()[0].drawing_order,
        document.notes()[0].drawing_order
    );
    assert_eq!(
        reparsed.annotations()[0].drawing_order,
        document.annotations()[0].drawing_order
    );
    assert_eq!(
        reparsed.shapes()[1].text_drawing_order,
        document.shapes()[1].text_drawing_order
    );
    assert_eq!(
        reparsed.legacy_text_boxes()[0].drawing_order,
        document.legacy_text_boxes()[0].drawing_order
    );
}

#[test]
fn root_drawings_round_trip_in_every_representable_story() {
    let story = unicode_drawing_story();
    let source = [
        r"{\rtf1 X{\footnote ",
        &story,
        r"}Y{\endnote ",
        &story,
        r"}Z{\*\atnid I}{\*\atnauthor Ada}\chatn{\*\annotation ",
        &story,
        r"}{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 202}}{\shptxt ",
        &story,
        r"}}}Q}",
    ]
    .concat();
    let document = RtfDocument::parse(&source).unwrap();
    assert_eq!(document.text(), "XYZQ");
    assert_eq!(document.shapes().len(), 1);
    assert_eq!(document.shape_groups().len(), 0);

    assert_eq!(document.notes().len(), 2);
    assert_eq!(document.notes()[0].position, 1);
    assert_eq!(document.notes()[1].position, 2);
    for note in document.notes() {
        assert_eq!(note.content, "a你b");
        assert_eq!(note.shapes[0].position, 1);
        assert_eq!(note.shape_groups[0].position, "a你".len());
    }

    let annotation = &document.annotations()[0];
    assert_eq!(annotation.text, "a你b");
    assert_eq!(annotation.shapes[0].position, 1);
    assert_eq!(annotation.shape_groups[0].position, "a你".len());

    let text_box = &document.shapes()[0];
    assert_eq!(text_box.text, "a你b");
    assert_eq!(text_box.text_shapes[0].position, 1);
    assert_eq!(text_box.text_shape_groups[0].position, "a你".len());

    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.text(), document.text());
    assert_eq!(reparsed.notes().len(), 2);
    assert_eq!(reparsed.notes()[0].shapes[0].position, 1);
    assert_eq!(reparsed.notes()[1].shape_groups[0].position, "a你".len());
    assert_eq!(reparsed.annotations(), document.annotations());
    assert_eq!(reparsed.shapes(), document.shapes());
}

#[test]
fn legacy_text_box_story_owns_root_drawings() {
    let story = unicode_drawing_story();
    let source = [r"{\rtf1 A{\*\do\dptxbx{\dptxbxtext ", &story, r"}}B}"].concat();
    let document = RtfDocument::parse(&source).unwrap();
    assert_eq!(document.text(), "AB");
    assert!(document.shapes().is_empty());
    assert!(document.shape_groups().is_empty());
    let text_box = &document.legacy_text_boxes()[0];
    assert_eq!(text_box.text, "a你b");
    assert_eq!(text_box.shapes[0].position, 1);
    assert_eq!(text_box.shape_groups[0].position, "a你".len());
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.legacy_text_boxes(), document.legacy_text_boxes());
}

#[test]
fn captures_note_drawings_after_text_inside_formatting_groups() {
    let source = format!(r"{{\rtf1 A{{\footnote pre{{\b mid{ROOT_SHAPE}tail}}post}}B}}");
    let document = RtfDocument::parse(&source).unwrap();
    let note = &document.notes()[0];
    assert_eq!(note.content, "premidtailpost");
    assert_eq!(note.shapes[0].position, "premid".len());
    assert!(document.shapes().is_empty());
    let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
    assert_eq!(reparsed.notes()[0].content, note.content);
    assert_eq!(reparsed.notes()[0].shapes[0].position, "premid".len());
}

#[test]
fn typed_story_apis_enforce_local_and_main_utf8_boundaries() {
    let mut shape = Shape::new(ShapeType::Rectangle);
    shape.position = 2;

    let mut note = Note::footnote(Cow::Borrowed("1"), Cow::Borrowed("A你B"));
    assert!(note.push_shape(shape.clone()).is_err());
    shape.position = "A你".len();
    note.push_shape(shape.clone()).unwrap();
    note.position = 2;

    let mut document = RtfDocument::parse(r"{\rtf1 A\u20320?B}").unwrap();
    assert!(document.push_note(note.clone()).is_err());
    note.position = "A你".len();
    document.push_note(note).unwrap();

    let mut annotation = Annotation::comment(1, Cow::Borrowed("A"), Cow::Borrowed("A你B"));
    let mut invalid = shape.clone();
    invalid.position = 2;
    assert!(annotation.push_shape(invalid).is_err());
    annotation.push_shape(shape.clone()).unwrap();

    let mut text_box = Shape::text_box(Cow::Borrowed("A你B"));
    let mut group = ShapeGroup::new();
    group.position = 2;
    assert!(text_box.push_text_shape_group(group).is_err());
    text_box.push_text_shape(shape).unwrap();
}

#[test]
fn rejects_starred_story_drawings_and_recursive_text_story_abuse() {
    for source in [
        r"{\rtf1{\footnote{\*\shp x}}}".to_string(),
        r"{\rtf1\chatn{\*\annotation{\*\shp x}}}".to_string(),
        r"{\rtf1{\shp{\*\shpinst{\shptxt{\*\shp x}}}}}".to_string(),
        r"{\rtf1{\endnote{\*\shpgrp x}}}".to_string(),
        r"{\rtf1{\*\do\dptxbx{\dptxbxtext{\*\shp x}}}}".to_string(),
    ] {
        assert!(RtfDocument::parse(&source).is_err(), "accepted {source}");
    }

    let mut nested = Shape::text_box(Cow::Borrowed(""));
    for _ in 0..65 {
        let mut parent = Shape::text_box(Cow::Borrowed(""));
        parent.text_shapes.push(nested);
        nested = parent;
    }
    assert!(nested.validate().is_err());

    let mut unordered = Note::footnote(Cow::Borrowed("1"), Cow::Borrowed("abcd"));
    let mut shape = Shape::new(ShapeType::Rectangle);
    shape.position = 1;
    let mut group = ShapeGroup::new();
    group.position = 3;
    unordered.shapes.push(shape);
    unordered.shape_groups.push(group);
    unordered.drawing_order = vec![StoryDrawing::ShapeGroup(0), StoryDrawing::Shape(0)];
    assert!(unordered.validate().is_err());
    unordered.drawing_order = vec![StoryDrawing::Shape(0), StoryDrawing::Shape(0)];
    assert!(unordered.validate().is_err());
}

#[test]
fn real_libreoffice_story_fixtures_round_trip_with_stable_ownership() {
    let manifest = std::path::Path::new(env!("CARGO_MANIFEST_DIR"));
    for relative in [
        "../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/copypaste-footnote.rtf",
        "../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/text-with-comment.rtf",
        "../../test-data/libreoffice-core/sw/qa/extras/rtfexport/data/shptxt-pard.rtf",
    ] {
        let source = std::fs::read(manifest.join(relative)).unwrap();
        let document = RtfDocument::parse_bytes(&source).unwrap();
        let note_count = document.notes().len();
        let annotation_count = document.annotations().len();
        let text_frame_count = document
            .shapes()
            .iter()
            .filter(|shape| shape.text_destination_present)
            .count();
        let reparsed = RtfDocument::parse_bytes(&write(&document)).unwrap();
        assert_eq!(reparsed.text(), document.text(), "{relative}");
        assert_eq!(reparsed.notes().len(), note_count, "{relative}");
        assert_eq!(reparsed.annotations().len(), annotation_count, "{relative}");
        assert_eq!(
            reparsed
                .shapes()
                .iter()
                .filter(|shape| shape.text_destination_present)
                .count(),
            text_frame_count,
            "{relative}"
        );
    }
}
