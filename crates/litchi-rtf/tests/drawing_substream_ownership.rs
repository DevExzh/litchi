#![allow(
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_same,
    clippy::shadow_unrelated,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design and rebind fixture names across steps"
)]

use litchi_rtf::{
    BodyStoryEvent, Cell, CellStoryEvent, FieldOwner, NoteSeparatorElement, NoteSeparatorKind,
    RtfDocument, RtfWriter, Shape, ShapeGroup, ShapeType, StoryDrawing, StoryEvent, StoryField,
};
use std::borrow::Cow;

const GROUP_THEN_SHAPE: &str = concat!(
    r"{\shpgrp{\*\shpinst{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 1}}}}}}",
    r"{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 1}}}}"
);
const ROOT_GROUP: &str = r"{\shpgrp{\*\shpinst{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 1}}}}}}";
const ROOT_SHAPE: &str = r"{\shp{\*\shpinst{\sp{\sn shapeType}{\sv 1}}}}";

fn write_document(document: &RtfDocument<'_>) -> String {
    let mut output = Vec::new();
    RtfWriter::new(&mut output)
        .write_document(document)
        .unwrap();
    String::from_utf8(output).unwrap()
}

fn assert_group_then_shape(
    shapes: &[Shape<'_>],
    groups: &[ShapeGroup<'_>],
    order: &[StoryDrawing],
    offset: usize,
) {
    assert_eq!(shapes.len(), 1);
    assert_eq!(groups.len(), 1);
    assert_eq!(shapes[0].position, offset);
    assert_eq!(groups[0].position, offset);
    assert_eq!(order, [StoryDrawing::ShapeGroup(0), StoryDrawing::Shape(0)]);
}

fn field_index(document: &RtfDocument<'_>, instruction: &str) -> usize {
    document
        .fields()
        .iter()
        .position(|field| field.instruction.trim() == instruction)
        .unwrap_or_else(|| {
            panic!(
                "missing {instruction}; parsed {:?}",
                document
                    .fields()
                    .iter()
                    .map(|field| field.instruction.as_ref())
                    .collect::<Vec<_>>()
            )
        })
}

#[test]
fn owns_and_round_trips_equal_offset_drawings_in_supported_substreams() {
    let source = format!(
        "{{\\rtf1\\ansi{{\\*\\ftnsep N{{\\b {GROUP_THEN_SHAPE}}}\\chftnsep\\par}}\\trowd\\cellx1200\\intbl C{{\\i {GROUP_THEN_SHAPE}}}D\\cell\\row}}",
    );
    let document = RtfDocument::parse(&source).unwrap();

    let separator = document
        .note_separators()
        .get(NoteSeparatorKind::FootnoteSeparator)
        .unwrap();
    assert_eq!(separator.text(), "N\n");
    assert_group_then_shape(
        &separator.shapes,
        &separator.shape_groups,
        &separator.drawing_order(),
        1,
    );
    assert!(matches!(
        separator.elements[1],
        NoteSeparatorElement::Drawing(StoryDrawing::ShapeGroup(0))
    ));
    assert!(matches!(
        separator.elements[2],
        NoteSeparatorElement::Drawing(StoryDrawing::Shape(0))
    ));

    let cell = &document.tables()[0].rows()[0].cells()[0];
    assert_eq!(cell.text(), "CD");
    assert_group_then_shape(cell.shapes(), cell.shape_groups(), cell.drawing_order(), 1);

    let reparsed = RtfDocument::parse(&write_document(&document)).unwrap();
    let reparsed_separator = reparsed
        .note_separators()
        .get(NoteSeparatorKind::FootnoteSeparator)
        .unwrap();
    assert_group_then_shape(
        &reparsed_separator.shapes,
        &reparsed_separator.shape_groups,
        &reparsed_separator.drawing_order(),
        1,
    );
    let reparsed_cell = &reparsed.tables()[0].rows()[0].cells()[0];
    assert_group_then_shape(
        reparsed_cell.shapes(),
        reparsed_cell.shape_groups(),
        reparsed_cell.drawing_order(),
        1,
    );
}

#[test]
fn field_result_drawings_round_trip_through_the_field_api() {
    let source = format!(
        "{{\\rtf1{{\\field{{\\*\\fldinst TEST}}{{\\fldrslt A{{\\b {GROUP_THEN_SHAPE}}}B}}}}}}"
    );
    let document = RtfDocument::parse(&source).unwrap();
    let field = &document.fields()[0];
    assert_eq!(field.result, "AB");
    assert_group_then_shape(&field.shapes, &field.shape_groups, &field.drawing_order, 1);

    let mut field_output = Vec::new();
    RtfWriter::new(&mut field_output)
        .write_field(field)
        .unwrap();
    let wrapped = format!("{{\\rtf1{}}}", String::from_utf8(field_output).unwrap());
    let reparsed = RtfDocument::parse(&wrapped).unwrap();
    let reparsed_field = &reparsed.fields()[0];
    assert_eq!(reparsed_field.result, "AB");
    assert_group_then_shape(
        &reparsed_field.shapes,
        &reparsed_field.shape_groups,
        &reparsed_field.drawing_order,
        1,
    );
}

#[test]
fn document_writer_places_body_fields_in_exact_drawing_source_order() {
    let source = format!(
        "{{\\rtf1 A{ROOT_GROUP}{{\\field{{\\*\\fldinst TEST}}{{\\fldrslt R{{\\b {GROUP_THEN_SHAPE}}}S}}}}{ROOT_SHAPE}B}}"
    );
    let document = RtfDocument::parse(&source).unwrap();
    assert_eq!(document.text(), "AB");
    assert_eq!(
        document.body_story_events(),
        [
            BodyStoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            BodyStoryEvent::Field(0),
            BodyStoryEvent::Drawing(StoryDrawing::Shape(0)),
        ]
    );
    let field = &document.fields()[0];
    assert_eq!(field.owner, FieldOwner::Body);
    assert_eq!((field.position, field.range_end), (1, 1));
    assert_eq!(field.result, "RS");
    assert_group_then_shape(&field.shapes, &field.shape_groups, &field.drawing_order, 1);

    let serialized = write_document(&document);
    let group_position = serialized.find("\\shpgrp").unwrap();
    let field_position = serialized.find("\\field").unwrap();
    let trailing_shape_position = serialized.rfind("{\\shp{").unwrap();
    assert!(group_position < field_position && field_position < trailing_shape_position);

    let reparsed = RtfDocument::parse(&serialized).unwrap();
    assert_eq!(reparsed.text(), "AB");
    assert_eq!(reparsed.body_story_events(), document.body_story_events());
    assert_eq!(reparsed.fields()[0].owner, FieldOwner::Body);
    assert_eq!(
        (
            reparsed.fields()[0].position,
            reparsed.fields()[0].range_end
        ),
        (1, 1)
    );
    assert_group_then_shape(
        &reparsed.fields()[0].shapes,
        &reparsed.fields()[0].shape_groups,
        &reparsed.fields()[0].drawing_order,
        1,
    );
}

#[test]
fn cell_story_events_preserve_nested_table_drawing_ties() {
    let source = format!(
        "{{\\rtf1\\trowd\\cellx5000\\intbl C{ROOT_GROUP}{{\\intbl\\itap2 X\\nestcell}}{{\\*\\nesttableprops\\itap2\\trowd\\cellx1700\\nestrow}}{{\\nonesttables\\par}}{ROOT_SHAPE}D\\cell\\row}}"
    );
    let document = RtfDocument::parse(&source).unwrap();
    let cell = &document.tables()[0].rows()[0].cells()[0];
    assert_eq!(cell.text(), "CD");
    assert_eq!(
        cell.story_events(),
        [
            CellStoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            CellStoryEvent::NestedTable(0),
            CellStoryEvent::Drawing(StoryDrawing::Shape(0)),
        ]
    );

    let reparsed = RtfDocument::parse(&write_document(&document)).unwrap();
    let cell = &reparsed.tables()[0].rows()[0].cells()[0];
    assert_eq!(cell.text(), "CD");
    assert_eq!(
        cell.story_events(),
        [
            CellStoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            CellStoryEvent::NestedTable(0),
            CellStoryEvent::Drawing(StoryDrawing::Shape(0)),
        ]
    );
}

#[test]
fn nested_table_cells_own_their_drawing_story() {
    let source = format!(
        "{{\\rtf1\\trowd\\cellx5000\\intbl {{\\intbl\\itap2 X{{\\b {GROUP_THEN_SHAPE}}}Y\\nestcell}}{{\\*\\nesttableprops\\itap2\\trowd\\cellx1700\\nestrow}}{{\\nonesttables\\par}}\\cell\\row}}"
    );
    let document = RtfDocument::parse(&source).unwrap();
    let nested = &document.tables()[0].rows()[0].cells()[0].nested_tables()[0].table;
    let cell = &nested.rows()[0].cells()[0];
    assert_eq!(cell.text(), "XY");
    assert_group_then_shape(cell.shapes(), cell.shape_groups(), cell.drawing_order(), 1);

    let reparsed = RtfDocument::parse(&write_document(&document)).unwrap();
    let nested = &reparsed.tables()[0].rows()[0].cells()[0].nested_tables()[0].table;
    let cell = &nested.rows()[0].cells()[0];
    assert_eq!(cell.text(), "XY");
    assert_group_then_shape(cell.shapes(), cell.shape_groups(), cell.drawing_order(), 1);
}

#[test]
fn typed_substream_apis_reject_invalid_utf8_and_order_references() {
    let mut cell = Cell::new(Cow::Borrowed("你"));
    let mut invalid = Shape::new(ShapeType::Rectangle);
    invalid.position = 1;
    assert!(cell.push_shape(invalid).is_err());

    let mut first = Shape::new(ShapeType::Rectangle);
    first.position = "你".len();
    cell.push_shape(first).unwrap();
    let mut backwards = ShapeGroup::new();
    backwards.position = 0;
    assert!(cell.push_shape_group(backwards).is_err());

    let source = format!("{{\\rtf1{{\\*\\ftnsep X{GROUP_THEN_SHAPE}}}}}");
    let document = RtfDocument::parse(&source).unwrap();
    let mut separator = document
        .note_separators()
        .get(NoteSeparatorKind::FootnoteSeparator)
        .unwrap()
        .clone();
    separator
        .elements
        .push(NoteSeparatorElement::Drawing(StoryDrawing::Shape(0)));
    assert!(separator.validate().is_err());
}

#[test]
fn header_footer_and_note_fields_preserve_equal_offset_drawing_order() {
    let source = format!(
        "{{\\rtf1{{\\header H{ROOT_GROUP}{{\\field{{\\*\\fldinst HEADER}}{{\\fldrslt hr}}}}{ROOT_SHAPE}I}}{{\\footer F{ROOT_GROUP}{{\\field{{\\*\\fldinst FOOTER}}{{\\fldrslt fr}}}}{ROOT_SHAPE}G}}A{{\\footnote N{ROOT_GROUP}{{\\field{{\\*\\fldinst FOOTNOTE}}{{\\fldrslt nr}}}}{ROOT_SHAPE}O}}B{{\\endnote E{ROOT_GROUP}{{\\field{{\\*\\fldinst ENDNOTE}}{{\\fldrslt er}}}}{ROOT_SHAPE}D}}C}}"
    );
    let document = RtfDocument::parse(&source).unwrap();
    let header_index = field_index(&document, "HEADER");
    let footer_index = field_index(&document, "FOOTER");
    let footnote_index = field_index(&document, "FOOTNOTE");
    let endnote_index = field_index(&document, "ENDNOTE");
    let header = &document.sections()[0].headers_footers[0];
    let footer = &document.sections()[0].headers_footers[1];
    assert_eq!(
        header.story_events,
        [
            StoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            StoryEvent::Field(StoryField {
                field_index: header_index,
                position: 1
            }),
            StoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    );
    assert_eq!(
        footer.story_events,
        [
            StoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            StoryEvent::Field(StoryField {
                field_index: footer_index,
                position: 1
            }),
            StoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    );
    assert_eq!(
        document.notes()[0].story_events,
        [
            StoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            StoryEvent::Field(StoryField {
                field_index: footnote_index,
                position: 1
            }),
            StoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    );
    assert_eq!(
        document.notes()[1].story_events,
        [
            StoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            StoryEvent::Field(StoryField {
                field_index: endnote_index,
                position: 1
            }),
            StoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    );
    assert_eq!(document.fields()[header_index].owner, FieldOwner::Header);
    assert_eq!(document.fields()[footer_index].owner, FieldOwner::Footer);
    assert_eq!(
        document.fields()[footnote_index].owner,
        FieldOwner::Footnote
    );
    assert_eq!(document.fields()[endnote_index].owner, FieldOwner::Endnote);

    let reparsed = RtfDocument::parse(&write_document(&document)).unwrap();
    let header = &reparsed.sections()[0].headers_footers[0];
    let footer = &reparsed.sections()[0].headers_footers[1];
    assert!(matches!(
        header.story_events.as_slice(),
        [
            StoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            StoryEvent::Field(_),
            StoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    ));
    assert!(matches!(
        footer.story_events.as_slice(),
        [
            StoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            StoryEvent::Field(_),
            StoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    ));
    for note in reparsed.notes() {
        assert!(matches!(
            note.story_events.as_slice(),
            [
                StoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
                StoryEvent::Field(_),
                StoryEvent::Drawing(StoryDrawing::Shape(0))
            ]
        ));
    }
}

#[test]
fn outer_and_nested_cells_preserve_generic_fields_among_equal_offset_events() {
    let source = format!(
        "{{\\rtf1\\trowd\\cellx5000\\intbl C{ROOT_GROUP}{{\\field{{\\*\\fldinst OUTER}}{{\\fldrslt o}}}}{{{{\\intbl\\itap2 X{ROOT_GROUP}{{\\field{{\\*\\fldinst INNER}}{{\\fldrslt i}}}}{ROOT_SHAPE}Y\\nestcell}}}}{{\\*\\nesttableprops\\itap2\\trowd\\cellx1700\\nestrow}}{{\\nonesttables\\par}}{ROOT_SHAPE}D\\cell\\row}}"
    );
    let document = RtfDocument::parse(&source).unwrap();
    let outer_index = field_index(&document, "OUTER");
    let inner_index = field_index(&document, "INNER");
    let outer = &document.tables()[0].rows()[0].cells()[0];
    assert_eq!(
        outer.story_events(),
        [
            CellStoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            CellStoryEvent::Field(StoryField {
                field_index: outer_index,
                position: 1
            }),
            CellStoryEvent::NestedTable(0),
            CellStoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    );
    let inner = &outer.nested_tables()[0].table.rows()[0].cells()[0];
    assert_eq!(
        inner.story_events(),
        [
            CellStoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            CellStoryEvent::Field(StoryField {
                field_index: inner_index,
                position: 1
            }),
            CellStoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    );
    assert_eq!(
        document.fields()[outer_index].owner,
        FieldOwner::TableCell(1)
    );
    assert_eq!(
        document.fields()[inner_index].owner,
        FieldOwner::TableCell(2)
    );

    let reparsed = RtfDocument::parse(&write_document(&document)).unwrap();
    let outer = &reparsed.tables()[0].rows()[0].cells()[0];
    assert!(matches!(
        outer.story_events(),
        [
            CellStoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            CellStoryEvent::Field(_),
            CellStoryEvent::NestedTable(0),
            CellStoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    ));
    let inner = &outer.nested_tables()[0].table.rows()[0].cells()[0];
    assert!(matches!(
        inner.story_events(),
        [
            CellStoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            CellStoryEvent::Field(_),
            CellStoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    ));
}

#[test]
fn nested_generic_fields_round_trip_in_parent_result_event_order() {
    let source = format!(
        "{{\\rtf1 A{{\\field{{\\*\\fldinst PARENT}}{{\\fldrslt R{ROOT_GROUP}{{\\field{{\\*\\fldinst CHILD}}{{\\fldrslt child}}}}{ROOT_SHAPE}S}}}}B}}"
    );
    let document = RtfDocument::parse(&source).unwrap();
    let child_index = field_index(&document, "CHILD");
    let parent_index = field_index(&document, "PARENT");
    let parent = &document.fields()[parent_index];
    assert_eq!(parent.owner, FieldOwner::Body);
    assert_eq!(
        document.fields()[child_index].owner,
        FieldOwner::FieldResult
    );
    assert_eq!(parent.result, "RS");
    assert_eq!(
        parent.result_events,
        [
            StoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            StoryEvent::Field(StoryField {
                field_index: child_index,
                position: 1
            }),
            StoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    );

    let reparsed = RtfDocument::parse(&write_document(&document)).unwrap();
    let parent_index = field_index(&reparsed, "PARENT");
    let parent = &reparsed.fields()[parent_index];
    assert_eq!(parent.result, "RS");
    assert!(matches!(
        parent.result_events.as_slice(),
        [
            StoryEvent::Drawing(StoryDrawing::ShapeGroup(0)),
            StoryEvent::Field(_),
            StoryEvent::Drawing(StoryDrawing::Shape(0))
        ]
    ));
}
