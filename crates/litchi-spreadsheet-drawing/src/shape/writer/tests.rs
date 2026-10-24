#![allow(clippy::unwrap_used, reason = "test assertions require direct values")]

use super::{ConnectionEndSpec, ConnectionShapeSpec, Emitter, GroupSpec, ObjectSpec, ShapeSpec};
use crate::shape::{Anchor, CellMarker, EditAs, Emu};
use litchi_drawingml::geom::Preset;

#[test]
fn emitter_writes_a_fresh_shape() {
    let anchor = Anchor::TwoCell {
        from: CellMarker {
            column: 0,
            column_offset: Emu(0),
            row: 0,
            row_offset: Emu(0),
        },
        to: CellMarker {
            column: 1,
            column_offset: Emu(0),
            row: 1,
            row_offset: Emu(0),
        },
        edit_as: EditAs::TwoCell,
    };
    let mut xml = String::new();
    Emitter::new(1, litchi_drawingml::geometry::Limits::DEFAULT)
        .write_anchored_shape(
            &mut xml,
            &ShapeSpec::text_box("Title", anchor, Preset::Rect, "Hello"),
        )
        .unwrap();
    assert!(xml.contains("Hello"));
}

#[test]
fn authored_spec_debug_redacts_names_descriptions_and_text() {
    let anchor = Anchor::TwoCell {
        from: CellMarker {
            column: 0,
            column_offset: Emu(0),
            row: 0,
            row_offset: Emu(0),
        },
        to: CellMarker {
            column: 1,
            column_offset: Emu(0),
            row: 1,
            row_offset: Emu(0),
        },
        edit_as: EditAs::TwoCell,
    };
    let mut shape = ShapeSpec::text_box(
        "AUTHORED_SHAPE_SECRET",
        anchor,
        Preset::Rect,
        "AUTHORED_TEXT_SECRET",
    );
    shape.description = Some("AUTHORED_DESCRIPTION_SECRET".to_owned());
    let start = ConnectionEndSpec {
        shape_name: "AUTHORED_START_SECRET".to_owned(),
        site: 1,
    };
    let end = ConnectionEndSpec {
        shape_name: "AUTHORED_END_SECRET".to_owned(),
        site: 2,
    };
    let mut connection = ConnectionShapeSpec::new(
        "AUTHORED_CONNECTION_SECRET",
        anchor,
        Preset::Rect,
        start.clone(),
        end.clone(),
    );
    connection.description = Some("AUTHORED_CONNECTION_DESCRIPTION_SECRET".to_owned());
    let group = GroupSpec::new("AUTHORED_GROUP_SECRET", anchor)
        .with_child(ObjectSpec::Shape(shape.clone()))
        .with_child(ObjectSpec::Connection(connection.clone()));
    let debug = format!(
        "{shape:?} {start:?} {end:?} {connection:?} {group:?} {:?}",
        ObjectSpec::Connection(connection.clone())
    );

    for secret in [
        "AUTHORED_SHAPE_SECRET",
        "AUTHORED_TEXT_SECRET",
        "AUTHORED_DESCRIPTION_SECRET",
        "AUTHORED_START_SECRET",
        "AUTHORED_END_SECRET",
        "AUTHORED_CONNECTION_SECRET",
        "AUTHORED_CONNECTION_DESCRIPTION_SECRET",
        "AUTHORED_GROUP_SECRET",
    ] {
        assert!(!debug.contains(secret), "debug leaked {secret}: {debug}");
    }
    assert!(debug.contains("paragraph_count"), "{debug}");
    assert!(debug.contains("child_count"), "{debug}");
}

#[test]
fn authoring_refuses_text_the_writer_cannot_write() {
    let anchor = Anchor::TwoCell {
        from: CellMarker {
            column: 0,
            column_offset: Emu(0),
            row: 0,
            row_offset: Emu(0),
        },
        to: CellMarker {
            column: 1,
            column_offset: Emu(0),
            row: 1,
            row_offset: Emu(0),
        },
        edit_as: EditAs::TwoCell,
    };
    let mut styled = ShapeSpec::text_box("Box", anchor, Preset::Rect, "Hello");
    styled.paragraphs[0].runs[0].properties.bold = Some(true);
    assert!(styled.validate(0).is_ok());

    let mut leveled = ShapeSpec::text_box("Box", anchor, Preset::Rect, "Hello");
    leveled.paragraphs[0].properties.level = Some(1);
    let error = leveled.validate(0).unwrap_err().to_string();
    assert!(error.contains("paragraph properties"), "{error}");

    let mut fonted = ShapeSpec::text_box("Box", anchor, Preset::Rect, "Hello");
    fonted.paragraphs[0].runs[0].properties.language = Some("en-US".to_owned());
    assert!(fonted.validate(0).is_err());
    let mut xml = "prefix".to_owned();
    assert!(
        Emitter::new(1, litchi_drawingml::geometry::Limits::DEFAULT)
            .write_anchored_shape(&mut xml, &fonted)
            .is_err()
    );
    assert_eq!(xml, "prefix");
}

#[test]
fn authoring_refuses_lossy_break_without_touching_existing_output() {
    let anchor = Anchor::TwoCell {
        from: CellMarker {
            column: 0,
            column_offset: Emu(0),
            row: 0,
            row_offset: Emu(0),
        },
        to: CellMarker {
            column: 1,
            column_offset: Emu(0),
            row: 1,
            row_offset: Emu(0),
        },
        edit_as: EditAs::TwoCell,
    };
    let mut shape = ShapeSpec::text_box("Box", anchor, Preset::Rect, "Hello");
    shape.paragraphs[0].runs[0].kind = litchi_drawingml::text::body::RunKind::Break;
    shape.paragraphs[0].runs[0].text = "wrong".to_owned();
    let mut xml = "prefix".to_owned();

    let error = Emitter::new(1, litchi_drawingml::geometry::Limits::DEFAULT)
        .write_anchored_shape(&mut xml, &shape)
        .unwrap_err();

    assert_eq!(xml, "prefix");
    assert!(error.to_string().contains("line breaks"));
}

#[test]
fn authoring_refuses_incomplete_text_projection_without_touching_existing_output() {
    let anchor = Anchor::TwoCell {
        from: CellMarker {
            column: 0,
            column_offset: Emu(0),
            row: 0,
            row_offset: Emu(0),
        },
        to: CellMarker {
            column: 1,
            column_offset: Emu(0),
            row: 1,
            row_offset: Emu(0),
        },
        edit_as: EditAs::TwoCell,
    };
    let mut shape = ShapeSpec::text_box("Box", anchor, Preset::Rect, "Hello");
    shape.properties.has_unmodeled_content = true;
    let mut xml = "prefix".to_owned();

    let error = Emitter::new(1, litchi_drawingml::geometry::Limits::DEFAULT)
        .write_anchored_shape(&mut xml, &shape)
        .unwrap_err();

    assert_eq!(xml, "prefix");
    assert!(error.to_string().contains("unmodeled content"));
}

#[test]
fn authoring_preserves_break_markup() {
    let anchor = Anchor::TwoCell {
        from: CellMarker {
            column: 0,
            column_offset: Emu(0),
            row: 0,
            row_offset: Emu(0),
        },
        to: CellMarker {
            column: 1,
            column_offset: Emu(0),
            row: 1,
            row_offset: Emu(0),
        },
        edit_as: EditAs::TwoCell,
    };
    let mut shape = ShapeSpec::text_box("Box", anchor, Preset::Rect, "Hello");
    shape.paragraphs[0].runs[0].kind = litchi_drawingml::text::body::RunKind::Break;
    shape.paragraphs[0].runs[0].text = "\n".to_owned();
    let mut xml = String::new();

    Emitter::new(1, litchi_drawingml::geometry::Limits::DEFAULT)
        .write_anchored_shape(&mut xml, &shape)
        .unwrap();

    assert!(xml.contains("<a:br><a:rPr/></a:br>"));
    assert!(!xml.contains("<a:t>"));
}
