#![allow(clippy::unwrap_used, reason = "test assertions require direct values")]

use super::{ConnectionEndSpec, ConnectionShapeSpec, Emitter, GroupSpec, ObjectSpec, ShapeSpec};
use crate::shape::{Anchor, CellMarker, EditAs, Emu};
use litchi_drawingml::geom::Preset;
use litchi_drawingml::geometry::{
    CustomGeometry, Formula, Guide, Limits as GeometryLimits, Path, PathCommand, Point,
};

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
        .unwrap()
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
            .unwrap()
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
        .unwrap()
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
        .unwrap()
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
        .unwrap()
        .write_anchored_shape(&mut xml, &shape)
        .unwrap();

    assert!(xml.contains("<a:br><a:rPr/></a:br>"));
    assert!(!xml.contains("<a:t>"));
}

#[test]
fn for_objects_validates_geometry_and_nesting_before_reserving_ids() {
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
    let geometry = CustomGeometry::new()
        .with_guide(Guide::new("g1", Formula::literal(1)))
        .with_guide(Guide::new("g2", Formula::literal(2)))
        .with_path(
            Path::new(21_600, 21_600)
                .with_command(PathCommand::MoveTo(Point::new(0, 0)))
                .with_command(PathCommand::LineTo(Point::new(21_600, 21_600)))
                .with_command(PathCommand::Close),
        );
    let shape = ShapeSpec::custom("shape", anchor, geometry, "");
    let tiny = GeometryLimits::new(4096, 4096, 1, 1024, 65_536, 1 << 20);
    assert!(Emitter::for_objects(1, std::slice::from_ref(&shape), &[], &[], tiny).is_err());

    let exact_bytes = litchi_drawingml::geometry::writer::measure_custom_geometry(
        shape.geometry.custom().unwrap(),
        &GeometryLimits::DEFAULT,
    )
    .unwrap();
    let exact = GeometryLimits {
        max_output_bytes: exact_bytes,
        ..GeometryLimits::DEFAULT
    };
    assert!(Emitter::for_objects(1, std::slice::from_ref(&shape), &[], &[], exact).is_ok());
    let one_under = GeometryLimits {
        max_output_bytes: exact_bytes - 1,
        ..GeometryLimits::DEFAULT
    };
    assert!(matches!(
        Emitter::for_objects(1, &[shape], &[], &[], one_under),
        Err(crate::Error::Drawing(litchi_drawingml::Error::Limit {
            kind: litchi_drawingml::LimitKind::OutputBytes,
            observed,
            maximum,
            ..
        })) if observed == exact_bytes && maximum == exact_bytes - 1
    ));

    let mut group = GroupSpec::new("g0", anchor);
    for depth in 1..=40 {
        group = GroupSpec::new(format!("g{depth}"), anchor).with_child(ObjectSpec::Group(group));
    }
    assert!(Emitter::for_objects(1, &[], &[group], &[], GeometryLimits::DEFAULT).is_err());
}

#[test]
fn few_top_level_groups_can_hold_many_children_before_connection_write() {
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
    let mut group = GroupSpec::new("outer", anchor);
    for group_index in 0..4 {
        let mut child_group = GroupSpec::new(format!("group-{group_index}"), anchor);
        for child_index in 0..1024 {
            child_group = child_group.with_child(ObjectSpec::Shape(ShapeSpec::shape(
                format!("shape-{group_index}-{child_index}"),
                anchor,
                Preset::Rect,
                "",
            )));
        }
        group = group.with_child(ObjectSpec::Group(child_group));
    }
    let connection = ConnectionShapeSpec::new(
        "connector",
        anchor,
        Preset::Rect,
        ConnectionEndSpec {
            shape_name: "shape-0-0".to_owned(),
            site: 0,
        },
        ConnectionEndSpec {
            shape_name: "shape-1-0".to_owned(),
            site: 0,
        },
    );
    let mut emitter = Emitter::for_objects(
        1,
        &[],
        &[group],
        std::slice::from_ref(&connection),
        GeometryLimits::DEFAULT,
    )
    .unwrap();
    let mut xml = String::new();
    emitter
        .write_anchored_connection(&mut xml, &connection)
        .unwrap();
    assert!(xml.contains("connector"));
}

#[test]
fn incremental_writes_preflight_geometry_before_anchor_or_id_state() {
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
    let geometry = CustomGeometry::new().with_path(
        Path::new(21_600, 21_600)
            .with_command(PathCommand::MoveTo(Point::new(0, 0)))
            .with_command(PathCommand::LineTo(Point::new(21_600, 21_600)))
            .with_command(PathCommand::Close),
    );
    let exact_bytes = litchi_drawingml::geometry::writer::measure_custom_geometry(
        &geometry,
        &GeometryLimits::DEFAULT,
    )
    .unwrap();
    let exact = GeometryLimits {
        max_output_bytes: exact_bytes,
        ..GeometryLimits::DEFAULT
    };
    let one_under = GeometryLimits {
        max_output_bytes: exact_bytes - 1,
        ..GeometryLimits::DEFAULT
    };
    let shape = ShapeSpec::custom("custom", anchor, geometry.clone(), "");
    let fallback = ShapeSpec::shape("fallback", anchor, Preset::Rect, "");

    let mut shape_emitter = Emitter::new(1, one_under).unwrap();
    let mut shape_xml = "prefix".to_owned();
    assert!(
        shape_emitter
            .write_anchored_shape(&mut shape_xml, &shape)
            .is_err()
    );
    assert_eq!(shape_xml, "prefix");
    shape_emitter
        .write_anchored_shape(&mut shape_xml, &fallback)
        .unwrap();
    assert!(shape_xml.contains(r#"id="1""#));

    let group = GroupSpec::new("group", anchor).with_child(ObjectSpec::Shape(shape.clone()));
    let mut group_emitter = Emitter::new(1, one_under).unwrap();
    let mut group_xml = "prefix".to_owned();
    assert!(
        group_emitter
            .write_anchored_group(&mut group_xml, &group)
            .is_err()
    );
    assert_eq!(group_xml, "prefix");
    group_emitter
        .write_anchored_shape(&mut group_xml, &fallback)
        .unwrap();
    assert!(group_xml.contains(r#"id="1""#));

    let connection = ConnectionShapeSpec::new(
        "connection",
        anchor,
        geometry,
        ConnectionEndSpec {
            shape_name: "start".to_owned(),
            site: 0,
        },
        ConnectionEndSpec {
            shape_name: "end".to_owned(),
            site: 0,
        },
    );
    let mut connection_emitter = Emitter::new(1, one_under).unwrap();
    let mut connection_xml = "prefix".to_owned();
    assert!(
        connection_emitter
            .write_anchored_connection(&mut connection_xml, &connection)
            .is_err()
    );
    assert_eq!(connection_xml, "prefix");
    connection_emitter
        .write_anchored_shape(&mut connection_xml, &fallback)
        .unwrap();
    assert!(connection_xml.contains(r#"id="1""#));

    let mut exact_emitter = Emitter::new(1, exact).unwrap();
    let mut exact_xml = String::new();
    exact_emitter
        .write_anchored_shape(&mut exact_xml, &shape)
        .unwrap();
    assert!(exact_xml.contains("<a:custGeom>"));
}
