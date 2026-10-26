//! Prost-differential tests of the straight-line codec (test-only oracle).

use prost::Message;

use super::{Error, Geometry, LINE_FLAGS, Limits, line, rewrite_line};
use crate::{tsd, tsp, tswp};

type TestResult = Result<(), Box<dyn std::error::Error>>;

fn element(kind: tsp::path::ElementType, x: f32) -> tsp::path::Element {
    tsp::path::Element {
        r#type: kind as i32,
        points: vec![tsp::Point { x, y: 0.0 }],
    }
}

fn shape(width: f32, flip: bool, elements: Vec<tsp::path::Element>) -> Vec<u8> {
    tswp::ShapeInfoArchive {
        super_: tsd::ShapeArchive {
            super_: tsd::DrawableArchive {
                geometry: Some(tsd::GeometryArchive {
                    position: Some(tsp::Point { x: 10.0, y: 20.0 }),
                    size: Some(tsp::Size { width, height: 0.0 }),
                    flags: Some(LINE_FLAGS),
                    angle: Some(30.0),
                }),
                ..Default::default()
            },
            pathsource: Some(tsd::PathSourceArchive {
                horizontal_flip: Some(flip),
                vertical_flip: Some(false),
                bezier_path_source: Some(tsd::BezierPathSourceArchive {
                    natural_size: Some(tsp::Size { width, height: 0.0 }),
                    path: Some(tsp::Path { elements }),
                    ..Default::default()
                }),
                ..Default::default()
            }),
            ..Default::default()
        },
        ..Default::default()
    }
    .encode_to_vec()
}

fn straight(width: f32, flip: bool) -> Vec<u8> {
    use tsp::path::ElementType;
    shape(
        width,
        flip,
        vec![
            element(ElementType::MoveTo, 0.0),
            element(ElementType::LineTo, width),
        ],
    )
}

#[test]
fn straight_lines_are_read() -> TestResult {
    let data = straight(100.0, true);
    let read = line(&data, Limits::DEFAULT)?.ok_or("not a line")?;
    assert_eq!(read.natural_width, 100.0);
    assert!(read.horizontal_flip);
    assert_eq!(read.geometry.position, Some([10.0, 20.0]));
    assert_eq!(read.geometry.size, Some([100.0, 0.0]));
    assert_eq!(read.geometry.flags, Some(LINE_FLAGS));
    assert_eq!(read.geometry.angle, Some(30.0));

    use tsp::path::ElementType;
    let curve = shape(
        100.0,
        false,
        vec![
            element(ElementType::MoveTo, 0.0),
            element(ElementType::CurveTo, 100.0),
        ],
    );
    assert_eq!(line(&curve, Limits::DEFAULT)?, None);
    let three = shape(
        100.0,
        false,
        vec![
            element(ElementType::MoveTo, 0.0),
            element(ElementType::LineTo, 50.0),
            element(ElementType::LineTo, 100.0),
        ],
    );
    assert_eq!(line(&three, Limits::DEFAULT)?, None);
    Ok(())
}

#[test]
fn rewriting_a_line_patches_in_place() -> TestResult {
    let data = straight(100.0, false);
    let geometry = Geometry {
        position: Some([5.0, 7.0]),
        size: Some([80.0, 0.0]),
        flags: Some(LINE_FLAGS),
        angle: Some(45.0),
    };
    let output = rewrite_line(&data, geometry, 80.0, Limits::DEFAULT)?;
    assert!(output.changed());
    assert_eq!(output.as_bytes().len(), data.len());
    let decoded = tswp::ShapeInfoArchive::decode(output.as_bytes())?;
    let stored = decoded.super_.super_.geometry.ok_or("geometry")?;
    assert_eq!(stored.position, Some(tsp::Point { x: 5.0, y: 7.0 }));
    assert_eq!(stored.angle, Some(45.0));
    let bezier = decoded
        .super_
        .pathsource
        .and_then(|path| path.bezier_path_source)
        .ok_or("bezier")?;
    assert_eq!(bezier.natural_size.map(|size| size.width), Some(80.0));
    let elements = bezier.path.ok_or("path")?.elements;
    assert_eq!(elements[1].points[0].x, 80.0);
    assert_eq!(elements[0].points[0].x, 0.0);

    let read = line(output.as_bytes(), Limits::DEFAULT)?.ok_or("line")?;
    assert_eq!(read.geometry, geometry);
    let same = rewrite_line(output.as_bytes(), geometry, 80.0, Limits::DEFAULT)?;
    assert!(!same.changed());
    Ok(())
}

#[test]
fn non_lines_are_refused() {
    use tsp::path::ElementType;
    let curve = shape(
        100.0,
        false,
        vec![
            element(ElementType::MoveTo, 0.0),
            element(ElementType::CurveTo, 100.0),
        ],
    );
    let geometry = Geometry {
        position: Some([0.0, 0.0]),
        size: Some([1.0, 0.0]),
        flags: Some(LINE_FLAGS),
        angle: Some(0.0),
    };
    assert!(matches!(
        rewrite_line(&curve, geometry, 1.0, Limits::DEFAULT),
        Err(Error::Malformed { .. })
    ));
    assert!(matches!(
        line(&[0x0a, 0x05], Limits::DEFAULT),
        Err(Error::Malformed { .. })
    ));
}
