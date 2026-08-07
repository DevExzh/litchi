#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{Coordinate, EscapeKind, Instruction, PathKind};
use crate::prop::{Id, Props};
use crate::{Error, Record, RecordKind};

const COMPLEX: u16 = 0x8000;

fn property(data: &mut Vec<u8>, id: Id, complex: bool, value: i32) {
    let raw = id.raw() | if complex { COMPLEX } else { 0 };
    data.extend_from_slice(&raw.to_le_bytes());
    data.extend_from_slice(&value.to_le_bytes());
}

fn points_array(points: &[(i32, i32)]) -> Vec<u8> {
    let mut data = Vec::with_capacity(6 + points.len() * 8);
    let count = u16::try_from(points.len()).expect("bounded test array");
    data.extend_from_slice(&count.to_le_bytes());
    data.extend_from_slice(&count.to_le_bytes());
    data.extend_from_slice(&8u16.to_le_bytes());
    for &(x, y) in points {
        data.extend_from_slice(&x.to_le_bytes());
        data.extend_from_slice(&y.to_le_bytes());
    }
    data
}

fn path_array(infos: &[u16]) -> Vec<u8> {
    let mut data = Vec::with_capacity(6 + infos.len() * 2);
    let count = u16::try_from(infos.len()).expect("bounded test array");
    data.extend_from_slice(&count.to_le_bytes());
    data.extend_from_slice(&count.to_le_bytes());
    data.extend_from_slice(&2u16.to_le_bytes());
    for info in infos {
        data.extend_from_slice(&info.to_le_bytes());
    }
    data
}

fn parse_properties<'data>(
    data: &'data mut Vec<u8>,
    shape_path: Option<u32>,
    points: Option<&[(i32, i32)]>,
    infos: Option<&[u16]>,
) -> Props<'data> {
    data.clear();
    let mut headers = Vec::new();
    let mut complex_data = Vec::new();
    if let Some(shape_path) = shape_path {
        property(&mut headers, Id::ShapePath, false, shape_path as i32);
    }
    if let Some(points) = points {
        let array = points_array(points);
        property(
            &mut headers,
            Id::Vertices,
            true,
            i32::try_from(array.len()).expect("bounded test array"),
        );
        complex_data.extend_from_slice(&array);
    }
    if let Some(infos) = infos {
        let array = path_array(infos);
        property(
            &mut headers,
            Id::SegmentInfo,
            true,
            i32::try_from(array.len()).expect("bounded test array"),
        );
        complex_data.extend_from_slice(&array);
    }
    headers.extend_from_slice(&complex_data);
    data.extend_from_slice(&headers);
    let count = if shape_path.is_some() { 1 } else { 0 }
        + usize::from(points.is_some())
        + usize::from(infos.is_some());
    let record = Record::from_parts(
        RecordKind::Opt,
        3,
        u16::try_from(count).expect("bounded test property table"),
        data.as_slice(),
    )
    .expect("valid test record");
    Props::parse(&record).expect("valid test properties")
}

#[test]
fn decodes_custom_vertices_and_path_info_without_copying_payloads() {
    let mut data = Vec::new();
    let points = [(0, 0), (100, 0), (100, 100), (0, 100)];
    let infos = [
        2,            // MoveTo, zero segments.
        2 << 3,       // LineTo, two points.
        1 << 3,       // LineTo, one point.
        3 | (1 << 3), // Close, one segment and no point.
    ];
    let properties = parse_properties(&mut data, Some(4), Some(&points), Some(&infos));
    let geometry = properties
        .geometry()
        .expect("valid geometry")
        .expect("geometry is present");

    assert_eq!(geometry.path_kind(), PathKind::Complex);
    assert_eq!(geometry.vertex_count(), 4);
    assert_eq!(geometry.segment_count(), 4);
    let expected_vertices = points_array(&points);
    let expected_segment_info = path_array(&infos);
    assert_eq!(geometry.raw_vertices(), Some(expected_vertices.as_slice()));
    assert_eq!(
        geometry.raw_segment_info(),
        Some(expected_segment_info.as_slice())
    );
    assert_eq!(
        geometry
            .vertices()
            .map(|point| point.raw())
            .collect::<Vec<_>>(),
        points
    );

    let decoded = geometry.segment_info().collect::<Vec<_>>();
    assert_eq!(decoded[0].instruction(), Instruction::MoveTo);
    assert_eq!(decoded[1].instruction(), Instruction::LineTo);
    assert_eq!(decoded[1].segments(), 2);
    assert_eq!(decoded[3].instruction(), Instruction::Close);
    assert_eq!(decoded[3].raw(), infos[3]);
    assert_eq!(geometry.raw_segment_info().map(|raw| raw.len()), Some(14));
}

#[test]
fn resolves_geometry_guide_references_losslessly() {
    let mut data = Vec::new();
    let points = [(i32::MIN, i32::MIN + 0x7F), (i32::MIN + 0x80, 9)];
    let properties = parse_properties(&mut data, Some(1), Some(&points), None);
    let geometry = properties
        .geometry()
        .expect("valid geometry")
        .expect("geometry is present");
    let decoded = geometry.vertices().collect::<Vec<_>>();

    assert_eq!(decoded[0].x(), Coordinate::Guide(0));
    assert_eq!(decoded[0].y(), Coordinate::Guide(0x7F));
    assert_eq!(decoded[0].x().raw(), i32::MIN);
    assert_eq!(decoded[0].y().raw(), i32::MIN + 0x7F);
    assert_eq!(decoded[1].x(), Coordinate::Value(i32::MIN + 0x80));
}

#[test]
fn preserves_unknown_path_instruction_and_escape_words() {
    let mut data = Vec::new();
    let points = [(10, 20)];
    let infos = [7, 5 | (31 << 3) | (1 << 8)];
    let properties = parse_properties(&mut data, Some(99), Some(&points), Some(&infos));
    let geometry = properties
        .geometry()
        .expect("unknown values remain lossless")
        .expect("geometry is present");
    let decoded = geometry.segment_info().collect::<Vec<_>>();

    assert_eq!(geometry.path_kind(), PathKind::Unknown(99));
    assert_eq!(decoded[0].instruction(), Instruction::Unknown(7));
    assert_eq!(decoded[0].raw(), infos[0]);
    assert_eq!(
        decoded[1].instruction(),
        Instruction::Escape(EscapeKind::Unknown(31))
    );
    assert_eq!(decoded[1].raw(), infos[1]);
    assert_eq!(
        &geometry.raw_segment_info().expect("raw array")[6..8],
        &infos[0].to_le_bytes()
    );
}

#[test]
fn rejects_invalid_geometry_array_width_and_point_consumption() {
    let mut invalid_data = Vec::new();
    let mut array = Vec::new();
    array.extend_from_slice(&1u16.to_le_bytes());
    array.extend_from_slice(&1u16.to_le_bytes());
    array.extend_from_slice(&4u16.to_le_bytes());
    array.extend_from_slice(&[0; 4]);
    property(
        &mut invalid_data,
        Id::Vertices,
        true,
        i32::try_from(array.len()).expect("bounded test array"),
    );
    invalid_data.extend_from_slice(&array);
    let record =
        Record::from_parts(RecordKind::Opt, 3, 1, invalid_data.as_slice()).expect("record");
    let properties = Props::parse(&record).expect("generic array remains structurally valid");
    assert!(matches!(
        properties.geometry(),
        Err(Error::MalformedGeometry {
            reason: "geometry IMsoArray element size is invalid"
        })
    ));

    let mut mismatch_data = Vec::new();
    let error = parse_properties(
        &mut mismatch_data,
        Some(4),
        Some(&[(0, 0)]),
        Some(&[2, 2 << 3]),
    );
    assert!(matches!(
        error.geometry(),
        Err(Error::MalformedGeometry {
            reason: "pSegmentInfo_complex consumes a different number of vertices"
        })
    ));

    let mut invalid_instruction_data = Vec::new();
    let properties = parse_properties(
        &mut invalid_instruction_data,
        Some(4),
        Some(&[(0, 0)]),
        Some(&[2 | (1 << 3)]),
    );
    assert!(matches!(
        properties.geometry(),
        Err(Error::MalformedGeometry {
            reason: "MSOPATHINFO segment count is invalid for its instruction"
        })
    ));
}

#[test]
fn requires_segment_info_for_complex_shape_path() {
    let mut data = Vec::new();
    let properties = parse_properties(&mut data, Some(4), Some(&[(0, 0)]), None);
    assert!(matches!(
        properties.geometry(),
        Err(Error::MalformedGeometry {
            reason: "complex shapePath requires non-empty pSegmentInfo_complex"
        })
    ));
}

#[test]
fn rejects_a_complex_shape_path_property() {
    let mut data = Vec::new();
    property(&mut data, Id::ShapePath, true, 4);
    data.extend_from_slice(&[0; 4]);
    let record = Record::from_parts(RecordKind::Opt, 3, 1, data.as_slice()).expect("record");
    let properties = Props::parse(&record).expect("generic property table");

    assert!(matches!(
        properties.geometry(),
        Err(Error::MalformedGeometry {
            reason: "shapePath must be a simple property"
        })
    ));
}
