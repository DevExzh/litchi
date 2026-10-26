#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::codec::text_from_ppt_records;
use super::{Container, RecordKind, anchor, parse, parse_drawing, text_from_drawing};

fn record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

#[test]
fn ppt_text_atoms_use_their_specified_encodings() {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&4000u16.to_le_bytes());
    bytes.extend_from_slice(&4u32.to_le_bytes());
    bytes.extend_from_slice(&[0x3d, 0xd8, 0x00, 0xde]);
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&4008u16.to_le_bytes());
    bytes.extend_from_slice(&2u32.to_le_bytes());
    bytes.extend_from_slice(&[0x80, 0xe9]);

    let mut text = String::new();
    text_from_ppt_records(
        &bytes,
        &mut text,
        super::package::MAX_HOST_RECORDS,
        crate::SourceLimits::DEFAULT_MAX_DEPTH,
    )
    .unwrap();

    assert_eq!(text, "😀\n\u{80}é");
}

#[test]
fn malformed_embedded_record_stops_without_panicking() {
    let mut text = String::new();
    assert!(
        text_from_ppt_records(
            &[0; 7],
            &mut text,
            super::package::MAX_HOST_RECORDS,
            crate::SourceLimits::DEFAULT_MAX_DEPTH,
        )
        .is_err()
    );
    assert!(text.is_empty());
}

#[test]
fn embedded_text_records_honor_the_caller_event_limit() {
    let first = record(0, 0, 4000, &[0; 2]);
    let second = record(0, 0, 4000, &[0; 2]);
    let data = [first.as_slice(), second.as_slice()].concat();
    let mut text = String::new();

    let error = text_from_ppt_records(&data, &mut text, 1, crate::SourceLimits::DEFAULT_MAX_DEPTH)
        .unwrap_err();
    assert!(matches!(
        error,
        crate::package::Error::Limit {
            kind: crate::SourceLimitKind::Events,
            observed: 2,
            maximum: 1,
        }
    ));
}

#[test]
fn embedded_text_records_honor_the_caller_depth_limit() {
    let text_atom = record(0, 0, 4000, &[0; 2]);
    let inner = record(0, 0, 1000, &text_atom);
    let outer = record(0, 0, 1000, &inner);
    let mut text = String::new();

    let error = text_from_ppt_records(&outer, &mut text, 16, 1).unwrap_err();
    assert!(matches!(
        error,
        crate::package::Error::Limit {
            kind: crate::SourceLimitKind::Depth,
            observed: 2,
            maximum: 1,
        }
    ));
}

#[test]
fn concatenated_drawing_roots_are_complete_and_strict() {
    let dg = record(0, 0, RecordKind::Dg.raw(), &[0; 8]);
    let root = record(0x0f, 0, RecordKind::DgContainer.raw(), &dg);
    let stream = [root.as_slice(), root.as_slice()].concat();

    assert!(
        parse(&stream, &super::Limits::default())
            .unwrap()
            .is_empty()
    );
    assert_eq!(
        text_from_drawing(&stream, crate::SourceLimits::default()).unwrap(),
        ""
    );

    let malformed = [stream.as_slice(), &[0]].concat();
    assert!(parse(&malformed, &super::Limits::default()).is_err());
    assert!(text_from_drawing(&malformed, crate::SourceLimits::default()).is_err());
}

#[test]
fn source_limits_intersect_officeart_safety_ceiling() {
    let dg = record(0, 0, RecordKind::Dg.raw(), &[0; 8]);
    let mut shape_atom = Vec::new();
    shape_atom.extend_from_slice(&42u32.to_le_bytes());
    shape_atom.extend_from_slice(&0x0800u32.to_le_bytes());
    let shape = record(
        0x0f,
        0,
        RecordKind::SpContainer.raw(),
        &record(2, 1, RecordKind::Sp.raw(), &shape_atom),
    );
    let shape_drawing_payload = [dg.as_slice(), shape.as_slice()].concat();
    let shape_drawing = record(
        0x0f,
        0,
        RecordKind::DgContainer.raw(),
        &shape_drawing_payload,
    );
    let source = crate::SourceLimits::default();
    let limits = super::limits_for_source(source);
    let expected_body_limit = source
        .max_record_bytes()
        .saturating_sub(8)
        .min(source.max_record_payload_bytes());
    assert_eq!(limits.max_record_bytes as usize, expected_body_limit);
    assert_eq!(limits.max_depth, super::Limits::default().max_depth);
    assert_eq!(parse(&shape_drawing, &limits).unwrap().len(), 1);

    let hard_depth_source = crate::SourceLimits::builder()
        .max_depth(64)
        .unwrap()
        .build()
        .unwrap();
    let hard_depth = super::limits_for_source(hard_depth_source);
    assert_eq!(hard_depth.max_depth, 64);
    assert_eq!(parse(&shape_drawing, &hard_depth).unwrap().len(), 1);

    let restrictive_source = crate::SourceLimits::builder()
        .max_record_bytes(1024)
        .unwrap()
        .max_events(16)
        .unwrap()
        .max_depth(1)
        .unwrap()
        .build()
        .unwrap();
    let restrictive = super::limits_for_source(restrictive_source);
    assert_eq!(restrictive.max_record_bytes, 1024 - 8);
    assert_eq!(restrictive.max_records, 16);
    assert_eq!(restrictive.max_depth, 1);
    assert!(parse(&shape_drawing, &restrictive).is_ok());

    let event_limited_source = crate::SourceLimits::builder()
        .max_events(1)
        .unwrap()
        .build()
        .unwrap();
    let event_limited = super::limits_for_source(event_limited_source);
    assert_eq!(event_limited.max_records, 1);
    assert!(parse(&shape_drawing, &event_limited).is_err());

    let payload_limited_source = crate::SourceLimits::builder()
        .max_record_bytes(1024)
        .unwrap()
        .max_record_payload_bytes(8)
        .unwrap()
        .build()
        .unwrap();
    let payload_limited = super::limits_for_source(payload_limited_source);
    assert_eq!(payload_limited.max_record_bytes, 8);
    assert!(parse(&shape_drawing, &payload_limited).is_err());

    let encoded_limited_source = crate::SourceLimits::builder()
        .max_record_bytes(12)
        .unwrap()
        .max_record_payload_bytes(1024)
        .unwrap()
        .build()
        .unwrap();
    let encoded_limited = super::limits_for_source(encoded_limited_source);
    assert_eq!(encoded_limited.max_record_bytes, 4);
    assert!(parse(&shape_drawing, &encoded_limited).is_err());

    let spgr = record(0, 0, RecordKind::Spgr.raw(), &[0; 16]);
    let group_header = |flags: u32| {
        let mut shape_atom = Vec::new();
        shape_atom.extend_from_slice(&1u32.to_le_bytes());
        shape_atom.extend_from_slice(&flags.to_le_bytes());
        let shape = record(2, 0, RecordKind::Sp.raw(), &shape_atom);
        [spgr.as_slice(), shape.as_slice()].concat()
    };
    let inner_group_header = group_header(0x0001);
    let inner_group = record(
        0x0f,
        0,
        RecordKind::SpgrContainer.raw(),
        &inner_group_header,
    );
    let outer_group_header = group_header(0x0005);
    let outer_group_payload = [outer_group_header.as_slice(), inner_group.as_slice()].concat();
    let outer_group = record(
        0x0f,
        0,
        RecordKind::SpgrContainer.raw(),
        &outer_group_payload,
    );
    let nested_drawing_payload = [dg.as_slice(), outer_group.as_slice()].concat();
    let nested_drawing = record(
        0x0f,
        0,
        RecordKind::DgContainer.raw(),
        &nested_drawing_payload,
    );
    assert!(
        parse(&nested_drawing, &restrictive).is_err(),
        "a nested group beyond the caller's depth must remain rejected"
    );
}

#[test]
fn admitted_shapes_keep_host_loop_event_limits_typed() {
    use crate::officeart_wire::{ShapeBuilder, record_type, write_atom, write_container};

    let child = record(0, 0, 1000, &[]);
    let host_records = child.repeat(9);
    let mut shape_children = Vec::new();
    ShapeBuilder::new(1, 1)
        .with_flags(0x0800u32)
        .write(&mut shape_children, Default::default())
        .unwrap();
    write_container(
        &mut shape_children,
        0,
        record_type::CLIENT_DATA,
        &host_records,
        Default::default(),
    )
    .unwrap();

    let mut shape = Vec::new();
    write_container(
        &mut shape,
        0,
        record_type::SP_CONTAINER,
        &shape_children,
        Default::default(),
    )
    .unwrap();
    let mut drawing_body = Vec::new();
    write_atom(
        &mut drawing_body,
        0,
        0,
        record_type::DG,
        &[0; 8],
        Default::default(),
    )
    .unwrap();
    drawing_body.extend_from_slice(&shape);
    let mut drawing = Vec::new();
    write_container(
        &mut drawing,
        0,
        record_type::DG_CONTAINER,
        &drawing_body,
        Default::default(),
    )
    .unwrap();

    let limits = super::Limits {
        max_records: 8,
        ..super::Limits::default()
    };
    let shapes = parse(&drawing, &limits).expect("OfficeArt shape admission");
    assert_eq!(shapes.len(), 1);
    assert!(matches!(
        super::codec::validate_host_framing(&shapes[0]),
        Err(crate::package::Error::Limit {
            kind: crate::SourceLimitKind::Events,
            observed: 9,
            maximum: 8,
        })
    ));
}

#[test]
fn drawing_retains_root_records_alongside_typed_shapes() {
    let mut shape_atom = Vec::new();
    shape_atom.extend_from_slice(&42u32.to_le_bytes());
    shape_atom.extend_from_slice(&0x0800u32.to_le_bytes());
    let shape = record(
        0x0f,
        0,
        RecordKind::SpContainer.raw(),
        &record(2, 1, RecordKind::Sp.raw(), &shape_atom),
    );
    let future = record(2, 7, 0xf123, &[0xaa, 0xbb]);
    let dg = record(0, 0, RecordKind::Dg.raw(), &[0; 8]);
    let drawing = record(
        0x0f,
        0,
        RecordKind::DgContainer.raw(),
        &[dg, future, shape].concat(),
    );

    let parsed = parse_drawing(&drawing, &super::Limits::default()).unwrap();
    assert_eq!(parsed.records().len(), 1);
    assert_eq!(parsed.records()[0].kind(), RecordKind::DgContainer);
    assert_eq!(parsed.shapes().len(), 1);
    assert_eq!(parsed.shapes()[0].id(), 42);

    let root = Container::try_new(parsed.records()[0].clone()).unwrap();
    let found = root.find(RecordKind::Unknown(0xf123)).unwrap().unwrap();
    assert_eq!(found.version(), 2);
    assert_eq!(found.instance(), 7);
    assert_eq!(found.data(), &[0xaa, 0xbb]);
    assert!(found.data_offset(&drawing).is_some());
}

#[test]
fn ppt_client_anchor_projects_small_rect_order() {
    let mut shape_atom = Vec::new();
    shape_atom.extend_from_slice(&42u32.to_le_bytes());
    shape_atom.extend_from_slice(&0x0A00u32.to_le_bytes());
    let mut sp_record = record(2, 1, RecordKind::Sp.raw(), &shape_atom);
    let anchor_data = [
        20i16.to_le_bytes(),
        10i16.to_le_bytes(),
        110i16.to_le_bytes(),
        70i16.to_le_bytes(),
    ]
    .concat();
    sp_record.extend(record(0, 0, RecordKind::ClientAnchor.raw(), &anchor_data));
    let shape = record(0x0f, 0, RecordKind::SpContainer.raw(), &sp_record);
    let mut drawing_children = record(0, 0, RecordKind::Dg.raw(), &[0; 8]);
    drawing_children.extend(shape);
    let drawing = record(0x0f, 0, RecordKind::DgContainer.raw(), &drawing_children);

    let shapes = parse(&drawing, &super::Limits::default()).unwrap();
    let anchor = anchor(&shapes[0]).unwrap().unwrap();
    assert_eq!((anchor.left(), anchor.top()), (10, 20));
    assert_eq!((anchor.right(), anchor.bottom()), (110, 70));
    assert_eq!((anchor.width(), anchor.height()), (100, 50));
}
