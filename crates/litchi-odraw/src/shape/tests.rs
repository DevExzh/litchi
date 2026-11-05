#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{Bounds, Flags, Kind, Native, Shape, parse};
use crate::builder::ShapeBuilder;
use crate::write::{self, Atom, Container as OutContainer};
use crate::{Error, LimitKind, Limits, Record, RecordKind};

fn shape(kind: Native, id: u32, child: bool) -> Vec<u8> {
    let mut body = Vec::new();
    let mut flags = Flags::HAVE_ANCHOR | Flags::HAVE_SPT;
    if child {
        flags |= Flags::CHILD;
    }
    ShapeBuilder::new(kind, id)
        .with_flags(flags)
        .write(&mut body, Limits::default())
        .expect("write shape atom");
    if child {
        write::child_anchor(&mut body, 10, 20, 110, 70, Limits::default())
            .expect("write child anchor");
    } else {
        let mut payload = Vec::new();
        for coordinate in [10_i32, 20, 110, 70] {
            payload.extend_from_slice(&coordinate.to_le_bytes());
        }
        write::atom(
            &mut body,
            0,
            Atom::ClientAnchor,
            &payload,
            Limits::default(),
        )
        .expect("write opaque host anchor");
    }
    let mut record = Vec::new();
    write::container(&mut record, 0, OutContainer::Sp, &body, Limits::default())
        .expect("write shape container");
    record
}

fn patriarch(id: u32) -> Vec<u8> {
    let mut body = Vec::new();
    write::spgr(&mut body, 0, 0, 0, 0, Limits::default()).expect("write patriarch bounds");
    ShapeBuilder::new(Native::FREEFORM, id)
        .with_flags(Flags::GROUP | Flags::PATRIARCH)
        .write(&mut body, Limits::default())
        .expect("write patriarch shape atom");
    let mut record = Vec::new();
    write::container(&mut record, 0, OutContainer::Sp, &body, Limits::default())
        .expect("write patriarch container");
    record
}

fn background(id: u32) -> Vec<u8> {
    let mut body = Vec::new();
    ShapeBuilder::new(Native::RECTANGLE, id)
        .with_flags(Flags::BACKGROUND | Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
        .write(&mut body, Limits::default())
        .expect("write background shape atom");
    let mut record = Vec::new();
    write::container(&mut record, 0, OutContainer::Sp, &body, Limits::default())
        .expect("write background container");
    record
}

fn drawing() -> Vec<u8> {
    let patriarch = patriarch(1);
    let rectangle = shape(Native::RECTANGLE, 2, false);

    let mut group_header_body = Vec::new();
    write::spgr(&mut group_header_body, 0, 0, 1000, 500, Limits::default())
        .expect("write group bounds");
    let future = Atom::unknown(0xF123, 0).expect("future atom kind");
    write::atom(
        &mut group_header_body,
        0,
        future,
        &[0xAA, 0xBB],
        Limits::default(),
    )
    .expect("write future atom");
    ShapeBuilder::new(Native::FREEFORM, 3)
        .with_flags(Flags::GROUP | Flags::HAVE_ANCHOR)
        .write(&mut group_header_body, Limits::default())
        .expect("write group shape");
    let mut group_anchor = Vec::new();
    for coordinate in [100_i32, 200, 500, 400] {
        group_anchor.extend_from_slice(&coordinate.to_le_bytes());
    }
    write::atom(
        &mut group_header_body,
        0,
        Atom::ClientAnchor,
        &group_anchor,
        Limits::default(),
    )
    .expect("write group anchor");
    let mut group_header = Vec::new();
    write::container(
        &mut group_header,
        0,
        OutContainer::Sp,
        &group_header_body,
        Limits::default(),
    )
    .expect("write group header");

    let mut nested_body = group_header;
    nested_body.extend_from_slice(&shape(Native::ELLIPSE, 4, true));
    let mut nested = Vec::new();
    write::container(
        &mut nested,
        0,
        OutContainer::Spgr,
        &nested_body,
        Limits::default(),
    )
    .expect("write nested group");

    let mut root_body = patriarch;
    root_body.extend_from_slice(&rectangle);
    root_body.extend_from_slice(&nested);
    let mut root = Vec::new();
    write::container(
        &mut root,
        0,
        OutContainer::Spgr,
        &root_body,
        Limits::default(),
    )
    .expect("write root group");

    let mut drawing_body = Vec::new();
    write::dg(&mut drawing_body, 5, 5, Limits::default()).expect("write drawing atom");
    drawing_body.extend_from_slice(&root);
    drawing_body.extend_from_slice(&background(5));
    let mut bytes = Vec::new();
    write::container(
        &mut bytes,
        0,
        OutContainer::Dg,
        &drawing_body,
        Limits::default(),
    )
    .expect("write drawing");
    bytes
}

fn standalone_nested_group() -> Vec<u8> {
    fn group_header(id: u32, child: bool) -> Vec<u8> {
        let mut body = Vec::new();
        write::spgr(&mut body, 0, 0, 1000, 500, Limits::default()).expect("write group bounds");
        let mut flags = Flags::GROUP | Flags::HAVE_ANCHOR;
        if child {
            flags |= Flags::CHILD;
        }
        ShapeBuilder::new(Native::FREEFORM, id)
            .with_flags(flags)
            .write(&mut body, Limits::default())
            .expect("write group shape");
        if child {
            write::child_anchor(&mut body, 10, 20, 110, 70, Limits::default())
                .expect("write child anchor");
        } else {
            let mut anchor = Vec::new();
            for coordinate in [10_i32, 20, 110, 70] {
                anchor.extend_from_slice(&coordinate.to_le_bytes());
            }
            write::atom(&mut body, 0, Atom::ClientAnchor, &anchor, Limits::default())
                .expect("write host anchor");
        }
        let mut record = Vec::new();
        write::container(&mut record, 0, OutContainer::Sp, &body, Limits::default())
            .expect("write group header");
        record
    }

    let inner_header = group_header(2, true);
    let inner_shape = shape(Native::ELLIPSE, 3, true);
    let mut inner_body = inner_header;
    inner_body.extend_from_slice(&inner_shape);
    let mut inner = Vec::new();
    write::container(
        &mut inner,
        0,
        OutContainer::Spgr,
        &inner_body,
        Limits::default(),
    )
    .expect("write nested group");

    let outer_header = group_header(1, false);
    let mut outer_body = outer_header;
    outer_body.extend_from_slice(&inner);
    let mut outer = Vec::new();
    write::container(
        &mut outer,
        0,
        OutContainer::Spgr,
        &outer_body,
        Limits::default(),
    )
    .expect("write standalone group");
    outer
}

#[test]
fn from_record_obeys_the_same_limits_as_parse() {
    let bytes = shape(Native::RECTANGLE, 2, false);
    let (record, _) = Record::parse(&bytes, 0, Limits::default()).expect("valid record");
    let parsed = Shape::from_record(record.clone(), Limits::default()).expect("valid shape");
    assert_eq!(parsed.id(), 2);
    assert_eq!(parsed.kind(), Kind::Rectangle);
    assert!(matches!(
        Shape::from_record(
            record.clone(),
            Limits {
                max_depth: 64,
                max_records: 1,
                ..Limits::default()
            }
        ),
        Err(Error::LimitExceeded {
            limit: LimitKind::Records,
            maximum: 1,
            ..
        })
    ));
    assert!(matches!(
        Shape::from_record(
            record,
            Limits {
                max_depth: 65,
                max_records: 1_000,
                ..Limits::default()
            }
        ),
        Err(Error::InvalidLimit {
            limit: LimitKind::Depth,
            maximum: 64,
            ..
        })
    ));
}

#[test]
fn from_record_applies_narrower_framing_limits_to_children() {
    let bytes = shape(Native::RECTANGLE, 2, false);
    let (record, _) = Record::parse(&bytes, 0, Limits::default()).expect("valid record");
    let record_len = record.len();
    let error = Shape::from_record(
        record,
        Limits {
            max_record_bytes: record_len - 1,
            max_depth: 64,
            max_records: 1_000,
        },
    )
    .expect_err("the root exceeds the narrower framing policy");

    assert!(matches!(
        error,
        Error::LimitExceeded {
            limit: LimitKind::RecordBytes,
            observed,
            maximum,
        } if observed == u64::from(record_len) && maximum == record_len - 1
    ));
}

#[test]
fn from_record_applies_narrower_depth_limits_to_nested_groups() {
    let bytes = standalone_nested_group();
    let (root_group, _) = Record::parse(&bytes, 0, Limits::default()).expect("valid group");

    let error = Shape::from_record(
        root_group,
        Limits {
            max_record_bytes: Limits::default().max_record_bytes,
            max_depth: 0,
            max_records: 1_000,
        },
    )
    .expect_err("the nested group exceeds the narrower depth policy");
    assert!(matches!(
        error,
        Error::LimitExceeded {
            limit: LimitKind::Depth,
            observed: 1,
            maximum: 0,
        }
    ));
}

#[test]
fn hides_root_patriarch_and_preserves_nested_group() {
    let bytes = drawing();
    let shapes = parse(&bytes, Default::default()).expect("parse drawing");

    assert_eq!(shapes.len(), 2);
    assert_eq!(shapes[0].kind(), Kind::Rectangle);
    assert_eq!(shapes[0].id(), 2);
    assert!(shapes[0].anchor().is_none());
    assert!(shapes[0].client_anchor().is_some());
    assert_eq!(shapes[1].kind(), Kind::Group);
    assert_eq!(shapes[1].id(), 3);
    assert_eq!(
        shapes[1].group_bounds().copied(),
        Some(Bounds::new(0, 0, 1000, 500))
    );
    assert_eq!(shapes[1].children()[0].kind(), Kind::Ellipse);
    assert_eq!(shapes[1].children()[0].id(), 4);
    assert!(
        shapes
            .iter()
            .all(|shape| !shape.flags().contains(Flags::BACKGROUND))
    );
}

#[test]
fn rejects_a_missing_anchor_for_a_user_shape() {
    let mut body = Vec::new();
    ShapeBuilder::new(Native::RECTANGLE, 1)
        .with_flags(Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
        .write(&mut body, Limits::default())
        .expect("write shape atom");
    let mut bytes = Vec::new();
    write::container(&mut bytes, 0, OutContainer::Sp, &body, Limits::default())
        .expect("write shape container");

    assert!(matches!(
        parse(&bytes, Default::default()),
        Err(Error::MalformedShape {
            reason: "shape HAVE_ANCHOR flag disagrees with its anchor records",
        })
    ));
}

#[test]
fn traversal_limits_are_enforced() {
    let error = parse(
        &drawing(),
        Limits {
            max_depth: 0,
            max_records: 1_000,
            ..Limits::default()
        },
    )
    .expect_err("nested group exceeds depth zero");

    assert!(matches!(
        error,
        Error::LimitExceeded {
            limit: LimitKind::Depth,
            ..
        }
    ));
}

#[test]
fn rejects_trailing_root_bytes() {
    let mut bytes = drawing();
    bytes.push(0);

    assert!(matches!(
        parse(&bytes, Default::default()),
        Err(Error::TrailingData { .. })
    ));
}

#[test]
fn rejects_an_unsafe_recursive_depth_limit() {
    assert!(matches!(
        parse(
            &drawing(),
            Limits {
                max_depth: 65,
                max_records: 1_000,
                ..Limits::default()
            },
        ),
        Err(Error::InvalidLimit {
            limit: LimitKind::Depth,
            maximum: 64,
            ..
        })
    ));
}

#[test]
fn metadata_records_consume_the_record_budget() {
    assert!(matches!(
        parse(
            &drawing(),
            Limits {
                max_depth: 64,
                max_records: 5,
                ..Limits::default()
            },
        ),
        Err(Error::LimitExceeded {
            limit: LimitKind::Records,
            maximum: 5,
            ..
        })
    ));
}

#[test]
fn rejects_a_group_coordinate_atom_with_the_wrong_length() {
    let data = [0_u8; 12];
    let record =
        Record::from_parts(RecordKind::Spgr, 1, 0, &data, Limits::default()).expect("test record");

    assert!(matches!(
        Bounds::from_record(&record),
        Err(Error::MalformedShape {
            reason: "OfficeArt atom payload length is invalid",
        })
    ));
}

#[test]
fn typed_group_bounds_do_not_drop_future_records() {
    let bytes = drawing();
    let shapes = parse(&bytes, Default::default()).expect("parse drawing");
    let group = &shapes[1];
    let extension = group
        .meta()
        .find(RecordKind::Unknown(0xF123))
        .expect("scan unknown record")
        .expect("future record");

    assert_eq!(extension.data(), &[0xAA, 0xBB]);
    assert!(extension.data_offset(&bytes).is_some());
}
