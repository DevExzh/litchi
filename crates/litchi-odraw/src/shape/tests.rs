#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{Bounds, Flags, Kind, Native, parse, parse_with};
use crate::write::{self, Atom, Container as OutContainer, ShapeBuilder};
use crate::{Error, Limit, Limits, Record, RecordKind};

fn shape(kind: Native, id: u32, child: bool) -> Vec<u8> {
    let mut body = Vec::new();
    let mut flags = Flags::HAVE_ANCHOR | Flags::HAVE_SPT;
    if child {
        flags |= Flags::CHILD;
    }
    ShapeBuilder::new(kind, id)
        .with_flags(flags)
        .write(&mut body)
        .expect("write shape atom");
    if child {
        write::child_anchor(&mut body, 10, 20, 110, 70).expect("write child anchor");
    } else {
        let mut payload = Vec::new();
        for coordinate in [10_i32, 20, 110, 70] {
            payload.extend_from_slice(&coordinate.to_le_bytes());
        }
        write::atom(&mut body, 0, Atom::ClientAnchor, &payload).expect("write opaque host anchor");
    }
    let mut record = Vec::new();
    write::container(&mut record, 0, OutContainer::Sp, &body).expect("write shape container");
    record
}

fn patriarch(id: u32) -> Vec<u8> {
    let mut body = Vec::new();
    write::spgr(&mut body, 0, 0, 0, 0).expect("write patriarch bounds");
    ShapeBuilder::new(Native::FREEFORM, id)
        .with_flags(Flags::GROUP | Flags::PATRIARCH)
        .write(&mut body)
        .expect("write patriarch shape atom");
    let mut record = Vec::new();
    write::container(&mut record, 0, OutContainer::Sp, &body).expect("write patriarch container");
    record
}

fn background(id: u32) -> Vec<u8> {
    let mut body = Vec::new();
    ShapeBuilder::new(Native::RECTANGLE, id)
        .with_flags(Flags::BACKGROUND | Flags::HAVE_ANCHOR | Flags::HAVE_SPT)
        .write(&mut body)
        .expect("write background shape atom");
    let mut record = Vec::new();
    write::container(&mut record, 0, OutContainer::Sp, &body).expect("write background container");
    record
}

fn drawing() -> Vec<u8> {
    let patriarch = patriarch(1);
    let rectangle = shape(Native::RECTANGLE, 2, false);

    let mut group_header_body = Vec::new();
    write::spgr(&mut group_header_body, 0, 0, 1000, 500).expect("write group bounds");
    let future = Atom::unknown(0xF123, 0).expect("future atom kind");
    write::atom(&mut group_header_body, 0, future, &[0xAA, 0xBB]).expect("write future atom");
    ShapeBuilder::new(Native::FREEFORM, 3)
        .with_flags(Flags::GROUP | Flags::HAVE_ANCHOR)
        .write(&mut group_header_body)
        .expect("write group shape");
    let mut group_anchor = Vec::new();
    for coordinate in [100_i32, 200, 500, 400] {
        group_anchor.extend_from_slice(&coordinate.to_le_bytes());
    }
    write::atom(&mut group_header_body, 0, Atom::ClientAnchor, &group_anchor)
        .expect("write group anchor");
    let mut group_header = Vec::new();
    write::container(&mut group_header, 0, OutContainer::Sp, &group_header_body)
        .expect("write group header");

    let mut nested_body = group_header;
    nested_body.extend_from_slice(&shape(Native::ELLIPSE, 4, true));
    let mut nested = Vec::new();
    write::container(&mut nested, 0, OutContainer::Spgr, &nested_body).expect("write nested group");

    let mut root_body = patriarch;
    root_body.extend_from_slice(&rectangle);
    root_body.extend_from_slice(&nested);
    let mut root = Vec::new();
    write::container(&mut root, 0, OutContainer::Spgr, &root_body).expect("write root group");

    let mut drawing_body = Vec::new();
    write::dg(&mut drawing_body, 5, 5).expect("write drawing atom");
    drawing_body.extend_from_slice(&root);
    drawing_body.extend_from_slice(&background(5));
    let mut bytes = Vec::new();
    write::container(&mut bytes, 0, OutContainer::Dg, &drawing_body).expect("write drawing");
    bytes
}

#[test]
fn hides_root_patriarch_and_preserves_nested_group() {
    let bytes = drawing();
    let shapes = parse(&bytes).expect("parse drawing");

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
        .write(&mut body)
        .expect("write shape atom");
    let mut bytes = Vec::new();
    write::container(&mut bytes, 0, OutContainer::Sp, &body).expect("write shape container");

    assert!(matches!(
        parse(&bytes),
        Err(Error::MalformedShape {
            reason: "shape HAVE_ANCHOR flag disagrees with its anchor records",
        })
    ));
}

#[test]
fn traversal_limits_are_enforced() {
    let error = parse_with(
        &drawing(),
        Limits {
            max_depth: 0,
            max_records: 1_000,
        },
    )
    .expect_err("nested group exceeds depth zero");

    assert!(matches!(
        error,
        Error::LimitExceeded {
            limit: Limit::Depth,
            ..
        }
    ));
}

#[test]
fn rejects_trailing_root_bytes() {
    let mut bytes = drawing();
    bytes.push(0);

    assert!(matches!(parse(&bytes), Err(Error::TrailingData { .. })));
}

#[test]
fn rejects_an_unsafe_recursive_depth_limit() {
    assert!(matches!(
        parse_with(
            &drawing(),
            Limits {
                max_depth: 65,
                max_records: 1_000,
            },
        ),
        Err(Error::InvalidLimit {
            limit: Limit::Depth,
            maximum: 64,
        })
    ));
}

#[test]
fn metadata_records_consume_the_record_budget() {
    assert!(matches!(
        parse_with(
            &drawing(),
            Limits {
                max_depth: 64,
                max_records: 5,
            },
        ),
        Err(Error::LimitExceeded {
            limit: Limit::Records,
            maximum: 5,
        })
    ));
}

#[test]
fn rejects_a_group_coordinate_atom_with_the_wrong_length() {
    let data = [0_u8; 12];
    let record = Record::from_parts(RecordKind::Spgr, 1, 0, &data).expect("test record");

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
    let shapes = parse(&bytes).expect("parse drawing");
    let group = &shapes[1];
    let extension = group
        .meta()
        .find(RecordKind::Unknown(0xF123))
        .expect("scan unknown record")
        .expect("future record");

    assert_eq!(extension.data(), &[0xAA, 0xBB]);
    assert!(extension.data_offset(&bytes).is_some());
}
