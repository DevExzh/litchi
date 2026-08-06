use super::*;

use crate::writer::shapes::shape_type;
use litchi_odraw::{Container, Parser, RecordKind};

#[test]
fn nested_group_emits_ms_odraw_shape_topology_and_ids() {
    let child = UserShapeData {
        shape_type: shape_type::RECTANGLE,
        fill_color: Some(0x0000FF00),
        ..Default::default()
    };
    let nested_child = GroupShape::nested(
        14,
        ChildAnchor {
            left: 400,
            top: 500,
            right: 800,
            bottom: 900,
        },
        EscherSpgrData {
            left: 0,
            top: 0,
            right: 400,
            bottom: 400,
        },
    )
    .with_shape(
        15,
        ChildAnchor {
            left: 20,
            top: 30,
            right: 200,
            bottom: 300,
        },
        UserShapeData::default(),
    );
    let nested = GroupShape::nested(
        12,
        ChildAnchor {
            left: 10,
            top: 20,
            right: 110,
            bottom: 120,
        },
        EscherSpgrData {
            left: 0,
            top: 0,
            right: 1000,
            bottom: 1000,
        },
    )
    .with_shape(
        13,
        ChildAnchor {
            left: 100,
            top: 200,
            right: 400,
            bottom: 500,
        },
        child,
    )
    .with_group(nested_child);
    let group = GroupShape::new(
        10,
        EscherSpgrData {
            left: 0,
            top: 0,
            right: 2000,
            bottom: 2000,
        },
    )
    .with_shape(
        11,
        ChildAnchor {
            left: 50,
            top: 60,
            right: 350,
            bottom: 360,
        },
        UserShapeData::default(),
    )
    .with_group(nested);

    let bytes = create_dg_container_with_group(2, &group).expect("group drawing");
    let view = EscherDrawing::parse(&bytes).expect("wire drawing");
    view.validate_shapes().expect("valid group topology");
    assert_eq!(
        view.shape_ids().expect("shape IDs"),
        vec![10, 11, 12, 13, 14, 15]
    );

    let root = Parser::new(&bytes)
        .root()
        .expect("parse root")
        .expect("root container");
    assert_eq!(root.record().kind(), RecordKind::DgContainer);
    let groups = root
        .find_all(RecordKind::SpgrContainer)
        .expect("group container");
    assert_eq!(groups.len(), 1);
    let group_container = Container::try_new(groups[0].clone()).expect("group");
    let children: Vec<_> = group_container
        .children()
        .collect::<Result<_, _>>()
        .expect("group children");
    assert_eq!(children.len(), 3);
    assert_eq!(children[0].kind(), RecordKind::SpContainer);
    assert_eq!(children[1].kind(), RecordKind::SpContainer);
    assert_eq!(children[2].kind(), RecordKind::SpgrContainer);
}

#[test]
fn group_validation_rejects_duplicate_shape_ids() {
    let shape = UserShapeData::default();
    let group = GroupShape::new(10, EscherSpgrData::ZERO)
        .with_shape(
            11,
            ChildAnchor {
                left: 0,
                top: 0,
                right: 1,
                bottom: 1,
            },
            shape.clone(),
        )
        .with_shape(
            11,
            ChildAnchor {
                left: 1,
                top: 1,
                right: 2,
                bottom: 2,
            },
            shape,
        );

    let error = create_group_shape_container(&group).expect_err("duplicate IDs must fail");
    assert_eq!(error.kind(), std::io::ErrorKind::InvalidInput);
    assert!(error.to_string().contains("unique"));
}

#[test]
fn group_validation_rejects_anchorless_nested_group() {
    let nested = GroupShape::new(11, EscherSpgrData::ZERO);
    let group = GroupShape::new(10, EscherSpgrData::ZERO).with_group(nested);

    let error = create_group_shape_container(&group).expect_err("nested anchor must be present");
    assert!(error.to_string().contains("ChildAnchor"));
}
