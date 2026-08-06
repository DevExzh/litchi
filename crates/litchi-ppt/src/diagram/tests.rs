use super::*;
use crate::animation::diagram_build::{self, Atom, Build as BuildAtom, BuildType, Container};
use crate::consts::RecordType;

fn record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + payload.len());
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

fn drawing_with_group() -> Vec<u8> {
    let patriarch_sp = record(
        2,
        0,
        0xF00A,
        &[1u32.to_le_bytes(), 0x0005u32.to_le_bytes()].concat(),
    );
    let patriarch_bounds = record(
        1,
        0,
        0xF009,
        &[
            0i32.to_le_bytes(),
            0i32.to_le_bytes(),
            1000i32.to_le_bytes(),
            1000i32.to_le_bytes(),
        ]
        .concat(),
    );
    let patriarch = record(0x0F, 0, 0xF004, &[patriarch_bounds, patriarch_sp].concat());

    let group_sp = record(
        2,
        0,
        0xF00A,
        &[42u32.to_le_bytes(), 0x0001u32.to_le_bytes()].concat(),
    );
    let group_bounds = record(
        1,
        0,
        0xF009,
        &[
            0i32.to_le_bytes(),
            0i32.to_le_bytes(),
            1000i32.to_le_bytes(),
            1000i32.to_le_bytes(),
        ]
        .concat(),
    );
    let group_header = record(0x0F, 0, 0xF004, &[group_bounds, group_sp].concat());

    let child_sp = record(
        2,
        1,
        0xF00A,
        &[77u32.to_le_bytes(), 0x0202u32.to_le_bytes()].concat(),
    );
    let child_anchor = record(
        0,
        0,
        0xF00F,
        &[
            10i32.to_le_bytes(),
            20i32.to_le_bytes(),
            110i32.to_le_bytes(),
            120i32.to_le_bytes(),
        ]
        .concat(),
    );
    let child = record(0x0F, 0, 0xF004, &[child_sp, child_anchor].concat());
    let group = record(0x0F, 0, 0xF003, &[group_header, child].concat());
    let dg = record(0, 0, 0xF008, &[0; 8]);
    record(0x0F, 0, 0xF002, &[dg, patriarch, group].concat())
}

fn build_list(builds: &[Container]) -> Vec<u8> {
    let children: Vec<_> = builds.iter().map(Container::to_bytes).collect();
    record(0, 0, RecordType::BuildList.as_u16(), &children.concat())
}

fn build(build_id: u32, shape_id: u32, mode: BuildType) -> Container {
    Container::new(
        BuildAtom::new(build_id, shape_id, true, false),
        Atom::new(mode),
    )
    .unwrap()
}

#[test]
fn inventory_groups_build_identity_shapes_and_inert_payloads() {
    let bytes = drawing_with_group();
    let inventory =
        parse_bytes(&build_list(&[build(9, 42, BuildType::AllAtOnce)]), &bytes).unwrap();
    assert_eq!(inventory.len(), 1);
    let diagram = &inventory.diagrams()[0];
    assert_eq!(diagram.id(), Id::new(9, 42));
    assert_eq!(diagram.build().mode(), BuildType::AllAtOnce);
    assert_eq!(
        diagram
            .shapes()
            .iter()
            .map(|shape| shape.id())
            .collect::<Vec<_>>(),
        vec![42, 77]
    );
    assert_eq!(diagram.payloads()[0].shape_id(), 42);
    assert_eq!(diagram.payloads()[0].kind(), PayloadKind::Shape);
    assert_eq!(
        diagram.payloads()[0].record_kind(),
        litchi_odraw::RecordKind::SpContainer
    );
    assert_eq!(inventory.shape(diagram.root()).unwrap().id(), 42);
}

#[test]
fn unknown_build_mode_is_retained_without_claiming_layout_support() {
    let bytes = drawing_with_group();
    let build = build(1, 42, BuildType::Unknown(0xDEAD_BEEF));
    let inventory = parse_bytes(&build_list(&[build]), &bytes).unwrap();
    assert_eq!(
        inventory.diagrams()[0].build().mode(),
        BuildType::Unknown(0xDEAD_BEEF)
    );
}

#[test]
fn rejects_duplicate_build_identity_and_missing_shape() {
    let bytes = drawing_with_group();
    let duplicate = build_list(&[
        build(1, 42, BuildType::AsOneObject),
        build(1, 42, BuildType::Down),
    ]);
    assert!(parse_bytes(&duplicate, &bytes).is_err());

    let missing = build_list(&[build(1, 999, BuildType::AsOneObject)]);
    assert!(parse_bytes(&missing, &bytes).is_err());
}

#[test]
fn rejects_malformed_build_list_and_limit_overflow() {
    let bytes = drawing_with_group();
    let valid = build_list(&[build(1, 42, BuildType::AsOneObject)]);
    assert!(parse_bytes(&valid[..valid.len() - 1], &bytes).is_err());
    assert!(
        parse_bytes_with_limits(
            &valid,
            &bytes,
            Limits {
                max_diagrams: 0,
                ..Limits::default()
            }
        )
        .is_err()
    );
}

#[test]
fn shared_record_parser_remains_the_single_diagram_wire_owner() {
    let value = build(3, 42, BuildType::Custom);
    let record = value.to_record();
    assert_eq!(diagram_build::parse_record(&record).unwrap(), value);
    assert_eq!(record.record_type, RecordType::DiagramBuild);
}

#[test]
fn malformed_drawing_is_not_silently_dropped() {
    let value = build_list(&[build(3, 42, BuildType::Custom)]);
    assert!(parse_bytes(&value, &[0; 7]).is_err());
}
