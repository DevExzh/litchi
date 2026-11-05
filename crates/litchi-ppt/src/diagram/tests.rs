#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::animation::diagram_build::{self, Atom, Build as BuildAtom, BuildType, Container, Kind};
use crate::consts::RecordType;

fn record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + payload.len());
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
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
    let root_group = record(0x0F, 0, 0xF003, &[patriarch, group].concat());
    record(0x0F, 0, 0xF002, &[dg, root_group].concat())
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
    let inventory = parse_bytes(
        &build_list(&[build(9, 42, BuildType::AllAtOnce)]),
        &bytes,
        &Limits::default(),
    )
    .unwrap();
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
    let inventory = parse_bytes(&build_list(&[build]), &bytes, &Limits::default()).unwrap();
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
    assert!(parse_bytes(&duplicate, &bytes, &Limits::default()).is_err());

    let missing = build_list(&[build(1, 999, BuildType::AsOneObject)]);
    assert!(parse_bytes(&missing, &bytes, &Limits::default()).is_err());
}

#[test]
fn rejects_malformed_build_list_and_limit_overflow() {
    let bytes = drawing_with_group();
    let valid = build_list(&[build(1, 42, BuildType::AsOneObject)]);
    assert!(parse_bytes(&valid[..valid.len() - 1], &bytes, &Limits::default()).is_err());
    assert!(
        codec::parse_bytes_with_limits(
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
    assert!(parse_bytes(&value, &[0; 7], &Limits::default()).is_err());
}

#[test]
fn transaction_noop_replays_the_exact_build_list() {
    let drawing = drawing_with_group();
    let source = build_list(&[build(9, 42, BuildType::AllAtOnce)]);
    let snapshot = Snapshot::parse(&source, &drawing, &EditLimits::default()).unwrap();

    let commit = snapshot.edit().commit().unwrap();
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().bytes(), source.as_slice());
    assert_eq!(
        commit.patch().apply(&snapshot).unwrap().bytes(),
        source.as_slice()
    );
}

#[test]
fn transaction_mode_edit_changes_only_the_fixed_mode_field() {
    let drawing = drawing_with_group();
    let source = build_list(&[build(9, 42, BuildType::AllAtOnce)]);
    let snapshot = Snapshot::parse(&source, &drawing, &EditLimits::default()).unwrap();
    let id = Id::new(9, 42);

    let mut edit = snapshot.edit();
    edit.set_mode(id, BuildType::DepthByNode).unwrap();
    let commit = edit.commit().unwrap();
    let after = commit.snapshot().bytes();
    assert_eq!(&source[..48], &after[..48]);
    assert_eq!(&source[52..], &after[52..]);
    assert_eq!(&after[48..52], &1u32.to_le_bytes());
    assert_eq!(
        commit.snapshot().get(id).unwrap().mode(),
        BuildType::DepthByNode
    );
    assert_eq!(commit.patch().changes().len(), 1);
}

#[test]
fn transaction_shape_edit_checks_the_officeart_graph_and_identity() {
    let drawing = drawing_with_group();
    let source = build_list(&[
        build(9, 42, BuildType::AllAtOnce),
        build(9, 77, BuildType::Down),
    ]);
    let snapshot = Snapshot::parse(&source, &drawing, &EditLimits::default()).unwrap();
    let mut edit = snapshot.edit();

    assert!(edit.set_shape_id(Id::new(9, 42), 999).is_err());
    assert!(edit.changes().is_empty());
    assert!(edit.set_shape_id(Id::new(9, 42), 77).is_err());
    assert!(edit.changes().is_empty());
}

#[test]
fn transaction_shape_edit_updates_to_an_existing_shape() {
    let drawing = drawing_with_group();
    let source = build_list(&[build(9, 42, BuildType::AllAtOnce)]);
    let snapshot = Snapshot::parse(&source, &drawing, &EditLimits::default()).unwrap();
    let mut edit = snapshot.edit();
    let target = edit.set_shape_id(Id::new(9, 42), 77).unwrap();
    assert_eq!(target, Id::new(9, 77));
    let commit = edit.commit().unwrap();
    assert_eq!(commit.snapshot().get(target).unwrap().shape_id(), 77);
    let after = commit.snapshot().bytes();
    assert_eq!(&source[..32], &after[..32]);
    assert_eq!(&source[36..], &after[36..]);
    assert_eq!(&after[32..36], &77u32.to_le_bytes());
}

#[test]
fn transaction_preserves_unknown_kind_reserved_bytes_and_opaque_children() {
    let drawing = drawing_with_group();
    let diagram = build(9, 42, BuildType::Unknown(0xDEAD_BEEF)).to_bytes();
    let opaque = record(0, 0, 0x7FFE, &[0xA1, 0xB2, 0xC3, 0xD4]);
    let mut source = record(
        0,
        0,
        RecordType::BuildList.as_u16(),
        &[diagram.as_slice(), opaque.as_slice()].concat(),
    );
    source[8 + 16..8 + 20].copy_from_slice(&0xCAFE_BABEu32.to_le_bytes());
    source[8 + 30..8 + 32].copy_from_slice(&[0xA5, 0x5A]);
    let snapshot = Snapshot::parse(&source, &drawing, &EditLimits::default()).unwrap();
    let before = snapshot.get(Id::new(9, 42)).unwrap();
    assert_eq!(before.record().build().kind(), Kind::Unknown(0xCAFE_BABE));
    assert_eq!(before.record().build().reserved(), [0xA5, 0x5A]);

    let mut edit = snapshot.edit();
    edit.set_mode(Id::new(9, 42), BuildType::Down).unwrap();
    let target = edit.commit().unwrap().snapshot().clone();
    assert_eq!(
        target.get(Id::new(9, 42)).unwrap().record().build().kind(),
        Kind::Unknown(0xCAFE_BABE)
    );
    assert_eq!(
        target
            .get(Id::new(9, 42))
            .unwrap()
            .record()
            .build()
            .reserved(),
        [0xA5, 0x5A]
    );
    assert_eq!(
        &target.bytes()[diagram.len() + 8..],
        &source[diagram.len() + 8..]
    );
}

#[test]
fn transaction_rejects_stale_sources_and_supports_inverse_replay() {
    let drawing = drawing_with_group();
    let source = build_list(&[build(9, 42, BuildType::AllAtOnce)]);
    let snapshot = Snapshot::parse(&source, &drawing, &EditLimits::default()).unwrap();
    let mut edit = snapshot.edit();
    edit.set_mode(Id::new(9, 42), BuildType::DepthByNode)
        .unwrap();
    let commit = edit.commit().unwrap();

    let mut other = snapshot.edit();
    other.set_mode(Id::new(9, 42), BuildType::Down).unwrap();
    let other_commit = other.commit().unwrap();
    assert!(commit.patch().apply(other_commit.snapshot()).is_err());

    let target = commit.patch().apply(&snapshot).unwrap();
    let restored = commit.patch().undo(&target).unwrap();
    assert_eq!(restored.bytes(), snapshot.bytes());
    assert_eq!(
        commit.patch().redo(&restored).unwrap().bytes(),
        target.bytes()
    );
}

/// A built two-textbox package whose first slide carries a `___PPT10`
/// `BuildList` with one diagram build on its first shape, and the ids of
/// the slide's `OfficeArt` shapes.
fn package_with_diagram() -> (crate::Package, Vec<u8>, Vec<u32>) {
    use litchi_core::Position;

    use crate::record_wire as wire;
    use crate::slide_order::records::{Owner, Rewrite, publish};

    let mut builder = crate::Builder::new();
    let slide = builder.add_slide().unwrap();
    builder.add_textbox(slide, 10, 10, 200, 40, "One").unwrap();
    builder.add_textbox(slide, 10, 80, 200, 40, "Two").unwrap();
    let package = builder.build().unwrap();
    let base = crate::scoped::Base::capture(&package).unwrap();
    let native = base.native().unwrap();
    let owner = Owner::Slide(Position::new(0));
    let before = base.owner_record(&native, owner).unwrap();
    let children = wire::child_list(&before).unwrap();
    let drawing = children
        .iter()
        .find(|child| wire::record_type(child).unwrap() == RecordType::PPDrawing.as_u16())
        .expect("the built slide has a drawing");
    let parsed =
        crate::odraw::parse_drawing(&drawing[8..], &crate::odraw::Limits::default()).unwrap();
    let mut shape_ids = validation::shape_ids(&parsed, usize::MAX).unwrap();
    shape_ids.sort_unstable();
    let target = *shape_ids.last().unwrap();
    let list = record(
        0x0F,
        0,
        RecordType::BuildList.as_u16(),
        &build(7, target, BuildType::AllAtOnce).to_bytes(),
    );
    let blob =
        crate::binary_tag_edit::locate(&before, RecordType::Slide.as_u16(), "___PPT10", "diagram")
            .unwrap();
    let mut records = blob
        .records()
        .iter()
        .map(|record| record.to_vec())
        .collect::<Vec<_>>();
    records.push(list);
    let after = blob.rebuild(&before, &records, &[]).unwrap();
    let published = publish(
        &native,
        &[Rewrite {
            owner,
            before: before.into(),
            after: after.into(),
        }],
    )
    .unwrap();
    let bytes = published.bytes().to_vec();
    (
        crate::Package::from_bytes(&bytes).unwrap(),
        bytes,
        shape_ids,
    )
}

#[test]
fn package_diagram_edits_publish_replay_and_invert_exactly() {
    use litchi_core::Position;
    use litchi_core::edit::RefusalKind;

    let (source, source_bytes, shape_ids) = package_with_diagram();
    let slide = Position::new(0);
    let id = source
        .edit_diagrams(slide)
        .unwrap()
        .builds()
        .unwrap()
        .next()
        .unwrap()
        .id();
    assert_eq!(id.build_id(), 7);

    let no_op = source.edit_diagrams(slide).unwrap().commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_diagrams(slide).unwrap();
    edit.set_mode(id, BuildType::Down).unwrap();
    drop(edit);
    assert_eq!(source.to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_diagrams(slide).unwrap();
    let Err(Error::Refused(refusal)) = edit.set_mode(Id::new(99, 1), BuildType::Down) else {
        panic!("a missing build must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
    let Err(Error::Refused(refusal)) = edit.set_shape_id(id, 0xFFFF_FFF0) else {
        panic!("a missing shape must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::InvalidValue);
    assert!(edit.changes().unwrap().is_empty());
    edit.set_mode(id, BuildType::Down).unwrap();
    let other = shape_ids[shape_ids.len() - 2];
    let moved = edit.set_shape_id(id, other).unwrap();
    assert_eq!(moved, Id::new(7, other));
    assert_eq!(edit.changes().unwrap().len(), 2);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().source_len(), source_bytes.len());
    assert_eq!(
        commit.patch().target_len(),
        commit.package().to_bytes().unwrap().len()
    );
    let published = commit.package().edit_diagrams(slide).unwrap();
    let build = published.builds().unwrap().next().unwrap();
    assert_eq!(build.id(), moved);
    assert_eq!(build.mode(), BuildType::Down);
    assert_eq!(
        source
            .edit_diagrams(slide)
            .unwrap()
            .builds()
            .unwrap()
            .next()
            .unwrap()
            .mode(),
        BuildType::AllAtOnce
    );

    let replayed = source.apply_diagrams(commit.patch()).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    let restored = commit
        .package()
        .apply_diagrams(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
    assert!(matches!(
        commit.package().apply_diagrams(commit.patch()),
        Err(Error::PatchConflict)
    ));

    let Err(Error::Refused(refusal)) = source
        .edit_diagrams(Position::new(1))
        .and_then(|edit| edit.builds().map(|_| ()))
    else {
        panic!("a missing slide must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
    let mut builder = crate::Builder::new();
    builder.add_slide().unwrap();
    let plain = builder.build().unwrap();
    let Err(Error::Refused(refusal)) = plain
        .edit_diagrams(slide)
        .and_then(|edit| edit.builds().map(|_| ()))
    else {
        panic!("a slide without a build list must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::UnsupportedTarget);
}

#[test]
fn package_diagram_edits_admit_a_signed_package_and_refuse_a_changed_commit() {
    use litchi_core::Position;

    let (source, _, _) = package_with_diagram();
    let signed = crate::scoped::signed(&source);
    let bytes = signed.to_bytes().unwrap();
    let slide = Position::new(0);
    let id = signed
        .edit_diagrams(slide)
        .unwrap()
        .builds()
        .unwrap()
        .next()
        .unwrap()
        .id();

    let no_op = signed.edit_diagrams(slide).unwrap().commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), bytes);

    let mut edit = signed.edit_diagrams(slide).unwrap();
    edit.set_mode(id, BuildType::Down).unwrap();
    assert!(matches!(edit.commit(), Err(Error::Signed)));
    assert_eq!(signed.to_bytes().unwrap(), bytes);
}
