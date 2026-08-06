use super::*;
use crate::animation::diagram_build::{self, Atom, Build as BuildAtom, BuildType, Container};
use crate::consts::RecordType;
use crate::diagram::Id;

fn record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + payload.len());
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&kind.to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

fn drawing() -> Vec<u8> {
    let patriarch_sp = record(
        2,
        0,
        0xF00A,
        &[1u32.to_le_bytes(), 0x0005u32.to_le_bytes()].concat(),
    );
    let bounds = record(
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
    let patriarch = record(0x0F, 0, 0xF004, &[bounds.clone(), patriarch_sp].concat());
    let group_sp = record(
        2,
        0,
        0xF00A,
        &[42u32.to_le_bytes(), 0x0001u32.to_le_bytes()].concat(),
    );
    let group_header = record(0x0F, 0, 0xF004, &[bounds, group_sp].concat());
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

fn build(build_id: u32, shape_id: u32, mode: BuildType) -> Container {
    Container::new(
        BuildAtom::new(build_id, shape_id, true, false),
        Atom::new(mode),
    )
    .expect("diagram build fixture")
}

fn build_list() -> Vec<u8> {
    let child = build(7, 42, BuildType::AllAtOnce).to_bytes();
    record(0x0F, 0, RecordType::BuildList.as_u16(), &child)
}

fn slide() -> Vec<u8> {
    let build_list = build_list();
    let opaque_before = record(0, 0, 0x7FF0, &[0xA1, 0xB2, 0xC3]);
    let opaque_inside = record(0, 0, 0x7FF1, &[0xD4, 0xE5]);
    let opaque_after = record(0, 0, 0x7FF2, &[0xF6, 0x07]);
    let tag_name = record(
        0,
        0,
        RecordType::CString.as_u16(),
        &"___PPT10"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>(),
    );
    let binary_data = record(
        0,
        0,
        RecordType::BinaryTagData.as_u16(),
        &[opaque_inside, build_list, opaque_after].concat(),
    );
    let binary_tag = record(
        0x0F,
        0,
        RecordType::ProgBinaryTag.as_u16(),
        &[tag_name, binary_data].concat(),
    );
    let tags = record(0x0F, 0, RecordType::ProgTags.as_u16(), &binary_tag);
    let drawing = record(0x0F, 0, RecordType::PPDrawing.as_u16(), &drawing());
    record(
        0x0F,
        0,
        RecordType::Slide.as_u16(),
        &[opaque_before, drawing, tags].concat(),
    )
}

#[test]
fn no_op_publication_replays_the_exact_slide_envelope() {
    let source = slide();
    let snapshot = SlideSnapshot::parse(&source).expect("slide snapshot");
    assert_eq!(snapshot.len(), 1);
    assert_eq!(
        snapshot.get(Id::new(7, 42)).unwrap().mode(),
        BuildType::AllAtOnce
    );

    let commit = snapshot.edit().commit().expect("no-op commit");
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().bytes(), source.as_slice());
    assert_eq!(
        commit.patch().apply(&snapshot).unwrap().bytes(),
        source.as_slice()
    );
}

#[test]
fn mode_publication_changes_only_the_fixed_build_field() {
    let source = slide();
    let snapshot = SlideSnapshot::parse(&source).expect("slide snapshot");
    let original_drawing = snapshot.drawing().to_vec();
    let original_prefix = source[..snapshot.build_range.start].to_vec();
    let original_suffix = source[snapshot.build_range.end..].to_vec();

    let mut editor = snapshot.edit();
    editor
        .set_mode(Id::new(7, 42), BuildType::DepthByNode)
        .expect("mode edit");
    let commit = editor.commit().expect("publication commit");
    assert!(!commit.patch().is_empty());
    assert_eq!(commit.snapshot().drawing(), original_drawing.as_slice());
    assert_eq!(
        &commit.snapshot().bytes()[..snapshot.build_range.start],
        original_prefix
    );
    assert_eq!(
        &commit.snapshot().bytes()[snapshot.build_range.end..],
        original_suffix
    );
    assert_eq!(
        commit.snapshot().get(Id::new(7, 42)).unwrap().mode(),
        BuildType::DepthByNode
    );
    assert_eq!(commit.patch().before(), source.as_slice());
    assert_eq!(commit.patch().after(), commit.snapshot().bytes());
}

#[test]
fn publication_is_source_checked_and_supports_inverse_replay() {
    let source = slide();
    let snapshot = SlideSnapshot::parse(&source).expect("slide snapshot");
    let mut editor = snapshot.edit();
    editor
        .set_mode(Id::new(7, 42), BuildType::Down)
        .expect("mode edit");
    let commit = editor.commit().expect("publication commit");

    let target = commit.patch().apply(&snapshot).expect("forward patch");
    assert_eq!(target.bytes(), commit.snapshot().bytes());
    assert_eq!(
        commit.patch().undo(&target).unwrap().bytes(),
        source.as_slice()
    );
    assert_eq!(
        commit.patch().redo(&snapshot).unwrap().bytes(),
        target.bytes()
    );

    let mut stale_bytes = source.clone();
    let last = stale_bytes.len() - 1;
    stale_bytes[last] ^= 1;
    let stale = SlideSnapshot::parse(&stale_bytes).expect("stale valid slide");
    assert!(commit.patch().apply(&stale).is_err());
}

#[test]
fn invalid_shape_edit_is_failure_atomic_at_the_publication_boundary() {
    let source = slide();
    let snapshot = SlideSnapshot::parse(&source).expect("slide snapshot");
    let mut editor = snapshot.edit();
    assert!(editor.set_shape_id(Id::new(7, 42), 999).is_err());
    assert!(!editor.is_changed());
    assert_eq!(
        editor.commit().unwrap().snapshot().bytes(),
        source.as_slice()
    );
}

#[test]
fn malformed_or_ambiguous_owning_envelopes_are_rejected() {
    let source = slide();
    assert!(SlideSnapshot::parse(&source[..source.len() - 1]).is_err());

    let mut duplicate = source.clone();
    duplicate.extend_from_slice(&[0]);
    assert!(SlideSnapshot::parse(&duplicate).is_err());

    let no_slide = record(0x0F, 0, RecordType::Document.as_u16(), &source[8..]);
    assert!(SlideSnapshot::parse(no_slide).is_err());

    let malformed_build = record(0, 0, RecordType::BuildList.as_u16(), &[0; 3]);
    let _ = diagram_build::parse_bytes(&malformed_build).is_err();
}
