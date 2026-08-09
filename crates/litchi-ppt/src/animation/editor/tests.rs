#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::super::{
    AnimationInfo, ExtendedTimeNode, TimeNodeAtom, TimeNodeKind, parse_slide_animation_extension,
    write_animation_info, write_extended_time_node,
};
use super::semantic::{ESCHER_CLIENT_DATA, ESCHER_SP, ESCHER_SP_CONTAINER, EditorLimits, Scope};
use super::transaction::{
    atom, collect_shapes_and_legacy, container, escher_record, rewrite_extension_payload,
    rewrite_shape_animation,
};
use super::validation::validate_node;
use crate::consts::RecordType;
use crate::records::Record;
use std::collections::BTreeSet;

#[test]
fn ppt10_payload_replacement_preserves_unknown_records() {
    let unknown = atom(0, 0, 0x7777, b"opaque").unwrap();
    let old = write_extended_time_node(&ExtendedTimeNode::default()).unwrap();
    let mut payload = unknown.clone();
    payload.extend(old);
    let rewritten = rewrite_extension_payload(
        Some(&payload),
        Some(&ExtendedTimeNode {
            atom: TimeNodeAtom {
                node_type: Some(TimeNodeKind::Sequential),
                duration_ms: Some(500),
                ..Default::default()
            },
            ..Default::default()
        }),
        None,
    )
    .unwrap();
    assert!(
        rewritten
            .windows(unknown.len())
            .any(|value| value == unknown)
    );
    assert_eq!(
        parse_slide_animation_extension(&rewritten)
            .unwrap()
            .time_node
            .unwrap()
            .atom
            .duration_ms,
        Some(500)
    );
}

#[test]
fn nested_timeline_limits_and_bad_reorders_roll_back_before_staging() {
    let mut node = ExtendedTimeNode::default();
    node.children.push(ExtendedTimeNode::default());
    let mut count = 0;
    assert!(
        validate_node(
            &node,
            1,
            &mut count,
            &BTreeSet::new(),
            EditorLimits {
                max_timeline_depth: 1,
                ..Default::default()
            }
        )
        .is_err()
    );
}

#[test]
fn malformed_ppt10_payload_is_rejected() {
    assert!(rewrite_extension_payload(Some(&[0; 7]), None, None).is_err());
}

#[test]
fn legacy_shape_edit_preserves_inert_interactive_records() {
    let interactive = atom(0, 0, RecordType::InteractiveInfoAtom.as_u16(), &[0; 16]).unwrap();
    let animation = write_animation_info(&AnimationInfo::new()).unwrap().0;
    let mut client_payload = interactive.clone();
    client_payload.extend(animation);
    let client = escher_record(0x0f, 0, ESCHER_CLIENT_DATA, &client_payload).unwrap();
    let mut shape_payload = escher_record(2, 0, ESCHER_SP, &[42, 0, 0, 0, 0, 0, 0, 0]).unwrap();
    shape_payload.extend(client);
    let shape = escher_record(0x0f, 0, ESCHER_SP_CONTAINER, &shape_payload).unwrap();
    let drawing = atom(0, 0, RecordType::PPDrawing.as_u16(), &shape).unwrap();
    let slide = container(0, RecordType::Slide.as_u16(), &drawing).unwrap();

    let (rewritten, found) = rewrite_shape_animation(&slide, 42, None).unwrap();
    assert!(found);
    assert!(
        rewritten
            .windows(interactive.len())
            .any(|value| value == interactive)
    );
    let (record, used) = Record::parse(&rewritten, 0).unwrap();
    assert_eq!(used, rewritten.len());
    let (shapes, legacy) =
        collect_shapes_and_legacy(1, Scope::Slide, &record, EditorLimits::default()).unwrap();
    assert!(shapes.contains(&42));
    assert!(legacy.is_empty());
}
