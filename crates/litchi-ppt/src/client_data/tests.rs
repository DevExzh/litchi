#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::codec::{OFFICE_ART_CLIENT_DATA_RECORD_TYPE, encode_record};
use super::model::{ClientData, ClientDataChild, ClientDataChildKind, ClientDataLimits};

fn child(version: u16, instance: u16, kind: ClientDataChildKind, data: &[u8]) -> Vec<u8> {
    encode_record(version, instance, kind.record_type(), data).unwrap()
}

fn container(payload: &[u8]) -> Vec<u8> {
    encode_record(0x0F, 0, OFFICE_ART_CLIENT_DATA_RECORD_TYPE, payload).unwrap()
}

fn opaque(version: u16, instance: u16, record_type: u16, data: &[u8]) -> Vec<u8> {
    encode_record(version, instance, record_type, data).unwrap()
}

#[test]
fn parses_accesses_and_round_trips_the_complete_ordered_grammar() {
    let mut payload = Vec::new();
    payload.extend(child(0, 0, ClientDataChildKind::ShapeFlags, &[1]));
    payload.extend(child(0, 0, ClientDataChildKind::ShapeFlags10, &[4]));
    payload.extend(child(
        0,
        0,
        ClientDataChildKind::ExternalObjectReference,
        &77u32.to_le_bytes(),
    ));
    payload.extend(child(
        0,
        0,
        ClientDataChildKind::Placeholder,
        &[3, 0, 0, 0, 13, 1, 0, 0],
    ));
    payload.extend(child(
        0x0F,
        7,
        ClientDataChildKind::ProgrammableTags,
        &[0xAA, 0xBB],
    ));
    payload.extend(child(
        0,
        0,
        ClientDataChildKind::RoundTripShapeId12,
        &0x1020u32.to_le_bytes(),
    ));
    payload.extend(child(
        0,
        0,
        ClientDataChildKind::RoundTripNewPlaceholderId12,
        &[26],
    ));
    payload.extend(child(
        0,
        0,
        ClientDataChildKind::RoundTripShapeChecksumForCustomLayouts12,
        &[1, 0, 0, 0, 2, 0, 0, 0],
    ));
    let bytes = container(&payload);

    let parsed = ClientData::parse(&bytes, &ClientDataLimits::default()).unwrap();
    assert!(parsed.shape_flags().is_some());
    assert_eq!(
        parsed
            .external_object_reference()
            .unwrap()
            .external_object_id(),
        Some(77)
    );
    assert_eq!(
        parsed
            .child(ClientDataChildKind::ProgrammableTags)
            .unwrap()
            .instance(),
        7
    );
    assert_eq!(
        parsed
            .child(ClientDataChildKind::RoundTripShapeId12)
            .unwrap()
            .round_trip_shape_id(),
        Some(0x1020)
    );
    assert_eq!(parsed.round_trip_records().count(), 4);
    assert_eq!(parsed.to_bytes().unwrap(), bytes);
}

#[test]
fn canonical_constructor_writes_valid_records() {
    let records = vec![
        ClientDataChild::new(
            ClientDataChildKind::ExternalObjectReference,
            9u32.to_le_bytes().to_vec(),
        )
        .unwrap(),
        ClientDataChild::new(
            ClientDataChildKind::RoundTripHeaderFooterPlaceholder12,
            vec![7],
        )
        .unwrap(),
    ];
    let value = ClientData::new(records).unwrap();
    assert_eq!(
        ClientData::parse(&value.to_bytes().unwrap(), &ClientDataLimits::default()).unwrap(),
        value
    );
}

#[test]
fn rejects_outer_header_length_truncation_and_trailing_data() {
    let valid = container(&[]);
    for index in [0usize, 1, 2] {
        let mut bad = valid.clone();
        bad[index] ^= 1;
        assert!(ClientData::parse(&bad, &ClientDataLimits::default()).is_err());
    }
    let mut bad_length = valid.clone();
    bad_length[4..8].copy_from_slice(&1u32.to_le_bytes());
    assert!(ClientData::parse(&bad_length, &ClientDataLimits::default()).is_err());
    let mut trailing = valid.clone();
    trailing.push(0);
    assert!(ClientData::parse(&trailing, &ClientDataLimits::default()).is_err());
    assert!(ClientData::parse(&valid[..7], &ClientDataLimits::default()).is_err());
}

#[test]
fn preserves_unknown_children_and_rejects_duplicate_and_out_of_order_known_children() {
    let flags = child(0, 0, ClientDataChildKind::ShapeFlags, &[0]);
    let placeholder = child(
        0,
        0,
        ClientDataChildKind::Placeholder,
        &[0, 0, 0, 0, 13, 0, 0, 0],
    );
    assert!(
        ClientData::parse(
            &container(&[flags.clone(), flags.clone()].concat()),
            &ClientDataLimits::default()
        )
        .is_err()
    );
    assert!(
        ClientData::parse(
            &container(&[placeholder, flags].concat()),
            &ClientDataLimits::default()
        )
        .is_err()
    );

    let unknown = opaque(7, 0x123, 0x2222, &[0xCA, 0xFE]);
    let unknown_container = container(&unknown);
    let parsed = ClientData::parse(&unknown_container, &ClientDataLimits::default()).unwrap();
    assert_eq!(parsed.children().len(), 1);
    assert_eq!(parsed.children()[0].kind(), ClientDataChildKind::Unknown);
    assert_eq!(parsed.children()[0].record_type(), 0x2222);
    assert_eq!(parsed.children()[0].version(), 7);
    assert_eq!(parsed.children()[0].instance(), 0x123);
    assert_eq!(parsed.children()[0].payload(), &[0xCA, 0xFE]);
    assert_eq!(parsed.to_bytes().unwrap(), unknown_container);

    let shape_id = child(0, 0, ClientDataChildKind::RoundTripShapeId12, &[0; 4]);
    assert!(
        ClientData::parse(
            &container(&[shape_id.clone(), shape_id].concat()),
            &ClientDataLimits::default()
        )
        .is_err()
    );
}

#[test]
fn rejects_bad_child_headers_reserved_values_and_boundaries() {
    assert!(
        ClientData::parse(
            &container(&child(1, 0, ClientDataChildKind::ShapeFlags, &[0])),
            &ClientDataLimits::default()
        )
        .is_err()
    );
    assert!(
        ClientData::parse(
            &container(&child(0, 1, ClientDataChildKind::ShapeFlags, &[0])),
            &ClientDataLimits::default()
        )
        .is_err()
    );
    assert!(
        ClientData::parse(
            &container(&child(0, 0, ClientDataChildKind::ShapeFlags, &[2])),
            &ClientDataLimits::default()
        )
        .is_err()
    );
    assert!(
        ClientData::parse(
            &container(&child(
                0,
                0,
                ClientDataChildKind::Placeholder,
                &[0, 0, 0, 0, 0, 0, 0, 0]
            )),
            &ClientDataLimits::default()
        )
        .is_err()
    );
    assert!(
        ClientData::parse(
            &container(&child(
                0,
                0,
                ClientDataChildKind::RoundTripNewPlaceholderId12,
                &[24]
            )),
            &ClientDataLimits::default()
        )
        .is_err()
    );
    assert!(
        ClientData::parse(
            &container(&child(
                0,
                2,
                ClientDataChildKind::MouseClickInteractiveInfo,
                &[]
            )),
            &ClientDataLimits::default()
        )
        .is_err()
    );

    let mut truncated = child(0, 0, ClientDataChildKind::ExternalObjectReference, &[0; 4]);
    truncated.pop();
    assert!(ClientData::parse(&container(&truncated), &ClientDataLimits::default()).is_err());
}

#[test]
fn enforces_payload_child_size_and_count_limits() {
    let record = container(&child(
        0x0F,
        0,
        ClientDataChildKind::ProgrammableTags,
        &[1, 2],
    ));
    assert!(
        ClientData::parse_with_limits(
            &record,
            ClientDataLimits {
                max_payload_bytes: 1,
                ..Default::default()
            }
        )
        .is_err()
    );
    assert!(
        ClientData::parse_with_limits(
            &record,
            ClientDataLimits {
                max_child_payload_bytes: 1,
                ..Default::default()
            }
        )
        .is_err()
    );
    assert!(
        ClientData::parse_with_limits(
            &record,
            ClientDataLimits {
                max_child_records: 0,
                ..Default::default()
            }
        )
        .is_err()
    );
}

#[test]
fn snapshot_noop_is_source_exact_and_unknown_order_is_retained() {
    let mut payload = Vec::new();
    payload.extend(child(0, 0, ClientDataChildKind::ShapeFlags, &[1]));
    payload.extend(opaque(7, 0x123, 0x2222, &[0xCA, 0xFE]));
    payload.extend(child(
        0,
        0,
        ClientDataChildKind::Placeholder,
        &[3, 0, 0, 0, 13, 1, 0, 0],
    ));
    let source_bytes = container(&payload);
    let source = super::Snapshot::parse(&source_bytes, &ClientDataLimits::default()).unwrap();

    let commit = source.edit().commit().unwrap();
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().bytes(), source_bytes.as_slice());
    assert_eq!(commit.snapshot(), &source);
}

#[test]
fn transaction_edits_are_checked_ordered_and_reversible() {
    let mut payload = Vec::new();
    payload.extend(child(0, 0, ClientDataChildKind::ShapeFlags, &[1]));
    payload.extend(opaque(7, 0x123, 0x2222, &[0xCA, 0xFE]));
    payload.extend(child(
        0,
        0,
        ClientDataChildKind::Placeholder,
        &[3, 0, 0, 0, 13, 1, 0, 0],
    ));
    payload.extend(child(
        0,
        0,
        ClientDataChildKind::RoundTripShapeId12,
        &0x1020u32.to_le_bytes(),
    ));
    let source = super::Snapshot::parse(container(&payload), &ClientDataLimits::default()).unwrap();
    let mut edit = source.edit();

    let before = edit
        .replace(
            0,
            ClientDataChild::new(ClientDataChildKind::ShapeFlags, vec![0]).unwrap(),
        )
        .unwrap();
    assert_eq!(before.payload(), &[1]);
    let removed = edit.remove(3).unwrap();
    assert_eq!(removed.kind(), ClientDataChildKind::RoundTripShapeId12);
    edit.insert(
        3,
        ClientDataChild::new(
            ClientDataChildKind::RoundTripHeaderFooterPlaceholder12,
            vec![7],
        )
        .unwrap(),
    )
    .unwrap();

    let commit = edit.commit().unwrap();
    assert_eq!(commit.patch().changes().len(), 3);
    assert_eq!(commit.snapshot().children()[1].record_type(), 0x2222);
    assert_eq!(commit.snapshot().children()[1].payload(), &[0xCA, 0xFE]);
    assert_eq!(
        commit.snapshot().children()[3].kind(),
        ClientDataChildKind::RoundTripHeaderFooterPlaceholder12
    );
    assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
    assert!(commit.patch().apply(commit.snapshot()).is_err());
    assert_eq!(commit.patch().undo(commit.snapshot()).unwrap(), source);
    assert_eq!(commit.patch().redo(&source).unwrap(), *commit.snapshot());
}

#[test]
fn failed_child_edits_leave_the_candidate_unchanged() {
    let source = super::Snapshot::parse(
        container(
            &[
                child(0, 0, ClientDataChildKind::ShapeFlags, &[1]),
                child(
                    0,
                    0,
                    ClientDataChildKind::Placeholder,
                    &[3, 0, 0, 0, 13, 1, 0, 0],
                ),
            ]
            .concat(),
        ),
        &ClientDataLimits::default(),
    )
    .unwrap();
    let mut edit = source.edit();
    let original = edit.children().to_vec();

    assert!(
        edit.insert(
            99,
            ClientDataChild::opaque(0, 0, 0x2222, vec![0xAA]).unwrap(),
        )
        .is_err()
    );
    assert_eq!(edit.children(), original.as_slice());

    assert!(
        edit.replace(
            0,
            ClientDataChild::new(ClientDataChildKind::RoundTripShapeId12, vec![0, 0, 0, 0],)
                .unwrap(),
        )
        .is_err()
    );
    assert_eq!(edit.children(), original.as_slice());

    let limited = super::Snapshot::parse_with_limits(
        container(&opaque(7, 0x123, 0x2222, &[0xCA, 0xFE])),
        ClientDataLimits {
            max_child_payload_bytes: 2,
            ..Default::default()
        },
    )
    .unwrap();
    let mut limited_edit = limited.edit();
    let limited_original = limited_edit.children().to_vec();
    assert!(
        limited_edit
            .append(ClientDataChild::opaque(0, 0, 0x3333, vec![0; 3]).unwrap())
            .is_err()
    );
    assert_eq!(limited_edit.children(), limited_original.as_slice());
}

/// A built one-slide package with two textboxes; the first gains an
/// `OfficeArtClientData` holding one `ShapeFlags10Atom`.
fn package_with_client_data() -> (crate::Package, Vec<u8>) {
    use litchi_core::Position;

    use crate::record_wire as wire;
    use crate::slide_order::records::{Owner, Rewrite, publish};

    let mut builder = crate::Builder::new();
    let slide = builder.add_slide().unwrap();
    builder.add_textbox(slide, 10, 10, 200, 40, "One").unwrap();
    builder.add_textbox(slide, 10, 80, 200, 40, "Two").unwrap();
    let package = builder.build().unwrap();
    let shape_id = {
        let presentation = package.presentation().unwrap();
        let slides = presentation.slides().unwrap();
        let shapes = slides[0].shapes().unwrap();
        crate::text_edit::native_shape_id(&shapes[0])
    };
    let base = crate::scoped::Base::capture(&package).unwrap();
    let native = base.native().unwrap();
    let owner = Owner::Slide(Position::new(0));
    let before = base.owner_record(&native, owner).unwrap();
    let record = container(&child(0, 0, ClientDataChildKind::ShapeFlags10, &[0]));
    let after =
        crate::hyperlink::edit::rewrite_slide_shape(&before, shape_id, &mut |shape| {
            let children = wire::child_list(shape)?;
            assert!(children.iter().all(
                |child| wire::record_type(child).unwrap() != OFFICE_ART_CLIENT_DATA_RECORD_TYPE
            ));
            let mut data = Vec::new();
            let mut inserted = false;
            for child in &children {
                if !inserted && wire::record_type(child)? == 0xF00D {
                    data.extend_from_slice(&record);
                    inserted = true;
                }
                data.extend_from_slice(child);
            }
            if !inserted {
                data.extend_from_slice(&record);
            }
            Ok(wire::rebuild(shape, &data)?)
        })
        .unwrap();
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
    (crate::Package::from_bytes(&bytes).unwrap(), bytes)
}

#[test]
fn package_client_data_edits_publish_replay_and_invert_exactly() {
    use litchi_core::Position;
    use litchi_core::edit::RefusalKind;

    use crate::text_edit::Target;

    let (source, source_bytes) = package_with_client_data();
    let target = Target::new(Position::new(0), Position::new(0));
    let kinds = |package: &crate::Package| {
        package
            .edit_client_data(target)
            .unwrap()
            .children()
            .unwrap()
            .iter()
            .map(ClientDataChild::kind)
            .collect::<Vec<_>>()
    };
    assert_eq!(kinds(&source), vec![ClientDataChildKind::ShapeFlags10]);

    let no_op = source.edit_client_data(target).unwrap().commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), source_bytes);

    let flags = ClientDataChild::new(ClientDataChildKind::ShapeFlags, vec![1]).unwrap();
    let mut edit = source.edit_client_data(target).unwrap();
    edit.insert(0, flags.clone()).unwrap();
    drop(edit);
    assert_eq!(source.to_bytes().unwrap(), source_bytes);

    let mut edit = source.edit_client_data(target).unwrap();
    let Err(super::Error::Refused(refusal)) = edit.remove(9) else {
        panic!("a missing child must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
    let Err(super::Error::Refused(refusal)) = edit.replace(9, flags.clone()) else {
        panic!("a missing child must be refused");
    };
    assert_eq!(refusal.kind(), RefusalKind::NotFound);
    assert!(edit.changes().unwrap().is_empty());
    edit.insert(0, flags).unwrap();
    assert_eq!(edit.changes().unwrap().len(), 1);
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().source_len(), source_bytes.len());
    assert_eq!(
        commit.patch().target_len(),
        commit.package().to_bytes().unwrap().len()
    );
    assert_eq!(
        kinds(commit.package()),
        vec![
            ClientDataChildKind::ShapeFlags,
            ClientDataChildKind::ShapeFlags10
        ]
    );

    let replayed = source.apply_client_data(commit.patch()).unwrap();
    assert_eq!(
        replayed.package().to_bytes().unwrap(),
        commit.package().to_bytes().unwrap()
    );
    let restored = commit
        .package()
        .apply_client_data(&commit.patch().inverse())
        .unwrap();
    assert_eq!(restored.package().to_bytes().unwrap(), source_bytes);
    assert!(matches!(
        commit.package().apply_client_data(commit.patch()),
        Err(super::Error::PatchConflict)
    ));

    let refusal_kind = |target| match source
        .edit_client_data(target)
        .and_then(|edit| edit.children().map(|_| ()))
    {
        Err(super::Error::Refused(refusal)) => refusal.kind(),
        other => panic!("expected a refusal, got {other:?}"),
    };
    assert_eq!(
        refusal_kind(Target::new(Position::new(0), Position::new(1))),
        RefusalKind::UnsupportedTarget
    );
    assert_eq!(
        refusal_kind(Target::new(Position::new(0), Position::new(9))),
        RefusalKind::NotFound
    );
    assert_eq!(
        refusal_kind(Target::new(Position::new(3), Position::new(0))),
        RefusalKind::NotFound
    );
}

#[test]
fn package_client_data_edits_admit_a_signed_package_and_refuse_a_changed_commit() {
    use litchi_core::Position;

    use crate::text_edit::Target;

    let (source, _) = package_with_client_data();
    let signed = crate::scoped::signed(&source);
    let bytes = signed.to_bytes().unwrap();
    let target = Target::new(Position::new(0), Position::new(0));

    let no_op = signed.edit_client_data(target).unwrap().commit().unwrap();
    assert!(!no_op.diagnostics().changed());
    assert!(no_op.patch().is_noop());
    assert_eq!(no_op.package().to_bytes().unwrap(), bytes);

    let flags = ClientDataChild::new(ClientDataChildKind::ShapeFlags, vec![1]).unwrap();
    let mut edit = signed.edit_client_data(target).unwrap();
    edit.insert(0, flags).unwrap();
    assert!(matches!(edit.commit(), Err(super::Error::Signed)));
    assert_eq!(signed.to_bytes().unwrap(), bytes);
}

#[test]
fn duplicate_candidate_is_an_invalid_value_but_duplicate_source_remains_corrupt() {
    use crate::text_edit::Target;
    use litchi_core::{Position, edit::RefusalKind};

    let wire = child(0, 0, ClientDataChildKind::ShapeFlags10, &[4]);
    let mut duplicate = wire.clone();
    duplicate.extend_from_slice(&wire);
    assert!(matches!(
        super::Snapshot::parse(container(&duplicate), &ClientDataLimits::default()),
        Err(crate::package::Error::Corrupted(_))
    ));

    let (source, bytes) = package_with_client_data();
    let mut edit = source
        .edit_client_data(Target::new(Position::new(0), Position::new(0)))
        .unwrap();
    let original = edit.children().unwrap()[0].clone();
    edit.insert(
        0,
        ClientDataChild::new(ClientDataChildKind::ShapeFlags, vec![1]).unwrap(),
    )
    .unwrap();
    let staged = edit.children().unwrap().to_vec();
    let Err(super::Error::Refused(refusal)) = edit.append(original) else {
        panic!("invalid candidate must be a typed refusal");
    };
    assert_eq!(refusal.kind(), RefusalKind::InvalidValue);
    assert_eq!(edit.children().unwrap(), staged);
    assert_eq!(edit.changes().unwrap().len(), 1);
    assert!(edit.commit().unwrap().diagnostics().changed());
    assert_eq!(source.to_bytes().unwrap(), bytes);
}
