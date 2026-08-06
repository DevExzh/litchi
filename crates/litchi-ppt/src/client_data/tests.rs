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

    let parsed = ClientData::parse(&bytes).unwrap();
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
        ClientData::parse(&value.to_bytes().unwrap()).unwrap(),
        value
    );
}

#[test]
fn rejects_outer_header_length_truncation_and_trailing_data() {
    let valid = container(&[]);
    for index in [0usize, 1, 2] {
        let mut bad = valid.clone();
        bad[index] ^= 1;
        assert!(ClientData::parse(&bad).is_err());
    }
    let mut bad_length = valid.clone();
    bad_length[4..8].copy_from_slice(&1u32.to_le_bytes());
    assert!(ClientData::parse(&bad_length).is_err());
    let mut trailing = valid.clone();
    trailing.push(0);
    assert!(ClientData::parse(&trailing).is_err());
    assert!(ClientData::parse(&valid[..7]).is_err());
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
    assert!(ClientData::parse(&container(&[flags.clone(), flags.clone()].concat())).is_err());
    assert!(ClientData::parse(&container(&[placeholder, flags].concat())).is_err());

    let unknown = opaque(7, 0x123, 0x2222, &[0xCA, 0xFE]);
    let unknown_container = container(&unknown);
    let parsed = ClientData::parse(&unknown_container).unwrap();
    assert_eq!(parsed.children().len(), 1);
    assert_eq!(parsed.children()[0].kind(), ClientDataChildKind::Unknown);
    assert_eq!(parsed.children()[0].record_type(), 0x2222);
    assert_eq!(parsed.children()[0].version(), 7);
    assert_eq!(parsed.children()[0].instance(), 0x123);
    assert_eq!(parsed.children()[0].payload(), &[0xCA, 0xFE]);
    assert_eq!(parsed.to_bytes().unwrap(), unknown_container);

    let shape_id = child(0, 0, ClientDataChildKind::RoundTripShapeId12, &[0; 4]);
    assert!(ClientData::parse(&container(&[shape_id.clone(), shape_id].concat())).is_err());
}

#[test]
fn rejects_bad_child_headers_reserved_values_and_boundaries() {
    assert!(
        ClientData::parse(&container(&child(
            1,
            0,
            ClientDataChildKind::ShapeFlags,
            &[0]
        )))
        .is_err()
    );
    assert!(
        ClientData::parse(&container(&child(
            0,
            1,
            ClientDataChildKind::ShapeFlags,
            &[0]
        )))
        .is_err()
    );
    assert!(
        ClientData::parse(&container(&child(
            0,
            0,
            ClientDataChildKind::ShapeFlags,
            &[2]
        )))
        .is_err()
    );
    assert!(
        ClientData::parse(&container(&child(
            0,
            0,
            ClientDataChildKind::Placeholder,
            &[0, 0, 0, 0, 0, 0, 0, 0]
        )))
        .is_err()
    );
    assert!(
        ClientData::parse(&container(&child(
            0,
            0,
            ClientDataChildKind::RoundTripNewPlaceholderId12,
            &[24]
        )))
        .is_err()
    );
    assert!(
        ClientData::parse(&container(&child(
            0,
            2,
            ClientDataChildKind::MouseClickInteractiveInfo,
            &[]
        )))
        .is_err()
    );

    let mut truncated = child(0, 0, ClientDataChildKind::ExternalObjectReference, &[0; 4]);
    truncated.pop();
    assert!(ClientData::parse(&container(&truncated)).is_err());
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
    let source = super::Snapshot::parse(&source_bytes).unwrap();

    let snapshot = source.edit().snapshot().unwrap();
    assert_eq!(snapshot.bytes(), source_bytes.as_slice());
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
    let source = super::Snapshot::parse(container(&payload)).unwrap();
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
    let source = super::Snapshot::parse(container(
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
    ))
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
    let original = limited_edit.children().to_vec();
    assert!(
        limited_edit
            .append(ClientDataChild::opaque(0, 0, 0x3333, vec![0; 3]).unwrap())
            .is_err()
    );
    assert_eq!(limited_edit.children(), original.as_slice());
}
