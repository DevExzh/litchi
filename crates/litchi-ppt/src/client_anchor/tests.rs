use super::*;

fn record(payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&0u16.to_le_bytes());
    bytes.extend_from_slice(&RECORD_TYPE.to_le_bytes());
    bytes.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

#[test]
fn small_rect_round_trips_with_normative_field_order() {
    let payload = [
        0x9C, 0xFF, // top = -100
        0x38, 0xFF, // left = -200
        0x2C, 0x01, // right = 300
        0x90, 0x01, // bottom = 400
    ];
    let bytes = record(&payload);
    let anchor = Anchor::parse(&bytes).unwrap();

    assert_eq!(anchor.encoding(), Encoding::Small);
    assert_eq!((anchor.left(), anchor.top()), (-200, -100));
    assert_eq!((anchor.right(), anchor.bottom()), (300, 400));
    assert_eq!((anchor.width(), anchor.height()), (500, 500));
    assert_eq!(anchor.to_bytes(), bytes);
}

#[test]
fn full_rect_round_trips_extreme_coordinates_without_overflow() {
    let anchor = Anchor::full(i32::MIN, -7, i32::MAX, 9).unwrap();
    let parsed = Anchor::parse(anchor.to_bytes()).unwrap();

    assert_eq!(parsed.encoding(), Encoding::Full);
    assert_eq!(parsed.width(), u32::MAX as i64);
    assert_eq!(parsed.height(), 16);
}

#[test]
fn strict_codec_rejects_bad_headers_lengths_bounds_and_limits() {
    let valid = Anchor::small(1, 2, 3, 4).unwrap().to_bytes();
    for index in [0, 1, 2] {
        let mut bad = valid.clone();
        bad[index] ^= 1;
        assert!(Anchor::parse(&bad).is_err());
    }
    let mut bad_length = valid.clone();
    bad_length[4..8].copy_from_slice(&12u32.to_le_bytes());
    assert!(Anchor::parse(&bad_length).is_err());
    assert!(Anchor::parse(&valid[..valid.len() - 1]).is_err());
    assert!(Anchor::small(4, 2, 3, 5).is_err());
    assert!(Anchor::full(1, 8, 3, 4).is_err());

    let full = Anchor::full(1, 2, 3, 4).unwrap().to_bytes();
    assert!(
        Anchor::parse_with_limits(
            full,
            Limits {
                max_payload_bytes: 8,
            },
        )
        .is_err()
    );
}

#[test]
fn no_op_commit_preserves_the_exact_source_snapshot() {
    let bytes = Anchor::small(-2, -1, 4, 8).unwrap().to_bytes();
    let source = Snapshot::parse(&bytes).unwrap();
    let commit = source.edit().commit().unwrap();

    assert_eq!(commit.snapshot().bytes(), bytes);
    assert_eq!(commit.snapshot().revision(), source.revision());
    assert!(commit.patch().is_empty());
}

#[test]
fn transaction_changes_encoding_and_patch_is_reversible() {
    let source = Snapshot::from_anchor(Anchor::small(10, 20, 30, 40).unwrap());
    let mut edit = source.edit();
    edit.set_full(-100_000, -200_000, 300_000, 400_000).unwrap();
    let commit = edit.commit().unwrap();

    assert_eq!(commit.anchor().encoding(), Encoding::Full);
    assert_eq!(commit.anchor().left(), -100_000);
    let change = commit.patch().change().unwrap();
    assert_eq!(change.before(), source.anchor());
    assert_eq!(change.after(), commit.anchor());
    assert_eq!(commit.patch().apply(&source).unwrap(), *commit.snapshot());
    assert_eq!(
        commit.patch().inverse().apply(commit.snapshot()).unwrap(),
        source
    );
}

#[test]
fn failed_edit_is_atomic_and_patch_rejects_an_unrelated_source() {
    let source = Snapshot::from_anchor(Anchor::small(1, 2, 3, 4).unwrap());
    let mut edit = source.edit();
    assert!(edit.set_small(9, 2, 3, 4).is_err());
    assert!(!edit.is_changed());

    edit.set_small(2, 3, 4, 5).unwrap();
    let commit = edit.commit().unwrap();
    let unrelated = Snapshot::from_anchor(Anchor::small(0, 0, 1, 1).unwrap());
    assert!(commit.patch().apply(&unrelated).is_err());
}
