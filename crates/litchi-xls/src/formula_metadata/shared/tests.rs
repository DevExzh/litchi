//! Focused ShrFmla owner and payload regressions.

use super::codec::parse;
use super::{Cell, Owner, Range};

#[test]
fn owner_checks_range_anchor_participants_and_record_order() {
    let range = Range::try_new(0, 0, 2, 2).unwrap();
    let anchor = Cell::new(0, 1);
    let owner = Owner::new(range, anchor, &[0x1E, 2, 0]).unwrap();
    let owner = owner
        .with_participants(&[anchor, Cell::new(1, 1), Cell::new(2, 2)])
        .unwrap();

    assert_eq!(owner.range(), range);
    assert_eq!(owner.anchor(), anchor);
    assert_eq!(owner.count(), 3);
    assert_eq!(owner.c_use().unwrap(), 3);
    assert_eq!(owner.anchor_tokens(), [0x01, 0, 0, 1, 0]);
}

#[test]
fn owner_rejects_inconsistent_participants() {
    let range = Range::try_new(0, 0, 2, 2).unwrap();
    let anchor = Cell::new(1, 1);
    let owner = Owner::new(range, anchor, &[0x1E, 2, 0]).unwrap();

    assert!(
        owner
            .clone()
            .with_participants(&[Cell::new(0, 0), anchor])
            .is_err()
    );
    assert!(owner.with_participants(&[anchor, anchor]).is_err());
    assert!(Range::try_new(0, 2, 1, 1).is_err());
}

#[test]
fn parser_round_trips_refu_reserved_count_and_shared_tokens() {
    let data = [0, 0, 2, 0, 1, 3, 0, 2, 3, 0, 0x1E, 2, 0];
    let parsed = parse(&data).unwrap();
    assert_eq!(parsed.range, Range::try_new(0, 1, 2, 3).unwrap());
    assert_eq!(parsed.reserved, 0);
    assert_eq!(parsed.count, 2);
    assert_eq!(parsed.tokens, [0x1E, 2, 0]);
}

#[test]
fn parser_rejects_reserved_bytes_empty_count_and_length_mismatches() {
    let mut data = [0, 0, 0, 0, 0, 0, 0, 1, 2, 0, 0x1E];
    assert!(parse(&data).is_err());

    data[6] = 1;
    assert!(parse(&data).is_err());

    data[6] = 0;
    data[7] = 0;
    assert!(parse(&data).is_err());
}
