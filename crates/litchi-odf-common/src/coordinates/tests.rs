#![allow(
    clippy::unwrap_used,
    reason = "Test assertions intentionally unwrap known-valid coordinate fixture construction failures."
)]

use super::codec::{alpha_to_digit, digit_to_alpha};
use super::model::{CellCoord, CellRange, MAX_INDEX};

#[test]
fn alpha_to_digit_converts_columns_and_rejects_malformed_or_oversized_input() {
    assert_eq!(alpha_to_digit("A").unwrap(), 0);
    assert_eq!(alpha_to_digit("B").unwrap(), 1);
    assert_eq!(alpha_to_digit("Z").unwrap(), 25);
    assert_eq!(alpha_to_digit("AA").unwrap(), 26);
    assert_eq!(alpha_to_digit("AB").unwrap(), 27);
    assert_eq!(alpha_to_digit("AZ").unwrap(), 51);
    assert_eq!(alpha_to_digit("BA").unwrap(), 52);
    assert_eq!(alpha_to_digit("a").unwrap(), 0);
    assert_eq!(alpha_to_digit("aa").unwrap(), 26);

    for value in ["", "A1", "1A"] {
        assert!(
            alpha_to_digit(value).is_err(),
            "accepted malformed {value:?}"
        );
    }
    let oversized = "Z".repeat(usize::BITS as usize);
    assert!(alpha_to_digit(&oversized).is_err());
}

#[test]
fn digit_to_alpha_round_trips_the_representable_coordinate_domain() {
    for index in 0..100 {
        let alpha = digit_to_alpha(index);
        assert_eq!(alpha_to_digit(&alpha).unwrap(), index);
    }
    assert_eq!(digit_to_alpha(0), "A");
    assert_eq!(digit_to_alpha(25), "Z");
    assert_eq!(digit_to_alpha(26), "AA");
    assert_eq!(digit_to_alpha(27), "AB");
    assert_eq!(digit_to_alpha(51), "AZ");
    assert_eq!(digit_to_alpha(52), "BA");
    assert_eq!(
        alpha_to_digit(&digit_to_alpha(usize::MAX)).unwrap(),
        usize::MAX
    );
}

#[test]
fn cell_coordinate_parsing_and_display_preserve_a1_behavior() {
    let first_coord: CellCoord = "A1".parse().unwrap();
    assert_eq!(first_coord.column(), 0);
    assert_eq!(first_coord.row(), 0);

    let second_coord: CellCoord = "B3".parse().unwrap();
    assert_eq!(second_coord.column(), 1);
    assert_eq!(second_coord.row(), 2);

    let wide_coord: CellCoord = "AA10".parse().unwrap();
    assert_eq!(wide_coord.column(), 26);
    assert_eq!(wide_coord.row(), 9);
    assert_eq!(wide_coord.to_a1(), "AA10");

    for value in ["A0", "1A", "A"] {
        assert!(value.parse::<CellCoord>().is_err(), "accepted {value}");
    }
    assert!(CellCoord::new(MAX_INDEX + 1, 0).is_err());
    assert!(CellCoord::new(0, MAX_INDEX + 1).is_err());
    assert_eq!(CellCoord::new(0, 0).unwrap().to_string(), "A1");
    assert_eq!(CellCoord::new(1, 2).unwrap().to_string(), "B3");
}

#[test]
fn cell_range_is_checked_and_keeps_inclusive_dimensions() {
    let first_range: CellRange = "A1:B3".parse().unwrap();
    assert_eq!(first_range.start().column(), 0);
    assert_eq!(first_range.start().row(), 0);
    assert_eq!(first_range.end().column(), 1);
    assert_eq!(first_range.end().row(), 2);
    assert_eq!(first_range.width(), 2);
    assert_eq!(first_range.height(), 3);
    assert!(first_range.contains(CellCoord::new(1, 2).unwrap()));
    assert!(!first_range.contains(CellCoord::new(2, 2).unwrap()));

    let wide_range: CellRange = "AA10:AB20".parse().unwrap();
    assert_eq!(wide_range.width(), 2);
    assert_eq!(wide_range.height(), 11);
    assert_eq!(wide_range.to_string(), "AA10:AB20");

    for value in ["A1", "A1:", ":B3", "A1:B3:C4"] {
        assert!(value.parse::<CellRange>().is_err(), "accepted {value}");
    }

    let start = CellCoord::new(1, 1).unwrap();
    let end = CellCoord::new(0, 0).unwrap();
    assert!(CellRange::new(start, end).is_err());
    let single_cell_range = CellRange::new(CellCoord::new(0, 0).unwrap(), end).unwrap();
    assert_eq!(single_cell_range.to_string(), "A1:A1");
}
