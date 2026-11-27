use super::codec::{
    ViewCollector, parse_pane, parse_scl, parse_selection, parse_window2, read_u16,
};
use super::{
    PANE_RECORD_TYPE, PaneType, SCL_RECORD_TYPE, SELECTION_RECORD_TYPE, WINDOW2_RECORD_TYPE,
};
use crate::error::Error;

fn window2(flags: u16) -> [u8; 18] {
    let mut data = [0u8; 18];
    data[0..2].copy_from_slice(&flags.to_le_bytes());
    data[6..8].copy_from_slice(&64u16.to_le_bytes());
    data[10..12].copy_from_slice(&60u16.to_le_bytes());
    data[12..14].copy_from_slice(&75u16.to_le_bytes());
    data
}

fn pane() -> [u8; 10] {
    [5, 0, 7, 0, 7, 0, 34, 0, 0, 0]
}

fn selection(pane: u8) -> [u8; 15] {
    [pane, 7, 0, 34, 0, 0, 0, 1, 0, 7, 0, 7, 0, 34, 34]
}

#[test]
fn binary_field_reads_return_typed_errors_for_truncation_and_overflow() {
    let truncated = read_u16(&[0], 0, WINDOW2_RECORD_TYPE, "WINDOW2.flags").unwrap_err();
    assert!(matches!(
        truncated,
        Error::InvalidRecord {
            record_type: WINDOW2_RECORD_TYPE,
            message,
        } if message == "truncated WINDOW2.flags"
    ));

    let overflowing = read_u16(&[], usize::MAX, WINDOW2_RECORD_TYPE, "WINDOW2.flags").unwrap_err();
    assert!(matches!(
        overflowing,
        Error::InvalidRecord {
            record_type: WINDOW2_RECORD_TYPE,
            message,
        } if message == "WINDOW2.flags offset overflows"
    ));
}

#[test]
fn parses_window_zoom_pane_and_selection() {
    let mut collector = ViewCollector::new();
    collector
        .feed_record(WINDOW2_RECORD_TYPE, &window2(0x07be))
        .unwrap();
    collector
        .feed_record(SCL_RECORD_TYPE, &[3, 0, 4, 0])
        .unwrap();
    collector.feed_record(PANE_RECORD_TYPE, &pane()).unwrap();
    collector
        .feed_record(SELECTION_RECORD_TYPE, &selection(0))
        .unwrap();
    let views = collector.finish().unwrap();
    let view = &views[0];
    assert!(view.has_frozen_panes());
    assert!(view.is_frozen_without_split());
    assert_eq!(view.zoom_fraction(), Some((3, 4)));
    assert_eq!(view.normal_zoom_percent(), Some(75));
    assert_eq!(view.pane().unwrap().right_pane_left_column(), 34);
    assert_eq!(view.selections()[0].ranges()[0].first_row(), 7);
}

#[test]
fn rejects_malformed_and_out_of_order_view_records() {
    assert!(parse_window2(&[0; 17]).is_err());
    assert!(parse_scl(&[0, 0, 1, 0]).is_err());
    assert!(parse_pane(&[0; 9], false).is_err());
    assert!(parse_selection(&[0; 8]).is_err());

    let mut collector = ViewCollector::new();
    collector
        .feed_record(WINDOW2_RECORD_TYPE, &window2(0x002e))
        .unwrap();
    collector.feed_record(PANE_RECORD_TYPE, &pane()).unwrap();
    assert!(
        collector
            .feed_record(SCL_RECORD_TYPE, &[1, 0, 1, 0])
            .is_err()
    );
}

#[test]
fn ignores_custom_view_selections_after_window_closes() {
    let mut collector = ViewCollector::new();
    collector
        .feed_record(WINDOW2_RECORD_TYPE, &window2(0x0026))
        .unwrap();
    collector
        .feed_record(SELECTION_RECORD_TYPE, &selection(3))
        .unwrap();
    collector.feed_record(0x01aa, &[]).unwrap();
    collector
        .feed_record(SELECTION_RECORD_TYPE, &selection(0))
        .unwrap();
    let views = collector.finish().unwrap();
    assert_eq!(views[0].selections().len(), 1);
}

#[test]
fn validates_active_range_across_contiguous_selection_records() {
    let mut first = selection(0);
    first[5..7].copy_from_slice(&1u16.to_le_bytes());
    let mut collector = ViewCollector::new();
    collector
        .feed_record(WINDOW2_RECORD_TYPE, &window2(0x002e))
        .unwrap();
    collector
        .feed_record(SELECTION_RECORD_TYPE, &first)
        .unwrap();
    let mut second = selection(0);
    second[5..7].copy_from_slice(&1u16.to_le_bytes());
    collector
        .feed_record(SELECTION_RECORD_TYPE, &second)
        .unwrap();
    assert!(collector.finish().is_ok());

    let mut invalid = selection(0);
    invalid[1..3].copy_from_slice(&8u16.to_le_bytes());
    let mut collector = ViewCollector::new();
    collector
        .feed_record(WINDOW2_RECORD_TYPE, &window2(0x002e))
        .unwrap();
    collector
        .feed_record(SELECTION_RECORD_TYPE, &invalid)
        .unwrap();
    assert!(collector.finish().is_err());
}

#[test]
fn rejects_cross_record_view_inconsistencies() {
    assert!(parse_scl(&[0x00, 0x80, 1, 0]).is_err());

    let mut bad_active_pane = pane();
    bad_active_pane[0..2].copy_from_slice(&0u16.to_le_bytes());
    let mut collector = ViewCollector::new();
    collector
        .feed_record(WINDOW2_RECORD_TYPE, &window2(0x0026))
        .unwrap();
    collector
        .feed_record(PANE_RECORD_TYPE, &bad_active_pane)
        .unwrap();
    assert!(collector.finish().is_err());
}

/// One row per record sequence: the outcome of every record and what `finish`
/// returns (the number of views, or the error text).
#[test]
fn window_production_order_accepts_and_rejects_the_same_sequences() {
    const PLV_RECORD_TYPE: u16 = 0x088b;
    const OTHER: u16 = 0x01aa;
    const PLV_ORDER: &str =
        "Invalid record 0x088B: record is out of order in the WINDOW production";
    const SCL_ORDER: &str =
        "Invalid record 0x00A0: record is out of order in the WINDOW production";
    const PANE_ORDER: &str =
        "Invalid record 0x0041: record is out of order in the WINDOW production";
    let payload = |record_type: u16| -> Vec<u8> {
        match record_type {
            WINDOW2_RECORD_TYPE => window2(0x07be).to_vec(),
            SCL_RECORD_TYPE => vec![3, 0, 4, 0],
            PANE_RECORD_TYPE => pane().to_vec(),
            SELECTION_RECORD_TYPE => selection(0).to_vec(),
            _ => Vec::new(),
        }
    };
    const W: u16 = WINDOW2_RECORD_TYPE;
    const L: u16 = PLV_RECORD_TYPE;
    const Z: u16 = SCL_RECORD_TYPE;
    const P: u16 = PANE_RECORD_TYPE;
    const S: u16 = SELECTION_RECORD_TYPE;
    type Case = (&'static [u16], &'static [&'static str], usize);
    let cases: &[Case] = &[
        (&[], &[], 0),
        (&[L, Z, P, S, OTHER], &["ok"; 5], 0),
        (&[W], &["ok"], 1),
        (&[W, L, Z, P, S], &["ok"; 5], 1),
        (&[W, L, P, S, S], &["ok"; 5], 1),
        (&[W, Z, S], &["ok"; 3], 1),
        (&[W, L, L, L], &["ok", "ok", PLV_ORDER, PLV_ORDER], 1),
        (&[W, Z, L], &["ok", "ok", PLV_ORDER], 1),
        (&[W, P, L], &["ok", "ok", PLV_ORDER], 1),
        (&[W, S, L], &["ok", "ok", PLV_ORDER], 1),
        (&[W, Z, Z], &["ok", "ok", SCL_ORDER], 1),
        (&[W, P, Z], &["ok", "ok", SCL_ORDER], 1),
        (&[W, S, Z], &["ok", "ok", SCL_ORDER], 1),
        (&[W, P, P], &["ok", "ok", PANE_ORDER], 1),
        (&[W, S, P], &["ok", "ok", PANE_ORDER], 1),
        (&[W, L, OTHER, L, Z, P], &["ok"; 6], 1),
        (&[W, L, W, L, Z, W, Z, P], &["ok"; 8], 3),
        (
            &[W, L, L, Z, P, S],
            &["ok", "ok", PLV_ORDER, "ok", "ok", "ok"],
            1,
        ),
    ];
    for (records, expected, views) in cases {
        let mut collector = ViewCollector::new();
        let outcomes: Vec<String> = records
            .iter()
            .map(
                |record_type| match collector.feed_record(*record_type, &payload(*record_type)) {
                    Ok(()) => "ok".to_string(),
                    Err(error) => error.to_string(),
                },
            )
            .collect();
        assert_eq!(outcomes, *expected, "{records:04x?}");
        assert_eq!(collector.finish().unwrap().len(), *views, "{records:04x?}");
    }
}

#[test]
fn reads_poi_pane_and_zoom_fixtures() {
    use crate::raw::Reader;
    use std::fs::File;
    use std::path::Path;

    let fixture = |name: &str| {
        Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet")
            .join(name)
    };

    let zoomed = Reader::new(File::open(fixture("41139.xls")).unwrap()).unwrap();
    let view = zoomed.xls_worksheet(0).unwrap().worksheet_view().unwrap();
    assert_eq!(view.zoom_fraction(), Some((3, 4)));
    assert_eq!(view.normal_zoom_percent(), Some(75));
    assert_eq!(view.selections().len(), 4);
    let pane = view.pane().unwrap();
    assert_eq!((pane.horizontal_split(), pane.vertical_split()), (5, 7));
    assert_eq!(
        (pane.bottom_pane_top_row(), pane.right_pane_left_column()),
        (7, 34)
    );
    assert_eq!(pane.active_pane(), PaneType::LowerRight);

    let split = Reader::new(File::open(fixture("50939.xls")).unwrap()).unwrap();
    let view = split.xls_worksheet(0).unwrap().worksheet_view().unwrap();
    assert!(view.has_frozen_panes());
    assert!(!view.is_frozen_without_split());
    let pane = view.pane().unwrap();
    assert_eq!((pane.horizontal_split(), pane.vertical_split()), (8, 4));
    assert_eq!(
        (pane.bottom_pane_top_row(), pane.right_pane_left_column()),
        (4, 26)
    );
}
