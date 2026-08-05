//! Regression coverage for the BIFF8 formula error-checking owner.

use super::codec::{FEAT_HDR_RECORD_TYPE, FEAT_RECORD_TYPE, FormulaErrorCollector};
use super::model::{Checks, Feature, Header, Range};

const POI_HEADER: [u8; 19] = [
    0x67, 0x08, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3, 0, 1, 0, 0, 0, 0,
];
const POI_FEATURE: [u8; 39] = [
    0x68, 0x08, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 3, 0, 0, 0, 0, 0, 0, 1, 0, 4, 0, 0, 0, 0, 0, 0, 0, 0,
    0, 0, 0, 0, 0, 4, 0, 0, 0,
];

#[test]
fn parses_and_round_trips_poi_reference_records() {
    let header = Header::parse_payload(&POI_HEADER).unwrap();
    assert_eq!(header.to_payload(), POI_HEADER);
    assert_eq!(
        &header.to_record_bytes().unwrap()[0..4],
        &[0x67, 0x08, 19, 0]
    );
    let feature = Feature::parse_payload(&POI_FEATURE).unwrap();
    assert_eq!(feature.ranges().len(), 1);
    assert_eq!(feature.ranges()[0], Range::try_new(0, 0, 0, 0).unwrap());
    assert!(feature.checks().numbers_stored_as_text());
    assert!(!feature.checks().calculation_errors());
    assert_eq!(feature.to_payload().unwrap(), POI_FEATURE);
    assert_eq!(
        &feature.to_record_bytes().unwrap()[0..4],
        &[0x68, 0x08, 39, 0]
    );
}

#[test]
fn constructs_multiple_ranges_and_all_flags() {
    let ranges = vec![
        Range::try_new(1, 7, 3, 3).unwrap(),
        Range::try_new(12, 18, 3, 3).unwrap(),
        Range::try_new(22, 28, 3, 3).unwrap(),
    ];
    let feature = Feature::try_new(ranges.clone(), Checks::from_bits(0xFF)).unwrap();
    let reparsed = Feature::parse_payload(&feature.to_payload().unwrap()).unwrap();
    assert_eq!(reparsed.ranges(), ranges);
    let checks = reparsed.checks();
    assert!(checks.calculation_errors());
    assert!(checks.empty_cell_references());
    assert!(checks.numbers_stored_as_text());
    assert!(checks.inconsistent_ranges());
    assert!(checks.inconsistent_formulas());
    assert!(checks.insufficient_date_time_formats());
    assert!(checks.unprotected_formulas());
    assert!(checks.data_validation());
}

#[test]
fn rejects_malformed_headers_lengths_ranges_flags_and_ordering() {
    let mut bad = POI_HEADER.to_vec();
    bad[0] = 0x68;
    assert!(Header::parse_payload(&bad).is_err());
    let mut bad = POI_HEADER.to_vec();
    bad[2] = 1;
    assert!(Header::parse_payload(&bad).is_err());
    let mut bad = POI_HEADER.to_vec();
    bad[4] = 1;
    assert!(Header::parse_payload(&bad).is_err());
    let mut bad = POI_HEADER.to_vec();
    bad[14] = 0;
    assert!(Header::parse_payload(&bad).is_err());
    let mut bad = POI_HEADER.to_vec();
    bad[15] = 1;
    assert!(Header::parse_payload(&bad).is_err());
    assert!(Header::parse_payload(&POI_HEADER[..18]).is_err());

    let mut bad = POI_FEATURE.to_vec();
    bad[14] = 1;
    assert!(Feature::parse_payload(&bad).is_err());
    let mut bad = POI_FEATURE.to_vec();
    bad[19] = 2;
    assert!(Feature::parse_payload(&bad).is_err());
    let mut bad = POI_FEATURE.to_vec();
    bad[21] = 5;
    assert!(Feature::parse_payload(&bad).is_err());
    let mut bad = POI_FEATURE.to_vec();
    bad[27] = 1;
    assert!(Feature::parse_payload(&bad).is_err());
    let mut bad = POI_FEATURE.to_vec();
    bad[31] = 0;
    bad[32] = 1;
    assert!(Feature::parse_payload(&bad).is_err());
    let mut bad = POI_FEATURE.to_vec();
    bad[38] = 1;
    assert!(Feature::parse_payload(&bad).is_err());
    assert!(Feature::parse_payload(&POI_FEATURE[..38]).is_err());

    let range = Range::try_new(0, 0, 0, 0).unwrap();
    assert!(
        Feature::try_new(vec![range; super::codec::MAX_RANGES + 1], Checks::default(),).is_err()
    );

    let mut collector = FormulaErrorCollector::new();
    assert!(
        collector
            .feed_record(FEAT_RECORD_TYPE, &POI_FEATURE)
            .is_err()
    );
    let mut collector = FormulaErrorCollector::new();
    collector
        .feed_record(FEAT_HDR_RECORD_TYPE, &POI_HEADER)
        .unwrap();
    assert!(
        collector
            .feed_record(FEAT_HDR_RECORD_TYPE, &POI_HEADER)
            .is_err()
    );
    let mut collector = FormulaErrorCollector::new();
    collector
        .feed_record(FEAT_HDR_RECORD_TYPE, &POI_HEADER)
        .unwrap();
    assert!(collector.finish().is_err());
}
