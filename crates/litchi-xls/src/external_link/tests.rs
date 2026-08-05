//! Regression tests for the layered external-link owner.

use super::codec::{parse_cache_row, parse_sup_book};
use super::model::{CachedValue, NameBody};
use super::package::ExternalLinkCollector;
use super::{
    CONTINUE_RECORD_TYPE, EXTERN_NAME_RECORD_TYPE, EXTERN_SHEET_RECORD_TYPE, SUP_BOOK_RECORD_TYPE,
    XCT_RECORD_TYPE,
};

#[test]
fn malformed_supbook_xct_and_crn_are_rejected() {
    assert!(parse_sup_book(&[1, 0, 2, 4]).is_err());
    assert!(parse_cache_row(&[0, 1, 0, 0]).is_err());
    assert!(parse_cache_row(&[0, 0, 0, 0, 4, 2, 0, 0, 0, 0, 0, 0, 0]).is_err());

    let mut collector = ExternalLinkCollector::new();
    assert!(
        collector
            .feed_record(XCT_RECORD_TYPE, &[1, 0, 0, 0])
            .is_err()
    );
}

#[test]
fn xct_cardinality_and_xti_bounds_are_strict() {
    let external = [1, 0, 2, 0, 0, 1, b'A', 1, 0, 0, b'S'];
    let mut collector = ExternalLinkCollector::new();
    collector
        .feed_record(SUP_BOOK_RECORD_TYPE, &external)
        .unwrap();
    collector
        .feed_record(XCT_RECORD_TYPE, &[1, 0, 0, 0])
        .unwrap();
    assert!(
        collector
            .feed_record(EXTERN_SHEET_RECORD_TYPE, &[0, 0])
            .is_err()
    );

    let mut collector = ExternalLinkCollector::new();
    collector
        .feed_record(SUP_BOOK_RECORD_TYPE, &external)
        .unwrap();
    collector
        .feed_record(EXTERN_SHEET_RECORD_TYPE, &[1, 0, 1, 0, 0, 0, 0, 0])
        .unwrap();
    assert!(collector.finish(1).is_err());
}

#[test]
fn extern_names_are_contextual_bounded_and_continue_complete_moper_values() {
    let addin = [1, 0, 1, 0x3a];
    let addin_name = [0, 0, 0, 0, 0, 0, 1, 0, b'F', 2, 0, 0x1c, 0x17];
    let mut collector = ExternalLinkCollector::new();
    assert!(
        collector
            .feed_record(EXTERN_NAME_RECORD_TYPE, &addin_name)
            .is_err()
    );
    collector.feed_record(SUP_BOOK_RECORD_TYPE, &addin).unwrap();
    collector
        .feed_record(EXTERN_NAME_RECORD_TYPE, &addin_name)
        .unwrap();
    assert!(collector.feed_record(CONTINUE_RECORD_TYPE, &[0]).is_err());

    let dde = [0, 0, 3, 0, 0, b'A', 3, b'B'];
    let dde_name = [2, 0, 0, 0, 0, 0, 1, 0, b'X', 0, 0, 0];
    let mut collector = ExternalLinkCollector::new();
    collector.feed_record(SUP_BOOK_RECORD_TYPE, &dde).unwrap();
    collector
        .feed_record(EXTERN_NAME_RECORD_TYPE, &dde_name)
        .unwrap();
    collector
        .feed_record(CONTINUE_RECORD_TYPE, &[0; 9])
        .unwrap();
    let links = collector.finish(1).unwrap();
    let NameBody::DdeOrOle {
        matrix: Some(matrix),
        ..
    } = links.external_names()[0].body()
    else {
        panic!("expected DDE/OLE body")
    };
    assert_eq!(matrix.last_column, 0);
    assert_eq!(matrix.last_row, 0);
    assert_eq!(matrix.values, [CachedValue::Blank]);
}
