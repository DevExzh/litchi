//! Focused tests for workbook codec boundaries.

use super::validation::cell_record_xf;
use super::wire::{parse_shared_formula_template, pivot_cache_stream_paths};
use crate::records::CellRecord;

#[test]
fn shared_formula_wire_parser_preserves_range_shape() {
    let data = [0, 0, 2, 0, 1, 3, 0, 0, 1, 0, 0x01];
    let template = parse_shared_formula_template(0x04bc, &data).expect("valid ShrFmla");
    assert_eq!(template.first_row, 0);
    assert_eq!(template.last_row, 2);
    assert_eq!(template.first_col, 1);
    assert_eq!(template.last_col, 3);
    assert!(template.contains(1, 2));
    assert!(!template.contains(3, 2));
}

#[test]
fn shared_formula_wire_parser_rejects_empty_tokens() {
    let data = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
    let error = parse_shared_formula_template(0x04bc, &data).expect_err("empty token stream");
    assert!(error.to_string().contains("token stream is empty"));
}

#[test]
fn cell_record_xf_is_variant_independent() {
    let record = CellRecord::Blank {
        row: 0,
        col: 0,
        xf_index: 37,
    };
    assert_eq!(cell_record_xf(&record), 37);
}

#[test]
fn pivot_cache_paths_are_case_insensitive_and_sorted() {
    let paths = pivot_cache_stream_paths([
        vec!["_SX_DB_CUR".into(), "0002".into()],
        vec!["_sx_db_cur".into(), "0001".into()],
        vec!["_SX_DB_CUR".into(), "not-hex".into()],
    ]);
    assert_eq!(
        paths.iter().map(|(id, _)| *id).collect::<Vec<_>>(),
        vec![1, 2]
    );
}
