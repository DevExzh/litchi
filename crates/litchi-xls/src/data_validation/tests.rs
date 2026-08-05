//! Regression coverage for the BIFF8 data-validation owner.

use super::codec::{parse_dv, parse_dval};
use super::model::{ErrorStyle, Kind, Operator};
use crate::Result;

fn string(data: &mut Vec<u8>, value: &str) {
    data.extend_from_slice(&(value.len() as u16).to_le_bytes());
    data.push(0);
    data.extend_from_slice(value.as_bytes());
}

fn formula(data: &mut Vec<u8>, tokens: &[u8]) {
    data.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
    data.extend_from_slice(&0u16.to_le_bytes());
    data.extend_from_slice(tokens);
}

fn valid_dv() -> Vec<u8> {
    let options = 1u32 | (1 << 4) | (1 << 8) | (1 << 18) | (1 << 19);
    let mut data = options.to_le_bytes().to_vec();
    string(&mut data, "Input");
    string(&mut data, "Error");
    string(&mut data, "Enter a value");
    string(&mut data, "Invalid");
    formula(&mut data, &[0x1E, 1, 0]);
    formula(&mut data, &[0x1E, 10, 0]);
    data.extend_from_slice(&1u16.to_le_bytes());
    for value in [2u16, 4, 3, 5] {
        data.extend_from_slice(&value.to_le_bytes());
    }
    data
}

fn assert_rejected_without_panicking<T>(parse: impl FnOnce() -> Result<T>) {
    match std::panic::catch_unwind(std::panic::AssertUnwindSafe(parse)) {
        Ok(result) => assert!(result.is_err()),
        Err(_) => panic!("malformed data-validation input must not panic"),
    }
}

#[test]
fn parses_dval_and_dv_with_raw_formulas() {
    let mut dval = Vec::new();
    dval.extend_from_slice(&1u16.to_le_bytes());
    dval.extend_from_slice(&10u32.to_le_bytes());
    dval.extend_from_slice(&20u32.to_le_bytes());
    dval.extend_from_slice(&(-1i32).to_le_bytes());
    dval.extend_from_slice(&1u32.to_le_bytes());
    let settings = parse_dval(&dval).unwrap();
    assert!(settings.window_closed());
    assert_eq!(settings.declared_rule_count(), 1);

    let rule = parse_dv(&valid_dv(), None).unwrap();
    assert_eq!(rule.kind(), Kind::Whole);
    assert_eq!(rule.error_style(), ErrorStyle::Warning);
    assert_eq!(rule.formula1().unwrap().tokens(), &[0x1E, 1, 0]);
    assert_eq!(rule.formula1().unwrap().rendered(), Some("=1"));
    assert_eq!(rule.formula2().unwrap().tokens(), &[0x1E, 10, 0]);
    assert_eq!(rule.ranges()[0].first_row(), 2);
    assert_eq!(rule.ranges()[0].last_column(), 5);
}

#[test]
fn rejects_reserved_bits_and_malformed_rule_shape() {
    let mut dval = [0u8; 18];
    dval[0] = 2;
    assert!(parse_dval(&dval).is_err());
    let mut dv = valid_dv();
    dv[3] = 0x80;
    assert!(parse_dv(&dv, None).is_err());
    let mut dv = valid_dv();
    let end = dv.len();
    dv[end - 8..end - 6].copy_from_slice(&5u16.to_le_bytes());
    dv[end - 6..end - 4].copy_from_slice(&4u16.to_le_bytes());
    assert!(parse_dv(&dv, None).is_err());
}

#[test]
fn rejects_truncated_dval_payloads_without_panicking() {
    for length in 0..18 {
        let data = vec![0u8; length];
        assert_rejected_without_panicking(|| parse_dval(&data));
    }
}

#[test]
fn rejects_truncated_dv_fields_without_panicking() {
    let data = valid_dv();
    for length in 0..data.len() {
        let truncated = &data[..length];
        assert_rejected_without_panicking(|| parse_dv(truncated, None));
    }
}

#[test]
fn enforces_formula_cardinality_and_range_limit() {
    let mut dv = valid_dv();
    dv[0] = 0;
    assert!(parse_dv(&dv, None).is_err());
    let mut data = 3u32.to_le_bytes().to_vec();
    for _ in 0..4 {
        string(&mut data, "\0");
    }
    formula(&mut data, &[0x17, 1, 0, b'A']);
    formula(&mut data, &[]);
    data.extend_from_slice(&0u16.to_le_bytes());
    assert!(parse_dv(&data, None).is_err());
}

#[test]
fn ignores_undefined_operator_bits_for_list_validations() {
    let options = 3u32 | (15 << 20);
    let mut data = options.to_le_bytes().to_vec();
    for _ in 0..4 {
        string(&mut data, "\0");
    }
    formula(&mut data, &[0x17, 1, 0, b'A']);
    formula(&mut data, &[]);
    data.extend_from_slice(&1u16.to_le_bytes());
    for value in [0u16, 0, 0, 0] {
        data.extend_from_slice(&value.to_le_bytes());
    }
    let rule = parse_dv(&data, None).unwrap();
    assert_eq!(rule.kind(), Kind::List);
    assert_eq!(rule.operator(), Operator::Between);
}
