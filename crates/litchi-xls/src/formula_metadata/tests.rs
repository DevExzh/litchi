//! Focused Formula-record metadata regressions.

use super::Metadata;
use super::codec::parse_record;
use super::validation::{decode_flags, encode_flags, is_ptg_exp};

fn payload(flags: u16, cache: u32, tokens: &[u8]) -> Vec<u8> {
    let mut data = Vec::with_capacity(22 + tokens.len());
    data.extend_from_slice(&4u16.to_le_bytes());
    data.extend_from_slice(&7u16.to_le_bytes());
    data.extend_from_slice(&9u16.to_le_bytes());
    data.extend_from_slice(&0f64.to_le_bytes());
    data.extend_from_slice(&flags.to_le_bytes());
    data.extend_from_slice(&cache.to_le_bytes());
    data.extend_from_slice(&(tokens.len() as u16).to_le_bytes());
    data.extend_from_slice(tokens);
    data
}

#[test]
fn parses_all_formula_flags_and_application_cache() {
    let parsed = parse_record(&payload(0x002D, 0x1234_5678, &[0x01, 4, 0, 7, 0])).unwrap();
    assert_eq!(parsed.row, 4);
    assert_eq!(parsed.col, 7);
    assert_eq!(parsed.xf_index, 9);
    assert!(parsed.metadata.always_calculate());
    assert!(parsed.metadata.fill_alignment());
    assert!(parsed.metadata.shared_formula());
    assert!(parsed.metadata.clear_errors());
    assert_eq!(parsed.metadata.calculation_cache(), 0x1234_5678);
}

#[test]
fn rejects_reserved_flags_empty_tokens_and_trailing_payload() {
    assert!(decode_flags(0x0002, &[0x16]).is_err());
    assert!(parse_record(&payload(0, 0, &[])).is_err());

    let mut data = payload(0, 0, &[0x16]);
    data.push(0xFF);
    assert!(parse_record(&data).is_err());
}

#[test]
fn string_cached_values_allow_an_empty_formula_token_stream() {
    let mut data = payload(0, 0, &[]);
    data[6] = 0;
    data[12..14].copy_from_slice(&[0xFF, 0xFF]);

    let parsed = parse_record(&data).unwrap();
    assert!(matches!(
        parsed.value,
        crate::records::FormulaValue::StringPending
    ));
    assert!(parsed.formula.is_empty());
}

#[test]
fn shared_flag_requires_ptg_exp_and_writer_refuses_orphaned_shared_state() {
    assert!(!is_ptg_exp(&[0x16, 0, 0, 0, 0]));
    assert!(decode_flags(0x0008, &[0x16]).is_err());
    assert!(
        encode_flags(
            Metadata::new().with_shared_formula(true),
            &[0x01, 0, 0, 0, 0]
        )
        .is_err()
    );
}

#[test]
fn metadata_builders_are_copyable_and_wire_stable() {
    let metadata = Metadata::new()
        .with_always_calculate(true)
        .with_fill_alignment(true)
        .with_clear_errors(true)
        .with_calculation_cache(17);
    assert_eq!(metadata, metadata);
    assert_eq!(encode_flags(metadata, &[0x16]).unwrap(), 0x0025);
}
