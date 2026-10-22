#![allow(
    clippy::cast_possible_truncation,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]

use super::{Definition, area3d_formula, validate_name};
use crate::formula::{ParsedFormula, ptg_types};
use crate::raw::Limits;

fn name_record(flags: u32, ch_key: u8, name: &str) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&flags.to_le_bytes());
    data.push(ch_key);
    data.extend_from_slice(&u32::MAX.to_le_bytes());
    let utf16: Vec<_> = name.encode_utf16().collect();
    data.extend_from_slice(&(utf16.len() as u32).to_le_bytes());
    for code_unit in utf16 {
        data.extend_from_slice(&code_unit.to_le_bytes());
    }
    data.extend_from_slice(
        &ParsedFormula {
            rgce: vec![ptg_types::PTG_INT, 1, 0],
            rgcb: Vec::new(),
        }
        .to_bytes(Limits::DEFAULT)
        .unwrap(),
    );
    data.extend_from_slice(&u32::MAX.to_le_bytes());
    data
}

#[test]
fn builder_preserves_defined_name_state() {
    let range = Definition::new("MyRange".to_string(), None)
        .with_hidden(true)
        .with_formula(vec![1, 2, 3]);

    assert_eq!(range.name, "MyRange");
    assert!(range.hidden);
    assert_eq!(range.formula, Some(vec![1, 2, 3]));
}

#[test]
fn creates_area3d_formula() {
    let formula = area3d_formula(0, 1, 3, 1, 1).unwrap();
    assert_eq!(formula[0], ptg_types::PTG_AREA_3D);
    assert_eq!(u16::from_le_bytes([formula[1], formula[2]]), 2);
    assert_eq!(u32::from_le_bytes(formula[3..7].try_into().unwrap()), 1);
    assert_eq!(u32::from_le_bytes(formula[7..11].try_into().unwrap()), 3);
    assert_eq!(u16::from_le_bytes([formula[11], formula[12]]), 1);
    assert_eq!(u16::from_le_bytes([formula[13], formula[14]]), 1);
}

#[test]
fn validates_defined_name_grammar() {
    for name in ["SalesData", "_rate.2026", "\\Print_Area", "数据1"] {
        validate_name(name).unwrap();
    }
    for name in [
        "",
        "1Sales",
        "Sales Data",
        "TRUE",
        "xfd1048576",
        "R1total",
        "C16384x",
    ] {
        assert!(validate_name(name).is_err(), "accepted {name:?}");
    }
    validate_name("XFE1").unwrap();
    validate_name("R1048577total").unwrap();
    validate_name("C16385x").unwrap();
}

#[test]
fn parses_complete_brt_name_and_rejects_malformed_records() {
    let record = name_record(1, 0, "SalesData");
    let parsed = Definition::parse(&record, Limits::DEFAULT).unwrap();
    assert_eq!(parsed.name, "SalesData");
    assert!(parsed.hidden);
    assert!(!parsed.function);
    assert_eq!(parsed.formula, Some(vec![ptg_types::PTG_INT, 1, 0]));

    let mut reserved = record.clone();
    reserved[3] = 0x80;
    assert!(Definition::parse(&reserved, Limits::DEFAULT).is_err());

    let mut shortcut = record.clone();
    shortcut[4] = b'A';
    assert!(Definition::parse(&shortcut, Limits::DEFAULT).is_err());

    let mut trailing = record.clone();
    trailing.push(0);
    assert!(Definition::parse(&trailing, Limits::DEFAULT).is_err());

    let mut macro_record = name_record(0x0008, b'A', "MacroName");
    for _ in 0..4 {
        macro_record.extend_from_slice(&u32::MAX.to_le_bytes());
    }
    assert!(
        Definition::parse(&macro_record, Limits::DEFAULT)
            .unwrap()
            .function
    );

    let mut truncated_comment = record;
    let comment_offset = truncated_comment.len() - 4;
    truncated_comment[comment_offset..].copy_from_slice(&1_u32.to_le_bytes());
    assert!(Definition::parse(&truncated_comment, Limits::DEFAULT).is_err());
}

#[test]
fn parse_uses_the_required_payload_and_string_limits() {
    let record = name_record(0, 0, "BoundedName");
    let payload_limit = Limits::new(record.len() - 1, Limits::DEFAULT.string_units());
    let error = Definition::parse(&record, payload_limit).unwrap_err();
    assert!(matches!(
        error,
        super::Error::Wire(crate::raw::Error::PayloadLimit {
            length,
            limit,
            ..
        }) if length == record.len() && limit == record.len() - 1
    ));

    let string_limit = Limits::new(record.len(), 1);
    let error = Definition::parse(&record, string_limit).unwrap_err();
    assert!(matches!(
        error,
        super::Error::Wire(crate::raw::Error::StringLimit { units, limit, .. })
            if units > limit && limit == 1
    ));

    let parsed = Definition::parse(&record, Limits::new(record.len(), 64)).unwrap();
    assert_eq!(parsed.name, "BoundedName");
}
