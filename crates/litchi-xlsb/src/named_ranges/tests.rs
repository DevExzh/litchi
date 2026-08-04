use super::{Definition, area3d_formula, validate_name};
use crate::formula::{CellParsedFormula, ptg_types};

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
        &CellParsedFormula {
            rgce: vec![ptg_types::PTG_INT, 1, 0],
            rgcb: Vec::new(),
        }
        .to_bytes()
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
    let parsed = Definition::parse(&record).unwrap();
    assert_eq!(parsed.name, "SalesData");
    assert!(parsed.hidden);
    assert!(!parsed.function);
    assert_eq!(parsed.formula, Some(vec![ptg_types::PTG_INT, 1, 0]));

    let mut reserved = record.clone();
    reserved[3] = 0x80;
    assert!(Definition::parse(&reserved).is_err());

    let mut shortcut = record.clone();
    shortcut[4] = b'A';
    assert!(Definition::parse(&shortcut).is_err());

    let mut trailing = record.clone();
    trailing.push(0);
    assert!(Definition::parse(&trailing).is_err());

    let mut macro_record = name_record(0x0008, b'A', "MacroName");
    for _ in 0..4 {
        macro_record.extend_from_slice(&u32::MAX.to_le_bytes());
    }
    assert!(Definition::parse(&macro_record).unwrap().function);

    let mut truncated_comment = record;
    let comment_offset = truncated_comment.len() - 4;
    truncated_comment[comment_offset..].copy_from_slice(&1_u32.to_le_bytes());
    assert!(Definition::parse(&truncated_comment).is_err());
}
