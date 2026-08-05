//! Focused formula codec round-trip and bounded-input tests.

use super::super::model::{
    Group, GroupKind, ParsedFormula, Range, TableColumns, TableDataType, TableReference,
    TableRowType, Token, ptg_types,
};
use super::super::{Error, MAX_CELL_FORMULA_BYTES};
use super::Parser;

#[test]
fn parses_and_serializes_cell_formula_lengths() {
    // [MS-XLSB] 2.5.98.4: cce, rgce, cb, and rgbExtra.
    let formula = ParsedFormula {
        rgce: vec![ptg_types::PTG_INT, 42, 0],
        rgcb: vec![0xAA, 0xBB],
    };
    let encoded = formula.to_bytes().unwrap();
    let (decoded, consumed) = ParsedFormula::parse(&encoded).unwrap();
    assert_eq!(consumed, encoded.len());
    assert_eq!(decoded, formula);
}

#[test]
fn parses_scalar_ptgs_and_preserves_unknown_tokens() {
    // [MS-XLSB] 2.5.98.34, 2.5.98.63, and the extensible Ptg prefix.
    let mut data = vec![ptg_types::PTG_BOOL, 1, ptg_types::PTG_NUM];
    data.extend_from_slice(&42.5_f64.to_le_bytes());
    data.extend_from_slice(&[0x7F, 0x7E]);
    let mut parser = Parser::new(&data);
    let tokens = parser.parse().unwrap();
    assert!(matches!(tokens[0], Token::Bool(true)));
    assert!(matches!(tokens[1], Token::Number(value) if value == 42.5));
    assert!(matches!(tokens[2], Token::Unknown(0x7F)));
    assert!(matches!(tokens[3], Token::Unknown(0x7E)));
}

#[test]
fn parses_and_rejects_grouped_formula_records_without_unbounded_allocations() {
    let formula = ParsedFormula::exp(3, 4).unwrap();
    let group = Group {
        kind: GroupKind::Array,
        range: Range::new(3, 3, 4, 4).unwrap(),
        formula,
        always_calculate: true,
    };
    let data = group.to_record_data().unwrap();
    assert_eq!(Group::parse_array(&data).unwrap(), group);

    let mut oversized = vec![0u8; MAX_CELL_FORMULA_BYTES + 1 + 8];
    oversized[..4].copy_from_slice(
        &u32::try_from(MAX_CELL_FORMULA_BYTES + 1)
            .unwrap()
            .to_le_bytes(),
    );
    assert!(matches!(
        ParsedFormula::parse(&oversized),
        Err(Error::InvalidFormula(message)) if message.contains("exceeds")
    ));
}

#[test]
fn extended_table_token_round_trips() {
    let token = Token::TableReference(TableReference {
        sheet_index: 2,
        row_type: Some(TableRowType::Data),
        columns: Some(TableColumns::One(1)),
        square_bracket_space: false,
        comma_space: true,
        data_type: TableDataType::Reference,
        invalid: false,
        list_index: Some(7),
        external: None,
    });
    let (rgce, rgcb) = token.to_extended_binary().unwrap();
    let mut parser = Parser::with_extra(&rgce, &rgcb);
    assert_eq!(parser.parse().unwrap(), vec![token]);
}
