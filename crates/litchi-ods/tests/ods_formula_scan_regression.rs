//! Compatibility checks for the legacy formula scan boundaries.
//!
//! These cases exercise token and error semantics around inputs that can look
//! like one compact identifier while the established parser intentionally
//! emits several partial cell references.

use litchi_core::Error;
use litchi_ods::codec::formula::{CellRef, FormulaParser, Token};

fn assert_cell(
    token: &Token,
    expected_sheet: Option<&str>,
    expected_column: &str,
    expected_row: u32,
    expected_column_absolute: bool,
    expected_row_absolute: bool,
) {
    let Token::CellRef(cell) = token else {
        panic!("expected a cell token, got {token:?}");
    };
    assert_cell_ref(
        cell,
        expected_sheet,
        expected_column,
        expected_row,
        expected_column_absolute,
        expected_row_absolute,
    );
}

fn assert_cell_ref(
    cell: &CellRef,
    expected_sheet: Option<&str>,
    expected_column: &str,
    expected_row: u32,
    expected_column_absolute: bool,
    expected_row_absolute: bool,
) {
    assert_eq!(cell.sheet.as_deref(), expected_sheet);
    assert_eq!(cell.column, expected_column);
    assert_eq!(cell.row, expected_row);
    assert_eq!(cell.column_absolute, expected_column_absolute);
    assert_eq!(cell.row_absolute, expected_row_absolute);
}

fn assert_invalid_format(source: &str) {
    let error = FormulaParser::new(source)
        .parse()
        .expect_err("malformed legacy scan suffix should be rejected");
    assert!(
        matches!(error, Error::InvalidFormat(_)),
        "unexpected error for {source:?}: {error}"
    );
}

#[test]
fn repeated_partial_cells_keep_boundaries_with_spaces_and_tabs() {
    let source = "=A1 A1\tA1\nA1";
    let formula = FormulaParser::new(source)
        .parse()
        .expect("whitespace-separated partial cells should remain independent");

    assert_eq!(formula.text, source);
    assert_eq!(formula.tokens.len(), 4);
    for token in &formula.tokens {
        assert_cell(token, None, "A", 1, false, false);
    }
}

#[test]
fn contiguous_partial_cells_and_a_dot_tail_keep_legacy_tokens() {
    let source = "=A1A1A1A1";
    let formula = FormulaParser::new(source)
        .parse()
        .expect("contiguous cell-shaped text should retain partial-cell tokens");

    assert_eq!(formula.text, source);
    assert_eq!(formula.tokens.len(), 4);
    for token in &formula.tokens {
        assert_cell(token, None, "A", 1, false, false);
    }

    let source = "=A1A1A1 .B2";
    let formula = FormulaParser::new(source)
        .parse()
        .expect("a separated current-sheet dot tail should remain parseable");
    assert_eq!(formula.text, source);
    assert_eq!(formula.tokens.len(), 1);
    assert_cell(&formula.tokens[0], Some("A1A1A1 "), "B", 2, false, false);

    let source = "=A1A1A1.A1";
    let formula = FormulaParser::new(source)
        .parse()
        .expect("a dot directly after a partial-cell run should retain its sheet prefix");
    assert_eq!(formula.text, source);
    assert_eq!(formula.tokens.len(), 1);
    assert_cell(&formula.tokens[0], Some("A1A1A1"), "A", 1, false, false);
}

#[test]
fn tabs_underscores_and_digit_boundaries_preserve_tokens_and_refusals() {
    let source = "=A1\tSheet_2.C3+AA10 2";
    let formula = FormulaParser::new(source)
        .parse()
        .expect("tab, underscore, and digit boundaries should parse");

    assert_eq!(formula.text, source);
    assert_eq!(formula.tokens.len(), 5);
    assert_cell(&formula.tokens[0], None, "A", 1, false, false);
    assert_cell(&formula.tokens[1], Some("Sheet_2"), "C", 3, false, false);
    assert!(matches!(formula.tokens[2], Token::Operator('+')));
    assert_cell(&formula.tokens[3], None, "AA", 10, false, false);
    assert!(matches!(formula.tokens[4], Token::Number(value) if value == 2.0));

    for source in ["=A1_A2", "=A1.2", "=A1..B2"] {
        assert_invalid_format(source);
    }
}

#[test]
fn a_dot_after_a_long_suffix_uses_the_legacy_sheet_fallback() {
    let sheet = "VeryLongSheetName_0123456789";
    let source = "=VeryLongSheetName_0123456789.A1";
    let formula = FormulaParser::new(source)
        .parse()
        .expect("a long sheet prefix followed by a dot should parse");

    assert_eq!(formula.text, source);
    assert_eq!(formula.tokens.len(), 1);
    assert_cell(&formula.tokens[0], Some(sheet), "A", 1, false, false);
}

#[test]
fn absolute_and_range_forms_retain_the_legacy_fallback_shape() {
    let absolute = FormulaParser::new("=$a$1")
        .parse()
        .expect("absolute cell should use the legacy parser");
    assert_eq!(absolute.tokens.len(), 1);
    assert_cell(&absolute.tokens[0], None, "A", 1, true, true);

    let source = "=Sheet_2.$a$1:.$B$2";
    let formula = FormulaParser::new(source)
        .parse()
        .expect("absolute range should retain both fallback endpoints");
    assert_eq!(formula.text, source);
    assert_eq!(formula.tokens.len(), 1);
    let Token::RangeRef(range) = &formula.tokens[0] else {
        panic!("expected a range token, got {:?}", formula.tokens[0]);
    };
    assert_cell_ref(&range.start, Some("Sheet_2"), "A", 1, true, true);
    assert_cell_ref(&range.end, None, "B", 2, true, true);
}

#[test]
fn formula_prefixes_normalize_for_scanning_but_retain_original_text() {
    for source in ["=A1+$b$2", "OF:=A1+$b$2", "  of:=A1+$b$2  "] {
        let formula = FormulaParser::new(source)
            .parse()
            .expect("supported formula prefixes should share token semantics");
        assert_eq!(formula.text, source);
        assert_eq!(formula.tokens.len(), 3);
        assert_cell(&formula.tokens[0], None, "A", 1, false, false);
        assert!(matches!(formula.tokens[1], Token::Operator('+')));
        assert_cell(&formula.tokens[2], None, "B", 2, true, true);
    }
}
