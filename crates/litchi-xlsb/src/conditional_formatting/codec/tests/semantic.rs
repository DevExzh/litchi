//! Semantic validation and formula-boundary tests.

use super::super::super::model::RuleType;
use super::super::Error;
use super::super::semantic::{TextCompiler, validate_formula_count, validate_template};

#[test]
fn formula_compiler_emits_bounded_tokens() {
    let formula = TextCompiler::compile("=A1+1").unwrap();
    assert!(!formula.rgce.is_empty());
    assert!(formula.rgce.len() <= crate::formula::MAX_CELL_FORMULA_BYTES);
}

#[test]
fn formula_compiler_rejects_unsupported_constructs() {
    let error = TextCompiler::compile("SUM(A1)").unwrap_err();
    assert!(matches!(error, Error::UnsupportedFeature(_)));
}

#[test]
fn template_and_formula_cardinality_are_validated_semantically() {
    assert!(validate_template(RuleType::CellIs, 0).is_ok());
    assert!(validate_template(RuleType::CellIs, 2).is_err());
    assert!(validate_formula_count(RuleType::CellIs, 0, 1, 2).is_ok());
    assert!(validate_formula_count(RuleType::CellIs, 0, 1, 1).is_err());
}
