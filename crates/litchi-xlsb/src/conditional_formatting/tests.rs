#![allow(
    clippy::expect_used,
    reason = "test fixture uses bounded literal casts, panic-on-failure extraction, exact floating sentinels, or explicit negative fallback solely to state its assertion"
)]

//! Public-facade regression coverage for the conditional-formatting owner.

use super::*;
use crate::raw::Writer;

#[test]
fn canonical_facade_builds_a_classic_rule() {
    let mut formatting = Formatting::new(vec!["A1:A10".to_string()]);
    formatting.add_rule(Rule::new(RuleType::CellIs, 1));

    assert_eq!(formatting.ranges, ["A1:A10"]);
    assert_eq!(formatting.rules[0].rule_type, RuleType::CellIs);
}

#[test]
fn canonical_facade_exposes_owner_result_and_codec() {
    let mut formatting = Formatting::new(vec!["A1".to_string()]);
    let mut rule = Rule::new(RuleType::Expression, 1);
    rule.formula_texts.push("1".to_string());
    formatting.add_rule(rule);
    let mut bytes = Vec::new();

    let result: Result<()> =
        write_conditional_formattings(&mut Writer::new(&mut bytes), &[formatting]);

    result.expect("canonical owner codec should write the rule");
    assert!(!bytes.is_empty());
}
