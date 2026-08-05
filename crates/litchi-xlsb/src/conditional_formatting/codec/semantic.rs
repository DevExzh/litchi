//! Semantic conditional-formatting validation and formula boundary.
//!
//! This layer owns source-formula compilation, model invariants, and
//! cross-record relationships. It deliberately does not write or consume
//! Brt* payloads.

use crate::formula::{
    ArrayValue, Compiler, MAX_CELL_FORMULA_BYTES, ParsedFormula, Parser, Resolution,
};
use std::collections::{HashMap, HashSet};

use super::super::model::*;
use super::{Error, Result, invalid};

#[derive(Debug, Clone, Copy, Default)]
pub(super) struct EmptyFormulaResolution;

impl Resolution for EmptyFormulaResolution {
    fn sheet_prefix(&self, index: u16) -> crate::formula::Result<String> {
        Err(crate::formula::Error::InvalidFormula(format!(
            "formula references unresolved sheet index {index}"
        )))
    }

    fn defined_name(&self, index: u32) -> crate::formula::Result<String> {
        Err(crate::formula::Error::InvalidFormula(format!(
            "formula references unresolved defined name {index}"
        )))
    }

    fn external_name(&self, sheet_index: u16, name_index: u32) -> crate::formula::Result<String> {
        Err(crate::formula::Error::InvalidFormula(format!(
            "formula references unresolved external name {sheet_index}:{name_index}"
        )))
    }

    fn table_reference(
        &self,
        _reference: &crate::formula::TableReference,
    ) -> crate::formula::Result<String> {
        Err(crate::formula::Error::InvalidFormula(
            "formula references unresolved table".to_string(),
        ))
    }

    fn pivot_name(&self, index: u32) -> crate::formula::Result<String> {
        Err(crate::formula::Error::InvalidFormula(format!(
            "formula references unresolved pivot name {index}"
        )))
    }
}

/// Small, owner-local formula emitter used when a conditional-formatting
/// model was authored from formula text instead of retained binary tokens.
///
/// The full workbook formula compiler remains a host concern. Conditional
/// formatting only needs the bounded literal/reference/operator subset here;
/// unsupported constructs are rejected rather than silently rewritten.
pub(super) struct TextCompiler;

impl TextCompiler {
    pub(super) fn compile(input: &str) -> Result<ParsedFormula> {
        let input = input.strip_prefix('=').unwrap_or(input).trim();
        if input.is_empty() {
            return Err(Error::InvalidFormula(
                "conditional-format formula is empty".to_string(),
            ));
        }
        let mut rgce = Vec::new();
        let mut rgcb = Vec::new();
        compile_formula_expression(input, &mut rgce, &mut rgcb)?;
        if rgce.is_empty() || rgce.len() > MAX_CELL_FORMULA_BYTES {
            return Err(Error::InvalidFormula(format!(
                "compiled conditional-format formula length {} is outside 1..={MAX_CELL_FORMULA_BYTES}",
                rgce.len()
            )));
        }
        Ok(ParsedFormula { rgce, rgcb })
    }
}

fn compile_formula_expression(input: &str, rgce: &mut Vec<u8>, rgcb: &mut Vec<u8>) -> Result<()> {
    let input = input.trim();
    if input.is_empty() {
        return Err(Error::InvalidFormula(
            "conditional-format formula has an empty operand".to_string(),
        ));
    }
    if let Some(inner) = strip_outer_parentheses(input) {
        compile_formula_expression(inner, rgce, rgcb)?;
        rgce.push(crate::formula::ptg_types::PTG_PAREN);
        return Ok(());
    }

    for operators in [
        ["<>", "<=", ">=", "=", "<", ">"].as_slice(),
        ["&"].as_slice(),
        ["+", "-"].as_slice(),
        ["*", "/"].as_slice(),
        ["^"].as_slice(),
    ] {
        if let Some((offset, operator)) = find_formula_operator(input, operators) {
            let (left, right) = input.split_at(offset);
            let right = &right[operator.len()..];
            compile_formula_expression(left, rgce, rgcb)?;
            compile_formula_expression(right, rgce, rgcb)?;
            rgce.push(match operator {
                "<>" => crate::formula::ptg_types::PTG_NE,
                "<=" => crate::formula::ptg_types::PTG_LE,
                ">=" => crate::formula::ptg_types::PTG_GE,
                "=" => crate::formula::ptg_types::PTG_EQ,
                "<" => crate::formula::ptg_types::PTG_LT,
                ">" => crate::formula::ptg_types::PTG_GT,
                "&" => crate::formula::ptg_types::PTG_CONCAT,
                "+" => crate::formula::ptg_types::PTG_ADD,
                "-" => crate::formula::ptg_types::PTG_SUB,
                "*" => crate::formula::ptg_types::PTG_MUL,
                "/" => crate::formula::ptg_types::PTG_DIV,
                "^" => crate::formula::ptg_types::PTG_POWER,
                _ => unreachable!("operator was selected from the fixed table"),
            });
            return Ok(());
        }
    }

    if let Some(rest) = input.strip_prefix('+') {
        compile_formula_expression(rest, rgce, rgcb)?;
        rgce.push(crate::formula::ptg_types::PTG_UPLUS);
        return Ok(());
    }
    if let Some(rest) = input.strip_prefix('-') {
        compile_formula_expression(rest, rgce, rgcb)?;
        rgce.push(crate::formula::ptg_types::PTG_UMINUS);
        return Ok(());
    }
    if let Some(rest) = input.strip_suffix('%') {
        compile_formula_expression(rest, rgce, rgcb)?;
        rgce.push(crate::formula::ptg_types::PTG_PERCENT);
        return Ok(());
    }

    if input.starts_with('{') && input.ends_with('}') {
        return compile_formula_array(&input[1..input.len() - 1], rgce, rgcb);
    }
    if let Some(value) = parse_formula_string(input)? {
        let units = value.encode_utf16().count();
        let units = u16::try_from(units)
            .map_err(|_| Error::InvalidFormula("formula string is too long".to_string()))?;
        rgce.push(crate::formula::ptg_types::PTG_STR);
        rgce.extend_from_slice(&units.to_le_bytes());
        rgce.extend(value.encode_utf16().flat_map(u16::to_le_bytes));
        return Ok(());
    }
    if input.eq_ignore_ascii_case("TRUE") || input.eq_ignore_ascii_case("FALSE") {
        rgce.extend([
            crate::formula::ptg_types::PTG_BOOL,
            u8::from(input.eq_ignore_ascii_case("TRUE")),
        ]);
        return Ok(());
    }
    if let Some(error) = formula_error_code(input) {
        rgce.extend([crate::formula::ptg_types::PTG_ERR, error]);
        return Ok(());
    }
    if let Ok(number) = input.parse::<f64>() {
        if !number.is_finite() {
            return Err(Error::InvalidFormula(
                "formula number is not finite".to_string(),
            ));
        }
        if number.fract() == 0.0 && (0.0..=f64::from(u16::MAX)).contains(&number) {
            rgce.push(crate::formula::ptg_types::PTG_INT);
            rgce.extend_from_slice(&(number as u16).to_le_bytes());
        } else {
            rgce.push(crate::formula::ptg_types::PTG_NUM);
            rgce.extend_from_slice(&number.to_le_bytes());
        }
        return Ok(());
    }

    if let Some((row, _col, bits)) = parse_formula_reference(input) {
        rgce.push(crate::formula::ptg_types::PTG_REF | 0x20);
        rgce.extend_from_slice(&row.to_le_bytes());
        rgce.extend_from_slice(&bits.to_le_bytes());
        return Ok(());
    }
    if let Some((first, last)) = input.split_once(':') {
        let Some((first_row, first_col, first_bits)) = parse_formula_reference(first.trim()) else {
            return Err(Error::InvalidFormula(format!(
                "invalid conditional-format range {input:?}"
            )));
        };
        let Some((last_row, last_col, last_bits)) = parse_formula_reference(last.trim()) else {
            return Err(Error::InvalidFormula(format!(
                "invalid conditional-format range {input:?}"
            )));
        };
        if first_row > last_row || first_col > last_col {
            return Err(Error::InvalidFormula(
                "conditional-format range is reversed".to_string(),
            ));
        }
        rgce.push(crate::formula::ptg_types::PTG_AREA | 0x20);
        rgce.extend_from_slice(&first_row.to_le_bytes());
        rgce.extend_from_slice(&last_row.to_le_bytes());
        rgce.extend_from_slice(&first_bits.to_le_bytes());
        rgce.extend_from_slice(&last_bits.to_le_bytes());
        return Ok(());
    }

    Err(Error::UnsupportedFeature(format!(
        "conditional-format formula construct {input:?} is not supported by the owner-local emitter"
    )))
}

fn compile_formula_array(input: &str, rgce: &mut Vec<u8>, rgcb: &mut Vec<u8>) -> Result<()> {
    let rows = split_formula_list(input, ';');
    if rows.is_empty() || rows.iter().any(|row| row.trim().is_empty()) {
        return Err(Error::InvalidFormula(
            "conditional-format array has an empty row".to_string(),
        ));
    }
    let columns = rows
        .iter()
        .map(|row| split_formula_list(row, ','))
        .collect::<Vec<_>>();
    let column_count = columns[0].len();
    if column_count == 0 || columns.iter().any(|row| row.len() != column_count) {
        return Err(Error::InvalidFormula(
            "conditional-format array rows have different widths".to_string(),
        ));
    }
    let row_count = u32::try_from(columns.len())
        .map_err(|_| Error::InvalidFormula("array row count overflow".to_string()))?;
    let column_count = u32::try_from(column_count)
        .map_err(|_| Error::InvalidFormula("array column count overflow".to_string()))?;
    if row_count > 1_048_576 || column_count > 16_384 {
        return Err(Error::InvalidFormula(
            "conditional-format array exceeds worksheet bounds".to_string(),
        ));
    }

    let mut values = Vec::new();
    for row in columns {
        for value in row {
            values.push(parse_formula_array_value(value.trim())?);
        }
    }
    // PtgArray uses the VALUE operand class (0x40); 0x20 is the base token
    // value and is not a valid array token on the BIFF12 wire.
    rgce.push(0x40);
    rgce.extend([0; 14]);
    rgcb.extend_from_slice(&row_count.to_le_bytes());
    rgcb.extend_from_slice(&column_count.to_le_bytes());
    for value in values {
        match value {
            ArrayValue::Number(value) => {
                if !value.is_finite() {
                    return Err(Error::InvalidFormula(
                        "conditional-format array number is not finite".to_string(),
                    ));
                }
                rgcb.push(0);
                rgcb.extend_from_slice(&value.to_le_bytes());
            },
            ArrayValue::String(value) => {
                let units = u16::try_from(value.encode_utf16().count()).map_err(|_| {
                    Error::InvalidFormula("conditional-format array string is too long".to_string())
                })?;
                rgcb.push(1);
                rgcb.extend_from_slice(&units.to_le_bytes());
                rgcb.extend(value.encode_utf16().flat_map(u16::to_le_bytes));
            },
            ArrayValue::Bool(value) => rgcb.extend([2, u8::from(value)]),
            ArrayValue::Error(value) => rgcb.extend([4, value, 0, 0, 0]),
        }
    }
    Ok(())
}

fn parse_formula_array_value(input: &str) -> Result<ArrayValue> {
    if let Some(value) = parse_formula_string(input)? {
        return Ok(ArrayValue::String(value));
    }
    if input.eq_ignore_ascii_case("TRUE") || input.eq_ignore_ascii_case("FALSE") {
        return Ok(ArrayValue::Bool(input.eq_ignore_ascii_case("TRUE")));
    }
    if let Some(error) = formula_error_code(input) {
        return Ok(ArrayValue::Error(error));
    }
    let value = input.parse::<f64>().map_err(|_| {
        Error::InvalidFormula(format!("invalid conditional-format array value {input:?}"))
    })?;
    Ok(ArrayValue::Number(value))
}

fn parse_formula_string(input: &str) -> Result<Option<String>> {
    if !input.starts_with('"') {
        return Ok(None);
    }
    if !input.ends_with('"') || input.len() < 2 {
        return Err(Error::InvalidFormula(
            "unterminated conditional-format string".to_string(),
        ));
    }
    let mut value = String::new();
    let mut chars = input[1..input.len() - 1].chars().peekable();
    while let Some(character) = chars.next() {
        if character == '"' && chars.peek() == Some(&'"') {
            chars.next();
        }
        value.push(character);
    }
    Ok(Some(value))
}

fn formula_error_code(input: &str) -> Option<u8> {
    [
        ("#NULL!", 0x00),
        ("#DIV/0!", 0x07),
        ("#VALUE!", 0x0F),
        ("#REF!", 0x17),
        ("#NAME?", 0x1D),
        ("#NUM!", 0x24),
        ("#N/A", 0x2A),
        ("#GETTING_DATA", 0x2B),
    ]
    .into_iter()
    .find_map(|(literal, code)| input.eq_ignore_ascii_case(literal).then_some(code))
}

fn parse_formula_reference(input: &str) -> Option<(u32, u32, u16)> {
    let input = input.trim();
    let bytes = input.as_bytes();
    let mut offset = 0;
    let col_relative = bytes.get(offset) != Some(&b'$');
    if !col_relative {
        offset += 1;
    }
    let col_start = offset;
    while bytes.get(offset).is_some_and(u8::is_ascii_alphabetic) {
        offset += 1;
    }
    if offset == col_start {
        return None;
    }
    let mut col = 0_u32;
    for byte in bytes[col_start..offset].iter().map(u8::to_ascii_uppercase) {
        col = col
            .checked_mul(26)?
            .checked_add(u32::from(byte - b'A' + 1))?;
    }
    let row_relative = bytes.get(offset) != Some(&b'$');
    if !row_relative {
        offset += 1;
    }
    let row_start = offset;
    while bytes.get(offset).is_some_and(u8::is_ascii_digit) {
        offset += 1;
    }
    if offset == row_start || offset != bytes.len() {
        return None;
    }
    let row = input[row_start..offset].parse::<u32>().ok()?;
    if col == 0 || col > 16_384 || row == 0 || row > 1_048_576 {
        return None;
    }
    let mut bits = u16::try_from(col - 1).ok()?;
    if col_relative {
        bits |= 0x4000;
    }
    if row_relative {
        bits |= 0x8000;
    }
    Some((row - 1, col - 1, bits))
}

fn split_formula_list(input: &str, separator: char) -> Vec<&str> {
    let mut result = Vec::new();
    let mut start = 0;
    let mut in_string = false;
    for (offset, character) in input.char_indices() {
        if character == '"' {
            in_string = !in_string;
        } else if character == separator && !in_string {
            result.push(&input[start..offset]);
            start = offset + character.len_utf8();
        }
    }
    result.push(&input[start..]);
    result
}

fn find_formula_operator(input: &str, operators: &[&'static str]) -> Option<(usize, &'static str)> {
    let mut parentheses = 0_usize;
    let mut braces = 0_usize;
    let mut in_string = false;
    for (offset, character) in input.char_indices() {
        if character == '"' {
            in_string = !in_string;
            continue;
        }
        if in_string {
            continue;
        }
        match character {
            '(' => parentheses += 1,
            ')' => parentheses = parentheses.saturating_sub(1),
            '{' => braces += 1,
            '}' => braces = braces.saturating_sub(1),
            _ => {},
        }
        if parentheses != 0 || braces != 0 {
            continue;
        }
        for operator in operators {
            if !input[offset..].starts_with(operator) {
                continue;
            }
            if matches!(*operator, "+" | "-") {
                let previous = input[..offset]
                    .chars()
                    .rev()
                    .find(|character| !character.is_whitespace());
                if previous.is_none_or(|character| "+-*/^&=<>(".contains(character))
                    || (*operator == "+" || *operator == "-")
                        && previous.is_some_and(|character| matches!(character, 'e' | 'E'))
                {
                    continue;
                }
            }
            return Some((offset, operator));
        }
    }
    None
}

fn strip_outer_parentheses(input: &str) -> Option<&str> {
    if !input.starts_with('(') || !input.ends_with(')') {
        return None;
    }
    let mut depth = 0_usize;
    let mut in_string = false;
    for (offset, character) in input.char_indices() {
        if character == '"' {
            in_string = !in_string;
        } else if !in_string {
            match character {
                '(' => depth += 1,
                ')' => {
                    depth = depth.checked_sub(1)?;
                    if depth == 0 && offset + character.len_utf8() != input.len() {
                        return None;
                    }
                },
                _ => {},
            }
        }
    }
    (depth == 0).then_some(&input[1..input.len() - 1])
}

/// Return the number of threshold icons defined by an Office 2013 icon set.
pub fn icon_count14(icon_set_type: u8) -> usize {
    match icon_set_type {
        0..=7 | 17 | 18 => 3,
        8..=12 => 4,
        13..=16 | 19 => 5,
        _ => 0,
    }
}

pub fn validate_template(rule_type: RuleType, template: u32) -> Result<()> {
    let valid = match rule_type {
        RuleType::CellIs => template == 0,
        RuleType::Expression => matches!(template, 1 | 7..=12 | 15..=27 | 29 | 30),
        RuleType::ColorScale => template == 2,
        RuleType::DataBar => template == 3,
        RuleType::TopN => template == 5,
        RuleType::IconSet => template == 4,
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "BrtBeginCFRule",
            format!("template {template} is invalid for {rule_type:?}"),
        ))
    }
}

pub(super) fn validate_extension14_template(rule_type: RuleType, template: u32) -> Result<()> {
    if rule_type == RuleType::DataBar && template == 0 {
        Ok(())
    } else {
        validate_template(rule_type, template).map_err(|_| {
            invalid(
                "BrtBeginCFRule14",
                format!("template {template} is invalid for {rule_type:?}"),
            )
        })
    }
}

pub fn validate_formula_count(
    rule_type: RuleType,
    template: u32,
    parameter: u32,
    count: usize,
) -> Result<()> {
    let expected = if rule_type == RuleType::CellIs {
        if matches!(parameter, 1 | 2) { 2 } else { 1 }
    } else if rule_type == RuleType::Expression && matches!(template, 1 | 8..=12 | 15..=24) {
        1
    } else {
        0
    };
    let valid = if matches!(
        rule_type,
        RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
    ) {
        count <= 1
    } else {
        count == expected
    };
    if valid {
        Ok(())
    } else {
        Err(invalid(
            "BrtBeginCFRule",
            format!("formula count {count} does not match required {expected}"),
        ))
    }
}

pub(super) fn validate_formula_slots(
    rule_type: RuleType,
    template: u32,
    parameter: u32,
    slots: &[Option<ParsedFormula>; 3],
) -> Result<()> {
    let expected = if rule_type == RuleType::CellIs {
        [true, matches!(parameter, 1 | 2), false]
    } else if rule_type == RuleType::Expression && matches!(template, 1 | 8..=12 | 15..=24) {
        [true, false, false]
    } else if matches!(
        rule_type,
        RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
    ) {
        [false, false, slots[2].is_some()]
    } else {
        [false, false, false]
    };
    let found = slots.each_ref().map(Option::is_some);
    if found == expected {
        Ok(())
    } else {
        Err(invalid(
            "BrtBeginCFRule",
            format!("formula slots {found:?} do not match required {expected:?}"),
        ))
    }
}

pub(super) fn validate_parameter_and_flags(
    rule_type: RuleType,
    template: u32,
    parameter: u32,
    above_average: bool,
    bottom: bool,
    percent: bool,
) -> Result<()> {
    let valid_parameter = match (rule_type, template) {
        (RuleType::CellIs, 0) => (1..=8).contains(&parameter),
        (RuleType::Expression, 8) => parameter <= 3,
        (RuleType::Expression, 15) => parameter == 0,
        (RuleType::Expression, 16) => parameter == 6,
        (RuleType::Expression, 17) => parameter == 1,
        (RuleType::Expression, 18) => parameter == 2,
        (RuleType::Expression, 19) => parameter == 5,
        (RuleType::Expression, 20) => parameter == 8,
        (RuleType::Expression, 21) => parameter == 3,
        (RuleType::Expression, 22) => parameter == 7,
        (RuleType::Expression, 23) => parameter == 4,
        (RuleType::Expression, 24) => parameter == 9,
        (RuleType::Expression, 25 | 26) => parameter < 4,
        (RuleType::TopN, 5) if percent => parameter <= 100,
        (RuleType::TopN, 5) => (1..=1_000).contains(&parameter),
        _ => parameter == 0,
    };
    if !valid_parameter {
        return Err(invalid(
            "BrtBeginCFRule",
            format!("invalid parameter {parameter} for template {template}"),
        ));
    }
    if above_average != matches!(template, 25 | 29) {
        return Err(invalid(
            "BrtBeginCFRule",
            format!("invalid above-average flag for template {template}"),
        ));
    }
    if rule_type != RuleType::TopN && (bottom || percent) {
        return Err(invalid(
            "BrtBeginCFRule",
            "bottom/percent flags are set on a non-filter rule",
        ));
    }
    Ok(())
}

pub(super) fn render_formula(
    formula: &ParsedFormula,
    base: (u32, u32),
    context: &impl Resolution,
) -> Result<String> {
    let tokens =
        Parser::with_base_cell_and_extra(&formula.rgce, &formula.rgcb, base.0, base.1).parse()?;
    Ok(Compiler::try_tokens_to_string_with_resolution(
        &tokens, context,
    )?)
}

pub(super) fn format_number(value: f64) -> String {
    if value.fract() == 0.0 {
        format!("{value:.0}")
    } else {
        value.to_string()
    }
}

pub(super) fn effective_rule_parameter(rule: &Rule) -> Result<u32> {
    if rule.rule_type != RuleType::CellIs {
        if rule.operator.is_some() {
            return Err(invalid(
                "BrtBeginCFRule14",
                "operator is set on a non-cell-comparison rule",
            ));
        }
        return Ok(rule.parameter);
    }
    let parameter = rule.operator.map_or(rule.parameter, u32::from);
    if rule.parameter != 0 && rule.parameter != parameter {
        return Err(invalid(
            "BrtBeginCFRule14",
            "operator and exact parameter disagree",
        ));
    }
    Ok(parameter)
}

pub(super) fn effective_rule_formulas(rule: &Rule) -> Result<Vec<ParsedFormula>> {
    if !rule.formulas.is_empty() {
        if !rule.formula_extras.is_empty() && rule.formula_extras.len() != rule.formulas.len() {
            return Err(Error::InvalidFormula(
                "conditional-format ancillary stream count does not match formulas".to_string(),
            ));
        }
        return rule
            .formulas
            .iter()
            .enumerate()
            .map(|(index, rgce)| {
                if rgce.is_empty() || rgce.len() > MAX_CELL_FORMULA_BYTES {
                    return Err(Error::InvalidFormula(format!(
                        "conditional-format formula length {} is outside 1..={MAX_CELL_FORMULA_BYTES}",
                        rgce.len()
                    )));
                }
                Ok(ParsedFormula {
                    rgce: rgce.clone(),
                    rgcb: rule.formula_extras.get(index).cloned().unwrap_or_default(),
                })
            })
            .collect();
    }
    rule.formula_texts
        .iter()
        .map(|formula| TextCompiler::compile(formula))
        .collect()
}

pub(super) fn validate_rule_text(
    template: u32,
    text: Option<&str>,
    record: &'static str,
) -> Result<()> {
    if template == 8 {
        if text.is_none_or(|text| text.is_empty() || text.encode_utf16().count() > 255) {
            return Err(invalid(record, "invalid text parameter"));
        }
    } else if text.is_some() {
        return Err(invalid(record, "non-text template has a text parameter"));
    }
    Ok(())
}

pub(super) fn validate_extension_links(cond_fmts: &[Formatting]) -> Result<()> {
    let mut classic = HashMap::new();
    for formatting in cond_fmts {
        if formatting.record_kind != RecordKind::Classic {
            continue;
        }
        for rule in &formatting.rules {
            let Some(guid) = rule.classic_extension_guid else {
                continue;
            };
            let bar = rule
                .data_bar
                .as_ref()
                .ok_or_else(|| invalid("BrtCFRuleExt", "is attached to a non-data-bar rule"))?;
            if classic
                .insert(guid, (rule.priority, bar.min_length, bar.max_length))
                .is_some()
            {
                return Err(invalid("BrtCFRuleExt", "duplicate GUID"));
            }
        }
    }
    let mut matched = HashSet::new();
    for formatting in cond_fmts {
        if formatting.record_kind != RecordKind::Extension14 {
            continue;
        }
        for rule in &formatting.rules {
            let Some(metadata) = rule.extension14 else {
                continue;
            };
            if metadata.priority != -1 || !metadata.guid_present {
                continue;
            }
            let Some(&(priority, classic_min, classic_max)) = classic.get(&metadata.guid) else {
                if metadata.linked_classic_priority.is_some() {
                    return Err(invalid(
                        "BrtBeginCFRule14",
                        "resolved classic priority has no matching GUID",
                    ));
                }
                continue;
            };
            if !matched.insert(metadata.guid) {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "multiple data-bar extensions use the same GUID",
                ));
            }
            if metadata
                .linked_classic_priority
                .is_some_and(|linked| linked != priority)
            {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "resolved classic priority disagrees with its GUID",
                ));
            }
            let bar = rule
                .data_bar14
                .as_ref()
                .ok_or_else(|| invalid("BrtBeginDatabar14", "missing linked data bar"))?;
            let expected_lengths = if bar.min_length == 0 && bar.max_length == 100 {
                (10, 90)
            } else {
                (bar.min_length, bar.max_length)
            };
            if (classic_min, classic_max) != expected_lengths {
                return Err(invalid(
                    "BrtBeginDatabar14",
                    "widths do not agree with the linked classic data bar",
                ));
            }
        }
    }
    if classic.keys().any(|guid| !matched.contains(guid)) {
        return Err(invalid(
            "BrtCFRuleExt",
            "GUID has no matching data-bar extension",
        ));
    }
    Ok(())
}

pub(super) fn validate_rule_metadata(rule: &Rule) -> Result<()> {
    if rule.extension14.is_some()
        || rule.color_scale14.is_some()
        || rule.data_bar14.is_some()
        || rule.icon_set14.is_some()
    {
        return Err(invalid(
            "BrtBeginCFRule",
            "Office 2013 fields are set on a classic rule",
        ));
    }
    validate_template(rule.rule_type, rule.template)?;
    if rule.priority == 0 || rule.priority > i32::MAX as u32 {
        return Err(invalid(
            "BrtBeginCFRule",
            format!("invalid priority {}", rule.priority),
        ));
    }
    if rule.dxf_id.is_some_and(|id| id > i32::MAX as u32) {
        return Err(invalid(
            "BrtBeginCFRule",
            "differential-format index overflow",
        ));
    }
    let visual = matches!(
        rule.rule_type,
        RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
    );
    if visual && (rule.dxf_id.is_some() || rule.stop_if_true) {
        return Err(invalid(
            "BrtBeginCFRule",
            "visual rule has a DXF or stop-if-true flag",
        ));
    }
    let expected_visual = match rule.rule_type {
        RuleType::ColorScale => {
            rule.color_scale.is_some() && rule.data_bar.is_none() && rule.icon_set.is_none()
        },
        RuleType::DataBar => {
            rule.color_scale.is_none() && rule.data_bar.is_some() && rule.icon_set.is_none()
        },
        RuleType::IconSet => {
            rule.color_scale.is_none() && rule.data_bar.is_none() && rule.icon_set.is_some()
        },
        _ => rule.color_scale.is_none() && rule.data_bar.is_none() && rule.icon_set.is_none(),
    };
    if !expected_visual {
        return Err(invalid(
            "BrtBeginCFRule",
            "visualization does not match rule type",
        ));
    }
    if rule.template == 8 {
        let valid = rule
            .text
            .as_deref()
            .is_some_and(|text| !text.is_empty() && text.encode_utf16().count() <= 255);
        if !valid {
            return Err(invalid("BrtBeginCFRule", "invalid text parameter"));
        }
    } else if rule.text.is_some() {
        return Err(invalid(
            "BrtBeginCFRule",
            "non-text template has a text parameter",
        ));
    }
    validate_parameter_and_flags(
        rule.rule_type,
        rule.template,
        effective_rule_parameter(rule)?,
        rule.above_average,
        rule.bottom,
        rule.percent,
    )
}

pub(super) fn validate_scale_thresholds14(scale: &Scale) -> Result<()> {
    if matches!(scale.min_cfvo.cfvo_type, 3 | 8 | 9)
        || matches!(scale.max_cfvo.cfvo_type, 2 | 8 | 9)
    {
        return Err(invalid(
            "BrtBeginColorScale14",
            "minimum/maximum threshold type is reversed",
        ));
    }
    if scale.mid_cfvo.is_some() != scale.mid_color_record.is_some()
        || scale.mid_cfvo.is_some() != scale.mid_color.is_some()
    {
        return Err(invalid(
            "BrtBeginColorScale14",
            "middle threshold and color must both be present or absent",
        ));
    }
    if scale
        .mid_cfvo
        .as_ref()
        .is_some_and(|cfvo| matches!(cfvo.cfvo_type, 2 | 3 | 8 | 9))
    {
        return Err(invalid(
            "BrtBeginColorScale14",
            "middle threshold cannot be a boundary",
        ));
    }
    Ok(())
}

pub(super) fn validate_data_bar14(bar: &Bar14, priority: i32) -> Result<()> {
    if matches!(bar.min_cfvo.cfvo_type, 3 | 9) || matches!(bar.max_cfvo.cfvo_type, 2 | 8) {
        return Err(invalid(
            "BrtBeginDatabar14",
            "minimum/maximum threshold type is reversed",
        ));
    }
    let valid_colors = bar.positive_color.is_some() == (priority != -1)
        && bar.border_color.is_some() == bar.border
        && bar.negative_color.is_some() == bar.custom_negative_fill
        && bar.negative_border_color.is_some() == (bar.custom_negative_border && bar.border)
        && bar.axis_color.is_some() == (bar.axis_position != AxisPosition14::None);
    if !valid_colors {
        return Err(invalid(
            "BrtBeginDatabar14",
            "color records do not match data-bar flags",
        ));
    }
    Ok(())
}

pub(super) fn validate_icon_set14(set: &IconSet14) -> Result<()> {
    let expected = icon_count14(set.icon_set_type);
    if expected == 0 || set.cfvos.len() != expected {
        return Err(invalid(
            "BrtBeginIconSet14",
            format!("expected {expected} thresholds, found {}", set.cfvos.len()),
        ));
    }
    if set
        .cfvos
        .iter()
        .any(|cfvo| matches!(cfvo.cfvo_type, 2 | 3 | 8 | 9) || !cfvo.save_greater_than_or_equal)
    {
        return Err(invalid(
            "BrtBeginIconSet14",
            "invalid threshold type or fSaveGTE flag",
        ));
    }
    if set
        .custom_icons
        .as_ref()
        .is_some_and(|icons| icons.len() != expected)
    {
        return Err(invalid(
            "BrtBeginIconSet14",
            "custom icon count does not match thresholds",
        ));
    }
    Ok(())
}

pub(super) fn validate_scale_thresholds(scale: &Scale) -> Result<()> {
    validate_boundary_thresholds(&scale.min_cfvo, &scale.max_cfvo, "BrtBeginColorScale")?;
    if scale.mid_cfvo.is_some() != scale.mid_color_record.is_some()
        || scale.mid_cfvo.is_some() != scale.mid_color.is_some()
    {
        return Err(invalid(
            "BrtBeginColorScale",
            "middle threshold and color must both be present or absent",
        ));
    }
    if scale
        .mid_cfvo
        .as_ref()
        .is_some_and(|cfvo| matches!(cfvo.cfvo_type, 2 | 3))
    {
        return Err(invalid(
            "BrtBeginColorScale",
            "middle threshold cannot be min/max",
        ));
    }
    Ok(())
}

pub(super) fn validate_boundary_thresholds(
    minimum: &Value,
    maximum: &Value,
    record: &str,
) -> Result<()> {
    if minimum.cfvo_type == 3 || maximum.cfvo_type == 2 {
        return Err(invalid(
            record,
            "minimum/maximum threshold type is reversed",
        ));
    }
    Ok(())
}

pub(super) fn icon_count(icon_set_type: u8) -> Result<usize> {
    match icon_set_type {
        0..=7 => Ok(3),
        8..=12 => Ok(4),
        13..=16 => Ok(5),
        value => Err(invalid("BrtBeginIconSet", format!("invalid set {value}"))),
    }
}

pub(super) fn effective_cfvo_formula(cfvo: &Value) -> Result<Option<ParsedFormula>> {
    if let Some(formula) = &cfvo.formula_binary {
        return Ok(Some(formula.clone()));
    }
    let Some(value) = cfvo.value.as_deref().filter(|value| !value.is_empty()) else {
        return Ok(None);
    };
    if matches!(cfvo.cfvo_type, 1 | 4 | 5) && value.parse::<f64>().is_ok() {
        return Ok(None);
    }
    if cfvo.cfvo_type == 7 || matches!(cfvo.cfvo_type, 1 | 4 | 5) {
        return TextCompiler::compile(value).map(Some);
    }
    Ok(None)
}
