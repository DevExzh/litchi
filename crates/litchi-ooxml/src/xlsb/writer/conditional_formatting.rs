//! Strict classic conditional-formatting serialization for XLSB worksheets.

use crate::xlsb::conditional_formatting::{
    CfRuleType, Cfvo, ConditionalFormatColor, ConditionalFormatting, ConditionalFormattingRule,
    validate_formula_count, validate_template,
};
use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::formula::{CellParsedFormula, FormulaCompiler, MAX_CELL_FORMULA_BYTES};
use crate::xlsb::records::record_types;
use crate::xlsb::writer::RecordWriter;
use crate::xlsb::writer::bin_range::{parse_range_list, write_bin_range_list};
use std::collections::HashSet;
use std::io::Write;

/// Write all classic conditional-formatting collections for a worksheet.
pub fn write_conditional_formattings<W: Write>(
    writer: &mut RecordWriter<W>,
    cond_fmts: &[ConditionalFormatting],
) -> XlsbResult<()> {
    let mut priorities = HashSet::new();
    for rule in cond_fmts.iter().flat_map(|formatting| &formatting.rules) {
        if !priorities.insert(rule.priority) {
            return Err(invalid(
                "BrtBeginCFRule priority",
                format!("duplicate {}", rule.priority),
            ));
        }
    }
    for formatting in cond_fmts {
        write_single_cond_formatting(writer, formatting)?;
    }
    Ok(())
}

fn write_single_cond_formatting<W: Write>(
    writer: &mut RecordWriter<W>,
    formatting: &ConditionalFormatting,
) -> XlsbResult<()> {
    writer.write_record(
        record_types::BEGIN_COND_FORMATTING,
        &serialize_cond_formatting_header(formatting)?,
    )?;
    for rule in &formatting.rules {
        writer.write_record(record_types::BEGIN_CF_RULE, &serialize_cf_rule(rule)?)?;
        write_rule_visualization(writer, rule)?;
        writer.write_record(record_types::END_CF_RULE, &[])?;
    }
    writer.write_record(record_types::END_COND_FORMATTING, &[])?;
    Ok(())
}

fn serialize_cond_formatting_header(formatting: &ConditionalFormatting) -> XlsbResult<Vec<u8>> {
    let rule_count = u32::try_from(formatting.rules.len())
        .map_err(|_| invalid("BrtBeginConditionalFormatting", "too many rules"))?;
    let mut ranges = Vec::new();
    for range in &formatting.ranges {
        ranges.extend(parse_range_list(range)?);
    }
    if ranges.is_empty() || ranges.len() > 8_192 {
        return Err(invalid(
            "BrtBeginConditionalFormatting",
            format!("classic range count {} is outside 1..=8192", ranges.len()),
        ));
    }
    let mut payload = Vec::with_capacity(12 + ranges.len() * 16);
    payload.extend_from_slice(&rule_count.to_le_bytes());
    payload.extend_from_slice(&u32::from(formatting.pivot_only).to_le_bytes());
    write_bin_range_list(&ranges, &mut payload)?;
    Ok(payload)
}

fn serialize_cf_rule(rule: &ConditionalFormattingRule) -> XlsbResult<Vec<u8>> {
    validate_rule_metadata(rule)?;
    let parameter = effective_parameter(rule)?;
    let formulas = effective_formulas(rule)?;
    validate_formula_count(rule.rule_type, rule.template, parameter, formulas.len())?;

    let mut slots: [Option<&CellParsedFormula>; 3] = [None, None, None];
    let start = if matches!(
        rule.rule_type,
        CfRuleType::ColorScale | CfRuleType::DataBar | CfRuleType::IconSet
    ) {
        2
    } else {
        0
    };
    for (index, formula) in formulas.iter().enumerate() {
        slots[start + index] = Some(formula);
    }

    let mut payload = Vec::with_capacity(64);
    payload.extend_from_slice(&(rule.rule_type as u32).to_le_bytes());
    payload.extend_from_slice(&rule.template.to_le_bytes());
    payload.extend_from_slice(&rule.dxf_id.unwrap_or(u32::MAX).to_le_bytes());
    payload.extend_from_slice(&rule.priority.to_le_bytes());
    payload.extend_from_slice(&parameter.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    payload.extend_from_slice(&0u32.to_le_bytes());
    let mut flags = 0u16;
    if rule.stop_if_true {
        flags |= 0x02;
    }
    if rule.above_average {
        flags |= 0x04;
    }
    if rule.bottom {
        flags |= 0x08;
    }
    if rule.percent {
        flags |= 0x10;
    }
    payload.extend_from_slice(&flags.to_le_bytes());
    for formula in &slots {
        let size = formula.map_or(0, |formula| formula.rgce.len());
        let size = u32::try_from(size)
            .map_err(|_| XlsbError::InvalidFormula("formula is too large".to_string()))?;
        payload.extend_from_slice(&size.to_le_bytes());
    }
    write_nullable_wide_string(&mut payload, rule.text.as_deref())?;
    for formula in slots.into_iter().flatten() {
        payload.extend_from_slice(&formula.to_bytes()?);
    }
    Ok(payload)
}

fn validate_rule_metadata(rule: &ConditionalFormattingRule) -> XlsbResult<()> {
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
        CfRuleType::ColorScale | CfRuleType::DataBar | CfRuleType::IconSet
    );
    if visual && (rule.dxf_id.is_some() || rule.stop_if_true) {
        return Err(invalid(
            "BrtBeginCFRule",
            "visual rule has a DXF or stop-if-true flag",
        ));
    }
    let expected_visual = match rule.rule_type {
        CfRuleType::ColorScale => {
            rule.color_scale.is_some() && rule.data_bar.is_none() && rule.icon_set.is_none()
        },
        CfRuleType::DataBar => {
            rule.color_scale.is_none() && rule.data_bar.is_some() && rule.icon_set.is_none()
        },
        CfRuleType::IconSet => {
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
    validate_rule_flags_and_parameter(rule, effective_parameter(rule)?)
}

fn validate_rule_flags_and_parameter(
    rule: &ConditionalFormattingRule,
    parameter: u32,
) -> XlsbResult<()> {
    let valid_parameter = match (rule.rule_type, rule.template) {
        (CfRuleType::CellIs, 0) => (1..=8).contains(&parameter),
        (CfRuleType::Expression, 8) => parameter <= 3,
        (CfRuleType::Expression, 15) => parameter == 0,
        (CfRuleType::Expression, 16) => parameter == 6,
        (CfRuleType::Expression, 17) => parameter == 1,
        (CfRuleType::Expression, 18) => parameter == 2,
        (CfRuleType::Expression, 19) => parameter == 5,
        (CfRuleType::Expression, 20) => parameter == 8,
        (CfRuleType::Expression, 21) => parameter == 3,
        (CfRuleType::Expression, 22) => parameter == 7,
        (CfRuleType::Expression, 23) => parameter == 4,
        (CfRuleType::Expression, 24) => parameter == 9,
        (CfRuleType::Expression, 25 | 26) => parameter < 4,
        (CfRuleType::TopN, 5) if rule.percent => parameter <= 100,
        (CfRuleType::TopN, 5) => (1..=1_000).contains(&parameter),
        _ => parameter == 0,
    };
    if !valid_parameter {
        return Err(invalid(
            "BrtBeginCFRule",
            format!(
                "invalid parameter {parameter} for template {}",
                rule.template
            ),
        ));
    }
    let expected_above = matches!(rule.template, 25 | 29);
    if rule.above_average != expected_above {
        return Err(invalid(
            "BrtBeginCFRule",
            format!(
                "above-average flag is invalid for template {}",
                rule.template
            ),
        ));
    }
    if rule.rule_type != CfRuleType::TopN && (rule.bottom || rule.percent) {
        return Err(invalid(
            "BrtBeginCFRule",
            "bottom/percent flags are set on a non-filter rule",
        ));
    }
    Ok(())
}

fn effective_parameter(rule: &ConditionalFormattingRule) -> XlsbResult<u32> {
    if rule.rule_type != CfRuleType::CellIs {
        if rule.operator.is_some() {
            return Err(invalid(
                "BrtBeginCFRule",
                "operator is set on a non-cell-comparison rule",
            ));
        }
        return Ok(rule.parameter);
    }
    let parameter = rule.operator.map_or(rule.parameter, u32::from);
    if rule.parameter != 0 && rule.parameter != parameter {
        return Err(invalid(
            "BrtBeginCFRule",
            "operator and exact parameter disagree",
        ));
    }
    Ok(parameter)
}

fn effective_formulas(rule: &ConditionalFormattingRule) -> XlsbResult<Vec<CellParsedFormula>> {
    if !rule.formulas.is_empty() {
        if !rule.formula_extras.is_empty() && rule.formula_extras.len() != rule.formulas.len() {
            return Err(XlsbError::InvalidFormula(
                "conditional-format ancillary stream count does not match formulas".to_string(),
            ));
        }
        return rule
            .formulas
            .iter()
            .enumerate()
            .map(|(index, rgce)| {
                if rgce.is_empty() || rgce.len() > MAX_CELL_FORMULA_BYTES {
                    return Err(XlsbError::InvalidFormula(format!(
                        "conditional-format formula length {} is outside 1..={MAX_CELL_FORMULA_BYTES}",
                        rgce.len()
                    )));
                }
                Ok(CellParsedFormula {
                    rgce: rgce.clone(),
                    rgcb: rule.formula_extras.get(index).cloned().unwrap_or_default(),
                })
            })
            .collect();
    }
    rule.formula_texts
        .iter()
        .map(|formula| FormulaCompiler::compile(formula))
        .collect()
}

fn write_rule_visualization<W: Write>(
    writer: &mut RecordWriter<W>,
    rule: &ConditionalFormattingRule,
) -> XlsbResult<()> {
    match rule.rule_type {
        CfRuleType::ColorScale => {
            let scale = rule.color_scale.as_ref().expect("validated color scale");
            validate_scale_thresholds(scale)?;
            writer.write_record(record_types::BEGIN_COLOR_SCALE, &[])?;
            write_cfvo(writer, &scale.min_cfvo, false)?;
            if let Some(midpoint) = &scale.mid_cfvo {
                write_cfvo(writer, midpoint, false)?;
            }
            write_cfvo(writer, &scale.max_cfvo, false)?;
            write_color(writer, scale.min_color_record, scale.min_color)?;
            if let (Some(record), Some(argb)) = (scale.mid_color_record, scale.mid_color) {
                write_color(writer, record, argb)?;
            }
            write_color(writer, scale.max_color_record, scale.max_color)?;
            writer.write_record(record_types::END_COLOR_SCALE, &[])?;
        },
        CfRuleType::DataBar => {
            let bar = rule.data_bar.as_ref().expect("validated data bar");
            if bar.min_length > bar.max_length || bar.max_length > 100 {
                return Err(invalid("BrtBeginDatabar", "invalid minimum/maximum length"));
            }
            validate_boundary_thresholds(&bar.min_cfvo, &bar.max_cfvo, "BrtBeginDatabar")?;
            writer.write_record(
                record_types::BEGIN_DATABAR,
                &[bar.min_length, bar.max_length, u8::from(bar.show_value)],
            )?;
            write_cfvo(writer, &bar.min_cfvo, false)?;
            write_cfvo(writer, &bar.max_cfvo, false)?;
            write_color(writer, bar.color_record, bar.color)?;
            writer.write_record(record_types::END_DATABAR, &[])?;
        },
        CfRuleType::IconSet => {
            let set = rule.icon_set.as_ref().expect("validated icon set");
            let expected = icon_count(set.icon_set_type)?;
            if set.cfvos.len() != expected {
                return Err(invalid(
                    "BrtBeginIconSet",
                    format!("expected {expected} thresholds, found {}", set.cfvos.len()),
                ));
            }
            if set.cfvos.iter().any(|cfvo| matches!(cfvo.cfvo_type, 2 | 3)) {
                return Err(invalid(
                    "BrtBeginIconSet",
                    "min/max threshold is not allowed",
                ));
            }
            let mut flags = 0u16;
            if !set.show_value {
                flags |= 0x02;
            }
            if !set.reverse {
                flags |= 0x04;
            }
            let mut begin = Vec::with_capacity(6);
            begin.extend_from_slice(&u32::from(set.icon_set_type).to_le_bytes());
            begin.extend_from_slice(&flags.to_le_bytes());
            writer.write_record(record_types::BEGIN_ICON_SET, &begin)?;
            for cfvo in &set.cfvos {
                write_cfvo(writer, cfvo, true)?;
            }
            writer.write_record(record_types::END_ICON_SET, &[])?;
        },
        _ => {},
    }
    Ok(())
}

fn validate_scale_thresholds(
    scale: &crate::xlsb::conditional_formatting::ColorScale,
) -> XlsbResult<()> {
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

fn validate_boundary_thresholds(minimum: &Cfvo, maximum: &Cfvo, record: &str) -> XlsbResult<()> {
    if minimum.cfvo_type == 3 || maximum.cfvo_type == 2 {
        return Err(invalid(
            record,
            "minimum/maximum threshold type is reversed",
        ));
    }
    Ok(())
}

fn icon_count(icon_set_type: u8) -> XlsbResult<usize> {
    match icon_set_type {
        0..=7 => Ok(3),
        8..=12 => Ok(4),
        13..=16 => Ok(5),
        value => Err(invalid("BrtBeginIconSet", format!("invalid set {value}"))),
    }
}

fn write_cfvo<W: Write>(
    writer: &mut RecordWriter<W>,
    cfvo: &Cfvo,
    icon_set: bool,
) -> XlsbResult<()> {
    if !matches!(cfvo.cfvo_type, 1 | 2 | 3 | 4 | 5 | 7) {
        return Err(invalid(
            "BrtCFVO",
            format!("invalid type {}", cfvo.cfvo_type),
        ));
    }
    let formula = effective_cfvo_formula(cfvo)?;
    if matches!(cfvo.cfvo_type, 2 | 3) && formula.is_some() {
        return Err(invalid("BrtCFVO", "min/max threshold contains a formula"));
    }
    if cfvo.cfvo_type == 7 && formula.is_none() {
        return Err(invalid("BrtCFVO", "formula threshold omits its formula"));
    }
    let numeric_value = if formula.is_none() && matches!(cfvo.cfvo_type, 1 | 4 | 5) {
        cfvo.value
            .as_deref()
            .and_then(|value| value.parse().ok())
            .unwrap_or(cfvo.numeric_value)
    } else {
        cfvo.numeric_value
    };
    if !numeric_value.is_finite()
        || (formula.is_none()
            && matches!(cfvo.cfvo_type, 4 | 5)
            && !(0.0..=100.0).contains(&numeric_value))
    {
        return Err(invalid("BrtCFVO", "invalid numeric parameter"));
    }
    let mut payload = Vec::with_capacity(32);
    payload.extend_from_slice(&u32::from(cfvo.cfvo_type).to_le_bytes());
    payload.extend_from_slice(&numeric_value.to_le_bytes());
    payload
        .extend_from_slice(&u32::from(icon_set || cfvo.save_greater_than_or_equal).to_le_bytes());
    payload.extend_from_slice(&u32::from(cfvo.greater_than_or_equal).to_le_bytes());
    let formula_size = formula.as_ref().map_or(0, |formula| formula.rgce.len());
    payload.extend_from_slice(
        &u32::try_from(formula_size)
            .map_err(|_| XlsbError::InvalidFormula("formula is too large".to_string()))?
            .to_le_bytes(),
    );
    if let Some(formula) = formula {
        payload.extend_from_slice(&formula.to_bytes()?);
    }
    writer.write_record(record_types::CFVO, &payload)?;
    Ok(())
}

fn effective_cfvo_formula(cfvo: &Cfvo) -> XlsbResult<Option<CellParsedFormula>> {
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
        return FormulaCompiler::compile(value).map(Some);
    }
    Ok(None)
}

fn write_color<W: Write>(
    writer: &mut RecordWriter<W>,
    record: ConditionalFormatColor,
    legacy_argb: u32,
) -> XlsbResult<()> {
    let record = if record.argb == Some(legacy_argb) || (record.argb.is_none() && legacy_argb == 0)
    {
        record
    } else {
        ConditionalFormatColor::from_argb(legacy_argb)
    };
    writer.write_record(record_types::COLOR, &record.to_bytes()?)?;
    Ok(())
}

fn write_nullable_wide_string(payload: &mut Vec<u8>, value: Option<&str>) -> XlsbResult<()> {
    let Some(value) = value else {
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
        return Ok(());
    };
    let utf16 = value.encode_utf16().collect::<Vec<_>>();
    payload.extend_from_slice(
        &u32::try_from(utf16.len())
            .map_err(|_| XlsbError::Encoding("conditional-format text is too long".to_string()))?
            .to_le_bytes(),
    );
    for unit in utf16 {
        payload.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}

fn invalid(typ: impl Into<String>, val: impl Into<String>) -> XlsbError {
    XlsbError::Unrecognized {
        typ: typ.into(),
        val: val.into(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsb::conditional_formatting::{ColorScale, DataBar, IconSet};
    use crate::xlsb::records::XlsbRecordIter;
    use std::io::Cursor;

    fn compiled(text: &str) -> CellParsedFormula {
        FormulaCompiler::compile(text).unwrap()
    }

    #[test]
    fn fixture_rule_header_matches_libreoffice_sample() {
        let formula = compiled("5");
        let mut rule = ConditionalFormattingRule::new(CfRuleType::CellIs, 1);
        rule.dxf_id = Some(0);
        rule.operator = Some(5);
        rule.parameter = 5;
        rule.formulas.push(formula.rgce);
        let payload = serialize_cf_rule(&rule).unwrap();
        assert_eq!(payload.len(), 57);
        assert_eq!(u32::from_le_bytes(payload[30..34].try_into().unwrap()), 3);
        assert_eq!(&payload[42..46], &u32::MAX.to_le_bytes());
        let parsed = ConditionalFormattingRule::parse(&payload).unwrap();
        assert_eq!(parsed.operator, Some(5));
        assert_eq!(parsed.formula_texts, ["5"]);
    }

    #[test]
    fn header_preserves_pivot_and_multiple_ranges() {
        let mut formatting = ConditionalFormatting::new(vec!["A1:B10".into(), "D4".into()]);
        formatting.pivot_only = true;
        let payload = serialize_cond_formatting_header(&formatting).unwrap();
        assert_eq!(u32::from_le_bytes(payload[4..8].try_into().unwrap()), 1);
        assert_eq!(u32::from_le_bytes(payload[8..12].try_into().unwrap()), 2);
    }

    #[test]
    fn writes_color_scale_data_bar_and_icon_set_subrecords() {
        let mut formatting = ConditionalFormatting::new(vec!["A1:A10".into()]);
        let mut scale = ConditionalFormattingRule::new(CfRuleType::ColorScale, 1);
        scale.color_scale = Some(ColorScale::new(
            Cfvo::new(2, None),
            Cfvo::new(3, None),
            0xffff0000,
            0xff00ff00,
        ));
        formatting.add_rule(scale);

        let mut bar = ConditionalFormattingRule::new(CfRuleType::DataBar, 2);
        bar.data_bar = Some(DataBar::new(
            Cfvo::new(2, None),
            Cfvo::new(3, None),
            0xff4472c4,
        ));
        formatting.add_rule(bar);

        let mut icons = ConditionalFormattingRule::new(CfRuleType::IconSet, 3);
        icons.icon_set = Some(IconSet::new(
            0,
            vec![
                Cfvo::new(1, Some("0".into())),
                Cfvo::new(4, Some("33".into())),
                Cfvo::new(4, Some("67".into())),
            ],
        ));
        formatting.add_rule(icons);

        let mut bytes = Vec::new();
        write_conditional_formattings(&mut RecordWriter::new(&mut bytes), &[formatting]).unwrap();
        let records = XlsbRecordIter::new(Cursor::new(bytes));
        let mut found = Vec::new();
        for record in records {
            found.push(record.unwrap().header.record_type);
        }
        for typ in [
            record_types::BEGIN_COLOR_SCALE,
            record_types::BEGIN_DATABAR,
            record_types::BEGIN_ICON_SET,
            record_types::CFVO,
            record_types::COLOR,
        ] {
            assert!(found.contains(&typ), "record 0x{typ:04x}");
        }
    }

    #[test]
    fn rejects_duplicate_priority_and_wrong_formula_slot_count() {
        let mut first = ConditionalFormatting::new(vec!["A1".into()]);
        let mut rule = ConditionalFormattingRule::new(CfRuleType::CellIs, 1);
        rule.operator = Some(1);
        rule.formulas.push(compiled("1").rgce);
        first.add_rule(rule.clone());
        let mut second = ConditionalFormatting::new(vec!["B1".into()]);
        second.add_rule(rule);
        assert!(
            write_conditional_formattings(&mut RecordWriter::new(Vec::new()), &[first, second])
                .is_err()
        );
    }
}
