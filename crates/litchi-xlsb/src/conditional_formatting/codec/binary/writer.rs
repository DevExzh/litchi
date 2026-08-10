#![allow(
    clippy::expect_used,
    clippy::map_err_ignore,
    clippy::wildcard_enum_match_arm,
    reason = "legacy module confines extraction after an immediately preceding structural invariant check, normalization into the module's stable typed public error, an intentional opaque or future-variant fallback to this codec boundary"
)]

//! Record-stream writer for classic and Office 2013 conditional formatting.

use crate::conditional_formatting::model::*;
use crate::formula::ParsedFormula;
use crate::raw::{Writer, kind};
use std::collections::HashSet;
use std::io::Write;

use super::super::semantic::{
    effective_cfvo_formula, effective_rule_formulas, effective_rule_parameter, icon_count,
    validate_boundary_thresholds, validate_data_bar14, validate_extension_links,
    validate_formula_count, validate_icon_set14, validate_rule_metadata, validate_scale_thresholds,
    validate_scale_thresholds14,
};
use super::super::{Error, Result, invalid};
use super::wire::{parse_range_list, serialize_rule_extension_guid, write_bin_range_list};

/// Write all classic and Office 2013 conditional-formatting collections for a worksheet.
pub fn write_conditional_formattings<W: Write>(
    writer: &mut Writer<W>,
    cond_fmts: &[Formatting],
) -> Result<()> {
    validate_extension_links(cond_fmts)?;
    let mut priorities = HashSet::new();
    for rule in cond_fmts.iter().flat_map(|formatting| &formatting.rules) {
        let priority = rule
            .extension14
            .map_or(i64::from(rule.priority), |metadata| {
                i64::from(metadata.priority)
            });
        if priority > 0 && !priorities.insert(priority) {
            return Err(invalid(
                "BrtBeginCFRule priority",
                format!("duplicate {priority}"),
            ));
        }
    }
    for formatting in cond_fmts {
        match formatting.record_kind {
            RecordKind::Classic => write_single_cond_formatting(writer, formatting)?,
            RecordKind::Extension14 => write_single_cond_formatting14(writer, formatting)?,
        }
    }
    Ok(())
}

fn write_single_cond_formatting<W: Write>(
    writer: &mut Writer<W>,
    formatting: &Formatting,
) -> Result<()> {
    writer.write_record(
        kind::BEGIN_COND_FORMATTING,
        &serialize_cond_formatting_header(formatting)?,
    )?;
    for rule in &formatting.rules {
        writer.write_record(kind::BEGIN_CF_RULE, &serialize_cf_rule(rule)?)?;
        write_rule_visualization(writer, rule)?;
        if let Some(guid) = rule.classic_extension_guid {
            writer.write_record(kind::CF_RULE_EXT, &serialize_rule_extension_guid(guid))?;
        }
        writer.write_record(kind::END_CF_RULE, &[])?;
    }
    writer.write_record(kind::END_COND_FORMATTING, &[])?;
    Ok(())
}

fn write_single_cond_formatting14<W: Write>(
    writer: &mut Writer<W>,
    formatting: &Formatting,
) -> Result<()> {
    writer.write_record(
        kind::BEGIN_COND_FORMATTING14,
        &formatting.serialize_extension14_header()?,
    )?;
    for rule in &formatting.rules {
        writer.write_record(kind::BEGIN_CF_RULE14, &rule.serialize_extension14()?)?;
        write_rule_visualization14(writer, rule)?;
        writer.write_record(kind::END_CF_RULE14, &[])?;
    }
    writer.write_record(kind::END_COND_FORMATTING14, &[])?;
    Ok(())
}

pub(crate) fn serialize_cond_formatting_header(formatting: &Formatting) -> Result<Vec<u8>> {
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

pub(crate) fn serialize_cf_rule(rule: &Rule) -> Result<Vec<u8>> {
    validate_rule_metadata(rule)?;
    let parameter = effective_rule_parameter(rule)?;
    let formulas = effective_rule_formulas(rule)?;
    validate_formula_count(rule.rule_type, rule.template, parameter, formulas.len())?;

    let mut slots: [Option<&ParsedFormula>; 3] = [None, None, None];
    let start = if matches!(
        rule.rule_type,
        RuleType::ColorScale | RuleType::DataBar | RuleType::IconSet
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
            .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?;
        payload.extend_from_slice(&size.to_le_bytes());
    }
    write_nullable_wide_string(&mut payload, rule.text.as_deref())?;
    for formula in slots.into_iter().flatten() {
        payload.extend_from_slice(&formula.to_bytes()?);
    }
    Ok(payload)
}

fn write_rule_visualization<W: Write>(writer: &mut Writer<W>, rule: &Rule) -> Result<()> {
    match rule.rule_type {
        RuleType::ColorScale => {
            let scale = rule.color_scale.as_ref().expect("validated color scale");
            validate_scale_thresholds(scale)?;
            writer.write_record(kind::BEGIN_COLOR_SCALE, &[])?;
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
            writer.write_record(kind::END_COLOR_SCALE, &[])?;
        },
        RuleType::DataBar => {
            let bar = rule.data_bar.as_ref().expect("validated data bar");
            if bar.min_length > bar.max_length || bar.max_length > 100 {
                return Err(invalid("BrtBeginDatabar", "invalid minimum/maximum length"));
            }
            validate_boundary_thresholds(&bar.min_cfvo, &bar.max_cfvo, "BrtBeginDatabar")?;
            writer.write_record(
                kind::BEGIN_DATABAR,
                &[bar.min_length, bar.max_length, u8::from(bar.show_value)],
            )?;
            write_cfvo(writer, &bar.min_cfvo, false)?;
            write_cfvo(writer, &bar.max_cfvo, false)?;
            write_color(writer, bar.color_record, bar.color)?;
            writer.write_record(kind::END_DATABAR, &[])?;
        },
        RuleType::IconSet => {
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
            writer.write_record(kind::BEGIN_ICON_SET, &begin)?;
            for cfvo in &set.cfvos {
                write_cfvo(writer, cfvo, true)?;
            }
            writer.write_record(kind::END_ICON_SET, &[])?;
        },
        _ => {},
    }
    Ok(())
}

fn write_rule_visualization14<W: Write>(writer: &mut Writer<W>, rule: &Rule) -> Result<()> {
    if rule.color_scale.is_some() || rule.data_bar.is_some() || rule.icon_set.is_some() {
        return Err(invalid(
            "BrtBeginCFRule14",
            "classic visualization is set on an Office 2013 rule",
        ));
    }
    match rule.rule_type {
        RuleType::ColorScale => {
            let scale = rule
                .color_scale14
                .as_ref()
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing Office 2013 color scale"))?;
            if rule.data_bar14.is_some() || rule.icon_set14.is_some() {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "visualization does not match rule type",
                ));
            }
            validate_scale_thresholds14(scale)?;
            writer.write_record(kind::BEGIN_COLOR_SCALE14, &[])?;
            write_cfvo14(writer, &scale.min_cfvo, false)?;
            if let Some(midpoint) = &scale.mid_cfvo {
                write_cfvo14(writer, midpoint, false)?;
            }
            write_cfvo14(writer, &scale.max_cfvo, false)?;
            write_color14(writer, scale.min_color_record, scale.min_color)?;
            if let (Some(record), Some(argb)) = (scale.mid_color_record, scale.mid_color) {
                write_color14(writer, record, argb)?;
            }
            write_color14(writer, scale.max_color_record, scale.max_color)?;
            writer.write_record(kind::END_COLOR_SCALE14, &[])?;
        },
        RuleType::DataBar => {
            let bar = rule
                .data_bar14
                .as_ref()
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing Office 2013 data bar"))?;
            if rule.color_scale14.is_some() || rule.icon_set14.is_some() {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "visualization does not match rule type",
                ));
            }
            let priority = rule
                .extension14
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing extension metadata"))?
                .priority;
            validate_data_bar14(bar, priority)?;
            writer.write_record(kind::BEGIN_DATABAR14, &bar.serialize_header()?)?;
            write_cfvo14(writer, &bar.min_cfvo, false)?;
            write_cfvo14(writer, &bar.max_cfvo, false)?;
            for color in [
                bar.positive_color,
                bar.border_color,
                bar.negative_color,
                bar.negative_border_color,
                bar.axis_color,
            ]
            .into_iter()
            .flatten()
            {
                writer.write_record(kind::COLOR14, &color.serialize_extension14()?)?;
            }
            writer.write_record(kind::END_DATABAR14, &[])?;
        },
        RuleType::IconSet => {
            let set = rule
                .icon_set14
                .as_ref()
                .ok_or_else(|| invalid("BrtBeginCFRule14", "missing Office 2013 icon set"))?;
            if rule.color_scale14.is_some() || rule.data_bar14.is_some() {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "visualization does not match rule type",
                ));
            }
            validate_icon_set14(set)?;
            writer.write_record(kind::BEGIN_ICON_SET14, &set.serialize_header()?)?;
            for cfvo in &set.cfvos {
                write_cfvo14(writer, cfvo, true)?;
            }
            if let Some(icons) = &set.custom_icons {
                for icon in icons {
                    writer.write_record(kind::CF_ICON, &icon.serialize()?)?;
                }
            }
            writer.write_record(kind::END_ICON_SET14, &[])?;
        },
        _ => {
            if rule.color_scale14.is_some()
                || rule.data_bar14.is_some()
                || rule.icon_set14.is_some()
            {
                return Err(invalid(
                    "BrtBeginCFRule14",
                    "non-visual rule contains a visualization",
                ));
            }
        },
    }
    Ok(())
}

fn write_cfvo14<W: Write>(writer: &mut Writer<W>, cfvo: &Value, icon_set: bool) -> Result<()> {
    let formula = effective_cfvo_formula(cfvo)?;
    let numeric_value = if formula.is_none() && matches!(cfvo.cfvo_type, 1 | 4 | 5) {
        cfvo.value
            .as_deref()
            .and_then(|value| value.parse().ok())
            .unwrap_or(cfvo.numeric_value)
    } else {
        cfvo.numeric_value
    };
    writer.write_record(
        kind::CFVO14,
        &cfvo.serialize_extension14_with(
            formula.as_ref(),
            numeric_value,
            icon_set || cfvo.save_greater_than_or_equal,
        )?,
    )?;
    Ok(())
}

fn write_color14<W: Write>(writer: &mut Writer<W>, record: Color, legacy_argb: u32) -> Result<()> {
    let record = if record.argb == Some(legacy_argb) || (record.argb.is_none() && legacy_argb == 0)
    {
        record
    } else {
        Color::from_argb(legacy_argb)
    };
    writer.write_record(kind::COLOR14, &record.serialize_extension14()?)?;
    Ok(())
}

fn write_cfvo<W: Write>(writer: &mut Writer<W>, cfvo: &Value, icon_set: bool) -> Result<()> {
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
            .map_err(|_| Error::InvalidFormula("formula is too large".to_string()))?
            .to_le_bytes(),
    );
    if let Some(formula) = formula {
        payload.extend_from_slice(&formula.to_bytes()?);
    }
    writer.write_record(kind::CFVO, &payload)?;
    Ok(())
}

fn write_color<W: Write>(writer: &mut Writer<W>, record: Color, legacy_argb: u32) -> Result<()> {
    let record = if record.argb == Some(legacy_argb) || (record.argb.is_none() && legacy_argb == 0)
    {
        record
    } else {
        Color::from_argb(legacy_argb)
    };
    writer.write_record(kind::COLOR, &record.to_bytes()?)?;
    Ok(())
}

fn write_nullable_wide_string(payload: &mut Vec<u8>, value: Option<&str>) -> Result<()> {
    let Some(value) = value else {
        payload.extend_from_slice(&u32::MAX.to_le_bytes());
        return Ok(());
    };
    let units = value.encode_utf16().count();
    payload.extend_from_slice(
        &u32::try_from(units)
            .map_err(|_| Error::Encoding("conditional-format text is too long".to_string()))?
            .to_le_bytes(),
    );
    payload.reserve(units.saturating_mul(2));
    for unit in value.encode_utf16() {
        payload.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(())
}
