//! Shared scalar and structural validation validation concerns for the `ChartEx` graph.

use super::{
    CX, DoubleOrAutomatic, MAX_SUBTOTALS, MiniNode, QuartileMethod, Result, invalid, invalid_error,
    limit, optional, parse_u32, reject_unknown, required, valid_xml_double,
};
use std::collections::HashSet;

pub(super) fn parse_statistics(node: &MiniNode) -> Result<Option<QuartileMethod>> {
    reject_unknown(&node.attributes, &[("", "quartileMethod")], "statistics")?;
    require_empty_content(node, "statistics")?;
    optional(&node.attributes, "", "quartileMethod")
        .map(|value| match value {
            "inclusive" => Ok(QuartileMethod::Inclusive),
            "exclusive" => Ok(QuartileMethod::Exclusive),
            _ => invalid("invalid  statistics quartileMethod"),
        })
        .transpose()
}

pub(super) fn parse_subtotals(node: &MiniNode) -> Result<Vec<u32>> {
    reject_unknown(&node.attributes, &[], "subtotals")?;
    if !node.text.trim().is_empty() {
        return invalid("unexpected text in  subtotals");
    }
    let mut values = Vec::new();
    let mut unique = HashSet::new();
    for child in &node.children {
        if values.len() >= MAX_SUBTOTALS {
            return limit(" subtotal count");
        }
        if child.namespace != CX
            || child.name != "idx"
            || !child.attributes.is_empty()
            || !child.children.is_empty()
        {
            return invalid("invalid  subtotal index");
        }
        let value = parse_u32(child.text.trim(), "subtotal index")?;
        if !unique.insert(value) {
            return invalid("duplicate  subtotal index");
        }
        values.push(value);
    }
    Ok(values)
}

pub(super) fn parse_double_or_auto(value: &str, label: &str) -> Result<DoubleOrAutomatic> {
    if value == "auto" {
        return Ok(DoubleOrAutomatic::Automatic);
    }
    if !valid_xml_double(value) {
        return invalid(format!("invalid  {label}"));
    }
    Ok(DoubleOrAutomatic::Number(value.to_owned()))
}

pub(super) fn parse_nonnegative_or_auto(value: &str, label: &str) -> Result<DoubleOrAutomatic> {
    if value == "auto" {
        return Ok(DoubleOrAutomatic::Automatic);
    }
    let number = value
        .parse::<f64>()
        .map_err(|_err| invalid_error(format!("invalid  {label}")))?;
    if number.is_nan() || number < 0.0 {
        return invalid(format!("invalid  {label}"));
    }
    Ok(DoubleOrAutomatic::Number(value.to_owned()))
}

pub(super) fn parse_positive_or_auto(value: &str, label: &str) -> Result<DoubleOrAutomatic> {
    if value == "auto" {
        return Ok(DoubleOrAutomatic::Automatic);
    }
    let number = value
        .parse::<f64>()
        .map_err(|_err| invalid_error(format!("invalid  {label}")))?;
    if number.is_nan() || number <= 0.0 {
        return invalid(format!("invalid  {label}"));
    }
    Ok(DoubleOrAutomatic::Number(value.to_owned()))
}

pub(super) fn require_empty_element(node: &MiniNode, label: &str) -> Result<()> {
    reject_unknown(&node.attributes, &[], label)?;
    require_empty_content(node, label)
}

pub(super) fn require_empty_content(node: &MiniNode, label: &str) -> Result<()> {
    if !node.children.is_empty() || !node.text.trim().is_empty() {
        invalid(format!(" {label} must be empty"))
    } else {
        Ok(())
    }
}

pub(super) fn bounded_required(node: &MiniNode, name: &str, max: usize) -> Result<String> {
    let value = required(&node.attributes, "", name)?;
    if value.len() > max {
        return limit(" attribute string");
    }
    Ok(value.to_owned())
}
