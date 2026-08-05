//! Shared bounded XML, lexical, and attribute helpers.

#![allow(
    clippy::wildcard_imports,
    reason = "semantic field owners share the stable model facade namespace"
)]
use super::*;

pub(crate) fn validate_dynamic_value(
    name: &str,
    value: Option<&str>,
    required_non_empty: bool,
    aggregate: &mut usize,
) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    if required_non_empty && value.trim().is_empty() {
        return Err(Error::InvalidFormat(format!("{name} must not be empty")));
    }
    if value.len() > MAX_DYNAMIC_FIELD_VALUE {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds {MAX_DYNAMIC_FIELD_VALUE} bytes"
        )));
    }
    if !value.chars().all(is_xml_1_0_char) {
        return Err(Error::InvalidFormat(format!(
            "{name} contains a character forbidden by XML 1.0"
        )));
    }
    *aggregate = aggregate
        .checked_add(value.len())
        .ok_or_else(|| Error::InvalidFormat("dynamic field size overflow".to_string()))?;
    if *aggregate > MAX_DYNAMIC_FIELD_AGGREGATE {
        return Err(Error::InvalidFormat(format!(
            "dynamic field exceeds {MAX_DYNAMIC_FIELD_AGGREGATE} aggregate bytes"
        )));
    }
    Ok(())
}

pub(crate) const fn is_xml_1_0_char(value: char) -> bool {
    matches!(value, '\u{9}' | '\u{A}' | '\u{D}')
        || (value as u32 >= 0x20 && value as u32 <= 0xD7FF)
        || (value as u32 >= 0xE000 && value as u32 <= 0xFFFD)
        || (value as u32 >= 0x10000 && value as u32 <= 0x10FFFF)
}

pub(crate) fn set_data_style(element: &mut Element, value: Option<&str>) {
    if let Some(value) = value {
        element.set_attribute("xmlns:style", STYLE_NAMESPACE);
        element.set_attribute("style:data-style-name", value);
    }
}

pub(crate) fn validate_double(value: &str) -> Result<()> {
    if matches!(value, "INF" | "-INF" | "NaN") || value.parse::<f64>().is_ok() {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "invalid XML Schema double '{value}'"
        )))
    }
}

pub(crate) fn push_xml_attribute(output: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '"' => output.push_str("&quot;"),
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(ch),
        }
    }
}

pub(crate) fn push_xml_text(output: &mut String, value: &str) {
    for ch in value.chars() {
        match ch {
            '&' => output.push_str("&amp;"),
            '<' => output.push_str("&lt;"),
            '>' => output.push_str("&gt;"),
            _ => output.push(ch),
        }
    }
}

pub(crate) fn parse_drop_down_boolean(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "invalid text:current-selected boolean '{value}'"
        ))),
    }
}
