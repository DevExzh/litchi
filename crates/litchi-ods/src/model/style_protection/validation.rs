//! Safety and semantic invariants for table-cell protection styles.

use super::model::{
    ConditionalStyle, MAX_CONDITIONAL_ATTRIBUTE_BYTES, MAX_CONDITIONAL_RULES,
    MAX_CONDITIONAL_STYLES, MAX_CONDITIONAL_TEXT_BYTES, MAX_RULES_PER_STYLE, Rule, TableStyle,
};
use litchi_core::{Error, Result};
use std::collections::HashSet;

pub(crate) fn validate_style_name(name: &str, label: &str) -> Result<()> {
    if name.is_empty() {
        return Err(Error::InvalidFormat(format!("{label} must not be empty")));
    }
    check_conditional_attribute_size(label, name)
}

pub(crate) fn validate_conditional_style_collection(
    styles: &[ConditionalStyle],
    common_styles: &HashSet<String>,
) -> Result<()> {
    if styles.len() > MAX_CONDITIONAL_STYLES {
        return Err(Error::InvalidFormat(format!(
            "document exceeds the {MAX_CONDITIONAL_STYLES} conditional style limit"
        )));
    }
    let mut names = HashSet::with_capacity(styles.len());
    let mut total_rules = 0usize;
    let mut total_text = 0usize;
    for style in styles {
        validate_style_name(&style.style_name, "conditional style name")?;
        if !names.insert(style.style_name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate conditional style name '{}'",
                style.style_name
            )));
        }
        if style.rules.is_empty() {
            return Err(Error::InvalidFormat(format!(
                "conditional style '{}' must contain at least one rule",
                style.style_name
            )));
        }
        if style.rules.len() > MAX_RULES_PER_STYLE {
            return Err(Error::InvalidFormat(format!(
                "conditional style '{}' exceeds the {MAX_RULES_PER_STYLE} rule limit",
                style.style_name
            )));
        }
        if let Some(parent) = &style.parent_style_name {
            validate_style_name(parent, "parent style name")?;
        }
        total_rules = total_rules
            .checked_add(style.rules.len())
            .ok_or_else(|| Error::InvalidFormat("conditional rule count overflow".to_string()))?;
        if total_rules > MAX_CONDITIONAL_RULES {
            return Err(Error::InvalidFormat(format!(
                "document exceeds the {MAX_CONDITIONAL_RULES} conditional rule limit"
            )));
        }
        for rule in &style.rules {
            check_conditional_attribute_size("style:condition", &rule.condition)?;
            if rule.condition.is_empty() {
                return Err(Error::InvalidFormat(
                    "style:condition must not be empty".to_string(),
                ));
            }
            validate_style_name(&rule.apply_style_name, "style:apply-style-name")?;
            if !common_styles.contains(&rule.apply_style_name) {
                return Err(Error::InvalidFormat(format!(
                    "conditional style '{}' references missing, automatic, or non-table-cell common style '{}'",
                    style.style_name, rule.apply_style_name
                )));
            }
            if let Some(base) = &rule.base_cell_address {
                check_conditional_attribute_size("style:base-cell-address", base)?;
                if base.is_empty() {
                    return Err(Error::InvalidFormat(
                        "style:base-cell-address must not be empty".to_string(),
                    ));
                }
            }
            validate_formula_namespace(rule)?;
            total_text = total_text
                .checked_add(rule.condition.len())
                .and_then(|value| value.checked_add(rule.apply_style_name.len()))
                .and_then(|value| {
                    value.checked_add(rule.base_cell_address.as_deref().map_or(0, str::len))
                })
                .ok_or_else(|| {
                    Error::InvalidFormat("conditional style text size overflow".to_string())
                })?;
            if total_text > MAX_CONDITIONAL_TEXT_BYTES {
                return Err(Error::InvalidFormat(format!(
                    "conditional style text exceeds the {MAX_CONDITIONAL_TEXT_BYTES} byte limit"
                )));
            }
        }
    }
    Ok(())
}

fn validate_formula_namespace(rule: &Rule) -> Result<()> {
    let lexical_prefix = formula_prefix(&rule.condition);
    match (lexical_prefix, &rule.formula_namespace) {
        (None, None) => Ok(()),
        (Some(prefix), Some(namespace)) if prefix == namespace.prefix => {
            validate_xml_prefix(&namespace.prefix)?;
            if namespace.uri.is_empty() {
                return Err(Error::InvalidFormat(
                    "conditional formula namespace URI must not be empty".to_string(),
                ));
            }
            check_conditional_attribute_size("formula namespace URI", &namespace.uri)
        },
        (Some(prefix), Some(namespace)) => Err(Error::InvalidFormat(format!(
            "condition prefix '{prefix}' does not match formula namespace prefix '{}'",
            namespace.prefix
        ))),
        (Some(prefix), None) => Err(Error::InvalidFormat(format!(
            "conditional style condition uses unbound namespace prefix '{prefix}'"
        ))),
        (None, Some(namespace)) => Err(Error::InvalidFormat(format!(
            "formula namespace prefix '{}' is not used by the condition",
            namespace.prefix
        ))),
    }
}

fn formula_prefix(condition: &str) -> Option<&str> {
    let (prefix, _) = condition.split_once(':')?;
    let mut characters = prefix.chars();
    if characters
        .next()
        .is_some_and(|character| character == '_' || character.is_alphabetic())
        && characters.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        })
    {
        Some(prefix)
    } else {
        None
    }
}

fn validate_xml_prefix(prefix: &str) -> Result<()> {
    if formula_prefix(&format!("{prefix}:x")) != Some(prefix) || matches!(prefix, "xml" | "xmlns") {
        return Err(Error::InvalidFormat(format!(
            "invalid conditional formula namespace prefix '{prefix}'"
        )));
    }
    Ok(())
}

pub(crate) fn validate_protection_style_collection(styles: &[TableStyle]) -> Result<()> {
    if styles.len() > MAX_CONDITIONAL_STYLES {
        return Err(Error::InvalidFormat(format!(
            "document exceeds the {MAX_CONDITIONAL_STYLES} automatic protection style limit"
        )));
    }
    let mut names = HashSet::with_capacity(styles.len());
    for style in styles {
        validate_style_name(&style.style_name, "protection style name")?;
        if !names.insert(style.style_name.as_str()) {
            return Err(Error::InvalidFormat(format!(
                "duplicate protection style name '{}'",
                style.style_name
            )));
        }
        if let Some(parent) = &style.parent_style_name {
            validate_style_name(parent, "parent style name")?;
        }
    }
    Ok(())
}

pub(super) fn check_conditional_attribute_size(name: &str, value: &str) -> Result<()> {
    if value.len() > MAX_CONDITIONAL_ATTRIBUTE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{name} exceeds the {MAX_CONDITIONAL_ATTRIBUTE_BYTES} byte limit"
        )));
    }
    Ok(())
}
