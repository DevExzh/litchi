//! Semantic and resource validation for data-validation models.

use super::super::model::{
    Collection, Formula, ListSource, Source, Validation, ValidationOperator, ValidationType,
};
use super::super::{MAX_FORMULA_BYTES, MAX_VALIDATIONS};
use super::wire::{invalid, parse_sqref, sqref_text};
use crate::error::Result;
use litchi_ooxml_common::custom_xml::valid_guid;
use std::collections::HashSet;

pub(super) fn validate_formula_cardinality(
    kind: ValidationType,
    operator: ValidationOperator,
    f1: &Option<ListSource>,
    f2: &Option<Formula>,
) -> Result<()> {
    match kind {
        ValidationType::None => {
            if f1.is_some() || f2.is_some() {
                return Err(invalid("type none must not contain formulas"));
            }
        },
        ValidationType::List | ValidationType::Custom => {
            if f1.is_none() || f2.is_some() {
                return Err(invalid(
                    "list/custom validation requires exactly formula1 or a quoted list",
                ));
            }
        },
        _ if matches!(
            operator,
            ValidationOperator::Between | ValidationOperator::NotBetween
        ) =>
        {
            if f1.is_none() || f2.is_none() {
                return Err(invalid("between validation requires formula1 and formula2"));
            }
        },
        ValidationType::Whole
        | ValidationType::Decimal
        | ValidationType::Date
        | ValidationType::Time
        | ValidationType::TextLength => {
            if f1.is_none() || f2.is_some() {
                return Err(invalid("validation requires formula1 and forbids formula2"));
            }
        },
    }
    Ok(())
}

pub fn validate_data_validation_collections(values: &[Collection]) -> Result<()> {
    let mut sources = HashSet::new();
    let mut count = 0usize;
    for collection in values {
        if !sources.insert(collection.source) {
            return Err(invalid("duplicate dataValidations collection source"));
        }
        validate_collection(collection)?;
        count = count
            .checked_add(collection.validations.len())
            .ok_or_else(|| invalid("data-validation count overflow"))?;
        if count > MAX_VALIDATIONS {
            return Err(invalid("too many data validations"));
        }
    }
    Ok(())
}

pub(crate) fn validate_collection(value: &Collection) -> Result<()> {
    if value.validations.is_empty() || value.validations.len() > MAX_VALIDATIONS {
        return Err(invalid("dataValidations has an invalid rule count"));
    }
    if value.x_window.is_some_and(|v| v > 65_535) || value.y_window.is_some_and(|v| v > 65_535) {
        return Err(invalid("dataValidations window coordinate exceeds 65535"));
    }
    for rule in &value.validations {
        if rule.source != value.source {
            return Err(invalid(
                "dataValidation source does not match its collection",
            ));
        }
        validate_rule(rule)?;
    }
    Ok(())
}

pub(crate) fn validate_rule(value: &Validation) -> Result<()> {
    validate_formula_cardinality(
        value.validation_type,
        value.operator,
        &value.formula1,
        &value.formula2,
    )?;
    validate_optional_text(value.error_title.as_deref(), 32, "errorTitle")?;
    validate_optional_text(value.error.as_deref(), 224, "error")?;
    validate_optional_text(value.prompt_title.as_deref(), 32, "promptTitle")?;
    validate_optional_text(value.prompt.as_deref(), 255, "prompt")?;
    if value.uid.as_deref().is_some_and(|uid| !valid_guid(uid)) {
        return Err(invalid("invalid data-validation uid"));
    }
    if value.source == Source::Core
        && (value.sqref.edited || value.sqref.split || value.sqref.adjusted || value.sqref.adjust)
    {
        return Err(invalid(
            "Office 2010 sqref flags are not valid on core data validation",
        ));
    }
    parse_sqref(
        &sqref_text(&value.sqref)?,
        value.sqref.edited,
        value.sqref.split,
        value.sqref.adjusted,
        value.sqref.adjust,
    )?;
    match value.formula1.as_ref() {
        Some(ListSource::Formula(value)) => validate_text(&value.0, MAX_FORMULA_BYTES, "formula1")?,
        Some(ListSource::QuotedList(list)) => {
            if value.source != Source::Office2010 {
                return Err(invalid(
                    "quoted-list source requires Office 2010 data validation",
                ));
            }
            validate_text(list, MAX_FORMULA_BYTES, "quoted validation list")?;
        },
        None => {},
    }
    if let Some(value) = value.formula2.as_ref() {
        validate_text(&value.0, MAX_FORMULA_BYTES, "formula2")?;
    }
    Ok(())
}

pub(crate) fn validate_optional_text(value: Option<&str>, max: usize, field: &str) -> Result<()> {
    if let Some(value) = value {
        if value.chars().count() > max {
            return Err(invalid(format!("{field} exceeds {max} characters")));
        }
        validate_xml_chars(value, field)?;
    }
    Ok(())
}

pub(crate) fn validate_text(value: &str, max_bytes: usize, field: &str) -> Result<()> {
    if value.len() > max_bytes {
        return Err(invalid(format!("{field} is too large")));
    }
    validate_xml_chars(value, field)
}

fn validate_xml_chars(value: &str, field: &str) -> Result<()> {
    if value.chars().any(|ch| {
        let code = ch as u32;
        !matches!(code, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x0010_FFFF)
    }) {
        return Err(invalid(format!(
            "{field} contains an invalid XML character"
        )));
    }
    Ok(())
}
