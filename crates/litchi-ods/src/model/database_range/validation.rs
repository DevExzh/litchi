//! Semantic validation and resource limits for database ranges.

use super::semantic::{ConditionSource, Expression, Filter, Range, Source};
use crate::model::data_pilot;
use litchi_core::{Error, Result};

pub(super) const MAX_FILTER_DEPTH: usize = 128;
const MAX_DATABASE_RANGES: usize = 65_536;
const MAX_DATABASE_VALUE_BYTES: usize = 1_048_576;
const MAX_DATABASE_ITEMS: usize = 262_144;

impl Range {
    /// Validate required values and recursive schema constraints.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn validate(&self) -> Result<()> {
        if self.target_range_address.is_empty() {
            return Err(Error::InvalidFormat(
                "database range target address cannot be empty".to_string(),
            ));
        }
        validate_text("database range name", self.name.as_deref(), true)?;
        validate_text(
            "database range target address",
            Some(&self.target_range_address),
            true,
        )?;
        data_pilot::parse_data_pilot_range(&self.target_range_address)?;
        if let Some(delay) = self.refresh_delay.as_deref()
            && !is_xsd_duration(delay)
        {
            return Err(invalid("table:refresh-delay", delay));
        }
        if let Some(filter) = &self.filter {
            validate_filter(filter)?;
            for address in [
                filter.target_range_address.as_deref(),
                filter.condition_source_range_address.as_deref(),
            ]
            .into_iter()
            .flatten()
            {
                validate_text("database filter range address", Some(address), true)?;
                data_pilot::parse_data_pilot_range(address)?;
            }
        }
        if self.sort.as_ref().is_some_and(|sort| sort.keys.is_empty()) {
            return Err(Error::InvalidFormat(
                "database sort requires at least one sort key".to_string(),
            ));
        }
        if let Some(sort) = &self.sort {
            if sort.keys.len() > MAX_DATABASE_ITEMS {
                return too_many("database sort keys");
            }
            if let Some(address) = &sort.target_range_address {
                validate_text("database sort target address", Some(address), true)?;
                data_pilot::parse_data_pilot_range(address)?;
            }
            for key in &sort.keys {
                validate_text("database sort data type", key.data_type.as_deref(), false)?;
            }
        }
        if let Some(source) = &self.source {
            match source {
                Source::Sql {
                    database_name,
                    statement,
                    ..
                } => {
                    validate_text("database source name", Some(database_name), true)?;
                    validate_text("database SQL statement", Some(statement), true)?;
                },
                Source::Table {
                    database_name,
                    table_name,
                } => {
                    validate_text("database source name", Some(database_name), true)?;
                    validate_text("database table name", Some(table_name), true)?;
                },
                Source::Query {
                    database_name,
                    query_name,
                } => {
                    validate_text("database source name", Some(database_name), true)?;
                    validate_text("database query name", Some(query_name), true)?;
                },
            }
        }
        if let Some(subtotals) = &self.subtotals {
            if subtotals.rules.len() > MAX_DATABASE_ITEMS {
                return too_many("database subtotal rules");
            }
            if let Some(groups) = &subtotals.sort_groups {
                validate_text(
                    "subtotal sort data type",
                    groups.data_type.as_deref(),
                    false,
                )?;
            }
            for rule in &subtotals.rules {
                if rule.fields.len() > MAX_DATABASE_ITEMS {
                    return too_many("database subtotal fields");
                }
                for field in &rule.fields {
                    validate_text("database subtotal function", Some(&field.function), true)?;
                }
            }
        }
        Ok(())
    }
}

/// # Errors
///
/// Returns an error when a value violates the format or resource constraints.
pub fn validate_database_range_collection(ranges: &[Range]) -> Result<()> {
    use std::collections::HashSet;
    if ranges.len() > MAX_DATABASE_RANGES {
        return too_many("database ranges");
    }
    let mut names = HashSet::new();
    for range in ranges {
        range.validate()?;
        if let Some(name) = &range.name
            && !names.insert(name.as_str())
        {
            return Err(Error::InvalidFormat(format!(
                "duplicate database range name '{name}'"
            )));
        }
    }
    Ok(())
}

fn validate_text(label: &str, value: Option<&str>, required: bool) -> Result<()> {
    let Some(value) = value else {
        return Ok(());
    };
    if required && value.trim().is_empty() {
        return Err(Error::InvalidFormat(format!("{label} must not be empty")));
    }
    if value.len() > MAX_DATABASE_VALUE_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{label} exceeds {MAX_DATABASE_VALUE_BYTES} bytes"
        )));
    }
    Ok(())
}

fn too_many(label: &str) -> Result<()> {
    Err(Error::InvalidFormat(format!(
        "{label} exceed the supported resource limit"
    )))
}

/// # Errors
///
/// Returns an error when a value violates the format or resource constraints.
pub fn validate_filter(filter: &Filter) -> Result<()> {
    validate_filter_expression(&filter.expression, 0, None)?;
    if filter.condition_source == Some(ConditionSource::CellRange)
        && filter.condition_source_range_address.is_none()
    {
        return Err(Error::InvalidFormat(
            "cell-range filter source requires table:condition-source-range-address".to_string(),
        ));
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum FilterParent {
    And,
    Or,
}

pub(super) fn validate_filter_expression(
    expression: &Expression,
    depth: usize,
    parent: Option<FilterParent>,
) -> Result<()> {
    if depth > MAX_FILTER_DEPTH {
        return Err(Error::InvalidFormat(
            "filter expression exceeds the supported nesting limit".to_string(),
        ));
    }
    let (children, kind) = match expression {
        Expression::Condition(condition) => {
            validate_text("filter operator", Some(&condition.operator), true)?;
            validate_text("filter value", Some(&condition.value), false)?;
            if condition.set_items.len() > MAX_DATABASE_ITEMS {
                return too_many("filter set items");
            }
            for item in &condition.set_items {
                validate_text("filter set item", Some(item), false)?;
            }
            return Ok(());
        },
        Expression::And(children) => (children, FilterParent::And),
        Expression::Or(children) => (children, FilterParent::Or),
    };
    if children.is_empty() {
        return Err(Error::InvalidFormat(
            "filter boolean group cannot be empty".to_string(),
        ));
    }
    if children.len() > MAX_DATABASE_ITEMS {
        return too_many("filter expressions");
    }
    if parent == Some(kind) {
        return Err(Error::InvalidFormat(
            "ODF filter groups must alternate AND and OR operators".to_string(),
        ));
    }
    for child in children {
        validate_filter_expression(child, depth + 1, Some(kind))?;
    }
    Ok(())
}

fn is_xsd_duration(value: &str) -> bool {
    let bytes = value.as_bytes();
    let mut index = usize::from(bytes.first() == Some(&b'-'));
    if bytes.get(index) != Some(&b'P') {
        return false;
    }
    index += 1;
    let mut any = false;
    any |= consume_integer_component(bytes, &mut index, b'Y');
    any |= consume_integer_component(bytes, &mut index, b'M');
    any |= consume_integer_component(bytes, &mut index, b'D');
    if bytes.get(index) == Some(&b'T') {
        index += 1;
        let mut any_time = false;
        any_time |= consume_integer_component(bytes, &mut index, b'H');
        any_time |= consume_integer_component(bytes, &mut index, b'M');
        any_time |= consume_seconds(bytes, &mut index);
        if !any_time {
            return false;
        }
        any = true;
    }
    any && index == bytes.len()
}

fn consume_integer_component(bytes: &[u8], index: &mut usize, suffix: u8) -> bool {
    let mut end = *index;
    while bytes.get(end).is_some_and(u8::is_ascii_digit) {
        end += 1;
    }
    if end > *index && bytes.get(end) == Some(&suffix) {
        *index = end + 1;
        true
    } else {
        false
    }
}

fn consume_seconds(bytes: &[u8], index: &mut usize) -> bool {
    let mut end = *index;
    while bytes.get(end).is_some_and(u8::is_ascii_digit) {
        end += 1;
    }
    if end == *index {
        return false;
    }
    if bytes.get(end) == Some(&b'.') {
        end += 1;
        let start = end;
        while bytes.get(end).is_some_and(u8::is_ascii_digit) {
            end += 1;
        }
        if end == start {
            return false;
        }
    }
    if bytes.get(end) == Some(&b'S') {
        *index = end + 1;
        true
    } else {
        false
    }
}

pub(super) fn invalid(attribute: &str, value: &str) -> Error {
    Error::InvalidFormat(format!("invalid {attribute} value '{value}'"))
}

pub(super) fn missing(attribute: &str) -> Error {
    Error::InvalidFormat(format!("missing required {attribute}"))
}

pub(super) fn xml_error(error: quick_xml::Error) -> Error {
    Error::InvalidFormat(format!("database-range XML parsing error: {error}"))
}

pub(super) fn unexpected_eof(element: &str) -> Error {
    Error::InvalidFormat(format!("unexpected end of XML inside {element}"))
}
