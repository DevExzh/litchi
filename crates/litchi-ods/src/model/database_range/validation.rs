//! Semantic validation and resource limits for database ranges.

use super::semantic::{ConditionSource, Expression, Filter, Range, Source};
use crate::model::data_pilot;
use litchi_core::{Error, Result};

pub(super) const MAX_FILTER_DEPTH: usize = 128;
pub(super) const MAX_DATABASE_RANGES: usize = 65_536;
pub(super) const MAX_DATABASE_VALUE_BYTES: usize = 1_048_576;
pub(super) const MAX_DATABASE_ITEMS: usize = 262_144;
/// Aggregate authored metadata budget for one database-range owner.
///
/// Per-value limits alone would allow a large number of individually valid
/// strings to create an unbounded owned graph. This budget covers all strings
/// and collection nodes before a caller clones or serializes the owner.
pub const MAX_DATABASE_GRAPH_BYTES: usize = 32 * 1024 * 1024;
/// Aggregate collection-node budget for one database-range owner.
pub const MAX_DATABASE_GRAPH_NODES: usize = 262_144;

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
        data_pilot::parse_database_range_address(&self.target_range_address)?;
        if let Some(delay) = self.refresh_delay.as_deref()
            && !is_xsd_duration(delay)
        {
            return Err(invalid("table:refresh-delay", delay));
        }
        if let Some(filter) = &self.filter {
            validate_filter_for_owner(filter)?;
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
                data_pilot::parse_database_range_address(address)?;
            }
            for key in &sort.keys {
                validate_text("database sort data type", key.data_type.as_deref(), false)?;
            }
            validate_text("database sort algorithm", sort.algorithm.as_deref(), false)?;
            validate_sort_lexicals(sort)?;
        }
        if let Some(source) = &self.source {
            validate_source(source)?;
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

fn validate_sort_lexicals(sort: &super::semantic::Sort) -> Result<()> {
    if let Some(language) = sort.language.as_deref()
        && !is_language_code(language)
    {
        return Err(invalid("table:language", language));
    }
    if let Some(country) = sort.country.as_deref()
        && !is_code(country)
    {
        return Err(invalid("table:country", country));
    }
    if let Some(script) = sort.script.as_deref()
        && !is_code(script)
    {
        return Err(invalid("table:script", script));
    }
    if let Some(language) = sort.rfc_language_tag.as_deref()
        && !is_language_tag(language)
    {
        return Err(invalid("table:rfc-language-tag", language));
    }
    Ok(())
}

fn is_language_code(value: &str) -> bool {
    (1..=8).contains(&value.len()) && value.bytes().all(|byte| byte.is_ascii_alphabetic())
}

fn is_code(value: &str) -> bool {
    (1..=8).contains(&value.len()) && value.bytes().all(|byte| byte.is_ascii_alphanumeric())
}

fn is_language_tag(value: &str) -> bool {
    let mut parts = value.split('-');
    let Some(primary) = parts.next() else {
        return false;
    };
    is_language_code(primary)
        && parts.all(|part| {
            (1..=8).contains(&part.len()) && part.bytes().all(|byte| byte.is_ascii_alphanumeric())
        })
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
    let mut saw_unnamed = false;
    let mut graph_bytes = 0usize;
    let mut graph_nodes = 0usize;
    for range in ranges {
        range.validate()?;
        account_range(range, &mut graph_bytes, &mut graph_nodes)?;
        if let Some(name) = &range.name {
            if !names.insert(name.as_str()) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate database range name '{name}'"
                )));
            }
        } else if saw_unnamed {
            return Err(Error::InvalidFormat(
                "only one unnamed database range is allowed".to_string(),
            ));
        } else {
            saw_unnamed = true;
        }
    }
    Ok(())
}

fn account_range(range: &Range, bytes: &mut usize, nodes: &mut usize) -> Result<()> {
    account_node(256, bytes, nodes)?;
    for value in [
        range.name.as_deref(),
        Some(range.target_range_address.as_str()),
        range.refresh_delay.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        account_string(value, bytes, nodes)?;
    }
    if let Some(source) = &range.source {
        account_source(source, bytes, nodes)?;
    }
    if let Some(filter) = &range.filter {
        account_filter(filter, bytes, nodes)?;
    }
    if let Some(sort) = &range.sort {
        account_node(128, bytes, nodes)?;
        if let Some(address) = sort.target_range_address.as_deref() {
            account_string(address, bytes, nodes)?;
        }
        for value in [
            sort.language.as_deref(),
            sort.country.as_deref(),
            sort.script.as_deref(),
            sort.rfc_language_tag.as_deref(),
            sort.algorithm.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            account_string(value, bytes, nodes)?;
        }
        for key in &sort.keys {
            account_node(48, bytes, nodes)?;
            if let Some(data_type) = key.data_type.as_deref() {
                account_string(data_type, bytes, nodes)?;
            }
        }
    }
    if let Some(subtotals) = &range.subtotals {
        account_node(128, bytes, nodes)?;
        if let Some(groups) = &subtotals.sort_groups {
            account_node(48, bytes, nodes)?;
            if let Some(data_type) = groups.data_type.as_deref() {
                account_string(data_type, bytes, nodes)?;
            }
        }
        for rule in &subtotals.rules {
            account_node(48, bytes, nodes)?;
            for field in &rule.fields {
                account_node(48, bytes, nodes)?;
                account_string(&field.function, bytes, nodes)?;
            }
        }
    }
    Ok(())
}

fn account_source(source: &Source, bytes: &mut usize, nodes: &mut usize) -> Result<()> {
    account_node(48, bytes, nodes)?;
    match source {
        Source::Sql {
            database_name,
            statement,
            ..
        } => {
            account_string(database_name, bytes, nodes)?;
            account_string(statement, bytes, nodes)?;
        },
        Source::Table {
            database_name,
            table_name,
        } => {
            account_string(database_name, bytes, nodes)?;
            account_string(table_name, bytes, nodes)?;
        },
        Source::Query {
            database_name,
            query_name,
        } => {
            account_string(database_name, bytes, nodes)?;
            account_string(query_name, bytes, nodes)?;
        },
    }
    Ok(())
}

fn account_filter(filter: &Filter, bytes: &mut usize, nodes: &mut usize) -> Result<()> {
    account_node(96, bytes, nodes)?;
    for value in [
        filter.target_range_address.as_deref(),
        filter.condition_source_range_address.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        account_string(value, bytes, nodes)?;
    }
    account_expression(&filter.expression, bytes, nodes)
}

fn account_expression(expression: &Expression, bytes: &mut usize, nodes: &mut usize) -> Result<()> {
    account_node(64, bytes, nodes)?;
    match expression {
        Expression::Condition(condition) => {
            account_string(&condition.value, bytes, nodes)?;
            account_string(&condition.operator, bytes, nodes)?;
            for item in &condition.set_items {
                account_string(item, bytes, nodes)?;
            }
        },
        Expression::And(children) | Expression::Or(children) => {
            for child in children {
                account_expression(child, bytes, nodes)?;
            }
        },
    }
    Ok(())
}

fn account_string(value: &str, bytes: &mut usize, nodes: &mut usize) -> Result<()> {
    account_node(16, bytes, nodes)?;
    let encoded = value
        .len()
        .checked_mul(2)
        .ok_or_else(|| Error::InvalidFormat("database-range graph size overflow".to_string()))?;
    *bytes = bytes
        .checked_add(encoded)
        .ok_or_else(|| Error::InvalidFormat("database-range graph size overflow".to_string()))?;
    if *bytes > MAX_DATABASE_GRAPH_BYTES {
        return Err(Error::InvalidFormat(
            "database-range metadata graph exceeds the aggregate byte limit".to_string(),
        ));
    }
    Ok(())
}

fn account_node(weight: usize, bytes: &mut usize, nodes: &mut usize) -> Result<()> {
    *nodes = nodes.checked_add(1).ok_or_else(|| {
        Error::InvalidFormat("database-range graph node count overflow".to_string())
    })?;
    if *nodes > MAX_DATABASE_GRAPH_NODES {
        return Err(Error::InvalidFormat(
            "database-range metadata graph exceeds the node limit".to_string(),
        ));
    }
    *bytes = bytes
        .checked_add(weight)
        .ok_or_else(|| Error::InvalidFormat("database-range graph size overflow".to_string()))?;
    if *bytes > MAX_DATABASE_GRAPH_BYTES {
        return Err(Error::InvalidFormat(
            "database-range metadata graph exceeds the aggregate byte limit".to_string(),
        ));
    }
    Ok(())
}

pub(crate) fn validate_source(source: &Source) -> Result<()> {
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
    Ok(())
}

pub(crate) fn validate_source_for_output(source: &Source) -> Result<()> {
    validate_source(source)?;
    let mut bytes = 0;
    let mut nodes = 0;
    account_source(source, &mut bytes, &mut nodes)
}

/// Validate filter metadata including the optional addresses that are only
/// interpreted by an owner. Public writers use this same admission path so a
/// standalone filter cannot bypass the database-range limits.
pub(crate) fn validate_filter_for_owner(filter: &Filter) -> Result<()> {
    validate_filter(filter)?;
    for address in [
        filter.target_range_address.as_deref(),
        filter.condition_source_range_address.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        validate_text("database filter range address", Some(address), true)?;
        data_pilot::parse_database_range_address(address)?;
    }
    Ok(())
}

pub(crate) fn validate_filter_for_output(filter: &Filter) -> Result<()> {
    validate_filter_for_owner(filter)?;
    let mut bytes = 0;
    let mut nodes = 0;
    account_filter(filter, &mut bytes, &mut nodes)
}

pub(crate) fn validate_range_for_output(range: &Range) -> Result<()> {
    range.validate()?;
    let mut bytes = 0;
    let mut nodes = 0;
    account_range(range, &mut bytes, &mut nodes)
}

pub(crate) fn validate_text(label: &str, value: Option<&str>, required: bool) -> Result<()> {
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
    if !value.chars().all(is_xml_text_char) {
        return Err(Error::InvalidFormat(format!(
            "{label} contains an XML 1.0 control character"
        )));
    }
    Ok(())
}

fn is_xml_text_char(value: char) -> bool {
    matches!(
        value,
        '\u{9}'
            | '\u{a}'
            | '\u{d}'
            | '\u{20}'..='\u{d7ff}'
            | '\u{e000}'..='\u{fffd}'
            | '\u{10000}'..='\u{10ffff}'
    )
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
    if filter.condition_source != Some(ConditionSource::CellRange)
        && filter.condition_source_range_address.is_some()
    {
        return Err(Error::InvalidFormat(
            "table:condition-source-range-address requires a cell-range filter source".to_string(),
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
            if let Some(data_type) = condition.data_type {
                let valid = match data_type {
                    super::semantic::DataType::Text | super::semantic::DataType::Number => true,
                    super::semantic::DataType::TextColor
                    | super::semantic::DataType::DataStyleColor => {
                        is_color(&condition.value) || condition.value == "window-font-color"
                    },
                    super::semantic::DataType::BackgroundColor => {
                        is_color(&condition.value) || condition.value == "transparent"
                    },
                };
                if !valid {
                    return Err(Error::InvalidFormat(
                        "filter color value does not match its data type".to_string(),
                    ));
                }
            }
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

fn is_color(value: &str) -> bool {
    let bytes = value.as_bytes();
    bytes.len() == 7 && bytes[0] == b'#' && bytes[1..].iter().all(|byte| byte.is_ascii_hexdigit())
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
