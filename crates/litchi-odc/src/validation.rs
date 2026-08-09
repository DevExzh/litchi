//! ODF 1.4 chart range and inert formula grammar checks.

use crate::{Definition, Limits};
use litchi_core::{Error, Result};

#[derive(Clone, Copy, PartialEq, Eq)]
enum AddressKind {
    Cell,
    Column,
    Row,
}

/// Validate a whitespace-separated ODF 1.4 cell-range-address list.
///
/// This follows ODF 1.4 Part 3, sections 9.2.1 through 9.2.5. Quoted table
/// names use doubled apostrophes. Relative second endpoints such as `.B4` are
/// accepted, as emitted by `OpenDocument` producers.
///
/// # Errors
///
/// Returns an invalid-format error for an empty, over-limit, or malformed list.
pub fn validate_range_list(value: &str) -> Result<()> {
    let limits = Limits::default();
    validate_range_list_with_limits(value, limits.max_scalar_bytes(), limits.max_range_items())
}

/// Validate the inert lexical grammar of an `OpenDocument` formula.
///
/// Formula text is never evaluated. The check enforces the ODF 1.4 Part 3
/// section 19.646 prefix rule and the Part 4 delimiter, string, and reference
/// grammar used by cached chart cells.
///
/// # Errors
///
/// Returns an invalid-format error for malformed formula structure.
pub fn validate_formula(value: &str) -> Result<()> {
    validate_formula_with_limit(value, Limits::default().max_scalar_bytes())
}

pub(crate) fn validate_definition(definition: &Definition, limits: Limits) -> Result<()> {
    definition.validate()?;
    validate_definition_scalars(definition, limits.max_scalar_bytes())?;
    if definition.plot_area.axes.len() > limits.max_axes() {
        return invalid("ODC axis count exceeds the caller-selected limit");
    }
    if definition.plot_area.series.len() > limits.max_series() {
        return invalid("ODC series count exceeds the caller-selected limit");
    }
    for value in [definition.plot_area.cell_range_address.as_deref()]
        .into_iter()
        .flatten()
    {
        validate_range_list_with_limits(
            value,
            limits.max_scalar_bytes(),
            limits.max_range_items(),
        )?;
    }
    for text in [&definition.title, &definition.subtitle, &definition.footer]
        .into_iter()
        .flatten()
    {
        if let Some(value) = text.cell_range.as_deref() {
            validate_range_list_with_limits(
                value,
                limits.max_scalar_bytes(),
                limits.max_range_items(),
            )?;
        }
    }
    let mut expanded_points = 0usize;
    for axis in &definition.plot_area.axes {
        if let Some(value) = axis.categories_cell_range_address.as_deref() {
            validate_range_list_with_limits(
                value,
                limits.max_scalar_bytes(),
                limits.max_range_items(),
            )?;
        }
    }
    for series in &definition.plot_area.series {
        for value in [
            series.values_cell_range_address.as_deref(),
            series.label_cell_address.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            validate_range_list_with_limits(
                value,
                limits.max_scalar_bytes(),
                limits.max_range_items(),
            )?;
        }
        for domain in &series.domains {
            validate_range_list_with_limits(
                &domain.cell_range_address,
                limits.max_scalar_bytes(),
                limits.max_range_items(),
            )?;
        }
        for point in &series.data_points {
            let repeated = usize::try_from(point.repeated).map_err(|error| {
                Error::InvalidFormat(format!(
                    "ODC data-point repeat count exceeds this platform: {error}"
                ))
            })?;
            expanded_points = expanded_points
                .checked_add(repeated)
                .ok_or_else(|| Error::InvalidFormat("ODC data-point count overflow".into()))?;
            if expanded_points > limits.max_data_points() {
                return invalid("ODC data-point count exceeds the caller-selected limit");
            }
        }
    }
    if let Some(table) = &definition.cached_table {
        let mut expanded_rows = 0usize;
        let mut expanded_cells = 0usize;
        for row in table.header_rows.iter().chain(&table.rows) {
            let row_repeat = usize::try_from(row.repeated).map_err(|error| {
                Error::InvalidFormat(format!(
                    "ODC cached-row repeat count exceeds this platform: {error}"
                ))
            })?;
            expanded_rows = expanded_rows
                .checked_add(row_repeat)
                .ok_or_else(|| Error::InvalidFormat("ODC cached-row count overflow".into()))?;
            if expanded_rows > limits.max_cached_rows() {
                return invalid("ODC cached-row count exceeds the caller-selected limit");
            }
            let mut row_cells = 0usize;
            for cell in &row.cells {
                let cell_repeat = usize::try_from(cell.repeated).map_err(|error| {
                    Error::InvalidFormat(format!(
                        "ODC cached-cell repeat count exceeds this platform: {error}"
                    ))
                })?;
                row_cells = row_cells
                    .checked_add(cell_repeat)
                    .ok_or_else(|| Error::InvalidFormat("ODC cached-cell count overflow".into()))?;
                if let Some(formula) = cell.formula.as_deref() {
                    validate_formula_with_limit(formula, limits.max_scalar_bytes())?;
                }
            }
            expanded_cells = expanded_cells
                .checked_add(row_cells.checked_mul(row_repeat).ok_or_else(|| {
                    Error::InvalidFormat("ODC expanded cached-cell count overflow".into())
                })?)
                .ok_or_else(|| {
                    Error::InvalidFormat("ODC expanded cached-cell count overflow".into())
                })?;
            if expanded_cells > limits.max_cached_cells() {
                return invalid("ODC cached-cell count exceeds the caller-selected limit");
            }
        }
    }
    Ok(())
}

pub(crate) fn validate_range_list_with_limits(
    value: &str,
    max_bytes: usize,
    max_items: usize,
) -> Result<()> {
    if value.is_empty() || value.len() > max_bytes {
        return invalid("ODC cell range is empty or exceeds its byte limit");
    }
    let mut start = 0usize;
    let mut quoted = false;
    let bytes = value.as_bytes();
    let mut index = 0usize;
    let mut count = 0usize;
    while index <= bytes.len() {
        let boundary = index == bytes.len() || (!quoted && bytes[index].is_ascii_whitespace());
        if boundary {
            if start < index {
                validate_range(&value[start..index])?;
                count += 1;
                if count > max_items {
                    return invalid("ODC cell range list exceeds its item limit");
                }
            }
            index += 1;
            while index < bytes.len() && bytes[index].is_ascii_whitespace() {
                index += 1;
            }
            start = index;
            continue;
        }
        if bytes[index] == b'\'' {
            if quoted && bytes.get(index + 1) == Some(&b'\'') {
                index += 2;
                continue;
            }
            quoted = !quoted;
        }
        index += 1;
    }
    if quoted || count == 0 {
        return invalid("ODC cell range contains an unterminated quoted table name");
    }
    Ok(())
}

fn validate_definition_scalars(definition: &Definition, max_bytes: usize) -> Result<()> {
    for value in [
        definition.style_name.as_deref(),
        definition.width.as_deref(),
        definition.height.as_deref(),
        definition.plot_area.cell_range_address.as_deref(),
        definition.plot_area.style_name.as_deref(),
        definition.plot_area.x.as_deref(),
        definition.plot_area.y.as_deref(),
        definition.plot_area.width.as_deref(),
        definition.plot_area.height.as_deref(),
    ]
    .into_iter()
    .flatten()
    {
        validate_scalar_bytes(value, max_bytes)?;
    }
    for text in [&definition.title, &definition.subtitle, &definition.footer]
        .into_iter()
        .flatten()
    {
        for value in [
            Some(text.text.as_str()),
            text.cell_range.as_deref(),
            text.style_name.as_deref(),
            text.x.as_deref(),
            text.y.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            validate_scalar_bytes(value, max_bytes)?;
        }
    }
    if let Some(legend) = &definition.legend {
        for value in [
            legend.style_name.as_deref(),
            legend.title.as_deref(),
            legend.x.as_deref(),
            legend.y.as_deref(),
            legend.expansion.as_deref(),
            legend.expansion_aspect_ratio.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            validate_scalar_bytes(value, max_bytes)?;
        }
    }
    for axis in &definition.plot_area.axes {
        for value in [
            axis.name.as_deref(),
            axis.style_name.as_deref(),
            axis.categories_cell_range_address.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            validate_scalar_bytes(value, max_bytes)?;
        }
    }
    for series in &definition.plot_area.series {
        for value in [
            series.xml_id.as_deref(),
            series.values_cell_range_address.as_deref(),
            series.label_cell_address.as_deref(),
            series.attached_axis.as_deref(),
            series.style_name.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            validate_scalar_bytes(value, max_bytes)?;
        }
        for domain in &series.domains {
            validate_scalar_bytes(&domain.cell_range_address, max_bytes)?;
        }
        for point in &series.data_points {
            if let Some(style) = point.style_name.as_deref() {
                validate_scalar_bytes(style, max_bytes)?;
            }
        }
    }
    if let Some(table) = &definition.cached_table {
        validate_scalar_bytes(&table.name, max_bytes)?;
        for row in table.header_rows.iter().chain(&table.rows) {
            for cell in &row.cells {
                if let Some(formula) = cell.formula.as_deref() {
                    validate_scalar_bytes(formula, max_bytes)?;
                }
                match &cell.value {
                    crate::CachedValue::Currency { currency, .. } => {
                        validate_scalar_bytes(currency, max_bytes)?;
                    },
                    crate::CachedValue::Date(value)
                    | crate::CachedValue::Time(value)
                    | crate::CachedValue::String(value) => {
                        validate_scalar_bytes(value, max_bytes)?;
                    },
                    crate::CachedValue::Empty
                    | crate::CachedValue::Float(_)
                    | crate::CachedValue::Percentage(_)
                    | crate::CachedValue::Boolean(_) => {},
                }
            }
        }
    }
    Ok(())
}

fn validate_scalar_bytes(value: &str, max_bytes: usize) -> Result<()> {
    if value.len() > max_bytes {
        return invalid("ODC scalar exceeds the caller-selected byte limit");
    }
    Ok(())
}

fn validate_range(value: &str) -> Result<()> {
    let colon = find_unquoted(value, b':')?;
    match colon {
        Some(index) => {
            if find_unquoted(&value[index + 1..], b':')?.is_some() {
                return invalid("ODC cell range contains more than one colon");
            }
            let start = validate_address(&value[..index])?;
            let end = validate_address(&value[index + 1..])?;
            if start != end {
                return invalid("ODC range endpoints use different address kinds");
            }
        },
        None if validate_address(value)? != AddressKind::Cell => {
            return invalid("ODC whole-row or whole-column address requires a range");
        },
        None => {},
    }
    Ok(())
}

fn validate_address(value: &str) -> Result<AddressKind> {
    if value.is_empty() {
        return invalid("ODC cell address is empty");
    }
    let dot = rfind_unquoted(value, b'.')?;
    let cell = if let Some(index) = dot {
        validate_table_name(&value[..index])?;
        &value[index + 1..]
    } else {
        value
    };
    validate_cell(cell)
}

fn validate_table_name(value: &str) -> Result<()> {
    let table_name = value.strip_prefix('$').unwrap_or(value);
    if table_name.is_empty() {
        return Ok(());
    }
    if table_name.starts_with('\'') {
        if !table_name.ends_with('\'') || table_name.len() < 3 {
            return invalid("ODC quoted table name is incomplete");
        }
        let inner = &table_name[1..table_name.len() - 1];
        let mut chars = inner.chars();
        while let Some(character) = chars.next() {
            if character == '\'' && chars.next() != Some('\'') {
                return invalid("ODC quoted table-name apostrophe is not doubled");
            }
        }
        return Ok(());
    }
    if table_name.chars().any(|character| {
        character.is_whitespace()
            || matches!(character, '.' | '\'' | ':' | '[' | ']' | '#')
            || character.is_control()
    }) {
        return invalid("ODC unquoted table name contains a reserved character");
    }
    Ok(())
}

fn validate_cell(value: &str) -> Result<AddressKind> {
    let mut rest = value.strip_prefix('$').unwrap_or(value);
    let column_bytes = rest.bytes().take_while(u8::is_ascii_alphabetic).count();
    if column_bytes == 0 {
        let row = rest.strip_prefix('$').unwrap_or(rest);
        if row.is_empty() || row.starts_with('0') || !row.bytes().all(|byte| byte.is_ascii_digit())
        {
            return invalid("ODC row address must be a positive decimal integer");
        }
        return Ok(AddressKind::Row);
    }
    if !rest[..column_bytes]
        .bytes()
        .all(|byte| byte.is_ascii_uppercase())
    {
        return invalid("ODC cell column must contain uppercase ASCII letters");
    }
    rest = &rest[column_bytes..];
    if rest.is_empty() {
        return Ok(AddressKind::Column);
    }
    rest = rest.strip_prefix('$').unwrap_or(rest);
    if rest.is_empty() || rest.starts_with('0') || !rest.bytes().all(|byte| byte.is_ascii_digit()) {
        return invalid("ODC cell row must be a positive decimal integer");
    }
    Ok(AddressKind::Cell)
}

pub(crate) fn validate_formula_with_limit(value: &str, max_bytes: usize) -> Result<()> {
    if value.is_empty() || value.len() > max_bytes {
        return invalid("ODC formula is empty or exceeds its byte limit");
    }
    if value.chars().any(char::is_control) {
        return invalid("ODC formula contains a control character");
    }
    let prefixed_expression = strip_formula_prefix(value)?;
    let expression = prefixed_expression
        .strip_prefix("==")
        .or_else(|| prefixed_expression.strip_prefix('='))
        .unwrap_or(prefixed_expression);
    if expression.trim().is_empty() {
        return invalid("ODC formula expression is empty");
    }
    let bytes = expression.as_bytes();
    let mut index = 0usize;
    let mut parentheses = 0usize;
    let mut braces = 0usize;
    while index < bytes.len() {
        match bytes[index] {
            b'"' => {
                index += 1;
                loop {
                    let Some(byte) = bytes.get(index) else {
                        return invalid("ODC formula string is unterminated");
                    };
                    if *byte == b'"' {
                        if bytes.get(index + 1) == Some(&b'"') {
                            index += 2;
                            continue;
                        }
                        index += 1;
                        break;
                    }
                    index += 1;
                }
                continue;
            },
            b'[' => {
                let end = find_reference_end(&expression[index + 1..])? + index + 1;
                validate_formula_reference(&expression[index + 1..end])?;
                index = end + 1;
                continue;
            },
            b']' => return invalid("ODC formula contains an unmatched closing bracket"),
            b'(' => parentheses += 1,
            b')' => {
                parentheses = parentheses
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("ODC formula has unmatched ')'".into()))?;
            },
            b'{' => braces += 1,
            b'}' => {
                braces = braces
                    .checked_sub(1)
                    .ok_or_else(|| Error::InvalidFormat("ODC formula has unmatched '}'".into()))?;
            },
            _ => {},
        }
        index += 1;
    }
    if parentheses != 0 || braces != 0 {
        return invalid("ODC formula contains an unterminated delimiter");
    }
    Ok(())
}

fn strip_formula_prefix(value: &str) -> Result<&str> {
    let Some(colon) = value.find(':') else {
        return Ok(value);
    };
    let prefix = &value[..colon];
    if prefix.contains(['=', '(', '[', ' ']) {
        return Ok(value);
    }
    if !is_ncname(prefix) {
        return invalid("ODC formula namespace prefix is not an NCName");
    }
    let expression = &value[colon + 1..];
    if expression.is_empty() {
        return invalid("ODC formula after namespace prefix is empty");
    }
    Ok(expression)
}

fn is_ncname(value: &str) -> bool {
    let mut chars = value.chars();
    chars
        .next()
        .is_some_and(|character| character == '_' || character.is_alphabetic())
        && chars.all(|character| {
            character == '_' || character == '-' || character == '.' || character.is_alphanumeric()
        })
}

fn find_reference_end(value: &str) -> Result<usize> {
    let mut quoted = false;
    let bytes = value.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'\'' {
            if quoted && bytes.get(index + 1) == Some(&b'\'') {
                index += 2;
                continue;
            }
            quoted = !quoted;
        } else if bytes[index] == b']' && !quoted {
            return Ok(index);
        }
        index += 1;
    }
    invalid("ODC formula reference is unterminated")
}

fn validate_formula_reference(value: &str) -> Result<()> {
    if value == "#REF!" {
        return Ok(());
    }
    let range = if value.starts_with('\'') {
        let mut index = 1usize;
        let bytes = value.as_bytes();
        loop {
            let Some(byte) = bytes.get(index) else {
                return invalid("ODC formula reference source is unterminated");
            };
            if *byte == b'\'' {
                if bytes.get(index + 1) == Some(&b'\'') {
                    index += 2;
                    continue;
                }
                if bytes.get(index + 1) != Some(&b'#') {
                    return invalid("ODC formula reference source lacks '#'");
                }
                break &value[index + 2..];
            }
            index += 1;
        }
    } else {
        value
    };
    validate_range(range)
}

fn find_unquoted(value: &str, needle: u8) -> Result<Option<usize>> {
    let mut quoted = false;
    let bytes = value.as_bytes();
    let mut index = 0usize;
    while index < bytes.len() {
        if bytes[index] == b'\'' {
            if quoted && bytes.get(index + 1) == Some(&b'\'') {
                index += 2;
                continue;
            }
            quoted = !quoted;
        } else if bytes[index] == needle && !quoted {
            return Ok(Some(index));
        }
        index += 1;
    }
    if quoted {
        return invalid("ODC address contains an unterminated quote");
    }
    Ok(None)
}

fn rfind_unquoted(value: &str, needle: u8) -> Result<Option<usize>> {
    let mut result = None;
    let mut offset = 0usize;
    while let Some(index) = find_unquoted(&value[offset..], needle)? {
        result = Some(offset + index);
        offset += index + 1;
    }
    Ok(result)
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

#[cfg(test)]
mod tests {
    use super::{validate_formula, validate_range_list};

    #[test]
    fn odf_range_lists_accept_quoted_and_relative_addresses() {
        assert!(validate_range_list("Sheet1.$B$1:.$B$3").is_ok());
        assert!(validate_range_list("'My Sheet'.$A$1:'My Sheet'.$C$4 Other.D2").is_ok());
        assert!(validate_range_list("'Don''t'.$A$1").is_ok());
        assert!(validate_range_list("Sheet.a1").is_err());
        assert!(validate_range_list("'bad.A1").is_err());
        assert!(validate_range_list("Sheet.A0").is_err());
    }

    #[test]
    fn inert_formula_grammar_checks_prefix_delimiters_and_references() {
        assert!(validate_formula("of:=SUM([.B2:.B4])&\"x\"").is_ok());
        assert!(validate_formula("=1+2").is_ok());
        assert!(validate_formula("of:=SUM([.b2:.B4])").is_err());
        assert!(validate_formula("of:=SUM([.B2:.B4]").is_err());
        assert!(validate_formula("1bad:=1").is_err());
        assert!(validate_formula("of:").is_err());
    }
}
