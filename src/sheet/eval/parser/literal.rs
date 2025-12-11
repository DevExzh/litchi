//! Literal parsing for formula expressions.

use crate::sheet::CellValue;

/// Try to parse a literal formula expression into a `CellValue`.
///
/// Supported forms:
/// - String literals: "text" (with "" as escaped quote)
/// - Boolean literals: TRUE / FALSE (case-insensitive)
/// - Numeric literals: integers or floats accepted by Rust's parsers
pub fn parse_literal(s: &str) -> Option<CellValue> {
    let trimmed = s.trim();
    if trimmed.is_empty() {
        return None;
    }

    if let Some(text) = parse_string_literal(trimmed) {
        return Some(CellValue::String(text));
    }

    if trimmed.eq_ignore_ascii_case("TRUE") {
        return Some(CellValue::Bool(true));
    }
    if trimmed.eq_ignore_ascii_case("FALSE") {
        return Some(CellValue::Bool(false));
    }

    if let Some(num) = parse_number_literal(trimmed) {
        return Some(num);
    }

    None
}

fn parse_string_literal(s: &str) -> Option<String> {
    let bytes = s.as_bytes();
    if bytes.len() < 2 || bytes.first() != Some(&b'"') || bytes.last() != Some(&b'"') {
        return None;
    }

    let inner = &s[1..s.len() - 1];
    if inner.is_empty() {
        return Some(String::new());
    }

    let mut out = String::new();
    let mut chars = inner.chars().peekable();
    while let Some(ch) = chars.next() {
        if ch == '"' {
            if chars.peek() == Some(&'"') {
                // Escaped quote "" -> '"'.
                out.push('"');
                chars.next();
            } else {
                // Lone quote inside string – treat as is.
                out.push(ch);
            }
        } else {
            out.push(ch);
        }
    }

    Some(out)
}

fn parse_number_literal(s: &str) -> Option<CellValue> {
    // First try integer parsing.
    if let Ok(int_val) = s.parse::<i64>() {
        return Some(CellValue::Int(int_val));
    }

    // Then fall back to floating-point.
    if let Ok(float_val) = s.parse::<f64>() {
        return Some(CellValue::Float(float_val));
    }

    None
}
