//! Efficient ODF attribute parsing with compile-time validation.
//!
//! This module provides high-performance attribute parsing for ODF XML elements,
//! with compile-time validation of attribute names and types using perfect hash maps.
//!
//! **Note**: This module provides a complete public API for ODF attribute parsing.
//! Not all functions are used internally, but they are available for advanced users
//! who need direct attribute parsing capabilities.
//!
//! # Performance Optimizations
//!
//! - **Zero-copy parsing**: Attributes are parsed directly from byte slices
//! - **SIMD acceleration**: Uses `atoi_simd` and `fast_float2` for numeric parsing
//! - **Compile-time validation**: Uses `phf` for O(1) attribute lookup

#![allow(dead_code)] // Public API - complete attribute parsing utilities
//! - **Type-safe parsing**: Strong typing prevents runtime errors
//!
//! # References
//!
//! - ODF 1.2 Specification: Attribute types and valid values
//! - odfpy: `3rdparty/odfpy/odf/attrconverters.py`
//! - odfdo: `3rdparty/odfdo/src/odfdo/element.py`
use crate::common::{Error, Result};
use atoi_simd::parse_skipped;
use fast_float2::parse_partial;
use phf::{Map, Set, phf_map, phf_set};
use std::borrow::Cow;

// ============================================================================
// ATTRIBUTE TYPE ENUMERATION
// ============================================================================

/// ODF attribute types for type-safe parsing
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum AttrType {
    /// String attribute (most common)
    String,
    /// Boolean attribute (true/false)
    Boolean,
    /// Integer attribute
    Integer,
    /// Floating-point number
    Float,
    /// Length with unit (e.g., "2.5cm", "10pt")
    Length,
    /// Percentage (e.g., "50%")
    Percentage,
    /// Color (e.g., "#FF0000", "red")
    Color,
    /// Date/time (ISO 8601 format)
    DateTime,
    /// Duration (ISO 8601 duration format)
    Duration,
    /// URI/URL reference
    Uri,
    /// Enum value (limited set of valid values)
    Enum,
}

// ============================================================================
// COMMON ATTRIBUTE NAMES (COMPILE-TIME MAP)
// ============================================================================

/// Common ODF attribute names to their types
///
/// This map provides O(1) lookup for attribute validation.
/// Only frequently used attributes are included to keep compile times reasonable.
static ATTR_TYPES: Map<&'static str, AttrType> = phf_map! {
    // Common text attributes
    "text:style-name" => AttrType::String,
    "text:class-names" => AttrType::String,
    "text:cond-style-name" => AttrType::String,
    "text:outline-level" => AttrType::Integer,
    "text:restart-numbering" => AttrType::Boolean,
    "text:start-value" => AttrType::Integer,
    "text:is-list-header" => AttrType::Boolean,

    // Table attributes
    "table:name" => AttrType::String,
    "table:style-name" => AttrType::String,
    "table:number-columns-repeated" => AttrType::Integer,
    "table:number-rows-repeated" => AttrType::Integer,
    "table:number-columns-spanned" => AttrType::Integer,
    "table:number-rows-spanned" => AttrType::Integer,
    "table:protected" => AttrType::Boolean,
    "table:print" => AttrType::Boolean,
    "table:display" => AttrType::Boolean,

    // Office value attributes
    "office:value-type" => AttrType::Enum,
    "office:value" => AttrType::Float,
    "office:string-value" => AttrType::String,
    "office:boolean-value" => AttrType::Boolean,
    "office:date-value" => AttrType::DateTime,
    "office:time-value" => AttrType::Duration,
    "office:currency" => AttrType::String,

    // Style attributes
    "style:name" => AttrType::String,
    "style:family" => AttrType::Enum,
    "style:parent-style-name" => AttrType::String,
    "style:display-name" => AttrType::String,
    "style:class" => AttrType::String,
    "style:data-style-name" => AttrType::String,
    "style:list-style-name" => AttrType::String,
    "style:master-page-name" => AttrType::String,
    "style:page-layout-name" => AttrType::String,

    // Formatting attributes
    "fo:font-weight" => AttrType::Enum,
    "fo:font-style" => AttrType::Enum,
    "fo:font-size" => AttrType::Length,
    "fo:color" => AttrType::Color,
    "fo:background-color" => AttrType::Color,
    "fo:text-align" => AttrType::Enum,
    "fo:margin" => AttrType::Length,
    "fo:margin-left" => AttrType::Length,
    "fo:margin-right" => AttrType::Length,
    "fo:margin-top" => AttrType::Length,
    "fo:margin-bottom" => AttrType::Length,
    "fo:padding" => AttrType::Length,
    "fo:border" => AttrType::String,
    "fo:text-indent" => AttrType::Length,
    "fo:line-height" => AttrType::Length,

    // Drawing attributes
    "draw:name" => AttrType::String,
    "draw:style-name" => AttrType::String,
    "draw:text-style-name" => AttrType::String,
    "draw:layer" => AttrType::String,
    "draw:id" => AttrType::String,
    "draw:z-index" => AttrType::Integer,

    // SVG attributes
    "svg:width" => AttrType::Length,
    "svg:height" => AttrType::Length,
    "svg:x" => AttrType::Length,
    "svg:y" => AttrType::Length,
    "svg:viewBox" => AttrType::String,

    // XLink attributes
    "xlink:href" => AttrType::Uri,
    "xlink:type" => AttrType::Enum,
    "xlink:show" => AttrType::Enum,
    "xlink:actuate" => AttrType::Enum,

    // Presentation attributes
    "presentation:class" => AttrType::Enum,
    "presentation:style-name" => AttrType::String,
    "presentation:user-transformed" => AttrType::Boolean,
    "presentation:placeholder" => AttrType::Boolean,

    // Chart attributes
    "chart:class" => AttrType::Enum,
    "chart:style-name" => AttrType::String,

    // Form attributes
    "form:name" => AttrType::String,
    "form:id" => AttrType::String,
    "form:control-implementation" => AttrType::String,

    // Meta attributes
    "meta:name" => AttrType::String,
    "meta:value-type" => AttrType::Enum,
};

// ============================================================================
// ENUM VALUE VALIDATION SETS
// ============================================================================

/// Valid values for office:value-type
static VALUE_TYPES: Set<&'static str> = phf_set! {
    "float", "percentage", "currency", "date", "time", "boolean", "string",
};

/// Valid values for style:family
static STYLE_FAMILIES: Set<&'static str> = phf_set! {
    "text", "paragraph", "section", "ruby", "table", "table-column",
    "table-row", "table-cell", "graphic", "presentation", "drawing-page",
    "chart",
};

/// Valid values for fo:font-weight
static FONT_WEIGHTS: Set<&'static str> = phf_set! {
    "normal", "bold", "100", "200", "300", "400", "500", "600", "700", "800", "900",
};

/// Valid values for fo:font-style
static FONT_STYLES: Set<&'static str> = phf_set! {
    "normal", "italic", "oblique",
};

/// Valid values for fo:text-align
static TEXT_ALIGNS: Set<&'static str> = phf_set! {
    "start", "end", "left", "right", "center", "justify",
};

// ============================================================================
// ATTRIBUTE PARSING FUNCTIONS
// ============================================================================

/// Parse a boolean attribute value
///
/// ODF uses "true"/"false" for boolean values.
///
/// # Arguments
///
/// * `value` - The attribute value as bytes
///
/// # Returns
///
/// The parsed boolean value, or error if invalid
///
/// # Examples
///
/// ```
/// # use litchi::odf::elements::attr_parser::parse_bool;
/// assert_eq!(parse_bool(b"true").unwrap(), true);
/// assert_eq!(parse_bool(b"false").unwrap(), false);
/// ```
#[inline]
pub fn parse_bool(value: &[u8]) -> Result<bool> {
    match value {
        b"true" => Ok(true),
        b"false" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "Invalid boolean value: {}",
            String::from_utf8_lossy(value)
        ))),
    }
}

/// Parse an integer attribute value using SIMD acceleration
///
/// # Arguments
///
/// * `value` - The attribute value as bytes
///
/// # Returns
///
/// The parsed integer value, or error if invalid
///
/// # Examples
///
/// ```
/// # use litchi::odf::elements::attr_parser::parse_int;
/// assert_eq!(parse_int(b"42").unwrap(), 42);
/// assert_eq!(parse_int(b"-123").unwrap(), -123);
/// ```
#[inline]
pub fn parse_int(value: &[u8]) -> Result<i64> {
    parse_skipped::<i64>(value).ok().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Invalid integer value: {}",
            String::from_utf8_lossy(value)
        ))
    })
}

/// Parse a positive integer attribute value using SIMD acceleration
///
/// This is optimized for common attributes like repetition counts.
///
/// # Arguments
///
/// * `value` - The attribute value as bytes
///
/// # Returns
///
/// The parsed unsigned integer value, or error if invalid
#[inline]
pub fn parse_uint(value: &[u8]) -> Result<u64> {
    parse_skipped::<u64>(value).ok().ok_or_else(|| {
        Error::InvalidFormat(format!(
            "Invalid unsigned integer value: {}",
            String::from_utf8_lossy(value)
        ))
    })
}

/// Parse a floating-point attribute value using fast_float2
///
/// # Arguments
///
/// * `value` - The attribute value as bytes
///
/// # Returns
///
/// The parsed float value, or error if invalid
///
/// # Examples
///
/// ```
/// # use litchi::odf::elements::attr_parser::parse_float;
/// assert!((parse_float(b"3.14").unwrap() - 3.14).abs() < 0.0001);
/// assert!((parse_float(b"-2.5e10").unwrap() + 2.5e10).abs() < 1.0);
/// ```
#[inline]
pub fn parse_float(value: &[u8]) -> Result<f64> {
    let s = std::str::from_utf8(value)
        .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in float value".to_string()))?;

    match parse_partial::<f64, _>(s) {
        Ok((num, _)) => Ok(num),
        Err(_) => Err(Error::InvalidFormat(format!("Invalid float value: {}", s))),
    }
}

/// Parse a length attribute with unit (e.g., "2.5cm", "10pt")
///
/// Returns the value and unit as separate components.
///
/// # Arguments
///
/// * `value` - The attribute value as bytes
///
/// # Returns
///
/// A tuple of (numeric_value, unit_string), or error if invalid
///
/// # Examples
///
/// ```
/// # use litchi::odf::elements::attr_parser::parse_length;
/// let (val, unit) = parse_length(b"2.5cm").unwrap();
/// assert!((val - 2.5).abs() < 0.0001);
/// assert_eq!(unit, "cm");
/// ```
pub fn parse_length(value: &[u8]) -> Result<(f64, Cow<'_, str>)> {
    let s = std::str::from_utf8(value)
        .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in length value".to_string()))?;

    // Try to parse the numeric part
    match parse_partial::<f64, _>(s) {
        Ok((num, consumed)) => {
            let unit = &s[consumed..].trim();
            if unit.is_empty() {
                // Default unit is points
                Ok((num, Cow::Borrowed("pt")))
            } else {
                Ok((num, Cow::Borrowed(unit)))
            }
        },
        Err(_) => Err(Error::InvalidFormat(format!("Invalid length value: {}", s))),
    }
}

/// Parse a percentage attribute (e.g., "50%")
///
/// Returns the percentage as a float (e.g., 0.5 for "50%").
///
/// # Arguments
///
/// * `value` - The attribute value as bytes
///
/// # Returns
///
/// The percentage as a decimal value (0.0 to 1.0+), or error if invalid
#[inline]
pub fn parse_percentage(value: &[u8]) -> Result<f64> {
    let s = std::str::from_utf8(value)
        .map_err(|_| Error::InvalidFormat("Invalid UTF-8 in percentage value".to_string()))?;

    let trimmed = s.trim_end_matches('%').trim();
    match parse_partial::<f64, _>(trimmed) {
        Ok((num, _)) => Ok(num / 100.0),
        Err(_) => Err(Error::InvalidFormat(format!(
            "Invalid percentage value: {}",
            s
        ))),
    }
}

/// Parse a color attribute (hex format like "#FF0000")
///
/// Returns RGB components as (r, g, b) where each is 0-255.
///
/// # Arguments
///
/// * `value` - The attribute value as bytes
///
/// # Returns
///
/// RGB tuple (r, g, b), or error if invalid
///
/// # Examples
///
/// ```
/// # use litchi::odf::elements::attr_parser::parse_color;
/// let (r, g, b) = parse_color(b"#FF0000").unwrap();
/// assert_eq!((r, g, b), (255, 0, 0));
/// ```
pub fn parse_color(value: &[u8]) -> Result<(u8, u8, u8)> {
    if value.is_empty() || value[0] != b'#' {
        return Err(Error::InvalidFormat("Color must start with #".to_string()));
    }

    let hex = &value[1..];
    if hex.len() != 6 {
        return Err(Error::InvalidFormat(
            "Color hex must be 6 characters".to_string(),
        ));
    }

    // Parse hex values using SIMD if possible
    let r = parse_hex_byte(&hex[0..2])?;
    let g = parse_hex_byte(&hex[2..4])?;
    let b = parse_hex_byte(&hex[4..6])?;

    Ok((r, g, b))
}

/// Parse a 2-digit hex value to u8
#[inline]
fn parse_hex_byte(hex: &[u8]) -> Result<u8> {
    let high = hex_digit(hex[0])?;
    let low = hex_digit(hex[1])?;
    Ok((high << 4) | low)
}

/// Convert a single hex digit character to its numeric value
#[inline]
fn hex_digit(c: u8) -> Result<u8> {
    match c {
        b'0'..=b'9' => Ok(c - b'0'),
        b'A'..=b'F' => Ok(c - b'A' + 10),
        b'a'..=b'f' => Ok(c - b'a' + 10),
        _ => Err(Error::InvalidFormat(format!(
            "Invalid hex digit: {}",
            c as char
        ))),
    }
}

/// Validate an enum attribute value against a set of valid values
///
/// # Arguments
///
/// * `value` - The attribute value as string
/// * `valid_set` - The set of valid values
///
/// # Returns
///
/// Ok if valid, error otherwise
#[inline]
pub fn validate_enum(value: &str, valid_set: &Set<&'static str>) -> Result<()> {
    if valid_set.contains(value) {
        Ok(())
    } else {
        Err(Error::InvalidFormat(format!(
            "Invalid enum value: {}",
            value
        )))
    }
}

/// Validate office:value-type attribute
#[inline]
pub fn validate_value_type(value: &str) -> Result<()> {
    validate_enum(value, &VALUE_TYPES)
}

/// Validate style:family attribute
#[inline]
pub fn validate_style_family(value: &str) -> Result<()> {
    validate_enum(value, &STYLE_FAMILIES)
}

/// Validate fo:font-weight attribute
#[inline]
pub fn validate_font_weight(value: &str) -> Result<()> {
    validate_enum(value, &FONT_WEIGHTS)
}

/// Validate fo:font-style attribute
#[inline]
pub fn validate_font_style(value: &str) -> Result<()> {
    validate_enum(value, &FONT_STYLES)
}

/// Validate fo:text-align attribute
#[inline]
pub fn validate_text_align(value: &str) -> Result<()> {
    validate_enum(value, &TEXT_ALIGNS)
}

// ============================================================================
// ATTRIBUTE PARSER
// ============================================================================

/// High-level attribute parser with type validation
///
/// This provides a convenient interface for parsing attributes with automatic
/// type detection and validation.
pub struct AttrParser;

impl AttrParser {
    /// Get the expected type for an attribute name
    #[inline]
    pub fn attr_type(name: &str) -> Option<AttrType> {
        ATTR_TYPES.get(name).copied()
    }

    /// Parse an attribute value based on its expected type
    pub fn parse_typed(name: &str, value: &[u8]) -> Result<ParsedValue> {
        let attr_type = Self::attr_type(name).unwrap_or(AttrType::String);

        match attr_type {
            AttrType::String => Ok(ParsedValue::String(
                String::from_utf8(value.to_vec())
                    .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?,
            )),
            AttrType::Boolean => Ok(ParsedValue::Boolean(parse_bool(value)?)),
            AttrType::Integer => Ok(ParsedValue::Integer(parse_int(value)?)),
            AttrType::Float => Ok(ParsedValue::Float(parse_float(value)?)),
            AttrType::Length => {
                let (num, unit) = parse_length(value)?;
                Ok(ParsedValue::Length(num, unit.into_owned()))
            },
            AttrType::Percentage => Ok(ParsedValue::Percentage(parse_percentage(value)?)),
            AttrType::Color => {
                let (r, g, b) = parse_color(value)?;
                Ok(ParsedValue::Color(r, g, b))
            },
            _ => Ok(ParsedValue::String(
                String::from_utf8(value.to_vec())
                    .map_err(|_| Error::InvalidFormat("Invalid UTF-8".to_string()))?,
            )),
        }
    }
}

/// Parsed attribute value (strongly typed)
#[derive(Debug, Clone, PartialEq)]
pub enum ParsedValue {
    String(String),
    Boolean(bool),
    Integer(i64),
    Float(f64),
    Length(f64, String),
    Percentage(f64),
    Color(u8, u8, u8),
}

// ============================================================================
// TESTS
// ============================================================================

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_parse_bool() {
        assert_eq!(parse_bool(b"true").unwrap(), true);
        assert_eq!(parse_bool(b"false").unwrap(), false);
        assert!(parse_bool(b"invalid").is_err());
    }

    #[test]
    fn test_parse_int() {
        assert_eq!(parse_int(b"42").unwrap(), 42);
        assert_eq!(parse_int(b"-123").unwrap(), -123);
        assert_eq!(parse_int(b"0").unwrap(), 0);
    }

    #[test]
    fn test_parse_float() {
        assert!((parse_float(b"3.14").unwrap() - 3.14).abs() < 0.0001);
        assert!((parse_float(b"-2.5").unwrap() + 2.5).abs() < 0.0001);
        assert!((parse_float(b"1e10").unwrap() - 1e10).abs() < 1.0);
    }

    #[test]
    fn test_parse_length() {
        let (val, unit) = parse_length(b"2.5cm").unwrap();
        assert!((val - 2.5).abs() < 0.0001);
        assert_eq!(unit, "cm");

        let (val, unit) = parse_length(b"10pt").unwrap();
        assert!((val - 10.0).abs() < 0.0001);
        assert_eq!(unit, "pt");
    }

    #[test]
    fn test_parse_percentage() {
        assert!((parse_percentage(b"50%").unwrap() - 0.5).abs() < 0.0001);
        assert!((parse_percentage(b"100%").unwrap() - 1.0).abs() < 0.0001);
        assert!((parse_percentage(b"25.5%").unwrap() - 0.255).abs() < 0.0001);
    }

    #[test]
    fn test_parse_color() {
        let (r, g, b) = parse_color(b"#FF0000").unwrap();
        assert_eq!((r, g, b), (255, 0, 0));

        let (r, g, b) = parse_color(b"#00FF00").unwrap();
        assert_eq!((r, g, b), (0, 255, 0));

        let (r, g, b) = parse_color(b"#0000FF").unwrap();
        assert_eq!((r, g, b), (0, 0, 255));
    }

    #[test]
    fn test_validate_value_type() {
        assert!(validate_value_type("float").is_ok());
        assert!(validate_value_type("string").is_ok());
        assert!(validate_value_type("invalid").is_err());
    }

    #[test]
    fn test_validate_style_family() {
        assert!(validate_style_family("text").is_ok());
        assert!(validate_style_family("paragraph").is_ok());
        assert!(validate_style_family("invalid").is_err());
    }

    #[test]
    fn test_attr_parser() {
        let value = AttrParser::parse_typed("office:boolean-value", b"true").unwrap();
        assert_eq!(value, ParsedValue::Boolean(true));

        let value = AttrParser::parse_typed("table:number-columns-repeated", b"5").unwrap();
        assert_eq!(value, ParsedValue::Integer(5));

        let value = AttrParser::parse_typed("fo:font-size", b"12pt").unwrap();
        match value {
            ParsedValue::Length(num, unit) => {
                assert!((num - 12.0).abs() < 0.0001);
                assert_eq!(unit, "pt");
            },
            _ => panic!("Expected Length value"),
        }
    }
}
