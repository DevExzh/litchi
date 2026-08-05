//! Contextual validators for bounded ODF lexical values.
//!
//! These helpers validate neutral lexical contracts while leaving each format
//! feature responsible for its own limits and surrounding semantics.

use litchi_core::{Error, Result};

/// Validate a finite numeric lexical value and preserve the caller's context
/// in any diagnostic.
pub fn validate_finite_number(context: &str, value: &str) -> Result<()> {
    let parsed: f64 = value.parse().map_err(|_| {
        Error::InvalidFormat(format!(
            "{context} requires a numeric value, found '{value}'"
        ))
    })?;
    if !parsed.is_finite() {
        return Err(Error::InvalidFormat(format!(
            "{context} requires a finite numeric value, found '{value}'"
        )));
    }
    Ok(())
}

/// Validate an ODF `#RRGGBB` lexical color and preserve the caller's context
/// in any diagnostic.
pub fn validate_rgb_color(context: &str, value: &str) -> Result<()> {
    if value.len() != 7
        || !value.starts_with('#')
        || !value[1..].bytes().all(|byte| byte.is_ascii_hexdigit())
    {
        return Err(Error::InvalidFormat(format!(
            "invalid {context} color '{value}'"
        )));
    }
    Ok(())
}

/// Validate that one lexical value fits within a caller-owned byte limit.
pub fn validate_byte_limit(context: &str, value: &str, maximum: usize) -> Result<()> {
    if value.len() > maximum {
        return Err(Error::InvalidFormat(format!(
            "{context} exceeds the {maximum} byte safety limit"
        )));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn validates_finite_numbers_with_context() {
        assert!(validate_finite_number("calcext:value", "-1.25e2").is_ok());
        assert_eq!(
            validate_finite_number("calcext:value", "NaN")
                .unwrap_err()
                .to_string(),
            "Invalid format: calcext:value requires a finite numeric value, found 'NaN'"
        );
        assert_eq!(
            validate_finite_number("calcext:value", "not-a-number")
                .unwrap_err()
                .to_string(),
            "Invalid format: calcext:value requires a numeric value, found 'not-a-number'"
        );
    }

    #[test]
    fn validates_rgb_colors() {
        assert!(validate_rgb_color("calcext:color", "#0aFf19").is_ok());
        assert_eq!(
            validate_rgb_color("calcext:color", "blue")
                .unwrap_err()
                .to_string(),
            "Invalid format: invalid calcext:color color 'blue'"
        );
        assert!(validate_rgb_color("calcext:color", "#ééé").is_err());
    }

    #[test]
    fn enforces_the_caller_owned_byte_limit() {
        assert!(validate_byte_limit("calcext:value", "12345", 5).is_ok());
        assert_eq!(
            validate_byte_limit("calcext:value", "123456", 5)
                .unwrap_err()
                .to_string(),
            "Invalid format: calcext:value exceeds the 5 byte safety limit"
        );
    }
}
