//! Writer-side constraints for the typed DrawingML chart model.

/// Construct an invalid-input error without coupling record writers to a
/// concrete semantic family.
#[inline]
pub(super) fn invalid_chart_input(message: impl Into<String>) -> std::io::Error {
    std::io::Error::new(std::io::ErrorKind::InvalidInput, message.into())
}

/// Validate an optional unsigned percentage-like value.
#[inline]
pub(super) fn validate_optional_u32_range(
    value: Option<u32>,
    minimum: u32,
    maximum: u32,
    description: &str,
) -> std::io::Result<()> {
    if value.is_some_and(|value| !(minimum..=maximum).contains(&value)) {
        return Err(invalid_chart_input(format!(
            "{description} must be between {minimum} and {maximum}"
        )));
    }
    Ok(())
}

/// Validate the ECMA-376 Part 1 §21.2.3.46 `ST_Style` chart-style range.
///
/// The chart-space `style/@val` value is one of the 48 built-in chart styles;
/// omitting the element is the model's way to request the schema default.
#[inline]
pub(super) fn validate_chart_style(style: Option<u32>) -> std::io::Result<()> {
    if style.is_some_and(|style| !(1..=48).contains(&style)) {
        return Err(invalid_chart_input("chart style must be between 1 and 48"));
    }
    Ok(())
}
