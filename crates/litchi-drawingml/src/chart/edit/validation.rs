//! Bounds and value validation for source-preserving chart edits.

use crate::{Error, Result};

pub(crate) const MAX_XML_BYTES: usize = 8 * 1024 * 1024;
pub(crate) const MAX_XML_NODES: usize = 250_000;
pub(crate) const MAX_XML_DEPTH: usize = 512;
pub(crate) const MAX_TEXT_BYTES: usize = 1 << 20;

pub(crate) fn validate_text(value: &str, description: &'static str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(Error::Limit {
            resource: description,
            limit: MAX_TEXT_BYTES,
        });
    }
    if value.chars().any(|character| character == '\0') {
        return Err(Error::Invalid(format!(
            "chart {description} contains a NUL character"
        )));
    }
    Ok(())
}

pub(crate) fn validate_style(style: Option<u32>) -> Result<()> {
    if style.is_some_and(|value| !(1..=48).contains(&value)) {
        return Err(Error::Invalid(
            "chart style must be between 1 and 48".into(),
        ));
    }
    Ok(())
}

pub(crate) fn validate_axis_range(min: Option<f64>, max: Option<f64>) -> Result<()> {
    if min.is_some_and(|value| !value.is_finite()) || max.is_some_and(|value| !value.is_finite()) {
        return Err(Error::Invalid("chart axis bounds must be finite".into()));
    }
    if min.zip(max).is_some_and(|(min, max)| min > max) {
        return Err(Error::Invalid(
            "chart axis minimum cannot exceed its maximum".into(),
        ));
    }
    Ok(())
}

pub(crate) fn limit(resource: &'static str, limit: usize) -> Error {
    Error::Limit { resource, limit }
}
