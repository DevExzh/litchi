use super::super::super::model::Limits;
use crate::{Error, Result};
pub(in crate::web) fn format_f64(value: f64) -> String {
    let mut buffer = ryu::Buffer::new();
    buffer.format_finite(value).to_owned()
}

pub(in crate::web) fn escape_attr(out: &mut String, value: &str) {
    for character in value.chars() {
        match character {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(character),
        }
    }
}

pub(in crate::web) fn require_nonempty(label: &str, value: &str) -> Result<()> {
    if value.is_empty() {
        invalid(format!("{label} must not be empty"))
    } else {
        Ok(())
    }
}

pub(in crate::web) fn enforce_count_with(
    label: &'static str,
    count: usize,
    limits: &Limits,
) -> Result<()> {
    if count > limits.items {
        limit(label, limits.items, count)
    } else {
        Ok(())
    }
}

pub(in crate::web) fn parse_bool(value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => invalid(format!("invalid XML boolean '{value}'")),
    }
}

pub(in crate::web) fn invalid<T>(message: String) -> Result<T> {
    Err(Error::Invalid(message))
}

pub(in crate::web) fn limit<T>(resource: &'static str, max: usize, actual: usize) -> Result<T> {
    Err(Error::Limit {
        resource,
        max,
        actual,
    })
}
