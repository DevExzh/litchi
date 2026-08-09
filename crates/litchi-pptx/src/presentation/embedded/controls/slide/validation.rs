//! Shared bounds for detached `ActiveX` control transactions.

use super::super::{MAX_BINARY_BYTES, MAX_SLIDE_XML_BYTES};
use crate::presentation::embedded::{MAX_ATTRIBUTE_BYTES, invalid, limit};
use crate::{Error, Result};

pub(crate) fn validate_text(value: &str, label: &'static str, allow_empty: bool) -> Result<()> {
    if value.len() > MAX_ATTRIBUTE_BYTES {
        return Err(limit(label, MAX_ATTRIBUTE_BYTES));
    }
    if !allow_empty && value.is_empty() {
        return Err(invalid(format!("{label} cannot be empty")));
    }
    if value
        .chars()
        .any(|character| matches!(character, '\u{0000}'..='\u{0008}' | '\u{000B}'..='\u{000C}' | '\u{000E}'..='\u{001F}'))
    {
        return Err(invalid(format!("{label} contains an invalid XML character")));
    }
    Ok(())
}

pub(crate) fn validate_binary(bytes: &[u8]) -> Result<()> {
    if bytes.len() > MAX_BINARY_BYTES {
        return Err(limit("control binary bytes", MAX_BINARY_BYTES));
    }
    Ok(())
}

pub(crate) fn validate_source_size(bytes: &[u8], label: &'static str) -> Result<()> {
    if bytes.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit(label, MAX_SLIDE_XML_BYTES));
    }
    Ok(())
}

pub(crate) fn invalid_revision() -> Error {
    invalid("ActiveX control patch source is stale")
}
