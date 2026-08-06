//! Shared bounded validation for ODT parser XML codecs.

use super::super::{MAX_SEMANTIC_DEPTH, MAX_SEMANTIC_ITEMS};
use litchi_core::{Error, Result};

pub(super) fn parse_tracked_change_bool(name: &str, value: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "text:{name} is not an XML Schema boolean"
        ))),
    }
}

pub(super) fn validate_tracked_change_text(
    value: &str,
    description: &str,
    allow_empty: bool,
) -> Result<()> {
    const MAX_TRACKED_CHANGE_TEXT_BYTES: usize = 1024 * 1024;
    if value.len() > MAX_TRACKED_CHANGE_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{description} exceeds {MAX_TRACKED_CHANGE_TEXT_BYTES} bytes"
        )));
    }
    if !allow_empty && value.is_empty() {
        return Err(Error::InvalidFormat(format!(
            "{description} cannot be empty"
        )));
    }
    if value.chars().any(|character| {
        character == '\0' || (character.is_control() && !matches!(character, '\t' | '\n' | '\r'))
    }) {
        return Err(Error::InvalidFormat(format!(
            "{description} contains invalid XML characters"
        )));
    }
    Ok(())
}

pub(super) fn validate_protection_key(value: &str) -> Result<()> {
    validate_tracked_change_text(value, "text:protection-key", false)?;
    let mut symbols = 0usize;
    let mut padding = 0usize;
    let mut saw_padding = false;
    for byte in value.bytes() {
        if byte.is_ascii_whitespace() {
            continue;
        }
        symbols += 1;
        if byte == b'=' {
            saw_padding = true;
            padding += 1;
        } else if byte.is_ascii_alphanumeric() || matches!(byte, b'+' | b'/') {
            if saw_padding {
                return Err(Error::InvalidFormat(
                    "text:protection-key has data after base64 padding".to_string(),
                ));
            }
        } else {
            return Err(Error::InvalidFormat(
                "text:protection-key is not base64Binary".to_string(),
            ));
        }
    }
    if symbols == 0 || !symbols.is_multiple_of(4) || padding > 2 {
        return Err(Error::InvalidFormat(
            "text:protection-key is not base64Binary".to_string(),
        ));
    }
    Ok(())
}

pub(super) fn ensure_pending_capacity(length: usize) -> Result<()> {
    if length >= MAX_SEMANTIC_ITEMS {
        return Err(Error::InvalidFormat(format!(
            "document exceeds {MAX_SEMANTIC_ITEMS} annotation ranges"
        )));
    }
    Ok(())
}

pub(super) fn parse_boolean(value: &str, context: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(Error::InvalidFormat(format!(
            "{context} must be true, false, 1, or 0"
        ))),
    }
}

pub(super) fn checked_semantic_depth(depth: usize, context: &str) -> Result<usize> {
    let depth = depth
        .checked_add(1)
        .ok_or_else(|| Error::InvalidFormat(format!("{context} nesting depth overflow")))?;
    if depth > MAX_SEMANTIC_DEPTH {
        return Err(Error::InvalidFormat(format!(
            "{context} nesting exceeds {MAX_SEMANTIC_DEPTH} levels"
        )));
    }
    Ok(depth)
}
