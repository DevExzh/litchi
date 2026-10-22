//! Semantic safety checks for classic worksheet comments.

use super::model::Comments;
use super::{MAX_TEXT_BYTES, codec};
use crate::error::{Result, invalid};

/// Validate the complete typed graph through the existing comments codec.
pub(crate) fn comments(value: &Comments) -> Result<()> {
    codec::validate_comments(value)
}

/// Validate author/comment text before it enters a staged graph.
pub(crate) fn text(value: &str, field: &str) -> Result<()> {
    if value.len() > MAX_TEXT_BYTES {
        return Err(invalid(format!(
            "classic comments {field} exceeds the configured text bound"
        )));
    }
    if value.chars().any(|character| {
        !matches!(
            character,
            '\u{9}'
                | '\u{A}'
                | '\u{D}'
                | '\u{20}'..='\u{D7FF}'
                | '\u{E000}'..='\u{FFFD}'
                | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(invalid(format!(
            "classic comments {field} contains a character forbidden by XML 1.0"
        )));
    }
    Ok(())
}
