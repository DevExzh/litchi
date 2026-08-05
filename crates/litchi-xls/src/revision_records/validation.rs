//! Shared MS-XLS invariant checks for revision-record payloads.

use crate::{Error, Result};

use super::codec::{REN_SHEET_MAX_COMPRESSED_CHARS, REN_SHEET_MAX_UTF16_CHARS, STRING_HIGH_BYTE};

pub(crate) fn invalid(record_type: u16, message: impl Into<String>) -> Error {
    Error::InvalidRecord {
        record_type,
        message: message.into(),
    }
}

/// Validate the MS-XLS name-length split for `RRDRenSheet` and `RRInsertSh`.
pub(crate) fn validate_sheet_name_chars(
    record_type: u16,
    field: &[u8],
    cch: u16,
    context: &str,
) -> Result<()> {
    let wide = field
        .first()
        .map(|flags| flags & STRING_HIGH_BYTE != 0)
        .unwrap_or(false);
    let maximum = if wide {
        REN_SHEET_MAX_UTF16_CHARS
    } else {
        REN_SHEET_MAX_COMPRESSED_CHARS
    };
    if cch > maximum {
        return Err(invalid(
            record_type,
            format!("{context} has {cch} characters; maximum is {maximum}"),
        ));
    }
    Ok(())
}

/// Validate that an empty begin/end marker record has no payload.
pub(crate) fn validate_empty_marker(record_type: u16, data: &[u8], name: &str) -> Result<()> {
    if !data.is_empty() {
        return Err(invalid(
            record_type,
            format!("{name} payload has {} bytes; expected 0", data.len()),
        ));
    }
    Ok(())
}
