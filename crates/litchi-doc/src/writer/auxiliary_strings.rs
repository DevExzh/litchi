//! Associated-document and save-history string-table authoring.

use super::core::WriteError;
use super::fib::FibBuilder;
use crate::{DocumentAssociatedStrings, SavedByTable};

fn invalid(error: impl std::fmt::Display) -> WriteError {
    WriteError::InvalidData(error.to_string())
}

fn checked_length(data: &[u8], table: &str) -> Result<u32, WriteError> {
    u32::try_from(data.len()).map_err(|_| WriteError::InvalidData(format!("{table} exceeds 4 GiB")))
}

fn checked_end(offset: u32, length: u32) -> Result<u32, WriteError> {
    offset
        .checked_add(length)
        .ok_or_else(|| WriteError::InvalidData("DOC table offset overflows".into()))
}

/// Validate and append the mandatory `SttbfAssoc` and optional `SttbSavedBy`.
///
/// Both tables and their complete offset plan are built before the stream is
/// modified, so validation and range failures cannot produce partial output.
pub(super) fn append_auxiliary_string_tables(
    fib: &mut FibBuilder,
    table_stream: &mut Vec<u8>,
    associated: &DocumentAssociatedStrings,
    saved_by: Option<&SavedByTable>,
    table_offset: u32,
) -> Result<u32, WriteError> {
    let actual_offset = u32::try_from(table_stream.len())
        .map_err(|_| WriteError::InvalidData("DOC table stream exceeds 4 GiB".into()))?;
    if actual_offset != table_offset {
        return Err(WriteError::InvalidData(
            "DOC table-stream offset plan is inconsistent".into(),
        ));
    }

    let associated = associated.to_bytes().map_err(invalid)?;
    let saved_by = saved_by
        .map(SavedByTable::to_bytes)
        .transpose()
        .map_err(invalid)?
        .map(|data| checked_length(&data, "SttbSavedBy").map(|length| (data, length)))
        .transpose()?;
    let associated_length = checked_length(&associated, "SttbfAssoc")?;
    let saved_by_length = saved_by.as_ref().map_or(0, |(_, length)| *length);
    let saved_by_offset = checked_end(table_offset, associated_length)?;
    let final_offset = checked_end(saved_by_offset, saved_by_length)?;

    fib.set_sttbf_assoc(table_offset, associated_length);
    table_stream.extend_from_slice(&associated);
    if let Some((data, length)) = saved_by {
        fib.set_sttb_saved_by(saved_by_offset, length);
        table_stream.extend_from_slice(&data);
    }
    Ok(final_offset)
}
