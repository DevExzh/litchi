//! Glossary-only DOC table authoring.

use super::core::WriteError;
use super::fib::FibBuilder;
use crate::GlossaryMetadata;

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

/// Validate, plan, and append all three parallel glossary tables atomically.
pub(super) fn append_glossary_tables(
    fib: &mut FibBuilder,
    table_stream: &mut Vec<u8>,
    metadata: Option<&GlossaryMetadata>,
    table_offset: u32,
    main_text_length: u32,
    text_stream: &[u8],
) -> Result<u32, WriteError> {
    let Some(metadata) = metadata else {
        return Ok(table_offset);
    };
    if metadata.main_text_length() != main_text_length {
        return Err(WriteError::InvalidData(format!(
            "glossary ccpText is {}; generated main story has {main_text_length} CPs",
            metadata.main_text_length()
        )));
    }
    metadata
        .validate_utf16_bytes(text_stream)
        .map_err(invalid)?;
    let actual_offset = u32::try_from(table_stream.len())
        .map_err(|_| WriteError::InvalidData("DOC table stream exceeds 4 GiB".into()))?;
    if actual_offset != table_offset {
        return Err(WriteError::InvalidData(
            "DOC table-stream offset plan is inconsistent".into(),
        ));
    }

    let tables = metadata.to_table_bytes().map_err(invalid)?;
    let (items, positions, styles) = tables.into_parts();
    let items_length = checked_length(&items, "SttbfGlsy")?;
    let positions_length = checked_length(&positions, "PlcfGlsy")?;
    let styles_length = checked_length(&styles, "SttbGlsyStyle")?;
    let positions_offset = checked_end(table_offset, items_length)?;
    let styles_offset = checked_end(positions_offset, positions_length)?;
    let final_offset = checked_end(styles_offset, styles_length)?;

    fib.set_glossary_document(true);
    fib.set_sttbf_glsy(table_offset, items_length);
    fib.set_plcf_glsy(positions_offset, positions_length);
    fib.set_sttb_glsy_style(styles_offset, styles_length);
    table_stream.extend_from_slice(&items);
    table_stream.extend_from_slice(&positions);
    table_stream.extend_from_slice(&styles);
    Ok(final_offset)
}
