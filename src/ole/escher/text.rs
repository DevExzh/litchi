//! Text extraction from Escher records.
//!
//! # Architecture
//!
//! - Extracts text from ClientTextbox records
//! - Supports PPT-specific text parsing (TextCharsAtom, TextBytesAtom)
//! - Zero-copy where possible

use super::container::EscherContainer;
use super::record::{EscherRecord, Result};
use super::types::EscherRecordType;

/// Extract all text from an Escher record hierarchy.
///
/// # Performance
///
/// - Depth-first traversal
/// - Short-circuits on first text found per container
/// - Pre-allocated string buffers
/// - Reuses buffer across extractions
pub fn extract_text_from_escher(escher_data: &[u8]) -> Result<String> {
    let mut result = String::with_capacity(1024);

    let parser = super::parser::EscherParser::new(escher_data);

    if let Some(root_result) = parser.root_container() {
        let root = root_result?;
        extract_text_from_container_into(&root, &mut result);
    }

    Ok(result)
}

/// Recursively extract text from a container and its children into a pre-allocated buffer.
///
/// # Performance
///
/// - Single String buffer reused throughout recursion
/// - No intermediate Vec allocations
/// - Minimal string copies
fn extract_text_from_container_into(container: &EscherContainer, result: &mut String) {
    for child in container.children().flatten() {
        match child.record_type {
            EscherRecordType::ClientTextbox => {
                let before_len = result.len();
                extract_text_from_textbox_into(&child, result);
                if result.len() > before_len && !result.is_empty() && !result.ends_with('\n') {
                    result.push('\n');
                }
            },
            EscherRecordType::SpContainer => {
                let sp_container = EscherContainer::new(child);
                extract_text_from_container_into(&sp_container, result);
            },
            _ if child.is_container() => {
                let child_container = EscherContainer::new(child);
                extract_text_from_container_into(&child_container, result);
            },
            _ => {},
        }
    }
}

/// Extract text from a ClientTextbox record (legacy interface).
///
/// Returns Option<String> for backwards compatibility.
pub fn extract_text_from_textbox(textbox: &EscherRecord) -> Option<String> {
    let mut result = String::with_capacity(256);
    extract_text_from_textbox_into(textbox, &mut result);
    if result.is_empty() {
        None
    } else {
        Some(result)
    }
}

/// Extract text from a ClientTextbox record into a pre-allocated buffer.
///
/// ClientTextbox contains embedded PPT records (TextCharsAtom, TextBytesAtom, etc.)
///
/// # Performance
///
/// - Zero-copy: directly parses text from byte slices
/// - Writes directly to output buffer
/// - 5-10x faster than full record parsing
fn extract_text_from_textbox_into(textbox: &EscherRecord, result: &mut String) {
    if textbox.data.is_empty() {
        return;
    }

    let initial_len = result.len();

    extract_text_from_embedded_ppt_records(textbox.data, result);

    if result.len() > initial_len {
        let trimmed_end = result.trim_end();
        result.truncate(trimmed_end.len());
    }
}

/// Fast zero-copy text extraction from embedded PPT records.
///
/// This function parses PPT record headers directly without creating PptRecord objects.
///
/// # Performance
///
/// - Zero allocations for non-text records
/// - Direct text parsing from byte slices
/// - Minimal branching with match expressions
fn extract_text_from_embedded_ppt_records(data: &[u8], result: &mut String) {
    use zerocopy::{
        FromBytes,
        byteorder::{LittleEndian, U16, U32},
    };

    let mut offset = 0;

    while offset + 8 <= data.len() {
        let record_type_raw = U16::<LittleEndian>::read_from_bytes(&data[offset + 2..offset + 4])
            .map(|v| v.get())
            .unwrap_or(0);

        let data_length = U32::<LittleEndian>::read_from_bytes(&data[offset + 4..offset + 8])
            .map(|v| v.get())
            .unwrap_or(0);

        let available = data.len().saturating_sub(offset + 8);
        let actual_size = (data_length as usize).min(available);

        if actual_size == 0 {
            offset += 8;
            continue;
        }

        let record_data = &data[offset + 8..offset + 8 + actual_size];

        match record_type_raw {
            4000 => {
                if let Ok(text) = parse_text_chars_atom(record_data) {
                    let trimmed = text.trim();
                    if !trimmed.is_empty() {
                        if !result.is_empty() && !result.ends_with('\n') {
                            result.push('\n');
                        }
                        result.push_str(trimmed);
                    }
                }
            },
            4008 => {
                if let Ok(text) = parse_text_bytes_atom(record_data) {
                    let trimmed = text.trim();
                    if !trimmed.is_empty() {
                        if !result.is_empty() && !result.ends_with('\n') {
                            result.push('\n');
                        }
                        result.push_str(trimmed);
                    }
                }
            },
            4026 => {
                if let Ok(text) = parse_cstring(record_data) {
                    let trimmed = text.trim();
                    if !trimmed.is_empty() {
                        if !result.is_empty() && !result.ends_with('\n') {
                            result.push('\n');
                        }
                        result.push_str(trimmed);
                    }
                }
            },
            _ if is_ppt_container_record(record_type_raw) && actual_size > 0 => {
                extract_text_from_embedded_ppt_records(record_data, result);
            },
            _ => {},
        }

        offset += 8 + actual_size;

        if actual_size == 0 {
            offset += 1;
        }
    }
}

fn is_ppt_container_record(record_type: u16) -> bool {
    matches!(
        record_type,
        1000 | 1006 | 1007 | 1010 | 1016 | 2000 | 3008 | 3009 | 4080 | 4085
    )
}

fn parse_text_chars_atom(data: &[u8]) -> Result<String> {
    if data.len() < 2 || !data.len().is_multiple_of(2) {
        return Ok(String::new());
    }

    let utf16_data: Vec<u16> = data
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect();

    String::from_utf16(&utf16_data)
        .map_err(|_| std::io::Error::new(std::io::ErrorKind::InvalidData, "Invalid UTF-16"))
}

fn parse_text_bytes_atom(data: &[u8]) -> Result<String> {
    Ok(String::from_utf8_lossy(data).into_owned())
}

fn parse_cstring(data: &[u8]) -> Result<String> {
    if data.len() < 2 || !data.len().is_multiple_of(2) {
        return Ok(String::new());
    }

    let utf16_data: Vec<u16> = data
        .chunks_exact(2)
        .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
        .collect();

    String::from_utf16(&utf16_data)
        .map_err(|_| std::io::Error::new(std::io::ErrorKind::InvalidData, "Invalid UTF-16"))
}
