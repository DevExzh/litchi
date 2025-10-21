//! Text extraction from Escher records.
//!
//! # Architecture
//!
//! - Extracts text from ClientTextbox records
//! - Parses embedded PPT text records (TextCharsAtom, TextBytesAtom)
//! - Zero-copy where possible

use super::container::EscherContainer;
use super::record::EscherRecord;
use super::types::EscherRecordType;
use crate::ole::consts::PptRecordType;
use crate::ole::ppt::package::Result;
use crate::ole::ppt::records::PptRecord;

/// Extract all text from an Escher record hierarchy.
///
/// # Performance
///
/// - Depth-first traversal
/// - Short-circuits on first text found per container
/// - Pre-allocated string buffers
pub fn extract_text_from_escher(escher_data: &[u8]) -> Result<String> {
    let mut text_parts = Vec::new();

    // Parse the Escher structure
    let parser = super::parser::EscherParser::new(escher_data);

    // Get root container
    if let Some(root_result) = parser.root_container() {
        let root = root_result?;
        extract_text_from_container(&root, &mut text_parts);
    }

    Ok(if text_parts.is_empty() {
        String::new()
    } else {
        text_parts.join("\n")
    })
}

/// Recursively extract text from a container and its children.
fn extract_text_from_container(container: &EscherContainer, text_parts: &mut Vec<String>) {
    for child in container.children().flatten() {
        match child.record_type {
            // ClientTextbox contains embedded PPT text records
            EscherRecordType::ClientTextbox => {
                if let Some(text) = extract_text_from_textbox(&child)
                    && !text.trim().is_empty()
                {
                    text_parts.push(text);
                }
            },
            // SpContainer may contain shapes with text
            EscherRecordType::SpContainer => {
                let sp_container = EscherContainer::new(child);
                extract_text_from_container(&sp_container, text_parts);
            },
            // Other container types - recurse
            _ if child.is_container() => {
                let child_container = EscherContainer::new(child);
                extract_text_from_container(&child_container, text_parts);
            },
            _ => {},
        }
    }
}

/// Extract text from a ClientTextbox record.
///
/// ClientTextbox contains embedded PPT records (TextCharsAtom, TextBytesAtom, etc.)
pub(crate) fn extract_text_from_textbox(textbox: &EscherRecord) -> Option<String> {
    if textbox.data.is_empty() {
        return None;
    }

    // Try to parse embedded PPT records
    let mut offset = 0;
    let mut text_parts = Vec::new();

    while offset + 8 <= textbox.data.len() {
        match PptRecord::parse(textbox.data, offset) {
            Ok((record, consumed)) => {
                // Extract text from this record
                if let Ok(record_text) = record.extract_text() {
                    let trimmed = record_text.trim();
                    if !trimmed.is_empty() {
                        text_parts.push(trimmed.to_string());
                    }
                }

                // Also check children recursively
                extract_text_from_ppt_record(&record, &mut text_parts);

                offset += consumed;
                if consumed == 0 {
                    break;
                }
            },
            Err(_) => {
                // Move forward to try next position
                offset += 1;
            },
        }
    }

    if text_parts.is_empty() {
        None
    } else {
        Some(text_parts.join("\n"))
    }
}

/// Recursively extract text from PPT record and children.
fn extract_text_from_ppt_record(record: &PptRecord, text_parts: &mut Vec<String>) {
    // Check specific text record types
    match record.record_type {
        PptRecordType::TextCharsAtom | PptRecordType::TextBytesAtom | PptRecordType::CString => {
            if let Ok(text) = record.extract_text() {
                let trimmed = text.trim();
                if !trimmed.is_empty() {
                    text_parts.push(trimmed.to_string());
                }
            }
        },
        _ => {},
    }

    // Recursively process children
    for child in &record.children {
        extract_text_from_ppt_record(child, text_parts);
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_empty_escher_data() {
        let data = vec![];
        let text = extract_text_from_escher(&data).unwrap();
        assert_eq!(text, "");
    }

    #[test]
    fn test_escher_without_text() {
        // DgContainer with no ClientTextbox
        let data = vec![
            0x0F, 0x00, // version=0xF (container), instance=0
            0x02, 0xF0, // record type = 0xF002 (DgContainer)
            0x04, 0x00, 0x00, 0x00, // length = 4
            0x01, 0x02, 0x03, 0x04, // data
        ];

        let text = extract_text_from_escher(&data).unwrap();
        assert_eq!(text, "");
    }
}
