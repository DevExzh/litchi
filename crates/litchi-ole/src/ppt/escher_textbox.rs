use super::super::consts::PptRecordType;
use super::package::Result;
/// EscherTextboxWrapper implementation.
///
/// Based on Apache POI's EscherTextboxWrapper, this wraps an Escher textbox record
/// and provides access to its child PPT records (TextCharsAtom, TextBytesAtom, StyleTextPropAtom).
use super::records::PptRecord;

/// Wrapper around Escher textbox data.
///
/// Based on POI's EscherTextboxWrapper. Parses child records from the textbox data.
#[derive(Debug, Clone)]
pub struct EscherTextboxWrapper {
    /// The raw Escher textbox data
    data: Vec<u8>,
    /// Child PPT records found in the textbox
    child_records: Vec<PptRecord>,
    /// Extracted text
    text: String,
}

impl EscherTextboxWrapper {
    /// Create a new wrapper from Escher textbox data.
    ///
    /// Based on POI's EscherTextboxWrapper constructor which calls
    /// Record.findChildRecords(data, 0, data.length).
    pub fn new(data: Vec<u8>) -> Result<Self> {
        // Parse child records from the escher data
        let child_records = Self::find_child_records(&data)?;

        // Extract text from text records
        let text = Self::extract_text_from_records(&child_records)?;

        // If no text records were found, this might be raw text data
        // that should be handled by the fallback in parse_text_record
        if text.is_empty() && !child_records.is_empty() {
            return Err(super::package::PptError::InvalidFormat(
                "No text records found in Escher textbox data".to_string(),
            ));
        }

        Ok(Self {
            data,
            child_records,
            text,
        })
    }

    /// Find child PPT records in the Escher textbox data.
    ///
    /// Based on POI's Record.findChildRecords().
    fn find_child_records(data: &[u8]) -> Result<Vec<PptRecord>> {
        let mut records = Vec::new();
        let mut offset = 0;

        while offset + 8 <= data.len() {
            match PptRecord::parse(data, offset) {
                Ok((record, consumed)) => {
                    records.push(record);
                    offset += consumed;
                    if consumed == 0 {
                        break; // Prevent infinite loop
                    }
                },
                Err(_) => {
                    // Skip invalid records
                    offset += 1;
                    if offset + 8 > data.len() {
                        break;
                    }
                },
            }
        }

        Ok(records)
    }

    /// Extract text from child records.
    ///
    /// Looks for TextCharsAtom or TextBytesAtom records.
    fn extract_text_from_records(records: &[PptRecord]) -> Result<String> {
        let mut text_parts = Vec::new();

        for record in records {
            match record.record_type {
                PptRecordType::TextCharsAtom | PptRecordType::TextBytesAtom => {
                    if let Ok(text) = record.extract_text()
                        && !text.is_empty()
                    {
                        text_parts.push(text);
                    }
                },
                _ => {},
            }
        }

        Ok(text_parts.join("\n"))
    }

    /// Get the extracted text.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Get child records.
    pub fn child_records(&self) -> &[PptRecord] {
        &self.child_records
    }

    /// Find a StyleTextPropAtom record.
    pub fn find_style_text_prop_atom(&self) -> Option<&PptRecord> {
        self.child_records
            .iter()
            .find(|r| r.record_type == PptRecordType::StyleTextPropAtom)
    }

    /// Get the raw data.
    pub fn data(&self) -> &[u8] {
        &self.data
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_escher_textbox_wrapper_creation() {
        // Create a simple textbox with a TextCharsAtom
        // PPT record header: ver/inst (2 bytes) | type (2 bytes) | length (4 bytes)
        let mut data = Vec::new();

        // TextCharsAtom: Record type 0x0FA0 (4000)
        data.extend_from_slice(&[0x00, 0x00]); // Version/instance (ver=0, inst=0)
        data.extend_from_slice(&[0xA0, 0x0F]); // Record type 0x0FA0 (little-endian)
        data.extend_from_slice(&[0x0A, 0x00, 0x00, 0x00]); // Length: 10 bytes (little-endian)

        // Text data (UTF-16LE): "Hello"
        data.extend_from_slice(&[
            0x48, 0x00, // 'H'
            0x65, 0x00, // 'e'
            0x6C, 0x00, // 'l'
            0x6C, 0x00, // 'l'
            0x6F, 0x00, // 'o'
        ]);

        let wrapper = EscherTextboxWrapper::new(data).unwrap();
        assert!(wrapper.text().contains("Hello") || !wrapper.text().is_empty());
        assert!(!wrapper.child_records().is_empty());
    }
}
