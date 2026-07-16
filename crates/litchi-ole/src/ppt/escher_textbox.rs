use super::super::consts::PptRecordType;
use super::package::{PptError, Result};
/// EscherTextboxWrapper implementation.
///
/// Based on Apache POI's EscherTextboxWrapper, this wraps an Escher textbox record
/// and provides access to its child PPT records (TextCharsAtom, TextBytesAtom, StyleTextPropAtom).
use super::records::PptRecord;
use super::text_run::{TextRun, TextRunExtractor};

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
    /// Text split into character-formatting runs
    runs: Vec<TextRun>,
}

impl EscherTextboxWrapper {
    /// Create a new wrapper from Escher textbox data.
    ///
    /// Based on POI's EscherTextboxWrapper constructor which calls
    /// Record.findChildRecords(data, 0, data.length).
    pub fn new(data: Vec<u8>) -> Result<Self> {
        // Parse child records from the escher data
        let child_records = Self::find_child_records(&data)?;

        let mut extractor = TextRunExtractor::new();
        extractor.extract_from_records(&child_records)?;
        let text = extractor.text().to_string();
        let runs = extractor.runs().to_vec();

        Ok(Self {
            data,
            child_records,
            text,
            runs,
        })
    }

    /// Find child PPT records in the Escher textbox data.
    ///
    /// Based on POI's Record.findChildRecords().
    fn find_child_records(data: &[u8]) -> Result<Vec<PptRecord>> {
        let mut records = Vec::new();
        let mut offset = 0;

        while offset < data.len() {
            let header_end = offset.checked_add(8).ok_or_else(|| {
                PptError::Corrupted("ClientTextbox record offset overflow".to_string())
            })?;
            if header_end > data.len() {
                return Err(PptError::Corrupted(
                    "Truncated record header in ClientTextbox".to_string(),
                ));
            }

            let payload_length = u32::from_le_bytes([
                data[offset + 4],
                data[offset + 5],
                data[offset + 6],
                data[offset + 7],
            ]);
            let payload_length = usize::try_from(payload_length).map_err(|_| {
                PptError::Corrupted("ClientTextbox record size overflow".to_string())
            })?;
            let record_end = header_end.checked_add(payload_length).ok_or_else(|| {
                PptError::Corrupted("ClientTextbox record size overflow".to_string())
            })?;
            if record_end > data.len() {
                return Err(PptError::Corrupted(
                    "Record extends beyond ClientTextbox data".to_string(),
                ));
            }

            let (record, consumed) = PptRecord::parse(&data[offset..record_end], 0)?;
            if consumed != record_end - offset {
                return Err(PptError::Corrupted(
                    "ClientTextbox child record was only partially parsed".to_string(),
                ));
            }
            records.push(record);
            offset = record_end;
        }

        Ok(records)
    }

    /// Get the extracted text.
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Get child records.
    pub fn child_records(&self) -> &[PptRecord] {
        &self.child_records
    }

    /// Get the character-formatting runs extracted from `StyleTextPropAtom`.
    pub fn runs(&self) -> &[TextRun] {
        &self.runs
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

    #[test]
    fn rejects_truncated_child_record_instead_of_resynchronizing() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&4008u16.to_le_bytes());
        data.extend_from_slice(&5u32.to_le_bytes());
        data.extend_from_slice(b"bad");

        let error = EscherTextboxWrapper::new(data).unwrap_err();
        assert!(error.to_string().contains("extends beyond ClientTextbox"));
    }

    #[test]
    fn allows_client_textboxes_without_inline_text() {
        let mut data = Vec::new();
        data.extend_from_slice(&0u16.to_le_bytes());
        data.extend_from_slice(&3999u16.to_le_bytes());
        data.extend_from_slice(&4u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());

        let wrapper = EscherTextboxWrapper::new(data).unwrap();
        assert_eq!(wrapper.text(), "");
        assert!(wrapper.runs().is_empty());
    }
}
