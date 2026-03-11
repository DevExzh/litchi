//! Headers and footers writer for DOC files
//!
//! Generates the PlcfHdd structure and header/footer subdocument content.

use std::io::Write;

/// Header/footer type
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderFooterType {
    /// First page header
    FirstPageHeader,
    /// Odd page header (default)
    OddPageHeader,
    /// Even page header
    EvenPageHeader,
    /// First page footer
    FirstPageFooter,
    /// Odd page footer (default)
    OddPageFooter,
    /// Even page footer
    EvenPageFooter,
}

/// A header or footer entry
#[derive(Debug, Clone)]
pub struct HeaderFooterEntry {
    /// Type of header/footer
    pub hf_type: HeaderFooterType,
    /// Text content
    pub text: String,
}

impl HeaderFooterEntry {
    /// Create a new header/footer entry
    pub fn new(hf_type: HeaderFooterType, text: impl Into<String>) -> Self {
        Self {
            hf_type,
            text: text.into(),
        }
    }
}

/// Headers and footers writer
#[derive(Debug)]
pub struct HeadersWriter {
    entries: Vec<HeaderFooterEntry>,
}

impl HeadersWriter {
    /// Create a new headers writer
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    /// Add a header or footer
    pub fn add_entry(&mut self, entry: HeaderFooterEntry) {
        self.entries.push(entry);
    }

    /// Add a header
    pub fn add_header(&mut self, hf_type: HeaderFooterType, text: impl Into<String>) {
        self.add_entry(HeaderFooterEntry::new(hf_type, text));
    }

    /// Get the subdocument text content
    ///
    /// Returns the concatenated text of all headers/footers and the character positions
    pub fn build_subdocument_text(&self) -> (Vec<u8>, Vec<u32>) {
        let mut text_bytes = Vec::new();
        let mut char_positions = vec![0u32];
        let mut current_pos = 0u32;

        for entry in &self.entries {
            // Convert text to CP1252 bytes
            let text = entry.text.as_bytes();
            text_bytes.extend_from_slice(text);
            current_pos += text.len() as u32;
            char_positions.push(current_pos);
        }

        (text_bytes, char_positions)
    }

    /// Generate the PlcfHdd structure
    ///
    /// The PlcfHdd is a PLCF with element_size=0 (just character positions)
    /// that maps character positions in the header subdocument
    pub fn build_plcfhdd(&self) -> Vec<u8> {
        let (_text, char_positions) = self.build_subdocument_text();

        let mut plcf = Vec::new();

        // Write character positions
        for &cp in &char_positions {
            plcf.write_all(&cp.to_le_bytes()).unwrap();
        }

        plcf
    }

    /// Get character count for the header subdocument
    pub fn char_count(&self) -> u32 {
        self.entries.iter().map(|e| e.text.len() as u32).sum()
    }

    /// Check if there are any headers/footers
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }
}

impl Default for HeadersWriter {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_header_footer_type_variants() {
        // Test all HeaderFooterType variants exist and are distinct
        let types = vec![
            HeaderFooterType::FirstPageHeader,
            HeaderFooterType::OddPageHeader,
            HeaderFooterType::EvenPageHeader,
            HeaderFooterType::FirstPageFooter,
            HeaderFooterType::OddPageFooter,
            HeaderFooterType::EvenPageFooter,
        ];
        // All 6 types should be distinct
        assert_eq!(types.len(), 6);
    }

    #[test]
    fn test_header_footer_entry_new() {
        let entry = HeaderFooterEntry::new(HeaderFooterType::OddPageHeader, "Header text");
        assert_eq!(entry.hf_type, HeaderFooterType::OddPageHeader);
        assert_eq!(entry.text, "Header text");
    }

    #[test]
    fn test_header_footer_entry_with_string() {
        let entry =
            HeaderFooterEntry::new(HeaderFooterType::FirstPageFooter, "Footer text".to_string());
        assert_eq!(entry.hf_type, HeaderFooterType::FirstPageFooter);
        assert_eq!(entry.text, "Footer text");
    }

    #[test]
    fn test_headers_writer_new() {
        let writer = HeadersWriter::new();
        assert!(writer.is_empty());
        assert_eq!(writer.char_count(), 0);
    }

    #[test]
    fn test_headers_writer_default() {
        let writer: HeadersWriter = Default::default();
        assert!(writer.is_empty());
    }

    #[test]
    fn test_add_entry() {
        let mut writer = HeadersWriter::new();
        let entry = HeaderFooterEntry::new(HeaderFooterType::OddPageHeader, "Test");
        writer.add_entry(entry);
        assert!(!writer.is_empty());
        assert_eq!(writer.char_count(), 4);
    }

    #[test]
    fn test_add_header() {
        let mut writer = HeadersWriter::new();
        writer.add_header(HeaderFooterType::OddPageHeader, "Odd Header");
        writer.add_header(HeaderFooterType::EvenPageHeader, "Even Header");
        assert_eq!(writer.entries.len(), 2);
    }

    #[test]
    fn test_build_subdocument_text_empty() {
        let writer = HeadersWriter::new();
        let (text_bytes, char_positions) = writer.build_subdocument_text();
        assert!(text_bytes.is_empty());
        assert_eq!(char_positions, vec![0u32]);
    }

    #[test]
    fn test_build_subdocument_text_single() {
        let mut writer = HeadersWriter::new();
        writer.add_header(HeaderFooterType::OddPageHeader, "Hello");
        let (text_bytes, char_positions) = writer.build_subdocument_text();
        assert_eq!(text_bytes, b"Hello");
        assert_eq!(char_positions, vec![0u32, 5u32]);
    }

    #[test]
    fn test_build_subdocument_text_multiple() {
        let mut writer = HeadersWriter::new();
        writer.add_header(HeaderFooterType::OddPageHeader, "First");
        writer.add_header(HeaderFooterType::EvenPageHeader, "Second");
        let (text_bytes, char_positions) = writer.build_subdocument_text();
        assert_eq!(text_bytes, b"FirstSecond");
        assert_eq!(char_positions, vec![0u32, 5u32, 11u32]);
    }

    #[test]
    fn test_build_plcfhdd_empty() {
        let writer = HeadersWriter::new();
        let plcf = writer.build_plcfhdd();
        // Should contain just one character position (0)
        assert_eq!(plcf.len(), 4);
        assert_eq!(
            u32::from_le_bytes([plcf[0], plcf[1], plcf[2], plcf[3]]),
            0u32
        );
    }

    #[test]
    fn test_build_plcfhdd_with_entries() {
        let mut writer = HeadersWriter::new();
        writer.add_header(HeaderFooterType::OddPageHeader, "Header");
        writer.add_header(HeaderFooterType::OddPageFooter, "Footer");
        let plcf = writer.build_plcfhdd();
        // Should contain 3 character positions (0, 6, 12)
        assert_eq!(plcf.len(), 12); // 3 * 4 bytes

        let cp0 = u32::from_le_bytes([plcf[0], plcf[1], plcf[2], plcf[3]]);
        let cp1 = u32::from_le_bytes([plcf[4], plcf[5], plcf[6], plcf[7]]);
        let cp2 = u32::from_le_bytes([plcf[8], plcf[9], plcf[10], plcf[11]]);

        assert_eq!(cp0, 0u32);
        assert_eq!(cp1, 6u32);
        assert_eq!(cp2, 12u32);
    }

    #[test]
    fn test_char_count_multiple() {
        let mut writer = HeadersWriter::new();
        writer.add_header(HeaderFooterType::OddPageHeader, "Hello");
        writer.add_header(HeaderFooterType::EvenPageHeader, "World");
        assert_eq!(writer.char_count(), 10);
    }

    #[test]
    fn test_is_empty_with_entries() {
        let mut writer = HeadersWriter::new();
        assert!(writer.is_empty());
        writer.add_header(HeaderFooterType::OddPageHeader, "Test");
        assert!(!writer.is_empty());
    }

    #[test]
    fn test_header_footer_entry_clone() {
        let entry = HeaderFooterEntry::new(HeaderFooterType::OddPageHeader, "Test");
        let cloned = entry.clone();
        assert_eq!(entry.hf_type, cloned.hf_type);
        assert_eq!(entry.text, cloned.text);
    }

    #[test]
    fn test_header_footer_entry_debug() {
        let entry = HeaderFooterEntry::new(HeaderFooterType::OddPageHeader, "Test");
        let debug_str = format!("{:?}", entry);
        assert!(debug_str.contains("HeaderFooterEntry"));
    }

    #[test]
    fn test_headers_writer_debug() {
        let writer = HeadersWriter::new();
        let debug_str = format!("{:?}", writer);
        assert!(debug_str.contains("HeadersWriter"));
    }

    #[test]
    fn test_all_header_footer_types() {
        let mut writer = HeadersWriter::new();

        writer.add_header(HeaderFooterType::FirstPageHeader, "First Header");
        writer.add_header(HeaderFooterType::OddPageHeader, "Odd Header");
        writer.add_header(HeaderFooterType::EvenPageHeader, "Even Header");
        writer.add_header(HeaderFooterType::FirstPageFooter, "First Footer");
        writer.add_header(HeaderFooterType::OddPageFooter, "Odd Footer");
        writer.add_header(HeaderFooterType::EvenPageFooter, "Even Footer");

        assert_eq!(writer.entries.len(), 6);

        // Verify all types are present
        assert!(
            writer
                .entries
                .iter()
                .any(|e| matches!(e.hf_type, HeaderFooterType::FirstPageHeader))
        );
        assert!(
            writer
                .entries
                .iter()
                .any(|e| matches!(e.hf_type, HeaderFooterType::OddPageHeader))
        );
        assert!(
            writer
                .entries
                .iter()
                .any(|e| matches!(e.hf_type, HeaderFooterType::EvenPageHeader))
        );
        assert!(
            writer
                .entries
                .iter()
                .any(|e| matches!(e.hf_type, HeaderFooterType::FirstPageFooter))
        );
        assert!(
            writer
                .entries
                .iter()
                .any(|e| matches!(e.hf_type, HeaderFooterType::OddPageFooter))
        );
        assert!(
            writer
                .entries
                .iter()
                .any(|e| matches!(e.hf_type, HeaderFooterType::EvenPageFooter))
        );
    }
}
