//! Headers and footers writer for DOC files
//!
//! Generates the PlcfHdd structure and header/footer subdocument content.

use super::WriteError;

fn add_cp(current: u32, text_units: usize, suffix: u32) -> Result<u32, WriteError> {
    let text_units = u32::try_from(text_units).map_err(|_| {
        WriteError::InvalidData("DOC header story exceeds the 32-bit CP range".to_string())
    })?;
    current
        .checked_add(text_units)
        .and_then(|next| next.checked_add(suffix))
        .ok_or_else(|| {
            WriteError::InvalidData("DOC header story exceeds the 32-bit CP range".to_string())
        })
}

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
    fn slots(&self) -> [Option<&str>; 12] {
        let mut slots: [Option<&str>; 12] = [None; 12];
        for entry in &self.entries {
            let index = match entry.hf_type {
                HeaderFooterType::EvenPageHeader => 6,
                HeaderFooterType::OddPageHeader => 7,
                HeaderFooterType::EvenPageFooter => 8,
                HeaderFooterType::OddPageFooter => 9,
                HeaderFooterType::FirstPageHeader => 10,
                HeaderFooterType::FirstPageFooter => 11,
            };
            slots[index] = Some(&entry.text);
        }
        slots
    }

    /// Create a new headers writer
    #[must_use]
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
    /// Returns UTF-16LE story bytes and all 14 `PlcfHdd` CPs for one section.
    ///
    /// # Errors
    ///
    /// Returns [`WriteError::InvalidData`] when the UTF-16 story length cannot
    /// be represented by DOC's 32-bit character positions or allocated safely.
    pub fn build_subdocument_text(&self) -> Result<(Vec<u8>, Vec<u32>), WriteError> {
        if self.entries.is_empty() {
            return Ok((Vec::new(), vec![0]));
        }

        let char_count = self.char_count()?;
        let byte_count = usize::try_from(char_count)
            .ok()
            .and_then(|count| count.checked_mul(2))
            .ok_or_else(|| {
                WriteError::InvalidData(
                    "DOC header story byte length exceeds this platform".to_string(),
                )
            })?;

        let mut text_bytes = Vec::new();
        text_bytes.try_reserve_exact(byte_count).map_err(|_| {
            WriteError::InvalidData("DOC header story allocation is too large".to_string())
        })?;
        let mut char_positions = Vec::with_capacity(14);
        let mut current_pos = 0u32;

        for text in self.slots() {
            char_positions.push(current_pos);
            if let Some(text) = text {
                let start = text_bytes.len();
                for unit in text.encode_utf16() {
                    text_bytes.extend_from_slice(&unit.to_le_bytes());
                }
                let story_units = text_bytes
                    .len()
                    .checked_sub(start)
                    .map(|bytes| bytes / size_of::<u16>())
                    .ok_or_else(|| {
                        WriteError::InvalidData(
                            "DOC header story byte length is invalid".to_string(),
                        )
                    })?;
                for unit in [0x000D_u16, 0x000D] {
                    text_bytes.extend_from_slice(&unit.to_le_bytes());
                }
                current_pos = add_cp(current_pos, story_units, 2)?;
            }
        }

        // The second-to-last CP terminates the last story at ccpHdd - 1. The header document has
        // one final paragraph mark whose position is recorded by the ignored last CP.
        char_positions.push(current_pos);
        text_bytes.extend_from_slice(&0x000Du16.to_le_bytes());
        current_pos = add_cp(current_pos, 0, 1)?;
        char_positions.push(current_pos);
        debug_assert_eq!(current_pos, char_count);

        Ok((text_bytes, char_positions))
    }

    /// Generate the `PlcfHdd` structure
    ///
    /// The `PlcfHdd` is a PLCF with `element_size=0` (just character positions)
    /// that maps character positions in the header subdocument
    ///
    /// # Errors
    ///
    /// Returns [`WriteError::InvalidData`] when the header story exceeds DOC's
    /// 32-bit character-position range or cannot be allocated safely.
    pub fn build_plcfhdd(&self) -> Result<Vec<u8>, WriteError> {
        let (_text, char_positions) = self.build_subdocument_text()?;

        let mut plcf = Vec::new();

        for &cp in &char_positions {
            plcf.extend_from_slice(&cp.to_le_bytes());
        }

        Ok(plcf)
    }

    /// Get character count for the header subdocument.
    ///
    /// # Errors
    ///
    /// Returns [`WriteError::InvalidData`] when the UTF-16 story length exceeds
    /// DOC's 32-bit character-position range.
    pub fn char_count(&self) -> Result<u32, WriteError> {
        if self.entries.is_empty() {
            return Ok(0);
        }
        self.slots()
            .into_iter()
            .flatten()
            .try_fold(1u32, |total, text| {
                add_cp(total, text.encode_utf16().count(), 2)
            })
    }

    /// Check if there are any headers/footers
    #[must_use]
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
        let types = [
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
        assert_eq!(writer.char_count().unwrap(), 0);
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
        assert_eq!(writer.char_count().unwrap(), 7);
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
        let (text_bytes, char_positions) = writer.build_subdocument_text().unwrap();
        assert!(text_bytes.is_empty());
        assert_eq!(char_positions, vec![0u32]);
    }

    #[test]
    fn test_build_subdocument_text_single() {
        let mut writer = HeadersWriter::new();
        writer.add_header(HeaderFooterType::OddPageHeader, "Hello");
        let (text_bytes, char_positions) = writer.build_subdocument_text().unwrap();
        let text = String::from_utf16(
            &text_bytes
                .chunks_exact(2)
                .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
                .collect::<Vec<_>>(),
        )
        .unwrap();
        assert_eq!(text, "Hello\r\r\r");
        assert_eq!(char_positions.len(), 14);
        assert_eq!(&char_positions[..8], &[0; 8]);
        assert_eq!(&char_positions[8..13], &[7; 5]);
        assert_eq!(char_positions[13], 8);
    }

    #[test]
    fn test_build_subdocument_text_multiple() {
        let mut writer = HeadersWriter::new();
        writer.add_header(HeaderFooterType::OddPageHeader, "First");
        writer.add_header(HeaderFooterType::EvenPageHeader, "Second");
        let (text_bytes, char_positions) = writer.build_subdocument_text().unwrap();
        let units = text_bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        assert_eq!(String::from_utf16(&units).unwrap(), "Second\r\rFirst\r\r\r");
        assert_eq!(char_positions.len(), 14);
        assert_eq!(char_positions[6], 0);
        assert_eq!(char_positions[7], 8);
        assert_eq!(char_positions[12], 15);
        assert_eq!(char_positions[13], 16);
    }

    #[test]
    fn test_build_plcfhdd_empty() {
        let writer = HeadersWriter::new();
        let plcf = writer.build_plcfhdd().unwrap();
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
        let plcf = writer.build_plcfhdd().unwrap();
        assert_eq!(plcf.len(), 56); // 14 CPs
        let cps = plcf
            .chunks_exact(4)
            .map(|chunk| u32::from_le_bytes(chunk.try_into().unwrap()))
            .collect::<Vec<_>>();
        assert_eq!(&cps[..8], &[0; 8]);
        assert_eq!(cps[8], 8);
        assert_eq!(cps[9], 8);
        assert_eq!(cps[10], 16);
        assert_eq!(cps[12], 16);
        assert_eq!(cps[13], 17);
    }

    #[test]
    fn test_char_count_multiple() {
        let mut writer = HeadersWriter::new();
        writer.add_header(HeaderFooterType::OddPageHeader, "Hello");
        writer.add_header(HeaderFooterType::EvenPageHeader, "World");
        assert_eq!(writer.char_count().unwrap(), 15);
    }

    #[test]
    fn supplementary_text_uses_utf16_cps_and_duplicate_empty_story_cps() {
        let mut writer = HeadersWriter::new();
        writer.add_header(HeaderFooterType::OddPageHeader, "😀");

        let (bytes, cps) = writer.build_subdocument_text().unwrap();
        assert_eq!(bytes.len(), 10); // surrogate pair + two story EOPs + document EOP
        assert_eq!(&cps[..8], &[0; 8]);
        assert_eq!(&cps[8..13], &[4; 5]);
        assert_eq!(cps[13], 5);
    }

    #[test]
    fn cp_sum_accepts_the_boundary_and_rejects_overflow() {
        assert_eq!(add_cp(u32::MAX - 2, 0, 2).unwrap(), u32::MAX);
        assert!(add_cp(u32::MAX - 1, 0, 2).is_err());
        assert!(add_cp(0, usize::MAX, 1).is_err());
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
