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
