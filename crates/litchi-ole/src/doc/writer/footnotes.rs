//! Footnotes and endnotes writer for DOC files
//!
//! Generates footnote/endnote reference PLCFs and subdocument content.

use std::io::Write;

/// A footnote entry
#[derive(Debug, Clone)]
pub struct FootnoteEntry {
    /// Reference position in main document (character position)
    pub ref_position: u32,
    /// Text content of the footnote
    pub text: String,
    /// Footnote number (1-based)
    pub number: u16,
}

impl FootnoteEntry {
    /// Create a new footnote
    pub fn new(ref_position: u32, text: impl Into<String>, number: u16) -> Self {
        Self {
            ref_position,
            text: text.into(),
            number,
        }
    }
}

/// Footnotes writer
#[derive(Debug)]
pub struct FootnotesWriter {
    footnotes: Vec<FootnoteEntry>,
}

impl FootnotesWriter {
    /// Create a new footnotes writer
    pub fn new() -> Self {
        Self {
            footnotes: Vec::new(),
        }
    }

    /// Add a footnote
    pub fn add_footnote(&mut self, footnote: FootnoteEntry) {
        self.footnotes.push(footnote);
    }

    /// Generate the footnote reference PLCF (PlcfFndRef)
    ///
    /// Format: CP array followed by FRD (Footnote Reference Descriptor) array
    /// FRD is 2 bytes: footnote number
    pub fn build_plcf_fnd_ref(&self) -> Vec<u8> {
        let mut plcf = Vec::new();

        // Write character positions
        for footnote in &self.footnotes {
            plcf.write_all(&footnote.ref_position.to_le_bytes())
                .unwrap();
        }
        // Write final CP
        let last_cp = self.footnotes.last().map_or(0, |f| f.ref_position + 1);
        plcf.write_all(&last_cp.to_le_bytes()).unwrap();

        // Write FRD descriptors (2 bytes each)
        for footnote in &self.footnotes {
            plcf.write_all(&footnote.number.to_le_bytes()).unwrap();
        }

        plcf
    }

    /// Generate footnote text PLCF (PlcfFndTxt)
    ///
    /// Maps character positions in the footnote subdocument
    pub fn build_plcf_fnd_txt(&self) -> Vec<u8> {
        let mut plcf = Vec::new();
        let mut current_cp = 0u32;

        // Initial CP (start of first footnote)
        plcf.write_all(&current_cp.to_le_bytes()).unwrap();

        // Each footnote contributes its character length + terminating chEop (0x0D)
        for footnote in &self.footnotes {
            let footnote_cp = footnote.text.chars().count() as u32 + 1; // include paragraph mark
            current_cp += footnote_cp;
            plcf.write_all(&current_cp.to_le_bytes()).unwrap();
        }

        plcf
    }

    /// Get the subdocument text content
    pub fn build_subdocument_text(&self) -> Vec<u8> {
        let mut text_bytes = Vec::new();
        for footnote in &self.footnotes {
            text_bytes.extend_from_slice(footnote.text.as_bytes());
        }
        text_bytes
    }

    /// Get total character count in footnote text
    pub fn char_count(&self) -> u32 {
        self.footnotes
            .iter()
            .map(|f| f.text.chars().count() as u32 + 1)
            .sum()
    }

    /// Get footnote entries
    pub fn footnotes(&self) -> &[FootnoteEntry] {
        &self.footnotes
    }

    /// Check if empty
    pub fn is_empty(&self) -> bool {
        self.footnotes.is_empty()
    }
}

impl Default for FootnotesWriter {
    fn default() -> Self {
        Self::new()
    }
}

/// Endnotes writer (same structure as footnotes)
pub type EndnotesWriter = FootnotesWriter;
