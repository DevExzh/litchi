//! Footnotes and endnotes writer for DOC files
//!
//! Generates footnote/endnote reference PLCFs and subdocument content.

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
    fn ordered_footnotes(&self) -> Vec<&FootnoteEntry> {
        let mut ordered = self.footnotes.iter().collect::<Vec<_>>();
        ordered.sort_by_key(|footnote| footnote.ref_position);
        ordered
    }

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
        let final_cp = self
            .footnotes
            .iter()
            .map(|footnote| footnote.ref_position)
            .max()
            .and_then(|cp| cp.checked_add(1))
            .unwrap_or(0);
        self.build_plcf_fnd_ref_with_text_length(final_cp)
    }

    /// Generate `PlcfFndRef` with the main-document character count as its ignored final CP.
    pub fn build_plcf_fnd_ref_with_text_length(&self, ccp_text: u32) -> Vec<u8> {
        let mut plcf = Vec::new();

        for footnote in self.ordered_footnotes() {
            plcf.extend_from_slice(&footnote.ref_position.to_le_bytes());
        }
        plcf.extend_from_slice(&ccp_text.to_le_bytes());

        for footnote in self.ordered_footnotes() {
            plcf.extend_from_slice(&footnote.number.max(1).to_le_bytes());
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
        plcf.extend_from_slice(&current_cp.to_le_bytes());

        // Each range contains its automatic reference, body, and terminating paragraph mark.
        for footnote in self.ordered_footnotes() {
            let footnote_cp = u32::try_from(footnote.text.encode_utf16().count())
                .expect("DOC footnote exceeds the 32-bit CP range")
                .checked_add(2)
                .expect("DOC footnote exceeds the 32-bit CP range");
            current_cp = current_cp
                .checked_add(footnote_cp)
                .expect("DOC footnote story exceeds the 32-bit CP range");
            plcf.extend_from_slice(&current_cp.to_le_bytes());
        }

        if !self.footnotes.is_empty() {
            current_cp = current_cp
                .checked_add(1)
                .expect("DOC footnote story exceeds the 32-bit CP range");
            plcf.extend_from_slice(&current_cp.to_le_bytes());
        }

        plcf
    }

    /// Get the subdocument text content
    pub fn build_subdocument_text(&self) -> Vec<u8> {
        let mut text_bytes = Vec::new();
        for footnote in self.ordered_footnotes() {
            text_bytes.extend_from_slice(&0x0002u16.to_le_bytes());
            for unit in footnote.text.encode_utf16() {
                text_bytes.extend_from_slice(&unit.to_le_bytes());
            }
            text_bytes.extend_from_slice(&0x000Du16.to_le_bytes());
        }
        if !self.footnotes.is_empty() {
            text_bytes.extend_from_slice(&0x000Du16.to_le_bytes());
        }
        text_bytes
    }

    /// Get total character count in footnote text
    pub fn char_count(&self) -> u32 {
        let stories = self
            .footnotes
            .iter()
            .map(|footnote| {
                u32::try_from(footnote.text.encode_utf16().count())
                    .expect("DOC footnote exceeds the 32-bit CP range")
                    .checked_add(2)
                    .expect("DOC footnote exceeds the 32-bit CP range")
            })
            .try_fold(0u32, |total, length| total.checked_add(length))
            .expect("DOC footnote story exceeds the 32-bit CP range");
        if self.footnotes.is_empty() {
            0
        } else {
            stories
                .checked_add(1)
                .expect("DOC footnote story exceeds the 32-bit CP range")
        }
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn counts_supplementary_characters_as_two_utf16_code_units() {
        let mut writer = FootnotesWriter::new();
        writer.add_footnote(FootnoteEntry::new(0, "A😀", 1));

        assert_eq!(writer.char_count(), 6);
        let text_plcf = writer.build_plcf_fnd_txt();
        assert_eq!(u32::from_le_bytes(text_plcf[4..8].try_into().unwrap()), 5);
        assert_eq!(u32::from_le_bytes(text_plcf[8..12].try_into().unwrap()), 6);

        let text = writer.build_subdocument_text();
        let units = text
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        assert_eq!(String::from_utf16(&units).unwrap(), "\u{2}A😀\r\r");

        let references = writer.build_plcf_fnd_ref_with_text_length(20);
        assert_eq!(u32::from_le_bytes(references[4..8].try_into().unwrap()), 20);
        assert_eq!(u16::from_le_bytes(references[8..10].try_into().unwrap()), 1);
    }

    #[test]
    fn orders_references_and_text_stories_together() {
        let mut writer = FootnotesWriter::new();
        writer.add_footnote(FootnoteEntry::new(8, "second", 2));
        writer.add_footnote(FootnoteEntry::new(2, "first", 1));

        let references = writer.build_plcf_fnd_ref_with_text_length(20);
        assert_eq!(u32::from_le_bytes(references[0..4].try_into().unwrap()), 2);
        assert_eq!(u32::from_le_bytes(references[4..8].try_into().unwrap()), 8);
        assert_eq!(
            u16::from_le_bytes(references[12..14].try_into().unwrap()),
            1
        );
        assert_eq!(
            u16::from_le_bytes(references[14..16].try_into().unwrap()),
            2
        );

        let text = writer.build_subdocument_text();
        let units = text
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]]))
            .collect::<Vec<_>>();
        assert_eq!(
            String::from_utf16(&units).unwrap(),
            "\u{2}first\r\u{2}second\r\r"
        );
    }
}
