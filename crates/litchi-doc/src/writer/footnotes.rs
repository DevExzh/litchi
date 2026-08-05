//! Footnotes and endnotes writer for DOC files
//!
//! Generates footnote/endnote reference PLCFs and subdocument content.

use super::WriteError;

fn add_cp(current: u32, text_units: usize, suffix: u32) -> Result<u32, WriteError> {
    let text_units = u32::try_from(text_units).map_err(|_| {
        WriteError::InvalidData("DOC footnote story exceeds the 32-bit CP range".to_string())
    })?;
    current
        .checked_add(text_units)
        .and_then(|next| next.checked_add(suffix))
        .ok_or_else(|| {
            WriteError::InvalidData("DOC footnote story exceeds the 32-bit CP range".to_string())
        })
}

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
    pub fn build_plcf_fnd_ref(&self) -> Result<Vec<u8>, WriteError> {
        let final_cp = match self
            .footnotes
            .iter()
            .map(|footnote| footnote.ref_position)
            .max()
        {
            Some(cp) => cp.checked_add(1).ok_or_else(|| {
                WriteError::InvalidData(
                    "DOC footnote reference exceeds the 32-bit CP range".to_string(),
                )
            })?,
            None => 0,
        };
        Ok(self.build_plcf_fnd_ref_with_text_length(final_cp))
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
    ///
    /// # Errors
    ///
    /// Returns [`WriteError::InvalidData`] when the UTF-16 footnote story length
    /// exceeds DOC's 32-bit character-position range or the PLCF cannot be
    /// allocated safely.
    pub fn build_plcf_fnd_txt(&self) -> Result<Vec<u8>, WriteError> {
        self.char_count()?;
        let final_mark = usize::from(!self.footnotes.is_empty());
        let cp_count = self
            .footnotes
            .len()
            .checked_add(1)
            .and_then(|count| count.checked_add(final_mark))
            .ok_or_else(|| WriteError::InvalidData("DOC footnote PLCF is too large".to_string()))?;
        let byte_count = cp_count
            .checked_mul(size_of::<u32>())
            .ok_or_else(|| WriteError::InvalidData("DOC footnote PLCF is too large".to_string()))?;
        let mut plcf = Vec::new();
        plcf.try_reserve_exact(byte_count).map_err(|_| {
            WriteError::InvalidData("DOC footnote PLCF allocation is too large".to_string())
        })?;
        let mut current_cp = 0u32;

        // Initial CP (start of first footnote)
        plcf.extend_from_slice(&current_cp.to_le_bytes());

        // Each range contains its automatic reference, body, and terminating paragraph mark.
        for footnote in self.ordered_footnotes() {
            current_cp = add_cp(current_cp, footnote.text.encode_utf16().count(), 2)?;
            plcf.extend_from_slice(&current_cp.to_le_bytes());
        }

        if !self.footnotes.is_empty() {
            current_cp = add_cp(current_cp, 0, 1)?;
            plcf.extend_from_slice(&current_cp.to_le_bytes());
        }

        Ok(plcf)
    }

    /// Get the subdocument text content
    pub fn build_subdocument_text(&self) -> Result<Vec<u8>, WriteError> {
        let char_count = self.char_count()?;
        let byte_count = usize::try_from(char_count)
            .ok()
            .and_then(|count| count.checked_mul(2))
            .ok_or_else(|| {
                WriteError::InvalidData(
                    "DOC footnote story byte length exceeds this platform".to_string(),
                )
            })?;
        let mut text_bytes = Vec::new();
        text_bytes.try_reserve_exact(byte_count).map_err(|_| {
            WriteError::InvalidData("DOC footnote story allocation is too large".to_string())
        })?;
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
        Ok(text_bytes)
    }

    /// Get total character count in footnote text.
    ///
    /// # Errors
    ///
    /// Returns [`WriteError::InvalidData`] when the UTF-16 story length exceeds
    /// DOC's 32-bit character-position range.
    pub fn char_count(&self) -> Result<u32, WriteError> {
        let stories = self.footnotes.iter().try_fold(0u32, |total, footnote| {
            add_cp(total, footnote.text.encode_utf16().count(), 2)
        })?;
        if self.footnotes.is_empty() {
            Ok(0)
        } else {
            add_cp(stories, 0, 1)
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

        assert_eq!(writer.char_count().unwrap(), 6);
        let text_plcf = writer.build_plcf_fnd_txt().unwrap();
        assert_eq!(u32::from_le_bytes(text_plcf[4..8].try_into().unwrap()), 5);
        assert_eq!(u32::from_le_bytes(text_plcf[8..12].try_into().unwrap()), 6);

        let text = writer.build_subdocument_text().unwrap();
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
    fn cp_sum_accepts_the_boundary_and_rejects_overflow() {
        assert_eq!(add_cp(u32::MAX - 2, 0, 2).unwrap(), u32::MAX);
        assert!(add_cp(u32::MAX - 1, 0, 2).is_err());
        assert!(add_cp(0, usize::MAX, 1).is_err());

        let mut writer = FootnotesWriter::new();
        writer.add_footnote(FootnoteEntry::new(u32::MAX, "note", 1));
        assert!(writer.build_plcf_fnd_ref().is_err());
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

        let text = writer.build_subdocument_text().unwrap();
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
