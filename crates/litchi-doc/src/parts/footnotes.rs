/// Footnotes and endnotes parser for Word binary format.
///
/// Based on Apache POI's `FootnotesTables` and `LibreOffice`'s implementation.
/// Footnotes and endnotes are stored in separate subdocuments with references in the main text.
use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;
use crate::plcf::Plcf;

/// Reference descriptor for footnote/endnote (FRD structure - 2 bytes)
#[derive(Debug, Clone, Copy)]
pub struct FootnoteDescriptor {
    /// Reference number (auto-numbered or custom)
    pub number: u16,
}

impl FootnoteDescriptor {
    /// Parse a footnote descriptor from 2 bytes
    #[must_use]
    pub fn from_bytes(data: &[u8]) -> Option<Self> {
        if data.len() < 2 {
            return None;
        }

        // FRD structure (2 bytes):
        // Bit 0-15: nAuto (auto-number or custom number mark)
        let number = litchi_core::binary::read_u16_le(data, 0).ok()?;

        Some(Self { number })
    }
}

/// A footnote or endnote reference in the main document
#[derive(Debug, Clone)]
pub struct FootnoteReference {
    /// Character position of the reference in the main document
    pub ref_cp: u32,
    /// Character position range in the footnote/endnote subdocument
    pub text_start_cp: u32,
    pub text_end_cp: u32,
    /// Reference descriptor
    pub descriptor: FootnoteDescriptor,
}

impl FootnoteReference {
    /// Create a new footnote reference
    #[must_use]
    pub fn new(
        ref_cp: u32,
        text_start_cp: u32,
        text_end_cp: u32,
        descriptor: FootnoteDescriptor,
    ) -> Self {
        Self {
            ref_cp,
            text_start_cp,
            text_end_cp,
            descriptor,
        }
    }

    /// Get the length of the footnote/endnote text
    #[must_use]
    pub fn text_length(&self) -> u32 {
        self.text_end_cp.saturating_sub(self.text_start_cp)
    }
}

/// Footnotes table parser
pub struct FootnotesTable {
    /// All footnote references
    references: Vec<FootnoteReference>,
}

impl FootnotesTable {
    /// Parse footnotes from the FIB and table stream
    ///
    /// # Arguments
    ///
    /// * `fib` - File Information Block
    /// * `table_stream` - The table stream (0Table or 1Table)
    ///
    /// # Returns
    ///
    /// A parsed `FootnotesTable`
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let mut references = Vec::new();

        // Check if footnotes exist in the document
        if let Some((subdoc_start, subdoc_end)) = fib.get_footnote_range() {
            // Parse footnote reference PLCF (plcfFndRef)
            // FIB index 2: fcPlcfFndRef and lcbPlcfFndRef
            if let Some((offset, length)) = fib.get_table_pointer(2)
                && length > 0
                && (offset as usize) < table_stream.len()
            {
                let plcf_data = &table_stream[offset as usize..];
                let plcf_len = length.min((table_stream.len() - offset as usize) as u32) as usize;

                if plcf_len >= 4 {
                    // Parse reference PLCF with 2-byte FRD descriptors
                    if let Some(ref_plcf) = Plcf::parse(&plcf_data[..plcf_len], 2) {
                        // Parse footnote text PLCF (plcfFndTxt)
                        // FIB index 3: fcPlcfFndTxt and lcbPlcfFndTxt
                        if let Some((txt_offset, txt_length)) = fib.get_table_pointer(3)
                            && txt_length > 0
                            && (txt_offset as usize) < table_stream.len()
                        {
                            let txt_plcf_data = &table_stream[txt_offset as usize..];
                            let txt_plcf_len = txt_length
                                .min((table_stream.len() - txt_offset as usize) as u32)
                                as usize;

                            references = Self::parse_footnote_plcfs(
                                &ref_plcf,
                                &txt_plcf_data[..txt_plcf_len],
                                subdoc_start,
                                subdoc_end,
                            )?;
                        }
                    }
                }
            }
        }

        Ok(Self { references })
    }

    /// Parse footnote PLCF structures
    fn parse_footnote_plcfs(
        ref_plcf: &Plcf<'_>,
        txt_plcf_data: &[u8],
        subdoc_start: u32,
        subdoc_end: u32,
    ) -> Result<Vec<FootnoteReference>> {
        // Parse text PLCF with element_size = 0 (just CPs)
        // Manually parse since `Plcf` expects element_size > 0.
        if !txt_plcf_data.len().is_multiple_of(4) {
            return Err(PackageError::Corrupted(
                "note text PLCF contains a partial CP".to_string(),
            ));
        }
        let cp_count = txt_plcf_data.len() / 4;
        if cp_count != ref_plcf.count() + 2 {
            return Err(PackageError::Corrupted(
                "note text PLCF count does not match its reference PLCF".to_string(),
            ));
        }

        let mut text_cps = Vec::with_capacity(cp_count);
        for i in 0..cp_count {
            if let Ok(cp) = litchi_core::binary::read_u32_le(txt_plcf_data, i * 4) {
                text_cps.push(cp);
            }
        }

        let subdoc_len = subdoc_end.checked_sub(subdoc_start).ok_or_else(|| {
            PackageError::Corrupted("note subdocument range is reversed".to_string())
        })?;
        if subdoc_len == 0
            || text_cps[..text_cps.len() - 1]
                .iter()
                .any(|&cp| cp >= subdoc_len)
            || text_cps[..text_cps.len() - 1]
                .windows(2)
                .any(|pair| pair[0] >= pair[1])
        {
            return Err(PackageError::Corrupted(
                "note text PLCF has invalid character positions".to_string(),
            ));
        }
        if text_cps[text_cps.len() - 2] != subdoc_len - 1 {
            return Err(PackageError::Corrupted(
                "note text PLCF terminator must equal subdocument length minus one".to_string(),
            ));
        }

        let mut references = Vec::with_capacity(ref_plcf.count());
        for i in 0..ref_plcf.count() {
            if let Some((ref_cp, _)) = ref_plcf.range(i)
                && let Some(desc_data) = ref_plcf.property(i)
                && let Some(descriptor) = FootnoteDescriptor::from_bytes(desc_data)
            {
                let text_start = subdoc_start.checked_add(text_cps[i]).ok_or_else(|| {
                    PackageError::Corrupted("note text start CP overflows".to_string())
                })?;
                let text_end = subdoc_start.checked_add(text_cps[i + 1]).ok_or_else(|| {
                    PackageError::Corrupted("note text end CP overflows".to_string())
                })?;

                references.push(FootnoteReference::new(
                    ref_cp, text_start, text_end, descriptor,
                ));
            }
        }

        if references.len() != ref_plcf.count() {
            return Err(PackageError::Corrupted(
                "note reference PLCF contains an invalid descriptor".to_string(),
            ));
        }
        if references
            .windows(2)
            .any(|pair| pair[0].ref_cp >= pair[1].ref_cp)
        {
            return Err(PackageError::Corrupted(
                "note reference PLCF character positions are not unique and increasing".to_string(),
            ));
        }

        Ok(references)
    }

    /// Get all footnote references
    #[must_use]
    pub fn references(&self) -> &[FootnoteReference] {
        &self.references
    }

    /// Get the count of footnotes
    #[must_use]
    pub fn count(&self) -> usize {
        self.references.len()
    }

    /// Find a footnote at a specific character position in the main document
    #[must_use]
    pub fn find_at_position(&self, cp: u32) -> Option<&FootnoteReference> {
        self.references.iter().find(|f| f.ref_cp == cp)
    }

    /// Get the footnote reference at a specific index
    #[must_use]
    pub fn get_at_index(&self, index: usize) -> Option<&FootnoteReference> {
        self.references.get(index)
    }

    /// Check if a footnote exists at a specific character position
    #[must_use]
    pub fn exists_at_position(&self, cp: u32) -> bool {
        self.references.iter().any(|f| f.ref_cp == cp)
    }
}

/// Endnotes table parser
pub struct EndnotesTable {
    /// All endnote references
    references: Vec<FootnoteReference>,
}

impl EndnotesTable {
    /// Parse endnotes from the FIB and table stream
    ///
    /// # Arguments
    ///
    /// * `fib` - File Information Block
    /// * `table_stream` - The table stream (0Table or 1Table)
    ///
    /// # Returns
    ///
    /// A parsed `EndnotesTable`
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let mut references = Vec::new();

        // Check if endnotes exist in the document
        if let Some((subdoc_start, subdoc_end)) = fib.get_endnote_range() {
            // Parse endnote reference PLCF (plcfEndRef)
            // FIB index 46: fcPlcfEndRef and lcbPlcfEndRef
            if let Some((offset, length)) = fib.get_table_pointer(46)
                && length > 0
                && (offset as usize) < table_stream.len()
            {
                let plcf_data = &table_stream[offset as usize..];
                let plcf_len = length.min((table_stream.len() - offset as usize) as u32) as usize;

                if plcf_len >= 4 {
                    // Parse reference PLCF with 2-byte FRD descriptors
                    if let Some(ref_plcf) = Plcf::parse(&plcf_data[..plcf_len], 2) {
                        // Parse endnote text PLCF (plcfEndTxt)
                        // FIB index 47: fcPlcfEndTxt and lcbPlcfEndTxt
                        if let Some((txt_offset, txt_length)) = fib.get_table_pointer(47)
                            && txt_length > 0
                            && (txt_offset as usize) < table_stream.len()
                        {
                            let txt_plcf_data = &table_stream[txt_offset as usize..];
                            let txt_plcf_len = txt_length
                                .min((table_stream.len() - txt_offset as usize) as u32)
                                as usize;

                            references = FootnotesTable::parse_footnote_plcfs(
                                &ref_plcf,
                                &txt_plcf_data[..txt_plcf_len],
                                subdoc_start,
                                subdoc_end,
                            )?;
                        }
                    }
                }
            }
        }

        Ok(Self { references })
    }

    /// Get all endnote references
    #[must_use]
    pub fn references(&self) -> &[FootnoteReference] {
        &self.references
    }

    /// Find an endnote at a specific character position in the main document
    #[must_use]
    pub fn find_at_position(&self, cp: u32) -> Option<&FootnoteReference> {
        self.references.iter().find(|e| e.ref_cp == cp)
    }

    /// Get the count of endnotes
    #[must_use]
    pub fn count(&self) -> usize {
        self.references.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn reference_plcf_bytes(cps: &[u32], descriptors: &[u16]) -> Vec<u8> {
        let mut bytes = Vec::new();
        for cp in cps {
            bytes.extend_from_slice(&cp.to_le_bytes());
        }
        for descriptor in descriptors {
            bytes.extend_from_slice(&descriptor.to_le_bytes());
        }
        bytes
    }

    #[test]
    fn test_footnote_descriptor_parsing() {
        let data = [0x01, 0x00]; // number = 1
        let desc = FootnoteDescriptor::from_bytes(&data).unwrap();
        assert_eq!(desc.number, 1);
    }

    #[test]
    fn test_footnote_reference() {
        let desc = FootnoteDescriptor { number: 1 };
        let reference = FootnoteReference::new(100, 5000, 5100, desc);
        assert_eq!(reference.ref_cp, 100);
        assert_eq!(reference.text_length(), 100);
    }

    #[test]
    fn parses_spec_terminal_note_character_positions() {
        let reference_data = reference_plcf_bytes(&[1, 10], &[1]);
        let references = Plcf::parse(&reference_data, 2).unwrap();
        let text_cps = [0u32, 5, 6]
            .into_iter()
            .flat_map(u32::to_le_bytes)
            .collect::<Vec<_>>();

        let parsed =
            FootnotesTable::parse_footnote_plcfs(&references, &text_cps, 100, 106).unwrap();
        assert_eq!(parsed.len(), 1);
        assert_eq!((parsed[0].text_start_cp, parsed[0].text_end_cp), (100, 105));
        assert_eq!(parsed[0].descriptor.number, 1);
    }

    #[test]
    fn rejects_malformed_note_character_positions() {
        let reference_data = reference_plcf_bytes(&[1, 10], &[1]);
        let references = Plcf::parse(&reference_data, 2).unwrap();
        let encode = |cps: &[u32]| {
            cps.iter()
                .flat_map(|cp| cp.to_le_bytes())
                .collect::<Vec<_>>()
        };

        assert!(
            FootnotesTable::parse_footnote_plcfs(&references, &encode(&[0, 0, 5]), 100, 106)
                .is_err()
        );
        assert!(
            FootnotesTable::parse_footnote_plcfs(&references, &encode(&[0, 3, 5]), 100, 106)
                .is_err()
        );
        assert!(FootnotesTable::parse_footnote_plcfs(&references, &[0; 11], 100, 106).is_err());
    }
}
