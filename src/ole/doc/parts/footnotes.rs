/// Footnotes and endnotes parser for Word binary format.
///
/// Based on Apache POI's FootnotesTables and LibreOffice's implementation.
/// Footnotes and endnotes are stored in separate subdocuments with references in the main text.
use super::super::package::Result;
use super::fib::FileInformationBlock;
use crate::ole::plcf::PlcfParser;

/// Reference descriptor for footnote/endnote (FRD structure - 2 bytes)
#[derive(Debug, Clone, Copy)]
pub struct FootnoteDescriptor {
    /// Reference number (auto-numbered or custom)
    pub number: u16,
}

impl FootnoteDescriptor {
    /// Parse a footnote descriptor from 2 bytes
    pub fn from_bytes(data: &[u8]) -> Option<Self> {
        if data.len() < 2 {
            return None;
        }

        // FRD structure (2 bytes):
        // Bit 0-15: nAuto (auto-number or custom number mark)
        let number = crate::common::binary::read_u16_le(data, 0).ok()?;

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
    /// A parsed FootnotesTable
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let mut references = Vec::new();

        // Check if footnotes exist in the document
        if let Some((subdoc_start, _subdoc_end)) = fib.get_footnote_range() {
            // Parse footnote reference PLCF (plcfFndRef)
            // FIB index 5: fcPlcfFndRef and lcbPlcfFndRef
            if let Some((offset, length)) = fib.get_table_pointer(5)
                && length > 0
                && (offset as usize) < table_stream.len()
            {
                let plcf_data = &table_stream[offset as usize..];
                let plcf_len = length.min((table_stream.len() - offset as usize) as u32) as usize;

                if plcf_len >= 4 {
                    // Parse reference PLCF with 2-byte FRD descriptors
                    if let Some(ref_plcf) = PlcfParser::parse(&plcf_data[..plcf_len], 2) {
                        // Parse footnote text PLCF (plcfFndTxt)
                        // FIB index 6: fcPlcfFndTxt and lcbPlcfFndTxt
                        if let Some((txt_offset, txt_length)) = fib.get_table_pointer(6)
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
                            );
                        }
                    }
                }
            }
        }

        Ok(Self { references })
    }

    /// Parse footnote PLCF structures
    fn parse_footnote_plcfs(
        ref_plcf: &PlcfParser,
        txt_plcf_data: &[u8],
        subdoc_start: u32,
    ) -> Vec<FootnoteReference> {
        // Parse text PLCF with element_size = 0 (just CPs)
        // Manually parse since PlcfParser expects element_size > 0
        let cp_count = txt_plcf_data.len() / 4;
        if cp_count < 2 {
            return Vec::new();
        }

        let mut text_cps = Vec::with_capacity(cp_count);
        for i in 0..cp_count {
            if let Ok(cp) = crate::common::binary::read_u32_le(txt_plcf_data, i * 4) {
                text_cps.push(cp);
            }
        }

        let mut references = Vec::new();
        for i in 0..ref_plcf.count() {
            if let Some((ref_cp, _)) = ref_plcf.range(i)
                && let Some(desc_data) = ref_plcf.property(i)
                && let Some(descriptor) = FootnoteDescriptor::from_bytes(desc_data)
                && i < text_cps.len() - 1
            {
                let text_start = subdoc_start + text_cps[i];
                let text_end = subdoc_start + text_cps[i + 1];

                references.push(FootnoteReference::new(
                    ref_cp, text_start, text_end, descriptor,
                ));
            }
        }

        references
    }

    /// Get all footnote references
    pub fn references(&self) -> &[FootnoteReference] {
        &self.references
    }

    /// Get the count of footnotes
    pub fn count(&self) -> usize {
        self.references.len()
    }

    /// Find a footnote at a specific character position in the main document
    pub fn find_at_position(&self, cp: u32) -> Option<&FootnoteReference> {
        self.references.iter().find(|f| f.ref_cp == cp)
    }

    /// Get the footnote reference at a specific index
    pub fn get_at_index(&self, index: usize) -> Option<&FootnoteReference> {
        self.references.get(index)
    }

    /// Check if a footnote exists at a specific character position
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
    /// A parsed EndnotesTable
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Self> {
        let mut references = Vec::new();

        // Check if endnotes exist in the document
        if let Some((subdoc_start, _subdoc_end)) = fib.get_endnote_range() {
            // Parse endnote reference PLCF (plcfEndRef)
            // FIB index 7: fcPlcfEndRef and lcbPlcfEndRef
            if let Some((offset, length)) = fib.get_table_pointer(7)
                && length > 0
                && (offset as usize) < table_stream.len()
            {
                let plcf_data = &table_stream[offset as usize..];
                let plcf_len = length.min((table_stream.len() - offset as usize) as u32) as usize;

                if plcf_len >= 4 {
                    // Parse reference PLCF with 2-byte FRD descriptors
                    if let Some(ref_plcf) = PlcfParser::parse(&plcf_data[..plcf_len], 2) {
                        // Parse endnote text PLCF (plcfEndTxt)
                        // FIB index 8: fcPlcfEndTxt and lcbPlcfEndTxt
                        if let Some((txt_offset, txt_length)) = fib.get_table_pointer(8)
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
                            );
                        }
                    }
                }
            }
        }

        Ok(Self { references })
    }

    /// Get all endnote references
    pub fn references(&self) -> &[FootnoteReference] {
        &self.references
    }

    /// Find an endnote at a specific character position in the main document
    pub fn find_at_position(&self, cp: u32) -> Option<&FootnoteReference> {
        self.references.iter().find(|e| e.ref_cp == cp)
    }

    /// Get the count of endnotes
    pub fn count(&self) -> usize {
        self.references.len()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

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
}
