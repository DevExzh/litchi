/// Proper paragraph extraction from binary structures.
///
/// Based on Apache POI's HWPF paragraph parsing logic, this module
/// extracts paragraphs using PLCF (Property List with Character Positions)
/// structures for paragraph boundaries (PAP) and character runs (CHP).
use super::super::package::Result;
use super::fib::FileInformationBlock;
use super::pap::ParagraphProperties;
use super::chp::CharacterProperties;
use crate::ole::{binary::PlcfParser, sprm::parse_sprms};

/// Type alias for extracted paragraph data: (text, properties, runs).
type ExtractedParagraph = (String, ParagraphProperties, Vec<(String, CharacterProperties)>);

/// Paragraph extractor using binary structures.
///
/// Based on Apache POI's ParagraphPropertiesTable (PAPBinTable) and
/// CharacterPropertiesTable (CHPBinTable).
pub struct ParagraphExtractor {
    /// Paragraph property data (PLCF)
    pap_plcf: Option<PlcfParser>,
    /// Character property data (PLCF)
    chp_plcf: Option<PlcfParser>,
    /// The extracted text
    text: String,
    /// Text piece character positions
    text_ranges: Vec<(u32, u32, usize)>, // (cp_start, cp_end, text_offset)
}

impl ParagraphExtractor {
    /// Create a new paragraph extractor.
    ///
    /// # Arguments
    ///
    /// * `fib` - File Information Block
    /// * `table_stream` - Table stream (0Table or 1Table) data
    /// * `text` - Extracted document text
    pub fn new(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        text: String,
    ) -> Result<Self> {
        // Get PAP bin table location from FIB
        // Index 2 in FibRgFcLcb97 is fcPlcfBtePapx/lcbPlcfBtePapx
        let pap_plcf = if let Some((offset, length)) = fib.get_table_pointer(2) {
            if length > 0 && (offset as usize) < table_stream.len() {
                let pap_data = &table_stream[offset as usize..];
                let pap_len = length.min((table_stream.len() - offset as usize) as u32) as usize;
                if pap_len >= 4 {
                    // PAP PLCF uses 4-byte property descriptors initially
                    // Each entry points to a PAPX (paragraph properties) structure
                    PlcfParser::parse(&pap_data[..pap_len], 4)
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            None
        };

        // Get CHP bin table location from FIB
        // Index 1 in FibRgFcLcb97 is fcPlcfBteChpx/lcbPlcfBteChpx
        let chp_plcf = if let Some((offset, length)) = fib.get_table_pointer(1) {
            if length > 0 && (offset as usize) < table_stream.len() {
                let chp_data = &table_stream[offset as usize..];
                let chp_len = length.min((table_stream.len() - offset as usize) as u32) as usize;
                if chp_len >= 4 {
                    // CHP PLCF uses 4-byte property descriptors
                    PlcfParser::parse(&chp_data[..chp_len], 4)
                } else {
                    None
                }
            } else {
                None
            }
        } else {
            None
        };

        // Build text ranges for mapping CPs to text offsets
        let text_ranges = Self::build_text_ranges(&text);

        Ok(Self {
            pap_plcf,
            chp_plcf,
            text,
            text_ranges,
        })
    }

    /// Build mapping from character positions to text offsets.
    fn build_text_ranges(text: &str) -> Vec<(u32, u32, usize)> {
        let mut ranges = Vec::new();
        let mut offset = 0usize;

        for (cp, ch) in text.chars().enumerate() {
            let char_len = ch.len_utf8();
            let cp_u32 = cp as u32;
            ranges.push((cp_u32, cp_u32 + 1, offset));
            offset += char_len;
        }

        ranges
    }

    /// Extract paragraphs with properties.
    ///
    /// Returns a vector of (text, paragraph_properties, character_runs) tuples.
    pub fn extract_paragraphs(&self) -> Result<Vec<ExtractedParagraph>> {
        let mut paragraphs = Vec::new();

        if let Some(ref pap_plcf) = self.pap_plcf {
            // Iterate through paragraph boundaries
            for i in 0..pap_plcf.count() {
                if let Some((para_start, para_end)) = pap_plcf.range(i) {
                    // Extract paragraph text
                    let para_text = self.extract_text_range(para_start, para_end);

                    // Parse paragraph properties
                    let para_props = if let Some(prop_data) = pap_plcf.property(i) {
                        // Property data points to a PAPX structure
                        // For now, use default properties - full implementation would
                        // follow the PAPX pointer to get actual properties
                        Self::parse_papx(prop_data).unwrap_or_default()
                    } else {
                        ParagraphProperties::default()
                    };

                    // Extract character runs within this paragraph
                    let runs = self.extract_runs(para_start, para_end)?;

                    paragraphs.push((para_text, para_props, runs));
                }
            }
        } else {
            // Fallback: split by newlines if no PLCF data
            for line in self.text.lines() {
                paragraphs.push((
                    line.to_string(),
                    ParagraphProperties::default(),
                    vec![(line.to_string(), CharacterProperties::default())],
                ));
            }
        }

        Ok(paragraphs)
    }

    /// Extract text for a character position range.
    fn extract_text_range(&self, cp_start: u32, cp_end: u32) -> String {
        let start_idx = cp_start as usize;
        let end_idx = cp_end as usize;

        if start_idx < self.text_ranges.len() && end_idx <= self.text_ranges.len() {
            let start_offset = self.text_ranges[start_idx].2;
            let end_offset = if end_idx < self.text_ranges.len() {
                self.text_ranges[end_idx].2
            } else {
                self.text.len()
            };

            self.text[start_offset..end_offset].to_string()
        } else {
            String::new()
        }
    }

    /// Extract character runs (formatted text segments) within a paragraph.
    fn extract_runs(
        &self,
        para_start: u32,
        para_end: u32,
    ) -> Result<Vec<(String, CharacterProperties)>> {
        let mut runs = Vec::new();

        if let Some(ref chp_plcf) = self.chp_plcf {
            // Find character runs that overlap with this paragraph
            for i in 0..chp_plcf.count() {
                if let Some((run_start, run_end)) = chp_plcf.range(i) {
                    // Check if this run overlaps with the paragraph
                    if run_end <= para_start || run_start >= para_end {
                        continue;
                    }

                    // Calculate actual run boundaries within paragraph
                    let actual_start = run_start.max(para_start);
                    let actual_end = run_end.min(para_end);

                    // Extract run text
                    let run_text = self.extract_text_range(actual_start, actual_end);

                    // Parse character properties
                    let char_props = if let Some(prop_data) = chp_plcf.property(i) {
                        Self::parse_chpx(prop_data).unwrap_or_default()
                    } else {
                        CharacterProperties::default()
                    };

                    runs.push((run_text, char_props));
                }
            }
        }

        // If no runs found, return the whole paragraph as one run
        if runs.is_empty() {
            let para_text = self.extract_text_range(para_start, para_end);
            runs.push((para_text, CharacterProperties::default()));
        }

        Ok(runs)
    }

    /// Parse PAPX (Paragraph Property eXceptions) data.
    ///
    /// Based on Apache POI's PAPX.getParagraphProperties().
    fn parse_papx(prop_data: &[u8]) -> Result<ParagraphProperties> {
        if prop_data.len() < 4 {
            return Ok(ParagraphProperties::default());
        }

        // PAPX format (simplified):
        // - 4 bytes: pointer to actual PAPX data in data stream (for file-based) or length
        // In memory structures, this often contains inline SPRM data

        // For now, try to parse as SPRM data directly
        let sprms = parse_sprms(prop_data);

        // Apply SPRMs to create paragraph properties
        let mut props = ParagraphProperties::default();

        for sprm in &sprms {
            // Apply common paragraph SPRMs
            match sprm.opcode {
                0x2403 | 0x0003 => {
                    // Justification
                    if let Some(val) = sprm.operand_byte() {
                        props.justification = match val {
                            0 => super::pap::Justification::Left,
                            1 => super::pap::Justification::Center,
                            2 => super::pap::Justification::Right,
                            3 => super::pap::Justification::Justified,
                            _ => super::pap::Justification::Left,
                        };
                    }
                }
                0x840F | 0x000F => {
                    // Left indent
                    props.indent_left = sprm.operand_i16().map(|v| v as i32);
                }
                0x8411 | 0x0011 => {
                    // Right indent
                    props.indent_right = sprm.operand_i16().map(|v| v as i32);
                }
                _ => {}
            }
        }

        Ok(props)
    }

    /// Parse CHPX (Character Property eXceptions) data.
    ///
    /// Based on Apache POI's CHPX.getCharacterProperties().
    fn parse_chpx(prop_data: &[u8]) -> Result<CharacterProperties> {
        if prop_data.is_empty() {
            return Ok(CharacterProperties::default());
        }

        // Parse as SPRM data
        let sprms = parse_sprms(prop_data);

        // Apply SPRMs to create character properties
        let mut props = CharacterProperties::default();

        for sprm in &sprms {
            match sprm.opcode {
                0x0835 | 0x0085 => {
                    // Bold
                    props.is_bold = Some(sprm.operand_byte().unwrap_or(0) != 0);
                }
                0x0836 | 0x0086 => {
                    // Italic
                    props.is_italic = Some(sprm.operand_byte().unwrap_or(0) != 0);
                }
                0x4A43 | 0x0043 => {
                    // Font size
                    props.font_size = sprm.operand_word();
                }
                _ => {}
            }
        }

        Ok(props)
    }
}

