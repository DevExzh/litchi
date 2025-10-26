/// Proper paragraph extraction from binary structures.
///
/// Based on Apache POI's HWPF paragraph parsing logic, this module
/// extracts paragraphs using PLCF (Property List with Character Positions)
/// structures for paragraph boundaries (PAP) and character runs (CHP).
use super::super::package::Result;
use super::chp::CharacterProperties;
use super::chp_bin_table::ChpBinTable;
use super::fib::FileInformationBlock;
use super::pap::ParagraphProperties;
use super::piece_table::PieceTable;
use crate::ole::plcf::PlcfParser;
use crate::ole::sprm::parse_sprms;

/// Type alias for extracted paragraph data: (text, properties, runs).
pub(crate) type ExtractedParagraph = (
    String,
    ParagraphProperties,
    Vec<(String, CharacterProperties)>,
);

/// Paragraph extractor using binary structures.
///
/// Based on Apache POI's ParagraphPropertiesTable (PAPBinTable) and
/// CharacterPropertiesTable (CHPBinTable).
pub struct ParagraphExtractor {
    /// Paragraph property data (PLCF)
    pap_plcf: Option<PlcfParser>,
    /// Character property bin table (properly parsed with FKP)
    chp_bin_table: Option<ChpBinTable>,
    /// The extracted text
    text: String,
    /// Text piece character positions
    text_ranges: Vec<(u32, u32, usize)>, // (cp_start, cp_end, text_offset)
    /// Character position range to extract (for subdocuments)
    cp_range: Option<(u32, u32)>,
}

impl ParagraphExtractor {
    /// Create a new paragraph extractor.
    ///
    /// # Arguments
    ///
    /// * `fib` - File Information Block
    /// * `table_stream` - Table stream (0Table or 1Table) data
    /// * `word_document` - WordDocument stream data
    /// * `text` - Extracted document text
    pub fn new(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        word_document: &[u8],
        text: String,
    ) -> Result<Self> {
        // Get PAP bin table location from FIB
        // Index 13 in FibRgFcLcb97 is fcPlcfBtePapx/lcbPlcfBtePapx (PLCFBTEPAPX)
        let pap_plcf = if let Some((offset, length)) = fib.get_table_pointer(13) {
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

        // Parse piece table from CLX (Complex file information)
        // According to [MS-DOC], fcClx is at FIB offset 0x01A2
        // In FibRgFcLcb97 (starting at FIB offset 154), this is index 33: (0x01A2 - 154) / 8 = 33
        let piece_table = if let Some((offset, length)) = fib.get_table_pointer(33) {
            if length > 0 && (offset as usize) < table_stream.len() {
                let clx_data = &table_stream[offset as usize..];
                let clx_len = length.min((table_stream.len() - offset as usize) as u32) as usize;
                PieceTable::parse(&clx_data[..clx_len])
            } else {
                None
            }
        } else {
            None
        };

        // Get CHP bin table location from FIB and parse it properly with FKP
        // Index 12 in FibRgFcLcb97 is fcPlcfBteChpx/lcbPlcfBteChpx (PLCFBTECHPX)
        // Requires piece table for FC-to-CP conversion
        let chp_bin_table = if let (Some((offset, length)), Some(pt)) =
            (fib.get_table_pointer(12), &piece_table)
        {
            if length > 0 && (offset as usize) < table_stream.len() {
                let chp_data = &table_stream[offset as usize..];
                let chp_len = length.min((table_stream.len() - offset as usize) as u32) as usize;
                if chp_len >= 8 {
                    // Parse CHPBinTable (PlcfBteChpx with FKP pages)
                    // FKP pages are in WordDocument stream, not table stream!
                    ChpBinTable::parse(&chp_data[..chp_len], word_document, pt)
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
            chp_bin_table,
            text,
            text_ranges,
            cp_range: None,
        })
    }

    /// Create a new paragraph extractor for a specific character position range.
    ///
    /// This is used to extract paragraphs from subdocuments (footnotes, headers, etc.).
    ///
    /// # Arguments
    ///
    /// * `fib` - File Information Block
    /// * `table_stream` - Table stream (0Table or 1Table) data
    /// * `word_document` - WordDocument stream data
    /// * `text` - Extracted document text
    /// * `cp_range` - Character position range (start_cp, end_cp)
    pub fn new_with_range(
        fib: &FileInformationBlock,
        table_stream: &[u8],
        word_document: &[u8],
        text: String,
        cp_range: (u32, u32),
    ) -> Result<Self> {
        let mut extractor = Self::new(fib, table_stream, word_document, text)?;
        extractor.cp_range = Some(cp_range);
        Ok(extractor)
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
    ///
    /// Based on MS-DOC specification and Apache POI's approach:
    /// Paragraphs in Word documents are delimited by CR (\r = 0x000D) characters.
    /// The PAP PLCF stores formatting properties, but doesn't define paragraph boundaries.
    pub fn extract_paragraphs(&self) -> Result<Vec<ExtractedParagraph>> {
        let mut paragraphs = Vec::new();

        // Determine the CP range to process
        let doc_start_cp = self.cp_range.map(|(start, _)| start).unwrap_or(0);
        let doc_end_cp = self
            .cp_range
            .map(|(_, end)| end)
            .unwrap_or(self.text.chars().count() as u32);

        // Find all paragraph breaks (CR characters) in the text
        // CR (0x000D / '\r') marks the end of each paragraph in Word documents
        let mut para_boundaries = vec![doc_start_cp];
        let mut current_cp = doc_start_cp;

        for c in self.text.chars().skip(doc_start_cp as usize) {
            if c == '\r' {
                para_boundaries.push(current_cp + 1); // Position after CR
            }
            current_cp += 1;
            if current_cp >= doc_end_cp {
                break;
            }
        }

        // Ensure we have an end boundary
        if para_boundaries.last() != Some(&doc_end_cp) && current_cp > doc_start_cp {
            para_boundaries.push(current_cp.min(doc_end_cp));
        }

        // Extract each paragraph
        for i in 0..para_boundaries.len().saturating_sub(1) {
            let para_start = para_boundaries[i];
            let para_end = para_boundaries[i + 1];

            if para_start >= para_end {
                continue;
            }

            // Extract paragraph text (excluding the CR marker itself)
            let mut para_text = self.extract_text_range(para_start, para_end);
            // Remove trailing CR if present
            if para_text.ends_with('\r') {
                para_text.pop();
            }

            // Find matching PAP properties for this paragraph
            // PAP PLCF entries define formatting, not boundaries
            let para_props = if let Some(ref pap_plcf) = self.pap_plcf {
                let mut found_props = None;
                for j in 0..pap_plcf.count() {
                    if let Some((pap_start, pap_end)) = pap_plcf.range(j) {
                        // Check if this PAP entry overlaps with our paragraph
                        if pap_start <= para_start && para_start < pap_end {
                            if let Some(prop_data) = pap_plcf.property(j) {
                                found_props = Self::parse_papx(prop_data).ok();
                            }
                            break;
                        }
                    }
                }
                found_props.unwrap_or_default()
            } else {
                ParagraphProperties::default()
            };

            // Extract character runs within this paragraph (excluding the CR)
            let para_text_end = if para_end > para_start
                && self.text.chars().nth((para_end - 1) as usize) == Some('\r')
            {
                para_end - 1
            } else {
                para_end
            };
            let runs = self.extract_runs(para_start, para_text_end)?;

            paragraphs.push((para_text, para_props, runs));
        }

        // Fallback if no paragraphs were found
        if paragraphs.is_empty() && !self.text.is_empty() {
            let runs = self.extract_runs(doc_start_cp, doc_end_cp)?;
            paragraphs.push((self.text.clone(), ParagraphProperties::default(), runs));
        }

        Ok(paragraphs)
    }

    /// Extract text for a character position range.
    fn extract_text_range(&self, cp_start: u32, cp_end: u32) -> String {
        // Clamp CPs to valid range
        let max_cp = self.text_ranges.len() as u32;
        let cp_start_clamped = cp_start.min(max_cp);
        let cp_end_clamped = cp_end.min(max_cp);

        if cp_start_clamped >= cp_end_clamped {
            return String::new();
        }

        let start_idx = cp_start_clamped as usize;
        let end_idx = cp_end_clamped as usize;

        if start_idx < self.text_ranges.len() {
            let start_offset = self.text_ranges[start_idx].2;
            let end_offset = if end_idx < self.text_ranges.len() {
                self.text_ranges[end_idx].2
            } else {
                self.text.len()
            };

            if start_offset <= end_offset {
                self.text[start_offset..end_offset].to_string()
            } else {
                String::new()
            }
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

        if let Some(ref chp_bin_table) = self.chp_bin_table {
            // Get runs that overlap with this paragraph
            let overlapping_runs = chp_bin_table.runs_in_range(para_start, para_end);

            for run in overlapping_runs {
                // Calculate actual run boundaries within paragraph
                let actual_start = run.start_cp.max(para_start);
                let actual_end = run.end_cp.min(para_end);

                if actual_start >= actual_end {
                    continue;
                }

                // Extract run text
                let run_text = self.extract_text_range(actual_start, actual_end);

                // Skip empty runs
                if run_text.is_empty() {
                    continue;
                }

                runs.push((run_text, run.properties.clone()));
            }
        }

        // If no runs found, return the whole paragraph as one run
        if runs.is_empty() {
            let para_text = self.extract_text_range(para_start, para_end);
            if !para_text.is_empty() {
                runs.push((para_text, CharacterProperties::default()));
            }
        } else {
            // Filter out empty runs before consolidation to prevent style markers without text
            runs.retain(|(text, _)| !text.is_empty());

            // Consolidate consecutive runs with identical formatting
            // This prevents markdown like **a****b****c** instead of **abc**
            if !runs.is_empty() {
                runs = Self::consolidate_runs(runs);
            }
        }

        Ok(runs)
    }

    /// Consolidate consecutive runs with identical formatting properties.
    ///
    /// This prevents creating separate runs for each character when they have
    /// the same styling, which would result in markdown like `**a****b****c**`
    /// instead of the desired `**abc**`.
    ///
    /// # Note
    ///
    /// Empty runs are skipped to prevent style markers without text (e.g., `****`).
    fn consolidate_runs(
        runs: Vec<(String, CharacterProperties)>,
    ) -> Vec<(String, CharacterProperties)> {
        if runs.is_empty() {
            return runs;
        }

        let mut consolidated = Vec::new();
        let mut current_text = String::new();
        let mut current_props = runs[0].1.clone();

        for (text, props) in runs {
            // Skip empty runs entirely to prevent empty style markers
            if text.is_empty() {
                continue;
            }

            // Compare only visual formatting properties (not object markers)
            // pic_offset and is_ole2 are position/reference markers, not formatting
            let props_match = props.is_bold == current_props.is_bold
                && props.is_italic == current_props.is_italic
                && props.underline == current_props.underline
                && props.is_strikethrough == current_props.is_strikethrough;

            if props_match {
                // Same formatting - append to current run
                current_text.push_str(&text);
            } else {
                // Different formatting - save current run and start new one
                if !current_text.is_empty() {
                    consolidated.push((current_text.clone(), current_props.clone()));
                }
                current_text = text;
                current_props = props;
            }
        }

        // Don't forget the last run (only if non-empty)
        if !current_text.is_empty() {
            consolidated.push((current_text, current_props));
        }

        consolidated
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
        // Parse SPRMs (always 2-byte opcodes per Apache POI)
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
                },
                0x840F | 0x000F => {
                    // Left indent
                    props.indent_left = sprm.operand_i16().map(|v| v as i32);
                },
                0x8411 | 0x0011 => {
                    // Right indent
                    props.indent_right = sprm.operand_i16().map(|v| v as i32);
                },
                0x2416 => {
                    // sprmPFInTable - paragraph is in a table
                    props.in_table = sprm.operand_byte().unwrap_or(0) != 0;
                },
                0x2417 => {
                    // sprmPFTtp - table row end marker (table trailer paragraph)
                    props.is_table_row_end = sprm.operand_byte().unwrap_or(0) != 0;
                },
                0x6649 => {
                    // sprmPItap - table nesting level (4-byte operand)
                    props.table_nesting_level = sprm.operand_dword().unwrap_or(0) as i32;
                },
                _ => {},
            }
        }

        Ok(props)
    }
}
