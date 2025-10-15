/// Proper paragraph extraction from binary structures.
///
/// Based on Apache POI's HWPF paragraph parsing logic, this module
/// extracts paragraphs using PLCF (Property List with Character Positions)
/// structures for paragraph boundaries (PAP) and character runs (CHP).
use super::super::package::Result;
use super::fib::FileInformationBlock;
use super::pap::ParagraphProperties;
use super::chp::CharacterProperties;
use super::chp_bin_table::ChpBinTable;
use super::piece_table::PieceTable;
use crate::ole::binary::PlcfParser;
use crate::ole::sprm::parse_sprms;

/// Type alias for extracted paragraph data: (text, properties, runs).
type ExtractedParagraph = (String, ParagraphProperties, Vec<(String, CharacterProperties)>);

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
        // Debug: Print all FIB table pointers to find CLX
        eprintln!("DEBUG: FIB table pointers with non-zero length:");
        for i in 0..50 {
            if let Some((offset, length)) = fib.get_table_pointer(i)
                && length > 0 {
                    eprintln!("DEBUG:   Index {}: offset={}, length={}", i, offset, length);
                }
        }
        
        // Get PAP bin table location from FIB
        // Index 2 in FibRgFcLcb97 is fcPlcfBtePapx/lcbPlcfBtePapx
        let pap_plcf = if let Some((offset, length)) = fib.get_table_pointer(2) {
            eprintln!("DEBUG: PAP table pointer: offset={}, length={}", offset, length);
            if length > 0 && (offset as usize) < table_stream.len() {
                let pap_data = &table_stream[offset as usize..];
                let pap_len = length.min((table_stream.len() - offset as usize) as u32) as usize;
                eprintln!("DEBUG: PAP data length={}", pap_len);
                if pap_len >= 4 {
                    // PAP PLCF uses 4-byte property descriptors initially
                    // Each entry points to a PAPX (paragraph properties) structure
                    let result = PlcfParser::parse(&pap_data[..pap_len], 4);
                    eprintln!("DEBUG: PAP PLCF parsed: {:?}", result.is_some());
                    if let Some(ref plcf) = result {
                        eprintln!("DEBUG: PAP PLCF has {} entries", plcf.count());
                    }
                    result
                } else {
                    eprintln!("DEBUG: PAP data too small");
                    None
                }
            } else {
                eprintln!("DEBUG: Invalid PAP offset/length");
                None
            }
        } else {
            eprintln!("DEBUG: No PAP table pointer in FIB");
            None
        };

        // Parse piece table from CLX (Complex file information)
        // According to [MS-DOC], fcClx is at FIB offset 0x01A2
        // In FibRgFcLcb97 (starting at FIB offset 154), this is index 33: (0x01A2 - 154) / 8 = 33
        let piece_table = if let Some((offset, length)) = fib.get_table_pointer(33) {
            eprintln!("DEBUG: CLX pointer: offset={}, length={}", offset, length);
            if length > 0 && (offset as usize) < table_stream.len() {
                let clx_data = &table_stream[offset as usize..];
                let clx_len = length.min((table_stream.len() - offset as usize) as u32) as usize;
                eprintln!("DEBUG: CLX data length={}", clx_len);
                let result = PieceTable::parse(&clx_data[..clx_len]);
                if let Some(ref pt) = result {
                    eprintln!("DEBUG: PieceTable has {} pieces, total_cps={}", pt.pieces().len(), pt.total_cps());
                } else {
                    eprintln!("DEBUG: Failed to parse PieceTable");
                }
                result
            } else {
                eprintln!("DEBUG: Invalid CLX offset/length");
                None
            }
        } else {
            eprintln!("DEBUG: No CLX pointer in FIB");
            None
        };

        // Get CHP bin table location from FIB and parse it properly with FKP
        // Index 1 in FibRgFcLcb97 is fcPlcfBteChpx/lcbPlcfBteChpx
        // Requires piece table for FC-to-CP conversion
        let chp_bin_table = if let (Some((offset, length)), Some(pt)) = (fib.get_table_pointer(1), &piece_table) {
            eprintln!("DEBUG: CHP table pointer: offset={}, length={}", offset, length);
            if length > 0 && (offset as usize) < table_stream.len() {
                let chp_data = &table_stream[offset as usize..];
                let chp_len = length.min((table_stream.len() - offset as usize) as u32) as usize;
                eprintln!("DEBUG: CHP data length={}", chp_len);
                if chp_len >= 8 {
                    // Parse CHPBinTable (PlcfBteChpx with FKP pages)
                    // FKP pages are in WordDocument stream, not table stream!
                    let result = ChpBinTable::parse(&chp_data[..chp_len], word_document, pt);
                    if let Some(ref bin_table) = result {
                        eprintln!("DEBUG: ChpBinTable has {} runs", bin_table.runs().len());
                    } else {
                        eprintln!("DEBUG: Failed to parse ChpBinTable");
                    }
                    result
                } else {
                    eprintln!("DEBUG: CHP data too small");
                    None
                }
            } else {
                eprintln!("DEBUG: Invalid CHP offset/length");
                None
            }
        } else {
            if piece_table.is_none() {
                eprintln!("DEBUG: No piece table available for CHP parsing");
            } else {
                eprintln!("DEBUG: No CHP table pointer in FIB");
            }
            None
        };

        // Build text ranges for mapping CPs to text offsets
        let text_ranges = Self::build_text_ranges(&text);

        Ok(Self {
            pap_plcf,
            chp_bin_table,
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
            eprintln!("DEBUG: Using fallback paragraph extraction (no PAP PLCF)");
            // Fallback: split by newlines if no PLCF data
            // But still try to extract character runs based on CHP
            let mut char_pos = 0u32;
            for line in self.text.lines() {
                let line_len = line.chars().count() as u32 + 1; // +1 for newline
                let line_end = char_pos + line_len;
                
                // Extract character runs for this line
                let runs = if !line.is_empty() {
                    self.extract_runs(char_pos, line_end).unwrap_or_else(|_| {
                        vec![(line.to_string(), CharacterProperties::default())]
                    })
                } else {
                    vec![(String::new(), CharacterProperties::default())]
                };
                
                paragraphs.push((
                    line.to_string(),
                    ParagraphProperties::default(),
                    runs,
                ));
                
                char_pos = line_end;
            }
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
            static mut DEBUG_COUNT: usize = 0;
            unsafe {
                DEBUG_COUNT += 1;
                if DEBUG_COUNT <= 3 {
                    eprintln!("DEBUG: extract_runs called, para_start={}, para_end={}, total_runs={}", 
                              para_start, para_end, chp_bin_table.runs().len());
                }
            }
            
            // Get runs that overlap with this paragraph
            let overlapping_runs = chp_bin_table.runs_in_range(para_start, para_end);
            
            let debug_count = unsafe { DEBUG_COUNT };
            if debug_count <= 3 {
                eprintln!("DEBUG:   Found {} overlapping runs", overlapping_runs.len());
            }
            
            for run in overlapping_runs {
                // Calculate actual run boundaries within paragraph
                let actual_start = run.start_cp.max(para_start);
                let actual_end = run.end_cp.min(para_end);

                // Extract run text
                let run_text = self.extract_text_range(actual_start, actual_end);
                
                if debug_count <= 3 && runs.len() < 5 {
                    eprintln!("DEBUG:     Run: cp={}..{}, is_ole2={}, pic_offset={:?}, text_len={}", 
                             actual_start, actual_end, run.properties.is_ole2, 
                             run.properties.pic_offset, run_text.len());
                }

                runs.push((run_text, run.properties.clone()));
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

}

