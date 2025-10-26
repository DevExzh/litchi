use super::chp::CharacterProperties;
/// CHPBinTable (Character Property Bin Table) parser.
///
/// Based on Apache POI's CHPBinTable class.
/// This handles the two-level structure of character properties in DOC files:
/// 1. PlcfBteChpx: Contains BTE entries with page numbers
/// 2. CHPXFKP pages: 512-byte pages containing actual character runs
///
/// References:
/// - org.apache.poi.hwpf.model.CHPBinTable
/// - org.apache.poi.hwpf.model.CHPFormattedDiskPage
/// - [MS-DOC] 2.8.5 PlcfBteChpx
use super::fkp::ChpxFkp;
use super::piece_table::PieceTable;
use crate::common::binary::read_u32_le;
use crate::ole::sprm::parse_sprms;

/// A character run with properties.
///
/// Represents a contiguous range of text with the same formatting.
#[derive(Debug, Clone)]
pub struct CharacterRun {
    /// Start character position
    pub start_cp: u32,
    /// End character position
    pub end_cp: u32,
    /// Character properties
    pub properties: CharacterProperties,
}

/// CHPBinTable - manages character property bin table.
///
/// Based on Apache POI's CHPBinTable.
#[derive(Debug)]
pub struct ChpBinTable {
    /// All character runs extracted from FKP pages
    runs: Vec<CharacterRun>,
}

impl ChpBinTable {
    /// Parse CHPBinTable from PlcfBteChpx data.
    ///
    /// # Arguments
    ///
    /// * `plcf_bte_chpx_data` - The PlcfBteChpx data from table stream
    /// * `word_document` - The WordDocument stream (FKP pages are stored here!)
    /// * `piece_table` - The piece table for FC-to-CP conversion
    ///
    /// # Returns
    ///
    /// Parsed CHPBinTable or None if invalid
    pub fn parse(
        plcf_bte_chpx_data: &[u8],
        word_document: &[u8],
        piece_table: &PieceTable,
    ) -> Option<Self> {
        // PlcfBteChpx structure:
        // - Array of FC positions (4 bytes each, n+1 entries)
        // - Array of PnBteChpx structures (4 bytes each, n entries)
        //
        // Each PnBteChpx contains:
        // - pn (PN): Page number (offset / 512) of CHPXFKP page

        if plcf_bte_chpx_data.len() < 8 {
            return None;
        }

        // Calculate number of BTE entries
        // PlcfBteChpx structure: array of (n+1) FCs followed by n PnFkpChpx structures
        // Each FC is 4 bytes, each PnFkpChpx is 4 bytes
        // Size = (n+1)*4 + n*4 = 8n + 4
        // So n = (size - 4) / 8
        let n = (plcf_bte_chpx_data.len() - 4) / 8;

        // Parsing PlcfBteChpx with n BTE entries

        let mut all_runs = Vec::new();

        // Parse each BTE entry
        for i in 0..n {
            // Read PN from PnFkpChpx array (starts after FC array)
            let pn_offset = (n + 1) * 4 + i * 4;
            if pn_offset + 4 > plcf_bte_chpx_data.len() {
                continue;
            }

            let pn_raw = read_u32_le(plcf_bte_chpx_data, pn_offset).unwrap_or(0);

            // PnFkpChpx structure (from [MS-DOC]):
            // - Bits 0-21 (22 bits): pn (Page Number)
            // - Bits 22-31 (10 bits): unused (MUST be ignored)
            // The ChpxFkp structure begins at offset pn * 512 in WordDocument stream
            let pn = pn_raw & 0x3FFFFF; // Mask to get 22-bit page number

            // PN (Page Number) format:
            // pageOffset = 512 * pn
            // FKP pages are stored in the WordDocument stream, not table stream!
            if pn == 0 || pn == 0x3FFFFF {
                // Invalid PN - skip this entry
                continue;
            }

            let page_offset = (pn as usize) * 512;

            if page_offset + 512 > word_document.len() {
                // Page offset exceeds document size - skip
                continue;
            }

            // Extract the 512-byte FKP page from WordDocument stream
            let fkp_page = &word_document[page_offset..page_offset + 512];

            // Parse CHPX FKP
            if let Some(fkp) = ChpxFkp::parse(fkp_page, word_document) {
                // Process each entry in the FKP
                for j in 0..fkp.count() {
                    if let Some(entry) = fkp.entry(j) {
                        // Get the next entry to determine end position
                        let end_fc = if j + 1 < fkp.count() {
                            fkp.entry(j + 1).map(|e| e.fc).unwrap_or(entry.fc)
                        } else {
                            // Last entry - use document end or next FKP's first FC
                            entry.fc + 1000000 // Placeholder
                        };

                        // Convert FC positions to CP using the piece table
                        let start_cp = piece_table.fc_to_cp(entry.fc).unwrap_or(entry.fc);
                        let end_cp = piece_table.fc_to_cp(end_fc).unwrap_or(end_fc);

                        // Parse CHPX (grpprl) to get character properties
                        let properties = Self::parse_chpx(&entry.grpprl);

                        if all_runs.len() < 5 && !entry.grpprl.is_empty() {
                            eprint!(
                                "DEBUG:       Entry {}: fc={}..{} -> cp={}..{}, grpprl_len={}, is_ole2={}, pic_offset={:?}, grpprl_bytes=",
                                j,
                                entry.fc,
                                end_fc,
                                start_cp,
                                end_cp,
                                entry.grpprl.len(),
                                properties.is_ole2,
                                properties.pic_offset
                            );
                            for b in entry.grpprl.iter().take(20) {
                                eprint!("{:02X} ", b);
                            }
                            eprintln!();
                        }

                        all_runs.push(CharacterRun {
                            start_cp,
                            end_cp,
                            properties,
                        });
                    }
                }
            } else {
                eprintln!("DEBUG: Failed to parse FKP at page offset {}", page_offset);
            }
        }

        eprintln!(
            "DEBUG: ChpBinTable parsed {} total runs (before deduplication)",
            all_runs.len()
        );

        // Sort runs by start CP, then by end CP
        // This is essential for proper merging and overlap detection
        all_runs.sort_by_key(|r| (r.start_cp, r.end_cp));

        // Remove overlapping/duplicate runs and fix boundaries
        // Following Apache POI's approach: consecutive CHPXs should not overlap
        let mut merged_runs = Vec::new();
        let mut last_end_cp = 0u32;

        for mut run in all_runs {
            // Clamp run boundaries to avoid overlaps
            if run.start_cp < last_end_cp {
                // This run overlaps with the previous one - skip or adjust
                if run.end_cp <= last_end_cp {
                    // Completely contained in previous run - skip
                    continue;
                }
                // Partially overlaps - adjust start to avoid overlap
                run.start_cp = last_end_cp;
            }

            // Skip invalid runs
            if run.start_cp >= run.end_cp {
                continue;
            }

            last_end_cp = run.end_cp;
            merged_runs.push(run);
        }

        eprintln!(
            "DEBUG: ChpBinTable after deduplication: {} runs",
            merged_runs.len()
        );

        Some(Self { runs: merged_runs })
    }

    /// Parse CHPX data (grpprl) into CharacterProperties.
    fn parse_chpx(grpprl: &[u8]) -> CharacterProperties {
        if grpprl.is_empty() {
            return CharacterProperties::default();
        }

        // Parse SPRMs (always 2-byte opcodes per Apache POI)
        let sprms = parse_sprms(grpprl);
        let mut props = CharacterProperties::default();

        for sprm in &sprms {
            match sprm.opcode {
                0x0835 | 0x0085 => {
                    // Bold
                    props.is_bold = Some(sprm.operand_byte().unwrap_or(0) != 0);
                },
                0x0836 | 0x0086 => {
                    // Italic
                    props.is_italic = Some(sprm.operand_byte().unwrap_or(0) != 0);
                },
                0x4A43 | 0x0043 => {
                    // Font size
                    props.font_size = sprm.operand_word();
                },
                0x080A => {
                    // OLE2 object flag (SPRM_FOLE2)
                    let operand = sprm.operand_byte().unwrap_or(0);
                    props.is_ole2 = operand != 0;
                    eprintln!(
                        "DEBUG: Found SPRM_FOLE2 in FKP, operand=0x{:02X}, operand_len={}, is_ole2={}",
                        operand,
                        sprm.operand.len(),
                        props.is_ole2
                    );
                },
                0x6A03 => {
                    // Picture location (sprmCPicLocation)
                    // This is the FILE character position (fc) of the picture/object data
                    props.pic_offset = sprm.operand_dword();
                    eprintln!(
                        "DEBUG: Found sprmCPicLocation (0x6A03) in FKP, pic_offset={:?}",
                        props.pic_offset
                    );
                },
                0x680E => {
                    // Object location/pic offset (SPRM_OBJLOCATION)
                    props.pic_offset = sprm.operand_dword();
                    eprintln!(
                        "DEBUG: Found SPRM_OBJLOCATION (0x680E) in FKP, pic_offset={:?}",
                        props.pic_offset
                    );
                },
                _ => {},
            }
        }

        props
    }

    /// Get all character runs.
    #[inline]
    pub fn runs(&self) -> &[CharacterRun] {
        &self.runs
    }

    /// Find runs that overlap with a character position range.
    pub fn runs_in_range(&self, start_cp: u32, end_cp: u32) -> Vec<&CharacterRun> {
        self.runs
            .iter()
            .filter(|run| {
                // Check if run overlaps with [start_cp, end_cp)
                run.end_cp > start_cp && run.start_cp < end_cp
            })
            .collect()
    }
}
