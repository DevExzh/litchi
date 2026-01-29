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
        // Pre-allocate: estimate ~10 runs per FKP page (conservative estimate)
        let estimated_runs = n.saturating_mul(10);
        let mut all_runs = Vec::with_capacity(estimated_runs);

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
                let entry_count = fkp.count();

                // Process entries in batch, caching next FC to avoid redundant lookups
                for j in 0..entry_count {
                    if let Some(entry) = fkp.entry(j) {
                        // Cache end_fc by looking ahead once
                        let end_fc = if j + 1 < entry_count {
                            // Get next entry's FC (more efficient than two entry() calls)
                            fkp.entry(j + 1).map(|e| e.fc).unwrap_or(entry.fc)
                        } else {
                            // Last entry - use large placeholder for document end
                            entry.fc.saturating_add(1_000_000)
                        };

                        // Convert FC positions to CP using the piece table
                        let start_cp = piece_table.fc_to_cp(entry.fc).unwrap_or(entry.fc);
                        let end_cp = piece_table.fc_to_cp(end_fc).unwrap_or(end_fc);

                        // Skip invalid ranges early
                        if start_cp >= end_cp {
                            continue;
                        }

                        // Parse CHPX (grpprl) to get character properties
                        let properties = Self::parse_chpx(&entry.grpprl);

                        all_runs.push(CharacterRun {
                            start_cp,
                            end_cp,
                            properties,
                        });
                    }
                }
            }
        }

        // Sort runs by start CP, then by end CP
        // This is essential for proper merging and overlap detection
        all_runs.sort_unstable_by_key(|r| (r.start_cp, r.end_cp));

        // Remove overlapping/duplicate runs and fix boundaries
        // Following Apache POI's approach: consecutive CHPXs should not overlap
        // Use retain_mut for in-place filtering (Rust 1.61+)
        let mut last_end_cp = 0u32;

        all_runs.retain_mut(|run| {
            // Clamp run boundaries to avoid overlaps
            if run.start_cp < last_end_cp {
                // This run overlaps with the previous one - skip or adjust
                if run.end_cp <= last_end_cp {
                    // Completely contained in previous run - skip
                    return false;
                }
                // Partially overlaps - adjust start to avoid overlap
                run.start_cp = last_end_cp;
            }

            // Skip invalid runs
            if run.start_cp >= run.end_cp {
                return false;
            }

            last_end_cp = run.end_cp;
            true
        });

        Some(Self { runs: all_runs })
    }

    /// Parse CHPX data (grpprl) into CharacterProperties.
    ///
    /// Delegates to CharacterProperties::from_sprm for consistent behavior
    /// with the full SPRM parser (handles is_spec, is_obj, is_data flags correctly).
    #[inline]
    fn parse_chpx(grpprl: &[u8]) -> CharacterProperties {
        if grpprl.is_empty() {
            return CharacterProperties::default();
        }

        // Use the complete SPRM parser from CharacterProperties
        // This ensures all flags (is_spec, is_obj, is_data, etc.) are set correctly
        CharacterProperties::from_sprm(grpprl).unwrap_or_default()
    }

    /// Get all character runs.
    #[inline]
    pub fn runs(&self) -> &[CharacterRun] {
        &self.runs
    }

    /// Find runs that overlap with a character position range.
    ///
    /// Returns an iterator to avoid unnecessary allocations.
    /// Use `.collect()` only if you need a Vec.
    pub fn runs_in_range(&self, start_cp: u32, end_cp: u32) -> impl Iterator<Item = &CharacterRun> {
        self.runs.iter().filter(move |run| {
            // Check if run overlaps with [start_cp, end_cp)
            run.end_cp > start_cp && run.start_cp < end_cp
        })
    }
}
