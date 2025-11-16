//! Bin table (PLCFBTE) generation for DOC files
//!
//! Bin tables map file offsets (FC, in bytes from start of `WordDocument`) to
//! FKP page numbers (PNs). This follows Apache POI's implementation where the
//! PLCF CP array actually stores FC values (`CharIndexTranslator.getByteIndex`).
//! See [MS-DOC] Section 2.8 and POI's CHPBinTable/PAPBinTable.

/// Generate a bin table (PLCF of FKP page numbers)
///
/// # Arguments
///
/// * `start_fc` - Starting file offset (FC)
/// * `end_fc` - Ending file offset (FC)
/// * `page_number` - FKP page number in WordDocument stream (byte offset / 512)
///
/// # Returns
///
/// Bin table as bytes (PLCF structure)
pub fn generate_single_entry_bin_table(start_fc: u32, end_fc: u32, page_number: u32) -> Vec<u8> {
    let mut bte = Vec::new();

    // PLCF structure for 1 entry:
    // - 2 FCs: start_fc, end_fc (8 bytes)
    // - 1 page number (4 bytes)
    // Total: 12 bytes

    // Write the 2 FCs
    bte.extend_from_slice(&start_fc.to_le_bytes());
    bte.extend_from_slice(&end_fc.to_le_bytes());

    // Write the page number
    bte.extend_from_slice(&page_number.to_le_bytes());

    bte
}

/// Generate a bin table (PLCF of FKP page numbers) - legacy API
///
/// # Arguments
///
/// * `fkp_positions` - Vector of (fc, page_number) tuples where:
///   - fc: file offset (FC)
///   - page_number: Page number in WordDocument stream (byte offset / 512)
///
/// # Returns
///
/// Bin table as bytes (PLCF structure)
pub fn generate_bin_table(fkp_positions: Vec<(u32, u32)>) -> Vec<u8> {
    let mut bte = Vec::new();

    // PLCF structure based on Apache POI's PlexOfCps.toByteArray():
    // For n properties:
    // - Array of n+1 FCs (file offsets) - 4 bytes each
    // - Array of n data structures (page numbers for bin table) - 4 bytes each
    //
    // For 0 properties (empty bin table):
    // - Just 1 FC (the ending offset = 0) - 4 bytes total

    if fkp_positions.is_empty() {
        // Empty bin table - just one FC set to 0 (POI line 99-101, 117)
        bte.extend_from_slice(&0u32.to_le_bytes());
        return bte;
    }

    // Write n FCs (starting offsets)
    for (fc, _) in &fkp_positions {
        bte.extend_from_slice(&fc.to_le_bytes());
    }

    // Write the final FC (ending offset) - POI line 117
    if let Some((last_fc, _)) = fkp_positions.last() {
        bte.extend_from_slice(&last_fc.to_le_bytes());
    }

    // Write n page numbers (PNs) - POI line 114
    for (_, pn) in &fkp_positions {
        bte.extend_from_slice(&pn.to_le_bytes());
    }

    bte
}
