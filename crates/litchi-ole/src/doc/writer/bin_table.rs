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

/// Generate a bin table from multi-page FKP results.
///
/// Takes the FC ranges from [`FkpPages`] and associates each with a page number.
///
/// # Arguments
///
/// * `ranges` - FC ranges from `FkpPages::ranges`: `(fc_first, fc_last)` per page
/// * `first_page_number` - Page number of the first FKP page in the WordDocument stream
///
/// # Returns
///
/// Bin table as bytes: `(n+1) × 4` FCs followed by `n × 4` page numbers.
pub fn generate_bin_table_from_pages(ranges: &[(u32, u32)], first_page_number: u32) -> Vec<u8> {
    let n = ranges.len();
    // Size: (n+1)*4 FCs + n*4 PNs
    let mut bte = Vec::with_capacity((n + 1) * 4 + n * 4);

    // Write (n+1) FCs: one start FC per page + final end FC
    for (fc_first, _) in ranges {
        bte.extend_from_slice(&fc_first.to_le_bytes());
    }
    // Final FC = end of last page
    if let Some((_, fc_last)) = ranges.last() {
        bte.extend_from_slice(&fc_last.to_le_bytes());
    }

    // Write n page numbers (contiguous from first_page_number)
    for i in 0..n {
        let pn = first_page_number + i as u32;
        bte.extend_from_slice(&pn.to_le_bytes());
    }

    bte
}
