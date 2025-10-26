//! FKP (Formatted Disk Page) parser for DOC files.
//!
//! FKPs are 512-byte pages used to store character and paragraph properties.
//! Based on Apache POI's FormattedDiskPage and CHPFormattedDiskPage.
//!
//! References:
//! - org.apache.poi.hwpf.model.FormattedDiskPage
//! - org.apache.poi.hwpf.model.CHPFormattedDiskPage
//! - org.apache.poi.hwpf.model.PAPFormattedDiskPage
//! - [MS-DOC] 2.4.3 PnFkp* (Page Number FKP)

use crate::common::binary::{read_i32_le, read_u16_le, read_u32_le};

/// Size of an FKP page in bytes (always 512)
const FKP_PAGE_SIZE: usize = 512;

/// Size of BX structure in PAPX FKP (13 bytes: 1 byte offset + 12 bytes PHE)
const PAPX_BX_SIZE: usize = 13;

/// Size of File Character position entry (4 bytes)
const FC_SIZE: usize = 4;

/// ParagraphHeight (PHE) structure - 12 bytes.
///
/// This structure is part of PAPX BX entries and contains paragraph height information.
/// Based on org.apache.poi.hwpf.model.ParagraphHeight.
#[derive(Debug, Clone, Copy)]
pub struct ParagraphHeight {
    /// Info field containing various bit flags
    pub info_field: u16,
    /// Reserved field
    pub reserved: u16,
    /// Column width
    pub dxa_col: i32,
    /// Line or height dimension
    pub dym_line_or_height: i32,
}

impl ParagraphHeight {
    /// Parse ParagraphHeight from 12 bytes.
    ///
    /// # Arguments
    ///
    /// * `data` - Slice containing at least 12 bytes
    /// * `offset` - Offset in data to start reading
    ///
    /// # Returns
    ///
    /// Parsed ParagraphHeight or None if insufficient data
    #[inline]
    pub fn parse(data: &[u8], offset: usize) -> Option<Self> {
        if offset + 12 > data.len() {
            return None;
        }

        Some(Self {
            info_field: read_u16_le(data, offset).ok()?,
            reserved: read_u16_le(data, offset + 2).ok()?,
            dxa_col: read_i32_le(data, offset + 4).ok()?,
            dym_line_or_height: read_i32_le(data, offset + 8).ok()?,
        })
    }
}

/// A single entry in an FKP page.
///
/// Each entry contains a File Character position (FC) and associated property data.
#[derive(Debug, Clone)]
pub struct FkpEntry {
    /// File Character position (byte offset in WordDocument stream)
    pub fc: u32,
    /// Property data (grpprl - group of SPRMs)
    /// For CHP: CHPX data
    /// For PAP: PAPX data
    pub grpprl: Vec<u8>,
    /// Paragraph height (only for PAPX entries, None for CHPX)
    pub paragraph_height: Option<ParagraphHeight>,
}

/// CHPX FKP (Character Property Formatted Disk Page).
///
/// Based on Apache POI's CHPFormattedDiskPage.
/// Each 512-byte page contains:
/// - FC array at start (4 bytes each)
/// - BX array after FCs (1 byte each, offset to grpprl)
/// - grpprl data at end of page (grows backwards)
/// - crun count at byte 511
#[derive(Debug, Clone)]
pub struct ChpxFkp {
    /// Entries in this FKP page
    entries: Vec<FkpEntry>,
}

impl ChpxFkp {
    /// Parse a CHPX FKP from a 512-byte page.
    ///
    /// # Arguments
    ///
    /// * `page_data` - 512-byte FKP page data
    /// * `data_stream` - The main data stream (WordDocument) for reading full CHPXs
    ///
    /// # Returns
    ///
    /// Parsed CHPX FKP or None if invalid
    pub fn parse(page_data: &[u8], _data_stream: &[u8]) -> Option<Self> {
        if page_data.len() != FKP_PAGE_SIZE {
            return None;
        }

        // Last byte contains crun (count of runs)
        let crun = page_data[511] as usize;
        if crun == 0 || crun > 101 {
            // Sanity check: max 101 entries per page
            return None;
        }

        let mut entries = Vec::with_capacity(crun);

        // Parse FC array (4 bytes each, crun+1 entries)
        // FCs define boundaries: [fc0..fc1), [fc1..fc2), etc.
        let mut fcs = Vec::with_capacity(crun + 1);
        for i in 0..=crun {
            let offset = i * 4;
            if offset + 4 > page_data.len() {
                return None;
            }
            let fc = read_u32_le(page_data, offset).unwrap_or(0);
            fcs.push(fc);
        }

        // Parse BX array (1 byte each, crun entries)
        // Each BX is offset*2 from start of page to grpprl
        let bx_offset = (crun + 1) * 4;
        for (i, fc_start) in fcs.iter().enumerate().take(crun) {
            let bx_index = bx_offset + i;

            if bx_index >= page_data.len() {
                return None;
            }

            let bx = page_data[bx_index] as usize;

            // BX = 0 means no formatting (use default)
            if bx == 0 {
                entries.push(FkpEntry {
                    fc: *fc_start,
                    grpprl: Vec::new(),
                    paragraph_height: None,
                });
                continue;
            }

            // Calculate actual offset: bx * 2
            let grpprl_offset = bx * 2;
            if grpprl_offset >= FKP_PAGE_SIZE {
                return None;
            }

            // First byte at grpprl_offset is cb (count of bytes)
            if grpprl_offset >= page_data.len() {
                return None;
            }

            let cb = page_data[grpprl_offset] as usize;

            // grpprl data starts at grpprl_offset + 1
            let grpprl_start = grpprl_offset + 1;
            let grpprl_end = grpprl_start + cb;

            if grpprl_end > FKP_PAGE_SIZE {
                return None;
            }

            let grpprl = page_data[grpprl_start..grpprl_end].to_vec();

            entries.push(FkpEntry {
                fc: *fc_start,
                grpprl,
                paragraph_height: None,
            });
        }

        Some(Self { entries })
    }

    /// Get the number of entries in this FKP.
    #[inline]
    pub fn count(&self) -> usize {
        self.entries.len()
    }

    /// Get an entry by index.
    #[inline]
    pub fn entry(&self, index: usize) -> Option<&FkpEntry> {
        self.entries.get(index)
    }

    /// Get all entries.
    #[inline]
    pub fn entries(&self) -> &[FkpEntry] {
        &self.entries
    }
}

/// PAPX FKP (Paragraph Property Formatted Disk Page).
///
/// Based on Apache POI's PAPFormattedDiskPage.
/// Each 512-byte page contains:
/// - FC array at start (4 bytes each)
/// - BX array after FCs (13 bytes each: 1 byte offset + 12 bytes PHE)
/// - grpprl data at end of page (grows backwards)
/// - cpara count at byte 511
///
/// The BX structure for PAPX is different from CHPX:
/// - Byte 0: offset/2 to the grpprl
/// - Bytes 1-12: ParagraphHeight (PHE) structure
#[derive(Debug, Clone)]
pub struct PapxFkp {
    /// Entries in this FKP page
    entries: Vec<FkpEntry>,
}

impl PapxFkp {
    /// Parse a PAPX FKP from a 512-byte page.
    ///
    /// # Arguments
    ///
    /// * `page_data` - 512-byte FKP page data
    /// * `_data_stream` - The data stream (for huge PAPX, not yet implemented)
    ///
    /// # Returns
    ///
    /// Parsed PAPX FKP or None if invalid
    pub fn parse(page_data: &[u8], _data_stream: &[u8]) -> Option<Self> {
        if page_data.len() != FKP_PAGE_SIZE {
            return None;
        }

        // Last byte contains cpara (count of paragraphs)
        let cpara = page_data[511] as usize;
        if cpara == 0 || cpara > 101 {
            // Sanity check: max 101 entries per page
            return None;
        }

        let mut entries = Vec::with_capacity(cpara);

        // Parse FC array (4 bytes each, cpara+1 entries)
        // FCs define boundaries: [fc0..fc1), [fc1..fc2), etc.
        let mut fcs = Vec::with_capacity(cpara + 1);
        for i in 0..=cpara {
            let offset = i * FC_SIZE;
            if offset + FC_SIZE > page_data.len() {
                return None;
            }
            let fc = read_u32_le(page_data, offset).unwrap_or(0);
            fcs.push(fc);
        }

        // Parse BX array (13 bytes each for PAPX)
        // Structure: [offset byte][12-byte ParagraphHeight]
        let bx_offset = (cpara + 1) * FC_SIZE;
        for (i, fc_start) in fcs.iter().enumerate().take(cpara) {
            let bx_index = bx_offset + (i * PAPX_BX_SIZE);

            if bx_index + PAPX_BX_SIZE > page_data.len() {
                return None;
            }

            // First byte of BX: offset/2 to grpprl (or 0 for no formatting)
            let bx_offset_byte = page_data[bx_index] as usize;

            // Parse ParagraphHeight from bytes 1-12 of BX entry
            let phe = ParagraphHeight::parse(page_data, bx_index + 1)?;

            // If bx_offset_byte is 0, use default formatting (empty grpprl)
            if bx_offset_byte == 0 {
                entries.push(FkpEntry {
                    fc: *fc_start,
                    grpprl: Vec::new(),
                    paragraph_height: Some(phe),
                });
                continue;
            }

            // Calculate actual offset: bx_offset_byte * 2
            let papx_offset = bx_offset_byte * 2;
            if papx_offset >= FKP_PAGE_SIZE {
                return None;
            }

            // Parse grpprl size and data
            // According to POI's getGrpprl implementation:
            // - Read first byte at papx_offset
            // - If it's 0, read next byte for size (in words, multiply by 2)
            // - Otherwise, use first byte as size (in words, multiply by 2) and decrement by 1
            let grpprl = Self::parse_papx_grpprl(page_data, papx_offset)?;

            entries.push(FkpEntry {
                fc: *fc_start,
                grpprl,
                paragraph_height: Some(phe),
            });
        }

        Some(Self { entries })
    }

    /// Parse PAPX grpprl data at the given offset.
    ///
    /// The size encoding is complex:
    /// - If first byte is 0: next byte contains size (in words)
    /// - Otherwise: first byte contains size (in words), subtract 1
    ///
    /// Based on POI's PAPFormattedDiskPage.getGrpprl()
    #[inline]
    fn parse_papx_grpprl(page_data: &[u8], papx_offset: usize) -> Option<Vec<u8>> {
        if papx_offset >= page_data.len() {
            return None;
        }

        let first_byte = page_data[papx_offset] as usize;
        let (size_in_words, grpprl_start) = if first_byte == 0 {
            // Size is in next byte
            if papx_offset + 1 >= page_data.len() {
                return None;
            }
            let size = page_data[papx_offset + 1] as usize;
            (size, papx_offset + 2)
        } else {
            // Size is in first byte, but subtract 1
            (first_byte - 1, papx_offset + 1)
        };

        // Convert size from words to bytes
        let size_in_bytes = size_in_words * 2;

        // Validate bounds
        let grpprl_end = grpprl_start + size_in_bytes;
        if grpprl_end > FKP_PAGE_SIZE {
            return None;
        }

        // Use slice borrowing to avoid allocation when possible,
        // but we need to return Vec<u8> for the API consistency
        Some(page_data[grpprl_start..grpprl_end].to_vec())
    }

    /// Get the number of entries in this FKP.
    #[inline]
    pub fn count(&self) -> usize {
        self.entries.len()
    }

    /// Get an entry by index.
    #[inline]
    pub fn entry(&self, index: usize) -> Option<&FkpEntry> {
        self.entries.get(index)
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_fkp_page_size() {
        assert_eq!(FKP_PAGE_SIZE, 512);
    }

    #[test]
    fn test_invalid_page_size() {
        let small_page = vec![0u8; 100];
        assert!(ChpxFkp::parse(&small_page, &[]).is_none());

        let large_page = vec![0u8; 1000];
        assert!(ChpxFkp::parse(&large_page, &[]).is_none());
    }

    #[test]
    fn test_invalid_crun() {
        let mut page = vec![0u8; 512];
        page[511] = 0; // crun = 0
        assert!(ChpxFkp::parse(&page, &[]).is_none());

        page[511] = 200; // crun > 101
        assert!(ChpxFkp::parse(&page, &[]).is_none());
    }

    #[test]
    fn test_chpx_fkp_basic_parsing() {
        // Create a simple CHPX FKP with 2 entries
        let mut page = vec![0u8; 512];

        // Set crun = 2
        page[511] = 2;

        // Set FC array (3 FCs for 2 runs: boundaries at 0, 100, 200)
        // FC[0] = 0
        page[0..4].copy_from_slice(&0u32.to_le_bytes());
        // FC[1] = 100
        page[4..8].copy_from_slice(&100u32.to_le_bytes());
        // FC[2] = 200
        page[8..12].copy_from_slice(&200u32.to_le_bytes());

        // BX array starts at offset 12 (3 FCs * 4 bytes)
        // BX[0] points to offset 250 (250/2 = 125)
        page[12] = 125;
        // BX[1] points to offset 260 (260/2 = 130)
        page[13] = 130;

        // Place grpprl data at offset 250
        page[250] = 4; // size = 4 bytes
        page[251..255].copy_from_slice(&[0x35, 0x08, 0x01, 0x00]); // Example: bold

        // Place grpprl data at offset 260
        page[260] = 6; // size = 6 bytes
        page[261..267].copy_from_slice(&[0x36, 0x08, 0x01, 0x00, 0x35, 0x08]); // Example: italic + bold

        let fkp = ChpxFkp::parse(&page, &[]).expect("Failed to parse CHPX FKP");

        assert_eq!(fkp.count(), 2);

        // Check first entry
        let entry0 = fkp.entry(0).expect("Failed to get entry 0");
        assert_eq!(entry0.fc, 0);
        assert_eq!(entry0.grpprl.len(), 4);
        assert!(entry0.paragraph_height.is_none());

        // Check second entry
        let entry1 = fkp.entry(1).expect("Failed to get entry 1");
        assert_eq!(entry1.fc, 100);
        assert_eq!(entry1.grpprl.len(), 6);
        assert!(entry1.paragraph_height.is_none());
    }

    #[test]
    fn test_chpx_fkp_zero_bx() {
        // Test that BX=0 creates an entry with empty grpprl
        let mut page = vec![0u8; 512];

        page[511] = 1; // crun = 1

        // Set FC array
        page[0..4].copy_from_slice(&0u32.to_le_bytes());
        page[4..8].copy_from_slice(&100u32.to_le_bytes());

        // BX[0] = 0 (no formatting)
        page[8] = 0;

        let fkp = ChpxFkp::parse(&page, &[]).expect("Failed to parse CHPX FKP");

        assert_eq!(fkp.count(), 1);

        let entry = fkp.entry(0).expect("Failed to get entry 0");
        assert_eq!(entry.fc, 0);
        assert!(entry.grpprl.is_empty());
        assert!(entry.paragraph_height.is_none());
    }

    #[test]
    fn test_papx_fkp_basic_parsing() {
        // Create a simple PAPX FKP with 1 entry
        let mut page = vec![0u8; 512];

        // Set cpara = 1
        page[511] = 1;

        // Set FC array (2 FCs for 1 paragraph: boundaries at 0, 1000)
        page[0..4].copy_from_slice(&0u32.to_le_bytes());
        page[4..8].copy_from_slice(&1000u32.to_le_bytes());

        // BX array starts at offset 8 (2 FCs * 4 bytes)
        // BX structure: [offset byte][12-byte PHE]
        // Offset byte points to grpprl at offset 200 (200/2 = 100)
        page[8] = 100;

        // Set ParagraphHeight (12 bytes starting at offset 9)
        page[9..11].copy_from_slice(&0x0001u16.to_le_bytes()); // info_field
        page[11..13].copy_from_slice(&0x0000u16.to_le_bytes()); // reserved
        page[13..17].copy_from_slice(&240i32.to_le_bytes()); // dxa_col
        page[17..21].copy_from_slice(&360i32.to_le_bytes()); // dym_line_or_height

        // Place grpprl data at offset 200
        // Size encoding: if first byte != 0, size = (first_byte - 1) * 2
        page[200] = 3; // size = (3-1)*2 = 4 bytes
        page[201..205].copy_from_slice(&[0x01, 0x02, 0x03, 0x04]); // grpprl data

        let fkp = PapxFkp::parse(&page, &[]).expect("Failed to parse PAPX FKP");

        assert_eq!(fkp.count(), 1);

        let entry = fkp.entry(0).expect("Failed to get entry 0");
        assert_eq!(entry.fc, 0);
        assert_eq!(entry.grpprl.len(), 4);

        let phe = entry
            .paragraph_height
            .expect("ParagraphHeight should exist");
        assert_eq!(phe.info_field, 0x0001);
        assert_eq!(phe.reserved, 0x0000);
        assert_eq!(phe.dxa_col, 240);
        assert_eq!(phe.dym_line_or_height, 360);
    }

    #[test]
    fn test_papx_fkp_zero_size_encoding() {
        // Test the size=0 encoding where size is in the next byte
        let mut page = vec![0u8; 512];

        page[511] = 1; // cpara = 1

        // Set FC array
        page[0..4].copy_from_slice(&0u32.to_le_bytes());
        page[4..8].copy_from_slice(&1000u32.to_le_bytes());

        // BX array
        page[8] = 100; // offset to grpprl

        // Set ParagraphHeight (minimal)
        page[9..21].fill(0);

        // Place grpprl data at offset 200 with zero-byte encoding
        page[200] = 0; // First byte is 0, size is in next byte
        page[201] = 5; // size = 5 words = 10 bytes
        page[202..212]
            .copy_from_slice(&[0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A]);

        let fkp = PapxFkp::parse(&page, &[]).expect("Failed to parse PAPX FKP");

        assert_eq!(fkp.count(), 1);

        let entry = fkp.entry(0).expect("Failed to get entry 0");
        assert_eq!(entry.grpprl.len(), 10);
        assert_eq!(
            &entry.grpprl[..],
            &[0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0A]
        );
    }

    #[test]
    fn test_papx_fkp_zero_bx() {
        // Test that BX offset byte = 0 creates entry with empty grpprl
        let mut page = vec![0u8; 512];

        page[511] = 1; // cpara = 1

        // Set FC array
        page[0..4].copy_from_slice(&0u32.to_le_bytes());
        page[4..8].copy_from_slice(&1000u32.to_le_bytes());

        // BX array
        page[8] = 0; // offset = 0 (no formatting)

        // Set ParagraphHeight (12 bytes starting at offset 9)
        page[9..11].copy_from_slice(&0x0002u16.to_le_bytes()); // info_field
        page[11..13].copy_from_slice(&0x0000u16.to_le_bytes()); // reserved
        page[13..17].copy_from_slice(&0i32.to_le_bytes()); // dxa_col
        page[17..21].copy_from_slice(&0i32.to_le_bytes()); // dym_line_or_height

        let fkp = PapxFkp::parse(&page, &[]).expect("Failed to parse PAPX FKP");

        assert_eq!(fkp.count(), 1);

        let entry = fkp.entry(0).expect("Failed to get entry 0");
        assert_eq!(entry.fc, 0);
        assert!(entry.grpprl.is_empty());

        let phe = entry
            .paragraph_height
            .expect("ParagraphHeight should exist");
        assert_eq!(phe.info_field, 0x0002);
    }

    #[test]
    fn test_paragraph_height_parsing() {
        let data = [
            0x01, 0x00, // info_field = 1
            0x02, 0x00, // reserved = 2
            0xF0, 0x00, 0x00, 0x00, // dxa_col = 240
            0x68, 0x01, 0x00, 0x00, // dym_line_or_height = 360
        ];

        let phe = ParagraphHeight::parse(&data, 0).expect("Failed to parse ParagraphHeight");

        assert_eq!(phe.info_field, 1);
        assert_eq!(phe.reserved, 2);
        assert_eq!(phe.dxa_col, 240);
        assert_eq!(phe.dym_line_or_height, 360);
    }

    #[test]
    fn test_paragraph_height_insufficient_data() {
        let data = [0u8; 10]; // Only 10 bytes, need 12
        assert!(ParagraphHeight::parse(&data, 0).is_none());

        let data = [0u8; 20];
        assert!(ParagraphHeight::parse(&data, 10).is_none()); // offset 10 + 12 > 20
    }
}
