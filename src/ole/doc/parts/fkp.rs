/// FKP (Formatted Disk Page) parser for DOC files.
///
/// FKPs are 512-byte pages used to store character and paragraph properties.
/// Based on Apache POI's FormattedDiskPage and CHPFormattedDiskPage.
///
/// References:
/// - org.apache.poi.hwpf.model.FormattedDiskPage
/// - org.apache.poi.hwpf.model.CHPFormattedDiskPage
/// - [MS-DOC] 2.4.3 PnFkp* (Page Number FKP)

use super::super::package::Result;

/// Size of an FKP page in bytes (always 512)
const FKP_PAGE_SIZE: usize = 512;

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
            let fc = u32::from_le_bytes([
                page_data[offset],
                page_data[offset + 1],
                page_data[offset + 2],
                page_data[offset + 3],
            ]);
            fcs.push(fc);
        }

        // Parse BX array (1 byte each, crun entries)
        // Each BX is offset*2 from start of page to grpprl
        let bx_offset = (crun + 1) * 4;
        for i in 0..crun {
            let fc_start = fcs[i];
            let bx_index = bx_offset + i;
            
            if bx_index >= page_data.len() {
                return None;
            }
            
            let bx = page_data[bx_index] as usize;
            
            // BX = 0 means no formatting (use default)
            if bx == 0 {
                entries.push(FkpEntry {
                    fc: fc_start,
                    grpprl: Vec::new(),
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
            
            if entries.len() < 2 && cb > 0 {
                eprintln!("DEBUG:           FKP entry {}: bx={}, grpprl_offset={}, cb={}", i, bx, grpprl_offset, cb);
            }
            
            // grpprl data starts at grpprl_offset + 1
            let grpprl_start = grpprl_offset + 1;
            let grpprl_end = grpprl_start + cb;

            if grpprl_end > FKP_PAGE_SIZE {
                return None;
            }

            let grpprl = page_data[grpprl_start..grpprl_end].to_vec();

            entries.push(FkpEntry {
                fc: fc_start,
                grpprl,
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
/// Similar to CHPX FKP but for paragraph properties.
/// Currently not fully implemented as we're focusing on character properties.
#[derive(Debug, Clone)]
pub struct PapxFkp {
    /// Entries in this FKP page
    entries: Vec<FkpEntry>,
}

impl PapxFkp {
    /// Parse a PAPX FKP from a 512-byte page.
    pub fn parse(page_data: &[u8], _data_stream: &[u8]) -> Option<Self> {
        if page_data.len() != FKP_PAGE_SIZE {
            return None;
        }

        // Last byte contains cpara (count of paragraphs)
        let cpara = page_data[511] as usize;
        if cpara == 0 || cpara > 101 {
            return None;
        }

        let mut entries = Vec::with_capacity(cpara);

        // Parse FC array
        let mut fcs = Vec::with_capacity(cpara + 1);
        for i in 0..=cpara {
            let offset = i * 4;
            if offset + 4 > page_data.len() {
                return None;
            }
            let fc = u32::from_le_bytes([
                page_data[offset],
                page_data[offset + 1],
                page_data[offset + 2],
                page_data[offset + 3],
            ]);
            fcs.push(fc);
        }

        // Parse BX array (13 bytes each for PAPX, not 1 byte like CHPX)
        let bx_offset = (cpara + 1) * 4;
        for i in 0..cpara {
            let fc_start = fcs[i];
            let bx_index = bx_offset + (i * 13);
            
            if bx_index + 13 > page_data.len() {
                return None;
            }
            
            // For PAPX, BX structure is more complex
            // For now, use simplified parsing
            entries.push(FkpEntry {
                fc: fc_start,
                grpprl: Vec::new(),
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
}

