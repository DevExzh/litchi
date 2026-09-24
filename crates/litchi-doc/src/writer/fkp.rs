//! FKP (Formatted Disk Page) generation for DOC files
//!
//! FKPs store character and paragraph formatting information in a compact,
//! page-based structure. Each FKP is 512 bytes and contains multiple formatting
//! entries (CHPXs for character properties, PAPXs for paragraph properties).
//!
//! When entries exceed the capacity of a single FKP page, multiple pages are
//! generated automatically. The bin table (PlcfBte) maps FC ranges to FKP pages.
//!
//! Based on Microsoft's "[MS-DOC]" specification and Apache POI's
//! `CHPFormattedDiskPage` / `PAPFormattedDiskPage`.

/// I/O error returned while generating FKP pages.
pub type IoError = std::io::Error;

/// Character property (CHPX) entry
/// Represents a formatting run from `fc_start` to `fc_end`
#[derive(Debug, Clone)]
pub struct ChpxEntry {
    /// Start file character position (byte offset in `WordDocument` stream)
    pub fc_start: u32,
    /// End file character position (byte offset in `WordDocument` stream)
    pub fc_end: u32,
    /// Character properties (SPRM sequence)
    pub grpprl: Vec<u8>,
}

/// Paragraph property (PAPX) entry
/// Represents a formatting run from `fc_start` to `fc_end`
#[derive(Debug, Clone)]
pub struct PapxEntry {
    /// Start file character position (byte offset in `WordDocument` stream)
    pub fc_start: u32,
    /// End file character position (byte offset in `WordDocument` stream)
    pub fc_end: u32,
    /// Paragraph properties (SPRM sequence)
    pub grpprl: Vec<u8>,
}

/// Result of multi-page FKP generation: a list of 512-byte pages, each
/// covering a contiguous FC range.
#[derive(Debug)]
pub struct FkpPages {
    /// Each page is exactly 512 bytes.
    pub pages: Vec<Vec<u8>>,
    /// For each page: (`fc_first`, `fc_last`) — the FC range it covers.
    /// Used to build the bin table (`PlcfBte`).
    pub ranges: Vec<(u32, u32)>,
}

// ───────────────────── CHPX FKP ─────────────────────

/// Character FKP (CHPX FKP) builder.
///
/// Generates one or more 512-byte FKP pages from collected CHPX entries.
#[derive(Debug)]
pub struct ChpxFkpBuilder {
    entries: Vec<ChpxEntry>,
}

impl ChpxFkpBuilder {
    /// Create a new character FKP builder
    #[must_use]
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    /// Add a character formatting entry (a formatting run)
    pub fn add_entry(&mut self, fc_start: u32, fc_end: u32, grpprl: Vec<u8>) {
        self.entries.push(ChpxEntry {
            fc_start,
            fc_end,
            grpprl,
        });
    }

    /// Generate one or more 512-byte CHPX FKP pages.
    ///
    /// If all entries fit in a single page, a single page is returned.
    /// Otherwise entries are split across multiple pages.
    pub fn generate_pages(&self) -> Result<FkpPages, IoError> {
        if self.entries.is_empty() {
            // Return one empty page
            return Ok(FkpPages {
                pages: vec![vec![0u8; 512]],
                ranges: vec![(0, 0)],
            });
        }

        let mut pages = Vec::new();
        let mut ranges = Vec::new();
        let mut start = 0usize;

        while start < self.entries.len() {
            // Determine how many entries fit in this page.
            // Layout: (n+1)*4 FCs + n*1 RGB + grpprl data + 1 count = 512
            // Data grows from the end backwards; FCs+RGB grow from the front.
            let mut n = 0usize;
            let mut data_used = 1usize; // 1 byte for count at [511]
            for entry in &self.entries[start..] {
                let front = (n + 2) * 4 + (n + 1); // (n+1+1)*4 FCs + (n+1) RGB
                let grpprl_cost = if entry.grpprl.is_empty() {
                    0
                } else {
                    // 1 byte cb + grpprl, word-aligned
                    let raw = 1 + entry.grpprl.len();
                    raw + (raw % 2) // round up to even
                };
                if front + data_used + grpprl_cost > 512 {
                    break;
                }
                n += 1;
                data_used += grpprl_cost;
            }
            // At least 1 entry per page
            if n == 0 {
                n = 1;
            }
            // The estimate charges each CHPX its even-rounded size, but the
            // first CHPX placed below the count byte can need one byte more,
            // so a page filled to the estimate could overwrite its last BX.
            while n > 1 && !Self::page_fits(&self.entries[start..start + n]) {
                n -= 1;
            }

            let end = (start + n).min(self.entries.len());
            let page = Self::build_page(&self.entries[start..end])?;
            let fc_first = self.entries[start].fc_start;
            let fc_last = self.entries[end - 1].fc_end;
            pages.push(page);
            ranges.push((fc_first, fc_last));
            start = end;
        }

        Ok(FkpPages { pages, ranges })
    }

    /// Whether [`Self::build_page`] places every CHPX of `entries` after the
    /// page's FC and BX arrays.
    ///
    /// `build_page` stores the CHPXs downward from the count byte at offset
    /// 511 and aligns each start to an even offset, which is where a BX
    /// (offset / 2) must point.
    fn page_fits(entries: &[ChpxEntry]) -> bool {
        let property_start = (entries.len() + 1) * 4 + entries.len();
        let mut data_offset = 511usize;
        for entry in entries
            .iter()
            .rev()
            .filter(|entry| !entry.grpprl.is_empty())
        {
            let Some(chpx_start) = data_offset.checked_sub(1 + entry.grpprl.len().min(255)) else {
                return false;
            };
            data_offset = chpx_start - chpx_start % 2;
        }
        data_offset >= property_start
    }

    /// Build a single 512-byte CHPX FKP page from a slice of entries.
    fn build_page(entries: &[ChpxEntry]) -> Result<Vec<u8>, IoError> {
        let mut fkp = vec![0u8; 512];
        let n = entries.len();
        if n == 0 {
            return Ok(fkp);
        }

        // FC array: (n+1) × 4 bytes
        for (i, entry) in entries.iter().enumerate() {
            let off = i * 4;
            fkp[off..off + 4].copy_from_slice(&entry.fc_start.to_le_bytes());
        }
        // Final FC
        let last_fc_off = n * 4;
        fkp[last_fc_off..last_fc_off + 4].copy_from_slice(&entries[n - 1].fc_end.to_le_bytes());

        // Count byte
        fkp[511] = n as u8;

        // RGB array offset
        let rgb_off = (n + 1) * 4;
        let mut data_offset = 511usize; // grows backwards from count byte

        // Fill grpprl data in reverse order
        for (i, entry) in entries.iter().enumerate().rev() {
            if entry.grpprl.is_empty() {
                fkp[rgb_off + i] = 0;
            } else {
                let sz = entry.grpprl.len().min(255);
                let chpx_size = 1 + sz;
                data_offset -= chpx_size;
                data_offset -= data_offset % 2; // word-align

                fkp[data_offset] = sz as u8;
                fkp[data_offset + 1..data_offset + 1 + sz].copy_from_slice(&entry.grpprl[..sz]);
                fkp[rgb_off + i] = (data_offset / 2) as u8;
            }
        }

        Ok(fkp)
    }

    /// Legacy helper: generate a single 512-byte FKP (first page only).
    pub fn generate(&self) -> Result<Vec<u8>, IoError> {
        let pages = self.generate_pages()?;
        Ok(pages
            .pages
            .into_iter()
            .next()
            .unwrap_or_else(|| vec![0u8; 512]))
    }
}

impl Default for ChpxFkpBuilder {
    fn default() -> Self {
        Self::new()
    }
}

// ───────────────────── PAPX FKP ─────────────────────

/// Paragraph FKP (PAPX FKP) builder.
///
/// Generates one or more 512-byte FKP pages from collected PAPX entries.
#[derive(Debug)]
pub struct PapxFkpBuilder {
    entries: Vec<PapxEntry>,
}

/// BX entry size in a PAPX FKP (1 byte offset + 12 bytes PHE = 13)
const BX_SIZE: usize = 13;

impl PapxFkpBuilder {
    /// Create a new paragraph FKP builder
    #[must_use]
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    /// Add a paragraph formatting entry
    pub fn add_entry(&mut self, fc_start: u32, fc_end: u32, grpprl: Vec<u8>) {
        self.entries.push(PapxEntry {
            fc_start,
            fc_end,
            grpprl,
        });
    }

    /// Generate one or more 512-byte PAPX FKP pages.
    pub fn generate_pages(&self) -> Result<FkpPages, IoError> {
        if self.entries.is_empty() {
            return Ok(FkpPages {
                pages: vec![vec![0u8; 512]],
                ranges: vec![(0, 0)],
            });
        }

        let mut pages = Vec::new();
        let mut ranges = Vec::new();
        let mut start = 0usize;

        while start < self.entries.len() {
            // Determine capacity: (n+1)*4 FCs + n*13 BX + grpprl data + 1 count = 512
            let mut n = 0usize;
            let mut data_used = 1usize; // count byte
            for entry in &self.entries[start..] {
                let front = (n + 2) * 4 + (n + 1) * BX_SIZE;
                // grpprl cost: istd(2) + sprms, then cb (1 or 2 bytes), word-aligned
                let full_len = 2 + entry.grpprl.len();
                let extra = if (full_len % 2) > 0 { 1 } else { 2 };
                let grpprl_cost = full_len + extra;
                let grpprl_cost = grpprl_cost + (grpprl_cost % 2); // ensure even
                if front + data_used + grpprl_cost > 512 {
                    break;
                }
                n += 1;
                data_used += grpprl_cost;
            }
            if n == 0 {
                n = 1;
            }
            // As for CHPX pages: the first PAPX placed below the count byte
            // takes one byte more than its even-rounded estimate, so a page
            // filled to the estimate could overwrite its last BX entry.
            while n > 1 && !Self::page_fits(&self.entries[start..start + n]) {
                n -= 1;
            }

            let end = (start + n).min(self.entries.len());
            let page = Self::build_page(&self.entries[start..end])?;
            let fc_first = self.entries[start].fc_start;
            let fc_last = self.entries[end - 1].fc_end;
            pages.push(page);
            ranges.push((fc_first, fc_last));
            start = end;
        }

        Ok(FkpPages { pages, ranges })
    }

    /// Whether [`Self::build_page`] places every PAPX of `entries` after the
    /// page's FC and BX arrays.
    fn page_fits(entries: &[PapxEntry]) -> bool {
        let property_start = (entries.len() + 1) * 4 + entries.len() * BX_SIZE;
        let mut grpprl_offset = 511usize;
        for entry in entries {
            let len = 2 + entry.grpprl.len();
            let extra = if (len % 2) > 0 { 1 } else { 2 };
            let Some(papx_start) = grpprl_offset.checked_sub(len + extra) else {
                return false;
            };
            grpprl_offset = papx_start - papx_start % 2;
        }
        grpprl_offset >= property_start
    }

    /// Build a single 512-byte PAPX FKP page from a slice of entries.
    fn build_page(entries: &[PapxEntry]) -> Result<Vec<u8>, IoError> {
        let mut fkp = vec![0u8; 512];
        let n = entries.len();
        if n == 0 {
            return Ok(fkp);
        }

        // FC array
        for (i, entry) in entries.iter().enumerate() {
            let off = i * 4;
            fkp[off..off + 4].copy_from_slice(&entry.fc_start.to_le_bytes());
        }
        let last_fc_off = n * 4;
        fkp[last_fc_off..last_fc_off + 4].copy_from_slice(&entries[n - 1].fc_end.to_le_bytes());

        // Count byte
        fkp[511] = n as u8;

        // BX array offset
        let bx_off = (n + 1) * 4;
        let mut grpprl_offset = 511usize;

        // Fill grpprl data forward (matching the forward BX index order)
        for (i, entry) in entries.iter().enumerate().take(n) {
            // Build full grpprl: istd(2) + sprms
            let mut grpprl_full = Vec::with_capacity(2 + entry.grpprl.len());
            grpprl_full.extend_from_slice(&0u16.to_le_bytes()); // istd = 0 (Normal)
            grpprl_full.extend_from_slice(&entry.grpprl);
            let len = grpprl_full.len();

            let extra = if (len % 2) > 0 { 1 } else { 2 };
            grpprl_offset -= len + extra;
            grpprl_offset -= grpprl_offset % 2;

            // BX entry: word-offset pointer + 12 bytes PHE (zeros)
            let bx_pos = bx_off + i * BX_SIZE;
            fkp[bx_pos] = (grpprl_offset / 2) as u8;

            // Write cb + grpprl
            let mut copy_off = grpprl_offset;
            if (len % 2) > 0 {
                fkp[copy_off] = len.div_ceil(2) as u8;
                copy_off += 1;
            } else {
                fkp[copy_off + 1] = (len / 2) as u8;
                copy_off += 2;
            }
            fkp[copy_off..copy_off + len].copy_from_slice(&grpprl_full);
        }

        Ok(fkp)
    }

    /// Legacy helper: generate a single 512-byte FKP (first page only).
    pub fn generate(&self) -> Result<Vec<u8>, IoError> {
        let pages = self.generate_pages()?;
        Ok(pages
            .pages
            .into_iter()
            .next()
            .unwrap_or_else(|| vec![0u8; 512]))
    }
}

impl Default for PapxFkpBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_chpx_fkp() {
        let mut builder = ChpxFkpBuilder::new();
        builder.add_entry(0, 100, vec![0x80, 0x00]);
        builder.add_entry(100, 200, vec![0x80, 0x01]);

        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
        assert_eq!(fkp[511], 2);
    }

    #[test]
    fn test_chpx_fkp_empty() {
        let builder = ChpxFkpBuilder::new();
        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
    }

    #[test]
    fn test_chpx_fkp_single_entry() {
        let mut builder = ChpxFkpBuilder::new();
        builder.add_entry(0, 1000, vec![0x80, 0x00, 0x01, 0x02]);

        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
        assert_eq!(fkp[511], 1);
    }

    #[test]
    fn test_chpx_fkp_multiple_entries() {
        let mut builder = ChpxFkpBuilder::new();
        for i in 0..10 {
            builder.add_entry(i * 100, (i + 1) * 100, vec![0x80, i as u8]);
        }

        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
        assert_eq!(fkp[511], 10);
    }

    #[test]
    fn test_chpx_fkp_default() {
        let builder: ChpxFkpBuilder = Default::default();
        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
    }

    #[test]
    fn test_papx_fkp() {
        let mut builder = PapxFkpBuilder::new();
        builder.add_entry(0, 100, vec![0x80, 0x00]);

        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
    }

    #[test]
    fn test_papx_fkp_empty() {
        let builder = PapxFkpBuilder::new();
        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
    }

    #[test]
    fn test_papx_fkp_single_entry() {
        let mut builder = PapxFkpBuilder::new();
        builder.add_entry(0, 500, vec![0x40, 0x42, 0x08, 0x00]);

        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
    }

    #[test]
    fn test_papx_fkp_multiple_entries() {
        let mut builder = PapxFkpBuilder::new();
        for i in 0..20 {
            builder.add_entry(i * 100, (i + 1) * 100, vec![0x40, 0x42, i as u8, 0x00]);
        }

        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
    }

    #[test]
    fn test_papx_fkp_default() {
        let builder: PapxFkpBuilder = Default::default();
        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
    }

    #[test]
    fn test_multi_page_papx() {
        let mut builder = PapxFkpBuilder::new();
        // Add 40 entries — should overflow a single page (max ~29)
        for i in 0..40u32 {
            builder.add_entry(i * 100, (i + 1) * 100, vec![0x40, 0x42, 0x08, 0x00]);
        }
        let pages = builder.generate_pages().unwrap();
        assert!(pages.pages.len() >= 2, "Should produce multiple FKP pages");
        // Verify ranges are contiguous
        for i in 1..pages.ranges.len() {
            assert_eq!(pages.ranges[i].0, pages.ranges[i - 1].1);
        }
    }

    #[test]
    fn test_multi_page_chpx() {
        let mut builder = ChpxFkpBuilder::new();
        // Add many entries to trigger multi-page
        for i in 0..50u32 {
            builder.add_entry(i * 50, (i + 1) * 50, vec![0x80, 0x00, i as u8]);
        }
        let pages = builder.generate_pages().unwrap();
        assert!(
            !pages.pages.is_empty(),
            "Should produce at least one FKP page"
        );
    }

    /// Three CHPXs whose even-rounded estimate fills the page exactly while the
    /// odd-sized last CHPX needs one more byte when placed below the count byte.
    #[test]
    fn chpx_page_filled_to_the_estimate_never_overwrites_its_bx_array() {
        let mut builder = ChpxFkpBuilder::new();
        for (index, size) in [255usize, 200, 33].into_iter().enumerate() {
            let fc = u32::try_from(index).unwrap() * 10;
            builder.add_entry(fc, fc + 10, vec![0xA5; size]);
        }
        let pages = builder.generate_pages().unwrap();
        assert_eq!(
            pages.pages.len(),
            2,
            "the overflowing CHPX moves to a new page"
        );
        let mut entries = Vec::new();
        for page in &pages.pages {
            let parsed = crate::parts::fkp::ChpxFkp::parse(page, &[]).expect("valid CHPX FKP");
            entries.extend(parsed.entries().iter().map(|entry| entry.grpprl.len()));
        }
        assert_eq!(entries, [255, 200, 33]);
        assert_eq!(pages.ranges, [(0, 20), (20, 30)]);
    }

    /// A page whose estimate and exact placement both fit is unchanged.
    #[test]
    fn chpx_page_below_the_estimate_keeps_its_single_page_layout() {
        let mut builder = ChpxFkpBuilder::new();
        for (index, size) in [255usize, 200, 32].into_iter().enumerate() {
            let fc = u32::try_from(index).unwrap() * 10;
            builder.add_entry(fc, fc + 10, vec![0xA5; size]);
        }
        let pages = builder.generate_pages().unwrap();
        assert_eq!(pages.pages.len(), 1);
        let parsed = crate::parts::fkp::ChpxFkp::parse(&pages.pages[0], &[]).unwrap();
        assert_eq!(parsed.count(), 3);
    }

    /// Three PAPXs whose even-rounded estimate is one byte short of the page
    /// because the first placement below the count byte is realigned.
    #[test]
    fn papx_page_filled_to_the_estimate_never_overwrites_its_bx_array() {
        let mut builder = PapxFkpBuilder::new();
        for index in 0..3u32 {
            builder.add_entry(index * 10, index * 10 + 10, vec![0x5A; 148]);
        }
        let pages = builder.generate_pages().unwrap();
        assert_eq!(
            pages.pages.len(),
            2,
            "the overflowing PAPX moves to a new page"
        );
        let mut count = 0;
        for page in &pages.pages {
            let parsed = crate::parts::fkp::PapxFkp::parse(page, &[]).expect("valid PAPX FKP");
            for index in 0..parsed.count() {
                let entry = parsed.entry(index).unwrap();
                assert_eq!(entry.grpprl.len(), 150);
                assert_eq!(&entry.grpprl[2..], &[0x5A; 148]);
                count += 1;
            }
        }
        assert_eq!(count, 3);
        assert_eq!(pages.ranges, [(0, 20), (20, 30)]);
    }

    #[test]
    fn test_fkp_pages_collection() {
        let pages = FkpPages {
            pages: vec![vec![0u8; 512], vec![0u8; 512]],
            ranges: vec![(0, 1000), (1000, 2000)],
        };
        assert_eq!(pages.pages.len(), 2);
        assert_eq!(pages.ranges.len(), 2);
    }
}
