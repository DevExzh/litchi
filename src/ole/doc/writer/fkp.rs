//! FKP (Formatted Disk Page) generation for DOC files
//!
//! FKPs store character and paragraph formatting information in a compact,
//! page-based structure. Each FKP is 512 bytes and contains multiple formatting
//! entries (CHPXs for character properties, PAPXs for paragraph properties).
//!
//! Based on Microsoft's "[MS-DOC]" specification and Apache POI's CHPFormattedDiskPage.

/// Error type for DOC operations
pub type DocError = std::io::Error;

/// Character property (CHPX) entry
/// Represents a formatting run from fc_start to fc_end
#[derive(Debug, Clone)]
pub struct ChpxEntry {
    /// Start file character position (byte offset in WordDocument stream)
    pub fc_start: u32,
    /// End file character position (byte offset in WordDocument stream)
    pub fc_end: u32,
    /// Character properties (SPRM sequence)
    pub grpprl: Vec<u8>,
}

/// Paragraph property (PAPX) entry
/// Represents a formatting run from fc_start to fc_end
#[derive(Debug, Clone)]
pub struct PapxEntry {
    /// Start file character position (byte offset in WordDocument stream)
    pub fc_start: u32,
    /// End file character position (byte offset in WordDocument stream)
    pub fc_end: u32,
    /// Paragraph properties (SPRM sequence)
    pub grpprl: Vec<u8>,
}

/// Character FKP (CHPX FKP) builder
#[derive(Debug)]
pub struct ChpxFkpBuilder {
    entries: Vec<ChpxEntry>,
}

impl ChpxFkpBuilder {
    /// Create a new character FKP builder
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    /// Add a character formatting entry (a formatting run)
    ///
    /// # Arguments
    ///
    /// * `fc_start` - Start file character position (byte offset in WordDocument stream)
    /// * `fc_end` - End file character position (byte offset in WordDocument stream)
    /// * `grpprl` - Character properties (SPRM sequence)
    pub fn add_entry(&mut self, fc_start: u32, fc_end: u32, grpprl: Vec<u8>) {
        self.entries.push(ChpxEntry {
            fc_start,
            fc_end,
            grpprl,
        });
    }

    /// Generate the FKP as a 512-byte page
    ///
    /// Based on Apache POI's CHPFormattedDiskPage and MS-DOC ChpxFkp specification
    pub fn generate(&self) -> Result<Vec<u8>, DocError> {
        let mut fkp = vec![0u8; 512];

        // Maximum entries per CHPX FKP (based on MS-DOC)
        let entry_count = self.entries.len().min(101);

        if entry_count == 0 {
            return Ok(fkp);
        }

        // Write FC array (file character positions) - MS-DOC Section 2.9.86
        // FC array has (n+1) entries for n formatting runs
        // Following Apache POI CHPFormattedDiskPage.toByteArray() lines 177-193
        let mut fc_offset = 0;
        for entry in self.entries.iter().take(entry_count) {
            // Write START fc for this run
            fkp[fc_offset..fc_offset + 4].copy_from_slice(&entry.fc_start.to_le_bytes());
            fc_offset += 4;
        }
        // Write END fc of the last run (POI line 193)
        if let Some(last_entry) = self.entries.get(entry_count - 1) {
            fkp[fc_offset..fc_offset + 4].copy_from_slice(&last_entry.fc_end.to_le_bytes());
        }

        // Write entry count at byte 511
        fkp[511] = entry_count as u8;

        // Write RGB array (property offset array) - MS-DOC Section 2.9.86
        // RGB array starts after FC array and contains offsets to CHPXs
        // Each rgb entry is 1 byte containing the offset (in words) to the CHPX
        // If rgb[i] = 0, it means no formatting for that range (use default)
        let rgb_offset = (entry_count + 1) * 4; // After (n+1) FCs
        let mut data_offset = 511; // Start from end, before the count byte

        // Process entries in reverse order to fill from end of page backwards
        for (i, entry) in self.entries.iter().take(entry_count).enumerate().rev() {
            let prop_size = entry.grpprl.len();

            if prop_size == 0 {
                // No formatting - set rgb to 0
                fkp[rgb_offset + i] = 0;
            } else {
                // Allocate space for CHPX: 1 byte (size) + grpprl bytes
                let chpx_size = 1 + prop_size.min(255);
                data_offset -= chpx_size;
                // Align to word boundary as in POI (grpprlOffset -= grpprlOffset % 2)
                data_offset -= data_offset % 2;

                // Write CHPX: size byte followed by grpprl
                fkp[data_offset] = prop_size.min(255) as u8;
                fkp[data_offset + 1..data_offset + 1 + prop_size.min(255)]
                    .copy_from_slice(&entry.grpprl[..prop_size.min(255)]);

                // Write offset in RGB (offset in words, so divide by 2)
                fkp[rgb_offset + i] = (data_offset / 2) as u8;
            }
        }

        Ok(fkp)
    }
}

impl Default for ChpxFkpBuilder {
    fn default() -> Self {
        Self::new()
    }
}

/// Paragraph FKP (PAPX FKP) builder
#[derive(Debug)]
pub struct PapxFkpBuilder {
    entries: Vec<PapxEntry>,
}

const BX_SIZE: usize = 13;

impl PapxFkpBuilder {
    /// Create a new paragraph FKP builder
    pub fn new() -> Self {
        Self {
            entries: Vec::new(),
        }
    }

    /// Add a paragraph formatting entry (a formatting run)
    ///
    /// # Arguments
    ///
    /// * `fc_start` - Start file character position (byte offset in WordDocument stream)
    /// * `fc_end` - End file character position (byte offset in WordDocument stream)
    /// * `grpprl` - Paragraph properties (SPRM sequence)
    pub fn add_entry(&mut self, fc_start: u32, fc_end: u32, grpprl: Vec<u8>) {
        self.entries.push(PapxEntry {
            fc_start,
            fc_end,
            grpprl,
        });
    }

    /// Generate the FKP as a 512-byte page
    ///
    /// Based on Apache POI's PAPFormattedDiskPage and MS-DOC PapxFkp specification
    pub fn generate(&self) -> Result<Vec<u8>, DocError> {
        let mut fkp = vec![0u8; 512];

        // Maximum entries per PAPX FKP (based on MS-DOC)
        let entry_count = self.entries.len().min(29);

        if entry_count == 0 {
            return Ok(fkp);
        }

        // Write FC array (file character positions) - MS-DOC Section 2.9.179
        // FC array has (n+1) entries for n formatting runs
        // Following Apache POI PAPFormattedDiskPage.toByteArray() logic
        let mut fc_offset = 0;
        for entry in self.entries.iter().take(entry_count) {
            // Write START fc for this run
            fkp[fc_offset..fc_offset + 4].copy_from_slice(&entry.fc_start.to_le_bytes());
            fc_offset += 4;
        }
        // Write END fc of the last run
        if let Some(last_entry) = self.entries.get(entry_count - 1) {
            fkp[fc_offset..fc_offset + 4].copy_from_slice(&last_entry.fc_end.to_le_bytes());
        }

        // Write entry count at byte 511
        fkp[511] = entry_count as u8;

        // Write BX array (PAPX descriptors) - MS-DOC Section 2.9.179
        // BX array starts after FC array, each entry is 13 bytes
        let bx_offset = (entry_count + 1) * 4;
        let mut grpprl_offset = 511; // Start from end, before count byte

        // Process entries in reverse order to fill from end of page backwards
        // Implement PAPX layout exactly as in POI: grpprl includes istd (2 bytes) followed by SPRMs
        // For our entries, we synthesize istd=0 and prepend it to grpprl bytes
        for i in 0..entry_count {
            let entry = &self.entries[i];
            let mut grpprl_full = Vec::with_capacity(2 + entry.grpprl.len());
            grpprl_full.extend_from_slice(&0u16.to_le_bytes()); // istd = 0
            grpprl_full.extend_from_slice(&entry.grpprl);
            let len = grpprl_full.len();

            // Adjust grpprl_offset for this PAPX: reserve grpprl bytes plus 1 or 2 bytes for cb
            let extra = if (len % 2) > 0 { 1 } else { 2 };
            grpprl_offset -= len + extra;
            // Align to word boundary
            grpprl_offset -= grpprl_offset % 2;

            // Write BX entry: pointer to PAPX and PHE (zeros)
            let bx_pos = bx_offset + i * BX_SIZE;
            fkp[bx_pos] = (grpprl_offset / 2) as u8; // Offset in words
            // PHE (12 bytes) remain zeroed

            // Write cb and grpprl per POI rules
            let mut copy_offset = grpprl_offset;
            if (len % 2) > 0 {
                // odd length: write cb at current offset, then grpprl
                fkp[copy_offset] = len.div_ceil(2) as u8;
                copy_offset += 1;
            } else {
                // even length: skip 1 byte, write cb next, then grpprl
                fkp[copy_offset + 1] = (len / 2) as u8;
                copy_offset += 2;
            }
            fkp[copy_offset..copy_offset + len].copy_from_slice(&grpprl_full);
        }

        // Write END fc of the last run
        if let Some(last_entry) = self.entries.get(entry_count - 1) {
            fkp[fc_offset..fc_offset + 4].copy_from_slice(&last_entry.fc_end.to_le_bytes());
        }

        Ok(fkp)
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
        builder.add_entry(0, 100, vec![0x80, 0x00]); // One run from FC 0 to 100
        builder.add_entry(100, 200, vec![0x80, 0x01]); // Another run from 100 to 200

        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
        assert_eq!(fkp[511], 2); // 2 entries
    }

    #[test]
    fn test_papx_fkp() {
        let mut builder = PapxFkpBuilder::new();
        builder.add_entry(0, 100, vec![0x80, 0x00]); // One run from FC 0 to 100

        let fkp = builder.generate().unwrap();
        assert_eq!(fkp.len(), 512);
    }
}
