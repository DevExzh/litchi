//! DIFAT (Double Indirect FAT) generation for OLE2 files
//!
//! The DIFAT is used when the FAT requires more than 109 sectors.
//! The header can store the first 109 FAT sector IDs, but additional
//! FAT sectors need to be tracked in DIFAT sectors.
//!
//! # Implementation Notes
//!
//! Based on Apache POI's BATBlock implementation.
//! - Each DIFAT sector contains FAT sector IDs and a pointer to the next DIFAT sector
//! - For 512-byte sectors: 127 FAT sector IDs + 1 next pointer (128 * 4 = 512)
//! - For 4096-byte sectors: 1023 FAT sector IDs + 1 next pointer (1024 * 4 = 4096)

use super::super::consts::*;

/// DIFAT builder for large file support
///
/// Manages DIFAT sector allocation when more than 109 FAT sectors are needed.
///
/// # Performance Optimizations
///
/// - Pre-calculates DIFAT sector requirements
/// - Efficiently packs FAT sector IDs into DIFAT sectors
/// - Uses zero-copy where possible
#[derive(Debug)]
pub struct DifatBuilder {
    /// FAT sector IDs beyond the first 109
    fat_sector_ids: Vec<u32>,
    /// Sector size (512 or 4096 bytes)
    sector_size: usize,
}

#[allow(dead_code)] // These methods are part of the public API for future use
impl DifatBuilder {
    /// Create a new DIFAT builder
    ///
    /// # Arguments
    ///
    /// * `sector_size` - Sector size in bytes (512 or 4096)
    pub fn new(sector_size: usize) -> Self {
        assert!(
            sector_size == 512 || sector_size == 4096,
            "Sector size must be 512 or 4096"
        );

        Self {
            fat_sector_ids: Vec::new(),
            sector_size,
        }
    }

    /// Add FAT sector IDs beyond the first 109
    ///
    /// # Arguments
    ///
    /// * `fat_sectors` - Complete list of FAT sector IDs
    ///
    /// The first 109 IDs will be skipped (they go in the header),
    /// and the rest will be stored in DIFAT sectors.
    pub fn set_fat_sectors(&mut self, fat_sectors: &[u32]) {
        if fat_sectors.len() > 109 {
            self.fat_sector_ids = fat_sectors[109..].to_vec();
        } else {
            self.fat_sector_ids.clear();
        }
    }

    /// Calculate the number of DIFAT sectors needed
    ///
    /// # Returns
    ///
    /// * `u32` - Number of DIFAT sectors required
    pub fn calculate_difat_sector_count(&self) -> u32 {
        if self.fat_sector_ids.is_empty() {
            return 0;
        }

        // Each DIFAT sector can hold: (sector_size / 4) - 1 FAT sector IDs
        // The -1 is for the next DIFAT sector pointer
        let ids_per_difat_sector = (self.sector_size / 4) - 1;

        self.fat_sector_ids.len().div_ceil(ids_per_difat_sector) as u32
    }

    /// Generate DIFAT sectors as bytes
    ///
    /// # Arguments
    ///
    /// * `first_difat_sector` - Starting sector ID for DIFAT chain
    ///
    /// # Returns
    ///
    /// * `Vec<Vec<u8>>` - Vector of DIFAT sectors
    ///
    /// # Format
    ///
    /// Each DIFAT sector contains:
    /// - FAT sector IDs (as many as will fit)
    /// - Next DIFAT sector ID (or ENDOFCHAIN for last)
    /// - Padding with FREESECT
    pub fn generate_difat_sectors(&self, first_difat_sector: u32) -> Vec<Vec<u8>> {
        if self.fat_sector_ids.is_empty() {
            return Vec::new();
        }

        let ids_per_difat_sector = (self.sector_size / 4) - 1;
        let num_difat_sectors = self.calculate_difat_sector_count();

        let mut difat_sectors = Vec::with_capacity(num_difat_sectors as usize);

        for difat_idx in 0..num_difat_sectors {
            let mut sector_data = vec![0xFFu8; self.sector_size]; // Initialize with FREESECT

            let start_id_idx = (difat_idx as usize) * ids_per_difat_sector;
            let end_id_idx =
                ((difat_idx as usize + 1) * ids_per_difat_sector).min(self.fat_sector_ids.len());

            // Write FAT sector IDs
            for (i, &fat_sector_id) in self.fat_sector_ids[start_id_idx..end_id_idx]
                .iter()
                .enumerate()
            {
                let offset = i * 4;
                sector_data[offset..offset + 4].copy_from_slice(&fat_sector_id.to_le_bytes());
            }

            // Write next DIFAT sector pointer (last u32 in sector)
            let next_pointer_offset = self.sector_size - 4;
            let next_difat_sector = if difat_idx < num_difat_sectors - 1 {
                first_difat_sector + difat_idx + 1
            } else {
                ENDOFCHAIN
            };
            sector_data[next_pointer_offset..next_pointer_offset + 4]
                .copy_from_slice(&next_difat_sector.to_le_bytes());

            difat_sectors.push(sector_data);
        }

        difat_sectors
    }

    /// Check if DIFAT is needed (more than 109 FAT sectors)
    pub fn is_needed(&self) -> bool {
        !self.fat_sector_ids.is_empty()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_no_difat_needed() {
        let difat = DifatBuilder::new(512);
        assert!(!difat.is_needed());
        assert_eq!(difat.calculate_difat_sector_count(), 0);
    }

    #[test]
    fn test_difat_calculation() {
        let mut difat = DifatBuilder::new(512);

        // Create 150 FAT sector IDs (109 fit in header, 41 need DIFAT)
        let fat_sectors: Vec<u32> = (0..150).collect();
        difat.set_fat_sectors(&fat_sectors);

        assert!(difat.is_needed());

        // For 512-byte sectors: (512 / 4) - 1 = 127 IDs per DIFAT sector
        // 41 IDs need 1 DIFAT sector
        assert_eq!(difat.calculate_difat_sector_count(), 1);
    }

    #[test]
    fn test_difat_generation() {
        let mut difat = DifatBuilder::new(512);

        // Create 150 FAT sector IDs
        let fat_sectors: Vec<u32> = (0..150).collect();
        difat.set_fat_sectors(&fat_sectors);

        let sectors = difat.generate_difat_sectors(200);
        assert_eq!(sectors.len(), 1);
        assert_eq!(sectors[0].len(), 512);

        // Check that the last 4 bytes contain ENDOFCHAIN (since it's the only DIFAT sector)
        let last_u32_bytes = &sectors[0][508..512];
        let last_u32 = u32::from_le_bytes([
            last_u32_bytes[0],
            last_u32_bytes[1],
            last_u32_bytes[2],
            last_u32_bytes[3],
        ]);
        assert_eq!(last_u32, ENDOFCHAIN);
    }

    #[test]
    fn test_difat_multiple_sectors() {
        let mut difat = DifatBuilder::new(512);

        // Create 250 FAT sector IDs (109 in header, 141 in DIFAT)
        // 141 IDs need 2 DIFAT sectors (127 + 14)
        let fat_sectors: Vec<u32> = (0..250).collect();
        difat.set_fat_sectors(&fat_sectors);

        assert_eq!(difat.calculate_difat_sector_count(), 2);

        let sectors = difat.generate_difat_sectors(300);
        assert_eq!(sectors.len(), 2);

        // Check that first DIFAT sector points to second
        let next_pointer_offset = 512 - 4;
        let next_sector_bytes = &sectors[0][next_pointer_offset..next_pointer_offset + 4];
        let next_sector = u32::from_le_bytes([
            next_sector_bytes[0],
            next_sector_bytes[1],
            next_sector_bytes[2],
            next_sector_bytes[3],
        ]);
        assert_eq!(next_sector, 301); // 300 + 1

        // Check that second DIFAT sector has ENDOFCHAIN
        let last_bytes = &sectors[1][next_pointer_offset..next_pointer_offset + 4];
        let last_sector =
            u32::from_le_bytes([last_bytes[0], last_bytes[1], last_bytes[2], last_bytes[3]]);
        assert_eq!(last_sector, ENDOFCHAIN);
    }

    #[test]
    fn test_difat_4096_sectors() {
        let mut difat = DifatBuilder::new(4096);

        // For 4096-byte sectors: (4096 / 4) - 1 = 1023 IDs per DIFAT sector
        // Create 1200 FAT sector IDs (109 in header, 1091 in DIFAT)
        // 1091 IDs need 2 DIFAT sectors (1023 + 68)
        let fat_sectors: Vec<u32> = (0..1200).collect();
        difat.set_fat_sectors(&fat_sectors);

        assert_eq!(difat.calculate_difat_sector_count(), 2);
    }
}
