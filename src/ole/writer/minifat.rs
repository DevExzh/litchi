//! MiniFAT (Mini File Allocation Table) generation for OLE2 files
//!
//! The MiniFAT is used for small streams (< 4096 bytes) to avoid wasting
//! space in regular sectors. Small streams are stored in a ministream
//! (stored in the root entry), and the MiniFAT tracks mini sector allocation.
//!
//! # Implementation Notes
//!
//! Based on Apache POI's POIFSMiniStore implementation.
//! - Mini sectors are typically 64 bytes each
//! - Streams < 4096 bytes use the ministream
//! - The ministream itself is stored in regular sectors
//! - The MiniFAT is stored in regular sectors but tracks mini sectors

use super::super::consts::*;

/// MiniFAT builder for small stream allocation
///
/// Manages mini sector allocation and builds the Mini File Allocation Table
/// for small streams in an OLE compound document.
///
/// # Performance Optimizations
///
/// - Pre-allocates MiniFAT entries to avoid frequent reallocations
/// - Uses efficient mini sector chain building
/// - Tracks ministream size for efficient writing
#[derive(Debug)]
pub struct MiniFatBuilder {
    /// The MiniFAT table (maps mini sector ID to next mini sector in chain)
    minifat: Vec<u32>,
    /// Next available mini sector
    next_mini_sector: u32,
    /// Mini sector size (typically 64 bytes)
    mini_sector_size: usize,
    /// Ministream data (concatenated small streams)
    ministream_data: Vec<u8>,
}

#[allow(dead_code)] // These methods are part of the public API for future use
impl MiniFatBuilder {
    /// Create a new MiniFAT builder
    ///
    /// # Arguments
    ///
    /// * `mini_sector_size` - Size of each mini sector in bytes (typically 64)
    pub fn new(mini_sector_size: usize) -> Self {
        Self {
            minifat: Vec::new(),
            next_mini_sector: 0,
            mini_sector_size,
            ministream_data: Vec::new(),
        }
    }

    /// Allocate a chain of mini sectors for a small stream
    ///
    /// # Arguments
    ///
    /// * `data` - Stream data (must be < 4096 bytes)
    ///
    /// # Returns
    ///
    /// * `u32` - The starting mini sector of the allocated chain
    ///
    /// # Performance
    ///
    /// This method pre-allocates all MiniFAT entries and ministream space needed.
    pub fn allocate_mini_chain(&mut self, data: &[u8]) -> u32 {
        if data.is_empty() {
            return ENDOFCHAIN;
        }

        let num_mini_sectors = data.len().div_ceil(self.mini_sector_size);
        let start_mini_sector = self.next_mini_sector;

        // Pre-allocate MiniFAT entries to avoid resizing
        let new_size = (self.next_mini_sector as usize + num_mini_sectors).max(self.minifat.len());
        if new_size > self.minifat.len() {
            self.minifat.resize(new_size, FREESECT);
        }

        // Allocate mini sectors and link them
        for i in 0..num_mini_sectors {
            let current_mini_sector = self.next_mini_sector;
            self.next_mini_sector += 1;

            // Set MiniFAT entry
            let next_value = if i < num_mini_sectors - 1 {
                current_mini_sector + 1
            } else {
                ENDOFCHAIN
            };

            self.minifat[current_mini_sector as usize] = next_value;
        }

        // Add data to ministream (padded to mini sector boundary)
        let padded_size = num_mini_sectors * self.mini_sector_size;
        let current_offset = self.ministream_data.len();
        self.ministream_data.resize(current_offset + padded_size, 0);
        self.ministream_data[current_offset..current_offset + data.len()].copy_from_slice(data);

        start_mini_sector
    }

    /// Get the ministream data
    ///
    /// This data should be written to regular sectors and referenced
    /// from the root entry.
    pub fn ministream_data(&self) -> &[u8] {
        &self.ministream_data
    }

    /// Get the ministream size
    pub fn ministream_size(&self) -> u64 {
        self.ministream_data.len() as u64
    }

    /// Generate MiniFAT sectors as bytes
    ///
    /// # Arguments
    ///
    /// * `sector_size` - Regular sector size (512 or 4096 bytes)
    ///
    /// # Returns
    ///
    /// * `Vec<Vec<u8>>` - Vector of MiniFAT sectors
    pub fn generate_minifat_sectors(&self, sector_size: usize) -> Vec<Vec<u8>> {
        if self.minifat.is_empty() {
            return Vec::new();
        }

        let entries_per_sector = sector_size / 4;
        let num_minifat_sectors = self.minifat.len().div_ceil(entries_per_sector);

        let mut minifat_sectors = Vec::with_capacity(num_minifat_sectors);

        for sector_idx in 0..num_minifat_sectors {
            // Initialize with FREESECT (0xFFFFFFFF)
            let mut sector_data = vec![0xFFu8; sector_size];
            let start_entry = sector_idx * entries_per_sector;
            let end_entry = (start_entry + entries_per_sector).min(self.minifat.len());

            // Copy MiniFAT entries as little-endian u32 values
            for (i, &minifat_value) in self.minifat[start_entry..end_entry].iter().enumerate() {
                let offset = i * 4;
                sector_data[offset..offset + 4].copy_from_slice(&minifat_value.to_le_bytes());
            }

            minifat_sectors.push(sector_data);
        }

        minifat_sectors
    }

    /// Get the number of mini sectors allocated
    pub fn mini_sector_count(&self) -> u32 {
        self.next_mini_sector
    }

    /// Check if MiniFAT has any allocations
    pub fn is_empty(&self) -> bool {
        self.minifat.is_empty()
    }

    /// Get the MiniFAT table
    pub fn minifat(&self) -> &[u32] {
        &self.minifat
    }
}

impl Default for MiniFatBuilder {
    fn default() -> Self {
        Self::new(64) // Default mini sector size
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_allocate_mini_chain() {
        let mut minifat = MiniFatBuilder::new(64);

        // Allocate a small stream (100 bytes = 2 mini sectors of 64 bytes)
        let data = vec![0xAAu8; 100];
        let start = minifat.allocate_mini_chain(&data);

        assert_eq!(start, 0);
        assert_eq!(minifat.mini_sector_count(), 2);

        // Check MiniFAT entries
        assert_eq!(minifat.minifat()[0], 1); // First mini sector points to second
        assert_eq!(minifat.minifat()[1], ENDOFCHAIN); // Second mini sector is end

        // Check ministream data (should be padded to 128 bytes = 2 * 64)
        assert_eq!(minifat.ministream_size(), 128);
    }

    #[test]
    fn test_empty_mini_chain() {
        let mut minifat = MiniFatBuilder::new(64);
        let start = minifat.allocate_mini_chain(&[]);

        assert_eq!(start, ENDOFCHAIN);
        assert_eq!(minifat.mini_sector_count(), 0);
        assert!(minifat.is_empty());
    }

    #[test]
    fn test_multiple_allocations() {
        let mut minifat = MiniFatBuilder::new(64);

        let data1 = vec![0xAAu8; 50]; // 1 mini sector
        let data2 = vec![0xBBu8; 100]; // 2 mini sectors

        let start1 = minifat.allocate_mini_chain(&data1);
        let start2 = minifat.allocate_mini_chain(&data2);

        assert_eq!(start1, 0);
        assert_eq!(start2, 1); // Starts after first allocation
        assert_eq!(minifat.mini_sector_count(), 3);

        // First chain: sector 0 -> ENDOFCHAIN
        assert_eq!(minifat.minifat()[0], ENDOFCHAIN);

        // Second chain: sector 1 -> 2 -> ENDOFCHAIN
        assert_eq!(minifat.minifat()[1], 2);
        assert_eq!(minifat.minifat()[2], ENDOFCHAIN);
    }

    #[test]
    fn test_generate_minifat_sectors() {
        let mut minifat = MiniFatBuilder::new(64);

        // Allocate some mini sectors
        minifat.allocate_mini_chain(&[0u8; 100]);

        let sectors = minifat.generate_minifat_sectors(512);
        assert!(!sectors.is_empty());
        assert_eq!(sectors[0].len(), 512);
    }
}
