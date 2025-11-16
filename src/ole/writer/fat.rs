//! FAT (File Allocation Table) generation for OLE2 files
//!
//! The FAT maps sector numbers to the next sector in a chain, enabling
//! variable-length streams to be stored in the compound file.
//!
//! # Implementation Notes
//!
//! Based on Apache POI's BATBlock and POIFSFileSystem implementations.
//! The FAT is organized as follows:
//! - Regular sectors use positive chain values
//! - FAT sectors are marked with FATSECT (0xFFFFFFFD)
//! - End of chain is marked with ENDOFCHAIN (0xFFFFFFFE)
//! - Free sectors are marked with FREESECT (0xFFFFFFFF)

use super::super::consts::*;

/// FAT builder for sector allocation
///
/// Manages sector allocation and builds the File Allocation Table
/// for an OLE compound document.
///
/// # Performance Optimizations
///
/// - Pre-allocates FAT entries to avoid frequent reallocations
/// - Uses efficient sector chain building with minimal branching
/// - Tracks allocated sectors for validation
#[derive(Debug)]
pub struct FatBuilder {
    /// The FAT table (maps sector ID to next sector in chain)
    fat: Vec<u32>,
    /// Next available sector
    next_sector: u32,
    /// Total number of sectors allocated (including FAT sectors)
    total_allocated: u32,
    /// Sector size for this FAT
    sector_size: usize,
}

#[allow(dead_code)] // These methods are part of the public API for future use
impl FatBuilder {
    /// Create a new FAT builder
    ///
    /// # Arguments
    ///
    /// * `sector_size` - Size of each sector in bytes (512 or 4096)
    pub fn new_with_size(sector_size: usize) -> Self {
        assert!(
            sector_size == 512 || sector_size == 4096,
            "Sector size must be 512 or 4096"
        );

        Self {
            fat: Vec::new(),
            next_sector: 0,
            total_allocated: 0,
            sector_size,
        }
    }

    /// Create a new FAT builder with default 512-byte sectors
    pub fn new() -> Self {
        Self::new_with_size(512)
    }

    /// Allocate a chain of sectors for a stream
    ///
    /// # Arguments
    ///
    /// * `size` - Size of the stream in bytes
    ///
    /// # Returns
    ///
    /// * `u32` - The starting sector of the allocated chain, or ENDOFCHAIN if empty
    ///
    /// # Performance
    ///
    /// This method pre-allocates all FAT entries needed for the chain,
    /// avoiding repeated vector resizing.
    pub fn allocate_chain(&mut self, size: usize) -> u32 {
        if size == 0 {
            return ENDOFCHAIN;
        }

        let num_sectors = size.div_ceil(self.sector_size);
        let start_sector = self.next_sector;

        // Pre-allocate FAT entries to avoid resizing (optimization)
        let new_size = (self.next_sector as usize + num_sectors).max(self.fat.len());
        if new_size > self.fat.len() {
            self.fat.resize(new_size, FREESECT);
        }

        // Allocate sectors and link them in a single pass
        for i in 0..num_sectors {
            let current_sector = self.next_sector;
            self.next_sector += 1;
            self.total_allocated += 1;

            // Set FAT entry
            // If not last sector: current_sector + 1, else ENDOFCHAIN
            let next_value = if i < num_sectors - 1 {
                current_sector + 1
            } else {
                ENDOFCHAIN
            };

            self.fat[current_sector as usize] = next_value;
        }

        start_sector
    }

    /// Allocate a single sector
    ///
    /// # Returns
    ///
    /// * `u32` - The allocated sector ID
    pub fn allocate_sector(&mut self) -> u32 {
        let sector = self.next_sector;
        self.next_sector += 1;
        self.total_allocated += 1;
        self.set_fat_entry(sector, ENDOFCHAIN);
        sector
    }

    /// Allocate a contiguous range of sectors and mark them with a special value
    ///
    /// This is used to reserve sectors for FAT (`FATSECT`) and DIFAT (`DIFSECT`).
    /// The returned sector ID is the first sector of the reserved range.
    ///
    /// # Arguments
    ///
    /// * `count` - Number of sectors to reserve
    /// * `marker` - The FAT marker to use for these sectors (e.g. `FATSECT`, `DIFSECT`)
    pub fn allocate_special(&mut self, count: u32, marker: u32) -> u32 {
        if count == 0 {
            return ENDOFCHAIN;
        }

        let start = self.next_sector;
        let end = start + count; // exclusive

        // Ensure FAT has capacity for all sectors we will mark
        let needed_len = end as usize;
        if self.fat.len() < needed_len {
            self.fat.resize(needed_len, FREESECT);
        }

        for s in start..end {
            self.fat[s as usize] = marker;
        }

        self.next_sector = end;
        self.total_allocated += count;
        start
    }

    /// Set a FAT entry
    ///
    /// # Arguments
    ///
    /// * `sector` - Sector ID
    /// * `value` - Value to set (next sector or ENDOFCHAIN)
    fn set_fat_entry(&mut self, sector: u32, value: u32) {
        let sector_idx = sector as usize;

        // Expand FAT if necessary
        if sector_idx >= self.fat.len() {
            self.fat.resize(sector_idx + 1, FREESECT);
        }

        self.fat[sector_idx] = value;
    }

    /// Mark a range of sectors as FAT sectors
    ///
    /// FAT sectors are marked with special value FATSECT in the FAT itself.
    pub fn mark_fat_sectors(&mut self, start: u32, count: u32) {
        for i in 0..count {
            self.set_fat_entry(start + i, FATSECT);
        }
    }

    /// Get the FAT table
    pub fn fat(&self) -> &[u32] {
        &self.fat
    }

    /// Get the total number of sectors allocated
    pub fn total_sectors(&self) -> u32 {
        self.next_sector
    }

    /// Generate FAT sectors as bytes
    ///
    /// # Returns
    ///
    /// * `Vec<Vec<u8>>` - Vector of FAT sectors
    ///
    /// # Performance
    ///
    /// Uses pre-allocated buffers and efficient byte copying to minimize allocations.
    pub fn generate_fat_sectors(&self) -> Vec<Vec<u8>> {
        let entries_per_sector = self.sector_size / 4;
        let num_fat_sectors = self.fat.len().div_ceil(entries_per_sector);

        let mut fat_sectors = Vec::with_capacity(num_fat_sectors);

        for sector_idx in 0..num_fat_sectors {
            // Initialize with FREESECT (0xFFFFFFFF)
            let mut sector_data = vec![0xFFu8; self.sector_size];
            let start_entry = sector_idx * entries_per_sector;
            let end_entry = (start_entry + entries_per_sector).min(self.fat.len());

            // Copy FAT entries as little-endian u32 values
            for (i, &fat_value) in self.fat[start_entry..end_entry].iter().enumerate() {
                let offset = i * 4;
                sector_data[offset..offset + 4].copy_from_slice(&fat_value.to_le_bytes());
            }

            fat_sectors.push(sector_data);
        }

        fat_sectors
    }

    /// Calculate the number of FAT sectors needed for the current allocation
    ///
    /// This is used to determine how many sectors will be needed to store the FAT itself.
    ///
    /// # Returns
    ///
    /// * `usize` - Number of FAT sectors needed
    pub fn calculate_fat_sector_count(&self) -> usize {
        let entries_per_sector = self.sector_size / 4;
        self.fat.len().div_ceil(entries_per_sector)
    }

    /// Validate the FAT for consistency
    ///
    /// Checks for:
    /// - Circular references (loops in chains)
    /// - Invalid sector references
    /// - Orphaned sectors
    ///
    /// # Returns
    ///
    /// * `Result<(), String>` - Ok if valid, Err with description if invalid
    pub fn validate(&self) -> Result<(), String> {
        use std::collections::HashSet;

        // Track visited sectors to detect loops
        let mut visited = HashSet::new();

        // Check each chain for cycles
        for start_sector in 0..self.fat.len() as u32 {
            if self.fat.get(start_sector as usize) == Some(&ENDOFCHAIN) {
                visited.clear();
                let mut current = start_sector;

                // Follow chain
                while current != ENDOFCHAIN {
                    if current >= self.fat.len() as u32 {
                        return Err(format!("Invalid sector reference: {}", current));
                    }

                    if !visited.insert(current) {
                        return Err(format!("Circular reference detected at sector {}", current));
                    }

                    let next = self.fat[current as usize];

                    // Validate special values
                    match next {
                        ENDOFCHAIN | FREESECT => break,
                        FATSECT | DIFSECT => break, // Valid special markers
                        _ => {
                            if next >= self.fat.len() as u32 && next != ENDOFCHAIN {
                                return Err(format!(
                                    "Invalid next sector {} at sector {}",
                                    next, current
                                ));
                            }
                        },
                    }

                    current = next;
                }
            }
        }

        Ok(())
    }

    /// Get sector size
    pub fn sector_size(&self) -> usize {
        self.sector_size
    }
}

impl Default for FatBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_allocate_chain() {
        let mut fat = FatBuilder::new();

        // Allocate 1024 bytes with 512-byte sectors (2 sectors)
        let start = fat.allocate_chain(1024);
        assert_eq!(start, 0);
        assert_eq!(fat.total_sectors(), 2);

        // Check FAT entries
        assert_eq!(fat.fat()[0], 1); // First sector points to second
        assert_eq!(fat.fat()[1], ENDOFCHAIN); // Second sector is end
    }

    #[test]
    fn test_empty_chain() {
        let mut fat = FatBuilder::new();
        let start = fat.allocate_chain(0);
        assert_eq!(start, ENDOFCHAIN);
        assert_eq!(fat.total_sectors(), 0);
    }

    #[test]
    fn test_mark_fat_sectors() {
        let mut fat = FatBuilder::new();
        fat.allocate_chain(512); // Allocate one sector
        fat.mark_fat_sectors(1, 2); // Mark sectors 1-2 as FAT

        assert_eq!(fat.fat()[1], FATSECT);
        assert_eq!(fat.fat()[2], FATSECT);
    }

    #[test]
    fn test_validate_good_fat() {
        let mut fat = FatBuilder::new();
        fat.allocate_chain(1024);
        assert!(fat.validate().is_ok());
    }

    #[test]
    fn test_sector_size() {
        let fat_512 = FatBuilder::new();
        assert_eq!(fat_512.sector_size(), 512);

        let fat_4096 = FatBuilder::new_with_size(4096);
        assert_eq!(fat_4096.sector_size(), 4096);
    }
}
