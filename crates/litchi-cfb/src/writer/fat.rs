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

use super::super::consts::{DIFSECT, ENDOFCHAIN, FATSECT, FREESECT, MAXREGSECT};
use super::super::file::OleError;

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
pub(super) struct FatBuilder {
    /// The FAT table (maps sector ID to next sector in chain)
    fat: Vec<u32>,
    /// Next available sector
    next_sector: u32,
    /// Sector size for this FAT
    sector_size: usize,
}

#[allow(
    dead_code,
    reason = "builder API kept complete for symmetry and future use"
)]
impl FatBuilder {
    /// Create a new FAT builder
    ///
    /// # Arguments
    ///
    /// * `sector_size` - Size of each sector in bytes (512 or 4096)
    pub(super) fn new_with_size(sector_size: usize) -> Result<Self, OleError> {
        if !matches!(sector_size, 512 | 4096) {
            return Err(OleError::InvalidData(format!(
                "CFB sector size must be 512 or 4096 bytes, got {sector_size}"
            )));
        }
        Ok(Self::with_valid_sector_size(sector_size))
    }

    fn with_valid_sector_size(sector_size: usize) -> Self {
        Self {
            fat: Vec::new(),
            next_sector: 0,
            sector_size,
        }
    }

    /// Create a new FAT builder with default 512-byte sectors
    pub(super) fn new() -> Self {
        Self::with_valid_sector_size(512)
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
    pub(super) fn allocate_chain(&mut self, size: usize) -> Result<u32, OleError> {
        if size == 0 {
            return Ok(ENDOFCHAIN);
        }

        let num_sectors = size.div_ceil(self.sector_size);
        let sector_count = u32::try_from(num_sectors).map_err(|_err| {
            OleError::InvalidData("CFB sector count exceeds MAXREGSECT".to_string())
        })?;
        let start_sector = self.next_sector;
        let end_sector = checked_sector_end(start_sector, sector_count)?;
        let new_len = usize::try_from(end_sector).map_err(|_err| {
            OleError::InvalidData("CFB FAT length does not fit usize".to_string())
        })?;

        if new_len > self.fat.len() {
            self.fat
                .try_reserve_exact(new_len - self.fat.len())
                .map_err(|source| OleError::allocation("FAT entries", source))?;
            self.fat.resize(new_len, FREESECT);
        }

        for current_sector in start_sector..end_sector {
            let next_value = if current_sector + 1 < end_sector {
                current_sector + 1
            } else {
                ENDOFCHAIN
            };
            self.fat[current_sector as usize] = next_value;
        }
        self.next_sector = end_sector;

        Ok(start_sector)
    }

    /// Allocate a single sector
    ///
    /// # Returns
    ///
    /// * `u32` - The allocated sector ID
    pub(super) fn allocate_sector(&mut self) -> Result<u32, OleError> {
        self.allocate_chain(self.sector_size)
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
    pub(super) fn allocate_special(&mut self, count: u32, marker: u32) -> Result<u32, OleError> {
        if count == 0 {
            return Ok(ENDOFCHAIN);
        }
        if !matches!(marker, FATSECT | DIFSECT) {
            return Err(OleError::InvalidData(
                "CFB special allocation requires FATSECT or DIFSECT".to_string(),
            ));
        }

        let start = self.next_sector;
        let end = checked_sector_end(start, count)?;

        let needed_len = usize::try_from(end).map_err(|_err| {
            OleError::InvalidData("CFB FAT length does not fit usize".to_string())
        })?;
        if self.fat.len() < needed_len {
            self.fat
                .try_reserve_exact(needed_len - self.fat.len())
                .map_err(|source| OleError::allocation("FAT entries", source))?;
            self.fat.resize(needed_len, FREESECT);
        }

        for s in start..end {
            self.fat[s as usize] = marker;
        }

        self.next_sector = end;
        Ok(start)
    }

    /// Mark a range of sectors as FAT sectors
    ///
    /// FAT sectors are marked with special value FATSECT in the FAT itself.
    pub(super) fn mark_fat_sectors(&mut self, start: u32, count: u32) -> Result<(), OleError> {
        if start != self.next_sector {
            return Err(OleError::InvalidData(
                "CFB FAT sectors must begin at the next free sector".to_string(),
            ));
        }
        self.allocate_special(count, FATSECT)?;
        Ok(())
    }

    /// Get the FAT table
    pub(super) fn fat(&self) -> &[u32] {
        &self.fat
    }

    /// Get the total number of sectors allocated
    pub(super) fn total_sectors(&self) -> u32 {
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
    pub(super) fn generate_fat_sectors(&self) -> Result<Vec<Vec<u8>>, OleError> {
        let entries_per_sector = self.sector_size / 4;
        let num_fat_sectors = self.fat.len().div_ceil(entries_per_sector);

        let mut fat_sectors = Vec::new();
        fat_sectors
            .try_reserve_exact(num_fat_sectors)
            .map_err(|source| OleError::allocation("serialized FAT sectors", source))?;

        for entries in self.fat.chunks(entries_per_sector) {
            let mut sector_data = filled_sector(self.sector_size, "serialized FAT sector")?;

            for (i, &fat_value) in entries.iter().enumerate() {
                let offset = i * 4;
                sector_data[offset..offset + 4].copy_from_slice(&fat_value.to_le_bytes());
            }

            fat_sectors.push(sector_data);
        }

        Ok(fat_sectors)
    }

    /// Calculate the number of FAT sectors needed for the current allocation
    ///
    /// This is used to determine how many sectors will be needed to store the FAT itself.
    ///
    /// # Returns
    ///
    /// * `usize` - Number of FAT sectors needed
    pub(super) fn calculate_fat_sector_count(&self) -> usize {
        let entries_per_sector = self.sector_size / 4;
        self.fat.len().div_ceil(entries_per_sector)
    }

    /// Validate the FAT for consistency
    ///
    /// Checks for invalid sector references and backward links, which would
    /// permit a cycle in a writer-produced chain.
    ///
    /// # Returns
    ///
    /// * `Result<(), OleError>` - Ok if valid, Err with description if invalid
    pub(super) fn validate(&self) -> Result<(), OleError> {
        for (current, &next) in self.fat.iter().enumerate() {
            match next {
                ENDOFCHAIN | FREESECT | FATSECT | DIFSECT => {},
                0..MAXREGSECT => {
                    let next_index = usize::try_from(next).map_err(|_err| {
                        OleError::InvalidData("CFB FAT reference does not fit usize".to_string())
                    })?;
                    if next_index >= self.fat.len() {
                        return Err(OleError::InvalidData(format!(
                            "invalid next FAT sector {next} at sector {current}"
                        )));
                    }
                    if next_index <= current {
                        return Err(OleError::InvalidData(format!(
                            "backward FAT reference from sector {current} to {next}"
                        )));
                    }
                },
                _ => {
                    return Err(OleError::InvalidData(format!(
                        "invalid FAT marker 0x{next:08X} at sector {current}"
                    )));
                },
            }
        }

        Ok(())
    }

    /// Get sector size
    pub(super) fn sector_size(&self) -> usize {
        self.sector_size
    }
}

impl Default for FatBuilder {
    fn default() -> Self {
        Self::new()
    }
}

fn checked_sector_end(start: u32, count: u32) -> Result<u32, OleError> {
    let end = start
        .checked_add(count)
        .ok_or_else(|| OleError::InvalidData("CFB sector count overflows u32".to_string()))?;
    if end > MAXREGSECT {
        return Err(OleError::InvalidData(
            "CFB sector count exceeds MAXREGSECT".to_string(),
        ));
    }
    Ok(end)
}

fn filled_sector(size: usize, resource: &'static str) -> Result<Vec<u8>, OleError> {
    let mut sector = Vec::new();
    sector
        .try_reserve_exact(size)
        .map_err(|source| OleError::allocation(resource, source))?;
    sector.resize(size, 0xff);
    Ok(sector)
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
    use super::*;

    #[test]
    fn test_allocate_chain() {
        let mut fat = FatBuilder::new();

        // Allocate 1024 bytes with 512-byte sectors (2 sectors)
        let start = fat.allocate_chain(1024).unwrap();
        assert_eq!(start, 0);
        assert_eq!(fat.total_sectors(), 2);

        // Check FAT entries
        assert_eq!(fat.fat()[0], 1); // First sector points to second
        assert_eq!(fat.fat()[1], ENDOFCHAIN); // Second sector is end
    }

    #[test]
    fn test_empty_chain() {
        let mut fat = FatBuilder::new();
        let start = fat.allocate_chain(0).unwrap();
        assert_eq!(start, ENDOFCHAIN);
        assert_eq!(fat.total_sectors(), 0);
    }

    #[test]
    fn test_mark_fat_sectors() {
        let mut fat = FatBuilder::new();
        fat.allocate_chain(512).unwrap(); // Allocate one sector
        fat.mark_fat_sectors(1, 2).unwrap(); // Mark sectors 1-2 as FAT

        assert_eq!(fat.fat()[1], FATSECT);
        assert_eq!(fat.fat()[2], FATSECT);
    }

    #[test]
    fn test_validate_good_fat() {
        let mut fat = FatBuilder::new();
        fat.allocate_chain(1024).unwrap();
        assert!(fat.validate().is_ok());
    }

    #[test]
    fn test_sector_size() {
        let fat_512 = FatBuilder::new();
        assert_eq!(fat_512.sector_size(), 512);

        let fat_4096 = FatBuilder::new_with_size(4096).unwrap();
        assert_eq!(fat_4096.sector_size(), 4096);
    }

    #[test]
    fn sector_limit_uses_maxregsect_as_an_exclusive_count() {
        assert_eq!(checked_sector_end(MAXREGSECT - 1, 1).unwrap(), MAXREGSECT);
        assert!(checked_sector_end(MAXREGSECT, 1).is_err());
        assert!(checked_sector_end(u32::MAX, 1).is_err());
    }

    #[test]
    fn invalid_sector_geometry_is_typed() {
        assert!(matches!(
            FatBuilder::new_with_size(1024),
            Err(OleError::InvalidData(_))
        ));
    }
}
