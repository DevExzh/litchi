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

use super::super::consts::{
    DIFSECT, ENDOFCHAIN, FATSECT, FREESECT, MAXREGSECT, RANGE_LOCK_SECTOR_V4, SECTOR_SIZE_V4,
    TWO_GIB_BYTES,
};
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
    /// Fixed range-lock sector for version 4 files.  The entry remains free
    /// until `finalize_range_lock` knows whether the physical output crosses
    /// the 2-GiB boundary.
    range_lock_sector: Option<u32>,
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
            range_lock_sector: (sector_size == SECTOR_SIZE_V4).then_some(RANGE_LOCK_SECTOR_V4),
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
        self.allocate_chain_u64(u64::try_from(size).map_err(|_err| {
            OleError::InvalidData("CFB stream size does not fit u64".to_string())
        })?)
    }

    /// Allocate a chain from a predeclared byte length without narrowing it
    /// through `usize` first.
    pub(super) fn allocate_chain_u64(&mut self, size: u64) -> Result<u32, OleError> {
        if size == 0 {
            return Ok(ENDOFCHAIN);
        }

        let num_sectors = size.div_ceil(u64::try_from(self.sector_size).map_err(|_err| {
            OleError::InvalidData("CFB sector size does not fit u64".to_string())
        })?);
        let sector_count = u32::try_from(num_sectors).map_err(|_err| {
            OleError::InvalidData("CFB sector count exceeds MAXREGSECT".to_string())
        })?;
        let start_sector = self.next_sector;
        let allocation_start = if self.is_range_lock(start_sector) {
            start_sector.checked_add(1).ok_or_else(|| {
                OleError::InvalidData("CFB range-lock sector overflows u32".to_string())
            })?
        } else {
            start_sector
        };
        let end_sector = self.physical_end_after(start_sector, sector_count)?;
        let new_len = usize::try_from(end_sector).map_err(|_err| {
            OleError::InvalidData("CFB FAT length does not fit usize".to_string())
        })?;

        if new_len > self.fat.len() {
            self.fat
                .try_reserve_exact(new_len - self.fat.len())
                .map_err(|source| OleError::allocation("FAT entries", source))?;
            self.fat.resize(new_len, FREESECT);
        }

        let mut current_sector = start_sector;
        let mut remaining = sector_count;
        while remaining != 0 {
            // The range-lock sector is a physical hole.  Consume its slot in
            // the output layout without assigning it to a user chain.
            if self.is_range_lock(current_sector) {
                current_sector = current_sector.checked_add(1).ok_or_else(|| {
                    OleError::InvalidData("CFB range-lock sector overflows u32".to_string())
                })?;
                continue;
            }
            let mut next_sector = current_sector.checked_add(1).ok_or_else(|| {
                OleError::InvalidData("CFB sector chain overflows u32".to_string())
            })?;
            if self.is_range_lock(next_sector) {
                next_sector = next_sector.checked_add(1).ok_or_else(|| {
                    OleError::InvalidData("CFB range-lock sector overflows u32".to_string())
                })?;
            }
            self.fat[current_sector as usize] = if remaining == 1 {
                ENDOFCHAIN
            } else {
                next_sector
            };
            current_sector = next_sector;
            remaining -= 1;
        }
        self.next_sector = end_sector;

        Ok(allocation_start)
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
        let allocation_start = if self.is_range_lock(start) {
            start.checked_add(1).ok_or_else(|| {
                OleError::InvalidData("CFB range-lock sector overflows u32".to_string())
            })?
        } else {
            start
        };
        let end = self.physical_end_after(start, count)?;

        let needed_len = usize::try_from(end).map_err(|_err| {
            OleError::InvalidData("CFB FAT length does not fit usize".to_string())
        })?;
        if self.fat.len() < needed_len {
            self.fat
                .try_reserve_exact(needed_len - self.fat.len())
                .map_err(|source| OleError::allocation("FAT entries", source))?;
            self.fat.resize(needed_len, FREESECT);
        }

        let mut current = start;
        let mut remaining = count;
        while remaining != 0 {
            if self.is_range_lock(current) {
                current = current.checked_add(1).ok_or_else(|| {
                    OleError::InvalidData("CFB range-lock sector overflows u32".to_string())
                })?;
                continue;
            }
            self.fat[current as usize] = marker;
            current = current.checked_add(1).ok_or_else(|| {
                OleError::InvalidData("CFB special sector range overflows u32".to_string())
            })?;
            remaining -= 1;
        }

        self.next_sector = end;
        Ok(allocation_start)
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

    /// Returns the fixed range-lock sector, when this is a version 4 FAT.
    pub(super) fn range_lock_sector(&self) -> Option<u32> {
        self.range_lock_sector
    }

    /// Finalize the fixed range-lock entry after every allocation is known.
    ///
    /// A file that crosses 2 GiB marks the sector `ENDOFCHAIN`; a file that
    /// merely reaches the location leaves it `FREESECT`, matching the
    /// shrink-back rule in MS-CFB 2.8.  The sector is never returned by an
    /// allocation method.
    pub(super) fn finalize_range_lock(&mut self) -> Result<(), OleError> {
        let Some(lock) = self.range_lock_sector else {
            return Ok(());
        };
        if self.next_sector <= lock {
            return Ok(());
        }
        let index = usize::try_from(lock).map_err(|_err| {
            OleError::InvalidData("CFB range-lock index does not fit usize".to_string())
        })?;
        let entry = self.fat.get_mut(index).ok_or_else(|| {
            OleError::InvalidData("CFB range-lock sector is missing from the FAT".to_string())
        })?;
        let bytes = (u64::from(self.next_sector) + 1)
            .checked_mul(u64::try_from(self.sector_size).map_err(|_err| {
                OleError::InvalidData("CFB sector size does not fit u64".to_string())
            })?)
            .ok_or_else(|| OleError::InvalidData("CFB output size overflows u64".to_string()))?;
        *entry = if bytes > TWO_GIB_BYTES {
            ENDOFCHAIN
        } else {
            FREESECT
        };
        Ok(())
    }

    /// Return physical IDs for an allocation of `count` logical sectors.
    ///
    /// FAT and DIFAT sectors are marked independently rather than linked, so
    /// this helper skips the fixed range-lock hole while preserving their
    /// allocation order.
    pub(super) fn allocation_sector_ids(
        &self,
        start: u32,
        count: u32,
    ) -> Result<Vec<u32>, OleError> {
        if count == 0 {
            return Ok(Vec::new());
        }
        if start >= self.next_sector || start == ENDOFCHAIN {
            return Err(OleError::InvalidData(
                "CFB allocation starts outside the FAT".to_string(),
            ));
        }
        let capacity = usize::try_from(count).map_err(|_err| {
            OleError::InvalidData("CFB allocation count does not fit usize".to_string())
        })?;
        let mut ids = Vec::new();
        ids.try_reserve_exact(capacity)
            .map_err(|source| OleError::allocation("CFB allocation sector IDs", source))?;
        let mut current = start;
        while ids.len() < capacity {
            if self.is_range_lock(current) {
                current = current.checked_add(1).ok_or_else(|| {
                    OleError::InvalidData("CFB range-lock sector overflows u32".to_string())
                })?;
                continue;
            }
            if current >= self.next_sector {
                return Err(OleError::InvalidData(
                    "CFB allocation sector range exceeds the FAT".to_string(),
                ));
            }
            ids.push(current);
            current = current.checked_add(1).ok_or_else(|| {
                OleError::InvalidData("CFB allocation sector range overflows u32".to_string())
            })?;
        }
        Ok(ids)
    }

    fn is_range_lock(&self, sector: u32) -> bool {
        self.range_lock_sector == Some(sector)
    }

    fn physical_end_after(&self, start: u32, count: u32) -> Result<u32, OleError> {
        let end = checked_sector_end(start, count)?;
        let crosses_lock = self
            .range_lock_sector
            .is_some_and(|lock| start <= lock && lock < end);
        let end = if crosses_lock {
            end.checked_add(1).ok_or_else(|| {
                OleError::InvalidData("CFB range-lock sector causes sector overflow".to_string())
            })?
        } else {
            end
        };
        checked_sector_end(
            start,
            end.checked_sub(start)
                .ok_or_else(|| OleError::InvalidData("CFB sector range underflows".to_string()))?,
        )
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
    fn version_four_range_lock_is_reserved_and_chain_crossing_skips_it() {
        let mut fat = FatBuilder::new_with_size(SECTOR_SIZE_V4).unwrap();
        let before_lock = u64::from(RANGE_LOCK_SECTOR_V4) * u64::try_from(SECTOR_SIZE_V4).unwrap();
        assert_eq!(fat.allocate_chain_u64(before_lock).unwrap(), 0);
        assert_eq!(fat.total_sectors(), RANGE_LOCK_SECTOR_V4);

        let crossing = fat.allocate_chain_u64(SECTOR_SIZE_V4 as u64).unwrap();
        assert_eq!(crossing, RANGE_LOCK_SECTOR_V4 + 1);
        assert_eq!(fat.total_sectors(), RANGE_LOCK_SECTOR_V4 + 2);
        assert_eq!(
            fat.fat()[usize::try_from(RANGE_LOCK_SECTOR_V4 - 1).unwrap()],
            ENDOFCHAIN
        );
        assert_eq!(
            fat.fat()[usize::try_from(RANGE_LOCK_SECTOR_V4).unwrap()],
            FREESECT
        );
        assert_eq!(fat.fat()[usize::try_from(crossing).unwrap()], ENDOFCHAIN);

        fat.finalize_range_lock().unwrap();
        assert_eq!(
            fat.fat()[usize::try_from(RANGE_LOCK_SECTOR_V4).unwrap()],
            ENDOFCHAIN
        );
        assert!(!fat.fat().contains(&RANGE_LOCK_SECTOR_V4));
        assert_eq!(
            fat.allocation_sector_ids(RANGE_LOCK_SECTOR_V4 - 1, 2)
                .unwrap(),
            [RANGE_LOCK_SECTOR_V4 - 1, crossing]
        );
        fat.validate().unwrap();
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
