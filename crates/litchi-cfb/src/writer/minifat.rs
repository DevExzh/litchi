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
use super::super::file::OleError;

const MINI_STREAM_CUTOFF: usize = 4096;

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
pub(super) struct MiniFatBuilder {
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
    pub(super) fn new(mini_sector_size: usize) -> Self {
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
    pub(super) fn allocate_mini_chain(&mut self, data: &[u8]) -> Result<u32, OleError> {
        if data.is_empty() {
            return Ok(ENDOFCHAIN);
        }
        // MS-CFB 2.2/2.4 require streams at or above the 4,096-byte cutoff
        // to use regular FAT sectors, never MiniFAT sectors.
        if data.len() >= MINI_STREAM_CUTOFF {
            return Err(OleError::InvalidData(
                "CFB streams at or above the MiniFAT cutoff must use regular FAT sectors"
                    .to_string(),
            ));
        }
        if self.mini_sector_size == 0 {
            return Err(OleError::InvalidData(
                "CFB mini sector size must be nonzero".to_string(),
            ));
        }

        let num_mini_sectors = data.len().div_ceil(self.mini_sector_size);
        let sector_count = u32::try_from(num_mini_sectors)
            .map_err(|_| OleError::InvalidData("CFB mini sector count exceeds u32".to_string()))?;
        let end_mini_sector = checked_end(self.next_mini_sector, sector_count)?;

        let start_mini_sector = self.next_mini_sector;
        let start_index = usize::try_from(start_mini_sector).map_err(|_| {
            OleError::InvalidData("CFB mini sector index does not fit usize".to_string())
        })?;
        let new_minifat_len = usize::try_from(end_mini_sector).map_err(|_| {
            OleError::InvalidData("CFB MiniFAT length does not fit usize".to_string())
        })?;
        let padded_size = num_mini_sectors
            .checked_mul(self.mini_sector_size)
            .ok_or_else(|| {
                OleError::InvalidData("CFB ministream size overflows usize".to_string())
            })?;
        let current_offset = self.ministream_data.len();
        let new_ministream_len = current_offset.checked_add(padded_size).ok_or_else(|| {
            OleError::InvalidData("CFB ministream offset overflows usize".to_string())
        })?;
        let data_end = current_offset.checked_add(data.len()).ok_or_else(|| {
            OleError::InvalidData("CFB ministream data range overflows usize".to_string())
        })?;

        // Reserve both buffers before mutating either chain, so an allocation
        // failure cannot leave partially committed MiniFAT metadata.
        if new_minifat_len > self.minifat.len() {
            self.minifat
                .try_reserve_exact(new_minifat_len - self.minifat.len())
                .map_err(|source| OleError::allocation("MiniFAT entries", source))?;
        }
        if new_ministream_len > self.ministream_data.len() {
            self.ministream_data
                .try_reserve_exact(new_ministream_len - self.ministream_data.len())
                .map_err(|source| OleError::allocation("ministream data", source))?;
        }
        self.minifat.resize(new_minifat_len, FREESECT);
        self.ministream_data.resize(new_ministream_len, 0);

        // Allocate mini sectors and link them
        let last_mini_sector = end_mini_sector - 1;
        for (index, current_mini_sector) in
            (start_index..new_minifat_len).zip(start_mini_sector..end_mini_sector)
        {
            let next_value = if current_mini_sector != last_mini_sector {
                current_mini_sector + 1
            } else {
                ENDOFCHAIN
            };
            self.minifat[index] = next_value;
        }
        self.next_mini_sector = end_mini_sector;

        // Add data to ministream (padded to mini sector boundary)
        let destination = self
            .ministream_data
            .get_mut(current_offset..data_end)
            .ok_or_else(|| {
                OleError::InvalidData("CFB ministream destination is unavailable".to_string())
            })?;
        destination.copy_from_slice(data);

        Ok(start_mini_sector)
    }

    /// Get the ministream data
    ///
    /// This data should be written to regular sectors and referenced
    /// from the root entry.
    pub(super) fn ministream_data(&self) -> &[u8] {
        &self.ministream_data
    }

    /// Get the ministream size
    pub(super) fn ministream_size(&self) -> Result<u64, OleError> {
        u64::try_from(self.ministream_data.len())
            .map_err(|_| OleError::InvalidData("CFB ministream size does not fit u64".to_string()))
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
    pub(super) fn generate_minifat_sectors(
        &self,
        sector_size: usize,
    ) -> Result<Vec<Vec<u8>>, OleError> {
        if !matches!(sector_size, 512 | 4096) {
            return Err(OleError::InvalidData(format!(
                "CFB sector size must be 512 or 4096 bytes, got {sector_size}"
            )));
        }
        if self.minifat.is_empty() {
            return Ok(Vec::new());
        }

        let entries_per_sector = sector_size / 4;
        let num_minifat_sectors = self.minifat.len().div_ceil(entries_per_sector);

        let mut minifat_sectors = Vec::new();
        minifat_sectors
            .try_reserve_exact(num_minifat_sectors)
            .map_err(|source| OleError::allocation("serialized MiniFAT sectors", source))?;

        for entries in self.minifat.chunks(entries_per_sector) {
            let mut sector_data = Vec::new();
            sector_data
                .try_reserve_exact(sector_size)
                .map_err(|source| OleError::allocation("serialized MiniFAT sector", source))?;
            sector_data.resize(sector_size, 0xff);

            for (i, &minifat_value) in entries.iter().enumerate() {
                let offset = i * 4;
                sector_data[offset..offset + 4].copy_from_slice(&minifat_value.to_le_bytes());
            }

            minifat_sectors.push(sector_data);
        }

        Ok(minifat_sectors)
    }

    /// Get the number of mini sectors allocated
    pub(super) fn mini_sector_count(&self) -> u32 {
        self.next_mini_sector
    }

    /// Check if MiniFAT has any allocations
    pub(super) fn is_empty(&self) -> bool {
        self.minifat.is_empty()
    }

    /// Get the MiniFAT table
    pub(super) fn minifat(&self) -> &[u32] {
        &self.minifat
    }
}

impl Default for MiniFatBuilder {
    fn default() -> Self {
        Self::new(64) // Default mini sector size
    }
}

fn checked_end(start: u32, count: u32) -> Result<u32, OleError> {
    let end = start
        .checked_add(count)
        .ok_or_else(|| OleError::InvalidData("CFB mini sector index overflows u32".to_string()))?;
    // `MAXREGSECT` is the first reserved marker value, so it is also the
    // exclusive upper bound for the count of zero-based mini sectors.
    if end > MAXREGSECT {
        return Err(OleError::InvalidData(
            "CFB mini sector index exceeds MAXREGSECT".to_string(),
        ));
    }
    Ok(end)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_allocate_mini_chain() {
        let mut minifat = MiniFatBuilder::new(64);

        // Allocate a small stream (100 bytes = 2 mini sectors of 64 bytes)
        let data = vec![0xAAu8; 100];
        let start = minifat.allocate_mini_chain(&data).unwrap();

        assert_eq!(start, 0);
        assert_eq!(minifat.mini_sector_count(), 2);

        // Check MiniFAT entries
        assert_eq!(minifat.minifat()[0], 1); // First mini sector points to second
        assert_eq!(minifat.minifat()[1], ENDOFCHAIN); // Second mini sector is end

        // Check ministream data (should be padded to 128 bytes = 2 * 64)
        assert_eq!(minifat.ministream_size().unwrap(), 128);
    }

    #[test]
    fn test_empty_mini_chain() {
        let mut minifat = MiniFatBuilder::new(64);
        let start = minifat.allocate_mini_chain(&[]).unwrap();

        assert_eq!(start, ENDOFCHAIN);
        assert_eq!(minifat.mini_sector_count(), 0);
        assert!(minifat.is_empty());
    }

    #[test]
    fn allocation_rejects_invalid_sector_geometry_before_mutation() {
        let mut zero_sized = MiniFatBuilder::new(0);
        assert!(zero_sized.allocate_mini_chain(&[1]).is_err());
        assert!(zero_sized.is_empty());

        let mut exhausted = MiniFatBuilder::new(64);
        exhausted.next_mini_sector = MAXREGSECT;
        assert!(exhausted.allocate_mini_chain(&[1]).is_err());
        assert!(exhausted.is_empty());

        assert_eq!(checked_end(MAXREGSECT - 1, 1).unwrap(), MAXREGSECT);
        assert!(checked_end(MAXREGSECT, 1).is_err());
    }

    #[test]
    fn allocation_rejects_ministream_cutoff_before_mutation() {
        let mut minifat = MiniFatBuilder::new(64);
        minifat.allocate_mini_chain(&[0xAA]).unwrap();
        let prior_minifat = minifat.minifat().to_vec();
        let prior_ministream = minifat.ministream_data().to_vec();
        let prior_count = minifat.mini_sector_count();

        assert!(
            minifat
                .allocate_mini_chain(&[0u8; MINI_STREAM_CUTOFF])
                .is_err()
        );

        assert_eq!(minifat.minifat(), prior_minifat.as_slice());
        assert_eq!(minifat.ministream_data(), prior_ministream.as_slice());
        assert_eq!(minifat.mini_sector_count(), prior_count);
    }

    #[test]
    fn test_multiple_allocations() {
        let mut minifat = MiniFatBuilder::new(64);

        let data1 = vec![0xAAu8; 50]; // 1 mini sector
        let data2 = vec![0xBBu8; 100]; // 2 mini sectors

        let start1 = minifat.allocate_mini_chain(&data1).unwrap();
        let start2 = minifat.allocate_mini_chain(&data2).unwrap();

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
        minifat.allocate_mini_chain(&[0u8; 100]).unwrap();

        let sectors = minifat.generate_minifat_sectors(512).unwrap();
        assert!(!sectors.is_empty());
        assert_eq!(sectors[0].len(), 512);
    }

    #[test]
    fn serialized_minifat_rejects_invalid_sector_geometry() {
        let mut minifat = MiniFatBuilder::new(64);
        minifat.allocate_mini_chain(&[1]).unwrap();
        assert!(matches!(
            minifat.generate_minifat_sectors(0),
            Err(OleError::InvalidData(_))
        ));
    }

    #[test]
    fn empty_minifat_rejects_invalid_sector_geometry() {
        let minifat = MiniFatBuilder::new(64);
        assert!(matches!(
            minifat.generate_minifat_sectors(0),
            Err(OleError::InvalidData(_))
        ));
    }
}
