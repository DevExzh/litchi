//! OLE2 header generation
//!
//! Generates the 512-byte OLE2 file header with proper magic bytes,
//! version information, and FAT/directory locations.

use super::super::consts::*;

/// OLE2 header builder
pub struct HeaderBuilder {
    /// Sector size (512 or 4096)
    sector_size: usize,
    /// First sector of directory stream
    first_dir_sector: u32,
    /// Number of directory sectors (csectDir). For 512-byte sector files, must be 0.
    num_dir_sectors: u32,
    /// First sector of MiniFAT
    first_minifat_sector: u32,
    /// Number of MiniFAT sectors
    num_minifat_sectors: u32,
    /// First sector of DIFAT
    first_difat_sector: u32,
    /// Number of DIFAT sectors
    num_difat_sectors: u32,
    /// FAT sector IDs (first 109 in header)
    fat_sectors: Vec<u32>,
}

impl HeaderBuilder {
    /// Create a new header builder
    ///
    /// # Arguments
    ///
    /// * `sector_size` - Sector size (512 or 4096 bytes)
    pub fn new(sector_size: usize) -> Self {
        Self {
            sector_size,
            first_dir_sector: 0,
            num_dir_sectors: 0,
            first_minifat_sector: ENDOFCHAIN,
            num_minifat_sectors: 0,
            first_difat_sector: ENDOFCHAIN,
            num_difat_sectors: 0,
            fat_sectors: Vec::new(),
        }
    }

    /// Set the first directory sector
    pub fn set_first_dir_sector(&mut self, sector: u32) {
        self.first_dir_sector = sector;
    }

    /// Set MiniFAT information
    pub fn set_minifat(&mut self, first_sector: u32, num_sectors: u32) {
        self.first_minifat_sector = first_sector;
        self.num_minifat_sectors = num_sectors;
    }

    /// Set number of directory sectors (csectDir)
    /// For 512-byte sectors, this must be set to 0.
    pub fn set_num_dir_sectors(&mut self, num: u32) {
        self.num_dir_sectors = if self.sector_size == 512 { 0 } else { num };
    }

    /// Set DIFAT information
    pub fn set_difat(&mut self, first_sector: u32, num_sectors: u32) {
        self.first_difat_sector = first_sector;
        self.num_difat_sectors = num_sectors;
    }

    /// Add FAT sectors to the header
    ///
    /// The first 109 FAT sector IDs are stored in the header.
    pub fn add_fat_sectors(&mut self, sectors: &[u32]) {
        self.fat_sectors.extend_from_slice(sectors);
    }

    /// Generate the OLE2 header block
    pub fn generate(&self) -> Vec<u8> {
        // The on-disk header data is 512 bytes, but for DLL version 4 (4096-byte sectors)
        // the first big block spans 4096 bytes. We return a buffer of sector_size bytes
        // with the 512-byte header populated and the rest zero-filled.
        let mut header = vec![0u8; self.sector_size];

        // Magic bytes (8 bytes)
        header[0..8].copy_from_slice(MAGIC);

        // CLSID (16 bytes, all zeros)
        // header[8..24] already zeros

        // Minor version (2 bytes) - 0x003E
        header[24..26].copy_from_slice(&0x003Eu16.to_le_bytes());

        // DLL version (2 bytes) - 3 for 512-byte sectors, 4 for 4096-byte sectors
        let dll_version = if self.sector_size == 512 { 3u16 } else { 4u16 };
        header[26..28].copy_from_slice(&dll_version.to_le_bytes());

        // Byte order (2 bytes) - 0xFFFE for little-endian
        header[28..30].copy_from_slice(&0xFFFEu16.to_le_bytes());

        // Sector shift (2 bytes) - 9 for 512, 12 for 4096
        let sector_shift = if self.sector_size == 512 { 9u16 } else { 12u16 };
        header[30..32].copy_from_slice(&sector_shift.to_le_bytes());

        // Mini sector shift (2 bytes) - always 6 (64 bytes)
        header[32..34].copy_from_slice(&6u16.to_le_bytes());

        // Reserved (6 bytes)
        // header[34..40] already zeros

        // Number of directory sectors (csectDir). For 512-byte sectors, must be 0.
        header[40..44].copy_from_slice(&self.num_dir_sectors.to_le_bytes());

        // Number of FAT sectors (4 bytes)
        let num_fat_sectors = self.fat_sectors.len() as u32;
        header[44..48].copy_from_slice(&num_fat_sectors.to_le_bytes());

        // First directory sector (4 bytes)
        header[48..52].copy_from_slice(&self.first_dir_sector.to_le_bytes());

        // Transaction signature (4 bytes) - 0
        // header[52..56] already zeros

        // Mini stream cutoff size (4 bytes) - 4096
        header[56..60].copy_from_slice(&4096u32.to_le_bytes());

        // First MiniFAT sector (4 bytes)
        header[60..64].copy_from_slice(&self.first_minifat_sector.to_le_bytes());

        // Number of MiniFAT sectors (4 bytes)
        header[64..68].copy_from_slice(&self.num_minifat_sectors.to_le_bytes());

        // First DIFAT sector (4 bytes)
        header[68..72].copy_from_slice(&self.first_difat_sector.to_le_bytes());

        // Number of DIFAT sectors (4 bytes)
        header[72..76].copy_from_slice(&self.num_difat_sectors.to_le_bytes());

        // First 109 FAT sector IDs (436 bytes = 109 * 4)
        for (i, &sector_id) in self.fat_sectors.iter().take(109).enumerate() {
            let offset = 76 + i * 4;
            header[offset..offset + 4].copy_from_slice(&sector_id.to_le_bytes());
        }

        // Fill remaining FAT sector slots with FREESECT
        for i in self.fat_sectors.len()..109 {
            let offset = 76 + i * 4;
            header[offset..offset + 4].copy_from_slice(&FREESECT.to_le_bytes());
        }

        header
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_header_generation() {
        let mut builder = HeaderBuilder::new(512);
        builder.set_first_dir_sector(10);
        builder.add_fat_sectors(&[1, 2, 3]);

        let header = builder.generate();

        assert_eq!(header.len(), 512);
        assert_eq!(&header[0..8], MAGIC);
        assert_eq!(&header[28..30], &0xFFFEu16.to_le_bytes()); // Little-endian marker
    }

    #[test]
    fn test_sector_size_512() {
        let builder = HeaderBuilder::new(512);
        let header = builder.generate();

        // DLL version should be 3
        assert_eq!(&header[26..28], &3u16.to_le_bytes());
        // Sector shift should be 9 (2^9 = 512)
        assert_eq!(&header[30..32], &9u16.to_le_bytes());
    }

    #[test]
    fn test_sector_size_4096() {
        let builder = HeaderBuilder::new(4096);
        let header = builder.generate();

        assert_eq!(header.len(), 4096);
        // DLL version should be 4
        assert_eq!(&header[26..28], &4u16.to_le_bytes());
        // Sector shift should be 12 (2^12 = 4096)
        assert_eq!(&header[30..32], &12u16.to_le_bytes());
    }
}
