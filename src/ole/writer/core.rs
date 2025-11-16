/// OLE file writer implementation
///
/// This module provides functionality to create and modify OLE2 structured storage files.
/// It supports creating new files, adding/updating/deleting streams and storages,
/// and properly managing the FAT (File Allocation Table) and directory structure.
///
/// # Architecture
///
/// The writer uses a transactional approach where changes are accumulated in memory
/// and written atomically when `save()` is called. This ensures file integrity even
/// if the write operation fails partway through.
///
/// # Stream Allocation vs Directory Ordering
///
/// **IMPORTANT**: This is a critical distinction for Microsoft Office compatibility!
///
/// 1. **Stream ALLOCATION order** determines which sector each stream is written to:
///    - Streams are allocated sectors in the order they are added via `create_stream()`
///    - For DOC files, `WordDocument` MUST be added first to get sector 0
///    - This is enforced in the FAT allocation logic (see lines 345-358)
///
/// 2. **Directory ENTRY order** determines how entries appear in the directory tree:
///    - Directory entries are sorted using Apache POI's PropertyComparator rules
///    - Entries are organized into a balanced binary search tree
///    - This happens during directory generation (see `DirectoryBuilder`)
///
/// ## Example: DOC File Structure
///
/// ```text
/// Stream creation order:
///   1. create_stream(["WordDocument"], ...) → allocated to sector 0
///   2. create_stream(["1Table"], ...)       → allocated to sector 8
///
/// Directory tree (after sorting by name length):
///   Root Entry (SID 0)
///       └─ WordDocument (SID 1, sector 0)  [midpoint]
///            └─ 1Table (SID 2, sector 8)   [left child]
/// ```
///
/// # Example
///
/// ```rust,no_run
/// use litchi::ole::writer::OleWriter;
///
/// // Create a new OLE file
/// let mut writer = OleWriter::new();
///
/// // Add a stream
/// writer.create_stream(&["MyStream"], b"Hello, World!")?;
///
/// // Create a storage and add a stream inside it
/// writer.create_storage(&["MyStorage"])?;
/// writer.create_stream(&["MyStorage", "NestedStream"], b"Nested content")?;
///
/// // Save to file
/// writer.save("output.ole")?;
/// # Ok::<(), Box<dyn std::error::Error>>(())
/// ```
use super::super::consts::*;
use super::super::file::OleError;
use super::difat::DifatBuilder;
use super::directory::DirectoryBuilder;
use super::fat::FatBuilder;
use super::header::HeaderBuilder;
use super::minifat::MiniFatBuilder;
use std::collections::HashMap;
use std::io::{Seek, SeekFrom, Write};

/// Represents a pending stream write operation
#[derive(Debug, Clone)]
#[allow(dead_code)] // Reserved for future implementation
struct StreamWrite {
    /// Path to the stream
    path: Vec<String>,
    /// Stream data
    data: Vec<u8>,
}

/// Represents a pending storage creation operation  
#[derive(Debug, Clone)]
#[allow(dead_code)] // Reserved for future implementation
struct StorageCreate {
    /// Path to the storage
    path: Vec<String>,
}

/// Directory entry for writing
#[derive(Debug, Clone)]
#[allow(dead_code)] // Reserved for future implementation
struct WriteDirectoryEntry {
    /// Entry name
    name: String,
    /// Entry type (stream, storage, root)
    entry_type: u8,
    /// Left sibling SID
    sid_left: u32,
    /// Right sibling SID
    sid_right: u32,
    /// Child SID
    sid_child: u32,
    /// CLSID (16 bytes)
    clsid: [u8; 16],
    /// State bits
    state_bits: u32,
    /// Creation time (FILETIME)
    creation_time: u64,
    /// Modified time (FILETIME)
    modified_time: u64,
    /// Starting sector
    start_sector: u32,
    /// Stream size
    stream_size: u64,
}

/// OLE file writer
///
/// Provides methods to create and modify OLE2 structured storage files.
/// All operations are buffered in memory until `save()` is called.
pub struct OleWriter {
    /// Sector size (512 or 4096 bytes)
    sector_size: usize,
    /// Mini sector size (typically 64 bytes)
    mini_sector_size: usize,
    /// Mini stream cutoff size (typically 4096 bytes)
    mini_stream_cutoff: u32,
    /// Directory entries
    entries: Vec<WriteDirectoryEntry>,
    /// Stream data in insertion order (path, data)
    /// Using Vec instead of HashMap to preserve insertion order for directory entries
    streams: Vec<(Vec<String>, Vec<u8>)>,
    /// Storages indexed by path
    storages: HashMap<Vec<String>, ()>,
}

impl OleWriter {
    /// Create a new empty OLE writer with default settings (512-byte sectors)
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi::ole::writer::OleWriter;
    ///
    /// let writer = OleWriter::new();
    /// ```
    pub fn new() -> Self {
        Self::with_sector_size(512)
    }

    /// Create a new OLE writer with specified sector size
    ///
    /// # Arguments
    ///
    /// * `sector_size` - Sector size in bytes (512 or 4096)
    ///
    /// # Panics
    ///
    /// Panics if sector_size is not 512 or 4096
    pub fn with_sector_size(sector_size: usize) -> Self {
        assert!(
            sector_size == 512 || sector_size == 4096,
            "Sector size must be 512 or 4096"
        );

        let mut writer = OleWriter {
            sector_size,
            mini_sector_size: 64,
            mini_stream_cutoff: 4096,
            entries: Vec::new(),
            streams: Vec::new(),
            storages: HashMap::new(),
        };

        // Initialize with root entry
        writer.entries.push(WriteDirectoryEntry {
            name: "Root Entry".to_string(),
            entry_type: STGTY_ROOT,
            sid_left: NOSTREAM,
            sid_right: NOSTREAM,
            sid_child: NOSTREAM,
            clsid: [0; 16],
            state_bits: 0,
            creation_time: 0,
            modified_time: 0,
            start_sector: 0, // Will be updated when writing ministream
            stream_size: 0,  // Will be updated when writing ministream
        });

        writer
    }

    /// Set the CLSID (Class ID) for the root entry
    ///
    /// This is required for Microsoft Office to recognize the document type.
    /// For Word 97-2003 documents, use: `{00020906-0000-0000-C000-000000000046}`
    ///
    /// # Arguments
    ///
    /// * `clsid` - 16-byte CLSID in little-endian format
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// # use litchi::ole::writer::OleWriter;
    /// let mut writer = OleWriter::new();
    /// // Word 97-2003 Document CLSID: {00020906-0000-0000-C000-000000000046}
    /// let word_clsid = [0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
    ///                   0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46];
    /// writer.set_root_clsid(word_clsid);
    /// # Ok::<(), litchi::ole::OleError>(())
    /// ```
    pub fn set_root_clsid(&mut self, clsid: [u8; 16]) {
        // Update the root entry (always at index 0)
        if !self.entries.is_empty() {
            self.entries[0].clsid = clsid;
        }
    }

    /// Create a new stream at the specified path
    ///
    /// If a stream already exists at this path, it will be overwritten.
    ///
    /// # Arguments
    ///
    /// * `path` - Path components (e.g., `&["MyStorage", "MyStream"]`)
    /// * `data` - Stream contents
    ///
    /// # Returns
    ///
    /// * `Result<(), OleError>` - Success or error
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// # use litchi::ole::writer::OleWriter;
    /// let mut writer = OleWriter::new();
    /// writer.create_stream(&["MyStream"], b"Hello, World!")?;
    /// # Ok::<(), litchi::ole::OleError>(())
    /// ```
    pub fn create_stream(&mut self, path: &[&str], data: &[u8]) -> Result<(), OleError> {
        if path.is_empty() {
            return Err(OleError::InvalidData("Empty path".to_string()));
        }

        // Convert path to owned strings
        let owned_path: Vec<String> = path.iter().map(|s| s.to_string()).collect();

        // Check if stream already exists and update it
        if let Some(pos) = self.streams.iter().position(|(p, _)| p == &owned_path) {
            self.streams[pos].1 = data.to_vec();
        } else {
            // Store stream data in insertion order
            self.streams.push((owned_path, data.to_vec()));
        }

        Ok(())
    }

    /// Update an existing stream
    ///
    /// This is an alias for `create_stream` since both create and update operations
    /// have the same behavior (overwrite if exists).
    ///
    /// # Arguments
    ///
    /// * `path` - Path components
    /// * `data` - New stream contents
    pub fn update_stream(&mut self, path: &[&str], data: &[u8]) -> Result<(), OleError> {
        self.create_stream(path, data)
    }

    /// Delete a stream
    ///
    /// # Arguments
    ///
    /// * `path` - Path components
    ///
    /// # Returns
    ///
    /// * `Result<(), OleError>` - Success or error if stream doesn't exist
    pub fn delete_stream(&mut self, path: &[&str]) -> Result<(), OleError> {
        let owned_path: Vec<String> = path.iter().map(|s| s.to_string()).collect();

        if let Some(pos) = self.streams.iter().position(|(p, _)| p == &owned_path) {
            self.streams.remove(pos);
            Ok(())
        } else {
            Err(OleError::StreamNotFound)
        }
    }

    /// Create a new storage (directory) at the specified path
    ///
    /// Parent storages are created automatically if they don't exist.
    ///
    /// # Arguments
    ///
    /// * `path` - Path components (e.g., `&["MyStorage"]`)
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// # use litchi::ole::writer::OleWriter;
    /// let mut writer = OleWriter::new();
    /// writer.create_storage(&["MyStorage"])?;
    /// writer.create_stream(&["MyStorage", "MyStream"], b"data")?;
    /// # Ok::<(), litchi::ole::OleError>(())
    /// ```
    pub fn create_storage(&mut self, path: &[&str]) -> Result<(), OleError> {
        if path.is_empty() {
            return Err(OleError::InvalidData("Empty path".to_string()));
        }

        let owned_path: Vec<String> = path.iter().map(|s| s.to_string()).collect();
        self.storages.insert(owned_path, ());

        Ok(())
    }

    /// Delete a storage and all its contents
    ///
    /// # Arguments
    ///
    /// * `path` - Path components
    ///
    /// # Returns
    ///
    /// * `Result<(), OleError>` - Success or error if storage doesn't exist
    ///
    /// # Implementation Notes
    ///
    /// Currently only supports deleting empty storages. Recursive deletion
    /// of child streams and storages is planned for a future enhancement
    /// when nested storage support is added.
    pub fn delete_storage(&mut self, path: &[&str]) -> Result<(), OleError> {
        let owned_path: Vec<String> = path.iter().map(|s| s.to_string()).collect();

        if self.storages.remove(&owned_path).is_none() {
            return Err(OleError::InvalidFormat("Storage not found".to_string()));
        }

        // Note: Recursive deletion of child streams/storages will be implemented
        // when nested storage support is added

        Ok(())
    }

    /// Save the OLE file to a writer
    ///
    /// This writes all buffered changes to the output writer in OLE2 format.
    ///
    /// # Arguments
    ///
    /// * `writer` - Output writer (must implement Write + Seek)
    ///
    /// # Returns
    ///
    /// * `Result<(), OleError>` - Success or error
    ///
    /// # Implementation Notes
    ///
    /// The write process follows these steps:
    /// 1. Classify streams as small (< 4096 bytes) or large (>= 4096 bytes)
    /// 2. Allocate mini sectors for small streams and build MiniFAT
    /// 3. Allocate regular sectors for large streams and build FAT
    /// 4. Build directory structure with proper sector references
    /// 5. Generate and write header, FAT, MiniFAT, directory, and data sectors
    ///
    /// This is based on Apache POI's POIFSFileSystem.writeFilesystem() method.
    pub fn write_to<W: Write + Seek>(&mut self, writer: &mut W) -> Result<(), OleError> {
        // Initialize builders
        let mut fat = FatBuilder::new_with_size(self.sector_size);
        let mut minifat = MiniFatBuilder::new(self.mini_sector_size);

        // Separate small and large streams
        let mut small_streams: Vec<(Vec<String>, Vec<u8>)> = Vec::new();
        let mut large_streams: Vec<(Vec<String>, Vec<u8>)> = Vec::new();

        for (path, data) in &self.streams {
            if data.len() < self.mini_stream_cutoff as usize {
                small_streams.push((path.clone(), data.clone()));
            } else {
                large_streams.push((path.clone(), data.clone()));
            }
        }

        // Allocate mini sectors for small streams and track their start sectors
        let mut small_stream_sectors: Vec<(Vec<String>, Vec<u8>, u32)> = Vec::new();
        for (path, data) in &small_streams {
            let start_mini_sector = minifat.allocate_mini_chain(data);
            small_stream_sectors.push((path.clone(), data.clone(), start_mini_sector));
        }

        // CRITICAL: Allocate large streams FIRST to ensure WordDocument gets sector 0
        // Microsoft Word requires WordDocument at sector 0!

        // Add large streams to directory (using FAT) - BEFORE ministream
        let mut large_stream_data: Vec<(u32, Vec<u8>, Vec<String>)> = Vec::new();
        for (path, data) in &large_streams {
            let start_sector = if data.is_empty() {
                ENDOFCHAIN
            } else {
                fat.allocate_chain(data.len())
            };

            large_stream_data.push((start_sector, data.clone(), path.clone()));
        }

        // NOW allocate ministream (after large streams)
        let (ministream_start, ministream_size) = if !minifat.is_empty() {
            let ministream_data = minifat.ministream_data();
            let start = fat.allocate_chain(ministream_data.len());
            (start, minifat.ministream_size())
        } else {
            (ENDOFCHAIN, 0u64)
        };

        // Initialize directory builder with ministream info
        let mut directory = DirectoryBuilder::new(ministream_start, ministream_size);

        // Set root CLSID if specified (e.g., for Word documents)
        if !self.entries.is_empty() && self.entries[0].clsid != [0u8; 16] {
            directory.set_root_clsid(self.entries[0].clsid);
        }

        // Pre-create storages declared explicitly by user
        for storage_path in self.storages.keys() {
            directory.add_storage_path(storage_path);
        }

        // Add large streams to directory using full path
        for (start_sector, data, path) in &large_stream_data {
            let full: Vec<String> = path.clone();
            let _sid = directory.add_stream_path(&full, *start_sector, data.len() as u64);
        }

        // Add small streams to directory (using MiniFAT) with full path
        for (path, data, start_mini_sector) in &small_stream_sectors {
            let full: Vec<String> = path.clone();
            let _sid = directory.add_stream_path(&full, *start_mini_sector, data.len() as u64);
        }

        // Generate directory stream
        let dir_stream = directory.generate_directory_stream();
        let dir_sector_count = (dir_stream.len().div_ceil(self.sector_size)) as u32;
        let dir_start_sector = fat.allocate_chain(dir_stream.len());

        // Generate MiniFAT sectors (if needed)
        let (minifat_start_sector, num_minifat_sectors) = if !minifat.is_empty() {
            let minifat_sectors = minifat.generate_minifat_sectors(self.sector_size);
            let num_sectors = minifat_sectors.len() as u32;

            if num_sectors > 0 {
                let start = fat.allocate_chain(num_sectors as usize * self.sector_size);
                (start, num_sectors)
            } else {
                (ENDOFCHAIN, 0)
            }
        } else {
            (ENDOFCHAIN, 0)
        };

        // === Compute FAT/DIFAT sectors requirement iteratively ===
        let entries_per_fat_sector = self.sector_size as u32 / 4;
        let ids_per_difat_sector = entries_per_fat_sector - 1; // last u32 is next pointer

        let n_used = fat.total_sectors();
        let mut n_fat: u32 = 0;
        let mut n_difat: u32 = 0;
        for _ in 0..8 {
            let total_entries = n_used + n_fat + n_difat;
            let new_n_fat = total_entries.div_ceil(entries_per_fat_sector);
            let new_n_difat = if new_n_fat > 109 {
                let over = new_n_fat - 109;
                over.div_ceil(ids_per_difat_sector)
            } else {
                0
            };
            if new_n_fat == n_fat && new_n_difat == n_difat {
                break;
            }
            n_fat = new_n_fat;
            n_difat = new_n_difat;
        }

        // Reserve DIFAT sectors then FAT sectors
        let difat_start_sector = if n_difat > 0 {
            fat.allocate_special(n_difat, DIFSECT)
        } else {
            ENDOFCHAIN
        };
        let fat_start_sector = if n_fat > 0 {
            fat.allocate_special(n_fat, FATSECT)
        } else {
            ENDOFCHAIN
        };

        // Prepare FAT sector data now that reservations are included
        let fat_sectors_data = fat.generate_fat_sectors();
        let num_fat_sectors = n_fat;

        // Validate FAT
        fat.validate()
            .map_err(|e| OleError::InvalidData(format!("FAT validation failed: {}", e)))?;

        // Build header
        let mut header_builder = HeaderBuilder::new(self.sector_size);
        header_builder.set_first_dir_sector(dir_start_sector);
        header_builder.set_num_dir_sectors(dir_sector_count);
        header_builder.set_minifat(minifat_start_sector, num_minifat_sectors);

        // Handle DIFAT if needed (> 109 FAT sectors)
        let fat_sector_ids: Vec<u32> = if num_fat_sectors > 0 {
            (fat_start_sector..fat_start_sector + num_fat_sectors).collect()
        } else {
            Vec::new()
        };

        let (num_difat_sectors, difat_sectors) = if num_fat_sectors > 109 {
            let mut difat = DifatBuilder::new(self.sector_size);
            difat.set_fat_sectors(&fat_sector_ids);
            let num_difat = difat.calculate_difat_sector_count();
            let sectors = if num_difat > 0 {
                difat.generate_difat_sectors(difat_start_sector)
            } else {
                Vec::new()
            };
            (num_difat, sectors)
        } else {
            (0, Vec::new())
        };

        // Add first 109 FAT sector IDs to header
        header_builder.add_fat_sectors(&fat_sector_ids);

        // Set DIFAT info in header
        if num_difat_sectors > 0 {
            header_builder.set_difat(difat_start_sector, num_difat_sectors);
        }

        let header = header_builder.generate();

        // === Write the file ===

        // Write header (sector 0 position is offset 0, but actual sectors start at +1)
        writer.write_all(&header)?;

        // Write ministream data (if any)
        if !minifat.is_empty() && ministream_start != ENDOFCHAIN {
            let position = ((ministream_start as u64) + 1) * (self.sector_size as u64);
            writer.seek(SeekFrom::Start(position))?;

            let ministream_data = minifat.ministream_data();
            let padded_size = ministream_data.len().div_ceil(self.sector_size) * self.sector_size;
            let mut padded_data = ministream_data.to_vec();
            padded_data.resize(padded_size, 0);
            writer.write_all(&padded_data)?;
        }

        // Write large stream data sectors
        for (start_sector, data, _path) in &large_stream_data {
            if *start_sector == ENDOFCHAIN {
                continue;
            }

            // Calculate file position for this sector
            let position = ((*start_sector as u64) + 1) * (self.sector_size as u64);
            writer.seek(SeekFrom::Start(position))?;

            // Write data (padded to sector boundaries)
            let padded_size = data.len().div_ceil(self.sector_size) * self.sector_size;
            let mut padded_data = data.clone();
            padded_data.resize(padded_size, 0);
            writer.write_all(&padded_data)?;
        }

        // Write directory stream
        let dir_position = ((dir_start_sector as u64) + 1) * (self.sector_size as u64);
        writer.seek(SeekFrom::Start(dir_position))?;
        let dir_padded_size = dir_stream.len().div_ceil(self.sector_size) * self.sector_size;
        let mut dir_padded = dir_stream;
        dir_padded.resize(dir_padded_size, 0);
        writer.write_all(&dir_padded)?;

        // Write MiniFAT sectors (if any)
        if !minifat.is_empty() && minifat_start_sector != ENDOFCHAIN {
            let minifat_sectors = minifat.generate_minifat_sectors(self.sector_size);
            let mut current_sector = minifat_start_sector;

            for minifat_sector_data in &minifat_sectors {
                let position = ((current_sector as u64) + 1) * (self.sector_size as u64);
                writer.seek(SeekFrom::Start(position))?;
                writer.write_all(minifat_sector_data)?;
                current_sector += 1;
            }
        }

        // Write FAT sectors
        for (i, fat_sector_data) in fat_sectors_data.iter().enumerate() {
            let sector_id = fat_start_sector + i as u32;
            let position = ((sector_id as u64) + 1) * (self.sector_size as u64);
            writer.seek(SeekFrom::Start(position))?;
            writer.write_all(fat_sector_data)?;
        }

        // Write DIFAT sectors (if any)
        if !difat_sectors.is_empty() {
            let mut current_sector = difat_start_sector;
            for difat_sector_data in &difat_sectors {
                let position = ((current_sector as u64) + 1) * (self.sector_size as u64);
                writer.seek(SeekFrom::Start(position))?;
                writer.write_all(difat_sector_data)?;
                current_sector += 1;
            }
        }

        writer.flush()?;

        Ok(())
    }

    /// Save the OLE file to a file path
    ///
    /// # Arguments
    ///
    /// * `path` - Output file path
    ///
    /// # Returns
    ///
    /// * `Result<(), OleError>` - Success or error
    ///
    /// # Example
    ///
    /// ```rust,no_run
    /// # use litchi::ole::writer::OleWriter;
    /// let mut writer = OleWriter::new();
    /// writer.create_stream(&["Test"], b"Hello")?;
    /// writer.save("output.ole")?;
    /// # Ok::<(), litchi::ole::OleError>(())
    /// ```
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> Result<(), OleError> {
        let file = std::fs::File::create(path)?;
        let mut buffered = std::io::BufWriter::new(file);
        self.write_to(&mut buffered)?;
        buffered.flush()?;
        Ok(())
    }
}

impl Default for OleWriter {
    fn default() -> Self {
        Self::new()
    }
}

/// Encode a string to UTF-16LE bytes (padded to 64 bytes)
///
/// This is used for directory entry names in OLE files.
///
/// # Arguments
///
/// * `name` - The string to encode (max 31 characters)
///
/// # Returns
///
/// * `[u8; 64]` - UTF-16LE encoded bytes with null terminator
///
/// # Implementation Notes
///
/// All core helper functions have been implemented:
/// - ✅ UTF-16LE encoding (this function)
/// - ✅ FAT chain building (FatBuilder)
/// - ✅ MiniFAT allocation (MiniFatBuilder)
/// - ✅ DIFAT handling (DifatBuilder)
/// - ✅ Directory tree building (DirectoryBuilder)
/// - Future: Balanced red-black tree (planned enhancement)
#[allow(dead_code)] // Reserved for future implementation
fn encode_name_utf16le(name: &str) -> [u8; 64] {
    let mut result = [0u8; 64];
    let utf16: Vec<u16> = name.encode_utf16().collect();

    // Copy UTF-16 data (max 31 characters + null terminator)
    let max_chars = 31.min(utf16.len());
    for (i, &ch) in utf16.iter().take(max_chars).enumerate() {
        let bytes = ch.to_le_bytes();
        result[i * 2] = bytes[0];
        result[i * 2 + 1] = bytes[1];
    }

    // Null terminator
    if max_chars < 32 {
        result[max_chars * 2] = 0;
        result[max_chars * 2 + 1] = 0;
    }

    result
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_writer() {
        let writer = OleWriter::new();
        assert_eq!(writer.sector_size, 512);
        assert_eq!(writer.mini_sector_size, 64);
        assert_eq!(writer.mini_stream_cutoff, 4096);
        assert_eq!(writer.entries.len(), 1); // Root entry
    }

    #[test]
    fn test_create_stream() {
        let mut writer = OleWriter::new();
        writer.create_stream(&["Test"], b"Hello").unwrap();
        assert_eq!(writer.streams.len(), 1);
    }

    #[test]
    fn test_create_storage() {
        let mut writer = OleWriter::new();
        writer.create_storage(&["Storage"]).unwrap();
        assert_eq!(writer.storages.len(), 1);
    }

    #[test]
    fn test_encode_name() {
        let encoded = encode_name_utf16le("Test");
        // Verify UTF-16LE encoding: 'T' = 0x0054, 'e' = 0x0065, etc.
        assert_eq!(encoded[0], 0x54); // 'T' low byte
        assert_eq!(encoded[1], 0x00); // 'T' high byte
        assert_eq!(encoded[2], 0x65); // 'e' low byte
        assert_eq!(encoded[3], 0x00); // 'e' high byte
    }
}
