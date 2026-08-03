/// OLE file writer implementation
///
/// This module provides functionality to create and modify OLE2 structured storage files.
/// It supports creating new files, adding/updating/deleting streams and storages,
/// and properly managing the FAT (File Allocation Table) and directory structure.
///
/// # Architecture
///
/// The writer accumulates a complete logical model in memory before serialization.
/// `write_to` and `save` write directly to their destination, so a sink failure can
/// leave partial output; callers that require filesystem replacement semantics must
/// finalize through an atomic temporary-file layer.
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
/// use litchi_cfb::writer::OleWriter;
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
use std::hash::Hash;
use std::io::{Seek, SeekFrom, Write};

const V3_MAX_FILE_BYTES: u64 = 2 * 1024 * 1024 * 1024;

#[derive(Debug, Clone, Copy)]
struct StreamPlan {
    index: usize,
    start_sector: u32,
}

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
    /// Non-zero CLSIDs assigned to individual storages.
    storage_clsids: HashMap<Vec<String>, [u8; 16]>,
}

impl OleWriter {
    /// Create a new empty OLE writer with default settings (512-byte sectors)
    ///
    /// # Example
    ///
    /// ```rust
    /// use litchi_cfb::writer::OleWriter;
    ///
    /// let writer = OleWriter::new();
    /// ```
    pub fn new() -> Self {
        Self::with_valid_sector_size(512)
    }

    /// Create a new OLE writer with specified sector size
    ///
    /// # Arguments
    ///
    /// * `sector_size` - Sector size in bytes (512 or 4096)
    ///
    /// Returns a typed error when `sector_size` is not 512 or 4096 bytes.
    pub fn with_sector_size(sector_size: usize) -> Result<Self, OleError> {
        if !matches!(sector_size, 512 | 4096) {
            return Err(OleError::InvalidData(format!(
                "sector size must be 512 or 4096 bytes, got {sector_size}"
            )));
        }
        Ok(Self::with_valid_sector_size(sector_size))
    }

    fn with_valid_sector_size(sector_size: usize) -> Self {
        let mut writer = Self {
            sector_size,
            mini_sector_size: 64,
            mini_stream_cutoff: 4096,
            entries: Vec::new(),
            streams: Vec::new(),
            storages: HashMap::new(),
            storage_clsids: HashMap::new(),
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
    /// # use litchi_cfb::writer::OleWriter;
    /// let mut writer = OleWriter::new();
    /// // Word 97-2003 Document CLSID: {00020906-0000-0000-C000-000000000046}
    /// let word_clsid = [0x06, 0x09, 0x02, 0x00, 0x00, 0x00, 0x00, 0x00,
    ///                   0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46];
    /// writer.set_root_clsid(word_clsid);
    /// # Ok::<(), litchi_cfb::OleError>(())
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
    /// # use litchi_cfb::writer::OleWriter;
    /// let mut writer = OleWriter::new();
    /// writer.create_stream(&["MyStream"], b"Hello, World!")?;
    /// # Ok::<(), litchi_cfb::OleError>(())
    /// ```
    pub fn create_stream(&mut self, path: &[&str], data: &[u8]) -> Result<(), OleError> {
        if path.is_empty() {
            return Err(OleError::InvalidData("Empty path".to_string()));
        }
        let mut owned = Vec::new();
        owned
            .try_reserve_exact(data.len())
            .map_err(|source| OleError::allocation("stream payload", source))?;
        owned.extend_from_slice(data);
        self.create_stream_owned(path, owned)
    }

    /// Create or replace a stream by taking ownership of its payload allocation.
    ///
    /// Unlike [`Self::create_stream`], this method never clones `data`. The
    /// allocation remains owned by the writer after this call and is reused by
    /// subsequent [`Self::write_to`] calls.
    pub fn create_stream_owned(&mut self, path: &[&str], data: Vec<u8>) -> Result<(), OleError> {
        if path.is_empty() {
            return Err(OleError::InvalidData("Empty path".to_string()));
        }

        if let Some(position) = self.stream_position(path) {
            self.streams[position].1 = data;
            return Ok(());
        }

        let owned_path = own_path(path, "stream path", "stream path component")?;
        self.streams
            .try_reserve(1)
            .map_err(|source| OleError::allocation("stream table", source))?;
        self.streams.push((owned_path, data));
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

    fn stream_position(&self, path: &[&str]) -> Option<usize> {
        self.streams.iter().position(|(candidate, _)| {
            candidate.len() == path.len()
                && candidate
                    .iter()
                    .zip(path)
                    .all(|(owned, borrowed)| owned == borrowed)
        })
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
        let owned_path = own_path(path, "stream path", "stream path component")?;

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
    /// # use litchi_cfb::writer::OleWriter;
    /// let mut writer = OleWriter::new();
    /// writer.create_storage(&["MyStorage"])?;
    /// writer.create_stream(&["MyStorage", "MyStream"], b"data")?;
    /// # Ok::<(), litchi_cfb::OleError>(())
    /// ```
    pub fn create_storage(&mut self, path: &[&str]) -> Result<(), OleError> {
        if path.is_empty() {
            return Err(OleError::InvalidData("Empty path".to_string()));
        }

        let owned_path = own_path(path, "storage path", "storage path component")?;
        reserve_hash_map_entry(&mut self.storages, &owned_path, 1, "storage table")?;
        self.storages.insert(owned_path, ());

        Ok(())
    }

    /// Assign a CLSID to an explicitly created storage.
    ///
    /// `path` must identify a storage previously registered with
    /// [`Self::create_storage`]. CLSIDs are encoded in CFB byte order.
    pub fn set_storage_clsid(&mut self, path: &[&str], clsid: [u8; 16]) -> Result<(), OleError> {
        let owned_path = own_path(path, "storage path", "storage path component")?;
        if !self.storages.contains_key(&owned_path) {
            return Err(OleError::InvalidData(format!(
                "CFB storage path {owned_path:?} does not exist"
            )));
        }
        if clsid == [0; 16] {
            self.storage_clsids.remove(&owned_path);
        } else {
            reserve_hash_map_entry(
                &mut self.storage_clsids,
                &owned_path,
                1,
                "storage CLSID table",
            )?;
            self.storage_clsids.insert(owned_path, clsid);
        }
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
    /// The root entry is represented by an empty path and cannot be deleted.
    /// A successful deletion removes the target storage and every stream,
    /// storage, and storage CLSID whose full path is beneath it.
    pub fn delete_storage(&mut self, path: &[&str]) -> Result<(), OleError> {
        if path.is_empty() {
            return Err(OleError::InvalidData("Empty path".to_string()));
        }

        let owned_path = own_path(path, "storage path", "storage path component")?;

        // Validate the complete operation before mutating any table. The
        // retain calls below are infallible, so a failed lookup leaves the
        // writer unchanged.
        if !self.storages.contains_key(&owned_path) {
            return Err(OleError::InvalidFormat("Storage not found".to_string()));
        }

        self.streams
            .retain(|(candidate, _)| !candidate.starts_with(owned_path.as_slice()));
        self.storages
            .retain(|candidate, _| !candidate.starts_with(owned_path.as_slice()));
        self.storage_clsids
            .retain(|candidate, _| !candidate.starts_with(owned_path.as_slice()));

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
        for (_, data) in &self.streams {
            validate_stream_size(self.sector_size, data.len(), "user stream")?;
        }

        // Initialize builders
        let mut fat = FatBuilder::new_with_size(self.sector_size)?;
        let mut minifat = MiniFatBuilder::new(self.mini_sector_size);

        // Classify by insertion-order index. Plans retain only compact metadata;
        // stream paths and payload allocations stay in `self.streams`.
        let mut small_streams = Vec::new();
        let mut large_streams = Vec::new();
        let small_count = self
            .streams
            .iter()
            .filter(|(_, data)| data.len() < self.mini_stream_cutoff as usize)
            .count();
        let large_count = self.streams.len().checked_sub(small_count).ok_or_else(|| {
            OleError::InvalidData("CFB stream classification underflow".to_string())
        })?;
        small_streams
            .try_reserve_exact(small_count)
            .map_err(|source| OleError::allocation("small-stream plan", source))?;
        large_streams
            .try_reserve_exact(large_count)
            .map_err(|source| OleError::allocation("large-stream plan", source))?;

        for (index, (_, data)) in self.streams.iter().enumerate() {
            if data.len() < self.mini_stream_cutoff as usize {
                small_streams.push(StreamPlan {
                    index,
                    start_sector: ENDOFCHAIN,
                });
            } else {
                large_streams.push(StreamPlan {
                    index,
                    start_sector: ENDOFCHAIN,
                });
            }
        }

        // Allocate mini sectors for small streams and track their start sectors
        for plan in &mut small_streams {
            let data = &self.streams[plan.index].1;
            plan.start_sector = minifat.allocate_mini_chain(data)?;
        }

        // CRITICAL: Allocate large streams FIRST to ensure WordDocument gets sector 0
        // Microsoft Word requires WordDocument at sector 0!

        // Add large streams to directory (using FAT) - BEFORE ministream
        for plan in &mut large_streams {
            let data = &self.streams[plan.index].1;
            plan.start_sector = if data.is_empty() {
                ENDOFCHAIN
            } else {
                fat.allocate_chain(data.len())?
            };
        }

        // NOW allocate ministream (after large streams)
        let (ministream_start, ministream_size) = if !minifat.is_empty() {
            let ministream_data = minifat.ministream_data();
            validate_stream_size(self.sector_size, ministream_data.len(), "mini stream")?;
            let start = fat.allocate_chain(ministream_data.len())?;
            (start, minifat.ministream_size()?)
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
            directory.add_storage_path(storage_path)?;
        }
        for (storage_path, clsid) in &self.storage_clsids {
            directory.set_storage_clsid(storage_path, *clsid)?;
        }

        // Add large streams to directory using full path
        for plan in &large_streams {
            let (path, data) = &self.streams[plan.index];
            let size = u64::try_from(data.len()).map_err(|_| {
                OleError::InvalidData("CFB stream size does not fit u64".to_string())
            })?;
            let _sid = directory.add_stream_path(path, plan.start_sector, size)?;
        }

        // Add small streams to directory (using MiniFAT) with full path
        for plan in &small_streams {
            let (path, data) = &self.streams[plan.index];
            let size = u64::try_from(data.len()).map_err(|_| {
                OleError::InvalidData("CFB stream size does not fit u64".to_string())
            })?;
            let _sid = directory.add_stream_path(path, plan.start_sector, size)?;
        }

        // Generate directory stream
        let dir_stream = directory.generate_directory_stream()?;
        validate_stream_size(self.sector_size, dir_stream.len(), "directory stream")?;
        let dir_sector_count =
            sector_count(dir_stream.len(), self.sector_size, "directory stream")?;
        let dir_start_sector = fat.allocate_chain(dir_stream.len())?;

        // Generate MiniFAT sectors (if needed)
        let minifat_sectors = if minifat.is_empty() {
            Vec::new()
        } else {
            minifat.generate_minifat_sectors(self.sector_size)?
        };
        let num_minifat_sectors = u32::try_from(minifat_sectors.len())
            .map_err(|_| OleError::InvalidData("too many MiniFAT sectors".to_string()))?;
        let minifat_bytes = minifat_sectors
            .len()
            .checked_mul(self.sector_size)
            .ok_or_else(|| OleError::InvalidData("MiniFAT size overflows usize".to_string()))?;
        let minifat_start_sector = if minifat_bytes == 0 {
            ENDOFCHAIN
        } else {
            fat.allocate_chain(minifat_bytes)?
        };

        // === Compute FAT/DIFAT sectors requirement iteratively ===
        let n_used = fat.total_sectors();
        let (n_fat, n_difat) = allocation_table_sector_counts(n_used, self.sector_size)?;

        // Reserve DIFAT sectors then FAT sectors
        let difat_start_sector = if n_difat > 0 {
            fat.allocate_special(n_difat, DIFSECT)?
        } else {
            ENDOFCHAIN
        };
        let fat_start_sector = if n_fat > 0 {
            fat.allocate_special(n_fat, FATSECT)?
        } else {
            ENDOFCHAIN
        };
        validate_output_size(self.sector_size, fat.total_sectors())?;

        // Prepare FAT sector data now that reservations are included
        let fat_sectors_data = fat.generate_fat_sectors()?;
        let num_fat_sectors = n_fat;
        let generated_fat_count = u32::try_from(fat_sectors_data.len())
            .map_err(|_| OleError::InvalidData("too many serialized FAT sectors".to_string()))?;
        if generated_fat_count != num_fat_sectors {
            return Err(OleError::InvalidData(format!(
                "CFB FAT planning mismatch: planned {num_fat_sectors}, generated {generated_fat_count}"
            )));
        }

        // Validate FAT
        fat.validate()?;

        // Build header
        let mut header_builder = HeaderBuilder::new(self.sector_size)?;
        header_builder.set_first_dir_sector(dir_start_sector);
        header_builder.set_num_dir_sectors(dir_sector_count);
        header_builder.set_minifat(minifat_start_sector, num_minifat_sectors);

        // Handle DIFAT if needed (> 109 FAT sectors)
        let fat_sector_ids = sector_ids(fat_start_sector, num_fat_sectors, "FAT sector IDs")?;

        let (num_difat_sectors, difat_sectors) = if num_fat_sectors > 109 {
            let mut difat = DifatBuilder::new(self.sector_size)?;
            difat.set_fat_sectors(&fat_sector_ids)?;
            let num_difat = difat.calculate_difat_sector_count()?;
            if num_difat != n_difat {
                return Err(OleError::InvalidData(format!(
                    "CFB DIFAT planning mismatch: planned {n_difat}, generated {num_difat}"
                )));
            }
            let sectors = if num_difat > 0 {
                difat.generate_difat_sectors(difat_start_sector)?
            } else {
                Vec::new()
            };
            (num_difat, sectors)
        } else {
            (0, Vec::new())
        };

        // Add first 109 FAT sector IDs to header
        header_builder.set_fat_sectors(&fat_sector_ids)?;

        // Set DIFAT info in header
        if num_difat_sectors > 0 {
            header_builder.set_difat(difat_start_sector, num_difat_sectors);
        }

        let header = header_builder.generate()?;

        // === Write the file ===

        // All sector offsets are absolute from the CFB header, so normalize a
        // caller-provided seekable sink before emitting any bytes.
        writer.seek(SeekFrom::Start(0))?;
        writer.write_all(&header)?;

        // Write ministream data (if any)
        if !minifat.is_empty() && ministream_start != ENDOFCHAIN {
            let position = sector_offset(ministream_start, self.sector_size)?;
            writer.seek(SeekFrom::Start(position))?;

            let ministream_data = minifat.ministream_data();
            write_sector_aligned(writer, ministream_data, self.sector_size)?;
        }

        // Write large stream data sectors
        for plan in &large_streams {
            if plan.start_sector == ENDOFCHAIN {
                continue;
            }
            let data = &self.streams[plan.index].1;

            // Calculate file position for this sector
            let position = sector_offset(plan.start_sector, self.sector_size)?;
            writer.seek(SeekFrom::Start(position))?;

            // Write the retained payload allocation, then at most one sector of
            // zero padding from a fixed stack buffer.
            write_sector_aligned(writer, data, self.sector_size)?;
        }

        // Write directory stream
        let dir_position = sector_offset(dir_start_sector, self.sector_size)?;
        writer.seek(SeekFrom::Start(dir_position))?;
        write_sector_aligned(writer, &dir_stream, self.sector_size)?;

        // Write MiniFAT sectors (if any)
        if minifat_start_sector != ENDOFCHAIN {
            for (index, minifat_sector_data) in minifat_sectors.iter().enumerate() {
                let current_sector = sector_at(minifat_start_sector, index)?;
                let position = sector_offset(current_sector, self.sector_size)?;
                writer.seek(SeekFrom::Start(position))?;
                writer.write_all(minifat_sector_data)?;
            }
        }

        // Write FAT sectors
        for (i, fat_sector_data) in fat_sectors_data.iter().enumerate() {
            let sector_id = sector_at(fat_start_sector, i)?;
            let position = sector_offset(sector_id, self.sector_size)?;
            writer.seek(SeekFrom::Start(position))?;
            writer.write_all(fat_sector_data)?;
        }

        // Write DIFAT sectors (if any)
        if !difat_sectors.is_empty() {
            for (index, difat_sector_data) in difat_sectors.iter().enumerate() {
                let current_sector = sector_at(difat_start_sector, index)?;
                let position = sector_offset(current_sector, self.sector_size)?;
                writer.seek(SeekFrom::Start(position))?;
                writer.write_all(difat_sector_data)?;
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
    /// # use litchi_cfb::writer::OleWriter;
    /// let mut writer = OleWriter::new();
    /// writer.create_stream(&["Test"], b"Hello")?;
    /// writer.save("output.ole")?;
    /// # Ok::<(), litchi_cfb::OleError>(())
    /// ```
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> Result<(), OleError> {
        let file = std::fs::File::create(path)?;
        let mut buffered = std::io::BufWriter::new(file);
        self.write_to(&mut buffered)?;
        buffered.flush()?;
        Ok(())
    }
}

fn own_path(
    path: &[&str],
    path_resource: &'static str,
    component_resource: &'static str,
) -> Result<Vec<String>, OleError> {
    let mut owned = Vec::new();
    owned
        .try_reserve_exact(path.len())
        .map_err(|source| OleError::allocation(path_resource, source))?;
    for component in path {
        let mut value = String::new();
        value
            .try_reserve_exact(component.len())
            .map_err(|source| OleError::allocation(component_resource, source))?;
        value.push_str(component);
        owned.push(value);
    }
    Ok(owned)
}

fn reserve_hash_map_entry<K, V>(
    map: &mut HashMap<K, V>,
    key: &K,
    additional: usize,
    resource: &'static str,
) -> Result<(), OleError>
where
    K: Eq + Hash,
{
    if map.contains_key(key) {
        return Ok(());
    }
    map.try_reserve(additional)
        .map_err(|source| OleError::allocation(resource, source))
}

fn validate_stream_size(
    sector_size: usize,
    size: usize,
    resource: &'static str,
) -> Result<(), OleError> {
    checked_sector_size(sector_size)?;
    if sector_size != 512 {
        return Ok(());
    }
    let size = u64::try_from(size)
        .map_err(|_| OleError::InvalidData("CFB stream size does not fit u64".to_string()))?;
    if size >= V3_MAX_FILE_BYTES {
        return Err(OleError::InvalidData(format!(
            "version 3 CFB {resource} must be smaller than 2 GiB"
        )));
    }
    Ok(())
}

fn validate_output_size(sector_size: usize, sector_count: u32) -> Result<(), OleError> {
    let sector_size_u64 = checked_sector_size(sector_size)?;
    if sector_count > MAXREGSECT {
        return Err(OleError::InvalidData(
            "CFB output exceeds MAXREGSECT".to_string(),
        ));
    }
    let bytes = u64::from(sector_count)
        .checked_add(1)
        .and_then(|count| count.checked_mul(sector_size_u64))
        .ok_or_else(|| OleError::InvalidData("CFB output size overflows u64".to_string()))?;
    if sector_size == 512 && bytes > V3_MAX_FILE_BYTES {
        return Err(OleError::InvalidData(
            "version 3 CFB output cannot exceed 2 GiB".to_string(),
        ));
    }
    Ok(())
}

fn sector_count(
    byte_len: usize,
    sector_size: usize,
    resource: &'static str,
) -> Result<u32, OleError> {
    checked_sector_size(sector_size)?;
    let count = byte_len.div_ceil(sector_size);
    let count = u32::try_from(count)
        .map_err(|_| OleError::InvalidData(format!("CFB {resource} has too many sectors")))?;
    if count > MAXREGSECT {
        return Err(OleError::InvalidData(format!(
            "CFB {resource} exceeds MAXREGSECT"
        )));
    }
    Ok(count)
}

fn allocation_table_sector_counts(used: u32, sector_size: usize) -> Result<(u32, u32), OleError> {
    checked_sector_size(sector_size)?;
    let entries_per_fat_sector = u32::try_from(sector_size / 4)
        .map_err(|_| OleError::InvalidData("CFB FAT geometry exceeds u32".to_string()))?;
    let ids_per_difat_sector = entries_per_fat_sector
        .checked_sub(1)
        .ok_or_else(|| OleError::InvalidData("CFB DIFAT sector has no ID slots".to_string()))?;
    let mut fat = 0u32;
    let mut difat = 0u32;
    for _ in 0..32 {
        let total = used
            .checked_add(fat)
            .and_then(|value| value.checked_add(difat))
            .ok_or_else(|| OleError::InvalidData("CFB sector count overflows u32".to_string()))?;
        if total > MAXREGSECT {
            return Err(OleError::InvalidData(
                "CFB sector count exceeds MAXREGSECT".to_string(),
            ));
        }
        let new_fat = total.div_ceil(entries_per_fat_sector);
        let new_difat = new_fat.saturating_sub(109).div_ceil(ids_per_difat_sector);
        if new_fat == fat && new_difat == difat {
            return Ok((fat, difat));
        }
        fat = new_fat;
        difat = new_difat;
    }
    Err(OleError::InvalidData(
        "CFB FAT/DIFAT planning did not converge".to_string(),
    ))
}

fn sector_ids(start: u32, count: u32, resource: &'static str) -> Result<Vec<u32>, OleError> {
    if count == 0 {
        return Ok(Vec::new());
    }
    let end = start
        .checked_add(count)
        .ok_or_else(|| OleError::InvalidData(format!("CFB {resource} range overflows u32")))?;
    if start >= MAXREGSECT || end > MAXREGSECT {
        return Err(OleError::InvalidData(format!(
            "CFB {resource} range exceeds MAXREGSECT"
        )));
    }
    let count = usize::try_from(count)
        .map_err(|_| OleError::InvalidData(format!("CFB {resource} count exceeds usize")))?;
    let mut ids = Vec::new();
    ids.try_reserve_exact(count)
        .map_err(|source| OleError::allocation(resource, source))?;
    ids.extend(start..end);
    Ok(ids)
}

fn sector_at(start: u32, index: usize) -> Result<u32, OleError> {
    let index = u32::try_from(index)
        .map_err(|_| OleError::InvalidData("CFB sector offset exceeds u32".to_string()))?;
    let sector = start
        .checked_add(index)
        .ok_or_else(|| OleError::InvalidData("CFB sector index overflows u32".to_string()))?;
    if sector >= MAXREGSECT {
        return Err(OleError::InvalidData(
            "CFB sector index exceeds MAXREGSECT".to_string(),
        ));
    }
    Ok(sector)
}

fn sector_offset(sector: u32, sector_size: usize) -> Result<u64, OleError> {
    let sector_size = checked_sector_size(sector_size)?;
    if sector >= MAXREGSECT {
        return Err(OleError::InvalidData(
            "CFB sector index exceeds MAXREGSECT".to_string(),
        ));
    }
    (u64::from(sector) + 1)
        .checked_mul(sector_size)
        .ok_or_else(|| OleError::InvalidData("CFB sector offset overflows u64".to_string()))
}

fn checked_sector_size(sector_size: usize) -> Result<u64, OleError> {
    if !matches!(sector_size, 512 | 4096) {
        return Err(OleError::InvalidData(format!(
            "CFB sector size must be 512 or 4096 bytes, got {sector_size}"
        )));
    }
    u64::try_from(sector_size)
        .map_err(|_| OleError::InvalidData("CFB sector size does not fit u64".to_string()))
}

fn write_sector_aligned<W: Write>(
    writer: &mut W,
    data: &[u8],
    sector_size: usize,
) -> Result<(), OleError> {
    if !matches!(sector_size, 512 | 4096) {
        return Err(OleError::InvalidData(format!(
            "CFB sector size must be 512 or 4096 bytes, got {sector_size}"
        )));
    }
    writer.write_all(data)?;
    let remainder = data.len() % sector_size;
    if remainder != 0 {
        const ZEROES: [u8; 4096] = [0; 4096];
        let padding = sector_size - remainder;
        let bytes = ZEROES.get(..padding).ok_or_else(|| {
            OleError::InvalidData("CFB sector padding exceeds 4096 bytes".to_string())
        })?;
        writer.write_all(bytes)?;
    }
    Ok(())
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
    let mut max_chars = 0;
    for (i, ch) in name.encode_utf16().take(31).enumerate() {
        let bytes = ch.to_le_bytes();
        result[i * 2] = bytes[0];
        result[i * 2 + 1] = bytes[1];
        max_chars = i + 1;
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
    use std::error::Error as _;

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
    fn owned_stream_retains_payload_allocation_through_write() {
        let mut writer = OleWriter::new();
        let data = vec![0x5a; 5_003];
        let pointer = data.as_ptr();

        writer
            .create_stream_owned(&["Owned"], data)
            .expect("move stream payload");
        assert_eq!(writer.streams[0].1.as_ptr(), pointer);

        let mut output = std::io::Cursor::new(Vec::new());
        writer.write_to(&mut output).expect("write owned stream");
        assert_eq!(writer.streams[0].1.as_ptr(), pointer);
    }

    #[test]
    fn serialization_normalizes_the_sink_position() {
        let mut writer = OleWriter::new();
        writer.create_stream(&["Test"], b"payload").unwrap();
        let mut output = std::io::Cursor::new(Vec::new());
        output.set_position(17);

        writer.write_to(&mut output).unwrap();

        assert_eq!(&output.into_inner()[..MAGIC.len()], MAGIC);
    }

    #[test]
    fn test_create_storage() {
        let mut writer = OleWriter::new();
        writer.create_storage(&["Storage"]).unwrap();
        assert_eq!(writer.storages.len(), 1);
    }

    #[test]
    fn delete_storage_removes_descendants_and_clsids_without_prefix_collisions() {
        let mut writer = OleWriter::new();
        writer.create_storage(&["Root"]).unwrap();
        writer.create_storage(&["Root", "Nested"]).unwrap();
        writer.create_storage(&["Root", "Nested", "Deep"]).unwrap();
        writer.create_storage(&["Root", "Sibling"]).unwrap();
        writer.create_storage(&["RootSibling"]).unwrap();

        let root_clsid = [0x11; 16];
        let nested_clsid = [0x22; 16];
        let deep_clsid = [0x33; 16];
        let removed_sibling_clsid = [0x44; 16];
        let sibling_clsid = [0x55; 16];
        writer.set_storage_clsid(&["Root"], root_clsid).unwrap();
        writer
            .set_storage_clsid(&["Root", "Nested"], nested_clsid)
            .unwrap();
        writer
            .set_storage_clsid(&["Root", "Nested", "Deep"], deep_clsid)
            .unwrap();
        writer
            .set_storage_clsid(&["Root", "Sibling"], removed_sibling_clsid)
            .unwrap();
        writer
            .set_storage_clsid(&["RootSibling"], sibling_clsid)
            .unwrap();

        writer
            .create_stream(&["Root", "Nested", "Removed"], b"removed")
            .unwrap();
        writer
            .create_stream(&["Root", "Nested", "Deep", "AlsoRemoved"], b"removed")
            .unwrap();
        writer
            .create_stream(&["Root", "Sibling", "Removed"], b"removed")
            .unwrap();
        writer
            .create_stream(&["RootSibling", "Preserved"], b"preserved")
            .unwrap();

        writer.delete_storage(&["Root"]).unwrap();

        let root = vec!["Root".to_string()];
        let nested = vec!["Root".to_string(), "Nested".to_string()];
        let deep = vec!["Root".to_string(), "Nested".to_string(), "Deep".to_string()];
        let nested_sibling = vec!["Root".to_string(), "Sibling".to_string()];
        let root_sibling = vec!["RootSibling".to_string()];

        assert!(!writer.storages.contains_key(&root));
        assert!(!writer.storages.contains_key(&nested));
        assert!(!writer.storages.contains_key(&deep));
        assert!(!writer.storages.contains_key(&nested_sibling));
        assert!(writer.storages.contains_key(&root_sibling));

        assert!(!writer.storage_clsids.contains_key(&root));
        assert!(!writer.storage_clsids.contains_key(&nested));
        assert!(!writer.storage_clsids.contains_key(&deep));
        assert!(!writer.storage_clsids.contains_key(&nested_sibling));
        assert_eq!(
            writer.storage_clsids.get(&root_sibling),
            Some(&sibling_clsid)
        );

        assert!(
            writer
                .streams
                .iter()
                .all(|(path, _)| !path.starts_with(root.as_slice()))
        );
        assert!(writer.streams.iter().any(|(path, data)| {
            path == &["RootSibling".to_string(), "Preserved".to_string()] && data == b"preserved"
        }));
    }

    #[test]
    fn delete_storage_rejects_root_and_missing_paths_atomically() {
        let mut writer = OleWriter::new();
        writer.create_storage(&["Root"]).unwrap();
        writer.create_storage(&["Root", "Nested"]).unwrap();
        writer
            .create_stream(&["Root", "Nested", "Stream"], b"payload")
            .unwrap();
        writer
            .set_storage_clsid(&["Root", "Nested"], [0x55; 16])
            .unwrap();

        let streams_before = writer.streams.clone();
        let storages_before = writer.storages.clone();
        let clsids_before = writer.storage_clsids.clone();

        assert!(matches!(
            writer.delete_storage(&[]),
            Err(OleError::InvalidData(message)) if message == "Empty path"
        ));
        assert_eq!(writer.streams, streams_before);
        assert_eq!(writer.storages, storages_before);
        assert_eq!(writer.storage_clsids, clsids_before);

        assert!(matches!(
            writer.delete_storage(&["Missing"]),
            Err(OleError::InvalidFormat(message)) if message == "Storage not found"
        ));
        assert_eq!(writer.streams, streams_before);
        assert_eq!(writer.storages, storages_before);
        assert_eq!(writer.storage_clsids, clsids_before);

        assert!(matches!(
            writer.delete_storage(&["Root", "Nested", "Stream"]),
            Err(OleError::InvalidFormat(message)) if message == "Storage not found"
        ));
        assert_eq!(writer.streams, streams_before);
        assert_eq!(writer.storages, storages_before);
        assert_eq!(writer.storage_clsids, clsids_before);
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

    #[test]
    fn version_three_limits_are_checked_without_large_allocations() {
        let last_stream_byte = usize::try_from(V3_MAX_FILE_BYTES - 1).unwrap();
        let first_invalid_stream = usize::try_from(V3_MAX_FILE_BYTES).unwrap();
        assert!(validate_stream_size(512, last_stream_byte, "test stream").is_ok());
        assert!(validate_stream_size(512, first_invalid_stream, "test stream").is_err());
        assert!(validate_stream_size(4096, first_invalid_stream, "test stream").is_ok());

        let maximum_v3_sectors = u32::try_from(V3_MAX_FILE_BYTES / 512).unwrap();
        assert!(validate_output_size(512, maximum_v3_sectors - 1).is_ok());
        assert!(validate_output_size(512, maximum_v3_sectors).is_err());
        assert!(validate_output_size(4096, maximum_v3_sectors).is_ok());
    }

    #[test]
    fn allocation_table_planning_is_checked_and_converges() {
        assert_eq!(allocation_table_sector_counts(1, 512).unwrap(), (1, 0));
        assert_eq!(allocation_table_sector_counts(128, 512).unwrap(), (2, 0));
        assert!(allocation_table_sector_counts(MAXREGSECT, 512).is_err());
        assert!(allocation_table_sector_counts(1, 0).is_err());
    }

    #[test]
    fn sector_id_ranges_use_maxregsect_as_an_exclusive_end() {
        assert_eq!(
            sector_ids(MAXREGSECT - 1, 1, "test IDs").unwrap(),
            [MAXREGSECT - 1]
        );
        assert!(sector_ids(MAXREGSECT, 1, "test IDs").is_err());
    }

    #[test]
    fn allocation_errors_preserve_resource_and_source() {
        let mut probe = Vec::<u8>::new();
        let source = probe.try_reserve(usize::MAX).unwrap_err();
        let error = OleError::allocation("test buffer", source);
        assert!(matches!(
            &error,
            OleError::Allocation {
                resource: "test buffer",
                ..
            }
        ));
        assert!(error.source().is_some());
    }

    #[test]
    fn path_table_reservation_reports_overflow_without_mutation() {
        let mut table = HashMap::<Vec<String>, ()>::new();
        let key = vec!["Storage".to_string()];

        let error =
            reserve_hash_map_entry(&mut table, &key, usize::MAX, "test path table").unwrap_err();

        assert!(matches!(
            error,
            OleError::Allocation {
                resource: "test path table",
                ..
            }
        ));
        assert!(table.is_empty());
        assert!(!table.contains_key(&key));
    }
}
