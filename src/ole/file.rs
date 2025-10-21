use super::consts::*;
use std::io::{self, Read, Seek, SeekFrom};
use zerocopy::{FromBytes, LE, U16, U32, U64};
use zerocopy_derive::FromBytes as DeriveFromBytes;

/// Raw OLE directory entry structure (128 bytes)
///
/// This represents the on-disk format of a directory entry.
/// Based on Microsoft OLE2 specification.
#[derive(Debug, Clone, DeriveFromBytes)]
#[repr(C)]
struct RawDirectoryEntry {
    /// Entry name in UTF-16LE (64 bytes, null-padded)
    name: [u8; 64],
    /// Length of name in bytes (including null terminator)
    name_len: U16<LE>,
    /// Entry type (1 = storage, 2 = stream, 5 = root)
    entry_type: u8,
    /// Node color (0 = red, 1 = black)
    node_color: u8,
    /// Left sibling SID
    sid_left: U32<LE>,
    /// Right sibling SID
    sid_right: U32<LE>,
    /// Child SID
    sid_child: U32<LE>,
    /// CLSID (16 bytes)
    clsid: [u8; 16],
    /// State bits
    state_bits: U32<LE>,
    /// Creation time (FILETIME)
    creation_time: U64<LE>,
    /// Modified time (FILETIME)
    modified_time: U64<LE>,
    /// Starting sector
    start_sector: U32<LE>,
    /// Stream size
    stream_size: U64<LE>,
}

/// Main OLE file parser structure
///
/// This struct represents an OLE2 structured storage file and provides
/// methods to access its contents (streams and storages).
#[derive(Debug)]
pub struct OleFile<R: Read + Seek> {
    /// File handle or reader
    reader: R,
    /// Total file size in bytes
    file_size: u64,
    /// Sector size (512 or 4096 bytes)
    sector_size: usize,
    /// Mini sector size (typically 64 bytes)
    mini_sector_size: usize,
    /// Mini stream cutoff size (typically 4096 bytes)
    mini_stream_cutoff: u32,
    /// File Allocation Table - maps sector to next sector in chain
    fat: Vec<u32>,
    /// Mini FAT - for streams smaller than cutoff size
    minifat: Vec<u32>,
    /// First sector of directory stream
    first_dir_sector: u32,
    /// Root directory entry
    root: Option<DirectoryEntry>,
    /// All directory entries indexed by SID
    dir_entries: Vec<Option<DirectoryEntry>>,
    /// Mini stream data (loaded on demand)
    ministream: Option<Vec<u8>>,
}

/// Represents an OLE directory entry (stream or storage)
#[derive(Debug, Clone)]
pub struct DirectoryEntry {
    /// Storage ID (index in directory)
    pub sid: u32,
    /// Entry name (UTF-16 decoded to UTF-8)
    pub name: String,
    /// Entry type (stream, storage, root, etc.)
    pub entry_type: u8,
    /// Index of left sibling in red-black tree
    pub sid_left: u32,
    /// Index of right sibling in red-black tree
    pub sid_right: u32,
    /// Index of child node in red-black tree
    pub sid_child: u32,
    /// CLSID of this entry
    pub clsid: String,
    /// First sector of the stream
    pub start_sector: u32,
    /// Size of the stream in bytes
    pub size: u64,
    /// Whether this stream is in MiniFAT
    pub is_minifat: bool,
    /// Child entries (for storages)
    pub children: Vec<DirectoryEntry>,
}

/// Error types for OLE file parsing
#[derive(Debug)]
pub enum OleError {
    Io(io::Error),
    InvalidFormat(String),
    InvalidData(String),
    NotOleFile,
    CorruptedFile(String),
    StreamNotFound,
}

impl From<io::Error> for OleError {
    fn from(err: io::Error) -> Self {
        OleError::Io(err)
    }
}

impl From<crate::common::binary::BinaryError> for OleError {
    fn from(err: crate::common::binary::BinaryError) -> Self {
        OleError::InvalidData(err.to_string())
    }
}

impl std::fmt::Display for OleError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            OleError::Io(e) => write!(f, "IO error: {}", e),
            OleError::InvalidFormat(s) => write!(f, "Invalid format: {}", s),
            OleError::InvalidData(s) => write!(f, "Invalid data: {}", s),
            OleError::NotOleFile => write!(f, "Not an OLE file"),
            OleError::CorruptedFile(s) => write!(f, "Corrupted file: {}", s),
            OleError::StreamNotFound => write!(f, "Stream not found"),
        }
    }
}

impl std::error::Error for OleError {}

impl<R: Read + Seek> OleFile<R> {
    /// Open and parse an OLE file from a reader
    ///
    /// # Arguments
    /// * `reader` - A reader that implements Read + Seek
    ///
    /// # Returns
    /// * `Result<OleFile<R>, OleError>` - The parsed OLE file or an error
    pub fn open(mut reader: R) -> Result<Self, OleError> {
        // Get file size
        let file_size = reader.seek(SeekFrom::End(0))?;
        reader.seek(SeekFrom::Start(0))?;

        // Check minimum size
        if file_size < MINIMAL_OLEFILE_SIZE as u64 {
            return Err(OleError::NotOleFile);
        }

        // Read and validate header (512 bytes)
        let mut header = [0u8; 512];
        reader.read_exact(&mut header)?;

        // Validate magic bytes
        if &header[0..8] != MAGIC {
            return Err(OleError::NotOleFile);
        }

        // Parse header fields (little-endian)
        let dll_version = U16::<LE>::read_from_bytes(&header[0x1A..0x1C])
            .map(|v| v.get())
            .unwrap_or(0);
        let byte_order = U16::<LE>::read_from_bytes(&header[0x1C..0x1E])
            .map(|v| v.get())
            .unwrap_or(0);
        let sector_shift = U16::<LE>::read_from_bytes(&header[0x1E..0x20])
            .map(|v| v.get())
            .unwrap_or(0);
        let mini_sector_shift = U16::<LE>::read_from_bytes(&header[0x20..0x22])
            .map(|v| v.get())
            .unwrap_or(0);
        let first_dir_sector = U32::<LE>::read_from_bytes(&header[0x30..0x34])
            .map(|v| v.get())
            .unwrap_or(0);
        let mini_stream_cutoff = U32::<LE>::read_from_bytes(&header[0x38..0x3C])
            .map(|v| v.get())
            .unwrap_or(0);
        let first_minifat_sector = U32::<LE>::read_from_bytes(&header[0x3C..0x40])
            .map(|v| v.get())
            .unwrap_or(0);
        let num_minifat_sectors = U32::<LE>::read_from_bytes(&header[0x40..0x44])
            .map(|v| v.get())
            .unwrap_or(0);
        let first_difat_sector = U32::<LE>::read_from_bytes(&header[0x44..0x48])
            .map(|v| v.get())
            .unwrap_or(0);
        let num_difat_sectors = U32::<LE>::read_from_bytes(&header[0x48..0x4C])
            .map(|v| v.get())
            .unwrap_or(0);

        // Validate byte order (must be little-endian)
        if byte_order != 0xFFFE {
            return Err(OleError::InvalidFormat("Invalid byte order".to_string()));
        }

        // Calculate sector sizes
        let sector_size = 1usize << sector_shift;
        let mini_sector_size = 1usize << mini_sector_shift;

        // Validate sector size matches DLL version
        if (dll_version == 3 && sector_size != 512) || (dll_version == 4 && sector_size != 4096) {
            return Err(OleError::InvalidFormat("Sector size mismatch".to_string()));
        }

        let mut ole = OleFile {
            reader,
            file_size,
            sector_size,
            mini_sector_size,
            mini_stream_cutoff,
            fat: Vec::new(),
            minifat: Vec::new(),
            first_dir_sector,
            root: None,
            dir_entries: Vec::new(),
            ministream: None,
        };

        // Load FAT (File Allocation Table)
        ole.load_fat(&header, first_difat_sector, num_difat_sectors)?;

        // Load directory
        ole.load_directory()?;

        // Load MiniFAT if needed
        if num_minifat_sectors > 0 {
            ole.load_minifat(first_minifat_sector)?;
        }

        Ok(ole)
    }

    pub fn file_size(&self) -> u64 {
        self.file_size
    }

    /// Load the File Allocation Table (FAT)
    ///
    /// The FAT maps each sector to the next sector in the chain.
    /// First 109 FAT sector indexes are stored in the header, additional
    /// indexes are stored in DIFAT sectors.
    fn load_fat(
        &mut self,
        header: &[u8; 512],
        first_difat_sector: u32,
        num_difat_sectors: u32,
    ) -> Result<(), OleError> {
        // First 109 FAT sector indexes are in header at offset 0x4C
        let mut fat_sectors = Vec::new();
        for i in 0..109 {
            let offset = 0x4C + i * 4;
            if offset + 4 > 512 {
                break;
            }
            let sector = U32::<LE>::read_from_bytes(&header[offset..offset + 4])
                .map(|v| v.get())
                .unwrap_or(0);
            if sector == FREESECT || sector == ENDOFCHAIN {
                break;
            }
            fat_sectors.push(sector);
        }

        // Load additional FAT sectors from DIFAT if needed
        if num_difat_sectors > 0 {
            let mut difat_sector = first_difat_sector;
            let entries_per_sector = (self.sector_size / 4) - 1; // -1 for next DIFAT pointer

            for _ in 0..num_difat_sectors {
                let sector_data = self.read_sector(difat_sector)?;

                // Read FAT sector indexes from DIFAT sector
                for i in 0..entries_per_sector {
                    let offset = i * 4;
                    let sector = U32::<LE>::read_from_bytes(&sector_data[offset..offset + 4])
                        .map(|v| v.get())
                        .unwrap_or(0);
                    if sector == FREESECT || sector == ENDOFCHAIN {
                        break;
                    }
                    fat_sectors.push(sector);
                }

                // Get next DIFAT sector
                let next_offset = entries_per_sector * 4;
                difat_sector =
                    U32::<LE>::read_from_bytes(&sector_data[next_offset..next_offset + 4])
                        .map(|v| v.get())
                        .unwrap_or(0);

                if difat_sector == ENDOFCHAIN || difat_sector == FREESECT {
                    break;
                }
            }
        }

        // Now read all FAT sectors and build the FAT table
        let entries_per_sector = self.sector_size / 4;
        self.fat.reserve(fat_sectors.len() * entries_per_sector);

        for &sector_id in &fat_sectors {
            let sector_data = self.read_sector(sector_id)?;

            // Parse sector as array of u32 (little-endian)
            for i in 0..entries_per_sector {
                let offset = i * 4;
                let entry = U32::<LE>::read_from_bytes(&sector_data[offset..offset + 4])
                    .map(|v| v.get())
                    .unwrap_or(0);
                self.fat.push(entry);
            }
        }

        Ok(())
    }

    /// Load the Mini FAT (for small streams)
    fn load_minifat(&mut self, first_minifat_sector: u32) -> Result<(), OleError> {
        // Read the MiniFAT stream using the FAT
        let minifat_data = self.read_stream_from_fat(first_minifat_sector)?;

        // Parse as array of u32 (little-endian)
        let entries_count = minifat_data.len() / 4;
        self.minifat.reserve(entries_count);

        for i in 0..entries_count {
            let offset = i * 4;
            let entry = U32::<LE>::read_from_bytes(&minifat_data[offset..offset + 4])
                .map_err(|_| OleError::InvalidFormat("Failed to read u32".to_string()))?;
            self.minifat.push(entry.get());
        }

        Ok(())
    }

    /// Load directory entries
    fn load_directory(&mut self) -> Result<(), OleError> {
        // Read the entire directory stream
        let dir_data = self.read_stream_from_fat(self.first_dir_sector)?;

        // Each directory entry is 128 bytes
        let num_entries = dir_data.len() / DIRENTRY_SIZE;
        self.dir_entries = vec![None; num_entries];

        // Parse root entry first (always at index 0)
        if num_entries > 0 {
            let root = self.parse_directory_entry(&dir_data[0..DIRENTRY_SIZE], 0)?;
            let root_child_sid = root.sid_child;
            self.root = Some(root);

            // Build storage tree starting from root's children
            self.build_storage_tree(root_child_sid, &dir_data)?;
        }

        Ok(())
    }

    /// Parse a single directory entry from 128 bytes
    fn parse_directory_entry(&self, data: &[u8], sid: u32) -> Result<DirectoryEntry, OleError> {
        // Parse the raw directory entry
        let raw = RawDirectoryEntry::read_from_bytes(data)
            .map_err(|_| OleError::InvalidFormat("Failed to parse directory entry".to_string()))?;

        // Decode name from UTF-16LE
        let name_len = raw.name_len.get() as usize;
        let name_bytes = &raw.name[0..name_len.saturating_sub(2).min(64)];
        let name = decode_utf16le(name_bytes);

        // Format CLSID
        let clsid = format_clsid(&raw.clsid);

        // Handle size based on sector size (512-byte sectors only use low 32 bits)
        let size = if self.sector_size == 512 {
            raw.stream_size.get() & 0xFFFFFFFF
        } else {
            raw.stream_size.get()
        };

        // Determine if stream should use MiniFAT
        let is_minifat = size < self.mini_stream_cutoff as u64 && raw.entry_type == STGTY_STREAM;

        Ok(DirectoryEntry {
            sid,
            name,
            entry_type: raw.entry_type,
            sid_left: raw.sid_left.get(),
            sid_right: raw.sid_right.get(),
            sid_child: raw.sid_child.get(),
            clsid,
            start_sector: raw.start_sector.get(),
            size,
            is_minifat,
            children: Vec::new(),
        })
    }

    /// Recursively build storage tree from directory entries
    fn build_storage_tree(&mut self, child_sid: u32, dir_data: &[u8]) -> Result<(), OleError> {
        if child_sid == NOSTREAM {
            return Ok(());
        }

        let sid = child_sid as usize;
        if sid >= dir_data.len() / DIRENTRY_SIZE {
            return Err(OleError::CorruptedFile(
                "Invalid directory entry index".to_string(),
            ));
        }

        // Parse this entry if not already parsed
        if self.dir_entries[sid].is_none() {
            let offset = sid * DIRENTRY_SIZE;
            let entry =
                self.parse_directory_entry(&dir_data[offset..offset + DIRENTRY_SIZE], sid as u32)?;
            self.dir_entries[sid] = Some(entry);
        }

        // Get the entry (we know it exists now)
        let entry = self.dir_entries[sid].as_ref().unwrap();
        let left_sid = entry.sid_left;
        let right_sid = entry.sid_right;
        let child_sid = entry.sid_child;

        // Recursively build left subtree
        self.build_storage_tree(left_sid, dir_data)?;

        // Recursively build right subtree
        self.build_storage_tree(right_sid, dir_data)?;

        // Recursively build children if this is a storage
        self.build_storage_tree(child_sid, dir_data)?;

        Ok(())
    }

    /// Read a single sector from the file
    fn read_sector(&mut self, sector_id: u32) -> Result<Vec<u8>, OleError> {
        // Sector position in file: (sector_id + 1) * sector_size
        let position = ((sector_id as u64) + 1) * (self.sector_size as u64);
        self.reader.seek(SeekFrom::Start(position))?;

        let mut buffer = vec![0u8; self.sector_size];
        self.reader.read_exact(&mut buffer)?;
        Ok(buffer)
    }

    /// Read a stream by following the FAT chain
    fn read_stream_from_fat(&mut self, start_sector: u32) -> Result<Vec<u8>, OleError> {
        let mut data = Vec::new();
        let mut sector = start_sector;

        // Follow the chain in FAT
        loop {
            if sector == ENDOFCHAIN {
                break;
            }

            if sector >= self.fat.len() as u32 {
                return Err(OleError::CorruptedFile(
                    "Invalid sector index in FAT".to_string(),
                ));
            }

            // Read this sector
            let sector_data = self.read_sector(sector)?;
            data.extend_from_slice(&sector_data);

            // Get next sector from FAT
            sector = self.fat[sector as usize];
        }

        Ok(data)
    }

    /// Read a stream by following the MiniFAT chain
    fn read_stream_from_minifat(
        &mut self,
        start_sector: u32,
        size: u64,
    ) -> Result<Vec<u8>, OleError> {
        // Ensure ministream is loaded
        if self.ministream.is_none() {
            if let Some(ref root) = self.root {
                let ministream_data = self.read_stream_from_fat(root.start_sector)?;
                self.ministream = Some(ministream_data);
            } else {
                return Err(OleError::CorruptedFile("No root entry".to_string()));
            }
        }

        let ministream = self.ministream.as_ref().unwrap();
        let mut data = Vec::new();
        let mut sector = start_sector;

        // Follow the chain in MiniFAT
        loop {
            if sector == ENDOFCHAIN {
                break;
            }

            if sector >= self.minifat.len() as u32 {
                return Err(OleError::CorruptedFile(
                    "Invalid sector index in MiniFAT".to_string(),
                ));
            }

            // Calculate position in ministream
            let position = (sector as usize) * self.mini_sector_size;
            if position + self.mini_sector_size > ministream.len() {
                return Err(OleError::CorruptedFile(
                    "Mini sector out of bounds".to_string(),
                ));
            }

            // Read this mini sector
            data.extend_from_slice(&ministream[position..position + self.mini_sector_size]);

            // Get next sector from MiniFAT
            sector = self.minifat[sector as usize];
        }

        // Truncate to actual size
        data.truncate(size as usize);
        Ok(data)
    }

    /// List all streams in the OLE file
    ///
    /// Returns a list of stream paths (as vectors of storage/stream names)
    pub fn list_streams(&self) -> Vec<Vec<String>> {
        let mut streams = Vec::new();
        if let Some(ref root) = self.root {
            self.collect_streams(root, &mut Vec::new(), &mut streams);
        }
        streams
    }

    /// List all entries (streams and storages) in a directory
    ///
    /// # Arguments
    /// * `path` - Path to the directory as a slice of strings (empty for root)
    ///
    /// # Returns
    /// * `Result<Vec<DirectoryEntry>, OleError>` - List of directory entries
    pub fn list_directory_entries(&self, path: &[&str]) -> Result<Vec<DirectoryEntry>, OleError> {
        let mut entries = Vec::new();

        // Get the directory entry
        let dir_entry = if path.is_empty() {
            self.root.as_ref().ok_or(OleError::StreamNotFound)?
        } else {
            &self.find_entry(path)?
        };

        // Ensure it's a storage/directory
        if dir_entry.entry_type != STGTY_STORAGE && dir_entry.entry_type != STGTY_ROOT {
            return Err(OleError::InvalidFormat("Not a directory".to_string()));
        }

        // Collect children
        if dir_entry.sid_child != NOSTREAM {
            self.collect_directory_children(dir_entry.sid_child, &mut entries);
        }

        Ok(entries)
    }

    /// Recursively collect all children from a directory
    fn collect_directory_children(&self, sid: u32, entries: &mut Vec<DirectoryEntry>) {
        if sid == NOSTREAM || sid as usize >= self.dir_entries.len() {
            return;
        }

        if let Some(ref entry) = self.dir_entries[sid as usize] {
            // Traverse left
            if entry.sid_left != NOSTREAM {
                self.collect_directory_children(entry.sid_left, entries);
            }

            // Add current entry
            entries.push(entry.clone());

            // Traverse right
            if entry.sid_right != NOSTREAM {
                self.collect_directory_children(entry.sid_right, entries);
            }
        }
    }

    /// Check if a directory exists at the given path
    ///
    /// # Arguments
    /// * `path` - Path to check as a slice of strings
    ///
    /// # Returns
    /// * `bool` - True if directory exists
    pub fn directory_exists(&self, path: &[&str]) -> bool {
        match self.find_entry(path) {
            Ok(entry) => entry.entry_type == STGTY_STORAGE || entry.entry_type == STGTY_ROOT,
            Err(_) => false,
        }
    }

    /// Recursively collect streams from directory tree
    fn collect_streams(
        &self,
        entry: &DirectoryEntry,
        path: &mut [String],
        streams: &mut Vec<Vec<String>>,
    ) {
        // Add current entry to path
        let mut current_path = path.to_owned();
        if !entry.name.is_empty() && entry.entry_type != STGTY_ROOT {
            current_path.push(entry.name.clone());
        }

        // If this is a stream, add it to the list
        if entry.entry_type == STGTY_STREAM {
            streams.push(current_path);
            return;
        }

        // If this is a storage, recurse into children
        if entry.entry_type == STGTY_STORAGE || entry.entry_type == STGTY_ROOT {
            // Process children by traversing the red-black tree
            if entry.sid_child != NOSTREAM {
                self.traverse_children(entry.sid_child, &current_path, streams);
            }
        }
    }

    /// Traverse children in red-black tree
    fn traverse_children(&self, sid: u32, path: &Vec<String>, streams: &mut Vec<Vec<String>>) {
        if sid == NOSTREAM || sid as usize >= self.dir_entries.len() {
            return;
        }

        if let Some(ref entry) = self.dir_entries[sid as usize] {
            // Traverse left
            if entry.sid_left != NOSTREAM {
                self.traverse_children(entry.sid_left, path, streams);
            }

            // Process current
            let mut current_path = path.clone();
            self.collect_streams(entry, &mut current_path, streams);

            // Traverse right
            if entry.sid_right != NOSTREAM {
                self.traverse_children(entry.sid_right, path, streams);
            }
        }
    }

    /// Open a stream by path and return its contents
    ///
    /// # Arguments
    /// * `path` - Path to the stream as a slice of strings
    ///
    /// # Returns
    /// * `Result<Vec<u8>, OleError>` - Stream contents or error
    pub fn open_stream(&mut self, path: &[&str]) -> Result<Vec<u8>, OleError> {
        // Find the entry
        let entry = self.find_entry(path)?;

        // Ensure it's a stream
        if entry.entry_type != STGTY_STREAM {
            return Err(OleError::InvalidFormat("Not a stream".to_string()));
        }

        // Read the stream based on whether it uses FAT or MiniFAT
        if entry.is_minifat {
            self.read_stream_from_minifat(entry.start_sector, entry.size)
        } else {
            let mut data = self.read_stream_from_fat(entry.start_sector)?;
            data.truncate(entry.size as usize);
            Ok(data)
        }
    }

    /// Find a directory entry by path
    fn find_entry(&self, path: &[&str]) -> Result<DirectoryEntry, OleError> {
        if path.is_empty() {
            return self.root.clone().ok_or(OleError::StreamNotFound);
        }

        // Start from root
        let root = self.root.as_ref().ok_or(OleError::StreamNotFound)?;
        let mut current_sid = root.sid_child;

        // Traverse path
        for (i, &name) in path.iter().enumerate() {
            let entry = self.find_child_by_name(current_sid, name)?;

            // If this is the last component, return it
            if i == path.len() - 1 {
                return Ok(entry);
            }

            // Otherwise, move to its children
            current_sid = entry.sid_child;
        }

        Err(OleError::StreamNotFound)
    }

    /// Find a child entry by name in a red-black tree
    fn find_child_by_name(&self, sid: u32, name: &str) -> Result<DirectoryEntry, OleError> {
        if sid == NOSTREAM || sid as usize >= self.dir_entries.len() {
            return Err(OleError::StreamNotFound);
        }

        let entry = self.dir_entries[sid as usize]
            .as_ref()
            .ok_or(OleError::StreamNotFound)?;

        // Case-insensitive comparison
        if entry.name.to_lowercase() == name.to_lowercase() {
            return Ok(entry.clone());
        }

        // Search left subtree
        if entry.sid_left != NOSTREAM
            && let Ok(found) = self.find_child_by_name(entry.sid_left, name)
        {
            return Ok(found);
        }

        // Search right subtree
        if entry.sid_right != NOSTREAM
            && let Ok(found) = self.find_child_by_name(entry.sid_right, name)
        {
            return Ok(found);
        }

        Err(OleError::StreamNotFound)
    }

    /// Get the root entry name
    pub fn get_root_name(&self) -> Option<&str> {
        self.root.as_ref().map(|r| r.name.as_str())
    }

    /// Check if a stream exists
    pub fn exists(&self, path: &[&str]) -> bool {
        self.find_entry(path).is_ok()
    }
}

/// Decode UTF-16LE bytes to String
fn decode_utf16le(bytes: &[u8]) -> String {
    let mut utf16_chars = Vec::new();

    for chunk in bytes.chunks_exact(2) {
        let code_unit = U16::<LE>::read_from_bytes(chunk)
            .map(|v| v.get())
            .unwrap_or(0);
        utf16_chars.push(code_unit);
    }

    // Decode UTF-16 to String, replacing invalid sequences
    String::from_utf16_lossy(&utf16_chars)
        .trim_end_matches('\0')
        .to_string()
}

/// Format CLSID as a human-readable string
fn format_clsid(bytes: &[u8]) -> String {
    if bytes.len() != 16 {
        return String::new();
    }

    // Check if all zeros
    if bytes.iter().all(|&b| b == 0) {
        return String::new();
    }

    // Format as: {XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}
    format!(
        "{:08X}-{:04X}-{:04X}-{:02X}{:02X}-{:02X}{:02X}{:02X}{:02X}{:02X}{:02X}",
        U32::<LE>::read_from_bytes(&bytes[0..4])
            .map(|v| v.get())
            .unwrap_or(0),
        U16::<LE>::read_from_bytes(&bytes[4..6])
            .map(|v| v.get())
            .unwrap_or(0),
        U16::<LE>::read_from_bytes(&bytes[6..8])
            .map(|v| v.get())
            .unwrap_or(0),
        bytes[8],
        bytes[9],
        bytes[10],
        bytes[11],
        bytes[12],
        bytes[13],
        bytes[14],
        bytes[15],
    )
}

/// Check if a file/data is an OLE file by checking magic bytes
pub fn is_ole_file(data: &[u8]) -> bool {
    data.len() >= MINIMAL_OLEFILE_SIZE && &data[0..8] == MAGIC
}
