use super::consts::*;
use crate::directory_name::{DirectoryNameData, directory_name_data};
use fixedbitset::FixedBitSet;
use smallvec::SmallVec;
use std::cmp::Ordering;
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[repr(u8)]
enum DirectoryNodeColor {
    Red = 0,
    Black = 1,
}

impl TryFrom<u8> for DirectoryNodeColor {
    type Error = OleError;

    fn try_from(value: u8) -> Result<Self, Self::Error> {
        match value {
            0 => Ok(Self::Red),
            1 => Ok(Self::Black),
            _ => Err(OleError::CorruptedFile(format!(
                "invalid CFB directory node color {value}"
            ))),
        }
    }
}

#[derive(Debug)]
struct ValidatedDirectoryEntry {
    sid: u32,
    name: String,
    name_data: DirectoryNameData,
    entry_type: u8,
    node_color: DirectoryNodeColor,
    sid_left: u32,
    sid_right: u32,
    sid_child: u32,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
enum PhysicalSectorRole {
    #[default]
    Unclaimed,
    Fat,
    Difat,
    Directory,
    MiniFat,
    MiniStream,
    RegularStream,
}

impl PhysicalSectorRole {
    fn label(self) -> &'static str {
        match self {
            Self::Unclaimed => "unclaimed",
            Self::Fat => "FAT",
            Self::Difat => "DIFAT",
            Self::Directory => "directory",
            Self::MiniFat => "MiniFAT",
            Self::MiniStream => "mini stream",
            Self::RegularStream => "regular stream",
        }
    }
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
    /// Exclusive ownership of every physical sector in the file.
    sector_roles: Vec<PhysicalSectorRole>,
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
    /// A bounded writer could not reserve the memory required for `resource`.
    Allocation {
        resource: &'static str,
        source: std::collections::TryReserveError,
    },
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

impl OleError {
    pub(crate) fn allocation(
        resource: &'static str,
        source: std::collections::TryReserveError,
    ) -> Self {
        Self::Allocation { resource, source }
    }
}

impl From<litchi_core::binary::BinaryError> for OleError {
    fn from(err: litchi_core::binary::BinaryError) -> Self {
        OleError::InvalidData(err.to_string())
    }
}

impl std::fmt::Display for OleError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            OleError::Io(e) => write!(f, "IO error: {}", e),
            OleError::Allocation { resource, source } => {
                write!(f, "could not reserve memory for CFB {resource}: {source}")
            },
            OleError::InvalidFormat(s) => write!(f, "Invalid format: {}", s),
            OleError::InvalidData(s) => write!(f, "Invalid data: {}", s),
            OleError::NotOleFile => write!(f, "Not an OLE file"),
            OleError::CorruptedFile(s) => write!(f, "Corrupted file: {}", s),
            OleError::StreamNotFound => write!(f, "Stream not found"),
        }
    }
}

impl std::error::Error for OleError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(source) => Some(source),
            Self::Allocation { source, .. } => Some(source),
            _ => None,
        }
    }
}

// Convert CFB-substrate errors into the unified `litchi_core::Error`.
//
// The reverse direction (`OleError -> litchi_core::Error`) used to live in the
// umbrella crate's `error_ext.rs`, but the orphan rule forbids implementing
// `From<external> for external` outside the crate that defines either side.
// This impl is local to `litchi-cfb` because `OleError` is defined here.
impl From<OleError> for litchi_core::Error {
    fn from(err: OleError) -> Self {
        match err {
            OleError::Io(e) => litchi_core::Error::Io(e),
            OleError::Allocation { resource, source } => litchi_core::Error::Other(format!(
                "could not reserve memory for CFB {resource}: {source}"
            )),
            OleError::InvalidFormat(s) => litchi_core::Error::InvalidFormat(s),
            OleError::InvalidData(s) => litchi_core::Error::InvalidFormat(s),
            OleError::NotOleFile => litchi_core::Error::NotOfficeFile,
            OleError::CorruptedFile(s) => litchi_core::Error::CorruptedFile(s),
            OleError::StreamNotFound => {
                litchi_core::Error::ComponentNotFound("Stream not found".to_string())
            },
        }
    }
}

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
        let num_dir_sectors = U32::<LE>::read_from_bytes(&header[0x28..0x2C])
            .map(|v| v.get())
            .unwrap_or(0);
        let first_dir_sector = U32::<LE>::read_from_bytes(&header[0x30..0x34])
            .map(|v| v.get())
            .unwrap_or(0);
        let num_fat_sectors = U32::<LE>::read_from_bytes(&header[0x2C..0x30])
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
        if header[0x22..0x28] != [0; 6] {
            return Err(OleError::InvalidFormat(
                "Reserved CFB header bytes must be zero".to_string(),
            ));
        }

        // Validate shifts before using them. Untrusted shift counts must never reach `<<`.
        let expected_sector_shift = match dll_version {
            3 => 9,
            4 => 12,
            _ => {
                return Err(OleError::InvalidFormat(format!(
                    "Unsupported CFB major version {dll_version}"
                )));
            },
        };
        if sector_shift != expected_sector_shift {
            return Err(OleError::InvalidFormat(format!(
                "Invalid sector shift {sector_shift} for CFB version {dll_version}"
            )));
        }
        if (dll_version == 3 && num_dir_sectors != 0) || (dll_version == 4 && num_dir_sectors == 0)
        {
            return Err(OleError::InvalidFormat(format!(
                "Invalid directory-sector count {num_dir_sectors} for CFB version {dll_version}"
            )));
        }
        if mini_sector_shift != 6 {
            return Err(OleError::InvalidFormat(format!(
                "Invalid mini sector shift {mini_sector_shift}"
            )));
        }
        if mini_stream_cutoff != 4096 {
            return Err(OleError::InvalidFormat(format!(
                "Invalid mini stream cutoff {mini_stream_cutoff}"
            )));
        }
        let sector_size = 1usize << sector_shift;
        let mini_sector_size = 1usize << mini_sector_shift;
        if file_size < sector_size as u64 {
            return Err(OleError::InvalidFormat(
                "File is smaller than its CFB header sector".to_string(),
            ));
        }
        // MS-CFB does not require the file length to be a whole number of
        // sectors, and real documents are routinely stored truncated at the end
        // of their last used sector. A short final sector is therefore accepted
        // and reads as zeroes past the end of the file; a sector that starts at
        // or beyond the end is still rejected.
        if (num_minifat_sectors == 0) != (first_minifat_sector == ENDOFCHAIN) {
            return Err(OleError::InvalidFormat(
                "MiniFAT start sector and sector count disagree".to_string(),
            ));
        }
        if (num_difat_sectors == 0) != (first_difat_sector == ENDOFCHAIN) {
            return Err(OleError::InvalidFormat(
                "DIFAT start sector and sector count disagree".to_string(),
            ));
        }
        if first_dir_sector >= MAXREGSECT {
            return Err(OleError::InvalidFormat(
                "Directory starts at an invalid sector".to_string(),
            ));
        }
        let physical_sector_count = usize::try_from(file_size / sector_size as u64 - 1)
            .map_err(|_| OleError::InvalidFormat("Too many physical sectors".to_string()))?;

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
            sector_roles: vec![PhysicalSectorRole::Unclaimed; physical_sector_count],
        };

        // Load FAT (File Allocation Table)
        ole.load_fat(
            &header,
            num_fat_sectors,
            first_difat_sector,
            num_difat_sectors,
        )?;

        // Load directory
        ole.load_directory((dll_version == 4).then_some(num_dir_sectors))?;

        // Load MiniFAT if needed
        if num_minifat_sectors > 0 {
            ole.load_minifat(first_minifat_sector, num_minifat_sectors)?;
        }
        ole.validate_stream_allocations()?;

        Ok(ole)
    }

    pub fn file_size(&self) -> u64 {
        self.file_size
    }

    /// Sector size used by this compound file.
    pub fn sector_size(&self) -> usize {
        self.sector_size
    }

    /// Load the File Allocation Table (FAT)
    ///
    /// The FAT maps each sector to the next sector in the chain.
    /// First 109 FAT sector indexes are stored in the header, additional
    /// indexes are stored in DIFAT sectors.
    fn load_fat(
        &mut self,
        header: &[u8; 512],
        num_fat_sectors: u32,
        first_difat_sector: u32,
        num_difat_sectors: u32,
    ) -> Result<(), OleError> {
        let physical_sector_count = self.sector_roles.len() as u64;
        if u64::from(num_fat_sectors) > physical_sector_count {
            return Err(OleError::CorruptedFile(
                "Declared FAT sector count exceeds the physical file".to_string(),
            ));
        }
        let expected_fat_sectors = usize::try_from(num_fat_sectors)
            .map_err(|_| OleError::CorruptedFile("FAT sector count is too large".to_string()))?;

        // First 109 FAT sector indexes are in header at offset 0x4C
        let mut fat_sectors = Vec::with_capacity(expected_fat_sectors);
        let mut difat_sectors = Vec::with_capacity(num_difat_sectors as usize);
        let header_fat_count = HEADER_DIFAT_ENTRIES.min(expected_fat_sectors);
        for i in 0..header_fat_count {
            let offset = HEADER_DIFAT_OFFSET + i * 4;
            let sector = U32::<LE>::read_from_bytes(&header[offset..offset + 4])
                .map(|v| v.get())
                .unwrap_or(0);
            if sector == FREESECT || sector == ENDOFCHAIN {
                return Err(OleError::CorruptedFile(
                    "FAT sector list ends before its declared count".to_string(),
                ));
            }
            self.claim_sector(sector, PhysicalSectorRole::Fat)?;
            fat_sectors.push(sector);
        }
        // Entries past the declared FAT sector count are not part of the FAT
        // sector list. MS-CFB 2.2 describes the header DIFAT only as holding
        // "the first 109 FAT sector locations" and never constrains the unused
        // tail, so writers leave zeroes or stale values there. The count field
        // already says where the list ends, and the FAT chain validation below
        // catches a count that disagrees with the file, so the tail is ignored.

        let mut difat_sector = first_difat_sector;
        let entries_per_sector = (self.sector_size / 4) - 1;
        for difat_index in 0..num_difat_sectors as usize {
            self.claim_sector(difat_sector, PhysicalSectorRole::Difat)?;
            difat_sectors.push(difat_sector);
            let sector_data = self.read_sector(difat_sector)?;

            for i in 0..entries_per_sector {
                let offset = i * 4;
                let sector = U32::<LE>::read_from_bytes(&sector_data[offset..offset + 4])
                    .map(|v| v.get())
                    .unwrap_or(0);
                if fat_sectors.len() < expected_fat_sectors {
                    if sector >= MAXREGSECT {
                        return Err(OleError::CorruptedFile(
                            "DIFAT sector list ends before the declared FAT count".to_string(),
                        ));
                    }
                    self.claim_sector(sector, PhysicalSectorRole::Fat)?;
                    fat_sectors.push(sector);
                } else if sector != FREESECT {
                    return Err(OleError::CorruptedFile(
                        "Unused DIFAT entries must be FREESECT".to_string(),
                    ));
                }
            }

            let next_offset = entries_per_sector * 4;
            let next = U32::<LE>::read_from_bytes(&sector_data[next_offset..next_offset + 4])
                .map(|v| v.get())
                .unwrap_or(0);
            if difat_index + 1 == num_difat_sectors as usize {
                if next != ENDOFCHAIN {
                    return Err(OleError::CorruptedFile(
                        "DIFAT chain exceeds its declared length".to_string(),
                    ));
                }
            } else if next >= MAXREGSECT {
                return Err(OleError::CorruptedFile(
                    "DIFAT chain ends before its declared length".to_string(),
                ));
            }
            difat_sector = next;
        }
        if fat_sectors.len() != expected_fat_sectors {
            return Err(OleError::CorruptedFile(format!(
                "Expected {expected_fat_sectors} FAT sectors, found {}",
                fat_sectors.len()
            )));
        }

        // Now read all FAT sectors and build the FAT table
        let entries_per_sector = self.sector_size / 4;

        // Pre-allocate with exact capacity needed (optimization)
        self.fat = Vec::with_capacity(fat_sectors.len() * entries_per_sector);

        for &sector_id in &fat_sectors {
            let sector_data = self.read_sector(sector_id)?;

            // Parse sector as array of u32 (little-endian) - use chunks for efficiency
            for chunk in sector_data.chunks_exact(4) {
                let entry = U32::<LE>::read_from_bytes(chunk)
                    .map(|v| v.get())
                    .unwrap_or(0);
                self.fat.push(entry);
            }
        }

        for sector in fat_sectors {
            if self.fat.get(sector as usize) != Some(&FATSECT) {
                return Err(OleError::CorruptedFile(format!(
                    "FAT sector {sector} is not marked FATSECT"
                )));
            }
        }
        for sector in difat_sectors {
            if self.fat.get(sector as usize) != Some(&DIFSECT) {
                return Err(OleError::CorruptedFile(format!(
                    "DIFAT sector {sector} is not marked DIFSECT"
                )));
            }
        }

        Ok(())
    }

    /// Load the Mini FAT (for small streams)
    fn load_minifat(
        &mut self,
        first_minifat_sector: u32,
        sector_count: u32,
    ) -> Result<(), OleError> {
        let sectors = collect_sector_chain_exact(
            &self.fat,
            first_minifat_sector,
            sector_count as usize,
            "MiniFAT",
        )?;
        self.claim_chain(&sectors, PhysicalSectorRole::MiniFat)?;
        let mut minifat_data = Vec::with_capacity(sectors.len() * self.sector_size);
        for sector in sectors {
            minifat_data.extend_from_slice(&self.read_sector(sector)?);
        }

        // Parse as array of u32 (little-endian) - use chunks for efficiency
        let entries_count = minifat_data.len() / 4;
        self.minifat = Vec::with_capacity(entries_count);

        for chunk in minifat_data.chunks_exact(4) {
            let entry = U32::<LE>::read_from_bytes(chunk)
                .map_err(|_| OleError::InvalidFormat("Failed to read u32".to_string()))?;
            self.minifat.push(entry.get());
        }

        Ok(())
    }

    /// Load directory entries with optimized iterative parsing
    fn load_directory(&mut self, declared_sector_count: Option<u32>) -> Result<(), OleError> {
        let sectors = match declared_sector_count {
            Some(count) => collect_sector_chain_exact(
                &self.fat,
                self.first_dir_sector,
                count as usize,
                "directory",
            )?,
            None => collect_sector_chain(&self.fat, self.first_dir_sector, "directory")?,
        };
        self.claim_chain(&sectors, PhysicalSectorRole::Directory)?;
        let mut dir_data = Vec::with_capacity(sectors.len() * self.sector_size);
        for sector in sectors {
            dir_data.extend_from_slice(&self.read_sector(sector)?);
        }

        Self::validate_directory(&dir_data, self.sector_size)?;

        // Each directory entry is 128 bytes
        let num_entries = dir_data.len() / DIRENTRY_SIZE;
        self.dir_entries = vec![None; num_entries];

        // Parse root entry first (always at index 0)
        if num_entries > 0 {
            let root = self.parse_directory_entry(&dir_data[0..DIRENTRY_SIZE], 0)?;
            let root_child_sid = root.sid_child;
            self.root = Some(root);

            // Build storage tree using iterative approach (avoids recursion overhead)
            self.build_storage_tree_iterative(root_child_sid, &dir_data)?;
        }

        Ok(())
    }

    fn claim_sector(&mut self, sector: u32, role: PhysicalSectorRole) -> Result<(), OleError> {
        let slot = self.sector_roles.get_mut(sector as usize).ok_or_else(|| {
            OleError::CorruptedFile(format!(
                "{} sector {sector} is outside the file",
                role.label()
            ))
        })?;
        if *slot != PhysicalSectorRole::Unclaimed {
            return Err(OleError::CorruptedFile(format!(
                "Sector {sector} is claimed by both {} and {}",
                slot.label(),
                role.label()
            )));
        }
        *slot = role;
        Ok(())
    }

    fn claim_chain(&mut self, sectors: &[u32], role: PhysicalSectorRole) -> Result<(), OleError> {
        for &sector in sectors {
            self.claim_sector(sector, role)?;
        }
        Ok(())
    }

    fn validate_stream_allocations(&mut self) -> Result<(), OleError> {
        let root = self
            .root
            .as_ref()
            .ok_or_else(|| OleError::CorruptedFile("Missing root directory entry".to_string()))?;
        let root_start = root.start_sector;
        let root_size = root.size;
        let root_sector_count = usize::try_from(root_size.div_ceil(self.sector_size as u64))
            .map_err(|_| OleError::CorruptedFile("Root mini stream is too large".to_string()))?;
        let root_chain = collect_sector_chain_exact(
            &self.fat,
            root_start,
            root_sector_count,
            "root mini stream",
        )?;
        self.claim_chain(&root_chain, PhysicalSectorRole::MiniStream)?;

        let mini_sector_capacity =
            usize::try_from(root_size.div_ceil(self.mini_sector_size as u64)).map_err(|_| {
                OleError::CorruptedFile("Root mini stream is too large".to_string())
            })?;
        let mut claimed_mini_sectors = FixedBitSet::with_capacity(mini_sector_capacity);
        let streams: Vec<_> = self
            .dir_entries
            .iter()
            .flatten()
            .filter(|entry| entry.entry_type == STGTY_STREAM)
            .map(|entry| (entry.is_minifat, entry.start_sector, entry.size))
            .collect();

        for (is_minifat, start_sector, size) in streams {
            if is_minifat {
                let sector_count = usize::try_from(size.div_ceil(self.mini_sector_size as u64))
                    .map_err(|_| OleError::CorruptedFile("Mini stream is too large".to_string()))?;
                let chain = collect_sector_chain_exact(
                    &self.minifat,
                    start_sector,
                    sector_count,
                    "mini stream",
                )?;
                for sector in chain {
                    let sector = sector as usize;
                    if sector >= mini_sector_capacity {
                        return Err(OleError::CorruptedFile(
                            "Mini stream references storage outside the root mini stream"
                                .to_string(),
                        ));
                    }
                    if claimed_mini_sectors.contains(sector) {
                        return Err(OleError::CorruptedFile(format!(
                            "Mini sector {sector} is claimed by multiple streams"
                        )));
                    }
                    claimed_mini_sectors.insert(sector);
                }
            } else {
                let sector_count = usize::try_from(size.div_ceil(self.sector_size as u64))
                    .map_err(|_| {
                        OleError::CorruptedFile("Regular stream is too large".to_string())
                    })?;
                let chain = collect_sector_chain_exact(
                    &self.fat,
                    start_sector,
                    sector_count,
                    "regular stream",
                )?;
                self.claim_chain(&chain, PhysicalSectorRole::RegularStream)?;
            }
        }
        Ok(())
    }

    pub(crate) fn validate_directory(dir_data: &[u8], sector_size: usize) -> Result<(), OleError> {
        if dir_data.is_empty() || !dir_data.len().is_multiple_of(DIRENTRY_SIZE) {
            return Err(OleError::CorruptedFile(
                "CFB directory stream must contain complete 128-byte entries".to_string(),
            ));
        }

        let mut entries = Vec::with_capacity(dir_data.len() / DIRENTRY_SIZE);
        for (sid, data) in dir_data.chunks_exact(DIRENTRY_SIZE).enumerate() {
            let sid = u32::try_from(sid).map_err(|_| {
                OleError::CorruptedFile("CFB directory contains too many entries".to_string())
            })?;
            entries.push(Self::parse_validated_directory_entry(
                data,
                sid,
                sector_size,
            )?);
        }

        let root = entries
            .first()
            .and_then(Option::as_ref)
            .ok_or_else(|| OleError::CorruptedFile("CFB root entry is missing".to_string()))?;
        if root.entry_type != STGTY_ROOT
            || root.name != "Root Entry"
            || root.sid_left != NOSTREAM
            || root.sid_right != NOSTREAM
        {
            return Err(OleError::CorruptedFile(
                "invalid CFB root directory entry".to_string(),
            ));
        }

        let mut owned = FixedBitSet::with_capacity(entries.len());
        owned.insert(0);
        let mut pending_trees = vec![root.sid_child];
        while let Some(tree_root) = pending_trees.pop() {
            if tree_root == NOSTREAM {
                continue;
            }
            let mut stack = vec![(tree_root, None, None, 0usize)];
            while let Some((sid, lower, upper, black_depth)) = stack.pop() {
                if sid == NOSTREAM {
                    // Some widely deployed Office producers wrote unbalanced
                    // color metadata. Bounds, ordering, acyclicity, and unique
                    // ownership are sufficient for safe deterministic traversal.
                    continue;
                }

                let entry = Self::validated_entry(&entries, sid)?;
                if owned.contains(sid as usize) {
                    return Err(OleError::CorruptedFile(format!(
                        "CFB directory tree contains repeated SID {sid} or cross-storage ownership"
                    )));
                }
                owned.insert(sid as usize);

                let violates_lower = if let Some(bound) = lower {
                    Self::compare_validated(Self::validated_entry(&entries, bound)?, entry)
                        != Ordering::Less
                } else {
                    false
                };
                let violates_upper = if let Some(bound) = upper {
                    Self::compare_validated(entry, Self::validated_entry(&entries, bound)?)
                        != Ordering::Less
                } else {
                    false
                };
                if violates_lower || violates_upper {
                    return Err(OleError::CorruptedFile(format!(
                        "CFB sibling tree violates name ordering at SID {sid}"
                    )));
                }

                if entry.entry_type == STGTY_STORAGE && entry.sid_child != NOSTREAM {
                    pending_trees.push(entry.sid_child);
                }
                let black_depth =
                    black_depth + usize::from(entry.node_color == DirectoryNodeColor::Black);
                stack.push((entry.sid_right, Some(sid), upper, black_depth));
                stack.push((entry.sid_left, lower, Some(sid), black_depth));
            }
        }

        for entry in entries.iter().flatten() {
            if !owned.contains(entry.sid as usize) {
                return Err(OleError::CorruptedFile(format!(
                    "CFB directory SID {} is not owned by a storage",
                    entry.sid
                )));
            }
        }
        Ok(())
    }

    fn parse_validated_directory_entry(
        data: &[u8],
        sid: u32,
        sector_size: usize,
    ) -> Result<Option<ValidatedDirectoryEntry>, OleError> {
        let raw = RawDirectoryEntry::read_from_bytes(data)
            .map_err(|_| OleError::InvalidFormat("Failed to parse directory entry".to_string()))?;
        if raw.entry_type == STGTY_EMPTY {
            if raw.name_len.get() != 0 {
                return Err(OleError::CorruptedFile(format!(
                    "empty CFB directory SID {sid} has a nonzero name length"
                )));
            }
            return Ok(None);
        }
        if !matches!(raw.entry_type, STGTY_STORAGE | STGTY_STREAM | STGTY_ROOT) {
            return Err(OleError::CorruptedFile(format!(
                "invalid CFB directory entry type {} at SID {sid}",
                raw.entry_type
            )));
        }

        let name_len = usize::from(raw.name_len.get());
        if !(2..=64).contains(&name_len) || name_len % 2 != 0 {
            return Err(OleError::CorruptedFile(format!(
                "invalid CFB directory name length {name_len} at SID {sid}"
            )));
        }
        let mut name_utf16 = SmallVec::<[u16; 32]>::with_capacity(name_len / 2);
        for pair in raw.name[..name_len].chunks_exact(2) {
            name_utf16.push(u16::from_le_bytes([pair[0], pair[1]]));
        }
        // Some classic Mac Excel writers store SID 0 as the two-byte sequence
        // `00 52` (an abbreviated, byte-swapped "R") instead of a terminated
        // UTF-16LE "Root Entry". Accept only that exact root-entry encoding and
        // canonicalize it; all other directory names remain strictly validated.
        let legacy_mac_root = sid == 0
            && raw.entry_type == STGTY_ROOT
            && name_len == 2
            && raw.name[0] == 0
            && raw.name[1] == b'R'
            && raw.name[2..].iter().all(|&byte| byte == 0);
        if legacy_mac_root {
            name_utf16.clear();
            name_utf16.extend("Root Entry".encode_utf16());
        } else if name_utf16.pop() != Some(0) || name_utf16.contains(&0) {
            return Err(OleError::CorruptedFile(format!(
                "CFB directory name at SID {sid} is not correctly NUL-terminated"
            )));
        }
        let name = String::from_utf16(&name_utf16).map_err(|_| {
            OleError::CorruptedFile(format!("invalid UTF-16 CFB directory name at SID {sid}"))
        })?;
        let name_data = directory_name_data(&name)
            .map_err(|error| OleError::CorruptedFile(error.to_string()))?;
        if name_data.utf16.as_slice() != name_utf16.as_slice() {
            return Err(OleError::CorruptedFile(format!(
                "CFB directory name encoding mismatch at SID {sid}"
            )));
        }

        let node_color = DirectoryNodeColor::try_from(raw.node_color)?;
        let sid_left = raw.sid_left.get();
        let sid_right = raw.sid_right.get();
        let sid_child = raw.sid_child.get();
        // MS-CFB 2.6.1: a version 3 file's stream size must fit in 32 bits, but
        // older writers left the high word uninitialized. The spec explicitly
        // recommends that parsers ignore those bits and treat them as zero
        // rather than reject the file, so mask instead of failing — matching
        // what `DirectoryEntry` reading already does.
        let stream_size = mask_v3_stream_size(raw.stream_size.get(), sector_size);

        match raw.entry_type {
            STGTY_ROOT if sid != 0 => {
                return Err(OleError::CorruptedFile(
                    "CFB root entry must have SID 0".to_string(),
                ));
            },
            STGTY_STORAGE
                if sid == 0
                    || !matches!(raw.start_sector.get(), 0 | ENDOFCHAIN)
                    || stream_size != 0 =>
            {
                return Err(OleError::CorruptedFile(format!(
                    "invalid CFB storage fields at SID {sid}"
                )));
            },
            STGTY_STORAGE => {},
            STGTY_STREAM if sid == 0 || sid_child != NOSTREAM => {
                return Err(OleError::CorruptedFile(format!(
                    "invalid CFB stream fields at SID {sid}"
                )));
            },
            STGTY_STREAM => {},
            _ => {},
        }

        Ok(Some(ValidatedDirectoryEntry {
            sid,
            name,
            name_data,
            entry_type: raw.entry_type,
            node_color,
            sid_left,
            sid_right,
            sid_child,
        }))
    }

    fn validated_entry(
        entries: &[Option<ValidatedDirectoryEntry>],
        sid: u32,
    ) -> Result<&ValidatedDirectoryEntry, OleError> {
        entries
            .get(sid as usize)
            .and_then(Option::as_ref)
            .ok_or_else(|| {
                OleError::CorruptedFile(format!("invalid or empty CFB directory SID {sid}"))
            })
    }

    fn compare_validated(
        left: &ValidatedDirectoryEntry,
        right: &ValidatedDirectoryEntry,
    ) -> Ordering {
        left.name_data.compare(&right.name_data)
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

        // Version 3 sizes only use the low 32 bits; see `mask_v3_stream_size`.
        let size = mask_v3_stream_size(raw.stream_size.get(), self.sector_size);

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

    /// Build storage tree using iterative approach (optimized, no recursion)
    ///
    /// This replaces the recursive `build_storage_tree` with an iterative
    /// traversal using a work queue, eliminating function call overhead.
    /// Uses FixedBitSet for better cache locality and memory efficiency.
    fn build_storage_tree_iterative(
        &mut self,
        root_sid: u32,
        dir_data: &[u8],
    ) -> Result<(), OleError> {
        if root_sid == NOSTREAM {
            return Ok(());
        }

        let max_entries = dir_data.len() / DIRENTRY_SIZE;

        // Use a work queue for iterative traversal (pre-allocate for common case)
        let mut queue = Vec::with_capacity(64);
        queue.push(root_sid);

        // Track visited SIDs using FixedBitSet for better cache locality
        // Uses ~8x less memory than Vec<bool> (1 bit vs 1 byte per entry)
        let mut visited = FixedBitSet::with_capacity(max_entries);

        while let Some(sid) = queue.pop() {
            if sid == NOSTREAM {
                continue;
            }

            let sid_usize = sid as usize;

            // Validate SID
            if sid_usize >= max_entries {
                return Err(OleError::CorruptedFile(
                    "Invalid directory entry index".to_string(),
                ));
            }

            // Every directory SID can occur only once in the storage tree.
            if visited.contains(sid_usize) {
                return Err(OleError::CorruptedFile(format!(
                    "Directory tree contains a cycle or repeated SID {sid}"
                )));
            }
            visited.insert(sid_usize);

            // Parse entry if not already parsed
            if self.dir_entries[sid_usize].is_none() {
                let offset = sid_usize * DIRENTRY_SIZE;
                let entry =
                    self.parse_directory_entry(&dir_data[offset..offset + DIRENTRY_SIZE], sid)?;

                // Extract child SIDs before moving entry
                let left_sid = entry.sid_left;
                let right_sid = entry.sid_right;
                let child_sid = entry.sid_child;

                self.dir_entries[sid_usize] = Some(entry);

                // Add children to queue (in reverse order for depth-first-like traversal)
                if child_sid != NOSTREAM {
                    queue.push(child_sid);
                }
                if right_sid != NOSTREAM {
                    queue.push(right_sid);
                }
                if left_sid != NOSTREAM {
                    queue.push(left_sid);
                }
            }
        }

        Ok(())
    }

    /// Read a single sector from the file
    fn read_sector(&mut self, sector_id: u32) -> Result<Vec<u8>, OleError> {
        // Sector position in file: (sector_id + 1) * sector_size
        let position = (u64::from(sector_id) + 1)
            .checked_mul(self.sector_size as u64)
            .ok_or_else(|| OleError::CorruptedFile("Sector offset overflow".to_string()))?;
        position
            .checked_add(self.sector_size as u64)
            .ok_or_else(|| OleError::CorruptedFile("Sector end overflow".to_string()))?;
        if position >= self.file_size {
            return Err(OleError::CorruptedFile(format!(
                "Sector {sector_id} is outside the file"
            )));
        }
        self.reader.seek(SeekFrom::Start(position))?;

        // Keep whatever a truncated final sector actually contains; the tail
        // stays zero. See the file-length note in the header parser.
        let present = self.present_sector_bytes(position, self.sector_size);
        let mut buffer = vec![0u8; self.sector_size];
        self.reader.read_exact(&mut buffer[..present])?;
        Ok(buffer)
    }

    /// How many of the `wanted` bytes starting at `position` are present in the
    /// file, for files whose length is not a whole number of sectors.
    #[inline]
    fn present_sector_bytes(&self, position: u64, wanted: usize) -> usize {
        self.file_size.saturating_sub(position).min(wanted as u64) as usize
    }

    /// Read a stream by following the FAT chain with optimized batching
    ///
    /// This implementation batches contiguous sector reads to minimize
    /// system calls (lseek + read), which is a major performance bottleneck.
    fn read_stream_from_fat(&mut self, start_sector: u32) -> Result<Vec<u8>, OleError> {
        if start_sector == ENDOFCHAIN {
            return Ok(Vec::new());
        }

        let sectors = collect_sector_chain(&self.fat, start_sector, "FAT")?;

        // Pre-allocate result buffer
        let data_len = sectors
            .len()
            .checked_mul(self.sector_size)
            .ok_or_else(|| OleError::CorruptedFile("FAT stream size overflow".to_string()))?;
        let mut data = vec![0u8; data_len];

        // Batch read contiguous sectors
        self.read_sectors_batched(&sectors, &mut data)?;

        Ok(data)
    }

    /// Read multiple sectors with batching optimization
    ///
    /// Groups contiguous sectors and reads them in a single I/O operation
    /// to minimize the number of lseek + read system calls.
    fn read_sectors_batched(&mut self, sectors: &[u32], buffer: &mut [u8]) -> Result<(), OleError> {
        if sectors.is_empty() {
            return Ok(());
        }

        let mut i = 0;
        while i < sectors.len() {
            // Find run of contiguous sectors
            let start_sector = sectors[i];
            let mut count = 1;

            while i + count < sectors.len() && sectors[i + count] == sectors[i + count - 1] + 1 {
                count += 1;
            }

            // Read the entire contiguous run in one I/O operation
            let position = (u64::from(start_sector) + 1)
                .checked_mul(self.sector_size as u64)
                .ok_or_else(|| OleError::CorruptedFile("Sector offset overflow".to_string()))?;
            if position >= self.file_size {
                return Err(OleError::CorruptedFile(format!(
                    "Sector {start_sector} is outside the file"
                )));
            }
            let read_size = count * self.sector_size;
            let buffer_offset = i * self.sector_size;

            // `buffer` arrives zero-filled, so a truncated final sector keeps
            // its real bytes and reads as zeroes beyond the end of the file.
            let present = self.present_sector_bytes(position, read_size);
            self.reader.seek(SeekFrom::Start(position))?;
            self.reader
                .read_exact(&mut buffer[buffer_offset..buffer_offset + present])?;

            i += count;
        }

        Ok(())
    }

    /// Read a stream by following the MiniFAT chain with optimized copying
    ///
    /// This implementation pre-allocates and copies data more efficiently
    /// than the naive sector-by-sector approach.
    fn read_stream_from_minifat(
        &mut self,
        start_sector: u32,
        size: u64,
    ) -> Result<Vec<u8>, OleError> {
        // Ensure ministream is loaded
        if self.ministream.is_none() {
            let (start_sector, size) = self
                .root
                .as_ref()
                .map(|root| (root.start_sector, root.size))
                .ok_or_else(|| OleError::CorruptedFile("No root entry".to_string()))?;
            let mut ministream_data = self.read_stream_from_fat(start_sector)?;
            let size = usize::try_from(size)
                .map_err(|_| OleError::CorruptedFile("Mini stream is too large".to_string()))?;
            if ministream_data.len() < size {
                return Err(OleError::CorruptedFile(
                    "Mini stream chain is shorter than its declared size".to_string(),
                ));
            }
            ministream_data.truncate(size);
            self.ministream = Some(ministream_data);
        }

        let ministream = self
            .ministream
            .as_ref()
            .ok_or_else(|| OleError::CorruptedFile("No mini stream".to_string()))?;
        let sectors = collect_sector_chain(&self.minifat, start_sector, "MiniFAT")?;
        let size = usize::try_from(size)
            .map_err(|_| OleError::CorruptedFile("MiniFAT stream is too large".to_string()))?;
        let chain_capacity = sectors
            .len()
            .checked_mul(self.mini_sector_size)
            .ok_or_else(|| OleError::CorruptedFile("MiniFAT stream size overflow".to_string()))?;
        if chain_capacity < size {
            return Err(OleError::CorruptedFile(
                "MiniFAT chain is shorter than the declared stream size".to_string(),
            ));
        }

        // Pre-allocate result buffer with exact size needed
        let mut data = Vec::with_capacity(size);

        // Copy all mini sectors
        for &sector in &sectors {
            let position = (sector as usize) * self.mini_sector_size;
            if position + self.mini_sector_size > ministream.len() {
                return Err(OleError::CorruptedFile(
                    "Mini sector out of bounds".to_string(),
                ));
            }

            data.extend_from_slice(&ministream[position..position + self.mini_sector_size]);
        }

        // Truncate to actual size
        data.truncate(size);
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
    /// * `Result<Vec<&DirectoryEntry>, OleError>` - List of directory entry references (zero-copy)
    pub fn list_directory_entries(&self, path: &[&str]) -> Result<Vec<&DirectoryEntry>, OleError> {
        let mut entries = Vec::new();

        // Get the directory entry
        let dir_entry = if path.is_empty() {
            self.root.as_ref().ok_or(OleError::StreamNotFound)?
        } else {
            self.find_entry(path)?
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

    /// Recursively collect all children from a directory (as references - zero-copy)
    fn collect_directory_children<'a>(&'a self, sid: u32, entries: &mut Vec<&'a DirectoryEntry>) {
        if sid == NOSTREAM || sid as usize >= self.dir_entries.len() {
            return;
        }

        if let Some(entry) = self.dir_entries[sid as usize].as_ref() {
            // Traverse left
            if entry.sid_left != NOSTREAM {
                self.collect_directory_children(entry.sid_left, entries);
            }

            // Add reference instead of clone - zero-copy!
            entries.push(entry);

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
        path: &mut Vec<String>,
        streams: &mut Vec<Vec<String>>,
    ) {
        // Add current entry to path
        let path_len_before = path.len();
        if !entry.name.is_empty() && entry.entry_type != STGTY_ROOT {
            path.push(entry.name.clone()); // Clone needed as we're building the path
        }

        // If this is a stream, add it to the list
        if entry.entry_type == STGTY_STREAM {
            streams.push(path.clone()); // Clone needed to save the path
            path.truncate(path_len_before); // Restore path
            return;
        }

        // If this is a storage, recurse into children
        if entry.entry_type == STGTY_STORAGE || entry.entry_type == STGTY_ROOT {
            // Process children by traversing the red-black tree
            if entry.sid_child != NOSTREAM {
                self.traverse_children(entry.sid_child, path, streams);
            }
        }

        // Restore path to original state
        path.truncate(path_len_before);
    }

    /// Traverse children in red-black tree
    fn traverse_children(&self, sid: u32, path: &mut Vec<String>, streams: &mut Vec<Vec<String>>) {
        if sid == NOSTREAM || sid as usize >= self.dir_entries.len() {
            return;
        }

        if let Some(ref entry) = self.dir_entries[sid as usize] {
            // Traverse left
            if entry.sid_left != NOSTREAM {
                self.traverse_children(entry.sid_left, path, streams);
            }

            // Process current (path is modified in-place and restored)
            self.collect_streams(entry, path, streams);

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
        // Find the entry and extract needed values to avoid borrow conflicts
        let (is_minifat, start_sector, size) = {
            let entry = self.find_entry(path)?;

            // Ensure it's a stream
            if entry.entry_type != STGTY_STREAM {
                return Err(OleError::InvalidFormat("Not a stream".to_string()));
            }

            (entry.is_minifat, entry.start_sector, entry.size)
        };

        // Read the stream based on whether it uses FAT or MiniFAT
        if is_minifat {
            self.read_stream_from_minifat(start_sector, size)
        } else {
            let mut data = self.read_stream_from_fat(start_sector)?;
            let size = usize::try_from(size)
                .map_err(|_| OleError::CorruptedFile("Stream is too large".to_string()))?;
            if data.len() < size {
                return Err(OleError::CorruptedFile(
                    "FAT chain is shorter than the declared stream size".to_string(),
                ));
            }
            data.truncate(size);
            Ok(data)
        }
    }

    /// Find a directory entry by path
    fn find_entry(&self, path: &[&str]) -> Result<&DirectoryEntry, OleError> {
        if path.is_empty() {
            return self.root.as_ref().ok_or(OleError::StreamNotFound);
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

    /// Find a child entry by name in a red-black tree (iterative, optimized)
    ///
    /// OLE directory entries are organized in a red-black tree, though not all
    /// implementations guarantee perfect ordering. This uses an iterative traversal
    /// with zero-allocation string comparison.
    ///
    /// Optimizations:
    /// - Iterative instead of recursive (eliminates function call overhead)
    /// - Zero-allocation case-insensitive comparison using eq_ignore_ascii_case
    /// - Full tree traversal using work queue (handles all tree structures)
    fn find_child_by_name(&self, sid: u32, name: &str) -> Result<&DirectoryEntry, OleError> {
        if sid == NOSTREAM || sid as usize >= self.dir_entries.len() {
            return Err(OleError::StreamNotFound);
        }

        // Use iterative in-order traversal with a work queue (pre-allocated for efficiency)
        // This handles all tree structures correctly, including improperly ordered trees
        let mut queue = smallvec::SmallVec::<[u32; 32]>::new();
        queue.push(sid);

        while let Some(current_sid) = queue.pop() {
            if current_sid == NOSTREAM || current_sid as usize >= self.dir_entries.len() {
                continue;
            }

            let entry = self.dir_entries[current_sid as usize]
                .as_ref()
                .ok_or(OleError::StreamNotFound)?;

            // Fast zero-allocation case-insensitive comparison (ASCII-optimized)
            if entry.name.eq_ignore_ascii_case(name) {
                return Ok(entry);
            }

            // Add children to queue (right first for depth-first-like order)
            if entry.sid_right != NOSTREAM {
                queue.push(entry.sid_right);
            }
            if entry.sid_left != NOSTREAM {
                queue.push(entry.sid_left);
            }
        }

        Err(OleError::StreamNotFound)
    }

    /// Get the root entry name
    pub fn get_root_name(&self) -> Option<&str> {
        self.root.as_ref().map(|r| r.name.as_str())
    }

    /// Root directory entry, including its CLSID.
    pub fn root_entry(&self) -> Option<&DirectoryEntry> {
        self.root.as_ref()
    }

    /// Check if a stream exists
    pub fn exists(&self, path: &[&str]) -> bool {
        self.find_entry(path).is_ok()
    }
}

fn collect_sector_chain(
    allocation_table: &[u32],
    start_sector: u32,
    table_name: &str,
) -> Result<Vec<u32>, OleError> {
    if start_sector == ENDOFCHAIN {
        return Ok(Vec::new());
    }

    let mut sectors = Vec::new();
    let mut visited = FixedBitSet::with_capacity(allocation_table.len());
    let mut sector = start_sector;
    while sector != ENDOFCHAIN {
        let index = usize::try_from(sector).map_err(|_| {
            OleError::CorruptedFile(format!("Invalid sector index in {table_name}"))
        })?;
        if index >= allocation_table.len() {
            return Err(OleError::CorruptedFile(format!(
                "Invalid sector index {sector} in {table_name}"
            )));
        }
        if visited.contains(index) {
            return Err(OleError::CorruptedFile(format!(
                "Cycle detected in {table_name} chain at sector {sector}"
            )));
        }
        visited.insert(index);
        sectors.push(sector);
        let next = allocation_table[index];
        if next != ENDOFCHAIN && next >= MAXREGSECT {
            return Err(OleError::CorruptedFile(format!(
                "Invalid sector marker 0x{next:08X} in {table_name} chain"
            )));
        }
        sector = next;
    }
    Ok(sectors)
}

fn collect_sector_chain_exact(
    allocation_table: &[u32],
    start_sector: u32,
    expected_count: usize,
    table_name: &str,
) -> Result<Vec<u32>, OleError> {
    if expected_count == 0 {
        if start_sector != ENDOFCHAIN {
            return Err(OleError::CorruptedFile(format!(
                "Empty {table_name} chain must start with ENDOFCHAIN"
            )));
        }
        return Ok(Vec::new());
    }
    if start_sector >= MAXREGSECT {
        return Err(OleError::CorruptedFile(format!(
            "Invalid start marker for {table_name} chain"
        )));
    }

    let mut sectors = Vec::with_capacity(expected_count);
    let mut visited = FixedBitSet::with_capacity(allocation_table.len());
    let mut sector = start_sector;
    for index in 0..expected_count {
        let slot = sector as usize;
        if slot >= allocation_table.len() {
            return Err(OleError::CorruptedFile(format!(
                "Invalid sector index {sector} in {table_name}"
            )));
        }
        if visited.contains(slot) {
            return Err(OleError::CorruptedFile(format!(
                "Cycle detected in {table_name} chain at sector {sector}"
            )));
        }
        visited.insert(slot);
        sectors.push(sector);
        let next = allocation_table[slot];
        if index + 1 == expected_count {
            if next != ENDOFCHAIN {
                return Err(OleError::CorruptedFile(format!(
                    "{table_name} chain exceeds its declared length"
                )));
            }
        } else {
            if next == ENDOFCHAIN {
                return Err(OleError::CorruptedFile(format!(
                    "{table_name} chain ends before its declared length"
                )));
            }
            if next >= MAXREGSECT {
                return Err(OleError::CorruptedFile(format!(
                    "Invalid sector marker 0x{next:08X} in {table_name} chain"
                )));
            }
            sector = next;
        }
    }
    Ok(sectors)
}

/// Drop the high stream-size word that version 3 compound files do not use.
///
/// MS-CFB 2.6.1 requires a version 3 stream size to be at most 2 GB, so the
/// most significant 32 bits must be zero. Some older writers never initialized
/// those bits, and the specification recommends that parsers ignore them and
/// treat them as zero rather than reject an otherwise valid file. Version 4
/// files use the full 64-bit value unchanged.
#[inline]
fn mask_v3_stream_size(stream_size: u64, sector_size: usize) -> u64 {
    if sector_size == SECTOR_SIZE_V3 {
        stream_size & u64::from(u32::MAX)
    } else {
        stream_size
    }
}

/// Decode UTF-16LE bytes to String (optimized version)
///
/// Pre-allocates the UTF-16 buffer with exact capacity to avoid reallocations.
fn decode_utf16le(bytes: &[u8]) -> String {
    // Pre-allocate with exact capacity needed
    let capacity = bytes.len() / 2;
    let mut utf16_chars = Vec::with_capacity(capacity);

    for chunk in bytes.chunks_exact(2) {
        let code_unit = U16::<LE>::read_from_bytes(chunk)
            .map(|v| v.get())
            .unwrap_or(0);
        utf16_chars.push(code_unit);
    }

    // Decode UTF-16 to String, replacing invalid sequences
    // Note: trim_end_matches returns a &str, so we only allocate once
    let decoded = String::from_utf16_lossy(&utf16_chars);

    // Check if trimming is needed to avoid unnecessary allocation
    if decoded.ends_with('\0') {
        decoded.trim_end_matches('\0').to_string()
    } else {
        decoded
    }
}

/// Format CLSID as a human-readable string (SIMD-optimized version)
///
/// Uses SIMD-accelerated hex encoding and comparison for optimal performance.
/// CLSID format: XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX (36 characters)
///
/// CLSID uses little-endian byte order for the first three fields (data1-data3),
/// and big-endian for the last field (data4).
///
/// # Performance Optimizations
///
/// - **SIMD hex encoding**: AVX-512, AVX2, SSSE3, or NEON depending on CPU
/// - **SIMD zero check**: Uses movemask instructions for single-cycle check
/// - **Zero heap allocations**: Stack-allocated arrays for byte reversal
/// - **Pre-allocated buffer**: Exact capacity to avoid reallocation
/// - **2-4x faster** than standard formatting on modern CPUs
fn format_clsid(bytes: &[u8]) -> String {
    use litchi_core::simd::cmp::is_all_zero;
    use litchi_core::simd::fmt::hex_encode_to_string;

    if bytes.len() != 16 {
        return String::new();
    }

    // Check if all zeros (empty CLSID) using truly SIMD method
    // Uses movemask instructions (SSE2/AVX2) or horizontal min (NEON)
    // This is a single instruction on x86_64, not a loop!
    if is_all_zero(bytes) {
        return String::new();
    }

    // Pre-allocate with exact capacity: 36 chars for CLSID format
    // XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
    let mut result = String::with_capacity(36);

    // Format with hyphens at correct positions
    // Note: CLSID uses little-endian byte order for the first three fields
    // Use stack-allocated arrays instead of heap Vec for reversed bytes (zero-copy)

    // Data1: 4 bytes (little-endian)
    let mut data1 = [0u8; 4];
    data1.copy_from_slice(&bytes[0..4]);
    data1.reverse();
    hex_encode_to_string(&data1, &mut result, false);
    result.push('-');

    // Data2: 2 bytes (little-endian)
    let mut data2 = [0u8; 2];
    data2.copy_from_slice(&bytes[4..6]);
    data2.reverse();
    hex_encode_to_string(&data2, &mut result, false);
    result.push('-');

    // Data3: 2 bytes (little-endian)
    let mut data3 = [0u8; 2];
    data3.copy_from_slice(&bytes[6..8]);
    data3.reverse();
    hex_encode_to_string(&data3, &mut result, false);
    result.push('-');

    // Data4: remaining bytes (big-endian, no reversal needed)
    hex_encode_to_string(&bytes[8..10], &mut result, false);
    result.push('-');
    hex_encode_to_string(&bytes[10..16], &mut result, false);

    result
}

/// Check if a file/data is an OLE file by checking magic bytes
pub fn is_ole_file(data: &[u8]) -> bool {
    data.len() >= MINIMAL_OLEFILE_SIZE && &data[0..8] == MAGIC
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::writer::OleWriter;
    use std::io::Cursor;

    fn sample_file() -> Vec<u8> {
        let mut writer = OleWriter::new();
        writer.create_stream(&["Data"], b"payload").unwrap();
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        output.into_inner()
    }

    #[test]
    fn rejects_invalid_header_shifts_and_versions_without_panicking() {
        for (offset, value) in [(0x1A, 5u16), (0x1E, 63), (0x20, 63)] {
            let mut data = sample_file();
            data[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
            assert!(matches!(
                OleFile::open(Cursor::new(data)),
                Err(OleError::InvalidFormat(_))
            ));
        }

        let mut data = sample_file();
        data[0x38..0x3C].copy_from_slice(&2048u32.to_le_bytes());
        assert!(matches!(
            OleFile::open(Cursor::new(data)),
            Err(OleError::InvalidFormat(_))
        ));
    }

    #[test]
    fn detects_cycles_in_allocation_chains() {
        assert_eq!(
            collect_sector_chain(&[1, ENDOFCHAIN], 0, "FAT").unwrap(),
            [0, 1]
        );
        assert!(matches!(
            collect_sector_chain(&[1, 0], 0, "FAT"),
            Err(OleError::CorruptedFile(message)) if message.contains("Cycle detected")
        ));
        assert!(matches!(
            collect_sector_chain(&[0], 0, "MiniFAT"),
            Err(OleError::CorruptedFile(message)) if message.contains("Cycle detected")
        ));
    }

    #[test]
    fn rejects_self_referential_difat_chains() {
        let mut bytes = vec![0u8; 114 * 512];
        let mut header = [0u8; 512];
        for sector in 0..109u32 {
            let offset = 0x4C + sector as usize * 4;
            header[offset..offset + 4].copy_from_slice(&sector.to_le_bytes());
        }
        let difat_offset = (109 + 1) * 512;
        bytes[difat_offset..difat_offset + 4].copy_from_slice(&110u32.to_le_bytes());
        bytes[difat_offset + 4..difat_offset + 8].copy_from_slice(&111u32.to_le_bytes());
        for offset in (difat_offset + 8..difat_offset + 508).step_by(4) {
            bytes[offset..offset + 4].copy_from_slice(&FREESECT.to_le_bytes());
        }
        bytes[difat_offset + 508..difat_offset + 512].copy_from_slice(&109u32.to_le_bytes());

        let mut file = OleFile {
            reader: Cursor::new(bytes),
            file_size: (114 * 512) as u64,
            sector_size: 512,
            mini_sector_size: 64,
            mini_stream_cutoff: 4096,
            fat: Vec::new(),
            minifat: Vec::new(),
            first_dir_sector: ENDOFCHAIN,
            root: None,
            dir_entries: Vec::new(),
            ministream: None,
            sector_roles: vec![PhysicalSectorRole::Unclaimed; 113],
        };
        assert!(matches!(
            file.load_fat(&header, 111, 109, 2),
            Err(OleError::CorruptedFile(message)) if message.contains("claimed by both")
        ));
    }

    #[test]
    fn rejects_cyclic_directory_trees() {
        let mut data = sample_file();
        let directory_sector = u32::from_le_bytes(data[0x30..0x34].try_into().unwrap());
        let directory_offset = (directory_sector as usize + 1) * 512;
        data[directory_offset + 76..directory_offset + 80].copy_from_slice(&0u32.to_le_bytes());

        assert!(matches!(
            OleFile::open(Cursor::new(data)),
            Err(OleError::CorruptedFile(message)) if message.contains("repeated SID")
        ));
    }

    #[test]
    fn rejects_sector_reads_past_the_physical_file() {
        let mut file = OleFile {
            reader: Cursor::new(vec![0; 512]),
            file_size: 512,
            sector_size: 512,
            mini_sector_size: 64,
            mini_stream_cutoff: 4096,
            fat: Vec::new(),
            minifat: Vec::new(),
            first_dir_sector: ENDOFCHAIN,
            root: None,
            dir_entries: Vec::new(),
            ministream: None,
            sector_roles: Vec::new(),
        };
        assert!(matches!(
            file.read_sector(0),
            Err(OleError::CorruptedFile(message)) if message.contains("outside the file")
        ));
    }
}
