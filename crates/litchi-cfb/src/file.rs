use super::consts::{
    DIFSECT, DIRENTRY_SIZE, ENDOFCHAIN, FATSECT, FREESECT, HEADER_DIFAT_ENTRIES,
    HEADER_DIFAT_OFFSET, MAGIC, MAXREGSECT, MINIMAL_OLEFILE_SIZE, NOSTREAM, SECTOR_SIZE_V3,
    SECTOR_SIZE_V4, STGTY_EMPTY, STGTY_ROOT, STGTY_STORAGE, STGTY_STREAM,
};
use crate::directory_name::{DirectoryNameData, directory_name_data};
use smallvec::SmallVec;
use std::cmp::Ordering;
use std::io::{self, Read, Seek, SeekFrom};
use zerocopy::{FromBytes, LE, U16, U32, U64};
use zerocopy_derive::FromBytes as DeriveFromBytes;

const BITSET_WORD_BITS: usize = u64::BITS as usize;

/// A fallible, compact bit set for indexes originating in a compound file.
///
/// `fixedbitset::FixedBitSet` is efficient, but its constructor and `insert`
/// operation intentionally panic on allocation and bounds failures.  CFB
/// indexes are untrusted, so the read path uses this small equivalent with
/// fallible allocation and checked insertion instead.
#[derive(Debug, Default)]
struct CheckedBitSet {
    words: Vec<u64>,
    bit_len: usize,
}

impl CheckedBitSet {
    fn try_with_capacity(bit_len: usize, resource: &'static str) -> Result<Self, OleError> {
        let word_count = bit_len.div_ceil(BITSET_WORD_BITS);
        let mut words = Vec::new();
        words
            .try_reserve_exact(word_count)
            .map_err(|source| OleError::allocation(resource, source))?;
        words.resize(word_count, 0);
        Ok(Self { words, bit_len })
    }

    #[inline]
    fn contains(&self, bit: usize) -> bool {
        if bit >= self.bit_len {
            return false;
        }
        let word = bit / BITSET_WORD_BITS;
        let mask = 1u64 << (bit % BITSET_WORD_BITS);
        self.words.get(word).is_some_and(|value| value & mask != 0)
    }

    fn insert(&mut self, bit: usize) -> Result<(), OleError> {
        if bit >= self.bit_len {
            return Err(OleError::CorruptedFile(format!(
                "bit index {bit} exceeds checked bit-set capacity {}",
                self.bit_len
            )));
        }
        let word = bit / BITSET_WORD_BITS;
        let mask = 1u64 << (bit % BITSET_WORD_BITS);
        let value = self.words.get_mut(word).ok_or_else(|| {
            OleError::CorruptedFile(format!("bit index {bit} has no backing word"))
        })?;
        *value |= mask;
        Ok(())
    }
}

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
    /// Validated physical sectors of the root mini stream in logical order.
    root_chain: Vec<u32>,
    /// First sector of directory stream
    first_dir_sector: u32,
    /// Root directory entry
    root: Option<DirectoryEntry>,
    /// All directory entries indexed by SID
    dir_entries: Vec<Option<DirectoryEntry>>,
    /// Validated MS-CFB directory-name comparison data indexed by SID.
    dir_name_data: Vec<Option<DirectoryNameData>>,
    /// Mini stream data (loaded on demand)
    ministream: Option<Vec<u8>>,
    /// Exclusive ownership of every physical sector in the file.
    sector_roles: Vec<PhysicalSectorRole>,
}

/// The validated, cursor-independent portion of an OLE file.
///
/// This stays crate-private so the positional reader can retain exactly the
/// metadata validated by [`OleFile::open`] without exposing CFB physical
/// sector identifiers as public API.
pub(crate) struct ParsedOleIndex {
    pub(crate) file_size: u64,
    pub(crate) sector_size: usize,
    pub(crate) mini_sector_size: usize,
    pub(crate) fat: Vec<u32>,
    pub(crate) minifat: Vec<u32>,
    /// Validated physical sectors of the root mini stream in logical order.
    pub(crate) root_chain: Vec<u32>,
    pub(crate) first_dir_sector: u32,
    pub(crate) root: Option<DirectoryEntry>,
    pub(crate) dir_entries: Vec<Option<DirectoryEntry>>,
    pub(crate) dir_name_data: Vec<Option<DirectoryNameData>>,
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
    /// Whether this stream is in `MiniFAT`
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
    /// The destination was replaced, but its parent directory could not be
    /// synchronized, so durability of the new name is not known.
    Committed {
        source: io::Error,
    },
    /// The source exceeded a finite CFB ingress limit before parsing began.
    LimitExceeded {
        /// Resource whose configured ceiling was crossed.
        resource: &'static str,
        /// Observed source or metadata size.
        observed: u64,
        /// Configured maximum.
        maximum: u64,
    },
    /// A caller supplied a limit outside the CFB hard safety ceiling.
    InvalidLimit {
        /// Resource whose limit was invalid.
        resource: &'static str,
        /// Requested limit.
        value: u64,
        /// Largest supported limit.
        maximum: u64,
    },
    InvalidFormat(String),
    InvalidData(String),
    NotOleFile,
    CorruptedFile(String),
    StreamNotFound,
    /// The positional source changed while a shared CFB view was opening or
    /// reading a stream.
    SourceChanged {
        /// Version captured before the operation started.
        expected: litchi_core::SourceVersion,
        /// Version observed after the operation completed.
        observed: litchi_core::SourceVersion,
    },
}

/// Finite resource limits for the low-level CFB reader.
///
/// The limit is checked immediately after obtaining the reader length and
/// before the parser allocates its physical-sector map or traverses any CFB
/// metadata. The default is deliberately finite and matches the hard ingress
/// ceiling used by [`crate::SharedOleFile`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct OleFileLimits {
    max_input_bytes: u64,
}

impl OleFileLimits {
    /// Largest source accepted by the low-level CFB reader.
    pub const MAX_INPUT_BYTES: u64 = 2 * 1024 * 1024 * 1024;

    /// Creates a finite input ceiling for one CFB source.
    ///
    /// # Errors
    ///
    /// Returns [`OleError::InvalidLimit`] if the ceiling is zero or exceeds
    /// the low-level CFB hard ingress ceiling.
    pub const fn new(max_input_bytes: u64) -> Result<Self, OleError> {
        if max_input_bytes == 0 || max_input_bytes > Self::MAX_INPUT_BYTES {
            return Err(OleError::InvalidLimit {
                resource: "CFB input bytes",
                value: max_input_bytes,
                maximum: Self::MAX_INPUT_BYTES,
            });
        }
        Ok(Self { max_input_bytes })
    }

    /// Maximum source length accepted before parsing begins.
    #[must_use]
    pub const fn max_input_bytes(self) -> u64 {
        self.max_input_bytes
    }
}

impl Default for OleFileLimits {
    fn default() -> Self {
        Self {
            max_input_bytes: Self::MAX_INPUT_BYTES,
        }
    }
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
            OleError::Io(e) => write!(f, "IO error: {e}"),
            OleError::Allocation { resource, source } => {
                write!(f, "could not reserve memory for CFB {resource}: {source}")
            },
            OleError::Committed { source } => write!(
                f,
                "CFB destination was replaced but directory durability could not be confirmed: {source}"
            ),
            OleError::LimitExceeded {
                resource,
                observed,
                maximum,
            } => write!(
                f,
                "CFB {resource} limit exceeded: observed {observed}, maximum {maximum}"
            ),
            OleError::InvalidLimit {
                resource,
                value,
                maximum,
            } => write!(
                f,
                "invalid CFB {resource} limit {value}; maximum is {maximum}"
            ),
            OleError::InvalidFormat(s) => write!(f, "Invalid format: {s}"),
            OleError::InvalidData(s) => write!(f, "Invalid data: {s}"),
            OleError::NotOleFile => write!(f, "Not an OLE file"),
            OleError::CorruptedFile(s) => write!(f, "Corrupted file: {s}"),
            OleError::StreamNotFound => write!(f, "Stream not found"),
            OleError::SourceChanged { expected, observed } => write!(
                f,
                "CFB positional source changed during read (expected {expected:?}, observed {observed:?})"
            ),
        }
    }
}

impl std::error::Error for OleError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Io(source) | Self::Committed { source } => Some(source),
            Self::Allocation { source, .. } => Some(source),
            Self::InvalidFormat(_)
            | Self::InvalidData(_)
            | Self::LimitExceeded { .. }
            | Self::InvalidLimit { .. }
            | Self::NotOleFile
            | Self::CorruptedFile(_)
            | Self::StreamNotFound
            | Self::SourceChanged { .. } => None,
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
            OleError::Allocation { resource, source } => {
                litchi_core::Error::Allocation { resource, source }
            },
            error @ OleError::Committed { .. } => litchi_core::Error::Other(error.to_string()),
            OleError::LimitExceeded {
                resource,
                observed,
                maximum,
            } => litchi_core::Error::ResourceLimit(litchi_core::ResourceLimit {
                resource: litchi_core::Resource::InputBytes,
                observed,
                limit: maximum,
                scope: resource.into(),
            }),
            error @ OleError::InvalidLimit { .. } => {
                litchi_core::Error::InvalidFormat(error.to_string())
            },
            OleError::InvalidFormat(s) | OleError::InvalidData(s) => {
                litchi_core::Error::InvalidFormat(s)
            },
            OleError::NotOleFile => litchi_core::Error::NotOfficeFile,
            OleError::CorruptedFile(s) => litchi_core::Error::CorruptedFile(s),
            OleError::StreamNotFound => {
                litchi_core::Error::ComponentNotFound("Stream not found".to_string())
            },
            error @ OleError::SourceChanged { .. } => litchi_core::Error::Other(error.to_string()),
        }
    }
}

impl<R: Read + Seek> OleFile<R> {
    /// Discard the cursor and validation-only state, retaining the parsed CFB
    /// index for the crate's immutable positional reader.
    pub(crate) fn into_parsed_index(self) -> ParsedOleIndex {
        ParsedOleIndex {
            file_size: self.file_size,
            sector_size: self.sector_size,
            mini_sector_size: self.mini_sector_size,
            fat: self.fat,
            minifat: self.minifat,
            root_chain: self.root_chain,
            first_dir_sector: self.first_dir_sector,
            root: self.root,
            dir_entries: self.dir_entries,
            dir_name_data: self.dir_name_data,
        }
    }

    /// Open and parse an OLE file from a reader
    ///
    /// # Arguments
    /// * `reader` - A reader that implements Read + Seek
    ///
    /// # Returns
    /// * `Result<OleFile<R>, OleError>` - The parsed OLE file or an error
    ///
    /// # Errors
    /// Returns an `OleError` if the reader fails, or if the data is not a
    /// valid OLE compound file (bad magic, truncated header, corrupt FAT).
    pub fn open(reader: R) -> Result<Self, OleError> {
        Self::open_with_limits(reader, OleFileLimits::default())
    }

    /// Open and parse an OLE file under an explicit finite ingress limit.
    ///
    /// The source length is checked immediately after the reader reports it,
    /// before any CFB metadata can cause an allocation or index traversal.
    /// Existing callers should use [`Self::open`], which retains the same
    /// source-compatible entry point with finite default limits.
    pub fn open_with_limits(mut reader: R, limits: OleFileLimits) -> Result<Self, OleError> {
        // Get file size
        let file_size = reader.seek(SeekFrom::End(0))?;
        if file_size > limits.max_input_bytes() {
            return Err(OleError::LimitExceeded {
                resource: "input bytes",
                observed: file_size,
                maximum: limits.max_input_bytes(),
            });
        }
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
        let dll_version = read_u16_le(&header[0x1A..0x1C], "DLL version")?;
        let byte_order = read_u16_le(&header[0x1C..0x1E], "byte order")?;
        let sector_shift = read_u16_le(&header[0x1E..0x20], "sector shift")?;
        let mini_sector_shift = read_u16_le(&header[0x20..0x22], "mini sector shift")?;
        let num_dir_sectors = read_u32_le(&header[0x28..0x2C], "directory sector count")?;
        let first_dir_sector = read_u32_le(&header[0x30..0x34], "directory start sector")?;
        let num_fat_sectors = read_u32_le(&header[0x2C..0x30], "FAT sector count")?;
        let mini_stream_cutoff = read_u32_le(&header[0x38..0x3C], "mini-stream cutoff")?;
        let first_minifat_sector = read_u32_le(&header[0x3C..0x40], "MiniFAT start sector")?;
        let num_minifat_sectors = read_u32_le(&header[0x40..0x44], "MiniFAT sector count")?;
        let first_difat_sector = read_u32_le(&header[0x44..0x48], "DIFAT start sector")?;
        let num_difat_sectors = read_u32_le(&header[0x48..0x4C], "DIFAT sector count")?;

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
            .map_err(|_err| OleError::InvalidFormat("Too many physical sectors".to_string()))?;

        // Validate all count metadata that can otherwise size an allocation or
        // traversal before installing the physical-sector index. The checks
        // remain duplicated in the table loaders as defense in depth for
        // crate-private future callers.
        let physical_sector_count_u64 = u64::try_from(physical_sector_count)
            .map_err(|_err| OleError::InvalidFormat("Too many physical sectors".to_string()))?;
        if u64::from(num_fat_sectors) > physical_sector_count_u64 {
            return Err(OleError::CorruptedFile(
                "Declared FAT sector count exceeds the physical file".to_string(),
            ));
        }
        if u64::from(num_difat_sectors) > physical_sector_count_u64 {
            return Err(OleError::CorruptedFile(
                "Declared DIFAT sector count exceeds the physical file".to_string(),
            ));
        }
        if u64::from(num_minifat_sectors) > physical_sector_count_u64 {
            return Err(OleError::CorruptedFile(
                "Declared MiniFAT sector count exceeds the physical file".to_string(),
            ));
        }
        if dll_version == 4 && u64::from(num_dir_sectors) > physical_sector_count_u64 {
            return Err(OleError::CorruptedFile(
                "Declared directory sector count exceeds the physical file".to_string(),
            ));
        }
        let sector_roles = try_filled_vec(
            physical_sector_count,
            PhysicalSectorRole::Unclaimed,
            "physical sector roles",
        )?;

        let mut ole = OleFile {
            reader,
            file_size,
            sector_size,
            mini_sector_size,
            mini_stream_cutoff,
            fat: Vec::new(),
            minifat: Vec::new(),
            root_chain: Vec::new(),
            first_dir_sector,
            root: None,
            dir_entries: Vec::new(),
            dir_name_data: Vec::new(),
            ministream: None,
            sector_roles,
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
        ole.validate_physical_sector_layout()?;

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
        let physical_sector_count = u64::try_from(self.sector_roles.len())
            .map_err(|_err| OleError::CorruptedFile("Too many physical sectors".to_string()))?;
        if u64::from(num_fat_sectors) > physical_sector_count {
            return Err(OleError::CorruptedFile(
                "Declared FAT sector count exceeds the physical file".to_string(),
            ));
        }
        if u64::from(num_difat_sectors) > physical_sector_count {
            return Err(OleError::CorruptedFile(
                "Declared DIFAT sector count exceeds the physical file".to_string(),
            ));
        }
        let expected_fat_sectors = usize::try_from(num_fat_sectors)
            .map_err(|_err| OleError::CorruptedFile("FAT sector count is too large".to_string()))?;
        let expected_difat_sectors = usize::try_from(num_difat_sectors).map_err(|_err| {
            OleError::CorruptedFile("DIFAT sector count is too large".to_string())
        })?;

        // First 109 FAT sector indexes are in header at offset 0x4C
        let mut fat_sectors = try_vec_with_capacity(expected_fat_sectors, "FAT sector locations")?;
        let mut difat_sectors =
            try_vec_with_capacity(expected_difat_sectors, "DIFAT sector locations")?;
        let header_fat_count = HEADER_DIFAT_ENTRIES.min(expected_fat_sectors);
        for i in 0..header_fat_count {
            let offset = HEADER_DIFAT_OFFSET + i * 4;
            let sector = read_u32_le(&header[offset..offset + 4], "header DIFAT entry")?;
            if sector == FREESECT || sector == ENDOFCHAIN {
                return Err(OleError::CorruptedFile(
                    "FAT sector list ends before its declared count".to_string(),
                ));
            }
            self.claim_sector(sector, PhysicalSectorRole::Fat)?;
            try_push(&mut fat_sectors, sector, "FAT sector locations")?;
        }
        // Entries past the declared FAT sector count are not part of the FAT
        // sector list. MS-CFB 2.2 describes the header DIFAT only as holding
        // "the first 109 FAT sector locations" and never constrains the unused
        // tail, so writers leave zeroes or stale values there. The count field
        // already says where the list ends, and the FAT chain validation below
        // catches a count that disagrees with the file, so the tail is ignored.

        // CFB sector sizes are fixed to 512 or 4096 bytes. Reuse one bounded
        // stack buffer while decoding metadata, rather than allocating a new
        // `Vec` for each DIFAT/FAT sector.
        let mut sector_data = [0u8; SECTOR_SIZE_V4];
        let sector_data = &mut sector_data[..self.sector_size];

        let mut difat_sector = first_difat_sector;
        let entries_per_sector = (self.sector_size / 4) - 1;
        for difat_index in 0..expected_difat_sectors {
            self.claim_sector(difat_sector, PhysicalSectorRole::Difat)?;
            try_push(&mut difat_sectors, difat_sector, "DIFAT sector locations")?;
            self.read_sector_into(difat_sector, sector_data)?;

            for i in 0..entries_per_sector {
                let offset = i * 4;
                let sector = read_u32_le(&sector_data[offset..offset + 4], "DIFAT entry")?;
                if fat_sectors.len() < expected_fat_sectors {
                    if sector >= MAXREGSECT {
                        return Err(OleError::CorruptedFile(
                            "DIFAT sector list ends before the declared FAT count".to_string(),
                        ));
                    }
                    self.claim_sector(sector, PhysicalSectorRole::Fat)?;
                    try_push(&mut fat_sectors, sector, "FAT sector locations")?;
                } else if sector != FREESECT {
                    return Err(OleError::CorruptedFile(
                        "Unused DIFAT entries must be FREESECT".to_string(),
                    ));
                }
            }

            let next_offset = entries_per_sector * 4;
            let next = read_u32_le(
                &sector_data[next_offset..next_offset + 4],
                "DIFAT continuation entry",
            )?;
            if difat_index + 1 == expected_difat_sectors {
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
        let fat_entries_per_sector = self.sector_size / 4;

        // Pre-allocate with exact capacity needed (optimization)
        let fat_entry_count = fat_sectors
            .len()
            .checked_mul(fat_entries_per_sector)
            .ok_or_else(|| OleError::CorruptedFile("FAT entry count overflow".to_string()))?;
        let mut fat = try_vec_with_capacity(fat_entry_count, "FAT entries")?;

        for &sector_id in &fat_sectors {
            self.read_sector_into(sector_id, sector_data)?;

            // Parse sector as array of u32 (little-endian) - use chunks for efficiency
            for chunk in sector_data.chunks_exact(4) {
                let entry = read_u32_le(chunk, "FAT entry")?;
                try_push(&mut fat, entry, "FAT entries")?;
            }
        }

        for sector in fat_sectors {
            let sector_index = usize::try_from(sector).map_err(|_err| {
                OleError::CorruptedFile("FAT sector index does not fit usize".to_string())
            })?;
            if fat.get(sector_index) != Some(&FATSECT) {
                return Err(OleError::CorruptedFile(format!(
                    "FAT sector {sector} is not marked FATSECT"
                )));
            }
        }
        for sector in difat_sectors {
            let sector_index = usize::try_from(sector).map_err(|_err| {
                OleError::CorruptedFile("DIFAT sector index does not fit usize".to_string())
            })?;
            if fat.get(sector_index) != Some(&DIFSECT) {
                return Err(OleError::CorruptedFile(format!(
                    "DIFAT sector {sector} is not marked DIFSECT"
                )));
            }
        }

        self.fat = fat;

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
            usize::try_from(sector_count).map_err(|_err| {
                OleError::CorruptedFile("MiniFAT sector count does not fit usize".to_string())
            })?,
            "MiniFAT",
        )?;
        self.claim_chain(&sectors, PhysicalSectorRole::MiniFat)?;
        let data_len = sectors
            .len()
            .checked_mul(self.sector_size)
            .ok_or_else(|| OleError::CorruptedFile("MiniFAT data size overflow".to_string()))?;
        let entries_count = data_len / 4;
        // Preserve the historical allocation resource label: this reservation
        // has the same byte size and occurs at the same point as the removed
        // aggregate MiniFAT byte-buffer reservation.
        let mut minifat = try_vec_with_capacity(entries_count, "MiniFAT data")?;
        let mut sector_data = [0u8; SECTOR_SIZE_V4];
        let sector_data = &mut sector_data[..self.sector_size];

        // Parse each bounded sector directly into the final table. Keeping only
        // one reusable bounded buffer avoids allocating and copying an
        // aggregate byte buffer the same size as the complete MiniFAT.
        for sector in sectors {
            self.read_sector_into(sector, sector_data)?;
            for chunk in sector_data.chunks_exact(4) {
                let entry = read_u32_le(chunk, "MiniFAT entry")?;
                try_push(&mut minifat, entry, "MiniFAT entries")?;
            }
        }
        self.minifat = minifat;

        Ok(())
    }

    /// Load directory entries with optimized iterative parsing
    fn load_directory(&mut self, declared_sector_count: Option<u32>) -> Result<(), OleError> {
        let sectors = match declared_sector_count {
            Some(count) => collect_sector_chain_exact(
                &self.fat,
                self.first_dir_sector,
                usize::try_from(count).map_err(|_err| {
                    OleError::CorruptedFile("directory sector count does not fit usize".to_string())
                })?,
                "directory",
            )?,
            None => collect_sector_chain(&self.fat, self.first_dir_sector, "directory")?,
        };
        self.claim_chain(&sectors, PhysicalSectorRole::Directory)?;
        let data_len = sectors
            .len()
            .checked_mul(self.sector_size)
            .ok_or_else(|| OleError::CorruptedFile("directory data size overflow".to_string()))?;
        // Initialize the final buffer up front so each sector can be read into
        // its permanent location. This also leaves the unread tail of a short
        // final sector zero-filled, matching `read_sector`.
        let mut dir_data = try_filled_vec(data_len, 0u8, "directory data")?;
        self.read_sectors_batched(&sectors, &mut dir_data)?;

        let validated_entries = Self::validated_directory_entries(&dir_data, self.sector_size)?;

        // Each directory entry is 128 bytes
        let num_entries = dir_data.len() / DIRENTRY_SIZE;
        let mut dir_entries = try_filled_vec(num_entries, None, "directory entries")?;
        if validated_entries.len() != num_entries {
            return Err(OleError::CorruptedFile(
                "validated directory cache has the wrong length".to_string(),
            ));
        }

        // Parse root entry first (always at index 0)
        let mut root_entry = None;
        if num_entries > 0 {
            let root = self.parse_directory_entry(&dir_data[0..DIRENTRY_SIZE], 0)?;
            let root_child_sid = root.sid_child;
            dir_entries[0] = Some(root.clone());

            // Build storage tree using iterative approach (avoids recursion overhead)
            self.build_storage_tree_iterative(root_child_sid, &dir_data, &mut dir_entries)?;
            root_entry = Some(root);
        }

        let mut dir_name_data =
            try_filled_vec(num_entries, None, "directory name comparison data")?;
        for ((name_data, entry), validated_entry) in dir_name_data
            .iter_mut()
            .zip(&dir_entries)
            .zip(validated_entries)
        {
            match (entry, validated_entry) {
                (None, None) => {},
                (Some(entry), Some(validated_entry))
                    if entry.sid == validated_entry.sid
                        // Validation canonicalizes the explicitly supported
                        // classic-Mac root encoding to `Root Entry`, while the
                        // public entry parser retains its historical decoding.
                        && (entry.sid == 0 || entry.name == validated_entry.name) =>
                {
                    *name_data = Some(validated_entry.name_data);
                },
                _ => {
                    return Err(OleError::CorruptedFile(
                        "directory entries and validated name cache disagree".to_string(),
                    ));
                },
            }
        }

        // Install the validated graph only after every directory entry has
        // been traversed successfully. A malformed tree therefore cannot
        // leave a partially populated object behind for a caller that uses
        // this loader in a staged open path.
        self.root = root_entry;
        self.dir_entries = dir_entries;
        self.dir_name_data = dir_name_data;

        Ok(())
    }

    fn claim_sector(&mut self, sector: u32, role: PhysicalSectorRole) -> Result<(), OleError> {
        let slot = self
            .sector_roles
            .get_mut(usize::try_from(sector).map_err(|_err| {
                OleError::CorruptedFile(format!(
                    "{} sector {sector} does not fit usize",
                    role.label()
                ))
            })?)
            .ok_or_else(|| {
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
            .map_err(|_err| OleError::CorruptedFile("Root mini stream is too large".to_string()))?;
        let root_chain = collect_sector_chain_exact(
            &self.fat,
            root_start,
            root_sector_count,
            "root mini stream",
        )?;
        self.claim_chain(&root_chain, PhysicalSectorRole::MiniStream)?;
        self.root_chain = root_chain;

        let mini_sector_capacity =
            usize::try_from(root_size.div_ceil(self.mini_sector_size as u64)).map_err(|_err| {
                OleError::CorruptedFile("Root mini stream is too large".to_string())
            })?;
        let mut claimed_mini_sectors =
            CheckedBitSet::try_with_capacity(mini_sector_capacity, "mini-sector ownership map")?;
        // MiniFAT and FAT normally have different table lengths. Separate
        // scratch instances avoid retaining the larger map in the smaller
        // table's validation path while still eliminating per-stream
        // allocations within each table.
        let mut mini_scratch = SectorChainScratch::default();
        let mut regular_scratch = SectorChainScratch::default();

        for index in 0..self.dir_entries.len() {
            let Some(entry) = self.dir_entries[index].as_ref() else {
                continue;
            };
            if entry.entry_type != STGTY_STREAM {
                continue;
            }
            let (is_minifat, start_sector, size) =
                (entry.is_minifat, entry.start_sector, entry.size);
            if is_minifat {
                let sector_count = usize::try_from(size.div_ceil(self.mini_sector_size as u64))
                    .map_err(|_err| {
                        OleError::CorruptedFile("Mini stream is too large".to_string())
                    })?;
                mini_scratch.collect_exact(
                    &self.minifat,
                    start_sector,
                    sector_count,
                    "mini stream",
                )?;
                for &sector in mini_scratch.sectors() {
                    let sector_index = usize::try_from(sector).map_err(|_err| {
                        OleError::CorruptedFile(
                            "mini stream sector index does not fit usize".to_string(),
                        )
                    })?;
                    if sector_index >= mini_sector_capacity {
                        return Err(OleError::CorruptedFile(
                            "Mini stream references storage outside the root mini stream"
                                .to_string(),
                        ));
                    }
                    if claimed_mini_sectors.contains(sector_index) {
                        return Err(OleError::CorruptedFile(format!(
                            "Mini sector {sector} is claimed by multiple streams"
                        )));
                    }
                    claimed_mini_sectors.insert(sector_index)?;
                }
            } else {
                let sector_count = usize::try_from(size.div_ceil(self.sector_size as u64))
                    .map_err(|_err| {
                        OleError::CorruptedFile("Regular stream is too large".to_string())
                    })?;
                regular_scratch.collect_exact(
                    &self.fat,
                    start_sector,
                    sector_count,
                    "regular stream",
                )?;
                self.claim_chain(regular_scratch.sectors(), PhysicalSectorRole::RegularStream)?;
            }
        }
        Ok(())
    }

    /// Ensure the physical file is reconciled with the FAT after all chains
    /// have been claimed. FAT sectors contain padding entries for their full
    /// sector width, but only entries addressing physical sectors may carry
    /// allocation markers.
    ///
    /// Entries beyond the physical file are tolerated with any marker: MS-CFB
    /// 2.3 requires past-end-of-file entries to be FREESECT, but real-world
    /// producers (including Word and Apache POI) routinely fill the padding
    /// with ENDOFCHAIN instead. Those entries address no physical bytes and
    /// every chain walk is bounds-checked against the physical file, so the
    /// markers are inert; rejecting them would refuse files that Word itself
    /// opens.
    fn validate_physical_sector_layout(&self) -> Result<(), OleError> {
        for (sector, role) in self.sector_roles.iter().enumerate() {
            let entry = self.fat.get(sector).copied().ok_or_else(|| {
                OleError::CorruptedFile(format!(
                    "FAT does not contain an entry for physical sector {sector}"
                ))
            })?;
            if *role == PhysicalSectorRole::Unclaimed && entry != FREESECT {
                return Err(OleError::CorruptedFile(format!(
                    "unclaimed physical sector {sector} has FAT marker 0x{entry:08X}"
                )));
            }
        }
        Ok(())
    }

    #[cfg(test)]
    pub(crate) fn validate_directory(dir_data: &[u8], sector_size: usize) -> Result<(), OleError> {
        Self::validated_directory_entries(dir_data, sector_size).map(drop)
    }

    fn validated_directory_entries(
        dir_data: &[u8],
        sector_size: usize,
    ) -> Result<Vec<Option<ValidatedDirectoryEntry>>, OleError> {
        if dir_data.is_empty() || !dir_data.len().is_multiple_of(DIRENTRY_SIZE) {
            return Err(OleError::CorruptedFile(
                "CFB directory stream must contain complete 128-byte entries".to_string(),
            ));
        }

        let mut entries = try_vec_with_capacity(
            dir_data.len() / DIRENTRY_SIZE,
            "validated directory entries",
        )?;
        for (sid, data) in dir_data.chunks_exact(DIRENTRY_SIZE).enumerate() {
            let entry_sid = u32::try_from(sid).map_err(|_err| {
                OleError::CorruptedFile("CFB directory contains too many entries".to_string())
            })?;
            try_push(
                &mut entries,
                Self::parse_validated_directory_entry(data, entry_sid, sector_size)?,
                "validated directory entries",
            )?;
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

        let mut owned = CheckedBitSet::try_with_capacity(entries.len(), "directory ownership map")?;
        owned.insert(0)?;
        let mut pending_trees = try_vec_with_capacity(1, "directory traversal stack")?;
        try_push(
            &mut pending_trees,
            root.sid_child,
            "directory traversal stack",
        )?;
        while let Some(tree_root) = pending_trees.pop() {
            if tree_root == NOSTREAM {
                continue;
            }
            let mut stack = try_vec_with_capacity(1, "directory traversal stack")?;
            try_push(
                &mut stack,
                (tree_root, None, None, 0usize),
                "directory traversal stack",
            )?;
            while let Some((sid, lower, upper, black_depth)) = stack.pop() {
                if sid == NOSTREAM {
                    // Some widely deployed Office producers wrote unbalanced
                    // color metadata. Bounds, ordering, acyclicity, and unique
                    // ownership are sufficient for safe deterministic traversal.
                    continue;
                }

                let entry = Self::validated_entry(&entries, sid)?;
                let sid_index = usize::try_from(sid).map_err(|_err| {
                    OleError::CorruptedFile("directory SID does not fit usize".to_string())
                })?;
                if owned.contains(sid_index) {
                    return Err(OleError::CorruptedFile(format!(
                        "CFB directory tree contains repeated SID {sid} or cross-storage ownership"
                    )));
                }
                owned.insert(sid_index)?;

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
                    try_push(
                        &mut pending_trees,
                        entry.sid_child,
                        "directory traversal stack",
                    )?;
                }
                let child_black_depth =
                    black_depth + usize::from(entry.node_color == DirectoryNodeColor::Black);
                try_push(
                    &mut stack,
                    (entry.sid_right, Some(sid), upper, child_black_depth),
                    "directory traversal stack",
                )?;
                try_push(
                    &mut stack,
                    (entry.sid_left, lower, Some(sid), child_black_depth),
                    "directory traversal stack",
                )?;
            }
        }

        for entry in entries.iter().flatten() {
            let sid = usize::try_from(entry.sid).map_err(|_err| {
                OleError::CorruptedFile("directory SID does not fit usize".to_string())
            })?;
            if !owned.contains(sid) {
                return Err(OleError::CorruptedFile(format!(
                    "CFB directory SID {} is not owned by a storage",
                    entry.sid
                )));
            }
        }
        Ok(entries)
    }

    fn parse_validated_directory_entry(
        data: &[u8],
        sid: u32,
        sector_size: usize,
    ) -> Result<Option<ValidatedDirectoryEntry>, OleError> {
        let raw = RawDirectoryEntry::read_from_bytes(data).map_err(|_err| {
            OleError::InvalidFormat("Failed to parse directory entry".to_string())
        })?;
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
        let name = String::from_utf16(&name_utf16).map_err(|_err| {
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
            STGTY_STREAM if sid == 0 || sid_child != NOSTREAM => {
                return Err(OleError::CorruptedFile(format!(
                    "invalid CFB stream fields at SID {sid}"
                )));
            },
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
        let sid_index = usize::try_from(sid).map_err(|_err| {
            OleError::CorruptedFile("directory SID does not fit usize".to_string())
        })?;
        entries
            .get(sid_index)
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
        let raw = RawDirectoryEntry::read_from_bytes(data).map_err(|_err| {
            OleError::InvalidFormat("Failed to parse directory entry".to_string())
        })?;

        // Decode name from UTF-16LE
        let name_len = usize::from(raw.name_len.get());
        if !(2..=64).contains(&name_len) || !name_len.is_multiple_of(2) {
            return Err(OleError::CorruptedFile(format!(
                "invalid CFB directory name length {name_len} at SID {sid}"
            )));
        }
        let name = decode_utf16le(&raw.name[..name_len - 2])?;

        // Format CLSID
        let clsid = format_clsid(&raw.clsid);

        // Version 3 sizes only use the low 32 bits; see `mask_v3_stream_size`.
        let size = mask_v3_stream_size(raw.stream_size.get(), self.sector_size);

        // Determine if stream should use MiniFAT
        let is_minifat =
            size < u64::from(self.mini_stream_cutoff) && raw.entry_type == STGTY_STREAM;

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
    /// Uses a compact checked bit set for cache locality and memory efficiency.
    fn build_storage_tree_iterative(
        &mut self,
        root_sid: u32,
        dir_data: &[u8],
        dir_entries: &mut [Option<DirectoryEntry>],
    ) -> Result<(), OleError> {
        if root_sid == NOSTREAM {
            return Ok(());
        }

        let max_entries = dir_data.len() / DIRENTRY_SIZE;

        // Use a work queue for iterative traversal (pre-allocate for common case)
        let mut queue = try_vec_with_capacity(64, "directory build queue")?;
        try_push(&mut queue, root_sid, "directory build queue")?;

        // Track visited SIDs using a compact checked bit set for cache locality
        // Uses ~8x less memory than Vec<bool> (1 bit vs 1 byte per entry)
        let mut visited = CheckedBitSet::try_with_capacity(max_entries, "directory traversal map")?;

        while let Some(sid) = queue.pop() {
            if sid == NOSTREAM {
                continue;
            }

            let sid_usize = usize::try_from(sid).map_err(|_err| {
                OleError::CorruptedFile("directory SID does not fit usize".to_string())
            })?;

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
            visited.insert(sid_usize)?;

            // Parse entry if not already parsed
            if dir_entries[sid_usize].is_none() {
                let offset = sid_usize.checked_mul(DIRENTRY_SIZE).ok_or_else(|| {
                    OleError::CorruptedFile("directory entry offset overflow".to_string())
                })?;
                let end = offset.checked_add(DIRENTRY_SIZE).ok_or_else(|| {
                    OleError::CorruptedFile("directory entry end offset overflow".to_string())
                })?;
                let entry = self.parse_directory_entry(&dir_data[offset..end], sid)?;

                // Extract child SIDs before moving entry
                let left_sid = entry.sid_left;
                let right_sid = entry.sid_right;
                let child_sid = entry.sid_child;

                dir_entries[sid_usize] = Some(entry);

                // Add children to queue (in reverse order for depth-first-like traversal)
                if child_sid != NOSTREAM {
                    try_push(&mut queue, child_sid, "directory build queue")?;
                }
                if right_sid != NOSTREAM {
                    try_push(&mut queue, right_sid, "directory build queue")?;
                }
                if left_sid != NOSTREAM {
                    try_push(&mut queue, left_sid, "directory build queue")?;
                }
            }
        }

        Ok(())
    }

    /// Read a single sector from the file
    #[allow(
        dead_code,
        reason = "kept as the allocating sector-read API and covered by its unit tests"
    )]
    fn read_sector(&mut self, sector_id: u32) -> Result<Vec<u8>, OleError> {
        // Preserve the original error ordering for this allocating API: reject
        // an invalid sector (and perform its seek) before reserving the result.
        let position = self.sector_position(sector_id)?;
        self.reader.seek(SeekFrom::Start(position))?;
        let present = self.present_sector_bytes(position, self.sector_size);
        let mut buffer = try_filled_vec(self.sector_size, 0u8, "sector buffer")?;
        self.reader.read_exact(&mut buffer[..present])?;
        Ok(buffer)
    }

    /// Read one sector into a caller-provided, sector-sized buffer.
    ///
    /// The destination is cleared before the on-disk prefix is read, so short
    /// final sectors retain the CFB reader's zero-filled tail semantics.
    fn read_sector_into(&mut self, sector_id: u32, buffer: &mut [u8]) -> Result<(), OleError> {
        let position = self.sector_position(sector_id)?;
        if buffer.len() != self.sector_size {
            return Err(OleError::CorruptedFile(format!(
                "Sector read destination has length {}, expected {}",
                buffer.len(),
                self.sector_size
            )));
        }
        self.reader.seek(SeekFrom::Start(position))?;

        // Keep whatever a truncated final sector actually contains; the tail
        // stays zero. See the file-length note in the header parser.
        let present = self.present_sector_bytes(position, self.sector_size);
        buffer.fill(0);
        self.reader.read_exact(&mut buffer[..present])?;
        Ok(())
    }

    fn sector_position(&self, sector_id: u32) -> Result<u64, OleError> {
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
        Ok(position)
    }

    /// How many of the `wanted` bytes starting at `position` are present in the
    /// file, for files whose length is not a whole number of sectors.
    #[inline]
    fn present_sector_bytes(&self, position: u64, wanted: usize) -> usize {
        let remaining = self.file_size.saturating_sub(position);
        usize::try_from(remaining.min(wanted as u64)).unwrap_or(wanted)
    }

    /// Read a stream by following the FAT chain with optimized batching
    ///
    /// This implementation batches contiguous sector reads to minimize
    /// system calls (lseek + read), which is a major performance bottleneck.
    fn read_stream_from_fat(
        &mut self,
        start_sector: u32,
        declared_size: u64,
    ) -> Result<Vec<u8>, OleError> {
        let sectors = collect_sector_chain(&self.fat, start_sector, "FAT")?;
        let size = usize::try_from(declared_size)
            .map_err(|_err| OleError::CorruptedFile("FAT stream is too large".to_string()))?;
        let required_sectors = size.div_ceil(self.sector_size);
        if sectors.len() < required_sectors {
            return Err(OleError::CorruptedFile(
                "FAT chain is shorter than the declared stream size".to_string(),
            ));
        }

        // Allocate only the declared stream size. A valid chain may contain
        // unused sectors after the logical end, and a corrupt file must not
        // turn those into an avoidable allocation.
        let mut data = try_filled_vec(size, 0u8, "FAT stream data")?;

        // Batch read contiguous sectors
        self.read_sectors_batched(&sectors[..required_sectors], &mut data)?;

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

            while let Some(next_index) = i.checked_add(count) {
                if next_index >= sectors.len()
                    || sectors[next_index]
                        != sectors[next_index - 1].checked_add(1).ok_or_else(|| {
                            OleError::CorruptedFile("contiguous sector index overflow".to_string())
                        })?
                {
                    break;
                }
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
            let read_size = count
                .checked_mul(self.sector_size)
                .ok_or_else(|| OleError::CorruptedFile("batched read size overflow".to_string()))?;
            let buffer_offset = i.checked_mul(self.sector_size).ok_or_else(|| {
                OleError::CorruptedFile("batched buffer offset overflow".to_string())
            })?;
            let buffer_remaining = buffer.len().checked_sub(buffer_offset).ok_or_else(|| {
                OleError::CorruptedFile("batched read buffer offset overflow".to_string())
            })?;
            let requested = read_size.min(buffer_remaining);

            // The buffer arrives zero-filled, so a truncated final sector keeps
            // its real bytes and reads as zeroes beyond the end of the file.
            if requested > 0 {
                let present = self.present_sector_bytes(position, requested);
                self.reader.seek(SeekFrom::Start(position))?;
                self.reader
                    .read_exact(&mut buffer[buffer_offset..buffer_offset + present])?;
            }

            i += count;
        }

        Ok(())
    }

    /// Read a stream by following the `MiniFAT` chain with optimized copying
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
            let (ministream_start, ministream_size) = self
                .root
                .as_ref()
                .map(|root| (root.start_sector, root.size))
                .ok_or_else(|| OleError::CorruptedFile("No root entry".to_string()))?;
            let mut ministream_data =
                self.read_stream_from_fat(ministream_start, ministream_size)?;
            let ministream_len = usize::try_from(ministream_size)
                .map_err(|_err| OleError::CorruptedFile("Mini stream is too large".to_string()))?;
            if ministream_data.len() < ministream_len {
                return Err(OleError::CorruptedFile(
                    "Mini stream chain is shorter than its declared size".to_string(),
                ));
            }
            ministream_data.truncate(ministream_len);
            self.ministream = Some(ministream_data);
        }

        let ministream = self
            .ministream
            .as_ref()
            .ok_or_else(|| OleError::CorruptedFile("No mini stream".to_string()))?;
        let sectors = collect_sector_chain(&self.minifat, start_sector, "MiniFAT")?;
        let stream_len = usize::try_from(size)
            .map_err(|_err| OleError::CorruptedFile("MiniFAT stream is too large".to_string()))?;
        let chain_capacity = sectors
            .len()
            .checked_mul(self.mini_sector_size)
            .ok_or_else(|| OleError::CorruptedFile("MiniFAT stream size overflow".to_string()))?;
        if chain_capacity < stream_len {
            return Err(OleError::CorruptedFile(
                "MiniFAT chain is shorter than the declared stream size".to_string(),
            ));
        }

        // Pre-allocate result buffer with exact size needed
        let mut data = try_vec_with_capacity(stream_len, "MiniFAT stream data")?;

        // Copy all mini sectors
        for &sector in &sectors {
            let position = usize::try_from(sector)
                .ok()
                .and_then(|sector_id| sector_id.checked_mul(self.mini_sector_size))
                .ok_or_else(|| {
                    OleError::CorruptedFile("Mini sector offset overflow".to_string())
                })?;
            let end = position
                .checked_add(self.mini_sector_size)
                .ok_or_else(|| OleError::CorruptedFile("Mini sector end overflow".to_string()))?;
            if end > ministream.len() {
                return Err(OleError::CorruptedFile(
                    "Mini sector out of bounds".to_string(),
                ));
            }

            let copy_len = self
                .mini_sector_size
                .min(stream_len.saturating_sub(data.len()));
            if copy_len == 0 {
                break;
            }
            data.extend_from_slice(&ministream[position..position + copy_len]);
        }

        // Truncate to actual size
        data.truncate(stream_len);
        Ok(data)
    }

    /// List all streams in the OLE file
    ///
    /// Returns a list of stream paths (as vectors of storage/stream names).
    ///
    /// The directory graph is supplied by untrusted input, so traversal uses
    /// an explicit work stack rather than consuming the call stack.
    pub fn list_streams(&self) -> Vec<Vec<String>> {
        enum Work {
            Tree { sid: u32, path_len: usize },
            Entry { sid: u32, path_len: usize },
            Restore { path_len: usize },
        }

        let mut streams = Vec::new();
        let Some(root) = self.root.as_ref() else {
            return streams;
        };

        let mut pending = Vec::new();
        if root.sid_child != NOSTREAM {
            pending.push(Work::Tree {
                sid: root.sid_child,
                path_len: 0,
            });
        }
        let mut path = Vec::new();

        while let Some(work) = pending.pop() {
            match work {
                Work::Tree { sid, path_len } => {
                    let Some(sid_index) = usize::try_from(sid).ok() else {
                        continue;
                    };
                    if sid == NOSTREAM || sid_index >= self.dir_entries.len() {
                        continue;
                    }
                    let Some(entry) = self.dir_entries.get(sid_index).and_then(Option::as_ref)
                    else {
                        continue;
                    };

                    // Push in reverse in-order so the left subtree is visited
                    // first without recursion.
                    pending.push(Work::Tree {
                        sid: entry.sid_right,
                        path_len,
                    });
                    pending.push(Work::Entry { sid, path_len });
                    pending.push(Work::Tree {
                        sid: entry.sid_left,
                        path_len,
                    });
                },
                Work::Entry { sid, path_len } => {
                    let Some(sid_index) = usize::try_from(sid).ok() else {
                        continue;
                    };
                    let Some(entry) = self.dir_entries.get(sid_index).and_then(Option::as_ref)
                    else {
                        continue;
                    };
                    path.truncate(path_len);
                    if !entry.name.is_empty() && entry.entry_type != STGTY_ROOT {
                        path.push(entry.name.clone());
                    }

                    match entry.entry_type {
                        STGTY_STREAM => {
                            streams.push(path.clone());
                            path.truncate(path_len);
                        },
                        STGTY_STORAGE | STGTY_ROOT if entry.sid_child != NOSTREAM => {
                            let child_path_len = path.len();
                            pending.push(Work::Restore { path_len });
                            pending.push(Work::Tree {
                                sid: entry.sid_child,
                                path_len: child_path_len,
                            });
                        },
                        _ => path.truncate(path_len),
                    }
                },
                Work::Restore { path_len } => path.truncate(path_len),
            }
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
    ///
    /// # Errors
    /// Returns `OleError::StreamNotFound` if no entry exists at `path`, or
    /// `OleError::InvalidFormat` if the entry at `path` is not a directory.
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

    /// Collect all children from a directory (as references - zero-copy).
    ///
    /// The sibling tree is untrusted input, so this uses an explicit stack
    /// instead of recursive calls.
    fn collect_directory_children<'a>(&'a self, sid: u32, entries: &mut Vec<&'a DirectoryEntry>) {
        let mut pending = Vec::new();
        let mut current = sid;

        while current != NOSTREAM || !pending.is_empty() {
            while current != NOSTREAM {
                let Some(current_index) = usize::try_from(current).ok() else {
                    break;
                };
                if current_index >= self.dir_entries.len() {
                    break;
                }
                let Some(entry) = self.dir_entries.get(current_index).and_then(Option::as_ref)
                else {
                    break;
                };
                pending.push(current);
                current = entry.sid_left;
            }

            let Some(current_sid) = pending.pop() else {
                break;
            };
            let Some(current_index) = usize::try_from(current_sid).ok() else {
                continue;
            };
            let Some(entry) = self.dir_entries.get(current_index).and_then(Option::as_ref) else {
                continue;
            };

            // Add reference instead of clone - zero-copy!
            entries.push(entry);
            current = entry.sid_right;
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

    /// Open a stream by path and return its contents
    ///
    /// # Arguments
    /// * `path` - Path to the stream as a slice of strings
    ///
    /// # Returns
    /// * `Result<Vec<u8>, OleError>` - Stream contents or error
    ///
    /// # Errors
    /// Returns `OleError::StreamNotFound` if `path` does not name a stream,
    /// or an `OleError` if reading the underlying sectors fails.
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
            self.read_stream_from_fat(start_sector, size)
        }
    }

    /// Return a stream's declared length without materializing its contents.
    ///
    /// # Errors
    ///
    /// Returns an error if the stream is not found or the path points to a
    /// non-stream entry.
    pub fn stream_len(&self, path: &[&str]) -> Result<u64, OleError> {
        let entry = self.find_entry(path)?;
        if entry.entry_type != STGTY_STREAM {
            return Err(OleError::InvalidFormat("Not a stream".to_string()));
        }
        Ok(entry.size)
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

    /// Find a child entry by name in the validated sibling tree.
    ///
    /// Directory validation establishes strict MS-CFB name ordering, so lookup
    /// can descend through one branch at each node. Limit the number of visited
    /// nodes as a final defense against a malformed in-memory graph.
    fn find_child_by_name(&self, sid: u32, name: &str) -> Result<&DirectoryEntry, OleError> {
        let target_name = directory_name_data(name).map_err(|_error| OleError::StreamNotFound)?;
        let mut current_sid = sid;

        for _ in 0..self.dir_entries.len() {
            if current_sid == NOSTREAM {
                return Err(OleError::StreamNotFound);
            }
            let Some(current_index) = usize::try_from(current_sid).ok() else {
                return Err(OleError::StreamNotFound);
            };
            let entry = self
                .dir_entries
                .get(current_index)
                .and_then(Option::as_ref)
                .ok_or(OleError::StreamNotFound)?;
            let entry_name = self
                .dir_name_data
                .get(current_index)
                .and_then(Option::as_ref)
                .ok_or(OleError::StreamNotFound)?;

            current_sid = match target_name.compare(entry_name) {
                Ordering::Less => entry.sid_left,
                Ordering::Equal => return Ok(entry),
                Ordering::Greater => entry.sid_right,
            };
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

fn try_vec_with_capacity<T>(capacity: usize, resource: &'static str) -> Result<Vec<T>, OleError> {
    let mut values = Vec::new();
    values
        .try_reserve_exact(capacity)
        .map_err(|source| OleError::allocation(resource, source))?;
    Ok(values)
}

fn try_push<T>(values: &mut Vec<T>, value: T, resource: &'static str) -> Result<(), OleError> {
    values
        .try_reserve(1)
        .map_err(|source| OleError::allocation(resource, source))?;
    values.push(value);
    Ok(())
}

fn try_filled_vec<T: Clone>(
    len: usize,
    value: T,
    resource: &'static str,
) -> Result<Vec<T>, OleError> {
    let mut values = try_vec_with_capacity(len, resource)?;
    values.resize(len, value);
    Ok(values)
}

fn read_u16_le(bytes: &[u8], description: &str) -> Result<u16, OleError> {
    let value: [u8; 2] = bytes
        .try_into()
        .map_err(|_err| OleError::InvalidFormat(format!("{description} is truncated")))?;
    Ok(u16::from_le_bytes(value))
}

fn read_u32_le(bytes: &[u8], description: &str) -> Result<u32, OleError> {
    let value: [u8; 4] = bytes
        .try_into()
        .map_err(|_err| OleError::InvalidFormat(format!("{description} is truncated")))?;
    Ok(u32::from_le_bytes(value))
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
    let mut visited = CheckedBitSet::try_with_capacity(allocation_table.len(), "sector-chain map")?;
    let mut sector = start_sector;
    while sector != ENDOFCHAIN {
        let index = usize::try_from(sector).map_err(|_err| {
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
        visited.insert(index)?;
        try_push(&mut sectors, sector, "sector-chain entries")?;
        let next = *allocation_table.get(index).ok_or_else(|| {
            OleError::CorruptedFile(format!("Invalid sector index {sector} in {table_name}"))
        })?;
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
    if expected_count > allocation_table.len() {
        return Err(OleError::CorruptedFile(format!(
            "{table_name} chain length exceeds its allocation table"
        )));
    }

    let mut sectors = try_vec_with_capacity(expected_count, "sector-chain entries")?;
    let mut visited = CheckedBitSet::try_with_capacity(allocation_table.len(), "sector-chain map")?;
    let mut sector = start_sector;
    for index in 0..expected_count {
        let slot = usize::try_from(sector).map_err(|_err| {
            OleError::CorruptedFile(format!("Invalid sector index {sector} in {table_name}"))
        })?;
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
        visited.insert(slot)?;
        try_push(&mut sectors, sector, "sector-chain entries")?;
        let next = *allocation_table.get(slot).ok_or_else(|| {
            OleError::CorruptedFile(format!("Invalid sector index {sector} in {table_name}"))
        })?;
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

/// Reusable fallible buffers for exact chain validation.
///
/// `collect_sector_chain_exact` remains the general-purpose helper for
/// callers that need an owned result. Stream-allocation validation walks many
/// chains against the same FAT or MiniFAT, so retaining these buffers removes
/// the two transient allocations that would otherwise occur for every
/// stream. The buffers are deliberately private to that validation path.
#[derive(Debug, Default)]
struct SectorChainScratch {
    sectors: Vec<u32>,
    visited: CheckedBitSet,
}

impl SectorChainScratch {
    fn sectors(&self) -> &[u32] {
        &self.sectors
    }

    fn reset_visited(&mut self) {
        self.visited.bit_len = 0;
    }

    fn reset(&mut self) {
        self.sectors.clear();
        self.reset_visited();
    }

    fn prepare_visited(&mut self, bit_len: usize) -> Result<(), OleError> {
        let word_count = bit_len.div_ceil(BITSET_WORD_BITS);
        let visited = &mut self.visited;
        if visited.words.len() < word_count {
            visited
                .words
                .try_reserve_exact(word_count - visited.words.len())
                .map_err(|source| OleError::allocation("sector-chain map", source))?;
            // The fallible reserve above makes this resize infallible.
            visited.words.resize(word_count, 0);
        }
        visited.bit_len = bit_len;
        visited.words.fill(0);
        Ok(())
    }

    fn collect_exact(
        &mut self,
        allocation_table: &[u32],
        start_sector: u32,
        expected_count: usize,
        table_name: &str,
    ) -> Result<(), OleError> {
        self.reset();
        let result = (|| {
            if expected_count == 0 {
                if start_sector != ENDOFCHAIN {
                    return Err(OleError::CorruptedFile(format!(
                        "Empty {table_name} chain must start with ENDOFCHAIN"
                    )));
                }
                return Ok(());
            }
            if start_sector >= MAXREGSECT {
                return Err(OleError::CorruptedFile(format!(
                    "Invalid start marker for {table_name} chain"
                )));
            }
            if expected_count > allocation_table.len() {
                return Err(OleError::CorruptedFile(format!(
                    "{table_name} chain length exceeds its allocation table"
                )));
            }

            // Preserve the original allocation order: the chain vector is
            // reserved before the visited map, with the same resource labels
            // as the owned-result helper.
            if self.sectors.capacity() < expected_count {
                self.sectors
                    .try_reserve_exact(expected_count)
                    .map_err(|source| OleError::allocation("sector-chain entries", source))?;
            }
            self.prepare_visited(allocation_table.len())?;
            let mut sector = start_sector;
            for index in 0..expected_count {
                let slot = usize::try_from(sector).map_err(|_err| {
                    OleError::CorruptedFile(format!(
                        "Invalid sector index {sector} in {table_name}"
                    ))
                })?;
                if slot >= allocation_table.len() {
                    return Err(OleError::CorruptedFile(format!(
                        "Invalid sector index {sector} in {table_name}"
                    )));
                }
                if self.visited.contains(slot) {
                    return Err(OleError::CorruptedFile(format!(
                        "Cycle detected in {table_name} chain at sector {sector}"
                    )));
                }
                self.visited.insert(slot)?;
                try_push(&mut self.sectors, sector, "sector-chain entries")?;
                let next = *allocation_table.get(slot).ok_or_else(|| {
                    OleError::CorruptedFile(format!(
                        "Invalid sector index {sector} in {table_name}"
                    ))
                })?;
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
            Ok(())
        })();
        if result.is_err() {
            self.reset();
        }
        result
    }
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
fn decode_utf16le(bytes: &[u8]) -> Result<String, OleError> {
    if !bytes.len().is_multiple_of(2) {
        return Err(OleError::InvalidFormat(
            "UTF-16LE directory name has an odd byte length".to_string(),
        ));
    }
    let unit_count = bytes.len() / 2;
    let capacity = unit_count.checked_mul(4).ok_or_else(|| {
        OleError::InvalidFormat("UTF-16LE directory name is too large".to_string())
    })?;
    let mut decoded = String::new();
    decoded
        .try_reserve(capacity)
        .map_err(|source| OleError::allocation("decoded CFB directory name", source))?;
    for value in std::char::decode_utf16(
        bytes
            .chunks_exact(2)
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]])),
    ) {
        decoded.push(value.unwrap_or('\u{FFFD}'));
    }
    while decoded.ends_with('\0') {
        decoded.pop();
    }
    Ok(decoded)
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
#[must_use]
pub fn is_ole_file(data: &[u8]) -> bool {
    data.len() >= MINIMAL_OLEFILE_SIZE && &data[0..8] == MAGIC
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        clippy::expect_used,
        reason = "test assertions panic on failure by design"
    )]
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

    fn file_with_streams<I, S>(names: I) -> OleFile<Cursor<Vec<u8>>>
    where
        I: IntoIterator<Item = S>,
        S: AsRef<str>,
    {
        let mut writer = OleWriter::new();
        for name in names {
            writer.create_stream(&[name.as_ref()], b"").unwrap();
        }
        let mut output = Cursor::new(Vec::new());
        writer.write_to(&mut output).unwrap();
        OleFile::open(Cursor::new(output.into_inner())).unwrap()
    }

    #[test]
    fn finds_exact_names_in_a_wide_sibling_tree() {
        let names: Vec<_> = (0..257).map(|index| format!("Entry {index:03}")).collect();
        let file = file_with_streams(&names);

        for name in &names {
            assert_eq!(file.stream_len(&[name]).unwrap(), 0);
        }
    }

    #[test]
    fn parsed_name_cache_is_sid_aligned_with_directory_entries() {
        let file = file_with_streams(["Alpha", "Bravo", "Charlie"]);

        assert_eq!(file.dir_name_data.len(), file.dir_entries.len());
        for (sid, (entry, cached_name)) in
            file.dir_entries.iter().zip(&file.dir_name_data).enumerate()
        {
            assert_eq!(entry.is_some(), cached_name.is_some());
            if let (Some(entry), Some(cached_name)) = (entry, cached_name) {
                assert_eq!(entry.sid as usize, sid);
                assert_eq!(*cached_name, directory_name_data(&entry.name).unwrap());
            }
        }
    }

    #[test]
    fn lookup_refuses_a_missing_validated_name_cache_entry() {
        let mut file = file_with_streams(["Alpha"]);
        let child_sid = file.root.as_ref().unwrap().sid_child as usize;
        file.dir_name_data[child_sid] = None;

        assert!(matches!(
            file.stream_len(&["Alpha"]),
            Err(OleError::StreamNotFound)
        ));
    }

    #[test]
    fn returns_not_found_for_missing_and_invalid_caller_names() {
        let file = file_with_streams(["Alpha", "Bravo", "Charlie"]);
        let too_long = "x".repeat(crate::directory_name::MAX_DIRECTORY_NAME_CODE_UNITS + 1);

        for name in ["Delta", "", "bad/name", "nul\0name", &too_long] {
            assert!(matches!(
                file.stream_len(&[name]),
                Err(OleError::StreamNotFound)
            ));
        }
    }

    #[test]
    fn lookup_uses_ascii_case_equivalence() {
        let file = file_with_streams(["Quarterly Report"]);

        assert_eq!(file.stream_len(&["qUaRtErLy rEpOrT"]).unwrap(), 0);
    }

    #[test]
    fn lookup_uses_cfb_simple_uppercase_comparison() {
        assert!(!"ſtream".eq_ignore_ascii_case("Stream"));
        let file = file_with_streams(["ſtream", "élan", "straße"]);

        assert_eq!(file.stream_len(&["Stream"]).unwrap(), 0);
        assert_eq!(file.stream_len(&["ÉLAN"]).unwrap(), 0);
        assert!(matches!(
            file.stream_len(&["STRASSE"]),
            Err(OleError::StreamNotFound)
        ));
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
    fn directory_name_decoder_checks_utf16_extents_without_panicking() {
        assert_eq!(decode_utf16le(b"A\0\0\0").unwrap(), "A");
        assert_eq!(decode_utf16le(&[0x00, 0xD8]).unwrap(), "\u{FFFD}");
        assert!(matches!(
            decode_utf16le(&[0x41]),
            Err(OleError::InvalidFormat(message)) if message.contains("odd byte length")
        ));
    }

    #[test]
    fn rejects_chain_lengths_before_reserving_chain_storage() {
        let result = collect_sector_chain_exact(&[ENDOFCHAIN], 0, 2, "FAT");
        assert!(matches!(
            result,
            Err(OleError::CorruptedFile(message))
                if message.contains("exceeds its allocation table")
        ));
    }

    #[test]
    fn reusable_chain_scratch_reuses_buffers_for_different_lengths() {
        let mut scratch = SectorChainScratch::default();
        let long_table = [
            1, 2, 3, ENDOFCHAIN, ENDOFCHAIN, ENDOFCHAIN, ENDOFCHAIN, ENDOFCHAIN,
        ];
        scratch
            .collect_exact(&long_table, 0, 4, "regular stream")
            .unwrap();
        let sectors_ptr = scratch.sectors.as_ptr();
        let sectors_capacity = scratch.sectors.capacity();
        let visited_ptr = scratch.visited.words.as_ptr();
        let visited_words = scratch.visited.words.len();
        assert_eq!(scratch.sectors(), &[0, 1, 2, 3]);

        scratch.reset();
        scratch
            .collect_exact(&[ENDOFCHAIN], 0, 1, "mini stream")
            .unwrap();

        assert_eq!(scratch.sectors(), &[0]);
        assert_eq!(scratch.sectors.as_ptr(), sectors_ptr);
        assert_eq!(scratch.sectors.capacity(), sectors_capacity);
        assert_eq!(scratch.visited.words.as_ptr(), visited_ptr);
        assert_eq!(scratch.visited.words.len(), visited_words);
        assert_eq!(scratch.visited.bit_len, 1);
    }

    #[test]
    fn reusable_chain_scratch_reserves_growth_before_walking() {
        let mut scratch = SectorChainScratch::default();
        scratch
            .collect_exact(&[ENDOFCHAIN], 0, 1, "mini stream")
            .unwrap();

        let early = scratch.collect_exact(&[ENDOFCHAIN; 8], 0, 8, "mini stream");
        assert!(matches!(
            early,
            Err(OleError::CorruptedFile(message))
                if message == "mini stream chain ends before its declared length"
        ));
        assert!(scratch.sectors.capacity() >= 8);
        assert!(scratch.sectors.is_empty());
    }

    #[test]
    fn reusable_chain_scratch_resets_after_success_and_empty_chain() {
        let mut scratch = SectorChainScratch::default();
        scratch
            .collect_exact(&[1, ENDOFCHAIN], 0, 2, "regular stream")
            .unwrap();
        assert!(!scratch.sectors.is_empty());

        scratch.reset();
        assert!(scratch.sectors.is_empty());
        assert_eq!(scratch.visited.bit_len, 0);

        scratch
            .collect_exact(&[], ENDOFCHAIN, 0, "empty stream")
            .unwrap();
        assert!(scratch.sectors.is_empty());
        assert_eq!(scratch.visited.bit_len, 0);
    }

    #[test]
    fn reusable_chain_scratch_resets_after_cycle_and_reuses_after_error() {
        let mut scratch = SectorChainScratch::default();
        let result = scratch.collect_exact(&[1, 0, ENDOFCHAIN], 0, 3, "regular stream");
        assert!(matches!(
            result,
            Err(OleError::CorruptedFile(message))
                if message == "Cycle detected in regular stream chain at sector 0"
        ));
        assert!(scratch.sectors.is_empty());
        assert_eq!(scratch.visited.bit_len, 0);

        scratch
            .collect_exact(&[ENDOFCHAIN], 0, 1, "regular stream")
            .unwrap();
        assert_eq!(scratch.sectors(), &[0]);
    }

    #[test]
    fn reusable_chain_scratch_preserves_early_and_late_end_errors() {
        let mut scratch = SectorChainScratch::default();

        let early = scratch.collect_exact(&[ENDOFCHAIN, ENDOFCHAIN], 0, 2, "mini stream");
        assert!(matches!(
            early,
            Err(OleError::CorruptedFile(message))
                if message == "mini stream chain ends before its declared length"
        ));
        assert!(scratch.sectors.is_empty());

        let late = scratch.collect_exact(&[1, ENDOFCHAIN], 0, 1, "mini stream");
        assert!(matches!(
            late,
            Err(OleError::CorruptedFile(message))
                if message == "mini stream chain exceeds its declared length"
        ));
        assert!(scratch.sectors.is_empty());
        assert_eq!(scratch.visited.bit_len, 0);
    }

    #[test]
    fn malformed_large_declarations_do_not_unwind() {
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            let mut data = sample_file();
            data[0x48..0x4C].copy_from_slice(&u32::MAX.to_le_bytes());
            OleFile::open(Cursor::new(data))
        }));

        assert!(result.is_ok(), "malformed CFB input must not panic");
        assert!(result.as_ref().is_ok_and(Result::is_err));
    }

    #[test]
    fn fat_stream_reads_only_the_declared_logical_size() {
        let mut bytes = vec![0u8; 3 * 512];
        bytes[512..515].copy_from_slice(b"abc");
        let mut file = OleFile {
            reader: Cursor::new(bytes),
            file_size: (3 * 512) as u64,
            sector_size: 512,
            mini_sector_size: 64,
            mini_stream_cutoff: 4096,
            fat: vec![1, ENDOFCHAIN],
            minifat: Vec::new(),
            root_chain: Vec::new(),
            first_dir_sector: ENDOFCHAIN,
            root: None,
            dir_entries: Vec::new(),
            dir_name_data: Vec::new(),
            ministream: None,
            sector_roles: vec![PhysicalSectorRole::Unclaimed; 2],
        };

        assert_eq!(file.read_stream_from_fat(0, 3).unwrap(), b"abc");
    }

    #[test]
    fn batched_sector_reads_place_noncontiguous_sectors_in_order() {
        let mut bytes = vec![0u8; 4 * SECTOR_SIZE_V3];
        bytes[SECTOR_SIZE_V3..SECTOR_SIZE_V3 + 3].copy_from_slice(b"one");
        bytes[3 * SECTOR_SIZE_V3..3 * SECTOR_SIZE_V3 + 3].copy_from_slice(b"two");
        let mut file = OleFile {
            reader: Cursor::new(bytes),
            file_size: (4 * SECTOR_SIZE_V3) as u64,
            sector_size: SECTOR_SIZE_V3,
            mini_sector_size: 64,
            mini_stream_cutoff: 4096,
            fat: Vec::new(),
            minifat: Vec::new(),
            root_chain: Vec::new(),
            first_dir_sector: ENDOFCHAIN,
            root: None,
            dir_entries: Vec::new(),
            dir_name_data: Vec::new(),
            ministream: None,
            sector_roles: vec![PhysicalSectorRole::Unclaimed; 3],
        };
        let mut data = vec![0xFF; 2 * SECTOR_SIZE_V3];

        file.read_sectors_batched(&[0, 2], &mut data).unwrap();

        assert_eq!(&data[..3], b"one");
        assert_eq!(&data[SECTOR_SIZE_V3..SECTOR_SIZE_V3 + 3], b"two");
        assert!(data[3..SECTOR_SIZE_V3].iter().all(|&byte| byte == 0));
        assert!(data[SECTOR_SIZE_V3 + 3..].iter().all(|&byte| byte == 0));
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
            root_chain: Vec::new(),
            first_dir_sector: ENDOFCHAIN,
            root: None,
            dir_entries: Vec::new(),
            dir_name_data: Vec::new(),
            ministream: None,
            sector_roles: vec![PhysicalSectorRole::Unclaimed; 113],
        };
        assert!(matches!(
            file.load_fat(&header, 111, 109, 2),
            Err(OleError::CorruptedFile(message)) if message.contains("claimed by both")
        ));
    }

    #[test]
    fn rejects_difat_counts_beyond_the_physical_file_before_reserving() {
        let mut data = sample_file();
        data[0x44..0x48].copy_from_slice(&0u32.to_le_bytes());
        data[0x48..0x4C].copy_from_slice(&u32::MAX.to_le_bytes());

        assert!(matches!(
            OleFile::open(Cursor::new(data)),
            Err(OleError::CorruptedFile(message))
                if message.contains("DIFAT sector count exceeds the physical file")
        ));
    }

    #[test]
    fn rejects_minifat_buffer_size_overflow_before_reading_a_sector() {
        let mut file = OleFile {
            reader: Cursor::new(Vec::new()),
            file_size: 0,
            sector_size: usize::MAX,
            mini_sector_size: 64,
            mini_stream_cutoff: 4096,
            fat: vec![1, ENDOFCHAIN],
            minifat: Vec::new(),
            root_chain: Vec::new(),
            first_dir_sector: ENDOFCHAIN,
            root: None,
            dir_entries: Vec::new(),
            dir_name_data: Vec::new(),
            ministream: None,
            sector_roles: vec![PhysicalSectorRole::Unclaimed; 2],
        };

        assert!(matches!(
            file.load_minifat(0, 2),
            Err(OleError::CorruptedFile(message))
                if message.contains("MiniFAT data size overflow")
        ));
    }

    #[test]
    fn reports_minifat_allocation_failure_without_panicking() {
        let mut file = OleFile {
            reader: Cursor::new(Vec::new()),
            file_size: 0,
            sector_size: usize::MAX,
            mini_sector_size: 64,
            mini_stream_cutoff: 4096,
            fat: vec![ENDOFCHAIN],
            minifat: Vec::new(),
            root_chain: Vec::new(),
            first_dir_sector: ENDOFCHAIN,
            root: None,
            dir_entries: Vec::new(),
            dir_name_data: Vec::new(),
            ministream: None,
            sector_roles: vec![PhysicalSectorRole::Unclaimed],
        };

        assert!(matches!(
            file.load_minifat(0, 1),
            Err(OleError::Allocation {
                resource: "MiniFAT data",
                ..
            })
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
            root_chain: Vec::new(),
            first_dir_sector: ENDOFCHAIN,
            root: None,
            dir_entries: Vec::new(),
            dir_name_data: Vec::new(),
            ministream: None,
            sector_roles: Vec::new(),
        };
        assert!(matches!(
            file.read_sector(0),
            Err(OleError::CorruptedFile(message)) if message.contains("outside the file")
        ));
    }

    #[test]
    fn read_sector_into_zero_fills_a_truncated_final_sector() {
        let mut bytes = vec![0u8; SECTOR_SIZE_V3 + 3];
        bytes[SECTOR_SIZE_V3..].copy_from_slice(b"CFB");
        let mut file = OleFile {
            reader: Cursor::new(bytes),
            file_size: (SECTOR_SIZE_V3 + 3) as u64,
            sector_size: SECTOR_SIZE_V3,
            mini_sector_size: 64,
            mini_stream_cutoff: 4096,
            fat: Vec::new(),
            minifat: Vec::new(),
            root_chain: Vec::new(),
            first_dir_sector: ENDOFCHAIN,
            root: None,
            dir_entries: Vec::new(),
            dir_name_data: Vec::new(),
            ministream: None,
            sector_roles: Vec::new(),
        };
        let mut destination = [0xFF; SECTOR_SIZE_V3];

        file.read_sector_into(0, &mut destination).unwrap();

        assert_eq!(&destination[..3], b"CFB");
        assert!(destination[3..].iter().all(|&byte| byte == 0));
    }

    #[test]
    fn tolerates_nonfree_fat_padding_beyond_the_physical_file() {
        // MS-CFB 2.3 requires past-end-of-file FAT entries to be FREESECT,
        // but real producers (Word, Apache POI) fill the padding with
        // ENDOFCHAIN; the markers address no physical bytes and are inert.
        let mut data = sample_file();
        let physical_sector_count = data.len() / 512 - 1;
        let fat_sector = u32::from_le_bytes(data[0x4c..0x50].try_into().unwrap());
        let fat_offset = (usize::try_from(fat_sector).unwrap() + 1) * 512;
        let padding_offset = fat_offset + physical_sector_count * 4;
        assert_eq!(
            &data[padding_offset..padding_offset + 4],
            &FREESECT.to_le_bytes()
        );
        data[padding_offset..padding_offset + 4].copy_from_slice(&ENDOFCHAIN.to_le_bytes());

        assert!(OleFile::open(Cursor::new(data)).is_ok());
    }

    #[test]
    fn lists_deep_directory_trees_without_call_stack_growth() {
        const DEPTH: usize = 16_384;

        let mut dir_entries = vec![None; DEPTH + 1];
        let mut dir_name_data = vec![None; DEPTH + 1];
        for index in (0..DEPTH).rev() {
            let sid = u32::try_from(index + 1).unwrap();
            let right = if index + 1 < DEPTH {
                u32::try_from(index + 2).unwrap()
            } else {
                NOSTREAM
            };
            let name = format!("Stream {index:05}");
            dir_name_data[sid as usize] = Some(directory_name_data(&name).unwrap());
            dir_entries[sid as usize] = Some(DirectoryEntry {
                sid,
                name,
                entry_type: STGTY_STREAM,
                sid_left: NOSTREAM,
                sid_right: right,
                sid_child: NOSTREAM,
                clsid: String::new(),
                start_sector: ENDOFCHAIN,
                size: 0,
                is_minifat: false,
                children: Vec::new(),
            });
        }

        let file = OleFile {
            reader: Cursor::new(Vec::new()),
            file_size: 0,
            sector_size: 512,
            mini_sector_size: 64,
            mini_stream_cutoff: 4096,
            fat: Vec::new(),
            minifat: Vec::new(),
            root_chain: Vec::new(),
            first_dir_sector: ENDOFCHAIN,
            root: Some(DirectoryEntry {
                sid: 0,
                name: "Root Entry".to_string(),
                entry_type: STGTY_ROOT,
                sid_left: NOSTREAM,
                sid_right: NOSTREAM,
                sid_child: 1,
                clsid: String::new(),
                start_sector: ENDOFCHAIN,
                size: 0,
                is_minifat: false,
                children: Vec::new(),
            }),
            dir_entries,
            dir_name_data,
            ministream: None,
            sector_roles: Vec::new(),
        };

        let streams = file.list_streams();
        assert_eq!(streams.len(), DEPTH);
        assert_eq!(streams[0], vec!["Stream 00000"]);
        let last_name = format!("Stream {:05}", DEPTH - 1);
        assert_eq!(streams[DEPTH - 1], vec![last_name]);

        let entries = file.list_directory_entries(&[]).unwrap();
        assert_eq!(entries.len(), DEPTH);
        assert_eq!(entries.first().map(|entry| entry.sid), Some(1));
        assert_eq!(
            entries.last().map(|entry| entry.sid),
            Some(u32::try_from(DEPTH).unwrap())
        );
    }
}
