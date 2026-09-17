//! Source-anchored sector layout for [`OleWriter`](super::OleWriter).
//!
//! [`OleWriter`](super::OleWriter) is a from-scratch builder: it holds the
//! complete logical model of a compound file and, by default before this
//! module existed, allocated every sector afresh on each serialization. That
//! is the smallest possible output, but it moves every stream's bytes even
//! when an edit touched one of them.
//!
//! This module adds the other half of the policy pair required by decision 10
//! of `docs/performance/0652-owner-decisions-for-the-third-wave.md`: when the
//! caller adopts the artifact it is republishing with
//! [`OleWriter::adopt_source_layout`](super::OleWriter::adopt_source_layout),
//! the writer reuses that artifact's sector assignment. Streams whose byte
//! length still fits their existing allocation keep their sectors; a stream
//! that outgrows its allocation keeps the prefix it has and takes further
//! sectors from the free list before appending; sectors a shrinking stream
//! gives up are reclaimed for the same save.
//!
//! The reuse path never invents a layout of its own. Every gate it cannot
//! satisfy — a changed directory shape, a changed geometry, a source that
//! declares DIFAT sectors, or an output that would need them — makes it
//! decline, and the writer falls back to the from-scratch serialization with
//! the reason recorded in [`SectorLayoutReport`].

use super::super::consts::{
    DIRENTRY_SIZE, ENDOFCHAIN, FATSECT, FREESECT, HEADER_DIFAT_ENTRIES, HEADER_DIFAT_OFFSET,
    MAXREGSECT, STGTY_ROOT, STGTY_STORAGE, STGTY_STREAM,
};
use super::super::file::{OleError, OleFile, OleFileLimits};
use std::collections::{BTreeMap, BTreeSet};
use std::io::{self, Cursor, Read, Seek, SeekFrom, Write};

/// Header offset of the Number of Directory Sectors field (MS-CFB 2.2).
const NUM_DIR_SECTORS_OFFSET: usize = 0x28;
/// Header offset of the Number of FAT Sectors field (MS-CFB 2.2).
const NUM_FAT_SECTORS_OFFSET: usize = 0x2C;
/// Header offset of the First Directory Sector Location field (MS-CFB 2.2).
const FIRST_DIR_SECTOR_OFFSET: usize = 0x30;
/// Header offset of the Mini Stream Cutoff Size field (MS-CFB 2.2).
const MINI_STREAM_CUTOFF_OFFSET: usize = 0x38;
/// Header offset of the First Mini FAT Sector Location field (MS-CFB 2.2).
const FIRST_MINIFAT_SECTOR_OFFSET: usize = 0x3C;
/// Header offset of the Number of Mini FAT Sectors field (MS-CFB 2.2).
const NUM_MINIFAT_SECTORS_OFFSET: usize = 0x40;
/// Header offset of the First DIFAT Sector Location field (MS-CFB 2.2).
const FIRST_DIFAT_SECTOR_OFFSET: usize = 0x44;
/// Header offset of the Number of DIFAT Sectors field (MS-CFB 2.2).
const NUM_DIFAT_SECTORS_OFFSET: usize = 0x48;
/// Directory entry offset of the Starting Sector Location field (MS-CFB 2.6.1).
const ENTRY_START_SECTOR_OFFSET: usize = 0x74;
/// Directory entry offset of the Stream Size field (MS-CFB 2.6.1).
const ENTRY_STREAM_SIZE_OFFSET: usize = 0x78;
/// Bound on the FAT-sector fixed point, matching the from-scratch planner.
const LAYOUT_FIXED_POINT_ROUNDS: usize = 32;

/// Where [`OleWriter`](super::OleWriter) places the sectors it serializes.
///
/// The policy only has an effect once an artifact has been adopted with
/// [`OleWriter::adopt_source_layout`](super::OleWriter::adopt_source_layout).
/// Without an adopted source there is no layout to reuse and both variants
/// produce the same from-scratch serialization.
///
/// Both variants publish the same logical document: identical stream bytes,
/// identical hierarchy, identical names and identical class identifiers. They
/// differ only in where the bytes land inside the container, and in how much
/// of the source's directory metadata survives — see the crate documentation
/// of [`SectorLayoutReport`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
#[non_exhaustive]
pub enum SectorLayoutPolicy {
    /// Reuse the adopted source's sector assignment; append only what does not
    /// fit. This is the default.
    ///
    /// An unchanged stream keeps its sectors and its bytes are written through
    /// at the offsets they already occupy rather than re-laid out. A changed
    /// stream that still needs the same number of sectors is written into
    /// them. A stream that outgrows its allocation keeps the sectors it has
    /// and takes the rest from sectors this save released or the source left
    /// unallocated, appending beyond the source's last sector only when the
    /// free list runs out. Nothing is moved for the sake of compaction.
    #[default]
    Reuse,
    /// Re-serialize the whole container from scratch, allocating every sector
    /// afresh in the writer's deterministic order.
    ///
    /// This produces the smallest output — there are no unallocated sectors in
    /// it at all — at the cost of relocating every stream and re-deriving the
    /// whole directory, including its red-black colours and its timestamps.
    Rewrite,
}

/// Why a serialization did not reuse the adopted source's sector layout.
///
/// A declined reuse is never an error: the writer serializes from scratch
/// instead, producing exactly the bytes [`SectorLayoutPolicy::Rewrite`] would
/// have produced. The reason is reported so that callers and tests can tell a
/// deliberate policy choice from a gate that fired.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum SectorLayoutFallback {
    /// The caller selected [`SectorLayoutPolicy::Rewrite`].
    PolicySelected,
    /// No source artifact was adopted, so there is no layout to reuse.
    NoAdoptedSource,
    /// The writer's sector geometry differs from the adopted source's.
    GeometryChanged,
    /// The set of streams or storages differs from the adopted source's.
    ///
    /// Creating, deleting, moving or renaming an entry changes the directory
    /// tree's shape and its storage identifier assignment, which only the
    /// from-scratch serializer performs.
    DirectoryShapeChanged,
    /// A class identifier the caller set differs from the adopted source's.
    ClassIdChanged,
    /// The adopted source declares DIFAT sectors.
    ///
    /// Reuse rebuilds the FAT, so it would have to rebuild the DIFAT chain as
    /// well. No fixture in this repository declares one, so the branch would
    /// ship untested; the gate can be lifted when one exists.
    SourceDeclaresDifat,
    /// The output would need more than the 109 FAT sector locations the header
    /// itself can hold, and therefore a DIFAT chain.
    OutputNeedsDifat,
    /// The adopted source's length is not a whole number of sectors.
    SourceNotSectorAligned,
    /// The adopted source's mini stream is not a whole number of mini sectors.
    SourceMiniStreamUnaligned,
    /// The planner could not produce a layout that satisfies every invariant.
    ///
    /// The from-scratch serialization is always available, so this declines
    /// rather than refusing.
    PlanRejected,
}

/// What one serialization did with the adopted source's sectors.
///
/// The counts are in sectors of the artifact's own sector size. They describe
/// the emitted container, not the logical document: two outputs with different
/// reports still expose byte-identical streams.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct SectorLayoutReport {
    reused: bool,
    fallback: Option<SectorLayoutFallback>,
    output_sectors: u32,
    kept_sectors: u32,
    rewritten_sectors: u32,
    appended_sectors: u32,
    reclaimed_sectors: u32,
    free_sectors: u32,
}

impl SectorLayoutReport {
    /// Whether the adopted source's sector layout was reused.
    #[must_use]
    pub const fn reused_source_layout(self) -> bool {
        self.reused
    }

    /// Why the source layout was not reused, when it was not.
    #[must_use]
    pub const fn fallback(self) -> Option<SectorLayoutFallback> {
        self.fallback
    }

    /// Sectors in the emitted artifact, excluding the header sector.
    #[must_use]
    pub const fn output_sectors(self) -> u32 {
        self.output_sectors
    }

    /// Sectors whose owner and logical position are the source's own.
    ///
    /// Their bytes are written through from the caller's payloads to the same
    /// offsets they already occupy.
    #[must_use]
    pub const fn kept_sectors(self) -> u32 {
        self.kept_sectors
    }

    /// Sectors that received different content than the source holds there:
    /// the rebuilt FAT, MiniFAT, directory and mini stream, plus every sector
    /// an allocation moved.
    #[must_use]
    pub const fn rewritten_sectors(self) -> u32 {
        self.rewritten_sectors
    }

    /// Sectors appended beyond the adopted source's last sector.
    #[must_use]
    pub const fn appended_sectors(self) -> u32 {
        self.appended_sectors
    }

    /// Sectors taken from the free list rather than appended.
    #[must_use]
    pub const fn reclaimed_sectors(self) -> u32 {
        self.reclaimed_sectors
    }

    /// Sectors the emitted artifact leaves unallocated.
    #[must_use]
    pub const fn free_sectors(self) -> u32 {
        self.free_sectors
    }

    pub(super) const fn declined(fallback: SectorLayoutFallback) -> Self {
        Self {
            reused: false,
            fallback: Some(fallback),
            output_sectors: 0,
            kept_sectors: 0,
            rewritten_sectors: 0,
            appended_sectors: 0,
            reclaimed_sectors: 0,
            free_sectors: 0,
        }
    }

    pub(super) const fn with_output_sectors(mut self, sectors: u32) -> Self {
        self.output_sectors = sectors;
        self
    }
}

/// What claims one physical sector of the adopted source.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum SectorOwner {
    Free,
    Fat,
    Difat,
    Directory,
    MiniFat,
    MiniStream,
    Stream(u32),
}

/// One directory entry of the adopted source, as the parser validated it.
#[derive(Debug, Clone)]
struct SourceEntry {
    sid: u32,
    entry_type: u8,
    start_sector: u32,
    size: u64,
    is_mini: bool,
    class_id: [u8; 16],
}

/// The bounded metadata of the artifact an [`OleWriter`](super::OleWriter)
/// republishes.
///
/// Only bounded source state is retained: the header sector, the FAT, the
/// MiniFAT, the directory image, the packed mini-stream image and one record
/// per directory entry. The packed mini-stream image includes source bytes
/// because mini-sector allocations are addressed inside it; regular stream
/// payloads and a second full copy of the source artifact are not retained by
/// this object. Its retained bytes are bounded by the source's sector count.
#[derive(Debug, Clone)]
pub(super) struct SourceLayout {
    sector_size: usize,
    mini_sector_size: usize,
    mini_stream_cutoff: u32,
    sector_count: u32,
    header_sector: Vec<u8>,
    fat: Vec<u32>,
    minifat: Vec<u32>,
    directory_image: Vec<u8>,
    ministream_image: Vec<u8>,
    ministream_size: u64,
    dir_chain: Vec<u32>,
    minifat_chain: Vec<u32>,
    root_chain: Vec<u32>,
    fat_sectors: Vec<u32>,
    entries: Vec<SourceEntry>,
    stream_paths: BTreeMap<Vec<String>, u32>,
    storage_paths: BTreeMap<Vec<String>, u32>,
    root_class_id: [u8; 16],
}

fn invalid(message: &str) -> OleError {
    OleError::InvalidData(message.to_string())
}

fn reserved_vec<T>(len: usize, resource: &'static str) -> Result<Vec<T>, OleError> {
    let mut value = Vec::new();
    value
        .try_reserve_exact(len)
        .map_err(|source| OleError::allocation(resource, source))?;
    Ok(value)
}

fn zeroed_vec(len: usize, resource: &'static str) -> Result<Vec<u8>, OleError> {
    let mut value = reserved_vec::<u8>(len, resource)?;
    value.resize(len, 0);
    Ok(value)
}

fn filled_vec(len: usize, value: u32, resource: &'static str) -> Result<Vec<u32>, OleError> {
    let mut out = reserved_vec::<u32>(len, resource)?;
    out.resize(len, value);
    Ok(out)
}

const BITSET_WORD_BITS: usize = u64::BITS as usize;

/// Fallible membership tracking for source-sized indexes.  Source directory
/// and MiniFAT counts are untrusted, so a `vec![false; count]` would turn a
/// hostile but representable count into an infallible allocation.
#[derive(Debug)]
struct LayoutBitSet {
    words: Vec<u64>,
    bit_len: usize,
}

impl LayoutBitSet {
    fn try_with_capacity(bit_len: usize, resource: &'static str) -> Result<Self, OleError> {
        let word_count = bit_len.div_ceil(BITSET_WORD_BITS);
        let mut words = Vec::new();
        words
            .try_reserve_exact(word_count)
            .map_err(|source| OleError::allocation(resource, source))?;
        words.resize(word_count, 0);
        Ok(Self { words, bit_len })
    }

    fn contains(&self, bit: usize) -> bool {
        if bit >= self.bit_len {
            return false;
        }
        self.words
            .get(bit / BITSET_WORD_BITS)
            .is_some_and(|word| word & (1u64 << (bit % BITSET_WORD_BITS)) != 0)
    }

    fn insert(&mut self, bit: usize) -> Result<(), OleError> {
        if bit >= self.bit_len {
            return Err(invalid("CFB source layout bit index is outside its bound"));
        }
        let word = self
            .words
            .get_mut(bit / BITSET_WORD_BITS)
            .ok_or_else(|| invalid("CFB source layout bit index has no word"))?;
        *word |= 1u64 << (bit % BITSET_WORD_BITS);
        Ok(())
    }
}

fn usize_from_u32(value: u32, resource: &'static str) -> Result<usize, OleError> {
    usize::try_from(value).map_err(|_err| invalid(resource))
}

fn read_u32(bytes: &[u8], offset: usize) -> Result<u32, OleError> {
    let slice = bytes
        .get(offset..offset + 4)
        .ok_or_else(|| invalid("CFB source layout field is outside the header"))?;
    let mut raw = [0u8; 4];
    raw.copy_from_slice(slice);
    Ok(u32::from_le_bytes(raw))
}

fn write_u32(bytes: &mut [u8], offset: usize, value: u32) -> Result<(), OleError> {
    let slice = bytes
        .get_mut(offset..offset + 4)
        .ok_or_else(|| invalid("CFB source layout field is outside its record"))?;
    slice.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

fn write_u64(bytes: &mut [u8], offset: usize, value: u64) -> Result<(), OleError> {
    let slice = bytes
        .get_mut(offset..offset + 8)
        .ok_or_else(|| invalid("CFB source layout field is outside its record"))?;
    slice.copy_from_slice(&value.to_le_bytes());
    Ok(())
}

/// Number of units of `unit` bytes needed to hold `len` bytes, as `u32`.
fn units_for(len: u64, unit: usize) -> Result<u32, OleError> {
    let unit = u64::try_from(unit).map_err(|_err| invalid("CFB allocation unit exceeds u64"))?;
    if unit == 0 {
        return Err(invalid("CFB allocation unit is zero"));
    }
    let count = len.div_ceil(unit);
    let count = u32::try_from(count).map_err(|_err| invalid("CFB chain length exceeds u32"))?;
    if count > MAXREGSECT {
        return Err(invalid("CFB chain length exceeds MAXREGSECT"));
    }
    Ok(count)
}

impl SourceLayout {
    /// Parses `source` through the ordinary validating CFB parser and retains
    /// the bounded metadata a reused layout needs.
    ///
    /// # Errors
    ///
    /// Returns the parser's own errors for a malformed artifact, and
    /// [`OleError::InvalidData`] when a declared chain or geometry cannot be
    /// represented.
    pub(super) fn parse(source: &[u8]) -> Result<Self, OleError> {
        let ole = OleFile::open(Cursor::new(source))?;
        let file_size = ole.file_size();
        let index = ole.into_parsed_index();
        let sector_size = index.sector_size;
        if !matches!(sector_size, 512 | 4096) {
            return Err(invalid("CFB source layout sector size must be 512 or 4096"));
        }
        let sector_size_u64 = sector_size as u64;
        let body = file_size
            .checked_sub(sector_size_u64)
            .ok_or_else(|| invalid("CFB source layout is shorter than its header sector"))?;
        if body % sector_size_u64 != 0 {
            return Err(invalid(
                "CFB source layout is not a whole number of sectors",
            ));
        }
        let sector_count = u32::try_from(body / sector_size_u64)
            .map_err(|_err| invalid("CFB source layout sector count exceeds u32"))?;

        let header_sector = source
            .get(..sector_size)
            .ok_or_else(|| invalid("CFB source layout header sector is truncated"))?
            .to_vec();
        let mini_stream_cutoff = read_u32(&header_sector, MINI_STREAM_CUTOFF_OFFSET)?;
        let minifat_start = read_u32(&header_sector, FIRST_MINIFAT_SECTOR_OFFSET)?;

        // The owner map is local: it proves the source's sectors partition
        // cleanly before any of the derived chains are trusted.
        let sector_count_usize = usize_from_u32(
            sector_count,
            "CFB source sector count does not fit this platform",
        )?;
        let mut owners =
            reserved_vec::<SectorOwner>(sector_count_usize, "CFB source sector owners")?;
        owners.resize(sector_count_usize, SectorOwner::Free);
        let fat = index.fat;
        let minifat = index.minifat;

        // FAT and DIFAT sectors announce themselves inside the FAT itself.
        let mut fat_sectors = Vec::new();
        for sector in 0..sector_count {
            let sector_index =
                usize_from_u32(sector, "CFB source sector index does not fit this platform")?;
            let entry = fat.get(sector_index).copied().unwrap_or(FREESECT);
            match entry {
                FATSECT => {
                    fat_sectors
                        .try_reserve(1)
                        .map_err(|source| OleError::allocation("CFB source FAT sectors", source))?;
                    fat_sectors.push(sector);
                    claim(&mut owners, sector, SectorOwner::Fat)?;
                },
                super::super::consts::DIFSECT => claim(&mut owners, sector, SectorOwner::Difat)?,
                _ => {},
            }
        }

        let dir_chain = walk_chain(&fat, index.first_dir_sector, sector_count, "directory")?;
        for sector in &dir_chain {
            claim(&mut owners, *sector, SectorOwner::Directory)?;
        }
        let minifat_chain = walk_chain(&fat, minifat_start, sector_count, "MiniFAT")?;
        for sector in &minifat_chain {
            claim(&mut owners, *sector, SectorOwner::MiniFat)?;
        }
        let root_chain = index.root_chain.clone();
        for sector in &root_chain {
            claim(&mut owners, *sector, SectorOwner::MiniStream)?;
        }

        let directory_image = gather_sectors(source, &dir_chain, sector_size, "directory image")?;
        let ministream_image =
            gather_sectors(source, &root_chain, sector_size, "mini stream image")?;

        let root = index
            .root
            .as_ref()
            .ok_or_else(|| invalid("CFB source layout has no root entry"))?;
        let ministream_size = root.size;
        let root_class_id = entry_class_id(&directory_image, 0)?;

        let mut entries =
            reserved_vec::<SourceEntry>(index.dir_entries.len(), "CFB source entries")?;
        let mut stream_paths = BTreeMap::new();
        let mut storage_paths = BTreeMap::new();
        for (sid, entry) in index.dir_entries.iter().enumerate() {
            let sid_u32 =
                u32::try_from(sid).map_err(|_err| invalid("CFB source SID exceeds u32"))?;
            match entry {
                Some(entry) => entries.push(SourceEntry {
                    sid: sid_u32,
                    entry_type: entry.entry_type,
                    start_sector: entry.start_sector,
                    size: entry.size,
                    is_mini: entry.is_minifat,
                    class_id: entry_class_id(&directory_image, sid_u32)?,
                }),
                None => entries.push(SourceEntry {
                    sid: sid_u32,
                    entry_type: super::super::consts::STGTY_EMPTY,
                    start_sector: ENDOFCHAIN,
                    size: 0,
                    is_mini: false,
                    class_id: [0u8; 16],
                }),
            }
        }
        collect_paths(&index.dir_entries, &mut stream_paths, &mut storage_paths)?;

        for entry in &entries {
            if entry.entry_type != STGTY_STREAM || entry.size == 0 {
                continue;
            }
            if entry.is_mini {
                continue;
            }
            let chain = walk_chain(&fat, entry.start_sector, sector_count, "stream")?;
            for sector in &chain {
                claim(&mut owners, *sector, SectorOwner::Stream(entry.sid))?;
            }
        }

        Ok(Self {
            sector_size,
            mini_sector_size: index.mini_sector_size,
            mini_stream_cutoff,
            sector_count,
            header_sector,
            fat,
            minifat,
            directory_image,
            ministream_image,
            ministream_size,
            dir_chain,
            minifat_chain,
            root_chain,
            fat_sectors,
            entries,
            stream_paths,
            storage_paths,
            root_class_id,
        })
    }

    pub(super) fn difat_declared(&self) -> Result<bool, OleError> {
        Ok(
            read_u32(&self.header_sector, NUM_DIFAT_SECTORS_OFFSET)? != 0
                || read_u32(&self.header_sector, FIRST_DIFAT_SECTOR_OFFSET)? != ENDOFCHAIN,
        )
    }

    pub(super) const fn root_class_id(&self) -> [u8; 16] {
        self.root_class_id
    }

    pub(super) fn storage_paths(&self) -> impl Iterator<Item = &Vec<String>> {
        self.storage_paths.keys()
    }

    pub(super) fn storage_class_id(&self, path: &[String]) -> Option<[u8; 16]> {
        let sid = self.storage_paths.get(path).copied()?;
        let index = usize::try_from(sid).ok()?;
        self.entries.get(index).map(|entry| entry.class_id)
    }
}

fn claim(owners: &mut [SectorOwner], sector: u32, owner: SectorOwner) -> Result<(), OleError> {
    let index = usize_from_u32(sector, "CFB source layout sector does not fit usize")?;
    let slot = owners
        .get_mut(index)
        .ok_or_else(|| invalid("CFB source layout sector is outside the artifact"))?;
    if *slot != SectorOwner::Free {
        return Err(invalid("CFB source layout sector is claimed twice"));
    }
    *slot = owner;
    Ok(())
}

fn walk_chain(
    fat: &[u32],
    start: u32,
    sector_count: u32,
    resource: &'static str,
) -> Result<Vec<u32>, OleError> {
    let mut chain = Vec::new();
    if start == ENDOFCHAIN || start == FREESECT {
        return Ok(chain);
    }
    let mut current = start;
    for _ in 0..=u64::from(sector_count) {
        if current == ENDOFCHAIN {
            return Ok(chain);
        }
        if current >= sector_count {
            return Err(invalid(&format!(
                "CFB source {resource} chain leaves the artifact"
            )));
        }
        chain
            .try_reserve(1)
            .map_err(|source| OleError::allocation(resource, source))?;
        chain.push(current);
        let index = usize_from_u32(current, "CFB source layout chain index does not fit usize")?;
        current = fat
            .get(index)
            .copied()
            .ok_or_else(|| invalid(&format!("CFB source {resource} chain leaves the FAT")))?;
    }
    Err(invalid(&format!(
        "CFB source {resource} chain does not end"
    )))
}

fn gather_sectors(
    source: &[u8],
    chain: &[u32],
    sector_size: usize,
    resource: &'static str,
) -> Result<Vec<u8>, OleError> {
    let bytes = chain
        .len()
        .checked_mul(sector_size)
        .ok_or_else(|| invalid("CFB source layout image size overflows usize"))?;
    let mut image = reserved_vec::<u8>(bytes, resource)?;
    for sector in chain {
        let start = usize_from_u32(*sector, "CFB source layout sector does not fit usize")?
            .checked_add(1)
            .and_then(|index| index.checked_mul(sector_size))
            .ok_or_else(|| invalid("CFB source layout sector offset overflows usize"))?;
        let end = start
            .checked_add(sector_size)
            .ok_or_else(|| invalid("CFB source layout sector end overflows usize"))?;
        let slice = source
            .get(start..end)
            .ok_or_else(|| invalid("CFB source layout sector is outside the artifact"))?;
        image.extend_from_slice(slice);
    }
    Ok(image)
}

fn entry_class_id(directory_image: &[u8], sid: u32) -> Result<[u8; 16], OleError> {
    let start = usize_from_u32(sid, "CFB directory SID does not fit usize")?
        .checked_mul(DIRENTRY_SIZE)
        .and_then(|offset| offset.checked_add(0x50))
        .ok_or_else(|| invalid("CFB directory entry offset overflows usize"))?;
    let end = start
        .checked_add(16)
        .ok_or_else(|| invalid("CFB directory entry class-ID offset overflows usize"))?;
    let slice = directory_image
        .get(start..end)
        .ok_or_else(|| invalid("CFB directory entry is outside the directory image"))?;
    let mut class_id = [0u8; 16];
    class_id.copy_from_slice(slice);
    Ok(class_id)
}

fn collect_paths(
    entries: &[Option<super::super::file::DirectoryEntry>],
    streams: &mut BTreeMap<Vec<String>, u32>,
    storages: &mut BTreeMap<Vec<String>, u32>,
) -> Result<(), OleError> {
    let root = entries
        .first()
        .and_then(Option::as_ref)
        .ok_or_else(|| invalid("CFB source layout has no root entry"))?;
    let mut visited = LayoutBitSet::try_with_capacity(entries.len(), "CFB source directory walk")?;
    visited.insert(0)?;
    let mut pending: Vec<(Vec<String>, u32)> = Vec::new();
    if root.sid_child != super::super::consts::NOSTREAM {
        pending.push((Vec::new(), root.sid_child));
    }
    while let Some((prefix, sid)) = pending.pop() {
        let sid_index = usize_from_u32(sid, "CFB source layout child SID does not fit usize")?;
        if sid_index >= entries.len() {
            return Err(invalid(
                "CFB source layout child SID is outside the directory",
            ));
        }
        if visited.contains(sid_index) {
            return Err(invalid(
                "CFB source layout directory tree revisits an entry",
            ));
        }
        visited.insert(sid_index)?;
        let entry = entries
            .get(sid_index)
            .and_then(Option::as_ref)
            .ok_or_else(|| invalid("CFB source layout child SID is unallocated"))?;
        let mut path = prefix.clone();
        path.push(entry.name.clone());
        match entry.entry_type {
            STGTY_STREAM => {
                if streams.insert(path.clone(), sid).is_some() {
                    return Err(invalid(
                        "CFB source layout declares a duplicate stream path",
                    ));
                }
            },
            STGTY_STORAGE | STGTY_ROOT => {
                if storages.insert(path.clone(), sid).is_some() {
                    return Err(invalid(
                        "CFB source layout declares a duplicate storage path",
                    ));
                }
                if entry.sid_child != super::super::consts::NOSTREAM {
                    pending.push((path.clone(), entry.sid_child));
                }
            },
            _ => {
                return Err(invalid(
                    "CFB source layout declares an unsupported entry type",
                ));
            },
        }
        if entry.sid_left != super::super::consts::NOSTREAM {
            pending.push((prefix.clone(), entry.sid_left));
        }
        if entry.sid_right != super::super::consts::NOSTREAM {
            pending.push((prefix, entry.sid_right));
        }
    }
    Ok(())
}

/// One stream of the writer's logical model, borrowed for planning.
#[derive(Debug, Clone, Copy)]
pub(super) struct StreamInput<'a> {
    pub(super) path: &'a [String],
    pub(super) bytes: &'a [u8],
}

/// What a planned output sector holds.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum PlannedSector {
    Free,
    Fat(u32),
    MiniFat(u32),
    Directory(u32),
    MiniStream(u32),
    Stream { index: u32, chunk: u32 },
}

/// A complete source-anchored serialization, ready to emit.
#[derive(Debug)]
pub(super) struct ReusePlan {
    sector_size: usize,
    header_sector: Vec<u8>,
    sectors: Vec<PlannedSector>,
    fat_image: Vec<u8>,
    minifat_image: Vec<u8>,
    directory_image: Vec<u8>,
    ministream_image: Vec<u8>,
    report: SectorLayoutReport,
}

/// Cursor over the bytes a reuse plan would publish. It lets the ordinary
/// reader validate a planned artifact without first allocating a second full
/// output buffer.
struct PlanCursor<'a> {
    plan: &'a ReusePlan,
    streams: &'a [StreamInput<'a>],
    position: u64,
}

impl<'a> PlanCursor<'a> {
    fn new(plan: &'a ReusePlan, streams: &'a [StreamInput<'a>]) -> Self {
        Self {
            plan,
            streams,
            position: 0,
        }
    }

    fn io_error(error: OleError) -> io::Error {
        io::Error::new(io::ErrorKind::InvalidData, error.to_string())
    }
}

impl Read for PlanCursor<'_> {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        let read = self
            .plan
            .read_at(self.streams, self.position, output)
            .map_err(Self::io_error)?;
        self.position = self
            .position
            .checked_add(u64::try_from(read).map_err(|_error| {
                io::Error::new(io::ErrorKind::InvalidData, "planned read exceeds u64")
            })?)
            .ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidData, "planned cursor overflows")
            })?;
        Ok(read)
    }
}

impl Seek for PlanCursor<'_> {
    fn seek(&mut self, from: SeekFrom) -> io::Result<u64> {
        let length = self.plan.output_len().map_err(Self::io_error)?;
        let (base, offset) = match from {
            SeekFrom::Start(position) => (i128::from(position), 0),
            SeekFrom::Current(offset) => (i128::from(self.position), i128::from(offset)),
            SeekFrom::End(offset) => (i128::from(length), i128::from(offset)),
        };
        let position = base
            .checked_add(offset)
            .filter(|position| *position >= 0)
            .and_then(|position| u64::try_from(position).ok())
            .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "planned seek overflows"))?;
        self.position = position;
        Ok(position)
    }
}

impl ReusePlan {
    pub(super) const fn report(&self) -> SectorLayoutReport {
        self.report
    }

    /// Reopens a read-only composed view of the planned artifact before a
    /// caller-owned sink sees any bytes. The ordinary CFB reader checks the
    /// complete FAT/MiniFAT, directory and physical-sector partition; stream
    /// readback proves that the planner emitted the model payloads at their
    /// assigned chains. The validation view is positional and does not
    /// allocate a second full output artifact.
    pub(super) fn validate(&self, streams: &[StreamInput<'_>]) -> Result<(), OleError> {
        let output_length = self.output_len()?;
        let directory_length = u64::try_from(self.directory_image.len())
            .map_err(|_error| invalid("CFB planned directory length exceeds u64"))?;
        let limits = OleFileLimits::for_writer(output_length, directory_length)?;
        let mut check = OleFile::open_with_limits(PlanCursor::new(self, streams), limits)?;
        for stream in streams {
            let mut refs = Vec::new();
            refs.try_reserve_exact(stream.path.len())
                .map_err(|source| OleError::allocation("CFB planned stream path", source))?;
            refs.extend(stream.path.iter().map(String::as_str));
            let actual = check.open_stream(&refs)?;
            if actual.as_slice() != stream.bytes {
                return Err(invalid("CFB reused layout stream readback differs"));
            }
        }
        Ok(())
    }

    fn output_len(&self) -> Result<u64, OleError> {
        let body = self
            .sectors
            .len()
            .checked_mul(self.sector_size)
            .ok_or_else(|| invalid("CFB reused layout output size overflows usize"))?;
        let length = self
            .header_sector
            .len()
            .checked_add(body)
            .ok_or_else(|| invalid("CFB reused layout output length overflows usize"))?;
        u64::try_from(length).map_err(|_error| invalid("CFB reused layout output exceeds u64"))
    }

    fn read_at(
        &self,
        streams: &[StreamInput<'_>],
        offset: u64,
        output: &mut [u8],
    ) -> Result<usize, OleError> {
        let length = self.output_len()?;
        if output.is_empty() || offset >= length {
            return Ok(0);
        }
        let available = usize::try_from(
            (length - offset).min(
                u64::try_from(output.len())
                    .map_err(|_error| invalid("CFB planned read length exceeds u64"))?,
            ),
        )
        .map_err(|_error| invalid("CFB planned read length exceeds usize"))?;
        let header_len = u64::try_from(self.header_sector.len())
            .map_err(|_error| invalid("CFB planned header length exceeds u64"))?;
        let mut written = 0usize;
        while written < available {
            let absolute = offset
                .checked_add(
                    u64::try_from(written)
                        .map_err(|_error| invalid("CFB planned read offset exceeds u64"))?,
                )
                .ok_or_else(|| invalid("CFB planned read offset overflows u64"))?;
            if absolute < header_len {
                let source_offset = usize::try_from(absolute)
                    .map_err(|_error| invalid("CFB planned header offset exceeds usize"))?;
                let count = (self.header_sector.len() - source_offset).min(available - written);
                output[written..written + count]
                    .copy_from_slice(&self.header_sector[source_offset..source_offset + count]);
                written += count;
                continue;
            }

            let body_offset = absolute - header_len;
            let sector_size_u64 = u64::try_from(self.sector_size)
                .map_err(|_error| invalid("CFB sector size exceeds u64"))?;
            let sector_index_u64 = body_offset / sector_size_u64;
            let sector_index = usize::try_from(sector_index_u64)
                .map_err(|_error| invalid("CFB planned sector index exceeds usize"))?;
            let within = usize::try_from(body_offset % sector_size_u64)
                .map_err(|_error| invalid("CFB planned sector offset exceeds usize"))?;
            let count = (self.sector_size - within).min(available - written);
            let planned = *self
                .sectors
                .get(sector_index)
                .ok_or_else(|| invalid("CFB planned read sector is outside the output"))?;
            self.read_sector(
                planned,
                streams,
                within,
                &mut output[written..written + count],
            )?;
            written += count;
        }
        Ok(written)
    }

    fn read_sector(
        &self,
        planned: PlannedSector,
        streams: &[StreamInput<'_>],
        within: usize,
        output: &mut [u8],
    ) -> Result<(), OleError> {
        match planned {
            PlannedSector::Free => output.fill(0),
            PlannedSector::Fat(position) => {
                let image = Self::image_run(&self.fat_image, position, 1, self.sector_size, "FAT")?;
                output.copy_from_slice(&image[within..within + output.len()]);
            },
            PlannedSector::MiniFat(position) => {
                let image = Self::image_run(
                    &self.minifat_image,
                    position,
                    1,
                    self.sector_size,
                    "MiniFAT",
                )?;
                output.copy_from_slice(&image[within..within + output.len()]);
            },
            PlannedSector::Directory(position) => {
                let image = Self::image_run(
                    &self.directory_image,
                    position,
                    1,
                    self.sector_size,
                    "directory",
                )?;
                output.copy_from_slice(&image[within..within + output.len()]);
            },
            PlannedSector::MiniStream(position) => {
                let image = Self::image_run(
                    &self.ministream_image,
                    position,
                    1,
                    self.sector_size,
                    "mini stream",
                )?;
                output.copy_from_slice(&image[within..within + output.len()]);
            },
            PlannedSector::Stream { index, chunk } => {
                let stream_index =
                    usize_from_u32(index, "CFB reused layout stream index does not fit usize")?;
                let bytes = streams
                    .get(stream_index)
                    .ok_or_else(|| invalid("CFB reused layout names a missing stream"))?
                    .bytes;
                output.fill(0);
                let source_start = usize_from_u32(chunk, "CFB stream chunk does not fit usize")?
                    .checked_mul(self.sector_size)
                    .and_then(|start| start.checked_add(within))
                    .ok_or_else(|| invalid("CFB reused layout chunk offset overflows usize"))?;
                if source_start < bytes.len() {
                    let count = output.len().min(bytes.len() - source_start);
                    output[..count].copy_from_slice(&bytes[source_start..source_start + count]);
                }
            },
        }
        Ok(())
    }
}

/// The planner's verdict.
#[derive(Debug)]
pub(super) enum Outcome {
    /// A gate declined; the caller serializes from scratch instead.
    Declined(SectorLayoutFallback),
    /// A layout that satisfies every invariant.
    Planned(Box<ReusePlan>),
}

/// Everything the writer's logical model contributes to a reused layout.
pub(super) struct ModelInputs<'a> {
    pub(super) sector_size: usize,
    pub(super) mini_sector_size: usize,
    pub(super) mini_stream_cutoff: u32,
    pub(super) streams: &'a [StreamInput<'a>],
    pub(super) storages: &'a BTreeSet<Vec<String>>,
    pub(super) storage_class_ids: &'a BTreeMap<Vec<String>, [u8; 16]>,
    /// `None` preserves the source root CLSID; `Some([0; 16])` explicitly
    /// clears it while `Some(nonzero)` requests the supplied identifier.
    pub(super) root_class_id: Option<[u8; 16]>,
}

struct SectorPool {
    claimed: Vec<SectorOwner>,
    free: BTreeSet<u32>,
    reclaimed: u32,
    appended: u32,
}

impl SectorPool {
    fn claim(&mut self, sector: u32, owner: SectorOwner) -> Result<(), OleError> {
        let index = usize_from_u32(sector, "CFB reused layout sector does not fit usize")?;
        let slot = self
            .claimed
            .get_mut(index)
            .ok_or_else(|| invalid("CFB reused layout claims a sector outside the artifact"))?;
        if *slot != SectorOwner::Free {
            return Err(invalid("CFB reused layout claims a sector twice"));
        }
        *slot = owner;
        Ok(())
    }

    fn release(&mut self, sector: u32) -> Result<(), OleError> {
        let index = usize_from_u32(sector, "CFB reused layout sector does not fit usize")?;
        let slot = self
            .claimed
            .get_mut(index)
            .ok_or_else(|| invalid("CFB reused layout releases a sector outside the artifact"))?;
        *slot = SectorOwner::Free;
        self.free.insert(sector);
        Ok(())
    }

    fn allocate(&mut self, owner: SectorOwner) -> Result<u32, OleError> {
        if let Some(sector) = self.free.iter().next().copied() {
            self.free.remove(&sector);
            self.claim(sector, owner)?;
            self.reclaimed = self.reclaimed.saturating_add(1);
            return Ok(sector);
        }
        let sector = u32::try_from(self.claimed.len())
            .map_err(|_err| invalid("CFB reused layout sector index exceeds u32"))?;
        if sector >= MAXREGSECT {
            return Err(invalid("CFB reused layout sector index exceeds MAXREGSECT"));
        }
        self.claimed
            .try_reserve(1)
            .map_err(|source| OleError::allocation("CFB reused layout sector map", source))?;
        self.claimed.push(owner);
        self.appended = self.appended.saturating_add(1);
        Ok(sector)
    }

    fn high_water(&self) -> Result<u32, OleError> {
        let mut high = 0u32;
        for (index, owner) in self.claimed.iter().enumerate() {
            if *owner != SectorOwner::Free {
                let index = u32::try_from(index)
                    .map_err(|_err| invalid("CFB reused layout sector index exceeds u32"))?;
                high = index
                    .checked_add(1)
                    .ok_or_else(|| invalid("CFB reused layout high-water mark overflows u32"))?;
            }
        }
        Ok(high)
    }
}

/// Plans a serialization that reuses `source`'s sector layout, or declines.
///
/// # Errors
///
/// Returns [`OleError::InvalidData`] when a geometry or chain length cannot be
/// represented, and an allocation error when a bounded reservation fails. A
/// gate that the model does not satisfy is reported as
/// [`Outcome::Declined`] rather than as an error, because the from-scratch
/// serialization is always available.
#[allow(
    clippy::too_many_lines,
    reason = "the planner is one allocation pass over the source; splitting it would hide the ordering the invariants depend on"
)]
pub(super) fn plan_reuse(
    source: &SourceLayout,
    model: &ModelInputs<'_>,
) -> Result<Outcome, OleError> {
    if source.sector_size != model.sector_size
        || source.mini_sector_size != model.mini_sector_size
        || source.mini_stream_cutoff != model.mini_stream_cutoff
    {
        return Ok(Outcome::Declined(SectorLayoutFallback::GeometryChanged));
    }
    if source.difat_declared()? {
        return Ok(Outcome::Declined(SectorLayoutFallback::SourceDeclaresDifat));
    }
    let mini_unit = model.mini_sector_size;
    let mini_unit_u64 =
        u64::try_from(mini_unit).map_err(|_err| invalid("CFB mini sector size exceeds u64"))?;
    if mini_unit_u64 == 0 || source.ministream_size % mini_unit_u64 != 0 {
        return Ok(Outcome::Declined(
            SectorLayoutFallback::SourceMiniStreamUnaligned,
        ));
    }

    // --- the directory shape must be the source's ---
    let mut model_streams: BTreeMap<&[String], usize> = BTreeMap::new();
    for (index, stream) in model.streams.iter().enumerate() {
        if model_streams.insert(stream.path, index).is_some() {
            return Ok(Outcome::Declined(
                SectorLayoutFallback::DirectoryShapeChanged,
            ));
        }
    }
    if model_streams.len() != source.stream_paths.len()
        || model.storages.len() != source.storage_paths.len()
    {
        return Ok(Outcome::Declined(
            SectorLayoutFallback::DirectoryShapeChanged,
        ));
    }
    let mut stream_sid = filled_vec(model.streams.len(), 0, "CFB reused layout stream SIDs")?;
    for (path, sid) in &source.stream_paths {
        let Some(index) = model_streams.get(path.as_slice()).copied() else {
            return Ok(Outcome::Declined(
                SectorLayoutFallback::DirectoryShapeChanged,
            ));
        };
        stream_sid[index] = *sid;
    }
    for path in model.storages {
        if !source.storage_paths.contains_key(path) {
            return Ok(Outcome::Declined(
                SectorLayoutFallback::DirectoryShapeChanged,
            ));
        }
    }
    // An explicitly cleared class identifier is safe to patch in place. A
    // different nonzero identifier still declines, keeping the historical
    // class-ID gate for callers that request a new object identity.
    if let Some(root_class_id) = model.root_class_id
        && root_class_id != source.root_class_id
        && root_class_id != [0u8; 16]
    {
        return Ok(Outcome::Declined(SectorLayoutFallback::ClassIdChanged));
    }
    for (path, class_id) in model.storage_class_ids {
        let Some(sid) = source.storage_paths.get(path).copied() else {
            return Ok(Outcome::Declined(
                SectorLayoutFallback::DirectoryShapeChanged,
            ));
        };
        let sid_index = usize_from_u32(sid, "CFB reused layout storage SID does not fit usize")?;
        let entry = source
            .entries
            .get(sid_index)
            .ok_or_else(|| invalid("CFB reused layout storage SID is outside the directory"))?;
        if *class_id != entry.class_id && *class_id != [0u8; 16] {
            return Ok(Outcome::Declined(SectorLayoutFallback::ClassIdChanged));
        }
    }

    // --- mini sector allocation, inside the mini stream ---
    let source_mini_count = u32::try_from(source.ministream_size / mini_unit_u64)
        .map_err(|_err| invalid("CFB source mini sector count exceeds u32"))?;
    let source_mini_count_usize = usize_from_u32(
        source_mini_count,
        "CFB source mini sector count does not fit this platform",
    )?;
    let mut mini_seen = LayoutBitSet::try_with_capacity(
        source_mini_count_usize,
        "CFB reused layout mini-sector bitmap",
    )?;
    let mut mini_free: BTreeSet<u32> = BTreeSet::new();
    let mut mini_chains: Vec<Vec<u32>> = Vec::new();
    mini_chains
        .try_reserve_exact(model.streams.len())
        .map_err(|source| OleError::allocation("CFB reused layout mini chains", source))?;
    let mut needs: Vec<(bool, u32)> = Vec::new();
    needs
        .try_reserve_exact(model.streams.len())
        .map_err(|source| OleError::allocation("CFB reused layout stream needs", source))?;

    for (index, stream) in model.streams.iter().enumerate() {
        let sid = stream_sid[index];
        let sid_index = usize_from_u32(sid, "CFB reused layout stream SID does not fit usize")?;
        let entry = source
            .entries
            .get(sid_index)
            .ok_or_else(|| invalid("CFB reused layout stream SID is outside the directory"))?;
        if entry.entry_type != STGTY_STREAM {
            return Ok(Outcome::Declined(
                SectorLayoutFallback::DirectoryShapeChanged,
            ));
        }
        let new_len = u64::try_from(stream.bytes.len())
            .map_err(|_err| invalid("CFB stream length exceeds u64"))?;
        let new_is_mini = new_len > 0 && new_len < u64::from(model.mini_stream_cutoff);
        let need = if new_is_mini {
            units_for(new_len, mini_unit)?
        } else {
            0
        };
        needs.push((new_is_mini, need));
        let mut kept: Vec<u32> = Vec::new();
        if entry.is_mini && entry.size > 0 {
            let old = walk_chain(
                &source.minifat,
                entry.start_sector,
                source_mini_count,
                "mini stream",
            )?;
            if u64::from(units_for(entry.size, mini_unit)?)
                != u64::try_from(old.len())
                    .map_err(|_err| invalid("CFB source mini chain length exceeds u64"))?
            {
                return Ok(Outcome::Declined(SectorLayoutFallback::PlanRejected));
            }
            let keep = if new_is_mini {
                need.min(
                    u32::try_from(old.len())
                        .map_err(|_err| invalid("CFB source mini chain length exceeds u32"))?,
                )
            } else {
                0
            };
            for (position, mini) in old.iter().enumerate() {
                let mini_index =
                    usize_from_u32(*mini, "CFB source mini sector does not fit this platform")?;
                mini_seen.insert(mini_index)?;
                if u32::try_from(position)
                    .map_err(|_err| invalid("CFB source mini position exceeds u32"))?
                    < keep
                {
                    kept.try_reserve(1).map_err(|source| {
                        OleError::allocation("CFB reused layout mini chain", source)
                    })?;
                    kept.push(*mini);
                } else {
                    mini_free.insert(*mini);
                }
            }
        }
        mini_chains.push(kept);
    }
    for mini in 0..source_mini_count {
        let mini_index = usize_from_u32(mini, "CFB source mini sector does not fit this platform")?;
        let seen = mini_seen.contains(mini_index);
        let entry = source.minifat.get(mini_index).copied().unwrap_or(FREESECT);
        if !seen && entry == FREESECT {
            mini_free.insert(mini);
        }
    }

    let mut mini_append = source_mini_count;
    for (index, (is_mini, need)) in needs.iter().copied().enumerate() {
        if !is_mini {
            continue;
        }
        while u32::try_from(mini_chains[index].len())
            .map_err(|_err| invalid("CFB reused layout mini chain length exceeds u32"))?
            < need
        {
            let mini = if let Some(free) = mini_free.iter().next().copied() {
                mini_free.remove(&free);
                free
            } else {
                let mini = mini_append;
                mini_append = mini_append
                    .checked_add(1)
                    .ok_or_else(|| invalid("CFB mini sector index overflows u32"))?;
                mini
            };
            mini_chains[index]
                .try_reserve(1)
                .map_err(|source| OleError::allocation("CFB reused layout mini chain", source))?;
            mini_chains[index].push(mini);
        }
    }
    let mut mini_count = 0u32;
    for chain in &mini_chains {
        for mini in chain {
            mini_count = mini_count.max(mini.saturating_add(1));
        }
    }
    let ministream_bytes = u64::from(mini_count)
        .checked_mul(mini_unit_u64)
        .ok_or_else(|| invalid("CFB mini stream size overflows u64"))?;

    // --- sector allocation, inside the artifact ---
    let mut claimed = Vec::new();
    let source_sector_count = usize_from_u32(
        source.sector_count,
        "CFB source sector count does not fit this platform",
    )?;
    claimed
        .try_reserve_exact(source_sector_count)
        .map_err(|source| OleError::allocation("CFB reused layout sector map", source))?;
    claimed.resize(source_sector_count, SectorOwner::Free);
    let mut pool = SectorPool {
        claimed,
        free: BTreeSet::new(),
        reclaimed: 0,
        appended: 0,
    };

    for sector in &source.dir_chain {
        pool.claim(*sector, SectorOwner::Directory)?;
    }
    for sector in &source.fat_sectors {
        pool.claim(*sector, SectorOwner::Fat)?;
    }

    let mut stream_chains: Vec<Vec<u32>> = Vec::new();
    stream_chains
        .try_reserve_exact(model.streams.len())
        .map_err(|source| OleError::allocation("CFB reused layout stream chains", source))?;
    let mut stream_needs: Vec<u32> = Vec::new();
    stream_needs
        .try_reserve_exact(model.streams.len())
        .map_err(|source| OleError::allocation("CFB reused layout stream needs", source))?;
    let mut kept_sectors = 0u32;
    for (index, stream) in model.streams.iter().enumerate() {
        let sid = stream_sid[index];
        let sid_index = usize_from_u32(sid, "CFB reused layout stream SID does not fit usize")?;
        let entry = source
            .entries
            .get(sid_index)
            .ok_or_else(|| invalid("CFB reused layout stream SID is outside the directory"))?;
        let (is_mini, _) = needs[index];
        let new_len = u64::try_from(stream.bytes.len())
            .map_err(|_err| invalid("CFB stream length exceeds u64"))?;
        let need = if is_mini || new_len == 0 {
            0
        } else {
            units_for(new_len, model.sector_size)?
        };
        stream_needs.push(need);
        let mut kept: Vec<u32> = Vec::new();
        if !entry.is_mini && entry.size > 0 {
            let old = walk_chain(
                &source.fat,
                entry.start_sector,
                source.sector_count,
                "stream",
            )?;
            if u64::from(units_for(entry.size, model.sector_size)?)
                != u64::try_from(old.len())
                    .map_err(|_err| invalid("CFB source stream chain length exceeds u64"))?
            {
                return Ok(Outcome::Declined(SectorLayoutFallback::PlanRejected));
            }
            let keep = need.min(
                u32::try_from(old.len())
                    .map_err(|_err| invalid("CFB source stream chain length exceeds u32"))?,
            );
            let keep_usize = usize_from_u32(keep, "CFB stream chain length does not fit usize")?;
            for sector in old.iter().take(keep_usize) {
                pool.claim(*sector, SectorOwner::Stream(sid))?;
                kept.try_reserve(1).map_err(|source| {
                    OleError::allocation("CFB reused layout stream chain", source)
                })?;
                kept.push(*sector);
                kept_sectors = kept_sectors.saturating_add(1);
            }
        }
        stream_chains.push(kept);
    }

    let root_need = units_for(ministream_bytes, model.sector_size)?;
    let mut root_chain: Vec<u32> = Vec::new();
    let root_need_usize = usize_from_u32(root_need, "CFB root chain length does not fit usize")?;
    for sector in source.root_chain.iter().take(root_need_usize) {
        pool.claim(*sector, SectorOwner::MiniStream)?;
        root_chain.try_reserve(1).map_err(|source| {
            OleError::allocation("CFB reused layout mini stream chain", source)
        })?;
        root_chain.push(*sector);
    }

    let minifat_bytes = u64::from(mini_count)
        .checked_mul(4)
        .ok_or_else(|| invalid("CFB MiniFAT size overflows u64"))?;
    let minifat_need = units_for(minifat_bytes, model.sector_size)?;
    let mut minifat_chain: Vec<u32> = Vec::new();
    let minifat_need_usize =
        usize_from_u32(minifat_need, "CFB MiniFAT chain length does not fit usize")?;
    for sector in source.minifat_chain.iter().take(minifat_need_usize) {
        pool.claim(*sector, SectorOwner::MiniFat)?;
        minifat_chain
            .try_reserve(1)
            .map_err(|source| OleError::allocation("CFB reused layout MiniFAT chain", source))?;
        minifat_chain.push(*sector);
    }

    // Everything the source held that this save did not keep is available.
    for sector in 0..source.sector_count {
        let sector_index = usize_from_u32(sector, "CFB reused layout sector does not fit usize")?;
        if pool.claimed[sector_index] == SectorOwner::Free {
            pool.free.insert(sector);
        }
    }

    for (index, need) in stream_needs.iter().copied().enumerate() {
        let sid = stream_sid[index];
        while u32::try_from(stream_chains[index].len())
            .map_err(|_err| invalid("CFB reused layout stream chain length exceeds u32"))?
            < need
        {
            let sector = pool.allocate(SectorOwner::Stream(sid))?;
            stream_chains[index]
                .try_reserve(1)
                .map_err(|source| OleError::allocation("CFB reused layout stream chain", source))?;
            stream_chains[index].push(sector);
        }
    }
    while u32::try_from(root_chain.len())
        .map_err(|_err| invalid("CFB reused layout root chain length exceeds u32"))?
        < root_need
    {
        let sector = pool.allocate(SectorOwner::MiniStream)?;
        root_chain.try_reserve(1).map_err(|source| {
            OleError::allocation("CFB reused layout mini stream chain", source)
        })?;
        root_chain.push(sector);
    }
    while u32::try_from(minifat_chain.len())
        .map_err(|_err| invalid("CFB reused layout MiniFAT chain length exceeds u32"))?
        < minifat_need
    {
        let sector = pool.allocate(SectorOwner::MiniFat)?;
        minifat_chain
            .try_reserve(1)
            .map_err(|source| OleError::allocation("CFB reused layout MiniFAT chain", source))?;
        minifat_chain.push(sector);
    }

    // --- the FAT covers the sectors it is stored in, so its size is a fixed point ---
    let entries_per_fat_sector = u32::try_from(model.sector_size / 4)
        .map_err(|_err| invalid("CFB FAT geometry exceeds u32"))?;
    if entries_per_fat_sector == 0 {
        return Err(invalid("CFB FAT sector holds no entries"));
    }
    let mut fat_sectors = source.fat_sectors.clone();
    let mut converged = false;
    for _ in 0..LAYOUT_FIXED_POINT_ROUNDS {
        let high = pool.high_water()?;
        let need = high.div_ceil(entries_per_fat_sector);
        let header_difat_entries = u32::try_from(HEADER_DIFAT_ENTRIES)
            .map_err(|_err| invalid("CFB header DIFAT entry count exceeds u32"))?;
        if need > header_difat_entries {
            return Ok(Outcome::Declined(SectorLayoutFallback::OutputNeedsDifat));
        }
        let have = u32::try_from(fat_sectors.len())
            .map_err(|_err| invalid("CFB reused layout FAT sector count exceeds u32"))?;
        if need == have {
            converged = true;
            break;
        }
        if need > have {
            for _ in 0..(need - have) {
                let sector = pool.allocate(SectorOwner::Fat)?;
                fat_sectors.try_reserve(1).map_err(|source| {
                    OleError::allocation("CFB reused layout FAT sectors", source)
                })?;
                fat_sectors.push(sector);
            }
        } else {
            fat_sectors.sort_unstable();
            for _ in 0..(have - need) {
                let sector = fat_sectors
                    .pop()
                    .ok_or_else(|| invalid("CFB reused layout released a missing FAT sector"))?;
                pool.release(sector)?;
            }
        }
    }
    if !converged {
        return Ok(Outcome::Declined(SectorLayoutFallback::PlanRejected));
    }
    fat_sectors.sort_unstable();
    let output_sectors = pool.high_water()?;
    if output_sectors == 0 {
        return Ok(Outcome::Declined(SectorLayoutFallback::PlanRejected));
    }
    super::core::validate_output_size(model.sector_size, output_sectors)?;

    // --- images ---
    let fat_sector_count = u32::try_from(fat_sectors.len())
        .map_err(|_err| invalid("CFB reused layout FAT sector count exceeds u32"))?;
    let fat_entry_count = u64::from(fat_sector_count)
        .checked_mul(u64::from(entries_per_fat_sector))
        .ok_or_else(|| invalid("CFB FAT entry count overflows u64"))?;
    let fat_entries = usize::try_from(fat_entry_count)
        .map_err(|_err| invalid("CFB FAT entry count exceeds usize"))?;
    let output_sectors_usize = usize_from_u32(
        output_sectors,
        "CFB reused layout output sector count does not fit this platform",
    )?;
    if fat_entries < output_sectors_usize {
        return Ok(Outcome::Declined(SectorLayoutFallback::PlanRejected));
    }
    let mut fat = filled_vec(fat_entries, FREESECT, "CFB reused layout FAT")?;
    let link = |chain: &[u32], fat: &mut Vec<u32>| -> Result<(), OleError> {
        for (position, sector) in chain.iter().enumerate() {
            let next = chain.get(position + 1).copied().unwrap_or(ENDOFCHAIN);
            let sector_index = usize_from_u32(
                *sector,
                "CFB reused layout FAT sector does not fit this platform",
            )?;
            let slot = fat
                .get_mut(sector_index)
                .ok_or_else(|| invalid("CFB reused layout FAT entry is outside the table"))?;
            if *slot != FREESECT {
                return Err(invalid("CFB reused layout links a sector into two chains"));
            }
            *slot = next;
        }
        Ok(())
    };
    for chain in &stream_chains {
        link(chain, &mut fat)?;
    }
    link(&root_chain, &mut fat)?;
    link(&minifat_chain, &mut fat)?;
    link(&source.dir_chain, &mut fat)?;
    for sector in &fat_sectors {
        let sector_index = usize_from_u32(
            *sector,
            "CFB reused layout FAT sector does not fit this platform",
        )?;
        let slot = fat
            .get_mut(sector_index)
            .ok_or_else(|| invalid("CFB reused layout FAT sector is outside the table"))?;
        if *slot != FREESECT {
            return Err(invalid("CFB reused layout stores the FAT inside a chain"));
        }
        *slot = FATSECT;
    }
    let fat_image_bytes = fat_entries
        .checked_mul(4)
        .ok_or_else(|| invalid("CFB reused layout FAT image size overflows usize"))?;
    let mut fat_image = reserved_vec::<u8>(fat_image_bytes, "CFB reused layout FAT image")?;
    for entry in &fat {
        fat_image.extend_from_slice(&entry.to_le_bytes());
    }

    let minifat_entry_count = minifat_need_usize
        .checked_mul(model.sector_size / 4)
        .ok_or_else(|| invalid("CFB reused layout MiniFAT entry count overflows usize"))?;
    let mut minifat = filled_vec(minifat_entry_count, FREESECT, "CFB reused layout MiniFAT")?;
    for chain in &mini_chains {
        for (position, mini) in chain.iter().enumerate() {
            let next = chain.get(position + 1).copied().unwrap_or(ENDOFCHAIN);
            let mini_index = usize_from_u32(
                *mini,
                "CFB reused layout mini sector does not fit this platform",
            )?;
            let slot = minifat
                .get_mut(mini_index)
                .ok_or_else(|| invalid("CFB reused layout MiniFAT entry is outside the table"))?;
            if *slot != FREESECT {
                return Err(invalid(
                    "CFB reused layout links a mini sector into two chains",
                ));
            }
            *slot = next;
        }
    }
    let minifat_image_bytes = minifat
        .len()
        .checked_mul(4)
        .ok_or_else(|| invalid("CFB reused layout MiniFAT image size overflows usize"))?;
    let mut minifat_image =
        reserved_vec::<u8>(minifat_image_bytes, "CFB reused layout MiniFAT image")?;
    for entry in &minifat {
        minifat_image.extend_from_slice(&entry.to_le_bytes());
    }

    let ministream_image_len = root_need_usize
        .checked_mul(model.sector_size)
        .ok_or_else(|| invalid("CFB reused layout mini stream image size overflows usize"))?;
    let mut ministream_image =
        zeroed_vec(ministream_image_len, "CFB reused layout mini stream image")?;
    let carried = source.ministream_image.len().min(ministream_image_len);
    ministream_image[..carried].copy_from_slice(&source.ministream_image[..carried]);
    for (index, chain) in mini_chains.iter().enumerate() {
        let bytes = model.streams[index].bytes;
        for (position, mini) in chain.iter().enumerate() {
            let mini_index = usize_from_u32(*mini, "CFB mini sector does not fit this platform")?;
            let start = mini_index
                .checked_mul(mini_unit)
                .ok_or_else(|| invalid("CFB mini sector offset overflows usize"))?;
            let end = start
                .checked_add(mini_unit)
                .ok_or_else(|| invalid("CFB mini sector end overflows usize"))?;
            let taken = position
                .checked_mul(mini_unit)
                .ok_or_else(|| invalid("CFB mini stream payload offset overflows usize"))?;
            let chunk_end = taken
                .checked_add(mini_unit)
                .ok_or_else(|| invalid("CFB mini stream payload end overflows usize"))?;
            let chunk = bytes
                .get(taken..chunk_end.min(bytes.len()))
                .ok_or_else(|| invalid("CFB mini sector chunk is outside the stream"))?;
            let target = ministream_image
                .get_mut(start..end)
                .ok_or_else(|| invalid("CFB mini sector is outside the mini stream image"))?;
            target[..chunk.len()].copy_from_slice(chunk);
            target[chunk.len()..].fill(0);
        }
    }

    let mut directory_image = source.directory_image.clone();
    for (index, stream) in model.streams.iter().enumerate() {
        let sid = stream_sid[index];
        let new_len = u64::try_from(stream.bytes.len())
            .map_err(|_err| invalid("CFB stream length exceeds u64"))?;
        let (is_mini, _) = needs[index];
        let start = if new_len == 0 {
            ENDOFCHAIN
        } else if is_mini {
            mini_chains[index].first().copied().unwrap_or(ENDOFCHAIN)
        } else {
            stream_chains[index].first().copied().unwrap_or(ENDOFCHAIN)
        };
        patch_entry(&mut directory_image, sid, start, new_len)?;
    }
    if let Some(root_class_id) = model.root_class_id
        && root_class_id == [0u8; 16]
        && source.root_class_id != [0u8; 16]
    {
        patch_class_id(&mut directory_image, 0, root_class_id)?;
    }
    for (path, class_id) in model.storage_class_ids {
        let Some(sid) = source.storage_paths.get(path).copied() else {
            return Ok(Outcome::Declined(
                SectorLayoutFallback::DirectoryShapeChanged,
            ));
        };
        if *class_id == [0u8; 16] {
            let sid_index =
                usize_from_u32(sid, "CFB reused layout storage SID does not fit usize")?;
            let source_class_id = source
                .entries
                .get(sid_index)
                .ok_or_else(|| invalid("CFB reused layout storage SID is outside the directory"))?
                .class_id;
            if source_class_id != [0u8; 16] {
                patch_class_id(&mut directory_image, sid, *class_id)?;
            }
        }
    }
    let root_start = root_chain.first().copied().unwrap_or(ENDOFCHAIN);
    patch_entry(&mut directory_image, 0, root_start, ministream_bytes)?;

    let mut header_sector = source.header_sector.clone();
    let fat_sector_count = u32::try_from(fat_sectors.len())
        .map_err(|_err| invalid("CFB reused layout FAT sector count exceeds u32"))?;
    let minifat_chain_count = u32::try_from(minifat_chain.len())
        .map_err(|_err| invalid("CFB reused layout MiniFAT sector count exceeds u32"))?;
    write_u32(&mut header_sector, NUM_FAT_SECTORS_OFFSET, fat_sector_count)?;
    write_u32(
        &mut header_sector,
        FIRST_DIR_SECTOR_OFFSET,
        source.dir_chain.first().copied().unwrap_or(ENDOFCHAIN),
    )?;
    write_u32(
        &mut header_sector,
        FIRST_MINIFAT_SECTOR_OFFSET,
        minifat_chain.first().copied().unwrap_or(ENDOFCHAIN),
    )?;
    write_u32(
        &mut header_sector,
        NUM_MINIFAT_SECTORS_OFFSET,
        minifat_chain_count,
    )?;
    write_u32(&mut header_sector, FIRST_DIFAT_SECTOR_OFFSET, ENDOFCHAIN)?;
    write_u32(&mut header_sector, NUM_DIFAT_SECTORS_OFFSET, 0)?;
    for slot in 0..HEADER_DIFAT_ENTRIES {
        let value = fat_sectors.get(slot).copied().unwrap_or(FREESECT);
        write_u32(&mut header_sector, HEADER_DIFAT_OFFSET + slot * 4, value)?;
    }
    let _ = read_u32(&header_sector, NUM_DIR_SECTORS_OFFSET)?;

    // --- the emission plan, and the partition invariant it must satisfy ---
    let mut sectors = Vec::new();
    sectors
        .try_reserve_exact(output_sectors_usize)
        .map_err(|source| OleError::allocation("CFB reused layout sector plan", source))?;
    sectors.resize(output_sectors_usize, PlannedSector::Free);
    let mut place = |sector: u32, value: PlannedSector| -> Result<(), OleError> {
        let sector_index = usize_from_u32(sector, "CFB reused layout sector does not fit usize")?;
        let slot = sectors
            .get_mut(sector_index)
            .ok_or_else(|| invalid("CFB reused layout plans a sector outside the output"))?;
        if *slot != PlannedSector::Free {
            return Err(invalid("CFB reused layout plans a sector twice"));
        }
        *slot = value;
        Ok(())
    };
    for (position, sector) in fat_sectors.iter().enumerate() {
        let position = u32::try_from(position)
            .map_err(|_err| invalid("CFB reused layout FAT position exceeds u32"))?;
        place(*sector, PlannedSector::Fat(position))?;
    }
    for (position, sector) in minifat_chain.iter().enumerate() {
        let position = u32::try_from(position)
            .map_err(|_err| invalid("CFB reused layout MiniFAT position exceeds u32"))?;
        place(*sector, PlannedSector::MiniFat(position))?;
    }
    for (position, sector) in source.dir_chain.iter().enumerate() {
        let position = u32::try_from(position)
            .map_err(|_err| invalid("CFB reused layout directory position exceeds u32"))?;
        place(*sector, PlannedSector::Directory(position))?;
    }
    for (position, sector) in root_chain.iter().enumerate() {
        let position = u32::try_from(position)
            .map_err(|_err| invalid("CFB reused layout mini stream position exceeds u32"))?;
        place(*sector, PlannedSector::MiniStream(position))?;
    }
    for (index, chain) in stream_chains.iter().enumerate() {
        for (position, sector) in chain.iter().enumerate() {
            let index_u32 = u32::try_from(index)
                .map_err(|_err| invalid("CFB reused layout stream index exceeds u32"))?;
            let position = u32::try_from(position)
                .map_err(|_err| invalid("CFB reused layout stream position exceeds u32"))?;
            place(
                *sector,
                PlannedSector::Stream {
                    index: index_u32,
                    chunk: position,
                },
            )?;
        }
    }

    let free_sectors = u32::try_from(
        sectors
            .iter()
            .filter(|planned| **planned == PlannedSector::Free)
            .count(),
    )
    .map_err(|_err| invalid("CFB reused layout free-sector count exceeds u32"))?;
    let appended_sectors = output_sectors.saturating_sub(source.sector_count);
    let report = SectorLayoutReport {
        reused: true,
        fallback: None,
        output_sectors,
        kept_sectors,
        rewritten_sectors: output_sectors
            .saturating_sub(kept_sectors)
            .saturating_sub(free_sectors),
        appended_sectors,
        reclaimed_sectors: pool.reclaimed,
        free_sectors,
    };

    Ok(Outcome::Planned(Box::new(ReusePlan {
        sector_size: model.sector_size,
        header_sector,
        sectors,
        fat_image,
        minifat_image,
        directory_image,
        ministream_image,
        report,
    })))
}

fn patch_entry(
    directory_image: &mut [u8],
    sid: u32,
    start_sector: u32,
    size: u64,
) -> Result<(), OleError> {
    let base = usize_from_u32(sid, "CFB directory SID does not fit usize")?
        .checked_mul(DIRENTRY_SIZE)
        .ok_or_else(|| invalid("CFB directory entry offset overflows usize"))?;
    let end = base
        .checked_add(DIRENTRY_SIZE)
        .ok_or_else(|| invalid("CFB directory entry end overflows usize"))?;
    let entry = directory_image
        .get_mut(base..end)
        .ok_or_else(|| invalid("CFB directory entry is outside the directory image"))?;
    write_u32(entry, ENTRY_START_SECTOR_OFFSET, start_sector)?;
    write_u64(entry, ENTRY_STREAM_SIZE_OFFSET, size)?;
    Ok(())
}

fn patch_class_id(
    directory_image: &mut [u8],
    sid: u32,
    class_id: [u8; 16],
) -> Result<(), OleError> {
    let base = usize_from_u32(sid, "CFB directory SID does not fit usize")?
        .checked_mul(DIRENTRY_SIZE)
        .ok_or_else(|| invalid("CFB directory entry offset overflows usize"))?;
    let start = base
        .checked_add(0x50)
        .ok_or_else(|| invalid("CFB directory class-ID offset overflows usize"))?;
    let end = start
        .checked_add(16)
        .ok_or_else(|| invalid("CFB directory class-ID end overflows usize"))?;
    let entry = directory_image
        .get_mut(start..end)
        .ok_or_else(|| invalid("CFB directory entry class-ID is outside the image"))?;
    entry.copy_from_slice(&class_id);
    Ok(())
}

/// One sector of zeroes, reused for every unallocated run.
const ZEROES: [u8; 4096] = [0; 4096];

fn write_zeroes<W: Write>(sink: &mut W, mut bytes: usize) -> Result<(), OleError> {
    while bytes > 0 {
        let take = bytes.min(ZEROES.len());
        let chunk = ZEROES
            .get(..take)
            .ok_or_else(|| invalid("CFB reused layout padding exceeds one sector"))?;
        sink.write_all(chunk)?;
        bytes -= take;
    }
    Ok(())
}

fn follows(previous: PlannedSector, next: PlannedSector) -> bool {
    match (previous, next) {
        (PlannedSector::Free, PlannedSector::Free) => true,
        (PlannedSector::Fat(before), PlannedSector::Fat(after))
        | (PlannedSector::MiniFat(before), PlannedSector::MiniFat(after))
        | (PlannedSector::Directory(before), PlannedSector::Directory(after))
        | (PlannedSector::MiniStream(before), PlannedSector::MiniStream(after)) => {
            after == before.wrapping_add(1)
        },
        (
            PlannedSector::Stream {
                index: before_index,
                chunk: before_chunk,
            },
            PlannedSector::Stream {
                index: after_index,
                chunk: after_chunk,
            },
        ) => before_index == after_index && after_chunk == before_chunk.wrapping_add(1),
        _ => false,
    }
}

impl ReusePlan {
    fn image_run<'a>(
        image: &'a [u8],
        position: u32,
        run: usize,
        sector_size: usize,
        resource: &'static str,
    ) -> Result<&'a [u8], OleError> {
        let start = usize_from_u32(position, "CFB image position does not fit usize")?
            .checked_mul(sector_size)
            .ok_or_else(|| invalid("CFB reused layout image offset overflows usize"))?;
        let end = run
            .checked_mul(sector_size)
            .and_then(|bytes| start.checked_add(bytes))
            .ok_or_else(|| invalid("CFB reused layout image end overflows usize"))?;
        image
            .get(start..end)
            .ok_or_else(|| invalid(&format!("CFB reused layout {resource} run is truncated")))
    }

    /// Writes the already validated planned artifact to `sink` in ascending
    /// sector order. Callers use [`Self::validate`] first when the destination
    /// must not observe a candidate before ordinary CFB validation.
    ///
    /// Runs of consecutive sectors that come from the same image — a stream's
    /// payload, the FAT, the directory — are written in one call, so a stream
    /// whose chain is contiguous costs a single copy of its bytes and no
    /// padding pass.
    ///
    /// # Errors
    ///
    /// Returns [`OleError::Io`] when `sink` fails, and
    /// [`OleError::InvalidData`] when a planned run leaves its image.
    pub(super) fn emit<W: Write>(
        &self,
        streams: &[StreamInput<'_>],
        sink: &mut W,
    ) -> Result<(), OleError> {
        sink.write_all(&self.header_sector)?;
        let sector_size = self.sector_size;
        let mut index = 0usize;
        while index < self.sectors.len() {
            let planned = self.sectors[index];
            let mut run = 1usize;
            while index
                .checked_add(run)
                .is_some_and(|next| next < self.sectors.len())
                && follows(self.sectors[index + run - 1], self.sectors[index + run])
            {
                run += 1;
            }
            match planned {
                PlannedSector::Free => {
                    let bytes = run
                        .checked_mul(sector_size)
                        .ok_or_else(|| invalid("CFB reused layout free run overflows usize"))?;
                    write_zeroes(sink, bytes)?;
                },
                PlannedSector::Fat(position) => sink.write_all(Self::image_run(
                    &self.fat_image,
                    position,
                    run,
                    sector_size,
                    "FAT",
                )?)?,
                PlannedSector::MiniFat(position) => sink.write_all(Self::image_run(
                    &self.minifat_image,
                    position,
                    run,
                    sector_size,
                    "MiniFAT",
                )?)?,
                PlannedSector::Directory(position) => sink.write_all(Self::image_run(
                    &self.directory_image,
                    position,
                    run,
                    sector_size,
                    "directory",
                )?)?,
                PlannedSector::MiniStream(position) => sink.write_all(Self::image_run(
                    &self.ministream_image,
                    position,
                    run,
                    sector_size,
                    "mini stream",
                )?)?,
                PlannedSector::Stream {
                    index: stream,
                    chunk,
                } => {
                    let stream_index = usize_from_u32(
                        stream,
                        "CFB reused layout stream index does not fit usize",
                    )?;
                    let bytes = streams
                        .get(stream_index)
                        .ok_or_else(|| invalid("CFB reused layout names a missing stream"))?
                        .bytes;
                    let start = usize_from_u32(chunk, "CFB stream chunk does not fit usize")?
                        .checked_mul(sector_size)
                        .ok_or_else(|| invalid("CFB reused layout chunk offset overflows usize"))?;
                    let end = run
                        .checked_mul(sector_size)
                        .and_then(|span| start.checked_add(span))
                        .ok_or_else(|| invalid("CFB reused layout chunk end overflows usize"))?;
                    let available = end.min(bytes.len());
                    let payload = bytes
                        .get(start..available)
                        .ok_or_else(|| invalid("CFB reused layout chunk is outside its stream"))?;
                    sink.write_all(payload)?;
                    write_zeroes(sink, end - available)?;
                },
            }
            index += run;
        }
        sink.flush()?;
        Ok(())
    }
}
