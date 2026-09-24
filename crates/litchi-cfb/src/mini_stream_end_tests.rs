//! Record 0769: open-time validation and every reader agree on the mini
//! streams at the end of a root mini stream whose size is not a multiple of
//! the 64-byte mini sector.
//!
//! MS-CFB 2.6.1 makes the root entry's stream size the size of the mini
//! stream and does not require a multiple of 64; MS-CFB 2.7 leaves the tail
//! of a stream's last sector unused. A mini stream is therefore admitted
//! exactly when every byte it takes from its mini sectors lies inside the
//! root size, and every reader then returns those bytes. A stream that needs
//! a byte past the root size is refused when the file is opened, with the
//! typed error open-time validation already gave mini storage outside the
//! root mini stream.
//!
//! Before this record, open admitted every stream whose mini sectors began
//! inside the root size, and the whole-stream readers (`OleFile::open_stream`,
//! the in-place readback comparison and the shared reader's mini-stream
//! cache) then refused a stream whose last sector the root size cut, even
//! when all of its bytes lay inside the root size.
//!
//! The expected verdicts come from an independent byte-bound oracle over the
//! crafted layouts (the least root size that holds every byte of every mini
//! stream), not from the validation code.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::SharedOleFile;
use crate::writer::{OleWriter, SectorLayoutFallback};
use litchi_core::SourceVersion;
use std::collections::BTreeMap;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use std::sync::Arc;

/// The refusal for mini storage outside the root mini stream.
const OUTSIDE: &str = "Mini stream references storage outside the root mini stream";
const MINI: usize = 64;
const MINI_U64: u64 = 64;

fn read_u32(bytes: &[u8], at: usize) -> u32 {
    u32::from_le_bytes(bytes[at..at + 4].try_into().unwrap())
}

fn write_u32(bytes: &mut [u8], at: usize, value: u32) {
    bytes[at..at + 4].copy_from_slice(&value.to_le_bytes());
}

fn index(value: u32) -> usize {
    usize::try_from(value).unwrap()
}

/// A small deterministic generator (xorshift64*), so every case replays.
struct Rng(u64);

impl Rng {
    fn new(seed: u64) -> Self {
        Self(seed.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1)
    }

    fn next(&mut self) -> u64 {
        let mut value = self.0;
        value ^= value >> 12;
        value ^= value << 25;
        value ^= value >> 27;
        self.0 = value;
        value.wrapping_mul(0x2545_F491_4F6C_DD1D)
    }

    fn below(&mut self, bound: usize) -> usize {
        usize::try_from(self.next() % u64::try_from(bound.max(1)).unwrap()).unwrap()
    }
}

/// A deterministic payload that differs by stream and by position.
fn payload(len: usize, salt: usize) -> Vec<u8> {
    (0..len)
        .map(|position| u8::try_from((position * 31 + salt * 17 + 1) % 251).unwrap())
        .collect()
}

/// A compound file of root-level streams, written from scratch.
fn written(sector_size: usize, streams: &[(String, Vec<u8>)]) -> Vec<u8> {
    let mut writer = OleWriter::with_sector_size(sector_size).unwrap();
    for (name, data) in streams {
        writer.create_stream(&[name.as_str()], data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

/// A mini stream: its path, SID, size and mini sectors in chain order.
#[derive(Clone, Debug)]
struct MiniStream {
    path: Vec<String>,
    sid: usize,
    size: u64,
    chain: Vec<u32>,
}

/// Where a valid compound file keeps the metadata these tests rewrite, taken
/// from a validated open of it.
struct Layout {
    sector_size: usize,
    fat: Vec<u32>,
    root_chain: Vec<u32>,
    first_dir_sector: u32,
    minifat_start: u32,
    root_size: u64,
    mini_streams: Vec<MiniStream>,
    /// Every stream's path and bytes, as the unmodified file reads.
    contents: Vec<(Vec<String>, Vec<u8>)>,
}

impl Layout {
    fn of(bytes: &[u8]) -> Self {
        let mut file = OleFile::open(Cursor::new(bytes)).unwrap();
        let mut mini_streams = Vec::new();
        let mut contents = Vec::new();
        for path in file.list_streams() {
            let refs: Vec<&str> = path.iter().map(String::as_str).collect();
            let entry = file.find_entry(&refs).unwrap().clone();
            if entry.is_minifat && entry.size > 0 {
                mini_streams.push(MiniStream {
                    path: path.clone(),
                    sid: index(entry.sid),
                    size: entry.size,
                    chain: collect_sector_chain(&file.minifat, entry.start_sector, "MiniFAT")
                        .unwrap(),
                });
            }
            let data = file.open_stream(&refs).unwrap();
            contents.push((path.clone(), data));
        }
        Self {
            sector_size: file.sector_size,
            fat: file.fat.clone(),
            root_chain: file.root_chain.clone(),
            first_dir_sector: file.first_dir_sector,
            minifat_start: read_u32(bytes, 0x3C),
            root_size: file.root.as_ref().unwrap().size,
            mini_streams,
            contents,
        }
    }

    /// The absolute offset of byte `offset` of the FAT chain from `start`.
    fn chain_offset(&self, start: u32, offset: usize) -> usize {
        let mut sector = start;
        for _ in 0..offset / self.sector_size {
            sector = self.fat[index(sector)];
        }
        (index(sector) + 1) * self.sector_size + offset % self.sector_size
    }

    fn entry_offset(&self, sid: usize) -> usize {
        self.chain_offset(self.first_dir_sector, sid * DIRENTRY_SIZE)
    }

    fn minifat_offset(&self, sector: u32) -> usize {
        self.chain_offset(self.minifat_start, index(sector) * 4)
    }

    fn ministream_offset(&self, offset: usize) -> usize {
        (index(self.root_chain[offset / self.sector_size]) + 1) * self.sector_size
            + offset % self.sector_size
    }

    /// Mini sectors of the unmodified mini stream.
    fn mini_sectors(&self) -> usize {
        usize::try_from(self.root_size.div_ceil(MINI_U64)).unwrap()
    }
}

/// Rewrites the root entry's stream size, the mini stream's size.
fn set_root_size(bytes: &mut [u8], layout: &Layout, size: u64) {
    let at = layout.entry_offset(0) + 0x78;
    bytes[at..at + 8].copy_from_slice(&size.to_le_bytes());
}

/// Moves every mini sector `i` of the mini stream to `perm[i]`, with its
/// bytes, its MiniFAT link and any stream start that names it, and returns
/// the mini streams as they are now chained.
fn permute_mini_sectors(bytes: &mut [u8], layout: &Layout, perm: &[u32]) -> Vec<MiniStream> {
    let count = perm.len();
    let blocks: Vec<Vec<u8>> = (0..count)
        .map(|sector| {
            let at = layout.ministream_offset(sector * MINI);
            bytes[at..at + MINI].to_vec()
        })
        .collect();
    let links: Vec<u32> = (0..count)
        .map(|sector| read_u32(bytes, layout.minifat_offset(u32::try_from(sector).unwrap())))
        .collect();
    for sector in 0..count {
        let at = layout.ministream_offset(index(perm[sector]) * MINI);
        bytes[at..at + MINI].copy_from_slice(&blocks[sector]);
        let link = links[sector];
        let moved = if index(link) < count {
            perm[index(link)]
        } else {
            link
        };
        write_u32(bytes, layout.minifat_offset(perm[sector]), moved);
    }
    layout
        .mini_streams
        .iter()
        .map(|stream| {
            let at = layout.entry_offset(stream.sid) + 0x74;
            write_u32(bytes, at, perm[index(stream.chain[0])]);
            MiniStream {
                chain: stream
                    .chain
                    .iter()
                    .map(|&sector| perm[index(sector)])
                    .collect(),
                ..stream.clone()
            }
        })
        .collect()
}

/// The byte-bound oracle: the least root size that holds every byte every
/// mini stream takes from its mini sectors.
fn required_root_size(streams: &[MiniStream]) -> u64 {
    streams
        .iter()
        .flat_map(|stream| {
            stream
                .chain
                .iter()
                .enumerate()
                .map(move |(position, &sector)| {
                    let taken =
                        (stream.size - MINI_U64 * u64::try_from(position).unwrap()).min(MINI_U64);
                    u64::from(sector) * MINI_U64 + taken
                })
        })
        .max()
        .unwrap_or(0)
}

/// Whether some mini stream takes bytes from the partial mini sector a root
/// size of `root_size` ends in.
fn uses_partial_sector(streams: &[MiniStream], root_size: u64) -> bool {
    root_size % MINI_U64 != 0
        && streams.iter().any(|stream| {
            stream
                .chain
                .contains(&u32::try_from(root_size / MINI_U64).unwrap())
        })
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord)]
enum Expected {
    /// Every byte of every mini stream lies inside the root size.
    Admitted,
    /// Some mini stream needs a byte past the root size.
    Outside,
    /// The root size needs another number of regular sectors than its chain
    /// has, so the root chain's own length check refuses first.
    RootChainLength,
}

fn oracle(layout: &Layout, streams: &[MiniStream], root_size: u64) -> Expected {
    let sector_size = u64::try_from(layout.sector_size).unwrap();
    if root_size.div_ceil(sector_size) != layout.root_size.div_ceil(sector_size) {
        Expected::RootChainLength
    } else if required_root_size(streams) <= root_size {
        Expected::Admitted
    } else {
        Expected::Outside
    }
}

/// A file's bytes as the in-place readback comparison's candidate.
struct FileBytes<'a>(&'a [u8]);

impl CandidateBytes for FileBytes<'_> {
    fn equals_at(&self, offset: u64, expected: &[u8]) -> Result<bool, OleError> {
        let start = usize::try_from(offset).unwrap();
        let bytes = self
            .0
            .get(start..start + expected.len())
            .ok_or_else(|| OleError::InvalidData("comparison leaves the file".to_string()))?;
        Ok(bytes == expected)
    }

    fn check_readable(&self, offset: u64, len: usize) -> Result<(), OleError> {
        let start = usize::try_from(offset).unwrap();
        self.0
            .get(start..start + len)
            .map(drop)
            .ok_or_else(|| OleError::InvalidData("comparison leaves the file".to_string()))
    }
}

fn shared(bytes: &[u8]) -> Result<SharedOleFile, OleError> {
    SharedOleFile::open_owned(
        Arc::from(bytes.to_vec().into_boxed_slice()),
        SourceVersion::new(769, 0),
    )
}

/// Every reader returns `expected` for every stream of `bytes`: the cursor
/// reader's whole read and ranges, the in-place readback comparison, and the
/// shared reader's bounded direct read, root mini-stream cache, ranges and
/// cursor.
fn assert_every_reader_returns(label: &str, bytes: &[u8], expected: &[(Vec<String>, Vec<u8>)]) {
    let mut ole =
        OleFile::open(Cursor::new(bytes)).unwrap_or_else(|error| panic!("{label}: {error}"));
    let compared = OleFile::open(Cursor::new(bytes)).unwrap();
    let candidate = FileBytes(bytes);
    let mut comparer = compared.stream_comparer(&candidate);
    let cached = shared(bytes).unwrap();
    for (path, data) in expected {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        let what = |reader: &str| format!("{label}: {path:?} through {reader}");
        assert_eq!(
            ole.open_stream(&refs).as_ref().ok(),
            Some(data),
            "{}",
            what("OleFile::open_stream")
        );
        let mut range = vec![0u8; data.len()];
        ole.read_stream_range(&refs, 0, &mut range)
            .unwrap_or_else(|error| panic!("{}: {error}", what("OleFile::read_stream_range")));
        assert_eq!(&range, data, "{}", what("OleFile::read_stream_range"));
        if let Some(&last) = data.last() {
            // The stream's last byte alone, the one nearest the root size.
            let mut one = [0u8];
            ole.read_stream_range(&refs, u64::try_from(data.len() - 1).unwrap(), &mut one)
                .unwrap();
            assert_eq!(one[0], last, "{}", what("a one-byte range"));
        }
        assert_eq!(
            comparer.stream_equals(&refs, data).ok(),
            Some(true),
            "{}",
            what("StreamComparer")
        );
        // A fresh shared reader takes the bounded direct path when the
        // stream is eligible; the kept one takes the root mini-stream cache.
        assert_eq!(
            shared(bytes).unwrap().open_stream(&refs).as_ref().ok(),
            Some(data),
            "{}",
            what("SharedOleFile::open_stream")
        );
        assert_eq!(
            cached.open_stream_force_cache(&refs).as_ref().ok(),
            Some(data),
            "{}",
            what("the shared mini-stream cache")
        );
        let mut shared_range = vec![0u8; data.len()];
        cached
            .read_stream_range(&refs, 0, &mut shared_range)
            .unwrap_or_else(|error| {
                panic!("{}: {error}", what("SharedOleFile::read_stream_range"))
            });
        assert_eq!(
            &shared_range,
            data,
            "{}",
            what("SharedOleFile::read_stream_range")
        );
        let mut via_cursor = vec![0u8; data.len()];
        cached
            .stream_cursor_at(&refs, 0)
            .and_then(|mut cursor| cursor.read_exact(&mut via_cursor))
            .unwrap_or_else(|error| panic!("{}: {error}", what("the shared cursor")));
        assert_eq!(&via_cursor, data, "{}", what("the shared cursor"));
    }
}

/// Both opens refuse `bytes` with the same message, `message` when given.
fn assert_refused(label: &str, bytes: &[u8], message: Option<&str>) {
    let cursor = OleFile::open(Cursor::new(bytes))
        .err()
        .unwrap_or_else(|| panic!("{label}: OleFile::open admitted it"))
        .to_string();
    let positional = shared(bytes)
        .err()
        .unwrap_or_else(|| panic!("{label}: SharedOleFile::open_owned admitted it"))
        .to_string();
    assert_eq!(cursor, positional, "{label}");
    if let Some(message) = message {
        assert_eq!(cursor, format!("Corrupted file: {message}"), "{label}");
    }
}

/// Republishes `expected` over the adopted source and reads it back. A
/// source whose mini stream is not a whole number of mini sectors is written
/// from scratch (the reused layout declines it), so the output's root size
/// is a multiple of 64 again.
fn assert_round_trip(label: &str, bytes: &[u8], expected: &[(Vec<String>, Vec<u8>)]) {
    let layout = Layout::of(bytes);
    let mut writer = OleWriter::with_sector_size(layout.sector_size).unwrap();
    assert!(writer.adopt_source_layout(bytes).unwrap(), "{label}");
    for (path, data) in expected {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        writer.create_stream(&refs, data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let report = writer.last_sector_layout().unwrap();
    if layout.root_size % MINI_U64 == 0 {
        assert!(report.reused_source_layout(), "{label}: {report:?}");
    } else {
        assert_eq!(
            report.fallback(),
            Some(SectorLayoutFallback::SourceMiniStreamUnaligned),
            "{label}"
        );
    }
    let output = output.into_inner();
    assert_eq!(Layout::of(&output).root_size % MINI_U64, 0, "{label}");
    assert_every_reader_returns(&format!("{label}, republished"), &output, expected);
}

/// Checks `bytes` with root size `root_size` against the oracle; returns the
/// expectation it met.
fn check_root_size(
    label: &str,
    bytes: &[u8],
    layout: &Layout,
    streams: &[MiniStream],
    contents: &[(Vec<String>, Vec<u8>)],
    root_size: u64,
) -> Expected {
    let mut crafted = bytes.to_vec();
    set_root_size(&mut crafted, layout, root_size);
    let label = format!("{label}, root size {root_size}");
    let verdict = oracle(layout, streams, root_size);
    match verdict {
        Expected::Admitted => assert_every_reader_returns(&label, &crafted, contents),
        Expected::Outside => assert_refused(&label, &crafted, Some(OUTSIDE)),
        Expected::RootChainLength => assert_refused(&label, &crafted, None),
    }
    verdict
}

/// Root sizes 64n-1, 64n and 64n+1 around the last two mini sectors, and the
/// byte-bound's own boundary minus one, at, and plus one.
fn interesting_root_sizes(layout: &Layout, streams: &[MiniStream]) -> Vec<u64> {
    let full = layout.root_size / MINI_U64;
    let bound = required_root_size(streams);
    let mut sizes = Vec::new();
    for sectors in [full.saturating_sub(1), full] {
        for delta in [-1i64, 0, 1] {
            sizes.push((sectors * MINI_U64).saturating_add_signed(delta));
        }
    }
    for delta in [-1i64, 0, 1] {
        sizes.push(bound.saturating_add_signed(delta));
    }
    sizes.retain(|&size| size > 0);
    sizes.sort_unstable();
    sizes.dedup();
    sizes
}

#[test]
fn root_sizes_around_a_partial_last_mini_sector_follow_the_byte_bound() {
    let mut reached = BTreeMap::<(Expected, bool), usize>::new();
    for sector_size in [SECTOR_SIZE_V3, SECTOR_SIZE_V4] {
        for last_len in [1usize, 2, 36, 62, 63, 64, 65, 100, 127, 128, 129, 4095] {
            let streams = vec![
                ("A".to_string(), payload(700, 1)),
                ("B".to_string(), payload(90, 2)),
                ("Z".to_string(), payload(last_len, 3)),
                ("Regular".to_string(), payload(5_000, 4)),
            ];
            let bytes = written(sector_size, &streams);
            let layout = Layout::of(&bytes);
            assert_eq!(layout.root_size % MINI_U64, 0, "the writer aligns its root");
            for root_size in interesting_root_sizes(&layout, &layout.mini_streams) {
                let label = format!("{sector_size}-byte sectors, last stream {last_len} bytes");
                let verdict = check_root_size(
                    &label,
                    &bytes,
                    &layout,
                    &layout.mini_streams,
                    &layout.contents,
                    root_size,
                );
                let partial = uses_partial_sector(&layout.mini_streams, root_size);
                *reached.entry((verdict, partial)).or_default() += 1;
                if verdict == Expected::Admitted && partial {
                    assert_round_trip(
                        &format!("{label}, root size {root_size}"),
                        &bytes_with_root(&bytes, &layout, root_size),
                        &layout.contents,
                    );
                }
            }
        }
    }
    eprintln!("0769 root-size matrix: {reached:?}");
    // The newly admitted shape (a stream's bytes end inside the partial
    // last mini sector) and the newly refused one (a stream needs bytes of
    // it past the root size, which open used to admit) both occur.
    assert!(
        reached
            .get(&(Expected::Admitted, true))
            .copied()
            .unwrap_or(0)
            >= 20,
        "{reached:?}"
    );
    assert!(
        reached
            .get(&(Expected::Outside, true))
            .copied()
            .unwrap_or(0)
            >= 10,
        "{reached:?}"
    );
    assert!(
        reached.contains_key(&(Expected::Admitted, false)),
        "{reached:?}"
    );
    assert!(
        reached.contains_key(&(Expected::RootChainLength, false))
            || reached.contains_key(&(Expected::RootChainLength, true)),
        "{reached:?}"
    );
}

fn bytes_with_root(bytes: &[u8], layout: &Layout, root_size: u64) -> Vec<u8> {
    let mut crafted = bytes.to_vec();
    set_root_size(&mut crafted, layout, root_size);
    crafted
}

#[test]
fn a_partial_mini_sector_inside_a_chain_is_refused_at_open() {
    // One mini stream of three sectors, moved so that its first sector, of
    // which it takes all 64 bytes, is the mini stream's last.
    for sector_size in [SECTOR_SIZE_V3, SECTOR_SIZE_V4] {
        let streams = vec![("S".to_string(), payload(150, 5))];
        let bytes = written(sector_size, &streams);
        let layout = Layout::of(&bytes);
        assert_eq!(layout.mini_sectors(), 3);
        let mut permuted = bytes.clone();
        let chained = permute_mini_sectors(&mut permuted, &layout, &[2, 0, 1]);
        assert_eq!(chained[0].chain, vec![2, 0, 1]);
        // Whole: every reader follows the permuted chain.
        assert_every_reader_returns("permuted", &permuted, &layout.contents);
        // The root size cuts the first sector of the chain: 30 of its 64
        // bytes lie inside. Open used to admit this and the reads refused.
        for root_size in [64 * 2 + 1, 64 * 2 + 22, 64 * 2 + 30, 64 * 3 - 1] {
            let mut crafted = permuted.clone();
            set_root_size(&mut crafted, &layout, root_size);
            assert_eq!(oracle(&layout, &chained, root_size), Expected::Outside);
            assert_refused(&format!("root size {root_size}"), &crafted, Some(OUTSIDE));
        }
        let mut whole = permuted.clone();
        set_root_size(&mut whole, &layout, 64 * 3);
        assert_every_reader_returns("root size 192", &whole, &layout.contents);
    }
}

#[test]
fn seeded_mini_layouts_follow_the_byte_bound() {
    let mut rng = Rng::new(0x0769_0001);
    let mut reached = BTreeMap::<(Expected, bool), usize>::new();
    for case in 0..160 {
        let sector_size = if rng.below(2) == 0 {
            SECTOR_SIZE_V3
        } else {
            SECTOR_SIZE_V4
        };
        let count = 1 + rng.below(6);
        let mut streams: Vec<(String, Vec<u8>)> = (0..count)
            .map(|stream| {
                let longest = if rng.below(4) == 0 { 4095 } else { 400 };
                let len = 1 + rng.below(longest);
                (format!("M{stream}"), payload(len, case * 8 + stream))
            })
            .collect();
        if rng.below(2) == 0 {
            streams.push((
                "Regular".to_string(),
                payload(4096 + rng.below(3_000), case),
            ));
        }
        let bytes = written(sector_size, &streams);
        let layout = Layout::of(&bytes);
        // Shuffle the mini sectors, so any stream, at any position of its
        // chain, can own the last one.
        let sectors = layout.mini_sectors();
        let mut perm: Vec<u32> = (0..u32::try_from(sectors).unwrap()).collect();
        for slot in (1..perm.len()).rev() {
            perm.swap(slot, rng.below(slot + 1));
        }
        let mut permuted = bytes.clone();
        let chained = permute_mini_sectors(&mut permuted, &layout, &perm);
        let full = layout.root_size;
        for _ in 0..4 {
            let low = full.saturating_sub(2 * MINI_U64).max(1);
            let root_size = low
                + u64::try_from(rng.below(usize::try_from(full + MINI_U64 - low).unwrap()))
                    .unwrap();
            let verdict = check_root_size(
                &format!("case {case}"),
                &permuted,
                &layout,
                &chained,
                &layout.contents,
                root_size,
            );
            *reached
                .entry((verdict, uses_partial_sector(&chained, root_size)))
                .or_default() += 1;
        }
    }
    eprintln!("0769 seeded layouts: {reached:?}");
    for key in [
        (Expected::Admitted, false),
        (Expected::Admitted, true),
        (Expected::Outside, true),
        (Expected::Outside, false),
    ] {
        assert!(
            reached.contains_key(&key),
            "{key:?} not reached: {reached:?}"
        );
    }
}

/// The offset of every directory entry of a file the validating open
/// refuses: its FAT from the header's sector list and its directory chain,
/// walked directly.
fn raw_directory_entries(bytes: &[u8], sector_size: usize) -> Vec<usize> {
    let fat_sectors = index(read_u32(bytes, 0x2C));
    assert!(fat_sectors <= 109, "the header lists every FAT sector");
    let mut fat = Vec::new();
    for slot in 0..fat_sectors {
        let at = (index(read_u32(bytes, 0x4C + 4 * slot)) + 1) * sector_size;
        fat.extend((0..sector_size / 4).map(|entry| read_u32(bytes, at + 4 * entry)));
    }
    let mut entries = Vec::new();
    let mut sector = read_u32(bytes, 0x30);
    while sector != ENDOFCHAIN {
        let at = (index(sector) + 1) * sector_size;
        entries.extend((0..sector_size / DIRENTRY_SIZE).map(|slot| at + slot * DIRENTRY_SIZE));
        sector = fat[index(sector)];
        assert!(
            entries.len() <= fat.len() * sector_size,
            "the directory chain ends"
        );
    }
    entries
}

fn repository_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(Path::parent)
        .unwrap()
        .to_path_buf()
}

/// POI's `BlockSize512.zvi`, a Zeiss AxioVision image written by that
/// producer, vendored unmodified. Its root mini stream is 11,564 bytes
/// (180 mini sectors and 44 bytes), and its last mini stream,
/// `\u{5}DocumentSummaryInformation` (172 bytes, mini sectors 178, 179 and
/// 180), ends exactly at the root size, 44 bytes into the partial sector.
const ZVI: &str = "test-data/poi/test-data/poifs/BlockSize512.zvi";

/// The stream's SHA-256, computed from the raw file by an independent
/// reader (record 0769's census script).
const ZVI_SUMMARY_SHA256: &str = "6f74c82ea0b1d6a867de8949a0e85b36a256de14c82ea429785a3f70c0efdbc9";

#[test]
fn a_real_producer_stream_ending_in_the_partial_mini_sector_reads() {
    let original = std::fs::read(repository_root().join(ZVI)).unwrap();
    assert_eq!(original.len(), 51_712);
    // The producer also writes FREESECT into its storages' start sector
    // fields, which the directory validation refuses before any stream
    // allocation is examined (a separate question this record does not
    // change). Normalizing those fields to ENDOFCHAIN, and nothing else,
    // lets the file reach the mini-stream checks.
    assert_refused(
        "unmodified",
        &original,
        Some("invalid CFB storage fields at SID 2"),
    );
    let mut normalized = original.clone();
    let mut storages = 0;
    for entry in raw_directory_entries(&original, SECTOR_SIZE_V3) {
        if normalized[entry + 0x42] == STGTY_STORAGE
            && read_u32(&normalized, entry + 0x74) == FREESECT
        {
            write_u32(&mut normalized, entry + 0x74, ENDOFCHAIN);
            storages += 1;
        }
    }
    assert_eq!(storages, 11);

    let layout = Layout::of(&normalized);
    assert_eq!(layout.root_size, 11_564);
    assert_eq!(layout.root_size % MINI_U64, 44);
    let summary = layout
        .mini_streams
        .iter()
        .find(|stream| {
            stream.path.last().map(String::as_str) == Some("\u{5}DocumentSummaryInformation")
        })
        .unwrap();
    assert_eq!(summary.size, 172);
    assert_eq!(summary.chain, vec![178, 179, 180]);
    assert_eq!(required_root_size(&layout.mini_streams), 11_564);
    let bytes = &layout
        .contents
        .iter()
        .find(|(path, _)| *path == summary.path)
        .unwrap()
        .1;
    assert_eq!(
        litchi_core::patch::BlobId::of(bytes).as_hex(),
        ZVI_SUMMARY_SHA256
    );

    assert_every_reader_returns("BlockSize512.zvi", &normalized, &layout.contents);
    // One byte less and the stream's last byte lies outside the root.
    assert_eq!(
        check_root_size(
            "BlockSize512.zvi",
            &normalized,
            &layout,
            &layout.mini_streams,
            &layout.contents,
            11_563
        ),
        Expected::Outside
    );
}

/// Every file under `dir` of 512 bytes to 48 MiB, in path order.
fn collect(dir: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(dir) else {
        return;
    };
    let mut sorted: Vec<PathBuf> = entries.flatten().map(|entry| entry.path()).collect();
    sorted.sort();
    for path in sorted {
        if path.is_dir() {
            collect(&path, out);
        } else if std::fs::metadata(&path)
            .is_ok_and(|metadata| (512..=48 * 1024 * 1024).contains(&metadata.len()))
        {
            out.push(path);
        }
    }
}

/// The compound files under `test-data/<dir>`, with their bytes.
fn compound_files(dir: &str) -> Vec<(String, Vec<u8>)> {
    let root = repository_root();
    let mut paths = Vec::new();
    collect(&root.join(dir), &mut paths);
    paths
        .into_iter()
        .filter_map(|path| {
            let bytes = std::fs::read(&path).ok()?;
            (bytes.get(..8) == Some(MAGIC.as_slice())).then(|| {
                let label = path
                    .strip_prefix(&root)
                    .unwrap_or(&path)
                    .display()
                    .to_string();
                (label, bytes)
            })
        })
        .collect()
}

/// The census: for every compound file under `test-data`, both opens reach
/// the same verdict, and every stream an admitting open lists is returned
/// by every reader. No file is admitted and then refused by a read.
#[test]
fn every_admitted_fixture_stream_reads_through_every_reader() {
    let (mut files, mut opened, mut streams) = (0usize, 0usize, 0usize);
    for (label, bytes) in compound_files("test-data") {
        files += 1;
        match OleFile::open(Cursor::new(bytes.as_slice())) {
            Err(error) => {
                assert_eq!(
                    shared(&bytes).err().map(|error| error.to_string()),
                    Some(error.to_string()),
                    "{label}"
                );
            },
            Ok(mut file) => {
                opened += 1;
                let mut contents = Vec::new();
                for path in file.list_streams() {
                    let refs: Vec<&str> = path.iter().map(String::as_str).collect();
                    let data = file.open_stream(&refs).unwrap_or_else(|error| {
                        panic!("{label}: {path:?} is admitted by the open but not read: {error}")
                    });
                    contents.push((path, data));
                }
                streams += contents.len();
                assert_every_reader_returns(&label, &bytes, &contents);
            },
        }
    }
    eprintln!("0769 census: {files} compound files, {opened} admitted, {streams} streams");
    assert!(
        files >= 200 && opened >= 200 && streams >= 1_200,
        "{files} / {opened} / {streams}"
    );
}

/// Root sizes around the end of every `test-data/ole` fixture's mini
/// streams (64n-1, 64n and 64n+1 for the last two mini sectors, and the
/// byte-bound minus one, at and plus one) follow the oracle, and every
/// admitted file reads the fixture's own bytes.
#[test]
fn root_sizes_around_each_fixture_mini_stream_end_follow_the_byte_bound() {
    let mut reached = BTreeMap::<(Expected, bool), usize>::new();
    let mut fixtures = 0usize;
    for (label, bytes) in compound_files("test-data/ole") {
        let layout = Layout::of(&bytes);
        if layout.mini_streams.is_empty() {
            continue;
        }
        fixtures += 1;
        for root_size in interesting_root_sizes(&layout, &layout.mini_streams) {
            let verdict = check_root_size(
                &label,
                &bytes,
                &layout,
                &layout.mini_streams,
                &layout.contents,
                root_size,
            );
            *reached
                .entry((
                    verdict,
                    uses_partial_sector(&layout.mini_streams, root_size),
                ))
                .or_default() += 1;
        }
    }
    eprintln!("0769 fixture root sizes over {fixtures} fixtures: {reached:?}");
    assert!(fixtures >= 60, "{fixtures}");
    for key in [
        (Expected::Admitted, false),
        (Expected::Admitted, true),
        (Expected::Outside, true),
    ] {
        assert!(
            reached.contains_key(&key),
            "{key:?} not reached: {reached:?}"
        );
    }
}
