//! Default-repair fixtures for the compound-file tolerance contract.
//!
//! Each defect class has a hand-built minimal file that carries the
//! deviation and must open by default with the repair recorded; the same
//! file without the deviation must open with a clean report.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]

use std::io::Cursor;
use std::sync::Arc;

use litchi_core::SourceVersion;

use crate::{
    AllocationDefect, OleError, OleFile, SharedOleFile, ToleranceReport, writer::OleWriter,
};

fn build(streams: &[(&[&str], &[u8])]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for (path, data) in streams {
        writer.create_stream(path, data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

/// Opens `bytes` through both readers and returns the cursor reader's report
/// after checking that the positional reader agrees.
fn open_report(bytes: &[u8]) -> ToleranceReport {
    let file = OleFile::open(Cursor::new(bytes.to_vec())).unwrap();
    let shared =
        SharedOleFile::open_owned(Arc::from(bytes.to_vec()), SourceVersion::new(0x7a, 0)).unwrap();
    assert_eq!(file.tolerance_report(), shared.tolerance_report());
    file.tolerance_report().clone()
}

#[test]
fn conformant_files_open_with_a_clean_report() {
    let small = build(&[(&["Data"], b"payload")]);
    let regular = build(&[
        (&["Data"], &[0x5a; 4096]),
        (&["Store", "Inner"], &[1; 9000]),
    ]);
    let empty = build(&[(&["Empty"], b"")]);
    for bytes in [small, regular, empty] {
        assert!(open_report(&bytes).is_clean());
    }
}

#[test]
fn real_world_compound_files_open_with_a_clean_report() {
    for bytes in [
        include_bytes!("../../../test-data/ole/doc/FancyFoot.doc").as_slice(),
        include_bytes!("../../../test-data/ole/doc/Lists.doc").as_slice(),
    ] {
        assert!(open_report(bytes).is_clean());
    }
}

// --- hand-built fixtures over writer output (512-byte sectors) ---

const SECTOR: usize = 512;

fn read_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn write_u32(bytes: &mut [u8], offset: usize, value: u32) {
    bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

fn sector_offset(sector: u32) -> usize {
    (sector as usize + 1) * SECTOR
}

/// The FAT entry of `sector`; fixtures stay within the first FAT sector.
fn fat_offset(bytes: &[u8], sector: u32) -> usize {
    assert!(
        sector < 128,
        "fixture FAT entry beyond the first FAT sector"
    );
    sector_offset(read_u32(bytes, 0x4c)) + sector as usize * 4
}

fn minifat_offset(bytes: &[u8], sector: u32) -> usize {
    assert!(
        sector < 128,
        "fixture MiniFAT entry beyond the first sector"
    );
    sector_offset(read_u32(bytes, 0x3c)) + sector as usize * 4
}

/// Every directory entry offset, in SID order, following the FAT.
fn directory_offsets(bytes: &[u8]) -> Vec<usize> {
    let mut offsets = Vec::new();
    let mut sector = read_u32(bytes, 0x30);
    while sector != crate::consts::ENDOFCHAIN {
        let base = sector_offset(sector);
        offsets.extend((0..SECTOR / 128).map(|index| base + index * 128));
        sector = read_u32(bytes, fat_offset(bytes, sector));
    }
    offsets
}

fn entry_name(bytes: &[u8], entry: usize) -> String {
    let len = usize::from(u16::from_le_bytes([bytes[entry + 64], bytes[entry + 65]]));
    let units: Vec<u16> = bytes[entry..entry + len.saturating_sub(2)]
        .chunks_exact(2)
        .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
        .collect();
    String::from_utf16(&units).unwrap()
}

/// The SID and entry offset of the entry named `name`.
fn entry(bytes: &[u8], name: &str) -> (u32, usize) {
    directory_offsets(bytes)
        .into_iter()
        .enumerate()
        .find(|&(_, offset)| bytes[offset + 66] != 0 && entry_name(bytes, offset) == name)
        .map(|(sid, offset)| (u32::try_from(sid).unwrap(), offset))
        .unwrap()
}

fn chain(bytes: &[u8], start: u32, count: usize, mini: bool) -> Vec<u32> {
    let mut sectors = vec![start];
    while sectors.len() < count {
        let last = *sectors.last().unwrap();
        let offset = if mini {
            minifat_offset(bytes, last)
        } else {
            fat_offset(bytes, last)
        };
        sectors.push(read_u32(bytes, offset));
    }
    sectors
}

/// Appends one physical sector of `fill` and returns its sector index.
fn append_sector(bytes: &mut Vec<u8>, fill: u8) -> u32 {
    assert_eq!(bytes.len() % SECTOR, 0);
    bytes.resize(bytes.len() + SECTOR, fill);
    u32::try_from(bytes.len() / SECTOR - 2).unwrap()
}

fn recorded(report: &ToleranceReport) -> Vec<(crate::Defect, u32)> {
    report
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context))
        .collect()
}

fn read_both(bytes: &[u8], path: &[&str]) -> Vec<u8> {
    let mut file = OleFile::open(Cursor::new(bytes.to_vec())).unwrap();
    let shared =
        SharedOleFile::open_owned(Arc::from(bytes.to_vec()), SourceVersion::new(0x7b, 0)).unwrap();
    let data = file.open_stream(path).unwrap();
    assert_eq!(shared.open_stream(path).unwrap(), data);
    let mut tail = vec![0; 3.min(data.len())];
    let offset = (data.len() - tail.len()) as u64;
    shared.read_stream_range(path, offset, &mut tail).unwrap();
    assert_eq!(tail, data[data.len() - tail.len()..]);
    data
}

fn pattern(len: usize, seed: u8) -> Vec<u8> {
    (0..len)
        .map(|index| (index % 251) as u8 ^ seed.wrapping_mul(37))
        .collect()
}

// --- F7.1: FAT chain longer than the declared stream size ---

#[test]
fn regular_stream_chain_longer_than_its_size_is_truncated_and_recorded() {
    let data = pattern(4096, 1);
    let mut bytes = build(&[(&["Data"], &data), (&["Small"], b"tiny")]);
    assert!(open_report(&bytes).is_clean());

    let (sid, offset) = entry(&bytes, "Data");
    let sectors = chain(&bytes, read_u32(&bytes, offset + 116), 8, false);
    let terminal = *sectors.last().unwrap();
    let orphan = append_sector(&mut bytes, 0xee);
    let orphan_entry = fat_offset(&bytes, orphan);
    write_u32(&mut bytes, orphan_entry, crate::consts::ENDOFCHAIN);
    let terminal_entry = fat_offset(&bytes, terminal);
    write_u32(&mut bytes, terminal_entry, orphan);

    assert_eq!(
        recorded(&open_report(&bytes)),
        [
            (AllocationDefect::ChainLongerThanSize.into(), sid),
            (AllocationDefect::UnclaimedAllocatedSector.into(), orphan),
        ]
    );
    assert_eq!(read_both(&bytes, &["Data"]), data);
    assert_eq!(read_both(&bytes, &["Small"]), b"tiny");
}

#[test]
fn mini_stream_chain_longer_than_its_size_is_truncated_and_recorded() {
    let data = pattern(128, 2);
    let mut bytes = build(&[(&["Data"], &data)]);
    assert!(open_report(&bytes).is_clean());

    let (sid, offset) = entry(&bytes, "Data");
    let sectors = chain(&bytes, read_u32(&bytes, offset + 116), 2, true);
    let terminal = *sectors.last().unwrap();
    let terminal_entry = minifat_offset(&bytes, terminal);
    write_u32(&mut bytes, terminal_entry, terminal + 1);

    assert_eq!(
        recorded(&open_report(&bytes)),
        [(AllocationDefect::ChainLongerThanSize.into(), sid)]
    );
    assert_eq!(read_both(&bytes, &["Data"]), data);
}

#[test]
fn root_mini_stream_chain_longer_than_its_size_is_truncated_and_recorded() {
    let data = pattern(100, 3);
    let mut bytes = build(&[(&["Data"], &data)]);
    let root = directory_offsets(&bytes)[0];
    let root_size = usize::try_from(read_u32(&bytes, root + 120)).unwrap();
    let sectors = chain(
        &bytes,
        read_u32(&bytes, root + 116),
        root_size.div_ceil(SECTOR),
        false,
    );
    let terminal = *sectors.last().unwrap();
    let orphan = append_sector(&mut bytes, 0);
    let orphan_entry = fat_offset(&bytes, orphan);
    write_u32(&mut bytes, orphan_entry, crate::consts::ENDOFCHAIN);
    let terminal_entry = fat_offset(&bytes, terminal);
    write_u32(&mut bytes, terminal_entry, orphan);

    assert_eq!(
        recorded(&open_report(&bytes)),
        [
            (AllocationDefect::ChainLongerThanSize.into(), 0),
            (AllocationDefect::UnclaimedAllocatedSector.into(), orphan),
        ]
    );
    assert_eq!(read_both(&bytes, &["Data"]), data);
}

#[test]
fn an_allocated_sector_no_chain_claims_is_recorded() {
    let mut bytes = build(&[(&["Data"], &pattern(1000, 4))]);
    let orphan = append_sector(&mut bytes, 0x55);
    let orphan_entry = fat_offset(&bytes, orphan);
    write_u32(&mut bytes, orphan_entry, crate::consts::ENDOFCHAIN);

    assert_eq!(
        recorded(&open_report(&bytes)),
        [(AllocationDefect::UnclaimedAllocatedSector.into(), orphan)]
    );
}

/// The surplus repair never relaxes the checks inside the declared length.
#[test]
fn chain_faults_within_the_declared_length_stay_fatal() {
    let bytes = build(&[(&["Data"], &pattern(4096, 5))]);
    let (_, offset) = entry(&bytes, "Data");
    let sectors = chain(&bytes, read_u32(&bytes, offset + 116), 8, false);
    for (link, expected) in [
        (sectors[1], "Cycle detected in regular stream chain"),
        (crate::consts::ENDOFCHAIN, "ends before its declared length"),
        (5000, "Invalid sector index 5000"),
        (crate::consts::FREESECT, "Invalid sector marker"),
    ] {
        let mut faulted = bytes.clone();
        let at = fat_offset(&faulted, sectors[3]);
        write_u32(&mut faulted, at, link);
        match OleFile::open(Cursor::new(faulted)) {
            Err(OleError::CorruptedFile(message)) => {
                assert!(message.contains(expected), "{expected}: {message}");
            },
            other => panic!("{expected}: expected a corruption error, got {other:?}"),
        }
    }
}

// --- F7.2: zero-size stream whose start sector is not ENDOFCHAIN ---

#[test]
fn empty_stream_with_a_stale_start_sector_reads_as_empty_and_is_recorded() {
    let data = pattern(5000, 6);
    let bytes = build(&[(&["Data"], &data), (&["Empty"], b"")]);
    assert!(open_report(&bytes).is_clean());

    let (sid, offset) = entry(&bytes, "Empty");
    assert_eq!(read_u32(&bytes, offset + 116), crate::consts::ENDOFCHAIN);
    // A sector another stream owns, a free marker, and a dangling index.
    let (_, data_offset) = entry(&bytes, "Data");
    let owned = read_u32(&bytes, data_offset + 116);
    for stale in [owned, crate::consts::FREESECT, 5000] {
        let mut faulted = bytes.clone();
        write_u32(&mut faulted, offset + 116, stale);
        assert_eq!(
            recorded(&open_report(&faulted)),
            [(AllocationDefect::EmptyStreamStartSector.into(), sid)],
            "stale start {stale:#x}"
        );
        assert!(read_both(&faulted, &["Empty"]).is_empty());
        assert_eq!(read_both(&faulted, &["Data"]), data);
        let file = OleFile::open(Cursor::new(faulted)).unwrap();
        let repaired = file.directory_entry_by_sid(sid).unwrap();
        assert_eq!(repaired.start_sector, crate::consts::ENDOFCHAIN);
    }
}

#[test]
fn empty_root_mini_stream_with_a_stale_start_sector_is_recorded() {
    let data = pattern(4096, 7);
    let mut bytes = build(&[(&["Data"], &data)]);
    let root = directory_offsets(&bytes)[0];
    assert_eq!(
        read_u32(&bytes, root + 120),
        0,
        "fixture has no mini stream"
    );
    assert!(open_report(&bytes).is_clean());

    write_u32(&mut bytes, root + 116, 0);
    assert_eq!(
        recorded(&open_report(&bytes)),
        [(AllocationDefect::EmptyStreamStartSector.into(), 0)]
    );
    assert_eq!(read_both(&bytes, &["Data"]), data);
}

// --- F7.3: red-black sibling-tree name-ordering violation ---

/// Swaps the left and right links of the sibling-tree node at `entry`.
fn swap_sibling_links(bytes: &mut [u8], entry: usize) {
    let left = read_u32(bytes, entry + 68);
    let right = read_u32(bytes, entry + 72);
    assert_ne!(left, right, "fixture node needs two distinct links");
    write_u32(bytes, entry + 68, right);
    write_u32(bytes, entry + 72, left);
}

fn sibling_names() -> Vec<String> {
    [
        "Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf",
    ]
    .map(str::to_owned)
    .to_vec()
}

#[test]
fn misordered_sibling_tree_opens_and_every_name_still_resolves() {
    let names = sibling_names();
    let contents: Vec<Vec<u8>> = (0..names.len())
        .map(|index| pattern(40 + index * 700, u8::try_from(index).unwrap()))
        .collect();
    let paths: Vec<[&str; 1]> = names.iter().map(|name| [name.as_str()]).collect();
    let nested = pattern(300, 9);
    let mut streams: Vec<(&[&str], &[u8])> = paths
        .iter()
        .zip(&contents)
        .map(|(path, data)| (path.as_slice(), data.as_slice()))
        .collect();
    let nested_paths: Vec<[&str; 2]> = names.iter().map(|name| ["Store", name.as_str()]).collect();
    for path in &nested_paths {
        streams.push((path.as_slice(), &nested));
    }
    let mut bytes = build(&streams);
    assert!(open_report(&bytes).is_clean());

    // Reverse the root storage's tree and the nested storage's tree at their
    // tree roots, so a descent by name goes the wrong way.
    let directory = directory_offsets(&bytes);
    let root_tree = read_u32(&bytes, directory[0] + 76);
    swap_sibling_links(&mut bytes, directory[root_tree as usize]);
    let (_, store) = entry(&bytes, "Store");
    let store_tree = read_u32(&bytes, store + 76);
    swap_sibling_links(&mut bytes, directory[store_tree as usize]);

    let report = open_report(&bytes);
    assert!(!report.is_clean());
    assert!(
        report
            .iter()
            .all(|tolerated| tolerated.defect == crate::DirectoryDefect::SiblingNameOrder.into())
    );
    for (path, data) in paths.iter().zip(&contents) {
        assert_eq!(&read_both(&bytes, path), data, "{path:?}");
    }
    for path in &nested_paths {
        assert_eq!(read_both(&bytes, path), nested, "{path:?}");
    }
    // Lookups are case-insensitive as before, and a missing name is still
    // not found by either reader.
    assert_eq!(read_both(&bytes, &["ALPHA"]), contents[0]);
    let mut file = OleFile::open(Cursor::new(bytes.clone())).unwrap();
    assert!(matches!(
        file.open_stream(&["Missing"]),
        Err(OleError::StreamNotFound)
    ));
    let shared = SharedOleFile::open_owned(Arc::from(bytes), SourceVersion::new(0x7c, 0)).unwrap();
    assert!(matches!(
        shared.open_stream(&["Store", "Missing"]),
        Err(OleError::StreamNotFound)
    ));
}

/// The ordering repair never relaxes acyclicity or unique ownership.
#[test]
fn sibling_tree_cycles_stay_fatal_when_ordering_is_tolerated() {
    let names = sibling_names();
    let paths: Vec<[&str; 1]> = names.iter().map(|name| [name.as_str()]).collect();
    let streams: Vec<(&[&str], &[u8])> = paths
        .iter()
        .map(|path| (path.as_slice(), b"x".as_slice()))
        .collect();
    let mut bytes = build(&streams);
    let directory = directory_offsets(&bytes);
    let root_tree = read_u32(&bytes, directory[0] + 76);
    // Point the tree root's left link back at itself.
    write_u32(&mut bytes, directory[root_tree as usize] + 68, root_tree);
    assert!(matches!(
        OleFile::open(Cursor::new(bytes)),
        Err(OleError::CorruptedFile(message)) if message.contains("repeated SID")
    ));
}

// --- F7.4: directory entries unreachable from the root ---

/// Unlinks the entry named `name` from its parent sibling tree, provided it
/// is a leaf there, and returns its SID.
fn detach_leaf(bytes: &mut [u8], name: &str) -> u32 {
    let (sid, offset) = entry(bytes, name);
    assert_eq!(read_u32(bytes, offset + 68), crate::consts::NOSTREAM);
    assert_eq!(read_u32(bytes, offset + 72), crate::consts::NOSTREAM);
    for link in directory_offsets(bytes)
        .into_iter()
        .filter(|&entry| bytes[entry + 66] != 0)
        .flat_map(|entry| [entry + 68, entry + 72, entry + 76])
    {
        if read_u32(bytes, link) == sid {
            write_u32(bytes, link, crate::consts::NOSTREAM);
            return sid;
        }
    }
    panic!("{name} has no parent link");
}

/// A sibling-tree leaf among `names` in `bytes`.
fn leaf_name<'a>(bytes: &[u8], names: &'a [String]) -> &'a str {
    names
        .iter()
        .find(|name| {
            let (_, offset) = entry(bytes, name);
            read_u32(bytes, offset + 68) == crate::consts::NOSTREAM
                && read_u32(bytes, offset + 72) == crate::consts::NOSTREAM
        })
        .unwrap()
}

#[test]
fn unreachable_directory_entry_is_ignored_and_recorded() {
    let names = sibling_names();
    let paths: Vec<[&str; 1]> = names.iter().map(|name| [name.as_str()]).collect();
    let contents: Vec<Vec<u8>> = (0..names.len())
        .map(|index| pattern(64 + index, u8::try_from(index).unwrap()))
        .collect();
    let streams: Vec<(&[&str], &[u8])> = paths
        .iter()
        .zip(&contents)
        .map(|(path, data)| (path.as_slice(), data.as_slice()))
        .collect();
    let mut bytes = build(&streams);
    assert!(open_report(&bytes).is_clean());

    let orphan_name = leaf_name(&bytes, &names).to_owned();
    let orphan = detach_leaf(&mut bytes, &orphan_name);
    assert_eq!(
        recorded(&open_report(&bytes)),
        [(crate::DirectoryDefect::UnreachableEntry.into(), orphan)]
    );

    let mut file = OleFile::open(Cursor::new(bytes.clone())).unwrap();
    assert!(file.directory_entry_by_sid(orphan).is_none());
    assert!(matches!(
        file.open_stream(&[orphan_name.as_str()]),
        Err(OleError::StreamNotFound)
    ));
    assert_eq!(file.list_streams().len(), names.len() - 1);
    for (path, data) in paths.iter().zip(&contents) {
        if path[0] != orphan_name {
            assert_eq!(&read_both(&bytes, path), data, "{path:?}");
        }
    }
}

#[test]
fn unreachable_storage_subtree_and_its_sectors_are_recorded() {
    let big = pattern(4096, 11);
    let mut bytes = build(&[
        (&["Keep"], b"kept"),
        (&["Store", "Inner"], &big),
        (&["Store", "Tiny"], b"tiny"),
    ]);
    assert!(open_report(&bytes).is_clean());
    let (inner, inner_offset) = entry(&bytes, "Inner");
    let (tiny, _) = entry(&bytes, "Tiny");
    let inner_sectors = chain(&bytes, read_u32(&bytes, inner_offset + 116), 8, false);

    // The root's sibling tree holds exactly Keep and Store. Make Keep the
    // tree root on its own, leaving Store (and its subtree) unreachable.
    let (store, store_offset) = entry(&bytes, "Store");
    let (keep, keep_offset) = entry(&bytes, "Keep");
    let root = directory_offsets(&bytes)[0];
    write_u32(&mut bytes, root + 76, keep);
    for offset in [keep_offset, store_offset] {
        write_u32(&mut bytes, offset + 68, crate::consts::NOSTREAM);
        write_u32(&mut bytes, offset + 72, crate::consts::NOSTREAM);
    }

    let report = open_report(&bytes);
    let mut unreachable: Vec<u32> = report
        .iter()
        .filter(|tolerated| tolerated.defect == crate::DirectoryDefect::UnreachableEntry.into())
        .map(|tolerated| tolerated.context)
        .collect();
    unreachable.sort_unstable();
    let mut expected = vec![store, inner, tiny];
    expected.sort_unstable();
    assert_eq!(unreachable, expected);
    let mut unclaimed: Vec<u32> = report
        .iter()
        .filter(|tolerated| tolerated.defect == AllocationDefect::UnclaimedAllocatedSector.into())
        .map(|tolerated| tolerated.context)
        .collect();
    unclaimed.sort_unstable();
    let mut inner_sorted = inner_sectors;
    inner_sorted.sort_unstable();
    assert_eq!(unclaimed, inner_sorted);
    assert_eq!(read_both(&bytes, &["Keep"]), b"kept");
}

// --- F7.6: CFB header reserved bytes must be zero ---

#[test]
fn nonzero_reserved_header_bytes_are_ignored_and_recorded() {
    let data = pattern(700, 12);
    let mut bytes = build(&[(&["Data"], &data)]);
    assert_eq!(bytes[0x22..0x28], [0; 6]);
    bytes[0x22..0x28].copy_from_slice(&[1, 2, 3, 4, 5, 6]);
    assert_eq!(
        recorded(&open_report(&bytes)),
        [(crate::HeaderDefect::ReservedBytes.into(), 0)]
    );
    assert_eq!(read_both(&bytes, &["Data"]), data);
}

// --- F7.5: MiniFAT/DIFAT start-sector vs count cross-check ---

#[test]
fn table_start_sectors_disagreeing_with_a_zero_count_are_ignored_and_recorded() {
    let data = pattern(5000, 13);
    let bytes = build(&[(&["Data"], &data)]);
    assert_eq!(read_u32(&bytes, 0x40), 0, "fixture has no MiniFAT");
    assert_eq!(read_u32(&bytes, 0x48), 0, "fixture has no DIFAT");
    assert!(open_report(&bytes).is_clean());

    // Zero-initializing writers store sector 0 as both start sectors.
    let mut zeroed = bytes.clone();
    write_u32(&mut zeroed, 0x3c, 0);
    write_u32(&mut zeroed, 0x44, 0);
    assert_eq!(
        recorded(&open_report(&zeroed)),
        [
            (crate::HeaderDefect::MiniFatStartCount.into(), 0),
            (crate::HeaderDefect::DifatStartCount.into(), 0),
        ]
    );
    assert_eq!(read_both(&zeroed, &["Data"]), data);

    // A nonzero MiniFAT count with no start sector reads as no MiniFAT.
    let mut counted = bytes;
    write_u32(&mut counted, 0x40, 1);
    assert_eq!(
        recorded(&open_report(&counted)),
        [(
            crate::HeaderDefect::MiniFatStartCount.into(),
            crate::consts::ENDOFCHAIN
        )]
    );
    assert_eq!(read_both(&counted, &["Data"]), data);
}

#[test]
fn table_start_count_repairs_never_invent_missing_tables() {
    // A mini stream still needs the MiniFAT the header no longer locates.
    let mut mini = build(&[(&["Data"], b"small")]);
    assert_ne!(read_u32(&mini, 0x40), 0);
    write_u32(&mut mini, 0x3c, crate::consts::ENDOFCHAIN);
    assert!(OleFile::open(Cursor::new(mini)).is_err());

    // A nonzero DIFAT count without a start sector stays fatal.
    let mut difat = build(&[(&["Data"], &pattern(5000, 14))]);
    write_u32(&mut difat, 0x48, 1);
    assert!(matches!(
        OleFile::open(Cursor::new(difat)),
        Err(OleError::InvalidFormat(message)) if message.contains("DIFAT start sector")
    ));
}

// --- F7.7: directory-entry field byte-exactness bundle ---

#[test]
fn meaningless_directory_entry_fields_are_ignored_and_recorded() {
    let data = pattern(300, 15);
    let bytes = build(&[
        (&["Data"], &data),
        (&["Store", "Inner"], b"inner"),
        (&["Extra"], b"extra"),
    ]);
    assert!(open_report(&bytes).is_clean());
    let (data_sid, data_entry) = entry(&bytes, "Data");
    let (store_sid, store_entry) = entry(&bytes, "Store");
    let (inner_sid, _) = entry(&bytes, "Inner");
    let root = directory_offsets(&bytes)[0];
    use crate::DirectoryDefect as D;

    // An undefined node color.
    let mut color = bytes.clone();
    color[data_entry + 67] = 7;
    assert_eq!(
        recorded(&open_report(&color)),
        [(D::NodeColor.into(), data_sid)]
    );

    // A storage carrying a start sector and size (a real producer writes
    // FREESECT there).
    let mut storage = bytes.clone();
    write_u32(&mut storage, store_entry + 116, crate::consts::FREESECT);
    write_u32(&mut storage, store_entry + 120, 77);
    assert_eq!(
        recorded(&open_report(&storage)),
        [(D::StorageAllocationFields.into(), store_sid)]
    );
    assert_eq!(read_both(&storage, &["Store", "Inner"]), b"inner");

    // A stream whose child link names another entry: never followed, so the
    // other entry keeps its one owner and no path runs through the stream.
    let mut child = bytes.clone();
    write_u32(&mut child, data_entry + 76, inner_sid);
    assert_eq!(
        recorded(&open_report(&child)),
        [(D::StreamChildLink.into(), data_sid)]
    );
    let mut file = OleFile::open(Cursor::new(child.clone())).unwrap();
    assert!(matches!(
        file.open_stream(&["Data", "Inner"]),
        Err(OleError::StreamNotFound)
    ));
    assert_eq!(read_both(&child, &["Data"]), data);
    assert_eq!(read_both(&child, &["Store", "Inner"]), b"inner");

    // A root entry with another name and stale sibling links.
    let mut root_fields = bytes.clone();
    root_fields[root..root + 64].fill(0);
    root_fields[root..root + 4].copy_from_slice(&[b'R', 0, 0, 0]);
    root_fields[root + 64..root + 66].copy_from_slice(&4u16.to_le_bytes());
    write_u32(&mut root_fields, root + 68, data_sid);
    assert_eq!(
        recorded(&open_report(&root_fields)),
        [(D::RootEntryFields.into(), 0)]
    );
    assert_eq!(read_both(&root_fields, &["Data"]), data);

    // An unused entry with leftover name bytes.
    let unused = directory_offsets(&bytes)
        .into_iter()
        .position(|offset| bytes[offset + 66] == 0)
        .expect("the fixture directory has an unused slot");
    let unused_offset = directory_offsets(&bytes)[unused];
    let mut residue = bytes.clone();
    residue[unused_offset + 64..unused_offset + 66].copy_from_slice(&10u16.to_le_bytes());
    residue[unused_offset..unused_offset + 8].copy_from_slice(b"o\0l\0d\0\0\0");
    assert_eq!(
        recorded(&open_report(&residue)),
        [(D::EmptyEntryResidue.into(), u32::try_from(unused).unwrap())]
    );
}

#[test]
fn names_with_forbidden_characters_stay_readable_and_are_recorded() {
    let data = pattern(200, 16);
    let mut bytes = build(&[(&["Data"], &data)]);
    let (sid, offset) = entry(&bytes, "Data");
    // "Data" -> "D:ta"; a lone stream has no siblings to reorder.
    bytes[offset + 2] = b':';
    assert_eq!(
        recorded(&open_report(&bytes)),
        [(crate::DirectoryDefect::NameCharacters.into(), sid)]
    );
    assert_eq!(read_both(&bytes, &["D:ta"]), data);
    assert_eq!(read_both(&bytes, &["d:TA"]), data);
    let file = OleFile::open(Cursor::new(bytes)).unwrap();
    assert_eq!(file.list_streams(), [vec!["D:ta".to_owned()]]);
}

/// The bundle never relaxes entry framing.
#[test]
fn directory_entry_framing_stays_fatal() {
    let bytes = build(&[(&["Data"], b"framing")]);
    let (_, offset) = entry(&bytes, "Data");
    for (field, value) in [(66usize, 9u8), (64, 3), (0, 0)] {
        let mut faulted = bytes.clone();
        faulted[offset + field] = value;
        if field == 0 {
            // Clear the terminator so the name is not NUL-terminated.
            faulted[offset + 8..offset + 10].copy_from_slice(b"x\0");
        }
        assert!(
            OleFile::open(Cursor::new(faulted)).is_err(),
            "field {field} = {value}"
        );
    }
}
