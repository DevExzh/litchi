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
