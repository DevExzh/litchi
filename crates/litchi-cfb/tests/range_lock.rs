#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused boundary tests use panic-on-failure assertions"
)]

//! End-to-end version-4 range-lock tests.
//!
//! The CFB files exercised here are larger than two GiB on the wire, but the
//! test source and sink stay small. `RepeatReader` supplies the stream payload
//! on demand, while `SparseCfb` stores only header/table/directory bytes and
//! synthesizes the known payload when the reader reopens the result.

use litchi_cfb::consts::{ENDOFCHAIN, RANGE_LOCK_SECTOR_V4, SECTOR_SIZE_V4, TWO_GIB_BYTES};
use litchi_cfb::writer::{SequentialOleWriter, SequentialWriterLimits, SequentialWriterOptions};
use litchi_cfb::{OleFile, OleFileLimits};
use std::cmp::min;
use std::collections::BTreeMap;
use std::io::{self, Read, Seek, SeekFrom, Write};
use std::ops::Range;

const PAYLOAD_BYTE: u8 = 0xA7;
const PUBLICATION_BUFFER_BYTES: usize = 64 * 1024;

#[derive(Clone, Debug)]
struct StoredRange {
    start: u64,
    bytes: Vec<u8>,
}

/// A bounded virtual file that retains metadata but synthesizes one known
/// repeated regular stream. The writer only moves forward, while the reader
/// reopens the same object through arbitrary seeks.
#[derive(Clone, Debug)]
struct SparseCfb {
    position: u64,
    length: u64,
    payload_ranges: Vec<Range<u64>>,
    range_lock: Range<u64>,
    payload_byte: u8,
    stored: Vec<StoredRange>,
    patches: BTreeMap<u64, Vec<u8>>,
    stored_bytes: u64,
    skipped_payload_bytes: u64,
    skipped_range_lock_bytes: u64,
}

impl SparseCfb {
    fn new(payload_len: u64, payload_byte: u8) -> Self {
        let sector_size = u64::try_from(SECTOR_SIZE_V4).unwrap();
        let data_start = sector_size;
        let lock_start = (u64::from(RANGE_LOCK_SECTOR_V4) + 1) * sector_size;
        let lock_end = lock_start + sector_size;
        let before_lock = lock_start - data_start;
        let mut payload_ranges = Vec::new();
        let first_len = min(payload_len, before_lock);
        if first_len != 0 {
            payload_ranges.push(data_start..data_start + first_len);
        }
        if payload_len > before_lock {
            let second_start = lock_end;
            payload_ranges.push(second_start..second_start + payload_len - before_lock);
        }
        Self {
            position: 0,
            length: 0,
            payload_ranges,
            range_lock: lock_start..lock_end,
            payload_byte,
            stored: Vec::new(),
            patches: BTreeMap::new(),
            stored_bytes: 0,
            skipped_payload_bytes: 0,
            skipped_range_lock_bytes: 0,
        }
    }

    fn stored_bytes(&self) -> u64 {
        self.stored_bytes
    }

    fn skipped_payload_bytes(&self) -> u64 {
        self.skipped_payload_bytes
    }

    fn skipped_range_lock_bytes(&self) -> u64 {
        self.skipped_range_lock_bytes
    }

    fn read_at(&self, start: u64, output: &mut [u8]) {
        output.fill(0);
        let end = start.saturating_add(u64::try_from(output.len()).unwrap_or(u64::MAX));

        for range in &self.payload_ranges {
            let overlap_start = start.max(range.start);
            let overlap_end = end.min(range.end);
            if overlap_start < overlap_end {
                let output_start = usize::try_from(overlap_start - start).unwrap();
                let output_end = usize::try_from(overlap_end - start).unwrap();
                output[output_start..output_end].fill(self.payload_byte);
            }
        }

        for range in &self.stored {
            let overlap_start = start.max(range.start);
            let overlap_end = end.min(range.start + u64::try_from(range.bytes.len()).unwrap());
            if overlap_start < overlap_end {
                let output_start = usize::try_from(overlap_start - start).unwrap();
                let range_start = usize::try_from(overlap_start - range.start).unwrap();
                let count = usize::try_from(overlap_end - overlap_start).unwrap();
                output[output_start..output_start + count]
                    .copy_from_slice(&range.bytes[range_start..range_start + count]);
            }
        }

        // Patches are applied last, allowing a test to mutate one FAT entry
        // without expanding the sparse representation.
        for (&patch_start, patch) in &self.patches {
            let patch_end = patch_start + u64::try_from(patch.len()).unwrap();
            let overlap_start = start.max(patch_start);
            let overlap_end = end.min(patch_end);
            if overlap_start < overlap_end {
                let output_start = usize::try_from(overlap_start - start).unwrap();
                let patch_offset = usize::try_from(overlap_start - patch_start).unwrap();
                let count = usize::try_from(overlap_end - overlap_start).unwrap();
                output[output_start..output_start + count]
                    .copy_from_slice(&patch[patch_offset..patch_offset + count]);
            }
        }
    }

    fn patch_at(&mut self, offset: u64, bytes: &[u8]) {
        assert!(offset.saturating_add(u64::try_from(bytes.len()).unwrap()) <= self.length);
        self.patches.insert(offset, bytes.to_vec());
    }

    fn read_u32_at(&self, offset: u64) -> u32 {
        let mut bytes = [0u8; 4];
        self.read_at(offset, &mut bytes);
        u32::from_le_bytes(bytes)
    }

    fn fat_sector_ids(&self) -> Vec<u32> {
        let sector_size = u64::try_from(SECTOR_SIZE_V4).unwrap();
        let fat_count = usize::try_from(self.read_u32_at(0x2C)).unwrap();
        let mut sectors = Vec::with_capacity(fat_count);
        let header_count = fat_count.min(109);
        for index in 0..header_count {
            sectors.push(self.read_u32_at(0x4C + u64::try_from(index * 4).unwrap()));
        }

        let mut difat_sector = self.read_u32_at(0x44);
        let difat_count = usize::try_from(self.read_u32_at(0x48)).unwrap();
        for _ in 0..difat_count {
            let sector_start = (u64::from(difat_sector) + 1) * sector_size;
            let entries = (SECTOR_SIZE_V4 / 4) - 1;
            for index in 0..entries {
                if sectors.len() == fat_count {
                    break;
                }
                sectors.push(self.read_u32_at(sector_start + u64::try_from(index * 4).unwrap()));
            }
            difat_sector = self.read_u32_at(sector_start + sector_size - 4);
        }
        assert_eq!(sectors.len(), fat_count);
        sectors
    }

    fn fat_entry_offset(&self, sector: u32) -> u64 {
        let entries_per_sector = SECTOR_SIZE_V4 / 4;
        let index = usize::try_from(sector).unwrap();
        let fat_sector = self.fat_sector_ids()[index / entries_per_sector];
        (u64::from(fat_sector) + 1) * u64::try_from(SECTOR_SIZE_V4).unwrap()
            + u64::try_from((index % entries_per_sector) * 4).unwrap()
    }

    fn fat_entry(&self, sector: u32) -> u32 {
        self.read_u32_at(self.fat_entry_offset(sector))
    }

    fn patch_fat_entry(&mut self, sector: u32, value: u32) {
        self.patch_at(self.fat_entry_offset(sector), &value.to_le_bytes());
    }

    fn assert_range_lock_is_zero(&self) {
        let mut bytes =
            vec![0xFF; usize::try_from(self.range_lock.end - self.range_lock.start).unwrap()];
        self.read_at(self.range_lock.start, &mut bytes);
        assert!(bytes.iter().all(|&byte| byte == 0));
    }
}

impl Read for SparseCfb {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.position >= self.length || output.is_empty() {
            return Ok(0);
        }
        let available = self.length - self.position;
        let count = min(available, u64::try_from(output.len()).unwrap())
            .try_into()
            .unwrap();
        self.read_at(self.position, &mut output[..count]);
        self.position += u64::try_from(count).unwrap();
        Ok(count)
    }
}

impl Seek for SparseCfb {
    fn seek(&mut self, from: SeekFrom) -> io::Result<u64> {
        let next = match from {
            SeekFrom::Start(offset) => i128::from(offset),
            SeekFrom::Current(offset) => i128::from(self.position) + i128::from(offset),
            SeekFrom::End(offset) => i128::from(self.length) + i128::from(offset),
        };
        if next < 0 || next > i128::from(u64::MAX) {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "virtual CFB seek is outside u64",
            ));
        }
        self.position = u64::try_from(next).unwrap();
        Ok(self.position)
    }
}

impl Write for SparseCfb {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let start = self.position;
        let end = start
            .checked_add(u64::try_from(bytes.len()).map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidInput,
                    "virtual CFB write is too large",
                )
            })?)
            .ok_or_else(|| {
                io::Error::new(io::ErrorKind::InvalidInput, "virtual CFB length overflow")
            })?;
        let mut offset = 0usize;
        while offset < bytes.len() {
            let current = start + u64::try_from(offset).unwrap();
            let mut next = end;
            enum Kind {
                Payload,
                RangeLock,
                Metadata,
            }
            let kind = if self.range_lock.contains(&current) {
                next = next.min(self.range_lock.end);
                Kind::RangeLock
            } else if let Some(range) = self
                .payload_ranges
                .iter()
                .find(|range| range.contains(&current))
            {
                next = next.min(range.end);
                Kind::Payload
            } else {
                for range in &self.payload_ranges {
                    if range.start > current {
                        next = next.min(range.start);
                    }
                }
                if self.range_lock.start > current {
                    next = next.min(self.range_lock.start);
                }
                Kind::Metadata
            };
            let count = usize::try_from(next - current).unwrap();
            let chunk = &bytes[offset..offset + count];
            match kind {
                Kind::Payload => {
                    if chunk.iter().any(|&byte| byte != self.payload_byte) {
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidData,
                            "regular stream bytes did not match the bounded test pattern",
                        ));
                    }
                    self.skipped_payload_bytes += u64::try_from(count).unwrap();
                },
                Kind::RangeLock => {
                    if chunk.iter().any(|&byte| byte != 0) {
                        return Err(io::Error::new(
                            io::ErrorKind::InvalidData,
                            "CFB range-lock sector received user data",
                        ));
                    }
                    self.skipped_range_lock_bytes += u64::try_from(count).unwrap();
                },
                Kind::Metadata => {
                    self.stored.push(StoredRange {
                        start: current,
                        bytes: chunk.to_vec(),
                    });
                    self.stored_bytes += u64::try_from(count).unwrap();
                },
            }
            offset += count;
        }
        self.position = end;
        self.length = self.length.max(end);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Clone, Debug)]
struct RepeatReader {
    remaining: u64,
    byte: u8,
}

impl Read for RepeatReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.remaining == 0 || output.is_empty() {
            return Ok(0);
        }
        let count = min(self.remaining, u64::try_from(output.len()).unwrap())
            .try_into()
            .unwrap();
        output[..count].fill(self.byte);
        self.remaining -= u64::try_from(count).unwrap();
        Ok(count)
    }
}

fn v4_options() -> SequentialWriterOptions {
    SequentialWriterOptions::default()
        .with_sector_size(SECTOR_SIZE_V4)
        .with_publication_buffer_bytes(PUBLICATION_BUFFER_BYTES)
        .with_limits(SequentialWriterLimits {
            max_stream_bytes: TWO_GIB_BYTES + 16 * 1024 * 1024,
            max_output_bytes: TWO_GIB_BYTES + 16 * 1024 * 1024,
            ..SequentialWriterLimits::default()
        })
}

fn publish_large_stream(length: u64) -> SparseCfb {
    let mut writer = SequentialOleWriter::with_options(v4_options()).unwrap();
    writer
        .add_stream(
            &["Boundary"],
            length,
            RepeatReader {
                remaining: length,
                byte: PAYLOAD_BYTE,
            },
        )
        .unwrap();
    let mut output = SparseCfb::new(length, PAYLOAD_BYTE);
    writer.write_to(&mut output).unwrap();
    assert!(output.stored_bytes() < 32 * 1024 * 1024);
    assert!(output.skipped_payload_bytes() >= length);
    output
}

fn reopen(output: SparseCfb) -> OleFile<SparseCfb> {
    let limits = OleFileLimits::new(output.length).unwrap();
    OleFile::open_with_limits(output, limits).unwrap()
}

fn assert_payload_range(file: &mut OleFile<SparseCfb>, offset: u64, length: usize) {
    let mut bytes = vec![0u8; length];
    file.read_stream_range(&["Boundary"], offset, &mut bytes)
        .unwrap();
    assert!(bytes.iter().all(|&byte| byte == PAYLOAD_BYTE));
}

#[test]
fn v4_metadata_crossing_uses_the_exact_fat_fixed_point() {
    // There is one directory sector for Root Entry plus Boundary. The two
    // adjacent logical-used counts exercise the fixed point immediately
    // before and at the extra-FAT-sector boundary. The regular payload ends
    // before the range-lock sector in both files; only metadata reaches it.
    for (stream_sectors, expected_fat_sectors) in [(523_773_u64, 512_u32), (523_774, 513)] {
        let length = stream_sectors * u64::try_from(SECTOR_SIZE_V4).unwrap();
        let output = publish_large_stream(length);
        assert!(output.length > TWO_GIB_BYTES);
        assert_eq!(
            output.skipped_range_lock_bytes(),
            u64::try_from(SECTOR_SIZE_V4).unwrap()
        );
        output.assert_range_lock_is_zero();
        assert_eq!(output.read_u32_at(0x2C), expected_fat_sectors);
        assert_eq!(output.read_u32_at(0x48), 1);
        assert_eq!(output.fat_entry(RANGE_LOCK_SECTOR_V4), ENDOFCHAIN);

        let mut file = reopen(output);
        assert_eq!(file.sector_size(), SECTOR_SIZE_V4);
        assert_eq!(file.stream_len(&["Boundary"]).unwrap(), length);
        assert_payload_range(&mut file, length - 8 * 1024, 8 * 1024);
    }
}

#[test]
fn v4_payload_chain_crosses_the_lock_and_rejects_an_inbound_lock_edge() {
    // The logical stream contains exactly 2 GiB. Its FAT chain has 524,286
    // sectors before the fixed hole and two sectors after it.
    let length = TWO_GIB_BYTES;
    let mut output = publish_large_stream(length);
    assert!(output.length > TWO_GIB_BYTES);
    assert_eq!(
        output.skipped_range_lock_bytes(),
        u64::try_from(SECTOR_SIZE_V4).unwrap()
    );
    output.assert_range_lock_is_zero();
    assert_eq!(output.fat_entry(RANGE_LOCK_SECTOR_V4), ENDOFCHAIN);

    let default_error =
        OleFile::open_with_limits(output.clone(), OleFileLimits::default()).unwrap_err();
    assert!(matches!(
        default_error,
        litchi_cfb::OleError::LimitExceeded {
            resource: "input bytes",
            observed,
            maximum,
        } if observed == output.length && maximum == TWO_GIB_BYTES
    ));
    let exact_length = output.length;
    let under_error = OleFile::open_with_limits(
        output.clone(),
        OleFileLimits::new(exact_length - 1).unwrap(),
    )
    .unwrap_err();
    assert!(matches!(
        under_error,
        litchi_cfb::OleError::LimitExceeded {
            resource: "input bytes",
            observed,
            maximum,
        } if observed == exact_length && maximum == exact_length - 1
    ));

    let mut file = reopen(output.clone());
    assert_eq!(file.stream_len(&["Boundary"]).unwrap(), length);
    let lock_start = (u64::from(RANGE_LOCK_SECTOR_V4) + 1) * u64::try_from(SECTOR_SIZE_V4).unwrap();
    let logical_before_lock = lock_start - u64::try_from(SECTOR_SIZE_V4).unwrap();
    assert_payload_range(&mut file, logical_before_lock - 8 * 1024, 16 * 1024);

    // The previous stream sector must skip the reserved sector. Pointing it
    // at the fixed hole is rejected before any user chain can consume it.
    output.patch_fat_entry(RANGE_LOCK_SECTOR_V4 - 1, RANGE_LOCK_SECTOR_V4);
    let error =
        OleFile::open_with_limits(output, OleFileLimits::new(exact_length).unwrap()).unwrap_err();
    assert!(error.to_string().contains("range-lock"));
}
