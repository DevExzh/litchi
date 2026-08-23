#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use std::io::{self, Cursor, Read, Seek, SeekFrom};
use std::sync::{
    Arc,
    atomic::{AtomicUsize, Ordering},
};

use crate::consts::{
    DIRENTRY_SIZE, ENDOFCHAIN, FREESECT, HEADER_DIFAT_ENTRIES, HEADER_DIFAT_OFFSET, MAXREGSECT,
    NUM_FAT_SECTORS_OFFSET, SECTOR_SHIFT_OFFSET, SECTOR_SHIFT_V3, SECTOR_SIZE_V3,
};
use crate::{OleError, OleFile, OleFileLimits, writer::OleWriter};

fn sample_file() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Data"], b"payload").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn mini_stream_file() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Data"], &[0u8; 128]).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn regular_stream_file() -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Data"], &[0u8; 4096]).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn read_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn write_u32(bytes: &mut [u8], offset: usize, value: u32) {
    bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

fn data_entry_offset(bytes: &[u8]) -> usize {
    (read_u32(bytes, 0x30) as usize + 1) * SECTOR_SIZE_V3 + DIRENTRY_SIZE
}

fn data_start_sector(bytes: &[u8]) -> u32 {
    read_u32(bytes, data_entry_offset(bytes) + 116)
}

fn fat_entry_offset(bytes: &[u8], sector: u32) -> usize {
    (read_u32(bytes, 0x4c) as usize + 1) * SECTOR_SIZE_V3 + sector as usize * 4
}

fn minifat_entry_offset(bytes: &[u8], sector: u32) -> usize {
    (read_u32(bytes, 0x3c) as usize + 1) * SECTOR_SIZE_V3 + sector as usize * 4
}

fn regular_data_terminal_sector(bytes: &[u8]) -> u32 {
    let mut sector = data_start_sector(bytes);
    for _ in 1..8 {
        sector = read_u32(bytes, fat_entry_offset(bytes, sector));
    }
    sector
}

fn mini_data_terminal_sector(bytes: &[u8]) -> u32 {
    let mut sector = data_start_sector(bytes);
    sector = read_u32(bytes, minifat_entry_offset(bytes, sector));
    sector
}

fn assert_open_corrupted(bytes: Vec<u8>, expected: &str) {
    assert!(matches!(
        OleFile::open(Cursor::new(bytes)),
        Err(OleError::CorruptedFile(message)) if message.contains(expected)
    ));
}

#[derive(Debug)]
struct ReportedLengthReader {
    bytes: Vec<u8>,
    position: u64,
    reported_length: u64,
    reads: Arc<AtomicUsize>,
}

impl Read for ReportedLengthReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::Relaxed);
        let Some(start) = usize::try_from(self.position).ok() else {
            return Ok(0);
        };
        let Some(input) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        self.position = self
            .position
            .checked_add(u64::try_from(count).unwrap())
            .unwrap();
        Ok(count)
    }
}

impl Seek for ReportedLengthReader {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let next = match position {
            SeekFrom::Start(offset) => offset,
            SeekFrom::End(offset) => self
                .reported_length
                .checked_add_signed(offset)
                .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "seek overflow"))?,
            SeekFrom::Current(offset) => self
                .position
                .checked_add_signed(offset)
                .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "seek overflow"))?,
        };
        self.position = next;
        Ok(next)
    }
}

#[test]
fn rejects_version_3_directory_sector_count() {
    let mut bytes = sample_file();
    write_u32(&mut bytes, 0x28, 1);
    assert!(OleFile::open(Cursor::new(bytes)).is_err());
}

#[test]
fn low_level_default_ingress_limit_is_finite() {
    assert!(OleFileLimits::default().max_input_bytes() < u64::MAX);
    assert!(matches!(
        OleFileLimits::new(0),
        Err(OleError::InvalidLimit {
            resource: "CFB input bytes",
            value: 0,
            maximum: OleFileLimits::MAX_INPUT_BYTES,
        })
    ));
    assert!(matches!(
        OleFileLimits::new(OleFileLimits::MAX_INPUT_BYTES + 1),
        Err(OleError::InvalidLimit {
            resource: "CFB input bytes",
            value,
            maximum,
        }) if value == maximum + 1
    ));
}

#[test]
fn low_level_ingress_accepts_exact_limit_and_rejects_max_plus_one() {
    let bytes = sample_file();
    let exact = OleFileLimits::new(bytes.len() as u64).unwrap();
    OleFile::open_with_limits(Cursor::new(bytes.clone()), exact)
        .expect("a source exactly at its configured limit opens");

    let over_limit = OleFileLimits::new(bytes.len() as u64).unwrap();
    let mut oversized = bytes;
    oversized.push(0);
    assert!(matches!(
        OleFile::open_with_limits(Cursor::new(oversized), over_limit),
        Err(OleError::LimitExceeded {
            resource: "input bytes",
            observed,
            maximum,
        }) if observed == maximum + 1
    ));
}

#[test]
fn hostile_reported_length_is_rejected_before_header_read() {
    let bytes = sample_file();
    let reads = Arc::new(AtomicUsize::new(0));
    let reader = ReportedLengthReader {
        reported_length: bytes.len() as u64 + 1,
        bytes,
        position: 0,
        reads: reads.clone(),
    };
    let limits = OleFileLimits::new(reader.reported_length - 1).unwrap();

    assert!(matches!(
        OleFile::open_with_limits(reader, limits),
        Err(OleError::LimitExceeded {
            resource: "input bytes",
            observed,
            maximum,
        }) if observed == maximum + 1
    ));
    assert_eq!(reads.load(Ordering::Relaxed), 0);
}

#[test]
fn hostile_count_metadata_is_rejected_before_index_allocation() {
    let mut bytes = sample_file();
    write_u32(&mut bytes, NUM_FAT_SECTORS_OFFSET, u32::MAX);

    assert!(matches!(
        OleFile::open(Cursor::new(bytes)),
        Err(OleError::CorruptedFile(message))
            if message.contains("Declared FAT sector count exceeds the physical file")
    ));
}

#[test]
fn hostile_difat_and_minifat_counts_are_rejected_before_index_allocation() {
    for (count_offset, first_offset, message) in [
        (
            0x48,
            0x44,
            "Declared DIFAT sector count exceeds the physical file",
        ),
        (
            0x40,
            0x3C,
            "Declared MiniFAT sector count exceeds the physical file",
        ),
    ] {
        let mut bytes = sample_file();
        write_u32(&mut bytes, count_offset, u32::MAX);
        write_u32(&mut bytes, first_offset, 0);

        assert!(matches!(
            OleFile::open(Cursor::new(bytes)),
            Err(OleError::CorruptedFile(actual)) if actual == message
        ));
    }
}

#[test]
fn hostile_version_4_directory_count_is_rejected_before_index_allocation() {
    let mut writer = OleWriter::with_sector_size(crate::consts::SECTOR_SIZE_V4).unwrap();
    writer.create_stream(&["Payload"], b"bounded").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let mut bytes = output.into_inner();
    write_u32(&mut bytes, 0x28, u32::MAX);

    assert!(matches!(
        OleFile::open(Cursor::new(bytes)),
        Err(OleError::CorruptedFile(message))
            if message == "Declared directory sector count exceeds the physical file"
    ));
}

#[test]
fn rejects_difat_start_count_disagreement() {
    let mut bytes = sample_file();
    write_u32(&mut bytes, 0x48, 1);
    write_u32(&mut bytes, 0x44, 0xFFFF_FFFE);
    assert!(OleFile::open(Cursor::new(bytes)).is_err());
}

#[test]
fn rejects_directory_and_fat_sector_overlap() {
    let mut bytes = sample_file();
    let first_fat_sector = read_u32(&bytes, 0x4C);
    write_u32(&mut bytes, 0x30, first_fat_sector);
    assert!(OleFile::open(Cursor::new(bytes)).is_err());
}

#[test]
fn rejects_fat_sector_without_fatsect_marker() {
    let mut bytes = sample_file();
    let sector_size = 1usize << u16::from_le_bytes([bytes[0x1E], bytes[0x1F]]);
    let first_fat_sector = read_u32(&bytes, 0x4C) as usize;
    let fat_offset = (first_fat_sector + 1) * sector_size + first_fat_sector * 4;
    write_u32(&mut bytes, fat_offset, 0xFFFF_FFFF);
    assert!(OleFile::open(Cursor::new(bytes)).is_err());
}

#[test]
fn rejects_incorrect_minifat_sector_count() {
    let mut bytes = sample_file();
    let count = read_u32(&bytes, 0x40);
    assert!(count > 0, "sample must contain a MiniFAT");
    write_u32(&mut bytes, 0x40, count + 1);
    assert!(OleFile::open(Cursor::new(bytes)).is_err());
}

#[test]
fn open_rejects_minifat_cycles_and_invalid_markers() {
    fn corrupt_first_data_mini_sector(bytes: &mut [u8], next: u32) {
        write_u32(
            bytes,
            minifat_entry_offset(bytes, data_start_sector(bytes)),
            next,
        );
    }

    let mut cycle = mini_stream_file();
    let mini_start = data_start_sector(&cycle);
    corrupt_first_data_mini_sector(&mut cycle, mini_start);
    assert!(matches!(
        OleFile::open(Cursor::new(cycle)),
        Err(OleError::CorruptedFile(message)) if message.contains("Cycle detected")
    ));

    let mut invalid_marker = mini_stream_file();
    corrupt_first_data_mini_sector(&mut invalid_marker, MAXREGSECT);
    assert!(matches!(
        OleFile::open(Cursor::new(invalid_marker)),
        Err(OleError::CorruptedFile(message)) if message.contains("Invalid sector marker")
    ));
}

#[test]
fn open_rejects_fat_stream_cycles_and_invalid_markers() {
    fn corrupt_first_data_sector(bytes: &mut [u8], next: u32) {
        write_u32(
            bytes,
            fat_entry_offset(bytes, data_start_sector(bytes)),
            next,
        );
    }

    let mut cycle = regular_stream_file();
    let data_start = data_start_sector(&cycle);
    corrupt_first_data_sector(&mut cycle, data_start);
    assert!(matches!(
        OleFile::open(Cursor::new(cycle)),
        Err(OleError::CorruptedFile(message)) if message.contains("Cycle detected")
    ));

    let mut invalid_marker = regular_stream_file();
    corrupt_first_data_sector(&mut invalid_marker, MAXREGSECT);
    assert!(matches!(
        OleFile::open(Cursor::new(invalid_marker)),
        Err(OleError::CorruptedFile(message)) if message.contains("Invalid sector marker")
    ));
}

#[test]
fn open_rejects_fat_stream_short_excess_and_terminal_markers() {
    let mut short = regular_stream_file();
    let start = data_start_sector(&short);
    let offset = fat_entry_offset(&short, start);
    write_u32(&mut short, offset, ENDOFCHAIN);
    assert_open_corrupted(
        short,
        "regular stream chain ends before its declared length",
    );

    let mut excess = regular_stream_file();
    let terminal = regular_data_terminal_sector(&excess);
    let offset = fat_entry_offset(&excess, terminal);
    write_u32(&mut excess, offset, terminal + 1);
    assert_open_corrupted(excess, "regular stream chain exceeds its declared length");

    let mut free = regular_stream_file();
    let terminal = regular_data_terminal_sector(&free);
    let offset = fat_entry_offset(&free, terminal);
    write_u32(&mut free, offset, FREESECT);
    assert_open_corrupted(free, "regular stream chain exceeds its declared length");

    let mut explicit_terminal = regular_stream_file();
    let terminal = regular_data_terminal_sector(&explicit_terminal);
    let offset = fat_entry_offset(&explicit_terminal, terminal);
    write_u32(&mut explicit_terminal, offset, ENDOFCHAIN);
    assert!(OleFile::open(Cursor::new(explicit_terminal)).is_ok());
}

#[test]
fn open_rejects_minifat_stream_short_excess_and_terminal_markers() {
    let mut short = mini_stream_file();
    let start = data_start_sector(&short);
    let offset = minifat_entry_offset(&short, start);
    write_u32(&mut short, offset, ENDOFCHAIN);
    assert_open_corrupted(short, "mini stream chain ends before its declared length");

    let mut excess = mini_stream_file();
    let terminal = mini_data_terminal_sector(&excess);
    let offset = minifat_entry_offset(&excess, terminal);
    write_u32(&mut excess, offset, terminal + 1);
    assert_open_corrupted(excess, "mini stream chain exceeds its declared length");

    let mut free = mini_stream_file();
    let terminal = mini_data_terminal_sector(&free);
    let offset = minifat_entry_offset(&free, terminal);
    write_u32(&mut free, offset, FREESECT);
    assert_open_corrupted(free, "mini stream chain exceeds its declared length");

    let mut explicit_terminal = mini_stream_file();
    let terminal = mini_data_terminal_sector(&explicit_terminal);
    let offset = minifat_entry_offset(&explicit_terminal, terminal);
    write_u32(&mut explicit_terminal, offset, ENDOFCHAIN);
    assert!(OleFile::open(Cursor::new(explicit_terminal)).is_ok());
}

#[test]
fn opens_real_world_word_compound_files() {
    for (name, bytes) in [
        (
            "FancyFoot.doc",
            include_bytes!("../../../test-data/ole/doc/FancyFoot.doc").as_slice(),
        ),
        (
            "Lists.doc",
            include_bytes!("../../../test-data/ole/doc/Lists.doc").as_slice(),
        ),
    ] {
        OleFile::open(Cursor::new(bytes))
            .unwrap_or_else(|error| panic!("failed to open {name}: {error}"));
    }
}

#[test]
fn opens_the_complete_legacy_office_compound_file_corpus() {
    fn collect(dir: &std::path::Path, files: &mut Vec<std::path::PathBuf>) {
        for entry in std::fs::read_dir(dir)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", dir.display()))
        {
            let path = entry.unwrap().path();
            if path.is_dir() {
                collect(&path, files);
            } else if path
                .extension()
                .and_then(|extension| extension.to_str())
                .is_some_and(|extension| {
                    extension.eq_ignore_ascii_case("doc")
                        || extension.eq_ignore_ascii_case("xls")
                        || extension.eq_ignore_ascii_case("ppt")
                })
            {
                files.push(path);
            }
        }
    }

    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ole");
    let mut files = Vec::new();
    collect(&root, &mut files);
    files.sort();
    assert!(!files.is_empty(), "legacy Office fixture corpus is empty");

    for path in files {
        let bytes = std::fs::read(&path)
            .unwrap_or_else(|error| panic!("failed to read {}: {error}", path.display()));
        OleFile::open(Cursor::new(bytes))
            .unwrap_or_else(|error| panic!("failed to open {}: {error}", path.display()));
    }
}

/// MS-CFB 2.6.1 notes that older writers left the most significant 32 bits of a
/// version 3 stream size uninitialized, and recommends that parsers ignore
/// those bits rather than reject the file. This fixture carries such a
/// directory entry and previously failed to open.
#[test]
fn opens_version_3_files_with_an_uninitialized_stream_size_high_word() {
    const FIXTURE: &[u8] =
        include_bytes!("../../../test-data/ole/doc/cfb-v3-uninitialized-size-high-word.doc");

    // Guard the fixture's premise: a 512-byte-sector file with a directory
    // entry whose high stream-size word is nonzero.
    assert_eq!(
        FIXTURE[SECTOR_SHIFT_OFFSET], SECTOR_SHIFT_V3,
        "fixture must use version 3 512-byte sectors"
    );
    let has_high_word = FIXTURE[SECTOR_SIZE_V3..]
        .chunks_exact(DIRENTRY_SIZE)
        .any(|entry| entry[DIRENTRY_SIZE - 4..].iter().any(|&byte| byte != 0));
    assert!(
        has_high_word,
        "fixture must contain a nonzero high stream-size word"
    );

    let file = OleFile::open(Cursor::new(FIXTURE)).expect("uninitialized high word is ignored");
    let entries = file.list_streams();
    assert!(
        entries
            .iter()
            .any(|path| path.iter().any(|part| part == "WordDocument")),
        "expected the Word streams to be reachable, got {entries:?}"
    );
}

/// MS-CFB imposes no requirement that a compound file's length be a whole
/// number of sectors, and real documents are commonly stored truncated at the
/// end of their last used sector. Such a file must open, with the short final
/// sector reading as zeroes past the end.
#[test]
fn opens_files_whose_length_is_not_a_whole_number_of_sectors() {
    const FIXTURE: &[u8] =
        include_bytes!("../../../test-data/ole/doc/cfb-truncated-final-sector.doc");

    assert_ne!(
        FIXTURE.len() % SECTOR_SIZE_V3,
        0,
        "fixture must end mid-sector"
    );

    let file = OleFile::open(Cursor::new(FIXTURE)).expect("a short final sector is tolerated");
    let entries = file.list_streams();
    assert!(
        entries
            .iter()
            .any(|path| path.iter().any(|part| part == "WordDocument")),
        "expected the Word streams to be reachable, got {entries:?}"
    );
}

/// A sector that begins at or beyond the end of the file is still an error:
/// tolerating a short final sector must not mask a genuinely missing one.
#[test]
fn still_rejects_sectors_that_start_past_the_end_of_the_file() {
    let bytes = sample_file();
    let truncated = &bytes[..SECTOR_SIZE_V3];
    assert!(
        OleFile::open(Cursor::new(truncated.to_vec())).is_err(),
        "a header-only file has no sectors to read"
    );
}

/// MS-CFB 2.2 describes the header DIFAT only as holding "the first 109 FAT
/// sector locations" and never constrains the entries past the declared FAT
/// sector count. Writers leave zeroes or stale values there, so the tail must
/// not be validated — the count field already says where the list ends.
#[test]
fn ignores_the_unused_tail_of_the_header_difat() {
    let mut bytes = sample_file();
    let used = read_u32(&bytes, NUM_FAT_SECTORS_OFFSET) as usize;
    assert!(
        used < HEADER_DIFAT_ENTRIES,
        "sample must leave part of the header DIFAT unused"
    );

    // Dirty every unused entry with values a writer might plausibly leave.
    for (index, value) in
        (used..HEADER_DIFAT_ENTRIES).zip([0u32, 1, 0xDEAD_BEEF].into_iter().cycle())
    {
        write_u32(&mut bytes, HEADER_DIFAT_OFFSET + index * 4, value);
    }

    let file = OleFile::open(Cursor::new(bytes)).expect("a dirty DIFAT tail is not fatal");
    let entries = file.list_streams();
    assert!(
        entries
            .iter()
            .any(|path| path.iter().any(|part| part == "Data")),
        "expected the stream to remain reachable, got {entries:?}"
    );
}

/// The used part of the list is still validated: an entry inside the declared
/// count that terminates the list early remains an error.
#[test]
fn still_rejects_a_header_difat_list_shorter_than_its_count() {
    let mut bytes = sample_file();
    let used = read_u32(&bytes, NUM_FAT_SECTORS_OFFSET) as usize;
    assert!(used > 0, "sample must declare at least one FAT sector");

    write_u32(&mut bytes, HEADER_DIFAT_OFFSET + (used - 1) * 4, FREESECT);
    assert!(OleFile::open(Cursor::new(bytes)).is_err());
}
