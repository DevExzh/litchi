//! `OleFile::stream_equals` held to `OleFile::open_stream` (change 0749).
//!
//! The in-place comparison must reach the verdict `open_stream` followed by a
//! byte comparison reaches — equal, different, or the same error — for every
//! stream of every OLE2 fixture, for expectations that differ in one byte or
//! in length, for files whose final sector is short, and for files whose
//! allocation tables were corrupted after they were written.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]

use std::io::Cursor;
use std::path::{Path, PathBuf};

use crate::file::{CandidateBytes, StreamCompareScratch};
use crate::{OleError, OleFile, writer::OleWriter};

const MAGIC: &[u8; 8] = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1";
const MAX_FIXTURE_BYTES: u64 = 48 * 1024 * 1024;

/// A file's bytes, examined where they lie.
struct FileBytes<'a>(&'a [u8]);

impl FileBytes<'_> {
    fn range(&self, offset: u64, len: usize) -> Result<&[u8], OleError> {
        let start = usize::try_from(offset).unwrap();
        self.0
            .get(start..start + len)
            .ok_or_else(|| OleError::InvalidData("comparison leaves the file".to_string()))
    }
}

impl CandidateBytes for FileBytes<'_> {
    fn equals_at(&self, offset: u64, expected: &[u8]) -> Result<bool, OleError> {
        Ok(self.range(offset, expected.len())? == expected)
    }

    fn check_readable(&self, offset: u64, len: usize) -> Result<(), OleError> {
        self.range(offset, len).map(drop)
    }
}

type Verdict = Result<bool, String>;

/// The readback verdict and the in-place verdict for one expectation.
fn verdicts(
    file: &[u8],
    ole: &mut OleFile<Cursor<&[u8]>>,
    scratch: &mut StreamCompareScratch,
    path: &[&str],
    expected: &[u8],
) -> (Verdict, Verdict) {
    let readback = ole
        .open_stream(path)
        .map(|data| data == expected)
        .map_err(|error| error.to_string());
    let in_place = ole
        .stream_equals(path, expected, scratch, &FileBytes(file))
        .map_err(|error| error.to_string());
    (in_place, readback)
}

/// Expectations around `actual`: itself, one byte flipped at the start, the
/// middle and the end, one byte longer and shorter, and empty.
fn expectations(actual: &[u8]) -> Vec<Vec<u8>> {
    let mut variants = vec![actual.to_vec(), Vec::new()];
    for position in [0, actual.len() / 2, actual.len().saturating_sub(1)] {
        if position < actual.len() {
            let mut flipped = actual.to_vec();
            flipped[position] ^= 0x01;
            variants.push(flipped);
        }
    }
    let mut longer = actual.to_vec();
    longer.push(0);
    variants.push(longer);
    if !actual.is_empty() {
        variants.push(actual[..actual.len() - 1].to_vec());
    }
    variants
}

/// Holds every stream of `file` to the readback verdict; returns how many
/// expectations were compared.
fn agree_on_every_stream(label: &str, file: &[u8]) -> usize {
    let Ok(mut ole) = OleFile::open(Cursor::new(file)) else {
        return 0;
    };
    let mut scratch = StreamCompareScratch::default();
    let mut compared = 0;
    for path in ole.list_streams() {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        let actual = ole.open_stream(&refs).unwrap_or_default();
        for expected in expectations(&actual) {
            let (in_place, readback) = verdicts(file, &mut ole, &mut scratch, &refs, &expected);
            assert_eq!(
                in_place,
                readback,
                "{label}: stream {path:?}, expectation of {} bytes",
                expected.len()
            );
            compared += 1;
        }
    }
    compared
}

fn repository_root() -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .and_then(Path::parent)
        .expect("crate lives two levels below the repository root")
        .to_path_buf()
}

fn collect(directory: &Path, out: &mut Vec<PathBuf>) {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return;
    };
    let mut sorted: Vec<PathBuf> = entries.flatten().map(|entry| entry.path()).collect();
    sorted.sort();
    for path in sorted {
        if path.is_dir() {
            collect(&path, out);
        } else if std::fs::metadata(&path)
            .is_ok_and(|metadata| metadata.len() >= 512 && metadata.len() <= MAX_FIXTURE_BYTES)
        {
            out.push(path);
        }
    }
}

#[test]
fn every_fixture_stream_reaches_the_readback_verdict() {
    let mut paths = Vec::new();
    collect(&repository_root().join("test-data"), &mut paths);
    let mut files = 0;
    let mut compared = 0;
    for path in paths {
        let Ok(bytes) = std::fs::read(&path) else {
            continue;
        };
        if bytes.get(..8) != Some(MAGIC) {
            continue;
        }
        files += 1;
        compared += agree_on_every_stream(&path.display().to_string(), &bytes);
    }
    assert!(files > 100, "the OLE2 corpus is present ({files} files)");
    assert!(
        compared > 1_000,
        "the corpus exercises the comparison ({compared})"
    );
    eprintln!("0749 corpus: {files} OLE2 files, {compared} expectations agree");
}

fn written(sector_size: usize) -> Vec<u8> {
    let mut writer = OleWriter::with_sector_size(sector_size).unwrap();
    writer
        .create_stream(
            &["Regular"],
            &(0..9_000u32).map(|v| v as u8).collect::<Vec<_>>(),
        )
        .unwrap();
    writer.create_stream(&["Mini"], &[7u8; 700]).unwrap();
    writer
        .create_stream(&["Storage", "Nested"], &[9u8; 130])
        .unwrap();
    writer.create_stream(&["Empty"], &[]).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn short_final_sectors_reach_the_readback_verdict() {
    for sector_size in [512, 4096] {
        let full = written(sector_size);
        for cut in [1, 17, 100, sector_size / 2, sector_size - 1] {
            let short = &full[..full.len() - cut];
            agree_on_every_stream(&format!("{sector_size}-byte sectors less {cut}"), short);
        }
    }
}

/// Deterministic xorshift for reproducible corruption.
struct Corruptor(u64);

impl Corruptor {
    fn below(&mut self, bound: usize) -> usize {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        usize::try_from(self.0 % u64::try_from(bound.max(1)).unwrap()).unwrap()
    }
}

#[test]
fn corrupted_files_reach_the_readback_verdict() {
    let mut corruptor = Corruptor(0x0749_c0de);
    let mut opened = 0;
    for sector_size in [512, 4096] {
        let original = written(sector_size);
        for round in 0..1_500 {
            let mut file = original.clone();
            // Corrupt one to three 32-bit words anywhere past the header's
            // fixed fields: FAT, MiniFAT and directory entries, sector links
            // and sizes alike.
            for _ in 0..=corruptor.below(3) {
                let word = 0x4C + corruptor.below((file.len() - 0x4C) / 4) * 4;
                let value = match corruptor.below(5) {
                    0 => 0xFFFF_FFFE,
                    1 => 0xFFFF_FFFF,
                    2 => u32::try_from(corruptor.below(40)).unwrap(),
                    3 => u32::try_from(corruptor.below(10_000)).unwrap(),
                    _ => u32::from_le_bytes(file[word..word + 4].try_into().unwrap()) ^ 1,
                };
                file[word..word + 4].copy_from_slice(&value.to_le_bytes());
            }
            if agree_on_every_stream(&format!("{sector_size}/{round}"), &file) > 0 {
                opened += 1;
            }
        }
    }
    assert!(opened > 100, "enough corrupted files still open ({opened})");
    eprintln!("0749 corruption: {opened} corrupted files opened and agree");
}
