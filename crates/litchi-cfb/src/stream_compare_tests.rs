//! `StreamComparer::stream_equals` held to `OleFile::open_stream` (change
//! 0749).
//!
//! The in-place comparison must reach the verdict `open_stream` followed by a
//! byte comparison reaches — equal, different, or the same error — for every
//! stream of every OLE2 fixture, for expectations that differ in one byte or
//! in length, for files whose final sector is short, and for files whose
//! allocation tables were corrupted after they were written. The verdicts are
//! compared as sequences: a fresh reader and a fresh comparer see the same
//! calls in the same order, so the one-time root mini-stream load of each is
//! exercised the same way. Further tests pin that load to once per comparer,
//! a contiguous mini stream to one comparison, and a failed load to no
//! cached result.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]

use std::cell::Cell;
use std::io::Cursor;
use std::path::{Path, PathBuf};

use crate::file::CandidateBytes;
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
    let Ok(mut first) = OleFile::open(Cursor::new(file)) else {
        return 0;
    };
    let mut cases = Vec::new();
    for path in first.list_streams() {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        let actual = first.open_stream(&refs).unwrap_or_default();
        for expected in expectations(&actual) {
            cases.push((path.clone(), expected));
        }
    }
    // A fresh reader, whose mini-stream cache starts empty as a comparer's
    // root load does, answers the readback in order.
    let mut reader = OleFile::open(Cursor::new(file)).unwrap();
    let readback: Vec<Verdict> = cases
        .iter()
        .map(|(path, expected)| {
            let refs: Vec<&str> = path.iter().map(String::as_str).collect();
            reader
                .open_stream(&refs)
                .map(|data| data == *expected)
                .map_err(|error| error.to_string())
        })
        .collect();
    let candidate = FileBytes(file);
    let mut comparer = reader.stream_comparer(&candidate);
    for ((path, expected), readback) in cases.iter().zip(&readback) {
        let refs: Vec<&str> = path.iter().map(String::as_str).collect();
        let in_place = comparer
            .stream_equals(&refs, expected)
            .map_err(|error| error.to_string());
        assert_eq!(
            &in_place,
            readback,
            "{label}: stream {path:?}, expectation of {} bytes",
            expected.len()
        );
    }
    cases.len()
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

/// A candidate that counts what it is asked, optionally failing its first
/// few readability checks.
struct Counting<'a> {
    bytes: FileBytes<'a>,
    equals: Cell<usize>,
    readable: Cell<usize>,
    failing_checks: Cell<usize>,
}

impl<'a> Counting<'a> {
    fn new(bytes: &'a [u8], failing_checks: usize) -> Self {
        Self {
            bytes: FileBytes(bytes),
            equals: Cell::new(0),
            readable: Cell::new(0),
            failing_checks: Cell::new(failing_checks),
        }
    }
}

impl CandidateBytes for Counting<'_> {
    fn equals_at(&self, offset: u64, expected: &[u8]) -> Result<bool, OleError> {
        self.equals.set(self.equals.get() + 1);
        self.bytes.equals_at(offset, expected)
    }

    fn check_readable(&self, offset: u64, len: usize) -> Result<(), OleError> {
        self.readable.set(self.readable.get() + 1);
        if self.failing_checks.get() > 0 {
            self.failing_checks.set(self.failing_checks.get() - 1);
            return Err(OleError::InvalidData("injected read failure".to_string()));
        }
        self.bytes.check_readable(offset, len)
    }
}

/// `count` mini streams of `size` bytes each, written from scratch.
fn many_mini_streams(sector_size: usize, count: usize, size: usize) -> (Vec<u8>, Vec<Vec<u8>>) {
    let mut writer = OleWriter::with_sector_size(sector_size).unwrap();
    let mut payloads = Vec::new();
    for index in 0..count {
        let payload: Vec<u8> = (0..size)
            .map(|byte| u8::try_from((byte * 7 + index * 13) % 251).unwrap())
            .collect();
        let name = format!("S{index:05}");
        writer.create_stream(&[name.as_str()], &payload).unwrap();
        payloads.push(payload);
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    (output.into_inner(), payloads)
}

#[test]
fn the_root_mini_stream_is_loaded_once_and_each_contiguous_stream_compared_once() {
    for sector_size in [512, 4096] {
        let (file, payloads) = many_mini_streams(sector_size, 300, 2_000);
        let reader = OleFile::open(Cursor::new(file.as_slice())).unwrap();
        let candidate = Counting::new(&file, 0);
        let mut comparer = reader.stream_comparer(&candidate);
        let mut root_checks = None;
        for (index, payload) in payloads.iter().enumerate() {
            let name = format!("S{index:05}");
            assert!(comparer.stream_equals(&[name.as_str()], payload).unwrap());
            // The root load's readability checks happen for the first mini
            // stream only; the quadratic form repeated them for every one.
            let checks = *root_checks.get_or_insert(candidate.readable.get());
            assert_eq!(
                candidate.readable.get(),
                checks,
                "{sector_size}: stream {index}"
            );
        }
        assert!(root_checks.is_some_and(|checks| checks >= 1));
        // Each stream's 32 mini sectors follow one another in the file, so
        // they are compared as one range, not 64 bytes at a time.
        assert_eq!(candidate.equals.get(), payloads.len(), "{sector_size}");
    }
}

#[test]
fn a_failed_root_load_is_not_kept_and_is_repeated_like_open_stream() {
    let (file, payloads) = many_mini_streams(512, 3, 700);
    let reader = OleFile::open(Cursor::new(file.as_slice())).unwrap();
    // The first readability check fails once: the first mini stream reports
    // the error, and the next mini stream loads the root again, successfully.
    let candidate = Counting::new(&file, 1);
    let mut comparer = reader.stream_comparer(&candidate);
    let error = comparer
        .stream_equals(&["S00000"], &payloads[0])
        .unwrap_err();
    assert!(
        error.to_string().contains("injected read failure"),
        "{error}"
    );
    assert!(comparer.stream_equals(&["S00001"], &payloads[1]).unwrap());
    assert!(comparer.stream_equals(&["S00000"], &payloads[0]).unwrap());
    let loaded = candidate.readable.get();
    assert!(comparer.stream_equals(&["S00002"], &payloads[2]).unwrap());
    assert_eq!(
        candidate.readable.get(),
        loaded,
        "a successful load is kept"
    );

    // A candidate whose reads always fail fails every mini stream the same
    // way, as `open_stream` fails every read of an unreadable mini stream.
    let failing = Counting::new(&file, usize::MAX);
    let mut comparer = reader.stream_comparer(&failing);
    for (index, payload) in payloads.iter().enumerate() {
        let name = format!("S{index:05}");
        let error = comparer
            .stream_equals(&[name.as_str()], payload)
            .unwrap_err();
        assert!(
            error.to_string().contains("injected read failure"),
            "{error}"
        );
    }
}

#[test]
fn many_mini_stream_files_reach_the_readback_verdict() {
    for sector_size in [512, 4096] {
        let (file, _) = many_mini_streams(sector_size, 120, 1_000);
        assert!(agree_on_every_stream(&format!("{sector_size}-byte sectors"), &file) > 0);
    }
}
