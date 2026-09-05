//! Standalone ZIP strict-read/index probe for change 0416.
//!
//! The probe deliberately uses the public `soapberry-zip` Office APIs.  The
//! `borrowed` operation constructs an in-memory `ArchiveReader` and exercises
//! `read_stored_borrowed` for every indexed member.  The `indexed` operation
//! opens an `IndexedArchive` over a `ReaderAt` source and builds its
//! `PreservationIndex`, which validates local records and descriptors without
//! reading or decompressing member payloads.  A separate `count` command
//! repeats the indexed operation with a ReaderAt wrapper that records calls
//! and accepted bytes; the wrapper is intentionally absent from latency
//! samples.
//!
//! Usage:
//!
//! ```text
//! probe index borrowed PATH SAMPLES WARMUPS
//! probe index indexed PATH SAMPLES WARMUPS
//! probe capability borrowed|indexed PATH
//! probe count PATH
//! ```
//!
//! Input archives are generated outside this process and are never generated
//! in a timed loop.  The timed oracle only consumes declared lengths, CRC
//! values and raw/normalized member-name bytes through `black_box`; payloads
//! are not decompressed.

use serde_json::{Value, json};
use soapberry_zip::office::{ArchiveReader, IndexedArchive};
use soapberry_zip::{RECOMMENDED_BUFFER_SIZE, ReaderAt, ZipArchive};
use std::collections::HashMap;
use std::fs;
use std::hint::black_box;
use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};
use std::time::Instant;

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

#[derive(Debug, Clone)]
struct OracleMember {
    normalized_name: Option<String>,
    raw_name: Vec<u8>,
    compressed_size: u64,
    uncompressed_size: u64,
    crc32: u32,
}

#[derive(Debug)]
struct Oracle {
    files: HashMap<String, OracleMember>,
    raw: HashMap<Vec<u8>, OracleMember>,
    file_count: usize,
    total_count: usize,
    name_bytes: u64,
    compressed_bytes: u64,
    uncompressed_bytes: u64,
    crc32_xor: u32,
}

impl Oracle {
    fn from_bytes(data: &[u8]) -> Result<Self> {
        let archive = ZipArchive::from_slice(data)?;
        let mut files = HashMap::new();
        let mut raw = HashMap::new();
        let mut file_count = 0usize;
        let mut total_count = 0usize;
        let mut name_bytes = 0u64;
        let mut compressed_bytes = 0u64;
        let mut uncompressed_bytes = 0u64;
        let mut crc32_xor = 0u32;
        let mut entries = archive.entries();
        while let Some(record) = entries.next_entry()? {
            total_count = total_count
                .checked_add(1)
                .ok_or("oracle entry count overflow")?;
            let raw_name = record.file_path().as_ref().to_vec();
            let raw_len = u64::try_from(raw_name.len())?;
            name_bytes = name_bytes
                .checked_add(raw_len)
                .ok_or("oracle name-byte count overflow")?;
            let member = OracleMember {
                normalized_name: if record.is_dir() {
                    None
                } else {
                    Some(record.file_path().try_normalize()?.as_ref().to_owned())
                },
                raw_name: raw_name.clone(),
                compressed_size: record.compressed_size_hint(),
                uncompressed_size: record.uncompressed_size_hint(),
                crc32: record.crc32(),
            };
            compressed_bytes = compressed_bytes
                .checked_add(member.compressed_size)
                .ok_or("oracle compressed-byte count overflow")?;
            uncompressed_bytes = uncompressed_bytes
                .checked_add(member.uncompressed_size)
                .ok_or("oracle uncompressed-byte count overflow")?;
            crc32_xor ^= member.crc32;
            if let Some(name) = member.normalized_name.as_ref() {
                file_count = file_count
                    .checked_add(1)
                    .ok_or("oracle file count overflow")?;
                if files.insert(name.clone(), member.clone()).is_some() {
                    return Err(format!("duplicate normalized oracle name: {name}").into());
                }
            }
            if raw.insert(raw_name, member).is_some() {
                return Err("duplicate raw oracle name".into());
            }
        }
        Ok(Self {
            files,
            raw,
            file_count,
            total_count,
            name_bytes,
            compressed_bytes,
            uncompressed_bytes,
            crc32_xor,
        })
    }

    fn json(&self) -> Value {
        json!({
            "file_count": self.file_count,
            "central_entry_count": self.total_count,
            "name_bytes": self.name_bytes,
            "compressed_bytes": self.compressed_bytes,
            "uncompressed_bytes": self.uncompressed_bytes,
            "crc32_xor": self.crc32_xor,
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Observation {
    digest: u64,
    entries: u64,
    names: u64,
    compressed_bytes: u64,
    uncompressed_bytes: u64,
    crc32_xor: u32,
    borrowed_some: u64,
    borrowed_none: u64,
    borrowed_crc_scan_bytes: u64,
}

impl Observation {
    fn json(self) -> Value {
        json!({
            "digest": self.digest,
            "entries": self.entries,
            "name_bytes": self.names,
            "compressed_bytes": self.compressed_bytes,
            "uncompressed_bytes": self.uncompressed_bytes,
            "crc32_xor": self.crc32_xor,
            "borrowed_some": self.borrowed_some,
            "borrowed_none": self.borrowed_none,
            "borrowed_crc_scan_bytes": self.borrowed_crc_scan_bytes,
        })
    }
}

#[derive(Debug, Default)]
struct ReadCounters {
    calls: AtomicU64,
    bytes: AtomicU64,
}

#[derive(Debug)]
struct CountingReaderAt<R> {
    inner: R,
    counters: Arc<ReadCounters>,
}

impl<R> CountingReaderAt<R> {
    fn new(inner: R) -> (Self, Arc<ReadCounters>) {
        let counters = Arc::new(ReadCounters::default());
        (
            Self {
                inner,
                counters: Arc::clone(&counters),
            },
            counters,
        )
    }
}

impl<R: ReaderAt> ReaderAt for CountingReaderAt<R> {
    fn read_at(&self, buffer: &mut [u8], offset: u64) -> io::Result<usize> {
        let read = self.inner.read_at(buffer, offset)?;
        self.counters.calls.fetch_add(1, Ordering::Relaxed);
        self.counters.bytes.fetch_add(
            u64::try_from(read).map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidData,
                    "ReaderAt read count overflows u64",
                )
            })?,
            Ordering::Relaxed,
        );
        Ok(read)
    }
}

fn mix_bytes(mut digest: u64, bytes: &[u8]) -> u64 {
    let bytes = black_box(bytes);
    for byte in bytes {
        digest = digest.rotate_left(5) ^ u64::from(black_box(*byte));
    }
    digest
}

fn mix_member(
    mut digest: u64,
    raw_name: &[u8],
    compressed_size: u64,
    uncompressed_size: u64,
    crc32: u32,
) -> u64 {
    digest = mix_bytes(digest, raw_name);
    digest = digest.rotate_left(7) ^ black_box(compressed_size);
    digest = digest.rotate_left(7) ^ black_box(uncompressed_size);
    digest = digest.rotate_left(7) ^ u64::from(black_box(crc32));
    digest
}

fn consume_borrowed(archive: &ArchiveReader<'_>, oracle: &Oracle) -> Result<Observation> {
    let mut digest = 0x0416_5eed_cafe_beefu64;
    let mut entries = 0u64;
    let mut names = 0u64;
    let mut compressed_bytes = 0u64;
    let mut uncompressed_bytes = 0u64;
    let mut crc32_xor = 0u32;
    let mut borrowed_some = 0u64;
    let mut borrowed_none = 0u64;
    let mut borrowed_crc_scan_bytes = 0u64;

    for name in archive.file_names() {
        let member = oracle
            .files
            .get(name)
            .ok_or_else(|| format!("borrowed oracle has no member {name:?}"))?;
        let metadata = archive.metadata(name)?;
        if metadata.is_directory()
            || metadata.compressed_size() != member.compressed_size
            || metadata.uncompressed_size() != member.uncompressed_size
        {
            return Err(format!("borrowed metadata mismatch for {name:?}").into());
        }
        let bytes = archive.read_stored_borrowed(name)?;
        match bytes {
            Some(bytes) => {
                if u64::try_from(bytes.len())? != member.uncompressed_size {
                    return Err(format!("borrowed payload length mismatch for {name:?}").into());
                }
                borrowed_some = borrowed_some
                    .checked_add(1)
                    .ok_or("borrowed Some count overflow")?;
                borrowed_crc_scan_bytes = borrowed_crc_scan_bytes
                    .checked_add(u64::try_from(bytes.len())?)
                    .ok_or("borrowed CRC scan byte count overflow")?;
                digest = digest.rotate_left(3) ^ black_box(u64::try_from(bytes.len())?);
            },
            None => {
                borrowed_none = borrowed_none
                    .checked_add(1)
                    .ok_or("borrowed None count overflow")?;
            },
        }
        let name_len = u64::try_from(black_box(name.as_bytes()).len())?;
        names = names
            .checked_add(name_len)
            .ok_or("borrowed name count overflow")?;
        compressed_bytes = compressed_bytes
            .checked_add(metadata.compressed_size())
            .ok_or("borrowed compressed count overflow")?;
        uncompressed_bytes = uncompressed_bytes
            .checked_add(metadata.uncompressed_size())
            .ok_or("borrowed uncompressed count overflow")?;
        crc32_xor ^= black_box(member.crc32);
        digest = mix_member(
            digest,
            name.as_bytes(),
            metadata.compressed_size(),
            metadata.uncompressed_size(),
            member.crc32,
        );
        entries = entries
            .checked_add(1)
            .ok_or("borrowed entry count overflow")?;
    }

    if entries != u64::try_from(oracle.file_count)? {
        return Err(format!(
            "borrowed entry count {} differs from oracle {}",
            entries, oracle.file_count
        )
        .into());
    }
    if borrowed_some == 0 {
        return Err(
            "borrowed guard requires a Store member to exercise strict layout validation".into(),
        );
    }
    Ok(Observation {
        digest,
        entries,
        names,
        compressed_bytes,
        uncompressed_bytes,
        crc32_xor,
        borrowed_some,
        borrowed_none,
        borrowed_crc_scan_bytes,
    })
}

fn consume_preservation<R: ReaderAt>(
    index: &soapberry_zip::PreservationIndex<'_, R>,
    oracle: &Oracle,
) -> Result<Observation> {
    let mut digest = 0x0416_5eed_cafe_beefu64;
    let mut entries = 0u64;
    let mut names = 0u64;
    let mut compressed_bytes = 0u64;
    let mut uncompressed_bytes = 0u64;
    let mut crc32_xor = 0u32;

    for entry in index.entries() {
        let raw_name = entry.raw_name_bytes();
        let member = oracle
            .raw
            .get(raw_name)
            .ok_or_else(|| format!("preservation oracle has no raw member {raw_name:?}"))?;
        if member.raw_name.as_slice() != raw_name {
            return Err("preservation raw-name oracle mismatch".into());
        }
        if entry.compressed_size() != member.compressed_size
            || entry.uncompressed_size() != member.uncompressed_size
        {
            return Err(format!("preservation metadata mismatch for {raw_name:?}").into());
        }
        let raw_len = u64::try_from(black_box(raw_name).len())?;
        names = names
            .checked_add(raw_len)
            .ok_or("preservation name count overflow")?;
        compressed_bytes = compressed_bytes
            .checked_add(entry.compressed_size())
            .ok_or("preservation compressed count overflow")?;
        uncompressed_bytes = uncompressed_bytes
            .checked_add(entry.uncompressed_size())
            .ok_or("preservation uncompressed count overflow")?;
        crc32_xor ^= black_box(member.crc32);
        digest = mix_member(
            digest,
            raw_name,
            entry.compressed_size(),
            entry.uncompressed_size(),
            member.crc32,
        );
        entries = entries
            .checked_add(1)
            .ok_or("preservation entry count overflow")?;
    }

    if entries != u64::try_from(oracle.total_count)? {
        return Err(format!(
            "preservation entry count {} differs from oracle {}",
            entries, oracle.total_count
        )
        .into());
    }
    Ok(Observation {
        digest,
        entries,
        names,
        compressed_bytes,
        uncompressed_bytes,
        crc32_xor,
        borrowed_some: 0,
        borrowed_none: 0,
        borrowed_crc_scan_bytes: 0,
    })
}

fn borrowed_once(data: &[u8], oracle: &Oracle) -> Result<Observation> {
    let archive = ArchiveReader::new(data)?;
    consume_borrowed(&archive, oracle)
}

fn indexed_once<R: ReaderAt>(
    source: R,
    source_len: u64,
    scratch: &mut [u8],
    oracle: &Oracle,
) -> Result<Observation> {
    let archive = IndexedArchive::from_reader(source, source_len)?;
    let index = archive.preservation_index(scratch)?;
    consume_preservation(&index, oracle)
}

fn run_timed(mode: &str, path: &str, samples: usize, warmups: usize) -> Result<Value> {
    if samples == 0 {
        return Err("samples must be positive".into());
    }
    let data = fs::read(path)?;
    let source_len = u64::try_from(data.len())?;
    let oracle = Oracle::from_bytes(&data)?;
    let mut sample_ns = Vec::with_capacity(samples);
    let mut expected = None;

    for iteration in 0..warmups.checked_add(samples).ok_or("iteration overflow")? {
        let start = Instant::now();
        let observation = match mode {
            "borrowed" => borrowed_once(&data, &oracle)?,
            "indexed" | "reader-at" | "reader_at" => {
                let mut scratch = vec![0u8; RECOMMENDED_BUFFER_SIZE];
                indexed_once(data.as_slice(), source_len, &mut scratch, &oracle)?
            },
            _ => return Err(format!("unknown reader mode {mode:?}").into()),
        };
        let elapsed_ns = u64::try_from(start.elapsed().as_nanos())?;
        if let Some(previous) = expected {
            if previous != observation {
                return Err(format!(
                    "metadata oracle changed between iterations: {previous:?} != {observation:?}"
                )
                .into());
            }
        } else {
            expected = Some(observation);
        }
        if iteration >= warmups {
            sample_ns.push(elapsed_ns.max(1));
        }
    }

    let observation = expected.ok_or("at least one operation is required")?;
    Ok(json!({
        "schema_version": 1,
        "tool": "litchi-goal-0416-zip-strict-probe",
        "operation": "index",
        "mode": mode,
        "input_path": path,
        "input_bytes": source_len,
        "warmups": warmups,
        "sample_count": samples,
        "samples_ns": sample_ns,
        "oracle": oracle.json(),
        "observation": observation.json(),
        "reader_at_instrumentation": "disabled in latency mode",
        "payload_decompression": false,
        "borrowed_store_crc_scan": mode == "borrowed",
    }))
}

fn run_counted(path: &str) -> Result<Value> {
    let data = fs::read(path)?;
    let oracle = Oracle::from_bytes(&data)?;
    let source_len = u64::try_from(data.len())?;
    let (source, counters) = CountingReaderAt::new(data.as_slice());
    let mut scratch = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    let start = Instant::now();
    let observation = indexed_once(source, source_len, &mut scratch, &oracle)?;
    let elapsed_ns = u64::try_from(start.elapsed().as_nanos())?;
    Ok(json!({
        "schema_version": 1,
        "tool": "litchi-goal-0416-zip-strict-probe",
        "operation": "reader-at-count",
        "mode": "indexed",
        "input_path": path,
        "input_bytes": source_len,
        "elapsed_ns": elapsed_ns.max(1),
        "oracle": oracle.json(),
        "observation": observation.json(),
        "reader_at": {
            "calls": counters.calls.load(Ordering::Relaxed),
            "bytes": counters.bytes.load(Ordering::Relaxed),
        },
        "payload_decompression": false,
        "borrowed_store_crc_scan": false,
        "latency_samples_include_reader_at_instrumentation": false,
    }))
}

fn run_capability(mode: &str, path: &str) -> Value {
    let started = Instant::now();
    let input = fs::read(path);
    let input_bytes = input.as_ref().map_or(0, |data| data.len() as u64);
    let result = (|| -> Result<Observation> {
        let data = input?;
        let oracle = Oracle::from_bytes(&data)?;
        match mode {
            "borrowed" => borrowed_once(&data, &oracle),
            "indexed" | "reader-at" | "reader_at" => {
                let source_len = u64::try_from(data.len())?;
                let mut scratch = vec![0u8; RECOMMENDED_BUFFER_SIZE];
                indexed_once(data.as_slice(), source_len, &mut scratch, &oracle)
            },
            _ => Err(format!("unknown reader mode {mode:?}").into()),
        }
    })();
    match result {
        Ok(observation) => json!({
            "schema_version": 1,
            "tool": "litchi-goal-0416-zip-strict-probe",
            "operation": "capability",
            "mode": mode,
            "input_path": path,
            "input_bytes": input_bytes,
            "status": "ok",
            "elapsed_ns": u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX).max(1),
            "observation": observation.json(),
            "payload_decompression": false,
            "borrowed_store_crc_scan": mode == "borrowed",
        }),
        Err(error) => json!({
            "schema_version": 1,
            "tool": "litchi-goal-0416-zip-strict-probe",
            "operation": "capability",
            "mode": mode,
            "input_path": path,
            "input_bytes": input_bytes,
            "status": "error",
            "elapsed_ns": u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX).max(1),
            "error": error.to_string(),
        "payload_decompression": false,
        "borrowed_store_crc_scan": mode == "borrowed",
        }),
    }
}

fn usage() -> &'static str {
    "usage: probe index borrowed|indexed PATH SAMPLES WARMUPS | probe capability borrowed|indexed PATH | probe count PATH"
}

fn main() -> Result<()> {
    let args: Vec<String> = std::env::args().collect();
    let output = match args.get(1).map(String::as_str) {
        Some("index") => {
            let mode = args.get(2).ok_or(usage())?;
            let path = args.get(3).ok_or(usage())?;
            let samples: usize = args.get(4).ok_or(usage())?.parse()?;
            let warmups: usize = args.get(5).ok_or(usage())?.parse()?;
            run_timed(mode, path, samples, warmups)?
        },
        Some("capability") => {
            let mode = args.get(2).ok_or(usage())?;
            let path = args.get(3).ok_or(usage())?;
            run_capability(mode, path)
        },
        Some("count") => {
            let path = args.get(2).ok_or(usage())?;
            run_counted(path)?
        },
        Some(mode @ ("borrowed" | "indexed" | "reader-at" | "reader_at")) => {
            let path = args.get(2).ok_or(usage())?;
            let samples: usize = args.get(3).ok_or(usage())?.parse()?;
            let warmups: usize = args.get(4).ok_or(usage())?.parse()?;
            run_timed(mode, path, samples, warmups)?
        },
        _ => return Err(usage().into()),
    };
    println!("{}", serde_json::to_string_pretty(&output)?);
    Ok(())
}
