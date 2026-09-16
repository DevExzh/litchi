//! Scratch probe for change 0644 (OLE2 snapshot source-identity fence design).
//!
//! Path dependencies point at the read-only before checkout; this record
//! changes no crate, so there is only one leg.
//!
//! trace   <mode> <path>        : every ReadAt call of one operation, classified
//! sweep   <mode> <path>        : mutate after read ordinal k, for every k
//! census  <root>               : per-.doc admission, reads and bytes
//! filever <path>               : can a FileSource SourceVersion see an
//!                                in-place write whose mtime is restored?
//! profile <mode> <path> <w> <n>: loop, for perf stat isolation pairs
//!
//! modes: doc-open  doc-open-read  ppt-open  cfb-index

use std::fs;
use std::hint::black_box;
use std::io;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

use litchi_core::{FileSource, ReadAt, SourceVersion};

type BoxError = Box<dyn std::error::Error>;

const FINGERPRINT_CHUNK: u64 = 1024 * 1024;

// ---------------------------------------------------------------- adapters

/// Records every call the fence makes, in order.
struct RecordingSource {
    inner: FileSource,
    calls: Mutex<Vec<(&'static str, u64, u64)>>,
}

impl RecordingSource {
    fn open(path: &Path) -> io::Result<Self> {
        Ok(Self {
            inner: FileSource::open(path)?,
            calls: Mutex::new(Vec::new()),
        })
    }

    fn push(&self, kind: &'static str, offset: u64, len: u64) {
        self.calls.lock().expect("probe lock").push((kind, offset, len));
    }
}

impl ReadAt for RecordingSource {
    fn len(&self) -> io::Result<u64> {
        let value = self.inner.len()?;
        self.push("len", 0, value);
        Ok(value)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let read = self.inner.read_at(offset, output)?;
        self.push("read", offset, read as u64);
        Ok(read)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let value = self.inner.version()?;
        self.push("version", 0, 0);
        Ok(value)
    }
}

/// Counts calls without recording them.
struct CountingSource {
    inner: FileSource,
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
    len_calls: AtomicU64,
    version_calls: AtomicU64,
}

impl CountingSource {
    fn open(path: &Path) -> io::Result<Self> {
        Ok(Self {
            inner: FileSource::open(path)?,
            read_calls: AtomicU64::new(0),
            read_bytes: AtomicU64::new(0),
            len_calls: AtomicU64::new(0),
            version_calls: AtomicU64::new(0),
        })
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        self.len_calls.fetch_add(1, Ordering::Relaxed);
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.read_calls.fetch_add(1, Ordering::Relaxed);
        let read = self.inner.read_at(offset, output)?;
        self.read_bytes.fetch_add(read as u64, Ordering::Relaxed);
        Ok(read)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.version_calls.fetch_add(1, Ordering::Relaxed);
        self.inner.version()
    }
}

/// A hostile adapter: owns its bytes, flips one of them after a chosen read
/// ordinal, and never moves its version token or its length. This is the
/// adapter class every complete-artifact fingerprint on the path exists to
/// defeat; `FileVersionPolicy`'s own rustdoc says a writer that restores the
/// tracked metadata is indistinguishable from it.
struct ScheduledMutationSource {
    bytes: Mutex<Vec<u8>>,
    length: u64,
    version: SourceVersion,
    trigger: u64,
    revert: u64,
    offset: u64,
    reads: AtomicU64,
    fired: AtomicU64,
}

impl ScheduledMutationSource {
    fn new(bytes: Vec<u8>, trigger: u64, offset: u64) -> Self {
        Self::transient(bytes, trigger, u64::MAX, offset)
    }

    fn transient(bytes: Vec<u8>, trigger: u64, revert: u64, offset: u64) -> Self {
        let length = bytes.len() as u64;
        Self {
            bytes: Mutex::new(bytes),
            length,
            version: SourceVersion::new(0x0644_0001, 7),
            trigger,
            revert,
            offset,
            reads: AtomicU64::new(0),
            fired: AtomicU64::new(0),
        }
    }

    fn maybe_fire(&self, reads_done: u64) {
        if reads_done != self.trigger && reads_done != self.revert {
            return;
        }
        let mut bytes = self.bytes.lock().expect("probe lock");
        let index = self.offset as usize;
        bytes[index] ^= 0xff;
        self.fired.fetch_add(1, Ordering::Relaxed);
    }
}

impl ReadAt for ScheduledMutationSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.length)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let bytes = self.bytes.lock().expect("probe lock");
        let start = usize::try_from(offset).map_err(io::Error::other)?;
        if start >= bytes.len() {
            drop(bytes);
            let done = self.reads.fetch_add(1, Ordering::Relaxed) + 1;
            self.maybe_fire(done);
            return Ok(0);
        }
        let count = output.len().min(bytes.len() - start);
        output[..count].copy_from_slice(&bytes[start..start + count]);
        drop(bytes);
        let done = self.reads.fetch_add(1, Ordering::Relaxed) + 1;
        self.maybe_fire(done);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

/// The same owned adapter with no mutation, for the clean control.
struct OwnedBytesSource {
    bytes: Vec<u8>,
    version: SourceVersion,
}

impl ReadAt for OwnedBytesSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).map_err(io::Error::other)?;
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
}

// ------------------------------------------------------------------- modes

fn short(text: String) -> String {
    let text = text.replace('\n', " ");
    if text.len() > 150 {
        format!("{}...", &text[..150])
    } else {
        text
    }
}

/// Runs one operation and reports its outcome as a stable string.
fn run(mode: &str, source: Arc<dyn ReadAt>) -> String {
    match mode {
        "doc-open" => match litchi_doc::body_text::source::SourceSnapshot::open(source) {
            Ok(snapshot) => {
                black_box(&snapshot);
                "OK".to_owned()
            },
            Err(error) => format!("open-ERR {}", short(format!("{error:?}"))),
        },
        "doc-open-read" => match litchi_doc::body_text::source::SourceSnapshot::open(source) {
            Ok(snapshot) => match snapshot.paragraph(litchi_core::Position::new(0)) {
                Ok(paragraph) => {
                    black_box(&paragraph);
                    "OK".to_owned()
                },
                Err(error) => format!("read-ERR {}", short(format!("{error:?}"))),
            },
            Err(error) => format!("open-ERR {}", short(format!("{error:?}"))),
        },
        "ppt-open" => match litchi_ppt::text_edit::SourceSnapshot::open(source) {
            Ok(snapshot) => {
                black_box(&snapshot);
                "OK".to_owned()
            },
            Err(error) => format!("open-ERR {}", short(format!("{error:?}"))),
        },
        "cfb-index" => match litchi_cfb::SharedOleFile::open(source) {
            Ok(shared) => {
                black_box(&shared);
                "OK".to_owned()
            },
            Err(error) => format!("open-ERR {}", short(format!("{error:?}"))),
        },
        other => format!("unknown-mode {other}"),
    }
}

/// True when this read is one chunk of a complete sequential fingerprint scan.
fn is_scan_chunk(offset: u64, len: u64, length: u64) -> bool {
    offset % FINGERPRINT_CHUNK == 0 && len == FINGERPRINT_CHUNK.min(length - offset)
}

fn trace(mode: &str, path: &Path) -> Result<(), BoxError> {
    let recorder = Arc::new(RecordingSource::open(path)?);
    let length = fs::metadata(path)?.len();
    let outcome = run(mode, recorder.clone() as Arc<dyn ReadAt>);
    let calls = recorder.calls.lock().expect("probe lock").clone();

    println!("# fixture={} mode={mode} bytes={length}", path.display());
    println!("# outcome={outcome}");
    println!("# fingerprint_chunk={FINGERPRINT_CHUNK}");
    println!("call_ordinal\tread_ordinal\tkind\toffset\tlen\tclass");
    let mut read_ordinal = 0_u64;
    let mut scan_chunks = 0_u64;
    for (call_ordinal, (kind, offset, len)) in calls.iter().enumerate() {
        let mut class = "-";
        if *kind == "read" {
            read_ordinal += 1;
            if is_scan_chunk(*offset, *len, length) {
                class = "scan-chunk";
                scan_chunks += 1;
            } else {
                class = "index-or-range";
            }
        }
        println!(
            "{}\t{}\t{kind}\t{offset}\t{len}\t{class}",
            call_ordinal + 1,
            if *kind == "read" {
                read_ordinal.to_string()
            } else {
                String::new()
            },

        );
    }
    let chunks_per_scan = length.div_ceil(FINGERPRINT_CHUNK);
    println!("# read_calls={read_ordinal}");
    println!("# scan_chunks={scan_chunks} chunks_per_scan={chunks_per_scan}");
    println!(
        "# complete_scans={}",
        scan_chunks as f64 / chunks_per_scan as f64
    );
    Ok(())
}

/// Offsets that the index parse and the semantic parse actually consume, so a
/// mutation witness can be placed outside them and exercise the fence rather
/// than the parser.
fn parsed_coverage(mode: &str, path: &Path) -> Result<(Vec<bool>, u64), BoxError> {
    let recorder = Arc::new(RecordingSource::open(path)?);
    let length = fs::metadata(path)?.len();
    let _ = run(mode, recorder.clone() as Arc<dyn ReadAt>);
    let calls = recorder.calls.lock().expect("probe lock").clone();
    let mut covered = vec![false; length as usize];
    for (kind, offset, len) in calls {
        if kind != "read" || is_scan_chunk(offset, len, length) {
            continue;
        }
        let start = offset as usize;
        let end = (start + len as usize).min(covered.len());
        for flag in &mut covered[start..end] {
            *flag = true;
        }
    }
    Ok((covered, length))
}

fn witness_offset(mode: &str, path: &Path) -> Result<u64, BoxError> {
    let (covered, _length) = parsed_coverage(mode, path)?;
    for index in (0..covered.len()).rev() {
        if !covered[index] {
            return Ok(index as u64);
        }
    }
    Err(format!("{}: every byte is consumed by the parse", path.display()).into())
}

fn sweep(mode: &str, path: &Path, forced: Option<u64>) -> Result<(), BoxError> {
    let bytes = fs::read(path)?;
    let length = bytes.len() as u64;
    let offset = match forced {
        Some(value) => value,
        None => witness_offset(mode, path)?,
    };

    // Clean control on the same owned adapter: the sweep must not pass by
    // refusing everything.
    let clean: Arc<dyn ReadAt> = Arc::new(OwnedBytesSource {
        bytes: bytes.clone(),
        version: SourceVersion::new(0x0644_0001, 7),
    });
    let clean_outcome = run(mode, clean);
    let clean_reads = {
        let source = Arc::new(ScheduledMutationSource::new(bytes.clone(), u64::MAX, offset));
        let _ = run(mode, source.clone() as Arc<dyn ReadAt>);
        source.reads.load(Ordering::Relaxed)
    };

    println!("# fixture={} mode={mode} bytes={length}", path.display());
    println!("# witness_offset={offset} forced={}", forced.is_some());
    println!("# clean_outcome={clean_outcome}");
    println!("# clean_read_calls={clean_reads}");
    println!("trigger_after_read\tfired\treads\toutcome");
    for trigger in 0..=clean_reads {
        let source = Arc::new(ScheduledMutationSource::new(bytes.clone(), trigger, offset));
        let outcome = run(mode, source.clone() as Arc<dyn ReadAt>);
        println!(
            "{trigger}\t{}\t{}\t{outcome}",
            source.fired.load(Ordering::Relaxed),
            source.reads.load(Ordering::Relaxed)
        );
    }
    Ok(())
}

/// One flip-then-revert witness: the mutation exists only between two read
/// ordinals and leaves the artifact byte-identical afterwards.
fn transient(
    mode: &str,
    path: &Path,
    flip: u64,
    revert: u64,
    offset: Option<u64>,
) -> Result<(), BoxError> {
    let bytes = fs::read(path)?;
    let offset = match offset {
        Some(value) => value,
        None => witness_offset(mode, path)?,
    };
    let original = bytes.clone();
    let source = Arc::new(ScheduledMutationSource::transient(
        bytes, flip, revert, offset,
    ));
    let outcome = run(mode, source.clone() as Arc<dyn ReadAt>);
    let final_bytes = source.bytes.lock().expect("probe lock").clone();
    println!("# fixture={} mode={mode}", path.display());
    println!("flip_after_read\t{flip}");
    println!("revert_after_read\t{revert}");
    println!("offset\t{offset}");
    println!("reads\t{}", source.reads.load(Ordering::Relaxed));
    println!("mutations_applied\t{}", source.fired.load(Ordering::Relaxed));
    println!("restored_byte_identical\t{}", final_bytes == original);
    println!("outcome\t{outcome}");
    Ok(())
}

fn walk(root: &Path, ext: &str, out: &mut Vec<PathBuf>) {
    if let Ok(entries) = fs::read_dir(root) {
        let mut paths: Vec<PathBuf> = entries.flatten().map(|entry| entry.path()).collect();
        paths.sort();
        for path in paths {
            if path.is_dir() {
                walk(&path, ext, out);
            } else if path
                .extension()
                .and_then(|value| value.to_str())
                .is_some_and(|value| value.eq_ignore_ascii_case(ext))
            {
                out.push(path);
            }
        }
    }
}

fn census(mode: &str, root: &Path, ext: &str) -> Result<(), BoxError> {
    let mut paths = Vec::new();
    walk(root, ext, &mut paths);
    println!("file\tbytes\tread_calls\tread_bytes\tfull_artifact_reads\tlen_calls\tversion_calls\toutcome");
    for path in paths {
        let source = Arc::new(CountingSource::open(&path)?);
        let length = fs::metadata(&path)?.len();
        let outcome = run(mode, source.clone() as Arc<dyn ReadAt>);
        let read_bytes = source.read_bytes.load(Ordering::Relaxed);
        println!(
            "{}\t{length}\t{}\t{read_bytes}\t{:.4}\t{}\t{}\t{outcome}",
            path.display(),
            source.read_calls.load(Ordering::Relaxed),
            read_bytes as f64 / length as f64,
            source.len_calls.load(Ordering::Relaxed),
            source.version_calls.load(Ordering::Relaxed),
        );
    }
    Ok(())
}

/// The witness for "could a `SourceVersion` observation replace a complete
/// read on a `FileSource`?". Writes one byte in place through a second
/// descriptor on the same inode and restores the modification time.
fn filever(path: &Path) -> Result<(), BoxError> {
    let scratch = std::env::temp_dir().join("fence-0644-witness.bin");
    fs::copy(path, &scratch)?;

    let source = FileSource::open(&scratch)?;
    let before_version = source.version()?;
    let before_length = source.len()?;
    let mut probe = [0_u8; 1];
    source.read_exact_at(0, &mut probe)?;
    let before_byte = probe[0];

    let metadata = fs::metadata(&scratch)?;
    let times = fs::FileTimes::new()
        .set_accessed(metadata.accessed()?)
        .set_modified(metadata.modified()?);

    {
        let writer = fs::File::options().write(true).open(&scratch)?;
        std::os::unix::fs::FileExt::write_all_at(&writer, &[before_byte ^ 0xff], 0)?;
        writer.sync_all()?;
        writer.set_times(times)?;
    }

    let after_version = source.version()?;
    let after_length = source.len()?;
    source.read_exact_at(0, &mut probe)?;
    let after_byte = probe[0];

    println!("# fixture={}", path.display());
    println!("policy\t{:?}", source.version_policy());
    println!("byte_before\t{before_byte}");
    println!("byte_after\t{after_byte}");
    println!("bytes_changed\t{}", before_byte != after_byte);
    println!("length_before\t{before_length}");
    println!("length_after\t{after_length}");
    println!("length_changed\t{}", before_length != after_length);
    println!("version_before\t{before_version:?}");
    println!("version_after\t{after_version:?}");
    println!("version_changed\t{}", before_version != after_version);
    println!(
        "verdict\t{}",
        if before_version == after_version && before_length == after_length && before_byte != after_byte {
            "SourceVersion and length are both blind to this mutation"
        } else {
            "detected"
        }
    );
    fs::remove_file(&scratch)?;
    Ok(())
}

fn profile(mode: &str, path: &Path, warmup: u64, samples: u64) -> Result<(), BoxError> {
    let mut sink = 0_usize;
    for _ in 0..warmup {
        let source: Arc<dyn ReadAt> = Arc::new(FileSource::open(path)?);
        sink += run(mode, source).len();
    }
    for _ in 0..samples {
        let source: Arc<dyn ReadAt> = Arc::new(FileSource::open(path)?);
        sink += run(mode, source).len();
    }
    black_box(sink);
    Ok(())
}

fn main() -> Result<(), BoxError> {
    let args: Vec<String> = std::env::args().collect();
    match args.get(1).map(String::as_str) {
        Some("trace") => trace(&args[2], Path::new(&args[3])),
        Some("sweep") => sweep(
            &args[2],
            Path::new(&args[3]),
            args.get(4).map(|value| value.parse()).transpose()?,
        ),
        Some("census") => census(&args[2], Path::new(&args[3]), &args[4]),
        Some("transient") => transient(
            &args[2],
            Path::new(&args[3]),
            args[4].parse()?,
            args[5].parse()?,
            args.get(6).map(|value| value.parse()).transpose()?,
        ),
        Some("filever") => filever(Path::new(&args[2])),
        Some("profile") => profile(
            &args[2],
            Path::new(&args[3]),
            args[4].parse()?,
            args[5].parse()?,
        ),
        other => Err(format!("usage: {} <trace|sweep|transient|census|filever|profile> ...; got {other:?}", args[0]).into()),
    }
}
