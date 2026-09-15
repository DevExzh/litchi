//! Peak-retained-bytes probe for ZIP passthrough save paths (change 0578, stage 1).
//!
//! Scenario A  copy-through: the production source-backed OOXML save shape.
//!             PreservationAction::Copy for every unchanged member, Regenerate
//!             for one small edited member, positional FileReader source,
//!             sequential non-seek sink.
//! Scenario B  precompressed regeneration: the cross-document copy shape.
//!             capture the large member as a VerifiedPrecompressedEntry and
//!             republish it through RegeneratedEntry::new_precompressed_shared.
//! Scenario C  ZipArchiveWriter::write_precompressed_file(&[u8]), the API named
//!             by the hypothesis. Zero production callers; measured for record.

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::Write;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

use soapberry_zip::office::{
    ArchiveLimits, IndexedArchive, StreamingArchiveLimits, StreamingArchiveWriter,
};
use soapberry_zip::{
    CompressionMethod, FileReader, PreservationAction, PreservationPlan, RECOMMENDED_BUFFER_SIZE,
    ReaderAt, RegeneratedEntry, ZipArchive, ZipArchiveWriter,
};

// ---------------------------------------------------------------- allocator

static LIVE: AtomicU64 = AtomicU64::new(0);
static PEAK: AtomicU64 = AtomicU64::new(0);
static ALLOCS: AtomicU64 = AtomicU64::new(0);
static ALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static ON: AtomicBool = AtomicBool::new(false);

fn bump(delta: i64) {
    if !ON.load(Ordering::Relaxed) {
        return;
    }
    let live = if delta >= 0 {
        LIVE.fetch_add(delta as u64, Ordering::Relaxed) + delta as u64
    } else {
        LIVE.fetch_sub((-delta) as u64, Ordering::Relaxed) - (-delta) as u64
    };
    PEAK.fetch_max(live, Ordering::Relaxed);
}

struct Counting;

// SAFETY: every method delegates to `std::alloc::System`; the counter updates
// are side-effect-only and change neither pointer, layout, nor ownership.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller upholds the `GlobalAlloc` layout contract.
        let p = unsafe { System.alloc(layout) };
        if !p.is_null() {
            ALLOCS.fetch_add(1, Ordering::Relaxed);
            ALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
            bump(layout.size() as i64);
        }
        p
    }
    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller upholds the `GlobalAlloc` layout contract.
        let p = unsafe { System.alloc_zeroed(layout) };
        if !p.is_null() {
            ALLOCS.fetch_add(1, Ordering::Relaxed);
            ALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
            bump(layout.size() as i64);
        }
        p
    }
    unsafe fn dealloc(&self, p: *mut u8, layout: Layout) {
        // SAFETY: the caller upholds the `GlobalAlloc` pointer/layout contract.
        unsafe { System.dealloc(p, layout) };
        bump(-(layout.size() as i64));
    }
    unsafe fn realloc(&self, p: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the caller upholds the `GlobalAlloc` pointer/layout contract.
        let r = unsafe { System.realloc(p, layout, new_size) };
        if !r.is_null() {
            ALLOCS.fetch_add(1, Ordering::Relaxed);
            if new_size > layout.size() {
                ALLOC_BYTES.fetch_add((new_size - layout.size()) as u64, Ordering::Relaxed);
            }
            bump(new_size as i64 - layout.size() as i64);
        }
        r
    }
}

#[global_allocator]
static GLOBAL: Counting = Counting;

struct Region {
    peak: u64,
    allocs: u64,
    bytes: u64,
    live_start: u64,
}

fn region<T>(body: impl FnOnce() -> T) -> (T, Region) {
    let live_start = LIVE.load(Ordering::Relaxed);
    PEAK.store(live_start, Ordering::Relaxed);
    let a0 = ALLOCS.load(Ordering::Relaxed);
    let b0 = ALLOC_BYTES.load(Ordering::Relaxed);
    ON.store(true, Ordering::Relaxed);
    let out = body();
    ON.store(false, Ordering::Relaxed);
    let r = Region {
        peak: PEAK.load(Ordering::Relaxed).saturating_sub(live_start),
        allocs: ALLOCS.load(Ordering::Relaxed) - a0,
        bytes: ALLOC_BYTES.load(Ordering::Relaxed) - b0,
        live_start,
    };
    (out, r)
}

// ------------------------------------------------------------------- sink

/// A sequential, non-seek sink that accepts and discards every byte.
struct NullSink {
    written: u64,
}
impl Write for NullSink {
    fn write(&mut self, buf: &[u8]) -> std::io::Result<usize> {
        self.written += buf.len() as u64;
        Ok(buf.len())
    }
    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

// -------------------------------------------------------------- synthesis

/// Deterministic incompressible bytes (xorshift64*), so Deflate cannot shrink
/// the member and compressed size tracks the requested size.
fn incompressible(len: usize, seed: u64) -> Vec<u8> {
    let mut out = vec![0u8; len];
    let mut s = seed | 1;
    for chunk in out.chunks_mut(8) {
        s ^= s << 13;
        s ^= s >> 7;
        s ^= s << 17;
        let b = s.to_le_bytes();
        chunk.copy_from_slice(&b[..chunk.len()]);
    }
    out
}

/// A package with `small_members` XML-ish members plus one large media member.
fn build_archive(path: &std::path::Path, large_bytes: usize, method: CompressionMethod) {
    let mut w = StreamingArchiveWriter::new();
    w.write_deflated(
        "[Content_Types].xml",
        br#"<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>"#,
    )
    .unwrap();
    for i in 0..20 {
        let body = format!("<?xml version=\"1.0\"?><p id=\"{i}\">{}</p>", "text ".repeat(64));
        w.write_deflated(&format!("word/part{i}.xml"), body.as_bytes())
            .unwrap();
    }
    let media = incompressible(large_bytes, 0x9E3779B97F4A7C15);
    match method {
        CompressionMethod::Store => w.write_stored("word/media/big.bin", &media).unwrap(),
        _ => w.write_deflated("word/media/big.bin", &media).unwrap(),
    }
    let bytes = w.finish_to_bytes().unwrap();
    std::fs::write(path, &bytes).unwrap();
}

// ------------------------------------------------------------- scenarios

struct Outcome {
    peak: u64,
    allocs: u64,
    bytes: u64,
    output: u64,
    compressed: u64,
}

/// Scenario A: the production source-backed save shape.
fn scenario_copy_through(path: &std::path::Path) -> Outcome {
    let len = std::fs::metadata(path).unwrap().len();
    let file = std::fs::File::open(path).unwrap();
    let archive = IndexedArchive::from_reader(FileReader::from(file), len).unwrap();
    let mut scratch = vec![0u8; RECOMMENDED_BUFFER_SIZE];

    let replacement = b"<?xml version=\"1.0\"?><p id=\"0\">edited</p>".to_vec();
    let mut sink = NullSink { written: 0 };
    let mut large = 0u64;

    let (_, r) = region(|| {
        let index = archive.preservation_index(&mut scratch).unwrap();
        let mut plan = PreservationPlan::new();
        for entry in index.entries() {
            let name = entry.raw_name_bytes();
            if name == b"word/media/big.bin" {
                large = entry.compressed_size();
            }
            if name == b"word/part0.xml" {
                plan.push(PreservationAction::Regenerate {
                    id: entry.id(),
                    entry: RegeneratedEntry::new("word/part0.xml", replacement.clone())
                        .compression_method(CompressionMethod::Deflate),
                });
            } else {
                plan.push(PreservationAction::Copy(entry.id()));
            }
        }
        index.write_to(&plan, &mut sink).unwrap();
    });
    Outcome { peak: r.peak, allocs: r.allocs, bytes: r.bytes, output: sink.written, compressed: large }
}

/// Scenario B: the cross-document precompressed-token copy shape.
fn scenario_precompressed_token(path: &std::path::Path) -> Outcome {
    let len = std::fs::metadata(path).unwrap().len();
    let file = std::fs::File::open(path).unwrap();
    let archive = IndexedArchive::from_reader(FileReader::from(file), len).unwrap();
    let mut scratch = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    let entry_id = archive.entry_id("word/media/big.bin").unwrap();

    let mut sink = NullSink { written: 0 };
    let mut large = 0u64;

    let (_, r) = region(|| {
        let (token, _decoded) = archive
            .read_entry_precompressed_and_decoded_with_progress(entry_id, |_| {
                Ok::<(), std::io::Error>(())
            })
            .unwrap();
        large = token.compressed_size();
        let index = archive.preservation_index(&mut scratch).unwrap();
        let mut plan = PreservationPlan::new();
        for entry in index.entries() {
            if entry.raw_name_bytes() == b"word/media/big.bin" {
                plan.push(PreservationAction::Omit(entry.id()));
            } else {
                plan.push(PreservationAction::Copy(entry.id()));
            }
        }
        plan.try_append(RegeneratedEntry::new_precompressed_shared(
            "word/media/big.bin",
            token,
        ))
        .unwrap();
        index.write_to(&plan, &mut sink).unwrap();
    });
    Outcome { peak: r.peak, allocs: r.allocs, bytes: r.bytes, output: sink.written, compressed: large }
}

/// Scenario C: the `&[u8]` API named by the hypothesis. It has zero production
/// callers, and `VerifiedPrecompressedEntry::compressed_payload` is
/// `pub(crate)`, so an out-of-crate caller cannot even feed it from a verified
/// token. The only way to call it is to materialize the member's compressed
/// bytes yourself, which is what this scenario measures.
fn scenario_write_precompressed_slice(path: &std::path::Path) -> Outcome {
    let file = std::fs::File::open(path).unwrap();
    let mut locate = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    let archive = ZipArchive::from_file(file, &mut locate).unwrap();

    let mut target = None;
    let mut scan = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    let mut entries = archive.entries(&mut scan);
    while let Some(record) = entries.next_entry().unwrap() {
        if record.file_path().as_ref() == b"word/media/big.bin" {
            target = Some((
                record.wayfinder(),
                record.crc32(),
                record.uncompressed_size_hint(),
                record.compression_method(),
            ));
        }
    }
    let (wayfinder, crc, uncompressed, method) = target.unwrap();
    let (start, end) = archive.get_entry(wayfinder).unwrap().compressed_data_range();

    let mut sink = NullSink { written: 0 };
    let mut large = 0u64;

    let (_, r) = region(|| {
        // What any caller of the `&[u8]` entry point must do first.
        let len = usize::try_from(end - start).unwrap();
        let mut compressed = Vec::new();
        compressed.try_reserve_exact(len).unwrap();
        compressed.resize(len, 0u8);
        archive.get_ref().read_exact_at(&mut compressed, start).unwrap();
        large = compressed.len() as u64;

        let mut writer = ZipArchiveWriter::new(&mut sink);
        writer
            .write_precompressed_file("word/media/big.bin", method, crc, uncompressed, &compressed)
            .unwrap();
        writer.finish().unwrap();
    });
    Outcome { peak: r.peak, allocs: r.allocs, bytes: r.bytes, output: sink.written, compressed: large }
}

// ------------------------------------------------------------------- main

/// A `Read` that emits deterministic incompressible bytes without ever
/// materializing them, so the generator can build a ZIP64-scale member.
struct Noise {
    remaining: u64,
    state: u64,
}
impl std::io::Read for Noise {
    fn read(&mut self, buf: &mut [u8]) -> std::io::Result<usize> {
        if self.remaining == 0 {
            return Ok(0);
        }
        let n = buf.len().min(usize::try_from(self.remaining).unwrap_or(usize::MAX));
        for chunk in buf[..n].chunks_mut(8) {
            self.state ^= self.state << 13;
            self.state ^= self.state >> 7;
            self.state ^= self.state << 17;
            let b = self.state.to_le_bytes();
            chunk.copy_from_slice(&b[..chunk.len()]);
        }
        self.remaining -= n as u64;
        Ok(n)
    }
}

/// Finite ceilings large enough for a ZIP64-scale member. Deliberately NOT
/// `ArchiveLimits::UNBOUNDED`: every zip-bomb and size check stays enforced.
fn big_limits() -> ArchiveLimits {
    ArchiveLimits {
        max_files: 100_000,
        max_member_name_bytes: 4 * 1024,
        max_metadata_bytes: 64 * 1024 * 1024,
        max_compressed_size: 8 * 1024 * 1024 * 1024,
        max_entry_size: 8 * 1024 * 1024 * 1024,
        max_total_size: 16 * 1024 * 1024 * 1024,
    }
}

/// Scenario A against an arbitrary archive under explicit large-but-finite
/// limits, replacing nothing (pure copy-all).
fn scenario_copy_all_big(path: &std::path::Path) -> Outcome {
    let len = std::fs::metadata(path).unwrap().len();
    let file = std::fs::File::open(path).unwrap();
    let archive =
        IndexedArchive::from_reader_with_limits(FileReader::from(file), len, big_limits()).unwrap();
    let mut scratch = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    let mut sink = NullSink { written: 0 };
    let mut largest = 0u64;
    let (_, r) = region(|| {
        let index = archive
            .preservation_index_with_limits(&mut scratch, big_limits())
            .unwrap();
        let mut plan = PreservationPlan::new();
        for entry in index.entries() {
            largest = largest.max(entry.compressed_size());
            plan.push(PreservationAction::Copy(entry.id()));
        }
        index.write_to(&plan, &mut sink).unwrap();
    });
    Outcome { peak: r.peak, allocs: r.allocs, bytes: r.bytes, output: sink.written, compressed: largest }
}

/// Scenario A against an arbitrary archive, replacing nothing (pure copy-all).
fn scenario_copy_all(path: &std::path::Path) -> Outcome {
    let len = std::fs::metadata(path).unwrap().len();
    let file = std::fs::File::open(path).unwrap();
    let archive = IndexedArchive::from_reader(FileReader::from(file), len).unwrap();
    let mut scratch = vec![0u8; RECOMMENDED_BUFFER_SIZE];
    let mut sink = NullSink { written: 0 };
    let mut largest = 0u64;
    let (_, r) = region(|| {
        let index = archive.preservation_index(&mut scratch).unwrap();
        let mut plan = PreservationPlan::new();
        for entry in index.entries() {
            largest = largest.max(entry.compressed_size());
            plan.push(PreservationAction::Copy(entry.id()));
        }
        index.write_to(&plan, &mut sink).unwrap();
    });
    Outcome { peak: r.peak, allocs: r.allocs, bytes: r.bytes, output: sink.written, compressed: largest }
}

fn build_many(path: &std::path::Path, count: usize, each: usize) {
    let mut w = StreamingArchiveWriter::new();
    let body = incompressible(each, 0x243F6A8885A308D3);
    for i in 0..count {
        w.write_stored(&format!("word/media/m{i:05}.bin"), &body).unwrap();
    }
    std::fs::write(path, w.finish_to_bytes().unwrap()).unwrap();
}

fn main() {
    let mut args = std::env::args().skip(1);
    let mode = args.next().expect("usage: probe <mode> ...");
    let dir = args.next().expect("workdir");
    let dir = std::path::Path::new(&dir);
    std::fs::create_dir_all(dir).unwrap();

    match mode.as_str() {
        // Axis 1: peak as a function of the largest member.
        "sizes" => {
            let sizes: Vec<usize> = vec![
                64 * 1024, 256 * 1024, 1024 * 1024,
                4 * 1024 * 1024, 16 * 1024 * 1024, 64 * 1024 * 1024,
            ];
            println!("method,member_bytes,scenario,archive_bytes,member_compressed,region_peak_bytes,allocations,allocated_bytes,output_bytes");
            for method in [CompressionMethod::Store, CompressionMethod::Deflate] {
                let mname = if method == CompressionMethod::Store { "store" } else { "deflate" };
                for &size in &sizes {
                    let path = dir.join(format!("pkg-{mname}-{size}.zip"));
                    build_archive(&path, size, method);
                    let archive_bytes = std::fs::metadata(&path).unwrap().len();
                    for (label, f) in [
                        ("A_copy_through", scenario_copy_through as fn(&std::path::Path) -> Outcome),
                        ("B_precompressed_token", scenario_precompressed_token),
                        ("C_write_precompressed_slice", scenario_write_precompressed_slice),
                    ] {
                        let o = f(&path);
                        println!("{mname},{size},{label},{archive_bytes},{},{},{},{},{}",
                            o.compressed, o.peak, o.allocs, o.bytes, o.output);
                    }
                    std::fs::remove_file(&path).unwrap();
                }
            }
        }
        // Axis 1b: the same size series with a pure copy-all plan (no edited
        // member), which removes the one-off Deflate encoder state.
        "sizes-copyall" => {
            let sizes: Vec<usize> = vec![
                64 * 1024, 256 * 1024, 1024 * 1024,
                4 * 1024 * 1024, 16 * 1024 * 1024, 64 * 1024 * 1024,
            ];
            println!("method,member_bytes,scenario,archive_bytes,largest_member,region_peak_bytes,allocations,allocated_bytes,output_bytes");
            for method in [CompressionMethod::Store, CompressionMethod::Deflate] {
                let mname = if method == CompressionMethod::Store { "store" } else { "deflate" };
                for &size in &sizes {
                    let path = dir.join(format!("ca-{mname}-{size}.zip"));
                    build_archive(&path, size, method);
                    let archive_bytes = std::fs::metadata(&path).unwrap().len();
                    let o = scenario_copy_all(&path);
                    println!("{mname},{size},A_copy_all,{archive_bytes},{},{},{},{},{}",
                        o.compressed, o.peak, o.allocs, o.bytes, o.output);
                    std::fs::remove_file(&path).unwrap();
                }
            }
        }
        // Axis 2: peak as a function of the member COUNT, member size fixed.
        "counts" => {
            println!("member_count,member_bytes,scenario,archive_bytes,largest_member,region_peak_bytes,allocations,allocated_bytes,output_bytes");
            for &count in &[20usize, 100, 500, 2000, 8000] {
                let path = dir.join(format!("many-{count}.zip"));
                build_many(&path, count, 4096);
                let archive_bytes = std::fs::metadata(&path).unwrap().len();
                let o = scenario_copy_all(&path);
                println!("{count},4096,A_copy_all,{archive_bytes},{},{},{},{},{}",
                    o.compressed, o.peak, o.allocs, o.bytes, o.output);
                std::fs::remove_file(&path).unwrap();
            }
        }
        // A real corpus fixture, copied through verbatim.
        "fixture" => {
            println!("fixture,scenario,archive_bytes,largest_member,region_peak_bytes,allocations,allocated_bytes,output_bytes");
            for f in args {
                let path = std::path::PathBuf::from(&f);
                let archive_bytes = std::fs::metadata(&path).unwrap().len();
                let o = scenario_copy_all(&path);
                println!("{},A_copy_all,{archive_bytes},{},{},{},{},{}",
                    path.file_name().unwrap().to_string_lossy(),
                    o.compressed, o.peak, o.allocs, o.bytes, o.output);
            }
        }
        // A single member past the ZIP64 4 GiB promotion boundary.
        "zip64" => {
            let big: u64 = 4 * 1024 * 1024 * 1024 + 64 * 1024 * 1024;
            let path = dir.join("zip64-huge.zip");
            if !path.exists() {
                let out = std::fs::File::create(&path).unwrap();
                let mut limits = StreamingArchiveLimits::default();
                limits.max_compressed_size = 8 * 1024 * 1024 * 1024;
                limits.max_entry_size = 8 * 1024 * 1024 * 1024;
                limits.max_total_size = 16 * 1024 * 1024 * 1024;
                limits.max_output_bytes = 16 * 1024 * 1024 * 1024;
                let mut w = StreamingArchiveWriter::with_writer_and_limits(
                    std::io::BufWriter::new(out), limits);
                w.write_deflated("[Content_Types].xml", b"<Types/>").unwrap();
                w.write_stored_stream("word/media/huge.bin",
                    Noise { remaining: big, state: 0x9E3779B97F4A7C15 }).unwrap();
                w.write_deflated("word/part0.xml", b"<p>a</p>").unwrap();
                w.finish().unwrap().into_inner().unwrap().sync_all().unwrap();
            }
            let archive_bytes = std::fs::metadata(&path).unwrap().len();
            println!("member_bytes,scenario,archive_bytes,largest_member,region_peak_bytes,allocations,allocated_bytes,output_bytes");
            let o = scenario_copy_all_big(&path);
            println!("{big},A_copy_all,{archive_bytes},{},{},{},{},{}",
                o.compressed, o.peak, o.allocs, o.bytes, o.output);
        }
        other => panic!("unknown mode {other}"),
    }
}
