//! End-to-end peak-heap probe for the two production OOXML save paths
//! (change 0578, stage 1). Both save to a sequential, non-seek sink.
//!
//! A: OpcPackage::open + get_part_mut().set_blob + PackageWriter::write_to_stream
//! B: SourceBackedPackage::from_path + write_part_overlay_to_stream

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::Write;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, SourceBackedPackage};

static LIVE: AtomicU64 = AtomicU64::new(0);
static PEAK: AtomicU64 = AtomicU64::new(0);
static ALLOCS: AtomicU64 = AtomicU64::new(0);
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
// SAFETY: every method delegates to `std::alloc::System`; counter updates are
// side-effect-only and change neither pointer, layout, nor ownership.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, l: Layout) -> *mut u8 {
        // SAFETY: the caller upholds the `GlobalAlloc` layout contract.
        let p = unsafe { System.alloc(l) };
        if !p.is_null() { ALLOCS.fetch_add(1, Ordering::Relaxed); bump(l.size() as i64); }
        p
    }
    unsafe fn alloc_zeroed(&self, l: Layout) -> *mut u8 {
        // SAFETY: the caller upholds the `GlobalAlloc` layout contract.
        let p = unsafe { System.alloc_zeroed(l) };
        if !p.is_null() { ALLOCS.fetch_add(1, Ordering::Relaxed); bump(l.size() as i64); }
        p
    }
    unsafe fn dealloc(&self, p: *mut u8, l: Layout) {
        // SAFETY: the caller upholds the `GlobalAlloc` pointer/layout contract.
        unsafe { System.dealloc(p, l) };
        bump(-(l.size() as i64));
    }
    unsafe fn realloc(&self, p: *mut u8, l: Layout, n: usize) -> *mut u8 {
        // SAFETY: the caller upholds the `GlobalAlloc` pointer/layout contract.
        let r = unsafe { System.realloc(p, l, n) };
        if !r.is_null() { ALLOCS.fetch_add(1, Ordering::Relaxed); bump(n as i64 - l.size() as i64); }
        r
    }
}
#[global_allocator]
static GLOBAL: Counting = Counting;

fn region<T>(body: impl FnOnce() -> T) -> (T, u64, u64) {
    let start = LIVE.load(Ordering::Relaxed);
    PEAK.store(start, Ordering::Relaxed);
    let a0 = ALLOCS.load(Ordering::Relaxed);
    ON.store(true, Ordering::Relaxed);
    let out = body();
    ON.store(false, Ordering::Relaxed);
    (out, PEAK.load(Ordering::Relaxed).saturating_sub(start),
     ALLOCS.load(Ordering::Relaxed) - a0)
}

/// A sequential, non-seek sink that accepts and discards every byte while
/// folding an FNV-1a digest of the whole output. Hashing allocates nothing, so
/// it does not disturb the peak-heap measurement.
struct NullSink { written: u64, digest: u64 }
impl NullSink {
    fn new() -> Self { Self { written: 0, digest: 0xcbf2_9ce4_8422_2325 } }
}
impl Write for NullSink {
    fn write(&mut self, b: &[u8]) -> std::io::Result<usize> {
        self.written += b.len() as u64;
        for byte in b {
            self.digest ^= u64::from(*byte);
            self.digest = self.digest.wrapping_mul(0x1000_0000_01b3);
        }
        Ok(b.len())
    }
    fn flush(&mut self) -> std::io::Result<()> { Ok(()) }
}

fn incompressible(len: usize, seed: u64) -> Vec<u8> {
    let mut out = vec![0u8; len];
    let mut s = seed | 1;
    for c in out.chunks_mut(8) {
        s ^= s << 13; s ^= s >> 7; s ^= s << 17;
        let b = s.to_le_bytes();
        c.copy_from_slice(&b[..c.len()]);
    }
    out
}

const MEDIA: &str = "/word/media/probe-big.bin";
const SMALL: &str = "/word/media/probe-small.bin";
const CTYPE: &str = "application/vnd.litchi.perf.probe";

/// Build a fixture from a real DOCX by adding one large and one small binary
/// part, so the package carries a big unchanged member on save.
fn make_fixture(base: &std::path::Path, out: &std::path::Path, media_bytes: usize) {
    let mut package = OpcPackage::open(base).expect("open base");
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(MEDIA).unwrap(),
            CTYPE.to_owned(),
            incompressible(media_bytes, 0x9E3779B97F4A7C15),
        )))
        .expect("add media part");
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new(SMALL).unwrap(),
            CTYPE.to_owned(),
            incompressible(4096, 0x243F6A8885A308D3),
        )))
        .expect("add small part");
    PackageWriter::write(out, &package).expect("write fixture");
}

fn main() {
    let mut a = std::env::args().skip(1);
    let base = a.next().expect("base docx");
    let dir = a.next().expect("workdir");
    let base = std::path::Path::new(&base);
    let dir = std::path::Path::new(&dir);
    std::fs::create_dir_all(dir).unwrap();

    println!("media_bytes,path,archive_bytes,region_peak_bytes,allocations,output_bytes,output_fnv1a64");
    for &media in &[64 * 1024usize, 1024 * 1024, 4 * 1024 * 1024, 16 * 1024 * 1024, 64 * 1024 * 1024] {
        let fixture = dir.join(format!("fx-{media}.docx"));
        make_fixture(base, &fixture, media);
        let archive_bytes = std::fs::metadata(&fixture).unwrap().len();
        let replacement = incompressible(4096, 0xB7E151628AED2A6B);

        // Path B — source-backed, positional source.
        {
            let mut sink = NullSink::new();
            let (_, peak, allocs) = region(|| {
                let package = SourceBackedPackage::from_path(&fixture).expect("source-backed open");
                package
                    .write_part_overlay_to_stream(&mut sink, &PackURI::new(SMALL).unwrap(), replacement.clone())
                    .expect("overlay publish");
            });
            println!("{media},B_source_backed,{archive_bytes},{peak},{allocs},{},{:016x}", sink.written, sink.digest);
        }

        // Path A — OpcPackage + PackageWriter.
        {
            let mut sink = NullSink::new();
            let (_, peak, allocs) = region(|| {
                let mut package = OpcPackage::open(&fixture).expect("eager open");
                package
                    .get_part_mut(&PackURI::new(SMALL).unwrap())
                    .expect("small part")
                    .set_blob(replacement.clone());
                PackageWriter::write_to_stream(&mut sink, &package).expect("eager publish");
            });
            println!("{media},A_package_writer,{archive_bytes},{peak},{allocs},{},{:016x}", sink.written, sink.digest);
        }

        std::fs::remove_file(&fixture).unwrap();
    }
}
