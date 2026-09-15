//! Stage-1 attribution probe for change 0581. Extends change 0578's
//! `opc-save-peak-probe.rs` (same allocator, same sink, same fixture builder)
//! by splitting each save path into its open and its publish region, and by
//! reporting the bytes each region *retains* as well as its peak.
//!
//! A: OpcPackage::open  + PackageWriter::write_to_stream   (eager package)
//! B: SourceBackedPackage::from_path + write_part_overlay_to_stream

use std::alloc::{GlobalAlloc, Layout, System};
use std::io::Write;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, SourceBackedPackage, SourceTopologyPlan};

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

struct Region { peak: u64, allocs: u64, retained: i64 }

/// Measure a region. The returned value stays alive in the caller, so
/// `retained` is the heap the region's product still holds afterwards.
fn region<T>(body: impl FnOnce() -> T) -> (T, Region) {
    let start = LIVE.load(Ordering::Relaxed);
    PEAK.store(start, Ordering::Relaxed);
    let a0 = ALLOCS.load(Ordering::Relaxed);
    ON.store(true, Ordering::Relaxed);
    let out = body();
    let end = LIVE.load(Ordering::Relaxed);
    ON.store(false, Ordering::Relaxed);
    let r = Region {
        peak: PEAK.load(Ordering::Relaxed).saturating_sub(start),
        allocs: ALLOCS.load(Ordering::Relaxed) - a0,
        retained: end as i64 - start as i64,
    };
    (out, r)
}

/// A sequential, non-seek sink that accepts and discards every byte while
/// folding an FNV-1a digest of the whole output. Hashing allocates nothing.
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

/// change 0578's fixture builder: one large and one small binary part added to
/// a real DOCX.
fn make_fixture(base: &std::path::Path, out: &std::path::Path, media_bytes: usize) {
    let mut package = OpcPackage::open(base).expect("open base");
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(MEDIA).unwrap(), CTYPE.to_owned(),
        incompressible(media_bytes, 0x9E3779B97F4A7C15)))).expect("add media part");
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(SMALL).unwrap(), CTYPE.to_owned(),
        incompressible(4096, 0x243F6A8885A308D3)))).expect("add small part");
    PackageWriter::write(out, &package).expect("write fixture");
}

/// Same base DOCX, but `count` small binary parts and no large one: isolates
/// scaling in the member count from scaling in the largest member.
fn make_count_fixture(base: &std::path::Path, out: &std::path::Path, count: usize, each: usize) {
    let mut package = OpcPackage::open(base).expect("open base");
    for i in 0..count {
        let uri = format!("/word/media/probe-{i:05}.bin");
        package.try_add_part(Box::new(BlobPart::new(
            PackURI::new(&uri).unwrap(), CTYPE.to_owned(),
            incompressible(each, 0x243F6A8885A308D3 ^ i as u64)))).expect("add part");
    }
    package.try_add_part(Box::new(BlobPart::new(
        PackURI::new(SMALL).unwrap(), CTYPE.to_owned(),
        incompressible(4096, 0xB7E151628AED2A6B)))).expect("add small part");
    PackageWriter::write(out, &package).expect("write fixture");
}

/// Drive both paths over one fixture, editing the small probe part.
/// Emits one CSV row per (path, phase).
fn drive(tag: &str, axis_value: u64, fixture: &std::path::Path) {
    let archive_bytes = std::fs::metadata(fixture).unwrap().len();
    let replacement = incompressible(4096, 0xB7E151628AED2A6B);
    let uri = PackURI::new(SMALL).unwrap();

    // Path B — source-backed, positional source.
    {
        let mut sink = NullSink::new();
        let (package, open_r) = region(|| SourceBackedPackage::from_path(fixture).expect("B open"));
        // NOTE: the source-backed publish takes `self` by value, so this
        // region also covers dropping the package. There is no separate drop
        // row for B.
        let (_, save_r) = region(|| {
            package.write_part_overlay_to_stream(&mut sink, &uri, replacement.clone()).expect("B publish");
        });
        println!("{tag},{axis_value},B_source_backed,open,{archive_bytes},{},{},{},0,0",
            open_r.peak, open_r.allocs, open_r.retained);
        println!("{tag},{axis_value},B_source_backed,save_and_drop,{archive_bytes},{},{},{},{},{:016x}",
            save_r.peak, save_r.allocs, save_r.retained, sink.written, sink.digest);
    }

    // Path A — eager OpcPackage + PackageWriter.
    {
        let mut sink = NullSink::new();
        let (mut package, open_r) = region(|| OpcPackage::open(fixture).expect("A open"));
        let (_, save_r) = region(|| {
            package.get_part_mut(&uri).expect("small part").set_blob(replacement.clone());
            PackageWriter::write_to_stream(&mut sink, &package).expect("A publish");
        });
        let (_, drop_r) = region(|| drop(package));
        println!("{tag},{axis_value},A_package_writer,open,{archive_bytes},{},{},{},0,0",
            open_r.peak, open_r.allocs, open_r.retained);
        println!("{tag},{axis_value},A_package_writer,save,{archive_bytes},{},{},{},{},{:016x}",
            save_r.peak, save_r.allocs, save_r.retained, sink.written, sink.digest);
        println!("{tag},{axis_value},A_package_writer,drop,{archive_bytes},{},{},{},0,0",
            drop_r.peak, drop_r.allocs, drop_r.retained);
    }
}

/// Unedited open-and-save of a real corpus fixture through both paths.
fn drive_real(name: &str, fixture: &std::path::Path) {
    let archive_bytes = std::fs::metadata(fixture).unwrap().len();
    {
        let mut sink = NullSink::new();
        let (package, open_r) = region(|| SourceBackedPackage::from_path(fixture).expect("B open"));
        let (_, save_r) = region(|| {
            package.write_topology_to_stream(&mut sink, SourceTopologyPlan::new()).expect("B publish")
        });
        println!("real,{name},B_source_backed,open,{archive_bytes},{},{},{},0,0",
            open_r.peak, open_r.allocs, open_r.retained);
        println!("real,{name},B_source_backed,save_and_drop,{archive_bytes},{},{},{},{},{:016x}",
            save_r.peak, save_r.allocs, save_r.retained, sink.written, sink.digest);
    }
    {
        let mut sink = NullSink::new();
        let (package, open_r) = region(|| OpcPackage::open(fixture).expect("A open"));
        let (_, save_r) = region(|| PackageWriter::write_to_stream(&mut sink, &package).expect("A publish"));
        println!("real,{name},A_package_writer,open,{archive_bytes},{},{},{},0,0",
            open_r.peak, open_r.allocs, open_r.retained);
        println!("real,{name},A_package_writer,save,{archive_bytes},{},{},{},{},{:016x}",
            save_r.peak, save_r.allocs, save_r.retained, sink.written, sink.digest);
        drop(package);
    }
}

fn main() {
    let mut a = std::env::args().skip(1);
    let base = a.next().expect("base docx");
    let dir = a.next().expect("workdir");
    let mode = a.next().unwrap_or_else(|| "all".to_owned());
    let base = std::path::Path::new(&base);
    let dir = std::path::Path::new(&dir);
    std::fs::create_dir_all(dir).unwrap();

    println!("axis,axis_value,path,phase,archive_bytes,region_peak_bytes,allocations,retained_bytes,output_bytes,output_fnv1a64");

    if mode == "all" || mode == "size" {
        for &media in &[64 * 1024usize, 256 * 1024, 1024 * 1024, 4 * 1024 * 1024,
                        16 * 1024 * 1024, 64 * 1024 * 1024, 128 * 1024 * 1024] {
            let fixture = dir.join(format!("fx-size-{media}.docx"));
            make_fixture(base, &fixture, media);
            drive("size", media as u64, &fixture);
            std::fs::remove_file(&fixture).unwrap();
        }
    }

    if mode == "all" || mode == "count" {
        for &count in &[0usize, 20, 100, 500, 2000, 8000] {
            let fixture = dir.join(format!("fx-count-{count}.docx"));
            make_count_fixture(base, &fixture, count, 4096);
            drive("count", count as u64, &fixture);
            std::fs::remove_file(&fixture).unwrap();
        }
    }

    if mode == "all" || mode == "real" {
        for arg in a {
            let p = std::path::PathBuf::from(&arg);
            let name = p.file_name().unwrap().to_string_lossy().into_owned();
            drive_real(&name, &p);
        }
    }
}
