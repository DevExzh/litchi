//! Change 0628 probe C: deterministic counts for OPC open and save.
//!
//! The change alters only which of several equally matching relationships
//! `Relationships::get_or_add` / `get_or_add_ext_rel` reuses, and neither is
//! called by open or by publication. Every number below must therefore be
//! identical on the before and after legs.
//!
//! Counted per fixture: positional source requests and bytes for a
//! source-backed open (a `ReadAt` provider that logs every call), allocations
//! and allocated bytes for the eager open and for `PackageWriter::to_bytes`,
//! and the published byte count.
//!
//! Usage: open_save_counts <fixture>...
use std::alloc::{GlobalAlloc, Layout, System};
use std::io;
use std::sync::Mutex;
use std::sync::atomic::{AtomicU64, Ordering};

struct CountingAlloc;
static ALLOCS: AtomicU64 = AtomicU64::new(0);
static ALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
unsafe impl GlobalAlloc for CountingAlloc {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        ALLOCS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.alloc(layout) }
    }
    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) }
    }
    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        ALLOCS.fetch_add(1, Ordering::Relaxed);
        ALLOC_BYTES.fetch_add(new_size as u64, Ordering::Relaxed);
        unsafe { System.realloc(pointer, layout, new_size) }
    }
}
#[global_allocator]
static GLOBAL: CountingAlloc = CountingAlloc;

fn snapshot() -> (u64, u64) {
    (
        ALLOCS.load(Ordering::Relaxed),
        ALLOC_BYTES.load(Ordering::Relaxed),
    )
}

use litchi_core::{ReadAt, SourceVersion};
use litchi_opc::{ReadLimits, SourceBackedPackage, SourceCacheLimits, SourceReadPolicy};

struct Counting {
    inner: Vec<u8>,
    calls: AtomicU64,
    bytes: AtomicU64,
    log: Mutex<Vec<(u64, u64)>>,
}

impl Counting {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            inner: bytes,
            calls: AtomicU64::new(0),
            bytes: AtomicU64::new(0),
            log: Mutex::new(Vec::new()),
        }
    }
}

impl ReadAt for Counting {
    fn len(&self) -> io::Result<u64> {
        Ok(self.inner.len() as u64)
    }
    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        self.calls.fetch_add(1, Ordering::SeqCst);
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let take = if start >= self.inner.len() {
            0
        } else {
            output.len().min(self.inner.len() - start)
        };
        if take > 0 {
            output[..take].copy_from_slice(&self.inner[start..start + take]);
        }
        self.bytes.fetch_add(take as u64, Ordering::SeqCst);
        self.log.lock().expect("log").push((offset, take as u64));
        Ok(take)
    }
    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x5a1, 1))
    }
}

fn main() {
    for fixture in std::env::args().skip(1) {
        let bytes = std::fs::read(&fixture).expect("fixture");
        println!("### {fixture}");
        println!("  source_bytes={}", bytes.len());

        // Source-backed open through a logging provider.
        let provider = std::sync::Arc::new(Counting::new(bytes.clone()));
        let dynamic: std::sync::Arc<dyn ReadAt + Send + Sync> = provider.clone();
        let before = snapshot();
        let source_backed =
            SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy(
                dynamic,
                ReadLimits::default(),
                SourceCacheLimits::default(),
                SourceReadPolicy::exact(),
            );
        let after = snapshot();
        match source_backed {
            Ok(package) => {
                let log = provider.log.lock().expect("log");
                let offsets: std::collections::BTreeSet<u64> =
                    log.iter().map(|(offset, _)| *offset).collect();
                println!(
                    "  [source-backed open] requests={} request_bytes={} distinct_offsets={} parts={} allocs={} alloc_bytes={}",
                    provider.calls.load(Ordering::SeqCst),
                    provider.bytes.load(Ordering::SeqCst),
                    offsets.len(),
                    package.iter_parts().count(),
                    after.0 - before.0,
                    after.1 - before.1
                );
            },
            Err(error) => println!("  [source-backed open] REFUSED {error}"),
        }

        // Eager open.
        let before = snapshot();
        let package = litchi_opc::OpcPackage::from_bytes(&bytes);
        let after = snapshot();
        let Ok(package) = package else {
            println!("  [eager open] REFUSED");
            continue;
        };
        let part_rels: usize = package.iter_parts().map(|part| part.rels().len()).sum();
        println!(
            "  [eager open] parts={} pkg_rels={} part_rels={} allocs={} alloc_bytes={}",
            package.part_count(),
            package.rels().len(),
            part_rels,
            after.0 - before.0,
            after.1 - before.1
        );

        // Publication.
        let before = snapshot();
        let published = litchi_opc::PackageWriter::to_bytes(&package);
        let after = snapshot();
        match published {
            Ok(output) => println!(
                "  [save] out_bytes={} allocs={} alloc_bytes={}",
                output.len(),
                after.0 - before.0,
                after.1 - before.1
            ),
            Err(error) => println!("  [save] REFUSED {error}"),
        }
    }
}
