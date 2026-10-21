use std::{hint::black_box, time::Instant};
use litchi_iwa_core::{Archive, ArchiveObject, RawMessage, SnappyStream};
use sha2::{Digest, Sha256};
#[cfg(feature = "allocations")]
mod allocation {
    use std::{alloc::{GlobalAlloc, Layout, System}, sync::atomic::{AtomicUsize, Ordering::Relaxed}};
    pub static LIVE: AtomicUsize = AtomicUsize::new(0);
    pub static PEAK: AtomicUsize = AtomicUsize::new(0);
    pub static BYTES: AtomicUsize = AtomicUsize::new(0);
    pub static CALLS: AtomicUsize = AtomicUsize::new(0);
    pub struct Allocator;
    unsafe impl GlobalAlloc for Allocator {
        unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
            let pointer = unsafe { System.alloc(layout) };
            if !pointer.is_null() { add(layout.size()); }
            pointer
        }
        unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
            LIVE.fetch_sub(layout.size(), Relaxed);
            unsafe { System.dealloc(pointer, layout); }
        }
        unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
            let result = unsafe { System.realloc(pointer, layout, size) };
            if !result.is_null() { LIVE.fetch_sub(layout.size(), Relaxed); add(size); }
            result
        }
    }
    fn add(size: usize) {
        let live = LIVE.fetch_add(size, Relaxed) + size;
        PEAK.fetch_max(live, Relaxed);
        BYTES.fetch_add(size, Relaxed);
        CALLS.fetch_add(1, Relaxed);
    }
    pub fn start() -> usize {
        let live = LIVE.load(Relaxed);
        PEAK.store(live, Relaxed); BYTES.store(0, Relaxed); CALLS.store(0, Relaxed);
        live
    }
    pub fn result(base: usize) -> (usize, usize, usize, usize) {
        (PEAK.load(Relaxed) - base, LIVE.load(Relaxed).saturating_sub(base), BYTES.load(Relaxed), CALLS.load(Relaxed))
    }
}
#[cfg(feature = "allocations")]
#[global_allocator]
static ALLOCATOR: allocation::Allocator = allocation::Allocator;

#[cfg(feature="before")]
fn compress(data: &[u8]) -> Vec<u8> { SnappyStream::compress(data).unwrap() }
#[cfg(not(feature="before"))]
fn compress(data: &[u8]) -> Vec<u8> { SnappyStream::compress(data, Default::default()).unwrap() }
#[cfg(feature="before")]
fn decompress(data: &[u8]) -> Vec<u8> { SnappyStream::decompress(data).unwrap().into_bytes() }
#[cfg(not(feature="before"))]
fn decompress(data: &[u8]) -> Vec<u8> { SnappyStream::decompress(data, Default::default()).unwrap().into_bytes() }
#[cfg(feature="before")]
fn bound(size: usize) -> usize { SnappyStream::maximum_compressed_len(size).unwrap() }
#[cfg(not(feature="before"))]
fn bound(size: usize) -> usize { SnappyStream::maximum_compressed_len(size, Default::default()).unwrap() }
#[cfg(feature="before")]
fn object(id: u64, messages: Vec<RawMessage>) -> ArchiveObject { ArchiveObject::new(id,messages).unwrap() }
#[cfg(not(feature="before"))]
fn object(id: u64, messages: Vec<RawMessage>) -> ArchiveObject { ArchiveObject::new(id,messages,Default::default()).unwrap() }
#[cfg(feature="before")]
fn archive_bytes(archive: &Archive) -> Vec<u8> { archive.to_bytes().unwrap() }
#[cfg(not(feature="before"))]
fn archive_bytes(archive: &Archive) -> Vec<u8> { archive.to_bytes(Default::default()).unwrap() }
#[cfg(feature="before")]
fn parse(bytes: &[u8]) -> Archive { Archive::parse(bytes).unwrap() }
#[cfg(not(feature="before"))]
fn parse(bytes: &[u8]) -> Archive { Archive::parse(bytes,Default::default()).unwrap() }
#[cfg(feature="before")]
fn framing(archive: &Archive, bytes: &[u8]) { archive.validate_canonical_object_framing(bytes).unwrap(); }
#[cfg(not(feature="before"))]
fn framing(archive: &Archive, bytes: &[u8]) { archive.validate_canonical_object_framing(bytes,Default::default()).unwrap(); }
#[cfg(feature="before")]
fn catalog(bytes: &[u8]) -> litchi_iwa_archive::package::Catalog { litchi_iwa_archive::package::Catalog::from_bytes(bytes).unwrap() }
#[cfg(not(feature="before"))]
fn catalog(bytes: &[u8]) -> litchi_iwa_archive::package::Catalog { litchi_iwa_archive::package::Catalog::from_bytes(bytes,Default::default()).unwrap() }
#[cfg(feature="before")]
fn copy(catalog: &litchi_iwa_archive::package::Catalog) -> Vec<u8> { catalog.to_bytes().unwrap() }
#[cfg(not(feature="before"))]
fn copy(catalog: &litchi_iwa_archive::package::Catalog) -> Vec<u8> { catalog.to_bytes(Default::default()).unwrap() }
#[cfg(feature="before")]
fn write(catalog: &litchi_iwa_archive::package::Catalog, sink: &mut Vec<u8>) { catalog.write_to(sink).unwrap(); }
#[cfg(not(feature="before"))]
fn write(catalog: &litchi_iwa_archive::package::Catalog, sink: &mut Vec<u8>) { catalog.write_to(sink,Default::default()).unwrap(); }
fn main() {
 let args:Vec<_>=std::env::args().collect(); let route=&args[1]; let iterations:usize=args[2].parse().unwrap();
 let size=if route.ends_with("small") {1024} else {4*1024*1024};
 let mut seed=0x123456789abcdefu64;
 let input:Vec<u8>=(0..size).map(|i| { seed^=seed<<13; seed^=seed>>7; seed^=seed<<17; if route.ends_with("noise") || route.starts_with("copy-") || route.starts_with("write-") {seed as u8} else {(i%31) as u8} }).collect();
 let encoded=compress(&input);
 let original=Archive { objects: (0..512).map(|id| object(id, vec![RawMessage { type_: 1, data: vec![42;1024] },RawMessage { type_:2, data:vec![5;1024] }])).collect() };
 let archive_source=archive_bytes(&original);let archive=parse(&archive_source);
 let zip=litchi_iwa_archive::package::to_bytes([("Data/blob",input.as_slice())], Default::default()).unwrap();let catalog=catalog(&zip);let mut sink=Vec::with_capacity(zip.len());
 let mut run=|| -> Vec<u8> { if route.starts_with("compress-") {compress(black_box(&input))} else if route.starts_with("decompress-") {decompress(black_box(&encoded))} else if route.starts_with("copy-") {copy(black_box(&catalog))} else if route.starts_with("write-") {sink.clear();write(black_box(&catalog),black_box(&mut sink));black_box(sink.as_slice());Vec::new()} else if route=="bound-large" {black_box(bound(black_box(size)));Vec::new()} else if route=="framing" {framing(black_box(&archive),black_box(&archive_source));Vec::new()} else {panic!("unknown route")} };
 let expected=run();let checksum=Sha256::digest(&expected).iter().map(|byte| format!("{:02x}",byte)).collect::<String>();
 assert_eq!(decompress(&encoded),input);
 let expected_bound=bound(input.len());
 #[cfg(feature="allocations")] let allocation_base=allocation::start();
 let started=Instant::now();for _ in 0..iterations {drop(black_box(run()));}let ns=started.elapsed().as_nanos();
 #[cfg(feature="allocations")] let metrics=allocation::result(allocation_base);
 #[cfg(not(feature="allocations"))] let metrics=(0usize,0usize,0usize,0usize);
 if route.starts_with("write-") {assert_eq!(sink,zip);}
 println!("{{\"route\":\"{}\",\"iterations\":{},\"ns\":{},\"checksum\":\"{}\",\"bytes\":{},\"bound\":{},\"peak\":{},\"retained\":{},\"allocated\":{},\"allocations\":{}}}", route, iterations, ns,checksum,expected.len(),expected_bound,metrics.0,metrics.1,metrics.2,metrics.3);
}
