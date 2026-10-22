use std::{hint::black_box, io::Write, time::Instant};
use litchi_odraw::{Record, RecordKind, prop::{Id, picture}};

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

fn fixture(large: bool) -> Vec<u8> {
    let name: Vec<u8> = "old".encode_utf16().chain([0]).flat_map(u16::to_le_bytes).collect();
    let opaque = vec![0xa5; if large { 2*1024*1024 } else { 24 }];
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&0x43_u16.to_le_bytes());
    bytes.extend_from_slice(&RecordKind::Opt.raw().to_le_bytes());
    bytes.extend_from_slice(&u32::try_from(24+name.len()+opaque.len()).unwrap().to_le_bytes());
    for (id, value) in [(0x0600, -7), (0x8000|Id::PictureFileName.raw(), name.len() as i32), (Id::BlipFlags.raw(), 0x8000_0009_u32 as i32), (0xc601, opaque.len() as i32)] {
        bytes.extend_from_slice(&id.to_le_bytes()); bytes.extend_from_slice(&value.to_le_bytes());
    }
    bytes.extend_from_slice(&name); bytes.extend_from_slice(&opaque); bytes
}

#[derive(Default)]
struct Sink { hash: u64, length: usize }
impl Write for Sink {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        // Rolling hash is independent of chunk boundaries. It reads every byte,
        // preventing optimization of the actual output work.
        for byte in bytes { self.hash = self.hash.wrapping_mul(31).wrapping_add(u64::from(*byte)); }
        self.length += bytes.len(); Ok(bytes.len())
    }
    fn flush(&mut self) -> std::io::Result<()> { Ok(()) }
}

#[cfg(feature = "before")]
fn open(bytes: &[u8]) -> picture::Snapshot<'_> {
    let (record, consumed) = Record::parse(bytes, 0).unwrap(); assert_eq!(consumed, bytes.len());
    picture::Snapshot::parse(&record).unwrap()
}
#[cfg(not(feature = "before"))]
fn open(bytes: &[u8]) -> picture::Package<'_> { picture::Package::open(bytes).unwrap() }

fn properties_probe(route: &str, iterations: usize) {
    use litchi_odraw::prop::Props;
    let count: u16 = route.rsplit('-').next().unwrap().parse().unwrap();
    let mut bytes = Vec::new();
    bytes.extend_from_slice(&((count << 4) | 3).to_le_bytes());
    bytes.extend_from_slice(&RecordKind::Opt.raw().to_le_bytes());
    bytes.extend_from_slice(&(u32::from(count)*6).to_le_bytes());
    for index in 0..count {
        bytes.extend_from_slice(&(0x600 + index).to_le_bytes());
        bytes.extend_from_slice(&i32::from(index).to_le_bytes());
    }
    let (record, _) = Record::parse(&bytes, 0).unwrap();
    let properties = Props::parse(&record).unwrap();
    #[cfg(feature = "allocations")]
    let base = allocation::start();
    let started = Instant::now();
    let mut checksum = 0u64;
    for _ in 0..iterations {
        if route.starts_with("props-open-") {
            checksum ^= black_box(Props::parse(black_box(&record)).unwrap()).len() as u64;
        } else {
            checksum ^= black_box(properties.get_int(black_box(Id::from(0x600 + count - 1)))).unwrap() as u64;
            checksum ^= black_box(properties.has(black_box(Id::from(0x1000)))) as u64;
        }
    }
    let ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let (peak, live, allocated, calls) = allocation::result(base);
    #[cfg(not(feature = "allocations"))]
    let (peak, live, allocated, calls) = (0,0,0,0);
    println!("{{\"route\":\"{route}\",\"iterations\":{iterations},\"ns\":{ns},\"checksum\":{checksum},\"peak\":{peak},\"retained\":{live},\"allocated\":{allocated},\"calls\":{calls}}}");
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let route = &args[1]; let iterations: usize = args[2].parse().unwrap();
    if route.starts_with("props-") { properties_probe(route, iterations); return; }
    let bytes = fixture(route.ends_with("large"));
    let source = open(&bytes);
    #[cfg(feature = "allocations")]
    let base = allocation::start();
    let started = Instant::now();
    let mut checksum = 0u64;
    let mut retained = Vec::new();
    for _ in 0..iterations {
        if route.starts_with("open-") {
            let package = open(black_box(&bytes));
            black_box(package);
            checksum ^= bytes.len() as u64;
        } else if route.starts_with("view-") {
            let (record, _) = Record::parse(black_box(&bytes), 0).unwrap();
            #[cfg(feature = "before")]
            let view = picture::Snapshot::parse(&record).unwrap();
            #[cfg(not(feature = "before"))]
            let view = picture::View::parse(&record, picture::Limits::new()).unwrap();
            black_box(view);
            checksum ^= bytes.len() as u64;
        } else if route.starts_with("metadata-") {
            checksum ^= black_box(source.metadata().unwrap().unwrap().flags().raw()) as u64;
        } else {
            let mut edit = source.edit();
            if !route.starts_with("noop-") { edit.set_name(black_box("new")).unwrap(); }
            let commit = black_box(edit.commit().unwrap());
            if route.starts_with("write-") {
                let mut sink = Sink::default();
                #[cfg(feature = "before")]
                sink.write_all(black_box(commit.snapshot().as_bytes())).unwrap();
                #[cfg(not(feature = "before"))]
                commit.package().write_to(&mut sink).unwrap();
                checksum ^= black_box(sink.hash ^ sink.length as u64);
            } else {
                #[cfg(feature = "before")]
                let noop = commit.patch().is_empty();
                #[cfg(not(feature = "before"))]
                let noop = commit.patch().is_noop();
                checksum ^= black_box(noop) as u64;
            }
            if route.starts_with("history-") { retained.push(commit); }
        }
    }
    let ns = started.elapsed().as_nanos();
    #[cfg(feature = "allocations")]
    let (peak, live, allocated, calls) = allocation::result(base);
    #[cfg(not(feature = "allocations"))]
    let (peak, live, allocated, calls) = (0,0,0,0);
    println!("{{\"route\":\"{route}\",\"iterations\":{iterations},\"ns\":{ns},\"checksum\":{checksum},\"peak\":{peak},\"retained\":{live},\"allocated\":{allocated},\"calls\":{calls}}}");
    black_box(retained);
}
