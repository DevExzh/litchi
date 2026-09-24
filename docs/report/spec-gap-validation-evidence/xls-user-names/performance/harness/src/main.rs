use std::alloc::{GlobalAlloc, Layout, System};
use std::hint::black_box;
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::Arc;
use std::time::{Duration, Instant};

use litchi_xls::{UserNamesLimits, UserNamesSnapshot as Snapshot};

struct CountingAllocator;

static ALLOCS: AtomicU64 = AtomicU64::new(0);
static DEALLOCS: AtomicU64 = AtomicU64::new(0);
static ALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);

unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            ALLOCS.fetch_add(1, Ordering::Relaxed);
            ALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
            let live = LIVE_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed)
                + layout.size() as u64;
            PEAK_LIVE_BYTES.fetch_max(live, Ordering::Relaxed);
        }
        ptr
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        DEALLOCS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(layout.size() as u64, Ordering::Relaxed);
        LIVE_BYTES.fetch_sub(layout.size() as u64, Ordering::Relaxed);
        unsafe { System.dealloc(ptr, layout) };
    }

    unsafe fn realloc(&self, ptr: *mut u8, old: Layout, size: usize) -> *mut u8 {
        let result = unsafe { System.realloc(ptr, old, size) };
        if !result.is_null() {
            ALLOCS.fetch_add(1, Ordering::Relaxed);
            ALLOC_BYTES.fetch_add(size as u64, Ordering::Relaxed);
            DEALLOCS.fetch_add(1, Ordering::Relaxed);
            DEALLOC_BYTES.fetch_add(old.size() as u64, Ordering::Relaxed);
            let live = if size >= old.size() {
                let delta = (size - old.size()) as u64;
                LIVE_BYTES.fetch_add(delta, Ordering::Relaxed) + delta
            } else {
                let delta = (old.size() - size) as u64;
                LIVE_BYTES.fetch_sub(delta, Ordering::Relaxed) - delta
            };
            PEAK_LIVE_BYTES.fetch_max(live, Ordering::Relaxed);
        }
        result
    }
}

#[global_allocator]
static GLOBAL: CountingAllocator = CountingAllocator;

#[derive(Clone, Copy, Default)]
struct Counters {
    allocs: u64,
    deallocs: u64,
    alloc_bytes: u64,
    dealloc_bytes: u64,
    live_bytes: u64,
    peak_live_bytes: u64,
}

fn counters() -> Counters {
    Counters {
        allocs: ALLOCS.load(Ordering::Relaxed),
        deallocs: DEALLOCS.load(Ordering::Relaxed),
        alloc_bytes: ALLOC_BYTES.load(Ordering::Relaxed),
        dealloc_bytes: DEALLOC_BYTES.load(Ordering::Relaxed),
        live_bytes: LIVE_BYTES.load(Ordering::Relaxed),
        peak_live_bytes: PEAK_LIVE_BYTES.load(Ordering::Relaxed),
    }
}

fn reset_counters() {
    ALLOCS.store(0, Ordering::Relaxed);
    DEALLOCS.store(0, Ordering::Relaxed);
    ALLOC_BYTES.store(0, Ordering::Relaxed);
    DEALLOC_BYTES.store(0, Ordering::Relaxed);
    let live = LIVE_BYTES.load(Ordering::Relaxed);
    PEAK_LIVE_BYTES.store(live, Ordering::Relaxed);
}

fn frame(record_type: u16, payload: &[u8], out: &mut Vec<u8>) {
    out.extend_from_slice(&record_type.to_le_bytes());
    out.extend_from_slice(&(payload.len() as u16).to_le_bytes());
    out.extend_from_slice(payload);
}

fn stream(user_count: usize, name_len: usize) -> Vec<u8> {
    let mut users = Vec::new();
    let mut sizes = [0u16; 256];
    for index in 0..user_count {
        let mut payload = Vec::with_capacity(32 + name_len);
        payload.extend_from_slice(&(index as i32).to_le_bytes());
        payload.extend_from_slice(&[index as u8; 16]);
        payload.extend_from_slice(&[0xE8, 0x07, 1, 2, 3, 4, 5, 1]);
        payload.extend_from_slice(&(name_len as u16).to_le_bytes());
        payload.push(0);
        for byte in 0..name_len {
            payload.push(b'A' + ((byte + index) % 26) as u8);
        }
        payload.push(0xA7);
        sizes[index] = payload.len() as u16;
        frame(403, &payload, &mut users);
    }
    let mut out = Vec::with_capacity(4 + 6 + 516 + users.len() + 4);
    frame(401, &(user_count as u16).to_le_bytes(), &mut out);
    frame(408, &[0x00, 0x06, 0xAA, 0xBB], &mut out);
    let mut cbusr = Vec::with_capacity(512);
    for size in sizes {
        cbusr.extend_from_slice(&size.to_le_bytes());
    }
    frame(402, &cbusr, &mut out);
    frame(407, &(user_count as u16).to_le_bytes(), &mut out);
    out.extend_from_slice(&users);
    out
}

fn millis(duration: Duration, n: usize) -> f64 {
    duration.as_secs_f64() * 1e6 / n as f64
}

fn main() {
    let cases = [(1usize, 8usize), (16, 16), (255, 54)];
    println!("build=release");
    for (user_count, name_len) in cases {
        let bytes = stream(user_count, name_len);
        let limits = UserNamesLimits::default();
        let source = Snapshot::parse_with_limits(&bytes, limits).expect("valid stream");
        let shared = Arc::<[u8]>::from(bytes.clone().into_boxed_slice());
        let shared_source = Snapshot::parse_shared_with_limits(Arc::clone(&shared), limits)
            .expect("valid shared");
        assert_eq!(source, shared_source);

        for _ in 0..10 {
            black_box(Snapshot::parse_with_limits(&bytes, limits).expect("parse"));
            black_box(source.edit());
            let mut tx = source.edit();
            tx.set_user_name(0, "Zed").expect("metadata edit");
            black_box(tx.commit().expect("commit"));
        }

        let runs = if user_count == 255 { 60 } else { 300 };
        println!("CASE users={} name_len={} bytes={} runs={}", user_count, name_len, bytes.len(), runs);
        measure("parse_slice", runs, || {
            black_box(Snapshot::parse_with_limits(&bytes, limits).expect("parse"));
        });
        measure("parse_shared", runs, || {
            black_box(Snapshot::parse_shared_with_limits(Arc::clone(&shared), limits).expect("parse"));
        });
        measure("edit_create", runs, || {
            black_box(source.edit());
        });
        measure("noop_commit", runs, || {
            let tx = source.edit();
            black_box(tx.commit().expect("noop commit"));
        });
        measure("rename_commit", runs, || {
            let mut tx = source.edit();
            tx.set_user_name(0, "Zed").expect("rename");
            black_box(tx.commit().expect("commit"));
        });
        println!();
    }
}

fn measure(label: &str, runs: usize, mut operation: impl FnMut()) {
    reset_counters();
    let baseline_live = counters().live_bytes;
    let start = Instant::now();
    for _ in 0..runs {
        operation();
    }
    let elapsed = start.elapsed();
    let c = counters();
    let peak_delta = c.peak_live_bytes.saturating_sub(baseline_live);
    println!(
        "  {label:13} us/op={:10.2} alloc/op={:8.2} alloc_bytes/op={:10.1} dealloc/op={:8.2} dealloc_bytes/op={:10.1} peak_live_delta={:10.1}",
        millis(elapsed, runs),
        c.allocs as f64 / runs as f64,
        c.alloc_bytes as f64 / runs as f64,
        c.deallocs as f64 / runs as f64,
        c.dealloc_bytes as f64 / runs as f64,
        peak_delta as f64,
    );
}
