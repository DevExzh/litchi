//! Matched per-lifecycle ODP source-owner probe for record 0362ye.
use std::alloc::{GlobalAlloc, Layout, System};
use std::hash::Hasher;
use std::io::Write;
use std::sync::{
    Arc,
    atomic::{AtomicUsize, Ordering},
};

struct Counting;

static CURRENT: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);
static ALLOCS: AtomicUsize = AtomicUsize::new(0);
static REALLOCS: AtomicUsize = AtomicUsize::new(0);
static ALLOCATED: AtomicUsize = AtomicUsize::new(0);
static ALLOC_BUCKETS: [AtomicUsize; 24] = [const { AtomicUsize::new(0) }; 24];
static ALLOC_LOG_LEN: AtomicUsize = AtomicUsize::new(0);
static ALLOC_LOG: [AtomicUsize; 256] = [const { AtomicUsize::new(0) }; 256];

fn record_allocation(size: usize) {
    let index = ALLOC_LOG_LEN.fetch_add(1, Ordering::Relaxed);
    if index < ALLOC_LOG.len() {
        ALLOC_LOG[index].store(size, Ordering::Relaxed);
    }
}

fn alloc_bucket(size: usize) -> usize {
    if size == 0 {
        return 0;
    }
    let bucket = (usize::BITS - size.leading_zeros()) as usize;
    bucket.min(23)
}

fn phase_reset() -> usize {
    let baseline = CURRENT.load(Ordering::Relaxed);
    PEAK.store(baseline, Ordering::Relaxed);
    ALLOCS.store(0, Ordering::Relaxed);
    REALLOCS.store(0, Ordering::Relaxed);
    ALLOCATED.store(0, Ordering::Relaxed);
    for bucket in &ALLOC_BUCKETS {
        bucket.store(0, Ordering::Relaxed);
    }
    ALLOC_LOG_LEN.store(0, Ordering::Relaxed);
    baseline
}

fn phase_report(name: &str, baseline: usize) {
    let mut sizes = String::new();
    for (index, bucket) in ALLOC_BUCKETS.iter().enumerate() {
        let count = bucket.load(Ordering::Relaxed);
        if count != 0 {
            use std::fmt::Write as _;
            let _ = write!(sizes, "{index}:{count},");
        }
    }
    let mut log = String::new();
    let count = ALLOC_LOG_LEN.load(Ordering::Relaxed).min(ALLOC_LOG.len());
    for allocation in ALLOC_LOG.iter().take(count) {
        use std::fmt::Write as _;
        let _ = write!(log, "{}:", allocation.load(Ordering::Relaxed));
    }
    println!(
        "{{\"phase\":\"{name}\",\"peak\":{},\"allocs\":{},\"reallocs\":{},\"allocated\":{},\"live_delta\":{},\"alloc_sizes\":\"{sizes}\",\"alloc_log\":\"{log}\"}}",
        PEAK.load(Ordering::Relaxed).saturating_sub(baseline),
        ALLOCS.load(Ordering::Relaxed),
        REALLOCS.load(Ordering::Relaxed),
        ALLOCATED.load(Ordering::Relaxed),
        CURRENT.load(Ordering::Relaxed).saturating_sub(baseline),
    );
}

// SAFETY: every operation forwards to System and only observes sizes.
unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: forwarded with the caller's layout.
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            record_allocation(layout.size());
            ALLOC_BUCKETS[alloc_bucket(layout.size())].fetch_add(1, Ordering::Relaxed);
            ALLOCS.fetch_add(1, Ordering::Relaxed);
            ALLOCATED.fetch_add(layout.size(), Ordering::Relaxed);
            let now = CURRENT.fetch_add(layout.size(), Ordering::Relaxed) + layout.size();
            PEAK.fetch_max(now, Ordering::Relaxed);
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: forwarded with the caller's pointer and layout.
        unsafe { System.dealloc(pointer, layout) };
        CURRENT.fetch_sub(layout.size(), Ordering::Relaxed);
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, size: usize) -> *mut u8 {
        // SAFETY: forwarded with the caller's arguments.
        let moved = unsafe { System.realloc(pointer, layout, size) };
        if !moved.is_null() {
            REALLOCS.fetch_add(1, Ordering::Relaxed);
            if size >= layout.size() {
                let delta = size - layout.size();
                record_allocation(delta);
                ALLOC_BUCKETS[alloc_bucket(delta)].fetch_add(1, Ordering::Relaxed);
                ALLOCATED.fetch_add(delta, Ordering::Relaxed);
                let now = CURRENT.fetch_add(delta, Ordering::Relaxed) + delta;
                PEAK.fetch_max(now, Ordering::Relaxed);
            } else {
                CURRENT.fetch_sub(layout.size() - size, Ordering::Relaxed);
            }
        }
        moved
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

const ODP_SLIDES: usize = 2_000;
const MEDIA_BYTES: usize = 4 * 1024 * 1024;

/// A sink retaining only output length and a deterministic SipHash.
struct Digest {
    hasher: std::collections::hash_map::DefaultHasher,
    bytes: u64,
}

impl Digest {
    fn new() -> Self {
        Self {
            hasher: std::collections::hash_map::DefaultHasher::new(),
            bytes: 0,
        }
    }

    fn finish(self) -> (u64, u64) {
        (self.bytes, self.hasher.finish())
    }

    fn update_bytes(&mut self, bytes: &[u8]) {
        self.hasher.write(bytes);
        self.bytes += bytes.len() as u64;
    }

    fn update_text(&mut self, text: &str) {
        self.update_bytes(text.as_bytes());
        // Keep adjacent semantic values distinguishable in the receipt.
        self.update_bytes(&[0]);
    }
}

impl Write for Digest {
    fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
        self.hasher.write(buffer);
        self.bytes += buffer.len() as u64;
        Ok(buffer.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

fn media() -> Vec<u8> {
    let mut state = 0x9e37_79b9_7f4a_7c15_u64;
    (0..MEDIA_BYTES)
        .map(|_| {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            state as u8
        })
        .collect()
}

fn odp() -> Vec<u8> {
    let mut content = String::from(
        r#"<?xml version="1.0" encoding="UTF-8"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" office:version="1.4"><office:body><office:presentation>"#,
    );
    for slide in 0..ODP_SLIDES {
        content.push_str(&format!(
            r#"<draw:page draw:name="page{}"><draw:frame presentation:class="title"><draw:text-box><text:p>Slide {slide}</text:p></draw:text-box></draw:frame><draw:frame presentation:class="outline"><draw:text-box><text:p>Body text of slide {slide} with a few words.</text:p></draw:text-box></draw:frame></draw:page>"#,
            slide + 1
        ));
    }
    content.push_str("</office:presentation></office:body></office:document-content>");
    let mut writer = litchi_odf_common::core::PackageWriter::new(
        litchi_odf_common::core::StreamingArchiveLimits::default(),
    );
    writer
        .set_mimetype("application/vnd.oasis.opendocument.presentation")
        .expect("mimetype");
    writer
        .add_file("content.xml", content.as_bytes())
        .expect("content");
    writer
        .add_file("Pictures/blob.bin", &media())
        .expect("media");
    writer.finish_to_bytes().expect("odp")
}

fn measure(
    source: &[u8],
    operation: fn(Arc<dyn litchi_core::ReadAt>, &mut Digest),
) -> (usize, usize, usize, usize, u64, u64) {
    // Keep the source clone outside the measured lifecycle, as in 0358a.
    let reader: Arc<dyn litchi_core::ReadAt> =
        Arc::new(litchi_core::OwnedSource::new(source.to_vec()));
    let start = CURRENT.load(Ordering::Relaxed);
    PEAK.store(start, Ordering::Relaxed);
    ALLOCS.store(0, Ordering::Relaxed);
    REALLOCS.store(0, Ordering::Relaxed);
    ALLOCATED.store(0, Ordering::Relaxed);
    for bucket in &ALLOC_BUCKETS {
        bucket.store(0, Ordering::Relaxed);
    }
    ALLOC_LOG_LEN.store(0, Ordering::Relaxed);
    let mut digest = Digest::new();
    operation(reader, &mut digest);
    let (bytes, hash) = digest.finish();
    (
        PEAK.load(Ordering::Relaxed).saturating_sub(start),
        ALLOCS.load(Ordering::Relaxed),
        REALLOCS.load(Ordering::Relaxed),
        ALLOCATED.load(Ordering::Relaxed),
        bytes,
        hash,
    )
}

fn run(name: &str, source: &[u8]) {
    let operation = match name {
        "open" => route::open,
        "first_slide" => route::first_slide,
        "repeat_slide" => route::repeat_slide,
        "first_text" => route::first_text,
        "repeat_text" => route::repeat_text,
        "write_text" => route::write_text,
        "repeat_write_text" => route::repeat_write_text,
        "unchanged_write" => route::unchanged_write,
        "append_commit_write" => route::append_commit_write,
        "replay" => route::replay,
        "append_phases" => route::append_phases,
        "replay_phases" => route::replay_phases,
        _ => panic!("unknown route: {name}"),
    };
    let (peak, allocs, reallocs, allocated, output_bytes, output_siphash) =
        measure(source, operation);
    println!(
        "{{\"route\":\"{name}\",\"mode\":\"{}\",\"peak\":{peak},\"allocs\":{allocs},\"reallocs\":{reallocs},\"allocated\":{allocated},\"output_bytes\":{output_bytes},\"output_siphash\":{output_siphash}}}",
        route::MODE,
    );
}

fn main() {
    let source = odp();
    let routes = [
        "open",
        "first_slide",
        "repeat_slide",
        "first_text",
        "repeat_text",
        "write_text",
        "repeat_write_text",
        "unchanged_write",
        "append_commit_write",
        "replay",
    ];
    match std::env::args().nth(1) {
        Some(name) => run(&name, &source),
        None => {
            for name in routes {
                run(name, &source);
            }
        },
    }
}

#[path = "../../../docs/unified-ops/results/0362ye/route_after.rs"]
mod route;
