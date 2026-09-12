#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The harness keeps allocator, source, correctness, and receipt helpers together for relocation."
)]

use litchi_core::ReadAt;
use litchi_xlsx::form_control::{ControlSelector, FormControlCollection, OwnerLimits};
use litchi_xlsx::{Error, SourceBackedWorkbook, Workbook};
use std::alloc::{GlobalAlloc, Layout, System};
use std::hint::black_box;
use std::io;
use std::path::PathBuf;
use std::sync::{
    Arc,
    atomic::{AtomicBool, AtomicU64, Ordering},
};
use std::time::Instant;

struct TrackingAllocator;

struct AllocationCounters {
    enabled: AtomicBool,
    alloc_calls: AtomicU64,
    alloc_bytes: AtomicU64,
    dealloc_calls: AtomicU64,
    dealloc_bytes: AtomicU64,
    realloc_calls: AtomicU64,
    realloc_bytes: AtomicU64,
    requested_event_bytes: AtomicU64,
    live_bytes: AtomicU64,
    phase_baseline_live_bytes: AtomicU64,
    phase_peak_live_bytes: AtomicU64,
}

impl AllocationCounters {
    const fn new() -> Self {
        Self {
            enabled: AtomicBool::new(false),
            alloc_calls: AtomicU64::new(0),
            alloc_bytes: AtomicU64::new(0),
            dealloc_calls: AtomicU64::new(0),
            dealloc_bytes: AtomicU64::new(0),
            realloc_calls: AtomicU64::new(0),
            realloc_bytes: AtomicU64::new(0),
            requested_event_bytes: AtomicU64::new(0),
            live_bytes: AtomicU64::new(0),
            phase_baseline_live_bytes: AtomicU64::new(0),
            phase_peak_live_bytes: AtomicU64::new(0),
        }
    }

    fn reset_and_enable(&self) {
        self.enabled.store(false, Ordering::SeqCst);
        self.alloc_calls.store(0, Ordering::Relaxed);
        self.alloc_bytes.store(0, Ordering::Relaxed);
        self.dealloc_calls.store(0, Ordering::Relaxed);
        self.dealloc_bytes.store(0, Ordering::Relaxed);
        self.realloc_calls.store(0, Ordering::Relaxed);
        self.realloc_bytes.store(0, Ordering::Relaxed);
        self.requested_event_bytes.store(0, Ordering::Relaxed);
        let baseline = self.live_bytes.load(Ordering::SeqCst);
        self.phase_baseline_live_bytes
            .store(baseline, Ordering::SeqCst);
        self.phase_peak_live_bytes.store(baseline, Ordering::SeqCst);
        self.enabled.store(true, Ordering::SeqCst);
    }

    fn disable(&self) -> AllocationSnapshot {
        self.enabled.store(false, Ordering::SeqCst);
        AllocationSnapshot {
            alloc_calls: self.alloc_calls.load(Ordering::Relaxed),
            alloc_bytes: self.alloc_bytes.load(Ordering::Relaxed),
            dealloc_calls: self.dealloc_calls.load(Ordering::Relaxed),
            dealloc_bytes: self.dealloc_bytes.load(Ordering::Relaxed),
            realloc_calls: self.realloc_calls.load(Ordering::Relaxed),
            realloc_bytes: self.realloc_bytes.load(Ordering::Relaxed),
            requested_event_bytes: self.requested_event_bytes.load(Ordering::Relaxed),
            live_bytes_delta: signed_delta(
                self.live_bytes.load(Ordering::SeqCst),
                self.phase_baseline_live_bytes.load(Ordering::SeqCst),
            ),
            peak_live_bytes: self
                .phase_peak_live_bytes
                .load(Ordering::SeqCst)
                .saturating_sub(self.phase_baseline_live_bytes.load(Ordering::SeqCst)),
        }
    }

    fn record_alloc_event(&self, size: usize) {
        self.alloc_calls.fetch_add(1, Ordering::Relaxed);
        self.alloc_bytes.fetch_add(size as u64, Ordering::Relaxed);
        self.record_requested(size);
    }

    fn record_dealloc_event(&self, size: usize) {
        self.dealloc_calls.fetch_add(1, Ordering::Relaxed);
        self.dealloc_bytes.fetch_add(size as u64, Ordering::Relaxed);
    }

    fn record_realloc_event(&self, size: usize) {
        self.realloc_calls.fetch_add(1, Ordering::Relaxed);
        self.realloc_bytes.fetch_add(size as u64, Ordering::Relaxed);
        self.record_requested(size);
    }

    fn record_requested(&self, size: usize) {
        self.requested_event_bytes
            .fetch_add(size as u64, Ordering::Relaxed);
    }

    fn record_live_delta(&self, delta: i128) {
        if delta >= 0 {
            self.live_bytes.fetch_add(delta as u64, Ordering::SeqCst);
        } else {
            let amount = (-delta) as u64;
            let _ = self
                .live_bytes
                .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |current| {
                    Some(current.saturating_sub(amount))
                });
        }
        if self.enabled.load(Ordering::SeqCst) {
            let live = self.live_bytes.load(Ordering::SeqCst);
            let mut peak = self.phase_peak_live_bytes.load(Ordering::SeqCst);
            while live > peak {
                match self.phase_peak_live_bytes.compare_exchange_weak(
                    peak,
                    live,
                    Ordering::SeqCst,
                    Ordering::SeqCst,
                ) {
                    Ok(_) => break,
                    Err(observed) => peak = observed,
                }
            }
        }
    }
}

unsafe impl GlobalAlloc for TrackingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc(layout) };
        if !pointer.is_null() {
            TRACKER.record_live_delta(layout.size() as i128);
            if TRACKER.enabled.load(Ordering::Relaxed) {
                TRACKER.record_alloc_event(layout.size());
            }
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if !pointer.is_null() {
            TRACKER.record_live_delta(layout.size() as i128);
            if TRACKER.enabled.load(Ordering::Relaxed) {
                TRACKER.record_alloc_event(layout.size());
            }
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) };
        TRACKER.record_live_delta(-(layout.size() as i128));
        if TRACKER.enabled.load(Ordering::Relaxed) {
            TRACKER.record_dealloc_event(layout.size());
        }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let replacement = unsafe { System.realloc(pointer, layout, new_size) };
        if !replacement.is_null() {
            TRACKER.record_live_delta(new_size as i128 - layout.size() as i128);
            if TRACKER.enabled.load(Ordering::Relaxed) {
                TRACKER.record_realloc_event(new_size);
            }
        }
        replacement
    }
}

#[global_allocator]
static GLOBAL_ALLOCATOR: TrackingAllocator = TrackingAllocator;
static TRACKER: AllocationCounters = AllocationCounters::new();

#[derive(Clone, Copy)]
struct AllocationSnapshot {
    alloc_calls: u64,
    alloc_bytes: u64,
    dealloc_calls: u64,
    dealloc_bytes: u64,
    realloc_calls: u64,
    realloc_bytes: u64,
    requested_event_bytes: u64,
    live_bytes_delta: i128,
    peak_live_bytes: u64,
}

struct SourceCounters {
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
    request_bytes: AtomicU64,
    max_request_bytes: AtomicU64,
    len_calls: AtomicU64,
    version_calls: AtomicU64,
}

impl SourceCounters {
    const fn new() -> Self {
        Self {
            read_calls: AtomicU64::new(0),
            read_bytes: AtomicU64::new(0),
            request_bytes: AtomicU64::new(0),
            max_request_bytes: AtomicU64::new(0),
            len_calls: AtomicU64::new(0),
            version_calls: AtomicU64::new(0),
        }
    }

    fn reset(&self) {
        self.read_calls.store(0, Ordering::Relaxed);
        self.read_bytes.store(0, Ordering::Relaxed);
        self.request_bytes.store(0, Ordering::Relaxed);
        self.max_request_bytes.store(0, Ordering::Relaxed);
        self.len_calls.store(0, Ordering::Relaxed);
        self.version_calls.store(0, Ordering::Relaxed);
    }

    fn record_request(&self, requested: usize, actual: usize) {
        self.read_calls.fetch_add(1, Ordering::Relaxed);
        self.request_bytes
            .fetch_add(requested as u64, Ordering::Relaxed);
        self.read_bytes.fetch_add(actual as u64, Ordering::Relaxed);
        let mut max = self.max_request_bytes.load(Ordering::Relaxed);
        while requested as u64 > max {
            match self.max_request_bytes.compare_exchange_weak(
                max,
                requested as u64,
                Ordering::Relaxed,
                Ordering::Relaxed,
            ) {
                Ok(_) => break,
                Err(observed) => max = observed,
            }
        }
    }

    fn snapshot(&self) -> SourceSnapshot {
        SourceSnapshot {
            read_calls: self.read_calls.load(Ordering::Relaxed),
            read_bytes: self.read_bytes.load(Ordering::Relaxed),
            request_bytes: self.request_bytes.load(Ordering::Relaxed),
            max_request_bytes: self.max_request_bytes.load(Ordering::Relaxed),
            len_calls: self.len_calls.load(Ordering::Relaxed),
            version_calls: self.version_calls.load(Ordering::Relaxed),
        }
    }
}

struct CountingSource {
    bytes: Arc<[u8]>,
    counters: SourceCounters,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::from(bytes),
            counters: SourceCounters::new(),
        }
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        self.counters.len_calls.fetch_add(1, Ordering::Relaxed);
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if offset >= self.bytes.len() {
            self.counters.record_request(output.len(), 0);
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        let count = end - offset;
        output[..count].copy_from_slice(&self.bytes[offset..end]);
        self.counters.record_request(output.len(), count);
        Ok(count)
    }

    fn version(&self) -> io::Result<litchi_core::SourceVersion> {
        self.counters.version_calls.fetch_add(1, Ordering::Relaxed);
        Ok(litchi_core::SourceVersion::new(0x464f524d_u64, 0))
    }
}

#[derive(Clone, Copy)]
struct SourceSnapshot {
    read_calls: u64,
    read_bytes: u64,
    request_bytes: u64,
    max_request_bytes: u64,
    len_calls: u64,
    version_calls: u64,
}

#[derive(Clone, Copy)]
struct Receipt {
    elapsed_ns: u128,
    rss_before_bytes: Option<u64>,
    rss_after_bytes: Option<u64>,
    allocation: AllocationSnapshot,
    source: SourceSnapshot,
}

impl Receipt {
    fn emit(self, fixture: &str, lane: &str, iteration: usize) {
        println!(
            "{{\"record\":\"receipt\",\"fixture\":\"{}\",\"lane\":\"{}\",\"iteration\":{},\"elapsed_ns\":{},\"rss_before_bytes\":{},\"rss_after_bytes\":{},\"alloc_calls\":{},\"alloc_bytes_requested\":{},\"dealloc_calls\":{},\"dealloc_bytes_requested\":{},\"realloc_calls\":{},\"realloc_bytes_requested\":{},\"requested_event_bytes\":{},\"live_bytes_delta\":{},\"peak_live_bytes\":{},\"source_read_calls\":{},\"source_read_bytes\":{},\"source_request_bytes\":{},\"source_max_request_bytes\":{},\"source_len_calls\":{},\"source_version_calls\":{}}}",
            fixture,
            lane,
            iteration,
            self.elapsed_ns,
            optional_u64(self.rss_before_bytes),
            optional_u64(self.rss_after_bytes),
            self.allocation.alloc_calls,
            self.allocation.alloc_bytes,
            self.allocation.dealloc_calls,
            self.allocation.dealloc_bytes,
            self.allocation.realloc_calls,
            self.allocation.realloc_bytes,
            self.allocation.requested_event_bytes,
            self.allocation.live_bytes_delta,
            self.allocation.peak_live_bytes,
            self.source.read_calls,
            self.source.read_bytes,
            self.source.request_bytes,
            self.source.max_request_bytes,
            self.source.len_calls,
            self.source.version_calls,
        );
    }
}

fn optional_u64(value: Option<u64>) -> String {
    value.map_or_else(|| "null".to_owned(), |number| number.to_string())
}

fn signed_delta(current: u64, baseline: u64) -> i128 {
    if current >= baseline {
        i128::from(current - baseline)
    } else {
        -i128::from(baseline - current)
    }
}

fn rss_bytes() -> Option<u64> {
    #[cfg(target_os = "linux")]
    {
        let status = std::fs::read_to_string("/proc/self/status").ok()?;
        let line = status.lines().find(|line| line.starts_with("VmRSS:"))?;
        let kib = line.split_whitespace().nth(1)?.parse::<u64>().ok()?;
        return Some(kib.saturating_mul(1024));
    }
    #[cfg(not(target_os = "linux"))]
    {
        None
    }
}

fn measure<T>(source: Option<&CountingSource>, operation: impl FnOnce() -> T) -> (Receipt, T) {
    if let Some(source) = source {
        source.counters.reset();
    }
    let rss_before = rss_bytes();
    TRACKER.reset_and_enable();
    let started = Instant::now();
    let value = operation();
    let elapsed_ns = started.elapsed().as_nanos();
    let allocation = TRACKER.disable();
    let rss_after = rss_bytes();
    let source_snapshot = source.map_or(
        SourceSnapshot {
            read_calls: 0,
            read_bytes: 0,
            request_bytes: 0,
            max_request_bytes: 0,
            len_calls: 0,
            version_calls: 0,
        },
        |source| source.counters.snapshot(),
    );
    (
        Receipt {
            elapsed_ns,
            rss_before_bytes: rss_before,
            rss_after_bytes: rss_after,
            allocation,
            source: source_snapshot,
        },
        value,
    )
}

fn fixture_collection_eager(bytes: &[u8]) -> FormControlCollection {
    Workbook::from_bytes(bytes.to_vec())
        .expect("eager workbook open")
        .sheet(0)
        .expect("eager sheet lookup")
        .expect("eager worksheet")
        .form_controls()
        .expect("eager form-control projection")
}

fn fixture_collection_source(bytes: &[u8]) -> (FormControlCollection, Arc<CountingSource>) {
    let source_impl = Arc::new(CountingSource::new(bytes.to_vec()));
    let source: Arc<dyn ReadAt> = source_impl.clone();
    let collection = SourceBackedWorkbook::from_read_at(source)
        .expect("source-backed workbook open")
        .sheet(0)
        .expect("source-backed sheet lookup")
        .expect("source-backed worksheet")
        .form_controls()
        .expect("source-backed form-control projection");
    (collection, source_impl)
}

fn typed_limit<T>(result: litchi_xlsx::Result<T>) -> bool {
    matches!(
        result,
        Err(Error::ResourceLimit(_))
            | Err(Error::FormControl(
                litchi_xlsx::form_control::FormControlError::Limit { .. }
            ))
    )
}

fn assert_cap_refusals(bytes: &[u8], expected_count: usize) -> (bool, bool, bool, bool) {
    let count_limits = OwnerLimits::new().with_max_controls(expected_count.saturating_sub(1));
    let eager_count = Workbook::from_bytes(bytes.to_vec())
        .expect("eager cap workbook open")
        .sheet(0)
        .expect("eager cap sheet lookup")
        .expect("eager cap worksheet")
        .form_controls_with_limits(count_limits);
    let eager_count = typed_limit(eager_count);
    let source_count =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.to_vec())))
            .expect("source cap workbook open")
            .sheet(0)
            .expect("source cap sheet lookup")
            .expect("source cap worksheet")
            .form_controls_with_limits(count_limits);
    let source_count = typed_limit(source_count);

    let projection_limits = OwnerLimits::new().with_max_projection_bytes(1);
    let eager_projection = Workbook::from_bytes(bytes.to_vec())
        .expect("eager projection cap workbook open")
        .sheet(0)
        .expect("eager projection cap sheet lookup")
        .expect("eager projection cap worksheet")
        .form_controls_with_limits(projection_limits);
    let eager_projection = typed_limit(eager_projection);
    let source_projection =
        SourceBackedWorkbook::from_read_at(Arc::new(CountingSource::new(bytes.to_vec())))
            .expect("source projection cap workbook open")
            .sheet(0)
            .expect("source projection cap sheet lookup")
            .expect("source projection cap worksheet")
            .form_controls_with_limits(projection_limits);
    let source_projection = typed_limit(source_projection);
    (
        eager_count,
        source_count,
        eager_projection,
        source_projection,
    )
}

fn correctness(bytes: &[u8], expected_count: usize) {
    let eager = fixture_collection_eager(bytes);
    let (source, source_impl) = fixture_collection_source(bytes);
    assert_eq!(eager.len(), expected_count);
    assert_eq!(source.len(), expected_count);
    assert_eq!(eager.profile(), source.profile());
    assert!(eager.read_set().is_none());
    assert!(source.read_set().is_some());
    assert!(source.source_version().is_some());
    let source_properties = source.read_set().expect("source read set").properties();
    assert_eq!(source_properties.len(), expected_count);
    for (position, (eager_view, source_view)) in eager.iter().zip(source.iter()).enumerate() {
        assert_eq!(eager_view.position(), position);
        assert_eq!(eager_view, source_view);
        let eager_bytes = eager_view
            .properties()
            .source_bytes()
            .expect("eager exact properties bytes");
        let source_bytes = source_view
            .properties()
            .source_bytes()
            .expect("source exact properties bytes");
        assert_eq!(eager_bytes, source_bytes);
        assert_eq!(eager_bytes, source_properties[position].bytes());
    }

    let eager_clone = eager.clone();
    let source_clone = source.clone();
    assert_eq!(eager_clone, eager);
    assert_eq!(source_clone, source);
    let eager_view_clone = eager
        .get(ControlSelector::position(0))
        .unwrap()
        .unwrap()
        .clone();
    let source_view_clone = source
        .get(ControlSelector::position(0))
        .unwrap()
        .unwrap()
        .clone();
    assert_eq!(
        eager_view_clone,
        eager
            .get(ControlSelector::position(0))
            .unwrap()
            .unwrap()
            .clone()
    );
    assert_eq!(
        source_view_clone,
        source
            .get(ControlSelector::position(0))
            .unwrap()
            .unwrap()
            .clone()
    );
    assert!(source_impl.counters.snapshot().read_bytes > 0);

    if expected_count > 1 {
        let first_name = eager.iter().next().and_then(|view| view.name());
        let duplicate = eager
            .iter()
            .filter(|view| view.name() == first_name)
            .count();
        if duplicate > 1 {
            let name = first_name.expect("duplicate controls have names");
            assert!(eager.get(ControlSelector::name(name)).is_err());
            assert!(source.get(ControlSelector::name(name)).is_err());
        }
    }

    let (eager_count, source_count, eager_projection, source_projection) =
        assert_cap_refusals(bytes, expected_count);
    assert!(eager_count, "eager max_controls cap admitted input");
    assert!(source_count, "source max_controls cap admitted input");
    assert!(eager_projection, "eager projection-byte cap admitted input");
    assert!(
        source_projection,
        "source projection-byte cap admitted input"
    );
}

fn parse_arg(args: &[String], name: &str) -> String {
    args.windows(2)
        .find(|pair| pair[0] == name)
        .map(|pair| pair[1].clone())
        .unwrap_or_else(|| panic!("missing required argument {name}"))
}

fn parse_optional_arg(args: &[String], name: &str, default: usize) -> usize {
    args.windows(2)
        .find(|pair| pair[0] == name)
        .map_or(default, |pair| {
            pair[1]
                .parse::<usize>()
                .unwrap_or_else(|_| panic!("invalid value for {name}"))
        })
}

fn main() {
    let args: Vec<String> = std::env::args().collect();
    let fixture_path = PathBuf::from(parse_arg(&args, "--fixture"));
    let fixture_label = parse_arg(&args, "--label");
    let expected_count = parse_optional_arg(&args, "--expected-count", 1);
    let iterations = parse_optional_arg(&args, "--iterations", 7);
    let bytes = std::fs::read(&fixture_path).expect("read fixture");

    correctness(&bytes, expected_count);
    println!(
        "{{\"record\":\"correctness\",\"fixture\":\"{}\",\"fixture_bytes\":{},\"expected_controls\":{},\"status\":\"pass\",\"caps\":{{\"max_controls\":\"typed_refusal\",\"max_projection_bytes\":\"typed_refusal\"}}}}",
        fixture_label,
        bytes.len(),
        expected_count
    );

    for iteration in 0..iterations {
        let (receipt, workbook) = measure(None, || {
            Workbook::from_bytes(bytes.clone()).expect("eager package open")
        });
        receipt.emit(&fixture_label, "eager_package_open", iteration);
        let (receipt, collection) = measure(None, || {
            workbook
                .sheet(0)
                .expect("eager projection sheet lookup")
                .expect("eager projection worksheet")
                .form_controls()
                .expect("eager projection")
        });
        receipt.emit(&fixture_label, "eager_owner_projection", iteration);
        assert_eq!(collection.len(), expected_count);

        let source_impl = Arc::new(CountingSource::new(bytes.clone()));
        let source: Arc<dyn ReadAt> = source_impl.clone();
        let (receipt, workbook) = measure(Some(&source_impl), || {
            SourceBackedWorkbook::from_read_at(source.clone()).expect("source package open")
        });
        receipt.emit(&fixture_label, "source_package_open", iteration);
        let (receipt, collection) = measure(Some(&source_impl), || {
            workbook
                .sheet(0)
                .expect("source projection sheet lookup")
                .expect("source projection worksheet")
                .form_controls()
                .expect("source projection")
        });
        receipt.emit(&fixture_label, "source_owner_projection", iteration);
        assert_eq!(collection.len(), expected_count);

        let eager_collection = fixture_collection_eager(&bytes);
        let source_collection = fixture_collection_source(&bytes).0;
        let (receipt, checksum) = measure(None, || {
            let view = eager_collection
                .get(ControlSelector::position(0))
                .expect("eager position query")
                .expect("eager selected control");
            black_box(view.position())
        });
        receipt.emit(&fixture_label, "eager_single_query", iteration);
        black_box(checksum);
        let (receipt, checksum) = measure(None, || {
            let view = source_collection
                .get(ControlSelector::position(0))
                .expect("source position query")
                .expect("source selected control");
            black_box(view.position())
        });
        receipt.emit(&fixture_label, "source_single_query", iteration);
        black_box(checksum);

        let (receipt, checksum) = measure(None, || {
            let mut checksum = 0usize;
            for view in eager_collection.iter() {
                checksum ^= view.position();
            }
            for position in 0..eager_collection.len() {
                checksum ^= eager_collection
                    .get(ControlSelector::position(position))
                    .expect("eager many query")
                    .expect("eager many selected control")
                    .position();
            }
            checksum
        });
        receipt.emit(&fixture_label, "eager_many_query", iteration);
        black_box(checksum);
        let (receipt, checksum) = measure(None, || {
            let mut checksum = 0usize;
            for view in source_collection.iter() {
                checksum ^= view.position();
            }
            for position in 0..source_collection.len() {
                checksum ^= source_collection
                    .get(ControlSelector::position(position))
                    .expect("source many query")
                    .expect("source many selected control")
                    .position();
            }
            checksum
        });
        receipt.emit(&fixture_label, "source_many_query", iteration);
        black_box(checksum);

        let (receipt, clones) = measure(None, || {
            let collection = eager_collection.clone();
            let view = collection
                .get(ControlSelector::position(0))
                .expect("eager clone query")
                .expect("eager clone selected control")
                .clone();
            (collection, view)
        });
        receipt.emit(&fixture_label, "eager_cheap_clone", iteration);
        black_box(clones);
        let (receipt, clones) = measure(None, || {
            let collection = source_collection.clone();
            let view = collection
                .get(ControlSelector::position(0))
                .expect("source clone query")
                .expect("source clone selected control")
                .clone();
            (collection, view)
        });
        receipt.emit(&fixture_label, "source_cheap_clone", iteration);
        black_box(clones);
    }
}
