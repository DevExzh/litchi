//! Process-isolated allocator and runtime evidence for the shared
//! `drawing/2015/06/chart:formatcode2` element owner.
//!
//! The harness deliberately keeps fixture construction and semantic checks
//! outside the timed interval. It is an opt-in report binary under this
//! evidence directory; it is not a production benchmark dependency.

#![allow(
    unsafe_code,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::print_stdout,
    clippy::shadow_reuse,
    reason = "the opt-in profile owns a process-local allocator observer and emits JSON"
)]

use litchi_drawingml::chart::extension::formatcode2::{self, Attribute, Element, Value};
use std::alloc::{GlobalAlloc, Layout, System};
use std::env;
use std::error::Error;
use std::fmt::Write as FmtWrite;
use std::hint::black_box;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::Instant;
use std::{io, sync::Arc};

type BoxError = Box<dyn Error + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

const DEFAULT_WARMUP: usize = 2;
const DEFAULT_SAMPLES: usize = 20;
const SMALL_VALUE: &str = "[$-en-US]#,##0.00";
const SMALL_UPDATED: &str = "[$-fr-FR]#,##0.00";
const SMALL_ATTRIBUTE_VALUE: &str = "A";
const SMALL_ATTRIBUTE_UPDATED: &str = "B";
const MALFORMED_SMALL_VALUE: &str = "_xD800_";

const LANES: &[&str] = &[
    "element_small_read",
    "element_small_noop",
    "element_small_scalar_edit",
    "element_small_clone",
    "element_small_read_shared",
    "element_small_write_to",
    "element_small_malformed",
    "element_near_limit_read",
    "element_near_limit_noop",
    "element_near_limit_scalar_edit",
    "element_near_limit_clone",
    "element_near_limit_read_shared",
    "element_near_limit_write_to",
    "element_near_limit_malformed",
    "attribute_small_read",
    "attribute_small_noop",
    "attribute_small_scalar_edit",
    "attribute_small_clone",
    "attribute_small_read_shared",
    "attribute_small_write_to",
    "attribute_small_malformed",
    "attribute_near_limit_read",
    "attribute_near_limit_noop",
    "attribute_near_limit_scalar_edit",
    "attribute_near_limit_clone",
    "attribute_near_limit_read_shared",
    "attribute_near_limit_write_to",
    "attribute_near_limit_malformed",
];

#[global_allocator]
static GLOBAL_ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static REALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DIRECT_BYTES: AtomicU64 = AtomicU64::new(0);
static REALLOC_OLD: AtomicU64 = AtomicU64::new(0);
static REALLOC_NEW: AtomicU64 = AtomicU64::new(0);
static DEALLOC_BYTES: AtomicU64 = AtomicU64::new(0);
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_FAILED: AtomicU64 = AtomicU64::new(0);
static INVALID: AtomicBool = AtomicBool::new(false);

// SAFETY: every method forwards the allocation contract to `System`; the
// atomics observe successful operations without changing pointer semantics.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid layout.
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid layout.
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        // SAFETY: the pointer/layout pair belongs to the caller.
        unsafe { System.dealloc(pointer, layout) };
        DEALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
        DEALLOC_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
        subtract_live(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the caller supplies the valid pointer/layout contract.
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if result.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            REALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
            REALLOC_OLD.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
            REALLOC_NEW.fetch_add(as_u64(new_size), Ordering::Relaxed);
            if new_size >= layout.size() {
                observe_growth(new_size - layout.size());
            } else {
                subtract_live(layout.size() - new_size);
            }
        }
        result
    }
}

fn as_u64(value: usize) -> u64 {
    u64::try_from(value).unwrap_or(u64::MAX)
}

fn observe_alloc(size: usize) {
    ALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
    DIRECT_BYTES.fetch_add(as_u64(size), Ordering::Relaxed);
    observe_growth(size);
}

fn observe_growth(size: usize) {
    let size = as_u64(size);
    let live = LIVE_BYTES
        .fetch_add(size, Ordering::Relaxed)
        .saturating_add(size);
    let mut old = PEAK_BYTES.load(Ordering::Relaxed);
    while live > old {
        match PEAK_BYTES.compare_exchange_weak(old, live, Ordering::Relaxed, Ordering::Relaxed) {
            Ok(_) => break,
            Err(observed) => old = observed,
        }
    }
}

fn subtract_live(size: usize) {
    let size = as_u64(size);
    let before = LIVE_BYTES.fetch_sub(size, Ordering::Relaxed);
    if before < size {
        INVALID.store(true, Ordering::Release);
    }
}

#[derive(Clone, Copy)]
struct AllocSnapshot {
    calls: u64,
    realloc_calls: u64,
    dealloc_calls: u64,
    direct: u64,
    realloc_old: u64,
    realloc_new: u64,
    deallocated: u64,
    live: u64,
    peak: u64,
    failed: u64,
    invalid: bool,
}

impl AllocSnapshot {
    fn now() -> Self {
        Self {
            calls: ALLOC_CALLS.load(Ordering::Acquire),
            realloc_calls: REALLOC_CALLS.load(Ordering::Acquire),
            dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
            direct: DIRECT_BYTES.load(Ordering::Acquire),
            realloc_old: REALLOC_OLD.load(Ordering::Acquire),
            realloc_new: REALLOC_NEW.load(Ordering::Acquire),
            deallocated: DEALLOC_BYTES.load(Ordering::Acquire),
            live: LIVE_BYTES.load(Ordering::Acquire),
            peak: PEAK_BYTES.load(Ordering::Acquire),
            failed: ALLOC_FAILED.load(Ordering::Acquire),
            invalid: INVALID.load(Ordering::Acquire),
        }
    }

    fn delta(self, after: Self) -> AllocDelta {
        AllocDelta {
            calls: after.calls.saturating_sub(self.calls),
            realloc_calls: after.realloc_calls.saturating_sub(self.realloc_calls),
            dealloc_calls: after.dealloc_calls.saturating_sub(self.dealloc_calls),
            direct: after.direct.saturating_sub(self.direct),
            realloc_old: after.realloc_old.saturating_sub(self.realloc_old),
            realloc_new: after.realloc_new.saturating_sub(self.realloc_new),
            deallocated: after.deallocated.saturating_sub(self.deallocated),
            live_before: self.live,
            live_after: after.live,
            peak_delta: after.peak.saturating_sub(self.peak),
            failed: after.failed.saturating_sub(self.failed),
            invalid: self.invalid || after.invalid,
        }
    }
}

#[derive(Clone, Copy)]
struct AllocDelta {
    calls: u64,
    realloc_calls: u64,
    dealloc_calls: u64,
    direct: u64,
    realloc_old: u64,
    realloc_new: u64,
    deallocated: u64,
    live_before: u64,
    live_after: u64,
    peak_delta: u64,
    failed: u64,
    invalid: bool,
}

impl AllocDelta {
    fn requested(self) -> u64 {
        self.direct.saturating_add(self.realloc_new)
    }

    fn balanced(self) -> bool {
        self.live_before
            .checked_add(self.direct)
            .and_then(|value| value.checked_add(self.realloc_new))
            .and_then(|value| value.checked_sub(self.realloc_old))
            .and_then(|value| value.checked_sub(self.deallocated))
            == Some(self.live_after)
    }
}

fn reset_counters() {
    PEAK_BYTES.store(LIVE_BYTES.load(Ordering::Acquire), Ordering::Release);
    ALLOC_CALLS.store(0, Ordering::Release);
    REALLOC_CALLS.store(0, Ordering::Release);
    DEALLOC_CALLS.store(0, Ordering::Release);
    DIRECT_BYTES.store(0, Ordering::Release);
    REALLOC_OLD.store(0, Ordering::Release);
    REALLOC_NEW.store(0, Ordering::Release);
    DEALLOC_BYTES.store(0, Ordering::Release);
    ALLOC_FAILED.store(0, Ordering::Release);
    INVALID.store(false, Ordering::Release);
}

fn allocator_counter_self_test() -> Result<()> {
    reset_counters();
    let before = AllocSnapshot::now();
    let layout = Layout::from_size_align(8, std::mem::align_of::<usize>())?;
    // SAFETY: this deliberately exercises the observer with a valid layout.
    let pointer = unsafe { std::alloc::alloc(layout) };
    if pointer.is_null() {
        reset_counters();
        return Err("allocator self-test allocation failed".into());
    }
    // SAFETY: `pointer` was returned with `layout`, and the new size is valid.
    let resized = unsafe { std::alloc::realloc(pointer, layout, 32) };
    if resized.is_null() {
        // SAFETY: a failed reallocation leaves the original allocation valid.
        unsafe { std::alloc::dealloc(pointer, layout) };
        reset_counters();
        return Err("allocator self-test reallocation failed".into());
    }
    let resized_layout = Layout::from_size_align(32, layout.align())?;
    // SAFETY: `resized` uses the same alignment and requested new size.
    unsafe { std::alloc::dealloc(resized, resized_layout) };
    let delta = before.delta(AllocSnapshot::now());
    let valid = delta.calls == 1
        && delta.realloc_calls == 1
        && delta.dealloc_calls == 1
        && delta.direct == 8
        && delta.realloc_old == 8
        && delta.realloc_new == 32
        && delta.deallocated == 32
        && delta.requested() == 40
        && delta.balanced()
        && !delta.invalid
        && delta.failed == 0;
    reset_counters();
    if valid {
        Ok(())
    } else {
        Err("allocator self-test counters did not balance".into())
    }
}

const FNV_OFFSET: u64 = 0xcbf29ce484222325;
const FNV_PRIME: u64 = 0x100000001b3;

/// A non-allocating sink used to distinguish `write_to` from `write`.
///
/// The sink retains only a byte count, a deterministic FNV-1a digest, and the
/// number of calls. The expected count and digest are computed after the
/// timed closure from the immutable source bytes.
#[derive(Clone, Copy)]
struct HashSink {
    bytes: u64,
    hash: u64,
    writes: u64,
}

impl HashSink {
    fn new() -> Self {
        Self {
            bytes: 0,
            hash: FNV_OFFSET,
            writes: 0,
        }
    }
}

impl io::Write for HashSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes = self
            .bytes
            .checked_add(as_u64(bytes.len()))
            .ok_or_else(|| io::Error::other("formatcode2 sink byte count overflow"))?;
        self.hash = hash_bytes(self.hash, bytes);
        self.writes = self
            .writes
            .checked_add(1)
            .ok_or_else(|| io::Error::other("formatcode2 sink write count overflow"))?;
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn hash_bytes(mut hash: u64, bytes: &[u8]) -> u64 {
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(FNV_PRIME);
    }
    hash
}

struct Inputs {
    element_small: Vec<u8>,
    element_near_limit: Vec<u8>,
    element_malformed_small: Vec<u8>,
    element_malformed_near_limit: Vec<u8>,
    element_small_shared: Arc<[u8]>,
    element_near_limit_shared: Arc<[u8]>,
    attribute_small: Vec<u8>,
    attribute_near_limit: Vec<u8>,
    attribute_malformed_small: Vec<u8>,
    attribute_malformed_near_limit: Vec<u8>,
    attribute_small_shared: Arc<[u8]>,
    attribute_near_limit_shared: Arc<[u8]>,
    element_small_prepared: Element,
    element_near_prepared: Element,
    attribute_small_prepared: Attribute,
    attribute_near_prepared: Attribute,
    near_value: String,
    near_updated: String,
}

impl Inputs {
    fn load() -> Result<Self> {
        let element_small = small_element_source();
        let attribute_small = small_attribute_source();
        let near_value = "0".repeat(formatcode2::MAX_VALUE_BYTES.saturating_sub(32));
        let mut near_updated = near_value.clone();
        near_updated.replace_range(0..1, "1");
        let element_near_limit = near_element_source(&near_value)?;
        let attribute_near_limit = near_attribute_source(&near_value)?;
        let element_malformed_small = malformed_element_source();
        let attribute_malformed_small = malformed_attribute_source();
        let mut element_malformed_near_limit = element_near_limit.clone();
        element_malformed_near_limit.push(b'X');
        let mut attribute_malformed_near_limit = attribute_near_limit.clone();
        attribute_malformed_near_limit.push(b'X');
        let element_small_shared = Arc::from(element_small.clone());
        let element_near_limit_shared = Arc::from(element_near_limit.clone());
        let attribute_small_shared = Arc::from(attribute_small.clone());
        let attribute_near_limit_shared = Arc::from(attribute_near_limit.clone());

        for (name, length) in [
            ("element", element_near_limit.len()),
            ("attribute", attribute_near_limit.len()),
        ] {
            if length != formatcode2::MAX_XML_BYTES - 1 {
                return Err(
                    format!("{name} near-limit fixture did not reach the source bound").into(),
                );
            }
        }
        for (name, length) in [
            ("element", element_malformed_near_limit.len()),
            ("attribute", attribute_malformed_near_limit.len()),
        ] {
            if length != formatcode2::MAX_XML_BYTES {
                return Err(
                    format!("{name} near-limit malformed fixture has the wrong size").into(),
                );
            }
        }
        let element_small_prepared = formatcode2::read(&element_small)?;
        let element_near_prepared = formatcode2::read(&element_near_limit)?;
        let attribute_small_prepared = formatcode2::read_attribute(&attribute_small)?;
        let attribute_near_prepared = formatcode2::read_attribute(&attribute_near_limit)?;
        if element_small_prepared.value() != SMALL_VALUE
            || element_near_prepared.value() != near_value
            || attribute_small_prepared.value() != SMALL_ATTRIBUTE_VALUE
            || attribute_near_prepared.value() != near_value
        {
            return Err("prepared formatcode2 fixtures have unexpected values".into());
        }
        let detached = Value::new(SMALL_VALUE)?;
        if detached.as_str() != SMALL_VALUE {
            return Err("detached formatcode2 Value has an unexpected value".into());
        }
        if formatcode2::read(&element_malformed_small).is_ok()
            || formatcode2::read(&element_malformed_near_limit).is_ok()
            || formatcode2::read_attribute(&attribute_malformed_small).is_ok()
            || formatcode2::read_attribute(&attribute_malformed_near_limit).is_ok()
        {
            return Err("malformed formatcode2 fixture was accepted".into());
        }
        Ok(Self {
            element_small,
            element_near_limit,
            element_malformed_small,
            element_malformed_near_limit,
            element_small_shared,
            element_near_limit_shared,
            attribute_small,
            attribute_near_limit,
            attribute_malformed_small,
            attribute_malformed_near_limit,
            attribute_small_shared,
            attribute_near_limit_shared,
            element_small_prepared,
            element_near_prepared,
            attribute_small_prepared,
            attribute_near_prepared,
            near_value,
            near_updated,
        })
    }
}

fn small_element_source() -> Vec<u8> {
    format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!-- retained-small -->\n<f:formatcode2 xmlns:f=\"{}\">{SMALL_VALUE}</f:formatcode2>\n",
        formatcode2::NAMESPACE
    )
    .into_bytes()
}

fn small_attribute_source() -> Vec<u8> {
    format!(
        r#"<c:numFmt xmlns:c="{}" formatCode="0" c:formatcode2="_x0041_" data="keep">"#,
        formatcode2::NAMESPACE
    )
    .into_bytes()
}

fn malformed_element_source() -> Vec<u8> {
    format!(
        "<f:formatcode2 xmlns:f=\"{}\">{MALFORMED_SMALL_VALUE}</f:formatcode2>",
        formatcode2::NAMESPACE
    )
    .into_bytes()
}

fn malformed_attribute_source() -> Vec<u8> {
    format!(
        r#"<c:numFmt xmlns:c="{}" c:formatcode2="{}">"#,
        formatcode2::NAMESPACE,
        MALFORMED_SMALL_VALUE
    )
    .into_bytes()
}

fn near_element_source(value: &str) -> Result<Vec<u8>> {
    let prefix = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<f:formatcode2 xmlns:f=\"{}\">",
        formatcode2::NAMESPACE
    );
    let suffix = b"</f:formatcode2>";
    let comment_open = b"<!--";
    let comment_close = b"-->";
    let target = formatcode2::MAX_XML_BYTES - 1;
    let root_len = prefix
        .len()
        .checked_add(value.len())
        .and_then(|length| length.checked_add(suffix.len()))
        .ok_or("near-limit fixture size overflow")?;
    let comment_overhead = comment_open
        .len()
        .checked_add(comment_close.len())
        .ok_or("near-limit comment size overflow")?;
    let body_len = target
        .checked_sub(root_len)
        .and_then(|length| length.checked_sub(comment_overhead))
        .ok_or("near-limit fixture has no room for comment")?;
    let mut output = Vec::with_capacity(target);
    output.extend_from_slice(prefix.as_bytes());
    output.extend_from_slice(value.as_bytes());
    output.extend_from_slice(suffix);
    output.extend_from_slice(comment_open);
    output.resize(output.len() + body_len, b'x');
    output.extend_from_slice(comment_close);
    Ok(output)
}

fn near_attribute_source(value: &str) -> Result<Vec<u8>> {
    let prefix = format!(
        r#"<c:numFmt xmlns:c="{}" formatCode="0" c:formatcode2="{}""#,
        formatcode2::NAMESPACE,
        value
    );
    let filler_count = formatcode2::MAX_ATTRIBUTES
        .checked_sub(3)
        .ok_or("attribute count profile constants underflow")?;
    let mut names = Vec::with_capacity(filler_count);
    for index in 0..filler_count {
        names.push(format!(" data{index}=\""));
    }
    let fixed = names
        .iter()
        .map(|name| name.len() + 1)
        .try_fold(1usize, usize::checked_add)
        .ok_or("attribute filler size overflow")?;
    let target = formatcode2::MAX_XML_BYTES - 1;
    let payload = target
        .checked_sub(prefix.len())
        .and_then(|length| length.checked_sub(fixed))
        .ok_or("attribute near-limit fixture has no room for filler")?;
    let max_each = formatcode2::MAX_VALUE_BYTES - 1;
    if payload > max_each * filler_count {
        return Err("attribute near-limit fixture exceeds filler capacity".into());
    }
    let mut output = prefix.into_bytes();
    let mut remaining = payload;
    for (index, name) in names.iter().enumerate() {
        output.extend_from_slice(name.as_bytes());
        let value_len = remaining.min(max_each);
        output.resize(output.len() + value_len, b'x');
        output.push(b'\"');
        remaining -= value_len;
        if index + 1 == filler_count && remaining != 0 {
            return Err("attribute near-limit fixture filler did not fit".into());
        }
    }
    output.push(b'>');
    if output.len() != target {
        return Err("attribute near-limit fixture has the wrong size".into());
    }
    Ok(output)
}

enum Work {
    Element(Element),
    Attribute(Attribute),
    Bytes(Vec<u8>),
    Sink(HashSink),
    ElementEdit { value: Element, bytes: Vec<u8> },
    AttributeEdit { value: Attribute, bytes: Vec<u8> },
    Rejected,
}

fn is_malformed(lane: &str) -> bool {
    lane.ends_with("_malformed")
}

fn is_near(lane: &str) -> bool {
    lane.contains("_near_limit_")
}

fn is_attribute(lane: &str) -> bool {
    lane.starts_with("attribute_")
}

fn expected_success(lane: &str) -> bool {
    !is_malformed(lane)
}

fn operation(inputs: &Inputs, lane: &str) -> Result<Work> {
    let source = if is_attribute(lane) {
        if is_near(lane) {
            &inputs.attribute_near_limit
        } else {
            &inputs.attribute_small
        }
    } else if is_near(lane) {
        &inputs.element_near_limit
    } else {
        &inputs.element_small
    };
    let element_prepared = if is_near(lane) {
        &inputs.element_near_prepared
    } else {
        &inputs.element_small_prepared
    };
    let attribute_prepared = if is_near(lane) {
        &inputs.attribute_near_prepared
    } else {
        &inputs.attribute_small_prepared
    };
    let element_shared = if is_near(lane) {
        &inputs.element_near_limit_shared
    } else {
        &inputs.element_small_shared
    };
    let attribute_shared = if is_near(lane) {
        &inputs.attribute_near_limit_shared
    } else {
        &inputs.attribute_small_shared
    };
    match lane {
        "element_small_read" | "element_near_limit_read" => {
            Ok(Work::Element(formatcode2::read(source)?))
        },
        "element_small_read_shared" | "element_near_limit_read_shared" => Ok(Work::Element(
            formatcode2::read_shared(Arc::clone(element_shared))?,
        )),
        "element_small_noop" | "element_near_limit_noop" => {
            Ok(Work::Bytes(formatcode2::write(element_prepared)?))
        },
        "element_small_write_to" | "element_near_limit_write_to" => {
            let mut sink = HashSink::new();
            formatcode2::write_to(&mut sink, element_prepared)?;
            Ok(Work::Sink(sink))
        },
        "element_small_scalar_edit" | "element_near_limit_scalar_edit" => {
            let mut value = element_prepared.clone();
            let updated = if is_near(lane) {
                &inputs.near_updated
            } else {
                SMALL_UPDATED
            };
            value.set_value(updated)?;
            let bytes = formatcode2::write(&value)?;
            Ok(Work::ElementEdit { value, bytes })
        },
        "element_small_clone" | "element_near_limit_clone" => {
            Ok(Work::Element(element_prepared.clone()))
        },
        "attribute_small_read" | "attribute_near_limit_read" => {
            Ok(Work::Attribute(formatcode2::read_attribute(source)?))
        },
        "attribute_small_read_shared" | "attribute_near_limit_read_shared" => Ok(Work::Attribute(
            formatcode2::read_attribute_shared(Arc::clone(attribute_shared))?,
        )),
        "attribute_small_noop" | "attribute_near_limit_noop" => Ok(Work::Bytes(
            formatcode2::write_attribute(attribute_prepared)?,
        )),
        "attribute_small_write_to" | "attribute_near_limit_write_to" => {
            let mut sink = HashSink::new();
            formatcode2::write_attribute_to(&mut sink, attribute_prepared)?;
            Ok(Work::Sink(sink))
        },
        "attribute_small_scalar_edit" | "attribute_near_limit_scalar_edit" => {
            let mut value = attribute_prepared.clone();
            let updated = if is_near(lane) {
                &inputs.near_updated
            } else {
                SMALL_ATTRIBUTE_UPDATED
            };
            value.set_value(updated)?;
            let bytes = formatcode2::write_attribute(&value)?;
            Ok(Work::AttributeEdit { value, bytes })
        },
        "attribute_small_clone" | "attribute_near_limit_clone" => {
            Ok(Work::Attribute(attribute_prepared.clone()))
        },
        "element_small_malformed" => match formatcode2::read(&inputs.element_malformed_small) {
            Ok(value) => Ok(Work::Element(value)),
            Err(_) => Ok(Work::Rejected),
        },
        "element_near_limit_malformed" => {
            match formatcode2::read(&inputs.element_malformed_near_limit) {
                Ok(value) => Ok(Work::Element(value)),
                Err(_) => Ok(Work::Rejected),
            }
        },
        "attribute_small_malformed" => {
            match formatcode2::read_attribute(&inputs.attribute_malformed_small) {
                Ok(value) => Ok(Work::Attribute(value)),
                Err(_) => Ok(Work::Rejected),
            }
        },
        "attribute_near_limit_malformed" => {
            match formatcode2::read_attribute(&inputs.attribute_malformed_near_limit) {
                Ok(value) => Ok(Work::Attribute(value)),
                Err(_) => Ok(Work::Rejected),
            }
        },
        other => Err(format!("unknown lane {other}").into()),
    }
}

#[derive(Clone, Copy)]
struct Sample {
    elapsed_ns: u64,
    allocation: AllocDelta,
    expected_success: bool,
    actual_success: bool,
    semantic_ok: bool,
    source_shared: bool,
    output_exact: bool,
    sink_bytes: u64,
    sink_hash: u64,
    sink_writes: u64,
}

fn validate_work(
    inputs: &Inputs,
    lane: &str,
    work: &Work,
) -> Result<(bool, bool, bool, u64, u64, u64)> {
    let source = if is_attribute(lane) {
        if is_near(lane) {
            &inputs.attribute_near_limit
        } else {
            &inputs.attribute_small
        }
    } else if is_near(lane) {
        &inputs.element_near_limit
    } else {
        &inputs.element_small
    };
    let prepared = if is_attribute(lane) {
        if is_near(lane) {
            SourceValue::Attribute(&inputs.attribute_near_prepared)
        } else {
            SourceValue::Attribute(&inputs.attribute_small_prepared)
        }
    } else if is_near(lane) {
        SourceValue::Element(&inputs.element_near_prepared)
    } else {
        SourceValue::Element(&inputs.element_small_prepared)
    };
    let expected_value = if is_attribute(lane) && !is_near(lane) {
        SMALL_ATTRIBUTE_VALUE
    } else if is_near(lane) {
        &inputs.near_value
    } else {
        SMALL_VALUE
    };
    let updated_value = if is_attribute(lane) && !is_near(lane) {
        SMALL_ATTRIBUTE_UPDATED
    } else if is_near(lane) {
        &inputs.near_updated
    } else {
        SMALL_UPDATED
    };
    let shared = if is_attribute(lane) {
        if is_near(lane) {
            &inputs.attribute_near_limit_shared
        } else {
            &inputs.attribute_small_shared
        }
    } else if is_near(lane) {
        &inputs.element_near_limit_shared
    } else {
        &inputs.element_small_shared
    };
    let empty_sink = (0_u64, 0_u64, 0_u64);
    match (lane, work) {
        ("element_small_read" | "element_near_limit_read", Work::Element(value)) => Ok((
            value.value() == expected_value && value.source() == Some(source.as_slice()),
            false,
            true,
            empty_sink.0,
            empty_sink.1,
            empty_sink.2,
        )),
        ("element_small_read_shared" | "element_near_limit_read_shared", Work::Element(value)) => {
            let source_shared = value
                .source()
                .zip(Some(shared.as_ref()))
                .is_some_and(|(actual, original)| actual.as_ptr() == original.as_ptr());
            Ok((
                value.value() == expected_value && value.source() == Some(source.as_slice()),
                source_shared,
                true,
                empty_sink.0,
                empty_sink.1,
                empty_sink.2,
            ))
        },
        ("attribute_small_read" | "attribute_near_limit_read", Work::Attribute(value)) => Ok((
            value.value() == expected_value && value.source() == Some(source.as_slice()),
            false,
            true,
            empty_sink.0,
            empty_sink.1,
            empty_sink.2,
        )),
        (
            "attribute_small_read_shared" | "attribute_near_limit_read_shared",
            Work::Attribute(value),
        ) => {
            let source_shared = value
                .source()
                .zip(Some(shared.as_ref()))
                .is_some_and(|(actual, original)| actual.as_ptr() == original.as_ptr());
            Ok((
                value.value() == expected_value && value.source() == Some(source.as_slice()),
                source_shared,
                true,
                empty_sink.0,
                empty_sink.1,
                empty_sink.2,
            ))
        },
        ("element_small_noop" | "element_near_limit_noop", Work::Bytes(bytes)) => {
            let reopened = formatcode2::read(bytes)?;
            Ok((
                reopened.value() == expected_value,
                false,
                bytes == source,
                empty_sink.0,
                empty_sink.1,
                empty_sink.2,
            ))
        },
        ("element_small_write_to" | "element_near_limit_write_to", Work::Sink(sink)) => {
            let expected_hash = hash_bytes(FNV_OFFSET, source);
            Ok((
                sink.bytes == as_u64(source.len()) && sink.hash == expected_hash && sink.writes > 0,
                false,
                sink.bytes == as_u64(source.len()) && sink.hash == expected_hash,
                sink.bytes,
                sink.hash,
                sink.writes,
            ))
        },
        ("attribute_small_noop" | "attribute_near_limit_noop", Work::Bytes(bytes)) => {
            let reopened = formatcode2::read_attribute(bytes)?;
            Ok((
                reopened.value() == expected_value,
                false,
                bytes == source,
                empty_sink.0,
                empty_sink.1,
                empty_sink.2,
            ))
        },
        ("attribute_small_write_to" | "attribute_near_limit_write_to", Work::Sink(sink)) => {
            let expected_hash = hash_bytes(FNV_OFFSET, source);
            Ok((
                sink.bytes == as_u64(source.len()) && sink.hash == expected_hash && sink.writes > 0,
                false,
                sink.bytes == as_u64(source.len()) && sink.hash == expected_hash,
                sink.bytes,
                sink.hash,
                sink.writes,
            ))
        },
        (
            "element_small_scalar_edit" | "element_near_limit_scalar_edit",
            Work::ElementEdit { value, bytes },
        ) => {
            let reopened = formatcode2::read(bytes)?;
            Ok((
                value.value() == updated_value && reopened.value() == updated_value,
                false,
                reopened.source() == Some(bytes.as_slice()),
                empty_sink.0,
                empty_sink.1,
                empty_sink.2,
            ))
        },
        (
            "attribute_small_scalar_edit" | "attribute_near_limit_scalar_edit",
            Work::AttributeEdit { value, bytes },
        ) => {
            let reopened = formatcode2::read_attribute(bytes)?;
            Ok((
                value.value() == updated_value && reopened.value() == updated_value,
                false,
                reopened.source() == Some(bytes.as_slice()),
                empty_sink.0,
                empty_sink.1,
                empty_sink.2,
            ))
        },
        ("element_small_clone" | "element_near_limit_clone", Work::Element(value)) => {
            let SourceValue::Element(original) = prepared else {
                return Err("element clone lane has an attribute preparation".into());
            };
            let source_shared = value
                .source()
                .zip(original.source())
                .is_some_and(|(actual, original)| actual.as_ptr() == original.as_ptr());
            Ok((
                value.value() == expected_value,
                source_shared,
                true,
                empty_sink.0,
                empty_sink.1,
                empty_sink.2,
            ))
        },
        ("attribute_small_clone" | "attribute_near_limit_clone", Work::Attribute(value)) => {
            let SourceValue::Attribute(original) = prepared else {
                return Err("attribute clone lane has an element preparation".into());
            };
            let source_shared = value
                .source()
                .zip(original.source())
                .is_some_and(|(actual, original)| actual.as_ptr() == original.as_ptr());
            Ok((
                value.value() == expected_value,
                source_shared,
                true,
                empty_sink.0,
                empty_sink.1,
                empty_sink.2,
            ))
        },
        (
            "element_small_malformed"
            | "element_near_limit_malformed"
            | "attribute_small_malformed"
            | "attribute_near_limit_malformed",
            Work::Rejected,
        ) => Ok((true, false, true, empty_sink.0, empty_sink.1, empty_sink.2)),
        (lane, _) if is_malformed(lane) => {
            Err(format!("malformed lane {lane} returned an accepted value").into())
        },
        _ => Err(format!("lane {lane} returned an unexpected work result").into()),
    }
}

enum SourceValue<'a> {
    Element(&'a Element),
    Attribute(&'a Attribute),
}

fn run_sample(inputs: &Inputs, lane: &str) -> Result<Sample> {
    reset_counters();
    let before = AllocSnapshot::now();
    let started = Instant::now();
    let work = operation(inputs, lane)?;
    black_box(&work);
    let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
    let allocation = before.delta(AllocSnapshot::now());
    let actual_success = !matches!(&work, Work::Rejected);
    let expected = expected_success(lane);
    if actual_success != expected {
        return Err(
            format!("lane {lane} returned success={actual_success}, expected={expected}").into(),
        );
    }
    if allocation.invalid || allocation.failed != 0 || !allocation.balanced() {
        return Err(format!("allocator accounting failed for {lane}").into());
    }
    let (semantic_ok, source_shared, output_exact, sink_bytes, sink_hash, sink_writes) =
        validate_work(inputs, lane, &work)?;
    if !semantic_ok || (is_malformed(lane) && !output_exact) {
        return Err(format!("post-timer semantic gate failed for {lane}").into());
    }
    Ok(Sample {
        elapsed_ns,
        allocation,
        expected_success: expected,
        actual_success,
        semantic_ok,
        source_shared,
        output_exact,
        sink_bytes,
        sink_hash,
        sink_writes,
    })
}

fn run_lane(inputs: &Inputs, lane: &str, warmup: usize, samples: usize) -> Result<String> {
    for _ in 0..warmup {
        let _ = run_sample(inputs, lane)?;
    }
    let mut values = Vec::with_capacity(samples);
    for _ in 0..samples {
        values.push(run_sample(inputs, lane)?);
    }
    let input_bytes = if is_attribute(lane) {
        if is_malformed(lane) {
            if is_near(lane) {
                inputs.attribute_malformed_near_limit.len()
            } else {
                inputs.attribute_malformed_small.len()
            }
        } else if is_near(lane) {
            inputs.attribute_near_limit.len()
        } else {
            inputs.attribute_small.len()
        }
    } else if is_malformed(lane) {
        if is_near(lane) {
            inputs.element_malformed_near_limit.len()
        } else {
            inputs.element_malformed_small.len()
        }
    } else if is_near(lane) {
        inputs.element_near_limit.len()
    } else {
        inputs.element_small.len()
    };
    let mut output = String::new();
    write!(
        &mut output,
        "{{\"schema\":\"chart-formatcode2-profile-v1\",\"lane\":\"{lane}\",\"warmup\":{warmup},\"sample_count\":{},\"expected_success\":{},\"input_bytes\":{input_bytes},\"samples\":[",
        values.len(),
        expected_success(lane)
    )?;
    for (index, sample) in values.iter().enumerate() {
        if index != 0 {
            output.push(',');
        }
        write!(
            &mut output,
            "{{\"elapsed_ns\":{},\"requested_alloc_bytes\":{},\"direct_allocated_bytes\":{},\"realloc_old_bytes\":{},\"realloc_new_bytes\":{},\"deallocated_bytes\":{},\"alloc_calls\":{},\"realloc_calls\":{},\"dealloc_calls\":{},\"live_before\":{},\"live_after\":{},\"peak_live_delta\":{},\"alloc_balance_ok\":{},\"alloc_invalid\":{},\"alloc_failed\":{},\"expected_success\":{},\"actual_success\":{},\"semantic_ok\":{},\"source_shared\":{},\"output_exact\":{},\"sink_bytes\":{},\"sink_hash\":{},\"sink_writes\":{}}}",
            sample.elapsed_ns,
            sample.allocation.requested(),
            sample.allocation.direct,
            sample.allocation.realloc_old,
            sample.allocation.realloc_new,
            sample.allocation.deallocated,
            sample.allocation.calls,
            sample.allocation.realloc_calls,
            sample.allocation.dealloc_calls,
            sample.allocation.live_before,
            sample.allocation.live_after,
            sample.allocation.peak_delta,
            sample.allocation.balanced(),
            sample.allocation.invalid,
            sample.allocation.failed,
            sample.expected_success,
            sample.actual_success,
            sample.semantic_ok,
            sample.source_shared,
            sample.output_exact,
            sample.sink_bytes,
            sample.sink_hash,
            sample.sink_writes,
        )?;
    }
    output.push_str("]}\n");
    Ok(output)
}

fn parse_positive(value: &str, name: &str) -> Result<usize> {
    let value = value.parse::<usize>()?;
    if value == 0 {
        return Err(format!("{name} must be positive").into());
    }
    Ok(value)
}

fn main() -> Result<()> {
    let mut lane = None;
    let mut warmup = DEFAULT_WARMUP;
    let mut samples = DEFAULT_SAMPLES;
    let mut args = env::args().skip(1);
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--lane" => lane = Some(args.next().ok_or("missing --lane value")?),
            "--warmup" => {
                warmup = parse_positive(&args.next().ok_or("missing --warmup value")?, "--warmup")?
            },
            "--samples" => {
                samples =
                    parse_positive(&args.next().ok_or("missing --samples value")?, "--samples")?
            },
            "--help" | "-h" => {
                println!("--lane <element_*|attribute_*> [--warmup N] [--samples N]");
                return Ok(());
            },
            other => return Err(format!("unknown argument {other}").into()),
        }
    }
    let lane = lane.ok_or("--lane is required")?;
    if !LANES.contains(&lane.as_str()) {
        return Err(format!("unknown lane {lane}").into());
    }
    allocator_counter_self_test()?;
    let inputs = Inputs::load()?;
    print!("{}", run_lane(&inputs, &lane, warmup, samples)?);
    Ok(())
}
