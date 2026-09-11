#![allow(
    unsafe_code,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::print_stdout,
    reason = "the opt-in profile owns a process-local allocator observer and emits JSON"
)]

use litchi_drawingml::theme::family::Family;
use litchi_drawingml::theme::family::part;
use std::alloc::{GlobalAlloc, Layout, System};
use std::env;
use std::error::Error;
use std::fmt::Write as FmtWrite;
use std::fs;
use std::hint::black_box;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::Instant;

type BoxError = Box<dyn Error + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

const MAX_NAME: &str = "hardening profile update";
const UNKNOWN_URI: &str = "urn:litchi:unknown-theme-family-owner";
const NATIVE_ID: &str = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}";
const NATIVE_VID: &str = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}";

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

// SAFETY: each method forwards the allocator contract to System; atomics only
// observe successful operations and preserve the caller's pointer semantics.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller provides a valid layout.
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller provides a valid layout.
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
        // SAFETY: the caller provides the valid pointer/layout contract.
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
    let live = LIVE_BYTES.fetch_add(size, Ordering::Relaxed).saturating_add(size);
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
            .and_then(|v| v.checked_add(self.realloc_new))
            .and_then(|v| v.checked_sub(self.realloc_old))
            .and_then(|v| v.checked_sub(self.deallocated))
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

#[derive(Clone, Copy)]
struct Sample {
    elapsed_ns: u64,
    alloc: AllocDelta,
    expected_success: bool,
    actual_success: bool,
}

struct Inputs {
    native: Vec<u8>,
    family_free: Vec<u8>,
    family: Family,
    updated: Family,
    unknown_32: Vec<u8>,
    unknown_1000: Vec<u8>,
    duplicate: Vec<u8>,
}

impl Inputs {
    fn load(path: &str) -> Result<Self> {
        let native = fs::read(path)?;
        let snapshot = part::read(&native)?;
        let family = snapshot.family().ok_or("native fixture has no family")?.clone();
        let mut updated = family.clone();
        updated.set_name(MAX_NAME)?;
        let family_free = part::remove_family(&native)?;
        let unknown_32 = unknown_theme(&family_free, 32, 200)?;
        let unknown_1000 = unknown_theme(&family_free, 1000, 200)?;
        let fragment = raw_family_fragment(&native)?;
        let duplicate = duplicate_theme(&native, &fragment)?;
        let checks = [
            part::read(&unknown_32)?.family().is_none(),
            part::read(&unknown_1000)?.family().is_none(),
            part::read(&duplicate).is_err(),
        ];
        if checks != [true, true, true] {
            return Err("hardening fixture semantic preflight failed".into());
        }
        Ok(Self {
            native,
            family_free,
            family,
            updated,
            unknown_32,
            unknown_1000,
            duplicate,
        })
    }
}

fn raw_family_fragment(source: &[u8]) -> Result<Vec<u8>> {
    let start = source
        .windows(b"<thm15:themeFamily".len())
        .position(|window| window == b"<thm15:themeFamily")
        .ok_or("native family root missing")?;
    let tail = &source[start..];
    let end = if let Some(offset) = tail.windows(2).position(|window| window == b"/>") {
        start + offset + 2
    } else {
        let offset = tail
            .windows(b"</thm15:themeFamily>".len())
            .position(|window| window == b"</thm15:themeFamily>")
            .ok_or("native family close missing")?;
        start + offset + b"</thm15:themeFamily>".len()
    };
    Ok(source[start..end].to_vec())
}

fn unknown_theme(source: &[u8], count: usize, bindings: usize) -> Result<Vec<u8>> {
    let root_end = source
        .windows(b"</a:theme>".len())
        .position(|w| w == b"</a:theme>")
        .ok_or("theme root close missing")?;
    let mut declarations = String::new();
    for index in 0..bindings {
        write!(&mut declarations, " xmlns:p{index}=\"urn:litchi:scope:{index}\"")?;
    }
    let mut body = String::with_capacity(count * 180);
    for index in 0..count {
        write!(
            &mut body,
            "<u:themeFamily xmlns:u=\"urn:litchi:unknown-family:{index}\" name=\"F{index}\" id=\"{NATIVE_ID}\" vid=\"{NATIVE_VID}\"/>"
        )?;
    }
    let extension = format!(
        "<a:extLst><a:ext uri=\"{UNKNOWN_URI}\">{body}</a:ext></a:extLst>"
    );
    let mut root = source.to_vec();
    let open = root
        .windows(b"<a:theme".len())
        .position(|w| w == b"<a:theme")
        .ok_or("theme root open missing")?;
    let open_end = root[open..]
        .iter()
        .position(|byte| *byte == b'>')
        .map(|offset| open + offset)
        .ok_or("theme root open end missing")?;
    root.splice(open_end..open_end, declarations.bytes());
    let insertion = root_end + declarations.len();
    root.splice(insertion..insertion, extension.bytes());
    Ok(root)
}

fn duplicate_theme(source: &[u8], fragment: &[u8]) -> Result<Vec<u8>> {
    let start = source
        .windows(fragment.len())
        .position(|window| window == fragment)
        .ok_or("native family fragment missing")?;
    let mut replacement = Vec::with_capacity(fragment.len() * 2);
    replacement.extend_from_slice(fragment);
    replacement.extend_from_slice(fragment);
    let mut output = Vec::with_capacity(source.len() + fragment.len());
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&source[start + fragment.len()..]);
    Ok(output)
}

fn operation(inputs: &Inputs, lane: &str) -> Result<(bool, bool)> {
    match lane {
        "native_read" => {
            let parsed = part::read(&inputs.native)?;
            Ok((true, parsed.family().is_some()))
        },
        "native_replace" => {
            let output = part::replace_family_with_limit(
                &inputs.native,
                &inputs.updated,
                part::MAX_XML_BYTES,
            )?;
            let parsed = part::read(&output)?;
            Ok((true, parsed.family().is_some_and(|f| f.name() == MAX_NAME)))
        },
        "native_remove" => {
            let output = part::remove_family_with_limit(&inputs.native, part::MAX_XML_BYTES)?;
            let parsed = part::read(&output)?;
            Ok((true, parsed.family().is_none()))
        },
        "native_add" => {
            let output = part::add_family_with_uri_limit(
                &inputs.family_free,
                &inputs.family,
                part::NATIVE_EXTENSION_URI,
                part::MAX_XML_BYTES,
            )?;
            let parsed = part::read(&output)?;
            Ok((true, parsed.family().is_some()))
        },
        "unknown_32" => {
            let parsed = part::read(&inputs.unknown_32)?;
            Ok((true, parsed.family().is_none()))
        },
        "unknown_1000" => {
            let parsed = part::read(&inputs.unknown_1000)?;
            Ok((true, parsed.family().is_none()))
        },
        "duplicate" => Ok((false, part::read(&inputs.duplicate).is_ok())),
        "limit_replace" => Ok((false, part::replace_family_with_limit(&inputs.native, &inputs.updated, 1).is_ok())),
        "limit_add" => Ok((false, part::add_family_with_uri_limit(&inputs.family_free, &inputs.family, part::NATIVE_EXTENSION_URI, 1).is_ok())),
        other => Err(format!("unknown lane {other}").into()),
    }
}

fn run_lane(inputs: &Inputs, lane: &str, warmup: usize, samples: usize) -> Result<String> {
    let expected_success = !matches!(lane, "duplicate" | "limit_replace" | "limit_add");
    for _ in 0..warmup {
        let (expected, actual) = operation(inputs, lane)?;
        if expected != expected_success || actual != expected_success {
            return Err(format!("warmup expectation failed for {lane}: {expected} {actual}").into());
        }
    }
    let mut values = Vec::with_capacity(samples);
    for _ in 0..samples {
        reset_counters();
        let before = AllocSnapshot::now();
        let started = Instant::now();
        let outcome = operation(inputs, lane);
        let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
        let (expected, actual) = match outcome {
            Ok((expected, actual)) => (expected, actual),
            Err(error) => return Err(format!("lane {lane} failed: {error}").into()),
        };
        black_box((expected, actual));
        let after = AllocSnapshot::now();
        let alloc = before.delta(after);
        if expected != expected_success || actual != expected_success {
            return Err(format!("sample expectation failed for {lane}: {expected} {actual}").into());
        }
        if alloc.invalid || alloc.failed != 0 || !alloc.balanced() {
            return Err(format!("allocator accounting failed for {lane}").into());
        }
        values.push(Sample {
            elapsed_ns,
            alloc,
            expected_success,
            actual_success: actual,
        });
    }
    let mut json = String::new();
    write!(
        &mut json,
        "{{\"schema\":\"theme-family-hardening-profile-v1\",\"lane\":\"{lane}\",\"warmup\":{warmup},\"sample_count\":{},\"expected_success\":{},\"samples\":[",
        values.len(), expected_success
    )?;
    for (index, sample) in values.iter().enumerate() {
        if index != 0 {
            json.push(',');
        }
        write!(
            &mut json,
            "{{\"elapsed_ns\":{},\"requested_alloc_bytes\":{},\"direct_allocated_bytes\":{},\"realloc_old_bytes\":{},\"realloc_new_bytes\":{},\"deallocated_bytes\":{},\"alloc_calls\":{},\"realloc_calls\":{},\"dealloc_calls\":{},\"live_before\":{},\"live_after\":{},\"peak_live_delta\":{},\"alloc_balance_ok\":{},\"alloc_invalid\":{},\"alloc_failed\":{},\"expected_success\":{},\"actual_success\":{}}}",
            sample.elapsed_ns,
            sample.alloc.requested(),
            sample.alloc.direct,
            sample.alloc.realloc_old,
            sample.alloc.realloc_new,
            sample.alloc.deallocated,
            sample.alloc.calls,
            sample.alloc.realloc_calls,
            sample.alloc.dealloc_calls,
            sample.alloc.live_before,
            sample.alloc.live_after,
            sample.alloc.peak_delta,
            sample.alloc.balanced(),
            sample.alloc.invalid,
            sample.alloc.failed,
            sample.expected_success,
            sample.actual_success,
        )?;
    }
    json.push_str("]}\n");
    Ok(json)
}

fn main() -> Result<()> {
    let mut fixture = None;
    let mut lane = None;
    let mut warmup = 2usize;
    let mut samples = 20usize;
    let mut args = env::args().skip(1);
    while let Some(arg) = args.next() {
        match arg.as_str() {
            "--fixture" => fixture = Some(args.next().ok_or("missing --fixture value")?),
            "--lane" => lane = Some(args.next().ok_or("missing --lane value")?),
            "--warmup" => warmup = args.next().ok_or("missing --warmup value")?.parse()?,
            "--samples" => samples = args.next().ok_or("missing --samples value")?.parse()?,
            "--help" | "-h" => {
                println!("--fixture PATH --lane <native_read|native_replace|native_remove|native_add|unknown_32|unknown_1000|duplicate|limit_replace|limit_add> [--warmup N] [--samples N]");
                return Ok(());
            },
            other => return Err(format!("unknown argument {other}").into()),
        }
    }
    let fixture = fixture.ok_or("--fixture is required")?;
    let lane = lane.ok_or("--lane is required")?;
    let inputs = Inputs::load(&fixture)?;
    print!("{}", run_lane(&inputs, &lane, warmup, samples)?);
    Ok(())
}
