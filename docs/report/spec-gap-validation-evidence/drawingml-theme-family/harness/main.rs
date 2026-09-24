//! Small isolated profile for the shared DrawingML `themeFamily` owner.
//!
//! The harness deliberately keeps the timed closures narrow.  Fixture
//! construction, source hashing, byte identity checks, and inverse checks run
//! outside those closures so allocator and elapsed-time observations describe
//! the owner operation rather than the report machinery.

#![allow(
    unsafe_code,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::print_stdout,
    clippy::shadow_reuse,
    reason = "the opt-in profile owns a process-local allocator observer and emits JSON"
)]

use litchi_drawingml::theme::family::{self, Snapshot};
use std::alloc::{GlobalAlloc, Layout, System};
use std::env;
use std::error::Error;
use std::fmt::Write as _;
use std::hint::black_box;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::Instant;

const DEFAULT_WARMUP: usize = 3;
const DEFAULT_SAMPLES: usize = 30;
const SMALL: &[u8] = br#"<thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" xmlns="" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/>"#;
const OPAQUE_MARKER: &[u8] = br#"v:opaqueMarker="keep-theme-family-extension""#;
const CHANGED_NAME: &str = "Office Theme changed";
const OPAQUE_EXTENSION_COUNT: usize = 96;

#[global_allocator]
static GLOBAL_ALLOCATOR: CountingAllocator = CountingAllocator;

struct CountingAllocator;

static ALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static REALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DEALLOC_CALLS: AtomicU64 = AtomicU64::new(0);
static DIRECT_ALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);
static REALLOC_OLD_BYTES: AtomicU64 = AtomicU64::new(0);
static REALLOC_NEW_BYTES: AtomicU64 = AtomicU64::new(0);
static DEALLOCATED_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_FAILED: AtomicU64 = AtomicU64::new(0);
static ALLOC_INVALID: AtomicBool = AtomicBool::new(false);

// SAFETY: each method forwards the allocator contract unchanged to `System`;
// the atomics only observe successful operations.
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
        DEALLOCATED_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
        subtract_live(layout.size());
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        // SAFETY: the caller supplies the valid pointer/layout contract.
        let result = unsafe { System.realloc(pointer, layout, new_size) };
        if result.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            REALLOC_CALLS.fetch_add(1, Ordering::Relaxed);
            REALLOC_OLD_BYTES.fetch_add(as_u64(layout.size()), Ordering::Relaxed);
            REALLOC_NEW_BYTES.fetch_add(as_u64(new_size), Ordering::Relaxed);
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
    DIRECT_ALLOCATED_BYTES.fetch_add(as_u64(size), Ordering::Relaxed);
    observe_growth(size);
}

fn observe_growth(size: usize) {
    let size = as_u64(size);
    let live = ALIVE_ADD
        .fetch_add(size, Ordering::Relaxed)
        .saturating_add(size);
    update_peak(live);
}

// Keep the live-byte update separate from the public allocation-call counter.
// This alias avoids a second fetch in the hot wrapper while retaining one
// authoritative counter for all direct and realloc growth.
static ALIVE_ADD: AtomicU64 = AtomicU64::new(0);

fn subtract_live(size: usize) {
    let size = as_u64(size);
    let before = ALIVE_ADD.fetch_sub(size, Ordering::Relaxed);
    if before < size {
        ALLOC_INVALID.store(true, Ordering::Release);
    }
}

fn update_peak(live: u64) {
    let mut current = PEAK_LIVE_BYTES.load(Ordering::Relaxed);
    while live > current {
        match PEAK_LIVE_BYTES.compare_exchange_weak(
            current,
            live,
            Ordering::Relaxed,
            Ordering::Relaxed,
        ) {
            Ok(_) => break,
            Err(observed) => current = observed,
        }
    }
}

fn live_bytes() -> u64 {
    ALIVE_ADD.load(Ordering::Acquire)
}

#[derive(Clone, Copy)]
struct AllocSnapshot {
    calls: u64,
    realloc_calls: u64,
    dealloc_calls: u64,
    direct_allocated_bytes: u64,
    realloc_old_bytes: u64,
    realloc_new_bytes: u64,
    deallocated_bytes: u64,
    live_bytes: u64,
    peak_live_bytes: u64,
    failed: u64,
    invalid: bool,
}

impl AllocSnapshot {
    fn now() -> Self {
        Self {
            calls: ALLOC_CALLS.load(Ordering::Acquire),
            realloc_calls: REALLOC_CALLS.load(Ordering::Acquire),
            dealloc_calls: DEALLOC_CALLS.load(Ordering::Acquire),
            direct_allocated_bytes: DIRECT_ALLOCATED_BYTES.load(Ordering::Acquire),
            realloc_old_bytes: REALLOC_OLD_BYTES.load(Ordering::Acquire),
            realloc_new_bytes: REALLOC_NEW_BYTES.load(Ordering::Acquire),
            deallocated_bytes: DEALLOCATED_BYTES.load(Ordering::Acquire),
            live_bytes: live_bytes(),
            peak_live_bytes: PEAK_LIVE_BYTES.load(Ordering::Acquire),
            failed: ALLOC_FAILED.load(Ordering::Acquire),
            invalid: ALLOC_INVALID.load(Ordering::Acquire),
        }
    }

    fn delta(self, after: Self) -> AllocDelta {
        AllocDelta {
            calls: after.calls.saturating_sub(self.calls),
            realloc_calls: after.realloc_calls.saturating_sub(self.realloc_calls),
            dealloc_calls: after.dealloc_calls.saturating_sub(self.dealloc_calls),
            direct_allocated_bytes: after
                .direct_allocated_bytes
                .saturating_sub(self.direct_allocated_bytes),
            realloc_old_bytes: after
                .realloc_old_bytes
                .saturating_sub(self.realloc_old_bytes),
            realloc_new_bytes: after
                .realloc_new_bytes
                .saturating_sub(self.realloc_new_bytes),
            deallocated_bytes: after
                .deallocated_bytes
                .saturating_sub(self.deallocated_bytes),
            live_before: self.live_bytes,
            live_after: after.live_bytes,
            peak_live_bytes: after.peak_live_bytes.saturating_sub(self.peak_live_bytes),
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
    direct_allocated_bytes: u64,
    realloc_old_bytes: u64,
    realloc_new_bytes: u64,
    deallocated_bytes: u64,
    live_before: u64,
    live_after: u64,
    peak_live_bytes: u64,
    failed: u64,
    invalid: bool,
}

impl AllocDelta {
    fn requested_bytes(self) -> u64 {
        self.direct_allocated_bytes
            .saturating_add(self.realloc_new_bytes)
    }

    fn balance_ok(self) -> bool {
        let Some(expected_live) = self
            .live_before
            .checked_add(self.direct_allocated_bytes)
            .and_then(|live| live.checked_add(self.realloc_new_bytes))
            .and_then(|live| live.checked_sub(self.realloc_old_bytes))
            .and_then(|live| live.checked_sub(self.deallocated_bytes))
        else {
            return false;
        };
        expected_live == self.live_after
    }
}

fn reset_alloc() {
    let live = live_bytes();
    PEAK_LIVE_BYTES.store(live, Ordering::Release);
    ALLOC_CALLS.store(0, Ordering::Release);
    REALLOC_CALLS.store(0, Ordering::Release);
    DEALLOC_CALLS.store(0, Ordering::Release);
    DIRECT_ALLOCATED_BYTES.store(0, Ordering::Release);
    REALLOC_OLD_BYTES.store(0, Ordering::Release);
    REALLOC_NEW_BYTES.store(0, Ordering::Release);
    DEALLOCATED_BYTES.store(0, Ordering::Release);
    ALLOC_FAILED.store(0, Ordering::Release);
    ALLOC_INVALID.store(false, Ordering::Release);
}

fn allocator_counter_self_test() -> Result<(), BoxError> {
    reset_alloc();
    let before = AllocSnapshot::now();
    let layout = Layout::from_size_align(8, std::mem::align_of::<usize>())
        .map_err(|_| "allocator self-test layout is invalid")?;
    // SAFETY: this deliberately exercises the process-local observer with a
    // valid allocation, reallocation, and matching deallocation.
    let pointer = unsafe { std::alloc::alloc(layout) };
    if pointer.is_null() {
        reset_alloc();
        return Err("allocator self-test allocation failed".into());
    }
    // SAFETY: `pointer` was returned by the preceding allocation with
    // `layout`, and the new size is non-zero.
    let resized = unsafe { std::alloc::realloc(pointer, layout, 32) };
    if resized.is_null() {
        // SAFETY: a failed reallocation leaves the original allocation valid.
        unsafe { std::alloc::dealloc(pointer, layout) };
        reset_alloc();
        return Err("allocator self-test reallocation failed".into());
    }
    let resized_layout = Layout::from_size_align(32, layout.align())
        .expect("allocator self-test resized layout is constant-valid");
    // SAFETY: `resized` was returned by the preceding reallocation with the
    // same alignment and its new size.
    unsafe { std::alloc::dealloc(resized, resized_layout) };
    let delta = before.delta(AllocSnapshot::now());
    let valid = delta.calls == 1
        && delta.realloc_calls == 1
        && delta.dealloc_calls == 1
        && delta.direct_allocated_bytes == 8
        && delta.realloc_old_bytes == 8
        && delta.realloc_new_bytes == 32
        && delta.deallocated_bytes == 32
        && delta.requested_bytes() == 40
        && delta.balance_ok()
        && !delta.invalid
        && delta.failed == 0;
    reset_alloc();
    if !valid {
        return Err("allocator self-test counters did not balance".into());
    }
    Ok(())
}

fn observe_source(snapshot: &Snapshot, fixture: &Fixture) -> Result<Observation, BoxError> {
    let bytes = snapshot.xml_bytes();
    let value = snapshot.value();
    let source_equal = bytes == fixture.bytes;
    let semantic_ok = value.name() == fixture.name
        && value.id().as_str() == fixture.id
        && value.variant_id().as_str() == fixture.vid;
    let opaque_preserved = !fixture.opaque
        || bytes
            .windows(OPAQUE_MARKER.len())
            .any(|w| w == OPAQUE_MARKER);
    if !source_equal && !opaque_preserved {
        return Err("source bytes and opaque marker were both lost".into());
    }
    Ok(Observation {
        source_hash: digest(bytes),
        semantic_hash: semantic_digest(value),
        source_equal,
        semantic_ok,
        opaque_preserved,
    })
}

fn semantic_digest(value: &family::Family) -> u64 {
    let mut digest = 0xcbf29ce484222325;
    digest = digest_bytes(digest, value.name().as_bytes());
    digest = digest_bytes(digest, value.id().as_str().as_bytes());
    digest_bytes(digest, value.variant_id().as_str().as_bytes())
}

fn digest(bytes: &[u8]) -> u64 {
    digest_bytes(0xcbf29ce484222325, bytes)
}

fn digest_bytes(mut digest: u64, bytes: &[u8]) -> u64 {
    for byte in bytes {
        digest ^= u64::from(*byte);
        digest = digest.wrapping_mul(0x100000001b3);
    }
    digest
}

type BoxError = Box<dyn Error + Send + Sync>;

#[derive(Clone, Copy, PartialEq, Eq)]
enum FixtureKind {
    Small,
    Opaque,
}

impl FixtureKind {
    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "small" => Ok(Self::Small),
            "opaque" => Ok(Self::Opaque),
            other => Err(format!("unknown fixture {other:?}; expected small or opaque").into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Small => "small",
            Self::Opaque => "opaque",
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Operation {
    Read,
    Clone,
    Noop,
    Change,
}

impl Operation {
    fn parse(value: &str) -> Result<Self, BoxError> {
        match value {
            "read" => Ok(Self::Read),
            "clone" => Ok(Self::Clone),
            "noop" => Ok(Self::Noop),
            "change" => Ok(Self::Change),
            other => Err(format!(
                "unknown operation {other:?}; expected read, clone, noop, or change"
            )
            .into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Read => "read",
            Self::Clone => "clone",
            Self::Noop => "noop",
            Self::Change => "change",
        }
    }
}

struct Args {
    fixture: FixtureKind,
    operation: Operation,
    warmup: usize,
    samples: usize,
}

impl Args {
    fn parse() -> Result<Self, BoxError> {
        let mut fixture = None;
        let mut operation = None;
        let mut warmup = DEFAULT_WARMUP;
        let mut samples = DEFAULT_SAMPLES;
        let mut args = env::args().skip(1);
        while let Some(arg) = args.next() {
            let mut value = || -> Result<String, BoxError> {
                args.next()
                    .ok_or_else(|| format!("missing value for {arg}").into())
            };
            match arg.as_str() {
                "--fixture" => fixture = Some(FixtureKind::parse(&value()?)?),
                "--operation" => operation = Some(Operation::parse(&value()?)?),
                "--warmup" => warmup = parse_positive(&value()?, "--warmup")?,
                "--samples" => samples = parse_positive(&value()?, "--samples")?,
                "--help" | "-h" => {
                    return Err("usage: drawingml-theme-family-profile --fixture <small|opaque> --operation <read|clone|noop|change> [--warmup N] [--samples N]".into());
                },
                other => return Err(format!("unknown argument {other:?}").into()),
            }
        }
        Ok(Self {
            fixture: fixture.ok_or("--fixture is required")?,
            operation: operation.ok_or("--operation is required")?,
            warmup,
            samples,
        })
    }
}

fn parse_positive(value: &str, name: &str) -> Result<usize, BoxError> {
    let parsed = value.parse::<usize>()?;
    if parsed == 0 {
        return Err(format!("{name} must be positive").into());
    }
    Ok(parsed)
}

struct Fixture {
    kind: FixtureKind,
    bytes: Vec<u8>,
    name: &'static str,
    id: &'static str,
    vid: &'static str,
    opaque: bool,
    shape_ok: bool,
}

impl Fixture {
    fn load(kind: FixtureKind) -> Result<Self, BoxError> {
        let bytes = match kind {
            FixtureKind::Small => SMALL.to_vec(),
            FixtureKind::Opaque => opaque_fixture(),
        };
        let fixture = Self {
            kind,
            bytes,
            name: if kind == FixtureKind::Opaque {
                "Office Theme & Vendor"
            } else {
                "Office Theme"
            },
            id: "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}",
            vid: "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}",
            opaque: kind == FixtureKind::Opaque,
            shape_ok: false,
        };
        let snapshot = Snapshot::from_xml(fixture.bytes.clone())?;
        if fixture.opaque {
            validate_opaque_shape(&fixture.bytes)?;
        }
        let observation = observe_source(&snapshot, &fixture)?;
        if !observation.semantic_ok {
            return Err("fixture typed values do not match the profile manifest".into());
        }
        Ok(Self {
            shape_ok: true,
            ..fixture
        })
    }
}

fn validate_opaque_shape(bytes: &[u8]) -> Result<(), BoxError> {
    let required_fragments = [
        b"<thm15:themeFamily " as &[u8],
        b"xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\"",
        b"xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"",
        b"name=\"Office Theme &amp; Vendor\"",
        b" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\"",
        b" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\"",
        b"<thm15:extLst>",
        b"</thm15:extLst>",
    ];
    for fragment in required_fragments {
        if !bytes
            .windows(fragment.len())
            .any(|window| window == fragment)
        {
            return Err(format!(
                "opaque profile fixture is missing required shape fragment {:?}",
                String::from_utf8_lossy(fragment)
            )
            .into());
        }
    }
    let family_extension = [b"<thm15:ext ".as_slice(), b"<thm15:ext>".as_slice()];
    if family_extension
        .iter()
        .any(|marker| bytes.windows(marker.len()).any(|window| window == *marker))
    {
        return Err("opaque profile fixture uses the family namespace for an extension".into());
    }
    let extension_start = bytes
        .windows(b"<a:ext uri=\"".len())
        .filter(|window| *window == b"<a:ext uri=\"")
        .count();
    let extension_end = bytes
        .windows(b"</a:ext>".len())
        .filter(|window| *window == b"</a:ext>")
        .count();
    if extension_start != OPAQUE_EXTENSION_COUNT || extension_end != OPAQUE_EXTENSION_COUNT {
        return Err(format!(
            "opaque profile fixture has {extension_start} a:ext starts and {extension_end} ends; expected {OPAQUE_EXTENSION_COUNT}"
        )
        .into());
    }
    Ok(())
}

fn opaque_fixture() -> Vec<u8> {
    let mut xml = String::with_capacity(24 * 1024);
    xml.push_str("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n<!-- retained-before -->\n");
    xml.push_str(
        "<thm15:themeFamily xmlns:thm15=\"http://schemas.microsoft.com/office/thememl/2012/main\" xmlns:v=\"urn:vendor:theme-family\" xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" name=\"Office Theme &amp; Vendor\" id=\"{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}\" vid=\"{4A3C46E8-61CC-4603-A589-7422A47A8E4A}\" v:opaqueMarker=\"keep-theme-family-extension\">\n",
    );
    xml.push_str("  <thm15:extLst>\n");
    for index in 0..96 {
        let _ = writeln!(
            xml,
            "    <a:ext uri=\"{{00000000-0000-0000-0000-{index:012X}}}\" v:slot=\"slot-{index}\"><v:opaque v:key=\"key-{index}\">opaque payload {index:04} &amp; retained</v:opaque><!-- ext-comment-{index} --></a:ext>",
        );
    }
    xml.push_str("  </thm15:extLst>\n</thm15:themeFamily>\n<!-- retained-after -->\n");
    xml.into_bytes()
}

#[derive(Clone, Copy)]
struct Observation {
    source_hash: u64,
    semantic_hash: u64,
    source_equal: bool,
    semantic_ok: bool,
    opaque_preserved: bool,
}

#[derive(Clone, Copy)]
struct Sample {
    elapsed_ns: u64,
    allocation: AllocDelta,
    observation: Observation,
    source_shared: bool,
    changed_ok: bool,
    inverse_ok: bool,
}

fn run_sample(
    fixture: &Fixture,
    operation: Operation,
    prepared: Option<&Snapshot>,
) -> Result<Sample, BoxError> {
    reset_alloc();
    let allocation_before = AllocSnapshot::now();
    let started = Instant::now();
    let mut source_shared = false;
    let mut changed_ok = true;
    let mut inverse_ok = true;
    let result = match operation {
        Operation::Read => Snapshot::from_xml(&fixture.bytes),
        Operation::Clone => Ok(black_box(
            prepared
                .ok_or("clone operation lacks prepared snapshot")?
                .clone(),
        )),
        Operation::Noop => {
            let snapshot = prepared.ok_or("no-op operation lacks prepared snapshot")?;
            let commit = snapshot.edit().commit()?;
            Ok(black_box(commit.into_snapshot()))
        },
        Operation::Change => {
            let snapshot = prepared.ok_or("change operation lacks prepared snapshot")?;
            let mut edit = snapshot.edit();
            edit.set_name(CHANGED_NAME)?;
            Ok(black_box(edit.commit()?.into_snapshot()))
        },
    }?;
    let elapsed_ns = u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX);
    let allocation = allocation_before.delta(AllocSnapshot::now());
    if allocation.invalid || allocation.failed != 0 || !allocation.balance_ok() {
        return Err("allocator counters did not balance for the timed sample".into());
    }

    // Everything below is intentionally after the timer and allocator snapshot.
    // In particular, no source hash or large-byte comparison is timed.
    let mut observation = observe_source(&result, fixture)?;
    match operation {
        Operation::Read => {
            if !observation.source_equal {
                return Err("read did not retain the complete source fragment".into());
            }
        },
        Operation::Clone => {
            source_shared = result.xml_bytes().as_ptr()
                == prepared
                    .ok_or("clone operation lacks prepared snapshot")?
                    .xml_bytes()
                    .as_ptr();
        },
        Operation::Noop => {
            let original = prepared.ok_or("no-op operation lacks prepared snapshot")?;
            source_shared = result.xml_bytes().as_ptr() == original.xml_bytes().as_ptr()
                && result.xml_bytes().len() == original.xml_bytes().len();
            if !observation.source_equal || !source_shared {
                return Err("semantic no-op did not preserve and share source bytes".into());
            }
        },
        Operation::Change => {
            changed_ok = result.value().name() == CHANGED_NAME
                && result.xml_bytes() != fixture.bytes.as_slice();
            observation.semantic_ok = result.value().name() == CHANGED_NAME
                && result.value().id().as_str() == fixture.id
                && result.value().variant_id().as_str() == fixture.vid;
            if fixture.opaque {
                changed_ok &= result
                    .xml_bytes()
                    .windows(OPAQUE_MARKER.len())
                    .any(|w| w == OPAQUE_MARKER);
            }
            let original = prepared.ok_or("change operation lacks prepared snapshot")?;
            let mut edit = original.edit();
            edit.set_name(CHANGED_NAME)?;
            let commit = edit.commit()?;
            let restored = commit.patch().clone().inverse().apply(&result)?;
            inverse_ok = restored.xml_bytes() == original.xml_bytes();
            if !changed_ok || !inverse_ok {
                return Err("changed edit failed its source-preservation or inverse gate".into());
            }
        },
    }
    Ok(Sample {
        elapsed_ns,
        allocation,
        observation,
        source_shared,
        changed_ok,
        inverse_ok,
    })
}

fn main() -> Result<(), BoxError> {
    let args = Args::parse()?;
    allocator_counter_self_test()?;
    let fixture = Fixture::load(args.fixture)?;
    let prepared = if args.operation == Operation::Read {
        None
    } else {
        Some(Snapshot::from_xml(fixture.bytes.clone())?)
    };
    for _ in 0..args.warmup {
        let _ = run_sample(&fixture, args.operation, prepared.as_ref())?;
    }
    let mut samples = Vec::with_capacity(args.samples);
    for _ in 0..args.samples {
        samples.push(run_sample(&fixture, args.operation, prepared.as_ref())?);
    }
    println!(
        "{}",
        report_json(&fixture, args.operation, args.warmup, &samples)
    );
    Ok(())
}

fn report_json(
    fixture: &Fixture,
    operation: Operation,
    warmup: usize,
    samples: &[Sample],
) -> String {
    let mut elapsed = samples
        .iter()
        .map(|sample| sample.elapsed_ns)
        .collect::<Vec<_>>();
    let mut allocated = samples
        .iter()
        .map(|sample| sample.allocation.requested_bytes())
        .collect::<Vec<_>>();
    let mut peak = samples
        .iter()
        .map(|sample| sample.allocation.peak_live_bytes)
        .collect::<Vec<_>>();
    elapsed.sort_unstable();
    allocated.sort_unstable();
    peak.sort_unstable();
    let first = samples[0];
    let expected_hash = digest(&fixture.bytes);
    let mut output = String::new();
    output.push('{');
    json_str(&mut output, "schema", "drawingml-theme-family-profile-v1");
    json_str(&mut output, "fixture", fixture.kind.name());
    json_str(&mut output, "operation", operation.name());
    json_num(&mut output, "input_bytes", fixture.bytes.len() as u64);
    json_num(&mut output, "input_hash", expected_hash);
    json_num(&mut output, "warmup", warmup as u64);
    json_num(&mut output, "sample_count", samples.len() as u64);
    json_bool(&mut output, "allocator_instrumented", true);
    json_bool(&mut output, "allocator_self_test", true);
    json_bool(&mut output, "peak_live_is_delta", true);
    json_bool(&mut output, "fixture_shape_ok", fixture.shape_ok);
    json_num(&mut output, "p50_ns", percentile(&elapsed, 50));
    json_num(&mut output, "p95_ns", percentile(&elapsed, 95));
    json_num(&mut output, "p99_ns", percentile(&elapsed, 99));
    json_num(&mut output, "allocated_p50", percentile(&allocated, 50));
    json_num(&mut output, "allocated_p95", percentile(&allocated, 95));
    json_num(&mut output, "peak_live_p50", percentile(&peak, 50));
    json_num(&mut output, "peak_live_p95", percentile(&peak, 95));
    json_bool(
        &mut output,
        "source_shared_all",
        samples.iter().all(|s| s.source_shared),
    );
    json_bool(
        &mut output,
        "changed_ok_all",
        samples.iter().all(|s| s.changed_ok),
    );
    json_bool(
        &mut output,
        "inverse_ok_all",
        samples.iter().all(|s| s.inverse_ok),
    );
    json_bool(
        &mut output,
        "semantic_ok_all",
        samples.iter().all(|s| s.observation.semantic_ok),
    );
    json_bool(
        &mut output,
        "opaque_preserved_all",
        samples.iter().all(|s| s.observation.opaque_preserved),
    );
    json_num(
        &mut output,
        "first_source_hash",
        first.observation.source_hash,
    );
    json_num(
        &mut output,
        "first_semantic_hash",
        first.observation.semantic_hash,
    );
    output.push_str(",\"samples\":[");
    for (index, sample) in samples.iter().enumerate() {
        if index != 0 {
            output.push(',');
        }
        output.push('{');
        json_num(&mut output, "elapsed_ns", sample.elapsed_ns);
        json_num(&mut output, "alloc_calls", sample.allocation.calls);
        json_num(
            &mut output,
            "realloc_calls",
            sample.allocation.realloc_calls,
        );
        json_num(
            &mut output,
            "dealloc_calls",
            sample.allocation.dealloc_calls,
        );
        json_num(
            &mut output,
            "allocated_bytes",
            sample.allocation.requested_bytes(),
        );
        json_num(
            &mut output,
            "direct_allocated_bytes",
            sample.allocation.direct_allocated_bytes,
        );
        json_num(
            &mut output,
            "realloc_old_bytes",
            sample.allocation.realloc_old_bytes,
        );
        json_num(
            &mut output,
            "realloc_new_bytes",
            sample.allocation.realloc_new_bytes,
        );
        json_num(
            &mut output,
            "deallocated_bytes",
            sample.allocation.deallocated_bytes,
        );
        json_num(
            &mut output,
            "peak_live_bytes",
            sample.allocation.peak_live_bytes,
        );
        json_num(&mut output, "live_before", sample.allocation.live_before);
        json_num(&mut output, "live_after", sample.allocation.live_after);
        json_num(&mut output, "alloc_failed", sample.allocation.failed);
        json_bool(&mut output, "alloc_invalid", sample.allocation.invalid);
        json_bool(
            &mut output,
            "alloc_balance_ok",
            sample.allocation.balance_ok(),
        );
        json_bool(&mut output, "source_shared", sample.source_shared);
        json_bool(&mut output, "changed_ok", sample.changed_ok);
        json_bool(&mut output, "inverse_ok", sample.inverse_ok);
        output.push('}');
    }
    output.push_str("]}");
    output
}

fn percentile(values: &[u64], percentile: usize) -> u64 {
    if values.is_empty() {
        return 0;
    }
    let rank = values.len().saturating_mul(percentile).saturating_add(99) / 100;
    values[rank.saturating_sub(1).min(values.len() - 1)]
}

fn json_str(output: &mut String, key: &str, value: &str) {
    json_separator(output);
    let _ = write!(output, "\"{key}\":\"{value}\"");
}

fn json_num(output: &mut String, key: &str, value: u64) {
    json_separator(output);
    let _ = write!(output, "\"{key}\":{value}");
}

fn json_bool(output: &mut String, key: &str, value: bool) {
    json_separator(output);
    let _ = write!(output, "\"{key}\":{value}");
}

fn json_separator(output: &mut String) {
    if !output.ends_with('{') && !output.ends_with(',') {
        output.push(',');
    }
}
