//! Bounded process-isolated profile for the XLSB host of DrawingML
//! `themeFamily` metadata.
//!
//! The harness deliberately keeps source hashing, pointer identity checks, and
//! semantic gates outside timed/allocation intervals.  The transaction lanes
//! measure one forward host edit and its exact inverse on a prepared workbook;
//! they do not claim production latency for a different caller workflow.

#![allow(
    unsafe_code,
    clippy::arbitrary_source_item_ordering,
    clippy::cast_possible_truncation,
    clippy::cast_precision_loss,
    clippy::print_stdout,
    clippy::shadow_reuse,
    clippy::similar_names,
    reason = "the opt-in profile owns a process-local allocator observer and emits JSON"
)]

use litchi_core::{OwnedSource, ReadAt, SourceVersion};
use litchi_drawingml::theme::family::part as family_part;
use litchi_drawingml::theme::{Color, Slot, Theme, codec};
use litchi_xlsb::theme::Family;
use litchi_xlsb::{SourceBackedWorkbook, Workbook};
use sha2::{Digest, Sha256};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::alloc::{GlobalAlloc, Layout, System};
use std::env;
use std::error::Error;
use std::fmt::Write as FmtWrite;
use std::fs;
use std::hint::black_box;
use std::io::{self, Cursor, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::time::Instant;

type BoxError = Box<dyn Error + Send + Sync>;
type Result<T> = std::result::Result<T, BoxError>;

const DEFAULT_WARMUP: usize = 3;
const DEFAULT_SAMPLES: usize = 30;
const THEME_PART: &str = "/xl/theme/theme1.xml";
const FAMILY_START: &[u8] = b"<thm15:themeFamily";
const FAMILY_CLOSE: &[u8] = b"</thm15:themeFamily>";
const FAMILY_NAMESPACE: &str = "http://schemas.microsoft.com/office/thememl/2012/main";
const DRAWINGML_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const VENDOR_NAMESPACE: &str = "urn:litchi:theme-family-profile";
const NATIVE_EXTENSION_URI: &str = "{05A4C25C-085E-4340-85A3-A5531E510DB2}";
const NATIVE_ID: &str = "{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}";
const NATIVE_VID: &str = "{4A3C46E8-61CC-4603-A589-7422A47A8E4A}";
const OPAQUE_EXTENSION_COUNT: usize = 96;
const UPDATED_NAME: &str = "Office Theme profile update";

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
static LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static PEAK_LIVE_BYTES: AtomicU64 = AtomicU64::new(0);
static ALLOC_FAILED: AtomicU64 = AtomicU64::new(0);
static ALLOC_INVALID: AtomicBool = AtomicBool::new(false);

// SAFETY: every method forwards the allocator contract unchanged to System;
// atomics only observe successful operations.
unsafe impl GlobalAlloc for CountingAllocator {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
        let pointer = unsafe { System.alloc(layout) };
        if pointer.is_null() {
            ALLOC_FAILED.fetch_add(1, Ordering::Relaxed);
        } else {
            observe_alloc(layout.size());
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        // SAFETY: the caller supplies a valid allocation layout.
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
    let live = LIVE_BYTES
        .fetch_add(size, Ordering::Relaxed)
        .saturating_add(size);
    update_peak(live);
}

fn subtract_live(size: usize) {
    let size = as_u64(size);
    let before = LIVE_BYTES.fetch_sub(size, Ordering::Relaxed);
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
            live_bytes: LIVE_BYTES.load(Ordering::Acquire),
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
            peak_live_delta: after.peak_live_bytes.saturating_sub(self.peak_live_bytes),
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
    peak_live_delta: u64,
    failed: u64,
    invalid: bool,
}

impl AllocDelta {
    /// Requested allocation bytes: direct allocation requests plus the new
    /// size requested by successful reallocations.  Old realloc sizes are
    /// reported separately and are never double-counted here.
    fn requested_bytes(self) -> u64 {
        self.direct_allocated_bytes
            .saturating_add(self.realloc_new_bytes)
    }

    fn balance_ok(self) -> bool {
        let Some(expected_live) = self
            .live_before
            .checked_add(self.direct_allocated_bytes)
            .and_then(|value| value.checked_add(self.realloc_new_bytes))
            .and_then(|value| value.checked_sub(self.realloc_old_bytes))
            .and_then(|value| value.checked_sub(self.deallocated_bytes))
        else {
            return false;
        };
        expected_live == self.live_after
    }
}

fn reset_alloc() {
    PEAK_LIVE_BYTES.store(LIVE_BYTES.load(Ordering::Acquire), Ordering::Release);
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

fn allocator_counter_self_test() -> Result<()> {
    reset_alloc();
    let before = AllocSnapshot::now();
    let layout = Layout::from_size_align(8, std::mem::align_of::<usize>())?;
    // SAFETY: the profile deliberately exercises one valid allocation.
    let pointer = unsafe { std::alloc::alloc(layout) };
    if pointer.is_null() {
        reset_alloc();
        return Err("allocator self-test allocation failed".into());
    }
    // SAFETY: pointer and layout are from the preceding allocation.
    let resized = unsafe { std::alloc::realloc(pointer, layout, 32) };
    if resized.is_null() {
        // SAFETY: failed realloc leaves the original allocation valid.
        unsafe { std::alloc::dealloc(pointer, layout) };
        reset_alloc();
        return Err("allocator self-test reallocation failed".into());
    }
    let resized_layout = Layout::from_size_align(32, layout.align())?;
    // SAFETY: resized uses the same alignment and the requested new size.
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

#[derive(Default)]
struct ReadCounterState {
    calls: AtomicU64,
    requested_bytes: AtomicU64,
    returned_bytes: AtomicU64,
}

#[derive(Clone, Copy, Default)]
struct ReadCounters {
    calls: u64,
    requested_bytes: u64,
    returned_bytes: u64,
}

impl ReadCounterState {
    fn snapshot(&self) -> ReadCounters {
        ReadCounters {
            calls: self.calls.load(Ordering::Acquire),
            requested_bytes: self.requested_bytes.load(Ordering::Acquire),
            returned_bytes: self.returned_bytes.load(Ordering::Acquire),
        }
    }
}

impl ReadCounters {
    fn delta(self, after: Self) -> Self {
        Self {
            calls: after.calls.saturating_sub(self.calls),
            requested_bytes: after.requested_bytes.saturating_sub(self.requested_bytes),
            returned_bytes: after.returned_bytes.saturating_sub(self.returned_bytes),
        }
    }
}

struct CountingReadAt {
    inner: OwnedSource,
    counters: Arc<ReadCounterState>,
}

impl CountingReadAt {
    fn new(bytes: &[u8], counters: Arc<ReadCounterState>) -> Self {
        Self {
            inner: OwnedSource::new(bytes.to_vec()),
            counters,
        }
    }
}

impl ReadAt for CountingReadAt {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.counters.calls.fetch_add(1, Ordering::Relaxed);
        self.counters
            .requested_bytes
            .fetch_add(as_u64(output.len()), Ordering::Relaxed);
        let read = self.inner.read_at(offset, output)?;
        self.counters
            .returned_bytes
            .fetch_add(as_u64(read), Ordering::Relaxed);
        Ok(read)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

struct SourceFixture {
    source: Arc<dyn ReadAt>,
    counters: Arc<ReadCounterState>,
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum FixtureKind {
    Native,
    Opaque,
}

impl FixtureKind {
    fn parse(value: &str) -> Result<Self> {
        match value {
            "native" => Ok(Self::Native),
            "opaque" => Ok(Self::Opaque),
            other => Err(format!("unknown fixture {other:?}; expected native or opaque").into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Native => "native",
            Self::Opaque => "opaque",
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum Operation {
    CodecRead,
    MetadataRead,
    SourceRead,
    FamilyClone,
    Noop,
    Add,
    Update,
    Remove,
    BaseEdit,
}

impl Operation {
    fn parse(value: &str) -> Result<Self> {
        Ok(match value {
            "codec_read" => Self::CodecRead,
            "metadata_read" => Self::MetadataRead,
            "source_read" => Self::SourceRead,
            "family_clone" => Self::FamilyClone,
            "noop" => Self::Noop,
            "add" => Self::Add,
            "update" => Self::Update,
            "remove" => Self::Remove,
            "base_edit" => Self::BaseEdit,
            other => {
                return Err(format!(
                    "unknown operation {other:?}; expected codec_read, metadata_read, source_read, family_clone, noop, add, update, remove, or base_edit"
                )
                .into())
            },
        })
    }

    const fn name(self) -> &'static str {
        match self {
            Self::CodecRead => "codec_read",
            Self::MetadataRead => "metadata_read",
            Self::SourceRead => "source_read",
            Self::FamilyClone => "family_clone",
            Self::Noop => "noop",
            Self::Add => "add",
            Self::Update => "update",
            Self::Remove => "remove",
            Self::BaseEdit => "base_edit",
        }
    }

    const fn timing_scope(self) -> &'static str {
        match self {
            Self::CodecRead => {
                "shared DrawingML Theme codec read of the selected Theme XML; no XLSB graph or family discovery"
            },
            Self::MetadataRead => {
                "fresh eager XLSB Workbook open, Theme read, and family metadata discovery"
            },
            Self::SourceRead => {
                "fresh source-backed XLSB open, selected Theme read, and family metadata discovery"
            },
            Self::FamilyClone => "clone of one parsed shared DrawingML Family value",
            Self::Noop => {
                "prepared eager Workbook Theme no-op commit and source-checked no-op publication"
            },
            Self::Add => {
                "prepared eager Workbook family add, forward publication, and exact inverse publication"
            },
            Self::Update => {
                "prepared eager Workbook family metadata update, forward publication, and exact inverse publication"
            },
            Self::Remove => {
                "prepared eager Workbook family removal, forward publication, and exact inverse publication"
            },
            Self::BaseEdit => {
                "prepared eager Workbook base Theme edit, forward publication, and exact inverse publication"
            },
        }
    }
}

struct Args {
    fixture: PathBuf,
    kind: FixtureKind,
    operation: Operation,
    warmup: usize,
    samples: usize,
}

impl Args {
    fn parse() -> Result<Self> {
        let mut fixture = None;
        let mut kind = None;
        let mut operation = None;
        let mut warmup = DEFAULT_WARMUP;
        let mut samples = DEFAULT_SAMPLES;
        let mut args = env::args().skip(1);
        while let Some(arg) = args.next() {
            let mut next = || -> Result<String> {
                args.next()
                    .ok_or_else(|| format!("missing value for {arg}").into())
            };
            match arg.as_str() {
                "--fixture" => fixture = Some(PathBuf::from(next()?)),
                "--kind" => kind = Some(FixtureKind::parse(&next()?)?),
                "--operation" => operation = Some(Operation::parse(&next()?)?),
                "--warmup" => warmup = parse_positive(&next()?, "--warmup")?,
                "--samples" => samples = parse_positive(&next()?, "--samples")?,
                "--help" | "-h" => {
                    return Err("usage: xlsb-theme-family-profile --fixture PATH --kind <native|opaque> --operation <codec_read|metadata_read|source_read|family_clone|noop|add|update|remove|base_edit> [--warmup N] [--samples N]".into())
                },
                other => return Err(format!("unknown argument {other:?}").into()),
            }
        }
        Ok(Self {
            fixture: fixture.ok_or("--fixture is required")?,
            kind: kind.ok_or("--kind is required")?,
            operation: operation.ok_or("--operation is required")?,
            warmup,
            samples,
        })
    }
}

fn parse_positive(value: &str, name: &str) -> Result<usize> {
    let parsed = value.parse::<usize>()?;
    if parsed == 0 {
        return Err(format!("{name} must be positive").into());
    }
    Ok(parsed)
}

struct Fixture {
    kind: FixtureKind,
    path: PathBuf,
    package_bytes: Vec<u8>,
    absent_package_bytes: Vec<u8>,
    theme_xml: Vec<u8>,
    family_xml: Vec<u8>,
    family: Family,
    updated_family: Family,
    source: SourceFixture,
    package_digest: u64,
    theme_digest: u64,
    family_digest: u64,
    package_sha256: String,
    theme_sha256: String,
    family_sha256: String,
    theme_semantic_digest: u64,
    family_semantic_digest: u64,
    opaque_shape_ok: bool,
    foreign_namespace_declarations: usize,
}

impl Fixture {
    fn load(path: &Path, kind: FixtureKind) -> Result<Self> {
        let native = fs::read(path)?;
        let package_bytes = if kind == FixtureKind::Opaque {
            let workbook = Workbook::new(Cursor::new(native.clone()))?;
            let snapshot = workbook.theme()?.ok_or("native workbook has no Theme")?;
            let xml = opaque_theme_xml(snapshot.source_xml())?;
            replace_theme_xml(&native, &xml)?
        } else {
            native
        };
        let workbook = Workbook::new(Cursor::new(package_bytes.clone()))?;
        let snapshot = workbook.theme()?.ok_or("fixture has no Theme")?;
        let family = snapshot
            .family()
            .ok_or("fixture has no DrawingML themeFamily metadata")?
            .clone();
        let theme_xml = snapshot.source_xml().to_vec();
        let family_xml = family
            .source()
            .ok_or("parsed Family did not retain source XML")?
            .to_vec();
        let opaque_shape_ok = if kind == FixtureKind::Opaque {
            validate_opaque_shape(&family_xml)?;
            true
        } else {
            true
        };
        let foreign_namespace_declarations = if kind == FixtureKind::Opaque {
            OPAQUE_EXTENSION_COUNT
        } else {
            0
        };
        let absent_theme_xml = remove_family(&theme_xml)?;
        let absent_package_bytes = replace_theme_xml(&package_bytes, &absent_theme_xml)?;
        let mut updated_family = family.clone();
        updated_family.set_name(UPDATED_NAME)?;
        let counters = Arc::new(ReadCounterState::default());
        let source: Arc<dyn ReadAt> =
            Arc::new(CountingReadAt::new(&package_bytes, Arc::clone(&counters)));
        Ok(Self {
            kind,
            path: path.to_path_buf(),
            package_digest: digest(&package_bytes),
            theme_digest: digest(&theme_xml),
            family_digest: digest(&family_xml),
            package_sha256: sha256_hex(&package_bytes),
            theme_sha256: sha256_hex(&theme_xml),
            family_sha256: sha256_hex(&family_xml),
            theme_semantic_digest: theme_digest(snapshot.theme()),
            family_semantic_digest: family_digest(&family),
            package_bytes,
            absent_package_bytes,
            theme_xml,
            family_xml,
            family,
            updated_family,
            source: SourceFixture { source, counters },
            opaque_shape_ok,
            foreign_namespace_declarations,
        })
    }
}

fn replace_theme_xml(source: &[u8], xml: &[u8]) -> Result<Vec<u8>> {
    let archive = ArchiveReader::new(source)?;
    let names = archive
        .file_names()
        .map(ToOwned::to_owned)
        .collect::<Vec<_>>();
    let mut writer = StreamingArchiveWriter::new();
    for name in names {
        let payload = if name == THEME_PART.trim_start_matches('/') {
            xml.to_vec()
        } else {
            archive.read(&name)?
        };
        writer.write_deflated(&name, &payload)?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn opaque_theme_xml(source: &[u8]) -> Result<Vec<u8>> {
    // The native fixture's declaration line uses CRLF.  The workbook's
    // publication path accepts compact XML for a generated synthetic package,
    // so remove formatting line endings in this setup-only derivative.
    let compact_source = source
        .iter()
        .copied()
        .filter(|byte| *byte != b'\r' && *byte != b'\n')
        .collect::<Vec<_>>();
    let source = compact_source.as_slice();
    let native_outer = format!("<a:ext uri=\"{NATIVE_EXTENSION_URI}\">");
    if !source
        .windows(native_outer.len())
        .any(|window| window == native_outer.as_bytes())
    {
        return Err("native Theme family host has no required Office extension URI".into());
    }
    let start = source
        .windows(FAMILY_START.len())
        .position(|window| window == FAMILY_START)
        .ok_or("native Theme has no themeFamily root")?;
    let root_end = source[start..]
        .windows(2)
        .position(|window| window == b"/>")
        .map(|offset| start + offset)
        .ok_or("native themeFamily is not self-closing")?;
    let mut replacement = Vec::with_capacity(24 * 1024);
    replacement.extend_from_slice(&source[start..root_end]);
    replacement.extend_from_slice(
        format!(" xmlns:a=\"{DRAWINGML_NAMESPACE}\" xmlns:v=\"{VENDOR_NAMESPACE}\"><thm15:extLst>")
            .as_bytes(),
    );
    for index in 0..OPAQUE_EXTENSION_COUNT {
        write!(
            &mut replacement,
            "<a:ext uri=\"{{00000000-0000-0000-0000-{index:012X}}}\" v:slot=\"slot-{index}\"><v:opaque xmlns:vn{index}=\"urn:litchi:theme-family-profile:foreign:{index}\" v:key=\"key-{index}\"><vn{index}:payload vn{index}:kind=\"foreign-{index}\">opaque payload {index:04} &amp; retained</vn{index}:payload></v:opaque><!-- profile-ext-{index} --></a:ext>"
        )?;
    }
    replacement.extend_from_slice(b"</thm15:extLst></thm15:themeFamily>");
    let mut output = Vec::with_capacity(source.len() + replacement.len());
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(&replacement);
    output.extend_from_slice(&source[root_end + 2..]);
    Ok(output)
}

fn remove_family(source: &[u8]) -> Result<Vec<u8>> {
    let start = source
        .windows(FAMILY_START.len())
        .position(|window| window == FAMILY_START)
        .ok_or("Theme has no themeFamily root")?;
    let after_root = &source[start..];
    let close = if let Some(offset) = after_root.windows(2).position(|w| w == b"/>") {
        start + offset + 2
    } else {
        let offset = after_root
            .windows(FAMILY_CLOSE.len())
            .position(|w| w == FAMILY_CLOSE)
            .ok_or("themeFamily closing tag is missing")?;
        start + offset + FAMILY_CLOSE.len()
    };
    let mut output = Vec::with_capacity(source.len().saturating_sub(close - start));
    output.extend_from_slice(&source[..start]);
    output.extend_from_slice(&source[close..]);
    Ok(output)
}

fn validate_opaque_shape(source: &[u8]) -> Result<()> {
    if !source
        .windows(FAMILY_START.len())
        .any(|window| window == FAMILY_START)
        || !source
            .windows(b"<thm15:extLst>".len())
            .any(|window| window == b"<thm15:extLst>")
        || !source
            .windows(b"</thm15:extLst>".len())
            .any(|window| window == b"</thm15:extLst>")
    {
        return Err("opaque family fixture lacks required extension-list shape".into());
    }
    let required_root = [
        format!("xmlns:thm15=\"{FAMILY_NAMESPACE}\""),
        format!("xmlns:a=\"{DRAWINGML_NAMESPACE}\""),
        format!("xmlns:v=\"{VENDOR_NAMESPACE}\""),
        format!("name=\"Office Theme\""),
        format!("id=\"{NATIVE_ID}\""),
        format!("vid=\"{NATIVE_VID}\""),
    ];
    if required_root.iter().any(|required| {
        !source
            .windows(required.len())
            .any(|window| window == required.as_bytes())
    }) {
        return Err("opaque family fixture lost native namespace or required attributes".into());
    }
    let count = source
        .windows(b"<a:ext uri=\"".len())
        .filter(|window| *window == b"<a:ext uri=\"")
        .count();
    if count != OPAQUE_EXTENSION_COUNT {
        return Err(format!(
            "opaque family fixture has {count} extensions; expected {OPAQUE_EXTENSION_COUNT}"
        )
        .into());
    }
    for index in 0..OPAQUE_EXTENSION_COUNT {
        let declaration =
            format!("xmlns:vn{index}=\"urn:litchi:theme-family-profile:foreign:{index}\"");
        if !source
            .windows(declaration.len())
            .any(|window| window == declaration.as_bytes())
        {
            return Err(format!(
                "opaque fixture is missing foreign namespace declaration vn{index}"
            )
            .into());
        }
    }
    if source
        .windows(b"<thm15:ext ".len())
        .any(|window| window == b"<thm15:ext ")
    {
        return Err("opaque fixture incorrectly uses thm15:ext for an extension".into());
    }
    Ok(())
}

fn theme_digest(theme: &Theme) -> u64 {
    let mut value = digest_bytes(0xcbf29ce484222325, theme.name.as_bytes());
    value = digest_bytes(value, theme.colors.name().as_bytes());
    for slot in Slot::ALL {
        value = digest_bytes(value, slot.token().as_bytes());
        if let Some(color) = theme.colors.color(slot) {
            value = match color {
                Color::Rgb(rgb) => digest_bytes(digest_bytes(value, b"rgb"), rgb.as_bytes()),
                Color::System { kind, last } => {
                    let mut value = digest_bytes(value, b"system");
                    value = digest_bytes(value, kind.token().as_bytes());
                    if let Some(last) = last {
                        value = digest_bytes(value, last.as_bytes());
                    }
                    value
                },
            };
        }
    }
    value = digest_bytes(value, theme.fonts.name().as_bytes());
    value = face_digest(value, theme.fonts.major());
    face_digest(value, theme.fonts.minor())
}

fn face_digest(mut value: u64, face: &litchi_drawingml::theme::Face) -> u64 {
    value = digest_bytes(value, face.latin.as_bytes());
    value = digest_bytes(value, face.east_asian.as_bytes());
    value = digest_bytes(value, face.complex_script.as_bytes());
    for script in &face.scripts {
        value = digest_bytes(value, script.code.as_bytes());
        value = digest_bytes(value, script.typeface.as_bytes());
    }
    value
}

fn family_digest(value: &Family) -> u64 {
    let mut digest = digest_bytes(0xcbf29ce484222325, value.name().as_bytes());
    digest = digest_bytes(digest, value.id().as_str().as_bytes());
    digest_bytes(digest, value.variant_id().as_str().as_bytes())
}

fn digest(bytes: &[u8]) -> u64 {
    digest_bytes(0xcbf29ce484222325, bytes)
}

fn sha256_hex(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

fn digest_bytes(mut digest: u64, bytes: &[u8]) -> u64 {
    for byte in bytes {
        digest ^= u64::from(*byte);
        digest = digest.wrapping_mul(0x100000001b3);
    }
    digest
}

struct HostState {
    workbook: Workbook,
    operation: Operation,
    baseline_family: Option<Family>,
    target_family: Option<Family>,
    updated_family: Option<Family>,
    base_theme: Theme,
    changed_theme: Theme,
    baseline_theme_xml: Vec<u8>,
    baseline_theme_without_family: Vec<u8>,
    baseline_family_xml: Option<Vec<u8>>,
    baseline_source_ptr: usize,
}

impl HostState {
    fn new(fixture: &Fixture, operation: Operation) -> Result<Self> {
        let bytes = if operation == Operation::Add {
            &fixture.absent_package_bytes
        } else {
            &fixture.package_bytes
        };
        let workbook = Workbook::new(Cursor::new(bytes.clone()))?;
        let snapshot = workbook.theme()?.ok_or("prepared workbook has no Theme")?;
        let baseline_family = snapshot.family().cloned();
        let baseline_family_xml = baseline_family
            .as_ref()
            .and_then(|value| value.source().map(ToOwned::to_owned));
        let baseline_theme_without_family = family_part::remove_family(snapshot.source_xml())?;
        let target_family = (operation == Operation::Add).then(|| fixture.family.clone());
        let updated_family =
            (operation == Operation::Update).then(|| fixture.updated_family.clone());
        let mut changed_theme = snapshot.theme().clone();
        let accent = Color::rgb("010203")?;
        let colors = changed_theme
            .colors
            .clone()
            .with(litchi_drawingml::theme::Slot::Accent1, accent);
        changed_theme.colors = colors;
        Ok(Self {
            workbook,
            operation,
            baseline_source_ptr: snapshot.source_xml().as_ptr() as usize,
            baseline_theme_xml: snapshot.source_xml().to_vec(),
            baseline_theme_without_family,
            baseline_family_xml,
            base_theme: snapshot.theme().clone(),
            changed_theme,
            baseline_family,
            target_family,
            updated_family,
        })
    }
}

#[derive(Clone, Copy)]
struct Sample {
    elapsed_ns: u64,
    allocation: AllocDelta,
    reads: ReadCounters,
    theme_hash: u64,
    family_hash: u64,
    restored_theme_hash: u64,
    restored_family_hash: u64,
    source_shared: bool,
    semantic_ok: bool,
    forward_preservation_ok: bool,
    preservation_ok: bool,
    inverse_ok: bool,
    changed_ok: bool,
}

fn run_codec_read(fixture: &Fixture) -> Result<Sample> {
    let input = fixture.theme_xml.as_slice();
    reset_alloc();
    let before = AllocSnapshot::now();
    let started = Instant::now();
    let parsed = codec::read(input)?;
    black_box(&parsed);
    let elapsed_ns = elapsed_ns(started);
    let allocation = before.delta(AllocSnapshot::now());
    let theme_hash = theme_digest(&parsed);
    drop(parsed);
    validate_alloc(allocation)?;
    Ok(Sample {
        elapsed_ns,
        allocation,
        reads: ReadCounters::default(),
        theme_hash,
        family_hash: 0,
        restored_theme_hash: theme_hash,
        restored_family_hash: 0,
        source_shared: false,
        semantic_ok: theme_hash == fixture.theme_semantic_digest,
        forward_preservation_ok: true,
        preservation_ok: true,
        inverse_ok: true,
        changed_ok: true,
    })
}

fn run_metadata_read(fixture: &Fixture) -> Result<Sample> {
    // The input Vec is prepared before the timed interval.  This keeps the
    // interval focused on eager workbook/theme parsing rather than caller
    // source cloning; the scope is recorded in the report.
    let input = fixture.package_bytes.clone();
    let input_ptr = input.as_ptr();
    reset_alloc();
    let before = AllocSnapshot::now();
    let started = Instant::now();
    let workbook = Workbook::new(Cursor::new(input))?;
    let snapshot = workbook.theme()?.ok_or("metadata read has no Theme")?;
    let family = snapshot
        .family()
        .ok_or("metadata read has no themeFamily")?;
    black_box((snapshot.theme(), family));
    let elapsed_ns = elapsed_ns(started);
    let allocation = before.delta(AllocSnapshot::now());
    let theme_hash = theme_digest(snapshot.theme());
    let family_hash = family_digest(family);
    let source_shared = snapshot.source_xml().as_ptr() == input_ptr;
    drop(snapshot);
    drop(workbook);
    validate_alloc(allocation)?;
    Ok(Sample {
        elapsed_ns,
        allocation,
        reads: ReadCounters::default(),
        theme_hash,
        family_hash,
        restored_theme_hash: theme_hash,
        restored_family_hash: family_hash,
        source_shared,
        semantic_ok: theme_hash == fixture.theme_semantic_digest
            && family_hash == fixture.family_semantic_digest,
        forward_preservation_ok: true,
        preservation_ok: digest(fixture.theme_xml.as_slice()) == fixture.theme_digest,
        inverse_ok: true,
        changed_ok: true,
    })
}

fn run_source_read(fixture: &Fixture) -> Result<Sample> {
    let reads_before = fixture.source.counters.snapshot();
    reset_alloc();
    let allocation_before = AllocSnapshot::now();
    let started = Instant::now();
    let workbook = SourceBackedWorkbook::from_read_at(Arc::clone(&fixture.source.source))?;
    let view = workbook.theme()?.ok_or("source read has no Theme")?;
    let family = view.family().ok_or("source read has no themeFamily")?;
    black_box((view.theme(), family));
    let elapsed_ns = elapsed_ns(started);
    let allocation = allocation_before.delta(AllocSnapshot::now());
    let reads = reads_before.delta(fixture.source.counters.snapshot());
    let theme_hash = theme_digest(view.theme());
    let family_hash = family_digest(family);
    drop(view);
    drop(workbook);
    validate_alloc(allocation)?;
    Ok(Sample {
        elapsed_ns,
        allocation,
        reads,
        theme_hash,
        family_hash,
        restored_theme_hash: theme_hash,
        restored_family_hash: family_hash,
        source_shared: false,
        semantic_ok: theme_hash == fixture.theme_semantic_digest
            && family_hash == fixture.family_semantic_digest,
        forward_preservation_ok: true,
        preservation_ok: true,
        inverse_ok: true,
        changed_ok: true,
    })
}

fn run_family_clone(fixture: &Fixture) -> Result<Sample> {
    reset_alloc();
    let before = AllocSnapshot::now();
    let started = Instant::now();
    let cloned = black_box(fixture.family.clone());
    let elapsed_ns = elapsed_ns(started);
    let allocation = before.delta(AllocSnapshot::now());
    let source_shared = cloned
        .source()
        .zip(fixture.family.source())
        .is_some_and(|(left, right)| left.as_ptr() == right.as_ptr());
    let family_hash = family_digest(&cloned);
    drop(cloned);
    validate_alloc(allocation)?;
    Ok(Sample {
        elapsed_ns,
        allocation,
        reads: ReadCounters::default(),
        theme_hash: fixture.theme_semantic_digest,
        family_hash,
        restored_theme_hash: fixture.theme_semantic_digest,
        restored_family_hash: family_hash,
        source_shared,
        semantic_ok: family_hash == fixture.family_semantic_digest,
        forward_preservation_ok: true,
        preservation_ok: true,
        inverse_ok: true,
        changed_ok: true,
    })
}

fn run_host_transaction(_fixture: &Fixture, state: &mut HostState) -> Result<Sample> {
    reset_alloc();
    let before = AllocSnapshot::now();
    let started = Instant::now();
    let mut edit = state.workbook.edit_theme()?;
    let staged_changed = match state.operation {
        Operation::Noop => false,
        Operation::Add => {
            edit.set_family(state.target_family.clone().ok_or("missing add Family")?)?
        },
        Operation::Update => edit.set_family(
            state
                .updated_family
                .clone()
                .ok_or("missing update Family")?,
        )?,
        Operation::Remove => edit.remove_family()?,
        Operation::BaseEdit => edit.replace(state.changed_theme.clone())?,
        _ => return Err("invalid host transaction operation".into()),
    };
    let commit = edit.commit()?;
    let commit_changed = commit.changed();
    let forward = black_box(state.workbook.apply_theme(&commit)?);
    let inverse = commit.patch().inverse();
    let restored = black_box(state.workbook.apply_theme_patch(&inverse)?);
    let elapsed_ns = elapsed_ns(started);
    let allocation = before.delta(AllocSnapshot::now());

    // Source hashes and pointer comparisons are intentionally after the timed
    // interval.  Large Theme XML is never hashed by the timed closure.
    let forward_theme_hash = theme_digest(forward.theme());
    let forward_family_hash = forward.family().map(family_digest).unwrap_or(0);
    let restored_theme_hash = theme_digest(restored.theme());
    let restored_family_hash = restored.family().map(family_digest).unwrap_or(0);
    let restored_xml = restored.source_xml();
    let source_shared = state.operation == Operation::Noop
        && restored_xml.as_ptr() == state.baseline_source_ptr as *const u8;
    // Compare the forward source outside the timed interval. Family edits must
    // leave every byte outside the selected owner unchanged; a base Theme edit
    // must retain the complete namespace-rich Family source and its opaque
    // foreign children.
    let forward_preservation_ok = match state.operation {
        Operation::BaseEdit => state.baseline_family_xml.as_ref().is_none_or(|expected| {
            forward
                .family()
                .and_then(Family::source)
                .is_some_and(|actual| actual == expected.as_slice())
        }),
        _ => {
            family_part::remove_family(forward.source_xml())?
                == state.baseline_theme_without_family.as_slice()
        },
    };
    let forward_edit_ok = match state.operation {
        Operation::Add => forward.family().is_some(),
        Operation::Update => forward
            .family()
            .is_some_and(|value| value.name() == UPDATED_NAME),
        Operation::Remove => forward.family().is_none(),
        Operation::BaseEdit => forward_theme_hash != theme_digest(&state.base_theme),
        Operation::Noop => true,
        _ => false,
    };
    let semantic_ok = forward_edit_ok
        && restored_theme_hash == theme_digest(&state.base_theme)
        && restored_family_hash
            == state
                .baseline_family
                .as_ref()
                .map(family_digest)
                .unwrap_or(0);
    let preservation_ok = restored_xml == state.baseline_theme_xml.as_slice()
        && state.baseline_family_xml.as_ref().is_none_or(|expected| {
            restored
                .family()
                .and_then(Family::source)
                .is_some_and(|actual| actual == expected.as_slice())
        });
    let inverse_ok = preservation_ok
        && restored_family_hash
            == state
                .baseline_family
                .as_ref()
                .map(family_digest)
                .unwrap_or(0);
    let changed_ok = match state.operation {
        Operation::Noop => {
            !staged_changed && !commit_changed && forward_family_hash == restored_family_hash
        },
        Operation::Add | Operation::Update | Operation::Remove | Operation::BaseEdit => {
            staged_changed && commit_changed
        },
        _ => false,
    };
    drop(forward);
    drop(restored);
    drop(inverse);
    drop(commit);
    validate_alloc(allocation)?;
    Ok(Sample {
        elapsed_ns,
        allocation,
        reads: ReadCounters::default(),
        theme_hash: forward_theme_hash,
        family_hash: forward_family_hash,
        restored_theme_hash,
        restored_family_hash,
        source_shared,
        semantic_ok,
        forward_preservation_ok,
        preservation_ok,
        inverse_ok,
        changed_ok,
    })
}

fn elapsed_ns(started: Instant) -> u64 {
    u64::try_from(started.elapsed().as_nanos()).unwrap_or(u64::MAX)
}

fn validate_alloc(delta: AllocDelta) -> Result<()> {
    if delta.invalid || delta.failed != 0 || !delta.balance_ok() {
        return Err(format!(
            "allocator counters did not balance: invalid={} failed={} live_before={} live_after={} direct={} realloc_old={} realloc_new={} deallocated={}",
            delta.invalid,
            delta.failed,
            delta.live_before,
            delta.live_after,
            delta.direct_allocated_bytes,
            delta.realloc_old_bytes,
            delta.realloc_new_bytes,
            delta.deallocated_bytes,
        )
        .into());
    }
    Ok(())
}

fn run(args: &Args) -> Result<String> {
    allocator_counter_self_test()?;
    let fixture = Fixture::load(&args.fixture, args.kind)?;
    let mut state = match args.operation {
        Operation::Noop
        | Operation::Add
        | Operation::Update
        | Operation::Remove
        | Operation::BaseEdit => Some(HostState::new(&fixture, args.operation)?),
        _ => None,
    };
    let mut samples = Vec::with_capacity(args.samples);
    for _ in 0..args.warmup {
        let _ = run_one(&fixture, args.operation, state.as_mut())?;
    }
    for _ in 0..args.samples {
        samples.push(run_one(&fixture, args.operation, state.as_mut())?);
    }
    Ok(report_json(&fixture, args.operation, args.warmup, &samples))
}

fn run_one(
    fixture: &Fixture,
    operation: Operation,
    state: Option<&mut HostState>,
) -> Result<Sample> {
    match operation {
        Operation::CodecRead => run_codec_read(fixture),
        Operation::MetadataRead => run_metadata_read(fixture),
        Operation::SourceRead => run_source_read(fixture),
        Operation::FamilyClone => run_family_clone(fixture),
        Operation::Noop
        | Operation::Add
        | Operation::Update
        | Operation::Remove
        | Operation::BaseEdit => run_host_transaction(
            fixture,
            state.ok_or("host transaction has no prepared state")?,
        ),
    }
}

fn report_json(
    fixture: &Fixture,
    operation: Operation,
    warmup: usize,
    samples: &[Sample],
) -> String {
    let mut elapsed = samples.iter().map(|s| s.elapsed_ns).collect::<Vec<_>>();
    let mut allocated = samples
        .iter()
        .map(|s| s.allocation.requested_bytes())
        .collect::<Vec<_>>();
    let mut peak = samples
        .iter()
        .map(|s| s.allocation.peak_live_delta)
        .collect::<Vec<_>>();
    elapsed.sort_unstable();
    allocated.sort_unstable();
    peak.sort_unstable();
    let mut output = String::new();
    output.push('{');
    json_str(&mut output, "schema", "xlsb-theme-family-profile-v1");
    json_str(&mut output, "fixture", fixture.kind.name());
    json_str(&mut output, "operation", operation.name());
    json_str(&mut output, "timing_scope", operation.timing_scope());
    json_str(&mut output, "fixture_path", &fixture.path.to_string_lossy());
    json_num(
        &mut output,
        "package_bytes",
        fixture.package_bytes.len() as u64,
    );
    json_num(&mut output, "theme_bytes", fixture.theme_xml.len() as u64);
    json_num(&mut output, "family_bytes", fixture.family_xml.len() as u64);
    json_num(
        &mut output,
        "foreign_namespace_declarations",
        fixture.foreign_namespace_declarations as u64,
    );
    json_num(&mut output, "package_digest", fixture.package_digest);
    json_num(&mut output, "theme_digest", fixture.theme_digest);
    json_num(&mut output, "family_digest", fixture.family_digest);
    json_str(&mut output, "package_sha256", &fixture.package_sha256);
    json_str(&mut output, "theme_sha256", &fixture.theme_sha256);
    json_str(&mut output, "family_sha256", &fixture.family_sha256);
    json_num(
        &mut output,
        "expected_theme_semantic_digest",
        fixture.theme_semantic_digest,
    );
    json_num(
        &mut output,
        "expected_family_semantic_digest",
        fixture.family_semantic_digest,
    );
    json_num(&mut output, "warmup", warmup as u64);
    json_num(&mut output, "sample_count", samples.len() as u64);
    json_bool(&mut output, "allocator_instrumented", true);
    json_bool(&mut output, "allocator_self_test", true);
    json_bool(&mut output, "peak_live_is_incremental_delta", true);
    json_bool(
        &mut output,
        "timed_input_clone_excluded",
        matches!(operation, Operation::MetadataRead),
    );
    json_bool(&mut output, "fixture_shape_ok", fixture.opaque_shape_ok);
    json_bool(
        &mut output,
        "semantic_ok_all",
        samples.iter().all(|s| s.semantic_ok),
    );
    json_bool(
        &mut output,
        "forward_preservation_ok_all",
        samples.iter().all(|s| s.forward_preservation_ok),
    );
    json_bool(
        &mut output,
        "preservation_ok_all",
        samples.iter().all(|s| s.preservation_ok),
    );
    json_bool(
        &mut output,
        "inverse_ok_all",
        samples.iter().all(|s| s.inverse_ok),
    );
    json_bool(
        &mut output,
        "changed_ok_all",
        samples.iter().all(|s| s.changed_ok),
    );
    json_bool(
        &mut output,
        "source_shared_all",
        samples.iter().all(|s| s.source_shared),
    );
    let source_sharing_applicable = matches!(operation, Operation::FamilyClone | Operation::Noop);
    json_bool(
        &mut output,
        "source_sharing_applicable",
        source_sharing_applicable,
    );
    json_bool(
        &mut output,
        "source_sharing_gate",
        !source_sharing_applicable || samples.iter().all(|s| s.source_shared),
    );
    json_bool(
        &mut output,
        "allocation_balance_all",
        samples.iter().all(|s| s.allocation.balance_ok()),
    );
    json_bool(
        &mut output,
        "source_observation_stable",
        source_observation_stable(samples),
    );
    json_num(&mut output, "p50_ns", percentile(&elapsed, 50));
    json_num(&mut output, "p95_ns", percentile(&elapsed, 95));
    json_num(&mut output, "p99_ns", percentile(&elapsed, 99));
    json_num(
        &mut output,
        "requested_alloc_p50",
        percentile(&allocated, 50),
    );
    json_num(
        &mut output,
        "requested_alloc_p95",
        percentile(&allocated, 95),
    );
    json_num(&mut output, "peak_live_delta_p50", percentile(&peak, 50));
    json_num(&mut output, "peak_live_delta_p95", percentile(&peak, 95));
    output.push_str(",\"samples\":[");
    for (index, sample) in samples.iter().enumerate() {
        if index != 0 {
            output.push(',');
        }
        output.push('{');
        json_num(&mut output, "elapsed_ns", sample.elapsed_ns);
        json_num(
            &mut output,
            "requested_alloc_bytes",
            sample.allocation.requested_bytes(),
        );
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
            "peak_live_delta",
            sample.allocation.peak_live_delta,
        );
        json_num(&mut output, "live_before", sample.allocation.live_before);
        json_num(&mut output, "live_after", sample.allocation.live_after);
        json_num(&mut output, "read_calls", sample.reads.calls);
        json_num(
            &mut output,
            "read_requested_bytes",
            sample.reads.requested_bytes,
        );
        json_num(
            &mut output,
            "read_returned_bytes",
            sample.reads.returned_bytes,
        );
        json_num(&mut output, "alloc_failed", sample.allocation.failed);
        json_bool(&mut output, "alloc_invalid", sample.allocation.invalid);
        json_num(&mut output, "theme_semantic_hash", sample.theme_hash);
        json_num(&mut output, "family_semantic_hash", sample.family_hash);
        json_num(
            &mut output,
            "restored_theme_semantic_hash",
            sample.restored_theme_hash,
        );
        json_num(
            &mut output,
            "restored_family_semantic_hash",
            sample.restored_family_hash,
        );
        json_bool(&mut output, "source_shared", sample.source_shared);
        json_bool(&mut output, "semantic_ok", sample.semantic_ok);
        json_bool(
            &mut output,
            "forward_preservation_ok",
            sample.forward_preservation_ok,
        );
        json_bool(&mut output, "preservation_ok", sample.preservation_ok);
        json_bool(&mut output, "inverse_ok", sample.inverse_ok);
        json_bool(&mut output, "changed_ok", sample.changed_ok);
        json_bool(
            &mut output,
            "alloc_balance_ok",
            sample.allocation.balance_ok(),
        );
        output.push('}');
    }
    output.push_str("]}");
    output
}

fn source_observation_stable(samples: &[Sample]) -> bool {
    samples.windows(2).all(|pair| {
        pair[0].reads.calls == pair[1].reads.calls
            && pair[0].reads.requested_bytes == pair[1].reads.requested_bytes
            && pair[0].reads.returned_bytes == pair[1].reads.returned_bytes
    })
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
    write!(output, "\"{key}\":\"").expect("String cannot fail");
    for character in value.chars() {
        match character {
            '"' => output.push_str("\\\""),
            '\\' => output.push_str("\\\\"),
            '\n' => output.push_str("\\n"),
            '\r' => output.push_str("\\r"),
            '\t' => output.push_str("\\t"),
            character if character.is_control() => {
                write!(output, "\\u{:04x}", character as u32).expect("String cannot fail")
            },
            character => output.push(character),
        }
    }
    output.push('"');
}

fn json_num(output: &mut String, key: &str, value: u64) {
    json_separator(output);
    write!(output, "\"{key}\":{value}").expect("String cannot fail");
}

fn json_bool(output: &mut String, key: &str, value: bool) {
    json_separator(output);
    write!(output, "\"{key}\":{value}").expect("String cannot fail");
}

fn json_separator(output: &mut String) {
    if !matches!(output.as_bytes().last(), Some(b'{') | Some(b'[')) {
        output.push(',');
    }
}

fn main() -> Result<()> {
    let args = Args::parse()?;
    println!("{}", run(&args)?);
    Ok(())
}
