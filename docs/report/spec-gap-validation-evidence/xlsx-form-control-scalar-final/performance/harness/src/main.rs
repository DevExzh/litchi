#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "The harness keeps measurement, allocator, source, and correctness helpers together for relocation."
)]
#![allow(
    clippy::too_many_lines,
    reason = "Each lifecycle lane is kept explicit so its timed boundary is reviewable."
)]

use litchi_core::ReadAt;
use litchi_opc::{OpcPackage, PackageWriter};
use litchi_xlsx::form_control::{
    Checked, ControlSelector, FormControlCollection, FormControlCommit, FormControlPatch,
    ScalarField, ScalarValue, SourceBackedFormControlEditor,
};
use litchi_xlsx::{SourceBackedWorkbook, Workbook};
use std::alloc::{GlobalAlloc, Layout, System};
use std::hint::black_box;
use std::io;
use std::path::PathBuf;
use std::sync::{
    Arc,
    atomic::{AtomicBool, AtomicI64, AtomicU64, Ordering},
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
    live_bytes: AtomicI64,
    phase_baseline_live_bytes: AtomicI64,
    phase_peak_live_bytes: AtomicI64,
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
            live_bytes: AtomicI64::new(0),
            phase_baseline_live_bytes: AtomicI64::new(0),
            phase_peak_live_bytes: AtomicI64::new(0),
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
        let baseline = self.phase_baseline_live_bytes.load(Ordering::SeqCst);
        let live = self.live_bytes.load(Ordering::SeqCst);
        let peak = self.phase_peak_live_bytes.load(Ordering::SeqCst);
        AllocationSnapshot {
            alloc_calls: self.alloc_calls.load(Ordering::Relaxed),
            alloc_bytes: self.alloc_bytes.load(Ordering::Relaxed),
            dealloc_calls: self.dealloc_calls.load(Ordering::Relaxed),
            dealloc_bytes: self.dealloc_bytes.load(Ordering::Relaxed),
            realloc_calls: self.realloc_calls.load(Ordering::Relaxed),
            realloc_bytes: self.realloc_bytes.load(Ordering::Relaxed),
            requested_event_bytes: self.requested_event_bytes.load(Ordering::Relaxed),
            live_bytes_delta: i128::from(live) - i128::from(baseline),
            peak_live_bytes: u64::try_from((peak - baseline).max(0)).unwrap_or(u64::MAX),
        }
    }

    fn record_alloc(&self, size: usize) {
        self.alloc_calls.fetch_add(1, Ordering::Relaxed);
        self.alloc_bytes.fetch_add(size as u64, Ordering::Relaxed);
        self.requested_event_bytes
            .fetch_add(size as u64, Ordering::Relaxed);
    }

    fn record_dealloc(&self, size: usize) {
        self.dealloc_calls.fetch_add(1, Ordering::Relaxed);
        self.dealloc_bytes.fetch_add(size as u64, Ordering::Relaxed);
    }

    fn record_realloc(&self, size: usize) {
        self.realloc_calls.fetch_add(1, Ordering::Relaxed);
        self.realloc_bytes.fetch_add(size as u64, Ordering::Relaxed);
        self.requested_event_bytes
            .fetch_add(size as u64, Ordering::Relaxed);
    }

    fn record_live(&self, delta: i64) {
        let live = self
            .live_bytes
            .fetch_add(delta, Ordering::SeqCst)
            .saturating_add(delta);
        if self.enabled.load(Ordering::Relaxed) {
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
            TRACKER.record_live(layout.size() as i64);
            if TRACKER.enabled.load(Ordering::Relaxed) {
                TRACKER.record_alloc(layout.size());
            }
        }
        pointer
    }

    unsafe fn alloc_zeroed(&self, layout: Layout) -> *mut u8 {
        let pointer = unsafe { System.alloc_zeroed(layout) };
        if !pointer.is_null() {
            TRACKER.record_live(layout.size() as i64);
            if TRACKER.enabled.load(Ordering::Relaxed) {
                TRACKER.record_alloc(layout.size());
            }
        }
        pointer
    }

    unsafe fn dealloc(&self, pointer: *mut u8, layout: Layout) {
        unsafe { System.dealloc(pointer, layout) };
        TRACKER.record_live(-(layout.size() as i64));
        if TRACKER.enabled.load(Ordering::Relaxed) {
            TRACKER.record_dealloc(layout.size());
        }
    }

    unsafe fn realloc(&self, pointer: *mut u8, layout: Layout, new_size: usize) -> *mut u8 {
        let replacement = unsafe { System.realloc(pointer, layout, new_size) };
        if !replacement.is_null() {
            TRACKER.record_live(new_size as i64 - layout.size() as i64);
            if TRACKER.enabled.load(Ordering::Relaxed) {
                TRACKER.record_realloc(new_size);
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

    fn record_read(&self, requested: usize, actual: usize) {
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
            self.counters.record_read(output.len(), 0);
            return Ok(0);
        }
        let end = offset.saturating_add(output.len()).min(self.bytes.len());
        let count = end - offset;
        output[..count].copy_from_slice(&self.bytes[offset..end]);
        self.counters.record_read(output.len(), count);
        Ok(count)
    }

    fn version(&self) -> io::Result<litchi_core::SourceVersion> {
        self.counters.version_calls.fetch_add(1, Ordering::Relaxed);
        Ok(litchi_core::SourceVersion::new(0x5343_414c_4152_u64, 0))
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
struct MemorySnapshot {
    rss_bytes: Option<u64>,
    hwm_bytes: Option<u64>,
}

#[derive(Clone, Copy)]
struct Receipt {
    elapsed_ns: u128,
    memory_before: MemorySnapshot,
    memory_after: MemorySnapshot,
    allocation: AllocationSnapshot,
    source: SourceSnapshot,
}

#[derive(Clone, Debug)]
struct Outcome {
    checksum: u64,
    output_bytes: usize,
    changed: bool,
}

#[derive(Clone, Debug)]
struct ScalarCase {
    field: ScalarField,
    field_name: &'static str,
    noop: ScalarValue,
    changed: ScalarValue,
}

impl Receipt {
    fn emit(
        self,
        fixture: &str,
        lane: &str,
        iteration: usize,
        case: &ScalarCase,
        outcome: &Outcome,
    ) {
        println!(
            "{{\"record\":\"sample\",\"fixture\":{},\"lane\":{},\"iteration\":{},\"field\":{},\"elapsed_ns\":{},\"rss_before_bytes\":{},\"rss_after_bytes\":{},\"hwm_before_bytes\":{},\"hwm_after_bytes\":{},\"alloc_calls\":{},\"alloc_bytes_requested\":{},\"dealloc_calls\":{},\"dealloc_bytes_requested\":{},\"realloc_calls\":{},\"realloc_bytes_requested\":{},\"requested_event_bytes\":{},\"live_bytes_delta\":{},\"peak_live_bytes\":{},\"source_read_calls\":{},\"source_read_bytes\":{},\"source_request_bytes\":{},\"source_max_request_bytes\":{},\"source_len_calls\":{},\"source_version_calls\":{},\"output_bytes\":{},\"output_checksum\":{},\"changed\":{}}}",
            json_string(fixture),
            json_string(lane),
            iteration,
            json_string(case.field_name),
            self.elapsed_ns,
            optional_u64(self.memory_before.rss_bytes),
            optional_u64(self.memory_after.rss_bytes),
            optional_u64(self.memory_before.hwm_bytes),
            optional_u64(self.memory_after.hwm_bytes),
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
            outcome.output_bytes,
            outcome.checksum,
            outcome.changed,
        );
    }
}

fn optional_u64(value: Option<u64>) -> String {
    value.map_or_else(|| "null".to_owned(), |number| number.to_string())
}

fn json_string(value: &str) -> String {
    let mut output = String::with_capacity(value.len() + 2);
    output.push('"');
    for byte in value.bytes() {
        match byte {
            b'"' => output.push_str("\\\""),
            b'\\' => output.push_str("\\\\"),
            b'\n' => output.push_str("\\n"),
            b'\r' => output.push_str("\\r"),
            b'\t' => output.push_str("\\t"),
            0..=0x1f => {
                use std::fmt::Write as _;
                let _ = write!(output, "\\u{byte:04x}");
            },
            _ => output.push(byte as char),
        }
    }
    output.push('"');
    output
}

fn rss_snapshot() -> MemorySnapshot {
    #[cfg(target_os = "linux")]
    {
        let Ok(status) = std::fs::read_to_string("/proc/self/status") else {
            return MemorySnapshot {
                rss_bytes: None,
                hwm_bytes: None,
            };
        };
        let value = |name: &str| {
            status
                .lines()
                .find(|line| line.starts_with(name))
                .and_then(|line| line.split_whitespace().nth(1))
                .and_then(|value| value.parse::<u64>().ok())
                .map(|kib| kib.saturating_mul(1024))
        };
        MemorySnapshot {
            rss_bytes: value("VmRSS:"),
            hwm_bytes: value("VmHWM:"),
        }
    }
    #[cfg(not(target_os = "linux"))]
    {
        MemorySnapshot {
            rss_bytes: None,
            hwm_bytes: None,
        }
    }
}

fn measure<T>(source: Option<&CountingSource>, operation: impl FnOnce() -> T) -> (Receipt, T) {
    if let Some(source) = source {
        source.counters.reset();
    }
    let memory_before = rss_snapshot();
    TRACKER.reset_and_enable();
    let started = Instant::now();
    let value = operation();
    let elapsed_ns = started.elapsed().as_nanos();
    let allocation = TRACKER.disable();
    let memory_after = rss_snapshot();
    let source = source.map_or(
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
            memory_before,
            memory_after,
            allocation,
            source,
        },
        value,
    )
}

fn fnv1a(bytes: &[u8]) -> u64 {
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    for byte in bytes {
        hash ^= u64::from(*byte);
        hash = hash.wrapping_mul(0x1000_0000_01b3);
    }
    hash
}

fn collection_checksum(collection: &FormControlCollection) -> u64 {
    let mut bytes = Vec::new();
    for control in collection.iter() {
        bytes.extend_from_slice(&(control.position() as u64).to_le_bytes());
        if let Some(name) = control.name() {
            bytes.extend_from_slice(name.as_bytes());
        }
        if let Some(source) = control.properties().source_bytes() {
            bytes.extend_from_slice(source);
        }
    }
    fnv1a(&bytes)
}

fn collection_from_eager(bytes: &[u8]) -> FormControlCollection {
    Workbook::from_bytes(bytes.to_vec())
        .expect("eager workbook open")
        .sheet(0)
        .expect("eager sheet lookup")
        .expect("eager worksheet")
        .form_controls()
        .expect("eager form-control read")
}

fn source_editor(bytes: &[u8]) -> (Arc<CountingSource>, SourceBackedFormControlEditor) {
    let source_impl = Arc::new(CountingSource::new(bytes.to_vec()));
    let source: Arc<dyn ReadAt> = source_impl.clone();
    let editor = SourceBackedFormControlEditor::from_read_at(source).expect("source editor open");
    (source_impl, editor)
}

fn collection_from_source(bytes: &[u8]) -> FormControlCollection {
    let source = Arc::new(CountingSource::new(bytes.to_vec()));
    let source_trait: Arc<dyn ReadAt> = source;
    SourceBackedWorkbook::from_read_at(source_trait)
        .expect("source workbook open")
        .sheet(0)
        .expect("source sheet lookup")
        .expect("source worksheet")
        .form_controls()
        .expect("source form-control read")
}

fn toggle_scalar(value: &ScalarValue) -> Option<ScalarValue> {
    match value {
        ScalarValue::Boolean(value) => Some(ScalarValue::Boolean(!value)),
        ScalarValue::Checked(Checked::Checked) => Some(ScalarValue::Checked(Checked::Unchecked)),
        ScalarValue::Checked(Checked::Unchecked | Checked::Mixed) => {
            Some(ScalarValue::Checked(Checked::Checked))
        },
        ScalarValue::Unsigned(value) => Some(ScalarValue::Unsigned(value.saturating_add(1))),
        _ => None,
    }
}

fn derive_case(
    collection: &FormControlCollection,
    requested_field: Option<ScalarField>,
) -> ScalarCase {
    let control = collection
        .get(ControlSelector::position(0))
        .expect("first control selector")
        .expect("non-empty control collection");
    let fields = [
        (ScalarField::Checked, "Checked"),
        (ScalarField::LockText, "LockText"),
        (ScalarField::NoThreeD, "NoThreeD"),
        (ScalarField::JustLastX, "JustLastX"),
        (ScalarField::NoThreeD2, "NoThreeD2"),
        (ScalarField::Colored, "Colored"),
    ];
    let candidates = requested_field
        .map(|field| {
            fields
                .iter()
                .copied()
                .filter(move |(candidate, _)| *candidate == field)
                .collect::<Vec<_>>()
        })
        .unwrap_or_else(|| fields.to_vec());
    candidates
        .into_iter()
        .find_map(|(field, field_name)| {
            let noop = control.properties().scalar(field)?;
            let changed = toggle_scalar(&noop)?;
            Some(ScalarCase {
                field,
                field_name,
                noop,
                changed,
            })
        })
        .unwrap_or_else(|| panic!("fixture has no supported authored scalar"))
}

fn parse_field(value: Option<String>) -> Option<ScalarField> {
    value.map(|value| match value.as_str() {
        "Checked" => ScalarField::Checked,
        "LockText" => ScalarField::LockText,
        "NoThreeD" => ScalarField::NoThreeD,
        "JustLastX" => ScalarField::JustLastX,
        "NoThreeD2" => ScalarField::NoThreeD2,
        "Colored" => ScalarField::Colored,
        _ => panic!("invalid scalar field: {value}"),
    })
}

fn value_label(value: &ScalarValue) -> String {
    format!("{value:?}")
}

fn reopened_outcome(bytes: &[u8], case: &ScalarCase, expected: &ScalarValue) -> Outcome {
    let workbook = Workbook::from_bytes(bytes.to_vec()).expect("reopen saved workbook");
    let collection = workbook
        .sheet(0)
        .expect("reopened sheet lookup")
        .expect("reopened worksheet")
        .form_controls()
        .expect("reopened form-control read");
    let control = collection
        .get(ControlSelector::position(0))
        .expect("reopened control selector")
        .expect("reopened control");
    assert_eq!(
        control.properties().scalar(case.field),
        Some(expected.clone()),
        "typed scalar readback disagreed with the requested value"
    );
    Outcome {
        checksum: collection_checksum(&collection),
        output_bytes: bytes.len(),
        changed: *expected != case.noop,
    }
}

fn source_publish(bytes: &[u8], case: &ScalarCase, value: &ScalarValue) -> Vec<u8> {
    let (_source_impl, editor) = source_editor(bytes);
    let mut edit = editor.edit("Sheet1").expect("source worksheet selection");
    edit.set_scalar(
        ControlSelector::position(0),
        case.field,
        Some(value.clone()),
    )
    .expect("stage source scalar");
    let commit = edit.commit().expect("source scalar commit");
    let mut output = Vec::new();
    editor
        .publish_commit_to_stream(&mut output, &commit)
        .expect("source scalar publication");
    output
}

fn eager_publish(bytes: &[u8], case: &ScalarCase, value: &ScalarValue) -> Vec<u8> {
    let workbook = Workbook::from_bytes(bytes.to_vec()).expect("eager workbook open");
    let mut edit = workbook.edit().expect("eager workbook edit");
    {
        let mut sheet = edit
            .sheet("Sheet1")
            .expect("eager worksheet selection")
            .expect("eager worksheet");
        sheet
            .set_form_control_scalar(
                ControlSelector::position(0),
                case.field,
                Some(value.clone()),
            )
            .expect("stage eager scalar");
    }
    edit.commit()
        .expect("eager scalar commit")
        .into_workbook()
        .to_plain_bytes()
        .expect("eager scalar serialization")
}

fn source_commit_for(bytes: &[u8], case: &ScalarCase, value: &ScalarValue) -> FormControlCommit {
    let (_source_impl, editor) = source_editor(bytes);
    let mut edit = editor.edit("Sheet1").expect("source worksheet selection");
    edit.set_scalar(
        ControlSelector::position(0),
        case.field,
        Some(value.clone()),
    )
    .expect("stage source scalar");
    edit.commit().expect("source scalar commit")
}

fn source_patch_for(
    bytes: &[u8],
    case: &ScalarCase,
    value: &ScalarValue,
) -> (FormControlPatch, Vec<u8>) {
    let commit = source_commit_for(bytes, case, value);
    let canonical =
        PackageWriter::to_bytes(&OpcPackage::from_bytes(bytes).expect("canonical source package"))
            .expect("canonical source serialization");
    (commit.patch().clone(), canonical)
}

fn sample_eager_read(bytes: &[u8], _case: &ScalarCase) -> (Receipt, Outcome) {
    let workbook = Workbook::from_bytes(bytes.to_vec()).expect("eager workbook open");
    measure(None, || {
        let collection = workbook
            .sheet(0)
            .expect("eager sheet lookup")
            .expect("eager worksheet")
            .form_controls()
            .expect("eager form-control read");
        Outcome {
            checksum: collection_checksum(&collection),
            output_bytes: 0,
            changed: false,
        }
    })
}

fn sample_source_read(bytes: &[u8], _case: &ScalarCase) -> (Receipt, Outcome) {
    let (source_impl, editor) = source_editor(bytes);
    measure(Some(&source_impl), || {
        let snapshot = editor
            .snapshot("Sheet1")
            .expect("source form-control snapshot");
        Outcome {
            checksum: collection_checksum(snapshot.form_controls()),
            output_bytes: 0,
            changed: false,
        }
    })
}

fn sample_eager_save_reopen(
    bytes: &[u8],
    case: &ScalarCase,
    value: &ScalarValue,
) -> (Receipt, Outcome) {
    let workbook = Workbook::from_bytes(bytes.to_vec()).expect("eager workbook open");
    measure(None, || {
        let mut edit = workbook.edit().expect("eager workbook edit");
        {
            let mut sheet = edit
                .sheet("Sheet1")
                .expect("eager worksheet selection")
                .expect("eager worksheet");
            sheet
                .set_form_control_scalar(
                    ControlSelector::position(0),
                    case.field,
                    Some(value.clone()),
                )
                .expect("stage eager scalar");
        }
        let output = edit
            .commit()
            .expect("eager scalar commit")
            .into_workbook()
            .to_plain_bytes()
            .expect("eager scalar serialization");
        let outcome = reopened_outcome(&output, case, value);
        black_box(&output);
        outcome
    })
}

fn sample_source_save_reopen(
    bytes: &[u8],
    case: &ScalarCase,
    value: &ScalarValue,
) -> (Receipt, Outcome) {
    let (source_impl, editor) = source_editor(bytes);
    measure(Some(&source_impl), || {
        let mut edit = editor.edit("Sheet1").expect("source worksheet selection");
        edit.set_scalar(
            ControlSelector::position(0),
            case.field,
            Some(value.clone()),
        )
        .expect("stage source scalar");
        let commit = edit.commit().expect("source scalar commit");
        let mut output = Vec::new();
        editor
            .publish_commit_to_stream(&mut output, &commit)
            .expect("source scalar publication");
        let outcome = reopened_outcome(&output, case, value);
        black_box(&output);
        outcome
    })
}

fn sample_source_forward_apply(bytes: &[u8], case: &ScalarCase) -> (Receipt, Outcome) {
    let (patch, _canonical) = source_patch_for(bytes, case, &case.changed);
    let mut package = OpcPackage::from_bytes(bytes).expect("source package for forward patch");
    measure(None, || {
        patch
            .apply(&mut package)
            .expect("forward source patch application");
        Outcome {
            checksum: collection_checksum(patch.after().form_controls()),
            output_bytes: 0,
            changed: true,
        }
    })
}

fn sample_source_inverse_apply(bytes: &[u8], case: &ScalarCase) -> (Receipt, Outcome) {
    let (patch, _canonical) = source_patch_for(bytes, case, &case.changed);
    let mut package = OpcPackage::from_bytes(bytes).expect("source package for inverse patch");
    patch
        .apply(&mut package)
        .expect("prepare forward source patch");
    let inverse = patch.inverse();
    measure(None, || {
        inverse
            .apply(&mut package)
            .expect("inverse source patch application");
        Outcome {
            checksum: collection_checksum(inverse.after().form_controls()),
            output_bytes: 0,
            changed: true,
        }
    })
}

fn correctness(
    bytes: &[u8],
    expected_controls: Option<usize>,
    requested_field: Option<ScalarField>,
) -> ScalarCase {
    let eager = collection_from_eager(bytes);
    let source = collection_from_source(bytes);
    assert!(!eager.is_empty(), "fixture has no effective form controls");
    assert_eq!(
        eager.len(),
        source.len(),
        "eager/source control count differs"
    );
    if let Some(expected) = expected_controls {
        assert_eq!(eager.len(), expected, "fixture control count changed");
    }
    assert_eq!(eager.profile(), source.profile(), "owner profiles differ");
    let case = derive_case(&eager, requested_field);
    assert_eq!(derive_case(&source, Some(case.field)).field, case.field);

    let source_noop = source_publish(bytes, &case, &case.noop);
    assert_eq!(
        source_noop, bytes,
        "source exact scalar no-op changed bytes"
    );
    let eager_noop = eager_publish(bytes, &case, &case.noop);
    assert_eq!(eager_noop, bytes, "eager exact scalar no-op changed bytes");

    let source_changed = source_publish(bytes, &case, &case.changed);
    assert_ne!(
        source_changed, bytes,
        "source scalar edit produced no byte change"
    );
    let _ = reopened_outcome(&source_changed, &case, &case.changed);
    let eager_changed = eager_publish(bytes, &case, &case.changed);
    let _ = reopened_outcome(&eager_changed, &case, &case.changed);

    let (patch, source_canonical) = source_patch_for(bytes, &case, &case.changed);
    let mut package = OpcPackage::from_bytes(bytes).expect("source package for inverse check");
    patch
        .apply(&mut package)
        .expect("correctness forward patch");
    let forward_canonical = PackageWriter::to_bytes(&package).expect("forward canonical bytes");
    let mut forward = OpcPackage::from_bytes(&forward_canonical).expect("reopen forward package");
    patch
        .inverse()
        .apply(&mut forward)
        .expect("correctness inverse patch");
    let restored_canonical = PackageWriter::to_bytes(&forward).expect("inverse canonical bytes");
    assert_eq!(
        restored_canonical, source_canonical,
        "inverse did not restore source"
    );

    println!(
        "{{\"record\":\"correctness\",\"fixture_bytes\":{},\"controls\":{},\"field\":{},\"noop_value\":{},\"changed_value\":{},\"source_noop_exact\":true,\"eager_noop_exact\":true,\"source_changed_reopen\":true,\"eager_changed_reopen\":true,\"inverse_canonical_exact\":true}}",
        bytes.len(),
        eager.len(),
        json_string(case.field_name),
        json_string(&value_label(&case.noop)),
        json_string(&value_label(&case.changed)),
    );
    case
}

fn parse_arg(args: &[String], name: &str) -> String {
    args.windows(2)
        .find(|pair| pair[0] == name)
        .map(|pair| pair[1].clone())
        .unwrap_or_else(|| panic!("missing required argument {name}"))
}

fn parse_optional(args: &[String], name: &str, default: usize) -> usize {
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
    let label = parse_arg(&args, "--label");
    let expected = args
        .windows(2)
        .find(|pair| pair[0] == "--expected-controls")
        .map(|pair| {
            pair[1]
                .parse::<usize>()
                .unwrap_or_else(|_| panic!("invalid value for --expected-controls"))
        });
    let requested_field = parse_field(
        args.windows(2)
            .find(|pair| pair[0] == "--field")
            .map(|pair| pair[1].clone()),
    );
    let warmups = parse_optional(&args, "--warmups", 3);
    let iterations = parse_optional(&args, "--iterations", 15);
    let bytes = std::fs::read(&fixture_path).expect("read fixture");
    let case = correctness(&bytes, expected, requested_field);

    type Sampler = fn(&[u8], &ScalarCase) -> (Receipt, Outcome);
    let lanes: [(&str, Sampler); 8] = [
        ("eager_read", sample_eager_read),
        ("source_read", sample_source_read),
        ("eager_noop_save_reopen", |bytes, case| {
            sample_eager_save_reopen(bytes, case, &case.noop)
        }),
        ("source_noop_save_reopen", |bytes, case| {
            sample_source_save_reopen(bytes, case, &case.noop)
        }),
        ("eager_scalar_save_reopen", |bytes, case| {
            sample_eager_save_reopen(bytes, case, &case.changed)
        }),
        ("source_scalar_save_reopen", |bytes, case| {
            sample_source_save_reopen(bytes, case, &case.changed)
        }),
        ("source_forward_apply", sample_source_forward_apply),
        ("source_inverse_apply", sample_source_inverse_apply),
    ];

    for (lane, sampler) in lanes {
        for _ in 0..warmups {
            let _ = sampler(&bytes, &case);
        }
        for iteration in 0..iterations {
            let (receipt, outcome) = sampler(&bytes, &case);
            receipt.emit(&label, lane, iteration, &case, &outcome);
        }
    }
    println!(
        "{{\"record\":\"complete\",\"fixture\":{},\"warmups\":{},\"iterations\":{},\"lanes\":8}}",
        json_string(&label),
        warmups,
        iterations,
    );
}
