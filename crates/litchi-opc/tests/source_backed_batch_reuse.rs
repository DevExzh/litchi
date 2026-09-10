#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "focused scheduler tests deliberately fail on fixture or contract errors"
)]

//! Worker-lifetime and terminal-boundary tests for ordered source-backed reads.
//!
//! The source probe records only positional reads that overlap selected payload
//! ranges.  It is armed after package construction, so catalog-ingress reads
//! cannot make the scheduler appear to use a caller thread.  A batch contains
//! enough distinct parts for several bounded waves; the test therefore checks
//! the worker lifetime directly rather than inferring it from elapsed time.

use std::collections::HashSet;
use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::panic::{AssertUnwindSafe, catch_unwind};
use std::sync::{
    Arc, Condvar, Mutex,
    atomic::{AtomicBool, Ordering},
    mpsc,
};
use std::thread::{self, ThreadId};
use std::time::{Duration, Instant};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::{OpcError, PackURI, ReadLimits, SourceBackedPackage, SourceCacheLimits};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT_MEMBER: &str = "word/document.xml";
const WORKERS: usize = 3;
const PART_COUNT: usize = 11;
const PAYLOAD_BYTES: usize = 16 * 1024;

#[derive(Debug, Clone, Copy)]
struct PayloadRange {
    start: u64,
    end: u64,
}

struct Fixture {
    archive: Vec<u8>,
    requests: Vec<PackURI>,
    expected: Vec<Vec<u8>>,
    payload_ranges: Vec<PayloadRange>,
    first_request_range: PayloadRange,
}

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).unwrap()
}

fn payload(seed: u8) -> Vec<u8> {
    (0..PAYLOAD_BYTES)
        .map(|index| seed.wrapping_add(u8::try_from(index % 251).unwrap().wrapping_mul(17)))
        .collect()
}

fn fixture() -> Fixture {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="{DOCUMENT_MEMBER}"/></Relationships>"#
    );

    let mut names = Vec::with_capacity(PART_COUNT);
    let mut uris = Vec::with_capacity(PART_COUNT);
    let mut expected_by_index = Vec::with_capacity(PART_COUNT);
    for index in 0..PART_COUNT {
        let name = format!("custom/reuse-{index:02}.bin");
        names.push(name);
        uris.push(pack(&format!("/custom/reuse-{index:02}.bin")));
        expected_by_index.push(payload(u8::try_from(index * 23 + 7).unwrap()));
    }

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .unwrap();
    writer
        .write_stored(DOCUMENT_MEMBER, b"<document/>")
        .unwrap();
    for (name, bytes) in names.iter().zip(&expected_by_index) {
        writer.write_stored(name, bytes).unwrap();
    }
    let archive = writer.finish_to_bytes().unwrap();
    let payload_ranges = local_payload_ranges(&archive, &names);
    assert_eq!(payload_ranges.len(), PART_COUNT);

    // The order deliberately crosses wave boundaries and is not archive order.
    let order = [7, 0, 10, 3, 6, 1, 9, 4, 8, 2, 5];
    let requests = order.iter().map(|&index| uris[index].clone()).collect();
    let expected = order
        .iter()
        .map(|&index| expected_by_index[index].clone())
        .collect();

    Fixture {
        archive,
        requests,
        expected,
        first_request_range: payload_ranges[7],
        payload_ranges,
    }
}

fn read_u16(bytes: &[u8], offset: usize) -> usize {
    usize::from(u16::from_le_bytes(
        bytes[offset..offset + 2].try_into().unwrap(),
    ))
}

fn read_u32(bytes: &[u8], offset: usize) -> usize {
    usize::try_from(u32::from_le_bytes(
        bytes[offset..offset + 4].try_into().unwrap(),
    ))
    .unwrap()
}

fn local_payload_ranges(bytes: &[u8], selected_names: &[String]) -> Vec<PayloadRange> {
    let mut ranges = Vec::new();
    let mut cursor = 0usize;
    while cursor + 30 <= bytes.len() && bytes[cursor..cursor + 4] == 0x0403_4b50_u32.to_le_bytes() {
        let name_len = read_u16(bytes, cursor + 26);
        let extra_len = read_u16(bytes, cursor + 28);
        let compressed_len = read_u32(bytes, cursor + 18);
        let name_start = cursor + 30;
        let data_start = name_start + name_len + extra_len;
        let data_end = data_start + compressed_len;
        let name = std::str::from_utf8(&bytes[name_start..name_start + name_len]).unwrap();
        if selected_names.iter().any(|selected| selected == name) {
            ranges.push(PayloadRange {
                start: u64::try_from(data_start).unwrap(),
                end: u64::try_from(data_end).unwrap(),
            });
        }
        cursor = data_end;
    }
    ranges
}

#[derive(Debug, Default)]
struct ThreadProbeState {
    armed: bool,
    payload_reads: usize,
    ids: HashSet<ThreadId>,
}

struct ThreadProbe {
    inner: Arc<dyn ReadAt>,
    ranges: Vec<PayloadRange>,
    state: Mutex<ThreadProbeState>,
}

impl ThreadProbe {
    fn new(inner: Arc<dyn ReadAt>, ranges: Vec<PayloadRange>) -> Self {
        Self {
            inner,
            ranges,
            state: Mutex::new(ThreadProbeState::default()),
        }
    }

    fn arm(&self) {
        let mut state = self.state.lock().unwrap();
        *state = ThreadProbeState {
            armed: true,
            ..ThreadProbeState::default()
        };
    }

    fn disarm(&self) {
        self.state.lock().unwrap().armed = false;
    }

    fn snapshot(&self) -> (HashSet<ThreadId>, usize) {
        let state = self.state.lock().unwrap();
        (state.ids.clone(), state.payload_reads)
    }

    fn overlaps_payload(&self, offset: u64, length: usize) -> bool {
        let end = offset.saturating_add(u64::try_from(length).unwrap_or(u64::MAX));
        self.ranges
            .iter()
            .any(|range| offset < range.end && range.start < end)
    }
}

impl ReadAt for ThreadProbe {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if self.overlaps_payload(offset, output.len()) {
            let mut state = self.state.lock().unwrap();
            if state.armed {
                state.payload_reads += 1;
                state.ids.insert(thread::current().id());
            }
        }
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

struct PanicProbe {
    inner: Arc<dyn ReadAt>,
    ranges: Vec<PayloadRange>,
    armed: AtomicBool,
    panicked: AtomicBool,
}

#[derive(Debug, Default)]
struct PanicGateState {
    armed: bool,
    first_started: bool,
    release: bool,
    panicked: bool,
}

struct PanicGateProbe {
    inner: Arc<dyn ReadAt>,
    range: PayloadRange,
    state: Arc<(Mutex<PanicGateState>, Condvar)>,
}

impl PanicGateProbe {
    fn new(inner: Arc<dyn ReadAt>, range: PayloadRange) -> Self {
        Self {
            inner,
            range,
            state: Arc::new((Mutex::new(PanicGateState::default()), Condvar::new())),
        }
    }

    fn arm(&self) {
        let (lock, changed) = &*self.state;
        let mut state = lock.lock().unwrap();
        *state = PanicGateState {
            armed: true,
            ..PanicGateState::default()
        };
        changed.notify_all();
    }

    fn disarm(&self) {
        let (lock, changed) = &*self.state;
        let mut state = lock.lock().unwrap();
        state.armed = false;
        state.release = true;
        changed.notify_all();
    }

    fn wait_for_first(&self, timeout: Duration) -> bool {
        let deadline = Instant::now() + timeout;
        let (lock, changed) = &*self.state;
        let mut state = lock.lock().unwrap();
        while !state.first_started {
            let remaining = deadline.saturating_duration_since(Instant::now());
            if remaining.is_zero() {
                return false;
            }
            let (next, timed_out) = changed.wait_timeout(state, remaining).unwrap();
            state = next;
            if timed_out.timed_out() && !state.first_started {
                return false;
            }
        }
        true
    }

    fn release(&self) {
        let (lock, changed) = &*self.state;
        let mut state = lock.lock().unwrap();
        state.release = true;
        changed.notify_all();
    }

    fn overlaps_payload(&self, offset: u64, length: usize) -> bool {
        let end = offset.saturating_add(u64::try_from(length).unwrap_or(u64::MAX));
        offset < self.range.end && self.range.start < end
    }
}

impl ReadAt for PanicGateProbe {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if self.overlaps_payload(offset, output.len()) {
            let (lock, changed) = &*self.state;
            let mut state = lock.lock().unwrap();
            if state.armed && !state.panicked {
                state.first_started = true;
                changed.notify_all();
                while !state.release {
                    state = changed.wait(state).unwrap();
                }
                state.panicked = true;
                drop(state);
                panic!("intentional gated source-backed batch worker panic");
            }
        }
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

impl PanicProbe {
    fn new(inner: Arc<dyn ReadAt>, ranges: Vec<PayloadRange>) -> Self {
        Self {
            inner,
            ranges,
            armed: AtomicBool::new(false),
            panicked: AtomicBool::new(false),
        }
    }

    fn arm(&self) {
        self.panicked.store(false, Ordering::Release);
        self.armed.store(true, Ordering::Release);
    }

    fn overlaps_payload(&self, offset: u64, length: usize) -> bool {
        let end = offset.saturating_add(u64::try_from(length).unwrap_or(u64::MAX));
        self.ranges
            .iter()
            .any(|range| offset < range.end && range.start < end)
    }
}

impl ReadAt for PanicProbe {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if self.armed.load(Ordering::Acquire)
            && self.overlaps_payload(offset, output.len())
            && !self.panicked.swap(true, Ordering::AcqRel)
        {
            panic!("intentional source-backed batch worker panic");
        }
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

fn managed_package(source: Arc<dyn ReadAt>) -> (SourceBackedPackage, Budget) {
    let budget = Budget::root(
        "opc-source-backed-batch-reuse-test",
        Limits::new(
            256 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (_cancellation, token) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(WORKERS).unwrap(),
        NonZeroUsize::new(WORKERS).unwrap(),
        NonZeroU64::new(u64::try_from(WORKERS * PAYLOAD_BYTES).unwrap()).unwrap(),
        1,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), token, execution_limits);
    let cache_limits =
        SourceCacheLimits::new(PAYLOAD_BYTES * PART_COUNT * 2, PART_COUNT + 4).unwrap();
    let package =
        SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source,
            ReadLimits::default(),
            cache_limits,
            context,
        )
        .unwrap();
    (package, budget)
}

fn assert_batch_bytes(batch: &litchi_opc::PartBatch, expected: &[Vec<u8>]) {
    assert_eq!(batch.len(), expected.len());
    for (part, expected) in batch.iter().zip(expected) {
        assert_eq!(part.as_bytes(), expected.as_slice());
    }
}

#[test]
fn batch_reuses_at_most_configured_worker_threads_across_many_waves() {
    let fixture = fixture();
    let probe = Arc::new(ThreadProbe::new(
        Arc::new(OwnedSource::new(fixture.archive.clone())),
        fixture.payload_ranges.clone(),
    ));
    let (package, budget) = managed_package(probe.clone());

    probe.arm();
    let batch = package.read_parts_ordered(&fixture.requests).unwrap();
    assert_batch_bytes(&batch, &fixture.expected);
    let (ids, payload_reads) = probe.snapshot();
    assert!(
        payload_reads > 0,
        "the armed probe saw no selected payload read"
    );
    assert!(
        ids.len() <= WORKERS,
        "worker reuse exceeded the configured bound: {} > {WORKERS}",
        ids.len()
    );
    drop(batch);
    probe.disarm();
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn separate_batch_operations_do_not_reuse_worker_threads() {
    let fixture = fixture();
    let probe = Arc::new(ThreadProbe::new(
        Arc::new(OwnedSource::new(fixture.archive.clone())),
        fixture.payload_ranges.clone(),
    ));
    let source: Arc<dyn ReadAt> = probe.clone();
    let (first_package, first_budget) = managed_package(source.clone());

    probe.arm();
    let first = first_package.read_parts_ordered(&fixture.requests).unwrap();
    assert_batch_bytes(&first, &fixture.expected);
    let (first_ids, first_reads) = probe.snapshot();
    assert!(first_reads > 0);
    assert!(!first_ids.is_empty());
    assert!(first_ids.len() <= WORKERS);
    drop(first);
    probe.disarm();
    drop(first_package);
    assert_eq!(first_budget.used(Resource::Memory), 0);
    assert_eq!(first_budget.used(Resource::Objects), 0);

    // Construct the second package while the probe is disarmed, then arm only
    // its selected reads.  This keeps catalog-ingress activity out of either
    // operation's worker identity set.
    let (second_package, second_budget) = managed_package(source);
    probe.arm();
    let second = second_package
        .read_parts_ordered(&fixture.requests)
        .unwrap();
    assert_batch_bytes(&second, &fixture.expected);
    let (second_ids, second_reads) = probe.snapshot();
    assert!(second_reads > 0);
    assert!(!second_ids.is_empty());
    assert!(second_ids.len() <= WORKERS);
    assert!(
        first_ids.is_disjoint(&second_ids),
        "worker identities leaked across separate batch operations"
    );
    drop(second);
    probe.disarm();
    drop(second_package);
    assert_eq!(second_budget.used(Resource::Memory), 0);
    assert_eq!(second_budget.used(Resource::Objects), 0);
}

#[cfg(panic = "unwind")]
#[test]
fn provider_panic_is_a_typed_terminal_batch_error() {
    let fixture = fixture();
    let probe = Arc::new(PanicProbe::new(
        Arc::new(OwnedSource::new(fixture.archive)),
        fixture.payload_ranges,
    ));
    let (package, budget) = managed_package(probe.clone());
    probe.arm();

    let result = catch_unwind(AssertUnwindSafe(|| {
        package.read_parts_ordered(&fixture.requests)
    }));
    let result = result.expect("provider panic must be contained by the worker boundary");
    assert!(matches!(
        result,
        Err(OpcError::SourceBackedBatchWorkerPanic { .. })
    ));
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    // The one-shot provider panic is now consumed.  A subsequent ordinary
    // read must be able to load the selected Part, proving that a failed
    // worker did not strand its cache flight or reservation.
    let recovered = package.part(&fixture.requests[0]).unwrap().data().unwrap();
    assert_eq!(recovered.as_bytes(), fixture.expected[0].as_slice());
    probe.armed.store(false, Ordering::Release);
    drop(package);
    assert!(budget.used(Resource::Memory) > 0);
    drop(recovered);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[cfg(panic = "unwind")]
#[test]
fn duplicate_waiter_recovers_after_a_gated_provider_panic() {
    let fixture = fixture();
    let probe = Arc::new(PanicGateProbe::new(
        Arc::new(OwnedSource::new(fixture.archive)),
        fixture.first_request_range,
    ));
    let (package, budget) = managed_package(probe.clone());
    let package = Arc::new(package);
    let requests = vec![
        fixture.requests[0].clone(),
        fixture.requests[0].clone(),
        fixture.requests[1].clone(),
        fixture.requests[2].clone(),
    ];
    probe.arm();
    let (sender, receiver) = mpsc::channel();
    let operation_package = Arc::clone(&package);
    let operation = thread::spawn(move || {
        let result = operation_package.read_parts_ordered(&requests);
        sender.send(result).unwrap();
    });

    let first_started = probe.wait_for_first(Duration::from_secs(2));
    let deadline = Instant::now() + Duration::from_secs(2);
    let mut waiter_joined = false;
    while Instant::now() < deadline {
        if package.cache_diagnostics().waiter_joins > 0 {
            waiter_joined = true;
            break;
        }
        thread::sleep(Duration::from_millis(1));
    }
    // Always release the provider before asserting either observation so a
    // failed setup cannot strand the operation thread behind the test gate.
    probe.release();
    assert!(first_started, "the duplicate loader never reached the gate");
    assert!(
        waiter_joined,
        "the duplicate request did not join the in-flight loader"
    );

    let result = receiver
        .recv_timeout(Duration::from_secs(5))
        .expect("a failed duplicate loader must not strand its waiter");
    operation
        .join()
        .expect("the operation thread must terminate after the provider panic");
    assert!(matches!(
        result,
        Err(OpcError::SourceBackedBatchWorkerPanic { .. })
    ));
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);

    let recovered = package.part(&fixture.requests[0]).unwrap().data().unwrap();
    assert_eq!(recovered.as_bytes(), fixture.expected[0].as_slice());
    probe.disarm();
    drop(package);
    assert!(budget.used(Resource::Memory) > 0);
    drop(recovered);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}
