#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::too_many_lines,
    reason = "focused integration assertions deliberately fail on fixture or contract errors"
)]

//! Integration coverage for the bounded source-backed multi-Part read.
//!
//! The provider wrappers in this file are deliberately positional and
//! observable.  The overlap assertions use a release gate and an active-call
//! counter; they do not infer parallelism from elapsed time.

use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc, Condvar, Mutex,
    atomic::{AtomicU64, Ordering},
};
use std::thread;
use std::time::{Duration, Instant};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::{
    OpcError, PackURI, ReadLimits, ReadResource, SourceBackedPackage, SourceCacheLimits,
};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT_MEMBER: &str = "word/document.xml";
const DOCUMENT_URI: &str = "/word/document.xml";
const ALPHA_MEMBER: &str = "custom/alpha.bin";
const ALPHA_URI: &str = "/custom/alpha.bin";
const BETA_MEMBER: &str = "custom/beta.bin";
const BETA_URI: &str = "/custom/beta.bin";
const GAMMA_MEMBER: &str = "custom/gamma.bin";
const GAMMA_URI: &str = "/custom/gamma.bin";
const PART_BYTES: usize = 8192;

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).unwrap()
}

fn payload(seed: u8) -> Vec<u8> {
    (0..PART_BYTES)
        .map(|index| {
            let index = u8::try_from(index % 251).unwrap();
            seed.wrapping_add(index.wrapping_mul(17))
        })
        .collect()
}

fn archive_bytes(corrupt_member: Option<&str>) -> (Vec<u8>, Vec<(PackURI, Vec<u8>)>) {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="{DOCUMENT_MEMBER}"/></Relationships>"#
    );
    let parts = vec![
        (pack(DOCUMENT_URI), DOCUMENT_MEMBER, payload(0x11)),
        (pack(ALPHA_URI), ALPHA_MEMBER, payload(0x31)),
        (pack(BETA_URI), BETA_MEMBER, payload(0x51)),
        (pack(GAMMA_URI), GAMMA_MEMBER, payload(0x71)),
    ];

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .unwrap();
    for (_, member, bytes) in &parts {
        writer.write_stored(member, bytes).unwrap();
    }
    let mut archive = writer.finish_to_bytes().unwrap();
    if let Some(member) = corrupt_member {
        corrupt_member_crc(&mut archive, member);
    }

    let expected = parts
        .into_iter()
        .map(|(uri, _, bytes)| (uri, bytes))
        .collect();
    (archive, expected)
}

fn deflated_payload(seed: u8, length: usize, repetitive: bool) -> Vec<u8> {
    (0..length)
        .map(|index| {
            if repetitive {
                seed.wrapping_add(u8::try_from((index / 4096) % 13).unwrap())
            } else {
                seed.wrapping_add(u8::try_from(index % 251).unwrap().wrapping_mul(29))
            }
        })
        .collect()
}

fn deflated_archive_bytes() -> (Vec<u8>, Vec<(PackURI, Vec<u8>)>) {
    const DEFLATED_DOCUMENT_MEMBER: &str = "word/document.xml";
    const DEFLATED_DOCUMENT_URI: &str = "/word/document.xml";
    const A_MEMBER: &str = "custom/deflated-a.bin";
    const A_URI: &str = "/custom/deflated-a.bin";
    const B_MEMBER: &str = "custom/deflated-b.bin";
    const B_URI: &str = "/custom/deflated-b.bin";
    const C_MEMBER: &str = "custom/deflated-c.bin";
    const C_URI: &str = "/custom/deflated-c.bin";
    const D_MEMBER: &str = "custom/deflated-d.bin";
    const D_URI: &str = "/custom/deflated-d.bin";

    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="{DEFLATED_DOCUMENT_MEMBER}"/></Relationships>"#
    );
    let parts = vec![
        (
            pack(DEFLATED_DOCUMENT_URI),
            DEFLATED_DOCUMENT_MEMBER,
            deflated_payload(0x11, 64 * 1024, true),
        ),
        (
            pack(A_URI),
            A_MEMBER,
            deflated_payload(0x31, 96 * 1024, true),
        ),
        (
            pack(B_URI),
            B_MEMBER,
            deflated_payload(0x51, 128 * 1024, false),
        ),
        (
            pack(C_URI),
            C_MEMBER,
            deflated_payload(0x71, 192 * 1024, true),
        ),
        (
            pack(D_URI),
            D_MEMBER,
            deflated_payload(0x91, 256 * 1024, false),
        ),
    ];

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .unwrap();
    for (_, member, bytes) in &parts {
        writer.write_deflated_sized(member, bytes).unwrap();
    }
    let archive = writer.finish_to_bytes().unwrap();
    let expected = parts
        .into_iter()
        .map(|(uri, _, bytes)| (uri, bytes))
        .collect();
    (archive, expected)
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

fn corrupt_member_crc(bytes: &mut [u8], wanted: &str) {
    let signature = 0x0605_4b50_u32.to_le_bytes();
    let eocd = bytes
        .windows(signature.len())
        .rposition(|window| window == signature)
        .unwrap();
    let count = read_u16(bytes, eocd + 10);
    let mut cursor = read_u32(bytes, eocd + 16);
    for _ in 0..count {
        assert_eq!(&bytes[cursor..cursor + 4], &0x0201_4b50_u32.to_le_bytes());
        let name_len = read_u16(bytes, cursor + 28);
        let extra_len = read_u16(bytes, cursor + 30);
        let comment_len = read_u16(bytes, cursor + 32);
        let name_start = cursor + 46;
        let local_offset = read_u32(bytes, cursor + 42);
        if &bytes[name_start..name_start + name_len] == wanted.as_bytes() {
            let central_crc =
                u32::from_le_bytes(bytes[cursor + 16..cursor + 20].try_into().unwrap()) ^ 1;
            bytes[cursor + 16..cursor + 20].copy_from_slice(&central_crc.to_le_bytes());
            let local_crc = u32::from_le_bytes(
                bytes[local_offset + 14..local_offset + 18]
                    .try_into()
                    .unwrap(),
            ) ^ 1;
            bytes[local_offset + 14..local_offset + 18].copy_from_slice(&local_crc.to_le_bytes());
            return;
        }
        cursor += 46 + name_len + extra_len + comment_len;
    }
    panic!("missing ZIP member {wanted}");
}

#[derive(Debug, Clone, Copy)]
struct MemberRange {
    name: &'static str,
    start: u64,
    end: u64,
}

fn stored_member_ranges(bytes: &[u8]) -> Vec<MemberRange> {
    let mut ranges = Vec::new();
    let mut cursor = 0usize;
    while cursor + 30 <= bytes.len() && bytes[cursor..cursor + 4] == 0x0403_4b50_u32.to_le_bytes() {
        let name_len = read_u16(bytes, cursor + 26);
        let extra_len = read_u16(bytes, cursor + 28);
        let compressed_len = read_u32(bytes, cursor + 18);
        let name_start = cursor + 30;
        let data_start = name_start + name_len + extra_len;
        let data_end = data_start + compressed_len;
        let name = match &bytes[name_start..name_start + name_len] {
            value if value == ALPHA_MEMBER.as_bytes() => Some(ALPHA_MEMBER),
            value if value == BETA_MEMBER.as_bytes() => Some(BETA_MEMBER),
            value if value == GAMMA_MEMBER.as_bytes() => Some(GAMMA_MEMBER),
            value if value == DOCUMENT_MEMBER.as_bytes() => Some(DOCUMENT_MEMBER),
            _ => None,
        };
        if let Some(name) = name {
            ranges.push(MemberRange {
                name,
                start: u64::try_from(data_start).unwrap(),
                end: u64::try_from(data_end).unwrap(),
            });
        }
        cursor = data_end;
    }
    ranges
}

#[derive(Debug, Clone, Copy, Default)]
struct MemberWaveSnapshot {
    started_members: usize,
    max_active_members: usize,
    max_active_declared_bytes: u64,
}

#[derive(Debug, Default)]
struct MemberWaveState {
    armed: bool,
    release_generation: usize,
    started_members: usize,
    active_members: Vec<usize>,
    max_active_members: usize,
    max_active_declared_bytes: u64,
}

struct MemberWaveProbe {
    inner: Arc<dyn ReadAt>,
    ranges: Vec<MemberRange>,
    state: Arc<(Mutex<MemberWaveState>, Condvar)>,
}

impl MemberWaveProbe {
    fn new(inner: Arc<dyn ReadAt>, ranges: Vec<MemberRange>) -> Self {
        Self {
            inner,
            ranges,
            state: Arc::new((Mutex::new(MemberWaveState::default()), Condvar::new())),
        }
    }

    fn arm(&self) {
        let (lock, changed) = &*self.state;
        let mut state = lock.lock().unwrap();
        *state = MemberWaveState {
            armed: true,
            ..MemberWaveState::default()
        };
        changed.notify_all();
    }

    fn member_at(&self, offset: u64) -> Option<usize> {
        self.ranges
            .iter()
            .position(|range| range.start <= offset && offset < range.end)
    }

    fn wait_for_started(&self, target: usize, timeout: Duration) -> bool {
        let deadline = Instant::now() + timeout;
        let (lock, changed) = &*self.state;
        let mut state = lock.lock().unwrap();
        while state.started_members < target {
            let remaining = deadline.saturating_duration_since(Instant::now());
            if remaining.is_zero() {
                return false;
            }
            let (next, timed_out) = changed.wait_timeout(state, remaining).unwrap();
            state = next;
            if timed_out.timed_out() && state.started_members < target {
                return false;
            }
        }
        true
    }

    fn release_wave(&self) {
        let (lock, changed) = &*self.state;
        let mut state = lock.lock().unwrap();
        state.release_generation = state.release_generation.saturating_add(1);
        changed.notify_all();
    }

    fn stop(&self) {
        let (lock, changed) = &*self.state;
        let mut state = lock.lock().unwrap();
        state.armed = false;
        state.release_generation = state.release_generation.saturating_add(1);
        changed.notify_all();
    }

    fn snapshot(&self) -> MemberWaveSnapshot {
        let (lock, _) = &*self.state;
        let state = lock.lock().unwrap();
        MemberWaveSnapshot {
            started_members: state.started_members,
            max_active_members: state.max_active_members,
            max_active_declared_bytes: state.max_active_declared_bytes,
        }
    }
}

impl ReadAt for MemberWaveProbe {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let Some(member) = self.member_at(offset) else {
            return self.inner.read_at(offset, output);
        };
        let (lock, changed) = &*self.state;
        let generation = {
            let mut state = lock.lock().unwrap();
            if !state.armed {
                drop(state);
                return self.inner.read_at(offset, output);
            }
            state.started_members += 1;
            state.active_members.push(member);
            let active_members = state.active_members.len();
            state.max_active_members = state.max_active_members.max(active_members);
            state.max_active_declared_bytes = state
                .max_active_declared_bytes
                .max(u64::try_from(active_members * PART_BYTES).unwrap());
            changed.notify_all();
            state.release_generation
        };
        let mut state = lock.lock().unwrap();
        while state.release_generation == generation {
            state = changed.wait(state).unwrap();
        }
        if let Some(position) = state
            .active_members
            .iter()
            .position(|value| *value == member)
        {
            state.active_members.remove(position);
        }
        drop(state);
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

#[derive(Debug, Default)]
struct FailureState {
    armed: bool,
    alpha_started: bool,
    gamma_started: bool,
    later_started: bool,
    release_alpha: bool,
}

struct OrderedFailureProbe {
    inner: Arc<dyn ReadAt>,
    ranges: Vec<MemberRange>,
    state: Arc<(Mutex<FailureState>, Condvar)>,
}

impl OrderedFailureProbe {
    fn new(inner: Arc<dyn ReadAt>, ranges: Vec<MemberRange>) -> Self {
        Self {
            inner,
            ranges,
            state: Arc::new((Mutex::new(FailureState::default()), Condvar::new())),
        }
    }

    fn member_at(&self, offset: u64) -> Option<&'static str> {
        self.ranges
            .iter()
            .find(|range| range.start <= offset && offset < range.end)
            .map(|range| range.name)
    }

    fn wait_for_first_wave(&self, timeout: Duration) -> bool {
        let deadline = Instant::now() + timeout;
        let (lock, changed) = &*self.state;
        let mut state = lock.lock().unwrap();
        while !(state.alpha_started && state.gamma_started) {
            let remaining = deadline.saturating_duration_since(Instant::now());
            if remaining.is_zero() {
                return false;
            }
            let (next, timed_out) = changed.wait_timeout(state, remaining).unwrap();
            state = next;
            if timed_out.timed_out() && !(state.alpha_started && state.gamma_started) {
                return false;
            }
        }
        true
    }

    fn arm(&self) {
        self.state.0.lock().unwrap().armed = true;
    }

    fn release_alpha(&self) {
        let (lock, changed) = &*self.state;
        let mut state = lock.lock().unwrap();
        state.release_alpha = true;
        changed.notify_all();
    }

    fn later_started(&self) -> bool {
        self.state.0.lock().unwrap().later_started
    }
}

impl ReadAt for OrderedFailureProbe {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if !self.state.0.lock().unwrap().armed {
            return self.inner.read_at(offset, output);
        }
        match self.member_at(offset) {
            Some(name) if name == ALPHA_MEMBER => {
                let (lock, changed) = &*self.state;
                let mut state = lock.lock().unwrap();
                state.alpha_started = true;
                changed.notify_all();
                while !state.release_alpha {
                    state = changed.wait(state).unwrap();
                }
                Err(io::Error::new(
                    io::ErrorKind::BrokenPipe,
                    "low ordinal failure",
                ))
            },
            Some(name) if name == GAMMA_MEMBER => {
                let (lock, changed) = &*self.state;
                let mut state = lock.lock().unwrap();
                state.gamma_started = true;
                changed.notify_all();
                Err(io::Error::new(
                    io::ErrorKind::InvalidData,
                    "high ordinal failure",
                ))
            },
            Some(_) => {
                self.state.0.lock().unwrap().later_started = true;
                self.inner.read_at(offset, output)
            },
            None => self.inner.read_at(offset, output),
        }
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

#[derive(Debug, Clone, Copy, Default)]
struct ReadStatsSnapshot {
    calls: usize,
    bytes_requested: u64,
    max_active_calls: usize,
    max_active_bytes: u64,
}

#[derive(Debug, Default)]
struct ReadStatsState {
    calls: usize,
    bytes_requested: u64,
    active_calls: usize,
    active_bytes: u64,
    max_active_calls: usize,
    max_active_bytes: u64,
}

#[derive(Debug, Default)]
struct ReadStats {
    state: Mutex<ReadStatsState>,
}

impl ReadStats {
    fn reset(&self) {
        let mut state = self.state.lock().unwrap();
        *state = ReadStatsState::default();
    }

    fn enter(&self, requested: usize) {
        let requested = u64::try_from(requested).unwrap();
        let mut state = self.state.lock().unwrap();
        state.calls += 1;
        state.bytes_requested = state.bytes_requested.saturating_add(requested);
        state.active_calls += 1;
        state.active_bytes = state.active_bytes.saturating_add(requested);
        state.max_active_calls = state.max_active_calls.max(state.active_calls);
        state.max_active_bytes = state.max_active_bytes.max(state.active_bytes);
    }

    fn leave(&self, requested: usize) {
        let requested = u64::try_from(requested).unwrap();
        let mut state = self.state.lock().unwrap();
        state.active_calls = state.active_calls.saturating_sub(1);
        state.active_bytes = state.active_bytes.saturating_sub(requested);
    }

    fn snapshot(&self) -> ReadStatsSnapshot {
        let state = self.state.lock().unwrap();
        ReadStatsSnapshot {
            calls: state.calls,
            bytes_requested: state.bytes_requested,
            max_active_calls: state.max_active_calls,
            max_active_bytes: state.max_active_bytes,
        }
    }
}

#[derive(Debug, Default)]
struct GateState {
    armed: bool,
    first_call_entered: bool,
    admitted_calls: usize,
    active_calls: usize,
    overlap: bool,
    released: bool,
}

#[derive(Debug, Default)]
struct ReadGate {
    state: Mutex<GateState>,
    changed: Condvar,
}

impl ReadGate {
    fn arm(&self) {
        let mut state = self.state.lock().unwrap();
        *state = GateState::default();
        state.armed = true;
    }

    fn before_read(&self) {
        let mut state = self.state.lock().unwrap();
        if !state.armed {
            return;
        }
        if state.admitted_calls >= 2 {
            return;
        }
        state.admitted_calls += 1;
        state.first_call_entered = true;
        state.active_calls += 1;
        self.changed.notify_all();
        if state.active_calls >= 2 {
            state.overlap = true;
            state.released = true;
            self.changed.notify_all();
        }
        while !state.released {
            state = self.changed.wait(state).unwrap();
        }
        state.active_calls = state.active_calls.saturating_sub(1);
    }

    fn wait_for_first(&self, timeout: Duration) -> bool {
        let deadline = Instant::now() + timeout;
        let mut state = self.state.lock().unwrap();
        while !state.first_call_entered {
            let remaining = deadline.saturating_duration_since(Instant::now());
            if remaining.is_zero() {
                return false;
            }
            let (next, timed_out) = self.changed.wait_timeout(state, remaining).unwrap();
            state = next;
            if timed_out.timed_out() && !state.first_call_entered {
                return false;
            }
        }
        true
    }

    fn wait_for_overlap(&self, timeout: Duration) -> bool {
        let deadline = Instant::now() + timeout;
        let mut state = self.state.lock().unwrap();
        while !state.overlap {
            let remaining = deadline.saturating_duration_since(Instant::now());
            if remaining.is_zero() {
                break;
            }
            let (next, timed_out) = self.changed.wait_timeout(state, remaining).unwrap();
            state = next;
            if timed_out.timed_out() {
                break;
            }
        }
        state.overlap
    }

    fn release(&self) {
        let mut state = self.state.lock().unwrap();
        state.released = true;
        state.admitted_calls = 2;
        self.changed.notify_all();
    }
}

struct ProbeSource {
    inner: Arc<dyn ReadAt>,
    stats: Arc<ReadStats>,
    gate: Option<Arc<ReadGate>>,
    short_read: Option<usize>,
}

impl ProbeSource {
    fn new(inner: Arc<dyn ReadAt>) -> Self {
        Self {
            inner,
            stats: Arc::new(ReadStats::default()),
            gate: None,
            short_read: None,
        }
    }

    fn with_gate(mut self, gate: Arc<ReadGate>) -> Self {
        self.gate = Some(gate);
        self
    }

    fn with_short_reads(mut self, maximum: usize) -> Self {
        self.short_read = Some(maximum);
        self
    }
}

impl ReadAt for ProbeSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let requested = self
            .short_read
            .map_or(output.len(), |maximum| output.len().min(maximum));
        if requested == 0 {
            return Ok(0);
        }
        self.stats.enter(requested);
        if let Some(gate) = self.gate.as_ref() {
            gate.before_read();
        }
        let result = self.inner.read_at(offset, &mut output[..requested]);
        self.stats.leave(requested);
        result
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

struct MutableSource {
    bytes: Mutex<Vec<u8>>,
    revision: AtomicU64,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Mutex::new(bytes),
            revision: AtomicU64::new(0),
        }
    }

    fn mutate(&self) {
        let mut bytes = self.bytes.lock().unwrap();
        if let Some(byte) = bytes.last_mut() {
            *byte ^= 0x5a;
        }
        self.revision.fetch_add(1, Ordering::AcqRel);
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.lock().unwrap().len())
            .map_err(|_| io::Error::other("fixture source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let bytes = self.bytes.lock().unwrap();
        let Ok(start) = usize::try_from(offset) else {
            return Ok(0);
        };
        let Some(input) = bytes.get(start..) else {
            return Ok(0);
        };
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x534f_5552_4345,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

fn managed_context(
    workers: usize,
    max_in_flight_tasks: usize,
    max_in_flight_bytes: u64,
    min_parallel_bytes: u64,
    work_limit: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "opc-source-backed-batch-test",
        Limits::new(
            64 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            work_limit,
        ),
    );
    let (cancellation, token) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(workers).unwrap(),
        NonZeroUsize::new(max_in_flight_tasks).unwrap(),
        NonZeroU64::new(max_in_flight_bytes).unwrap(),
        min_parallel_bytes,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), token, limits);
    (budget, cancellation, context)
}

fn package_with_context(source: Arc<dyn ReadAt>, context: ExecutionContext) -> SourceBackedPackage {
    SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
        source,
        ReadLimits::default(),
        SourceCacheLimits::new(PART_BYTES * 8, 16).unwrap(),
        context,
    )
    .unwrap()
}

fn requested_uris() -> Vec<PackURI> {
    vec![
        pack(DOCUMENT_URI),
        pack(ALPHA_URI),
        pack(BETA_URI),
        pack(GAMMA_URI),
    ]
}

fn assert_batch_bytes(batch: &litchi_opc::PartBatch, expected: &[Vec<u8>]) {
    assert_eq!(batch.len(), expected.len());
    assert!(!batch.is_empty());
    for (part, expected) in batch.iter().zip(expected) {
        assert_eq!(part.as_bytes(), expected.as_slice());
    }
}

#[test]
fn batch_preserves_request_order_duplicates_and_exact_serial_bytes() {
    let (archive, expected_parts) = archive_bytes(None);
    let serial = SourceBackedPackage::from_vec(archive.clone()).unwrap();
    let requests = vec![
        pack(BETA_URI),
        pack(ALPHA_URI),
        pack(BETA_URI),
        pack(GAMMA_URI),
    ];
    let expected = requests
        .iter()
        .map(|uri| {
            serial
                .part(uri)
                .unwrap()
                .data()
                .unwrap()
                .as_bytes()
                .to_vec()
        })
        .collect::<Vec<_>>();

    let package = SourceBackedPackage::from_vec(archive).unwrap();
    let batch = package.read_parts_ordered(&requests).unwrap();
    assert_batch_bytes(&batch, &expected);
    assert_eq!(batch.get(0).unwrap().as_bytes(), expected[0].as_slice());
    assert!(
        batch
            .get(0)
            .unwrap()
            .shares_allocation_with(batch.get(2).unwrap())
    );
    assert_eq!(
        package.cache_diagnostics().successful_loads,
        expected_parts.len() as u64 - 1
    );
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
}

#[test]
fn batch_short_read_provider_preserves_exact_bytes() {
    let (archive, expected_parts) = archive_bytes(None);
    let expected = expected_parts
        .iter()
        .map(|(_, bytes)| bytes.clone())
        .collect::<Vec<_>>();
    let inner: Arc<dyn ReadAt> = Arc::new(litchi_core::OwnedSource::new(archive));
    let probe = ProbeSource::new(inner).with_short_reads(7);
    let package = SourceBackedPackage::from_read_at(Arc::new(probe)).unwrap();
    let requests = requested_uris();
    let batch = package.read_parts_ordered(&requests).unwrap();
    assert_batch_bytes(&batch, &expected);
    assert!(package.cache_diagnostics().in_flight_loads == 0);
}

#[test]
fn batch_parallel_overlap_obeys_task_and_byte_caps() {
    let (archive, expected_parts) = archive_bytes(None);
    let expected = expected_parts
        .iter()
        .map(|(_, bytes)| bytes.clone())
        .collect::<Vec<_>>();
    let ranges = stored_member_ranges(&archive);
    assert_eq!(ranges.len(), 4);
    let inner: Arc<dyn ReadAt> = Arc::new(litchi_core::OwnedSource::new(archive));
    let probe = Arc::new(MemberWaveProbe::new(inner, ranges));
    let (budget, _cancellation, context) =
        managed_context(4, 4, u64::try_from(PART_BYTES * 2).unwrap(), 1, u64::MAX);
    let package = package_with_context(probe.clone(), context);
    probe.arm();
    let requests = requested_uris();
    let worker = thread::spawn(move || {
        let result = package.read_parts_ordered(&requests);
        (result, package)
    });

    let first_wave = probe.wait_for_started(2, Duration::from_secs(2));
    let mut second_wave = false;
    if first_wave {
        probe.release_wave();
        second_wave = probe.wait_for_started(4, Duration::from_secs(2));
    }
    probe.stop();
    let (result, package) = worker.join().unwrap();
    let batch = result.unwrap();
    assert_batch_bytes(&batch, &expected);
    let snapshot = probe.snapshot();
    assert!(first_wave, "the first bounded worker wave did not start");
    assert!(second_wave, "the second bounded worker wave did not start");
    assert!(snapshot.started_members >= 4);
    assert!(snapshot.max_active_members >= 2);
    assert!(snapshot.max_active_members <= 2);
    assert!(snapshot.max_active_declared_bytes <= u64::try_from(PART_BYTES * 2).unwrap());
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    drop(batch);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn batch_below_parallel_threshold_uses_serial_fallback() {
    let (archive, expected_parts) = archive_bytes(None);
    let expected = expected_parts
        .iter()
        .skip(1)
        .take(2)
        .map(|(_, bytes)| bytes.clone())
        .collect::<Vec<_>>();
    let inner: Arc<dyn ReadAt> = Arc::new(litchi_core::OwnedSource::new(archive));
    let gate = Arc::new(ReadGate::default());
    let probe = Arc::new(ProbeSource::new(inner).with_gate(Arc::clone(&gate)));
    let stats = Arc::clone(&probe.stats);
    let aggregate = u64::try_from(PART_BYTES * 2).unwrap();
    let (_budget, _cancellation, context) =
        managed_context(4, 4, aggregate + 2, aggregate + 1, u64::MAX);
    let package = package_with_context(probe, context);
    stats.reset();
    gate.arm();
    let requests = vec![pack(ALPHA_URI), pack(BETA_URI)];
    let worker = thread::spawn(move || {
        let result = package.read_parts_ordered(&requests);
        (result, package)
    });

    let first = gate.wait_for_first(Duration::from_secs(2));
    let overlapped = first && gate.wait_for_overlap(Duration::from_secs(2));
    gate.release();
    let (result, package) = worker.join().unwrap();
    let batch = result.unwrap();
    assert_batch_bytes(&batch, &expected);
    assert!(first, "serial fallback did not enter the source gate");
    assert!(
        !overlapped,
        "below-threshold work must stay on the serial path"
    );
    let stats = stats.snapshot();
    assert!(stats.calls >= 2);
    assert!(stats.bytes_requested >= aggregate);
    assert_eq!(stats.max_active_calls, 1);
    assert!(stats.max_active_bytes > 0);
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
}

#[test]
fn batch_missing_and_corrupt_parts_fail_without_partial_results_or_flights() {
    let (archive, _) = archive_bytes(None);
    let package = SourceBackedPackage::from_vec(archive).unwrap();
    let missing =
        package.read_parts_ordered(&[pack(ALPHA_URI), pack("/custom/missing.bin"), pack(BETA_URI)]);
    assert!(matches!(missing, Err(OpcError::PartNotFound(_))));
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    assert_eq!(package.cache_diagnostics().successful_loads, 0);

    let (archive, _) = archive_bytes(Some(GAMMA_MEMBER));
    let (budget, _cancellation, context) =
        managed_context(2, 2, u64::try_from(PART_BYTES * 2).unwrap(), 1, u64::MAX);
    let package = package_with_context(Arc::new(litchi_core::OwnedSource::new(archive)), context);
    let work_before = budget.used(Resource::Work);
    let corrupt = package.read_parts_ordered(&[pack(GAMMA_URI)]);
    assert!(matches!(corrupt, Err(OpcError::ZipError(_))));
    assert!(budget.used(Resource::Work) > work_before);
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.in_flight_loads, 0);
    assert!(diagnostics.successful_loads <= diagnostics.cold_loads);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn batch_request_count_limit_remains_a_typed_read_limit() {
    let (archive, _) = archive_bytes(None);
    let limits = ReadLimits::builder().max_parts(4).unwrap().build().unwrap();
    let package = SourceBackedPackage::from_read_at_with_limits(
        Arc::new(litchi_core::OwnedSource::new(archive)),
        limits,
    )
    .unwrap();
    let result = package.read_parts_ordered(&[
        pack(ALPHA_URI),
        pack(BETA_URI),
        pack(GAMMA_URI),
        pack(DOCUMENT_URI),
        pack(ALPHA_URI),
    ]);
    assert!(matches!(
        result,
        Err(OpcError::ReadLimit {
            resource: ReadResource::Parts,
            actual,
            maximum,
        }) if actual == 5 && maximum == 4
    ));
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
}

#[test]
fn managed_batch_releases_finite_memory_and_object_reservations_after_drop() {
    let (archive, expected_parts) = archive_bytes(None);
    let expected = expected_parts
        .iter()
        .map(|(_, bytes)| bytes.clone())
        .collect::<Vec<_>>();
    let (budget, _cancellation, context) =
        managed_context(2, 2, u64::try_from(PART_BYTES * 2).unwrap(), 1, u64::MAX);
    let package = package_with_context(Arc::new(litchi_core::OwnedSource::new(archive)), context);
    let baseline_memory = budget.used(Resource::Memory);
    let baseline_objects = budget.used(Resource::Objects);
    let requests = requested_uris();
    let batch = package.read_parts_ordered(&requests).unwrap();
    assert_batch_bytes(&batch, &expected);
    assert!(batch.get(0).unwrap().into_arc().is_err());
    let cold_work = budget.used(Resource::Work);
    assert!(cold_work > 0);
    drop(batch);
    assert!(budget.used(Resource::Memory) >= baseline_memory);
    assert!(budget.used(Resource::Objects) >= baseline_objects);

    let warm = package.read_parts_ordered(&requests).unwrap();
    assert_eq!(budget.used(Resource::Work), cold_work);
    drop(warm);
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
    assert!(budget.used(Resource::Work) >= cold_work);
}

#[test]
fn warm_batch_holds_structural_memory_until_the_returned_batch_drops() {
    let (archive, _) = archive_bytes(None);
    let (budget, _cancellation, context) =
        managed_context(2, 2, u64::try_from(PART_BYTES * 2).unwrap(), 1, u64::MAX);
    let package = package_with_context(Arc::new(litchi_core::OwnedSource::new(archive)), context);
    let requests = requested_uris();

    // Establish a warm payload cache and release the first returned
    // collection before taking the baseline for the structural reservation.
    let first = package.read_parts_ordered(&requests).unwrap();
    drop(first);
    let warm_cache = package.cache_diagnostics();
    let before_memory = budget.used(Resource::Memory);
    let before_objects = budget.used(Resource::Objects);

    let batch = package.read_parts_ordered(&requests).unwrap();
    let after_memory = budget.used(Resource::Memory);
    let after_objects = budget.used(Resource::Objects);
    let after_cache = package.cache_diagnostics();
    assert_eq!(
        after_cache.budget_cache_reserved_bytes,
        warm_cache.budget_cache_reserved_bytes
    );
    assert_eq!(
        after_cache.budget_cache_reserved_objects,
        warm_cache.budget_cache_reserved_objects
    );
    assert!(
        after_memory > before_memory,
        "returned PartBatch must retain structural memory"
    );
    assert!(after_objects >= before_objects);

    drop(batch);
    assert_eq!(budget.used(Resource::Memory), before_memory);
    assert_eq!(budget.used(Resource::Objects), before_objects);
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn batch_empty_one_worker_and_oversized_single_part_use_serial_fallbacks() {
    let (archive, expected_parts) = archive_bytes(None);
    let empty = SourceBackedPackage::from_vec(archive.clone()).unwrap();
    let empty_batch = empty.read_parts_ordered(&[]).unwrap();
    assert!(empty_batch.is_empty());
    assert_eq!(empty_batch.len(), 0);
    drop(empty_batch);
    drop(empty);

    let expected_alpha = expected_parts
        .iter()
        .find(|(uri, _)| uri.as_str() == ALPHA_URI)
        .map(|(_, bytes)| bytes.clone())
        .unwrap();
    let (budget, _cancellation, context) =
        managed_context(1, 1, u64::try_from(PART_BYTES).unwrap(), 0, u64::MAX);
    let package = package_with_context(
        Arc::new(litchi_core::OwnedSource::new(archive.clone())),
        context,
    );
    let one = package.read_parts_ordered(&[pack(ALPHA_URI)]).unwrap();
    assert_eq!(one.len(), 1);
    assert_eq!(one.get(0).unwrap().as_bytes(), expected_alpha.as_slice());
    drop(one);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);

    let (budget, _cancellation, context) = managed_context(4, 4, 1, 1, u64::MAX);
    let package = package_with_context(Arc::new(litchi_core::OwnedSource::new(archive)), context);
    let oversized = package.read_parts_ordered(&[pack(GAMMA_URI)]).unwrap();
    assert_eq!(oversized.len(), 1);
    assert_eq!(oversized.get(0).unwrap().as_bytes().len(), PART_BYTES);
    drop(oversized);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn deflated_multi_part_batch_decodes_varied_payloads_on_an_explicit_two_mib_stack() {
    let (archive, expected_parts) = deflated_archive_bytes();
    let requests = expected_parts
        .iter()
        .skip(1)
        .map(|(uri, _)| uri.clone())
        .collect::<Vec<_>>();
    let expected = expected_parts
        .iter()
        .skip(1)
        .map(|(_, bytes)| bytes.clone())
        .collect::<Vec<_>>();
    let (budget, _cancellation, context) =
        managed_context(3, 3, u64::try_from(512 * 1024).unwrap(), 1, u64::MAX);
    let package = package_with_context(Arc::new(litchi_core::OwnedSource::new(archive)), context);
    let worker = thread::Builder::new()
        .name("opc-deflated-batch-2mib".to_owned())
        .stack_size(2 * 1024 * 1024)
        .spawn(move || {
            let result = package.read_parts_ordered(&requests);
            (result, package)
        })
        .unwrap();
    let (result, package) = worker.join().unwrap();
    let batch = result.unwrap();
    assert_batch_bytes(&batch, &expected);
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    assert!(package.cache_diagnostics().successful_loads >= 4);
    drop(batch);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn batch_work_limit_preserves_the_dynamic_resource_limit_scope() {
    let (archive, _) = archive_bytes(None);
    let work_limit = u64::try_from(PART_BYTES * 2 - 1).unwrap();
    let (budget, _cancellation, context) =
        managed_context(2, 2, u64::try_from(PART_BYTES * 2).unwrap(), 1, work_limit);
    let package = package_with_context(Arc::new(litchi_core::OwnedSource::new(archive)), context);
    let result = package.read_parts_ordered(&[pack(ALPHA_URI), pack(BETA_URI)]);
    assert!(matches!(
        result,
        Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
            if limit.resource == Resource::Work
                && limit.scope.as_ref() == "opc-source-backed-batch-test"
                && limit.limit == work_limit
    ));
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn batch_reports_low_ordinal_failure_and_does_not_start_a_later_wave() {
    let (archive, _) = archive_bytes(None);
    let ranges = stored_member_ranges(&archive);
    assert_eq!(ranges.len(), 4);
    let inner: Arc<dyn ReadAt> = Arc::new(litchi_core::OwnedSource::new(archive));
    let probe = Arc::new(OrderedFailureProbe::new(inner, ranges));
    let (budget, _cancellation, context) =
        managed_context(2, 2, u64::try_from(PART_BYTES * 2).unwrap(), 1, u64::MAX);
    let package = package_with_context(probe.clone(), context);
    probe.arm();
    let requests = vec![
        pack(ALPHA_URI),
        pack(GAMMA_URI),
        pack(BETA_URI),
        pack(DOCUMENT_URI),
    ];
    let worker = thread::spawn(move || {
        let result = package.read_parts_ordered(&requests);
        (result, package)
    });

    let first_wave = probe.wait_for_first_wave(Duration::from_secs(2));
    // Always release the low ordinal failure before joining.  The assertion
    // below must never strand a worker behind the provider gate.
    probe.release_alpha();
    let (result, package) = worker.join().unwrap();
    assert!(
        first_wave,
        "the two first-wave failure members did not start"
    );
    assert!(matches!(
        result,
        Err(OpcError::IoError(error)) if error.kind() == io::ErrorKind::BrokenPipe
    ));
    assert!(
        !probe.later_started(),
        "a failed first wave must not dispatch a later wave"
    );
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn batch_cancellation_during_a_gated_load_returns_cancelled_and_joins_workers() {
    let (archive, _) = archive_bytes(None);
    let inner: Arc<dyn ReadAt> = Arc::new(litchi_core::OwnedSource::new(archive));
    let gate = Arc::new(ReadGate::default());
    let probe = Arc::new(ProbeSource::new(inner).with_gate(Arc::clone(&gate)));
    let (budget, cancellation, context) =
        managed_context(2, 2, u64::try_from(PART_BYTES * 2).unwrap(), 1, u64::MAX);
    let package = package_with_context(probe, context);
    gate.arm();
    let worker = thread::spawn(move || {
        let result = package.read_parts_ordered(&requested_uris());
        (result, package)
    });
    let first = gate.wait_for_first(Duration::from_secs(2));
    cancellation.cancel();
    gate.release();
    let (result, package) = worker.join().unwrap();
    assert!(first, "cancellation test did not enter the source gate");
    assert!(matches!(result, Err(OpcError::Cancelled)));
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn batch_source_mutation_during_a_gated_load_returns_source_changed() {
    let (archive, _) = archive_bytes(None);
    let mutable = Arc::new(MutableSource::new(archive));
    let inner: Arc<dyn ReadAt> = mutable.clone();
    let gate = Arc::new(ReadGate::default());
    let probe = Arc::new(ProbeSource::new(inner).with_gate(Arc::clone(&gate)));
    let package = SourceBackedPackage::from_read_at(probe).unwrap();
    gate.arm();
    let worker = thread::spawn(move || {
        let result = package.read_parts_ordered(&[pack(ALPHA_URI), pack(BETA_URI)]);
        (result, package)
    });
    let first = gate.wait_for_first(Duration::from_secs(2));
    if first {
        mutable.mutate();
    }
    gate.release();
    let (result, package) = worker.join().unwrap();
    assert!(first, "source mutation test did not enter the source gate");
    assert!(matches!(result, Err(OpcError::SourceChanged { .. })));
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    drop(package);
}
