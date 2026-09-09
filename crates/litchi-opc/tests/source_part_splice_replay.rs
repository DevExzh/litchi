#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "focused replay-boundary assertions intentionally panic on fixture errors"
)]

//! Public-boundary coverage for the format-neutral authored replay payload.
//!
//! The fixture is deliberately small.  It has one XML target and one opaque
//! member, so the tests can check both logical decoded output and physical
//! preservation without depending on private ZIP implementation types.

use std::io::{self, Cursor, Read, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc, Condvar, Mutex, RwLock,
    atomic::{AtomicUsize, Ordering},
};
use std::thread;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, ReadAt, Resource, SourceVersion,
};
use litchi_opc::{
    OpcError, PackURI, ReadLimits, SourceArtifactFingerprint, SourceBackedPackage,
    SourcePartSpliceLimits, SourcePartSpliceProof, SourcePartSpliceReplay,
    SourcePartSpliceReplayError, SourcePartSpliceReplayProof, SpliceResource,
};
use sha2::{Digest as _, Sha256};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const TARGET_MEMBER: &str = "word/document.xml";
const TARGET_URI: &str = "/word/document.xml";
const OPAQUE_MEMBER: &str = "opaque.bin";
const SOURCE: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?><root><before/></root>"#;
const INSERTION_OFFSET: usize = SOURCE.len() - b"</root>".len();
const FRAGMENT: &[u8] = b"<inserted/>";
const CHANGED_FRAGMENT: &[u8] = b"<inserteD/>";

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).expect("fixture URI must be canonical")
}

fn digest(bytes: &[u8]) -> [u8; 32] {
    Sha256::digest(bytes).into()
}

fn candidate(source: &[u8], fragment: &[u8]) -> Vec<u8> {
    candidate_at(source, INSERTION_OFFSET, fragment)
}

fn candidate_at(source: &[u8], insertion_offset: usize, fragment: &[u8]) -> Vec<u8> {
    [
        &source[..insertion_offset],
        fragment,
        &source[insertion_offset..],
    ]
    .concat()
}

fn archive(deflated: bool) -> Vec<u8> {
    archive_with_source(deflated, SOURCE)
}

fn archive_with_source(deflated: bool, source: &[u8]) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#
    );
    let relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rDoc" Type="{OFFICE_DOCUMENT_REL}" Target="{TARGET_MEMBER}"/></Relationships>"#
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content-types fixture must be writable");
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .expect("relationships fixture must be writable");
    writer
        .write_stored(OPAQUE_MEMBER, b"opaque physical bytes\0\xff")
        .expect("opaque fixture must be writable");
    if deflated {
        writer
            .write_deflated_sized(TARGET_MEMBER, source)
            .expect("Deflate target fixture must be writable");
    } else {
        writer
            .write_stored(TARGET_MEMBER, source)
            .expect("Store target fixture must be writable");
    }
    writer
        .finish_to_bytes()
        .expect("fixture archive must finish")
}

fn multi_window_source() -> (Vec<u8>, usize) {
    let mut source = br#"<?xml version="1.0" encoding="UTF-8"?><root><before>"#.to_vec();
    source.extend(std::iter::repeat_n(b'a', 2 * 512 + 73));
    source.extend_from_slice(b"</before></root>");
    let insertion_offset = source.len() - b"</root>".len();
    (source, insertion_offset)
}

fn multi_window_fragment() -> Vec<u8> {
    let mut fragment = b"<inserted>".to_vec();
    fragment.extend(std::iter::repeat_n(b'b', 3 * 512 + 113));
    fragment.extend_from_slice(b"</inserted>");
    fragment
}

fn open(deflated: bool) -> (SourceBackedPackage, Vec<u8>) {
    let bytes = archive(deflated);
    let package = SourceBackedPackage::from_vec(bytes.clone()).expect("fixture package must open");
    (package, bytes)
}

fn proof(package: &SourceBackedPackage, fragment: &[u8]) -> SourcePartSpliceProof {
    proof_for(package, SOURCE, INSERTION_OFFSET, fragment)
}

fn proof_for(
    package: &SourceBackedPackage,
    source: &[u8],
    insertion_offset: usize,
    fragment: &[u8],
) -> SourcePartSpliceProof {
    let candidate = candidate_at(source, insertion_offset, fragment);
    SourcePartSpliceProof {
        source_version: package
            .source_version()
            .expect("fixture source version must be available"),
        source_len: source.len() as u64,
        source_sha256: digest(source),
        insertion_offset: insertion_offset as u64,
        fragment_len: fragment.len() as u64,
        fragment_sha256: digest(fragment),
        candidate_len: candidate.len() as u64,
        candidate_sha256: digest(&candidate),
    }
}

fn target_bytes(archive: &[u8]) -> Vec<u8> {
    SourceBackedPackage::from_vec(archive.to_vec())
        .expect("published archive must open")
        .part(&pack(TARGET_URI))
        .expect("published target must exist")
        .data()
        .expect("published target must read")
        .as_bytes()
        .to_vec()
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    managed_context_with_output(memory, u64::MAX)
}

fn managed_context_with_output(
    memory: u64,
    output: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let budget = Budget::root(
        "source-part-splice-replay-test",
        Limits::new(memory, u64::MAX, output, u64::MAX, u64::MAX, u64::MAX),
    );
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("worker limit is nonzero"),
        NonZeroUsize::new(1).expect("operation limit is nonzero"),
        NonZeroU64::new(memory).expect("memory limit is nonzero"),
        0,
    )
    .expect("execution limits must be valid");
    (
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn publication_memory_envelope(archive_bytes: &[u8], limits: SourcePartSpliceLimits) -> u64 {
    let indexed = soapberry_zip::office::IndexedArchive::from_reader(
        Cursor::new(archive_bytes.to_vec()),
        archive_bytes.len() as u64,
    )
    .expect("fixture archive must have an indexed memory envelope");
    let target = indexed
        .entry_id(TARGET_MEMBER)
        .expect("fixture target must be indexed");
    indexed
        .preservation_memory_upper_bound()
        .expect("preservation bound must fit u64")
        .checked_add(soapberry_zip::RECOMMENDED_BUFFER_SIZE as u64)
        .expect("preservation scratch bound must fit u64")
        .checked_add(soapberry_zip::replay_memory_upper_bound().expect("replay bound must fit u64"))
        .and_then(|value| {
            value.checked_add(
                indexed
                    .verified_entry_reader_memory_upper_bound(target)
                    .expect("verified reader bound must be available"),
            )
        })
        .and_then(|value| {
            value.checked_add(
                u64::try_from(
                    limits
                        .xml_audit_limits
                        .streaming_memory_upper_bound()
                        .expect("fixture XML audit profile must have a finite bound"),
                )
                .expect("XML workspace must fit u64"),
            )
        })
        .expect("publication memory envelope must fit u64")
}

#[derive(Clone, Copy, Debug)]
enum ReplayMode {
    Exact,
    Truncated,
    Extra,
    HashChanged,
    OpenError,
    ReadError,
    ReadErrorAtOpen(usize),
    ChangeAfterFirstOpen,
    ChangeAtOpen(usize),
}

struct ReplayProvider {
    payload: Arc<Vec<u8>>,
    mode: ReplayMode,
    opens: AtomicUsize,
}

impl ReplayProvider {
    fn new(payload: &[u8], mode: ReplayMode) -> Arc<Self> {
        Arc::new(Self {
            payload: Arc::new(payload.to_vec()),
            mode,
            opens: AtomicUsize::new(0),
        })
    }

    fn open_count(&self) -> usize {
        self.opens.load(Ordering::Acquire)
    }

    fn bytes_for(&self, open_number: usize) -> Vec<u8> {
        let mut bytes = self.payload.as_ref().clone();
        match self.mode {
            ReplayMode::Exact => {},
            ReplayMode::Truncated => {
                bytes.pop();
            },
            ReplayMode::Extra => bytes.push(b'X'),
            ReplayMode::HashChanged => {
                bytes = CHANGED_FRAGMENT.to_vec();
            },
            ReplayMode::OpenError | ReplayMode::ReadError | ReplayMode::ReadErrorAtOpen(_) => {},
            ReplayMode::ChangeAfterFirstOpen if open_number > 1 => {
                bytes = CHANGED_FRAGMENT.to_vec();
            },
            ReplayMode::ChangeAfterFirstOpen => {},
            ReplayMode::ChangeAtOpen(change_at) if open_number >= change_at => {
                bytes = CHANGED_FRAGMENT.to_vec();
            },
            ReplayMode::ChangeAtOpen(_) => {},
        }
        bytes
    }
}

impl SourcePartSpliceReplay for ReplayProvider {
    fn proof(&self) -> SourcePartSpliceReplayProof {
        SourcePartSpliceReplayProof {
            encoded_len: self.payload.len() as u64,
            encoded_sha256: digest(self.payload.as_slice()),
        }
    }

    fn open(&self) -> Result<Box<dyn Read + '_>, SourcePartSpliceReplayError> {
        let open_number = self.opens.fetch_add(1, Ordering::AcqRel) + 1;
        match self.mode {
            ReplayMode::OpenError => Err(SourcePartSpliceReplayError::Io(io::Error::new(
                io::ErrorKind::NotFound,
                "replay fixture open failed",
            ))),
            ReplayMode::ReadError => {
                Ok(Box::new(FailingReader::new(self.bytes_for(open_number), 2)))
            },
            ReplayMode::ReadErrorAtOpen(change_at) if open_number >= change_at => {
                Ok(Box::new(FailingReader::new(self.bytes_for(open_number), 2)))
            },
            ReplayMode::ReadErrorAtOpen(_) => {
                Ok(Box::new(Cursor::new(self.bytes_for(open_number))))
            },
            _ => Ok(Box::new(Cursor::new(self.bytes_for(open_number)))),
        }
    }
}

struct ChunkedReplayProvider {
    payload: Arc<Vec<u8>>,
    chunk_size: usize,
    opens: AtomicUsize,
    short_final_reads: Arc<AtomicUsize>,
}

impl ChunkedReplayProvider {
    fn new(payload: Vec<u8>, chunk_size: usize) -> Arc<Self> {
        assert!(
            chunk_size != 0,
            "chunked replay fixture needs a nonzero chunk"
        );
        Arc::new(Self {
            payload: Arc::new(payload),
            chunk_size,
            opens: AtomicUsize::new(0),
            short_final_reads: Arc::new(AtomicUsize::new(0)),
        })
    }

    fn open_count(&self) -> usize {
        self.opens.load(Ordering::Acquire)
    }

    fn short_final_reads(&self) -> usize {
        self.short_final_reads.load(Ordering::Acquire)
    }
}

impl SourcePartSpliceReplay for ChunkedReplayProvider {
    fn proof(&self) -> SourcePartSpliceReplayProof {
        SourcePartSpliceReplayProof {
            encoded_len: self.payload.len() as u64,
            encoded_sha256: digest(self.payload.as_slice()),
        }
    }

    fn open(&self) -> Result<Box<dyn Read + '_>, SourcePartSpliceReplayError> {
        self.opens.fetch_add(1, Ordering::AcqRel);
        Ok(Box::new(ChunkedReplayReader {
            bytes: self.payload.as_ref().clone(),
            position: 0,
            chunk_size: self.chunk_size,
            short_final_reads: Arc::clone(&self.short_final_reads),
        }))
    }
}

struct ChunkedReplayReader {
    bytes: Vec<u8>,
    position: usize,
    chunk_size: usize,
    short_final_reads: Arc<AtomicUsize>,
}

impl Read for ChunkedReplayReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.position >= self.bytes.len() {
            return Ok(0);
        }
        let remaining = self.bytes.len() - self.position;
        let count = output.len().min(self.chunk_size).min(remaining);
        output[..count].copy_from_slice(&self.bytes[self.position..self.position + count]);
        self.position += count;
        if self.position == self.bytes.len() && count < self.chunk_size {
            self.short_final_reads.fetch_add(1, Ordering::AcqRel);
        }
        Ok(count)
    }
}

struct MutableProofProvider {
    payload: Arc<Vec<u8>>,
    proof_calls: AtomicUsize,
    bytes_read: Arc<AtomicUsize>,
}

impl MutableProofProvider {
    fn new() -> Arc<Self> {
        let mut payload = FRAGMENT.to_vec();
        payload.extend(std::iter::repeat_n(b'a', 128 * 1024));
        Arc::new(Self {
            payload: Arc::new(payload),
            proof_calls: AtomicUsize::new(0),
            bytes_read: Arc::new(AtomicUsize::new(0)),
        })
    }

    fn bytes_read(&self) -> usize {
        self.bytes_read.load(Ordering::Acquire)
    }
}

impl SourcePartSpliceReplay for MutableProofProvider {
    fn proof(&self) -> SourcePartSpliceReplayProof {
        if self.proof_calls.fetch_add(1, Ordering::AcqRel) == 0 {
            SourcePartSpliceReplayProof {
                encoded_len: FRAGMENT.len() as u64,
                encoded_sha256: digest(FRAGMENT),
            }
        } else {
            SourcePartSpliceReplayProof {
                encoded_len: u64::MAX,
                encoded_sha256: [0xa5; 32],
            }
        }
    }

    fn open(&self) -> Result<Box<dyn Read + '_>, SourcePartSpliceReplayError> {
        Ok(Box::new(CountingReader {
            inner: Cursor::new(self.payload.as_ref().clone()),
            bytes_read: Arc::clone(&self.bytes_read),
        }))
    }
}

struct CountingReader {
    inner: Cursor<Vec<u8>>,
    bytes_read: Arc<AtomicUsize>,
}

impl Read for CountingReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        let read = self.inner.read(output)?;
        self.bytes_read.fetch_add(read, Ordering::AcqRel);
        Ok(read)
    }
}

struct FailingReader {
    bytes: Vec<u8>,
    position: usize,
    fail_at: usize,
}

impl FailingReader {
    fn new(bytes: Vec<u8>, fail_at: usize) -> Self {
        Self {
            bytes,
            position: 0,
            fail_at,
        }
    }
}

impl Read for FailingReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.position >= self.fail_at {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "replay fixture reader failed",
            ));
        }
        let count = output
            .len()
            .min(self.fail_at - self.position)
            .min(self.bytes.len() - self.position);
        output[..count].copy_from_slice(&self.bytes[self.position..self.position + count]);
        self.position += count;
        Ok(count)
    }
}

struct CancellingReplayProvider {
    payload: Arc<Vec<u8>>,
    cancellation: CancellationSource,
}

impl CancellingReplayProvider {
    fn new(cancellation: CancellationSource) -> Arc<Self> {
        Arc::new(Self {
            payload: Arc::new(FRAGMENT.to_vec()),
            cancellation,
        })
    }
}

impl SourcePartSpliceReplay for CancellingReplayProvider {
    fn proof(&self) -> SourcePartSpliceReplayProof {
        SourcePartSpliceReplayProof {
            encoded_len: self.payload.len() as u64,
            encoded_sha256: digest(self.payload.as_slice()),
        }
    }

    fn open(&self) -> Result<Box<dyn Read + '_>, SourcePartSpliceReplayError> {
        Ok(Box::new(CancellingReader {
            inner: Cursor::new(self.payload.as_ref().clone()),
            cancellation: self.cancellation.clone(),
            cancelled: false,
        }))
    }
}

struct CancellingReader {
    inner: Cursor<Vec<u8>>,
    cancellation: CancellationSource,
    cancelled: bool,
}

impl Read for CancellingReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        let read = self.inner.read(output)?;
        if !self.cancelled {
            self.cancellation.cancel();
            self.cancelled = true;
        }
        Ok(read)
    }
}

struct BlockingReplayProvider {
    payload: Arc<Vec<u8>>,
    opens: AtomicUsize,
    entered: Arc<(Mutex<bool>, Condvar)>,
    release: Arc<(Mutex<bool>, Condvar)>,
}

impl BlockingReplayProvider {
    fn new(payload: &[u8]) -> Arc<Self> {
        Arc::new(Self {
            payload: Arc::new(payload.to_vec()),
            opens: AtomicUsize::new(0),
            entered: Arc::new((Mutex::new(false), Condvar::new())),
            release: Arc::new((Mutex::new(false), Condvar::new())),
        })
    }

    fn wait_until_first_publication_reader(&self) {
        let (lock, condition) = &*self.entered;
        let mut entered = lock.lock().expect("entered lock must work");
        while !*entered {
            entered = condition
                .wait(entered)
                .expect("entered condition wait must work");
        }
    }

    fn release_first_publication_reader(&self) {
        let (lock, condition) = &*self.release;
        *lock.lock().expect("release lock must work") = true;
        condition.notify_all();
    }
}

impl SourcePartSpliceReplay for BlockingReplayProvider {
    fn proof(&self) -> SourcePartSpliceReplayProof {
        SourcePartSpliceReplayProof {
            encoded_len: self.payload.len() as u64,
            encoded_sha256: digest(self.payload.as_slice()),
        }
    }

    fn open(&self) -> Result<Box<dyn Read + '_>, SourcePartSpliceReplayError> {
        let open_number = self.opens.fetch_add(1, Ordering::AcqRel) + 1;
        if open_number == 3 {
            let (lock, condition) = &*self.entered;
            *lock.lock().expect("entered lock must work") = true;
            condition.notify_all();

            let (lock, condition) = &*self.release;
            let mut released = lock.lock().expect("release lock must work");
            while !*released {
                released = condition
                    .wait(released)
                    .expect("release condition wait must work");
            }
        }
        Ok(Box::new(Cursor::new(self.payload.as_ref().clone())))
    }
}

fn replay_plan<'a, R>(
    package: &'a SourceBackedPackage,
    provider: Arc<R>,
    limits: SourcePartSpliceLimits,
) -> litchi_opc::SourcePartSplicePlan<'a>
where
    R: SourcePartSpliceReplay + 'static,
{
    package
        .prepare_source_part_splice_with_replay(
            &pack(TARGET_URI),
            proof(package, FRAGMENT),
            provider,
            limits,
        )
        .expect("replay fixture proof must prepare")
}

fn expected_replay_artifact(deflated: bool) -> (Vec<u8>, u64, SourceArtifactFingerprint) {
    let (package, _) = open(deflated);
    let provider = ReplayProvider::new(FRAGMENT, ReplayMode::Exact);
    let plan = replay_plan(&package, provider, SourcePartSpliceLimits::default());
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("reference replay publication must succeed");
    let length = output.len() as u64;
    let hash = SourceArtifactFingerprint::from_sha256(digest(&output));
    (output, length, hash)
}

fn member_central_record(bytes: &[u8], wanted: &str) -> Vec<u8> {
    let eocd = bytes.len() - 22;
    let mut offset = u32_at(bytes, eocd + 16);
    for _ in 0..u16_at(bytes, eocd + 10) {
        assert_eq!(&bytes[offset..offset + 4], b"PK\x01\x02");
        let name_len = u16_at(bytes, offset + 28);
        let extra_len = u16_at(bytes, offset + 30);
        let comment_len = u16_at(bytes, offset + 32);
        let end = offset + 46 + name_len + extra_len + comment_len;
        if &bytes[offset + 46..offset + 46 + name_len] == wanted.as_bytes() {
            let mut record = bytes[offset..end].to_vec();
            record[42..46].fill(0);
            return record;
        }
        offset = end;
    }
    panic!("central member {wanted} is missing")
}

fn member_local_record(bytes: &[u8], wanted: &str) -> Vec<u8> {
    let eocd = bytes.len() - 22;
    let mut central = u32_at(bytes, eocd + 16);
    for _ in 0..u16_at(bytes, eocd + 10) {
        assert_eq!(&bytes[central..central + 4], b"PK\x01\x02");
        let name_len = u16_at(bytes, central + 28);
        let extra_len = u16_at(bytes, central + 30);
        let comment_len = u16_at(bytes, central + 32);
        let end = central + 46 + name_len + extra_len + comment_len;
        if &bytes[central + 46..central + 46 + name_len] == wanted.as_bytes() {
            let local = u32_at(bytes, central + 42);
            let local_name_len = u16_at(bytes, local + 26);
            let local_extra_len = u16_at(bytes, local + 28);
            let payload = local + 30 + local_name_len + local_extra_len;
            let compressed_len = u32_at(bytes, central + 20);
            return bytes[local..payload + compressed_len].to_vec();
        }
        central = end;
    }
    panic!("local member {wanted} is missing")
}

fn u16_at(bytes: &[u8], offset: usize) -> usize {
    usize::from(u16::from_le_bytes(
        bytes[offset..offset + 2]
            .try_into()
            .expect("ZIP16 fixture field must fit"),
    ))
}

fn u32_at(bytes: &[u8], offset: usize) -> usize {
    usize::try_from(u32::from_le_bytes(
        bytes[offset..offset + 4]
            .try_into()
            .expect("ZIP32 fixture field must fit"),
    ))
    .expect("ZIP32 fixture offset must fit usize")
}

#[test]
fn replay_store_and_deflate_preserve_opaque_members_and_target_bytes() {
    for deflated in [false, true] {
        let (package, source_archive) = open(deflated);
        let provider = ReplayProvider::new(FRAGMENT, ReplayMode::Exact);
        let plan = replay_plan(
            &package,
            provider.clone(),
            SourcePartSpliceLimits::default(),
        );
        let mut output = Vec::new();
        let publication = plan
            .write_to_stream(&mut output)
            .expect("replay splice must publish");

        assert_eq!(target_bytes(&output), candidate(SOURCE, FRAGMENT));
        assert_eq!(publication.candidate_artifact_len(), output.len() as u64);
        assert_eq!(
            member_local_record(&source_archive, OPAQUE_MEMBER),
            member_local_record(&output, OPAQUE_MEMBER),
            "opaque local record must remain byte exact"
        );
        assert_eq!(
            member_central_record(&source_archive, OPAQUE_MEMBER),
            member_central_record(&output, OPAQUE_MEMBER),
            "opaque central metadata must remain byte exact apart from relocation"
        );
        assert_eq!(
            compression_method(&source_archive, TARGET_MEMBER),
            compression_method(&output, TARGET_MEMBER),
            "the target compression method must be preserved"
        );
    }
}

#[test]
fn replay_512_byte_window_handles_multi_window_source_and_payload() {
    let (source, insertion_offset) = multi_window_source();
    let fragment = multi_window_fragment();
    assert!(
        source.len() > 2 * 512,
        "the source fixture must cross multiple configured replay windows"
    );
    assert!(
        fragment.len() > 3 * 512,
        "the replay fixture must cross multiple configured replay windows"
    );
    let expected = candidate_at(&source, insertion_offset, &fragment);
    let limits = SourcePartSpliceLimits::default().with_max_authored_replay_window_bytes(512);

    for deflated in [false, true] {
        let source_archive = archive_with_source(deflated, &source);
        let package = SourceBackedPackage::from_vec(source_archive.clone())
            .expect("multi-window fixture package must open");
        let provider = ChunkedReplayProvider::new(fragment.clone(), 173);
        let plan = package
            .prepare_source_part_splice_with_replay(
                &pack(TARGET_URI),
                proof_for(&package, &source, insertion_offset, &fragment),
                provider.clone(),
                limits,
            )
            .expect("a 512-byte replay window must prepare");
        let mut output = Vec::new();
        plan.write_to_stream(&mut output)
            .expect("a multi-window replay must publish");

        assert_eq!(target_bytes(&output), expected);
        assert_eq!(provider.open_count(), 3);
        assert_eq!(
            provider.short_final_reads(),
            3,
            "candidate validation, ZIP measurement, and emission must each accept a short final read"
        );
        assert_eq!(
            member_local_record(&source_archive, OPAQUE_MEMBER),
            member_local_record(&output, OPAQUE_MEMBER)
        );
        assert_eq!(
            member_central_record(&source_archive, OPAQUE_MEMBER),
            member_central_record(&output, OPAQUE_MEMBER)
        );
        assert_eq!(
            compression_method(&source_archive, TARGET_MEMBER),
            compression_method(&output, TARGET_MEMBER)
        );
    }
}

#[test]
fn replay_expected_artifact_accepts_exact_store_and_deflate_output() {
    for deflated in [false, true] {
        let (expected, expected_len, expected_hash) = expected_replay_artifact(deflated);
        let (package, _) = open(deflated);
        let plan = replay_plan(
            &package,
            ReplayProvider::new(FRAGMENT, ReplayMode::Exact),
            SourcePartSpliceLimits::default(),
        );
        let mut output = Vec::new();
        plan.write_to_stream_with_expected_artifact(&mut output, expected_len, expected_hash)
            .expect("the exact expected artifact must authorize publication");
        assert_eq!(output, expected);
    }
}

#[test]
fn replay_expected_artifact_mismatch_writes_zero_bytes() {
    for deflated in [false, true] {
        let (_, expected_len, expected_hash) = expected_replay_artifact(deflated);
        for (length, hash) in [
            (expected_len + 1, expected_hash),
            (
                expected_len,
                SourceArtifactFingerprint::from_sha256([0xa5; 32]),
            ),
        ] {
            let (package, _) = open(deflated);
            let plan = replay_plan(
                &package,
                ReplayProvider::new(FRAGMENT, ReplayMode::Exact),
                SourcePartSpliceLimits::default(),
            );
            let mut output = Vec::new();
            let error = plan
                .write_to_stream_with_expected_artifact(&mut output, length, hash)
                .expect_err("an expected artifact mismatch must refuse publication");
            assert!(
                matches!(error, OpcError::SourceArtifactMismatch { .. }),
                "unexpected expected-artifact mismatch: {error:?}"
            );
            assert!(
                output.is_empty(),
                "preflight mismatch must precede sink output"
            );
        }
    }
}

#[test]
fn replay_expected_artifact_preview_failure_keeps_caller_sink_empty() {
    let (_, expected_len, expected_hash) = expected_replay_artifact(false);
    let (package, _) = open(false);
    let plan = replay_plan(
        &package,
        ReplayProvider::new(FRAGMENT, ReplayMode::ReadErrorAtOpen(3)),
        SourcePartSpliceLimits::default(),
    );
    let mut output = Vec::new();
    let error = plan
        .write_to_stream_with_expected_artifact(&mut output, expected_len, expected_hash)
        .expect_err("a preview reader failure must refuse publication");
    assert!(
        output.is_empty(),
        "preview failure must not touch caller output"
    );
    assert!(
        !matches!(error, OpcError::IncompleteOutput { .. }),
        "internal preview progress must not be reported as caller output: {error:?}"
    );
    assert!(
        matches!(error, OpcError::IoError(_)),
        "unexpected preview error: {error:?}"
    );
}

#[test]
fn replay_expected_artifact_charges_one_managed_output_artifact() {
    for deflated in [false, true] {
        let (expected, expected_len, expected_hash) = expected_replay_artifact(deflated);
        let (budget, _cancellation, context) =
            managed_context_with_output(128 * 1024 * 1024, expected_len);
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(OwnedSource::new(archive(deflated))),
            ReadLimits::default(),
            context,
        )
        .expect("managed expected-artifact fixture must open");
        let plan = replay_plan(
            &package,
            ReplayProvider::new(FRAGMENT, ReplayMode::Exact),
            SourcePartSpliceLimits::default(),
        );
        let before = budget.used(Resource::OutputBytes);
        let mut output = Vec::new();
        plan.write_to_stream_with_expected_artifact(&mut output, expected_len, expected_hash)
            .expect("the exact output ceiling must permit one artifact");
        assert_eq!(output, expected);
        assert_eq!(
            budget.used(Resource::OutputBytes) - before,
            expected_len,
            "the preview must not charge output; the actual artifact is charged once"
        );
    }
}

#[test]
fn replay_expected_artifact_change_during_measurement_writes_zero_bytes() {
    let (_, expected_len, expected_hash) = expected_replay_artifact(false);
    let (package, _) = open(false);
    let plan = replay_plan(
        &package,
        ReplayProvider::new(FRAGMENT, ReplayMode::ChangeAtOpen(4)),
        SourcePartSpliceLimits::default(),
    );
    let mut output = Vec::new();
    let error = plan
        .write_to_stream_with_expected_artifact(&mut output, expected_len, expected_hash)
        .expect_err("a changed replay during actual measurement must be refused");
    assert!(
        matches!(
            error,
            OpcError::SourceArtifactMismatch { .. }
                | OpcError::SourceBackedOverlayUnavailable { .. }
                | OpcError::IoError(_)
        ),
        "unexpected measurement mismatch: {error:?}"
    );
    assert!(output.is_empty());
}

#[test]
fn replay_expected_artifact_change_during_emission_is_incomplete() {
    let (_, expected_len, expected_hash) = expected_replay_artifact(false);
    let (package, _) = open(false);
    let plan = replay_plan(
        &package,
        ReplayProvider::new(FRAGMENT, ReplayMode::ChangeAtOpen(5)),
        SourcePartSpliceLimits::default(),
    );
    let mut output = Vec::new();
    let error = plan
        .write_to_stream_with_expected_artifact(&mut output, expected_len, expected_hash)
        .expect_err("a changed replay during emission must report incomplete output");
    assert!(
        matches!(error, OpcError::IncompleteOutput { .. }),
        "unexpected emission mismatch: {error:?}"
    );
    assert!(
        !output.is_empty(),
        "emission mismatch must retain accepted prefix"
    );
}

#[test]
fn replay_expected_artifact_noop_releases_preview_memory_and_charges_once() {
    let source_archive = archive(false);
    let expected_len = source_archive.len() as u64;
    let expected_hash = SourceArtifactFingerprint::from_sha256(digest(&source_archive));
    let (budget, _cancellation, context) =
        managed_context_with_output(128 * 1024 * 1024, expected_len);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(source_archive.clone())),
        ReadLimits::default(),
        context,
    )
    .expect("managed no-op fixture must open");
    let provider = ReplayProvider::new(&[], ReplayMode::Exact);
    let plan = package
        .prepare_source_part_splice_with_replay(
            &pack(TARGET_URI),
            proof_for(&package, SOURCE, INSERTION_OFFSET, &[]),
            provider,
            SourcePartSpliceLimits::default(),
        )
        .expect("empty replay insertion must prepare as an exact no-op");
    let baseline_memory = budget.used(Resource::Memory);
    let peak_memory = Arc::new(AtomicUsize::new(
        usize::try_from(baseline_memory).expect("memory baseline must fit usize"),
    ));
    let mut sink = MemoryObservingSink {
        bytes: Vec::new(),
        budget: budget.clone(),
        peak_memory: Arc::clone(&peak_memory),
    };
    plan.write_to_stream_with_expected_artifact(&mut sink, expected_len, expected_hash)
        .expect("the exact no-op artifact must publish");
    assert_eq!(sink.bytes, source_archive);
    assert_eq!(
        budget.used(Resource::OutputBytes),
        expected_len,
        "the preview must not charge output for the no-op artifact"
    );
    assert_eq!(budget.used(Resource::Memory), baseline_memory);
    assert!(
        peak_memory.load(Ordering::Acquire) >= baseline_memory as usize + 64 * 1024,
        "each managed no-op pass must own its bounded 64 KiB copy workspace"
    );
}

struct MemoryObservingSink {
    bytes: Vec<u8>,
    budget: Budget,
    peak_memory: Arc<AtomicUsize>,
}

impl Write for MemoryObservingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        let used = usize::try_from(self.budget.used(Resource::Memory))
            .map_err(|_| io::Error::other("memory usage does not fit usize"))?;
        self.peak_memory.fetch_max(used, Ordering::AcqRel);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn compression_method(bytes: &[u8], wanted: &str) -> usize {
    let eocd = bytes.len() - 22;
    let mut central = u32_at(bytes, eocd + 16);
    for _ in 0..u16_at(bytes, eocd + 10) {
        let name_len = u16_at(bytes, central + 28);
        let extra_len = u16_at(bytes, central + 30);
        let comment_len = u16_at(bytes, central + 32);
        let end = central + 46 + name_len + extra_len + comment_len;
        if &bytes[central + 46..central + 46 + name_len] == wanted.as_bytes() {
            return u16_at(bytes, central + 10);
        }
        central = end;
    }
    panic!("compression member {wanted} is missing")
}

#[test]
fn replay_opens_fresh_readers_for_candidate_measurement_and_emission() {
    let (package, _) = open(false);
    let provider = ReplayProvider::new(FRAGMENT, ReplayMode::Exact);
    let plan = replay_plan(
        &package,
        provider.clone(),
        SourcePartSpliceLimits::default(),
    );
    assert_eq!(
        provider.open_count(),
        1,
        "candidate verification needs one reader"
    );
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("replay publication must succeed");
    assert_eq!(
        provider.open_count(),
        3,
        "candidate, ZIP measurement, and ZIP emission must each open a fresh reader"
    );
    assert_eq!(target_bytes(&output), candidate(SOURCE, FRAGMENT));
}

#[test]
fn replay_rejects_truncated_extra_and_hash_changed_readers_before_output() {
    for mode in [
        ReplayMode::Truncated,
        ReplayMode::Extra,
        ReplayMode::HashChanged,
    ] {
        let (package, _) = open(false);
        let provider = ReplayProvider::new(FRAGMENT, mode);
        let error = package
            .prepare_source_part_splice_with_replay(
                &pack(TARGET_URI),
                proof(&package, FRAGMENT),
                provider,
                SourcePartSpliceLimits::default(),
            )
            .expect_err("an unauthenticated replay reader must be refused");
        assert!(
            matches!(
                error,
                OpcError::IoError(_) | OpcError::SourceBackedOverlayUnavailable { .. }
            ),
            "unexpected replay refusal for {mode:?}: {error:?}"
        );
    }
}

#[test]
fn mutable_replay_proof_cannot_expand_a_bounded_candidate_reader() {
    let (package, _) = open(false);
    let provider = MutableProofProvider::new();
    let error = package
        .prepare_source_part_splice_with_replay(
            &pack(TARGET_URI),
            proof(&package, FRAGMENT),
            provider.clone(),
            SourcePartSpliceLimits::default(),
        )
        .expect_err("a mutable provider proof must not expand the replay window");
    assert!(
        matches!(
            error,
            OpcError::IoError(_)
                | OpcError::SourceBackedOverlayUnavailable { .. }
                | OpcError::XmlPublication { .. }
        ),
        "unexpected mutable-proof refusal: {error:?}"
    );
    assert!(
        provider.bytes_read() <= FRAGMENT.len() + 1,
        "a later u64::MAX proof must not make the reader consume the excessive payload"
    );
}

#[test]
fn replay_provider_open_and_mid_read_errors_are_typed() {
    for mode in [ReplayMode::OpenError, ReplayMode::ReadError] {
        let (package, _) = open(false);
        let provider = ReplayProvider::new(FRAGMENT, mode);
        let error = package
            .prepare_source_part_splice_with_replay(
                &pack(TARGET_URI),
                proof(&package, FRAGMENT),
                provider,
                SourcePartSpliceLimits::default(),
            )
            .expect_err("provider I/O failure must refuse preparation");
        assert!(
            matches!(error, OpcError::IoError(_)),
            "unexpected error: {error:?}"
        );
    }
}

#[test]
fn replay_reader_cancellation_is_reported_as_typed_cancelled() {
    let (_budget, cancellation, context) = managed_context(64 * 1024 * 1024);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(archive(false))),
        ReadLimits::default(),
        context,
    )
    .expect("managed replay fixture must open");
    let provider = CancellingReplayProvider::new(cancellation);
    let error = package
        .prepare_source_part_splice_with_replay(
            &pack(TARGET_URI),
            proof(&package, FRAGMENT),
            provider,
            SourcePartSpliceLimits::default(),
        )
        .expect_err("cancellation from a replay reader must stop preparation");
    assert!(
        matches!(error, OpcError::Cancelled),
        "unexpected error: {error:?}"
    );
}

#[test]
fn replay_zero_window_limit_is_rejected_before_retaining_the_plan() {
    let (package, _) = open(false);
    let provider = ReplayProvider::new(FRAGMENT, ReplayMode::Exact);
    let error = package
        .prepare_source_part_splice_with_replay(
            &pack(TARGET_URI),
            proof(&package, FRAGMENT),
            provider,
            SourcePartSpliceLimits::default().with_max_authored_replay_window_bytes(0),
        )
        .expect_err("a zero replay window limit must be refused");
    assert!(matches!(
        error,
        OpcError::InvalidSourcePartSpliceLimit {
            resource: SpliceResource::AuthoredReplayMemoryBytes,
            value: 0,
        }
    ));
}

#[test]
fn replay_provider_change_on_measurement_stops_before_sink_output() {
    let (package, _) = open(false);
    let provider = ReplayProvider::new(FRAGMENT, ReplayMode::ChangeAfterFirstOpen);
    let plan = replay_plan(
        &package,
        provider.clone(),
        SourcePartSpliceLimits::default(),
    );
    assert_eq!(provider.open_count(), 1);
    let mut output = Vec::new();
    let error = plan
        .write_to_stream(&mut output)
        .expect_err("provider bytes changing for measurement must be refused");
    assert!(
        output.is_empty(),
        "measurement failure must precede ZIP output"
    );
    assert!(
        matches!(
            error,
            OpcError::SourceBackedOverlayUnavailable { .. } | OpcError::IoError(_)
        ),
        "unexpected provider-change error: {error:?}"
    );
    assert_eq!(provider.open_count(), 2);
}

fn contains_memory_limit(error: &OpcError) -> bool {
    match error {
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Memory =>
        {
            true
        },
        OpcError::IncompleteOutput { source, .. } => contains_memory_limit(source),
        _ => false,
    }
}

#[test]
fn concurrent_replay_publications_charge_one_window_per_live_reader() {
    let source_archive = archive(false);
    let limits = SourcePartSpliceLimits::default();
    let (budget, _cancellation, context) = managed_context(128 * 1024 * 1024);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(source_archive.clone())),
        ReadLimits::default(),
        context,
    )
    .expect("managed replay fixture must open");
    let provider = BlockingReplayProvider::new(FRAGMENT);
    let used_before_plans = budget.used(Resource::Memory);
    let first_plan = replay_plan(&package, provider.clone(), limits);
    let used_after_first_plan = budget.used(Resource::Memory);
    let second_plan = replay_plan(&package, provider.clone(), limits);
    let used_after_second_plan = budget.used(Resource::Memory);
    assert_eq!(
        used_after_first_plan, used_before_plans,
        "a prepared replay plan must retain no parser or replay-window lease"
    );
    assert_eq!(
        used_after_second_plan, used_after_first_plan,
        "a second prepared replay plan must retain no parser or replay-window lease"
    );

    // Each live publication owns its parser and publication state. Leave
    // enough managed memory for two such envelopes plus one 64 KiB replay
    // window. The second publication must therefore reach, and fail at, its
    // independent authored-window reservation while the first callback is
    // blocked.
    let publication_envelope = publication_memory_envelope(&source_archive, limits);
    let replay_window = soapberry_zip::RECOMMENDED_BUFFER_SIZE as u64;
    let required = publication_envelope
        .checked_mul(2)
        .and_then(|value| value.checked_add(replay_window))
        .expect("concurrency memory envelope must fit u64");
    let used = budget.used(Resource::Memory);
    let guard_amount = budget
        .limit(Resource::Memory)
        .checked_sub(used)
        .and_then(|available| available.checked_sub(required))
        .expect("managed budget must have room for the concurrency fixture");
    let guard = budget
        .reserve(Resource::Memory, guard_amount)
        .expect("memory guard must reserve the intended remaining envelope");

    thread::scope(|scope| {
        let first = scope.spawn(move || {
            let mut output = Vec::new();
            let result = first_plan.write_to_stream(&mut output);
            (result, output)
        });
        provider.wait_until_first_publication_reader();

        let second = scope.spawn(move || {
            let mut output = Vec::new();
            let result = second_plan.write_to_stream(&mut output);
            (result, output)
        });
        let (second_result, second_output) = second
            .join()
            .expect("second publication thread must not panic");
        assert!(
            second_output.is_empty(),
            "replay-window refusal must occur before second sink output"
        );
        let second_error = second_result.expect_err("second live replay window must be refused");
        assert!(
            contains_memory_limit(&second_error),
            "unexpected concurrent budget error: {second_error:?}"
        );

        provider.release_first_publication_reader();
        let (first_result, first_output) = first
            .join()
            .expect("first publication thread must not panic");
        first_result.expect("the first publication must complete after release");
        assert_eq!(target_bytes(&first_output), candidate(SOURCE, FRAGMENT));
    });
    drop(guard);
}

struct PrefixFailSink {
    bytes: Vec<u8>,
    remaining: usize,
}

impl PrefixFailSink {
    fn new(remaining: usize) -> Self {
        Self {
            bytes: Vec::new(),
            remaining,
        }
    }
}

impl Write for PrefixFailSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "short fixture sink",
            ));
        }
        let accepted = self.remaining.min(bytes.len());
        self.bytes.extend_from_slice(&bytes[..accepted]);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn replay_short_sink_reports_exact_partial_output() {
    let (package, _) = open(true);
    let provider = ReplayProvider::new(FRAGMENT, ReplayMode::Exact);
    let plan = replay_plan(&package, provider, SourcePartSpliceLimits::default());
    let mut sink = PrefixFailSink::new(19);
    let error = plan
        .write_to_stream(&mut sink)
        .expect_err("short sink must fail the replay publication");
    let written = match error {
        OpcError::IncompleteOutput { written, .. } => written,
        other => panic!("short sink must report IncompleteOutput, got {other:?}"),
    };
    assert_eq!(written, sink.bytes.len() as u64);
    assert_eq!(written, 19);
}

#[test]
fn replay_reader_error_during_publication_reports_accepted_output() {
    for deflated in [false, true] {
        let (package, _) = open(deflated);
        // Preparation and measurement succeed. The publishing reader returns
        // a short prefix, then fails while that prefix can still be buffered.
        let provider = ReplayProvider::new(FRAGMENT, ReplayMode::ReadErrorAtOpen(3));
        let plan = replay_plan(
            &package,
            Arc::clone(&provider),
            SourcePartSpliceLimits::default(),
        );
        let mut output = Vec::new();
        let error = plan
            .write_to_stream(&mut output)
            .expect_err("reader failure must refuse an incomplete publication");
        match error {
            OpcError::IncompleteOutput { written, source } => {
                assert_eq!(written, output.len() as u64);
                assert!(written > 0, "ZIP publication must have accepted a prefix");
                assert!(
                    matches!(*source, OpcError::IoError(ref error) if error.kind() == io::ErrorKind::BrokenPipe),
                    "the original replay I/O error must remain authoritative: {source:?}"
                );
            },
            other => panic!("expected incomplete publication, got {other:?}"),
        }
        assert_eq!(provider.open_count(), 3);
    }
}

#[derive(Debug)]
struct MutableSource {
    bytes: RwLock<Vec<u8>>,
    revision: AtomicUsize,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes: RwLock::new(bytes),
            revision: AtomicUsize::new(0),
        })
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::AcqRel);
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.read().expect("source lock must work").len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let bytes = self.bytes.read().expect("source lock must work");
        let start = usize::try_from(offset).map_err(|_| io::Error::other("offset overflow"))?;
        if start >= bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(bytes.len() - start);
        output[..count].copy_from_slice(&bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x5245_504c_4159,
            self.revision.load(Ordering::Acquire) as u64,
        ))
    }
}

struct MutatingSink {
    source: Arc<MutableSource>,
    bytes: Vec<u8>,
    mutated: bool,
}

impl Write for MutatingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        if !self.mutated {
            self.source.bump_revision();
            self.mutated = true;
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn contains_source_change(error: &OpcError) -> bool {
    match error {
        OpcError::SourceChanged { .. } => true,
        OpcError::IncompleteOutput { source, .. } => contains_source_change(source),
        _ => false,
    }
}

#[test]
fn replay_source_change_after_sink_progress_is_incomplete() {
    let source_bytes = archive(false);
    let source = MutableSource::new(source_bytes);
    let package = SourceBackedPackage::from_read_at(source.clone())
        .expect("mutable replay fixture must open");
    let provider = ReplayProvider::new(FRAGMENT, ReplayMode::Exact);
    let plan = replay_plan(&package, provider, SourcePartSpliceLimits::default());
    let mut sink = MutatingSink {
        source,
        bytes: Vec::new(),
        mutated: false,
    };
    let error = plan
        .write_to_stream(&mut sink)
        .expect_err("source mutation during replay publication must fail");
    assert!(
        contains_source_change(&error),
        "unexpected error: {error:?}"
    );
    assert!(!sink.bytes.is_empty());
    assert!(matches!(error, OpcError::IncompleteOutput { written, .. } if written > 0));
}
