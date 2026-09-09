#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "focused stream fixtures make invalid test data construction explicit"
)]

//! Public-boundary coverage for replayable source-backed DOCX paragraph streams.
//!
//! These tests intentionally construct a small OPC package instead of using
//! private scanner or ZIP-preservation helpers.  The source cursor owns the
//! event strings and returns borrowed `&str` values, which keeps the fixture
//! close to the contract exercised by callers.

use std::fmt;
use std::io::{self, Cursor, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicU64, AtomicUsize, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, CancellationToken, ExecutionContext, ExecutionLimits,
    Limits as CoreLimits, ReadAt, Resource, SourceVersion,
};
use litchi_docx::source_backed::tail_append::Error as TailAppendError;
use litchi_docx::source_backed::tail_append_stream::patch::{
    AuthoredReplayResolver, ExactInverseAuthorization, OriginalArtifactError,
    OriginalArtifactProvider, OriginalArtifactReference, ParagraphStreamPatch, PatchError,
    PatchLimits, ReplayResolverError, apply_exact_inverse, artifact_proof,
};
use litchi_docx::source_backed::tail_append_stream::{
    AuthoredReplayError, AuthoredReplayHandle, AuthoredReplayReader, AuthoredReplayReference,
    AuthoredReplayStore, AuthoredStreamProof, Error as StreamError, MemoryReplayHandle,
    MemoryReplayStore, ParagraphCursor, ParagraphEventSink, ParagraphStreamLimits,
    PlainParagraphEvent, ReplayableParagraphSource,
};
use litchi_docx::source_backed::{self, TailAppendLimits};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{PackURI, SourceBackedPackage};
use sha2::{Digest as _, Sha256};
use soapberry_zip::office::StreamingArchiveWriter;

const WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const MAIN: &str = "word/document.xml";
const MAIN_URI: &str = "/word/document.xml";
const OPAQUE: &str = "word/opaque.bin";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Compression {
    Store,
    Deflate,
}

#[derive(Clone, Debug, PartialEq, Eq)]
enum EventData {
    Start,
    Text(String),
    End,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CursorError;

impl fmt::Display for CursorError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("test cursor failure")
    }
}

impl std::error::Error for CursorError {}

#[derive(Debug)]
struct EventCursor<'source> {
    events: &'source [EventData],
    position: usize,
}

impl ParagraphCursor for EventCursor<'_> {
    type Error = CursorError;

    fn next<'event>(&'event mut self) -> Result<Option<PlainParagraphEvent<'event>>, Self::Error> {
        let event = self.events.get(self.position);
        self.position = self.position.saturating_add(1);
        Ok(event.map(|event| match event {
            EventData::Start => PlainParagraphEvent::ParagraphStart,
            EventData::Text(text) => PlainParagraphEvent::TextChunk(text.as_str()),
            EventData::End => PlainParagraphEvent::ParagraphEnd,
        }))
    }
}

#[derive(Debug)]
struct EventSource {
    events: Vec<EventData>,
    alternate: Option<Vec<EventData>>,
    opens: Arc<AtomicUsize>,
    durable: Option<AuthoredReplayReference>,
}

impl EventSource {
    fn new(events: Vec<EventData>) -> Self {
        Self::counted(events).0
    }

    fn counted(events: Vec<EventData>) -> (Self, Arc<AtomicUsize>) {
        let opens = Arc::new(AtomicUsize::new(0));
        (
            Self {
                events,
                alternate: None,
                opens: Arc::clone(&opens),
                durable: None,
            },
            opens,
        )
    }

    fn changing(events: Vec<EventData>, alternate: Vec<EventData>) -> Self {
        Self {
            events,
            alternate: Some(alternate),
            opens: Arc::new(AtomicUsize::new(0)),
            durable: None,
        }
    }
}

impl ReplayableParagraphSource for EventSource {
    type Error = CursorError;
    type Cursor<'source> = EventCursor<'source>;

    fn open<'source>(&'source self) -> Result<Self::Cursor<'source>, Self::Error> {
        let opening = self.opens.fetch_add(1, Ordering::AcqRel);
        let events = match (opening, self.alternate.as_deref()) {
            (opening, Some(alternate)) if opening > 0 => alternate,
            _ => self.events.as_slice(),
        };
        Ok(EventCursor {
            events,
            position: 0,
        })
    }

    fn durable_reference(&self) -> Option<AuthoredReplayReference> {
        self.durable.clone()
    }
}

#[derive(Debug)]
struct FailingSource;

impl ReplayableParagraphSource for FailingSource {
    type Error = CursorError;
    type Cursor<'source> = EventCursor<'source>;

    fn open<'source>(&'source self) -> Result<Self::Cursor<'source>, Self::Error> {
        Err(CursorError)
    }
}

fn events_for(texts: &[&str]) -> Vec<EventData> {
    texts
        .iter()
        .flat_map(|text| {
            [
                EventData::Start,
                EventData::Text((*text).to_owned()),
                EventData::End,
            ]
        })
        .collect()
}

fn events_for_owned(texts: impl IntoIterator<Item = String>) -> Vec<EventData> {
    texts
        .into_iter()
        .flat_map(|text| [EventData::Start, EventData::Text(text), EventData::End])
        .collect()
}

fn source_xml(strict: bool, padding: usize) -> Vec<u8> {
    let (prefix, namespace) = if strict {
        ("s", STRICT_WORD)
    } else {
        ("w", WORD)
    };
    let section = String::from_utf8(section_xml(strict)).expect("section fixture is UTF-8");
    let padding_token = format!("<{prefix}:t>x</{prefix}:t>");
    let padding = padding_token.repeat(padding / padding_token.len());
    format!(
        r#"<?xml version="1.0" encoding="UTF-8"?><{prefix}:document xmlns:{prefix}="{namespace}"><{prefix}:body><{prefix}:p><{prefix}:r><{prefix}:t>seed</{prefix}:t>{padding}</{prefix}:r></{prefix}:p>{section}</{prefix}:body></{prefix}:document>"#,
    )
    .into_bytes()
}

fn section_xml(strict: bool) -> Vec<u8> {
    let (prefix, namespace) = if strict {
        ("s", STRICT_WORD)
    } else {
        ("w", WORD)
    };
    format!(
        r#"<{prefix}:sectPr xmlns:{prefix}="{namespace}" xmlns:v="urn:vendor-section" v:keep="yes"><v:opaque a="&quot;"/><{prefix}:pgSz {prefix}:w="11906" {prefix}:h="16838"/></{prefix}:sectPr>"#
    )
    .into_bytes()
}

fn source_xml_without_section(strict: bool) -> Vec<u8> {
    let (prefix, namespace) = if strict {
        ("s", STRICT_WORD)
    } else {
        ("w", WORD)
    };
    format!(
        r#"<?xml version="1.0"?><{prefix}:document xmlns:{prefix}="{namespace}"><{prefix}:body><{prefix}:p><{prefix}:r><{prefix}:t>seed</{prefix}:t></{prefix}:r></{prefix}:p></{prefix}:body></{prefix}:document>"#
    )
    .into_bytes()
}

fn shadowed_source_xml() -> Vec<u8> {
    format!(
        r#"<document xmlns="{WORD}" xmlns:w="urn:root-shadow"><body><p><r><t>seed</t></r></p><sectPr xmlns="{WORD}" xmlns:w="urn:vendor-shadow"><w:opaque/></sectPr></body></document>"#
    )
    .into_bytes()
}

fn archive(main: &[u8], strict: bool, compression: Compression) -> Vec<u8> {
    let office_document = if strict {
        rt::STRICT_OFFICE_DOCUMENT
    } else {
        rt::OFFICE_DOCUMENT
    };
    let content_types = format!(
        r#"<Types xmlns="{content_types}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/{MAIN}" ContentType="{document}"/></Types>"#,
        content_types = "http://schemas.openxmlformats.org/package/2006/content-types",
        document = ct::WML_DOCUMENT_MAIN,
    );
    let relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS}"><Relationship Id="rIdDocument" Type="{office_document}" Target="{MAIN}"/></Relationships>"#
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content-types fixture must be writable");
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .expect("relationship fixture must be writable");
    writer
        .write_stored(OPAQUE, b"opaque physical payload\0\xff\x01")
        .expect("opaque fixture must be writable");
    match compression {
        Compression::Store => writer
            .write_stored(MAIN, main)
            .expect("stored main fixture must be writable"),
        Compression::Deflate => writer
            .write_deflated_sized(MAIN, main)
            .expect("deflated main fixture must be writable"),
    }
    writer
        .finish_to_bytes()
        .expect("archive fixture must finish")
}

fn open_package(bytes: Vec<u8>) -> source_backed::Package {
    source_backed::Package::from_reader(Cursor::new(bytes))
        .expect("source-backed DOCX fixture must open")
}

fn limits(source_xml_bytes: usize) -> ParagraphStreamLimits {
    let source = TailAppendLimits {
        max_source_xml_bytes: source_xml_bytes as u64 + 256 * 1024,
        max_text_bytes: 64 * 1024 * 1024,
        max_fragment_bytes: 16 * 1024 * 1024,
        max_candidate_xml_bytes: source_xml_bytes as u64 + 16 * 1024 * 1024,
        max_events: 65_536,
        max_depth: 32,
        max_paragraphs: 65_536,
        max_settings_xml_bytes: 64 * 1024,
        max_workspace_bytes: 16 * 1024 * 1024,
        max_output_bytes: 32 * 1024 * 1024,
        max_token_bytes: 4 * 1024,
    };

    let mut stream = ParagraphStreamLimits::new(source);
    stream.max_authored_paragraphs = 16_384;
    stream.max_authored_events = 100_000;
    stream.max_authored_chunk_bytes = 8 * 1024;
    stream.max_authored_text_bytes = 16 * 1024 * 1024;
    stream.max_authored_xml_bytes = 16 * 1024 * 1024;
    stream.max_replay_bytes = 16 * 1024 * 1024;
    stream.max_replay_window_bytes = 64 * 1024;
    stream.max_patch_bytes = 64 * 1024;
    stream
}

fn managed_context(memory: u64, output: u64) -> (Budget, ExecutionContext) {
    let (budget, _cancellation_source, context) =
        managed_context_with_objects(memory, output, u64::MAX);
    (budget, context)
}

fn managed_context_with_objects(
    memory: u64,
    output: u64,
    objects: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "docx-tail-append-stream-managed-test",
        CoreLimits::new(memory, u64::MAX, output, objects, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory.max(1)).expect("managed memory limit must be nonzero"),
        0,
    )
    .expect("managed execution limits must be valid");
    (
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn managed_package(archive: Vec<u8>, memory: u64, output: u64) -> (Budget, source_backed::Package) {
    let (budget, context) = managed_context(memory, output);
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(VersionedArchiveSource::new(archive)),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed source-backed DOCX fixture must open");
    (budget, package)
}

fn managed_package_with_cancellation(
    archive: Vec<u8>,
    memory: u64,
    output: u64,
) -> (Budget, CancellationSource, source_backed::Package) {
    let (budget, cancellation_source, context) =
        managed_context_with_objects(memory, output, u64::MAX);
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(VersionedArchiveSource::new(archive)),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed source-backed DOCX fixture must open");
    (budget, cancellation_source, package)
}

fn main_bytes(archive: &[u8]) -> Vec<u8> {
    SourceBackedPackage::from_vec(archive.to_vec())
        .expect("published archive must reopen")
        .part(&PackURI::new(MAIN_URI).expect("main URI must be valid"))
        .expect("main part must exist")
        .data()
        .expect("main part must decode")
        .as_bytes()
        .to_vec()
}

fn semantic_texts(archive: &[u8]) -> Vec<String> {
    let package = open_package(archive.to_vec());
    package
        .document_snapshot()
        .expect("published document must reopen")
        .paragraphs()
        .iter()
        .map(|paragraph| paragraph.text().expect("paragraph text must decode"))
        .collect()
}

#[derive(Debug)]
struct PrefixFailSink {
    bytes: Vec<u8>,
    remaining: usize,
}

impl PrefixFailSink {
    fn after(bytes: usize) -> Self {
        Self {
            bytes: Vec::new(),
            remaining: bytes,
        }
    }
}

impl Write for PrefixFailSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(io::ErrorKind::BrokenPipe, "sink stopped"));
        }
        let accepted = bytes.len().min(self.remaining);
        self.bytes.extend_from_slice(&bytes[..accepted]);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct VersionedArchiveSource {
    bytes: Mutex<Vec<u8>>,
    revision: AtomicU64,
}

impl VersionedArchiveSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Mutex::new(bytes),
            revision: AtomicU64::new(0),
        }
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::AcqRel);
    }

    fn mutate_opaque_payload(&self) {
        let marker = b"opaque physical payload\0\xff\x01";
        let mut bytes = self
            .bytes
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        let offset = bytes
            .windows(marker.len())
            .position(|window| window == marker)
            .expect("opaque payload marker must be present");
        bytes[offset] ^= 1;
        self.bump();
    }
}

impl ReadAt for VersionedArchiveSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self
            .bytes
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        let bytes = self
            .bytes
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner());
        if offset >= bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(bytes.len() - offset);
        output[..count].copy_from_slice(&bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            48_404,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

#[derive(Debug)]
struct TestOriginalProvider {
    bytes: Option<Vec<u8>>,
}

impl OriginalArtifactProvider for TestOriginalProvider {
    fn open(
        &self,
        _reference: &OriginalArtifactReference,
    ) -> Result<litchi_opc::SourceArtifact, OriginalArtifactError> {
        let bytes = self.bytes.clone().ok_or(OriginalArtifactError::Missing)?;
        let package = SourceBackedPackage::from_vec(bytes)
            .map_err(|error| OriginalArtifactError::Provider(Box::new(error)))?;
        Ok(package.source_artifact())
    }
}

#[derive(Debug)]
struct CountingOriginalProvider {
    bytes: Vec<u8>,
    opens: Arc<AtomicUsize>,
}

impl OriginalArtifactProvider for CountingOriginalProvider {
    fn open(
        &self,
        _reference: &OriginalArtifactReference,
    ) -> Result<litchi_opc::SourceArtifact, OriginalArtifactError> {
        self.opens.fetch_add(1, Ordering::AcqRel);
        let package = SourceBackedPackage::from_vec(self.bytes.clone())
            .map_err(|error| OriginalArtifactError::Provider(Box::new(error)))?;
        Ok(package.source_artifact())
    }
}

#[derive(Debug)]
struct MutatingOriginalProvider {
    source: Arc<VersionedArchiveSource>,
}

impl OriginalArtifactProvider for MutatingOriginalProvider {
    fn open(
        &self,
        _reference: &OriginalArtifactReference,
    ) -> Result<litchi_opc::SourceArtifact, OriginalArtifactError> {
        self.source.mutate_opaque_payload();
        Err(OriginalArtifactError::Missing)
    }
}

fn mutate_opaque_payload(archive: &[u8]) -> Vec<u8> {
    let mut mutated = archive.to_vec();
    let marker = b"opaque physical payload\0\xff\x01";
    let offset = mutated
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("opaque payload marker must be present");
    mutated[offset] ^= 1;
    mutated
}

#[derive(Debug)]
struct CapturingReplayStore {
    inner: Option<MemoryReplayStore>,
    captured: Arc<Mutex<Option<MemoryReplayHandle>>>,
}

impl AuthoredReplayStore for CapturingReplayStore {
    type Handle = MemoryReplayHandle;

    fn prepare_for_operation(
        &mut self,
        limits: ParagraphStreamLimits,
        context: Option<&ExecutionContext>,
        cancellation: Option<&CancellationToken>,
    ) -> Result<(), AuthoredReplayError> {
        self.inner
            .as_mut()
            .expect("capturing store must remain available")
            .prepare_for_operation(limits, context, cancellation)
    }

    fn append(&mut self, chunk: &[u8]) -> Result<(), AuthoredReplayError> {
        self.inner
            .as_mut()
            .expect("capturing store must remain available")
            .append(chunk)
    }

    fn finish(mut self, proof: AuthoredStreamProof) -> Result<Self::Handle, AuthoredReplayError> {
        let handle = self
            .inner
            .take()
            .expect("capturing store must remain available")
            .finish(proof)?;
        *self
            .captured
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner()) = Some(handle.clone());
        Ok(handle)
    }
}

#[derive(Debug)]
struct ForgedReplayStore {
    accepted: Option<MemoryReplayStore>,
    replacement: Option<MemoryReplayHandle>,
}

impl AuthoredReplayStore for ForgedReplayStore {
    type Handle = MemoryReplayHandle;

    fn append(&mut self, chunk: &[u8]) -> Result<(), AuthoredReplayError> {
        self.accepted
            .as_mut()
            .expect("forged store must remain available while producing")
            .append(chunk)
    }

    fn finish(mut self, _proof: AuthoredStreamProof) -> Result<Self::Handle, AuthoredReplayError> {
        let _accepted = self
            .accepted
            .take()
            .expect("forged store must have accepted the producer bytes");
        Ok(self
            .replacement
            .take()
            .expect("forged store must have a presealed replacement"))
    }
}

fn forged_payload() -> (Vec<u8>, AuthoredStreamProof) {
    let text = "forged";
    let xml =
        format!(r#"<w:p xmlns:w="{WORD}"><w:r><w:t xml:space="preserve">{text}</w:t></w:r></w:p>"#)
            .into_bytes();
    let mut event_hash = Sha256::new();
    event_hash.update([0]);
    event_hash.update([1]);
    event_hash.update((text.len() as u64).to_le_bytes());
    event_hash.update(text.as_bytes());
    event_hash.update([2]);
    let proof = AuthoredStreamProof {
        strict_namespace: false,
        paragraph_count: 1,
        event_count: 3,
        text_bytes: text.len() as u64,
        encoded_xml_bytes: xml.len() as u64,
        event_sha256: event_hash.finalize().into(),
        encoded_sha256: Sha256::digest(&xml).into(),
    };
    (xml, proof)
}

fn presealed_forged_handle(maximum: u64) -> MemoryReplayHandle {
    let (xml, proof) = forged_payload();
    let mut store = MemoryReplayStore::new(maximum).expect("forged store limit is finite");
    store.append(&xml).expect("forged payload fits its store");
    store
        .finish(proof)
        .expect("forged payload must seal legally")
}

fn managed_forged_handle(
    maximum: u64,
    limits: ParagraphStreamLimits,
    context: &ExecutionContext,
) -> MemoryReplayHandle {
    let (xml, proof) = forged_payload();
    let mut store = MemoryReplayStore::new(maximum).expect("managed store limit is finite");
    store
        .prepare_for_operation(limits, Some(context), None)
        .expect("managed store must reserve its retained owner");
    store.append(&xml).expect("managed forged payload must fit");
    store
        .finish(proof)
        .expect("managed forged payload must seal legally")
}

#[derive(Debug)]
struct TestReplayResolver {
    expected: AuthoredReplayReference,
    handle: Arc<MemoryReplayHandle>,
}

impl AuthoredReplayResolver for TestReplayResolver {
    fn resolve(
        &self,
        reference: &AuthoredReplayReference,
    ) -> Result<Arc<dyn AuthoredReplayHandle>, ReplayResolverError> {
        if reference.as_bytes() != self.expected.as_bytes() {
            return Err(ReplayResolverError::Missing);
        }
        let handle: Arc<dyn AuthoredReplayHandle> = self.handle.clone();
        Ok(handle)
    }
}

#[derive(Debug)]
struct CountingReplayHandle {
    inner: Arc<MemoryReplayHandle>,
    opens: Arc<AtomicUsize>,
}

impl AuthoredReplayHandle for CountingReplayHandle {
    fn proof(&self) -> AuthoredStreamProof {
        self.inner.proof()
    }

    fn open(&self) -> Result<Box<dyn AuthoredReplayReader + '_>, AuthoredReplayError> {
        self.opens.fetch_add(1, Ordering::AcqRel);
        self.inner.open()
    }

    fn durable_reference(&self) -> Option<AuthoredReplayReference> {
        self.inner.durable_reference()
    }
}

#[derive(Debug)]
struct CancellingReplayResolver {
    expected: AuthoredReplayReference,
    handle: Arc<CountingReplayHandle>,
    cancellation: CancellationSource,
}

impl AuthoredReplayResolver for CancellingReplayResolver {
    fn resolve(
        &self,
        reference: &AuthoredReplayReference,
    ) -> Result<Arc<dyn AuthoredReplayHandle>, ReplayResolverError> {
        if reference.as_bytes() != self.expected.as_bytes() {
            return Err(ReplayResolverError::Missing);
        }
        self.cancellation.cancel();
        let handle: Arc<dyn AuthoredReplayHandle> = self.handle.clone();
        Ok(handle)
    }
}

#[derive(Debug)]
struct MutatingMissingReplayResolver {
    source: Arc<VersionedArchiveSource>,
}

impl AuthoredReplayResolver for MutatingMissingReplayResolver {
    fn resolve(
        &self,
        _reference: &AuthoredReplayReference,
    ) -> Result<Arc<dyn AuthoredReplayHandle>, ReplayResolverError> {
        self.source.mutate_opaque_payload();
        Err(ReplayResolverError::Missing)
    }
}

#[test]
fn replayable_stream_escapes_borrowed_chunks_and_preserves_section_and_inverse() {
    let authored_events = vec![
        EventData::Start,
        EventData::Text("<&>".to_owned()),
        EventData::Text(String::new()),
        EventData::Text("\"'“”🙂".to_owned()),
        EventData::End,
        EventData::Start,
        EventData::Text("second".to_owned()),
        EventData::End,
    ];

    for strict in [false, true] {
        for compression in [Compression::Store, Compression::Deflate] {
            let source_xml = source_xml(strict, 0);
            let source_archive = archive(&source_xml, strict, compression);
            let section = section_xml(strict);
            let package = open_package(source_archive.clone());
            let plan = package
                .tail_append_plain_paragraphs(
                    EventSource::new(authored_events.clone()),
                    limits(source_xml.len()),
                )
                .prepare()
                .expect("replayable stream must prepare");
            let proof = plan.authored_proof();
            assert_eq!(proof.paragraph_count, 2);
            assert_eq!(proof.event_count, 8);
            assert_eq!(proof.text_bytes, "<&>\"'“”🙂second".len() as u64);
            assert!(proof.encoded_xml_bytes > proof.text_bytes);
            assert_eq!(proof.strict_namespace, strict);

            let mut output = Vec::new();
            let publication = plan
                .write_to_stream(&mut output)
                .expect("replayable stream must publish");
            assert_ne!(output, source_archive);
            assert_eq!(
                semantic_texts(&output),
                vec!["seed", "<&>\"'“”🙂", "second"]
            );
            let output_main = main_bytes(&output);
            assert!(
                output_main
                    .windows(b"&lt;".len())
                    .any(|window| window == b"&lt;")
            );
            assert!(
                output_main
                    .windows(b"&amp;".len())
                    .any(|window| window == b"&amp;")
            );
            assert!(
                output_main
                    .windows(b"xml:space=\"preserve\"".len())
                    .any(|window| window == b"xml:space=\"preserve\"")
            );
            assert!(
                output_main
                    .windows(section.len())
                    .any(|window| window == section.as_slice()),
                "the final sectPr bytes must remain source-owned"
            );
            assert!(
                output
                    .windows(b"opaque physical payload\0\xff\x01".len())
                    .any(|window| window == b"opaque physical payload\0\xff\x01")
            );

            let candidate = open_package(output.clone());
            let mut inverse = Vec::new();
            publication
                .write_inverse_to_stream(&candidate, &mut inverse)
                .expect("immediate inverse must restore the exact source archive");
            assert_eq!(inverse, source_archive);
        }
    }
}

#[test]
fn default_and_shadowed_namespaces_cannot_capture_generated_prefix() {
    let source_archive = archive(&shadowed_source_xml(), false, Compression::Store);
    let package = open_package(source_archive);
    let plan = package
        .tail_append_plain_paragraphs(
            EventSource::new(events_for(&["tail"])),
            limits(shadowed_source_xml().len()),
        )
        .prepare()
        .expect("default namespace source must prepare");
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("shadowed namespace source must publish");
    assert_eq!(semantic_texts(&output), vec!["seed", "tail"]);
    assert!(
        main_bytes(&output)
            .windows(b"<w:p xmlns:w=\"".len())
            .any(|window| window == b"<w:p xmlns:w=\"")
    );
}

#[test]
fn default_stream_limits_accept_a_tiny_valid_append() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let package = open_package(source_archive);
    let plan = package
        .tail_append_plain_paragraphs(
            EventSource::new(events_for(&["default limits"])),
            ParagraphStreamLimits::default(),
        )
        .prepare()
        .expect("default stream limits must admit a tiny valid append");
    assert_eq!(plan.authored_proof().paragraph_count, 1);
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("default stream limits must publish a tiny valid append");
    assert_eq!(semantic_texts(&output), vec!["seed", "default limits"]);
}

#[test]
fn malformed_order_empty_stream_invalid_xml_and_provider_errors_refuse_before_output() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let invalid_streams = [
        vec![],
        vec![EventData::Text("outside".to_owned())],
        vec![EventData::Start],
        vec![EventData::Start, EventData::Start, EventData::End],
        vec![EventData::Start, EventData::End, EventData::End],
        vec![
            EventData::Start,
            EventData::Text("\u{1}".to_owned()),
            EventData::End,
        ],
        vec![
            EventData::Start,
            EventData::Text("\r".to_owned()),
            EventData::End,
        ],
        vec![
            EventData::Start,
            EventData::Text("\n".to_owned()),
            EventData::End,
        ],
        vec![
            EventData::Start,
            EventData::Text("\t".to_owned()),
            EventData::End,
        ],
    ];
    for events in invalid_streams {
        let package = open_package(source_archive.clone());
        let result = package
            .tail_append_plain_paragraphs(EventSource::new(events), limits(source_xml.len()))
            .prepare();
        assert!(matches!(result, Err(StreamError::Replay(_))));
    }

    let package = open_package(source_archive);
    let result = package
        .tail_append_plain_paragraphs(FailingSource, limits(source_xml.len()))
        .prepare();
    assert!(matches!(result, Err(StreamError::Replay(_))));
}

#[test]
fn deterministic_replay_change_and_truncated_event_eof_are_typed_refusals() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Deflate);
    let package = open_package(source_archive.clone());
    let changing = EventSource::changing(events_for(&["first"]), events_for(&["different"]));
    let result = package
        .tail_append_plain_paragraphs(changing, limits(source_xml.len()))
        .prepare();
    assert!(
        result.is_err(),
        "changed replay source unexpectedly prepared"
    );

    let package = open_package(source_archive.clone());
    assert!(
        package
            .tail_append_plain_paragraphs(
                EventSource::new(vec![
                    EventData::Start,
                    EventData::Text("missing end".to_owned()),
                ]),
                limits(source_xml.len()),
            )
            .prepare()
            .is_err(),
        "a replay pass ending with an open paragraph must refuse"
    );
    let package = open_package(source_archive);
    assert!(
        package
            .tail_append_plain_paragraphs(
                EventSource::new(vec![EventData::Start, EventData::End]),
                limits(source_xml.len()),
            )
            .prepare()
            .is_ok(),
        "a complete empty paragraph should be valid"
    );
}

#[test]
fn changed_text_event_boundaries_are_rejected_even_when_encoded_xml_matches() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let package = open_package(source_archive);
    let first = vec![
        EventData::Start,
        EventData::Text("same bytes".to_owned()),
        EventData::End,
    ];
    let second = vec![
        EventData::Start,
        EventData::Text("same ".to_owned()),
        EventData::Text("bytes".to_owned()),
        EventData::End,
    ];
    assert!(
        package
            .tail_append_plain_paragraphs(
                EventSource::changing(first, second),
                limits(source_xml.len()),
            )
            .prepare()
            .is_err(),
        "event-boundary changes must fail the event proof even when bytes match"
    );
}

#[test]
fn each_replay_pass_opens_one_cursor_independent_of_encoded_window_count() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let mut authored = vec![EventData::Start];
    authored.extend((0..512).map(|_| EventData::Text("x".repeat(64))));
    authored.push(EventData::End);
    let (source, opens) = EventSource::counted(authored);
    let mut stream_limits = limits(source_xml.len());
    stream_limits.max_replay_window_bytes = 512;
    stream_limits.max_authored_chunk_bytes = 64;
    stream_limits.source.max_token_bytes = 40 * 1024;
    stream_limits.source.max_workspace_bytes = 32 * 1024 * 1024;
    let package = open_package(source_archive);
    let plan = package
        .tail_append_plain_paragraphs(source, stream_limits)
        .prepare()
        .expect("large authored output must prepare");
    assert_eq!(
        opens.load(Ordering::Acquire),
        3,
        "sealing, DOCX candidate validation, and OPC candidate validation each open once"
    );
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("large authored output must publish");
    assert_eq!(
        opens.load(Ordering::Acquire),
        5,
        "measurement and emission each add one cursor pass rather than restarting for each output window"
    );
}

#[test]
fn source_size_and_authored_count_are_independent() {
    for padding in [0, 131_072] {
        let source_xml = source_xml(false, padding);
        let source_archive = archive(&source_xml, false, Compression::Store);
        let mut previous_source_len = None;
        for authored_count in [64_usize, 256, 4_096] {
            let texts = (0..authored_count)
                .map(|index| format!("authored-{index}"))
                .collect::<Vec<_>>();
            let package = open_package(source_archive.clone());
            let plan = package
                .tail_append_plain_paragraphs(
                    EventSource::new(events_for_owned(texts)),
                    limits(source_xml.len()),
                )
                .prepare()
                .expect("independent source/authored counts must prepare");
            assert_eq!(plan.source_proof().paragraph_count, 1);
            assert_eq!(plan.authored_proof().paragraph_count, authored_count as u64);
            assert_eq!(
                plan.authored_proof().event_count,
                (authored_count * 3) as u64
            );
            if let Some(previous) = previous_source_len {
                assert_eq!(previous, plan.source_proof().source_len);
            }
            previous_source_len = Some(plan.source_proof().source_len);
        }
    }

    let small = source_xml(false, 0);
    let large = source_xml(false, 131_072);
    assert_ne!(small.len(), large.len());
}

#[test]
fn one_shot_producer_requires_and_uses_explicit_replay_store() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let package = open_package(source_archive);
    let runs = Arc::new(AtomicUsize::new(0));
    let producer_runs = Arc::clone(&runs);
    let producer = move |sink: &mut dyn ParagraphEventSink| -> Result<(), AuthoredReplayError> {
        producer_runs.fetch_add(1, Ordering::AcqRel);
        sink.push(PlainParagraphEvent::ParagraphStart)?;
        sink.push(PlainParagraphEvent::TextChunk("one-shot"))?;
        sink.push(PlainParagraphEvent::ParagraphEnd)?;
        Ok(())
    };
    let limits = limits(source_xml.len());
    let edit = package
        .tail_append_plain_paragraphs_from_producer(
            producer,
            MemoryReplayStore::new(limits.max_replay_bytes).expect("store limit is finite"),
            limits,
        )
        .expect("explicit replay store route must be accepted");
    let plan = edit.prepare().expect("one-shot stream must prepare");
    assert_eq!(runs.load(Ordering::Acquire), 1);
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("one-shot stream must publish");
    assert_eq!(semantic_texts(&output), vec!["seed", "one-shot"]);
}

#[test]
fn one_shot_store_must_return_the_encoder_proof_before_candidate_or_output() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let package = open_package(source_archive);
    let stream_limits = limits(source_xml.len());
    let runs = Arc::new(AtomicUsize::new(0));
    let producer_runs = Arc::clone(&runs);
    let producer = move |sink: &mut dyn ParagraphEventSink| -> Result<(), AuthoredReplayError> {
        producer_runs.fetch_add(1, Ordering::AcqRel);
        sink.push(PlainParagraphEvent::ParagraphStart)?;
        sink.push(PlainParagraphEvent::TextChunk("producer bytes"))?;
        sink.push(PlainParagraphEvent::ParagraphEnd)?;
        Ok(())
    };
    let store = ForgedReplayStore {
        accepted: Some(
            MemoryReplayStore::new(stream_limits.max_replay_bytes)
                .expect("replay store limit is finite"),
        ),
        replacement: Some(presealed_forged_handle(stream_limits.max_replay_bytes)),
    };
    let edit = package
        .tail_append_plain_paragraphs_from_producer(producer, store, stream_limits)
        .expect("explicit replay store route must be accepted");
    let result = edit.prepare();
    assert!(matches!(
        result,
        Err(StreamError::Replay(AuthoredReplayError::Changed))
    ));
    assert_eq!(runs.load(Ordering::Acquire), 1);
}

#[test]
fn explicit_store_capacity_must_fit_selected_replay_limit_before_producer() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let package = open_package(source_archive);
    let store_capacity = 4 * 1024_u64;
    let selected_replay = 1024_u64;
    let mut stream_limits = limits(source_xml.len());
    stream_limits.max_replay_bytes = selected_replay;
    let runs = Arc::new(AtomicUsize::new(0));
    let producer_runs = Arc::clone(&runs);
    let producer = move |sink: &mut dyn ParagraphEventSink| -> Result<(), AuthoredReplayError> {
        producer_runs.fetch_add(1, Ordering::AcqRel);
        sink.push(PlainParagraphEvent::ParagraphStart)?;
        sink.push(PlainParagraphEvent::TextChunk("must not run"))?;
        sink.push(PlainParagraphEvent::ParagraphEnd)?;
        Ok(())
    };
    let edit = package
        .tail_append_plain_paragraphs_from_producer(
            producer,
            MemoryReplayStore::new(store_capacity).expect("store capacity is finite"),
            stream_limits,
        )
        .expect("selected replay limits must be valid");
    let result = edit.prepare();
    assert!(matches!(
        result,
        Err(StreamError::Replay(AuthoredReplayError::Limit {
            resource: "replay bytes",
            actual: 4096,
            maximum: 1024,
        }))
    ));
    assert_eq!(runs.load(Ordering::Acquire), 0);
}

#[test]
fn managed_store_memory_budget_refuses_before_one_shot_producer() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let store_capacity = 4 * 1024 * 1024;
    let (budget, package) = managed_package(source_archive, store_capacity, u64::MAX);
    let mut stream_limits = limits(source_xml.len());
    stream_limits.max_replay_bytes = store_capacity;
    let runs = Arc::new(AtomicUsize::new(0));
    let producer_runs = Arc::clone(&runs);
    let producer = move |sink: &mut dyn ParagraphEventSink| -> Result<(), AuthoredReplayError> {
        producer_runs.fetch_add(1, Ordering::AcqRel);
        sink.push(PlainParagraphEvent::ParagraphStart)?;
        sink.push(PlainParagraphEvent::TextChunk("budget refusal"))?;
        sink.push(PlainParagraphEvent::ParagraphEnd)?;
        Ok(())
    };
    let edit = package
        .tail_append_plain_paragraphs_from_producer(
            producer,
            MemoryReplayStore::new(store_capacity).expect("store capacity is finite"),
            stream_limits,
        )
        .expect("managed replay limits must be valid");
    let result = edit.prepare();
    assert!(matches!(
        result,
        Err(StreamError::Replay(AuthoredReplayError::Provider(_)))
    ));
    assert_eq!(runs.load(Ordering::Acquire), 0);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_store_lease_survives_plan_and_handle_then_releases() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let store_capacity = 2 * 1024 * 1024;
    let (budget, package) = managed_package(source_archive, 64 * 1024 * 1024, u64::MAX);
    let baseline = budget.used(Resource::Memory);
    let captured = Arc::new(Mutex::new(None));
    let producer = |sink: &mut dyn ParagraphEventSink| -> Result<(), AuthoredReplayError> {
        sink.push(PlainParagraphEvent::ParagraphStart)?;
        sink.push(PlainParagraphEvent::TextChunk("retained"))?;
        sink.push(PlainParagraphEvent::ParagraphEnd)?;
        Ok(())
    };
    let mut stream_limits = limits(source_xml.len());
    stream_limits.max_replay_bytes = store_capacity;
    let store = CapturingReplayStore {
        inner: Some(MemoryReplayStore::new(store_capacity).expect("store capacity is finite")),
        captured: Arc::clone(&captured),
    };
    let plan = package
        .tail_append_plain_paragraphs_from_producer(producer, store, stream_limits)
        .expect("managed replay store route must be accepted")
        .prepare()
        .expect("managed replay store must prepare");
    assert!(
        budget.used(Resource::Memory) >= baseline + store_capacity,
        "managed replay capacity must be charged while the plan owns it"
    );
    let replay_handle = captured
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
        .take()
        .expect("capturing store must retain its handle");
    drop(plan);
    assert!(
        budget.used(Resource::Memory) >= baseline + store_capacity,
        "the replay handle must retain the store lease after plan drop"
    );
    drop(replay_handle);
    assert_eq!(
        budget.used(Resource::Memory),
        baseline,
        "dropping the last replay handle must release its retained lease"
    );
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_external_output_is_charged_once_for_one_publication() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let stream_limits = limits(source_xml.len());
    let expected = {
        let package = open_package(source_archive.clone());
        let plan = package
            .tail_append_plain_paragraphs(
                EventSource::new(events_for(&["one output charge"])),
                stream_limits,
            )
            .prepare()
            .expect("unmanaged candidate must prepare");
        let mut output = Vec::new();
        plan.write_to_stream(&mut output)
            .expect("unmanaged candidate must publish");
        output
    };
    let expected_len = expected.len() as u64;
    let (budget, package) = managed_package(source_archive, 64 * 1024 * 1024, expected_len);
    let plan = package
        .tail_append_plain_paragraphs(
            EventSource::new(events_for(&["one output charge"])),
            stream_limits,
        )
        .prepare()
        .expect("managed candidate must prepare");
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("one publication must fit the exact external output budget");
    assert_eq!(output, expected);
    assert_eq!(output.len() as u64, expected_len);
    assert_eq!(budget.used(Resource::OutputBytes), expected_len);
    drop(package);
    assert_eq!(budget.used(Resource::OutputBytes), expected_len);
}

#[test]
fn managed_memory_store_zero_object_budget_refuses_before_append() {
    let (budget, _cancellation_source, context) =
        managed_context_with_objects(8 * 1024 * 1024, u64::MAX, 0);
    let mut store = MemoryReplayStore::new(64 * 1024).expect("store capacity is finite");
    let result =
        store.prepare_for_operation(ParagraphStreamLimits::default(), Some(&context), None);
    let error = match result {
        Err(AuthoredReplayError::Provider(error)) => error.to_string(),
        other => panic!("unexpected zero-object store result: {other:?}"),
    };
    assert!(
        error.contains("object") || error.contains("Object"),
        "zero-object refusal must identify the object resource: {error}"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn managed_replay_readers_each_hold_object_lease_until_final_handle_drop() {
    let (budget, _cancellation_source, context) =
        managed_context_with_objects(8 * 1024 * 1024, u64::MAX, 16);
    let handle = managed_forged_handle(64 * 1024, ParagraphStreamLimits::default(), &context);
    let owner_objects = budget.used(Resource::Objects);
    assert!(
        owner_objects > 0,
        "managed replay owner must hold an object lease"
    );

    let reader_a = handle
        .open()
        .expect("first managed replay reader must open");
    assert_eq!(budget.used(Resource::Objects), owner_objects + 1);
    let reader_b = handle
        .open()
        .expect("second managed replay reader must open");
    assert_eq!(budget.used(Resource::Objects), owner_objects + 2);
    drop(reader_a);
    assert_eq!(budget.used(Resource::Objects), owner_objects + 1);
    drop(reader_b);
    assert_eq!(budget.used(Resource::Objects), owner_objects);
    drop(handle);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn durable_resolver_cancellation_refuses_before_replay_reader_or_output() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let stream_limits = limits(source_xml.len());
    let authored_reference = AuthoredReplayReference::try_from_bytes(
        b"resolver-cancellation-provider",
        stream_limits.max_patch_bytes,
    )
    .expect("bounded authored provider reference must be accepted");
    let captured = Arc::new(Mutex::new(None));
    let producer = |sink: &mut dyn ParagraphEventSink| -> Result<(), AuthoredReplayError> {
        sink.push(PlainParagraphEvent::ParagraphStart)?;
        sink.push(PlainParagraphEvent::TextChunk("resolver cancellation"))?;
        sink.push(PlainParagraphEvent::ParagraphEnd)?;
        Ok(())
    };
    let store = CapturingReplayStore {
        inner: Some(
            MemoryReplayStore::new(stream_limits.max_replay_bytes)
                .expect("replay store limit is finite")
                .with_durable_reference(authored_reference.clone()),
        ),
        captured: Arc::clone(&captured),
    };
    let original = open_package(source_archive.clone());
    let plan = original
        .tail_append_plain_paragraphs_from_producer(producer, store, stream_limits)
        .expect("durable producer route must be accepted")
        .prepare()
        .expect("durable producer must prepare");
    let mut candidate = Vec::new();
    let publication = plan
        .write_to_stream(&mut candidate)
        .expect("durable producer must publish");
    let patch = publication
        .durable_patch()
        .expect("durable publication must retain its patch");
    let wire = patch.to_bytes().expect("durable patch must serialize");
    let decoded = ParagraphStreamPatch::from_bytes(&wire).expect("durable patch must decode");
    let captured_handle = Arc::new(
        captured
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .clone()
            .expect("durable store must expose its sealed handle"),
    );
    let opens = Arc::new(AtomicUsize::new(0));
    let counted_handle = Arc::new(CountingReplayHandle {
        inner: captured_handle,
        opens: Arc::clone(&opens),
    });
    let (budget, cancellation, current) =
        managed_package_with_cancellation(source_archive, 64 * 1024 * 1024, u64::MAX);
    let resolver = CancellingReplayResolver {
        expected: authored_reference,
        handle: counted_handle,
        cancellation,
    };
    let mut output = Vec::new();
    let result = current.apply_tail_append_stream_patch(&decoded, &resolver, &mut output);
    assert!(matches!(
        result,
        Err(StreamError::TailAppend(TailAppendError::Execution(
            litchi_core::ExecutionError::Cancelled
        )))
    ));
    assert!(
        output.is_empty(),
        "resolver cancellation must precede output"
    );
    assert_eq!(
        opens.load(Ordering::Acquire),
        0,
        "resolver cancellation must precede the first replay reader"
    );
    drop(current);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn durable_forward_and_inverse_authorizations_are_compact_canonical_wire() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let package = open_package(source_archive.clone());
    let stream_limits = limits(source_xml.len());
    let authored_reference = AuthoredReplayReference::try_from_bytes(
        b"caller-owned-authored-provider",
        stream_limits.max_patch_bytes,
    )
    .expect("bounded authored provider reference must be accepted");
    let plan = package
        .tail_append_plain_paragraphs(EventSource::new(events_for(&["durable"])), stream_limits)
        .prepare()
        .expect("durable patch fixture must prepare");
    let source_proof = plan.source_proof();
    let authored_proof = plan.authored_proof();
    let candidate_proof = plan.candidate_proof();
    let mut candidate_archive = Vec::new();
    plan.write_to_stream(&mut candidate_archive)
        .expect("durable patch fixture must publish");
    let source_artifact = artifact_proof(
        &SourceBackedPackage::from_vec(source_archive.clone())
            .expect("source artifact package must open")
            .source_artifact(),
    )
    .expect("source artifact fingerprint must be bounded");
    let candidate_artifact = artifact_proof(
        &SourceBackedPackage::from_vec(candidate_archive.clone())
            .expect("candidate artifact package must open")
            .source_artifact(),
    )
    .expect("candidate artifact fingerprint must be bounded");
    let patch = ParagraphStreamPatch::from_tail_append_proofs(
        PatchLimits::from_stream_limits(&stream_limits),
        MAIN_URI,
        source_proof,
        authored_proof,
        candidate_proof,
        source_artifact,
        candidate_artifact,
        Some(authored_reference),
    )
    .expect("durable forward patch must authenticate its proof facts");
    let bytes = patch.to_bytes().expect("durable patch must serialize");
    assert!(bytes.len() <= stream_limits.max_patch_bytes as usize);
    assert!(!bytes.windows(b"<w:p".len()).any(|window| window == b"<w:p"));
    assert_eq!(
        ParagraphStreamPatch::from_bytes(&bytes).expect("wire must round-trip"),
        patch
    );

    let original_reference = OriginalArtifactReference::try_from_bytes(
        b"caller-owned-original-provider",
        stream_limits.max_patch_bytes,
    )
    .expect("bounded original provider reference must be accepted");
    let inverse = patch
        .inverse_authorization(original_reference)
        .expect("durable inverse must require explicit original provider");
    let inverse_bytes = inverse
        .to_bytes()
        .expect("inverse authorization must serialize");
    assert_eq!(
        inverse,
        ExactInverseAuthorization::from_bytes(&inverse_bytes)
            .expect("inverse wire must round-trip")
    );

    let decoded_inverse = ExactInverseAuthorization::from_bytes(&inverse_bytes)
        .expect("decoded inverse must be applicable");
    let reopened_current = source_backed::Package::from_read_at(Arc::new(
        VersionedArchiveSource::new(candidate_archive.clone()),
    ))
    .expect("reopened candidate with a different source version must open");
    let original_provider = TestOriginalProvider {
        bytes: Some(source_archive.clone()),
    };
    let mut restored = Vec::new();
    apply_exact_inverse(
        &reopened_current,
        &decoded_inverse,
        &original_provider,
        &mut restored,
    )
    .expect("explicit original provider must restore across reopen");
    assert_eq!(restored, source_archive);

    let mut missing_output = Vec::new();
    assert!(
        apply_exact_inverse(
            &reopened_current,
            &decoded_inverse,
            &TestOriginalProvider { bytes: None },
            &mut missing_output,
        )
        .is_err()
    );
    assert!(
        missing_output.is_empty(),
        "missing provider must fail before output"
    );

    let mut mutated_original_output = Vec::new();
    assert!(
        apply_exact_inverse(
            &reopened_current,
            &decoded_inverse,
            &TestOriginalProvider {
                bytes: Some(mutate_opaque_payload(&source_archive)),
            },
            &mut mutated_original_output,
        )
        .is_err()
    );
    assert!(
        mutated_original_output.is_empty(),
        "mutated original artifact must fail before output"
    );

    let mismatched_current = source_backed::Package::from_read_at(Arc::new(
        VersionedArchiveSource::new(mutate_opaque_payload(&candidate_archive)),
    ))
    .expect("mutated candidate fixture must open");
    let provider_opens = Arc::new(AtomicUsize::new(0));
    let counting_provider = CountingOriginalProvider {
        bytes: source_archive.clone(),
        opens: Arc::clone(&provider_opens),
    };
    let mut mismatched_output = Vec::new();
    assert!(
        apply_exact_inverse(
            &mismatched_current,
            &decoded_inverse,
            &counting_provider,
            &mut mismatched_output,
        )
        .is_err()
    );
    assert!(
        mismatched_output.is_empty(),
        "candidate artifact mismatch must fail before output"
    );
    assert_eq!(
        provider_opens.load(Ordering::Acquire),
        0,
        "stale current artifact must be rejected before opening the original provider"
    );

    let mutable_current_source = Arc::new(VersionedArchiveSource::new(candidate_archive));
    let mutable_current = source_backed::Package::from_read_at(mutable_current_source.clone())
        .expect("mutable inverse current source must open");
    let mutating_provider = MutatingOriginalProvider {
        source: mutable_current_source,
    };
    let mut provider_changed_output = Vec::new();
    let provider_changed = apply_exact_inverse(
        &mutable_current,
        &decoded_inverse,
        &mutating_provider,
        &mut provider_changed_output,
    );
    assert!(matches!(
        provider_changed,
        Err(PatchError::Opc(
            litchi_opc::error::OpcError::SourceChanged { .. }
                | litchi_opc::error::OpcError::SourceArtifactMismatch { .. }
        ))
    ));
    assert!(
        provider_changed_output.is_empty(),
        "current source mismatch must precede original-provider output"
    );
}

#[test]
fn durable_forward_patch_resolves_after_reopen_and_refuses_wrong_source_or_provider() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let stream_limits = limits(source_xml.len());
    let authored_reference = AuthoredReplayReference::try_from_bytes(
        b"reopenable-authored-provider",
        stream_limits.max_patch_bytes,
    )
    .expect("bounded authored provider reference must be accepted");
    let captured = Arc::new(Mutex::new(None));
    let producer = |sink: &mut dyn ParagraphEventSink| -> Result<(), AuthoredReplayError> {
        sink.push(PlainParagraphEvent::ParagraphStart)?;
        sink.push(PlainParagraphEvent::TextChunk("reopened"))?;
        sink.push(PlainParagraphEvent::ParagraphEnd)?;
        Ok(())
    };
    let store = CapturingReplayStore {
        inner: Some(
            MemoryReplayStore::new(stream_limits.max_replay_bytes)
                .expect("replay store limit is finite")
                .with_durable_reference(authored_reference.clone()),
        ),
        captured: Arc::clone(&captured),
    };
    let original = open_package(source_archive.clone());
    let edit = original
        .tail_append_plain_paragraphs_from_producer(producer, store, stream_limits)
        .expect("durable producer route must be accepted");
    let plan = edit.prepare().expect("durable producer must prepare");
    let mut expected_candidate = Vec::new();
    let publication = plan
        .write_to_stream(&mut expected_candidate)
        .expect("durable producer must publish");
    let patch = publication
        .durable_patch()
        .expect("publication must retain a durable forward patch");
    assert!(patch.is_durable());
    let wire = patch.to_bytes().expect("durable patch must serialize");
    let decoded = ParagraphStreamPatch::from_bytes(&wire).expect("durable patch must decode");
    assert_eq!(decoded, patch);
    let captured_handle = Arc::new(
        captured
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .clone()
            .expect("one-shot store must expose its sealed handle to the resolver"),
    );
    let resolver = TestReplayResolver {
        expected: authored_reference,
        handle: captured_handle,
    };

    let reopened = source_backed::Package::from_read_at(Arc::new(VersionedArchiveSource::new(
        source_archive.clone(),
    )))
    .expect("reopened source with a different process-local version must open");
    let mut applied = Vec::new();
    reopened
        .apply_tail_append_stream_patch(&decoded, &resolver, &mut applied)
        .expect("durable forward patch must apply after reopening");
    assert_eq!(applied, expected_candidate);
    assert_eq!(semantic_texts(&applied), vec!["seed", "reopened"]);

    let (managed_budget, managed_reopened) = managed_package(
        source_archive.clone(),
        64 * 1024 * 1024,
        expected_candidate.len() as u64,
    );
    let mut managed_applied = Vec::new();
    managed_reopened
        .apply_tail_append_stream_patch(&decoded, &resolver, &mut managed_applied)
        .expect("managed durable forward patch must apply once to the external sink");
    assert_eq!(managed_applied, expected_candidate);
    assert_eq!(
        managed_budget.used(Resource::OutputBytes),
        expected_candidate.len() as u64,
        "managed durable publication must charge accepted external output once"
    );
    drop(managed_reopened);
    assert_eq!(
        managed_budget.used(Resource::OutputBytes),
        expected_candidate.len() as u64
    );

    let mutated_source = source_backed::Package::from_read_at(Arc::new(
        VersionedArchiveSource::new(mutate_opaque_payload(&source_archive)),
    ))
    .expect("mutated source fixture must open");
    let mut source_mismatch_output = Vec::new();
    assert!(
        mutated_source
            .apply_tail_append_stream_patch(&decoded, &resolver, &mut source_mismatch_output,)
            .is_err()
    );
    assert!(
        source_mismatch_output.is_empty(),
        "durable source artifact mismatch must fail before output"
    );

    let wrong_reference = AuthoredReplayReference::try_from_bytes(
        b"wrong-authored-provider",
        stream_limits.max_patch_bytes,
    )
    .expect("wrong reference fixture must be bounded");
    let wrong_resolver = TestReplayResolver {
        expected: wrong_reference,
        handle: resolver.handle.clone(),
    };
    let mut provider_mismatch_output = Vec::new();
    let provider_mismatch = reopened.apply_tail_append_stream_patch(
        &decoded,
        &wrong_resolver,
        &mut provider_mismatch_output,
    );
    assert!(matches!(
        provider_mismatch,
        Err(StreamError::Patch(PatchError::ReplayResolver(
            ReplayResolverError::Missing
        )))
    ));
    assert!(
        provider_mismatch_output.is_empty(),
        "wrong replay provider must fail before output"
    );
}

#[test]
fn durable_resolver_error_yields_source_mismatch_when_resolver_mutates_current() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let stream_limits = limits(source_xml.len());
    let authored_reference = AuthoredReplayReference::try_from_bytes(
        b"mutating-resolver-provider",
        stream_limits.max_patch_bytes,
    )
    .expect("bounded authored provider reference must be accepted");
    let producer = |sink: &mut dyn ParagraphEventSink| -> Result<(), AuthoredReplayError> {
        sink.push(PlainParagraphEvent::ParagraphStart)?;
        sink.push(PlainParagraphEvent::TextChunk("source precedence"))?;
        sink.push(PlainParagraphEvent::ParagraphEnd)?;
        Ok(())
    };
    let original = open_package(source_archive.clone());
    let plan = original
        .tail_append_plain_paragraphs_from_producer(
            producer,
            MemoryReplayStore::new(stream_limits.max_replay_bytes)
                .expect("replay store limit is finite")
                .with_durable_reference(authored_reference),
            stream_limits,
        )
        .expect("durable producer route must be accepted")
        .prepare()
        .expect("durable producer must prepare");
    let mut candidate = Vec::new();
    let publication = plan
        .write_to_stream(&mut candidate)
        .expect("durable producer must publish");
    let patch = publication
        .durable_patch()
        .expect("durable publication must retain its patch");
    let wire = patch.to_bytes().expect("durable patch must serialize");
    let decoded = ParagraphStreamPatch::from_bytes(&wire).expect("durable patch must decode");

    let source = Arc::new(VersionedArchiveSource::new(source_archive));
    let current = source_backed::Package::from_read_at(source.clone())
        .expect("mutable current source must open");
    let resolver = MutatingMissingReplayResolver { source };
    let mut output = Vec::new();
    let result = current.apply_tail_append_stream_patch(&decoded, &resolver, &mut output);
    assert!(matches!(
        result,
        Err(StreamError::Opc(
            litchi_opc::error::OpcError::SourceChanged { .. }
                | litchi_opc::error::OpcError::SourceArtifactMismatch { .. }
        ))
    ));
    assert!(
        output.is_empty(),
        "source freshness mismatch must precede resolver error output"
    );
}

#[test]
fn finite_limits_cancellation_and_partial_sink_errors_remain_at_public_boundary() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);

    let mut too_few = limits(source_xml.len());
    too_few.max_authored_paragraphs = 1;
    let package = open_package(source_archive.clone());
    assert!(
        package
            .tail_append_plain_paragraphs(EventSource::new(events_for(&["one", "two"])), too_few,)
            .prepare()
            .is_err(),
        "paragraph ceiling must be enforced before publication"
    );

    let mut too_large_chunk = limits(source_xml.len());
    too_large_chunk.max_authored_chunk_bytes = 2;
    let package = open_package(source_archive.clone());
    assert!(
        package
            .tail_append_plain_paragraphs(
                EventSource::new(events_for(&["three"])),
                too_large_chunk,
            )
            .prepare()
            .is_err(),
        "borrowed chunk ceiling must be enforced before retention"
    );

    let mut too_small_xml = limits(source_xml.len());
    too_small_xml.max_authored_xml_bytes = 1;
    let package = open_package(source_archive.clone());
    assert!(
        package
            .tail_append_plain_paragraphs(EventSource::new(events_for(&["one"])), too_small_xml,)
            .prepare()
            .is_err(),
        "generated XML ceiling must be finite and enforced"
    );

    let package = open_package(source_archive.clone());
    let (cancellation_source, cancellation) = CancellationSource::pair();
    cancellation_source.cancel();
    assert!(
        package
            .tail_append_plain_paragraphs(
                EventSource::new(events_for(&["cancelled"])),
                limits(source_xml.len()),
            )
            .with_cancellation_token(&cancellation)
            .prepare()
            .is_err(),
        "cancelled stream must refuse before output"
    );

    let package = open_package(source_archive);
    let plan = package
        .tail_append_plain_paragraphs(
            EventSource::new(events_for(&["sink failure"])),
            limits(source_xml.len()),
        )
        .prepare()
        .expect("partial sink fixture plan must prepare");
    let mut sink = PrefixFailSink::after(32);
    assert!(plan.write_to_stream(&mut sink).is_err());
    assert!(
        !sink.bytes.is_empty(),
        "sink failure must retain accepted prefix"
    );
}

#[test]
fn source_change_after_preparation_refuses_without_archive_output() {
    let source_xml = source_xml_without_section(false);
    let source_archive = archive(&source_xml, false, Compression::Store);
    let source = Arc::new(VersionedArchiveSource::new(source_archive.clone()));
    let package = source_backed::Package::from_read_at(source.clone())
        .expect("versioned source fixture must open");
    let plan = package
        .tail_append_plain_paragraphs(
            EventSource::new(events_for(&["tail"])),
            limits(source_xml.len()),
        )
        .prepare()
        .expect("versioned source plan must prepare");
    source.bump();
    let mut output = Vec::new();
    assert!(
        plan.write_to_stream(&mut output).is_err(),
        "stale source must refuse publication"
    );
    assert!(
        output.is_empty(),
        "stale preflight must emit no archive bytes"
    );
}
