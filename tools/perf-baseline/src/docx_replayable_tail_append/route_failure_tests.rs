//! End-to-end failure probes for the explicit file-backed replay route.
//!
//! These tests deliberately pass through the DOCX stream planner and OPC
//! publication path.  The sink accepts a one-byte publication prefix before
//! mutating the sealed replay file, so the subsequent publication pass must
//! retain both its exact accepted-byte count and the typed replay failure.

use super::{
    AuthoredCounters, AuthoredSpec, ChunkMode, CompressionProfile, FileReplayStore, FileSyncPolicy,
    GeneratedParagraphProducer, MeasureSource, ReplayCounters, SourceCounters, TextMode,
    build_fixture,
};
use litchi_core::ReadAt;
use litchi_docx::source_backed;
use litchi_docx::source_backed::tail_append_stream::{AuthoredReplayError, Error as StreamError};
use litchi_opc::OpcError;
use std::error::Error as StdError;
use std::fs::{self, OpenOptions};
use std::io::{self, Read as _, Seek, SeekFrom, Write};
use std::path::{Path, PathBuf};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

static NEXT_REPLAY_FAILURE_TEST: AtomicU64 = AtomicU64::new(0);

struct TempReplayPath {
    directory: PathBuf,
    file: PathBuf,
}

impl TempReplayPath {
    fn new() -> Self {
        let serial = NEXT_REPLAY_FAILURE_TEST.fetch_add(1, Ordering::Relaxed);
        let directory = std::env::temp_dir().join(format!(
            "litchi-docx-replay-route-failure-{}-{serial}",
            std::process::id()
        ));
        fs::create_dir(&directory).expect("create bounded replay-failure directory");
        let file = directory.join("authored.replay");
        Self { directory, file }
    }

    fn path(&self) -> &Path {
        &self.file
    }
}

impl Drop for TempReplayPath {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.file);
        let _ = fs::remove_dir(&self.directory);
    }
}

struct MutatingPrefixSink {
    replay_path: PathBuf,
    bytes: Vec<u8>,
    mutated: bool,
}

impl MutatingPrefixSink {
    fn new(replay_path: &Path) -> Self {
        Self {
            replay_path: replay_path.to_path_buf(),
            bytes: Vec::new(),
            mutated: false,
        }
    }

    fn mutate_replay_file(&self) {
        let mut file = OpenOptions::new()
            .read(true)
            .write(true)
            .open(&self.replay_path)
            .expect("open sealed replay file for deterministic mutation");
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)
            .expect("sealed replay file must contain authored bytes");
        let marker = b"authored-";
        let offset = bytes
            .windows(marker.len())
            .position(|window| window == marker)
            .and_then(|position| position.checked_add(marker.len()))
            .expect("authored replay must contain a deterministic text marker");
        assert!(bytes[offset].is_ascii_digit());
        bytes[offset] = if bytes[offset] == b'0' { b'1' } else { b'0' };
        file.seek(SeekFrom::Start(
            u64::try_from(offset).expect("bounded replay offset fits u64"),
        ))
        .expect("seek to deterministic replay mutation byte");
        file.write_all(&bytes[offset..offset + 1])
            .expect("mutate one replay byte without changing its length");
        file.flush().expect("flush deterministic replay mutation");
    }
}

impl Write for MutatingPrefixSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.is_empty() {
            return Ok(0);
        }

        if self.mutated {
            self.bytes.extend_from_slice(bytes);
            return Ok(bytes.len());
        }

        // Accept exactly one publication byte first.  The OPC preservation
        // writer retries the remaining prefix, and its later target callback
        // then has to consume the mutated replay file.
        self.bytes.push(bytes[0]);
        self.mutate_replay_file();
        self.mutated = true;
        Ok(1)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn contains_authored_changed(error: &OpcError) -> bool {
    match error {
        OpcError::IncompleteOutput { source, .. } => contains_authored_changed(source),
        OpcError::IoError(error) => contains_authored_changed_in_io(error),
        _ => false,
    }
}

fn contains_authored_changed_in_io(error: &io::Error) -> bool {
    error
        .get_ref()
        .is_some_and(|source| contains_authored_changed_in_source(source))
}

fn contains_authored_changed_in_source(error: &(dyn StdError + 'static)) -> bool {
    if error
        .downcast_ref::<AuthoredReplayError>()
        .is_some_and(|value| matches!(value, AuthoredReplayError::Changed))
    {
        return true;
    }
    if let Some(io_error) = error.downcast_ref::<io::Error>()
        && contains_authored_changed_in_io(io_error)
    {
        return true;
    }
    // quick-xml retains I/O failures in Arc; the XML audit carries that
    // owner through io::Error. Arc's Error::source skips its inner I/O node,
    // so inspect the retained owner before following the generic chain.
    if let Some(io_error) = error.downcast_ref::<Arc<io::Error>>()
        && contains_authored_changed_in_io(io_error)
    {
        return true;
    }
    error
        .source()
        .is_some_and(contains_authored_changed_in_source)
}

fn stream_opc_error(error: &StreamError) -> &OpcError {
    match error {
        StreamError::Opc(error) => error,
        other => panic!("publication must preserve an OPC output error: {other:?}"),
    }
}

#[test]
fn file_replay_mutation_after_publication_prefix_is_incomplete_and_typed() {
    let fixture = build_fixture(
        64,
        AuthoredSpec {
            count: 2,
            chunk_mode: ChunkMode::Fixed64,
            text_mode: TextMode::Short,
        },
        CompressionProfile::Current,
    )
    .expect("bounded DOCX fixture");
    let temporary = TempReplayPath::new();
    let replay_max_bytes = 8 * 1024_u64;
    let mut limits = fixture.limits;
    limits.max_replay_bytes = replay_max_bytes;

    let source_counters = Arc::new(SourceCounters::default());
    let source = Arc::new(MeasureSource::new(
        Arc::clone(&fixture.source_archive),
        source_counters,
    ));
    let package = source_backed::Package::from_read_at(Arc::clone(&source) as Arc<dyn ReadAt>)
        .expect("source-backed package");
    let authored_counters = Arc::new(AuthoredCounters::default());
    let replay_counters = Arc::new(ReplayCounters::default());
    let producer =
        GeneratedParagraphProducer::new(fixture.authored, authored_counters, replay_counters);
    let store = FileReplayStore::new(temporary.path(), replay_max_bytes, FileSyncPolicy::None)
        .expect("file replay store");
    let edit = package
        .tail_append_plain_paragraphs_from_producer(producer, store, limits)
        .expect("one-shot file replay edit");
    let plan = edit.prepare().expect("prepare and authenticate replay");

    let mut sink = MutatingPrefixSink::new(temporary.path());
    let error = plan
        .write_to_stream(&mut sink)
        .expect_err("mutation during DOCX publication must fail");
    assert!(sink.mutated, "sink must mutate after accepting its prefix");
    assert!(!sink.bytes.is_empty(), "publication must accept a prefix");

    let opc_error = stream_opc_error(&error);
    let written = match opc_error {
        OpcError::IncompleteOutput { written, .. } => *written,
        other => panic!("publication must report IncompleteOutput: {other:?}"),
    };
    assert_eq!(
        written,
        u64::try_from(sink.bytes.len()).expect("bounded sink bytes fit u64"),
        "IncompleteOutput.written must equal bytes accepted by the publication sink"
    );
    assert!(written > 0, "the mutation must occur after output progress");
    assert!(
        contains_authored_changed(opc_error),
        "typed AuthoredReplayError::Changed must survive the publication error chain: {error:?}"
    );
}
