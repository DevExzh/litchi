//! End-to-end smoke tests for the explicit authored replay routes.
//!
//! These tests deliberately call the same `run_case` path used by the CLI.
//! They inspect the returned sample counters against that case's proof so a
//! route cannot pass by returning a hand-built observation that only matches
//! the expected shape.

use super::*;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::atomic::{AtomicU64, Ordering};

static NEXT_ROUTE_SMOKE: AtomicU64 = AtomicU64::new(0);

struct TempRouteDirectory {
    path: PathBuf,
}

impl TempRouteDirectory {
    fn new() -> Self {
        let serial = NEXT_ROUTE_SMOKE.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "litchi-docx-replay-route-smoke-{}-{serial}",
            std::process::id()
        ));
        fs::create_dir(&path).expect("create isolated route-smoke directory");
        Self { path }
    }

    fn path(&self) -> &Path {
        &self.path
    }
}

impl Drop for TempRouteDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.path);
    }
}

const TEST_REPLAY_MAX_BYTES: u64 = 1024 * 1024;

fn config_for(
    provider: AuthoredProvider,
    replay_directory: Option<PathBuf>,
    replay_max_bytes: Option<u64>,
) -> Config {
    Config {
        source_counts: vec![2],
        authored_counts: vec![1, 2],
        chunk_modes: vec![ChunkMode::Fixed64],
        text_modes: vec![TextMode::Empty, TextMode::Short, TextMode::NearLimit],
        samples: 1,
        warmups: 0,
        sink_write_bytes: SINK_WRITE_BYTES[0],
        json_path: None,
        fixture_dir: None,
        authored_provider: provider,
        replay_dir: replay_directory,
        replay_max_bytes,
        replay_sync: if matches!(provider, AuthoredProvider::FileStore) {
            ReplaySync::Data
        } else {
            ReplaySync::None
        },
        compression: CompressionProfile::Current,
        input_mode: None,
        input_file: None,
        input_max_range_bytes: None,
        input_delay_us: 0,
        input_overhead_us: 0,
        input_bytes_per_second: None,
        publication: PublicationMode::HashingSink,
    }
}

fn store_case(
    provider: AuthoredProvider,
    authored_count: usize,
    text_mode: TextMode,
    replay_directory: Option<PathBuf>,
) -> CaseRecord {
    let config = config_for(provider, replay_directory, Some(TEST_REPLAY_MAX_BYTES));
    run_case(2, authored_count, ChunkMode::Fixed64, text_mode, &config)
        .expect("full explicit replay route case")
}

fn assert_store_sample(case: &CaseRecord, provider: AuthoredProvider) {
    assert_eq!(case.provider, provider);
    assert_eq!(case.samples.len(), 1);
    let sample = &case.samples[0];
    assert_eq!(sample.source_count, 2);
    assert_eq!(sample.authored_count, case.authored_count);
    assert_eq!(sample.chunk_mode, ChunkMode::Fixed64);
    let replay = sample
        .replay
        .expect("store route must report actual replay counters");
    assert_eq!(replay.route, Some(provider));
    assert_eq!(replay.producer_invocations, 1);
    assert_eq!(replay.prepare_calls, 1);
    assert!(replay.append_calls > 0);
    assert_eq!(replay.appended_bytes, case.authored.encoded_xml_bytes);
    assert_eq!(replay.store_finish_calls, 1);
    assert_eq!(replay.replay_opens, provider.replay_opens());
    assert!(replay.replay_read_calls > 0);
    assert_eq!(
        replay.replay_returned_bytes,
        case.authored
            .encoded_xml_bytes
            .checked_mul(provider.replay_opens())
            .expect("bounded replay byte count")
    );
    assert_eq!(replay.replay_finish_calls, provider.replay_opens());
    assert_eq!(replay.replay_sha256_checks, provider.replay_opens());

    // The one-shot producer has no source cursor opens.  These exact event and
    // text counters exercise empty text as well as nonempty chunk framing.
    assert_eq!(sample.authored.opens, 0);
    assert_eq!(sample.authored.events, case.authored.event_count);
    assert_eq!(sample.authored.text_chunks, case.authored.text_chunk_count);
    assert_eq!(sample.authored.text_bytes, case.authored.text_bytes);
    assert_eq!(
        sample
            .sink
            .accepted_bytes
            .expect("hashing sink must report accepted bytes"),
        u64::try_from(case.oracle.candidate_archive_bytes).expect("archive length fits u64")
    );

    match provider {
        AuthoredProvider::MemoryStore => {
            assert!(replay.file.is_none());
            assert_eq!(
                replay.retained_logical_bytes,
                Some(case.authored.encoded_xml_bytes)
            );
            assert!(
                replay
                    .retained_capacity_bytes
                    .is_some_and(|capacity| capacity >= case.authored.encoded_xml_bytes)
            );
            assert_eq!(replay.durable_reference_kind, Some("none"));
            assert_eq!(replay.durable_reference_bytes, 0);
            assert!(replay.durable_reference_sha256.is_none());
        },
        AuthoredProvider::FileStore => {
            let file = replay
                .file
                .expect("file route must report post-drop observations");
            assert!(file.write_calls > 0);
            assert!(file.read_calls > 0);
            assert_eq!(
                file.returned_bytes,
                case.authored
                    .encoded_xml_bytes
                    .checked_mul(provider.replay_opens())
                    .expect("bounded file read count")
            );
            assert_eq!(file.replay_sha256_checks, provider.replay_opens());
            assert_eq!(file.seal_sha256_checks, 1);
            assert_eq!(file.cleanup_sha256_checks, 1);
            assert_eq!(file.logical_bytes, Some(case.authored.encoded_xml_bytes));
            assert!(file.cleanup_verified);
            assert_eq!(replay.durable_reference_kind, Some("file"));
            assert!(replay.durable_reference_bytes > 0);
            assert!(replay.durable_reference_sha256.is_some());
        },
        AuthoredProvider::Deterministic => panic!("smoke helper requires a store provider"),
    }
}

#[test]
fn memory_store_run_case_short_and_empty_routes() {
    let short = store_case(AuthoredProvider::MemoryStore, 2, TextMode::Short, None);
    assert_store_sample(&short, AuthoredProvider::MemoryStore);

    // Empty text is a separate route invocation because it is the regression
    // case where a zero text-byte stream can hide a missing producer pass.
    let empty = store_case(AuthoredProvider::MemoryStore, 2, TextMode::Empty, None);
    assert_store_sample(&empty, AuthoredProvider::MemoryStore);
    assert_eq!(empty.authored.text_bytes, 0);
    assert_eq!(empty.authored.text_chunk_count, 0);
}

#[test]
fn file_store_run_case_short_route() {
    let temporary = TempRouteDirectory::new();
    let case = store_case(
        AuthoredProvider::FileStore,
        2,
        TextMode::Short,
        Some(temporary.path().to_owned()),
    );
    assert_store_sample(&case, AuthoredProvider::FileStore);
    assert!(
        fs::read_dir(temporary.path())
            .expect("inspect file route cleanup")
            .next()
            .is_none(),
        "file replay route must clean its exclusive artifact"
    );
}

#[test]
fn memory_store_run_case_near_limit_route() {
    let case = store_case(AuthoredProvider::MemoryStore, 1, TextMode::NearLimit, None);
    assert_store_sample(&case, AuthoredProvider::MemoryStore);
    assert_eq!(case.authored.text_bytes, MAX_CURSOR_TEXT_BYTES as u64);
}

#[test]
fn file_store_run_case_near_limit_route() {
    let temporary = TempRouteDirectory::new();
    let case = store_case(
        AuthoredProvider::FileStore,
        1,
        TextMode::NearLimit,
        Some(temporary.path().to_owned()),
    );
    assert_store_sample(&case, AuthoredProvider::FileStore);
    assert_eq!(case.authored.text_bytes, MAX_CURSOR_TEXT_BYTES as u64);
    assert!(
        fs::read_dir(temporary.path())
            .expect("inspect near-limit file cleanup")
            .next()
            .is_none(),
        "near-limit file replay route must clean its exclusive artifact"
    );
}
