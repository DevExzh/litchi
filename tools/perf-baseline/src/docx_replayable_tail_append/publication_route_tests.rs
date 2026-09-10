//! Focused tests for the after-only publication extension.
//!
//! The legacy hashing route is exercised by `route_smoke_tests`.  These tests
//! drive the counting and atomic routes through `run_case`, including every
//! authored provider, so a report cannot claim a route while using an
//! unrelated or materializing fallback.

use super::*;
use std::ffi::OsString;
use std::fs;
use std::path::PathBuf;
use std::sync::atomic::{AtomicU64, Ordering};

static NEXT_PUBLICATION_TEST: AtomicU64 = AtomicU64::new(0);

struct ReplayDirectory(PathBuf);

impl ReplayDirectory {
    fn new() -> Self {
        let serial = NEXT_PUBLICATION_TEST.fetch_add(1, Ordering::Relaxed);
        let path = std::env::temp_dir().join(format!(
            "litchi-docx-publication-test-{}-{serial}",
            std::process::id()
        ));
        fs::create_dir(&path).expect("create isolated publication-test directory");
        Self(path)
    }
}

impl Drop for ReplayDirectory {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.0);
    }
}

fn config(
    publication: PublicationMode,
    provider: AuthoredProvider,
    replay: Option<PathBuf>,
) -> Config {
    Config {
        source_counts: vec![2],
        authored_counts: vec![1],
        chunk_modes: vec![ChunkMode::Fixed64],
        text_modes: vec![TextMode::Short],
        samples: 1,
        warmups: 0,
        sink_write_bytes: SINK_WRITE_BYTES[0],
        json_path: None,
        fixture_dir: None,
        authored_provider: provider,
        replay_dir: replay,
        replay_max_bytes: provider.is_store().then_some(1024 * 1024),
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
        publication,
    }
}

fn providers() -> [AuthoredProvider; 3] {
    [
        AuthoredProvider::Deterministic,
        AuthoredProvider::MemoryStore,
        AuthoredProvider::FileStore,
    ]
}

#[test]
fn publication_cli_modes_are_explicit_and_bounded() {
    for (argument, expected) in [
        ("hashing-sink", PublicationMode::HashingSink),
        ("counting-sink", PublicationMode::CountingSink),
        ("atomic-path", PublicationMode::AtomicPath),
    ] {
        let config = parse_args([OsString::from("--publication"), OsString::from(argument)])
            .expect("finite publication mode");
        assert_eq!(config.publication, expected);
    }
    assert!(parse_args([OsString::from("--publication"), OsString::from("unknown")]).is_err());
}

#[test]
fn counting_route_uses_production_artifact_proof_for_every_provider() {
    let replay = ReplayDirectory::new();
    for provider in providers() {
        let replay_dir = matches!(provider, AuthoredProvider::FileStore).then(|| replay.0.clone());
        let case = run_case(
            2,
            1,
            ChunkMode::Fixed64,
            TextMode::Short,
            &config(PublicationMode::CountingSink, provider, replay_dir),
        )
        .expect("counting publication route");
        let sample = &case.samples[0];
        let publication = sample
            .publication
            .as_ref()
            .expect("counting route extension record");
        assert_eq!(publication.schema, PUBLICATION_SCHEMA);
        assert_eq!(publication.route, PublicationMode::CountingSink);
        assert_eq!(
            publication.timing_scope,
            "source_admission_prepare_sequential_sink_publication_drop"
        );
        assert!(publication.timed_candidate_matches_oracle);
        assert!(publication.atomic.is_none());
        assert_eq!(sample.sink.sha256, None);
        assert_eq!(
            sample
                .sink
                .accepted_bytes
                .expect("counting sink byte counter"),
            u64::try_from(case.oracle.candidate_archive_bytes).expect("candidate length")
        );
    }
}

#[test]
fn atomic_route_publishes_and_cleans_a_real_destination_for_every_provider() {
    let replay = ReplayDirectory::new();
    for provider in providers() {
        let replay_dir = matches!(provider, AuthoredProvider::FileStore).then(|| replay.0.clone());
        let case = run_case(
            2,
            1,
            ChunkMode::Fixed64,
            TextMode::Short,
            &config(PublicationMode::AtomicPath, provider, replay_dir),
        )
        .expect("atomic publication route");
        let sample = &case.samples[0];
        let publication = sample
            .publication
            .as_ref()
            .expect("atomic route extension record");
        assert_eq!(publication.route, PublicationMode::AtomicPath);
        assert_eq!(
            publication.timing_scope,
            "source_admission_prepare_atomic_write_data_sync_rename_parent_directory_sync_publication_drop"
        );
        assert!(publication.timed_candidate_matches_oracle);
        assert!(sample.sink.accepted_bytes.is_none());
        assert!(sample.sink.write_calls.is_none());
        assert!(sample.sink.histogram.is_none());
        let json = serde_json::to_value(sample).expect("serialize atomic sample");
        assert!(json["sink"]["accepted_bytes"].is_null());
        assert!(json["sink"]["sha256"].is_null());
        let atomic = publication.atomic.as_ref().expect("atomic output record");
        assert!(!atomic.before.exists);
        assert!(atomic.after.exists);
        assert!(atomic.after.regular_file);
        assert!(atomic.output_bytes_exact);
        assert!(atomic.output_sha256_exact);
        assert!(atomic.post_timer_oracle.candidate_xml_exact);
        assert!(atomic.post_timer_oracle.candidate_semantic_exact);
        assert!(atomic.post_timer_oracle.untouched_raw_members_preserved);
        assert!(atomic.post_timer_oracle.inverse_exact);
        assert_eq!(
            atomic.inverse_oracle_scope,
            "untimed_fixture_publication_inverse_exact; timed_atomic_publication_inverse_not_reexecuted"
        );
        assert!(atomic.cleanup.destination_removed);
        assert!(atomic.cleanup.parent_removed);
        assert!(!PathBuf::from(&atomic.destination_path).exists());
        assert!(!PathBuf::from(&atomic.private_parent_path).exists());
    }
}

#[test]
fn default_route_omits_the_publication_extension_and_keeps_hashing_sink_shape() {
    let case = run_case(
        2,
        1,
        ChunkMode::Fixed64,
        TextMode::Short,
        &config(
            PublicationMode::HashingSink,
            AuthoredProvider::Deterministic,
            None,
        ),
    )
    .expect("legacy hashing publication route");
    let sample = &case.samples[0];
    assert!(sample.publication.is_none());
    assert!(sample.sink.sha256.is_some());
    let json = serde_json::to_value(&case).expect("serialize case");
    assert!(json["samples"][0].get("publication").is_none());
    assert!(json["samples"][0]["sink"]["sha256"].is_string());
}
