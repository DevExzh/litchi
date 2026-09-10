//! Verified cold-filesystem DOCX one-edit/save evidence.
//!
//! This is intentionally a separate child-process runner.  The parent builds
//! all expected values, creates one page-aligned private source file, and
//! asks a fresh child to repeat the residency probe immediately before the
//! timed source-backed edit.  An ineligible probe is a successful, explicit
//! evidence outcome with zero timed rows.

use std::{
    env,
    error::Error,
    ffi::OsString,
    fs::{self, OpenOptions},
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Stdio},
    sync::Arc,
    time::Instant,
};

use litchi_core::{FileSource, OwnedSource, ReadAt, SourceVersion};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};

use super::{ReadStats, SequentialSink};
use crate::{cold_verified, process_metrics};

const SCHEMA: &str = "docx_edit_provider_cold_v1";
const CHILD_SCHEMA: &str = "docx_edit_provider_cold_child_v1";
const BENCHMARK: &str = "one DOCX paragraph replacement through verified cold FileSource open/commit/sequential publication";
const PROVIDER_SCOPE: &str = "fresh-child verified cold regular-file FileSource for one source-backed DOCX paragraph replacement";
const TIMING_SCOPE: &str = "FileSource open + source-backed DOCX open + one paragraph edit staging/commit + sequential publication + commit/package/document drops; source version fences, commit diagnostics and source/candidate XML identity comparisons are inside; cold preparation, expected-value construction, sink reservation, output hashing, semantic/media verification and preflight patch oracles are outside";
const SETUP_SCOPE: &str = "aligned source construction, source staging, expected publication, all semantic and patch oracles, counter reservation, and cold verifier setup are outside the operation clock";
const COLD_SCOPE: &str = cold_verified::CLAIM_SCOPE;
const PROCESS_METRICS_SCOPE: &str = "fresh measured child process /proc interval only; the parent and child supervisor process tree are excluded; the after snapshot includes procfs probe overhead";
const RSS_SCOPE: &str = "fresh measured child /proc/self/status RSS and VmHWM only; this is process-local evidence and excludes the parent process tree";
const MAX_HASH_HEX_BYTES: usize = 64;
const SINK_MAX_WRITE_BYTES: u64 = 1024 * 1024;

#[derive(Debug)]
struct ColdCase {
    corpus: crate::Corpus,
    aligned_source: Vec<u8>,
    expected_output: Vec<u8>,
    expected_output_sha256: String,
    source_document_xml: Vec<u8>,
    candidate_document_xml: Vec<u8>,
    media_ranges: Vec<std::ops::Range<u64>>,
    preflight: ColdPreflight,
}

#[derive(Clone, Debug, Serialize)]
struct ColdPreflight {
    expected_materializations: usize,
    source_document_xml_sha256: String,
    candidate_document_xml_sha256: String,
    output_exact_source_changed: bool,
    semantic_reopen_verified: bool,
    unchanged_media_preserved: bool,
    replay_forward_verified: bool,
    inverse_restores_source_verified: bool,
    stale_target_refusal_verified: bool,
    foreign_source_refusal_verified: bool,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct SourceVersionRecord {
    id: u64,
    revision: u64,
}

impl From<SourceVersion> for SourceVersionRecord {
    fn from(version: SourceVersion) -> Self {
        Self {
            id: version.id(),
            revision: version.revision(),
        }
    }
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct PatchOracle {
    scope: String,
    replay_forward: bool,
    inverse_restores_source: bool,
    stale_target_refused: bool,
    foreign_source_refused: bool,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct OracleEvidence {
    commit_identity_verified: bool,
    output_exact_bytes: bool,
    semantic_reopen: bool,
    unchanged_media_preserved: bool,
    source_version_unchanged: bool,
    cache_load_count: bool,
    logical_source_reads_positive: bool,
    patch_oracles: PatchOracle,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct ColdReadCounter {
    availability: String,
    scope: String,
    calls: Option<u64>,
    empty_calls: Option<u64>,
    requested_bytes: Option<u64>,
    returned_bytes: Option<u64>,
    short_reads: Option<u64>,
    min_request_bytes: Option<u64>,
    max_request_bytes: Option<u64>,
    traced_ranges: Option<usize>,
    ranges: Option<Vec<super::RangeRecord>>,
    requested_media_overlap_bytes: Option<u64>,
    returned_media_overlap_bytes: Option<u64>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct ColdReadEvidence {
    source: ColdReadCounter,
    scope: String,
}

#[derive(Clone, Copy, Debug, Serialize, Deserialize)]
struct ColdSinkRecord {
    accepted_bytes: u64,
    write_calls: u64,
    largest_write: u64,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct ColdCacheEvidence {
    successful_loads: usize,
    expected_successful_loads: usize,
    exactly_one_main_part_materialization: bool,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct ColdRow {
    sample_index: usize,
    child_process_id: u32,
    latency_ns: u64,
    output_bytes: usize,
    output_sha256: String,
    materializations: usize,
    commit_changed: bool,
    commit_operations: usize,
    source_version_before: SourceVersionRecord,
    source_version_after: SourceVersionRecord,
    source_version_unchanged: bool,
    reads: ColdReadEvidence,
    sink: ColdSinkRecord,
    cache: ColdCacheEvidence,
    #[serde(skip_serializing_if = "Option::is_none")]
    process_metrics: Option<process_metrics::Delta>,
    process_metrics_scope: String,
    rss_scope: String,
    oracles: OracleEvidence,
    #[serde(skip_serializing_if = "Option::is_none")]
    allocation: Option<crate::allocation_metrics::Sample>,
}

#[derive(Clone, Debug, Serialize, Deserialize)]
struct ColdChildOutput {
    schema: String,
    cold_verified: cold_verified::Sample,
    #[serde(skip_serializing_if = "Option::is_none")]
    row: Option<ColdRow>,
    #[serde(skip_serializing_if = "Option::is_none")]
    error: Option<String>,
}

#[derive(Clone, Debug, Serialize)]
struct ColdReport {
    limits: super::LimitsRecord,
    schema: &'static str,
    benchmark: &'static str,
    provider_scope: &'static str,
    timing_scope: &'static str,
    setup_scope: &'static str,
    cold_claim_scope: &'static str,
    process_metrics_scope: &'static str,
    rss_scope: &'static str,
    provider: &'static str,
    cache_state: &'static str,
    filesystem_root_selected: bool,
    source_archive_sha256: String,
    source_archive_bytes: usize,
    aligned_source_sha256: Option<String>,
    aligned_source_bytes: Option<usize>,
    expected_output_sha256: Option<String>,
    expected_output_bytes: Option<usize>,
    source_revision: String,
    corpus: Option<crate::CorpusManifest>,
    preflight: Option<ColdPreflight>,
    cold_verified_status: cold_verified::Status,
    cold_verified_samples: Vec<cold_verified::Sample>,
    cold_verified_fincore_command: &'static str,
    warmup: usize,
    samples: usize,
    rows: Vec<ColdRow>,
}

#[derive(Debug)]
struct StagedAlignedFile {
    root: PathBuf,
    path: PathBuf,
}

impl Drop for StagedAlignedFile {
    fn drop(&mut self) {
        let _ = fs::remove_file(&self.path);
        let _ = fs::remove_dir(&self.root);
    }
}

#[derive(Clone, Debug)]
struct ParentConfig {
    samples: usize,
    warmup: usize,
    source_revision: String,
    output: PathBuf,
    filesystem_root: Option<PathBuf>,
}

#[derive(Clone, Debug)]
struct ChildConfig {
    source: PathBuf,
    expected_aligned_source_sha256: String,
    expected_output_sha256: String,
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(MAX_HASH_HEX_BYTES);
    for byte in digest {
        let _ = std::fmt::Write::write_fmt(&mut output, format_args!("{byte:02x}"));
    }
    output
}

fn parse_hash(value: &str, flag: &str) -> Result<String, Box<dyn Error>> {
    if value.len() != MAX_HASH_HEX_BYTES || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(
            format!("{flag} must be exactly {MAX_HASH_HEX_BYTES} hexadecimal characters").into(),
        );
    }
    Ok(value.to_ascii_lowercase())
}

fn need_value(args: &[OsString], index: &mut usize, flag: &str) -> Result<String, Box<dyn Error>> {
    *index = index
        .checked_add(1)
        .ok_or("argument index overflows usize")?;
    args.get(*index)
        .ok_or_else(|| format!("{flag} requires a value").into())
        .and_then(|value| {
            value
                .to_str()
                .map(str::to_owned)
                .ok_or_else(|| format!("{flag} value must be valid UTF-8").into())
        })
}

fn parse_parent(args: &[OsString]) -> Result<ParentConfig, Box<dyn Error>> {
    let mut samples = None;
    let mut warmup = None;
    let mut source_revision = None;
    let mut output = None;
    let mut filesystem_root = None;
    let mut provider = None;
    let mut cache_state = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index]
            .to_str()
            .ok_or("cold DOCX argument must be valid UTF-8")?;
        match flag {
            "--samples" => {
                if samples.is_some() {
                    return Err("--samples was specified more than once".into());
                }
                samples = Some(
                    need_value(args, &mut index, flag)?
                        .parse::<usize>()
                        .map_err(|_| "--samples requires an unsigned decimal integer")?,
                );
            },
            "--warmup" | "--warmups" => {
                if warmup.is_some() {
                    return Err("--warmup was specified more than once".into());
                }
                warmup = Some(
                    need_value(args, &mut index, flag)?
                        .parse::<usize>()
                        .map_err(|_| "--warmup requires an unsigned decimal integer")?,
                );
            },
            "--source-revision" => {
                if source_revision.is_some() {
                    return Err("--source-revision was specified more than once".into());
                }
                let value = need_value(args, &mut index, flag)?;
                if value.len() != 40 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                    return Err(
                        "--source-revision must be exactly 40 hexadecimal characters".into(),
                    );
                }
                source_revision = Some(value);
            },
            "--output" | "--json" => {
                if output.is_some() {
                    return Err("--output was specified more than once".into());
                }
                let value = need_value(args, &mut index, flag)?;
                if value.is_empty() || value == "-" {
                    return Err("--output must be a new regular file path".into());
                }
                output = Some(PathBuf::from(value));
            },
            "--filesystem-root" => {
                if filesystem_root.is_some() {
                    return Err("--filesystem-root was specified more than once".into());
                }
                filesystem_root = Some(PathBuf::from(need_value(args, &mut index, flag)?));
            },
            "--provider" => {
                if provider.is_some() {
                    return Err("--provider was specified more than once".into());
                }
                provider = Some(need_value(args, &mut index, flag)?);
            },
            "--filesystem-cache" => {
                if cache_state.is_some() {
                    return Err("--filesystem-cache was specified more than once".into());
                }
                cache_state = Some(need_value(args, &mut index, flag)?);
            },
            other => return Err(format!("unknown cold DOCX argument: {other}").into()),
        }
        index = index
            .checked_add(1)
            .ok_or("argument index overflows usize")?;
    }
    let samples = samples.ok_or("--samples is required")?;
    let warmup = warmup.unwrap_or(0);
    if samples != 1 || warmup != 0 {
        return Err("verified cold DOCX runner requires exactly --samples 1 and --warmup 0".into());
    }
    if let Some(provider) = provider.as_deref()
        && provider != "file-cold-verified"
        && provider != "file"
    {
        return Err("verified cold DOCX runner requires provider file-cold-verified".into());
    }
    if let Some(cache_state) = cache_state.as_deref()
        && cache_state != "cold-verified"
    {
        return Err("verified cold DOCX runner requires filesystem-cache cold-verified".into());
    }
    Ok(ParentConfig {
        samples,
        warmup,
        source_revision: source_revision.ok_or("--source-revision is required")?,
        output: output.ok_or("--output is required")?,
        filesystem_root,
    })
}

fn parse_child(args: &[OsString]) -> Result<ChildConfig, Box<dyn Error>> {
    let mut source = None;
    let mut expected_aligned_source_sha256 = None;
    let mut expected_output_sha256 = None;
    let mut index = 0;
    while index < args.len() {
        let flag = args[index]
            .to_str()
            .ok_or("cold DOCX child argument must be valid UTF-8")?;
        match flag {
            "--child" => {},
            "--source" => {
                if source.is_some() {
                    return Err("--source was specified more than once".into());
                }
                source = Some(PathBuf::from(need_value(args, &mut index, flag)?));
            },
            "--expected-aligned-source-sha256" => {
                if expected_aligned_source_sha256.is_some() {
                    return Err("aligned source hash was specified more than once".into());
                }
                expected_aligned_source_sha256 =
                    Some(parse_hash(&need_value(args, &mut index, flag)?, flag)?);
            },
            "--expected-output-sha256" => {
                if expected_output_sha256.is_some() {
                    return Err("expected output hash was specified more than once".into());
                }
                expected_output_sha256 =
                    Some(parse_hash(&need_value(args, &mut index, flag)?, flag)?);
            },
            other => return Err(format!("unknown cold DOCX child argument: {other}").into()),
        }
        index = index
            .checked_add(1)
            .ok_or("argument index overflows usize")?;
    }
    Ok(ChildConfig {
        source: source.ok_or("cold DOCX child is missing --source")?,
        expected_aligned_source_sha256: expected_aligned_source_sha256
            .ok_or("cold DOCX child is missing aligned source hash")?,
        expected_output_sha256: expected_output_sha256
            .ok_or("cold DOCX child is missing expected output hash")?,
    })
}

fn prepare_case(
    corpus: crate::Corpus,
    aligned_source: Vec<u8>,
) -> Result<ColdCase, Box<dyn Error>> {
    let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(aligned_source.clone()));
    let mut expected_output = Vec::new();
    let (expected_materializations, commit) =
        crate::publish_docx_source_edit(source, &mut expected_output)?;
    if expected_materializations != 1 {
        return Err(format!(
            "verified cold DOCX preflight materialized {expected_materializations} main Parts; exactly one is required"
        )
        .into());
    }
    if expected_output == aligned_source {
        return Err("verified cold DOCX preflight produced an exact no-op".into());
    }
    crate::verify_docx_source_edit_output(&corpus, &expected_output)?;
    let source_document_xml = commit.patch().source().xml_bytes().to_vec();
    let candidate_document_xml = commit.snapshot().xml_bytes().to_vec();
    let replay_forward_verified =
        commit.patch().apply(commit.patch().source())?.xml_bytes() == commit.snapshot().xml_bytes();
    let inverse_restores_source_verified = commit
        .patch()
        .inverse()
        .apply(commit.snapshot())?
        .xml_bytes()
        == commit.patch().source().xml_bytes();
    let stale_target_refusal_verified = commit.patch().apply(commit.snapshot()).is_err();
    let foreign_package = litchi_docx::source_backed::Package::from_read_at(Arc::new(
        OwnedSource::new(expected_output.clone()),
    ))?;
    let foreign_snapshot = foreign_package.document_snapshot()?;
    let foreign_source_refusal_verified = commit.patch().apply(&foreign_snapshot).is_err();
    if !replay_forward_verified
        || !inverse_restores_source_verified
        || !stale_target_refusal_verified
        || !foreign_source_refusal_verified
    {
        return Err("verified cold DOCX patch preflight oracle failed".into());
    }
    let preflight = ColdPreflight {
        expected_materializations,
        source_document_xml_sha256: sha256_hex(&source_document_xml),
        candidate_document_xml_sha256: sha256_hex(&candidate_document_xml),
        output_exact_source_changed: true,
        semantic_reopen_verified: true,
        unchanged_media_preserved: true,
        replay_forward_verified,
        inverse_restores_source_verified,
        stale_target_refusal_verified,
        foreign_source_refusal_verified,
    };
    drop(commit);
    Ok(ColdCase {
        media_ranges: super::media_ranges(&aligned_source)?,
        expected_output_sha256: sha256_hex(&expected_output),
        expected_output,
        source_document_xml,
        candidate_document_xml,
        preflight,
        corpus,
        aligned_source,
    })
}

fn stage_aligned_file(
    bytes: &[u8],
    root_base: Option<&Path>,
) -> Result<StagedAlignedFile, Box<dyn Error>> {
    static NEXT_STAGE_ID: std::sync::atomic::AtomicU64 = std::sync::atomic::AtomicU64::new(0);
    let base = root_base
        .map(Path::to_path_buf)
        .unwrap_or_else(env::temp_dir);
    if !base.is_dir() {
        return Err(format!("filesystem root is not a directory: {}", base.display()).into());
    }
    let id = NEXT_STAGE_ID.fetch_add(1, std::sync::atomic::Ordering::Relaxed);
    let root = base.join(format!(
        "litchi-docx-edit-provider-cold-{}-{id}",
        std::process::id()
    ));
    fs::create_dir(&root)?;
    let path = root.join("source.cold-verified.docx");
    let result = (|| -> Result<(), Box<dyn Error>> {
        let mut file = OpenOptions::new()
            .create_new(true)
            .write(true)
            .open(&path)?;
        file.write_all(bytes)?;
        file.sync_all()?;
        Ok(())
    })();
    if let Err(error) = result {
        let _ = fs::remove_file(&path);
        let _ = fs::remove_dir(&root);
        return Err(error);
    }
    Ok(StagedAlignedFile { root, path })
}

fn write_report(path: &Path, report: &ColdReport) -> Result<(), Box<dyn Error>> {
    let mut output = OpenOptions::new().create_new(true).write(true).open(path)?;
    serde_json::to_writer_pretty(&mut output, report)?;
    output.write_all(b"\n")?;
    output.sync_all()?;
    Ok(())
}

fn cold_read_counter(record: super::ReadCounter) -> ColdReadCounter {
    ColdReadCounter {
        availability: record.availability.to_owned(),
        scope: record.scope.to_owned(),
        calls: record.calls,
        empty_calls: record.empty_calls,
        requested_bytes: record.requested_bytes,
        returned_bytes: record.returned_bytes,
        short_reads: record.short_reads,
        min_request_bytes: record.min_request_bytes,
        max_request_bytes: record.max_request_bytes,
        traced_ranges: record.traced_ranges,
        ranges: record.ranges,
        requested_media_overlap_bytes: record.requested_media_overlap_bytes,
        returned_media_overlap_bytes: record.returned_media_overlap_bytes,
    }
}

fn cold_sink_record(record: super::SinkRecord) -> ColdSinkRecord {
    ColdSinkRecord {
        accepted_bytes: record.accepted_bytes,
        write_calls: record.write_calls,
        largest_write: record.largest_write,
    }
}

fn logical_source_reads_positive(reads: &ColdReadCounter) -> bool {
    matches!(reads.calls, Some(calls) if calls > 0)
        && matches!(reads.requested_bytes, Some(bytes) if bytes > 0)
        && matches!(reads.returned_bytes, Some(bytes) if bytes > 0)
}

fn base_report(
    config: &ParentConfig,
    case: Option<&ColdCase>,
    filesystem_root_selected: bool,
    status: cold_verified::Status,
    samples: Vec<cold_verified::Sample>,
    rows: Vec<ColdRow>,
) -> ColdReport {
    ColdReport {
        limits: super::limits_record(),
        schema: SCHEMA,
        benchmark: BENCHMARK,
        provider_scope: PROVIDER_SCOPE,
        timing_scope: TIMING_SCOPE,
        setup_scope: SETUP_SCOPE,
        cold_claim_scope: COLD_SCOPE,
        process_metrics_scope: PROCESS_METRICS_SCOPE,
        rss_scope: RSS_SCOPE,
        provider: "file-cold-verified",
        cache_state: "cold-verified",
        filesystem_root_selected,
        source_archive_sha256: case
            .map(|case| sha256_hex(&case.corpus.archive))
            .unwrap_or_default(),
        source_archive_bytes: case.map_or(0, |case| case.corpus.archive.len()),
        aligned_source_sha256: case.map(|case| sha256_hex(&case.aligned_source)),
        aligned_source_bytes: case.map(|case| case.aligned_source.len()),
        expected_output_sha256: case.map(|case| case.expected_output_sha256.clone()),
        expected_output_bytes: case.map(|case| case.expected_output.len()),
        source_revision: config.source_revision.clone(),
        corpus: case.map(|case| case.corpus.manifest.clone()),
        preflight: case.map(|case| case.preflight.clone()),
        cold_verified_status: status,
        cold_verified_samples: samples,
        cold_verified_fincore_command: cold_verified::FINCORE_COMMAND,
        warmup: config.warmup,
        samples: config.samples,
        rows,
    }
}

fn spawn_child(
    source: &Path,
    expected_aligned_source_sha256: &str,
    expected_output_sha256: &str,
) -> Result<ColdChildOutput, Box<dyn Error>> {
    let executable = env::current_exe()?;
    let output = Command::new(executable)
        .arg("docx-edit-provider-cold")
        .arg("--child")
        .arg("--source")
        .arg(source)
        .arg("--expected-aligned-source-sha256")
        .arg(expected_aligned_source_sha256)
        .arg("--expected-output-sha256")
        .arg(expected_output_sha256)
        .stdin(Stdio::null())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .output()?;
    if !output.status.success() {
        return Err(format!(
            "verified cold DOCX child failed with {}: {}",
            output.status,
            String::from_utf8_lossy(&output.stderr)
        )
        .into());
    }
    let child: ColdChildOutput = serde_json::from_slice(&output.stdout)
        .map_err(|error| format!("verified cold DOCX child emitted invalid JSON: {error}"))?;
    if child.schema != CHILD_SCHEMA {
        return Err("verified cold DOCX child schema differs from the contract".into());
    }
    Ok(child)
}

fn run_parent(args: &[OsString]) -> Result<(), Box<dyn Error>> {
    let config = parse_parent(args)?;
    let corpus = crate::build_docx_source_edit_corpus()?;
    let source_archive_sha256 = sha256_hex(&corpus.archive);
    let page_size = match cold_verified::page_size_for_harness() {
        Ok(page_size) => page_size,
        Err(status) => {
            let report = ColdReport {
                limits: super::limits_record(),
                schema: SCHEMA,
                benchmark: BENCHMARK,
                provider_scope: PROVIDER_SCOPE,
                timing_scope: TIMING_SCOPE,
                setup_scope: SETUP_SCOPE,
                cold_claim_scope: COLD_SCOPE,
                process_metrics_scope: PROCESS_METRICS_SCOPE,
                rss_scope: RSS_SCOPE,
                provider: "file-cold-verified",
                cache_state: "cold-verified",
                filesystem_root_selected: config.filesystem_root.is_some(),
                source_archive_sha256,
                source_archive_bytes: corpus.archive.len(),
                aligned_source_sha256: None,
                aligned_source_bytes: None,
                expected_output_sha256: None,
                expected_output_bytes: None,
                source_revision: config.source_revision.clone(),
                corpus: Some(corpus.manifest.clone()),
                preflight: None,
                cold_verified_status: status,
                cold_verified_samples: vec![cold_verified::Sample::ineligible(status)],
                cold_verified_fincore_command: cold_verified::FINCORE_COMMAND,
                warmup: config.warmup,
                samples: config.samples,
                rows: Vec::new(),
            };
            return write_report(&config.output, &report);
        },
    };
    let aligned_source = match cold_verified::page_aligned_archive(&corpus.archive, page_size, true)
    {
        Ok(bytes) => bytes,
        Err(status) => {
            let report = ColdReport {
                limits: super::limits_record(),
                schema: SCHEMA,
                benchmark: BENCHMARK,
                provider_scope: PROVIDER_SCOPE,
                timing_scope: TIMING_SCOPE,
                setup_scope: SETUP_SCOPE,
                cold_claim_scope: COLD_SCOPE,
                process_metrics_scope: PROCESS_METRICS_SCOPE,
                rss_scope: RSS_SCOPE,
                provider: "file-cold-verified",
                cache_state: "cold-verified",
                filesystem_root_selected: config.filesystem_root.is_some(),
                source_archive_sha256,
                source_archive_bytes: corpus.archive.len(),
                aligned_source_sha256: None,
                aligned_source_bytes: None,
                expected_output_sha256: None,
                expected_output_bytes: None,
                source_revision: config.source_revision.clone(),
                corpus: Some(corpus.manifest.clone()),
                preflight: None,
                cold_verified_status: status,
                cold_verified_samples: vec![cold_verified::Sample::ineligible(status)],
                cold_verified_fincore_command: cold_verified::FINCORE_COMMAND,
                warmup: config.warmup,
                samples: config.samples,
                rows: Vec::new(),
            };
            return write_report(&config.output, &report);
        },
    };
    let case = prepare_case(corpus, aligned_source)?;
    let staged = stage_aligned_file(&case.aligned_source, config.filesystem_root.as_deref())?;
    let setup_proof = cold_verified::prepare(&staged.path);
    if !setup_proof.status.is_eligible() {
        let report = base_report(
            &config,
            Some(&case),
            config.filesystem_root.is_some(),
            setup_proof.status,
            vec![setup_proof],
            Vec::new(),
        );
        return write_report(&config.output, &report);
    }
    let expected_aligned_source_sha256 = sha256_hex(&case.aligned_source);
    let child = spawn_child(
        &staged.path,
        &expected_aligned_source_sha256,
        &case.expected_output_sha256,
    )?;
    let status = child.cold_verified.status;
    if let Some(error) = child.error {
        return Err(format!("verified cold DOCX child operation failed: {error}").into());
    }
    let rows = child.row.into_iter().collect::<Vec<_>>();
    if status.is_eligible() && rows.len() != 1 {
        return Err("eligible verified cold DOCX child omitted its one measured row".into());
    }
    if !status.is_eligible() && !rows.is_empty() {
        return Err("ineligible verified cold DOCX child emitted a timed row".into());
    }
    if status.is_eligible()
        && (child.cold_verified.aligned_source_sha256.as_deref()
            != Some(expected_aligned_source_sha256.as_str())
            || child.cold_verified.aligned_source_bytes != Some(case.aligned_source.len() as u64))
    {
        return Err("verified cold DOCX child proof does not bind the aligned source".into());
    }
    if let Some(row) = rows.first()
        && row.output_sha256 != case.expected_output_sha256
    {
        return Err("verified cold DOCX child output digest differs from preflight".into());
    }
    let report = base_report(
        &config,
        Some(&case),
        config.filesystem_root.is_some(),
        status,
        vec![child.cold_verified],
        rows,
    );
    write_report(&config.output, &report)
}

fn run_child(args: &[OsString]) -> Result<(), Box<dyn Error>> {
    let config = parse_child(args)?;
    // All source reads needed for expected values and every semantic/patch
    // oracle happen before the final cold probe.  The probe then evicts this
    // source and is immediately followed by the process-I/O bracket.
    let aligned_source = fs::read(&config.source)?;
    let aligned_source_sha256 = sha256_hex(&aligned_source);
    if aligned_source_sha256 != config.expected_aligned_source_sha256 {
        return Err("aligned source changed before the cold child probe".into());
    }
    let corpus = crate::build_docx_source_edit_corpus()?;
    let case = prepare_case(corpus, aligned_source)?;
    if case.expected_output_sha256 != config.expected_output_sha256 {
        return Err("cold child expected output differs from parent preflight".into());
    }

    // Reserve all measurement-side memory before the final probe.  In
    // particular, this keeps the bounded range trace and publication sink
    // allocation outside both the probe and the timed lifecycle.
    let stats = Arc::new(ReadStats::new()?);
    let maximum = u64::try_from(case.expected_output.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("cold DOCX output ceiling overflows u64")?;
    let mut sink = SequentialSink::new(maximum, SINK_MAX_WRITE_BYTES)?;
    let preparation = cold_verified::prepare(&config.source);
    if !preparation.status.is_eligible() {
        let output = ColdChildOutput {
            schema: CHILD_SCHEMA.to_owned(),
            cold_verified: preparation,
            row: None,
            error: None,
        };
        serde_json::to_writer(std::io::stdout().lock(), &output)?;
        return Ok(());
    }
    let before_io = match process_metrics::Snapshot::read() {
        Ok(snapshot) => snapshot,
        Err(_) => {
            let mut proof = preparation;
            proof.status = cold_verified::Status::IneligibleProcIoUnavailable;
            let output = ColdChildOutput {
                schema: CHILD_SCHEMA.to_owned(),
                cold_verified: proof,
                row: None,
                error: None,
            };
            serde_json::to_writer(std::io::stdout().lock(), &output)?;
            return Ok(());
        },
    };
    let allocation_region = crate::allocation_metrics::begin();
    let started = Instant::now();
    let operation =
        (|| -> Result<(usize, bool, usize, bool, SourceVersion, SourceVersion), Box<dyn Error>> {
            let file: Arc<dyn ReadAt> = Arc::new(FileSource::open(&config.source)?);
            let source: Arc<dyn ReadAt> =
                Arc::new(super::CountingReadAt::new(file, Arc::clone(&stats)));
            let source_version_before = source.version()?;
            let (materializations, commit) =
                crate::publish_docx_source_edit(Arc::clone(&source), &mut sink)?;
            let changed = commit.patch().changed();
            let operations = commit.diagnostics().operations();
            let commit_identity_verified = commit.patch().source().xml_bytes()
                == case.source_document_xml.as_slice()
                && commit.snapshot().xml_bytes() == case.candidate_document_xml.as_slice();
            let source_version_after = source.version()?;
            // Both the source-backed package and its FileSource are owned by this
            // local Arc.  Drop them before stopping the clock so provider teardown
            // remains part of the measured lifecycle.
            drop(commit);
            drop(source);
            Ok((
                materializations,
                changed,
                operations,
                commit_identity_verified,
                source_version_before,
                source_version_after,
            ))
        })();
    let elapsed = started.elapsed();
    let allocation = allocation_region.finish();
    // Complete the verifier immediately after the timed interval.  No source
    // fingerprint, output read, semantic reopen, or counter snapshot precedes
    // this proof completion.
    let after_io = process_metrics::Snapshot::read().ok();
    let process_metrics = after_io.map(|after| after.delta(before_io));
    let proof = cold_verified::complete(preparation, Some(before_io), after_io);
    let (
        materializations,
        commit_changed,
        commit_operations,
        commit_identity_verified,
        source_version_before,
        source_version_after,
    ) = match operation {
        Ok(operation) => operation,
        Err(error) => {
            let output = ColdChildOutput {
                schema: CHILD_SCHEMA.to_owned(),
                cold_verified: proof,
                row: None,
                error: Some(error.to_string()),
            };
            serde_json::to_writer(std::io::stdout().lock(), &output)?;
            return Ok(());
        },
    };
    let source_version_unchanged = source_version_before == source_version_after;
    let reads = ColdReadEvidence {
        source: cold_read_counter(super::snapshot_record(
            Some(&stats),
            "FileSource positional ReadAt calls; logical source evidence",
            &case.media_ranges,
        )?),
        scope: "FileSource positional ReadAt calls; counters are logical ranges and process read_bytes is the cold admission proof".to_owned(),
    };
    let latency_ns = u64::try_from(elapsed.as_nanos())?;
    if !proof.status.is_eligible() {
        let output = ColdChildOutput {
            schema: CHILD_SCHEMA.to_owned(),
            cold_verified: proof,
            row: None,
            error: None,
        };
        serde_json::to_writer(std::io::stdout().lock(), &output)?;
        return Ok(());
    }
    let output_exact_bytes = sink.bytes == case.expected_output;
    if !output_exact_bytes {
        return Err("cold DOCX output differs from aligned-source preflight".into());
    }
    let output_sha256 = sha256_hex(&sink.bytes);
    if output_sha256 != case.expected_output_sha256 {
        return Err("cold DOCX output digest differs from aligned-source preflight".into());
    }
    crate::verify_docx_source_edit_output(&case.corpus, &sink.bytes)?;
    let exactly_one_materialization = materializations == 1;
    let cache_load_count =
        exactly_one_materialization && materializations == case.preflight.expected_materializations;
    let logical_source_reads_positive = logical_source_reads_positive(&reads.source);
    let lifecycle_oracle_verified = commit_changed
        && commit_operations == 1
        && exactly_one_materialization
        && cache_load_count
        && source_version_unchanged
        && output_exact_bytes
        && commit_identity_verified
        && logical_source_reads_positive;
    if !lifecycle_oracle_verified {
        return Err("cold DOCX lifecycle oracle failed".into());
    }
    let row = ColdRow {
        sample_index: 0,
        child_process_id: std::process::id(),
        latency_ns,
        output_bytes: sink.bytes.len(),
        output_sha256,
        materializations,
        commit_changed,
        commit_operations,
        source_version_before: source_version_before.into(),
        source_version_after: source_version_after.into(),
        source_version_unchanged,
        reads,
        sink: cold_sink_record(sink.record()),
        cache: ColdCacheEvidence {
            successful_loads: materializations,
            expected_successful_loads: case.preflight.expected_materializations,
            exactly_one_main_part_materialization: cache_load_count,
        },
        process_metrics,
        process_metrics_scope: PROCESS_METRICS_SCOPE.to_owned(),
        rss_scope: RSS_SCOPE.to_owned(),
        oracles: OracleEvidence {
            commit_identity_verified,
            output_exact_bytes,
            semantic_reopen: case.preflight.semantic_reopen_verified,
            unchanged_media_preserved: case.preflight.unchanged_media_preserved,
            source_version_unchanged,
            cache_load_count,
            logical_source_reads_positive,
            patch_oracles: PatchOracle {
                scope: "untimed cold-child preflight commit patch oracles".to_owned(),
                replay_forward: case.preflight.replay_forward_verified,
                inverse_restores_source: case.preflight.inverse_restores_source_verified,
                stale_target_refused: case.preflight.stale_target_refusal_verified,
                foreign_source_refused: case.preflight.foreign_source_refusal_verified,
            },
        },
        allocation,
    };
    let output = ColdChildOutput {
        schema: CHILD_SCHEMA.to_owned(),
        cold_verified: proof,
        row: Some(row),
        error: None,
    };
    serde_json::to_writer(std::io::stdout().lock(), &output)?;
    Ok(())
}

/// Runs the parent or fresh-child cold evidence protocol.
pub fn run_from_args(args: impl IntoIterator<Item = OsString>) -> Result<(), Box<dyn Error>> {
    let args: Vec<OsString> = args.into_iter().collect();
    if args
        .iter()
        .any(|argument| argument == "--help" || argument == "-h")
    {
        println!(
            "docx-edit-provider-cold --samples 1 --warmup 0 --source-revision <40 hex chars> --output PATH [--filesystem-root PATH]"
        );
        return Ok(());
    }
    if args.iter().any(|argument| argument == "--child") {
        run_child(&args)
    } else {
        run_parent(&args)
    }
}

#[cfg(test)]
mod tests {
    use super::{ColdReadCounter, logical_source_reads_positive, parse_hash};

    fn reads(
        calls: Option<u64>,
        requested_bytes: Option<u64>,
        returned_bytes: Option<u64>,
    ) -> ColdReadCounter {
        ColdReadCounter {
            availability: "available".to_owned(),
            scope: "test".to_owned(),
            calls,
            empty_calls: Some(0),
            requested_bytes,
            returned_bytes,
            short_reads: Some(0),
            min_request_bytes: Some(1),
            max_request_bytes: Some(1),
            traced_ranges: calls.and_then(|value| usize::try_from(value).ok()),
            ranges: None,
            requested_media_overlap_bytes: Some(0),
            returned_media_overlap_bytes: Some(0),
        }
    }

    #[test]
    fn logical_source_gate_requires_positive_calls_and_bytes() {
        assert!(logical_source_reads_positive(&reads(
            Some(1),
            Some(8),
            Some(8)
        )));
        assert!(!logical_source_reads_positive(&reads(
            Some(0),
            Some(8),
            Some(8)
        )));
        assert!(!logical_source_reads_positive(&reads(
            Some(1),
            Some(0),
            Some(8)
        )));
        assert!(!logical_source_reads_positive(&reads(
            Some(1),
            Some(8),
            Some(0)
        )));
        assert!(!logical_source_reads_positive(&reads(None, None, None)));
    }

    #[test]
    fn aligned_hash_parser_rejects_short_or_non_hex_values() {
        assert!(parse_hash(&"a".repeat(64), "--hash").is_ok());
        assert!(parse_hash(&"a".repeat(63), "--hash").is_err());
        assert!(parse_hash(&format!("{}g", "a".repeat(63)), "--hash").is_err());
    }
}
