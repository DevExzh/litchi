//! Small managed source-backed DOCX paragraph-edit benchmark.
//!
//! The repeated mode stages sorted replacements by calling
//! Edit::replace_paragraph_text once per paragraph. The batch mode stages the
//! same replacement list with Edit::replace_body_paragraph_texts. Both modes
//! use the public managed source-backed package constructor and the same
//! publication lifecycle. They are reported as two API paths; rows do not
//! authorize a same-API before/after claim.
//!
//! The fixture contains a direct main-story body, one deterministic 256 KiB
//! media member, and one deterministic 256 KiB opaque member. Source/output
//! identities, all paragraph text, untouched member payloads, patch
//! forward/inverse oracles, and managed resource gauges are checked outside
//! the timed lifecycle.

extern crate self as litchi_perf_baseline;

#[allow(
    dead_code,
    reason = "The shared observer includes harness-only identity helpers."
)]
#[path = "../../../../../../tools/perf-baseline/src/allocation_metrics.rs"]
pub mod allocation_metrics;
#[cfg(feature = "allocator-metrics")]
#[path = "../../../../../../tools/perf-baseline/src/bin/support/counting_allocator.rs"]
mod counting_allocator;

use std::error::Error;
use std::fmt::Write as _;
use std::fs;
use std::io::{self, Cursor, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::path::{Path, PathBuf};
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};
use std::time::Instant;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, FileSource, Limits, OwnedSource,
    Position, ReadAt, Resource, SourceVersion,
};
use litchi_docx::document::ParagraphTextReplacement;
use litchi_docx::{ReadLimits, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, SourceCacheDiagnostics, SourceCacheLimits};
use sha2::{Digest, Sha256};
use serde::Serialize;
use soapberry_zip::office::StreamingArchiveWriter;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const MAIN: &str = "word/document.xml";
const MEDIA: &str = "word/media/image1.png";
const OPAQUE: &str = "word/opaque.bin";
const MEDIA_BYTES: usize = 256 * 1024;
const OPAQUE_BYTES: usize = 256 * 1024;

const MEMORY_LIMIT: u64 = 64 * 1024 * 1024;
const INPUT_LIMIT: u64 = 64 * 1024 * 1024;
const OUTPUT_LIMIT: u64 = 64 * 1024 * 1024;
const OBJECT_LIMIT: u64 = 1_000_000;
const DEPTH_LIMIT: u64 = 1024;
const WORK_LIMIT: u64 = 1 << 50;
const IN_FLIGHT_BYTES: u64 = 64 * 1024 * 1024;
const CACHE_BYTES: usize = 8 * 1024 * 1024;
const CACHE_ENTRIES: usize = 128;

type AnyResult<T> = Result<T, Box<dyn Error>>;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum SourceKind {
    Owned,
    File,
}

impl SourceKind {
    const fn name(self) -> &'static str {
        match self {
            Self::Owned => "owned",
            Self::File => "file",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum EditMode {
    Repeated,
    Batch,
}

impl EditMode {
    const fn name(self) -> &'static str {
        match self {
            Self::Repeated => "repeated",
            Self::Batch => "batch",
        }
    }

    const fn api_name(self) -> &'static str {
        match self {
            Self::Repeated => "replace_paragraph_text",
            Self::Batch => "replace_body_paragraph_texts",
        }
    }
}

#[derive(Debug)]
struct Config {
    paragraphs: usize,
    replacements: usize,
    source: SourceKind,
    mode: EditMode,
    samples: usize,
    warmups: usize,
    repeats: usize,
    output: PathBuf,
    artifact_dir: PathBuf,
}

#[derive(Debug)]
struct Fixture {
    archive: Vec<u8>,
    archive_sha256: String,
    media: Vec<u8>,
    opaque: Vec<u8>,
    original_text: Vec<String>,
    replacement_text: Vec<String>,
    replacements: Vec<ParagraphTextReplacement>,
}

#[derive(Clone, Copy, Debug)]
struct BudgetSnapshot {
    memory: u64,
    input: u64,
    output: u64,
    objects: u64,
    work: u64,
}

impl BudgetSnapshot {
    fn from_budget(budget: &Budget) -> Self {
        Self {
            memory: budget.used(Resource::Memory),
            input: budget.used(Resource::InputBytes),
            output: budget.used(Resource::OutputBytes),
            objects: budget.used(Resource::Objects),
            work: budget.used(Resource::Work),
        }
    }
}

#[derive(Debug)]
struct ReadCounters {
    calls: AtomicU64,
    requested: AtomicU64,
    returned: AtomicU64,
    zero_length: AtomicU64,
}

impl ReadCounters {
    const fn new() -> Self {
        Self {
            calls: AtomicU64::new(0),
            requested: AtomicU64::new(0),
            returned: AtomicU64::new(0),
            zero_length: AtomicU64::new(0),
        }
    }

    fn snapshot(&self) -> (u64, u64, u64, u64) {
        (
            self.calls.load(Ordering::Relaxed),
            self.requested.load(Ordering::Relaxed),
            self.returned.load(Ordering::Relaxed),
            self.zero_length.load(Ordering::Relaxed),
        )
    }
}

struct CountingSource {
    inner: Arc<dyn ReadAt>,
    counters: ReadCounters,
}

impl CountingSource {
    fn new(inner: Arc<dyn ReadAt>) -> Self {
        Self {
            inner,
            counters: ReadCounters::new(),
        }
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.counters.calls.fetch_add(1, Ordering::Relaxed);
        self.counters.requested.fetch_add(
            u64::try_from(output.len()).unwrap_or(u64::MAX),
            Ordering::Relaxed,
        );
        if output.is_empty() {
            self.counters.zero_length.fetch_add(1, Ordering::Relaxed);
        }
        let result = self.inner.read_at(offset, output);
        if let Ok(read) = result {
            self.counters
                .returned
                .fetch_add(u64::try_from(read).unwrap_or(u64::MAX), Ordering::Relaxed);
        }
        result
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

#[derive(Debug)]
struct Preflight {
    expected_output: Vec<u8>,
    expected_output_sha256: String,
    unmanaged_forward_ok: bool,
    unmanaged_inverse_ok: bool,
    managed_forward_ok: bool,
    managed_inverse_ok: bool,
    managed_output_exact_ok: bool,
}

#[derive(Debug)]
struct Sample {
    repeat: usize,
    ordinal: usize,
    warmup: bool,
    elapsed_ns: u64,
    open_ns: u64,
    edit_ns: u64,
    commit_ns: u64,
    publish_ns: u64,
    allocation_sample: allocation_metrics::Sample,
    drop_ns: u64,
    output_bytes: usize,
    output_sha256: String,
    source_version_before: SourceVersion,
    source_version_after: SourceVersion,
    source_calls: u64,
    source_requested_bytes: u64,
    source_returned_bytes: u64,
    source_zero_length_calls: u64,
    cache_before: SourceCacheDiagnostics,
    cache_live: SourceCacheDiagnostics,
    budget_before: BudgetSnapshot,
    budget_live: BudgetSnapshot,
    budget_after_drop: BudgetSnapshot,
    semantic_ok: bool,
    raw_untouched_member_payloads_ok: bool,
    output_exact_ok: bool,
    memory_released: bool,
    objects_released: bool,
    managed_preflight_forward_ok: bool,
    managed_preflight_inverse_ok: bool,
    input_monotonic: bool,
    work_monotonic: bool,
    source_version_unchanged: bool,
}

struct ArtifactDirGuard {
    path: PathBuf,
}

impl ArtifactDirGuard {
    fn new(path: &Path) -> AnyResult<Self> {
        if path.exists() {
            return Err(format!(
                "--artifact-dir must be a fresh absent path: {}",
                path.display()
            )
            .into());
        }
        if let Some(parent) = path.parent()
            && !parent.as_os_str().is_empty()
        {
            fs::create_dir_all(parent)?;
        }
        fs::create_dir(path)?;
        Ok(Self {
            path: path.to_owned(),
        })
    }
}

impl Drop for ArtifactDirGuard {
    fn drop(&mut self) {
        let _ = fs::remove_dir_all(&self.path);
    }
}

fn main() -> AnyResult<()> {
    #[cfg(feature = "allocator-metrics")]
    allocation_metrics::enable();
    let config = parse_args(std::env::args_os().skip(1))?;
    if config.output.starts_with(&config.artifact_dir) {
        return Err("--output must be outside the owned --artifact-dir".into());
    }
    let _artifact_dir = ArtifactDirGuard::new(&config.artifact_dir)?;
    let fixture = build_fixture(config.paragraphs, config.replacements)?;
    let fixture_path = config.artifact_dir.join(format!(
        "fixture-p{}-k{}.docx",
        config.paragraphs, config.replacements
    ));
    fs::write(&fixture_path, &fixture.archive)?;
    let preflight = build_preflight(&fixture, config.mode)?;
    if !preflight.unmanaged_forward_ok
        || !preflight.unmanaged_inverse_ok
        || !preflight.managed_forward_ok
        || !preflight.managed_inverse_ok
        || !preflight.managed_output_exact_ok
    {
        return Err("preflight patch forward/inverse or managed output oracle failed".into());
    }

    let mut rows = Vec::new();
    for repeat in 0..config.repeats {
        for ordinal in 0..config.warmups {
            rows.push(run_sample_and_emit(
                &config,
                &fixture,
                &preflight,
                &fixture_path,
                repeat,
                ordinal,
                true,
            )?);
        }
        for ordinal in 0..config.samples {
            rows.push(run_sample_and_emit(
                &config,
                &fixture,
                &preflight,
                &fixture_path,
                repeat,
                ordinal,
                false,
            )?);
        }
    }
    write_csv(&config, &fixture, &preflight, &fixture_path, &rows)?;
    Ok(())
}

fn parse_args<I>(args: I) -> AnyResult<Config>
where
    I: IntoIterator<Item = std::ffi::OsString>,
{
    let mut paragraphs = None;
    let mut replacements = None;
    let mut source = SourceKind::Owned;
    let mut mode = EditMode::Repeated;
    let mut samples = 30;
    let mut warmups = 3;
    let mut repeats = 1;
    let mut output = None;
    let mut artifact_dir = PathBuf::from("target/managed-paragraph-batch");
    let values: Vec<std::ffi::OsString> = args.into_iter().collect();
    let mut index = 0;
    while index < values.len() {
        let flag = values[index].to_string_lossy();
        if flag == "--help" || flag == "-h" {
            print_help();
            std::process::exit(0);
        }
        let next = |index: &mut usize, name: &str| -> AnyResult<std::ffi::OsString> {
            *index = (*index)
                .checked_add(1)
                .ok_or_else(|| format!("{name} argument index overflow"))?;
            values
                .get(*index)
                .cloned()
                .ok_or_else(|| Box::<dyn Error>::from(format!("missing value for {name}")))
        };
        match flag.as_ref() {
            "--paragraphs" => {
                paragraphs = Some(parse_count(
                    &next(&mut index, "--paragraphs")?,
                    "--paragraphs",
                )?);
            },
            "--replacements" => {
                replacements = Some(parse_count(
                    &next(&mut index, "--replacements")?,
                    "--replacements",
                )?);
            },
            "--source" => {
                source = match next(&mut index, "--source")?.to_string_lossy().as_ref() {
                    "owned" => SourceKind::Owned,
                    "file" => SourceKind::File,
                    value => return Err(format!("unknown --source value {value:?}").into()),
                };
            },
            "--mode" => {
                mode = match next(&mut index, "--mode")?.to_string_lossy().as_ref() {
                    "repeated" => EditMode::Repeated,
                    "batch" => EditMode::Batch,
                    value => return Err(format!("unknown --mode value {value:?}").into()),
                };
            },
            "--samples" => {
                samples = parse_count(&next(&mut index, "--samples")?, "--samples")?;
            },
            "--warmups" => {
                warmups = parse_count(&next(&mut index, "--warmups")?, "--warmups")?;
            },
            "--repeats" => {
                repeats = parse_count(&next(&mut index, "--repeats")?, "--repeats")?;
            },
            "--output" => output = Some(PathBuf::from(next(&mut index, "--output")?)),
            "--artifact-dir" => {
                artifact_dir = PathBuf::from(next(&mut index, "--artifact-dir")?);
            },
            value if value.starts_with('-') => return Err(format!("unknown option {value}").into()),
            value => return Err(format!("unexpected positional argument {value}").into()),
        }
        index = index.checked_add(1).ok_or("argument index overflow")?;
    }
    let paragraphs = paragraphs.ok_or("--paragraphs is required")?;
    let replacements = replacements.ok_or("--replacements is required")?;
    if paragraphs == 0 || paragraphs > 4096 {
        return Err("--paragraphs must be between 1 and 4096".into());
    }
    if replacements == 0 || replacements > paragraphs {
        return Err("--replacements must be between 1 and --paragraphs".into());
    }
    if samples == 0 || warmups > 10_000 || samples > 10_000 || repeats == 0 || repeats > 100 {
        return Err("samples, warmups, or repeats is outside the bounded harness range".into());
    }
    Ok(Config {
        paragraphs,
        replacements,
        source,
        mode,
        samples,
        warmups,
        repeats,
        output: output.ok_or("--output is required")?,
        artifact_dir,
    })
}

fn parse_count(value: &std::ffi::OsStr, name: &str) -> AnyResult<usize> {
    value
        .to_str()
        .ok_or_else(|| Box::<dyn Error>::from(format!("{name} must be UTF-8")))?
        .parse::<usize>()
        .map_err(|error| Box::<dyn Error>::from(format!("invalid {name}: {error}")))
}

fn print_help() {
    println!(
        "managed_paragraph_batch_perf --paragraphs N --replacements K --source owned|file --mode repeated|batch --samples N --warmups N --repeats N --output PATH --artifact-dir PATH"
    );
}

fn build_fixture(paragraph_count: usize, replacement_count: usize) -> AnyResult<Fixture> {
    let mut media = Vec::with_capacity(MEDIA_BYTES);
    let mut opaque = Vec::with_capacity(OPAQUE_BYTES);
    for index in 0..MEDIA_BYTES {
        media.push((index as u32).wrapping_mul(131).wrapping_add(17) as u8);
    }
    for index in 0..OPAQUE_BYTES {
        opaque.push((index as u32).wrapping_mul(197).wrapping_add(43) as u8);
    }
    let mut document = String::new();
    write!(document, r#"<w:document xmlns:w="{W}"><w:body>"#)?;
    let mut original_text = Vec::with_capacity(paragraph_count);
    for index in 0..paragraph_count {
        let text = format!("paragraph-{index:04}-original");
        write!(document, r#"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"#)?;
        original_text.push(text);
    }
    document.push_str("</w:body></w:document>");

    let mut replacement_text = original_text.clone();
    let mut replacements = Vec::with_capacity(replacement_count);
    for index in 0..replacement_count {
        let position = ((index + 1) * paragraph_count) / (replacement_count + 1);
        let text = format!("paragraph-{position:04}-replacement-k{replacement_count:02}");
        replacement_text[position] = text.clone();
        replacements.push(ParagraphTextReplacement::new(Position::new(position), text));
    }

    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/{MAIN}" ContentType="{document_content_type}"/></Types>"#,
        document_content_type = ct::WML_DOCUMENT_MAIN,
    );
    let relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rIdDocument" Type="{office_document}" Target="{MAIN}"/></Relationships>"#,
        office_document = rt::OFFICE_DOCUMENT,
    );
    let mut writer = StreamingArchiveWriter::new();
    writer.write_stored("[Content_Types].xml", content_types.as_bytes())?;
    writer.write_stored("_rels/.rels", relationships.as_bytes())?;
    writer.write_stored(MEDIA, &media)?;
    writer.write_stored(OPAQUE, &opaque)?;
    writer.write_stored(MAIN, document.as_bytes())?;
    let archive = writer.finish_to_bytes()?;
    let archive_sha256 = sha256_hex(&archive);
    Ok(Fixture {
        archive,
        archive_sha256,
        media,
        opaque,
        original_text,
        replacement_text,
        replacements,
    })
}

fn build_preflight(fixture: &Fixture, mode: EditMode) -> AnyResult<Preflight> {
    let package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(fixture.archive.clone())))?;
    let mut edit = package.edit_document()?;
    stage_edit(&mut edit, fixture, mode)?;
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    let forward = patch.apply(patch.source())?;
    let inverse = patch.inverse().apply(commit.snapshot())?;
    let forward_ok = forward.xml_bytes() == commit.snapshot().xml_bytes();
    let inverse_ok = inverse.xml_bytes() == patch.source().xml_bytes();
    let mut expected_output = Vec::new();
    package.publish_document_commit_to_stream(&mut expected_output, &commit)?;
    let expected_output_sha256 = sha256_hex(&expected_output);
    if !verify_output(fixture, &expected_output, &fixture.replacement_text)? {
        return Err("unmanaged preflight semantic oracle failed".into());
    }
    let (managed_forward_ok, managed_inverse_ok, managed_output_exact_ok) =
        build_managed_preflight(fixture, mode, &expected_output)?;
    Ok(Preflight {
        expected_output,
        expected_output_sha256,
        unmanaged_forward_ok: forward_ok,
        unmanaged_inverse_ok: inverse_ok,
        managed_forward_ok,
        managed_inverse_ok,
        managed_output_exact_ok,
    })
}

fn build_managed_preflight(
    fixture: &Fixture,
    mode: EditMode,
    expected_output: &[u8],
) -> AnyResult<(bool, bool, bool)> {
    let (budget, _cancellation_source, context) = managed_context()?;
    let budget_before = BudgetSnapshot::from_budget(&budget);
    let package =
        source_backed::Package::from_read_at_with_limits_and_cache_limits_and_execution_context(
            Arc::new(OwnedSource::new(fixture.archive.clone())),
            ReadLimits::default(),
            SourceCacheLimits::new(CACHE_BYTES, CACHE_ENTRIES)?,
            context,
        )?;
    let mut edit = package.edit_document()?;
    stage_edit(&mut edit, fixture, mode)?;
    let commit = edit.commit()?;
    let patch = commit.patch().clone();
    let forward = patch.apply(patch.source())?;
    let inverse = patch.inverse().apply(commit.snapshot())?;
    let forward_ok = forward.xml_bytes() == commit.snapshot().xml_bytes();
    let inverse_ok = inverse.xml_bytes() == patch.source().xml_bytes();
    drop(forward);
    drop(inverse);
    let mut output = Vec::with_capacity(expected_output.len());
    let published = package.publish_document_commit_to_stream(&mut output, &commit)?;
    drop(published);
    drop(commit);
    drop(patch);
    let output_exact_ok = output == expected_output;
    let budget_after_drop = BudgetSnapshot::from_budget(&budget);
    if budget_after_drop.memory != budget_before.memory
        || budget_after_drop.objects != budget_before.objects
    {
        return Err(format!(
            "managed preflight retained memory/objects: {} / {}",
            budget_after_drop.memory, budget_after_drop.objects
        )
        .into());
    }
    Ok((forward_ok, inverse_ok, output_exact_ok))
}

fn stage_edit(
    edit: &mut litchi_docx::document::Edit,
    fixture: &Fixture,
    mode: EditMode,
) -> AnyResult<()> {
    match mode {
        EditMode::Repeated => {
            for replacement in &fixture.replacements {
                edit.replace_paragraph_text(replacement.position(), replacement.text())?;
            }
        },
        EditMode::Batch => {
            edit.replace_body_paragraph_texts(&fixture.replacements)?;
        },
    }
    Ok(())
}

fn managed_context() -> AnyResult<(Budget, CancellationSource, ExecutionContext)> {
    let budget = Budget::root(
        "managed-paragraph-batch-perf",
        Limits::new(
            MEMORY_LIMIT,
            INPUT_LIMIT,
            OUTPUT_LIMIT,
            OBJECT_LIMIT,
            DEPTH_LIMIT,
            WORK_LIMIT,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(IN_FLIGHT_BYTES).ok_or("in-flight limit must be nonzero")?,
        0,
    )?;
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    Ok((budget, cancellation_source, context))
}

fn source_for_sample(
    fixture: &Fixture,
    source_kind: SourceKind,
    fixture_path: &Path,
) -> AnyResult<Arc<CountingSource>> {
    let inner: Arc<dyn ReadAt> = match source_kind {
        SourceKind::Owned => Arc::new(OwnedSource::new(fixture.archive.clone())),
        SourceKind::File => Arc::new(FileSource::open(fixture_path)?),
    };
    Ok(Arc::new(CountingSource::new(inner)))
}

#[derive(Serialize)]
struct AllocationSampleRecord<'a> {
    tag: &'static str,
    scope: &'static str,
    case: String,
    repeat: usize,
    ordinal: usize,
    warmup: bool,
    #[serde(rename = "allocationSample")]
    allocation_sample: &'a allocation_metrics::Sample,
}

fn allocation_case(config: &Config) -> String {
    format!(
        "p{}-k{}-{}-{}",
        config.paragraphs,
        config.replacements,
        config.source.name(),
        config.mode.name()
    )
}

fn emit_allocation_sample(config: &Config, sample: &Sample) -> AnyResult<()> {
    let record = AllocationSampleRecord {
        tag: "allocationSample",
        scope: "publish_document_commit_to_stream_method_only_before_returned_snapshot_drop",
        case: allocation_case(config),
        repeat: sample.repeat,
        ordinal: sample.ordinal,
        warmup: sample.warmup,
        allocation_sample: &sample.allocation_sample,
    };
    let stdout = io::stdout();
    let mut output = stdout.lock();
    serde_json::to_writer(&mut output, &record)?;
    output.write_all(b"\n")?;
    output.flush()?;
    Ok(())
}

fn run_sample_and_emit(
    config: &Config,
    fixture: &Fixture,
    preflight: &Preflight,
    fixture_path: &Path,
    repeat: usize,
    ordinal: usize,
    warmup: bool,
) -> AnyResult<Sample> {
    let sample = run_sample(
        config,
        fixture,
        preflight,
        fixture_path,
        repeat,
        ordinal,
        warmup,
    )?;
    emit_allocation_sample(config, &sample)?;
    Ok(sample)
}

fn run_sample(
    config: &Config,
    fixture: &Fixture,
    preflight: &Preflight,
    fixture_path: &Path,
    repeat: usize,
    ordinal: usize,
    warmup: bool,
) -> AnyResult<Sample> {
    let source = source_for_sample(fixture, config.source, fixture_path)?;
    let source_version_before = source.version()?;
    let (budget, _cancellation_source, context) = managed_context()?;
    let budget_before = BudgetSnapshot::from_budget(&budget);
    let mut output = Vec::with_capacity(preflight.expected_output.len());
    let operation_started = Instant::now();

    let phase_started = Instant::now();
    let package =
        source_backed::Package::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source.clone(),
            ReadLimits::default(),
            SourceCacheLimits::new(CACHE_BYTES, CACHE_ENTRIES)?,
            context,
        )?;
    let open_ns = elapsed_ns(phase_started)?;
    let cache_before = package.cache_diagnostics();

    let phase_started = Instant::now();
    let mut edit = package.edit_document()?;
    stage_edit(&mut edit, fixture, config.mode)?;
    let edit_ns = elapsed_ns(phase_started)?;

    let phase_started = Instant::now();
    let commit = edit.commit()?;
    let commit_ns = elapsed_ns(phase_started)?;
    let cache_live = package.cache_diagnostics();
    let budget_live = BudgetSnapshot::from_budget(&budget);

    let phase_started = Instant::now();
    let allocation_region = allocation_metrics::begin();
    let published = package.publish_document_commit_to_stream(&mut output, &commit)?;
    let allocation_sample = allocation_region
        .finish()
        .unwrap_or_else(allocation_metrics::unavailable_sample);
    drop(published);
    let publish_ns = elapsed_ns(phase_started)?;

    let phase_started = Instant::now();
    drop(commit);
    let drop_ns = elapsed_ns(phase_started)?;
    let elapsed_ns = elapsed_ns(operation_started)?;
    let budget_after_drop = BudgetSnapshot::from_budget(&budget);
    let source_version_after = source.version()?;
    let (source_calls, source_requested_bytes, source_returned_bytes, source_zero_length_calls) =
        source.counters.snapshot();
    drop(_cancellation_source);

    let output_sha256 = sha256_hex(&output);
    let semantic_ok = verify_output(fixture, &output, &fixture.replacement_text)?;
    let raw_untouched_member_payloads_ok = verify_untouched_members(fixture, &output)?;
    let output_exact_ok = output == preflight.expected_output;
    let memory_released = budget_after_drop.memory == budget_before.memory;
    let objects_released = budget_after_drop.objects == budget_before.objects;
    let input_monotonic =
        budget_before.input <= budget_live.input && budget_live.input <= budget_after_drop.input;
    let work_monotonic =
        budget_before.work <= budget_live.work && budget_live.work <= budget_after_drop.work;
    let source_version_unchanged = source_version_before == source_version_after;
    if !cache_live.budget_managed {
        return Err("managed sample did not report a managed cache budget".into());
    }
    if !semantic_ok
        || !raw_untouched_member_payloads_ok
        || !output_exact_ok
        || !memory_released
        || !objects_released
        || !input_monotonic
        || !work_monotonic
        || !source_version_unchanged
    {
        return Err(format!(
            "sample verification failed: semantic={semantic_ok} raw_member_payloads={raw_untouched_member_payloads_ok} exact={output_exact_ok} memory={memory_released} objects={objects_released} input={input_monotonic} work={work_monotonic} source={source_version_unchanged}"
        )
        .into());
    }
    Ok(Sample {
        repeat,
        ordinal,
        warmup,
        elapsed_ns,
        open_ns,
        edit_ns,
        commit_ns,
        publish_ns,
        allocation_sample,
        drop_ns,
        output_bytes: output.len(),
        output_sha256,
        source_version_before,
        source_version_after,
        source_calls,
        source_requested_bytes,
        source_returned_bytes,
        source_zero_length_calls,
        cache_before,
        cache_live,
        budget_before,
        budget_live,
        budget_after_drop,
        semantic_ok,
        raw_untouched_member_payloads_ok,
        output_exact_ok,
        memory_released,
        objects_released,
        managed_preflight_forward_ok: preflight.managed_forward_ok,
        managed_preflight_inverse_ok: preflight.managed_inverse_ok,
        input_monotonic,
        work_monotonic,
        source_version_unchanged,
    })
}

fn verify_output(fixture: &Fixture, output: &[u8], expected: &[String]) -> AnyResult<bool> {
    let package = source_backed::Package::from_reader(Cursor::new(output))?;
    let snapshot = package.document_snapshot()?;
    if snapshot.paragraph_count() != fixture.original_text.len() {
        return Ok(false);
    }
    for (index, expected_text) in expected.iter().enumerate() {
        let paragraph = snapshot
            .paragraph(Position::new(index))
            .ok_or_else(|| format!("missing output paragraph {index}"))?;
        if paragraph.text()?.as_str() != expected_text {
            return Ok(false);
        }
    }
    Ok(true)
}

fn verify_untouched_members(fixture: &Fixture, output: &[u8]) -> AnyResult<bool> {
    Ok(part_blob(output, MEDIA)? == fixture.media && part_blob(output, OPAQUE)? == fixture.opaque)
}

fn part_blob(bytes: &[u8], name: &str) -> AnyResult<Vec<u8>> {
    let package = OpcPackage::from_bytes(bytes)?;
    let part = package.get_part(&PackURI::new(format!("/{name}"))?)?;
    Ok(part.blob().to_vec())
}

fn elapsed_ns(started: Instant) -> AnyResult<u64> {
    u64::try_from(started.elapsed().as_nanos())
        .map_err(|_| "timed duration does not fit in u64 nanoseconds".into())
}

fn sha256_hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut output = String::with_capacity(64);
    for byte in digest {
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn write_csv(
    config: &Config,
    fixture: &Fixture,
    preflight: &Preflight,
    fixture_path: &Path,
    rows: &[Sample],
) -> AnyResult<()> {
    if let Some(parent) = config.output.parent() {
        if !parent.as_os_str().is_empty() {
            fs::create_dir_all(parent)?;
        }
    }
    let mut file = fs::File::create(&config.output)?;
    writeln!(
        file,
        "schema,version,api_path,comparison_scope,paragraphs,replacements,source,mode,fixture_sha256,fixture_bytes,fixture_name,expected_output_sha256,expected_output_bytes,repeat,ordinal,warmup,elapsed_ns,open_ns,edit_ns,commit_ns,publish_ns,drop_ns,output_bytes,output_sha256,source_version_before_id,source_version_before_revision,source_version_after_id,source_version_after_revision,source_version_unchanged,source_read_calls,source_requested_bytes,source_returned_bytes,source_zero_length_calls,cache_before_cold_loads,cache_before_successful_loads,cache_before_hits,cache_live_cold_loads,cache_live_successful_loads,cache_live_hits,cache_live_retained_bytes,cache_live_retained_entries,cache_live_in_flight_loads,budget_managed,budget_before_memory,budget_live_memory,budget_after_memory,budget_before_input,budget_live_input,budget_after_input,budget_before_output,budget_live_output,budget_after_output,budget_before_objects,budget_live_objects,budget_after_objects,budget_before_work,budget_live_work,budget_after_work,memory_released,objects_released,input_monotonic,work_monotonic,semantic_ok,raw_untouched_ok,raw_untouched_member_payloads_ok,output_exact_ok,managed_preflight_forward_ok,managed_preflight_inverse_ok,forward_ok,inverse_ok,unmanaged_preflight_forward_ok,unmanaged_preflight_inverse_ok"
    )?;
    for row in rows {
        let fields = [
            csv_field("managed_paragraph_batch_perf_v1"),
            csv_field(1),
            csv_field(config.mode.api_name()),
            csv_field(
                "descriptive_same_fixture_alternative_api_path_no_same_api_before_after_claim",
            ),
            csv_field(config.paragraphs),
            csv_field(config.replacements),
            csv_field(config.source.name()),
            csv_field(config.mode.name()),
            csv_field(&fixture.archive_sha256),
            csv_field(fixture.archive.len()),
            csv_field(
                fixture_path
                    .file_name()
                    .and_then(std::ffi::OsStr::to_str)
                    .unwrap_or("fixture.docx"),
            ),
            csv_field(&preflight.expected_output_sha256),
            csv_field(preflight.expected_output.len()),
            csv_field(row.repeat),
            csv_field(row.ordinal),
            csv_field(row.warmup),
            csv_field(row.elapsed_ns),
            csv_field(row.open_ns),
            csv_field(row.edit_ns),
            csv_field(row.commit_ns),
            csv_field(row.publish_ns),
            csv_field(row.drop_ns),
            csv_field(row.output_bytes),
            csv_field(&row.output_sha256),
            csv_field(row.source_version_before.id()),
            csv_field(row.source_version_before.revision()),
            csv_field(row.source_version_after.id()),
            csv_field(row.source_version_after.revision()),
            csv_field(row.source_version_unchanged),
            csv_field(row.source_calls),
            csv_field(row.source_requested_bytes),
            csv_field(row.source_returned_bytes),
            csv_field(row.source_zero_length_calls),
            csv_field(row.cache_before.cold_loads),
            csv_field(row.cache_before.successful_loads),
            csv_field(row.cache_before.hits),
            csv_field(row.cache_live.cold_loads),
            csv_field(row.cache_live.successful_loads),
            csv_field(row.cache_live.hits),
            csv_field(row.cache_live.retained_bytes),
            csv_field(row.cache_live.retained_entries),
            csv_field(row.cache_live.in_flight_loads),
            csv_field(row.cache_live.budget_managed),
            csv_field(row.budget_before.memory),
            csv_field(row.budget_live.memory),
            csv_field(row.budget_after_drop.memory),
            csv_field(row.budget_before.input),
            csv_field(row.budget_live.input),
            csv_field(row.budget_after_drop.input),
            csv_field(row.budget_before.output),
            csv_field(row.budget_live.output),
            csv_field(row.budget_after_drop.output),
            csv_field(row.budget_before.objects),
            csv_field(row.budget_live.objects),
            csv_field(row.budget_after_drop.objects),
            csv_field(row.budget_before.work),
            csv_field(row.budget_live.work),
            csv_field(row.budget_after_drop.work),
            csv_field(row.memory_released),
            csv_field(row.objects_released),
            csv_field(row.input_monotonic),
            csv_field(row.work_monotonic),
            csv_field(row.semantic_ok),
            csv_field(row.raw_untouched_member_payloads_ok),
            csv_field(row.raw_untouched_member_payloads_ok),
            csv_field(row.output_exact_ok),
            csv_field(row.managed_preflight_forward_ok),
            csv_field(row.managed_preflight_inverse_ok),
            csv_field(row.managed_preflight_forward_ok),
            csv_field(row.managed_preflight_inverse_ok),
            csv_field(preflight.unmanaged_forward_ok),
            csv_field(preflight.unmanaged_inverse_ok),
        ];
        writeln!(file, "{}", fields.join(","))?;
    }
    Ok(())
}

fn csv_field(value: impl std::fmt::Display) -> String {
    let value = value.to_string();
    if value
        .bytes()
        .any(|byte| matches!(byte, b',' | b'\"' | b'\n' | b'\r'))
    {
        format!("\"{}\"", value.replace('\"', "\"\""))
    } else {
        value
    }
}
