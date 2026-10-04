//! Paired PPTX owning-input versus materialized-ReadAt benchmark control.
//!
//! This is a standalone, opt-in process benchmark for the exact media-rich
//! corpus qualified by change 0909.  It deliberately does not add a selector
//! to the normal harness matrix: the caller runs one route per process and
//! can arrange its own AB/BA schedule and process-level measurements.

use super::{
    Corpus, CorpusManifest, PPTX_SOURCE_TEXT_BOXES_PER_SLIDE, SourceBackedPackage,
    build_pptx_source_edit_corpus, pptx_batch_replacements, pptx_multi_slide_positions, sha256_hex,
    verify_pptx_multi_slide_edit_output,
};
use litchi_core::{OwnedSource, ReadAt};
use serde::Serialize;
use std::{
    error::Error,
    fs::{self, OpenOptions},
    io::{self, Write},
    path::Path,
    sync::Arc,
    time::Instant,
};

const SCHEMA: &str = "pptx-owned-input-benchmark-v1";
const SOURCE_FILE: &str = "source.pptx";
const OUTPUT_FILE: &str = "output.pptx";
const REPORT_FILE: &str = "report.json";
const FAILURE_FILE: &str = "failure.json";

// These identities are the source fixture established by 0909.  Keeping the
// values here prevents an accidental change in the generated corpus from
// becoming an unannounced benchmark change.
const EXPECTED_SOURCE_BYTES: usize = 17_017_139;
const EXPECTED_SOURCE_SHA256: &str =
    "61b2b99083ca27ebd37955db600955e3f41289b93dba71951983164239eff757";
const EXPECTED_ENTRY_COUNT: usize = 229;
const EXPECTED_MEMBER_COUNT: usize = 445;
const EXPECTED_UNCOMPRESSED_BYTES: usize = 17_568_429;
const MAX_SINK_WRITE_BYTES: usize = 64 * 1024;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Route {
    Materialized,
    Owned,
}

impl Route {
    pub(crate) fn parse(value: &str) -> Result<Self, Box<dyn Error>> {
        match value {
            "materialized" => Ok(Self::Materialized),
            "owned" => Ok(Self::Owned),
            _ => Err(format!(
                "unknown PPTX owning benchmark route {value:?}; expected materialized or owned"
            )
            .into()),
        }
    }

    const fn name(self) -> &'static str {
        match self {
            Self::Materialized => "materialized",
            Self::Owned => "owned",
        }
    }

    const fn implementation(self) -> &'static str {
        match self {
            Self::Materialized => {
                "OwnedSource(Vec) -> SourceBackedPackage::from_read_at -> into_opc_package -> Package::from_opc_package"
            },
            Self::Owned => "Package::from_vec(Vec)",
        }
    }
}

#[derive(Debug)]
struct PhaseDurations {
    clone_ns: u64,
    ingress_ns: u64,
    snapshot_ns: u64,
    edit_commit_ns: u64,
    publish_ns: u64,
}

#[derive(Debug)]
struct TimedIteration {
    phases: PhaseDurations,
    total_ns: u64,
}

/// A preallocated sequential consumer shared by every route and iteration.
/// Its allocation and capacity reservation happen before the primary timer;
/// the bounded writes themselves are part of publication for both routes.
#[derive(Debug)]
struct PreallocatedSink {
    bytes: Vec<u8>,
    max_bytes: usize,
    max_write: usize,
    write_calls: usize,
    largest_write: usize,
}

impl PreallocatedSink {
    fn new(max_bytes: usize, max_write: usize) -> Result<Self, Box<dyn Error>> {
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(max_bytes).map_err(|error| {
            io::Error::other(format!(
                "cannot reserve PPTX benchmark sink budget: {error}"
            ))
        })?;
        Ok(Self {
            bytes,
            max_bytes,
            max_write,
            write_calls: 0,
            largest_write: 0,
        })
    }

    fn reset(&mut self) {
        self.bytes.clear();
        self.write_calls = 0;
        self.largest_write = 0;
    }
}

impl Write for PreallocatedSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.len() > self.max_write {
            return Err(io::Error::other(
                "PPTX benchmark sink write exceeds the configured maximum",
            ));
        }
        let remaining = self
            .max_bytes
            .checked_sub(self.bytes.len())
            .ok_or_else(|| io::Error::other("PPTX benchmark sink byte budget overflow"))?;
        if bytes.len() > remaining {
            return Err(io::Error::other("PPTX benchmark sink byte budget exceeded"));
        }
        self.bytes.extend_from_slice(bytes);
        self.write_calls = self
            .write_calls
            .checked_add(1)
            .ok_or_else(|| io::Error::other("PPTX benchmark sink write count overflow"))?;
        self.largest_write = self.largest_write.max(bytes.len());
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug, Serialize)]
struct Artifact {
    file: &'static str,
    bytes: usize,
    sha256: String,
}

#[derive(Debug, Serialize)]
struct Sample {
    index: usize,
    clone_ns: u64,
    ingress_ns: u64,
    snapshot_ns: u64,
    edit_commit_ns: u64,
    publish_ns: u64,
    total_ns: u64,
    output_sha256: String,
    output_bytes: usize,
    sink_write_calls: usize,
    sink_largest_write: usize,
}

#[derive(Debug, Serialize)]
struct Gates {
    source_identity_verified: bool,
    source_shape_verified: bool,
    semantic_output_verified: bool,
    stable_output_verified: bool,
    sink_verified: bool,
    source_immutability_verified: bool,
}

#[derive(Debug, Serialize)]
struct BenchmarkReport {
    schema: &'static str,
    status: &'static str,
    route: &'static str,
    implementation: &'static str,
    warmup_iterations: usize,
    measured_samples: usize,
    corpus: CorpusManifest,
    source: Artifact,
    output: Artifact,
    sink_capacity_bytes: usize,
    sink_max_write_bytes: usize,
    selected_slide_positions: Vec<usize>,
    replacement_count: usize,
    timing_scope: &'static str,
    phase_scope: &'static str,
    teardown_scope: &'static str,
    verification_scope: &'static str,
    performance_claim: &'static str,
    samples: Vec<Sample>,
    gates: Gates,
}

#[derive(Debug, Serialize)]
struct FailureReport {
    schema: &'static str,
    status: &'static str,
    route: &'static str,
    warmup_iterations: usize,
    measured_samples: usize,
    error: String,
    corpus: Option<CorpusManifest>,
}

fn elapsed_ns(start: Instant) -> Result<u64, Box<dyn Error>> {
    u64::try_from(start.elapsed().as_nanos())
        .map_err(|_| "PPTX benchmark elapsed duration exceeds u64 nanoseconds".into())
}

fn write_new_bytes(path: &Path, bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let mut file = OpenOptions::new().write(true).create_new(true).open(path)?;
    file.write_all(bytes)?;
    file.flush()?;
    file.sync_all()?;
    Ok(())
}

fn write_new_json<T: Serialize>(path: &Path, value: &T) -> Result<(), Box<dyn Error>> {
    let mut file = OpenOptions::new().write(true).create_new(true).open(path)?;
    serde_json::to_writer_pretty(&mut file, value)?;
    file.write_all(b"\n")?;
    file.flush()?;
    file.sync_all()?;
    Ok(())
}

fn check_primary_corpus(corpus: &Corpus) -> Result<(), Box<dyn Error>> {
    if corpus.manifest.generator != super::PPTX_SOURCE_EDIT_CORPUS_GENERATOR
        || corpus.archive.len() != EXPECTED_SOURCE_BYTES
        || sha256_hex(&corpus.archive) != EXPECTED_SOURCE_SHA256
        || corpus.manifest.entry_count != EXPECTED_ENTRY_COUNT
        || corpus.manifest.archive_member_count != EXPECTED_MEMBER_COUNT
        || corpus.manifest.uncompressed_payload_bytes != EXPECTED_UNCOMPRESSED_BYTES
        || corpus.manifest.shape != "media-rich"
        || corpus.manifest.entry_bytes != super::PPTX_SOURCE_MEDIA_ENTRY_BYTES
    {
        return Err("PPTX benchmark source differs from the frozen 0909 corpus".into());
    }
    Ok(())
}

/// Run one route against the borrowed corpus bytes.
///
/// The total timer starts before the `Vec` clone and stops only after this
/// function's inner scope has dropped the package, transaction, source owner,
/// output buffer, and other route-local state. The caller-provided sink is
/// retained after that boundary. All routes execute the same
/// opened transaction, 64 shape-text replacements, commit, publication, and
/// `to_bytes` path after ingress.
fn timed_iteration(
    source: &[u8],
    route: Route,
    sink: &mut PreallocatedSink,
) -> Result<TimedIteration, Box<dyn Error>> {
    // Resetting a retained, already-reserved sink is outside the primary
    // timer. This keeps capacity and allocator state matched across routes.
    sink.reset();
    let total_start = Instant::now();
    let phases = {
        let clone_start = Instant::now();
        let input = std::hint::black_box(source.to_vec());
        let clone_ns = elapsed_ns(clone_start)?;

        let ingress_start = Instant::now();
        let mut package = match route {
            Route::Owned => litchi_pptx::Package::from_vec(input)?,
            Route::Materialized => {
                let source: Arc<dyn ReadAt> = Arc::new(OwnedSource::new(input));
                let package = SourceBackedPackage::from_read_at(source)?;
                let opc = package.into_opc_package()?;
                litchi_pptx::Package::from_opc_package(opc)?
            },
        };
        let ingress_ns = elapsed_ns(ingress_start)?;

        let snapshot_start = Instant::now();
        let mut edit = package.opened_presentation_transaction()?;
        let snapshot_ns = elapsed_ns(snapshot_start)?;

        let edit_start = Instant::now();
        for slide_position in pptx_multi_slide_positions() {
            let texts = (0..PPTX_SOURCE_TEXT_BOXES_PER_SLIDE)
                .map(|shape| super::semantic_pptx_text(slide_position, shape, true))
                .collect::<Vec<_>>();
            let replacements = pptx_batch_replacements(&texts);
            if edit.set_shape_texts(slide_position, &replacements)?
                != PPTX_SOURCE_TEXT_BOXES_PER_SLIDE
            {
                return Err("PPTX owning benchmark changed an unexpected shape count".into());
            }
        }
        let commit = edit.commit()?;
        if !commit.is_changed() {
            return Err("PPTX owning benchmark produced an unchanged commit".into());
        }
        let edit_commit_ns = elapsed_ns(edit_start)?;

        let publish_start = Instant::now();
        package.apply_opened_presentation_commit(commit)?;
        let output = package.to_bytes()?;
        for chunk in output.chunks(MAX_SINK_WRITE_BYTES) {
            sink.write_all(chunk)?;
        }
        let publish_ns = elapsed_ns(publish_start)?;

        PhaseDurations {
            clone_ns,
            ingress_ns,
            snapshot_ns,
            edit_commit_ns,
            publish_ns,
        }
    };
    let total_ns = elapsed_ns(total_start)?;
    Ok(TimedIteration { phases, total_ns })
}

fn verify_output(
    corpus: &Corpus,
    output: &[u8],
    expected_digest: &mut Option<String>,
    expected_bytes: &mut Option<usize>,
) -> Result<String, Box<dyn Error>> {
    verify_pptx_multi_slide_edit_output(corpus, output)?;
    let digest = sha256_hex(output);
    if let Some(expected) = expected_digest.as_ref() {
        if expected != &digest || expected_bytes.as_ref() != Some(&output.len()) {
            return Err("PPTX benchmark output changed between iterations".into());
        }
    } else {
        *expected_digest = Some(digest.clone());
        *expected_bytes = Some(output.len());
    }
    Ok(digest)
}

fn artifact(file: &'static str, bytes: &[u8]) -> Artifact {
    Artifact {
        file,
        bytes: bytes.len(),
        sha256: sha256_hex(bytes),
    }
}

fn run(
    route: Route,
    warmup_iterations: usize,
    samples: usize,
    directory: &Path,
) -> Result<(), Box<dyn Error>> {
    if samples == 0 {
        return Err("PPTX benchmark requires at least one measured sample".into());
    }
    if warmup_iterations > 100_000 || samples > 100_000 {
        return Err("PPTX benchmark warmups and samples are capped at 100000".into());
    }
    if directory.exists() {
        return Err("PPTX benchmark output directory already exists".into());
    }
    fs::create_dir(directory)?;

    let mut corpus_manifest = None;
    let result = (|| -> Result<(), Box<dyn Error>> {
        let corpus = build_pptx_source_edit_corpus()?;
        check_primary_corpus(&corpus)?;
        corpus_manifest = Some(corpus.manifest.clone());
        let source_digest_before = sha256_hex(&corpus.archive);
        let sink_capacity_bytes = corpus
            .archive
            .len()
            .checked_mul(2)
            .and_then(|bytes| bytes.checked_add(MAX_SINK_WRITE_BYTES))
            .ok_or("PPTX benchmark sink capacity overflows usize")?;
        let mut sink = PreallocatedSink::new(sink_capacity_bytes, MAX_SINK_WRITE_BYTES)?;

        let mut expected_digest = None;
        let mut expected_bytes = None;
        for _ in 0..warmup_iterations {
            timed_iteration(&corpus.archive, route, &mut sink)?;
            verify_output(
                &corpus,
                &sink.bytes,
                &mut expected_digest,
                &mut expected_bytes,
            )?;
            if sink.write_calls == 0 || sink.largest_write > MAX_SINK_WRITE_BYTES {
                return Err("PPTX benchmark warmup sink result is invalid".into());
            }
        }

        let mut measured = Vec::with_capacity(samples);
        for index in 0..samples {
            let iteration = timed_iteration(&corpus.archive, route, &mut sink)?;
            let digest = verify_output(
                &corpus,
                &sink.bytes,
                &mut expected_digest,
                &mut expected_bytes,
            )?;
            if sink.write_calls == 0 || sink.largest_write > MAX_SINK_WRITE_BYTES {
                return Err("PPTX benchmark measured sink result is invalid".into());
            }
            measured.push(Sample {
                index,
                clone_ns: iteration.phases.clone_ns,
                ingress_ns: iteration.phases.ingress_ns,
                snapshot_ns: iteration.phases.snapshot_ns,
                edit_commit_ns: iteration.phases.edit_commit_ns,
                publish_ns: iteration.phases.publish_ns,
                total_ns: iteration.total_ns,
                output_sha256: digest,
                output_bytes: sink.bytes.len(),
                sink_write_calls: sink.write_calls,
                sink_largest_write: sink.largest_write,
            });
        }

        let source_artifact = artifact(SOURCE_FILE, &corpus.archive);
        let output_artifact = artifact(OUTPUT_FILE, &sink.bytes);
        let sink_verified = measured.iter().all(|sample| {
            sample.sink_write_calls > 0
                && sample.sink_largest_write <= MAX_SINK_WRITE_BYTES
                && sample.output_bytes > 0
        });
        write_new_bytes(&directory.join(SOURCE_FILE), &corpus.archive)?;
        write_new_bytes(&directory.join(OUTPUT_FILE), &sink.bytes)?;
        let report = BenchmarkReport {
            schema: SCHEMA,
            status: "completed",
            route: route.name(),
            implementation: route.implementation(),
            warmup_iterations,
            measured_samples: samples,
            corpus: corpus.manifest.clone(),
            source: source_artifact,
            output: output_artifact,
            sink_capacity_bytes,
            sink_max_write_bytes: MAX_SINK_WRITE_BYTES,
            selected_slide_positions: pptx_multi_slide_positions().to_vec(),
            replacement_count: pptx_multi_slide_positions()
                .len()
                .checked_mul(PPTX_SOURCE_TEXT_BOXES_PER_SLIDE)
                .ok_or("PPTX benchmark replacement count overflows usize")?,
            timing_scope: "one fresh Vec clone plus route ingress, opened transaction snapshot, 64 text replacements and commit, commit publication, Package::to_bytes, and the identical bounded sink writes; one route per process",
            phase_scope: "clone_ns is source.to_vec; ingress_ns begins after the clone and ends after route package construction; snapshot_ns is opened_presentation_transaction; edit_commit_ns is all selected shape replacements plus commit; publish_ns is apply_opened_presentation_commit, to_bytes, and identical 64 KiB sequential sink writes",
            teardown_scope: "total_ns stops after the timed body drops the package, transaction, source owner, output Vec, and route-local state; the preallocated sink is retained after the boundary",
            verification_scope: "semantic output, digest stability, source identity, sink counters, and artifact writes occur outside measured timers",
            performance_claim: "none; descriptive paired phase evidence only, with no inferred speedup, allocation result, or production recommendation",
            samples: measured,
            gates: Gates {
                source_identity_verified: corpus.archive.len() == EXPECTED_SOURCE_BYTES
                    && sha256_hex(&corpus.archive) == EXPECTED_SOURCE_SHA256,
                source_shape_verified: corpus.manifest.entry_count == EXPECTED_ENTRY_COUNT
                    && corpus.manifest.archive_member_count == EXPECTED_MEMBER_COUNT
                    && corpus.manifest.uncompressed_payload_bytes == EXPECTED_UNCOMPRESSED_BYTES,
                semantic_output_verified: true,
                stable_output_verified: true,
                sink_verified,
                source_immutability_verified: sha256_hex(&corpus.archive) == source_digest_before,
            },
        };
        write_new_json(&directory.join(REPORT_FILE), &report)?;
        Ok(())
    })();

    if let Err(error) = result {
        let failure = FailureReport {
            schema: SCHEMA,
            status: "failed",
            route: route.name(),
            warmup_iterations,
            measured_samples: samples,
            error: error.to_string(),
            corpus: corpus_manifest,
        };
        let _ = write_new_json(&directory.join(FAILURE_FILE), &failure);
        return Err(error);
    }
    Ok(())
}

pub(crate) fn run_from_args(
    route: &str,
    warmup_iterations: usize,
    samples: usize,
    directory: &Path,
) -> Result<(), Box<dyn Error>> {
    run(Route::parse(route)?, warmup_iterations, samples, directory)
}
