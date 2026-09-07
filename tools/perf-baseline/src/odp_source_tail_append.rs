//! Opt-in source-backed ODP tail-append publication evidence.
//!
//! The selector measures the advanced source publication path owned by the ODP
//! crate.  Its candidate is captured once before the measured iterations and
//! is used only as an output oracle; the timed path opens a positional source,
//! proves the bounded append, prepares the insertion plan, and publishes to a
//! hashing discard sink.  It does not measure or claim the retained-result
//! contract of the ordinary owned `Snapshot`/`Commit`/`Patch` path.
//!
//! The formal capture contract is carried by `OdpSourceTailAppendSummary`:
//! source and candidate archive/content hashes, slide count and actual tail
//! title/body, complete member identities and manifest/raw-preservation gates,
//! proof insertion offset/page/source version, publication-report bytes and
//! source-version vectors, positional read calls/bytes, and per-sample output
//! digest/length gates.  A capture may compare timing vectors only after all
//! of those untimed oracle fields and the stale-source refusal gate pass.

use super::odp_existing_append::{
    self, ODP_EXISTING_APPEND_CORPUS_GENERATOR, OPAQUE_PATH, OdpExistingAppendCorpus,
    OdpExistingAppendMemberIdentity,
};
use super::{
    Case, CaseResult, HashingDiscardSink, SemanticShape, SinkSummary, SourceSummary,
    allocation_metrics, deterministic_sink_summary, elapsed_ns, iteration_count, operation_metrics,
    process_metrics, record_elapsed, sha256_hex, statistics,
};
use litchi_core::{ReadAt, SourceVersion};
use litchi_odf_common::core::{OwnedPackage, SourceBackedPackage, SourceContentPublicationOptions};
use serde::Serialize;
use std::{
    error::Error,
    fs, io,
    path::Path,
    sync::{
        Arc,
        atomic::{AtomicU64, Ordering},
    },
    time::Duration,
};

/// This selector deliberately reuses the existing append corpus generator so
/// source and owned control rows can be compared on the same deterministic
/// 64/4096/8192 source archives.
pub(crate) const ODP_SOURCE_TAIL_APPEND_CORPUS_GENERATOR: &str =
    ODP_EXISTING_APPEND_CORPUS_GENERATOR;

static NEXT_MEASURE_SOURCE_ID: AtomicU64 = AtomicU64::new(1);

/// Positional source adapter used by the source-backed append lane.
///
/// The archive bytes are owned before timing begins.  Only logical `ReadAt`
/// calls and returned bytes are counted; ZIP, XML, and sink accounting stays
/// in the owning operation layers.
#[derive(Debug)]
pub(crate) struct MeasureBytesSource {
    bytes: Arc<[u8]>,
    id: u64,
    revision: AtomicU64,
    read_calls: AtomicU64,
    read_bytes: AtomicU64,
}

impl MeasureBytesSource {
    pub(crate) fn new(bytes: Arc<[u8]>) -> Self {
        Self {
            bytes,
            id: NEXT_MEASURE_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
            revision: AtomicU64::new(0),
            read_calls: AtomicU64::new(0),
            read_bytes: AtomicU64::new(0),
        }
    }

    fn metrics(&self) -> (u64, u64) {
        (
            self.read_calls.load(Ordering::Relaxed),
            self.read_bytes.load(Ordering::Relaxed),
        )
    }

    fn matches_bytes(&self, expected: &[u8]) -> bool {
        self.bytes.as_ref() == expected
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::AcqRel);
    }
}

impl ReadAt for MeasureBytesSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "source length overflows u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.read_calls.fetch_add(1, Ordering::Relaxed);
        let offset = usize::try_from(offset).map_err(|_| {
            io::Error::new(
                io::ErrorKind::InvalidInput,
                "source offset does not fit usize",
            )
        })?;
        if offset >= self.bytes.len() || output.is_empty() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        self.read_bytes.fetch_add(
            u64::try_from(count).map_err(|_| {
                io::Error::new(
                    io::ErrorKind::InvalidData,
                    "source read length overflows u64",
                )
            })?,
            Ordering::Relaxed,
        );
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            self.id,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

#[derive(Clone, Debug, Serialize)]
pub(crate) struct OdpSourceTailAppendSummary {
    pub(crate) role: &'static str,
    pub(crate) implementation: &'static str,
    pub(crate) timing_scope: &'static str,
    pub(crate) performance_claim: &'static str,
    pub(crate) corpus_generator: &'static str,
    pub(crate) shape: &'static str,
    pub(crate) source_archive_sha256: String,
    pub(crate) source_archive_bytes: usize,
    pub(crate) output_archive_sha256: String,
    pub(crate) output_archive_bytes: usize,
    pub(crate) source_content_xml_sha256: String,
    pub(crate) source_content_xml_bytes: usize,
    pub(crate) output_content_xml_sha256: String,
    pub(crate) output_content_xml_bytes: usize,
    pub(crate) source_slide_count: usize,
    pub(crate) output_slide_count: usize,
    pub(crate) append_count: usize,
    pub(crate) appended_title: String,
    pub(crate) appended_body: String,
    pub(crate) appended_title_sha256: String,
    pub(crate) appended_body_sha256: String,
    pub(crate) opaque_member_path: &'static str,
    pub(crate) opaque_bytes: usize,
    pub(crate) opaque_sha256: String,
    pub(crate) opaque_member_compressed_bytes: u64,
    pub(crate) opaque_member_compressed_sha256: String,
    pub(crate) source_members: Vec<OdpExistingAppendMemberIdentity>,
    pub(crate) output_members: Vec<OdpExistingAppendMemberIdentity>,
    pub(crate) source_member_count: usize,
    pub(crate) output_member_count: usize,
    pub(crate) manifest_entry_count: usize,
    pub(crate) source_manifest_bindings_verified: bool,
    pub(crate) output_manifest_bindings_verified: bool,
    pub(crate) source_member_raw_preservation_verified: bool,
    pub(crate) source_semantic_reopen_verified: bool,
    pub(crate) output_semantic_reopen_verified: bool,
    pub(crate) append_exactly_one_verified: bool,
    pub(crate) source_unchanged_verified: bool,
    pub(crate) stale_source_refusal_verified: bool,
    pub(crate) text_contract: &'static str,
    pub(crate) source_semantic_sha256: String,
    pub(crate) output_semantic_sha256: String,
    pub(crate) source_order_sha256: String,
    pub(crate) output_order_sha256: String,
    pub(crate) source_text_projection_sha256: String,
    pub(crate) output_text_projection_sha256: String,
    pub(crate) source_text_projection_bytes: usize,
    pub(crate) output_text_projection_bytes: usize,
    pub(crate) proof_slide_count: usize,
    pub(crate) proof_content_xml_bytes: u64,
    pub(crate) proof_insert_at: u64,
    pub(crate) proof_page_name: String,
    pub(crate) proof_source_version_id: u64,
    pub(crate) proof_source_version_revision: u64,
    pub(crate) publication_report_bytes: Vec<u64>,
    pub(crate) publication_report_source_version_id: Vec<u64>,
    pub(crate) publication_report_source_version_revision: Vec<u64>,
    pub(crate) runtime_proof_source_version_id: Vec<u64>,
    pub(crate) runtime_proof_source_version_revision: Vec<u64>,
    pub(crate) source_read_calls: Vec<u64>,
    pub(crate) source_read_bytes: Vec<u64>,
    pub(crate) runtime_output_digest_verified: bool,
    pub(crate) runtime_sink_length_verified: bool,
    pub(crate) lifecycle_ns: Vec<u64>,
    pub(crate) open_ns: Vec<u64>,
    pub(crate) append_plan_ns: Vec<u64>,
    pub(crate) publication_ns: Vec<u64>,
    pub(crate) output_sha256: Vec<String>,
}

#[derive(Debug)]
pub(crate) struct OdpSourceTailAppendCorpus {
    pub(crate) source: OdpExistingAppendCorpus,
    output_sha256: String,
    output_bytes: usize,
    source_content_xml_sha256: String,
    output_content_xml_sha256: String,
    output_content_xml_bytes: usize,
    output_members: Vec<OdpExistingAppendMemberIdentity>,
    source_semantic_sha256: String,
    output_semantic_sha256: String,
    source_order_sha256: String,
    output_order_sha256: String,
    source_text_projection_sha256: String,
    output_text_projection_sha256: String,
    source_text_projection_bytes: usize,
    output_text_projection_bytes: usize,
    source_manifest_bindings_verified: bool,
    output_manifest_bindings_verified: bool,
    source_member_raw_preservation_verified: bool,
    manifest_entry_count: usize,
    source_slide_count: usize,
    output_slide_count: usize,
    appended_title: String,
    appended_body: String,
    appended_title_sha256: String,
    appended_body_sha256: String,
    source_unchanged_verified: bool,
    proof_slide_count: usize,
    proof_content_xml_bytes: u64,
    proof_insert_at: u64,
    proof_page_name: String,
    proof_source_version_id: u64,
    proof_source_version_revision: u64,
    stale_source_refusal_verified: bool,
}

fn source_arc(bytes: &[u8]) -> Arc<[u8]> {
    Arc::<[u8]>::from(bytes.to_vec())
}

fn open_source_package(
    source: &Arc<MeasureBytesSource>,
) -> Result<Arc<SourceBackedPackage>, Box<dyn Error>> {
    let package = SourceBackedPackage::from_read_at(Arc::clone(source) as Arc<dyn ReadAt>)?;
    Ok(Arc::new(package))
}

fn append_options() -> SourceContentPublicationOptions {
    SourceContentPublicationOptions::new()
}

const fn expected_source_slide_count(shape: SemanticShape) -> usize {
    match shape {
        SemanticShape::Tiny => 64,
        SemanticShape::Medium => 4_096,
        SemanticShape::Large => 8_192,
    }
}

fn source_member_raw_preservation(
    source_members: &[OdpExistingAppendMemberIdentity],
    output_members: &[OdpExistingAppendMemberIdentity],
) -> Result<bool, Box<dyn Error>> {
    for source in source_members {
        if source.path == "content.xml" {
            continue;
        }
        let output = odp_existing_append::find_member(output_members, &source.path)?;
        if source != output {
            return Ok(false);
        }
    }
    Ok(source_members.len() == output_members.len())
}

fn candidate_semantics(
    shape: SemanticShape,
    source_bytes: &[u8],
    output_bytes: &[u8],
    source_members: &[OdpExistingAppendMemberIdentity],
    output_members: &[OdpExistingAppendMemberIdentity],
) -> Result<CandidateFacts, Box<dyn Error>> {
    odp_existing_append::verify_append_output(
        shape,
        source_bytes,
        output_bytes,
        source_members,
        output_members,
    )?;
    let source_slides = odp_existing_append::semantic_slides(source_bytes)?;
    let output_slides = odp_existing_append::semantic_slides(output_bytes)?;
    let expected_source_count = expected_source_slide_count(shape);
    if source_slides.len() != expected_source_count
        || output_slides.len() != expected_source_count + 1
        || output_slides[..expected_source_count] != source_slides[..]
    {
        return Err(
            "ODP source append candidate changed the source semantic slide sequence".into(),
        );
    }
    let source_semantic_sha256 = odp_existing_append::semantic_digest(&source_slides)?;
    let output_semantic_sha256 = odp_existing_append::semantic_digest(&output_slides)?;
    let source_order_sha256 = odp_existing_append::order_digest(&source_slides)?;
    let output_order_sha256 = odp_existing_append::order_digest(&output_slides)?;
    let source_text = odp_existing_append::text_projection(&source_slides)?;
    let output_text = odp_existing_append::text_projection(&output_slides)?;
    let (Some(appended_title), appended_body) = output_slides
        .last()
        .cloned()
        .ok_or("ODP source append output has no final slide")?
    else {
        return Err("ODP source append output final slide has no title".into());
    };
    let expected_title = odp_existing_append::appended_title(shape);
    let expected_body = odp_existing_append::appended_body(shape);
    if appended_title != expected_title || appended_body != expected_body {
        return Err(
            "ODP source append candidate title/body differs from the actual request".into(),
        );
    }
    let appended_title_sha256 = sha256_hex(appended_title.as_bytes());
    let appended_body_sha256 = sha256_hex(appended_body.as_bytes());
    let source_package = OwnedPackage::from_bytes(source_bytes.to_vec())?;
    let output_package = OwnedPackage::from_bytes(output_bytes.to_vec())?;
    let source_content_xml = source_package.get_file("content.xml")?;
    let output_content_xml = output_package.get_file("content.xml")?;
    let source_manifest_bindings_verified =
        odp_existing_append::expected_manifest_bindings(&source_package)?;
    let output_manifest_bindings_verified =
        odp_existing_append::expected_manifest_bindings(&output_package)?;
    let source_manifest_entry_count = source_package.package()?.manifest().entries.len();
    let output_manifest_entry_count = output_package.package()?.manifest().entries.len();
    if source_manifest_entry_count != output_manifest_entry_count {
        return Err("ODP source append manifest entry count changed".into());
    }
    if !source_manifest_bindings_verified || !output_manifest_bindings_verified {
        return Err("ODP source append manifest bindings differ from the contract".into());
    }
    let source_member_raw_preservation_verified =
        source_member_raw_preservation(source_members, output_members)?;
    if !source_member_raw_preservation_verified {
        return Err("ODP source append changed an untouched member's raw identity".into());
    }
    if !odp_existing_append::expected_archive_shape(source_bytes)?
        || !odp_existing_append::expected_archive_shape(output_bytes)?
    {
        return Err("ODP source append archive topology differs from the contract".into());
    }
    let source_opaque = odp_existing_append::find_member(source_members, OPAQUE_PATH)?;
    let output_opaque = odp_existing_append::find_member(output_members, OPAQUE_PATH)?;
    if source_opaque != output_opaque {
        return Err("ODP source append opaque member identity changed".into());
    }
    Ok(CandidateFacts {
        source_content_xml_sha256: sha256_hex(&source_content_xml),
        output_content_xml_sha256: sha256_hex(&output_content_xml),
        output_content_xml_bytes: output_content_xml.len(),
        source_semantic_sha256,
        output_semantic_sha256,
        source_order_sha256,
        output_order_sha256,
        source_text_projection_sha256: sha256_hex(&source_text),
        output_text_projection_sha256: sha256_hex(&output_text),
        source_text_projection_bytes: source_text.len(),
        output_text_projection_bytes: output_text.len(),
        source_slide_count: source_slides.len(),
        output_slide_count: output_slides.len(),
        appended_title,
        appended_body,
        appended_title_sha256,
        appended_body_sha256,
        source_manifest_bindings_verified,
        output_manifest_bindings_verified,
        source_member_raw_preservation_verified,
        manifest_entry_count: source_manifest_entry_count,
    })
}

#[derive(Debug)]
struct CandidateFacts {
    source_content_xml_sha256: String,
    output_content_xml_sha256: String,
    output_content_xml_bytes: usize,
    source_semantic_sha256: String,
    output_semantic_sha256: String,
    source_order_sha256: String,
    output_order_sha256: String,
    source_text_projection_sha256: String,
    output_text_projection_sha256: String,
    source_text_projection_bytes: usize,
    output_text_projection_bytes: usize,
    source_slide_count: usize,
    output_slide_count: usize,
    appended_title: String,
    appended_body: String,
    appended_title_sha256: String,
    appended_body_sha256: String,
    source_manifest_bindings_verified: bool,
    output_manifest_bindings_verified: bool,
    source_member_raw_preservation_verified: bool,
    manifest_entry_count: usize,
}

#[derive(Debug)]
struct SourceTailAppendCandidateCapture {
    output: Vec<u8>,
    facts: CandidateFacts,
    proof_slide_count: usize,
    proof_content_xml_bytes: u64,
    proof_insert_at: u64,
    proof_page_name: String,
    proof_source_version: SourceVersion,
}

const ODP_SOURCE_TAIL_APPEND_FIXTURE_SCHEMA: &str = "litchi-odp-source-tail-append-fixtures-v1";

#[derive(Debug, Serialize)]
struct OdpSourceTailAppendFixtureManifest {
    schema: &'static str,
    generator: &'static str,
    package_format: &'static str,
    artifact_count: usize,
    artifacts: Vec<OdpSourceTailAppendFixtureArtifact>,
}

#[derive(Debug, Serialize)]
struct OdpSourceTailAppendFixtureArtifact {
    shape: &'static str,
    source_file: String,
    candidate_file: String,
    title: String,
    body: String,
    source_archive_sha256: String,
    source_archive_bytes: usize,
    candidate_archive_sha256: String,
    candidate_archive_bytes: usize,
    source_content_xml_sha256: String,
    source_content_xml_bytes: usize,
    candidate_content_xml_sha256: String,
    candidate_content_xml_bytes: usize,
    source_slide_count: usize,
    candidate_slide_count: usize,
    append_count: usize,
    source_semantic_sha256: String,
    candidate_semantic_sha256: String,
    source_order_sha256: String,
    candidate_order_sha256: String,
    source_text_projection_sha256: String,
    candidate_text_projection_sha256: String,
    source_text_projection_bytes: usize,
    candidate_text_projection_bytes: usize,
    source_members: Vec<OdpExistingAppendMemberIdentity>,
    candidate_members: Vec<OdpExistingAppendMemberIdentity>,
    source_member_count: usize,
    candidate_member_count: usize,
    manifest_entry_count: usize,
    proof_slide_count: usize,
    proof_content_xml_bytes: u64,
    proof_insert_at: u64,
    proof_page_name: String,
    source_manifest_bindings_verified: bool,
    candidate_manifest_bindings_verified: bool,
    source_member_raw_preservation_verified: bool,
    append_exactly_one_verified: bool,
    source_unchanged_verified: bool,
    stale_source_refusal_verified: bool,
}

fn capture_source_tail_append_candidate(
    shape: SemanticShape,
    source: &OdpExistingAppendCorpus,
) -> Result<SourceTailAppendCandidateCapture, Box<dyn Error>> {
    let source_bytes = source.corpus.archive.as_slice();
    let title = odp_existing_append::appended_title(shape);
    let body = odp_existing_append::appended_body(shape);
    let options = append_options();
    let provider = Arc::new(MeasureBytesSource::new(source_arc(source_bytes)));
    let package = open_source_package(&provider)?;
    let edit = litchi_odp::SourceBackedTailAppendEdit::new(package, title, body)?;
    let plan = edit.plan(&options)?;
    let proof = plan.proof().clone();
    let mut output = Vec::new();
    let report = plan.write_to(&mut output, options.clone())?;
    let report_bytes = report.bytes();
    let proof_source_version = proof.source_version();
    let report_source_version = report.source_version();
    let provider_source_version = provider.version()?;
    if report_source_version != proof_source_version
        || proof_source_version != provider_source_version
    {
        return Err(
            "ODP source append proof, publication report, and provider source versions differ"
                .into(),
        );
    }
    if proof.slide_count() == 0 {
        return Err("ODP source append proof has no source slides".into());
    }
    if proof.content_bytes() != u64::try_from(source.source_content_xml.len())? {
        return Err("ODP source append proof content length differs from source XML".into());
    }
    let source_metrics = provider.metrics();
    if report_bytes != u64::try_from(output.len())? {
        return Err("ODP source append publication report length differs from output".into());
    }
    if source_metrics.1 == 0 || output.is_empty() {
        return Err("ODP source append candidate did not read and publish bytes".into());
    }
    let output_members = odp_existing_append::member_identities(&output)?;
    let facts = candidate_semantics(
        shape,
        source_bytes,
        &output,
        &source.source_members,
        &output_members,
    )?;
    if provider.metrics().1 == 0 || !provider.matches_bytes(source_bytes) {
        return Err("ODP source append candidate source counters or bytes changed".into());
    }
    drop(plan);
    drop(provider);
    Ok(SourceTailAppendCandidateCapture {
        output,
        facts,
        proof_slide_count: proof.slide_count(),
        proof_content_xml_bytes: proof.content_bytes(),
        proof_insert_at: proof.insert_at(),
        proof_page_name: proof.page_name().to_owned(),
        proof_source_version,
    })
}

fn build_odp_source_tail_append_corpus_with_output(
    shape: SemanticShape,
) -> Result<(OdpSourceTailAppendCorpus, Vec<u8>), Box<dyn Error>> {
    let source = odp_existing_append::build_odp_existing_append_corpus(shape)?;
    if source.corpus.manifest.generator != ODP_SOURCE_TAIL_APPEND_CORPUS_GENERATOR {
        return Err("ODP source append received the wrong deterministic corpus".into());
    }
    let source_bytes = source.corpus.archive.as_slice();
    let title = odp_existing_append::appended_title(shape);
    let body = odp_existing_append::appended_body(shape);
    let options = append_options();

    let stale_provider = Arc::new(MeasureBytesSource::new(source_arc(source_bytes)));
    let stale_package = open_source_package(&stale_provider)?;
    let stale_edit =
        litchi_odp::SourceBackedTailAppendEdit::new(stale_package, title.clone(), body.clone())?;
    let stale_plan = stale_edit.plan(&options)?;
    stale_provider.bump_revision();
    let mut stale_output = Vec::new();
    let stale_refusal_verified = stale_plan
        .write_to(&mut stale_output, options.clone())
        .is_err()
        && stale_output.is_empty();
    drop(stale_plan);
    drop(stale_output);
    drop(stale_provider);
    if !stale_refusal_verified {
        return Err("ODP source append accepted a stale source publication plan".into());
    }

    let capture = capture_source_tail_append_candidate(shape, &source)?;
    let SourceTailAppendCandidateCapture {
        output,
        facts,
        proof_slide_count,
        proof_content_xml_bytes,
        proof_insert_at,
        proof_page_name,
        proof_source_version,
    } = capture;
    let output_sha256 = sha256_hex(&output);
    let output_bytes = output.len();
    let corpus = OdpSourceTailAppendCorpus {
        source,
        output_sha256,
        output_bytes,
        source_content_xml_sha256: facts.source_content_xml_sha256,
        output_content_xml_sha256: facts.output_content_xml_sha256,
        output_content_xml_bytes: facts.output_content_xml_bytes,
        output_members: odp_existing_append::member_identities(&output)?,
        source_semantic_sha256: facts.source_semantic_sha256,
        output_semantic_sha256: facts.output_semantic_sha256,
        source_order_sha256: facts.source_order_sha256,
        output_order_sha256: facts.output_order_sha256,
        source_text_projection_sha256: facts.source_text_projection_sha256,
        output_text_projection_sha256: facts.output_text_projection_sha256,
        source_text_projection_bytes: facts.source_text_projection_bytes,
        output_text_projection_bytes: facts.output_text_projection_bytes,
        source_manifest_bindings_verified: facts.source_manifest_bindings_verified,
        output_manifest_bindings_verified: facts.output_manifest_bindings_verified,
        source_member_raw_preservation_verified: facts.source_member_raw_preservation_verified,
        manifest_entry_count: facts.manifest_entry_count,
        source_slide_count: facts.source_slide_count,
        output_slide_count: facts.output_slide_count,
        appended_title: facts.appended_title,
        appended_body: facts.appended_body,
        appended_title_sha256: facts.appended_title_sha256,
        appended_body_sha256: facts.appended_body_sha256,
        source_unchanged_verified: true,
        proof_slide_count,
        proof_content_xml_bytes,
        proof_insert_at,
        proof_page_name,
        proof_source_version_id: proof_source_version.id(),
        proof_source_version_revision: proof_source_version.revision(),
        stale_source_refusal_verified: stale_refusal_verified,
    };
    Ok((corpus, output))
}

/// Capture and gate the source lane's own candidate once, outside timed work.
pub(crate) fn build_odp_source_tail_append_corpus(
    shape: SemanticShape,
) -> Result<OdpSourceTailAppendCorpus, Box<dyn Error>> {
    let (corpus, output) = build_odp_source_tail_append_corpus_with_output(shape)?;
    drop(output);
    Ok(corpus)
}

/// Export the exact source and source-backed candidate archives used by this lane.
pub(crate) fn export_odp_source_tail_append_fixtures(
    directory: &Path,
) -> Result<(), Box<dyn Error>> {
    fs::create_dir(directory)?;
    let mut artifacts = Vec::with_capacity(SemanticShape::ALL.len());
    for shape in SemanticShape::ALL {
        let (corpus, output) = build_odp_source_tail_append_corpus_with_output(shape)?;
        let name = shape.name();
        let source_file = format!("{name}-source.odp");
        let candidate_file = format!("{name}-candidate.odp");
        let source_archive_sha256 = sha256_hex(&corpus.source.corpus.archive);
        let candidate_archive_sha256 = sha256_hex(&output);
        if source_archive_sha256 != corpus.source.corpus.manifest.archive_sha256
            || candidate_archive_sha256 != corpus.output_sha256
            || output.len() != corpus.output_bytes
        {
            return Err("ODP source append fixture bytes changed during export".into());
        }
        fs::write(
            directory.join(source_file.as_str()),
            &corpus.source.corpus.archive,
        )?;
        fs::write(directory.join(candidate_file.as_str()), &output)?;
        artifacts.push(OdpSourceTailAppendFixtureArtifact {
            shape: name,
            source_file,
            candidate_file,
            title: corpus.appended_title.clone(),
            body: corpus.appended_body.clone(),
            source_archive_sha256,
            source_archive_bytes: corpus.source.corpus.archive.len(),
            candidate_archive_sha256,
            candidate_archive_bytes: output.len(),
            source_content_xml_sha256: corpus.source_content_xml_sha256.clone(),
            source_content_xml_bytes: corpus.source.source_content_xml.len(),
            candidate_content_xml_sha256: corpus.output_content_xml_sha256.clone(),
            candidate_content_xml_bytes: corpus.output_content_xml_bytes,
            source_slide_count: corpus.source_slide_count,
            candidate_slide_count: corpus.output_slide_count,
            append_count: 1,
            source_semantic_sha256: corpus.source_semantic_sha256.clone(),
            candidate_semantic_sha256: corpus.output_semantic_sha256.clone(),
            source_order_sha256: corpus.source_order_sha256.clone(),
            candidate_order_sha256: corpus.output_order_sha256.clone(),
            source_text_projection_sha256: corpus.source_text_projection_sha256.clone(),
            candidate_text_projection_sha256: corpus.output_text_projection_sha256.clone(),
            source_text_projection_bytes: corpus.source_text_projection_bytes,
            candidate_text_projection_bytes: corpus.output_text_projection_bytes,
            source_member_count: corpus.source.source_members.len(),
            candidate_member_count: corpus.output_members.len(),
            source_members: corpus.source.source_members.clone(),
            candidate_members: corpus.output_members.clone(),
            manifest_entry_count: corpus.manifest_entry_count,
            proof_slide_count: corpus.proof_slide_count,
            proof_content_xml_bytes: corpus.proof_content_xml_bytes,
            proof_insert_at: corpus.proof_insert_at,
            proof_page_name: corpus.proof_page_name.clone(),
            source_manifest_bindings_verified: corpus.source_manifest_bindings_verified,
            candidate_manifest_bindings_verified: corpus.output_manifest_bindings_verified,
            source_member_raw_preservation_verified: corpus.source_member_raw_preservation_verified,
            append_exactly_one_verified: corpus.output_slide_count == corpus.source_slide_count + 1,
            source_unchanged_verified: corpus.source_unchanged_verified,
            stale_source_refusal_verified: corpus.stale_source_refusal_verified,
        });
        drop(output);
    }
    let manifest = OdpSourceTailAppendFixtureManifest {
        schema: ODP_SOURCE_TAIL_APPEND_FIXTURE_SCHEMA,
        generator: ODP_SOURCE_TAIL_APPEND_CORPUS_GENERATOR,
        package_format: "ODP/ODF/ZIP",
        artifact_count: artifacts.len(),
        artifacts,
    };
    fs::write(
        directory.join("manifest.json"),
        serde_json::to_vec_pretty(&manifest)?,
    )?;
    Ok(())
}

fn sink_observation(sink: SinkSummary) -> operation_metrics::SinkObservation {
    operation_metrics::SinkObservation {
        accepted_bytes: sink.accepted_bytes,
        write_calls: sink.write_calls,
        largest_write: sink.largest_write,
        bytes_0: sink.write_size_buckets.bytes_0,
        bytes_1_to_512: sink.write_size_buckets.bytes_1_to_512,
        bytes_513_to_4096: sink.write_size_buckets.bytes_513_to_4096,
        bytes_4097_to_16384: sink.write_size_buckets.bytes_4097_to_16384,
        bytes_16385_to_65536: sink.write_size_buckets.bytes_16385_to_65536,
        bytes_over_65536: sink.write_size_buckets.bytes_over_65536,
    }
}

fn ns(duration: Duration) -> Result<u64, Box<dyn Error>> {
    elapsed_ns(duration)
}

/// Run the source-backed ODP append publication lifecycle.
pub(crate) fn run_odp_source_tail_append_lifecycle(
    case: Case,
    corpus: &OdpSourceTailAppendCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if case != Case::OdpSourceTailAppendLifecycle
        || corpus.source.corpus.manifest.generator != ODP_SOURCE_TAIL_APPEND_CORPUS_GENERATOR
    {
        return Err("non-source-tail-append case passed to ODP source append runner".into());
    }
    let mut elapsed = Vec::with_capacity(samples);
    let mut open_ns = Vec::with_capacity(samples);
    let mut append_plan_ns = Vec::with_capacity(samples);
    let mut publication_ns = Vec::with_capacity(samples);
    let mut source_read_calls = Vec::with_capacity(samples);
    let mut source_read_bytes = Vec::with_capacity(samples);
    let mut publication_report_bytes = Vec::with_capacity(samples);
    let mut publication_report_source_version_id = Vec::with_capacity(samples);
    let mut publication_report_source_version_revision = Vec::with_capacity(samples);
    let mut runtime_proof_source_version_id = Vec::with_capacity(samples);
    let mut runtime_proof_source_version_revision = Vec::with_capacity(samples);
    let mut sink_summaries = Vec::with_capacity(samples);
    let mut output_digests = Vec::with_capacity(samples);
    let mut observations = Vec::with_capacity(samples);
    let iterations = iteration_count(warmup_iterations, samples)?;
    let source_bytes = Arc::<[u8]>::from(corpus.source.corpus.archive.clone());
    let title = odp_existing_append::appended_title(corpus.source.shape);
    let body = odp_existing_append::appended_body(corpus.source.shape);
    let options = append_options();
    for iteration in 0..iterations {
        // Source bytes, source provider, strings, publication options, and
        // hashing sink are all prepared before the lifecycle clock.  The
        // clock includes direct SourceBackedPackage opening, the bounded ODP
        // scan/proof and insertion-plan preparation, and sequential write.
        let iteration_title = title.clone();
        let iteration_body = body.clone();
        let iteration_options = options.clone();
        let provider = Arc::new(MeasureBytesSource::new(Arc::clone(&source_bytes)));
        let mut sink =
            HashingDiscardSink::without_authoring_window(u64::try_from(corpus.output_bytes)?);
        let process_before = process_metrics::Snapshot::read().ok();
        let allocation_region = allocation_metrics::begin();
        let lifecycle_started = std::time::Instant::now();
        let open_started = std::time::Instant::now();
        let package = open_source_package(&provider)?;
        let opened = open_started.elapsed();
        let plan_started = std::time::Instant::now();
        let edit =
            litchi_odp::SourceBackedTailAppendEdit::new(package, iteration_title, iteration_body)?;
        let plan = edit.plan(&iteration_options)?;
        let planned = plan_started.elapsed();
        let publication_started = std::time::Instant::now();
        let report = plan.write_to(&mut sink, iteration_options)?;
        let published = publication_started.elapsed();
        let lifecycle = lifecycle_started.elapsed();
        let allocation_metrics = allocation_region.finish();
        let process_after = process_metrics::Snapshot::read().ok();

        let report_bytes = report.bytes();
        let report_source_version = report.source_version();
        let proof_source_version = plan.proof().source_version();
        let provider_source_version = provider.version()?;
        let proof_slide_count = plan.proof().slide_count();
        let proof_content_xml_bytes = plan.proof().content_bytes();
        let proof_insert_at = plan.proof().insert_at();
        let proof_page_name = plan.proof().page_name().to_owned();
        let (sink_summary, candidate_digest) = sink.finish();
        let (read_calls, read_bytes) = provider.metrics();
        if candidate_digest != corpus.output_sha256
            || sink_summary.accepted_bytes != u64::try_from(corpus.output_bytes)?
            || report_bytes != u64::try_from(corpus.output_bytes)?
        {
            return Err(
                "ODP source append runtime output digest or length differs from candidate".into(),
            );
        }
        if read_calls == 0 || read_bytes == 0 {
            return Err("ODP source append runtime did not observe positional source reads".into());
        }
        if !provider.matches_bytes(corpus.source.corpus.archive.as_slice()) {
            return Err("ODP source append runtime changed its source bytes".into());
        }
        if proof_slide_count != corpus.proof_slide_count
            || proof_content_xml_bytes != corpus.proof_content_xml_bytes
            || proof_insert_at != corpus.proof_insert_at
            || proof_page_name != corpus.proof_page_name
        {
            return Err("ODP source append proof changed across iterations".into());
        }
        if proof_source_version != report_source_version
            || report_source_version != provider_source_version
        {
            return Err(
                "ODP source append proof, publication report, and provider source versions differ"
                    .into(),
            );
        }
        std::hint::black_box(candidate_digest.as_str());
        drop(plan);
        drop(provider);

        if iteration >= warmup_iterations {
            sink_summaries.push(sink_summary);
            output_digests.push(candidate_digest);
            open_ns.push(ns(opened)?);
            append_plan_ns.push(ns(planned)?);
            publication_ns.push(ns(published)?);
            publication_report_bytes.push(report_bytes);
            publication_report_source_version_id.push(report_source_version.id());
            publication_report_source_version_revision.push(report_source_version.revision());
            runtime_proof_source_version_id.push(proof_source_version.id());
            runtime_proof_source_version_revision.push(proof_source_version.revision());
            source_read_calls.push(read_calls);
            source_read_bytes.push(read_bytes);
            observations.push(operation_metrics::InProcessObservation {
                elapsed_ns: ns(lifecycle)?,
                process_metrics: process_before
                    .zip(process_after)
                    .map(|(before, after)| after.delta(before)),
                allocation_metrics,
            });
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, lifecycle)?;
    }

    let sink = deterministic_sink_summary(&sink_summaries, "ODP source tail append lifecycle")?;
    if sink.retained_output_bytes != Some(0) || sink.retained_authoring_window_bytes.is_some() {
        return Err("ODP source append lifecycle unexpectedly retained output/window bytes".into());
    }
    if output_digests
        .iter()
        .any(|digest| digest != &corpus.output_sha256)
    {
        return Err("ODP source append output digest changed across samples".into());
    }
    let source_observations = source_read_calls
        .iter()
        .zip(&source_read_bytes)
        .map(
            |(&read_calls, &read_bytes)| operation_metrics::InProcessSourceObservation {
                read_calls,
                read_bytes,
                max_concurrent_reads: if read_calls == 0 { 0 } else { 1 },
            },
        )
        .collect::<Vec<_>>();
    let elapsed_statistics = statistics(elapsed);
    let sample_order = elapsed_statistics.sample_order.clone();
    for values in [
        &mut open_ns,
        &mut append_plan_ns,
        &mut publication_ns,
        &mut publication_report_bytes,
        &mut publication_report_source_version_id,
        &mut publication_report_source_version_revision,
        &mut runtime_proof_source_version_id,
        &mut runtime_proof_source_version_revision,
        &mut source_read_calls,
        &mut source_read_bytes,
    ] {
        super::reorder_sample_vector(values, &sample_order)?;
    }
    super::reorder_sample_vector(&mut output_digests, &sample_order)?;
    let operation_metrics = Some(
        operation_metrics::from_in_process_source_and_sink_observations(
            &observations,
            &source_observations,
            sink_observation(sink),
        )?,
    );
    let source_member =
        odp_existing_append::find_member(&corpus.source.source_members, OPAQUE_PATH)?;
    let source = SourceSummary {
        read_calls: source_read_calls.clone(),
        read_bytes: source_read_bytes.clone(),
        odp_source_tail_append: Some(OdpSourceTailAppendSummary {
            role: "source_tail_append",
            implementation: "SourceBackedPackage::from_read_at + SourceBackedTailAppendEdit::plan + SourceBackedTailAppendPublicationPlan::write_to",
            timing_scope: "source archive Arc, positional provider, append strings/options, and HashingDiscardSink construction outside; lifecycle_ns includes direct SourceBackedPackage opening, bounded source scan/proof, insertion-plan preparation, and sequential publication-plan write; output digest finalization, semantic/member/raw/source/stale oracles, report assembly, and operation-owned drops occur outside the clock",
            performance_claim: "source-backed ODP advanced publication-plan lifecycle evidence only; the returned SourceContentPublicationReport is a publication report and has a different retained-result contract from owned Snapshot/Commit/Patch; no ordinary Commit/Patch optimization or general CRUD claim",
            corpus_generator: ODP_SOURCE_TAIL_APPEND_CORPUS_GENERATOR,
            shape: corpus.source.shape.name(),
            source_archive_sha256: corpus.source.corpus.manifest.archive_sha256.clone(),
            source_archive_bytes: corpus.source.corpus.archive.len(),
            output_archive_sha256: corpus.output_sha256.clone(),
            output_archive_bytes: corpus.output_bytes,
            source_content_xml_sha256: corpus.source_content_xml_sha256.clone(),
            source_content_xml_bytes: corpus.source.source_content_xml.len(),
            output_content_xml_sha256: corpus.output_content_xml_sha256.clone(),
            output_content_xml_bytes: corpus.output_content_xml_bytes,
            source_slide_count: corpus.source_slide_count,
            output_slide_count: corpus.output_slide_count,
            append_count: 1,
            appended_title: corpus.appended_title.clone(),
            appended_body: corpus.appended_body.clone(),
            appended_title_sha256: corpus.appended_title_sha256.clone(),
            appended_body_sha256: corpus.appended_body_sha256.clone(),
            opaque_member_path: OPAQUE_PATH,
            opaque_bytes: source_member.decoded_bytes,
            opaque_sha256: source_member.decoded_sha256.clone(),
            opaque_member_compressed_bytes: source_member.compressed_bytes,
            opaque_member_compressed_sha256: source_member.compressed_sha256.clone(),
            source_members: corpus.source.source_members.clone(),
            output_members: corpus.output_members.clone(),
            source_member_count: corpus.source.source_members.len(),
            output_member_count: corpus.output_members.len(),
            manifest_entry_count: corpus.manifest_entry_count,
            source_manifest_bindings_verified: corpus.source_manifest_bindings_verified,
            output_manifest_bindings_verified: corpus.output_manifest_bindings_verified,
            source_member_raw_preservation_verified: corpus.source_member_raw_preservation_verified,
            source_semantic_reopen_verified: true,
            output_semantic_reopen_verified: true,
            append_exactly_one_verified: corpus.output_slide_count == corpus.source_slide_count + 1,
            source_unchanged_verified: corpus.source_unchanged_verified,
            stale_source_refusal_verified: corpus.stale_source_refusal_verified,
            text_contract: "UTF-8 mixed Unicode/entities/plain text; one interior ASCII space; no CR, LF, tab, edge space, or repeated interior spaces",
            source_semantic_sha256: corpus.source_semantic_sha256.clone(),
            output_semantic_sha256: corpus.output_semantic_sha256.clone(),
            source_order_sha256: corpus.source_order_sha256.clone(),
            output_order_sha256: corpus.output_order_sha256.clone(),
            source_text_projection_sha256: corpus.source_text_projection_sha256.clone(),
            output_text_projection_sha256: corpus.output_text_projection_sha256.clone(),
            source_text_projection_bytes: corpus.source_text_projection_bytes,
            output_text_projection_bytes: corpus.output_text_projection_bytes,
            proof_slide_count: corpus.proof_slide_count,
            proof_content_xml_bytes: corpus.proof_content_xml_bytes,
            proof_insert_at: corpus.proof_insert_at,
            proof_page_name: corpus.proof_page_name.clone(),
            proof_source_version_id: corpus.proof_source_version_id,
            proof_source_version_revision: corpus.proof_source_version_revision,
            publication_report_bytes,
            publication_report_source_version_id,
            publication_report_source_version_revision,
            runtime_proof_source_version_id,
            runtime_proof_source_version_revision,
            source_read_calls,
            source_read_bytes,
            runtime_output_digest_verified: true,
            runtime_sink_length_verified: true,
            lifecycle_ns: elapsed_statistics.samples.clone(),
            open_ns,
            append_plan_ns,
            publication_ns,
            output_sha256: output_digests,
        }),
        ..SourceSummary::default()
    };
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.source.corpus.manifest.clone(),
        elapsed_ns: elapsed_statistics,
        sink: Some(sink),
        source: Some(Box::new(source)),
        execution: None,
        output_sha256: Some(corpus.output_sha256.clone()),
        operation_metrics,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn source_tail_append_is_opt_in_and_has_independent_candidate_gates() {
        let corpus = build_odp_source_tail_append_corpus(SemanticShape::Tiny).unwrap();
        assert_eq!(
            corpus.source.corpus.manifest.generator,
            ODP_SOURCE_TAIL_APPEND_CORPUS_GENERATOR
        );
        assert!(corpus.stale_source_refusal_verified);
        assert!(corpus.source_member_raw_preservation_verified);
        assert!(corpus.source_unchanged_verified);
        assert_eq!(corpus.output_slide_count, corpus.source_slide_count + 1);
        assert_eq!(
            corpus.appended_title,
            odp_existing_append::appended_title(SemanticShape::Tiny)
        );
        assert_eq!(
            corpus.appended_body,
            odp_existing_append::appended_body(SemanticShape::Tiny)
        );
        assert_ne!(
            corpus.output_sha256,
            corpus.source.corpus.manifest.archive_sha256
        );
        assert_eq!(
            super::super::parse_case("odp_source_tail_append_lifecycle"),
            Some(Case::OdpSourceTailAppendLifecycle)
        );
        assert!(!Case::DEFAULT.contains(&Case::OdpSourceTailAppendLifecycle));
    }

    #[test]
    fn runner_binds_source_reads_and_publication_report_to_candidate() {
        let corpus = build_odp_source_tail_append_corpus(SemanticShape::Tiny).unwrap();
        let result =
            run_odp_source_tail_append_lifecycle(Case::OdpSourceTailAppendLifecycle, &corpus, 0, 1)
                .unwrap();
        assert_eq!(
            result.output_sha256.as_deref(),
            Some(corpus.output_sha256.as_str())
        );
        let source = result.source.unwrap();
        let evidence = source.odp_source_tail_append.unwrap();
        assert!(evidence.runtime_output_digest_verified);
        assert!(evidence.runtime_sink_length_verified);
        assert_eq!(evidence.source_read_calls.len(), 1);
        assert!(evidence.source_read_calls[0] > 0);
        assert_eq!(
            evidence.publication_report_bytes,
            vec![corpus.output_bytes as u64]
        );
        assert_eq!(evidence.publication_report_source_version_id.len(), 1);
        assert_eq!(
            evidence.publication_report_source_version_id[0],
            evidence.runtime_proof_source_version_id[0]
        );
        assert_eq!(
            evidence.publication_report_source_version_revision[0],
            evidence.runtime_proof_source_version_revision[0]
        );
    }
}
