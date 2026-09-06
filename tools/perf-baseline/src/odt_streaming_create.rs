//! Matched fresh ODT paragraph creation through buffered and bounded writers.
//!
//! Corpus construction and semantic/package gates are outside the measured
//! loop. Both roles generate fresh paragraph strings and publish to a hashing
//! discard sink inside the operation clock.

use super::{
    Case, CaseResult, Corpus, CorpusManifest, HashingDiscardSink, SemanticShape, SourceSummary,
    allocation_metrics, deterministic_sink_summary, elapsed_ns, iteration_count, operation_metrics,
    process_metrics, record_elapsed, semantic_shape, sha256_hex, statistics, streaming_context,
};
use serde::Serialize;
use sha2::Digest as _;
use soapberry_zip::office::ArchiveReader;
use soapberry_zip::{CompressionMethod, ZipArchive};
use std::{collections::BTreeSet, error::Error, io::Write, time::Instant};

/// Stable corpus identity for the buffered ODT role.
pub(crate) const ODT_BUFFERED_CORPUS_GENERATOR: &str = "litchi-odt-buffered-paragraphs-v1";
const ODT_BUFFERED_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.text";
const ODT_BUFFERED_MEMBER_NAMES: [&str; 5] = [
    "mimetype",
    "content.xml",
    "styles.xml",
    "meta.xml",
    "META-INF/manifest.xml",
];
// These are independently pinned bytes from the fixed, empty Builder
// package template.  They are intentionally not obtained by constructing a
// second Builder during corpus inspection: a production change must either
// preserve this package contract or fail the baseline gate visibly.
const ODT_BUFFERED_DEFAULT_STYLES_BYTES: usize = 2_477;
const ODT_BUFFERED_DEFAULT_STYLES_SHA256: &str =
    "4d1dfc46e4c369722193548ff59269167f3e838a4b61bf1ddca1ecafe5092277";
const ODT_BUFFERED_DEFAULT_META_BYTES: usize = 387;
const ODT_BUFFERED_DEFAULT_META_SHA256: &str =
    "c7e55a3560c73aa42da85eec4751c3e78b5cc53ff964f50acba6c5cd105e6719";
const ODT_BUFFERED_COMPRESSION: &str = "mimetype=stored;xml=deflate";
pub(crate) const ODT_STREAMING_CORPUS_GENERATOR: &str = "litchi-odt-streaming-paragraphs-v1";
const ODT_STREAMING_OUTPUT_LIMIT: u64 = 64 * 1024 * 1024;
const ODT_STREAMING_PARAGRAPH_XML_WINDOW: usize = 4_096;
const ODT_STREAMING_MAX_PARAGRAPH_TEXT_BYTES: usize = 1 << 20;
const ODT_STREAMING_MAX_TOTAL_TEXT_BYTES: usize = 16 << 20;
const ODT_STREAMING_MAX_CONTENT_XML_BYTES: usize = 32 << 20;
const ODT_STREAMING_WORK_LIMIT: u64 = 64 * 1024 * 1024;

/// Source evidence for one role-local fresh ODT paragraph corpus.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct OdtParagraphsSummary {
    pub(crate) role: &'static str,
    pub(crate) implementation: &'static str,
    pub(crate) timing_scope: &'static str,
    pub(crate) performance_claim: &'static str,
    pub(crate) semantic_sha256: String,
    pub(crate) content_xml_sha256: String,
    pub(crate) styles_xml_sha256: String,
    pub(crate) meta_xml_sha256: String,
    pub(crate) archive_member_set_verified: bool,
    pub(crate) manifest_bindings_verified: bool,
    pub(crate) semantic_reopen_verified: bool,
    pub(crate) immutable_styles_meta_verified: bool,
    pub(crate) paragraph_count: usize,
    pub(crate) run_count: usize,
    pub(crate) text_contract: &'static str,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct OdtBufferedIdentity {
    archive_bytes: usize,
    archive_member_count: usize,
    entry_count: usize,
    entry_bytes: usize,
    archive_sha256: String,
    semantic_sha256: String,
    semantic_input_bytes: usize,
    target_payload_bytes: usize,
    target_payload_sha256: String,
    content_xml_bytes: usize,
    content_xml_sha256: String,
    styles_xml_sha256: String,
    meta_xml_sha256: String,
}

const fn odt_buffered_paragraph_count(shape: SemanticShape) -> usize {
    match shape {
        SemanticShape::Tiny => 64,
        SemanticShape::Medium => 8_192,
        SemanticShape::Large => 32_768,
    }
}

/// Return the fixed mixed-text paragraph used by the fresh Builder baseline.
///
/// ODF 1.3 whitespace folding is intentionally not part of this corpus: each
/// paragraph has only single interior spaces and no edge/control whitespace.
/// The fixture still exercises ordinary text, non-ASCII UTF-8, and literal
/// XML-significant characters that the Builder must escape and the Document
/// facade must restore.
fn odt_buffered_paragraph_text(index: usize) -> String {
    let variant = match index % 4 {
        0 => "plain paragraph",
        1 => "Unicode café Δ 中",
        2 => "entities <&> \"quoted\"",
        _ => "mixed façade Ω <&> value",
    };
    format!("litchi-perf-odt-buffered-{index:05} {variant}")
}

fn odt_buffered_expected_paragraphs(shape: SemanticShape) -> Vec<String> {
    (0..odt_buffered_paragraph_count(shape))
        .map(odt_buffered_paragraph_text)
        .collect()
}

fn odt_buffered_semantic_digest(paragraphs: &[String]) -> Result<String, Box<dyn Error>> {
    let mut hasher = sha2::Sha256::new();
    hasher.update(b"litchi-odt-buffered-semantic-v1\0");
    hasher.update(
        u64::try_from(paragraphs.len())
            .map_err(|_error| "ODT paragraph count does not fit semantic digest")?
            .to_le_bytes(),
    );
    for paragraph in paragraphs {
        hasher.update(
            u64::try_from(paragraph.len())
                .map_err(|_error| "ODT paragraph length does not fit semantic digest")?
                .to_le_bytes(),
        );
        hasher.update(paragraph.as_bytes());
    }
    let digest = hasher.finalize();
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    Ok(output)
}

fn odt_buffered_semantic_input_bytes(paragraphs: &[String]) -> Result<usize, Box<dyn Error>> {
    paragraphs
        .iter()
        .enumerate()
        .try_fold(0usize, |total, (index, paragraph)| {
            let separator = if index + 1 < paragraphs.len() { 1 } else { 0 };
            total
                .checked_add(paragraph.len())
                .and_then(|value| value.checked_add(separator))
                .ok_or_else(|| "ODT semantic projection byte count overflows usize".into())
        })
}

fn odt_buffered_bytes(shape: SemanticShape) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut builder = litchi_odt::Builder::new();
    for index in 0..odt_buffered_paragraph_count(shape) {
        let paragraph = odt_buffered_paragraph_text(index);
        builder.add_paragraph(&paragraph)?;
    }
    Ok(builder.build()?)
}

fn verify_odt_buffered_default_parts(
    styles_xml: &[u8],
    meta_xml: &[u8],
) -> Result<(), Box<dyn Error>> {
    if styles_xml.len() != ODT_BUFFERED_DEFAULT_STYLES_BYTES
        || sha256_hex(styles_xml) != ODT_BUFFERED_DEFAULT_STYLES_SHA256
    {
        return Err("buffered ODT styles.xml differs from the pinned Builder default".into());
    }
    if meta_xml.len() != ODT_BUFFERED_DEFAULT_META_BYTES
        || sha256_hex(meta_xml) != ODT_BUFFERED_DEFAULT_META_SHA256
    {
        return Err("buffered ODT meta.xml differs from the pinned Builder default".into());
    }
    Ok(())
}

fn verify_odt_buffered_compression(bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    let archive = ZipArchive::from_slice(bytes)?;
    let mut file_count = 0usize;
    for entry in archive.entries() {
        let entry = entry?;
        if entry.is_dir() {
            continue;
        }
        file_count = file_count
            .checked_add(1)
            .ok_or("buffered ODT ZIP member count overflows usize")?;
        let path = entry.file_path().try_normalize()?;
        let expected = if path.as_str() == "mimetype" {
            CompressionMethod::Store
        } else {
            CompressionMethod::Deflate
        };
        if entry.compression_method() != expected {
            return Err(format!(
                "buffered ODT member {} uses {:?}, expected {:?}",
                path.as_str(),
                entry.compression_method(),
                expected
            )
            .into());
        }
    }
    if file_count != ODT_BUFFERED_MEMBER_NAMES.len() {
        return Err("buffered ODT ZIP member count differs from the package contract".into());
    }
    Ok(())
}

fn verify_odt_buffered_manifest(archive: &ArchiveReader<'_>) -> Result<(), Box<dyn Error>> {
    let manifest_xml = String::from_utf8(archive.read("META-INF/manifest.xml")?)?;
    let manifest = litchi_odf_common::core::Manifest::parse(&manifest_xml)?;
    let root = manifest
        .get_entry("/")
        .ok_or("buffered ODT manifest root entry is missing")?;
    let content = manifest
        .get_entry("content.xml")
        .ok_or("buffered ODT manifest content.xml entry is missing")?;
    let styles = manifest
        .get_entry("styles.xml")
        .ok_or("buffered ODT manifest styles.xml entry is missing")?;
    let meta = manifest
        .get_entry("meta.xml")
        .ok_or("buffered ODT manifest meta.xml entry is missing")?;
    if manifest.entries.len() != 4
        || manifest.mimetype != "application/vnd.oasis.opendocument.text"
        || root.media_type != "application/vnd.oasis.opendocument.text"
        || root.size.is_some()
        || root.encryption.is_some()
        || content.media_type != "text/xml"
        || content.size.is_some()
        || content.encryption.is_some()
        || styles.media_type != "text/xml"
        || styles.size.is_some()
        || styles.encryption.is_some()
        || meta.media_type != "text/xml"
        || meta.size.is_some()
        || meta.encryption.is_some()
    {
        return Err("buffered ODT manifest bindings differ from the Builder contract".into());
    }
    Ok(())
}

fn inspect_odt_buffered_archive(
    bytes: &[u8],
    shape: SemanticShape,
) -> Result<OdtBufferedIdentity, Box<dyn Error>> {
    let archive = ArchiveReader::new(bytes)?;
    let names = archive.file_names().collect::<Vec<_>>();
    let actual_set = names.iter().copied().collect::<BTreeSet<_>>();
    let expected_set = ODT_BUFFERED_MEMBER_NAMES
        .iter()
        .copied()
        .collect::<BTreeSet<_>>();
    if names.len() != ODT_BUFFERED_MEMBER_NAMES.len() || actual_set != expected_set {
        return Err(format!(
            "buffered ODT archive member set differs: actual={names:?}, expected={ODT_BUFFERED_MEMBER_NAMES:?}"
        )
        .into());
    }
    if archive.read("mimetype")? != ODT_BUFFERED_MIMETYPE {
        return Err("buffered ODT mimetype member differs from the ODT MIME".into());
    }
    verify_odt_buffered_compression(bytes)?;
    verify_odt_buffered_manifest(&archive)?;

    let content_xml = archive.read("content.xml")?;
    let styles_xml = archive.read("styles.xml")?;
    let meta_xml = archive.read("meta.xml")?;
    verify_odt_buffered_default_parts(&styles_xml, &meta_xml)?;

    let document = litchi_odt::Document::from_bytes(bytes.to_vec())?;
    let paragraphs = document.paragraphs()?;
    let expected = odt_buffered_expected_paragraphs(shape);
    if paragraphs.len() != expected.len() {
        return Err("buffered ODT paragraph count differs from the fixture".into());
    }
    let mut actual = Vec::with_capacity(paragraphs.len());
    for paragraph in paragraphs {
        actual.push(paragraph.text()?);
    }
    if actual != expected {
        return Err("buffered ODT paragraph semantic projection differs from the fixture".into());
    }
    let expected_text = expected.join("\n");
    if document.text()? != expected_text {
        return Err("buffered ODT full-text projection differs from the fixture".into());
    }

    Ok(OdtBufferedIdentity {
        archive_bytes: bytes.len(),
        archive_member_count: names.len(),
        entry_count: expected.len(),
        entry_bytes: odt_buffered_paragraph_text(0).len(),
        archive_sha256: sha256_hex(bytes),
        semantic_sha256: odt_buffered_semantic_digest(&actual)?,
        semantic_input_bytes: odt_buffered_semantic_input_bytes(&expected)?,
        target_payload_bytes: content_xml.len(),
        target_payload_sha256: sha256_hex(&content_xml),
        content_xml_bytes: content_xml.len(),
        content_xml_sha256: sha256_hex(&content_xml),
        styles_xml_sha256: sha256_hex(&styles_xml),
        meta_xml_sha256: sha256_hex(&meta_xml),
    })
}

fn verify_odt_paragraph_corpus_binding(
    corpus: &Corpus,
    shape: SemanticShape,
    identity: &OdtBufferedIdentity,
    expected_name: &str,
    expected_generator: &str,
) -> Result<(), Box<dyn Error>> {
    let expected_semantic_sha256 =
        odt_buffered_semantic_digest(&odt_buffered_expected_paragraphs(shape))?;
    let reopened_target = ArchiveReader::new(&corpus.archive)?.read("content.xml")?;
    if corpus.target_name != "content.xml"
        || corpus.manifest.target_entry != "content.xml"
        || corpus.target_payload != reopened_target
    {
        return Err("buffered ODT corpus target entry is not bound to reopened content.xml".into());
    }
    if corpus.manifest.name != expected_name
        || corpus.manifest.generator != expected_generator
        || corpus.manifest.package_format != "ODT/ODF/ZIP"
        || corpus.manifest.shape != shape.name()
        || corpus.manifest.payload_kind != "deterministic-mixed-unicode-entities-plain-paragraphs"
        || corpus.manifest.compression != ODT_BUFFERED_COMPRESSION
    {
        return Err("buffered ODT corpus manifest identity differs from its contract".into());
    }
    if corpus.archive.len() != identity.archive_bytes
        || corpus.manifest.archive_bytes != identity.archive_bytes
        || corpus.manifest.archive_member_count != identity.archive_member_count
        || corpus.manifest.entry_count != identity.entry_count
        || corpus.manifest.entry_bytes != identity.entry_bytes
        || corpus.manifest.uncompressed_payload_bytes != identity.semantic_input_bytes
        || corpus.manifest.archive_sha256 != identity.archive_sha256
        || corpus.manifest.target_payload_bytes != identity.target_payload_bytes
        || corpus.manifest.target_payload_sha256 != identity.target_payload_sha256
    {
        return Err(
            "ODT paragraph corpus manifest is not bound to reopened archive identity".into(),
        );
    }
    if corpus.target_payload.len() != identity.target_payload_bytes
        || sha256_hex(&corpus.target_payload) != identity.target_payload_sha256
        || sha256_hex(&corpus.archive) != identity.archive_sha256
        || reopened_target.len() != identity.target_payload_bytes
        || sha256_hex(&reopened_target) != identity.target_payload_sha256
        || identity.semantic_input_bytes
            != odt_buffered_semantic_input_bytes(&odt_buffered_expected_paragraphs(shape))?
        || identity.semantic_sha256 != expected_semantic_sha256
    {
        return Err("ODT paragraph corpus payload or semantic projection is not bound".into());
    }
    Ok(())
}

/// Build and fully gate the role-local buffered ODT corpus before warmups.
pub(crate) fn build_odt_buffered_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let archive = odt_buffered_bytes(shape)?;
    let identity = inspect_odt_buffered_archive(&archive, shape)?;
    let target_payload = ArchiveReader::new(&archive)?.read("content.xml")?;
    let entry_bytes = odt_buffered_paragraph_text(0).len();
    let corpus = Corpus {
        manifest: CorpusManifest {
            name: format!("odt-buffered-paragraphs-{}", shape.name()),
            generator: ODT_BUFFERED_CORPUS_GENERATOR,
            package_format: "ODT/ODF/ZIP",
            shape: shape.name(),
            payload_kind: "deterministic-mixed-unicode-entities-plain-paragraphs",
            compression: ODT_BUFFERED_COMPRESSION,
            entry_count: odt_buffered_paragraph_count(shape),
            archive_member_count: ODT_BUFFERED_MEMBER_NAMES.len(),
            entry_bytes,
            uncompressed_payload_bytes: identity.semantic_input_bytes,
            archive_bytes: archive.len(),
            archive_sha256: identity.archive_sha256.clone(),
            target_entry: "content.xml".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "content.xml".to_owned(),
        target_payload,
        xlsx: None,
    };
    let expected_name = format!("odt-buffered-paragraphs-{}", shape.name());
    verify_odt_paragraph_corpus_binding(
        &corpus,
        shape,
        &identity,
        &expected_name,
        ODT_BUFFERED_CORPUS_GENERATOR,
    )?;
    Ok(corpus)
}

/// Run the buffered ODT authoring baseline.
pub(crate) fn run_odt_buffered_creation(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if case != Case::OdtBufferedCreate || corpus.manifest.generator != ODT_BUFFERED_CORPUS_GENERATOR
    {
        return Err("non-buffered ODT case passed to buffered creation runner".into());
    }
    let shape = semantic_shape(corpus)?;
    let paragraphs = odt_buffered_paragraph_count(shape);
    let expected = odt_buffered_expected_paragraphs(shape);
    let expected_semantic_sha256 = odt_buffered_semantic_digest(&expected)?;
    let expected_input_bytes = odt_buffered_semantic_input_bytes(&expected)?;
    let corpus_identity = inspect_odt_buffered_archive(&corpus.archive, shape)?;
    let expected_name = format!("odt-buffered-paragraphs-{}", shape.name());
    verify_odt_paragraph_corpus_binding(
        corpus,
        shape,
        &corpus_identity,
        &expected_name,
        ODT_BUFFERED_CORPUS_GENERATOR,
    )?;
    if corpus_identity.archive_sha256 != corpus.manifest.archive_sha256
        || corpus_identity.semantic_sha256 != expected_semantic_sha256
        || corpus_identity.semantic_input_bytes != expected_input_bytes
    {
        return Err("buffered ODT corpus identity differs from its manifest".into());
    }
    let maximum = u64::try_from(corpus.manifest.archive_bytes)?
        .checked_add(64 * 1024)
        .ok_or("buffered ODT sink ceiling overflows")?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut summaries = Vec::with_capacity(samples);
    let mut digests = Vec::with_capacity(samples);
    let mut observations = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        // Sink reservation and process/allocator setup are outside the clock.
        // Fresh paragraph model construction, Builder publication, and the
        // hashing sink write are the measured baseline operation.
        let mut sink = HashingDiscardSink::without_authoring_window(maximum);
        let process_before = process_metrics::Snapshot::read().ok();
        let allocation_region = allocation_metrics::begin();
        let started = Instant::now();
        let output = odt_buffered_bytes(shape)?;
        sink.write_all(&output)?;
        std::hint::black_box(output.len());
        drop(output);
        let duration = started.elapsed();
        let allocation_metrics = allocation_region.finish();
        let process_after = process_metrics::Snapshot::read().ok();
        let process_metrics = process_before
            .zip(process_after)
            .map(|(before, after)| after.delta(before));

        let (mut summary, digest) = sink.finish();
        if digest != corpus.manifest.archive_sha256
            || summary.accepted_bytes != u64::try_from(corpus.manifest.archive_bytes)?
        {
            return Err("buffered ODT creation digest or sink length differs from corpus".into());
        }
        summary.paragraphs = Some(u64::try_from(paragraphs)?);
        summary.runs = Some(u64::try_from(paragraphs)?);
        summary.input_bytes = Some(u64::try_from(expected_input_bytes)?);
        summary.authored_part_bytes = Some(u64::try_from(corpus.manifest.target_payload_bytes)?);
        if iteration >= warmup_iterations {
            summaries.push(summary);
            digests.push(digest);
            observations.push(operation_metrics::InProcessObservation {
                elapsed_ns: elapsed_ns(duration)?,
                process_metrics,
                allocation_metrics,
            });
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let sink = deterministic_sink_summary(&summaries, "buffered ODT creation")?;
    if sink.retained_output_bytes != Some(0) || sink.retained_authoring_window_bytes.is_some() {
        return Err("buffered ODT creation reported a retained output/window bound".into());
    }
    if digests
        .iter()
        .any(|digest| digest != &corpus.manifest.archive_sha256)
    {
        return Err("buffered ODT creation output digest changed across samples".into());
    }
    let sink_observation = operation_metrics::SinkObservation {
        accepted_bytes: sink.accepted_bytes,
        write_calls: sink.write_calls,
        largest_write: sink.largest_write,
        bytes_0: sink.write_size_buckets.bytes_0,
        bytes_1_to_512: sink.write_size_buckets.bytes_1_to_512,
        bytes_513_to_4096: sink.write_size_buckets.bytes_513_to_4096,
        bytes_4097_to_16384: sink.write_size_buckets.bytes_4097_to_16384,
        bytes_16385_to_65536: sink.write_size_buckets.bytes_16385_to_65536,
        bytes_over_65536: sink.write_size_buckets.bytes_over_65536,
    };
    let operation_metrics = Some(operation_metrics::from_in_process_observations(
        &observations,
        sink_observation,
    )?);
    let source = SourceSummary {
        odt_paragraphs: Some(OdtParagraphsSummary {
            role: "buffered",
            implementation: "litchi_odt::Builder",
            timing_scope: "fresh paragraph model construction, Builder::add_paragraph/build, and HashingDiscardSink write; output Vec release occurs before the clock stops and allocator/process endpoint snapshots; corpus setup, reopen, digest, sink finalization, and package/semantic gates are outside",
            performance_claim: "baseline timing and process/RSS/allocator evidence only; no fixed retained-window, throughput, or cross-role lexical-byte claim",
            semantic_sha256: expected_semantic_sha256,
            content_xml_sha256: corpus_identity.content_xml_sha256,
            styles_xml_sha256: corpus_identity.styles_xml_sha256,
            meta_xml_sha256: corpus_identity.meta_xml_sha256,
            archive_member_set_verified: true,
            manifest_bindings_verified: true,
            semantic_reopen_verified: true,
            immutable_styles_meta_verified: true,
            paragraph_count: paragraphs,
            run_count: paragraphs,
            text_contract: "four-cycle plain/Unicode/XML-significant UTF-8 text with single interior spaces; one logical text run per paragraph",
        }),
        ..SourceSummary::default()
    };
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(Box::new(source)),
        execution: None,
        output_sha256: Some(corpus.manifest.archive_sha256.clone()),
        operation_metrics,
    })
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct OdtStreamingReport {
    paragraphs: usize,
    input_text_bytes: usize,
    content_xml_bytes: usize,
}

fn odt_streaming_input_bytes(paragraphs: &[String]) -> Result<usize, Box<dyn Error>> {
    paragraphs.iter().try_fold(0usize, |total, paragraph| {
        total
            .checked_add(paragraph.len())
            .ok_or_else(|| "ODT streaming input byte count overflows usize".into())
    })
}

fn odt_streaming_limits(
    paragraphs: usize,
) -> Result<litchi_odt::streaming::StreamingLimits, Box<dyn Error>> {
    Ok(litchi_odt::streaming::StreamingLimits::new(
        paragraphs,
        ODT_STREAMING_MAX_PARAGRAPH_TEXT_BYTES,
        ODT_STREAMING_MAX_TOTAL_TEXT_BYTES,
        ODT_STREAMING_PARAGRAPH_XML_WINDOW,
        ODT_STREAMING_MAX_CONTENT_XML_BYTES,
        ODT_STREAMING_OUTPUT_LIMIT,
        litchi_odf_common::core::GeneratedXmlLimits::default(),
    )?)
}

fn odt_streaming_context(
    paragraphs: usize,
    input_text_bytes: usize,
    limits: litchi_odt::streaming::StreamingLimits,
) -> Result<litchi_core::ExecutionContext, Box<dyn Error>> {
    streaming_context(
        limits.required_memory_bytes()?,
        u64::try_from(input_text_bytes)?,
        ODT_STREAMING_OUTPUT_LIMIT,
        u64::try_from(paragraphs)?,
        ODT_STREAMING_WORK_LIMIT,
    )
}

fn odt_streaming_source(shape: SemanticShape) -> impl Iterator<Item = String> {
    (0..odt_buffered_paragraph_count(shape)).map(odt_buffered_paragraph_text)
}

fn stream_odt_bytes(shape: SemanticShape) -> Result<(Vec<u8>, OdtStreamingReport), Box<dyn Error>> {
    let expected = odt_buffered_expected_paragraphs(shape);
    let expected_input_bytes = odt_streaming_input_bytes(&expected)?;
    let paragraphs = expected.len();
    let limits = odt_streaming_limits(paragraphs)?;
    let context = odt_streaming_context(paragraphs, expected_input_bytes, limits)?;
    let mut output = Vec::new();
    let report = litchi_odt::streaming::stream_plain_paragraphs_to(
        &mut output,
        odt_streaming_source(shape),
        &context,
        limits,
    )?;
    drop(context);
    let report = OdtStreamingReport {
        paragraphs: report.paragraphs(),
        input_text_bytes: report.input_text_bytes(),
        content_xml_bytes: report.content_xml_bytes(),
    };
    if report.paragraphs != paragraphs || report.input_text_bytes != expected_input_bytes {
        return Err("streaming ODT provider report differs from the paragraph fixture".into());
    }
    Ok((output, report))
}

/// Build and fully gate the role-local streaming ODT corpus before warmups.
pub(crate) fn build_odt_streaming_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let expected = odt_buffered_expected_paragraphs(shape);
    let expected_input_bytes = odt_streaming_input_bytes(&expected)?;
    let expected_semantic_sha256 = odt_buffered_semantic_digest(&expected)?;
    let (archive, report) = stream_odt_bytes(shape)?;
    let identity = inspect_odt_buffered_archive(&archive, shape)?;
    if identity.semantic_sha256 != expected_semantic_sha256
        || report.input_text_bytes != expected_input_bytes
        || report.content_xml_bytes != identity.target_payload_bytes
    {
        return Err(
            "streaming ODT corpus report differs from its semantic/package identity".into(),
        );
    }
    let target_payload = ArchiveReader::new(&archive)?.read("content.xml")?;
    let corpus = Corpus {
        manifest: CorpusManifest {
            name: format!("odt-streaming-paragraphs-{}", shape.name()),
            generator: ODT_STREAMING_CORPUS_GENERATOR,
            package_format: "ODT/ODF/ZIP",
            shape: shape.name(),
            payload_kind: "deterministic-mixed-unicode-entities-plain-paragraphs",
            compression: ODT_BUFFERED_COMPRESSION,
            entry_count: expected.len(),
            archive_member_count: ODT_BUFFERED_MEMBER_NAMES.len(),
            entry_bytes: odt_buffered_paragraph_text(0).len(),
            uncompressed_payload_bytes: odt_buffered_semantic_input_bytes(&expected)?,
            archive_bytes: archive.len(),
            archive_sha256: identity.archive_sha256.clone(),
            target_entry: "content.xml".to_owned(),
            target_payload_bytes: target_payload.len(),
            target_payload_sha256: sha256_hex(&target_payload),
            rtf_variant: None,
            xlsx: None,
        },
        archive,
        target_name: "content.xml".to_owned(),
        target_payload,
        xlsx: None,
    };
    let expected_name = format!("odt-streaming-paragraphs-{}", shape.name());
    verify_odt_paragraph_corpus_binding(
        &corpus,
        shape,
        &identity,
        &expected_name,
        ODT_STREAMING_CORPUS_GENERATOR,
    )?;
    Ok(corpus)
}

/// Run the bounded ODT paragraph streaming role.
pub(crate) fn run_odt_streaming_creation(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if case != Case::OdtStreamingCreate
        || corpus.manifest.generator != ODT_STREAMING_CORPUS_GENERATOR
    {
        return Err("non-streaming ODT case passed to streaming creation runner".into());
    }
    let shape = semantic_shape(corpus)?;
    let expected = odt_buffered_expected_paragraphs(shape);
    let paragraphs = expected.len();
    let expected_semantic_sha256 = odt_buffered_semantic_digest(&expected)?;
    let expected_projection_bytes = odt_buffered_semantic_input_bytes(&expected)?;
    let expected_provider_input_bytes = odt_streaming_input_bytes(&expected)?;
    let corpus_identity = inspect_odt_buffered_archive(&corpus.archive, shape)?;
    let expected_name = format!("odt-streaming-paragraphs-{}", shape.name());
    verify_odt_paragraph_corpus_binding(
        corpus,
        shape,
        &corpus_identity,
        &expected_name,
        ODT_STREAMING_CORPUS_GENERATOR,
    )?;
    if corpus_identity.semantic_sha256 != expected_semantic_sha256 {
        return Err("streaming ODT corpus semantic identity differs from its fixture".into());
    }
    let limits = odt_streaming_limits(paragraphs)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut summaries = Vec::with_capacity(samples);
    let mut digests = Vec::with_capacity(samples);
    let mut observations = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        // Sink, provider limits, and context construction are outside the
        // clock.  The provider call itself constructs each paragraph String
        // lazily inside the timed API call.  Keep the fixed provider window
        // visible in the sink evidence as well as in StreamingLimits.
        let mut sink = HashingDiscardSink::new(
            ODT_STREAMING_OUTPUT_LIMIT,
            u64::try_from(ODT_STREAMING_PARAGRAPH_XML_WINDOW)?,
        );
        let context = odt_streaming_context(paragraphs, expected_provider_input_bytes, limits)?;
        let process_before = process_metrics::Snapshot::read().ok();
        let allocation_region = allocation_metrics::begin();
        let started = Instant::now();
        let report = litchi_odt::streaming::stream_plain_paragraphs_to(
            &mut sink,
            odt_streaming_source(shape),
            &context,
            limits,
        )?;
        let duration = started.elapsed();
        let allocation_metrics = allocation_region.finish();
        let process_after = process_metrics::Snapshot::read().ok();
        let process_metrics = process_before
            .zip(process_after)
            .map(|(before, after)| after.delta(before));
        // Context construction remains outside the timed region and its drop
        // is after both allocator and process endpoint snapshots.  This keeps
        // setup/destruction out of the operation clock without creating a
        // negative live-allocation delta.
        drop(context);

        if report.paragraphs() != paragraphs
            || report.input_text_bytes() != expected_provider_input_bytes
            || report.content_xml_bytes() != corpus.manifest.target_payload_bytes
        {
            return Err("streaming ODT provider report changed during sample".into());
        }
        std::hint::black_box(report);
        let (mut summary, digest) = sink.finish();
        if digest != corpus.manifest.archive_sha256
            || summary.accepted_bytes != u64::try_from(corpus.manifest.archive_bytes)?
        {
            return Err(
                "streaming ODT output digest or sink length differs from its corpus".into(),
            );
        }
        summary.paragraphs = Some(u64::try_from(paragraphs)?);
        summary.runs = Some(u64::try_from(paragraphs)?);
        // Keep the sink's semantic input metric aligned with the buffered
        // corpus projection (paragraphs joined by LF).  The provider report
        // above separately proves its actual source bytes exclude separators.
        summary.input_bytes = Some(u64::try_from(expected_projection_bytes)?);
        summary.authored_part_bytes = Some(u64::try_from(report.content_xml_bytes())?);
        if iteration >= warmup_iterations {
            summaries.push(summary);
            digests.push(digest);
            observations.push(operation_metrics::InProcessObservation {
                elapsed_ns: elapsed_ns(duration)?,
                process_metrics,
                allocation_metrics,
            });
        }
        record_elapsed(&mut elapsed, iteration, warmup_iterations, duration)?;
    }
    let sink = deterministic_sink_summary(&summaries, "streaming ODT creation")?;
    if sink.retained_output_bytes != Some(0)
        || sink.retained_authoring_window_bytes
            != Some(u64::try_from(ODT_STREAMING_PARAGRAPH_XML_WINDOW)?)
    {
        return Err("streaming ODT creation did not report the fixed authoring window".into());
    }
    if digests
        .iter()
        .any(|digest| digest != &corpus.manifest.archive_sha256)
    {
        return Err("streaming ODT output digest changed across samples".into());
    }
    let sink_observation = operation_metrics::SinkObservation {
        accepted_bytes: sink.accepted_bytes,
        write_calls: sink.write_calls,
        largest_write: sink.largest_write,
        bytes_0: sink.write_size_buckets.bytes_0,
        bytes_1_to_512: sink.write_size_buckets.bytes_1_to_512,
        bytes_513_to_4096: sink.write_size_buckets.bytes_513_to_4096,
        bytes_4097_to_16384: sink.write_size_buckets.bytes_4097_to_16384,
        bytes_16385_to_65536: sink.write_size_buckets.bytes_16385_to_65536,
        bytes_over_65536: sink.write_size_buckets.bytes_over_65536,
    };
    let operation_metrics = Some(operation_metrics::from_in_process_observations(
        &observations,
        sink_observation,
    )?);
    let source = SourceSummary {
        odt_paragraphs: Some(OdtParagraphsSummary {
            role: "streaming",
            implementation: "litchi_odt::streaming::stream_plain_paragraphs_to",
            timing_scope: "fresh paragraph String generation, provider validation/XML emission/package publication, and HashingDiscardSink writes inside stream_plain_paragraphs_to; sink/context/limits setup, corpus construction, reopen, digest, diagnostics, and package/semantic gates are outside; the provider's report is checked after the clock stops; context destruction follows allocator/process endpoint snapshots",
            performance_claim: "timing and process/RSS/allocator evidence only; fixed 4096-byte provider paragraph XML window; no physical-I/O, throughput, or lexical content-XML/archive equality claim",
            semantic_sha256: expected_semantic_sha256,
            content_xml_sha256: corpus_identity.content_xml_sha256,
            styles_xml_sha256: corpus_identity.styles_xml_sha256,
            meta_xml_sha256: corpus_identity.meta_xml_sha256,
            archive_member_set_verified: true,
            manifest_bindings_verified: true,
            semantic_reopen_verified: true,
            immutable_styles_meta_verified: true,
            paragraph_count: paragraphs,
            run_count: paragraphs,
            text_contract: "four-cycle plain/Unicode/XML-significant UTF-8 text with single interior spaces; one logical text run per paragraph",
        }),
        ..SourceSummary::default()
    };
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.manifest.clone(),
        elapsed_ns: statistics(elapsed),
        sink: Some(sink),
        source: Some(Box::new(source)),
        execution: None,
        output_sha256: Some(corpus.manifest.archive_sha256.clone()),
        operation_metrics,
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use soapberry_zip::office::StreamingArchiveWriter;

    #[test]
    fn buffered_text_projection_scales_with_stable_distinct_hashes() {
        let tiny = odt_buffered_expected_paragraphs(SemanticShape::Tiny);
        let medium = odt_buffered_expected_paragraphs(SemanticShape::Medium);
        let large = odt_buffered_expected_paragraphs(SemanticShape::Large);
        assert_eq!(tiny.len(), 64);
        assert_eq!(medium.len(), 8_192);
        assert_eq!(large.len(), 32_768);
        let tiny_again = odt_buffered_expected_paragraphs(SemanticShape::Tiny);
        assert_eq!(
            odt_buffered_semantic_digest(&tiny).unwrap(),
            odt_buffered_semantic_digest(&tiny_again).unwrap()
        );
        assert_ne!(
            odt_buffered_semantic_digest(&tiny).unwrap(),
            odt_buffered_semantic_digest(&large).unwrap()
        );
        assert!(tiny.iter().any(|value| value.contains("café Δ 中")));
        assert!(tiny.iter().any(|value| value.contains("<&>")));
    }

    #[test]
    fn buffered_tiny_archive_has_exact_package_and_semantic_oracle() {
        let archive = odt_buffered_bytes(SemanticShape::Tiny).unwrap();
        let identity = inspect_odt_buffered_archive(&archive, SemanticShape::Tiny).unwrap();
        assert_eq!(
            identity.semantic_input_bytes,
            odt_buffered_semantic_input_bytes(&odt_buffered_expected_paragraphs(
                SemanticShape::Tiny
            ),)
            .unwrap()
        );
        assert_eq!(identity.archive_sha256, sha256_hex(&archive));
        assert!(identity.content_xml_bytes > identity.semantic_input_bytes);
        assert_ne!(identity.styles_xml_sha256, identity.meta_xml_sha256);
    }

    #[test]
    fn buffered_oracle_rejects_manifest_binding_mutation() {
        let mut corpus = build_odt_buffered_corpus(SemanticShape::Tiny).unwrap();
        let reader = ArchiveReader::new(&corpus.archive).unwrap();
        let mut writer = StreamingArchiveWriter::new();
        for name in reader.file_names() {
            let mut payload = reader.read(name).unwrap();
            if name == "META-INF/manifest.xml" {
                let manifest = String::from_utf8(payload).unwrap();
                let mutated = manifest.replacen(
                    "manifest:full-path=\"content.xml\"",
                    "manifest:full-path=\"content-corrupt.xml\"",
                    1,
                );
                assert_ne!(mutated, manifest);
                payload = mutated.into_bytes();
            }
            if name == "mimetype" {
                writer.write_stored(name, &payload).unwrap();
            } else {
                writer.write_deflated(name, &payload).unwrap();
            }
        }
        corpus.archive = writer.finish_to_bytes().unwrap();
        let error = run_odt_buffered_creation(Case::OdtBufferedCreate, &corpus, 0, 1)
            .err()
            .expect("buffered ODT runner accepted a mutated manifest binding")
            .to_string();
        assert!(
            error.contains("manifest") || error.contains("content.xml"),
            "manifest mutation failed for the wrong reason: {error}"
        );
    }

    #[test]
    fn streaming_tiny_corpus_reuses_semantic_and_package_gates() {
        let expected = odt_buffered_expected_paragraphs(SemanticShape::Tiny);
        let (archive, report) = stream_odt_bytes(SemanticShape::Tiny).unwrap();
        assert_eq!(report.paragraphs, expected.len());
        assert_eq!(
            report.input_text_bytes,
            odt_streaming_input_bytes(&expected).unwrap()
        );
        assert_eq!(
            odt_buffered_semantic_input_bytes(&expected).unwrap(),
            report.input_text_bytes + expected.len() - 1
        );
        let identity = inspect_odt_buffered_archive(&archive, SemanticShape::Tiny).unwrap();
        assert_eq!(
            identity.semantic_sha256,
            odt_buffered_semantic_digest(&expected).unwrap()
        );
        assert_eq!(report.content_xml_bytes, identity.target_payload_bytes);
        assert_eq!(
            identity.archive_member_count,
            ODT_BUFFERED_MEMBER_NAMES.len()
        );
    }
}
