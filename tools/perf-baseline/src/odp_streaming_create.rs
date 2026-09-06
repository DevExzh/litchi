//! Matched fresh ODP titled-slide creation through the bounded streaming API.
//!
//! The input fixture and package gates are shared with the buffered Builder
//! baseline.  Corpus construction materializes one role-local streaming
//! artifact before warmups; measured iterations lazily create each title/body
//! string and pass it directly to the provider and a hashing discard sink.

use super::odp_buffered_create::{self, OdpSlidesSummary, OdpTextShapeCounts};
use super::{
    Case, CaseResult, Corpus, CorpusManifest, HashingDiscardSink, SemanticShape, SourceSummary,
    allocation_metrics, deterministic_sink_summary, elapsed_ns, iteration_count, operation_metrics,
    process_metrics, record_elapsed, semantic_shape, sha256_hex, statistics, streaming_context,
};
use std::{error::Error, time::Instant};

/// Stable role-local corpus identity for the streaming ODP producer.
pub(crate) const ODP_STREAMING_CORPUS_GENERATOR: &str = "litchi-odp-streaming-slides-v1";
const ODP_STREAMING_AUTHORING_WINDOW_BYTES: u64 = 4_096;
const ODP_STREAMING_MAX_TITLE_TEXT_BYTES: usize = 1 << 20;
const ODP_STREAMING_MAX_BODY_TEXT_BYTES: usize = 1 << 20;
const ODP_STREAMING_MAX_TOTAL_TEXT_BYTES: usize = 16 << 20;
const ODP_STREAMING_MAX_SLIDE_XML_BYTES: usize = 4_096;
const ODP_STREAMING_MAX_CONTENT_XML_BYTES: usize = 32 << 20;
const ODP_STREAMING_OUTPUT_LIMIT: u64 = 64 << 20;
const ODP_STREAMING_WORK_LIMIT: u64 = 64 << 20;
const ODP_STREAMING_IMPLEMENTATION: &str = "litchi_odp::streaming::stream_plain_slides_to";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct OdpStreamingReport {
    slides: usize,
    title_count: usize,
    body_count: usize,
    title_text_bytes: usize,
    body_text_bytes: usize,
    provider_input_text_bytes: usize,
    content_xml_bytes: usize,
}

fn odp_streaming_input_bytes(counts: &OdpTextShapeCounts) -> Result<usize, Box<dyn Error>> {
    counts
        .title_text_bytes
        .checked_add(counts.body_text_bytes)
        .ok_or_else(|| "ODP streaming provider input byte count overflows usize".into())
}

fn odp_streaming_limits(
    slides: usize,
) -> Result<litchi_odp::streaming::StreamingLimits, Box<dyn Error>> {
    Ok(litchi_odp::streaming::StreamingLimits::new(
        slides,
        ODP_STREAMING_MAX_TITLE_TEXT_BYTES,
        ODP_STREAMING_MAX_BODY_TEXT_BYTES,
        ODP_STREAMING_MAX_TOTAL_TEXT_BYTES,
        ODP_STREAMING_MAX_SLIDE_XML_BYTES,
        ODP_STREAMING_MAX_CONTENT_XML_BYTES,
        ODP_STREAMING_OUTPUT_LIMIT,
        litchi_odp::streaming::XmlAuditLimits::default(),
    )?)
}

fn odp_streaming_context(
    slides: usize,
    provider_input_text_bytes: usize,
    limits: litchi_odp::streaming::StreamingLimits,
) -> Result<litchi_core::ExecutionContext, Box<dyn Error>> {
    streaming_context(
        limits.required_memory_bytes()?,
        u64::try_from(provider_input_text_bytes)?,
        ODP_STREAMING_OUTPUT_LIMIT,
        u64::try_from(slides)?,
        ODP_STREAMING_WORK_LIMIT,
    )
}

/// Construct the matched source lazily so the fresh text allocation is inside
/// the measured public provider operation.
fn odp_streaming_source(
    shape: SemanticShape,
) -> impl Iterator<Item = litchi_odp::streaming::PlainSlide<String>> {
    (0..odp_buffered_create::odp_buffered_slide_count(shape)).map(|index| {
        litchi_odp::streaming::PlainSlide::new(
            Some(odp_buffered_create::odp_buffered_title(index)),
            odp_buffered_create::odp_buffered_body(index),
        )
    })
}

fn stream_odp_bytes(shape: SemanticShape) -> Result<(Vec<u8>, OdpStreamingReport), Box<dyn Error>> {
    let counts = odp_buffered_create::odp_buffered_text_shape_counts(shape)?;
    let slides = counts.slide_count;
    let provider_input_text_bytes = odp_streaming_input_bytes(&counts)?;
    let limits = odp_streaming_limits(slides)?;
    let context = odp_streaming_context(slides, provider_input_text_bytes, limits)?;
    let mut output = Vec::new();
    let report = litchi_odp::streaming::stream_plain_slides_to(
        &mut output,
        odp_streaming_source(shape),
        &context,
        limits,
    )?;
    drop(context);
    let report = OdpStreamingReport {
        slides: report.slides(),
        title_count: report.title_count(),
        body_count: report.body_count(),
        title_text_bytes: report.title_text_bytes(),
        body_text_bytes: report.body_text_bytes(),
        provider_input_text_bytes: report
            .title_text_bytes()
            .checked_add(report.body_text_bytes())
            .ok_or("ODP streaming provider input byte count overflows usize")?,
        content_xml_bytes: report.content_xml_bytes(),
    };
    if report.slides != counts.slide_count
        || report.title_count != counts.title_count
        || report.body_count != counts.body_count
        || report.title_text_bytes != counts.title_text_bytes
        || report.body_text_bytes != counts.body_text_bytes
        || report.provider_input_text_bytes != provider_input_text_bytes
    {
        return Err("streaming ODP provider report differs from the fixed slide fixture".into());
    }
    Ok((output, report))
}

/// Build and gate the role-local streaming artifact before warmups.
pub(crate) fn build_odp_streaming_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let counts = odp_buffered_create::odp_buffered_text_shape_counts(shape)?;
    let expected_provider_input_text_bytes = odp_streaming_input_bytes(&counts)?;
    let expected_semantic_sha256 = odp_buffered_create::odp_buffered_semantic_digest(shape)?;
    let (archive, report) = stream_odp_bytes(shape)?;
    let identity = odp_buffered_create::inspect_odp_buffered_archive(&archive, shape)?;
    if identity.semantic_sha256 != expected_semantic_sha256
        || report.provider_input_text_bytes != expected_provider_input_text_bytes
        || report.content_xml_bytes != identity.target_payload_bytes
    {
        return Err(
            "streaming ODP corpus report differs from its package/semantic identity".into(),
        );
    }
    let target_payload =
        soapberry_zip::office::ArchiveReader::new(&archive)?.read("content.xml")?;
    let corpus = Corpus {
        manifest: CorpusManifest {
            name: format!("odp-streaming-slides-{}", shape.name()),
            generator: ODP_STREAMING_CORPUS_GENERATOR,
            package_format: "ODP/ODF/ZIP",
            shape: shape.name(),
            payload_kind: "deterministic-mixed-unicode-entities-plain-titled-slides",
            compression: odp_buffered_create::ODP_BUFFERED_COMPRESSION,
            entry_count: identity.slide_count,
            archive_member_count: identity.archive_member_count,
            entry_bytes: identity.representative_slide_bytes,
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
    odp_buffered_create::verify_odp_corpus_binding(
        &corpus,
        shape,
        &identity,
        ODP_STREAMING_CORPUS_GENERATOR,
        &format!("odp-streaming-slides-{}", shape.name()),
    )?;
    Ok(corpus)
}

fn expected_sink_observation(sink: super::SinkSummary) -> operation_metrics::SinkObservation {
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

/// Run the bounded ODP streaming role against its own pre-gated corpus.
pub(crate) fn run_odp_streaming_creation(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if case != Case::OdpStreamingCreate
        || corpus.manifest.generator != ODP_STREAMING_CORPUS_GENERATOR
    {
        return Err("non-streaming ODP case passed to streaming creation runner".into());
    }
    let shape = semantic_shape(corpus)?;
    let counts = odp_buffered_create::odp_buffered_text_shape_counts(shape)?;
    let expected_semantic_sha256 = odp_buffered_create::odp_buffered_semantic_digest(shape)?;
    let expected_provider_input_text_bytes = odp_streaming_input_bytes(&counts)?;
    let identity = odp_buffered_create::inspect_odp_buffered_archive(&corpus.archive, shape)?;
    odp_buffered_create::verify_odp_corpus_binding(
        corpus,
        shape,
        &identity,
        ODP_STREAMING_CORPUS_GENERATOR,
        &format!("odp-streaming-slides-{}", shape.name()),
    )?;
    if identity.semantic_sha256 != expected_semantic_sha256 {
        return Err("streaming ODP corpus semantic identity differs from its fixture".into());
    }
    let limits = odp_streaming_limits(counts.slide_count)?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut summaries = Vec::with_capacity(samples);
    let mut digests = Vec::with_capacity(samples);
    let mut observations = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        // Sink, provider limits, and context setup are outside the operation
        // clock.  The lazy title/body source, provider publication, and sink
        // writes are the measured operation.  Context destruction is kept
        // after both resource endpoint snapshots, matching the buffered/ODT
        // harness scope and avoiding a negative live-allocation delta.
        let mut sink = HashingDiscardSink::new(
            ODP_STREAMING_OUTPUT_LIMIT,
            ODP_STREAMING_AUTHORING_WINDOW_BYTES,
        );
        let context = odp_streaming_context(
            counts.slide_count,
            expected_provider_input_text_bytes,
            limits,
        )?;
        let process_before = process_metrics::Snapshot::read().ok();
        let allocation_region = allocation_metrics::begin();
        let started = Instant::now();
        let report = litchi_odp::streaming::stream_plain_slides_to(
            &mut sink,
            odp_streaming_source(shape),
            &context,
            limits,
        )?;
        let duration = started.elapsed();
        let allocation_metrics = allocation_region.finish();
        let process_after = process_metrics::Snapshot::read().ok();
        let process_metrics = process_before
            .zip(process_after)
            .map(|(before, after)| after.delta(before));
        drop(context);

        let report = OdpStreamingReport {
            slides: report.slides(),
            title_count: report.title_count(),
            body_count: report.body_count(),
            title_text_bytes: report.title_text_bytes(),
            body_text_bytes: report.body_text_bytes(),
            provider_input_text_bytes: report
                .title_text_bytes()
                .checked_add(report.body_text_bytes())
                .ok_or("ODP streaming provider input byte count overflows usize")?,
            content_xml_bytes: report.content_xml_bytes(),
        };
        if report.slides != counts.slide_count
            || report.title_count != counts.title_count
            || report.body_count != counts.body_count
            || report.title_text_bytes != counts.title_text_bytes
            || report.body_text_bytes != counts.body_text_bytes
            || report.provider_input_text_bytes != expected_provider_input_text_bytes
            || report.content_xml_bytes != identity.content_xml_bytes
        {
            return Err("streaming ODP provider report changed during a sample".into());
        }
        std::hint::black_box(report);
        let (mut summary, digest) = sink.finish();
        if digest != corpus.manifest.archive_sha256
            || summary.accepted_bytes != u64::try_from(corpus.manifest.archive_bytes)?
        {
            return Err(
                "streaming ODP output digest or sink length differs from its corpus".into(),
            );
        }
        summary.paragraphs = Some(u64::try_from(counts.slide_count)?);
        summary.runs = Some(
            u64::try_from(counts.title_count)?
                .checked_add(u64::try_from(counts.body_count)?)
                .ok_or("ODP streaming text object count overflows u64")?,
        );
        // The provider report independently exposes its raw title/body bytes.
        // Keep the generic sink input metric on the same raw source contract;
        // the corpus manifest retains the presentation LF projection.
        summary.input_bytes = Some(u64::try_from(expected_provider_input_text_bytes)?);
        summary.authored_part_bytes = Some(u64::try_from(report.content_xml_bytes)?);
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
    let sink = deterministic_sink_summary(&summaries, "streaming ODP creation")?;
    if sink.retained_output_bytes != Some(0)
        || sink.retained_authoring_window_bytes != Some(ODP_STREAMING_AUTHORING_WINDOW_BYTES)
    {
        return Err("streaming ODP creation did not report the fixed authoring window".into());
    }
    if digests
        .iter()
        .any(|digest| digest != &corpus.manifest.archive_sha256)
    {
        return Err("streaming ODP creation output digest changed across samples".into());
    }
    let operation_metrics = Some(operation_metrics::from_in_process_observations(
        &observations,
        expected_sink_observation(sink),
    )?);
    let source = SourceSummary {
        odp_slides: Some(OdpSlidesSummary {
            role: "streaming",
            implementation: ODP_STREAMING_IMPLEMENTATION,
            timing_scope: "fresh title/body String generation, bounded provider validation/XML emission/package publication, and HashingDiscardSink writes inside the operation timer; sink/context/limits setup, corpus construction, archive reopen, digest extraction, report diagnostics, and package/semantic gates are outside; context destruction follows allocator/process endpoint snapshots",
            performance_claim: "timing and process/RSS/allocator evidence only; fixed 4096-byte provider authoring window and zero retained output; no physical-I/O, throughput, fixed-memory, or cross-role lexical content/archive equality claim",
            semantic_sha256: expected_semantic_sha256,
            content_xml_sha256: identity.content_xml_sha256.clone(),
            styles_xml_sha256: identity.styles_xml_sha256.clone(),
            meta_xml_sha256: identity.meta_xml_sha256.clone(),
            archive_member_set_verified: true,
            manifest_bindings_verified: true,
            semantic_reopen_verified: true,
            immutable_styles_meta_verified: true,
            runtime_output_digest_verified: true,
            runtime_sink_length_verified: true,
            page_structure_verified: true,
            page_geometry_verified: true,
            provider_input_text_bytes: Some(expected_provider_input_text_bytes),
            slide_count: counts.slide_count,
            title_count: counts.title_count,
            body_count: counts.body_count,
            title_text_bytes: counts.title_text_bytes,
            body_text_bytes: counts.body_text_bytes,
            title_variant_counts: counts.title_variant_counts,
            body_variant_counts: counts.body_variant_counts,
            text_contract: "four-cycle plain/Unicode/XML-significant UTF-8 title and body text with single interior spaces; one title and one body text object per slide",
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

    #[test]
    fn streaming_tiny_report_matches_shared_fixture() {
        let (archive, report) = stream_odp_bytes(SemanticShape::Tiny).unwrap();
        let counts =
            odp_buffered_create::odp_buffered_text_shape_counts(SemanticShape::Tiny).unwrap();
        let identity =
            odp_buffered_create::inspect_odp_buffered_archive(&archive, SemanticShape::Tiny)
                .unwrap();
        assert_eq!(report.slides, counts.slide_count);
        assert_eq!(report.title_count, counts.title_count);
        assert_eq!(report.body_count, counts.body_count);
        assert_eq!(report.title_text_bytes, counts.title_text_bytes);
        assert_eq!(report.body_text_bytes, counts.body_text_bytes);
        assert_eq!(
            report.provider_input_text_bytes,
            counts.title_text_bytes + counts.body_text_bytes
        );
        assert_eq!(report.content_xml_bytes, identity.content_xml_bytes);
        assert_eq!(
            identity.semantic_sha256,
            odp_buffered_create::odp_buffered_semantic_digest(SemanticShape::Tiny).unwrap()
        );
        assert_eq!(
            identity.semantic_sha256,
            "167428219fc1603b63046c349a06ef8d0033ac9f720567c5383bb79e379d10d1"
        );
    }

    #[test]
    fn streaming_tiny_runner_binds_raw_sink_input_and_fixed_window() {
        let corpus = build_odp_streaming_corpus(SemanticShape::Tiny).unwrap();
        let result = run_odp_streaming_creation(Case::OdpStreamingCreate, &corpus, 0, 1).unwrap();
        let sink = result.sink.unwrap();
        let counts =
            odp_buffered_create::odp_buffered_text_shape_counts(SemanticShape::Tiny).unwrap();
        assert_eq!(sink.retained_output_bytes, Some(0));
        assert_eq!(
            sink.retained_authoring_window_bytes,
            Some(ODP_STREAMING_AUTHORING_WINDOW_BYTES)
        );
        assert_eq!(
            sink.input_bytes,
            Some(u64::try_from(counts.title_text_bytes + counts.body_text_bytes).unwrap())
        );
        let source = result.source.unwrap();
        let summary = source.odp_slides.unwrap();
        assert_eq!(summary.role, "streaming");
        assert_eq!(summary.implementation, ODP_STREAMING_IMPLEMENTATION);
        assert_eq!(
            summary.provider_input_text_bytes,
            Some(counts.title_text_bytes + counts.body_text_bytes)
        );
        assert!(summary.page_structure_verified);
        assert!(summary.page_geometry_verified);
    }
}
