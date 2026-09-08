//! Opt-in fresh DOCX paragraph creation through the public streaming writer.
//!
//! The materialized package is built once outside the timed loop and is used
//! only as a deterministic reopen oracle. Timed iterations write to the
//! scalar-only [`HashingDiscardSink`] so the operation does not retain the
//! generated package.

use super::{
    Case, CaseResult, CorpusManifest, HashingDiscardSink, SemanticShape, SourceSummary,
    allocation_metrics, deterministic_sink_summary, elapsed_ns, iteration_count, operation_metrics,
    process_metrics, record_elapsed, sha256_hex, statistics, streaming_context,
};
use litchi_docx::{StreamingDocumentLimits, StreamingDocumentWriter};
use litchi_opc::phys_pkg::OwnedPhysPkgReader;
use serde::Serialize;
use sha2::{Digest as _, Sha256};
use std::{
    error::Error,
    io::{Cursor, Write},
};

pub(crate) const DOCX_STREAMING_CORPUS_GENERATOR: &str = "litchi-docx-streaming-paragraphs-v1";
const DOCX_STREAMING_SCRATCH_BYTES: u64 = 64;
const DOCX_STREAMING_OUTPUT_HEADROOM: u64 = 256 * 1024;
const DOCX_STREAMING_MEMBER_NAMES: [&str; 3] =
    ["[Content_Types].xml", "_rels/.rels", "word/document.xml"];
const DOCX_STREAMING_DOCUMENT_SUFFIX: &[u8] = b"<w:sectPr/></w:body></w:document>";

/// Source and semantic identity for one role-local fresh DOCX stream corpus.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct DocxParagraphsSummary {
    pub(crate) role: &'static str,
    pub(crate) implementation: &'static str,
    pub(crate) timing_scope: &'static str,
    pub(crate) performance_claim: &'static str,
    pub(crate) semantic_sha256: String,
    pub(crate) full_text_sha256: String,
    pub(crate) archive_sha256: String,
    pub(crate) target_payload_sha256: String,
    pub(crate) archive_member_set_verified: bool,
    pub(crate) semantic_reopen_verified: bool,
    pub(crate) deterministic_output_verified: bool,
    pub(crate) paragraph_count: usize,
    pub(crate) run_count: usize,
    pub(crate) input_text_bytes: usize,
    pub(crate) authored_part_bytes: usize,
    pub(crate) scratch_bytes: u64,
    pub(crate) text_contract: &'static str,
}

#[derive(Clone, Debug)]
pub(crate) struct DocxStreamingCorpus {
    pub(crate) manifest: CorpusManifest,
    spec: DocxStreamingSpec,
}

#[derive(Clone, Debug, PartialEq, Eq)]
struct DocxStreamingSpec {
    paragraphs: usize,
    input_text_bytes: usize,
    max_run_text_bytes: usize,
    semantic_sha256: String,
    full_text_sha256: String,
}

fn docx_streaming_text(index: usize) -> String {
    format!("litchi-perf-docx-streaming-v1-{index:06}-café-<&>")
}

fn digest_hex(hasher: Sha256) -> String {
    let digest = hasher.finalize();
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    output
}

fn build_spec(shape: SemanticShape) -> Result<DocxStreamingSpec, Box<dyn Error>> {
    let paragraphs = shape.streaming_units();
    let mut input_text_bytes = 0usize;
    let mut max_run_text_bytes = 0usize;
    let mut semantic = Sha256::new();
    semantic.update(b"litchi-docx-streaming-semantic-v1\0");
    semantic.update(u64::try_from(paragraphs)?.to_le_bytes());
    let mut full_text = Sha256::new();
    full_text.update(b"litchi-docx-streaming-full-text-v1\0");
    full_text.update(u64::try_from(paragraphs)?.to_le_bytes());
    for index in 0..paragraphs {
        let text = docx_streaming_text(index);
        let text_bytes = text.len();
        input_text_bytes = input_text_bytes
            .checked_add(text_bytes)
            .ok_or("DOCX streaming input byte count overflows usize")?;
        max_run_text_bytes = max_run_text_bytes.max(text_bytes);
        semantic.update(u64::try_from(text_bytes)?.to_le_bytes());
        semantic.update(text.as_bytes());
        full_text.update(text.as_bytes());
    }
    Ok(DocxStreamingSpec {
        paragraphs,
        input_text_bytes,
        max_run_text_bytes,
        semantic_sha256: digest_hex(semantic),
        full_text_sha256: digest_hex(full_text),
    })
}

fn document_xml_upper_bound(spec: &DocxStreamingSpec) -> Result<u64, Box<dyn Error>> {
    // This bound is deliberately independent of the serializer's private
    // constants. It leaves enough room for the fixed XML framing and every
    // legal XML escape in the generated text.
    let max_text = u64::try_from(spec.max_run_text_bytes)?
        .checked_mul(8)
        .and_then(|value| value.checked_add(256))
        .ok_or("DOCX streaming paragraph XML bound overflows")?;
    let paragraph_count = u64::try_from(spec.paragraphs)?;
    let paragraph_bytes = max_text
        .checked_mul(paragraph_count)
        .ok_or("DOCX streaming document XML bound overflows")?;
    paragraph_bytes
        .checked_add(16 * 1024)
        .ok_or_else(|| "DOCX streaming document XML bound overflows".into())
}

fn write_limits(spec: &DocxStreamingSpec) -> Result<StreamingDocumentLimits, Box<dyn Error>> {
    let input = u64::try_from(spec.input_text_bytes)?;
    let document_xml = document_xml_upper_bound(spec)?;
    let output = document_xml
        .checked_add(DOCX_STREAMING_OUTPUT_HEADROOM)
        .ok_or("DOCX streaming output bound overflows")?;
    let paragraphs = u64::try_from(spec.paragraphs)?;
    Ok(StreamingDocumentLimits::new(
        input,
        output,
        paragraphs,
        paragraphs,
        1,
        u64::try_from(spec.max_run_text_bytes)?,
        document_xml,
        document_xml,
        DOCX_STREAMING_SCRATCH_BYTES,
    ))
}

fn write_docx_stream<W: Write>(
    sink: W,
    spec: &DocxStreamingSpec,
) -> Result<(W, u64), Box<dyn Error>> {
    let limits = write_limits(spec)?;
    let context = streaming_context(
        DOCX_STREAMING_SCRATCH_BYTES,
        u64::try_from(spec.input_text_bytes)?,
        limits.max_output_bytes,
        limits
            .max_paragraphs
            .checked_mul(2)
            .and_then(|value| value.checked_add(16))
            .ok_or("DOCX streaming object budget overflows")?,
        limits
            .max_paragraphs
            .checked_mul(4)
            .and_then(|value| value.checked_add(u64::try_from(spec.input_text_bytes).ok()?))
            .and_then(|value| value.checked_add(32))
            .ok_or("DOCX streaming work budget overflows")?,
    )?;
    let mut writer = StreamingDocumentWriter::new(sink, context, limits)?;
    for index in 0..spec.paragraphs {
        writer.start_paragraph()?;
        writer.start_run()?;
        let text = docx_streaming_text(index);
        writer.write_text(&text)?;
        writer.finish_run()?;
        writer.finish_paragraph()?;
    }
    if writer.paragraph_count() != u64::try_from(spec.paragraphs)?
        || writer.run_count() != u64::try_from(spec.paragraphs)?
        || writer.input_bytes() != u64::try_from(spec.input_text_bytes)?
    {
        return Err("DOCX streaming writer counters differ from its paragraph fixture".into());
    }
    let document_xml_bytes_before_finish = writer.document_xml_bytes();
    Ok((writer.finish()?, document_xml_bytes_before_finish))
}

fn inspect_materialized_archive(
    archive: &[u8],
    spec: &DocxStreamingSpec,
) -> Result<Vec<u8>, Box<dyn Error>> {
    let physical = OwnedPhysPkgReader::from_bytes(archive.to_vec())?;
    let member_names = physical.member_names()?;
    if member_names != DOCX_STREAMING_MEMBER_NAMES {
        return Err(format!(
            "DOCX streaming archive members differ: actual={member_names:?}, expected={DOCX_STREAMING_MEMBER_NAMES:?}"
        )
        .into());
    }
    let package = litchi_docx::Package::from_reader(Cursor::new(archive.to_vec()))?;
    let document = package.document()?;
    if document.paragraph_count()? != spec.paragraphs {
        return Err("DOCX streaming paragraph count differs from its fixture".into());
    }
    let paragraphs = document.paragraphs()?;
    if paragraphs.len() != spec.paragraphs {
        return Err("DOCX streaming paragraph list differs from its fixture".into());
    }
    let mut semantic = Sha256::new();
    semantic.update(b"litchi-docx-streaming-semantic-v1\0");
    semantic.update(u64::try_from(spec.paragraphs)?.to_le_bytes());
    let mut full_text = String::new();
    for (index, paragraph) in paragraphs.into_iter().enumerate() {
        let runs = paragraph.runs()?;
        if runs.len() != 1 {
            return Err(format!(
                "DOCX streaming paragraph {index} has {} runs, expected one",
                runs.len()
            )
            .into());
        }
        let actual = paragraph.text()?;
        let expected = docx_streaming_text(index);
        if actual != expected || runs[0].text()? != expected {
            return Err(
                format!("DOCX streaming paragraph {index} differs from its fixture").into(),
            );
        }
        semantic.update(u64::try_from(actual.len())?.to_le_bytes());
        semantic.update(actual.as_bytes());
        full_text.push_str(&actual);
    }
    if digest_hex(semantic) != spec.semantic_sha256 {
        return Err("DOCX streaming paragraph digest differs from its fixture".into());
    }
    let mut full_text_hasher = Sha256::new();
    full_text_hasher.update(b"litchi-docx-streaming-full-text-v1\0");
    full_text_hasher.update(u64::try_from(spec.paragraphs)?.to_le_bytes());
    full_text_hasher.update(full_text.as_bytes());
    if digest_hex(full_text_hasher) != spec.full_text_sha256 || document.text()? != full_text {
        return Err("DOCX streaming full-text digest differs from its fixture".into());
    }
    let target = physical.read_member("word/document.xml")?;
    Ok(target)
}

/// Build and fully gate one role-local materialized DOCX streaming corpus.
pub(crate) fn build_corpus(shape: SemanticShape) -> Result<DocxStreamingCorpus, Box<dyn Error>> {
    let spec = build_spec(shape)?;
    let (archive, document_xml_bytes_before_finish) = write_docx_stream(Vec::new(), &spec)?;
    let target_payload = inspect_materialized_archive(&archive, &spec)?;
    let expected_before_finish = target_payload
        .len()
        .checked_sub(DOCX_STREAMING_DOCUMENT_SUFFIX.len())
        .ok_or("DOCX streaming document payload is unexpectedly short")?;
    if document_xml_bytes_before_finish != u64::try_from(expected_before_finish)? {
        return Err("DOCX streaming document XML counter differs from finalized part".into());
    }
    let archive_sha256 = sha256_hex(&archive);
    let target_payload_sha256 = sha256_hex(&target_payload);
    let manifest = CorpusManifest {
        name: format!("docx-streaming-paragraphs-{}", shape.name()),
        generator: DOCX_STREAMING_CORPUS_GENERATOR,
        package_format: "DOCX/OOXML/ZIP",
        shape: shape.name(),
        payload_kind: "deterministic-plain-unicode-xml-significant-paragraphs",
        compression: "deflate",
        entry_count: spec.paragraphs,
        archive_member_count: DOCX_STREAMING_MEMBER_NAMES.len(),
        entry_bytes: docx_streaming_text(0).len(),
        uncompressed_payload_bytes: spec.input_text_bytes,
        archive_bytes: archive.len(),
        archive_sha256,
        target_entry: "word/document.xml".to_owned(),
        target_payload_bytes: target_payload.len(),
        target_payload_sha256,
        rtf_variant: None,
        xlsx: None,
    };
    if target_payload.len() < DOCX_STREAMING_DOCUMENT_SUFFIX.len()
        || target_payload.len() - DOCX_STREAMING_DOCUMENT_SUFFIX.len() == 0
    {
        return Err("DOCX streaming document payload is unexpectedly short".into());
    }
    Ok(DocxStreamingCorpus { manifest, spec })
}

/// Run the bounded public DOCX streaming creation selector.
pub(crate) fn run(
    case: Case,
    corpus: &DocxStreamingCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if case != Case::DocxStreamingCreate
        || corpus.manifest.generator != DOCX_STREAMING_CORPUS_GENERATOR
    {
        return Err("non-streaming DOCX case passed to DOCX streaming runner".into());
    }
    let expected_archive_sha256 = &corpus.manifest.archive_sha256;
    let expected_target_sha256 = &corpus.manifest.target_payload_sha256;
    let expected_document_xml_bytes = u64::try_from(corpus.manifest.target_payload_bytes)?;
    let maximum = u64::try_from(corpus.manifest.archive_bytes)?
        .checked_add(DOCX_STREAMING_OUTPUT_HEADROOM)
        .ok_or("DOCX streaming sink ceiling overflows")?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut summaries = Vec::with_capacity(samples);
    let mut digests = Vec::with_capacity(samples);
    let mut observations = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let sink = HashingDiscardSink::new(maximum, DOCX_STREAMING_SCRATCH_BYTES);
        let process_before = process_metrics::Snapshot::read().ok();
        let allocation_region = allocation_metrics::begin();
        let started = std::time::Instant::now();
        let (sink, document_xml_bytes_before_finish) = write_docx_stream(sink, &corpus.spec)?;
        let duration = started.elapsed();
        let allocation_metrics = allocation_region.finish();
        let process_after = process_metrics::Snapshot::read().ok();
        let process_metrics = process_before
            .zip(process_after)
            .map(|(before, after)| after.delta(before));
        let (mut summary, digest) = sink.finish();
        let expected_before_finish = expected_document_xml_bytes
            .checked_sub(u64::try_from(DOCX_STREAMING_DOCUMENT_SUFFIX.len())?)
            .ok_or("DOCX streaming document payload is unexpectedly short")?;
        if document_xml_bytes_before_finish != expected_before_finish {
            return Err("DOCX streaming document XML counter changed during sample".into());
        }
        if summary.accepted_bytes != u64::try_from(corpus.manifest.archive_bytes)?
            || digest.as_str() != expected_archive_sha256
        {
            return Err(
                "DOCX streaming output digest or sink length differs from untimed artifact".into(),
            );
        }
        summary.paragraphs = Some(u64::try_from(corpus.spec.paragraphs)?);
        summary.runs = Some(u64::try_from(corpus.spec.paragraphs)?);
        summary.input_bytes = Some(u64::try_from(corpus.spec.input_text_bytes)?);
        summary.authored_part_bytes = Some(expected_document_xml_bytes);
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
    let sink = deterministic_sink_summary(&summaries, "streaming DOCX creation")?;
    if sink.retained_output_bytes != Some(0)
        || sink.retained_authoring_window_bytes != Some(DOCX_STREAMING_SCRATCH_BYTES)
    {
        return Err("DOCX streaming creation did not report its scratch reservation and zero retained output".into());
    }
    if digests
        .iter()
        .any(|digest| digest != expected_archive_sha256)
    {
        return Err("DOCX streaming output digest changed across samples".into());
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
    let operation_metrics =
        operation_metrics::from_in_process_observations(&observations, sink_observation)?;
    let source = SourceSummary {
        docx_paragraphs: Some(DocxParagraphsSummary {
            role: "streaming",
            implementation: "litchi_docx::StreamingDocumentWriter",
            timing_scope: "fresh paragraph String generation, public StreamingDocumentWriter creation, forward paragraph/run/text writes, DOCX package finalization, and writer destruction inside the operation clock; sink setup, materialized corpus construction, reopen, exact-member/semantic digest gates, and sink digest finalization are outside",
            performance_claim: "descriptive timing plus process/RSS and allocator observations for the public plain-text writer; 64-byte XML escaping scratch reservation and zero retained sink output; no total-RSS, physical-I/O, native-producer, or broad DOCX creation claim",
            semantic_sha256: corpus.spec.semantic_sha256.clone(),
            full_text_sha256: corpus.spec.full_text_sha256.clone(),
            archive_sha256: corpus.manifest.archive_sha256.clone(),
            target_payload_sha256: expected_target_sha256.clone(),
            archive_member_set_verified: true,
            semantic_reopen_verified: true,
            deterministic_output_verified: true,
            paragraph_count: corpus.spec.paragraphs,
            run_count: corpus.spec.paragraphs,
            input_text_bytes: corpus.spec.input_text_bytes,
            authored_part_bytes: corpus.manifest.target_payload_bytes,
            scratch_bytes: DOCX_STREAMING_SCRATCH_BYTES,
            text_contract: "one plain Unicode text run per paragraph with deterministic café and XML-significant &<> characters",
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
        operation_metrics: Some(operation_metrics),
    })
}

#[cfg(test)]
mod tests {
    use super::*;
    use soapberry_zip::office::StreamingArchiveWriter;

    fn rewritten_archive(
        archive: &[u8],
        mutate: impl FnOnce(&mut Vec<(String, Vec<u8>)>),
    ) -> Vec<u8> {
        let source = OwnedPhysPkgReader::from_bytes(archive.to_vec()).expect("source package");
        let mut members = source
            .member_names()
            .expect("source member names")
            .into_iter()
            .map(|name| {
                let payload = source.read_member(&name).expect("source member payload");
                (name, payload)
            })
            .collect::<Vec<_>>();
        mutate(&mut members);
        let mut writer = StreamingArchiveWriter::new();
        for (name, payload) in members {
            writer
                .write_deflated(&name, &payload)
                .expect("rewritten member");
        }
        writer.finish_to_bytes().expect("rewritten package")
    }

    #[test]
    fn spec_uses_requested_scaling_shapes_and_is_deterministic() {
        for (shape, paragraphs) in [
            (SemanticShape::Tiny, 64),
            (SemanticShape::Medium, 8_192),
            (SemanticShape::Large, 131_072),
        ] {
            let first = build_spec(shape).expect("DOCX streaming spec");
            let second = build_spec(shape).expect("DOCX streaming spec");
            assert_eq!(first, second);
            assert_eq!(first.paragraphs, paragraphs);
            assert_eq!(
                first.input_text_bytes,
                (0..paragraphs)
                    .map(|index| docx_streaming_text(index).len())
                    .sum::<usize>()
            );
        }
    }

    #[test]
    fn tiny_materialized_package_has_exact_members_and_semantics() {
        let corpus = build_corpus(SemanticShape::Tiny).expect("DOCX streaming corpus");
        assert_eq!(corpus.manifest.archive_member_count, 3);
        assert_eq!(corpus.manifest.entry_count, 64);
        assert_eq!(corpus.manifest.target_entry, "word/document.xml");
        let (archive, _) = write_docx_stream(Vec::new(), &corpus.spec).expect("test artifact");
        let target = inspect_materialized_archive(&archive, &corpus.spec).expect("test oracle");
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&archive));
        assert_eq!(corpus.manifest.target_payload_sha256, sha256_hex(&target));
    }

    #[test]
    fn materialized_oracle_rejects_extra_member_changed_text_and_split_run() {
        let corpus = build_corpus(SemanticShape::Tiny).expect("DOCX streaming corpus");

        let (archive, _) = write_docx_stream(Vec::new(), &corpus.spec).expect("test artifact");
        let extra = rewritten_archive(&archive, |members| {
            members.push(("extra.bin".to_owned(), b"unexpected".to_vec()));
        });
        let error = inspect_materialized_archive(&extra, &corpus.spec)
            .expect_err("extra member must fail the package gate");
        assert!(error.to_string().contains("archive members differ"));

        let changed = rewritten_archive(&archive, |members| {
            let (_, document) = members
                .iter_mut()
                .find(|(name, _)| name == "word/document.xml")
                .expect("document member");
            let xml = String::from_utf8(document.clone()).expect("document XML");
            let old = docx_streaming_text(0)
                .replace('&', "&amp;")
                .replace('<', "&lt;")
                .replace('>', "&gt;");
            let changed = old.replacen("000000", "999999", 1);
            assert_ne!(old, changed);
            *document = xml.replacen(&old, &changed, 1).into_bytes();
        });
        let error = inspect_materialized_archive(&changed, &corpus.spec)
            .expect_err("changed paragraph must fail the semantic gate");
        assert!(error.to_string().contains("paragraph 0 differs"));

        let split = rewritten_archive(&archive, |members| {
            let (_, document) = members
                .iter_mut()
                .find(|(name, _)| name == "word/document.xml")
                .expect("document member");
            let xml = String::from_utf8(document.clone()).expect("document XML");
            let old = docx_streaming_text(0)
                .replace('&', "&amp;")
                .replace('<', "&lt;")
                .replace('>', "&gt;");
            let split_at = old.find("café-").expect("split point");
            let first = &old[..split_at];
            let second = &old[split_at..];
            let open = r#"<w:r><w:t xml:space="preserve">"#;
            let close = r#"</w:t></w:r>"#;
            let one_run = format!("{open}{old}{close}");
            let two_runs = format!("{open}{first}{close}{open}{second}{close}");
            *document = xml.replacen(&one_run, &two_runs, 1).into_bytes();
        });
        let error = inspect_materialized_archive(&split, &corpus.spec)
            .expect_err("split run must fail the exact-run gate");
        assert!(error.to_string().contains("has 2 runs, expected one"));
    }

    #[test]
    fn tiny_timed_run_reports_window_counters_and_operation_metrics() {
        let corpus = build_corpus(SemanticShape::Tiny).expect("DOCX streaming corpus");
        let measured = run(Case::DocxStreamingCreate, &corpus, 0, 1).expect("DOCX streaming run");
        assert_eq!(measured.elapsed_ns.samples.len(), 1);
        assert_eq!(
            measured.output_sha256.as_deref(),
            Some(corpus.manifest.archive_sha256.as_str())
        );
        let sink = measured.sink.as_ref().expect("sink summary");
        assert_eq!(sink.accepted_bytes, corpus.manifest.archive_bytes as u64);
        assert_eq!(sink.retained_output_bytes, Some(0));
        assert_eq!(sink.retained_authoring_window_bytes, Some(64));
        assert_eq!(sink.paragraphs, Some(64));
        assert_eq!(sink.runs, Some(64));
        assert_eq!(
            sink.input_bytes,
            Some(corpus.manifest.uncompressed_payload_bytes as u64)
        );
        let source = measured.source.as_ref().expect("source summary");
        let docx = source.docx_paragraphs.as_ref().expect("DOCX summary");
        assert!(docx.archive_member_set_verified);
        assert!(docx.semantic_reopen_verified);
        assert!(docx.deterministic_output_verified);
        assert_eq!(docx.scratch_bytes, 64);
        assert_eq!(measured.operation_metrics.as_ref().unwrap().sample_count, 1);
    }
}
