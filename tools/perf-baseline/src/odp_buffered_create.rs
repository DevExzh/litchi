//! Matched fresh ODP titled-slide creation through the public buffered Builder.
//!
//! Corpus construction and semantic/package gates are outside the measured
//! loop. Each measured iteration creates fresh title/body strings, appends
//! them to a new public Builder, builds the package, and publishes it to a
//! hashing discard sink.

mod content_structure;

use super::{
    Case, CaseResult, Corpus, CorpusManifest, HashingDiscardSink, SemanticShape, SourceSummary,
    allocation_metrics, deterministic_sink_summary, elapsed_ns, iteration_count, operation_metrics,
    process_metrics, record_elapsed, semantic_shape, sha256_hex, statistics,
};
use serde::Serialize;
use sha2::Digest as _;
use soapberry_zip::office::ArchiveReader;
use soapberry_zip::{CompressionMethod, ZipArchive};
use std::{collections::BTreeSet, error::Error, io::Write, time::Instant};

/// Stable corpus identity for the buffered ODP role.
pub(crate) const ODP_BUFFERED_CORPUS_GENERATOR: &str = "litchi-odp-buffered-slides-v1";
pub(super) const ODP_BUFFERED_MIMETYPE: &[u8] = b"application/vnd.oasis.opendocument.presentation";
pub(super) const ODP_BUFFERED_MEMBER_NAMES: [&str; 5] = [
    "mimetype",
    "content.xml",
    "styles.xml",
    "meta.xml",
    "META-INF/manifest.xml",
];
pub(super) const ODP_BUFFERED_COMPRESSION: &str = "mimetype=stored;xml=deflate";
pub(super) const ODP_BUFFERED_DEFAULT_STYLES_BYTES: usize = 1_960;
pub(super) const ODP_BUFFERED_DEFAULT_STYLES_SHA256: &str =
    "d9881e91085516246a19c30d9e5cde39a8b10d7e42120b135f48f5ca8afef8d2";
pub(super) const ODP_BUFFERED_DEFAULT_META_BYTES: usize = 387;
pub(super) const ODP_BUFFERED_DEFAULT_META_SHA256: &str =
    "c7e55a3560c73aa42da85eec4751c3e78b5cc53ff964f50acba6c5cd105e6719";
/// Source evidence for one role-local fresh ODP titled-slide corpus.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct OdpSlidesSummary {
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
    pub(crate) runtime_output_digest_verified: bool,
    pub(crate) runtime_sink_length_verified: bool,
    pub(crate) page_structure_verified: bool,
    pub(crate) page_geometry_verified: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub(crate) provider_input_text_bytes: Option<usize>,
    pub(crate) slide_count: usize,
    pub(crate) title_count: usize,
    pub(crate) body_count: usize,
    pub(crate) title_text_bytes: usize,
    pub(crate) body_text_bytes: usize,
    pub(crate) title_variant_counts: [usize; 4],
    pub(crate) body_variant_counts: [usize; 4],
    pub(crate) text_contract: &'static str,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct OdpBufferedIdentity {
    pub(super) archive_bytes: usize,
    pub(super) archive_member_count: usize,
    pub(super) slide_count: usize,
    pub(super) representative_slide_bytes: usize,
    pub(super) archive_sha256: String,
    pub(super) semantic_sha256: String,
    pub(super) semantic_input_bytes: usize,
    pub(super) target_payload_bytes: usize,
    pub(super) target_payload_sha256: String,
    pub(super) content_xml_bytes: usize,
    pub(super) content_xml_sha256: String,
    pub(super) styles_xml_sha256: String,
    pub(super) meta_xml_sha256: String,
    pub(super) title_text_bytes: usize,
    pub(super) body_text_bytes: usize,
    pub(super) title_variant_counts: [usize; 4],
    pub(super) body_variant_counts: [usize; 4],
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(super) struct OdpTextShapeCounts {
    pub(super) slide_count: usize,
    pub(super) title_count: usize,
    pub(super) body_count: usize,
    pub(super) title_text_bytes: usize,
    pub(super) body_text_bytes: usize,
    pub(super) title_variant_counts: [usize; 4],
    pub(super) body_variant_counts: [usize; 4],
}

pub(super) const fn odp_buffered_slide_count(shape: SemanticShape) -> usize {
    match shape {
        SemanticShape::Tiny => 64,
        SemanticShape::Medium => 4_096,
        SemanticShape::Large => 8_192,
    }
}

const fn odp_buffered_variant(index: usize) -> usize {
    index % 4
}

const fn odp_buffered_variant_text(index: usize) -> &'static str {
    match odp_buffered_variant(index) {
        0 => "plain slide",
        1 => "Unicode café Δ 中",
        2 => "entities <&> \"quoted\"",
        _ => "mixed façade Ω <&> value",
    }
}

/// Return the fixed mixed-text title used by the fresh Builder baseline.
///
/// The corpus intentionally has single interior spaces and no control or edge
/// whitespace. It still exercises ordinary text, non-ASCII UTF-8, and XML
/// significant characters that the Builder must escape and the Presentation
/// facade must restore.
pub(super) fn odp_buffered_title(index: usize) -> String {
    format!(
        "litchi-perf-odp-buffered-title-{index:05} {}",
        odp_buffered_variant_text(index)
    )
}

pub(super) fn odp_buffered_body(index: usize) -> String {
    format!(
        "litchi-perf-odp-buffered-body-{index:05} {}",
        odp_buffered_variant_text(index)
    )
}

pub(super) fn odp_buffered_semantic_digest(shape: SemanticShape) -> Result<String, Box<dyn Error>> {
    let slide_count = odp_buffered_slide_count(shape);
    let mut hasher = sha2::Sha256::new();
    hasher.update(b"litchi-odp-buffered-semantic-v1\0");
    hasher.update(u64::try_from(slide_count)?.to_le_bytes());
    for index in 0..slide_count {
        let title = odp_buffered_title(index);
        let body = odp_buffered_body(index);
        hasher.update(u64::try_from(title.len())?.to_le_bytes());
        hasher.update(title.as_bytes());
        hasher.update(u64::try_from(body.len())?.to_le_bytes());
        hasher.update(body.as_bytes());
    }
    let digest = hasher.finalize();
    let mut output = String::with_capacity(digest.len() * 2);
    for byte in digest {
        use std::fmt::Write as _;
        let _ = write!(output, "{byte:02x}");
    }
    Ok(output)
}

pub(super) fn odp_buffered_semantic_input_bytes(
    shape: SemanticShape,
) -> Result<usize, Box<dyn Error>> {
    let slide_count = odp_buffered_slide_count(shape);
    (0..slide_count).try_fold(0usize, |total, index| {
        let title = odp_buffered_title(index);
        let body = odp_buffered_body(index);
        total
            .checked_add(title.len())
            .and_then(|value| value.checked_add(1))
            .and_then(|value| value.checked_add(body.len()))
            .and_then(|value| {
                if index + 1 < slide_count {
                    value.checked_add(2)
                } else {
                    Some(value)
                }
            })
            .ok_or_else(|| "ODP semantic projection byte count overflows usize".into())
    })
}

pub(super) fn odp_buffered_text_shape_counts(
    shape: SemanticShape,
) -> Result<OdpTextShapeCounts, Box<dyn Error>> {
    let slide_count = odp_buffered_slide_count(shape);
    let mut title_text_bytes = 0usize;
    let mut body_text_bytes = 0usize;
    let mut title_variant_counts = [0usize; 4];
    let mut body_variant_counts = [0usize; 4];
    for index in 0..slide_count {
        let title = odp_buffered_title(index);
        let body = odp_buffered_body(index);
        title_text_bytes = title_text_bytes
            .checked_add(title.len())
            .ok_or("ODP title byte count overflows usize")?;
        body_text_bytes = body_text_bytes
            .checked_add(body.len())
            .ok_or("ODP body byte count overflows usize")?;
        let variant = odp_buffered_variant(index);
        title_variant_counts[variant] = title_variant_counts[variant]
            .checked_add(1)
            .ok_or("ODP title variant count overflows usize")?;
        body_variant_counts[variant] = body_variant_counts[variant]
            .checked_add(1)
            .ok_or("ODP body variant count overflows usize")?;
    }
    Ok(OdpTextShapeCounts {
        slide_count,
        title_count: slide_count,
        body_count: slide_count,
        title_text_bytes,
        body_text_bytes,
        title_variant_counts,
        body_variant_counts,
    })
}

fn odp_buffered_bytes(shape: SemanticShape) -> Result<Vec<u8>, Box<dyn Error>> {
    let mut builder = litchi_odp::Builder::new();
    for index in 0..odp_buffered_slide_count(shape) {
        let title = odp_buffered_title(index);
        let body = odp_buffered_body(index);
        builder.add_slide_with_title(&title, &body)?;
    }
    Ok(builder.build()?)
}

fn verify_odp_buffered_default_parts(
    styles_xml: &[u8],
    meta_xml: &[u8],
) -> Result<(), Box<dyn Error>> {
    if styles_xml.len() != ODP_BUFFERED_DEFAULT_STYLES_BYTES
        || sha256_hex(styles_xml) != ODP_BUFFERED_DEFAULT_STYLES_SHA256
    {
        return Err("buffered ODP styles.xml differs from the pinned Builder default".into());
    }
    if meta_xml.len() != ODP_BUFFERED_DEFAULT_META_BYTES
        || sha256_hex(meta_xml) != ODP_BUFFERED_DEFAULT_META_SHA256
    {
        return Err("buffered ODP meta.xml differs from the pinned Builder default".into());
    }
    Ok(())
}

fn verify_odp_buffered_compression(bytes: &[u8]) -> Result<(), Box<dyn Error>> {
    // ODF's MIME marker is a stored, offset-zero local entry with no extra
    // field. Inspect its fixed framing independently of the semantic reader.
    let marker_end = 38 + ODP_BUFFERED_MIMETYPE.len();
    let header = bytes.get(..marker_end).ok_or("truncated ODP MIME header")?;
    let expected_size = u32::try_from(ODP_BUFFERED_MIMETYPE.len())?.to_le_bytes();
    let flags = u16::from_le_bytes(header[6..8].try_into()?);
    if &header[..4] != b"PK\x03\x04"
        || flags & !0x0800 != 0
        || header[8..10] != [0, 0]
        || header[18..22] != expected_size
        || header[22..26] != expected_size
        || header[26..28] != [8, 0]
        || header[28..30] != [0, 0]
        || &header[30..38] != b"mimetype"
        || &header[38..marker_end] != ODP_BUFFERED_MIMETYPE
    {
        return Err("ODP MIME local-header layout differs from its fixed contract".into());
    }
    let archive = ZipArchive::from_slice(bytes)?;
    let mut file_count = 0usize;
    for entry in archive.entries() {
        let entry = entry?;
        if entry.is_dir() {
            continue;
        }
        file_count = file_count
            .checked_add(1)
            .ok_or("buffered ODP ZIP member count overflows usize")?;
        let path = entry.file_path().try_normalize()?;
        if path.as_str() == "mimetype"
            && (file_count != 1
                || entry.local_header_offset() != 0
                || header[14..18] != entry.crc32().to_le_bytes())
        {
            return Err("ODP MIME central entry is not bound to the first local header".into());
        }
        let expected = if path.as_str() == "mimetype" {
            CompressionMethod::Store
        } else {
            CompressionMethod::Deflate
        };
        if entry.compression_method() != expected {
            return Err(format!(
                "buffered ODP member {} uses {:?}, expected {:?}",
                path.as_str(),
                entry.compression_method(),
                expected
            )
            .into());
        }
    }
    if file_count != ODP_BUFFERED_MEMBER_NAMES.len() {
        return Err("buffered ODP ZIP member count differs from the package contract".into());
    }
    Ok(())
}

fn verify_odp_buffered_manifest(archive: &ArchiveReader<'_>) -> Result<(), Box<dyn Error>> {
    let manifest_xml = String::from_utf8(archive.read("META-INF/manifest.xml")?)?;
    let manifest = litchi_odf_common::core::Manifest::parse(&manifest_xml)?;
    let root = manifest
        .get_entry("/")
        .ok_or("buffered ODP manifest root entry is missing")?;
    let content = manifest
        .get_entry("content.xml")
        .ok_or("buffered ODP manifest content.xml entry is missing")?;
    let styles = manifest
        .get_entry("styles.xml")
        .ok_or("buffered ODP manifest styles.xml entry is missing")?;
    let meta = manifest
        .get_entry("meta.xml")
        .ok_or("buffered ODP manifest meta.xml entry is missing")?;
    if manifest.entries.len() != 4
        || manifest.mimetype != "application/vnd.oasis.opendocument.presentation"
        || root.media_type != "application/vnd.oasis.opendocument.presentation"
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
        return Err("buffered ODP manifest bindings differ from the Builder contract".into());
    }
    Ok(())
}

pub(super) fn inspect_odp_buffered_archive(
    bytes: &[u8],
    shape: SemanticShape,
) -> Result<OdpBufferedIdentity, Box<dyn Error>> {
    let archive = ArchiveReader::new(bytes)?;
    let names = archive.file_names().collect::<Vec<_>>();
    let actual_set = names.iter().copied().collect::<BTreeSet<_>>();
    let expected_set = ODP_BUFFERED_MEMBER_NAMES
        .iter()
        .copied()
        .collect::<BTreeSet<_>>();
    if names.len() != ODP_BUFFERED_MEMBER_NAMES.len() || actual_set != expected_set {
        return Err(format!(
            "buffered ODP archive member set differs: actual={names:?}, expected={ODP_BUFFERED_MEMBER_NAMES:?}"
        )
        .into());
    }
    if archive.read("mimetype")? != ODP_BUFFERED_MIMETYPE {
        return Err("buffered ODP mimetype member differs from the ODP MIME".into());
    }
    verify_odp_buffered_compression(bytes)?;
    verify_odp_buffered_manifest(&archive)?;

    let content_xml = archive.read("content.xml")?;
    let styles_xml = archive.read("styles.xml")?;
    let meta_xml = archive.read("meta.xml")?;
    verify_odp_buffered_default_parts(&styles_xml, &meta_xml)?;
    content_structure::verify_odp_buffered_content_structure(&content_xml, shape)?;

    let presentation = litchi_odp::Presentation::from_bytes(bytes.to_vec())?;
    let slides = presentation.slides()?;
    let expected_slide_count = odp_buffered_slide_count(shape);
    if slides.len() != expected_slide_count {
        return Err("buffered ODP slide count differs from the fixture".into());
    }
    let mut title_text_bytes = 0usize;
    let mut body_text_bytes = 0usize;
    let mut title_variant_counts = [0usize; 4];
    let mut body_variant_counts = [0usize; 4];
    let mut semantic_projection = Vec::new();
    semantic_projection.extend_from_slice(b"litchi-odp-buffered-semantic-v1\0");
    semantic_projection.extend_from_slice(&u64::try_from(slides.len())?.to_le_bytes());
    for (index, slide) in slides.iter().enumerate() {
        let title = slide
            .title()?
            .ok_or("buffered ODP slide title is missing")?;
        let body = slide.text()?;
        let expected_title = odp_buffered_title(index);
        let expected_body = odp_buffered_body(index);
        if slide.index() != index
            || title != expected_title.as_str()
            || body != expected_body.as_str()
            || slide.all_text() != format!("{expected_title}\n{expected_body}")
        {
            return Err("buffered ODP slide semantic projection differs from the fixture".into());
        }
        title_text_bytes = title_text_bytes
            .checked_add(title.len())
            .ok_or("buffered ODP title byte count overflows usize")?;
        body_text_bytes = body_text_bytes
            .checked_add(body.len())
            .ok_or("buffered ODP body byte count overflows usize")?;
        let variant = odp_buffered_variant(index);
        title_variant_counts[variant] = title_variant_counts[variant]
            .checked_add(1)
            .ok_or("buffered ODP title variant count overflows usize")?;
        body_variant_counts[variant] = body_variant_counts[variant]
            .checked_add(1)
            .ok_or("buffered ODP body variant count overflows usize")?;
        semantic_projection.extend_from_slice(&u64::try_from(title.len())?.to_le_bytes());
        semantic_projection.extend_from_slice(title.as_bytes());
        semantic_projection.extend_from_slice(&u64::try_from(body.len())?.to_le_bytes());
        semantic_projection.extend_from_slice(body.as_bytes());
    }
    let expected_text = (0..expected_slide_count)
        .map(|index| {
            format!(
                "{}\n{}",
                odp_buffered_title(index),
                odp_buffered_body(index)
            )
        })
        .collect::<Vec<_>>()
        .join("\n\n");
    if presentation.text()? != expected_text {
        return Err("buffered ODP full-text projection differs from the fixture".into());
    }
    if sha256_hex(&semantic_projection) != odp_buffered_semantic_digest(shape)? {
        return Err("buffered ODP semantic digest differs from the fixture".into());
    }

    Ok(OdpBufferedIdentity {
        archive_bytes: bytes.len(),
        archive_member_count: names.len(),
        slide_count: expected_slide_count,
        representative_slide_bytes: odp_buffered_title(0).len() + 1 + odp_buffered_body(0).len(),
        archive_sha256: sha256_hex(bytes),
        semantic_sha256: odp_buffered_semantic_digest(shape)?,
        semantic_input_bytes: odp_buffered_semantic_input_bytes(shape)?,
        target_payload_bytes: content_xml.len(),
        target_payload_sha256: sha256_hex(&content_xml),
        content_xml_bytes: content_xml.len(),
        content_xml_sha256: sha256_hex(&content_xml),
        styles_xml_sha256: sha256_hex(&styles_xml),
        meta_xml_sha256: sha256_hex(&meta_xml),
        title_text_bytes,
        body_text_bytes,
        title_variant_counts,
        body_variant_counts,
    })
}

pub(super) fn verify_odp_corpus_binding(
    corpus: &Corpus,
    shape: SemanticShape,
    identity: &OdpBufferedIdentity,
    expected_generator: &str,
    expected_name: &str,
) -> Result<(), Box<dyn Error>> {
    let reopened_target = ArchiveReader::new(&corpus.archive)?.read("content.xml")?;
    if corpus.target_name != "content.xml"
        || corpus.manifest.target_entry != "content.xml"
        || corpus.target_payload != reopened_target
    {
        return Err("buffered ODP corpus target entry is not bound to reopened content.xml".into());
    }
    if corpus.manifest.name != expected_name
        || corpus.manifest.generator != expected_generator
        || corpus.manifest.package_format != "ODP/ODF/ZIP"
        || corpus.manifest.shape != shape.name()
        || corpus.manifest.payload_kind
            != "deterministic-mixed-unicode-entities-plain-titled-slides"
        || corpus.manifest.compression != ODP_BUFFERED_COMPRESSION
    {
        return Err("buffered ODP corpus manifest identity differs from its contract".into());
    }
    if corpus.archive.len() != identity.archive_bytes
        || corpus.manifest.archive_bytes != identity.archive_bytes
        || corpus.manifest.archive_member_count != identity.archive_member_count
        || corpus.manifest.entry_count != identity.slide_count
        || corpus.manifest.entry_bytes != identity.representative_slide_bytes
        || corpus.manifest.uncompressed_payload_bytes != identity.semantic_input_bytes
        || corpus.manifest.archive_sha256 != identity.archive_sha256
        || corpus.manifest.target_payload_bytes != identity.target_payload_bytes
        || corpus.manifest.target_payload_sha256 != identity.target_payload_sha256
    {
        return Err("ODP slide corpus manifest is not bound to reopened archive identity".into());
    }
    if corpus.target_payload.len() != identity.target_payload_bytes
        || sha256_hex(&corpus.target_payload) != identity.target_payload_sha256
        || sha256_hex(&corpus.archive) != identity.archive_sha256
        || reopened_target.len() != identity.target_payload_bytes
        || sha256_hex(&reopened_target) != identity.target_payload_sha256
        || identity.semantic_input_bytes != odp_buffered_semantic_input_bytes(shape)?
        || identity.semantic_sha256 != odp_buffered_semantic_digest(shape)?
    {
        return Err("ODP slide corpus payload or semantic projection is not bound".into());
    }
    Ok(())
}

/// Build and fully gate the role-local buffered ODP corpus before warmups.
pub(crate) fn build_odp_buffered_corpus(shape: SemanticShape) -> Result<Corpus, Box<dyn Error>> {
    let archive = odp_buffered_bytes(shape)?;
    let identity = inspect_odp_buffered_archive(&archive, shape)?;
    let target_payload = ArchiveReader::new(&archive)?.read("content.xml")?;
    let corpus = Corpus {
        manifest: CorpusManifest {
            name: format!("odp-buffered-slides-{}", shape.name()),
            generator: ODP_BUFFERED_CORPUS_GENERATOR,
            package_format: "ODP/ODF/ZIP",
            shape: shape.name(),
            payload_kind: "deterministic-mixed-unicode-entities-plain-titled-slides",
            compression: ODP_BUFFERED_COMPRESSION,
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
    verify_odp_corpus_binding(
        &corpus,
        shape,
        &identity,
        ODP_BUFFERED_CORPUS_GENERATOR,
        &format!("odp-buffered-slides-{}", shape.name()),
    )?;
    Ok(corpus)
}

/// Run the buffered ODP authoring baseline.
pub(crate) fn run_odp_buffered_creation(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if case != Case::OdpBufferedCreate || corpus.manifest.generator != ODP_BUFFERED_CORPUS_GENERATOR
    {
        return Err("non-buffered ODP case passed to buffered creation runner".into());
    }
    let shape = semantic_shape(corpus)?;
    let identity = inspect_odp_buffered_archive(&corpus.archive, shape)?;
    verify_odp_corpus_binding(
        corpus,
        shape,
        &identity,
        ODP_BUFFERED_CORPUS_GENERATOR,
        &format!("odp-buffered-slides-{}", shape.name()),
    )?;
    let expected_semantic_sha256 = odp_buffered_semantic_digest(shape)?;
    let expected_input_bytes = odp_buffered_semantic_input_bytes(shape)?;
    let maximum = u64::try_from(corpus.manifest.archive_bytes)?
        .checked_add(64 * 1024)
        .ok_or("buffered ODP sink ceiling overflows")?;
    let text_counts = odp_buffered_text_shape_counts(shape)?;
    let text_object_count = u64::try_from(text_counts.title_count)?
        .checked_add(u64::try_from(text_counts.body_count)?)
        .ok_or("buffered ODP title/body object count overflows u64")?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut summaries = Vec::with_capacity(samples);
    let mut digests = Vec::with_capacity(samples);
    let mut observations = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        // Sink reservation and process/allocator setup are outside the clock.
        // Fresh title/body strings, Builder model construction, package
        // publication, and the hashing sink write are the measured operation.
        let mut sink = HashingDiscardSink::without_authoring_window(maximum);
        let process_before = process_metrics::Snapshot::read().ok();
        let allocation_region = allocation_metrics::begin();
        let started = Instant::now();
        let output = odp_buffered_bytes(shape)?;
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
            return Err("buffered ODP creation digest or sink length differs from corpus".into());
        }
        summary.paragraphs = Some(u64::try_from(text_counts.slide_count)?);
        summary.runs = Some(text_object_count);
        summary.input_bytes = Some(u64::try_from(expected_input_bytes)?);
        summary.authored_part_bytes = Some(u64::try_from(identity.content_xml_bytes)?);
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
    let sink = deterministic_sink_summary(&summaries, "buffered ODP creation")?;
    if sink.retained_output_bytes != Some(0) || sink.retained_authoring_window_bytes.is_some() {
        return Err("buffered ODP creation reported a retained output/window bound".into());
    }
    if digests
        .iter()
        .any(|digest| digest != &corpus.manifest.archive_sha256)
    {
        return Err("buffered ODP creation output digest changed across samples".into());
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
        odp_slides: Some(OdpSlidesSummary {
            role: "buffered",
            implementation: "litchi_odp::Builder::add_slide_with_title/build",
            timing_scope: "fresh title/body String generation, Builder::add_slide_with_title/build, and HashingDiscardSink write; output Vec release occurs before the clock stops and allocator/process endpoint snapshots; corpus setup, reopen, digest, sink finalization, and package/semantic gates are outside",
            performance_claim: "baseline timing and process/RSS/allocator evidence only; no fixed retained-window, throughput, or cross-role lexical-byte claim",
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
            provider_input_text_bytes: None,
            slide_count: text_counts.slide_count,
            title_count: text_counts.title_count,
            body_count: text_counts.body_count,
            title_text_bytes: text_counts.title_text_bytes,
            body_text_bytes: text_counts.body_text_bytes,
            title_variant_counts: text_counts.title_variant_counts,
            body_variant_counts: text_counts.body_variant_counts,
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
    use soapberry_zip::office::StreamingArchiveWriter;

    #[test]
    fn slide_text_projection_scales_with_stable_distinct_hashes() {
        assert_eq!(odp_buffered_slide_count(SemanticShape::Tiny), 64);
        assert_eq!(odp_buffered_slide_count(SemanticShape::Medium), 4_096);
        assert_eq!(odp_buffered_slide_count(SemanticShape::Large), 8_192);
        assert_eq!(
            odp_buffered_semantic_digest(SemanticShape::Tiny).unwrap(),
            odp_buffered_semantic_digest(SemanticShape::Tiny).unwrap()
        );
        assert_ne!(
            odp_buffered_semantic_digest(SemanticShape::Tiny).unwrap(),
            odp_buffered_semantic_digest(SemanticShape::Medium).unwrap()
        );
        assert_eq!(
            odp_buffered_semantic_digest(SemanticShape::Tiny).unwrap(),
            "167428219fc1603b63046c349a06ef8d0033ac9f720567c5383bb79e379d10d1"
        );
        assert_eq!(
            odp_buffered_semantic_input_bytes(SemanticShape::Tiny).unwrap(),
            7_358
        );
        assert!(odp_buffered_title(1).contains("café Δ 中"));
        assert!(odp_buffered_body(2).contains("<&>"));
    }

    #[test]
    fn oversized_titled_deck_preserves_the_default_xml_attribute_refusal() {
        let mut builder = litchi_odp::Builder::new();
        for index in 0..32_768 {
            builder
                .add_slide_with_title(&odp_buffered_title(index), &odp_buffered_body(index))
                .unwrap();
        }
        let error = builder.build().unwrap_err();
        let litchi_core::Error::InvalidFormat(message) = error else {
            panic!("unexpected oversized deck refusal: {error}");
        };
        assert!(message.contains("XML publication rejected for 'content.xml'"));
        assert!(message.contains("XML Attributes limit 250000 exceeded by 250001"));
        println!("32768-slide buffered refusal: {message}");
    }

    #[test]
    fn buffered_tiny_archive_has_exact_package_and_semantic_oracle() {
        let archive = odp_buffered_bytes(SemanticShape::Tiny).unwrap();
        let identity = inspect_odp_buffered_archive(&archive, SemanticShape::Tiny).unwrap();
        assert_eq!(identity.slide_count, 64);
        assert_eq!(identity.title_text_bytes, 3_616);
        assert_eq!(identity.body_text_bytes, 3_552);
        assert_eq!(identity.title_variant_counts, [16; 4]);
        assert_eq!(identity.body_variant_counts, [16; 4]);
        assert_eq!(
            identity.semantic_input_bytes,
            odp_buffered_semantic_input_bytes(SemanticShape::Tiny).unwrap()
        );
        assert_eq!(identity.archive_sha256, sha256_hex(&archive));
        assert!(identity.content_xml_bytes > identity.semantic_input_bytes);
        assert_ne!(identity.styles_xml_sha256, identity.meta_xml_sha256);
    }

    #[test]
    fn mime_header_gate_rejects_extra_fields_descriptors_and_reordered_members() {
        let original = odp_buffered_bytes(SemanticShape::Tiny).unwrap();
        for (offset, value) in [(6, 8), (8, 8), (28, 1)] {
            let mut mutated = original.clone();
            mutated[offset] = value;
            assert!(verify_odp_buffered_compression(&mutated).is_err());
        }
        let reader = ArchiveReader::new(&original).unwrap();
        let mut writer = StreamingArchiveWriter::new();
        for path in [
            "content.xml",
            "mimetype",
            "styles.xml",
            "meta.xml",
            "META-INF/manifest.xml",
        ] {
            let content = reader.read(path).unwrap();
            if path == "mimetype" {
                writer.write_stored(path, &content).unwrap();
            } else {
                writer.write_deflated(path, &content).unwrap();
            }
        }
        let reordered = writer.finish_to_bytes().unwrap();
        assert!(verify_odp_buffered_compression(&reordered).is_err());
    }

    #[test]
    fn buffered_oracle_rejects_manifest_binding_mutation() {
        let mut corpus = build_odp_buffered_corpus(SemanticShape::Tiny).unwrap();
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
        let error = run_odp_buffered_creation(Case::OdpBufferedCreate, &corpus, 0, 1)
            .err()
            .expect("buffered ODP runner accepted a mutated manifest binding")
            .to_string();
        assert!(
            error.contains("manifest") || error.contains("content.xml"),
            "manifest mutation failed for the wrong reason: {error}"
        );
    }

    #[test]
    fn buffered_oracle_rejects_semantic_content_mutation() {
        let mut corpus = build_odp_buffered_corpus(SemanticShape::Tiny).unwrap();
        let reader = ArchiveReader::new(&corpus.archive).unwrap();
        let mut writer = StreamingArchiveWriter::new();
        for name in reader.file_names() {
            let mut payload = reader.read(name).unwrap();
            if name == "content.xml" {
                let content = String::from_utf8(payload).unwrap();
                let mutated = content.replacen(&odp_buffered_title(0), "corrupt-title", 1);
                assert_ne!(mutated, content);
                payload = mutated.into_bytes();
            }
            if name == "mimetype" {
                writer.write_stored(name, &payload).unwrap();
            } else {
                writer.write_deflated(name, &payload).unwrap();
            }
        }
        corpus.archive = writer.finish_to_bytes().unwrap();
        let error = run_odp_buffered_creation(Case::OdpBufferedCreate, &corpus, 0, 1)
            .err()
            .expect("buffered ODP runner accepted a mutated semantic content part")
            .to_string();
        assert!(
            error.contains("semantic") || error.contains("slide"),
            "semantic mutation failed for the wrong reason: {error}"
        );
    }
}
