//! Opt-in fresh PPTX slide creation through the public streaming writer.
//!
//! The materialized package is built once outside the timed loop and is used
//! only for deterministic semantic and physical-package gates. Timed
//! iterations write to the scalar-only [`HashingDiscardSink`], so the
//! generated archive is not retained by the operation.

use super::{
    Case, CaseResult, CorpusManifest, HashingDiscardSink, SemanticShape, SourceSummary,
    allocation_metrics, deterministic_sink_summary, elapsed_ns, iteration_count, operation_metrics,
    process_metrics, record_elapsed, sha256_hex, statistics,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::phys_pkg::OwnedPhysPkgReader;
use litchi_opc::{PackURI, Part, TargetMode};
use litchi_pptx::{
    StreamingPresentationLimits, StreamingPresentationOptions, StreamingPresentationWriter,
    TextBoxSpec,
};
use serde::Serialize;
use sha2::{Digest as _, Sha256};
use std::{error::Error, io::Write, time::Instant};

/// Stable corpus identity for one role-local fresh PPTX stream corpus.
pub(crate) const PPTX_STREAMING_CORPUS_GENERATOR: &str =
    "litchi-pptx-streaming-plaintext-slides-v1";

const PPTX_FIXED_MEMBER_COUNT: usize = 37;
const PPTX_SLIDE_MEMBERS_PER_SLIDE: usize = 2;
const PPTX_STATIC_MEMBER_NAMES: [&str; PPTX_FIXED_MEMBER_COUNT] = [
    "[Content_Types].xml",
    "_rels/.rels",
    "docProps/core.xml",
    "docProps/app.xml",
    "ppt/presProps.xml",
    "ppt/viewProps.xml",
    "ppt/tableStyles.xml",
    "ppt/slideMasters/slideMaster1.xml",
    "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    "ppt/slideLayouts/slideLayout1.xml",
    "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    "ppt/slideLayouts/slideLayout2.xml",
    "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
    "ppt/slideLayouts/slideLayout3.xml",
    "ppt/slideLayouts/_rels/slideLayout3.xml.rels",
    "ppt/slideLayouts/slideLayout4.xml",
    "ppt/slideLayouts/_rels/slideLayout4.xml.rels",
    "ppt/slideLayouts/slideLayout5.xml",
    "ppt/slideLayouts/_rels/slideLayout5.xml.rels",
    "ppt/slideLayouts/slideLayout6.xml",
    "ppt/slideLayouts/_rels/slideLayout6.xml.rels",
    "ppt/slideLayouts/slideLayout7.xml",
    "ppt/slideLayouts/_rels/slideLayout7.xml.rels",
    "ppt/slideLayouts/slideLayout8.xml",
    "ppt/slideLayouts/_rels/slideLayout8.xml.rels",
    "ppt/slideLayouts/slideLayout9.xml",
    "ppt/slideLayouts/_rels/slideLayout9.xml.rels",
    "ppt/slideLayouts/slideLayout10.xml",
    "ppt/slideLayouts/_rels/slideLayout10.xml.rels",
    "ppt/slideLayouts/slideLayout11.xml",
    "ppt/slideLayouts/_rels/slideLayout11.xml.rels",
    "ppt/theme/theme1.xml",
    "ppt/theme/theme2.xml",
    "ppt/notesMasters/notesMaster1.xml",
    "ppt/notesMasters/_rels/notesMaster1.xml.rels",
    "ppt/presentation.xml",
    "ppt/_rels/presentation.xml.rels",
];

const PPTX_TEXT_X: i64 = 914_400;
const PPTX_TEXT_Y: i64 = 914_400;
const PPTX_TEXT_WIDTH: i64 = 7_315_200;
const PPTX_TEXT_HEIGHT: i64 = 914_400;
const PPTX_MAX_TEXT_BYTES_PER_BOX: usize = 1 << 20;
const PPTX_MAX_TOTAL_TEXT_BYTES: usize = 16 << 20;
const PPTX_MAX_OUTPUT_HEADROOM: u64 = 4 * 1024 * 1024;
const PPTX_OUTPUT_BYTES_PER_SLIDE: u64 = 8 * 1024;
const PPTX_MAX_SLIDE_XML_HEADROOM: u64 = 16 * 1024;
/// Source and semantic identity for one role-local fresh PPTX stream corpus.
#[derive(Clone, Debug, Serialize)]
pub(crate) struct PptxSlidesSummary {
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
    pub(crate) slide_count: usize,
    pub(crate) text_box_count: usize,
    pub(crate) input_text_bytes: usize,
    pub(crate) authored_part_bytes: usize,
    pub(crate) observed_max_slide_xml_bytes: usize,
    pub(crate) max_slide_xml_bytes: usize,
    pub(crate) structural_metadata_fixed_member_count: usize,
    pub(crate) structural_metadata_members_per_slide: usize,
    pub(crate) structural_metadata_scope: &'static str,
    pub(crate) text_contract: &'static str,
}

#[derive(Clone, Debug)]
pub(crate) struct PptxStreamingCorpus {
    pub(crate) manifest: CorpusManifest,
    spec: PptxStreamingSpec,
    observed_max_slide_xml_bytes: usize,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub(crate) struct PptxStreamingSpec {
    slides: usize,
    input_text_bytes: usize,
    max_text_bytes: usize,
    max_slide_xml_bytes: usize,
    semantic_sha256: String,
    full_text_sha256: String,
}

#[derive(Debug)]
pub(crate) struct PptxArchiveIdentity {
    target_payload: Vec<u8>,
    observed_max_slide_xml_bytes: usize,
}

fn pptx_streaming_slide_count(shape: SemanticShape) -> usize {
    match shape {
        SemanticShape::Tiny => 8,
        SemanticShape::Medium => 256,
        SemanticShape::Large => 8_192,
    }
}

fn pptx_streaming_text(index: usize) -> String {
    format!("litchi-perf-pptx-streaming-v1-{index:06}-café-&<>")
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

fn update_semantic_digest(
    hasher: &mut Sha256,
    index: usize,
    text: &str,
) -> Result<(), Box<dyn Error>> {
    hasher.update(u64::try_from(index)?.to_le_bytes());
    hasher.update(u64::try_from(text.len())?.to_le_bytes());
    hasher.update(text.as_bytes());
    for coordinate in [PPTX_TEXT_X, PPTX_TEXT_Y, PPTX_TEXT_WIDTH, PPTX_TEXT_HEIGHT] {
        hasher.update(coordinate.to_le_bytes());
    }
    Ok(())
}

fn build_spec(shape: SemanticShape) -> Result<PptxStreamingSpec, Box<dyn Error>> {
    let slides = pptx_streaming_slide_count(shape);
    let mut input_text_bytes = 0usize;
    let mut max_text_bytes = 0usize;
    let mut semantic = Sha256::new();
    semantic.update(b"litchi-pptx-streaming-semantic-v1\0");
    semantic.update(u64::try_from(slides)?.to_le_bytes());
    let mut full_text = Sha256::new();
    full_text.update(b"litchi-pptx-streaming-full-text-v1\0");
    full_text.update(u64::try_from(slides)?.to_le_bytes());
    for index in 0..slides {
        let text = pptx_streaming_text(index);
        input_text_bytes = input_text_bytes
            .checked_add(text.len())
            .ok_or("PPTX streaming input byte count overflows usize")?;
        max_text_bytes = max_text_bytes.max(text.len());
        update_semantic_digest(&mut semantic, index, &text)?;
        full_text.update(text.as_bytes());
    }
    if max_text_bytes > PPTX_MAX_TEXT_BYTES_PER_BOX {
        return Err("PPTX streaming fixture exceeds its text-box limit".into());
    }
    if input_text_bytes > PPTX_MAX_TOTAL_TEXT_BYTES {
        return Err("PPTX streaming fixture exceeds its total-text limit".into());
    }
    let escaped_upper_bound = u64::try_from(max_text_bytes)?
        .checked_mul(5)
        .ok_or("PPTX streaming escaped text bound overflows")?;
    let max_slide_xml_bytes = escaped_upper_bound
        .checked_add(PPTX_MAX_SLIDE_XML_HEADROOM)
        .ok_or("PPTX streaming slide XML bound overflows")?;
    let max_slide_xml_bytes = usize::try_from(max_slide_xml_bytes)?;
    Ok(PptxStreamingSpec {
        slides,
        input_text_bytes,
        max_text_bytes,
        max_slide_xml_bytes,
        semantic_sha256: digest_hex(semantic),
        full_text_sha256: digest_hex(full_text),
    })
}

fn write_limits(spec: &PptxStreamingSpec) -> Result<StreamingPresentationLimits, Box<dyn Error>> {
    let max_output_bytes = PPTX_MAX_OUTPUT_HEADROOM
        .checked_add(
            u64::try_from(spec.slides)?
                .checked_mul(PPTX_OUTPUT_BYTES_PER_SLIDE)
                .ok_or("PPTX streaming output bound overflows")?,
        )
        .ok_or("PPTX streaming output bound overflows")?;
    Ok(StreamingPresentationLimits {
        max_slides: spec.slides,
        max_text_boxes_per_slide: 1,
        max_text_bytes_per_box: spec.max_text_bytes,
        max_total_text_bytes: spec.input_text_bytes,
        max_slide_xml_bytes: spec.max_slide_xml_bytes,
        max_output_bytes,
    })
}

fn write_pptx_stream<W: Write>(sink: W, spec: &PptxStreamingSpec) -> Result<W, Box<dyn Error>> {
    let limits = write_limits(spec)?;
    let writer = StreamingPresentationWriter::with_options(
        sink,
        spec.slides,
        StreamingPresentationOptions::standard(),
        limits,
    )?;
    finish_pptx_stream(writer, spec)
}

/// Complete the deterministic slide/text-box sequence after a caller has
/// selected the public writer constructor.  The constructor itself is kept
/// outside this helper so diagnostics can compare the ordinary and explicit
/// metadata-spool APIs while sharing this exact corpus producer.
pub(crate) fn finish_pptx_stream<W: Write>(
    mut writer: StreamingPresentationWriter<W>,
    spec: &PptxStreamingSpec,
) -> Result<W, Box<dyn Error>> {
    let mut expected_text_bytes = 0usize;
    for index in 0..spec.slides {
        let mut slide = writer.start_slide(None)?;
        let text = pptx_streaming_text(index);
        slide.write_text_box(TextBoxSpec::new(
            &text,
            PPTX_TEXT_X,
            PPTX_TEXT_Y,
            PPTX_TEXT_WIDTH,
            PPTX_TEXT_HEIGHT,
        ))?;
        expected_text_bytes = expected_text_bytes
            .checked_add(text.len())
            .ok_or("PPTX streaming expected text byte count overflows")?;
        if slide.text_box_count() != 1 {
            return Err("PPTX streaming slide counter differs from its fixture".into());
        }
        writer = slide.finish()?;
        if writer.completed_slides() != index + 1
            || writer.total_text_bytes() != expected_text_bytes
        {
            return Err("PPTX streaming writer counters differ from its fixture".into());
        }
    }
    if writer.completed_slides() != spec.slides
        || writer.total_text_bytes() != spec.input_text_bytes
    {
        return Err("PPTX streaming final counters differ from its fixture".into());
    }
    Ok(writer.finish()?)
}

/// Return the opaque writer specification attached to a validated corpus.
/// Keeping the fields private prevents a diagnostic from creating a corpus
/// that bypasses the existing text, geometry, and limit calculations.
pub(crate) fn corpus_spec(corpus: &PptxStreamingCorpus) -> &PptxStreamingSpec {
    &corpus.spec
}

pub(crate) fn corpus_archive_sha256(corpus: &PptxStreamingCorpus) -> &str {
    &corpus.manifest.archive_sha256
}

pub(crate) fn corpus_archive_bytes(corpus: &PptxStreamingCorpus) -> usize {
    corpus.manifest.archive_bytes
}

pub(crate) fn corpus_archive_member_count(corpus: &PptxStreamingCorpus) -> usize {
    corpus.manifest.archive_member_count
}

pub(crate) fn corpus_manifest_slide_count(corpus: &PptxStreamingCorpus) -> usize {
    corpus.manifest.entry_count
}

pub(crate) fn corpus_source_bytes(corpus: &PptxStreamingCorpus) -> usize {
    corpus.manifest.uncompressed_payload_bytes
}

pub(crate) fn corpus_semantic_sha256(corpus: &PptxStreamingCorpus) -> &str {
    &corpus.spec.semantic_sha256
}

pub(crate) fn corpus_full_text_sha256(corpus: &PptxStreamingCorpus) -> &str {
    &corpus.spec.full_text_sha256
}

/// Build the existing deterministic PPTX corpus for one of its public
/// scaling counts.  The count mapping is deliberately explicit so the
/// metadata-spool diagnostic cannot silently drift to a different fixture.
pub(crate) fn build_corpus_for_slide_count(
    slide_count: usize,
) -> Result<PptxStreamingCorpus, Box<dyn Error>> {
    let shape = match slide_count {
        8 => SemanticShape::Tiny,
        256 => SemanticShape::Medium,
        8_192 => SemanticShape::Large,
        _ => return Err(format!("unsupported PPTX streaming slide count: {slide_count}").into()),
    };
    build_corpus(shape)
}

/// Compute the same finite limits used by the existing PPTX streaming
/// producer.  This is kept beside the corpus specification so the new
/// metadata-spool constructor receives exactly the old writer limits.
pub(crate) fn corpus_limits(
    corpus: &PptxStreamingCorpus,
) -> Result<StreamingPresentationLimits, Box<dyn Error>> {
    write_limits(&corpus.spec)
}

/// Re-run the complete physical member, semantic slide, geometry, and
/// relationship graph oracle used by the original PPTX streaming diagnostic.
/// The returned identity is intentionally opaque; callers only need the
/// validation side effect and the archive digest they compute at their own
/// boundary.
pub(crate) fn validate_materialized_archive(
    archive: &[u8],
    corpus: &PptxStreamingCorpus,
) -> Result<PptxArchiveIdentity, Box<dyn Error>> {
    inspect_materialized_archive(archive, &corpus.spec)
}

fn expected_member_names(slides: usize) -> Result<Vec<String>, Box<dyn Error>> {
    let additional = slides
        .checked_mul(PPTX_SLIDE_MEMBERS_PER_SLIDE)
        .ok_or("PPTX streaming member count overflows usize")?;
    let capacity = PPTX_FIXED_MEMBER_COUNT
        .checked_add(additional)
        .ok_or("PPTX streaming member count overflows usize")?;
    let mut names = Vec::with_capacity(capacity);
    names.extend(
        PPTX_STATIC_MEMBER_NAMES
            .iter()
            .map(|name| (*name).to_owned()),
    );
    for index in 1..=slides {
        names.push(format!("ppt/slides/slide{index}.xml"));
        names.push(format!("ppt/slides/_rels/slide{index}.xml.rels"));
    }
    Ok(names)
}

fn require_internal_relationship(
    part: &dyn Part,
    relationship_id: &str,
    relationship_type: &str,
    target_reference: &str,
    target_member: &str,
    owner: &str,
) -> Result<PackURI, Box<dyn Error>> {
    let relationship = part.rels().get(relationship_id).ok_or_else(|| {
        format!("PPTX streaming {owner} relationship {relationship_id} is missing")
    })?;
    if relationship.reltype() != relationship_type
        || relationship.target_mode() != TargetMode::Internal
        || relationship.target_ref() != target_reference
    {
        return Err(format!(
            "PPTX streaming {owner} relationship {relationship_id} has an unexpected type, mode, or target"
        )
        .into());
    }
    let target = relationship.target_partname()?;
    if target.membername() != target_member {
        return Err(format!(
            "PPTX streaming {owner} relationship {relationship_id} resolves to {}, expected {target_member}",
            target.membername()
        )
        .into());
    }
    Ok(target)
}

fn verify_presentation_graph(
    presentation: &litchi_pptx::presentation::Presentation<'_>,
    slides: usize,
) -> Result<(), Box<dyn Error>> {
    let references = presentation.slide_references()?;
    if references.len() != slides {
        return Err(
            "PPTX streaming presentation slide-reference count differs from its fixture".into(),
        );
    }
    for (index, reference) in references.iter().enumerate() {
        let expected_id = 256u32
            .checked_add(u32::try_from(index)?)
            .ok_or("PPTX streaming slide reference ID overflows")?;
        let expected_relationship = format!("rId{}", 4 + index);
        if reference.id() != expected_id
            || reference.relationship_id() != expected_relationship.as_str()
        {
            return Err(format!(
                "PPTX streaming slide reference {index} differs: id={} relationship={}",
                reference.id(),
                reference.relationship_id()
            )
            .into());
        }
    }
    let masters = presentation.slide_masters()?;
    if masters.len() != 1 {
        return Err("PPTX streaming presentation must resolve one slide master".into());
    }
    let presentation_part = presentation.part().part();
    let master_references = presentation.part().slide_master_references()?;
    if master_references.len() != 1 || master_references[0] != "rId1" {
        return Err(
            "PPTX streaming presentation master-reference list differs from its fixture".into(),
        );
    }
    let master = masters
        .first()
        .ok_or("PPTX streaming slide master is missing")?;
    let master_part = master.part().part();
    let _ = require_internal_relationship(
        presentation_part,
        "rId1",
        rt::SLIDE_MASTER,
        "slideMasters/slideMaster1.xml",
        "ppt/slideMasters/slideMaster1.xml",
        "presentation slide master",
    )?;
    let _ = require_internal_relationship(
        master_part,
        "rId12",
        rt::THEME,
        "../theme/theme1.xml",
        "ppt/theme/theme1.xml",
        "slide master theme",
    )?;
    if master.theme()?.is_none() {
        return Err("PPTX streaming slide master theme relationship is missing".into());
    }
    let master_layouts = master.layouts()?;
    if master_layouts.len() != 11 {
        return Err("PPTX streaming slide master layout count differs from its fixture".into());
    }
    for index in 1..=11 {
        let relationship_id = format!("rId{index}");
        let target_reference = format!("../slideLayouts/slideLayout{index}.xml");
        let target_member = format!("ppt/slideLayouts/slideLayout{index}.xml");
        let _ = require_internal_relationship(
            master_part,
            &relationship_id,
            rt::SLIDE_LAYOUT,
            &target_reference,
            &target_member,
            "slide master layout",
        )?;
    }
    let layouts = presentation.slide_layouts()?;
    if layouts.len() != 11 {
        return Err("PPTX streaming presentation layout count differs from its fixture".into());
    }
    for layout in &layouts {
        let _ = layout.master()?;
        let _ = require_internal_relationship(
            layout.part().part(),
            "rId1",
            rt::SLIDE_MASTER,
            "../slideMasters/slideMaster1.xml",
            "ppt/slideMasters/slideMaster1.xml",
            "slide layout master",
        )?;
    }
    let notes_master = require_internal_relationship(
        presentation_part,
        "rIdNotesMaster",
        rt::NOTES_MASTER,
        "notesMasters/notesMaster1.xml",
        "ppt/notesMasters/notesMaster1.xml",
        "presentation notes master",
    )?;
    let notes_master_part = presentation.package().get_part(&notes_master)?;
    if notes_master_part.content_type() != ct::PML_NOTES_MASTER {
        return Err("PPTX streaming notes master content type differs from its fixture".into());
    }
    let _ = require_internal_relationship(
        notes_master_part,
        "rId1",
        rt::THEME,
        "../theme/theme2.xml",
        "ppt/theme/theme2.xml",
        "notes master theme",
    )?;
    if presentation_part.rel_ref_count("rIdNotesMaster") != 1 {
        return Err(
            "PPTX streaming presentation notes-master ID reference differs from its fixture".into(),
        );
    }
    // `Presentation::notes` materializes every notes-slide owner and has a
    // deliberate 4,096-slide policy bound. The large 8,192-slide shape is
    // still checked through the same typed OPC relationships above; smaller
    // shapes additionally exercise the public notes graph itself.
    if slides <= 4_096 && presentation.notes()?.is_none() {
        return Err("PPTX streaming notes master graph is missing".into());
    }
    Ok(())
}

fn inspect_materialized_archive(
    archive: &[u8],
    spec: &PptxStreamingSpec,
) -> Result<PptxArchiveIdentity, Box<dyn Error>> {
    let physical = OwnedPhysPkgReader::from_bytes(archive.to_vec())?;
    let expected_names = expected_member_names(spec.slides)?;
    let member_names = physical.member_names()?;
    if member_names != expected_names {
        return Err(format!(
            "PPTX streaming archive members differ: actual count={} expected count={}",
            member_names.len(),
            expected_names.len()
        )
        .into());
    }
    let expected_member_count = PPTX_FIXED_MEMBER_COUNT
        .checked_add(
            spec.slides
                .checked_mul(PPTX_SLIDE_MEMBERS_PER_SLIDE)
                .ok_or("PPTX streaming member count overflows usize")?,
        )
        .ok_or("PPTX streaming member count overflows usize")?;
    if member_names.len() != expected_member_count {
        return Err("PPTX streaming archive member count differs from its fixture".into());
    }

    let package = litchi_pptx::Package::from_bytes(archive)?;
    let presentation = package.presentation()?;
    if presentation.slide_count()? != spec.slides {
        return Err("PPTX streaming slide count differs from its fixture".into());
    }
    if presentation.slide_size()? != (9_144_000, 6_858_000) {
        return Err("PPTX streaming slide dimensions differ from its fixture".into());
    }
    verify_presentation_graph(&presentation, spec.slides)?;
    let mut semantic = Sha256::new();
    semantic.update(b"litchi-pptx-streaming-semantic-v1\0");
    semantic.update(u64::try_from(spec.slides)?.to_le_bytes());
    let mut full_text = Sha256::new();
    full_text.update(b"litchi-pptx-streaming-full-text-v1\0");
    full_text.update(u64::try_from(spec.slides)?.to_le_bytes());
    let mut observed_max_slide_xml_bytes = 0usize;
    let semantic_slides = presentation.slides()?;
    if semantic_slides.len() != spec.slides {
        return Err("PPTX streaming semantic slide list differs from its fixture".into());
    }
    for (index, slide) in semantic_slides.into_iter().enumerate() {
        let shape_count = slide.shape_count()?;
        if shape_count != 1 {
            return Err(format!(
                "PPTX streaming slide {index} has {} shapes, expected one",
                shape_count
            )
            .into());
        }
        let layout = slide.layout()?.ok_or_else(|| {
            format!("PPTX streaming slide {index} layout relationship is missing")
        })?;
        let _ = layout.master()?;
        if layout.part().part().partname().membername() != "ppt/slideLayouts/slideLayout1.xml" {
            return Err(format!(
                "PPTX streaming slide {index} layout resolves to {}, expected ppt/slideLayouts/slideLayout1.xml",
                layout.part().part().partname().membername()
            )
            .into());
        }
        let expected = pptx_streaming_text(index);
        if slide.text()? != expected {
            return Err(
                format!("PPTX streaming slide {index} text differs from its fixture").into(),
            );
        }
        let shapes = slide.shapes()?;
        let shape = shapes.shape(0)?;
        if shape.text() != Some(expected.as_str()) {
            return Err(format!(
                "PPTX streaming slide {index} shape text differs from its fixture"
            )
            .into());
        }
        let bounds = shape
            .bounds()
            .ok_or_else(|| format!("PPTX streaming slide {index} shape bounds are missing"))?;
        let actual_bounds = (bounds.x(), bounds.y(), bounds.width(), bounds.height());
        let expected_bounds = (PPTX_TEXT_X, PPTX_TEXT_Y, PPTX_TEXT_WIDTH, PPTX_TEXT_HEIGHT);
        if actual_bounds != expected_bounds {
            return Err(format!(
                "PPTX streaming slide {index} geometry differs: actual={actual_bounds:?} expected={expected_bounds:?}"
            )
            .into());
        }
        update_semantic_digest(&mut semantic, index, &expected)?;
        full_text.update(expected.as_bytes());
        let slide_xml = physical.read_member(&format!("ppt/slides/slide{}.xml", index + 1))?;
        observed_max_slide_xml_bytes = observed_max_slide_xml_bytes.max(slide_xml.len());
    }
    if digest_hex(semantic) != spec.semantic_sha256 {
        return Err("PPTX streaming slide semantic digest differs from its fixture".into());
    }
    if digest_hex(full_text) != spec.full_text_sha256 {
        return Err("PPTX streaming full-text digest differs from its fixture".into());
    }
    let target_payload = physical.read_member("ppt/presentation.xml")?;
    Ok(PptxArchiveIdentity {
        target_payload,
        observed_max_slide_xml_bytes,
    })
}

/// Build and fully gate one role-local materialized PPTX streaming corpus.
pub(crate) fn build_corpus(shape: SemanticShape) -> Result<PptxStreamingCorpus, Box<dyn Error>> {
    let spec = build_spec(shape)?;
    let archive = write_pptx_stream(Vec::new(), &spec)?;
    let identity = inspect_materialized_archive(&archive, &spec)?;
    if identity.observed_max_slide_xml_bytes > spec.max_slide_xml_bytes {
        return Err("PPTX streaming materialized slide XML exceeds its finite limit".into());
    }
    let archive_sha256 = sha256_hex(&archive);
    let target_payload_sha256 = sha256_hex(&identity.target_payload);
    let archive_member_count = expected_member_names(spec.slides)?.len();
    let manifest = CorpusManifest {
        name: format!("pptx-streaming-slides-{}", shape.name()),
        generator: PPTX_STREAMING_CORPUS_GENERATOR,
        package_format: "PPTX/OOXML/ZIP",
        shape: shape.name(),
        payload_kind: "deterministic-plain-unicode-xml-significant-text-box-per-slide",
        compression: "deflate",
        entry_count: spec.slides,
        archive_member_count,
        entry_bytes: pptx_streaming_text(0).len(),
        uncompressed_payload_bytes: spec.input_text_bytes,
        archive_bytes: archive.len(),
        archive_sha256,
        target_entry: "ppt/presentation.xml".to_owned(),
        target_payload_bytes: identity.target_payload.len(),
        target_payload_sha256,
        rtf_variant: None,
        xlsx: None,
    };
    let observed_max_slide_xml_bytes = identity.observed_max_slide_xml_bytes;
    drop(identity.target_payload);
    drop(archive);
    Ok(PptxStreamingCorpus {
        manifest,
        spec,
        observed_max_slide_xml_bytes,
    })
}

/// Run the public PPTX streaming creation selector.
pub(crate) fn run(
    case: Case,
    corpus: &PptxStreamingCorpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn Error>> {
    if case != Case::PptxStreamingCreate
        || corpus.manifest.generator != PPTX_STREAMING_CORPUS_GENERATOR
    {
        return Err("non-streaming PPTX case passed to PPTX streaming runner".into());
    }
    let expected_archive_sha256 = &corpus.manifest.archive_sha256;
    let expected_target_sha256 = &corpus.manifest.target_payload_sha256;
    let expected_target_bytes = u64::try_from(corpus.manifest.target_payload_bytes)?;
    let maximum = u64::try_from(corpus.manifest.archive_bytes)?
        .checked_add(64 * 1024)
        .ok_or("PPTX streaming sink ceiling overflows")?;
    let mut elapsed = Vec::with_capacity(samples);
    let mut summaries = Vec::with_capacity(samples);
    let mut digests = Vec::with_capacity(samples);
    let mut observations = Vec::with_capacity(samples);
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let sink = HashingDiscardSink::without_authoring_window(maximum);
        let process_before = process_metrics::Snapshot::read().ok();
        let allocation_region = allocation_metrics::begin();
        let started = Instant::now();
        let sink = write_pptx_stream(sink, &corpus.spec)?;
        let duration = started.elapsed();
        let allocation_metrics = allocation_region.finish();
        let process_after = process_metrics::Snapshot::read().ok();
        let process_metrics = process_before
            .zip(process_after)
            .map(|(before, after)| after.delta(before));
        let (mut summary, digest) = sink.finish();
        if digest.as_str() != expected_archive_sha256
            || summary.accepted_bytes != u64::try_from(corpus.manifest.archive_bytes)?
        {
            return Err(
                "PPTX streaming output digest or sink length differs from untimed artifact".into(),
            );
        }
        summary.input_bytes = Some(u64::try_from(corpus.spec.input_text_bytes)?);
        summary.authored_part_bytes = Some(expected_target_bytes);
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
    let sink = deterministic_sink_summary(&summaries, "streaming PPTX creation")?;
    if sink.retained_output_bytes != Some(0) || sink.retained_authoring_window_bytes.is_some() {
        return Err(
            "PPTX streaming creation retained output or reported an authoring window".into(),
        );
    }
    if digests
        .iter()
        .any(|digest| digest != expected_archive_sha256)
    {
        return Err("PPTX streaming output digest changed across samples".into());
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
        pptx_slides: Some(PptxSlidesSummary {
            role: "streaming",
            implementation: "litchi_pptx::StreamingPresentationWriter",
            timing_scope: "fresh deterministic text String generation, public StreamingPresentationWriter creation, forward slide/text-box writes, PPTX package finalization, OPC part-name validation maps, relationship serialization, and writer destruction inside the operation clock; sink setup, materialized corpus construction, semantic reopen, exact 37+2N-member topology gate, graph/geometry/text digest gate, and sink digest finalization are outside",
            performance_claim: "descriptive timing plus process/RSS and allocator observations for the public plain-text PPTX writer; zero retained sink output and no authoring-window reservation are reported; slide XML byte limits are counters, while ZIP central-directory/member-name storage and OPC part-name validation maps grow with slide count; target presentation-part bytes identify the untimed corpus and are not an all-slide memory bound; no total-RSS, physical-I/O, native-producer, or broad PPTX creation claim",
            semantic_sha256: corpus.spec.semantic_sha256.clone(),
            full_text_sha256: corpus.spec.full_text_sha256.clone(),
            archive_sha256: corpus.manifest.archive_sha256.clone(),
            target_payload_sha256: expected_target_sha256.clone(),
            archive_member_set_verified: true,
            semantic_reopen_verified: true,
            deterministic_output_verified: true,
            slide_count: corpus.spec.slides,
            text_box_count: corpus.spec.slides,
            input_text_bytes: corpus.spec.input_text_bytes,
            authored_part_bytes: corpus.manifest.target_payload_bytes,
            observed_max_slide_xml_bytes: corpus.observed_max_slide_xml_bytes,
            max_slide_xml_bytes: corpus.spec.max_slide_xml_bytes,
            structural_metadata_fixed_member_count: PPTX_FIXED_MEMBER_COUNT,
            structural_metadata_members_per_slide: PPTX_SLIDE_MEMBERS_PER_SLIDE,
            structural_metadata_scope: "mandatory ZIP package topology and member-name/central-directory metadata; payload, descriptors, and compressor state are runtime work",
            text_contract: "one plain Unicode text box per slide with deterministic café and XML-significant &<> characters, no title shape, and fixed standard-slide geometry",
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
        for (shape, slides) in [
            (SemanticShape::Tiny, 8),
            (SemanticShape::Medium, 256),
            (SemanticShape::Large, 8_192),
        ] {
            let first = build_spec(shape).expect("PPTX streaming spec");
            let second = build_spec(shape).expect("PPTX streaming spec");
            assert_eq!(first, second);
            assert_eq!(first.slides, slides);
            assert_eq!(
                first.input_text_bytes,
                (0..slides)
                    .map(|index| pptx_streaming_text(index).len())
                    .sum::<usize>()
            );
        }
    }

    #[test]
    fn tiny_materialized_package_has_exact_members_and_all_semantics() {
        let corpus = build_corpus(SemanticShape::Tiny).expect("PPTX streaming corpus");
        assert_eq!(corpus.manifest.archive_member_count, 37 + 2 * 8);
        assert_eq!(corpus.manifest.entry_count, 8);
        assert_eq!(corpus.manifest.target_entry, "ppt/presentation.xml");
        let archive = write_pptx_stream(Vec::new(), &corpus.spec).expect("test artifact");
        let identity = inspect_materialized_archive(&archive, &corpus.spec).expect("test oracle");
        assert_eq!(corpus.manifest.archive_sha256, sha256_hex(&archive));
        assert_eq!(
            corpus.manifest.target_payload_sha256,
            sha256_hex(&identity.target_payload)
        );
    }

    #[test]
    fn materialized_oracle_rejects_extra_member_changed_text_and_geometry() {
        let corpus = build_corpus(SemanticShape::Tiny).expect("PPTX streaming corpus");
        let archive = write_pptx_stream(Vec::new(), &corpus.spec).expect("test artifact");

        let extra = rewritten_archive(&archive, |members| {
            members.push(("extra.bin".to_owned(), b"unexpected".to_vec()));
        });
        let error = inspect_materialized_archive(&extra, &corpus.spec)
            .expect_err("extra member must fail the package gate");
        assert!(error.to_string().contains("archive members differ"));

        let changed = rewritten_archive(&archive, |members| {
            let (_, slide) = members
                .iter_mut()
                .find(|(name, _)| name == "ppt/slides/slide1.xml")
                .expect("slide member");
            let xml = String::from_utf8(slide.clone()).expect("slide XML");
            let old = pptx_streaming_text(0)
                .replace('&', "&amp;")
                .replace('<', "&lt;")
                .replace('>', "&gt;");
            let changed = old.replacen("000000", "999999", 1);
            assert_ne!(old, changed);
            *slide = xml.replacen(&old, &changed, 1).into_bytes();
        });
        let error = inspect_materialized_archive(&changed, &corpus.spec)
            .expect_err("changed text must fail the semantic gate");
        assert!(error.to_string().contains("slide 0 text differs"));

        let changed_geometry = rewritten_archive(&archive, |members| {
            let (_, slide) = members
                .iter_mut()
                .find(|(name, _)| name == "ppt/slides/slide1.xml")
                .expect("slide member");
            let xml = String::from_utf8(slide.clone()).expect("slide XML");
            let changed = xml.replacen("x=\"914400\"", "x=\"914401\"", 1);
            assert_ne!(xml, changed);
            *slide = changed.into_bytes();
        });
        let error = inspect_materialized_archive(&changed_geometry, &corpus.spec)
            .expect_err("changed geometry must fail the semantic gate");
        assert!(error.to_string().contains("geometry differs"));

        let changed_layout = rewritten_archive(&archive, |members| {
            let (_, relationships) = members
                .iter_mut()
                .find(|(name, _)| name == "ppt/slides/_rels/slide1.xml.rels")
                .expect("slide relationship member");
            let xml = String::from_utf8(relationships.clone()).expect("slide relationships XML");
            let changed = xml.replacen(
                "../slideLayouts/slideLayout1.xml",
                "../slideLayouts/slideLayout2.xml",
                1,
            );
            assert_ne!(xml, changed);
            *relationships = changed.into_bytes();
        });
        let error = inspect_materialized_archive(&changed_layout, &corpus.spec)
            .expect_err("retargeted layout must fail the ownership gate");
        assert!(error.to_string().contains("layout resolves to"));
    }

    #[test]
    fn tiny_timed_run_reports_zero_retention_and_operation_metrics() {
        let corpus = build_corpus(SemanticShape::Tiny).expect("PPTX streaming corpus");
        let measured = run(Case::PptxStreamingCreate, &corpus, 0, 1).expect("PPTX streaming run");
        assert_eq!(measured.elapsed_ns.samples.len(), 1);
        assert_eq!(
            measured.output_sha256.as_deref(),
            Some(corpus.manifest.archive_sha256.as_str())
        );
        let sink = measured.sink.as_ref().expect("sink summary");
        assert_eq!(sink.accepted_bytes, corpus.manifest.archive_bytes as u64);
        assert_eq!(sink.retained_output_bytes, Some(0));
        assert_eq!(sink.retained_authoring_window_bytes, None);
        assert_eq!(
            sink.input_bytes,
            Some(corpus.manifest.uncompressed_payload_bytes as u64)
        );
        let source = measured.source.as_ref().expect("source summary");
        let pptx = source.pptx_slides.as_ref().expect("PPTX summary");
        assert_eq!(pptx.slide_count, corpus.manifest.entry_count);
        assert_eq!(pptx.text_box_count, corpus.manifest.entry_count);
        assert!(pptx.archive_member_set_verified);
        assert!(pptx.semantic_reopen_verified);
        assert!(pptx.deterministic_output_verified);
        assert_eq!(pptx.structural_metadata_fixed_member_count, 37);
        assert_eq!(pptx.structural_metadata_members_per_slide, 2);
        assert_eq!(measured.operation_metrics.as_ref().unwrap().sample_count, 1);
    }
}
