//! Opt-in DOCX story-hyperlink planning evidence.
//!
//! The corpus and the source-backed snapshot are prepared outside the timed
//! interval.  Each retained sample then performs a bounded sequence of
//! `Snapshot::plan_target_urls` calls, with the resulting plans black-boxed so
//! the elapsed vector represents planning rather than package ingress,
//! relationship discovery, publication, or correctness validation.

use super::{Case, CaseResult, Corpus as BaseCorpus, CorpusManifest, SourceSummary};
use litchi_core::{OwnedSource, ReadAt};
use litchi_docx::source_backed;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcPackage, PackURI, PackageWriter, Part, TargetMode};
use serde::Serialize;
use soapberry_zip::office::ArchiveReader;
use std::fmt::Write as _;
use std::sync::Arc;
use std::time::Instant;

const CORPUS_GENERATOR: &str = "litchi-docx-story-hyperlink-plan-v1";
const CORPUS_SHAPE: &str = "49-stories-1152-links";
const PAYLOAD_KIND: &str = "deterministic-story-hyperlinks";
const STORY_COUNT: usize = 48;
const RELATIONSHIPS_PER_STORY: usize = 24;
const PLAN_CALLS_PER_SAMPLE: usize = 8;
const SHARED_TARGET: &str = "https://litchi-perf.invalid/shared-target";
const MISSING_TARGET: &str = "https://litchi-perf.invalid/missing-target";
const SHARED_RELATIONSHIP_ID: &str = "rShared";

const EXPECTED_STORY_COUNT: usize = STORY_COUNT + 1;
const EXPECTED_RELATIONSHIP_COUNT: usize = STORY_COUNT * RELATIONSHIPS_PER_STORY;
const EXPECTED_SELECTED_RELATIONSHIPS: usize = STORY_COUNT;
const EXPECTED_OUTPUT_RELATIONSHIP_COUNT: usize =
    EXPECTED_RELATIONSHIP_COUNT - EXPECTED_SELECTED_RELATIONSHIPS;

/// Content-free correctness and phase evidence for the repeated planning
/// selector.  The untimed gates are intentionally retained beside the plan
/// vector so a result cannot be mistaken for a plan-only latency claim without
/// its package, inventory, and publication identity.
#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
pub(super) struct DocxStoryHyperlinkPlanSummary {
    implementation: &'static str,
    timing_scope: &'static str,
    performance_claim: &'static str,
    plan_calls_per_sample: usize,
    story_count: usize,
    relationship_count: usize,
    selected_target: String,
    selected_relationship_count: usize,
    changed_story_count: usize,
    scanned_story_count: usize,
    source_archive_bytes: u64,
    source_archive_sha256: String,
    output_archive_bytes: u64,
    output_sha256: String,
    plan_ns: Vec<u64>,
    inventory_verified: bool,
    complete_inventory_verified: bool,
    plan_effect_verified: bool,
    duplicate_selector_dedup_verified: bool,
    missing_selector_refusal_verified: bool,
    publication_verified: bool,
    output_inventory_verified: bool,
    output_relationships_verified: bool,
    output_text_verified: bool,
    deterministic_output_verified: bool,
    source_immutability_verified: bool,
}

/// The complete deterministic corpus plus the selector metadata required by
/// the dedicated runner.  Keeping the base `Corpus` private to the harness
/// preserves the existing report shape and default matrix.
#[derive(Debug)]
pub(super) struct Corpus {
    base: BaseCorpus,
    target_url: String,
    expected_story_count: usize,
    expected_relationship_count: usize,
}

#[derive(Debug)]
struct Prepared {
    snapshot: litchi_docx::story_hyperlinks::Snapshot,
    output: Vec<u8>,
    output_sha256: String,
    duplicate_selector_dedup_verified: bool,
    missing_selector_refusal_verified: bool,
    publication_verified: bool,
    output_inventory_verified: bool,
    output_relationships_verified: bool,
    output_text_verified: bool,
    deterministic_output_verified: bool,
    source_immutability_verified: bool,
}

fn reserve_exact<T>(
    values: &mut Vec<T>,
    amount: usize,
    label: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    values
        .try_reserve_exact(amount)
        .map_err(|error| format!("{label} allocation failed: {error}").into())
}

fn reserve_string(
    value: &mut String,
    amount: usize,
    label: &str,
) -> Result<(), Box<dyn std::error::Error>> {
    value
        .try_reserve_exact(amount)
        .map_err(|error| format!("{label} allocation failed: {error}").into())
}

fn generated_string<F>(
    capacity: usize,
    label: &str,
    write_value: F,
) -> Result<String, Box<dyn std::error::Error>>
where
    F: FnOnce(&mut String) -> std::fmt::Result,
{
    let mut value = String::new();
    reserve_string(&mut value, capacity, label)?;
    write_value(&mut value).map_err(|_| format!("{label} formatting failed"))?;
    Ok(value)
}

fn story_file_name(index: usize) -> Result<String, Box<dyn std::error::Error>> {
    generated_string(32, "DOCX story file name", |value| {
        if index % 2 == 0 {
            write!(value, "header{:03}.xml", index / 2 + 1)
        } else {
            write!(value, "footer{:03}.xml", index / 2 + 1)
        }
    })
}

fn story_root(index: usize) -> &'static str {
    if index % 2 == 0 { "hdr" } else { "ftr" }
}

fn story_content_type(index: usize) -> &'static str {
    if index % 2 == 0 {
        ct::WML_HEADER
    } else {
        ct::WML_FOOTER
    }
}

fn story_relationship_type(index: usize) -> &'static str {
    if index % 2 == 0 {
        rt::HEADER
    } else {
        rt::FOOTER
    }
}

fn story_owner_relationship_id(index: usize) -> Result<String, Box<dyn std::error::Error>> {
    generated_string(24, "DOCX story owner relationship ID", |value| {
        write!(value, "rStory{:03}", index + 1)
    })
}

fn hyperlink_relationship_id(index: usize) -> Result<String, Box<dyn std::error::Error>> {
    if index == 0 {
        let mut value = String::new();
        reserve_string(
            &mut value,
            SHARED_RELATIONSHIP_ID.len(),
            "DOCX shared hyperlink relationship ID",
        )?;
        value.push_str(SHARED_RELATIONSHIP_ID);
        Ok(value)
    } else {
        generated_string(24, "DOCX hyperlink relationship ID", |value| {
            write!(value, "rLink{:02}", index)
        })
    }
}

fn hyperlink_target(
    story_index: usize,
    relationship_index: usize,
) -> Result<String, Box<dyn std::error::Error>> {
    if relationship_index == 0 {
        let mut value = String::new();
        reserve_string(
            &mut value,
            SHARED_TARGET.len(),
            "DOCX shared hyperlink target",
        )?;
        value.push_str(SHARED_TARGET);
        Ok(value)
    } else {
        generated_string(96, "DOCX hyperlink target", |value| {
            write!(
                value,
                "https://litchi-perf.invalid/story/{story_index:03}/target/{relationship_index:02}"
            )
        })
    }
}

fn main_xml() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let bytes = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>deterministic main story</w:t></w:r></w:p></w:body></w:document>"#;
    let mut output = Vec::new();
    reserve_exact(&mut output, bytes.len(), "DOCX main story XML")?;
    output.extend_from_slice(bytes);
    Ok(output)
}

fn story_xml(index: usize) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut xml = String::new();
    // This is a fixed upper bound for the generated XML.  The exact payload
    // is smaller, but reserving once makes the generator's allocation path
    // explicit and fallible.
    reserve_string(
        &mut xml,
        512 + RELATIONSHIPS_PER_STORY * 192,
        "DOCX story XML",
    )?;
    write!(
        &mut xml,
        r#"<w:{root} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p>"#,
        root = story_root(index)
    )
    .map_err(|_| "DOCX story XML formatting failed")?;
    for relationship_index in 0..RELATIONSHIPS_PER_STORY {
        let relationship_id = hyperlink_relationship_id(relationship_index)?;
        write!(
            &mut xml,
            r#"<w:hyperlink r:id="{relationship_id}"><w:r><w:t>story-{index:03}-link-{relationship_index:02}</w:t></w:r></w:hyperlink>"#
        )
        .map_err(|_| "DOCX story hyperlink XML formatting failed")?;
    }
    write!(&mut xml, r#"</w:p></w:{root}>"#, root = story_root(index))
        .map_err(|_| "DOCX story XML close formatting failed")?;
    Ok(xml.into_bytes())
}

fn build_archive() -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut package = OpcPackage::new();
    let main_uri = PackURI::new("/word/document.xml")?;
    let mut main = BlobPart::new(main_uri, ct::WML_DOCUMENT_MAIN.to_owned(), main_xml()?);
    for story_index in 0..STORY_COUNT {
        let file_name = story_file_name(story_index)?;
        let relationship_id = story_owner_relationship_id(story_index)?;
        main.rels_mut().try_add_relationship(
            story_relationship_type(story_index).to_owned(),
            file_name.clone(),
            relationship_id,
            TargetMode::Internal,
        )?;
        let part_name = generated_string(48, "DOCX story Part URI", |value| {
            write!(value, "/word/{file_name}")
        })?;
        let mut story = BlobPart::new(
            PackURI::new(part_name)?,
            story_content_type(story_index).to_owned(),
            story_xml(story_index)?,
        );
        for relationship_index in 0..RELATIONSHIPS_PER_STORY {
            story.rels_mut().try_add_relationship(
                rt::HYPERLINK.to_owned(),
                hyperlink_target(story_index, relationship_index)?,
                hyperlink_relationship_id(relationship_index)?,
                TargetMode::External,
            )?;
        }
        package.try_add_part(Box::new(story))?;
    }
    package.try_add_part(Box::new(main))?;
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    Ok(PackageWriter::to_bytes(&package)?)
}

/// Build the deterministic high-cardinality DOCX package used by the case.
pub(super) fn build_corpus() -> Result<Corpus, Box<dyn std::error::Error>> {
    let archive = build_archive()?;
    let opc = OpcPackage::from_bytes(&archive)?;
    let archive_member_count = ArchiveReader::new(&archive)?.file_names().count();
    let uncompressed_payload_bytes = opc.iter_parts().try_fold(0usize, |total, part| {
        total
            .checked_add(part.blob().len())
            .ok_or("DOCX story hyperlink logical byte count overflows usize")
    })?;
    let story_uri = PackURI::new("/word/header001.xml")?;
    let story_bytes = opc.get_part(&story_uri)?.blob().len();
    let mut target_payload = Vec::new();
    reserve_exact(
        &mut target_payload,
        SHARED_TARGET.len(),
        "DOCX story hyperlink target payload",
    )?;
    target_payload.extend_from_slice(SHARED_TARGET.as_bytes());
    let archive_sha256 = super::sha256_hex(&archive);
    let manifest = CorpusManifest {
        name: "docx-story-hyperlink-plan".to_owned(),
        generator: CORPUS_GENERATOR,
        package_format: "DOCX/OPC/ZIP",
        shape: CORPUS_SHAPE,
        payload_kind: PAYLOAD_KIND,
        compression: "deflate",
        entry_count: EXPECTED_RELATIONSHIP_COUNT,
        archive_member_count,
        entry_bytes: story_bytes,
        uncompressed_payload_bytes,
        archive_bytes: archive.len(),
        archive_sha256,
        target_entry: "shared-external-hyperlink-relationship".to_owned(),
        target_payload_bytes: target_payload.len(),
        target_payload_sha256: super::sha256_hex(&target_payload),
        rtf_variant: None,
        xlsx: None,
    };
    let mut target_url = String::new();
    reserve_string(
        &mut target_url,
        SHARED_TARGET.len(),
        "DOCX story hyperlink target selector",
    )?;
    target_url.push_str(SHARED_TARGET);
    Ok(Corpus {
        base: BaseCorpus {
            manifest,
            archive,
            target_name: "shared-external-hyperlink-relationship".to_owned(),
            target_payload,
            xlsx: None,
        },
        target_url,
        expected_story_count: EXPECTED_STORY_COUNT,
        expected_relationship_count: EXPECTED_RELATIONSHIP_COUNT,
    })
}

fn owned_source(archive: &[u8]) -> Result<Arc<OwnedSource>, Box<dyn std::error::Error>> {
    let mut owned = Vec::new();
    reserve_exact(&mut owned, archive.len(), "DOCX story hyperlink source")?;
    owned.extend_from_slice(archive);
    Ok(Arc::new(OwnedSource::new(owned)))
}

fn source_package(
    source: Arc<OwnedSource>,
) -> Result<source_backed::Package, Box<dyn std::error::Error>> {
    let source: Arc<dyn ReadAt> = source;
    Ok(source_backed::Package::from_read_at(source)?)
}

fn verify_plan_effect(plan: &litchi_docx::story_hyperlinks::Plan, corpus: &Corpus) -> bool {
    let effect = plan.effect_report();
    effect.selected_targets() == 1
        && effect.scanned_stories() == corpus.expected_story_count
        && effect.removed_relationships() == EXPECTED_SELECTED_RELATIONSHIPS
        && effect.unwrapped_hyperlinks() == EXPECTED_SELECTED_RELATIONSHIPS
}

fn verify_output_parts(output: &[u8]) -> Result<bool, Box<dyn std::error::Error>> {
    let package = OpcPackage::from_bytes(output)?;
    for story_index in 0..STORY_COUNT {
        let file_name = story_file_name(story_index)?;
        let part_name = generated_string(48, "DOCX output story URI", |value| {
            write!(value, "/word/{file_name}")
        })?;
        let part = package.get_part(&PackURI::new(part_name)?)?;
        if part.rels().iter().any(|relationship| {
            relationship.r_id() == SHARED_RELATIONSHIP_ID
                || relationship.target_ref() == SHARED_TARGET
        }) {
            return Ok(false);
        }
        if part
            .blob()
            .windows(SHARED_RELATIONSHIP_ID.len())
            .any(|window| window == SHARED_RELATIONSHIP_ID.as_bytes())
        {
            return Ok(false);
        }
        let expected_remaining = RELATIONSHIPS_PER_STORY - 1;
        if part.rels().iter().count() != expected_remaining {
            return Ok(false);
        }
        let label = format!("story-{story_index:03}-link-00");
        if !part
            .blob()
            .windows(label.len())
            .any(|window| window == label.as_bytes())
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn prepare(corpus: &Corpus) -> Result<Prepared, Box<dyn std::error::Error>> {
    let source = owned_source(&corpus.base.archive)?;
    let package = source_package(Arc::clone(&source))?;
    let snapshot = package.story_hyperlinks_only_snapshot()?;
    let inventory_verified = snapshot.inventory().story_count() == corpus.expected_story_count
        && snapshot.inventory().relationship_count() == corpus.expected_relationship_count;
    let complete_inventory_verified = snapshot.diagnostics().is_empty();
    if !inventory_verified || !complete_inventory_verified {
        return Err("DOCX story hyperlink inventory differs from deterministic corpus".into());
    }

    let target = corpus.target_url.as_str();
    let plan = snapshot.plan_target_urls(&[target])?;
    let plan_effect_verified = verify_plan_effect(&plan, corpus);
    if !plan_effect_verified {
        return Err("DOCX story hyperlink plan effect differs from deterministic corpus".into());
    }
    let duplicate_plan = snapshot.plan_target_urls(&[target, target])?;
    let duplicate_selector_dedup_verified = duplicate_plan.effect_report().selected_targets() == 1
        && verify_plan_effect(&duplicate_plan, corpus);
    let missing_selector_refusal_verified = snapshot.plan_target_urls(&[MISSING_TARGET]).is_err();
    if !duplicate_selector_dedup_verified || !missing_selector_refusal_verified {
        return Err("DOCX story hyperlink selector guards are not verified".into());
    }

    let commit = plan.apply()?;
    let mut output = Vec::new();
    reserve_exact(
        &mut output,
        corpus.base.archive.len(),
        "DOCX story hyperlink output",
    )?;
    package.publish_story_hyperlink_redaction_to_stream(&mut output, &commit)?;
    let output_sha256 = super::sha256_hex(&output);
    let output_parts_verified = verify_output_parts(&output)?;
    let output_package = source_package(owned_source(&output)?)?;
    let output_snapshot = output_package.story_hyperlinks_only_snapshot()?;
    let output_inventory_verified = output_snapshot.inventory().story_count()
        == corpus.expected_story_count
        && output_snapshot.diagnostics().is_empty();
    let output_relationships_verified = output_snapshot.inventory().relationship_count()
        == EXPECTED_OUTPUT_RELATIONSHIP_COUNT
        && output_snapshot
            .relationships()
            .iter()
            .all(|relationship| relationship.target_url() != SHARED_TARGET);
    let output_text_verified = output_parts_verified;
    let mut second_output = Vec::new();
    reserve_exact(
        &mut second_output,
        corpus.base.archive.len(),
        "DOCX repeated story hyperlink output",
    )?;
    let second_package = source_package(owned_source(&corpus.base.archive)?)?;
    let second_snapshot = second_package.story_hyperlinks_only_snapshot()?;
    let second_plan = second_snapshot.plan_target_urls(&[target])?;
    let second_commit = second_plan.apply()?;
    second_package
        .publish_story_hyperlink_redaction_to_stream(&mut second_output, &second_commit)?;
    let deterministic_output_verified = second_output == output;
    let publication_verified = output_inventory_verified
        && output_relationships_verified
        && output_text_verified
        && deterministic_output_verified;
    if !publication_verified {
        return Err("DOCX story hyperlink publication gates failed".into());
    }
    let mut retained_source = Vec::new();
    reserve_exact(
        &mut retained_source,
        corpus.base.archive.len(),
        "DOCX retained source verification",
    )?;
    retained_source.resize(corpus.base.archive.len(), 0);
    source.read_exact_at(0, &mut retained_source)?;
    let source_immutability_verified = source.len()? == u64::try_from(retained_source.len())?
        && retained_source == corpus.base.archive
        && super::sha256_hex(&retained_source) == corpus.base.manifest.archive_sha256;
    if !source_immutability_verified {
        return Err("DOCX story hyperlink source bytes changed during preparation".into());
    }
    Ok(Prepared {
        snapshot,
        output,
        output_sha256,
        duplicate_selector_dedup_verified,
        missing_selector_refusal_verified,
        publication_verified,
        output_inventory_verified,
        output_relationships_verified,
        output_text_verified,
        deterministic_output_verified,
        source_immutability_verified,
    })
}

/// Measure repeated planning on one pre-opened immutable source-backed
/// snapshot.  Corpus construction, snapshot capture, selector guards,
/// publication, output reopening, and all semantic checks are outside timing.
pub(super) fn run(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn std::error::Error>> {
    if case != Case::DocxStoryHyperlinkPlan {
        return Err("DOCX story hyperlink runner received an unrelated case".into());
    }
    let prepared = prepare(corpus)?;
    let mut elapsed = Vec::new();
    reserve_exact(&mut elapsed, samples, "DOCX story hyperlink timing samples")?;
    let total_iterations = warmup_iterations
        .checked_add(samples)
        .ok_or("DOCX story hyperlink iteration count overflows usize")?;
    for iteration in 0..total_iterations {
        let started = Instant::now();
        for _ in 0..PLAN_CALLS_PER_SAMPLE {
            let plan = prepared
                .snapshot
                .plan_target_urls(&[corpus.target_url.as_str()])?;
            std::hint::black_box(plan);
        }
        let duration = started.elapsed();
        if iteration >= warmup_iterations {
            elapsed.push(super::elapsed_ns(duration)?);
        }
    }
    let source_archive_bytes = u64::try_from(corpus.base.archive.len())?;
    let output_archive_bytes = u64::try_from(prepared.output.len())?;
    let source_archive_sha256 = corpus.base.manifest.archive_sha256.clone();
    let summary = DocxStoryHyperlinkPlanSummary {
        implementation: "litchi-docx::source_backed::Package + story_hyperlinks::Snapshot",
        timing_scope: "eight repeated Snapshot::plan_target_urls calls on a prepared immutable snapshot",
        performance_claim: "planning evidence only; no end-to-end, I/O, allocation, RSS, or speedup claim",
        plan_calls_per_sample: PLAN_CALLS_PER_SAMPLE,
        story_count: corpus.expected_story_count,
        relationship_count: corpus.expected_relationship_count,
        selected_target: corpus.target_url.clone(),
        selected_relationship_count: EXPECTED_SELECTED_RELATIONSHIPS,
        changed_story_count: EXPECTED_SELECTED_RELATIONSHIPS,
        scanned_story_count: corpus.expected_story_count,
        source_archive_bytes,
        source_archive_sha256,
        output_archive_bytes,
        output_sha256: prepared.output_sha256.clone(),
        plan_ns: elapsed.clone(),
        inventory_verified: true,
        complete_inventory_verified: true,
        plan_effect_verified: true,
        duplicate_selector_dedup_verified: prepared.duplicate_selector_dedup_verified,
        missing_selector_refusal_verified: prepared.missing_selector_refusal_verified,
        publication_verified: prepared.publication_verified,
        output_inventory_verified: prepared.output_inventory_verified,
        output_relationships_verified: prepared.output_relationships_verified,
        output_text_verified: prepared.output_text_verified,
        deterministic_output_verified: prepared.deterministic_output_verified,
        source_immutability_verified: prepared.source_immutability_verified,
    };
    let mut source = SourceSummary::default();
    source.docx_story_hyperlinks = Some(summary);
    Ok(CaseResult {
        case: case.name(),
        cache_state: None,
        corpus: corpus.base.manifest.clone(),
        elapsed_ns: super::statistics(elapsed),
        sink: None,
        source: Some(source),
        execution: None,
        output_sha256: Some(prepared.output_sha256),
        operation_metrics: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn corpus_is_bounded_deterministic_and_out_of_default() {
        let first = build_corpus().expect("first corpus");
        let second = build_corpus().expect("second corpus");
        assert_eq!(first.base.archive, second.base.archive);
        assert_eq!(
            first.base.manifest.archive_sha256,
            second.base.manifest.archive_sha256
        );
        assert_eq!(first.expected_story_count, EXPECTED_STORY_COUNT);
        assert_eq!(
            first.expected_relationship_count,
            EXPECTED_RELATIONSHIP_COUNT
        );
        assert!(!Case::DEFAULT.contains(&Case::DocxStoryHyperlinkPlan));
        assert_eq!(
            super::super::parse_case("docx_story_hyperlink_plan"),
            Some(Case::DocxStoryHyperlinkPlan)
        );
    }

    #[test]
    fn runner_keeps_plan_timing_separate_from_correctness_gates() {
        let corpus = build_corpus().expect("story hyperlink corpus");
        let result = run(Case::DocxStoryHyperlinkPlan, &corpus, 0, 1).expect("run");
        let source = result.source.expect("source evidence");
        let summary = source
            .docx_story_hyperlinks
            .expect("story hyperlink evidence");
        assert_eq!(summary.plan_calls_per_sample, PLAN_CALLS_PER_SAMPLE);
        assert_eq!(summary.plan_ns.len(), 1);
        assert!(summary.plan_ns[0] > 0);
        assert!(summary.inventory_verified);
        assert!(summary.complete_inventory_verified);
        assert!(summary.plan_effect_verified);
        assert!(summary.duplicate_selector_dedup_verified);
        assert!(summary.missing_selector_refusal_verified);
        assert!(summary.publication_verified);
        assert!(summary.output_inventory_verified);
        assert!(summary.output_relationships_verified);
        assert!(summary.output_text_verified);
        assert!(summary.deterministic_output_verified);
        assert!(summary.source_immutability_verified);
    }
}
