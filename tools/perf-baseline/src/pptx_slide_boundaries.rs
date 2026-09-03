//! Opt-in PPTX whole-slide boundary CRUD evidence.
//!
//! This module deliberately keeps the corpus small and the claim narrow.  It
//! exercises the dependency-free first/middle/last removal boundary and the
//! first/last move boundary, while retaining the phase vectors and the exact
//! package/sink checks needed to make the result useful as correctness and
//! publication evidence.  It is not part of the default performance matrix.

use super::{
    Case, CaseResult, CorpusManifest, CountingSink, PrefixFailSink, SourceSummary,
    deterministic_sink_summary, elapsed_ns, iteration_count, sha256_hex, statistics,
};
use litchi_pptx::opened::{Limits, Patch};
use litchi_pptx::{Error, Package, SlideRemovalPatch, SlideRemovalRefusal};
use litchi_opc::OpcPackage;
use serde::Serialize;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};
use std::error::Error as StdError;
use std::time::Instant;

const CORPUS_GENERATOR: &str = "litchi-pptx-slide-boundary-save-v1";
const CORPUS_SHAPE: &str = "four-dependency-free-plain-slides";
const SLIDE_COUNT: usize = 4;
const REMOVE_POSITIONS: [usize; 3] = [0, 1, 3];
const MOVE_PAIRS: [(usize, usize); 2] = [(0, 3), (3, 0)];
const REPRESENTATIVE_REMOVE_POSITION: usize = 1;
const REPRESENTATIVE_MOVE: (usize, usize) = (0, 3);
const DEPENDENCY_SLIDE_POSITION: usize = 1;
const EXPECTED_SOURCE_ARCHIVE_BYTES: usize = 32_396;
const EXPECTED_SOURCE_ARCHIVE_SHA256: &str =
    "685a1805ad291e8f9852d3ccd584320f20847bd0ac8fdf29857f96efe1109477";
const EXPECTED_SOURCE_ARCHIVE_MEMBER_COUNT: usize = 45;
const MAX_WRITE: u64 = 64 * 1024;

/// Per-sample phase and untimed gate evidence for the boundary selectors.
#[derive(Clone, Debug, Serialize, PartialEq, Eq)]
pub(super) struct PptxSlideBoundarySummary {
    implementation: &'static str,
    operation: &'static str,
    timing_scope: &'static str,
    performance_claim: &'static str,
    source_archive_bytes: u64,
    source_archive_sha256: String,
    output_archive_bytes: u64,
    expected_output_sha256: String,
    source_slide_count: usize,
    output_slide_count: usize,
    remove_positions: Vec<usize>,
    move_pairs: Vec<(usize, usize)>,
    plan_ns: Vec<u64>,
    commit_ns: Vec<u64>,
    publication_ns: Vec<u64>,
    reopen_ns: Vec<u64>,
    output_sha256: Vec<String>,
    boundary_semantics_verified: bool,
    no_op_verified: bool,
    untouched_raw_members_verified: bool,
    deterministic_output_verified: bool,
    durable_forward_verified: bool,
    durable_inverse_verified: bool,
    stale_refusal_verified: bool,
    foreign_refusal_verified: bool,
    partial_sink_verified: bool,
    zero_sink_verified: bool,
    one_slide_refusal_verified: bool,
    unknown_member_refusal_verified: bool,
    dependency_refusal_verified: bool,
    markup_compatibility_refusal_verified: bool,
    signed_refusal_verified: bool,
    limits_verified: bool,
    source_immutability_verified: bool,
}

#[derive(Debug)]
pub(super) struct Corpus {
    manifest: CorpusManifest,
    archive: Vec<u8>,
    foreign_archive: Vec<u8>,
    dependency_archive: Vec<u8>,
    slide_names: Vec<String>,
    remove_outputs: Vec<Vec<u8>>,
    move_outputs: Vec<Vec<u8>>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum Operation {
    Remove,
    Move,
}

impl Operation {
    const fn name(self) -> &'static str {
        match self {
            Self::Remove => "remove",
            Self::Move => "move",
        }
    }

}

fn reserve_exact<T>(values: &mut Vec<T>, amount: usize, label: &str) -> Result<(), Box<dyn StdError>> {
    values
        .try_reserve_exact(amount)
        .map_err(|error| format!("{label} allocation failed: {error}").into())
}

fn build_plain_archive(slide_count: usize, prefix: &str) -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut package = Package::new()?;
    let presentation = package.presentation_mut()?;
    for index in 0..slide_count {
        let slide = presentation.add_slide()?;
        slide.set_title(&format!("{prefix}-title-{index}"));
        slide.add_text_box(
            &format!("{prefix}-body-{index}"),
            36,
            36,
            540,
            72,
        );
    }
    Ok(package.to_bytes()?)
}

fn package_slide_names(bytes: &[u8]) -> Result<Vec<String>, Box<dyn StdError>> {
    let package = Package::from_vec(bytes.to_vec())?;
    let snapshot = package.opened_presentation()?;
    let mut names = Vec::new();
    reserve_exact(&mut names, snapshot.slides().len(), "PPTX boundary slide names")?;
    names.extend(snapshot.slides().iter().map(|slide| slide.name().to_owned()));
    Ok(names)
}

fn remove_bytes(bytes: &[u8], position: usize) -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut package = Package::from_vec(bytes.to_vec())?;
    let snapshot = package.opened_presentation()?;
    let plan = snapshot.plan_slide_removal(position)?;
    package.apply_slide_removal_plan(&plan)?;
    Ok(package.to_bytes()?)
}

fn move_bytes(
    bytes: &[u8],
    from: usize,
    to: usize,
) -> Result<Vec<u8>, Box<dyn StdError>> {
    let mut package = Package::from_vec(bytes.to_vec())?;
    let snapshot = package.opened_presentation()?;
    let mut transaction = snapshot.edit();
    let changed = transaction.move_slide(from, to)?;
    let commit = transaction.commit()?;
    if changed {
        package.apply_opened_presentation_commit(commit)?;
    }
    Ok(package.to_bytes()?)
}

fn rewrite_member<F>(bytes: &[u8], member_name: &str, rewrite: F) -> Result<Vec<u8>, Box<dyn StdError>>
where
    F: FnOnce(&[u8]) -> Result<Vec<u8>, Box<dyn StdError>>,
{
    let archive = ArchiveReader::new(bytes)?;
    let mut writer = StreamingArchiveWriter::new();
    let mut rewrite = Some(rewrite);
    for name in archive.file_names() {
        let member = archive.read(name)?;
        let member = if name == member_name {
            rewrite
                .take()
                .ok_or("PPTX boundary rewrite member was encountered twice")?(&member)?
        } else {
            member
        };
        writer.write_stored(name, &member)?;
    }
    Ok(writer.finish_to_bytes()?)
}

fn add_unknown_member(bytes: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
    let archive = ArchiveReader::new(bytes)?;
    let mut writer = StreamingArchiveWriter::new();
    for name in archive.file_names() {
        writer.write_stored(name, &archive.read(name)?)?;
    }
    writer.write_stored("boundary-opaque.bin", b"unmodeled physical payload")?;
    Ok(writer.finish_to_bytes()?)
}

fn add_mce_marker(bytes: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
    rewrite_member(bytes, "ppt/presentation.xml", |member| {
        let xml = std::str::from_utf8(member)?;
        let updated = xml.replacen(
            "<p:presentation ",
            "<p:presentation xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" ",
            1,
        );
        Ok(updated.into_bytes())
    })
}

fn add_signature_marker(bytes: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
    rewrite_member(bytes, "_rels/.rels", |member| {
        let xml = std::str::from_utf8(member)?;
        let relationship = "<Relationship Id=\"rIdBoundarySignature\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"_xmlsignatures/origin.sigs\"/>";
        let updated = xml.replacen("</Relationships>", &format!("{relationship}</Relationships>"), 1);
        Ok(updated.into_bytes())
    })
}

fn add_external_dependency(bytes: &[u8]) -> Result<Vec<u8>, Box<dyn StdError>> {
    rewrite_member(bytes, "ppt/slides/_rels/slide2.xml.rels", |member| {
        let xml = std::str::from_utf8(member)?;
        let relationship = "<Relationship Id=\"rIdBoundaryExternal\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink\" Target=\"https://litchi-perf.invalid/boundary\" TargetMode=\"External\"/>";
        let updated = xml.replacen("</Relationships>", &format!("{relationship}</Relationships>"), 1);
        Ok(updated.into_bytes())
    })
}

fn expected_manifest(archive: &[u8]) -> Result<CorpusManifest, Box<dyn StdError>> {
    let opc = OpcPackage::from_bytes(archive)?;
    let uncompressed_payload_bytes = opc.iter_parts().try_fold(0usize, |total, part| {
        total
            .checked_add(part.blob().len())
            .ok_or("PPTX boundary logical payload count overflows usize")
    })?;
    let entry_bytes = opc
        .iter_parts()
        .find(|part| part.partname().as_str() == "/ppt/slides/slide2.xml")
        .map_or(0, |part| part.blob().len());
    Ok(CorpusManifest {
        name: "pptx-slide-boundary-save".to_owned(),
        generator: CORPUS_GENERATOR,
        package_format: "PPTX/OPC/ZIP",
        shape: CORPUS_SHAPE,
        payload_kind: "dependency-free-plain-slide-text",
        compression: "deflate",
        entry_count: opc.part_count(),
        archive_member_count: ArchiveReader::new(archive)?.file_names().count(),
        entry_bytes,
        uncompressed_payload_bytes,
        archive_bytes: archive.len(),
        archive_sha256: sha256_hex(archive),
        target_entry: "slide-boundary-position".to_owned(),
        target_payload_bytes: entry_bytes,
        target_payload_sha256: sha256_hex(
            opc.iter_parts()
                .find(|part| part.partname().as_str() == "/ppt/slides/slide2.xml")
                .map_or(&[][..], |part| part.blob()),
        ),
        rtf_variant: None,
        xlsx: None,
    })
}

/// Build the fixed four-slide corpus and all untimed expected boundary outputs.
pub(super) fn build_corpus() -> Result<Corpus, Box<dyn StdError>> {
    let archive = build_plain_archive(SLIDE_COUNT, "boundary")?;
    let manifest = expected_manifest(&archive)?;
    if archive.len() != EXPECTED_SOURCE_ARCHIVE_BYTES
        || manifest.archive_sha256 != EXPECTED_SOURCE_ARCHIVE_SHA256
        || manifest.archive_member_count != EXPECTED_SOURCE_ARCHIVE_MEMBER_COUNT
    {
        return Err(format!(
            "PPTX boundary corpus identity changed: bytes={}, sha256={}, members={}",
            archive.len(), manifest.archive_sha256, manifest.archive_member_count
        )
        .into());
    }
    let foreign_archive = move_bytes(&archive, 1, 2)?;
    let dependency_archive = add_external_dependency(&archive)?;
    let slide_names = package_slide_names(&archive)?;
    if slide_names.len() != SLIDE_COUNT {
        return Err("PPTX boundary corpus did not produce four slides".into());
    }
    let mut remove_outputs = Vec::new();
    reserve_exact(
        &mut remove_outputs,
        REMOVE_POSITIONS.len(),
        "PPTX boundary removal outputs",
    )?;
    for position in REMOVE_POSITIONS {
        remove_outputs.push(remove_bytes(&archive, position)?);
    }
    let mut move_outputs = Vec::new();
    reserve_exact(
        &mut move_outputs,
        MOVE_PAIRS.len(),
        "PPTX boundary move outputs",
    )?;
    for (from, to) in MOVE_PAIRS {
        move_outputs.push(move_bytes(&archive, from, to)?);
    }
    let corpus = Corpus {
        manifest,
        archive,
        foreign_archive,
        dependency_archive,
        slide_names,
        remove_outputs,
        move_outputs,
    };
    verify_untimed_gates(&corpus)?;
    Ok(corpus)
}

fn changed_member_allowed(name: &str, operation: Operation, selected_slide: Option<&str>) -> bool {
    if (operation == Operation::Remove && name == "[Content_Types].xml")
        || name == "ppt/presentation.xml"
        || name == "ppt/_rels/presentation.xml.rels"
    {
        return true;
    }
    match (operation, selected_slide) {
        (Operation::Remove, Some(slide)) => {
            name == slide
                || name == slide.replace("ppt/slides/", "ppt/slides/_rels/").replace(
                    ".xml",
                    ".xml.rels",
                )
        },
        (Operation::Move, _) => false,
        _ => false,
    }
}

fn verify_untouched_raw_members(
    source: &[u8],
    candidate: &[u8],
    operation: Operation,
    selected_slide: Option<&str>,
) -> Result<bool, Box<dyn StdError>> {
    let source_members = super::raw_zip_members(source)?;
    let candidate_members = super::raw_zip_members(candidate)?;
    for (name, source_member) in &source_members {
        if changed_member_allowed(name, operation, selected_slide) {
            continue;
        }
        let Some(candidate_member) = candidate_members.get(name) else {
            return Ok(false);
        };
        if candidate_member != source_member {
            return Ok(false);
        }
    }
    for name in candidate_members.keys() {
        if !source_members.contains_key(name)
            && !changed_member_allowed(name, operation, selected_slide)
        {
            return Ok(false);
        }
    }
    Ok(true)
}

fn output_names(bytes: &[u8]) -> Result<Vec<String>, Box<dyn StdError>> {
    package_slide_names(bytes)
}

fn verify_semantic_output(
    bytes: &[u8],
    operation: Operation,
    position: usize,
    corpus: &Corpus,
) -> Result<bool, Box<dyn StdError>> {
    let actual = output_names(bytes)?;
    let expected = match operation {
        Operation::Remove => corpus
            .slide_names
            .iter()
            .enumerate()
            .filter(|(index, _)| *index != position)
            .map(|(_, name)| name.clone())
            .collect::<Vec<_>>(),
        Operation::Move => {
            let (from, to) = MOVE_PAIRS
                .get(position)
                .copied()
                .ok_or("PPTX boundary move output index is outside the fixed pair set")?;
            let mut names = corpus.slide_names.clone();
            let moved = names.remove(from);
            names.insert(to, moved);
            names
        },
    };
    Ok(actual == expected)
}

fn apply_boundary_to_package(
    package: &mut Package,
    operation: Operation,
    position: usize,
) -> Result<(), Box<dyn StdError>> {
    match operation {
        Operation::Remove => {
            let snapshot = package.opened_presentation()?;
            let plan = snapshot.plan_slide_removal(position)?;
            package.apply_slide_removal_plan(&plan)?;
        },
        Operation::Move => {
            let (from, to) = MOVE_PAIRS
                .get(position)
                .copied()
                .ok_or("PPTX boundary move operation index is outside the fixed pair set")?;
            let snapshot = package.opened_presentation()?;
            let mut transaction = snapshot.edit();
            transaction.move_slide(from, to)?;
            package.apply_opened_presentation_commit(transaction.commit()?)?;
        },
    }
    Ok(())
}

fn verify_partial_zero_sinks(
    source: &[u8],
    operation: Operation,
    position: usize,
    expected_output: &[u8],
) -> Result<(bool, bool), Box<dyn StdError>> {
    let mut package = Package::from_vec(source.to_vec())?;
    apply_boundary_to_package(&mut package, operation, position)?;
    let fail_after = u64::try_from(expected_output.len() / 2)
        .map_err(|_error| "PPTX boundary partial sink length exceeds u64")?
        .max(1);
    let mut partial = PrefixFailSink {
        accepted: 0,
        fail_after,
    };
    let partial_result = package.opc()?.to_stream(&mut partial);
    let partial_verified = partial_result.is_err()
        && partial.accepted > 0
        && partial.accepted < u64::try_from(expected_output.len())?;
    let mut zero = PrefixFailSink {
        accepted: 0,
        fail_after: 0,
    };
    let zero_result = package.opc()?.to_stream(&mut zero);
    let zero_verified = zero_result.is_err() && zero.accepted == 0;
    Ok((partial_verified, zero_verified))
}

fn mutate_order(package: &mut Package) -> Result<(), Box<dyn StdError>> {
    let snapshot = package.opened_presentation()?;
    let mut transaction = snapshot.edit();
    transaction.move_slide(0, 1)?;
    package.apply_opened_presentation_commit(transaction.commit()?)?;
    Ok(())
}

fn verify_remove_durable(corpus: &Corpus) -> Result<(bool, bool, bool, bool), Box<dyn StdError>> {
    let source_package = Package::from_vec(corpus.archive.clone())?;
    let plan = source_package
        .opened_presentation()?
        .plan_slide_removal(REPRESENTATIVE_REMOVE_POSITION)?;
    let forward = SlideRemovalPatch::from_bytes(&plan.patch().to_bytes()?)?;
    let inverse = SlideRemovalPatch::from_bytes(&plan.patch().inverse().to_bytes()?)?;
    let mut applied = Package::from_vec(corpus.archive.clone())?;
    applied.apply_slide_removal_patch(&forward)?;
    let applied_bytes = applied.to_bytes()?;
    let forward_verified = output_names(&applied_bytes)?.len() == SLIDE_COUNT - 1;
    let mut restored = Package::from_vec(applied_bytes)?;
    restored.apply_slide_removal_patch(&inverse)?;
    let restored_bytes = restored.to_bytes()?;
    let inverse_verified = package_slide_names(&restored_bytes)? == corpus.slide_names
        && super::raw_zip_members(&restored_bytes)? == super::raw_zip_members(&corpus.archive)?;

    let mut stale = Package::from_vec(corpus.archive.clone())?;
    mutate_order(&mut stale)?;
    let stale_verified = stale.apply_slide_removal_patch(&forward).is_err();
    let mut foreign = Package::from_vec(corpus.foreign_archive.clone())?;
    let foreign_verified = foreign.apply_slide_removal_patch(&forward).is_err();
    Ok((
        forward_verified,
        inverse_verified,
        stale_verified,
        foreign_verified,
    ))
}

fn verify_move_durable(corpus: &Corpus) -> Result<(bool, bool, bool, bool), Box<dyn StdError>> {
    let package = Package::from_vec(corpus.archive.clone())?;
    let snapshot = package.opened_presentation()?;
    let mut transaction = snapshot.edit();
    transaction.move_slide(REPRESENTATIVE_MOVE.0, REPRESENTATIVE_MOVE.1)?;
    let commit = transaction.commit()?;
    let forward = Patch::from_bytes(&commit.patch().to_bytes()?)?;
    let inverse = Patch::from_bytes(&forward.inverse().to_bytes()?)?;
    let mut applied = Package::from_vec(corpus.archive.clone())?;
    applied.apply_opened_presentation_patch(&forward)?;
    let applied_bytes = applied.to_bytes()?;
    let forward_verified = verify_semantic_output(&applied_bytes, Operation::Move, 0, corpus)?
        && verify_untouched_raw_members(&corpus.archive, &applied_bytes, Operation::Move, None)?;
    let mut restored = Package::from_vec(applied_bytes)?;
    restored.apply_opened_presentation_patch(&inverse)?;
    let restored_bytes = restored.to_bytes()?;
    let inverse_verified = restored_bytes == corpus.archive;

    let mut stale = Package::from_vec(corpus.archive.clone())?;
    mutate_order(&mut stale)?;
    let stale_verified = stale.apply_opened_presentation_patch(&forward).is_err();
    let mut foreign = Package::from_vec(corpus.foreign_archive.clone())?;
    let foreign_verified = foreign.apply_opened_presentation_patch(&forward).is_err();
    Ok((
        forward_verified,
        inverse_verified,
        stale_verified,
        foreign_verified,
    ))
}

fn verify_dependency_refusal(corpus: &Corpus) -> Result<bool, Box<dyn StdError>> {
    let source = corpus.dependency_archive.as_slice();
    let source_before = source.to_vec();
    let mut package = Package::from_vec(source_before.clone())?;
    let snapshot = package.opened_presentation()?;
    let source_revision = snapshot.revision();
    let before_names = snapshot
        .slides()
        .iter()
        .map(|slide| slide.name().to_owned())
        .collect::<Vec<_>>();
    let mut transaction = snapshot.edit();
    let refusal = transaction.remove_slide(DEPENDENCY_SLIDE_POSITION);
    let typed_refusal = matches!(
        refusal,
        Err(Error::SlideRemovalPlan {
            kind: SlideRemovalRefusal::UnsupportedRelationship,
            ..
        })
    );
    let staging_names = transaction
        .slides()
        .iter()
        .map(|slide| slide.name().to_owned())
        .collect::<Vec<_>>();
    let staging_unchanged = !transaction.is_changed()
        && transaction.source().revision() == source_revision;
    let rolled_back = transaction.rollback();
    let rollback_names = rolled_back
        .slides()
        .iter()
        .map(|slide| slide.name().to_owned())
        .collect::<Vec<_>>();
    let output = package.to_bytes()?;
    Ok(typed_refusal
        && staging_names == before_names
        && staging_unchanged
        && rollback_names == before_names
        && source == source_before.as_slice()
        && output == source_before)
}

fn verify_refusal_gates(
    corpus: &Corpus,
) -> Result<(bool, bool, bool, bool, bool, bool), Box<dyn StdError>> {
    let one_slide = Package::from_vec(build_plain_archive(1, "sole")?)?;
    let one_slide_verified = matches!(
        one_slide
            .opened_presentation()?
            .plan_slide_removal(0),
        Err(Error::SlideRemovalPlan {
            kind: SlideRemovalRefusal::FinalSlide,
            ..
        })
    );

    let unknown = Package::from_vec(add_unknown_member(&corpus.archive)?)?;
    let unknown_verified = matches!(
        unknown
            .opened_presentation()?
            .plan_slide_removal(0),
        Err(Error::SlideRemovalPlan {
            kind: SlideRemovalRefusal::UnknownPhysicalMember,
            ..
        })
    );

    let dependency_verified = verify_dependency_refusal(corpus)?;

    let mce = Package::from_vec(add_mce_marker(&corpus.archive)?)?;
    let mce_verified = matches!(
        mce.opened_presentation()?.plan_slide_removal(0),
        Err(Error::SlideRemovalPlan {
            kind: SlideRemovalRefusal::MarkupCompatibility,
            ..
        })
    );

    let signed = Package::from_vec(add_signature_marker(&corpus.archive)?)?;
    let signed_verified = matches!(
        signed
            .opened_presentation()?
            .plan_slide_removal(0),
        Err(Error::SlideRemovalPlan {
            kind: SlideRemovalRefusal::SignedPackage,
            ..
        })
    );

    let limits = Limits::new(64, 1, 1_024, 1, 1)
        .ok_or("PPTX boundary test limits must be nonzero")?;
    let remove_limited = Package::from_vec(corpus.archive.clone())?
        .opened_presentation_with_limits(limits)?
        .plan_slide_removal(REPRESENTATIVE_REMOVE_POSITION)
        .is_err();
    let move_limited = {
        let package = Package::from_vec(corpus.archive.clone())?;
        let snapshot = package.opened_presentation_with_limits(limits)?;
        let mut transaction = snapshot.edit();
        transaction.move_slide(REPRESENTATIVE_MOVE.0, REPRESENTATIVE_MOVE.1)?;
        transaction.commit().is_err()
    };
    Ok((
        one_slide_verified,
        unknown_verified,
        dependency_verified,
        mce_verified,
        signed_verified,
        remove_limited && move_limited,
    ))
}

fn verify_boundary_semantics(corpus: &Corpus) -> Result<(bool, bool), Box<dyn StdError>> {
    let mut remove_verified = true;
    for (output, position) in corpus.remove_outputs.iter().zip(REMOVE_POSITIONS) {
        let package = Package::from_vec(output.clone())?;
        let snapshot = package.opened_presentation()?;
        let expected = corpus
            .slide_names
            .iter()
            .enumerate()
            .filter(|(index, _)| *index != position)
            .map(|(_, name)| name.clone())
            .collect::<Vec<_>>();
        remove_verified &= snapshot.slides().iter().map(|slide| slide.name()).eq(expected.iter().map(String::as_str));
        let selected = corpus
            .slide_names
            .get(position)
            .ok_or("PPTX removal boundary selected name is missing")?;
        let selected_slide = format!("ppt/slides/slide{}.xml", position + 1);
        remove_verified &= verify_untouched_raw_members(
            &corpus.archive,
            output,
            Operation::Remove,
            Some(&selected_slide),
        )?;
        remove_verified &= !selected.is_empty();
    }

    let mut move_verified = true;
    for (index, output) in corpus.move_outputs.iter().enumerate() {
        move_verified &= verify_semantic_output(output, Operation::Move, index, corpus)?;
        move_verified &= verify_untouched_raw_members(&corpus.archive, output, Operation::Move, None)?;
    }
    let noop = move_bytes(&corpus.archive, 0, 0)?;
    let noop_verified = noop == corpus.archive;
    Ok((remove_verified && corpus.remove_outputs.len() == REMOVE_POSITIONS.len(), move_verified && noop_verified))
}

fn verify_untimed_gates(corpus: &Corpus) -> Result<(), Box<dyn StdError>> {
    let (boundary_remove, boundary_move) = verify_boundary_semantics(corpus)?;
    let (remove_forward, remove_inverse, remove_stale, remove_foreign) = verify_remove_durable(corpus)?;
    let (move_forward, move_inverse, move_stale, move_foreign) = verify_move_durable(corpus)?;
    let (one_slide, unknown, dependency, mce, signed, limits) = verify_refusal_gates(corpus)?;
    if !(boundary_remove
        && boundary_move
        && remove_forward
        && remove_inverse
        && remove_stale
        && remove_foreign
        && move_forward
        && move_inverse
        && move_stale
        && move_foreign
        && one_slide
        && unknown
        && dependency
        && mce
        && signed
        && limits)
    {
        return Err(format!(
            "PPTX slide-boundary untimed gate failed: boundary_remove={boundary_remove}, boundary_move={boundary_move}, remove=({remove_forward},{remove_inverse},{remove_stale},{remove_foreign}), move=({move_forward},{move_inverse},{move_stale},{move_foreign}), refusals=({one_slide},{unknown},{dependency},{mce},{signed}), limits={limits}"
        )
        .into());
    }
    Ok(())
}

fn selected_slide_name(corpus: &Corpus, position: usize) -> Result<String, Box<dyn StdError>> {
    let package = Package::from_vec(corpus.archive.clone())?;
    let snapshot = package.opened_presentation()?;
    let slide = snapshot
        .slides()
        .get(position)
        .ok_or("PPTX boundary selected slide is missing")?;
    Ok(slide.part_name().membername().to_owned())
}

/// Run one boundary selector.  Setup and every correctness/refusal gate are
/// outside the phase clocks; the reported elapsed vector is only plan +
/// commit + sequential publication.
pub(super) fn run(
    case: Case,
    corpus: &Corpus,
    warmup_iterations: usize,
    samples: usize,
) -> Result<CaseResult, Box<dyn StdError>> {
    let operation = match case {
        Case::PptxSlideRemoveBoundarySave => Operation::Remove,
        Case::PptxSlideMoveBoundarySave => Operation::Move,
        _ => return Err("PPTX boundary runner received an unrelated case".into()),
    };
    let representative_position = match operation {
        Operation::Remove => REMOVE_POSITIONS
            .iter()
            .position(|position| *position == REPRESENTATIVE_REMOVE_POSITION)
            .ok_or("PPTX removal representative is not in the boundary set")?,
        Operation::Move => 0,
    };
    let expected_output = match operation {
        Operation::Remove => corpus
            .remove_outputs
            .get(representative_position)
            .ok_or("PPTX removal representative output is missing")?,
        Operation::Move => corpus
            .move_outputs
            .get(representative_position)
            .ok_or("PPTX move representative output is missing")?,
    };
    let expected_output_sha256 = sha256_hex(expected_output);
    let selected_slide = (operation == Operation::Remove)
        .then(|| selected_slide_name(corpus, REPRESENTATIVE_REMOVE_POSITION))
        .transpose()?;

    let (boundary_semantics_verified, no_op_verified) = verify_boundary_semantics(corpus)?;
    let (durable_forward_verified, durable_inverse_verified, stale_refusal_verified, foreign_refusal_verified) =
        match operation {
            Operation::Remove => verify_remove_durable(corpus)?,
            Operation::Move => verify_move_durable(corpus)?,
        };
    let (
        one_slide_refusal_verified,
        unknown_member_refusal_verified,
        dependency_refusal_verified,
        markup_compatibility_refusal_verified,
        signed_refusal_verified,
        limits_verified,
    ) =
        verify_refusal_gates(corpus)?;
    if !(boundary_semantics_verified
        && durable_forward_verified
        && durable_inverse_verified
        && stale_refusal_verified
        && foreign_refusal_verified
        && one_slide_refusal_verified
        && unknown_member_refusal_verified
        && dependency_refusal_verified
        && markup_compatibility_refusal_verified
        && signed_refusal_verified
        && limits_verified)
    {
        return Err("PPTX boundary selector has incomplete untimed gates".into());
    }

    let mut elapsed = Vec::new();
    let mut plan_ns = Vec::new();
    let mut commit_ns = Vec::new();
    let mut publication_ns = Vec::new();
    let mut reopen_ns = Vec::new();
    let mut sink_summaries = Vec::new();
    let mut output_digests = Vec::new();
    reserve_exact(&mut elapsed, samples, "PPTX boundary elapsed samples")?;
    reserve_exact(&mut plan_ns, samples, "PPTX boundary plan samples")?;
    reserve_exact(&mut commit_ns, samples, "PPTX boundary commit samples")?;
    reserve_exact(&mut publication_ns, samples, "PPTX boundary publication samples")?;
    reserve_exact(&mut reopen_ns, samples, "PPTX boundary reopen samples")?;
    reserve_exact(&mut sink_summaries, samples, "PPTX boundary sink samples")?;
    reserve_exact(&mut output_digests, samples, "PPTX boundary output digests")?;

    let maximum = u64::try_from(expected_output.len())?
        .checked_mul(2)
        .and_then(|value| value.checked_add(64 * 1024))
        .ok_or("PPTX boundary sink budget overflows u64")?;
    for iteration in 0..iteration_count(warmup_iterations, samples)? {
        let mut package = Package::from_vec(corpus.archive.clone())?;
        let snapshot = package.opened_presentation()?;
        let mut transaction = (operation == Operation::Move).then(|| snapshot.edit());

        let started = Instant::now();
        let removal_plan = if operation == Operation::Remove {
            Some(snapshot.plan_slide_removal(REPRESENTATIVE_REMOVE_POSITION)?)
        } else {
            transaction
                .as_mut()
                .ok_or("PPTX move transaction disappeared before planning")?
                .move_slide(REPRESENTATIVE_MOVE.0, REPRESENTATIVE_MOVE.1)?;
            None
        };
        let plan_duration = started.elapsed();

        let started = Instant::now();
        match operation {
            Operation::Remove => {
                package.apply_slide_removal_plan(
                    removal_plan
                        .as_ref()
                        .ok_or("PPTX removal plan disappeared before commit")?,
                )?;
            },
            Operation::Move => {
                let commit = transaction
                    .take()
                    .ok_or("PPTX move transaction disappeared before commit")?
                    .commit()?;
                package.apply_opened_presentation_commit(commit)?;
            },
        }
        let commit_duration = started.elapsed();

        let mut sink = CountingSink::bounded(maximum, MAX_WRITE);
        sink.reserve_budget()?;
        let started = Instant::now();
        package.opc()?.to_stream(&mut sink)?;
        let publication_duration = started.elapsed();
        let sink_summary = sink.summary();
        let output = sink.bytes;

        let started = Instant::now();
        let reopened = Package::from_vec(output.clone())?;
        let reopened_snapshot = reopened.opened_presentation()?;
        let reopened_names = reopened_snapshot
            .slides()
            .iter()
            .map(|slide| slide.name().to_owned())
            .collect::<Vec<_>>();
        let reopen_duration = started.elapsed();
        let expected_names = match operation {
            Operation::Remove => corpus
                .slide_names
                .iter()
                .enumerate()
                .filter(|(index, _)| *index != REPRESENTATIVE_REMOVE_POSITION)
                .map(|(_, name)| name.clone())
                .collect::<Vec<_>>(),
            Operation::Move => {
                let mut names = corpus.slide_names.clone();
                let moved = names.remove(REPRESENTATIVE_MOVE.0);
                names.insert(REPRESENTATIVE_MOVE.1, moved);
                names
            },
        };
        if reopened_names != expected_names
            || reopened_snapshot.slides().len() != expected_names.len()
            || output != *expected_output
            || !verify_untouched_raw_members(
                &corpus.archive,
                &output,
                operation,
                selected_slide.as_deref(),
            )?
        {
            return Err("PPTX boundary measured output failed semantic/raw gates".into());
        }
        let digest = sha256_hex(&output);
        let plan_elapsed = elapsed_ns(plan_duration)?;
        let commit_elapsed = elapsed_ns(commit_duration)?;
        let publication_elapsed = elapsed_ns(publication_duration)?;
        let total = plan_elapsed
            .checked_add(commit_elapsed)
            .and_then(|value| value.checked_add(publication_elapsed))
            .ok_or("PPTX boundary phase duration overflows u64")?;
        if iteration >= warmup_iterations {
            elapsed.push(total);
            plan_ns.push(plan_elapsed);
            commit_ns.push(commit_elapsed);
            publication_ns.push(publication_elapsed);
            reopen_ns.push(elapsed_ns(reopen_duration)?);
            sink_summaries.push(sink_summary);
            output_digests.push(digest);
        }
        std::hint::black_box(output);
    }

    let sink = deterministic_sink_summary(&sink_summaries, case.name())?;
    if output_digests.iter().any(|digest| digest != &expected_output_sha256) {
        return Err("PPTX boundary output digest is not deterministic".into());
    }
    if sha256_hex(&corpus.archive) != corpus.manifest.archive_sha256 {
        return Err("PPTX boundary source archive changed during measurement".into());
    }
    let (partial_sink_verified, zero_sink_verified) = verify_partial_zero_sinks(
        &corpus.archive,
        operation,
        representative_position,
        expected_output,
    )?;
    let summary = PptxSlideBoundarySummary {
        implementation: "litchi_pptx::opened::Snapshot + Transaction + Package publication",
        operation: operation.name(),
        timing_scope: "representative boundary plan/staging + commit + OPC sequential publication; setup, reopen, topology, raw-member, durable-patch, refusal, sink, and bounds checks excluded",
        performance_claim: "correctness and phase evidence only; no latency, speedup, allocation, RSS, or physical-I/O claim",
        source_archive_bytes: u64::try_from(corpus.archive.len())?,
        source_archive_sha256: corpus.manifest.archive_sha256.clone(),
        output_archive_bytes: u64::try_from(expected_output.len())?,
        expected_output_sha256: expected_output_sha256.clone(),
        source_slide_count: SLIDE_COUNT,
        output_slide_count: SLIDE_COUNT - usize::from(operation == Operation::Remove),
        remove_positions: REMOVE_POSITIONS.to_vec(),
        move_pairs: MOVE_PAIRS.to_vec(),
        plan_ns,
        commit_ns,
        publication_ns,
        reopen_ns,
        output_sha256: output_digests,
        boundary_semantics_verified,
        no_op_verified,
        untouched_raw_members_verified: true,
        deterministic_output_verified: true,
        durable_forward_verified,
        durable_inverse_verified,
        stale_refusal_verified,
        foreign_refusal_verified,
        partial_sink_verified,
        zero_sink_verified,
        one_slide_refusal_verified,
        unknown_member_refusal_verified,
        dependency_refusal_verified,
        markup_compatibility_refusal_verified,
        signed_refusal_verified,
        limits_verified,
        source_immutability_verified: true,
    };
    if !summary.partial_sink_verified || !summary.zero_sink_verified {
        return Err("PPTX boundary partial/zero sink gates failed".into());
    }
    let source = SourceSummary {
        pptx_slide_boundaries: Some(summary),
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
        output_sha256: Some(expected_output_sha256),
        operation_metrics: None,
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn corpus_is_bounded_and_opt_in() {
        let corpus = build_corpus().expect("PPTX boundary corpus");
        let again = build_corpus().expect("PPTX boundary corpus rebuild");
        assert_eq!(corpus.slide_names.len(), SLIDE_COUNT);
        assert_eq!(corpus.remove_outputs.len(), REMOVE_POSITIONS.len());
        assert_eq!(corpus.move_outputs.len(), MOVE_PAIRS.len());
        assert_eq!(corpus.archive, again.archive);
        assert_eq!(
            serde_json::to_vec(&corpus.manifest).expect("PPTX boundary manifest encoding"),
            serde_json::to_vec(&again.manifest).expect("PPTX boundary manifest encoding")
        );
        assert_eq!(corpus.manifest.archive_bytes, EXPECTED_SOURCE_ARCHIVE_BYTES);
        assert_eq!(corpus.manifest.archive_sha256, EXPECTED_SOURCE_ARCHIVE_SHA256);
        assert_eq!(
            corpus.manifest.archive_member_count,
            EXPECTED_SOURCE_ARCHIVE_MEMBER_COUNT
        );
        assert!(!Case::DEFAULT.contains(&Case::PptxSlideRemoveBoundarySave));
        assert!(!Case::DEFAULT.contains(&Case::PptxSlideMoveBoundarySave));
        assert_eq!(super::super::parse_case("pptx_slide_remove_boundary_save"), Some(Case::PptxSlideRemoveBoundarySave));
        assert_eq!(super::super::parse_case("pptx_slide_move_boundary_save"), Some(Case::PptxSlideMoveBoundarySave));
    }

    #[test]
    fn selectors_emit_phases_and_all_boundary_gates() {
        let corpus = build_corpus().expect("PPTX boundary corpus");
        for case in [Case::PptxSlideRemoveBoundarySave, Case::PptxSlideMoveBoundarySave] {
            let result = run(case, &corpus, 0, 1).expect("PPTX boundary run");
            let summary = result
                .source
                .expect("PPTX boundary source evidence")
                .pptx_slide_boundaries
                .expect("PPTX boundary summary");
            assert_eq!(summary.plan_ns.len(), 1);
            assert_eq!(summary.commit_ns.len(), 1);
            assert_eq!(summary.publication_ns.len(), 1);
            assert_eq!(summary.reopen_ns.len(), 1);
            assert!(summary.boundary_semantics_verified);
            assert!(summary.no_op_verified);
            assert!(summary.untouched_raw_members_verified);
            assert!(summary.durable_forward_verified);
            assert!(summary.durable_inverse_verified);
            assert!(summary.stale_refusal_verified);
            assert!(summary.foreign_refusal_verified);
            assert!(summary.partial_sink_verified);
            assert!(summary.zero_sink_verified);
            assert!(summary.one_slide_refusal_verified);
            assert!(summary.unknown_member_refusal_verified);
            assert!(summary.dependency_refusal_verified);
            assert!(summary.markup_compatibility_refusal_verified);
            assert!(summary.signed_refusal_verified);
            assert!(summary.limits_verified);
            assert!(summary.source_immutability_verified);
        }
    }
}
