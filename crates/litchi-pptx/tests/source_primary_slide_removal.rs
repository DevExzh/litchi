#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "integration fixtures and assertions intentionally panic"
)]

//! Primary source-owner integration coverage for dependency-free slide removal.
//!
//! These tests exercise removal through the immutable source facade. They keep
//! the original owner untouched, compose removal with order and text edits, and
//! transport the resulting topology through the durable patch wire format.

use std::collections::BTreeMap;
use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, Ordering};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, Position,
    ReadAt, SourceVersion,
};
use litchi_opc::{OpcPackage, PackURI, PackageWriter, SourceLimits};
use litchi_pptx::builder::Builder;
use litchi_pptx::{Error, SlideRemovalRefusal, source::Package};
use soapberry_zip::office::ArchiveReader;

macro_rules! assert_refused_kind {
    ($operation:expr, $expected:expr) => {{
        match $operation {
            Err(Error::Refused(refusal)) => {
                assert_eq!(refusal.kind(), $expected);
            },
            Err(error) => panic!("unexpected error: {error:?}"),
            Ok(_) => panic!("operation unexpectedly succeeded"),
        }
    }};
}

const MAIN: &str = "/ppt/presentation.xml";

fn fixture() -> Vec<u8> {
    let mut builder = Builder::new().unwrap();
    {
        let presentation = builder.presentation().unwrap();
        for text in ["A", "B", "C"] {
            presentation
                .add_slide()
                .unwrap()
                .add_text_box(text, 0, 0, 914_400, 914_400);
        }
    }
    let authored = builder.build().unwrap();
    authored.to_bytes().unwrap()
}

fn fresh(bytes: &[u8]) -> Package {
    Package::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap()
}

fn texts(package: &Package) -> Vec<String> {
    let presentation = package.presentation().unwrap();
    (0..presentation.slide_count().unwrap())
        .map(|position| presentation.slide(position).unwrap().text().unwrap())
        .collect()
}

fn source_positions(package: &Package) -> Vec<Position> {
    let count = package.presentation().unwrap().slide_count().unwrap();
    (0..count)
        .map(|position| package.source_position(Position::new(position)).unwrap())
        .collect()
}

fn decoded_members(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    let archive =
        ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default()).unwrap();
    let names = archive.file_names().map(str::to_owned).collect::<Vec<_>>();
    names
        .into_iter()
        .map(|name| {
            let payload = archive.read(&name).unwrap();
            (name, payload)
        })
        .collect()
}

fn replace_once(bytes: &[u8], before: &[u8], after: &[u8]) -> Vec<u8> {
    let start = bytes
        .windows(before.len())
        .position(|window| window == before)
        .expect("fixture fragment must exist");
    let mut result = Vec::with_capacity(bytes.len() - before.len() + after.len());
    result.extend_from_slice(&bytes[..start]);
    result.extend_from_slice(after);
    result.extend_from_slice(&bytes[start + before.len()..]);
    result
}

fn with_part(bytes: &[u8], path: &str, replacement: &[u8]) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
    package
        .get_part_mut(&PackURI::new(path).unwrap())
        .unwrap()
        .set_blob(replacement.to_vec());
    PackageWriter::to_bytes(&package).unwrap()
}

fn with_hidden_selected_slide(bytes: &[u8]) -> Vec<u8> {
    let package = OpcPackage::from_bytes(bytes, SourceLimits::default()).unwrap();
    let presentation = package.get_part(&PackURI::new(MAIN).unwrap()).unwrap();
    let selected_relationship = presentation
        .rels()
        .iter()
        .find(|relationship| {
            relationship
                .target_partname()
                .is_ok_and(|target| target.as_str() == "/ppt/slides/slide2.xml")
        })
        .expect("the fixture must bind slide 2")
        .r_id()
        .to_owned();
    let original = presentation.blob();
    let marker = format!(r#"r:id="{selected_relationship}"/>"#);
    let replacement = format!(r#"show="0" r:id="{selected_relationship}"/>"#);
    let updated = replace_once(original, marker.as_bytes(), replacement.as_bytes());
    with_part(bytes, MAIN, &updated)
}

#[derive(Debug)]
struct VersionedSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset too large"))?;
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let end = start.saturating_add(output.len()).min(self.bytes.len());
        output[..end - start].copy_from_slice(&self.bytes[start..end]);
        Ok(end - start)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(43, self.revision.load(Ordering::SeqCst)))
    }
}

fn managed_context() -> (CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "pptx-source-primary-removal",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(u64::MAX).unwrap(),
        0,
    )
    .unwrap();
    (
        cancellation_source,
        ExecutionContext::new(budget, cancellation, limits),
    )
}

#[test]
fn primary_remove_commits_count_readback_and_reopen_without_mutating_source() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_slide(Position::new(1)).unwrap());
    let changed = edit.commit().unwrap().into_package();

    assert_eq!(changed.presentation().unwrap().slide_count().unwrap(), 2);
    assert_eq!(texts(&changed), ["A", "C"]);
    assert_eq!(
        source_positions(&changed),
        [Position::new(0), Position::new(2)]
    );
    assert_eq!(source.to_bytes().unwrap(), bytes);

    let reopened = fresh(&changed.to_bytes().unwrap());
    assert_eq!(reopened.presentation().unwrap().slide_count().unwrap(), 2);
    assert_eq!(texts(&reopened), ["A", "C"]);
    assert_eq!(
        source_positions(&reopened),
        [Position::new(0), Position::new(1)]
    );
}

#[test]
fn dropped_primary_removal_rolls_back_and_order_text_removal_compose_atomically() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut dropped = source.edit().unwrap();
    assert!(dropped.remove_slide(Position::new(1)).unwrap());
    drop(dropped);
    assert_eq!(source.to_bytes().unwrap(), bytes);
    assert_eq!(texts(&source), ["A", "B", "C"]);

    let mut edit = source.edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    assert!(
        edit.set_shape_text(Position::new(0), 0_usize, "B*")
            .unwrap()
    );
    assert!(edit.remove_slide(Position::new(1)).unwrap());
    let changed = edit.commit().unwrap().into_package();

    assert_eq!(texts(&changed), ["B*", "A"]);
    assert_eq!(
        source_positions(&changed),
        [Position::new(1), Position::new(0)]
    );
    assert_eq!(source.to_bytes().unwrap(), bytes);
}

#[test]
fn removal_then_order_and_text_composition_keeps_stable_identity() {
    let source = fresh(&fixture());
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_slide(Position::new(1)).unwrap());
    assert!(edit.move_slide(Position::new(0), Position::new(1)).unwrap());
    assert!(
        edit.set_shape_text(Position::new(0), 0_usize, "C*")
            .unwrap()
    );
    let changed = edit.commit().unwrap().into_package();

    assert_eq!(texts(&changed), ["C*", "A"]);
    assert_eq!(
        source_positions(&changed),
        [Position::new(2), Position::new(0)]
    );
}

#[test]
fn repeated_removals_stop_at_the_final_slide_guard() {
    let source = fresh(&fixture());
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_slide(Position::new(1)).unwrap());
    assert!(edit.remove_slide(Position::new(1)).unwrap());
    let error = edit.remove_slide(Position::new(0)).err().unwrap();
    assert!(matches!(
        error,
        Error::Refused(refusal) if refusal.kind() == SlideRemovalRefusal::FinalSlide.kind()
    ));
    let changed = edit.commit().unwrap().into_package();
    assert_eq!(texts(&changed), ["A"]);
}

#[test]
fn inverse_of_removing_an_already_text_edited_slide_restores_original_payload() {
    let source = fresh(&fixture());
    let mut edit = source.edit().unwrap();
    assert!(
        edit.set_shape_text(Position::new(1), 0_usize, "B*")
            .unwrap()
    );
    assert!(edit.remove_slide(Position::new(1)).unwrap());
    let (changed, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(texts(&changed), ["A", "C"]);

    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(texts(&restored), ["A", "B", "C"]);
}

#[test]
fn local_removal_patch_replays_and_inverse_restores_decoded_state() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut edit = source.edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    assert!(
        edit.set_shape_text(Position::new(0), 0_usize, "B*")
            .unwrap()
    );
    assert!(edit.remove_slide(Position::new(1)).unwrap());
    let (changed, patch) = edit.commit().unwrap().into_parts();

    let replayed = source.apply(&patch).unwrap().into_package();
    assert_eq!(texts(&replayed), ["B*", "A"]);
    assert_eq!(
        source_positions(&replayed),
        [Position::new(1), Position::new(0)]
    );

    let restored = changed.apply(&patch.inverse()).unwrap().into_package();
    assert_eq!(texts(&restored), ["A", "B", "C"]);
    assert_eq!(
        source_positions(&restored),
        [Position::new(0), Position::new(1), Position::new(2)]
    );
    assert_eq!(
        decoded_members(&restored.to_bytes().unwrap()),
        decoded_members(&bytes)
    );
}

#[test]
fn durable_removal_patch_reopens_forward_and_inverse_after_removed_output() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut edit = source.edit().unwrap();
    assert!(edit.move_slide(Position::new(0), Position::new(2)).unwrap());
    assert!(
        edit.set_shape_text(Position::new(0), 0_usize, "B*")
            .unwrap()
    );
    assert!(edit.remove_slide(Position::new(1)).unwrap());
    let (changed, patch) = edit.commit().unwrap().into_parts();
    let durable = litchi_pptx::source::Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();

    let replayed = fresh(&bytes).apply(&durable).unwrap().into_package();
    assert_eq!(texts(&replayed), ["B*", "A"]);
    assert_eq!(replayed.presentation().unwrap().slide_count().unwrap(), 2);

    let removed_output = changed.to_bytes().unwrap();
    let reopened_removed = fresh(&removed_output);
    let inverse =
        litchi_pptx::source::Patch::from_bytes(&durable.inverse().to_bytes().unwrap()).unwrap();
    let restored = reopened_removed.apply(&inverse).unwrap().into_package();
    assert_eq!(texts(&restored), ["A", "B", "C"]);
    assert_eq!(
        decoded_members(&restored.to_bytes().unwrap()),
        decoded_members(&bytes)
    );
}

#[test]
fn restored_middle_slide_can_be_read_edited_reordered_replayed_and_removed_again() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut remove = source.edit().unwrap();
    assert!(remove.remove_slide(Position::new(1)).unwrap());
    let (removed, removal_patch) = remove.commit().unwrap().into_parts();
    let removal =
        litchi_pptx::source::Patch::from_bytes(&removal_patch.to_bytes().unwrap()).unwrap();

    let reopened_removed = fresh(&removed.to_bytes().unwrap());
    let restored = reopened_removed
        .apply(
            &litchi_pptx::source::Patch::from_bytes(&removal.inverse().to_bytes().unwrap())
                .unwrap(),
        )
        .unwrap()
        .into_package();
    let middle = restored
        .presentation()
        .unwrap()
        .slide(Position::new(1))
        .unwrap();
    assert_eq!(middle.text().unwrap(), "B");
    assert!(!middle.scene().unwrap().is_empty());

    let mut edit = restored.edit().unwrap();
    assert!(
        edit.set_shape_text(Position::new(1), 0_usize, "B-restored")
            .unwrap()
    );
    assert!(edit.move_slide(Position::new(1), Position::new(2)).unwrap());
    let (ordered, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(texts(&ordered), ["A", "C", "B-restored"]);

    let published_ordered = ordered.to_bytes().unwrap();
    let reopened_ordered = fresh(&published_ordered);
    assert_eq!(texts(&reopened_ordered), ["A", "C", "B-restored"]);

    let durable = litchi_pptx::source::Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
    assert_eq!(
        texts(&restored.apply(&durable).unwrap().into_package()),
        ["A", "C", "B-restored"]
    );
    let reopened_restored = fresh(&restored.to_bytes().unwrap());
    let replayed = reopened_restored.apply(&durable).unwrap().into_package();
    assert_eq!(texts(&replayed), ["A", "C", "B-restored"]);
    let inverse =
        litchi_pptx::source::Patch::from_bytes(&durable.inverse().to_bytes().unwrap()).unwrap();
    let replayed_inverse = replayed.apply(&inverse).unwrap().into_package();
    assert_eq!(texts(&replayed_inverse), ["A", "B", "C"]);

    let mut remove_again = ordered.edit().unwrap();
    assert!(remove_again.remove_slide(Position::new(2)).unwrap());
    let removed_again = remove_again.commit().unwrap().into_package();
    assert_eq!(texts(&removed_again), ["A", "C"]);
    assert!(removed_again.same_content(&reopened_removed).unwrap());
    assert_eq!(
        removed_again.to_bytes().unwrap(),
        reopened_removed.to_bytes().unwrap()
    );
}

#[test]
fn durable_inverse_reopens_after_two_removals_and_restores_every_member() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_slide(Position::new(2)).unwrap());
    assert!(edit.remove_slide(Position::new(1)).unwrap());
    let (changed, patch) = edit.commit().unwrap().into_parts();
    assert_eq!(texts(&changed), ["A"]);

    let durable = litchi_pptx::source::Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
    let removed = fresh(&changed.to_bytes().unwrap());
    let inverse =
        litchi_pptx::source::Patch::from_bytes(&durable.inverse().to_bytes().unwrap()).unwrap();
    let restored = removed.apply(&inverse).unwrap().into_package();
    assert_eq!(texts(&restored), ["A", "B", "C"]);
    assert_eq!(
        decoded_members(&restored.to_bytes().unwrap()),
        decoded_members(&bytes)
    );
    // The physical source now contains only A, but the effective catalog has
    // three slides. Removing A must use the effective final-slide guard.
    let mut remove_physical = restored.edit().unwrap();
    remove_physical.remove_slide(Position::new(0)).unwrap();
    assert_eq!(
        texts(&remove_physical.commit().unwrap().into_package()),
        ["B", "C"]
    );
}

#[test]
fn final_slide_and_hidden_slide_are_refused_without_changing_the_edit() {
    let one_slide = {
        let mut builder = Builder::new().unwrap();
        builder
            .presentation()
            .unwrap()
            .add_slide()
            .unwrap()
            .add_text_box("only", 0, 0, 914_400, 914_400);
        let package = builder.build().unwrap();
        package.to_bytes().unwrap()
    };
    let source = fresh(&one_slide);
    let mut edit = source.edit().unwrap();
    let error = edit.remove_slide(Position::new(0)).err().unwrap();
    assert!(matches!(
        error,
        Error::Refused(refusal) if refusal.kind() == SlideRemovalRefusal::FinalSlide.kind()
    ));
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), one_slide);

    let hidden_bytes = with_hidden_selected_slide(&fixture());
    let hidden = fresh(&hidden_bytes);
    let mut edit = hidden.edit().unwrap();
    let error = edit.remove_slide(Position::new(1)).err().unwrap();
    assert!(matches!(
        error,
        Error::Refused(refusal) if refusal.kind() == SlideRemovalRefusal::UnknownSemanticSurface.kind()
    ));
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.package().to_bytes().unwrap(), hidden_bytes);
}

#[test]
fn stale_source_and_cancellation_reject_removal_commit_atomically() {
    let source_handle = Arc::new(VersionedSource::new(fixture()));
    let source = Package::from_read_at(source_handle.clone()).unwrap();
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_slide(Position::new(1)).unwrap());
    source_handle.changed();
    assert!(edit.commit().is_err());

    let (cancellation_source, context) = managed_context();
    let source = Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(fixture())),
        litchi_pptx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = source.edit().unwrap();
    assert!(edit.remove_slide(Position::new(1)).unwrap());
    cancellation_source.cancel();
    assert!(edit.commit().is_err());
}

fn reseal_topology_wire(bytes: &mut [u8]) {
    use sha2::{Digest, Sha256};
    assert_eq!(&bytes[..8], b"LPSX0003");
    let body = bytes.len() - 32;
    let mut hash = Sha256::new();
    hash.update(b"litchi-pptx-source-topology-wire-v3\0");
    hash.update(&bytes[..body]);
    bytes[body..].copy_from_slice(&hash.finalize());
}

#[test]
fn resealed_topology_wire_cannot_change_unrelated_root_xml() {
    let bytes = fixture();
    let opc = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
    let root = opc.get_part(&PackURI::new(MAIN).unwrap()).unwrap().blob();
    let root = replace_once(
        root,
        b"</p:presentation>",
        b"<!--canary--></p:presentation>",
    );
    let bytes = with_part(&bytes, MAIN, &root);
    let source = fresh(&bytes);
    let mut edit = source.edit().unwrap();
    edit.remove_slide(Position::new(1)).unwrap();
    let (_, patch) = edit.commit().unwrap().into_parts();
    let mut wire = patch.to_bytes().unwrap();
    let canaries = wire
        .windows(6)
        .enumerate()
        .filter_map(|(i, bytes)| (bytes == b"canary").then_some(i))
        .collect::<Vec<_>>();
    assert_eq!(canaries.len(), 2);
    wire[canaries[1]] = b'f';
    reseal_topology_wire(&mut wire);
    let forged = litchi_pptx::source::Patch::from_bytes(&wire).unwrap();
    assert!(source.apply(&forged).is_err());
    assert_eq!(source.to_bytes().unwrap(), bytes);
}

#[test]
fn resealed_topology_wire_cannot_change_shape_preset_with_text() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut edit = source.edit().unwrap();
    edit.set_shape_text(Position::new(0), 0_usize, "unique-new-text")
        .unwrap();
    edit.remove_slide(Position::new(2)).unwrap();
    let (_, patch) = edit.commit().unwrap().into_parts();
    let mut wire = patch.to_bytes().unwrap();
    let text = wire
        .windows(b"unique-new-text".len())
        .position(|b| b == b"unique-new-text")
        .unwrap();
    // Bounds are now an admitted edit surface. Preset shape geometry remains
    // outside it even when a sender recomputes the transport checksum.
    let marker = b"prst=\"rect\"";
    let offset = wire[..text]
        .windows(marker.len())
        .rposition(|b| b == marker)
        .unwrap();
    wire[offset..offset + marker.len()].copy_from_slice(b"prst=\"line\"");
    reseal_topology_wire(&mut wire);
    let forged = litchi_pptx::source::Patch::from_bytes(&wire).unwrap();
    assert!(source.apply(&forged).is_err());
    assert_eq!(source.to_bytes().unwrap(), bytes);
}

#[test]
fn removal_patch_and_inverse_retain_managed_images_until_last_drop() {
    use litchi_core::Resource;
    let budget = Budget::root(
        "primary-removal-retention",
        Limits::new(
            32 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (_cancel, cancellation) = CancellationSource::pair();
    let context = ExecutionContext::new(
        budget.clone(),
        cancellation,
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MAX, 0).unwrap(),
    );
    let source = Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(fixture())),
        litchi_pptx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = source.edit().unwrap();
    edit.remove_slide(Position::new(1)).unwrap();
    let commit = edit.commit().unwrap();
    let patch = commit.patch().clone();
    let inverse = patch.inverse();
    drop(commit);
    drop(source);
    assert!(budget.used(Resource::Memory) > 0);
    drop(patch);
    assert!(budget.used(Resource::Memory) > 0);
    drop(inverse);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn removal_authored_after_a_committed_text_edit_guards_effective_preimages() {
    let bytes = fixture();
    let source = fresh(&bytes);
    let mut text = source.edit().unwrap();
    text.set_shape_text(Position::new(1), 0_usize, "B-previous")
        .unwrap();
    let edited = text.commit().unwrap().into_package();
    let mut removal = edited.edit().unwrap();
    removal.remove_slide(Position::new(1)).unwrap();
    let (removed, patch) = removal.commit().unwrap().into_parts();
    let durable = litchi_pptx::source::Patch::from_bytes(&patch.to_bytes().unwrap()).unwrap();
    assert_eq!(
        texts(&edited.apply(&durable).unwrap().into_package()),
        ["A", "C"]
    );
    assert!(source.apply(&durable).is_err());
    let restored = removed.apply(&durable.inverse()).unwrap().into_package();
    assert_eq!(texts(&restored), ["A", "B-previous", "C"]);
    let reopened = fresh(&removed.to_bytes().unwrap());
    assert_eq!(
        texts(&reopened.apply(&durable.inverse()).unwrap().into_package()),
        ["A", "B-previous", "C"]
    );
}

#[test]
fn alternate_layout_target_spelling_is_refused_before_noninvertible_removal() {
    let mut members = decoded_members(&fixture());
    let rels = members.get_mut("ppt/slides/_rels/slide2.xml.rels").unwrap();
    *rels = replace_once(rels, b"../slideLayouts/", b"/ppt/slideLayouts/");
    let mut archive = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    for (name, bytes) in members {
        archive.write_stored(&name, &bytes).unwrap();
    }
    let bytes = archive.finish_to_bytes().unwrap();
    let source = fresh(&bytes);
    let mut edit = source.edit().unwrap();
    assert_refused_kind!(
        edit.remove_slide(Position::new(1)),
        SlideRemovalRefusal::AmbiguousTopology.kind()
    );
    let commit = edit.commit().unwrap();
    assert!(commit.patch().is_noop());
    assert_eq!(commit.into_package().to_bytes().unwrap(), bytes);
}

#[test]
fn repeated_structural_revision_edits_do_not_retain_prior_candidate_chains() {
    use litchi_core::Resource;
    let budget = Budget::root(
        "primary-removal-revisions",
        Limits::new(
            32 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (_cancel, cancellation) = CancellationSource::pair();
    let context = ExecutionContext::new(
        budget.clone(),
        cancellation,
        ExecutionLimits::new(NonZeroUsize::MIN, NonZeroUsize::MIN, NonZeroU64::MAX, 0).unwrap(),
    );
    let mut package = Package::from_read_at_with_options(
        Arc::new(OwnedSource::new(fixture())),
        litchi_pptx::OpenOptions::default()
            .source_limits(SourceLimits::default())
            .execution_context(context),
    )
    .unwrap();
    let mut edit = package.edit().unwrap();
    edit.remove_slide(Position::new(1)).unwrap();
    package = edit.commit().unwrap().into_package();
    let mut steady = 0;
    for index in 0..16 {
        let mut edit = package.edit().unwrap();
        edit.set_shape_text(
            Position::new(0),
            0_usize,
            if index % 2 == 0 { "xx" } else { "yy" },
        )
        .unwrap();
        package = edit.commit().unwrap().into_package();
        if index == 2 {
            steady = budget.used(Resource::Memory);
        }
        if index > 2 {
            assert_eq!(budget.used(Resource::Memory), steady);
        }
    }
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}
