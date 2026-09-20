//! Focused correctness and resource tests for the 0704 bounded slide-MCE
//! projection retention candidate.
//!
//! This module is included by `opened/mod.rs` only under `#[cfg(test)]`. It
//! exercises the production policy, observability, and release methods plus
//! the private test-only seams described below. It adds no runtime observer,
//! counter, or global cache.
//!
//! The module uses these private test hooks (their implementation remains in
//! the production owner's `cfg(test)` module):
//!
//! ```text
//! super::model::mce_retention_test_hooks::slide_output(
//!     snapshot: &super::Snapshot,
//!     slide_index: usize,
//! ) -> Option<Arc<Vec<u8>>>
//! ```
//!
//! ```text
//! super::model::mce_retention_test_hooks::checked_charge_for_test(
//!     entries_capacity: usize,
//!     output_capacity: usize,
//! ) -> Option<usize>
//! ```
//!
//! It must return the cached owned default-MCE output for that slide, if one
//! was retained, while returning `None` for a borrowed/uncached output.  The
//! charge hook must call the same checked production arithmetic used by table
//! admission. Both hooks are test-only and must not expose storage through the
//! normal API. The tests use the output hook only for Arc identity and
//! resource assertions; semantic bytes are checked through public APIs.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused retention tests use panic-on-fixture-failure assertions"
)]

use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_opc::constants::content_type as ct;
use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, Relationships, TargetMode};

use super::{Limits, Snapshot};
use crate::{Error, Package, Result};

use super::model::mce_retention_test_hooks;

const MCE_NAMESPACE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const PRESENTATION_NAMESPACE: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";

fn marked_slides_package(count: usize, marked: &[usize]) -> Result<Package> {
    let mut package = Package::new()?;
    for index in 0..count {
        let presentation = package.presentation_mut()?;
        let slide = presentation.add_slide()?;
        slide.set_title(&format!("retention slide {index}"));
        if index == 0 {
            slide.set_notes("First notes");
        }
    }
    let bytes = package.to_bytes()?;
    let mut package = Package::from_vec(bytes)?;
    for &index in marked {
        add_mce_marker(&mut package, index)?;
    }
    Package::from_vec(package.to_bytes()?)
}

fn add_mce_marker(package: &mut Package, index: usize) -> Result<()> {
    let part_name =
        PackURI::new(format!("/ppt/slides/slide{}.xml", index + 1)).map_err(Error::Invalid)?;
    let xml = std::str::from_utf8(package.opc.get_part(&part_name)?.blob())
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned();
    let marker = format!(
        r#"xmlns:mc="{MCE_NAMESPACE}" xmlns:p14="urn:litchi-retention-test" mc:Ignorable="p14" "#
    );
    let marked = xml.replacen("<p:sld ", &format!("<p:sld {marker}"), 1);
    if marked == xml {
        return Err(Error::Invalid("retention fixture has no p:sld root".into()));
    }
    package
        .opc
        .get_part_mut(&part_name)?
        .set_blob(marked.into_bytes());
    Ok(())
}

fn plain_slides_package(count: usize) -> Result<Package> {
    marked_slides_package(count, &[])
}

fn retention_limits(maximum: usize) -> Limits {
    Limits::default().with_max_retained_mce_bytes(maximum)
}

fn output(snapshot: &Snapshot, index: usize) -> Option<Arc<Vec<u8>>> {
    mce_retention_test_hooks::slide_output(snapshot, index)
}

fn assert_same_output(left: &Snapshot, right: &Snapshot, index: usize) {
    let left = output(left, index).expect("the unchanged marked slide is retained");
    let right = output(right, index).expect("the unchanged marked slide remains retained");
    assert!(
        Arc::ptr_eq(&left, &right),
        "unchanged slide output must be shared across the retained snapshots"
    );
}

fn assert_distinct_output(left: &Snapshot, right: &Snapshot, index: usize) {
    let left = output(left, index).expect("the source marked slide is retained");
    let right = output(right, index).expect("the changed marked slide is retained");
    assert!(
        !Arc::ptr_eq(&left, &right),
        "a changed slide must not reuse the source slide's processed output"
    );
}

#[test]
fn the_retention_policy_defaults_to_one_mib_and_zero_disables_only_mce_retention() {
    let defaults = Limits::default();
    assert_eq!(defaults.max_retained_mce_bytes(), 1024 * 1024);

    let disabled = defaults.with_max_retained_mce_bytes(0);
    assert_eq!(disabled.max_retained_mce_bytes(), 0);
    assert_eq!(disabled.max_parts(), defaults.max_parts());
    assert_eq!(disabled.max_patch_bytes(), defaults.max_patch_bytes());
    assert_eq!(disabled.max_text_bytes(), defaults.max_text_bytes());
    assert_eq!(
        disabled.max_history_entries(),
        defaults.max_history_entries()
    );
    assert_eq!(disabled.max_history_bytes(), defaults.max_history_bytes());
    assert_eq!(
        disabled.max_retained_candidate_bytes(),
        defaults.max_retained_candidate_bytes()
    );
}

#[test]
fn marker_free_slides_never_create_a_retained_mce_projection() -> Result<()> {
    let package = plain_slides_package(2)?;
    let snapshot = package.opened_presentation_with_limits(Limits::default())?;

    assert_eq!(snapshot.retained_mce_bytes(), 0);
    assert!(output(&snapshot, 0).is_none());
    assert!(output(&snapshot, 1).is_none());
    Ok(())
}

#[test]
fn disabled_retention_keeps_the_same_snapshot_semantics_without_owned_outputs() -> Result<()> {
    let package = marked_slides_package(2, &[0, 1])?;
    let retained = package.opened_presentation_with_limits(Limits::default())?;
    let disabled = package.opened_presentation_with_limits(retention_limits(0))?;

    assert!(retained.retained_mce_bytes() > 0);
    assert_eq!(disabled.retained_mce_bytes(), 0);
    assert_eq!(disabled.slides(), retained.slides());
    assert_eq!(disabled.revision(), retained.revision());
    assert!(output(&disabled, 0).is_none());
    assert!(output(&disabled, 1).is_none());
    Ok(())
}

#[test]
fn fresh_captures_are_equal_but_snapshot_local_retention_is_not_a_global_cache() -> Result<()> {
    let package = marked_slides_package(1, &[0])?;
    let first = package.opened_presentation_with_limits(Limits::default())?;
    let second = package.opened_presentation_with_limits(Limits::default())?;

    assert_eq!(first.slides(), second.slides());
    assert_eq!(first.revision(), second.revision());
    assert_eq!(first.retained_mce_bytes(), second.retained_mce_bytes());
    let first_output = output(&first, 0).expect("first capture retains marked output");
    let second_output = output(&second, 0).expect("second capture retains marked output");
    assert!(
        !Arc::ptr_eq(&first_output, &second_output),
        "fresh captures must not share a global transformed-XML cache"
    );
    Ok(())
}

#[test]
fn the_exact_observed_budget_fits_and_one_byte_below_falls_back_without_refusal() -> Result<()> {
    let package = marked_slides_package(1, &[0])?;
    let wide = package.opened_presentation_with_limits(Limits::default())?;
    let required = wide.retained_mce_bytes();
    assert!(required > 0, "the fixture must exercise owned MCE output");

    let exact = package.opened_presentation_with_limits(retention_limits(required))?;
    assert_eq!(exact.retained_mce_bytes(), required);
    assert!(output(&exact, 0).is_some());
    assert_eq!(exact.slides(), wide.slides());
    assert_eq!(exact.revision(), wide.revision());

    let below = package.opened_presentation_with_limits(retention_limits(required - 1))?;
    assert_eq!(below.retained_mce_bytes(), 0);
    assert!(output(&below, 0).is_none());
    assert_eq!(below.slides(), wide.slides());
    assert_eq!(below.revision(), wide.revision());
    Ok(())
}

#[test]
fn tiny_ceilings_and_maximum_usize_are_fallbacks_or_successes_never_refusals() -> Result<()> {
    let package = marked_slides_package(2, &[0, 1])?;

    for maximum in [0, 1] {
        let snapshot = package.opened_presentation_with_limits(retention_limits(maximum))?;
        assert_eq!(snapshot.retained_mce_bytes(), 0, "ceiling {maximum}");
        assert_eq!(snapshot.slides().len(), 2);
    }

    let unbounded_by_policy =
        package.opened_presentation_with_limits(retention_limits(usize::MAX))?;
    assert!(unbounded_by_policy.retained_mce_bytes() > 0);
    assert!(unbounded_by_policy.retained_mce_bytes() < usize::MAX);
    Ok(())
}

#[test]
fn checked_retention_charge_rejects_entry_and_output_overflow() {
    assert!(mce_retention_test_hooks::checked_charge_for_test(1, 0).is_some());
    assert!(
        mce_retention_test_hooks::checked_charge_for_test(usize::MAX, 0).is_none(),
        "entry metadata multiplication must fail closed on overflow"
    );
    assert!(
        mce_retention_test_hooks::checked_charge_for_test(1, usize::MAX).is_none(),
        "output plus metadata must fail closed on overflow"
    );
}

#[test]
fn a_failed_optional_table_reservation_releases_both_candidate_owners() {
    let (uncached, owners_released) = crate::parts::refused_mce_reservation_releases_for_test();
    assert!(uncached, "a reservation refusal must leave the memo empty");
    assert!(
        owners_released,
        "failed admission must release both candidate owners"
    );
}

#[test]
fn a_clone_shares_outputs_release_is_local_and_drop_does_not_evict_the_original() -> Result<()> {
    let package = marked_slides_package(1, &[0])?;
    let source = package.opened_presentation_with_limits(Limits::default())?;
    let mut clone = source.clone();
    assert_same_output(&source, &clone, 0);
    assert_eq!(clone.retained_mce_bytes(), source.retained_mce_bytes());

    clone.release_retained_mce();
    assert_eq!(clone.retained_mce_bytes(), 0);
    assert!(output(&clone, 0).is_none());
    assert!(source.retained_mce_bytes() > 0);
    assert!(output(&source, 0).is_some());

    drop(clone);
    assert!(source.retained_mce_bytes() > 0);
    Ok(())
}

#[test]
fn transaction_and_commit_release_hooks_are_value_preserving() -> Result<()> {
    let package = marked_slides_package(1, &[0])?;
    let source = package.opened_presentation_with_limits(Limits::default())?;
    let revision = source.revision();
    let slides = source.slides().to_vec();

    let mut transaction = source.edit();
    assert!(transaction.retained_mce_bytes() > 0);
    transaction.release_retained_mce();
    assert_eq!(transaction.retained_mce_bytes(), 0);
    let mut commit = transaction.commit()?;
    assert!(!commit.is_changed());
    assert_eq!(commit.snapshot().revision(), revision);
    assert_eq!(commit.snapshot().slides(), slides.as_slice());
    assert_eq!(commit.snapshot().retained_mce_bytes(), 0);

    let mut changed = source.edit();
    changed.set_shape_text(0, crate::shape::Key::Index(0), "release candidate")?;
    let mut changed_commit = changed.commit()?;
    assert!(changed_commit.is_changed());
    let changed_revision = changed_commit.snapshot().revision();
    changed_commit.release_retained_mce();
    assert_eq!(changed_commit.snapshot().retained_mce_bytes(), 0);
    assert_eq!(changed_commit.snapshot().revision(), changed_revision);
    assert!(changed_commit.is_changed());
    commit.release_retained_mce();
    Ok(())
}

#[test]
fn no_op_commit_returns_the_same_retained_snapshot_and_publication_preserves_it() -> Result<()> {
    let mut package = marked_slides_package(2, &[0, 1])?;
    let source = package.opened_presentation_with_limits(Limits::default())?;

    let commit = source.edit().commit()?;
    assert!(!commit.is_changed());
    assert_eq!(commit.snapshot().revision(), source.revision());
    assert_eq!(
        commit.snapshot().retained_mce_bytes(),
        source.retained_mce_bytes()
    );
    assert_same_output(&source, commit.snapshot(), 0);
    assert_same_output(&source, commit.snapshot(), 1);

    let published = package.apply_opened_presentation_commit(commit)?;
    assert_eq!(published.revision(), source.revision());
    assert_same_output(&source, &published, 0);
    assert_same_output(&source, &published, 1);
    Ok(())
}

#[test]
fn one_edit_misses_the_changed_slide_but_shares_the_untouched_slide_through_publication()
-> Result<()> {
    let mut package = marked_slides_package(2, &[0, 1])?;
    let source = package.opened_presentation_with_limits(Limits::default())?;

    let mut transaction = source.edit();
    assert!(transaction.set_shape_text(0, crate::shape::Key::Index(0), "changed once")?);
    let commit = transaction.commit()?;
    assert!(commit.is_changed());
    assert_distinct_output(&source, commit.snapshot(), 0);
    assert_same_output(&source, commit.snapshot(), 1);

    let candidate = commit.snapshot().clone();
    let published = package.apply_opened_presentation_commit(commit)?;
    assert_distinct_output(&source, &published, 0);
    assert_same_output(&candidate, &published, 1);
    assert_eq!(
        published.retained_mce_bytes(),
        candidate.retained_mce_bytes()
    );
    Ok(())
}

#[test]
fn two_edits_keep_the_previous_unchanged_output_and_miss_each_newly_changed_slide() -> Result<()> {
    let package = marked_slides_package(3, &[0, 1, 2])?;
    let source = package.opened_presentation_with_limits(Limits::default())?;

    let mut first_edit = source.edit();
    first_edit.set_shape_text(0, crate::shape::Key::Index(0), "first edit")?;
    let first = first_edit.commit()?;
    let first_snapshot = first.snapshot().clone();
    assert_distinct_output(&source, &first_snapshot, 0);
    assert_same_output(&source, &first_snapshot, 1);
    assert_same_output(&source, &first_snapshot, 2);

    let mut second_edit = first_snapshot.edit();
    second_edit.set_shape_text(1, crate::shape::Key::Index(0), "second edit")?;
    let second = second_edit.commit()?;
    assert_same_output(&first_snapshot, second.snapshot(), 0);
    assert_distinct_output(&first_snapshot, second.snapshot(), 1);
    assert_same_output(&first_snapshot, second.snapshot(), 2);
    Ok(())
}

#[test]
fn a_root_relationship_edit_runs_capture_validation_while_slide_outputs_hit() -> Result<()> {
    let package = marked_slides_package(2, &[0, 1])?;
    let source = package.opened_presentation_with_limits(Limits::default())?;
    let first_name = source.slides[0].part_name.clone();
    let second_name = source.slides[1].part_name.clone();

    let mut transaction = source.edit();
    assert!(transaction.move_slide(0, 1)?);
    let commit = transaction.commit()?;
    let candidate = commit.snapshot();
    assert_eq!(candidate.slides[0].part_name, second_name);
    assert_eq!(candidate.slides[1].part_name, first_name);
    assert_same_output_for_part(&source, candidate, &first_name);
    assert_same_output_for_part(&source, candidate, &second_name);
    Ok(())
}

#[test]
fn a_notes_edit_runs_notes_validation_while_unchanged_slide_outputs_hit() -> Result<()> {
    let package = marked_slides_package(2, &[0, 1])?;
    let source = package.opened_presentation_with_limits(Limits::default())?;
    let first_name = source.slides[0].part_name.clone();
    let second_name = source.slides[1].part_name.clone();

    let mut transaction = source.edit();
    assert!(transaction.set_notes_text(0, "retained notes edit")?);
    let commit = transaction.commit()?;
    assert_same_output_for_part(&source, commit.snapshot(), &first_name);
    assert_same_output_for_part(&source, commit.snapshot(), &second_name);
    Ok(())
}

#[test]
fn changed_content_type_on_an_unchanged_raw_slide_remains_a_typed_refusal() -> Result<()> {
    let package = marked_slides_package(1, &[0])?;
    let source = package.opened_presentation_with_limits(Limits::default())?;
    let slide_name = source.slides[0].part_name.clone();
    let mut candidate = source.package.as_ref().clone();
    candidate
        .get_part_mut(&slide_name)?
        .set_content_type("application/octet-stream".into())?;

    let error = capture_committed_candidate(&source, &candidate)
        .expect_err("content-type metadata must be checked before a cache hit can publish");
    assert!(matches!(error, Error::ContentType { .. }));
    Ok(())
}

#[test]
fn changed_root_relationship_metadata_remains_a_typed_refusal_on_slide_cache_hits() -> Result<()> {
    let package = marked_slides_package(2, &[0, 1])?;
    let source = package.opened_presentation_with_limits(Limits::default())?;
    let presentation_name = source.presentation_name.clone();
    let relationship_id = crate::parts::PresentationPart::from_package(&source.package)?
        .slide_references()?
        .get(1)
        .expect("second slide reference")
        .relationship_id()
        .to_owned();
    let mut candidate = source.package.as_ref().clone();
    candidate
        .get_part_mut(&presentation_name)?
        .rels_mut()
        .remove(&relationship_id);

    let error = capture_committed_candidate(&source, &candidate)
        .expect_err("root relationship validation must not be skipped by slide reuse");
    assert!(error.to_string().contains("missing relationship"));
    Ok(())
}

#[test]
fn changed_notes_root_remains_a_typed_refusal_while_slides_are_cache_hits() -> Result<()> {
    let package = marked_slides_package(2, &[0, 1])?;
    let source = package.opened_presentation_with_limits(Limits::default())?;
    let notes_name = PackURI::new("/ppt/notesSlides/notesSlide1.xml").map_err(Error::Invalid)?;
    let mut candidate = source.package.as_ref().clone();
    candidate
        .get_part_mut(&notes_name)?
        .set_blob(format!(r#"<p:wrong xmlns:p="{PRESENTATION_NAMESPACE}"/>"#).into_bytes());

    let error = capture_committed_candidate(&source, &candidate)
        .expect_err("notes validation must remain after reusable slide projections");
    assert!(
        error
            .to_string()
            .contains("invalid notes root or namespace")
    );
    Ok(())
}

fn capture_committed_candidate(source: &Snapshot, candidate: &OpcPackage) -> Result<Snapshot> {
    let (revision, digests) =
        super::model::package_fingerprint_with_memo(candidate, Some(&source.part_digests))?;
    super::model::capture_with_revision_and_digests_and_mce(
        candidate,
        source.limits,
        source.physical_source_provenance,
        revision,
        digests,
        source.retained_mce.as_deref(),
    )
}

#[test]
fn a_cache_hit_adds_no_blob_or_blob_arc_observations_over_the_uncached_route() -> Result<()> {
    let (package, observations) = counted_marked_slides_package()?;
    let source = package.opened_presentation_with_limits(Limits::default())?;
    observations.store(0, Ordering::SeqCst);
    let mut cached_candidate = source.package.as_ref().clone();
    rewrite_notes_on_opc(&mut cached_candidate, "cached route")?;
    let cached = capture_committed_candidate(&source, &cached_candidate)?;
    let cached_observations = observations.load(Ordering::SeqCst);
    assert_same_output(&source, &cached, 0);

    let disabled_package = Package::from_opc_package(source.package.as_ref().clone())?;
    let disabled = disabled_package.opened_presentation_with_limits(retention_limits(0))?;
    observations.store(0, Ordering::SeqCst);
    let mut uncached_candidate = disabled.package.as_ref().clone();
    rewrite_notes_on_opc(&mut uncached_candidate, "uncached route")?;
    capture_committed_candidate(&disabled, &uncached_candidate)?;
    let uncached_observations = observations.load(Ordering::SeqCst);

    assert_eq!(
        cached_observations, uncached_observations,
        "MCE cache lookup must not add raw blob observations"
    );
    Ok(())
}

fn counted_marked_slides_package() -> Result<(Package, Arc<AtomicUsize>)> {
    let package = marked_slides_package(1, &[0])?;
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").map_err(Error::Invalid)?;
    let observations = Arc::new(AtomicUsize::new(0));
    let mut opc = package.opc.clone();
    let counted =
        CountingBlobPart::from_part(opc.get_part(&slide_name)?, Arc::clone(&observations))?;
    assert!(opc.remove_part(&slide_name));
    opc.try_add_part(Box::new(counted))?;
    Ok((Package::from_opc_package(opc)?, observations))
}

fn rewrite_notes_on_opc(opc: &mut OpcPackage, text: &str) -> Result<()> {
    let notes_name = PackURI::new("/ppt/notesSlides/notesSlide1.xml").map_err(Error::Invalid)?;
    let xml = std::str::from_utf8(opc.get_part(&notes_name)?.blob())
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned();
    let replacement = xml.replacen("First notes", text, 1);
    opc.get_part_mut(&notes_name)?
        .set_blob(replacement.into_bytes());
    Ok(())
}

fn assert_same_output_for_part(left: &Snapshot, right: &Snapshot, part_name: &PackURI) {
    let left_index = left
        .slides
        .iter()
        .position(|slide| slide.part_name == *part_name)
        .expect("source part must be present");
    let right_index = right
        .slides
        .iter()
        .position(|slide| slide.part_name == *part_name)
        .expect("candidate part must be present");
    assert_same_output_at(left, right, left_index, right_index);
}

fn assert_same_output_at(left: &Snapshot, right: &Snapshot, left_index: usize, right_index: usize) {
    let left = output(left, left_index).expect("the unchanged marked slide is retained");
    let right = output(right, right_index).expect("the unchanged marked slide remains retained");
    assert!(
        Arc::ptr_eq(&left, &right),
        "unchanged slide output must be shared across the retained snapshots"
    );
}

#[test]
fn retained_bytes_are_aggregate_capacity_and_metadata_and_partial_admission_is_a_fallback()
-> Result<()> {
    let package = marked_slides_package(2, &[0, 1])?;
    let wide = package.opened_presentation_with_limits(Limits::default())?;
    let total = wide.retained_mce_bytes();
    assert!(total > 0);

    // The one-slide amount is measured from an isolated capture so the
    // admission check is against the implementation's actual vector capacity
    // and metadata charge, rather than guessed XML length.
    let one = marked_slides_package(1, &[0])?.opened_presentation_with_limits(Limits::default())?;
    let one_bytes = one.retained_mce_bytes();
    assert!(one_bytes > 0 && one_bytes < total);
    let one_output = output(&one, 0).expect("the one-slide fixture retains its output");
    assert!(
        one_output.capacity() > one_output.len(),
        "the fixture must exercise capacity-based, rather than length-based, charging"
    );
    assert!(
        one_bytes > one_output.capacity(),
        "retained charge must include entry/table metadata in addition to output capacity"
    );

    let partial = package.opened_presentation_with_limits(retention_limits(one_bytes))?;
    assert!(partial.retained_mce_bytes() <= one_bytes);
    assert!(partial.retained_mce_bytes() < total);
    assert_eq!(partial.slides(), wide.slides());
    assert_eq!(partial.revision(), wide.revision());

    // A checked reservation must not turn an arithmetic overflow at the
    // policy boundary into an error or an over-budget resident cache.
    let maximum = package.opened_presentation_with_limits(retention_limits(usize::MAX))?;
    assert!(maximum.retained_mce_bytes() < usize::MAX);
    Ok(())
}

#[test]
fn malformed_root_name_and_notes_inputs_keep_typed_refusal_precedence() -> Result<()> {
    // These checks deliberately run through the public capture route.  A
    // retention candidate may skip only successful MCE transformation; it may
    // never cache an error or move the established validation order.
    let mut root = marked_slides_package(2, &[0, 1])?;
    replace_slide_root(&mut root, 1, "p:wrong")?;
    let error = root
        .opened_presentation_with_limits(Limits::default())
        .expect_err("invalid later root must remain a typed refusal");
    assert!(
        error
            .to_string()
            .contains("slide part does not have a p:sld root")
    );

    let mut name = marked_slides_package(2, &[0, 1])?;
    duplicate_slide_name(&mut name, 0)?;
    let error = name
        .opened_presentation_with_limits(Limits::default())
        .expect_err("malformed slide name must remain a typed refusal");
    assert!(error.to_string().contains("duplicated attribute"));

    let mut notes = marked_slides_package(2, &[0, 1])?;
    replace_notes_root(&mut notes)?;
    let error = notes
        .opened_presentation_with_limits(Limits::default())
        .expect_err("malformed notes root must remain a typed refusal");
    assert!(
        error
            .to_string()
            .contains("invalid notes root or namespace")
    );

    let mut missing_relationship = marked_slides_package(2, &[0, 1])?;
    remove_slide_relationship(&mut missing_relationship, 1)?;
    let error = missing_relationship
        .opened_presentation_with_limits(Limits::default())
        .expect_err("missing relationship must remain a typed refusal");
    assert!(error.to_string().contains("missing relationship"));
    Ok(())
}

fn remove_slide_relationship(package: &mut Package, index: usize) -> Result<()> {
    let presentation_name = PackURI::new("/ppt/presentation.xml").map_err(Error::Invalid)?;
    let relationship_id = crate::parts::PresentationPart::from_package(&package.opc)?
        .slide_references()?
        .get(index)
        .ok_or_else(|| Error::Invalid("retention fixture slide is absent".into()))?
        .relationship_id()
        .to_owned();
    package
        .opc
        .get_part_mut(&presentation_name)?
        .rels_mut()
        .remove(&relationship_id);
    Ok(())
}

fn replace_slide_root(package: &mut Package, index: usize, root: &str) -> Result<()> {
    rewrite_slide(package, index, |xml| {
        xml.replacen("<p:sld ", &format!("<{root} "), 1).replacen(
            "</p:sld>",
            &format!("</{root}>"),
            1,
        )
    })
}

fn duplicate_slide_name(package: &mut Package, index: usize) -> Result<()> {
    rewrite_slide(package, index, |xml| {
        let needle = format!(r#" name="Slide {}""#, 256 + index);
        xml.replacen(&needle, &format!(r#"{needle} name="duplicate""#), 1)
    })
}

fn rewrite_slide(
    package: &mut Package,
    index: usize,
    rewrite: impl FnOnce(String) -> String,
) -> Result<()> {
    let part_name =
        PackURI::new(format!("/ppt/slides/slide{}.xml", index + 1)).map_err(Error::Invalid)?;
    let xml = std::str::from_utf8(package.opc.get_part(&part_name)?.blob())
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned();
    package
        .opc
        .get_part_mut(&part_name)?
        .set_blob(rewrite(xml).into_bytes());
    Ok(())
}

fn replace_notes_root(package: &mut Package) -> Result<()> {
    let part_name = PackURI::new("/ppt/notesSlides/notesSlide1.xml").map_err(Error::Invalid)?;
    package
        .opc
        .get_part_mut(&part_name)?
        .set_blob(format!(r#"<p:wrong xmlns:p="{PRESENTATION_NAMESPACE}"/>"#).into_bytes());
    Ok(())
}

#[derive(Clone, Debug)]
struct CountingBlobPart {
    inner: BlobPart,
    observations: Arc<AtomicUsize>,
}

impl CountingBlobPart {
    fn from_part(part: &dyn Part, observations: Arc<AtomicUsize>) -> Result<Self> {
        let mut inner = BlobPart::new(
            part.partname().clone(),
            part.content_type().to_owned(),
            part.blob().to_vec(),
        );
        copy_relationships(part, &mut inner)?;
        Ok(Self {
            inner,
            observations,
        })
    }
}

impl Part for CountingBlobPart {
    fn blob(&self) -> &[u8] {
        self.observations.fetch_add(1, Ordering::SeqCst);
        self.inner.blob()
    }

    fn blob_arc(&self) -> Arc<Vec<u8>> {
        self.observations.fetch_add(1, Ordering::SeqCst);
        self.inner.blob_arc()
    }

    fn content_type(&self) -> &str {
        self.inner.content_type()
    }

    fn partname(&self) -> &PackURI {
        self.inner.partname()
    }

    fn rels(&self) -> &Relationships {
        self.inner.rels()
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        self.inner.rels_mut()
    }

    fn set_blob(&mut self, blob: Vec<u8>) {
        self.inner.set_blob(blob);
    }

    fn set_content_type(&mut self, content_type: String) -> litchi_opc::Result<()> {
        self.inner.set_content_type(content_type)
    }
}

/// A foreign part whose visible payload and returned `Arc` intentionally
/// disagree.  The retention key must apply the same alias proof as the
/// existing payload digest memo: this payload may be read, but its foreign
/// `Arc` must never be pinned or used as a source identity.
#[derive(Clone, Debug)]
struct MismatchedBlobArcPart {
    inner: BlobPart,
    foreign: Arc<Vec<u8>>,
}

impl MismatchedBlobArcPart {
    fn new(name: PackURI, visible: Vec<u8>, foreign: Arc<Vec<u8>>) -> Self {
        Self {
            inner: BlobPart::new(name, ct::PML_SLIDE.into(), visible),
            foreign,
        }
    }

    fn from_part(part: &dyn Part, foreign: Arc<Vec<u8>>) -> Result<Self> {
        let mut output = Self::new(part.partname().clone(), part.blob().to_vec(), foreign);
        copy_relationships(part, &mut output.inner)?;
        Ok(output)
    }
}

fn copy_relationships(source: &dyn Part, destination: &mut BlobPart) -> Result<()> {
    for relationship in source.rels().iter() {
        destination.rels_mut().try_add_relationship(
            relationship.reltype().to_owned(),
            relationship.target_ref().to_owned(),
            relationship.r_id().to_owned(),
            if relationship.is_external() {
                TargetMode::External
            } else {
                TargetMode::Internal
            },
        )?;
    }
    Ok(())
}

impl Part for MismatchedBlobArcPart {
    fn blob(&self) -> &[u8] {
        self.inner.blob()
    }

    fn blob_arc(&self) -> Arc<Vec<u8>> {
        Arc::clone(&self.foreign)
    }

    fn content_type(&self) -> &str {
        self.inner.content_type()
    }

    fn partname(&self) -> &PackURI {
        self.inner.partname()
    }

    fn rels(&self) -> &Relationships {
        self.inner.rels()
    }

    fn rels_mut(&mut self) -> &mut Relationships {
        self.inner.rels_mut()
    }

    fn set_blob(&mut self, blob: Vec<u8>) {
        self.inner.set_blob(blob);
    }

    fn set_content_type(&mut self, content_type: String) -> litchi_opc::Result<()> {
        self.inner.set_content_type(content_type)
    }
}

#[test]
fn foreign_blob_arc_and_clone_allocations_cannot_pin_or_answer_for_foreign_payloads() -> Result<()>
{
    let package = marked_slides_package(1, &[0])?;
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").map_err(Error::Invalid)?;
    let foreign = Arc::new(b"foreign payload with a different identity".to_vec());

    let mut mismatched = package.opc.clone();
    let foreign_part =
        MismatchedBlobArcPart::from_part(mismatched.get_part(&slide_name)?, Arc::clone(&foreign))?;
    assert!(mismatched.remove_part(&slide_name));
    mismatched.try_add_part(Box::new(foreign_part))?;
    let mismatched = Package::from_opc_package(mismatched)?;
    let foreign_weak = Arc::downgrade(&foreign);
    let snapshot = mismatched.opened_presentation_with_limits(Limits::default())?;
    // The normal XML route still succeeds, but the foreign Arc is not retained
    // as a source owner.  The slide may have an owned output produced from the
    // visible bytes; it must never report the foreign allocation itself.
    if let Some(retained) = output(&snapshot, 0) {
        assert!(!Arc::ptr_eq(&retained, &foreign));
    }

    // A separate allocation carrying equal visible bytes is not an identity
    // proof. Reopening it must not answer from the first snapshot's output.
    let mut separate = mismatched.opc.clone();
    let copied = separate.get_part(&slide_name)?.blob().to_vec();
    let relationships = separate.get_part(&slide_name)?.rels().clone();
    assert!(separate.remove_part(&slide_name));
    let mut honest = BlobPart::new(slide_name, ct::PML_SLIDE.into(), copied);
    for relationship in relationships.iter() {
        honest.rels_mut().try_add_relationship(
            relationship.reltype().to_owned(),
            relationship.target_ref().to_owned(),
            relationship.r_id().to_owned(),
            if relationship.is_external() {
                TargetMode::External
            } else {
                TargetMode::Internal
            },
        )?;
    }
    separate.try_add_part(Box::new(honest))?;
    let separate = Package::from_opc_package(separate)?;
    let second = separate.opened_presentation_with_limits(Limits::default())?;
    if let (Some(first), Some(second)) = (output(&snapshot, 0), output(&second, 0)) {
        assert!(!Arc::ptr_eq(&first, &second));
    }

    // The package and its snapshot are the only legitimate owners of the
    // foreign `Arc`.  Once both are gone, a retained projection must not keep
    // that unrelated allocation alive through an unproven source key.
    drop(snapshot);
    drop(mismatched);
    drop(foreign);
    assert!(
        foreign_weak.upgrade().is_none(),
        "a non-aliasing foreign blob_arc must never be pinned by retention"
    );
    Ok(())
}

#[test]
fn rebound_drops_unproven_foreign_owners() -> Result<()> {
    let package = marked_slides_package(1, &[0])?;
    let source = package.opened_presentation_with_limits(Limits::default())?;
    let slide_name = source.slides[0].part_name.clone();
    let original = source.package.get_part(&slide_name)?;
    let foreign = Arc::new(b"rebind foreign owner".to_vec());
    let foreign_part = MismatchedBlobArcPart::from_part(original, Arc::clone(&foreign))?;

    let mut rebound_package = source.package.as_ref().clone();
    assert!(rebound_package.remove_part(&slide_name));
    rebound_package.try_add_part(Box::new(foreign_part))?;
    let rebound = source.rebound_to(&rebound_package);

    assert_eq!(rebound.revision(), source.revision());
    assert_eq!(rebound.slides(), source.slides());
    assert_eq!(rebound.retained_mce_bytes(), 0);
    assert!(output(&rebound, 0).is_none());
    drop(rebound);
    drop(foreign);
    Ok(())
}

#[test]
fn rebound_keeps_a_package_owned_allocation_but_new_equal_slide_bytes_miss() -> Result<()> {
    let mut package = marked_slides_package(1, &[0])?;
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").map_err(Error::Invalid)?;
    let unrelated_name = PackURI::new("/retained-source.bin").map_err(Error::Invalid)?;
    let raw = package.opc.get_part(&slide_name)?.blob_arc();
    package.opc.try_add_part(Box::new(BlobPart::new_shared(
        unrelated_name.clone(),
        "application/octet-stream".into(),
        Arc::clone(&raw),
    )))?;
    let package = Package::from_opc_package(package.opc)?;
    let source = package.opened_presentation_with_limits(Limits::default())?;
    let original_output = output(&source, 0).expect("marked source output");

    let mut moved_owner = source.package.as_ref().clone();
    moved_owner
        .get_part_mut(&slide_name)?
        .set_blob(raw.as_ref().clone());
    let rebound = source.rebound_to(&moved_owner);
    assert_eq!(rebound.revision(), source.revision());
    assert_eq!(rebound.retained_mce_bytes(), source.retained_mce_bytes());
    assert!(output(&rebound, 0).is_none(), "equal new bytes are a miss");
    let recaptured = capture_committed_candidate(&source, &moved_owner)?;
    let fresh_output = output(&recaptured, 0).expect("freshly processed slide");
    assert!(!Arc::ptr_eq(&original_output, &fresh_output));
    assert_eq!(original_output.as_slice(), fresh_output.as_slice());

    moved_owner
        .get_part_mut(&unrelated_name)?
        .set_blob(raw.as_ref().clone());
    let removed_owner = rebound.rebound_to(&moved_owner);
    assert_eq!(removed_owner.revision(), rebound.revision());
    assert_eq!(removed_owner.retained_mce_bytes(), 0);
    Ok(())
}
