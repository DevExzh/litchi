//! Binding proofs for the retained candidate archive and its budget.
//!
//! A cross-package copy plan retains the serialized candidate archive it
//! already built when the operation's intersected
//! [`Limits::max_retained_candidate_bytes`] admits its length, and applying
//! the plan reuses those bytes instead of serializing and deflating the
//! candidate a second time.  These tests pin the three things that make that
//! safe: the retained bytes are the bytes the application publishes, a
//! substituted archive is refused before anything is published, and no
//! application verdict moves between the retained and the rebuilt route.
//! They also pin the budget itself: a candidate above the ceiling falls back
//! to the rebuild rather than refusing, and the tighter of the two snapshots'
//! ceilings binds.
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "retention tests use panic-on-fixture-failure assertions"
)]

use std::sync::Arc;

use super::{CrossSlideCopyPlan, physical_package_fingerprint, seal_physical_revision};
use crate::opened::model::Limits;
use crate::{Error, Package, Result};

fn owned_slides(titles: &[&str]) -> Result<Package> {
    let mut package = Package::new()?;
    for title in titles {
        package.presentation_mut()?.add_slide()?.set_title(title);
    }
    let bytes = package.to_bytes()?;
    Package::from_vec(bytes)
}

/// Give every slide a distinct producer-visible name so a cross-package copy
/// is not refused for a name collision.
fn rename_slide(package: &mut Package, index: usize, name: &str) -> Result<()> {
    let slide = package
        .opened_presentation()?
        .slides()
        .get(index)
        .cloned()
        .ok_or(Error::SlideIndexOutOfBounds {
            index,
            len: package.opened_presentation()?.slides().len(),
        })?;
    let part_name = slide.part_name().clone();
    let xml = std::str::from_utf8(package.opc.get_part(&part_name)?.blob())
        .map_err(|error| Error::Xml(format!("slide XML is not UTF-8: {error}")))?;
    let renamed = xml.replacen(
        &format!(r#"name="Slide {}""#, slide.id()),
        &format!(r#"name="{name}""#),
        1,
    );
    if renamed == xml {
        return Err(Error::Invalid(format!(
            "test slide {index} has no canonical producer name"
        )));
    }
    package
        .opc
        .get_part_mut(&part_name)?
        .set_blob(renamed.into_bytes());
    Ok(())
}

fn cross_copy_fixture() -> Result<(Package, Package)> {
    let mut source = owned_slides(&["retain-source"])?;
    rename_slide(&mut source, 0, "retain-source")?;
    let source = Package::from_vec(source.to_bytes()?)?;

    let mut destination = owned_slides(&["retain-destination-a", "retain-destination-b"])?;
    rename_slide(&mut destination, 0, "retain-destination-a")?;
    rename_slide(&mut destination, 1, "retain-destination-b")?;
    let destination = Package::from_vec(destination.to_bytes()?)?;
    Ok((source, destination))
}

/// The default policy with one member replaced.
fn limits_retaining(max_retained_candidate_bytes: usize) -> Result<Limits> {
    let default = Limits::default();
    Limits::new(
        default.max_parts(),
        default.max_patch_bytes(),
        default.max_text_bytes(),
        default.max_history_entries(),
        default.max_history_bytes(),
        max_retained_candidate_bytes,
    )
    .ok_or_else(|| Error::Invalid("test limits are invalid".into()))
}

fn plan_for(source: &Package, destination: &Package) -> Result<CrossSlideCopyPlan> {
    let source_snapshot = source.opened_presentation()?;
    let destination_snapshot = destination.opened_presentation()?;
    destination_snapshot.plan_cross_slide_copy(&source_snapshot, 0_usize, 0_usize, 1)
}

fn plan_under(
    source: &Package,
    destination: &Package,
    source_limits: Limits,
    destination_limits: Limits,
) -> Result<CrossSlideCopyPlan> {
    let source_snapshot = source.opened_presentation_with_limits(source_limits)?;
    let destination_snapshot = destination.opened_presentation_with_limits(destination_limits)?;
    destination_snapshot.plan_cross_slide_copy(&source_snapshot, 0_usize, 0_usize, 1)
}

/// The default limits retain, the retained bytes are exactly the bytes the
/// application publishes, and they seal to the plan's own physical revision.
#[test]
fn a_plan_retains_the_archive_its_application_publishes() -> Result<()> {
    let (source, mut destination) = cross_copy_fixture()?;
    let plan = plan_for(&source, &destination)?;

    let held = plan
        .candidate
        .0
        .as_ref()
        .expect("the default retained-candidate budget retains an ordinary candidate");
    assert_eq!(plan.retained_candidate_bytes(), Some(held.archive.len()));
    assert!(
        held.archive.len() <= Limits::default().max_retained_candidate_bytes(),
        "a retained archive is never larger than the budget that admitted it"
    );
    assert_eq!(
        held.bound,
        Limits::default().max_patch_bytes(),
        "the retained archive records the archive bound it was produced under"
    );
    let archive = Arc::clone(&held.archive);
    assert_eq!(
        seal_physical_revision(
            {
                use sha2::{Digest, Sha256};
                let mut digest = Sha256::new();
                digest.update(archive.as_slice());
                digest.finalize().into()
            },
            archive.len(),
        )?,
        plan.target_physical_revision(),
        "the retained archive seals to the plan's own physical revision"
    );

    let mut applied = Package::from_vec(destination.to_bytes()?)?;
    applied.apply_cross_slide_copy_plan(&source, &plan)?;
    assert_eq!(
        applied.to_bytes()?,
        *archive,
        "the published archive is byte-identical to the archive the plan retained"
    );
    Ok(())
}

/// A candidate above the budget is not retained and the plan still applies:
/// exceeding the budget is a fallback to the rebuild, never a refusal.
#[test]
fn a_candidate_above_the_budget_falls_back_to_the_rebuild() -> Result<()> {
    let (source, mut destination) = cross_copy_fixture()?;

    let retained = plan_for(&source, &destination)?;
    let archive_bytes = retained
        .retained_candidate_bytes()
        .expect("the default budget retains this candidate");

    // One byte below the candidate is the smallest budget that refuses it.
    let tight = limits_retaining(archive_bytes - 1)?;
    let unretained = plan_under(&source, &destination, tight, tight)?;
    assert_eq!(
        unretained.retained_candidate_bytes(),
        None,
        "a candidate above the retained-candidate budget is not retained"
    );
    // Exactly the candidate's length is admitted: the comparison is `<=`.
    let exact = limits_retaining(archive_bytes)?;
    assert_eq!(
        plan_under(&source, &destination, exact, exact)?.retained_candidate_bytes(),
        Some(archive_bytes),
        "a candidate exactly at the budget is retained"
    );

    // The unretained plan is not refused: it applies, and publishes the same
    // archive the retained plan publishes.
    let mut rebuilt = Package::from_vec(destination.to_bytes()?)?;
    let rebuilt_snapshot = rebuilt.apply_cross_slide_copy_plan(&source, &unretained)?;
    let mut reused = Package::from_vec(destination.to_bytes()?)?;
    let reused_snapshot = reused.apply_cross_slide_copy_plan(&source, &retained)?;
    assert_eq!(rebuilt.to_bytes()?, reused.to_bytes()?);
    assert_eq!(rebuilt_snapshot.revision(), reused_snapshot.revision());
    assert_eq!(
        unretained.target_physical_revision(),
        retained.target_physical_revision(),
        "the budget changes what is held, never what is planned"
    );
    Ok(())
}

/// The budget is intersected like every other member: the tighter of the two
/// snapshots' ceilings binds, whichever side carries it.
#[test]
fn the_tighter_retained_candidate_budget_binds() -> Result<()> {
    let (source, destination) = cross_copy_fixture()?;
    let archive_bytes = plan_for(&source, &destination)?
        .retained_candidate_bytes()
        .expect("the default budget retains this candidate");
    let wide = Limits::default();
    let tight = limits_retaining(archive_bytes - 1)?;

    assert_eq!(
        plan_under(&source, &destination, tight, wide)?.retained_candidate_bytes(),
        None,
        "a tight source budget binds the operation"
    );
    assert_eq!(
        plan_under(&source, &destination, wide, tight)?.retained_candidate_bytes(),
        None,
        "a tight destination budget binds the operation"
    );
    assert_eq!(
        plan_under(&source, &destination, wide, wide)?.retained_candidate_bytes(),
        Some(archive_bytes)
    );
    Ok(())
}

/// Releasing the archive returns the bytes, keeps the plan's value, and leaves
/// it applicable through the rebuild; dropping the plan returns them too.
#[test]
fn releasing_and_dropping_return_the_retained_bytes() -> Result<()> {
    let (source, mut destination) = cross_copy_fixture()?;
    let mut plan = plan_for(&source, &destination)?;
    let archive = Arc::clone(
        &plan
            .candidate
            .0
            .as_ref()
            .expect("the default budget retains this candidate")
            .archive,
    );
    assert_eq!(
        Arc::strong_count(&archive),
        2,
        "planning drops the candidate package, leaving the plan the only owner"
    );

    let released = {
        let mut released = plan.clone();
        released.release_retained_candidate();
        assert_eq!(released.retained_candidate_bytes(), None);
        assert_eq!(
            released, plan,
            "releasing the archive preserves the plan's value"
        );
        released
    };
    assert_eq!(
        Arc::strong_count(&archive),
        2,
        "the released clone returned its handle"
    );

    // A released plan still applies, through the rebuild.
    let mut rebuilt = Package::from_vec(destination.to_bytes()?)?;
    rebuilt.apply_cross_slide_copy_plan(&source, &released)?;
    let mut reused = Package::from_vec(destination.to_bytes()?)?;
    reused.apply_cross_slide_copy_plan(&source, &plan)?;
    assert_eq!(rebuilt.to_bytes()?, reused.to_bytes()?);

    // Applying does not release: a plan may be applied again.
    assert_eq!(plan.retained_candidate_bytes(), Some(archive.len()));
    plan.release_retained_candidate();
    assert_eq!(plan.retained_candidate_bytes(), None);
    drop(plan);
    assert_eq!(
        Arc::strong_count(&archive),
        1,
        "dropping the plan returns the retained bytes"
    );
    Ok(())
}

/// `Debug` reports the retained length, never the retained bytes.
#[test]
fn debug_does_not_print_the_retained_archive() -> Result<()> {
    let (source, destination) = cross_copy_fixture()?;
    let plan = plan_for(&source, &destination)?;
    let rendered = format!("{plan:?}");
    let archive_bytes = plan
        .retained_candidate_bytes()
        .expect("the default budget retains this candidate");
    assert!(rendered.contains(&format!("retained_bytes: Some({archive_bytes})")));
    // A rendered archive would begin with the ZIP local-file signature
    // `PK\x03\x04` as the first four bytes of a `Vec<u8>` Debug. This is a
    // structural check rather than a size comparison, so no fixture can make
    // it pass vacuously.
    assert!(
        !rendered.contains("[80, 75, 3, 4"),
        "a plan's Debug rendering must not carry the archive"
    );
    Ok(())
}

/// Substituting the retained archive is refused before publication and leaves
/// the destination byte-identical.  This is the test that fails if a retained
/// archive were ever trusted without a recomputed proof.
///
/// Two layers catch it and this fixture pins both.  In a debug or test build
/// the `debug_assert` in `build_candidate` re-serializes the candidate and
/// panics; in a release build that assertion is compiled out and the refusal
/// comes from the recomputed complete-package revision, which is the layer the
/// design rests on.  The release half runs under `cargo test --release`.
#[test]
#[cfg(debug_assertions)]
#[should_panic(expected = "cross-slide reused a retained candidate archive")]
fn a_substituted_retained_archive_trips_the_debug_assertion() {
    substitute_and_apply().expect("fixture");
}

#[test]
#[cfg(not(debug_assertions))]
fn a_substituted_retained_archive_is_refused_and_publishes_nothing() -> Result<()> {
    substitute_and_apply()
}

fn substitute_and_apply() -> Result<()> {
    let (source, mut destination) = cross_copy_fixture()?;
    let mut plan = plan_for(&source, &destination)?;

    // A different, structurally valid PPTX archive: the destination's own.
    let foreign = destination.to_bytes()?;
    plan.candidate
        .0
        .as_mut()
        .expect("the default budget retains this candidate")
        .archive = Arc::new(foreign);

    let mut applied = Package::from_vec(destination.to_bytes()?)?;
    let before = applied.to_bytes()?;
    let outcome = applied.apply_cross_slide_copy_plan(&source, &plan);
    let rendered = format!("{outcome:?}");
    assert!(
        outcome.is_err(),
        "a substituted candidate archive must be refused, got {rendered}"
    );
    assert!(
        rendered.contains("cross-slide candidate did not publish the reserved slide identity"),
        "the refusal that catches a substituted archive in a release build moved: {rendered}"
    );
    assert_eq!(
        applied.to_bytes()?,
        before,
        "a refused application leaves the destination byte-identical"
    );
    Ok(())
}

/// Reuse moves no verdict: the same application with and without the retained
/// archive publishes the same bytes on the accepting route and produces the
/// identical `Result` on each rejecting route.
#[test]
fn retention_changes_no_apply_verdict() -> Result<()> {
    let (mut source, mut destination) = cross_copy_fixture()?;
    let retained = plan_for(&source, &destination)?;
    let mut unretained = retained.clone();
    unretained.release_retained_candidate();
    let limits = Limits::default();
    let destination_bytes = destination.to_bytes()?;

    // Accepting route: identical published bytes and identical revisions.
    let mut with = Package::from_vec(destination_bytes.clone())?;
    let mut without = Package::from_vec(destination_bytes.clone())?;
    let with_snapshot = with.apply_cross_slide_copy_plan(&source, &retained)?;
    let without_snapshot = without.apply_cross_slide_copy_plan(&source, &unretained)?;
    assert_eq!(with.to_bytes()?, without.to_bytes()?);
    assert_eq!(with_snapshot.revision(), without_snapshot.revision());
    assert_eq!(
        physical_package_fingerprint(&with.opc, limits)?,
        retained.target_physical_revision()
    );

    // Rejecting route 1: a destination whose slide XML drifted.
    let drift = |plan: &CrossSlideCopyPlan| -> Result<String> {
        let mut drifted = Package::from_vec(destination_bytes.clone())?;
        let slide = drifted
            .opened_presentation()?
            .slides()
            .first()
            .expect("destination has slides")
            .part_name()
            .clone();
        let stale = String::from_utf8_lossy(drifted.opc.get_part(&slide)?.blob())
            .replace("retain-destination-a", "retain-destination-z")
            .into_bytes();
        drifted.opc.get_part_mut(&slide)?.set_blob(stale);
        Ok(format!(
            "{:?}",
            drifted.apply_cross_slide_copy_plan(&source, plan)
        ))
    };
    assert_eq!(drift(&retained)?, drift(&unretained)?);
    assert!(drift(&retained)?.contains("UnsafeEdit"));

    // Rejecting route 2: a foreign source package.
    let foreign_source = owned_slides(&["foreign"])?;
    let foreign = |plan: &CrossSlideCopyPlan| -> Result<String> {
        let mut target = Package::from_vec(destination_bytes.clone())?;
        Ok(format!(
            "{:?}",
            target.apply_cross_slide_copy_plan(&foreign_source, plan)
        ))
    };
    assert_eq!(foreign(&retained)?, foreign(&unretained)?);
    assert!(foreign(&retained)?.contains("UnsafeEdit"));

    // Rejecting route 3: a source whose own bytes drifted after planning.
    let mut stale_source = |plan: &CrossSlideCopyPlan| -> Result<String> {
        let mut drifted = Package::from_vec(source.to_bytes()?)?;
        let slide = drifted
            .opened_presentation()?
            .slides()
            .first()
            .expect("source has slides")
            .part_name()
            .clone();
        let stale = String::from_utf8_lossy(drifted.opc.get_part(&slide)?.blob())
            .replace("retain-source", "retain-source-z")
            .into_bytes();
        drifted.opc.get_part_mut(&slide)?.set_blob(stale);
        let mut target = Package::from_vec(destination_bytes.clone())?;
        Ok(format!(
            "{:?}",
            target.apply_cross_slide_copy_plan(&drifted, plan)
        ))
    };
    assert_eq!(stale_source(&retained)?, stale_source(&unretained)?);
    assert!(stale_source(&retained)?.contains("UnsafeEdit"));

    // Rejecting route 4: borrowed graph-only ingress carries no physical
    // provenance and is refused before any candidate is built.
    let borrowed = |plan: &CrossSlideCopyPlan| -> Result<String> {
        let mut target = Package::from_bytes(&destination_bytes)?;
        Ok(format!(
            "{:?}",
            target.apply_cross_slide_copy_plan(&source, plan)
        ))
    };
    assert_eq!(borrowed(&retained)?, borrowed(&unretained)?);
    assert!(borrowed(&retained)?.contains("SlideCopyPlan"));
    Ok(())
}

/// The durable patch carries no archive, and a patch-driven application
/// retains nothing: the encoded bytes and the published result are identical
/// whether or not the plan that produced the patch retained its candidate.
#[test]
fn the_durable_patch_is_untouched_by_retention() -> Result<()> {
    let (source, mut destination) = cross_copy_fixture()?;
    let retained = plan_for(&source, &destination)?;
    let mut unretained = retained.clone();
    unretained.release_retained_candidate();
    assert_eq!(
        retained.patch().to_bytes()?,
        unretained.patch().to_bytes()?,
        "the encoded durable patch does not depend on retention"
    );
    // The family prefix, not the version: change 0655 bumps the version in the
    // same wave and retention has nothing to do with it.
    assert!(retained.patch().to_bytes()?.starts_with(b"LPCP"));

    let destination_bytes = destination.to_bytes()?;
    let mut from_patch = Package::from_vec(destination_bytes.clone())?;
    from_patch.apply_cross_slide_copy_patch(&source, retained.patch())?;
    let mut from_plan = Package::from_vec(destination_bytes)?;
    from_plan.apply_cross_slide_copy_plan(&source, &retained)?;
    assert_eq!(from_patch.to_bytes()?, from_plan.to_bytes()?);
    Ok(())
}

/// Source ratchet: nothing on the retention path may reach the filesystem, a
/// temporary file or a scratch provider.  ADR 0005 forbids litchi from
/// spilling document content automatically, and a candidate archive is a
/// complete presentation package.
#[test]
fn the_retention_path_never_spills() {
    const FORBIDDEN: &[&str] = &[
        "std::fs",
        "std :: fs",
        "fs::File",
        "File::create",
        "File::open",
        "tempfile",
        "TempDir",
        "NamedTempFile",
        "temp_dir",
        "scratch",
        "Scratch",
        "mmap",
        "MmapMut",
        "OpenOptions",
    ];
    let source = include_str!("../cross_copy_plan.rs");
    for marker in FORBIDDEN {
        assert!(
            !source.contains(marker),
            "the cross-slide copy plan must not reach {marker}: a retained candidate archive is \
             document content and is never spilled"
        );
    }
}

/// Planning against a destination whose exact-source authorization has been
/// revoked, then applying the plan to a clean reopen of that destination's own
/// serialized bytes.
///
/// This is the one pairing in which the planning and applying destinations
/// have *different* preservation sources while passing all four staleness
/// proofs: the plan-time destination republishes its pristine members out of
/// the archive it was opened from, and the apply-time one out of the archive
/// that republication produced.  It is therefore the pairing in which a
/// retained candidate archive could publish bytes a rebuild would not.
#[test]
fn a_dirty_planning_destination_and_a_clean_applying_one_agree() -> Result<()> {
    let (source, mut destination) = cross_copy_fixture()?;
    let mut dirty = Package::from_vec(destination.to_bytes()?)?;
    {
        // An in-place part edit revokes exact-source authorization while owned
        // ingress and physical provenance survive, which is the state this
        // pairing needs. Slide names stay distinct so the copy is admissible.
        let slide = dirty
            .opened_presentation()?
            .slides()
            .get(1)
            .expect("destination has two slides")
            .part_name()
            .clone();
        let renamed = String::from_utf8_lossy(dirty.opc.get_part(&slide)?.blob())
            .replace("retain-destination-b", "retain-destination-c")
            .into_bytes();
        dirty.opc.get_part_mut(&slide)?.set_blob(renamed);
    }
    assert!(
        !dirty.opc.is_unmodified_owned_source(),
        "an in-place part edit revokes exact-source authorization"
    );

    let retained = {
        let source_snapshot = source.opened_presentation()?;
        let destination_snapshot = dirty.opened_presentation()?;
        destination_snapshot.plan_cross_slide_copy(&source_snapshot, 0_usize, 0_usize, 1)?
    };
    assert!(
        retained.retained_candidate_bytes().is_some(),
        "a dirty-but-owned destination still retains: the candidate reopen \
         authorizes its own bytes whatever the destination's state"
    );
    let mut released = retained.clone();
    released.release_retained_candidate();

    let clean_bytes = dirty.to_bytes()?;
    let mut with = Package::from_vec(clean_bytes.clone())?;
    let mut without = Package::from_vec(clean_bytes)?;
    // In a debug build this call runs the `debug_assert` that re-serializes
    // the apply-time candidate graph and compares it with the retained bytes,
    // so a preservation-source difference that reached the published archive
    // would panic here rather than being argued away.
    let reused = format!("{:?}", with.apply_cross_slide_copy_plan(&source, &retained));
    let rebuilt = format!(
        "{:?}",
        without.apply_cross_slide_copy_plan(&source, &released)
    );
    assert!(
        reused.starts_with("Ok("),
        "the pairing must be admissible: {reused}"
    );
    assert_eq!(reused, rebuilt, "the verdict must not depend on retention");
    assert_eq!(
        with.to_bytes()?,
        without.to_bytes()?,
        "the published bytes must not depend on retention"
    );
    Ok(())
}
