#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "revision cache tests use panic-on-fixture-failure assertions"
)]

use super::{
    candidate_physical_revision, physical_package_fingerprint, remember_physical_revision,
    snapshot_physical_revision,
};
use crate::opened::model::{Limits, Snapshot};
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
        .ok_or(Error::SlideIndexOutOfBounds { index, len: index })?;
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

/// Append a ZIP end-of-central-directory comment. The package graph is
/// unchanged, so the semantic fingerprint is identical while the retained
/// archive — and therefore the physical revision — is not.
fn with_eocd_comment(mut archive: Vec<u8>, comment: &[u8]) -> Result<Vec<u8>> {
    let comment_len = u16::try_from(comment.len())
        .map_err(|_error| Error::Invalid("test ZIP comment exceeds u16".into()))?;
    let eocd = archive
        .len()
        .checked_sub(22)
        .ok_or_else(|| Error::Invalid("test ZIP has no EOCD".into()))?;
    if archive.get(eocd..eocd + 4) != Some(b"PK\x05\x06") {
        return Err(Error::Invalid("test ZIP has an invalid EOCD".into()));
    }
    archive[eocd + 20..eocd + 22].copy_from_slice(&comment_len.to_le_bytes());
    archive.extend_from_slice(comment);
    Ok(archive)
}

fn cross_copy_fixture() -> Result<(Package, Package)> {
    let mut source = owned_slides(&["cache-source"])?;
    rename_slide(&mut source, 0, "cache-source")?;
    let source = Package::from_vec(source.to_bytes()?)?;

    let mut destination = owned_slides(&["cache-destination-a", "cache-destination-b"])?;
    rename_slide(&mut destination, 0, "cache-destination-a")?;
    rename_slide(&mut destination, 1, "cache-destination-b")?;
    let destination = Package::from_vec(destination.to_bytes()?)?;
    Ok((source, destination))
}

fn cached(snapshot: &Snapshot) -> Option<(usize, [u8; 32])> {
    snapshot.physical_revision.get().copied()
}

#[test]
fn snapshot_physical_revision_caches_the_freshly_computed_value() -> Result<()> {
    let (source, _destination) = cross_copy_fixture()?;
    let snapshot = source.opened_presentation()?;
    let limits = snapshot.limits();

    assert!(cached(&snapshot).is_none());
    let fresh = physical_package_fingerprint(snapshot.package.as_ref(), limits)?;
    let first = snapshot_physical_revision(&snapshot, limits)?;
    assert_eq!(first, fresh);
    assert_eq!(cached(&snapshot), Some((limits.max_patch_bytes(), fresh)));

    // The cached read must return the value the recomputation returns.
    let second = snapshot_physical_revision(&snapshot, limits)?;
    assert_eq!(second, fresh);
    assert_eq!(
        physical_package_fingerprint(snapshot.package.as_ref(), limits)?,
        fresh
    );

    // A snapshot clone shares the immutable package and the same cache.
    let clone = snapshot.clone();
    assert_eq!(cached(&clone), Some((limits.max_patch_bytes(), fresh)));
    assert_eq!(snapshot_physical_revision(&clone, limits)?, fresh);
    Ok(())
}

#[test]
fn snapshot_physical_revision_is_keyed_by_the_archive_bound() -> Result<()> {
    let (source, _destination) = cross_copy_fixture()?;
    let snapshot = source.opened_presentation()?;
    let limits = snapshot.limits();
    let fresh = snapshot_physical_revision(&snapshot, limits)?;

    // A different bound is a cache miss and recomputes the same digest: the
    // value does not depend on the bound, only the refusal does.
    let wider = Limits::new(
        limits.max_parts(),
        limits.max_patch_bytes() - 1,
        limits.max_text_bytes(),
        limits.max_history_entries(),
        limits.max_history_bytes(),
    )
    .ok_or_else(|| Error::Invalid("test limits are invalid".into()))?;
    assert_eq!(snapshot_physical_revision(&snapshot, wider)?, fresh);
    assert_eq!(cached(&snapshot), Some((limits.max_patch_bytes(), fresh)));

    // A bound below the serialized archive still refuses, cache or no cache.
    let tiny = Limits::new(
        limits.max_parts(),
        1,
        limits.max_text_bytes(),
        limits.max_history_entries(),
        limits.max_history_bytes(),
    )
    .ok_or_else(|| Error::Invalid("test limits are invalid".into()))?;
    assert!(matches!(
        snapshot_physical_revision(&snapshot, tiny),
        Err(Error::Limit {
            resource: "cross-slide serialized archive bytes",
            ..
        })
    ));
    assert_eq!(snapshot_physical_revision(&snapshot, limits)?, fresh);
    Ok(())
}

#[test]
fn rebound_snapshot_starts_an_empty_physical_revision_cache() -> Result<()> {
    let (mut source, _destination) = cross_copy_fixture()?;
    let snapshot = source.opened_presentation()?;
    let limits = snapshot.limits();
    let original = snapshot_physical_revision(&snapshot, limits)?;
    assert!(cached(&snapshot).is_some());

    // Same graph, different retained archive: `packages_equal` holds and the
    // semantic revision is identical, but the serialized archive is not.
    let commented = Package::from_vec(with_eocd_comment(
        source.to_bytes()?,
        b"revision-cache-rebind",
    )?)?;
    let commented_snapshot = commented.opened_presentation()?;
    assert_eq!(commented_snapshot.revision(), snapshot.revision());
    assert!(crate::opened::model::packages_equal(
        snapshot.package.as_ref(),
        commented_snapshot.package.as_ref()
    ));

    let rebound = snapshot.rebound_to(commented_snapshot.package.as_ref());
    assert!(cached(&rebound).is_none());
    let rebound_revision = snapshot_physical_revision(&rebound, limits)?;
    assert_eq!(
        rebound_revision,
        physical_package_fingerprint(commented_snapshot.package.as_ref(), limits)?
    );
    assert_ne!(rebound_revision, original);
    Ok(())
}

#[test]
fn planned_and_published_revisions_equal_fresh_serializations() -> Result<()> {
    let (source, mut destination) = cross_copy_fixture()?;
    let source_snapshot = source.opened_presentation()?;
    let destination_snapshot = destination.opened_presentation()?;
    let limits = destination_snapshot.limits();
    let plan = destination_snapshot.plan_cross_slide_copy(&source_snapshot, 0, 1, 1)?;

    assert_eq!(
        plan.source_physical_revision(),
        physical_package_fingerprint(&source.opc, limits)?
    );
    assert_eq!(
        plan.destination_physical_revision(),
        physical_package_fingerprint(&destination.opc, limits)?
    );

    let published = destination.apply_cross_slide_copy_plan(&source, &plan)?;
    assert_eq!(published.slides().len(), 3);
    // The candidate archive hashed while it was built is exactly what a fresh
    // serializing hash sink reads back from the published package.
    assert_eq!(
        plan.target_physical_revision(),
        physical_package_fingerprint(&destination.opc, limits)?
    );
    assert_eq!(
        plan.target_revision(),
        crate::opened::model::package_fingerprint(&destination.opc)?
    );
    Ok(())
}

#[test]
fn warm_snapshot_caches_do_not_hide_a_stale_source_or_destination() -> Result<()> {
    let (mut source, mut destination) = cross_copy_fixture()?;
    let source_bytes = source.to_bytes()?;
    let destination_bytes = destination.to_bytes()?;
    let source = Package::from_vec(source_bytes.clone())?;
    let destination = Package::from_vec(destination_bytes.clone())?;
    let source_snapshot = source.opened_presentation()?;
    let destination_snapshot = destination.opened_presentation()?;
    let limits = destination_snapshot.limits();
    let plan = destination_snapshot.plan_cross_slide_copy(&source_snapshot, 0, 1, 1)?;
    // Warm both caches by planning again from the very same snapshots.
    let replanned = destination_snapshot.plan_cross_slide_copy(&source_snapshot, 0, 1, 1)?;
    assert_eq!(
        replanned.source_physical_revision(),
        plan.source_physical_revision()
    );
    assert_eq!(
        replanned.destination_physical_revision(),
        plan.destination_physical_revision()
    );
    assert_eq!(
        replanned.target_physical_revision(),
        plan.target_physical_revision()
    );
    assert!(cached(&source_snapshot).is_some());
    assert!(cached(&destination_snapshot).is_some());

    // The live destination moves on after planning. The snapshot caches cover
    // the immutable captured packages, never a live one, so the apply-time
    // proof still runs and still refuses.
    let mut drifted = Package::from_vec(destination_bytes.clone())?;
    let slide = drifted.opened_presentation()?.slides()[0]
        .part_name()
        .clone();
    let stale = std::str::from_utf8(drifted.opc.get_part(&slide)?.blob())
        .map_err(|error| Error::Xml(error.to_string()))?
        .replace("cache-destination-a", "cache-destination-x")
        .into_bytes();
    drifted.opc.get_part_mut(&slide)?.set_blob(stale);
    let drifted_before = drifted.opc.clone();
    assert!(matches!(
        drifted.apply_cross_slide_copy_plan(&source, &plan),
        Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the complete destination package graph changed after cross-slide planning",
        })
    ));
    assert_eq!(
        physical_package_fingerprint(&drifted.opc, limits)?,
        physical_package_fingerprint(&drifted_before, limits)?
    );

    // A source whose graph is identical but whose retained archive is not
    // reaches the physical proof, which still refuses it.
    let mut repackaged_destination = Package::from_vec(destination_bytes.clone())?;
    let recomment = Package::from_vec(with_eocd_comment(
        source_bytes.clone(),
        b"revision-cache-foreign",
    )?)?;
    assert_eq!(
        recomment.opened_presentation()?.revision(),
        source_snapshot.revision()
    );
    assert!(matches!(
        repackaged_destination.apply_cross_slide_copy_plan(&recomment, &plan),
        Err(Error::UnsafeEdit {
            operation: "apply_cross_slide_copy_plan",
            reason: "the serialized source package changed after cross-slide planning",
        })
    ));

    // The unchanged pairing still applies, and publishes the planned archive.
    let mut applied = Package::from_vec(destination_bytes)?;
    let published = applied.apply_cross_slide_copy_plan(&source, &plan)?;
    assert_eq!(published.slides().len(), 3);
    assert_eq!(
        physical_package_fingerprint(&applied.opc, limits)?,
        plan.target_physical_revision()
    );
    Ok(())
}

#[test]
fn candidate_physical_revision_without_a_known_value_recomputes() -> Result<()> {
    let (source, _destination) = cross_copy_fixture()?;
    let snapshot = source.opened_presentation()?;
    let limits = snapshot.limits();
    let fresh = physical_package_fingerprint(&source.opc, limits)?;
    assert_eq!(
        candidate_physical_revision(&source.opc, limits, None)?,
        fresh
    );
    assert_eq!(
        candidate_physical_revision(&source.opc, limits, Some(fresh))?,
        fresh
    );

    // Seeding is the same value the snapshot would compute for itself.
    remember_physical_revision(&snapshot, limits, fresh);
    assert_eq!(cached(&snapshot), Some((limits.max_patch_bytes(), fresh)));
    assert_eq!(snapshot_physical_revision(&snapshot, limits)?, fresh);
    Ok(())
}
