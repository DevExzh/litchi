//! Focused tests for the proved slide-root classification memo (ADR 0032;
//! change 0760, re-applying change 0743's withdrawn `99ce9c5e34`).
//!
//! An opened capture classifies every slide root with the complete notes-graph
//! scan. The snapshot keeps those classifications keyed by the payload
//! allocation each one read, and a commit's recapture reuses them for every
//! slide the transaction did not rewrite. These tests pin that the reuse is
//! real (the scan is skipped), that it is exact (a memo-assisted capture
//! equals a cold one, value and refusal alike), that it can never answer for
//! bytes it did not classify, and every condition ADR 0032 and ADR 0005's
//! 2026-09-16 memo amendment place on it:
//!
//! * keyed on allocation identity with a strong reference;
//! * entries only for allocations the owning snapshot's package holds;
//! * projected, never inherited, on a rebind;
//! * a miss is an ordinary recomputation, and no value, refusal or published
//!   byte depends on hits;
//! * built or projected, never mutated;
//! * resident cost bounded and charged through fallible reservation, whose
//!   refusal is a typed error;
//! * every hit re-derived in test and debug builds.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused memo tests use panic-on-fixture-failure assertions"
)]

use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use litchi_opc::{BlobPart, OpcPackage, PackURI, Part, Relationships, TargetMode};

use super::{Limits, Snapshot};
use crate::notes::{
    Conformance, SlideRootMemo, SlideRootRecord, with_refused_slide_root_reservation,
};
use crate::parts::{take_proved_root_hits, take_root_scans};
use crate::{Error, Package, Result};

/// Resource named by a refused memo reservation (ADR 0032 section 4).
const MEMO_RESOURCE: &str = "opened-presentation slide-root memo";

fn text_box_package(slides: usize, boxes: usize) -> Result<Package> {
    let mut package = Package::new()?;
    let presentation = package.presentation_mut()?;
    for slide_index in 0..slides {
        let slide = presentation.add_slide()?;
        for box_index in 0..boxes {
            let offset = i64::try_from(box_index).map_err(|_| Error::Invalid("offset".into()))?;
            slide.add_text_box(
                &format!("slide {slide_index} box {box_index}"),
                36,
                36 + offset * 90,
                144,
                54,
            );
        }
    }
    Package::from_vec(package.to_bytes()?)
}

/// Recapture `candidate` the way a commit does, with or without the source
/// snapshot's proved slide-root classifications.
fn recapture(source: &Snapshot, candidate: &OpcPackage, memo: bool) -> Result<Snapshot> {
    let (revision, digests) =
        super::model::package_fingerprint_with_memo(candidate, Some(&source.part_digests))?;
    super::model::capture_with_revision_and_digests_and_mce(
        candidate,
        source.limits,
        source.physical_source_provenance,
        revision,
        digests,
        source.retained_mce.as_deref(),
        memo.then(|| source.slide_roots.as_ref()),
    )
}

fn cold_capture(snapshot: &Snapshot) -> Result<Snapshot> {
    super::model::capture_with_provenance(
        snapshot.package.as_ref(),
        snapshot.limits,
        snapshot.physical_source_provenance,
    )
}

fn assert_same_snapshot(label: &str, left: &Snapshot, right: &Snapshot) {
    assert_eq!(left.revision(), right.revision(), "{label}: revision");
    assert_eq!(left.slides(), right.slides(), "{label}: slides");
    assert_eq!(
        left.retained_mce_bytes(),
        right.retained_mce_bytes(),
        "{label}: retained MCE"
    );
    assert_eq!(
        left.slide_roots.len(),
        right.slide_roots.len(),
        "{label}: proved roots"
    );
}

fn assert_same_result(label: &str, left: &Result<Snapshot>, right: &Result<Snapshot>) {
    match (left, right) {
        (Ok(left), Ok(right)) => assert_same_snapshot(label, left, right),
        (Err(left), Err(right)) => {
            assert_eq!(
                format!("{left:?}"),
                format!("{right:?}"),
                "{label}: refusal"
            );
        },
        (left, right) => panic!("{label}: memo changed the outcome: {left:?} versus {right:?}"),
    }
}

fn assert_memo_refusal<T: std::fmt::Debug>(label: &str, result: Result<T>) {
    match result {
        Err(Error::Allocation { resource, .. }) => assert_eq!(resource, MEMO_RESOURCE, "{label}"),
        other => panic!("{label}: expected the typed memo reservation refusal, got {other:?}"),
    }
}

fn slide_blob(snapshot: &Snapshot, index: usize) -> &[u8] {
    snapshot
        .package
        .get_part(&snapshot.slides[index].part_name)
        .expect("captured slide part")
        .blob()
}

fn slide_arc(snapshot: &Snapshot, index: usize) -> Arc<Vec<u8>> {
    snapshot
        .package
        .get_part(&snapshot.slides[index].part_name)
        .expect("captured slide part")
        .blob_arc()
}

fn reallocate_slide(package: &mut OpcPackage, part_name: &PackURI) -> Result<()> {
    let copy = package.get_part(part_name)?.blob().to_vec();
    package.get_part_mut(part_name)?.set_blob(copy);
    Ok(())
}

/// Every entry of `snapshot`'s memo names, and retains, an allocation one of
/// its own package's parts holds; returns the number of entries checked.
fn assert_entries_owned_by(label: &str, snapshot: &Snapshot) -> usize {
    let package = snapshot.package.as_ref();
    let held: Vec<Arc<Vec<u8>>> = package
        .iter_parts()
        .map(|metadata| {
            package
                .get_part(metadata.partname())
                .expect("listed part")
                .blob_arc()
        })
        .collect();
    let mut checked = 0;
    for (key, raw, _conformance) in snapshot.slide_roots.retained() {
        assert_eq!(
            key,
            (raw.as_ptr() as usize, raw.len()),
            "{label}: an entry's key names the allocation it retains"
        );
        assert!(
            held.iter().any(|blob| Arc::ptr_eq(blob, raw)),
            "{label}: an entry names an allocation its snapshot's package does not hold"
        );
        checked += 1;
    }
    checked
}

/// The observable content of a memo: each entry's key, retained allocation
/// address and classification, in table order.
fn memo_content(memo: &SlideRootMemo) -> Vec<((usize, usize), usize, Conformance)> {
    memo.retained()
        .map(|(key, raw, conformance)| (key, Arc::as_ptr(raw) as usize, conformance))
        .collect()
}

#[test]
fn a_capture_proves_one_classification_per_slide_it_scanned() -> Result<()> {
    let package = text_box_package(5, 3)?;
    let snapshot = package.opened_presentation()?;
    assert_eq!(snapshot.slide_roots.len(), 5);
    for index in 0..5 {
        assert_eq!(
            snapshot.slide_roots.lookup(slide_blob(&snapshot, index)),
            Some(Conformance::Transitional),
            "slide {index} must be classified by the exact bytes the snapshot holds"
        );
    }
    Ok(())
}

#[test]
fn a_one_slide_commit_rescans_only_the_rewritten_slide_and_equals_a_cold_capture() -> Result<()> {
    let package = text_box_package(5, 3)?;
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    assert!(edit.set_shape_text(2, 0, "rewritten")?);
    take_proved_root_hits();
    let commit = edit.commit()?;
    assert_eq!(
        take_proved_root_hits(),
        4,
        "the four untouched slides reuse their classification"
    );
    let cold = cold_capture(commit.snapshot())?;
    assert_same_snapshot("commit versus cold capture", commit.snapshot(), &cold);
    // The rewritten slide is classified afresh, so the committed memo still
    // covers every slide of the committed package.
    for index in 0..5 {
        assert!(
            commit
                .snapshot()
                .slide_roots
                .lookup(slide_blob(commit.snapshot(), index))
                .is_some()
        );
    }
    Ok(())
}

#[test]
fn a_no_op_commit_and_an_all_slide_commit_keep_their_exact_meaning() -> Result<()> {
    let package = text_box_package(4, 2)?;
    let source = package.opened_presentation()?;
    take_proved_root_hits();
    let no_op = source.edit().commit()?;
    assert!(!no_op.is_changed());
    assert_eq!(no_op.snapshot().revision(), source.revision());
    assert!(
        Arc::ptr_eq(&no_op.snapshot().slide_roots, &source.slide_roots),
        "an exact no-op returns the source snapshot, memo included"
    );
    assert_eq!(
        take_proved_root_hits(),
        0,
        "a no-op commit captures nothing"
    );

    let mut edit = source.edit();
    for slide in 0..4 {
        assert!(edit.set_shape_text(slide, 1, format!("all {slide}"))?);
    }
    take_proved_root_hits();
    let commit = edit.commit()?;
    assert_eq!(take_proved_root_hits(), 0, "every slide was rewritten");
    let cold = cold_capture(commit.snapshot())?;
    assert_same_snapshot("all-slide commit", commit.snapshot(), &cold);
    Ok(())
}

#[test]
fn a_new_allocation_with_equal_bytes_is_a_miss_not_a_hit() -> Result<()> {
    let package = text_box_package(5, 2)?;
    let source = package.opened_presentation()?;
    let mut candidate = source.package.as_ref().clone();
    reallocate_slide(&mut candidate, &source.slides[0].part_name)?;
    take_proved_root_hits();
    let assisted = recapture(&source, &candidate, true)?;
    assert_eq!(take_proved_root_hits(), 4);
    let plain = recapture(&source, &candidate, false)?;
    assert_eq!(take_proved_root_hits(), 0);
    assert_same_snapshot("re-allocated slide", &assisted, &plain);
    // Equal bytes are not an identity: the source memo does not answer for the
    // copy, although it answers for the original.
    let copy = candidate.get_part(&source.slides[0].part_name)?.blob();
    assert_eq!(copy, slide_blob(&source, 0));
    assert_eq!(source.slide_roots.lookup(copy), None);
    assert!(source.slide_roots.lookup(slide_blob(&source, 0)).is_some());
    Ok(())
}

fn replace_first_text(package: &mut OpcPackage, part_name: &PackURI, with: &str) -> Result<()> {
    let xml = std::str::from_utf8(package.get_part(part_name)?.blob())
        .map_err(|error| Error::Xml(error.to_string()))?
        .to_owned();
    let start = xml.find("<a:t>").expect("slide has a text run") + "<a:t>".len();
    let end = start + xml[start..].find("</a:t>").expect("text run ends");
    let mut replaced = String::with_capacity(xml.len() + with.len());
    replaced.push_str(&xml[..start]);
    replaced.push_str(with);
    replaced.push_str(&xml[end..]);
    package
        .get_part_mut(part_name)?
        .set_blob(replaced.into_bytes());
    Ok(())
}

#[test]
fn the_memo_never_turns_a_refusal_into_a_success() -> Result<()> {
    let package = text_box_package(5, 2)?;
    let source = package.opened_presentation()?;
    // CDATA passes the slide's own root and name projections but fails the
    // complete notes-root scan, so the capture refuses at the notes graph.
    for (label, index, text) in [
        ("first slide CDATA", 0, "<![CDATA[first]]>"),
        ("middle slide CDATA", 2, "<![CDATA[middle]]>"),
        ("last slide unbound prefix", 4, "<q:x/>"),
    ] {
        let mut candidate = source.package.as_ref().clone();
        replace_first_text(&mut candidate, &source.slides[index].part_name, text)?;
        let assisted = recapture(&source, &candidate, true);
        let plain = recapture(&source, &candidate, false);
        let cold = super::model::capture_with_provenance(
            &candidate,
            source.limits,
            source.physical_source_provenance,
        );
        assert!(
            assisted.is_err(),
            "{label}: the malformed slide must refuse"
        );
        assert_same_result(label, &assisted, &plain);
        assert_same_result(label, &assisted, &cold);
    }
    Ok(())
}

#[test]
fn a_published_snapshot_keeps_classifications_only_for_its_own_allocations() -> Result<()> {
    let mut package = text_box_package(4, 2)?;
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    assert!(edit.set_shape_text(1, 0, "published")?);
    let commit = edit.commit()?;
    let published = package.apply_opened_presentation_commit(commit)?;
    assert_eq!(published.slide_roots.len(), 4);
    for index in 0..4 {
        assert!(
            published
                .slide_roots
                .lookup(slide_blob(&published, index))
                .is_some()
        );
    }
    assert_eq!(assert_entries_owned_by("published", &published), 4);
    // A rebind onto a package whose slide allocation moved keeps no entry for
    // the moved slide: equal bytes in a new allocation are not an identity.
    let mut moved = published.package.as_ref().clone();
    reallocate_slide(&mut moved, &published.slides[0].part_name)?;
    let rebound = published.rebound_to(&moved);
    assert_eq!(rebound.slide_roots.len(), 3);
    assert!(
        rebound
            .slide_roots
            .lookup(slide_blob(&rebound, 0))
            .is_none()
    );
    assert_eq!(assert_entries_owned_by("rebound", &rebound), 3);
    Ok(())
}

/// A foreign part whose visible payload and returned `Arc` disagree.
#[derive(Clone, Debug)]
struct MismatchedBlobArcPart {
    inner: BlobPart,
    foreign: Arc<Vec<u8>>,
}

/// A part that counts every raw payload observation.
#[derive(Clone, Debug)]
struct CountingBlobPart {
    inner: BlobPart,
    observations: Arc<AtomicUsize>,
}

fn copied_part(part: &dyn Part) -> Result<BlobPart> {
    let mut inner = BlobPart::new(
        part.partname().clone(),
        part.content_type().to_owned(),
        part.blob().to_vec(),
    );
    for relationship in part.rels().iter() {
        inner.rels_mut().try_add_relationship(
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
    Ok(inner)
}

macro_rules! delegate_part {
    ($type:ty, $blob:expr, $blob_arc:expr) => {
        impl Part for $type {
            fn blob(&self) -> &[u8] {
                $blob(self)
            }

            fn blob_arc(&self) -> Arc<Vec<u8>> {
                $blob_arc(self)
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
    };
}

fn mismatched_blob(part: &MismatchedBlobArcPart) -> &[u8] {
    part.inner.blob()
}

fn mismatched_blob_arc(part: &MismatchedBlobArcPart) -> Arc<Vec<u8>> {
    Arc::clone(&part.foreign)
}

delegate_part!(MismatchedBlobArcPart, mismatched_blob, mismatched_blob_arc);

fn counted_blob(part: &CountingBlobPart) -> &[u8] {
    part.observations.fetch_add(1, Ordering::SeqCst);
    part.inner.blob()
}

fn counted_blob_arc(part: &CountingBlobPart) -> Arc<Vec<u8>> {
    part.observations.fetch_add(1, Ordering::SeqCst);
    part.inner.blob_arc()
}

delegate_part!(CountingBlobPart, counted_blob, counted_blob_arc);

fn replace_part(package: &mut OpcPackage, part: Box<dyn Part>) -> Result<()> {
    let name = part.partname().clone();
    assert!(package.remove_part(&name));
    package.try_add_part(part)?;
    Ok(())
}

#[test]
fn a_foreign_arc_that_does_not_alias_its_payload_is_never_memoized_or_pinned() -> Result<()> {
    let package = text_box_package(3, 2)?;
    let mut opc = package.opc.clone();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").map_err(Error::Invalid)?;
    let foreign = Arc::new(b"foreign payload with a different identity".to_vec());
    let inner = copied_part(opc.get_part(&slide_name)?)?;
    replace_part(
        &mut opc,
        Box::new(MismatchedBlobArcPart {
            inner,
            foreign: Arc::clone(&foreign),
        }),
    )?;
    let package = Package::from_opc_package(opc)?;
    let snapshot = package.opened_presentation_with_limits(Limits::default())?;
    let foreign_index = snapshot
        .slides
        .iter()
        .position(|slide| slide.part_name == slide_name)
        .expect("foreign slide is captured");
    assert_eq!(
        snapshot.slide_roots.len(),
        2,
        "no owner proves the foreign key"
    );
    assert!(
        snapshot
            .slide_roots
            .lookup(slide_blob(&snapshot, foreign_index))
            .is_none()
    );
    let weak = Arc::downgrade(&foreign);
    drop(snapshot);
    drop(package);
    drop(foreign);
    assert!(
        weak.upgrade().is_none(),
        "the memo must not pin a foreign Arc"
    );
    Ok(())
}

#[test]
fn a_memo_hit_adds_no_payload_observation_over_a_plain_recapture() -> Result<()> {
    let package = text_box_package(2, 2)?;
    let mut opc = package.opc.clone();
    let slide_name = PackURI::new("/ppt/slides/slide1.xml").map_err(Error::Invalid)?;
    let observations = Arc::new(AtomicUsize::new(0));
    let inner = copied_part(opc.get_part(&slide_name)?)?;
    replace_part(
        &mut opc,
        Box::new(CountingBlobPart {
            inner,
            observations: Arc::clone(&observations),
        }),
    )?;
    let package = Package::from_opc_package(opc)?;
    let source = package.opened_presentation()?;
    let candidate = source.package.as_ref().clone();

    observations.store(0, Ordering::SeqCst);
    take_proved_root_hits();
    let assisted = recapture(&source, &candidate, true)?;
    let assisted_observations = observations.load(Ordering::SeqCst);
    assert_eq!(take_proved_root_hits(), 2);

    observations.store(0, Ordering::SeqCst);
    let plain = recapture(&source, &candidate, false)?;
    let plain_observations = observations.load(Ordering::SeqCst);
    assert_eq!(take_proved_root_hits(), 0);

    assert_same_snapshot("counted part", &assisted, &plain);
    assert_eq!(
        assisted_observations, plain_observations,
        "a memo lookup must not add raw payload observations"
    );
    Ok(())
}

#[test]
fn memo_assisted_recaptures_equal_plain_ones_over_the_pptx_corpus() -> Result<()> {
    let root = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data");
    let mut fixtures = Vec::new();
    collect_pptx_fixtures(&root, &mut fixtures);
    fixtures.sort();
    assert!(fixtures.len() >= 70, "expected the repository PPTX corpus");
    let mut captured = 0_usize;
    let mut hits = 0_usize;
    let mut committed = 0_usize;
    for fixture in &fixtures {
        let label = fixture.display().to_string();
        let Ok(bytes) = std::fs::read(fixture) else {
            continue;
        };
        let Ok(package) = Package::from_bytes(&bytes) else {
            continue;
        };
        let Ok(source) = package.opened_presentation() else {
            continue;
        };
        captured += 1;
        assert_entries_owned_by(&label, &source);
        let identical = source.package.as_ref().clone();
        take_proved_root_hits();
        let assisted = recapture(&source, &identical, true);
        let fixture_hits = take_proved_root_hits();
        hits += fixture_hits;
        let plain = recapture(&source, &identical, false);
        assert_same_result(&format!("{label}: identical"), &assisted, &plain);

        // Re-allocate every other slide: those miss, the rest may hit.
        let mut mixed = source.package.as_ref().clone();
        for slide in source.slides.iter().step_by(2) {
            reallocate_slide(&mut mixed, &slide.part_name)?;
        }
        let assisted = recapture(&source, &mixed, true);
        let plain = recapture(&source, &mixed, false);
        assert_same_result(&format!("{label}: mixed"), &assisted, &plain);

        // A real one-shape edit of the first, middle and last slide, committed
        // from the source with its memo and from the same source with an
        // empty one.
        let mut without = source.clone();
        without.slide_roots = Arc::new(SlideRootMemo::default());
        let count = source.slides.len();
        let mut positions = vec![0, count / 2, count.saturating_sub(1)];
        positions.dedup();
        for slide in positions.into_iter().filter(|&slide| slide < count) {
            let mut with_memo = source.edit();
            let mut plain_edit = without.edit();
            let staged = with_memo.set_shape_text(slide, 0, "0760 corpus edit");
            let staged_plain = plain_edit.set_shape_text(slide, 0, "0760 corpus edit");
            assert_eq!(
                format!("{staged:?}"),
                format!("{staged_plain:?}"),
                "{label}: the edit itself does not consult the memo"
            );
            if !matches!(staged, Ok(true)) {
                continue;
            }
            let assisted = with_memo.commit();
            let plain = plain_edit.commit();
            match (&assisted, &plain) {
                (Ok(assisted), Ok(plain)) => {
                    assert_same_snapshot(
                        &format!("{label}: slide {slide} commit"),
                        assisted.snapshot(),
                        plain.snapshot(),
                    );
                    assert_eq!(
                        assisted.patch().to_bytes()?,
                        plain.patch().to_bytes()?,
                        "{label}: slide {slide} patch"
                    );
                    committed += 1;
                },
                (Err(assisted), Err(plain)) => assert_eq!(
                    format!("{assisted:?}"),
                    format!("{plain:?}"),
                    "{label}: slide {slide} refusal"
                ),
                (assisted, plain) => panic!(
                    "{label}: slide {slide}: memo changed the commit: {assisted:?} versus {plain:?}"
                ),
            }
        }
        take_proved_root_hits();
    }
    println!(
        "0760-memo-oracle fixtures={} captured={captured} identical_hits={hits} commits={committed}",
        fixtures.len()
    );
    assert!(captured > 0 && hits > 0 && committed > 0);
    Ok(())
}

fn collect_pptx_fixtures(directory: &std::path::Path, found: &mut Vec<std::path::PathBuf>) {
    let Ok(entries) = std::fs::read_dir(directory) else {
        return;
    };
    for entry in entries.flatten() {
        let path = entry.path();
        if path.is_dir() {
            collect_pptx_fixtures(&path, found);
        } else if path
            .extension()
            .is_some_and(|extension| extension == "pptx")
        {
            found.push(path);
        }
    }
}

#[test]
fn the_notes_owning_deck_reuses_proofs_and_keeps_its_notes_graph() -> Result<()> {
    let mut package = Package::new()?;
    for index in 0..3 {
        let slide = package.presentation_mut()?.add_slide()?;
        slide.set_title(&format!("notes slide {index}"));
        slide.set_notes(&format!("speaker notes {index}"));
    }
    let package = Package::from_vec(package.to_bytes()?)?;
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    assert!(edit.set_notes_text(1, "rewritten notes")?);
    take_proved_root_hits();
    let commit = edit.commit()?;
    assert_eq!(take_proved_root_hits(), 3, "no slide payload changed");
    let cold = cold_capture(commit.snapshot())?;
    assert_same_snapshot("notes commit", commit.snapshot(), &cold);
    let notes = crate::notes::load_snapshot(
        commit.snapshot().package.as_ref(),
        &commit.snapshot().presentation_name,
    )?
    .expect("notes graph");
    assert_eq!(notes.slides().len(), 3);
    Ok(())
}

// ---------------------------------------------------------------------------
// ADR 0032 conditions, each proven by behaviour rather than asserted.
// ---------------------------------------------------------------------------

/// Condition: keyed on allocation identity, with a strong reference that keeps
/// the key from being recycled by a different payload while the entry lives.
#[test]
fn an_entry_holds_its_allocation_strongly_so_its_key_cannot_be_recycled() -> Result<()> {
    let payload = Arc::new(b"<p:sld/>".to_vec());
    let record = SlideRootRecord::for_test(payload.as_slice(), Some(Conformance::Transitional));
    let owner_key = (payload.as_ptr() as usize, payload.len());
    let owned = Arc::clone(&payload);
    let memo = SlideRootMemo::from_records([record].into_iter(), move |key| {
        (key == owner_key).then(|| Arc::clone(&owned))
    })?;
    assert_eq!(memo.len(), 1);
    assert_eq!(
        Arc::strong_count(&payload),
        2,
        "the entry holds one strong reference of its own"
    );
    let weak = Arc::downgrade(&payload);
    drop(payload);
    let alive = weak
        .upgrade()
        .expect("the entry keeps its allocation, and so its address, alive");
    assert_eq!(
        memo.lookup(alive.as_slice()),
        Some(Conformance::Transitional)
    );
    // Equal bytes elsewhere, and any sub-slice of the allocation, are misses.
    let copy = alive.as_slice().to_vec();
    assert_eq!(memo.lookup(&copy), None);
    assert_eq!(memo.lookup(&alive[..alive.len() - 1]), None);
    assert_eq!(memo.lookup(&alive[1..]), None);
    drop(alive);
    drop(memo);
    assert!(
        weak.upgrade().is_none(),
        "dropping the memo releases the only remaining reference"
    );
    Ok(())
}

/// Condition: every entry names an allocation the owning snapshot's own
/// package holds, through capture, commit, publication and rebind; and the
/// admission itself refuses unowned, non-aliasing and unsuccessful records.
#[test]
fn every_entry_names_an_allocation_its_snapshots_package_holds() -> Result<()> {
    let mut package = text_box_package(6, 2)?;
    let source = package.opened_presentation()?;
    assert_eq!(assert_entries_owned_by("capture", &source), 6);
    let mut edit = source.edit();
    assert!(edit.set_shape_text(3, 1, "owned")?);
    let commit = edit.commit()?;
    assert_eq!(assert_entries_owned_by("commit", commit.snapshot()), 6);
    // No committed entry retains the source's pre-edit allocation of slide 3.
    let replaced = slide_arc(&source, 3);
    assert!(
        commit
            .snapshot()
            .slide_roots
            .retained()
            .all(|(_key, raw, _)| !Arc::ptr_eq(raw, &replaced))
    );
    let published = package.apply_opened_presentation_commit(commit)?;
    assert_eq!(assert_entries_owned_by("published", &published), 6);
    let mut moved = published.package.as_ref().clone();
    reallocate_slide(&mut moved, &published.slides[5].part_name)?;
    assert_eq!(
        assert_entries_owned_by("rebound", &published.rebound_to(&moved)),
        5
    );
    Ok(())
}

#[test]
fn from_records_admits_only_owned_aliasing_successful_classifications() -> Result<()> {
    let owned = Arc::new(b"<p:sld owned/>".to_vec());
    let unowned = Arc::new(b"<p:sld unowned/>".to_vec());
    let refused = Arc::new(b"<p:sld refused/>".to_vec());
    let aliased = Arc::new(b"<p:sld aliased/>".to_vec());
    let impostor = Arc::new(b"<p:sld impostor/>".to_vec());
    let key = |payload: &Arc<Vec<u8>>| (payload.as_ptr() as usize, payload.len());
    let records = [
        SlideRootRecord::for_test(&owned, Some(Conformance::Strict)),
        SlideRootRecord::for_test(&unowned, Some(Conformance::Transitional)),
        // A slide the scan refused is recomputed from its bytes every time.
        SlideRootRecord::for_test(&refused, None),
        SlideRootRecord::for_test(&aliased, Some(Conformance::Transitional)),
    ];
    let (owned_key, refused_key, aliased_key) = (key(&owned), key(&refused), key(&aliased));
    let (owner_owned, owner_refused, owner_impostor) = (
        Arc::clone(&owned),
        Arc::clone(&refused),
        Arc::clone(&impostor),
    );
    let memo = SlideRootMemo::from_records(records.into_iter(), move |candidate| {
        if candidate == owned_key {
            Some(Arc::clone(&owner_owned))
        } else if candidate == refused_key {
            Some(Arc::clone(&owner_refused))
        } else if candidate == aliased_key {
            // An owner that answers with an allocation the key does not name
            // must not be admitted, let alone pinned.
            Some(Arc::clone(&owner_impostor))
        } else {
            None
        }
    })?;
    assert_eq!(memo.len(), 1, "only the owned, aliasing, successful record");
    assert_eq!(memo.lookup(&owned), Some(Conformance::Strict));
    for (label, payload) in [
        ("unowned", &unowned),
        ("refused", &refused),
        ("aliased", &aliased),
        ("impostor", &impostor),
    ] {
        assert_eq!(memo.lookup(payload), None, "{label}");
    }
    assert!(
        memo.retained()
            .all(|(_key, raw, _)| Arc::ptr_eq(raw, &owned)),
        "the memo retains nothing but the owned allocation"
    );
    Ok(())
}

/// Condition: a memo carried onto another package is projected onto that
/// package's own allocations, never inherited.
#[test]
fn a_rebind_projects_the_memo_and_never_inherits_a_moved_allocation() -> Result<()> {
    let package = text_box_package(4, 2)?;
    let source = package.opened_presentation()?;
    let mut moved = source.package.as_ref().clone();
    reallocate_slide(&mut moved, &source.slides[0].part_name)?;
    let stale = slide_arc(&source, 0);
    assert!(
        source
            .slide_roots
            .retained()
            .any(|(_key, raw, _)| Arc::ptr_eq(raw, &stale)),
        "the source memo names the moved slide's old allocation"
    );
    let references = Arc::strong_count(&stale);
    let rebound = source.rebound_to(&moved);
    assert_eq!(rebound.slide_roots.len(), 3);
    assert_eq!(assert_entries_owned_by("rebound", &rebound), 3);
    // Each kept entry retains the rebound package's own allocation.
    for index in 1..4 {
        let held = slide_arc(&rebound, index);
        assert!(
            rebound
                .slide_roots
                .retained()
                .any(|(_key, raw, _)| Arc::ptr_eq(raw, &held)),
            "slide {index} is projected onto the rebound package's allocation"
        );
    }
    assert!(!Arc::ptr_eq(&rebound.slide_roots, &source.slide_roots));
    // The source's allocation of the moved slide is not inherited: neither
    // rebound memo names it and the rebind took no reference to it. (The
    // package's own lazily decoded payload state, shared by every clone of
    // the package, may keep it alive; that is the OPC's, not a memo's.)
    assert_eq!(
        Arc::strong_count(&stale),
        references,
        "the rebind must take no reference to the moved slide's old allocation"
    );
    assert!(
        rebound
            .slide_roots
            .retained()
            .all(|(_key, raw, _)| !Arc::ptr_eq(raw, &stale))
    );
    assert!(
        rebound
            .part_digests
            .retained()
            .all(|(_key, raw)| !Arc::ptr_eq(raw, &stale))
    );
    drop(stale);
    drop(moved);
    drop(source);
    // A commit from the rebound snapshot misses the moved slide and still
    // equals a cold capture.
    let mut edit = rebound.edit();
    assert!(edit.set_shape_text(2, 0, "after rebind")?);
    take_proved_root_hits();
    let commit = edit.commit()?;
    assert_eq!(take_proved_root_hits(), 2, "slides 1 and 3 hit");
    assert_same_snapshot(
        "rebound commit",
        commit.snapshot(),
        &cold_capture(commit.snapshot())?,
    );
    Ok(())
}

/// ADR 0032 section 3: a rebind is infallible, so a projection that cannot
/// reserve its table is an empty memo, exactly as the digest projection
/// degrades; the rebound snapshot means the same and later commits only miss.
#[test]
fn a_refused_projection_is_an_empty_memo_like_the_digest_projection() -> Result<()> {
    let package = text_box_package(4, 2)?;
    let source = package.opened_presentation()?;
    let identical = source.package.as_ref().clone();
    let rebound = with_refused_slide_root_reservation(|| source.rebound_to(&identical));
    assert_eq!(rebound.slide_roots.len(), 0);
    assert_eq!(rebound.revision(), source.revision());
    assert_eq!(rebound.slides(), source.slides());
    let mut edit = rebound.edit();
    assert!(edit.set_shape_text(1, 0, "empty memo")?);
    take_proved_root_hits();
    let commit = edit.commit()?;
    assert_eq!(take_proved_root_hits(), 0, "an empty memo is only a miss");
    assert_same_snapshot(
        "empty-memo commit",
        commit.snapshot(),
        &cold_capture(commit.snapshot())?,
    );
    Ok(())
}

/// ADR 0032 section 3 and verification 2: a refused reservation while building
/// the memo is the typed `Error::Allocation`, and no snapshot, facade memo or
/// published byte is left behind.
#[test]
fn a_refused_memo_reservation_is_a_typed_allocation_error_and_leaves_no_partial_snapshot()
-> Result<()> {
    let mut package = text_box_package(4, 2)?;
    let published_before = package.to_bytes()?;

    // A capture refuses before it offers its digest memo to the facade.
    assert_memo_refusal(
        "capture",
        with_refused_slide_root_reservation(|| package.opened_presentation()),
    );
    assert!(package.part_digest_memo().is_none());

    // A commit refuses and leaves its source snapshot exactly as it was.
    let source = package.opened_presentation()?;
    let source_memo = Arc::clone(&source.slide_roots);
    let source_content = memo_content(&source.slide_roots);
    let mut edit = source.edit();
    assert!(edit.set_shape_text(2, 0, "refused")?);
    assert_memo_refusal(
        "commit",
        with_refused_slide_root_reservation(|| edit.clone().commit()),
    );
    assert!(Arc::ptr_eq(&source.slide_roots, &source_memo));
    assert_eq!(memo_content(&source.slide_roots), source_content);

    // A publication that must capture its candidate refuses atomically: the
    // facade's package, and so its published bytes, do not move.
    let commit = edit.clone().commit()?;
    let patch = commit.patch().clone();
    assert_memo_refusal(
        "publication",
        with_refused_slide_root_reservation(|| package.apply_opened_presentation_patch(&patch)),
    );
    assert_eq!(package.to_bytes()?, published_before);

    // Without the refusal every step succeeds and equals the cold path.
    let published = package.apply_opened_presentation_patch(&patch)?;
    assert_same_snapshot("published", &published, &cold_capture(&published)?);
    assert_same_snapshot(
        "commit",
        commit.snapshot(),
        &cold_capture(commit.snapshot())?,
    );
    Ok(())
}

/// Condition: a miss is an ordinary recomputation, so no value, refusal,
/// patch or published byte depends on which entries the memo holds.
#[test]
fn published_bytes_patches_and_revisions_never_depend_on_memo_contents() -> Result<()> {
    let package = text_box_package(6, 3)?;
    let source = package.opened_presentation()?;
    let full = source.clone();
    let mut empty = source.clone();
    empty.slide_roots = Arc::new(SlideRootMemo::default());
    let mut partial = source.clone();
    let keep: Vec<Arc<Vec<u8>>> = (0..6).step_by(2).map(|i| slide_arc(&source, i)).collect();
    partial.slide_roots = Arc::new(source.slide_roots.project(|key| {
        keep.iter()
            .find(|blob| (blob.as_ptr() as usize, blob.len()) == key)
            .cloned()
    }));
    assert_eq!(partial.slide_roots.len(), 3);

    let scenarios: [&[(usize, usize)]; 4] = [&[(0, 0)], &[(2, 1), (5, 2)], &[(1, 0), (1, 1)], &[]];
    for (index, edits) in scenarios.iter().enumerate() {
        let mut outcomes = Vec::new();
        for (label, root) in [("full", &full), ("empty", &empty), ("partial", &partial)] {
            let mut edit = root.edit();
            for &(slide, shape) in *edits {
                assert!(edit.set_shape_text(slide, shape, format!("scenario {index}"))?);
            }
            let commit = edit.commit()?;
            let patch = commit.patch().to_bytes()?;
            let revision = commit.snapshot().revision();
            let slides = commit.snapshot().slides().to_vec();
            // A second commit from the committed snapshot consults the memo
            // that recapture built.
            let mut next = commit.snapshot().edit();
            assert!(next.set_shape_text(4, 0, format!("second {index}"))?);
            let second = next.commit()?;
            let mut facade = text_box_package(6, 3)?;
            facade.apply_opened_presentation_commit(commit)?;
            facade.apply_opened_presentation_commit(second.clone())?;
            outcomes.push((
                label,
                patch,
                revision,
                slides,
                second.patch().to_bytes()?,
                second.snapshot().revision(),
                facade.to_bytes()?,
            ));
        }
        let (_, first_patch, first_revision, first_slides, second_patch, second_revision, bytes) =
            &outcomes[0];
        for (label, patch, revision, slides, next_patch, next_revision, published) in &outcomes[1..]
        {
            assert_eq!(patch, first_patch, "scenario {index} {label}: patch");
            assert_eq!(
                revision, first_revision,
                "scenario {index} {label}: revision"
            );
            assert_eq!(slides, first_slides, "scenario {index} {label}: slides");
            assert_eq!(
                next_patch, second_patch,
                "scenario {index} {label}: second patch"
            );
            assert_eq!(
                next_revision, second_revision,
                "scenario {index} {label}: second revision"
            );
            assert_eq!(
                published, bytes,
                "scenario {index} {label}: published bytes"
            );
        }
    }
    Ok(())
}

/// Condition: the memo is built by a capture or projected by a rebind, and
/// never mutated in place; clones share it.
#[test]
fn a_memo_is_built_or_projected_and_never_mutated() -> Result<()> {
    let mut package = text_box_package(5, 2)?;
    let source = package.opened_presentation()?;
    let memo = Arc::clone(&source.slide_roots);
    let content = memo_content(&memo);
    assert_eq!(content.len(), 5);
    let clone = source.clone();
    assert!(
        Arc::ptr_eq(&clone.slide_roots, &memo),
        "clones share one memo"
    );

    let mut edit = source.edit();
    assert!(edit.set_shape_text(0, 0, "never mutated")?);
    let commit = edit.commit()?;
    let committed = Arc::clone(&commit.snapshot().slide_roots);
    assert!(
        !Arc::ptr_eq(&committed, &memo),
        "a commit builds its own memo"
    );
    let committed_content = memo_content(&committed);
    let rebound = commit
        .snapshot()
        .rebound_to(commit.snapshot().package.as_ref());
    assert!(
        !Arc::ptr_eq(&rebound.slide_roots, &committed),
        "a rebind projects a new memo"
    );
    assert_eq!(memo_content(&rebound.slide_roots), committed_content);
    let published = package.apply_opened_presentation_commit(commit.clone())?;
    let mut again = published.edit();
    assert!(again.set_shape_text(4, 1, "again")?);
    let _second = again.commit()?;
    let _refused = with_refused_slide_root_reservation(|| clone.edit().commit());

    // Nothing that consulted or projected the memos changed them.
    assert_eq!(memo_content(&memo), content);
    assert_eq!(memo_content(&source.slide_roots), content);
    assert_eq!(memo_content(&committed), committed_content);
    assert_eq!(
        memo_content(&commit.snapshot().slide_roots),
        committed_content
    );
    Ok(())
}

/// Condition: the resident cost is one 32-byte entry per captured slide plus
/// the `Arc` header, bounded by the capture's slide limit.
#[test]
fn the_memo_costs_one_32_byte_entry_per_captured_slide() -> Result<()> {
    if size_of::<usize>() == 8 {
        assert_eq!(SlideRootMemo::ENTRY_BYTES, 32);
        // Strong and weak counts, then the table's vector header.
        assert_eq!(2 * size_of::<usize>() + size_of::<SlideRootMemo>(), 40);
    }
    let package = text_box_package(7, 1)?;
    let snapshot = package.opened_presentation()?;
    assert_eq!(snapshot.slide_roots.len(), 7);
    assert_eq!(
        snapshot.slide_roots.resident_bytes(),
        7 * SlideRootMemo::ENTRY_BYTES,
        "the table is reserved exactly, one entry per captured slide"
    );
    // The slide count is bounded by the capture's limit, and so is the memo.
    let limits = Limits::new(6, 1 << 20, 1 << 20, 4, 1 << 20, 1 << 20).expect("limits");
    match package.opened_presentation_with_limits(limits) {
        Err(Error::Limit { resource, limit }) => {
            assert_eq!(resource, "opened-presentation slides");
            assert_eq!(limit, 6);
        },
        other => panic!("seven slides exceed a six-part limit: {other:?}"),
    }
    Ok(())
}

/// Condition: every reused value is re-derived by a `debug_assert!` in test and
/// debug builds; in such a build every hit still runs the scan once.
#[test]
fn every_hit_is_re_derived_in_debug_builds() -> Result<()> {
    let package = text_box_package(6, 2)?;
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    assert!(edit.set_shape_text(1, 0, "re-derived")?);
    assert!(edit.set_shape_text(4, 0, "re-derived")?);
    take_proved_root_hits();
    take_root_scans();
    let commit = edit.commit()?;
    let hits = take_proved_root_hits();
    let scans = take_root_scans();
    assert_eq!(hits, 4, "four untouched slides hit");
    let misses = 2;
    if cfg!(debug_assertions) {
        assert_eq!(scans, misses + hits, "each hit is re-derived once");
    } else {
        assert_eq!(scans, misses);
    }
    assert_same_snapshot(
        "re-derived commit",
        commit.snapshot(),
        &cold_capture(commit.snapshot())?,
    );
    Ok(())
}

/// The re-derivation is not vacuous: a planted classification that disagrees
/// with the bytes it names fails the debug assertion at its first hit.
#[cfg(debug_assertions)]
#[test]
#[should_panic(expected = "a proved slide-root classification answered for different bytes")]
fn a_planted_wrong_classification_is_caught_by_the_re_derivation() {
    let package = text_box_package(3, 1).expect("package");
    let mut source = package.opened_presentation().expect("snapshot");
    let records: Vec<SlideRootRecord> = (0..3)
        .map(|index| {
            SlideRootRecord::for_test(slide_blob(&source, index), Some(Conformance::Strict))
        })
        .collect();
    let package_ref = Arc::clone(&source.package);
    let names: Vec<PackURI> = source
        .slides
        .iter()
        .map(|slide| slide.part_name.clone())
        .collect();
    let planted = SlideRootMemo::from_records(records.into_iter(), |key| {
        names.iter().find_map(|name| {
            let blob = package_ref.get_part(name).ok()?.blob_arc();
            ((blob.as_ptr() as usize, blob.len()) == key).then_some(blob)
        })
    })
    .expect("planted memo");
    assert_eq!(planted.len(), 3);
    source.slide_roots = Arc::new(planted);
    let mut edit = source.edit();
    assert!(edit.set_shape_text(1, 0, "planted").expect("edit"));
    let _ = edit.commit();
}

/// The facade publishes snapshots but keeps no slide-root memo of its own, so
/// its adoption condition is vacuous: once every snapshot is dropped, each
/// slide payload is held only by the facade's package and its digest memo.
#[test]
fn the_facade_retains_no_slide_root_memo() -> Result<()> {
    let mut package = text_box_package(3, 1)?;
    let source = package.opened_presentation()?;
    let mut edit = source.edit();
    assert!(edit.set_shape_text(0, 0, "facade")?);
    let commit = edit.commit()?;
    let published = package.apply_opened_presentation_commit(commit)?;
    let names: Vec<PackURI> = published
        .slides
        .iter()
        .map(|slide| slide.part_name.clone())
        .collect();
    assert_eq!(published.slide_roots.len(), 3);
    drop(published);
    drop(source);
    assert!(package.part_digest_memo().is_some());
    for name in &names {
        let blob = package.opc.get_part(name)?.blob_arc();
        // The facade's package, its digest memo and the `blob` handle above.
        assert_eq!(Arc::strong_count(&blob), 3, "{name}");
    }
    Ok(())
}
