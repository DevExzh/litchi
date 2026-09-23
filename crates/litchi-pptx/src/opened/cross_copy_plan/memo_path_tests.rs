//! Positive proof that the owned cross-copy's plan and application take the
//! memoized routes of change 0751, and that each memo answers only for
//! allocations it provably names.
//!
//! `digest_reuse_tests` proves the values did not move. These tests prove the
//! saving is real and bounded: which fingerprints consulted a memo, how many
//! payloads each still hashed, that the live physical revision of an
//! unmodified owned source is sealed from the digest bound to its archive, and
//! that the facade keeps a capture's memo only as re-projected onto the
//! allocations its own graph holds.
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "memo-path tests use panic-on-fixture-failure assertions"
)]

use std::sync::Arc;

use litchi_opc::{OpcPackage, PackURI};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

use super::{
    CrossSlideCopyPatch, CrossSlideCopyPlan, physical_package_fingerprint,
    streamed_physical_revision,
};
use crate::media_parts::Resource;
use crate::opened::Limits;
use crate::opened::model::fingerprint_log::{self, Entry};
use crate::opened::tests::{CopyingPart, assert_memo_retains_only_allocations_of};
use crate::{Error, Package, Result};

const PHOTO: &str = "/ppt/media/memo-photo.png";
const FLAT: &str = "/ppt/media/memo-flat.png";
const SOURCE_SLIDE: usize = 2;
const DESTINATION_SLIDE: usize = 1;
const POSITION: usize = 1;

fn noisy(len: usize, mut state: u32) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(len);
    while bytes.len() < len {
        state ^= state << 13;
        state ^= state >> 17;
        state ^= state << 5;
        bytes.push((state >> 24) as u8);
    }
    bytes[..8].copy_from_slice(b"\x89PNG\r\n\x1a\n");
    bytes
}

fn authored(prefix: &str, slides: usize) -> Result<Vec<u8>> {
    let mut package = Package::new()?;
    let presentation = package.presentation_mut()?;
    for index in 0..slides {
        presentation
            .add_slide()?
            .set_title(&format!("change-0751-memo-{prefix}-{index}"));
    }
    package.to_bytes()
}

fn with_pictures(bytes: &[u8], slide: usize, pictures: &[(&str, &[u8])]) -> Result<Vec<u8>> {
    let mut package = Package::from_bytes(bytes)?;
    let mut edit = package.opened_presentation_transaction()?;
    for (index, (part_name, data)) in pictures.iter().enumerate() {
        edit.add_picture(
            slide,
            format!("change-0751-memo-picture-{index}"),
            &Resource::new(*part_name, "image/png", data.to_vec()),
            (800, 800, 72, 72),
        )?;
    }
    let commit = edit.commit()?;
    package.apply_opened_presentation_commit(commit)?;
    package.to_bytes()
}

fn repack_stored(bytes: &[u8], stored: &str) -> Vec<u8> {
    let reader = ArchiveReader::new(bytes).expect("fixture archive");
    let mut writer = StreamingArchiveWriter::new();
    for name in reader.file_names() {
        let payload = reader.read(name).expect("fixture member");
        if name == stored {
            writer.write_stored(name, &payload)
        } else {
            writer.write_deflated_sized(name, &payload)
        }
        .expect("repacked member");
    }
    writer.finish_to_bytes().expect("repacked archive")
}

/// A source whose copied slide carries two images, one stored, and a
/// destination already using the same media names.
///
/// The candidate reopen takes a payload from the built graph only when the
/// donor's vector capacity is no larger than its own decode's. The source's
/// deferred decode allocates one byte of spare capacity, while the eager
/// reopen reads a Stored member at exactly its length, so the stored photo's
/// copy keeps the reopen's own allocation and no memo names it: its capture
/// hashes it, as does the base. The deflated image's copy is donated.
fn media_pair() -> Result<(Vec<u8>, Vec<u8>)> {
    let photo = noisy(32 * 1024, 0x0751_0002);
    let flat: Vec<u8> = noisy(16 * 1024, 0x0751_0003);
    let pictures: [(&str, &[u8]); 2] = [(PHOTO, &photo), (FLAT, &flat)];
    let source = repack_stored(
        &with_pictures(&authored("source", 3)?, SOURCE_SLIDE, &pictures)?,
        &PHOTO[1..],
    );
    let destination = with_pictures(&authored("destination", 2)?, 0, &pictures)?;
    Ok((source, destination))
}

fn plan_copy(source: &Package, destination: &Package) -> Result<CrossSlideCopyPlan> {
    destination.opened_presentation()?.plan_cross_slide_copy(
        &source.opened_presentation()?,
        SOURCE_SLIDE,
        DESTINATION_SLIDE,
        POSITION,
    )
}

/// Payloads of the candidate no snapshot memo names: the rewritten
/// presentation part and the stored photo's copy (see [`media_pair`]).
const UNSHARED: usize = 2;

fn hashed(entries: &[Entry]) -> Vec<usize> {
    entries.iter().map(|entry| entry.hashed).collect()
}

/// Planning captures its candidate consulting both snapshots' memos, so the
/// only payloads it hashes are the presentation part the copy rewrote and the
/// stored photo's copy, which the reopen did not take from the built graph
/// (see [`media_pair`]).
#[test]
fn planning_hashes_only_payloads_no_snapshot_holds() -> Result<()> {
    let (source_bytes, destination_bytes) = media_pair()?;
    let source = Package::from_vec(source_bytes)?;
    let destination = Package::from_vec(destination_bytes)?;
    let source_snapshot = source.opened_presentation()?;
    let destination_snapshot = destination.opened_presentation()?;
    let _cold = fingerprint_log::take();
    let plan = destination_snapshot.plan_cross_slide_copy(
        &source_snapshot,
        SOURCE_SLIDE,
        DESTINATION_SLIDE,
        POSITION,
    )?;
    assert!(plan.transfers_source_compressed_media());
    let consulted = fingerprint_log::take_with_parent();
    assert_eq!(
        hashed(&consulted),
        [UNSHARED],
        "the candidate capture hashes only the payloads no snapshot holds: {consulted:?}"
    );
    Ok(())
}

/// An application consults the facades' memos for both live revisions, the
/// snapshots' memos for its candidate capture, and that capture's memo for
/// the capture it publishes: four memo-consulting fingerprints that hash only
/// the candidate's unshared payloads between them, where the base hashed every
/// payload of four packages.
#[test]
fn an_application_hashes_only_payloads_no_memo_names() -> Result<()> {
    let (source_bytes, destination_bytes) = media_pair()?;
    let source = Package::from_vec(source_bytes)?;
    let mut destination = Package::from_vec(destination_bytes)?;
    let plan = plan_copy(&source, &destination)?;
    assert!(plan.retained_candidate_bytes().is_some());
    assert!(
        source.part_digest_memo().is_some(),
        "planning's capture filled it"
    );
    assert!(
        destination.part_digest_memo().is_some(),
        "planning's capture filled it"
    );
    let _planning = fingerprint_log::take();

    destination.apply_cross_slide_copy_plan(&source, &plan)?;
    let consulted = fingerprint_log::take_with_parent();
    assert_eq!(
        hashed(&consulted),
        [0, 0, UNSHARED, 0],
        "live source, live destination, candidate capture, published capture: {consulted:?}"
    );
    assert_eq!(consulted[0].parts, source.opc.part_count());
    assert_eq!(consulted[2].parts, destination.opc.part_count());
    assert_eq!(consulted[3].parts, destination.opc.part_count());
    // The facade adopted the published capture's memo.
    assert_eq!(
        destination.part_digest_memo().map(|memo| memo.len()),
        Some(destination.opc.part_count())
    );
    Ok(())
}

/// Without facade memos the live revisions are computed from cold, as before,
/// and the candidate captures still consult the memos those passes filled.
#[test]
fn an_application_without_facade_memos_hashes_the_live_packages_once() -> Result<()> {
    let (source_bytes, destination_bytes) = media_pair()?;
    let plan = {
        let source = Package::from_vec(source_bytes.clone())?;
        let destination = Package::from_vec(destination_bytes.clone())?;
        plan_copy(&source, &destination)?
    };
    let source = Package::from_vec(source_bytes)?;
    let mut destination = Package::from_vec(destination_bytes)?;
    assert!(source.part_digest_memo().is_none());
    assert!(destination.part_digest_memo().is_none());
    let _setup = fingerprint_log::take();
    destination.apply_cross_slide_copy_plan(&source, &plan)?;
    let log = fingerprint_log::take();
    let cold: Vec<&Entry> = log.iter().filter(|entry| !entry.with_parent).collect();
    let consulted: Vec<&Entry> = log.iter().filter(|entry| entry.with_parent).collect();
    // The two live passes hash every payload; debug builds add memo-free
    // re-derivations, all of which also appear as cold entries.
    assert!(cold.len() >= 2, "{log:?}");
    assert_eq!(cold[0].hashed, cold[0].parts, "{log:?}");
    assert_eq!(cold[1].hashed, cold[1].parts, "{log:?}");
    assert_eq!(
        consulted
            .iter()
            .map(|entry| entry.hashed)
            .collect::<Vec<_>>(),
        [UNSHARED, 0],
        "candidate capture and published capture: {log:?}"
    );
    Ok(())
}

/// Both durable-patch routes consult the memos. The forward route does so as
/// an application does. The inverse route's live revisions and its restored
/// clone's capture consult the facade's and the live destination's memos, and
/// the capture it publishes consults the restored capture's memo. Its proof
/// replans the forward copy against the restored clone, which is not an
/// unmodified owned source, so a transferring copy is built from the reopen of
/// that clone's serialization: those allocations are fresh, no memo names
/// them, and that capture hashes as the base's does.
#[test]
fn durable_patch_routes_consult_the_memos() -> Result<()> {
    let (source_bytes, destination_bytes) = media_pair()?;
    let source = Package::from_vec(source_bytes)?;
    let mut destination = Package::from_vec(destination_bytes.clone())?;
    let plan = plan_copy(&source, &destination)?;
    let forward = CrossSlideCopyPatch::from_bytes(&plan.patch().to_bytes()?)?;
    let _planning = fingerprint_log::take();

    destination.apply_cross_slide_copy_patch(&source, &forward)?;
    let consulted = fingerprint_log::take_with_parent();
    assert_eq!(
        hashed(&consulted),
        [0, 0, UNSHARED, 0],
        "forward: {consulted:?}"
    );

    destination.apply_cross_slide_copy_patch(&source, &forward.inverse())?;
    let consulted = fingerprint_log::take_with_parent();
    let summary = hashed(&consulted);
    assert_eq!(summary.len(), 5, "{consulted:?}");
    assert_eq!(summary[..2], [0, 0], "the live revisions: {consulted:?}");
    assert_eq!(
        summary[2], 1,
        "the restored clone's capture hashes only the presentation part the inverse restored: {consulted:?}"
    );
    assert_eq!(
        summary[4], 0,
        "the inverse's published capture hashes nothing: {consulted:?}"
    );
    assert_eq!(destination.to_bytes()?, destination_bytes);
    Ok(())
}

/// A capture of the facade's current graph offers its memo, and the next
/// capture consults it and hashes nothing.
#[test]
fn a_capture_fills_the_facade_memo_and_the_next_capture_reuses_it() -> Result<()> {
    let (source_bytes, _destination) = media_pair()?;
    let package = Package::from_vec(source_bytes)?;
    assert!(package.part_digest_memo().is_none());
    let first = package.opened_presentation()?;
    let memo = package
        .part_digest_memo()
        .expect("the capture offered its memo");
    assert_eq!(memo.len(), package.opc.part_count());
    let _cold = fingerprint_log::take();
    let second = package.opened_presentation()?;
    assert_eq!(second.revision(), first.revision());
    let consulted = fingerprint_log::take_with_parent();
    assert_eq!(hashed(&consulted), [0], "{consulted:?}");
    Ok(())
}

/// A capture's memo is re-projected onto the facade's own allocations (change
/// 0751's review). A facade holding a caller-defined part that copies its
/// payload when cloned keeps every entry but that part's: the snapshot's memo
/// names the snapshot's copy, which the facade does not hold, and every entry
/// the facade keeps retains the facade's own `Arc`. A later capture still
/// computes every value.
#[test]
fn a_capture_memo_is_projected_onto_the_facade_allocations() -> Result<()> {
    let (source_bytes, _destination) = media_pair()?;
    let mut package = Package::from_vec(source_bytes)?;
    let custom = PackURI::new("/custom/copying.bin").expect("URI");
    package.opc.try_add_part(Box::new(CopyingPart::new(
        custom.clone(),
        b"copied on clone".to_vec(),
    )))?;
    assert!(!package.opc.holds_only_built_in_parts());
    let first = package.opened_presentation()?;
    let copy = first.package.get_part(&custom)?.blob_arc();
    let held = package.opc.get_part(&custom)?.blob_arc();
    assert!(!Arc::ptr_eq(&copy, &held), "the part copies its payload");
    let copy_key = (copy.as_ptr() as usize, copy.len());
    assert!(
        first.part_digests.get_for_test(copy_key).is_some(),
        "the snapshot's memo names its copy, an entry the facade would not hold"
    );
    let memo = package
        .part_digest_memo()
        .expect("the capture offered its memo");
    assert_eq!(
        memo.get_for_test(copy_key),
        None,
        "the entry naming the snapshot's copy is dropped"
    );
    assert_eq!(memo.len(), first.part_digests.len() - 1);
    assert_memo_retains_only_allocations_of("capture memo", memo, &package.opc);
    let second = package.opened_presentation()?;
    assert_eq!(second.revision(), first.revision());
    Ok(())
}

/// An unmodified owned source's physical revision is sealed from the digest
/// bound to its archive, equals the streamed revision, and is refused over
/// the archive bound exactly as the stream refuses it; an edited package is
/// streamed.
#[test]
fn an_exact_source_seals_its_bound_digest_and_keeps_the_bound() -> Result<()> {
    let (source_bytes, _destination) = media_pair()?;
    let package = OpcPackage::from_vec(source_bytes.clone())?;
    let limits = Limits::default();
    assert_eq!(package.exact_source_sha256(), {
        use sha2::Digest as _;
        Some(<[u8; 32]>::from(sha2::Sha256::digest(&source_bytes)))
    });
    let sealed = physical_package_fingerprint(&package, limits)?;
    assert_eq!(sealed, streamed_physical_revision(&package, limits)?);
    // A clone shares the archive and its digest.
    assert_eq!(
        physical_package_fingerprint(&package.clone(), limits)?,
        sealed
    );

    let tight = Limits::new(4096, source_bytes.len() - 1, 1024, 1, 1, 1).expect("limits");
    let refused = physical_package_fingerprint(&package, tight);
    let streamed = streamed_physical_revision(&package, tight);
    assert!(
        matches!(
            refused,
            Err(Error::Limit {
                resource: "cross-slide serialized archive bytes",
                limit,
            }) if limit == source_bytes.len() - 1
        ),
        "{refused:?}"
    );
    assert_eq!(format!("{refused:?}"), format!("{streamed:?}"));
    let exact = Limits::new(4096, source_bytes.len(), 1024, 1, 1, 1).expect("limits");
    assert_eq!(physical_package_fingerprint(&package, exact)?, sealed);

    let mut edited = package.clone();
    edited.set_save_options(litchi_opc::SaveOptions::default());
    assert_eq!(edited.exact_source_sha256(), None);
    assert_eq!(
        physical_package_fingerprint(&edited, limits)?,
        streamed_physical_revision(&edited, limits)?
    );
    Ok(())
}
