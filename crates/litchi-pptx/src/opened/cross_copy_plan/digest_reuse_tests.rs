//! Byte identity of the owned cross-copy's durable patches, revisions,
//! published archives and refusals across change 0751.
//!
//! Change 0751 stops re-hashing bytes whose identity is already proven: the
//! application's live revisions consult the facade's payload-digest memos,
//! the candidate captures consult the snapshots' memos, an unmodified owned
//! source's physical revision is sealed from the digest `litchi-opc` binds to
//! its archive, and an application shares a plan's retained archive instead of
//! copying it. None of that may move a value. This module records, for a set
//! of deterministic scenarios, the SHA-256 of every durable patch encoding, the
//! six recorded revisions, every published archive and returned snapshot
//! revision, and the text of every refusal, and compares the transcript with
//! the one base `6d989cad63` produces. The module compiles unchanged on the
//! base tree, which is how `GOLDEN` was generated.
//!
//! After 0751's review the transcript also covers a destination holding a
//! caller-defined part, by plan and by durable patch in both directions, and
//! size-limit refusals end to end: durable patches read under limits that
//! admit the patch but not the source or destination archive, and planning
//! under the same limits.
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "golden-transcript tests use panic-on-fixture-failure assertions"
)]

use std::fmt::Write as _;

use sha2::{Digest, Sha256};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

use super::{CrossSlideCopyPatch, CrossSlideCopyPlan};
use crate::media_parts::Resource;
use crate::opened::Limits;
use crate::{Package, Result};

const PHOTO: &str = "/ppt/media/digest-photo.png";
const FLAT: &str = "/ppt/media/digest-flat.png";
const SOURCE_SLIDE: usize = 2;
const DESTINATION_SLIDE: usize = 1;
const POSITION: usize = 1;

fn hex(bytes: &[u8]) -> String {
    let digest = Sha256::digest(bytes);
    let mut text = String::with_capacity(64);
    for byte in digest {
        write!(text, "{byte:02x}").unwrap();
    }
    text
}

fn hex32(value: [u8; 32]) -> String {
    let mut text = String::with_capacity(64);
    for byte in value {
        write!(text, "{byte:02x}").unwrap();
    }
    text
}

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

fn flat(len: usize) -> Vec<u8> {
    let mut bytes: Vec<u8> = (0..len).map(|index| (index % 5) as u8).collect();
    bytes[..8].copy_from_slice(b"\x89PNG\r\n\x1a\n");
    bytes
}

fn authored(prefix: &str, slides: usize) -> Result<Vec<u8>> {
    let mut package = Package::new()?;
    let presentation = package.presentation_mut()?;
    for index in 0..slides {
        presentation
            .add_slide()?
            .set_title(&format!("change-0751-{prefix}-{index}"));
    }
    package.to_bytes()
}

fn with_pictures(bytes: &[u8], slide: usize, pictures: &[(&str, &str, &[u8])]) -> Result<Vec<u8>> {
    let mut package = Package::from_bytes(bytes)?;
    let mut edit = package.opened_presentation_transaction()?;
    for (index, (part_name, content_type, data)) in pictures.iter().enumerate() {
        edit.add_picture(
            slide,
            format!("change-0751-picture-{index}"),
            &Resource::new(*part_name, *content_type, data.to_vec()),
            (800, 800, 72, 72),
        )?;
    }
    let commit = edit.commit()?;
    package.apply_opened_presentation_commit(commit)?;
    package.to_bytes()
}

/// Rewrite every member through a fresh writer, storing the members `stored`
/// selects and deflating the rest with known sizes.
fn repack(bytes: &[u8], stored: impl Fn(&str) -> bool) -> Vec<u8> {
    let reader = ArchiveReader::new(bytes).expect("fixture archive");
    let mut writer = StreamingArchiveWriter::new();
    for name in reader.file_names() {
        let payload = reader.read(name).expect("fixture member");
        if stored(name) {
            writer.write_stored(name, &payload)
        } else {
            writer.write_deflated_sized(name, &payload)
        }
        .expect("repacked member");
    }
    writer.finish_to_bytes().expect("repacked archive")
}

/// A source whose copied slide carries a stored incompressible photo and a
/// deflated flat image, and a destination already using the same media names.
fn media_pair() -> Result<(Vec<u8>, Vec<u8>)> {
    let photo = noisy(40 * 1024, 0x0751_0001);
    let flat = flat(24 * 1024);
    let pictures: [(&str, &str, &[u8]); 2] =
        [(PHOTO, "image/png", &photo), (FLAT, "image/png", &flat)];
    let source = repack(
        &with_pictures(&authored("source", 3)?, SOURCE_SLIDE, &pictures)?,
        |name| name == &PHOTO[1..],
    );
    let destination = with_pictures(&authored("destination", 2)?, 0, &pictures)?;
    Ok((source, destination))
}

/// A pair whose copied slide has no image.
fn plain_pair() -> Result<(Vec<u8>, Vec<u8>)> {
    Ok((
        authored("plain-source", 3)?,
        authored("plain-destination", 2)?,
    ))
}

fn plan_copy(source: &Package, destination: &Package) -> Result<CrossSlideCopyPlan> {
    destination.opened_presentation()?.plan_cross_slide_copy(
        &source.opened_presentation()?,
        SOURCE_SLIDE,
        DESTINATION_SLIDE,
        POSITION,
    )
}

fn record_patch(transcript: &mut String, label: &str, patch: &CrossSlideCopyPatch) -> Result<()> {
    writeln!(transcript, "{label}.bytes={}", hex(&patch.to_bytes()?)).unwrap();
    writeln!(
        transcript,
        "{label}.inverse={}",
        hex(&patch.inverse().to_bytes()?)
    )
    .unwrap();
    for (name, value) in [
        ("source", patch.source_revision()),
        ("destination", patch.destination_revision()),
        ("target", patch.target_revision()),
        ("source_physical", patch.source_physical_revision()),
        (
            "destination_physical",
            patch.destination_physical_revision(),
        ),
        ("target_physical", patch.target_physical_revision()),
    ] {
        writeln!(transcript, "{label}.{name}={}", hex32(value)).unwrap();
    }
    writeln!(
        transcript,
        "{label}.transfers={}",
        patch.transfers_source_compressed_media()
    )
    .unwrap();
    Ok(())
}

fn record_outcome(
    transcript: &mut String,
    label: &str,
    package: &mut Package,
    outcome: Result<crate::opened::Snapshot>,
) -> Result<()> {
    match outcome {
        Ok(snapshot) => {
            writeln!(
                transcript,
                "{label}.published={}",
                hex(&package.to_bytes()?)
            )
            .unwrap();
            writeln!(
                transcript,
                "{label}.revision={}",
                hex32(snapshot.revision())
            )
            .unwrap();
        },
        Err(error) => {
            writeln!(transcript, "{label}.refused={error:?}").unwrap();
            writeln!(
                transcript,
                "{label}.untouched={}",
                hex(&package.to_bytes()?)
            )
            .unwrap();
        },
    }
    Ok(())
}

/// Everything one pair publishes and refuses, by plan and by durable patch.
fn pair_transcript(label: &str, source_bytes: &[u8], destination_bytes: &[u8]) -> Result<String> {
    let mut transcript = String::new();
    let source = Package::from_vec(source_bytes.to_vec())?;
    let destination = Package::from_vec(destination_bytes.to_vec())?;
    let plan = plan_copy(&source, &destination)?;
    writeln!(
        transcript,
        "{label}.plan.retained={}",
        plan.retained_candidate_bytes().is_some()
    )
    .unwrap();
    record_patch(&mut transcript, &format!("{label}.plan"), plan.patch())?;
    let durable = CrossSlideCopyPatch::from_bytes(&plan.patch().to_bytes()?)?;

    // The plan, retained and released, into the destination it was planned
    // against and into a fresh open of the same bytes.
    let mut retained = Package::from_vec(destination_bytes.to_vec())?;
    let outcome = retained.apply_cross_slide_copy_plan(&source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.apply.retained"),
        &mut retained,
        outcome,
    )?;
    let mut released_plan = plan.clone();
    released_plan.release_retained_candidate();
    let mut released = Package::from_vec(destination_bytes.to_vec())?;
    let outcome = released.apply_cross_slide_copy_plan(&source, &released_plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.apply.released"),
        &mut released,
        outcome,
    )?;
    // The same plan applied again, to the destination the planning facade
    // captured (its memo is warm).
    let mut warm = destination;
    let outcome = warm.apply_cross_slide_copy_plan(&source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.apply.warm"),
        &mut warm,
        outcome,
    )?;

    // The durable patch forward, its inverse (undo) and the forward again
    // (redo), by patch; then redo by plan.
    let mut patched = Package::from_vec(destination_bytes.to_vec())?;
    let outcome = patched.apply_cross_slide_copy_patch(&source, &durable);
    record_outcome(
        &mut transcript,
        &format!("{label}.patch.forward"),
        &mut patched,
        outcome,
    )?;
    let outcome = patched.apply_cross_slide_copy_patch(&source, &durable.inverse());
    record_outcome(
        &mut transcript,
        &format!("{label}.patch.undo"),
        &mut patched,
        outcome,
    )?;
    let outcome = patched.apply_cross_slide_copy_patch(&source, &durable);
    record_outcome(
        &mut transcript,
        &format!("{label}.patch.redo"),
        &mut patched,
        outcome,
    )?;
    let outcome = patched.apply_cross_slide_copy_patch(&source, &durable.inverse());
    record_outcome(
        &mut transcript,
        &format!("{label}.patch.undo2"),
        &mut patched,
        outcome,
    )?;
    let outcome = patched.apply_cross_slide_copy_plan(&source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.plan.redo"),
        &mut patched,
        outcome,
    )?;

    // Applying a forward patch twice is stale: the destination moved.
    let outcome = patched.apply_cross_slide_copy_patch(&source, &durable);
    record_outcome(
        &mut transcript,
        &format!("{label}.patch.stale"),
        &mut patched,
        outcome,
    )?;

    // A destination that moved after planning is refused.
    let mut moved = Package::from_vec(destination_bytes.to_vec())?;
    {
        let mut edit = moved.opened_presentation_transaction()?;
        edit.set_shape_text(0, crate::shape::Key::Index(0), "change-0751 moved")?;
        let commit = edit.commit()?;
        moved.apply_opened_presentation_commit(commit)?;
    }
    let outcome = moved.apply_cross_slide_copy_plan(&source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.stale.destination"),
        &mut moved,
        outcome,
    )?;

    // A source that moved after planning, and a foreign source, are refused.
    let mut moved_source = Package::from_vec(source_bytes.to_vec())?;
    {
        let mut edit = moved_source.opened_presentation_transaction()?;
        edit.set_shape_text(0, crate::shape::Key::Index(0), "change-0751 moved source")?;
        let commit = edit.commit()?;
        moved_source.apply_opened_presentation_commit(commit)?;
    }
    let mut target = Package::from_vec(destination_bytes.to_vec())?;
    let outcome = target.apply_cross_slide_copy_plan(&moved_source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.stale.source"),
        &mut target,
        outcome,
    )?;
    let foreign = Package::from_vec(destination_bytes.to_vec())?;
    let outcome = target.apply_cross_slide_copy_plan(&foreign, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.foreign.source"),
        &mut target,
        outcome,
    )?;

    // A destination with the same content and different archive bytes is
    // refused by the physical proof.
    let repacked = repack(destination_bytes, |_name| true);
    let mut physical = Package::from_vec(repacked)?;
    let outcome = physical.apply_cross_slide_copy_plan(&source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.stale.physical"),
        &mut physical,
        outcome,
    )?;

    // A destination whose exact-source authorization an edit revoked while
    // keeping its bytes accepts the copy and publishes the same output.
    let mut revoked = Package::from_vec(destination_bytes.to_vec())?;
    revoked.edit_opc(|opc| {
        opc.set_save_options(litchi_opc::SaveOptions::default());
        Ok(())
    })?;
    let outcome = revoked.apply_cross_slide_copy_plan(&source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.revoked.destination"),
        &mut revoked,
        outcome,
    )?;

    // A source edited back to its own bytes plans and publishes the same copy.
    let mut resaved_source = Package::from_vec(source_bytes.to_vec())?;
    resaved_source.edit_opc(|opc| {
        opc.set_save_options(litchi_opc::SaveOptions::default());
        Ok(())
    })?;
    let fresh = Package::from_vec(destination_bytes.to_vec())?;
    let replanned = plan_copy(&resaved_source, &fresh)?;
    record_patch(
        &mut transcript,
        &format!("{label}.revoked.source"),
        replanned.patch(),
    )?;
    let mut target = fresh;
    let outcome = target.apply_cross_slide_copy_plan(&resaved_source, &replanned);
    record_outcome(
        &mut transcript,
        &format!("{label}.revoked.source.apply"),
        &mut target,
        outcome,
    )?;
    Ok(transcript)
}

/// A caller-defined part that copies its payload whenever it is cloned.
struct CallerPart {
    inner: litchi_opc::BlobPart,
}

impl CallerPart {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            inner: litchi_opc::BlobPart::new(
                litchi_opc::PackURI::new("/custom/change-0751.bin").expect("part name"),
                "application/octet-stream".to_owned(),
                bytes,
            ),
        }
    }
}

impl Clone for CallerPart {
    fn clone(&self) -> Self {
        Self::new(litchi_opc::Part::blob(&self.inner).to_vec())
    }
}

impl litchi_opc::Part for CallerPart {
    fn blob(&self) -> &[u8] {
        litchi_opc::Part::blob(&self.inner)
    }
    fn blob_arc(&self) -> std::sync::Arc<Vec<u8>> {
        litchi_opc::Part::blob_arc(&self.inner)
    }
    fn content_type(&self) -> &str {
        litchi_opc::Part::content_type(&self.inner)
    }
    fn partname(&self) -> &litchi_opc::PackURI {
        litchi_opc::Part::partname(&self.inner)
    }
    fn rels(&self) -> &litchi_opc::Relationships {
        litchi_opc::Part::rels(&self.inner)
    }
    fn rels_mut(&mut self) -> &mut litchi_opc::Relationships {
        litchi_opc::Part::rels_mut(&mut self.inner)
    }
    fn set_blob(&mut self, blob: Vec<u8>) {
        litchi_opc::Part::set_blob(&mut self.inner, blob);
    }
}

/// Open `bytes` and add a [`CallerPart`].
fn with_caller_part(bytes: &[u8]) -> Result<Package> {
    let mut package = Package::from_vec(bytes.to_vec())?;
    package.opc.try_add_part(Box::new(CallerPart::new(
        b"change-0751 caller-defined payload".to_vec(),
    )))?;
    Ok(package)
}

/// A destination holding a caller-defined part: planned against it, by plan
/// and by durable patch forward and back; and a plan and patch planned against
/// an ordinary destination with the same bytes, applied to it.
fn caller_transcript(label: &str, source_bytes: &[u8], destination_bytes: &[u8]) -> Result<String> {
    let mut transcript = String::new();
    let source = Package::from_vec(source_bytes.to_vec())?;
    let mut custom = with_caller_part(destination_bytes)?;
    let custom_bytes = litchi_opc::PackageWriter::to_bytes(&custom.opc)?;
    writeln!(transcript, "{label}.caller.bytes={}", hex(&custom_bytes)).unwrap();
    let plan = plan_copy(&source, &custom)?;
    record_patch(
        &mut transcript,
        &format!("{label}.caller.plan"),
        plan.patch(),
    )?;
    let durable = CrossSlideCopyPatch::from_bytes(&plan.patch().to_bytes()?)?;
    let outcome = custom.apply_cross_slide_copy_plan(&source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.caller.apply"),
        &mut custom,
        outcome,
    )?;
    let mut patched = with_caller_part(destination_bytes)?;
    let outcome = patched.apply_cross_slide_copy_patch(&source, &durable);
    record_outcome(
        &mut transcript,
        &format!("{label}.caller.patch.forward"),
        &mut patched,
        outcome,
    )?;
    let outcome = patched.apply_cross_slide_copy_patch(&source, &durable.inverse());
    record_outcome(
        &mut transcript,
        &format!("{label}.caller.patch.undo"),
        &mut patched,
        outcome,
    )?;

    let ordinary = Package::from_vec(custom_bytes)?;
    let ordinary_plan = plan_copy(&source, &ordinary)?;
    record_patch(
        &mut transcript,
        &format!("{label}.caller.ordinary.plan"),
        ordinary_plan.patch(),
    )?;
    let mut target = with_caller_part(destination_bytes)?;
    let outcome = target.apply_cross_slide_copy_plan(&source, &ordinary_plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.caller.ordinary.apply"),
        &mut target,
        outcome,
    )?;
    let ordinary_durable = CrossSlideCopyPatch::from_bytes(&ordinary_plan.patch().to_bytes()?)?;
    let mut target = with_caller_part(destination_bytes)?;
    let outcome = target.apply_cross_slide_copy_patch(&source, &ordinary_durable);
    record_outcome(
        &mut transcript,
        &format!("{label}.caller.ordinary.patch"),
        &mut target,
        outcome,
    )?;
    Ok(transcript)
}

/// Size-limit refusals end to end: the durable patch read under limits that
/// admit it but bound the archives, applied; and planning under those limits,
/// applied when it plans.
fn limits_transcript(label: &str, source_bytes: &[u8], destination_bytes: &[u8]) -> Result<String> {
    let mut transcript = String::new();
    let source = Package::from_vec(source_bytes.to_vec())?;
    let destination = Package::from_vec(destination_bytes.to_vec())?;
    let bytes = plan_copy(&source, &destination)?.patch().to_bytes()?;
    for (name, bound) in [
        ("patch", bytes.len()),
        ("source-under", source_bytes.len() - 1),
        ("source-exact", source_bytes.len()),
        ("destination-under", destination_bytes.len() - 1),
        ("destination-exact", destination_bytes.len()),
    ] {
        let limits = Limits::new(4096, bound, 8 * 1024 * 1024, 64, 256 * 1024 * 1024, 1)
            .expect("finite nonzero limits");
        let stage = format!("{label}.limits.{name}");
        writeln!(transcript, "{stage}.bound={bound}").unwrap();
        match CrossSlideCopyPatch::from_bytes_with_limits(&bytes, limits) {
            Ok(patch) => {
                let mut target = Package::from_vec(destination_bytes.to_vec())?;
                let outcome = target.apply_cross_slide_copy_patch(&source, &patch);
                record_outcome(
                    &mut transcript,
                    &format!("{stage}.patch"),
                    &mut target,
                    outcome,
                )?;
            },
            Err(error) => writeln!(transcript, "{stage}.patch.parse.refused={error:?}").unwrap(),
        }
        let planned = source
            .opened_presentation_with_limits(limits)
            .and_then(|source_snapshot| {
                destination
                    .opened_presentation_with_limits(limits)?
                    .plan_cross_slide_copy(
                        &source_snapshot,
                        SOURCE_SLIDE,
                        DESTINATION_SLIDE,
                        POSITION,
                    )
            });
        match planned {
            Ok(plan) => {
                record_patch(&mut transcript, &format!("{stage}.plan"), plan.patch())?;
                let mut target = Package::from_vec(destination_bytes.to_vec())?;
                let outcome = target.apply_cross_slide_copy_plan(&source, &plan);
                record_outcome(
                    &mut transcript,
                    &format!("{stage}.plan.apply"),
                    &mut target,
                    outcome,
                )?;
            },
            Err(error) => writeln!(transcript, "{stage}.plan.refused={error:?}").unwrap(),
        }
    }
    Ok(transcript)
}

fn full_transcript() -> Result<String> {
    let (source, destination) = media_pair()?;
    let mut transcript = pair_transcript("media", &source, &destination)?;
    let (plain_source, plain_destination) = plain_pair()?;
    transcript.push_str(&pair_transcript(
        "plain",
        &plain_source,
        &plain_destination,
    )?);
    transcript.push_str(&caller_transcript("media", &source, &destination)?);
    transcript.push_str(&caller_transcript(
        "plain",
        &plain_source,
        &plain_destination,
    )?);
    transcript.push_str(&limits_transcript("media", &source, &destination)?);
    transcript.push_str(&limits_transcript(
        "plain",
        &plain_source,
        &plain_destination,
    )?);
    Ok(transcript)
}

/// The transcript base `6d989cad63` produces, generated there with
/// `print_the_golden_transcript`: durable patch encodings and their inverses,
/// the six recorded revisions, published archives, returned snapshot
/// revisions and refusals.
const GOLDEN: &str = r#"media.plan.retained=true
media.plan.bytes=4c853fc0835f77df95e345626d80c529505083c2814f5859766929925de007aa
media.plan.inverse=638e792ffdd9e9b9b12a169d400b76f141bf00947e46b8e28ea2de64e200b48a
media.plan.source=18deeee36c7eb857d871c149950b9581180b2dce31dc602f6103622fd0e3dc52
media.plan.destination=3b5d7eeaa2760d75a1f54329cca73475b9b14b67b4afe0c4ce6e836f63fd46a2
media.plan.target=eb7c1a3e673d5fe956bd8ece62baad39cd6c162a97d90c26f6f0feb77e803bb8
media.plan.source_physical=770d39d094ed3fbcad1778de65b02a4ad6ea7c5ef2e30f0abb4fd69576e57462
media.plan.destination_physical=e2dadd323f0031f6322094e50a73ff022cdac94e1bd1a2840688b60cc746879f
media.plan.target_physical=1120b6f2555f43c587c068c62151f0c23308f18ad784d1accedc1205bf85839d
media.plan.transfers=true
media.apply.retained.published=6545174ab647b7e9bc2ed1cd9d032e5196f108b05f73d8b9a7e3e21d964c1771
media.apply.retained.revision=eb7c1a3e673d5fe956bd8ece62baad39cd6c162a97d90c26f6f0feb77e803bb8
media.apply.released.published=6545174ab647b7e9bc2ed1cd9d032e5196f108b05f73d8b9a7e3e21d964c1771
media.apply.released.revision=eb7c1a3e673d5fe956bd8ece62baad39cd6c162a97d90c26f6f0feb77e803bb8
media.apply.warm.published=6545174ab647b7e9bc2ed1cd9d032e5196f108b05f73d8b9a7e3e21d964c1771
media.apply.warm.revision=eb7c1a3e673d5fe956bd8ece62baad39cd6c162a97d90c26f6f0feb77e803bb8
media.patch.forward.published=6545174ab647b7e9bc2ed1cd9d032e5196f108b05f73d8b9a7e3e21d964c1771
media.patch.forward.revision=eb7c1a3e673d5fe956bd8ece62baad39cd6c162a97d90c26f6f0feb77e803bb8
media.patch.undo.published=c1b8ae334472a75240e2590b9925f5fe50387d45e1c623b3d99d16baae91d58d
media.patch.undo.revision=3b5d7eeaa2760d75a1f54329cca73475b9b14b67b4afe0c4ce6e836f63fd46a2
media.patch.redo.published=6545174ab647b7e9bc2ed1cd9d032e5196f108b05f73d8b9a7e3e21d964c1771
media.patch.redo.revision=eb7c1a3e673d5fe956bd8ece62baad39cd6c162a97d90c26f6f0feb77e803bb8
media.patch.undo2.published=c1b8ae334472a75240e2590b9925f5fe50387d45e1c623b3d99d16baae91d58d
media.patch.undo2.revision=3b5d7eeaa2760d75a1f54329cca73475b9b14b67b4afe0c4ce6e836f63fd46a2
media.plan.redo.published=6545174ab647b7e9bc2ed1cd9d032e5196f108b05f73d8b9a7e3e21d964c1771
media.plan.redo.revision=eb7c1a3e673d5fe956bd8ece62baad39cd6c162a97d90c26f6f0feb77e803bb8
media.patch.stale.refused=UnsafeEdit { operation: "apply_cross_slide_copy_patch", reason: "the complete destination package graph differs from the cross-slide patch source" }
media.patch.stale.untouched=6545174ab647b7e9bc2ed1cd9d032e5196f108b05f73d8b9a7e3e21d964c1771
media.stale.destination.refused=UnsafeEdit { operation: "apply_cross_slide_copy_plan", reason: "the complete destination package graph changed after cross-slide planning" }
media.stale.destination.untouched=632b53ac43e445b7bbe04e91b572c73e8279692549fc847bd5e2d81f4e6649ba
media.stale.source.refused=UnsafeEdit { operation: "apply_cross_slide_copy_plan", reason: "the complete source package graph changed after cross-slide planning" }
media.stale.source.untouched=c1b8ae334472a75240e2590b9925f5fe50387d45e1c623b3d99d16baae91d58d
media.foreign.source.refused=UnsafeEdit { operation: "apply_cross_slide_copy_plan", reason: "the complete source package graph changed after cross-slide planning" }
media.foreign.source.untouched=c1b8ae334472a75240e2590b9925f5fe50387d45e1c623b3d99d16baae91d58d
media.stale.physical.refused=UnsafeEdit { operation: "apply_cross_slide_copy_plan", reason: "the serialized destination package changed after cross-slide planning" }
media.stale.physical.untouched=a02526ca7578418779cbb042e51a60e2c4c34927e36e894e42bdb4dbf879fdb2
media.revoked.destination.published=6545174ab647b7e9bc2ed1cd9d032e5196f108b05f73d8b9a7e3e21d964c1771
media.revoked.destination.revision=eb7c1a3e673d5fe956bd8ece62baad39cd6c162a97d90c26f6f0feb77e803bb8
media.revoked.source.bytes=4c853fc0835f77df95e345626d80c529505083c2814f5859766929925de007aa
media.revoked.source.inverse=638e792ffdd9e9b9b12a169d400b76f141bf00947e46b8e28ea2de64e200b48a
media.revoked.source.source=18deeee36c7eb857d871c149950b9581180b2dce31dc602f6103622fd0e3dc52
media.revoked.source.destination=3b5d7eeaa2760d75a1f54329cca73475b9b14b67b4afe0c4ce6e836f63fd46a2
media.revoked.source.target=eb7c1a3e673d5fe956bd8ece62baad39cd6c162a97d90c26f6f0feb77e803bb8
media.revoked.source.source_physical=770d39d094ed3fbcad1778de65b02a4ad6ea7c5ef2e30f0abb4fd69576e57462
media.revoked.source.destination_physical=e2dadd323f0031f6322094e50a73ff022cdac94e1bd1a2840688b60cc746879f
media.revoked.source.target_physical=1120b6f2555f43c587c068c62151f0c23308f18ad784d1accedc1205bf85839d
media.revoked.source.transfers=true
media.revoked.source.apply.published=6545174ab647b7e9bc2ed1cd9d032e5196f108b05f73d8b9a7e3e21d964c1771
media.revoked.source.apply.revision=eb7c1a3e673d5fe956bd8ece62baad39cd6c162a97d90c26f6f0feb77e803bb8
plain.plan.retained=true
plain.plan.bytes=d7c5e42ce4f77af57b9967a6e51e6c0c007ecd85ed9f6cee0a6468c51d3dbe91
plain.plan.inverse=19d31f502e81326ca108be54c474f17f52c23aea1d0d07b549735e87772effda
plain.plan.source=db34b5c259c836a4dfa2a9db172f5778ccc7bbc23bf1fdc7316f5026ad56d980
plain.plan.destination=a2d91ca61d9019c584e33862ae72549b48929004232ac67d61dae3dcf75a633b
plain.plan.target=0efbb26e7c5f17d10013ff269f43d1912a05037d809ab6c7eeab74e0b7e64632
plain.plan.source_physical=427b61c25dd2a2369be5aa553df944b00bac2ff9c04cf9aa1133614e7b114fcc
plain.plan.destination_physical=4c1174b1b954078ef29ee9ae202ac9e81b2223996374bc6843d8d2ac21c54643
plain.plan.target_physical=a34fa32c78bbcc904cfa093a7f8c5a8f60cffda928fce1f73c9f6836e96ad68a
plain.plan.transfers=false
plain.apply.retained.published=15aa2b596658ca2baa30e0b9b2274ed69ee508a2a6da97a82982eb30bc4a530d
plain.apply.retained.revision=0efbb26e7c5f17d10013ff269f43d1912a05037d809ab6c7eeab74e0b7e64632
plain.apply.released.published=15aa2b596658ca2baa30e0b9b2274ed69ee508a2a6da97a82982eb30bc4a530d
plain.apply.released.revision=0efbb26e7c5f17d10013ff269f43d1912a05037d809ab6c7eeab74e0b7e64632
plain.apply.warm.published=15aa2b596658ca2baa30e0b9b2274ed69ee508a2a6da97a82982eb30bc4a530d
plain.apply.warm.revision=0efbb26e7c5f17d10013ff269f43d1912a05037d809ab6c7eeab74e0b7e64632
plain.patch.forward.published=15aa2b596658ca2baa30e0b9b2274ed69ee508a2a6da97a82982eb30bc4a530d
plain.patch.forward.revision=0efbb26e7c5f17d10013ff269f43d1912a05037d809ab6c7eeab74e0b7e64632
plain.patch.undo.published=6cdd00f339051b5f0a35d4809233fd3f64e7cbee86d9524bd8fe748c44722db4
plain.patch.undo.revision=a2d91ca61d9019c584e33862ae72549b48929004232ac67d61dae3dcf75a633b
plain.patch.redo.published=15aa2b596658ca2baa30e0b9b2274ed69ee508a2a6da97a82982eb30bc4a530d
plain.patch.redo.revision=0efbb26e7c5f17d10013ff269f43d1912a05037d809ab6c7eeab74e0b7e64632
plain.patch.undo2.published=6cdd00f339051b5f0a35d4809233fd3f64e7cbee86d9524bd8fe748c44722db4
plain.patch.undo2.revision=a2d91ca61d9019c584e33862ae72549b48929004232ac67d61dae3dcf75a633b
plain.plan.redo.published=15aa2b596658ca2baa30e0b9b2274ed69ee508a2a6da97a82982eb30bc4a530d
plain.plan.redo.revision=0efbb26e7c5f17d10013ff269f43d1912a05037d809ab6c7eeab74e0b7e64632
plain.patch.stale.refused=UnsafeEdit { operation: "apply_cross_slide_copy_patch", reason: "the complete destination package graph differs from the cross-slide patch source" }
plain.patch.stale.untouched=15aa2b596658ca2baa30e0b9b2274ed69ee508a2a6da97a82982eb30bc4a530d
plain.stale.destination.refused=UnsafeEdit { operation: "apply_cross_slide_copy_plan", reason: "the complete destination package graph changed after cross-slide planning" }
plain.stale.destination.untouched=216a1dd43508078d378f461986537efed6c1ed31f874d67bdd1bc469715b3d9b
plain.stale.source.refused=UnsafeEdit { operation: "apply_cross_slide_copy_plan", reason: "the complete source package graph changed after cross-slide planning" }
plain.stale.source.untouched=6cdd00f339051b5f0a35d4809233fd3f64e7cbee86d9524bd8fe748c44722db4
plain.foreign.source.refused=UnsafeEdit { operation: "apply_cross_slide_copy_plan", reason: "the complete source package graph changed after cross-slide planning" }
plain.foreign.source.untouched=6cdd00f339051b5f0a35d4809233fd3f64e7cbee86d9524bd8fe748c44722db4
plain.stale.physical.refused=UnsafeEdit { operation: "apply_cross_slide_copy_plan", reason: "the serialized destination package changed after cross-slide planning" }
plain.stale.physical.untouched=c7e61fcbbd418b98d29d47aed30a6e8202e697b4443440bf648e2f8501c83a7d
plain.revoked.destination.published=15aa2b596658ca2baa30e0b9b2274ed69ee508a2a6da97a82982eb30bc4a530d
plain.revoked.destination.revision=0efbb26e7c5f17d10013ff269f43d1912a05037d809ab6c7eeab74e0b7e64632
plain.revoked.source.bytes=d7c5e42ce4f77af57b9967a6e51e6c0c007ecd85ed9f6cee0a6468c51d3dbe91
plain.revoked.source.inverse=19d31f502e81326ca108be54c474f17f52c23aea1d0d07b549735e87772effda
plain.revoked.source.source=db34b5c259c836a4dfa2a9db172f5778ccc7bbc23bf1fdc7316f5026ad56d980
plain.revoked.source.destination=a2d91ca61d9019c584e33862ae72549b48929004232ac67d61dae3dcf75a633b
plain.revoked.source.target=0efbb26e7c5f17d10013ff269f43d1912a05037d809ab6c7eeab74e0b7e64632
plain.revoked.source.source_physical=427b61c25dd2a2369be5aa553df944b00bac2ff9c04cf9aa1133614e7b114fcc
plain.revoked.source.destination_physical=4c1174b1b954078ef29ee9ae202ac9e81b2223996374bc6843d8d2ac21c54643
plain.revoked.source.target_physical=a34fa32c78bbcc904cfa093a7f8c5a8f60cffda928fce1f73c9f6836e96ad68a
plain.revoked.source.transfers=false
plain.revoked.source.apply.published=15aa2b596658ca2baa30e0b9b2274ed69ee508a2a6da97a82982eb30bc4a530d
plain.revoked.source.apply.revision=0efbb26e7c5f17d10013ff269f43d1912a05037d809ab6c7eeab74e0b7e64632
media.caller.bytes=0f9c088a35ff12285497a2c55156f57712059c4941f8405f26b2f7f0c7ba98f0
media.caller.plan.bytes=4a73d4d428e226ce2ebd05edf9f2747dc7d7c7050c70408ef03a83ca8dffdaae
media.caller.plan.inverse=559c4f2efb14ce8f892964ffa6d043a19db8a56bb65bf0f24e502db06b5c0fe6
media.caller.plan.source=18deeee36c7eb857d871c149950b9581180b2dce31dc602f6103622fd0e3dc52
media.caller.plan.destination=92e724ea72b99f612be92ba883dd1d91d0c8eba178389b81eb35a962f6e707d9
media.caller.plan.target=44594b917c833377b04450c3daad7d4db66b57fb828f3b1802a42b6b7a8222fa
media.caller.plan.source_physical=770d39d094ed3fbcad1778de65b02a4ad6ea7c5ef2e30f0abb4fd69576e57462
media.caller.plan.destination_physical=14512ef6686232f1459e79c19c8f0dfd259ad2bbc653b699020435cfc512b30f
media.caller.plan.target_physical=4e9979a6efe2c6dadc1617984342651c125337f2ca4f17bda6614bef2c7886d4
media.caller.plan.transfers=false
media.caller.apply.published=b0bc10be4a18d1f80c32a4955d4a5dcfad41d400e01d2812edbfc9b72e53d331
media.caller.apply.revision=44594b917c833377b04450c3daad7d4db66b57fb828f3b1802a42b6b7a8222fa
media.caller.patch.forward.published=b0bc10be4a18d1f80c32a4955d4a5dcfad41d400e01d2812edbfc9b72e53d331
media.caller.patch.forward.revision=44594b917c833377b04450c3daad7d4db66b57fb828f3b1802a42b6b7a8222fa
media.caller.patch.undo.published=0f9c088a35ff12285497a2c55156f57712059c4941f8405f26b2f7f0c7ba98f0
media.caller.patch.undo.revision=92e724ea72b99f612be92ba883dd1d91d0c8eba178389b81eb35a962f6e707d9
media.caller.ordinary.plan.bytes=32f4519cc30e207157f9888a2dbd0405d0787e44a4d7eea9dcb1df55ce4720d8
media.caller.ordinary.plan.inverse=93d56a404c612d461711663438d3cb1cd63cfac3a04501b4ee1f125ef709b688
media.caller.ordinary.plan.source=18deeee36c7eb857d871c149950b9581180b2dce31dc602f6103622fd0e3dc52
media.caller.ordinary.plan.destination=92e724ea72b99f612be92ba883dd1d91d0c8eba178389b81eb35a962f6e707d9
media.caller.ordinary.plan.target=44594b917c833377b04450c3daad7d4db66b57fb828f3b1802a42b6b7a8222fa
media.caller.ordinary.plan.source_physical=770d39d094ed3fbcad1778de65b02a4ad6ea7c5ef2e30f0abb4fd69576e57462
media.caller.ordinary.plan.destination_physical=14512ef6686232f1459e79c19c8f0dfd259ad2bbc653b699020435cfc512b30f
media.caller.ordinary.plan.target_physical=71b1578456b8b11bc7a674b5ae91864cd56e3a9c4c9aa0fc7305292de12edc62
media.caller.ordinary.plan.transfers=true
media.caller.ordinary.apply.refused=SlideCopyPlan { kind: CallerDefinedPart, detail: "a cross-slide copy that frames source-compressed media publishes the reopen of the destination's bytes, which cannot keep its caller-defined parts; plan the copy against this destination to record the recompressing route" }
media.caller.ordinary.apply.untouched=0f9c088a35ff12285497a2c55156f57712059c4941f8405f26b2f7f0c7ba98f0
media.caller.ordinary.patch.refused=SlideCopyPlan { kind: CallerDefinedPart, detail: "a cross-slide copy that frames source-compressed media publishes the reopen of the destination's bytes, which cannot keep its caller-defined parts; plan the copy against this destination to record the recompressing route" }
media.caller.ordinary.patch.untouched=0f9c088a35ff12285497a2c55156f57712059c4941f8405f26b2f7f0c7ba98f0
plain.caller.bytes=0c13f8aaed85234850be3ba4c2e222d4345aabd98e71ce246ece2040816b7d5a
plain.caller.plan.bytes=ff2463aaa8651275719ef8c22553c14583f5ecf984aafe4078bc88432d0f2c43
plain.caller.plan.inverse=dcaab63bf885a4f33d5c938c0c2cc9b8486fc02a08c8a38256a9ef028f4114c7
plain.caller.plan.source=db34b5c259c836a4dfa2a9db172f5778ccc7bbc23bf1fdc7316f5026ad56d980
plain.caller.plan.destination=3dfaedf12382c76c8fecec10161dd55b2ed0f5333ebb952832dacb4673c4e315
plain.caller.plan.target=67cbaa3138287f6900b57dad783134da7fc89929d9316f54a5aba1463c47307f
plain.caller.plan.source_physical=427b61c25dd2a2369be5aa553df944b00bac2ff9c04cf9aa1133614e7b114fcc
plain.caller.plan.destination_physical=b483cc3356c9edc685e141cd164e39f00d00d5798c0b8dd6001074892a60ab24
plain.caller.plan.target_physical=68eb01c3613dc2b0d582f1d74b2ac8261944dae237e654781ae3b14716d97957
plain.caller.plan.transfers=false
plain.caller.apply.published=859d1f66d9aa849f86d89e27d2b3018fa24125f7024b34b489254c1a93c55d9c
plain.caller.apply.revision=67cbaa3138287f6900b57dad783134da7fc89929d9316f54a5aba1463c47307f
plain.caller.patch.forward.published=859d1f66d9aa849f86d89e27d2b3018fa24125f7024b34b489254c1a93c55d9c
plain.caller.patch.forward.revision=67cbaa3138287f6900b57dad783134da7fc89929d9316f54a5aba1463c47307f
plain.caller.patch.undo.published=0c13f8aaed85234850be3ba4c2e222d4345aabd98e71ce246ece2040816b7d5a
plain.caller.patch.undo.revision=3dfaedf12382c76c8fecec10161dd55b2ed0f5333ebb952832dacb4673c4e315
plain.caller.ordinary.plan.bytes=ff2463aaa8651275719ef8c22553c14583f5ecf984aafe4078bc88432d0f2c43
plain.caller.ordinary.plan.inverse=dcaab63bf885a4f33d5c938c0c2cc9b8486fc02a08c8a38256a9ef028f4114c7
plain.caller.ordinary.plan.source=db34b5c259c836a4dfa2a9db172f5778ccc7bbc23bf1fdc7316f5026ad56d980
plain.caller.ordinary.plan.destination=3dfaedf12382c76c8fecec10161dd55b2ed0f5333ebb952832dacb4673c4e315
plain.caller.ordinary.plan.target=67cbaa3138287f6900b57dad783134da7fc89929d9316f54a5aba1463c47307f
plain.caller.ordinary.plan.source_physical=427b61c25dd2a2369be5aa553df944b00bac2ff9c04cf9aa1133614e7b114fcc
plain.caller.ordinary.plan.destination_physical=b483cc3356c9edc685e141cd164e39f00d00d5798c0b8dd6001074892a60ab24
plain.caller.ordinary.plan.target_physical=68eb01c3613dc2b0d582f1d74b2ac8261944dae237e654781ae3b14716d97957
plain.caller.ordinary.plan.transfers=false
plain.caller.ordinary.apply.published=859d1f66d9aa849f86d89e27d2b3018fa24125f7024b34b489254c1a93c55d9c
plain.caller.ordinary.apply.revision=67cbaa3138287f6900b57dad783134da7fc89929d9316f54a5aba1463c47307f
plain.caller.ordinary.patch.published=859d1f66d9aa849f86d89e27d2b3018fa24125f7024b34b489254c1a93c55d9c
plain.caller.ordinary.patch.revision=67cbaa3138287f6900b57dad783134da7fc89929d9316f54a5aba1463c47307f
media.limits.patch.bound=71855
media.limits.patch.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 71855 }
media.limits.patch.patch.untouched=c1b8ae334472a75240e2590b9925f5fe50387d45e1c623b3d99d16baae91d58d
media.limits.patch.plan.refused=Limit { resource: "cross-slide candidate patch bytes", limit: 71855 }
media.limits.source-under.bound=71942
media.limits.source-under.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 71942 }
media.limits.source-under.patch.untouched=c1b8ae334472a75240e2590b9925f5fe50387d45e1c623b3d99d16baae91d58d
media.limits.source-under.plan.refused=Limit { resource: "cross-slide candidate patch bytes", limit: 71942 }
media.limits.source-exact.bound=71943
media.limits.source-exact.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 71943 }
media.limits.source-exact.patch.untouched=c1b8ae334472a75240e2590b9925f5fe50387d45e1c623b3d99d16baae91d58d
media.limits.source-exact.plan.refused=Limit { resource: "cross-slide candidate patch bytes", limit: 71943 }
media.limits.destination-under.bound=72047
media.limits.destination-under.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 72047 }
media.limits.destination-under.patch.untouched=c1b8ae334472a75240e2590b9925f5fe50387d45e1c623b3d99d16baae91d58d
media.limits.destination-under.plan.refused=Limit { resource: "cross-slide candidate patch bytes", limit: 72047 }
media.limits.destination-exact.bound=72048
media.limits.destination-exact.patch.refused=UnsafeEdit { operation: "apply_cross_slide_copy_patch", reason: "the durable cross-slide patch does not match a freshly proven candidate" }
media.limits.destination-exact.patch.untouched=c1b8ae334472a75240e2590b9925f5fe50387d45e1c623b3d99d16baae91d58d
media.limits.destination-exact.plan.refused=Limit { resource: "cross-slide candidate patch bytes", limit: 72048 }
plain.limits.patch.bound=4895
plain.limits.patch.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 4895 }
plain.limits.patch.patch.untouched=6cdd00f339051b5f0a35d4809233fd3f64e7cbee86d9524bd8fe748c44722db4
plain.limits.patch.plan.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 4895 }
plain.limits.source-under.bound=31374
plain.limits.source-under.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 31374 }
plain.limits.source-under.patch.untouched=6cdd00f339051b5f0a35d4809233fd3f64e7cbee86d9524bd8fe748c44722db4
plain.limits.source-under.plan.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 31374 }
plain.limits.source-exact.bound=31375
plain.limits.source-exact.patch.refused=UnsafeEdit { operation: "apply_cross_slide_copy_patch", reason: "the durable cross-slide patch does not match a freshly proven candidate" }
plain.limits.source-exact.patch.untouched=6cdd00f339051b5f0a35d4809233fd3f64e7cbee86d9524bd8fe748c44722db4
plain.limits.source-exact.plan.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 31375 }
plain.limits.destination-under.bound=30460
plain.limits.destination-under.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 30460 }
plain.limits.destination-under.patch.untouched=6cdd00f339051b5f0a35d4809233fd3f64e7cbee86d9524bd8fe748c44722db4
plain.limits.destination-under.plan.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 30460 }
plain.limits.destination-exact.bound=30461
plain.limits.destination-exact.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 30461 }
plain.limits.destination-exact.patch.untouched=6cdd00f339051b5f0a35d4809233fd3f64e7cbee86d9524bd8fe748c44722db4
plain.limits.destination-exact.plan.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 30461 }
"#;

/// Every durable patch byte, recorded revision, published byte, snapshot
/// revision and refusal of the scenarios equals the base's. Debug and test
/// builds also re-derive every memoized digest these routes reuse.
#[test]
fn every_patch_revision_output_and_refusal_matches_the_base() -> Result<()> {
    let transcript = full_transcript()?;
    let actual: Vec<&str> = transcript.lines().collect();
    let expected: Vec<&str> = GOLDEN.lines().collect();
    for (index, (actual, expected)) in actual.iter().zip(&expected).enumerate() {
        assert_eq!(
            actual, expected,
            "line {index} of the golden transcript moved"
        );
    }
    assert_eq!(
        actual.len(),
        expected.len(),
        "the golden transcript changed length"
    );
    Ok(())
}

/// Print the transcript, for regenerating `GOLDEN` on the base tree:
/// `cargo test -p litchi-pptx --lib digest_reuse_tests::print -- --ignored --nocapture`.
#[test]
#[ignore = "prints the golden transcript; run explicitly"]
fn print_the_golden_transcript() -> Result<()> {
    println!("{}", full_transcript()?);
    Ok(())
}
