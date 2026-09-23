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

fn full_transcript() -> Result<String> {
    let (source, destination) = media_pair()?;
    let mut transcript = pair_transcript("media", &source, &destination)?;
    let (source, destination) = plain_pair()?;
    transcript.push_str(&pair_transcript("plain", &source, &destination)?);
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
