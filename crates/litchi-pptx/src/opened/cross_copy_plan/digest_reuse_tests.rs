//! Byte identity of the owned cross-copy's durable patches, revisions,
//! published archives and refusals under digest reuse.
//!
//! The owned cross-copy does not re-hash bytes whose identity is already
//! proven: the application's live revisions consult the facade's payload-digest
//! memos, the candidate captures consult the snapshots' memos, an unmodified
//! owned source's physical revision is sealed from the digest `litchi-opc`
//! binds to its archive, and an application shares a plan's retained archive
//! instead of copying it. None of that may move a value. This module records,
//! for a set of deterministic scenarios, the SHA-256 of every durable patch
//! encoding, the six recorded revisions, every published archive and returned
//! snapshot revision, and the text of every refusal, and compares the
//! transcript with the one the encoder before `LPTX0003` produces. The module
//! also compiles against that earlier encoder, which is how `GOLDEN` was
//! generated.
//!
//! The transcript also covers a destination holding a caller-defined part, by
//! plan and by durable patch in both directions, and size-limit refusals end to
//! end: durable patches read under limits that admit the patch but not the
//! source or destination archive, and planning under the same limits.

use std::fmt::Write as _;

use sha2::{Digest, Sha256};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

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
    let mut builder = crate::builder::Builder::new()?;
    let presentation = builder.presentation()?;
    for index in 0..slides {
        presentation
            .add_slide()?
            .set_title(&format!("digest-reuse-{prefix}-{index}"));
    }
    let package = builder.build()?;
    package.to_bytes()
}

fn with_pictures(bytes: &[u8], slide: usize, pictures: &[(&str, &str, &[u8])]) -> Result<Vec<u8>> {
    let mut package = Package::from_bytes(bytes)?;
    let mut edit = package.edit()?;
    for (index, (part_name, content_type, data)) in pictures.iter().enumerate() {
        edit.add_picture(
            slide,
            format!("digest-reuse-picture-{index}"),
            &Resource::new(*part_name, *content_type, data.to_vec()),
            (800, 800, 72, 72),
        )?;
    }
    let commit = edit.commit()?;
    package = commit.into_package();
    package.to_bytes()
}

/// Rewrite every member through a fresh writer, storing the members `stored`
/// selects and deflating the rest with known sizes.
fn repack(bytes: &[u8], stored: impl Fn(&str) -> bool) -> Vec<u8> {
    let reader = ArchiveReader::new(bytes, soapberry_zip::office::ArchiveLimits::default())
        .expect("fixture archive");
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
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

fn plan_copy(source: &Package, destination: &Package) -> Result<crate::cross_slide_copy::Plan> {
    destination.opened_presentation()?.plan_cross_slide_copy(
        &source.opened_presentation()?,
        SOURCE_SLIDE,
        DESTINATION_SLIDE,
        POSITION,
    )
}

fn record_patch(
    transcript: &mut String,
    label: &str,
    patch: &crate::cross_slide_copy::Patch,
) -> Result<()> {
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
    let source = Package::from_reader(std::io::Cursor::new(source_bytes.to_vec()))?;
    let destination = Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
    let plan = plan_copy(&source, &destination)?;
    writeln!(
        transcript,
        "{label}.plan.retained={}",
        plan.retained_candidate_bytes().is_some()
    )
    .unwrap();
    record_patch(&mut transcript, &format!("{label}.plan"), plan.patch())?;
    let durable =
        crate::cross_slide_copy::Patch::from_bytes(&plan.patch().to_bytes()?, Limits::default())?;

    // The plan, retained and released, into the destination it was planned
    // against and into a fresh open of the same bytes.
    let mut retained = Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
    let outcome = retained.apply_cross_slide_copy_plan(&source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.apply.retained"),
        &mut retained,
        outcome,
    )?;
    let mut released_plan = plan.clone();
    released_plan.release_retained_candidate();
    let mut released = Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
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
    let mut patched = Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
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
    let mut moved = Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
    {
        let mut edit = moved.edit()?;
        edit.set_shape_text(0, crate::shape::Key::Index(0), "digest-reuse moved")?;
        let commit = edit.commit()?;
        moved = commit.into_package();
    }
    let outcome = moved.apply_cross_slide_copy_plan(&source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.stale.destination"),
        &mut moved,
        outcome,
    )?;

    // A source that moved after planning, and a foreign source, are refused.
    let mut moved_source = Package::from_reader(std::io::Cursor::new(source_bytes.to_vec()))?;
    {
        let mut edit = moved_source.edit()?;
        edit.set_shape_text(0, crate::shape::Key::Index(0), "digest-reuse moved source")?;
        let commit = edit.commit()?;
        moved_source = commit.into_package();
    }
    let mut target = Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
    let outcome = target.apply_cross_slide_copy_plan(&moved_source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.stale.source"),
        &mut target,
        outcome,
    )?;
    let foreign = Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
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
    let mut physical = Package::from_reader(std::io::Cursor::new(repacked))?;
    let outcome = physical.apply_cross_slide_copy_plan(&source, &plan);
    record_outcome(
        &mut transcript,
        &format!("{label}.stale.physical"),
        &mut physical,
        outcome,
    )?;

    // A destination whose exact-source authorization an edit revoked while
    // keeping its bytes accepts the copy and publishes the same output.
    let mut revoked = Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
    revoked.edit_opc_in_place(|opc| {
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
    let mut resaved_source = Package::from_reader(std::io::Cursor::new(source_bytes.to_vec()))?;
    resaved_source.edit_opc_in_place(|opc| {
        opc.set_save_options(litchi_opc::SaveOptions::default());
        Ok(())
    })?;
    let fresh = Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
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
                litchi_opc::PackURI::new("/custom/digest-reuse.bin").expect("part name"),
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
    let mut package = Package::from_reader(std::io::Cursor::new(bytes.to_vec()))?;
    std::sync::Arc::make_mut(&mut package.opc).try_add_part(Box::new(CallerPart::new(
        b"digest-reuse caller-defined payload".to_vec(),
    )))?;
    Ok(package)
}

/// A destination holding a caller-defined part: planned against it, by plan
/// and by durable patch forward and back; and a plan and patch planned against
/// an ordinary destination with the same bytes, applied to it.
fn caller_transcript(label: &str, source_bytes: &[u8], destination_bytes: &[u8]) -> Result<String> {
    let mut transcript = String::new();
    let source = Package::from_reader(std::io::Cursor::new(source_bytes.to_vec()))?;
    let mut custom = with_caller_part(destination_bytes)?;
    let custom_bytes = litchi_opc::PackageWriter::to_bytes(&custom.opc)?;
    writeln!(transcript, "{label}.caller.bytes={}", hex(&custom_bytes)).unwrap();
    let plan = plan_copy(&source, &custom)?;
    record_patch(
        &mut transcript,
        &format!("{label}.caller.plan"),
        plan.patch(),
    )?;
    let durable =
        crate::cross_slide_copy::Patch::from_bytes(&plan.patch().to_bytes()?, Limits::default())?;
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

    let ordinary = Package::from_reader(std::io::Cursor::new(custom_bytes))?;
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
    let ordinary_durable = crate::cross_slide_copy::Patch::from_bytes(
        &ordinary_plan.patch().to_bytes()?,
        Limits::default(),
    )?;
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
    let source = Package::from_reader(std::io::Cursor::new(source_bytes.to_vec()))?;
    let destination = Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
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
        match crate::cross_slide_copy::Patch::from_bytes(&bytes, limits) {
            Ok(patch) => {
                let mut target =
                    Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
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
                let mut target =
                    Package::from_reader(std::io::Cursor::new(destination_bytes.to_vec()))?;
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

/// The transcript the encoder before `LPTX0003` produces, generated with
/// `print_the_golden_transcript` against that encoder: durable patch
/// encodings and their inverses, the six recorded revisions, published
/// archives, returned snapshot revisions and refusals.
const GOLDEN: &str = r#"media.plan.retained=true
media.plan.bytes=f7dadf294a783c8fcf2667cd546f02f8bbaca14b4814e380f3ad0e34a4b4036b
media.plan.inverse=7868b2e102151ad2aeb1c9699a4a4a399b3d18fd17adb9045f7c3cbe6ed745f6
media.plan.source=07024199f93cdd1b6e087e7cd886e7763ec6a75ae8f7ea450531f3f165344937
media.plan.destination=d653dce73a5814ce4229555034d55d8e249edf7b710f33bf14d3cfb14034e54e
media.plan.target=65422a917be6da8fb13b90e4ae7e1da81617761693222a4535f11dc7bc586b74
media.plan.source_physical=88c6b1a49bad6800952511209d613349b5359a2322b8d483036503432e6c8704
media.plan.destination_physical=808ea58f02c24b1e6ba86c74abfb77fd0a034121fccfb307fb745e25d3c7afa1
media.plan.target_physical=22b72f40b2dd3291ec4e1c7d77719a6e11b4b6b3c335f113027e912ca5b2f36c
media.plan.transfers=true
media.apply.retained.published=7ca80a4c06c8faba374152f4eeac2111788f7981d419f238da74edf848fb2d35
media.apply.retained.revision=65422a917be6da8fb13b90e4ae7e1da81617761693222a4535f11dc7bc586b74
media.apply.released.published=7ca80a4c06c8faba374152f4eeac2111788f7981d419f238da74edf848fb2d35
media.apply.released.revision=65422a917be6da8fb13b90e4ae7e1da81617761693222a4535f11dc7bc586b74
media.apply.warm.published=7ca80a4c06c8faba374152f4eeac2111788f7981d419f238da74edf848fb2d35
media.apply.warm.revision=65422a917be6da8fb13b90e4ae7e1da81617761693222a4535f11dc7bc586b74
media.patch.forward.published=7ca80a4c06c8faba374152f4eeac2111788f7981d419f238da74edf848fb2d35
media.patch.forward.revision=65422a917be6da8fb13b90e4ae7e1da81617761693222a4535f11dc7bc586b74
media.patch.undo.published=d1dc2a4b199a0c74955a07fc4514c26066fc97b1205b630ebd727d390f858912
media.patch.undo.revision=d653dce73a5814ce4229555034d55d8e249edf7b710f33bf14d3cfb14034e54e
media.patch.redo.published=7ca80a4c06c8faba374152f4eeac2111788f7981d419f238da74edf848fb2d35
media.patch.redo.revision=65422a917be6da8fb13b90e4ae7e1da81617761693222a4535f11dc7bc586b74
media.patch.undo2.published=d1dc2a4b199a0c74955a07fc4514c26066fc97b1205b630ebd727d390f858912
media.patch.undo2.revision=d653dce73a5814ce4229555034d55d8e249edf7b710f33bf14d3cfb14034e54e
media.plan.redo.published=7ca80a4c06c8faba374152f4eeac2111788f7981d419f238da74edf848fb2d35
media.plan.redo.revision=65422a917be6da8fb13b90e4ae7e1da81617761693222a4535f11dc7bc586b74
media.patch.stale.refused=PatchConflict
media.patch.stale.untouched=7ca80a4c06c8faba374152f4eeac2111788f7981d419f238da74edf848fb2d35
media.stale.destination.refused=PatchConflict
media.stale.destination.untouched=676c6b2fb9fcad99d9fa10beb5efff9fb6776bd203d67a8045f9c28ba97a7cb1
media.stale.source.refused=PatchConflict
media.stale.source.untouched=d1dc2a4b199a0c74955a07fc4514c26066fc97b1205b630ebd727d390f858912
media.foreign.source.refused=PatchConflict
media.foreign.source.untouched=d1dc2a4b199a0c74955a07fc4514c26066fc97b1205b630ebd727d390f858912
media.stale.physical.refused=PatchConflict
media.stale.physical.untouched=a2f90125a208327ec8e997166723800256682e662fb5b5c9453f7a1bcb7c08e4
media.revoked.destination.published=7ca80a4c06c8faba374152f4eeac2111788f7981d419f238da74edf848fb2d35
media.revoked.destination.revision=65422a917be6da8fb13b90e4ae7e1da81617761693222a4535f11dc7bc586b74
media.revoked.source.bytes=f7dadf294a783c8fcf2667cd546f02f8bbaca14b4814e380f3ad0e34a4b4036b
media.revoked.source.inverse=7868b2e102151ad2aeb1c9699a4a4a399b3d18fd17adb9045f7c3cbe6ed745f6
media.revoked.source.source=07024199f93cdd1b6e087e7cd886e7763ec6a75ae8f7ea450531f3f165344937
media.revoked.source.destination=d653dce73a5814ce4229555034d55d8e249edf7b710f33bf14d3cfb14034e54e
media.revoked.source.target=65422a917be6da8fb13b90e4ae7e1da81617761693222a4535f11dc7bc586b74
media.revoked.source.source_physical=88c6b1a49bad6800952511209d613349b5359a2322b8d483036503432e6c8704
media.revoked.source.destination_physical=808ea58f02c24b1e6ba86c74abfb77fd0a034121fccfb307fb745e25d3c7afa1
media.revoked.source.target_physical=22b72f40b2dd3291ec4e1c7d77719a6e11b4b6b3c335f113027e912ca5b2f36c
media.revoked.source.transfers=true
media.revoked.source.apply.published=7ca80a4c06c8faba374152f4eeac2111788f7981d419f238da74edf848fb2d35
media.revoked.source.apply.revision=65422a917be6da8fb13b90e4ae7e1da81617761693222a4535f11dc7bc586b74
plain.plan.retained=true
plain.plan.bytes=f5f578890dced63fc248231f4ee899b797ce4329aca5cf67aaab058cfea49600
plain.plan.inverse=d91394a6668bc9321835c425f27a35bc2a652b1453b85a6b38cc7e1146a6107d
plain.plan.source=820f57f2cc9b2e45f3d06aaa1341994c24afee6824d94f77fb3079fabb0fa3c1
plain.plan.destination=56b79dd049d157806c5e061bc53a1d90d01c12274644e66ae5075bbb34ad04d1
plain.plan.target=0e451a267a499a46a70abc418cfd81e0ec92a4a568f1519a67e4346e1ad5107e
plain.plan.source_physical=4abfb1a5b7b2648f1343f16bbef7fcd494a0100e95743d8ce34bcf289d07c9e6
plain.plan.destination_physical=d91fbc1ec1e3894678ff5ef90f042171e6f2e8f78e8c9edf9bc3ebd4e7d7ca09
plain.plan.target_physical=bce9204f5117cdc642fd1e1889db03fa0dd8bda4a78e615ababd31b4fdee7a98
plain.plan.transfers=false
plain.apply.retained.published=935a9f282bbbef8f0aa42bd14cc1d3b5a42d975b1f9b36220a0fef7c630873f9
plain.apply.retained.revision=0e451a267a499a46a70abc418cfd81e0ec92a4a568f1519a67e4346e1ad5107e
plain.apply.released.published=935a9f282bbbef8f0aa42bd14cc1d3b5a42d975b1f9b36220a0fef7c630873f9
plain.apply.released.revision=0e451a267a499a46a70abc418cfd81e0ec92a4a568f1519a67e4346e1ad5107e
plain.apply.warm.published=935a9f282bbbef8f0aa42bd14cc1d3b5a42d975b1f9b36220a0fef7c630873f9
plain.apply.warm.revision=0e451a267a499a46a70abc418cfd81e0ec92a4a568f1519a67e4346e1ad5107e
plain.patch.forward.published=935a9f282bbbef8f0aa42bd14cc1d3b5a42d975b1f9b36220a0fef7c630873f9
plain.patch.forward.revision=0e451a267a499a46a70abc418cfd81e0ec92a4a568f1519a67e4346e1ad5107e
plain.patch.undo.published=e0beb28b70e0f02a41a449f5341d199e36417a5a2e85664f6f46e57eb80bc409
plain.patch.undo.revision=56b79dd049d157806c5e061bc53a1d90d01c12274644e66ae5075bbb34ad04d1
plain.patch.redo.published=935a9f282bbbef8f0aa42bd14cc1d3b5a42d975b1f9b36220a0fef7c630873f9
plain.patch.redo.revision=0e451a267a499a46a70abc418cfd81e0ec92a4a568f1519a67e4346e1ad5107e
plain.patch.undo2.published=e0beb28b70e0f02a41a449f5341d199e36417a5a2e85664f6f46e57eb80bc409
plain.patch.undo2.revision=56b79dd049d157806c5e061bc53a1d90d01c12274644e66ae5075bbb34ad04d1
plain.plan.redo.published=935a9f282bbbef8f0aa42bd14cc1d3b5a42d975b1f9b36220a0fef7c630873f9
plain.plan.redo.revision=0e451a267a499a46a70abc418cfd81e0ec92a4a568f1519a67e4346e1ad5107e
plain.patch.stale.refused=PatchConflict
plain.patch.stale.untouched=935a9f282bbbef8f0aa42bd14cc1d3b5a42d975b1f9b36220a0fef7c630873f9
plain.stale.destination.refused=PatchConflict
plain.stale.destination.untouched=c938cb3aa0b9104cf73b6bf4814e024a3210618620a104cb16d6a07b5365e532
plain.stale.source.refused=PatchConflict
plain.stale.source.untouched=e0beb28b70e0f02a41a449f5341d199e36417a5a2e85664f6f46e57eb80bc409
plain.foreign.source.refused=PatchConflict
plain.foreign.source.untouched=e0beb28b70e0f02a41a449f5341d199e36417a5a2e85664f6f46e57eb80bc409
plain.stale.physical.refused=PatchConflict
plain.stale.physical.untouched=e9187aa594838555d86adf8e58869815846e89ec3a116c47814599391d041d8b
plain.revoked.destination.published=935a9f282bbbef8f0aa42bd14cc1d3b5a42d975b1f9b36220a0fef7c630873f9
plain.revoked.destination.revision=0e451a267a499a46a70abc418cfd81e0ec92a4a568f1519a67e4346e1ad5107e
plain.revoked.source.bytes=f5f578890dced63fc248231f4ee899b797ce4329aca5cf67aaab058cfea49600
plain.revoked.source.inverse=d91394a6668bc9321835c425f27a35bc2a652b1453b85a6b38cc7e1146a6107d
plain.revoked.source.source=820f57f2cc9b2e45f3d06aaa1341994c24afee6824d94f77fb3079fabb0fa3c1
plain.revoked.source.destination=56b79dd049d157806c5e061bc53a1d90d01c12274644e66ae5075bbb34ad04d1
plain.revoked.source.target=0e451a267a499a46a70abc418cfd81e0ec92a4a568f1519a67e4346e1ad5107e
plain.revoked.source.source_physical=4abfb1a5b7b2648f1343f16bbef7fcd494a0100e95743d8ce34bcf289d07c9e6
plain.revoked.source.destination_physical=d91fbc1ec1e3894678ff5ef90f042171e6f2e8f78e8c9edf9bc3ebd4e7d7ca09
plain.revoked.source.target_physical=bce9204f5117cdc642fd1e1889db03fa0dd8bda4a78e615ababd31b4fdee7a98
plain.revoked.source.transfers=false
plain.revoked.source.apply.published=935a9f282bbbef8f0aa42bd14cc1d3b5a42d975b1f9b36220a0fef7c630873f9
plain.revoked.source.apply.revision=0e451a267a499a46a70abc418cfd81e0ec92a4a568f1519a67e4346e1ad5107e
media.caller.bytes=c45268167a26c6d6e2b1ad6ca988c31770f684cf4835f4c1fa00a2cf94f4f2b1
media.caller.plan.bytes=8b930ccfa6136caa3cff819fd34054751c27dee6befcef985e17963c48b476c5
media.caller.plan.inverse=67ccdd352dfc62b637c36b40e4340f36c2483d4c92c4b41ff10683afa05dc832
media.caller.plan.source=07024199f93cdd1b6e087e7cd886e7763ec6a75ae8f7ea450531f3f165344937
media.caller.plan.destination=24e41b192c86fca83f215298c2436748c67be1858550466268ef9814ee400db9
media.caller.plan.target=d4204f706cdf171ebf3531237e3eef184953e21e7a00c639ea73a7446591582a
media.caller.plan.source_physical=88c6b1a49bad6800952511209d613349b5359a2322b8d483036503432e6c8704
media.caller.plan.destination_physical=78f26cda5979a1a843d8cb06c0c24370ff0039c521f7bcbe246efcfbebb10fdb
media.caller.plan.target_physical=caf9437a66b27ddb2f82be67554b4a6cdc42b5124ccbc4e6de39b1f4ca7273be
media.caller.plan.transfers=false
media.caller.apply.published=4ae43939a419321c780a4ab91fdde467da014019e3bfb99b2b6c79052389a24f
media.caller.apply.revision=d4204f706cdf171ebf3531237e3eef184953e21e7a00c639ea73a7446591582a
media.caller.patch.forward.published=4ae43939a419321c780a4ab91fdde467da014019e3bfb99b2b6c79052389a24f
media.caller.patch.forward.revision=d4204f706cdf171ebf3531237e3eef184953e21e7a00c639ea73a7446591582a
media.caller.patch.undo.published=c45268167a26c6d6e2b1ad6ca988c31770f684cf4835f4c1fa00a2cf94f4f2b1
media.caller.patch.undo.revision=24e41b192c86fca83f215298c2436748c67be1858550466268ef9814ee400db9
media.caller.ordinary.plan.bytes=28f7558a6e2dc2a9a72e9e0220a560cb44ac745670f040b1b1a9719534de9a42
media.caller.ordinary.plan.inverse=d61f3210204da88f0538ebdaa5eb6ae7bd8f9ca03c42834c6e9056fcbd555317
media.caller.ordinary.plan.source=07024199f93cdd1b6e087e7cd886e7763ec6a75ae8f7ea450531f3f165344937
media.caller.ordinary.plan.destination=24e41b192c86fca83f215298c2436748c67be1858550466268ef9814ee400db9
media.caller.ordinary.plan.target=d4204f706cdf171ebf3531237e3eef184953e21e7a00c639ea73a7446591582a
media.caller.ordinary.plan.source_physical=88c6b1a49bad6800952511209d613349b5359a2322b8d483036503432e6c8704
media.caller.ordinary.plan.destination_physical=78f26cda5979a1a843d8cb06c0c24370ff0039c521f7bcbe246efcfbebb10fdb
media.caller.ordinary.plan.target_physical=1b241666c48dff1f8189560f8b324718dcfbc53667854a08ffe444c0a441550a
media.caller.ordinary.plan.transfers=true
media.caller.ordinary.apply.refused=Refused(Refusal { kind: UnsupportedOperation, operation: "slide_copy", reason: "the selected slide cannot be copied losslessly because the destination has caller-defined parts" })
media.caller.ordinary.apply.untouched=c45268167a26c6d6e2b1ad6ca988c31770f684cf4835f4c1fa00a2cf94f4f2b1
media.caller.ordinary.patch.refused=Refused(Refusal { kind: UnsupportedOperation, operation: "slide_copy", reason: "the selected slide cannot be copied losslessly because the destination has caller-defined parts" })
media.caller.ordinary.patch.untouched=c45268167a26c6d6e2b1ad6ca988c31770f684cf4835f4c1fa00a2cf94f4f2b1
plain.caller.bytes=52d5f13b15dbb56f935f11424ffdb9d6cd9b7a593b859c56632c8978435d5211
plain.caller.plan.bytes=268a5dde8b586393ad80e2bd4de4440779480093a0bfdba6501ca048e55ac006
plain.caller.plan.inverse=b5829818870bf71017d7efa674620058bef3e6d39f0fa954f736e93be77e0896
plain.caller.plan.source=820f57f2cc9b2e45f3d06aaa1341994c24afee6824d94f77fb3079fabb0fa3c1
plain.caller.plan.destination=45c181cc4dff24ab938d6da20545468b171075d8d5482bdf2155242732ddcaa8
plain.caller.plan.target=4465b6e4adf6f5210468c8395f4fb335ed3c2fb842f68d89f40dfa586a39ddd8
plain.caller.plan.source_physical=4abfb1a5b7b2648f1343f16bbef7fcd494a0100e95743d8ce34bcf289d07c9e6
plain.caller.plan.destination_physical=69a56b89b7716219b1c5b212128198e9ba2a5948cf033c973a2121ada24102b9
plain.caller.plan.target_physical=a981d5ef9ecf851f799c2ffb33c32976f79bf1df99bf57473cbeab707b91e773
plain.caller.plan.transfers=false
plain.caller.apply.published=1ac79056de3ed2a09006b1d0c49aa5aac4bebb9fe6f87fe2f0f153c6cf818eae
plain.caller.apply.revision=4465b6e4adf6f5210468c8395f4fb335ed3c2fb842f68d89f40dfa586a39ddd8
plain.caller.patch.forward.published=1ac79056de3ed2a09006b1d0c49aa5aac4bebb9fe6f87fe2f0f153c6cf818eae
plain.caller.patch.forward.revision=4465b6e4adf6f5210468c8395f4fb335ed3c2fb842f68d89f40dfa586a39ddd8
plain.caller.patch.undo.published=52d5f13b15dbb56f935f11424ffdb9d6cd9b7a593b859c56632c8978435d5211
plain.caller.patch.undo.revision=45c181cc4dff24ab938d6da20545468b171075d8d5482bdf2155242732ddcaa8
plain.caller.ordinary.plan.bytes=268a5dde8b586393ad80e2bd4de4440779480093a0bfdba6501ca048e55ac006
plain.caller.ordinary.plan.inverse=b5829818870bf71017d7efa674620058bef3e6d39f0fa954f736e93be77e0896
plain.caller.ordinary.plan.source=820f57f2cc9b2e45f3d06aaa1341994c24afee6824d94f77fb3079fabb0fa3c1
plain.caller.ordinary.plan.destination=45c181cc4dff24ab938d6da20545468b171075d8d5482bdf2155242732ddcaa8
plain.caller.ordinary.plan.target=4465b6e4adf6f5210468c8395f4fb335ed3c2fb842f68d89f40dfa586a39ddd8
plain.caller.ordinary.plan.source_physical=4abfb1a5b7b2648f1343f16bbef7fcd494a0100e95743d8ce34bcf289d07c9e6
plain.caller.ordinary.plan.destination_physical=69a56b89b7716219b1c5b212128198e9ba2a5948cf033c973a2121ada24102b9
plain.caller.ordinary.plan.target_physical=a981d5ef9ecf851f799c2ffb33c32976f79bf1df99bf57473cbeab707b91e773
plain.caller.ordinary.plan.transfers=false
plain.caller.ordinary.apply.published=1ac79056de3ed2a09006b1d0c49aa5aac4bebb9fe6f87fe2f0f153c6cf818eae
plain.caller.ordinary.apply.revision=4465b6e4adf6f5210468c8395f4fb335ed3c2fb842f68d89f40dfa586a39ddd8
plain.caller.ordinary.patch.published=1ac79056de3ed2a09006b1d0c49aa5aac4bebb9fe6f87fe2f0f153c6cf818eae
plain.caller.ordinary.patch.revision=4465b6e4adf6f5210468c8395f4fb335ed3c2fb842f68d89f40dfa586a39ddd8
media.limits.patch.bound=71858
media.limits.patch.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 71858 }
media.limits.patch.patch.untouched=d1dc2a4b199a0c74955a07fc4514c26066fc97b1205b630ebd727d390f858912
media.limits.patch.plan.refused=Limit { resource: "cross-slide candidate patch bytes", limit: 71858 }
media.limits.source-under.bound=71940
media.limits.source-under.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 71940 }
media.limits.source-under.patch.untouched=d1dc2a4b199a0c74955a07fc4514c26066fc97b1205b630ebd727d390f858912
media.limits.source-under.plan.refused=Limit { resource: "cross-slide candidate patch bytes", limit: 71940 }
media.limits.source-exact.bound=71941
media.limits.source-exact.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 71941 }
media.limits.source-exact.patch.untouched=d1dc2a4b199a0c74955a07fc4514c26066fc97b1205b630ebd727d390f858912
media.limits.source-exact.plan.refused=Limit { resource: "cross-slide candidate patch bytes", limit: 71941 }
media.limits.destination-under.bound=72046
media.limits.destination-under.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 72046 }
media.limits.destination-under.patch.untouched=d1dc2a4b199a0c74955a07fc4514c26066fc97b1205b630ebd727d390f858912
media.limits.destination-under.plan.refused=Limit { resource: "cross-slide candidate patch bytes", limit: 72046 }
media.limits.destination-exact.bound=72047
media.limits.destination-exact.patch.refused=PatchConflict
media.limits.destination-exact.patch.untouched=d1dc2a4b199a0c74955a07fc4514c26066fc97b1205b630ebd727d390f858912
media.limits.destination-exact.plan.refused=Limit { resource: "cross-slide candidate patch bytes", limit: 72047 }
plain.limits.patch.bound=4896
plain.limits.patch.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 4896 }
plain.limits.patch.patch.untouched=e0beb28b70e0f02a41a449f5341d199e36417a5a2e85664f6f46e57eb80bc409
plain.limits.patch.plan.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 4896 }
plain.limits.source-under.bound=31374
plain.limits.source-under.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 31374 }
plain.limits.source-under.patch.untouched=e0beb28b70e0f02a41a449f5341d199e36417a5a2e85664f6f46e57eb80bc409
plain.limits.source-under.plan.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 31374 }
plain.limits.source-exact.bound=31375
plain.limits.source-exact.patch.refused=PatchConflict
plain.limits.source-exact.patch.untouched=e0beb28b70e0f02a41a449f5341d199e36417a5a2e85664f6f46e57eb80bc409
plain.limits.source-exact.plan.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 31375 }
plain.limits.destination-under.bound=30458
plain.limits.destination-under.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 30458 }
plain.limits.destination-under.patch.untouched=e0beb28b70e0f02a41a449f5341d199e36417a5a2e85664f6f46e57eb80bc409
plain.limits.destination-under.plan.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 30458 }
plain.limits.destination-exact.bound=30459
plain.limits.destination-exact.patch.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 30459 }
plain.limits.destination-exact.patch.untouched=e0beb28b70e0f02a41a449f5341d199e36417a5a2e85664f6f46e57eb80bc409
plain.limits.destination-exact.plan.refused=Limit { resource: "cross-slide serialized archive bytes", limit: 30459 }
"#;

/// LPTX0003 changes wire hashes and size-dependent refusal points. The original
/// transcript remains frozen; all its non-wire, non-limit outcomes still agree.
/// The versioned transcript checks every current wire and refusal outcome too.
#[test]
fn versioned_patch_wire_preserves_base_revisions_and_publication() -> Result<()> {
    let transcript = full_transcript()?;
    let actual: Vec<&str> = transcript.lines().collect();
    let expected: Vec<&str> = include_str!("digest_reuse_lptx0003.golden.txt")
        .lines()
        .collect();
    for legacy in GOLDEN.lines() {
        let (key, _) = legacy.split_once('=').expect("golden row");
        if !key.contains(".limits.") && !key.ends_with(".bytes") && !key.ends_with(".inverse") {
            assert!(
                actual.contains(&legacy),
                "non-wire legacy outcome changed: {key}"
            );
        }
    }
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

/// Print the transcript, for regenerating `GOLDEN` against the earlier encoder:
/// `cargo test -p litchi-pptx --lib digest_reuse_tests::print -- --ignored --nocapture`.
#[test]
#[ignore = "prints the golden transcript; run explicitly"]
fn print_the_golden_transcript() -> Result<()> {
    println!("{}", full_transcript()?);
    Ok(())
}
