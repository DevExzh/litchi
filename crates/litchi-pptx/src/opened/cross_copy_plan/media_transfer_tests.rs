//! Proof obligations of the owned cross-copy's source-compressed media
//! transfer (change 0742).
//!
//! A copy planned against an unmodified owned destination publishes each
//! eligible copied image from the source member's verified compressed bytes
//! in fresh known-size framing, and records that encoding in the plan and in
//! the `LPCP0004` durable patch. These tests pin: the published span and its
//! framing; byte identity of every route that publishes the same copy; the
//! members that keep the recompressing route; the modified-destination
//! refusal and its remedy; source freshness, foreignness and provenance; the
//! refusal of genuine `LPCP0003` patches by name; and the new typed refusal of
//! a source member whose strict layout cannot be proven.
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "media-transfer tests use panic-on-fixture-failure assertions"
)]

use sha2::{Digest, Sha256};
use soapberry_zip::ZipArchive;
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

use super::{CrossSlideCopyPatch, CrossSlideCopyPlan};
use crate::media_parts::Resource;
use crate::{DurablePatchFormat, Error, Package, Result};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcError, PackURI};

const PHOTO: &str = "/ppt/media/transfer-photo.png";
const FLAT: &str = "/ppt/media/transfer-flat.png";
const VECTOR: &str = "/ppt/media/transfer-vector.svg";
const SOURCE_SLIDE: usize = 2;
const DESTINATION_SLIDE: usize = 1;
const POSITION: usize = 1;

fn png(mut body: Vec<u8>) -> Vec<u8> {
    body[..8].copy_from_slice(b"\x89PNG\r\n\x1a\n");
    body
}

/// Incompressible bytes: a Deflate stream over them is a run of stored
/// blocks, the shape of real PNG and JPEG payloads.
fn noisy(len: usize, mut state: u32) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(len);
    while bytes.len() < len {
        state ^= state << 13;
        state ^= state >> 17;
        state ^= state << 5;
        bytes.push((state >> 24) as u8);
    }
    png(bytes)
}

fn flat(len: usize) -> Vec<u8> {
    png((0..len).map(|index| (index % 7) as u8).collect())
}

const SVG: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?><svg xmlns="http://www.w3.org/2000/svg" width="1" height="1"/>"#;

fn authored(prefix: &str, slides: usize) -> Result<Vec<u8>> {
    let mut package = Package::new()?;
    let presentation = package.presentation_mut()?;
    for index in 0..slides {
        presentation
            .add_slide()?
            .set_title(&format!("change-0742-{prefix}-{index}"));
    }
    package.to_bytes()
}

/// Add `pictures` to one slide through the opened transaction, as a caller
/// would, and serialize the result.
fn with_pictures(bytes: &[u8], slide: usize, pictures: &[(&str, &str, &[u8])]) -> Result<Vec<u8>> {
    let mut package = Package::from_bytes(bytes)?;
    let mut edit = package.opened_presentation_transaction()?;
    for (index, (part_name, content_type, data)) in pictures.iter().enumerate() {
        edit.add_picture(
            slide,
            format!("change-0742-picture-{index}"),
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

fn photo_bytes() -> Vec<u8> {
    noisy(48 * 1024, 0x0742_1111)
}

fn flat_bytes() -> Vec<u8> {
    flat(32 * 1024)
}

/// A source whose slide carries a stored incompressible photo and a
/// deflated flat image, and a destination whose first slide already uses the
/// same media names, so the copy has to rename them.
fn media_fixture() -> Result<(Vec<u8>, Vec<u8>)> {
    let photo = photo_bytes();
    let flat = flat_bytes();
    let pictures: [(&str, &str, &[u8]); 2] =
        [(PHOTO, "image/png", &photo), (FLAT, "image/png", &flat)];
    let source = repack(
        &with_pictures(&authored("source", 3)?, SOURCE_SLIDE, &pictures)?,
        |name| name == &PHOTO[1..],
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

/// One archive member's raw framing and compressed bytes.
#[derive(Debug)]
struct RawMember {
    method: u16,
    flags: u16,
    local_time: [u16; 2],
    central_time: [u16; 2],
    local_crc: u32,
    central_crc: u32,
    local_compressed_size: u32,
    local_extra: usize,
    central_extra: usize,
    compressed: Vec<u8>,
}

fn u16_at(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes([bytes[offset], bytes[offset + 1]])
}

fn u32_at(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().expect("u32 field"))
}

fn raw_member(archive: &[u8], name: &str) -> RawMember {
    let zip = ZipArchive::from_slice(archive).expect("parse ZIP");
    for entry in zip.entries() {
        let entry = entry.expect("central record");
        if entry.file_path().as_ref() != name.as_bytes() {
            continue;
        }
        let local = usize::try_from(entry.local_header_offset()).expect("offset");
        let central = usize::try_from(entry.central_directory_offset()).expect("offset");
        let data = zip.get_entry(entry.wayfinder()).expect("local entry");
        return RawMember {
            method: u16_at(archive, local + 8),
            flags: u16_at(archive, local + 6),
            local_time: [u16_at(archive, local + 10), u16_at(archive, local + 12)],
            central_time: [u16_at(archive, central + 12), u16_at(archive, central + 14)],
            local_crc: u32_at(archive, local + 14),
            central_crc: u32_at(archive, central + 16),
            local_compressed_size: u32_at(archive, local + 18),
            local_extra: usize::from(u16_at(archive, local + 28)),
            central_extra: usize::from(u16_at(archive, central + 30)),
            compressed: data.data().to_vec(),
        };
    }
    panic!("archive has no member {name}");
}

fn assert_fresh_sized_framing(member: &RawMember) {
    assert_eq!(member.flags & 0x08, 0, "no data descriptor is declared");
    assert_eq!(member.local_time, [0, 0], "no local timestamp");
    assert_eq!(member.central_time, [0, 0], "no central timestamp");
    assert_eq!(member.local_extra, 0, "no local extras");
    assert_eq!(member.central_extra, 0, "no central extras");
    assert_eq!(
        member.local_crc, member.central_crc,
        "the CRC is in the local header"
    );
    assert_eq!(
        usize::try_from(member.local_compressed_size).expect("size"),
        member.compressed.len(),
        "the compressed size is in the local header"
    );
}

fn image_parts(plan: &CrossSlideCopyPlan) -> Vec<&crate::opened::SlideCopyPart> {
    plan.parts()
        .iter()
        .filter(|part| part.content_type().starts_with("image/"))
        .collect()
}

fn assert_unsafe_edit(error: &Error) {
    assert!(matches!(error, Error::UnsafeEdit { .. }), "{error:?}");
}

#[test]
fn copied_images_carry_their_source_compressed_span_in_fresh_framing() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let source = Package::from_vec(source_bytes.clone())?;
    let mut destination = Package::from_vec(destination_bytes)?;
    let plan = plan_copy(&source, &destination)?;
    assert!(plan.transfers_source_compressed_media());
    assert!(plan.patch().transfers_source_compressed_media());
    let images = image_parts(&plan);
    assert_eq!(images.len(), 2);

    destination.apply_cross_slide_copy_plan(&source, &plan)?;
    let output = destination.to_bytes()?;
    let reopened = Package::from_vec(output.clone())?;
    for part in images {
        assert_ne!(
            part.source(),
            part.target(),
            "the copy renames colliding media"
        );
        let original = raw_member(&source_bytes, part.source().membername());
        let published = raw_member(&output, part.target().membername());
        assert_eq!(
            published.compressed,
            original.compressed,
            "{}: the published payload is the source member's exact compressed span",
            part.target()
        );
        assert_eq!(
            published.method, original.method,
            "the source method is kept"
        );
        assert_eq!(published.central_crc, original.central_crc);
        assert_fresh_sized_framing(&published);
        let planned = source.opc.get_part(part.source())?.blob().to_vec();
        assert_eq!(planned.len(), part.bytes());
        assert_eq!(
            reopened.opc.get_part(part.target())?.blob(),
            planned.as_slice(),
            "the transferred span decodes to the planned bytes"
        );
    }
    let photo = raw_member(&output, images_target(&plan, PHOTO).membername());
    assert_eq!(photo.method, 0, "a stored source member stays stored");
    let flat = raw_member(&output, images_target(&plan, FLAT).membername());
    assert_eq!(flat.method, 8, "a deflated source member stays deflated");

    // XML members still take the recompressing route.
    let slide = plan
        .parts()
        .iter()
        .find(|part| part.content_type() == ct::PML_SLIDE)
        .expect("the copied slide");
    let slide_member = raw_member(&output, slide.target().membername());
    assert_eq!(slide_member.method, 8);
    assert_ne!(
        slide_member.flags & 0x08,
        0,
        "generated Deflate uses a descriptor"
    );
    Ok(())
}

fn images_target<'a>(plan: &'a CrossSlideCopyPlan, source: &str) -> &'a PackURI {
    plan.parts()
        .iter()
        .find(|part| part.source().as_str() == source)
        .map(|part| part.target())
        .expect("the planned image")
}

#[test]
fn retained_released_and_durable_outputs_are_byte_identical() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let source = Package::from_vec(source_bytes.clone())?;
    let plan = plan_copy(&source, &Package::from_vec(destination_bytes.clone())?)?;
    assert!(plan.transfers_source_compressed_media());
    assert!(plan.retained_candidate_bytes().is_some());

    let mut retained = Package::from_vec(destination_bytes.clone())?;
    retained.apply_cross_slide_copy_plan(&source, &plan)?;
    let retained_output = retained.to_bytes()?;
    assert_eq!(
        plan.candidate
            .0
            .as_ref()
            .map(|held| held.archive.as_slice()),
        Some(retained_output.as_slice()),
        "the retained archive is the published archive"
    );

    let mut released_plan = plan.clone();
    released_plan.release_retained_candidate();
    let mut released = Package::from_vec(destination_bytes.clone())?;
    released.apply_cross_slide_copy_plan(&source, &released_plan)?;
    assert_eq!(released.to_bytes()?, retained_output);

    let durable = CrossSlideCopyPatch::from_bytes(&plan.patch().to_bytes()?)?;
    assert!(durable.transfers_source_compressed_media());
    assert_eq!(durable, *plan.patch());
    let mut patched = Package::from_vec(destination_bytes.clone())?;
    patched.apply_cross_slide_copy_patch(&source, &durable)?;
    assert_eq!(patched.to_bytes()?, retained_output);

    // Independent planning over fresh opens of the same bytes is equal.
    let again = plan_copy(
        &Package::from_vec(source_bytes.clone())?,
        &Package::from_vec(destination_bytes.clone())?,
    )?;
    assert_eq!(again, plan);
    let mut independent = Package::from_vec(destination_bytes.clone())?;
    independent.apply_cross_slide_copy_plan(&Package::from_vec(source_bytes)?, &again)?;
    assert_eq!(independent.to_bytes()?, retained_output);

    // The inverse restores the destination byte for byte.
    let inverse = CrossSlideCopyPatch::from_bytes(&durable.inverse().to_bytes()?)?;
    assert!(inverse.transfers_source_compressed_media());
    patched.apply_cross_slide_copy_patch(&source, &inverse)?;
    assert_eq!(patched.to_bytes()?, destination_bytes);
    Ok(())
}

#[test]
fn a_replaced_payload_keeps_the_recompressing_route_beside_a_transferred_one() -> Result<()> {
    let photo = photo_bytes();
    let flat = flat_bytes();
    let pictures: [(&str, &str, &[u8]); 2] =
        [(PHOTO, "image/png", &photo), (FLAT, "image/png", &flat)];
    let source_bytes = repack(
        &with_pictures(&authored("source", 3)?, SOURCE_SLIDE, &pictures)?,
        |name| name.starts_with("ppt/media/"),
    );
    let destination_bytes = authored("destination", 2)?;
    let mut source = Package::from_vec(source_bytes.clone())?;
    // Replace the photo with byte-identical bytes: the payload is now the
    // caller's, not the source member's, so it is not transferred. The flat
    // image is untouched and is.
    let photo_uri = PackURI::new(PHOTO).map_err(Error::Invalid)?;
    let same = source.opc.get_part(&photo_uri)?.blob().to_vec();
    source.opc.get_part_mut(&photo_uri)?.set_blob(same);
    let mut destination = Package::from_vec(destination_bytes)?;
    let plan = plan_copy(&source, &destination)?;
    assert!(plan.transfers_source_compressed_media());
    destination.apply_cross_slide_copy_plan(&source, &plan)?;
    let output = destination.to_bytes()?;
    let original_photo = raw_member(&source_bytes, &PHOTO[1..]);
    let published_photo = raw_member(&output, images_target(&plan, PHOTO).membername());
    assert_eq!(original_photo.method, 0, "the fixture stores its media");
    assert_eq!(
        published_photo.method, 8,
        "the replaced photo is recompressed"
    );
    assert_ne!(published_photo.flags & 0x08, 0, "generated Deflate framing");
    let original_flat = raw_member(&source_bytes, &FLAT[1..]);
    let published_flat = raw_member(&output, images_target(&plan, FLAT).membername());
    assert_eq!(published_flat.method, 0, "the untouched image stays stored");
    assert_eq!(published_flat.compressed, original_flat.compressed);
    assert_fresh_sized_framing(&published_flat);
    let reopened = Package::from_vec(output)?;
    assert_eq!(
        reopened.opc.get_part(images_target(&plan, PHOTO))?.blob(),
        photo.as_slice()
    );
    Ok(())
}

/// An SVG picture is XML: the owned cross-copy refuses it as an unmodeled
/// XML media surface before any encoding is chosen, exactly as before.
#[test]
fn an_svg_picture_is_still_refused_before_any_encoding_is_chosen() -> Result<()> {
    let pictures: [(&str, &str, &[u8]); 1] = [(VECTOR, "image/svg+xml", SVG)];
    let source = Package::from_vec(with_pictures(
        &authored("source", 3)?,
        SOURCE_SLIDE,
        &pictures,
    )?)?;
    let destination = Package::from_vec(authored("destination", 2)?)?;
    assert!(matches!(
        plan_copy(&source, &destination),
        Err(Error::SlideCopyPlan {
            kind: crate::SlideCopyRefusal::UnknownSemanticSurface,
            ..
        })
    ));
    Ok(())
}

#[test]
fn plain_copies_record_the_recompressed_encoding() -> Result<()> {
    let source = Package::from_vec(authored("source", 3)?)?;
    let destination = Package::from_vec(authored("destination", 2)?)?;
    let plan = plan_copy(&source, &destination)?;
    assert!(!plan.transfers_source_compressed_media());
    assert!(!plan.patch().inverse().transfers_source_compressed_media());
    Ok(())
}

fn presentation_uri() -> PackURI {
    PackURI::new("/ppt/presentation.xml").expect("presentation URI")
}

#[test]
fn a_modified_destination_refuses_a_transferring_copy_until_it_is_replanned() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let source = Package::from_vec(source_bytes)?;
    let plan = plan_copy(&source, &Package::from_vec(destination_bytes.clone())?)?;
    assert!(plan.transfers_source_compressed_media());
    let durable = CrossSlideCopyPatch::from_bytes(&plan.patch().to_bytes()?)?;

    // A mutable access without a change leaves both revisions intact but
    // revokes the destination's exact-source authorization.
    let mut modified = Package::from_vec(destination_bytes.clone())?;
    modified.opc.get_part_mut(&presentation_uri())?;
    assert!(!modified.opc.is_unmodified_owned_source());
    let before = modified.to_bytes()?;
    assert_eq!(before, destination_bytes, "the revocation changed no byte");

    let error = modified
        .apply_cross_slide_copy_plan(&source, &plan)
        .expect_err("a transferring plan needs an unmodified owned destination");
    assert_unsafe_edit(&error);
    assert!(error.to_string().contains("unmodified owned destination"));
    assert_eq!(modified.to_bytes()?, before);
    let error = modified
        .apply_cross_slide_copy_patch(&source, &durable)
        .expect_err("a transferring durable patch needs an unmodified owned destination");
    assert_unsafe_edit(&error);
    assert_eq!(modified.to_bytes()?, before);

    // Planning against the modified destination records the recompressed
    // encoding and applies.
    let replanned = plan_copy(&source, &modified)?;
    assert!(!replanned.transfers_source_compressed_media());
    assert_ne!(
        replanned.target_physical_revision(),
        plan.target_physical_revision()
    );
    assert_eq!(replanned.target_revision(), plan.target_revision());
    modified.apply_cross_slide_copy_plan(&source, &replanned)?;
    let output = modified.to_bytes()?;
    for part in image_parts(&replanned) {
        assert_eq!(raw_member(&output, part.target().membername()).method, 8);
    }
    Ok(())
}

#[test]
fn the_inverse_of_a_transferring_copy_applies_to_a_modified_destination() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let source = Package::from_vec(source_bytes)?;
    let mut destination = Package::from_vec(destination_bytes.clone())?;
    let plan = plan_copy(&source, &destination)?;
    destination.apply_cross_slide_copy_plan(&source, &plan)?;
    let copied = destination.to_bytes()?;

    let mut modified = Package::from_vec(copied.clone())?;
    modified.opc.get_part_mut(&presentation_uri())?;
    assert!(!modified.opc.is_unmodified_owned_source());
    let inverse = CrossSlideCopyPatch::from_bytes(&plan.patch().inverse().to_bytes()?)?;
    modified.apply_cross_slide_copy_patch(&source, &inverse)?;
    assert_eq!(modified.to_bytes()?, destination_bytes);
    Ok(())
}

#[test]
fn stale_foreign_and_reprovenanced_sources_are_refused() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let source = Package::from_vec(source_bytes.clone())?;
    let plan = plan_copy(&source, &Package::from_vec(destination_bytes.clone())?)?;
    assert!(plan.transfers_source_compressed_media());

    let mut expected = Package::from_vec(destination_bytes.clone())?;
    expected.apply_cross_slide_copy_plan(&source, &plan)?;
    let expected = expected.to_bytes()?;

    // A fresh open of the same bytes is the same source.
    let mut reopened_destination = Package::from_vec(destination_bytes.clone())?;
    reopened_destination
        .apply_cross_slide_copy_plan(&Package::from_vec(source_bytes.clone())?, &plan)?;
    assert_eq!(reopened_destination.to_bytes()?, expected);

    // A foreign source is refused before anything is published.
    let foreign = Package::from_vec(authored("foreign", 3)?)?;
    let mut refused = Package::from_vec(destination_bytes.clone())?;
    assert_unsafe_edit(
        &refused
            .apply_cross_slide_copy_plan(&foreign, &plan)
            .expect_err("foreign source"),
    );
    assert_eq!(refused.to_bytes()?, destination_bytes);

    // A stale source whose copied image changed is refused.
    let changed_photo = noisy(48 * 1024, 0x0742_2222);
    let stale_bytes = repack(
        &with_pictures(
            &authored("source", 3)?,
            SOURCE_SLIDE,
            &[
                (PHOTO, "image/png", &changed_photo),
                (FLAT, "image/png", &flat_bytes()),
            ],
        )?,
        |name| name == &PHOTO[1..],
    );
    let stale = Package::from_vec(stale_bytes)?;
    assert_unsafe_edit(
        &refused
            .apply_cross_slide_copy_plan(&stale, &plan)
            .expect_err("stale source"),
    );
    assert_eq!(refused.to_bytes()?, destination_bytes);

    // A source whose photo payload was replaced with equal bytes after
    // planning proves the same revisions but no longer lends its member, so
    // the fresh candidate differs from the planned one and is refused.
    let mut reprovenanced = Package::from_vec(source_bytes.clone())?;
    let photo_uri = PackURI::new(PHOTO).map_err(Error::Invalid)?;
    let same = reprovenanced.opc.get_part(&photo_uri)?.blob().to_vec();
    reprovenanced.opc.get_part_mut(&photo_uri)?.set_blob(same);
    let error = refused
        .apply_cross_slide_copy_plan(&reprovenanced, &plan)
        .expect_err("a re-provenanced source cannot lend its member");
    assert_unsafe_edit(&error);
    assert!(error.to_string().contains("freshly proven candidate"));
    assert_eq!(refused.to_bytes()?, destination_bytes);
    // Retention moves no verdict: the released plan is refused identically,
    // because a retained archive is reused only for the transfer set it was
    // built with.
    let mut released = plan.clone();
    released.release_retained_candidate();
    let released_error = refused
        .apply_cross_slide_copy_plan(&reprovenanced, &released)
        .expect_err("the released plan is refused too");
    assert_eq!(released_error.to_string(), error.to_string());
    assert_eq!(refused.to_bytes()?, destination_bytes);

    // The reverse flip: a plan made while the photo was re-provenanced does
    // not transfer it, and a source that could lend it again is refused by
    // both routes.
    let partial = plan_copy(
        &reprovenanced,
        &Package::from_vec(destination_bytes.clone())?,
    )?;
    assert!(
        partial.transfers_source_compressed_media(),
        "the flat image still transfers"
    );
    assert_ne!(
        partial.target_physical_revision(),
        plan.target_physical_revision()
    );
    let mut released_partial = partial.clone();
    released_partial.release_retained_candidate();
    for candidate in [&partial, &released_partial] {
        let error = refused
            .apply_cross_slide_copy_plan(&source, candidate)
            .expect_err("a source that lends the photo again is a different candidate");
        assert_unsafe_edit(&error);
        assert_eq!(refused.to_bytes()?, destination_bytes);
    }
    // And the source it was planned against still applies it.
    refused.apply_cross_slide_copy_plan(&reprovenanced, &partial)?;
    Ok(())
}

#[test]
fn a_source_member_whose_local_header_disagrees_refuses_the_transferring_copy() -> Result<()> {
    let (mut source_bytes, destination_bytes) = media_fixture()?;
    // Rename the photo's local header only; the central record, which the
    // ordinary reader trusts, still names it.
    let zip = ZipArchive::from_slice(&source_bytes).expect("parse ZIP");
    let mut local = None;
    for entry in zip.entries() {
        let entry = entry.expect("central record");
        if entry.file_path().as_ref() == &PHOTO.as_bytes()[1..] {
            local = Some(usize::try_from(entry.local_header_offset()).expect("offset"));
        }
    }
    let position = local.expect("photo member") + 30 + PHOTO.len() - 1 - 5;
    assert_eq!(source_bytes[position], b'o');
    source_bytes[position] = b'0';

    let source = Package::from_vec(source_bytes)?;
    let error = plan_copy(&source, &Package::from_vec(destination_bytes.clone())?)
        .expect_err("a transfer needs a provable strict local layout");
    assert!(
        matches!(&error, Error::Opc(OpcError::ZipError(message)) if message.contains("names differ")),
        "{error:?}"
    );

    // The recompressing route decodes through the central record and still
    // copies the slide.
    let mut modified = Package::from_vec(destination_bytes)?;
    modified.opc.get_part_mut(&presentation_uri())?;
    let plan = plan_copy(&source, &modified)?;
    assert!(!plan.transfers_source_compressed_media());
    modified.apply_cross_slide_copy_plan(&source, &plan)?;
    Ok(())
}

fn legacy_image_payload() -> Vec<u8> {
    let mut bytes = Vec::with_capacity(64);
    let mut state = 0x0742_u32;
    bytes.extend_from_slice(b"\x89PNG\r\n\x1a\n");
    while bytes.len() < 64 {
        state = state.wrapping_mul(1_103_515_245).wrapping_add(12_345);
        bytes.push((state >> 16) as u8);
    }
    bytes
}

fn hex(bytes: &[u8]) -> String {
    Sha256::digest(bytes)
        .iter()
        .map(|byte| format!("{byte:02x}"))
        .collect()
}

/// Genuine `LPCP0003` patches, written by the unchanged base tree (009d515bef)
/// for a copy whose closure carries one PNG, are refused by name in both
/// directions. The test also rebuilds the inputs the base tree planned over,
/// proves they are the same bytes, and shows why the refusal is needed: the
/// semantic and input physical revisions are unchanged, but the copy's
/// target physical revision is not, so no current application could
/// reproduce the legacy proof.
#[test]
fn legacy_lpcp0003_patches_are_refused_by_name() -> Result<()> {
    const FORWARD: &[u8] = include_bytes!(
        "../../../../../test-data/ooxml/pptx/cross-copy-legacy/lpcp0003-forward.patch"
    );
    const INVERSE: &[u8] = include_bytes!(
        "../../../../../test-data/ooxml/pptx/cross-copy-legacy/lpcp0003-inverse.patch"
    );
    for (label, bytes) in [("forward", FORWARD), ("inverse", INVERSE)] {
        assert_eq!(&bytes[..8], b"LPCP0003", "{label} fixture magic");
        for parsed in [
            CrossSlideCopyPatch::from_bytes(bytes),
            CrossSlideCopyPatch::from_bytes_with_limits(bytes, super::Limits::default()),
        ] {
            let error = parsed.expect_err("a legacy patch is refused");
            assert!(
                matches!(
                    error,
                    Error::DurablePatchRevisionFormat {
                        found: DurablePatchFormat::CrossSlideCopyV3,
                        expected: DurablePatchFormat::CrossSlideCopyV4,
                    }
                ),
                "{label}: {error:?}"
            );
            assert!(error.to_string().contains("LPCP0003"));
        }
    }

    // The generator's inputs, rebuilt with the same authoring steps.
    let source_bytes = {
        let mut package = Package::from_bytes(&authored("source", 3)?)?;
        let mut edit = package.opened_presentation_transaction()?;
        edit.add_picture(
            SOURCE_SLIDE,
            "change-0742-legacy-picture",
            &Resource::new(
                "/ppt/media/change-0742-legacy.png",
                "image/png",
                legacy_image_payload(),
            ),
            (800, 800, 72, 72),
        )?;
        let commit = edit.commit()?;
        package.apply_opened_presentation_commit(commit)?;
        package.to_bytes()?
    };
    let destination_bytes = authored("destination", 2)?;
    assert_eq!(
        hex(&source_bytes),
        "204bfc1b3a334c2709327d077add72ecbd3a42eb0f319bdcff6dbe17cbaaeb8f"
    );
    assert_eq!(
        hex(&destination_bytes),
        "458d1e2787537506787aabb38b405ddb08708dd107e0d0462d91c1cf0455f12d"
    );
    let plan = plan_copy(
        &Package::from_vec(source_bytes)?,
        &Package::from_vec(destination_bytes)?,
    )?;
    assert!(plan.transfers_source_compressed_media());
    let revision = |offset: usize| -> [u8; 32] {
        FORWARD[8 + 32 * offset..8 + 32 * (offset + 1)]
            .try_into()
            .expect("revision")
    };
    assert_eq!(revision(0), plan.source_revision());
    assert_eq!(revision(1), plan.destination_revision());
    assert_eq!(revision(2), plan.target_revision());
    assert_eq!(revision(3), plan.source_physical_revision());
    assert_eq!(revision(4), plan.destination_physical_revision());
    assert_ne!(
        revision(5),
        plan.target_physical_revision(),
        "the copied image's encoding changed the serialized target"
    );
    Ok(())
}
