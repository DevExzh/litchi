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
use litchi_opc::PackURI;
use litchi_opc::constants::content_type as ct;

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

/// A source edited since it was opened lends exactly the members its own
/// serialization carries: a payload replaced with equal bytes serializes as
/// the original member, and a payload replaced with other bytes serializes as
/// a regenerated one. Either is transferred from those bytes, and the plan
/// equals the plan made over a fresh open of the same serialization.
#[test]
fn a_modified_source_lends_the_members_it_publishes() -> Result<()> {
    let photo = photo_bytes();
    let flat = flat_bytes();
    let pictures: [(&str, &str, &[u8]); 2] =
        [(PHOTO, "image/png", &photo), (FLAT, "image/png", &flat)];
    let source_bytes = repack(
        &with_pictures(&authored("source", 3)?, SOURCE_SLIDE, &pictures)?,
        |name| name.starts_with("ppt/media/"),
    );
    let destination_bytes = authored("destination", 2)?;
    let photo_uri = PackURI::new(PHOTO).map_err(Error::Invalid)?;
    let changed = noisy(48 * 1024, 0x0742_3333);
    for replacement in [photo.clone(), changed] {
        let mut source = Package::from_vec(source_bytes.clone())?;
        source
            .opc
            .get_part_mut(&photo_uri)?
            .set_blob(replacement.clone());
        assert!(!source.opc.is_unmodified_owned_source());
        let published_source = source.to_bytes()?;
        let mut destination = Package::from_vec(destination_bytes.clone())?;
        let plan = plan_copy(&source, &destination)?;
        assert!(plan.transfers_source_compressed_media());
        destination.apply_cross_slide_copy_plan(&source, &plan)?;
        let output = destination.to_bytes()?;
        for (name, bytes) in [(PHOTO, &replacement), (FLAT, &flat)] {
            let published = raw_member(&output, images_target(&plan, name).membername());
            assert_eq!(
                published.compressed,
                raw_member(&published_source, &name[1..]).compressed,
                "{name}: the member the source publishes is transferred"
            );
            assert_fresh_sized_framing(&published);
            let reopened = Package::from_vec(output.clone())?;
            assert_eq!(
                reopened.opc.get_part(images_target(&plan, name))?.blob(),
                bytes.as_slice()
            );
        }
        // The same plan and output from a fresh open of the serialization.
        let clean = Package::from_vec(published_source)?;
        let clean_plan = plan_copy(&clean, &Package::from_vec(destination_bytes.clone())?)?;
        assert_eq!(clean_plan, plan);
        let mut clean_destination = Package::from_vec(destination_bytes.clone())?;
        clean_destination.apply_cross_slide_copy_plan(&clean, &plan)?;
        assert_eq!(clean_destination.to_bytes()?, output);
    }
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

/// A destination edited since it was opened, but byte-identical, accepts a
/// transferring plan and durable patch and publishes exactly what an
/// unmodified destination publishes; planning against it gives the same plan.
#[test]
fn a_byte_identical_modified_destination_accepts_the_same_copy() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let source = Package::from_vec(source_bytes)?;
    let plan = plan_copy(&source, &Package::from_vec(destination_bytes.clone())?)?;
    assert!(plan.transfers_source_compressed_media());
    let durable = CrossSlideCopyPatch::from_bytes(&plan.patch().to_bytes()?)?;
    let mut expected = Package::from_vec(destination_bytes.clone())?;
    expected.apply_cross_slide_copy_plan(&source, &plan)?;
    let expected = expected.to_bytes()?;

    // A mutable access without a change leaves both revisions intact but
    // revokes the destination's exact-source authorization.
    let modified = || -> Result<Package> {
        let mut modified = Package::from_vec(destination_bytes.clone())?;
        modified.opc.get_part_mut(&presentation_uri())?;
        assert!(!modified.opc.is_unmodified_owned_source());
        Ok(modified)
    };
    let mut by_plan = modified()?;
    assert_eq!(by_plan.to_bytes()?, destination_bytes, "no byte changed");
    by_plan.apply_cross_slide_copy_plan(&source, &plan)?;
    assert_eq!(by_plan.to_bytes()?, expected);
    let mut by_patch = modified()?;
    by_patch.apply_cross_slide_copy_patch(&source, &durable)?;
    assert_eq!(by_patch.to_bytes()?, expected);

    // Planning against the modified destination is planning against its
    // bytes.
    let replanned = plan_copy(&source, &modified()?)?;
    assert_eq!(replanned, plan);
    let mut again = modified()?;
    again.apply_cross_slide_copy_plan(&source, &replanned)?;
    assert_eq!(again.to_bytes()?, expected);
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
fn stale_and_foreign_sources_are_refused_and_a_reprovenanced_one_is_not() -> Result<()> {
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
    // planning proves the same revisions and publishes the same bytes, so it
    // lends the same members: the plan applies, retained or released, and so
    // does a plan made against it.
    let mut reprovenanced = Package::from_vec(source_bytes)?;
    let photo_uri = PackURI::new(PHOTO).map_err(Error::Invalid)?;
    let same = reprovenanced.opc.get_part(&photo_uri)?.blob().to_vec();
    reprovenanced.opc.get_part_mut(&photo_uri)?.set_blob(same);
    let mut released = plan.clone();
    released.release_retained_candidate();
    for candidate in [&plan, &released] {
        let mut destination = Package::from_vec(destination_bytes.clone())?;
        destination.apply_cross_slide_copy_plan(&reprovenanced, candidate)?;
        assert_eq!(destination.to_bytes()?, expected);
    }
    let replanned = plan_copy(
        &reprovenanced,
        &Package::from_vec(destination_bytes.clone())?,
    )?;
    assert_eq!(replanned, plan);
    let mut destination = Package::from_vec(destination_bytes)?;
    destination.apply_cross_slide_copy_plan(&source, &replanned)?;
    assert_eq!(destination.to_bytes()?, expected);
    Ok(())
}

#[test]
fn a_source_member_whose_local_header_disagrees_keeps_the_recompressing_route() -> Result<()> {
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

    // The photo's strict layout is disproven by its headers alone, so it is
    // not eligible and keeps the recompressing route, decoded through the
    // central record as before; the untouched flat image still transfers.
    let source = Package::from_vec(source_bytes.clone())?;
    let mut destination = Package::from_vec(destination_bytes.clone())?;
    let plan = plan_copy(&source, &destination)?;
    assert!(plan.transfers_source_compressed_media());
    destination.apply_cross_slide_copy_plan(&source, &plan)?;
    let output = destination.to_bytes()?;
    let photo = raw_member(&output, images_target(&plan, PHOTO).membername());
    assert_eq!(photo.method, 8, "the photo is recompressed");
    assert_ne!(photo.flags & 0x08, 0, "generated Deflate framing");
    let flat = raw_member(&output, images_target(&plan, FLAT).membername());
    assert_eq!(
        flat.compressed,
        raw_member(&source_bytes, &FLAT[1..]).compressed,
        "the flat image is transferred"
    );
    assert_fresh_sized_framing(&flat);
    let reopened = Package::from_vec(output)?;
    assert_eq!(
        reopened.opc.get_part(images_target(&plan, PHOTO))?.blob(),
        photo_bytes().as_slice()
    );

    // The verdict does not depend on the destination: a modified
    // destination gives the same plan and the same bytes.
    let mut modified = Package::from_vec(destination_bytes)?;
    modified.opc.get_part_mut(&presentation_uri())?;
    let replanned = plan_copy(&source, &modified)?;
    assert_eq!(replanned, plan);
    modified.apply_cross_slide_copy_plan(&source, &replanned)?;
    assert_eq!(
        modified.to_bytes()?,
        Package::from_vec(output_bytes_of(&source, &plan)?)?.to_bytes()?
    );
    Ok(())
}

/// The bytes an unmodified destination of the media fixture publishes for
/// `plan`.
fn output_bytes_of(source: &Package, plan: &CrossSlideCopyPlan) -> Result<Vec<u8>> {
    let (_source_bytes, destination_bytes) = media_fixture()?;
    let mut destination = Package::from_vec(destination_bytes)?;
    destination.apply_cross_slide_copy_plan(source, plan)?;
    destination.to_bytes()
}

/// The format half of the eligibility rule, including the XML guard the
/// cross-copy's own SVG refusal otherwise leaves unexercised.
#[test]
fn only_relationship_free_non_xml_images_are_transferable_media() -> Result<()> {
    let part = |name: &str, content_type: &str| -> Result<litchi_opc::BlobPart> {
        Ok(litchi_opc::BlobPart::new(
            PackURI::new(name).map_err(Error::Invalid)?,
            content_type.to_owned(),
            b"payload".to_vec(),
        ))
    };
    assert!(super::is_transferable_media(&part(
        "/ppt/media/a.png",
        "image/png"
    )?));
    assert!(super::is_transferable_media(&part(
        "/ppt/media/a.jpeg",
        "IMAGE/JPEG"
    )?));
    assert!(!super::is_transferable_media(&part(
        "/ppt/media/a.svg",
        "image/svg+xml"
    )?));
    assert!(!super::is_transferable_media(&part(
        "/ppt/media/a.xml",
        "image/png"
    )?));
    assert!(!super::is_transferable_media(&part(
        "/ppt/media/a.bin",
        "application/octet-stream"
    )?));
    assert!(!super::is_transferable_media(&part(
        "/ppt/media/a.wmf",
        "imagex/wmf"
    )?));
    let mut linked = part("/ppt/media/b.png", "image/png")?;
    litchi_opc::Part::rels_mut(&mut linked).try_add_relationship(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink".to_owned(),
        "https://example.invalid/".to_owned(),
        "rId1".to_owned(),
        litchi_opc::TargetMode::External,
    )?;
    assert!(!super::is_transferable_media(&linked));
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
/// proves they are the same bytes, and shows what the format change is about:
/// the semantic and input physical revisions are unchanged, but the same copy
/// planned today transfers the PNG and carries a different target physical
/// revision, so a legacy header — which does not record its encoding — cannot
/// be compared with a fresh plan.
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

/// Insert `padding` zero bytes after the compressed payload of one member of a
/// ZIP32 archive written without data descriptors, fixing the member's local
/// and central compressed sizes and every later offset. A Deflate decoder
/// stops at the final block, so the ordinary reader still decodes the member.
fn pad_member_payload(archive: &[u8], name: &str, padding: usize) -> Vec<u8> {
    splice_member_payload(archive, name, false, &vec![0; padding])
}

/// Insert `inserted` at the start or the end of a sized member's compressed
/// payload, growing its declared compressed size and moving every later
/// offset.
fn splice_member_payload(archive: &[u8], name: &str, at_start: bool, inserted: &[u8]) -> Vec<u8> {
    let padding = inserted.len();
    let zip = ZipArchive::from_slice(archive).expect("parse ZIP");
    let directory = usize::try_from(zip.directory_offset()).expect("offset");
    let eocd = usize::try_from(zip.eocd_offset()).expect("offset");
    let mut target = None;
    for entry in zip.entries() {
        let entry = entry.expect("central record");
        if entry.file_path().as_ref() == name.as_bytes() {
            assert!(
                !entry.has_data_descriptor(),
                "the helper pads sized members"
            );
            let (start, end) = zip
                .get_entry(entry.wayfinder())
                .expect("local entry")
                .compressed_data_range();
            target = Some((
                usize::try_from(entry.local_header_offset()).expect("offset"),
                usize::try_from(entry.central_directory_offset()).expect("offset"),
                usize::try_from(if at_start { start } else { end }).expect("offset"),
            ));
        }
    }
    // Every offset at or after the insertion point moves; the member's own
    // local header precedes it.
    let (local, central, payload_end) = target.expect("the member exists");
    let grow = |bytes: &mut [u8], offset: usize| {
        let value = u32_at(bytes, offset) + u32::try_from(padding).expect("padding");
        bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    };
    let mut output = archive[..payload_end].to_vec();
    output.extend_from_slice(inserted);
    output.extend_from_slice(&archive[payload_end..]);
    grow(&mut output, local + 18);
    let shift = |offset: usize| {
        if offset >= payload_end {
            offset + padding
        } else {
            offset
        }
    };
    let central = shift(central);
    grow(&mut output, central + 20);
    // Every central record whose local header follows the payload moves.
    let mut record = shift(directory);
    let end = shift(eocd);
    while record < end {
        let local_offset = usize::try_from(u32_at(&output, record + 42)).expect("offset");
        if local_offset >= payload_end {
            grow(&mut output, record + 42);
        }
        record += 46
            + usize::from(u16_at(&output, record + 28))
            + usize::from(u16_at(&output, record + 30))
            + usize::from(u16_at(&output, record + 32));
    }
    grow(&mut output, end + 16);
    output
}

/// The reviewer's counterexample: 16 bytes after the final Deflate block of a
/// sized member. The ordinary reader decodes it; the capture's consumed-input
/// check refuses it.
#[test]
fn a_member_with_bytes_after_its_final_block_keeps_the_recompressing_route() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let padded = pad_member_payload(&source_bytes, &FLAT[1..], 16);
    let source = Package::from_vec(padded.clone())?;
    let flat_uri = PackURI::new(FLAT).map_err(Error::Invalid)?;
    assert_eq!(
        source.opc.get_part(&flat_uri)?.blob(),
        flat_bytes().as_slice()
    );
    let mut destination = Package::from_vec(destination_bytes.clone())?;
    let plan = plan_copy(&source, &destination)?;
    assert!(
        plan.transfers_source_compressed_media(),
        "the photo still transfers"
    );
    destination.apply_cross_slide_copy_plan(&source, &plan)?;
    let output = destination.to_bytes()?;
    let flat = raw_member(&output, images_target(&plan, FLAT).membername());
    assert_eq!(flat.method, 8);
    assert_ne!(flat.flags & 0x08, 0, "the padded member is recompressed");
    let photo = raw_member(&output, images_target(&plan, PHOTO).membername());
    assert_eq!(
        photo.compressed,
        raw_member(&padded, &PHOTO[1..]).compressed
    );
    // The durable patch reproduces the same decision.
    let mut patched = Package::from_vec(destination_bytes)?;
    patched.apply_cross_slide_copy_patch(
        &source,
        &CrossSlideCopyPatch::from_bytes(&plan.patch().to_bytes()?)?,
    )?;
    assert_eq!(patched.to_bytes()?, output);
    Ok(())
}

/// The reviewer's other counterexample: undo then redo of a transferring copy.
#[test]
fn redo_after_undo_publishes_the_first_copy_again() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let source = Package::from_vec(source_bytes)?;
    let plan = plan_copy(&source, &Package::from_vec(destination_bytes.clone())?)?;
    assert!(plan.transfers_source_compressed_media());
    let forward = CrossSlideCopyPatch::from_bytes(&plan.patch().to_bytes()?)?;
    let inverse = CrossSlideCopyPatch::from_bytes(&forward.inverse().to_bytes()?)?;

    let mut first = Package::from_vec(destination_bytes.clone())?;
    first.apply_cross_slide_copy_patch(&source, &forward)?;
    let first = first.to_bytes()?;
    let undone = || -> Result<Package> {
        let mut destination = Package::from_vec(destination_bytes.clone())?;
        destination.apply_cross_slide_copy_patch(&source, &forward)?;
        destination.apply_cross_slide_copy_patch(&source, &inverse)?;
        assert_eq!(destination.to_bytes()?, destination_bytes, "undo is exact");
        Ok(destination)
    };

    let mut by_patch = undone()?;
    by_patch.apply_cross_slide_copy_patch(&source, &forward)?;
    assert_eq!(by_patch.to_bytes()?, first, "redo by the durable patch");
    let mut by_plan = undone()?;
    by_plan.apply_cross_slide_copy_plan(&source, &plan)?;
    assert_eq!(by_plan.to_bytes()?, first, "redo by the plan");
    Ok(())
}

/// Offset of the copied-media encoding byte in an `LPCP0004` header.
fn encoding_offset(plan: &CrossSlideCopyPlan) -> usize {
    8 + 6 * 32
        + (4 + plan.source().part_name().as_str().len())
        + (4 + plan.destination().part_name().as_str().len())
        + (4 + plan.destination_layout().as_str().len())
        + 8
        + 4
        + (4 + plan.presentation_relationship_id().len())
}

/// An unknown encoding byte is a typed `Invalid`; flipping a transferring
/// patch to the recompressed encoding parses, but no candidate reproduces its
/// physical revisions, so application refuses and publishes nothing.
#[test]
fn a_tampered_encoding_byte_is_refused() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let source = Package::from_vec(source_bytes)?;
    let plan = plan_copy(&source, &Package::from_vec(destination_bytes.clone())?)?;
    let encoded = plan.patch().to_bytes()?;
    let offset = encoding_offset(&plan);
    assert_eq!(
        encoded[offset], 1,
        "a transferring patch records encoding 1"
    );

    let mut unknown = encoded.clone();
    unknown[offset] = 2;
    assert!(matches!(
        CrossSlideCopyPatch::from_bytes(&unknown),
        Err(Error::Invalid(message)) if message.contains("copied-media encoding")
    ));

    let mut flipped = encoded;
    flipped[offset] = 0;
    let flipped = CrossSlideCopyPatch::from_bytes(&flipped)?;
    assert!(!flipped.transfers_source_compressed_media());
    for patch in [flipped.clone(), flipped.inverse()] {
        let mut destination = Package::from_vec(destination_bytes.clone())?;
        assert_unsafe_edit(
            &destination
                .apply_cross_slide_copy_patch(&source, &patch)
                .expect_err("a flipped encoding is not a proven candidate"),
        );
        assert_eq!(destination.to_bytes()?, destination_bytes);
    }
    Ok(())
}

/// The captures a transferring copy holds are charged, before any capture is
/// taken, against the destination's `max_patch_bytes`: one byte below the
/// charge is refused by name although the copy fits without them.
#[test]
fn captures_are_charged_against_the_patch_byte_limit() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let source = Package::from_vec(source_bytes)?;
    let destination = Package::from_vec(destination_bytes)?;
    let plan = plan_copy(&source, &destination)?;
    let mut capture_bytes = 0usize;
    for part in image_parts(&plan) {
        let size = source
            .opc
            .compressed_transfer_size(part.source())?
            .expect("the fixture's images are eligible");
        capture_bytes += usize::try_from(size).expect("size");
    }
    let snapshot = destination.opened_presentation()?;
    let without = super::candidate_estimate(&snapshot, plan.parts(), plan.planned_bytes(), 0)?;
    let with =
        super::candidate_estimate(&snapshot, plan.parts(), plan.planned_bytes(), capture_bytes)?;
    assert_eq!(with, without + capture_bytes);
    let limits = |max_patch_bytes| {
        let default = super::Limits::default();
        super::Limits::new(
            default.max_parts(),
            max_patch_bytes,
            default.max_text_bytes(),
            default.max_history_entries(),
            default.max_history_bytes(),
            default.max_retained_candidate_bytes(),
        )
        .expect("finite limits")
    };
    let plan_under = |max_patch_bytes| {
        let limits = limits(max_patch_bytes);
        destination
            .opened_presentation_with_limits(limits)?
            .plan_cross_slide_copy(
                &source.opened_presentation_with_limits(limits)?,
                SOURCE_SLIDE,
                DESTINATION_SLIDE,
                POSITION,
            )
    };
    assert!(matches!(
        plan_under(with - 1),
        Err(Error::Limit {
            resource: "cross-slide candidate patch bytes",
            limit,
        }) if limit == with - 1
    ));
    assert!(plan_under(with)?.transfers_source_compressed_media());
    Ok(())
}

/// A candidate published by a transferring copy is an owned package whose
/// parts were materialized eagerly; it lends the same members onward, so a
/// second copy frames the original source's compressed bytes again, and its
/// durable patch round-trips.
#[test]
fn a_published_candidate_lends_its_media_to_a_further_copy() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let source = Package::from_vec(source_bytes.clone())?;
    let mut middle = Package::from_vec(destination_bytes)?;
    let first = plan_copy(&source, &middle)?;
    middle.apply_cross_slide_copy_plan(&source, &first)?;
    assert!(middle.opc.is_unmodified_owned_source());

    let third_bytes = authored("third", 2)?;
    let mut third = Package::from_vec(third_bytes.clone())?;
    let second = third.opened_presentation()?.plan_cross_slide_copy(
        &middle.opened_presentation()?,
        POSITION,
        0_usize,
        1,
    )?;
    assert!(second.transfers_source_compressed_media());
    third.apply_cross_slide_copy_plan(&middle, &second)?;
    let output = third.to_bytes()?;
    for name in [PHOTO, FLAT] {
        let in_middle = images_target(&first, name);
        let in_third = second
            .parts()
            .iter()
            .find(|part| part.source() == in_middle)
            .map(|part| part.target())
            .expect("the image is copied again");
        let published = raw_member(&output, in_third.membername());
        assert_eq!(
            published.compressed,
            raw_member(&source_bytes, &name[1..]).compressed,
            "{name}: the original source member's bytes"
        );
        assert_fresh_sized_framing(&published);
    }
    let durable = CrossSlideCopyPatch::from_bytes(&second.patch().to_bytes()?)?;
    let mut patched = Package::from_vec(third_bytes.clone())?;
    patched.apply_cross_slide_copy_patch(&middle, &durable)?;
    assert_eq!(patched.to_bytes()?, output);
    patched.apply_cross_slide_copy_patch(
        &middle,
        &CrossSlideCopyPatch::from_bytes(&durable.inverse().to_bytes()?)?,
    )?;
    assert_eq!(patched.to_bytes()?, third_bytes);
    Ok(())
}

/// A caller-defined part whose relationship-counting behaviour differs from
/// a built-in part's.
#[derive(Clone)]
struct Sentinel(litchi_opc::BlobPart);

impl litchi_opc::Part for Sentinel {
    fn blob(&self) -> &[u8] {
        self.0.blob()
    }
    fn blob_arc(&self) -> std::sync::Arc<Vec<u8>> {
        self.0.blob_arc()
    }
    fn content_type(&self) -> &str {
        self.0.content_type()
    }
    fn partname(&self) -> &PackURI {
        self.0.partname()
    }
    fn rel_ref_count(&self, _r_id: &str) -> usize {
        0x51_7e
    }
    fn rels(&self) -> &litchi_opc::Relationships {
        self.0.rels()
    }
    fn rels_mut(&mut self) -> &mut litchi_opc::Relationships {
        self.0.rels_mut()
    }
    fn set_blob(&mut self, blob: Vec<u8>) {
        self.0.set_blob(blob);
    }
}

/// A transferring copy into a destination that holds a caller-defined part
/// and save preferences publishes the reopened candidate: the part keeps its
/// bytes but becomes a built-in part, and the save preferences are carried.
/// The output is the output of the same copy into the destination's bytes.
#[test]
fn a_transferring_copy_keeps_save_options_but_not_caller_part_types() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let source = Package::from_vec(source_bytes)?;
    let sentinel = PackURI::new("/custom/sentinel.bin").map_err(Error::Invalid)?;
    let custom = || -> Result<Package> {
        let mut destination = Package::from_vec(destination_bytes.clone())?;
        destination
            .opc
            .try_add_part(Box::new(Sentinel(litchi_opc::BlobPart::new(
                sentinel.clone(),
                "application/octet-stream".to_owned(),
                b"sentinel".to_vec(),
            ))))?;
        destination.opc.set_save_options(litchi_opc::SaveOptions {
            fonts: litchi_opc::FontEmbedding::Full,
        });
        Ok(destination)
    };
    let mut destination = custom()?;
    assert_eq!(
        destination.opc.get_part(&sentinel)?.rel_ref_count("rId1"),
        0x51_7e
    );
    let plan = plan_copy(&source, &destination)?;
    assert!(plan.transfers_source_compressed_media());
    destination.apply_cross_slide_copy_plan(&source, &plan)?;
    assert_eq!(
        destination.opc.save_options().fonts,
        litchi_opc::FontEmbedding::Full
    );
    let part = destination.opc.get_part(&sentinel)?;
    assert_eq!(part.blob(), b"sentinel");
    assert_eq!(part.rel_ref_count("rId1"), 0, "a built-in part now");
    let published = litchi_opc::PackageWriter::to_bytes(&destination.opc)?;

    let clean_bytes = litchi_opc::PackageWriter::to_bytes(&custom()?.opc)?;
    let mut clean = Package::from_vec(clean_bytes)?;
    let clean_plan = plan_copy(&source, &clean)?;
    assert_eq!(clean_plan, plan);
    clean.apply_cross_slide_copy_plan(&source, &clean_plan)?;
    assert_eq!(litchi_opc::PackageWriter::to_bytes(&clean.opc)?, published);
    Ok(())
}

/// A producer that Stores a compressible image gets it transferred Stored:
/// the source's choice is kept, as the source-backed route keeps it, at the
/// cost of a larger member than recompression would give.
#[test]
fn a_stored_compressible_image_is_transferred_stored() -> Result<()> {
    let flat = flat_bytes();
    let source_bytes = repack(
        &with_pictures(
            &authored("source", 3)?,
            SOURCE_SLIDE,
            &[(FLAT, "image/png", &flat)],
        )?,
        |name| name.starts_with("ppt/media/"),
    );
    let source = Package::from_vec(source_bytes)?;
    let mut destination = Package::from_vec(authored("destination", 2)?)?;
    let plan = plan_copy(&source, &destination)?;
    assert!(plan.transfers_source_compressed_media());
    destination.apply_cross_slide_copy_plan(&source, &plan)?;
    let output = destination.to_bytes()?;
    let published = raw_member(&output, images_target(&plan, FLAT).membername());
    assert_eq!(published.method, 0);
    assert_eq!(published.compressed, flat);
    Ok(())
}

/// The reviewer's padding case: 200,000 empty stored blocks in front of the
/// flat image's Deflate stream. The member decodes to the same image, but its
/// declared compressed size fails the header-only guard, so it is recompressed
/// instead of being published padded; the photo still transfers.
#[test]
fn a_member_padded_with_empty_stored_blocks_is_recompressed() -> Result<()> {
    let (source_bytes, destination_bytes) = media_fixture()?;
    let empty_blocks = [0x00, 0x00, 0x00, 0xff, 0xff].repeat(200_000);
    let padded = splice_member_payload(&source_bytes, &FLAT[1..], true, &empty_blocks);
    let source = Package::from_vec(padded.clone())?;
    let flat_uri = PackURI::new(FLAT).map_err(Error::Invalid)?;
    assert_eq!(
        source.opc.get_part(&flat_uri)?.blob(),
        flat_bytes().as_slice()
    );
    assert_eq!(source.opc.compressed_transfer_size(&flat_uri)?, None);
    let mut destination = Package::from_vec(destination_bytes)?;
    let plan = plan_copy(&source, &destination)?;
    assert!(
        plan.transfers_source_compressed_media(),
        "the photo transfers"
    );
    destination.apply_cross_slide_copy_plan(&source, &plan)?;
    let output = destination.to_bytes()?;
    let flat = raw_member(&output, images_target(&plan, FLAT).membername());
    assert_eq!(flat.method, 8);
    assert_ne!(flat.flags & 0x08, 0, "the padded member is recompressed");
    assert!(flat.compressed.len() < 1024);
    assert!(output.len() < padded.len() / 2);
    Ok(())
}
