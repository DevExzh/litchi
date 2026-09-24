//! Proof obligations of the verified compressed transfer (change 0742).
//!
//! These tests pin, at the OPC boundary: which parts are eligible and that
//! eligibility never decodes; that every refusal after eligibility is typed;
//! that the published member carries exactly the source member's compressed
//! span inside fresh known-size framing; and that any mutation of a
//! transferred part discards its capture.
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "compressed-transfer tests use panic-on-fixture-failure assertions"
)]

use std::io::Write as _;

use soapberry_zip::extra_fields::ExtraFieldId;
use soapberry_zip::office::StreamingArchiveWriter;
use soapberry_zip::time::UtcDateTime;
use soapberry_zip::{CompressionMethod, Header, ZipArchive, ZipArchiveWriter};

use crate::limits::{ReadLimits, ReadResource};
use crate::packuri::PackURI;
use crate::part::{BlobPart, Part};
use crate::pkgwriter::PackageWriter;
use crate::{OpcError, OpcPackage};

const PHOTO: &str = "doc/media/photo.png";
const STORED: &str = "doc/media/stored.png";
const VECTOR: &str = "doc/media/vector.svg";
const LINKED: &str = "doc/media/linked.png";
const COPIED: &str = "/doc/media/copied.png";

#[derive(Clone, Copy)]
enum Mode {
    Stored,
    DeflatedDescriptor,
    DeflatedSized,
}

fn pseudo_random_bytes(len: usize, mut state: u32) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(len);
    while bytes.len() < len {
        state ^= state << 13;
        state ^= state >> 17;
        state ^= state << 5;
        bytes.push((state >> 24) as u8);
    }
    bytes
}

/// A compressible payload, so its Deflate stream is not a run of stored
/// blocks and a corrupted byte is certain to break decoding.
fn compressible_bytes(len: usize) -> Vec<u8> {
    (0..len)
        .map(|index| b"litchi compressed transfer "[index % 27])
        .collect()
}

fn part_uri(member: &str) -> PackURI {
    PackURI::new(format!("/{member}")).expect("test part URI")
}

fn content_types(signed: bool) -> Vec<u8> {
    let signature = if signed {
        r#"<Default Extension="sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/>"#
    } else {
        ""
    };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="svg" ContentType="image/svg+xml"/>{signature}<Override PartName="/doc/main.xml" ContentType="application/vnd.litchi.test.main+xml"/></Types>"#
    )
    .into_bytes()
}

fn package_relationships(signed: bool) -> Vec<u8> {
    let signature = if signed {
        r#"<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/>"#
    } else {
        ""
    };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="doc/main.xml"/>{signature}</Relationships>"#
    )
    .into_bytes()
}

const MAIN_RELATIONSHIPS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/photo.png"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/stored.png"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/vector.svg"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/linked.png"/></Relationships>"#;
const LINKED_RELATIONSHIPS: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/" TargetMode="External"/></Relationships>"#;
const MAIN_XML: &[u8] =
    br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><main xmlns="urn:litchi:test"/>"#;
const VECTOR_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><svg xmlns="http://www.w3.org/2000/svg"/>"#;

fn photo_bytes() -> Vec<u8> {
    compressible_bytes(64 * 1024)
}

fn stored_bytes() -> Vec<u8> {
    pseudo_random_bytes(16 * 1024, 0x0742_0742)
}

/// The fixed source package with its photo written in `photo_mode`.
fn source_archive(photo_mode: Mode, signed: bool) -> Vec<u8> {
    let photo = photo_bytes();
    let stored = stored_bytes();
    let content_types = content_types(signed);
    let package_relationships = package_relationships(signed);
    let mut members: Vec<(&str, &[u8], Mode)> = vec![
        ("[Content_Types].xml", &content_types, Mode::DeflatedSized),
        ("_rels/.rels", &package_relationships, Mode::DeflatedSized),
        ("doc/main.xml", MAIN_XML, Mode::DeflatedSized),
        (
            "doc/_rels/main.xml.rels",
            MAIN_RELATIONSHIPS,
            Mode::DeflatedSized,
        ),
        (PHOTO, &photo, photo_mode),
        (STORED, &stored, Mode::Stored),
        (VECTOR, VECTOR_XML, Mode::DeflatedSized),
        (LINKED, &stored, Mode::DeflatedSized),
        (
            "doc/media/_rels/linked.png.rels",
            LINKED_RELATIONSHIPS,
            Mode::DeflatedSized,
        ),
    ];
    if signed {
        members.push(("_xmlsignatures/origin.sigs", b"", Mode::Stored));
    }
    streaming_archive(&members)
}

fn streaming_archive(members: &[(&str, &[u8], Mode)]) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    for (name, bytes, mode) in members {
        match mode {
            Mode::Stored => writer.write_stored(name, bytes),
            Mode::DeflatedDescriptor => writer.write_deflated(name, bytes),
            Mode::DeflatedSized => writer.write_deflated_sized(name, bytes),
        }
        .expect("test member");
    }
    writer.finish_to_bytes().expect("test archive")
}

/// The same package with its stored photo written as a streaming member that
/// carries a data descriptor, a timestamp and a source extra field.
fn source_archive_with_source_framing() -> Vec<u8> {
    let photo = stored_bytes();
    let content_types = content_types(false);
    let package_relationships = package_relationships(false);
    let mut archive = ZipArchiveWriter::new(Vec::new());
    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_slice()),
        ("_rels/.rels", package_relationships.as_slice()),
        ("doc/main.xml", MAIN_XML),
        ("doc/_rels/main.xml.rels", MAIN_RELATIONSHIPS),
    ] {
        archive.write_stored_file(name, bytes).expect("test member");
    }
    let (mut entry, config) = archive
        .new_file(PHOTO)
        .compression_method(CompressionMethod::Store)
        .last_modified(UtcDateTime::from_unix(1_700_000_000))
        .extra_field(
            ExtraFieldId::new(0x6c74),
            b"source-extra",
            Header::default(),
        )
        .expect("test extra field")
        .start()
        .expect("test streaming member");
    let mut data = config.wrap(&mut entry);
    data.write_all(&photo).expect("test payload");
    let (_, descriptor) = data.finish().expect("test payload finish");
    entry.finish(descriptor).expect("test member finish");
    for (name, bytes) in [
        (STORED, stored_bytes()),
        (VECTOR, VECTOR_XML.to_vec()),
        (LINKED, stored_bytes()),
        (
            "doc/media/_rels/linked.png.rels",
            LINKED_RELATIONSHIPS.to_vec(),
        ),
    ] {
        archive
            .write_stored_file(name, &bytes)
            .expect("test member");
    }
    archive.finish().expect("test archive")
}

fn destination_archive() -> Vec<u8> {
    let content_types = content_types(false);
    let package_relationships = package_relationships(false);
    streaming_archive(&[
        ("[Content_Types].xml", &content_types, Mode::DeflatedSized),
        ("_rels/.rels", &package_relationships, Mode::DeflatedSized),
        ("doc/main.xml", MAIN_XML, Mode::DeflatedSized),
    ])
}

/// Publish `part` as a new member of a fresh owned destination.
fn publish_into_destination(part: BlobPart) -> Vec<u8> {
    let mut destination = OpcPackage::from_vec(destination_archive()).expect("destination opens");
    destination
        .try_add_part(Box::new(part))
        .expect("copied part is added");
    PackageWriter::to_bytes(&destination).expect("destination publishes")
}

/// One archive member's raw framing and compressed bytes.
#[derive(Debug, PartialEq, Eq)]
struct RawMember {
    method: u16,
    flags: u16,
    local_time: [u16; 2],
    central_time: [u16; 2],
    local_crc: u32,
    local_sizes: [u32; 2],
    central_crc: u32,
    local_extra: Vec<u8>,
    central_extra: Vec<u8>,
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
        let local_name = usize::from(u16_at(archive, local + 26));
        let local_extra_len = usize::from(u16_at(archive, local + 28));
        let local_extra_start = local + 30 + local_name;
        let central_name = usize::from(u16_at(archive, central + 28));
        let central_extra_len = usize::from(u16_at(archive, central + 30));
        let central_extra_start = central + 46 + central_name;
        let data = zip.get_entry(entry.wayfinder()).expect("local entry");
        return RawMember {
            method: u16_at(archive, local + 8),
            flags: u16_at(archive, local + 6),
            local_time: [u16_at(archive, local + 10), u16_at(archive, local + 12)],
            central_time: [u16_at(archive, central + 12), u16_at(archive, central + 14)],
            local_crc: u32_at(archive, local + 14),
            local_sizes: [u32_at(archive, local + 18), u32_at(archive, local + 22)],
            central_crc: u32_at(archive, central + 16),
            local_extra: archive[local_extra_start..local_extra_start + local_extra_len].to_vec(),
            central_extra: archive[central_extra_start..central_extra_start + central_extra_len]
                .to_vec(),
            compressed: data.data().to_vec(),
        };
    }
    panic!("archive has no member {name}");
}

fn assert_fresh_sized_framing(member: &RawMember, expected_method: u16) {
    assert_eq!(member.method, expected_method, "the source method is kept");
    assert_eq!(member.flags & 0x08, 0, "no data descriptor is declared");
    assert_eq!(member.local_time, [0, 0], "no source local timestamp");
    assert_eq!(member.central_time, [0, 0], "no source central timestamp");
    assert!(member.local_extra.is_empty(), "no source local extras");
    assert!(member.central_extra.is_empty(), "no source central extras");
    assert_eq!(member.local_crc, member.central_crc, "sized local CRC");
    assert_eq!(
        usize::try_from(member.local_sizes[0]).expect("size"),
        member.compressed.len(),
        "sized local compressed size"
    );
}

fn read_member(archive: &[u8], partname: &str) -> Vec<u8> {
    let package = OpcPackage::from_bytes(archive).expect("published package reopens");
    package
        .get_part(&PackURI::new(partname).expect("URI"))
        .expect("published part")
        .blob()
        .to_vec()
}

fn is_ineligible(error: &OpcError) -> bool {
    matches!(error, OpcError::PreservationUnavailable { .. })
}

#[test]
fn a_deflated_member_transfers_its_exact_compressed_span_in_fresh_sized_framing() {
    for mode in [Mode::DeflatedDescriptor, Mode::DeflatedSized] {
        let source_bytes = source_archive(mode, false);
        let source = OpcPackage::from_vec(source_bytes.clone()).expect("source opens");
        let photo = part_uri(PHOTO);
        assert!(
            source
                .compressed_transfer_size(&photo)
                .expect("part exists")
                .is_some()
        );
        let transfer = source
            .authorize_compressed_transfer(&photo)
            .expect("an untouched binary leaf transfers")
            .expect("its capture verifies");
        let source_member = raw_member(&source_bytes, PHOTO);
        assert_eq!(transfer.content_type(), "image/png");
        assert_eq!(transfer.decoded_size(), photo_bytes().len());
        assert_eq!(
            transfer.compressed_size(),
            u64::try_from(source_member.compressed.len()).expect("size")
        );

        let copied = PackURI::new(COPIED).expect("URI");
        let output =
            publish_into_destination(BlobPart::with_compressed_transfer(copied.clone(), transfer));
        let published = raw_member(&output, copied.membername());
        assert_eq!(
            published.compressed, source_member.compressed,
            "the published payload is the source member's exact compressed span"
        );
        assert_eq!(published.central_crc, source_member.central_crc);
        assert_fresh_sized_framing(&published, 8);
        assert_eq!(read_member(&output, COPIED), photo_bytes());
        // Deterministic: the same inputs publish the same bytes.
        let again = OpcPackage::from_vec(source_bytes).expect("source reopens");
        let transfer = again
            .authorize_compressed_transfer(&photo)
            .expect("transfer is repeatable")
            .expect("its capture verifies");
        assert_eq!(
            publish_into_destination(BlobPart::with_compressed_transfer(copied, transfer)),
            output
        );
    }
}

#[test]
fn a_stored_member_stays_stored_and_drops_source_descriptor_timestamp_and_extras() {
    let source_bytes = source_archive_with_source_framing();
    let source_member = raw_member(&source_bytes, PHOTO);
    assert_ne!(
        source_member.flags & 0x08,
        0,
        "the fixture uses a descriptor"
    );
    assert_ne!(
        source_member.local_time,
        [0, 0],
        "the fixture is timestamped"
    );
    assert!(
        !source_member.central_extra.is_empty(),
        "the fixture has extras"
    );

    let source = OpcPackage::from_vec(source_bytes).expect("source opens");
    let transfer = source
        .authorize_compressed_transfer(&part_uri(PHOTO))
        .expect("a stored binary leaf transfers")
        .expect("its capture verifies");
    let copied = PackURI::new(COPIED).expect("URI");
    let output = publish_into_destination(BlobPart::with_compressed_transfer(copied, transfer));
    let published = raw_member(&output, &COPIED[1..]);
    assert_eq!(published.compressed, source_member.compressed);
    assert_eq!(published.compressed, stored_bytes());
    assert_fresh_sized_framing(&published, 0);
}

#[test]
fn an_eager_owned_package_issues_the_same_capture_as_a_deferred_one() {
    let source_bytes = source_archive(Mode::DeflatedDescriptor, false);
    let deferred = OpcPackage::from_vec(source_bytes.clone()).expect("deferred open");
    let donor = OpcPackage::new();
    let eager = OpcPackage::from_vec_reusing_payloads(source_bytes, ReadLimits::default(), &donor)
        .expect("eager open");
    let photo = part_uri(PHOTO);
    assert!(
        eager
            .compressed_transfer_size(&photo)
            .expect("part")
            .is_some()
    );
    let copied = PackURI::new(COPIED).expect("URI");
    let from_deferred = publish_into_destination(BlobPart::with_compressed_transfer(
        copied.clone(),
        deferred
            .authorize_compressed_transfer(&photo)
            .expect("deferred transfer")
            .expect("its capture verifies"),
    ));
    let from_eager = publish_into_destination(BlobPart::with_compressed_transfer(
        copied,
        eager
            .authorize_compressed_transfer(&photo)
            .expect("eager transfer")
            .expect("its capture verifies"),
    ));
    assert_eq!(from_deferred, from_eager);
}

#[test]
fn ineligible_parts_are_classified_without_decoding_and_refused_by_type() {
    let source_bytes = source_archive(Mode::DeflatedSized, false);
    let source = OpcPackage::from_vec(source_bytes.clone()).expect("source opens");
    let before = source.deferred_decode_counters();
    for (member, eligible) in [
        (PHOTO, true),
        (STORED, true),
        (VECTOR, false),
        (LINKED, false),
        ("doc/main.xml", false),
    ] {
        assert_eq!(
            source
                .compressed_transfer_size(&part_uri(member))
                .expect("part exists")
                .is_some(),
            eligible,
            "{member}"
        );
    }
    assert_eq!(
        source.deferred_decode_counters(),
        before,
        "eligibility never decodes a payload"
    );
    for member in [VECTOR, LINKED, "doc/main.xml"] {
        let error = source
            .authorize_compressed_transfer(&part_uri(member))
            .expect_err("an ineligible part is refused");
        assert!(is_ineligible(&error), "{member}: {error}");
    }
    assert!(matches!(
        source.compressed_transfer_size(&PackURI::new("/doc/media/absent.png").expect("URI")),
        Err(OpcError::PartNotFound(_))
    ));

    // A payload replaced with byte-identical content is a caller's payload,
    // not the source member's.
    let mut replaced = OpcPackage::from_vec(source_bytes.clone()).expect("source opens");
    let photo = part_uri(PHOTO);
    let same = replaced.get_part(&photo).expect("photo").blob().to_vec();
    replaced.get_part_mut(&photo).expect("photo").set_blob(same);
    assert!(
        replaced
            .compressed_transfer_size(&photo)
            .expect("part")
            .is_none()
    );
    assert!(is_ineligible(
        &replaced
            .authorize_compressed_transfer(&photo)
            .expect_err("replaced payload is refused")
    ));

    // A retyped part no longer has the content type its member was admitted with.
    let mut retyped = OpcPackage::from_vec(source_bytes.clone()).expect("source opens");
    retyped
        .get_part_mut(&photo)
        .expect("photo")
        .set_content_type("image/x-litchi".to_owned())
        .expect("blob parts can be retyped");
    assert!(
        retyped
            .compressed_transfer_size(&photo)
            .expect("part")
            .is_none()
    );

    // Borrowed ingress retains no archive to capture from.
    let borrowed = OpcPackage::from_bytes(&source_bytes).expect("borrowed open");
    assert!(
        borrowed
            .compressed_transfer_size(&photo)
            .expect("part")
            .is_none()
    );

    // A part authored in memory has no source member.
    let mut authored = OpcPackage::from_vec(source_bytes).expect("source opens");
    let fresh = PackURI::new("/doc/media/fresh.png").expect("URI");
    authored
        .try_add_part(Box::new(BlobPart::new(
            fresh.clone(),
            "image/png".to_owned(),
            stored_bytes(),
        )))
        .expect("fresh part");
    assert!(
        authored
            .compressed_transfer_size(&fresh)
            .expect("part")
            .is_none()
    );
    assert!(
        authored
            .compressed_transfer_size(&photo)
            .expect("part")
            .is_some()
    );
}

#[test]
fn signed_packages_are_ineligible_and_refused_by_policy() {
    let source = OpcPackage::from_vec(source_archive(Mode::DeflatedSized, true))
        .expect("signed source opens");
    assert!(source.is_signed());
    let photo = part_uri(PHOTO);
    assert!(
        source
            .compressed_transfer_size(&photo)
            .expect("part")
            .is_none()
    );
    assert!(matches!(
        source.authorize_compressed_transfer(&photo),
        Err(OpcError::SignedSourceRequiresExplicitPolicy)
    ));
}

/// Offset of one member's compressed bytes in an archive.
fn compressed_range(archive: &[u8], name: &str) -> (usize, usize) {
    let zip = ZipArchive::from_slice(archive).expect("parse ZIP");
    for entry in zip.entries() {
        let entry = entry.expect("central record");
        if entry.file_path().as_ref() == name.as_bytes() {
            let (start, end) = zip
                .get_entry(entry.wayfinder())
                .expect("local entry")
                .compressed_data_range();
            return (
                usize::try_from(start).expect("offset"),
                usize::try_from(end).expect("offset"),
            );
        }
    }
    panic!("archive has no member {name}");
}

fn central_offset(archive: &[u8], name: &str) -> (usize, usize) {
    let zip = ZipArchive::from_slice(archive).expect("parse ZIP");
    for entry in zip.entries() {
        let entry = entry.expect("central record");
        if entry.file_path().as_ref() == name.as_bytes() {
            return (
                usize::try_from(entry.local_header_offset()).expect("offset"),
                usize::try_from(entry.central_directory_offset()).expect("offset"),
            );
        }
    }
    panic!("archive has no member {name}");
}

/// Assert that the transfer is refused with a typed error and that the error
/// is the one the package's own accessor reports, when that accessor fails.
fn assert_refused(archive: Vec<u8>) {
    let source = OpcPackage::from_vec(archive).expect("the central directory is well formed");
    let photo = part_uri(PHOTO);
    let transfer = source.authorize_compressed_transfer(&photo);
    let error = transfer.expect_err("a malformed member never yields a transfer");
    assert!(
        matches!(
            error,
            OpcError::ZipError(_) | OpcError::IoError(_) | OpcError::ReadLimit { .. }
        ),
        "the refusal is typed: {error:?}"
    );
    if let Err(access) = source.get_part(&photo) {
        assert_eq!(
            access.to_string(),
            error.to_string(),
            "the refusal is stable"
        );
    }
    // Refusal is idempotent: a second attempt reports the same error.
    let again = source
        .authorize_compressed_transfer(&photo)
        .expect_err("still refused");
    assert_eq!(again.to_string(), error.to_string());
}

#[test]
fn a_corrupt_compressed_stream_is_refused_with_a_typed_error() {
    let mut archive = source_archive(Mode::DeflatedSized, false);
    let (start, end) = compressed_range(&archive, PHOTO);
    for byte in &mut archive[start + 2..end.min(start + 64)] {
        *byte ^= 0x5a;
    }
    assert_refused(archive);
}

#[test]
fn a_crc_mismatch_is_refused_with_a_typed_error() {
    let mut archive = source_archive(Mode::DeflatedSized, false);
    let (local, central) = central_offset(&archive, PHOTO);
    for offset in [local + 14, central + 16] {
        archive[offset] ^= 0xff;
    }
    assert_refused(archive);
}

#[test]
fn a_local_header_that_disagrees_with_its_central_record_is_not_eligible() {
    let mut archive = source_archive(Mode::DeflatedSized, false);
    let (local, _central) = central_offset(&archive, PHOTO);
    // Rename the local header only: `photo.png` becomes `phot0.png`.
    let name_start = local + 30;
    let position = name_start + PHOTO.len() - 5;
    assert_eq!(archive[position], b'o');
    archive[position] = b'0';
    // The ordinary reader treats the central record as authoritative and
    // still decodes the payload; the strict layout a compressed capture needs
    // is disproven by the headers alone, so the part is classified
    // ineligible, deterministically and without decoding, and a caller keeps
    // its recompressing route. Other members stay eligible.
    let source = OpcPackage::from_vec(archive).expect("the archive opens");
    let before = source.deferred_decode_counters();
    assert!(
        source
            .compressed_transfer_size(&part_uri(PHOTO))
            .expect("part")
            .is_none()
    );
    assert!(
        source
            .compressed_transfer_size(&part_uri(STORED))
            .expect("part")
            .is_some()
    );
    assert_eq!(source.deferred_decode_counters(), before);
    assert_eq!(
        source
            .get_part(&part_uri(PHOTO))
            .expect("lenient decode")
            .blob(),
        photo_bytes().as_slice()
    );
    assert!(is_ineligible(
        &source
            .authorize_compressed_transfer(&part_uri(PHOTO))
            .expect_err("an unprovable layout is never captured")
    ));
}

/// The capture's own verification: bytes that are not the member's decode
/// are refused with a typed error even when every header is consistent.
#[test]
fn a_capture_that_does_not_decode_to_the_payload_is_refused() {
    let source =
        OpcPackage::from_vec(source_archive(Mode::DeflatedSized, false)).expect("source opens");
    let photo = part_uri(PHOTO);
    let part = source.get_part(&photo).expect("photo");
    let source_part = source.transfer_provenance(part).expect("provenance");
    let member = source
        .transfer_member(part, source_part)
        .expect("index")
        .expect("member");
    let mut wrong = part.blob().to_vec();
    wrong[1024] ^= 0x01;
    assert!(matches!(super::verified_capture(&member, &wrong), Ok(None)));
    let mut short = part.blob().to_vec();
    short.pop();
    assert!(matches!(super::verified_capture(&member, &short), Ok(None)));
    assert!(matches!(
        super::verified_capture(&member, part.blob()),
        Ok(Some(_))
    ));
}

/// A custom part that forwards its payload handle to a transferred part but
/// shows different bytes publishes the bytes it shows.
#[test]
fn a_capture_is_framed_only_for_the_allocation_it_was_verified_against() {
    #[derive(Clone)]
    struct Forwarding {
        inner: BlobPart,
        shown: std::sync::Arc<Vec<u8>>,
    }
    impl Part for Forwarding {
        fn blob(&self) -> &[u8] {
            &self.shown
        }
        fn blob_arc(&self) -> std::sync::Arc<Vec<u8>> {
            std::sync::Arc::clone(&self.shown)
        }
        fn content_type(&self) -> &str {
            self.inner.content_type()
        }
        fn partname(&self) -> &PackURI {
            self.inner.partname()
        }
        fn payload_handle(&self) -> crate::part::PayloadHandle {
            self.inner.payload_handle()
        }
        fn rels(&self) -> &crate::Relationships {
            self.inner.rels()
        }
        fn rels_mut(&mut self) -> &mut crate::Relationships {
            self.inner.rels_mut()
        }
        fn set_blob(&mut self, blob: Vec<u8>) {
            self.shown = std::sync::Arc::new(blob);
        }
    }
    let source =
        OpcPackage::from_vec(source_archive(Mode::DeflatedSized, false)).expect("source opens");
    let transfer = source
        .authorize_compressed_transfer(&part_uri(PHOTO))
        .expect("transfer")
        .expect("its capture verifies");
    let copied = PackURI::new(COPIED).expect("URI");
    let shown = stored_bytes();
    let mut destination = OpcPackage::from_vec(destination_archive()).expect("destination");
    destination
        .try_add_part(Box::new(Forwarding {
            inner: BlobPart::with_compressed_transfer(copied, transfer),
            shown: std::sync::Arc::new(shown.clone()),
        }))
        .expect("forwarding part");
    let output = PackageWriter::to_bytes(&destination).expect("publish");
    assert_eq!(read_member(&output, COPIED), shown);
    assert_eq!(raw_member(&output, &COPIED[1..]).method, 8);
}

/// The source archive with the photo's data-descriptor CRC flipped, so the
/// descriptor disagrees with the central record while the payload itself is
/// intact.
fn disagreeing_descriptor_archive() -> Vec<u8> {
    let mut archive = source_archive(Mode::DeflatedDescriptor, false);
    let (_start, end) = compressed_range(&archive, PHOTO);
    // The descriptor follows the payload: optional signature, then CRC.
    let crc = if u32_at(&archive, end) == 0x0807_4b50 {
        end + 4
    } else {
        end
    };
    archive[crc] ^= 0xff;
    archive
}

#[test]
fn a_data_descriptor_that_disagrees_is_refused() {
    // Every `OpcPackage` ingress is in-memory, and since merge 0759 the
    // in-memory (slice) admission indexes each member's local framing,
    // descriptor included, before the package exists. A descriptor that
    // disagrees with its central record is therefore refused at open, with a
    // typed error, so no package that could issue a transfer holds it. A
    // positional (source-backed) open does not check descriptors: it admits
    // the member and refuses at the first read.
    let error = OpcPackage::from_vec(disagreeing_descriptor_archive())
        .expect_err("a disagreeing descriptor is refused");
    assert!(
        matches!(error, OpcError::ZipError(_)),
        "the refusal is typed: {error:?}"
    );
}

#[test]
fn the_transfer_capture_never_captures_a_disagreeing_descriptor() {
    // The transfer copies a member's compressed span verbatim, so its own
    // capture must refuse this framing whatever the open checked. The open
    // above no longer lets such a member reach it, so the capture is driven
    // here through the positional index `owned_transfer_index` builds, which,
    // like every positional index, admits the member.
    let capture = |archive: Vec<u8>| {
        let length = archive.len() as u64;
        let limits = ReadLimits::default();
        let index = soapberry_zip::office::IndexedArchive::from_reader_with_limits(
            std::sync::Arc::new(archive),
            length,
            limits.zip_limits(),
        )
        .expect("the positional index admits the member");
        let entry_id = index.entry_id(PHOTO).expect("photo member");
        let member = super::TransferMember {
            index: &index,
            entry_id,
            limits,
        };
        super::verified_capture(&member, &photo_bytes()).map(|captured| captured.is_some())
    };
    // Control: the same member with its descriptor intact is captured, so the
    // refusal below is the descriptor's.
    assert!(
        capture(source_archive(Mode::DeflatedDescriptor, false)).expect("intact member"),
        "an intact descriptor-bearing member is captured"
    );
    let captured = capture(disagreeing_descriptor_archive());
    assert!(
        matches!(captured, Ok(false)),
        "a disagreeing descriptor is a content fault, never a capture: {captured:?}"
    );
}

#[test]
fn the_transfer_rechecks_the_read_limits_the_archive_was_admitted_under() {
    let source_bytes = source_archive(Mode::DeflatedSized, false);
    let member = raw_member(&source_bytes, PHOTO);
    let donor = OpcPackage::new();
    for (limits, resource) in [
        (
            ReadLimits::builder()
                .max_part_bytes(1024)
                .and_then(crate::ReadLimitsBuilder::build)
                .expect("limit"),
            ReadResource::PartBytes,
        ),
        (
            ReadLimits::builder()
                .max_archive_compressed_bytes(
                    u64::try_from(member.compressed.len()).expect("size") - 1,
                )
                .and_then(crate::ReadLimitsBuilder::build)
                .expect("limit"),
            ReadResource::ArchiveCompressedBytes,
        ),
    ] {
        // Admit the archive under the default policy, then bind the policy
        // under test: the transfer must re-check the member against it.
        let mut package = OpcPackage::from_vec_reusing_payloads(
            source_bytes.clone(),
            ReadLimits::default(),
            &donor,
        )
        .expect("eager open");
        package.source_limits = limits;
        let error = package
            .authorize_compressed_transfer(&part_uri(PHOTO))
            .expect_err("a member above the limit is refused");
        assert!(
            matches!(error, OpcError::ReadLimit { resource: found, .. } if found == resource),
            "{error:?}"
        );
    }
}

#[test]
fn replacing_or_retyping_a_transferred_part_discards_its_capture() {
    let source =
        OpcPackage::from_vec(source_archive(Mode::DeflatedSized, false)).expect("source opens");
    let transfer = source
        .authorize_compressed_transfer(&part_uri(PHOTO))
        .expect("transfer")
        .expect("its capture verifies");
    let copied = PackURI::new(COPIED).expect("URI");

    let replacement = stored_bytes();
    let mut replaced = BlobPart::with_compressed_transfer(copied.clone(), transfer.clone());
    replaced.set_blob(replacement.clone());
    let plain = BlobPart::new(copied.clone(), "image/png".to_owned(), replacement.clone());
    let replaced_output = publish_into_destination(replaced);
    assert_eq!(replaced_output, publish_into_destination(plain));
    assert_eq!(read_member(&replaced_output, COPIED), replacement);

    let mut retyped = BlobPart::with_compressed_transfer(copied.clone(), transfer);
    retyped
        .set_content_type("image/x-litchi".to_owned())
        .expect("retype");
    let plain = BlobPart::new(copied, "image/x-litchi".to_owned(), photo_bytes());
    assert_eq!(
        publish_into_destination(retyped),
        publish_into_destination(plain)
    );
}

#[test]
fn a_transferred_part_that_replaces_a_source_member_is_regenerated_from_its_capture() {
    let source_bytes = source_archive(Mode::DeflatedSized, false);
    let source = OpcPackage::from_vec(source_bytes.clone()).expect("source opens");
    let transfer = source
        .authorize_compressed_transfer(&part_uri(STORED))
        .expect("transfer")
        .expect("its capture verifies");
    // The destination is another open of the same archive whose photo is
    // replaced by the transferred stored payload under the same name.
    let mut destination = OpcPackage::from_vec(source_bytes.clone()).expect("destination");
    let photo = part_uri(PHOTO);
    assert!(destination.remove_part(&photo));
    destination
        .try_add_part(Box::new(BlobPart::with_compressed_transfer(
            photo.clone(),
            transfer,
        )))
        .expect("replacement part");
    let output = PackageWriter::to_bytes(&destination).expect("publish");
    let published = raw_member(&output, PHOTO);
    assert_eq!(
        published.compressed,
        raw_member(&source_bytes, STORED).compressed
    );
    assert_fresh_sized_framing(&published, 0);
    assert_eq!(read_member(&output, &format!("/{PHOTO}")), stored_bytes());
    // Every other member is still copied byte for byte.
    assert_eq!(
        raw_member(&output, VECTOR),
        raw_member(&source_bytes, VECTOR)
    );
    // A clone publishes the same bytes.
    assert_eq!(
        PackageWriter::to_bytes(&destination.clone()).expect("publish clone"),
        output
    );
}

#[test]
fn the_full_writer_still_publishes_a_transferred_part_from_its_decoded_bytes() {
    let source =
        OpcPackage::from_vec(source_archive(Mode::DeflatedSized, false)).expect("source opens");
    let transfer = source
        .authorize_compressed_transfer(&part_uri(PHOTO))
        .expect("transfer")
        .expect("its capture verifies");
    let mut authored = OpcPackage::new();
    let copied = PackURI::new(COPIED).expect("URI");
    authored
        .try_add_part(Box::new(BlobPart::with_compressed_transfer(
            copied, transfer,
        )))
        .expect("part");
    let output = PackageWriter::to_bytes(&authored).expect("full writer publishes");
    assert_eq!(read_member(&output, COPIED), photo_bytes());
}

/// Insert `padding` zero bytes after one sized member's compressed payload,
/// fixing its compressed sizes and every later offset. A Deflate decoder stops
/// at the final block, so the ordinary reader still decodes the member.
fn pad_member_payload(archive: &[u8], name: &str, padding: usize) -> Vec<u8> {
    let zip = ZipArchive::from_slice(archive).expect("parse ZIP");
    let directory = usize::try_from(zip.directory_offset()).expect("offset");
    let eocd = usize::try_from(zip.eocd_offset()).expect("offset");
    let (local, central) = central_offset(archive, name);
    let (_start, payload_end) = compressed_range(archive, name);
    let grow = |bytes: &mut [u8], offset: usize| {
        let value = u32_at(bytes, offset) + u32::try_from(padding).expect("padding");
        bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
    };
    let mut output = archive[..payload_end].to_vec();
    output.extend(std::iter::repeat_n(0_u8, padding));
    output.extend_from_slice(&archive[payload_end..]);
    let shift = |offset: usize| {
        if offset >= payload_end {
            offset + padding
        } else {
            offset
        }
    };
    grow(&mut output, local + 18);
    grow(&mut output, shift(central) + 20);
    let mut record = shift(directory);
    let end = shift(eocd);
    while record < end {
        if usize::try_from(u32_at(&output, record + 42)).expect("offset") >= payload_end {
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

/// Bytes after a member's final Deflate block: the ordinary reader decodes
/// the member, the headers are consistent, and only the capture's
/// consumed-input check disproves it. That is a deterministic `Ok(None)`, not
/// an error, so a caller recompresses the part.
#[test]
fn bytes_after_the_final_block_disprove_the_capture_deterministically() {
    let archive = pad_member_payload(&source_archive(Mode::DeflatedSized, false), PHOTO, 16);
    let source = OpcPackage::from_vec(archive).expect("the padded archive opens");
    let photo = part_uri(PHOTO);
    assert_eq!(
        source.get_part(&photo).expect("lenient decode").blob(),
        photo_bytes().as_slice()
    );
    assert!(
        source
            .compressed_transfer_size(&photo)
            .expect("part")
            .is_some(),
        "the headers are consistent"
    );
    for _ in 0..2 {
        assert!(
            source
                .authorize_compressed_transfer(&photo)
                .expect("a content fault is not an error")
                .is_none()
        );
    }
    // Another member of the same archive still transfers.
    assert!(
        source
            .authorize_compressed_transfer(&part_uri(STORED))
            .expect("stored member")
            .is_some()
    );
}

/// A Deflate stream padded with empty stored blocks decodes to the payload
/// but is far larger than any encoder would emit; the header-only size guard
/// makes it ineligible, so it is recompressed rather than published padded.
#[test]
fn a_member_padded_with_empty_stored_blocks_is_not_eligible() {
    let payload = stored_bytes();
    let build = |empty_blocks: usize| -> Vec<u8> {
        let mut stream = Vec::new();
        for _ in 0..empty_blocks {
            stream.extend_from_slice(&[0x00, 0x00, 0x00, 0xff, 0xff]);
        }
        let len = u16::try_from(payload.len()).expect("one stored block");
        stream.push(0x01);
        stream.extend_from_slice(&len.to_le_bytes());
        stream.extend_from_slice(&(!len).to_le_bytes());
        stream.extend_from_slice(&payload);
        let content_types = content_types(false);
        let package_relationships = package_relationships(false);
        let mut archive = ZipArchiveWriter::new(Vec::new());
        for (name, bytes) in [
            ("[Content_Types].xml", content_types.as_slice()),
            ("_rels/.rels", package_relationships.as_slice()),
            ("doc/main.xml", MAIN_XML),
            ("doc/_rels/main.xml.rels", MAIN_RELATIONSHIPS),
            (STORED, payload.as_slice()),
            (VECTOR, VECTOR_XML),
            (LINKED, payload.as_slice()),
            ("doc/media/_rels/linked.png.rels", LINKED_RELATIONSHIPS),
        ] {
            archive.write_stored_file(name, bytes).expect("member");
        }
        archive
            .write_precompressed_file(
                PHOTO,
                CompressionMethod::Deflate,
                soapberry_zip::crc32(&payload),
                payload.len() as u64,
                &stream,
            )
            .expect("padded member");
        archive.finish().expect("archive")
    };
    let photo = part_uri(PHOTO);
    // 16 KiB decoded allows 20 block headers plus 64 bytes of slack.
    let within = OpcPackage::from_vec(build(15)).expect("opens");
    assert_eq!(
        within.get_part(&photo).expect("decode").blob(),
        payload.as_slice()
    );
    assert!(
        within
            .compressed_transfer_size(&photo)
            .expect("part")
            .is_some()
    );
    for padded in [build(17), build(20_000)] {
        let source = OpcPackage::from_vec(padded).expect("opens");
        assert_eq!(
            source.get_part(&photo).expect("decode").blob(),
            payload.as_slice()
        );
        assert!(
            source
                .compressed_transfer_size(&photo)
                .expect("part")
                .is_none()
        );
        assert!(is_ineligible(
            &source
                .authorize_compressed_transfer(&photo)
                .expect_err("an ineligible member is never captured")
        ));
    }
}

/// A member framed with ZIP64 sizes and a ZIP64 data descriptor transfers,
/// and is republished with fresh ZIP32 known-size framing.
#[test]
fn a_zip64_framed_member_transfers_with_fresh_zip32_framing() {
    let payload = stored_bytes();
    let content_types = content_types(false);
    let package_relationships = package_relationships(false);
    let mut archive = ZipArchiveWriter::new(Vec::new());
    for (name, bytes) in [
        ("[Content_Types].xml", content_types.as_slice()),
        ("_rels/.rels", package_relationships.as_slice()),
        ("doc/main.xml", MAIN_XML),
        ("doc/_rels/main.xml.rels", MAIN_RELATIONSHIPS),
    ] {
        archive.write_stored_file(name, bytes).expect("member");
    }
    let (mut entry, config) = archive
        .new_file(PHOTO)
        .compression_method(CompressionMethod::Store)
        .zip64(true)
        .start()
        .expect("zip64 member");
    let mut data = config.wrap(&mut entry);
    data.write_all(&payload).expect("payload");
    let (_, descriptor) = data.finish().expect("payload finish");
    entry.finish(descriptor).expect("member finish");
    for (name, bytes) in [
        (STORED, payload.as_slice()),
        (VECTOR, VECTOR_XML),
        (LINKED, payload.as_slice()),
        ("doc/media/_rels/linked.png.rels", LINKED_RELATIONSHIPS),
    ] {
        archive.write_stored_file(name, bytes).expect("member");
    }
    let source_bytes = archive.finish().expect("archive");
    let source_member = raw_member(&source_bytes, PHOTO);
    assert_eq!(
        source_member.local_sizes,
        [u32::MAX, u32::MAX],
        "ZIP64 sizes"
    );
    assert_ne!(source_member.flags & 0x08, 0, "a ZIP64 data descriptor");

    let source = OpcPackage::from_vec(source_bytes).expect("source opens");
    let transfer = source
        .authorize_compressed_transfer(&part_uri(PHOTO))
        .expect("zip64 member")
        .expect("its capture verifies");
    let copied = PackURI::new(COPIED).expect("URI");
    let output = publish_into_destination(BlobPart::with_compressed_transfer(copied, transfer));
    let published = raw_member(&output, &COPIED[1..]);
    assert_eq!(published.compressed, payload);
    assert_fresh_sized_framing(&published, 0);
    assert!(published.local_extra.is_empty(), "no ZIP64 extra is needed");
    assert_eq!(read_member(&output, COPIED), payload);
}

/// Two raw names that normalize to one member name are refused when the
/// package is opened, so a transfer can never capture a different member
/// than the one its part was decoded from.
#[test]
fn members_whose_names_normalize_to_one_name_are_refused_at_open() {
    let mut archive = source_archive(Mode::DeflatedSized, false);
    let (local, central) = central_offset(&archive, STORED);
    // `doc/media/stored.png` and `doc/media//photo.png` have equal length;
    // the second normalizes to `doc/media/photo.png`, the photo's own name.
    let alias = b"doc/media//photo.png";
    assert_eq!(alias.len(), STORED.len());
    archive[local + 30..local + 30 + alias.len()].copy_from_slice(alias);
    archive[central + 46..central + 46 + alias.len()].copy_from_slice(alias);
    let error = OpcPackage::from_vec(archive).expect_err("an ambiguous member name is refused");
    assert!(
        matches!(&error, OpcError::ZipError(message) if message.contains("duplicate")),
        "{error:?}"
    );
}

#[test]
fn the_admitted_read_limits_are_reported_while_the_source_archive_is_retained() {
    let limits = ReadLimits::builder()
        .max_part_bytes(1 << 20)
        .and_then(crate::ReadLimitsBuilder::build)
        .expect("limit");
    let mut package =
        OpcPackage::from_vec_with_limits(source_archive(Mode::DeflatedSized, false), limits)
            .expect("source opens");
    assert_eq!(package.source_read_limits(), Some(limits));
    // An edit revokes exact publication but keeps the archive and its policy.
    package
        .get_part_mut(&part_uri(STORED))
        .expect("part")
        .set_blob(b"edited".to_vec());
    assert!(!package.is_unmodified_owned_source());
    assert_eq!(package.source_read_limits(), Some(limits));
    assert_eq!(package.clone().source_read_limits(), Some(limits));
    assert_eq!(OpcPackage::new().source_read_limits(), None);
}

/// soapberry-zip's own writer, like zlib at its default memory level, frames
/// incompressible data in 16 KiB stored blocks: 5 bytes per 16 KiB. That is
/// more than one header per 65,535 bytes plus a small constant allows for any
/// member above about 273 KiB, so the size guard budgets one header per
/// 4 KiB; such a member stays eligible and transfers.
#[test]
fn an_incompressible_member_from_a_default_encoder_stays_eligible() {
    let photo = pseudo_random_bytes(1024 * 1024, 0x0742_4096);
    let content_types = content_types(false);
    let package_relationships = package_relationships(false);
    let stored = stored_bytes();
    let archive = streaming_archive(&[
        ("[Content_Types].xml", &content_types, Mode::DeflatedSized),
        ("_rels/.rels", &package_relationships, Mode::DeflatedSized),
        ("doc/main.xml", MAIN_XML, Mode::DeflatedSized),
        (
            "doc/_rels/main.xml.rels",
            MAIN_RELATIONSHIPS,
            Mode::DeflatedSized,
        ),
        (PHOTO, &photo, Mode::DeflatedSized),
        (STORED, &stored, Mode::Stored),
        (VECTOR, VECTOR_XML, Mode::DeflatedSized),
        (LINKED, &stored, Mode::DeflatedSized),
        (
            "doc/media/_rels/linked.png.rels",
            LINKED_RELATIONSHIPS,
            Mode::DeflatedSized,
        ),
    ]);
    let member = raw_member(&archive, PHOTO);
    assert_eq!(member.method, 8);
    let decoded = u64::try_from(photo.len()).expect("size");
    let compressed = u64::try_from(member.compressed.len()).expect("size");
    // One 5-byte header per 16 KiB stored block, and a few bytes to end the
    // stream.
    let framing = compressed - decoded;
    let headers = 5 * decoded.div_ceil(16 * 1024);
    assert!((headers..=headers + 8).contains(&framing), "{framing}");
    assert!(framing > 5 * decoded.div_ceil(65_535) + 64);
    let source = OpcPackage::from_vec(archive).expect("source opens");
    let part = part_uri(PHOTO);
    assert_eq!(
        source.compressed_transfer_size(&part).expect("part"),
        Some(compressed)
    );
    let transfer = source
        .authorize_compressed_transfer(&part)
        .expect("capture")
        .expect("verified");
    assert_eq!(transfer.compressed_size(), compressed);
}

/// A deferred part's own decode index caches its refusal as an `OpcError`
/// that no longer says whether the bytes caused it. Eligibility classifies
/// the archive again through the package's transfer index, so a refusal the
/// bytes did not cause is not taken for a verdict, while the part's own
/// decode still reports it.
#[test]
fn a_refused_decode_index_is_not_taken_for_a_verdict_about_the_bytes() {
    let archive = source_archive(Mode::DeflatedSized, false);
    let compressed = raw_member(&archive, PHOTO).compressed.len();
    let source = OpcPackage::from_vec(archive).expect("source opens");
    let photo = part_uri(PHOTO);
    let deferred = source
        .preservation
        .as_deref()
        .and_then(|provenance| provenance.parts.get(&photo))
        .and_then(|part| part.blob.as_deferred())
        .expect("a deferred photo");
    deferred
        .source()
        .refuse_index_for_test(OpcError::ZipError("Buffer size too small".to_owned()));
    assert_eq!(
        source
            .compressed_transfer_size(&photo)
            .expect("classified again"),
        Some(u64::try_from(compressed).expect("size"))
    );
    assert!(matches!(
        source.authorize_compressed_transfer(&photo),
        Err(OpcError::ZipError(message)) if message == "Buffer size too small"
    ));
}

/// A caller-defined part is not one of this crate's own part types; a
/// package holding one says so, and the reopen of its serialization does not.
#[test]
fn a_caller_defined_part_is_not_built_in() {
    #[derive(Clone)]
    struct Custom(BlobPart);
    impl Part for Custom {
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
        fn rels(&self) -> &crate::Relationships {
            self.0.rels()
        }
        fn rels_mut(&mut self) -> &mut crate::Relationships {
            self.0.rels_mut()
        }
        fn set_blob(&mut self, blob: Vec<u8>) {
            self.0.set_blob(blob);
        }
    }
    let mut package =
        OpcPackage::from_vec(source_archive(Mode::DeflatedSized, false)).expect("source opens");
    assert!(package.holds_only_built_in_parts());
    package
        .get_part_mut(&part_uri(STORED))
        .expect("part")
        .set_blob(b"edited".to_vec());
    assert!(
        package.holds_only_built_in_parts(),
        "an edit keeps the built-in part"
    );
    let custom = PackURI::new("/doc/custom.bin").expect("URI");
    package
        .try_add_part(Box::new(Custom(BlobPart::new(
            custom.clone(),
            "application/octet-stream".to_owned(),
            b"custom".to_vec(),
        ))))
        .expect("added");
    assert!(!package.holds_only_built_in_parts());
    assert!(!package.clone().holds_only_built_in_parts());
    let reopened = OpcPackage::from_vec(PackageWriter::to_bytes(&package).expect("serialized"))
        .expect("reopened");
    assert!(reopened.holds_only_built_in_parts());
    assert_eq!(reopened.get_part(&custom).expect("part").blob(), b"custom");
}
