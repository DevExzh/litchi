#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused physical-reader assertions intentionally fail on fixture errors"
)]

//! Public borrowed-access tests for the physical OPC package reader.
//!
//! The archive is deliberately small.  Lower-layer tests own the exhaustive
//! ZIP descriptor, overlap, and ZIP64 matrix; this file checks only the OPC
//! adapter's public contract and its mapping of those outcomes.

use litchi_opc::phys_pkg::PhysPkgReader;
use litchi_opc::{OpcError, PackURI, ReadLimits, ReadResource};
use soapberry_zip::office::StreamingArchiveWriter;

const STORED_MEMBER: &str = "word/media/image.bin";
const STORED_URI: &str = "/word/media/image.bin";
const STORED_PAYLOAD: &[u8] = b"small stored OPC payload";
const DEFLATED_MEMBER: &str = "word/document.xml";
const DEFLATED_URI: &str = "/word/document.xml";
const DEFLATED_PAYLOAD: &[u8] = b"small deflated OPC payload";

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).expect("test URI must be canonical")
}

fn mixed_archive() -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored(STORED_MEMBER, STORED_PAYLOAD)
        .expect("stored fixture must be writable");
    writer
        .write_deflated(DEFLATED_MEMBER, DEFLATED_PAYLOAD)
        .expect("deflated fixture must be writable");
    writer
        .finish_to_bytes()
        .expect("mixed fixture must be finalized")
}

fn header_offset(
    archive: &[u8],
    signature: &[u8],
    name_length_offset: usize,
    name_offset: usize,
    name: &[u8],
) -> usize {
    archive
        .windows(signature.len())
        .enumerate()
        .find_map(|(offset, candidate)| {
            if candidate != signature {
                return None;
            }
            let length_start = offset.checked_add(name_length_offset)?;
            let length_end = length_start.checked_add(2)?;
            let length_bytes = archive.get(length_start..length_end)?;
            let name_length = usize::from(u16::from_le_bytes([length_bytes[0], length_bytes[1]]));
            let name_start = offset.checked_add(name_offset)?;
            let name_end = name_start.checked_add(name_length)?;
            (archive.get(name_start..name_end)? == name).then_some(offset)
        })
        .expect("fixture must contain the requested ZIP header")
}

fn local_header_offset(archive: &[u8], name: &str) -> usize {
    header_offset(archive, b"PK\x03\x04", 26, 30, name.as_bytes())
}

fn central_header_offset(archive: &[u8], name: &str) -> usize {
    header_offset(archive, b"PK\x01\x02", 28, 46, name.as_bytes())
}

fn stored_payload_offset(archive: &[u8], name: &str) -> usize {
    let local = local_header_offset(archive, name);
    let name_length = usize::from(u16::from_le_bytes([
        archive[local + 26],
        archive[local + 27],
    ]));
    let extra_length = usize::from(u16::from_le_bytes([
        archive[local + 28],
        archive[local + 29],
    ]));
    local + 30 + name_length + extra_length
}

fn mark_encrypted(archive: &mut [u8], name: &str) {
    let local = local_header_offset(archive, name);
    let central = central_header_offset(archive, name);
    let local_flags = u16::from_le_bytes([archive[local + 6], archive[local + 7]]) | 1;
    let central_flags = u16::from_le_bytes([archive[central + 8], archive[central + 9]]) | 1;
    archive[local + 6..local + 8].copy_from_slice(&local_flags.to_le_bytes());
    archive[central + 8..central + 10].copy_from_slice(&central_flags.to_le_bytes());
}

fn assert_zip_refusal(error: OpcError) {
    assert!(
        matches!(&error, OpcError::ZipError(_)),
        "borrowed access must retain a typed ZIP refusal, got {error:?}"
    );
}

#[test]
fn phys_pkg_borrowed_store_returns_source_slice_with_identity() {
    let archive = mixed_archive();
    let reader = PhysPkgReader::new(&archive).expect("fixture must parse");
    let uri = pack(STORED_URI);

    // This annotation also fixes the public contract at plain bytes: no
    // physical ZIP entry or archive type crosses the OPC boundary.
    let borrowed: Option<&[u8]> = reader
        .blob_for_borrowed(&uri)
        .expect("stored member must be readable");
    let borrowed = borrowed.expect("Store member must be borrowable");
    let payload_offset = stored_payload_offset(&archive, STORED_MEMBER);
    let source_payload = &archive[payload_offset..payload_offset + STORED_PAYLOAD.len()];

    assert_eq!(borrowed, STORED_PAYLOAD);
    assert_eq!(borrowed.as_ptr(), source_payload.as_ptr());
    assert_eq!(borrowed.len(), source_payload.len());
}

#[test]
fn phys_pkg_borrowed_normalizes_pack_uri_to_zip_member_name() {
    let archive = mixed_archive();
    let reader = PhysPkgReader::new(&archive).expect("fixture must parse");
    let uri = pack(STORED_URI);

    assert_eq!(uri.membername(), STORED_MEMBER);
    assert_eq!(
        reader
            .blob_for_borrowed(&uri)
            .expect("normalized member must be readable"),
        Some(STORED_PAYLOAD)
    );
}

#[test]
fn phys_pkg_borrowed_deflate_returns_none_and_owned_blob_is_exact() {
    let archive = mixed_archive();
    let reader = PhysPkgReader::new(&archive).expect("fixture must parse");
    let uri = pack(DEFLATED_URI);

    assert_eq!(
        reader
            .blob_for_borrowed(&uri)
            .expect("Deflate lookup must be supported"),
        None
    );
    assert_eq!(
        reader
            .blob_for(&uri)
            .expect("owned Deflate fallback must be readable"),
        DEFLATED_PAYLOAD
    );
}

#[test]
fn phys_pkg_borrowed_missing_part_is_typed_error() {
    let archive = mixed_archive();
    let reader = PhysPkgReader::new(&archive).expect("fixture must parse");
    let missing = pack("/missing.bin");

    let error = reader
        .blob_for_borrowed(&missing)
        .expect_err("missing member must not be treated as an empty result");
    assert!(matches!(
        error,
        OpcError::PartNotFound(name) if name == "/missing.bin"
    ));
}

#[test]
fn phys_pkg_borrowed_rejects_encrypted_store_and_deflate() {
    for (member, uri) in [(STORED_MEMBER, STORED_URI), (DEFLATED_MEMBER, DEFLATED_URI)] {
        let mut archive = mixed_archive();
        mark_encrypted(&mut archive, member);
        let reader = PhysPkgReader::new(&archive).expect("encrypted fixture must parse");

        let error = reader
            .blob_for_borrowed(&pack(uri))
            .expect_err("encrypted members must never publish borrowed bytes");
        assert_zip_refusal(error);
    }
}

#[test]
fn phys_pkg_borrowed_rejects_corrupt_stored_payload_and_crc() {
    let mut payload_corrupt = mixed_archive();
    let payload_offset = stored_payload_offset(&payload_corrupt, STORED_MEMBER);
    payload_corrupt[payload_offset] ^= 0x80;
    let reader = PhysPkgReader::new(&payload_corrupt).expect("corrupt payload must still parse");
    assert_zip_refusal(
        reader
            .blob_for_borrowed(&pack(STORED_URI))
            .expect_err("corrupt stored payload must be refused"),
    );

    let mut crc_corrupt = mixed_archive();
    let local = local_header_offset(&crc_corrupt, STORED_MEMBER);
    let central = central_header_offset(&crc_corrupt, STORED_MEMBER);
    let incorrect_crc = u32::from_le_bytes([
        crc_corrupt[local + 14],
        crc_corrupt[local + 15],
        crc_corrupt[local + 16],
        crc_corrupt[local + 17],
    ]) ^ 0xffff_ffff;
    crc_corrupt[local + 14..local + 18].copy_from_slice(&incorrect_crc.to_le_bytes());
    crc_corrupt[central + 16..central + 20].copy_from_slice(&incorrect_crc.to_le_bytes());
    let reader = PhysPkgReader::new(&crc_corrupt).expect("wrong CRC must still parse");
    assert_zip_refusal(
        reader
            .blob_for_borrowed(&pack(STORED_URI))
            .expect_err("incorrect stored CRC must be refused"),
    );
}

#[test]
fn phys_pkg_borrowed_enforces_part_limit_without_materialization_charge() {
    let archive = mixed_archive();
    let oversized_maximum = (STORED_PAYLOAD.len() - 1) as u64;
    let oversized_limits = ReadLimits::builder()
        .max_part_bytes(oversized_maximum)
        .expect("positive part limit must be accepted")
        .build()
        .expect("part limit profile must be valid");
    let oversized_reader = PhysPkgReader::new_with_limits(&archive, oversized_limits)
        .expect("archive metadata must fit the archive limits");
    let error = oversized_reader
        .blob_for_borrowed(&pack(STORED_URI))
        .expect_err("declared Store bytes above max_part_bytes must be rejected");
    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: ReadResource::PartBytes,
            actual,
            maximum,
        } if actual == STORED_PAYLOAD.len() as u64 && maximum == oversized_maximum
    ));

    let materialization_maximum = STORED_PAYLOAD.len() as u64;
    let materialization_limits = ReadLimits::builder()
        .max_parts(1)
        .expect("one materialized part must be accepted")
        .max_part_bytes(materialization_maximum)
        .expect("stored payload must fit the per-part limit")
        .max_total_part_bytes(materialization_maximum)
        .expect("stored payload must fit the aggregate limit")
        .build()
        .expect("materialization profile must be valid");
    let materialization_reader = PhysPkgReader::new_with_limits(&archive, materialization_limits)
        .expect("archive metadata must fit the materialization limits");

    assert_eq!(
        materialization_reader
            .blob_for_borrowed(&pack(STORED_URI))
            .expect("eligible Store member must be borrowable"),
        Some(STORED_PAYLOAD)
    );
    assert_eq!(
        materialization_reader
            .blob_for(&pack(STORED_URI))
            .expect("first owned read must consume the only materialization slot"),
        STORED_PAYLOAD
    );
    let error = materialization_reader
        .blob_for(&pack(STORED_URI))
        .expect_err("second owned read must be the first exhausted materialization");
    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: ReadResource::Parts,
            actual: 2,
            maximum: 1,
        }
    ));
}

fn borrow_stored_from_source<'data>(source: &'data [u8]) -> &'data [u8] {
    let reader = PhysPkgReader::new(source).expect("fixture must parse");
    reader
        .blob_for_borrowed(&pack(STORED_URI))
        .expect("stored member must be readable")
        .expect("stored member must be borrowable")
}

#[test]
fn phys_pkg_borrowed_slice_lifetime_tracks_input_source() {
    let archive = mixed_archive();
    let borrowed = borrow_stored_from_source(&archive);
    let source_start = archive.as_ptr() as usize;
    let source_end = source_start + archive.len();
    let borrowed_start = borrowed.as_ptr() as usize;
    let borrowed_end = borrowed_start + borrowed.len();

    // The reader can disappear after the helper returns because the slice is
    // tied to the caller-owned source, not to a temporary owned ZIP buffer.
    assert_eq!(borrowed, STORED_PAYLOAD);
    assert!(borrowed_start >= source_start);
    assert!(borrowed_end <= source_end);
}
