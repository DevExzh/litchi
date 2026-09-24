#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "ingress admission assertions intentionally panic on fixture errors"
)]

//! ZIP admission on every OPC ingress, and when each framing fault surfaces.
//!
//! Merge 0759 carried the spec-gap branch's admission onto every ingress: the
//! declared entry count is checked before indexing, and encryption flagged in
//! the central *or* the local header is refused, on members nothing reads.
//! What else an ingress checks at open depends on how it reads the archive:
//!
//! * the in-memory (slice) ingress indexes each member's local framing, so a
//!   data descriptor that disagrees with its central record is refused at
//!   open;
//! * the positional (source-backed) ingress and the catalog probe read only
//!   each member's 8-byte local-header prefix at admission, so the same
//!   member is admitted, and the positional ingress refuses it when it is
//!   first read;
//! * no ingress compares a sized member's local-header CRC with the central
//!   record at open;
//! * a payload that fails its CRC under consistent headers is refused at its
//!   first decode on every ingress (ADR 0030), which for the eager ingress is
//!   at open.
//!
//! These tests pin that behaviour, which record 0759's judgement 3 describes.

use std::io::Cursor;
use std::sync::Arc;

use litchi_opc::{OpcError, OpcPackage, PackURI, ReadLimits, SourceBackedPackage};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#;
const RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
const DOCUMENT: &str = "<doc>benign document payload, long enough to deflate a little</doc>";
const BLOB_MEMBER: &str = "custom/blob.bin";
const BLOB_URI: &str = "/custom/blob.bin";
const BLOB: &[u8] = b"an unrelated payload that nothing reads at open time";

/// A small package whose `custom/blob.bin` member is either streamed (bit 3,
/// zero local CRC and sizes, a data descriptor) or sized.
fn package(streaming_blob: bool) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_deflated_sized("[Content_Types].xml", CONTENT_TYPES.as_bytes())
        .unwrap();
    writer
        .write_deflated_sized("_rels/.rels", RELS.as_bytes())
        .unwrap();
    writer
        .write_deflated("word/document.xml", DOCUMENT.as_bytes())
        .unwrap();
    if streaming_blob {
        writer.write_deflated(BLOB_MEMBER, BLOB).unwrap();
    } else {
        writer.write_deflated_sized(BLOB_MEMBER, BLOB).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn find(archive: &[u8], signature: &[u8], name_len_at: usize, name_at: usize, name: &str) -> usize {
    archive
        .windows(4)
        .enumerate()
        .find_map(|(offset, window)| {
            if window != signature {
                return None;
            }
            let len = u16::from_le_bytes([
                archive[offset + name_len_at],
                archive[offset + name_len_at + 1],
            ]);
            let start = offset + name_at;
            (archive.get(start..start + usize::from(len))? == name.as_bytes()).then_some(offset)
        })
        .expect("header present")
}

fn local(archive: &[u8], name: &str) -> usize {
    find(archive, b"PK\x03\x04", 26, 30, name)
}

fn central(archive: &[u8], name: &str) -> usize {
    find(archive, b"PK\x01\x02", 28, 46, name)
}

fn set_flags(archive: &mut [u8], name: &str, local_bits: u16, central_bits: u16) {
    let l = local(archive, name);
    let c = central(archive, name);
    let local_flags = u16::from_le_bytes([archive[l + 6], archive[l + 7]]) | local_bits;
    let central_flags = u16::from_le_bytes([archive[c + 8], archive[c + 9]]) | central_bits;
    archive[l + 6..l + 8].copy_from_slice(&local_flags.to_le_bytes());
    archive[c + 8..c + 10].copy_from_slice(&central_flags.to_le_bytes());
}

fn open_all(archive: &[u8]) -> Vec<(&'static str, Result<(), String>)> {
    let map = |result: Result<(), OpcError>| result.map_err(|error| format!("{error:?}"));
    vec![
        (
            "from_vec (deferred)",
            map(OpcPackage::from_vec(archive.to_vec()).map(drop)),
        ),
        (
            "from_shared_vec_with_limits (deferred)",
            map(OpcPackage::from_shared_vec_with_limits(
                Arc::new(archive.to_vec()),
                ReadLimits::default(),
            )
            .map(drop)),
        ),
        (
            "from_bytes (eager)",
            map(OpcPackage::from_bytes(archive).map(drop)),
        ),
        (
            "source-backed from_vec",
            map(SourceBackedPackage::from_vec(archive.to_vec()).map(drop)),
        ),
        (
            "catalog probe",
            map(
                litchi_opc::probe_package_catalog_from_reader(&mut Cursor::new(archive.to_vec()))
                    .map(drop),
            ),
        ),
    ]
}

fn first_positional_read(archive: Vec<u8>) -> Result<(), OpcError> {
    let package = SourceBackedPackage::from_vec(archive)?;
    package
        .part(&PackURI::new(BLOB_URI).unwrap())?
        .data()
        .map(drop)
}

/// The in-memory ingresses: each indexes every member's local framing at
/// open, data descriptors included.
const IN_MEMORY_INGRESSES: [&str; 3] = [
    "from_vec (deferred)",
    "from_shared_vec_with_limits (deferred)",
    "from_bytes (eager)",
];

/// The ingresses that read only a bounded prefix of each member at open.
const PREFIX_ONLY_INGRESSES: [&str; 2] = ["source-backed from_vec", "catalog probe"];

#[test]
fn a_disagreeing_descriptor_is_refused_at_open_in_memory_and_at_first_read_positionally() {
    // Flip only the descriptor's CRC: the central record and the payload
    // still agree, the descriptor does not.
    let mut archive = package(true);
    let c = central(&archive, BLOB_MEMBER);
    let crc = u32::from_le_bytes(archive[c + 16..c + 20].try_into().unwrap());
    let l = local(&archive, BLOB_MEMBER);
    let relative = archive[l..c]
        .windows(4)
        .rposition(|window| window == crc.to_le_bytes())
        .expect("descriptor CRC present");
    archive[l + relative] ^= 1;

    for (ingress, outcome) in open_all(&archive) {
        if IN_MEMORY_INGRESSES.contains(&ingress) {
            let error = outcome.expect_err("an in-memory ingress refuses at open");
            assert!(error.starts_with("ZipError"), "{ingress}: {error}");
        } else {
            assert!(PREFIX_ONLY_INGRESSES.contains(&ingress), "{ingress}");
            outcome.unwrap_or_else(|error| panic!("{ingress} checks only a prefix: {error}"));
        }
    }
    let error = first_positional_read(archive).expect_err("the first positional read refuses");
    assert!(matches!(error, OpcError::ZipError(_)), "{error:?}");
}

#[test]
fn a_sized_members_local_crc_is_not_compared_at_open() {
    // Without a data descriptor, only the local header repeats the CRC. No
    // ingress compares it with the central record at open; a central CRC the
    // payload fails is refused where the payload is first decoded (at open
    // only by the eager ingress, which decodes everything).
    let mut archive = package(false);
    let l = local(&archive, BLOB_MEMBER);
    archive[l + 14] ^= 1;
    for (ingress, outcome) in open_all(&archive) {
        outcome.unwrap_or_else(|error| panic!("{ingress} compared the local CRC: {error}"));
    }

    let mut archive = package(false);
    let c = central(&archive, BLOB_MEMBER);
    archive[c + 16] ^= 1;
    for (ingress, outcome) in open_all(&archive) {
        if ingress == "from_bytes (eager)" {
            let error = outcome.expect_err("the eager ingress decodes at open");
            assert!(error.starts_with("ZipError"), "{error}");
        } else {
            outcome.unwrap_or_else(|error| panic!("{ingress} decoded at open: {error}"));
        }
    }
    let package = OpcPackage::from_vec(archive.clone()).unwrap();
    assert!(matches!(
        package.get_part(&PackURI::new(BLOB_URI).unwrap()),
        Err(OpcError::ZipError(_))
    ));
    let error = first_positional_read(archive).expect_err("the first positional read refuses");
    assert!(matches!(error, OpcError::ZipError(_)), "{error:?}");
}

#[test]
fn every_ingress_refuses_encryption_on_an_unread_member() {
    for streaming in [false, true] {
        for bit in [1_u16, 1 << 6] {
            for local_only in [false, true] {
                let mut archive = package(streaming);
                set_flags(
                    &mut archive,
                    BLOB_MEMBER,
                    bit,
                    if local_only { 0 } else { bit },
                );
                for (ingress, outcome) in open_all(&archive) {
                    let error = match outcome {
                        Ok(()) => panic!(
                            "{ingress} admitted encryption bit {bit:#x} \
                             local_only={local_only} streaming={streaming}"
                        ),
                        Err(error) => error,
                    };
                    assert!(
                        error.starts_with("ZipError"),
                        "{ingress}: untyped refusal {error}"
                    );
                }
            }
        }
    }
}

#[test]
fn benign_streaming_members_with_zero_local_crc_open_everywhere() {
    let archive = package(true);
    let l = local(&archive, BLOB_MEMBER);
    let flags = u16::from_le_bytes([archive[l + 6], archive[l + 7]]);
    assert_ne!(
        flags & (1 << 3),
        0,
        "fixture member must use a data descriptor"
    );
    assert_eq!(
        &archive[l + 14..l + 18],
        &[0, 0, 0, 0],
        "streaming local CRC is zero"
    );
    for (ingress, outcome) in open_all(&archive) {
        outcome.unwrap_or_else(|error| {
            panic!("{ingress} refused a benign streaming package: {error}")
        });
    }
    let package = OpcPackage::from_vec(archive.clone()).unwrap();
    let uri = PackURI::new(BLOB_URI).unwrap();
    assert_eq!(package.get_part(&uri).unwrap().blob(), BLOB);
    first_positional_read(archive).expect("positional read of a benign streaming member");
}

/// Flip the member's CRC consistently: central, and the local header or the
/// data descriptor, whichever carries it. The headers then agree with each
/// other and only the payload disagrees.
fn consistent_crc_flip(archive: &mut [u8], name: &str) {
    let c = central(archive, name);
    let crc = u32::from_le_bytes(archive[c + 16..c + 20].try_into().unwrap());
    let flipped = (crc ^ 1).to_le_bytes();
    archive[c + 16..c + 20].copy_from_slice(&flipped);
    let l = local(archive, name);
    let flags = u16::from_le_bytes([archive[l + 6], archive[l + 7]]);
    if flags & (1 << 3) == 0 {
        archive[l + 14..l + 18].copy_from_slice(&flipped);
    } else {
        let relative = archive[l..c]
            .windows(4)
            .rposition(|window| window == crc.to_le_bytes())
            .expect("descriptor CRC present");
        let at = l + relative;
        archive[at..at + 4].copy_from_slice(&flipped);
    }
}

#[test]
fn payload_crc_mismatch_with_consistent_headers_moves_to_first_decode() {
    for streaming in [false, true] {
        let mut archive = package(streaming);
        consistent_crc_flip(&mut archive, BLOB_MEMBER);
        let package = OpcPackage::from_vec(archive.clone()).unwrap_or_else(|error| {
            panic!("deferred open must not decode (streaming={streaming}): {error:?}")
        });
        let uri = PackURI::new(BLOB_URI).unwrap();
        let error = package
            .get_part(&uri)
            .map(drop)
            .expect_err("first decode refuses the CRC");
        assert!(
            matches!(error, OpcError::ZipError(_) | OpcError::IoError(_)),
            "{error:?}"
        );
        let error = first_positional_read(archive).expect_err("first positional read refuses");
        assert!(
            matches!(error, OpcError::ZipError(_) | OpcError::IoError(_)),
            "{error:?}"
        );
    }
}
