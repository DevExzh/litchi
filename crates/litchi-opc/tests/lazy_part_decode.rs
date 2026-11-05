#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "focused lazy-decoding assertions intentionally panic on fixture errors"
)]

//! Contract tests for the owned-source lazy OPC payload representation.
//!
//! These tests keep the package small enough to make the read boundary
//! observable. Metadata iteration and exact-source publication must not inflate
//! ordinary members; a fallible part access must inflate exactly the selected
//! member; and a failed first access must remain a typed, stable refusal.

use litchi_opc::SourceLimits;
use litchi_opc::{OpcError, OpcPackage, PackURI, PackageWriter};
use soapberry_zip::office::StreamingArchiveWriter;

const MANIFEST: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/></Types>"#;
const ROOT_RELATIONSHIPS: &[u8] =
    br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#;
const GOOD_MEMBER: &str = "custom/good.xml";
const BAD_MEMBER: &str = "custom/bad.bin";
const GOOD_URI: &str = "/custom/good.xml";
const BAD_URI: &str = "/custom/bad.bin";
const GOOD_PAYLOAD: &[u8] = b"<good/>";
const BAD_PAYLOAD: &[u8] = b"lazy-corrupt-payload";

fn uri(value: &str) -> PackURI {
    PackURI::new(value).expect("valid test part URI")
}

fn archive() -> Vec<u8> {
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_stored("[Content_Types].xml", MANIFEST)
        .expect("write manifest");
    writer
        .write_stored("_rels/.rels", ROOT_RELATIONSHIPS)
        .expect("write root relationships");
    writer
        .write_deflated_sized(GOOD_MEMBER, GOOD_PAYLOAD)
        .expect("write good member");
    writer
        .write_stored(BAD_MEMBER, BAD_PAYLOAD)
        .expect("write bad member");
    writer.finish_to_bytes().expect("finish test archive")
}

fn corrupt_bad_member(mut bytes: Vec<u8>) -> Vec<u8> {
    let offset = bytes
        .windows(BAD_PAYLOAD.len())
        .position(|window| window == BAD_PAYLOAD)
        .expect("stored test payload must be present in the archive");
    bytes[offset] ^= 1;
    bytes
}

fn misdeclare_bad_member(mut bytes: Vec<u8>) -> Vec<u8> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == b"PK\x05\x06")
        .expect("fixture EOCD must be present");
    let count = usize::from(u16::from_le_bytes(
        bytes[eocd + 10..eocd + 12]
            .try_into()
            .expect("fixture EOCD count must fit"),
    ));
    let mut central = usize::try_from(u32::from_le_bytes(
        bytes[eocd + 16..eocd + 20]
            .try_into()
            .expect("fixture central-directory offset must fit"),
    ))
    .expect("fixture central-directory offset must fit this platform");
    for _ in 0..count {
        assert_eq!(&bytes[central..central + 4], b"PK\x01\x02");
        let name_len = usize::from(u16::from_le_bytes(
            bytes[central + 28..central + 30]
                .try_into()
                .expect("fixture name length must fit"),
        ));
        let extra_len = usize::from(u16::from_le_bytes(
            bytes[central + 30..central + 32]
                .try_into()
                .expect("fixture extra length must fit"),
        ));
        let comment_len = usize::from(u16::from_le_bytes(
            bytes[central + 32..central + 34]
                .try_into()
                .expect("fixture comment length must fit"),
        ));
        if &bytes[central + 46..central + 46 + name_len] == BAD_MEMBER.as_bytes() {
            let declared = u32::from_le_bytes(
                bytes[central + 24..central + 28]
                    .try_into()
                    .expect("fixture uncompressed length must fit"),
            )
            .checked_add(1)
            .expect("fixture length must not overflow");
            bytes[central + 24..central + 28].copy_from_slice(&declared.to_le_bytes());

            let local = usize::try_from(u32::from_le_bytes(
                bytes[central + 42..central + 46]
                    .try_into()
                    .expect("fixture local-header offset must fit"),
            ))
            .expect("fixture local-header offset must fit this platform");
            assert_eq!(&bytes[local..local + 4], b"PK\x03\x04");
            bytes[local + 22..local + 26].copy_from_slice(&declared.to_le_bytes());
            let flags = u16::from_le_bytes(
                bytes[local + 6..local + 8]
                    .try_into()
                    .expect("fixture flags must fit"),
            );
            if flags & 8 != 0 {
                let local_name_len = usize::from(u16::from_le_bytes(
                    bytes[local + 26..local + 28]
                        .try_into()
                        .expect("fixture local name length must fit"),
                ));
                let local_extra_len = usize::from(u16::from_le_bytes(
                    bytes[local + 28..local + 30]
                        .try_into()
                        .expect("fixture local extra length must fit"),
                ));
                let compressed_len = usize::try_from(u32::from_le_bytes(
                    bytes[central + 20..central + 24]
                        .try_into()
                        .expect("fixture compressed length must fit"),
                ))
                .expect("fixture compressed length must fit this platform");
                let descriptor = local + 30 + local_name_len + local_extra_len + compressed_len;
                let descriptor_size = if &bytes[descriptor..descriptor + 4] == b"PK\x07\x08" {
                    descriptor + 12
                } else {
                    descriptor + 8
                };
                bytes[descriptor_size..descriptor_size + 4]
                    .copy_from_slice(&declared.to_le_bytes());
            }
            return bytes;
        }
        central += 46 + name_len + extra_len + comment_len;
    }
    panic!("fixture member {BAD_MEMBER} is missing");
}

#[test]
fn metadata_and_exact_noop_leave_owned_payloads_cold() {
    let source = archive();
    let package = OpcPackage::from_vec(source.clone(), SourceLimits::default())
        .expect("open lazy owned package");

    assert_eq!(package.deferred_decode_counters(), Some((0, 0)));
    let names: Vec<_> = package
        .iter_parts()
        .map(|part| {
            assert!(!part.payload_is_decoded());
            part.partname().as_str().to_owned()
        })
        .collect();
    assert_eq!(names.len(), 2);
    assert_eq!(package.deferred_decode_counters(), Some((0, 0)));

    let mut lengths: Vec<_> = package
        .iter_parts()
        .map(|part| part.payload_len().expect("admitted payload length"))
        .collect();
    lengths.sort_unstable();
    assert_eq!(lengths, vec![GOOD_PAYLOAD.len(), BAD_PAYLOAD.len()]);
    assert_eq!(package.deferred_decode_counters(), Some((0, 0)));

    let output = PackageWriter::to_bytes(&package).expect("exact no-op publication");
    assert_eq!(output, source);
    assert_eq!(package.deferred_decode_counters(), Some((0, 0)));
}

#[test]
fn fallible_access_decodes_only_the_selected_owned_part() {
    let package =
        OpcPackage::from_vec(archive(), SourceLimits::default()).expect("open lazy owned package");
    let selected = package.get_part(&uri(GOOD_URI)).expect("good part");
    assert_eq!(selected.blob(), GOOD_PAYLOAD);
    assert_eq!(
        package.deferred_decode_counters(),
        Some((1, GOOD_PAYLOAD.len() as u64))
    );
    assert!(
        package
            .iter_parts()
            .find(|part| part.partname().as_str() == BAD_URI)
            .is_some_and(|part| !part.payload_is_decoded())
    );

    // The capacity view reads only what is decoded, and decodes nothing.
    for part in package.iter_parts() {
        let capacity = part.decoded_payload_capacity();
        if part.partname().as_str() == GOOD_URI {
            assert_eq!(capacity, Some(selected.blob_arc().capacity()));
        } else {
            assert_eq!(capacity, None);
        }
    }
    assert_eq!(
        package.deferred_decode_counters(),
        Some((1, GOOD_PAYLOAD.len() as u64))
    );
}

#[test]
fn part_metadata_separates_absence_from_a_failed_decode_without_decoding() {
    let package = OpcPackage::from_vec(corrupt_bad_member(archive()), SourceLimits::default())
        .expect("corruption is deferred");

    let bad = package
        .part_metadata(&uri(BAD_URI))
        .expect("a corrupt member is still present");
    assert_eq!(bad.partname().as_str(), BAD_URI);
    assert_eq!(bad.content_type(), "application/octet-stream");
    assert!(!bad.payload_is_decoded());
    assert_eq!(bad.payload_len().unwrap(), BAD_PAYLOAD.len());
    assert_eq!(package.deferred_decode_counters(), Some((0, 0)));
    let good = package
        .part_metadata(&uri("/CUSTOM/Good.XML"))
        .expect("name resolution matches get_part");
    assert_eq!(good.partname().as_str(), GOOD_URI);
    assert!(package.part_metadata(&uri("/custom/absent.xml")).is_none());
    assert_eq!(package.deferred_decode_counters(), Some((0, 0)));

    // The decode refusal is still there for a reader, and is not absence.
    assert!(matches!(
        package.get_part(&uri(BAD_URI)),
        Err(OpcError::ZipError(_))
    ));
}

#[test]
fn misdeclared_lazy_member_is_rejected_after_a_bounded_decode_probe() {
    let source = misdeclare_bad_member(archive());
    let package = OpcPackage::from_vec(source.clone(), SourceLimits::default())
        .expect("misdeclaration is deferred");
    let bad = package
        .part_metadata(&uri(BAD_URI))
        .expect("the misdeclared member is still present");

    // Metadata admission reports the central-directory declaration without
    // inflating the member. The declared + 1 probe is only exercised on the
    // first payload access, and the source accounting remains untouched when
    // the decoded length fails verification.
    assert_eq!(bad.payload_len().unwrap(), BAD_PAYLOAD.len() + 1);
    assert_eq!(package.deferred_decode_counters(), Some((0, 0)));
    let error = match package.get_part(&uri(BAD_URI)) {
        Ok(_) => panic!("a misdeclared member must refuse on first access"),
        Err(error) => error,
    };
    assert!(matches!(error, OpcError::ZipError(_)));
    assert_eq!(package.deferred_decode_counters(), Some((0, 0)));

    // The failed read does not publish a replacement or alter exact-source
    // publication; no decoded payload is retained as a fallback.
    assert_eq!(PackageWriter::to_bytes(&package).unwrap(), source);
}

#[test]
fn batch_relationship_to_a_corrupt_present_part_targets_an_existing_part() {
    let mut package = OpcPackage::from_vec(corrupt_bad_member(archive()), SourceLimits::default())
        .expect("corruption is deferred");
    let root = uri("/");
    let content_types = package
        .source_content_types(SourceLimits::default())
        .expect("manifest token");
    let expected = package
        .source_relationships(&root, SourceLimits::default())
        .expect("root token");
    let replacement = expected
        .with_relationship(
            "urn:test:bad",
            "custom/bad.bin",
            "rIdBad",
            litchi_opc::TargetMode::Internal,
            1 << 20,
        )
        .expect("edge to the corrupt member");

    // Whether a relationship target exists is a question about names; the
    // target's payload is neither read nor needed, so its decode refusal
    // must not be reported as a missing part.
    package
        .try_add_parts_with_source_tokens(
            content_types.bytes(),
            &content_types,
            &expected,
            &replacement,
            Vec::new(),
        )
        .expect("a present part is an existing target");
    assert_eq!(package.deferred_decode_counters(), Some((0, 0)));
    assert!(matches!(
        package.get_part(&uri(BAD_URI)),
        Err(OpcError::ZipError(_))
    ));
}

#[test]
fn failed_first_access_is_stable_and_never_becomes_an_empty_generated_part() {
    let source = corrupt_bad_member(archive());
    let bad = uri(BAD_URI);

    let package = OpcPackage::from_vec(source.clone(), SourceLimits::default())
        .expect("corruption is deferred");
    let first = match package.get_part(&bad) {
        Ok(_) => panic!("corrupt member must refuse on first access"),
        Err(error) => error,
    };
    let first_text = first.to_string();
    let second = match package.get_part(&bad) {
        Ok(_) => panic!("the recorded refusal must be stable"),
        Err(error) => error,
    };
    assert_eq!(second.to_string(), first_text);
    assert!(matches!(first, OpcError::ZipError(_)));

    // An untouched package is intentionally allowed to copy its exact source
    // without decoding the member. The failure is therefore still present in
    // the output and is observed when the member is subsequently accessed.
    assert_eq!(PackageWriter::to_bytes(&package).unwrap(), source);

    // A targeted edit must validate and regenerate the changed member while
    // copying the failed, untouched member byte-for-byte. It must never turn
    // the failed payload into the empty fallback used by the infallible trait
    // method.
    let mut edited = OpcPackage::from_vec(source.clone(), SourceLimits::default())
        .expect("reopen corrupt source");
    edited
        .get_part_mut(&uri(GOOD_URI))
        .expect("good part remains editable")
        .set_blob(b"<changed/>".to_vec());
    let output = PackageWriter::to_bytes(&edited).expect("targeted edit publication");
    let reopened =
        OpcPackage::from_vec(output, SourceLimits::default()).expect("reopen targeted output");
    assert_eq!(
        reopened.get_part(&uri(GOOD_URI)).unwrap().blob(),
        b"<changed/>"
    );
    let preserved = match reopened.get_part(&bad) {
        Ok(_) => panic!("copied corrupt member must remain unreadable"),
        Err(error) => error,
    };
    assert_eq!(preserved.to_string(), first_text);
}
