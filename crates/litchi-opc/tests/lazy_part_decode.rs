#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused lazy-decoding assertions intentionally panic on fixture errors"
)]

//! Contract tests for the owned-source lazy OPC payload representation.
//!
//! These tests keep the package small enough to make the read boundary
//! observable. Metadata iteration and exact-source publication must not inflate
//! ordinary members; a fallible part access must inflate exactly the selected
//! member; and a failed first access must remain a typed, stable refusal.

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
    let mut writer = StreamingArchiveWriter::new();
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

#[test]
fn metadata_and_exact_noop_leave_owned_payloads_cold() {
    let source = archive();
    let package = OpcPackage::from_vec(source.clone()).expect("open lazy owned package");

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

    let output = PackageWriter::to_bytes(&package).expect("exact no-op publication");
    assert_eq!(output, source);
    assert_eq!(package.deferred_decode_counters(), Some((0, 0)));
}

#[test]
fn fallible_access_decodes_only_the_selected_owned_part() {
    let package = OpcPackage::from_vec(archive()).expect("open lazy owned package");
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
}

#[test]
fn failed_first_access_is_stable_and_never_becomes_an_empty_generated_part() {
    let source = corrupt_bad_member(archive());
    let bad = uri(BAD_URI);

    let package = OpcPackage::from_vec(source.clone()).expect("corruption is deferred");
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
    let mut edited = OpcPackage::from_vec(source.clone()).expect("reopen corrupt source");
    edited
        .get_part_mut(&uri(GOOD_URI))
        .expect("good part remains editable")
        .set_blob(b"<changed/>".to_vec());
    let output = PackageWriter::to_bytes(&edited).expect("targeted edit publication");
    let reopened = OpcPackage::from_vec(output).expect("reopen targeted output");
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
