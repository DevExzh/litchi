//! Regression coverage for encrypted Office files supplied by Apache POI.
//!
//! The passwords here are the public test credentials that accompany the
//! upstream fixtures. Keep them in `Zeroizing` owners and never include them
//! in diagnostics or assertions.

#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use std::fs::File;
use std::path::{Path, PathBuf};

use litchi_cfb::OleFile;
use litchi_crypto::{rc4, spaces};
use zeroize::Zeroizing;

const FILEPASS_RECORD: u16 = 0x002F;
const CRYPTOAPI_VERSION_MINOR: u16 = 2;

fn fixture(path: &[&str]) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(path.iter().collect::<PathBuf>())
}

fn stream(path: &Path, name: &str) -> Vec<u8> {
    let mut ole = OleFile::open(File::open(path).expect("open fixture")).expect("open OLE");
    ole.open_stream(&[name]).expect("read fixture stream")
}

/// Locate an exactly-sized `CryptoAPI` header and leave its validation to the
/// public parser. The surrounding PPT/DOC record grammars belong to their
/// respective format crates, not to this format-neutral crate.
fn cryptoapi_header(stream: &[u8]) -> &[u8] {
    for offset in 0..stream.len().saturating_sub(12) {
        let Some(version) = stream.get(offset..offset + 4) else {
            continue;
        };
        let major = u16::from_le_bytes(version[..2].try_into().expect("two-byte major"));
        let minor = u16::from_le_bytes(version[2..].try_into().expect("two-byte minor"));
        if !(2..=4).contains(&major) || minor != CRYPTOAPI_VERSION_MINOR {
            continue;
        }
        let Some(header_size_bytes) = stream.get(offset + 8..offset + 12) else {
            continue;
        };
        let declared_header_size =
            u32::from_le_bytes(header_size_bytes.try_into().expect("four-byte size"));
        let Ok(header_size) = usize::try_from(declared_header_size) else {
            continue;
        };
        let Some(length) = 12usize
            .checked_add(header_size)
            .and_then(|value| value.checked_add(60))
        else {
            continue;
        };
        let Some(candidate) = stream.get(offset..offset.saturating_add(length)) else {
            continue;
        };
        if rc4::parse_header(candidate).is_ok() {
            return candidate;
        }
    }
    panic!("fixture has no complete CryptoAPI header");
}

fn assert_cryptoapi_password(header: &[u8], expected_password: &str) {
    let parsed = rc4::parse_header(header).expect("validated fixture header");
    let correct_password = Zeroizing::new(expected_password.to_owned());
    let incorrect = Zeroizing::new("incorrect fixture password".to_owned());

    assert!(
        rc4::verify(&parsed, correct_password.as_str())
            .expect("verify known fixture password")
            .is_some(),
        "fixture's public test password must authenticate"
    );
    assert!(
        rc4::verify(&parsed, incorrect.as_str())
            .expect("reject an incorrect fixture password")
            .is_none(),
        "incorrect password must not yield key material"
    );
}

#[test]
fn authenticates_apache_poi_cryptoapi_powerpoint_fixture() {
    let path = fixture(&["poi", "test-data", "slideshow", "cryptoapi-proc2356.ppt"]);
    let document = stream(&path, "PowerPoint Document");

    assert_cryptoapi_password(cryptoapi_header(&document), "crypto");
}

#[test]
fn authenticates_apache_poi_cryptoapi_word_fixture() {
    let path = fixture(&[
        "poi",
        "test-data",
        "document",
        "password_password_cryptoapi.doc",
    ]);
    let table = stream(&path, "1Table");

    assert_cryptoapi_password(cryptoapi_header(&table), "password");
}

#[test]
fn inspects_apache_poi_xor_workbook_without_claiming_unsupported_decryption() {
    let path = fixture(&["poi", "test-data", "spreadsheet", "xor-encryption-abc.xls"]);
    let bytes = std::fs::read(&path).expect("read fixture");
    let workbook = stream(&path, "Workbook");

    // XOR FILEPASS is intentionally not a public litchi-crypto primitive:
    // ADR-0006 keeps weak legacy encryption decode-only. Still pin this real
    // fixture's on-disk classification so it cannot silently turn into an
    // ordinary workbook or an unrelated DataSpaces envelope.
    let mut offset = 0;
    let mut filepass = None;
    while offset < workbook.len() {
        let header = workbook
            .get(offset..offset + 4)
            .expect("complete BIFF record header");
        let record = u16::from_le_bytes(header[..2].try_into().expect("record id"));
        let length = usize::from(u16::from_le_bytes(
            header[2..].try_into().expect("record length"),
        ));
        let end = offset
            .checked_add(4)
            .and_then(|value| value.checked_add(length))
            .expect("BIFF record extent");
        let payload = workbook.get(offset + 4..end).expect("complete BIFF record");
        if record == FILEPASS_RECORD {
            filepass = Some(payload);
            break;
        }
        offset = end;
    }

    let filepass_payload = filepass.expect("XOR FILEPASS record");
    assert_eq!(filepass_payload.len(), 6, "XOR FILEPASS exact record size");
    assert_eq!(
        u16::from_le_bytes(filepass_payload[..2].try_into().expect("kind")),
        0
    );
    assert_eq!(
        spaces::inspect_bytes(&bytes).expect("inspect actual CFB fixture"),
        None,
        "legacy XOR workbook must not be mistaken for an IRM DataSpaces graph"
    );
}
