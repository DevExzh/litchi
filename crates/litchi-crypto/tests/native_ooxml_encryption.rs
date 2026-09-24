//! Native Office and LibreOffice OOXML encryption corpus checks.
//!
//! The corpus is kept in the repository's LibreOffice and Apache POI checkouts
//! because these files are producer evidence rather than synthetic litchi
//! fixtures. Run explicitly with `cargo test ... -- --ignored`.

#![cfg(feature = "ooxml")]
#![allow(
    clippy::expect_used,
    clippy::panic,
    reason = "ignored corpus tests report the exact missing or incompatible fixture"
)]

use std::io::{Cursor, Read};
use std::path::{Path, PathBuf};

use litchi_crypto::ooxml::{Error, IntegrityStatus, Kind, Limits, Mode, inspect_with, open_with};
use zip::ZipArchive;

const LIBREOFFICE_ROOT_ENV: &str = "LITCHI_LIBREOFFICE_CORE";
const POI_FIXTURE_ENV: &str = "LITCHI_POI60320_FIXTURE";
const PASSWORD: &str = "abc";
const POI_PASSWORD: &str = "Test001!!";

fn libreoffice_fixture_root() -> PathBuf {
    let source_root = std::env::var_os(LIBREOFFICE_ROOT_ENV)
        .map(PathBuf::from)
        .unwrap_or_else(|| {
            Path::new(env!("CARGO_MANIFEST_DIR")).join("../../3rdparty/libreoffice-core")
        });
    source_root.join("sw/qa/extras/ooxmlexport/data")
}

fn libreoffice_fixture(name: &str) -> PathBuf {
    libreoffice_fixture_root().join(name)
}

fn poi_fixture() -> PathBuf {
    std::env::var_os(POI_FIXTURE_ENV)
        .map(PathBuf::from)
        .unwrap_or_else(|| {
            Path::new(env!("CARGO_MANIFEST_DIR"))
                .join("../../3rdparty/poi/test-data/poifs/60320-protected.xlsx")
        })
}

fn assert_complete_opc_zip(bytes: &[u8], required_part: &str, fixture_name: &str) {
    let mut archive = ZipArchive::new(Cursor::new(bytes))
        .unwrap_or_else(|error| panic!("{fixture_name} clear payload is not a ZIP: {error}"));
    assert!(!archive.is_empty(), "{fixture_name} ZIP has no entries");

    // Reading every entry verifies local headers, decompression, and CRCs.
    for index in 0..archive.len() {
        let mut entry = archive
            .by_index(index)
            .unwrap_or_else(|error| panic!("{fixture_name} ZIP entry {index}: {error}"));
        std::io::copy(&mut entry, &mut std::io::sink())
            .unwrap_or_else(|error| panic!("{fixture_name} ZIP entry {index} read: {error}"));
    }

    for required in ["[Content_Types].xml", "_rels/.rels", required_part] {
        let mut entry = archive
            .by_name(required)
            .unwrap_or_else(|error| panic!("{fixture_name} missing OPC part {required}: {error}"));
        let mut contents = Vec::new();
        entry
            .read_to_end(&mut contents)
            .unwrap_or_else(|error| panic!("{fixture_name} OPC part {required} read: {error}"));
        assert!(
            !contents.is_empty(),
            "{fixture_name} OPC part {required} is empty"
        );
    }
}

fn assert_wrong_password(bytes: &[u8], limits: &Limits, fixture_name: &str, password: &str) {
    assert!(
        matches!(
            open_with(bytes.to_vec(), password, limits),
            Err(Error::Password)
        ),
        "{fixture_name} accepted an incorrect password"
    );
}

fn assert_native_fixture(
    name: &str,
    expected_mode: Mode,
    expected_integrity: IntegrityStatus,
    required_part: &str,
    limits: &Limits,
    password: &str,
) {
    let path = libreoffice_fixture(name);
    let bytes = std::fs::read(&path).expect("read native OOXML encryption fixture");
    assert_eq!(
        inspect_with(&bytes, limits).expect("classify native OOXML encryption fixture"),
        Kind::Encrypted(expected_mode),
        "native fixture profile: {name}"
    );
    let opened =
        open_with(bytes.clone(), password, limits).expect("open native OOXML encryption fixture");
    assert_eq!(
        opened.mode(),
        Some(expected_mode),
        "native fixture mode: {name}"
    );
    assert_eq!(
        opened.integrity(),
        Some(expected_integrity),
        "native fixture integrity status: {name}"
    );
    assert_complete_opc_zip(opened.bytes(), required_part, name);
    assert_wrong_password(&bytes, limits, name, "wrong-native-password");
}

#[test]
#[ignore = "requires the native Office/LibreOffice corpus checkout"]
fn native_office_profiles_open_with_default_graph_validation() {
    let limits = Limits::default();
    assert_native_fixture(
        "Encrypted_MSO2007_abc.docx",
        Mode::Standard,
        IntegrityStatus::Unauthenticated,
        "word/document.xml",
        &limits,
        PASSWORD,
    );
    assert_native_fixture(
        "Encrypted_MSO2010_abc.docx",
        Mode::Agile,
        IntegrityStatus::Authenticated,
        "word/document.xml",
        &limits,
        PASSWORD,
    );
    assert_native_fixture(
        "Encrypted_MSO2013_abc.docx",
        Mode::AgileAes256Sha512,
        IntegrityStatus::Authenticated,
        "word/document.xml",
        &limits,
        PASSWORD,
    );
}

#[test]
#[ignore = "requires the native Office/LibreOffice corpus checkout"]
fn libreoffice_standard_without_dataspaces_requires_explicit_compatibility() {
    let path = libreoffice_fixture("Encrypted_LO_Standard_abc.docx");
    let bytes = std::fs::read(&path).expect("read native LibreOffice fixture");
    let limits = Limits {
        allow_missing_data_spaces: true,
        ..Limits::default()
    };
    assert!(matches!(
        open_with(bytes.clone(), PASSWORD, &Limits::default()),
        Err(Error::Malformed(_))
    ));
    assert_eq!(
        inspect_with(&bytes, &limits).expect("classify LibreOffice Standard fixture"),
        Kind::Encrypted(Mode::Standard)
    );
    let opened = open_with(bytes.clone(), PASSWORD, &limits)
        .expect("open LibreOffice Standard fixture with explicit compatibility");
    assert_eq!(opened.mode(), Some(Mode::Standard));
    assert_eq!(opened.integrity(), Some(IntegrityStatus::Unauthenticated));
    assert_complete_opc_zip(
        opened.bytes(),
        "word/document.xml",
        "Encrypted_LO_Standard_abc.docx",
    );
    assert_wrong_password(
        &bytes,
        &limits,
        "Encrypted_LO_Standard_abc.docx",
        "wrong-native-password",
    );
}

#[test]
#[ignore = "requires the Apache POI 60320 fixture checkout"]
fn apache_poi_60320_mixed_wrapper_profile_opens() {
    let path = poi_fixture();
    let bytes = std::fs::read(&path).expect("read Apache POI 60320 fixture");
    let limits = Limits::default();
    assert_eq!(
        inspect_with(&bytes, &limits).expect("classify Apache POI 60320 fixture"),
        Kind::Encrypted(Mode::Agile)
    );
    let opened =
        open_with(bytes.clone(), POI_PASSWORD, &limits).expect("open Apache POI 60320 fixture");
    assert_eq!(opened.mode(), Some(Mode::Agile));
    assert_eq!(opened.integrity(), Some(IntegrityStatus::Authenticated));
    assert_complete_opc_zip(opened.bytes(), "xl/workbook.xml", "60320-protected.xlsx");
    assert_wrong_password(
        &bytes,
        &limits,
        "60320-protected.xlsx",
        "wrong-poi-password",
    );
}
