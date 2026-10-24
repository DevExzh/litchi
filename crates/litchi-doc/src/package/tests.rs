use super::{Error, OpenOptions};
use crate::raw::{OpenOutcome, Reader, open_package};
use litchi_cfb::OleFile;
use std::io::Cursor;
use std::path::Path;

#[test]
#[ignore = "requires an external DOC fixture"]
fn test_open_package() {
    let result = Reader::open("test.doc");
    assert!(result.is_ok());
}

#[test]
#[ignore = "requires an external DOC fixture"]
fn test_invalid_file() {
    // Create a non-DOC file
    std::fs::write("test_invalid.tmp", b"Not a DOC file").unwrap();
    let result = Reader::open("test_invalid.tmp");
    assert!(result.is_err());
    std::fs::remove_file("test_invalid.tmp").ok();
}

fn poi_fixture(name: &str) -> std::path::PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/document")
        .join(name)
}

#[test]
fn opens_apache_poi_binary_rc4_document() {
    let path = poi_fixture("password_tika_binaryrc4.doc");

    let mut package = Reader::open(&path).unwrap();
    assert!(matches!(package.document(), Err(Error::PasswordRequired)));

    let mut package = Reader::open(&path).unwrap();
    assert!(matches!(
        package.document_with_options(OpenOptions::default().password("wrong".to_owned().into())),
        Err(Error::InvalidPassword)
    ));

    let mut package = Reader::open(path).unwrap();
    let document = package
        .document_with_options(OpenOptions::default().password("tika".to_owned().into()))
        .unwrap();
    assert!(!document.text().unwrap().trim().is_empty());
}

#[test]
fn opens_apache_poi_cryptoapi_document() {
    let path = poi_fixture("password_password_cryptoapi.doc");

    let mut package = Reader::open(&path).unwrap();
    assert!(matches!(package.document(), Err(Error::PasswordRequired)));

    let mut package = Reader::open(&path).unwrap();
    assert!(matches!(
        package.document_with_options(OpenOptions::default().password("wrong".to_owned().into())),
        Err(Error::InvalidPassword)
    ));

    let mut package = Reader::open(path).unwrap();
    let document = package
        .document_with_options(OpenOptions::default().password("password".to_owned().into()))
        .unwrap();
    assert!(!document.text().unwrap().trim().is_empty());
}

#[test]
fn raw_open_outcome_retains_encrypted_cfb_for_later_password_ingress() {
    let bytes = std::fs::read(poi_fixture("password_tika_binaryrc4.doc")).unwrap();
    let ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let ole = match open_package(ole, OpenOptions::default()).unwrap() {
        OpenOutcome::PasswordRequired(ole) => *ole,
        OpenOutcome::Opened(_) => panic!("encrypted fixture unexpectedly opened without password"),
    };
    let package = match open_package(
        ole,
        OpenOptions::default().password("tika".to_owned().into()),
    )
    .unwrap()
    {
        OpenOutcome::Opened(package) => package,
        OpenOutcome::PasswordRequired(_) => panic!("correct password remained locked"),
    };
    assert!(!package.document().text().unwrap().trim().is_empty());
}

#[test]
fn ordinary_doc_exposes_an_empty_dataspaces_profile() {
    let mut package = Reader::open(poi_fixture("test.doc")).unwrap();
    assert!(package.data_spaces().unwrap().is_none());
    assert!(package.data_spaces_snapshot().unwrap().is_none());
}
