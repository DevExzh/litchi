use super::*;
use std::path::Path;

#[test]
#[ignore = "requires an external DOC fixture"]
fn test_open_package() {
    let result = Package::open("test.doc");
    assert!(result.is_ok());
}

#[test]
#[ignore = "requires an external DOC fixture"]
fn test_invalid_file() {
    // Create a non-DOC file
    std::fs::write("test_invalid.tmp", b"Not a DOC file").unwrap();
    let result = Package::open("test_invalid.tmp");
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

    let mut package = Package::open(&path).unwrap();
    assert!(matches!(package.document(), Err(Error::PasswordRequired)));

    let mut package = Package::open(&path).unwrap();
    assert!(matches!(
        package
            .document_with_options(OpenOptions::default().with_password("wrong".to_owned().into())),
        Err(Error::InvalidPassword)
    ));

    let mut package = Package::open(path).unwrap();
    let document = package
        .document_with_options(OpenOptions::default().with_password("tika".to_owned().into()))
        .unwrap();
    assert!(!document.text().unwrap().trim().is_empty());
}

#[test]
fn opens_apache_poi_cryptoapi_document() {
    let path = poi_fixture("password_password_cryptoapi.doc");

    let mut package = Package::open(&path).unwrap();
    assert!(matches!(package.document(), Err(Error::PasswordRequired)));

    let mut package = Package::open(&path).unwrap();
    assert!(matches!(
        package
            .document_with_options(OpenOptions::default().with_password("wrong".to_owned().into())),
        Err(Error::InvalidPassword)
    ));

    let mut package = Package::open(path).unwrap();
    let document = package
        .document_with_options(OpenOptions::default().with_password("password".to_owned().into()))
        .unwrap();
    assert!(!document.text().unwrap().trim().is_empty());
}

#[test]
fn ordinary_doc_exposes_an_empty_dataspaces_profile() {
    let mut package = Package::open(poi_fixture("test.doc")).unwrap();
    assert!(package.data_spaces().unwrap().is_none());
    assert!(package.data_spaces_snapshot().unwrap().is_none());
}
