#![cfg(feature = "encryption")]

use litchi_docx::encryption::{Error as CryptoError, Limits, Mode};
use litchi_docx::{Error, Package};

const PASSWORD: &str = "Litchi test password 42!";
const NEW_PASSWORD: &str = "Litchi changed password 7!";

#[test]
fn package_round_trips_the_supported_encryption_profiles() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("encrypted.docx");
    let mut package = Package::new().unwrap();
    package
        .save_encrypted(&path, PASSWORD, Mode::Standard)
        .unwrap();

    let package = Package::open_with_password(&path, PASSWORD).unwrap();
    assert_eq!(package.encryption(), Some(Mode::Standard));
    assert_eq!(package.document().unwrap().paragraph_count().unwrap(), 0);
}

#[test]
fn wrong_password_does_not_mutate_the_source() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("encrypted.docx");
    let mut package = Package::new().unwrap();
    package
        .save_encrypted(&path, PASSWORD, Mode::Agile)
        .unwrap();
    let before = std::fs::read(&path).unwrap();

    assert!(matches!(
        Package::open_with_password(&path, "wrong password"),
        Err(Error::Encryption(CryptoError::Password))
    ));
    assert_eq!(std::fs::read(path).unwrap(), before);
}

#[test]
fn encrypted_sources_require_an_explicit_plaintext_save() {
    let directory = tempfile::tempdir().unwrap();
    let source = directory.path().join("source.docx");
    let target = directory.path().join("target.docx");
    let sentinel = b"existing output must survive";
    let mut package = Package::new().unwrap();
    package
        .save_encrypted(&source, PASSWORD, Mode::Agile)
        .unwrap();
    let mut package = Package::open_with_password(&source, PASSWORD).unwrap();

    std::fs::write(&target, sentinel).unwrap();
    assert!(package.save(&target).is_err());
    assert_eq!(std::fs::read(&target).unwrap(), sentinel);
    package.save_plain(&target).unwrap();
    assert_eq!(&std::fs::read(&target).unwrap()[..2], b"PK");
}

#[test]
fn reencrypted_output_preserves_mode_but_not_the_old_password() {
    let directory = tempfile::tempdir().unwrap();
    let source = directory.path().join("source.docx");
    let changed = directory.path().join("changed.docx");
    let mut package = Package::new().unwrap();
    package
        .save_encrypted(&source, PASSWORD, Mode::Standard)
        .unwrap();

    let mut package = Package::open_with_password(&source, PASSWORD).unwrap();
    package.save_reencrypted(&changed, NEW_PASSWORD).unwrap();
    assert!(Package::open_with_password(&changed, PASSWORD).is_err());
    let reopened = Package::open_with_password(&changed, NEW_PASSWORD).unwrap();
    assert_eq!(reopened.encryption(), Some(Mode::Standard));
}

#[test]
fn bounded_reader_rejects_input_before_package_parsing() {
    let limits = Limits {
        max_input_bytes: 4,
        ..Limits::default()
    };
    let error = Package::from_reader_with(std::io::Cursor::new(vec![0u8; 5]), PASSWORD, &limits)
        .err()
        .expect("bounded reader must fail");
    assert!(matches!(
        error,
        Error::Encryption(CryptoError::Limit {
            resource: "input",
            actual: 5,
            maximum: 4,
        })
    ));
}
