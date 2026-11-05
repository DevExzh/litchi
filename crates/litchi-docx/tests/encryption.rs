#![cfg(feature = "encryption")]

use litchi_docx::encryption::{Error as CryptoError, Limits, Mode};
use litchi_docx::{Error, Package};

const PASSWORD: &str = "Litchi test password 42!";
const NEW_PASSWORD: &str = "Litchi changed password 7!";

#[test]
fn package_round_trips_the_supported_encryption_profiles() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("encrypted.docx");
    let package = Package::new().unwrap();
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
    let package = Package::new().unwrap();
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
    let package = Package::new().unwrap();
    package
        .save_encrypted(&source, PASSWORD, Mode::Agile)
        .unwrap();
    let package = Package::open_with_password(&source, PASSWORD).unwrap();

    std::fs::write(&target, sentinel).unwrap();
    assert!(matches!(
        package.save(&target),
        Err(Error::EncryptionPolicy {
            operation: "save",
            ..
        })
    ));
    assert_eq!(std::fs::read(&target).unwrap(), sentinel);
    package.save_plain(&target).unwrap();
    assert_eq!(&std::fs::read(&target).unwrap()[..2], b"PK");
}

#[test]
fn reencrypted_output_preserves_mode_but_not_the_old_password() {
    let directory = tempfile::tempdir().unwrap();
    let source = directory.path().join("source.docx");
    let changed = directory.path().join("changed.docx");
    let package = Package::new().unwrap();
    package
        .save_encrypted(&source, PASSWORD, Mode::Standard)
        .unwrap();

    let package = Package::open_with_password(&source, PASSWORD).unwrap();
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
    let error = Package::from_reader_with_password_and_options(
        std::io::Cursor::new(vec![0u8; 5]),
        PASSWORD,
        litchi_docx::OpenOptions::default().encryption_limits(limits),
    )
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

#[test]
fn cloned_encrypted_package_retains_policy_and_refuses_plain_apply() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("encrypted.docx");
    let plain = Package::new().unwrap();
    let noop = plain.edit().unwrap().commit().unwrap();
    plain
        .save_encrypted(&path, PASSWORD, Mode::Standard)
        .unwrap();
    let encrypted = Package::open_with_password(&path, PASSWORD).unwrap();
    assert!(
        encrypted
            .edit()
            .and_then(|edit| edit.commit().map(|_| ()))
            .is_err()
    );
    let shared = encrypted.clone();
    assert_eq!(shared.encryption(), Some(Mode::Standard));
    assert!(shared.apply(noop.patch()).is_err());
    let mut output = Vec::new();
    assert!(shared.write_to(&mut output).is_err());
    assert!(output.is_empty());
    shared.to_plain_stream(&mut output).unwrap();
    assert!(output.starts_with(b"PK"));
    assert_eq!(encrypted.encryption(), Some(Mode::Standard));
    assert_eq!(shared.encryption(), Some(Mode::Standard));
}

#[test]
fn byte_output_requires_explicit_plaintext_and_retains_provenance() {
    use litchi_ooxml_common::package_encryption::PolicyError;
    for mode in [Mode::Standard, Mode::Agile] {
        let mut builder = litchi_docx::Builder::new().unwrap();
        builder
            .document()
            .unwrap()
            .add_paragraph_with_text("byte output");
        let source = builder.build().unwrap();
        let plain = source.to_bytes().unwrap();
        let encrypted = source.to_encrypted(PASSWORD, mode).unwrap();
        let package = Package::from_bytes_with_password(&encrypted, PASSWORD).unwrap();
        assert!(matches!(package.to_bytes(), Err(Error::EncryptionPolicy {
            operation: "to_bytes", source: PolicyError::OrdinaryOutputWouldRemoveEncryption { mode: retained }
        }) if retained == mode));
        let mut refused = Vec::new();
        assert!(matches!(
            package.write_to(&mut refused),
            Err(Error::EncryptionPolicy {
                operation: "write_to",
                ..
            })
        ));
        assert!(refused.is_empty());
        assert_eq!(package.to_plain_bytes().unwrap(), plain);
        assert_eq!(package.encryption(), Some(mode));
        assert!(matches!(
            package.to_bytes(),
            Err(Error::EncryptionPolicy { .. })
        ));
        let cloned = package.clone();
        assert!(matches!(
            cloned.to_bytes(),
            Err(Error::EncryptionPolicy { .. })
        ));
        assert_eq!(cloned.to_plain_bytes().unwrap(), plain);
    }
}

#[test]
fn reencrypted_output_without_source_mode_is_a_typed_policy_refusal() {
    let package = Package::new().unwrap();
    assert!(matches!(
        package.to_reencrypted(PASSWORD),
        Err(Error::EncryptionPolicy {
            operation: "to_reencrypted",
            source: litchi_ooxml_common::package_encryption::PolicyError::NoRetainedMode,
        })
    ));
    assert!(package.to_bytes().unwrap().starts_with(b"PK"));
}
