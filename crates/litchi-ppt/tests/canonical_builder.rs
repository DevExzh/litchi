#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "contract assertions"
)]

use litchi_core::Position;
use litchi_ppt::{Builder, OpenOptions, Package, SourceLimits};

#[test]
fn minimal_packages_are_valid_deterministic_and_publish_exactly() {
    let first = Package::new().unwrap();
    let second = Package::new().unwrap();
    assert_eq!(first.presentation().unwrap().slide_count(), 0);
    let bytes = first.to_bytes().unwrap();
    assert_eq!(bytes, second.to_bytes().unwrap());
    assert_eq!(
        Package::from_bytes(&bytes).unwrap().to_bytes().unwrap(),
        bytes
    );
    let mut sink = Vec::new();
    first.write_to(&mut sink).unwrap();
    assert_eq!(sink, bytes);
}

#[test]
fn builder_publishes_detached_packages_that_use_canonical_edits() {
    let mut builder = Builder::new();
    let slide = builder.add_slide().unwrap();
    builder
        .add_textbox(slide, 10, 20, 200, 80, "first")
        .unwrap();
    let first = builder.build().unwrap();
    let before = first.to_bytes().unwrap();
    let slide = builder.add_slide().unwrap();
    builder
        .add_textbox(slide, 10, 20, 200, 80, "second")
        .unwrap();
    let second = builder.build().unwrap();
    assert_eq!(first.presentation().unwrap().slide_count(), 1);
    assert_eq!(second.presentation().unwrap().slide_count(), 2);
    assert_eq!(first.to_bytes().unwrap(), before);
    assert_eq!(
        second.to_bytes().unwrap(),
        builder.build().unwrap().to_bytes().unwrap()
    );
    let mut edit = second.edit().unwrap();
    edit.move_slide(Position::new(0), Position::new(1)).unwrap();
    let commit = edit.commit().unwrap();
    let texts: Vec<_> = commit
        .package()
        .presentation()
        .unwrap()
        .slides()
        .unwrap()
        .iter()
        .map(|slide| slide.text().unwrap().to_owned())
        .collect();
    assert_eq!(texts, vec!["second", "first"]);
    assert_eq!(
        commit
            .package()
            .apply(&commit.patch().inverse())
            .unwrap()
            .into_package()
            .to_bytes()
            .unwrap(),
        second.to_bytes().unwrap()
    );
}

#[test]
fn build_limits_refuse_output_without_destroying_the_authoring_state() {
    let mut builder = Builder::new();
    let slide = builder.add_slide().unwrap();
    builder
        .add_textbox(slide, 0, 0, 100, 100, "bounded")
        .unwrap();
    let expected = builder.build().unwrap().to_bytes().unwrap();
    let limits = SourceLimits::builder()
        .max_input_bytes(expected.len() - 1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        builder.build_with_options(OpenOptions::new().source_limits(limits)),
        Err(litchi_ppt::Error::Limit {
            kind: litchi_ppt::SourceLimitKind::InputBytes,
            ..
        })
    ));
    assert_eq!(builder.build().unwrap().to_bytes().unwrap(), expected);
    let limits = SourceLimits::builder()
        .max_input_bytes(expected.len())
        .unwrap()
        .build()
        .unwrap();
    let package = builder
        .build_with_options(OpenOptions::new().source_limits(limits))
        .unwrap();
    assert_eq!(package.source_limits(), limits);
    assert_eq!(package.to_bytes().unwrap(), expected);
}

#[cfg(feature = "encryption")]
#[test]
fn encrypted_build_requires_and_consumes_explicit_entropy() {
    use litchi_ppt::{EncryptionEntropy, EncryptionProfile};
    fn configured() -> Builder {
        let mut builder = Builder::new();
        let slide = builder.add_slide().unwrap();
        builder
            .add_textbox(slide, 0, 0, 100, 100, "secret")
            .unwrap();
        builder
            .set_password(
                "password",
                EncryptionProfile::CryptoApiRc4 { key_bits: 128 },
            )
            .unwrap();
        builder
    }
    let mut builder = configured();
    assert!(matches!(
        builder.build(),
        Err(litchi_ppt::Error::Authoring(
            litchi_ppt::writer::WriteError::EncryptionEntropyRequired
        ))
    ));
    builder.set_encryption_entropy(EncryptionEntropy::from_bytes([7; 16], [9; 16]));
    let package = builder.build().unwrap();
    let encrypted = package.to_encrypted_bytes().unwrap();
    assert!(package.is_encrypted());
    assert!(package.to_bytes().is_err());
    assert!(builder.build().is_err(), "entropy is single-use");
    let mut same = configured();
    same.set_encryption_entropy(EncryptionEntropy::from_bytes([7; 16], [9; 16]));
    assert_eq!(
        same.build().unwrap().to_encrypted_bytes().unwrap(),
        encrypted
    );
    let reopened =
        Package::from_bytes_with_options(&encrypted, OpenOptions::new().password("password"))
            .unwrap();
    assert_eq!(
        reopened.presentation().unwrap().text().unwrap(),
        package.presentation().unwrap().text().unwrap()
    );

    builder.set_encryption_entropy(EncryptionEntropy::from_bytes([11; 16], [13; 16]));
    let tiny = SourceLimits::builder()
        .max_input_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(
        builder
            .build_with_options(OpenOptions::new().source_limits(tiny))
            .is_err()
    );
    assert!(
        builder.build().is_err(),
        "failed admission also consumes entropy"
    );
}

#[test]
fn authoring_errors_retain_typed_causes_and_redact_content() {
    let error = litchi_ppt::Error::Authoring(litchi_ppt::writer::WriteError::InvalidData(
        "sensitive authored text".to_owned(),
    ));
    assert!(
        std::error::Error::source(&error)
            .unwrap()
            .is::<litchi_ppt::writer::WriteError>()
    );
    assert!(!format!("{error:?} {error}").contains("sensitive authored text"));
}

#[cfg(feature = "encryption")]
#[test]
fn plaintext_build_does_not_consume_pending_encryption_entropy() {
    let mut builder = Builder::new();
    builder.set_encryption_entropy(litchi_ppt::EncryptionEntropy::from_bytes(
        [29; 16], [31; 16],
    ));
    assert!(!builder.build().unwrap().is_encrypted());
    builder
        .set_password(
            "secret",
            litchi_ppt::EncryptionProfile::CryptoApiRc4 { key_bits: 128 },
        )
        .unwrap();
    assert!(builder.build().unwrap().is_encrypted());
    assert!(matches!(
        builder.build(),
        Err(litchi_ppt::Error::Authoring(
            litchi_ppt::writer::WriteError::EncryptionEntropyRequired
        ))
    ));
}
