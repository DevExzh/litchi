//! Canonical security metadata survives admission without changing publication policy.
#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions fail fast"
)]

use litchi_xls::package::Error;
use litchi_xls::{Builder, EncryptionProfile, OpenOptions, Package, WeakEncryptionPolicy};
use std::sync::Arc;

fn encrypted(profile: EncryptionProfile) -> Package {
    let mut builder = Builder::new();
    let sheet = builder.add_worksheet("Protected").unwrap();
    builder.write_string(sheet, 0, 0, "retained value").unwrap();
    match profile {
        EncryptionProfile::XorObfuscation => builder
            .set_xor_obfuscation_password("secret", WeakEncryptionPolicy::allow_xor_obfuscation())
            .unwrap(),
        _ => builder.set_password("secret", profile).unwrap(),
    }
    if profile != EncryptionProfile::XorObfuscation {
        builder.set_encryption_entropy(litchi_xls::EncryptionEntropy::generate().unwrap());
    }
    builder.build().unwrap()
}

#[test]
fn exact_encryption_profile_survives_build_open_and_clone() {
    let profiles = [
        EncryptionProfile::XorObfuscation,
        EncryptionProfile::OfficeBinaryRc4,
    ]
    .into_iter()
    .chain(
        (40..=128)
            .step_by(8)
            .map(|key_bits| EncryptionProfile::CryptoApiRc4 { key_bits }),
    );
    for profile in profiles {
        let package = encrypted(profile);
        assert_eq!(package.encryption(), Some(profile));
        let bytes = package.to_encrypted_bytes().unwrap();
        let reopened =
            Package::from_bytes_with_options(&bytes, OpenOptions::new().password(Some("secret")))
                .unwrap();
        let retained = reopened.clone();
        drop(reopened);
        assert_eq!(retained.encryption(), Some(profile));
        assert_eq!(retained.reader_ref().unwrap().encryption(), Some(profile));
        assert!(retained.is_encrypted());
        assert!(
            retained
                .workbook()
                .text()
                .unwrap()
                .contains("retained value")
        );
        assert_eq!(retained.source_bytes().unwrap(), bytes);
        assert!(matches!(retained.to_bytes(), Err(Error::EncryptionPolicy)));
    }
}

#[test]
fn plain_eager_and_source_owners_have_no_encryption_profile() {
    let package = Package::new().unwrap();
    assert_eq!(package.encryption(), None);
    let bytes = package.to_bytes().unwrap();
    let cfb = Arc::new(
        litchi_cfb::SharedOleFile::open(
            Arc::new(litchi_core::OwnedSource::new(bytes)),
            litchi_cfb::SharedOleFileLimits::default(),
        )
        .unwrap(),
    );
    let workbook = litchi_xls::raw::source_backed_workbook_from_shared_ole_file(
        cfb,
        litchi_xls::SourceBackedLimits::default(),
    )
    .unwrap();
    let source = Package::from_source(workbook).unwrap();
    assert_eq!(source.encryption(), None);
    assert!(!source.is_encrypted());
    assert_eq!(source.reader_ref().unwrap().encryption(), None);
}

#[cfg(any(unix, windows))]
#[test]
fn explicit_password_path_preserves_mode_and_typed_password_failures() {
    struct Scratch(std::path::PathBuf);
    impl Drop for Scratch {
        fn drop(&mut self) {
            let _ = std::fs::remove_dir_all(&self.0);
        }
    }
    let directory = Scratch(
        std::env::temp_dir().join(format!("litchi-xls-security-0201-{}", std::process::id())),
    );
    std::fs::create_dir(&directory.0).unwrap();
    let path = directory.0.join("protected.xls");
    let mode = EncryptionProfile::CryptoApiRc4 { key_bits: 72 };
    let bytes = encrypted(mode).to_encrypted_bytes().unwrap();
    std::fs::write(&path, &bytes).unwrap();
    assert!(matches!(
        Package::open(&path),
        Err(Error::Reader(litchi_xls::Error::PasswordRequired))
    ));
    assert!(matches!(
        Package::open_with_password(&path, "incorrect"),
        Err(Error::Reader(litchi_xls::Error::InvalidPassword))
    ));
    let package = Package::open_with_password(&path, "secret").unwrap();
    assert_eq!(package.encryption(), Some(mode));
    assert_eq!(package.source_bytes().unwrap(), bytes);
    assert_eq!(std::fs::read(&path).unwrap(), bytes);
}
