//! RC4 authoring consumes caller-owned entropy, never ambient randomness.
#![allow(clippy::unwrap_used, reason = "acceptance assertions fail fast")]
use litchi_xls::{Builder, EncryptionEntropy, EncryptionProfile};

fn builder(profile: EncryptionProfile) -> Builder {
    let mut builder = Builder::new();
    builder.set_password("secret", profile).unwrap();
    builder
}

fn missing_entropy(builder: &mut Builder) {
    assert!(matches!(
        builder.build(),
        Err(litchi_xls::package::Error::Reader(
            litchi_xls::Error::EncryptionEntropyRequired
        ))
    ));
}

#[test]
fn missing_entropy_refuses_before_default_sheet_is_staged() {
    let mut builder = builder(EncryptionProfile::OfficeBinaryRc4);
    missing_entropy(&mut builder);
    assert_eq!(builder.worksheet_count(), 0);
    builder.clear_password();
    assert!(builder.build().unwrap().to_bytes().is_ok());
}

#[test]
fn fixed_inputs_are_deterministic_and_entropy_is_consumed() {
    for mode in [
        EncryptionProfile::OfficeBinaryRc4,
        EncryptionProfile::CryptoApiRc4 { key_bits: 40 },
        EncryptionProfile::CryptoApiRc4 { key_bits: 128 },
    ] {
        let mut first = builder(mode);
        first.set_encryption_entropy(EncryptionEntropy::from_bytes([7; 16], [9; 16]));
        let output = first.build().unwrap().to_encrypted_bytes().unwrap();
        missing_entropy(&mut first);
        first.write_number(0, 0, 0, 42.0).unwrap();
        missing_entropy(&mut first);
        first.set_encryption_entropy(EncryptionEntropy::generate().unwrap());
        let changed = first.build().unwrap().to_encrypted_bytes().unwrap();
        assert_ne!(changed, output);
        let mut second = builder(mode);
        // Identical plaintext and fixed inputs deliberately reproduce the same artifact.
        second.set_encryption_entropy(EncryptionEntropy::from_bytes([7; 16], [9; 16]));
        assert_eq!(
            second.build().unwrap().to_encrypted_bytes().unwrap(),
            output
        );
    }
}

#[test]
fn plain_and_xor_builds_leave_unused_entropy_for_the_next_rc4_attempt() {
    let mut builder = Builder::new();
    builder.set_encryption_entropy(EncryptionEntropy::from_bytes([3; 16], [4; 16]));
    assert!(builder.build().unwrap().to_bytes().is_ok());
    builder
        .set_xor_obfuscation_password(
            "secret",
            litchi_xls::WeakEncryptionPolicy::allow_xor_obfuscation(),
        )
        .unwrap();
    let first = builder.build().unwrap().to_encrypted_bytes().unwrap();
    assert_eq!(
        first,
        builder.build().unwrap().to_encrypted_bytes().unwrap()
    );
    builder
        .set_password("secret", EncryptionProfile::OfficeBinaryRc4)
        .unwrap();
    assert!(builder.build().unwrap().is_encrypted());
    missing_entropy(&mut builder);
}
