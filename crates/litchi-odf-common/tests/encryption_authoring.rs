use litchi_odf_common::core::{
    Manifest, OdfEncryptionCipher, OdfEncryptionKdf, OdfEncryptionProfile, OdfEncryptionStartKey,
    OwnedPackage, PackageWriter,
};
use std::num::NonZeroU32;

const MIME: &str = "application/vnd.oasis.opendocument.text";
const PASSWORD: &str = "correct horse battery staple";
const CONTENT: &[u8] = b"<office:document-content>encrypted payload</office:document-content>";

fn nz(value: u32) -> NonZeroU32 {
    NonZeroU32::new(value).unwrap()
}

fn profile(cipher: OdfEncryptionCipher, argon2: bool) -> OdfEncryptionProfile {
    let kdf = if argon2 {
        OdfEncryptionKdf::Argon2id {
            iterations: nz(1),
            memory_kib: nz(8),
            lanes: nz(1),
        }
    } else {
        OdfEncryptionKdf::Pbkdf2 { iterations: nz(1) }
    };
    OdfEncryptionProfile::new(cipher, OdfEncryptionStartKey::Sha256, kdf).unwrap()
}

fn author(profile: OdfEncryptionProfile) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer.set_encryption(PASSWORD, profile).unwrap();
    writer.add_file("content.xml", CONTENT).unwrap();
    writer
        .add_file("META-INF/custom-metadata.xml", b"<metadata/>")
        .unwrap();
    writer.finish().unwrap()
}

#[test]
fn round_trips_every_cipher_with_both_kdfs() {
    let ciphers = [
        OdfEncryptionCipher::Aes128Cbc,
        OdfEncryptionCipher::Aes192Cbc,
        OdfEncryptionCipher::Aes256Cbc,
        OdfEncryptionCipher::Aes128Gcm,
        OdfEncryptionCipher::Aes192Gcm,
        OdfEncryptionCipher::Aes256Gcm,
        OdfEncryptionCipher::BlowfishCfb8 { key_size: 16 },
    ];
    for cipher in ciphers {
        for argon2 in [false, true] {
            let bytes = author(profile(cipher, argon2));
            let package = OwnedPackage::from_bytes_with_password(bytes, PASSWORD).unwrap();
            assert_eq!(package.get_file("content.xml").unwrap(), CONTENT);
            assert_eq!(
                package.get_file("META-INF/custom-metadata.xml").unwrap(),
                b"<metadata/>"
            );
        }
    }
}

#[test]
fn emits_typed_manifest_descriptors_and_fresh_randomness() {
    let first = author(OdfEncryptionProfile::compatible());
    let second = author(OdfEncryptionProfile::compatible());
    assert_ne!(first, second);

    let package = OwnedPackage::from_bytes(first).unwrap();
    let manifest_xml =
        String::from_utf8(package.get_file("META-INF/manifest.xml").unwrap()).unwrap();
    let manifest = Manifest::parse(&manifest_xml).unwrap();
    let content = manifest.get_entry("content.xml").unwrap();
    assert_eq!(content.size, Some(CONTENT.len() as u64));
    assert!(content.encryption.as_ref().unwrap().checksum.is_some());
    assert!(
        manifest
            .get_entry("META-INF/custom-metadata.xml")
            .unwrap()
            .encryption
            .is_none()
    );
    assert!(manifest_xml.contains("aes256-cbc"));
    assert!(manifest_xml.contains("PBKDF2"));
    assert!(manifest_xml.contains("sha256-1k"));
}

#[test]
fn rejects_wrong_password_and_ciphertext_tampering() {
    let bytes = author(OdfEncryptionProfile::authenticated());
    let wrong = OwnedPackage::from_bytes_with_password(bytes.clone(), "wrong").unwrap();
    assert!(wrong.get_file("content.xml").is_err());

    let archive = soapberry_zip::office::ArchiveReader::new(&bytes).unwrap();
    let ciphertext = archive.read("content.xml").unwrap();
    let offset = bytes
        .windows(ciphertext.len())
        .position(|window| window == ciphertext)
        .unwrap();
    let mut tampered = bytes;
    tampered[offset + ciphertext.len() / 2] ^= 1;
    let package = OwnedPackage::from_bytes_with_password(tampered, PASSWORD).unwrap();
    assert!(package.get_file("content.xml").is_err());
}

#[test]
fn rejects_invalid_profiles_and_late_configuration_without_mutation() {
    assert!(
        OdfEncryptionProfile::new(
            OdfEncryptionCipher::BlowfishCfb8 { key_size: 3 },
            OdfEncryptionStartKey::Sha1,
            OdfEncryptionKdf::Pbkdf2 { iterations: nz(1) },
        )
        .is_err()
    );
    assert!(
        OdfEncryptionProfile::new(
            OdfEncryptionCipher::Aes256Gcm,
            OdfEncryptionStartKey::Sha256,
            OdfEncryptionKdf::Pbkdf2 {
                iterations: nz(10_000_001)
            },
        )
        .is_err()
    );
    assert!(
        OdfEncryptionProfile::new(
            OdfEncryptionCipher::Aes256Gcm,
            OdfEncryptionStartKey::Sha256,
            OdfEncryptionKdf::Argon2id {
                iterations: nz(1),
                memory_kib: nz(262_145),
                lanes: nz(1),
            },
        )
        .is_err()
    );

    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer
        .set_encryption(PASSWORD, OdfEncryptionProfile::compatible())
        .unwrap();
    writer.add_file("content.xml", CONTENT).unwrap();
    assert!(
        writer
            .set_encryption("replacement", OdfEncryptionProfile::authenticated())
            .is_err()
    );
    assert!(writer.clear_encryption().is_err());
    let bytes = writer.finish().unwrap();
    let package = OwnedPackage::from_bytes_with_password(bytes, PASSWORD).unwrap();
    assert_eq!(package.get_file("content.xml").unwrap(), CONTENT);
}

#[test]
fn encryption_authoring_uses_no_unsafe_code() {
    assert!(!include_str!("../src/core/encryption.rs").contains("unsafe"));
    assert!(!include_str!("../src/core/writer.rs").contains("unsafe"));
}
