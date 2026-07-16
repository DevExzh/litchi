use super::manifest::{
    ManifestChecksumAlgorithm, ManifestEncryption, ManifestEncryptionAlgorithm,
    ManifestKeyDerivation, ManifestStartKeyGeneration,
};
use aes::{Aes128, Aes192, Aes256};
use aes::cipher::{BlockModeDecrypt, KeyIvInit, block_padding::NoPadding};
use aes_gcm::{Aes128Gcm, Aes256Gcm, AesGcm, Nonce};
use aes_gcm::aead::consts::U12;
use aes_gcm::aead::{Aead, KeyInit as AeadKeyInit};
use argon2::{Algorithm as Argon2Algorithm, Argon2, Params as Argon2Params, Version as Argon2Version};
use cbc::Decryptor;
use blowfish::Blowfish;
use blowfish::cipher::{Block, BlockEncrypt, KeyInit};
use flate2::read::DeflateDecoder;
use litchi_core::{Error, Result};
use pbkdf2::pbkdf2_hmac;
use sha1::{Digest as _, Sha1};
use sha2::Sha256;
use std::io::Read;
use subtle::ConstantTimeEq;
use zeroize::Zeroizing;

type Aes192Gcm = AesGcm<Aes192, U12>;

const MAX_PLAINTEXT_ENTRY_SIZE: u64 = 512 * 1024 * 1024;
const MAX_ENCRYPTED_ENTRY_SIZE: usize = 1024 * 1024 * 1024;

pub(crate) fn decrypt_entry(
    ciphertext: &[u8],
    password: &str,
    descriptor: &ManifestEncryption,
    plaintext_size: u64,
) -> Result<Vec<u8>> {
    if plaintext_size > MAX_PLAINTEXT_ENTRY_SIZE {
        return Err(Error::InvalidFormat(format!(
            "Encrypted ODF entry declares {plaintext_size} plaintext bytes, exceeding the limit"
        )));
    }
    if ciphertext.is_empty() || ciphertext.len() > MAX_ENCRYPTED_ENTRY_SIZE {
        return Err(encryption_failure());
    }
    if matches!(
        &descriptor.algorithm,
        ManifestEncryptionAlgorithm::Aes128Cbc { .. }
            | ManifestEncryptionAlgorithm::Aes192Cbc { .. }
            | ManifestEncryptionAlgorithm::Aes256Cbc { .. }
    ) && !ciphertext.len().is_multiple_of(16)
    {
        return Err(encryption_failure());
    }

    let start_key = Zeroizing::new(match descriptor.start_key {
        ManifestStartKeyGeneration::Sha1 => Sha1::digest(password.as_bytes()).to_vec(),
        ManifestStartKeyGeneration::Sha256 => Sha256::digest(password.as_bytes()).to_vec(),
    });
    let key = Zeroizing::new(derive_key(
        &start_key,
        &descriptor.key_derivation,
        &descriptor.algorithm,
    )?);
    let key_size = u16::try_from(key.len()).map_err(|_| encryption_failure())?;
    if !descriptor.algorithm.accepts_key_size(key_size) {
        return Err(encryption_failure());
    }
    let compressed = match &descriptor.algorithm {
        ManifestEncryptionAlgorithm::Aes128Cbc { iv } =>
            remove_w3c_padding(Decryptor::<Aes128>::new_from_slices(key.as_ref(), iv)
                .map_err(|_| encryption_failure())?
                .decrypt_padded_vec::<NoPadding>(ciphertext)
                .map_err(|_| encryption_failure())?)?,
        ManifestEncryptionAlgorithm::Aes192Cbc { iv } =>
            remove_w3c_padding(Decryptor::<Aes192>::new_from_slices(key.as_ref(), iv)
                .map_err(|_| encryption_failure())?
                .decrypt_padded_vec::<NoPadding>(ciphertext)
                .map_err(|_| encryption_failure())?)?,
        ManifestEncryptionAlgorithm::Aes256Cbc { iv } =>
            remove_w3c_padding(Decryptor::<Aes256>::new_from_slices(key.as_ref(), iv)
                .map_err(|_| encryption_failure())?
                .decrypt_padded_vec::<NoPadding>(ciphertext)
                .map_err(|_| encryption_failure())?)?,
        ManifestEncryptionAlgorithm::Aes128Gcm { iv } => {
            let encrypted = gcm_encrypted_payload(ciphertext, iv)?;
            Aes128Gcm::new_from_slice(key.as_ref())
                .map_err(|_| encryption_failure())?
                .decrypt(&Nonce::from(*iv), encrypted)
                .map_err(|_| encryption_failure())?
        },
        ManifestEncryptionAlgorithm::Aes192Gcm { iv } => {
            let encrypted = gcm_encrypted_payload(ciphertext, iv)?;
            Aes192Gcm::new_from_slice(key.as_ref())
                .map_err(|_| encryption_failure())?
                .decrypt(&Nonce::from(*iv), encrypted)
                .map_err(|_| encryption_failure())?
        },
        ManifestEncryptionAlgorithm::Aes256Gcm { iv } => {
            let encrypted = gcm_encrypted_payload(ciphertext, iv)?;
            Aes256Gcm::new_from_slice(key.as_ref())
                .map_err(|_| encryption_failure())?
                .decrypt(&Nonce::from(*iv), encrypted)
                .map_err(|_| encryption_failure())?
        },
        ManifestEncryptionAlgorithm::BlowfishCfb8 { iv } => {
            decrypt_blowfish_cfb8(ciphertext, key.as_ref(), iv)?
        },
    };

    if let Some(checksum) = &descriptor.checksum {
        let checksum_input = &compressed[..compressed.len().min(1024)];
        let actual = match checksum.algorithm {
            ManifestChecksumAlgorithm::Sha1First1024 => Sha1::digest(checksum_input).to_vec(),
            ManifestChecksumAlgorithm::Sha256First1024 => Sha256::digest(checksum_input).to_vec(),
        };
        if actual.len() != checksum.value.len() || !bool::from(actual.ct_eq(&checksum.value)) {
            return Err(encryption_failure());
        }
    } else if !descriptor.algorithm.is_aead() {
        return Err(encryption_failure());
    }

    let expected_size = usize::try_from(plaintext_size).map_err(|_| {
        Error::InvalidFormat("ODF entry size does not fit this platform".to_string())
    })?;
    let mut plaintext = Vec::new();
    plaintext
        .try_reserve_exact(expected_size)
        .map_err(|error| {
            Error::InvalidFormat(format!("Could not allocate decrypted ODF entry: {error}"))
        })?;
    DeflateDecoder::new(compressed.as_slice())
        .take(plaintext_size.saturating_add(1))
        .read_to_end(&mut plaintext)
        .map_err(|_| encryption_failure())?;
    if plaintext.len() != expected_size {
        return Err(encryption_failure());
    }
    Ok(plaintext)
}

fn derive_key(
    start_key: &[u8],
    derivation: &ManifestKeyDerivation,
    algorithm: &ManifestEncryptionAlgorithm,
) -> Result<Vec<u8>> {
    match derivation {
        ManifestKeyDerivation::Pbkdf2 {
            salt,
            iterations,
            key_size,
        } => {
            if !(4..=56).contains(key_size) {
                return Err(encryption_failure());
            }
            let mut key = vec![0u8; usize::from(*key_size)];
            pbkdf2_hmac::<Sha1>(start_key, salt, iterations.get(), &mut key);
            Ok(key)
        },
        ManifestKeyDerivation::Argon2id {
            salt,
            iterations,
            memory_kib,
            lanes,
            key_size,
        } => {
            let key_size = (*key_size)
                .or(algorithm.fixed_key_size())
                .ok_or_else(encryption_failure)?;
            if !matches!(key_size, 16 | 24 | 32) {
                return Err(encryption_failure());
            }
            let params = Argon2Params::new(
                memory_kib.get(),
                iterations.get(),
                lanes.get(),
                Some(usize::from(key_size)),
            )
            .map_err(|_| encryption_failure())?;
            let argon2 = Argon2::new(Argon2Algorithm::Argon2id, Argon2Version::V0x13, params);
            let mut key = vec![0u8; usize::from(key_size)];
            argon2
                .hash_password_into(start_key, salt, &mut key)
                .map_err(|_| encryption_failure())?;
            Ok(key)
        },
    }
}

fn decrypt_blowfish_cfb8(ciphertext: &[u8], key: &[u8], iv: &[u8; 8]) -> Result<Vec<u8>> {
    let cipher: Blowfish = Blowfish::new_from_slice(key).map_err(|_| encryption_failure())?;
    let mut feedback = *iv;
    let mut plaintext = Vec::new();
    plaintext
        .try_reserve_exact(ciphertext.len())
        .map_err(|_| encryption_failure())?;
    for &ciphertext_byte in ciphertext {
        let mut block = Block::<Blowfish>::clone_from_slice(&feedback);
        cipher.encrypt_block(&mut block);
        plaintext.push(ciphertext_byte ^ block[0]);
        feedback.copy_within(1.., 0);
        feedback[7] = ciphertext_byte;
    }
    Ok(plaintext)
}

fn gcm_encrypted_payload<'a>(payload: &'a [u8], iv: &[u8; 12]) -> Result<&'a [u8]> {
    const PREFIX_SIZE: usize = 12;
    const TAG_SIZE: usize = 16;
    if payload.len() <= PREFIX_SIZE + TAG_SIZE
        || !bool::from(payload[..PREFIX_SIZE].ct_eq(iv))
    {
        return Err(encryption_failure());
    }
    Ok(&payload[PREFIX_SIZE..])
}

fn remove_w3c_padding(mut decrypted: Vec<u8>) -> Result<Vec<u8>> {
    let padding = decrypted.last().copied().unwrap_or(0) as usize;
    if padding == 0 || padding > 16 || padding > decrypted.len() {
        return Err(encryption_failure());
    }
    decrypted.truncate(decrypted.len() - padding);
    Ok(decrypted)
}

fn encryption_failure() -> Error {
    Error::InvalidFormat("Incorrect ODF password or corrupted encrypted package entry".to_string())
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::core::package::OwnedPackage;
    use aes::cipher::BlockModeEncrypt;
    use base64::Engine as _;
    use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
    use cbc::Encryptor;
    use flate2::Compression;
    use flate2::write::DeflateEncoder;
    use std::io::{Cursor, Write};

    #[test]
    fn w3c_padding_accepts_random_filler_bytes() {
        let mut padded = vec![1, 2, 3, 0xaa, 0xbb, 3];
        padded.resize(16, 0);
        padded[15] = 10;
        assert_eq!(
            remove_w3c_padding(padded).unwrap(),
            [1, 2, 3, 0xaa, 0xbb, 3]
        );
    }

    #[test]
    fn decrypts_all_aes_cbc_odf_entries_and_rejects_wrong_password() {
        let password = "pässword";
        let plaintext = br#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text/></office:body></office:document-content>"#;
        let salt = [0x23u8; 16];
        let iv = [0x47u8; 16];
        let iterations = 1_000u32;

        let mut deflater = DeflateEncoder::new(Vec::new(), Compression::default());
        deflater.write_all(plaintext).unwrap();
        let compressed = deflater.finish().unwrap();
        let checksum = Sha256::digest(&compressed[..compressed.len().min(1024)]).to_vec();
        let start_key = Sha256::digest(password.as_bytes());
        let padding = 16 - (compressed.len() % 16);
        let mut padded = compressed.clone();
        padded.resize(padded.len() + padding, 0x5a);
        *padded.last_mut().unwrap() = padding as u8;

        for (key_size, algorithm) in [
            (16, "http://www.w3.org/2001/04/xmlenc#aes128-cbc"),
            (24, "http://www.w3.org/2001/04/xmlenc#aes192-cbc"),
            (32, "http://www.w3.org/2001/04/xmlenc#aes256-cbc"),
        ] {
            let mut key = vec![0u8; key_size];
            pbkdf2_hmac::<Sha1>(&start_key, &salt, iterations, &mut key);
            let ciphertext = match key_size {
                16 => Encryptor::<Aes128>::new_from_slices(&key, &iv)
                    .unwrap()
                    .encrypt_padded_vec::<NoPadding>(&padded),
                24 => Encryptor::<Aes192>::new_from_slices(&key, &iv)
                    .unwrap()
                    .encrypt_padded_vec::<NoPadding>(&padded),
                32 => Encryptor::<Aes256>::new_from_slices(&key, &iv)
                    .unwrap()
                    .encrypt_padded_vec::<NoPadding>(&padded),
                _ => unreachable!(),
            };

            let manifest = format!(
                r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml" m:size="{}"><m:encryption-data m:checksum-type="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0#sha256-1k" m:checksum="{}"><m:algorithm m:algorithm-name="{}" m:initialisation-vector="{}"/><m:start-key-generation m:start-key-generation-name="http://www.w3.org/2000/09/xmldsig#sha256" m:key-size="32"/><m:key-derivation m:key-derivation-name="PBKDF2" m:salt="{}" m:iteration-count="{}" m:key-size="{}"/></m:encryption-data></m:file-entry></m:manifest>"#,
                plaintext.len(),
                BASE64_STANDARD.encode(&checksum),
                algorithm,
                BASE64_STANDARD.encode(iv),
                BASE64_STANDARD.encode(salt),
                iterations,
                key_size,
            );

            let mut bytes = Vec::new();
            {
                let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
                let options = zip::write::SimpleFileOptions::default()
                    .compression_method(zip::CompressionMethod::Stored);
                zip.start_file("mimetype", options).unwrap();
                zip.write_all(b"application/vnd.oasis.opendocument.text")
                    .unwrap();
                zip.start_file("META-INF/manifest.xml", options).unwrap();
                zip.write_all(manifest.as_bytes()).unwrap();
                zip.start_file("content.xml", options).unwrap();
                zip.write_all(&ciphertext).unwrap();
                zip.finish().unwrap();
            }

            let unlocked = OwnedPackage::from_bytes_with_password(bytes.clone(), password).unwrap();
            assert_eq!(unlocked.get_file("content.xml").unwrap(), plaintext);
            assert!(crate::Document::from_bytes_with_password(bytes.clone(), password).is_ok());
            let locked = OwnedPackage::from_bytes(bytes.clone()).unwrap();
            assert!(locked.get_file("content.xml").is_err());
            let wrong = OwnedPackage::from_bytes_with_password(bytes, "wrong").unwrap();
            assert!(wrong.get_file("content.xml").is_err());
        }
    }

    #[test]
    fn decrypts_blowfish_cfb8_with_default_key_size() {
        let password = "legacy";
        let plaintext = b"legacy encrypted OpenDocument content";
        let salt = [0x31u8; 16];
        let iv = [0x72u8; 8];
        let iterations = 1_000u32;

        let mut deflater = DeflateEncoder::new(Vec::new(), Compression::default());
        deflater.write_all(plaintext).unwrap();
        let compressed = deflater.finish().unwrap();
        let checksum = Sha1::digest(&compressed[..compressed.len().min(1024)]).to_vec();
        let start_key = Sha1::digest(password.as_bytes());
        let mut key = [0u8; 16];
        pbkdf2_hmac::<Sha1>(&start_key, &salt, iterations, &mut key);
        let cipher: Blowfish = Blowfish::new_from_slice(&key).unwrap();
        let mut feedback = iv;
        let mut ciphertext = Vec::with_capacity(compressed.len());
        for &plaintext_byte in &compressed {
            let mut block = Block::<Blowfish>::clone_from_slice(&feedback);
            cipher.encrypt_block(&mut block);
            let ciphertext_byte = plaintext_byte ^ block[0];
            ciphertext.push(ciphertext_byte);
            feedback.copy_within(1.., 0);
            feedback[7] = ciphertext_byte;
        }

        let manifest = format!(
            r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml" m:size="{}"><m:encryption-data m:checksum-type="SHA1/1K" m:checksum="{}"><m:algorithm m:algorithm-name="Blowfish CFB" m:initialisation-vector="{}"/><m:start-key-generation m:start-key-generation-name="SHA1" m:key-size="20"/><m:key-derivation m:key-derivation-name="PBKDF2" m:salt="{}" m:iteration-count="{}"/></m:encryption-data></m:file-entry></m:manifest>"#,
            plaintext.len(),
            BASE64_STANDARD.encode(checksum),
            BASE64_STANDARD.encode(iv),
            BASE64_STANDARD.encode(salt),
            iterations,
        );
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(manifest.as_bytes()).unwrap();
            zip.start_file("content.xml", options).unwrap();
            zip.write_all(&ciphertext).unwrap();
            zip.finish().unwrap();
        }

        let unlocked = OwnedPackage::from_bytes_with_password(bytes.clone(), password).unwrap();
        assert_eq!(unlocked.get_file("content.xml").unwrap(), plaintext);
        let wrong = OwnedPackage::from_bytes_with_password(bytes, "wrong").unwrap();
        assert!(wrong.get_file("content.xml").is_err());
    }

    #[test]
    fn decrypts_all_aes_gcm_profiles_and_rejects_nonce_tampering() {
        let password = "authenticated";
        let plaintext = b"authenticated OpenDocument content";
        let salt = [0x19u8; 16];
        let iv = [0x63u8; 12];
        let iterations = 1_000u32;
        let mut deflater = DeflateEncoder::new(Vec::new(), Compression::default());
        deflater.write_all(plaintext).unwrap();
        let compressed = deflater.finish().unwrap();
        let start_key = Sha256::digest(password.as_bytes());

        for (key_size, algorithm) in [
            (16, "http://www.w3.org/2009/xmlenc11#aes128-gcm"),
            (24, "http://www.w3.org/2009/xmlenc11#aes192-gcm"),
            (32, "http://www.w3.org/2009/xmlenc11#aes256-gcm"),
        ] {
            let mut key = vec![0u8; key_size];
            pbkdf2_hmac::<Sha1>(&start_key, &salt, iterations, &mut key);
            let encrypted = match key_size {
                16 => Aes128Gcm::new_from_slice(&key)
                    .unwrap()
                    .encrypt(&Nonce::from(iv), compressed.as_slice())
                    .unwrap(),
                24 => Aes192Gcm::new_from_slice(&key)
                    .unwrap()
                    .encrypt(&Nonce::from(iv), compressed.as_slice())
                    .unwrap(),
                32 => Aes256Gcm::new_from_slice(&key)
                    .unwrap()
                    .encrypt(&Nonce::from(iv), compressed.as_slice())
                    .unwrap(),
                _ => unreachable!(),
            };
            let mut payload = iv.to_vec();
            payload.extend_from_slice(&encrypted);
            let manifest = format!(
                r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml" m:size="{}"><m:encryption-data><m:algorithm m:algorithm-name="{}" m:initialisation-vector="{}"/><m:start-key-generation m:start-key-generation-name="http://www.w3.org/2001/04/xmlenc#sha256" m:key-size="32"/><m:key-derivation m:key-derivation-name="PBKDF2" m:salt="{}" m:iteration-count="{}" m:key-size="{}"/></m:encryption-data></m:file-entry></m:manifest>"#,
                plaintext.len(),
                algorithm,
                BASE64_STANDARD.encode(iv),
                BASE64_STANDARD.encode(salt),
                iterations,
                key_size,
            );
            let package = |encrypted_content: &[u8]| {
                let mut bytes = Vec::new();
                {
                    let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
                    let options = zip::write::SimpleFileOptions::default()
                        .compression_method(zip::CompressionMethod::Stored);
                    zip.start_file("mimetype", options).unwrap();
                    zip.write_all(b"application/vnd.oasis.opendocument.text").unwrap();
                    zip.start_file("META-INF/manifest.xml", options).unwrap();
                    zip.write_all(manifest.as_bytes()).unwrap();
                    zip.start_file("content.xml", options).unwrap();
                    zip.write_all(encrypted_content).unwrap();
                    zip.finish().unwrap();
                }
                bytes
            };
            let bytes = package(&payload);
            let unlocked = OwnedPackage::from_bytes_with_password(bytes, password).unwrap();
            assert_eq!(unlocked.get_file("content.xml").unwrap(), plaintext);

            payload[0] ^= 1;
            let tampered = OwnedPackage::from_bytes_with_password(package(&payload), password).unwrap();
            assert!(tampered.get_file("content.xml").is_err());
            payload[0] ^= 1;
        }
    }

    #[test]
    fn decrypts_libreoffice_argon2id_aes256_gcm() {
        let password = "memory-hard";
        let plaintext = b"Argon2id protected OpenDocument content";
        let salt = [0x29u8; 16];
        let iv = [0x51u8; 12];
        let mut deflater = DeflateEncoder::new(Vec::new(), Compression::default());
        deflater.write_all(plaintext).unwrap();
        let compressed = deflater.finish().unwrap();
        let start_key = Sha256::digest(password.as_bytes());
        let params = Argon2Params::new(1024, 1, 1, Some(32)).unwrap();
        let argon2 = Argon2::new(Argon2Algorithm::Argon2id, Argon2Version::V0x13, params);
        let mut key = [0u8; 32];
        argon2
            .hash_password_into(&start_key, &salt, &mut key)
            .unwrap();
        let encrypted = Aes256Gcm::new_from_slice(&key)
            .unwrap()
            .encrypt(&Nonce::from(iv), compressed.as_slice())
            .unwrap();
        let mut payload = iv.to_vec();
        payload.extend_from_slice(&encrypted);
        let manifest = format!(
            r#"<m:manifest xmlns:m="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"><m:file-entry m:full-path="/" m:media-type="application/vnd.oasis.opendocument.text"/><m:file-entry m:full-path="content.xml" m:media-type="text/xml" m:size="{}"><m:encryption-data><m:algorithm m:algorithm-name="http://www.w3.org/2009/xmlenc11#aes256-gcm" m:initialisation-vector="{}"/><m:start-key-generation m:start-key-generation-name="http://www.w3.org/2001/04/xmlenc#sha256" m:key-size="32"/><m:key-derivation m:key-derivation-name="urn:org:documentfoundation:names:experimental:office:manifest:argon2id" loext:argon2-iterations="1" loext:argon2-memory="1024" loext:argon2-lanes="1" m:salt="{}"/></m:encryption-data></m:file-entry></m:manifest>"#,
            plaintext.len(),
            BASE64_STANDARD.encode(iv),
            BASE64_STANDARD.encode(salt),
        );
        let mut bytes = Vec::new();
        {
            let mut zip = zip::ZipWriter::new(Cursor::new(&mut bytes));
            let options = zip::write::SimpleFileOptions::default()
                .compression_method(zip::CompressionMethod::Stored);
            zip.start_file("mimetype", options).unwrap();
            zip.write_all(b"application/vnd.oasis.opendocument.text")
                .unwrap();
            zip.start_file("META-INF/manifest.xml", options).unwrap();
            zip.write_all(manifest.as_bytes()).unwrap();
            zip.start_file("content.xml", options).unwrap();
            zip.write_all(&payload).unwrap();
            zip.finish().unwrap();
        }
        let unlocked = OwnedPackage::from_bytes_with_password(bytes.clone(), password).unwrap();
        assert_eq!(unlocked.get_file("content.xml").unwrap(), plaintext);
        let wrong = OwnedPackage::from_bytes_with_password(bytes, "wrong").unwrap();
        assert!(wrong.get_file("content.xml").is_err());
    }
}
