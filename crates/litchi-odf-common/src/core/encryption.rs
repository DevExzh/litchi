use super::manifest::{
    ManifestChecksumAlgorithm, ManifestEncryption, ManifestEncryptionAlgorithm,
    ManifestKeyDerivation, ManifestStartKeyGeneration,
};
use aes::cipher::{BlockModeDecrypt, BlockModeEncrypt, KeyIvInit, block_padding::NoPadding};
use aes::{Aes128, Aes192, Aes256};
use aes_gcm::aead::consts::U12;
use aes_gcm::aead::{Aead, KeyInit as AeadKeyInit};
use aes_gcm::{Aes128Gcm, Aes256Gcm, AesGcm, Nonce};
use argon2::{
    Algorithm as Argon2Algorithm, Argon2, Params as Argon2Params, Version as Argon2Version,
};
use blowfish::Blowfish;
use blowfish::cipher::{Block, BlockEncrypt, KeyInit};
use cbc::{Decryptor, Encryptor};
use flate2::read::DeflateDecoder;
use litchi_core::{Error, Result};
use pbkdf2::pbkdf2_hmac;
use rand::{TryRng, rngs::SysRng};
use sha1::{Digest as _, Sha1};
use sha2::Sha256;
use std::io::{Read, Write};
use std::num::NonZeroU32;
use subtle::ConstantTimeEq;
use zeroize::Zeroizing;

const MAX_PLAINTEXT_ENTRY_SIZE: u64 = 512 * 1024 * 1024;
const MAX_ENCRYPTED_ENTRY_SIZE: usize = 1024 * 1024 * 1024;

type Aes192Gcm = AesGcm<Aes192, U12>;

/// Cipher used to encrypt ODF package payload entries.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Cipher {
    Aes128Cbc,
    Aes192Cbc,
    Aes256Cbc,
    Aes128Gcm,
    Aes192Gcm,
    Aes256Gcm,
    /// Legacy Blowfish CFB8 with a key size in bytes from 4 through 56.
    BlowfishCfb8 {
        key_size: u16,
    },
}

impl Cipher {
    fn key_size(self) -> u16 {
        match self {
            Self::Aes128Cbc | Self::Aes128Gcm => 16,
            Self::Aes192Cbc | Self::Aes192Gcm => 24,
            Self::Aes256Cbc | Self::Aes256Gcm => 32,
            Self::BlowfishCfb8 { key_size } => key_size,
        }
    }
}

/// Digest used to turn the password into KDF input.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StartKey {
    Sha1,
    Sha256,
}

/// Password-based key derivation settings.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kdf {
    Pbkdf2 {
        iterations: NonZeroU32,
    },
    Argon2id {
        iterations: NonZeroU32,
        memory_kib: NonZeroU32,
        lanes: NonZeroU32,
    },
}

/// Validated encryption settings for ODF package authoring.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Profile {
    cipher: Cipher,
    start_key: StartKey,
    kdf: Kdf,
}

impl Profile {
    /// Build and validate a custom encryption profile.
    ///
    /// # Errors
    ///
    /// Returns an error when the cipher key size or KDF parameters exceed the
    /// supported ODF encryption limits.
    pub fn new(cipher: Cipher, start_key: StartKey, kdf: Kdf) -> Result<Self> {
        let profile = Self {
            cipher,
            start_key,
            kdf,
        };
        profile.validate()?;
        Ok(profile)
    }

    /// AES-256-CBC with SHA-256 and PBKDF2 for broad ODF compatibility.
    #[must_use]
    pub fn compatible() -> Self {
        Self {
            cipher: Cipher::Aes256Cbc,
            start_key: StartKey::Sha256,
            kdf: Kdf::Pbkdf2 {
                iterations: NonZeroU32::new(100_000).unwrap_or(NonZeroU32::MIN),
            },
        }
    }

    /// AES-256-GCM with SHA-256 and Argon2id for authenticated encryption.
    #[must_use]
    pub fn authenticated() -> Self {
        Self {
            cipher: Cipher::Aes256Gcm,
            start_key: StartKey::Sha256,
            kdf: Kdf::Argon2id {
                iterations: NonZeroU32::new(3).unwrap_or(NonZeroU32::MIN),
                memory_kib: NonZeroU32::new(65_536).unwrap_or(NonZeroU32::MIN),
                lanes: NonZeroU32::MIN,
            },
        }
    }

    #[must_use]
    pub fn cipher(self) -> Cipher {
        self.cipher
    }

    #[must_use]
    pub fn start_key(self) -> StartKey {
        self.start_key
    }

    #[must_use]
    pub fn kdf(self) -> Kdf {
        self.kdf
    }

    fn validate(self) -> Result<()> {
        let key_size = self.cipher.key_size();
        if !(4..=56).contains(&key_size) {
            return Err(Error::InvalidFormat(format!(
                "Blowfish key size {key_size} is outside the supported range of 4 through 56 bytes"
            )));
        }
        match self.kdf {
            Kdf::Pbkdf2 { iterations }
                if iterations.get() > super::manifest::MAX_PBKDF2_ITERATIONS =>
            {
                Err(Error::InvalidFormat(
                    "PBKDF2 iteration count exceeds the supported limit".to_string(),
                ))
            },
            Kdf::Argon2id {
                iterations,
                memory_kib,
                lanes,
            } if iterations.get() > super::manifest::MAX_ARGON2_ITERATIONS
                || memory_kib.get() > super::manifest::MAX_ARGON2_MEMORY_KIB
                || lanes.get() > super::manifest::MAX_ARGON2_LANES
                || memory_kib.get() < 8 * lanes.get()
                || u64::from(memory_kib.get()) * u64::from(iterations.get())
                    > super::manifest::MAX_ARGON2_WORK_KIB_PASSES =>
            {
                Err(Error::InvalidFormat(
                    "Argon2id parameters exceed supported resource limits".to_string(),
                ))
            },
            Kdf::Argon2id { .. } if !matches!(key_size, 16 | 24 | 32) => Err(Error::InvalidFormat(
                "Argon2id derived key size must be 16, 24, or 32 bytes".to_string(),
            )),
            Kdf::Pbkdf2 { .. } | Kdf::Argon2id { .. } => Ok(()),
        }
    }
}

pub(crate) fn encrypt_entry(
    plaintext: &[u8],
    password: &str,
    profile: Profile,
) -> Result<(Vec<u8>, ManifestEncryption)> {
    profile.validate()?;
    let plaintext_length = u64::try_from(plaintext.len()).map_err(|error| {
        Error::InvalidFormat(format!(
            "ODF plaintext entry size does not fit u64: {error}"
        ))
    })?;
    if plaintext_length > MAX_PLAINTEXT_ENTRY_SIZE {
        return Err(Error::InvalidFormat(
            "ODF plaintext entry exceeds the supported size limit".to_string(),
        ));
    }

    let mut encoder =
        flate2::write::DeflateEncoder::new(Vec::new(), flate2::Compression::default());
    encoder.write_all(plaintext)?;
    let compressed = encoder.finish()?;

    let mut salt = vec![0_u8; 16];
    fill_random(&mut salt, "salt")?;
    let start_key = match profile.start_key {
        StartKey::Sha1 => ManifestStartKeyGeneration::Sha1,
        StartKey::Sha256 => ManifestStartKeyGeneration::Sha256,
    };
    let key_derivation = match profile.kdf {
        Kdf::Pbkdf2 { iterations } => ManifestKeyDerivation::Pbkdf2 {
            salt,
            iterations,
            key_size: profile.cipher.key_size(),
        },
        Kdf::Argon2id {
            iterations,
            memory_kib,
            lanes,
        } => ManifestKeyDerivation::Argon2id {
            salt,
            iterations,
            memory_kib,
            lanes,
            key_size: Some(profile.cipher.key_size()),
        },
    };
    let start_key_bytes = Zeroizing::new(match start_key {
        ManifestStartKeyGeneration::Sha1 => Sha1::digest(password.as_bytes()).to_vec(),
        ManifestStartKeyGeneration::Sha256 => Sha256::digest(password.as_bytes()).to_vec(),
    });
    let key = Zeroizing::new(derive_key(
        &start_key_bytes,
        &key_derivation,
        &algorithm_without_iv(profile.cipher),
    )?);

    let checksum = (!matches!(
        profile.cipher,
        Cipher::Aes128Gcm | Cipher::Aes192Gcm | Cipher::Aes256Gcm
    ))
    .then(|| {
        let prefix = &compressed[..compressed.len().min(1024)];
        match profile.start_key {
            StartKey::Sha1 => super::manifest::ManifestChecksum {
                algorithm: ManifestChecksumAlgorithm::Sha1First1024,
                value: Sha1::digest(prefix).to_vec(),
            },
            StartKey::Sha256 => super::manifest::ManifestChecksum {
                algorithm: ManifestChecksumAlgorithm::Sha256First1024,
                value: Sha256::digest(prefix).to_vec(),
            },
        }
    });
    let (ciphertext, algorithm) = encrypt_compressed(&compressed, &key, profile.cipher)?;
    Ok((
        ciphertext,
        ManifestEncryption {
            checksum,
            algorithm,
            start_key,
            key_derivation,
        },
    ))
}

fn fill_random(bytes: &mut [u8], field: &str) -> Result<()> {
    SysRng.try_fill_bytes(bytes).map_err(|error| {
        Error::InvalidFormat(format!(
            "Operating-system randomness unavailable for ODF {field}: {error}"
        ))
    })
}

fn algorithm_without_iv(cipher: Cipher) -> ManifestEncryptionAlgorithm {
    match cipher {
        Cipher::Aes128Cbc => ManifestEncryptionAlgorithm::Aes128Cbc { iv: [0; 16] },
        Cipher::Aes192Cbc => ManifestEncryptionAlgorithm::Aes192Cbc { iv: [0; 16] },
        Cipher::Aes256Cbc => ManifestEncryptionAlgorithm::Aes256Cbc { iv: [0; 16] },
        Cipher::Aes128Gcm => ManifestEncryptionAlgorithm::Aes128Gcm { iv: [0; 12] },
        Cipher::Aes192Gcm => ManifestEncryptionAlgorithm::Aes192Gcm { iv: [0; 12] },
        Cipher::Aes256Gcm => ManifestEncryptionAlgorithm::Aes256Gcm { iv: [0; 12] },
        Cipher::BlowfishCfb8 { .. } => ManifestEncryptionAlgorithm::BlowfishCfb8 { iv: [0; 8] },
    }
}

fn encrypt_compressed(
    compressed: &[u8],
    key: &[u8],
    cipher: Cipher,
) -> Result<(Vec<u8>, ManifestEncryptionAlgorithm)> {
    macro_rules! encrypt_cbc {
        ($aes:ty, $variant:ident) => {{
            let mut iv = [0_u8; 16];
            fill_random(&mut iv, "initialisation vector")?;
            let padding = 16 - compressed.len() % 16;
            let mut padded = compressed.to_vec();
            padded.resize(padded.len() + padding, 0);
            let padding_byte = padded.last_mut().ok_or_else(|| {
                Error::InvalidFormat("AES-CBC padding unexpectedly has no final byte".to_string())
            })?;
            *padding_byte = u8::try_from(padding).map_err(|error| {
                Error::InvalidFormat(format!("AES-CBC padding length exceeds one byte: {error}"))
            })?;
            let padded_len = padded.len();
            let encrypted = Encryptor::<$aes>::new_from_slices(key, &iv)
                .map_err(|error| {
                    Error::InvalidFormat(format!("Invalid AES-CBC key or IV: {error}"))
                })?
                .encrypt_padded::<NoPadding>(&mut padded, padded_len)
                .map_err(|error| {
                    Error::InvalidFormat(format!("Unable to encrypt ODF AES-CBC entry: {error}"))
                })?
                .to_vec();
            (encrypted, ManifestEncryptionAlgorithm::$variant { iv })
        }};
    }
    macro_rules! encrypt_gcm {
        ($aes:ty, $variant:ident) => {{
            let mut iv = [0_u8; 12];
            fill_random(&mut iv, "initialisation vector")?;
            let cipher = <$aes>::new_from_slice(key)
                .map_err(|error| Error::InvalidFormat(format!("Invalid AES-GCM key: {error}")))?;
            let mut encrypted = iv.to_vec();
            encrypted.extend(
                cipher
                    .encrypt(&Nonce::from(iv), compressed)
                    .map_err(|error| {
                        Error::InvalidFormat(format!(
                            "Unable to encrypt ODF AES-GCM entry: {error}"
                        ))
                    })?,
            );
            (encrypted, ManifestEncryptionAlgorithm::$variant { iv })
        }};
    }

    Ok(match cipher {
        Cipher::Aes128Cbc => encrypt_cbc!(Aes128, Aes128Cbc),
        Cipher::Aes192Cbc => encrypt_cbc!(Aes192, Aes192Cbc),
        Cipher::Aes256Cbc => encrypt_cbc!(Aes256, Aes256Cbc),
        Cipher::Aes128Gcm => encrypt_gcm!(Aes128Gcm, Aes128Gcm),
        Cipher::Aes192Gcm => encrypt_gcm!(Aes192Gcm, Aes192Gcm),
        Cipher::Aes256Gcm => encrypt_gcm!(Aes256Gcm, Aes256Gcm),
        Cipher::BlowfishCfb8 { .. } => {
            let mut iv = [0_u8; 8];
            fill_random(&mut iv, "initialisation vector")?;
            (
                encrypt_blowfish_cfb8(compressed, key, iv)?,
                ManifestEncryptionAlgorithm::BlowfishCfb8 { iv },
            )
        },
    })
}

fn encrypt_blowfish_cfb8(plaintext: &[u8], key: &[u8], iv: [u8; 8]) -> Result<Vec<u8>> {
    let cipher: Blowfish = Blowfish::new_from_slice(key).map_err(encryption_failure_from)?;
    let mut feedback = iv;
    let mut ciphertext = Vec::with_capacity(plaintext.len());
    for &byte in plaintext {
        let mut block = Block::<Blowfish>::default();
        block.copy_from_slice(&feedback);
        cipher.encrypt_block(&mut block);
        let encrypted = byte ^ block[0];
        ciphertext.push(encrypted);
        feedback.copy_within(1.., 0);
        feedback[7] = encrypted;
    }
    Ok(ciphertext)
}

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
    let key_size = u16::try_from(key.len()).map_err(encryption_failure_from)?;
    if !descriptor.algorithm.accepts_key_size(key_size) {
        return Err(encryption_failure());
    }
    let compressed = match &descriptor.algorithm {
        ManifestEncryptionAlgorithm::Aes128Cbc { iv } => remove_w3c_padding(
            Decryptor::<Aes128>::new_from_slices(key.as_ref(), iv)
                .map_err(encryption_failure_from)?
                .decrypt_padded_vec::<NoPadding>(ciphertext)
                .map_err(encryption_failure_from)?,
        )?,
        ManifestEncryptionAlgorithm::Aes192Cbc { iv } => remove_w3c_padding(
            Decryptor::<Aes192>::new_from_slices(key.as_ref(), iv)
                .map_err(encryption_failure_from)?
                .decrypt_padded_vec::<NoPadding>(ciphertext)
                .map_err(encryption_failure_from)?,
        )?,
        ManifestEncryptionAlgorithm::Aes256Cbc { iv } => remove_w3c_padding(
            Decryptor::<Aes256>::new_from_slices(key.as_ref(), iv)
                .map_err(encryption_failure_from)?
                .decrypt_padded_vec::<NoPadding>(ciphertext)
                .map_err(encryption_failure_from)?,
        )?,
        ManifestEncryptionAlgorithm::Aes128Gcm { iv } => {
            let encrypted = gcm_encrypted_payload(ciphertext, iv)?;
            Aes128Gcm::new_from_slice(key.as_ref())
                .map_err(encryption_failure_from)?
                .decrypt(&Nonce::from(*iv), encrypted)
                .map_err(encryption_failure_from)?
        },
        ManifestEncryptionAlgorithm::Aes192Gcm { iv } => {
            let encrypted = gcm_encrypted_payload(ciphertext, iv)?;
            Aes192Gcm::new_from_slice(key.as_ref())
                .map_err(encryption_failure_from)?
                .decrypt(&Nonce::from(*iv), encrypted)
                .map_err(encryption_failure_from)?
        },
        ManifestEncryptionAlgorithm::Aes256Gcm { iv } => {
            let encrypted = gcm_encrypted_payload(ciphertext, iv)?;
            Aes256Gcm::new_from_slice(key.as_ref())
                .map_err(encryption_failure_from)?
                .decrypt(&Nonce::from(*iv), encrypted)
                .map_err(encryption_failure_from)?
        },
        ManifestEncryptionAlgorithm::BlowfishCfb8 { iv } => {
            decrypt_blowfish_cfb8(ciphertext, key.as_ref(), *iv)?
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

    let expected_size = usize::try_from(plaintext_size).map_err(|error| {
        Error::InvalidFormat(format!(
            "ODF entry size does not fit this platform: {error}"
        ))
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
        .map_err(encryption_failure_from)?;
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
            let derived_key_size = (*key_size)
                .or(algorithm.fixed_key_size())
                .ok_or_else(encryption_failure)?;
            if !matches!(derived_key_size, 16 | 24 | 32) {
                return Err(encryption_failure());
            }
            let params = Argon2Params::new(
                memory_kib.get(),
                iterations.get(),
                lanes.get(),
                Some(usize::from(derived_key_size)),
            )
            .map_err(encryption_failure_from)?;
            let argon2 = Argon2::new(Argon2Algorithm::Argon2id, Argon2Version::V0x13, params);
            let mut key = vec![0u8; usize::from(derived_key_size)];
            argon2
                .hash_password_into(start_key, salt, &mut key)
                .map_err(encryption_failure_from)?;
            Ok(key)
        },
    }
}

fn decrypt_blowfish_cfb8(ciphertext: &[u8], key: &[u8], iv: [u8; 8]) -> Result<Vec<u8>> {
    let cipher: Blowfish = Blowfish::new_from_slice(key).map_err(encryption_failure_from)?;
    let mut feedback = iv;
    let mut plaintext = Vec::new();
    plaintext
        .try_reserve_exact(ciphertext.len())
        .map_err(encryption_failure_from)?;
    for &ciphertext_byte in ciphertext {
        let mut block = Block::<Blowfish>::default();
        block.copy_from_slice(&feedback);
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
    if payload.len() <= PREFIX_SIZE + TAG_SIZE || !bool::from(payload[..PREFIX_SIZE].ct_eq(iv)) {
        return Err(encryption_failure());
    }
    Ok(&payload[PREFIX_SIZE..])
}

fn remove_w3c_padding(mut decrypted: Vec<u8>) -> Result<Vec<u8>> {
    let padding = usize::from(decrypted.last().copied().unwrap_or(0));
    if padding == 0 || padding > 16 || padding > decrypted.len() {
        return Err(encryption_failure());
    }
    decrypted.truncate(decrypted.len() - padding);
    Ok(decrypted)
}

fn encryption_failure() -> Error {
    Error::InvalidFormat("Incorrect ODF password or corrupted encrypted package entry".to_string())
}

fn encryption_failure_from(error: impl std::fmt::Display) -> Error {
    Error::InvalidFormat(format!(
        "Incorrect ODF password or corrupted encrypted package entry: {error}"
    ))
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "Test fixtures use infallible setup operations, keeping assertions focused on decryption behavior."
)]
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
        let padding_byte = u8::try_from(padding).unwrap_or(u8::MAX);
        if let Some(last_byte) = padded.last_mut() {
            *last_byte = padding_byte;
        }

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
                _ => continue,
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

            let prepared = crate::detect::prepared(bytes.clone())
                .expect("encrypted ODT framing should remain detectable");
            assert!(prepared.package().get_file("content.xml").is_err());
            let unlocked = OwnedPackage::from_bytes_with_password(bytes.clone(), password).unwrap();
            assert_eq!(unlocked.get_file("content.xml").unwrap(), plaintext);
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
            let mut block = Block::<Blowfish>::default();
            block.copy_from_slice(&feedback);
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
                    zip.write_all(b"application/vnd.oasis.opendocument.text")
                        .unwrap();
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
            let tampered =
                OwnedPackage::from_bytes_with_password(package(&payload), password).unwrap();
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
