//! `[MS-OFFCRYPTO]` Agile Encryption (AES-128/CBC/SHA-1 profile).

use std::fmt::Write as _;

use aes::Aes128;
use aes::cipher::{BlockModeDecrypt, BlockModeEncrypt, KeyIvInit, block_padding::NoPadding};
use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64;
use cbc::{Decryptor, Encryptor};
use hmac::{Hmac, Mac, digest::KeyInit};
use quick_xml::XmlVersion;
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesDecl, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use rand::TryRng;
use rand::rngs::SysRng;
use sha1::{Digest, Sha1};
use subtle::ConstantTimeEq;
use zeroize::{Zeroize, Zeroizing};

use super::{
    Error, Limits, Result, SPEC_MAX_SPIN_COUNT, container, declared_size, malformed, password_bytes,
};

const BLOCK: usize = 16;
const KEY_BYTES: usize = 16;
const HASH_BYTES: usize = 20;
const ENCRYPTED_HASH_BYTES: usize = 32;
const SPIN_COUNT: u32 = 100_000;
const SEGMENT: usize = 4_096;
const ENC_NS: &[u8] = b"http://schemas.microsoft.com/office/2006/encryption";
const PASSWORD_NS: &[u8] = b"http://schemas.microsoft.com/office/2006/keyEncryptor/password";

const VERIFIER_INPUT_BLOCK: [u8; 8] = [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
const HASHED_VERIFIER_BLOCK: [u8; 8] = [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
const CRYPTO_KEY_BLOCK: [u8; 8] = [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];
const INTEGRITY_KEY_BLOCK: [u8; 8] = [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];
const INTEGRITY_VALUE_BLOCK: [u8; 8] = [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];

type HmacSha1 = Hmac<Sha1>;
type AesCbcEnc = Encryptor<Aes128>;
type AesCbcDec = Decryptor<Aes128>;

struct Material {
    verifier_salt: [u8; BLOCK],
    verifier: [u8; BLOCK],
    key_salt: [u8; BLOCK],
    content_key: [u8; KEY_BYTES],
    // MS-OFFCRYPTO 2.3.4.14 says this follows saltSize (16), while the
    // published 3.11 Office vector stores a 32-byte encrypted form consistent
    // with a 20-byte SHA-1-sized salt. Office interoperability takes priority.
    integrity_salt: [u8; HASH_BYTES],
}

impl Zeroize for Material {
    fn zeroize(&mut self) {
        self.verifier_salt.zeroize();
        self.verifier.zeroize();
        self.key_salt.zeroize();
        self.content_key.zeroize();
        self.integrity_salt.zeroize();
    }
}

struct Info {
    spin_count: u32,
    key_salt: [u8; BLOCK],
    verifier_salt: [u8; BLOCK],
    encrypted_verifier: [u8; BLOCK],
    encrypted_verifier_hash: [u8; ENCRYPTED_HASH_BYTES],
    encrypted_key: [u8; KEY_BYTES],
    encrypted_hmac_key: [u8; ENCRYPTED_HASH_BYTES],
    encrypted_hmac_value: [u8; ENCRYPTED_HASH_BYTES],
}

#[derive(Default)]
struct Parsed {
    key_salt: Option<[u8; BLOCK]>,
    verifier_salt: Option<[u8; BLOCK]>,
    encrypted_verifier: Option<[u8; BLOCK]>,
    encrypted_verifier_hash: Option<[u8; ENCRYPTED_HASH_BYTES]>,
    encrypted_key: Option<[u8; KEY_BYTES]>,
    encrypted_hmac_key: Option<[u8; ENCRYPTED_HASH_BYTES]>,
    encrypted_hmac_value: Option<[u8; ENCRYPTED_HASH_BYTES]>,
    spin_count: Option<u32>,
}

#[derive(Clone, Copy)]
enum Direction {
    Encrypt,
    Decrypt,
}

pub(super) fn encrypt(package: Vec<u8>, password: &str, limits: &Limits) -> Result<Vec<u8>> {
    let mut material = Zeroizing::new(Material {
        verifier_salt: [0; BLOCK],
        verifier: [0; BLOCK],
        key_salt: [0; BLOCK],
        content_key: [0; KEY_BYTES],
        integrity_salt: [0; HASH_BYTES],
    });
    let mut rng = SysRng;
    fill_random(&mut rng, &mut material.verifier_salt, "Agile verifier salt")?;
    fill_random(&mut rng, &mut material.verifier, "Agile verifier")?;
    fill_random(&mut rng, &mut material.key_salt, "Agile key salt")?;
    fill_random(&mut rng, &mut material.content_key, "Agile content key")?;
    fill_random(
        &mut rng,
        &mut material.integrity_salt,
        "Agile integrity salt",
    )?;
    let (info, encrypted) = encrypt_parts(package, password, &material, limits)?;
    container::write(&info, encrypted, limits)
}

pub(super) fn decrypt(
    info: &[u8],
    encrypted: Vec<u8>,
    password: &str,
    limits: &Limits,
) -> Result<Vec<u8>> {
    Limits::bytes("EncryptionInfo", info.len(), limits.max_info_bytes)?;
    Limits::bytes(
        "EncryptedPackage",
        encrypted.len(),
        limits.max_encrypted_bytes,
    )?;
    let parsed = parse(info, limits)?;
    let password_hash = password_hash(password, &parsed.verifier_salt, parsed.spin_count, limits)?;

    let verifier = decrypt_value::<BLOCK>(
        &parsed.verifier_salt,
        &password_hash,
        VERIFIER_INPUT_BLOCK,
        &parsed.encrypted_verifier,
    )?;
    let verifier_hash = decrypt_value::<HASH_BYTES>(
        &parsed.verifier_salt,
        &password_hash,
        HASHED_VERIFIER_BLOCK,
        &parsed.encrypted_verifier_hash,
    )?;
    let expected_hash = Zeroizing::new(<[u8; HASH_BYTES]>::from(Sha1::digest(verifier.as_slice())));
    if !bool::from(verifier_hash.ct_eq(expected_hash.as_slice())) {
        return Err(Error::Password);
    }

    let content_key = decrypt_value::<KEY_BYTES>(
        &parsed.verifier_salt,
        &password_hash,
        CRYPTO_KEY_BLOCK,
        &parsed.encrypted_key,
    )?;
    verify_integrity(&parsed, &content_key, &encrypted)?;
    let clear_len = package_size(&encrypted, limits)?;
    decrypt_package(encrypted, clear_len, &content_key, &parsed.key_salt)
}

pub(super) fn validate_info(info: &[u8], limits: &Limits) -> Result<()> {
    parse(info, limits).map(|_| ())
}

fn fill_random(rng: &mut SysRng, bytes: &mut [u8], name: &'static str) -> Result<()> {
    rng.try_fill_bytes(bytes)
        .map_err(|error| Error::Random(format!("{name}: {error}")))
}

fn encrypt_parts(
    package: Vec<u8>,
    password: &str,
    material: &Material,
    limits: &Limits,
) -> Result<(Vec<u8>, Vec<u8>)> {
    check_spin(SPIN_COUNT, limits)?;
    let password_hash = password_hash(password, &material.verifier_salt, SPIN_COUNT, limits)?;
    let encrypted_verifier = encrypt_value::<BLOCK>(
        &material.verifier_salt,
        &password_hash,
        VERIFIER_INPUT_BLOCK,
        &material.verifier,
    )?;
    let mut verifier_digest = Sha1::new();
    verifier_digest.update(material.verifier.as_slice());
    let verifier_hash = Zeroizing::new(<[u8; HASH_BYTES]>::from(verifier_digest.finalize()));
    let encrypted_verifier_hash = encrypt_value::<ENCRYPTED_HASH_BYTES>(
        &material.verifier_salt,
        &password_hash,
        HASHED_VERIFIER_BLOCK,
        verifier_hash.as_slice(),
    )?;
    let encrypted_key = encrypt_value::<KEY_BYTES>(
        &material.verifier_salt,
        &password_hash,
        CRYPTO_KEY_BLOCK,
        &material.content_key,
    )?;

    let encrypted = encrypt_package(package, &material.content_key, &material.key_salt, limits)?;
    let integrity_value = hmac(&material.integrity_salt, &encrypted)?;
    let encrypted_hmac_key = encrypt_content::<ENCRYPTED_HASH_BYTES>(
        &material.content_key,
        &material.key_salt,
        INTEGRITY_KEY_BLOCK,
        &material.integrity_salt,
    )?;
    let encrypted_hmac_value = encrypt_content::<ENCRYPTED_HASH_BYTES>(
        &material.content_key,
        &material.key_salt,
        INTEGRITY_VALUE_BLOCK,
        integrity_value.as_slice(),
    )?;

    let info = Info {
        spin_count: SPIN_COUNT,
        key_salt: material.key_salt,
        verifier_salt: material.verifier_salt,
        encrypted_verifier,
        encrypted_verifier_hash,
        encrypted_key,
        encrypted_hmac_key,
        encrypted_hmac_value,
    };
    Ok((build_info(&info, limits)?, encrypted))
}

fn password_hash(
    password: &str,
    salt: &[u8; BLOCK],
    spin_count: u32,
    limits: &Limits,
) -> Result<Zeroizing<[u8; HASH_BYTES]>> {
    check_spin(spin_count, limits)?;
    let encoded = password_bytes(password, limits)?;
    let mut hasher = Sha1::new();
    hasher.update(salt);
    hasher.update(encoded.as_slice());
    let mut hash = Zeroizing::new(<[u8; HASH_BYTES]>::from(hasher.finalize()));
    for iterator in 0..spin_count {
        let mut spin = Sha1::new();
        spin.update(iterator.to_le_bytes());
        spin.update(hash.as_slice());
        hash = Zeroizing::new(<[u8; HASH_BYTES]>::from(spin.finalize()));
    }
    Ok(hash)
}

fn check_spin(spin_count: u32, limits: &Limits) -> Result<()> {
    if spin_count > SPEC_MAX_SPIN_COUNT {
        return Err(malformed("Agile spinCount exceeds the schema maximum"));
    }
    if spin_count > limits.max_spin_count {
        return Err(Error::Limit {
            resource: "Agile spin count",
            actual: u64::from(spin_count),
            maximum: u64::from(limits.max_spin_count),
        });
    }
    Ok(())
}

fn derive_key(password_hash: &[u8; HASH_BYTES], block_key: [u8; 8]) -> Zeroizing<[u8; KEY_BYTES]> {
    let mut sha = Sha1::new();
    sha.update(password_hash);
    sha.update(block_key);
    let digest = Zeroizing::new(<[u8; HASH_BYTES]>::from(sha.finalize()));
    let mut key = Zeroizing::new([0x36; KEY_BYTES]);
    key.copy_from_slice(&digest[..KEY_BYTES]);
    key
}

fn iv(salt: &[u8; BLOCK], block_key: Option<&[u8]>) -> [u8; BLOCK] {
    let Some(segment) = block_key else {
        return *salt;
    };
    let mut sha = Sha1::new();
    sha.update(salt);
    sha.update(segment);
    let digest = <[u8; HASH_BYTES]>::from(sha.finalize());
    let mut output = [0u8; BLOCK];
    output.copy_from_slice(&digest[..BLOCK]);
    output
}

fn encrypt_value<const N: usize>(
    salt: &[u8; BLOCK],
    password_hash: &[u8; HASH_BYTES],
    block_key: [u8; 8],
    input: &[u8],
) -> Result<[u8; N]> {
    if round_up(input.len(), BLOCK)? != N {
        return Err(malformed("Agile password value has an invalid padded size"));
    }
    let key = derive_key(password_hash, block_key);
    let iv = iv(salt, None);
    let mut output = Zeroizing::new([0u8; N]);
    let destination = output
        .get_mut(..input.len())
        .ok_or_else(|| malformed("Agile password value exceeds its destination"))?;
    destination.copy_from_slice(input);
    cbc_encrypt(&key, &iv, &mut output[..])?;
    Ok(*output)
}

fn decrypt_value<const N: usize>(
    salt: &[u8; BLOCK],
    password_hash: &[u8; HASH_BYTES],
    block_key: [u8; 8],
    encrypted: &[u8],
) -> Result<Zeroizing<[u8; N]>> {
    if encrypted.is_empty()
        || encrypted.len() > ENCRYPTED_HASH_BYTES
        || !encrypted.len().is_multiple_of(BLOCK)
        || encrypted.len() < N
    {
        return Err(malformed(
            "Agile encrypted password value has an invalid size",
        ));
    }
    let key = derive_key(password_hash, block_key);
    let iv = iv(salt, None);
    let mut buffer = Zeroizing::new([0u8; ENCRYPTED_HASH_BYTES]);
    let destination = buffer
        .get_mut(..encrypted.len())
        .ok_or_else(|| malformed("Agile encrypted password value is too large"))?;
    destination.copy_from_slice(encrypted);
    cbc_decrypt(&key, &iv, destination)?;
    let source = buffer
        .get(..N)
        .ok_or_else(|| malformed("Agile decrypted password value is truncated"))?;
    let mut output = Zeroizing::new([0u8; N]);
    output.copy_from_slice(source);
    Ok(output)
}

fn encrypt_content<const N: usize>(
    key: &[u8; KEY_BYTES],
    salt: &[u8; BLOCK],
    block_key: [u8; 8],
    input: &[u8],
) -> Result<[u8; N]> {
    if round_up(input.len(), BLOCK)? != N {
        return Err(malformed(
            "Agile integrity value has an invalid padded size",
        ));
    }
    let mut output = Zeroizing::new([0u8; N]);
    let destination = output
        .get_mut(..input.len())
        .ok_or_else(|| malformed("Agile integrity value exceeds its destination"))?;
    destination.copy_from_slice(input);
    cbc_encrypt(key, &iv(salt, Some(&block_key)), &mut output[..])?;
    Ok(*output)
}

fn decrypt_content<const N: usize>(
    key: &[u8; KEY_BYTES],
    salt: &[u8; BLOCK],
    block_key: [u8; 8],
    encrypted: &[u8; ENCRYPTED_HASH_BYTES],
) -> Result<Zeroizing<[u8; N]>> {
    if N > ENCRYPTED_HASH_BYTES {
        return Err(malformed("Agile integrity output size is unsupported"));
    }
    let mut buffer = Zeroizing::new(*encrypted);
    cbc_decrypt(key, &iv(salt, Some(&block_key)), &mut buffer[..])?;
    let source = buffer
        .get(..N)
        .ok_or_else(|| malformed("Agile decrypted integrity value is truncated"))?;
    let mut output = Zeroizing::new([0u8; N]);
    output.copy_from_slice(source);
    Ok(output)
}

fn hmac(key: &[u8], bytes: &[u8]) -> Result<Zeroizing<[u8; HASH_BYTES]>> {
    let mut mac = <HmacSha1 as KeyInit>::new_from_slice(key)
        .map_err(|_err| malformed("HMAC-SHA1 key length invariant was violated"))?;
    mac.update(bytes);
    Ok(Zeroizing::new(<[u8; HASH_BYTES]>::from(
        mac.finalize().into_bytes(),
    )))
}

fn verify_integrity(info: &Info, content_key: &[u8; KEY_BYTES], encrypted: &[u8]) -> Result<()> {
    let integrity_salt = decrypt_content::<HASH_BYTES>(
        content_key,
        &info.key_salt,
        INTEGRITY_KEY_BLOCK,
        &info.encrypted_hmac_key,
    )?;
    let stored = decrypt_content::<HASH_BYTES>(
        content_key,
        &info.key_salt,
        INTEGRITY_VALUE_BLOCK,
        &info.encrypted_hmac_value,
    )?;
    let expected = hmac(integrity_salt.as_slice(), encrypted)?;
    if !bool::from(stored.ct_eq(expected.as_slice())) {
        return Err(Error::Integrity);
    }
    Ok(())
}

fn encrypt_package(
    mut package: Vec<u8>,
    content_key: &[u8; KEY_BYTES],
    key_salt: &[u8; BLOCK],
    limits: &Limits,
) -> Result<Vec<u8>> {
    let clear_len = package.len();
    let cipher_len = round_up(clear_len, BLOCK)?;
    let total = cipher_len
        .checked_add(8)
        .ok_or_else(|| malformed("Agile EncryptedPackage size overflows usize"))?;
    Limits::bytes("EncryptedPackage", total, limits.max_encrypted_bytes)?;
    package
        .try_reserve_exact(total.saturating_sub(package.len()))
        .map_err(|_err| Error::Allocation("Agile EncryptedPackage"))?;
    package.resize(total, 0);
    package.copy_within(0..clear_len, 8);
    let prefix = package
        .get_mut(..8)
        .ok_or_else(|| malformed("Agile EncryptedPackage prefix is unavailable"))?;
    prefix.copy_from_slice(
        &u64::try_from(clear_len)
            .map_err(|_err| malformed("plaintext size does not fit u64"))?
            .to_le_bytes(),
    );

    crypt_segments(
        &mut package,
        clear_len,
        content_key,
        key_salt,
        Direction::Encrypt,
    )?;
    Ok(package)
}

fn package_size(encrypted: &[u8], limits: &Limits) -> Result<usize> {
    if encrypted.len() < 8 + BLOCK {
        return Err(malformed("Agile EncryptedPackage is too short"));
    }
    let prefix: [u8; 8] = encrypted
        .get(..8)
        .ok_or_else(|| malformed("Agile EncryptedPackage has no StreamSize"))?
        .try_into()
        .map_err(|_err| malformed("Agile StreamSize has the wrong length"))?;
    let clear_len = declared_size(u64::from_le_bytes(prefix), limits)?;
    if clear_len == 0 {
        return Err(malformed(
            "Agile EncryptedPackage declares an empty package",
        ));
    }
    let expected = round_up(clear_len, BLOCK)?
        .checked_add(8)
        .ok_or_else(|| malformed("Agile EncryptedPackage length overflows usize"))?;
    if encrypted.len() != expected {
        return Err(malformed(
            "Agile EncryptedPackage length disagrees with StreamSize",
        ));
    }
    Ok(clear_len)
}

fn decrypt_package(
    mut encrypted: Vec<u8>,
    clear_len: usize,
    content_key: &[u8; KEY_BYTES],
    key_salt: &[u8; BLOCK],
) -> Result<Vec<u8>> {
    crypt_segments(
        &mut encrypted,
        clear_len,
        content_key,
        key_salt,
        Direction::Decrypt,
    )?;
    let source_end = clear_len
        .checked_add(8)
        .ok_or_else(|| malformed("Agile plaintext range overflows usize"))?;
    if encrypted.get(8..source_end).is_none() {
        return Err(malformed("Agile decrypted package is truncated"));
    }
    encrypted.copy_within(8..source_end, 0);
    encrypted.truncate(clear_len);
    Ok(encrypted)
}

fn crypt_segments(
    bytes: &mut [u8],
    clear_len: usize,
    content_key: &[u8; KEY_BYTES],
    key_salt: &[u8; BLOCK],
    direction: Direction,
) -> Result<()> {
    let segments = clear_len
        .checked_add(SEGMENT - 1)
        .ok_or_else(|| malformed("Agile segment count overflows usize"))?
        / SEGMENT;
    for index in 0..segments {
        let clear_start = index
            .checked_mul(SEGMENT)
            .ok_or_else(|| malformed("Agile segment offset overflows usize"))?;
        let clear_segment = (clear_len - clear_start).min(SEGMENT);
        let cipher_segment = round_up(clear_segment, BLOCK)?;
        let start = clear_start
            .checked_add(8)
            .ok_or_else(|| malformed("Agile segment start overflows usize"))?;
        let end = start
            .checked_add(cipher_segment)
            .ok_or_else(|| malformed("Agile segment end overflows usize"))?;
        let segment = bytes
            .get_mut(start..end)
            .ok_or_else(|| malformed("Agile encrypted segment is truncated"))?;
        let block = u32::try_from(index)
            .map_err(|_err| malformed("Agile segment index exceeds u32"))?
            .to_le_bytes();
        let iv = iv(key_salt, Some(&block));
        match direction {
            Direction::Encrypt => cbc_encrypt(content_key, &iv, segment)?,
            Direction::Decrypt => cbc_decrypt(content_key, &iv, segment)?,
        }
    }
    Ok(())
}

fn cbc_encrypt(key: &[u8; KEY_BYTES], iv: &[u8; BLOCK], bytes: &mut [u8]) -> Result<()> {
    let message_len = bytes.len();
    AesCbcEnc::new_from_slices(key, iv)
        .map_err(|_err| malformed("AES-128-CBC key or IV length invariant was violated"))?
        .encrypt_padded::<NoPadding>(bytes, message_len)
        .map(|_| ())
        .map_err(|_err| malformed("AES-128-CBC input is not block aligned"))
}

fn cbc_decrypt(key: &[u8; KEY_BYTES], iv: &[u8; BLOCK], bytes: &mut [u8]) -> Result<()> {
    AesCbcDec::new_from_slices(key, iv)
        .map_err(|_err| malformed("AES-128-CBC key or IV length invariant was violated"))?
        .decrypt_padded::<NoPadding>(bytes)
        .map(|_| ())
        .map_err(|_err| malformed("AES-128-CBC input is not block aligned"))
}

fn round_up(value: usize, multiple: usize) -> Result<usize> {
    value
        .checked_add(multiple - 1)
        .map(|padded| padded / multiple * multiple)
        .ok_or_else(|| malformed("Agile block length overflows usize"))
}

fn build_info(info: &Info, limits: &Limits) -> Result<Vec<u8>> {
    let key_salt = BASE64.encode(info.key_salt);
    let verifier_salt = BASE64.encode(info.verifier_salt);
    let encrypted_verifier = BASE64.encode(info.encrypted_verifier);
    let encrypted_verifier_hash = BASE64.encode(info.encrypted_verifier_hash);
    let encrypted_key = BASE64.encode(info.encrypted_key);
    let encrypted_hmac_key = BASE64.encode(info.encrypted_hmac_key);
    let encrypted_hmac_value = BASE64.encode(info.encrypted_hmac_value);

    let mut xml = String::new();
    xml.try_reserve(1_024)
        .map_err(|_err| Error::Allocation("Agile EncryptionInfo XML"))?;
    write!(
        xml,
        concat!(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#,
            r#"<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">"#,
            r#"<keyData saltSize="16" blockSize="16" keyBits="128" hashSize="20" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA-1" saltValue="{}"/>"#,
            r#"<dataIntegrity encryptedHmacKey="{}" encryptedHmacValue="{}"/>"#,
            r#"<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">"#,
            r#"<p:encryptedKey spinCount="{}" saltSize="16" blockSize="16" keyBits="128" hashSize="20" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA-1" saltValue="{}" encryptedVerifierHashInput="{}" encryptedVerifierHashValue="{}" encryptedKeyValue="{}"/>"#,
            r#"</keyEncryptor></keyEncryptors></encryption>"#,
        ),
        key_salt,
        encrypted_hmac_key,
        encrypted_hmac_value,
        info.spin_count,
        verifier_salt,
        encrypted_verifier,
        encrypted_verifier_hash,
        encrypted_key,
    )
    .map_err(|_err| Error::Allocation("Agile EncryptionInfo XML"))?;
    Limits::bytes("Agile XML", xml.len(), limits.max_xml_bytes)?;
    let total = xml
        .len()
        .checked_add(8)
        .ok_or_else(|| malformed("Agile EncryptionInfo length overflows usize"))?;
    Limits::bytes("EncryptionInfo", total, limits.max_info_bytes)?;
    let mut output = Vec::new();
    output
        .try_reserve_exact(total)
        .map_err(|_err| Error::Allocation("Agile EncryptionInfo"))?;
    output.extend_from_slice(&4u16.to_le_bytes());
    output.extend_from_slice(&4u16.to_le_bytes());
    output.extend_from_slice(&0x40u32.to_le_bytes());
    output.extend_from_slice(xml.as_bytes());
    Ok(output)
}

fn parse(bytes: &[u8], limits: &Limits) -> Result<Info> {
    Limits::bytes("EncryptionInfo", bytes.len(), limits.max_info_bytes)?;
    let prefix = bytes
        .get(..8)
        .ok_or_else(|| malformed("Agile EncryptionInfo is shorter than its header"))?;
    if prefix != [4, 0, 4, 0, 0x40, 0, 0, 0] {
        return Err(Error::Unsupported(
            "Agile EncryptionInfo must use version 4.4 and reserved value 0x40".into(),
        ));
    }
    let xml = bytes
        .get(8..)
        .ok_or_else(|| malformed("Agile EncryptionInfo has no XML descriptor"))?;
    if xml.is_empty() {
        return Err(malformed("Agile EncryptionInfo XML is empty"));
    }
    Limits::bytes("Agile XML", xml.len(), limits.max_xml_bytes)?;

    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().check_comments = true;
    reader.config_mut().expand_empty_elements = true;
    let mut phase = 0u8;
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut attributes = 0usize;
    let mut declaration_seen = false;
    let mut first_event = true;
    let mut parsed = Parsed::default();

    loop {
        let decoder = reader.decoder();
        let (namespace, event) = reader
            .read_resolved_event()
            .map_err(|error| Error::Xml(error.to_string()))?;
        let was_first = first_event;
        first_event = false;
        match event {
            Event::Decl(declaration) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                if !was_first || declaration_seen || phase != 0 {
                    return Err(malformed(
                        "Agile XML declaration must be the first event and occur once",
                    ));
                }
                validate_declaration(&declaration)?;
                declaration_seen = true;
            },
            Event::Start(element) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("Agile XML depth", usize::MAX, limits.max_xml_depth))?;
                if child_depth > limits.max_xml_depth {
                    return Err(limit("Agile XML depth", child_depth, limits.max_xml_depth));
                }
                match phase {
                    0 => {
                        element_is(&namespace, &element, ENC_NS, b"encryption")?;
                        exact_attributes(
                            &reader,
                            &element,
                            decoder,
                            &[],
                            &mut attributes,
                            limits,
                            |_, _| Ok(()),
                        )?;
                        phase = 1;
                    },
                    1 => {
                        element_is(&namespace, &element, ENC_NS, b"keyData")?;
                        parsed.key_salt = Some(parse_key_data(
                            &reader,
                            &element,
                            decoder,
                            &mut attributes,
                            limits,
                        )?);
                        phase = 2;
                    },
                    3 => {
                        element_is(&namespace, &element, ENC_NS, b"dataIntegrity")?;
                        let (key, value) = parse_data_integrity(
                            &reader,
                            &element,
                            decoder,
                            &mut attributes,
                            limits,
                        )?;
                        parsed.encrypted_hmac_key = Some(key);
                        parsed.encrypted_hmac_value = Some(value);
                        phase = 4;
                    },
                    5 => {
                        element_is(&namespace, &element, ENC_NS, b"keyEncryptors")?;
                        exact_attributes(
                            &reader,
                            &element,
                            decoder,
                            &[],
                            &mut attributes,
                            limits,
                            |_, _| Ok(()),
                        )?;
                        phase = 6;
                    },
                    6 => {
                        element_is(&namespace, &element, ENC_NS, b"keyEncryptor")?;
                        parse_key_encryptor(&reader, &element, decoder, &mut attributes, limits)?;
                        phase = 7;
                    },
                    7 => {
                        element_is(&namespace, &element, PASSWORD_NS, b"encryptedKey")?;
                        parse_password_key(
                            &reader,
                            &element,
                            decoder,
                            &mut attributes,
                            limits,
                            &mut parsed,
                        )?;
                        phase = 8;
                    },
                    _ => return Err(malformed("Agile XML element order is invalid")),
                }
                depth = child_depth;
            },
            Event::End(element) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                match phase {
                    2 => {
                        end_is(&namespace, &element, ENC_NS, b"keyData")?;
                        phase = 3;
                    },
                    4 => {
                        end_is(&namespace, &element, ENC_NS, b"dataIntegrity")?;
                        phase = 5;
                    },
                    8 => {
                        end_is(&namespace, &element, PASSWORD_NS, b"encryptedKey")?;
                        phase = 9;
                    },
                    9 => {
                        end_is(&namespace, &element, ENC_NS, b"keyEncryptor")?;
                        phase = 10;
                    },
                    10 => {
                        end_is(&namespace, &element, ENC_NS, b"keyEncryptors")?;
                        phase = 11;
                    },
                    11 => {
                        end_is(&namespace, &element, ENC_NS, b"encryption")?;
                        phase = 12;
                    },
                    _ => return Err(malformed("Agile XML end-element order is invalid")),
                }
                depth = depth
                    .checked_sub(1)
                    .ok_or_else(|| malformed("Agile XML has an unexpected end element"))?;
            },
            Event::Text(text) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                let value = text
                    .xml_content(XmlVersion::Explicit1_0)
                    .map_err(|error| Error::Xml(error.to_string()))?;
                if !value.trim().is_empty() {
                    return Err(malformed("Agile XML cannot contain character data"));
                }
            },
            Event::Comment(comment) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                comment
                    .decode()
                    .map_err(|error| Error::Xml(error.to_string()))?;
            },
            Event::PI(instruction) => {
                count(&mut nodes, limits.max_xml_nodes, "Agile XML nodes")?;
                decoder
                    .decode(instruction.as_ref())
                    .map_err(|error| Error::Xml(error.to_string()))?;
            },
            Event::DocType(_) => return Err(malformed("DTD is forbidden in Agile XML")),
            Event::CData(_) | Event::GeneralRef(_) => {
                return Err(malformed("Agile XML cannot contain CDATA or entity nodes"));
            },
            Event::Empty(_) => {
                return Err(malformed("Agile XML empty-element expansion failed"));
            },
            Event::Eof => break,
        }
    }
    if phase != 12 || depth != 0 {
        return Err(malformed("Agile XML descriptor is incomplete"));
    }

    Ok(Info {
        spin_count: parsed
            .spin_count
            .ok_or_else(|| malformed("Agile encryptedKey has no spinCount"))?,
        key_salt: parsed
            .key_salt
            .ok_or_else(|| malformed("Agile XML has no keyData"))?,
        verifier_salt: parsed
            .verifier_salt
            .ok_or_else(|| malformed("Agile XML has no password salt"))?,
        encrypted_verifier: parsed
            .encrypted_verifier
            .ok_or_else(|| malformed("Agile XML has no encrypted verifier"))?,
        encrypted_verifier_hash: parsed
            .encrypted_verifier_hash
            .ok_or_else(|| malformed("Agile XML has no encrypted verifier hash"))?,
        encrypted_key: parsed
            .encrypted_key
            .ok_or_else(|| malformed("Agile XML has no encrypted content key"))?,
        encrypted_hmac_key: parsed
            .encrypted_hmac_key
            .ok_or_else(|| malformed("Agile XML has no encrypted HMAC key"))?,
        encrypted_hmac_value: parsed
            .encrypted_hmac_value
            .ok_or_else(|| malformed("Agile XML has no encrypted HMAC value"))?,
    })
}

fn element_is(
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    expected_namespace: &[u8],
    expected_local: &[u8],
) -> Result<()> {
    namespace_is(namespace, expected_namespace)?;
    if element.local_name().as_ref() != expected_local {
        return Err(malformed(format!(
            "unexpected Agile XML element '{}'",
            String::from_utf8_lossy(element.name().as_ref())
        )));
    }
    Ok(())
}

fn end_is(
    namespace: &ResolveResult<'_>,
    element: &quick_xml::events::BytesEnd<'_>,
    expected_namespace: &[u8],
    expected_local: &[u8],
) -> Result<()> {
    namespace_is(namespace, expected_namespace)?;
    if element.local_name().as_ref() != expected_local {
        return Err(malformed(format!(
            "unexpected Agile XML end element '{}'",
            String::from_utf8_lossy(element.name().as_ref())
        )));
    }
    Ok(())
}

fn namespace_is(namespace: &ResolveResult<'_>, expected: &[u8]) -> Result<()> {
    match namespace {
        ResolveResult::Bound(Namespace(actual)) if *actual == expected => Ok(()),
        ResolveResult::Bound(Namespace(actual)) => Err(malformed(format!(
            "unexpected Agile XML namespace '{}'",
            String::from_utf8_lossy(actual)
        ))),
        ResolveResult::Unbound => Err(malformed("Agile XML element has no namespace")),
        ResolveResult::Unknown(prefix) => Err(malformed(format!(
            "Agile XML namespace prefix '{}' is unbound",
            String::from_utf8_lossy(prefix)
        ))),
    }
}

fn parse_key_data(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    total: &mut usize,
    limits: &Limits,
) -> Result<[u8; BLOCK]> {
    const NAMES: &[&[u8]] = &[
        b"saltSize",
        b"blockSize",
        b"keyBits",
        b"hashSize",
        b"cipherAlgorithm",
        b"cipherChaining",
        b"hashAlgorithm",
        b"saltValue",
    ];
    let mut salt = None;
    exact_attributes(
        reader,
        element,
        decoder,
        NAMES,
        total,
        limits,
        |name, value| match name {
            b"saltSize" => exact_number(value, 16, "keyData.saltSize"),
            b"blockSize" => exact_number(value, 16, "keyData.blockSize"),
            b"keyBits" => exact_number(value, 128, "keyData.keyBits"),
            b"hashSize" => exact_number(value, 20, "keyData.hashSize"),
            b"cipherAlgorithm" => exact_text(value, "AES", "keyData.cipherAlgorithm"),
            b"cipherChaining" => exact_text(value, "ChainingModeCBC", "keyData.cipherChaining"),
            b"hashAlgorithm" => exact_text(value, "SHA-1", "keyData.hashAlgorithm"),
            b"saltValue" => {
                salt = Some(decode_array(value, "keyData.saltValue")?);
                Ok(())
            },
            _ => Err(malformed("unknown keyData attribute")),
        },
    )?;
    salt.ok_or_else(|| malformed("keyData.saltValue is missing"))
}

fn parse_data_integrity(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    total: &mut usize,
    limits: &Limits,
) -> Result<([u8; ENCRYPTED_HASH_BYTES], [u8; ENCRYPTED_HASH_BYTES])> {
    const NAMES: &[&[u8]] = &[b"encryptedHmacKey", b"encryptedHmacValue"];
    let mut key = None;
    let mut value = None;
    exact_attributes(
        reader,
        element,
        decoder,
        NAMES,
        total,
        limits,
        |name, raw| match name {
            b"encryptedHmacKey" => {
                key = Some(decode_array(raw, "dataIntegrity.encryptedHmacKey")?);
                Ok(())
            },
            b"encryptedHmacValue" => {
                value = Some(decode_array(raw, "dataIntegrity.encryptedHmacValue")?);
                Ok(())
            },
            _ => Err(malformed("unknown dataIntegrity attribute")),
        },
    )?;
    Ok((
        key.ok_or_else(|| malformed("dataIntegrity.encryptedHmacKey is missing"))?,
        value.ok_or_else(|| malformed("dataIntegrity.encryptedHmacValue is missing"))?,
    ))
}

fn parse_key_encryptor(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    total: &mut usize,
    limits: &Limits,
) -> Result<()> {
    exact_attributes(
        reader,
        element,
        decoder,
        &[b"uri"],
        total,
        limits,
        |_, value| {
            exact_text(
                value,
                "http://schemas.microsoft.com/office/2006/keyEncryptor/password",
                "keyEncryptor.uri",
            )
        },
    )
}

fn parse_password_key(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    total: &mut usize,
    limits: &Limits,
    parsed: &mut Parsed,
) -> Result<()> {
    const NAMES: &[&[u8]] = &[
        b"spinCount",
        b"saltSize",
        b"blockSize",
        b"keyBits",
        b"hashSize",
        b"cipherAlgorithm",
        b"cipherChaining",
        b"hashAlgorithm",
        b"saltValue",
        b"encryptedVerifierHashInput",
        b"encryptedVerifierHashValue",
        b"encryptedKeyValue",
    ];
    exact_attributes(
        reader,
        element,
        decoder,
        NAMES,
        total,
        limits,
        |name, value| match name {
            b"spinCount" => {
                let count = number(value, "encryptedKey.spinCount")?;
                check_spin(count, limits)?;
                parsed.spin_count = Some(count);
                Ok(())
            },
            b"saltSize" => exact_number(value, 16, "encryptedKey.saltSize"),
            b"blockSize" => exact_number(value, 16, "encryptedKey.blockSize"),
            b"keyBits" => exact_number(value, 128, "encryptedKey.keyBits"),
            b"hashSize" => exact_number(value, 20, "encryptedKey.hashSize"),
            b"cipherAlgorithm" => exact_text(value, "AES", "encryptedKey.cipherAlgorithm"),
            b"cipherChaining" => {
                exact_text(value, "ChainingModeCBC", "encryptedKey.cipherChaining")
            },
            b"hashAlgorithm" => exact_text(value, "SHA-1", "encryptedKey.hashAlgorithm"),
            b"saltValue" => {
                parsed.verifier_salt = Some(decode_array(value, "encryptedKey.saltValue")?);
                Ok(())
            },
            b"encryptedVerifierHashInput" => {
                parsed.encrypted_verifier =
                    Some(decode_array(value, "encryptedVerifierHashInput")?);
                Ok(())
            },
            b"encryptedVerifierHashValue" => {
                parsed.encrypted_verifier_hash =
                    Some(decode_array(value, "encryptedVerifierHashValue")?);
                Ok(())
            },
            b"encryptedKeyValue" => {
                parsed.encrypted_key = Some(decode_array(value, "encryptedKeyValue")?);
                Ok(())
            },
            _ => Err(malformed("unknown encryptedKey attribute")),
        },
    )
}

fn exact_attributes(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    allowed: &[&[u8]],
    total: &mut usize,
    limits: &Limits,
    mut visitor: impl FnMut(&[u8], &str) -> Result<()>,
) -> Result<()> {
    if allowed.len()
        > u16::BITS.try_into().map_err(|_err| {
            Error::InvalidLimit("Agile attribute schema width does not fit usize")
        })?
    {
        return Err(Error::InvalidLimit(
            "Agile attribute schema exceeds its bitset",
        ));
    }
    let mut seen = 0u16;
    for raw_attribute in element.attributes().with_checks(true) {
        count(total, limits.max_xml_attributes, "Agile XML attributes")?;
        let attribute = raw_attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Explicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?;
        let name = attribute.key.as_ref();
        if is_namespace_declaration(name) {
            validate_namespace_declaration(name, &value)?;
            continue;
        }
        match reader.resolver().resolve_attribute(attribute.key).0 {
            ResolveResult::Unbound => {},
            ResolveResult::Bound(_) => {
                return Err(malformed(format!(
                    "Agile attribute '{}' must be unqualified",
                    String::from_utf8_lossy(name)
                )));
            },
            ResolveResult::Unknown(prefix) => {
                return Err(malformed(format!(
                    "Agile attribute prefix '{}' is unbound",
                    String::from_utf8_lossy(&prefix)
                )));
            },
        }
        let index = allowed
            .iter()
            .position(|expected| *expected == name)
            .ok_or_else(|| {
                malformed(format!(
                    "unexpected Agile attribute '{}'",
                    String::from_utf8_lossy(name)
                ))
            })?;
        let shift =
            u32::try_from(index).map_err(|_err| malformed("Agile attribute index exceeds u32"))?;
        let bit = 1u16
            .checked_shl(shift)
            .ok_or_else(|| malformed("Agile attribute index exceeds its bitset"))?;
        if seen & bit != 0 {
            return Err(malformed(format!(
                "duplicate Agile attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
        seen |= bit;
        visitor(name, &value)?;
    }
    for (index, name) in allowed.iter().enumerate() {
        let shift =
            u32::try_from(index).map_err(|_err| malformed("Agile attribute index exceeds u32"))?;
        let bit = 1u16
            .checked_shl(shift)
            .ok_or_else(|| malformed("Agile attribute index exceeds its bitset"))?;
        if seen & bit == 0 {
            return Err(malformed(format!(
                "missing Agile attribute '{}'",
                String::from_utf8_lossy(name)
            )));
        }
    }
    Ok(())
}

fn is_namespace_declaration(name: &[u8]) -> bool {
    name == b"xmlns" || name.starts_with(b"xmlns:")
}

fn validate_namespace_declaration(name: &[u8], value: &str) -> Result<()> {
    if value.is_empty()
        || value
            .bytes()
            .any(|byte| matches!(byte, b' ' | b'\t' | b'\r' | b'\n'))
    {
        return Err(malformed(
            "Agile namespace URI is empty or contains whitespace",
        ));
    }
    if value == "http://www.w3.org/2000/xmlns/" {
        return Err(malformed("the xmlns namespace cannot be rebound"));
    }
    if let Some(prefix) = name.strip_prefix(b"xmlns:") {
        if prefix.is_empty() || prefix.contains(&b':') || prefix == b"xmlns" {
            return Err(malformed("Agile namespace prefix is invalid"));
        }
        if (prefix == b"xml") != (value == "http://www.w3.org/XML/1998/namespace") {
            return Err(malformed("the XML namespace may be bound only to 'xml'"));
        }
    } else if value == "http://www.w3.org/XML/1998/namespace" {
        return Err(malformed("the XML namespace may be bound only to 'xml'"));
    }
    Ok(())
}

fn validate_declaration(declaration: &BytesDecl<'_>) -> Result<()> {
    let version = declaration
        .xml_version()
        .map_err(|error| Error::Xml(error.to_string()))?;
    if version != XmlVersion::Explicit1_0 {
        return Err(malformed("Agile XML declaration must use version 1.0"));
    }
    if let Some(encoding) = declaration.encoding() {
        let encoding_name = encoding.map_err(|error| Error::Xml(error.to_string()))?;
        if !encoding_name.eq_ignore_ascii_case(b"UTF-8") {
            return Err(Error::Unsupported(format!(
                "Agile XML encoding '{}'",
                String::from_utf8_lossy(&encoding_name)
            )));
        }
    }
    if let Some(standalone) = declaration.standalone() {
        let standalone_value = standalone.map_err(|error| Error::Xml(error.to_string()))?;
        if !matches!(standalone_value.as_ref(), b"yes" | b"no") {
            return Err(malformed("Agile XML standalone must be 'yes' or 'no'"));
        }
    }
    Ok(())
}

fn exact_number(value: &str, expected: u32, field: &'static str) -> Result<()> {
    let actual = number(value, field)?;
    if actual != expected {
        return Err(Error::Unsupported(format!(
            "{field} is {actual}, expected {expected}"
        )));
    }
    Ok(())
}

fn number(value: &str, field: &'static str) -> Result<u32> {
    value
        .parse()
        .map_err(|_err| malformed(format!("{field} is not a u32")))
}

fn exact_text(value: &str, expected: &str, field: &'static str) -> Result<()> {
    if value != expected {
        return Err(Error::Unsupported(format!(
            "{field} is '{value}', expected '{expected}'"
        )));
    }
    Ok(())
}

fn decode_array<const N: usize>(value: &str, field: &'static str) -> Result<[u8; N]> {
    let mut decoded = [0u8; N];
    let length = BASE64
        .decode_slice(value, &mut decoded)
        .map_err(|error| malformed(format!("{field} is not valid {N}-byte base64: {error}")))?;
    if length != N {
        return Err(malformed(format!(
            "{field} has {length} bytes, expected {N}"
        )));
    }
    Ok(decoded)
}

fn count(value: &mut usize, maximum: usize, resource: &'static str) -> Result<()> {
    let next = value
        .checked_add(1)
        .ok_or_else(|| limit(resource, usize::MAX, maximum))?;
    if next > maximum {
        return Err(limit(resource, next, maximum));
    }
    *value = next;
    Ok(())
}

fn limit(resource: &'static str, actual: usize, maximum: usize) -> Error {
    Error::Limit {
        resource,
        actual: u64::try_from(actual).unwrap_or(u64::MAX),
        maximum: u64::try_from(maximum).unwrap_or(u64::MAX),
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::expect_used,
        reason = "test code panics on failure; expect keeps assertions concise"
    )]
    use super::*;
    use crate::ooxml::{Mode, open_with};

    const MATERIAL: Material = Material {
        verifier_salt: [
            0x9a, 0x4c, 0x79, 0x4b, 0x45, 0x20, 0x8c, 0xf6, 0x2c, 0x8a, 0xf5, 0xcd, 0x3a, 0xb6,
            0x9c, 0xe4,
        ],
        verifier: *b"fixed verifier!!",
        key_salt: [
            0xfd, 0xae, 0x22, 0x5a, 0xa3, 0xf2, 0x22, 0xf1, 0x36, 0x71, 0x4a, 0x25, 0x24, 0xc2,
            0xab, 0x23,
        ],
        content_key: *b"content key 1234",
        integrity_salt: *b"integrity salt bytes",
    };

    #[test]
    fn agile_round_trip_has_typed_password_and_integrity_failures() {
        let limits = Limits::default();
        let clear = Vec::from(&b"PK\x03\x04deterministic Agile package"[..]);
        let (info, encrypted) = encrypt_parts(clear.clone(), "correct horse", &MATERIAL, &limits)
            .expect("encrypt Agile parts");
        let compound =
            container::write(&info, encrypted.clone(), &limits).expect("wrap Agile parts");
        let opened = open_with(compound, "correct horse", &limits).expect("open Agile package");
        assert_eq!(opened.mode(), Some(Mode::Agile));
        assert_eq!(opened.bytes(), clear);
        assert!(matches!(
            decrypt(&info, encrypted.clone(), "wrong", &limits),
            Err(Error::Password)
        ));

        let mut tampered = encrypted;
        let last = tampered.len() - 1;
        tampered[last] ^= 1;
        assert!(matches!(
            decrypt(&info, tampered, "correct horse", &limits),
            Err(Error::Integrity)
        ));

        let mut tampered_prefix =
            encrypt_package(clear, &MATERIAL.content_key, &MATERIAL.key_salt, &limits)
                .expect("encrypt package for prefix tamper");
        tampered_prefix[0] ^= 1;
        assert!(matches!(
            decrypt(&info, tampered_prefix, "correct horse", &limits),
            Err(Error::Integrity)
        ));
    }

    #[test]
    fn final_segment_accepts_spec_padding_for_boundary_lengths() {
        let limits = Limits::default();
        for length in [1, BLOCK, SEGMENT, SEGMENT + 1] {
            let clear = vec![0x5a; length];
            let encrypted = encrypt_package(
                clear.clone(),
                &MATERIAL.content_key,
                &MATERIAL.key_salt,
                &limits,
            )
            .expect("encrypt segmented package");
            let clear_len = package_size(&encrypted, &limits).expect("package size");
            let opened = decrypt_package(
                encrypted,
                clear_len,
                &MATERIAL.content_key,
                &MATERIAL.key_salt,
            )
            .expect("decrypt segmented package");
            assert_eq!(opened, clear);
        }
    }

    #[test]
    fn parses_published_ms_offcrypto_3_11_vector() {
        let limits = Limits::default();
        let info = published_info();
        let parsed = parse(&info, &limits).expect("published Agile vector");
        assert_eq!(parsed.spin_count, 100_000);
        assert_eq!(
            parsed.key_salt,
            [
                0xfd, 0xae, 0x22, 0x5a, 0xa3, 0xf2, 0x22, 0xf1, 0x36, 0x71, 0x4a, 0x25, 0x24, 0xc2,
                0xab, 0x23,
            ]
        );
        assert_eq!(
            parsed.verifier_salt,
            [
                0xa6, 0x9b, 0x3a, 0x07, 0x56, 0xe6, 0xa8, 0x21, 0x57, 0x82, 0x8a, 0x6c, 0x9b, 0x5a,
                0xd6, 0x9d,
            ]
        );
    }

    #[test]
    fn parser_enforces_attribute_spin_and_tree_budgets() {
        let info = published_info();
        let spin_limits = Limits {
            max_spin_count: 99_999,
            ..Limits::default()
        };
        assert!(matches!(
            parse(&info, &spin_limits),
            Err(Error::Limit {
                resource: "Agile spin count",
                actual: 100_000,
                maximum: 99_999,
            })
        ));

        let depth_limits = Limits {
            max_xml_depth: 3,
            ..Limits::default()
        };
        assert!(matches!(
            parse(&info, &depth_limits),
            Err(Error::Limit {
                resource: "Agile XML depth",
                actual: 4,
                maximum: 3,
            })
        ));

        let attribute_limits = Limits {
            max_xml_attributes: 1,
            ..Limits::default()
        };
        assert!(matches!(
            parse(&info, &attribute_limits),
            Err(Error::Limit {
                resource: "Agile XML attributes",
                actual: 2,
                maximum: 1,
            })
        ));
    }

    #[test]
    fn parser_rejects_unknown_attributes_and_schema_exceeding_spin() {
        let info = published_info();
        let xml = std::str::from_utf8(&info[8..]).expect("published XML is UTF-8");
        let unknown = xml.replacen("saltSize=\"16\"", "saltSize=\"16\" extra=\"1\"", 1);
        assert!(matches!(
            parse(&with_prefix(&unknown), &Limits::default()),
            Err(Error::Malformed(_))
        ));

        let excessive = xml.replacen("spinCount=\"100000\"", "spinCount=\"10000001\"", 1);
        assert!(matches!(
            parse(
                &with_prefix(&excessive),
                &Limits {
                    max_spin_count: SPEC_MAX_SPIN_COUNT,
                    ..Limits::default()
                }
            ),
            Err(Error::Malformed(_))
        ));
    }

    #[test]
    fn declared_size_is_bounded_before_decryption() {
        let limits = Limits {
            max_plaintext_bytes: 32,
            ..Limits::default()
        };
        let mut encrypted = vec![0u8; 8 + 48];
        encrypted[..8].copy_from_slice(&33u64.to_le_bytes());
        assert!(matches!(
            package_size(&encrypted, &limits),
            Err(Error::Limit {
                resource: "declared plaintext",
                actual: 33,
                maximum: 32,
            })
        ));
    }

    fn published_info() -> Vec<u8> {
        with_prefix(
            r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption" xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
<keyData saltSize="16" blockSize="16" keyBits="128" hashSize="20" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA-1" saltValue="/a4iWqPyIvE2cUolJMKrIw=="/>
<dataIntegrity encryptedHmacKey="uwpAEFW1hQyD2O01kz1lhjevNw0ECyAA0u2OxDygsfY=" encryptedHmacValue="uf6HbJjtryJOjSFqrkqkNQY9NjNQUPI+xck8Q8y4mko="/>
<keyEncryptors><keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
<p:encryptedKey spinCount="100000" saltSize="16" blockSize="16" keyBits="128" hashSize="20" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA-1" saltValue="pps6B1bmqCFXgopsm1rWnQ==" encryptedVerifierHashInput="JYU4Q0u2BhqzQA5D4J/voA==" encryptedVerifierHashValue="eB2jX5mvhBJ+9O7ffC+6X2Mydz2glHOXx0T9Pn6nK+w=" encryptedKeyValue="2F86HG+xV3nGa27DElgqgw=="/>
</keyEncryptor></keyEncryptors></encryption>"#,
        )
    }

    fn with_prefix(xml: &str) -> Vec<u8> {
        let mut info = Vec::from([4, 0, 4, 0, 0x40, 0, 0, 0]);
        info.extend_from_slice(xml.as_bytes());
        info
    }
}
