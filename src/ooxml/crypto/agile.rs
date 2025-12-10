use crate::ooxml::error::{OoxmlError, Result};
use aes::Aes128;
use aes::cipher::{
    BlockEncryptMut, KeyIvInit,
    block_padding::{NoPadding, Pkcs7},
};
use base64::Engine as _;
use base64::engine::general_purpose::STANDARD as BASE64_STANDARD;
use cbc::Encryptor as Aes128Cbc;
use hmac::{Hmac, Mac};
use rand::TryRngCore;
use rand::rngs::OsRng;
use sha1::{Digest, Sha1};

use super::ole_encrypted_package::build_ole_encrypted_package;
use super::password_to_utf16le;

const AGILE_BLOCK_SIZE: usize = 16;
const AGILE_KEY_BITS: u32 = 128;
const AGILE_KEY_BYTES: usize = (AGILE_KEY_BITS as usize) / 8;
const AGILE_HASH_SIZE: usize = 20; // SHA‑1
const AGILE_SPIN_COUNT: u32 = 100_000;
const AGILE_SEGMENT_SIZE: usize = 4096;
const AGILE_ENCRYPTION_VERSION_MAJOR: u16 = 4;
const AGILE_ENCRYPTION_VERSION_MINOR: u16 = 4;
const AGILE_ENCRYPTION_FLAGS: u32 = 0x0000_0040;

const K_VERIFIER_INPUT_BLOCK: [u8; 8] = [0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79];
const K_HASHED_VERIFIER_BLOCK: [u8; 8] = [0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e];
const K_CRYPTO_KEY_BLOCK: [u8; 8] = [0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6];
const K_INTEGRITY_KEY_BLOCK: [u8; 8] = [0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6];
const K_INTEGRITY_VALUE_BLOCK: [u8; 8] = [0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33];

type HmacSha1 = Hmac<Sha1>;
type Aes128CbcEnc = Aes128Cbc<Aes128>;

pub fn encrypt_ooxml_package_agile(package_bytes: &[u8], password: &str) -> Result<Vec<u8>> {
    if package_bytes.is_empty() {
        return Err(OoxmlError::InvalidFormat(
            "cannot encrypt empty OOXML package".to_string(),
        ));
    }

    // 1) Generate random salts/keys (all OsRng)
    let mut rng = OsRng;

    let mut verifier_salt = [0u8; AGILE_BLOCK_SIZE];
    let mut verifier = [0u8; AGILE_BLOCK_SIZE];
    let mut key_salt = [0u8; AGILE_BLOCK_SIZE];
    let mut content_key = [0u8; AGILE_KEY_BYTES];
    let mut integrity_salt = [0u8; AGILE_HASH_SIZE];

    rng.try_fill_bytes(&mut verifier_salt)
        .map_err(|e| OoxmlError::Other(format!("failed to generate Agile verifier salt: {e}")))?;
    rng.try_fill_bytes(&mut verifier)
        .map_err(|e| OoxmlError::Other(format!("failed to generate Agile verifier: {e}")))?;
    rng.try_fill_bytes(&mut key_salt)
        .map_err(|e| OoxmlError::Other(format!("failed to generate Agile key salt: {e}")))?;
    rng.try_fill_bytes(&mut content_key)
        .map_err(|e| OoxmlError::Other(format!("failed to generate Agile content key: {e}")))?;
    rng.try_fill_bytes(&mut integrity_salt)
        .map_err(|e| OoxmlError::Other(format!("failed to generate Agile integrity salt: {e}")))?;

    // 2) Password hash
    let pw_hash = hash_password_agile(password, &verifier_salt, AGILE_SPIN_COUNT);

    // 3) Verifier structures
    let encrypted_verifier = hash_input_agile(
        &verifier_salt,
        &pw_hash,
        &K_VERIFIER_INPUT_BLOCK,
        &verifier,
        AGILE_BLOCK_SIZE,
        AGILE_KEY_BYTES,
    )?;

    let mut sha = Sha1::new();
    sha.update(verifier);
    let verifier_hash = sha.finalize().to_vec();

    let encrypted_verifier_hash = hash_input_agile(
        &verifier_salt,
        &pw_hash,
        &K_HASHED_VERIFIER_BLOCK,
        &verifier_hash,
        AGILE_BLOCK_SIZE,
        AGILE_KEY_BYTES,
    )?;

    // 4) Encrypted content key
    let encrypted_key = hash_input_agile(
        &verifier_salt,
        &pw_hash,
        &K_CRYPTO_KEY_BLOCK,
        &content_key,
        AGILE_BLOCK_SIZE,
        AGILE_KEY_BYTES,
    )?;

    // 5) EncryptedPackage (StreamSize + segmented AES-CBC)
    let encrypted_package = encrypt_agile_package_stream(&content_key, &key_salt, package_bytes)?;

    // 6) DataIntegrity
    // Per MS-OFFCRYPTO 2.3.4.14 and Apache POI's AgileEncryptor:
    // - integritySalt (hashSize bytes) is the HMAC key
    // - integritySalt is zero-padded to a block multiple only for AES encryption
    let integrity_salt_padded = pad_zero_to_block_multiple(&integrity_salt, AGILE_BLOCK_SIZE);

    let iv_hmac_key = generate_iv_agile(&key_salt, Some(&K_INTEGRITY_KEY_BLOCK), AGILE_BLOCK_SIZE);
    let cipher = Aes128CbcEnc::new_from_slices(&content_key, &iv_hmac_key)
        .map_err(|_| OoxmlError::InvalidFormat("invalid AES key/iv for integrity key".into()))?;
    let encrypted_hmac_key = cipher.encrypt_padded_vec_mut::<NoPadding>(&integrity_salt_padded);

    let mut mac = HmacSha1::new_from_slice(&integrity_salt)
        .map_err(|e| OoxmlError::Other(format!("failed to init HMAC-SHA1: {e}")))?;
    mac.update(&encrypted_package);
    let hmac_value = mac.finalize().into_bytes().to_vec();
    let hmac_value_padded = pad_zero_to_block_multiple(&hmac_value, AGILE_BLOCK_SIZE);

    let iv_hmac_value =
        generate_iv_agile(&key_salt, Some(&K_INTEGRITY_VALUE_BLOCK), AGILE_BLOCK_SIZE);
    let cipher = Aes128CbcEnc::new_from_slices(&content_key, &iv_hmac_value)
        .map_err(|_| OoxmlError::InvalidFormat("invalid AES key/iv for integrity value".into()))?;
    let encrypted_hmac_value = cipher.encrypt_padded_vec_mut::<NoPadding>(&hmac_value_padded);

    // 7) Build EncryptionInfo XML + binary prefix
    let xml = build_agile_encryption_info_xml(
        &key_salt,
        &verifier_salt,
        &encrypted_verifier,
        &encrypted_verifier_hash,
        &encrypted_key,
        &encrypted_hmac_key,
        &encrypted_hmac_value,
    );
    let xml_bytes = xml.into_bytes();

    let mut encryption_info = Vec::with_capacity(8 + xml_bytes.len());
    encryption_info.extend_from_slice(&AGILE_ENCRYPTION_VERSION_MAJOR.to_le_bytes());
    encryption_info.extend_from_slice(&AGILE_ENCRYPTION_VERSION_MINOR.to_le_bytes());
    encryption_info.extend_from_slice(&AGILE_ENCRYPTION_FLAGS.to_le_bytes());
    encryption_info.extend_from_slice(&xml_bytes);

    build_ole_encrypted_package(&encryption_info, &encrypted_package)
}

fn hash_password_agile(password: &str, salt: &[u8], spin_count: u32) -> Vec<u8> {
    let mut sha = Sha1::new();
    sha.update(salt);
    // UTF‑16LE of password
    let pw_bytes = password_to_utf16le(password);
    sha.update(&pw_bytes);
    let mut hash = sha.finalize().to_vec();

    let mut iter = [0u8; 4];
    for i in 0..spin_count {
        iter.copy_from_slice(&i.to_le_bytes());
        let mut sha = Sha1::new();
        // iteratorFirst = true: H(iterator || hash)
        sha.update(iter);
        sha.update(&hash);
        hash = sha.finalize().to_vec();
    }

    hash
}

fn generate_key_agile(password_hash: &[u8], block_key: &[u8], key_size: usize) -> Vec<u8> {
    let mut sha = Sha1::new();
    sha.update(password_hash);
    sha.update(block_key);
    let key = sha.finalize().to_vec(); // H(H_n || blockKey)

    // pad/truncate with 0x36 to key_size
    if key.len() == key_size {
        return key;
    }
    let mut out = vec![0x36u8; key_size];
    let copy = out.len().min(key.len());
    out[..copy].copy_from_slice(&key[..copy]);
    out
}

fn pad_36_to_block(mut iv: Vec<u8>, block_size: usize) -> Vec<u8> {
    if iv.len() == block_size {
        return iv;
    }
    if iv.len() > block_size {
        iv.truncate(block_size);
        return iv;
    }
    iv.resize(block_size, 0x36);
    iv
}

fn generate_iv_agile(key_salt: &[u8], block_key: Option<&[u8]>, block_size: usize) -> Vec<u8> {
    let iv = if let Some(block_key) = block_key {
        let mut sha = Sha1::new();
        sha.update(key_salt);
        sha.update(block_key);
        sha.finalize().to_vec()
    } else {
        key_salt.to_vec()
    };
    pad_36_to_block(iv, block_size)
}

fn pad_zero_to_block_multiple(input: &[u8], block_size: usize) -> Vec<u8> {
    if input.is_empty() {
        return vec![0u8; block_size];
    }
    let mut out = Vec::from(input);
    let rem = out.len() % block_size;
    if rem != 0 {
        out.resize(out.len() + (block_size - rem), 0);
    }
    out
}

fn hash_input_agile(
    verifier_salt: &[u8],
    pw_hash: &[u8],
    block_key: &[u8],
    input: &[u8],
    block_size: usize,
    key_size: usize,
) -> Result<Vec<u8>> {
    if input.is_empty() {
        return Err(OoxmlError::InvalidFormat(
            "Agile hashInput called with empty input".to_string(),
        ));
    }

    let inter_key = generate_key_agile(pw_hash, block_key, key_size);
    let iv = generate_iv_agile(verifier_salt, None, block_size);
    let cipher = Aes128CbcEnc::new_from_slices(&inter_key, &iv)
        .map_err(|_| OoxmlError::InvalidFormat("invalid AES-128 key/iv".into()))?;

    let padded = pad_zero_to_block_multiple(input, block_size);
    // NoPadding because we padded manually
    let ciphertext = cipher.encrypt_padded_vec_mut::<NoPadding>(&padded);

    Ok(ciphertext)
}

fn encrypt_agile_package_stream(
    content_key: &[u8],
    key_salt: &[u8],
    plain: &[u8],
) -> Result<Vec<u8>> {
    let mut out = Vec::with_capacity(8 + plain.len() + 64);
    let stream_size = plain.len() as u64;
    out.extend_from_slice(&stream_size.to_le_bytes()); // StreamSize (unencrypted)

    let mut offset = 0usize;
    let mut block_index: u32 = 0;

    while offset < plain.len() {
        let remaining = plain.len() - offset;
        let this_len = remaining.min(AGILE_SEGMENT_SIZE);

        let is_last = offset + this_len == plain.len();

        let segment = &plain[offset..offset + this_len];

        let block_key = block_index.to_le_bytes(); // 4 bytes, LE
        let iv = generate_iv_agile(key_salt, Some(&block_key), AGILE_BLOCK_SIZE);

        if is_last {
            let cipher = Aes128CbcEnc::new_from_slices(content_key, &iv)
                .map_err(|_| OoxmlError::InvalidFormat("invalid AES key/iv".into()))?;
            let ct = cipher.encrypt_padded_vec_mut::<Pkcs7>(segment);
            out.extend_from_slice(&ct);
        } else {
            // 4096 is multiple of 16 => can use NoPadding
            let cipher = Aes128CbcEnc::new_from_slices(content_key, &iv)
                .map_err(|_| OoxmlError::InvalidFormat("invalid AES key/iv".into()))?;
            let ct = cipher.encrypt_padded_vec_mut::<NoPadding>(segment);
            out.extend_from_slice(&ct);
        }

        offset += this_len;
        block_index += 1;
    }

    Ok(out)
}

fn build_agile_encryption_info_xml(
    key_salt: &[u8],
    verifier_salt: &[u8],
    encrypted_verifier: &[u8],
    encrypted_verifier_hash: &[u8],
    encrypted_key: &[u8],
    encrypted_hmac_key: &[u8],
    encrypted_hmac_value: &[u8],
) -> String {
    let key_salt_b64 = BASE64_STANDARD.encode(key_salt);
    let verifier_salt_b64 = BASE64_STANDARD.encode(verifier_salt);
    let enc_ver_b64 = BASE64_STANDARD.encode(encrypted_verifier);
    let enc_ver_hash_b64 = BASE64_STANDARD.encode(encrypted_verifier_hash);
    let enc_key_b64 = BASE64_STANDARD.encode(encrypted_key);
    let enc_hmac_key_b64 = BASE64_STANDARD.encode(encrypted_hmac_key);
    let enc_hmac_val_b64 = BASE64_STANDARD.encode(encrypted_hmac_value);

    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<encryption xmlns="http://schemas.microsoft.com/office/2006/encryption"
 xmlns:p="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
  <keyData saltSize="{salt_sz}" blockSize="{blk_sz}" keyBits="{key_bits}" hashSize="{hash_sz}"
           cipherAlgorithm="AES" cipherChaining="ChainingModeCBC" hashAlgorithm="SHA1"
           saltValue="{key_salt}"/>
  <dataIntegrity encryptedHmacKey="{enc_hmac_key}" encryptedHmacValue="{enc_hmac_val}"/>
  <keyEncryptors>
    <keyEncryptor uri="http://schemas.microsoft.com/office/2006/keyEncryptor/password">
      <p:encryptedKey spinCount="{spin}" saltSize="{salt_sz}" blockSize="{blk_sz}" keyBits="{key_bits}"
                      hashSize="{hash_sz}" cipherAlgorithm="AES" cipherChaining="ChainingModeCBC"
                      hashAlgorithm="SHA1" saltValue="{ver_salt}"
                      encryptedVerifierHashInput="{enc_ver}" encryptedVerifierHashValue="{enc_ver_hash}"
                      encryptedKeyValue="{enc_key}"/>
    </keyEncryptor>
  </keyEncryptors>
</encryption>"#,
        salt_sz = AGILE_BLOCK_SIZE,
        blk_sz = AGILE_BLOCK_SIZE,
        key_bits = AGILE_KEY_BITS,
        hash_sz = AGILE_HASH_SIZE,
        key_salt = key_salt_b64,
        ver_salt = verifier_salt_b64,
        enc_ver = enc_ver_b64,
        enc_ver_hash = enc_ver_hash_b64,
        enc_key = enc_key_b64,
        enc_hmac_key = enc_hmac_key_b64,
        enc_hmac_val = enc_hmac_val_b64,
        spin = AGILE_SPIN_COUNT,
    )
}
