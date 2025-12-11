use crate::ooxml::error::{OoxmlError, Result};
use aes::Aes128;
use aes::cipher::{BlockDecrypt, BlockEncrypt, KeyInit, generic_array::GenericArray};
use rand::TryRngCore;
use rand::rngs::OsRng;
use sha1::{Digest, Sha1};

use super::ole_encrypted_package::build_ole_encrypted_package;
use super::password_to_utf16le;

#[derive(Debug, Clone, Copy)]
struct Standard2007Verifier {
    salt: [u8; 16],
    encrypted_verifier: [u8; 16],
    encrypted_verifier_hash: [u8; 32],
}

pub fn encrypt_ooxml_package_standard_2007(
    package_bytes: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    if package_bytes.is_empty() {
        return Err(OoxmlError::InvalidFormat(
            "cannot encrypt empty OOXML package".to_string(),
        ));
    }

    let mut salt = [0u8; 16];
    let mut verifier = [0u8; 16];
    let mut rng = OsRng;

    rng.try_fill_bytes(&mut salt)
        .map_err(|e| OoxmlError::Other(format!("failed to generate encryption salt: {e}")))?;
    rng.try_fill_bytes(&mut verifier)
        .map_err(|e| OoxmlError::Other(format!("failed to generate encryption verifier: {e}")))?;

    let key = derive_standard2007_key(password, &salt, 50_000, 16)?;
    let (encrypted_verifier, encrypted_verifier_hash) = encrypt_verifier(&key, &verifier);

    let encryption_info =
        build_encryption_info_standard2007(&salt, &encrypted_verifier, &encrypted_verifier_hash);

    let encrypted_package = encrypt_package_stream(&key, package_bytes)?;

    build_ole_encrypted_package(&encryption_info, &encrypted_package)
}

fn derive_standard2007_key(
    password: &str,
    salt: &[u8; 16],
    spin_count: u32,
    key_size: usize,
) -> Result<Vec<u8>> {
    if key_size == 0 || key_size > 32 {
        return Err(OoxmlError::InvalidFormat(
            "unsupported key size for Standard 2007 encryption".to_string(),
        ));
    }

    let pw_bytes = password_to_utf16le(password);

    let mut sha = Sha1::new();
    sha.update(salt);
    sha.update(&pw_bytes);
    let mut hash = sha.finalize().to_vec();

    for i in 0..spin_count {
        let mut iter = [0u8; 4];
        iter.copy_from_slice(&i.to_le_bytes());
        let mut sha = Sha1::new();
        sha.update(iter);
        sha.update(&hash);
        hash = sha.finalize().to_vec();
    }

    let block_key = [0u8; 4];
    let mut sha = Sha1::new();
    sha.update(&hash);
    sha.update(block_key);
    let intermediate = sha.finalize().to_vec();

    let mut key_material = vec![0x36u8; 20];
    let copy_len = key_material.len().min(intermediate.len());
    key_material[..copy_len].copy_from_slice(&intermediate[..copy_len]);

    let x1 = fill_and_xor_sha1(&key_material, 0x36);
    let x2 = fill_and_xor_sha1(&key_material, 0x5c);

    let mut combined = Vec::with_capacity(x1.len() + x2.len());
    combined.extend_from_slice(&x1);
    combined.extend_from_slice(&x2);

    Ok(combined[..key_size].to_vec())
}

fn fill_and_xor_sha1(input: &[u8], fill: u8) -> Vec<u8> {
    let mut buff = [fill; 64];
    let len = buff.len().min(input.len());
    for i in 0..len {
        buff[i] ^= input[i];
    }
    let mut sha = Sha1::new();
    sha.update(buff);
    sha.finalize().to_vec()
}

fn encrypt_verifier(key: &[u8], verifier: &[u8; 16]) -> ([u8; 16], [u8; 32]) {
    let cipher = Aes128::new_from_slice(key).expect("AES-128 key must be 16 bytes");

    let mut encrypted_verifier = *verifier;
    let block = GenericArray::from_mut_slice(&mut encrypted_verifier);
    cipher.encrypt_block(block);

    let mut sha = Sha1::new();
    sha.update(verifier);
    let hash = sha.finalize();

    let mut padded = [0u8; 32];
    padded[..hash.len()].copy_from_slice(&hash);

    let mut encrypted_hash = padded;
    for chunk in encrypted_hash.chunks_mut(16) {
        let block = GenericArray::from_mut_slice(chunk);
        cipher.encrypt_block(block);
    }

    (encrypted_verifier, encrypted_hash)
}

fn build_encryption_info_standard2007(
    salt: &[u8; 16],
    encrypted_verifier: &[u8; 16],
    encrypted_verifier_hash: &[u8; 32],
) -> Vec<u8> {
    let mut buf = Vec::with_capacity(256);

    buf.extend_from_slice(&3u16.to_le_bytes());
    buf.extend_from_slice(&2u16.to_le_bytes());
    buf.extend_from_slice(&0x24u32.to_le_bytes());

    let header_start = buf.len();
    buf.extend_from_slice(&0u32.to_le_bytes());

    buf.extend_from_slice(&0x24u32.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf.extend_from_slice(&0x660Eu32.to_le_bytes());
    buf.extend_from_slice(&0x8004u32.to_le_bytes());
    buf.extend_from_slice(&128u32.to_le_bytes());
    buf.extend_from_slice(&0x18u32.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());
    buf.extend_from_slice(&0u32.to_le_bytes());

    let csp_name = "Microsoft Enhanced RSA and AES Cryptographic Provider";
    for ch in csp_name.encode_utf16() {
        let bytes = ch.to_le_bytes();
        buf.push(bytes[0]);
        buf.push(bytes[1]);
    }
    buf.extend_from_slice(&0u16.to_le_bytes());

    let header_size = (buf.len() - header_start - 4) as u32;
    buf[header_start..header_start + 4].copy_from_slice(&header_size.to_le_bytes());

    buf.extend_from_slice(&16u32.to_le_bytes());
    buf.extend_from_slice(salt);
    buf.extend_from_slice(encrypted_verifier);
    buf.extend_from_slice(&20u32.to_le_bytes());
    buf.extend_from_slice(encrypted_verifier_hash);

    buf
}

fn encrypt_package_stream(key: &[u8], data: &[u8]) -> Result<Vec<u8>> {
    let cipher = Aes128::new_from_slice(key).map_err(|_| {
        OoxmlError::InvalidFormat("invalid AES-128 key length for package encryption".to_string())
    })?;

    let block_size = 16usize;
    let pad_len = block_size - (data.len() % block_size);
    let pad_len = if pad_len == 0 { block_size } else { pad_len };

    let mut buf = Vec::with_capacity(8 + data.len() + pad_len);
    let stream_size = data.len() as u64;
    buf.extend_from_slice(&stream_size.to_le_bytes());

    let mut padded = Vec::with_capacity(data.len() + pad_len);
    padded.extend_from_slice(data);
    padded.resize(data.len() + pad_len, pad_len as u8);

    for chunk in padded.chunks_mut(block_size) {
        let block = GenericArray::from_mut_slice(chunk);
        cipher.encrypt_block(block);
    }

    buf.extend_from_slice(&padded);
    Ok(buf)
}

pub fn decrypt_ooxml_package_standard_2007(
    encryption_info: &[u8],
    encrypted_package: &[u8],
    password: &str,
) -> Result<Vec<u8>> {
    if encryption_info.len() < 8 {
        return Err(OoxmlError::InvalidFormat(
            "EncryptionInfo stream too short for Standard 2007 header".to_string(),
        ));
    }

    if encrypted_package.len() < 8 + 16 {
        return Err(OoxmlError::InvalidFormat(
            "EncryptedPackage stream too short for Standard 2007 payload".to_string(),
        ));
    }

    let verifier = parse_encryption_info_standard2007(encryption_info)?;

    let key = derive_standard2007_key(password, &verifier.salt, 50_000, 16)?;

    verify_standard2007_password(&key, &verifier)?;

    decrypt_package_stream(&key, encrypted_package)
}

fn parse_encryption_info_standard2007(info: &[u8]) -> Result<Standard2007Verifier> {
    if info.len() < 8 + 4 {
        return Err(OoxmlError::InvalidFormat(
            "EncryptionInfo stream too short for Standard 2007 header".to_string(),
        ));
    }

    let version_major = u16::from_le_bytes([info[0], info[1]]);
    let version_minor = u16::from_le_bytes([info[2], info[3]]);
    if version_major != 3 || version_minor != 2 {
        return Err(OoxmlError::InvalidFormat(format!(
            "unsupported Standard 2007 EncryptionInfo version: {}.{}",
            version_major, version_minor
        )));
    }

    let header_size = u32::from_le_bytes([info[8], info[9], info[10], info[11]]) as usize;
    let header_start = 12usize;
    let header_end = header_start.checked_add(header_size).ok_or_else(|| {
        OoxmlError::InvalidFormat("EncryptionInfo header size overflow".to_string())
    })?;

    let mut offset = header_end;
    if info.len() < offset + 4 + 16 + 16 + 4 + 32 {
        return Err(OoxmlError::InvalidFormat(
            "EncryptionInfo stream too short for Standard 2007 verifier".to_string(),
        ));
    }

    let salt_size = u32::from_le_bytes([
        info[offset],
        info[offset + 1],
        info[offset + 2],
        info[offset + 3],
    ]);
    offset += 4;

    if salt_size != 16 {
        return Err(OoxmlError::InvalidFormat(format!(
            "unexpected Standard 2007 salt size: {} (expected 16)",
            salt_size
        )));
    }

    let mut salt = [0u8; 16];
    salt.copy_from_slice(&info[offset..offset + 16]);
    offset += 16;

    let mut encrypted_verifier = [0u8; 16];
    encrypted_verifier.copy_from_slice(&info[offset..offset + 16]);
    offset += 16;

    let hash_size = u32::from_le_bytes([
        info[offset],
        info[offset + 1],
        info[offset + 2],
        info[offset + 3],
    ]);
    offset += 4;

    if hash_size != 20 {
        return Err(OoxmlError::InvalidFormat(format!(
            "unexpected Standard 2007 verifier hash size: {} (expected 20)",
            hash_size
        )));
    }

    let mut encrypted_verifier_hash = [0u8; 32];
    encrypted_verifier_hash.copy_from_slice(&info[offset..offset + 32]);

    Ok(Standard2007Verifier {
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
    })
}

fn verify_standard2007_password(key: &[u8], verifier: &Standard2007Verifier) -> Result<()> {
    let cipher = Aes128::new_from_slice(key).map_err(|_| {
        OoxmlError::InvalidFormat(
            "invalid AES-128 key length for Standard 2007 password verification".to_string(),
        )
    })?;

    let mut decrypted_verifier = verifier.encrypted_verifier;
    let block = GenericArray::from_mut_slice(&mut decrypted_verifier);
    cipher.decrypt_block(block);

    let mut sha = Sha1::new();
    sha.update(decrypted_verifier);
    let verifier_hash = sha.finalize();

    let mut decrypted_hash = verifier.encrypted_verifier_hash;
    for chunk in decrypted_hash.chunks_mut(16) {
        let block = GenericArray::from_mut_slice(chunk);
        cipher.decrypt_block(block);
    }

    if decrypted_hash[..verifier_hash.len()] != verifier_hash[..] {
        return Err(OoxmlError::InvalidFormat(
            "incorrect password for Standard 2007 encrypted OOXML package".to_string(),
        ));
    }

    Ok(())
}
fn decrypt_package_stream(key: &[u8], encrypted: &[u8]) -> Result<Vec<u8>> {
    let cipher = Aes128::new_from_slice(key).map_err(|_| {
        OoxmlError::InvalidFormat("invalid AES-128 key length for package decryption".to_string())
    })?;

    if encrypted.len() < 8 + 16 {
        return Err(OoxmlError::InvalidFormat(
            "EncryptedPackage stream too short for Standard 2007 decryption".to_string(),
        ));
    }

    let mut size_bytes = [0u8; 8];
    size_bytes.copy_from_slice(&encrypted[..8]);
    let stream_size = u64::from_le_bytes(size_bytes) as usize;

    let mut data = encrypted[8..].to_vec();
    if !data.len().is_multiple_of(16) {
        return Err(OoxmlError::InvalidFormat(
            "EncryptedPackage ciphertext length is not a multiple of 16 bytes".to_string(),
        ));
    }

    for chunk in data.chunks_mut(16) {
        let block = GenericArray::from_mut_slice(chunk);
        cipher.decrypt_block(block);
    }

    if data.is_empty() {
        return Err(OoxmlError::InvalidFormat(
            "EncryptedPackage has no payload after decryption".to_string(),
        ));
    }

    let pad_len = *data.last().unwrap() as usize;
    if pad_len == 0 || pad_len > 16 || pad_len > data.len() {
        return Err(OoxmlError::InvalidFormat(
            "invalid PKCS7 padding in Standard 2007 EncryptedPackage".to_string(),
        ));
    }

    let plain_len = data.len() - pad_len;
    if plain_len < stream_size {
        return Err(OoxmlError::InvalidFormat(
            "decrypted stream smaller than declared StreamSize".to_string(),
        ));
    }

    data.truncate(stream_size);
    Ok(data)
}
