use crate::ooxml::error::{OoxmlError, Result};
use aes::Aes128;
use aes::cipher::{BlockEncrypt, KeyInit, generic_array::GenericArray};
use rand::TryRngCore;
use rand::rngs::OsRng;
use sha1::{Digest, Sha1};

use super::ole_encrypted_package::build_ole_encrypted_package;
use super::password_to_utf16le;

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
