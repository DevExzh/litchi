//! Format-neutral Office Binary Document RC4 CryptoAPI primitives.

use rc4::{KeyInit, Rc4, StreamCipher};
use sha1::{Digest, Sha1};
use subtle::ConstantTimeEq;
use zeroize::Zeroizing;

const CRYPTO_API_FLAG: u32 = 0x0000_0004;
const EXTERNAL_FLAG: u32 = 0x0000_0010;
const AES_FLAG: u32 = 0x0000_0020;
const CALG_RC4: u32 = 0x0000_6801;
const CALG_SHA1: u32 = 0x0000_8004;

#[derive(Debug)]
pub(crate) enum CryptoApiError {
    Malformed(String),
    UnsupportedVersion { major: u16, minor: u16 },
    UnsupportedAlgorithm,
}

pub(crate) struct CryptoApiHeader {
    salt: [u8; 16],
    encrypted_verifier: [u8; 16],
    encrypted_verifier_hash: [u8; 20],
    key_bits: usize,
}

#[derive(Clone)]
pub(crate) struct CryptoApiContext {
    secret: Zeroizing<[u8; 20]>,
    key_bits: usize,
}

pub(crate) fn parse_header(data: &[u8]) -> Result<CryptoApiHeader, CryptoApiError> {
    if data.len() < 12 {
        return malformed("CryptoAPI header is truncated");
    }
    let major = le_u16(data, 0)?;
    let minor = le_u16(data, 2)?;
    if !(2..=4).contains(&major) || minor != 2 {
        return Err(CryptoApiError::UnsupportedVersion { major, minor });
    }
    validate_flags(le_u32(data, 4)?)?;
    let header_size = usize::try_from(le_u32(data, 8)?)
        .map_err(|_| CryptoApiError::Malformed("CryptoAPI header size does not fit in memory".to_string()))?;
    if header_size < 32 {
        return malformed("CryptoAPI encryption header is shorter than 32 bytes");
    }
    let body = checked_slice(data, 12, header_size, "CryptoAPI encryption header")?;
    validate_flags(le_u32(body, 0)?)?;
    // SizeExtra and the two provider-specific reserved words do not affect RC4
    // and are ignored by Office and Apache POI.
    if le_u32(body, 8)? != CALG_RC4 {
        return Err(CryptoApiError::UnsupportedAlgorithm);
    }
    if !matches!(le_u32(body, 12)?, 0 | CALG_SHA1) {
        return Err(CryptoApiError::UnsupportedAlgorithm);
    }
    let raw_key_bits = le_u32(body, 16)?;
    let key_bits = if raw_key_bits == 0 { 40 } else { raw_key_bits };
    if !(40..=128).contains(&key_bits) || key_bits % 8 != 0 || le_u32(body, 20)? != 1 {
        return Err(CryptoApiError::UnsupportedAlgorithm);
    }
    let csp = &body[32..];
    if !csp.is_empty() && (csp.len() % 2 != 0 || !csp.ends_with(&[0, 0])) {
        return malformed("CryptoAPI provider name is not a terminated UTF-16LE string");
    }
    let verifier_offset = 12usize
        .checked_add(header_size)
        .ok_or_else(|| CryptoApiError::Malformed("CryptoAPI verifier offset overflow".to_string()))?;
    let verifier = checked_slice(data, verifier_offset, 60, "CryptoAPI verifier")?;
    if verifier_offset + 60 != data.len() || le_u32(verifier, 0)? != 16 {
        return malformed("CryptoAPI verifier has an invalid size");
    }
    if le_u32(verifier, 36)? != 20 {
        return malformed("CryptoAPI verifier hash size is not 20");
    }
    let mut salt = [0u8; 16];
    let mut encrypted_verifier = [0u8; 16];
    let mut encrypted_verifier_hash = [0u8; 20];
    salt.copy_from_slice(&verifier[4..20]);
    encrypted_verifier.copy_from_slice(&verifier[20..36]);
    encrypted_verifier_hash.copy_from_slice(&verifier[40..60]);
    Ok(CryptoApiHeader {
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
        key_bits: key_bits as usize,
    })
}

pub(crate) fn verify_password(
    header: &CryptoApiHeader,
    password: &str,
) -> Result<Option<CryptoApiContext>, CryptoApiError> {
    let password_bytes = Zeroizing::new(
        password
            .encode_utf16()
            .take(255)
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>(),
    );
    let mut hasher = Sha1::new();
    hasher.update(header.salt);
    hasher.update(password_bytes.as_slice());
    let context = CryptoApiContext {
        secret: Zeroizing::new(<[u8; 20]>::from(hasher.finalize())),
        key_bits: header.key_bits,
    };
    let key = derive_block_key(&context, 0);
    let mut cipher = Rc4::new_from_slice(key.as_slice())
        .map_err(|_| CryptoApiError::Malformed("invalid CryptoAPI RC4 key length".to_string()))?;
    let mut verifier = Zeroizing::new(header.encrypted_verifier);
    let mut verifier_hash = Zeroizing::new(header.encrypted_verifier_hash);
    cipher.apply_keystream(verifier.as_mut());
    cipher.apply_keystream(verifier_hash.as_mut());
    let calculated = Zeroizing::new(<[u8; 20]>::from(Sha1::digest(verifier.as_slice())));
    Ok(bool::from(calculated.ct_eq(verifier_hash.as_ref())).then_some(context))
}

pub(crate) fn apply_block_cipher(
    context: &CryptoApiContext,
    block: u32,
    data: &mut [u8],
) -> Result<(), CryptoApiError> {
    apply_block_cipher_at_offset(context, block, 0, data)
}

pub(crate) fn apply_block_cipher_at_offset(
    context: &CryptoApiContext,
    block: u32,
    offset: usize,
    data: &mut [u8],
) -> Result<(), CryptoApiError> {
    let key = derive_block_key(context, block);
    let mut cipher = Rc4::new_from_slice(key.as_slice())
        .map_err(|_| CryptoApiError::Malformed("invalid CryptoAPI RC4 key length".to_string()))?;
    if offset != 0 {
        let mut discarded = Zeroizing::new(vec![0u8; offset]);
        cipher.apply_keystream(discarded.as_mut_slice());
    }
    cipher.apply_keystream(data);
    Ok(())
}

fn derive_block_key(context: &CryptoApiContext, block: u32) -> Zeroizing<Vec<u8>> {
    let mut hasher = Sha1::new();
    hasher.update(context.secret.as_slice());
    hasher.update(block.to_le_bytes());
    let digest = Zeroizing::new(<[u8; 20]>::from(hasher.finalize()));
    let key_len = context.key_bits / 8;
    let output_len = if context.key_bits == 40 { 16 } else { key_len };
    let mut key = Zeroizing::new(vec![0u8; output_len]);
    key[..key_len].copy_from_slice(&digest[..key_len]);
    key
}

fn validate_flags(flags: u32) -> Result<(), CryptoApiError> {
    if flags & CRYPTO_API_FLAG == 0 || flags & (EXTERNAL_FLAG | AES_FLAG) != 0 {
        return Err(CryptoApiError::UnsupportedAlgorithm);
    }
    Ok(())
}

fn checked_slice<'a>(
    data: &'a [u8],
    offset: usize,
    len: usize,
    field: &str,
) -> Result<&'a [u8], CryptoApiError> {
    let end = offset
        .checked_add(len)
        .ok_or_else(|| CryptoApiError::Malformed(format!("{field} range overflow")))?;
    data.get(offset..end)
        .ok_or_else(|| CryptoApiError::Malformed(format!("{field} is truncated")))
}

fn le_u16(data: &[u8], offset: usize) -> Result<u16, CryptoApiError> {
    Ok(u16::from_le_bytes(
        checked_slice(data, offset, 2, "16-bit encryption field")?
            .try_into()
            .unwrap(),
    ))
}

fn le_u32(data: &[u8], offset: usize) -> Result<u32, CryptoApiError> {
    Ok(u32::from_le_bytes(
        checked_slice(data, offset, 4, "32-bit encryption field")?
            .try_into()
            .unwrap(),
    ))
}

fn malformed<T>(message: &str) -> Result<T, CryptoApiError> {
    Err(CryptoApiError::Malformed(message.to_string()))
}

#[cfg(test)]
pub(crate) fn test_context(secret: [u8; 20], key_bits: usize) -> CryptoApiContext {
    CryptoApiContext {
        secret: Zeroizing::new(secret),
        key_bits,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn block_cipher_supports_office_key_sizes_and_offsets() {
        for key_bits in [40, 56, 120, 128] {
            let context = test_context([0x5a; 20], key_bits);
            let original = vec![0xa5; 73];
            let mut encrypted = original.clone();
            apply_block_cipher_at_offset(&context, 7, 511, &mut encrypted).unwrap();
            assert_ne!(encrypted, original);
            apply_block_cipher_at_offset(&context, 7, 511, &mut encrypted).unwrap();
            assert_eq!(encrypted, original);
        }
    }
}
