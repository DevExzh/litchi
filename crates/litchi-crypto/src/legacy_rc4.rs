//! Format-neutral Office Binary Document legacy RC4 primitives.
//!
//! This module implements the non-CryptoAPI Office Binary Document RC4
//! profile from [MS-OFFCRYPTO] section 2.3.6.  It is deliberately separate
//! from [`crate::rc4`], which implements the CryptoAPI profile from section
//! 2.3.5.  DOC and XLS own the format-specific stream layout and choose the
//! block size required by their respective specifications.

use md5::{Digest, Md5};
use rc4::{KeyInit, Rc4, StreamCipher};
use std::fmt;
use subtle::ConstantTimeEq;
use zeroize::Zeroizing;

/// Size of an Office Binary Document RC4 encryption header.
pub const HEADER_LEN: usize = 52;

/// Maximum block size accepted by [`apply_at`].
///
/// DOC uses 512-byte blocks and XLS uses 1024-byte blocks.  Keeping this
/// bound in the shared primitive prevents an untrusted caller from turning a
/// block-size argument into an unbounded temporary allocation.
pub const MAX_BLOCK_SIZE: usize = 4096;

const MAX_PASSWORD_UNITS: usize = 255;

/// A malformed legacy RC4 header or an invalid bounded operation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    /// The input does not satisfy the fixed legacy RC4 grammar.
    Malformed(String),
    /// The header identifies a profile other than version 1.1.
    UnsupportedVersion { major: u16, minor: u16 },
    /// The password exceeds the format's maximum UTF-16 length.
    PasswordTooLong { units: usize },
    /// The caller supplied a block size outside the bounded shared range.
    InvalidBlockSize { size: usize },
    /// The requested stream range overflows `usize` or the format block index.
    StreamRangeOverflow,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Malformed(message) => write!(formatter, "malformed legacy RC4 data: {message}"),
            Self::UnsupportedVersion { major, minor } => {
                write!(formatter, "unsupported legacy RC4 version {major}.{minor}")
            },
            Self::PasswordTooLong { units } => write!(
                formatter,
                "legacy RC4 password contains {units} UTF-16 code units; maximum is {MAX_PASSWORD_UNITS}"
            ),
            Self::InvalidBlockSize { size } => write!(
                formatter,
                "legacy RC4 block size {size} is outside the bounded 1..={MAX_BLOCK_SIZE} range"
            ),
            Self::StreamRangeOverflow => {
                formatter.write_str("legacy RC4 stream range exceeds the supported block index")
            },
        }
    }
}

impl std::error::Error for Error {}

/// A validated Office Binary Document RC4 header.
pub struct Header {
    salt: [u8; 16],
    encrypted_verifier: [u8; 16],
    encrypted_verifier_hash: [u8; 16],
}

/// Password-derived legacy RC4 material with zeroizing storage.
///
/// Contexts are move-only so secret material is not duplicated accidentally.
/// Format consumers borrow a context for each absolute stream range they
/// decrypt or encrypt.
pub struct Context {
    secret: Zeroizing<[u8; 5]>,
}

/// Build a legacy RC4 header and its reusable password context.
pub fn build_header(
    password: &str,
    salt: &[u8; 16],
    verifier: &[u8; 16],
) -> Result<(Vec<u8>, Context), Error> {
    let context = context(password, salt)?;
    let key = derive_block_key(&context, 0);
    let mut encrypted = Zeroizing::new([0u8; 32]);
    encrypted[..16].copy_from_slice(verifier);
    encrypted[16..].copy_from_slice(&Md5::digest(verifier));
    let mut cipher = Rc4::new_from_slice(key.as_slice())
        .map_err(|_| Error::Malformed("invalid legacy RC4 key length".to_string()))?;
    // The verifier and its MD5 digest are one continuous RC4 stream.  The
    // stream must not be reset between the two 16-byte fields.
    cipher.apply_keystream(encrypted.as_mut());

    let mut header = Vec::with_capacity(HEADER_LEN);
    header.extend_from_slice(&1u16.to_le_bytes());
    header.extend_from_slice(&1u16.to_le_bytes());
    header.extend_from_slice(salt);
    header.extend_from_slice(encrypted.as_ref());
    Ok((header, context))
}

/// Derive reusable legacy RC4 material from a password and salt.
pub fn context(password: &str, salt: &[u8; 16]) -> Result<Context, Error> {
    let units = password.encode_utf16().count();
    if units > MAX_PASSWORD_UNITS {
        return Err(Error::PasswordTooLong { units });
    }

    let password_bytes = Zeroizing::new(
        password
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>(),
    );
    let initial_hash = Zeroizing::new(<[u8; 16]>::from(Md5::digest(password_bytes.as_slice())));
    let mut intermediate = Zeroizing::new([0u8; 336]);
    for chunk in intermediate.as_chunks_mut::<21>().0.iter_mut() {
        chunk[..5].copy_from_slice(&initial_hash[..5]);
        chunk[5..].copy_from_slice(salt);
    }
    let final_hash = Zeroizing::new(<[u8; 16]>::from(Md5::digest(intermediate.as_slice())));
    let mut secret = Zeroizing::new([0u8; 5]);
    secret.copy_from_slice(&final_hash[..5]);
    Ok(Context { secret })
}

/// Parse one complete 52-byte version 1.1 legacy RC4 header.
pub fn parse_header(data: &[u8]) -> Result<Header, Error> {
    if data.len() != HEADER_LEN {
        return Err(Error::Malformed(format!(
            "legacy RC4 header must contain exactly {HEADER_LEN} bytes, found {}",
            data.len()
        )));
    }
    let major = u16::from_le_bytes([data[0], data[1]]);
    let minor = u16::from_le_bytes([data[2], data[3]]);
    if (major, minor) != (1, 1) {
        return Err(Error::UnsupportedVersion { major, minor });
    }
    let mut salt = [0u8; 16];
    let mut encrypted_verifier = [0u8; 16];
    let mut encrypted_verifier_hash = [0u8; 16];
    salt.copy_from_slice(&data[4..20]);
    encrypted_verifier.copy_from_slice(&data[20..36]);
    encrypted_verifier_hash.copy_from_slice(&data[36..52]);
    Ok(Header {
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
    })
}

/// Verify a password and return its context only on a constant-time match.
pub fn verify(header: &Header, password: &str) -> Result<Option<Context>, Error> {
    let context = context(password, &header.salt)?;
    let key = derive_block_key(&context, 0);
    let mut cipher = Rc4::new_from_slice(key.as_slice())
        .map_err(|_| Error::Malformed("invalid legacy RC4 key length".to_string()))?;
    let mut verifier = Zeroizing::new(header.encrypted_verifier);
    let mut verifier_hash = Zeroizing::new(header.encrypted_verifier_hash);
    cipher.apply_keystream(verifier.as_mut());
    cipher.apply_keystream(verifier_hash.as_mut());
    let calculated = Zeroizing::new(<[u8; 16]>::from(Md5::digest(verifier.as_slice())));
    Ok(bool::from(calculated.ct_eq(verifier_hash.as_ref())).then_some(context))
}

/// Apply the legacy RC4 stream from the start of a format stream.
///
/// `block_size` is supplied by the format consumer: DOC uses 512 and XLS
/// uses 1024.  The complete requested range is checked before any bytes are
/// changed, so a caller-bound or offset failure cannot leave a partial edit.
pub fn apply(context: &Context, block_size: usize, data: &mut [u8]) -> Result<(), Error> {
    apply_at(context, block_size, 0, data)
}

/// Apply the legacy RC4 stream at an absolute byte offset.
///
/// `block_size` is supplied by the format consumer: DOC uses 512 and XLS
/// uses 1024.  The complete requested range is checked before any bytes are
/// changed, so a caller-bound or offset failure cannot leave a partial edit.
pub fn apply_at(
    context: &Context,
    block_size: usize,
    absolute_offset: usize,
    data: &mut [u8],
) -> Result<(), Error> {
    if block_size == 0 || block_size > MAX_BLOCK_SIZE {
        return Err(Error::InvalidBlockSize { size: block_size });
    }
    let end = absolute_offset
        .checked_add(data.len())
        .ok_or(Error::StreamRangeOverflow)?;
    if let Some(last_byte) = end.checked_sub(1) {
        u32::try_from(last_byte / block_size).map_err(|_| Error::StreamRangeOverflow)?;
    }

    let mut position = absolute_offset;
    let mut remaining = data;
    while !remaining.is_empty() {
        let block = u32::try_from(position / block_size).map_err(|_| Error::StreamRangeOverflow)?;
        let block_offset = position % block_size;
        let count = remaining.len().min(block_size - block_offset);
        let key = derive_block_key(context, block);
        let mut cipher = Rc4::new_from_slice(key.as_slice())
            .map_err(|_| Error::Malformed("invalid legacy RC4 key length".to_string()))?;
        if block_offset != 0 {
            let mut discarded = Zeroizing::new([0u8; MAX_BLOCK_SIZE]);
            cipher.apply_keystream(&mut discarded[..block_offset]);
        }
        cipher.apply_keystream(&mut remaining[..count]);
        position = position
            .checked_add(count)
            .ok_or(Error::StreamRangeOverflow)?;
        remaining = &mut remaining[count..];
    }
    Ok(())
}

fn derive_block_key(context: &Context, block: u32) -> Zeroizing<[u8; 16]> {
    let mut input = Zeroizing::new([0u8; 9]);
    input[..5].copy_from_slice(context.secret.as_slice());
    input[5..].copy_from_slice(&block.to_le_bytes());
    Zeroizing::new(<[u8; 16]>::from(Md5::digest(input.as_slice())))
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        reason = "test code panics on failure; unwrap keeps assertions concise"
    )]

    use super::*;

    #[test]
    fn password_derivation_matches_apache_poi_vector() {
        let salt = [
            0x17, 0xf6, 0xd1, 0x6b, 0x09, 0xb1, 0x5f, 0x7b, 0x4c, 0x9d, 0x03, 0xb4, 0x81, 0xb5,
            0xb4, 0x4a,
        ];
        let context = context("MoneyForNothing", &salt).unwrap();
        assert_eq!(context.secret.as_ref(), &[0xc2, 0xd9, 0x56, 0xb2, 0x6b]);
    }

    #[test]
    fn header_verifier_uses_one_continuous_cipher_and_rejects_wrong_password() {
        let salt = [0x31; 16];
        let verifier = [0x72; 16];
        let (bytes, _context) =
            build_header("correct horse battery staple", &salt, &verifier).unwrap();
        let header = parse_header(&bytes).unwrap();
        assert!(
            verify(&header, "correct horse battery staple")
                .unwrap()
                .is_some()
        );
        assert!(verify(&header, "wrong password").unwrap().is_none());
    }

    #[test]
    fn parser_rejects_truncation_extension_and_wrong_version() {
        let (valid, _) = build_header("password", &[0x11; 16], &[0x22; 16]).unwrap();
        for length in [0, HEADER_LEN - 1] {
            assert!(matches!(
                parse_header(&valid[..length]),
                Err(Error::Malformed(_))
            ));
        }
        let mut extended = valid.clone();
        extended.push(0);
        assert!(matches!(parse_header(&extended), Err(Error::Malformed(_))));
        let mut wrong_version = valid;
        wrong_version[..2].copy_from_slice(&2u16.to_le_bytes());
        assert!(matches!(
            parse_header(&wrong_version),
            Err(Error::UnsupportedVersion { major: 2, minor: 1 })
        ));
    }

    #[test]
    fn absolute_offsets_rekey_at_boundaries_for_both_consumers() {
        for block_size in [512, 1024] {
            let (_, context) = build_header("block-boundary", &[0x42; 16], &[0x24; 16]).unwrap();
            let original = vec![0xa5; block_size * 2 + 37];
            let mut encrypted = original.clone();
            apply_at(&context, block_size, block_size - 1, &mut encrypted).unwrap();
            assert_ne!(encrypted, original);
            apply_at(&context, block_size, block_size - 1, &mut encrypted).unwrap();
            assert_eq!(encrypted, original);
        }
    }

    #[test]
    fn invalid_bounds_fail_before_mutation() {
        let (_, context) = build_header("bounds", &[0x42; 16], &[0x24; 16]).unwrap();
        let mut bytes = vec![0x5a; 8];
        let original = bytes.clone();
        assert!(matches!(
            apply_at(&context, 0, 0, &mut bytes),
            Err(Error::InvalidBlockSize { size: 0 })
        ));
        assert_eq!(bytes, original);
        assert!(matches!(
            apply_at(&context, MAX_BLOCK_SIZE + 1, 0, &mut bytes),
            Err(Error::InvalidBlockSize { .. })
        ));
        assert_eq!(bytes, original);
        assert!(matches!(
            apply_at(&context, 512, usize::MAX, &mut bytes),
            Err(Error::StreamRangeOverflow)
        ));
        assert_eq!(bytes, original);
    }

    #[test]
    fn passwords_over_the_normative_limit_are_rejected() {
        let password = "a".repeat(MAX_PASSWORD_UNITS + 1);
        assert!(matches!(
            context(&password, &[0; 16]),
            Err(Error::PasswordTooLong { units }) if units == MAX_PASSWORD_UNITS + 1
        ));
    }
}
