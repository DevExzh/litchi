//! Format-neutral Office Binary Document RC4 `CryptoAPI` primitives.

use bitflags::bitflags;
use rc4::{KeyInit, Rc4, StreamCipher};
use sha1::{Digest, Sha1};
use std::fmt;
use subtle::ConstantTimeEq;
use zeroize::Zeroizing;

const CALG_RC4: u32 = 0x0000_6801;
const CALG_SHA1: u32 = 0x0000_8004;
const RC4_DISCARD_BUFFER_LEN: usize = 4 * 1024;

bitflags! {
    /// Flags stored in an Office Binary Document CryptoAPI header.
    #[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
    pub struct Flags: u32 {
        /// The header uses a CryptoAPI provider.
        const CRYPTO_API = 0x0000_0004;
        /// Document properties are encrypted with the document payload.
        const DOC_PROPERTIES = 0x0000_0008;
        /// The provider is external. RC4 helpers intentionally reject it.
        const EXTERNAL = 0x0000_0010;
        /// The provider uses AES. RC4 helpers intentionally reject it.
        const AES = 0x0000_0020;
    }
}

/// A malformed or unsupported RC4 `CryptoAPI` header or operation.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum Error {
    Malformed(String),
    UnsupportedVersion { major: u16, minor: u16 },
    UnsupportedAlgorithm,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Malformed(message) => write!(formatter, "malformed CryptoAPI data: {message}"),
            Self::UnsupportedVersion { major, minor } => {
                write!(formatter, "unsupported CryptoAPI version {major}.{minor}")
            },
            Self::UnsupportedAlgorithm => {
                formatter.write_str("unsupported CryptoAPI algorithm or flags")
            },
        }
    }
}

impl std::error::Error for Error {}

/// A validated RC4 `CryptoAPI` header.
pub struct Header {
    salt: [u8; 16],
    encrypted_verifier: [u8; 16],
    encrypted_verifier_hash: [u8; 20],
    key_bits: usize,
}

/// Password-derived RC4 `CryptoAPI` key material with zeroizing storage.
///
/// Contexts are intentionally move-only so secret material is not duplicated
/// accidentally. Cipher operations borrow a context for reuse.
///
/// ```compile_fail
/// let context = litchi_crypto::rc4::context("password", &[0; 16], 128)?;
/// let duplicate = context.clone();
/// # Ok::<(), litchi_crypto::rc4::Error>(())
/// ```
pub struct Context {
    secret: Zeroizing<[u8; 20]>,
    key_bits: usize,
}

struct BlockKey {
    bytes: Zeroizing<[u8; 16]>,
    len: usize,
}

impl BlockKey {
    fn as_slice(&self) -> &[u8] {
        &self.bytes[..self.len]
    }
}

/// Build a validated RC4 `CryptoAPI` header and its reusable cipher context.
///
/// # Errors
///
/// Returns [`Error::UnsupportedAlgorithm`] when `flags` request an external
/// or AES provider, or when `key_bits` is outside the validated RC4 range.
/// Returns [`Error::Malformed`] when the provider name is empty, contains a
/// NUL, or a length computation overflows.
pub fn build_header(
    password: &str,
    key_bits: usize,
    flags: Flags,
    provider: &str,
    salt: &[u8; 16],
    verifier: &[u8; 16],
) -> Result<(Vec<u8>, Context), Error> {
    validate_flags(flags)?;
    if provider.is_empty() || provider.encode_utf16().any(|unit| unit == 0) {
        return malformed("CryptoAPI provider name is empty or contains NUL");
    }
    let context = context(password, salt, key_bits)?;
    let key_bits_field = u32::try_from(key_bits).map_err(|_err| Error::UnsupportedAlgorithm)?;
    let mut encrypted = Zeroizing::new([0u8; 36]);
    encrypted[..16].copy_from_slice(verifier);
    encrypted[16..].copy_from_slice(&Sha1::digest(verifier));
    apply(&context, 0, encrypted.as_mut())?;

    let provider_len = provider
        .encode_utf16()
        .count()
        .checked_add(1)
        .and_then(|units| units.checked_mul(2))
        .ok_or_else(|| Error::Malformed("CryptoAPI provider length overflow".to_string()))?;
    let header_size = 32usize
        .checked_add(provider_len)
        .and_then(|len| u32::try_from(len).ok())
        .ok_or_else(|| Error::Malformed("CryptoAPI header size overflow".to_string()))?;
    let capacity = 12usize
        .checked_add(header_size as usize)
        .and_then(|len| len.checked_add(60))
        .ok_or_else(|| Error::Malformed("CryptoAPI header length overflow".to_string()))?;
    let mut header = Vec::with_capacity(capacity);
    header.extend_from_slice(&2u16.to_le_bytes());
    header.extend_from_slice(&2u16.to_le_bytes());
    header.extend_from_slice(&flags.bits().to_le_bytes());
    header.extend_from_slice(&header_size.to_le_bytes());
    header.extend_from_slice(&flags.bits().to_le_bytes());
    header.extend_from_slice(&0u32.to_le_bytes());
    header.extend_from_slice(&CALG_RC4.to_le_bytes());
    header.extend_from_slice(&CALG_SHA1.to_le_bytes());
    header.extend_from_slice(&key_bits_field.to_le_bytes());
    header.extend_from_slice(&1u32.to_le_bytes());
    header.extend_from_slice(&0u32.to_le_bytes());
    header.extend_from_slice(&0u32.to_le_bytes());
    for unit in provider.encode_utf16().chain(std::iter::once(0)) {
        header.extend_from_slice(&unit.to_le_bytes());
    }
    header.extend_from_slice(&16u32.to_le_bytes());
    header.extend_from_slice(salt);
    header.extend_from_slice(&encrypted[..16]);
    header.extend_from_slice(&20u32.to_le_bytes());
    header.extend_from_slice(&encrypted[16..]);
    Ok((header, context))
}

/// Derive a reusable cipher context from a password, salt, and key size.
///
/// # Errors
///
/// Returns [`Error::UnsupportedAlgorithm`] when `key_bits` is not a multiple
/// of 8 within the supported 40..=128 range.
pub fn context(password: &str, salt: &[u8; 16], key_bits: usize) -> Result<Context, Error> {
    if !(40..=128).contains(&key_bits) || !key_bits.is_multiple_of(8) {
        return Err(Error::UnsupportedAlgorithm);
    }
    let mut hasher = Sha1::new();
    hasher.update(salt);
    for unit in password.encode_utf16().take(255) {
        let bytes = Zeroizing::new(unit.to_le_bytes());
        hasher.update(bytes.as_slice());
    }
    Ok(Context {
        secret: Zeroizing::new(<[u8; 20]>::from(hasher.finalize())),
        key_bits,
    })
}

/// Parse and validate one complete RC4 `CryptoAPI` header.
///
/// # Errors
///
/// Returns [`Error::UnsupportedVersion`] for an unrecognized version pair,
/// [`Error::UnsupportedAlgorithm`] for non-RC4/SHA1 algorithm identifiers or
/// invalid key sizes, and [`Error::Malformed`] when fixed fields, lengths, or
/// the verifier layout do not match the RC4 `CryptoAPI` grammar.
pub fn parse_header(data: &[u8]) -> Result<Header, Error> {
    if data.len() < 12 {
        return malformed("CryptoAPI header is truncated");
    }
    let major = le_u16(data, 0)?;
    let minor = le_u16(data, 2)?;
    if !(2..=4).contains(&major) || minor != 2 {
        return Err(Error::UnsupportedVersion { major, minor });
    }
    let outer_flags = Flags::from_bits_retain(le_u32(data, 4)?);
    validate_flags(outer_flags)?;
    let header_size = usize::try_from(le_u32(data, 8)?).map_err(|_err| {
        Error::Malformed("CryptoAPI header size does not fit in memory".to_string())
    })?;
    if header_size < 32 {
        return malformed("CryptoAPI encryption header is shorter than 32 bytes");
    }
    let body = checked_slice(data, 12, header_size, "CryptoAPI encryption header")?;
    let header_flags = Flags::from_bits_retain(le_u32(body, 0)?);
    validate_flags(header_flags)?;
    if header_flags != outer_flags {
        return malformed("CryptoAPI header flags do not match their outer copy");
    }
    if le_u32(body, 4)? != 0 {
        return malformed("CryptoAPI SizeExtra field is not zero");
    }
    if le_u32(body, 8)? != CALG_RC4 {
        return Err(Error::UnsupportedAlgorithm);
    }
    if le_u32(body, 12)? != CALG_SHA1 {
        return Err(Error::UnsupportedAlgorithm);
    }
    let raw_key_bits = le_u32(body, 16)?;
    let key_bits = if raw_key_bits == 0 { 40 } else { raw_key_bits };
    if !(40..=128).contains(&key_bits) || key_bits % 8 != 0 || le_u32(body, 20)? != 1 {
        return Err(Error::UnsupportedAlgorithm);
    }
    // Reserved1 is undefined and intentionally ignored. Reserved2 is required
    // to be zero by the RC4 CryptoAPI Encryption Header grammar.
    if le_u32(body, 28)? != 0 {
        return malformed("CryptoAPI Reserved2 field is not zero");
    }
    let csp = &body[32..];
    if !csp.is_empty() && (csp.len() % 2 != 0 || !csp.ends_with(&[0, 0])) {
        return malformed("CryptoAPI provider name is not a terminated UTF-16LE string");
    }
    let verifier_offset = 12usize
        .checked_add(header_size)
        .ok_or_else(|| Error::Malformed("CryptoAPI verifier offset overflow".to_string()))?;
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
    Ok(Header {
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
        key_bits: key_bits as usize,
    })
}

/// Verify a password and return its context only on a constant-time match.
///
/// # Errors
///
/// Returns [`Error::UnsupportedAlgorithm`] when the header's key size is
/// invalid, and [`Error::Malformed`] when the derived RC4 key length is
/// rejected by the cipher.
pub fn verify(header: &Header, password: &str) -> Result<Option<Context>, Error> {
    let context = context(password, &header.salt, header.key_bits)?;
    let key = derive_block_key(&context, 0);
    let mut cipher = Rc4::new_from_slice(key.as_slice())
        .map_err(|_err| Error::Malformed("invalid CryptoAPI RC4 key length".to_string()))?;
    let mut verifier = Zeroizing::new(header.encrypted_verifier);
    let mut verifier_hash = Zeroizing::new(header.encrypted_verifier_hash);
    cipher.apply_keystream(verifier.as_mut());
    cipher.apply_keystream(verifier_hash.as_mut());
    let calculated = Zeroizing::new(<[u8; 20]>::from(Sha1::digest(verifier.as_slice())));
    Ok(bool::from(calculated.ct_eq(verifier_hash.as_ref())).then_some(context))
}

/// Apply a record-block RC4 keystream in place.
///
/// # Errors
///
/// Returns [`Error::Malformed`] when the derived RC4 key length is rejected
/// by the cipher.
pub fn apply(context: &Context, block: u32, data: &mut [u8]) -> Result<(), Error> {
    apply_at(context, block, 0, data)
}

/// Apply a record-block RC4 keystream after skipping `offset` bytes.
///
/// # Errors
///
/// Returns [`Error::Malformed`] when the derived RC4 key length is rejected
/// by the cipher.
pub fn apply_at(
    context: &Context,
    block: u32,
    offset: usize,
    data: &mut [u8],
) -> Result<(), Error> {
    let key = derive_block_key(context, block);
    let mut cipher = Rc4::new_from_slice(key.as_slice())
        .map_err(|_err| Error::Malformed("invalid CryptoAPI RC4 key length".to_string()))?;
    if offset != 0 {
        let mut discarded = Zeroizing::new([0u8; RC4_DISCARD_BUFFER_LEN]);
        let mut remaining = offset;
        while remaining != 0 {
            let chunk_len = remaining.min(discarded.len());
            cipher.apply_keystream(&mut discarded[..chunk_len]);
            remaining -= chunk_len;
        }
    }
    cipher.apply_keystream(data);
    Ok(())
}

fn derive_block_key(context: &Context, block: u32) -> BlockKey {
    let mut hasher = Sha1::new();
    hasher.update(context.secret.as_slice());
    hasher.update(block.to_le_bytes());
    let digest = Zeroizing::new(<[u8; 20]>::from(hasher.finalize()));
    let key_len = context.key_bits / 8;
    let output_len = if context.key_bits == 40 { 16 } else { key_len };
    let mut key = Zeroizing::new([0u8; 16]);
    key[..key_len].copy_from_slice(&digest[..key_len]);
    BlockKey {
        bytes: key,
        len: output_len,
    }
}

fn validate_flags(flags: Flags) -> Result<(), Error> {
    let has_unknown_bits = flags.bits() & !Flags::all().bits() != 0;
    if has_unknown_bits
        || !flags.contains(Flags::CRYPTO_API)
        || flags.intersects(Flags::EXTERNAL | Flags::AES)
    {
        return Err(Error::UnsupportedAlgorithm);
    }
    Ok(())
}

fn checked_slice<'a>(
    data: &'a [u8],
    offset: usize,
    len: usize,
    field: &str,
) -> Result<&'a [u8], Error> {
    let end = offset
        .checked_add(len)
        .ok_or_else(|| Error::Malformed(format!("{field} range overflow")))?;
    data.get(offset..end)
        .ok_or_else(|| Error::Malformed(format!("{field} is truncated")))
}

fn le_u16(data: &[u8], offset: usize) -> Result<u16, Error> {
    let bytes = checked_slice(data, offset, 2, "16-bit encryption field")?
        .try_into()
        .map_err(|_err| Error::Malformed("invalid 16-bit encryption field".to_string()))?;
    Ok(u16::from_le_bytes(bytes))
}

fn le_u32(data: &[u8], offset: usize) -> Result<u32, Error> {
    let bytes = checked_slice(data, offset, 4, "32-bit encryption field")?
        .try_into()
        .map_err(|_err| Error::Malformed("invalid 32-bit encryption field".to_string()))?;
    Ok(u32::from_le_bytes(bytes))
}

fn malformed<T>(message: &str) -> Result<T, Error> {
    Err(Error::Malformed(message.to_string()))
}

#[cfg(test)]
fn test_context(secret: [u8; 20], key_bits: usize) -> Context {
    Context {
        secret: Zeroizing::new(secret),
        key_bits,
    }
}

#[cfg(test)]
mod tests {
    #![allow(
        clippy::unwrap_used,
        reason = "test code panics on failure; unwrap keeps assertions concise"
    )]
    use super::*;

    fn assert_send_sync<T: Send + Sync>() {}

    fn crypto_api_header() -> Vec<u8> {
        let flags = Flags::CRYPTO_API.bits();
        let mut data = Vec::with_capacity(12 + 32 + 60);
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&32u32.to_le_bytes());
        data.extend_from_slice(&flags.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&CALG_RC4.to_le_bytes());
        data.extend_from_slice(&CALG_SHA1.to_le_bytes());
        data.extend_from_slice(&40u32.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0xfeed_beefu32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&16u32.to_le_bytes());
        data.extend_from_slice(&[0x11; 16]);
        data.extend_from_slice(&[0x22; 16]);
        data.extend_from_slice(&20u32.to_le_bytes());
        data.extend_from_slice(&[0x33; 20]);
        data
    }

    #[test]
    fn block_cipher_supports_office_key_sizes_and_offsets() {
        for key_bits in [40, 56, 120, 128] {
            let context = test_context([0x5a; 20], key_bits);
            let original = vec![0xa5; 73];
            let mut encrypted = original.clone();
            apply_at(&context, 7, 511, &mut encrypted).unwrap();
            assert_ne!(encrypted, original);
            apply_at(&context, 7, 511, &mut encrypted).unwrap();
            assert_eq!(encrypted, original);
        }
    }

    #[test]
    fn block_cipher_offset_matches_constant_memory_chunking() {
        let context = test_context([0x37; 20], 128);
        let mut combined = vec![0u8; RC4_DISCARD_BUFFER_LEN * 2 + 137 + 73];
        apply(&context, 19, &mut combined).unwrap();

        let offset = combined.len() - 73;
        let mut suffix = vec![0u8; 73];
        apply_at(&context, 19, offset, &mut suffix).unwrap();
        assert_eq!(suffix, combined[offset..]);
    }

    #[test]
    fn header_enforces_rc4_cryptoapi_fixed_fields() {
        let valid = crypto_api_header();
        assert!(parse_header(&valid).is_ok());

        for (offset, value) in [
            (12usize, (Flags::CRYPTO_API | Flags::DOC_PROPERTIES).bits()),
            (16, 1),
            (24, 0),
            (40, 1),
        ] {
            let mut malformed = valid.clone();
            malformed[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
            assert!(
                parse_header(&malformed).is_err(),
                "accepted field at {offset}"
            );
        }
    }

    #[test]
    fn typed_header_builder_round_trips_and_rejects_invalid_flags() {
        assert_send_sync::<Context>();
        let flags = Flags::CRYPTO_API | Flags::DOC_PROPERTIES;
        let (bytes, context) = build_header(
            "correct horse battery staple",
            128,
            flags,
            "Microsoft Enhanced Cryptographic Provider v1.0",
            &[0x31; 16],
            &[0x72; 16],
        )
        .unwrap();

        let header = parse_header(&bytes).unwrap();
        assert!(
            verify(&header, "correct horse battery staple")
                .unwrap()
                .is_some()
        );
        assert!(verify(&header, "wrong password").unwrap().is_none());

        let mut plaintext = b"typed, reusable context".to_vec();
        let original = plaintext.clone();
        apply(&context, 9, &mut plaintext).unwrap();
        assert_ne!(plaintext, original);
        apply(&context, 9, &mut plaintext).unwrap();
        assert_eq!(plaintext, original);

        for invalid in [
            Flags::empty(),
            Flags::CRYPTO_API | Flags::EXTERNAL,
            Flags::CRYPTO_API | Flags::AES,
            Flags::from_bits_retain(0x8000_0004),
        ] {
            assert!(matches!(
                build_header("password", 128, invalid, "provider", &[0; 16], &[0; 16],),
                Err(Error::UnsupportedAlgorithm)
            ));
        }
    }
}
