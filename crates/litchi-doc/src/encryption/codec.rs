//! Binary codecs for legacy Word encryption headers and streams.

use super::model::DocEncryptionProfile;
use crate::package::{EncryptionKind, Error as PackageError, Result};
use crate::parts::fib::FileInformationBlock;
use encoding_rs::{
    BIG5, EUC_KR, Encoding, GBK, SHIFT_JIS, WINDOWS_874, WINDOWS_1250, WINDOWS_1251, WINDOWS_1252,
    WINDOWS_1253, WINDOWS_1254, WINDOWS_1255, WINDOWS_1256, WINDOWS_1257, WINDOWS_1258,
};
use litchi_crypto::rc4 as office_rc4;
use litchi_crypto::rc4::{Context, Error, Flags};
use md5::{Digest, Md5};
use rand::{TryRng, rngs::SysRng};
use rc4::{KeyInit, Rc4, StreamCipher};
use subtle::ConstantTimeEq;
use zeroize::Zeroizing;

pub(super) const FIB_BASE_LEN: usize = 68;
const BINARY_RC4_HEADER_LEN: usize = 52;
const BINARY_RC4_BLOCK_SIZE: usize = 512;
const CRYPTO_API_VERIFIER_LEN: usize = 60;
const CRYPTO_API_PROVIDER: &str = "Microsoft Enhanced Cryptographic Provider v1.0";

impl DocEncryptionProfile {
    pub(crate) fn table_header_len(self) -> std::result::Result<usize, String> {
        self.validate()?;
        match self {
            Self::WordXorObfuscation => Ok(0),
            Self::OfficeBinaryRc4 => Ok(BINARY_RC4_HEADER_LEN),
            Self::CryptoApiRc4 { .. } => {
                let provider_len = CRYPTO_API_PROVIDER.encode_utf16().count() + 1;
                12usize
                    .checked_add(32)
                    .and_then(|len| len.checked_add(provider_len * 2))
                    .and_then(|len| len.checked_add(CRYPTO_API_VERIFIER_LEN))
                    .ok_or_else(|| "DOC CryptoAPI encryption header length overflow".to_string())
            },
        }
    }
}

pub(crate) fn validate_writer_password(
    profile: DocEncryptionProfile,
    password: &str,
) -> std::result::Result<(), String> {
    profile.validate()?;
    let units = password.encode_utf16().count();
    if units == 0 {
        return Err("DOC password-to-open password must not be empty".to_string());
    }
    let maximum = match profile {
        DocEncryptionProfile::WordXorObfuscation => 15,
        DocEncryptionProfile::OfficeBinaryRc4 | DocEncryptionProfile::CryptoApiRc4 { .. } => 255,
    };
    if units > maximum {
        return Err(format!(
            "DOC password contains {units} UTF-16 code units and exceeds the {maximum}-unit limit"
        ));
    }
    Ok(())
}

pub(crate) fn encrypt_document_streams_for_write(
    profile: DocEncryptionProfile,
    password: &str,
    word_document: &mut [u8],
    table_stream: &mut [u8],
    data_stream: &mut [u8],
) -> std::result::Result<(), String> {
    validate_writer_password(profile, password)?;
    let header_len = profile.table_header_len()?;
    if word_document.len() < FIB_BASE_LEN {
        return Err(format!(
            "DOC WordDocument stream is shorter than the {FIB_BASE_LEN}-byte clear FIB prefix"
        ));
    }
    if table_stream.len() < header_len {
        return Err("DOC table stream is shorter than its encryption header".to_string());
    }
    if table_stream[..header_len].iter().any(|byte| *byte != 0) {
        return Err("DOC reserved table encryption header is not clear".to_string());
    }
    let flags = u16::from_le_bytes([word_document[10], word_document[11]]);
    if flags & 0x0200 == 0 {
        return Err("DOC encrypted writer output must identify the 1Table stream".to_string());
    }

    match profile {
        DocEncryptionProfile::WordXorObfuscation => {
            let password_bytes = ansi_password_bytes(password, 0x0409);
            let verifier = xor_password_verifier(&password_bytes);
            patch_fib_encryption(word_document, true, verifier);
            let context = XorContext {
                array: Zeroizing::new(create_word_xor_array(&password_bytes)),
            };
            apply_xor_stream(&mut word_document[FIB_BASE_LEN..], FIB_BASE_LEN, &context)
                .map_err(|error| error.to_string())?;
            apply_xor_stream(table_stream, 0, &context).map_err(|error| error.to_string())?;
            apply_xor_stream(data_stream, 0, &context).map_err(|error| error.to_string())?;
        },
        DocEncryptionProfile::OfficeBinaryRc4 => {
            let mut salt = Zeroizing::new([0u8; 16]);
            let mut verifier = Zeroizing::new([0u8; 16]);
            fill_random(salt.as_mut(), "binary RC4 salt")?;
            fill_random(verifier.as_mut(), "binary RC4 verifier")?;
            let (header, secret) = build_binary_rc4_header(password, &salt, &verifier)?;
            table_stream[..header_len].copy_from_slice(&header);
            patch_fib_encryption(word_document, false, header_len as u32);
            apply_stream_cipher(&mut word_document[FIB_BASE_LEN..], FIB_BASE_LEN, &secret)
                .map_err(|error| error.to_string())?;
            apply_stream_cipher(&mut table_stream[header_len..], header_len, &secret)
                .map_err(|error| error.to_string())?;
            apply_stream_cipher(data_stream, 0, &secret).map_err(|error| error.to_string())?;
        },
        DocEncryptionProfile::CryptoApiRc4 { key_bits } => {
            let mut salt = Zeroizing::new([0u8; 16]);
            let mut verifier = Zeroizing::new([0u8; 16]);
            fill_random(salt.as_mut(), "CryptoAPI salt")?;
            fill_random(verifier.as_mut(), "CryptoAPI verifier")?;
            let (header, context) = build_cryptoapi_header(password, key_bits, &salt, &verifier)?;
            if header.len() != header_len {
                return Err("DOC CryptoAPI header length is inconsistent".to_string());
            }
            table_stream[..header_len].copy_from_slice(&header);
            patch_fib_encryption(word_document, false, header_len as u32);
            apply_cryptoapi_stream(&mut word_document[FIB_BASE_LEN..], FIB_BASE_LEN, &context)
                .map_err(|error| error.to_string())?;
            apply_cryptoapi_stream(&mut table_stream[header_len..], header_len, &context)
                .map_err(|error| error.to_string())?;
            apply_cryptoapi_stream(data_stream, 0, &context).map_err(|error| error.to_string())?;
        },
    }
    Ok(())
}

fn patch_fib_encryption(word_document: &mut [u8], obfuscated: bool, l_key: u32) {
    let mut flags = u16::from_le_bytes([word_document[10], word_document[11]]) | 0x0100;
    if obfuscated {
        flags |= 0x8000;
    } else {
        flags &= !0x8000;
    }
    word_document[10..12].copy_from_slice(&flags.to_le_bytes());
    word_document[14..18].copy_from_slice(&l_key.to_le_bytes());
}

fn fill_random(bytes: &mut [u8], field: &str) -> std::result::Result<(), String> {
    SysRng
        .try_fill_bytes(bytes)
        .map_err(|_| format!("operating-system randomness unavailable for DOC {field}"))
}

fn build_binary_rc4_header(
    password: &str,
    salt: &[u8; 16],
    verifier: &[u8; 16],
) -> std::result::Result<(Vec<u8>, Zeroizing<[u8; 5]>), String> {
    let secret = derive_secret(password, salt);
    let key = derive_block_key(&secret, 0);
    let mut encrypted = Zeroizing::new([0u8; 32]);
    encrypted[..16].copy_from_slice(verifier);
    encrypted[16..].copy_from_slice(&Md5::digest(verifier));
    let mut cipher = Rc4::new_from_slice(key.as_ref())
        .map_err(|_| "invalid DOC binary RC4 key length".to_string())?;
    cipher.apply_keystream(encrypted.as_mut());
    let mut header = Vec::with_capacity(BINARY_RC4_HEADER_LEN);
    header.extend_from_slice(&1u16.to_le_bytes());
    header.extend_from_slice(&1u16.to_le_bytes());
    header.extend_from_slice(salt);
    header.extend_from_slice(encrypted.as_ref());
    Ok((header, secret))
}

fn build_cryptoapi_header(
    password: &str,
    key_bits: u16,
    salt: &[u8; 16],
    verifier: &[u8; 16],
) -> std::result::Result<(Vec<u8>, Context), String> {
    office_rc4::build_header(
        password,
        usize::from(key_bits),
        Flags::CRYPTO_API,
        CRYPTO_API_PROVIDER,
        salt,
        verifier,
    )
    .map_err(|error| map_crypto_error(error).to_string())
}

const XOR_INITIAL_CODE: [u16; 15] = [
    0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9,
    0x280c, 0xa96a, 0x4ec3,
];

const XOR_PAD_ARRAY: [u8; 16] = [
    0xbb, 0xff, 0xff, 0xba, 0xff, 0xff, 0xb9, 0x80, 0x00, 0xbe, 0x0f, 0x00, 0xbf, 0x0f, 0x00, 0x00,
];

const XOR_MATRIX: [[u16; 7]; 15] = [
    [0xaefc, 0x4dd9, 0x9bb2, 0x2745, 0x4e8a, 0x9d14, 0x2a09],
    [0x7b61, 0xf6c2, 0xfda5, 0xeb6b, 0xc6f7, 0x9dcf, 0x2bbf],
    [0x4563, 0x8ac6, 0x05ad, 0x0b5a, 0x16b4, 0x2d68, 0x5ad0],
    [0x0375, 0x06ea, 0x0dd4, 0x1ba8, 0x3750, 0x6ea0, 0xdd40],
    [0xd849, 0xa0b3, 0x5147, 0xa28e, 0x553d, 0xaa7a, 0x44d5],
    [0x6f45, 0xde8a, 0xad35, 0x4a4b, 0x9496, 0x390d, 0x721a],
    [0xeb23, 0xc667, 0x9cef, 0x29ff, 0x53fe, 0xa7fc, 0x5fd9],
    [0x47d3, 0x8fa6, 0x0f6d, 0x1eda, 0x3db4, 0x7b68, 0xf6d0],
    [0xb861, 0x60e3, 0xc1c6, 0x93ad, 0x377b, 0x6ef6, 0xddec],
    [0x45a0, 0x8b40, 0x06a1, 0x0d42, 0x1a84, 0x3508, 0x6a10],
    [0xaa51, 0x4483, 0x8906, 0x022d, 0x045a, 0x08b4, 0x1168],
    [0x76b4, 0xed68, 0xcaf1, 0x85c3, 0x1ba7, 0x374e, 0x6e9c],
    [0x3730, 0x6e60, 0xdcc0, 0xa9a1, 0x4363, 0x86c6, 0x1dad],
    [0x3331, 0x6662, 0xccc4, 0x89a9, 0x0373, 0x06e6, 0x0dcc],
    [0x1021, 0x2042, 0x4084, 0x8108, 0x1231, 0x2462, 0x48c4],
];

#[derive(Debug, Clone, Copy)]
struct BinaryRc4Header {
    salt: [u8; 16],
    encrypted_verifier: [u8; 16],
    encrypted_verifier_hash: [u8; 16],
}

pub(super) struct XorContext {
    pub(super) array: Zeroizing<[u8; 16]>,
}

pub(crate) fn decrypt_document_streams(
    fib: &FileInformationBlock,
    word_document: &mut [u8],
    table_stream: &mut [u8],
    data_stream: Option<&mut [u8]>,
    password: Option<&str>,
) -> Result<()> {
    if fib.is_obfuscated() {
        if word_document.len() < FIB_BASE_LEN {
            return Err(PackageError::Corrupted(format!(
                "obfuscated WordDocument stream is shorter than the {FIB_BASE_LEN}-byte clear FIB base"
            )));
        }
        let password = password.ok_or(PackageError::PasswordRequired)?;
        let context =
            verify_xor_password(fib.xor_obfuscation_verifier(), password, fib.language_id())
                .ok_or(PackageError::InvalidPassword)?;
        apply_xor_stream(&mut word_document[FIB_BASE_LEN..], FIB_BASE_LEN, &context)?;
        apply_xor_stream(table_stream, 0, &context)?;
        if let Some(data_stream) = data_stream {
            apply_xor_stream(data_stream, 0, &context)?;
        }
        return Ok(());
    }

    let header_len = usize::try_from(fib.encryption_header_size()).map_err(|_| {
        PackageError::MalformedEncryptionHeader(
            "encryption header size does not fit in memory".to_string(),
        )
    })?;
    if header_len > table_stream.len() {
        return Err(PackageError::MalformedEncryptionHeader(format!(
            "encryption header is {header_len} bytes but the table stream is only {} bytes",
            table_stream.len()
        )));
    }
    if header_len < 4 {
        return Err(PackageError::MalformedEncryptionHeader(
            "encryption header is missing its version".to_string(),
        ));
    }

    let header = &table_stream[..header_len];
    let major = u16::from_le_bytes([header[0], header[1]]);
    let minor = u16::from_le_bytes([header[2], header[3]]);
    if matches!((major, minor), (2..=4, 2)) {
        let password = password.ok_or(PackageError::PasswordRequired)?;
        let parsed = office_rc4::parse_header(header).map_err(map_crypto_error)?;
        let context = office_rc4::verify(&parsed, password)
            .map_err(map_crypto_error)?
            .ok_or(PackageError::InvalidPassword)?;
        apply_cryptoapi_stream(&mut word_document[FIB_BASE_LEN..], FIB_BASE_LEN, &context)?;
        apply_cryptoapi_stream(&mut table_stream[header_len..], header_len, &context)?;
        if let Some(data_stream) = data_stream {
            apply_cryptoapi_stream(data_stream, 0, &context)?;
        }
        return Ok(());
    }
    if (major, minor) != (1, 1) {
        let kind = if matches!((major, minor), (2..=4, 2)) {
            EncryptionKind::CryptoApi
        } else {
            EncryptionKind::Unknown { major, minor }
        };
        return Err(PackageError::UnsupportedEncryption(kind));
    }
    if header_len != BINARY_RC4_HEADER_LEN {
        return Err(PackageError::MalformedEncryptionHeader(format!(
            "binary RC4 encryption header must contain exactly {BINARY_RC4_HEADER_LEN} bytes, found {header_len}"
        )));
    }
    if word_document.len() < FIB_BASE_LEN {
        return Err(PackageError::Corrupted(format!(
            "encrypted WordDocument stream is shorter than the {FIB_BASE_LEN}-byte clear FIB base"
        )));
    }

    let mut salt = [0u8; 16];
    let mut encrypted_verifier = [0u8; 16];
    let mut encrypted_verifier_hash = [0u8; 16];
    salt.copy_from_slice(&header[4..20]);
    encrypted_verifier.copy_from_slice(&header[20..36]);
    encrypted_verifier_hash.copy_from_slice(&header[36..52]);
    let header = BinaryRc4Header {
        salt,
        encrypted_verifier,
        encrypted_verifier_hash,
    };

    let password = password.ok_or(PackageError::PasswordRequired)?;
    let secret = verify(&header, password)?.ok_or(PackageError::InvalidPassword)?;

    apply_stream_cipher(&mut word_document[FIB_BASE_LEN..], FIB_BASE_LEN, &secret)?;
    apply_stream_cipher(&mut table_stream[header_len..], header_len, &secret)?;
    if let Some(data_stream) = data_stream {
        apply_stream_cipher(data_stream, 0, &secret)?;
    }
    Ok(())
}

fn verify_xor_password(stored: u32, password: &str, language_id: u16) -> Option<XorContext> {
    let canonical = xor_password_bytes(password);
    let ansi = ansi_password_bytes(password, language_id);
    let mut candidates = vec![canonical];
    if candidates[0].as_slice() != ansi.as_slice() {
        candidates.push(ansi);
    }

    for candidate in candidates {
        let derived = xor_password_verifier(&candidate);
        if bool::from(derived.to_le_bytes().ct_eq(&stored.to_le_bytes())) {
            return Some(XorContext {
                array: Zeroizing::new(create_word_xor_array(&candidate)),
            });
        }
    }
    None
}

pub(super) fn xor_password_bytes(password: &str) -> Zeroizing<Vec<u8>> {
    Zeroizing::new(
        password
            .encode_utf16()
            .take(15)
            .map(|unit| {
                let low = unit as u8;
                if low == 0 { (unit >> 8) as u8 } else { low }
            })
            .collect(),
    )
}

pub(super) fn ansi_password_bytes(password: &str, language_id: u16) -> Zeroizing<Vec<u8>> {
    let encoding = ansi_encoding_for_lcid(language_id);
    let mut bytes = Vec::with_capacity(15);
    for character in password.chars() {
        let text = character.to_string();
        let (encoded, _, had_errors) = encoding.encode(&text);
        if had_errors {
            bytes.push(b'?');
        } else {
            bytes.extend(encoded.iter().copied().take(15 - bytes.len()));
        }
        if bytes.len() == 15 {
            break;
        }
    }
    Zeroizing::new(bytes)
}

fn ansi_encoding_for_lcid(language_id: u16) -> &'static Encoding {
    let primary = language_id & 0x03ff;
    match primary {
        0x01 | 0x20 | 0x29 => WINDOWS_1256,
        0x02 | 0x19 | 0x22 | 0x23 | 0x28 | 0x2f | 0x3f | 0x40 | 0x44 | 0x50 => WINDOWS_1251,
        0x04 => match language_id {
            0x0804 | 0x1004 => GBK,
            _ => BIG5,
        },
        0x05 | 0x0e | 0x15 | 0x18 | 0x1b | 0x1c | 0x24 | 0x2e => WINDOWS_1250,
        0x08 => WINDOWS_1253,
        0x0d => WINDOWS_1255,
        0x11 => SHIFT_JIS,
        0x12 => EUC_KR,
        0x1a => match language_id {
            0x0c1a | 0x201a => WINDOWS_1251,
            _ => WINDOWS_1250,
        },
        0x1e => WINDOWS_874,
        0x1f => WINDOWS_1254,
        0x25..=0x27 => WINDOWS_1257,
        0x2a => WINDOWS_1258,
        0x2c => match language_id {
            0x082c => WINDOWS_1251,
            _ => WINDOWS_1254,
        },
        0x43 => match language_id {
            0x0843 => WINDOWS_1251,
            _ => WINDOWS_1254,
        },
        _ => WINDOWS_1252,
    }
}

pub(super) fn xor_password_verifier(password: &[u8]) -> u32 {
    (u32::from(create_xor_key(password)) << 16) | u32::from(create_xor_hash(password))
}

fn create_xor_key(password: &[u8]) -> u16 {
    if password.is_empty() {
        return 0;
    }
    let mut key = XOR_INITIAL_CODE[password.len() - 1];
    for (matrix_line, mut byte) in (15 - password.len()..).zip(password.iter().copied()) {
        for matrix_value in XOR_MATRIX[matrix_line] {
            if byte & 1 != 0 {
                key ^= matrix_value;
            }
            byte >>= 1;
        }
    }
    key
}

fn create_xor_hash(password: &[u8]) -> u16 {
    if password.is_empty() {
        return 0;
    }
    let mut verifier = 0u16;
    for byte in password.iter().rev() {
        verifier = ((verifier << 1) & 0x7fff) | ((verifier >> 14) & 1);
        verifier ^= u16::from(*byte);
    }
    verifier = ((verifier << 1) & 0x7fff) | ((verifier >> 14) & 1);
    verifier ^ password.len() as u16 ^ 0xce4b
}

pub(super) fn create_word_xor_array(password: &[u8]) -> [u8; 16] {
    let mut array = [0u8; 16];
    array[..password.len()].copy_from_slice(password);
    array[password.len()..].copy_from_slice(&XOR_PAD_ARRAY[..16 - password.len()]);
    let key = create_xor_key(password).to_le_bytes();
    for (index, byte) in array.iter_mut().enumerate() {
        *byte = (*byte ^ key[index & 1]).rotate_right(1);
    }
    array
}

pub(super) fn apply_xor_stream(
    data: &mut [u8],
    absolute_offset: usize,
    context: &XorContext,
) -> Result<()> {
    for (index, byte) in data.iter_mut().enumerate() {
        let absolute = absolute_offset
            .checked_add(index)
            .ok_or_else(|| PackageError::Corrupted("DOC XOR stream offset overflow".to_string()))?;
        let transformed = *byte ^ context.array[absolute & 0x0f];
        if *byte != 0 && transformed != 0 {
            *byte = transformed;
        }
    }
    Ok(())
}

fn map_crypto_error(error: Error) -> PackageError {
    match error {
        Error::Malformed(message) => PackageError::MalformedEncryptionHeader(message),
        Error::UnsupportedVersion { major, minor } => {
            PackageError::UnsupportedEncryption(EncryptionKind::Unknown { major, minor })
        },
        Error::UnsupportedAlgorithm => {
            PackageError::UnsupportedEncryption(EncryptionKind::CryptoApi)
        },
    }
}

pub(super) fn apply_cryptoapi_stream(
    mut data: &mut [u8],
    mut absolute_offset: usize,
    context: &Context,
) -> Result<()> {
    while !data.is_empty() {
        let block = u32::try_from(absolute_offset / BINARY_RC4_BLOCK_SIZE).map_err(|_| {
            PackageError::Corrupted(
                "encrypted DOC stream is too large for CryptoAPI RC4".to_string(),
            )
        })?;
        let block_offset = absolute_offset % BINARY_RC4_BLOCK_SIZE;
        let count = data
            .len()
            .min(BINARY_RC4_BLOCK_SIZE.saturating_sub(block_offset));
        office_rc4::apply_at(context, block, block_offset, &mut data[..count])
            .map_err(map_crypto_error)?;
        absolute_offset += count;
        data = &mut data[count..];
    }
    Ok(())
}

pub(super) fn derive_secret(password: &str, salt: &[u8; 16]) -> Zeroizing<[u8; 5]> {
    let password_bytes = Zeroizing::new(
        password
            .encode_utf16()
            .take(255)
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>(),
    );
    let initial_hash = Zeroizing::new(<[u8; 16]>::from(Md5::digest(password_bytes.as_slice())));
    let mut intermediate = Zeroizing::new([0u8; 336]);
    for chunk in intermediate.chunks_exact_mut(21) {
        chunk[..5].copy_from_slice(&initial_hash[..5]);
        chunk[5..].copy_from_slice(salt);
    }
    let final_hash = Zeroizing::new(<[u8; 16]>::from(Md5::digest(intermediate.as_slice())));
    let mut secret = Zeroizing::new([0u8; 5]);
    secret.copy_from_slice(&final_hash[..5]);
    secret
}

fn derive_block_key(secret: &[u8; 5], block: u32) -> Zeroizing<[u8; 16]> {
    let mut input = Zeroizing::new([0u8; 9]);
    input[..5].copy_from_slice(secret);
    input[5..].copy_from_slice(&block.to_le_bytes());
    Zeroizing::new(<[u8; 16]>::from(Md5::digest(input.as_slice())))
}

fn verify(header: &BinaryRc4Header, password: &str) -> Result<Option<Zeroizing<[u8; 5]>>> {
    let secret = derive_secret(password, &header.salt);
    let key = derive_block_key(&secret, 0);
    let mut cipher = Rc4::new_from_slice(key.as_ref()).map_err(|_| {
        PackageError::MalformedEncryptionHeader("invalid binary RC4 key length".to_string())
    })?;
    let mut verifier = Zeroizing::new(header.encrypted_verifier);
    let mut verifier_hash = Zeroizing::new(header.encrypted_verifier_hash);
    cipher.apply_keystream(verifier.as_mut());
    cipher.apply_keystream(verifier_hash.as_mut());
    let calculated = Zeroizing::new(<[u8; 16]>::from(Md5::digest(verifier.as_slice())));
    let difference = calculated
        .iter()
        .zip(verifier_hash.iter())
        .fold(0u8, |difference, (left, right)| difference | (left ^ right));
    Ok((difference == 0).then_some(secret))
}

pub(super) fn apply_stream_cipher(
    mut data: &mut [u8],
    mut absolute_offset: usize,
    secret: &[u8; 5],
) -> Result<()> {
    while !data.is_empty() {
        let block = u32::try_from(absolute_offset / BINARY_RC4_BLOCK_SIZE).map_err(|_| {
            PackageError::Corrupted("encrypted DOC stream is too large for binary RC4".to_string())
        })?;
        let block_offset = absolute_offset % BINARY_RC4_BLOCK_SIZE;
        let key = derive_block_key(secret, block);
        let mut cipher = Rc4::new_from_slice(key.as_ref()).map_err(|_| {
            PackageError::MalformedEncryptionHeader("invalid binary RC4 key length".to_string())
        })?;
        if block_offset != 0 {
            let mut discarded = Zeroizing::new([0u8; BINARY_RC4_BLOCK_SIZE]);
            cipher.apply_keystream(&mut discarded[..block_offset]);
        }

        let count = data
            .len()
            .min(BINARY_RC4_BLOCK_SIZE.saturating_sub(block_offset));
        cipher.apply_keystream(&mut data[..count]);
        absolute_offset += count;
        data = &mut data[count..];
    }
    Ok(())
}
