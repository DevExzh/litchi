//! BIFF8 password-to-open encryption handling.

use super::error::{XlsEncryptionKind, XlsError, XlsResult};
use crate::office_crypto::cryptoapi::{self, CryptoApiContext, CryptoApiError, CryptoApiHeader};
use md5::{Digest, Md5};
use rand::{TryRng, rngs::SysRng};
use rc4::{KeyInit, Rc4, StreamCipher};
use zeroize::Zeroizing;

const FILEPASS_SID: u16 = 0x002f;
const BOUNDSHEET8_SID: u16 = 0x0085;
const BINARY_RC4_FILEPASS_LEN: usize = 54;
const BINARY_RC4_BLOCK_SIZE: usize = 1024;
const CODEPAGE_SID: u16 = 0x0042;
const BOF_SID: u16 = 0x0809;
const EOF_SID: u16 = 0x000a;
const CRYPTOAPI_RC4_FLAGS: u32 = 0x0000_0004;
const CRYPTOAPI_RC4_PROVIDER: &str = "Microsoft Enhanced Cryptographic Provider v1.0";

/// Password-to-open encryption profiles supported by the BIFF8 writer.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XlsEncryptionProfile {
    /// Legacy BIFF8 XOR obfuscation. Passwords are limited to 15 ANSI characters.
    XorObfuscation,
    /// Office 97-2003 binary RC4 with its fixed 40-bit key derivation.
    OfficeBinaryRc4,
    /// CryptoAPI RC4 with a key length from 40 through 128 bits.
    CryptoApiRc4 { key_bits: u16 },
}

#[derive(Clone)]
pub(crate) struct XlsWriterEncryption {
    pub(crate) password: Zeroizing<String>,
    pub(crate) profile: XlsEncryptionProfile,
}

pub(crate) fn validate_writer_encryption(
    password: &str,
    profile: XlsEncryptionProfile,
) -> XlsResult<()> {
    if password.is_empty() {
        return Err(XlsError::InvalidData(
            "password-to-open password cannot be empty".to_string(),
        ));
    }
    match profile {
        XlsEncryptionProfile::XorObfuscation => {
            let characters = password.chars().count();
            if characters > 15 {
                return Err(XlsError::InvalidData(
                    "BIFF8 XOR passwords cannot exceed 15 ANSI characters".to_string(),
                ));
            }
            if password.chars().any(|character| u32::from(character) > 0xff) {
                return Err(XlsError::InvalidData(
                    "BIFF8 XOR passwords must contain only ANSI characters".to_string(),
                ));
            }
        },
        XlsEncryptionProfile::OfficeBinaryRc4
        | XlsEncryptionProfile::CryptoApiRc4 { .. } => {
            if password.encode_utf16().count() > 255 {
                return Err(XlsError::InvalidData(
                    "BIFF8 RC4 passwords cannot exceed 255 UTF-16 code units".to_string(),
                ));
            }
        },
    }
    if let XlsEncryptionProfile::CryptoApiRc4 { key_bits } = profile
        && (!(40..=128).contains(&key_bits) || key_bits % 8 != 0)
    {
        return Err(XlsError::InvalidData(
            "CryptoAPI RC4 key length must be a multiple of 8 from 40 through 128 bits"
                .to_string(),
        ));
    }
    Ok(())
}

const INITIAL_CODE_ARRAY: [u16; 15] = [
    0xe1f0, 0x1d0f, 0xcc9c, 0x84c0, 0x110c, 0x0e10, 0xf1ce, 0x313e, 0x1872, 0xe139, 0xd40f, 0x84f9,
    0x280c, 0xa96a, 0x4ec3,
];

const PAD_ARRAY: [u8; 15] = [
    0xbb, 0xff, 0xff, 0xba, 0xff, 0xff, 0xb9, 0x80, 0x00, 0xbe, 0x0f, 0x00, 0xbf, 0x0f, 0x00,
];

const ENCRYPTION_MATRIX: [[u16; 7]; 15] = [
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

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct XorObfuscation {
    key: u16,
    verifier: u16,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct BinaryRc4FilePass {
    salt: [u8; 16],
    encrypted_verifier: [u8; 16],
    encrypted_verifier_hash: [u8; 16],
}

enum FilePassRecord {
    Xor(XorObfuscation),
    BinaryRc4(BinaryRc4FilePass),
    CryptoApi(CryptoApiHeader),
    Unsupported(XlsEncryptionKind),
}

impl FilePassRecord {
    fn parse(data: &[u8]) -> XlsResult<Self> {
        if data.len() < 2 {
            return Err(XlsError::MalformedFilePass(
                "missing encryption type".to_string(),
            ));
        }
        let encryption_type = u16::from_le_bytes([data[0], data[1]]);
        match encryption_type {
            0 => {
                if data.len() != 6 {
                    return Err(XlsError::MalformedFilePass(format!(
                        "XOR FILEPASS must contain exactly 6 bytes, found {}",
                        data.len()
                    )));
                }
                Ok(Self::Xor(XorObfuscation {
                    key: u16::from_le_bytes([data[2], data[3]]),
                    verifier: u16::from_le_bytes([data[4], data[5]]),
                }))
            },
            1 => {
                if data.len() < 6 {
                    return Err(XlsError::MalformedFilePass(
                        "encrypted FILEPASS is missing its version".to_string(),
                    ));
                }
                let major = u16::from_le_bytes([data[2], data[3]]);
                let minor = u16::from_le_bytes([data[4], data[5]]);
                if matches!((major, minor), (2..=4, 2)) {
                    let header =
                        cryptoapi::parse_header(&data[2..]).map_err(map_cryptoapi_header_error)?;
                    return Ok(Self::CryptoApi(header));
                }
                if (major, minor) != (1, 1) {
                    return Ok(Self::Unsupported(XlsEncryptionKind::CryptoApi));
                }
                if data.len() != BINARY_RC4_FILEPASS_LEN {
                    return Err(XlsError::MalformedFilePass(format!(
                        "binary RC4 FILEPASS must contain exactly {BINARY_RC4_FILEPASS_LEN} bytes, found {}",
                        data.len()
                    )));
                }
                let mut salt = [0u8; 16];
                let mut encrypted_verifier = [0u8; 16];
                let mut encrypted_verifier_hash = [0u8; 16];
                salt.copy_from_slice(&data[6..22]);
                encrypted_verifier.copy_from_slice(&data[22..38]);
                encrypted_verifier_hash.copy_from_slice(&data[38..54]);
                Ok(Self::BinaryRc4(BinaryRc4FilePass {
                    salt,
                    encrypted_verifier,
                    encrypted_verifier_hash,
                }))
            },
            other => Ok(Self::Unsupported(XlsEncryptionKind::Unknown(other))),
        }
    }
}

enum WorkbookCipher {
    Xor([u8; 16]),
    BinaryRc4(Box<BinaryRc4Stream>),
    CryptoApi(CryptoApiContext),
}

struct BinaryRc4Stream {
    secret: Zeroizing<[u8; 5]>,
    block: Option<u32>,
    offset: usize,
    cipher: Option<Rc4>,
}

impl BinaryRc4Stream {
    fn new(secret: Zeroizing<[u8; 5]>) -> Self {
        Self {
            secret,
            block: None,
            offset: 0,
            cipher: None,
        }
    }

    fn apply_at(&mut self, mut data: &mut [u8], mut absolute: usize) -> XlsResult<()> {
        while !data.is_empty() {
            let block = u32::try_from(absolute / BINARY_RC4_BLOCK_SIZE).map_err(|_| {
                XlsError::InvalidData("Workbook stream is too large for binary RC4".to_string())
            })?;
            let block_offset = absolute % BINARY_RC4_BLOCK_SIZE;
            if self.block != Some(block) {
                let key = derive_binary_rc4_block_key(&self.secret, block);
                self.cipher = Some(Rc4::new_from_slice(key.as_ref()).map_err(|_| {
                    XlsError::InvalidData("invalid binary RC4 key length".to_string())
                })?);
                self.block = Some(block);
                self.offset = 0;
            }

            if self.offset > block_offset {
                return Err(XlsError::InvalidData(
                    "binary RC4 stream offsets moved backwards".to_string(),
                ));
            }
            let gap = block_offset - self.offset;
            if gap != 0 {
                let mut discarded = Zeroizing::new([0u8; BINARY_RC4_BLOCK_SIZE]);
                self.cipher
                    .as_mut()
                    .expect("binary RC4 cipher initialized")
                    .apply_keystream(&mut discarded[..gap]);
                self.offset = block_offset;
            }

            let count = data
                .len()
                .min(BINARY_RC4_BLOCK_SIZE.saturating_sub(block_offset));
            self.cipher
                .as_mut()
                .expect("binary RC4 cipher initialized")
                .apply_keystream(&mut data[..count]);
            self.offset += count;
            absolute += count;
            data = &mut data[count..];
        }
        Ok(())
    }
}

struct WriterEncryptionMaterial {
    filepass: Vec<u8>,
    cipher: WorkbookCipher,
}

fn random_16(field: &str) -> XlsResult<Zeroizing<[u8; 16]>> {
    let mut value = Zeroizing::new([0u8; 16]);
    SysRng.try_fill_bytes(value.as_mut()).map_err(|error| {
        XlsError::InvalidData(format!("failed to generate BIFF8 {field}: {error}"))
    })?;
    Ok(value)
}

fn prepare_writer_material(encryption: &XlsWriterEncryption) -> XlsResult<WriterEncryptionMaterial> {
    validate_writer_encryption(&encryption.password, encryption.profile)?;
    match encryption.profile {
        XlsEncryptionProfile::XorObfuscation => {
            let mut filepass = Vec::with_capacity(6);
            filepass.extend_from_slice(&0u16.to_le_bytes());
            filepass.extend_from_slice(&create_xor_key(&encryption.password).to_le_bytes());
            filepass.extend_from_slice(&create_xor_verifier(&encryption.password).to_le_bytes());
            Ok(WriterEncryptionMaterial {
                filepass,
                cipher: WorkbookCipher::Xor(create_xor_array(&encryption.password)),
            })
        },
        XlsEncryptionProfile::OfficeBinaryRc4 => {
            let salt = random_16("binary RC4 salt")?;
            let verifier = random_16("binary RC4 verifier")?;
            let secret = derive_binary_rc4_secret(&encryption.password, &salt);
            let key = derive_binary_rc4_block_key(&secret, 0);
            let mut encrypted = Zeroizing::new([0u8; 32]);
            encrypted[..16].copy_from_slice(verifier.as_ref());
            encrypted[16..].copy_from_slice(&Md5::digest(verifier.as_ref()));
            Rc4::new_from_slice(key.as_ref())
                .map_err(|_| XlsError::InvalidData("invalid binary RC4 key length".to_string()))?
                .apply_keystream(encrypted.as_mut());

            let mut filepass = Vec::with_capacity(BINARY_RC4_FILEPASS_LEN);
            filepass.extend_from_slice(&1u16.to_le_bytes());
            filepass.extend_from_slice(&1u16.to_le_bytes());
            filepass.extend_from_slice(&1u16.to_le_bytes());
            filepass.extend_from_slice(salt.as_ref());
            filepass.extend_from_slice(encrypted.as_ref());
            Ok(WriterEncryptionMaterial {
                filepass,
                cipher: WorkbookCipher::BinaryRc4(Box::new(BinaryRc4Stream::new(secret))),
            })
        },
        XlsEncryptionProfile::CryptoApiRc4 { key_bits } => {
            let salt = random_16("CryptoAPI salt")?;
            let verifier = random_16("CryptoAPI verifier")?;
            let (header, context) = cryptoapi::build_rc4_header_for_write(
                &encryption.password,
                usize::from(key_bits),
                CRYPTOAPI_RC4_FLAGS,
                CRYPTOAPI_RC4_PROVIDER,
                &salt,
                &verifier,
            )
            .map_err(map_cryptoapi_runtime_error)?;
            let mut filepass = Vec::with_capacity(2 + header.len());
            filepass.extend_from_slice(&1u16.to_le_bytes());
            filepass.extend_from_slice(&header);
            Ok(WriterEncryptionMaterial {
                filepass,
                cipher: WorkbookCipher::CryptoApi(context),
            })
        },
    }
}

struct ClearWorkbookPlan {
    filepass_offset: usize,
    boundsheet_payloads: Vec<usize>,
}

fn inspect_clear_workbook(workbook: &[u8]) -> XlsResult<ClearWorkbookPlan> {
    let mut position = 0usize;
    let mut record_positions = std::collections::HashSet::new();
    let mut boundsheet_payloads = Vec::new();
    let mut codepage_offset = None;
    let mut globals_eof = None;
    let mut saw_filepass = false;

    while position < workbook.len() {
        let header_end = position.checked_add(4).ok_or_else(|| {
            XlsError::InvalidData("BIFF record header offset overflow".to_string())
        })?;
        if header_end > workbook.len() {
            return Err(XlsError::InvalidData(
                "truncated BIFF record header in generated Workbook stream".to_string(),
            ));
        }
        record_positions.insert(position);
        let sid = u16::from_le_bytes([workbook[position], workbook[position + 1]]);
        let data_len = usize::from(u16::from_le_bytes([
            workbook[position + 2],
            workbook[position + 3],
        ]));
        let record_end = header_end.checked_add(data_len).ok_or_else(|| {
            XlsError::InvalidData("BIFF record length overflow".to_string())
        })?;
        if record_end > workbook.len() {
            return Err(XlsError::InvalidData(format!(
                "generated BIFF record 0x{sid:04x} extends beyond the Workbook stream"
            )));
        }
        let in_globals = globals_eof.is_none();
        if position == 0 {
            if sid != BOF_SID || data_len < 4 || workbook[header_end + 2..header_end + 4] != 0x0005u16.to_le_bytes() {
                return Err(XlsError::InvalidData(
                    "generated Workbook stream does not begin with a workbook-globals BOF"
                        .to_string(),
                ));
            }
        } else if sid == FILEPASS_SID {
            saw_filepass = true;
        }
        if in_globals && sid == CODEPAGE_SID && codepage_offset.replace(position).is_some() {
            return Err(XlsError::InvalidData(
                "generated workbook globals contain duplicate CODEPAGE records".to_string(),
            ));
        }
        if in_globals && sid == BOUNDSHEET8_SID {
            if data_len < 4 {
                return Err(XlsError::InvalidData(
                    "generated BOUNDSHEET8 record is truncated".to_string(),
                ));
            }
            boundsheet_payloads.push(header_end);
        }
        if in_globals && sid == EOF_SID {
            globals_eof = Some(record_end);
        }
        position = record_end;
    }
    if saw_filepass {
        return Err(XlsError::MalformedFilePass(
            "generated Workbook stream already contains FILEPASS".to_string(),
        ));
    }
    let filepass_offset = codepage_offset.ok_or_else(|| {
        XlsError::InvalidData("generated workbook globals are missing CODEPAGE".to_string())
    })?;
    let globals_eof = globals_eof.ok_or_else(|| {
        XlsError::InvalidData("generated workbook globals are missing EOF".to_string())
    })?;
    if filepass_offset >= globals_eof {
        return Err(XlsError::InvalidData(
            "generated CODEPAGE record is outside workbook globals".to_string(),
        ));
    }
    for &payload in &boundsheet_payloads {
        let target = u32::from_le_bytes(workbook[payload..payload + 4].try_into().unwrap()) as usize;
        if target < globals_eof || !record_positions.contains(&target) {
            return Err(XlsError::InvalidData(
                "BOUNDSHEET8 references an invalid substream offset".to_string(),
            ));
        }
        if u16::from_le_bytes([workbook[target], workbook[target + 1]]) != BOF_SID {
            return Err(XlsError::InvalidData(
                "BOUNDSHEET8 does not reference a BOF record".to_string(),
            ));
        }
    }
    Ok(ClearWorkbookPlan {
        filepass_offset,
        boundsheet_payloads,
    })
}

/// Insert FILEPASS into a completely generated clear Workbook stream and encrypt it.
pub(crate) fn encrypt_workbook_for_write(
    mut workbook: Vec<u8>,
    encryption: &XlsWriterEncryption,
) -> XlsResult<Vec<u8>> {
    let plan = inspect_clear_workbook(&workbook)?;
    let mut material = prepare_writer_material(encryption)?;
    let filepass_len = u16::try_from(material.filepass.len()).map_err(|_| {
        XlsError::InvalidData("FILEPASS payload exceeds the BIFF8 record limit".to_string())
    })?;
    let shift = material.filepass.len().checked_add(4).ok_or_else(|| {
        XlsError::InvalidData("FILEPASS record size overflow".to_string())
    })?;
    for payload in plan.boundsheet_payloads {
        let target = u32::from_le_bytes(workbook[payload..payload + 4].try_into().unwrap());
        let shifted = target
            .checked_add(u32::try_from(shift).map_err(|_| {
                XlsError::InvalidData("FILEPASS record size exceeds u32".to_string())
            })?)
            .ok_or_else(|| XlsError::InvalidData("BOUNDSHEET8 offset overflow".to_string()))?;
        workbook[payload..payload + 4].copy_from_slice(&shifted.to_le_bytes());
    }
    let mut record = Vec::with_capacity(shift);
    record.extend_from_slice(&FILEPASS_SID.to_le_bytes());
    record.extend_from_slice(&filepass_len.to_le_bytes());
    record.append(&mut material.filepass);
    workbook.splice(plan.filepass_offset..plan.filepass_offset, record);

    let mut position = 0usize;
    let mut saw_filepass = false;
    while position < workbook.len() {
        let header_end = position + 4;
        let sid = u16::from_le_bytes([workbook[position], workbook[position + 1]]);
        let data_len = usize::from(u16::from_le_bytes([
            workbook[position + 2],
            workbook[position + 3],
        ]));
        let record_end = header_end.checked_add(data_len).ok_or_else(|| {
            XlsError::InvalidData("encrypted BIFF record length overflow".to_string())
        })?;
        if record_end > workbook.len() {
            return Err(XlsError::InvalidData(
                "encrypted BIFF traversal exceeded the Workbook stream".to_string(),
            ));
        }
        if sid == FILEPASS_SID {
            if saw_filepass {
                return Err(XlsError::MalformedFilePass(
                    "multiple FILEPASS records".to_string(),
                ));
            }
            saw_filepass = true;
        } else if saw_filepass {
            let clear_prefix = if is_never_encrypted_record(sid) {
                data_len
            } else if sid == BOUNDSHEET8_SID {
                data_len.min(4)
            } else {
                0
            };
            match &mut material.cipher {
                WorkbookCipher::Xor(array) => {
                    for index in clear_prefix..data_len {
                        let absolute = header_end + index;
                        let array_index = (record_end + index) & 0x0f;
                        workbook[absolute] =
                            (workbook[absolute] ^ array[array_index]).rotate_right(3);
                    }
                },
                WorkbookCipher::BinaryRc4(stream) => {
                    let encrypted_start = header_end + clear_prefix;
                    stream.apply_at(&mut workbook[encrypted_start..record_end], encrypted_start)?;
                },
                WorkbookCipher::CryptoApi(context) => {
                    let encrypted_start = header_end + clear_prefix;
                    apply_cryptoapi_at(
                        &mut workbook[encrypted_start..record_end],
                        encrypted_start,
                        context,
                    )?;
                },
            }
        }
        position = record_end;
    }
    if !saw_filepass {
        return Err(XlsError::InvalidData(
            "FILEPASS insertion failed".to_string(),
        ));
    }
    Ok(workbook)
}

/// Validate `FILEPASS` and decrypt eligible BIFF8 record payloads in place.
pub(crate) fn prepare_workbook_stream(
    mut workbook: Vec<u8>,
    password: Option<&str>,
) -> XlsResult<Vec<u8>> {
    let mut position = 0usize;
    let mut cipher = None;
    let mut saw_filepass = false;

    while position < workbook.len() {
        let header_end = position.checked_add(4).ok_or_else(|| {
            XlsError::UnexpectedEndOfStream("BIFF record header offset overflow".to_string())
        })?;
        if header_end > workbook.len() {
            return Err(XlsError::UnexpectedEndOfStream(
                "truncated BIFF record header".to_string(),
            ));
        }
        let sid = u16::from_le_bytes([workbook[position], workbook[position + 1]]);
        let data_len =
            u16::from_le_bytes([workbook[position + 2], workbook[position + 3]]) as usize;
        let record_end = header_end.checked_add(data_len).ok_or_else(|| {
            XlsError::UnexpectedEndOfStream("BIFF record length overflow".to_string())
        })?;
        if record_end > workbook.len() {
            return Err(XlsError::UnexpectedEndOfStream(format!(
                "record 0x{sid:04x} extends beyond the Workbook stream"
            )));
        }

        // FILEPASS is a workbook-globals record. Once the globals EOF is
        // reached without one, worksheet substreams and trailing OLE padding
        // cannot change the workbook's encryption state.
        if sid == 0x000A && !saw_filepass {
            return Ok(workbook);
        }

        if sid == FILEPASS_SID {
            if saw_filepass {
                return Err(XlsError::MalformedFilePass(
                    "multiple FILEPASS records".to_string(),
                ));
            }
            saw_filepass = true;
            match FilePassRecord::parse(&workbook[header_end..record_end])? {
                FilePassRecord::Xor(obfuscation) => {
                    let password = password.ok_or(XlsError::PasswordRequired)?;
                    if create_xor_key(password) != obfuscation.key
                        || create_xor_verifier(password) != obfuscation.verifier
                    {
                        return Err(XlsError::InvalidPassword);
                    }
                    cipher = Some(WorkbookCipher::Xor(create_xor_array(password)));
                },
                FilePassRecord::BinaryRc4(filepass) => {
                    let password = password.ok_or(XlsError::PasswordRequired)?;
                    let secret = verify_binary_rc4_password(&filepass, password)?
                        .ok_or(XlsError::InvalidPassword)?;
                    cipher = Some(WorkbookCipher::BinaryRc4(Box::new(BinaryRc4Stream::new(secret))));
                },
                FilePassRecord::CryptoApi(header) => {
                    let password = password.ok_or(XlsError::PasswordRequired)?;
                    let context = cryptoapi::verify_password(&header, password)
                        .map_err(map_cryptoapi_runtime_error)?
                        .ok_or(XlsError::InvalidPassword)?;
                    cipher = Some(WorkbookCipher::CryptoApi(context));
                },
                FilePassRecord::Unsupported(kind) => {
                    return Err(XlsError::UnsupportedEncryption(kind));
                },
            }
        } else if let Some(active_cipher) = cipher.as_mut() {
            let clear_prefix = if is_never_encrypted_record(sid) {
                data_len
            } else if sid == BOUNDSHEET8_SID {
                data_len.min(4)
            } else {
                0
            };
            match active_cipher {
                WorkbookCipher::Xor(array) => {
                    for index in clear_prefix..data_len {
                        let absolute = header_end + index;
                        let array_index = (record_end + index) & 0x0f;
                        workbook[absolute] = workbook[absolute].rotate_left(3) ^ array[array_index];
                    }
                },
                WorkbookCipher::BinaryRc4(stream) => {
                    let encrypted_start = header_end + clear_prefix;
                    stream.apply_at(&mut workbook[encrypted_start..record_end], encrypted_start)?;
                },
                WorkbookCipher::CryptoApi(context) => {
                    let encrypted_start = header_end + clear_prefix;
                    apply_cryptoapi_at(
                        &mut workbook[encrypted_start..record_end],
                        encrypted_start,
                        context,
                    )?;
                },
            }
        }

        position = record_end;
    }

    Ok(workbook)
}

fn map_cryptoapi_header_error(error: CryptoApiError) -> XlsError {
    match error {
        CryptoApiError::Malformed(message) => XlsError::MalformedFilePass(message),
        CryptoApiError::UnsupportedVersion { .. } | CryptoApiError::UnsupportedAlgorithm => {
            XlsError::UnsupportedEncryption(XlsEncryptionKind::CryptoApi)
        },
    }
}

fn map_cryptoapi_runtime_error(error: CryptoApiError) -> XlsError {
    match error {
        CryptoApiError::Malformed(message) => XlsError::InvalidData(message),
        CryptoApiError::UnsupportedVersion { .. } | CryptoApiError::UnsupportedAlgorithm => {
            XlsError::UnsupportedEncryption(XlsEncryptionKind::CryptoApi)
        },
    }
}

fn apply_cryptoapi_at(
    mut data: &mut [u8],
    mut absolute: usize,
    context: &CryptoApiContext,
) -> XlsResult<()> {
    while !data.is_empty() {
        let block = u32::try_from(absolute / BINARY_RC4_BLOCK_SIZE).map_err(|_| {
            XlsError::InvalidData("Workbook stream is too large for CryptoAPI RC4".to_string())
        })?;
        let block_offset = absolute % BINARY_RC4_BLOCK_SIZE;
        let count = data
            .len()
            .min(BINARY_RC4_BLOCK_SIZE.saturating_sub(block_offset));
        cryptoapi::apply_block_cipher_at_offset(context, block, block_offset, &mut data[..count])
            .map_err(map_cryptoapi_runtime_error)?;
        absolute = absolute.checked_add(count).ok_or_else(|| {
            XlsError::InvalidData("Workbook CryptoAPI stream offset overflow".to_string())
        })?;
        data = &mut data[count..];
    }
    Ok(())
}

fn derive_binary_rc4_secret(password: &str, salt: &[u8; 16]) -> Zeroizing<[u8; 5]> {
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

fn derive_binary_rc4_block_key(secret: &[u8; 5], block: u32) -> Zeroizing<[u8; 16]> {
    let mut input = Zeroizing::new([0u8; 9]);
    input[..5].copy_from_slice(secret);
    input[5..].copy_from_slice(&block.to_le_bytes());
    Zeroizing::new(<[u8; 16]>::from(Md5::digest(input.as_slice())))
}

fn verify_binary_rc4_password(
    filepass: &BinaryRc4FilePass,
    password: &str,
) -> XlsResult<Option<Zeroizing<[u8; 5]>>> {
    let secret = derive_binary_rc4_secret(password, &filepass.salt);
    let key = derive_binary_rc4_block_key(&secret, 0);
    let mut cipher = Rc4::new_from_slice(key.as_ref())
        .map_err(|_| XlsError::InvalidData("invalid binary RC4 key length".to_string()))?;
    let mut verifier = Zeroizing::new(filepass.encrypted_verifier);
    let mut verifier_hash = Zeroizing::new(filepass.encrypted_verifier_hash);
    cipher.apply_keystream(verifier.as_mut());
    cipher.apply_keystream(verifier_hash.as_mut());
    let calculated = Zeroizing::new(<[u8; 16]>::from(Md5::digest(verifier.as_slice())));
    let difference = calculated
        .iter()
        .zip(verifier_hash.iter())
        .fold(0u8, |difference, (left, right)| difference | (left ^ right));
    Ok((difference == 0).then_some(secret))
}

fn is_never_encrypted_record(sid: u16) -> bool {
    matches!(
        sid,
        0x0809 // BOF
            | 0x002f // FILEPASS
            | 0x0194 // UsrExcl
            | 0x0195 // FileLock
            | 0x00e1 // InterfaceHdr
            | 0x0196 // RRDInfo
            | 0x0138 // RRDHead
    )
}

fn password_bytes(password: &str) -> Vec<u8> {
    password
        .encode_utf16()
        .take(15)
        .map(|unit| {
            let low = unit as u8;
            if low == 0 { (unit >> 8) as u8 } else { low }
        })
        .collect()
}

fn rotate_left_base15(value: u16) -> u16 {
    ((value << 1) & 0x7fff) | ((value >> 14) & 1)
}

fn create_xor_verifier(password: &str) -> u16 {
    let bytes = password_bytes(password);
    if bytes.is_empty() {
        return 0;
    }
    let mut verifier = 0u16;
    for byte in bytes.iter().rev() {
        verifier = rotate_left_base15(verifier) ^ u16::from(*byte);
    }
    rotate_left_base15(verifier) ^ bytes.len() as u16 ^ 0xce4b
}

fn create_xor_key(password: &str) -> u16 {
    let bytes = password_bytes(password);
    if bytes.is_empty() {
        return 0;
    }
    let mut key = INITIAL_CODE_ARRAY[bytes.len() - 1];
    for (matrix_line, mut byte) in (15 - bytes.len()..).zip(bytes) {
        for matrix_value in ENCRYPTION_MATRIX[matrix_line] {
            if byte & 1 != 0 {
                key ^= matrix_value;
            }
            byte >>= 1;
        }
    }
    key
}

fn create_xor_array(password: &str) -> [u8; 16] {
    let bytes = password_bytes(password);
    let mut array = [0u8; 16];
    array[..bytes.len()].copy_from_slice(&bytes);
    let padding_len = if bytes.is_empty() {
        15
    } else {
        16 - bytes.len()
    };
    array[bytes.len()..bytes.len() + padding_len].copy_from_slice(&PAD_ARRAY[..padding_len]);

    let key = create_xor_key(password).to_le_bytes();
    for (index, byte) in array.iter_mut().enumerate() {
        *byte = (*byte ^ key[index & 1]).rotate_left(2);
    }
    array
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xls::{XlsOpenOptions, XlsWorkbook};
    use litchi_core::sheet::{Cell, CellValue};

    fn poi_fixture(name: &str) -> std::path::PathBuf {
        std::path::PathBuf::from(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet")
            .join(name)
    }

    fn append_record(stream: &mut Vec<u8>, sid: u16, data: &[u8]) {
        stream.extend_from_slice(&sid.to_le_bytes());
        stream.extend_from_slice(&(data.len() as u16).to_le_bytes());
        stream.extend_from_slice(data);
    }

    fn binary_rc4_filepass(
        salt: [u8; 16],
        encrypted_verifier: [u8; 16],
        encrypted_verifier_hash: [u8; 16],
    ) -> Vec<u8> {
        let mut data = vec![1, 0, 1, 0, 1, 0];
        data.extend_from_slice(&salt);
        data.extend_from_slice(&encrypted_verifier);
        data.extend_from_slice(&encrypted_verifier_hash);
        data
    }

    fn cryptoapi_filepass(major: u16, key_bits: u32) -> Vec<u8> {
        let mut data = vec![1, 0];
        data.extend_from_slice(&major.to_le_bytes());
        data.extend_from_slice(&2u16.to_le_bytes());
        data.extend_from_slice(&4u32.to_le_bytes());
        data.extend_from_slice(&32u32.to_le_bytes());
        data.extend_from_slice(&4u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0x6801u32.to_le_bytes());
        data.extend_from_slice(&0x8004u32.to_le_bytes());
        data.extend_from_slice(&key_bits.to_le_bytes());
        data.extend_from_slice(&1u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&0u32.to_le_bytes());
        data.extend_from_slice(&16u32.to_le_bytes());
        data.extend_from_slice(&[0x11; 16]);
        data.extend_from_slice(&[0x22; 16]);
        data.extend_from_slice(&20u32.to_le_bytes());
        data.extend_from_slice(&[0x33; 20]);
        data
    }

    fn encrypt_payload(data: &mut [u8], record_end: usize, clear_prefix: usize, array: &[u8; 16]) {
        for index in clear_prefix..data.len() {
            data[index] = (data[index] ^ array[(record_end + index) & 0x0f]).rotate_right(3);
        }
    }

    #[test]
    fn method_one_vectors_match_apache_poi() {
        assert_eq!(create_xor_key("abc"), 20_810);
        assert_eq!(create_xor_verifier("abc"), 52_250);
        assert_eq!(
            create_xor_array("abc"),
            [
                0xac, 0xcc, 0xa4, 0xab, 0xd6, 0xba, 0xc3, 0xba, 0xd6, 0xa3, 0x2b, 0x45, 0xd3, 0x79,
                0x29, 0xbb,
            ]
        );
    }

    #[test]
    fn validates_filepass_structure_and_password() {
        assert!(matches!(
            FilePassRecord::parse(&[0]),
            Err(XlsError::MalformedFilePass(_))
        ));
        assert!(matches!(
            FilePassRecord::parse(&[0, 0, 1, 2]),
            Err(XlsError::MalformedFilePass(_))
        ));

        let mut stream = Vec::new();
        append_record(&mut stream, FILEPASS_SID, &[0, 0, 0x4a, 0x51, 0x1a, 0xcc]);
        assert!(matches!(
            prepare_workbook_stream(stream.clone(), None),
            Err(XlsError::PasswordRequired)
        ));
        assert!(matches!(
            prepare_workbook_stream(stream, Some("wrong")),
            Err(XlsError::InvalidPassword)
        ));
    }

    #[test]
    fn decrypts_payloads_with_clear_record_exceptions() {
        let array = create_xor_array("abc");
        let mut stream = Vec::new();
        append_record(&mut stream, FILEPASS_SID, &[0, 0, 0x4a, 0x51, 0x1a, 0xcc]);

        let expected = b"encrypted payload".to_vec();
        let mut encrypted = expected.clone();
        let record_end = stream.len() + 4 + encrypted.len();
        encrypt_payload(&mut encrypted, record_end, 0, &array);
        append_record(&mut stream, 0x1234, &encrypted);

        let mut bound_sheet = vec![4, 3, 2, 1, 9, 8, 7, 6];
        let expected_bound_sheet = bound_sheet.clone();
        let record_end = stream.len() + 4 + bound_sheet.len();
        encrypt_payload(&mut bound_sheet, record_end, 4, &array);
        append_record(&mut stream, BOUNDSHEET8_SID, &bound_sheet);

        let decrypted = prepare_workbook_stream(stream, Some("abc")).unwrap();
        let first_data = 10 + 4;
        assert_eq!(
            &decrypted[first_data..first_data + expected.len()],
            expected
        );
        let bound_data = first_data + expected.len() + 4;
        assert_eq!(
            &decrypted[bound_data..bound_data + expected_bound_sheet.len()],
            expected_bound_sheet
        );
    }

    #[test]
    fn validates_cryptoapi_filepass_headers() {
        for (major, key_bits) in [(2, 40), (3, 56), (4, 120), (4, 128)] {
            assert!(matches!(
                FilePassRecord::parse(&cryptoapi_filepass(major, key_bits)),
                Ok(FilePassRecord::CryptoApi(_))
            ));
        }

        let mut truncated = cryptoapi_filepass(4, 128);
        truncated.pop();
        assert!(matches!(
            FilePassRecord::parse(&truncated),
            Err(XlsError::MalformedFilePass(_))
        ));

        let mut trailing = cryptoapi_filepass(4, 128);
        trailing.push(0);
        assert!(matches!(
            FilePassRecord::parse(&trailing),
            Err(XlsError::MalformedFilePass(_))
        ));

        let mut unsupported = cryptoapi_filepass(4, 128);
        unsupported[22..26].copy_from_slice(&0x660eu32.to_le_bytes());
        assert!(matches!(
            FilePassRecord::parse(&unsupported),
            Err(XlsError::UnsupportedEncryption(
                XlsEncryptionKind::CryptoApi
            ))
        ));
    }

    #[test]
    fn cryptoapi_cursor_handles_key_sizes_clear_gaps_and_block_boundaries() {
        for key_bits in [40, 56, 120, 128] {
            let context = cryptoapi::test_context([0x5a; 20], key_bits);
            let original = vec![0xa5; 80];
            let mut encrypted = original.clone();
            apply_cryptoapi_at(&mut encrypted, 1000, &context).unwrap();
            assert_ne!(encrypted, original);
            apply_cryptoapi_at(&mut encrypted, 1000, &context).unwrap();
            assert_eq!(encrypted, original);

            let mut first = vec![0x11; 12];
            let mut second = vec![0x22; 12];
            apply_cryptoapi_at(&mut first, 1010, &context).unwrap();
            apply_cryptoapi_at(&mut second, 1040, &context).unwrap();
            apply_cryptoapi_at(&mut first, 1010, &context).unwrap();
            apply_cryptoapi_at(&mut second, 1040, &context).unwrap();
            assert_eq!(first, vec![0x11; 12]);
            assert_eq!(second, vec![0x22; 12]);
        }
    }

    #[test]
    fn opens_apache_poi_cryptoapi_workbook() {
        let path = poi_fixture("35897-type4.xls");
        let open = |password| {
            XlsWorkbook::new_with_options(
                std::fs::File::open(&path).unwrap(),
                XlsOpenOptions {
                    password,
                    ..XlsOpenOptions::default()
                },
            )
        };
        assert!(matches!(open(None), Err(XlsError::PasswordRequired)));
        assert!(matches!(
            open(Some("wrong")),
            Err(XlsError::InvalidPassword)
        ));
        open(Some("freedom")).unwrap();
    }

    #[test]
    fn binary_rc4_secret_matches_apache_poi_vector() {
        let salt = [
            0x17, 0xf6, 0xd1, 0x6b, 0x09, 0xb1, 0x5f, 0x7b, 0x4c, 0x9d, 0x03, 0xb4, 0x81, 0xb5,
            0xb4, 0x4a,
        ];
        assert_eq!(
            derive_binary_rc4_secret("MoneyForNothing", &salt).as_ref(),
            &[0xc2, 0xd9, 0x56, 0xb2, 0x6b]
        );
    }

    #[test]
    fn binary_rc4_verifier_uses_one_continuous_cipher() {
        let filepass = BinaryRc4FilePass {
            salt: [
                0xdf, 0x35, 0x52, 0x38, 0x0d, 0x75, 0x4a, 0xe6, 0x85, 0xc2, 0xfd, 0x78, 0xce, 0x3d,
                0xd1, 0xb6,
            ],
            encrypted_verifier: [
                0xd4, 0x04, 0x43, 0xec, 0xb7, 0xa7, 0x6f, 0x6a, 0xd2, 0x68, 0xc7, 0xdf, 0xcf, 0xa8,
                0x80, 0x68,
            ],
            encrypted_verifier_hash: [
                0x8d, 0xc2, 0x63, 0xcc, 0xe1, 0x1d, 0xe0, 0x05, 0x20, 0x16, 0x96, 0xaf, 0x48, 0x59,
                0x94, 0x64,
            ],
        };
        assert!(
            verify_binary_rc4_password(&filepass, "5ecret")
                .unwrap()
                .is_some()
        );
        assert!(
            verify_binary_rc4_password(&filepass, "Secret")
                .unwrap()
                .is_none()
        );
    }

    #[test]
    fn binary_rc4_filepass_length_is_exact() {
        let valid = binary_rc4_filepass([0; 16], [0; 16], [0; 16]);
        assert!(matches!(
            FilePassRecord::parse(&valid),
            Ok(FilePassRecord::BinaryRc4(_))
        ));
        for length in [53, 55] {
            let mut malformed = valid.clone();
            malformed.resize(length, 0);
            assert!(matches!(
                FilePassRecord::parse(&malformed),
                Err(XlsError::MalformedFilePass(_))
            ));
        }
    }

    #[test]
    fn binary_rc4_cursor_handles_clear_gaps_and_block_boundaries() {
        let secret = Zeroizing::new([1, 2, 3, 4, 5]);
        let plaintext = vec![0x5a; 80];
        let mut encrypted = plaintext.clone();
        BinaryRc4Stream::new(Zeroizing::new(*secret))
            .apply_at(&mut encrypted, 1000)
            .unwrap();
        assert_ne!(encrypted, plaintext);
        BinaryRc4Stream::new(secret)
            .apply_at(&mut encrypted, 1000)
            .unwrap();
        assert_eq!(encrypted, plaintext);

        let mut first = vec![0x11; 12];
        let mut second = vec![0x22; 12];
        let mut encoder = BinaryRc4Stream::new(Zeroizing::new([5, 4, 3, 2, 1]));
        encoder.apply_at(&mut first, 1010).unwrap();
        encoder.apply_at(&mut second, 1040).unwrap();
        let mut decoder = BinaryRc4Stream::new(Zeroizing::new([5, 4, 3, 2, 1]));
        decoder.apply_at(&mut first, 1010).unwrap();
        decoder.apply_at(&mut second, 1040).unwrap();
        assert_eq!(first, vec![0x11; 12]);
        assert_eq!(second, vec![0x22; 12]);
    }

    #[test]
    fn opens_apache_poi_xor_encrypted_workbook() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/poi/test-data/spreadsheet/xor-encryption-abc.xls");
        let file = std::fs::File::open(path).unwrap();
        let workbook = XlsWorkbook::new_with_options(
            file,
            XlsOpenOptions {
                password: Some("abc"),
                ..XlsOpenOptions::default()
            },
        )
        .unwrap();
        let cell = workbook.xls_worksheet(0).unwrap().get_cell(0, 0).unwrap();
        assert!(matches!(cell.value(), CellValue::Float(value) if *value == 1.0));
    }

    #[test]
    fn opens_apache_poi_binary_rc4_workbook() {
        let path = std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/ole/xls/password.xls");
        let wrong_password_file = std::fs::File::open(&path).unwrap();
        assert!(matches!(
            XlsWorkbook::new_with_options(
                wrong_password_file,
                XlsOpenOptions {
                    password: Some("wrong"),
                    ..XlsOpenOptions::default()
                },
            ),
            Err(XlsError::InvalidPassword)
        ));
        let file = std::fs::File::open(path).unwrap();
        let workbook = XlsWorkbook::new_with_options(
            file,
            XlsOpenOptions {
                password: Some("password"),
                ..XlsOpenOptions::default()
            },
        )
        .unwrap();
        let worksheet = workbook.xls_worksheet(0).unwrap();
        let found = (0..128).any(|row| {
            (0..32).any(|column| {
                matches!(
                    worksheet.get_cell(row, column).map(Cell::value),
                    Some(CellValue::String(value)) if value.contains("ZIP bomb")
                )
            })
        });
        assert!(
            found,
            "expected decrypted ZIP-bomb text in the first worksheet"
        );
    }
}
