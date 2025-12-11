pub mod agile;
mod ole_encrypted_package;
pub mod standard2007;

pub use agile::{decrypt_ooxml_package_agile, encrypt_ooxml_package_agile};
pub use standard2007::{decrypt_ooxml_package_standard_2007, encrypt_ooxml_package_standard_2007};

use crate::ole::is_ole_file;
use crate::ooxml::error::{OoxmlError, Result};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EncryptionMode {
    Standard2007,
    Agile,
}

#[derive(Debug, Clone)]
pub struct DecryptedPackage {
    pub mode: Option<EncryptionMode>,
    pub package_bytes: Vec<u8>,
}

pub fn decrypt_ooxml_ole_encrypted(ole_bytes: &[u8], password: &str) -> Result<DecryptedPackage> {
    let (encryption_info, encrypted_package) =
        ole_encrypted_package::parse_ole_encrypted_package(ole_bytes)?;

    let mode = detect_encryption_mode(&encryption_info)?;

    let package_bytes = match mode {
        EncryptionMode::Standard2007 => standard2007::decrypt_ooxml_package_standard_2007(
            &encryption_info,
            &encrypted_package,
            password,
        )?,
        EncryptionMode::Agile => {
            agile::decrypt_ooxml_package_agile(&encryption_info, &encrypted_package, password)?
        },
    };

    Ok(DecryptedPackage {
        mode: Some(mode),
        package_bytes,
    })
}

pub fn decrypt_ooxml_if_encrypted(bytes: &[u8], password: &str) -> Result<DecryptedPackage> {
    if is_ole_file(bytes) {
        return decrypt_ooxml_ole_encrypted(bytes, password);
    }

    Ok(DecryptedPackage {
        mode: None,
        package_bytes: bytes.to_vec(),
    })
}

fn detect_encryption_mode(encryption_info: &[u8]) -> Result<EncryptionMode> {
    if encryption_info.len() < 8 {
        return Err(OoxmlError::InvalidFormat(
            "EncryptionInfo stream too short for OOXML encryption header".to_string(),
        ));
    }

    let major = u16::from_le_bytes([encryption_info[0], encryption_info[1]]);
    let minor = u16::from_le_bytes([encryption_info[2], encryption_info[3]]);

    match (major, minor) {
        (3, 2) => Ok(EncryptionMode::Standard2007),
        (4, 4) => Ok(EncryptionMode::Agile),
        _ => Err(OoxmlError::InvalidFormat(format!(
            "unsupported OOXML EncryptionInfo version: {}.{}",
            major, minor
        ))),
    }
}

fn password_to_utf16le(password: &str) -> Vec<u8> {
    let mut buf = Vec::with_capacity(password.len() * 2);
    for ch in password.encode_utf16() {
        let bytes = ch.to_le_bytes();
        buf.push(bytes[0]);
        buf.push(bytes[1]);
    }
    buf
}
