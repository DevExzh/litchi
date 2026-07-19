//! Strict parser for the MS-XLDM `CryptKey.bin` container.
//!
//! Recovery only unwraps the exponent-of-one `SIMPLEBLOB` already present in
//! caller-supplied bytes.  It does not obtain credentials or contact a source.

use std::fmt;

const MAGIC: [u8; 16] = [
    0x98, 0xbc, 0x21, 0x5d, 0x2d, 0x8d, 0xe6, 0x4e, 0xa8, 0xe5, 0xd0, 0x38, 0xaa, 0xc9, 0x44, 0x41,
];
const HEADER_SIZE: usize = 44;
const TRAILER_SIZE: usize = 16;
const EXPONENT_ONE_RSA_BYTES: usize = 64;
const CALG_RSA_KEYX: u32 = 0x0000_a400;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XldmCryptProvider {
    Base,
    Enhanced,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum XldmCryptAlgorithm {
    TripleDes,
    Rc2_40,
    TripleDes112,
}

impl XldmCryptAlgorithm {
    pub const fn blob_algorithm(self) -> u32 {
        match self {
            Self::TripleDes => 0x0000_6603,
            Self::Rc2_40 => 0x0000_6602,
            Self::TripleDes112 => 0x0000_6609,
        }
    }

    pub const fn key_len(self) -> usize {
        match self {
            Self::TripleDes => 24,
            Self::Rc2_40 => 5,
            Self::TripleDes112 => 16,
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum XldmCryptKeyError {
    Invalid(&'static str),
    LimitExceeded,
    IntegerOverflow,
}

impl fmt::Display for XldmCryptKeyError {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Invalid(message) => write!(f, "invalid XLDM CryptKey.bin: {message}"),
            Self::LimitExceeded => f.write_str("XLDM CryptKey.bin exceeds the caller's size limit"),
            Self::IntegerOverflow => {
                f.write_str("integer overflow while parsing XLDM CryptKey.bin")
            },
        }
    }
}

impl std::error::Error for XldmCryptKeyError {}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct XldmCryptKey<'a> {
    pub provider: XldmCryptProvider,
    pub algorithm: XldmCryptAlgorithm,
    pub blob_version: u8,
    pub key_blob: &'a [u8],
    pub exchange_algorithm: u32,
    pub encrypted_key: &'a [u8],
    pub blob_padding: &'a [u8],
}

fn read_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(
        bytes[offset..offset + 4]
            .try_into()
            .expect("fixed-size field"),
    )
}

pub fn parse_xldm_crypt_key(
    bytes: &[u8],
    max_key_data_size: usize,
) -> Result<XldmCryptKey<'_>, XldmCryptKeyError> {
    if bytes.len() < HEADER_SIZE + TRAILER_SIZE {
        return Err(XldmCryptKeyError::Invalid(
            "file is shorter than header and trailer",
        ));
    }
    if bytes[..16] != MAGIC || bytes[bytes.len() - TRAILER_SIZE..] != MAGIC {
        return Err(XldmCryptKeyError::Invalid(
            "header or trailer magic GUID is wrong",
        ));
    }
    if read_u32(bytes, 16) != 4
        || read_u32(bytes, 20) != HEADER_SIZE as u32
        || read_u32(bytes, 28) != TRAILER_SIZE as u32
    {
        return Err(XldmCryptKeyError::Invalid(
            "version, header size, or trailer size is wrong",
        ));
    }
    let key_data_size =
        usize::try_from(read_u32(bytes, 24)).map_err(|_| XldmCryptKeyError::IntegerOverflow)?;
    if key_data_size > max_key_data_size {
        return Err(XldmCryptKeyError::LimitExceeded);
    }
    let expected_size = HEADER_SIZE
        .checked_add(key_data_size)
        .and_then(|value| value.checked_add(TRAILER_SIZE))
        .ok_or(XldmCryptKeyError::IntegerOverflow)?;
    if bytes.len() != expected_size {
        return Err(XldmCryptKeyError::Invalid(
            "KeyDataSize does not account for the exact file length",
        ));
    }
    let provider = match read_u32(bytes, 32) {
        0 => XldmCryptProvider::Base,
        1 => XldmCryptProvider::Enhanced,
        _ => {
            return Err(XldmCryptKeyError::Invalid(
                "unsupported cryptographic provider",
            ));
        },
    };
    let algorithm = match read_u32(bytes, 36) {
        0 | 1 | 2 | 3 | 4 | 7 => XldmCryptAlgorithm::TripleDes,
        5 => XldmCryptAlgorithm::Rc2_40,
        6 => XldmCryptAlgorithm::TripleDes112,
        _ => {
            return Err(XldmCryptKeyError::Invalid(
                "unsupported CryptKey header algorithm",
            ));
        },
    };
    if read_u32(bytes, 40) != u32::MAX {
        return Err(XldmCryptKeyError::Invalid("flags are not 0xffffffff"));
    }
    let key_blob = &bytes[HEADER_SIZE..HEADER_SIZE + key_data_size];
    let minimum_blob = 12usize
        .checked_add(EXPONENT_ONE_RSA_BYTES)
        .ok_or(XldmCryptKeyError::IntegerOverflow)?;
    if key_blob.len() < minimum_blob {
        return Err(XldmCryptKeyError::Invalid(
            "SIMPLEBLOB is shorter than the exponent-one key payload",
        ));
    }
    if key_blob[0] != 1 || key_blob[1] < 2 || key_blob[2..4] != [0, 0] {
        return Err(XldmCryptKeyError::Invalid(
            "PUBLICKEYSTRUC is not a supported SIMPLEBLOB",
        ));
    }
    if read_u32(key_blob, 4) != algorithm.blob_algorithm() {
        return Err(XldmCryptKeyError::Invalid(
            "PUBLICKEYSTRUC algorithm does not match the header",
        ));
    }
    let exchange_algorithm = read_u32(key_blob, 8);
    if exchange_algorithm != CALG_RSA_KEYX {
        return Err(XldmCryptKeyError::Invalid(
            "SIMPLEBLOB does not use RSA key exchange",
        ));
    }
    Ok(XldmCryptKey {
        provider,
        algorithm,
        blob_version: key_blob[1],
        key_blob,
        exchange_algorithm,
        encrypted_key: &key_blob[12..minimum_blob],
        blob_padding: &key_blob[minimum_blob..],
    })
}

impl XldmCryptKey<'_> {
    pub fn recover_session_key(&self) -> Result<Vec<u8>, XldmCryptKeyError> {
        if self.encrypted_key.len() != EXPONENT_ONE_RSA_BYTES {
            return Err(XldmCryptKeyError::Invalid(
                "exponent-one encrypted key is not 512 bits",
            ));
        }
        let encoded: Vec<u8> = self.encrypted_key.iter().rev().copied().collect();
        if encoded.get(0..2) != Some(&[0, 2]) {
            return Err(XldmCryptKeyError::Invalid(
                "RSAES-PKCS1-v1_5 block type is not 2",
            ));
        }
        let separator = encoded[2..]
            .iter()
            .position(|&byte| byte == 0)
            .map(|position| position + 2)
            .ok_or(XldmCryptKeyError::Invalid(
                "RSAES-PKCS1-v1_5 padding has no separator",
            ))?;
        if separator < 10 || encoded[2..separator].contains(&0) {
            return Err(XldmCryptKeyError::Invalid(
                "RSAES-PKCS1-v1_5 padding is shorter than eight nonzero bytes",
            ));
        }
        let key = &encoded[separator + 1..];
        if key.len() != self.algorithm.key_len() {
            return Err(XldmCryptKeyError::Invalid(
                "unwrapped session key length does not match its algorithm",
            ));
        }
        Ok(key.to_vec())
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn crypt_key_file(algorithm: XldmCryptAlgorithm, header_algorithm: u32, key: &[u8]) -> Vec<u8> {
        let mut encoded = vec![0, 2];
        encoded.extend(std::iter::repeat_n(
            0xa5,
            EXPONENT_ONE_RSA_BYTES - key.len() - 3,
        ));
        encoded.push(0);
        encoded.extend_from_slice(key);
        encoded.reverse();
        let mut blob = vec![1, 2, 0, 0];
        blob.extend_from_slice(&algorithm.blob_algorithm().to_le_bytes());
        blob.extend_from_slice(&CALG_RSA_KEYX.to_le_bytes());
        blob.extend_from_slice(&encoded);
        blob.extend_from_slice(&[0xcc; 4]);
        let mut file = MAGIC.to_vec();
        file.extend_from_slice(&4u32.to_le_bytes());
        file.extend_from_slice(&(HEADER_SIZE as u32).to_le_bytes());
        file.extend_from_slice(&(blob.len() as u32).to_le_bytes());
        file.extend_from_slice(&(TRAILER_SIZE as u32).to_le_bytes());
        file.extend_from_slice(&1u32.to_le_bytes());
        file.extend_from_slice(&header_algorithm.to_le_bytes());
        file.extend_from_slice(&u32::MAX.to_le_bytes());
        file.extend_from_slice(&blob);
        file.extend_from_slice(&MAGIC);
        file
    }

    #[test]
    fn parses_and_recovers_each_specified_algorithm() {
        for (algorithm, header, key) in [
            (XldmCryptAlgorithm::TripleDes, 7, vec![0x11; 24]),
            (XldmCryptAlgorithm::Rc2_40, 5, vec![0x22; 5]),
            (XldmCryptAlgorithm::TripleDes112, 6, vec![0x33; 16]),
        ] {
            let file = crypt_key_file(algorithm, header, &key);
            let parsed = parse_xldm_crypt_key(&file, 1024).unwrap();
            assert_eq!(parsed.algorithm, algorithm);
            assert_eq!(parsed.blob_padding, [0xcc; 4]);
            assert_eq!(parsed.recover_session_key().unwrap(), key);
        }
    }

    #[test]
    fn rejects_magic_algorithm_size_and_hostile_padding() {
        let mut file = crypt_key_file(XldmCryptAlgorithm::TripleDes, 7, &[0x11; 24]);
        file[0] ^= 1;
        assert!(parse_xldm_crypt_key(&file, 1024).is_err());
        let file = crypt_key_file(XldmCryptAlgorithm::TripleDes, 5, &[0x11; 24]);
        assert!(parse_xldm_crypt_key(&file, 1024).is_err());
        let file = crypt_key_file(XldmCryptAlgorithm::TripleDes, 7, &[0x11; 24]);
        assert_eq!(
            parse_xldm_crypt_key(&file, 8),
            Err(XldmCryptKeyError::LimitExceeded)
        );
        let mut file = crypt_key_file(XldmCryptAlgorithm::TripleDes, 7, &[0x11; 24]);
        let ciphertext_end = HEADER_SIZE + 12 + EXPONENT_ONE_RSA_BYTES;
        file[ciphertext_end - 1] = 1;
        assert!(
            parse_xldm_crypt_key(&file, 1024)
                .unwrap()
                .recover_session_key()
                .is_err()
        );
    }
}
