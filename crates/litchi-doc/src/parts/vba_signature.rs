//! Inert VBA digital-signature variables stored by Word in `StwUser`.
//!
//! `[MS-DOC]` stores legacy VBA signatures as `WordSigBlob` values in the
//! parallel Xst array following the `SttbNames` variable-name table. This
//! module bounds and validates that table, then delegates the nested blob to
//! the shared `[MS-OSHARED]` owner. Signature and certificate-store bytes stay
//! opaque, including the PKCS#7 `SignedData` and its
//! `SpcIndirectDataContent`/`SpcIndirectDataContentV2` `contentInfo` form: no
//! certificate trust is established and no VBA project is opened or executed.

use super::super::package::{Error as PackageError, Result};
use super::fib::FileInformationBlock;
use litchi_ole_common::vba_signature;
use std::collections::HashSet;

/// FIB `FibRgFcLcb` index for `fcStwUser`/`lcbStwUser`.
pub const FIB_INDEX_STW_USER: usize = 60;

const STTB_NAMES_HEADER_LEN: usize = 6;
const NAME_EXTRA_LEN: usize = 4;
/// Maximum `StwUser` payload retained by the deferred DOC reader.
pub const MAX_STW_USER_BYTES: usize = 16 * 1024 * 1024;
const MAX_VARIABLE_COUNT: usize = u16::MAX as usize;

fn corrupted(message: impl Into<String>) -> PackageError {
    PackageError::Corrupted(message.into())
}

fn read_u16(data: &[u8], offset: usize, field: &str) -> Result<u16> {
    litchi_core::binary::read_u16_le(data, offset)
        .map_err(|error| corrupted(format!("invalid {field}: {error}")))
}

/// The well-known Word variable names used for VBA digital signatures.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum SignatureName {
    /// Legacy VBA signature variable (`Sign`).
    Sign,
    /// Agile VBA signature variable (`SigAgile`).
    SigAgile,
    /// Version-3 VBA signature variable (`SigV3`).
    SigV3,
}

impl SignatureName {
    /// Returns the exact variable name used in `SttbNames`.
    #[must_use]
    pub const fn as_str(self) -> &'static str {
        match self {
            Self::Sign => "Sign",
            Self::SigAgile => "SigAgile",
            Self::SigV3 => "SigV3",
        }
    }

    fn from_name(name: &str) -> Option<Self> {
        match name {
            "Sign" => Some(Self::Sign),
            "SigAgile" => Some(Self::SigAgile),
            "SigV3" => Some(Self::SigV3),
            _ => None,
        }
    }
}

/// One recognized inert Word VBA signature variable.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct WordVbaSignature {
    name: SignatureName,
    snapshot: vba_signature::Snapshot,
}

impl WordVbaSignature {
    /// The recognized `SttbNames` variable name.
    #[must_use]
    pub const fn name(&self) -> SignatureName {
        self.name
    }

    /// The exact spelling of the recognized variable name.
    #[must_use]
    pub const fn name_str(&self) -> &'static str {
        self.name.as_str()
    }

    /// The validated inert `WordSigBlob` snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &vba_signature::Snapshot {
        &self.snapshot
    }

    /// The exact source bytes, including the Xst count prefix and padding.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.snapshot.bytes()
    }
}

/// Recognized VBA signature variables from one `StwUser` table.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentVbaSignatures {
    signatures: Vec<WordVbaSignature>,
}

impl DocumentVbaSignatures {
    /// Returns signatures in their original `SttbNames` order.
    #[must_use]
    pub fn signatures(&self) -> &[WordVbaSignature] {
        &self.signatures
    }

    /// Returns the recognized signature with the requested variable name.
    #[must_use]
    pub fn get(&self, name: SignatureName) -> Option<&WordVbaSignature> {
        self.signatures
            .iter()
            .find(|signature| signature.name == name)
    }

    /// Number of recognized signature variables.
    #[must_use]
    pub fn len(&self) -> usize {
        self.signatures.len()
    }

    /// Whether no recognized signature variable was present.
    #[must_use]
    pub fn is_empty(&self) -> bool {
        self.signatures.is_empty()
    }

    /// Parse the optional `StwUser` table selected by FIB index 60.
    ///
    /// Parsing is bounded and deferred by [`crate::Document`], so malformed
    /// optional signature metadata does not prevent ordinary document text
    /// from opening.
    pub fn parse(fib: &FileInformationBlock, table_stream: &[u8]) -> Result<Option<Self>> {
        let Some((offset, length)) = fib.get_table_pointer(FIB_INDEX_STW_USER) else {
            return Ok(None);
        };
        if length == 0 {
            return Ok(None);
        }
        let length = usize::try_from(length)
            .map_err(|_| corrupted("StwUser length does not fit in memory"))?;
        if length > MAX_STW_USER_BYTES {
            return Err(corrupted("StwUser exceeds its bounded table size"));
        }
        let start = usize::try_from(offset)
            .map_err(|_| corrupted("StwUser offset does not fit in memory"))?;
        let end = start
            .checked_add(length)
            .ok_or_else(|| corrupted("StwUser range overflows"))?;
        let data = table_stream
            .get(start..end)
            .ok_or_else(|| corrupted("StwUser extends beyond the table stream"))?;
        Self::parse_bytes(data).map(Some)
    }

    /// Parse one complete `StwUser` payload.
    pub fn parse_bytes(data: &[u8]) -> Result<Self> {
        if data.len() > MAX_STW_USER_BYTES {
            return Err(corrupted("StwUser exceeds its bounded table size"));
        }
        if data.len() < STTB_NAMES_HEADER_LEN {
            return Err(corrupted("StwUser SttbNames header is truncated"));
        }
        if read_u16(data, 0, "StwUser fExtend")? != 0xFFFF {
            return Err(corrupted("StwUser fExtend is not 0xFFFF"));
        }
        let count = usize::from(read_u16(data, 2, "StwUser cData")?);
        if count > MAX_VARIABLE_COUNT {
            return Err(corrupted("StwUser cData exceeds u16"));
        }
        if read_u16(data, 4, "StwUser cbExtra")? != NAME_EXTRA_LEN as u16 {
            return Err(corrupted("StwUser cbExtra is not 4"));
        }

        let minimum_names = count
            .checked_mul(2 + NAME_EXTRA_LEN)
            .and_then(|size| size.checked_add(STTB_NAMES_HEADER_LEN))
            .ok_or_else(|| corrupted("StwUser name table size overflows"))?;
        if minimum_names > data.len() {
            return Err(corrupted("StwUser name table is truncated"));
        }

        let mut names = Vec::new();
        names
            .try_reserve_exact(count)
            .map_err(|error| corrupted(format!("StwUser name allocation failed: {error}")))?;
        let mut unique_names = HashSet::new();
        unique_names
            .try_reserve(count)
            .map_err(|error| corrupted(format!("StwUser name index allocation failed: {error}")))?;
        let mut offset = STTB_NAMES_HEADER_LEN;
        for index in 0..count {
            let (name, next) = read_xst(data, offset, &format!("StwUser name {index}"))?;
            offset = next
                .checked_add(NAME_EXTRA_LEN)
                .ok_or_else(|| corrupted("StwUser name extra range overflows"))?;
            if data.get(next..offset).is_none() {
                return Err(corrupted(format!(
                    "StwUser name {index} extra is truncated"
                )));
            }
            if !unique_names.insert(name.clone()) {
                return Err(corrupted(format!("StwUser name {index} is not unique")));
            }
            names.push(name);
        }

        let minimum_values = count
            .checked_mul(2)
            .ok_or_else(|| corrupted("StwUser value table size overflows"))?;
        if data.len().saturating_sub(offset) < minimum_values {
            return Err(corrupted("StwUser value table is truncated"));
        }

        let mut signatures = Vec::new();
        signatures
            .try_reserve_exact(count.min(3))
            .map_err(|error| corrupted(format!("StwUser signature allocation failed: {error}")))?;
        for (index, name) in names.iter().enumerate() {
            let (value, next) = read_xst_bytes(data, offset, &format!("StwUser value {index}"))?;
            offset = next;
            if let Some(name) = SignatureName::from_name(name) {
                let snapshot = vba_signature::Snapshot::parse_word(value).map_err(|error| {
                    corrupted(format!(
                        "StwUser {name:?} value is not a valid WordSigBlob: {error}"
                    ))
                })?;
                signatures.push(WordVbaSignature { name, snapshot });
            }
        }
        if offset != data.len() {
            return Err(corrupted("StwUser has trailing bytes"));
        }
        Ok(Self { signatures })
    }
}

fn read_xst(data: &[u8], offset: usize, field: &str) -> Result<(String, usize)> {
    let (serialized, end) = read_xst_bytes(data, offset, field)?;
    let bytes = &serialized[2..];
    let mut units = Vec::new();
    units
        .try_reserve_exact(bytes.len() / 2)
        .map_err(|error| corrupted(format!("{field} allocation failed: {error}")))?;
    units.extend(
        bytes
            .as_chunks::<2>()
            .0
            .iter()
            .map(|chunk| u16::from_le_bytes([chunk[0], chunk[1]])),
    );
    let value = String::from_utf16(&units)
        .map_err(|_| corrupted(format!("{field} contains invalid UTF-16")))?;
    Ok((value, end))
}

fn read_xst_bytes<'a>(data: &'a [u8], offset: usize, field: &str) -> Result<(&'a [u8], usize)> {
    let cch = usize::from(read_u16(data, offset, &format!("{field} cch"))?);
    let chars_offset = offset
        .checked_add(2)
        .ok_or_else(|| corrupted(format!("{field} offset overflows")))?;
    let byte_count = cch
        .checked_mul(2)
        .ok_or_else(|| corrupted(format!("{field} byte count overflows")))?;
    let end = chars_offset
        .checked_add(byte_count)
        .ok_or_else(|| corrupted(format!("{field} range overflows")))?;
    let bytes = data
        .get(offset..end)
        .ok_or_else(|| corrupted(format!("{field} is truncated")))?;
    Ok((bytes, end))
}

#[cfg(test)]
mod tests {
    use super::*;

    fn word_blob() -> Vec<u8> {
        let mut bytes = vec![0; 56];
        bytes[0..2].copy_from_slice(&27u16.to_le_bytes());
        bytes[2..6].copy_from_slice(&45u32.to_le_bytes());
        bytes[6..10].copy_from_slice(&8u32.to_le_bytes());
        bytes[10..14].copy_from_slice(&2u32.to_le_bytes());
        bytes[14..18].copy_from_slice(&44u32.to_le_bytes());
        bytes[18..22].copy_from_slice(&3u32.to_le_bytes());
        bytes[22..26].copy_from_slice(&46u32.to_le_bytes());
        bytes[30..34].copy_from_slice(&49u32.to_le_bytes());
        bytes[42..46].copy_from_slice(&51u32.to_le_bytes());
        bytes[46..48].copy_from_slice(&[9, 8]);
        bytes[48..51].copy_from_slice(&[7, 6, 5]);
        bytes[55] = 0xEE;
        bytes
    }

    fn xst(value: &str) -> Vec<u8> {
        let units = value.encode_utf16().collect::<Vec<_>>();
        let mut bytes = Vec::with_capacity(2 + units.len() * 2);
        bytes.extend_from_slice(&(units.len() as u16).to_le_bytes());
        bytes.extend(units.into_iter().flat_map(u16::to_le_bytes));
        bytes
    }

    fn stw_user() -> Vec<u8> {
        let names = ["Sign", "Other"];
        let values = [word_blob(), xst("plain")];
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0xFFFFu16.to_le_bytes());
        bytes.extend_from_slice(&(names.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        for name in names {
            bytes.extend(xst(name));
            bytes.extend_from_slice(&0u32.to_le_bytes());
        }
        for value in values {
            bytes.extend(value);
        }
        bytes
    }

    fn stw_user_with_supported_names() -> Vec<u8> {
        let names = ["Sign", "SigAgile", "SigV3"];
        let values = [word_blob(), word_blob(), word_blob()];
        let mut bytes = Vec::new();
        bytes.extend_from_slice(&0xFFFFu16.to_le_bytes());
        bytes.extend_from_slice(&(names.len() as u16).to_le_bytes());
        bytes.extend_from_slice(&4u16.to_le_bytes());
        for name in names {
            bytes.extend(xst(name));
            bytes.extend_from_slice(&[0xA5, 0x5A, 0xC3, 0x3C]);
        }
        for value in values {
            bytes.extend(value);
        }
        bytes
    }

    #[test]
    fn parses_signature_variable_and_ignores_unrecognized_value() {
        let parsed = DocumentVbaSignatures::parse_bytes(&stw_user()).unwrap();
        let signature = parsed.get(SignatureName::Sign).unwrap();
        assert_eq!(signature.name_str(), "Sign");
        assert_eq!(signature.snapshot().kind(), vba_signature::Kind::Word);
        assert_eq!(signature.snapshot().info().signature(), [9, 8]);
        assert_eq!(parsed.len(), 1);
    }

    #[test]
    fn parses_each_supported_signature_name_and_preserves_word_blob_bytes() {
        let source = stw_user_with_supported_names();
        let parsed = DocumentVbaSignatures::parse_bytes(&source).unwrap();

        assert_eq!(parsed.len(), 3);
        assert_eq!(
            parsed
                .signatures()
                .iter()
                .map(WordVbaSignature::name)
                .collect::<Vec<_>>(),
            [
                SignatureName::Sign,
                SignatureName::SigAgile,
                SignatureName::SigV3
            ]
        );
        assert_eq!(
            parsed.get(SignatureName::SigAgile).unwrap().bytes(),
            word_blob().as_slice()
        );
        assert_eq!(
            parsed
                .get(SignatureName::SigV3)
                .unwrap()
                .snapshot()
                .info()
                .signature(),
            [9, 8]
        );
    }

    #[test]
    fn rejects_bad_header_duplicate_names_and_trailing_values() {
        let mut bad_header = stw_user();
        bad_header[0] = 0;
        assert!(DocumentVbaSignatures::parse_bytes(&bad_header).is_err());

        let mut truncated_count = stw_user();
        truncated_count[2..4].copy_from_slice(&u16::MAX.to_le_bytes());
        assert!(
            DocumentVbaSignatures::parse_bytes(&truncated_count)
                .expect_err("oversized name count must fail before allocation")
                .to_string()
                .contains("name table is truncated")
        );

        let mut duplicate = stw_user();
        // The second name's Xst payload starts after the first name and its
        // four-byte ignored extra field. Replace `Other` with `Sign`.
        let second_name = 6 + 2 + 8 + 4;
        duplicate[second_name..second_name + 2].copy_from_slice(&4u16.to_le_bytes());
        duplicate[second_name + 2..second_name + 10]
            .copy_from_slice(&[b'S', 0, b'i', 0, b'g', 0, b'n', 0]);
        assert!(DocumentVbaSignatures::parse_bytes(&duplicate).is_err());

        let mut trailing = stw_user();
        trailing.push(0);
        assert!(DocumentVbaSignatures::parse_bytes(&trailing).is_err());
    }
}
