//! Zero-copy codecs for MS-OFFCRYPTO IRM protected-content envelopes.
//!
//! These codecs expose the encrypted payload and its declared plaintext size.
//! They never acquire rights, activate a license, or decrypt content.

use std::fmt;

const LENGTH_FIELD_BYTES: usize = 8;
pub const IRM_AES_BLOCK_BYTES: usize = 16;

/// The semantic kind of an IRM content envelope.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtectedContentKind {
    /// Original OOXML or legacy binary document content.
    Document,
    /// LZX-compressed MHTML viewer representation.
    Viewer,
}

/// A borrowed MS-OFFCRYPTO protected-content stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ProtectedContentEnvelope<'a> {
    pub kind: ProtectedContentKind,
    /// Plaintext byte length for document content, or compressed plaintext
    /// length for viewer content.
    pub plaintext_size: u64,
    /// AES-128-ECB ciphertext, borrowed directly from the input.
    pub ciphertext: &'a [u8],
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum ProtectedContentError {
    Truncated,
    EmptyCiphertext,
    MisalignedCiphertext { length: usize },
    SizeOverflow,
    CiphertextTooShort { expected: u64, actual: usize },
}

impl fmt::Display for ProtectedContentError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Truncated => write!(
                formatter,
                "protected-content stream lacks its 8-byte length"
            ),
            Self::EmptyCiphertext => write!(formatter, "protected-content ciphertext is empty"),
            Self::MisalignedCiphertext { length } => write!(
                formatter,
                "protected-content ciphertext length {length} is not AES-block aligned"
            ),
            Self::SizeOverflow => write!(formatter, "protected-content padded size overflows u64"),
            Self::CiphertextTooShort { expected, actual } => write!(
                formatter,
                "protected-content ciphertext has {actual} bytes, fewer than the {expected} bytes required by its declared plaintext size"
            ),
        }
    }
}

impl std::error::Error for ProtectedContentError {}

/// Parse a document protected-content stream without copying its ciphertext.
pub fn parse_document_content(
    data: &[u8],
) -> Result<ProtectedContentEnvelope<'_>, ProtectedContentError> {
    parse_content(data, ProtectedContentKind::Document)
}

/// Parse an optional viewer-content stream without copying its ciphertext.
pub fn parse_viewer_content(
    data: &[u8],
) -> Result<ProtectedContentEnvelope<'_>, ProtectedContentError> {
    parse_content(data, ProtectedContentKind::Viewer)
}

/// Serialize an already-encrypted content envelope.
pub fn write_content(
    plaintext_size: u64,
    ciphertext: &[u8],
) -> Result<Vec<u8>, ProtectedContentError> {
    validate_ciphertext(plaintext_size, ciphertext)?;
    let mut output = Vec::with_capacity(LENGTH_FIELD_BYTES + ciphertext.len());
    output.extend_from_slice(&plaintext_size.to_le_bytes());
    output.extend_from_slice(ciphertext);
    Ok(output)
}

fn parse_content(
    data: &[u8],
    kind: ProtectedContentKind,
) -> Result<ProtectedContentEnvelope<'_>, ProtectedContentError> {
    let length = data
        .get(..LENGTH_FIELD_BYTES)
        .ok_or(ProtectedContentError::Truncated)?;
    let plaintext_size = u64::from_le_bytes(length.try_into().expect("slice length checked"));
    let ciphertext = &data[LENGTH_FIELD_BYTES..];
    validate_ciphertext(plaintext_size, ciphertext)?;
    Ok(ProtectedContentEnvelope {
        kind,
        plaintext_size,
        ciphertext,
    })
}

fn validate_ciphertext(
    plaintext_size: u64,
    ciphertext: &[u8],
) -> Result<(), ProtectedContentError> {
    if ciphertext.is_empty() && plaintext_size != 0 {
        return Err(ProtectedContentError::EmptyCiphertext);
    }
    if !ciphertext.len().is_multiple_of(IRM_AES_BLOCK_BYTES) {
        return Err(ProtectedContentError::MisalignedCiphertext {
            length: ciphertext.len(),
        });
    }
    let block_size =
        u64::try_from(IRM_AES_BLOCK_BYTES).map_err(|_| ProtectedContentError::SizeOverflow)?;
    let minimum = plaintext_size
        .checked_add(block_size - 1)
        .ok_or(ProtectedContentError::SizeOverflow)?
        / block_size
        * block_size;
    if u64::try_from(ciphertext.len()).map_err(|_| ProtectedContentError::SizeOverflow)? < minimum {
        return Err(ProtectedContentError::CiphertextTooShort {
            expected: minimum,
            actual: ciphertext.len(),
        });
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn document_envelope_is_zero_copy_and_round_trips() {
        let ciphertext = [0xA5; 32];
        let bytes = write_content(17, &ciphertext).unwrap();
        let envelope = parse_document_content(&bytes).unwrap();
        assert_eq!(envelope.kind, ProtectedContentKind::Document);
        assert_eq!(envelope.plaintext_size, 17);
        assert_eq!(envelope.ciphertext, ciphertext);
        assert_eq!(envelope.ciphertext.as_ptr(), bytes[8..].as_ptr());
    }

    #[test]
    fn viewer_length_describes_compressed_plaintext() {
        let mut bytes = 9u64.to_le_bytes().to_vec();
        bytes.extend_from_slice(&[0; 16]);
        assert_eq!(
            parse_viewer_content(&bytes).unwrap().kind,
            ProtectedContentKind::Viewer
        );
    }

    #[test]
    fn malformed_envelopes_are_rejected() {
        assert_eq!(
            parse_document_content(&[0; 7]).unwrap_err(),
            ProtectedContentError::Truncated
        );
        assert_eq!(
            parse_document_content(&0u64.to_le_bytes()).unwrap(),
            ProtectedContentEnvelope {
                kind: ProtectedContentKind::Document,
                plaintext_size: 0,
                ciphertext: &[],
            }
        );
        assert_eq!(
            parse_document_content(&1u64.to_le_bytes()).unwrap_err(),
            ProtectedContentError::EmptyCiphertext
        );
        let mut unaligned = 1u64.to_le_bytes().to_vec();
        unaligned.extend_from_slice(&[0; 15]);
        assert!(matches!(
            parse_document_content(&unaligned),
            Err(ProtectedContentError::MisalignedCiphertext { .. })
        ));
        let mut short = 17u64.to_le_bytes().to_vec();
        short.extend_from_slice(&[0; 16]);
        assert!(matches!(
            parse_document_content(&short),
            Err(ProtectedContentError::CiphertextTooShort { expected: 32, .. })
        ));
    }
}
