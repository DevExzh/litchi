//! Typed encryption profiles used by legacy Word documents.

/// Password-to-open encryption profile used by [`crate::DocWriter`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum EncryptionProfile {
    /// Legacy Word XOR obfuscation with an ANSI password of at most 15 characters.
    WordXorObfuscation,
    /// Office 97 binary RC4 with a 40-bit password-derived secret.
    OfficeBinaryRc4,
    /// Office CryptoAPI RC4/SHA-1 using a supported byte-aligned key size.
    CryptoApiRc4 {
        /// RC4 key size in bits. Supported values are 40 through 128 in steps of eight.
        key_bits: u16,
    },
}

impl EncryptionProfile {
    pub(crate) fn validate(self) -> Result<(), String> {
        if let Self::CryptoApiRc4 { key_bits } = self
            && (!(40..=128).contains(&key_bits) || key_bits % 8 != 0)
        {
            return Err(format!(
                "DOC CryptoAPI RC4 key size {key_bits} is not a byte-aligned value in 40..=128"
            ));
        }
        Ok(())
    }
}
