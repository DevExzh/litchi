//! Typed values for the bounded `SlideNameAtom` payload.

use crate::package::{Error, Result};

/// Maximum UTF-16 payload accepted by this semantic owner.
///
/// `[MS-PPT]` constrains `SlideNameAtom` to an even `recLen` but does not set
/// a protocol maximum. This implementation bound prevents untrusted names
/// from causing unbounded semantic allocations while remaining well above
/// normal presentation metadata sizes.
pub const MAX_NAME_BYTES: usize = 16 * 1024;

/// A validated UTF-16 name carried by a `SlideNameAtom`.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Name(String);

impl Name {
    /// Create a name under the semantic UTF-16 size bound.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let name = value.into();
        validate_unit_count(name.encode_utf16().count())?;
        Ok(Self(name))
    }

    /// Borrow the decoded name.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the value and return its decoded text.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0
    }

    pub(super) fn from_wire(bytes: &[u8]) -> Result<Self> {
        if !bytes.len().is_multiple_of(2) {
            return Err(Error::Corrupted(
                "SlideNameAtom payload length must be even".into(),
            ));
        }
        validate_unit_count(bytes.len() / 2)?;

        let mut units = Vec::new();
        units
            .try_reserve_exact(bytes.len() / 2)
            .map_err(|_err| Error::InvalidFormat("SlideNameAtom allocation failed".into()))?;
        units.extend(
            bytes
                .chunks_exact(2)
                .map(|pair| u16::from_le_bytes([pair[0], pair[1]])),
        );
        let value = String::from_utf16(&units)
            .map_err(|_err| Error::Corrupted("SlideNameAtom contains invalid UTF-16".into()))?;
        Ok(Self(value))
    }

    pub(super) fn wire(&self) -> Result<Vec<u8>> {
        let units = self.0.encode_utf16().collect::<Vec<_>>();
        validate_unit_count(units.len())?;
        let byte_len = units
            .len()
            .checked_mul(2)
            .ok_or_else(|| Error::InvalidFormat("SlideNameAtom payload size overflow".into()))?;
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(byte_len)
            .map_err(|_err| Error::InvalidFormat("SlideNameAtom allocation failed".into()))?;
        for unit in units {
            bytes.extend_from_slice(&unit.to_le_bytes());
        }
        Ok(bytes)
    }
}

impl AsRef<str> for Name {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for Name {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<&str> for Name {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

fn validate_unit_count(units: usize) -> Result<()> {
    let bytes = units
        .checked_mul(2)
        .ok_or_else(|| Error::InvalidFormat("SlideNameAtom payload size overflow".into()))?;
    if bytes > MAX_NAME_BYTES {
        return Err(Error::InvalidFormat(format!(
            "SlideNameAtom payload exceeds {MAX_NAME_BYTES} bytes"
        )));
    }
    if u32::try_from(bytes).is_err() {
        return Err(Error::InvalidFormat(
            "SlideNameAtom payload exceeds the PPT record length field".into(),
        ));
    }
    Ok(())
}
