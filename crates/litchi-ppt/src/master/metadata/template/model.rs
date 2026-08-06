//! Typed values for the bounded TemplateNameAtom payload.

use crate::package::{Error, Result};

/// The protocol maximum for TemplateNameAtom.recLen from [MS-PPT] §2.5.18.
pub const MAX_NAME_BYTES: usize = 4_168;

/// A validated UTF-16 design name carried by a TemplateNameAtom.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct Name(String);

impl Name {
    /// Create a name under the exact TemplateNameAtom byte bound.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        validate_units(value.encode_utf16().count())?;
        if value.encode_utf16().any(|unit| unit == 0) {
            return Err(Error::InvalidFormat(
                "TemplateNameAtom authoring value contains a NUL".into(),
            ));
        }
        Ok(Self(value))
    }

    /// Borrow the decoded design name.
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the value and return its decoded text.
    pub fn into_string(self) -> String {
        self.0
    }

    pub(super) fn from_wire(bytes: &[u8]) -> Result<Self> {
        if bytes.len() % 2 != 0 {
            return Err(Error::Corrupted(
                "TemplateNameAtom payload length must be even".into(),
            ));
        }
        if bytes.len() > MAX_NAME_BYTES {
            return Err(Error::Corrupted(format!(
                "TemplateNameAtom payload exceeds {MAX_NAME_BYTES} bytes"
            )));
        }
        let units = bytes
            .chunks_exact(2)
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]))
            .collect::<Vec<_>>();
        let length = units
            .iter()
            .position(|unit| *unit == 0)
            .unwrap_or(units.len());
        let value = String::from_utf16(&units[..length])
            .map_err(|_| Error::Corrupted("TemplateNameAtom contains invalid UTF-16".into()))?;
        Ok(Self(value))
    }

    pub(super) fn wire(&self) -> Result<Vec<u8>> {
        validate_units(self.0.encode_utf16().count())?;
        if self.0.encode_utf16().any(|unit| unit == 0) {
            return Err(Error::InvalidFormat(
                "TemplateNameAtom authoring value contains a NUL".into(),
            ));
        }
        let units = self.0.encode_utf16().collect::<Vec<_>>();
        let mut bytes = Vec::new();
        bytes
            .try_reserve_exact(units.len() * 2)
            .map_err(|_| Error::InvalidFormat("TemplateNameAtom allocation failed".into()))?;
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

fn validate_units(units: usize) -> Result<()> {
    let bytes = units
        .checked_mul(2)
        .ok_or_else(|| Error::InvalidFormat("TemplateNameAtom payload size overflow".into()))?;
    if bytes > MAX_NAME_BYTES {
        return Err(Error::InvalidFormat(format!(
            "TemplateNameAtom payload exceeds {MAX_NAME_BYTES} bytes"
        )));
    }
    Ok(())
}
