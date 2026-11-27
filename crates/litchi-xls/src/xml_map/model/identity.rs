//! Bounded XML-map identifiers and `XPath` values.

use crate::{Error, Result};
use std::fmt;

const MAX_MAP_ID: u32 = 2_147_483_647;
const MAX_XSTRING: usize = 65_535;
const MAX_XPATH: usize = 31_999;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct MapId(u32);

impl MapId {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(value: u32) -> Result<Self> {
        if (1..=MAX_MAP_ID).contains(&value) {
            Ok(Self(value))
        } else {
            Err(invalid(format!(
                "XML map ID {value} is outside 1..={MAX_MAP_ID}"
            )))
        }
    }

    #[must_use]
    pub const fn get(self) -> u32 {
        self.0
    }
}

impl fmt::Display for MapId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        self.0.fmt(formatter)
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct SchemaId(String);

impl SchemaId {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        Ok(Self(validate_string(
            value.into(),
            MAX_XSTRING,
            "schema ID",
            false,
        )?))
    }

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct XPath(String);

impl XPath {
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.encode_utf16().count() > MAX_XPATH {
            return Err(invalid(
                "XML column XPath must contain fewer than 32000 Unicode characters",
            ));
        }
        validate_xml_text(&value, "XML column XPath")?;
        Ok(Self(value))
    }

    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

pub(crate) fn validate_string(
    value: String,
    max_units: usize,
    label: &str,
    allow_empty: bool,
) -> Result<String> {
    let length = value.encode_utf16().count();
    if (!allow_empty && length == 0) || length > max_units {
        return Err(invalid(format!("{label} exceeds its Unicode length bound")));
    }
    validate_xml_text(&value, label)?;
    Ok(value)
}

pub(crate) fn validate_xml_text(value: &str, label: &str) -> Result<()> {
    if value.chars().any(|character| {
        (character <= '\u{1f}' && !matches!(character, '\t' | '\n' | '\r'))
            || matches!(character, '\u{fffe}' | '\u{ffff}')
    }) {
        return Err(invalid(format!(
            "{label} contains an XML-forbidden character"
        )));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidData(format!("XML map: {}", message.into()))
}
