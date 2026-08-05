use std::fmt;

use litchi_codepage::Ansi;

/// Resource limits for a smart-tag property-bag store.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    pub max_types: usize,
    pub max_strings: usize,
    pub max_bags: usize,
    pub max_properties: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_types: usize::from(u16::MAX) + 1,
            max_strings: 1_000_000,
            max_bags: 1_000_000,
            max_properties: 4_000_000,
        }
    }
}

/// A malformed or unsupported shared smart-tag structure.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Error(String);

impl Error {
    pub(crate) fn new(message: impl Into<String>) -> Self {
        Self(message.into())
    }
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl std::error::Error for Error {}

/// Original representation of a `PBString`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PropertyBagStringEncoding {
    Ansi,
    Utf16,
}

/// A decoded `PBString`, retaining its ANSI-versus-UTF-16 discriminator.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PropertyBagString {
    pub value: String,
    pub encoding: PropertyBagStringEncoding,
}

/// One `FactoidType` declaration.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Type {
    pub id: u16,
    pub namespace_uri: PropertyBagString,
    pub tag_name: PropertyBagString,
    pub download_url: PropertyBagString,
}

/// String-table indexes for one property.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Property {
    pub key_index: u32,
    pub value_index: u32,
}

/// One property bag attached to a format-specific smart-tag reference.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PropertyBag {
    pub type_id: u16,
    pub properties: Vec<Property>,
}

/// The shared type and string tables preceding format-specific property bags.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PropertyBagStore {
    /// Validated ANSI page used by ANSI `PBString` values.
    pub ansi: Ansi,
    /// Reserved `cfactoid` value. Consumers must not interpret it.
    pub reserved_factoid_count: u32,
    pub types: Vec<Type>,
    pub strings: Vec<PropertyBagString>,
}
