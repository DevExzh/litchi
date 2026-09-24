//! Public semantic values for the XLSB Custom Data owner.
//!
//! Package names, relationship identifiers, and BIFF12 spans intentionally
//! stay private to the package adapter.  Callers select storages with an
//! owned [`StorageId`] or a checked ordinal and work with the package-neutral
//! Custom Data values.

use std::fmt;

use super::{CustomData, CustomDataView, Properties};
use crate::package::error::{Error, Result};

/// An owned, checked Custom Data UID.
///
/// The delegated `ST_Xstring` vocabulary permits the empty string.  The
/// package adapter checks the UTF-16 and source XML bounds before constructing
/// one of these values.
#[derive(Clone, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct StorageId(String);

impl StorageId {
    /// Construct a checked UID.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        super::validate_id(&value)?;
        Ok(Self(value))
    }

    /// Borrow the decoded UID.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume this handle and return its decoded UID.
    #[must_use]
    pub fn into_inner(self) -> String {
        self.0
    }
}

impl fmt::Debug for StorageId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_tuple("StorageId").field(&self.0).finish()
    }
}

impl fmt::Display for StorageId {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0)
    }
}

impl AsRef<str> for StorageId {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for StorageId {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<&str> for StorageId {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

impl From<StorageId> for String {
    fn from(value: StorageId) -> Self {
        value.0
    }
}

/// Selector used by transaction verbs.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum StorageSelector<'a> {
    /// Select by decoded UID.
    Id(&'a str),
    /// Select by checked zero-based snapshot order.
    Index(usize),
}

/// Conversion accepted by selector-first transaction methods.
pub trait StorageSelectorInput {
    /// Resolve this selector against a semantic storage list.
    fn resolve_storage(&self, storages: &[Storage]) -> Option<usize>;
}

impl<'a> StorageSelectorInput for StorageSelector<'a> {
    fn resolve_storage(&self, storages: &[Storage]) -> Option<usize> {
        match self {
            Self::Id(id) => storages.iter().position(|storage| storage.id() == *id),
            Self::Index(index) => (*index < storages.len()).then_some(*index),
        }
    }
}

impl StorageSelectorInput for usize {
    fn resolve_storage(&self, storages: &[Storage]) -> Option<usize> {
        (*self < storages.len()).then_some(*self)
    }
}

impl StorageSelectorInput for &str {
    fn resolve_storage(&self, storages: &[Storage]) -> Option<usize> {
        storages.iter().position(|storage| storage.id() == *self)
    }
}

impl StorageSelectorInput for StorageId {
    fn resolve_storage(&self, storages: &[Storage]) -> Option<usize> {
        storages
            .iter()
            .position(|storage| storage.id() == self.as_str())
    }
}

impl StorageSelectorInput for &StorageId {
    fn resolve_storage(&self, storages: &[Storage]) -> Option<usize> {
        storages
            .iter()
            .position(|storage| storage.id() == self.as_str())
    }
}

/// One semantic Custom Data storage view.
///
/// The associated payload is inert bytes.  The private package identity is
/// retained by the source-bound snapshot and never appears in ordinary
/// accessors.
#[derive(Debug, Clone)]
pub struct Storage {
    pub(crate) value: CustomDataView,
    pub(crate) origin: Option<super::package::StorageOrigin>,
}

impl PartialEq for Storage {
    fn eq(&self, other: &Self) -> bool {
        self.value == other.value
    }
}

impl Eq for Storage {}

impl Storage {
    /// Borrow the shared semantic value.
    #[must_use]
    pub fn value(&self) -> &CustomDataView {
        &self.value
    }

    /// Borrow typed X14 properties.
    #[must_use]
    pub fn properties(&self) -> &Properties {
        self.value.properties()
    }

    /// Borrow the inert payload bytes.
    #[must_use]
    pub fn data(&self) -> &[u8] {
        self.value.data()
    }

    /// Borrow the decoded UID.
    #[must_use]
    pub fn id(&self) -> &str {
        &self.value.properties.id
    }

    /// Return an owned semantic value for an allocation-bearing handoff.
    #[must_use]
    pub fn to_owned(&self) -> CustomData {
        self.value.to_owned()
    }
}

/// Compatibility alias for code that calls a storage a part.
pub type Part = Storage;
