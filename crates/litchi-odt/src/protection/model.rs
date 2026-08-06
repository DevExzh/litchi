//! Typed semantic values for ODT document interaction policy.

use litchi_core::{Error, Result};

use super::MAX_KEY_BYTES;

/// Document-level interaction and protection metadata stored by common ODF
/// producers in the configuration settings tree.
///
/// These values describe policy metadata only. They are never used to block
/// edits, unlock tracked changes, or change how a document is opened.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Policy {
    /// Whether classic form controls are marked protected.
    pub forms: Option<bool>,
    /// Whether bookmark targets are marked protected.
    pub bookmarks: Option<bool>,
    /// Whether the producer requests read-only loading.
    pub read_only: Option<bool>,
    /// Stored tracked-change protection digest material.
    pub redline_key: Option<Key>,
}

impl Policy {
    /// Validate the complete policy before it crosses an XML or package seam.
    pub fn validate(&self) -> Result<()> {
        if let Some(key) = &self.redline_key {
            key.validate()?;
        }
        Ok(())
    }

    /// Mark or clear classic form protection metadata.
    pub fn with_forms(mut self, value: Option<bool>) -> Self {
        self.forms = value;
        self
    }

    /// Mark or clear bookmark protection metadata.
    pub fn with_bookmarks(mut self, value: Option<bool>) -> Self {
        self.bookmarks = value;
        self
    }

    /// Mark or clear the producer's read-only loading hint.
    pub fn with_read_only(mut self, value: Option<bool>) -> Self {
        self.read_only = value;
        self
    }

    /// Replace or clear tracked-change protection digest material.
    pub fn with_redline_key(mut self, value: Option<Key>) -> Self {
        self.redline_key = value;
        self
    }

    pub(super) fn is_empty(&self) -> bool {
        self == &Self::default()
    }
}

/// ODF `base64Binary` protection material.
///
/// The bytes are an inert producer digest, not a plaintext password. The
/// value is kept private so every instance is size-checked at construction.
#[derive(Clone, PartialEq, Eq)]
pub struct Key(Vec<u8>);

impl Key {
    /// Create a checked protection digest from decoded bytes.
    pub fn new(bytes: impl Into<Vec<u8>>) -> Result<Self> {
        let key = Self(bytes.into());
        key.validate()?;
        Ok(key)
    }

    /// Borrow the decoded digest bytes without copying.
    pub fn as_bytes(&self) -> &[u8] {
        &self.0
    }

    /// Return the decoded digest bytes.
    pub fn into_bytes(self) -> Vec<u8> {
        self.0
    }

    fn validate(&self) -> Result<()> {
        if self.0.len() > MAX_KEY_BYTES {
            return Err(Error::InvalidFormat(format!(
                "redline protection key exceeds the {MAX_KEY_BYTES} byte limit"
            )));
        }
        Ok(())
    }
}

impl std::fmt::Debug for Key {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Key")
            .field("length", &self.0.len())
            .finish()
    }
}
