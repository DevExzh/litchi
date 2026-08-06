//! Package-independent Designer design-element values.

use crate::{Error, Result};

const MAX_UNKNOWN_EXTENSIONS: usize = 1_024;
const MAX_UNKNOWN_EXTENSION_BYTES: usize = 1 << 20;
const MAX_UNKNOWN_BYTES: usize = 8 << 20;

/// An extension entry that this release does not interpret.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Opaque {
    pub(super) xml: Vec<u8>,
}

impl Opaque {
    /// Borrow the exact extension bytes retained by the bounded snapshot.
    #[inline]
    pub fn xml(&self) -> &[u8] {
        &self.xml
    }

    pub(crate) fn from_wire(xml: Vec<u8>) -> Result<Self> {
        if xml.is_empty() {
            return Err(Error::Invalid("designer opaque extension is empty".into()));
        }
        if xml.len() > MAX_UNKNOWN_EXTENSION_BYTES {
            return Err(Error::Limit {
                resource: "designer opaque extension bytes",
                limit: MAX_UNKNOWN_EXTENSION_BYTES,
            });
        }
        Ok(Self { xml })
    }
}

/// A detached, lossless snapshot of one shape's optional `designElem`.
///
/// `None` from [`Snapshot::value`] preserves a present, schema-valid
/// `designElem` whose optional `val` attribute was omitted. It is distinct
/// from `Some(false)`, which is serialized as `val="false"`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    pub(super) value: Option<bool>,
    pub(super) unknown_extensions: Vec<Opaque>,
}

impl Snapshot {
    /// Construct a snapshot with an explicit design-element value.
    #[inline]
    pub fn new(value: bool) -> Self {
        Self {
            value: Some(value),
            unknown_extensions: Vec::new(),
        }
    }

    /// Return the optional `designElem/@val` value.
    #[inline]
    pub const fn value(&self) -> Option<bool> {
        self.value
    }

    /// Borrow unrelated extension entries retained byte-for-byte.
    #[inline]
    pub fn unknown_extensions(&self) -> &[Opaque] {
        &self.unknown_extensions
    }

    /// Start a detached atomic edit of this snapshot.
    #[inline]
    pub fn edit(&self) -> crate::shape::designer::Editor {
        crate::shape::designer::Editor::new(self.clone())
    }

    pub(crate) fn from_wire(value: Option<bool>, unknown_extensions: Vec<Opaque>) -> Result<Self> {
        let snapshot = Self {
            value,
            unknown_extensions,
        };
        snapshot.validate()?;
        Ok(snapshot)
    }

    pub(crate) fn validate(&self) -> Result<()> {
        if self.unknown_extensions.len() > MAX_UNKNOWN_EXTENSIONS {
            return Err(Error::Limit {
                resource: "designer opaque extensions",
                limit: MAX_UNKNOWN_EXTENSIONS,
            });
        }
        let mut total = 0usize;
        for extension in &self.unknown_extensions {
            if extension.xml.is_empty() || extension.xml.len() > MAX_UNKNOWN_EXTENSION_BYTES {
                return Err(Error::Limit {
                    resource: "designer opaque extension bytes",
                    limit: MAX_UNKNOWN_EXTENSION_BYTES,
                });
            }
            total = total.checked_add(extension.xml.len()).ok_or(Error::Limit {
                resource: "designer opaque extension bytes",
                limit: MAX_UNKNOWN_BYTES,
            })?;
            if total > MAX_UNKNOWN_BYTES {
                return Err(Error::Limit {
                    resource: "designer opaque extension bytes",
                    limit: MAX_UNKNOWN_BYTES,
                });
            }
        }
        Ok(())
    }
}
