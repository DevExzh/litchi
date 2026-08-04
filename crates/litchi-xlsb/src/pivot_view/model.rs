//! Semantic PivotTable-view values.
//!
//! The owner module supplies the PivotTable-view context, so the canonical
//! model name is concise within the `pivot_view` owner.

use std::fmt;

/// A complete PivotTable definition stream with validated framing.
#[derive(Clone, PartialEq, Eq)]
pub struct Part {
    name: String,
    cache_id: u32,
    version_created: u8,
    bytes: Vec<u8>,
}

impl Part {
    pub(super) fn new(name: String, cache_id: u32, version_created: u8, bytes: Vec<u8>) -> Self {
        Self {
            name,
            cache_id,
            version_created,
            bytes,
        }
    }

    /// Unique PivotTable view name (`irstName`).
    #[must_use]
    pub fn name(&self) -> &str {
        &self.name
    }

    /// Workbook PivotCache identifier (`idCache`).
    #[must_use]
    pub const fn cache_id(&self) -> u32 {
        self.cache_id
    }

    /// Data functionality level that created the view (`bVerSxMacro`).
    #[must_use]
    pub const fn version_created(&self) -> u8 {
        self.version_created
    }

    /// Complete original PivotTable definition stream.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        &self.bytes
    }
}

impl fmt::Debug for Part {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Part")
            .field("name", &self.name)
            .field("cache_id", &self.cache_id)
            .field("version_created", &self.version_created)
            .field("bytes", &self.bytes.len())
            .finish()
    }
}

impl fmt::Display for Part {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Part")
            .field("name", &self.name)
            .field("cache_id", &self.cache_id)
            .field("version_created", &self.version_created)
            .field("bytes", &self.bytes.len())
            .finish()
    }
}
