#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Immutable source-preserving run snapshots.

use std::sync::Arc;

use crate::error::Result;

use super::codec;
use super::model::{Symbol, Symbols};

/// An immutable, cheaply clonable run snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    symbols: Symbols,
}

impl Snapshot {
    /// Parse and retain a bounded source-preserving `w:r` snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let symbols = codec::read(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            symbols,
        })
    }

    /// Borrow the original run XML without copying it.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow all direct symbols in source order.
    #[inline]
    pub const fn symbols(&self) -> &Symbols {
        &self.symbols
    }

    /// Return the first direct symbol, preserving absence as `None`.
    #[inline]
    #[must_use]
    pub fn symbol(&self) -> Option<&Symbol> {
        self.symbols.first()
    }

    /// Alias for [`Self::symbol`].
    #[inline]
    #[must_use]
    pub fn sym_ex(&self) -> Option<&Symbol> {
        self.symbol()
    }
}
