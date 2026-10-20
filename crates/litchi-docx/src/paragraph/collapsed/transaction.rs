#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Immutable source-preserving paragraph snapshots.

use std::sync::Arc;

use crate::error::Result;

use super::codec;
use super::model::Collapsed;

/// An immutable, cheaply clonable paragraph snapshot.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    value: Option<Collapsed>,
}

impl Snapshot {
    /// Parse and retain a bounded paragraph snapshot.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let value = codec::read(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            value,
        })
    }

    /// Return the current authored paragraph XML without copying it.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Return the direct `collapsed` state; absence remains observable.
    #[must_use]
    pub const fn collapsed(&self) -> Option<Collapsed> {
        self.value
    }
}
