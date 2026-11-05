#![expect(
    clippy::shadow_reuse,
    reason = "parser bindings are intentionally refined after validation"
)]
//! Source-preserving OpenType snapshots.

use std::sync::Arc;

use crate::error::Result;

use super::Limits;
use super::codec;
use super::model::OpenType;

/// An immutable complete run/rPr snapshot with its typed OpenType projection.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    value: OpenType,
}

impl Snapshot {
    /// Parse and retain a bounded source fragment.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation cannot be completed.
    pub fn from_xml(xml: impl Into<Vec<u8>>, limits: &Limits) -> Result<Self> {
        let xml = xml.into();
        super::limit::admit_xml(&xml, limits)?;
        let value = codec::parse(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            value,
        })
    }

    /// Borrow the exact source bytes retained by this snapshot.
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow the typed OpenType projection.
    #[must_use]
    pub const fn open_type(&self) -> &OpenType {
        &self.value
    }
}
