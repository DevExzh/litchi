//! Immutable ODS metadata snapshots.

use litchi_core::{Metadata as CoreMetadata, Result};
use litchi_odf_common::core::metadata::Metadata;

use super::transaction::Transaction;

/// Maximum retained `meta.xml` source size for one ODS snapshot.
pub(crate) const MAX_XML_BYTES: usize = 16 * 1024 * 1024;

/// An immutable metadata snapshot bound to one ODS package.
///
/// The common typed ODF value is retained alongside the ergonomic cross-format
/// projection.  The bounded source text is kept only for the metadata part, so
/// a transaction can patch known fields while copying unknown XML unchanged.
#[derive(Clone, Debug)]
pub struct Snapshot {
    pub(crate) source: Option<String>,
    pub(crate) odf: Metadata,
    pub(crate) value: CoreMetadata,
}

impl Snapshot {
    /// Decode an optional `meta.xml` part without normalizing its source text.
    pub(crate) fn from_source(source: Option<String>) -> Result<Self> {
        if let Some(xml) = &source {
            if xml.len() > MAX_XML_BYTES {
                return Err(litchi_core::Error::InvalidFormat(
                    "ODS meta.xml exceeds the size limit".to_string(),
                ));
            }
        }
        let odf = source
            .as_deref()
            .map(Metadata::from_xml)
            .transpose()?
            .unwrap_or_default();
        let value = odf.clone().into();
        Ok(Self { source, odf, value })
    }

    /// Whether the package contains a physical `meta.xml` part.
    pub fn is_present(&self) -> bool {
        self.source.is_some()
    }

    /// Borrow the complete typed ODF metadata model.
    pub fn odf(&self) -> &Metadata {
        &self.odf
    }

    /// Borrow the compact cross-format metadata projection.
    pub fn value(&self) -> &CoreMetadata {
        &self.value
    }

    /// Borrow the retained source XML, if the package had a metadata part.
    pub fn xml(&self) -> Option<&str> {
        self.source.as_deref()
    }

    /// Start an isolated metadata transaction.
    pub fn transaction(&self) -> Transaction<'_> {
        Transaction::from_snapshot(self)
    }
}
