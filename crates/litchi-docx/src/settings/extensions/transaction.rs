//! Source-preserving Word settings-extension snapshots and edits.

use std::sync::Arc;

use crate::{Error, Result};

use super::codec;
use super::model::{Extensions, Guid, OnOff};

/// An immutable complete `w:settings` snapshot with its typed extension view.
#[derive(Debug, Clone)]
pub struct Snapshot {
    xml: Arc<[u8]>,
    extensions: Extensions,
}

impl Snapshot {
    /// Parse and retain a bounded complete settings XML snapshot.
    pub fn from_xml(xml: impl Into<Vec<u8>>) -> Result<Self> {
        let xml = xml.into();
        let extensions = Extensions::parse(&xml)?;
        Ok(Self {
            xml: Arc::from(xml.into_boxed_slice()),
            extensions,
        })
    }

    /// Borrow the exact settings XML retained by this snapshot.
    #[inline]
    #[must_use]
    pub fn xml_bytes(&self) -> &[u8] {
        &self.xml
    }

    /// Borrow the contextual typed extension collection in source order.
    #[inline]
    #[must_use]
    pub const fn extensions(&self) -> &Extensions {
        &self.extensions
    }

    /// Start an isolated source-checked settings-extension edit.
    #[must_use]
    pub fn edit(&self) -> Transaction {
        Transaction {
            base: self.clone(),
            next: self.extensions.clone(),
        }
    }
}

/// A settings-extension edit that has not yet been published.
#[derive(Debug, Clone)]
pub struct Transaction {
    base: Snapshot,
    next: Extensions,
}

impl Transaction {
    /// Borrow the projected typed settings extensions.
    #[inline]
    #[must_use]
    pub const fn extensions(&self) -> &Extensions {
        &self.next
    }

    /// Set or remove the Word 2012 chart-reference tracking marker.
    pub fn set_chart_tracking_ref_based(&mut self, value: Option<OnOff>) -> Result<&mut Self> {
        self.next.set_chart_tracking_ref_based(value)?;
        Ok(self)
    }

    /// Set or remove the Word 2010 paragraph-ID context.
    pub fn set_document_id(&mut self, value: Option<u32>) -> Result<&mut Self> {
        self.next.set_document_id(value)?;
        Ok(self)
    }

    /// Set or remove the Word 2012 source-document GUID.
    pub fn set_source_document_id(&mut self, value: Option<Guid>) -> Result<&mut Self> {
        self.next.set_source_document_id(value)?;
        Ok(self)
    }

    /// Set or remove a present Word 2012 source-document element without a
    /// `val` attribute.
    pub fn set_source_document_id_without_value(&mut self, present: bool) -> Result<&mut Self> {
        self.next.set_source_document_id_without_value(present)?;
        Ok(self)
    }

    /// Set or remove the Word 2010 conflict-resolution save marker.
    pub fn set_conflict_mode(&mut self, value: Option<OnOff>) -> Result<&mut Self> {
        self.next.set_conflict_mode(value)?;
        Ok(self)
    }

    /// Set or remove the Word 2010 image-editing-data discard marker.
    pub fn set_discard_image_editing_data(&mut self, value: Option<OnOff>) -> Result<&mut Self> {
        self.next.set_discard_image_editing_data(value)?;
        Ok(self)
    }

    /// Set or remove the Word 2010 default image DPI value.
    pub fn set_default_image_dpi(&mut self, value: Option<i32>) -> Result<&mut Self> {
        self.next.set_default_image_dpi(value)?;
        Ok(self)
    }

    /// Validate and publish the edit without changing its source snapshot.
    pub fn commit(self) -> Result<Commit> {
        self.next.validate()?;
        if self.next == self.base.extensions {
            let patch = Patch {
                before: self.base.extensions.clone(),
                after: self.base.extensions.clone(),
                before_xml: self.base.xml.clone(),
                after_xml: self.base.xml.clone(),
            };
            return Ok(Commit {
                snapshot: self.base,
                patch,
            });
        }

        let xml = codec::rewrite(self.base.xml_bytes(), &self.next)?;
        let snapshot = Snapshot::from_xml(xml)?;
        let patch = Patch {
            before: self.base.extensions,
            after: self.next,
            before_xml: self.base.xml,
            after_xml: snapshot.xml.clone(),
        };
        Ok(Commit { snapshot, patch })
    }
}

/// A successful settings-extension publication.
#[derive(Debug, Clone)]
pub struct Commit {
    snapshot: Snapshot,
    patch: Patch,
}

impl Commit {
    /// Borrow the published settings snapshot.
    #[inline]
    #[must_use]
    pub const fn snapshot(&self) -> &Snapshot {
        &self.snapshot
    }

    /// Move the published settings snapshot out of the commit.
    #[must_use]
    pub fn into_snapshot(self) -> Snapshot {
        self.snapshot
    }

    /// Borrow the reversible source-checked patch.
    #[inline]
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Move the reversible patch out of the commit.
    #[must_use]
    pub fn into_patch(self) -> Patch {
        self.patch
    }
}

/// A reversible patch for the typed settings-extension projection.
#[derive(Debug, Clone)]
pub struct Patch {
    before: Extensions,
    after: Extensions,
    before_xml: Arc<[u8]>,
    after_xml: Arc<[u8]>,
}

impl Patch {
    /// Borrow the typed source precondition.
    #[inline]
    #[must_use]
    pub const fn before(&self) -> &Extensions {
        &self.before
    }

    /// Borrow the typed state produced by this patch.
    #[inline]
    #[must_use]
    pub const fn after(&self) -> &Extensions {
        &self.after
    }

    /// Return the inverse source-checked operation.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: self.after.clone(),
            after: self.before.clone(),
            before_xml: self.after_xml.clone(),
            after_xml: self.before_xml.clone(),
        }
    }

    /// Apply only to the exact source snapshot captured by this patch.
    pub fn apply(&self, source: &Snapshot) -> Result<Snapshot> {
        if source.xml.as_ref() != self.before_xml.as_ref() {
            return Err(Error::InvalidFormat(
                "settings extension patch source does not match its byte precondition".into(),
            ));
        }
        if source.extensions != self.before {
            return Err(Error::InvalidFormat(
                "settings extension patch source does not match its semantic precondition".into(),
            ));
        }
        if self.before == self.after {
            return Ok(source.clone());
        }
        Snapshot::from_xml(self.after_xml.to_vec())
    }
}
