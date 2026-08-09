//! Source-checked master-document title transactions.

use litchi_core::{Error, Metadata, Result};
use litchi_odf_common::core::{MetaXmlPatch, metadata::Metadata as OdfMetadata, patch_meta_xml};
use std::sync::Arc;

use crate::Master;

const MAX_TITLE_BYTES: usize = 16 * 1024;

/// An isolated title edit derived from one immutable master-document snapshot.
pub struct Edit<'source> {
    source: &'source Master,
    source_meta_xml: Option<String>,
    before: Option<String>,
    after: Option<String>,
}

impl<'source> Edit<'source> {
    pub(crate) fn new(source: &'source Master) -> Result<Self> {
        Ok(Self {
            source,
            source_meta_xml: source.meta_xml()?,
            before: source.title().map(str::to_owned),
            after: source.title().map(str::to_owned),
        })
    }

    /// Returns the title staged for publication.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        self.after.as_deref()
    }

    /// Replaces the document title.
    ///
    /// # Errors
    ///
    /// Returns an error when the value is too large or cannot be represented
    /// as XML 1.0 text.
    pub fn set(&mut self, value: impl Into<String>) -> Result<()> {
        let title = value.into();
        validate_title(&title)?;
        self.after = Some(title);
        Ok(())
    }

    /// Removes the document title element.
    pub fn clear(&mut self) {
        self.after = None;
    }

    /// Validates, reparses, and publishes the staged title atomically.
    ///
    /// A changed title requires an existing `meta.xml` with an `office:meta`
    /// container. This narrow initial boundary avoids inventing metadata or
    /// normalizing a package that cannot retain its metadata source.
    ///
    /// # Errors
    ///
    /// Returns an error when the source metadata cannot be patched, the
    /// candidate package cannot be validated, or title readback differs.
    pub fn commit(self) -> Result<Commit> {
        if self.before == self.after {
            let snapshot = self.source.clone();
            return Ok(Commit::new(self.source, snapshot));
        }
        let source_meta_xml = self
            .source_meta_xml
            .ok_or_else(|| invalid("ODM title editing requires an existing UTF-8 meta.xml part"))?;
        let source_metadata = OdfMetadata::from_xml(&source_meta_xml)?;
        let mut current = Metadata::from(source_metadata.clone());
        current.title.clone_from(&self.after);
        let patch = MetaXmlPatch::preserve_all().diff_simple_fields(&source_metadata, &current);
        let meta_xml = patch_meta_xml(&source_meta_xml, &patch)?
            .ok_or_else(|| invalid("ODM title editing requires an office:meta container"))?;
        if meta_xml.len() > source_meta_xml.len().saturating_add(MAX_TITLE_BYTES * 5) {
            return Err(invalid("ODM patched meta.xml exceeds the title edit limit"));
        }

        let snapshot = self.source.with_meta_xml(&meta_xml)?;
        if snapshot.title() != self.after.as_deref() {
            return Err(invalid(
                "ODM title transaction readback differs from the staged title",
            ));
        }
        Ok(Commit::new(self.source, snapshot))
    }
}

/// A validated publication result containing a new master snapshot and patch.
pub struct Commit {
    snapshot: Master,
    patch: Patch,
}

impl Commit {
    fn new(before: &Master, snapshot: Master) -> Self {
        Self {
            patch: Patch {
                before: before.shared_bytes(),
                after: snapshot.shared_bytes(),
            },
            snapshot,
        }
    }

    /// Returns the committed immutable snapshot.
    #[must_use]
    pub const fn snapshot(&self) -> &Master {
        &self.snapshot
    }

    /// Returns the exact-source-checked reversible patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes this result and returns the committed immutable snapshot.
    #[must_use]
    pub fn into_snapshot(self) -> Master {
        self.snapshot
    }
}

/// An exact-source-checked reversible master-document title patch.
#[derive(Clone)]
pub struct Patch {
    before: Arc<Vec<u8>>,
    after: Arc<Vec<u8>>,
}

impl Patch {
    /// Returns whether this patch applies to the exact source artifact.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Master) -> bool {
        source.as_bytes() == self.before.as_slice()
    }

    /// Applies this patch only to the exact package bytes from which it was made.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` is not byte-for-byte identical to the
    /// source snapshot accepted by this patch.
    pub fn apply(&self, source: &Master) -> Result<Master> {
        if !self.is_applicable_to(source) {
            return Err(invalid(
                "ODM title patch source does not match its expected snapshot",
            ));
        }
        Master::from_shared_bytes(Arc::clone(&self.after))
    }

    /// Returns the patch that restores the exact source package bytes.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            before: Arc::clone(&self.after),
            after: Arc::clone(&self.before),
        }
    }

    /// Returns whether this patch leaves the source bytes unchanged.
    #[must_use]
    pub fn is_noop(&self) -> bool {
        self.before.as_slice() == self.after.as_slice()
    }
}

fn validate_title(title: &str) -> Result<()> {
    if title.len() > MAX_TITLE_BYTES {
        return Err(invalid("ODM title exceeds the 16 KiB limit"));
    }
    if title.chars().any(|value| {
        !matches!(
            value,
            '\u{9}' | '\u{A}' | '\u{D}' | '\u{20}'..='\u{D7FF}' | '\u{E000}'..='\u{FFFD}' | '\u{10000}'..='\u{10FFFF}'
        )
    }) {
        return Err(invalid("ODM title contains a character forbidden by XML 1.0"));
    }
    Ok(())
}

fn invalid(message: &str) -> Error {
    Error::InvalidFormat(message.to_owned())
}
