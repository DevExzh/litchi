//! Source-checked master-document title edits: the `title` family.
//!
//! [`Package::edit_title`](crate::Package::edit_title) stages an [`Edit`];
//! [`Edit::commit`] publishes a new [`Package`] with its
//! [`Patch`] and [`Diagnostics`];
//! [`Package::apply_title`](crate::Package::apply_title) replays a patch onto
//! its exact source.

use litchi_core::{Error as CoreError, Metadata, Result};
use litchi_odf_common::core::{MetaXmlPatch, metadata::Metadata as OdfMetadata, patch_meta_xml};
use std::{borrow::Cow, sync::Arc};

use crate::Package;

type FamilyResult<T> = std::result::Result<T, Error>;

/// Whether a commit changed the package bytes.
pub use litchi_core::edit::Diagnostics;
/// The family error: typed limits, refusals and patch conflicts plus the
/// common-core payload.
pub use litchi_odf_common::Error;
/// The dimension named by [`Error::Limit`] and [`Error::InvalidLimit`].
pub use litchi_odf_common::LimitKind;

/// A committed title edit: the published package, its patch and diagnostics.
pub type Commit = litchi_core::edit::Commit<Package, Patch>;

const MAX_TITLE_BYTES: usize = 16 * 1024;

/// An isolated title edit derived from one immutable master-document snapshot.
pub struct Edit<'source> {
    source: &'source Package,
    before: Option<&'source str>,
    after: Option<Cow<'source, str>>,
}

impl<'source> Edit<'source> {
    pub(crate) fn new(source: &'source Package) -> Result<Self> {
        let before = source.title();
        Ok(Self {
            source,
            after: before.map(Cow::Borrowed),
            before,
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
    pub fn set(&mut self, value: impl Into<String>) -> FamilyResult<()> {
        let title = value.into();
        validate_title(&title)?;
        self.after = Some(Cow::Owned(title));
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
    pub fn commit(self) -> std::result::Result<Commit, Error> {
        Ok(self.commit_core()?)
    }

    fn commit_core(self) -> Result<Commit> {
        if self.before == self.after.as_deref() {
            let snapshot = self.source.clone();
            return Ok(commit(self.source, snapshot));
        }
        let source_meta_xml = self
            .source
            .meta_xml()?
            .ok_or_else(|| invalid("ODM title editing requires an existing UTF-8 meta.xml part"))?;
        let source_metadata =
            OdfMetadata::from_xml(&source_meta_xml, self.source.xml_limits_value())?;
        let mut current = Metadata::from(source_metadata.clone());
        current.title = self.after.as_deref().map(str::to_owned);
        let patch = MetaXmlPatch::preserve_all().diff_simple_fields(&source_metadata, &current);
        let meta_xml = patch_meta_xml(&source_meta_xml, &patch, self.source.xml_limits_value())?
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
        Ok(commit(self.source, snapshot))
    }
}

fn commit(before: &Package, snapshot: Package) -> Commit {
    let patch = Patch {
        before: before.shared_bytes(),
        after: snapshot.shared_bytes(),
    };
    let diagnostics = Diagnostics::new(!patch.is_noop());
    Commit::new(snapshot, patch, diagnostics)
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
    pub fn is_applicable_to(&self, source: &Package) -> bool {
        source.as_bytes() == self.before.as_slice()
    }

    /// Applies this patch only to the exact package bytes from which it was made.
    ///
    /// # Errors
    ///
    /// Returns an error when `source` is not byte-for-byte identical to the
    /// source snapshot accepted by this patch.
    pub(crate) fn apply(&self, source: &Package) -> std::result::Result<Commit, Error> {
        if !self.is_applicable_to(source) {
            return Err(Error::patch_conflict(
                "title.apply",
                "the package is not the patch source",
            ));
        }
        let package = Package::from_shared_bytes(Arc::clone(&self.after))?;
        Ok(Commit::new(
            package,
            self.clone(),
            Diagnostics::new(!self.is_noop()),
        ))
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

pub(crate) fn validate_title(title: &str) -> Result<()> {
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

fn invalid(message: &str) -> CoreError {
    CoreError::InvalidFormat(message.to_owned())
}
