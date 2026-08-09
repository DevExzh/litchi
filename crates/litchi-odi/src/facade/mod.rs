//! Concise family entry points.

use litchi_core::{Error, Metadata, Result};
use std::path::Path;

pub use crate::authoring::Builder;

/// Immutable document snapshot.
#[derive(Clone)]
pub struct Image {
    package: crate::package::Snapshot,
}

impl Image {
    /// Opens an image package from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        crate::package::Snapshot::open(path).map(|package| Self { package })
    }

    /// Opens an image package from in-memory bytes.
    ///
    /// # Errors
    ///
    /// Returns an error if the bytes are not a valid package.
    pub fn from_bytes(bytes: Vec<u8>) -> Result<Self> {
        crate::package::Snapshot::from_bytes(bytes).map(|package| Self { package })
    }

    /// Returns the `content.xml` document.
    #[must_use]
    pub fn content_xml(&self) -> &str {
        self.package.content_xml()
    }

    /// Returns the `styles.xml` document, if present.
    #[must_use]
    pub fn styles_xml(&self) -> Option<&str> {
        self.package.styles_xml()
    }

    /// Returns the document metadata, if present.
    #[must_use]
    pub fn metadata(&self) -> Option<&Metadata> {
        self.package.metadata()
    }

    /// Returns the raw package bytes.
    #[must_use]
    pub fn as_bytes(&self) -> &[u8] {
        self.package.as_bytes()
    }

    /// Lists the file entries stored in the package.
    ///
    /// # Errors
    ///
    /// Returns an error if the package entries cannot be enumerated.
    pub fn files(&self) -> Result<Vec<String>> {
        self.package.files()
    }

    /// Returns the inert semantic frame inventory from `content.xml`.
    ///
    /// Links and embedded bytes are reported only. They are never fetched,
    /// executed, or otherwise activated.
    #[must_use]
    pub fn frames(&self) -> &[crate::frame::Frame] {
        self.package.frames()
    }

    /// Starts a source-bound package frame-metadata transaction.
    ///
    /// This supports the same lossless existing-name and existing-source edits
    /// as [`crate::FlatImage`], while rebuilding a validated ODI package and
    /// preserving all untouched member payloads.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit {
            source: self,
            transaction: self.package.content_snapshot().transaction(),
        }
    }

    /// Consumes the snapshot and returns the raw package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }
}

/// A source-bound package frame-metadata transaction.
pub struct Edit<'a> {
    source: &'a Image,
    transaction: crate::FlatImageTransaction,
}

impl Edit<'_> {
    /// Stages a replacement for one frame's optional `draw:name`.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or an image without a
    /// losslessly editable `draw:frame` owner.
    pub fn set_frame_name(&mut self, frame: usize, name: Option<String>) -> Result<()> {
        self.transaction.set_frame_name(frame, name)
    }

    /// Stages replacement of an existing linked URI or inline image payload.
    ///
    /// Cross-kind changes are refused rather than reconstructing unknown XML.
    ///
    /// # Errors
    ///
    /// Returns an error for an invalid selector or lossy source representation.
    pub fn set_source(&mut self, frame: usize, source: crate::source::Source) -> Result<()> {
        self.transaction.set_source(frame, source)
    }

    /// Atomically validates, rebuilds, and publishes the package edit.
    ///
    /// # Errors
    ///
    /// Returns an error if the package cannot be safely rebuilt, including
    /// signed or non-compact source XML, or if semantic readback fails.
    pub fn commit(self) -> Result<Commit> {
        let content = self.transaction.commit()?;
        let changes = content.patch().changes().to_vec();
        let inverse_changes = content.patch().inverse().changes().to_vec();
        let snapshot = if changes.is_empty() {
            self.source.clone()
        } else {
            let xml = std::str::from_utf8(content.snapshot().as_bytes()).map_err(|_| {
                Error::InvalidFormat("ODI edited content.xml is not UTF-8".to_string())
            })?;
            Image {
                package: self.source.package.rebuild_with_content(xml)?,
            }
        };
        for change in &changes {
            let actual = snapshot.frames().get(change.frame()).ok_or_else(|| {
                Error::InvalidFormat("ODI edited frame disappeared during readback".to_string())
            })?;
            if actual.name() != change.after_name() || actual.source() != change.after_source() {
                return Err(Error::InvalidFormat(
                    "ODI package edit failed semantic readback".to_string(),
                ));
            }
        }
        Ok(Commit {
            snapshot: snapshot.clone(),
            patch: Patch {
                source: self.source.clone(),
                target: snapshot,
                changes,
                inverse_changes,
            },
        })
    }
}

/// A committed immutable image package and its exact-source patch.
pub struct Commit {
    snapshot: Image,
    patch: Patch,
}

impl Commit {
    /// Returns whether the package bytes changed.
    #[must_use]
    pub fn changed(&self) -> bool {
        !self.patch.changes.is_empty()
    }

    /// Returns the committed image snapshot.
    #[must_use]
    pub fn image(&self) -> &Image {
        &self.snapshot
    }

    /// Returns the reversible exact-source patch.
    #[must_use]
    pub fn patch(&self) -> &Patch {
        &self.patch
    }

    /// Consumes this commit into its published snapshot.
    #[must_use]
    pub fn into_image(self) -> Image {
        self.snapshot
    }
}

/// A source-checked reversible ODI package frame-metadata patch.
#[derive(Clone)]
pub struct Patch {
    source: Image,
    target: Image,
    changes: Vec<crate::FrameChange>,
    inverse_changes: Vec<crate::FrameChange>,
}

impl Patch {
    /// Returns whether the patch applies to this exact source byte sequence.
    #[must_use]
    pub fn is_applicable_to(&self, source: &Image) -> bool {
        self.source.as_bytes() == source.as_bytes()
    }

    /// Applies this patch only to its exact immutable source.
    ///
    /// # Errors
    ///
    /// Returns an error when the supplied source differs byte-for-byte.
    pub fn apply(&self, source: &Image) -> Result<Image> {
        if !self.is_applicable_to(source) {
            return Err(Error::InvalidFormat(
                "ODI package patch source does not match its expected snapshot".to_string(),
            ));
        }
        Ok(self.target.clone())
    }

    /// Returns the semantic changes in source order.
    #[must_use]
    pub fn changes(&self) -> &[crate::FrameChange] {
        &self.changes
    }

    /// Returns the patch that restores the exact source package.
    #[must_use]
    pub fn inverse(&self) -> Self {
        Self {
            source: self.target.clone(),
            target: self.source.clone(),
            changes: self.inverse_changes.clone(),
            inverse_changes: self.changes.clone(),
        }
    }
}

#[cfg(test)]
#[allow(
    clippy::unwrap_used,
    reason = "test code panics on failure; unwrap keeps assertions concise"
)]
mod tests {
    use super::{Builder, Image};

    #[test]
    fn builder_opens_as_validated_snapshot() {
        let bytes = Builder::new().build().unwrap();
        let document = Image::from_bytes(bytes).unwrap();
        assert!(document.content_xml().contains("<office:image"));
        assert!(!document.as_bytes().is_empty());
    }
}
