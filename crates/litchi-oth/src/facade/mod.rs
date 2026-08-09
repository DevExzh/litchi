//! Concise family entry points.

use litchi_core::{Error, Metadata, Result};
use std::path::Path;

pub use crate::authoring::Builder;

/// A read-only semantic text-web body projection.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct TextBody {
    paragraphs: Vec<crate::paragraph::Paragraph>,
}

impl TextBody {
    /// Returns projected paragraph character data in document order.
    #[must_use]
    pub fn paragraphs(&self) -> &[crate::paragraph::Paragraph] {
        &self.paragraphs
    }
}

/// Immutable document snapshot.
pub struct Template {
    package: crate::package::Snapshot,
}

impl Template {
    /// Opens a web-template package from a file path.
    ///
    /// # Errors
    ///
    /// Returns an error if the file cannot be read or is not a valid package.
    pub fn open(path: impl AsRef<Path>) -> Result<Self> {
        crate::package::Snapshot::open(path).map(|package| Self { package })
    }

    /// Opens a web-template package from in-memory bytes.
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

    /// Projects inert paragraph character data from the validated text body.
    ///
    /// Fields, links, scripts, forms, resources, and embedded objects are not
    /// evaluated, followed, activated, or otherwise executed.
    pub fn text_body(&self) -> Result<TextBody> {
        self.package
            .paragraphs()
            .map(|paragraphs| TextBody { paragraphs })
    }

    /// Starts an exact-byte, failure-atomic no-op edit transaction.
    ///
    /// The current OTH authoring surface creates new packages only. This seam
    /// gives callers a source-checked commit and patch lifecycle without
    /// pretending unsupported mutations exist.
    #[must_use]
    pub fn edit(&self) -> Edit<'_> {
        Edit { source: self }
    }

    /// Consumes the snapshot and returns the raw package bytes.
    #[must_use]
    pub fn into_bytes(self) -> Vec<u8> {
        self.package.into_bytes()
    }
}

/// A detached no-op edit over one exact template snapshot.
pub struct Edit<'a> {
    source: &'a Template,
}

impl<'a> Edit<'a> {
    /// Publishes an exact no-op commit without mutating the source snapshot.
    #[must_use]
    pub fn commit(self) -> Commit<'a> {
        Commit {
            template: self.source,
            patch: Patch {
                expected_source: self.source.as_bytes(),
            },
        }
    }
}

/// An exact no-op commit over a validated template snapshot.
pub struct Commit<'a> {
    template: &'a Template,
    patch: Patch<'a>,
}

impl<'a> Commit<'a> {
    /// Returns the committed immutable template snapshot.
    #[must_use]
    pub const fn template(&self) -> &'a Template {
        self.template
    }

    /// Returns the source-checked exact no-op patch.
    #[must_use]
    pub const fn patch(&self) -> &Patch<'a> {
        &self.patch
    }
}

/// A source-checked patch that preserves exact package bytes.
pub struct Patch<'a> {
    expected_source: &'a [u8],
}

impl Patch<'_> {
    /// Applies this patch only to a byte-identical template source.
    ///
    /// # Errors
    ///
    /// Returns a typed invalid-format error when the immutable source differs.
    pub fn apply<'a>(&self, source: &'a Template) -> Result<&'a [u8]> {
        if source.as_bytes() != self.expected_source {
            return Err(Error::InvalidFormat(
                "OTH patch source does not match its expected snapshot".to_string(),
            ));
        }
        Ok(source.as_bytes())
    }

    /// Returns this exact no-op patch as its own inverse.
    #[must_use]
    pub const fn inverse(&self) -> Self {
        Self {
            expected_source: self.expected_source,
        }
    }
}
