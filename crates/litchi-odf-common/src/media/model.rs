//! Public image media values.

use crate::drawing::{Frame, Part};

/// The inert source of an OpenDocument image.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Source {
    /// Base64 data stored in an `office:binary-data` child.
    Inline {
        bytes: Vec<u8>,
        /// An href present on the parent is ignored by ODF when inline data exists.
        ignored_href: Option<String>,
    },
    /// A verified file in the same OpenDocument package.
    PackagePart {
        href: String,
        path: String,
        manifest_media_type: Option<String>,
    },
    /// A safe package path which is referenced but absent from the archive.
    MissingPackagePart { href: String, resolved_path: String },
    /// An inert external, filesystem, fragment, query-bearing, or flat-document link.
    Linked { href: String },
    /// A malformed producer omitted both href and inline data.
    Missing,
}

/// One `draw:image` occurrence and its safely classified source.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Image {
    pub part: Part,
    pub source: Source,
    pub frame: Option<Frame>,
    pub xml_id: Option<String>,
    pub filter_name: Option<String>,
    pub declared_media_type: Option<String>,
    pub link_type: Option<String>,
    pub show: Option<String>,
    pub actuate: Option<String>,
    /// Zero-based position among alternative images in the same frame.
    pub alternative_index: usize,
}

impl Image {
    /// Return inline bytes without copying, if this is an inline image.
    pub fn inline_bytes(&self) -> Option<&[u8]> {
        match &self.source {
            Source::Inline { bytes, .. } => Some(bytes),
            _ => None,
        }
    }

    /// Return the resolved package path, if this references an existing package part.
    pub fn package_path(&self) -> Option<&str> {
        match &self.source {
            Source::PackagePart { path, .. } => Some(path),
            _ => None,
        }
    }
}
