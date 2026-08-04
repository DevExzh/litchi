//! Public image media values.

/// XML part containing an image occurrence.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ImagePart {
    Content,
    Styles,
    FlatDocument,
}

/// The inert source of an OpenDocument image.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum ImageSource {
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

/// Drawing-frame context for an image or embedded-object occurrence.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct ImageFrame {
    pub name: Option<String>,
    pub xml_id: Option<String>,
    /// Short alternative title from a direct `svg:title` child.
    pub title: Option<String>,
    /// Prose alternative description from a direct `svg:desc` child.
    pub description: Option<String>,
    pub anchor_type: Option<String>,
    pub x: Option<String>,
    pub y: Option<String>,
    pub width: Option<String>,
    pub height: Option<String>,
    /// Cell address anchoring the frame's bottom-right corner
    /// (`table:end-cell-address`, spreadsheets only).
    pub end_cell_address: Option<String>,
    pub page_name: Option<String>,
    pub sheet_name: Option<String>,
    /// Whether this frame is a direct child of a spreadsheet `table:shapes` container.
    pub sheet_shape: bool,
}

/// One `draw:image` occurrence and its safely classified source.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Image {
    pub part: ImagePart,
    pub source: ImageSource,
    pub frame: Option<ImageFrame>,
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
            ImageSource::Inline { bytes, .. } => Some(bytes),
            _ => None,
        }
    }

    /// Return the resolved package path, if this references an existing package part.
    pub fn package_path(&self) -> Option<&str> {
        match &self.source {
            ImageSource::PackagePart { path, .. } => Some(path),
            _ => None,
        }
    }
}
