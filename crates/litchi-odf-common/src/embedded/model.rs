//! Contextual models for inert embedded OpenDocument resources.

use crate::drawing::{Frame, Part};

/// Normative embedded-object element kind.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind {
    /// A regular `draw:object`.
    Object,
    /// An OLE `draw:object-ole`.
    ObjectOle,
    /// An inert `draw:applet` declaration.
    Applet,
    /// An inert `draw:plugin` declaration.
    Plugin,
    /// An inert `draw:floating-frame` declaration.
    FloatingFrame,
}

/// One ordered, inert applet or plugin parameter.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Parameter {
    pub name: String,
    pub value: String,
}

/// Root kind of an inline XML object payload.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Root {
    /// An inline `office:document` payload.
    OpenDocument,
    /// An inline MathML `math:math` payload.
    MathMl,
}

/// Inert storage classification for an embedded object.
#[derive(Clone, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Source {
    /// An inline `office:document` or MathML payload.
    InlineXml {
        root: Root,
        xml: String,
        ignored_href: Option<String>,
    },
    /// Base64 data stored in an `office:binary-data` child.
    InlineBinary {
        bytes: Vec<u8>,
        ignored_href: Option<String>,
    },
    /// A verified opaque file in the same OpenDocument package.
    PackageFile {
        href: String,
        path: String,
        manifest_media_type: Option<String>,
    },
    /// A verified package subdocument rooted at a directory.
    PackageSubdocument {
        href: String,
        root_path: String,
        content_path: String,
        manifest_media_type: Option<String>,
    },
    /// A safe package path which is referenced but absent from the archive.
    MissingPackagePart { href: String, resolved_path: String },
    /// An external, filesystem, fragment, query-bearing, or otherwise inert link.
    Linked { href: String },
    /// A malformed producer omitted both href and inline data.
    Missing,
}

/// One inert `draw:object` or `draw:object-ole` occurrence.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Object {
    pub part: Part,
    pub kind: Kind,
    pub source: Source,
    pub frame: Option<Frame>,
    pub xml_id: Option<String>,
    pub class_id: Option<String>,
    pub notify_on_update_of_ranges: Option<String>,
    pub link_type: Option<String>,
    pub show: Option<String>,
    pub actuate: Option<String>,
    pub code: Option<String>,
    pub object_name: Option<String>,
    pub archive: Option<String>,
    /// Stored applet scripting intent. No script or applet is ever started.
    pub may_script: Option<bool>,
    pub applet_name: Option<String>,
    pub mime_type: Option<String>,
    pub frame_name: Option<String>,
    pub parameters: Vec<Parameter>,
}
