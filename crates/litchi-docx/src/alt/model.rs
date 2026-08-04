use crate::{Error, Result};
use litchi_opc::constants::relationship_type;

pub(crate) const TRANSITIONAL_WORD_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/wordprocessingml/2006/main";
pub(crate) const STRICT_WORD_NAMESPACE: &[u8] = b"http://purl.oclc.org/ooxml/wordprocessingml/main";
pub(crate) const TRANSITIONAL_RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://schemas.openxmlformats.org/officeDocument/2006/relationships";
pub(crate) const STRICT_RELATIONSHIP_NAMESPACE: &[u8] =
    b"http://purl.oclc.org/ooxml/officeDocument/relationships";
pub(crate) const STRICT_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/afChunk";

/// Maximum payload accepted by the safe package-authoring facade.
pub const MAX_DATA_BYTES: usize = 128 * 1024 * 1024;
/// Maximum anchors accepted in one main-document part.
pub const MAX_CHUNKS: usize = 4096;
/// Maximum main-document XML accepted by the bounded scanner.
pub const MAX_XML_BYTES: usize = 32 * 1024 * 1024;
/// Maximum XML nesting accepted by the bounded scanner.
pub const MAX_XML_DEPTH: usize = 256;

pub(crate) const MAX_VISIBILITY_OFFSETS: usize = 1_000_000;
pub(crate) const MAX_MARKED_XML_BYTES: usize = 128 * 1024 * 1024;

/// A validated OPC relationship identifier used by a [`Chunk`].
///
/// This type is intentionally low-level. Package CRUD accepts semantic [`Import`]
/// values and allocates relationship identifiers itself.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Rel(Box<str>);

impl Rel {
    /// Validate an identifier that can be emitted without XML escaping.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.is_empty()
            || value.len() > 255
            || !value.bytes().all(|byte| {
                byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.' | b':')
            })
        {
            return Err(invalid(
                "altChunk relationship ID is not a safe XML attribute value",
            ));
        }
        Ok(Self(value.into_boxed_str()))
    }

    /// Borrow the underlying OPC identifier.
    #[inline]
    pub fn as_str(&self) -> &str {
        &self.0
    }
}

impl AsRef<str> for Rel {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// A validated external target URI that is retained but never accessed.
#[derive(Debug, Clone, PartialEq, Eq, PartialOrd, Ord, Hash)]
#[repr(transparent)]
pub struct Uri(Box<str>);

impl Uri {
    /// Validate a bounded, inert external target.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let value = value.into();
        if value.is_empty() || value.len() > 32_768 || value.chars().any(char::is_control) {
            return Err(invalid("external altChunk target is empty or unsafe"));
        }
        Ok(Self(value.into_boxed_str()))
    }

    /// Borrow the preserved target URI.
    #[inline]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Move the target URI out without copying it.
    #[inline]
    pub fn into_string(self) -> String {
        self.0.into()
    }
}

impl AsRef<str> for Uri {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

/// Namespace family used when emitting a new alternative-format anchor.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum Conformance {
    #[default]
    Transitional,
    Strict,
}

impl Conformance {
    /// Relationship type emitted for this conformance family.
    ///
    /// Word uses the case-sensitive `aFChunk` spelling in Transitional files.
    #[inline]
    pub const fn relationship(self) -> &'static str {
        match self {
            Self::Transitional => relationship_type::MS_ALTERNATIVE_FORMAT_IMPORT,
            Self::Strict => STRICT_RELATIONSHIP,
        }
    }

    pub(crate) const fn word_namespace(self) -> &'static str {
        match self {
            Self::Transitional => "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            Self::Strict => "http://purl.oclc.org/ooxml/wordprocessingml/main",
        }
    }

    pub(crate) const fn relationship_namespace(self) -> &'static str {
        match self {
            Self::Transitional => {
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
            },
            Self::Strict => "http://purl.oclc.org/ooxml/officeDocument/relationships",
        }
    }
}

/// Owned bytes for every alternative-format media type supported by Word.
#[derive(Debug, PartialEq, Eq)]
pub enum Data {
    Docx(Vec<u8>),
    Docm(Vec<u8>),
    Dotx(Vec<u8>),
    Dotm(Vec<u8>),
    Mime(Vec<u8>),
    Html(Vec<u8>),
    Xhtml(Vec<u8>),
    Rtf(Vec<u8>),
    Text(Vec<u8>),
    Xml(Vec<u8>),
}

impl Data {
    /// Canonical media type for the payload.
    pub const fn media_type(&self) -> &'static str {
        match self {
            Self::Docx(_) => {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
            },
            Self::Docm(_) => "application/vnd.ms-word.document.macroEnabled.main+xml",
            Self::Dotx(_) => {
                "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"
            },
            Self::Dotm(_) => "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
            Self::Mime(_) => "message/rfc822",
            Self::Html(_) => "text/html",
            Self::Xhtml(_) => "application/xhtml+xml",
            Self::Rtf(_) => "application/rtf",
            Self::Text(_) => "text/plain",
            Self::Xml(_) => "application/xml",
        }
    }

    /// Conventional package extension for the payload.
    pub const fn extension(&self) -> &'static str {
        match self {
            Self::Docx(_) => "docx",
            Self::Docm(_) => "docm",
            Self::Dotx(_) => "dotx",
            Self::Dotm(_) => "dotm",
            Self::Mime(_) => "eml",
            Self::Html(_) => "html",
            Self::Xhtml(_) => "xhtml",
            Self::Rtf(_) => "rtf",
            Self::Text(_) => "txt",
            Self::Xml(_) => "xml",
        }
    }

    /// Borrow the opaque payload bytes.
    pub fn bytes(&self) -> &[u8] {
        match self {
            Self::Docx(bytes)
            | Self::Docm(bytes)
            | Self::Dotx(bytes)
            | Self::Dotm(bytes)
            | Self::Mime(bytes)
            | Self::Html(bytes)
            | Self::Xhtml(bytes)
            | Self::Rtf(bytes)
            | Self::Text(bytes)
            | Self::Xml(bytes) => bytes,
        }
    }

    /// Payload size in bytes.
    #[inline]
    pub fn len(&self) -> usize {
        self.bytes().len()
    }

    /// Whether the payload is empty.
    #[inline]
    pub fn is_empty(&self) -> bool {
        self.bytes().is_empty()
    }

    /// Enforce the package-authoring resource limit.
    pub fn validate(&self) -> Result<()> {
        if self.len() > MAX_DATA_BYTES {
            return Err(invalid(
                "alternative-format part exceeds the 128 MiB authoring limit",
            ));
        }
        Ok(())
    }

    /// Move the opaque payload bytes out without copying them.
    pub fn into_bytes(self) -> Vec<u8> {
        match self {
            Self::Docx(bytes)
            | Self::Docm(bytes)
            | Self::Dotx(bytes)
            | Self::Dotm(bytes)
            | Self::Mime(bytes)
            | Self::Html(bytes)
            | Self::Xhtml(bytes)
            | Self::Rtf(bytes)
            | Self::Text(bytes)
            | Self::Xml(bytes) => bytes,
        }
    }
}

/// Package target to create for an alternative-format import.
#[derive(Debug, PartialEq, Eq)]
pub enum Import {
    /// An owned payload moved into the package.
    Data(Data),
    /// A validated URI that is preserved but never fetched or interpreted.
    Link(Uri),
}

impl Import {
    /// Wrap an owned internal payload.
    #[inline]
    pub const fn data(data: Data) -> Self {
        Self::Data(data)
    }

    /// Validate an inert external target.
    #[inline]
    pub fn link(uri: impl Into<String>) -> Result<Self> {
        Uri::new(uri).map(Self::Link)
    }
}

/// A block-level `<w:altChunk>` import anchor.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Chunk {
    relationship: Rel,
    match_source: Option<bool>,
}

impl Chunk {
    /// Bind a low-level anchor to a validated OPC relationship.
    #[inline]
    pub const fn new(relationship: Rel, match_source: Option<bool>) -> Self {
        Self {
            relationship,
            match_source,
        }
    }

    /// Relationship identifying the alternative-format import target.
    #[inline]
    pub const fn relationship(&self) -> &Rel {
        &self.relationship
    }

    /// Whether imported formatting should match the source formatting.
    ///
    /// `None` means `<w:matchSrc>` was absent. `Some(true)` includes the
    /// empty-element form, while `Some(false)` represents an explicit false.
    #[inline]
    pub const fn match_source(&self) -> Option<bool> {
        self.match_source
    }
}

/// Recognized MIME family for an opaque alternative-format import part.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Kind {
    Docx,
    Docm,
    Dotx,
    Dotm,
    Mime,
    Html,
    Xhtml,
    Rtf,
    Text,
    Xml,
    Unknown,
}

impl Kind {
    /// Classify a media type without inspecting its payload.
    pub fn from_media_type(value: &str) -> Self {
        let media_type = value.split(';').next().map_or(value, str::trim);
        const TYPES: &[(&str, Kind)] = &[
            (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                Kind::Docx,
            ),
            (
                "application/vnd.ms-word.document.macroEnabled.main+xml",
                Kind::Docm,
            ),
            (
                "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml",
                Kind::Dotx,
            ),
            (
                "application/vnd.ms-word.template.macroEnabledTemplate.main+xml",
                Kind::Dotm,
            ),
            ("message/rfc822", Kind::Mime),
            ("text/html", Kind::Html),
            ("application/xhtml+xml", Kind::Xhtml),
            ("application/rtf", Kind::Rtf),
            ("text/rtf", Kind::Rtf),
            ("text/plain", Kind::Text),
            ("application/xml", Kind::Xml),
            ("text/xml", Kind::Xml),
        ];
        match TYPES
            .iter()
            .find_map(|(known, kind)| media_type.eq_ignore_ascii_case(known).then_some(*kind))
        {
            Some(kind) => kind,
            None => Self::Unknown,
        }
    }
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}
