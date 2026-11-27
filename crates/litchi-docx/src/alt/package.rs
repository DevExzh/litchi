use litchi_opc::packuri::PackURI;
use litchi_opc::part::Part as OpcPart;

use super::model::Kind;

/// Resolved target of an alternative-format anchor.
pub enum Target<'a> {
    /// Borrowed internal package part.
    Part(Part<'a>),
    /// Borrowed external URI; it is never accessed.
    Link(&'a str),
}

/// A borrowed, opaque alternative-format import payload.
///
/// Access never parses the foreign format, opens nested packages, fetches
/// resources, or performs filesystem or network I/O.
pub struct Part<'a> {
    part: &'a dyn OpcPart,
    kind: Kind,
}

impl<'a> Part<'a> {
    /// Borrow an opaque OPC part without copying its payload.
    pub fn new(part: &'a dyn OpcPart) -> Self {
        Self {
            kind: Kind::from_media_type(part.content_type()),
            part,
        }
    }

    /// OPC part name.
    #[inline]
    #[must_use]
    pub fn name(&self) -> &PackURI {
        self.part.partname()
    }

    /// Preserved OPC media type.
    #[inline]
    #[must_use]
    pub fn media_type(&self) -> &str {
        self.part.content_type()
    }

    /// Classified media family.
    #[inline]
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Return the raw OPC part bytes without interpreting them.
    #[inline]
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        self.part.blob()
    }
}
