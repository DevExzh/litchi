//! Typed, inert `[MS-OLEDS]` `CompObj` metadata.

use std::sync::Arc;

/// A clipboard format declared by a `CompObj` stream.
#[derive(Clone, Debug, PartialEq, Eq)]
pub enum Clipboard {
    /// The stream does not declare a clipboard format.
    None,
    /// A standard clipboard format identifier.
    Standard(u32),
    /// A registered clipboard format name.
    Registered(String),
}

/// Parsed, inert metadata from a `\x01CompObj` stream.
///
/// The object never activates a server, loads a class, or interprets the
/// embedded payload. The complete source stream remains available through
/// [`Self::bytes`], including arbitrary reserved and trailing bytes.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct CompObj {
    pub(crate) reserved1: u32,
    pub(crate) version: u32,
    pub(crate) reserved2: [u8; 20],
    pub(crate) ansi_user_type: String,
    pub(crate) ansi_clipboard: Clipboard,
    pub(crate) reserved_ansi_present: bool,
    pub(crate) unicode_marker: Option<u32>,
    pub(crate) unicode_user_type: Option<String>,
    pub(crate) unicode_clipboard: Option<Clipboard>,
    pub(crate) reserved_unicode_present: bool,
    pub(crate) raw: Arc<[u8]>,
    pub(crate) trailing_start: usize,
}

impl CompObj {
    /// The arbitrary four-byte `Reserved1` value from the `CompObjHeader`.
    #[must_use]
    pub const fn reserved1(&self) -> u32 {
        self.reserved1
    }

    /// The arbitrary four-byte `Version` value from the `CompObjHeader`.
    #[must_use]
    pub const fn version(&self) -> u32 {
        self.version
    }

    /// The arbitrary twenty-byte `Reserved2` value from the header.
    #[must_use]
    pub const fn reserved2(&self) -> &[u8; 20] {
        &self.reserved2
    }

    /// The ANSI display name, decoded with the bounded Windows-1252 view.
    #[must_use]
    pub fn ansi_user_type(&self) -> &str {
        &self.ansi_user_type
    }

    /// The ANSI clipboard format declaration.
    #[must_use]
    pub const fn ansi_clipboard(&self) -> &Clipboard {
        &self.ansi_clipboard
    }

    /// Whether the optional ANSI reserved string was present.
    #[must_use]
    pub const fn has_reserved_ansi(&self) -> bool {
        self.reserved_ansi_present
    }

    /// The Unicode marker, when the optional Unicode section was present.
    #[must_use]
    pub const fn unicode_marker(&self) -> Option<u32> {
        self.unicode_marker
    }

    /// The Unicode display name, when the Unicode section was valid and
    /// present.
    #[must_use]
    pub fn unicode_user_type(&self) -> Option<&str> {
        self.unicode_user_type.as_deref()
    }

    /// The Unicode clipboard format declaration, when present.
    #[must_use]
    pub const fn unicode_clipboard(&self) -> Option<&Clipboard> {
        self.unicode_clipboard.as_ref()
    }

    /// Whether the optional Unicode reserved string was present.
    #[must_use]
    pub const fn has_reserved_unicode(&self) -> bool {
        self.reserved_unicode_present
    }

    /// Bytes after the last successfully decoded field.
    ///
    /// These bytes are never interpreted or discarded. The full stream is
    /// also available through [`Self::bytes`].
    #[must_use]
    pub fn trailing(&self) -> &[u8] {
        self.raw.get(self.trailing_start..).unwrap_or_default()
    }

    /// Exact source bytes of the `\x01CompObj` stream.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.raw
    }

    /// Shared ownership of the exact source stream allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.raw)
    }

    pub(in crate::embedded_object) fn from_parts(parts: Parts) -> Self {
        Self {
            reserved1: parts.reserved1,
            version: parts.version,
            reserved2: parts.reserved2,
            ansi_user_type: parts.ansi_user_type,
            ansi_clipboard: parts.ansi_clipboard,
            reserved_ansi_present: parts.reserved_ansi_present,
            unicode_marker: parts.unicode_marker,
            unicode_user_type: parts.unicode_user_type,
            unicode_clipboard: parts.unicode_clipboard,
            reserved_unicode_present: parts.reserved_unicode_present,
            raw: parts.raw,
            trailing_start: parts.trailing_start,
        }
    }
}

pub(in crate::embedded_object) struct Parts {
    pub(in crate::embedded_object) reserved1: u32,
    pub(in crate::embedded_object) version: u32,
    pub(in crate::embedded_object) reserved2: [u8; 20],
    pub(in crate::embedded_object) ansi_user_type: String,
    pub(in crate::embedded_object) ansi_clipboard: Clipboard,
    pub(in crate::embedded_object) reserved_ansi_present: bool,
    pub(in crate::embedded_object) unicode_marker: Option<u32>,
    pub(in crate::embedded_object) unicode_user_type: Option<String>,
    pub(in crate::embedded_object) unicode_clipboard: Option<Clipboard>,
    pub(in crate::embedded_object) reserved_unicode_present: bool,
    pub(in crate::embedded_object) raw: Arc<[u8]>,
    pub(in crate::embedded_object) trailing_start: usize,
}
