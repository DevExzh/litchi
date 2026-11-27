//! Semantic OLE metadata for the DOC writer.

/// A 16-byte OLE class identifier in the byte order used by OLE streams.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
pub struct ClassId([u8; 16]);

impl ClassId {
    /// The class identifier for `Word.Document.8`.
    pub const WORD_DOCUMENT: Self = Self([
        0x06, 0x09, 0x02, 0x00, // Data1
        0x00, 0x00, // Data2
        0x00, 0x00, // Data3
        0xC0, 0x00, // Data4[0..2]
        0x00, 0x00, 0x00, 0x00, 0x00, 0x46, // Data4[2..8]
    ]);

    /// Construct a class identifier from its OLE-ordered bytes.
    #[must_use]
    pub const fn from_bytes(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    /// Borrow the OLE-ordered class identifier bytes.
    #[must_use]
    pub const fn as_bytes(&self) -> &[u8; 16] {
        &self.0
    }
}

/// The semantic contents of a DOC `\x01CompObj` stream.
///
/// The writer profile intentionally uses borrowed strings and a compact class
/// identifier, so constructing the default Word metadata performs no heap
/// allocation. The fields are private to keep the stream's required metadata
/// combination valid.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CompObj {
    class_id: ClassId,
    user_type: &'static str,
    clipboard_format: &'static str,
    prog_id: &'static str,
}

impl CompObj {
    /// Construct the canonical metadata for a Word 97–2003 document.
    #[must_use]
    pub const fn word_document() -> Self {
        Self {
            class_id: ClassId::WORD_DOCUMENT,
            user_type: "Microsoft Word Document",
            clipboard_format: "MSWordDoc",
            prog_id: "Word.Document.8",
        }
    }

    /// Return the embedded object's class identifier.
    #[must_use]
    pub const fn class_id(&self) -> ClassId {
        self.class_id
    }

    /// Return the ANSI user-type display name.
    #[must_use]
    pub const fn user_type(&self) -> &'static str {
        self.user_type
    }

    /// Return the clipboard format name.
    #[must_use]
    pub const fn clipboard_format(&self) -> &'static str {
        self.clipboard_format
    }

    /// Return the programmatic identifier.
    #[must_use]
    pub const fn prog_id(&self) -> &'static str {
        self.prog_id
    }
}

impl Default for CompObj {
    fn default() -> Self {
        Self::word_document()
    }
}

/// The semantic contents of a DOC `\x01Ole` stream.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Ole {
    version: u32,
}

impl Ole {
    /// Construct the canonical OLE 2.1 version profile emitted by Word.
    #[must_use]
    pub const fn word_document() -> Self {
        Self {
            version: 0x0200_0001,
        }
    }

    /// Return the raw little-endian OLE version value.
    #[must_use]
    pub const fn version(&self) -> u32 {
        self.version
    }
}

impl Default for Ole {
    fn default() -> Self {
        Self::word_document()
    }
}

/// The complete fixed metadata profile written into a DOC compound file.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct Metadata {
    comp_obj: CompObj,
    ole: Ole,
}

impl Metadata {
    /// Construct the canonical metadata pair for a Word document.
    #[must_use]
    pub const fn word_document() -> Self {
        Self {
            comp_obj: CompObj::word_document(),
            ole: Ole::word_document(),
        }
    }

    /// Return the semantic `\x01CompObj` contents.
    #[must_use]
    pub const fn comp_obj(&self) -> CompObj {
        self.comp_obj
    }

    /// Return the semantic `\x01Ole` contents.
    #[must_use]
    pub const fn ole(&self) -> Ole {
        self.ole
    }
}
