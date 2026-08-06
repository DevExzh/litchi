use crate::package::{Error, Result};

/// Maximum bytes in one persisted `ExOleObjStg` record payload.
pub const MAX_STORED_BYTES: usize = 128 * 1_048_576;
/// Maximum producer-declared decompressed size for a compressed payload.
pub const MAX_DECLARED_BYTES: u32 = 256 * 1_048_576;

/// The PowerPoint record context that owns a persisted payload.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Kind {
    /// An embedded or linked OLE object.
    OleObject,
    /// The document's persisted VBA project.
    VbaProject,
    /// An ActiveX control's persisted storage.
    ActiveXControl,
}

/// The outer encoding used by one `ExOleObjStg` payload.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Compression {
    /// The stored bytes are the uncompressed payload.
    Uncompressed,
    /// The stored bytes are one complete zlib stream.
    Zlib,
}

/// One validated, inert PowerPoint persisted-storage payload.
///
/// The payload is kept opaque.  The record context, compression mode, and
/// declared decompressed size are private so callers cannot construct a
/// contradictory state or mutate bytes after validation.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Storage {
    pub(super) kind: Kind,
    pub(super) compression: Compression,
    pub(super) declared_uncompressed_len: Option<u32>,
    /// Raw uncompressed bytes, or raw zlib bytes without the size prefix.
    pub(super) data: Vec<u8>,
}

/// A persisted-storage record borrowed directly from `PowerPoint Document`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct Ref<'a> {
    pub(crate) kind: Kind,
    pub(crate) compression: Compression,
    pub(crate) declared_uncompressed_len: Option<u32>,
    pub(crate) data: &'a [u8],
}

impl Compression {
    /// Whether this storage requires zlib decoding.
    pub const fn is_zlib(self) -> bool {
        matches!(self, Self::Zlib)
    }

    pub(crate) const fn prefix_len(self) -> usize {
        if self.is_zlib() { 4 } else { 0 }
    }
}

impl Storage {
    /// Build an uncompressed persisted payload after checking its wire size.
    pub fn uncompressed(kind: Kind, data: Vec<u8>) -> Result<Self> {
        Self::from_parts(kind, Compression::Uncompressed, None, data)
    }

    /// Build a zlib persisted payload after checking its declared and stored
    /// wire sizes.  `data` excludes PowerPoint's four-byte size prefix.
    pub fn compressed(kind: Kind, uncompressed_len: u32, data: Vec<u8>) -> Result<Self> {
        Self::from_parts(kind, Compression::Zlib, Some(uncompressed_len), data)
    }

    pub(crate) fn from_parts(
        kind: Kind,
        compression: Compression,
        declared_uncompressed_len: Option<u32>,
        data: Vec<u8>,
    ) -> Result<Self> {
        let prefix_len = compression.prefix_len();
        let stored_len = data
            .len()
            .checked_add(prefix_len)
            .ok_or_else(|| Error::Corrupted("ExOleObjStg stored size overflows".into()))?;
        if stored_len > MAX_STORED_BYTES {
            return corrupted("ExOleObjStg exceeds 128 MiB stored data");
        }
        if let Some(declared) = declared_uncompressed_len {
            if !compression.is_zlib() {
                return corrupted("uncompressed ExOleObjStg has a compressed size");
            }
            if declared > MAX_DECLARED_BYTES {
                return corrupted("compressed ExOleObjStg declares more than 256 MiB");
            }
        } else if compression.is_zlib() {
            return corrupted("compressed ExOleObjStg is missing its size");
        }
        Ok(Self {
            kind,
            compression,
            declared_uncompressed_len,
            data,
        })
    }

    /// The record's PowerPoint reference context.
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// The outer storage encoding.
    pub const fn compression(&self) -> Compression {
        self.compression
    }

    /// The producer-declared decompressed size, when zlib is used.
    pub const fn declared_uncompressed_len(&self) -> Option<u32> {
        self.declared_uncompressed_len
    }

    /// Borrow the stored payload, excluding the compressed size prefix.
    pub fn stored_bytes(&self) -> &[u8] {
        &self.data
    }

    /// Return the stored payload length without allocating.
    pub const fn stored_payload_len(&self) -> usize {
        self.data.len()
    }

    /// Consume the storage and return its stored payload allocation.
    pub fn into_stored_bytes(self) -> Vec<u8> {
        self.data
    }
}

fn corrupted<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::Corrupted(message.into()))
}
