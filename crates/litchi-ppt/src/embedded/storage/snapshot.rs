//! Borrowed, lossless projections of one PowerPoint persisted storage.

use super::model::{Compression, Kind, Ref, Storage};

/// The payload-free metadata projection of one `ExOleObjStg` record.
///
/// The projection retains every variable scalar from the PowerPoint storage
/// envelope. Use [`Snapshot::stored_bytes`] when the opaque OLE2 bytes are
/// needed; no embedded content is opened or executed by this view.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Metadata {
    pub(crate) kind: Kind,
    pub(crate) compression: Compression,
    pub(crate) declared_uncompressed_len: Option<u32>,
    pub(crate) stored_payload_len: usize,
    pub(crate) record_payload_len: usize,
    pub(crate) record_len: usize,
    pub(crate) contains_data: bool,
}

impl Metadata {
    /// The PowerPoint record context that owns the payload.
    #[must_use]
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// The outer `ExOleObjStg` encoding.
    #[must_use]
    pub const fn compression(self) -> Compression {
        self.compression
    }

    /// The producer-declared decompressed size, when the payload is zlib
    /// encoded.
    #[must_use]
    pub const fn declared_uncompressed_len(self) -> Option<u32> {
        self.declared_uncompressed_len
    }

    /// The stored payload length after the compressed-size prefix.
    #[must_use]
    pub const fn stored_payload_len(self) -> usize {
        self.stored_payload_len
    }

    /// The payload length recorded by `rh.recLen`, including the compressed
    /// size prefix when present.
    #[must_use]
    pub const fn record_payload_len(self) -> usize {
        self.record_payload_len
    }

    /// The complete encoded record length, including its eight-byte header.
    #[must_use]
    pub const fn record_len(self) -> usize {
        self.record_len
    }

    /// Whether the producer-declared storage contains payload data.
    #[must_use]
    pub const fn contains_data(self) -> bool {
        self.contains_data
    }
}

/// A zero-copy immutable view over one validated `ExOleObjStg` payload.
///
/// The view is bounded by the same 128 MiB stored and 256 MiB declared limits
/// as [`Storage`]. It carries the raw stored representation, so serializing it
/// preserves the compressed bytes exactly. The lifetime ties the view to the
/// source snapshot and prevents mutation behind it.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Snapshot<'a> {
    kind: Kind,
    compression: Compression,
    declared_uncompressed_len: Option<u32>,
    stored_bytes: &'a [u8],
}

impl<'a> Snapshot<'a> {
    pub(super) const fn from_parts(
        kind: Kind,
        compression: Compression,
        declared_uncompressed_len: Option<u32>,
        stored_bytes: &'a [u8],
    ) -> Self {
        Self {
            kind,
            compression,
            declared_uncompressed_len,
            stored_bytes,
        }
    }

    /// The PowerPoint record context that owns the payload.
    #[must_use]
    pub const fn kind(self) -> Kind {
        self.kind
    }

    /// The outer `ExOleObjStg` encoding.
    #[must_use]
    pub const fn compression(self) -> Compression {
        self.compression
    }

    /// The producer-declared decompressed size, when the payload is zlib
    /// encoded.
    #[must_use]
    pub const fn declared_uncompressed_len(self) -> Option<u32> {
        self.declared_uncompressed_len
    }

    /// Borrow the exact stored payload, excluding the compressed-size prefix.
    #[must_use]
    pub const fn stored_bytes(self) -> &'a [u8] {
        self.stored_bytes
    }

    /// The stored payload length without the compressed-size prefix.
    #[must_use]
    pub const fn stored_payload_len(self) -> usize {
        self.stored_bytes.len()
    }

    /// The payload length recorded by `rh.recLen`, including any size prefix.
    #[must_use]
    pub const fn record_payload_len(self) -> usize {
        self.stored_bytes
            .len()
            .saturating_add(self.compression.prefix_len())
    }

    /// The complete encoded record length, including its eight-byte header.
    #[must_use]
    pub const fn record_len(self) -> usize {
        self.record_payload_len().saturating_add(8)
    }

    /// Return the validated scalar projection without copying the OLE2 bytes.
    #[must_use]
    pub const fn metadata(self) -> Metadata {
        Metadata {
            kind: self.kind,
            compression: self.compression,
            declared_uncompressed_len: self.declared_uncompressed_len,
            stored_payload_len: self.stored_payload_len(),
            record_payload_len: self.record_payload_len(),
            record_len: self.record_len(),
            contains_data: self.contains_data(),
        }
    }

    /// Whether the producer-declared storage contains payload data.
    #[must_use]
    pub const fn contains_data(self) -> bool {
        match self.compression {
            Compression::Uncompressed => !self.stored_bytes.is_empty(),
            Compression::Zlib => match self.declared_uncompressed_len {
                Some(length) => length != 0,
                None => false,
            },
        }
    }

    /// Copy this bounded view into an owned storage snapshot.
    pub fn to_storage(self) -> crate::package::Result<Storage> {
        Storage::from_parts(
            self.kind,
            self.compression,
            self.declared_uncompressed_len,
            self.stored_bytes.to_vec(),
        )
    }
}

impl Storage {
    /// Borrow a lossless, immutable view of this storage without copying its
    /// OLE2 payload.
    #[must_use]
    pub fn snapshot(&self) -> Snapshot<'_> {
        Snapshot::from_parts(
            self.kind,
            self.compression,
            self.declared_uncompressed_len,
            &self.data,
        )
    }

    /// Return the validated payload-free metadata projection.
    #[must_use]
    pub fn metadata(&self) -> Metadata {
        self.snapshot().metadata()
    }
}

#[allow(dead_code)]
impl<'a> Ref<'a> {
    pub(crate) const fn snapshot(self) -> Snapshot<'a> {
        Snapshot::from_parts(
            self.kind,
            self.compression,
            self.declared_uncompressed_len,
            self.data,
        )
    }
}
