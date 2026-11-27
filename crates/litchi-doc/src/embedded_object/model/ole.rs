//! Typed, inert `[MS-OLEDS]` `OLEStream` metadata.

use std::ops::Range;
use std::sync::Arc;

/// The storage role declared by an `OLEStream`.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Kind {
    /// The storage contains an embedded object.
    Embedded,
    /// The storage references an external linked object.
    Linked,
}

impl Kind {
    pub(crate) const fn is_linked(self) -> bool {
        matches!(self, Self::Linked)
    }
}

/// Parsed, inert metadata from a `\x01Ole` stream.
///
/// Monikers and ignored display-name data are exposed only as borrowed raw
/// bytes. No link is followed and no external resource is accessed.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Ole {
    pub(crate) version: u32,
    pub(crate) flags: u32,
    pub(crate) kind: Kind,
    pub(crate) cache_storage: bool,
    pub(crate) update_option: u32,
    pub(crate) reserved: u32,
    pub(crate) reserved_moniker: Option<Range<usize>>,
    pub(crate) relative_moniker: Option<Range<usize>>,
    pub(crate) absolute_moniker: Option<Range<usize>>,
    pub(crate) class_id: Option<[u8; 16]>,
    pub(crate) reserved_display_name: Option<Range<usize>>,
    pub(crate) reserved2: Option<u32>,
    pub(crate) local_update_time: Option<u64>,
    pub(crate) local_check_update_time: Option<u64>,
    pub(crate) remote_update_time: Option<u64>,
    pub(crate) raw: Arc<[u8]>,
    pub(crate) trailing_start: usize,
}

impl Ole {
    /// The required OLE stream version.
    #[must_use]
    pub const fn version(&self) -> u32 {
        self.version
    }

    /// The raw OLE stream flags, including implementation-defined bits.
    #[must_use]
    pub const fn flags(&self) -> u32 {
        self.flags
    }

    /// The declared storage role.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// Whether the implementation-defined cache-storage hint is set.
    #[must_use]
    pub const fn cache_storage(&self) -> bool {
        self.cache_storage
    }

    /// The implementation-defined link update option.
    #[must_use]
    pub const fn update_option(&self) -> u32 {
        self.update_option
    }

    /// The required zero reserved field.
    #[must_use]
    pub const fn reserved(&self) -> u32 {
        self.reserved
    }

    /// The ignored reserved moniker payload, when present.
    #[must_use]
    pub fn reserved_moniker(&self) -> Option<&[u8]> {
        self.range(self.reserved_moniker.as_ref())
    }

    /// The relative link moniker, when present.
    #[must_use]
    pub fn relative_moniker(&self) -> Option<&[u8]> {
        self.range(self.relative_moniker.as_ref())
    }

    /// The absolute link moniker, when present.
    #[must_use]
    pub fn absolute_moniker(&self) -> Option<&[u8]> {
        self.range(self.absolute_moniker.as_ref())
    }

    /// The creating application's packetized CLSID, when present.
    #[must_use]
    pub const fn class_id(&self) -> Option<&[u8; 16]> {
        self.class_id.as_ref()
    }

    /// The ignored reserved display-name payload, when present.
    #[must_use]
    pub fn reserved_display_name(&self) -> Option<&[u8]> {
        self.range(self.reserved_display_name.as_ref())
    }

    /// The arbitrary optional reserved value after the display name.
    #[must_use]
    pub const fn reserved2(&self) -> Option<u32> {
        self.reserved2
    }

    /// The local update timestamp, when the linked-object tail is present.
    #[must_use]
    pub const fn local_update_time(&self) -> Option<u64> {
        self.local_update_time
    }

    /// The local check timestamp, when the linked-object tail is present.
    #[must_use]
    pub const fn local_check_update_time(&self) -> Option<u64> {
        self.local_check_update_time
    }

    /// The remote update timestamp, when the linked-object tail is present.
    #[must_use]
    pub const fn remote_update_time(&self) -> Option<u64> {
        self.remote_update_time
    }

    /// Bytes after the last successfully decoded field.
    #[must_use]
    pub fn trailing(&self) -> &[u8] {
        self.raw.get(self.trailing_start..).unwrap_or_default()
    }

    /// Exact source bytes of the `\x01Ole` stream.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.raw
    }

    /// Shared ownership of the exact source stream allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.raw)
    }

    fn range(&self, range: Option<&Range<usize>>) -> Option<&[u8]> {
        range.and_then(|range| self.raw.get(range.clone()))
    }

    pub(in crate::embedded_object) fn from_parts(parts: Parts) -> Self {
        Self {
            version: parts.version,
            flags: parts.flags,
            kind: parts.kind,
            cache_storage: parts.cache_storage,
            update_option: parts.update_option,
            reserved: parts.reserved,
            reserved_moniker: parts.reserved_moniker,
            relative_moniker: parts.relative_moniker,
            absolute_moniker: parts.absolute_moniker,
            class_id: parts.class_id,
            reserved_display_name: parts.reserved_display_name,
            reserved2: parts.reserved2,
            local_update_time: parts.local_update_time,
            local_check_update_time: parts.local_check_update_time,
            remote_update_time: parts.remote_update_time,
            raw: parts.raw,
            trailing_start: parts.trailing_start,
        }
    }
}

pub(in crate::embedded_object) struct Parts {
    pub(in crate::embedded_object) version: u32,
    pub(in crate::embedded_object) flags: u32,
    pub(in crate::embedded_object) kind: Kind,
    pub(in crate::embedded_object) cache_storage: bool,
    pub(in crate::embedded_object) update_option: u32,
    pub(in crate::embedded_object) reserved: u32,
    pub(in crate::embedded_object) reserved_moniker: Option<Range<usize>>,
    pub(in crate::embedded_object) relative_moniker: Option<Range<usize>>,
    pub(in crate::embedded_object) absolute_moniker: Option<Range<usize>>,
    pub(in crate::embedded_object) class_id: Option<[u8; 16]>,
    pub(in crate::embedded_object) reserved_display_name: Option<Range<usize>>,
    pub(in crate::embedded_object) reserved2: Option<u32>,
    pub(in crate::embedded_object) local_update_time: Option<u64>,
    pub(in crate::embedded_object) local_check_update_time: Option<u64>,
    pub(in crate::embedded_object) remote_update_time: Option<u64>,
    pub(in crate::embedded_object) raw: Arc<[u8]>,
    pub(in crate::embedded_object) trailing_start: usize,
}
