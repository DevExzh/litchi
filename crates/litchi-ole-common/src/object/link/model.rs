//! Semantic OLEDS object-link values.

use crate::property_set::Guid;
use litchi_cfb::OleError;
use std::ops::Range;
use std::sync::Arc;

use super::codec;

/// The object state encoded by the low bit of an OLEDS link stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// The storage contains an embedded object.
    Embedded,
    /// The storage refers to an object outside the containing document.
    Linked,
}

impl Kind {
    pub(crate) const fn from_flags(flags: u32) -> Self {
        if flags & codec::LINKED_FLAG != 0 {
            Self::Linked
        } else {
            Self::Embedded
        }
    }

    /// Whether this state references an external object.
    #[must_use]
    pub const fn is_linked(self) -> bool {
        matches!(self, Self::Linked)
    }

    /// Whether this state contains an embedded object.
    #[must_use]
    pub const fn is_embedded(self) -> bool {
        matches!(self, Self::Embedded)
    }
}

/// Three OLEDS FILETIME values carried by a linked object.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Times {
    local_update: u64,
    local_check: u64,
    remote_update: u64,
}

impl Times {
    /// Creates a timestamp group from raw Windows FILETIME counters.
    #[must_use]
    pub const fn new(local_update: u64, local_check: u64, remote_update: u64) -> Self {
        Self {
            local_update,
            local_check,
            remote_update,
        }
    }

    /// The time when the container last updated the remote timestamp.
    #[must_use]
    pub const fn local_update(self) -> u64 {
        self.local_update
    }

    /// The time when the container last checked the remote object.
    #[must_use]
    pub const fn local_check(self) -> u64 {
        self.local_check
    }

    /// The last update time reported by the linked object.
    #[must_use]
    pub const fn remote_update(self) -> u64 {
        self.remote_update
    }
}

/// A bounded view of one OLEDS `MONIKERSTREAM`.
///
/// The class-specific moniker payload remains opaque and is never resolved.
/// The view borrows the parsed link allocation, so inspecting it does not
/// allocate or copy the reference bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Moniker<'a> {
    bytes: &'a [u8],
}

impl<'a> Moniker<'a> {
    pub(crate) const fn new(bytes: &'a [u8]) -> Self {
        Self { bytes }
    }

    /// The packetized moniker class identifier.
    #[must_use]
    pub fn class_id(self) -> Guid {
        let mut bytes = [0; 16];
        bytes.copy_from_slice(&self.bytes[..16]);
        Guid::from_bytes(bytes)
    }

    /// The implementation-specific moniker data after its class identifier.
    #[must_use]
    pub fn data(self) -> &'a [u8] {
        &self.bytes[16..]
    }

    /// The complete packetized moniker bytes.
    #[must_use]
    pub const fn bytes(self) -> &'a [u8] {
        self.bytes
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Link {
    pub(crate) wire: Arc<[u8]>,
    pub(crate) kind: Kind,
    pub(crate) flags: u32,
    pub(crate) update_option: u32,
    pub(crate) reserved_moniker: Option<Range<usize>>,
    pub(crate) relative_source: Option<Range<usize>>,
    pub(crate) absolute_source: Option<Range<usize>>,
    pub(crate) class_id: Option<Guid>,
    pub(crate) class_id_offset: Option<usize>,
    pub(crate) reserved_display: Option<Range<usize>>,
    pub(crate) reserved2: Option<u32>,
    pub(crate) times: Option<Times>,
    pub(crate) times_offsets: Option<[usize; 3]>,
    pub(crate) tail_offset: usize,
}

impl Link {
    /// OLEDS version required by the `\x01Ole` stream.
    pub const VERSION: u32 = 0x0200_0001;

    /// Parses a link stream and retains a private copy of its bytes.
    ///
    /// Use [`Self::parse_shared`] when the caller already owns the stream in
    /// an [`Arc`], as the object layer does.
    ///
    /// # Errors
    ///
    /// Returns an error when the stream is malformed or exceeds the OLEDS
    /// metadata limit.
    pub fn parse(bytes: &[u8]) -> Result<Self, OleError> {
        codec::parse(Arc::<[u8]>::from(bytes))
    }

    /// Parses a link stream without copying an existing allocation.
    ///
    /// # Errors
    ///
    /// Returns an error when the stream is malformed or exceeds the OLEDS
    /// metadata limit.
    pub fn parse_shared(bytes: Arc<[u8]>) -> Result<Self, OleError> {
        codec::parse(bytes)
    }

    /// The exact source bytes retained by this value.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.wire
    }

    /// Shared ownership of the exact source allocation.
    #[must_use]
    pub fn bytes_shared(&self) -> Arc<[u8]> {
        Arc::clone(&self.wire)
    }

    /// Whether the object is embedded or linked.
    #[must_use]
    pub const fn kind(&self) -> Kind {
        self.kind
    }

    /// The complete raw Flags field, including future or producer-specific
    /// bits that this crate does not interpret.
    #[must_use]
    pub const fn flags(&self) -> u32 {
        self.flags
    }

    /// Whether the OLEDS cache hint is set.
    #[must_use]
    pub const fn cache_hint(&self) -> bool {
        self.flags & codec::CACHE_HINT_FLAG != 0
    }

    /// Updates the cache hint while preserving all other flag bits.
    pub const fn set_cache_hint(&mut self, enabled: bool) {
        if enabled {
            self.flags |= codec::CACHE_HINT_FLAG;
        } else {
            self.flags &= !codec::CACHE_HINT_FLAG;
        }
    }

    /// Replaces the raw Flags field without changing embedded/linked state.
    ///
    /// The low bit controls the wire layout and therefore cannot be changed
    /// after parsing.  Unknown higher bits are retained for forward
    /// compatibility and round-trip losslessness.
    ///
    /// # Errors
    ///
    /// Returns an error when `flags` changes the embedded/linked layout.
    pub fn set_flags(&mut self, flags: u32) -> Result<(), OleError> {
        if (flags & codec::LINKED_FLAG) != (self.flags & codec::LINKED_FLAG) {
            return Err(OleError::InvalidFormat(
                "OLE link kind cannot change without rebuilding its wire layout".into(),
            ));
        }
        self.flags = flags;
        Ok(())
    }

    /// The implementation-specific OLEDS update option.
    #[must_use]
    pub const fn link_update_option(&self) -> u32 {
        self.update_option
    }

    /// Replaces the implementation-specific OLEDS update option.
    pub const fn set_link_update_option(&mut self, value: u32) {
        self.update_option = value;
    }

    /// The reserved moniker bytes, when the producer supplied them.
    #[must_use]
    pub fn reserved_moniker(&self) -> Option<&[u8]> {
        self.reserved_moniker
            .as_ref()
            .map(|range| &self.wire[range.clone()])
    }

    /// The relative source reference, when supplied by a linked object.
    #[must_use]
    pub fn relative_source(&self) -> Option<Moniker<'_>> {
        self.relative_source
            .as_ref()
            .map(|range| Moniker::new(&self.wire[range.clone()]))
    }

    /// The absolute source reference, when supplied by a linked object.
    #[must_use]
    pub fn absolute_source(&self) -> Option<Moniker<'_>> {
        self.absolute_source
            .as_ref()
            .map(|range| Moniker::new(&self.wire[range.clone()]))
    }

    /// The source selected by OLEDS: relative first, absolute otherwise.
    #[must_use]
    pub fn source(&self) -> Option<Moniker<'_>> {
        self.relative_source().or_else(|| self.absolute_source())
    }

    /// The linked object's class identifier from the stream, when present.
    #[must_use]
    pub const fn class_id(&self) -> Option<Guid> {
        self.class_id
    }

    /// Replaces the linked object's class identifier.
    ///
    /// # Errors
    ///
    /// Returns an error when this is an embedded link with no class identifier
    /// field in its wire layout.
    pub fn set_class_id(&mut self, value: Guid) -> Result<(), OleError> {
        if self.class_id_offset.is_none() {
            return Err(OleError::InvalidFormat(
                "embedded OLE links do not carry a stream class identifier".into(),
            ));
        }
        self.class_id = Some(value);
        Ok(())
    }

    /// The reserved Unicode display-name bytes, excluding their length field.
    #[must_use]
    pub fn reserved_display_name(&self) -> Option<&[u8]> {
        self.reserved_display
            .as_ref()
            .map(|range| &self.wire[range.clone()])
    }

    /// The uninterpreted Reserved2 value, when the linked tail is present.
    #[must_use]
    pub const fn reserved2(&self) -> Option<u32> {
        self.reserved2
    }

    /// The linked-object FILETIME group, when present.
    #[must_use]
    pub const fn times(&self) -> Option<Times> {
        self.times
    }

    /// Replaces the linked-object FILETIME group.
    ///
    /// # Errors
    ///
    /// Returns an error when this link has no timestamp group in its wire
    /// layout.
    pub fn set_times(&mut self, value: Times) -> Result<(), OleError> {
        if self.times_offsets.is_none() {
            return Err(OleError::InvalidFormat(
                "OLE link has no linked-object timestamp group".into(),
            ));
        }
        self.times = Some(value);
        Ok(())
    }

    /// Bytes after the currently understood OLEDS fields.
    #[must_use]
    pub fn unknown_tail(&self) -> &[u8] {
        &self.wire[self.tail_offset..]
    }

    /// Serializes the typed edits over the original wire layout.
    ///
    /// Unknown flags, reserved fields, monikers, display bytes, and trailing
    /// data are copied from the original stream unchanged.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        let mut output = self.wire.as_ref().to_vec();
        output[4..8].copy_from_slice(&self.flags.to_le_bytes());
        output[8..12].copy_from_slice(&self.update_option.to_le_bytes());
        if let (Some(offset), Some(class_id)) = (self.class_id_offset, self.class_id) {
            output[offset..offset + 16].copy_from_slice(class_id.as_bytes());
        }
        if let (Some(offsets), Some(times)) = (self.times_offsets, self.times) {
            for (offset, value) in offsets.into_iter().zip([
                times.local_update(),
                times.local_check(),
                times.remote_update(),
            ]) {
                output[offset..offset + 8].copy_from_slice(&value.to_le_bytes());
            }
        }
        output
    }
}
