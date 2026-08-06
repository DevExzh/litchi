//! Lossless typed projections of the CFB directory entry fields exposed by
//! the common object owner.

use crate::property_set::Guid;
use litchi_cfb::OleError;

/// The CFB terminator used for an empty sibling or child link.
pub const NOSTREAM: u32 = u32::MAX;

/// The greatest regular CFB directory identifier accepted by `[MS-CFB]`.
pub const MAX_REGULAR_SID: u32 = 0xFFFF_FFFA;

/// Resource bounds for a captured directory catalog.
///
/// The catalog owns only directory metadata.  It never reads or retains CFB
/// sectors, but names and producer-supplied child projections are still
/// bounded before they are cloned into a source-preserving snapshot.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum number of top-level directory entries in the catalog.
    pub max_entries: usize,
    /// Maximum UTF-8 byte length of one retained directory name or CLSID.
    pub max_name_bytes: usize,
    /// Maximum aggregate UTF-8 bytes retained by names and raw CLSIDs.
    pub max_total_bytes: usize,
    /// Maximum number of recursively supplied raw child entries.
    pub max_raw_children: usize,
    /// Maximum nesting depth of recursively supplied raw child entries.
    pub max_raw_depth: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_entries: 65_536,
            max_name_bytes: 256,
            max_total_bytes: 64 * 1024 * 1024,
            max_raw_children: 65_536,
            max_raw_depth: 64,
        }
    }
}

impl Limits {
    pub(crate) fn validate(self) -> Result<(), OleError> {
        if self.max_entries == 0
            || self.max_name_bytes == 0
            || self.max_total_bytes == 0
            || self.max_raw_depth == 0
        {
            return Err(OleError::InvalidFormat(
                "CFB directory catalog limits must be non-zero".into(),
            ));
        }
        Ok(())
    }
}

/// A checked CFB directory stream identifier.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
pub struct Sid(u32);

impl Sid {
    /// Creates a directory identifier from its wire value.
    ///
    /// The `NOSTREAM` terminator is represented by `None` in [`Links`] and is
    /// therefore not a valid `Sid`.
    pub fn new(raw: u32) -> Result<Self, OleError> {
        if raw <= MAX_REGULAR_SID {
            Ok(Self(raw))
        } else {
            Err(OleError::InvalidFormat(format!(
                "CFB directory SID {raw:#010X} is outside the regular SID range"
            )))
        }
    }

    /// The original unsigned CFB stream identifier.
    #[must_use]
    pub const fn raw(self) -> u32 {
        self.0
    }

    pub(crate) const fn from_checked(raw: u32) -> Self {
        Self(raw)
    }
}

/// The valid object kinds stored in a CFB directory entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum EntryKind {
    /// A child storage object.
    Storage,
    /// A byte stream object.
    Stream,
    /// The SID-zero root storage object.
    Root,
}

impl EntryKind {
    /// The `[MS-CFB]` wire value for this object kind.
    #[must_use]
    pub const fn raw(self) -> u8 {
        match self {
            Self::Storage => 0x01,
            Self::Stream => 0x02,
            Self::Root => 0x05,
        }
    }
}

/// The red-black-tree links carried by a CFB directory entry.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Default)]
pub struct Links {
    left: Option<Sid>,
    right: Option<Sid>,
    child: Option<Sid>,
}

impl Links {
    /// Creates checked containment links from optional directory SIDs.
    #[must_use]
    pub const fn new(left: Option<Sid>, right: Option<Sid>, child: Option<Sid>) -> Self {
        Self { left, right, child }
    }

    /// The left sibling, if present.
    #[must_use]
    pub const fn left(self) -> Option<Sid> {
        self.left
    }

    /// The right sibling, if present.
    #[must_use]
    pub const fn right(self) -> Option<Sid> {
        self.right
    }

    /// The first child tree node, if present.
    #[must_use]
    pub const fn child(self) -> Option<Sid> {
        self.child
    }

    /// Returns these links with a different left sibling.
    #[must_use]
    pub const fn with_left(self, left: Option<Sid>) -> Self {
        Self { left, ..self }
    }

    /// Returns these links with a different right sibling.
    #[must_use]
    pub const fn with_right(self, right: Option<Sid>) -> Self {
        Self { right, ..self }
    }

    /// Returns these links with a different first child.
    #[must_use]
    pub const fn with_child(self, child: Option<Sid>) -> Self {
        Self { child, ..self }
    }

    pub(crate) const fn from_raw(
        left: Option<Sid>,
        right: Option<Sid>,
        child: Option<Sid>,
    ) -> Self {
        Self { left, right, child }
    }
}

/// Checked directory metadata for one captured CFB object.
///
/// `start_sector`, `stream_size`, and `uses_mini_stream` are meaningful for a
/// stream.  For a storage they retain the parsed CFB values and are normally
/// zero; for the root they describe the mini-stream as specified by
/// `[MS-CFB]`.  Metadata is `Copy`, so exposing it never clones payload bytes.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Metadata {
    sid: Sid,
    kind: EntryKind,
    class_id: Option<Guid>,
    links: Links,
    start_sector: u32,
    stream_size: u64,
    uses_mini_stream: bool,
}

impl Metadata {
    pub(crate) const fn new(
        sid: Sid,
        kind: EntryKind,
        class_id: Option<Guid>,
        links: Links,
        start_sector: u32,
        stream_size: u64,
        uses_mini_stream: bool,
    ) -> Self {
        Self {
            sid,
            kind,
            class_id,
            links,
            start_sector,
            stream_size,
            uses_mini_stream,
        }
    }

    pub(crate) const fn with_class_id(self, class_id: Option<Guid>) -> Self {
        Self { class_id, ..self }
    }

    /// Replaces the storage or root class identifier in a transaction draft.
    pub const fn set_class_id(&mut self, class_id: Option<Guid>) -> &mut Self {
        self.class_id = class_id;
        self
    }

    /// Replaces the object kind in a transaction draft.
    pub const fn set_kind(&mut self, kind: EntryKind) -> &mut Self {
        self.kind = kind;
        self
    }

    /// Replaces the red-black-tree containment links in a transaction draft.
    pub const fn set_links(&mut self, links: Links) -> &mut Self {
        self.links = links;
        self
    }

    /// Replaces the starting FAT or MiniFAT sector in a transaction draft.
    pub const fn set_start_sector(&mut self, start_sector: u32) -> &mut Self {
        self.start_sector = start_sector;
        self
    }

    /// Replaces the parsed CFB stream size in a transaction draft.
    pub const fn set_stream_size(&mut self, stream_size: u64) -> &mut Self {
        self.stream_size = stream_size;
        self
    }

    /// Replaces MiniFAT placement in a transaction draft.
    pub const fn set_uses_mini_stream(&mut self, uses_mini_stream: bool) -> &mut Self {
        self.uses_mini_stream = uses_mini_stream;
        self
    }

    pub(crate) const fn staged_storage(class_id: Option<Guid>) -> Self {
        Self::new(
            Sid::from_checked(0),
            EntryKind::Storage,
            class_id,
            Links::from_raw(None, None, None),
            0,
            0,
            false,
        )
    }

    /// The physical directory identifier assigned by the CFB file.
    #[must_use]
    pub const fn sid(self) -> Sid {
        self.sid
    }

    /// The typed CFB object kind.
    #[must_use]
    pub const fn kind(self) -> EntryKind {
        self.kind
    }

    /// The storage/root class identifier, when one was stored.
    #[must_use]
    pub const fn class_id(self) -> Option<Guid> {
        self.class_id
    }

    /// The raw red-black-tree containment links.
    #[must_use]
    pub const fn links(self) -> Links {
        self.links
    }

    /// The starting FAT or MiniFAT sector location.
    #[must_use]
    pub const fn start_sector(self) -> u32 {
        self.start_sector
    }

    /// The parsed CFB stream size.
    #[must_use]
    pub const fn stream_size(self) -> u64 {
        self.stream_size
    }

    /// Whether the stream payload is stored through the MiniFAT.
    #[must_use]
    pub const fn uses_mini_stream(self) -> bool {
        self.uses_mini_stream
    }
}
