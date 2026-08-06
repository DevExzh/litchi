use crate::{Error, Result};

/// Maximum number of UTF-16 code units accepted in one OfficeArt name,
/// excluding its terminating NUL.
pub const MAX_NAME_UNITS: usize = 4096;

/// Maximum encoded byte length of one OfficeArt name, including its NUL.
pub const MAX_NAME_BYTES: usize = (MAX_NAME_UNITS + 1) * 2;

/// Maximum OfficeArt property-record body retained by a picture snapshot.
pub const MAX_SNAPSHOT_BYTES: usize = 16 * 1024 * 1024;

/// The semantic meaning of a picture name.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub enum Kind {
    /// The name is a producer comment.
    Comment,
    /// The name identifies a file.
    File,
    /// The name identifies a URL.
    Url,
}

/// Checked `[MS-ODRAW]` `MSOBLIPFLAGS` metadata.
///
/// Bits outside the specified low nibble are retained exactly as reserved
/// producer data.  The defined low-bit constraints are enforced by
/// [`Self::from_raw`] and [`Self::new`].
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct Flags(u32);

impl Flags {
    const KIND_MASK: u32 = 0x03;
    const DO_NOT_SAVE: u32 = 0x04;
    const LINK_TO_FILE: u32 = 0x08;
    const DEFINED_MASK: u32 = Self::KIND_MASK | Self::DO_NOT_SAVE | Self::LINK_TO_FILE;

    /// Creates checked flags from their semantic components.
    pub fn new(kind: Kind, link_to_file: bool, do_not_save: bool) -> Result<Self> {
        if do_not_save && !link_to_file {
            return Err(Error::MalformedProperties {
                reason: "picture do-not-save requires link-to-file",
            });
        }
        if link_to_file && matches!(kind, Kind::Comment) {
            return Err(Error::MalformedProperties {
                reason: "picture link-to-file requires a file or URL name",
            });
        }
        let kind_bits = match kind {
            Kind::Comment => 0,
            Kind::File => 1,
            Kind::Url => 2,
        };
        let raw = kind_bits
            | if link_to_file { Self::LINK_TO_FILE } else { 0 }
            | if do_not_save { Self::DO_NOT_SAVE } else { 0 };
        Ok(Self(raw))
    }

    /// Decodes flags while retaining undefined producer bits.
    pub fn from_raw(raw: u32) -> Result<Self> {
        let kind = raw & Self::KIND_MASK;
        if kind == 3 {
            return Err(Error::MalformedProperties {
                reason: "picture name kind uses a reserved flag value",
            });
        }
        let link_to_file = raw & Self::LINK_TO_FILE != 0;
        let do_not_save = raw & Self::DO_NOT_SAVE != 0;
        if do_not_save && !link_to_file {
            return Err(Error::MalformedProperties {
                reason: "picture do-not-save requires link-to-file",
            });
        }
        if link_to_file && kind == 0 {
            return Err(Error::MalformedProperties {
                reason: "picture link-to-file requires a file or URL name",
            });
        }
        Ok(Self(raw))
    }

    /// Returns the exact 32-bit flags value, including reserved bits.
    pub const fn raw(self) -> u32 {
        self.0
    }

    /// Returns the semantic name kind.
    pub const fn kind(self) -> Kind {
        match self.0 & Self::KIND_MASK {
            1 => Kind::File,
            2 => Kind::Url,
            _ => Kind::Comment,
        }
    }

    /// Returns whether the BLIP is linked to a file or URL.
    pub const fn link_to_file(self) -> bool {
        self.0 & Self::LINK_TO_FILE != 0
    }

    /// Returns whether the BLIP must not be embedded on save.
    pub const fn do_not_save(self) -> bool {
        self.0 & Self::DO_NOT_SAVE != 0
    }

    /// Returns the exact undefined bits retained from the producer.
    pub const fn reserved(self) -> u32 {
        self.0 & !Self::DEFINED_MASK
    }
}

/// A checked, zero-copy view of one null-terminated OfficeArt Unicode name.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Name<'data> {
    raw: &'data [u8],
}

impl<'data> Name<'data> {
    pub(crate) const fn from_raw(raw: &'data [u8]) -> Self {
        Self { raw }
    }

    /// Returns the exact UTF-16LE bytes, including the terminating NUL.
    pub const fn raw_bytes(self) -> &'data [u8] {
        self.raw
    }

    /// Returns the number of UTF-16 code units excluding the terminator.
    pub fn unit_len(self) -> usize {
        self.raw.len() / 2 - 1
    }

    /// Decodes the checked name into an owned Rust string.
    pub fn text(self) -> Result<String> {
        let units = self
            .raw
            .chunks_exact(2)
            .take(self.unit_len())
            .map(|pair| u16::from_le_bytes([pair[0], pair[1]]));
        String::from_utf16(&units.collect::<Vec<_>>()).map_err(|_| Error::MalformedProperties {
            reason: "picture name is not valid UTF-16",
        })
    }
}

/// The typed picture metadata projection for one shape property table.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Metadata<'data> {
    name: Option<Name<'data>>,
    flags: Flags,
}

impl<'data> Metadata<'data> {
    pub(crate) const fn new(name: Option<Name<'data>>, flags: Flags) -> Self {
        Self { name, flags }
    }

    /// Returns the optional picture comment, file name, or URL.
    pub const fn name(self) -> Option<Name<'data>> {
        self.name
    }

    /// Returns the effective BLIP-name flags, including retained reserved bits.
    pub const fn flags(self) -> Flags {
        self.flags
    }
}

/// One reversible before/after value in a picture snapshot patch.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Change<T> {
    before: Option<T>,
    after: Option<T>,
}

impl<T> Change<T> {
    pub(crate) const fn new(before: Option<T>, after: Option<T>) -> Self {
        Self { before, after }
    }

    /// Returns the value observed in the source snapshot.
    pub fn before(&self) -> Option<&T> {
        self.before.as_ref()
    }

    /// Returns the value published by the committed snapshot.
    pub fn after(&self) -> Option<&T> {
        self.after.as_ref()
    }
}
