//! Typed, context-local models for Word range-level protection.

/// The `ProtectionType` carried by a protection-range exception (`PRTI.iProt`,
/// MS-DOC 2.9.219).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Mode {
    /// `iProtNone`: allow all changes.
    None,
    /// `iProtReadWrite`: allow editing of marked editable form regions.
    ReadWrite,
    /// `iProtRevision`: allow annotations and track other changes.
    Revision,
    /// `iProtComment`: allow annotations but no other changes.
    Comment,
    /// `iProtRead`: allow no changes.
    Read,
    /// An unassigned or future wire value retained without interpretation.
    Unknown(u16),
}

impl Mode {
    /// Decode the complete known `ProtectionType` vocabulary without
    /// discarding an unrecognized value.
    pub(crate) const fn from_raw(raw: u16) -> Self {
        match raw {
            0x0000 => Self::None,
            0x0001 => Self::ReadWrite,
            0x0002 => Self::Revision,
            0x0003 => Self::Comment,
            0x0004 => Self::Read,
            value => Self::Unknown(value),
        }
    }

    /// The exact 16-bit value that was stored in `PRTI.iProt`.
    pub const fn raw(self) -> u16 {
        match self {
            Self::None => 0x0000,
            Self::ReadWrite => 0x0001,
            Self::Revision => 0x0002,
            Self::Comment => 0x0003,
            Self::Read => 0x0004,
            Self::Unknown(value) => value,
        }
    }
}

/// The permitted editor selector in a protection-range exception (`UidSel`,
/// MS-DOC 2.9.334).
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Selector {
    /// `uidEveryone`: all users.
    Everyone,
    /// `uidEditors`: editors of the document.
    Editors,
    /// `uidOwners`: owners of the document.
    Owners,
    /// A one-based index into the `SttbProtUser` table.
    User(u16),
    /// A reserved, future, or otherwise invalid-for-`UidSel` wire value.
    Unknown(u16),
}

impl Selector {
    pub(crate) const fn from_raw(raw: u16) -> Self {
        match raw {
            0xFFFB => Self::Editors,
            0xFFFC => Self::Owners,
            0xFFFF => Self::Everyone,
            1..=0x7FFF => Self::User(raw),
            value => Self::Unknown(value),
        }
    }

    /// The exact 16-bit selector value from `PRTI.uidSel`.
    pub const fn raw(self) -> u16 {
        match self {
            Self::Everyone => 0xFFFF,
            Self::Editors => 0xFFFB,
            Self::Owners => 0xFFFC,
            Self::User(index) | Self::Unknown(index) => index,
        }
    }
}

/// The role recorded for one username in `SttbProtUser`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Role {
    /// No role is specified for the username (`0x0000`).
    Unspecified,
    /// The username specifies an owner (`0xFFFC`).
    Owner,
    /// The username specifies an editor (`0xFFFB`).
    Editor,
    /// A reserved or future role value retained verbatim.
    Unknown(u16),
}

impl Role {
    pub(crate) const fn from_raw(raw: u16) -> Self {
        match raw {
            0x0000 => Self::Unspecified,
            0xFFFC => Self::Owner,
            0xFFFB => Self::Editor,
            value => Self::Unknown(value),
        }
    }

    /// The exact 16-bit role value from the username table.
    pub const fn raw(self) -> u16 {
        match self {
            Self::Unspecified => 0x0000,
            Self::Owner => 0xFFFC,
            Self::Editor => 0xFFFB,
            Self::Unknown(value) => value,
        }
    }
}

/// One username from `SttbProtUser`.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct User {
    /// A mapped account (`DOMAIN\\NAME`) or e-mail address, stored verbatim.
    pub name: String,
    /// The role associated with this username.
    pub role: Role,
}

/// Wire values that MS-DOC marks as undefined, ignored, or reserved for one
/// protection range.
///
/// These values have no semantic effect, but retaining them makes a future
/// lossless writer possible and prevents a read from silently normalizing a
/// producer's bytes.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Reserved {
    /// The raw `BKC` flags, including `fPub`, `fNative`, and `fCol`.
    bkc: u16,
    /// Undefined `PRTI.i`.
    prti_i: u16,
    /// Undefined `PRTI.fUseMe`.
    prti_use_me: u16,
    /// Raw non-semantic bytes that appeared in the `SttbfBkmkProt` string
    /// slot. The specification requires this string to be empty; a bounded
    /// copy is retained instead of silently dropping an unexpected value.
    bookmark_data: Box<[u8]>,
}

impl Reserved {
    pub(crate) fn new(bkc: u16, prti_i: u16, prti_use_me: u16, bookmark_data: Box<[u8]>) -> Self {
        Self {
            bkc,
            prti_i,
            prti_use_me,
            bookmark_data,
        }
    }

    /// The exact raw `BKC` word.
    pub const fn bkc(&self) -> u16 {
        self.bkc
    }

    /// The exact undefined `PRTI.i` word.
    pub const fn prti_i(&self) -> u16 {
        self.prti_i
    }

    /// The exact undefined `PRTI.fUseMe` word.
    pub const fn prti_use_me(&self) -> u16 {
        self.prti_use_me
    }

    /// Unexpected bytes from the otherwise-required empty bookmark string.
    pub fn bookmark_data(&self) -> &[u8] {
        &self.bookmark_data
    }
}

/// One range-level protection bookmark and its editing exception.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Range {
    /// Character position where the protected range begins.
    pub start: u32,
    /// Character position of the first character beyond the range.
    pub end: u32,
    /// Whether the bookmark is expected to survive RTF/HTML/XML export.
    pub is_native: bool,
    /// The single table column spanned by a column bookmark, if any.
    pub column: Option<u8>,
    /// The users permitted to edit this range.
    pub editor: Selector,
    /// The protection mode recorded by the exception. MS-DOC requires
    /// `ReadWrite` for ordinary range-level protection, but unknown values
    /// remain observable rather than being discarded.
    pub mode: Mode,
    reserved: Reserved,
}

impl Range {
    pub(crate) fn from_parts(
        start: u32,
        end: u32,
        is_native: bool,
        column: Option<u8>,
        editor: Selector,
        mode: Mode,
        reserved: Reserved,
    ) -> Self {
        Self {
            start,
            end,
            is_native,
            column,
            editor,
            mode,
            reserved,
        }
    }

    /// The ignored or reserved wire values associated with this range.
    pub fn reserved(&self) -> &Reserved {
        &self.reserved
    }
}

/// The usernames and range-level protection bookmarks from one Word document.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Ranges {
    users: Vec<User>,
    ranges: Vec<Range>,
}

impl Ranges {
    pub(crate) fn from_parts(users: Vec<User>, ranges: Vec<Range>) -> Self {
        Self { users, ranges }
    }

    /// Usernames from `SttbProtUser`, in table order.
    pub fn users(&self) -> &[User] {
        &self.users
    }

    /// Editable ranges in start-CP order.
    pub fn ranges(&self) -> &[Range] {
        &self.ranges
    }

    /// Resolve a one-based `Selector::User` index.
    pub fn user(&self, index: u16) -> Option<&User> {
        usize::from(index)
            .checked_sub(1)
            .and_then(|zero| self.users.get(zero))
    }

    /// Resolve the username selected by a range, when the selector is an
    /// indexed user rather than a well-known group.
    pub fn editor_for(&self, range: &Range) -> Option<&User> {
        match range.editor {
            Selector::User(index) => self.user(index),
            Selector::Everyone | Selector::Editors | Selector::Owners | Selector::Unknown(_) => {
                None
            },
        }
    }
}
