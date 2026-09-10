//! Semantic values owned by the BIFF8 `User Names` stream.

use crate::revision_records::ShortDtr;
use crate::{Error, Result};

/// A GUID-sized revision-log identity used by `UsrInfo.guid`.
pub type UserGuid = [u8; 16];

/// The `UsrChk` version record (MS-XLS 2.4.338).
///
/// The version and reserved words are retained as raw values.  This keeps
/// forward-compatible producer values observable while the stream owner
/// remains inert.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct UserCheck {
    version: u16,
    reserved: u16,
}

impl UserCheck {
    /// Construct a check record while retaining its reserved word.
    #[must_use]
    pub const fn new(version: u16, reserved: u16) -> Self {
        Self { version, reserved }
    }

    /// BIFF version reported by the last user.
    #[must_use]
    pub const fn version(self) -> u16 {
        self.version
    }

    /// Reserved bits retained from the source record.
    #[must_use]
    pub const fn reserved(self) -> u16 {
        self.reserved
    }
}

/// One `UsrInfo` entry in the shared-workbook user log.
///
/// The revision GUID is intentionally immutable through the metadata edit
/// API.  It is a dependency edge into the `Revision Log` stream and changing
/// it without updating that stream would create a dangling user-log entry.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserEntry {
    pub(crate) user_id: i32,
    pub(crate) guid: UserGuid,
    pub(crate) opened_at: ShortDtr,
    pub(crate) user_name: String,
    pub(crate) string_flags: u8,
    pub(crate) unused: u8,
}

impl UserEntry {
    /// Construct a new user-log entry.
    ///
    /// `user_name` is checked against the `UsrInfo.stUserName` character
    /// bounds.  The timestamp is already a validated [`ShortDtr`].
    /// # Errors
    ///
    /// Returns an error when the name is empty or too long for the BIFF user
    /// name field. NUL characters are retained as inert string data.
    pub fn new(
        user_id: i32,
        guid: UserGuid,
        opened_at: ShortDtr,
        user_name: impl Into<String>,
    ) -> Result<Self> {
        let user_name = user_name.into();
        validate_user_name(&user_name)?;
        Ok(Self {
            user_id,
            guid,
            opened_at,
            user_name,
            string_flags: 0,
            unused: 0,
        })
    }

    /// Unique signed user identifier (`lUsrId`).
    #[must_use]
    pub const fn user_id(&self) -> i32 {
        self.user_id
    }

    /// Revision-log GUID to which this user is synchronized.
    #[must_use]
    pub const fn guid(&self) -> &UserGuid {
        &self.guid
    }

    /// Time at which this user opened the shared workbook.
    #[must_use]
    pub const fn opened_at(&self) -> ShortDtr {
        self.opened_at
    }

    /// User name stored in `stUserName`.
    #[must_use]
    pub fn user_name(&self) -> &str {
        &self.user_name
    }

    /// Raw `XLUnicodeString` option flags, including unknown producer bits.
    #[must_use]
    pub const fn string_flags(&self) -> u8 {
        self.string_flags
    }

    /// Undefined trailing byte retained from `UsrInfo`.
    #[must_use]
    pub const fn unused(&self) -> u8 {
        self.unused
    }

    /// Return this entry with a new user name.
    /// # Errors
    ///
    /// Returns an error when the name is outside the `UsrInfo` bounds.
    pub fn with_user_name(&self, user_name: impl Into<String>) -> Result<Self> {
        let user_name = user_name.into();
        validate_user_name(&user_name)?;
        Ok(Self {
            user_name,
            ..self.clone()
        })
    }

    /// Return this entry with a new opening timestamp.
    #[must_use]
    pub fn with_opened_at(&self, opened_at: ShortDtr) -> Self {
        Self {
            opened_at,
            user_id: self.user_id,
            guid: self.guid,
            user_name: self.user_name.clone(),
            string_flags: self.string_flags,
            unused: self.unused,
        }
    }

    /// Return this entry with a preserved undefined byte.
    #[must_use]
    pub fn with_unused(&self, unused: u8) -> Self {
        Self {
            unused,
            user_id: self.user_id,
            guid: self.guid,
            opened_at: self.opened_at,
            user_name: self.user_name.clone(),
            string_flags: self.string_flags,
        }
    }
}

/// The complete typed semantic projection of a `User Names` stream.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserNames {
    pub(crate) user_check: UserCheck,
    pub(crate) briefcase_user_count: u16,
    pub(crate) user_record_sizes: [u16; 256],
    pub(crate) users: Vec<UserEntry>,
}

impl UserNames {
    /// Number of users represented by the `CUsr.iCount` field.
    #[must_use]
    pub fn user_count(&self) -> usize {
        self.users.len()
    }

    /// Raw `CUsr.iCount` value represented by this model.
    #[must_use]
    pub fn cusr_count(&self) -> u16 {
        self.users.len() as u16
    }

    /// User entries in source order.
    #[must_use]
    pub fn users(&self) -> &[UserEntry] {
        &self.users
    }

    /// The `UsrChk` record.
    #[must_use]
    pub const fn user_check(&self) -> UserCheck {
        self.user_check
    }

    /// The `BCUsrs.iCount` value.  It is kept separately from `CUsr.iCount`.
    #[must_use]
    pub const fn briefcase_user_count(&self) -> u16 {
        self.briefcase_user_count
    }

    /// Alias for [`Self::briefcase_user_count`].
    #[must_use]
    pub const fn bcusrs_count(&self) -> u16 {
        self.briefcase_user_count()
    }

    /// All 256 `CbUsr.rgCbUsr` entries, including reserved slots.
    #[must_use]
    pub const fn user_record_sizes(&self) -> &[u16; 256] {
        &self.user_record_sizes
    }

    /// Alias for [`Self::user_record_sizes`].
    #[must_use]
    pub const fn cbusr_sizes(&self) -> &[u16; 256] {
        self.user_record_sizes()
    }
}

/// Resource bounds for a `User Names` stream.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum complete stream bytes retained by a snapshot or candidate.
    /// The same ceiling is applied to the required `Revision Log` dependency
    /// before that stream is copied for bounded GUID-closure scanning.
    pub max_stream_bytes: usize,
    /// Maximum users accepted from `CUsr.iCount`.
    pub max_users: usize,
    /// Maximum framed records admitted while scanning the required Revision
    /// Log for User Names GUID closure.
    pub max_revision_records: usize,
    /// Maximum `RRDHead` GUIDs retained for User Names closure validation.
    pub max_revision_guids: usize,
}

impl Limits {
    /// Conservative default bounds for the small fixed-layout stream.
    pub const DEFAULT: Self = Self {
        max_stream_bytes: 64 * 1024 * 1024,
        max_users: 255,
        max_revision_records: 1_000_000,
        max_revision_guids: 1_000_000,
    };

    /// Hard stream ceiling accepted by this owner. It is intentionally finite
    /// even when a caller supplies a larger custom profile.
    pub const MAX_STREAM_BYTES: usize = 64 * 1024 * 1024;

    /// Hard cap on framed Revision Log records visited by the User Names
    /// closure scanner, even when a caller supplies a larger profile.
    pub const MAX_REVISION_RECORDS: usize = 1_000_000;

    /// Hard cap on `RRDHead` GUIDs retained by the User Names owner.
    pub const MAX_REVISION_GUIDS: usize = 1_000_000;

    /// Set the maximum complete stream size.
    #[must_use]
    pub const fn with_max_stream_bytes(mut self, value: usize) -> Self {
        self.max_stream_bytes = value;
        self
    }

    /// Set the maximum user count, capped by the BIFF `CUsr` field.
    #[must_use]
    pub const fn with_max_users(mut self, value: usize) -> Self {
        self.max_users = value;
        self
    }

    /// Set the maximum framed Revision Log records visited for closure.
    #[must_use]
    pub const fn with_max_revision_records(mut self, value: usize) -> Self {
        self.max_revision_records = value;
        self
    }

    /// Set the maximum Revision Log header GUIDs retained for closure.
    #[must_use]
    pub const fn with_max_revision_guids(mut self, value: usize) -> Self {
        self.max_revision_guids = value;
        self
    }

    pub(crate) fn validate(self) -> Result<Self> {
        if self.max_stream_bytes == 0 || self.max_stream_bytes > Self::MAX_STREAM_BYTES {
            return Err(Error::InvalidData(format!(
                "User Names stream byte limit must be in 1..={} bytes",
                Self::MAX_STREAM_BYTES
            )));
        }
        if self.max_users > 255 {
            return Err(Error::InvalidData(
                "User Names stream user limit exceeds CUsr.iCount".to_string(),
            ));
        }
        if self.max_revision_records < 4 || self.max_revision_records > Self::MAX_REVISION_RECORDS {
            return Err(Error::InvalidData(format!(
                "Revision Log record limit must be in 4..={} records",
                Self::MAX_REVISION_RECORDS
            )));
        }
        if self.max_revision_guids > Self::MAX_REVISION_GUIDS {
            return Err(Error::InvalidData(format!(
                "Revision Log GUID limit must be at most {} GUIDs",
                Self::MAX_REVISION_GUIDS
            )));
        }
        Ok(self)
    }
}

impl Default for Limits {
    fn default() -> Self {
        Self::DEFAULT
    }
}

pub(crate) fn validate_user_name(user_name: &str) -> Result<()> {
    let utf16_len = user_name.encode_utf16().count();
    if !(1..=54).contains(&utf16_len) {
        return Err(Error::UnsafeEdit(format!(
            "UsrInfo user name has {utf16_len} UTF-16 code units; expected 1..=54"
        )));
    }
    Ok(())
}
