//! Semantic values for the inert XLSB shared-workbook owner.
//!
//! The shared-workbook log is deliberately not interpreted here.  Its records
//! stay in source order as [`RawRecord`] values; callers can inspect the
//! common revision envelope through [`RecordView`] without applying or
//! replaying a revision.

use std::fmt;

/// The maximum number of users permitted by `BrtCUsr`.
pub const MAX_USERS: usize = 256;
/// A conservative bound for revision headers in one owner.
pub const MAX_HEADERS: usize = 65_535;
/// A conservative bound for opaque records in one revision log.
pub const MAX_RECORDS: usize = 1_000_000;
/// Maximum BIFF12 part bytes accepted by this owner.
pub const MAX_PART_BYTES: usize = 64 * 1024 * 1024;
/// Maximum UTF-16 code units in a metadata string.
pub const MAX_STRING_UNITS: usize = 1_048_576;
/// Maximum UTF-16 code units for names constrained by `[MS-XLSB]`.
pub const MAX_NAME_UNITS: usize = 54;

/// A GUID stored in a BIFF12 shared-workbook record.
#[derive(Clone, Copy, PartialEq, Eq, Hash)]
pub struct Guid([u8; 16]);

impl Guid {
    /// Construct a GUID from its on-wire 16-byte representation.
    #[must_use]
    pub const fn from_bytes(bytes: [u8; 16]) -> Self {
        Self(bytes)
    }

    /// Borrow the on-wire representation without allocation.
    #[must_use]
    pub const fn as_bytes(self) -> [u8; 16] {
        self.0
    }
}

impl fmt::Debug for Guid {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("Guid(")?;
        for (index, byte) in self.0.iter().enumerate() {
            if index != 0 {
                formatter.write_str(":")?;
            }
            write!(formatter, "{byte:02X}")?;
        }
        formatter.write_str(")")
    }
}

impl From<[u8; 16]> for Guid {
    fn from(value: [u8; 16]) -> Self {
        Self::from_bytes(value)
    }
}

/// A validated `ShortDtr` value from `[MS-XLSB]` section 2.5.133.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct ShortDateTime {
    /// Gregorian year, from 1900 through 9999.
    pub year: u16,
    /// Month, from 1 through 12.
    pub month: u8,
    /// Day of month, from 1 through 31.
    pub day: u8,
    /// Hour, from 0 through 23.
    pub hour: u8,
    /// Minute, from 0 through 59.
    pub minute: u8,
    /// Second, from 0 through 59.
    pub second: u8,
    /// ISO weekday, from 1 (Monday) through 7 (Sunday).
    pub weekday: u8,
}

impl Default for ShortDateTime {
    fn default() -> Self {
        Self {
            year: 1900,
            month: 1,
            day: 1,
            hour: 0,
            minute: 0,
            second: 0,
            weekday: 1,
        }
    }
}

/// Typed `BrtInfo` revision metadata.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Info {
    /// GUID for the latest revision header set.
    pub guid: Guid,
    /// GUID for the latest revision header set saved to disk.
    pub root_guid: Guid,
    /// Current shared-workbook revision number.
    pub revision_id: u32,
    /// Current shared-workbook version.
    pub version: i32,
    /// Whether the workbook contains revisions.
    pub has_revisions: bool,
    /// Whether revision history is not retained.
    pub no_revision_history: bool,
    /// Whether change tracking is protected.
    pub protected: bool,
    /// Number of days to retain history.  Zero is valid only when history is
    /// disabled, as required by `[MS-XLSB]`.
    pub revision_history_interval: u16,
}

impl Default for Info {
    fn default() -> Self {
        Self {
            guid: Guid::from_bytes([0; 16]),
            root_guid: Guid::from_bytes([0; 16]),
            revision_id: 0,
            version: 1,
            has_revisions: false,
            no_revision_history: true,
            protected: false,
            revision_history_interval: 0,
        }
    }
}

/// Typed `BrtUsr` metadata for one currently open user.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct User {
    /// Unique user identifier.
    pub id: u32,
    /// Revision-header GUID to which the user is synchronized.
    pub guid: Guid,
    /// Local date and time at which the workbook was opened.
    pub opened_at: ShortDateTime,
    /// Display name stored in the user log.
    pub name: String,
}

/// Typed `BrtRRHeader` metadata for one revision log.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Header {
    /// GUID identifying this revision-header/log pair.
    pub guid: Guid,
    /// Local date and time at which this set was saved.
    pub saved_at: ShortDateTime,
    /// Next unused sheet identifier, or `0xFFFF` when unknown.
    pub next_sheet_id: u16,
    /// Lowest reviewable revision identifier, or zero when absent.
    pub revision_min: u32,
    /// Highest reviewable revision identifier, or zero when absent.
    pub revision_max: u32,
    /// User responsible for this revision-header set.
    pub user_name: String,
    /// Relationship identifier for the corresponding revision-log part.
    pub relationship_id: String,
    /// Sheet identifiers in display order when this set was saved.
    pub sheet_ids: Vec<u16>,
    /// Revision identifiers already reviewed.
    pub reviewed: Vec<u32>,
}

/// The common 14-byte `RRd` envelope shown by an opaque revision-log record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct RevisionEnvelope {
    /// Reviewable revision identifier, or zero for a non-reviewable record.
    pub revision_id: u32,
    /// `[MS-XLSB]` `RevisionType` numeric value.
    pub revision_type: u16,
    /// Whether the revision was accepted.
    pub accepted: bool,
    /// Whether the revision was produced by an undo action.
    pub undo_action: bool,
    /// Reserved envelope bit retained for inert inspection.
    pub reserved_one: bool,
    /// Reserved envelope bit retained for inert inspection.
    pub reserved_two: bool,
    /// Associated sheet identifier, or `0xFFFF` for workbook-wide records.
    pub sheet_id: u16,
}

/// One raw BIFF12 record whose framing and payload are preserved.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RawRecord {
    /// Numeric BIFF12 record kind.
    pub kind: u16,
    /// Exact record payload bytes.
    pub payload: Vec<u8>,
}

impl RawRecord {
    /// Construct a raw record.  Range and payload checks happen when it is
    /// committed through a snapshot transaction.
    #[must_use]
    pub const fn new(kind: u16, payload: Vec<u8>) -> Self {
        Self { kind, payload }
    }
}

/// Read-only view over one opaque revision-log record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RecordView<'a> {
    /// Record kind.
    pub kind: u16,
    /// Full payload, borrowed from the owning log.
    pub payload: &'a [u8],
    /// Common envelope, when this record kind has a safely recognized `RRd`
    /// prefix and the payload is well formed.
    pub envelope: Option<RevisionEnvelope>,
}

/// The user-names part and its package identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct UserNames {
    /// Workbook relationship identifier for this part.
    pub relationship_id: String,
    /// Relationship type as authored in the OPC graph.
    pub relationship_type: String,
    /// Absolute OPC part name.
    pub part_name: String,
    /// Typed users in source order.
    pub users: Vec<User>,
    pub(crate) records: Vec<RawRecord>,
    pub(crate) count_slot: Option<usize>,
    pub(crate) user_slots: Vec<usize>,
}

impl UserNames {
    /// Construct a detached user-names value for a new package owner.
    #[must_use]
    pub fn new(users: Vec<User>) -> Self {
        Self {
            relationship_id: String::new(),
            relationship_type: String::new(),
            part_name: String::new(),
            users,
            records: Vec::new(),
            count_slot: None,
            user_slots: Vec::new(),
        }
    }

    /// Borrow source-order raw records, including unknown records.
    #[must_use]
    pub fn raw_records(&self) -> &[RawRecord] {
        &self.records
    }
}

/// The revision-headers part and its package identity.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RevisionHeaders {
    /// Workbook relationship identifier for this part.
    pub relationship_id: String,
    /// Relationship type as authored in the OPC graph.
    pub relationship_type: String,
    /// Absolute OPC part name.
    pub part_name: String,
    /// Typed workbook revision properties.
    pub info: Info,
    /// Typed headers in source order.
    pub headers: Vec<Header>,
    pub(crate) records: Vec<RawRecord>,
    pub(crate) info_slot: Option<usize>,
    pub(crate) header_slots: Vec<usize>,
}

impl RevisionHeaders {
    /// Construct detached revision headers for a new package owner.
    #[must_use]
    pub fn new(info: Info, headers: Vec<Header>) -> Self {
        Self {
            relationship_id: String::new(),
            relationship_type: String::new(),
            part_name: String::new(),
            info,
            headers,
            records: Vec::new(),
            info_slot: None,
            header_slots: Vec::new(),
        }
    }

    /// Borrow source-order raw records, including unknown records.
    #[must_use]
    pub fn raw_records(&self) -> &[RawRecord] {
        &self.records
    }
}

/// One revision-log part.  Records are intentionally opaque and ordered.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct RevisionLog {
    /// Relationship identifier from the revision-headers part.
    pub relationship_id: String,
    /// Relationship type as authored in the OPC graph.
    pub relationship_type: String,
    /// Absolute OPC part name.
    pub part_name: String,
    /// Opaque records in exact source order.
    pub records: Vec<RawRecord>,
}

impl RevisionLog {
    /// Construct a detached empty revision log.
    #[must_use]
    pub fn new(records: Vec<RawRecord>) -> Self {
        Self {
            relationship_id: String::new(),
            relationship_type: String::new(),
            part_name: String::new(),
            records,
        }
    }

    /// Iterate inert record views without interpreting or replaying changes.
    pub fn views(&self) -> impl Iterator<Item = RecordView<'_>> {
        self.records.iter().map(|record| RecordView {
            kind: record.kind,
            payload: &record.payload,
            envelope: super::codec::revision_envelope(record.kind, &record.payload),
        })
    }
}

/// The package-neutral shared-workbook metadata graph.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct Catalog {
    /// Optional user-names part.  Its presence identifies a shared workbook.
    pub users: Option<UserNames>,
    /// Optional revision-headers part.
    pub headers: Option<RevisionHeaders>,
    /// Revision-log parts in the same order as their headers.
    pub logs: Vec<RevisionLog>,
}

impl Catalog {
    /// Construct an empty, non-shared workbook metadata graph.
    #[must_use]
    pub const fn empty() -> Self {
        Self {
            users: None,
            headers: None,
            logs: Vec::new(),
        }
    }
}
