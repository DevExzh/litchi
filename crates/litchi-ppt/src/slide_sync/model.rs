//! Typed values and immutable snapshots for slide-library synchronization.

use crate::package::{Error, Result};
use crate::records::Record;
use chrono::NaiveDate;

use super::{codec, transaction, validation};

/// Maximum UTF-16 payload accepted for a synchronization string.
pub const MAX_TEXT_BYTES: usize = 16 * 1024;

/// Limits applied while capturing or editing one slide record.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct Limits {
    /// Maximum record nesting depth.
    pub max_depth: usize,
    /// Maximum number of records in the captured slide.
    pub max_records: usize,
    /// Maximum encoded size of the captured slide.
    pub max_bytes: usize,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_depth: 128,
            max_records: 262_144,
            max_bytes: 64 * 1024 * 1024,
        }
    }
}

impl Limits {
    pub(crate) fn validate(self) -> Result<Self> {
        if self.max_depth == 0 || self.max_records == 0 || self.max_bytes < 8 {
            return Err(Error::InvalidFormat(
                "slide synchronization limits must allow one record".into(),
            ));
        }
        Ok(self)
    }
}

/// A validated slide-library identifier.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct ServerId(String);

impl ServerId {
    /// Build an identifier from its semantic Unicode value.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let text = value.into();
        validate_printable(&text, "server slide identifier")?;
        Ok(Self(text))
    }

    /// Borrow the identifier text.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the value and return its text.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0
    }

    pub(crate) fn from_wire(bytes: &[u8]) -> Result<Self> {
        Self::new(decode_printable(bytes, "ServerIdAtom")?)
    }

    pub(crate) fn wire(&self) -> Result<Vec<u8>> {
        encode_printable(&self.0, "server slide identifier")
    }
}

impl AsRef<str> for ServerId {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for ServerId {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<&str> for ServerId {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// A validated HTTP slide-library URI.
#[derive(Debug, Clone, PartialEq, Eq, Hash)]
pub struct LibraryUrl(String);

impl LibraryUrl {
    /// Build a URL using the HTTP URI required by `[MS-PPT]`.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(value: impl Into<String>) -> Result<Self> {
        let text = value.into();
        validate_printable(&text, "slide-library URL")?;
        let parsed = url::Url::parse(&text)
            .map_err(|_err| Error::InvalidFormat("slide-library URL is not a valid URI".into()))?;
        if parsed.scheme() != "http" {
            return Err(Error::InvalidFormat(
                "slide-library URL must use the HTTP scheme".into(),
            ));
        }
        Ok(Self(text))
    }

    /// Borrow the exact URL text supplied by the producer.
    #[must_use]
    pub fn as_str(&self) -> &str {
        &self.0
    }

    /// Consume the value and return its exact text.
    #[must_use]
    pub fn into_string(self) -> String {
        self.0
    }

    pub(crate) fn from_wire(bytes: &[u8]) -> Result<Self> {
        Self::new(decode_printable(bytes, "SlideLibUrlAtom")?)
    }

    pub(crate) fn wire(&self) -> Result<Vec<u8>> {
        encode_printable(&self.0, "slide-library URL")
    }
}

impl AsRef<str> for LibraryUrl {
    fn as_ref(&self) -> &str {
        self.as_str()
    }
}

impl TryFrom<String> for LibraryUrl {
    type Error = Error;

    fn try_from(value: String) -> Result<Self> {
        Self::new(value)
    }
}

impl TryFrom<&str> for LibraryUrl {
    type Error = Error;

    fn try_from(value: &str) -> Result<Self> {
        Self::new(value)
    }
}

/// A validated Windows `SYSTEMTIME` carried by `SlideSyncInfoAtom12`.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct SystemTime {
    pub(crate) year: u16,
    pub(crate) month: u16,
    pub(crate) day_of_week: u16,
    pub(crate) day: u16,
    pub(crate) hour: u16,
    pub(crate) minute: u16,
    pub(crate) second: u16,
    pub(crate) millisecond: u16,
}

impl SystemTime {
    /// Construct and validate one Gregorian `SYSTEMTIME` value.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[allow(
        clippy::too_many_arguments,
        reason = "constructor parameters mirror the atom wire fields one-to-one"
    )]
    pub fn new(
        year: u16,
        month: u16,
        day_of_week: u16,
        day: u16,
        hour: u16,
        minute: u16,
        second: u16,
        millisecond: u16,
    ) -> Result<Self> {
        let value = Self {
            year,
            month,
            day_of_week,
            day,
            hour,
            minute,
            second,
            millisecond,
        };
        value.validate()?;
        Ok(value)
    }

    #[must_use]
    pub const fn year(self) -> u16 {
        self.year
    }

    #[must_use]
    pub const fn month(self) -> u16 {
        self.month
    }

    #[must_use]
    pub const fn day_of_week(self) -> u16 {
        self.day_of_week
    }

    #[must_use]
    pub const fn day(self) -> u16 {
        self.day
    }

    #[must_use]
    pub const fn hour(self) -> u16 {
        self.hour
    }

    #[must_use]
    pub const fn minute(self) -> u16 {
        self.minute
    }

    #[must_use]
    pub const fn second(self) -> u16 {
        self.second
    }

    #[must_use]
    pub const fn millisecond(self) -> u16 {
        self.millisecond
    }

    pub(crate) fn from_wire(data: &[u8], field: &str) -> Result<Self> {
        if data.len() != 16 {
            return Err(Error::Corrupted(format!(
                "{field} must contain exactly 16 bytes"
            )));
        }
        let field_at = |index| u16::from_le_bytes([data[index], data[index + 1]]);
        Self::new(
            field_at(0),
            field_at(2),
            field_at(4),
            field_at(6),
            field_at(8),
            field_at(10),
            field_at(12),
            field_at(14),
        )
        .map_err(|_err| Error::Corrupted(format!("{field} is not a valid SYSTEMTIME")))
    }

    pub(crate) fn write_wire(self, output: &mut Vec<u8>) {
        for value in [
            self.year,
            self.month,
            self.day_of_week,
            self.day,
            self.hour,
            self.minute,
            self.second,
            self.millisecond,
        ] {
            output.extend_from_slice(&value.to_le_bytes());
        }
    }

    fn validate(self) -> Result<()> {
        if !(1601..=30_827).contains(&self.year)
            || self.day_of_week > 6
            || self.hour > 23
            || self.minute > 59
            || self.second > 59
            || self.millisecond > 999
            || NaiveDate::from_ymd_opt(
                i32::from(self.year),
                u32::from(self.month),
                u32::from(self.day),
            )
            .is_none()
        {
            return Err(Error::InvalidFormat(
                "SYSTEMTIME contains an out-of-range field".into(),
            ));
        }
        Ok(())
    }
}

/// Typed semantic metadata connecting a slide to a slide library.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Synchronization {
    server_slide_id: ServerId,
    slide_library_url: LibraryUrl,
    server_modified: SystemTime,
    client_inserted: SystemTime,
}

impl Synchronization {
    /// Construct a complete synchronization value from ergonomic strings and
    /// validated timestamps.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn new(
        server_slide_id: impl Into<String>,
        slide_library_url: impl Into<String>,
        server_modified: SystemTime,
        client_inserted: SystemTime,
    ) -> Result<Self> {
        Self::from_parts(
            ServerId::new(server_slide_id)?,
            LibraryUrl::new(slide_library_url)?,
            server_modified,
            client_inserted,
        )
    }

    /// Construct from already validated wire-domain values.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_parts(
        server_slide_id: ServerId,
        slide_library_url: LibraryUrl,
        server_modified: SystemTime,
        client_inserted: SystemTime,
    ) -> Result<Self> {
        Ok(Self {
            server_slide_id,
            slide_library_url,
            server_modified,
            client_inserted,
        })
    }

    /// Borrow the server-side slide identifier.
    #[must_use]
    pub const fn server_slide_id(&self) -> &ServerId {
        &self.server_slide_id
    }

    /// Borrow the exact slide-library URL.
    #[must_use]
    pub const fn slide_library_url(&self) -> &LibraryUrl {
        &self.slide_library_url
    }

    /// Return the server modification timestamp.
    #[must_use]
    pub const fn server_modified(&self) -> SystemTime {
        self.server_modified
    }

    /// Return the client insertion timestamp.
    #[must_use]
    pub const fn client_inserted(&self) -> SystemTime {
        self.client_inserted
    }

    /// Parse the optional synchronization container below a slide record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(root: &Record) -> Result<Option<Self>> {
        codec::read(root)
    }
}

/// An immutable, lossless snapshot of one legacy `PowerPoint` slide record.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Snapshot {
    pub(crate) root: Record,
    pub(crate) bytes: Vec<u8>,
    pub(crate) limits: Limits,
}

impl Snapshot {
    /// Capture a slide record and normalize its container payloads.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn from_record(root: Record) -> Result<Self> {
        Self::from_record_with_limits(root, Limits::default())
    }

    /// Capture a slide record under explicit resource limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    #[allow(
        clippy::needless_pass_by_value,
        reason = "taking the record by value is the established public API; the snapshot stores a normalized re-parse, so the argument is only borrowed internally"
    )]
    pub fn from_record_with_limits(root: Record, limits: Limits) -> Result<Self> {
        let validated = limits.validate()?;
        validation::validate_slide_header(&root)?;
        let bytes = codec::encode(&root, validated)?;
        let normalized_root = codec::parse(&bytes, validated)?;
        validation::validate(&normalized_root, validated)?;
        Ok(Self {
            root: normalized_root,
            bytes,
            limits: validated,
        })
    }

    /// Parse and capture exactly one complete slide record.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse(bytes: &[u8]) -> Result<Self> {
        Self::parse_with_limits(bytes, Limits::default())
    }

    /// Parse one complete slide record under explicit limits.
    ///
    /// # Errors
    ///
    /// Returns an error if the input cannot be read or is malformed.
    pub fn parse_with_limits(bytes: &[u8], limits: Limits) -> Result<Self> {
        let validated = limits.validate()?;
        let root = codec::parse(bytes, validated)?;
        validation::validate(&root, validated)?;
        let encoded = codec::encode(&root, validated)?;
        if encoded != bytes {
            return Err(Error::Corrupted(
                "slide synchronization snapshot is not losslessly representable".into(),
            ));
        }
        Ok(Self {
            root,
            bytes: bytes.to_vec(),
            limits: validated,
        })
    }

    /// Borrow the complete slide record represented by this snapshot.
    #[must_use]
    pub const fn record(&self) -> &Record {
        &self.root
    }

    /// Borrow the exact source or committed bytes.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    /// Return the optional typed synchronization metadata.
    ///
    /// # Errors
    ///
    /// Returns an error if the operation fails.
    pub fn synchronization(&self) -> Result<Option<Synchronization>> {
        codec::read(&self.root)
    }

    /// Resource limits attached to this snapshot.
    #[must_use]
    pub const fn limits(&self) -> Limits {
        self.limits
    }

    /// Return a deterministic content revision for optimistic integration.
    #[must_use]
    pub fn revision(&self) -> transaction::Revision {
        transaction::Revision::from_bytes(&self.bytes)
    }

    /// Start an isolated atomic semantic edit.
    #[must_use]
    pub fn edit(&self) -> transaction::Editor {
        transaction::Editor::open(self.clone())
    }
}

fn validate_printable(value: &str, field: &str) -> Result<()> {
    let units = value.encode_utf16().collect::<Vec<_>>();
    if units.is_empty() {
        return Err(Error::InvalidFormat(format!("{field} must not be empty")));
    }
    if units
        .len()
        .checked_mul(2)
        .is_none_or(|bytes| bytes > MAX_TEXT_BYTES)
    {
        return Err(Error::InvalidFormat(format!(
            "{field} exceeds {MAX_TEXT_BYTES} UTF-16 bytes"
        )));
    }
    if units
        .iter()
        .any(|unit| *unit == 0 || matches!(*unit, 0x0001..=0x001f | 0x007f..=0x009f))
    {
        return Err(Error::InvalidFormat(format!(
            "{field} contains a non-printable character"
        )));
    }
    Ok(())
}

fn decode_printable(bytes: &[u8], field: &str) -> Result<String> {
    if !bytes.len().is_multiple_of(2) {
        return Err(Error::Corrupted(format!("{field} payload must be even")));
    }
    if bytes.len() > MAX_TEXT_BYTES {
        return Err(Error::InvalidFormat(format!(
            "{field} exceeds {MAX_TEXT_BYTES} UTF-16 bytes"
        )));
    }
    let mut units = Vec::with_capacity(bytes.len() / 2);
    for pair in bytes.chunks_exact(2) {
        let unit = u16::from_le_bytes([pair[0], pair[1]]);
        if unit == 0 {
            break;
        }
        if matches!(unit, 0x0001..=0x001f | 0x007f..=0x009f) {
            return Err(Error::Corrupted(format!(
                "{field} contains a non-printable character"
            )));
        }
        units.push(unit);
    }
    String::from_utf16(&units)
        .map_err(|_err| Error::Corrupted(format!("{field} contains invalid UTF-16")))
}

fn encode_printable(value: &str, field: &str) -> Result<Vec<u8>> {
    validate_printable(value, field)?;
    let units = value.encode_utf16().collect::<Vec<_>>();
    let byte_len = units
        .len()
        .checked_mul(2)
        .ok_or_else(|| Error::InvalidFormat(format!("{field} payload size overflow")))?;
    let mut bytes = Vec::new();
    bytes
        .try_reserve_exact(byte_len)
        .map_err(|_err| Error::InvalidFormat(format!("{field} allocation failed")))?;
    for unit in units {
        bytes.extend_from_slice(&unit.to_le_bytes());
    }
    Ok(bytes)
}
