//! RTF document information and properties.

#![allow(
    clippy::arbitrary_source_item_ordering,
    reason = "items stay grouped by RTF feature area rather than by item kind"
)]
use crate::{RtfError, RtfResult};
use std::borrow::Cow;
use std::ops::Range;

pub(crate) const MAX_INFO_TEXT_BYTES: usize = 1_048_576;
pub(crate) const PROTECTION_PASSWORD_HASH_BYTES: usize = 8;

/// Largest accepted `\passwordhash` payload, in decoded bytes.
///
/// The RTF grammar treats this destination as opaque SDATA.  The bounded
/// limit keeps that inert payload safe to retain while allowing the observed
/// record and future records with additional fields.
pub const MAX_PASSWORD_HASH_BYTES: usize = 1_048_576;
pub(crate) const PASSWORD_HASH_HEADER_BYTES: usize = 28;

/// A typed, inert `\passwordhash` record from the modern RTF protection
/// destination.
///
/// The record is never used to authenticate, decrypt, or execute anything.
/// Its complete byte payload is retained so unknown header values and trailing
/// fields survive a parse/write cycle.  The exposed fields cover the observed
/// fixed header layout (little-endian version, total size, flags, algorithm,
/// spin count, logical hash size, and salt size).
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PasswordHash<'a> {
    raw: Cow<'a, [u8]>,
    version: u32,
    total_size: u32,
    flags: u32,
    algorithm_id: u32,
    spin_count: u32,
    hash_size: u32,
    salt_size: u32,
    salt_range: Range<usize>,
    hash_range: Range<usize>,
}

impl<'a> PasswordHash<'a> {
    /// Parse an inert binary password-hash record.
    ///
    /// Unknown values in the fixed header and bytes after the logical hash
    /// are retained verbatim.  No password operation is performed.
    ///
    /// # Errors
    /// Returns an error when the record is truncated, inconsistent, or larger
    /// than [`MAX_PASSWORD_HASH_BYTES`].
    pub fn from_bytes(bytes: Cow<'a, [u8]>) -> RtfResult<Self> {
        if bytes.len() > MAX_PASSWORD_HASH_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF passwordhash payload exceeds the safety limit".to_string(),
            ));
        }
        if bytes.len() < PASSWORD_HASH_HEADER_BYTES {
            return Err(RtfError::MalformedDocument(
                "RTF passwordhash payload is shorter than its header".to_string(),
            ));
        }
        let read = |offset: usize| {
            let end = offset.saturating_add(4);
            bytes
                .get(offset..end)
                .and_then(|value| <[u8; 4]>::try_from(value).ok())
                .map(u32::from_le_bytes)
        };
        let version = read(0).ok_or_else(|| {
            RtfError::MalformedDocument("RTF passwordhash version is truncated".to_string())
        })?;
        let total_size = read(4).ok_or_else(|| {
            RtfError::MalformedDocument("RTF passwordhash size is truncated".to_string())
        })?;
        let expected_size = usize::try_from(total_size).map_err(|_err| {
            RtfError::MalformedDocument("RTF passwordhash size is not representable".to_string())
        })?;
        if expected_size != bytes.len() {
            return Err(RtfError::MalformedDocument(
                "RTF passwordhash size does not match its payload".to_string(),
            ));
        }
        let flags = read(8).ok_or_else(|| {
            RtfError::MalformedDocument("RTF passwordhash flags are truncated".to_string())
        })?;
        let algorithm_id = read(12).ok_or_else(|| {
            RtfError::MalformedDocument("RTF passwordhash algorithm is truncated".to_string())
        })?;
        let spin_count = read(16).ok_or_else(|| {
            RtfError::MalformedDocument("RTF passwordhash spin count is truncated".to_string())
        })?;
        let hash_size = read(20).ok_or_else(|| {
            RtfError::MalformedDocument("RTF passwordhash hash size is truncated".to_string())
        })?;
        let salt_size = read(24).ok_or_else(|| {
            RtfError::MalformedDocument("RTF passwordhash salt size is truncated".to_string())
        })?;
        let salt_size_usize = usize::try_from(salt_size).map_err(|_err| {
            RtfError::MalformedDocument(
                "RTF passwordhash salt size is not representable".to_string(),
            )
        })?;
        let hash_size_usize = usize::try_from(hash_size).map_err(|_err| {
            RtfError::MalformedDocument(
                "RTF passwordhash hash size is not representable".to_string(),
            )
        })?;
        let salt_end = PASSWORD_HASH_HEADER_BYTES
            .checked_add(salt_size_usize)
            .ok_or_else(|| {
                RtfError::MalformedDocument("RTF passwordhash salt size overflows".to_string())
            })?;
        let hash_end = salt_end.checked_add(hash_size_usize).ok_or_else(|| {
            RtfError::MalformedDocument("RTF passwordhash hash size overflows".to_string())
        })?;
        if hash_end > bytes.len() {
            return Err(RtfError::MalformedDocument(
                "RTF passwordhash salt and hash exceed the payload".to_string(),
            ));
        }
        let value = Self {
            raw: bytes,
            version,
            total_size,
            flags,
            algorithm_id,
            spin_count,
            hash_size,
            salt_size,
            salt_range: PASSWORD_HASH_HEADER_BYTES..salt_end,
            hash_range: salt_end..hash_end,
        };
        value.validate()?;
        Ok(value)
    }

    /// Parse a borrowed inert binary password-hash record without copying it.
    pub fn from_slice(bytes: &'a [u8]) -> RtfResult<Self> {
        Self::from_bytes(Cow::Borrowed(bytes))
    }

    /// Construct an authored inert record using the values shown in the RTF
    /// 1.9.1 example layout and no trailing extension bytes.
    ///
    /// This convenience constructor does not claim that the resulting record
    /// is a Word-compatible verifier. Use [`Self::from_parts`] when a source
    /// or producer supplies a different layout or trailing extension.
    pub fn new(salt: Cow<'a, [u8]>, hash: Cow<'a, [u8]>) -> RtfResult<Self> {
        Self::from_parts(1, 1, 0x8004, 50_000, salt, hash, Cow::Borrowed(&[]))
    }

    /// Construct an authored inert record with caller-supplied fixed-header
    /// values and trailing bytes. The fields are retained as data only; this
    /// crate never derives or verifies a password from them.
    pub fn from_parts(
        version: u32,
        flags: u32,
        algorithm_id: u32,
        spin_count: u32,
        salt: Cow<'a, [u8]>,
        hash: Cow<'a, [u8]>,
        trailing: Cow<'a, [u8]>,
    ) -> RtfResult<Self> {
        let total_size = PASSWORD_HASH_HEADER_BYTES
            .checked_add(salt.len())
            .and_then(|size| size.checked_add(hash.len()))
            .and_then(|size| size.checked_add(trailing.len()))
            .ok_or_else(|| {
                RtfError::InvalidStructure("RTF passwordhash size overflow".to_string())
            })?;
        let total_size_u32 = u32::try_from(total_size).map_err(|_err| {
            RtfError::InvalidStructure("RTF passwordhash size exceeds u32".to_string())
        })?;
        if total_size > MAX_PASSWORD_HASH_BYTES {
            return Err(RtfError::InvalidStructure(
                "RTF passwordhash payload exceeds the safety limit".to_string(),
            ));
        }
        let hash_size = u32::try_from(hash.len()).map_err(|_err| {
            RtfError::InvalidStructure("RTF passwordhash hash is too large".to_string())
        })?;
        let salt_size = u32::try_from(salt.len()).map_err(|_err| {
            RtfError::InvalidStructure("RTF passwordhash salt is too large".to_string())
        })?;
        let mut bytes = Vec::new();
        crate::error::try_reserve_additional(
            &mut bytes,
            total_size,
            "RTF passwordhash authored record",
        )?;
        for value in [
            version,
            total_size_u32,
            flags,
            algorithm_id,
            spin_count,
            hash_size,
            salt_size,
        ] {
            bytes.extend_from_slice(&value.to_le_bytes());
        }
        bytes.extend_from_slice(&salt);
        bytes.extend_from_slice(&hash);
        bytes.extend_from_slice(&trailing);
        Self::from_bytes(Cow::Owned(bytes))
    }

    /// Complete encoded payload, including any unknown or trailing fields.
    #[must_use]
    pub fn to_bytes(&self) -> Vec<u8> {
        self.raw.to_vec()
    }

    /// Complete encoded payload borrowed from the retained record.
    #[must_use]
    pub fn bytes(&self) -> &[u8] {
        &self.raw
    }

    /// RTF record version.
    #[must_use]
    pub const fn version(&self) -> u32 {
        self.version
    }
    /// Declared total record size.
    #[must_use]
    pub const fn total_size(&self) -> u32 {
        self.total_size
    }
    /// Opaque record flags.
    #[must_use]
    pub const fn flags(&self) -> u32 {
        self.flags
    }
    /// Opaque hash algorithm identifier.
    #[must_use]
    pub const fn algorithm_id(&self) -> u32 {
        self.algorithm_id
    }
    /// Opaque spin-count field from the retained record.
    #[must_use]
    pub const fn spin_count(&self) -> u32 {
        self.spin_count
    }
    /// Logical hash length.
    #[must_use]
    pub const fn hash_size(&self) -> u32 {
        self.hash_size
    }
    /// Salt length.
    #[must_use]
    pub const fn salt_size(&self) -> u32 {
        self.salt_size
    }
    /// Salt bytes from the inert record.
    #[must_use]
    pub fn salt(&self) -> &[u8] {
        self.raw.get(self.salt_range.clone()).unwrap_or_default()
    }
    /// Logical hash bytes from the inert record.
    #[must_use]
    pub fn hash(&self) -> &[u8] {
        self.raw.get(self.hash_range.clone()).unwrap_or_default()
    }

    /// Validate the record and its configured resource bounds.
    pub fn validate(&self) -> RtfResult<()> {
        let encoded_size = self.raw.len();
        if encoded_size > MAX_PASSWORD_HASH_BYTES
            || self.total_size as usize != encoded_size
            || self.salt_size as usize != self.salt().len()
            || self.hash_size as usize != self.hash().len()
            || self.raw.get(self.salt_range.clone()).is_none()
            || self.raw.get(self.hash_range.clone()).is_none()
        {
            return Err(RtfError::MalformedDocument(
                "RTF passwordhash record has inconsistent sizes".to_string(),
            ));
        }
        Ok(())
    }

    #[must_use]
    pub fn into_owned(self) -> PasswordHash<'static> {
        PasswordHash {
            raw: Cow::Owned(self.raw.into_owned()),
            version: self.version,
            total_size: self.total_size,
            flags: self.flags,
            algorithm_id: self.algorithm_id,
            spin_count: self.spin_count,
            hash_size: self.hash_size,
            salt_size: self.salt_size,
            salt_range: self.salt_range,
            hash_range: self.hash_range,
        }
    }
}

fn parse_legacy_timestamp_triplet(
    value: &str,
    separator: char,
    invalid_component: &'static str,
) -> RtfResult<[i32; 3]> {
    let mut parts = value.split(separator);
    let parse = |part: Option<&str>| {
        part.ok_or_else(|| {
            RtfError::MalformedDocument("RTF info time must use YYYY-MM-DDTHH:MM:SS".to_string())
        })?
        .parse()
        .map_err(|_err| RtfError::MalformedDocument(invalid_component.to_string()))
    };
    let components = [
        parse(parts.next())?,
        parse(parts.next())?,
        parse(parts.next())?,
    ];
    if parts.next().is_some() {
        return Err(RtfError::MalformedDocument(
            "RTF info time must use YYYY-MM-DDTHH:MM:SS".to_string(),
        ));
    }
    Ok(components)
}

/// A possibly partial timestamp from an RTF information destination.
#[derive(Debug, Clone, Copy, Default, PartialEq, Eq)]
pub struct RtfTimestamp {
    /// Raw values are signed because legacy producers use invalid values such
    /// as zero as sentinels. Use [`Self::validate`] or [`Self::is_valid`]
    /// before interpreting a parsed value as a calendar timestamp.
    pub year: Option<i32>,
    pub month: Option<i32>,
    pub day: Option<i32>,
    pub hour: Option<i32>,
    pub minute: Option<i32>,
    pub second: Option<i32>,
}

impl RtfTimestamp {
    #[must_use]
    pub fn is_valid(&self) -> bool {
        self.validate().is_ok()
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn validate(&self) -> RtfResult<()> {
        if self.year.is_some_and(|value| value > 9999)
            || self.month.is_some_and(|value| !(1..=12).contains(&value))
            || self.day.is_some_and(|value| !(1..=31).contains(&value))
            || self.hour.is_some_and(|value| value > 23)
            || self.minute.is_some_and(|value| value > 59)
            || self.second.is_some_and(|value| value > 59)
        {
            return Err(RtfError::MalformedDocument(
                "RTF info timestamp component is outside its valid range".to_string(),
            ));
        }
        if let (Some(year), Some(month), Some(day)) = (self.year, self.month, self.day) {
            let leap = year % 4 == 0 && (year % 100 != 0 || year % 400 == 0);
            let max_day = match month {
                2 if leap => 29,
                2 => 28,
                4 | 6 | 9 | 11 => 30,
                _ => 31,
            };
            if day > max_day {
                return Err(RtfError::MalformedDocument(
                    "RTF info timestamp contains an invalid calendar date".to_string(),
                ));
            }
        }
        Ok(())
    }
    ///
    /// # Errors
    /// Returns an error when the input is malformed or a configured limit is exceeded.
    pub fn from_legacy(value: &str) -> RtfResult<Self> {
        let (date, time) = value.split_once('T').ok_or_else(|| {
            RtfError::MalformedDocument("RTF info time must contain T".to_string())
        })?;
        let [year, month, day] =
            parse_legacy_timestamp_triplet(date, '-', "invalid RTF info date")?;
        let [hour, minute, second] =
            parse_legacy_timestamp_triplet(time, ':', "invalid RTF info time")?;
        let timestamp = Self {
            year: Some(year),
            month: Some(month),
            day: Some(day),
            hour: Some(hour),
            minute: Some(minute),
            second: Some(second),
        };
        timestamp.validate()?;
        Ok(timestamp)
    }

    #[must_use]
    pub fn legacy_string(self) -> String {
        format!(
            "{:04}-{:02}-{:02}T{:02}:{:02}:{:02}",
            self.year.unwrap_or(0),
            self.month.unwrap_or(0),
            self.day.unwrap_or(0),
            self.hour.unwrap_or(0),
            self.minute.unwrap_or(0),
            self.second.unwrap_or(0),
        )
    }
}

/// Document information/metadata.
#[derive(Debug, Clone, Default, PartialEq, Eq)]
pub struct DocumentInfo<'a> {
    pub title: Option<Cow<'a, str>>,
    pub subject: Option<Cow<'a, str>>,
    pub author: Option<Cow<'a, str>>,
    pub manager: Option<Cow<'a, str>>,
    pub company: Option<Cow<'a, str>>,
    pub operator: Option<Cow<'a, str>>,
    pub category: Option<Cow<'a, str>>,
    pub keywords: Option<Cow<'a, str>>,
    pub comment: Option<Cow<'a, str>>,
    pub document_comment: Option<Cow<'a, str>>,
    pub hyperlink_base: Option<Cow<'a, str>>,
    pub version: Option<u32>,
    pub revision: Option<u32>,
    /// Legacy complete timestamp mirror retained for API compatibility.
    pub creation_time: Option<Cow<'a, str>>,
    pub creation_timestamp: Option<RtfTimestamp>,
    /// Legacy complete timestamp mirror retained for API compatibility.
    pub revision_time: Option<Cow<'a, str>>,
    pub revision_timestamp: Option<RtfTimestamp>,
    /// Legacy complete timestamp mirror retained for API compatibility.
    pub print_time: Option<Cow<'a, str>>,
    pub print_timestamp: Option<RtfTimestamp>,
    /// Legacy complete timestamp mirror retained for API compatibility.
    pub backup_time: Option<Cow<'a, str>>,
    pub backup_timestamp: Option<RtfTimestamp>,
    pub editing_time: Option<u32>,
    pub pages: Option<u32>,
    pub words: Option<u32>,
    pub characters: Option<u32>,
    pub characters_with_spaces: Option<u32>,
    pub id: Option<u32>,
    pub protection: DocumentProtection<'a>,
}

impl<'a> DocumentInfo<'a> {
    #[must_use]
    pub fn new() -> Self {
        Self::default()
    }
    #[must_use]
    pub fn with_title(mut self, value: Cow<'a, str>) -> Self {
        self.title = Some(value);
        self
    }
    #[must_use]
    pub fn with_author(mut self, value: Cow<'a, str>) -> Self {
        self.author = Some(value);
        self
    }
    #[must_use]
    pub fn with_subject(mut self, value: Cow<'a, str>) -> Self {
        self.subject = Some(value);
        self
    }
    #[must_use]
    pub fn with_keywords(mut self, value: Cow<'a, str>) -> Self {
        self.keywords = Some(value);
        self
    }
    #[must_use]
    pub fn with_comment(mut self, value: Cow<'a, str>) -> Self {
        self.comment = Some(value);
        self
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        for value in [
            self.title.as_deref(),
            self.subject.as_deref(),
            self.author.as_deref(),
            self.manager.as_deref(),
            self.company.as_deref(),
            self.operator.as_deref(),
            self.category.as_deref(),
            self.keywords.as_deref(),
            self.comment.as_deref(),
            self.document_comment.as_deref(),
            self.hyperlink_base.as_deref(),
        ]
        .into_iter()
        .flatten()
        {
            if value.len() > MAX_INFO_TEXT_BYTES {
                return Err(RtfError::MalformedDocument(
                    "RTF info text exceeds the metadata safety limit".to_string(),
                ));
            }
        }
        for value in [
            self.version,
            self.revision,
            self.editing_time,
            self.pages,
            self.words,
            self.characters,
            self.characters_with_spaces,
            self.id,
        ]
        .into_iter()
        .flatten()
        {
            if value > i32::MAX as u32 {
                return Err(RtfError::MalformedDocument(
                    "RTF info numeric value exceeds the signed control-word range".to_string(),
                ));
            }
        }
        for (typed, legacy) in [
            (self.creation_timestamp, self.creation_time.as_deref()),
            (self.revision_timestamp, self.revision_time.as_deref()),
            (self.print_timestamp, self.print_time.as_deref()),
            (self.backup_timestamp, self.backup_time.as_deref()),
        ] {
            if let Some(timestamp) = typed {
                if let Some(legacy_text) = legacy
                    && legacy_text != timestamp.legacy_string()
                {
                    return Err(RtfError::MalformedDocument(
                        "conflicting typed and legacy RTF info timestamps".to_string(),
                    ));
                } else if legacy.is_none() {
                    // A matching legacy mirror is parser provenance for raw
                    // producer values. Newly authored typed values are strict.
                    timestamp.validate()?;
                }
            } else if let Some(legacy_text) = legacy {
                RtfTimestamp::from_legacy(legacy_text)?;
            }
        }
        self.protection.validate()?;
        Ok(())
    }
}

/// Protection type for document.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ProtectionType {
    #[default]
    None,
    ReadOnly,
    RevisionTracking,
    Comments,
    Forms,
    All,
}

/// The bounded numeric value carried by `\protlevel`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ProtectionLevel {
    Level0,
    Level1,
    Level2,
    Level3,
}

impl ProtectionLevel {
    pub(crate) fn from_rtf(value: i32) -> RtfResult<Self> {
        match value {
            0 => Ok(Self::Level0),
            1 => Ok(Self::Level1),
            2 => Ok(Self::Level2),
            3 => Ok(Self::Level3),
            _ => Err(RtfError::MalformedDocument(
                "RTF protection level must be in 0..=3".to_string(),
            )),
        }
    }

    #[must_use]
    pub fn rtf_value(self) -> i32 {
        match self {
            Self::Level0 => 0,
            Self::Level1 => 1,
            Self::Level2 => 2,
            Self::Level3 => 3,
        }
    }
}

/// Document protection settings.
#[derive(Debug, Clone, PartialEq, Eq, Default)]
pub struct DocumentProtection<'a> {
    pub forms: Option<bool>,
    pub annotations: Option<bool>,
    pub revisions: Option<bool>,
    pub read_only: Option<bool>,
    pub all: Option<bool>,
    pub enforced: Option<bool>,
    pub level: Option<ProtectionLevel>,
    /// Exact inert hexadecimal payload from `\password`; never interpreted.
    pub password_hash: Option<Cow<'a, str>>,
    /// Modern inert binary record from the starred `\passwordhash`
    /// destination.  It is retained separately from the legacy eight-digit
    /// `\password` value and is never used to authenticate or decrypt.
    pub password_hash_data: Option<PasswordHash<'a>>,
}

impl DocumentProtection<'_> {
    #[must_use]
    pub fn new(protection_type: ProtectionType) -> Self {
        let mut protection = Self {
            enforced: Some(true),
            ..Self::default()
        };
        match protection_type {
            ProtectionType::None => {},
            ProtectionType::ReadOnly => protection.read_only = Some(true),
            ProtectionType::RevisionTracking => protection.revisions = Some(true),
            ProtectionType::Comments => protection.annotations = Some(true),
            ProtectionType::Forms => protection.forms = Some(true),
            ProtectionType::All => protection.all = Some(true),
        }
        protection
    }

    #[must_use]
    pub fn protection_type(&self) -> ProtectionType {
        if self.read_only == Some(true) {
            ProtectionType::ReadOnly
        } else if self.revisions == Some(true) {
            ProtectionType::RevisionTracking
        } else if self.annotations == Some(true) {
            ProtectionType::Comments
        } else if self.forms == Some(true) {
            ProtectionType::Forms
        } else if self.all == Some(true) {
            ProtectionType::All
        } else {
            ProtectionType::None
        }
    }

    #[must_use]
    pub fn is_protected(&self) -> bool {
        self.enforced != Some(false) && self.protection_type() != ProtectionType::None
    }

    pub(crate) fn validate(&self) -> RtfResult<()> {
        if let Some(hash) = &self.password_hash
            && (hash.len() != PROTECTION_PASSWORD_HASH_BYTES
                || !hash.as_bytes().iter().all(u8::is_ascii_hexdigit))
        {
            return Err(RtfError::MalformedDocument(
                "RTF protection password hash must contain exactly eight hexadecimal digits"
                    .to_string(),
            ));
        }
        if let Some(hash) = &self.password_hash_data {
            hash.validate()?;
        }
        Ok(())
    }

    /// Modern inert `\passwordhash` record, if present.
    #[must_use]
    pub fn modern_password_hash(&self) -> Option<&PasswordHash<'_>> {
        self.password_hash_data.as_ref()
    }

    #[must_use]
    pub fn into_owned(self) -> DocumentProtection<'static> {
        DocumentProtection {
            forms: self.forms,
            annotations: self.annotations,
            revisions: self.revisions,
            read_only: self.read_only,
            all: self.all,
            enforced: self.enforced,
            level: self.level,
            password_hash: self
                .password_hash
                .map(|value| Cow::Owned(value.into_owned())),
            password_hash_data: self.password_hash_data.map(PasswordHash::into_owned),
        }
    }
}
