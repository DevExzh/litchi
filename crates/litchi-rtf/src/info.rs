//! RTF document information and properties.

use crate::{RtfError, RtfResult};
use std::borrow::Cow;

pub(crate) const MAX_INFO_TEXT_BYTES: usize = 1_048_576;
pub(crate) const PROTECTION_PASSWORD_HASH_BYTES: usize = 8;

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

    pub fn from_legacy(value: &str) -> RtfResult<Self> {
        let (date, time) = value.split_once('T').ok_or_else(|| {
            RtfError::MalformedDocument("RTF info time must contain T".to_string())
        })?;
        let date: Vec<i32> = date
            .split('-')
            .map(str::parse)
            .collect::<Result<_, _>>()
            .map_err(|_| RtfError::MalformedDocument("invalid RTF info date".to_string()))?;
        let time: Vec<i32> = time
            .split(':')
            .map(str::parse)
            .collect::<Result<_, _>>()
            .map_err(|_| RtfError::MalformedDocument("invalid RTF info time".to_string()))?;
        if date.len() != 3 || time.len() != 3 {
            return Err(RtfError::MalformedDocument(
                "RTF info time must use YYYY-MM-DDTHH:MM:SS".to_string(),
            ));
        }
        let timestamp = Self {
            year: Some(date[0]),
            month: Some(date[1]),
            day: Some(date[2]),
            hour: Some(time[0]),
            minute: Some(time[1]),
            second: Some(time[2]),
        };
        timestamp.validate()?;
        Ok(timestamp)
    }

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
    pub fn new() -> Self {
        Self::default()
    }
    pub fn with_title(mut self, value: Cow<'a, str>) -> Self {
        self.title = Some(value);
        self
    }
    pub fn with_author(mut self, value: Cow<'a, str>) -> Self {
        self.author = Some(value);
        self
    }
    pub fn with_subject(mut self, value: Cow<'a, str>) -> Self {
        self.subject = Some(value);
        self
    }
    pub fn with_keywords(mut self, value: Cow<'a, str>) -> Self {
        self.keywords = Some(value);
        self
    }
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
                if let Some(legacy) = legacy
                    && legacy != timestamp.legacy_string()
                {
                    return Err(RtfError::MalformedDocument(
                        "conflicting typed and legacy RTF info timestamps".to_string(),
                    ));
                } else if legacy.is_none() {
                    // A matching legacy mirror is parser provenance for raw
                    // producer values. Newly authored typed values are strict.
                    timestamp.validate()?;
                }
            } else if let Some(legacy) = legacy {
                RtfTimestamp::from_legacy(legacy)?;
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
}

impl<'a> DocumentProtection<'a> {
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
        Ok(())
    }

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
        }
    }
}
