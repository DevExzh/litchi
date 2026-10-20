//! The one document-metadata type shared by every format.
//!
//! [`Metadata`] is the union of the document properties the supported
//! formats store: the OPC core properties, the OOXML application
//! properties, the OLE summary information, the ODF `meta.xml` fields and
//! the iWork property lists. Each format reads the fields it stores and
//! leaves the others `None`. A write path stores what its format can carry
//! and refuses (or records a diagnostic for) a [`Field`] it cannot, so a
//! value is never dropped silently.
//!
//! Wire spellings stay at each codec: the author is `dc:creator` in OOXML
//! and `meta:initial-creator` in ODF, yet it is always
//! [`Metadata::author`] here. Timestamps are real time values
//! ([`Timestamp`]); lexical forms such as W3CDTF are a codec concern.
//!
//! Format-specific conversions (e.g. `From<crate::ole::OleMetadata>`) and
//! YAML front-matter serialization live in the umbrella crate at
//! `src/metadata_ext.rs`.

pub mod keywords;

use std::fmt;

use chrono::{DateTime, NaiveDateTime, SecondsFormat, Utc};
use serde::{Deserialize, Serialize};

pub use keywords::Keywords;

/// A document timestamp.
///
/// Formats either store an instant (an explicit UTC offset) or a
/// producer-local wall-clock time without a zone. The two stay distinct so a
/// zone-less value is never reinterpreted as UTC.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, Serialize, Deserialize)]
#[serde(untagged)]
pub enum Timestamp {
    /// An instant, normalized to UTC.
    Utc(DateTime<Utc>),
    /// A wall-clock time the source stores without a time zone.
    Unzoned(NaiveDateTime),
}

impl Timestamp {
    /// The instant, when the source stores one.
    #[must_use]
    pub const fn utc(self) -> Option<DateTime<Utc>> {
        match self {
            Self::Utc(value) => Some(value),
            Self::Unzoned(_) => None,
        }
    }

    /// The zone-less wall-clock time, when the source stores one.
    #[must_use]
    pub const fn unzoned(self) -> Option<NaiveDateTime> {
        match self {
            Self::Utc(_) => None,
            Self::Unzoned(value) => Some(value),
        }
    }

    /// The wall-clock reading: the UTC reading of an instant, or the stored
    /// zone-less time.
    #[must_use]
    pub fn naive(self) -> NaiveDateTime {
        match self {
            Self::Utc(value) => value.naive_utc(),
            Self::Unzoned(value) => value,
        }
    }

    /// Parse an RFC 3339 instant (`2024-01-02T03:04:05Z`, any offset), or a
    /// zone-less ISO 8601 date-time (`2024-01-02T03:04:05`).
    ///
    /// # Errors
    ///
    /// Returns [`TimestampError`] for any other text.
    pub fn parse(text: &str) -> Result<Self, TimestampError> {
        let text = text.trim();
        if let Ok(value) = DateTime::parse_from_rfc3339(text) {
            return Ok(Self::Utc(value.with_timezone(&Utc)));
        }
        NaiveDateTime::parse_from_str(text, "%Y-%m-%dT%H:%M:%S%.f")
            .map(Self::Unzoned)
            .map_err(|_| TimestampError(text.to_owned()))
    }
}

impl From<DateTime<Utc>> for Timestamp {
    fn from(value: DateTime<Utc>) -> Self {
        Self::Utc(value)
    }
}

impl From<NaiveDateTime> for Timestamp {
    fn from(value: NaiveDateTime) -> Self {
        Self::Unzoned(value)
    }
}

impl fmt::Display for Timestamp {
    /// RFC 3339 with `Z` for an instant; ISO 8601 without a zone otherwise.
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Utc(value) => {
                formatter.write_str(&value.to_rfc3339_opts(SecondsFormat::AutoSi, true))
            },
            Self::Unzoned(value) => write!(formatter, "{}", value.format("%Y-%m-%dT%H:%M:%S%.f")),
        }
    }
}

impl std::str::FromStr for Timestamp {
    type Err = TimestampError;

    fn from_str(text: &str) -> Result<Self, TimestampError> {
        Self::parse(text)
    }
}

/// Text that [`Timestamp::parse`] does not accept.
#[derive(Debug, Clone, PartialEq, Eq, thiserror::Error)]
#[error("'{0}' is not an RFC 3339 or ISO 8601 date-time")]
pub struct TimestampError(String);

/// Unified document metadata.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
pub struct Metadata {
    /// Document title
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    /// Document subject
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subject: Option<String>,
    /// Document author (OOXML `dc:creator`)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// Keywords, one entry per keyword; empty means none.
    ///
    /// Formats that store keywords as one string split it at `;` and `,`
    /// and write a changed list joined with `, `
    /// ([`keywords::split`], [`keywords::join`]); ODF stores one
    /// `meta:keyword` per entry. The OOXML codec keeps the structured
    /// [`Keywords`] mixed content of an unchanged list losslessly.
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub keywords: Vec<String>,
    /// Document description/comments
    #[serde(skip_serializing_if = "Option::is_none")]
    pub description: Option<String>,
    /// Stable document identifier
    #[serde(skip_serializing_if = "Option::is_none")]
    pub identifier: Option<String>,
    /// Document language tag
    #[serde(skip_serializing_if = "Option::is_none")]
    pub language: Option<String>,
    /// Template used to create the document
    #[serde(skip_serializing_if = "Option::is_none")]
    pub template: Option<String>,
    /// Last person to modify the document
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_modified_by: Option<String>,
    /// Revision lexical value
    #[serde(skip_serializing_if = "Option::is_none")]
    pub revision: Option<String>,
    /// Creation time
    #[serde(skip_serializing_if = "Option::is_none")]
    pub created: Option<Timestamp>,
    /// Last modification time
    #[serde(skip_serializing_if = "Option::is_none")]
    pub modified: Option<Timestamp>,
    /// Number of pages
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_count: Option<u32>,
    /// Number of words
    #[serde(skip_serializing_if = "Option::is_none")]
    pub word_count: Option<u32>,
    /// Number of characters
    #[serde(skip_serializing_if = "Option::is_none")]
    pub character_count: Option<u32>,
    /// Number of characters including spaces
    #[serde(skip_serializing_if = "Option::is_none")]
    pub character_count_with_spaces: Option<u32>,
    /// Total editing duration in minutes
    #[serde(skip_serializing_if = "Option::is_none")]
    pub editing_time_minutes: Option<u32>,
    /// Application that created the document
    #[serde(skip_serializing_if = "Option::is_none")]
    pub application: Option<String>,
    /// Document category
    #[serde(skip_serializing_if = "Option::is_none")]
    pub category: Option<String>,
    /// Company/organization
    #[serde(skip_serializing_if = "Option::is_none")]
    pub company: Option<String>,
    /// Manager name
    #[serde(skip_serializing_if = "Option::is_none")]
    pub manager: Option<String>,
    /// Content status (draft, final, etc.)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub content_status: Option<String>,
    /// Core-properties content type descriptor
    #[serde(skip_serializing_if = "Option::is_none")]
    pub content_type: Option<String>,
    /// Document version
    #[serde(skip_serializing_if = "Option::is_none")]
    pub version: Option<String>,
    /// Last printed time
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_printed: Option<Timestamp>,
    /// Last backup time
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_backup: Option<Timestamp>,
    /// Base URI used to resolve relative hyperlinks.
    ///
    /// This value is metadata only. Consumers decide whether and how to
    /// resolve it; the library does not follow the URI.
    #[serde(skip_serializing_if = "Option::is_none")]
    pub hyperlink_base: Option<String>,
    /// Security level
    #[serde(skip_serializing_if = "Option::is_none")]
    pub security: Option<u32>,
    /// Codepage for text encoding
    #[serde(skip_serializing_if = "Option::is_none")]
    pub codepage: Option<u32>,
}

/// One field of [`Metadata`], for errors and diagnostics that name what a
/// format can or cannot store.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash, PartialOrd, Ord)]
#[non_exhaustive]
pub enum Field {
    /// [`Metadata::title`]
    Title,
    /// [`Metadata::subject`]
    Subject,
    /// [`Metadata::author`]
    Author,
    /// [`Metadata::keywords`]
    Keywords,
    /// [`Metadata::description`]
    Description,
    /// [`Metadata::identifier`]
    Identifier,
    /// [`Metadata::language`]
    Language,
    /// [`Metadata::template`]
    Template,
    /// [`Metadata::last_modified_by`]
    LastModifiedBy,
    /// [`Metadata::revision`]
    Revision,
    /// [`Metadata::created`]
    Created,
    /// [`Metadata::modified`]
    Modified,
    /// [`Metadata::page_count`]
    PageCount,
    /// [`Metadata::word_count`]
    WordCount,
    /// [`Metadata::character_count`]
    CharacterCount,
    /// [`Metadata::character_count_with_spaces`]
    CharacterCountWithSpaces,
    /// [`Metadata::editing_time_minutes`]
    EditingTimeMinutes,
    /// [`Metadata::application`]
    Application,
    /// [`Metadata::category`]
    Category,
    /// [`Metadata::company`]
    Company,
    /// [`Metadata::manager`]
    Manager,
    /// [`Metadata::content_status`]
    ContentStatus,
    /// [`Metadata::content_type`]
    ContentType,
    /// [`Metadata::version`]
    Version,
    /// [`Metadata::last_printed`]
    LastPrinted,
    /// [`Metadata::last_backup`]
    LastBackup,
    /// [`Metadata::hyperlink_base`]
    HyperlinkBase,
    /// [`Metadata::security`]
    Security,
    /// [`Metadata::codepage`]
    Codepage,
}

impl Field {
    /// Every field, in declaration order.
    pub const ALL: [Self; 29] = [
        Self::Title,
        Self::Subject,
        Self::Author,
        Self::Keywords,
        Self::Description,
        Self::Identifier,
        Self::Language,
        Self::Template,
        Self::LastModifiedBy,
        Self::Revision,
        Self::Created,
        Self::Modified,
        Self::PageCount,
        Self::WordCount,
        Self::CharacterCount,
        Self::CharacterCountWithSpaces,
        Self::EditingTimeMinutes,
        Self::Application,
        Self::Category,
        Self::Company,
        Self::Manager,
        Self::ContentStatus,
        Self::ContentType,
        Self::Version,
        Self::LastPrinted,
        Self::LastBackup,
        Self::HyperlinkBase,
        Self::Security,
        Self::Codepage,
    ];

    /// The Rust field name, such as `last_modified_by`.
    #[must_use]
    pub const fn name(self) -> &'static str {
        match self {
            Self::Title => "title",
            Self::Subject => "subject",
            Self::Author => "author",
            Self::Keywords => "keywords",
            Self::Description => "description",
            Self::Identifier => "identifier",
            Self::Language => "language",
            Self::Template => "template",
            Self::LastModifiedBy => "last_modified_by",
            Self::Revision => "revision",
            Self::Created => "created",
            Self::Modified => "modified",
            Self::PageCount => "page_count",
            Self::WordCount => "word_count",
            Self::CharacterCount => "character_count",
            Self::CharacterCountWithSpaces => "character_count_with_spaces",
            Self::EditingTimeMinutes => "editing_time_minutes",
            Self::Application => "application",
            Self::Category => "category",
            Self::Company => "company",
            Self::Manager => "manager",
            Self::ContentStatus => "content_status",
            Self::ContentType => "content_type",
            Self::Version => "version",
            Self::LastPrinted => "last_printed",
            Self::LastBackup => "last_backup",
            Self::HyperlinkBase => "hyperlink_base",
            Self::Security => "security",
            Self::Codepage => "codepage",
        }
    }
}

impl fmt::Display for Field {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.name())
    }
}

impl Metadata {
    /// Check if the metadata contains any actual data.
    ///
    /// Returns true if at least one field is populated.
    #[must_use]
    pub fn has_data(&self) -> bool {
        Field::ALL.into_iter().any(|field| self.is_set(field))
    }

    /// Whether `field` holds a value.
    #[must_use]
    pub fn is_set(&self, field: Field) -> bool {
        match field {
            Field::Title => self.title.is_some(),
            Field::Subject => self.subject.is_some(),
            Field::Author => self.author.is_some(),
            Field::Keywords => !self.keywords.is_empty(),
            Field::Description => self.description.is_some(),
            Field::Identifier => self.identifier.is_some(),
            Field::Language => self.language.is_some(),
            Field::Template => self.template.is_some(),
            Field::LastModifiedBy => self.last_modified_by.is_some(),
            Field::Revision => self.revision.is_some(),
            Field::Created => self.created.is_some(),
            Field::Modified => self.modified.is_some(),
            Field::PageCount => self.page_count.is_some(),
            Field::WordCount => self.word_count.is_some(),
            Field::CharacterCount => self.character_count.is_some(),
            Field::CharacterCountWithSpaces => self.character_count_with_spaces.is_some(),
            Field::EditingTimeMinutes => self.editing_time_minutes.is_some(),
            Field::Application => self.application.is_some(),
            Field::Category => self.category.is_some(),
            Field::Company => self.company.is_some(),
            Field::Manager => self.manager.is_some(),
            Field::ContentStatus => self.content_status.is_some(),
            Field::ContentType => self.content_type.is_some(),
            Field::Version => self.version.is_some(),
            Field::LastPrinted => self.last_printed.is_some(),
            Field::LastBackup => self.last_backup.is_some(),
            Field::HyperlinkBase => self.hyperlink_base.is_some(),
            Field::Security => self.security.is_some(),
            Field::Codepage => self.codepage.is_some(),
        }
    }

    /// The fields whose values differ between `self` and `other`, in
    /// [`Field::ALL`] order.
    pub fn changed_fields<'a>(&'a self, other: &'a Self) -> impl Iterator<Item = Field> + 'a {
        Field::ALL
            .into_iter()
            .filter(move |field| !self.same_field(other, *field))
    }

    fn same_field(&self, other: &Self, field: Field) -> bool {
        match field {
            Field::Title => self.title == other.title,
            Field::Subject => self.subject == other.subject,
            Field::Author => self.author == other.author,
            Field::Keywords => self.keywords == other.keywords,
            Field::Description => self.description == other.description,
            Field::Identifier => self.identifier == other.identifier,
            Field::Language => self.language == other.language,
            Field::Template => self.template == other.template,
            Field::LastModifiedBy => self.last_modified_by == other.last_modified_by,
            Field::Revision => self.revision == other.revision,
            Field::Created => self.created == other.created,
            Field::Modified => self.modified == other.modified,
            Field::PageCount => self.page_count == other.page_count,
            Field::WordCount => self.word_count == other.word_count,
            Field::CharacterCount => self.character_count == other.character_count,
            Field::CharacterCountWithSpaces => {
                self.character_count_with_spaces == other.character_count_with_spaces
            },
            Field::EditingTimeMinutes => self.editing_time_minutes == other.editing_time_minutes,
            Field::Application => self.application == other.application,
            Field::Category => self.category == other.category,
            Field::Company => self.company == other.company,
            Field::Manager => self.manager == other.manager,
            Field::ContentStatus => self.content_status == other.content_status,
            Field::ContentType => self.content_type == other.content_type,
            Field::Version => self.version == other.version,
            Field::LastPrinted => self.last_printed == other.last_printed,
            Field::LastBackup => self.last_backup == other.last_backup,
            Field::HyperlinkBase => self.hyperlink_base == other.hyperlink_base,
            Field::Security => self.security == other.security,
            Field::Codepage => self.codepage == other.codepage,
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_metadata_has_data() {
        let empty_metadata = Metadata::default();
        assert!(!empty_metadata.has_data());

        let metadata_with_title = Metadata {
            title: Some("Test Document".to_string()),
            ..Default::default()
        };
        assert!(metadata_with_title.has_data());

        let metadata_with_unzoned_time = Metadata {
            created: chrono::NaiveDate::from_ymd_opt(2026, 7, 26)
                .and_then(|date| date.and_hms_opt(12, 0, 0))
                .map(Timestamp::Unzoned),
            ..Default::default()
        };
        assert!(metadata_with_unzoned_time.has_data());
    }

    #[test]
    fn changed_fields_names_exactly_the_differences() {
        let before = Metadata {
            title: Some("a".to_owned()),
            author: Some("x".to_owned()),
            ..Metadata::default()
        };
        let after = Metadata {
            title: Some("b".to_owned()),
            category: Some("c".to_owned()),
            ..before.clone()
        };
        assert_eq!(
            before.changed_fields(&after).collect::<Vec<_>>(),
            [Field::Title, Field::Category]
        );
        assert_eq!(Field::LastModifiedBy.to_string(), "last_modified_by");
    }

    #[test]
    fn timestamps_keep_zoned_and_unzoned_values_apart() {
        let zoned = Timestamp::parse("2024-01-02T08:04:05+05:00").unwrap();
        assert_eq!(zoned.to_string(), "2024-01-02T03:04:05Z");
        let unzoned = Timestamp::parse("2024-01-02T03:04:05").unwrap();
        assert!(unzoned.utc().is_none());
        assert_eq!(unzoned.to_string(), "2024-01-02T03:04:05");
        assert_eq!(zoned.naive(), unzoned.naive());
        assert!(Timestamp::parse("2024").is_err());

        let json = serde_json::to_string(&Metadata {
            created: Some(zoned),
            modified: Some(unzoned),
            ..Metadata::default()
        })
        .unwrap();
        let back: Metadata = serde_json::from_str(&json).unwrap();
        assert_eq!(back.created, Some(zoned));
        assert_eq!(back.modified, Some(unzoned));
    }
}
