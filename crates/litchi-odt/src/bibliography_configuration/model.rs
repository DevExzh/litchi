//! Semantic bibliography policy and sort-key model.

use super::MAX_SORT_KEYS;
use litchi_core::{Error, Result};

/// A bibliography field used for document-wide entry ordering.
#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum Field {
    Identifier,
    BibliographyType,
    Address,
    Annote,
    Author,
    BookTitle,
    Chapter,
    Edition,
    Editor,
    HowPublished,
    Institution,
    Journal,
    Month,
    Note,
    Number,
    Organizations,
    Pages,
    Publisher,
    School,
    Series,
    Title,
    ReportType,
    Volume,
    Year,
    Url,
    Custom1,
    Custom2,
    Custom3,
    Custom4,
    Custom5,
    Isbn,
    Issn,
}

impl Field {
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Identifier => "identifier",
            Self::BibliographyType => "bibliography-type",
            Self::Address => "address",
            Self::Annote => "annote",
            Self::Author => "author",
            Self::BookTitle => "booktitle",
            Self::Chapter => "chapter",
            Self::Edition => "edition",
            Self::Editor => "editor",
            Self::HowPublished => "howpublished",
            Self::Institution => "institution",
            Self::Journal => "journal",
            Self::Month => "month",
            Self::Note => "note",
            Self::Number => "number",
            Self::Organizations => "organizations",
            Self::Pages => "pages",
            Self::Publisher => "publisher",
            Self::School => "school",
            Self::Series => "series",
            Self::Title => "title",
            Self::ReportType => "report-type",
            Self::Volume => "volume",
            Self::Year => "year",
            Self::Url => "url",
            Self::Custom1 => "custom1",
            Self::Custom2 => "custom2",
            Self::Custom3 => "custom3",
            Self::Custom4 => "custom4",
            Self::Custom5 => "custom5",
            Self::Isbn => "isbn",
            Self::Issn => "issn",
        }
    }

    pub(super) fn parse(value: &str) -> Result<Self> {
        Ok(match value {
            "identifier" => Self::Identifier,
            "bibliography-type" => Self::BibliographyType,
            "address" => Self::Address,
            "annote" => Self::Annote,
            "author" => Self::Author,
            "booktitle" => Self::BookTitle,
            "chapter" => Self::Chapter,
            "edition" => Self::Edition,
            "editor" => Self::Editor,
            "howpublished" => Self::HowPublished,
            "institution" => Self::Institution,
            "journal" => Self::Journal,
            "month" => Self::Month,
            "note" => Self::Note,
            "number" => Self::Number,
            "organizations" => Self::Organizations,
            "pages" => Self::Pages,
            "publisher" => Self::Publisher,
            "school" => Self::School,
            "series" => Self::Series,
            "title" => Self::Title,
            "report-type" => Self::ReportType,
            "volume" => Self::Volume,
            "year" => Self::Year,
            "url" => Self::Url,
            "custom1" => Self::Custom1,
            "custom2" => Self::Custom2,
            "custom3" => Self::Custom3,
            "custom4" => Self::Custom4,
            "custom5" => Self::Custom5,
            "isbn" => Self::Isbn,
            "issn" => Self::Issn,
            _ => return invalid(format!("invalid bibliography sort key '{value}'")),
        })
    }
}

/// One ordered bibliography sort key.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct SortKey {
    pub field: Field,
    pub ascending: Option<bool>,
}

impl SortKey {
    /// ODF defaults `text:sort-ascending` to `true`.
    pub fn effective_ascending(&self) -> bool {
        self.ascending.unwrap_or(true)
    }
}

/// Document-wide bibliography formatting and ordering policy.
#[derive(Clone, Debug, Default, PartialEq, Eq)]
pub struct Configuration {
    pub prefix: Option<String>,
    pub suffix: Option<String>,
    pub numbered_entries: Option<bool>,
    pub sort_by_position: Option<bool>,
    pub sort_algorithm: Option<String>,
    pub language: Option<String>,
    pub country: Option<String>,
    pub script: Option<String>,
    pub rfc_language_tag: Option<String>,
    pub sort_keys: Vec<SortKey>,
}

impl Configuration {
    pub fn effective_numbered_entries(&self) -> bool {
        self.numbered_entries.unwrap_or(false)
    }

    pub fn effective_sort_by_position(&self) -> bool {
        self.sort_by_position.unwrap_or(true)
    }

    /// Validate serializable bibliography policy metadata.
    pub fn validate(&self) -> Result<()> {
        if self.sort_keys.len() > MAX_SORT_KEYS {
            return invalid("bibliography configuration has too many sort keys");
        }
        for (value, context) in [
            (&self.prefix, "bibliography prefix"),
            (&self.suffix, "bibliography suffix"),
            (&self.sort_algorithm, "bibliography sort algorithm"),
        ] {
            if let Some(value) = value {
                checked_value(value, context)?;
            }
        }
        if let Some(value) = &self.language {
            validate_language_code(value, "fo:language")?;
        }
        if let Some(value) = &self.country {
            validate_alphanumeric_code(value, "fo:country")?;
        }
        if let Some(value) = &self.script {
            validate_alphanumeric_code(value, "fo:script")?;
        }
        if let Some(value) = &self.rfc_language_tag {
            validate_language_tag(value)?;
        }
        Ok(())
    }
}

fn checked_value(value: &str, context: &str) -> Result<()> {
    if value.len() > super::MAX_VALUE_BYTES {
        invalid(format!("{context} exceeds 64 KiB"))
    } else {
        Ok(())
    }
}

fn validate_language_code(value: &str, context: &str) -> Result<()> {
    if (1..=8).contains(&value.len()) && value.bytes().all(|byte| byte.is_ascii_alphabetic()) {
        Ok(())
    } else {
        invalid(format!("invalid {context} lexical '{value}'"))
    }
}

fn validate_alphanumeric_code(value: &str, context: &str) -> Result<()> {
    if (1..=8).contains(&value.len()) && value.bytes().all(|byte| byte.is_ascii_alphanumeric()) {
        Ok(())
    } else {
        invalid(format!("invalid {context} lexical '{value}'"))
    }
}

fn validate_language_tag(value: &str) -> Result<()> {
    if value.split('-').all(|part| {
        (1..=8).contains(&part.len()) && part.bytes().all(|byte| byte.is_ascii_alphanumeric())
    }) {
        Ok(())
    } else {
        invalid(format!("invalid style:rfc-language-tag lexical '{value}'"))
    }
}

fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(Error::InvalidFormat(message.into()))
}

// These aliases are consumed by the unchanged crate-root facade. New code
// within this owner uses the shorter contextual vocabulary above.
