//! The one conversion from the unified [`litchi_core::Metadata`] into the
//! OLE property sets.
//!
//! Every read of a `litchi_core::Metadata` field by the legacy writers and
//! metadata transactions happens here, so a change to the unified type is a
//! local edit of [`SummaryProperties::merge_metadata`] and its time helpers.
//! The destructuring is exhaustive on purpose: a new unified field fails to
//! compile until it is mapped or refused, so no field is dropped silently.

use super::super::document_summary;
use super::super::model::CodePage;
use super::super::summary_information::{self, DocumentSecurity, FileTime};
use super::model::SummaryProperties;
use litchi_cfb::OleError;
use litchi_core::Metadata;
use litchi_core::metadata::{Timestamp, keywords};
use std::fmt;

/// FILETIME ticks (100 ns) per minute of editing time.
const TICKS_PER_MINUTE: u64 = 600_000_000;

/// Why a [`Metadata`] value cannot be stored in the OLE property sets.
///
/// The field is named by its unified key, structural metadata only; the
/// refused value is never carried.
#[derive(Debug)]
#[non_exhaustive]
pub enum MetadataError {
    /// The property sets have no place for the field: `identifier`,
    /// `hyperlink_base`, and timestamps without a time zone (FILETIME is
    /// UTC).
    Unsupported {
        /// The unified field name.
        field: &'static str,
    },
    /// The property refused the value (a revision number that is not a whole
    /// number, text with NUL or over 1 MiB, a time outside FILETIME, an
    /// unknown code page).
    Invalid {
        /// The unified field name.
        field: &'static str,
        /// The property-set refusal.
        source: OleError,
    },
}

impl MetadataError {
    /// The unified field name that was refused.
    #[must_use]
    pub const fn field(&self) -> &'static str {
        match self {
            Self::Unsupported { field } | Self::Invalid { field, .. } => field,
        }
    }
}

impl fmt::Display for MetadataError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Unsupported { field } => write!(
                formatter,
                "metadata field `{field}` cannot be stored in OLE property sets"
            ),
            Self::Invalid { field, .. } => {
                write!(formatter, "metadata field `{field}` has an invalid value")
            },
        }
    }
}

impl std::error::Error for MetadataError {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Unsupported { .. } => None,
            Self::Invalid { source, .. } => Some(source),
        }
    }
}

impl SummaryProperties {
    /// Builds the properties of a new document from unified metadata.
    ///
    /// # Errors
    ///
    /// See [`Self::merge_metadata`].
    pub fn from_metadata(metadata: &Metadata) -> Result<Self, MetadataError> {
        let mut properties = Self::new();
        properties.merge_metadata(metadata)?;
        Ok(properties)
    }

    /// Sets every field `metadata` carries, leaving the properties it does
    /// not carry (`None`) as they are.
    ///
    /// | Unified field | Property |
    /// |---|---|
    /// | `title`, `subject`, `author`, `keywords`, `template`, `application` | the PIDSI text of that name |
    /// | `description` | PIDSI `COMMENTS` |
    /// | `last_modified_by` | PIDSI `LASTAUTHOR` |
    /// | `revision` | PIDSI `REVNUMBER` (a whole number) |
    /// | `created`, `modified`, `last_printed` | PIDSI `CREATE_DTM`, `LASTSAVE_DTM`, `LASTPRINTED` |
    /// | `editing_time_minutes` | PIDSI `EDITTIME` |
    /// | `page_count`, `word_count`, `character_count` | PIDSI counts |
    /// | `security` | PIDSI `DOC_SECURITY` |
    /// | `category`, `manager`, `company`, `content_type`, `content_status`, `language` | the PIDDSI text of that name |
    /// | `version` | PIDDSI `DOCVERSION` |
    /// | `character_count_with_spaces` | PIDDSI `CCHWITHSPACES` |
    /// | `codepage` | both standard sections' code page, applied first |
    ///
    /// Text the code page cannot represent switches its section to UTF-16, as
    /// [`Self::set`] does. `identifier`, `hyperlink_base`, `last_backup` and
    /// unzoned timestamps ([`Timestamp::Unzoned`]) have no place in the
    /// property sets and are refused.
    ///
    /// The merge is atomic: on any error nothing changes. Returns whether a
    /// property changed.
    ///
    /// # Errors
    ///
    /// Returns [`MetadataError::Unsupported`] for a field the property sets
    /// cannot store and [`MetadataError::Invalid`] for a refused value.
    pub fn merge_metadata(&mut self, metadata: &Metadata) -> Result<bool, MetadataError> {
        let Metadata {
            title,
            subject,
            author,
            keywords,
            description,
            identifier,
            language,
            template,
            last_modified_by,
            revision,
            created,
            modified,
            page_count,
            word_count,
            character_count,
            character_count_with_spaces,
            editing_time_minutes,
            application,
            category,
            company,
            manager,
            content_status,
            content_type,
            version,
            last_printed,
            last_backup,
            hyperlink_base,
            security,
            codepage,
        } = metadata;

        refuse_present("identifier", identifier.as_ref())?;
        refuse_present("hyperlink_base", hyperlink_base.as_ref())?;
        let created = utc_time("created", *created)?;
        let modified = utc_time("modified", *modified)?;
        let last_printed = utc_time("last_printed", *last_printed)?;
        refuse_present("last_backup", last_backup.as_ref())?;

        let keywords = self.keyword_text(keywords)?;
        let mut draft = self.clone();
        if let Some(page) = *codepage {
            draft.set_code_page(page)?;
        }
        let texts: [(&'static str, &'static str, &Option<String>); 16] = [
            ("title", "title", title),
            ("subject", "subject", subject),
            ("author", "author", author),
            ("keywords", "keywords", &keywords),
            ("description", "comments", description),
            ("template", "template", template),
            ("last_modified_by", "lastauthor", last_modified_by),
            ("revision", "revisionnumber", revision),
            ("application", "applicationname", application),
            ("category", "category", category),
            ("manager", "manager", manager),
            ("company", "company", company),
            ("content_type", "contenttype", content_type),
            ("content_status", "contentstatus", content_status),
            ("language", "language", language),
            ("version", "documentversion", version),
        ];
        for (field, name, value) in texts {
            if let Some(value) = value {
                draft
                    .set(name, value)
                    .map_err(|source| invalid(field, source))?;
            }
        }
        let times: [(&'static str, Option<FileTime>, TimeSetter); 3] = [
            ("created", created, |edit, value| {
                edit.set_create_time(value)
            }),
            ("modified", modified, |edit, value| {
                edit.set_last_save_time(value)
            }),
            ("last_printed", last_printed, |edit, value| {
                edit.set_last_printed(value)
            }),
        ];
        for (field, value, setter) in times {
            if let Some(value) = value {
                draft
                    .edit_summary_information(|edit| setter(edit, value))
                    .map_err(|source| invalid(field, source))?;
            }
        }
        if let Some(minutes) = *editing_time_minutes {
            let ticks = FileTime::from_raw(u64::from(minutes).saturating_mul(TICKS_PER_MINUTE));
            draft
                .edit_summary_information(|edit| edit.set_edit_time(ticks))
                .map_err(|source| invalid("editing_time_minutes", source))?;
        }
        let counts: [(&'static str, Option<u32>, CountSetter); 3] = [
            ("page_count", *page_count, |edit, value| {
                edit.set_page_count(value)
            }),
            ("word_count", *word_count, |edit, value| {
                edit.set_word_count(value)
            }),
            ("character_count", *character_count, |edit, value| {
                edit.set_character_count(value)
            }),
        ];
        for (field, value, setter) in counts {
            if let Some(value) = value {
                draft
                    .edit_summary_information(|edit| setter(edit, value))
                    .map_err(|source| invalid(field, source))?;
            }
        }
        if let Some(bits) = *security {
            draft
                .edit_summary_information(|edit| {
                    edit.set_document_security(DocumentSecurity::new(bits))
                })
                .map_err(|source| invalid("security", source))?;
        }
        if let Some(count) = *character_count_with_spaces {
            let count = i32::try_from(count).map_err(|_| {
                invalid(
                    "character_count_with_spaces",
                    super::super::model::invalid("count exceeds VT_I4"),
                )
            })?;
            draft
                .edit_document_summary_information(|edit: &mut document_summary::Edit<'_>| {
                    edit.set_character_count_with_spaces(count)
                })
                .map_err(|source| invalid("character_count_with_spaces", source))?;
        }
        let changed = draft != *self;
        *self = draft;
        Ok(changed)
    }

    /// The PIDSI `KEYWORDS` text to write for `list`: `None` for an empty
    /// list (left as it is) and for the list the current text already
    /// splits into (its bytes are kept), else the list joined with `, `.
    fn keyword_text(&self, list: &[String]) -> Result<Option<String>, MetadataError> {
        if list.is_empty() {
            return Ok(None);
        }
        let current = self.summary_information().keywords().map(keywords::split);
        if current.as_deref() == Some(list) {
            return Ok(None);
        }
        keywords::join(list).map(Some).map_err(|_| {
            invalid(
                "keywords",
                super::super::model::invalid(
                    "a keyword contains ';' or ',' or surrounding whitespace, which the keyword string cannot keep",
                ),
            )
        })
    }

    /// Selects one code page for both standard sections.
    fn set_code_page(&mut self, page: u32) -> Result<(), MetadataError> {
        let page = u16::try_from(page)
            .ok()
            .and_then(CodePage::new)
            .ok_or_else(|| {
                invalid(
                    "codepage",
                    super::super::model::invalid("unsupported property-set code page"),
                )
            })?;
        self.edit_summary_information(|edit| edit.set_codepage(page))
            .map_err(|source| invalid("codepage", source))?;
        self.edit_document_summary_information(|edit| {
            edit.set_codepage(page);
            Ok(())
        })
        .map_err(|source| invalid("codepage", source))?;
        Ok(())
    }
}

fn invalid(field: &'static str, source: OleError) -> MetadataError {
    MetadataError::Invalid { field, source }
}

fn refuse_present<T>(field: &'static str, value: Option<&T>) -> Result<(), MetadataError> {
    match value {
        Some(_) => Err(MetadataError::Unsupported { field }),
        None => Ok(()),
    }
}

/// One unified timestamp as a FILETIME. FILETIME is UTC, so only the zoned
/// form is stored; an unzoned timestamp is refused rather than reinterpreted.
fn utc_time(
    field: &'static str,
    time: Option<Timestamp>,
) -> Result<Option<FileTime>, MetadataError> {
    match time {
        None => Ok(None),
        Some(Timestamp::Unzoned(_)) => Err(MetadataError::Unsupported { field }),
        Some(Timestamp::Utc(utc)) => FileTime::from_date_time(utc)
            .map(Some)
            .map_err(|source| invalid(field, source)),
    }
}

type TimeSetter = fn(&mut summary_information::Edit<'_>, FileTime) -> Result<(), OleError>;
type CountSetter = fn(&mut summary_information::Edit<'_>, u32) -> Result<(), OleError>;
