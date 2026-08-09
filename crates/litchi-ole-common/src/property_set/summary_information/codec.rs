//! Typed PIDSI accessors over the generic property-set values.

use super::super::model::Value;
use super::model::{
    APP_NAME, AUTHOR, CHARACTER_COUNT, COMMENTS, CREATE_DTM, DOC_SECURITY, DocumentSecurity,
    EDIT_TIME, FileTime, KEYWORDS, LAST_AUTHOR, LAST_PRINTED, LAST_SAVE_DTM, PAGE_COUNT,
    REVISION_NUMBER, SUBJECT, Snapshot, TEMPLATE, THUMBNAIL, TITLE, ThumbnailRef, WORD_COUNT,
};

impl Snapshot {
    /// Returns the document title.
    #[must_use]
    pub fn title(&self) -> Option<&str> {
        string(self.property(TITLE))
    }

    /// Returns the document subject.
    #[must_use]
    pub fn subject(&self) -> Option<&str> {
        string(self.property(SUBJECT))
    }

    /// Returns the document author.
    #[must_use]
    pub fn author(&self) -> Option<&str> {
        string(self.property(AUTHOR))
    }

    /// Returns the document keywords.
    #[must_use]
    pub fn keywords(&self) -> Option<&str> {
        string(self.property(KEYWORDS))
    }

    /// Returns the document comments.
    #[must_use]
    pub fn comments(&self) -> Option<&str> {
        string(self.property(COMMENTS))
    }

    /// Returns the application-specific template name.
    #[must_use]
    pub fn template(&self) -> Option<&str> {
        string(self.property(TEMPLATE))
    }

    /// Returns the last author.
    #[must_use]
    pub fn last_author(&self) -> Option<&str> {
        string(self.property(LAST_AUTHOR))
    }

    /// Returns the application-specific revision number.
    #[must_use]
    pub fn revision_number(&self) -> Option<&str> {
        string(self.property(REVISION_NUMBER))
    }

    /// Returns the total editing time as an exact FILETIME value.
    #[must_use]
    pub fn edit_time(&self) -> Option<FileTime> {
        filetime(self.property(EDIT_TIME))
    }

    /// Returns the most-recently-printed timestamp.
    #[must_use]
    pub fn last_printed(&self) -> Option<FileTime> {
        filetime(self.property(LAST_PRINTED))
    }

    /// Returns the document-creation timestamp.
    #[must_use]
    pub fn create_time(&self) -> Option<FileTime> {
        filetime(self.property(CREATE_DTM))
    }

    /// Returns the most-recently-saved timestamp.
    #[must_use]
    pub fn last_save_time(&self) -> Option<FileTime> {
        filetime(self.property(LAST_SAVE_DTM))
    }

    /// Returns the nonnegative page count.
    #[must_use]
    pub fn page_count(&self) -> Option<u32> {
        count(self.property(PAGE_COUNT))
    }

    /// Returns the nonnegative word count.
    #[must_use]
    pub fn word_count(&self) -> Option<u32> {
        count(self.property(WORD_COUNT))
    }

    /// Returns the nonnegative character count.
    #[must_use]
    pub fn character_count(&self) -> Option<u32> {
        count(self.property(CHARACTER_COUNT))
    }

    /// Returns the optional thumbnail without allocating.
    #[must_use]
    pub fn thumbnail(&self) -> Option<ThumbnailRef<'_>> {
        ThumbnailRef::from_value(self.property(THUMBNAIL)?).ok()
    }

    /// Returns the creating application name.
    #[must_use]
    pub fn app_name(&self) -> Option<&str> {
        string(self.property(APP_NAME))
    }

    /// Returns the checked document-security flags.
    #[must_use]
    pub fn document_security(&self) -> Option<DocumentSecurity> {
        match self.property(DOC_SECURITY) {
            Some(Value::I4(value)) => Some(DocumentSecurity::new(u32::from_ne_bytes(
                value.to_ne_bytes(),
            ))),
            _ => None,
        }
    }
}

fn string(property: Option<&Value>) -> Option<&str> {
    match property {
        Some(Value::Lpstr(text) | Value::Lpwstr(text)) => Some(text),
        _ => None,
    }
}

fn filetime(property: Option<&Value>) -> Option<FileTime> {
    match property {
        Some(Value::Filetime(raw)) => Some(FileTime::from_raw(*raw)),
        _ => None,
    }
}

fn count(property: Option<&Value>) -> Option<u32> {
    match property {
        Some(Value::I4(raw_count)) => u32::try_from(*raw_count).ok(),
        _ => None,
    }
}
