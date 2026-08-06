//! Typed PIDSI accessors over the generic property-set values.

use super::super::model::Value;
use super::model::*;

impl Snapshot {
    /// Returns the document title.
    pub fn title(&self) -> Option<&str> {
        string(self.property(TITLE))
    }

    /// Returns the document subject.
    pub fn subject(&self) -> Option<&str> {
        string(self.property(SUBJECT))
    }

    /// Returns the document author.
    pub fn author(&self) -> Option<&str> {
        string(self.property(AUTHOR))
    }

    /// Returns the document keywords.
    pub fn keywords(&self) -> Option<&str> {
        string(self.property(KEYWORDS))
    }

    /// Returns the document comments.
    pub fn comments(&self) -> Option<&str> {
        string(self.property(COMMENTS))
    }

    /// Returns the application-specific template name.
    pub fn template(&self) -> Option<&str> {
        string(self.property(TEMPLATE))
    }

    /// Returns the last author.
    pub fn last_author(&self) -> Option<&str> {
        string(self.property(LAST_AUTHOR))
    }

    /// Returns the application-specific revision number.
    pub fn revision_number(&self) -> Option<&str> {
        string(self.property(REVISION_NUMBER))
    }

    /// Returns the total editing time as an exact FILETIME value.
    pub fn edit_time(&self) -> Option<FileTime> {
        filetime(self.property(EDIT_TIME))
    }

    /// Returns the most-recently-printed timestamp.
    pub fn last_printed(&self) -> Option<FileTime> {
        filetime(self.property(LAST_PRINTED))
    }

    /// Returns the document-creation timestamp.
    pub fn create_time(&self) -> Option<FileTime> {
        filetime(self.property(CREATE_DTM))
    }

    /// Returns the most-recently-saved timestamp.
    pub fn last_save_time(&self) -> Option<FileTime> {
        filetime(self.property(LAST_SAVE_DTM))
    }

    /// Returns the nonnegative page count.
    pub fn page_count(&self) -> Option<u32> {
        count(self.property(PAGE_COUNT))
    }

    /// Returns the nonnegative word count.
    pub fn word_count(&self) -> Option<u32> {
        count(self.property(WORD_COUNT))
    }

    /// Returns the nonnegative character count.
    pub fn character_count(&self) -> Option<u32> {
        count(self.property(CHARACTER_COUNT))
    }

    /// Returns the optional thumbnail without allocating.
    pub fn thumbnail(&self) -> Option<ThumbnailRef<'_>> {
        ThumbnailRef::from_value(self.property(THUMBNAIL)?).ok()
    }

    /// Returns the creating application name.
    pub fn app_name(&self) -> Option<&str> {
        string(self.property(APP_NAME))
    }

    /// Returns the checked document-security flags.
    pub fn document_security(&self) -> Option<DocumentSecurity> {
        match self.property(DOC_SECURITY) {
            Some(Value::I4(value)) => Some(DocumentSecurity::new(*value as u32)),
            _ => None,
        }
    }
}

fn string(value: Option<&Value>) -> Option<&str> {
    match value {
        Some(Value::Lpstr(value) | Value::Lpwstr(value)) => Some(value),
        _ => None,
    }
}

fn filetime(value: Option<&Value>) -> Option<FileTime> {
    match value {
        Some(Value::Filetime(value)) => Some(FileTime::from_raw(*value)),
        _ => None,
    }
}

fn count(value: Option<&Value>) -> Option<u32> {
    match value {
        Some(Value::I4(value)) => u32::try_from(*value).ok(),
        _ => None,
    }
}
