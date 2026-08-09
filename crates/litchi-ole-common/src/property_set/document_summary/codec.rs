//! Typed PIDDSI accessors.

use super::super::model::{CodePage, DocParts, HeadingPairs, Value};
use super::model::{
    BYTE_COUNT, CATEGORY, CHARACTER_COUNT_WITH_SPACES, COMPANY, CONTENT_STATUS, CONTENT_TYPE,
    DOCUMENT_PARTS, DOCUMENT_VERSION, HEADING_PAIRS, HIDDEN_COUNT, HYPERLINKS_CHANGED, LANGUAGE,
    LINE_COUNT, LINKS_DIRTY, MANAGER, MULTIMEDIA_CLIP_COUNT, NOTE_COUNT, PARAGRAPH_COUNT,
    PRESENTATION_FORMAT, SCALE, SHARED_DOCUMENT, SLIDE_COUNT, Snapshot, VERSION, Version,
};

impl Snapshot {
    /// Returns the required section code page.
    #[must_use]
    pub const fn codepage(&self) -> Option<CodePage> {
        self.section.page()
    }

    /// Returns the category string.
    #[must_use]
    pub fn category(&self) -> Option<&str> {
        string(self.property(CATEGORY))
    }

    /// Returns the presentation-format string.
    #[must_use]
    pub fn presentation_format(&self) -> Option<&str> {
        string(self.property(PRESENTATION_FORMAT))
    }

    /// Returns the estimated document size in bytes.
    #[must_use]
    pub fn byte_count(&self) -> Option<i32> {
        i4(self.property(BYTE_COUNT))
    }

    /// Returns the estimated text-line count.
    #[must_use]
    pub fn line_count(&self) -> Option<i32> {
        i4(self.property(LINE_COUNT))
    }

    /// Returns the paragraph count.
    #[must_use]
    pub fn paragraph_count(&self) -> Option<i32> {
        i4(self.property(PARAGRAPH_COUNT))
    }

    /// Returns the slide count.
    #[must_use]
    pub fn slide_count(&self) -> Option<i32> {
        i4(self.property(SLIDE_COUNT))
    }

    /// Returns the note count.
    #[must_use]
    pub fn note_count(&self) -> Option<i32> {
        i4(self.property(NOTE_COUNT))
    }

    /// Returns the hidden-slide count.
    #[must_use]
    pub fn hidden_count(&self) -> Option<i32> {
        i4(self.property(HIDDEN_COUNT))
    }

    /// Returns the multimedia-clip count.
    #[must_use]
    pub fn multimedia_clip_count(&self) -> Option<i32> {
        i4(self.property(MULTIMEDIA_CLIP_COUNT))
    }

    /// Returns the MS-OSHARED scale flag, which is always false when present.
    #[must_use]
    pub fn scale(&self) -> Option<bool> {
        boolean(self.property(SCALE))
    }

    /// Returns the ordered heading pairs.
    #[must_use]
    pub fn heading_pairs(&self) -> Option<&HeadingPairs> {
        match self.property(HEADING_PAIRS) {
            Some(Value::HeadingPairs(value)) => Some(value),
            _ => None,
        }
    }

    /// Returns the ordered document-part names.
    #[must_use]
    pub fn document_parts(&self) -> Option<&DocParts> {
        match self.property(DOCUMENT_PARTS) {
            Some(Value::DocParts(value)) => Some(value),
            _ => None,
        }
    }

    /// Returns the manager string.
    #[must_use]
    pub fn manager(&self) -> Option<&str> {
        string(self.property(MANAGER))
    }

    /// Returns the company string.
    #[must_use]
    pub fn company(&self) -> Option<&str> {
        string(self.property(COMPANY))
    }

    /// Returns whether linked user-defined properties are dirty.
    #[must_use]
    pub fn links_dirty(&self) -> Option<bool> {
        boolean(self.property(LINKS_DIRTY))
    }

    /// Returns the estimated character count including spaces.
    #[must_use]
    pub fn character_count_with_spaces(&self) -> Option<i32> {
        i4(self.property(CHARACTER_COUNT_WITH_SPACES))
    }

    /// Returns the shared-document flag.
    #[must_use]
    pub fn shared_document(&self) -> Option<bool> {
        boolean(self.property(SHARED_DOCUMENT))
    }

    /// Returns whether the hyperlink property changed externally.
    #[must_use]
    pub fn hyperlinks_changed(&self) -> Option<bool> {
        boolean(self.property(HYPERLINKS_CHANGED))
    }

    /// Returns the decoded application version.
    #[must_use]
    pub fn version(&self) -> Option<Version> {
        match self.property(VERSION) {
            Some(Value::I4(value)) => Version::from_raw(*value).ok(),
            _ => None,
        }
    }

    /// Returns the content-type string.
    #[must_use]
    pub fn content_type(&self) -> Option<&str> {
        string(self.property(CONTENT_TYPE))
    }

    /// Returns the document-status string.
    #[must_use]
    pub fn content_status(&self) -> Option<&str> {
        string(self.property(CONTENT_STATUS))
    }

    /// Returns the language string when a producer supplied one.
    #[must_use]
    pub fn language(&self) -> Option<&str> {
        string(self.property(LANGUAGE))
    }

    /// Returns the producer-specific document-version string.
    #[must_use]
    pub fn document_version(&self) -> Option<&str> {
        string(self.property(DOCUMENT_VERSION))
    }
}

fn string(property: Option<&Value>) -> Option<&str> {
    match property {
        Some(Value::Lpstr(text) | Value::Lpwstr(text)) => Some(text),
        _ => None,
    }
}

fn i4(property: Option<&Value>) -> Option<i32> {
    match property {
        Some(Value::I4(count)) => Some(*count),
        _ => None,
    }
}

fn boolean(property: Option<&Value>) -> Option<bool> {
    match property {
        Some(Value::Bool(enabled)) => Some(*enabled),
        _ => None,
    }
}
