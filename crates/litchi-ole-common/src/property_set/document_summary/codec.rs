//! Typed PIDDSI accessors.

use super::super::model::{CodePage, DocParts, HeadingPairs, Value};
use super::model::*;

impl Snapshot {
    /// Returns the required section code page.
    pub const fn codepage(&self) -> Option<CodePage> {
        self.section.page()
    }

    /// Returns the category string.
    pub fn category(&self) -> Option<&str> {
        string(self.property(CATEGORY))
    }

    /// Returns the presentation-format string.
    pub fn presentation_format(&self) -> Option<&str> {
        string(self.property(PRESENTATION_FORMAT))
    }

    /// Returns the estimated document size in bytes.
    pub fn byte_count(&self) -> Option<i32> {
        i4(self.property(BYTE_COUNT))
    }

    /// Returns the estimated text-line count.
    pub fn line_count(&self) -> Option<i32> {
        i4(self.property(LINE_COUNT))
    }

    /// Returns the paragraph count.
    pub fn paragraph_count(&self) -> Option<i32> {
        i4(self.property(PARAGRAPH_COUNT))
    }

    /// Returns the slide count.
    pub fn slide_count(&self) -> Option<i32> {
        i4(self.property(SLIDE_COUNT))
    }

    /// Returns the note count.
    pub fn note_count(&self) -> Option<i32> {
        i4(self.property(NOTE_COUNT))
    }

    /// Returns the hidden-slide count.
    pub fn hidden_count(&self) -> Option<i32> {
        i4(self.property(HIDDEN_COUNT))
    }

    /// Returns the multimedia-clip count.
    pub fn multimedia_clip_count(&self) -> Option<i32> {
        i4(self.property(MULTIMEDIA_CLIP_COUNT))
    }

    /// Returns the MS-OSHARED scale flag, which is always false when present.
    pub fn scale(&self) -> Option<bool> {
        boolean(self.property(SCALE))
    }

    /// Returns the ordered heading pairs.
    pub fn heading_pairs(&self) -> Option<&HeadingPairs> {
        match self.property(HEADING_PAIRS) {
            Some(Value::HeadingPairs(value)) => Some(value),
            _ => None,
        }
    }

    /// Returns the ordered document-part names.
    pub fn document_parts(&self) -> Option<&DocParts> {
        match self.property(DOCUMENT_PARTS) {
            Some(Value::DocParts(value)) => Some(value),
            _ => None,
        }
    }

    /// Returns the manager string.
    pub fn manager(&self) -> Option<&str> {
        string(self.property(MANAGER))
    }

    /// Returns the company string.
    pub fn company(&self) -> Option<&str> {
        string(self.property(COMPANY))
    }

    /// Returns whether linked user-defined properties are dirty.
    pub fn links_dirty(&self) -> Option<bool> {
        boolean(self.property(LINKS_DIRTY))
    }

    /// Returns the estimated character count including spaces.
    pub fn character_count_with_spaces(&self) -> Option<i32> {
        i4(self.property(CHARACTER_COUNT_WITH_SPACES))
    }

    /// Returns the shared-document flag.
    pub fn shared_document(&self) -> Option<bool> {
        boolean(self.property(SHARED_DOCUMENT))
    }

    /// Returns whether the hyperlink property changed externally.
    pub fn hyperlinks_changed(&self) -> Option<bool> {
        boolean(self.property(HYPERLINKS_CHANGED))
    }

    /// Returns the decoded application version.
    pub fn version(&self) -> Option<Version> {
        match self.property(VERSION) {
            Some(Value::I4(value)) => Version::from_raw(*value).ok(),
            _ => None,
        }
    }

    /// Returns the content-type string.
    pub fn content_type(&self) -> Option<&str> {
        string(self.property(CONTENT_TYPE))
    }

    /// Returns the document-status string.
    pub fn content_status(&self) -> Option<&str> {
        string(self.property(CONTENT_STATUS))
    }

    /// Returns the language string when a producer supplied one.
    pub fn language(&self) -> Option<&str> {
        string(self.property(LANGUAGE))
    }

    /// Returns the producer-specific document-version string.
    pub fn document_version(&self) -> Option<&str> {
        string(self.property(DOCUMENT_VERSION))
    }
}

fn string(value: Option<&Value>) -> Option<&str> {
    match value {
        Some(Value::Lpstr(value) | Value::Lpwstr(value)) => Some(value),
        _ => None,
    }
}

fn i4(value: Option<&Value>) -> Option<i32> {
    match value {
        Some(Value::I4(value)) => Some(*value),
        _ => None,
    }
}

fn boolean(value: Option<&Value>) -> Option<bool> {
    match value {
        Some(Value::Bool(value)) => Some(*value),
        _ => None,
    }
}
