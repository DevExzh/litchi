//! Unified metadata representation for Word documents.
//!
//! This module provides a unified interface for document metadata
//! that works with both OLE (.doc) and OOXML (.docx) formats.
//!
//! Format-specific conversions (e.g. `From<crate::ole::OleMetadata>`) and
//! YAML front-matter serialization (which depends on `serde_saphyr`, a
//! workspace dep not present in `litchi-core`) live in the umbrella crate
//! at `src/metadata_ext.rs`.
use chrono::{DateTime, Utc};
use serde::{Deserialize, Serialize};

/// Unified document metadata structure.
///
/// Contains standard document properties that can be extracted from
/// both OLE and OOXML document formats.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Metadata {
    /// Document title
    #[serde(skip_serializing_if = "Option::is_none")]
    pub title: Option<String>,
    /// Document subject
    #[serde(skip_serializing_if = "Option::is_none")]
    pub subject: Option<String>,
    /// Document author/creator
    #[serde(skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    /// Keywords associated with the document
    #[serde(skip_serializing_if = "Option::is_none")]
    pub keywords: Option<String>,
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
    /// Revision number
    #[serde(skip_serializing_if = "Option::is_none")]
    pub revision: Option<String>,
    /// Creation date (Unix timestamp)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub created: Option<DateTime<Utc>>,
    /// Last modification date (Unix timestamp)
    #[serde(skip_serializing_if = "Option::is_none")]
    pub modified: Option<DateTime<Utc>>,
    /// Number of pages
    #[serde(skip_serializing_if = "Option::is_none")]
    pub page_count: Option<u32>,
    /// Number of words
    #[serde(skip_serializing_if = "Option::is_none")]
    pub word_count: Option<u32>,
    /// Number of characters
    #[serde(skip_serializing_if = "Option::is_none")]
    pub character_count: Option<u32>,
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
    /// Document format or content version
    #[serde(skip_serializing_if = "Option::is_none")]
    pub version: Option<String>,
    /// Last printed time
    #[serde(skip_serializing_if = "Option::is_none")]
    pub last_printed_time: Option<DateTime<Utc>>,
    /// Security level
    #[serde(skip_serializing_if = "Option::is_none")]
    pub security: Option<u32>,
    /// Codepage for text encoding
    #[serde(skip_serializing_if = "Option::is_none")]
    pub codepage: Option<u32>,
}

impl Metadata {
    /// Check if the metadata contains any actual data.
    ///
    /// Returns true if at least one field is populated.
    pub fn has_data(&self) -> bool {
        self.title.is_some()
            || self.subject.is_some()
            || self.author.is_some()
            || self.keywords.is_some()
            || self.description.is_some()
            || self.identifier.is_some()
            || self.language.is_some()
            || self.template.is_some()
            || self.last_modified_by.is_some()
            || self.revision.is_some()
            || self.created.is_some()
            || self.modified.is_some()
            || self.page_count.is_some()
            || self.word_count.is_some()
            || self.character_count.is_some()
            || self.application.is_some()
            || self.category.is_some()
            || self.company.is_some()
            || self.manager.is_some()
            || self.content_status.is_some()
            || self.content_type.is_some()
            || self.version.is_some()
            || self.last_printed_time.is_some()
            || self.security.is_some()
            || self.codepage.is_some()
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
    }
}
