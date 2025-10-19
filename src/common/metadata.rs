/// Unified metadata representation for Word documents.
///
/// This module provides a unified interface for document metadata
/// that works with both OLE (.doc) and OOXML (.docx) formats.
use serde::{Deserialize, Serialize};
use crate::common::Result;
use chrono::{DateTime, Utc};

/// Unified document metadata structure.
///
/// Contains standard document properties that can be extracted from
/// both OLE and OOXML document formats.
#[derive(Debug, Clone, Default, Serialize, Deserialize)]
pub struct Metadata {
    /// Document title
    pub title: Option<String>,
    /// Document subject
    pub subject: Option<String>,
    /// Document author/creator
    pub author: Option<String>,
    /// Keywords associated with the document
    pub keywords: Option<String>,
    /// Document description/comments
    pub description: Option<String>,
    /// Template used to create the document
    pub template: Option<String>,
    /// Last person to modify the document
    pub last_modified_by: Option<String>,
    /// Revision number
    pub revision: Option<String>,
    /// Creation date (Unix timestamp)
    pub created: Option<DateTime<Utc>>,
    /// Last modification date (Unix timestamp)
    pub modified: Option<DateTime<Utc>>,
    /// Number of pages
    pub page_count: Option<u32>,
    /// Number of words
    pub word_count: Option<u32>,
    /// Number of characters
    pub character_count: Option<u32>,
    /// Application that created the document
    pub application: Option<String>,
    /// Document category
    pub category: Option<String>,
    /// Company/organization
    pub company: Option<String>,
    /// Manager name
    pub manager: Option<String>,
    /// Content status (draft, final, etc.)
    pub content_status: Option<String>,
    /// Last printed time
    pub last_printed_time: Option<DateTime<Utc>>,
    /// Security level
    pub security: Option<u32>,
    /// Codepage for text encoding
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
            || self.security.is_some()
            || self.codepage.is_some()
    }

    /// Convert metadata to YAML front matter format.
    ///
    /// Returns a string containing the YAML front matter block,
    /// or an empty string if no metadata is available.
    pub fn to_yaml_front_matter(&self) -> Result<String> {
        if !self.has_data() {
            return Ok(String::new());
        }

        let yaml_string = serde_saphyr::to_string(self)
            .map_err(|e| crate::common::Error::Other(format!("Failed to serialize metadata to YAML: {}", e)))?;

        // Add YAML front matter delimiters
        Ok(format!("---\n{}---\n\n", yaml_string))
    }
}

#[cfg(feature = "ole")]
impl From<crate::ole::OleMetadata> for Metadata {
    fn from(ole_metadata: crate::ole::OleMetadata) -> Self {
        Self {
            title: ole_metadata.title,
            subject: ole_metadata.subject,
            author: ole_metadata.author,
            keywords: ole_metadata.keywords,
            description: ole_metadata.comments,
            template: ole_metadata.template,
            last_modified_by: ole_metadata.last_saved_by,
            revision: ole_metadata.revision_number,
            created: ole_metadata.create_time,
            modified: ole_metadata.last_saved_time,
            page_count: ole_metadata.num_pages,
            word_count: ole_metadata.num_words,
            character_count: ole_metadata.num_chars,
            application: ole_metadata.creating_application,
            category: ole_metadata.category,
            company: ole_metadata.company,
            manager: ole_metadata.manager,
            content_status: None, // OLE doesn't have this field
            last_printed_time: ole_metadata.last_printed_time,
            security: ole_metadata.security,
            codepage: ole_metadata.codepage,
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
    }

    #[test]
    fn test_metadata_to_yaml_front_matter() {
        let empty_metadata = Metadata::default();
        let yaml = empty_metadata.to_yaml_front_matter().unwrap();
        assert_eq!(yaml, "");

        let metadata = Metadata {
            title: Some("Test Document".to_string()),
            author: Some("Test Author".to_string()),
            subject: Some("Test Subject".to_string()),
            ..Default::default()
        };

        let yaml = metadata.to_yaml_front_matter().unwrap();
        assert!(yaml.starts_with("---\n"));
        assert!(yaml.ends_with("---\n\n"));
        assert!(yaml.contains("title: Test Document"));
        assert!(yaml.contains("author: Test Author"));
        assert!(yaml.contains("subject: Test Subject"));
    }

    #[test]
    #[cfg(feature = "ole")]
    fn test_ole_metadata_conversion() {
        let ole_metadata = crate::ole::OleMetadata {
            title: Some("OLE Document".to_string()),
            author: Some("OLE Author".to_string()),
            codepage: Some(65001),
            ..Default::default()
        };

        let metadata: Metadata = ole_metadata.into();
        assert_eq!(metadata.title, Some("OLE Document".to_string()));
        assert_eq!(metadata.author, Some("OLE Author".to_string()));
        assert_eq!(metadata.codepage, Some(65001));
    }
}
