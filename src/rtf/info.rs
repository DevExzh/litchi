//! RTF document information and properties.
//!
//! This module provides support for document metadata like title, author,
//! subject, keywords, and other document properties.

use std::borrow::Cow;

/// Document information/metadata
#[derive(Debug, Clone, Default)]
pub struct DocumentInfo<'a> {
    /// Document title
    pub title: Option<Cow<'a, str>>,
    /// Document subject
    pub subject: Option<Cow<'a, str>>,
    /// Document author
    pub author: Option<Cow<'a, str>>,
    /// Document manager
    pub manager: Option<Cow<'a, str>>,
    /// Company name
    pub company: Option<Cow<'a, str>>,
    /// Operator (last person to modify)
    pub operator: Option<Cow<'a, str>>,
    /// Document category
    pub category: Option<Cow<'a, str>>,
    /// Keywords
    pub keywords: Option<Cow<'a, str>>,
    /// Comments
    pub comment: Option<Cow<'a, str>>,
    /// Document version
    pub version: Option<i32>,
    /// Document revision number
    pub revision: Option<i32>,
    /// Creation time (RTF datetime format)
    pub creation_time: Option<Cow<'a, str>>,
    /// Revision time (last modified)
    pub revision_time: Option<Cow<'a, str>>,
    /// Print time (last printed)
    pub print_time: Option<Cow<'a, str>>,
    /// Backup time
    pub backup_time: Option<Cow<'a, str>>,
    /// Total editing time (in minutes)
    pub editing_time: Option<i32>,
    /// Number of pages
    pub pages: Option<i32>,
    /// Number of words
    pub words: Option<i32>,
    /// Number of characters
    pub characters: Option<i32>,
    /// Number of characters including spaces
    pub characters_with_spaces: Option<i32>,
    /// Document ID (internal identifier)
    pub id: Option<i32>,
}

impl<'a> DocumentInfo<'a> {
    /// Create a new document info
    #[inline]
    pub fn new() -> Self {
        Self::default()
    }

    /// Set the title
    #[inline]
    pub fn with_title(mut self, title: Cow<'a, str>) -> Self {
        self.title = Some(title);
        self
    }

    /// Set the author
    #[inline]
    pub fn with_author(mut self, author: Cow<'a, str>) -> Self {
        self.author = Some(author);
        self
    }

    /// Set the subject
    #[inline]
    pub fn with_subject(mut self, subject: Cow<'a, str>) -> Self {
        self.subject = Some(subject);
        self
    }

    /// Set keywords
    #[inline]
    pub fn with_keywords(mut self, keywords: Cow<'a, str>) -> Self {
        self.keywords = Some(keywords);
        self
    }

    /// Set comments
    #[inline]
    pub fn with_comment(mut self, comment: Cow<'a, str>) -> Self {
        self.comment = Some(comment);
        self
    }
}

/// Protection type for document
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum ProtectionType {
    /// No protection
    #[default]
    None,
    /// Read-only
    ReadOnly,
    /// Revision tracking only
    RevisionTracking,
    /// Comments only
    Comments,
    /// Forms only
    Forms,
}

/// Document protection settings
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub struct DocumentProtection {
    /// Protection type
    pub protection_type: ProtectionType,
    /// Whether protection is enforced
    pub enforced: bool,
}

impl DocumentProtection {
    /// Create a new document protection
    #[inline]
    pub fn new(protection_type: ProtectionType) -> Self {
        Self {
            protection_type,
            enforced: true,
        }
    }

    /// Check if document is protected
    #[inline]
    pub fn is_protected(&self) -> bool {
        self.enforced && self.protection_type != ProtectionType::None
    }
}
