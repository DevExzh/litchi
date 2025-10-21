//! Text Storage Structures
//!
//! iWork documents store text in TSWP (Text Word Processing) storage objects
//! that contain rich text with styling information.

use crate::iwa::Result;

/// Represents a contiguous block of text storage
#[derive(Debug, Clone)]
pub struct TextStorage {
    /// The raw text content
    pub text: String,
    /// Text runs with styling information
    pub runs: Vec<TextRun>,
    /// Storage identifier
    pub identifier: Option<u64>,
}

impl TextStorage {
    /// Create a new empty text storage
    pub fn new() -> Self {
        Self {
            text: String::new(),
            runs: Vec::new(),
            identifier: None,
        }
    }

    /// Create text storage from a string
    pub fn from_text(text: String) -> Self {
        let length = text.len();
        Self {
            text,
            runs: vec![TextRun {
                offset: 0,
                length,
                style: None,
            }],
            identifier: None,
        }
    }

    /// Get plain text content without styling
    pub fn plain_text(&self) -> &str {
        &self.text
    }

    /// Get text fragments with their styling
    pub fn fragments(&self) -> Vec<TextFragment> {
        self.runs
            .iter()
            .filter_map(|run| {
                let end = (run.offset + run.length).min(self.text.len());
                if run.offset < self.text.len() && run.offset < end {
                    Some(TextFragment {
                        text: self.text[run.offset..end].to_string(),
                        style: run.style,
                    })
                } else {
                    None
                }
            })
            .collect()
    }

    /// Check if storage is empty
    pub fn is_empty(&self) -> bool {
        self.text.is_empty()
    }

    /// Get length of stored text
    pub fn len(&self) -> usize {
        self.text.len()
    }
}

impl Default for TextStorage {
    fn default() -> Self {
        Self::new()
    }
}

/// Represents a run of text with consistent styling
#[derive(Debug, Clone)]
pub struct TextRun {
    /// Offset from start of text storage
    pub offset: usize,
    /// Length of this run
    pub length: usize,
    /// Style identifier (reference to style object)
    pub style: Option<u64>,
}

/// A fragment of text with its associated style
#[derive(Debug, Clone)]
pub struct TextFragment {
    /// The text content
    pub text: String,
    /// Optional style reference
    pub style: Option<u64>,
}

impl TextFragment {
    /// Create a new text fragment
    pub fn new(text: String) -> Self {
        Self { text, style: None }
    }

    /// Create a fragment with style
    pub fn with_style(text: String, style: u64) -> Self {
        Self {
            text,
            style: Some(style),
        }
    }
}

/// Parse text storage from protobuf StorageArchive message
pub fn parse_storage_archive(text_lines: &[String]) -> Result<TextStorage> {
    // StorageArchive in iWork protobuf contains text as repeated string field
    // Join all text lines with newlines to preserve structure
    let text = text_lines.join("\n");

    Ok(TextStorage::from_text(text))
}

/// Extract text from multiple storage archives
pub fn extract_text_from_storages(storages: Vec<TextStorage>) -> String {
    storages
        .into_iter()
        .map(|s| s.plain_text().to_string())
        .collect::<Vec<_>>()
        .join("\n")
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_text_storage_creation() {
        let storage = TextStorage::from_text("Hello, World!".to_string());
        assert_eq!(storage.plain_text(), "Hello, World!");
        assert_eq!(storage.len(), 13);
        assert!(!storage.is_empty());
    }

    #[test]
    fn test_text_fragments() {
        let mut storage = TextStorage::from_text("Hello World".to_string());
        storage.runs = vec![
            TextRun {
                offset: 0,
                length: 5,
                style: Some(1),
            },
            TextRun {
                offset: 6,
                length: 5,
                style: Some(2),
            },
        ];

        let fragments = storage.fragments();
        assert_eq!(fragments.len(), 2);
        assert_eq!(fragments[0].text, "Hello");
        assert_eq!(fragments[1].text, "World");
    }

    #[test]
    fn test_parse_storage_archive() {
        let lines = vec![
            "First line".to_string(),
            "Second line".to_string(),
            "Third line".to_string(),
        ];

        let storage = parse_storage_archive(&lines).unwrap();
        assert!(storage.plain_text().contains("First line"));
        assert!(storage.plain_text().contains("Second line"));
        assert!(storage.plain_text().contains("Third line"));
    }
}
