//! Text Storage Structures
//!
//! iWork documents store text in TSWP (Text Word Processing) storage objects
//! that contain rich text with styling information.

use crate::Result;

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

    /// Iterate over text fragments without copying their text.
    pub fn iter_fragments(&self) -> TextFragmentIter<'_> {
        TextFragmentIter {
            text: &self.text,
            runs: self.runs.iter(),
        }
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

/// A borrowed fragment of text with its associated style.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextFragment<'a> {
    /// The text content borrowed from its [`TextStorage`].
    pub text: &'a str,
    /// Optional style reference
    pub style: Option<u64>,
}

/// Lazy iterator over the non-empty, valid runs in a [`TextStorage`].
#[derive(Debug)]
pub struct TextFragmentIter<'a> {
    text: &'a str,
    runs: std::slice::Iter<'a, TextRun>,
}

impl<'a> Iterator for TextFragmentIter<'a> {
    type Item = TextFragment<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        self.runs.find_map(|run| {
            let end = run.offset.checked_add(run.length)?.min(self.text.len());
            let text = self.text.get(run.offset..end)?;
            (!text.is_empty()).then_some(TextFragment {
                text,
                style: run.style,
            })
        })
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
    let capacity = storages
        .iter()
        .fold(storages.len().saturating_sub(1), |capacity, storage| {
            capacity.saturating_add(storage.text.len())
        });
    let mut text = String::with_capacity(capacity);

    for (index, storage) in storages.into_iter().enumerate() {
        if index != 0 {
            text.push('\n');
        }
        text.push_str(&storage.text);
    }

    text
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
    fn test_text_fragments_are_borrowed() {
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

        let mut fragments = storage.iter_fragments();
        let first = fragments.next().expect("first fragment");
        let second = fragments.next().expect("second fragment");

        assert_eq!(first.text, "Hello");
        assert_eq!(first.style, Some(1));
        assert_eq!(second.text, "World");
        assert_eq!(second.style, Some(2));
        assert!(fragments.next().is_none());
        assert!(std::ptr::eq(first.text.as_ptr(), storage.text.as_ptr()));
        assert!(std::ptr::eq(
            second.text.as_ptr(),
            storage.text.as_ptr().wrapping_add(6)
        ));
    }

    #[test]
    fn test_text_fragments_skip_malformed_runs() {
        let storage = TextStorage {
            text: "éclair".to_string(),
            runs: vec![
                TextRun {
                    offset: usize::MAX,
                    length: 1,
                    style: None,
                },
                TextRun {
                    offset: 1,
                    length: 1,
                    style: None,
                },
                TextRun {
                    offset: 2,
                    length: 3,
                    style: Some(7),
                },
            ],
            identifier: None,
        };

        let fragments: Vec<_> = storage.iter_fragments().collect();
        assert_eq!(fragments.len(), 1);
        assert_eq!(fragments[0].text, "cla");
        assert_eq!(fragments[0].style, Some(7));
    }

    #[test]
    fn test_extract_text_uses_one_output_and_preserves_empty_storages() {
        let text = extract_text_from_storages(vec![
            TextStorage::from_text("First".to_string()),
            TextStorage::new(),
            TextStorage::from_text("第三".to_string()),
        ]);

        assert_eq!(text, "First\n\n第三");
        assert_eq!(text.capacity(), text.len());
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
