//! Shared rich-text value models for the iWork format crates.
//!
//! This crate intentionally has no archive, protobuf, or application-format
//! dependencies. It owns only the allocation-bearing text values that Pages,
//! Numbers, and Keynote can exchange without importing one another's package
//! semantics.

#![forbid(unsafe_code)]

pub mod character;
pub mod columns;
pub mod font;
pub mod paragraph;

pub use character::{
    Error as CharacterError, TextBaselineShift, TextCapitalization, TextCharacterSpacing,
    TextDecorations, TextLigatures, TextPointSize, TextScript, TextStrikethrough, TextStyle,
    TextUnderline,
};
pub use font::{Font, Name, NameError};

/// A contiguous rich-text storage value.
#[derive(Debug, Clone, Default)]
pub struct TextStorage {
    /// The UTF-8 text content.
    pub text: String,
    /// Runs with styling references relative to [`Self::text`].
    pub runs: Vec<TextRun>,
    /// The source storage identifier, when one exists.
    pub identifier: Option<u64>,
}

impl TextStorage {
    /// Creates an empty text storage.
    #[must_use]
    pub const fn new() -> Self {
        Self {
            text: String::new(),
            runs: Vec::new(),
            identifier: None,
        }
    }

    /// Creates a storage containing one unstyled run for `text`.
    #[must_use]
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

    /// Borrows the plain text content without copying it.
    #[must_use]
    pub fn plain_text(&self) -> &str {
        &self.text
    }

    /// Iterates over non-empty, valid text fragments without copying text.
    #[must_use]
    pub fn iter_fragments(&self) -> TextFragmentIter<'_> {
        TextFragmentIter {
            text: &self.text,
            runs: self.runs.iter(),
        }
    }

    /// Returns whether the storage contains no text.
    #[must_use]
    pub const fn is_empty(&self) -> bool {
        self.text.is_empty()
    }

    /// Returns the UTF-8 byte length of the stored text.
    #[must_use]
    pub const fn len(&self) -> usize {
        self.text.len()
    }
}

/// A text run with a shared style reference.
#[derive(Debug, Clone)]
pub struct TextRun {
    /// UTF-8 byte offset from the start of the storage.
    pub offset: usize,
    /// UTF-8 byte length of this run.
    pub length: usize,
    /// Optional style identifier.
    pub style: Option<u64>,
}

/// A borrowed fragment of text and its associated style reference.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextFragment<'a> {
    /// Text borrowed from the source storage.
    pub text: &'a str,
    /// Optional style identifier.
    pub style: Option<u64>,
}

/// Lazy iterator over valid, non-empty runs in a [`TextStorage`].
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

/// Converts protobuf-decoded text lines into one storage value.
///
/// The line representation has already been decoded by the owning format
/// crate, so joining it cannot fail. Returning the value directly removes an
/// unnecessary error wrapper from this allocation-free model layer.
#[must_use]
pub fn parse_storage_archive(text_lines: &[String]) -> TextStorage {
    TextStorage::from_text(text_lines.join("\n"))
}

/// Joins owned storages while preserving empty storage positions.
#[must_use]
pub fn extract_text_from_storages(storages: Vec<TextStorage>) -> String {
    let capacity = storages
        .len()
        .saturating_sub(1)
        .saturating_add(storages.iter().map(|storage| storage.text.len()).sum());
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
    fn text_storage_creation() {
        let storage = TextStorage::from_text("Hello, World!".to_owned());
        assert_eq!(storage.plain_text(), "Hello, World!");
        assert_eq!(storage.len(), 13);
        assert!(!storage.is_empty());
    }

    #[test]
    fn text_fragments_are_borrowed() -> Result<(), &'static str> {
        let mut storage = TextStorage::from_text("Hello World".to_owned());
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
        let first = fragments.next().ok_or("first fragment")?;
        let second = fragments.next().ok_or("second fragment")?;

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
        Ok(())
    }

    #[test]
    fn malformed_runs_are_skipped() {
        let storage = TextStorage {
            text: "éclair".to_owned(),
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
    fn joined_text_uses_one_output_and_preserves_empty_storages() {
        let text = extract_text_from_storages(vec![
            TextStorage::from_text("First".to_owned()),
            TextStorage::new(),
            TextStorage::from_text("第三".to_owned()),
        ]);

        assert_eq!(text, "First\n\n第三");
        assert_eq!(text.capacity(), text.len());
    }

    #[test]
    fn parse_storage_archive_joins_lines() {
        let lines = vec![
            "First line".to_owned(),
            "Second line".to_owned(),
            "Third line".to_owned(),
        ];

        let storage = parse_storage_archive(&lines);
        assert!(storage.plain_text().contains("First line"));
        assert!(storage.plain_text().contains("Second line"));
        assert!(storage.plain_text().contains("Third line"));
    }
}
