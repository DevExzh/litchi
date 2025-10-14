//! Shared strings table for Excel files.
//!
//! Excel uses a shared strings table to efficiently store string values.
//! This module provides parsing and access to the shared strings.
//!
//! Performance optimizations:
//! - Uses memchr for fast character searching
//! - Pre-allocates vectors with reasonable capacities
//! - Uses FxHashMap for better performance (if available)

use std::collections::HashMap;

use crate::sheet::Result;

// Performance: Pre-allocate typical capacities to reduce reallocations
const INITIAL_STRINGS_CAPACITY: usize = 10000;

/// Shared strings table for efficient string storage.
#[derive(Debug, Default)]
pub struct SharedStrings {
    /// The actual strings
    strings: Vec<String>,
    /// Reverse mapping from string to index for deduplication
    #[allow(dead_code)] // Used for deduplication logic
    string_to_index: HashMap<String, usize>,
}

impl SharedStrings {
    /// Create a new empty shared strings table.
    pub fn new() -> Self {
        Self::default()
    }

    /// Parse shared strings from xl/sharedStrings.xml content - optimized version.
    pub fn parse(content: &str) -> Result<Self> {
        let mut strings = Vec::with_capacity(INITIAL_STRINGS_CAPACITY);
        let string_to_index = HashMap::with_capacity(INITIAL_STRINGS_CAPACITY);

        let bytes = content.as_bytes();
        let mut pos = 0;

        while let Some(si_start) = memchr::memmem::find(&bytes[pos..], b"<si>") {
            let si_start_pos = pos + si_start;
            if let Some(si_end) = memchr::memmem::find(&bytes[si_start_pos..], b"</si>") {
                let si_content = &content[si_start_pos..si_start_pos + si_end + 5];

                if let Some(text) = Self::extract_text_from_si(si_content) {
                    // Performance: Always push to maintain order, deduplication handled by Excel
                    strings.push(text);
                }

                pos = si_start_pos + si_end + 5;
            } else {
                break;
            }
        }

        Ok(SharedStrings {
            strings,
            string_to_index,
        })
    }

    /// Get a string by its index.
    pub fn get(&self, index: usize) -> Option<&str> {
        self.strings.get(index).map(|s| s.as_str())
    }

    /// Get the number of strings in the table.
    pub fn len(&self) -> usize {
        self.strings.len()
    }

    /// Check if the table is empty.
    pub fn is_empty(&self) -> bool {
        self.strings.is_empty()
    }

    /// Get all strings.
    pub fn strings(&self) -> &[String] {
        &self.strings
    }

    /// Extract text content from <si> element - optimized version.
    fn extract_text_from_si(si_content: &str) -> Option<String> {
        let bytes = si_content.as_bytes();

        if let Some(t_start) = memchr::memmem::find(bytes, b"<t>") {
            let t_start_pos = t_start + 3;
            if let Some(t_end) = memchr::memmem::find(&bytes[t_start_pos..], b"</t>") {
                let text = &si_content[t_start_pos..t_start_pos + t_end];
                return Some(text.to_string());
            }
        }
        None
    }
}
