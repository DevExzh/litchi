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

use super::RichTextRun;
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
    /// Optional rich text runs per string index
    rich_text: HashMap<usize, Vec<RichTextRun>>,
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
        let mut rich_text: HashMap<usize, Vec<RichTextRun>> =
            HashMap::with_capacity(INITIAL_STRINGS_CAPACITY / 4);

        let bytes = content.as_bytes();
        let mut pos = 0;

        while let Some(si_start) = memchr::memmem::find(&bytes[pos..], b"<si>") {
            let si_start_pos = pos + si_start;
            if let Some(si_end) = memchr::memmem::find(&bytes[si_start_pos..], b"</si>") {
                let si_content = &content[si_start_pos..si_start_pos + si_end + 5];

                if let Some(text) = Self::extract_text_from_si(si_content) {
                    let index = strings.len();
                    // Performance: Always push to maintain order, deduplication handled by Excel
                    strings.push(text);

                    if let Some(runs) = Self::extract_rich_text_runs_from_si(si_content)
                        && !runs.is_empty()
                    {
                        rich_text.insert(index, runs);
                    }
                }

                pos = si_start_pos + si_end + 5;
            } else {
                break;
            }
        }

        Ok(SharedStrings {
            strings,
            string_to_index,
            rich_text,
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

    /// Get rich text runs for a specific shared string index, if present.
    pub fn rich_text_runs(&self, index: usize) -> Option<&[RichTextRun]> {
        self.rich_text.get(&index).map(|v| v.as_slice())
    }

    /// Extract text content from <si> element - optimized version.
    fn extract_text_from_si(si_content: &str) -> Option<String> {
        let bytes = si_content.as_bytes();
        let mut result = String::new();
        let mut search_start = 0;

        // An <si> may contain multiple <r><t>...</t></r> runs for rich text. Concatenate
        // the text from all <t> elements in document order.
        while let Some(rel_pos) = memchr::memmem::find(&bytes[search_start..], b"<t") {
            let t_pos = search_start + rel_pos;

            // Find the closing '>' of the <t ...> start tag.
            let after_t = &bytes[t_pos..];
            let gt_rel = match memchr::memchr(b'>', after_t) {
                Some(p) => p,
                None => break,
            };
            let text_start = t_pos + gt_rel + 1;

            // Find the corresponding </t> end tag.
            let after_text = &bytes[text_start..];
            if let Some(end_rel) = memchr::memmem::find(after_text, b"</t>") {
                let text_end = text_start + end_rel;
                result.push_str(&si_content[text_start..text_end]);
                // Advance search past this </t>
                search_start = text_end + 4; // len("</t>")
            } else {
                break;
            }
        }

        if result.is_empty() {
            None
        } else {
            Some(result)
        }
    }

    /// Extract rich text runs (if any) from an <si> element.
    fn extract_rich_text_runs_from_si(si_content: &str) -> Option<Vec<RichTextRun>> {
        let bytes = si_content.as_bytes();
        let mut runs: Vec<RichTextRun> = Vec::new();
        let mut pos = 0usize;
        let slice = &bytes;

        while let Some(rel) = memchr::memmem::find(&slice[pos..], b"<r") {
            let r_pos = pos + rel;
            let next = slice.get(r_pos + 2).copied();
            if !matches!(next, Some(b'>') | Some(b' ')) {
                pos = r_pos + 2;
                continue;
            }

            let after_r = &slice[r_pos..];
            let gt_rel = match memchr::memchr(b'>', after_r) {
                Some(p) => p,
                None => break,
            };
            let inner_start = r_pos + gt_rel + 1;
            let after_inner = &slice[inner_start..];
            let end_rel = match memchr::memmem::find(after_inner, b"</r>") {
                Some(p) => p,
                None => break,
            };
            let inner_end = inner_start + end_rel;
            let run_inner = &si_content[inner_start..inner_end];

            if let Some(run) = Self::parse_rich_text_run(run_inner) {
                runs.push(run);
            }

            pos = inner_end + 4; // len("</r>")
        }

        if runs.is_empty() {
            // Fallback: treat entire shared string as a single run if we have text
            // but no explicit <r> elements.
            if let Some(text) = Self::extract_text_from_si(si_content)
                && !text.is_empty()
            {
                runs.push(RichTextRun {
                    text,
                    font_name: None,
                    font_size: None,
                    bold: false,
                    italic: false,
                    underline: false,
                    color: None,
                });
            }
        }

        if runs.is_empty() { None } else { Some(runs) }
    }

    fn parse_rich_text_run(content: &str) -> Option<RichTextRun> {
        let bytes = content.as_bytes();
        let mut text = String::new();
        let mut search_start = 0;

        while let Some(rel_pos) = memchr::memmem::find(&bytes[search_start..], b"<t") {
            let t_pos = search_start + rel_pos;
            let after_t = &bytes[t_pos..];
            let gt_rel = match memchr::memchr(b'>', after_t) {
                Some(p) => p,
                None => break,
            };
            let text_start = t_pos + gt_rel + 1;

            let after_text = &bytes[text_start..];
            if let Some(end_rel) = memchr::memmem::find(after_text, b"</t>") {
                let text_end = text_start + end_rel;
                text.push_str(&content[text_start..text_end]);
                search_start = text_end + 4;
            } else {
                break;
            }
        }

        if text.is_empty() {
            return None;
        }

        let mut font_name: Option<String> = None;
        let mut font_size: Option<f64> = None;
        let mut bold = false;
        let mut italic = false;
        let mut underline = false;
        let mut color: Option<String> = None;

        if let Some(rpr_start) = content.find("<rPr") {
            let rpr_bytes = &bytes[rpr_start..];
            if let Some(rpr_end_rel) = memchr::memmem::find(rpr_bytes, b"</rPr>") {
                let rpr_end = rpr_start + rpr_end_rel + "</rPr>".len();
                let rpr_content = &content[rpr_start..rpr_end];

                if let Some(pos) = rpr_content.find("<rFont")
                    && let Some(val_pos) = rpr_content[pos..].find("val=\"")
                {
                    let start = pos + val_pos + 5;
                    if let Some(end_rel) = rpr_content[start..].find('"') {
                        font_name = Some(rpr_content[start..start + end_rel].to_string());
                    }
                }

                if let Some(pos) = rpr_content.find("<sz")
                    && let Some(val_pos) = rpr_content[pos..].find("val=\"")
                {
                    let start = pos + val_pos + 5;
                    if let Some(end_rel) = rpr_content[start..].find('"')
                        && let Ok(sz) = rpr_content[start..start + end_rel].parse::<f64>()
                    {
                        font_size = Some(sz);
                    }
                }

                if rpr_content.contains("<b/") || rpr_content.contains("<b ") {
                    bold = true;
                }
                if rpr_content.contains("<i/") || rpr_content.contains("<i ") {
                    italic = true;
                }
                if rpr_content.contains("<u/") || rpr_content.contains("<u ") {
                    underline = true;
                }

                if let Some(pos) = rpr_content.find("<color")
                    && let Some(rgb_pos) = rpr_content[pos..].find("rgb=\"")
                {
                    let start = pos + rgb_pos + 5;
                    if let Some(end_rel) = rpr_content[start..].find('"') {
                        color = Some(rpr_content[start..start + end_rel].to_string());
                    }
                }
            }
        }

        Some(RichTextRun {
            text,
            font_name,
            font_size,
            bold,
            italic,
            underline,
            color,
        })
    }
}
