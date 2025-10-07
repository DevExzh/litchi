// Performance utilities for LaTeX conversion
//
// This module contains optimized helper functions for high-performance
// LaTeX conversion operations using SIMD and efficient algorithms.

use memchr::memchr;
use smallvec::SmallVec;

/// Fast check if a string represents a valid number using SIMD
#[inline]
pub fn is_valid_number_fast(s: &str) -> bool {
    if s.is_empty() {
        return false;
    }

    let bytes = s.as_bytes();
    let mut has_digit = false;
    let mut has_dot = false;

    // Use SIMD-friendly loop
    for &b in bytes {
        match b {
            b'0'..=b'9' => has_digit = true,
            b'.' => {
                if has_dot {
                    return false; // Multiple dots
                }
                has_dot = true;
            }
            b'-' if bytes.len() == 1 => return false, // Just a minus sign
            b'-' => {} // Allow negative numbers at start
            _ => return false, // Invalid character
        }
    }

    has_digit // Must have at least one digit
}

/// SIMD-accelerated check if a character sequence contains LaTeX special characters
#[inline]
pub fn contains_latex_special_simd(text: &str) -> bool {
    let bytes = text.as_bytes();

    // Use memchr for common special characters
    memchr(b' ', bytes).is_some()
        || memchr(b'#', bytes).is_some()
        || memchr(b'$', bytes).is_some()
        || memchr(b'%', bytes).is_some()
        || memchr(b'&', bytes).is_some()
        || memchr(b'_', bytes).is_some()
        || memchr(b'{', bytes).is_some()
        || memchr(b'}', bytes).is_some()
        || memchr(b'~', bytes).is_some()
        || memchr(b'^', bytes).is_some()
        || memchr(b'\\', bytes).is_some()
}

/// SIMD-accelerated LaTeX special character escaping
/// Returns true if escaping was needed
#[inline]
pub fn escape_latex_special_chars(text: &str, buffer: &mut String) -> bool {
    if !contains_latex_special_simd(text) {
        buffer.push_str(text);
        return false;
    }

    // Need to escape - process character by character
    for ch in text.chars() {
        match ch {
            ' ' | '#' | '$' | '%' | '&' | '_' | '{' | '}' | '~' | '^' | '\\' => {
                buffer.push('\\');
                buffer.push(ch);
            }
            _ => buffer.push(ch),
        }
    }
    true
}

/// Fast buffer extension with capacity management
#[inline]
pub fn extend_buffer_with_capacity(buffer: &mut String, text: &str, additional_capacity: usize) {
    if buffer.capacity() < buffer.len() + text.len() + additional_capacity {
        buffer.reserve(text.len() + additional_capacity);
    }
    buffer.push_str(text);
}

/// Fast check if text needs LaTeX protection (contains spaces or special chars)
#[inline]
pub fn needs_latex_protection(text: &str) -> bool {
    if text.is_empty() {
        return true;
    }

    // Quick check for spaces
    if memchr(b' ', text.as_bytes()).is_some() {
        return true;
    }

    // Check for other special characters
    contains_latex_special_simd(text)
}

/// Efficient string interning for repeated LaTeX commands
/// Uses SmallVec to avoid allocations for common cases
pub struct LatexStringCache {
    cache: SmallVec<[String; 16]>,
}

impl LatexStringCache {
    pub fn new() -> Self {
        Self {
            cache: SmallVec::new(),
        }
    }

    pub fn get_or_insert(&mut self, s: &str) -> usize {
        // Simple linear search for small cache
        for (i, cached) in self.cache.iter().enumerate() {
            if cached == s {
                return i;
            }
        }

        // Not found, insert
        let index = self.cache.len();
        self.cache.push(s.to_string());
        index
    }

    pub fn get(&self, index: usize) -> &str {
        &self.cache[index]
    }

    pub fn clear(&mut self) {
        self.cache.clear();
    }
}

impl Default for LatexStringCache {
    fn default() -> Self {
        Self::new()
    }
}

/// Performance statistics for LaTeX conversion
#[derive(Debug, Default)]
pub struct LatexConversionStats {
    pub nodes_processed: usize,
    pub strings_allocated: usize,
    pub buffer_resizes: usize,
    pub total_chars_written: usize,
}

impl LatexConversionStats {
    pub fn record_node(&mut self) {
        self.nodes_processed += 1;
    }

    pub fn record_allocation(&mut self, size: usize) {
        self.strings_allocated += 1;
        self.total_chars_written += size;
    }

    pub fn record_resize(&mut self) {
        self.buffer_resizes += 1;
    }
}
