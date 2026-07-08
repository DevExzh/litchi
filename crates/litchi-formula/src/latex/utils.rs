// Performance utilities for LaTeX conversion
//
// This module contains optimized helper functions for high-performance
// LaTeX conversion operations using SIMD and efficient algorithms.

use memchr::memchr;

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
            },
            b'-' if bytes.len() == 1 => return false, // Just a minus sign
            b'-' => {},                               // Allow negative numbers at start
            _ => return false,                        // Invalid character
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
            },
            _ => buffer.push(ch),
        }
    }
    true
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
