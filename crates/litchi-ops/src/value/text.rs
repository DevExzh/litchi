//! Text helpers shared by the coercions: .NET-compatible number white space,
//! ASCII and Unicode case-insensitive comparison, sorted-table lookup, and a
//! fixed-capacity formatting buffer.

use std::cmp::Ordering;
use std::fmt;

/// White space .NET's `NumberStyles.AllowLeadingWhite` and
/// `AllowTrailingWhite` accept around a number: U+0009 to U+000D and U+0020.
/// (`string.Trim()`, which OfficeCLI applies first, removes every Unicode
/// white-space character; that is Rust's [`str::trim`].)
#[inline]
pub(super) const fn is_number_space(byte: u8) -> bool {
    matches!(byte, b'\t'..=b'\r' | b' ')
}

/// `text` ends with the ASCII `suffix`, ignoring ASCII case.
#[inline]
pub(super) fn ends_with_ci(text: &str, suffix: &str) -> bool {
    let (text, suffix) = (text.as_bytes(), suffix.as_bytes());
    text.len() >= suffix.len() && text[text.len() - suffix.len()..].eq_ignore_ascii_case(suffix)
}

/// `text` starts with the ASCII `prefix`, ignoring ASCII case.
#[inline]
pub(super) fn starts_with_ci(text: &str, prefix: &str) -> bool {
    let (text, prefix) = (text.as_bytes(), prefix.as_bytes());
    text.len() >= prefix.len() && text[..prefix.len()].eq_ignore_ascii_case(prefix)
}

/// Removes the ASCII `suffix` from `text`, ignoring ASCII case.
///
/// The cut is always on a character boundary: the matched bytes equal an
/// ASCII suffix, so they are ASCII themselves.
#[inline]
pub(super) fn strip_suffix_ci<'a>(text: &'a str, suffix: &str) -> Option<&'a str> {
    ends_with_ci(text, suffix).then(|| &text[..text.len() - suffix.len()])
}

/// Removes the ASCII `prefix` from `text`, ignoring ASCII case.
#[inline]
pub(super) fn strip_prefix_ci<'a>(text: &'a str, prefix: &str) -> Option<&'a str> {
    starts_with_ci(text, prefix).then(|| &text[prefix.len()..])
}

/// Orders a lower-case ASCII table key against `input` folded to ASCII lower
/// case. Non-ASCII input bytes are left as they are, so they never equal a
/// table key.
#[inline]
pub(super) fn cmp_folded(key: &str, input: &str) -> Ordering {
    key.bytes()
        .cmp(input.bytes().map(|byte| byte.to_ascii_lowercase()))
}

/// Finds `input` (ASCII case-insensitively) in `table`, which is sorted by the
/// lower-case ASCII key `key` returns. Allocation-free binary search.
#[inline]
pub(super) fn find_folded<'t, T>(
    table: &'t [T],
    key: impl Fn(&T) -> &str,
    input: &str,
) -> Option<&'t T> {
    table
        .binary_search_by(|entry| cmp_folded(key(entry), input))
        .ok()
        .map(|index| &table[index])
}

/// `true` when `a` sorts strictly before `b` byte-wise (const-evaluable).
pub(super) const fn str_lt(a: &str, b: &str) -> bool {
    let (a, b) = (a.as_bytes(), b.as_bytes());
    let mut index = 0;
    while index < a.len() && index < b.len() {
        if a[index] != b[index] {
            return a[index] < b[index];
        }
        index += 1;
    }
    a.len() < b.len()
}

/// `true` when `text` is non-empty ASCII without upper-case letters
/// (const-evaluable).
pub(super) const fn is_lower_ascii(text: &str) -> bool {
    let bytes = text.as_bytes();
    if bytes.is_empty() {
        return false;
    }
    let mut index = 0;
    while index < bytes.len() {
        if !bytes[index].is_ascii() || bytes[index].is_ascii_uppercase() {
            return false;
        }
        index += 1;
    }
    true
}

/// Folds one character the way .NET's `OrdinalIgnoreCase` does: to its
/// simple upper-case form. A character whose full upper-case mapping expands
/// to several characters (`ß`, `ŉ`, the Greek iota-subscript forms) is kept
/// as it is.
#[inline]
fn fold(c: char) -> char {
    if c.is_ascii() {
        return c.to_ascii_uppercase();
    }
    let mut upper = c.to_uppercase();
    match (upper.next(), upper.next()) {
        (Some(single), None) => single,
        _ => c,
    }
}

/// Case-insensitive equality (.NET `string.Equals(a, b, OrdinalIgnoreCase)`).
pub(crate) fn eq_ignore_case(a: &str, b: &str) -> bool {
    if a.is_ascii() && b.is_ascii() {
        return a.eq_ignore_ascii_case(b);
    }
    let (mut left, mut right) = (a.chars(), b.chars());
    loop {
        match (left.next(), right.next()) {
            (None, None) => return true,
            (Some(x), Some(y)) if fold(x) == fold(y) => {},
            _ => return false,
        }
    }
}

/// Case-insensitive substring test (.NET `string.Contains(needle,
/// OrdinalIgnoreCase)`). An empty needle is contained in every text.
pub(crate) fn contains_ignore_case(haystack: &str, needle: &str) -> bool {
    if needle.is_empty() {
        return true;
    }
    if haystack.is_ascii() && needle.is_ascii() {
        let needle = needle.as_bytes();
        return haystack
            .as_bytes()
            .windows(needle.len())
            .any(|window| window.eq_ignore_ascii_case(needle));
    }
    let mut rest = haystack;
    loop {
        if starts_with_folded(rest, needle) {
            return true;
        }
        let mut chars = rest.chars();
        if chars.next().is_none() {
            return false;
        }
        rest = chars.as_str();
    }
}

fn starts_with_folded(text: &str, prefix: &str) -> bool {
    let mut chars = text.chars();
    prefix
        .chars()
        .all(|p| chars.next().is_some_and(|c| fold(c) == fold(p)))
}

/// A fixed-capacity UTF-8 buffer that implements [`fmt::Write`] without
/// allocating. A write that does not fit fails as a whole and leaves the
/// buffer unchanged, so the contents are always valid UTF-8.
pub(super) struct StackStr<const N: usize> {
    bytes: [u8; N],
    len: usize,
}

impl<const N: usize> StackStr<N> {
    /// An empty buffer.
    pub(super) const fn new() -> Self {
        Self {
            bytes: [0; N],
            len: 0,
        }
    }

    /// The text written so far.
    pub(super) fn as_str(&self) -> &str {
        // Every write copies a whole `&str`, so the prefix is valid UTF-8.
        std::str::from_utf8(&self.bytes[..self.len]).unwrap_or_default()
    }
}

impl<const N: usize> fmt::Write for StackStr<N> {
    fn write_str(&mut self, text: &str) -> fmt::Result {
        let end = self.len.checked_add(text.len()).ok_or(fmt::Error)?;
        if end > N {
            return Err(fmt::Error);
        }
        self.bytes[self.len..end].copy_from_slice(text.as_bytes());
        self.len = end;
        Ok(())
    }
}
