//! Lexical classification and quoting shared by the path and query grammars.
//!
//! `Display` chooses the bare form of a name, key, or value only when the
//! parser would read the bare text back unchanged; otherwise it emits a quoted
//! form whose escapes the parser reverses exactly. That is the invariant behind
//! the parse -> display -> parse identity.

use std::borrow::Cow;
use std::fmt::{self, Write as _};

/// ASCII whitespace accepted between tokens.
#[inline]
pub(super) const fn is_ws(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r' | 0x0c)
}

/// A character allowed in an unquoted predicate key.
#[inline]
pub(super) fn is_key_char(c: char) -> bool {
    c.is_alphanumeric() || c == '.' || c == '_'
}

/// An ASCII byte allowed in a selector element type or pseudo-class name.
#[inline]
pub(super) const fn is_ident_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || byte == b'_' || byte == b'-'
}

#[inline]
pub(super) const fn is_quote(byte: u8) -> bool {
    byte == b'"' || byte == b'\''
}

/// Reverses the backslash escapes of a quoted token: `\q` (the delimiting
/// quote) and `\\` lose their backslash; any other backslash is literal.
/// Borrows when the inner text has no backslash.
pub(super) fn unescape(inner: &str, quote: char) -> Cow<'_, str> {
    if !inner.contains('\\') {
        return Cow::Borrowed(inner);
    }
    let mut out = String::with_capacity(inner.len());
    let mut chars = inner.chars();
    while let Some(c) = chars.next() {
        if c == '\\' {
            match chars.clone().next() {
                Some(next) if next == quote || next == '\\' => {
                    out.push(next);
                    chars.next();
                },
                _ => out.push('\\'),
            }
        } else {
            out.push(c);
        }
    }
    Cow::Owned(out)
}

/// Writes `text` double-quoted, escaping `"` and `\`.
pub(super) fn write_quoted(f: &mut fmt::Formatter<'_>, text: &str) -> fmt::Result {
    f.write_char('"')?;
    for c in text.chars() {
        if c == '"' || c == '\\' {
            f.write_char('\\')?;
        }
        f.write_char(c)?;
    }
    f.write_char('"')
}

/// A path segment name that reads back unquoted.
pub(super) fn is_plain_name(text: &str) -> bool {
    let bytes = text.as_bytes();
    match bytes.first() {
        None => false,
        Some(&first) if is_quote(first) => false,
        Some(_) => !bytes
            .iter()
            .any(|&b| matches!(b, b'/' | b'[' | b']') || b.is_ascii_control()),
    }
}

/// A predicate key that reads back unquoted.
pub(super) fn is_plain_key(text: &str) -> bool {
    !text.is_empty()
        && text.chars().all(is_key_char)
        // An all-digit bare token is a positional index, not a key.
        && !text.bytes().all(|b| b.is_ascii_digit())
}

/// A comparison value that reads back unquoted.
pub(super) fn is_plain_value(text: &str) -> bool {
    let bytes = text.as_bytes();
    match bytes.first() {
        None => false,
        Some(b'=' | b'~' | b'!') => false,
        Some(_) => !bytes.iter().any(|&b| {
            matches!(b, b'[' | b']' | b'(' | b')' | b'*' | b'\\')
                || is_quote(b)
                || is_ws(b)
                || b.is_ascii_control()
        }),
    }
}

/// A `:contains()` argument that reads back unquoted.
pub(super) fn is_plain_contains(text: &str) -> bool {
    let bytes = text.as_bytes();
    !bytes.is_empty()
        && !is_ws(bytes[0])
        && !is_ws(bytes[bytes.len() - 1])
        && !bytes.iter().any(|&b| {
            matches!(b, b'[' | b']' | b'(' | b')' | b'\\') || is_quote(b) || b.is_ascii_control()
        })
}

/// An Excel sheet name that reads back without `'...'` quoting.
pub(super) fn is_plain_sheet(text: &str) -> bool {
    let mut chars = text.chars();
    match chars.next() {
        None => false,
        Some(first) if first.is_ascii_digit() => false,
        Some(first) => {
            let ok = |c: char| c.is_alphanumeric() || c == '_' || c == '.';
            ok(first) && chars.all(ok)
        },
    }
}

/// Writes an Excel-quoted sheet name: `'...'` with `'` doubled.
pub(super) fn write_sheet_quoted(f: &mut fmt::Formatter<'_>, text: &str) -> fmt::Result {
    f.write_char('\'')?;
    for c in text.chars() {
        if c == '\'' {
            f.write_char('\'')?;
        }
        f.write_char(c)?;
    }
    f.write_char('\'')
}

/// Returns `true` when `text` begins with `keyword` (ASCII case-insensitive).
#[inline]
pub(super) fn starts_with_ci(text: &[u8], keyword: &[u8]) -> bool {
    text.len() >= keyword.len() && text[..keyword.len()].eq_ignore_ascii_case(keyword)
}
