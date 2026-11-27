//! Typed ODF `text:list-style` declarations (numbered, bullet, and image levels).
//!
//! The owner separates semantic values and validation from XML codecs and the
//! document/package facade while keeping the concise `list_style` API intact.

mod codec;
mod model;
mod package;

use litchi_core::{Error, Result};

pub(super) const MAX_XML: usize = 64 * 1024 * 1024;
pub(super) const MAX_DEPTH: usize = 256;
pub(super) const MAX_STYLES: usize = 65_536;
pub(super) const MAX_VALUE: usize = 4_096;
pub(super) const MAX_TOTAL: usize = 16 * 1024 * 1024;
pub(super) const MAX_BINARY: usize = 8 * 1024 * 1024;
/// Maximum `text:level` of a list level (ODF 1.2 allows deep lists).
pub const MAX_LEVEL: u16 = 1_024;

pub(super) fn bad(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

pub(super) fn name_ok(value: &str, field: &str) -> Result<()> {
    if value.is_empty() || value.len() > MAX_VALUE || value.chars().any(char::is_control) {
        return Err(bad(format!("invalid {field}")));
    }
    Ok(())
}

pub(super) fn parse_bool(value: &str, field: &str) -> Result<bool> {
    match value {
        "true" | "1" => Ok(true),
        "false" | "0" => Ok(false),
        _ => Err(bad(format!("{field} must be an XML Schema boolean"))),
    }
}

pub(super) fn percent(value: &str) -> bool {
    if value.len() > MAX_VALUE {
        return false;
    }
    let Some(number) = value.strip_suffix('%') else {
        return false;
    };
    let mut split = number.split('.');
    let whole = split.next().unwrap_or_default();
    let fraction = split.next();
    if split.next().is_some() {
        return false;
    }
    let digits = |part: &str| part.bytes().all(|byte| byte.is_ascii_digit());
    match fraction {
        None => !whole.is_empty() && digits(whole),
        Some(fraction) => {
            digits(whole) && digits(fraction) && (!whole.is_empty() || !fraction.is_empty())
        },
    }
}

pub use codec::parse;
pub use model::{
    BulletRelativeSize, BulletStyle, ImageSource, Kind, LevelStyle, NumberStyle, Style, Styles,
};

#[cfg(test)]
mod tests;
