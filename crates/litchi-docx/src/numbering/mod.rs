#![expect(
    clippy::module_name_repetitions,
    reason = "public names retain established OOXML facade terminology"
)]
//! Layered `WordprocessingML` numbering model and XML codec.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use codec::parse_numbering;
pub use model::{
    Collection, Definition, Format, Instance, Level, MultiLevel, Override, Paragraph,
    ParseFormatError, ParseMultiLevelError, PictureBullet, Restart, Suffix,
};
pub(crate) use package::parse_part;
pub(crate) use package::parse_snapshot_part;

/// The deepest list level Word models (`w:ilvl` 0–8).
pub(crate) const MAX_LIST_LEVEL: u8 = 8;

/// Parse a paragraph or style `w:ilvl` value.
///
/// `w:ilvl` is an unbounded `ST_DecimalNumber`, but Word models only levels
/// 0–8. A non-negative integer above 8 (including one too large for any
/// integer type) is clamped to 8 and reported as clamped (`true`);
/// LENIENCY_AUDIT 2.9. Anything that is not a non-negative decimal integer
/// returns `None`.
pub(crate) fn parse_list_level(raw: &str) -> Option<(u8, bool)> {
    let digits = raw.strip_prefix('+').unwrap_or(raw);
    if digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit()) {
        return None;
    }
    match digits.parse::<u8>() {
        Ok(level) if level <= MAX_LIST_LEVEL => Some((level, false)),
        _ => Some((MAX_LIST_LEVEL, true)),
    }
}
pub use transaction::{Commit, Patch, Snapshot, Transaction};
