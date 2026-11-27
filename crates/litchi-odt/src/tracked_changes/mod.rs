//! Semantic tracked-change declarations and lossless ODT marker edits.
//!
//! The owner keeps declaration data and stable story positions in `model`,
//! XML grammar and bounded scanning in `codec`, document mutations in
//! `package`, and regression coverage in `tests`.

mod codec;
mod model;
mod package;
#[cfg(test)]
mod tests;

use litchi_core::{Error, Result};

pub(super) const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
pub(super) const MAX_CHANGES: usize = 1_000_000;
pub(super) const MAX_VALUE_BYTES: usize = 65_536;
pub(super) const MAX_AGGREGATE_BYTES: usize = 16 * 1_048_576;

pub(super) fn invalid<T>(message: impl Into<String>) -> Result<T> {
    Err(make_error(message))
}

pub(super) fn make_error(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

pub use model::{Position, Story};

pub use package::{
    mark_tracked_change_range_xml, mark_tracked_deletion_xml, set_tracked_changes_xml,
    unmark_tracked_change_xml,
};
