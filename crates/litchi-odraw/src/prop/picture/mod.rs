//! Typed picture-name metadata from `[MS-ODRAW]` section 2.3.23.
//!
//! The public facade keeps the short contextual names [`Flags`], [`Name`],
//! [`Metadata`], and [`Snapshot`].  The implementation separates borrowed
//! semantic views, wire codecs, validation, and snapshot editing.  A snapshot
//! owns no source bytes; an edit allocates only when it commits a changed
//! property table and keeps every untouched descriptor byte-for-byte.

mod codec;
mod model;
mod validation;

pub use codec::{Commit, Edit, Owned, Patch, Snapshot, parse};
pub use model::{
    Change, Flags, Kind, MAX_NAME_BYTES, MAX_NAME_UNITS, MAX_SNAPSHOT_BYTES, Metadata, Name,
};

pub(crate) fn parse_properties<'data>(
    properties: &super::Props<'data>,
) -> crate::Result<Option<Metadata<'data>>> {
    validation::decode(properties)
}

#[cfg(test)]
mod tests;
