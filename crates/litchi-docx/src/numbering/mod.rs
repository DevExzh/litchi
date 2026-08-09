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
pub use transaction::{Commit, Patch, Snapshot, Transaction};
