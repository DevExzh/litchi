//! Bounded, exact-source `CommonMark` and GitHub Flavored Markdown snapshots.
//!
//! Parsing follows `CommonMark` 0.31.2. [`Dialect::GitHubFlavored`] additionally
//! enables GFM tables, task lists, strikethrough, alerts, and compatible
//! footnote definitions. A snapshot retains its complete UTF-8 source byte-for-byte;
//! parsing never renders or normalizes it. Edits replace exactly one parsed
//! top-level block or append one block, then fully reparse before publication.

mod model;
mod parse;
mod transaction;

pub use model::{Block, BlockKind, Blocks, Dialect, Error, ReadLimits, Snapshot};
pub use transaction::{Commit, Diagnostics, Edit, Patch};
