//! Bounded, exact-source `CommonMark` and GitHub Flavored Markdown snapshots.
//!
//! Parsing follows `CommonMark` 0.31.2. [`Dialect::GitHubFlavored`] additionally
//! enables GFM tables, task lists, strikethrough, extended autolinks, tag
//! filtering, alerts, and compatible footnote definitions. A snapshot retains
//! its complete UTF-8 source byte-for-byte; parsing never renders or normalizes
//! it. Edits target exact top-level, nested-block, or inline ranges and fully
//! reparse before publication.

mod gfm;
mod model;
mod parse;
mod transaction;

pub use model::{
    Block, BlockKind, Blocks, Dialect, Error, Inline, InlineKind, Inlines, NestedBlock,
    NestedBlocks, ProjectionCapabilities, ProjectionIssue, ProjectionIssueKind,
    ProjectionPreflight, ReadLimits, Reference, ReferenceKind, References, Snapshot,
};
pub use transaction::{
    Commit, Conflict, ConflictSet, Diagnostics, Edit, History, HistoryLimits, JoinError, MergePlan,
    Patch, PatchEnvelopeLimits, TransferPlan,
};
