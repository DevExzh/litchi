//! Inert, source-preserving `WordprocessingML` conflict markup.
//!
//! The module models only conflict annotations. It never evaluates, activates,
//! or follows embedded code, controls, actions, DDE, macros, or VBA content.

mod codec;
mod edit;
mod limit;
mod model;
mod package;
#[allow(
    dead_code,
    unreachable_pub,
    reason = "the story-level draft stages the family edit; its accessors serve unit tests"
)]
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{
    AttributeSpan, Conflict, Id, Inventory, Kind, Limits, Metadata, Range, Scope, Snapshot, Span,
};

pub use edit::{Commit, Diagnostics, Edit, EditOptions, Error, LimitKind, Patch};
pub use package::Story;
pub use transaction::Operation;
