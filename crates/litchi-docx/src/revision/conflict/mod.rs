//! Inert, source-preserving `WordprocessingML` conflict markup.
//!
//! The module models only conflict annotations. It never evaluates, activates,
//! or follows embedded code, controls, actions, DDE, macros, or VBA content.

mod codec;
mod model;
mod package;
mod transaction;
mod validation;

#[cfg(test)]
mod tests;

pub use model::{
    AttributeSpan, Conflict, Id, Inventory, Kind, Limits, Metadata, Range, Scope, Snapshot, Span,
};

pub use package::Story;
pub use transaction::{Commit, Edit, Patch, Transaction};
