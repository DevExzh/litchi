//! OPC graph facade for persisted web-extension parts.

mod discovery;
mod durable;
mod graph;
mod naming;
mod planning;
mod transaction;
mod validation;

pub use discovery::*;
pub use durable::{GraphMode, OwnerSelector, WebIntent, apply_durable, plan_durable};
pub use planning::*;
pub use transaction::{Commit, Diagnostics, Patch};

pub(in crate::web) use discovery::load_with_index_budget;
pub(in crate::web) use graph::*;
pub(in crate::web) use naming::*;
#[allow(
    unused_imports,
    reason = "the package facade preserves planning helpers for sibling owners"
)]
pub(in crate::web) use planning::*;
#[allow(
    unused_imports,
    reason = "the package facade preserves transaction helpers for sibling owners"
)]
pub(in crate::web) use transaction::*;
pub(in crate::web) use validation::*;
