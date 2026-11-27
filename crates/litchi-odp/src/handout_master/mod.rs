//! Typed ODF handout-master ownership and its package facade.
//!
//! A handout master is a singleton under `office:master-styles`.  Its
//! drawing children use the shared ODF master-child vocabulary, while the
//! handout-specific root and package transaction remain owned by ODP.

pub(crate) mod codec;
mod model;
pub(crate) mod package;
mod validation;

#[cfg(test)]
mod tests;

pub use litchi_odf_common::style::master::{Child, ChildKind};
pub use model::{Master, Resolved};
