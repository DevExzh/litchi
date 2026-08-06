//! Workbook-owned Office Add-in task panes.
//!
//! The XLSX facade keeps task-pane CRUD beside the package owner while the
//! shared [`litchi_ooxml_common::web`] owner supplies the MS-OWEXML model,
//! bounded XML codec, relationship planner, and opaque extension retention.
//! [`Transaction`] stages a complete OPC snapshot, so failed edits and
//! dropped transactions cannot publish partial package state.

mod package;
mod transaction;

#[cfg(test)]
mod tests;

pub use litchi_ooxml_common::web::{
    AddIn, Binding, BindingKind, Compression, Conformance, Dock, Effect, EffectKind, ExtKind,
    ExtList, Pane, Panes, Property, Reference, Selector, Snapshot, Store,
};
pub use package::{load, remove, store};
pub use transaction::Transaction;

pub(crate) use package::existing_conformance;
