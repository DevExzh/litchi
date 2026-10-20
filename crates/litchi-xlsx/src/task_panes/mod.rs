//! Workbook-owned Office Add-in task panes.
//!
//! The XLSX facade keeps task-pane CRUD beside the package owner while the
//! shared [`litchi_ooxml_common::web`] owner supplies the MS-OWEXML model,
//! bounded XML codec, relationship planner, and opaque extension retention.
//! [`Edit`] stages a private candidate graph, so failed and dropped edits
//! cannot publish partial package state.

mod edit;
mod package;

#[cfg(test)]
mod tests;

pub use edit::{Commit, Diagnostics, Edit, EditOptions, Error, LimitKind, Patch};
pub use litchi_ooxml_common::web::{
    AddIn, Binding, BindingKind, Compression, Conformance, Dock, Effect, EffectKind, ExtKind,
    ExtList, Pane, Panes, Property, Reference, Selector, Snapshot, Store,
};
pub use package::load;

pub(crate) use package::existing_conformance;
