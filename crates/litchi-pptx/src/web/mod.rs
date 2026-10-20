//! Inert Office Add-in task-pane values for `PresentationML` packages.
//!
//! [`crate::Package`] owns package discovery; [`crate::Package::edit_task_panes`]
//! publishes changes through the [`Edit`](crate::web::Edit) family. The shared OOXML
//! implementation supplies bounded MS-OWEXML parsing, compact XML output and
//! opaque extension retention. This module never resolves, downloads, or
//! executes add-in content.

mod edit;
pub use edit::{Commit, Diagnostics, Edit, Error, LimitKind, Patch};

pub use litchi_ooxml_common::web::{
    AddIn, Binding, BindingKind, Compression, Conformance, Dock, Effect, EffectKind, ExtKind,
    ExtList, Image, Limits, Link, Pane, Panes, Property, Reference, Selector, Snapshot, Store,
};
