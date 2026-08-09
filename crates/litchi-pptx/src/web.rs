//! Inert Office Add-in task-pane values for `PresentationML` packages.
//!
//! [`crate::Package`] owns package discovery and publication. The shared
//! OOXML implementation supplies bounded MS-OWEXML parsing, compact XML
//! output, source-checked reversible patches, and opaque extension retention.
//! This module never resolves, downloads, or executes add-in content.

pub use litchi_ooxml_common::web::{
    AddIn, Binding, BindingKind, Compression, Conformance, Dock, Effect, EffectKind, ExtKind,
    ExtList, Image, Limits, Link, Pane, Panes, Patch, Property, Reference, Selector, Snapshot,
    Store,
};
