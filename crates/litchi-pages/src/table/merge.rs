//! Archive-free merged-cell geometry for Pages tables.
//!
//! Pages owns the native formula storage and package traversal, while the
//! checked rectangular geometry is shared by all table-producing iWork
//! formats.  Keeping this compatibility module at the Pages table boundary
//! lets callers describe merge regions without importing a concrete wire
//! codec or native object identifier.

pub use litchi_iwa_common::table::merge::{
    AnchorRelocation, Axis, Deletion, Error, Region, Result, after_deletion, after_insertion,
    anchor_relocation_after_deletion,
};
