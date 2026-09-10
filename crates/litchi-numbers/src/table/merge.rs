//! Compatibility re-exports for shared merged-cell geometry.
//!
//! Merged-cell topology is format-independent and is implemented by the
//! common table model. Numbers keeps this module so existing callers can
//! migrate without changing their import paths.

pub use litchi_iwa_common::table::merge::{
    AnchorRelocation, Axis, Deletion, Error, Region, Result, after_deletion, after_insertion,
    anchor_relocation_after_deletion,
};
