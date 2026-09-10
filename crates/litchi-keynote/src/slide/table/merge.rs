//! Archive-free merged-cell geometry for Keynote tables.
//!
//! The package adapter owns the native merge-owner formula storage and
//! selector resolution.  This module re-exports the checked neutral geometry
//! so callers can share merge values across Numbers, Pages, and Keynote.

pub use litchi_iwa_common::table::merge::{
    AnchorRelocation, Axis, Deletion, Error, Region, Result, after_deletion, after_insertion,
    anchor_relocation_after_deletion,
};
