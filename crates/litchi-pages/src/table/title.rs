//! Lossless title visibility and outline settings for Pages body tables.
//!
//! The compact semantic value preserves native optional-field presence:
//! `None` is absent, while `Some(false)` is an explicit native false.

pub use crate::package::body_table_title::{
    Commit, Diagnostics, Edit, Error, LimitKind, Patch, Refusal,
};
pub use litchi_iwa_common::table::title::Settings;
