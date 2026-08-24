//! Lossless header, footer, freeze, and print-repetition settings for Pages
//! body tables.
//!
//! The semantic value is shared with the other iWork owners. Pages owns the
//! selector, native graph, strict wire, and exact-source transaction around
//! that value; this module deliberately exposes no archive identifiers.

pub use litchi_iwa_common::table::headers::{Count, Error, Settings};
