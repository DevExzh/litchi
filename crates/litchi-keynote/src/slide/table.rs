//! Archive-free semantic values for a Keynote slide table.
//!
//! The native IWA adapter owns object discovery, protobuf decoding, and
//! package mutation. These focused modules expose only the values callers use
//! to describe table formulas and sorting.

/// Formula values shared with the Numbers semantic model.
pub mod formula;
/// Sort values shared with the Numbers semantic model.
pub mod sort;
