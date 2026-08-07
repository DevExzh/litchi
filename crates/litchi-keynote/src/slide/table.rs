//! Archive-free semantic values for a Keynote slide table.
//!
//! The native IWA adapter owns object discovery, protobuf decoding, and
//! package mutation. These focused modules expose only the values callers use
//! to describe table formulas and sorting.

/// Formula values shared through the neutral iWork semantic model.
pub mod formula;
/// Sort values shared through the neutral iWork semantic model.
pub mod sort;
