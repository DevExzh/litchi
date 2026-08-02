//! Typed Excel Binary Workbook documents.
//!
//! The initial extraction owns the validated BIFF12 wire substrate. Semantic
//! workbook snapshots and edits will be layered over it without exposing
//! package identifiers in their ordinary APIs.

#![forbid(unsafe_code)]

pub mod raw;

pub use raw::Error;
