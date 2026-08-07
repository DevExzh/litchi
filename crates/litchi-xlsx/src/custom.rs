//! Inert typed custom document properties in the XLSX package context.
//!
//! The package-level OOXML owner remains in `litchi-ooxml-common`; this module
//! provides its types at the contextual XLSX API path without a second model
//! or codec.

pub use litchi_ooxml_common::custom::{Host, Props, Value};
