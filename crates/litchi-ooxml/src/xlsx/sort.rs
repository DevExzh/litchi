//! Compatibility exports for the canonical XLSX sort codec.
//!
//! The typed sort domains now live in `litchi_xlsx::sort`; this host module
//! keeps the historical path available to the OOXML facade.

pub use litchi_xlsx::sort::*;
