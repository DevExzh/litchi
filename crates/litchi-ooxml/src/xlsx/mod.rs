//! Temporary OOXML-host facade for the standalone XLSX owner.
//!
//! XLSX models, package orchestration, codecs, and resource graph operations
//! are implemented by [`litchi_xlsx`].  The migration host exposes that owner
//! only while the surrounding `litchi-ooxml` crate is being retired; it does
//! not contain a second SpreadsheetML implementation.

pub use litchi_xlsx::*;
