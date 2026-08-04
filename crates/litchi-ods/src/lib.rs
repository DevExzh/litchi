//! OpenDocument Spreadsheet (`.ods`) support.
//!
//! The crate is organized by responsibility: immutable spreadsheet vocabulary
//! in [`model`], XML codecs in [`codec`], package access in [`package`],
//! construction in [`authoring`], and the concise entry points in [`facade`].

pub mod authoring;
pub mod codec;
pub mod facade;
pub mod model;
pub mod package;

pub use facade::{Builder, MutableSpreadsheet, Spreadsheet};
pub use litchi_odf_common::rdf;
pub use model::*;
