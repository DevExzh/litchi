//! Typed Excel Open XML documents.
//!
//! The ordinary API exposes immutable, cheap-to-share workbook and sheet
//! handles. Package relationships and physical identifiers remain in [`raw`].

#![forbid(unsafe_code)]

mod error;
pub mod raw;
mod workbook;

pub use error::{Error, Result};
pub use workbook::{DateSystem, Flavor, Sheet, SheetKind, SheetSelector, Visibility, Workbook};
