//! Typed Excel Open XML documents.
//!
//! The ordinary API exposes immutable, cheap-to-share workbook and sheet
//! handles. Package relationships and physical identifiers remain in [`raw`].

#![forbid(unsafe_code)]

pub mod cell;
mod error;
pub mod formula;
pub mod raw;
mod workbook;

pub use cell::{Cell, Cells};
pub use error::{Error, Result};
pub use litchi_sheet::{At, Cell as Address, Column, Rect, Row};
pub use workbook::{DateSystem, Flavor, Sheet, SheetKind, SheetSelector, Visibility, Workbook};
