//! Typed Excel Open XML documents.
//!
//! The ordinary API exposes immutable, cheap-to-share workbook and sheet
//! handles. Package relationships and physical identifiers remain in [`raw`].

#![forbid(unsafe_code)]

pub mod cell;
mod error;
pub mod formula;
pub mod raw;
pub mod style;
mod workbook;

pub use cell::{Cell, Cells, Content, Date, ErrorValue, Extents, Number, Text, Value};
pub use error::{EditBlock, Error, Result};
pub use formula::Formula;
pub use litchi_sheet::{Area, At, Cell as Address, Column, Rect, Row};
pub use style::{LocalStyle, Style, StyleKey, StyleState, Styles, StylesIter};
pub use workbook::{
    Change, Commit, Conflict, ConflictSet, DateSystem, Edit, Flavor, JoinError, JoinFailure, Patch,
    Sheet, SheetEdit, SheetKind, SheetSelector, State, Visibility, Workbook,
};
