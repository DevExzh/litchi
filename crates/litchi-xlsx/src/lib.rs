//! Typed Excel Open XML documents.
//!
//! The ordinary API exposes immutable, cheap-to-share workbook and sheet
//! handles. Package relationships and physical identifiers remain in [`raw`].

#![forbid(unsafe_code)]

pub mod cell;
pub mod column;
mod error;
pub mod formula;
pub mod raw;
pub mod row;
pub mod sheet;
pub mod style;
mod workbook;

pub use cell::{Cell, Cells, Content, Date, ErrorValue, Extents, Number, Text, Value};
pub use column::{Column, Columns};
pub use error::{
    ColumnEditBlock, EditBlock, Error, RemoveBlock, RenameBlock, Result, RowEditBlock, TabEditBlock,
};
pub use formula::Formula;
pub use litchi_sheet::{
    Area, At, Cell as Address, Column as ColumnIndex, ColumnAt, Rect, Row as RowIndex, RowAt,
};
pub use row::{Row, Rows};
pub use style::{LocalStyle, Style, StyleKey, StyleState, Styles, StylesIter};
pub use workbook::{
    ActiveTab, Change, ColumnEdit, ColumnState, Commit, Conflict, ConflictSet, DateSystem, Edit,
    Flavor, JoinError, JoinFailure, NewSheet, Patch, RowEdit, RowState, Sheet, SheetEdit,
    SheetKind, SheetSelector, State, TabEdit, Visibility, Workbook,
};
