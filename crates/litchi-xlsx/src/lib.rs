//! Typed Excel Open XML documents.
//!
//! The ordinary API exposes immutable, cheap-to-share workbook and sheet
//! handles. Package relationships and physical identifiers remain in [`raw`].

#![forbid(unsafe_code)]

pub mod calculation_properties;
pub mod cell;
pub mod chain;
pub mod column;
mod error;
pub mod formula;
pub mod layout;
pub mod merge;
mod outline;
pub mod raw;
pub mod row;
pub mod sheet;
pub mod sheet_calculation_properties;
pub mod style;
pub mod web;
mod workbook;

pub use calculation_properties::{
    WorkbookCalculationMode, WorkbookCalculationProperties, WorkbookReferenceMode,
    parse_workbook_calculation_properties,
};
pub use cell::{Cell, Cells, Content, Date, ErrorValue, Extents, Number, Text, Value};
pub use column::{Column, Columns, Width, WidthAt};
pub use error::{
    ColumnEditBlock, DefaultsEditBlock, EditBlock, Error, MergeEditBlock, RemoveBlock, RenameBlock,
    Result, RowEditBlock, TabEditBlock,
};
pub use formula::Formula;
pub use litchi_sheet::{
    Area, At, Cell as Address, Column as ColumnIndex, ColumnAt, Rect, Row as RowIndex, RowAt,
};
pub use outline::{Outline, OutlineAt};
pub use row::{Height, HeightAt, Row, Rows};
pub use sheet_calculation_properties::{
    WorksheetSheetCalculationProperties, WorksheetSheetCalculationPropertiesConformance,
    parse_worksheet_sheet_calculation_properties, write_worksheet_sheet_calculation_properties,
};
pub use style::{LocalStyle, Style, StyleKey, StyleState, Styles, StylesIter};
pub use workbook::{
    ActiveTab, Change, ColumnEdit, Commit, Conflict, ConflictSet, DateSystem, DefaultsEdit, Edit,
    Flavor, JoinError, JoinFailure, NewSheet, PackageChange, Patch, RowEdit, Sheet, SheetEdit,
    SheetKind, SheetSelector, State, TabEdit, Visibility, Workbook,
};
