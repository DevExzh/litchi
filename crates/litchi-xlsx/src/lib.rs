//! Typed Excel Open XML documents.
//!
//! The ordinary API exposes immutable, cheap-to-share workbook and sheet
//! handles. Package relationships and physical identifiers remain in [`raw`].

#![forbid(unsafe_code)]

pub mod calculation_properties;
pub mod cell;
pub mod cell_watches;
pub mod chain;
pub mod column;
pub mod connections;
mod error;
pub mod formula;
pub mod ignored_errors;
pub mod layout;
pub mod merge;
mod outline;
pub mod outline_properties;
pub mod page_margins;
pub mod phonetic_properties;
pub mod print_options;
pub mod raw;
pub mod row;
pub mod scenarios;
pub mod sheet;
pub mod sheet_calculation_properties;
pub mod sheet_view;
pub mod style;
pub mod views;
pub mod volatile_dependencies;
pub mod web;
mod workbook;
pub mod xml_maps;

pub use calculation_properties::{
    WorkbookCalculationMode, WorkbookCalculationProperties, WorkbookReferenceMode,
    parse_workbook_calculation_properties,
};
pub use cell::{Cell, Cells, Content, Date, ErrorValue, Extents, Number, Text, Value};
pub use cell_watches::{
    CellWatchReference, WorksheetCellWatchConformance, WorksheetCellWatches,
    parse_worksheet_cell_watches, write_worksheet_cell_watches,
};
pub use column::{Column, Columns, Width, WidthAt};
pub use error::{
    ColumnEditBlock, DefaultsEditBlock, EditBlock, Error, MergeEditBlock, RemoveBlock, RenameBlock,
    Result, RowEditBlock, TabEditBlock,
};
pub use formula::Formula;
pub use ignored_errors::{
    IgnoredErrorRangeReference, WorksheetIgnoredError, WorksheetIgnoredErrorType,
    WorksheetIgnoredErrors, WorksheetIgnoredErrorsExtension, parse_worksheet_ignored_errors,
};
pub use litchi_sheet::{
    Area, At, Cell as Address, Column as ColumnIndex, ColumnAt, Rect, Row as RowIndex, RowAt,
};
pub use outline::{Outline, OutlineAt};
pub use outline_properties::{WorksheetOutlineProperties, parse_worksheet_outline_properties};
pub use page_margins::{PageMargin, WorksheetPageMargins, parse_worksheet_page_margins};
pub use phonetic_properties::{
    WorksheetPhoneticAlignment, WorksheetPhoneticProperties, WorksheetPhoneticType,
    parse_worksheet_phonetic_properties,
};
pub use print_options::{WorksheetPrintOptions, parse_worksheet_print_options};
pub use row::{Height, HeightAt, Row, Rows};
pub use scenarios::{
    ScenarioCellReference, ScenarioRangeReference, WorksheetScenario, WorksheetScenarioConformance,
    WorksheetScenarioInputCell, WorksheetScenarios, parse_worksheet_scenarios,
    write_worksheet_scenarios,
};
pub use sheet_calculation_properties::{
    WorksheetSheetCalculationProperties, WorksheetSheetCalculationPropertiesConformance,
    parse_worksheet_sheet_calculation_properties, write_worksheet_sheet_calculation_properties,
};
pub use sheet_view::{
    PivotAreaType, PivotSelectionAxis, WorksheetCellReference, WorksheetPanePosition,
    WorksheetPaneState, WorksheetPivotArea, WorksheetPivotSelection, WorksheetRangeReference,
    WorksheetViewCollection, WorksheetViewDefinition, WorksheetViewExtension, WorksheetViewPane,
    WorksheetViewSelection, WorksheetViewSqref, WorksheetViewType, parse_worksheet_views,
};
pub use style::{LocalStyle, Style, StyleKey, StyleState, Styles, StylesIter};
pub use views::{
    SheetPane, SheetPanePosition, SheetPaneState, SheetSelection, SheetView, SheetViewType,
};
pub use workbook::{
    ActiveTab, Change, ColumnEdit, Commit, Conflict, ConflictSet, DateSystem, DefaultsEdit, Edit,
    Flavor, JoinError, JoinFailure, NewSheet, PackageChange, Patch, RowEdit, Sheet, SheetEdit,
    SheetKind, SheetSelector, State, TabEdit, Visibility, Workbook,
};
