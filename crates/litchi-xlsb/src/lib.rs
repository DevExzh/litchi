//! Typed Excel Binary Workbook documents.
//!
//! [`raw`] owns the validated BIFF12 wire substrate, while [`calc`] owns the
//! strictly typed workbook calculation record. Additional semantic snapshots
//! and edits will be layered over them without exposing package identifiers in
//! ordinary APIs.

#![forbid(unsafe_code)]

pub mod calc;
pub mod comments;
pub mod conditional_formatting;
pub mod data_validation;
pub mod date_utils;
pub mod external_link;
pub mod formula;
pub mod hyperlinks;
pub mod merged_cells;
pub mod named_ranges;
pub mod pivot_view;
pub mod raw;
pub mod styles;

pub use raw::Error;

pub use pivot_view::{Part, PivotTableViewPart};

pub use data_validation::{
    DataValidation, DataValidationRecordKind, DataValidationSettings, FormulaBinary,
};
pub use formula::ptg_types;
pub use formula::{
    ArrayValue, BinaryOperator, CellParsedFormula, Compiler, Error as FormulaError,
    ExternalTableReference, FormulaArrayValue, FormulaConverter, FormulaExternalTableReference,
    FormulaGroup, FormulaGroupKind, FormulaMemoryKind, FormulaParser, FormulaRange,
    FormulaResolution, FormulaTableColumns, FormulaTableDataType, FormulaTableNamedColumns,
    FormulaTableReference, FormulaTableRowType, FormulaToken, Group, GroupKind,
    MAX_CELL_FORMULA_BYTES, MemoryKind, ParsedFormula, Parser, Range, Resolution,
    Result as FormulaResult, TableColumns, TableDataType, TableNamedColumns, TableReference,
    TableRowType, Token, UnaryOperator,
};
