//! Typed Excel Binary Workbook documents.
//!
//! [`raw`] owns the validated BIFF12 wire substrate, while [`calc`] owns the
//! strictly typed workbook calculation record. Additional semantic snapshots
//! and edits will be layered over them without exposing package identifiers in
//! ordinary APIs.

#![forbid(unsafe_code)]

pub mod calc;
pub mod conditional_formatting;
pub mod data_validation;
pub mod date_utils;
pub mod external_link;
pub mod formula;
pub mod hyperlinks;
pub mod merged_cells;
pub mod raw;

pub use raw::Error;

pub use data_validation::{
    DataValidation, DataValidationRecordKind, DataValidationSettings, FormulaBinary,
};
pub use formula::ptg_types;
pub use formula::{
    BinaryOperator, CellParsedFormula, Error as FormulaError, FormulaArrayValue,
    FormulaExternalTableReference, FormulaGroup, FormulaGroupKind, FormulaMemoryKind,
    FormulaParser, FormulaRange, FormulaTableColumns, FormulaTableDataType,
    FormulaTableNamedColumns, FormulaTableReference, FormulaTableRowType, FormulaToken,
    MAX_CELL_FORMULA_BYTES, Result as FormulaResult, UnaryOperator,
};
