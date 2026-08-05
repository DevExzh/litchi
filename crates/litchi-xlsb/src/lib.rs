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
pub mod package;
pub mod pivot_view;
pub mod raw;
pub mod sheet;
pub mod styles;
pub mod workbook;
pub mod writer;

pub use raw::Error;

pub use package::Package;
pub use sheet::Worksheet;
pub use workbook::Workbook;

pub use pivot_view::Part;

pub use data_validation::{
    DataValidation, DataValidationRecordKind, DataValidationSettings, FormulaBinary,
};
pub use formula::ptg_types;
pub use formula::{
    ArrayValue, BinaryOperator, Compiler, Error as FormulaError, ExternalTableReference, Group,
    GroupKind, MAX_CELL_FORMULA_BYTES, MemoryKind, ParsedFormula, Parser, Range, Resolution,
    Result as FormulaResult, TableColumns, TableDataType, TableNamedColumns, TableReference,
    TableRowType, Token, UnaryOperator,
};
