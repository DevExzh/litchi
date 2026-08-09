//! Typed Excel Binary Workbook documents.
//!
//! [`raw`] owns the validated BIFF12 wire substrate, while [`calc`] owns the
//! strictly typed workbook calculation record. Additional semantic snapshots
//! and edits will be layered over them without exposing package identifiers in
//! ordinary APIs.

#![forbid(unsafe_code)]

pub mod calc;
pub mod cell_values;
pub mod cell_watches;
pub mod chart;
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
mod pivot_chart;
pub mod pivot_view;
pub mod raw;
pub mod shapes;
pub mod shared_workbook;
pub mod sheet;
pub mod slicer;
pub mod sparkline;
pub mod styles;
pub mod timeline;
pub mod workbook;
pub mod xml_maps;

/// OPC resource limits used by XLSB package and workbook ingress.
pub use litchi_opc::ReadLimits;
pub mod writer;

pub use raw::Error;

pub use package::Package;
pub use package::scenarios;
pub use sheet::Worksheet;
pub use workbook::Workbook;

pub use pivot_view::Part;

pub use data_validation::{FormulaBinary, RecordKind, Settings, Validation};
pub use formula::ptg_types;
pub use formula::{
    ArrayValue, BinaryOperator, Compiler, Error as FormulaError, ExternalTableReference, Group,
    GroupKind, MAX_CELL_FORMULA_BYTES, MemoryKind, ParsedFormula, Parser, Range, Resolution,
    Result as FormulaResult, TableColumns, TableDataType, TableNamedColumns, TableReference,
    TableRowType, Token, UnaryOperator,
};
