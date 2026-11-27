//! Binary formula parser and typed token facade.
//!
//! The BIFF12 wire parser remains owned by [`crate::formula`]. This private
//! adapter keeps the host facade's parser surface in one contextual layer,
//! while the package-specific relationship resolution stays in `resolution`.

pub use crate::formula::{
    ArrayValue, BinaryOperator, Compiler, ExternalTableReference, Group, GroupKind,
    MAX_CELL_FORMULA_BYTES, MemoryKind, ParsedFormula, Parser, Range, Resolution, TableColumns,
    TableDataType, TableNamedColumns, TableReference, TableRowType, Token, UnaryOperator,
    ptg_types,
};
