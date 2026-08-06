//! Internal expression and reference model for formula-text compilation.

use super::super::{
    ArrayValue, BinaryOperator, TableNamedColumns, TableReference, TableRowType, UnaryOperator,
};
use super::builtin::BuiltinFunction;

#[derive(Debug)]
pub(super) enum CompileExpr {
    Number(f64),
    String(String),
    Bool(bool),
    Error(u8),
    MissingArg,
    Parenthesized(Box<CompileExpr>),
    Array {
        rows: u32,
        cols: u32,
        values: Vec<ArrayValue>,
    },
    Ref(A1Reference),
    Area(A1Reference, A1Reference),
    Ref3d(u16, A1Reference),
    Area3d(u16, A1Reference, A1Reference),
    Name(u32),
    TableReference(TableReference),
    Unary(UnaryOperator, Box<CompileExpr>),
    Binary(BinaryOperator, Box<CompileExpr>, Box<CompileExpr>),
    Function(BuiltinFunction, Vec<CompileExpr>),
}

#[derive(Debug)]
pub(super) struct ParsedStructuredReference {
    pub(super) row_type: TableRowType,
    pub(super) columns: TableNamedColumns,
    pub(super) square_bracket_space: bool,
    pub(super) comma_space: bool,
}

#[derive(Debug)]
pub(super) struct StructuredReferenceItem {
    pub(super) text: String,
    pub(super) first_character_escaped: bool,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct A1Reference {
    pub(super) row: u32,
    pub(super) col: u32,
    pub(super) row_relative: bool,
    pub(super) col_relative: bool,
}
