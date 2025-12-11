//! AST types for formula expressions.

use crate::sheet::CellValue;

/// Binary operators supported by the expression parser.
#[derive(Debug, Clone, Copy)]
pub enum BinaryOp {
    Add,
    Sub,
    Mul,
    Div,
    Eq,
    Ne,
    Gt,
    Ge,
    Lt,
    Le,
}

/// A rectangular cell range reference.
#[derive(Debug, Clone)]
pub struct RangeRef {
    pub sheet: String,
    pub start_row: u32,
    pub start_col: u32,
    pub end_row: u32,
    pub end_col: u32,
}

/// Minimal expression AST used by the evaluator runtime.
#[derive(Debug, Clone)]
pub enum Expr {
    /// Literal value (number, string, boolean).
    Literal(CellValue),
    /// Single-cell reference.
    Reference { sheet: String, row: u32, col: u32 },
    /// Rectangular range reference.
    Range(RangeRef),
    /// Unary minus (e.g., -A1, -1).
    UnaryMinus(Box<Expr>),
    /// Binary arithmetic operation.
    Binary {
        op: BinaryOp,
        left: Box<Expr>,
        right: Box<Expr>,
    },
    /// Function call, e.g. SUM(A1:B3).
    FunctionCall { name: String, args: Vec<Expr> },
}
