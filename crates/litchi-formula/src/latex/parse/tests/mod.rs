// Unit tests for the LaTeX parser
//
// These assert on AST shape. String-level round trips through `LatexConverter`
// live in the crate's `tests/latex_parser.rs` integration test.

mod errors;
mod syntax;
mod vocabulary;

use super::{DEFAULT_MAX_DEPTH, LatexParseError, LatexParser};
use crate::ast::{
    AccentType, Fence, FractionType, FunctionName, LargeOperator, MathNode, MatrixFence, Operator,
    Position, PredefinedSymbol, SpaceType, StyleType,
};
use std::borrow::Cow;

/// Parse `input`, failing the test with the parser's error if it does not.
pub(super) fn parse(input: &str) -> Vec<MathNode<'_>> {
    match LatexParser::new().parse(input) {
        Ok(nodes) => nodes,
        Err(error) => panic!("`{input}` should parse but failed: {error}"),
    }
}

/// Parse `input` expecting exactly one node.
pub(super) fn parse_one(input: &str) -> MathNode<'_> {
    let mut nodes = parse(input);
    assert_eq!(nodes.len(), 1, "`{input}` should yield exactly one node");
    nodes.remove(0)
}

/// Parse `input` expecting failure, returning the error.
pub(super) fn parse_err(input: &str) -> LatexParseError {
    match LatexParser::new().parse(input) {
        Ok(nodes) => panic!("`{input}` should fail but produced {nodes:?}"),
        Err(error) => error,
    }
}

/// Build a borrowed text node.
pub(super) fn text(value: &str) -> MathNode<'_> {
    MathNode::Text(Cow::Borrowed(value))
}

/// Build a borrowed number node.
pub(super) fn number(value: &str) -> MathNode<'_> {
    MathNode::Number(Cow::Borrowed(value))
}
