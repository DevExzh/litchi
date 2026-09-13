//! Bounded, non-evaluating OpenFormula expression syntax.
//!
//! This module implements the expression grammar in ODF 1.4 Part 4 §§5.1–5.14
//! and retains the original source spelling in an immutable flat expression
//! tree.  It deliberately performs no name lookup, reference resolution,
//! function evaluation, or recalculation.  The older [`crate::codec::formula::FormulaParser`]
//! remains the compatibility tokenizer; this parser is an additive strict
//! grammar surface.
//!
//! ```
//! use litchi_ods::codec::formula::expression::{Expression, Kind, Limits};
//!
//! let limits = Limits::default().with_max_bytes(4096).with_max_nodes(512);
//! let expression = Expression::parse_with_limits("of:==SUM([.A1];;2)", &limits)?;
//! assert!(expression.is_force_recalculate());
//! let call = expression.root();
//! assert!(matches!(call.kind(), Kind::Function { name: "SUM" }));
//! assert_eq!(call.child_count(), 3);
//! assert!(call.child(1).unwrap().is_missing());
//! assert_eq!(call.child(0).unwrap().text(), "[.A1]");
//! # Ok::<(), litchi_core::Error>(())
//! ```

use super::reference;
use litchi_core::{Error, Resource, ResourceLimit, Result};
use std::{convert::TryFrom, sync::Arc};

mod names;

use names::{is_combining, is_digit, is_letter};

/// Default maximum UTF-8 byte length of one expression source.
pub const DEFAULT_MAX_EXPRESSION_BYTES: usize = 1024 * 1024;
/// Default maximum number of flat expression nodes.
pub const DEFAULT_MAX_EXPRESSION_NODES: usize = 65_536;
/// Default maximum recursive expression nesting.
pub const DEFAULT_MAX_EXPRESSION_DEPTH: usize = 256;
/// Default maximum number of inline-array cells across one expression.
pub const DEFAULT_MAX_ARRAY_CELLS: usize = 65_536;
/// Hard ceiling protecting the recursive-descent parser from caller-provided
/// unbounded depth limits.
pub const HARD_MAX_EXPRESSION_DEPTH: usize = 256;

/// Finite limits for one strict OpenFormula expression parse.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct Limits {
    max_bytes: usize,
    max_nodes: usize,
    max_depth: usize,
    max_array_cells: usize,
    reference: reference::Limits,
}

impl Default for Limits {
    fn default() -> Self {
        Self {
            max_bytes: DEFAULT_MAX_EXPRESSION_BYTES,
            max_nodes: DEFAULT_MAX_EXPRESSION_NODES,
            max_depth: DEFAULT_MAX_EXPRESSION_DEPTH,
            max_array_cells: DEFAULT_MAX_ARRAY_CELLS,
            reference: reference::Limits::default(),
        }
    }
}

impl Limits {
    /// Set the maximum UTF-8 byte length of the source expression.
    #[must_use]
    pub const fn with_max_bytes(mut self, value: usize) -> Self {
        self.max_bytes = value;
        self
    }

    /// Set the maximum number of flat expression nodes.
    #[must_use]
    pub const fn with_max_nodes(mut self, value: usize) -> Self {
        self.max_nodes = value;
        self
    }

    /// Set the requested recursive expression depth.  The parser also applies
    /// [`HARD_MAX_EXPRESSION_DEPTH`] even when this value is larger.
    #[must_use]
    pub const fn with_max_depth(mut self, value: usize) -> Self {
        self.max_depth = value;
        self
    }

    /// Set the maximum number of inline-array cells across the expression.
    #[must_use]
    pub const fn with_max_array_cells(mut self, value: usize) -> Self {
        self.max_array_cells = value;
        self
    }

    /// Set the finite limits passed to each bracketed reference.
    #[must_use]
    pub const fn with_reference_limits(mut self, value: reference::Limits) -> Self {
        self.reference = value;
        self
    }

    /// Return the maximum source byte length.
    #[must_use]
    pub const fn max_bytes(self) -> usize {
        self.max_bytes
    }

    /// Return the maximum node count.
    #[must_use]
    pub const fn max_nodes(self) -> usize {
        self.max_nodes
    }

    /// Return the caller-requested depth limit.
    #[must_use]
    pub const fn max_depth(self) -> usize {
        self.max_depth
    }

    /// Return the maximum aggregate inline-array cell count.
    #[must_use]
    pub const fn max_array_cells(self) -> usize {
        self.max_array_cells
    }

    /// Return the limits used for bracketed references.
    #[must_use]
    pub const fn reference_limits(self) -> reference::Limits {
        self.reference
    }
}

/// Prefix operators in the OpenFormula grammar.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PrefixOperator {
    /// Unary plus.
    Plus,
    /// Unary minus.
    Minus,
}

/// Postfix operators in the OpenFormula grammar.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PostfixOperator {
    /// Percentage postfix operator.
    Percent,
}

/// Infix operators in Table 1 of ODF 1.4 Part 4 §5.5.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum InfixOperator {
    /// Reference range operator.
    Range,
    /// Reference intersection operator.
    Intersection,
    /// Reference concatenation/union operator.
    Union,
    /// Power operator.
    Power,
    /// Multiplication operator.
    Multiply,
    /// Division operator.
    Divide,
    /// Addition operator.
    Add,
    /// Subtraction operator.
    Subtract,
    /// String concatenation operator.
    Concatenate,
    /// Equality comparison.
    Equal,
    /// Inequality comparison.
    NotEqual,
    /// Less-than comparison.
    Less,
    /// Less-than-or-equal comparison.
    LessEqual,
    /// Greater-than comparison.
    Greater,
    /// Greater-than-or-equal comparison.
    GreaterEqual,
}

/// Scope and lexical qualification of a named expression.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum NameScope<'a> {
    /// A current-document name. `forced` records the `$$` spelling that
    /// forces named-expression interpretation.
    Simple {
        /// Whether the source used the `$$` forced-name marker.
        forced: bool,
    },
    /// A name attached to a quoted sheet name.
    SheetLocal {
        /// Raw quoted-sheet content, with doubled apostrophes retained.
        sheet: &'a str,
        /// Whether the sheet name carried its absolute `$` marker.
        absolute: bool,
        /// Whether the source used the `$$` spelling for the name.
        forced: bool,
    },
    /// A name qualified by an inert external source IRI.
    External {
        /// Raw source-IRI content, with doubled apostrophes retained.
        source: &'a str,
        /// Optional raw quoted-sheet content after the source marker.
        sheet: Option<&'a str>,
        /// Whether the optional sheet name carried its absolute `$` marker.
        sheet_absolute: bool,
        /// Whether the source used the `$$` spelling for the name.
        forced: bool,
    },
}

/// Dimensions observed for one inline array.
///
/// The grammar accepts nonempty ragged rows. `columns` is `Some` only when
/// every row has the same number of cells; `min_columns` and `max_columns`
/// retain the shape of a ragged array without imposing the evaluator-only
/// rectangularity condition from §5.13.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ArrayDimensions {
    rows: usize,
    columns: Option<usize>,
    min_columns: usize,
    max_columns: usize,
}

impl ArrayDimensions {
    /// Number of nonempty matrix rows.
    #[must_use]
    pub const fn rows(self) -> usize {
        self.rows
    }

    /// Common column count, or `None` for a ragged matrix.
    #[must_use]
    pub const fn columns(self) -> Option<usize> {
        self.columns
    }

    /// Smallest row width.
    #[must_use]
    pub const fn min_columns(self) -> usize {
        self.min_columns
    }

    /// Largest row width.
    #[must_use]
    pub const fn max_columns(self) -> usize {
        self.max_columns
    }

    /// Whether every row has the same width.
    #[must_use]
    pub const fn is_rectangular(self) -> bool {
        self.columns.is_some()
    }
}

/// A borrowed semantic kind for one expression node.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum Kind<'a> {
    /// A syntactically valid number; use [`Node::text`] for its exact lexeme.
    Number,
    /// A double-quoted string; use [`Node::text`] for its exact escaped lexeme.
    String,
    /// A prefix unary operation.
    Prefix(PrefixOperator),
    /// A postfix unary operation.
    Postfix(PostfixOperator),
    /// A binary operation.
    Infix(InfixOperator),
    /// Parentheses retained as a source-level grouping node.
    Parenthesized,
    /// A function call with a host-defined or standard name.
    Function {
        /// Function name span, excluding the opening parenthesis.
        name: &'a str,
    },
    /// A complete inert bracketed reference.
    Reference(&'a reference::Reference),
    /// A single-quoted label.
    QuotedLabel,
    /// Two quoted labels joined by automatic intersection (`!!`).
    AutomaticIntersection,
    /// A simple, sheet-local, or external named expression.
    NamedExpression {
        /// Name content, excluding `$$` and quote delimiters where present.
        name: &'a str,
        /// Name qualification and source metadata.
        scope: NameScope<'a>,
    },
    /// A constant formula error such as `#N/A`.
    Error,
    /// An inline array. Its direct children are [`ArrayRow`](Self::ArrayRow)
    /// nodes.
    Array(ArrayDimensions),
    /// One matrix row inside an inline array. Its children are cell expressions.
    ArrayRow,
    /// An explicitly missing function parameter.
    Missing,
}

/// An immutable, owned, flat OpenFormula expression tree.
#[derive(Debug)]
pub struct Expression {
    source: String,
    nodes: Vec<NodeRecord>,
    edges: Vec<NodeId>,
    references: Vec<reference::Reference>,
    force_recalculate: bool,
    root: NodeId,
}

impl Expression {
    /// Parse an expression using the default finite limits.
    pub fn parse(source: &str) -> Result<Self> {
        Self::parse_with_limits(source, &Limits::default())
    }

    /// Parse an expression using explicit finite limits.
    pub fn parse_with_limits(source: &str, limits: &Limits) -> Result<Self> {
        if source.len() > limits.max_bytes {
            return Err(limit_error(
                Resource::InputBytes,
                source.len(),
                limits.max_bytes,
            ));
        }

        // Admit the immutable source before any parser-owned storage is
        // created. Every node and edge allocation below is separately
        // fallible; the source remains authoritative for every Node::text view.
        let mut owned_source = String::new();
        owned_source
            .try_reserve_exact(source.len())
            .map_err(|error| Error::Allocation {
                resource: "formula expression source",
                source: error,
            })?;
        owned_source.push_str(source);

        let parser = Parser::new(&owned_source, *limits);
        let (root, nodes, edges, references, force_recalculate) = parser.parse()?;
        Ok(Self {
            source: owned_source,
            nodes,
            edges,
            references,
            force_recalculate,
            root,
        })
    }

    /// Return the exact source spelling supplied to the parser.
    #[must_use]
    pub fn source(&self) -> &str {
        &self.source
    }

    /// Return whether the formula used the optional forced-recalculation
    /// marker (`==` or `of:==`).
    #[must_use]
    pub fn is_force_recalculate(&self) -> bool {
        self.force_recalculate
    }

    /// Return the root node of the expression tree.
    #[must_use]
    pub fn root(&self) -> Node<'_> {
        Node {
            expression: self,
            id: self.root,
        }
    }

    /// Number of nodes in the flat arena.
    #[must_use]
    pub fn node_count(&self) -> usize {
        self.nodes.len()
    }

    /// Number of parent-to-child edges in the flat arena.
    #[must_use]
    pub fn edge_count(&self) -> usize {
        self.edges.len()
    }
}

/// A borrowed node view into an [`Expression`].
#[derive(Clone, Copy)]
pub struct Node<'a> {
    expression: &'a Expression,
    id: NodeId,
}

impl<'a> Node<'a> {
    /// Return the source span covered by this node, including its delimiters.
    #[must_use]
    pub fn text(&self) -> &'a str {
        slice(&self.expression.source, self.record().span).unwrap_or_default()
    }

    /// Return the typed node kind and borrowed names/references.
    #[must_use]
    pub fn kind(&self) -> Kind<'a> {
        let record = self.record();
        match &record.kind {
            RecordKind::Number => Kind::Number,
            RecordKind::String => Kind::String,
            RecordKind::Prefix(operator) => Kind::Prefix(*operator),
            RecordKind::Postfix(operator) => Kind::Postfix(*operator),
            RecordKind::Infix(operator) => Kind::Infix(*operator),
            RecordKind::Parenthesized => Kind::Parenthesized,
            RecordKind::Function { name } => Kind::Function {
                name: slice(&self.expression.source, *name).unwrap_or_default(),
            },
            RecordKind::Reference(reference) => {
                Kind::Reference(&self.expression.references[reference.0])
            },
            RecordKind::QuotedLabel { .. } => Kind::QuotedLabel,
            RecordKind::AutomaticIntersection => Kind::AutomaticIntersection,
            RecordKind::Named { name, scope } => Kind::NamedExpression {
                name: slice(&self.expression.source, *name).unwrap_or_default(),
                scope: scope.as_public(&self.expression.source),
            },
            RecordKind::Error => Kind::Error,
            RecordKind::Array { dimensions } => Kind::Array(*dimensions),
            RecordKind::ArrayRow => Kind::ArrayRow,
            RecordKind::Missing => Kind::Missing,
        }
    }

    /// Return the direct children in source/grammar order.
    #[must_use]
    pub fn children(&self) -> Children<'a> {
        let record = self.record();
        let end = record
            .children_start
            .checked_add(record.children_len)
            .expect("validated expression child span overflow");
        let ids = self
            .expression
            .edges
            .get(record.children_start..end)
            .expect("validated expression child span out of bounds");
        Children {
            expression: self.expression,
            ids,
            position: 0,
        }
    }

    /// Return one direct child by zero-based index.
    #[must_use]
    pub fn child(&self, index: usize) -> Option<Node<'a>> {
        let record = self.record();
        if index >= record.children_len {
            return None;
        }
        let edge = record.children_start.checked_add(index)?;
        let id = *self.expression.edges.get(edge)?;
        Some(Node {
            expression: self.expression,
            id,
        })
    }

    /// Return the number of direct children.
    #[must_use]
    pub fn child_count(&self) -> usize {
        self.record().children_len
    }

    pub(super) fn arena_index(&self) -> usize {
        self.id.0
    }

    /// Return a function or named-expression name span, if this node has one.
    #[must_use]
    pub fn name(&self) -> Option<&'a str> {
        match &self.record().kind {
            RecordKind::Function { name } | RecordKind::Named { name, .. } => {
                slice(&self.expression.source, *name)
            },
            _ => None,
        }
    }

    /// Return a function name span, if this node is a function call.
    #[must_use]
    pub fn function_name(&self) -> Option<&'a str> {
        match &self.record().kind {
            RecordKind::Function { name } => slice(&self.expression.source, *name),
            _ => None,
        }
    }

    /// Return the raw content of a quoted label, excluding quote delimiters.
    #[must_use]
    pub fn label(&self) -> Option<&'a str> {
        match &self.record().kind {
            RecordKind::QuotedLabel { value } => slice(&self.expression.source, *value),
            _ => None,
        }
    }

    /// Return a borrowed parsed reference, if this node is a reference.
    #[must_use]
    pub fn reference(&self) -> Option<&'a reference::Reference> {
        match &self.record().kind {
            RecordKind::Reference(reference) => Some(&self.expression.references[reference.0]),
            _ => None,
        }
    }

    /// Return inline-array dimensions, if this node is an array.
    #[must_use]
    pub fn array_dimensions(&self) -> Option<ArrayDimensions> {
        match &self.record().kind {
            RecordKind::Array { dimensions } => Some(*dimensions),
            _ => None,
        }
    }

    /// Return whether this node represents a missing function parameter.
    #[must_use]
    pub fn is_missing(&self) -> bool {
        matches!(self.record().kind, RecordKind::Missing)
    }

    fn record(&self) -> &NodeRecord {
        // Node values are created only from parser-validated IDs. Returning a
        // static empty record would conceal an internal invariant violation;
        // this branch is unreachable for safe public callers.
        &self.expression.nodes[self.id.0]
    }
}

impl std::fmt::Debug for Node<'_> {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter
            .debug_struct("Node")
            .field("kind", &self.kind())
            .field("text", &self.text())
            .finish()
    }
}

/// An allocation-free iterator over borrowed node children.
pub struct Children<'a> {
    expression: &'a Expression,
    ids: &'a [NodeId],
    position: usize,
}

impl<'a> Iterator for Children<'a> {
    type Item = Node<'a>;

    fn next(&mut self) -> Option<Self::Item> {
        let id = *self.ids.get(self.position)?;
        self.position += 1;
        Some(Node {
            expression: self.expression,
            id,
        })
    }

    fn size_hint(&self) -> (usize, Option<usize>) {
        let remaining = self.ids.len().saturating_sub(self.position);
        (remaining, Some(remaining))
    }
}

impl ExactSizeIterator for Children<'_> {}
impl std::iter::FusedIterator for Children<'_> {}

#[derive(Clone, Copy, Debug)]
struct Span {
    start: usize,
    end: usize,
}

#[derive(Clone, Copy, Debug)]
struct NodeId(usize);

#[derive(Clone, Copy, Debug)]
struct ReferenceId(usize);

#[derive(Debug)]
struct NodeRecord {
    span: Span,
    children_start: usize,
    children_len: usize,
    kind: RecordKind,
}

#[derive(Debug)]
enum RecordKind {
    Number,
    String,
    Prefix(PrefixOperator),
    Postfix(PostfixOperator),
    Infix(InfixOperator),
    Parenthesized,
    Function { name: Span },
    Reference(ReferenceId),
    QuotedLabel { value: Span },
    AutomaticIntersection,
    Named { name: Span, scope: RecordScope },
    Error,
    Array { dimensions: ArrayDimensions },
    ArrayRow,
    Missing,
}

#[derive(Clone, Copy, Debug)]
enum RecordScope {
    Simple {
        forced: bool,
    },
    SheetLocal {
        sheet: Span,
        absolute: bool,
        forced: bool,
    },
    External {
        source: Span,
        sheet: Option<Span>,
        sheet_absolute: bool,
        forced: bool,
    },
}

impl RecordScope {
    fn as_public<'a>(self, source: &'a str) -> NameScope<'a> {
        match self {
            Self::Simple { forced } => NameScope::Simple { forced },
            Self::SheetLocal {
                sheet,
                absolute,
                forced,
            } => NameScope::SheetLocal {
                sheet: slice(source, sheet).unwrap_or_default(),
                absolute,
                forced,
            },
            Self::External {
                source: source_span,
                sheet,
                sheet_absolute,
                forced,
            } => NameScope::External {
                source: slice(source, source_span).unwrap_or_default(),
                sheet: sheet.and_then(|span| slice(source, span)),
                sheet_absolute,
                forced,
            },
        }
    }
}

fn slice(source: &str, span: Span) -> Option<&str> {
    source.get(span.start..span.end)
}

struct Parser<'a> {
    source: &'a str,
    bytes: &'a [u8],
    position: usize,
    limits: Limits,
    effective_depth: usize,
    nodes: Vec<NodeRecord>,
    edges: Vec<NodeId>,
    references: Vec<reference::Reference>,
    array_cells: usize,
}

impl<'a> Parser<'a> {
    fn new(source: &'a str, limits: Limits) -> Self {
        Self {
            source,
            bytes: source.as_bytes(),
            position: 0,
            limits,
            effective_depth: limits.max_depth.min(HARD_MAX_EXPRESSION_DEPTH),
            nodes: Vec::new(),
            edges: Vec::new(),
            references: Vec::new(),
            array_cells: 0,
        }
    }

    fn parse(
        mut self,
    ) -> Result<(
        NodeId,
        Vec<NodeRecord>,
        Vec<NodeId>,
        Vec<reference::Reference>,
        bool,
    )> {
        self.skip_whitespace();
        let force_recalculate = self.parse_intro()?;
        self.skip_whitespace();
        let root = self.parse_expression(0, 0)?;
        self.skip_whitespace();
        if !self.at_end() {
            return Err(invalid("trailing bytes after OpenFormula expression"));
        }
        Ok((
            root,
            self.nodes,
            self.edges,
            self.references,
            force_recalculate,
        ))
    }

    fn parse_intro(&mut self) -> Result<bool> {
        if self.starts_with_ascii_case_insensitive(b"of:") {
            self.position += 3;
            self.skip_whitespace();
            if !self.consume_byte(b'=') {
                return Err(invalid(
                    "OpenFormula namespace wrapper must be followed by '='",
                ));
            }
            return Ok(self.consume_force_recalculate());
        }

        if self.consume_byte(b'=') {
            return Ok(self.consume_force_recalculate());
        }
        Ok(false)
    }

    fn consume_force_recalculate(&mut self) -> bool {
        let checkpoint = self.position;
        self.skip_whitespace();
        if self.consume_byte(b'=') {
            self.skip_whitespace();
            true
        } else {
            self.position = checkpoint;
            self.skip_whitespace();
            false
        }
    }

    fn parse_expression(&mut self, minimum_precedence: u8, depth: usize) -> Result<NodeId> {
        self.check_depth(depth)?;
        self.skip_whitespace();
        let mut left = self.parse_prefix(depth)?;

        loop {
            self.skip_whitespace();
            if let Some(operator) = self.peek_postfix() {
                if POSTFIX_PRECEDENCE < minimum_precedence {
                    break;
                }
                self.position += 1;
                let span = Span {
                    start: self.node(left).span.start,
                    end: self.position,
                };
                left = self.push_node(span, RecordKind::Postfix(operator), &[left])?;
                continue;
            }
            if let Some((operator, precedence, width)) = self.peek_infix() {
                if precedence < minimum_precedence {
                    break;
                }
                self.position += width;
                let right =
                    self.parse_expression(precedence.saturating_add(1), depth.saturating_add(1))?;
                let span = Span {
                    start: self.node(left).span.start,
                    end: self.node(right).span.end,
                };
                left = self.push_node(span, RecordKind::Infix(operator), &[left, right])?;
                continue;
            }
            break;
        }
        Ok(left)
    }

    fn parse_prefix(&mut self, depth: usize) -> Result<NodeId> {
        self.skip_whitespace();
        let start = self.position;
        let operator = match self.peek() {
            Some(b'+') => Some(PrefixOperator::Plus),
            Some(b'-') => Some(PrefixOperator::Minus),
            _ => None,
        };
        let Some(operator) = operator else {
            return self.parse_primary(depth);
        };
        self.position += 1;
        let operand = self.parse_expression(PREFIX_PRECEDENCE, depth.saturating_add(1))?;
        let span = Span {
            start,
            end: self.node(operand).span.end,
        };
        self.push_node(span, RecordKind::Prefix(operator), &[operand])
    }

    fn parse_primary(&mut self, depth: usize) -> Result<NodeId> {
        self.skip_whitespace();
        let start = self.position;
        match self.peek() {
            Some(b'(') => self.parse_parenthesized(start, depth),
            Some(b'[') => self.parse_reference(start),
            Some(b'{') => self.parse_array(start, depth),
            Some(b'"') => self.parse_string(start),
            Some(b'\'') => self.parse_quoted_primary(start, depth),
            Some(b'#') => self.parse_error(start),
            Some(byte) if byte.is_ascii_digit() => self.parse_number(start),
            Some(b'.') if self.bytes.get(start + 1).is_some_and(u8::is_ascii_digit) => {
                self.parse_number(start)
            },
            Some(b'$') => self.parse_dollar_name(start, depth),
            Some(_) if self.is_letter_at(start) => self.parse_identifier_primary(start, depth),
            Some(_) => Err(invalid("expected an OpenFormula expression")),
            None => Err(invalid("expected an OpenFormula expression")),
        }
    }

    fn parse_parenthesized(&mut self, start: usize, depth: usize) -> Result<NodeId> {
        self.position += 1;
        let child = self.parse_expression(0, depth.saturating_add(1))?;
        self.skip_whitespace();
        if !self.consume_byte(b')') {
            return Err(invalid("parenthesized expression is missing ')'"));
        }
        self.push_node(
            Span {
                start,
                end: self.position,
            },
            RecordKind::Parenthesized,
            &[child],
        )
    }

    fn parse_string(&mut self, start: usize) -> Result<NodeId> {
        self.position += 1;
        let content_start = self.position;
        let closing = loop {
            let Some(byte) = self.peek() else {
                return Err(invalid("unterminated OpenFormula string literal"));
            };
            if byte != b'"' {
                self.position += 1;
                continue;
            }
            if self.bytes.get(self.position + 1) == Some(&b'"') {
                self.position += 2;
                continue;
            }
            break self.position;
        };
        let content = &self.bytes[content_start..closing];
        if content.contains(&0) {
            return Err(invalid("NUL is not allowed in OpenFormula string literal"));
        }
        self.position = closing + 1;
        self.push_node(
            Span {
                start,
                end: self.position,
            },
            RecordKind::String,
            &[],
        )
    }

    fn parse_number(&mut self, start: usize) -> Result<NodeId> {
        if self.consume_byte(b'.') {
            if !self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                return Err(invalid("fractional number requires digits after '.'"));
            }
            self.consume_ascii_digits();
        } else {
            if !self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                return Err(invalid("number requires a digit"));
            }
            self.consume_ascii_digits();
            if self.consume_byte(b'.') {
                if !self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                    return Err(invalid("number requires digits after its decimal point"));
                }
                self.consume_ascii_digits();
            }
        }

        if matches!(self.peek(), Some(b'e' | b'E')) {
            self.position += 1;
            if matches!(self.peek(), Some(b'+' | b'-')) {
                self.position += 1;
            }
            if !self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                return Err(invalid("scientific number requires an exponent"));
            }
            self.consume_ascii_digits();
        }

        self.push_node(
            Span {
                start,
                end: self.position,
            },
            RecordKind::Number,
            &[],
        )
    }

    fn parse_error(&mut self, start: usize) -> Result<NodeId> {
        self.position += 1;
        let body_start = self.position;
        while self
            .peek()
            .is_some_and(|byte| byte.is_ascii_uppercase() || byte.is_ascii_digit())
        {
            self.position += 1;
        }
        if self.position == body_start {
            return Err(invalid(
                "formula error name requires uppercase letters or digits",
            ));
        }

        match self.peek() {
            Some(b'!' | b'?') => self.position += 1,
            Some(b'/') => {
                self.position += 1;
                if self.peek().is_some_and(|byte| byte.is_ascii_uppercase()) {
                    self.position += 1;
                } else if self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
                    self.position += 1;
                    if !matches!(self.peek(), Some(b'!' | b'?')) {
                        return Err(invalid("numeric formula error suffix needs ! or ?"));
                    }
                    self.position += 1;
                } else {
                    return Err(invalid("formula error slash suffix is malformed"));
                }
            },
            _ => return Err(invalid("formula error name is missing its suffix")),
        }

        self.push_node(
            Span {
                start,
                end: self.position,
            },
            RecordKind::Error,
            &[],
        )
    }

    fn parse_reference(&mut self, start: usize) -> Result<NodeId> {
        // A parsed reference owns its complete endpoint/source metadata. Do
        // the node admission before invoking the bounded reference parser so
        // a zero-node expression limit cannot trigger component allocations.
        self.reserve_node_slot()?;
        let closing = self.scan_reference_close(start)?;
        let body = self
            .source
            .get(start + 1..closing)
            .ok_or_else(|| invalid("reference span is not valid UTF-8"))?;
        let reference = reference::parse_body(body, &self.limits.reference)?;
        let reference_id = self.push_reference(reference)?;
        self.position = closing + 1;
        self.push_node(
            Span {
                start,
                end: self.position,
            },
            RecordKind::Reference(reference_id),
            &[],
        )
    }

    fn parse_array(&mut self, start: usize, depth: usize) -> Result<NodeId> {
        self.position += 1;
        let mut rows = Vec::new();
        let mut row_count = 0usize;
        let mut min_columns = usize::MAX;
        let mut max_columns = 0usize;

        loop {
            self.skip_whitespace();
            if matches!(self.peek(), Some(b'}' | b'|' | b';') | None) {
                return Err(invalid("inline array rows and cells cannot be empty"));
            }

            let row_start = self.position;
            let mut cells = Vec::new();
            loop {
                self.bump_array_cell()?;
                let cell = self.parse_expression(0, depth.saturating_add(1))?;
                push_child(&mut cells, cell)?;
                self.skip_whitespace();
                if !self.consume_byte(b';') {
                    break;
                }
                self.skip_whitespace();
                if matches!(self.peek(), Some(b'}' | b'|' | b';') | None) {
                    return Err(invalid("inline array cells cannot be empty"));
                }
            }

            let row_end = cells
                .last()
                .map(|id| self.node(*id).span.end)
                .unwrap_or(row_start);
            let row = self.push_node(
                Span {
                    start: row_start,
                    end: row_end,
                },
                RecordKind::ArrayRow,
                &cells,
            )?;
            push_child(&mut rows, row)?;
            row_count = row_count
                .checked_add(1)
                .ok_or_else(|| invalid("inline array row count overflow"))?;
            min_columns = min_columns.min(cells.len());
            max_columns = max_columns.max(cells.len());

            self.skip_whitespace();
            if self.consume_byte(b'|') {
                continue;
            }
            if !self.consume_byte(b'}') {
                return Err(invalid("inline array requires '|' or '}' after each row"));
            }
            break;
        }

        let dimensions = ArrayDimensions {
            rows: row_count,
            columns: (min_columns == max_columns).then_some(min_columns),
            min_columns,
            max_columns,
        };
        self.push_node(
            Span {
                start,
                end: self.position,
            },
            RecordKind::Array { dimensions },
            &rows,
        )
    }

    fn parse_identifier_primary(&mut self, start: usize, depth: usize) -> Result<NodeId> {
        let name = self.scan_identifier()?;
        if self.peek() == Some(b'(') {
            return self.parse_function(start, name, depth);
        }
        validate_named_identifier(self.slice(name))?;
        self.push_node(
            Span {
                start,
                end: name.end,
            },
            RecordKind::Named {
                name,
                scope: RecordScope::Simple { forced: false },
            },
            &[],
        )
    }

    fn parse_dollar_name(&mut self, start: usize, depth: usize) -> Result<NodeId> {
        if self
            .bytes
            .get(self.position..self.position + 2)
            .is_some_and(|bytes| bytes == b"$$")
        {
            let (name, forced) = self.scan_simple_name(true)?;
            self.push_node(
                Span {
                    start,
                    end: self.position,
                },
                RecordKind::Named {
                    name,
                    scope: RecordScope::Simple { forced },
                },
                &[],
            )
        } else if self.bytes.get(self.skip_whitespace_from(self.position + 1)) == Some(&b'\'') {
            self.position = self.skip_whitespace_from(self.position + 1);
            self.parse_sheet_local_name(start, true, depth)
        } else {
            Err(invalid("'$' must begin a $$ name or quoted sheet name"))
        }
    }

    fn parse_quoted_primary(&mut self, start: usize, depth: usize) -> Result<NodeId> {
        let (close, value) = self.scan_single_quoted(true)?;
        self.position = close + 1;
        let probe = self.skip_whitespace_from(self.position);
        match self.bytes.get(probe).copied() {
            Some(b'#') => {
                self.position = probe + 1;
                self.parse_external_name(start, value, depth)
            },
            Some(b'.') => {
                self.position = probe + 1;
                self.parse_sheet_local_suffix(start, value, depth)
            },
            _ => {
                if value.start == value.end {
                    return Err(invalid("quoted labels cannot be empty"));
                }
                let label = self.push_node(
                    Span {
                        start,
                        end: close + 1,
                    },
                    RecordKind::QuotedLabel { value },
                    &[],
                )?;
                let intersection_probe = self.skip_whitespace_from(close + 1);
                if self
                    .bytes
                    .get(intersection_probe..intersection_probe + 2)
                    .is_some_and(|bytes| bytes == b"!!")
                {
                    self.position = intersection_probe + 2;
                    self.skip_whitespace();
                    let second_start = self.position;
                    let (second_close, second_value) = self.scan_single_quoted(false)?;
                    self.position = second_close + 1;
                    let second = self.push_node(
                        Span {
                            start: second_start,
                            end: second_close + 1,
                        },
                        RecordKind::QuotedLabel {
                            value: second_value,
                        },
                        &[],
                    )?;
                    return self.push_node(
                        Span {
                            start,
                            end: self.position,
                        },
                        RecordKind::AutomaticIntersection,
                        &[label, second],
                    );
                }
                self.position = close + 1;
                Ok(label)
            },
        }
    }

    fn parse_sheet_local_name(
        &mut self,
        start: usize,
        absolute_marker: bool,
        depth: usize,
    ) -> Result<NodeId> {
        let (close, sheet) = self.scan_single_quoted(true)?;
        self.position = close + 1;
        let separator = self.skip_whitespace_from(self.position);
        if self.bytes.get(separator) != Some(&b'.') {
            return Err(invalid("quoted sheet name must be followed by '.'"));
        }
        self.position = separator + 1;
        self.parse_sheet_local_suffix_with_marker(start, sheet, absolute_marker, depth)
    }

    fn parse_sheet_local_suffix(
        &mut self,
        start: usize,
        sheet: Span,
        depth: usize,
    ) -> Result<NodeId> {
        self.parse_sheet_local_suffix_with_marker(start, sheet, false, depth)
    }

    fn parse_sheet_local_suffix_with_marker(
        &mut self,
        start: usize,
        sheet: Span,
        absolute_marker: bool,
        _depth: usize,
    ) -> Result<NodeId> {
        let (name, forced) = self.scan_simple_name(true)?;
        self.push_node(
            Span {
                start,
                end: self.position,
            },
            RecordKind::Named {
                name,
                scope: RecordScope::SheetLocal {
                    sheet,
                    absolute: absolute_marker,
                    forced,
                },
            },
            &[],
        )
    }

    fn parse_external_name(&mut self, start: usize, source: Span, _depth: usize) -> Result<NodeId> {
        self.reserve_node_slot()?;
        self.validate_source_iri(source)?;
        self.skip_whitespace();
        let absolute_sheet = self.peek() == Some(b'$')
            && self.bytes.get(self.skip_whitespace_from(self.position + 1)) == Some(&b'\'');
        if absolute_sheet || self.peek() == Some(b'\'') {
            if absolute_sheet {
                self.position += 1;
                self.skip_whitespace();
            }
            let (close, sheet) = self.scan_single_quoted(true)?;
            self.position = close + 1;
            let separator = self.skip_whitespace_from(self.position);
            if self.bytes.get(separator) != Some(&b'.') {
                return Err(invalid(
                    "external quoted sheet name must be followed by '.'",
                ));
            }
            self.position = separator + 1;
            let (name, forced) = self.scan_simple_name(true)?;
            return self.push_node(
                Span {
                    start,
                    end: self.position,
                },
                RecordKind::Named {
                    name,
                    scope: RecordScope::External {
                        source,
                        sheet: Some(sheet),
                        sheet_absolute: absolute_sheet,
                        forced,
                    },
                },
                &[],
            );
        }
        if self
            .bytes
            .get(self.position..self.position + 2)
            .is_some_and(|bytes| bytes == b"$$")
        {
            let (name, forced) = self.scan_simple_name(true)?;
            return self.push_node(
                Span {
                    start,
                    end: self.position,
                },
                RecordKind::Named {
                    name,
                    scope: RecordScope::External {
                        source,
                        sheet: None,
                        sheet_absolute: false,
                        forced,
                    },
                },
                &[],
            );
        }

        let (name, forced) = self.scan_simple_name(true)?;
        self.push_node(
            Span {
                start,
                end: self.position,
            },
            RecordKind::Named {
                name,
                scope: RecordScope::External {
                    source,
                    sheet: None,
                    sheet_absolute: false,
                    forced,
                },
            },
            &[],
        )
    }

    fn parse_function(&mut self, start: usize, name: Span, depth: usize) -> Result<NodeId> {
        self.position += 1;
        let mut parameters = Vec::new();
        self.skip_whitespace();
        if !self.consume_byte(b')') {
            loop {
                self.skip_whitespace();
                let parameter = if matches!(self.peek(), Some(b';' | b')') | None) {
                    self.push_node(
                        Span {
                            start: self.position,
                            end: self.position,
                        },
                        RecordKind::Missing,
                        &[],
                    )?
                } else {
                    self.parse_expression(0, depth.saturating_add(1))?
                };
                push_child(&mut parameters, parameter)?;
                self.skip_whitespace();
                if self.consume_byte(b')') {
                    break;
                }
                if !self.consume_byte(b';') {
                    return Err(invalid("function parameters require ';' or ')'"));
                }
            }
        }

        self.push_node(
            Span {
                start,
                end: self.position,
            },
            RecordKind::Function { name },
            &parameters,
        )
    }

    fn scan_identifier(&mut self) -> Result<Span> {
        let start = self.position;
        let Some((character, width)) = self.character_at(self.position) else {
            return Err(invalid("expected an identifier"));
        };
        if !is_letter(character) {
            return Err(invalid("identifier must begin with an XML Letter"));
        }
        self.position += width;
        while let Some((character, width)) = self.character_at(self.position) {
            if !(is_letter(character)
                || is_digit(character)
                || is_combining(character)
                || matches!(character, '_' | '.'))
            {
                break;
            }
            self.position += width;
        }
        Ok(Span {
            start,
            end: self.position,
        })
    }

    fn peek_postfix(&self) -> Option<PostfixOperator> {
        (self.peek() == Some(b'%')).then_some(PostfixOperator::Percent)
    }

    fn scan_simple_name(&mut self, allow_forced_marker: bool) -> Result<(Span, bool)> {
        // §5.14 permits whitespace between grammar components, while the
        // $$ marker and Identifier/SingleQuoted terminals remain contiguous.
        self.skip_whitespace();
        let forced = allow_forced_marker
            && self
                .bytes
                .get(self.position..self.position + 2)
                .is_some_and(|bytes| bytes == b"$$");
        if forced {
            self.position += 2;
            self.skip_whitespace();
        }
        if self.peek() == Some(b'\'') {
            if !forced {
                return Err(invalid("quoted named expressions require the $$ marker"));
            }
            let (close, value) = self.scan_single_quoted(false)?;
            self.position = close + 1;
            return Ok((value, true));
        }
        let name = self.scan_identifier()?;
        validate_named_identifier(self.slice(name))?;
        Ok((name, forced))
    }

    fn scan_single_quoted(&self, allow_empty: bool) -> Result<(usize, Span)> {
        if self.peek() != Some(b'\'') {
            return Err(invalid("expected a single-quoted component"));
        }
        let opening = self.position;
        let mut cursor = opening + 1;
        while cursor < self.bytes.len() {
            if self.bytes[cursor] != b'\'' {
                cursor += 1;
                continue;
            }
            if self.bytes.get(cursor + 1) == Some(&b'\'') {
                cursor += 2;
                continue;
            }
            if !allow_empty && cursor == opening + 1 {
                return Err(invalid("single-quoted component cannot be empty"));
            }
            return Ok((
                cursor,
                Span {
                    start: opening + 1,
                    end: cursor,
                },
            ));
        }
        Err(invalid("unterminated single-quoted component"))
    }

    fn validate_source_iri(&self, source: Span) -> Result<()> {
        let raw = self.slice(source);
        let mut pairs = 0usize;
        let mut cursor = 0usize;
        while cursor < raw.len() {
            if raw.as_bytes()[cursor] == b'\'' {
                if raw.as_bytes().get(cursor + 1) != Some(&b'\'') {
                    return Err(invalid("source IRI contains an invalid apostrophe escape"));
                }
                pairs = pairs
                    .checked_add(1)
                    .ok_or_else(|| invalid("source IRI escape count overflow"))?;
                cursor += 2;
            } else {
                cursor += 1;
            }
        }
        let decoded_len = raw
            .len()
            .checked_sub(pairs)
            .ok_or_else(|| invalid("source IRI length overflow"))?;
        let maximum = self.limits.reference.max_name_bytes();
        if decoded_len > maximum {
            return Err(limit_error(Resource::InputBytes, decoded_len, maximum));
        }
        if pairs == 0 {
            if reference::is_valid_iri_reference(raw) {
                return Ok(());
            }
            return Err(invalid("invalid OpenFormula source IRI"));
        }

        let mut decoded = String::new();
        decoded
            .try_reserve_exact(decoded_len)
            .map_err(|error| Error::Allocation {
                resource: "formula expression source IRI",
                source: error,
            })?;
        let mut segment_start = 0usize;
        cursor = 0;
        while cursor < raw.len() {
            if raw.as_bytes()[cursor] == b'\'' {
                if segment_start < cursor {
                    decoded.push_str(&raw[segment_start..cursor]);
                }
                decoded.push('\'');
                cursor += 2;
                segment_start = cursor;
            } else {
                cursor += 1;
            }
        }
        if segment_start < raw.len() {
            decoded.push_str(&raw[segment_start..]);
        }
        if reference::is_valid_iri_reference(&decoded) {
            Ok(())
        } else {
            Err(invalid("invalid OpenFormula source IRI"))
        }
    }

    fn scan_reference_close(&self, start: usize) -> Result<usize> {
        let mut cursor = start + 1;
        let mut quoted = false;
        while cursor < self.bytes.len() {
            match self.bytes[cursor] {
                b'\'' => {
                    if quoted && self.bytes.get(cursor + 1) == Some(&b'\'') {
                        cursor += 2;
                    } else {
                        quoted = !quoted;
                        cursor += 1;
                    }
                },
                b']' if !quoted => return Ok(cursor),
                _ => cursor += 1,
            }
        }
        Err(invalid("unterminated bracketed reference"))
    }

    fn check_depth(&self, depth: usize) -> Result<()> {
        if depth > self.effective_depth {
            return Err(limit_error(Resource::Depth, depth, self.effective_depth));
        }
        Ok(())
    }

    fn bump_array_cell(&mut self) -> Result<()> {
        self.array_cells = self
            .array_cells
            .checked_add(1)
            .ok_or_else(|| invalid("inline array cell count overflow"))?;
        if self.array_cells > self.limits.max_array_cells {
            return Err(limit_error(
                Resource::Objects,
                self.array_cells,
                self.limits.max_array_cells,
            ));
        }
        Ok(())
    }

    fn push_node(&mut self, span: Span, kind: RecordKind, children: &[NodeId]) -> Result<NodeId> {
        self.reserve_node_slot()?;
        self.reserve_edges(children.len())?;
        let children_start = self.edges.len();
        self.edges.extend_from_slice(children);
        let id = NodeId(self.nodes.len());
        self.nodes.push(NodeRecord {
            span,
            children_start,
            children_len: children.len(),
            kind,
        });
        Ok(id)
    }

    fn reserve_node_slot(&mut self) -> Result<()> {
        let observed = self
            .nodes
            .len()
            .checked_add(1)
            .ok_or_else(|| invalid("expression node count overflow"))?;
        if observed > self.limits.max_nodes {
            return Err(limit_error(
                Resource::Objects,
                observed,
                self.limits.max_nodes,
            ));
        }
        if self.nodes.len() == self.nodes.capacity() {
            self.nodes
                .try_reserve(1)
                .map_err(|error| Error::Allocation {
                    resource: "formula expression nodes",
                    source: error,
                })?;
        }
        Ok(())
    }

    fn push_reference(&mut self, reference: reference::Reference) -> Result<ReferenceId> {
        let observed = self
            .references
            .len()
            .checked_add(1)
            .ok_or_else(|| invalid("expression reference count overflow"))?;
        if observed > self.limits.max_nodes {
            return Err(limit_error(
                Resource::Objects,
                observed,
                self.limits.max_nodes,
            ));
        }
        if self.references.len() == self.references.capacity() {
            self.references
                .try_reserve(1)
                .map_err(|error| Error::Allocation {
                    resource: "formula expression references",
                    source: error,
                })?;
        }
        let id = ReferenceId(self.references.len());
        self.references.push(reference);
        Ok(id)
    }

    fn reserve_edges(&mut self, additional: usize) -> Result<()> {
        let observed = self
            .edges
            .len()
            .checked_add(additional)
            .ok_or_else(|| invalid("expression edge count overflow"))?;
        if observed > self.limits.max_nodes {
            return Err(limit_error(
                Resource::Objects,
                observed,
                self.limits.max_nodes,
            ));
        }
        self.edges
            .try_reserve(additional)
            .map_err(|error| Error::Allocation {
                resource: "formula expression edges",
                source: error,
            })?;
        Ok(())
    }

    fn consume_ascii_digits(&mut self) {
        while self.peek().is_some_and(|byte| byte.is_ascii_digit()) {
            self.position += 1;
        }
    }

    fn is_letter_at(&self, position: usize) -> bool {
        self.character_at(position)
            .is_some_and(|(character, _)| is_letter(character))
    }

    fn character_at(&self, position: usize) -> Option<(char, usize)> {
        let character = self.source.get(position..)?.chars().next()?;
        Some((character, character.len_utf8()))
    }

    fn skip_whitespace_from(&self, mut position: usize) -> usize {
        while matches!(self.bytes.get(position), Some(b' ' | b'\t' | b'\n' | b'\r')) {
            position += 1;
        }
        position
    }

    fn skip_whitespace(&mut self) {
        self.position = self.skip_whitespace_from(self.position);
    }

    fn peek_infix(&self) -> Option<(InfixOperator, u8, usize)> {
        let first = *self.bytes.get(self.position)?;
        let second = self.bytes.get(self.position + 1).copied();
        let (operator, width) = match (first, second) {
            (b'<', Some(b'>')) => (InfixOperator::NotEqual, 2),
            (b'<', Some(b'=')) => (InfixOperator::LessEqual, 2),
            (b'>', Some(b'=')) => (InfixOperator::GreaterEqual, 2),
            (b':', _) => (InfixOperator::Range, 1),
            (b'!', _) => (InfixOperator::Intersection, 1),
            (b'~', _) => (InfixOperator::Union, 1),
            (b'^', _) => (InfixOperator::Power, 1),
            (b'*', _) => (InfixOperator::Multiply, 1),
            (b'/', _) => (InfixOperator::Divide, 1),
            (b'+', _) => (InfixOperator::Add, 1),
            (b'-', _) => (InfixOperator::Subtract, 1),
            (b'&', _) => (InfixOperator::Concatenate, 1),
            (b'=', _) => (InfixOperator::Equal, 1),
            (b'<', _) => (InfixOperator::Less, 1),
            (b'>', _) => (InfixOperator::Greater, 1),
            _ => return None,
        };
        Some((operator, precedence(operator), width))
    }

    fn consume_byte(&mut self, expected: u8) -> bool {
        if self.peek() == Some(expected) {
            self.position += 1;
            true
        } else {
            false
        }
    }

    fn starts_with_ascii_case_insensitive(&self, expected: &[u8]) -> bool {
        self.bytes
            .get(self.position..self.position.saturating_add(expected.len()))
            .is_some_and(|value| value.eq_ignore_ascii_case(expected))
    }

    fn slice(&self, span: Span) -> &str {
        slice(self.source, span).unwrap_or_default()
    }

    fn node(&self, id: NodeId) -> &NodeRecord {
        &self.nodes[id.0]
    }

    fn peek(&self) -> Option<u8> {
        self.bytes.get(self.position).copied()
    }

    fn at_end(&self) -> bool {
        self.position == self.bytes.len()
    }
}

fn push_child(children: &mut Vec<NodeId>, child: NodeId) -> Result<()> {
    children.try_reserve(1).map_err(|error| Error::Allocation {
        resource: "formula expression child list",
        source: error,
    })?;
    children.push(child);
    Ok(())
}

fn validate_named_identifier(value: &str) -> Result<()> {
    if value.as_bytes().contains(&b'.') {
        return Err(invalid("periods are only allowed in function names"));
    }
    if value.eq_ignore_ascii_case("true") || value.eq_ignore_ascii_case("false") {
        return Err(invalid("True and False are not named expressions"));
    }
    let bytes = value.as_bytes();
    let mut position = 0;
    while bytes.get(position).is_some_and(u8::is_ascii_alphabetic) {
        position += 1;
    }
    if position != 0 && position < bytes.len() && bytes[position..].iter().all(u8::is_ascii_digit) {
        return Err(invalid("cell-shaped identifiers are not named expressions"));
    }
    Ok(())
}

fn precedence(operator: InfixOperator) -> u8 {
    match operator {
        InfixOperator::Range => 100,
        InfixOperator::Intersection => 95,
        InfixOperator::Union => 90,
        InfixOperator::Power => 70,
        InfixOperator::Multiply | InfixOperator::Divide => 60,
        InfixOperator::Add | InfixOperator::Subtract => 50,
        InfixOperator::Concatenate => 40,
        InfixOperator::Equal
        | InfixOperator::NotEqual
        | InfixOperator::Less
        | InfixOperator::LessEqual
        | InfixOperator::Greater
        | InfixOperator::GreaterEqual => 30,
    }
}

const PREFIX_PRECEDENCE: u8 = 80;
const POSTFIX_PRECEDENCE: u8 = 75;

fn invalid(message: impl Into<String>) -> Error {
    Error::InvalidFormat(message.into())
}

#[cold]
fn limit_error(resource: Resource, actual: usize, maximum: usize) -> Error {
    let Some(observed) = u64::try_from(actual).ok() else {
        return invalid("OpenFormula expression limit exceeds u64");
    };
    let Some(limit) = u64::try_from(maximum).ok() else {
        return invalid("OpenFormula expression limit exceeds u64");
    };
    Error::ResourceLimit(ResourceLimit {
        resource,
        observed,
        limit,
        scope: Arc::from("ods-formula-expression"),
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    fn child_kinds(node: Node<'_>) -> Vec<Kind<'_>> {
        node.children().map(|child| child.kind()).collect()
    }

    #[test]
    fn parses_intro_numbers_strings_and_operator_precedence() {
        let expression = Expression::parse("of:== -2^2+ .5e999 & \"a\"\"b\"")
            .expect("strict expression grammar");
        assert_eq!(expression.source(), "of:== -2^2+ .5e999 & \"a\"\"b\"");
        assert!(expression.is_force_recalculate());
        assert!(matches!(
            expression.root().kind(),
            Kind::Infix(InfixOperator::Concatenate)
        ));
        assert!(expression.node_count() >= 8);
    }

    #[test]
    fn parses_functions_with_missing_arguments_and_nested_children() {
        let expression = Expression::parse("F(;1;;\"x\")")
            .expect("missing function parameters are syntactically valid");
        let root = expression.root();
        assert!(matches!(root.kind(), Kind::Function { name: "F" }));
        assert_eq!(root.child_count(), 4);
        assert!(root.child(0).is_some_and(|node| node.is_missing()));
        assert!(root.child(2).is_some_and(|node| node.is_missing()));
    }

    #[test]
    fn parses_ragged_arrays_without_imposing_evaluator_rectangularity() {
        let expression = Expression::parse("{1;2|3}").expect("ragged grammar");
        let root = expression.root();
        let dimensions = root.array_dimensions().expect("array dimensions");
        assert_eq!(dimensions.rows(), 2);
        assert_eq!(dimensions.columns(), None);
        assert_eq!(dimensions.min_columns(), 1);
        assert_eq!(dimensions.max_columns(), 2);
        assert_eq!(root.children().count(), 2);
        assert!(matches!(child_kinds(root)[0], Kind::ArrayRow));
    }

    #[test]
    fn parses_reference_names_labels_and_errors_inertly() {
        let reference = Expression::parse("[.A1]").expect("reference");
        assert!(matches!(reference.root().kind(), Kind::Reference(_)));

        let name = Expression::parse("'Sheet'.Rate").expect("sheet-local name");
        assert!(matches!(
            name.root().kind(),
            Kind::NamedExpression {
                scope: NameScope::SheetLocal { sheet: "Sheet", .. },
                ..
            }
        ));

        let forced = Expression::parse("$$'Name'").expect("forced quoted name");
        assert!(matches!(
            forced.root().kind(),
            Kind::NamedExpression {
                scope: NameScope::Simple { forced: true },
                ..
            }
        ));
        assert!(Expression::parse("'Sheet'.'Name'").is_err());

        let absolute = Expression::parse("$'Sheet'.$$Name").expect("absolute sheet name");
        assert!(matches!(
            absolute.root().kind(),
            Kind::NamedExpression {
                scope: NameScope::SheetLocal {
                    sheet: "Sheet",
                    absolute: true,
                    forced: true,
                },
                ..
            }
        ));

        let label = Expression::parse("'row'!!'column'").expect("automatic intersection");
        assert!(matches!(label.root().kind(), Kind::AutomaticIntersection));

        let error = Expression::parse("#N/A").expect("constant error");
        assert!(matches!(error.root().kind(), Kind::Error));
    }

    #[test]
    fn depth_and_array_limits_are_typed() {
        let error =
            Expression::parse_with_limits("((((1))))", &Limits::default().with_max_depth(1))
                .expect_err("depth limit");
        assert!(matches!(error, Error::ResourceLimit(limit) if limit.resource == Resource::Depth));
        let error =
            Expression::parse_with_limits("{1;2}", &Limits::default().with_max_array_cells(1))
                .expect_err("array cell limit");
        assert!(
            matches!(error, Error::ResourceLimit(limit) if limit.resource == Resource::Objects)
        );

        let hostile = format!(
            "{}1{}",
            "(".repeat(HARD_MAX_EXPRESSION_DEPTH + 1),
            ")".repeat(HARD_MAX_EXPRESSION_DEPTH + 1)
        );
        let unbounded = Limits::default().with_max_depth(usize::MAX);
        let error = Expression::parse_with_limits(&hostile, &unbounded)
            .expect_err("the parser's hard depth ceiling must remain effective");
        assert!(matches!(error, Error::ResourceLimit(limit) if limit.resource == Resource::Depth));
    }
}
