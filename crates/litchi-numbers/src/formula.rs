//! Bounded, source-bound Numbers formula authoring values.
//!
//! This module deliberately exposes semantic references rather than native
//! table identifiers, UUID words, protobuf nodes, or encoded bytes. A cross-
//! table reference starts from a [`Table`] selected from the exact package
//! snapshot being edited. Formula values are opaque and can only be built by
//! constructors which enforce finite scalars and aggregate resource limits.

use std::fmt;

use crate::{
    Package, SheetSelector, TableSelector,
    cell::{FiniteF64, FiniteF64Error},
    package::table_cells::{Error as CellError, Path},
    table::{Dimensions, cells::Error as TableCellError},
};

/// Maximum UTF-8 bytes retained by one expression, including function names.
pub const MAX_OWNED_BYTES: usize = 1024 * 1024;
/// Maximum semantic nodes retained by one expression.
pub const MAX_NODES: usize = 65_536;
/// Maximum expression nesting accepted before allocation.
pub const MAX_DEPTH: usize = 64;
/// Maximum arguments retained by one function call.
pub const MAX_FUNCTION_ARGUMENTS: usize = 256;

/// A finite formula-construction resource.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum LimitKind {
    /// UTF-8 bytes owned by text literals, caches, and function names.
    OwnedBytes,
    /// Semantic expression nodes.
    Nodes,
    /// Expression nesting depth.
    Depth,
    /// Arguments in one function call.
    FunctionArguments,
    /// Encoded formula and dependency bytes admitted by package authoring.
    WireBytes,
    /// Fallible allocation events admitted by package authoring.
    AllocationEvents,
}

/// Failure from constructing a bounded formula value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
#[non_exhaustive]
pub enum Error {
    /// A numeric literal or cached result was not finite.
    NonFinite,
    /// A finite resource ceiling was exceeded.
    LimitExceeded {
        /// Resource which exceeded its ceiling.
        kind: LimitKind,
        /// Requested aggregate amount.
        observed: usize,
        /// Fixed public ceiling.
        maximum: usize,
    },
    /// A fallible allocation could not reserve the requested amount.
    Allocation {
        /// Resource being allocated.
        kind: LimitKind,
        /// Elements or bytes requested.
        amount: usize,
    },
    /// The function name is not in the authorable Numbers subset.
    UnsupportedFunction,
    /// The function argument count is invalid.
    InvalidArity,
    /// A range has reversed endpoints.
    ReversedRange,
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::NonFinite => formatter.write_str("Numbers formula scalar must be finite"),
            Self::LimitExceeded {
                observed, maximum, ..
            } => write!(
                formatter,
                "Numbers formula resource limit exceeded: observed {observed}, maximum {maximum}"
            ),
            Self::Allocation { amount, .. } => {
                write!(
                    formatter,
                    "could not allocate {amount} units for a Numbers formula"
                )
            },
            Self::UnsupportedFunction => {
                formatter.write_str("unsupported Numbers formula function")
            },
            Self::InvalidArity => formatter.write_str("invalid Numbers formula function arity"),
            Self::ReversedRange => formatter.write_str("Numbers formula range is reversed"),
        }
    }
}

impl std::error::Error for Error {}

impl From<FiniteF64Error> for Error {
    fn from(_error: FiniteF64Error) -> Self {
        Self::NonFinite
    }
}

/// A typed cached result stored beside an authored formula.
#[derive(Clone, PartialEq)]
pub struct CachedValue(CachedKind);

#[derive(Clone, PartialEq)]
pub(crate) enum CachedKind {
    Number(FiniteF64),
    Text(Box<str>),
    Boolean(bool),
    Date(FiniteF64),
    Duration(FiniteF64),
}

impl CachedValue {
    /// Construct a finite numeric result.
    pub fn number(value: f64) -> Result<Self, Error> {
        Ok(Self(CachedKind::Number(FiniteF64::new(value)?)))
    }

    /// Fallibly copy a bounded textual result.
    pub fn text(value: impl AsRef<str>) -> Result<Self, Error> {
        copy_text(value.as_ref(), MAX_OWNED_BYTES).map(|value| Self(CachedKind::Text(value)))
    }

    /// Construct a Boolean result.
    #[must_use]
    pub const fn boolean(value: bool) -> Self {
        Self(CachedKind::Boolean(value))
    }

    /// Construct a finite Apple-epoch date result.
    pub fn date(value: f64) -> Result<Self, Error> {
        Ok(Self(CachedKind::Date(FiniteF64::new(value)?)))
    }

    /// Construct a finite duration result.
    pub fn duration(value: f64) -> Result<Self, Error> {
        Ok(Self(CachedKind::Duration(FiniteF64::new(value)?)))
    }

    pub(crate) const fn owned_bytes(&self) -> usize {
        match &self.0 {
            CachedKind::Text(value) => value.len(),
            CachedKind::Number(_)
            | CachedKind::Boolean(_)
            | CachedKind::Date(_)
            | CachedKind::Duration(_) => 0,
        }
    }

    pub(crate) const fn kind(&self) -> &CachedKind {
        &self.0
    }
}

impl fmt::Debug for CachedValue {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        let kind = match &self.0 {
            CachedKind::Number(_) => "number",
            CachedKind::Text(_) => "text",
            CachedKind::Boolean(_) => "boolean",
            CachedKind::Date(_) => "date",
            CachedKind::Duration(_) => "duration",
        };
        formatter
            .debug_struct("CachedValue")
            .field("kind", &kind)
            .finish_non_exhaustive()
    }
}

/// A checked cell address used by a formula.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct CellReference {
    row: usize,
    column: usize,
    absolute_row: bool,
    absolute_column: bool,
}

impl CellReference {
    /// Construct a reference whose row and column move with the formula.
    #[must_use]
    pub const fn relative(row: usize, column: usize) -> Self {
        Self::mixed(row, column, false, false)
    }

    /// Construct a reference fixed to an absolute row and column.
    #[must_use]
    pub const fn absolute(row: usize, column: usize) -> Self {
        Self::mixed(row, column, true, true)
    }

    /// Construct a reference with independent row and column modes.
    #[must_use]
    pub const fn mixed(
        row: usize,
        column: usize,
        absolute_row: bool,
        absolute_column: bool,
    ) -> Self {
        Self {
            row,
            column,
            absolute_row,
            absolute_column,
        }
    }

    pub(crate) const fn coordinates(self) -> (usize, usize) {
        (self.row, self.column)
    }

    pub(crate) const fn modes(self) -> (bool, bool) {
        (self.absolute_row, self.absolute_column)
    }
}

/// A checked whole-row or whole-column endpoint.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub struct AxisReference {
    index: usize,
    absolute: bool,
}

impl AxisReference {
    /// Construct an endpoint that moves with the formula.
    #[must_use]
    pub const fn relative(index: usize) -> Self {
        Self {
            index,
            absolute: false,
        }
    }

    /// Construct an endpoint fixed to its absolute index.
    #[must_use]
    pub const fn absolute(index: usize) -> Self {
        Self {
            index,
            absolute: true,
        }
    }

    pub(crate) const fn parts(self) -> (usize, bool) {
        (self.index, self.absolute)
    }
}

/// A binary formula operator.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
#[non_exhaustive]
pub enum BinaryOperator {
    Add,
    Subtract,
    Multiply,
    Divide,
    Power,
    Concatenate,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
    Equal,
    NotEqual,
}

/// An opaque table selected from one exact package snapshot.
#[derive(Clone)]
pub struct Table {
    source: Package,
    path: Path,
    dimensions: Dimensions,
}

impl Table {
    pub(crate) const fn new(source: Package, path: Path, dimensions: Dimensions) -> Self {
        Self {
            source,
            path,
            dimensions,
        }
    }

    pub(crate) const fn source(&self) -> &Package {
        &self.source
    }

    pub(crate) const fn path(&self) -> Path {
        self.path
    }

    pub(crate) const fn dimensions(&self) -> Dimensions {
        self.dimensions
    }
}

impl fmt::Debug for Table {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.debug_struct("Table").finish_non_exhaustive()
    }
}

impl PartialEq for Table {
    fn eq(&self, other: &Self) -> bool {
        self.source.shares_snapshot(&other.source)
            && self.path == other.path
            && self.dimensions == other.dimensions
    }
}

/// One opaque, bounded formula expression.
#[derive(Clone, PartialEq)]
pub struct Expression {
    node: Node,
    metrics: Metrics,
}

#[derive(Clone, PartialEq)]
pub(crate) enum Node {
    Number(FiniteF64),
    Text(Box<str>),
    Boolean(bool),
    Cell(CellReference),
    TableCell(Table, CellReference),
    TableRange(Table, CellReference, CellReference),
    Rows(AxisReference, AxisReference),
    Columns(AxisReference, AxisReference),
    TableRows(Table, AxisReference, AxisReference),
    TableColumns(Table, AxisReference, AxisReference),
    Range(CellReference, CellReference),
    Function {
        name: Box<str>,
        arguments: Box<[Expression]>,
    },
    Binary {
        operator: BinaryOperator,
        operands: Box<[Expression]>,
    },
    Negate(Box<[Expression]>),
    Percent(Box<[Expression]>),
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct Metrics {
    nodes: usize,
    depth: usize,
    owned_bytes: usize,
}

impl Expression {
    /// Construct a finite numeric literal.
    pub fn number(value: f64) -> Result<Self, Error> {
        Ok(Self::leaf(Node::Number(FiniteF64::new(value)?)))
    }

    /// Fallibly copy a bounded text literal.
    pub fn text(value: impl AsRef<str>) -> Result<Self, Error> {
        let value = copy_text(value.as_ref(), MAX_OWNED_BYTES)?;
        let owned_bytes = value.len();
        Ok(Self {
            node: Node::Text(value),
            metrics: Metrics {
                nodes: 1,
                depth: 1,
                owned_bytes,
            },
        })
    }

    /// Construct a Boolean literal.
    #[must_use]
    pub const fn boolean(value: bool) -> Self {
        Self::leaf(Node::Boolean(value))
    }

    /// Construct a local cell reference.
    #[must_use]
    pub const fn cell(reference: CellReference) -> Self {
        Self::leaf(Node::Cell(reference))
    }

    /// Construct a cross-table cell reference.
    #[must_use]
    pub fn table_cell(table: &Table, reference: CellReference) -> Self {
        Self::leaf(Node::TableCell(table.clone(), reference))
    }

    /// Construct a local rectangular range.
    pub fn range(start: CellReference, end: CellReference) -> Result<Self, Error> {
        validate_cell_range(start, end)?;
        Ok(Self::leaf(Node::Range(start, end)))
    }

    /// Construct a cross-table rectangular range.
    pub fn table_range(
        table: &Table,
        start: CellReference,
        end: CellReference,
    ) -> Result<Self, Error> {
        validate_cell_range(start, end)?;
        Ok(Self::leaf(Node::TableRange(table.clone(), start, end)))
    }

    /// Construct a local whole-row range.
    pub fn rows(start: AxisReference, end: AxisReference) -> Result<Self, Error> {
        validate_axis_range(start, end)?;
        Ok(Self::leaf(Node::Rows(start, end)))
    }

    /// Construct a local whole-column range.
    pub fn columns(start: AxisReference, end: AxisReference) -> Result<Self, Error> {
        validate_axis_range(start, end)?;
        Ok(Self::leaf(Node::Columns(start, end)))
    }

    /// Construct a cross-table whole-row range.
    pub fn table_rows(
        table: &Table,
        start: AxisReference,
        end: AxisReference,
    ) -> Result<Self, Error> {
        validate_axis_range(start, end)?;
        Ok(Self::leaf(Node::TableRows(table.clone(), start, end)))
    }

    /// Construct a cross-table whole-column range.
    pub fn table_columns(
        table: &Table,
        start: AxisReference,
        end: AxisReference,
    ) -> Result<Self, Error> {
        validate_axis_range(start, end)?;
        Ok(Self::leaf(Node::TableColumns(table.clone(), start, end)))
    }

    /// Construct a checked call to one authorable Numbers function.
    pub fn function(
        name: impl AsRef<str>,
        arguments: impl IntoIterator<Item = Expression>,
    ) -> Result<Self, Error> {
        let name = name.as_ref();
        const MAX_SUPPORTED_FUNCTION_NAME_BYTES: usize = 7;
        if name.len() > MAX_SUPPORTED_FUNCTION_NAME_BYTES {
            return Err(Error::LimitExceeded {
                kind: LimitKind::OwnedBytes,
                observed: name.len(),
                maximum: MAX_SUPPORTED_FUNCTION_NAME_BYTES,
            });
        }
        let (canonical, minimum, maximum) =
            known_function(name).ok_or(Error::UnsupportedFunction)?;
        let mut retained = Vec::new();
        let mut metrics = Metrics {
            nodes: 1,
            depth: 1,
            owned_bytes: canonical.len(),
        };
        for argument in arguments {
            if retained.len() == MAX_FUNCTION_ARGUMENTS {
                return Err(Error::LimitExceeded {
                    kind: LimitKind::FunctionArguments,
                    observed: retained.len().saturating_add(1),
                    maximum: MAX_FUNCTION_ARGUMENTS,
                });
            }
            metrics = metrics.with_child(&argument)?;
            let required = retained.len().saturating_add(1);
            retained
                .try_reserve(1)
                .map_err(|_error| Error::Allocation {
                    kind: LimitKind::FunctionArguments,
                    amount: required,
                })?;
            retained.push(argument);
        }
        if !(minimum..=maximum).contains(&retained.len()) {
            return Err(Error::InvalidArity);
        }
        let name = copy_text(canonical, MAX_OWNED_BYTES)?;
        Ok(Self {
            node: Node::Function {
                name,
                arguments: retained.into_boxed_slice(),
            },
            metrics,
        })
    }

    /// Construct a binary expression.
    pub fn binary(
        operator: BinaryOperator,
        left: Expression,
        right: Expression,
    ) -> Result<Self, Error> {
        let metrics = Metrics::parent(&[&left, &right], 0)?;
        let operands = boxed_children([left, right], LimitKind::Nodes)?;
        Ok(Self {
            node: Node::Binary { operator, operands },
            metrics,
        })
    }

    /// Construct unary negation.
    pub fn negate(value: Expression) -> Result<Self, Error> {
        let metrics = Metrics::parent(&[&value], 0)?;
        Ok(Self {
            node: Node::Negate(boxed_children([value], LimitKind::Nodes)?),
            metrics,
        })
    }

    /// Construct a percentage expression.
    pub fn percent(value: Expression) -> Result<Self, Error> {
        let metrics = Metrics::parent(&[&value], 0)?;
        Ok(Self {
            node: Node::Percent(boxed_children([value], LimitKind::Nodes)?),
            metrics,
        })
    }

    const fn leaf(node: Node) -> Self {
        Self {
            node,
            metrics: Metrics {
                nodes: 1,
                depth: 1,
                owned_bytes: 0,
            },
        }
    }

    pub(crate) const fn root(&self) -> &Node {
        &self.node
    }

    pub(crate) const fn owned_bytes(&self) -> usize {
        self.metrics.owned_bytes
    }

    pub(crate) const fn node_count(&self) -> usize {
        self.metrics.nodes
    }

    pub(crate) fn validate_for(
        &self,
        source: &Package,
        local: Dimensions,
    ) -> Result<(), CellError> {
        validate_expression(self, source, local, 1)
    }
}

impl Metrics {
    fn parent(children: &[&Expression], own_bytes: usize) -> Result<Self, Error> {
        let mut metrics = Self {
            nodes: 1,
            depth: 1,
            owned_bytes: own_bytes,
        };
        for child in children {
            metrics = metrics.with_child(child)?;
        }
        Ok(metrics)
    }

    fn with_child(mut self, child: &Expression) -> Result<Self, Error> {
        self.nodes =
            checked_limit_add(self.nodes, child.metrics.nodes, MAX_NODES, LimitKind::Nodes)?;
        self.owned_bytes = checked_limit_add(
            self.owned_bytes,
            child.metrics.owned_bytes,
            MAX_OWNED_BYTES,
            LimitKind::OwnedBytes,
        )?;
        self.depth = self.depth.max(child.metrics.depth.saturating_add(1));
        if self.depth > MAX_DEPTH {
            return Err(Error::LimitExceeded {
                kind: LimitKind::Depth,
                observed: self.depth,
                maximum: MAX_DEPTH,
            });
        }
        Ok(self)
    }
}

impl fmt::Debug for Expression {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("Expression")
            .field("nodes", &self.metrics.nodes)
            .field("depth", &self.metrics.depth)
            .finish_non_exhaustive()
    }
}

pub(crate) fn resolve_table_handle<'sheet, 'table>(
    source: &Package,
    sheet_selector: impl Into<SheetSelector<'sheet>>,
    table_selector: impl Into<TableSelector<'table>>,
) -> Result<Table, TableCellError> {
    let sheet_selector = sheet_selector.into();
    let table_selector = table_selector.into();
    let (sheet_position, sheet) = match sheet_selector {
        SheetSelector::Index(position) => source
            .document()
            .sheets()
            .get(position)
            .map(|sheet| (position, sheet))
            .ok_or(TableCellError::SheetNotFound)?,
        SheetSelector::Name(name) => {
            let mut matches = source
                .document()
                .sheets()
                .iter()
                .enumerate()
                .filter(|(_position, sheet)| sheet.name() == name);
            let first = matches.next().ok_or(TableCellError::SheetNotFound)?;
            if matches.next().is_some() {
                return Err(TableCellError::AmbiguousSource {
                    path: Path::Package,
                });
            }
            first
        },
    };
    let compact_sheet =
        u32::try_from(sheet_position).map_err(|_error| TableCellError::InvalidSource {
            path: Path::Package,
        })?;
    let (table_position, table) = match table_selector {
        TableSelector::Index(position) => sheet
            .tables()
            .nth(position)
            .map(|table| (position, table))
            .ok_or(TableCellError::TableNotFound)?,
        TableSelector::Name(name) => {
            let mut matches = sheet
                .tables()
                .enumerate()
                .filter(|(_position, table)| table.name() == name);
            let first = matches.next().ok_or(TableCellError::TableNotFound)?;
            if matches.next().is_some() {
                return Err(TableCellError::AmbiguousSource {
                    path: Path::Table {
                        sheet: compact_sheet,
                        table: u32::try_from(first.0).unwrap_or(u32::MAX),
                    },
                });
            }
            first
        },
    };
    let compact_table =
        u32::try_from(table_position).map_err(|_error| TableCellError::InvalidSource {
            path: Path::Package,
        })?;
    Ok(Table::new(
        source.snapshot(),
        Path::Table {
            sheet: compact_sheet,
            table: compact_table,
        },
        table.dimensions(),
    ))
}

fn copy_text(value: &str, maximum: usize) -> Result<Box<str>, Error> {
    if value.len() > maximum {
        return Err(Error::LimitExceeded {
            kind: LimitKind::OwnedBytes,
            observed: value.len(),
            maximum,
        });
    }
    let mut owned = String::new();
    owned
        .try_reserve_exact(value.len())
        .map_err(|_error| Error::Allocation {
            kind: LimitKind::OwnedBytes,
            amount: value.len(),
        })?;
    owned.push_str(value);
    Ok(owned.into_boxed_str())
}

fn boxed_children<const N: usize>(
    children: [Expression; N],
    kind: LimitKind,
) -> Result<Box<[Expression]>, Error> {
    let mut retained = Vec::new();
    retained
        .try_reserve_exact(N)
        .map_err(|_error| Error::Allocation { kind, amount: N })?;
    if retained.capacity() != N {
        return Err(Error::Allocation {
            kind,
            amount: retained.capacity(),
        });
    }
    retained.extend(children);
    Ok(retained.into_boxed_slice())
}

fn checked_limit_add(
    current: usize,
    amount: usize,
    maximum: usize,
    kind: LimitKind,
) -> Result<usize, Error> {
    let observed = current.checked_add(amount).unwrap_or(usize::MAX);
    if observed > maximum {
        Err(Error::LimitExceeded {
            kind,
            observed,
            maximum,
        })
    } else {
        Ok(observed)
    }
}

fn validate_cell_range(start: CellReference, end: CellReference) -> Result<(), Error> {
    let (start_row, start_column) = start.coordinates();
    let (end_row, end_column) = end.coordinates();
    if start_row > end_row || start_column > end_column {
        Err(Error::ReversedRange)
    } else {
        Ok(())
    }
}

fn validate_axis_range(start: AxisReference, end: AxisReference) -> Result<(), Error> {
    if start.parts().0 > end.parts().0 {
        Err(Error::ReversedRange)
    } else {
        Ok(())
    }
}

fn validate_table(source: &Package, table: &Table) -> Result<(), CellError> {
    if source.shares_snapshot(table.source()) {
        Ok(())
    } else {
        Err(CellError::PatchConflict)
    }
}

fn validate_expression(
    expression: &Expression,
    source: &Package,
    local: Dimensions,
    depth: usize,
) -> Result<(), CellError> {
    if depth > MAX_DEPTH {
        return Err(CellError::UnsupportedSource {
            path: Path::Package,
        });
    }
    match expression.root() {
        Node::Cell(reference) => validate_cell(*reference, local),
        Node::TableCell(table, reference) => {
            validate_table(source, table)?;
            validate_cell(*reference, table.dimensions())
        },
        Node::Range(start, end) => {
            validate_cell(*start, local)?;
            validate_cell(*end, local)
        },
        Node::TableRange(table, start, end) => {
            validate_table(source, table)?;
            validate_cell(*start, table.dimensions())?;
            validate_cell(*end, table.dimensions())
        },
        Node::Rows(start, end) => {
            validate_axis(*start, local.rows())?;
            validate_axis(*end, local.rows())
        },
        Node::Columns(start, end) => {
            validate_axis(*start, local.columns())?;
            validate_axis(*end, local.columns())
        },
        Node::TableRows(table, start, end) => {
            validate_table(source, table)?;
            validate_axis(*start, table.dimensions().rows())?;
            validate_axis(*end, table.dimensions().rows())
        },
        Node::TableColumns(table, start, end) => {
            validate_table(source, table)?;
            validate_axis(*start, table.dimensions().columns())?;
            validate_axis(*end, table.dimensions().columns())
        },
        Node::Function { arguments, .. } => validate_children(arguments, source, local, depth),
        Node::Binary { operands, .. } | Node::Negate(operands) | Node::Percent(operands) => {
            validate_children(operands, source, local, depth)
        },
        Node::Number(_) | Node::Text(_) | Node::Boolean(_) => Ok(()),
    }
}

fn validate_children(
    children: &[Expression],
    source: &Package,
    local: Dimensions,
    depth: usize,
) -> Result<(), CellError> {
    let child_depth = depth.checked_add(1).ok_or(CellError::UnsupportedSource {
        path: Path::Package,
    })?;
    for child in children {
        validate_expression(child, source, local, child_depth)?;
    }
    Ok(())
}

fn validate_cell(reference: CellReference, dimensions: Dimensions) -> Result<(), CellError> {
    let (row, column) = reference.coordinates();
    if row >= dimensions.rows() as usize || column >= dimensions.columns() as usize {
        let (Ok(row), Ok(column)) = (u32::try_from(row), u32::try_from(column)) else {
            return Err(CellError::UnsupportedSource {
                path: Path::Package,
            });
        };
        Err(CellError::OutOfBounds {
            position: crate::table::CellPosition::new(row, column),
            dimensions,
        })
    } else {
        Ok(())
    }
}

fn validate_axis(reference: AxisReference, maximum: u32) -> Result<(), CellError> {
    let position = reference.parts().0;
    if position >= maximum as usize {
        Err(CellError::UnsupportedSource {
            path: Path::Package,
        })
    } else {
        Ok(())
    }
}

fn known_function(name: &str) -> Option<(&'static str, usize, usize)> {
    const FUNCTIONS: [(&str, usize, usize); 13] = [
        ("SUM", 1, MAX_FUNCTION_ARGUMENTS),
        ("AVERAGE", 1, MAX_FUNCTION_ARGUMENTS),
        ("MIN", 1, MAX_FUNCTION_ARGUMENTS),
        ("MAX", 1, MAX_FUNCTION_ARGUMENTS),
        ("COUNT", 1, MAX_FUNCTION_ARGUMENTS),
        ("COUNTA", 1, MAX_FUNCTION_ARGUMENTS),
        ("AND", 1, MAX_FUNCTION_ARGUMENTS),
        ("OR", 1, MAX_FUNCTION_ARGUMENTS),
        ("IF", 2, 3),
        ("IFERROR", 2, 2),
        ("NOT", 1, 1),
        ("ABS", 1, 1),
        ("ROUND", 2, 2),
    ];
    FUNCTIONS
        .into_iter()
        .find(|(candidate, _, _)| name.eq_ignore_ascii_case(candidate))
}
