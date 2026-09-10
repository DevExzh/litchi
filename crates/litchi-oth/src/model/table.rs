//! Read-only projection of ODF text tables.

use crate::paragraph::Paragraph;

/// The semantic kind of a table cell.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
#[non_exhaustive]
pub enum CellKind {
    /// A normal `table:table-cell`.
    Cell,
    /// A covered cell represented by `table:covered-table-cell`.
    Covered,
}

/// A projected table column declaration.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Column {
    default_cell_style_name: Option<String>,
    repeated: usize,
    style_name: Option<String>,
}

impl Column {
    pub(crate) const fn projected(
        style_name: Option<String>,
        default_cell_style_name: Option<String>,
        repeated: usize,
    ) -> Self {
        Self {
            default_cell_style_name,
            repeated,
            style_name,
        }
    }

    /// Column style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Default cell style reference.
    #[must_use]
    pub fn default_cell_style_name(&self) -> Option<&str> {
        self.default_cell_style_name.as_deref()
    }

    /// Number of physical columns represented by this declaration.
    #[must_use]
    pub const fn repeat_count(&self) -> usize {
        self.repeated
    }
}

/// A projected table cell.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Cell {
    columns_spanned: usize,
    formula: Option<String>,
    kind: CellKind,
    paragraphs: Vec<Paragraph>,
    repeated: usize,
    rows_spanned: usize,
    style_name: Option<String>,
    text: String,
    value: Option<String>,
    value_type: Option<String>,
}

impl Cell {
    #[allow(
        clippy::too_many_arguments,
        reason = "projection fields mirror ODF cell attributes"
    )]
    pub(crate) fn projected(
        kind: CellKind,
        style_name: Option<String>,
        repeated: usize,
        columns_spanned: usize,
        rows_spanned: usize,
        formula: Option<String>,
        value_type: Option<String>,
        value: Option<String>,
        text: String,
        paragraphs: Vec<Paragraph>,
    ) -> Self {
        Self {
            columns_spanned,
            formula,
            kind,
            paragraphs,
            repeated,
            rows_spanned,
            style_name,
            text,
            value,
            value_type,
        }
    }

    /// Whether this is a normal or covered cell.
    #[must_use]
    pub const fn kind(&self) -> CellKind {
        self.kind
    }

    /// Cell style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Number of repeated physical cells represented by this node.
    #[must_use]
    pub const fn repeat_count(&self) -> usize {
        self.repeated
    }

    /// Number of columns spanned by the cell, including the first column.
    #[must_use]
    pub const fn columns_spanned(&self) -> usize {
        self.columns_spanned
    }

    /// Number of rows spanned by the cell, including the first row.
    #[must_use]
    pub const fn rows_spanned(&self) -> usize {
        self.rows_spanned
    }

    /// Inert table formula, if present.
    #[must_use]
    pub fn formula(&self) -> Option<&str> {
        self.formula.as_deref()
    }

    /// ODF value type, if present.
    #[must_use]
    pub fn value_type(&self) -> Option<&str> {
        self.value_type.as_deref()
    }

    /// Stored lexical value, if present.
    #[must_use]
    pub fn value(&self) -> Option<&str> {
        self.value.as_deref()
    }

    /// Projected visible cell text. Formulas are never evaluated.
    #[must_use]
    pub fn text(&self) -> &str {
        &self.text
    }

    /// Direct paragraphs in source order.
    #[must_use]
    pub fn paragraphs(&self) -> &[Paragraph] {
        &self.paragraphs
    }
}

/// A projected table row.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Row {
    cells: Vec<Cell>,
    repeated: usize,
    style_name: Option<String>,
}

impl Row {
    pub(crate) const fn projected(
        style_name: Option<String>,
        repeated: usize,
        cells: Vec<Cell>,
    ) -> Self {
        Self {
            cells,
            repeated,
            style_name,
        }
    }

    /// Row style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Number of physical rows represented by this node.
    #[must_use]
    pub const fn repeat_count(&self) -> usize {
        self.repeated
    }

    /// Cells in source order.
    #[must_use]
    pub fn cells(&self) -> &[Cell] {
        &self.cells
    }
}

/// A projected `table:table` body structure.
#[derive(Clone, Debug, PartialEq, Eq)]
pub struct Table {
    columns: Vec<Column>,
    name: Option<String>,
    rows: Vec<Row>,
    style_name: Option<String>,
    declared_column_count: usize,
    column_count: usize,
}

impl Table {
    pub(crate) const fn projected(
        name: Option<String>,
        style_name: Option<String>,
        columns: Vec<Column>,
        rows: Vec<Row>,
        declared_column_count: usize,
        column_count: usize,
    ) -> Self {
        Self {
            columns,
            name,
            rows,
            style_name,
            declared_column_count,
            column_count,
        }
    }

    /// Producer-visible table name.
    #[must_use]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Table style reference.
    #[must_use]
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Explicit column declarations.
    #[must_use]
    pub fn columns(&self) -> &[Column] {
        &self.columns
    }

    /// Rows in source order.
    #[must_use]
    pub fn rows(&self) -> &[Row] {
        &self.rows
    }

    /// Number of physical columns declared by `table:table-column` nodes.
    /// Repeated declarations stay compact in [`Column::repeat_count`].
    #[must_use]
    pub const fn declared_column_count(&self) -> usize {
        self.declared_column_count
    }

    /// Logical table width after applying repeated cells and spans.
    ///
    /// This is the maximum of the declared column width and every row's
    /// checked expanded cell width. Repetition is retained as a count and is
    /// never materialized into additional cells.
    #[must_use]
    pub const fn column_count(&self) -> usize {
        self.column_count
    }
}
