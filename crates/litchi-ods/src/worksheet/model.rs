//! Typed worksheet vocabulary for the ODS facade.
//!
//! The model stores the physical ODF runs instead of expanding repeated rows
//! and cells.  A run is still addressable through the logical row/column
//! accessors, while round trips retain the producer's compact representation.

use litchi_core::{Error, Result};
use std::{num::NonZeroUsize, ops::Range};

/// A typed value stored by an ODF table cell.
#[derive(Clone, Debug, PartialEq)]
pub enum CellValue {
    /// A cell with no `office:value-type` and no displayed content.
    Empty,
    /// ODF string content.
    Text(String),
    /// ODF floating-point content.
    Number(f64),
    /// ODF currency content and its ISO/application currency token.
    Currency { value: f64, currency: String },
    /// ODF percentage content, represented in the stored fractional domain.
    Percentage(f64),
    /// ODF Boolean content.
    Boolean(bool),
    /// ODF date lexical value.
    Date(String),
    /// ODF duration/time lexical value.
    Time(String),
    /// A value type not interpreted by this worksheet facade.
    ///
    /// The type token and value are retained so an unsupported producer value
    /// is not silently converted into ordinary text during an edit.
    Unknown { kind: String, value: Option<String> },
}

impl CellValue {
    pub(crate) fn validate(&self) -> Result<()> {
        match self {
            Self::Empty => Ok(()),
            Self::Text(value) => super::validation::validate_text(value, "cell text"),
            Self::Number(value) | Self::Percentage(value) => {
                if value.is_finite() {
                    Ok(())
                } else {
                    Err(Error::InvalidFormat(
                        "ODS numeric cell values must be finite".to_string(),
                    ))
                }
            },
            Self::Currency { value, currency } => {
                if !value.is_finite() {
                    return Err(Error::InvalidFormat(
                        "ODS currency values must be finite".to_string(),
                    ));
                }
                super::validation::validate_text(currency, "cell currency")?;
                if currency.is_empty() {
                    return Err(Error::InvalidFormat(
                        "ODS currency values require a non-empty currency token".to_string(),
                    ));
                }
                Ok(())
            },
            Self::Boolean(_) => Ok(()),
            Self::Date(value) => super::validation::validate_text(value, "cell date value"),
            Self::Time(value) => super::validation::validate_text(value, "cell time value"),
            Self::Unknown { kind, value } => {
                super::validation::validate_text(kind, "cell value type")?;
                if kind.is_empty() {
                    return Err(Error::InvalidFormat(
                        "unknown ODS cell values require a non-empty type token".to_string(),
                    ));
                }
                if let Some(value) = value {
                    super::validation::validate_text(value, "cell value")?;
                }
                Ok(())
            },
        }
    }
}

/// The merge role carried by one physical cell run.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Merge {
    /// The cell occupies one ordinary grid position.
    None,
    /// The cell is the anchor for a rectangular span.
    Span {
        rows: NonZeroUsize,
        columns: NonZeroUsize,
    },
    /// The cell is a covered position belonging to another span anchor.
    Covered,
}

/// One physical ODF cell run.
#[derive(Clone, Debug, PartialEq)]
pub struct Cell {
    /// Typed stored value.
    pub value: CellValue,
    /// Displayed text, which may differ from the typed value's lexical form.
    pub text: String,
    /// Inert ODF formula token, if present.
    pub formula: Option<String>,
    /// Direct `table:style-name`, if present.
    pub style_name: Option<String>,
    /// Merge/covered-cell role for this physical cell run.
    pub merge: Merge,
    /// Number of adjacent logical cells represented by this physical cell.
    pub repeat: NonZeroUsize,
}

impl Cell {
    /// Create one ordinary cell.
    pub fn new(value: CellValue, text: impl Into<String>) -> Self {
        Self {
            value,
            text: text.into(),
            formula: None,
            style_name: None,
            merge: Merge::None,
            repeat: NonZeroUsize::MIN,
        }
    }

    /// Create a compact repeated cell run.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn repeated(value: CellValue, text: impl Into<String>, repeat: usize) -> Result<Self> {
        let repeat = NonZeroUsize::new(repeat).ok_or_else(|| {
            Error::InvalidFormat("ODS cell repetition must be positive".to_string())
        })?;
        Ok(Self {
            value,
            text: text.into(),
            formula: None,
            style_name: None,
            merge: Merge::None,
            repeat,
        })
    }

    /// Create an empty physical cell.
    #[must_use]
    pub fn empty() -> Self {
        Self::new(CellValue::Empty, "")
    }

    /// Number of logical cells represented by this run.
    #[must_use]
    pub fn repeat(&self) -> usize {
        self.repeat.get()
    }

    /// Set an inert formula after validating its lexical payload.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_formula(&mut self, formula: impl Into<String>) -> Result<()> {
        let formula = formula.into();
        super::validation::validate_text(&formula, "cell formula")?;
        if formula.is_empty() {
            return Err(Error::InvalidFormat(
                "ODS cell formulas must be non-empty".to_string(),
            ));
        }
        self.formula = Some(formula);
        Ok(())
    }

    /// Clear the inert formula while retaining the cached value and text.
    pub fn clear_formula(&mut self) {
        self.formula = None;
    }

    /// Set a direct ODF cell style reference.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_style_name(&mut self, style_name: impl Into<String>) -> Result<()> {
        let style_name = style_name.into();
        super::validation::validate_text(&style_name, "cell style name")?;
        if style_name.is_empty() {
            return Err(Error::InvalidFormat(
                "ODS cell style names must be non-empty".to_string(),
            ));
        }
        self.style_name = Some(style_name);
        Ok(())
    }

    /// Remove the direct cell style reference.
    pub fn clear_style_name(&mut self) {
        self.style_name = None;
    }

    /// Set a rectangular merge span on this cell.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_span(&mut self, rows: usize, columns: usize) -> Result<()> {
        let rows = NonZeroUsize::new(rows).ok_or_else(|| {
            Error::InvalidFormat("ODS merged-cell row span must be positive".to_string())
        })?;
        let columns = NonZeroUsize::new(columns).ok_or_else(|| {
            Error::InvalidFormat("ODS merged-cell column span must be positive".to_string())
        })?;
        self.merge = if rows.get() == 1 && columns.get() == 1 {
            Merge::None
        } else {
            Merge::Span { rows, columns }
        };
        Ok(())
    }

    /// Mark this cell as a covered position or restore it to an ordinary cell.
    pub fn set_covered(&mut self, covered: bool) {
        self.merge = if covered { Merge::Covered } else { Merge::None };
    }

    /// Logical column interval covered by this run when it starts at `start`.
    #[must_use]
    pub fn columns(&self, start: usize) -> Range<usize> {
        start..start.saturating_add(self.repeat())
    }

    pub(crate) fn with_repeat(&self, repeat: usize) -> Result<Self> {
        let repeat = NonZeroUsize::new(repeat).ok_or_else(|| {
            Error::InvalidFormat("ODS cell repetition must be positive".to_string())
        })?;
        let mut value = self.clone();
        value.repeat = repeat;
        Ok(value)
    }

    pub(crate) fn equivalent_run(&self, other: &Self) -> bool {
        self.value == other.value
            && self.text == other.text
            && self.formula == other.formula
            && self.style_name == other.style_name
            && self.merge == other.merge
    }
}

/// One physical ODF row run.
#[derive(Clone, Debug, PartialEq)]
pub struct Row {
    /// Physical cell runs in logical column order.
    pub cells: Vec<Cell>,
    /// Direct `table:style-name`, if present.
    pub style_name: Option<String>,
    /// Direct `table:default-cell-style-name`, if present.
    pub default_cell_style_name: Option<String>,
    /// Number of adjacent logical rows represented by this physical row.
    pub repeat: NonZeroUsize,
}

impl Row {
    /// Create one ordinary row.
    #[must_use]
    pub fn new() -> Self {
        Self {
            cells: Vec::new(),
            style_name: None,
            default_cell_style_name: None,
            repeat: NonZeroUsize::MIN,
        }
    }

    /// Create a compact repeated row run.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn repeated(repeat: usize) -> Result<Self> {
        let repeat = NonZeroUsize::new(repeat).ok_or_else(|| {
            Error::InvalidFormat("ODS row repetition must be positive".to_string())
        })?;
        Ok(Self {
            repeat,
            ..Self::new()
        })
    }

    /// Number of logical rows represented by this run.
    #[must_use]
    pub fn repeat(&self) -> usize {
        self.repeat.get()
    }

    /// Number of logical cells covered by the physical cell runs.
    #[must_use]
    pub fn logical_cell_count(&self) -> usize {
        self.cells
            .iter()
            .fold(0usize, |total, cell| total.saturating_add(cell.repeat()))
    }

    /// Return physical cell runs in logical order.
    #[must_use]
    pub fn cells(&self) -> &[Cell] {
        &self.cells
    }

    /// Return the physical run covering a logical column.
    #[must_use]
    pub fn cell(&self, column: usize) -> Option<&Cell> {
        let mut start = 0usize;
        for cell in &self.cells {
            let end = start.saturating_add(cell.repeat());
            if column < end {
                return Some(cell);
            }
            start = end;
        }
        None
    }

    /// Append one physical cell run.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn push_cell(&mut self, cell: Cell) -> Result<()> {
        cell.validate()?;
        self.cells.push(cell);
        Ok(())
    }

    pub(crate) fn with_repeat(&self, repeat: usize) -> Result<Self> {
        let repeat = NonZeroUsize::new(repeat).ok_or_else(|| {
            Error::InvalidFormat("ODS row repetition must be positive".to_string())
        })?;
        let mut value = self.clone();
        value.repeat = repeat;
        Ok(value)
    }

    pub(crate) fn equivalent_run(&self, other: &Self) -> bool {
        self.style_name == other.style_name
            && self.default_cell_style_name == other.default_cell_style_name
            && self.cells.len() == other.cells.len()
            && self
                .cells
                .iter()
                .zip(&other.cells)
                .all(|(left, right)| left.repeat() == right.repeat() && left.equivalent_run(right))
    }
}

impl Default for Row {
    fn default() -> Self {
        Self::new()
    }
}

/// An immutable lookup result that distinguishes missing cells from stored runs.
#[derive(Clone, Copy, Debug, PartialEq)]
pub enum CellView<'a> {
    /// No physical cell covers the requested coordinate.
    Missing,
    /// A physical cell run covers the requested coordinate.
    Stored(&'a Cell),
}

/// A typed ODS worksheet.
#[derive(Clone, Debug, PartialEq)]
pub struct Sheet {
    /// Exact ODF `table:name`.
    pub name: String,
    /// Physical row runs in document order.
    pub rows: Vec<Row>,
    /// Direct `table:style-name`, if present.
    pub style_name: Option<String>,
}

impl Sheet {
    /// Create an empty named worksheet.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn new(name: impl Into<String>) -> Result<Self> {
        let sheet = Self {
            name: name.into(),
            rows: Vec::new(),
            style_name: None,
        };
        sheet.validate()?;
        Ok(sheet)
    }

    /// Number of logical rows, including repeated rows.
    #[must_use]
    pub fn logical_row_count(&self) -> usize {
        self.rows
            .iter()
            .fold(0usize, |total, row| total.saturating_add(row.repeat()))
    }

    /// Number of logical columns in the widest physical row.
    pub fn logical_column_count(&self) -> usize {
        self.rows
            .iter()
            .map(Row::logical_cell_count)
            .max()
            .unwrap_or(0)
    }

    /// Return the physical row run covering a logical row.
    #[must_use]
    pub fn row(&self, index: usize) -> Option<&Row> {
        let mut start = 0usize;
        for row in &self.rows {
            let end = start.saturating_add(row.repeat());
            if index < end {
                return Some(row);
            }
            start = end;
        }
        None
    }

    /// Return the physical cell run covering a logical coordinate.
    #[must_use]
    pub fn cell(&self, row: usize, column: usize) -> Option<&Cell> {
        self.row(row).and_then(|row| row.cell(column))
    }

    /// Return a lookup view that distinguishes missing and stored coordinates.
    pub fn cell_view(&self, row: usize, column: usize) -> CellView<'_> {
        self.cell(row, column)
            .map_or(CellView::Missing, CellView::Stored)
    }

    /// Set a single logical cell as one atomic model operation.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_cell(&mut self, row: usize, column: usize, cell: Cell) -> Result<()> {
        let mut candidate = self.clone();
        candidate.set_cell_unchecked(row, column, cell)?;
        candidate.validate()?;
        *self = candidate;
        Ok(())
    }

    /// Clear one stored logical cell. Missing coordinates are a no-op.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn clear_cell(&mut self, row: usize, column: usize) -> Result<()> {
        let Some(existing) = self.cell(row, column).cloned() else {
            return Ok(());
        };
        let mut candidate = self.clone();
        let mut replacement = Cell::empty();
        replacement.style_name = existing.style_name.clone();
        candidate.set_cell_unchecked(row, column, replacement)?;
        candidate.validate()?;
        *self = candidate;
        Ok(())
    }

    /// Set a formula while retaining its cached typed value and displayed text.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_formula(
        &mut self,
        row: usize,
        column: usize,
        formula: impl Into<String>,
    ) -> Result<()> {
        let mut cell = self.cell(row, column).cloned().unwrap_or_else(Cell::empty);
        cell.repeat = NonZeroUsize::MIN;
        cell.set_formula(formula)?;
        self.set_cell(row, column, cell)
    }

    /// Set or replace the direct style on one logical cell.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_cell_style(
        &mut self,
        row: usize,
        column: usize,
        style_name: impl Into<String>,
    ) -> Result<()> {
        let mut cell = self.cell(row, column).cloned().unwrap_or_else(Cell::empty);
        cell.repeat = NonZeroUsize::MIN;
        cell.set_style_name(style_name)?;
        self.set_cell(row, column, cell)
    }

    /// Append a physical row run and validate the resulting graph.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn push_row(&mut self, row: Row) -> Result<()> {
        let mut candidate = self.clone();
        candidate.rows.push(row);
        candidate.validate()?;
        *self = candidate;
        Ok(())
    }

    /// Set the direct sheet style reference.
    ///
    /// # Errors
    /// Returns an error when the operation cannot be completed.
    pub fn set_style_name(&mut self, style_name: impl Into<String>) -> Result<()> {
        let style_name = style_name.into();
        super::validation::validate_text(&style_name, "sheet style name")?;
        if style_name.is_empty() {
            return Err(Error::InvalidFormat(
                "ODS sheet style names must be non-empty".to_string(),
            ));
        }
        self.style_name = Some(style_name);
        Ok(())
    }

    /// Remove the direct sheet style reference.
    pub fn clear_style_name(&mut self) {
        self.style_name = None;
    }

    pub(crate) fn validate(&self) -> Result<()> {
        super::validation::validate_sheet(self)
    }

    fn set_cell_unchecked(&mut self, row: usize, column: usize, mut cell: Cell) -> Result<()> {
        if cell.repeat.get() != 1 {
            return Err(Error::InvalidFormat(
                "setting one logical cell requires a non-repeated Cell".to_string(),
            ));
        }
        let row_index = isolate_row(&mut self.rows, row)?;
        let cell_index = isolate_cell(&mut self.rows[row_index].cells, column)?;
        cell.repeat = NonZeroUsize::MIN;
        self.rows[row_index].cells[cell_index] = cell;
        compact_cells(&mut self.rows[row_index].cells)?;
        compact_rows(&mut self.rows)?;
        Ok(())
    }
}

impl Cell {
    pub(crate) fn validate(&self) -> Result<()> {
        super::validation::validate_cell(self)
    }
}

fn isolate_row(rows: &mut Vec<Row>, target: usize) -> Result<usize> {
    let mut start = 0usize;
    for index in 0..rows.len() {
        let count = rows[index].repeat();
        let end = start.checked_add(count).ok_or_else(|| {
            Error::InvalidFormat("ODS row address overflows the logical grid".to_string())
        })?;
        if target < end {
            let offset = target - start;
            let original = rows[index].clone();
            let mut replacement = Vec::with_capacity(3);
            if offset > 0 {
                replacement.push(original.with_repeat(offset)?);
            }
            replacement.push(original.with_repeat(1)?);
            let suffix = count - offset - 1;
            if suffix > 0 {
                replacement.push(original.with_repeat(suffix)?);
            }
            rows.splice(index..=index, replacement);
            return Ok(index + usize::from(offset > 0));
        }
        start = end;
    }

    if target > start {
        rows.push(Row::repeated(target - start)?);
    }
    rows.push(Row::new());
    Ok(rows.len() - 1)
}

fn isolate_cell(cells: &mut Vec<Cell>, target: usize) -> Result<usize> {
    let mut start = 0usize;
    for index in 0..cells.len() {
        let count = cells[index].repeat();
        let end = start.checked_add(count).ok_or_else(|| {
            Error::InvalidFormat("ODS cell address overflows the logical grid".to_string())
        })?;
        if target < end {
            let offset = target - start;
            let original = cells[index].clone();
            let mut replacement = Vec::with_capacity(3);
            if offset > 0 {
                replacement.push(original.with_repeat(offset)?);
            }
            replacement.push(original.with_repeat(1)?);
            let suffix = count - offset - 1;
            if suffix > 0 {
                replacement.push(original.with_repeat(suffix)?);
            }
            cells.splice(index..=index, replacement);
            return Ok(index + usize::from(offset > 0));
        }
        start = end;
    }
    if target > start {
        cells.push(Cell::repeated(CellValue::Empty, "", target - start)?);
    }
    cells.push(Cell::empty());
    Ok(cells.len() - 1)
}

fn compact_cells(cells: &mut Vec<Cell>) -> Result<()> {
    let mut compacted: Vec<Cell> = Vec::with_capacity(cells.len());
    for cell in cells.drain(..) {
        if let Some(previous) = compacted.last_mut()
            && previous.equivalent_run(&cell)
        {
            let repeat = previous
                .repeat()
                .checked_add(cell.repeat())
                .ok_or_else(|| Error::InvalidFormat("ODS cell repetition overflows".to_string()))?;
            previous.repeat = NonZeroUsize::new(repeat).ok_or_else(|| {
                Error::InvalidFormat("ODS cell repetition must be positive".to_string())
            })?;
        } else {
            compacted.push(cell);
        }
    }
    *cells = compacted;
    Ok(())
}

fn compact_rows(rows: &mut Vec<Row>) -> Result<()> {
    let mut compacted: Vec<Row> = Vec::with_capacity(rows.len());
    for row in rows.drain(..) {
        if let Some(previous) = compacted.last_mut()
            && previous.equivalent_run(&row)
        {
            let repeat = previous
                .repeat()
                .checked_add(row.repeat())
                .ok_or_else(|| Error::InvalidFormat("ODS row repetition overflows".to_string()))?;
            previous.repeat = NonZeroUsize::new(repeat).ok_or_else(|| {
                Error::InvalidFormat("ODS row repetition must be positive".to_string())
            })?;
        } else {
            compacted.push(row);
        }
    }
    *rows = compacted;
    Ok(())
}
