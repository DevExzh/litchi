//! Cell data structures for ODS spreadsheets.

use super::{CellAnnotation, CellDetective, CellHyperlink, CellRangeSource, CellTextContent, Row};
use litchi_core::{Result, xml::escape_xml};
use std::{num::NonZeroUsize, ops::Range};

/// Cell data types supported by ODF spreadsheets.
///
/// This enum represents the various data types that can be stored in
/// spreadsheet cells, following the ODF specification.
#[derive(Debug, Clone, PartialEq)]
pub enum CellValue {
    /// Empty cell
    Empty,
    /// Text string
    Text(String),
    /// Numeric value
    Number(f64),
    /// Boolean value
    Boolean(bool),
    /// Date/time value (stored as ISO 8601 string)
    Date(String),
    /// Currency value with currency code
    Currency(f64, String),
    /// Percentage value
    Percentage(f64),
    /// Time duration
    Time(String),
}

/// Merge role of an ODF table cell.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub enum CellMerge {
    /// The cell is not part of a merged range.
    #[default]
    None,
    /// The cell anchors a merged range with the given row and column spans.
    Span {
        rows: NonZeroUsize,
        columns: NonZeroUsize,
    },
    /// The cell is covered by another cell's span.
    Covered,
}

/// Dimensions of an ODF matrix formula result range.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct CellMatrixSpan {
    rows: NonZeroUsize,
    columns: NonZeroUsize,
}

impl CellMatrixSpan {
    /// Create positive matrix-result dimensions.
    pub fn new(rows: usize, columns: usize) -> Result<Self> {
        let rows = NonZeroUsize::new(rows).ok_or_else(|| {
            litchi_core::Error::InvalidFormat("matrix row span must be positive".to_string())
        })?;
        let columns = NonZeroUsize::new(columns).ok_or_else(|| {
            litchi_core::Error::InvalidFormat("matrix column span must be positive".to_string())
        })?;
        Ok(Self { rows, columns })
    }

    /// Number of rows in the matrix result.
    pub fn rows(self) -> usize {
        self.rows.get()
    }

    /// Number of columns in the matrix result.
    pub fn columns(self) -> usize {
        self.columns.get()
    }
}

/// A cell in an ODS spreadsheet.
///
/// Cells contain typed values, optional formulas, and positioning information.
#[derive(Clone, Debug)]
pub struct Cell {
    /// The cell value
    pub value: CellValue,
    /// The raw text content of the cell
    pub text: String,
    /// The formula in the cell (if any), in ODF format
    pub formula: Option<String>,
    /// The optional ODF annotation (comment/note) attached to the cell.
    pub annotation: Option<CellAnnotation>,
    /// Hyperlinks (`text:a`) contained in the cell's text content, in
    /// document order.
    pub hyperlinks: Vec<CellHyperlink>,
    /// Retained mixed paragraph content for parsed rich-text cells.
    pub(crate) rich_text: Option<CellTextContent>,
    /// Optional inert metadata for an externally imported rectangular range.
    pub range_source: Option<CellRangeSource>,
    /// Optional inert formula-auditing highlights and operations.
    pub detective: Option<CellDetective>,
    /// Name of the document-level content validation applied to this cell.
    pub validation_name: Option<String>,
    /// Name of the ODF table-cell style applied directly to this cell.
    pub style_name: Option<String>,
    /// The cell's role in an ODF merged range.
    pub merge: CellMerge,
    /// Matrix formula result dimensions, if declared on this cell.
    pub matrix_span: Option<CellMatrixSpan>,
    /// Legacy ODF `table:protect` state, preserved independently when present.
    pub protect: Option<bool>,
    /// ODF `table:protected` state, preserved independently when present.
    pub protected: Option<bool>,
    /// The row index (0-based)
    pub row: usize,
    /// The column index (0-based)
    pub col: usize,
}

impl Cell {
    /// Create a cell with a typed value, displayed text, and zero-based position.
    pub fn new(value: CellValue, text: impl Into<String>, row: usize, col: usize) -> Self {
        Self {
            value,
            text: text.into(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            merge: CellMerge::None,
            matrix_span: None,
            protect: None,
            protected: None,
            row,
            col,
        }
    }

    /// Get the text content of the cell.
    ///
    /// Returns the displayed text value, which may differ from the
    /// underlying typed value for formatted numbers, dates, etc.
    pub fn text(&self) -> Result<&str> {
        Ok(&self.text)
    }

    /// Get the cell value.
    ///
    /// Returns the typed value stored in the cell.
    pub fn value(&self) -> Result<&CellValue> {
        Ok(&self.value)
    }

    /// Get the numeric value of the cell (if applicable).
    ///
    /// Returns `Some(value)` for Number, Currency, and Percentage types,
    /// `None` for other types.
    pub fn numeric_value(&self) -> Result<Option<f64>> {
        match &self.value {
            CellValue::Number(n) => Ok(Some(*n)),
            CellValue::Currency(n, _) => Ok(Some(*n)),
            CellValue::Percentage(p) => Ok(Some(*p)),
            _ => Ok(None),
        }
    }

    /// Get the formula in the cell.
    ///
    /// Returns the formula string if the cell contains a formula,
    /// None otherwise.
    pub fn formula(&self) -> Result<Option<&str>> {
        Ok(self.formula.as_deref())
    }

    /// Return the cell annotation, if present.
    pub fn annotation(&self) -> Option<&CellAnnotation> {
        self.annotation.as_ref()
    }

    /// Return a mutable cell annotation, if present.
    pub fn annotation_mut(&mut self) -> Option<&mut CellAnnotation> {
        self.annotation.as_mut()
    }

    /// Attach or replace the cell annotation.
    pub fn set_annotation(&mut self, annotation: CellAnnotation) {
        self.annotation = Some(annotation);
    }

    /// Remove and return the cell annotation.
    pub fn take_annotation(&mut self) -> Option<CellAnnotation> {
        self.annotation.take()
    }

    /// Check whether this cell has an annotation.
    pub fn has_annotation(&self) -> bool {
        self.annotation.is_some()
    }

    /// Return all hyperlinks in the cell's text content, in document order.
    pub fn hyperlinks(&self) -> &[CellHyperlink] {
        &self.hyperlinks
    }

    /// Return the first hyperlink in the cell's text content, if any.
    pub fn hyperlink(&self) -> Option<&CellHyperlink> {
        self.hyperlinks.first()
    }

    /// Check whether this cell contains at least one hyperlink.
    pub fn has_hyperlinks(&self) -> bool {
        !self.hyperlinks.is_empty()
    }

    /// Return the structure-preserving mixed paragraph content, when present.
    ///
    /// Parsed spans, fields, whitespace elements, and extension nodes remain
    /// available here and are preserved when the spreadsheet is saved.
    pub fn rich_text(&self) -> Option<&CellTextContent> {
        self.rich_text.as_ref()
    }

    /// Replace this cell's displayed content with one validated hyperlink.
    ///
    /// The operation clears a formula and stores a string value. To retain
    /// existing plain text outside an anchor, use [`Self::add_hyperlink`].
    pub fn set_hyperlink(&mut self, mut hyperlink: CellHyperlink) -> Result<()> {
        if self.merge == CellMerge::Covered {
            return Err(litchi_core::Error::InvalidFormat(
                "cannot author a hyperlink in a covered cell".to_string(),
            ));
        }
        hyperlink.validate()?;
        let text = hyperlink.text.clone();
        hyperlink.set_range(0..text.len());
        let rich_text = CellTextContent::from_hyperlink(&hyperlink)?;
        self.text = text;
        self.value = CellValue::Text(self.text.clone());
        self.formula = None;
        self.hyperlinks = vec![hyperlink];
        self.rich_text = Some(rich_text);
        Ok(())
    }

    /// Add a hyperlink around a UTF-8 byte range of this cell's existing text.
    ///
    /// The supplied link text must exactly match `self.text[range]`. Hyperlink
    /// ranges are kept in document order and may not overlap, so unrelated
    /// text cannot accidentally acquire a target. Hyperlink targets are inert
    /// metadata and are never followed by this crate.
    pub fn add_hyperlink(
        &mut self,
        range: Range<usize>,
        mut hyperlink: CellHyperlink,
    ) -> Result<()> {
        if self.merge == CellMerge::Covered {
            return Err(litchi_core::Error::InvalidFormat(
                "cannot author a hyperlink in a covered cell".to_string(),
            ));
        }
        hyperlink.validate()?;
        let Some(anchor) = self.text.get(range.clone()) else {
            return Err(litchi_core::Error::InvalidFormat(
                "cell hyperlink range is not on a UTF-8 character boundary".to_string(),
            ));
        };
        if hyperlink.text != anchor {
            return Err(litchi_core::Error::InvalidFormat(
                "cell hyperlink text must match its cell range".to_string(),
            ));
        }

        hyperlink.set_range(range.clone());
        let authored = hyperlink.clone();
        let mut hyperlinks = self.hyperlinks.clone();
        hyperlinks.push(hyperlink);
        hyperlinks.sort_by_key(|candidate| {
            let range = candidate.range();
            (range.start, range.start != range.end)
        });
        validate_hyperlink_ranges(&self.text, &hyperlinks)?;
        let mut rich_text = self.rich_text.clone();
        if let Some(content) = rich_text.as_mut() {
            content.wrap_hyperlink(range, &authored)?;
        }
        self.hyperlinks = hyperlinks;
        self.rich_text = rich_text;
        Ok(())
    }

    /// Remove every parsed or authored hyperlink while preserving the cell text.
    pub fn clear_hyperlinks(&mut self) -> Vec<CellHyperlink> {
        if let Some(rich_text) = self.rich_text.as_mut() {
            rich_text.clear_hyperlinks();
        }
        std::mem::take(&mut self.hyperlinks)
    }

    /// Remove and return one hyperlink by its document-order index.
    pub fn remove_hyperlink(&mut self, index: usize) -> Option<CellHyperlink> {
        if index >= self.hyperlinks.len() {
            return None;
        }
        if let Some(rich_text) = self.rich_text.as_mut() {
            debug_assert!(rich_text.remove_hyperlink(index));
        }
        Some(self.hyperlinks.remove(index))
    }

    pub(crate) fn validate_hyperlinks(&self) -> Result<()> {
        validate_hyperlink_ranges(&self.text, &self.hyperlinks)?;
        if let Some(rich_text) = &self.rich_text {
            if rich_text.plain_text() != self.text {
                return Err(litchi_core::Error::InvalidFormat(
                    "cell rich-text tree does not match its displayed text".to_string(),
                ));
            }
            if rich_text.hyperlink_count() != self.hyperlinks.len() {
                return Err(litchi_core::Error::InvalidFormat(
                    "cell rich-text hyperlink count does not match hyperlink metadata".to_string(),
                ));
            }
        }
        Ok(())
    }

    /// Return inert external-range metadata without accessing its URI.
    pub fn range_source(&self) -> Option<&CellRangeSource> {
        self.range_source.as_ref()
    }

    /// Mutably access inert external-range metadata.
    pub fn range_source_mut(&mut self) -> Option<&mut CellRangeSource> {
        self.range_source.as_mut()
    }

    /// Attach or replace inert external-range metadata.
    pub fn set_range_source(&mut self, source: CellRangeSource) {
        self.range_source = Some(source);
    }

    /// Remove and return inert external-range metadata.
    pub fn take_range_source(&mut self) -> Option<CellRangeSource> {
        self.range_source.take()
    }

    /// Return formula-auditing metadata attached to this cell.
    pub fn detective(&self) -> Option<&CellDetective> {
        self.detective.as_ref()
    }

    /// Mutably access formula-auditing metadata attached to this cell.
    pub fn detective_mut(&mut self) -> Option<&mut CellDetective> {
        self.detective.as_mut()
    }

    /// Attach or replace formula-auditing metadata.
    pub fn set_detective(&mut self, detective: CellDetective) {
        self.detective = Some(detective);
    }

    /// Remove and return formula-auditing metadata.
    pub fn take_detective(&mut self) -> Option<CellDetective> {
        self.detective.take()
    }

    /// Return the document-level content-validation name applied to this cell.
    pub fn validation_name(&self) -> Option<&str> {
        self.validation_name.as_deref()
    }

    /// Apply a named document-level content validation to this cell.
    pub fn set_validation_name(&mut self, name: impl Into<String>) {
        self.validation_name = Some(name.into());
    }

    /// Remove the content-validation reference from this cell.
    pub fn clear_validation(&mut self) {
        self.validation_name = None;
    }

    /// Return the directly applied table-cell style name.
    pub fn style_name(&self) -> Option<&str> {
        self.style_name.as_deref()
    }

    /// Apply a named table-cell style.
    pub fn set_style_name(&mut self, name: impl Into<String>) {
        self.style_name = Some(name.into());
    }

    /// Remove the directly applied table-cell style reference.
    pub fn clear_style_name(&mut self) {
        self.style_name = None;
    }

    /// Return this cell's merge role.
    pub fn merge(&self) -> CellMerge {
        self.merge
    }

    /// Return the `(row_span, column_span)` for a merge anchor.
    pub fn span(&self) -> Option<(usize, usize)> {
        match self.merge {
            CellMerge::Span { rows, columns } => Some((rows.get(), columns.get())),
            CellMerge::None | CellMerge::Covered => None,
        }
    }

    /// Set this cell as a merge anchor.
    pub fn set_span(&mut self, rows: usize, columns: usize) -> Result<()> {
        let rows = NonZeroUsize::new(rows).ok_or_else(|| {
            litchi_core::Error::InvalidFormat("cell row span must be positive".to_string())
        })?;
        let columns = NonZeroUsize::new(columns).ok_or_else(|| {
            litchi_core::Error::InvalidFormat("cell column span must be positive".to_string())
        })?;
        self.merge = if rows.get() == 1 && columns.get() == 1 {
            CellMerge::None
        } else {
            CellMerge::Span { rows, columns }
        };
        Ok(())
    }

    /// Mark this cell as covered by a merge anchor.
    pub fn set_covered(&mut self, covered: bool) {
        self.merge = if covered {
            CellMerge::Covered
        } else {
            CellMerge::None
        };
    }

    /// Return this cell's matrix formula result dimensions.
    pub fn matrix_span(&self) -> Option<CellMatrixSpan> {
        self.matrix_span
    }

    /// Set positive matrix formula result dimensions.
    pub fn set_matrix_span(&mut self, rows: usize, columns: usize) -> Result<()> {
        self.matrix_span = Some(CellMatrixSpan::new(rows, columns)?);
        Ok(())
    }

    /// Remove matrix formula result dimensions.
    pub fn clear_matrix_span(&mut self) {
        self.matrix_span = None;
    }

    /// Return the legacy `table:protect` state.
    pub fn protect(&self) -> Option<bool> {
        self.protect
    }

    /// Return the ODF `table:protected` state.
    pub fn protected(&self) -> Option<bool> {
        self.protected
    }

    /// Set both independently representable ODF cell-protection attributes.
    pub fn set_protection(&mut self, protect: Option<bool>, protected: Option<bool>) {
        self.protect = protect;
        self.protected = protected;
    }

    /// Remove both cell-protection attributes.
    pub fn clear_protection(&mut self) {
        self.protect = None;
        self.protected = None;
    }

    /// Parse and get the formula structure.
    ///
    /// Returns the parsed formula if the cell contains a formula,
    /// None otherwise. This provides access to the formula's tokens
    /// and structure for analysis.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = Spreadsheet::open("data.ods")?;
    /// let sheets = spreadsheet.sheets()?;
    /// if let Some(first_sheet) = sheets.first() {
    ///     if let Some(cell) = first_sheet.rows.get(0).and_then(|row| row.cells.get(0)) {
    ///         if let Some(parsed_formula) = cell.parsed_formula()? {
    ///             println!("Formula tokens: {:?}", parsed_formula.tokens);
    ///         }
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn parsed_formula(&self) -> Result<Option<super::formula::Formula>> {
        if let Some(formula_str) = &self.formula {
            let parser = super::formula::FormulaParser::new(formula_str);
            Ok(Some(parser.parse()?))
        } else {
            Ok(None)
        }
    }

    /// Check if the cell has a formula.
    ///
    /// Returns true if the cell contains a formula.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = Spreadsheet::open("data.ods")?;
    /// let sheets = spreadsheet.sheets()?;
    /// if let Some(first_sheet) = sheets.first() {
    ///     if let Some(cell) = first_sheet.rows.get(0).and_then(|row| row.cells.get(0)) {
    ///         if cell.has_formula() {
    ///             println!("Cell A1 contains a formula");
    ///         }
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    #[inline]
    pub fn has_formula(&self) -> bool {
        self.formula.is_some()
    }

    /// Extract cell references from the formula.
    ///
    /// Returns a list of cell references used in the formula.
    /// Returns an empty vector if the cell has no formula.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = Spreadsheet::open("data.ods")?;
    /// let sheets = spreadsheet.sheets()?;
    /// if let Some(first_sheet) = sheets.first() {
    ///     if let Some(cell) = first_sheet.rows.get(0).and_then(|row| row.cells.get(0)) {
    ///         let refs = cell.formula_cell_refs()?;
    ///         println!("Cell references: {:?}", refs);
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn formula_cell_refs(&self) -> Result<Vec<super::formula::CellRef>> {
        if let Some(formula) = self.parsed_formula()? {
            Ok(super::formula::extract_cell_refs(&formula)
                .into_iter()
                .cloned()
                .collect())
        } else {
            Ok(Vec::new())
        }
    }

    /// Extract function names used in the formula.
    ///
    /// Returns a list of function names used in the formula.
    /// Returns an empty vector if the cell has no formula.
    ///
    /// # Examples
    ///
    /// ```no_run
    /// use litchi_odf::Spreadsheet;
    ///
    /// # fn main() -> litchi_core::Result<()> {
    /// let mut spreadsheet = Spreadsheet::open("data.ods")?;
    /// let sheets = spreadsheet.sheets()?;
    /// if let Some(first_sheet) = sheets.first() {
    ///     if let Some(cell) = first_sheet.rows.get(0).and_then(|row| row.cells.get(0)) {
    ///         let funcs = cell.formula_functions()?;
    ///         println!("Functions used: {:?}", funcs);
    ///     }
    /// }
    /// # Ok(())
    /// # }
    /// ```
    pub fn formula_functions(&self) -> Result<Vec<String>> {
        if let Some(formula) = self.parsed_formula()? {
            Ok(super::formula::extract_functions(&formula)
                .into_iter()
                .map(|s| s.to_string())
                .collect())
        } else {
            Ok(Vec::new())
        }
    }

    /// Get the cell coordinates (row, column).
    ///
    /// Returns a tuple of (row_index, column_index), both 0-based.
    pub fn coordinates(&self) -> (usize, usize) {
        (self.row, self.col)
    }

    /// Check if the cell is empty.
    ///
    /// Returns true if the cell value is `Empty`.
    pub fn is_empty(&self) -> bool {
        matches!(self.value, CellValue::Empty)
    }
}

pub(crate) fn merge_cell_range(
    rows: &mut Vec<Row>,
    start_row: usize,
    start_col: usize,
    row_span: usize,
    column_span: usize,
) -> Result<()> {
    let end_row = start_row.checked_add(row_span).ok_or_else(|| {
        litchi_core::Error::InvalidFormat("merged row range overflows address space".to_string())
    })?;
    let end_col = start_col.checked_add(column_span).ok_or_else(|| {
        litchi_core::Error::InvalidFormat("merged column range overflows address space".to_string())
    })?;
    if row_span == 0 || column_span == 0 {
        return Err(litchi_core::Error::InvalidFormat(
            "merged cell ranges must have positive spans".to_string(),
        ));
    }

    if row_span == 1 && column_span == 1 {
        return Err(litchi_core::Error::InvalidFormat(
            "a merged range must cover more than one cell".to_string(),
        ));
    }
    let materialized_cells = row_span.checked_mul(end_col).ok_or_else(|| {
        litchi_core::Error::InvalidFormat("merged range allocation overflows".to_string())
    })?;
    if materialized_cells > 1_048_576 || end_row > 1_048_576 {
        return Err(litchi_core::Error::InvalidFormat(
            "merged range exceeds the expansion safety limit".to_string(),
        ));
    }

    // Validate the existing portion before growing the vectors so errors do
    // not leave a partially expanded sheet behind.
    for (row_index, row) in rows.iter().enumerate().take(end_row).skip(start_row) {
        for (column_index, cell) in row.cells.iter().enumerate().take(end_col).skip(start_col) {
            if cell.merge != CellMerge::None {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "merged range overlaps cell ({row_index}, {column_index})"
                )));
            }
            if (row_index != start_row || column_index != start_col)
                && (!cell.is_empty()
                    || !cell.text.is_empty()
                    || cell.formula.is_some()
                    || cell.annotation.is_some()
                    || cell.range_source.is_some()
                    || cell.detective.is_some())
            {
                return Err(litchi_core::Error::InvalidFormat(format!(
                    "merged range would cover populated cell ({row_index}, {column_index})"
                )));
            }
        }
    }

    while rows.len() < end_row {
        rows.push(Row {
            cells: Vec::new(),
            index: rows.len(),
            style_name: None,
            default_cell_style_name: None,
            visibility: Default::default(),
        });
    }
    for (row_index, row) in rows.iter_mut().enumerate().take(end_row).skip(start_row) {
        while row.cells.len() < end_col {
            row.cells.push(Cell {
                value: CellValue::Empty,
                text: String::new(),
                formula: None,
                annotation: None,
                hyperlinks: Vec::new(),
                rich_text: None,
                range_source: None,
                detective: None,
                validation_name: None,
                style_name: None,
                matrix_span: None,
                merge: CellMerge::None,
                protect: None,
                protected: None,
                row: row_index,
                col: row.cells.len(),
            });
        }
    }

    rows[start_row].cells[start_col].set_span(row_span, column_span)?;
    for (row_index, row) in rows.iter_mut().enumerate().take(end_row).skip(start_row) {
        for (column_index, cell) in row
            .cells
            .iter_mut()
            .enumerate()
            .take(end_col)
            .skip(start_col)
        {
            if row_index != start_row || column_index != start_col {
                cell.set_covered(true);
            }
        }
    }
    Ok(())
}

pub(crate) fn unmerge_cell_range(rows: &mut [Row], start_row: usize, start_col: usize) -> bool {
    let Some((row_span, column_span)) = rows
        .get(start_row)
        .and_then(|row| row.cells.get(start_col))
        .and_then(Cell::span)
    else {
        return false;
    };
    let end_row = start_row.saturating_add(row_span).min(rows.len());
    for (row_index, row) in rows.iter_mut().enumerate().take(end_row).skip(start_row) {
        let end_col = start_col.saturating_add(column_span).min(row.cells.len());
        for (column_index, cell) in row
            .cells
            .iter_mut()
            .enumerate()
            .take(end_col)
            .skip(start_col)
        {
            if (row_index == start_row && column_index == start_col)
                || cell.merge == CellMerge::Covered
            {
                cell.merge = CellMerge::None;
            }
        }
    }
    true
}

pub(crate) fn write_cell_xml(output: &mut String, cell: &Cell) {
    let has_hyperlinks = cell.has_hyperlinks();
    let has_rich_text = cell.rich_text.is_some();
    output.push_str(match cell.merge {
        CellMerge::Covered => "<table:covered-table-cell",
        CellMerge::None | CellMerge::Span { .. } => "<table:table-cell",
    });
    if let CellMerge::Span { rows, columns } = cell.merge {
        output.push_str(" table:number-rows-spanned=\"");
        output.push_str(&rows.get().to_string());
        output.push_str("\" table:number-columns-spanned=\"");
        output.push_str(&columns.get().to_string());
        output.push('"');
    }
    if let Some(matrix_span) = cell.matrix_span {
        output.push_str(" table:number-matrix-rows-spanned=\"");
        output.push_str(&matrix_span.rows().to_string());
        output.push_str("\" table:number-matrix-columns-spanned=\"");
        output.push_str(&matrix_span.columns().to_string());
        output.push('"');
    }
    if let Some(formula) = &cell.formula
        && cell.merge != CellMerge::Covered
    {
        output.push_str(" table:formula=\"");
        output.push_str(&escape_xml(formula));
        output.push('"');
    }
    if let Some(validation_name) = &cell.validation_name {
        output.push_str(" table:content-validation-name=\"");
        output.push_str(&escape_xml(validation_name));
        output.push('"');
    }
    if let Some(style_name) = &cell.style_name {
        output.push_str(" table:style-name=\"");
        output.push_str(&escape_xml(style_name));
        output.push('"');
    }
    if let Some(protect) = cell.protect {
        output.push_str(if protect {
            " table:protect=\"true\""
        } else {
            " table:protect=\"false\""
        });
    }
    if let Some(protected) = cell.protected {
        output.push_str(if protected {
            " table:protected=\"true\""
        } else {
            " table:protected=\"false\""
        });
    }

    if cell.merge == CellMerge::Covered
        && cell.range_source.is_none()
        && cell.annotation.is_none()
        && cell.detective.is_none()
        && !has_hyperlinks
        && !has_rich_text
        && cell.text.is_empty()
    {
        output.push_str("/>");
        return;
    }

    if cell.merge == CellMerge::Covered {
        output.push('>');
        if let Some(source) = &cell.range_source {
            super::source::write_cell_range_source(output, source);
        }
        if let Some(annotation) = &cell.annotation {
            annotation.write_xml(output);
        }
        if let Some(detective) = &cell.detective {
            super::detective::write_detective(output, detective);
        }
        if has_rich_text || has_hyperlinks || !cell.text.is_empty() {
            write_cell_text_content(output, cell);
        }
        output.push_str("</table:covered-table-cell>");
        return;
    }

    match &cell.value {
        CellValue::Text(_) => output.push_str(" office:value-type=\"string\""),
        CellValue::Number(value) => {
            output.push_str(" office:value-type=\"float\" office:value=\"");
            output.push_str(&value.to_string());
            output.push('"');
        },
        CellValue::Currency(value, currency) => {
            output.push_str(" office:value-type=\"currency\" office:value=\"");
            output.push_str(&value.to_string());
            output.push_str("\" office:currency=\"");
            output.push_str(&escape_xml(currency));
            output.push('"');
        },
        CellValue::Percentage(value) => {
            output.push_str(" office:value-type=\"percentage\" office:value=\"");
            output.push_str(&value.to_string());
            output.push('"');
        },
        CellValue::Date(value) => {
            output.push_str(" office:value-type=\"date\" office:date-value=\"");
            output.push_str(&escape_xml(value));
            output.push('"');
        },
        CellValue::Time(value) => {
            output.push_str(" office:value-type=\"time\" office:time-value=\"");
            output.push_str(&escape_xml(value));
            output.push('"');
        },
        CellValue::Boolean(value) => {
            output.push_str(" office:value-type=\"boolean\" office:boolean-value=\"");
            output.push_str(if *value { "true" } else { "false" });
            output.push('"');
        },
        CellValue::Empty
            if cell.annotation.is_none()
                && cell.range_source.is_none()
                && cell.detective.is_none()
                && !has_hyperlinks
                && !has_rich_text =>
        {
            output.push_str("/>");
            return;
        },
        CellValue::Empty => {},
    }

    output.push('>');
    if let Some(source) = &cell.range_source {
        super::source::write_cell_range_source(output, source);
    }
    if let Some(annotation) = &cell.annotation {
        annotation.write_xml(output);
    }
    if let Some(detective) = &cell.detective {
        super::detective::write_detective(output, detective);
    }
    if has_rich_text || has_hyperlinks {
        write_cell_text_content(output, cell);
    } else if !matches!(cell.value, CellValue::Empty) {
        output.push_str("<text:p>");
        output.push_str(&escape_xml(&cell.text));
        output.push_str("</text:p>");
    }
    output.push_str("</table:table-cell>");
}

fn write_cell_text_content(output: &mut String, cell: &Cell) {
    if let Some(rich_text) = &cell.rich_text {
        let mut rich_text = rich_text.clone();
        if rich_text.plain_text() == cell.text
            && rich_text.synchronize_hyperlinks(cell.hyperlinks())
        {
            rich_text.write_xml(output);
            return;
        }
    }
    output.push_str("<text:p>");
    write_cell_text_with_hyperlinks(output, cell);
    output.push_str("</text:p>");
}

fn validate_hyperlink_ranges(text: &str, hyperlinks: &[CellHyperlink]) -> Result<()> {
    let mut previous_end = 0usize;
    for hyperlink in hyperlinks {
        hyperlink.validate()?;
        let range = hyperlink.range();
        if range.start > range.end {
            return Err(litchi_core::Error::InvalidFormat(
                "cell hyperlink range starts after it ends".to_string(),
            ));
        }
        if range.start < previous_end {
            return Err(litchi_core::Error::InvalidFormat(
                "cell hyperlink ranges must be ordered and non-overlapping".to_string(),
            ));
        }
        let Some(anchor) = text.get(range.clone()) else {
            return Err(litchi_core::Error::InvalidFormat(
                "cell hyperlink range is not on a UTF-8 character boundary".to_string(),
            ));
        };
        if anchor != hyperlink.text {
            return Err(litchi_core::Error::InvalidFormat(
                "cell hyperlink text must match its cell range".to_string(),
            ));
        }
        previous_end = range.end;
    }
    Ok(())
}

fn write_cell_text_with_hyperlinks(output: &mut String, cell: &Cell) {
    let mut cursor = 0usize;
    for hyperlink in cell.hyperlinks() {
        let range = hyperlink.range();
        output.push_str(&escape_xml(&cell.text[cursor..range.start]));
        hyperlink.write_xml(output);
        cursor = range.end;
    }
    output.push_str(&escape_xml(&cell.text[cursor..]));
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn cell_new_initializes_optional_metadata() {
        let cell = Cell::new(CellValue::Number(42.0), "42", 3, 4);
        assert_eq!(cell.value, CellValue::Number(42.0));
        assert_eq!(cell.text, "42");
        assert_eq!((cell.row, cell.col), (3, 4));
        assert!(cell.formula.is_none());
        assert_eq!(cell.merge, CellMerge::None);
        assert!(cell.matrix_span.is_none());
        assert_eq!(cell.protect, None);
        assert_eq!(cell.protected, None);
    }

    #[test]
    fn set_hyperlink_replaces_the_complete_displayed_value() {
        let mut cell = Cell::new(CellValue::Number(42.0), "42", 0, 0);
        cell.formula = Some("of:=42".to_string());
        let hyperlink = CellHyperlink::with_text("https://example.test/", "Example").unwrap();
        cell.set_hyperlink(hyperlink.clone()).unwrap();

        assert_eq!(cell.value, CellValue::Text("Example".to_string()));
        assert_eq!(cell.text, "Example");
        assert!(cell.formula.is_none());
        assert_eq!(cell.hyperlinks(), &[hyperlink.clone()]);
        assert_eq!(cell.hyperlinks()[0].range(), 0.."Example".len());
        assert_eq!(cell.clear_hyperlinks(), vec![hyperlink]);
        assert!(!cell.has_hyperlinks());
    }

    #[test]
    fn mixed_text_hyperlinks_are_ordered_validated_and_serialized() {
        let text = "Before link & 第二";
        let mut cell = Cell::new(CellValue::Text(text.to_string()), text, 0, 0);
        let link_start = text.find("link").unwrap();
        let link_range = link_start..link_start + "link".len();
        let second_start = text.find("第二").unwrap();
        let second_range = second_start..text.len();

        cell.add_hyperlink(
            second_range.clone(),
            CellHyperlink::with_text("#Sheet2.B10", "第二").unwrap(),
        )
        .unwrap();
        cell.add_hyperlink(
            link_range.clone(),
            CellHyperlink::with_text("https://example.test/", "link").unwrap(),
        )
        .unwrap();

        assert_eq!(cell.hyperlinks().len(), 2);
        assert_eq!(cell.hyperlinks()[0].range(), link_range);
        assert_eq!(cell.hyperlinks()[1].range(), second_range);

        let mut xml = String::new();
        write_cell_xml(&mut xml, &cell);
        assert!(xml.contains(
            r##"<text:p>Before <text:a xlink:type="simple" xlink:href="https://example.test/">link</text:a> &amp; <text:a xlink:type="simple" xlink:href="#Sheet2.B10">第二</text:a></text:p>"##
        ));

        assert!(
            cell.add_hyperlink(
                link_start + 1..link_start + 2,
                CellHyperlink::with_text("https://overlap.example/", "i").unwrap(),
            )
            .is_err()
        );
        let mut unicode = Cell::new(CellValue::Text("é".to_string()), "é", 0, 0);
        assert!(
            unicode
                .add_hyperlink(
                    1..2,
                    CellHyperlink::with_text("https://example.test/", "").unwrap(),
                )
                .is_err()
        );
    }

    #[test]
    fn test_cell_value_empty() {
        let value = CellValue::Empty;
        assert_eq!(value, CellValue::Empty);
    }

    #[test]
    fn test_cell_value_text() {
        let value = CellValue::Text("Hello".to_string());
        assert_eq!(value, CellValue::Text("Hello".to_string()));
    }

    #[test]
    fn test_cell_value_number() {
        let value = CellValue::Number(42.5);
        assert_eq!(value, CellValue::Number(42.5));
    }

    #[test]
    fn test_cell_value_boolean() {
        let value = CellValue::Boolean(true);
        assert_eq!(value, CellValue::Boolean(true));
    }

    #[test]
    fn test_cell_value_date() {
        let value = CellValue::Date("2024-01-15".to_string());
        assert_eq!(value, CellValue::Date("2024-01-15".to_string()));
    }

    #[test]
    fn test_cell_value_currency() {
        let value = CellValue::Currency(100.0, "USD".to_string());
        match value {
            CellValue::Currency(amount, currency) => {
                assert!((amount - 100.0).abs() < f64::EPSILON);
                assert_eq!(currency, "USD");
            },
            _ => panic!("Expected Currency"),
        }
    }

    #[test]
    fn test_cell_value_percentage() {
        let value = CellValue::Percentage(0.25);
        assert_eq!(value, CellValue::Percentage(0.25));
    }

    #[test]
    fn test_cell_value_time() {
        let value = CellValue::Time("PT1H30M".to_string());
        assert_eq!(value, CellValue::Time("PT1H30M".to_string()));
    }

    #[test]
    fn test_cell_new() {
        let cell = Cell {
            value: CellValue::Empty,
            text: String::new(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert!(cell.is_empty());
        assert_eq!(cell.text, "");
        assert!(cell.formula.is_none());
    }

    #[test]
    fn writes_merge_anchors_and_covered_cells() {
        let mut cell = Cell {
            value: CellValue::Text("anchor".to_string()),
            text: "anchor".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: Some("Merged".to_string()),
            matrix_span: None,
            merge: CellMerge::None,
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert!(cell.set_span(0, 2).is_err());
        cell.set_span(2, 3).unwrap();
        cell.set_matrix_span(4, 2).unwrap();
        let mut xml = String::new();
        write_cell_xml(&mut xml, &cell);
        assert!(xml.starts_with(
            r#"<table:table-cell table:number-rows-spanned="2" table:number-columns-spanned="3""#
        ));
        assert!(xml.contains(r#"table:number-matrix-rows-spanned="4""#));

        cell.set_covered(true);
        xml.clear();
        write_cell_xml(&mut xml, &cell);
        assert_eq!(
            xml,
            r#"<table:covered-table-cell table:number-matrix-rows-spanned="4" table:number-matrix-columns-spanned="2" table:style-name="Merged"><text:p>anchor</text:p></table:covered-table-cell>"#
        );
    }

    #[test]
    fn writes_range_sources_before_cell_content_and_inside_covered_cells() {
        let source = CellRangeSource::new("Named", "source.ods", 2, 3).unwrap();
        let mut cell = Cell::new(CellValue::Text("value".to_string()), "value", 0, 0);
        cell.set_range_source(source.clone());

        let mut xml = String::new();
        write_cell_xml(&mut xml, &cell);
        let source_position = xml.find("<table:cell-range-source").unwrap();
        let text_position = xml.find("<text:p>").unwrap();
        assert!(source_position < text_position);
        assert!(!xml.contains("xmlns:"));

        cell.value = CellValue::Empty;
        cell.text.clear();
        cell.set_covered(true);
        xml.clear();
        write_cell_xml(&mut xml, &cell);
        assert!(xml.starts_with("<table:covered-table-cell>"));
        assert!(xml.contains("<table:cell-range-source"));
        assert!(xml.ends_with("</table:covered-table-cell>"));
        assert_eq!(cell.take_range_source(), Some(source));
    }

    #[test]
    fn leaves_an_uncached_formula_without_a_fabricated_zero_result() {
        let mut cell = Cell::new(CellValue::Empty, "", 0, 0);
        cell.formula = Some("of:=SUM([.A1:.B1])".to_string());

        let mut xml = String::new();
        write_cell_xml(&mut xml, &cell);

        assert!(xml.contains("table:formula=\"of:=SUM([.A1:.B1])\""));
        assert!(!xml.contains("office:value=\"0\""));
        assert!(!xml.contains("<text:p>0</text:p>"));
    }

    #[test]
    fn test_cell_text() {
        let cell = Cell {
            value: CellValue::Text("Hello".to_string()),
            text: "Hello".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.text().unwrap(), "Hello");
    }

    #[test]
    fn test_cell_value() {
        let cell = Cell {
            value: CellValue::Number(42.0),
            text: "42".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        match cell.value().unwrap() {
            CellValue::Number(n) => assert!((n - 42.0).abs() < f64::EPSILON),
            _ => panic!("Expected Number"),
        }
    }

    #[test]
    fn test_cell_numeric_value() {
        let cell = Cell {
            value: CellValue::Number(42.0),
            text: "42".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.numeric_value().unwrap(), Some(42.0));

        let cell = Cell {
            value: CellValue::Currency(100.0, "USD".to_string()),
            text: "$100".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.numeric_value().unwrap(), Some(100.0));

        let cell = Cell {
            value: CellValue::Percentage(0.5),
            text: "50%".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.numeric_value().unwrap(), Some(0.5));

        let cell = Cell {
            value: CellValue::Text("Hello".to_string()),
            text: "Hello".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.numeric_value().unwrap(), None);
    }

    #[test]
    fn test_cell_formula() {
        let cell = Cell {
            value: CellValue::Number(42.0),
            text: "42".to_string(),
            formula: Some("=A1+B1".to_string()),
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.formula().unwrap(), Some("=A1+B1"));
    }

    #[test]
    fn test_cell_no_formula() {
        let cell = Cell {
            value: CellValue::Text("Hello".to_string()),
            text: "Hello".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert_eq!(cell.formula().unwrap(), None);
    }

    #[test]
    fn test_cell_has_formula() {
        let cell_with = Cell {
            value: CellValue::Number(42.0),
            text: "42".to_string(),
            formula: Some("=A1".to_string()),
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert!(cell_with.has_formula());

        let cell_without = Cell {
            value: CellValue::Text("Hello".to_string()),
            text: "Hello".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert!(!cell_without.has_formula());
    }

    #[test]
    fn test_cell_coordinates() {
        let cell = Cell {
            value: CellValue::Empty,
            text: String::new(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 5,
            col: 10,
        };
        assert_eq!(cell.coordinates(), (5, 10));
    }

    #[test]
    fn test_cell_is_empty() {
        let empty_cell = Cell {
            value: CellValue::Empty,
            text: String::new(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert!(empty_cell.is_empty());

        let text_cell = Cell {
            value: CellValue::Text("Hello".to_string()),
            text: "Hello".to_string(),
            formula: None,
            annotation: None,
            hyperlinks: Vec::new(),
            rich_text: None,
            range_source: None,
            detective: None,
            validation_name: None,
            style_name: None,
            matrix_span: None,
            merge: Default::default(),
            protect: None,
            protected: None,
            row: 0,
            col: 0,
        };
        assert!(!text_cell.is_empty());
    }

    #[test]
    fn test_cell_equality() {
        let cell1 = CellValue::Number(42.0);
        let cell2 = CellValue::Number(42.0);
        let cell3 = CellValue::Number(43.0);

        assert_eq!(cell1, cell2);
        assert_ne!(cell1, cell3);
    }

    #[test]
    fn test_cell_clone() {
        let cell = CellValue::Text("Hello".to_string());
        let cloned = cell.clone();
        assert_eq!(cell, cloned);
    }
}
