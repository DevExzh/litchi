use std::collections::{HashMap, HashSet};

use super::comment::WritableComment;
use super::shape::ShapeWrite;
use super::shape_group::ShapeGroupWrite;
use super::{
    CellValue, ConditionalFormat, ConditionalFormat12Group, ConditionalFormatGroup,
    ConditionalFormatRange, ConditionalFormatRule, DataValidation, DataValidationBiffPayload,
    DataValidationOptions, DataValidationRange, DataValidationTableOptions, PageSetupOptions,
};
use crate::writer::biff::AutoFilterConditionWrite;
use crate::writer::formula::{FormulaTokenizer, compile_array_formula, encode_ptg_tokens};
use crate::{Error, Result};

use super::model::Writer;

use crate::formula_metadata::{Cell as FormulaCell, Owner, Range as FormulaRange, array};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum PivotCellXfRole {
    HeaderAccent,
    HeaderPlain,
    RowLabel,
    Value,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
pub(super) struct CellPos {
    row: u16,
    col: u8,
}

impl CellPos {
    pub(super) fn try_new(row: u32, col: u16) -> Result<Self> {
        let invalid = || {
            Error::InvalidCellReference(format!(
                "row {row}, column {col} is outside the BIFF8 grid"
            ))
        };
        let row = u16::try_from(row).map_err(|_error| invalid())?;
        let col = u8::try_from(col).map_err(|_error| invalid())?;
        Ok(Self { row, col })
    }

    pub(super) const fn row(self) -> u16 {
        self.row
    }

    pub(super) const fn col(self) -> u8 {
        self.col
    }
}

#[derive(Debug, Clone)]
pub(super) struct WritableCell {
    pos: CellPos,
    /// Cell value
    pub value: CellValue,
    pub formula_metadata: Option<crate::FormulaMetadata>,
    pub format_idx: u16,
    pub pivot_xf_role: Option<PivotCellXfRole>,
}

impl WritableCell {
    pub(super) const fn new(
        pos: CellPos,
        value: CellValue,
        format_idx: u16,
        pivot_xf_role: Option<PivotCellXfRole>,
    ) -> Self {
        Self {
            pos,
            value,
            formula_metadata: None,
            format_idx,
            pivot_xf_role,
        }
    }

    pub(super) fn with_formula_metadata(
        mut self,
        formula_metadata: Option<crate::FormulaMetadata>,
    ) -> Self {
        self.formula_metadata = formula_metadata;
        self
    }

    pub(super) const fn row(&self) -> u16 {
        self.pos.row()
    }

    pub(super) const fn col(&self) -> u8 {
        self.pos.col()
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct MergedRange {
    first: CellPos,
    last: CellPos,
}

impl MergedRange {
    pub(super) fn try_new(
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> Result<Self> {
        let first = CellPos::try_new(first_row, first_col)?;
        let last = CellPos::try_new(last_row, last_col)?;
        if first.row() > last.row() || first.col() > last.col() {
            return Err(Error::InvalidCellReference(format!(
                "range ({first_row}, {first_col})..=({last_row}, {last_col}) is reversed"
            )));
        }
        Ok(Self { first, last })
    }

    pub(super) const fn fields(self) -> (u16, u16, u8, u8) {
        (
            self.first.row(),
            self.last.row(),
            self.first.col(),
            self.last.col(),
        )
    }

    pub(super) const fn overlaps(self, other: Self) -> bool {
        self.first.row() <= other.last.row()
            && other.first.row() <= self.last.row()
            && self.first.col() <= other.last.col()
            && other.first.col() <= self.last.col()
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct HorizontalPageBreak {
    row: u16,
    col_start: u16,
    col_end: u16,
}

impl HorizontalPageBreak {
    pub(super) fn try_new(row: u32, col_start: u16, col_end: u16) -> Result<Self> {
        let row = u16::try_from(row).map_err(|_error| {
            Error::InvalidCellReference(format!(
                "horizontal page-break row {row} is outside the BIFF8 grid"
            ))
        })?;
        if col_end <= col_start || col_end > 16_383 {
            return Err(Error::InvalidCellReference(format!(
                "horizontal page-break columns {col_start}..={col_end} are outside the BIFF8 page-break bounds"
            )));
        }
        Ok(Self {
            row,
            col_start,
            col_end,
        })
    }

    pub(super) const fn overlaps(self, other: Self) -> bool {
        self.row == other.row && self.col_start <= other.col_end && other.col_start <= self.col_end
    }

    pub(super) const fn fields(self) -> (u16, u16, u16) {
        (self.row, self.col_start, self.col_end)
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) struct VerticalPageBreak {
    col: u8,
    row_start: u16,
    row_end: u16,
}

impl VerticalPageBreak {
    pub(super) fn try_new(col: u16, row_start: u32, row_end: u32) -> Result<Self> {
        let invalid = || {
            Error::InvalidCellReference(format!(
                "vertical page-break column {col}, rows {row_start}..={row_end} are outside the BIFF8 grid"
            ))
        };
        let col = u8::try_from(col).map_err(|_error| invalid())?;
        let row_start = u16::try_from(row_start).map_err(|_error| invalid())?;
        let row_end = u16::try_from(row_end).map_err(|_error| invalid())?;
        if row_end <= row_start {
            return Err(invalid());
        }
        Ok(Self {
            col,
            row_start,
            row_end,
        })
    }

    pub(super) const fn overlaps(self, other: Self) -> bool {
        self.col == other.col && self.row_start <= other.row_end && other.row_start <= self.row_end
    }

    pub(super) const fn fields(self) -> (u16, u16, u16) {
        (self.col as u16, self.row_start, self.row_end)
    }
}

#[derive(Debug, Clone, Copy)]
pub(super) struct AutoFilterRange {
    pub first_row: u32,
    pub last_row: u32,
    pub first_col: u16,
    pub last_col: u16,
}

#[derive(Debug, Clone, Copy)]
pub(super) struct SheetProtection {
    pub protect_objects: bool,
    pub protect_scenarios: bool,
    pub password_hash: Option<u16>,
}

#[derive(Debug, Clone)]
pub(super) struct WritableDataValidation {
    pub validation: DataValidation,
    pub payload: DataValidationBiffPayload,
    pub ranges: Vec<DataValidationRange>,
    pub options: DataValidationOptions,
}

/// Hyperlink target within a worksheet.
#[derive(Debug, Clone)]
pub(super) struct Hyperlink {
    /// First row (0-based) of the hyperlink range.
    pub first_row: u32,
    /// Last row (0-based) of the hyperlink range.
    pub last_row: u32,
    /// First column (0-based) of the hyperlink range.
    pub first_col: u16,
    /// Last column (0-based) of the hyperlink range.
    pub last_col: u16,
    /// Raw hyperlink target string.
    pub url: String,
}

/// Represents a worksheet in the writer
#[derive(Debug)]
pub(super) struct WritableWorksheet {
    /// Worksheet name
    pub name: String,
    /// Cells to write (indexed by (row, col))
    pub cells: HashMap<(u32, u16), WritableCell>,
    /// First used row
    pub first_row: u32,
    /// Last used row (exclusive)
    pub last_row: u32,
    /// First used column
    pub first_col: u16,
    /// Last used column (exclusive)
    pub last_col: u16,
    /// Per-column widths in 1/256 character units (BIFF8 COLINFO).
    pub column_widths: HashMap<u16, u16>,
    /// Hidden columns (0-based indices).
    pub hidden_columns: HashSet<u16>,
    /// Per-row heights in 1/20 point units (BIFF8 ROW).
    pub row_heights: HashMap<u32, u16>,
    /// Hidden rows (0-based indices).
    pub hidden_rows: HashSet<u32>,
    pub merged_ranges: Vec<MergedRange>,
    pub data_validations: Vec<WritableDataValidation>,
    pub data_validation_table_options: Option<DataValidationTableOptions>,
    pub conditional_formats: Vec<ConditionalFormatGroup>,
    pub conditional_formats12: Vec<ConditionalFormat12Group>,
    pub view: crate::writer::view::View,
    pub sheet_protection: Option<SheetProtection>,
    pub sheet_layout: super::WorksheetLayoutOptions,
    pub page_setup: Option<PageSetupOptions>,
    pub horizontal_page_breaks: Vec<HorizontalPageBreak>,
    pub vertical_page_breaks: Vec<VerticalPageBreak>,
    pub auto_filter: Option<AutoFilterRange>,
    /// Cell or range hyperlinks stored for this worksheet.
    pub hyperlinks: Vec<Hyperlink>,
    pub comments: Vec<WritableComment>,
    pub shapes: Vec<ShapeWrite>,
    pub shape_groups: Vec<ShapeGroupWrite>,
    /// Per-column `AutoFilter` conditions.
    pub auto_filter_columns: Vec<AutoFilterColumnDef>,
    /// Sort configuration.
    pub sort_config: Option<SortConfig>,
    /// Extended sort metadata and conditions.
    pub sort_data: Option<crate::sort_data::Config>,
    /// Pivot tables to write.
    pub pivot_tables: Vec<WritablePivotTable>,
    pub formulas_pending_recalculation: bool,
    pub scenario_manager: Option<crate::scenario::ScenarioManager>,
    pub vba_code_name: Option<String>,
    pub consolidation: Option<crate::consolidation::Consolidation>,
    pub list_objects: Vec<crate::ListObject>,
    /// Sheet tab color as a palette index (SHEETEXT `icvPlain`).
    pub tab_color: Option<u8>,
    /// What-if data tables: anchor formula cell and typed TABLE record.
    pub data_tables: Vec<(u32, u16, crate::DataTable)>,
    /// Default phonetic format and visible ranges (PHONETICINFO).
    pub phonetic_info: Option<crate::PhoneticInfo>,
    /// Web pages published from this sheet (`WebPub` records).
    pub web_publications: Vec<crate::WebPub>,
}

/// A column-level `AutoFilter` condition for the writer.
#[derive(Debug, Clone)]
pub(super) struct AutoFilterColumnDef {
    /// Column index within the filter range (0-based relative to filter start).
    pub column_index: u16,
    /// Join logic: true = OR, false = AND.
    pub join_or: bool,
    /// First condition.
    pub condition1: AutoFilterConditionWrite,
    /// Second condition.
    pub condition2: AutoFilterConditionWrite,
}

/// Sort configuration for the writer.
#[derive(Debug, Clone)]
pub(super) struct SortConfig {
    pub case_sensitive: bool,
    /// true = sort by columns (left-to-right), false = by rows (top-to-bottom)
    pub sort_by_columns: bool,
    /// Up to 3 sort keys: (`column_index`, descending).
    pub keys: Vec<(u16, bool)>,
}

impl WritableWorksheet {
    pub(super) fn new(name: String) -> Self {
        Self {
            name,
            cells: HashMap::new(),
            first_row: 0,
            last_row: 0,
            first_col: 0,
            last_col: 0,
            column_widths: HashMap::new(),
            hidden_columns: HashSet::new(),
            row_heights: HashMap::new(),
            hidden_rows: HashSet::new(),
            merged_ranges: Vec::new(),
            data_validations: Vec::new(),
            data_validation_table_options: None,
            conditional_formats: Vec::new(),
            conditional_formats12: Vec::new(),
            view: crate::writer::view::View::default(),
            sheet_protection: None,
            sheet_layout: super::WorksheetLayoutOptions::default(),
            page_setup: None,
            horizontal_page_breaks: Vec::new(),
            vertical_page_breaks: Vec::new(),
            auto_filter: None,
            hyperlinks: Vec::new(),
            comments: Vec::new(),
            shapes: Vec::new(),
            shape_groups: Vec::new(),
            auto_filter_columns: Vec::new(),
            sort_config: None,
            sort_data: None,
            pivot_tables: Vec::new(),
            formulas_pending_recalculation: false,
            scenario_manager: None,
            vba_code_name: None,
            consolidation: None,
            list_objects: Vec::new(),
            tab_color: None,
            data_tables: Vec::new(),
            phonetic_info: None,
            web_publications: Vec::new(),
        }
    }

    pub(super) fn add_cell(&mut self, cell: WritableCell) {
        let row = u32::from(cell.row());
        let col = u16::from(cell.col());

        // Update dimensions
        if self.cells.is_empty() {
            self.first_row = row;
            self.last_row = row + 1;
            self.first_col = col;
            self.last_col = col + 1;
        } else {
            self.first_row = self.first_row.min(row);
            self.last_row = self.last_row.max(row + 1);
            self.first_col = self.first_col.min(col);
            self.last_col = self.last_col.max(col + 1);
        }

        self.cells.insert((row, col), cell);
    }

    pub(super) fn include_list_object_range(&mut self, range: crate::ListObjectRange) {
        if self.cells.is_empty() && self.list_objects.is_empty() {
            self.first_row = u32::from(range.first_row());
            self.last_row = u32::from(range.last_row()) + 1;
            self.first_col = range.first_column();
            self.last_col = range.last_column() + 1;
        } else {
            self.first_row = self.first_row.min(u32::from(range.first_row()));
            self.last_row = self.last_row.max(u32::from(range.last_row()) + 1);
            self.first_col = self.first_col.min(range.first_column());
            self.last_col = self.last_col.max(range.last_column() + 1);
        }
    }

    pub(super) fn add_merged_range(&mut self, range: MergedRange) {
        self.merged_ranges.push(range);
    }

    pub(super) fn add_data_validation(
        &mut self,
        validation: DataValidation,
        payload: DataValidationBiffPayload,
        ranges: Vec<DataValidationRange>,
        options: DataValidationOptions,
    ) {
        self.data_validations.push(WritableDataValidation {
            validation,
            payload,
            ranges,
            options,
        });
    }

    pub(super) fn add_conditional_format(&mut self, cf: ConditionalFormat) {
        self.conditional_formats.push(ConditionalFormatGroup {
            ranges: vec![ConditionalFormatRange {
                first_row: cf.first_row,
                last_row: cf.last_row,
                first_col: cf.first_col,
                last_col: cf.last_col,
            }],
            rules: vec![ConditionalFormatRule {
                format_type: cf.format_type,
                pattern: cf.pattern,
            }],
        });
    }
    pub(super) fn add_conditional_format_group(&mut self, group: ConditionalFormatGroup) {
        self.conditional_formats.push(group);
    }
    pub(super) fn add_conditional_format12_group(&mut self, group: ConditionalFormat12Group) {
        self.conditional_formats12.push(group);
    }

    pub(super) fn set_freeze_panes(&mut self, panes: crate::writer::FrozenPanes) -> Result<()> {
        self.view
            .set_frozen(panes.rows().index(), panes.columns().index())
    }

    pub(super) fn clear_freeze_panes(&mut self) {
        self.view.clear_pane();
    }

    pub(super) fn put_scale(
        &mut self,
        scale: Option<crate::writer::view::Scale>,
    ) -> Option<crate::writer::view::Scale> {
        self.view.put_scale(scale)
    }

    pub(super) fn set_column_width(&mut self, column: crate::writer::Column, width: u16) {
        self.column_widths.insert(u16::from(column.index()), width);
    }

    pub(super) fn hide_column(&mut self, column: crate::writer::Column) {
        self.hidden_columns.insert(u16::from(column.index()));
    }

    pub(super) fn add_hyperlink(&mut self, hyperlink: Hyperlink) {
        self.hyperlinks.push(hyperlink);
    }

    pub(super) fn add_comment(&mut self, comment: WritableComment) -> Result<()> {
        self.comments
            .try_reserve(1)
            .map_err(|_error| Error::Allocation("reserving worksheet comment storage"))?;
        self.comments.push(comment);
        Ok(())
    }

    pub(super) fn show_column(&mut self, column: crate::writer::Column) {
        self.hidden_columns.remove(&u16::from(column.index()));
    }

    pub(super) fn set_row_height(&mut self, row: crate::writer::Row, height: u16) {
        self.row_heights.insert(u32::from(row.index()), height);
    }

    pub(super) fn hide_row(&mut self, row: crate::writer::Row) {
        self.hidden_rows.insert(u32::from(row.index()));
    }

    pub(super) fn show_row(&mut self, row: crate::writer::Row) {
        self.hidden_rows.remove(&u32::from(row.index()));
    }

    pub(super) fn add_auto_filter_column(&mut self, def: AutoFilterColumnDef) {
        self.auto_filter_columns.push(def);
    }

    pub(super) fn set_sort_config(&mut self, config: SortConfig) {
        self.sort_config = Some(config);
    }

    pub(super) fn put_sort(
        &mut self,
        sort: crate::sort_data::Config,
    ) -> Option<crate::sort_data::Config> {
        self.sort_data.replace(sort)
    }

    pub(super) fn remove_sort(&mut self) -> Option<crate::sort_data::Config> {
        self.sort_data.take()
    }

    pub(super) fn add_pivot_table(&mut self, pt: WritablePivotTable) {
        // Expand worksheet dimensions to encompass the pivot table output
        // range.  Excel validates that the DIMENSIONS record covers the
        // SXVIEW output area; a mismatch causes a "corrupt file" repair
        // dialog.
        let pt_first_row = u32::from(pt.first_row);
        let pt_last_row_excl = u32::from(pt.last_row) + 1; // DIMENSIONS uses exclusive end
        let pt_first_col = pt.first_col;
        let pt_last_col_excl = pt.last_col + 1;

        if self.cells.is_empty() && self.pivot_tables.is_empty() {
            self.first_row = pt_first_row;
            self.last_row = pt_last_row_excl;
            self.first_col = pt_first_col;
            self.last_col = pt_last_col_excl;
        } else {
            self.first_row = self.first_row.min(pt_first_row);
            self.last_row = self.last_row.max(pt_last_row_excl);
            self.first_col = self.first_col.min(pt_first_col);
            self.last_col = self.last_col.max(pt_last_col_excl);
        }

        self.pivot_tables.push(pt);
    }
}

impl Writer {
    /// Write one inert BIFF8 shared formula and its participating cells.
    ///
    /// `range` becomes the `ShrFmla` `RefU`. `participants` is the complete
    /// participating-cell set when non-empty and therefore must include
    /// `anchor`; an empty slice means that only the anchor participates. The
    /// anchor must not follow any participating cell in worksheet row-major
    /// order because the CELLTABLE grammar requires the anchor Formula and its
    /// `ShrFmla` to precede the other Formula records.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_shared_formula(
        &mut self,
        sheet: usize,
        range: FormulaRange,
        anchor: FormulaCell,
        formula: &str,
        participants: &[FormulaCell],
    ) -> Result<()> {
        self.write_shared_formula_with_format(sheet, range, anchor, formula, 0, participants)
    }

    /// Write a formatted BIFF8 shared formula.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_shared_formula_with_format(
        &mut self,
        sheet: usize,
        range: FormulaRange,
        anchor: FormulaCell,
        formula: &str,
        format_id: u16,
        participants: &[FormulaCell],
    ) -> Result<()> {
        if self.fmt.get_format(format_id).is_none() {
            return Err(Error::InvalidFormat(format_id));
        }

        let mut participant_cells = if participants.is_empty() {
            vec![anchor]
        } else {
            participants.to_vec()
        };
        participant_cells.sort_unstable();

        let expression = formula.strip_prefix('=').unwrap_or(formula);
        let tokens = FormulaTokenizer::new().tokenize(expression)?;
        let encoded = encode_ptg_tokens(&tokens);
        let mut owner = Owner::new(range, anchor, &encoded)?;
        if !participants.is_empty() {
            owner = owner.with_participants(&participant_cells)?;
        }
        let metadata = crate::FormulaMetadata::new()
            .with_always_calculate(true)
            .with_shared(owner);

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;

        for participant in &participant_cells {
            let key = (u32::from(participant.row()), u16::from(participant.col()));
            if worksheet.cells.contains_key(&key) {
                return Err(Error::InvalidData(format!(
                    "shared-formula participant cell ({}, {}) is already occupied",
                    participant.row(),
                    participant.col()
                )));
            }
            if worksheet
                .data_tables
                .iter()
                .any(|(row, col, _)| (*row, *col) == key)
            {
                return Err(Error::InvalidData(format!(
                    "shared-formula participant cell ({}, {}) overlaps a data-table anchor",
                    participant.row(),
                    participant.col()
                )));
            }
        }

        for participant in &participant_cells {
            let pos = CellPos::try_new(u32::from(participant.row()), u16::from(participant.col()))?;
            // The BIFF writer replaces participating-cell tokens with
            // PtgExp, but retaining the expression keeps staging valid and
            // lets the existing tokenizer validate the formula.
            let value = CellValue::Formula(formula.to_string());
            worksheet.add_cell(
                WritableCell::new(pos, value, format_id, None)
                    .with_formula_metadata(Some(metadata.clone())),
            );
        }

        Ok(())
    }

    /// Write one inert BIFF8 array formula over a complete rectangle.
    ///
    /// The upper-left cell owns the following `Array` record. Every cell in
    /// `range` is staged as a Formula containing the owner's exact `PtgExp`;
    /// sparse participation is intentionally not supported. The expression
    /// is compiled by the bounded, non-executing array compiler and cached
    /// results remain the canonical BIFF8 Empty value. Authored Array records
    /// request recalculation by default, but this crate never evaluates them.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_array_formula(
        &mut self,
        sheet: usize,
        range: FormulaRange,
        formula: &str,
    ) -> Result<()> {
        self.write_array_formula_with_format_and_limits(
            sheet,
            range,
            formula,
            0,
            array::Limits::default(),
        )
    }

    /// Write an inert BIFF8 array formula with explicit resource limits.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_array_formula_with_limits(
        &mut self,
        sheet: usize,
        range: FormulaRange,
        formula: &str,
        limits: array::Limits,
    ) -> Result<()> {
        self.write_array_formula_with_format_and_limits(sheet, range, formula, 0, limits)
    }

    /// Write a formatted, inert BIFF8 array formula over a complete rectangle.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_array_formula_with_format(
        &mut self,
        sheet: usize,
        range: FormulaRange,
        formula: &str,
        format_id: u16,
    ) -> Result<()> {
        self.write_array_formula_with_format_and_limits(
            sheet,
            range,
            formula,
            format_id,
            array::Limits::default(),
        )
    }

    /// Write a formatted, inert BIFF8 array formula with explicit limits.
    /// # Errors
    ///
    /// Returns an error if validation, decoding, encoding, or the requested operation fails.
    pub fn write_array_formula_with_format_and_limits(
        &mut self,
        sheet: usize,
        range: FormulaRange,
        formula: &str,
        format_id: u16,
        limits: array::Limits,
    ) -> Result<()> {
        if self.fmt.get_format(format_id).is_none() {
            return Err(Error::InvalidFormat(format_id));
        }

        let worksheet = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        let first = range.first();
        let last = range.last();
        let row_count = usize::from(last.row() - first.row()) + 1;
        let col_count = usize::from(last.col() - first.col()) + 1;
        let cell_count = row_count.checked_mul(col_count).ok_or_else(|| {
            Error::InvalidFormula("array-formula rectangle cardinality overflow".to_string())
        })?;
        if cell_count > limits.max_cells() {
            return Err(Error::InvalidFormula(format!(
                "array-formula rectangle contains {cell_count} cells, exceeding the limit of {}",
                limits.max_cells()
            )));
        }
        let tokens = compile_array_formula(formula, limits)?;

        let in_range = |row: u32, col: u16| {
            (u32::from(first.row())..=u32::from(last.row())).contains(&row)
                && (u16::from(first.col())..=u16::from(last.col())).contains(&col)
        };
        if let Some(&(row, col)) = worksheet
            .cells
            .keys()
            .find(|&&(row, col)| in_range(row, col))
        {
            return Err(Error::InvalidData(format!(
                "array-formula cell ({row}, {col}) is already occupied"
            )));
        }
        for &(anchor_row, anchor_col, table) in &worksheet.data_tables {
            if in_range(anchor_row, anchor_col) {
                return Err(Error::InvalidData(format!(
                    "array-formula range overlaps data-table anchor ({anchor_row}, {anchor_col})"
                )));
            }
            let table_range = table.range();
            let overlaps_rows = u32::from(first.row()) <= u32::from(table_range.last_row())
                && u32::from(table_range.first_row()) <= u32::from(last.row());
            let overlaps_cols =
                first.col() <= table_range.last_col() && table_range.first_col() <= last.col();
            if overlaps_rows && overlaps_cols {
                return Err(Error::InvalidData(
                    "array-formula range overlaps a data-table formula group".to_string(),
                ));
            }
        }

        let owner = if limits == array::Limits::default() {
            array::Owner::from_compiled(range, tokens)?
        } else {
            array::Owner::from_compiled_with_limits(range, tokens, limits)?
        };
        let metadata = crate::FormulaMetadata::new()
            .with_always_calculate(true)
            .with_array(owner);
        let mut candidates = Vec::new();
        candidates
            .try_reserve_exact(cell_count)
            .map_err(|_error| Error::Allocation("reserving array-formula cells"))?;
        for row in first.row()..=last.row() {
            for col in first.col()..=last.col() {
                candidates.push(
                    WritableCell::new(
                        CellPos { row, col },
                        CellValue::Formula(String::new()),
                        format_id,
                        None,
                    )
                    .with_formula_metadata(Some(metadata.clone())),
                );
            }
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| Error::WorksheetNotFound(format!("Sheet {sheet}")))?;
        worksheet
            .cells
            .try_reserve(cell_count)
            .map_err(|_error| Error::Allocation("reserving array-formula worksheet cells"))?;
        for cell in candidates {
            worksheet.add_cell(cell);
        }
        Ok(())
    }
}

/// A pivot table definition for the writer.
#[derive(Debug, Clone)]
pub(super) struct WritablePivotTable {
    /// Pivot table name.
    pub name: String,
    /// Source type (0x0001 = Worksheet).
    pub source_type: u16,

    // -- Source data range (for DCONREF + SXDB cache) --
    /// Name of the source worksheet.
    pub source_sheet_name: String,
    /// Source range (0-based, inclusive).
    pub source_first_row: u16,
    pub source_last_row: u16,
    pub source_first_col: u16,
    pub source_last_col: u16,

    // -- Output range --
    pub first_row: u16,
    pub last_row: u16,
    pub first_col: u16,
    pub last_col: u16,
    /// First header row.
    pub first_header_row: u16,
    /// First data row.
    pub first_data_row: u16,
    /// First data column.
    pub first_data_col: u16,
    /// Data field header name (e.g. "Values").
    pub data_field_name: String,
    /// Axis for data field header.
    pub data_axis: u16,
    /// Position of data label within axis.
    pub data_position: u16,
    /// Field definitions.
    pub fields: Vec<WritablePivotField>,
    /// Data item definitions.
    pub data_items: Vec<WritablePivotDataItem>,
    /// Page field entries: (`item_index`, `field_index`, `object_id`).
    pub page_entries: Vec<(u16, u16, u16)>,
    /// Source data rows for the pivot cache.
    pub source_data: Vec<Vec<super::PivotCacheValue>>,
}

/// A pivot field definition for the writer.
#[derive(Debug, Clone)]
pub(super) struct WritablePivotField {
    /// Axis: 0=none, 1=row, 2=col, 4=page, 8=data.
    pub axis: u16,
    pub subtotal_count: u16,
    pub subtotal_flags: u16,
    /// Items in this field.
    pub items: Vec<WritablePivotItem>,
    /// Optional SXVD display name override (`None` → use cache name).
    pub name: Option<String>,
    /// Source column name for the pivot cache SXFDB record.
    pub cache_name: String,
    /// Unique source data values for this field's cache items (SXSTRING records).
    pub cache_items: Vec<crate::PivotCacheItem>,
    /// Whether this field is numeric (data-axis).
    pub is_numeric: bool,
    pub grouping: Option<crate::PivotCacheGrouping>,
}

/// A pivot item for the writer.
#[derive(Debug, Clone)]
pub(super) struct WritablePivotItem {
    /// Item type: 0x0000=Data, 0x0001=Default subtotal, 0x0002=Sum, etc.
    pub item_type: u16,
    pub flags: u16,
    pub cache_index: u16,
    pub name: Option<String>,
}

/// A pivot data item (value field) for the writer.
#[derive(Debug, Clone)]
pub(super) struct WritablePivotDataItem {
    pub source_field_index: u16,
    /// Aggregation function: 0=Sum,1=Count,2=Average,...
    pub function: u16,
    pub display_format: u16,
    pub base_field_index: u16,
    pub base_item_index: u16,
    pub num_format_index: u16,
    pub name: String,
}
