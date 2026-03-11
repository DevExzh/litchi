//! XLS file writer implementation
//!
//! This module provides functionality to create and modify Microsoft Excel files
//! in the legacy binary format (.xls files) using the BIFF (Binary Interchange File Format).
//!
//! # Architecture
//!
//! The writer generates BIFF8 records and uses the OLE writer to create the
//! compound document structure. It supports:
//! - Creating workbooks with multiple worksheets
//! - Writing cell values (numbers, strings, formulas, booleans)
//! - Shared string table management
//! - Basic cell formatting
//!
//! # Example
//!
//! ```rust,no_run
//! use litchi::ole::xls::XlsWriter;
//!
//! let mut writer = XlsWriter::new();
//! let sheet = writer.add_worksheet("Sheet1")?;
//!
//! // Write some data
//! writer.write_string(sheet, 0, 0, "Hello")?;
//! writer.write_number(sheet, 0, 1, 42.0)?;
//!
//! // Save the file
//! writer.save("output.xls")?;
//! # Ok::<(), Box<dyn std::error::Error>>(())
//! ```

use super::super::error::{XlsError, XlsResult};
use super::biff::AutoFilterConditionWrite;
use super::formatting::{CellStyle, ExtendedFormat, FormattingManager};
use crate::ole::writer::OleWriter;
use std::collections::HashMap;

mod conditional_format;
mod data_validation;
mod named_range;
mod stream;
mod worksheet;

pub use self::conditional_format::{
    XlsConditionalFormat, XlsConditionalFormatType, XlsConditionalPattern,
};
pub use self::data_validation::{
    XlsDataValidation, XlsDataValidationOperator, XlsDataValidationType,
};
pub use self::named_range::XlsDefinedName;
use self::named_range::XlsDefinedName as InternalDefinedName;
use self::worksheet::{
    AutoFilterColumnDef, AutoFilterRange, MergedRange, PivotCellXfRole, SortConfig, WritableCell,
    WritablePivotDataItem, WritablePivotField, WritablePivotItem, WritablePivotTable,
    WritableWorksheet, XlsHyperlink, XlsSheetProtection,
};

/// Public configuration for adding a pivot table via [`XlsWriter::add_pivot_table`].
#[derive(Debug, Clone)]
pub struct XlsPivotTableConfig {
    /// Pivot table name.
    pub name: String,
    /// Source type (0x0001 = Worksheet, 0x0002 = External).
    pub source_type: u16,

    // -- Source data range --
    /// Name of the worksheet that holds the source data.
    pub source_sheet_name: String,
    /// First row of the source data range (0-based, **including** the header row).
    pub source_first_row: u16,
    /// Last row of the source data range (0-based, inclusive).
    pub source_last_row: u16,
    /// First column of the source data range (0-based).
    pub source_first_col: u16,
    /// Last column of the source data range (0-based).
    pub source_last_col: u16,

    // -- Output range --
    /// First row of the pivot table output.
    pub first_row: u16,
    pub last_row: u16,
    pub first_col: u16,
    pub last_col: u16,
    /// First header row in the output.
    pub first_header_row: u16,
    /// First data row in the output.
    pub first_data_row: u16,
    /// First data column in the output.
    pub first_data_col: u16,
    /// Data field header name (e.g. "Values").
    pub data_field_name: String,
    /// Axis for the data field header (0=none, 1=row, 2=col, 4=page, 8=data).
    pub data_axis: u16,
    /// Position of data label within the axis.
    pub data_position: u16,
    /// Field definitions.
    pub fields: Vec<XlsPivotFieldConfig>,
    /// Data item (value field) definitions.
    pub data_items: Vec<XlsPivotDataItemConfig>,
    /// Page field entries: `(item_index, field_index, object_id)`.
    pub page_entries: Vec<(u16, u16, u16)>,
    /// Source data rows for the pivot cache (fSaveData).
    ///
    /// Each inner `Vec` has one entry per field in the same order as `fields`.
    /// String fields use [`PivotCacheValue::StringIndex`] (index into that
    /// field's `cache_items`), numeric fields use [`PivotCacheValue::Number`].
    ///
    /// When non-empty, SXDBB + SXNUM records are written to the cache stream
    /// and the SXDB `fSaveData` flag is set.
    pub source_data: Vec<Vec<PivotCacheValue>>,
}

/// A single pivot field definition.
#[derive(Debug, Clone)]
pub struct XlsPivotFieldConfig {
    /// Axis: 0=none, 1=row, 2=col, 4=page, 8=data.
    pub axis: u16,
    /// Number of subtotals.
    pub subtotal_count: u16,
    /// Subtotal function bitmask.
    pub subtotal_flags: u16,
    /// Items belonging to this field.
    pub items: Vec<XlsPivotItemConfig>,
    /// Optional SXVD display name override (`None` → use cache name, i.e. cch=0xFFFF).
    pub name: Option<String>,
    /// Source column name used in the pivot cache SXFDB record.
    /// This is the actual header text from the source data range.
    pub cache_name: String,
    /// Unique source data values for this field's cache items.
    /// These become SXSTRING records in the pivot cache stream.
    /// For data-axis (numeric) fields, leave this empty.
    pub cache_items: Vec<String>,
    /// Whether this is a numeric (data-axis) field.
    ///
    /// Numeric fields use SXFDB flags `0x0560` and contribute SXNUM records
    /// (instead of SXDBB indices) in the cache source data.
    pub is_numeric: bool,
}

/// A single cell value in the pivot cache source data.
#[derive(Debug, Clone, Copy)]
pub enum PivotCacheValue {
    /// Index into the field's `cache_items` (for string fields).
    StringIndex(u8),
    /// Raw numeric value (for numeric/data-axis fields).
    Number(f64),
}

/// A single pivot item.
#[derive(Debug, Clone)]
pub struct XlsPivotItemConfig {
    /// Item type: 0x0000=Data, 0x0001=Default subtotal, 0x0002=Sum, etc.
    pub item_type: u16,
    /// Option flags.
    pub flags: u16,
    /// Cache index.
    pub cache_index: u16,
    /// Optional item name override.
    pub name: Option<String>,
}

/// A pivot data item (value field).
#[derive(Debug, Clone)]
pub struct XlsPivotDataItemConfig {
    /// Index of the source field in the pivot cache.
    pub source_field_index: u16,
    /// Aggregation function: 0=Sum, 1=Count, 2=Average, 3=Max, 4=Min, ...
    pub function: u16,
    /// Display format flags.
    pub display_format: u16,
    /// Base field index (for "show values as").
    pub base_field_index: u16,
    /// Base item index.
    pub base_item_index: u16,
    /// Number format index.
    pub num_format_index: u16,
    /// Optional name override.
    pub name: String,
}

fn column_to_letters(col: u16) -> String {
    let mut col_index = col as u32;
    let mut buf = Vec::new();

    loop {
        let rem = (col_index % 26) as u8;
        buf.push((b'A' + rem) as char);
        col_index /= 26;
        if col_index == 0 {
            break;
        }
        col_index -= 1;
    }

    buf.iter().rev().collect()
}

fn a1_cell(row: u32, col: u16) -> String {
    let col_str = column_to_letters(col);
    let row_idx = row + 1;
    format!("{col_str}{row_idx}")
}

/// Cell value type for writing
#[derive(Debug, Clone)]
pub enum XlsCellValue {
    /// String value
    String(String),
    /// Number value (f64)
    Number(f64),
    /// Boolean value
    Boolean(bool),
    /// Formula (stored as string)
    Formula(String),
    /// Blank/empty cell
    Blank,
}
#[derive(Debug, Clone, Copy)]
struct XlsWorkbookProtection {
    protect_structure: bool,
    protect_windows: bool,
    password_hash: Option<u16>,
}
/// XLS file writer
///
/// Provides methods to create and modify XLS (BIFF8) files.
pub struct XlsWriter {
    /// Worksheets to write
    worksheets: Vec<WritableWorksheet>,
    /// Shared string table
    shared_strings: Vec<String>,
    /// String to index mapping for deduplication
    string_map: HashMap<String, u32>,
    /// Workbook-level defined names (named ranges).
    defined_names: Vec<InternalDefinedName>,
    fmt: FormattingManager,
    /// Total number of string occurrences (including duplicates) for SST.cstTotal
    sst_total: u32,
    workbook_protection: Option<XlsWorkbookProtection>,
    /// Use 1904 date system (Mac) instead of 1900 (Windows)
    use_1904_dates: bool,
}

impl XlsWriter {
    /// Create a new XLS writer
    pub fn new() -> Self {
        Self {
            worksheets: Vec::new(),
            shared_strings: Vec::new(),
            string_map: HashMap::new(),
            defined_names: Vec::new(),
            sst_total: 0,
            fmt: FormattingManager::new(),
            workbook_protection: None,
            use_1904_dates: false,
        }
    }

    /// Add a new worksheet
    ///
    /// # Arguments
    ///
    /// * `name` - Worksheet name (max 31 characters)
    ///
    /// # Returns
    ///
    /// * `Result<usize, XlsError>` - Worksheet index or error
    pub fn add_worksheet(&mut self, name: &str) -> XlsResult<usize> {
        // Validate worksheet name
        if name.is_empty() || name.len() > 31 {
            return Err(XlsError::InvalidData(
                "Worksheet name must be 1-31 characters".to_string(),
            ));
        }

        // Check for duplicate names
        if self.worksheets.iter().any(|ws| ws.name == name) {
            return Err(XlsError::InvalidData(format!(
                "Worksheet '{}' already exists",
                name
            )));
        }

        let index = self.worksheets.len();
        self.worksheets
            .push(WritableWorksheet::new(name.to_string()));
        Ok(index)
    }

    /// Write a string value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - String value
    pub fn write_string(&mut self, sheet: usize, row: u32, col: u16, value: &str) -> XlsResult<()> {
        self.write_string_with_format(sheet, row, col, value, 0)
    }

    pub fn write_string_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: &str,
        format_id: u16,
    ) -> XlsResult<()> {
        self.write_cell(
            sheet,
            row,
            col,
            XlsCellValue::String(value.to_string()),
            format_id,
        )
    }

    /// Write a number value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Numeric value
    pub fn write_number(&mut self, sheet: usize, row: u32, col: u16, value: f64) -> XlsResult<()> {
        self.write_number_with_format(sheet, row, col, value, 0)
    }

    pub fn write_number_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: f64,
        format_id: u16,
    ) -> XlsResult<()> {
        self.write_cell(sheet, row, col, XlsCellValue::Number(value), format_id)
    }

    /// Write a boolean value to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `value` - Boolean value
    pub fn write_boolean(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: bool,
    ) -> XlsResult<()> {
        self.write_boolean_with_format(sheet, row, col, value, 0)
    }

    pub fn write_boolean_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: bool,
        format_id: u16,
    ) -> XlsResult<()> {
        self.write_cell(sheet, row, col, XlsCellValue::Boolean(value), format_id)
    }

    /// Write a formula to a cell
    ///
    /// # Arguments
    ///
    /// * `sheet` - Worksheet index
    /// * `row` - Row index (0-based)
    /// * `col` - Column index (0-based)
    /// * `formula` - Formula string (without leading '=')
    ///
    /// # Implementation Notes
    ///
    /// Formula tokenization is deferred as a future enhancement.
    /// Formulas are currently written as blank cells.
    pub fn write_formula(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        formula: &str,
    ) -> XlsResult<()> {
        self.write_formula_with_format(sheet, row, col, formula, 0)
    }

    pub fn write_formula_with_format(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        formula: &str,
        format_id: u16,
    ) -> XlsResult<()> {
        self.write_cell(
            sheet,
            row,
            col,
            XlsCellValue::Formula(formula.to_string()),
            format_id,
        )
    }

    /// Register a number format pattern and return its BIFF format index.
    ///
    /// This is a thin wrapper around the internal `FormattingManager`
    /// and mirrors Apache POI's `HSSFDataFormat.getFormat` API. The
    /// returned index can be stored in `ExtendedFormat.format_index`
    /// to apply number formats to cells.
    pub fn register_number_format(&mut self, pattern: &str) -> u16 {
        self.fmt.register_number_format(pattern)
    }

    /// Register a reusable cell style defined by `CellStyle`.
    ///
    /// The returned identifier can be passed to the `write_*_with_format`
    /// methods to apply this style to individual cells.
    pub fn add_cell_style(&mut self, style: CellStyle) -> u16 {
        self.fmt.register_cell_style(style)
    }

    pub fn add_cell_format(&mut self, format: ExtendedFormat) -> u16 {
        self.fmt.add_format(format)
    }

    /// Validate a defined name according to basic Excel constraints.
    ///
    /// This helper enforces only well-defined structural rules from the
    /// specification:
    /// - Name MUST NOT be empty.
    /// - Name length MUST be at most 255 characters (Lbl.cch is a byte).
    fn validate_defined_name(name: &str) -> XlsResult<()> {
        if name.is_empty() {
            return Err(XlsError::InvalidData(
                "Defined name must not be empty".to_string(),
            ));
        }

        let char_count = name.chars().count();
        if char_count > u8::MAX as usize {
            return Err(XlsError::InvalidData(
                "Defined name must be at most 255 characters".to_string(),
            ));
        }

        Ok(())
    }

    fn hash_password(password: &str) -> u16 {
        let bytes = password.as_bytes();
        if bytes.is_empty() {
            return 0;
        }

        let mut hash: u16 = 0;
        for &b in bytes.iter().rev() {
            let high_bit = (hash >> 14) & 0x0001;
            hash = ((hash << 1) & 0x7FFF) | high_bit;
            hash ^= b as u16;
        }

        let high_bit = (hash >> 14) & 0x0001;
        hash = ((hash << 1) & 0x7FFF) | high_bit;
        hash ^= bytes.len() as u16;
        hash ^= 0xCE4B;
        hash
    }

    /// Set a hyperlink for a single cell.
    ///
    /// Row and column indices are 0-based, matching the rest of the XLS
    /// writer APIs. The hyperlink target can be a standard URL (http, https,
    /// ftp, mailto) or an internal reference such as `Sheet1!A1` or
    /// `internal:Sheet1!A1`.
    pub fn set_hyperlink(&mut self, sheet: usize, row: u32, col: u16, url: &str) -> XlsResult<()> {
        if row > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "set_hyperlink: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        if col >= 256 {
            return Err(XlsError::InvalidData(
                "set_hyperlink: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        // Replace any existing hyperlink on this exact cell to match
        // XLSX writer semantics.
        worksheet.hyperlinks.retain(|h| {
            !(h.first_row == row && h.last_row == row && h.first_col == col && h.last_col == col)
        });

        worksheet.add_hyperlink(XlsHyperlink {
            first_row: row,
            last_row: row,
            first_col: col,
            last_col: col,
            url: url.to_string(),
        });

        Ok(())
    }

    pub fn set_auto_filter(
        &mut self,
        sheet: usize,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> XlsResult<()> {
        if first_row > last_row || first_col > last_col {
            return Err(XlsError::InvalidData(
                "set_auto_filter: first row/col must be <= last row/col".to_string(),
            ));
        }

        if last_row > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "set_auto_filter: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        if last_col >= 256 {
            return Err(XlsError::InvalidData(
                "set_auto_filter: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.auto_filter = Some(AutoFilterRange {
            first_row,
            last_row,
            first_col,
            last_col,
        });

        let itab = u16::try_from(sheet + 1).map_err(|_| {
            XlsError::InvalidData(
                "set_auto_filter: sheet index exceeds BIFF8 itab limit".to_string(),
            )
        })?;

        self.defined_names.retain(|n| {
            !(n.is_built_in && n.built_in_code == Some(0x0D) && n.local_sheet == Some(itab))
        });

        let start_ref = a1_cell(first_row, first_col);
        let end_ref = a1_cell(last_row, last_col);
        let reference = format!("{start_ref}:{end_ref}");

        self.defined_names.push(InternalDefinedName {
            name: "_FilterDatabase".to_string(),
            reference,
            comment: None,
            local_sheet: Some(itab),
            target_sheet: Some(sheet as u16),
            hidden: true,
            is_function: false,
            is_built_in: true,
            built_in_code: Some(0x0D),
        });

        Ok(())
    }

    /// Add a filter condition to a specific column within the AutoFilter range.
    ///
    /// The AutoFilter range must first be set via [`set_auto_filter`]. The
    /// `column_index` is 0-based relative to the filter range start column.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `column_index` — column within the filter range (0-based relative)
    /// * `join_or` — `true` to join conditions with OR, `false` for AND
    /// * `cond1` — first filter condition
    /// * `cond2` — second filter condition (use `AutoFilterConditionWrite::None` if unused)
    ///
    /// # Example
    ///
    /// ```ignore
    /// use litchi::ole::xls::writer::biff::AutoFilterConditionWrite;
    ///
    /// // Filter column 2: value > 100
    /// writer.add_filter_condition(
    ///     sheet_idx, 2, false,
    ///     AutoFilterConditionWrite::Number { operator: 0x04, value: 100.0 },
    ///     AutoFilterConditionWrite::None,
    /// )?;
    /// ```
    pub fn add_filter_condition(
        &mut self,
        sheet: usize,
        column_index: u16,
        join_or: bool,
        cond1: AutoFilterConditionWrite,
        cond2: AutoFilterConditionWrite,
    ) -> XlsResult<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        if worksheet.auto_filter.is_none() {
            return Err(XlsError::InvalidData(
                "add_filter_condition: call set_auto_filter first".to_string(),
            ));
        }

        worksheet.add_auto_filter_column(AutoFilterColumnDef {
            column_index,
            join_or,
            condition1: cond1,
            condition2: cond2,
        });

        Ok(())
    }

    /// Set the sort configuration for a worksheet.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `case_sensitive` — whether sorting is case-sensitive
    /// * `sort_by_columns` — `true` for left-to-right sort, `false` for top-to-bottom
    /// * `keys` — up to 3 sort keys as `(column_index, descending)` tuples
    pub fn set_sort(
        &mut self,
        sheet: usize,
        case_sensitive: bool,
        sort_by_columns: bool,
        keys: &[(u16, bool)],
    ) -> XlsResult<()> {
        if keys.is_empty() || keys.len() > 3 {
            return Err(XlsError::InvalidData(
                "set_sort: must provide 1..3 sort keys".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.set_sort_config(SortConfig {
            case_sensitive,
            sort_by_columns,
            keys: keys.to_vec(),
        });

        Ok(())
    }

    /// Add a pivot table definition to a worksheet.
    ///
    /// This writes the SX* record family (SXVS, SXVIEW, SXVD, SXVI, SXDI,
    /// SXPI) to the worksheet stream. The pivot table must be fully
    /// configured before calling this method.
    ///
    /// # Arguments
    ///
    /// * `sheet` — worksheet index (0-based)
    /// * `config` — pivot table configuration (see [`XlsPivotTableConfig`])
    pub fn add_pivot_table(&mut self, sheet: usize, config: XlsPivotTableConfig) -> XlsResult<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        self.fmt.enable_pivot_xfs();

        // Generate pivot output cells BEFORE consuming config.fields / config.data_items.
        // Excel validates that DIMENSIONS and cell content are consistent with the
        // pivot table definition; missing cells cause a "corrupt file" repair dialog.
        Self::generate_pivot_output_cells(worksheet, &config);

        let fields: Vec<WritablePivotField> = config
            .fields
            .into_iter()
            .map(|f| {
                let mut items: Vec<WritablePivotItem> = f
                    .items
                    .into_iter()
                    .map(|i| WritablePivotItem {
                        item_type: i.item_type,
                        flags: i.flags,
                        cache_index: i.cache_index,
                        name: i.name,
                    })
                    .collect();

                // Sort data items (item_type=0x0000) alphabetically by their
                // cache label to match Excel's default SXVI ordering.  Non-data
                // items (subtotals etc.) stay at the end.
                let data_end = items
                    .iter()
                    .position(|i| i.item_type != 0x0000)
                    .unwrap_or(items.len());
                items[..data_end].sort_unstable_by(|a, b| {
                    let al = f
                        .cache_items
                        .get(a.cache_index as usize)
                        .map(String::as_str)
                        .unwrap_or("");
                    let bl = f
                        .cache_items
                        .get(b.cache_index as usize)
                        .map(String::as_str)
                        .unwrap_or("");
                    al.cmp(bl)
                });

                WritablePivotField {
                    axis: f.axis,
                    subtotal_count: f.subtotal_count,
                    subtotal_flags: f.subtotal_flags,
                    items,
                    name: f.name,
                    cache_name: f.cache_name,
                    cache_items: f.cache_items,
                    is_numeric: f.is_numeric,
                }
            })
            .collect();

        let data_items: Vec<WritablePivotDataItem> = config
            .data_items
            .into_iter()
            .map(|d| WritablePivotDataItem {
                source_field_index: d.source_field_index,
                function: d.function,
                display_format: d.display_format,
                base_field_index: d.base_field_index,
                base_item_index: d.base_item_index,
                num_format_index: d.num_format_index,
                name: d.name,
            })
            .collect();

        worksheet.add_pivot_table(WritablePivotTable {
            name: config.name,
            source_type: config.source_type,
            source_sheet_name: config.source_sheet_name,
            source_first_row: config.source_first_row,
            source_last_row: config.source_last_row,
            source_first_col: config.source_first_col,
            source_last_col: config.source_last_col,
            first_row: config.first_row,
            last_row: config.last_row,
            first_col: config.first_col,
            last_col: config.last_col,
            first_header_row: config.first_header_row,
            first_data_row: config.first_data_row,
            first_data_col: config.first_data_col,
            data_field_name: config.data_field_name,
            data_axis: config.data_axis,
            data_position: config.data_position,
            fields,
            data_items,
            page_entries: config.page_entries,
            source_data: config.source_data,
        });

        Ok(())
    }

    /// Generate the cell data that Excel expects in the SXVIEW output area.
    ///
    /// The layout (for a single row-field, single col-field, single page-field,
    /// single data-field configuration) is:
    ///
    /// ```text
    /// (first_row-2, 0)       : page field name    (first_row-2, 1)       : "(All)"
    /// (first_row,   0)       : data item name      (first_row, first_data_col): "Column Labels"
    /// (first_header_row, 0)  : "Row Labels"        (fhr, fdc+j)           : col item names …
    /// (first_data_row+i, 0)  : row item name       (fdr+i, fdc+j)         : aggregated value
    /// (last_row, 0)          : "Grand Total"        (lr, fdc+j)            : column totals
    /// ```
    fn generate_pivot_output_cells(ws: &mut WritableWorksheet, cfg: &XlsPivotTableConfig) {
        // Identify fields per axis.
        let row_field = cfg.fields.iter().find(|f| f.axis == 0x0001);
        let col_field = cfg.fields.iter().find(|f| f.axis == 0x0002);
        let page_field = cfg.fields.iter().find(|f| f.axis == 0x0004);

        let data_item = cfg.data_items.first();

        // Helper: find the field index for a given field by cache_name.
        let field_idx_of =
            |name: &str| -> Option<usize> { cfg.fields.iter().position(|f| f.cache_name == name) };

        // Collect row/col item labels from cache_items, sorted alphabetically
        // to match Excel's default SXVI ordering.  Also build a mapping from
        // cache_index → sorted position so the aggregation grid uses the same
        // order as the output rows/columns.
        let (row_items, row_cache_to_sorted) = Self::sorted_cache_items(row_field);
        let (col_items, col_cache_to_sorted) = Self::sorted_cache_items(col_field);

        let fr = cfg.first_row;
        let fhr = cfg.first_header_row;
        let fdr = cfg.first_data_row;
        let fdc = cfg.first_data_col;
        let lr = cfg.last_row;
        let lc = cfg.last_col;
        let fc = cfg.first_col;

        let add = |ws: &mut WritableWorksheet,
                   r: u16,
                   c: u16,
                   v: XlsCellValue,
                   pivot_xf_role: Option<PivotCellXfRole>| {
            ws.add_cell(WritableCell {
                row: r as u32,
                col: c,
                value: v,
                format_idx: 0,
                pivot_xf_role,
            });
        };

        // --- Page field area (above SXVIEW range) ---
        if let Some(pf) = page_field {
            let page_row = fr.saturating_sub(2);
            add(
                ws,
                page_row,
                0,
                XlsCellValue::String(pf.cache_name.clone()),
                Some(PivotCellXfRole::HeaderAccent),
            );
            add(
                ws,
                page_row,
                1,
                XlsCellValue::String("(All)".to_string()),
                Some(PivotCellXfRole::HeaderPlain),
            );
        }

        // --- Row at first_row: data item name + "Column Labels" ---
        if let Some(di) = data_item {
            add(
                ws,
                fr,
                fc,
                XlsCellValue::String(di.name.clone()),
                Some(PivotCellXfRole::HeaderAccent),
            );
        }
        if col_field.is_some() {
            add(
                ws,
                fr,
                fdc,
                XlsCellValue::String("Column Labels".to_string()),
                Some(PivotCellXfRole::HeaderAccent),
            );
        }

        // --- Row at first_header_row: "Row Labels" + column item names + "Grand Total" ---
        add(
            ws,
            fhr,
            fc,
            XlsCellValue::String("Row Labels".to_string()),
            Some(PivotCellXfRole::HeaderAccent),
        );
        for (j, ci) in col_items.iter().enumerate() {
            add(
                ws,
                fhr,
                fdc + j as u16,
                XlsCellValue::String(ci.clone()),
                Some(PivotCellXfRole::HeaderPlain),
            );
        }
        add(
            ws,
            fhr,
            lc,
            XlsCellValue::String("Grand Total".to_string()),
            Some(PivotCellXfRole::HeaderPlain),
        );

        // --- Compute aggregated values from source_data ---
        let row_fi = row_field.and_then(|f| field_idx_of(&f.cache_name));
        let col_fi = col_field.and_then(|f| field_idx_of(&f.cache_name));
        let data_fi = data_item.map(|di| di.source_field_index as usize);

        let nr = row_items.len();
        let nc = col_items.len();
        let mut grid = vec![vec![0.0f64; nc]; nr];
        let mut row_totals = vec![0.0f64; nr];
        let mut col_totals = vec![0.0f64; nc];
        let mut grand_total = 0.0f64;

        for row_data in &cfg.source_data {
            // Map cache indices through the sorted permutation so that
            // grid positions match the alphabetically-sorted output.
            let ri = row_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::StringIndex(idx)) => {
                    row_cache_to_sorted.get(*idx as usize).copied()
                },
                _ => None,
            });
            let ci = col_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::StringIndex(idx)) => {
                    col_cache_to_sorted.get(*idx as usize).copied()
                },
                _ => None,
            });
            let val = data_fi.and_then(|fi| match row_data.get(fi) {
                Some(PivotCacheValue::Number(v)) => Some(*v),
                _ => None,
            });

            if let (Some(ri), Some(ci), Some(val)) = (ri, ci, val)
                && ri < nr
                && ci < nc
            {
                grid[ri][ci] += val;
                row_totals[ri] += val;
                col_totals[ci] += val;
                grand_total += val;
            }
        }

        // --- Data rows ---
        for (i, (ri_name, row_total)) in row_items.iter().zip(row_totals.iter()).enumerate() {
            let r = fdr + i as u16;
            add(
                ws,
                r,
                fc,
                XlsCellValue::String(ri_name.clone()),
                Some(PivotCellXfRole::RowLabel),
            );
            for (j, cell_val) in grid[i].iter().enumerate() {
                add(
                    ws,
                    r,
                    fdc + j as u16,
                    XlsCellValue::Number(*cell_val),
                    Some(PivotCellXfRole::Value),
                );
            }
            add(
                ws,
                r,
                lc,
                XlsCellValue::Number(*row_total),
                Some(PivotCellXfRole::Value),
            );
        }

        // --- Grand total row ---
        add(
            ws,
            lr,
            fc,
            XlsCellValue::String("Grand Total".to_string()),
            Some(PivotCellXfRole::RowLabel),
        );
        for (j, col_total) in col_totals.iter().enumerate() {
            add(
                ws,
                lr,
                fdc + j as u16,
                XlsCellValue::Number(*col_total),
                Some(PivotCellXfRole::Value),
            );
        }
        add(
            ws,
            lr,
            lc,
            XlsCellValue::Number(grand_total),
            Some(PivotCellXfRole::Value),
        );
    }

    /// Sort a field's cache items alphabetically and return the sorted labels
    /// plus a mapping from original cache index to sorted position.
    ///
    /// Returns `(sorted_labels, cache_to_sorted)` where `cache_to_sorted[i]`
    /// gives the position of original cache item `i` in the sorted output.
    fn sorted_cache_items(field: Option<&XlsPivotFieldConfig>) -> (Vec<String>, Vec<usize>) {
        let Some(f) = field else {
            return (Vec::new(), Vec::new());
        };

        // Build (original_index, label) pairs and sort by label.
        let mut indexed: Vec<(usize, &str)> = f
            .cache_items
            .iter()
            .enumerate()
            .map(|(i, s)| (i, s.as_str()))
            .collect();
        indexed.sort_unstable_by(|a, b| a.1.cmp(b.1));

        let sorted_labels: Vec<String> = indexed.iter().map(|(_, s)| (*s).to_string()).collect();

        // cache_to_sorted[original_cache_idx] = position in sorted output
        let mut cache_to_sorted = vec![0usize; f.cache_items.len()];
        for (sorted_pos, &(orig_idx, _)) in indexed.iter().enumerate() {
            cache_to_sorted[orig_idx] = sorted_pos;
        }

        (sorted_labels, cache_to_sorted)
    }

    /// Define a workbook-scoped named range.
    ///
    /// The reference must currently be a simple A1 or A1:B10 style range
    /// without sheet qualifiers. More complex formulas will be rejected
    /// at serialization time to avoid emitting invalid BIFF payloads.
    pub fn define_name(&mut self, name: &str, reference: &str) -> XlsResult<()> {
        Self::validate_defined_name(name)?;

        if self.worksheets.is_empty() {
            return Err(XlsError::InvalidData(
                "define_name: workbook must have at least one worksheet".to_string(),
            ));
        }

        // For now, workbook-scoped names that refer to cell ranges are
        // anchored to the first worksheet. Users who need explicit
        // sheet scoping can use `define_name_local`.
        let target_sheet = 0u16;

        self.defined_names.push(InternalDefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Define a sheet-scoped named range.
    ///
    /// `sheet` is a 0-based worksheet index.
    pub fn define_name_local(
        &mut self,
        name: &str,
        reference: &str,
        sheet: usize,
    ) -> XlsResult<()> {
        Self::validate_defined_name(name)?;

        let _ = self
            .worksheets
            .get(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let itab = u16::try_from(sheet + 1).map_err(|_| {
            XlsError::InvalidData(
                "define_name_local: sheet index exceeds BIFF8 itab limit".to_string(),
            )
        })?;

        self.defined_names.push(InternalDefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet: Some(itab),
            target_sheet: Some(sheet as u16),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Define a workbook-scoped named range with a user-visible comment.
    pub fn define_name_with_comment(
        &mut self,
        name: &str,
        reference: &str,
        comment: &str,
    ) -> XlsResult<()> {
        Self::validate_defined_name(name)?;

        if self.worksheets.is_empty() {
            return Err(XlsError::InvalidData(
                "define_name_with_comment: workbook must have at least one worksheet".to_string(),
            ));
        }

        let target_sheet = 0u16;

        self.defined_names.push(InternalDefinedName {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: Some(comment.to_string()),
            local_sheet: None,
            target_sheet: Some(target_sheet),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        });

        Ok(())
    }

    /// Remove all defined names with the given name.
    ///
    /// Returns `true` if at least one name was removed.
    pub fn remove_name(&mut self, name: &str) -> bool {
        let initial_len = self.defined_names.len();
        self.defined_names.retain(|n| n.name != name);
        self.defined_names.len() < initial_len
    }

    /// Get all defined names in this workbook.
    pub fn named_ranges(&self) -> &[XlsDefinedName] {
        &self.defined_names
    }

    /// Set the width of a column in character units.
    ///
    /// The column index is 0-based (0 = column A), matching the rest of the
    /// XLS writer API. The width is specified in the same units as Excel's
    /// UI, i.e. the number of characters of the "0" glyph in the default
    /// font. Internally this is converted to BIFF8 units of 1/256 characters
    /// for the COLINFO record.
    pub fn set_column_width(&mut self, sheet: usize, col: u16, width_chars: f64) -> XlsResult<()> {
        if col >= 256 {
            return Err(XlsError::InvalidData(
                "set_column_width: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        if !(width_chars.is_finite()) || width_chars <= 0.0 {
            return Err(XlsError::InvalidData(
                "set_column_width: width must be a positive finite value".to_string(),
            ));
        }

        let max_units = 255u32 * 256u32; // Excel maximum column width
        let width_units_f = (width_chars * 256.0).round();
        if width_units_f <= 0.0 || width_units_f > max_units as f64 {
            return Err(XlsError::InvalidData(
                "set_column_width: width exceeds Excel's maximum (255 characters)".to_string(),
            ));
        }

        let width_units = width_units_f as u16;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.set_column_width(col, width_units);
        Ok(())
    }

    /// Hide a column.
    pub fn hide_column(&mut self, sheet: usize, col: u16) -> XlsResult<()> {
        if col >= 256 {
            return Err(XlsError::InvalidData(
                "hide_column: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.hide_column(col);
        Ok(())
    }

    /// Show a previously hidden column.
    pub fn show_column(&mut self, sheet: usize, col: u16) -> XlsResult<()> {
        if col >= 256 {
            return Err(XlsError::InvalidData(
                "show_column: column index must be < 256 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.show_column(col);
        Ok(())
    }

    pub fn merge_cells(
        &mut self,
        sheet: usize,
        first_row: u32,
        last_row: u32,
        first_col: u16,
        last_col: u16,
    ) -> XlsResult<()> {
        if first_row > last_row || first_col > last_col {
            return Err(XlsError::InvalidData(
                "merge_cells: first row/col must be <= last row/col".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_merged_range(MergedRange {
            first_row,
            last_row,
            first_col,
            last_col,
        });

        Ok(())
    }

    /// Configure freeze panes for the specified worksheet.
    ///
    /// Row and column indices are 0-based and represent the number of
    /// rows/columns at the top/left that remain frozen.
    pub fn freeze_panes(
        &mut self,
        sheet: usize,
        freeze_rows: u32,
        freeze_cols: u16,
    ) -> XlsResult<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        if freeze_rows == 0 && freeze_cols == 0 {
            worksheet.clear_freeze_panes();
            return Ok(());
        }

        if freeze_rows > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "freeze_panes: freeze_rows must be <= 65535".to_string(),
            ));
        }

        worksheet.set_freeze_panes(freeze_rows, freeze_cols);
        Ok(())
    }

    /// Remove any freeze panes from the specified worksheet.
    pub fn unfreeze_panes(&mut self, sheet: usize) -> XlsResult<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.clear_freeze_panes();
        Ok(())
    }

    /// Set the height of a row in points.
    ///
    /// The row index is 0-based (0 = first row), and the height is specified
    /// in typographic points. Internally this is converted to twips
    /// (1/20th of a point) for the BIFF8 ROW record.
    pub fn set_row_height(&mut self, sheet: usize, row: u32, height_points: f64) -> XlsResult<()> {
        if !(height_points.is_finite()) || height_points <= 0.0 {
            return Err(XlsError::InvalidData(
                "set_row_height: height must be a positive finite value".to_string(),
            ));
        }

        if row > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "set_row_height: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let height_units_f = (height_points * 20.0).round();
        if height_units_f <= 0.0 || height_units_f > u16::MAX as f64 {
            return Err(XlsError::InvalidData(
                "set_row_height: height exceeds BIFF8 row height limit".to_string(),
            ));
        }

        let height_units = height_units_f as u16;

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.set_row_height(row, height_units);
        Ok(())
    }

    /// Hide a row.
    pub fn hide_row(&mut self, sheet: usize, row: u32) -> XlsResult<()> {
        if row > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "hide_row: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.hide_row(row);
        Ok(())
    }

    /// Show a previously hidden row.
    pub fn show_row(&mut self, sheet: usize, row: u32) -> XlsResult<()> {
        if row > u16::MAX as u32 {
            return Err(XlsError::InvalidData(
                "show_row: row index must be <= 65535 for BIFF8".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.show_row(row);
        Ok(())
    }

    /// Add a data validation rule to the specified worksheet.
    pub fn add_data_validation(
        &mut self,
        sheet: usize,
        validation: XlsDataValidation,
    ) -> XlsResult<()> {
        if validation.first_row > validation.last_row || validation.first_col > validation.last_col
        {
            return Err(XlsError::InvalidData(
                "add_data_validation: first row/col must be <= last row/col".to_string(),
            ));
        }

        if let Some(title) = validation.input_title.as_ref()
            && title.len() > 32
        {
            return Err(XlsError::InvalidData(
                "Input message title must be at most 32 characters".to_string(),
            ));
        }
        if let Some(text) = validation.input_message.as_ref()
            && text.len() > 255
        {
            return Err(XlsError::InvalidData(
                "Input message text must be at most 255 characters".to_string(),
            ));
        }
        if let Some(title) = validation.error_title.as_ref()
            && title.len() > 32
        {
            return Err(XlsError::InvalidData(
                "Error message title must be at most 32 characters".to_string(),
            ));
        }
        if let Some(text) = validation.error_message.as_ref()
            && text.len() > 255
        {
            return Err(XlsError::InvalidData(
                "Error message text must be at most 255 characters".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_data_validation(validation);

        Ok(())
    }

    pub fn add_conditional_format(
        &mut self,
        sheet: usize,
        cf: XlsConditionalFormat,
    ) -> XlsResult<()> {
        if cf.first_row > cf.last_row || cf.first_col > cf.last_col {
            return Err(XlsError::InvalidData(
                "add_conditional_format: first row/col must be <= last row/col".to_string(),
            ));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_conditional_format(cf);

        Ok(())
    }

    fn write_cell(
        &mut self,
        sheet: usize,
        row: u32,
        col: u16,
        value: XlsCellValue,
        format_id: u16,
    ) -> XlsResult<()> {
        if self.fmt.get_format(format_id).is_none() {
            return Err(XlsError::InvalidFormat(format_id));
        }

        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        worksheet.add_cell(WritableCell {
            row,
            col,
            value,
            format_idx: format_id,
            pivot_xf_role: None,
        });

        Ok(())
    }

    /// Set the date system (1900 vs 1904)
    ///
    /// # Arguments
    ///
    /// * `use_1904` - True to use 1904 date system (Mac), false for 1900 (Windows, default)
    pub fn set_1904_dates(&mut self, use_1904: bool) {
        self.use_1904_dates = use_1904;
    }

    pub fn protect_workbook(
        &mut self,
        password: Option<&str>,
        protect_structure: bool,
        protect_windows: bool,
    ) {
        if !protect_structure && !protect_windows && password.is_none() {
            self.workbook_protection = None;
            return;
        }

        let password_hash = password.map(Self::hash_password);
        self.workbook_protection = Some(XlsWorkbookProtection {
            protect_structure,
            protect_windows,
            password_hash,
        });
    }

    pub fn unprotect_workbook(&mut self) {
        self.workbook_protection = None;
    }

    pub fn protect_sheet(
        &mut self,
        sheet: usize,
        password: Option<&str>,
        protect_objects: bool,
        protect_scenarios: bool,
    ) -> XlsResult<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;

        let password_hash = password.map(Self::hash_password);
        worksheet.sheet_protection = Some(XlsSheetProtection {
            protect_objects,
            protect_scenarios,
            password_hash,
        });

        Ok(())
    }

    pub fn unprotect_sheet(&mut self, sheet: usize) -> XlsResult<()> {
        let worksheet = self
            .worksheets
            .get_mut(sheet)
            .ok_or_else(|| XlsError::WorksheetNotFound(format!("Sheet {}", sheet)))?;
        worksheet.sheet_protection = None;
        Ok(())
    }

    /// Save the XLS file
    ///
    /// # Arguments
    ///
    /// * `path` - Output file path
    ///
    /// # Returns
    ///
    /// * `Result<(), XlsError>` - Success or error
    ///
    /// # Implementation Status
    ///
    /// ✅ Basic structure generation (BOF, EOF, workbook globals)
    /// ✅ Cell record generation (Number, LabelSST, BoolErr)
    /// ✅ Shared string table (SST)
    /// ❌ Formula tokenization (formulas stored as values currently)
    /// ❌ Cell formatting (XF records)
    /// ❌ Column widths / row heights
    /// ❌ Merged cells
    /// ❌ Named ranges
    pub fn save<P: AsRef<std::path::Path>>(&mut self, path: P) -> XlsResult<()> {
        // Build shared string table
        self.build_shared_strings();

        // Generate the Workbook stream + pivot cache streams
        let streams = self.generate_workbook_streams()?;

        // Create OLE compound document
        let mut ole_writer = OleWriter::new();
        ole_writer.create_stream(&["Workbook"], &streams.workbook)?;

        // Pivot cache storage: _SX_DB_CUR/XXXX
        // Stream names use 4-digit uppercase hex per LO ScfTools::GetHexStr.
        if !streams.pivot_caches.is_empty() {
            ole_writer.create_storage(&["_SX_DB_CUR"])?;
            for (id, data) in &streams.pivot_caches {
                let name = format!("{:04X}", id);
                ole_writer.create_stream(&["_SX_DB_CUR", &name], data)?;
            }
        }

        // Save to file
        ole_writer.save(path)?;

        Ok(())
    }

    /// Write to a writer (useful for testing and in-memory generation)
    ///
    /// # Arguments
    ///
    /// * `writer` - Output writer
    ///
    /// # Returns
    ///
    /// * `Result<(), XlsError>` - Success or error
    pub fn write_to<W: std::io::Write + std::io::Seek>(&mut self, writer: &mut W) -> XlsResult<()> {
        // Build shared string table
        self.build_shared_strings();

        // Generate the Workbook stream + pivot cache streams
        let streams = self.generate_workbook_streams()?;

        // Create OLE compound document
        let mut ole_writer = OleWriter::new();
        ole_writer.create_stream(&["Workbook"], &streams.workbook)?;

        // Pivot cache storage: _SX_DB_CUR/XXXX
        if !streams.pivot_caches.is_empty() {
            ole_writer.create_storage(&["_SX_DB_CUR"])?;
            for (id, data) in &streams.pivot_caches {
                let name = format!("{:04X}", id);
                ole_writer.create_stream(&["_SX_DB_CUR", &name], data)?;
            }
        }

        // Write to the provided writer
        ole_writer.write_to(writer)?;

        Ok(())
    }

    /// Build the shared string table from all string cells
    fn build_shared_strings(&mut self) {
        self.shared_strings.clear();
        self.string_map.clear();
        self.sst_total = 0;

        // Collect all unique strings from all worksheets
        for worksheet in &self.worksheets {
            for cell in worksheet.cells.values() {
                if let XlsCellValue::String(ref s) = cell.value {
                    // Count total occurrences
                    self.sst_total = self.sst_total.saturating_add(1);
                    // Insert unique strings
                    if !self.string_map.contains_key(s) {
                        let index = self.shared_strings.len() as u32;
                        self.string_map.insert(s.clone(), index);
                        self.shared_strings.push(s.clone());
                    }
                }
            }
        }
    }

    /// Generate the complete Workbook stream (plus pivot cache streams) with
    /// all BIFF records.
    fn generate_workbook_streams(&self) -> XlsResult<stream::WorkbookStreams> {
        stream::generate_workbook_stream(
            self.use_1904_dates,
            &self.fmt,
            &self.defined_names,
            &self.shared_strings,
            self.sst_total,
            self.workbook_protection,
            &self.worksheets,
            &self.string_map,
        )
    }

    /// Get the number of worksheets in this workbook
    pub fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    /// Get worksheet name by index
    pub fn get_worksheet_name(&self, index: usize) -> Option<&str> {
        self.worksheets.get(index).map(|w| w.name.as_str())
    }

    // Implementation status notes:
    // ✅ Building shared string table (SST) with deduplication - IMPLEMENTED
    // ✅ Generating BIFF8 records for all cell types - IMPLEMENTED (Number, LabelSST, BoolErr)
    // ❌ Worksheet management (rename, delete, reorder) - Future enhancement
    // ❌ Cell formatting (fonts, colors, borders, number formats) - Future enhancement
    // ❌ Column widths and row heights - Future enhancement
    // ❌ Merged cells - Future enhancement
    // ✅ Named ranges (simple A1-style, workbook and sheet scoped) - IMPLEMENTED
    // ❌ Formulas (parsing and tokenization) - Future enhancement
}

impl Default for XlsWriter {
    fn default() -> Self {
        Self::new()
    }
}

/// Implementation notes for BIFF record generation:
///
/// All core BIFF8 records have been implemented in the `biff` module:
/// - ✅ write_bof() - Beginning of File (0x0809)
/// - ✅ write_eof() - End of File (0x000A)
/// - ✅ write_codepage() - Code page (0x0042)
/// - ✅ write_date1904() - Date system (0x0022)
/// - ✅ write_window1() - Workbook window properties (0x003D)
/// - ✅ write_boundsheet() - Sheet metadata (0x0085)
/// - ✅ write_dimensions() - Worksheet dimensions (0x0200)
/// - ✅ write_sst() - Shared string table with CONTINUE support (0x00FC)
/// - ✅ write_number() - Floating point cell (0x0203)
/// - ✅ write_labelsst() - String cell (0x00FD)
/// - ✅ write_boolerr() - Boolean/error cell (0x0205)
/// - ✅ write_continue() - Continuation record (0x003C)
///
/// Future enhancements:
/// - FORMULA record (0x0006) - For formula cells with RPN tokens
/// - XF records (0x00E0) - For cell formatting
/// - FONT records (0x0031) - For font definitions
/// - FORMAT records (0x041E) - For number formats
/// - COLINFO records (0x007D) - For column widths
/// - ROW records (0x0208) - For row heights
/// - MERGEDCELLS records (0x00E5) - For merged cell ranges
/// - NAME records (0x0018) - For named ranges
#[cfg(test)]
mod tests {
    use super::*;
    use std::io::Cursor;

    #[test]
    fn test_create_writer() {
        let writer = XlsWriter::new();
        assert_eq!(writer.worksheets.len(), 0);
        assert_eq!(writer.shared_strings.len(), 0);
    }

    #[test]
    fn test_add_worksheet() {
        let mut writer = XlsWriter::new();
        let idx = writer.add_worksheet("Sheet1").unwrap();
        assert_eq!(idx, 0);
        assert_eq!(writer.worksheets.len(), 1);
        assert_eq!(writer.worksheets[0].name, "Sheet1");
    }

    #[test]
    fn test_add_multiple_worksheets() {
        let mut writer = XlsWriter::new();
        let idx1 = writer.add_worksheet("Sheet1").unwrap();
        let idx2 = writer.add_worksheet("Sheet2").unwrap();
        let idx3 = writer.add_worksheet("Sheet3").unwrap();

        assert_eq!(idx1, 0);
        assert_eq!(idx2, 1);
        assert_eq!(idx3, 2);
        assert_eq!(writer.worksheets.len(), 3);
    }

    #[test]
    fn test_add_worksheet_empty_name() {
        let mut writer = XlsWriter::new();
        let result = writer.add_worksheet("");
        assert!(result.is_err());
    }

    #[test]
    fn test_add_worksheet_long_name() {
        let mut writer = XlsWriter::new();
        let long_name = "A".repeat(50);
        let result = writer.add_worksheet(&long_name);
        assert!(result.is_err()); // Name too long
    }

    #[test]
    fn test_add_worksheet_duplicate_name() {
        let mut writer = XlsWriter::new();
        writer.add_worksheet("Sheet1").unwrap();
        let result = writer.add_worksheet("Sheet1");
        assert!(result.is_err());
    }

    #[test]
    fn test_write_string() {
        let mut writer = XlsWriter::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_string(sheet, 0, 0, "Hello").unwrap();
        assert_eq!(writer.worksheets[0].cells.len(), 1);

        let cell = writer.worksheets[0].cells.get(&(0, 0)).unwrap();
        assert_eq!(cell.row, 0);
        assert_eq!(cell.col, 0);
        assert!(matches!(&cell.value, XlsCellValue::String(s) if s == "Hello"));
    }

    #[test]
    fn test_write_number() {
        let mut writer = XlsWriter::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_number(sheet, 0, 0, 42.5).unwrap();
        assert_eq!(writer.worksheets[0].cells.len(), 1);

        let cell = writer.worksheets[0].cells.get(&(0, 0)).unwrap();
        assert!(matches!(&cell.value, XlsCellValue::Number(n) if *n == 42.5));
    }

    #[test]
    fn test_write_boolean() {
        let mut writer = XlsWriter::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_boolean(sheet, 0, 0, true).unwrap();
        writer.write_boolean(sheet, 1, 0, false).unwrap();

        assert_eq!(writer.worksheets[0].cells.len(), 2);
        assert!(matches!(
            writer.worksheets[0].cells.get(&(0, 0)).unwrap().value,
            XlsCellValue::Boolean(true)
        ));
        assert!(matches!(
            writer.worksheets[0].cells.get(&(1, 0)).unwrap().value,
            XlsCellValue::Boolean(false)
        ));
    }

    #[test]
    fn test_write_formula() {
        let mut writer = XlsWriter::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_formula(sheet, 0, 0, "SUM(A1:B1)").unwrap();

        let cell = writer.worksheets[0].cells.get(&(0, 0)).unwrap();
        assert!(matches!(&cell.value, XlsCellValue::Formula(f) if f == "SUM(A1:B1)"));
    }

    #[test]
    fn test_write_multiple_cells() {
        let mut writer = XlsWriter::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();

        writer.write_string(sheet, 0, 0, "A1").unwrap();
        writer.write_string(sheet, 0, 1, "B1").unwrap();
        writer.write_string(sheet, 1, 0, "A2").unwrap();
        writer.write_string(sheet, 1, 1, "B2").unwrap();

        assert_eq!(writer.worksheets[0].cells.len(), 4);
    }

    #[test]
    fn test_shared_strings_build() {
        let mut writer = XlsWriter::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();

        writer.write_string(sheet, 0, 0, "Hello").unwrap();
        writer.write_string(sheet, 0, 1, "Hello").unwrap();
        writer.write_string(sheet, 1, 0, "World").unwrap();

        // Build shared strings table (normally done during write)
        writer.build_shared_strings();

        // Should only have 2 unique strings
        assert_eq!(writer.shared_strings.len(), 2);
    }

    #[test]
    fn test_write_to_memory() {
        let mut writer = XlsWriter::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_string(sheet, 0, 0, "Test").unwrap();
        writer.write_number(sheet, 0, 1, 123.45).unwrap();

        let mut cursor = Cursor::new(Vec::new());
        let result = writer.write_to(&mut cursor);
        assert!(result.is_ok());

        let data = cursor.into_inner();
        assert!(!data.is_empty());
        // Should start with OLE compound document signature
        assert_eq!(
            &data[0..8],
            [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1]
        );
    }

    #[test]
    fn test_save_to_file() {
        let mut writer = XlsWriter::new();
        let sheet = writer.add_worksheet("Sheet1").unwrap();
        writer.write_string(sheet, 0, 0, "Hello").unwrap();

        let temp_path = std::env::temp_dir().join("test_xls_writer.xls");
        let result = writer.save(&temp_path);
        assert!(result.is_ok());

        // Verify file was created
        assert!(temp_path.exists());

        // Clean up
        let _ = std::fs::remove_file(temp_path);
    }

    #[test]
    fn test_xls_writer_default() {
        let writer: XlsWriter = Default::default();
        assert_eq!(writer.worksheets.len(), 0);
        assert_eq!(writer.shared_strings.len(), 0);
    }

    #[test]
    fn test_xlscellvalue_variants() {
        let string_val = XlsCellValue::String("test".to_string());
        let number_val = XlsCellValue::Number(42.0);
        let bool_val = XlsCellValue::Boolean(true);
        let formula_val = XlsCellValue::Formula("A1+B1".to_string());
        let blank_val = XlsCellValue::Blank;

        assert!(matches!(string_val, XlsCellValue::String(_)));
        assert!(matches!(number_val, XlsCellValue::Number(_)));
        assert!(matches!(bool_val, XlsCellValue::Boolean(_)));
        assert!(matches!(formula_val, XlsCellValue::Formula(_)));
        assert!(matches!(blank_val, XlsCellValue::Blank));
    }

    #[test]
    fn test_xlscellvalue_debug() {
        let val = XlsCellValue::String("test".to_string());
        let debug = format!("{:?}", val);
        assert!(debug.contains("String"));
    }

    #[test]
    fn test_xlscellvalue_clone() {
        let val = XlsCellValue::Number(42.0);
        let cloned = val.clone();
        assert!(matches!(cloned, XlsCellValue::Number(42.0)));
    }

    #[test]
    fn test_writablecell_creation() {
        let cell = WritableCell {
            row: 5,
            col: 3,
            value: XlsCellValue::String("Test".to_string()),
            format_idx: 15,
            pivot_xf_role: None,
        };

        assert_eq!(cell.row, 5);
        assert_eq!(cell.col, 3);
        assert_eq!(cell.format_idx, 15);
    }

    #[test]
    fn test_writableworksheet_creation() {
        let ws = WritableWorksheet::new("TestSheet".to_string());
        assert_eq!(ws.name, "TestSheet");
        assert!(ws.cells.is_empty());
        assert!(ws.merged_ranges.is_empty());
        assert!(ws.column_widths.is_empty());
    }

    #[test]
    fn test_writableworksheet_add_cell() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        let cell = WritableCell {
            row: 0,
            col: 0,
            value: XlsCellValue::Number(100.0),
            format_idx: 0,
            pivot_xf_role: None,
        };
        ws.add_cell(cell);
        assert_eq!(ws.cells.len(), 1);
    }

    #[test]
    fn test_writableworksheet_set_column_width() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        ws.set_column_width(0, 2560); // ~10 characters
        assert_eq!(ws.column_widths.get(&0), Some(&2560));
    }

    #[test]
    fn test_writableworksheet_merge_cells() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        ws.add_merged_range(super::MergedRange {
            first_row: 0,
            last_row: 1,
            first_col: 0,
            last_col: 2,
        }); // Merge A1:C2
        assert_eq!(ws.merged_ranges.len(), 1);
        assert_eq!(ws.merged_ranges[0].first_row, 0);
        assert_eq!(ws.merged_ranges[0].last_row, 1);
        assert_eq!(ws.merged_ranges[0].first_col, 0);
        assert_eq!(ws.merged_ranges[0].last_col, 2);
    }

    #[test]
    fn test_writableworksheet_freeze_panes() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        assert!(ws.freeze_panes.is_none());
        ws.set_freeze_panes(1, 2);
        assert!(ws.freeze_panes.is_some());
        let fp = ws.freeze_panes.unwrap();
        assert_eq!(fp.freeze_rows, 1);
        assert_eq!(fp.freeze_cols, 2);
    }

    #[test]
    fn test_writableworksheet_add_conditional_format() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        let cf = XlsConditionalFormat {
            first_row: 0,
            last_row: 10,
            first_col: 0,
            last_col: 0,
            format_type: super::XlsConditionalFormatType::Formula {
                formula: "A1>100".to_string(),
            },
            pattern: None,
        };
        ws.add_conditional_format(cf);
        assert_eq!(ws.conditional_formats.len(), 1);
    }

    #[test]
    fn test_writableworksheet_add_data_validation() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        let dv = XlsDataValidation {
            first_row: 0,
            last_row: 10,
            first_col: 0,
            last_col: 0,
            validation_type: super::XlsDataValidationType::List {
                values: vec!["Option1".to_string(), "Option2".to_string()],
            },
            show_input_message: true,
            input_title: None,
            input_message: None,
            show_error_alert: true,
            error_title: None,
            error_message: None,
        };
        ws.add_data_validation(dv);
        assert_eq!(ws.data_validations.len(), 1);
    }

    #[test]
    fn test_writableworksheet_add_hyperlink() {
        let mut ws = WritableWorksheet::new("Sheet1".to_string());
        let link = super::XlsHyperlink {
            first_row: 0,
            last_row: 0,
            first_col: 0,
            last_col: 0,
            url: "https://example.com".to_string(),
        };
        ws.add_hyperlink(link);
        assert_eq!(ws.hyperlinks.len(), 1);
        assert_eq!(ws.hyperlinks[0].url, "https://example.com");
    }

    #[test]
    fn test_xls_defined_name_basic() {
        let name = XlsDefinedName {
            name: "TestRange".to_string(),
            reference: "A1:B10".to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: Some(0),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        };
        assert_eq!(name.name, "TestRange");
        assert_eq!(name.reference, "A1:B10");
        assert_eq!(name.target_sheet, Some(0));
    }

    #[test]
    fn test_xls_defined_name_to_biff_formula_area() {
        let name = XlsDefinedName {
            name: "TestRange".to_string(),
            reference: "A1:B10".to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: Some(0),
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        };
        let formula = name.to_biff_formula().unwrap();
        assert!(!formula.is_empty());
    }

    #[test]
    fn test_xls_defined_name_to_biff_formula_single() {
        let name = XlsDefinedName {
            name: "SingleCell".to_string(),
            reference: "C5".to_string(),
            comment: None,
            local_sheet: None,
            target_sheet: None,
            hidden: false,
            is_function: false,
            is_built_in: false,
            built_in_code: None,
        };
        let formula = name.to_biff_formula().unwrap();
        assert!(!formula.is_empty());
    }
}
