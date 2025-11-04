//! Worksheet implementation for Excel files.
//!
//! This module provides the concrete implementation of worksheets
//! for Excel (.xlsx) files.

use std::borrow::Cow;
use std::collections::HashMap;

use crate::ooxml::opc::PackURI;
use crate::sheet::{
    Cell as CellTrait, CellIterator, CellValue, Result as SheetResult, RowIterator,
    Worksheet as WorksheetTrait,
};

use super::cell::{Cell, CellIterator as XlsxCellIterator, RowIterator as XlsxRowIterator};
use super::format::{CellBorder, CellFill, CellFont, CellFormat};

/// Information about a worksheet
#[derive(Debug, Clone)]
pub struct WorksheetInfo {
    /// Worksheet name
    pub name: String,
    /// Relationship ID for the worksheet
    pub relationship_id: String,
    /// Sheet ID
    pub sheet_id: u32,
    /// Whether this is the active sheet
    pub is_active: bool,
}

/// Column information
#[derive(Debug, Clone)]
pub struct ColumnInfo {
    /// Column width (in character units)
    pub width: Option<f64>,
    /// Whether the column is hidden
    pub hidden: bool,
    /// Custom width set
    pub custom_width: bool,
}

/// Row information
#[derive(Debug, Clone)]
pub struct RowInfo {
    /// Row height (in points)
    pub height: Option<f64>,
    /// Whether the row is hidden
    pub hidden: bool,
    /// Custom height set
    pub custom_height: bool,
}

/// Hyperlink information
#[derive(Debug, Clone)]
pub struct Hyperlink {
    /// Cell reference (e.g., "A1")
    pub cell_ref: String,
    /// Target URL or reference
    pub target: String,
    /// Display text (tooltip)
    pub display: Option<String>,
}

/// Cell comment information
#[derive(Debug, Clone)]
pub struct Comment {
    /// Cell reference (e.g., "A1")
    pub cell_ref: String,
    /// Author of the comment
    pub author: Option<String>,
    /// Comment text
    pub text: String,
}

/// Data validation rule information (parsed from worksheet XML)
#[derive(Debug, Clone)]
pub struct DataValidationRule {
    /// Range (e.g., "A1:B10")
    pub range: String,
    /// Validation type (e.g., "list", "whole", "decimal")
    pub validation_type: String,
    /// Allowed values or formula
    pub formula: Option<String>,
}

/// Conditional formatting rule
#[derive(Debug, Clone)]
pub struct ConditionalFormatRule {
    /// Range (e.g., "A1:B10")
    pub range: String,
    /// Rule type (e.g., "cellIs", "colorScale")
    pub rule_type: String,
    /// Priority
    pub priority: u32,
}

/// Page setup information
#[derive(Debug, Clone, Default)]
pub struct PageSetup {
    /// Paper size (e.g., 9 for A4)
    pub paper_size: Option<u32>,
    /// Orientation (true = landscape, false = portrait)
    pub landscape: bool,
    /// Scale percentage
    pub scale: Option<u32>,
    /// Fit to page width
    pub fit_to_width: Option<u32>,
    /// Fit to page height
    pub fit_to_height: Option<u32>,
}

/// Auto-filter information
#[derive(Debug, Clone)]
pub struct AutoFilter {
    /// Range (e.g., "A1:D10")
    pub range: String,
}

/// Concrete implementation of the Worksheet trait for Excel files.
pub struct Worksheet<'a> {
    /// Reference to the parent workbook
    workbook: &'a Workbook,
    /// Worksheet information
    info: WorksheetInfo,
    /// Cached cell data (row -> column -> value)
    cells: HashMap<u32, HashMap<u32, CellValue>>,
    /// Cell style indices (row -> column -> style_index)
    cell_styles: HashMap<u32, HashMap<u32, u32>>,
    /// Dimensions of the worksheet (min_row, min_col, max_row, max_col)
    dimensions: Option<(u32, u32, u32, u32)>,
    /// Merged cell ranges (start_row, start_col, end_row, end_col)
    merged_regions: Vec<(u32, u32, u32, u32)>,
    /// Hyperlinks by cell reference
    hyperlinks: HashMap<String, Hyperlink>,
    /// Comments by cell reference
    comments: HashMap<String, Comment>,
    /// Column information by column number
    columns: HashMap<u32, ColumnInfo>,
    /// Row information by row number
    rows: HashMap<u32, RowInfo>,
    /// Data validations
    data_validations: Vec<DataValidationRule>,
    /// Conditional formatting rules
    conditional_formats: Vec<ConditionalFormatRule>,
    /// Page setup
    page_setup: PageSetup,
    /// Auto-filter
    auto_filter: Option<AutoFilter>,
}

impl<'a> Worksheet<'a> {
    /// Create a new worksheet.
    pub fn new(workbook: &'a Workbook, info: WorksheetInfo) -> Self {
        Self {
            workbook,
            info,
            cells: HashMap::new(),
            cell_styles: HashMap::new(),
            dimensions: None,
            merged_regions: Vec::new(),
            hyperlinks: HashMap::new(),
            comments: HashMap::new(),
            columns: HashMap::new(),
            rows: HashMap::new(),
            data_validations: Vec::new(),
            conditional_formats: Vec::new(),
            page_setup: PageSetup::default(),
            auto_filter: None,
        }
    }

    /// Load worksheet data from the XML.
    pub fn load_data(&mut self) -> SheetResult<()> {
        // Get the worksheet part using the relationship ID
        let worksheet_uri =
            PackURI::new(format!("/xl/worksheets/sheet{}.xml", self.info.sheet_id))?;

        let worksheet_part = self.workbook.package().get_part(&worksheet_uri)?;
        let content = std::str::from_utf8(worksheet_part.blob())?;

        // Parse worksheet data
        self.parse_worksheet_xml(content)?;

        Ok(())
    }

    /// Parse worksheet XML to extract cell data.
    fn parse_worksheet_xml(&mut self, content: &str) -> SheetResult<()> {
        // Parse sheetData section (cells)
        if let Some(sheet_data_start) = content.find("<sheetData>")
            && let Some(sheet_data_end) = content[sheet_data_start..].find("</sheetData>")
        {
            let sheet_data_content = &content[sheet_data_start..sheet_data_start + sheet_data_end];
            self.parse_sheet_data(sheet_data_content)?;
        }

        // Parse merged cells
        if let Some(merge_start) = content.find("<mergeCells")
            && let Some(merge_end) = content[merge_start..].find("</mergeCells>")
        {
            let merge_content = &content[merge_start..merge_start + merge_end + 13];
            self.parse_merged_cells(merge_content)?;
        }

        // Parse hyperlinks
        if let Some(hyperlink_start) = content.find("<hyperlinks>")
            && let Some(hyperlink_end) = content[hyperlink_start..].find("</hyperlinks>")
        {
            let hyperlink_content = &content[hyperlink_start..hyperlink_start + hyperlink_end];
            self.parse_hyperlinks(hyperlink_content)?;
        }

        // Parse column information
        if let Some(cols_start) = content.find("<cols>")
            && let Some(cols_end) = content[cols_start..].find("</cols>")
        {
            let cols_content = &content[cols_start..cols_start + cols_end];
            self.parse_columns(cols_content)?;
        }

        // Parse data validations
        if let Some(dv_start) = content.find("<dataValidations")
            && let Some(dv_end) = content[dv_start..].find("</dataValidations>")
        {
            let dv_content = &content[dv_start..dv_start + dv_end + 18];
            self.parse_data_validations(dv_content)?;
        }

        // Parse conditional formatting
        if let Some(cf_start) = content.find("<conditionalFormatting")
            && let Some(cf_end) = content[cf_start..].find("</conditionalFormatting>")
        {
            let cf_content = &content[cf_start..cf_start + cf_end + 24];
            self.parse_conditional_formatting(cf_content)?;
        }

        // Parse page setup
        if let Some(ps_start) = content.find("<pageSetup ")
            && let Some(ps_end) = content[ps_start..].find("/>")
        {
            let ps_content = &content[ps_start..ps_start + ps_end + 2];
            self.parse_page_setup(ps_content)?;
        }

        // Parse auto-filter
        if let Some(af_start) = content.find("<autoFilter ")
            && let Some(af_end) = content[af_start..].find("/>")
        {
            let af_content = &content[af_start..af_start + af_end + 2];
            self.parse_auto_filter(af_content)?;
        }

        Ok(())
    }

    /// Parse sheetData content.
    fn parse_sheet_data(&mut self, sheet_data: &str) -> SheetResult<()> {
        let mut pos = 0;
        let mut min_row = u32::MAX;
        let mut max_row = 0;
        let mut min_col = u32::MAX;
        let mut max_col = 0;

        while let Some(row_start) = sheet_data[pos..].find("<row ") {
            let row_start_pos = pos + row_start;
            if let Some(row_end) = sheet_data[row_start_pos..].find("</row>") {
                let row_content = &sheet_data[row_start_pos..row_start_pos + row_end + 6];

                if let Some((row_num, row_info, cells)) = self.parse_row_xml(row_content)? {
                    min_row = min_row.min(row_num);
                    max_row = max_row.max(row_num);

                    // Store row information if it has custom properties
                    if let Some(info) = row_info {
                        self.rows.insert(row_num, info);
                    }

                    for (col_num, value, style_idx) in cells {
                        min_col = min_col.min(col_num);
                        max_col = max_col.max(col_num);

                        self.cells
                            .entry(row_num)
                            .or_default()
                            .insert(col_num, value);

                        if let Some(idx) = style_idx {
                            self.cell_styles
                                .entry(row_num)
                                .or_default()
                                .insert(col_num, idx);
                        }
                    }
                }

                pos = row_start_pos + row_end + 6;
            } else {
                break;
            }
        }

        if min_row <= max_row && min_col <= max_col {
            self.dimensions = Some((min_row, min_col, max_row, max_col));
        }

        Ok(())
    }

    /// Parse a single row XML.
    #[allow(clippy::type_complexity)]
    fn parse_row_xml(
        &self,
        row_content: &str,
    ) -> SheetResult<Option<(u32, Option<RowInfo>, Vec<(u32, CellValue, Option<u32>)>)>> {
        // Extract row number
        let row_num = if let Some(r_start) = row_content.find("r=\"") {
            let r_content = &row_content[r_start + 3..];
            if let Some(quote_pos) = r_content.find('"') {
                r_content[..quote_pos].parse::<u32>().ok()
            } else {
                None
            }
        } else {
            None
        };

        let row_num = match row_num {
            Some(r) => r,
            None => return Ok(None),
        };

        // Extract row height and hidden status
        let height = if let Some(ht_start) = row_content.find("ht=\"") {
            let ht_content = &row_content[ht_start + 4..];
            ht_content
                .find('"')
                .and_then(|quote_pos| ht_content[..quote_pos].parse::<f64>().ok())
        } else {
            None
        };

        let hidden = row_content.contains("hidden=\"1\"");
        let custom_height = row_content.contains("customHeight=\"1\"");

        let row_info = if height.is_some() || hidden || custom_height {
            Some(RowInfo {
                height,
                hidden,
                custom_height,
            })
        } else {
            None
        };

        let mut cells = Vec::new();

        // Parse cells in this row
        let mut pos = 0;
        while let Some(c_start) = row_content[pos..].find("<c ") {
            let c_start_pos = pos + c_start;
            if let Some(c_end) = row_content[c_start_pos..].find("</c>") {
                let c_content = &row_content[c_start_pos..c_start_pos + c_end + 4];

                if let Some((col_num, value, style_idx)) = self.parse_cell_xml(c_content)? {
                    cells.push((col_num, value, style_idx));
                }

                pos = c_start_pos + c_end + 4;
            } else {
                break;
            }
        }

        Ok(Some((row_num, row_info, cells)))
    }

    /// Parse a single cell XML.
    fn parse_cell_xml(
        &self,
        cell_content: &str,
    ) -> SheetResult<Option<(u32, CellValue, Option<u32>)>> {
        // Extract cell reference (e.g., "A1")
        let reference = if let Some(r_start) = cell_content.find("r=\"") {
            let r_content = &cell_content[r_start + 3..];
            r_content
                .find('"')
                .map(|quote_pos| r_content[..quote_pos].to_string())
        } else {
            None
        };

        let reference = match reference {
            Some(r) => r,
            None => return Ok(None),
        };

        // Convert reference to row/col numbers
        let (col_num, _row_num) = Cell::reference_to_coords(&reference)?;

        // Extract style index (s attribute)
        let style_idx = if let Some(s_start) = cell_content.find(" s=\"") {
            let s_content = &cell_content[s_start + 4..];
            s_content
                .find('"')
                .and_then(|quote_pos| s_content[..quote_pos].parse::<u32>().ok())
        } else {
            None
        };

        // Extract cell type
        let cell_type = if let Some(t_start) = cell_content.find("t=\"") {
            let t_content = &cell_content[t_start + 3..];
            t_content
                .find('"')
                .map(|quote_pos| t_content[..quote_pos].to_string())
        } else {
            None
        };

        // Extract value
        let value = if let Some(v_start) = cell_content.find("<v>") {
            let v_start_pos = v_start + 3;
            cell_content[v_start_pos..]
                .find("</v>")
                .map(|v_end| cell_content[v_start_pos..v_start_pos + v_end].to_string())
        } else {
            None
        };

        let cell_value = match (cell_type.as_deref(), value.as_deref()) {
            (Some("str"), Some(v)) => CellValue::String(v.to_string()),
            (Some("s"), Some(v)) => {
                // Shared string reference - parse index and resolve later
                CellValue::String(format!("SHARED_STRING_{}", v))
            },
            (Some("b"), Some(v)) => match v {
                "1" => CellValue::Bool(true),
                "0" => CellValue::Bool(false),
                _ => CellValue::Error("Invalid boolean value".to_string()),
            },
            (_, Some(v)) => {
                // Try to parse as number - use fast parsing
                if let Ok(int_val) = atoi_simd::parse(v.as_bytes()) {
                    CellValue::Int(int_val)
                } else if let Ok(float_val) = fast_float2::parse(v) {
                    CellValue::Float(float_val)
                } else {
                    CellValue::String(v.to_string())
                }
            },
            _ => CellValue::Empty,
        };

        Ok(Some((col_num, cell_value, style_idx)))
    }

    /// Parse merged cells from XML.
    fn parse_merged_cells(&mut self, content: &str) -> SheetResult<()> {
        let mut pos = 0;
        while let Some(merge_start) = content[pos..].find("<mergeCell ") {
            let merge_start_pos = pos + merge_start;
            if let Some(merge_end) = content[merge_start_pos..].find("/>") {
                let merge_tag = &content[merge_start_pos..merge_start_pos + merge_end + 2];

                // Extract ref attribute (e.g., "A1:B2")
                if let Some(ref_start) = merge_tag.find("ref=\"")
                    && let Some(ref_end) = merge_tag[ref_start + 5..].find('"')
                {
                    let range_ref = &merge_tag[ref_start + 5..ref_start + 5 + ref_end];
                    if let Some(colon_pos) = range_ref.find(':') {
                        let start_ref = &range_ref[..colon_pos];
                        let end_ref = &range_ref[colon_pos + 1..];

                        if let Ok((start_col, start_row)) = Cell::reference_to_coords(start_ref)
                            && let Ok((end_col, end_row)) = Cell::reference_to_coords(end_ref)
                        {
                            self.merged_regions
                                .push((start_row, start_col, end_row, end_col));
                        }
                    }
                }

                pos = merge_start_pos + merge_end + 2;
            } else {
                break;
            }
        }
        Ok(())
    }

    /// Parse hyperlinks from XML.
    fn parse_hyperlinks(&mut self, content: &str) -> SheetResult<()> {
        let mut pos = 0;
        while let Some(hyperlink_start) = content[pos..].find("<hyperlink ") {
            let hyperlink_start_pos = pos + hyperlink_start;
            if let Some(hyperlink_end) = content[hyperlink_start_pos..].find("/>") {
                let hyperlink_tag =
                    &content[hyperlink_start_pos..hyperlink_start_pos + hyperlink_end + 2];

                let cell_ref = Self::extract_attribute(hyperlink_tag, "ref");
                let r_id = Self::extract_attribute(hyperlink_tag, "r:id");
                let display = Self::extract_attribute(hyperlink_tag, "display");

                if let Some(ref_val) = cell_ref {
                    self.hyperlinks.insert(
                        ref_val.clone(),
                        Hyperlink {
                            cell_ref: ref_val,
                            target: r_id.unwrap_or_else(|| String::from("")),
                            display,
                        },
                    );
                }

                pos = hyperlink_start_pos + hyperlink_end + 2;
            } else {
                break;
            }
        }
        Ok(())
    }

    /// Parse column information from XML.
    fn parse_columns(&mut self, content: &str) -> SheetResult<()> {
        let mut pos = 0;
        while let Some(col_start) = content[pos..].find("<col ") {
            let col_start_pos = pos + col_start;
            if let Some(col_end) = content[col_start_pos..].find("/>") {
                let col_tag = &content[col_start_pos..col_start_pos + col_end + 2];

                let min_col =
                    Self::extract_attribute(col_tag, "min").and_then(|s| s.parse::<u32>().ok());
                let max_col =
                    Self::extract_attribute(col_tag, "max").and_then(|s| s.parse::<u32>().ok());
                let width =
                    Self::extract_attribute(col_tag, "width").and_then(|s| s.parse::<f64>().ok());
                let hidden = col_tag.contains("hidden=\"1\"");
                let custom_width = col_tag.contains("customWidth=\"1\"");

                if let (Some(min), Some(max)) = (min_col, max_col) {
                    let col_info = ColumnInfo {
                        width,
                        hidden,
                        custom_width,
                    };
                    for col_num in min..=max {
                        self.columns.insert(col_num, col_info.clone());
                    }
                }

                pos = col_start_pos + col_end + 2;
            } else {
                break;
            }
        }
        Ok(())
    }

    /// Parse data validations from XML.
    fn parse_data_validations(&mut self, content: &str) -> SheetResult<()> {
        let mut pos = 0;
        while let Some(dv_start) = content[pos..].find("<dataValidation ") {
            let dv_start_pos = pos + dv_start;
            if let Some(dv_end) = content[dv_start_pos..].find("</dataValidation>") {
                let dv_tag = &content[dv_start_pos..dv_start_pos + dv_end + 17];

                let range = Self::extract_attribute(dv_tag, "sqref");
                let validation_type = Self::extract_attribute(dv_tag, "type");
                let formula = if let Some(formula_start) = dv_tag.find("<formula1>")
                    && let Some(formula_end) = dv_tag[formula_start..].find("</formula1>")
                {
                    Some(dv_tag[formula_start + 10..formula_start + formula_end].to_string())
                } else {
                    None
                };

                if let Some(range_val) = range {
                    self.data_validations.push(DataValidationRule {
                        range: range_val,
                        validation_type: validation_type.unwrap_or_else(|| String::from("list")),
                        formula,
                    });
                }

                pos = dv_start_pos + dv_end + 17;
            } else {
                break;
            }
        }
        Ok(())
    }

    /// Parse conditional formatting from XML.
    fn parse_conditional_formatting(&mut self, content: &str) -> SheetResult<()> {
        // Extract sqref attribute for the range
        let range = Self::extract_attribute(content, "sqref");

        if let Some(range_val) = range {
            let mut pos = 0;
            while let Some(rule_start) = content[pos..].find("<cfRule ") {
                let rule_start_pos = pos + rule_start;
                if let Some(rule_end) = content[rule_start_pos..].find("/>") {
                    let rule_tag = &content[rule_start_pos..rule_start_pos + rule_end + 2];

                    let rule_type = Self::extract_attribute(rule_tag, "type");
                    let priority = Self::extract_attribute(rule_tag, "priority")
                        .and_then(|s| s.parse::<u32>().ok());

                    if let (Some(type_val), Some(priority_val)) = (rule_type, priority) {
                        self.conditional_formats.push(ConditionalFormatRule {
                            range: range_val.clone(),
                            rule_type: type_val,
                            priority: priority_val,
                        });
                    }

                    pos = rule_start_pos + rule_end + 2;
                } else {
                    break;
                }
            }
        }

        Ok(())
    }

    /// Parse page setup from XML.
    fn parse_page_setup(&mut self, content: &str) -> SheetResult<()> {
        let paper_size =
            Self::extract_attribute(content, "paperSize").and_then(|s| s.parse::<u32>().ok());
        let landscape = content.contains("orientation=\"landscape\"");
        let scale = Self::extract_attribute(content, "scale").and_then(|s| s.parse::<u32>().ok());
        let fit_to_width =
            Self::extract_attribute(content, "fitToWidth").and_then(|s| s.parse::<u32>().ok());
        let fit_to_height =
            Self::extract_attribute(content, "fitToHeight").and_then(|s| s.parse::<u32>().ok());

        self.page_setup = PageSetup {
            paper_size,
            landscape,
            scale,
            fit_to_width,
            fit_to_height,
        };

        Ok(())
    }

    /// Parse auto-filter from XML.
    fn parse_auto_filter(&mut self, content: &str) -> SheetResult<()> {
        if let Some(range) = Self::extract_attribute(content, "ref") {
            self.auto_filter = Some(AutoFilter { range });
        }
        Ok(())
    }

    /// Helper method to extract attribute value from XML tag.
    fn extract_attribute(tag: &str, attr: &str) -> Option<String> {
        let search_str = format!("{}=\"", attr);
        if let Some(start) = tag.find(&search_str) {
            let value_start = start + search_str.len();
            tag[value_start..]
                .find('"')
                .map(|end| tag[value_start..value_start + end].to_string())
        } else {
            None
        }
    }

    /// Get cell value at specific coordinates.
    fn get_cell_value(&self, row: u32, col: u32) -> CellValue {
        match self.cells.get(&row).and_then(|row_data| row_data.get(&col)) {
            Some(cell_value) => self.resolve_shared_string(cell_value.clone()),
            None => CellValue::Empty,
        }
    }

    /// Resolve shared string references to actual string values.
    fn resolve_shared_string(&self, cell_value: CellValue) -> CellValue {
        match cell_value {
            CellValue::String(s) if s.starts_with("SHARED_STRING_") => {
                // Extract the index from the shared string reference
                if let Some(index_str) = s.strip_prefix("SHARED_STRING_")
                    && let Ok(index) = atoi_simd::parse(index_str.as_bytes())
                    && let Some(shared_string) = self.workbook.shared_strings().get(index)
                {
                    return CellValue::String(shared_string.to_string());
                }
                CellValue::Error("Invalid shared string reference".to_string())
            },
            other => other,
        }
    }

    /// Get all cells in a specific column.
    ///
    /// # Arguments
    /// * `column` - Column number (1-based)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// // Get all values in column A (column 1)
    /// let column_values = ws.column_values(1)?;
    /// for value in column_values {
    ///     println!("{:?}", value);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn column_values(&self, column: u32) -> SheetResult<Vec<CellValue>> {
        let mut values = Vec::new();

        if let Some((min_row, _, max_row, _)) = self.dimensions {
            for row in min_row..=max_row {
                values.push(self.get_cell_value(row, column));
            }
        }

        Ok(values)
    }

    /// Get all cells in a specific row.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// // Get all values in row 1
    /// let row_values = ws.row_values(1)?;
    /// for value in row_values {
    ///     println!("{:?}", value);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn row_values(&self, row: u32) -> SheetResult<Vec<CellValue>> {
        let mut values = Vec::new();

        if let Some((_, min_col, _, max_col)) = self.dimensions {
            for col in min_col..=max_col {
                values.push(self.get_cell_value(row, col));
            }
        }

        Ok(values)
    }

    /// Get a range of cells as a 2D vector.
    ///
    /// # Arguments
    /// * `start_row` - Starting row (1-based, inclusive)
    /// * `start_col` - Starting column (1-based, inclusive)
    /// * `end_row` - Ending row (1-based, inclusive)
    /// * `end_col` - Ending column (1-based, inclusive)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// // Get range A1:C3
    /// let range = ws.range(1, 1, 3, 3)?;
    /// for row in range {
    ///     for cell in row {
    ///         print!("{:?} ", cell);
    ///     }
    ///     println!();
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn range(
        &self,
        start_row: u32,
        start_col: u32,
        end_row: u32,
        end_col: u32,
    ) -> SheetResult<Vec<Vec<CellValue>>> {
        let mut result = Vec::new();

        for row in start_row..=end_row {
            let mut row_data = Vec::new();
            for col in start_col..=end_col {
                row_data.push(self.get_cell_value(row, col));
            }
            result.push(row_data);
        }

        Ok(result)
    }

    /// Find cells containing specific text.
    ///
    /// # Arguments
    /// * `query` - Text to search for
    ///
    /// # Returns
    /// Vector of (row, column) tuples where the cell contains the query text
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// // Find all cells containing "Total"
    /// let matches = ws.find_text("Total")?;
    /// for (row, col) in matches {
    ///     println!("Found at row {}, column {}", row, col);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn find_text(&self, query: &str) -> SheetResult<Vec<(u32, u32)>> {
        let mut matches = Vec::new();

        for (&row, row_data) in &self.cells {
            for (&col, value) in row_data {
                if let CellValue::String(s) = &value
                    && (s.contains(query)
                        || (s.starts_with("SHARED_STRING_") && {
                            let resolved = self.resolve_shared_string(value.clone());
                            matches!(resolved, CellValue::String(ref text) if text.contains(query))
                        }))
                {
                    matches.push((row, col));
                }
            }
        }

        Ok(matches)
    }

    /// Get the used range dimensions (min row, min col, max row, max col).
    ///
    /// Returns None if the worksheet is empty.
    pub fn used_range(&self) -> Option<(u32, u32, u32, u32)> {
        self.dimensions
    }

    /// Check if a cell is empty.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn is_cell_empty(&self, row: u32, column: u32) -> bool {
        matches!(self.get_cell_value(row, column), CellValue::Empty)
    }

    /// Count non-empty cells in the worksheet.
    pub fn non_empty_cell_count(&self) -> usize {
        self.cells
            .values()
            .map(|row| {
                row.values()
                    .filter(|v| !matches!(v, CellValue::Empty))
                    .count()
            })
            .sum()
    }

    /// Get worksheet information.
    pub fn info(&self) -> &WorksheetInfo {
        &self.info
    }

    // ===== Cell Formatting (Reading) =====

    /// Get the cell style for a specific cell.
    ///
    /// Returns the style information including font, fill, border, and number format.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// if let Some(style) = ws.get_cell_style(1, 1) {
    ///     println!("Cell has custom styling");
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_cell_style(&self, row: u32, column: u32) -> Option<&super::styles::CellStyle> {
        self.cell_styles
            .get(&row)
            .and_then(|row_styles| row_styles.get(&column))
            .and_then(|style_idx| self.workbook.styles().get_cell_style(*style_idx as usize))
    }

    /// Get the complete cell format information for a specific cell.
    ///
    /// Returns a `CellFormat` with resolved font, fill, border, and number format.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn get_cell_format(&self, row: u32, column: u32) -> Option<CellFormat> {
        let style = self.get_cell_style(row, column)?;
        let styles = self.workbook.styles();

        let font = style
            .font_id
            .and_then(|id| styles.get_font(id as usize))
            .map(|f| CellFont {
                name: f.name.clone(),
                size: f.size,
                bold: f.bold,
                italic: f.italic,
                underline: f.underline.is_some(),
                color: f.color.clone(),
            });

        let fill = style
            .fill_id
            .and_then(|id| styles.get_fill(id as usize))
            .and_then(|f| match f {
                super::styles::Fill::Pattern {
                    pattern_type,
                    fg_color,
                    bg_color,
                } => {
                    // Map pattern type string to enum
                    let pattern_enum = match pattern_type.as_str() {
                        "solid" => super::format::CellFillPatternType::Solid,
                        "gray125" => super::format::CellFillPatternType::Gray125,
                        "darkGray" => super::format::CellFillPatternType::DarkGray,
                        "mediumGray" => super::format::CellFillPatternType::MediumGray,
                        "lightGray" => super::format::CellFillPatternType::LightGray,
                        "gray0625" => super::format::CellFillPatternType::Gray0625,
                        "darkHorizontal" => super::format::CellFillPatternType::DarkHorizontal,
                        "darkVertical" => super::format::CellFillPatternType::DarkVertical,
                        "darkDown" => super::format::CellFillPatternType::DarkDown,
                        "darkUp" => super::format::CellFillPatternType::DarkUp,
                        "darkGrid" => super::format::CellFillPatternType::DarkGrid,
                        "darkTrellis" => super::format::CellFillPatternType::DarkTrellis,
                        _ => super::format::CellFillPatternType::None,
                    };
                    Some(CellFill {
                        pattern_type: pattern_enum,
                        fg_color: fg_color.clone(),
                        bg_color: bg_color.clone(),
                    })
                },
                _ => None,
            });

        let border = style
            .border_id
            .and_then(|id| styles.get_border(id as usize))
            .map(|b| CellBorder {
                left: b.left.as_ref().map(|s| super::format::CellBorderSide {
                    style: super::format::CellBorderLineStyle::Thin, // TODO: Map properly
                    color: s.color.clone(),
                }),
                right: b.right.as_ref().map(|s| super::format::CellBorderSide {
                    style: super::format::CellBorderLineStyle::Thin,
                    color: s.color.clone(),
                }),
                top: b.top.as_ref().map(|s| super::format::CellBorderSide {
                    style: super::format::CellBorderLineStyle::Thin,
                    color: s.color.clone(),
                }),
                bottom: b.bottom.as_ref().map(|s| super::format::CellBorderSide {
                    style: super::format::CellBorderLineStyle::Thin,
                    color: s.color.clone(),
                }),
                diagonal: b.diagonal.as_ref().map(|s| super::format::CellBorderSide {
                    style: super::format::CellBorderLineStyle::Thin,
                    color: s.color.clone(),
                }),
            });

        let number_format = style
            .num_fmt_id
            .and_then(|id| styles.get_number_format(id))
            .map(|nf| nf.code.clone());

        Some(CellFormat {
            font,
            fill,
            border,
            number_format,
        })
    }

    /// Check if a cell is formatted as a date.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn is_date_formatted(&self, row: u32, column: u32) -> bool {
        if let Some(style) = self.get_cell_style(row, column)
            && let Some(num_fmt_id) = style.num_fmt_id
            && let Some(num_fmt) = self.workbook.styles().get_number_format(num_fmt_id)
        {
            return num_fmt.is_date_format();
        }
        false
    }

    /// Get the date/time value from a cell formatted as a date.
    ///
    /// Returns None if the cell is not a date or doesn't contain a numeric value.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn get_date_cell_value(&self, row: u32, column: u32) -> Option<f64> {
        if !self.is_date_formatted(row, column) {
            return None;
        }

        match self.get_cell_value(row, column) {
            CellValue::Float(f) => Some(f),
            CellValue::Int(i) => Some(i as f64),
            _ => None,
        }
    }

    // ===== Merged Regions =====

    /// Get all merged cell regions in the worksheet.
    ///
    /// Returns a slice of tuples (start_row, start_col, end_row, end_col).
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// for (start_row, start_col, end_row, end_col) in ws.get_merged_regions() {
    ///     println!("Merged region: ({}, {}) to ({}, {})",
    ///              start_row, start_col, end_row, end_col);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_merged_regions(&self) -> &[(u32, u32, u32, u32)] {
        &self.merged_regions
    }

    /// Check if a cell is part of a merged region.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn is_merged_cell(&self, row: u32, column: u32) -> bool {
        self.get_merge_region(row, column).is_some()
    }

    /// Get the merged region that contains a specific cell.
    ///
    /// Returns None if the cell is not part of any merged region.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    pub fn get_merge_region(&self, row: u32, column: u32) -> Option<(u32, u32, u32, u32)> {
        self.merged_regions
            .iter()
            .find(|&&(sr, sc, er, ec)| row >= sr && row <= er && column >= sc && column <= ec)
            .copied()
    }

    // ===== Hyperlinks =====

    /// Get the hyperlink for a specific cell.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// if let Some(hyperlink) = ws.get_hyperlink(1, 1) {
    ///     println!("Cell A1 links to: {}", hyperlink.target);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_hyperlink(&self, row: u32, column: u32) -> Option<&Hyperlink> {
        let cell_ref = format!("{}{}", Cell::column_to_letters(column), row);
        self.hyperlinks.get(&cell_ref)
    }

    /// Get all hyperlinks in the worksheet.
    pub fn get_hyperlinks(&self) -> &HashMap<String, Hyperlink> {
        &self.hyperlinks
    }

    // ===== Comments =====

    /// Get the comment for a specific cell.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    /// * `column` - Column number (1-based)
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// if let Some(comment) = ws.get_cell_comment(1, 1) {
    ///     println!("Comment: {}", comment.text);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_cell_comment(&self, row: u32, column: u32) -> Option<&Comment> {
        let cell_ref = format!("{}{}", Cell::column_to_letters(column), row);
        self.comments.get(&cell_ref)
    }

    /// Get all comments in the worksheet.
    pub fn get_comments(&self) -> &HashMap<String, Comment> {
        &self.comments
    }

    // ===== Column Operations =====

    /// Get the width of a specific column.
    ///
    /// Returns the width in Excel's character units, or None if using default width.
    ///
    /// # Arguments
    /// * `column` - Column number (1-based)
    pub fn get_column_width(&self, column: u32) -> Option<f64> {
        self.columns.get(&column).and_then(|info| info.width)
    }

    /// Check if a column is hidden.
    ///
    /// # Arguments
    /// * `column` - Column number (1-based)
    pub fn is_column_hidden(&self, column: u32) -> bool {
        self.columns.get(&column).is_some_and(|info| info.hidden)
    }

    /// Get column information.
    ///
    /// # Arguments
    /// * `column` - Column number (1-based)
    pub fn get_column_info(&self, column: u32) -> Option<&ColumnInfo> {
        self.columns.get(&column)
    }

    // ===== Row Operations =====

    /// Get the height of a specific row.
    ///
    /// Returns the height in points, or None if using default height.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    pub fn get_row_height(&self, row: u32) -> Option<f64> {
        self.rows.get(&row).and_then(|info| info.height)
    }

    /// Check if a row is hidden.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    pub fn is_row_hidden(&self, row: u32) -> bool {
        self.rows.get(&row).is_some_and(|info| info.hidden)
    }

    /// Get row information.
    ///
    /// # Arguments
    /// * `row` - Row number (1-based)
    pub fn get_row_info(&self, row: u32) -> Option<&RowInfo> {
        self.rows.get(&row)
    }

    // ===== Data Validation =====

    /// Get all data validations in the worksheet.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// for validation in ws.get_data_validations() {
    ///     println!("Validation on range: {}", validation.range);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_data_validations(&self) -> &[DataValidationRule] {
        &self.data_validations
    }

    // ===== Conditional Formatting =====

    /// Get all conditional formatting rules in the worksheet.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// for rule in ws.get_conditional_formatting() {
    ///     println!("Conditional format on range: {}", rule.range);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_conditional_formatting(&self) -> &[ConditionalFormatRule] {
        &self.conditional_formats
    }

    // ===== Page Setup =====

    /// Get the page setup information.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// let page_setup = ws.get_page_setup();
    /// if page_setup.landscape {
    ///     println!("Page is in landscape orientation");
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_page_setup(&self) -> &PageSetup {
        &self.page_setup
    }

    // ===== Auto-Filter =====

    /// Get the auto-filter information.
    ///
    /// # Examples
    ///
    /// ```rust,no_run
    /// use litchi::ooxml::xlsx::Workbook;
    ///
    /// let wb = Workbook::open("workbook.xlsx")?;
    /// let ws = wb.worksheet_by_index(0)?;
    ///
    /// if let Some(auto_filter) = ws.get_auto_filter() {
    ///     println!("Auto-filter range: {}", auto_filter.range);
    /// }
    /// # Ok::<(), Box<dyn std::error::Error>>(())
    /// ```
    pub fn get_auto_filter(&self) -> Option<&AutoFilter> {
        self.auto_filter.as_ref()
    }

    // Previously TODO: Apache POI worksheet-level features - NOW IMPLEMENTED:
    // ✅ Cell formatting (reading): get_cell_style(), get_cell_format()
    // ✅ Cell types (advanced): get_cell_type() via CellValue enum
    // ✅ Date cells: is_date_formatted(), get_date_cell_value()
    // ✅ Cell hyperlinks: get_hyperlink(), get_hyperlinks()
    // ✅ Cell comments: get_cell_comment(), get_comments()
    // ✅ Merged regions: get_merged_regions(), is_merged_cell(), get_merge_region()
    // ✅ Column operations: get_column_width(), is_column_hidden(), get_column_info()
    // ✅ Row operations: is_row_hidden(), get_row_height(), get_row_info()
    // ✅ Auto-filter: get_auto_filter()
    // ✅ Data validation: get_data_validations()
    // ✅ Conditional formatting: get_conditional_formatting()
    // ✅ Page setup: get_page_setup()
    //
    // Still TODO (writing operations and advanced features):
    // - Formula evaluation: evaluate_formula(), get_formula_evaluator()
    // - Array formulas: set_array_formula(), get_array_formulas()
    // - Rich text cells: get_rich_string_cell_value(), set_rich_text_string()
    // - Set operations: set_hyperlink(), remove_hyperlink(), set_cell_comment(), remove_cell_comment()
    // - Column/row mutations: auto_size_column(), set_column_hidden(), set_row_hidden(), set_row_height()
    // - Sheet protection: protect_sheet(), is_protected(), get_protection_info()
    // - Set operations: set_auto_filter(), add_validation_data()
    // - Set operations: set_fit_to_page(), set_header(), set_footer()
    // - Repeating rows/columns: set_repeating_rows(), set_repeating_columns()
}

impl<'a> WorksheetTrait for Worksheet<'a> {
    fn name(&self) -> &str {
        &self.info.name
    }

    fn row_count(&self) -> usize {
        self.dimensions
            .map(|(_, _, max_row, _)| max_row as usize)
            .unwrap_or(0)
    }

    fn column_count(&self) -> usize {
        self.dimensions
            .map(|(_, _, _, max_col)| max_col as usize)
            .unwrap_or(0)
    }

    fn dimensions(&self) -> Option<(u32, u32, u32, u32)> {
        self.dimensions
    }

    fn cell(&self, row: u32, column: u32) -> SheetResult<Box<dyn CellTrait + '_>> {
        let value = self.get_cell_value(row, column);
        let cell = Cell::new(row, column, value);
        Ok(Box::new(cell))
    }

    fn cell_by_coordinate(&self, coordinate: &str) -> SheetResult<Box<dyn CellTrait + '_>> {
        let (col, row) = Cell::reference_to_coords(coordinate)?;
        self.cell(row, col)
    }

    fn cells(&self) -> Box<dyn CellIterator<'_> + '_> {
        let mut cells = Vec::new();

        for (&row, row_data) in &self.cells {
            for (&col, value) in row_data {
                cells.push(Cell::new(row, col, value.clone()));
            }
        }

        Box::new(XlsxCellIterator::new(cells))
    }

    fn rows(&self) -> Box<dyn RowIterator<'_> + '_> {
        let mut rows = Vec::new();

        if let Some((min_row, min_col, max_row, max_col)) = self.dimensions {
            for row in min_row..=max_row {
                let mut row_data = Vec::new();
                for col in min_col..=max_col {
                    let value = self.get_cell_value(row, col).clone();
                    row_data.push(value);
                }
                rows.push(row_data);
            }
        }

        Box::new(XlsxRowIterator::new(rows))
    }

    fn row(&self, row_idx: usize) -> SheetResult<Cow<'_, [CellValue]>> {
        if let Some((min_row, min_col, max_col)) =
            self.dimensions.map(|(mr, mc, _, mc2)| (mr, mc, mc2))
        {
            let row_num = min_row + row_idx as u32;
            if row_num > self.dimensions.unwrap().2 {
                return Ok(Cow::Owned(Vec::new()));
            }

            let mut row_data = Vec::new();
            for col in min_col..=max_col {
                let value = self.get_cell_value(row_num, col).clone();
                row_data.push(value);
            }
            Ok(Cow::Owned(row_data))
        } else {
            Ok(Cow::Owned(Vec::new()))
        }
    }

    fn cell_value(&self, row: u32, column: u32) -> SheetResult<Cow<'_, CellValue>> {
        // XLSX values need shared string resolution, so we return owned
        Ok(Cow::Owned(self.get_cell_value(row, column)))
    }
}

/// Iterator over worksheets in a workbook
pub struct WorksheetIterator<'a> {
    worksheets: Vec<WorksheetInfo>,
    workbook: &'a Workbook,
    index: usize,
}

impl<'a> WorksheetIterator<'a> {
    /// Create a new worksheet iterator.
    pub fn new(worksheets: Vec<WorksheetInfo>, workbook: &'a Workbook) -> Self {
        Self {
            worksheets,
            workbook,
            index: 0,
        }
    }
}

impl<'a> crate::sheet::WorksheetIterator<'a> for WorksheetIterator<'a> {
    fn next(&mut self) -> Option<SheetResult<Box<dyn WorksheetTrait + 'a>>> {
        if self.index >= self.worksheets.len() {
            return None;
        }

        let info = &self.worksheets[self.index];
        let mut worksheet = Worksheet::new(self.workbook, info.clone());

        match worksheet.load_data() {
            Ok(_) => {
                self.index += 1;
                Some(Ok(Box::new(worksheet) as Box<dyn WorksheetTrait + 'a>))
            },
            Err(e) => {
                self.index += 1;
                Some(Err(e))
            },
        }
    }
}

// Import Workbook from the workbook module
use super::workbook::Workbook;
