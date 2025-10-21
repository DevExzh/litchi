//! Worksheet implementation for Excel files.
//!
//! This module provides the concrete implementation of worksheets
//! for Excel (.xlsx) files.

use std::collections::HashMap;

use crate::ooxml::opc::PackURI;
use crate::sheet::{
    Cell as CellTrait, CellIterator, CellValue, Result as SheetResult, RowIterator,
    Worksheet as WorksheetTrait,
};

use super::cell::{Cell, CellIterator as XlsxCellIterator, RowIterator as XlsxRowIterator};

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

/// Concrete implementation of the Worksheet trait for Excel files.
pub struct Worksheet<'a> {
    /// Reference to the parent workbook
    workbook: &'a Workbook,
    /// Worksheet information
    info: WorksheetInfo,
    /// Cached cell data (row -> column -> value)
    cells: HashMap<u32, HashMap<u32, CellValue>>,
    /// Dimensions of the worksheet (min_row, min_col, max_row, max_col)
    dimensions: Option<(u32, u32, u32, u32)>,
}

impl<'a> Worksheet<'a> {
    /// Create a new worksheet.
    pub fn new(workbook: &'a Workbook, info: WorksheetInfo) -> Self {
        Self {
            workbook,
            info,
            cells: HashMap::new(),
            dimensions: None,
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
        // Find the sheetData section
        if let Some(sheet_data_start) = content.find("<sheetData>")
            && let Some(sheet_data_end) = content[sheet_data_start..].find("</sheetData>")
        {
            let sheet_data_content = &content[sheet_data_start..sheet_data_start + sheet_data_end];

            // Parse individual rows and cells
            self.parse_sheet_data(sheet_data_content)?;
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

                if let Some((row_num, cells)) = self.parse_row_xml(row_content)? {
                    min_row = min_row.min(row_num);
                    max_row = max_row.max(row_num);

                    for (col_num, value) in cells {
                        min_col = min_col.min(col_num);
                        max_col = max_col.max(col_num);

                        self.cells
                            .entry(row_num)
                            .or_default()
                            .insert(col_num, value);
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
    ) -> SheetResult<Option<(u32, Vec<(u32, CellValue)>)>> {
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

        let mut cells = Vec::new();

        // Parse cells in this row
        let mut pos = 0;
        while let Some(c_start) = row_content[pos..].find("<c ") {
            let c_start_pos = pos + c_start;
            if let Some(c_end) = row_content[c_start_pos..].find("</c>") {
                let c_content = &row_content[c_start_pos..c_start_pos + c_end + 4];

                if let Some((col_num, value)) = self.parse_cell_xml(c_content)? {
                    cells.push((col_num, value));
                }

                pos = c_start_pos + c_end + 4;
            } else {
                break;
            }
        }

        Ok(Some((row_num, cells)))
    }

    /// Parse a single cell XML.
    fn parse_cell_xml(&self, cell_content: &str) -> SheetResult<Option<(u32, CellValue)>> {
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

        Ok(Some((col_num, cell_value)))
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

    fn row(&self, row_idx: usize) -> SheetResult<Vec<CellValue>> {
        if let Some((min_row, min_col, max_col)) =
            self.dimensions.map(|(mr, mc, _, mc2)| (mr, mc, mc2))
        {
            let row_num = min_row + row_idx as u32;
            if row_num > self.dimensions.unwrap().2 {
                return Ok(Vec::new());
            }

            let mut row_data = Vec::new();
            for col in min_col..=max_col {
                let value = self.get_cell_value(row_num, col).clone();
                row_data.push(value);
            }
            Ok(row_data)
        } else {
            Ok(Vec::new())
        }
    }

    fn cell_value(&self, row: u32, column: u32) -> SheetResult<CellValue> {
        Ok(self.get_cell_value(row, column).clone())
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
