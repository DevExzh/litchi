//! Parser for Excel worksheet XML files.
//!
//! This module provides parsing functionality for individual worksheet
//! XML files (sheet1.xml, sheet2.xml, etc.) to extract cell data.
//!
//! Performance optimizations:
//! - Uses memchr for fast character searching
//! - Uses atoi_simd for fast integer parsing
//! - Uses fast_float2 for fast float parsing
//! - Minimizes allocations and string operations

use std::collections::HashMap;

use crate::sheet::{CellValue, Result};

// Performance: Pre-allocate typical capacities to reduce reallocations
const INITIAL_ROW_CAPACITY: usize = 1000;
const INITIAL_COL_CAPACITY: usize = 100;

/// Parse worksheet XML content to extract cell data.
pub fn parse_worksheet_xml(content: &str) -> Result<HashMap<u32, HashMap<u32, CellValue>>> {
    let mut cells = HashMap::with_capacity(INITIAL_ROW_CAPACITY);

    // Find the sheetData section using optimized search
    if let Some(sheet_data_start) = memchr::memmem::find(content.as_bytes(), b"<sheetData>")
        && let Some(sheet_data_end) =
            memchr::memmem::find(&content.as_bytes()[sheet_data_start..], b"</sheetData>")
    {
        let sheet_data_content = &content[sheet_data_start..sheet_data_start + sheet_data_end];

        // Parse individual rows and cells
        parse_sheet_data(sheet_data_content, &mut cells)?;
    }

    Ok(cells)
}

/// Parse sheetData content to extract cells.
pub fn parse_sheet_data(
    sheet_data: &str,
    cells: &mut HashMap<u32, HashMap<u32, CellValue>>,
) -> Result<()> {
    let bytes = sheet_data.as_bytes();
    let mut pos = 0;

    while let Some(row_start) = memchr::memmem::find(&bytes[pos..], b"<row ") {
        let row_start_pos = pos + row_start;
        if let Some(row_end) = memchr::memmem::find(&bytes[row_start_pos..], b"</row>") {
            let row_content = &sheet_data[row_start_pos..row_start_pos + row_end + 6];

            if let Some((row_num, row_cells)) = parse_row_xml(row_content)? {
                // Performance: Use entry API to avoid double hash lookups
                let row_map = cells
                    .entry(row_num)
                    .or_insert_with(|| HashMap::with_capacity(INITIAL_COL_CAPACITY));
                for (col_num, value) in row_cells {
                    row_map.insert(col_num, value);
                }
            }

            pos = row_start_pos + row_end + 6;
        } else {
            break;
        }
    }

    Ok(())
}

/// Parse a single row XML to extract cells.
#[allow(clippy::type_complexity)]
pub fn parse_row_xml(row_content: &str) -> Result<Option<(u32, Vec<(u32, CellValue)>)>> {
    // Extract row number - optimized attribute parsing
    let row_num = if let Some(r_start) = memchr::memmem::find(row_content.as_bytes(), b"r=\"") {
        let r_content = &row_content[r_start + 3..];
        if let Some(quote_pos) = memchr::memchr(b'"', r_content.as_bytes()) {
            // Performance: Use atoi_simd for fast integer parsing
            atoi_simd::parse(&r_content.as_bytes()[..quote_pos]).ok()
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
    let bytes = row_content.as_bytes();
    let mut pos = 0;

    // Parse cells in this row
    while let Some(c_start) = memchr::memmem::find(&bytes[pos..], b"<c ") {
        let c_start_pos = pos + c_start;
        if let Some(c_end) = memchr::memmem::find(&bytes[c_start_pos..], b"</c>") {
            let c_content = &row_content[c_start_pos..c_start_pos + c_end + 4];

            if let Some((col_num, value)) = parse_cell_xml(c_content)? {
                cells.push((col_num, value));
            }

            pos = c_start_pos + c_end + 4;
        } else {
            break;
        }
    }

    Ok(Some((row_num, cells)))
}

/// Parse a single cell XML to extract value and coordinates.
pub fn parse_cell_xml(cell_content: &str) -> Result<Option<(u32, CellValue)>> {
    // Extract cell reference (e.g., "A1") - optimized attribute parsing
    let reference = if let Some(r_start) = memchr::memmem::find(cell_content.as_bytes(), b"r=\"") {
        let r_content = &cell_content[r_start + 3..];
        memchr::memchr(b'"', r_content.as_bytes()).map(|quote_pos| &r_content[..quote_pos])
    } else {
        None
    };

    let reference = match reference {
        Some(r) => r,
        None => return Ok(None),
    };

    // Convert reference to coordinates - optimized version
    let (col_num, _row_num) = reference_to_coords(reference)?;

    // Extract cell type - optimized attribute parsing
    let cell_type = if let Some(t_start) = memchr::memmem::find(cell_content.as_bytes(), b"t=\"") {
        let t_content = &cell_content[t_start + 3..];
        memchr::memchr(b'"', t_content.as_bytes()).map(|quote_pos| &t_content[..quote_pos])
    } else {
        None
    };

    // Extract value - optimized tag search
    let value = if let Some(v_start) = memchr::memmem::find(cell_content.as_bytes(), b"<v>") {
        let v_start_pos = v_start + 3;
        memchr::memmem::find(&cell_content.as_bytes()[v_start_pos..], b"</v>")
            .map(|v_end| &cell_content[v_start_pos..v_start_pos + v_end])
    } else {
        None
    };

    let cell_value = match (cell_type, value) {
        (Some("str"), Some(v)) => CellValue::String(v.to_string()),
        (Some("s"), Some(v)) => {
            // Shared string reference - this will be resolved later
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

/// Convert Excel reference (e.g., "A1") to coordinates - optimized version.
pub fn reference_to_coords(reference: &str) -> Result<(u32, u32)> {
    let bytes = reference.as_bytes();
    let mut col_str_end = 0;

    // Find where column letters end and row numbers begin
    for (i, &byte) in bytes.iter().enumerate() {
        if byte.is_ascii_digit() {
            col_str_end = i;
            break;
        }
    }

    if col_str_end == 0 {
        return Err(format!("Invalid reference: {}", reference).into());
    }

    // Convert column letters to number (A=1, B=2, ..., Z=26, AA=27, etc.)
    let mut col_num = 0u32;
    for &byte in &bytes[..col_str_end] {
        if !byte.is_ascii_alphabetic() {
            return Err(format!("Invalid column in reference: {}", reference).into());
        }
        col_num = col_num * 26 + (byte.to_ascii_uppercase() - b'A' + 1) as u32;
    }

    // Parse row number using fast integer parsing
    let row_part = &bytes[col_str_end..];
    let row_num = atoi_simd::parse(row_part)
        .map_err(|_| format!("Invalid row number in reference: {}", reference))?;

    Ok((col_num, row_num))
}
