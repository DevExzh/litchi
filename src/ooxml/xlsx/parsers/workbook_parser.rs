//! Parser for Excel workbook.xml files.
//!
//! This module provides parsing functionality for the main workbook.xml
//! file which contains sheet definitions and workbook-level metadata.
//!
//! Performance optimizations:
//! - Uses memchr for fast character searching
//! - Uses atoi_simd for fast integer parsing
//! - Pre-allocates vectors with reasonable capacities

use crate::ooxml::xlsx::worksheet::WorksheetInfo;
use crate::sheet::Result;

// Performance: Pre-allocate typical capacity for worksheets
const INITIAL_SHEETS_CAPACITY: usize = 16;

/// Parse workbook.xml content to extract sheet information and active sheet.
pub fn parse_workbook_xml(content: &str) -> Result<(Vec<WorksheetInfo>, usize)> {
    let mut sheets = Vec::with_capacity(INITIAL_SHEETS_CAPACITY);
    let mut active_sheet_id = 0;

    let bytes = content.as_bytes();

    // Look for <sheets> section - optimized search
    if let Some(sheets_start) = memchr::memmem::find(bytes, b"<sheets>")
        && let Some(sheets_end) = memchr::memmem::find(&bytes[sheets_start..], b"</sheets>")
    {
        let sheets_content = &content[sheets_start..sheets_start + sheets_end];

        // Parse individual sheet entries - optimized parsing
        parse_sheets_section(sheets_content, &mut sheets)?;
    }

    // Look for active sheet - optimized search
    if let Some(book_views_start) = memchr::memmem::find(bytes, b"<bookViews>")
        && let Some(book_views_end) =
            memchr::memmem::find(&bytes[book_views_start..], b"</bookViews>")
    {
        let book_views_content = &content[book_views_start..book_views_start + book_views_end];

        if let Some(active_tab_start) =
            memchr::memmem::find(book_views_content.as_bytes(), b"activeTab=\"")
        {
            let active_tab_content = &book_views_content[active_tab_start + 11..];
            if let Some(quote_pos) = memchr::memchr(b'"', active_tab_content.as_bytes()) {
                // Performance: Use atoi_simd for fast integer parsing
                if let Ok(tab) = atoi_simd::parse(&active_tab_content.as_bytes()[..quote_pos]) {
                    active_sheet_id = tab;
                }
            }
        }
    }

    let final_active_sheet_index = active_sheet_id.min(sheets.len().saturating_sub(1));
    Ok((sheets, final_active_sheet_index))
}

/// Parse the sheets section to extract individual sheet information.
fn parse_sheets_section(sheets_content: &str, sheets: &mut Vec<WorksheetInfo>) -> Result<()> {
    let bytes = sheets_content.as_bytes();
    let mut sheet_start = 0;

    while let Some(sheet_pos) = memchr::memmem::find(&bytes[sheet_start..], b"<sheet ") {
        let sheet_start_pos = sheet_start + sheet_pos;
        if let Some(sheet_end_pos) = memchr::memmem::find(&bytes[sheet_start_pos..], b"/>") {
            let sheet_xml = &sheets_content[sheet_start_pos..sheet_start_pos + sheet_end_pos + 2];

            if let Some(info) = parse_sheet_xml(sheet_xml)? {
                sheets.push(info);
            }
            sheet_start = sheet_start_pos + sheet_end_pos + 2;
        } else {
            break;
        }
    }

    Ok(())
}

/// Parse individual sheet XML to extract worksheet information - optimized version.
pub fn parse_sheet_xml(sheet_xml: &str) -> Result<Option<WorksheetInfo>> {
    let bytes = sheet_xml.as_bytes();

    // Extract name attribute - optimized attribute parsing
    let name = if let Some(name_start) = memchr::memmem::find(bytes, b"name=\"") {
        let name_content = &sheet_xml[name_start + 6..];
        memchr::memchr(b'"', name_content.as_bytes())
            .map(|quote_pos| name_content[..quote_pos].to_string())
    } else {
        None
    };

    // Extract relationship ID - optimized attribute parsing
    let relationship_id = if let Some(r_start) = memchr::memmem::find(bytes, b"r:id=\"") {
        let r_content = &sheet_xml[r_start + 6..];
        memchr::memchr(b'"', r_content.as_bytes())
            .map(|quote_pos| r_content[..quote_pos].to_string())
    } else {
        None
    };

    // Extract sheet ID - optimized attribute parsing with fast integer conversion
    let sheet_id = if let Some(id_start) = memchr::memmem::find(bytes, b"sheetId=\"") {
        let id_content = &sheet_xml[id_start + 9..];
        if let Some(quote_pos) = memchr::memchr(b'"', id_content.as_bytes()) {
            // Performance: Use atoi_simd for fast integer parsing
            atoi_simd::parse(&id_content.as_bytes()[..quote_pos]).ok()
        } else {
            None
        }
    } else {
        None
    };

    match (name, relationship_id, sheet_id) {
        (Some(name), Some(relationship_id), Some(sheet_id)) => {
            Ok(Some(WorksheetInfo {
                name,
                relationship_id,
                sheet_id,
                is_active: false, // Will be set later
                print_area: None,
                repeating_rows: None,
                repeating_columns: None,
            }))
        },
        _ => Ok(None),
    }
}
