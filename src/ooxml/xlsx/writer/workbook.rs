/// Workbook data structure for XLSX.
use crate::sheet::Result as SheetResult;
use std::collections::HashMap;
use std::fmt::Write as FmtWrite;

use super::sheet::{MutableWorksheet, NamedRange};
use super::strings::MutableSharedStrings;
use super::styles::StylesBuilder;

/// Type alias for cell position to style index mapping.
type CellStyleMap = HashMap<(u32, u32), usize>;

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Escape worksheet names for use in definedName references.
///
/// This doubles single quotes, e.g. "Bob's Sheet" -> "Bob''s Sheet".
fn escape_sheet_name(name: &str) -> String {
    let mut escaped = String::with_capacity(name.len());
    for ch in name.chars() {
        if ch == '\'' {
            escaped.push_str("''");
        } else {
            escaped.push(ch);
        }
    }
    escaped
}

/// Workbook protection configuration.
#[derive(Debug, Clone)]
pub struct WorkbookProtection {
    /// Password hash (optional)
    pub password_hash: Option<String>,
    /// Lock structure (prevent adding/deleting sheets)
    pub lock_structure: bool,
    /// Lock windows (prevent resizing/moving workbook window)
    pub lock_windows: bool,
}

/// Mutable workbook for writing.
///
/// This is managed internally by the Workbook struct.
#[derive(Debug)]
pub struct MutableWorkbookData {
    /// Worksheets
    pub worksheets: Vec<MutableWorksheet>,
    /// Shared strings table
    pub shared_strings: MutableSharedStrings,
    /// Named ranges
    pub named_ranges: Vec<NamedRange>,
    /// Workbook protection
    pub protection: Option<WorkbookProtection>,
    /// Force formula recalculation on open
    pub force_formula_recalculation: bool,
    /// Calculation mode: "auto", "manual", or "autoNoTable"
    pub calculation_mode: String,
    /// Whether the workbook has been modified
    pub modified: bool,
}

impl MutableWorkbookData {
    /// Create a new workbook with one default worksheet.
    pub fn new() -> Self {
        let mut data = Self {
            worksheets: Vec::new(),
            shared_strings: MutableSharedStrings::new(),
            named_ranges: Vec::new(),
            protection: None,
            force_formula_recalculation: false,
            calculation_mode: "auto".to_string(),
            modified: false,
        };

        // Add a default worksheet
        data.add_worksheet("Sheet1".to_string());

        data
    }

    /// Add a new worksheet.
    pub fn add_worksheet(&mut self, name: String) -> &mut MutableWorksheet {
        let sheet_id = (self.worksheets.len() + 1) as u32;
        let worksheet = MutableWorksheet::new(name, sheet_id);
        self.worksheets.push(worksheet);
        self.modified = true;
        self.worksheets.last_mut().unwrap()
    }

    /// Get a worksheet by index.
    pub fn worksheet_mut(&mut self, index: usize) -> SheetResult<&mut MutableWorksheet> {
        self.worksheets
            .get_mut(index)
            .ok_or_else(|| "Worksheet index out of bounds".into())
    }

    /// Get the number of worksheets.
    pub fn worksheet_count(&self) -> usize {
        self.worksheets.len()
    }

    /// Define a named range.
    ///
    /// Named ranges allow you to refer to cells or ranges by meaningful names.
    pub fn define_name(&mut self, name: &str, reference: &str) {
        self.named_ranges.push(NamedRange {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet_id: None,
        });
        self.modified = true;
    }

    /// Define a sheet-scoped named range.
    ///
    /// Sheet-scoped names are only visible within the specified worksheet.
    pub fn define_name_local(&mut self, name: &str, reference: &str, sheet_id: u32) {
        self.named_ranges.push(NamedRange {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: None,
            local_sheet_id: Some(sheet_id),
        });
        self.modified = true;
    }

    /// Define a named range with a comment.
    pub fn define_name_with_comment(&mut self, name: &str, reference: &str, comment: &str) {
        self.named_ranges.push(NamedRange {
            name: name.to_string(),
            reference: reference.to_string(),
            comment: Some(comment.to_string()),
            local_sheet_id: None,
        });
        self.modified = true;
    }

    /// Remove a named range by name.
    pub fn remove_name(&mut self, name: &str) -> bool {
        let initial_len = self.named_ranges.len();
        self.named_ranges.retain(|r| r.name != name);
        let removed = self.named_ranges.len() < initial_len;
        if removed {
            self.modified = true;
        }
        removed
    }

    /// Get all named ranges.
    pub fn named_ranges(&self) -> &[NamedRange] {
        &self.named_ranges
    }

    /// Synchronize worksheet print settings with internal defined names.
    ///
    /// This maps worksheet-level page setup (print area, repeating rows/columns)
    /// to the Excel-reserved sheet-scoped defined names `_xlnm.Print_Area` and
    /// `_xlnm.Print_Titles` in `workbook.xml`.
    pub(crate) fn sync_print_settings_to_defined_names(&mut self) {
        // Preserve all existing named ranges except the internal print-related
        // ones for each sheet, which we rebuild from the current worksheet
        // settings.
        let mut new_ranges =
            Vec::with_capacity(self.named_ranges.len() + self.worksheets.len() * 2);

        for range in &self.named_ranges {
            let is_internal_print_name = (range.name == "_xlnm.Print_Area"
                || range.name == "_xlnm.Print_Titles")
                && range.local_sheet_id.is_some();

            if !is_internal_print_name {
                new_ranges.push(range.clone());
            }
        }

        // Rebuild per-sheet print area and print titles based on the
        // MutableWorksheet settings.
        for ws in &self.worksheets {
            let sheet_id = ws.sheet_id();
            let sheet_name = ws.name();
            let escaped_sheet_name = escape_sheet_name(sheet_name);

            // Print area -> _xlnm.Print_Area
            if let Some(range) = ws.get_print_area() {
                let reference = format!("'{}'!{}", escaped_sheet_name, range);
                new_ranges.push(NamedRange {
                    name: "_xlnm.Print_Area".to_string(),
                    reference,
                    comment: None,
                    local_sheet_id: Some(sheet_id),
                });
            }

            // Repeating rows/columns -> _xlnm.Print_Titles
            let repeating_cols = ws.get_repeating_columns();
            let repeating_rows = ws.get_repeating_rows();

            if repeating_cols.is_some() || repeating_rows.is_some() {
                let mut parts = Vec::new();

                if let Some(cols) = repeating_cols {
                    parts.push(format!("'{}'!{}", escaped_sheet_name, cols));
                }
                if let Some(rows) = repeating_rows {
                    parts.push(format!("'{}'!{}", escaped_sheet_name, rows));
                }

                let reference = parts.join(",");
                new_ranges.push(NamedRange {
                    name: "_xlnm.Print_Titles".to_string(),
                    reference,
                    comment: None,
                    local_sheet_id: Some(sheet_id),
                });
            }
        }

        self.named_ranges = new_ranges;
        self.modified = true;
    }

    /// Check if the workbook has been modified.
    pub fn is_modified(&self) -> bool {
        self.modified || self.worksheets.iter().any(|ws| ws.is_modified())
    }

    /// Build styles from all worksheets and return a StylesBuilder and cell position -> style index mappings.
    ///
    /// Returns a tuple of (StylesBuilder, Vec of per-worksheet CellStyleMap).
    pub fn build_styles(&self) -> SheetResult<(StylesBuilder, Vec<CellStyleMap>)> {
        let mut builder = StylesBuilder::new();
        let mut worksheet_style_indices = Vec::new();

        // For each worksheet, collect cell formats and build style indices
        for ws in &self.worksheets {
            let mut style_map = CellStyleMap::new();

            // Iterate through all cells with formats
            for (pos, format) in ws.cell_formats() {
                // Add format to builder and get its style index
                let style_index = builder.add_cell_format(format);
                style_map.insert(*pos, style_index);
            }

            worksheet_style_indices.push(style_map);
        }

        Ok((builder, worksheet_style_indices))
    }

    /// Generate workbook.xml content.
    pub fn generate_workbook_xml(&self) -> SheetResult<String> {
        // Default to calculating relationship IDs (legacy behavior)
        let rel_ids: Vec<String> = (1..=self.worksheets.len())
            .map(|i| format!("rId{}", i))
            .collect();
        self.generate_workbook_xml_with_rels(&rel_ids)
    }

    /// Generate workbook.xml content with actual relationship IDs.
    ///
    /// # Arguments
    /// * `worksheet_rel_ids` - Vector of relationship IDs for worksheets (e.g., ["rId1", "rId2", ...])
    pub(crate) fn generate_workbook_xml_with_rels(
        &self,
        worksheet_rel_ids: &[String],
    ) -> SheetResult<String> {
        let mut xml = String::with_capacity(2048);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);

        xml.push_str(
            r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
        );
        xml.push_str(
            r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">"#,
        );

        // Add fileVersion (recommended by Excel for compatibility)
        xml.push_str(
            r#"<fileVersion appName="xl" lastEdited="7" lowestEdited="7" rupBuild="16925"/>"#,
        );

        // Add workbookPr (required by Excel)
        xml.push_str(r#"<workbookPr defaultThemeVersion="166925"/>"#);

        // Write workbook protection if configured (must come after workbookPr per OOXML spec)
        if let Some(ref protection) = self.protection {
            xml.push_str("<workbookProtection");
            if let Some(ref hash) = protection.password_hash {
                write!(xml, r#" workbookPassword="{}""#, hash)
                    .map_err(|e| format!("XML write error: {}", e))?;
            }
            if protection.lock_structure {
                xml.push_str(r#" lockStructure="1""#);
            }
            if protection.lock_windows {
                xml.push_str(r#" lockWindows="1""#);
            }
            xml.push_str("/>");
        }

        // Add bookViews (required by Excel)
        xml.push_str("<bookViews>");
        xml.push_str(
            r#"<workbookView xWindow="0" yWindow="0" windowWidth="20000" windowHeight="10000"/>"#,
        );
        xml.push_str("</bookViews>");

        xml.push_str("<sheets>");
        for (index, ws) in self.worksheets.iter().enumerate() {
            let sheet_id = ws.sheet_id();
            let rel_id = worksheet_rel_ids
                .get(index)
                .map(|s| s.as_str())
                .unwrap_or("rId1"); // Fallback, shouldn't happen

            // Sheet elements are always self-closing (tab colors are in worksheet XML, not here)
            write!(
                xml,
                r#"<sheet name="{}" sheetId="{}" r:id="{}"/>"#,
                escape_xml(ws.name()),
                sheet_id,
                rel_id
            )
            .map_err(|e| format!("XML write error: {}", e))?;
        }
        xml.push_str("</sheets>");

        // Write defined names (named ranges)
        if !self.named_ranges.is_empty() {
            xml.push_str("<definedNames>");
            for named_range in &self.named_ranges {
                xml.push_str("<definedName name=\"");
                xml.push_str(&escape_xml(&named_range.name));
                xml.push('"');

                // Add localSheetId if it's a sheet-scoped name
                if let Some(sheet_id) = named_range.local_sheet_id {
                    write!(xml, " localSheetId=\"{}\"", sheet_id - 1)
                        .map_err(|e| format!("XML write error: {}", e))?;
                }

                // Add comment if present
                if let Some(ref comment) = named_range.comment {
                    write!(xml, " comment=\"{}\"", escape_xml(comment))
                        .map_err(|e| format!("XML write error: {}", e))?;
                }

                xml.push('>');
                xml.push_str(&escape_xml(&named_range.reference));
                xml.push_str("</definedName>");
            }
            xml.push_str("</definedNames>");
        }

        // Add calculation properties (recommended for Excel compatibility)
        xml.push_str(r#"<calcPr calcId="171027"/>"#);

        xml.push_str("</workbook>");

        Ok(xml)
    }

    // ===== Workbook-level Features =====

    /// Hide a worksheet by index.
    pub fn hide_sheet(&mut self, index: usize) -> SheetResult<()> {
        if let Some(ws) = self.worksheets.get_mut(index) {
            ws.set_hidden(true);
            self.modified = true;
            Ok(())
        } else {
            Err("Worksheet index out of bounds".into())
        }
    }

    /// Unhide a worksheet by index.
    pub fn unhide_sheet(&mut self, index: usize) -> SheetResult<()> {
        if let Some(ws) = self.worksheets.get_mut(index) {
            ws.set_hidden(false);
            self.modified = true;
            Ok(())
        } else {
            Err("Worksheet index out of bounds".into())
        }
    }

    /// Check if a worksheet is hidden.
    pub fn is_sheet_hidden(&self, index: usize) -> Option<bool> {
        self.worksheets.get(index).map(|ws| ws.is_hidden())
    }

    /// Move a worksheet to a new position.
    pub fn move_sheet(&mut self, from_index: usize, to_index: usize) -> SheetResult<()> {
        if from_index >= self.worksheets.len() || to_index >= self.worksheets.len() {
            return Err("Worksheet index out of bounds".into());
        }

        let worksheet = self.worksheets.remove(from_index);
        self.worksheets.insert(to_index, worksheet);
        self.modified = true;
        Ok(())
    }

    /// Set sheet visibility state.
    pub fn set_sheet_visibility(&mut self, index: usize, visibility: &str) -> SheetResult<()> {
        if let Some(ws) = self.worksheets.get_mut(index) {
            ws.set_visibility(visibility);
            self.modified = true;
            Ok(())
        } else {
            Err("Worksheet index out of bounds".into())
        }
    }

    /// Get sheet visibility state.
    pub fn get_sheet_visibility(&self, index: usize) -> Option<&str> {
        self.worksheets.get(index).map(|ws| ws.visibility())
    }

    /// Set the active worksheet index.
    pub fn set_active_sheet(&mut self, index: usize) {
        // Mark all worksheets as not active
        for ws in &mut self.worksheets {
            ws.set_active(false);
        }

        // Set the specified worksheet as active
        if let Some(ws) = self.worksheets.get_mut(index) {
            ws.set_active(true);
            self.modified = true;
        }
    }

    /// Force formula recalculation when the workbook is opened.
    pub fn set_force_formula_recalculation(&mut self, force: bool) {
        self.force_formula_recalculation = force;
        self.modified = true;
    }

    /// Get whether formula recalculation is forced.
    pub fn get_force_formula_recalculation(&self) -> bool {
        self.force_formula_recalculation
    }

    /// Set the calculation mode for the workbook.
    pub fn set_calculation_mode(&mut self, mode: &str) {
        self.calculation_mode = mode.to_string();
        self.modified = true;
    }

    /// Get the calculation mode for the workbook.
    pub fn get_calculation_mode(&self) -> Option<&str> {
        Some(&self.calculation_mode)
    }

    /// Protect the workbook with optional password.
    ///
    /// # Arguments
    /// * `password` - Optional password (will be hashed)
    /// * `lock_structure` - Prevent adding/deleting sheets
    /// * `lock_windows` - Prevent resizing/moving workbook window
    pub fn protect_workbook(
        &mut self,
        password: Option<&str>,
        lock_structure: bool,
        lock_windows: bool,
    ) {
        use super::sheet::MutableWorksheet;

        let password_hash = password.map(MutableWorksheet::hash_password);

        self.protection = Some(WorkbookProtection {
            password_hash,
            lock_structure,
            lock_windows,
        });
        self.modified = true;
    }

    /// Unprotect the workbook.
    pub fn unprotect_workbook(&mut self) {
        self.protection = None;
        self.modified = true;
    }

    /// Check if the workbook is protected.
    pub fn is_protected(&self) -> bool {
        self.protection.is_some()
    }

    /// Get the workbook protection configuration.
    pub fn get_protection(&self) -> Option<&WorkbookProtection> {
        self.protection.as_ref()
    }
}

impl Default for MutableWorkbookData {
    fn default() -> Self {
        Self::new()
    }
}
