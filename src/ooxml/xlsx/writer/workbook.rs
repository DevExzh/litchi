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

        xml.push_str("<sheets>");
        for (index, ws) in self.worksheets.iter().enumerate() {
            let sheet_id = ws.sheet_id();
            let rel_id = worksheet_rel_ids
                .get(index)
                .map(|s| s.as_str())
                .unwrap_or("rId1"); // Fallback, shouldn't happen
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

        xml.push_str("</workbook>");

        Ok(xml)
    }
}

impl Default for MutableWorkbookData {
    fn default() -> Self {
        Self::new()
    }
}
