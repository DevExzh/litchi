/// Table types and implementation for DOCX documents.
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::TableBorderStyle;
// Import paragraph types
use super::paragraph::MutableParagraph;

/// Border definition for table or cell.
#[derive(Debug, Clone)]
pub struct TableBorder {
    /// Border style
    pub style: TableBorderStyle,
    /// Border width in eighths of a point (e.g., 8 = 1pt, 24 = 3pt)
    pub size: u32,
    /// Border color in hex RGB format (e.g., "FF0000" for red)
    pub color: String,
}

impl Default for TableBorder {
    fn default() -> Self {
        Self {
            style: TableBorderStyle::Single,
            size: 4,
            color: "000000".to_string(),
        }
    }
}

/// Table borders (all sides).
#[derive(Debug, Clone, Default)]
pub struct TableBorders {
    pub top: Option<TableBorder>,
    pub left: Option<TableBorder>,
    pub bottom: Option<TableBorder>,
    pub right: Option<TableBorder>,
    pub inside_h: Option<TableBorder>,
    pub inside_v: Option<TableBorder>,
}

/// Table properties.
#[derive(Debug, Default)]
pub(crate) struct TableProperties {
    pub(crate) borders: TableBorders,
    pub(crate) width_pct: Option<u32>,
}

/// Cell properties.
#[derive(Debug, Default, Clone)]
pub struct CellProperties {
    /// Cell background color in hex RGB format
    pub background_color: Option<String>,
    /// Cell borders (if different from table borders)
    pub borders: Option<TableBorders>,
    /// Cell width in DXA units (twentieth of a point)
    pub width_dxa: Option<u32>,
}

/// A mutable table.
#[derive(Debug)]
pub struct MutableTable {
    /// Table rows
    pub(crate) rows: Vec<MutableRow>,
    /// Table properties
    pub(crate) properties: TableProperties,
}

impl MutableTable {
    pub(crate) fn new(rows: usize, cols: usize) -> Self {
        let mut table = Self {
            rows: Vec::with_capacity(rows),
            properties: TableProperties::default(),
        };
        for _ in 0..rows {
            table.add_row(cols);
        }
        table
    }

    /// Add a new row with specified column count.
    pub fn add_row(&mut self, cols: usize) -> &mut MutableRow {
        self.rows.push(MutableRow::new(cols));
        self.rows.last_mut().unwrap()
    }

    /// Set table width as percentage (100-500 where 100=20%, 500=100%).
    pub fn set_width_percent(&mut self, percent: u32) {
        self.properties.width_pct = Some(percent * 50);
    }

    /// Set all table borders at once.
    pub fn set_borders(&mut self, border: TableBorder) {
        self.properties.borders.top = Some(border.clone());
        self.properties.borders.left = Some(border.clone());
        self.properties.borders.bottom = Some(border.clone());
        self.properties.borders.right = Some(border.clone());
        self.properties.borders.inside_h = Some(border.clone());
        self.properties.borders.inside_v = Some(border);
    }

    /// Get a cell by row and column index.
    pub fn cell(&mut self, row: usize, col: usize) -> Option<&mut MutableCell> {
        self.rows.get_mut(row)?.cell(col)
    }

    /// Get the number of rows.
    pub fn row_count(&self) -> usize {
        self.rows.len()
    }

    /// Get a row by index.
    pub fn row(&mut self, index: usize) -> Option<&mut MutableRow> {
        self.rows.get_mut(index)
    }

    fn write_border(xml: &mut String, name: &str, border: &TableBorder) -> Result<()> {
        write!(
            xml,
            "<w:{} w:val=\"{}\" w:sz=\"{}\" w:space=\"0\" w:color=\"{}\"/>",
            name,
            border.style.as_str(),
            border.size,
            border.color
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))
    }

    pub(crate) fn to_xml(&self, xml: &mut String) -> Result<()> {
        xml.push_str("<w:tbl>");

        // Write table properties
        xml.push_str("<w:tblPr>");

        // Table width
        let width = self.properties.width_pct.unwrap_or(5000);
        write!(xml, "<w:tblW w:w=\"{}\" w:type=\"pct\"/>", width)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;

        // Table borders
        xml.push_str("<w:tblBorders>");
        if let Some(ref border) = self.properties.borders.top {
            Self::write_border(xml, "top", border)?;
        } else {
            xml.push_str("<w:top w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
        }
        if let Some(ref border) = self.properties.borders.left {
            Self::write_border(xml, "left", border)?;
        } else {
            xml.push_str("<w:left w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
        }
        if let Some(ref border) = self.properties.borders.bottom {
            Self::write_border(xml, "bottom", border)?;
        } else {
            xml.push_str(
                "<w:bottom w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>",
            );
        }
        if let Some(ref border) = self.properties.borders.right {
            Self::write_border(xml, "right", border)?;
        } else {
            xml.push_str("<w:right w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>");
        }
        if let Some(ref border) = self.properties.borders.inside_h {
            Self::write_border(xml, "insideH", border)?;
        } else {
            xml.push_str(
                "<w:insideH w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>",
            );
        }
        if let Some(ref border) = self.properties.borders.inside_v {
            Self::write_border(xml, "insideV", border)?;
        } else {
            xml.push_str(
                "<w:insideV w:val=\"single\" w:sz=\"4\" w:space=\"0\" w:color=\"000000\"/>",
            );
        }
        xml.push_str("</w:tblBorders>");
        xml.push_str("</w:tblPr>");

        // Write grid
        if let Some(first_row) = self.rows.first() {
            xml.push_str("<w:tblGrid>");
            for _ in 0..first_row.cell_count() {
                xml.push_str("<w:gridCol/>");
            }
            xml.push_str("</w:tblGrid>");
        }

        // Write rows
        for row in &self.rows {
            row.to_xml(xml)?;
        }

        xml.push_str("</w:tbl>");

        Ok(())
    }
}

/// A mutable table row.
#[derive(Debug)]
pub struct MutableRow {
    /// Table cells in this row
    pub(crate) cells: Vec<MutableCell>,
}

impl MutableRow {
    pub(crate) fn new(cols: usize) -> Self {
        let mut row = Self {
            cells: Vec::with_capacity(cols),
        };
        for _ in 0..cols {
            row.cells.push(MutableCell::new());
        }
        row
    }

    /// Get a cell by index.
    pub fn cell(&mut self, index: usize) -> Option<&mut MutableCell> {
        self.cells.get_mut(index)
    }

    /// Add a new cell.
    pub fn add_cell(&mut self) -> &mut MutableCell {
        self.cells.push(MutableCell::new());
        self.cells.last_mut().unwrap()
    }

    /// Get the number of cells.
    pub fn cell_count(&self) -> usize {
        self.cells.len()
    }

    pub(crate) fn to_xml(&self, xml: &mut String) -> Result<()> {
        xml.push_str("<w:tr>");

        for cell in &self.cells {
            cell.to_xml(xml)?;
        }

        xml.push_str("</w:tr>");

        Ok(())
    }
}

/// A mutable table cell.
#[derive(Debug)]
pub struct MutableCell {
    /// Paragraphs in this cell
    pub(crate) paragraphs: Vec<MutableParagraph>,
    /// Cell properties
    pub(crate) properties: CellProperties,
}

impl MutableCell {
    pub(crate) fn new() -> Self {
        Self {
            paragraphs: vec![MutableParagraph::new()],
            properties: CellProperties::default(),
        }
    }

    /// Add a new paragraph to the cell.
    pub fn add_paragraph(&mut self) -> &mut MutableParagraph {
        self.paragraphs.push(MutableParagraph::new());
        self.paragraphs.last_mut().unwrap()
    }

    /// Get the number of paragraphs.
    pub fn paragraph_count(&self) -> usize {
        self.paragraphs.len()
    }

    /// Get a paragraph by index.
    pub fn paragraph(&mut self, index: usize) -> Option<&mut MutableParagraph> {
        self.paragraphs.get_mut(index)
    }

    /// Set text in the first paragraph.
    pub fn set_text(&mut self, text: &str) {
        self.paragraphs.clear();
        let para = self.add_paragraph();
        para.add_run_with_text(text);
    }

    /// Set cell background color in hex RGB format (e.g., "FFFF00" for yellow).
    pub fn set_background_color(&mut self, color: &str) {
        self.properties.background_color = Some(color.to_string());
    }

    /// Set cell width in DXA units (twentieth of a point).
    pub fn set_width_dxa(&mut self, width: u32) {
        self.properties.width_dxa = Some(width);
    }

    pub(crate) fn to_xml(&self, xml: &mut String) -> Result<()> {
        xml.push_str("<w:tc>");

        // Write cell properties if any
        if self.properties.background_color.is_some() || self.properties.width_dxa.is_some() {
            xml.push_str("<w:tcPr>");

            if let Some(ref bg_color) = self.properties.background_color {
                write!(
                    xml,
                    "<w:shd w:val=\"clear\" w:color=\"auto\" w:fill=\"{}\"/>",
                    bg_color
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            if let Some(width) = self.properties.width_dxa {
                write!(xml, "<w:tcW w:w=\"{}\" w:type=\"dxa\"/>", width)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
            }

            xml.push_str("</w:tcPr>");
        }

        for para in &self.paragraphs {
            para.to_xml(xml)?;
        }

        xml.push_str("</w:tc>");

        Ok(())
    }
}
