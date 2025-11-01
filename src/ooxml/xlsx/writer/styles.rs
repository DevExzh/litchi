//! Styles.xml generator for XLSX files.
//!
//! This module handles the generation of the styles.xml file, which defines
//! all the formatting information (fonts, fills, borders, number formats, and
//! cell formats) used in an Excel workbook.

use crate::ooxml::xlsx::format::{
    CellBorder, CellBorderSide, CellFill, CellFillPatternType, CellFont, CellFormat,
};
use crate::sheet::Result as SheetResult;
use std::collections::HashMap;
use std::fmt::Write as FmtWrite;

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
}

/// Builder for generating styles.xml content.
///
/// This struct collects all unique fonts, fills, borders, and cell formats,
/// assigns indices to them, and generates the complete styles.xml content.
#[derive(Debug)]
pub struct StylesBuilder {
    /// Unique fonts (index -> font)
    fonts: Vec<CellFont>,
    /// Font lookup (font hash -> index)
    font_map: HashMap<u64, usize>,
    /// Unique fills (index -> fill)
    fills: Vec<CellFill>,
    /// Fill lookup (fill hash -> index)
    fill_map: HashMap<u64, usize>,
    /// Unique borders (index -> border)
    borders: Vec<CellBorder>,
    /// Border lookup (border hash -> index)
    border_map: HashMap<u64, usize>,
    /// Unique number formats (index -> format string)
    number_formats: Vec<String>,
    /// Number format lookup (format string -> index)
    number_format_map: HashMap<String, usize>,
    /// Cell formats (XF records) - index -> (font_id, fill_id, border_id, num_fmt_id)
    cell_formats: Vec<(usize, usize, usize, usize)>,
    /// Cell format lookup (format hash -> index)
    cell_format_map: HashMap<u64, usize>,
}

impl StylesBuilder {
    /// Create a new StylesBuilder with default styles.
    pub fn new() -> Self {
        let mut builder = Self {
            fonts: Vec::new(),
            font_map: HashMap::new(),
            fills: Vec::new(),
            fill_map: HashMap::new(),
            borders: Vec::new(),
            border_map: HashMap::new(),
            number_formats: Vec::new(),
            number_format_map: HashMap::new(),
            cell_formats: Vec::new(),
            cell_format_map: HashMap::new(),
        };

        // Add default font (required by Excel)
        builder.fonts.push(CellFont::default());
        builder
            .font_map
            .insert(Self::hash_font(&CellFont::default()), 0);

        // Add default fills (required by Excel - must be first two)
        // Fill 0: no fill
        builder.fills.push(CellFill {
            pattern_type: CellFillPatternType::None,
            fg_color: None,
            bg_color: None,
        });
        builder.fill_map.insert(
            Self::hash_fill(&CellFill {
                pattern_type: CellFillPatternType::None,
                fg_color: None,
                bg_color: None,
            }),
            0,
        );

        // Fill 1: gray125 (Excel default)
        builder.fills.push(CellFill {
            pattern_type: CellFillPatternType::Gray125,
            fg_color: None,
            bg_color: None,
        });
        builder.fill_map.insert(
            Self::hash_fill(&CellFill {
                pattern_type: CellFillPatternType::Gray125,
                fg_color: None,
                bg_color: None,
            }),
            1,
        );

        // Add default border (required by Excel)
        builder.borders.push(CellBorder::default());
        builder
            .border_map
            .insert(Self::hash_border(&CellBorder::default()), 0);

        // Add default cell format (style index 0)
        builder.cell_formats.push((0, 0, 0, 0)); // font=0, fill=0, border=0, numFmt=0

        builder
    }

    /// Add a cell format and return its style index.
    ///
    /// If the format has already been added, returns the existing index.
    pub fn add_cell_format(&mut self, format: &CellFormat) -> usize {
        let format_hash = Self::hash_cell_format(format);

        // Check if this format already exists
        if let Some(&index) = self.cell_format_map.get(&format_hash) {
            return index;
        }

        // Add font if present
        let font_id = if let Some(ref font) = format.font {
            self.add_font(font)
        } else {
            0 // Default font
        };

        // Add fill if present
        let fill_id = if let Some(ref fill) = format.fill {
            self.add_fill(fill)
        } else {
            0 // Default fill
        };

        // Add border if present
        let border_id = if let Some(ref border) = format.border {
            self.add_border(border)
        } else {
            0 // Default border
        };

        // Add number format if present
        let num_fmt_id = if let Some(ref num_fmt) = format.number_format {
            self.add_number_format(num_fmt)
        } else {
            0 // General format
        };

        // Add the cell format
        let index = self.cell_formats.len();
        self.cell_formats
            .push((font_id, fill_id, border_id, num_fmt_id));
        self.cell_format_map.insert(format_hash, index);

        index
    }

    /// Add a font and return its index.
    fn add_font(&mut self, font: &CellFont) -> usize {
        let hash = Self::hash_font(font);
        if let Some(&index) = self.font_map.get(&hash) {
            return index;
        }

        let index = self.fonts.len();
        self.fonts.push(font.clone());
        self.font_map.insert(hash, index);
        index
    }

    /// Add a fill and return its index.
    fn add_fill(&mut self, fill: &CellFill) -> usize {
        let hash = Self::hash_fill(fill);
        if let Some(&index) = self.fill_map.get(&hash) {
            return index;
        }

        let index = self.fills.len();
        self.fills.push(fill.clone());
        self.fill_map.insert(hash, index);
        index
    }

    /// Add a border and return its index.
    fn add_border(&mut self, border: &CellBorder) -> usize {
        let hash = Self::hash_border(border);
        if let Some(&index) = self.border_map.get(&hash) {
            return index;
        }

        let index = self.borders.len();
        self.borders.push(border.clone());
        self.border_map.insert(hash, index);
        index
    }

    /// Add a number format and return its index.
    fn add_number_format(&mut self, format: &str) -> usize {
        if let Some(&index) = self.number_format_map.get(format) {
            return index;
        }

        // Custom number formats start at index 164 (per Excel spec)
        let index = 164 + self.number_formats.len();
        self.number_formats.push(format.to_string());
        self.number_format_map.insert(format.to_string(), index);
        index
    }

    /// Generate the complete styles.xml content.
    pub fn to_xml(&self) -> SheetResult<String> {
        let mut xml = String::with_capacity(4096);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        xml.push_str(
            r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">"#,
        );

        // Write number formats (if any custom ones exist)
        if !self.number_formats.is_empty() {
            write!(xml, r#"<numFmts count="{}">"#, self.number_formats.len())
                .map_err(|e| format!("XML write error: {}", e))?;

            for (i, num_fmt) in self.number_formats.iter().enumerate() {
                let fmt_id = 164 + i; // Custom formats start at 164
                write!(
                    xml,
                    r#"<numFmt numFmtId="{}" formatCode="{}"/>"#,
                    fmt_id,
                    escape_xml(num_fmt)
                )
                .map_err(|e| format!("XML write error: {}", e))?;
            }

            xml.push_str("</numFmts>");
        }

        // Write fonts
        write!(xml, r#"<fonts count="{}">"#, self.fonts.len())
            .map_err(|e| format!("XML write error: {}", e))?;

        for font in &self.fonts {
            self.write_font(&mut xml, font)?;
        }

        xml.push_str("</fonts>");

        // Write fills
        write!(xml, r#"<fills count="{}">"#, self.fills.len())
            .map_err(|e| format!("XML write error: {}", e))?;

        for fill in &self.fills {
            self.write_fill(&mut xml, fill)?;
        }

        xml.push_str("</fills>");

        // Write borders
        write!(xml, r#"<borders count="{}">"#, self.borders.len())
            .map_err(|e| format!("XML write error: {}", e))?;

        for border in &self.borders {
            self.write_border(&mut xml, border)?;
        }

        xml.push_str("</borders>");

        // Write cell style XFs (required, even if empty)
        xml.push_str(r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#);

        // Write cell XFs (the actual cell formats)
        write!(xml, r#"<cellXfs count="{}">"#, self.cell_formats.len())
            .map_err(|e| format!("XML write error: {}", e))?;

        for (font_id, fill_id, border_id, num_fmt_id) in &self.cell_formats {
            write!(
                xml,
                r#"<xf numFmtId="{}" fontId="{}" fillId="{}" borderId="{}""#,
                num_fmt_id, font_id, fill_id, border_id
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            // Add applyXXX attributes if non-default
            if *font_id != 0 {
                xml.push_str(r#" applyFont="1""#);
            }
            if *fill_id != 0 {
                xml.push_str(r#" applyFill="1""#);
            }
            if *border_id != 0 {
                xml.push_str(r#" applyBorder="1""#);
            }
            if *num_fmt_id != 0 {
                xml.push_str(r#" applyNumberFormat="1""#);
            }

            xml.push_str("/>");
        }

        xml.push_str("</cellXfs>");

        // Write cell styles (required, even if minimal)
        xml.push_str(r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#);

        xml.push_str("</styleSheet>");

        Ok(xml)
    }

    /// Write a font element to XML.
    fn write_font(&self, xml: &mut String, font: &CellFont) -> SheetResult<()> {
        xml.push_str("<font>");

        if font.bold {
            xml.push_str("<b/>");
        }
        if font.italic {
            xml.push_str("<i/>");
        }
        if font.underline {
            xml.push_str("<u/>");
        }

        if let Some(ref size) = font.size {
            write!(xml, r#"<sz val="{}"/>"#, size)
                .map_err(|e| format!("XML write error: {}", e))?;
        }

        if let Some(ref color) = font.color {
            write!(xml, r#"<color rgb="{}"/>"#, escape_xml(color))
                .map_err(|e| format!("XML write error: {}", e))?;
        }

        if let Some(ref name) = font.name {
            write!(xml, r#"<name val="{}"/>"#, escape_xml(name))
                .map_err(|e| format!("XML write error: {}", e))?;
        } else {
            // Default font name
            xml.push_str(r#"<name val="Calibri"/>"#);
        }

        xml.push_str("</font>");
        Ok(())
    }

    /// Write a fill element to XML.
    fn write_fill(&self, xml: &mut String, fill: &CellFill) -> SheetResult<()> {
        xml.push_str("<fill>");

        write!(
            xml,
            r#"<patternFill patternType="{}">"#,
            fill.pattern_type.as_str()
        )
        .map_err(|e| format!("XML write error: {}", e))?;

        if let Some(ref fg_color) = fill.fg_color {
            write!(xml, r#"<fgColor rgb="{}"/>"#, escape_xml(fg_color))
                .map_err(|e| format!("XML write error: {}", e))?;
        }

        if let Some(ref bg_color) = fill.bg_color {
            write!(xml, r#"<bgColor rgb="{}"/>"#, escape_xml(bg_color))
                .map_err(|e| format!("XML write error: {}", e))?;
        }

        xml.push_str("</patternFill></fill>");
        Ok(())
    }

    /// Write a border element to XML.
    fn write_border(&self, xml: &mut String, border: &CellBorder) -> SheetResult<()> {
        xml.push_str("<border>");

        self.write_border_side(xml, "left", border.left.as_ref())?;
        self.write_border_side(xml, "right", border.right.as_ref())?;
        self.write_border_side(xml, "top", border.top.as_ref())?;
        self.write_border_side(xml, "bottom", border.bottom.as_ref())?;
        self.write_border_side(xml, "diagonal", border.diagonal.as_ref())?;

        xml.push_str("</border>");
        Ok(())
    }

    /// Write a single border side to XML.
    fn write_border_side(
        &self,
        xml: &mut String,
        side: &str,
        border_side: Option<&CellBorderSide>,
    ) -> SheetResult<()> {
        if let Some(bs) = border_side {
            write!(xml, r#"<{} style="{}">"#, side, bs.style.as_str())
                .map_err(|e| format!("XML write error: {}", e))?;

            if let Some(ref color) = bs.color {
                write!(xml, r#"<color rgb="{}"/>"#, escape_xml(color))
                    .map_err(|e| format!("XML write error: {}", e))?;
            }

            write!(xml, "</{}>", side).map_err(|e| format!("XML write error: {}", e))?;
        } else {
            write!(xml, "<{}/>", side).map_err(|e| format!("XML write error: {}", e))?;
        }

        Ok(())
    }

    /// Hash a font for deduplication.
    fn hash_font(font: &CellFont) -> u64 {
        use std::collections::hash_map::DefaultHasher;
        use std::hash::{Hash, Hasher};

        let mut hasher = DefaultHasher::new();
        font.bold.hash(&mut hasher);
        font.italic.hash(&mut hasher);
        font.underline.hash(&mut hasher);
        if let Some(ref name) = font.name {
            name.hash(&mut hasher);
        }
        if let Some(size) = font.size {
            size.to_bits().hash(&mut hasher);
        }
        if let Some(ref color) = font.color {
            color.hash(&mut hasher);
        }
        hasher.finish()
    }

    /// Hash a fill for deduplication.
    fn hash_fill(fill: &CellFill) -> u64 {
        use std::collections::hash_map::DefaultHasher;
        use std::hash::{Hash, Hasher};

        let mut hasher = DefaultHasher::new();
        // Hash the pattern type's discriminant
        std::mem::discriminant(&fill.pattern_type).hash(&mut hasher);
        if let Some(ref fg) = fill.fg_color {
            fg.hash(&mut hasher);
        }
        if let Some(ref bg) = fill.bg_color {
            bg.hash(&mut hasher);
        }
        hasher.finish()
    }

    /// Hash a border for deduplication.
    fn hash_border(border: &CellBorder) -> u64 {
        use std::collections::hash_map::DefaultHasher;
        use std::hash::Hasher;

        let mut hasher = DefaultHasher::new();
        Self::hash_border_side(&border.left, &mut hasher);
        Self::hash_border_side(&border.right, &mut hasher);
        Self::hash_border_side(&border.top, &mut hasher);
        Self::hash_border_side(&border.bottom, &mut hasher);
        Self::hash_border_side(&border.diagonal, &mut hasher);
        hasher.finish()
    }

    /// Hash a border side.
    fn hash_border_side(side: &Option<CellBorderSide>, hasher: &mut impl std::hash::Hasher) {
        use std::hash::Hash;

        if let Some(s) = side {
            std::mem::discriminant(&s.style).hash(hasher);
            if let Some(color) = &s.color {
                color.hash(hasher);
            }
        }
    }

    /// Hash a cell format for deduplication.
    fn hash_cell_format(format: &CellFormat) -> u64 {
        use std::collections::hash_map::DefaultHasher;
        use std::hash::{Hash, Hasher};

        let mut hasher = DefaultHasher::new();
        if let Some(ref font) = format.font {
            Self::hash_font(font).hash(&mut hasher);
        }
        if let Some(ref fill) = format.fill {
            Self::hash_fill(fill).hash(&mut hasher);
        }
        if let Some(ref border) = format.border {
            Self::hash_border(border).hash(&mut hasher);
        }
        if let Some(ref num_fmt) = format.number_format {
            num_fmt.hash(&mut hasher);
        }
        hasher.finish()
    }
}

impl Default for StylesBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn test_create_default_styles() {
        let builder = StylesBuilder::new();
        assert_eq!(builder.fonts.len(), 1); // Default font
        assert_eq!(builder.fills.len(), 2); // Two required default fills
        assert_eq!(builder.borders.len(), 1); // Default border
        assert_eq!(builder.cell_formats.len(), 1); // Default cell format
    }

    #[test]
    fn test_add_cell_format() {
        let mut builder = StylesBuilder::new();

        let format = CellFormat {
            font: Some(CellFont {
                bold: true,
                ..Default::default()
            }),
            ..Default::default()
        };

        let index = builder.add_cell_format(&format);
        assert_eq!(index, 1); // First custom format after default

        // Adding the same format again should return the same index
        let index2 = builder.add_cell_format(&format);
        assert_eq!(index, index2);
    }

    #[test]
    fn test_generate_xml() {
        let mut builder = StylesBuilder::new();

        // Add a custom format
        let format = CellFormat {
            font: Some(CellFont {
                bold: true,
                size: Some(12.0),
                ..Default::default()
            }),
            fill: Some(CellFill {
                pattern_type: CellFillPatternType::Solid,
                fg_color: Some("FFFF0000".to_string()),
                bg_color: None,
            }),
            ..Default::default()
        };

        builder.add_cell_format(&format);

        let xml = builder.to_xml().unwrap();
        assert!(xml.contains("<styleSheet"));
        assert!(xml.contains("<fonts"));
        assert!(xml.contains("<fills"));
        assert!(xml.contains("<borders"));
        assert!(xml.contains("<cellXfs"));
    }
}
