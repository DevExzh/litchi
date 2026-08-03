//! Styles.xml generator for XLSX files.
//!
//! This module handles the generation of the styles.xml file, which defines
//! all the formatting information (fonts, fills, borders, number formats, and
//! cell formats) used in an Excel workbook.

use crate::xlsx::format::{CellFill, CellFillPatternType, CellFont, CellFormat};
use crate::xlsx::styles::Alignment;
use crate::xlsx::styles::border::{Border, Color, Conformance, Side};
use litchi_core::sheet::Result as SheetResult;
use litchi_core::xml::escape_xml;
use std::collections::HashMap;
use std::fmt::Write as FmtWrite;
use std::sync::Arc;

#[derive(Debug, Clone, Copy, PartialEq, Eq, Hash)]
struct CellXf {
    font_id: usize,
    fill_id: usize,
    border_id: usize,
    num_fmt_id: usize,
    alignment: Option<Alignment>,
}

impl CellXf {
    const DEFAULT: Self = Self {
        font_id: 0,
        fill_id: 0,
        border_id: 0,
        num_fmt_id: 0,
        alignment: None,
    };
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
    borders: Vec<Arc<Border>>,
    /// Full-value lookup; `Arc` shares the single owned border allocation.
    border_map: HashMap<Arc<Border>, usize>,
    /// Unique number formats (index -> format string)
    number_formats: Vec<String>,
    /// Number format lookup (format string -> index)
    number_format_map: HashMap<String, usize>,
    /// Cell format records in publication order.
    cell_formats: Vec<CellXf>,
    /// Cell format lookup by resolved resource identity.
    cell_format_map: HashMap<CellXf, usize>,
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
        let border = Arc::new(Border::default());
        builder.borders.push(Arc::clone(&border));
        builder.border_map.insert(border, 0);

        // Add default cell format (style index 0)
        builder.cell_formats.push(CellXf::DEFAULT);
        builder.cell_format_map.insert(CellXf::DEFAULT, 0);

        builder
    }

    /// Add a cell format and return its style index.
    ///
    /// If the format has already been added, returns the existing index.
    pub fn add_cell_format(&mut self, format: &CellFormat) -> usize {
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

        let key = CellXf {
            font_id,
            fill_id,
            border_id,
            num_fmt_id,
            alignment: format.alignment,
        };
        if let Some(&index) = self.cell_format_map.get(&key) {
            return index;
        }

        // Add the cell format
        let index = self.cell_formats.len();
        self.cell_formats.push(key);
        self.cell_format_map.insert(key, index);

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
    fn add_border(&mut self, border: &Border) -> usize {
        if let Some(&index) = self.border_map.get(border) {
            return index;
        }

        let index = self.borders.len();
        let border = Arc::new(border.clone());
        self.borders.push(Arc::clone(&border));
        self.border_map.insert(border, index);
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
        self.to_xml_in(Conformance::Transitional)
    }

    /// Generate styles using the requested SpreadsheetML conformance.
    pub fn to_xml_in(&self, conformance: Conformance) -> SheetResult<String> {
        let mut xml = String::with_capacity(4096);

        xml.push_str(r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#);
        let namespace = match conformance {
            Conformance::Transitional => {
                "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
            },
            Conformance::Strict => "http://purl.oclc.org/ooxml/spreadsheetml/main",
        };
        write!(xml, r#"<styleSheet xmlns="{namespace}">"#)
            .map_err(|error| format!("XML write error: {error}"))?;

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
            self.write_border(&mut xml, border, conformance)?;
        }

        xml.push_str("</borders>");

        // Write cell style XFs (required, even if empty)
        xml.push_str(r#"<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#);

        // Write cell XFs (the actual cell formats)
        write!(xml, r#"<cellXfs count="{}">"#, self.cell_formats.len())
            .map_err(|e| format!("XML write error: {}", e))?;

        for format in &self.cell_formats {
            write!(
                xml,
                r#"<xf numFmtId="{}" fontId="{}" fillId="{}" borderId="{}""#,
                format.num_fmt_id, format.font_id, format.fill_id, format.border_id
            )
            .map_err(|e| format!("XML write error: {}", e))?;

            // Add applyXXX attributes if non-default
            if format.font_id != 0 {
                xml.push_str(r#" applyFont="1""#);
            }
            if format.fill_id != 0 {
                xml.push_str(r#" applyFill="1""#);
            }
            if format.border_id != 0 {
                xml.push_str(r#" applyBorder="1""#);
            }
            if format.num_fmt_id != 0 {
                xml.push_str(r#" applyNumberFormat="1""#);
            }
            if let Some(alignment) = format.alignment {
                xml.push_str(r#" applyAlignment="1">"#);
                Self::write_alignment(&mut xml, alignment, conformance)?;
                xml.push_str("</xf>");
            } else {
                xml.push_str("/>");
            }
        }

        xml.push_str("</cellXfs>");

        // Write cell styles (required, even if minimal)
        xml.push_str(r#"<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>"#);

        // Write dxfs (differential formats) - must come AFTER cellStyles per OOXML spec
        // These are used by conditional formatting
        xml.push_str(r#"<dxfs count="0"/>"#);

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
        if let Some(underline) = font.underline {
            write!(xml, r#"<u val="{}"/>"#, underline.as_str())
                .map_err(|error| format!("XML write error: {error}"))?;
        }
        if let Some(script) = font.script {
            write!(xml, r#"<vertAlign val="{}"/>"#, script.as_str())
                .map_err(|error| format!("XML write error: {error}"))?;
        }
        if let Some(scheme) = font.scheme {
            write!(xml, r#"<scheme val="{}"/>"#, scheme.as_str())
                .map_err(|error| format!("XML write error: {error}"))?;
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

    /// Write the complete typed `CT_CellAlignment` attribute set.
    fn write_alignment(
        xml: &mut String,
        alignment: Alignment,
        conformance: Conformance,
    ) -> SheetResult<()> {
        if conformance == Conformance::Strict
            && alignment
                .text_rotation
                .is_some_and(|rotation| rotation.is_contextual())
        {
            return Err(
                "context-dependent text rotation 254 requires transitional SpreadsheetML".into(),
            );
        }

        xml.push_str("<alignment");
        if let Some(horizontal) = alignment.horizontal {
            write!(xml, r#" horizontal="{}""#, horizontal.as_str())
                .map_err(|error| format!("XML write error: {error}"))?;
        }
        if let Some(vertical) = alignment.vertical {
            write!(xml, r#" vertical="{}""#, vertical.as_str())
                .map_err(|error| format!("XML write error: {error}"))?;
        }
        if let Some(rotation) = alignment.text_rotation {
            write!(xml, r#" textRotation="{}""#, rotation.get())
                .map_err(|error| format!("XML write error: {error}"))?;
        }
        if alignment.wrap_text {
            xml.push_str(r#" wrapText="1""#);
        }
        if let Some(indent) = alignment.indent {
            write!(xml, r#" indent="{}""#, indent.get())
                .map_err(|error| format!("XML write error: {error}"))?;
        }
        if let Some(relative_indent) = alignment.relative_indent {
            write!(xml, r#" relativeIndent="{relative_indent}""#)
                .map_err(|error| format!("XML write error: {error}"))?;
        }
        if alignment.justify_last_line {
            xml.push_str(r#" justifyLastLine="1""#);
        }
        if alignment.shrink_to_fit {
            xml.push_str(r#" shrinkToFit="1""#);
        }
        if let Some(reading_order) = alignment.reading_order {
            write!(xml, r#" readingOrder="{}""#, reading_order.get())
                .map_err(|error| format!("XML write error: {error}"))?;
        }
        xml.push_str("/>");
        Ok(())
    }

    /// Write a border element to XML.
    fn write_border(
        &self,
        xml: &mut String,
        border: &Border,
        conformance: Conformance,
    ) -> SheetResult<()> {
        match conformance {
            Conformance::Transitional if border.start.is_some() || border.end.is_some() => {
                return Err("strict start/end borders require strict SpreadsheetML".into());
            },
            Conformance::Strict if border.left.is_some() || border.right.is_some() => {
                return Err(
                    "physical left/right borders require transitional SpreadsheetML".into(),
                );
            },
            _ => {},
        }

        xml.push_str("<border");
        if let Some(direction) = border.diagonal.as_ref().and_then(|diagonal| diagonal.dir()) {
            if direction.is_up() {
                xml.push_str(r#" diagonalUp="1""#);
            }
            if direction.is_down() {
                xml.push_str(r#" diagonalDown="1""#);
            }
        }
        if let Some(outline) = border.outline {
            write!(xml, r#" outline="{}""#, u8::from(outline))
                .map_err(|error| format!("XML write error: {error}"))?;
        }
        xml.push('>');

        match conformance {
            Conformance::Transitional => {
                self.write_border_side(xml, "left", border.left.as_ref())?;
                self.write_border_side(xml, "right", border.right.as_ref())?;
            },
            Conformance::Strict => {
                self.write_border_side(xml, "start", border.start.as_ref())?;
                self.write_border_side(xml, "end", border.end.as_ref())?;
            },
        }
        self.write_border_side(xml, "top", border.top.as_ref())?;
        self.write_border_side(xml, "bottom", border.bottom.as_ref())?;
        self.write_border_side(
            xml,
            "diagonal",
            border
                .diagonal
                .as_ref()
                .and_then(|diagonal| diagonal.side()),
        )?;
        self.write_border_side(xml, "vertical", border.vertical.as_ref())?;
        self.write_border_side(xml, "horizontal", border.horizontal.as_ref())?;

        xml.push_str("</border>");
        Ok(())
    }

    /// Write a single border side to XML.
    fn write_border_side(
        &self,
        xml: &mut String,
        side: &str,
        border_side: Option<&Side>,
    ) -> SheetResult<()> {
        if let Some(bs) = border_side {
            write!(xml, r#"<{} style="{}">"#, side, bs.line.as_str())
                .map_err(|e| format!("XML write error: {}", e))?;

            if let Some(color) = bs.color.as_ref() {
                Self::write_border_color(xml, color)?;
            }

            write!(xml, "</{}>", side).map_err(|e| format!("XML write error: {}", e))?;
        } else {
            write!(xml, "<{}/>", side).map_err(|e| format!("XML write error: {}", e))?;
        }

        Ok(())
    }

    fn write_border_color(xml: &mut String, color: &Color) -> SheetResult<()> {
        xml.push_str("<color");
        match color {
            Color::Default { .. } => {},
            Color::Rgb { value, .. } => {
                write!(xml, r#" rgb="{value}""#)
                    .map_err(|error| format!("XML write error: {error}"))?;
            },
            Color::Theme { index, .. } => {
                write!(xml, r#" theme="{index}""#)
                    .map_err(|error| format!("XML write error: {error}"))?;
            },
            Color::Indexed { index, .. } => {
                write!(xml, r#" indexed="{index}""#)
                    .map_err(|error| format!("XML write error: {error}"))?;
            },
            Color::Auto { enabled, .. } => {
                write!(xml, r#" auto="{}""#, u8::from(*enabled))
                    .map_err(|error| format!("XML write error: {error}"))?;
            },
        }
        if let Some(tint) = color.tint() {
            write!(xml, r#" tint="{}""#, tint.get())
                .map_err(|error| format!("XML write error: {error}"))?;
        }
        xml.push_str("/>");
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
        font.scheme.hash(&mut hasher);
        font.script.hash(&mut hasher);
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
}

impl Default for StylesBuilder {
    fn default() -> Self {
        Self::new()
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::styles::alignment::{Horizontal, Indent, Reading, Rotation, Vertical};
    use crate::xlsx::styles::border::{Diagonal, Dir, Line, Tint};
    use crate::xlsx::styles::{Scheme, Script, Styles, Underline};

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
        assert_eq!(builder.add_cell_format(&CellFormat::default()), 0);

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

    #[test]
    fn typed_font_tokens_write_and_parse_losslessly() {
        let mut builder = StylesBuilder::new();
        let format = CellFormat {
            font: Some(CellFont {
                name: Some("Aptos".to_string()),
                underline: Some(Underline::DoubleAccounting),
                scheme: Some(Scheme::Major),
                script: Some(Script::Subscript),
                ..CellFont::default()
            }),
            ..CellFormat::default()
        };
        assert_eq!(builder.add_cell_format(&format), 1);

        let xml = builder.to_xml().unwrap();
        assert!(xml.contains(r#"<u val="doubleAccounting"/>"#));
        assert!(xml.contains(r#"<vertAlign val="subscript"/>"#));
        assert!(xml.contains(r#"<scheme val="major"/>"#));

        let styles = Styles::parse(&xml).unwrap();
        let font = &styles.fonts[1];
        assert_eq!(font.underline, Some(Underline::DoubleAccounting));
        assert_eq!(font.scheme, Some(Scheme::Major));
        assert_eq!(font.script, Some(Script::Subscript));
    }

    #[test]
    fn typed_border_line_writes_its_exact_token() {
        let mut builder = StylesBuilder::new();
        let format = CellFormat {
            border: Some(Border {
                bottom: Some(
                    Side::new(Line::MediumDashDot).with_color(Color::argb(0xFF, 0x10, 0x20, 0x30)),
                ),
                diagonal: Some(Diagonal::new(Side::new(Line::Hair), Dir::Both)),
                ..Border::default()
            }),
            ..CellFormat::default()
        };

        builder.add_cell_format(&format);
        let xml = builder.to_xml().unwrap();
        assert!(xml.contains(r#"<bottom style="mediumDashDot"><color rgb="FF102030"/></bottom>"#));
        assert!(xml.contains(r#"<border diagonalUp="1" diagonalDown="1">"#));
    }

    #[test]
    fn border_dedup_uses_full_value_equality() {
        let mut builder = StylesBuilder::new();
        let left = CellFormat {
            border: Some(Border {
                left: Some(Side::new(Line::Thin)),
                ..Border::default()
            }),
            ..CellFormat::default()
        };
        let bottom = CellFormat {
            border: Some(Border {
                bottom: Some(Side::new(Line::Thin)),
                ..Border::default()
            }),
            ..CellFormat::default()
        };

        assert_ne!(
            builder.add_cell_format(&left),
            builder.add_cell_format(&bottom)
        );
        assert_eq!(builder.borders.len(), 3);
    }

    #[test]
    fn strict_border_round_trips_all_typed_fields() {
        let mut builder = StylesBuilder::new();
        let tint = Tint::new(-0.25).unwrap();
        let border = Border {
            start: Some(Side::new(Line::Thin).with_color(Color::theme(2).with_tint(tint))),
            end: Some(Side::new(Line::Dotted).with_color(Color::indexed(64))),
            top: Some(Side::new(Line::Medium).with_color(Color::auto_value(false))),
            vertical: Some(Side::new(Line::Hair).with_color(Color::rgb(1, 2, 3))),
            horizontal: Some(Side::new(Line::Double)),
            diagonal: Some(Diagonal::new(Side::new(Line::Dashed), Dir::Down)),
            outline: Some(false),
            ..Border::default()
        };
        builder.add_cell_format(&CellFormat {
            border: Some(border.clone()),
            ..CellFormat::default()
        });

        let xml = builder.to_xml_in(Conformance::Strict).unwrap();
        assert!(xml.contains("http://purl.oclc.org/ooxml/spreadsheetml/main"));
        assert!(xml.contains(r#"<border diagonalDown="1" outline="0">"#));
        assert!(xml.contains(r#"<color theme="2" tint="-0.25"/>"#));
        assert!(xml.contains(r#"<color indexed="64"/>"#));
        assert!(xml.contains(r#"<color auto="0"/>"#));
        assert!(xml.contains(r#"<color rgb="FF010203"/>"#));
        let parsed = Styles::parse(&xml).unwrap();
        assert_eq!(parsed.borders.get(1), Some(&border));
    }

    #[test]
    fn edge_names_are_checked_against_conformance() {
        let mut builder = StylesBuilder::new();
        builder.add_cell_format(&CellFormat {
            border: Some(Border {
                start: Some(Side::new(Line::Thin)),
                ..Border::default()
            }),
            ..CellFormat::default()
        });
        assert!(builder.to_xml().is_err());
    }

    #[test]
    fn typed_alignment_round_trips_every_supported_field() {
        let alignment = Alignment {
            horizontal: Some(Horizontal::CenterContinuous),
            vertical: Some(Vertical::Distributed),
            text_rotation: Some(Rotation::stacked()),
            wrap_text: true,
            indent: Some(Indent::new(255)),
            relative_indent: Some(-7),
            justify_last_line: true,
            shrink_to_fit: true,
            reading_order: Some(Reading::RightToLeft),
        };
        let mut builder = StylesBuilder::new();
        builder.add_cell_format(&CellFormat {
            alignment: Some(alignment),
            ..CellFormat::default()
        });

        for conformance in [Conformance::Transitional, Conformance::Strict] {
            let xml = builder.to_xml_in(conformance).unwrap();
            assert!(xml.contains(r#"applyAlignment="1""#));
            assert!(xml.contains(r#"horizontal="centerContinuous""#));
            assert!(xml.contains(r#"vertical="distributed""#));
            assert!(xml.contains(r#"textRotation="255""#));
            assert!(xml.contains(r#"indent="255""#));
            assert!(xml.contains(r#"relativeIndent="-7""#));
            assert!(xml.contains(r#"readingOrder="2""#));

            let parsed = Styles::parse(&xml).unwrap();
            assert_eq!(parsed.cell_xfs[1].alignment, Some(alignment));
        }
    }

    #[test]
    fn contextual_rotation_is_transitional_only() {
        let mut builder = StylesBuilder::new();
        builder.add_cell_format(&CellFormat {
            alignment: Some(Alignment {
                text_rotation: Some(Rotation::contextual()),
                ..Alignment::new()
            }),
            ..CellFormat::default()
        });

        let xml = builder.to_xml_in(Conformance::Transitional).unwrap();
        assert!(xml.contains(r#"textRotation="254""#));
        assert_eq!(
            Styles::parse(&xml).unwrap().cell_xfs[1].alignment,
            Some(Alignment {
                text_rotation: Some(Rotation::contextual()),
                ..Alignment::new()
            })
        );

        let error = builder.to_xml_in(Conformance::Strict).unwrap_err();
        assert!(
            error
                .to_string()
                .contains("context-dependent text rotation 254")
        );
    }

    #[test]
    fn alignment_participates_in_exact_cell_format_deduplication() {
        let mut builder = StylesBuilder::new();
        let left = CellFormat {
            alignment: Some(Alignment::horizontal(Horizontal::Left)),
            ..CellFormat::default()
        };
        let right = CellFormat {
            alignment: Some(Alignment::horizontal(Horizontal::Right)),
            ..CellFormat::default()
        };

        let left_id = builder.add_cell_format(&left);
        assert_eq!(builder.add_cell_format(&left), left_id);
        assert_ne!(builder.add_cell_format(&right), left_id);
        assert_ne!(
            builder.add_cell_format(&CellFormat::default()),
            builder.add_cell_format(&CellFormat {
                alignment: Some(Alignment::new()),
                ..CellFormat::default()
            })
        );
    }
}
