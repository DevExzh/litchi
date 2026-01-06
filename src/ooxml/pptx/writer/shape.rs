//! Shape types and implementation for PPTX presentations.
use crate::common::xml::escape_xml;
use crate::ooxml::drawings::blip::write_a_blip_embed;
use crate::ooxml::drawings::fill::write_a_stretch_fill_rect;
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::{ImageFormat, TextFormat};

/// Optional relationship IDs for shapes that need external references.
///
/// This struct is used to pass relationship IDs to shapes during XML generation.
/// Different shape types require different relationship IDs:
/// - Pictures: single image relationship ID
/// - Charts: single chart relationship ID
/// - SmartArt: four relationship IDs (data, layout, style, colors)
#[derive(Debug, Default, Clone)]
pub struct ShapeRelIds<'a> {
    /// Image relationship ID (for Picture shapes)
    pub image_rel_id: Option<&'a str>,
    /// Chart relationship ID (for Chart shapes)
    pub chart_rel_id: Option<&'a str>,
    /// SmartArt relationship IDs (data, layout, style, colors)
    pub smartart_rel_ids: Option<(&'a str, &'a str, &'a str, &'a str)>,
}

/// A shape on a slide (text box, image, etc.).
#[derive(Debug, Clone)]
pub struct MutableShape {
    /// Shape ID
    pub(crate) shape_id: u32,
    /// Shape type
    pub(crate) shape_type: ShapeType,
}

#[derive(Debug, Clone)]
pub(crate) enum ShapeType {
    TextBox {
        text: String,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        format: TextFormat,
    },
    Rectangle {
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    },
    Ellipse {
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    },
    Picture {
        #[allow(dead_code)]
        data: Vec<u8>,
        #[allow(dead_code)]
        format: ImageFormat,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        description: String,
    },
    /// Table shape (p:graphicFrame containing a:tbl)
    Table {
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        /// Table data: rows of cells, each cell is a string
        data: Vec<Vec<String>>,
        /// Column widths in EMUs (optional, auto-calculated if not provided)
        col_widths: Option<Vec<i64>>,
        /// Row heights in EMUs (optional, auto-calculated if not provided)
        row_heights: Option<Vec<i64>>,
        /// First row is header
        first_row: bool,
        /// Band rows (alternating row colors)
        band_row: bool,
    },
    /// Group shape containing multiple child shapes
    GroupShape {
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        /// Child shapes within the group
        children: Vec<MutableShape>,
    },
    /// Chart graphic frame (embedded chart)
    Chart {
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        /// Chart relationship ID (e.g., "rId3")
        chart_rel_id: String,
        /// Chart index (for naming chart1.xml, chart2.xml, etc.)
        chart_idx: u32,
    },
    /// SmartArt/Diagram graphic frame
    SmartArt {
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        /// Relationship IDs for the 4 required parts
        data_rel_id: String,
        layout_rel_id: String,
        style_rel_id: String,
        colors_rel_id: String,
        /// Diagram index (for naming data1.xml, layout1.xml, etc.)
        diagram_idx: u32,
    },
}

#[cfg(feature = "fonts")]
use crate::fonts::CollectGlyphs;
#[cfg(feature = "fonts")]
use roaring::RoaringBitmap;
#[cfg(feature = "fonts")]
use std::collections::HashMap;

#[cfg(feature = "fonts")]
impl CollectGlyphs for MutableShape {
    fn collect_glyphs(&self) -> HashMap<String, RoaringBitmap> {
        let mut glyphs = HashMap::new();
        match &self.shape_type {
            ShapeType::TextBox { text, format, .. } => {
                let font_name = format.font.clone().unwrap_or_else(|| "Calibri".to_string());
                let bitmap = glyphs.entry(font_name).or_insert_with(RoaringBitmap::new);
                for c in text.chars() {
                    bitmap.insert(c as u32);
                }
            },
            ShapeType::Table { data, .. } => {
                // Table cells currently only support plain strings in PPTX writer
                let font_name = "Calibri".to_string(); // Default font for tables
                let bitmap = glyphs.entry(font_name).or_insert_with(RoaringBitmap::new);
                for row in data {
                    for cell_text in row {
                        for c in cell_text.chars() {
                            bitmap.insert(c as u32);
                        }
                    }
                }
            },
            ShapeType::GroupShape { children, .. } => {
                for child in children {
                    for (font, bitmap) in child.collect_glyphs() {
                        *glyphs.entry(font).or_insert_with(RoaringBitmap::new) |= bitmap;
                    }
                }
            },
            _ => {},
        }
        glyphs
    }
}

impl MutableShape {
    /// Create a new text box shape.
    pub(crate) fn new_text_box(
        shape_id: u32,
        text: String,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
    ) -> Self {
        Self {
            shape_id,
            shape_type: ShapeType::TextBox {
                text,
                x,
                y,
                width,
                height,
                format: TextFormat::default(),
            },
        }
    }

    /// Set text formatting for this shape (only for text boxes).
    pub fn set_text_format(&mut self, format: TextFormat) -> &mut Self {
        if let ShapeType::TextBox {
            format: ref mut f, ..
        } = self.shape_type
        {
            *f = format;
        }
        self
    }

    /// Builder method: set font.
    pub fn font(&mut self, font: &str) -> &mut Self {
        if let ShapeType::TextBox {
            format: ref mut f, ..
        } = self.shape_type
        {
            f.font = Some(font.to_string());
        }
        self
    }

    /// Builder method: set font size.
    pub fn font_size(&mut self, size: f64) -> &mut Self {
        if let ShapeType::TextBox {
            format: ref mut f, ..
        } = self.shape_type
        {
            f.size = Some(size);
        }
        self
    }

    /// Builder method: set bold.
    pub fn bold(&mut self, bold: bool) -> &mut Self {
        if let ShapeType::TextBox {
            format: ref mut f, ..
        } = self.shape_type
        {
            f.bold = Some(bold);
        }
        self
    }

    /// Builder method: set italic.
    pub fn italic(&mut self, italic: bool) -> &mut Self {
        if let ShapeType::TextBox {
            format: ref mut f, ..
        } = self.shape_type
        {
            f.italic = Some(italic);
        }
        self
    }

    /// Builder method: set underline.
    pub fn underline(&mut self, underline: bool) -> &mut Self {
        if let ShapeType::TextBox {
            format: ref mut f, ..
        } = self.shape_type
        {
            f.underline = Some(underline);
        }
        self
    }

    /// Builder method: set text color.
    pub fn color(&mut self, color: &str) -> &mut Self {
        if let ShapeType::TextBox {
            format: ref mut f, ..
        } = self.shape_type
        {
            f.color = Some(color.to_string());
        }
        self
    }

    /// Create a new rectangle shape.
    pub(crate) fn new_rectangle(
        shape_id: u32,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) -> Self {
        Self {
            shape_id,
            shape_type: ShapeType::Rectangle {
                x,
                y,
                width,
                height,
                fill_color,
            },
        }
    }

    /// Create a new ellipse (circle/oval) shape.
    pub(crate) fn new_ellipse(
        shape_id: u32,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        fill_color: Option<String>,
    ) -> Self {
        Self {
            shape_id,
            shape_type: ShapeType::Ellipse {
                x,
                y,
                width,
                height,
                fill_color,
            },
        }
    }

    /// Create a new picture shape.
    #[allow(clippy::too_many_arguments)]
    pub(crate) fn new_picture(
        shape_id: u32,
        data: Vec<u8>,
        format: ImageFormat,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        description: String,
    ) -> Result<Self> {
        Ok(Self {
            shape_id,
            shape_type: ShapeType::Picture {
                data,
                format,
                x,
                y,
                width,
                height,
                description,
            },
        })
    }

    /// Create a new table shape.
    #[allow(clippy::too_many_arguments)]
    pub(crate) fn new_table(
        shape_id: u32,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        data: Vec<Vec<String>>,
        col_widths: Option<Vec<i64>>,
        row_heights: Option<Vec<i64>>,
        first_row: bool,
        band_row: bool,
    ) -> Self {
        Self {
            shape_id,
            shape_type: ShapeType::Table {
                x,
                y,
                width,
                height,
                data,
                col_widths,
                row_heights,
                first_row,
                band_row,
            },
        }
    }

    /// Create a new group shape.
    pub(crate) fn new_group(
        shape_id: u32,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        children: Vec<MutableShape>,
    ) -> Self {
        Self {
            shape_id,
            shape_type: ShapeType::GroupShape {
                x,
                y,
                width,
                height,
                children,
            },
        }
    }

    /// Get image data if this shape is a picture.
    pub(crate) fn get_image_data(&self) -> Option<(&[u8], ImageFormat)> {
        match &self.shape_type {
            ShapeType::Picture { data, format, .. } => Some((data.as_slice(), *format)),
            _ => None,
        }
    }

    /// Get child shapes if this is a group shape.
    #[allow(dead_code)] // Public API for group shape access
    pub(crate) fn get_children(&self) -> Option<&[MutableShape]> {
        match &self.shape_type {
            ShapeType::GroupShape { children, .. } => Some(children.as_slice()),
            _ => None,
        }
    }

    /// Get mutable child shapes if this is a group shape.
    pub(crate) fn get_children_mut(&mut self) -> Option<&mut Vec<MutableShape>> {
        match &mut self.shape_type {
            ShapeType::GroupShape { children, .. } => Some(children),
            _ => None,
        }
    }

    /// Generate XML for this shape.
    ///
    /// # Arguments
    /// * `xml` - The string buffer to write XML to
    /// * `rel_ids` - Optional relationship IDs for shapes that need external references
    ///
    /// For pictures, charts, and SmartArt, the relationship IDs are used to reference
    /// external parts. If not provided, placeholder values or stored values are used.
    pub(crate) fn to_xml(&self, xml: &mut String, rel_ids: ShapeRelIds<'_>) -> Result<()> {
        match &self.shape_type {
            ShapeType::TextBox {
                text,
                x,
                y,
                width,
                height,
                format,
            } => {
                xml.push_str("<p:sp>");
                xml.push_str("<p:nvSpPr>");
                write!(
                    xml,
                    r#"<p:cNvPr id="{}" name="Text Box {}"/>"#,
                    self.shape_id, self.shape_id
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("<p:cNvSpPr txBox=\"1\"/>");
                xml.push_str("<p:nvPr/>");
                xml.push_str("</p:nvSpPr>");

                xml.push_str("<p:spPr>");
                xml.push_str("<a:xfrm>");
                write!(xml, r#"<a:off x="{}" y="{}"/>"#, x, y)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                write!(xml, r#"<a:ext cx="{}" cy="{}"/>"#, width, height)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("</a:xfrm>");
                xml.push_str(r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#);
                xml.push_str("</p:spPr>");

                xml.push_str("<p:txBody>");
                xml.push_str(r#"<a:bodyPr wrap="square" rtlCol="0">"#);
                xml.push_str(r#"<a:spAutoFit/>"#);
                xml.push_str("</a:bodyPr>");
                xml.push_str("<a:lstStyle/>");
                xml.push_str("<a:p>");
                xml.push_str("<a:r>");

                xml.push_str("<a:rPr lang=\"en-US\" dirty=\"0\"");

                if let Some(size) = format.size {
                    write!(xml, " sz=\"{}\"", (size * 100.0) as u32)
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }

                if let Some(true) = format.bold {
                    xml.push_str(" b=\"1\"");
                }

                if let Some(true) = format.italic {
                    xml.push_str(" i=\"1\"");
                }

                if let Some(true) = format.underline {
                    xml.push_str(" u=\"sng\"");
                }

                xml.push('>');

                if let Some(ref font) = format.font {
                    write!(xml, "<a:latin typeface=\"{}\"/>", escape_xml(font))
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }

                if let Some(ref color) = format.color {
                    write!(
                        xml,
                        "<a:solidFill><a:srgbClr val=\"{}\"/></a:solidFill>",
                        color
                    )
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                }

                xml.push_str("</a:rPr>");

                write!(xml, "<a:t>{}</a:t>", escape_xml(text))
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("</a:r>");
                xml.push_str("</a:p>");
                xml.push_str("</p:txBody>");

                xml.push_str("</p:sp>");
            },
            ShapeType::Rectangle {
                x,
                y,
                width,
                height,
                fill_color,
            } => {
                xml.push_str("<p:sp>");
                xml.push_str("<p:nvSpPr>");
                write!(
                    xml,
                    r#"<p:cNvPr id="{}" name="Rectangle {}"/>"#,
                    self.shape_id, self.shape_id
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("<p:cNvSpPr/>");
                xml.push_str("<p:nvPr/>");
                xml.push_str("</p:nvSpPr>");

                xml.push_str("<p:spPr>");
                xml.push_str("<a:xfrm>");
                write!(xml, r#"<a:off x="{}" y="{}"/>"#, x, y)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                write!(xml, r#"<a:ext cx="{}" cy="{}"/>"#, width, height)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("</a:xfrm>");
                xml.push_str(r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#);

                if let Some(color) = fill_color {
                    xml.push_str("<a:solidFill>");
                    write!(xml, r#"<a:srgbClr val="{}"/>"#, color)
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    xml.push_str("</a:solidFill>");
                }

                xml.push_str("</p:spPr>");
                xml.push_str("</p:sp>");
            },
            ShapeType::Ellipse {
                x,
                y,
                width,
                height,
                fill_color,
            } => {
                xml.push_str("<p:sp>");
                xml.push_str("<p:nvSpPr>");
                write!(
                    xml,
                    r#"<p:cNvPr id="{}" name="Ellipse {}"/>"#,
                    self.shape_id, self.shape_id
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("<p:cNvSpPr/>");
                xml.push_str("<p:nvPr/>");
                xml.push_str("</p:nvSpPr>");

                xml.push_str("<p:spPr>");
                xml.push_str("<a:xfrm>");
                write!(xml, r#"<a:off x="{}" y="{}"/>"#, x, y)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                write!(xml, r#"<a:ext cx="{}" cy="{}"/>"#, width, height)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("</a:xfrm>");
                xml.push_str(r#"<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>"#);

                if let Some(color) = fill_color {
                    xml.push_str("<a:solidFill>");
                    write!(xml, r#"<a:srgbClr val="{}"/>"#, color)
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    xml.push_str("</a:solidFill>");
                }

                xml.push_str("</p:spPr>");
                xml.push_str("</p:sp>");
            },
            ShapeType::Picture {
                data: _,
                format: _,
                x,
                y,
                width,
                height,
                description,
            } => {
                xml.push_str("<p:pic>");
                xml.push_str("<p:nvPicPr>");
                write!(
                    xml,
                    r#"<p:cNvPr id="{}" name="Picture {}" descr="{}"/>"#,
                    self.shape_id,
                    self.shape_id,
                    escape_xml(description)
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("<p:cNvPicPr/>");
                xml.push_str("<p:nvPr/>");
                xml.push_str("</p:nvPicPr>");

                xml.push_str("<p:blipFill>");
                let rid = rel_ids.image_rel_id.unwrap_or("rIdImagePlaceholder");
                write_a_blip_embed(xml, rid, false).map_err(|e| OoxmlError::Xml(e.to_string()))?;
                write_a_stretch_fill_rect(xml);
                xml.push_str("</p:blipFill>");

                xml.push_str("<p:spPr>");
                xml.push_str("<a:xfrm>");
                write!(xml, r#"<a:off x="{}" y="{}"/>"#, x, y)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                write!(xml, r#"<a:ext cx="{}" cy="{}"/>"#, width, height)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("</a:xfrm>");
                xml.push_str(r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>"#);
                xml.push_str("</p:spPr>");
                xml.push_str("</p:pic>");
            },
            ShapeType::Table {
                x,
                y,
                width,
                height,
                data,
                col_widths,
                row_heights,
                first_row,
                band_row,
            } => {
                self.write_table_xml(
                    xml,
                    *x,
                    *y,
                    *width,
                    *height,
                    data,
                    col_widths.as_deref(),
                    row_heights.as_deref(),
                    *first_row,
                    *band_row,
                )?;
            },
            ShapeType::GroupShape {
                x,
                y,
                width,
                height,
                children,
            } => {
                self.write_group_xml(xml, *x, *y, *width, *height, children, rel_ids.image_rel_id)?;
            },
            ShapeType::Chart {
                x,
                y,
                width,
                height,
                chart_rel_id,
                chart_idx,
            } => {
                // Use provided relationship ID if available, otherwise use stored one
                let actual_rel_id = rel_ids.chart_rel_id.unwrap_or(chart_rel_id.as_str());

                // Chart graphicFrame
                xml.push_str("<p:graphicFrame>");
                xml.push_str("<p:nvGraphicFramePr>");
                write!(
                    xml,
                    r#"<p:cNvPr id="{}" name="Chart {}"/>"#,
                    self.shape_id, chart_idx
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("<p:cNvGraphicFramePr/>");
                xml.push_str("<p:nvPr/>");
                xml.push_str("</p:nvGraphicFramePr>");

                xml.push_str("<p:xfrm>");
                write!(xml, r#"<a:off x="{}" y="{}"/>"#, x, y)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                write!(xml, r#"<a:ext cx="{}" cy="{}"/>"#, width, height)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("</p:xfrm>");

                xml.push_str(r#"<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#);
                xml.push_str(r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">"#);
                write!(
                    xml,
                    r#"<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="{}"/>"#,
                    actual_rel_id
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("</a:graphicData>");
                xml.push_str("</a:graphic>");
                xml.push_str("</p:graphicFrame>");
            },
            ShapeType::SmartArt {
                x,
                y,
                width,
                height,
                data_rel_id,
                layout_rel_id,
                style_rel_id,
                colors_rel_id,
                diagram_idx,
            } => {
                // Use provided relationship IDs if available, otherwise use stored ones
                let (actual_data_id, actual_layout_id, actual_style_id, actual_colors_id) =
                    if let Some((d, l, s, c)) = rel_ids.smartart_rel_ids {
                        (d, l, s, c)
                    } else {
                        (
                            data_rel_id.as_str(),
                            layout_rel_id.as_str(),
                            style_rel_id.as_str(),
                            colors_rel_id.as_str(),
                        )
                    };

                // SmartArt graphicFrame
                xml.push_str("<p:graphicFrame>");
                xml.push_str("<p:nvGraphicFramePr>");
                write!(
                    xml,
                    r#"<p:cNvPr id="{}" name="Diagram {}"/>"#,
                    self.shape_id, diagram_idx
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("<p:cNvGraphicFramePr/>");
                xml.push_str("<p:nvPr/>");
                xml.push_str("</p:nvGraphicFramePr>");

                xml.push_str("<p:xfrm>");
                write!(xml, r#"<a:off x="{}" y="{}"/>"#, x, y)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                write!(xml, r#"<a:ext cx="{}" cy="{}"/>"#, width, height)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("</p:xfrm>");

                xml.push_str(r#"<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">"#);
                xml.push_str(r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">"#);
                write!(
                    xml,
                    r#"<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:dm="{}" r:lo="{}" r:qs="{}" r:cs="{}"/>"#,
                    actual_data_id, actual_layout_id, actual_style_id, actual_colors_id
                )
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("</a:graphicData>");
                xml.push_str("</a:graphic>");
                xml.push_str("</p:graphicFrame>");
            },
        }

        Ok(())
    }

    /// Write table XML (p:graphicFrame containing a:tbl).
    #[allow(clippy::too_many_arguments)]
    fn write_table_xml(
        &self,
        xml: &mut String,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        data: &[Vec<String>],
        col_widths: Option<&[i64]>,
        row_heights: Option<&[i64]>,
        first_row: bool,
        band_row: bool,
    ) -> Result<()> {
        let num_rows = data.len();
        let num_cols = data.first().map(|r| r.len()).unwrap_or(0);

        if num_rows == 0 || num_cols == 0 {
            return Ok(());
        }

        // Calculate column widths if not provided
        let calculated_col_widths: Vec<i64> = col_widths
            .map(|w| w.to_vec())
            .unwrap_or_else(|| vec![width / num_cols as i64; num_cols]);

        // Calculate row heights if not provided
        let calculated_row_heights: Vec<i64> = row_heights
            .map(|h| h.to_vec())
            .unwrap_or_else(|| vec![height / num_rows as i64; num_rows]);

        // Start graphic frame
        xml.push_str("<p:graphicFrame>");
        xml.push_str("<p:nvGraphicFramePr>");
        write!(
            xml,
            r#"<p:cNvPr id="{}" name="Table {}"/>"#,
            self.shape_id, self.shape_id
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str(
            r#"<p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr>"#,
        );
        xml.push_str("<p:nvPr/>");
        xml.push_str("</p:nvGraphicFramePr>");

        // Transform
        xml.push_str("<p:xfrm>");
        write!(xml, r#"<a:off x="{}" y="{}"/>"#, x, y)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        write!(xml, r#"<a:ext cx="{}" cy="{}"/>"#, width, height)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str("</p:xfrm>");

        // Graphic element containing table
        xml.push_str("<a:graphic>");
        xml.push_str(
            r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">"#,
        );

        // Table element
        xml.push_str("<a:tbl>");

        // Table properties
        xml.push_str("<a:tblPr");
        if first_row {
            xml.push_str(" firstRow=\"1\"");
        }
        if band_row {
            xml.push_str(" bandRow=\"1\"");
        }
        xml.push('>');
        // Default table style
        xml.push_str("<a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>");
        xml.push_str("</a:tblPr>");

        // Table grid (column definitions)
        xml.push_str("<a:tblGrid>");
        for col_width in &calculated_col_widths {
            write!(xml, r#"<a:gridCol w="{}"/>"#, col_width)
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        }
        xml.push_str("</a:tblGrid>");

        // Table rows
        for (row_idx, row) in data.iter().enumerate() {
            let row_height = calculated_row_heights
                .get(row_idx)
                .copied()
                .unwrap_or(370840);
            write!(xml, r#"<a:tr h="{}">"#, row_height)
                .map_err(|e| OoxmlError::Xml(e.to_string()))?;

            for cell_text in row {
                xml.push_str("<a:tc>");
                xml.push_str("<a:txBody>");
                xml.push_str("<a:bodyPr/>");
                xml.push_str("<a:lstStyle/>");
                xml.push_str("<a:p>");
                if !cell_text.is_empty() {
                    xml.push_str("<a:r>");
                    xml.push_str(r#"<a:rPr lang="en-US" dirty="0"/>"#);
                    write!(xml, "<a:t>{}</a:t>", escape_xml(cell_text))
                        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                    xml.push_str("</a:r>");
                } else {
                    xml.push_str(r#"<a:endParaRPr lang="en-US"/>"#);
                }
                xml.push_str("</a:p>");
                xml.push_str("</a:txBody>");
                xml.push_str("<a:tcPr/>");
                xml.push_str("</a:tc>");
            }

            xml.push_str("</a:tr>");
        }

        xml.push_str("</a:tbl>");
        xml.push_str("</a:graphicData>");
        xml.push_str("</a:graphic>");
        xml.push_str("</p:graphicFrame>");

        Ok(())
    }

    /// Write group shape XML (p:grpSp).
    #[allow(clippy::too_many_arguments)]
    fn write_group_xml(
        &self,
        xml: &mut String,
        x: i64,
        y: i64,
        width: i64,
        height: i64,
        children: &[MutableShape],
        image_rel_id: Option<&str>,
    ) -> Result<()> {
        xml.push_str("<p:grpSp>");

        // Non-visual group shape properties
        xml.push_str("<p:nvGrpSpPr>");
        write!(
            xml,
            r#"<p:cNvPr id="{}" name="Group {}"/>"#,
            self.shape_id, self.shape_id
        )
        .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str("<p:cNvGrpSpPr/>");
        xml.push_str("<p:nvPr/>");
        xml.push_str("</p:nvGrpSpPr>");

        // Group shape properties with transforms
        xml.push_str("<p:grpSpPr>");
        xml.push_str("<a:xfrm>");
        write!(xml, r#"<a:off x="{}" y="{}"/>"#, x, y)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        write!(xml, r#"<a:ext cx="{}" cy="{}"/>"#, width, height)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        // Child offset and extents (same as outer for 1:1 mapping)
        write!(xml, r#"<a:chOff x="{}" y="{}"/>"#, x, y)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        write!(xml, r#"<a:chExt cx="{}" cy="{}"/>"#, width, height)
            .map_err(|e| OoxmlError::Xml(e.to_string()))?;
        xml.push_str("</a:xfrm>");
        xml.push_str("</p:grpSpPr>");

        // Write child shapes
        // Note: Group children currently only support image relationship IDs
        // Charts and SmartArt in groups would need additional handling
        for child in children {
            let rel_ids = ShapeRelIds {
                image_rel_id,
                ..Default::default()
            };
            child.to_xml(xml, rel_ids)?;
        }

        xml.push_str("</p:grpSp>");

        Ok(())
    }
}
