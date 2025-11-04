/// Shape types and implementation for PPTX presentations.
use crate::ooxml::error::{OoxmlError, Result};
use std::fmt::Write as FmtWrite;

// Import shared format types
pub use super::super::format::{ImageFormat, TextFormat};

/// Escape XML special characters.
fn escape_xml(s: &str) -> String {
    s.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
        .replace('"', "&quot;")
        .replace('\'', "&apos;")
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

    /// Get image data if this shape is a picture.
    pub(crate) fn get_image_data(&self) -> Option<(&[u8], ImageFormat)> {
        match &self.shape_type {
            ShapeType::Picture { data, format, .. } => Some((data.as_slice(), *format)),
            _ => None,
        }
    }

    /// Generate XML for this shape.
    ///
    /// For pictures, the relationship ID is optional. If not provided, a placeholder will be used.
    pub(crate) fn to_xml(&self, xml: &mut String, rel_id: Option<&str>) -> Result<()> {
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
                let rid = rel_id.unwrap_or("rIdImagePlaceholder");
                write!(xml, r#"<a:blip r:embed="{}"/>"#, rid)
                    .map_err(|e| OoxmlError::Xml(e.to_string()))?;
                xml.push_str("<a:stretch><a:fillRect/></a:stretch>");
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
        }

        Ok(())
    }
}
