//! Mutable, package-independent slide-shape authoring values.

use crate::format::TextFormat;
use crate::{Error, Result};

/// Optional relationship identifiers supplied by a package writer.
#[derive(Debug, Default, Clone, Copy)]
pub struct ShapeRelIds<'a> {
    /// Image relationship identifier.
    pub image_rel_id: Option<&'a str>,
    /// Chart relationship identifier.
    pub chart_rel_id: Option<&'a str>,
    /// SmartArt relationship identifiers.
    pub smartart_rel_ids: Option<(&'a str, &'a str, &'a str, &'a str)>,
}

/// The bounded set of shape kinds emitted by the standalone writer slice.
#[derive(Debug, Clone)]
#[non_exhaustive]
pub enum ShapeType {
    /// A text box with optional text formatting.
    TextBox {
        /// Plain text payload.
        text: String,
        /// X position in EMUs.
        x: i64,
        /// Y position in EMUs.
        y: i64,
        /// Width in EMUs.
        width: i64,
        /// Height in EMUs.
        height: i64,
        /// Text formatting.
        format: TextFormat,
    },
    /// A preset rectangle.
    Rectangle {
        /// X position in EMUs.
        x: i64,
        /// Y position in EMUs.
        y: i64,
        /// Width in EMUs.
        width: i64,
        /// Height in EMUs.
        height: i64,
        /// Optional sRGB fill.
        fill_color: Option<String>,
    },
    /// A preset ellipse.
    Ellipse {
        /// X position in EMUs.
        x: i64,
        /// Y position in EMUs.
        y: i64,
        /// Width in EMUs.
        width: i64,
        /// Height in EMUs.
        height: i64,
        /// Optional sRGB fill.
        fill_color: Option<String>,
    },
}

/// One mutable shape in a new presentation.
#[derive(Debug, Clone)]
pub struct MutableShape {
    pub(crate) shape_id: u32,
    pub(crate) shape_type: ShapeType,
    pub(crate) modified: bool,
}

impl MutableShape {
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
            modified: false,
        }
    }

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
            modified: false,
        }
    }

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
            modified: false,
        }
    }

    /// The stable non-visual shape ID.
    #[inline]
    pub fn shape_id(&self) -> u32 {
        self.shape_id
    }

    /// Borrow the current shape kind.
    #[inline]
    pub fn shape_type(&self) -> &ShapeType {
        &self.shape_type
    }

    /// Replace text in a text box.
    pub fn set_text(&mut self, text: impl Into<String>) -> &mut Self {
        if let ShapeType::TextBox { text: value, .. } = &mut self.shape_type {
            *value = text.into();
            self.modified = true;
        }
        self
    }

    /// Return text for a text box.
    pub fn text(&self) -> Option<&str> {
        match &self.shape_type {
            ShapeType::TextBox { text, .. } => Some(text),
            _ => None,
        }
    }

    /// Replace all text formatting properties.
    pub fn set_text_format(&mut self, format: TextFormat) -> &mut Self {
        if let ShapeType::TextBox { format: value, .. } = &mut self.shape_type {
            *value = format;
            self.modified = true;
        }
        self
    }

    /// Set the font family.
    pub fn font(&mut self, font: &str) -> &mut Self {
        if let ShapeType::TextBox { format, .. } = &mut self.shape_type {
            format.font = Some(font.to_string());
            self.modified = true;
        }
        self
    }

    /// Set the font size in points.
    pub fn font_size(&mut self, size: f64) -> &mut Self {
        if let ShapeType::TextBox { format, .. } = &mut self.shape_type {
            format.size = Some(size);
            self.modified = true;
        }
        self
    }

    /// Set bold text.
    pub fn bold(&mut self, bold: bool) -> &mut Self {
        if let ShapeType::TextBox { format, .. } = &mut self.shape_type {
            format.bold = Some(bold);
            self.modified = true;
        }
        self
    }

    /// Set italic text.
    pub fn italic(&mut self, italic: bool) -> &mut Self {
        if let ShapeType::TextBox { format, .. } = &mut self.shape_type {
            format.italic = Some(italic);
            self.modified = true;
        }
        self
    }

    /// Set underlined text.
    pub fn underline(&mut self, underline: bool) -> &mut Self {
        if let ShapeType::TextBox { format, .. } = &mut self.shape_type {
            format.underline = Some(underline);
            self.modified = true;
        }
        self
    }

    /// Set the text sRGB color.
    pub fn color(&mut self, color: &str) -> &mut Self {
        if let ShapeType::TextBox { format, .. } = &mut self.shape_type {
            format.color = Some(color.to_string());
            self.modified = true;
        }
        self
    }

    /// Return the shape as PresentationML owned by a slide part.
    pub(crate) fn to_xml(&self) -> Result<String> {
        let mut xml = String::new();
        match &self.shape_type {
            ShapeType::TextBox {
                text,
                x,
                y,
                width,
                height,
                format,
            } => {
                xml.push_str(&shape_start(
                    self.shape_id,
                    "TextBox",
                    *x,
                    *y,
                    *width,
                    *height,
                ));
                xml.push_str("<p:nvSpPr><p:cNvPr id=\"");
                xml.push_str(&self.shape_id.to_string());
                xml.push_str("\" name=\"TextBox\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>");
                xml.push_str("<p:spPr><a:xfrm><a:off x=\"");
                xml.push_str(&x.to_string());
                xml.push_str("\" y=\"");
                xml.push_str(&y.to_string());
                xml.push_str("\"/><a:ext cx=\"");
                xml.push_str(&width.to_string());
                xml.push_str("\" cy=\"");
                xml.push_str(&height.to_string());
                xml.push_str(
                    "\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></p:spPr>",
                );
                xml.push_str("<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"");
                if format.bold == Some(true) {
                    xml.push_str(" b=\"1\"");
                }
                if format.italic == Some(true) {
                    xml.push_str(" i=\"1\"");
                }
                if format.underline == Some(true) {
                    xml.push_str(" u=\"sng\"");
                }
                if let Some(size) = format.size {
                    if !size.is_finite() || size < 0.0 || size > 400.0 {
                        return Err(Error::Invalid(
                            "text font size is outside the bounded domain".into(),
                        ));
                    }
                    xml.push_str(" sz=\"");
                    xml.push_str(&(size * 100.0).round().to_string());
                    xml.push_str("\"");
                }
                if let Some(color) = &format.color {
                    xml.push_str("><a:solidFill><a:srgbClr val=\"");
                    xml.push_str(&escape_xml(color));
                    xml.push_str("\"/></a:solidFill><a:endParaRPr lang=\"en-US\"/></a:rPr><a:t>");
                } else {
                    xml.push_str("/><a:t>");
                }
                xml.push_str(&escape_xml(text));
                xml.push_str("</a:t></a:r></a:p></p:txBody></p:sp>");
            },
            ShapeType::Rectangle {
                x,
                y,
                width,
                height,
                fill_color,
            } => write_preset_shape(
                &mut xml,
                self.shape_id,
                "Rectangle",
                "rect",
                *x,
                *y,
                *width,
                *height,
                fill_color.as_deref(),
            ),
            ShapeType::Ellipse {
                x,
                y,
                width,
                height,
                fill_color,
            } => write_preset_shape(
                &mut xml,
                self.shape_id,
                "Ellipse",
                "ellipse",
                *x,
                *y,
                *width,
                *height,
                fill_color.as_deref(),
            ),
        }
        Ok(xml)
    }

    pub(crate) fn is_modified(&self) -> bool {
        self.modified
    }

    pub(crate) fn mark_clean(&mut self) {
        self.modified = false;
    }
}

fn shape_start(id: u32, name: &str, _x: i64, _y: i64, _width: i64, _height: i64) -> String {
    let _ = (id, name);
    "<p:sp>".to_string()
}

fn write_preset_shape(
    xml: &mut String,
    id: u32,
    name: &str,
    preset: &str,
    x: i64,
    y: i64,
    width: i64,
    height: i64,
    fill_color: Option<&str>,
) {
    xml.push_str(&shape_start(id, name, x, y, width, height));
    xml.push_str("<p:nvSpPr><p:cNvPr id=\"");
    xml.push_str(&id.to_string());
    xml.push_str("\" name=\"");
    xml.push_str(&escape_xml(name));
    xml.push_str("\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x=\"");
    xml.push_str(&x.to_string());
    xml.push_str("\" y=\"");
    xml.push_str(&y.to_string());
    xml.push_str("\"/><a:ext cx=\"");
    xml.push_str(&width.to_string());
    xml.push_str("\" cy=\"");
    xml.push_str(&height.to_string());
    xml.push_str("\"/></a:xfrm><a:prstGeom prst=\"");
    xml.push_str(preset);
    xml.push_str("\"><a:avLst/></a:prstGeom>");
    if let Some(color) = fill_color {
        xml.push_str("<a:solidFill><a:srgbClr val=\"");
        xml.push_str(&escape_xml(color));
        xml.push_str("\"/></a:solidFill>");
    }
    xml.push_str("</p:spPr></p:sp>");
}

pub(crate) fn escape_xml(value: &str) -> String {
    let mut escaped = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            _ => escaped.push(character),
        }
    }
    escaped
}
