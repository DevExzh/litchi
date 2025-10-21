//! ODP-specific parsing utilities.

use super::{Shape, Slide};
use crate::common::{Error, Result, ShapeType};
use quick_xml::Reader;
use quick_xml::events::Event;

/// Parser for ODP-specific structures.
///
/// This provides parsing logic specific to presentations,
/// including slide and shape parsing.
pub(crate) struct OdpParser;

/// Internal structure for building shapes during parsing
#[allow(dead_code)]
struct ShapeBuilder {
    shape_type: ShapeType,
    text: String,
    name: Option<String>,
    x: Option<String>,
    y: Option<String>,
    width: Option<String>,
    height: Option<String>,
    style_name: Option<String>,
}

#[allow(dead_code)]
impl ShapeBuilder {
    fn new() -> Self {
        Self {
            shape_type: ShapeType::AutoShape,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
        }
    }

    fn build(self) -> Shape {
        Shape {
            shape_type: self.shape_type,
            text: self.text,
            name: self.name,
            x: self.x,
            y: self.y,
            width: self.width,
            height: self.height,
            style_name: self.style_name,
        }
    }
}

impl OdpParser {
    /// Parse all slides from ODP content.xml
    pub fn parse_slides(xml_content: &str) -> Result<Vec<Slide>> {
        let mut reader = Reader::from_str(xml_content);
        reader.config_mut().trim_text(true);
        let mut buf = Vec::new();
        let mut slides = Vec::new();

        // State tracking
        let mut current_slide_text = String::new();
        let mut current_slide_title: Option<String> = None;
        let mut current_shapes: Vec<Shape> = Vec::new();
        let mut in_slide = false;
        let mut slide_index = 0;

        // Shape parsing state
        let mut current_shape: Option<ShapeBuilder> = None;
        let mut in_text_box = false;
        let mut shape_depth = 0;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    match e.name().as_ref() {
                        b"draw:page" => {
                            // Finish previous slide if any
                            if in_slide {
                                slides.push(Slide {
                                    title: current_slide_title.take(),
                                    text: std::mem::take(&mut current_slide_text)
                                        .trim()
                                        .to_string(),
                                    index: slide_index,
                                    notes: None,
                                    shapes: std::mem::take(&mut current_shapes),
                                });
                                slide_index += 1;
                            }

                            // Start new slide
                            current_slide_title = Self::get_attr(e.attributes(), b"draw:name")
                                .or_else(|| Some(format!("Slide{}", slide_index + 1)));
                            in_slide = true;
                        },
                        b"draw:frame" | b"draw:rect" | b"draw:ellipse" | b"draw:line"
                        | b"draw:custom-shape" | b"draw:circle" | b"draw:path"
                        | b"draw:polygon" | b"draw:polyline" => {
                            if in_slide && current_shape.is_none() {
                                let mut builder = ShapeBuilder::new();

                                // Determine shape type
                                builder.shape_type = match e.name().as_ref() {
                                    b"draw:frame" => {
                                        // Check presentation:class to determine if it's a placeholder
                                        if let Some(pres_class) =
                                            Self::get_attr(e.attributes(), b"presentation:class")
                                        {
                                            match pres_class.as_str() {
                                                "title" | "subtitle" | "object" => {
                                                    ShapeType::Placeholder
                                                },
                                                _ => ShapeType::TextBox,
                                            }
                                        } else {
                                            ShapeType::TextBox
                                        }
                                    },
                                    b"draw:rect" | b"draw:ellipse" | b"draw:circle" => {
                                        ShapeType::AutoShape
                                    },
                                    b"draw:line" => ShapeType::Line,
                                    b"draw:custom-shape" | b"draw:path" | b"draw:polygon"
                                    | b"draw:polyline" => ShapeType::AutoShape,
                                    _ => ShapeType::AutoShape,
                                };

                                // Extract attributes
                                builder.name = Self::get_attr(e.attributes(), b"draw:name");
                                builder.x = Self::get_attr(e.attributes(), b"svg:x");
                                builder.y = Self::get_attr(e.attributes(), b"svg:y");
                                builder.width = Self::get_attr(e.attributes(), b"svg:width");
                                builder.height = Self::get_attr(e.attributes(), b"svg:height");
                                builder.style_name = Self::get_attr(
                                    e.attributes(),
                                    b"draw:style-name",
                                )
                                .or_else(|| {
                                    Self::get_attr(e.attributes(), b"presentation:style-name")
                                });

                                current_shape = Some(builder);
                                shape_depth = 0;
                            } else if current_shape.is_some() {
                                shape_depth += 1;
                            }
                        },
                        b"draw:text-box" => {
                            if current_shape.is_some() {
                                in_text_box = true;
                            }
                        },
                        b"text:p" | b"text:span" => {
                            // Text will be collected in Text event
                        },
                        _ => {},
                    }
                },
                Ok(Event::Text(ref t)) => {
                    if in_slide && let Ok(text) = String::from_utf8(t.to_vec()) {
                        let trimmed = text.trim();
                        if !trimmed.is_empty() {
                            // Add to slide text
                            if !current_slide_text.is_empty() {
                                current_slide_text.push(' ');
                            }
                            current_slide_text.push_str(trimmed);

                            // Add to shape text if in a shape
                            if let Some(ref mut shape) = current_shape
                                && in_text_box
                            {
                                if !shape.text.is_empty() {
                                    shape.text.push(' ');
                                }
                                shape.text.push_str(trimmed);
                            }
                        }
                    }
                },
                Ok(Event::End(ref e)) => {
                    match e.name().as_ref() {
                        b"draw:page" => {
                            // Finish current slide
                            if in_slide {
                                slides.push(Slide {
                                    title: current_slide_title.take(),
                                    text: std::mem::take(&mut current_slide_text)
                                        .trim()
                                        .to_string(),
                                    index: slide_index,
                                    notes: None,
                                    shapes: std::mem::take(&mut current_shapes),
                                });
                                slide_index += 1;
                            }
                            in_slide = false;
                        },
                        b"draw:frame" | b"draw:rect" | b"draw:ellipse" | b"draw:line"
                        | b"draw:custom-shape" | b"draw:circle" | b"draw:path"
                        | b"draw:polygon" | b"draw:polyline" => {
                            if shape_depth > 0 {
                                shape_depth -= 1;
                            } else if let Some(builder) = current_shape.take() {
                                // Finish the shape and add it to the slide
                                current_shapes.push(builder.build());
                                in_text_box = false;
                            }
                        },
                        b"draw:text-box" => {
                            in_text_box = false;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Eof) => break,
                Err(e) => {
                    return Err(Error::InvalidFormat(format!("XML parsing error: {}", e)));
                },
                _ => {},
            }
            buf.clear();
        }

        Ok(slides)
    }

    /// Helper to extract attribute values
    fn get_attr(attrs: quick_xml::events::attributes::Attributes, name: &[u8]) -> Option<String> {
        for attr_result in attrs {
            if let Ok(attr) = attr_result
                && attr.key.as_ref() == name
            {
                return String::from_utf8(attr.value.to_vec()).ok();
            }
        }
        None
    }
}
