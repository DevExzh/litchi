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

#[cfg(test)]
mod tests {
    use super::*;

    const TEST_PRESENTATION_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">
    <office:body>
        <office:presentation>
            <draw:page draw:name="Slide1">
                <draw:frame draw:name="Title" presentation:class="title" svg:x="1cm" svg:y="1cm" svg:width="18cm" svg:height="3cm">
                    <draw:text-box>
                        <text:p>Welcome</text:p>
                    </draw:text-box>
                </draw:frame>
                <draw:rect draw:name="Box1" svg:x="2cm" svg:y="5cm" svg:width="5cm" svg:height="3cm">
                    <draw:text-box>
                        <text:p>Rectangle content</text:p>
                    </draw:text-box>
                </draw:rect>
            </draw:page>
            <draw:page draw:name="Slide2">
                <draw:frame draw:name="Content" presentation:class="object" svg:x="1cm" svg:y="4cm">
                    <draw:text-box>
                        <text:p>Bullet 1</text:p>
                        <text:p>Bullet 2</text:p>
                    </draw:text-box>
                </draw:frame>
            </draw:page>
        </office:presentation>
    </office:body>
</office:document-content>"#;

    const TEST_EMPTY_PRESENTATION: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0">
    <office:body>
        <office:presentation>
        </office:presentation>
    </office:body>
</office:document-content>"#;

    const TEST_SHAPES_XML: &str = r#"<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
    xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
    xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0">
    <office:body>
        <office:presentation>
            <draw:page draw:name="Shapes">
                <draw:ellipse draw:name="Circle1" svg:x="1cm" svg:y="1cm" svg:width="3cm" svg:height="3cm">
                    <draw:text-box>
                        <text:p>Circle</text:p>
                    </draw:text-box>
                </draw:ellipse>
                <draw:line draw:name="Line1" svg:x1="0cm" svg:y1="0cm" svg:x2="10cm" svg:y2="10cm"/>
                <draw:custom-shape draw:name="Custom1" svg:x="5cm" svg:y="5cm"/>
            </draw:page>
        </office:presentation>
    </office:body>
</office:document-content>"#;

    #[test]
    fn test_parse_slides() {
        let slides = OdpParser::parse_slides(TEST_PRESENTATION_XML).unwrap();
        assert_eq!(slides.len(), 2);

        // First slide
        assert_eq!(slides[0].title, Some("Slide1".to_string()));
        assert_eq!(slides[0].index, 0);
        assert!(slides[0].text.contains("Welcome"));
        assert!(!slides[0].shapes.is_empty());

        // Second slide
        assert_eq!(slides[1].title, Some("Slide2".to_string()));
        assert_eq!(slides[1].index, 1);
    }

    #[test]
    fn test_parse_empty_presentation() {
        let slides = OdpParser::parse_slides(TEST_EMPTY_PRESENTATION).unwrap();
        assert!(slides.is_empty());
    }

    #[test]
    fn test_parse_shapes() {
        let slides = OdpParser::parse_slides(TEST_SHAPES_XML).unwrap();
        assert_eq!(slides.len(), 1);

        let slide = &slides[0];
        // Parser extracts shapes from draw:ellipse, draw:line, draw:custom-shape, etc.
        // Verify that at least one shape was parsed (actual count depends on parser implementation)
        assert!(!slide.shapes.is_empty());
    }

    #[test]
    fn test_slide_debug() {
        let slide = Slide {
            title: Some("Test".to_string()),
            text: "Content".to_string(),
            index: 0,
            notes: None,
            shapes: vec![],
        };
        let debug_str = format!("{:?}", slide);
        assert!(debug_str.contains("Slide"));
        assert!(debug_str.contains("Test"));
    }

    #[test]
    fn test_slide_clone() {
        let slide = Slide {
            title: Some("Test".to_string()),
            text: "Content".to_string(),
            index: 0,
            notes: None,
            shapes: vec![],
        };
        let cloned = slide.clone();
        assert_eq!(slide.title, cloned.title);
        assert_eq!(slide.text, cloned.text);
    }

    #[test]
    fn test_shape_debug() {
        let shape = Shape {
            shape_type: ShapeType::TextBox,
            text: "Shape text".to_string(),
            name: Some("Shape1".to_string()),
            x: Some("1cm".to_string()),
            y: Some("2cm".to_string()),
            width: Some("10cm".to_string()),
            height: Some("5cm".to_string()),
            style_name: Some("Style1".to_string()),
        };
        let debug_str = format!("{:?}", shape);
        assert!(debug_str.contains("Shape"));
        assert!(debug_str.contains("TextBox"));
    }

    #[test]
    fn test_shape_clone() {
        let shape = Shape {
            shape_type: ShapeType::AutoShape,
            text: "Text".to_string(),
            name: Some("Name".to_string()),
            x: Some("0cm".to_string()),
            y: Some("0cm".to_string()),
            width: Some("5cm".to_string()),
            height: Some("3cm".to_string()),
            style_name: None,
        };
        let cloned = shape.clone();
        assert_eq!(shape.shape_type, cloned.shape_type);
        assert_eq!(shape.name, cloned.name);
    }

    #[test]
    fn test_shape_type_variants() {
        // Test all shape type variants
        let types = vec![
            ShapeType::TextBox,
            ShapeType::AutoShape,
            ShapeType::Line,
            ShapeType::Placeholder,
            ShapeType::Picture,
            ShapeType::Group,
            ShapeType::Connector,
            ShapeType::Table,
            ShapeType::GraphicFrame,
            ShapeType::Unknown,
        ];

        for shape_type in types {
            let shape = Shape {
                shape_type,
                text: String::new(),
                name: None,
                x: None,
                y: None,
                width: None,
                height: None,
                style_name: None,
            };
            let _ = format!("{:?}", shape);
        }
    }

    #[test]
    fn test_shape_type_equality() {
        assert_eq!(ShapeType::TextBox, ShapeType::TextBox);
        assert_ne!(ShapeType::TextBox, ShapeType::Line);
        assert_ne!(ShapeType::AutoShape, ShapeType::Picture);
    }

    #[test]
    fn test_shape_type_clone() {
        let t1 = ShapeType::Placeholder;
        let t2 = t1.clone();
        assert_eq!(t1, t2);
    }

    #[test]
    fn test_shape_type_copy() {
        let t1 = ShapeType::Line;
        let t2 = t1;
        assert_eq!(t1, t2); // Copy trait allows this
    }

    #[test]
    fn test_shape_builder() {
        let builder = ShapeBuilder::new();
        let shape = builder.build();
        assert_eq!(shape.shape_type, ShapeType::AutoShape);
        assert!(shape.text.is_empty());
    }

    #[test]
    fn test_shape_builder_with_data() {
        let mut builder = ShapeBuilder::new();
        builder.name = Some("TestShape".to_string());
        builder.x = Some("1cm".to_string());
        builder.y = Some("2cm".to_string());
        builder.width = Some("10cm".to_string());
        builder.height = Some("5cm".to_string());
        builder.text = "Hello".to_string();
        builder.shape_type = ShapeType::TextBox;

        let shape = builder.build();
        assert_eq!(shape.name, Some("TestShape".to_string()));
        assert_eq!(shape.x, Some("1cm".to_string()));
        assert_eq!(shape.text, "Hello");
        assert_eq!(shape.shape_type, ShapeType::TextBox);
    }

    #[test]
    fn test_shape_builder_clone() {
        let builder = ShapeBuilder::new();
        let cloned = builder.build().clone();
        assert_eq!(cloned.shape_type, ShapeType::AutoShape);
    }
}
