//! ODP-specific parsing utilities.

use super::{Shape, Slide};
use litchi_core::{Error, Result, ShapeType};
use quick_xml::Reader;
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event, attributes::Attribute};

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
    image_href: Option<String>,
    is_title: bool,
    has_paragraph: bool,
}

#[derive(Default)]
struct ParagraphText {
    value: String,
    trailing_collapsible_space: bool,
}

impl ParagraphText {
    fn push_text(&mut self, text: &str) {
        for character in text.chars() {
            if character.is_whitespace() {
                if !self.value.is_empty()
                    && !self
                        .value
                        .chars()
                        .next_back()
                        .is_some_and(char::is_whitespace)
                {
                    self.value.push(' ');
                    self.trailing_collapsible_space = true;
                }
            } else {
                self.value.push(character);
                self.trailing_collapsible_space = false;
            }
        }
    }

    fn push_explicit(&mut self, character: char, count: usize) {
        self.value.extend(std::iter::repeat_n(character, count));
        self.trailing_collapsible_space = false;
    }

    fn finish(mut self) -> String {
        if self.trailing_collapsible_space {
            self.value.pop();
        }
        self.value
    }
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
            image_href: None,
            is_title: false,
            has_paragraph: false,
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
            image_href: self.image_href,
        }
    }

    fn push_paragraph(&mut self, text: &str) {
        if self.has_paragraph {
            self.text.push('\n');
        }
        self.text.push_str(text);
        self.has_paragraph = true;
    }
}

impl OdpParser {
    fn is_shape_element(name: &[u8]) -> bool {
        matches!(
            name,
            b"draw:frame"
                | b"draw:rect"
                | b"draw:ellipse"
                | b"draw:line"
                | b"draw:custom-shape"
                | b"draw:circle"
                | b"draw:path"
                | b"draw:polygon"
                | b"draw:polyline"
                | b"draw:connector"
                | b"draw:g"
        )
    }

    fn shape_builder(element: &BytesStart<'_>) -> ShapeBuilder {
        let mut builder = ShapeBuilder::new();
        let presentation_class = Self::get_attr(element.attributes(), b"presentation:class");
        builder.is_title = presentation_class.as_deref() == Some("title");
        builder.shape_type = match element.name().as_ref() {
            b"draw:frame" => match presentation_class.as_deref() {
                Some(_) => ShapeType::Placeholder,
                _ => ShapeType::TextBox,
            },
            b"draw:line" => ShapeType::Line,
            b"draw:connector" => ShapeType::Connector,
            b"draw:g" => ShapeType::Group,
            _ => ShapeType::AutoShape,
        };
        builder.name = Self::get_attr(element.attributes(), b"draw:name");
        if matches!(element.name().as_ref(), b"draw:line" | b"draw:connector") {
            builder.x = Self::get_attr(element.attributes(), b"svg:x1");
            builder.y = Self::get_attr(element.attributes(), b"svg:y1");
            builder.width = Self::get_attr(element.attributes(), b"svg:x2");
            builder.height = Self::get_attr(element.attributes(), b"svg:y2");
        } else {
            builder.x = Self::get_attr(element.attributes(), b"svg:x");
            builder.y = Self::get_attr(element.attributes(), b"svg:y");
            builder.width = Self::get_attr(element.attributes(), b"svg:width");
            builder.height = Self::get_attr(element.attributes(), b"svg:height");
        }
        builder.style_name = Self::get_attr(element.attributes(), b"draw:style-name")
            .or_else(|| Self::get_attr(element.attributes(), b"presentation:style-name"));
        builder
    }

    fn append_segment(target: &mut String, has_segment: &mut bool, text: &str) {
        if *has_segment {
            target.push('\n');
        }
        target.push_str(text);
        *has_segment = true;
    }

    fn finish_shape(
        builder: ShapeBuilder,
        slide_title: &mut Option<String>,
        slide_text: &mut String,
        slide_has_segment: &mut bool,
        shapes: &mut Vec<Shape>,
    ) {
        let is_title = builder.is_title;
        let shape = builder.build();
        if is_title {
            *slide_title = Some(shape.text);
        } else if matches!(
            shape.shape_type,
            ShapeType::TextBox | ShapeType::Placeholder
        ) && shape.has_text()
        {
            Self::append_segment(slide_text, slide_has_segment, &shape.text);
        } else {
            shapes.push(shape);
        }
    }

    fn decode_text(text: &quick_xml::events::BytesText<'_>) -> Result<String> {
        let decoded = text
            .xml_content(XmlVersion::Explicit1_0)
            .map_err(|error| Error::InvalidFormat(format!("invalid presentation text: {error}")))?;
        Ok(decoded.into_owned())
    }

    fn decode_reference(reference: &BytesRef<'_>) -> Result<String> {
        if let Some(character) = reference.resolve_char_ref().map_err(|error| {
            Error::InvalidFormat(format!("invalid presentation character reference: {error}"))
        })? {
            return Ok(character.to_string());
        }
        let name = reference.decode().map_err(|error| {
            Error::InvalidFormat(format!("invalid presentation entity reference: {error}"))
        })?;
        match name.as_ref() {
            "amp" => Ok("&".to_string()),
            "lt" => Ok("<".to_string()),
            "gt" => Ok(">".to_string()),
            "quot" => Ok("\"".to_string()),
            "apos" => Ok("'".to_string()),
            _ => Err(Error::InvalidFormat(format!(
                "unsupported presentation entity reference '&{name};'"
            ))),
        }
    }

    fn push_parsed_paragraph(
        text: &str,
        in_notes: bool,
        notes: &mut String,
        notes_has_paragraph: &mut bool,
        shape: Option<&mut ShapeBuilder>,
        slide_text: &mut String,
        slide_has_segment: &mut bool,
    ) {
        if in_notes {
            Self::append_segment(notes, notes_has_paragraph, text);
        } else if let Some(shape) = shape {
            shape.push_paragraph(text);
        } else {
            Self::append_segment(slide_text, slide_has_segment, text);
        }
    }

    fn push_text_control(element: &BytesStart<'_>, paragraph: &mut ParagraphText) -> Result<()> {
        match element.name().as_ref() {
            b"text:line-break" => paragraph.push_explicit('\n', 1),
            b"text:tab" => paragraph.push_explicit('\t', 1),
            b"text:s" => {
                let count = Self::get_attr(element.attributes(), b"text:c")
                    .map(|value| {
                        value.parse::<usize>().map_err(|_| {
                            Error::InvalidFormat(format!("invalid text:s count '{value}'"))
                        })
                    })
                    .transpose()?
                    .unwrap_or(1);
                if count > 1_000_000 {
                    return Err(Error::InvalidFormat(
                        "text:s count exceeds the supported safety limit".to_string(),
                    ));
                }
                paragraph.push_explicit(' ', count);
            },
            _ => {},
        }
        Ok(())
    }

    /// Parse all slides from ODP content.xml
    pub fn parse_slides(xml_content: &str) -> Result<Vec<Slide>> {
        let mut reader = Reader::from_str(xml_content);
        let mut buf = Vec::new();
        let mut slides = Vec::new();

        // State tracking
        let mut current_slide_text = String::new();
        let mut current_slide_title: Option<String> = None;
        let mut current_shapes: Vec<Shape> = Vec::new();
        let mut in_slide = false;
        let mut slide_index = 0;
        let mut current_notes_text = String::new();
        let mut current_notes_has_paragraph = false;
        let mut in_notes = false;
        let mut current_slide_has_segment = false;

        // Shape parsing state
        let mut current_shape: Option<ShapeBuilder> = None;
        let mut shape_depth = 0;
        let mut current_paragraph: Option<ParagraphText> = None;

        loop {
            match reader.read_event_into(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    match e.name().as_ref() {
                        b"draw:page" => {
                            // Finish previous slide if any
                            if in_slide {
                                slides.push(Slide {
                                    title: current_slide_title.take(),
                                    text: std::mem::take(&mut current_slide_text),
                                    index: slide_index,
                                    notes: (!current_notes_text.is_empty())
                                        .then(|| std::mem::take(&mut current_notes_text)),
                                    shapes: std::mem::take(&mut current_shapes),
                                });
                                slide_index += 1;
                            }

                            // Start new slide
                            current_slide_title = None;
                            current_slide_has_segment = false;
                            current_notes_has_paragraph = false;
                            in_slide = true;
                        },
                        b"presentation:notes" if in_slide => {
                            in_notes = true;
                        },
                        b"text:p" | b"text:h" if in_slide => {
                            if current_paragraph.is_some() {
                                return Err(Error::InvalidFormat(
                                    "nested ODP text paragraphs are not supported".to_string(),
                                ));
                            }
                            current_paragraph = Some(ParagraphText::default());
                        },
                        b"text:s" | b"text:tab" | b"text:line-break"
                            if current_paragraph.is_some() =>
                        {
                            Self::push_text_control(
                                e,
                                current_paragraph.as_mut().expect("paragraph checked above"),
                            )?;
                        },
                        _ if in_notes => {},
                        name if Self::is_shape_element(name) => {
                            if in_slide && current_shape.is_none() {
                                current_shape = Some(Self::shape_builder(e));
                                shape_depth = 0;
                            } else if current_shape.is_some() {
                                shape_depth += 1;
                            }
                        },
                        b"draw:image" if current_shape.is_some() => {
                            let builder = current_shape.as_mut().expect("shape checked above");
                            builder.shape_type = ShapeType::Picture;
                            builder.image_href = Self::get_attr(e.attributes(), b"xlink:href");
                        },
                        b"table:table" if current_shape.is_some() => {
                            current_shape
                                .as_mut()
                                .expect("shape checked above")
                                .shape_type = ShapeType::Table;
                        },
                        b"draw:object" | b"draw:object-ole" | b"draw:plugin"
                            if current_shape.is_some() =>
                        {
                            current_shape
                                .as_mut()
                                .expect("shape checked above")
                                .shape_type = ShapeType::GraphicFrame;
                        },
                        _ => {},
                    }
                },
                Ok(Event::Text(ref t)) if current_paragraph.is_some() => {
                    let text = Self::decode_text(t)?;
                    current_paragraph
                        .as_mut()
                        .expect("paragraph checked above")
                        .push_text(&text);
                },
                Ok(Event::CData(ref text)) if current_paragraph.is_some() => {
                    let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid presentation CDATA: {error}"))
                    })?;
                    current_paragraph
                        .as_mut()
                        .expect("paragraph checked above")
                        .push_text(&decoded);
                },
                Ok(Event::GeneralRef(ref reference)) if current_paragraph.is_some() => {
                    let text = Self::decode_reference(reference)?;
                    current_paragraph
                        .as_mut()
                        .expect("paragraph checked above")
                        .push_text(&text);
                },
                Ok(Event::Empty(ref e))
                    if in_slide && matches!(e.name().as_ref(), b"text:p" | b"text:h") =>
                {
                    Self::push_parsed_paragraph(
                        "",
                        in_notes,
                        &mut current_notes_text,
                        &mut current_notes_has_paragraph,
                        current_shape.as_mut(),
                        &mut current_slide_text,
                        &mut current_slide_has_segment,
                    );
                },
                Ok(Event::Empty(ref e))
                    if current_paragraph.is_some()
                        && matches!(
                            e.name().as_ref(),
                            b"text:s" | b"text:tab" | b"text:line-break"
                        ) =>
                {
                    Self::push_text_control(
                        e,
                        current_paragraph.as_mut().expect("paragraph checked above"),
                    )?;
                },
                Ok(Event::Empty(ref e)) if !in_notes => match e.name().as_ref() {
                    b"draw:image" => {
                        if let Some(builder) = current_shape.as_mut() {
                            builder.shape_type = ShapeType::Picture;
                            builder.image_href = Self::get_attr(e.attributes(), b"xlink:href");
                        }
                    },
                    b"table:table" => {
                        if let Some(builder) = current_shape.as_mut() {
                            builder.shape_type = ShapeType::Table;
                        }
                    },
                    b"draw:object" | b"draw:object-ole" | b"draw:plugin" => {
                        if let Some(builder) = current_shape.as_mut() {
                            builder.shape_type = ShapeType::GraphicFrame;
                        }
                    },
                    name if in_slide && current_shape.is_none() && Self::is_shape_element(name) => {
                        Self::finish_shape(
                            Self::shape_builder(e),
                            &mut current_slide_title,
                            &mut current_slide_text,
                            &mut current_slide_has_segment,
                            &mut current_shapes,
                        );
                    },
                    _ => {},
                },
                Ok(Event::End(ref e)) => {
                    if matches!(e.name().as_ref(), b"text:p" | b"text:h")
                        && current_paragraph.is_some()
                    {
                        let paragraph = current_paragraph
                            .take()
                            .expect("paragraph checked above")
                            .finish();
                        Self::push_parsed_paragraph(
                            &paragraph,
                            in_notes,
                            &mut current_notes_text,
                            &mut current_notes_has_paragraph,
                            current_shape.as_mut(),
                            &mut current_slide_text,
                            &mut current_slide_has_segment,
                        );
                        buf.clear();
                        continue;
                    }
                    if e.name().as_ref() == b"presentation:notes" {
                        in_notes = false;
                        buf.clear();
                        continue;
                    }
                    if in_notes {
                        buf.clear();
                        continue;
                    }
                    match e.name().as_ref() {
                        b"draw:page" => {
                            // Finish current slide
                            if in_slide {
                                slides.push(Slide {
                                    title: current_slide_title.take(),
                                    text: std::mem::take(&mut current_slide_text),
                                    index: slide_index,
                                    notes: (!current_notes_text.is_empty())
                                        .then(|| std::mem::take(&mut current_notes_text)),
                                    shapes: std::mem::take(&mut current_shapes),
                                });
                                slide_index += 1;
                            }
                            current_slide_has_segment = false;
                            current_notes_has_paragraph = false;
                            in_slide = false;
                        },
                        b"draw:frame" | b"draw:rect" | b"draw:ellipse" | b"draw:line"
                        | b"draw:custom-shape" | b"draw:circle" | b"draw:path"
                        | b"draw:polygon" | b"draw:polyline" | b"draw:connector" | b"draw:g" => {
                            if shape_depth > 0 {
                                shape_depth -= 1;
                            } else if let Some(builder) = current_shape.take() {
                                // Finish the shape and add it to the slide
                                Self::finish_shape(
                                    builder,
                                    &mut current_slide_title,
                                    &mut current_slide_text,
                                    &mut current_slide_has_segment,
                                    &mut current_shapes,
                                );
                            }
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
                return Self::normalize_attr(&attr);
            }
        }
        None
    }

    fn normalize_attr(attr: &Attribute<'_>) -> Option<String> {
        attr.normalized_value(XmlVersion::Implicit1_0)
            .ok()
            .map(|value| value.into_owned())
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
    xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
    xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0">
    <office:body>
        <office:presentation>
            <draw:page draw:name="Shapes">
                <draw:ellipse draw:name="Circle1" svg:x="1cm" svg:y="1cm" svg:width="3cm" svg:height="3cm">
                    <draw:text-box>
                        <text:p>Circle</text:p>
                    </draw:text-box>
                </draw:ellipse>
                <draw:line draw:name="Line1" svg:x1="0cm" svg:y1="0cm" svg:x2="10cm" svg:y2="10cm"/>
                <draw:connector draw:name="Connector1" svg:x1="1cm" svg:y1="2cm" svg:x2="3cm" svg:y2="4cm"/>
                <draw:custom-shape draw:name="Custom1" svg:x="5cm" svg:y="5cm"/>
                <presentation:notes><draw:frame><draw:text-box><text:p>Speaker note</text:p></draw:text-box></draw:frame></presentation:notes>
            </draw:page>
        </office:presentation>
    </office:body>
</office:document-content>"#;

    #[test]
    fn test_parse_slides() {
        let slides = OdpParser::parse_slides(TEST_PRESENTATION_XML).unwrap();
        assert_eq!(slides.len(), 2);

        // First slide
        assert_eq!(slides[0].title, Some("Welcome".to_string()));
        assert_eq!(slides[0].index, 0);
        assert!(slides[0].text.is_empty());
        assert_eq!(slides[0].shapes.len(), 1);
        assert_eq!(slides[0].shapes[0].text, "Rectangle content");
        assert_eq!(slides[0].all_text(), "Welcome\nRectangle content");

        // Second slide
        assert_eq!(slides[1].title, None);
        assert_eq!(slides[1].index, 1);
        assert_eq!(slides[1].text, "Bullet 1\nBullet 2");
        assert!(slides[1].shapes.is_empty());
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
        assert_eq!(slide.shapes.len(), 4);
        assert!(
            slide
                .shapes
                .iter()
                .any(|shape| shape.shape_type == ShapeType::Connector)
        );
        assert_eq!(slide.notes.as_deref(), Some("Speaker note"));
        assert!(!slide.all_text().contains("Speaker note"));
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
            image_href: None,
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
            image_href: None,
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
                image_href: None,
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
        let t2 = t1;
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

    #[test]
    fn parses_picture_shape_and_unescapes_href() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page draw:name="Images"><draw:frame draw:name="Picture"><draw:image xlink:href="Pictures/a&amp;b.png"/></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

        let slides = OdpParser::parse_slides(xml).unwrap();
        let shape = &slides[0].shapes[0];
        assert_eq!(shape.shape_type, ShapeType::Picture);
        assert_eq!(shape.image_href(), Some("Pictures/a&b.png"));
    }

    #[test]
    fn identifies_shapes_that_cannot_be_losslessly_rebuilt() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"><office:body><office:presentation><draw:page><draw:g draw:name="Group"><draw:rect/></draw:g><draw:frame draw:name="Table"><table:table/></draw:frame><draw:frame draw:name="Object"><draw:object/></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

        let slides = OdpParser::parse_slides(xml).unwrap();
        let types: Vec<_> = slides[0]
            .shapes
            .iter()
            .map(|shape| shape.shape_type)
            .collect();
        assert_eq!(
            types,
            [ShapeType::Group, ShapeType::Table, ShapeType::GraphicFrame]
        );
    }

    #[test]
    fn preserves_text_across_spans_and_odf_whitespace_elements() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:presentation><draw:page><draw:frame presentation:class="object"><draw:text-box><text:p><text:s/>Hel<text:span>lo</text:span> <text:span>world</text:span><text:s text:c="2"/>again<text:tab/>tab<text:line-break/>line &amp; more</text:p><text:p/><text:p>second paragraph<text:s/></text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

        let slides = OdpParser::parse_slides(xml).unwrap();
        assert_eq!(
            slides[0].text,
            " Hello world  again\ttab\nline & more\n\nsecond paragraph "
        );
    }

    #[test]
    fn rejects_excessive_explicit_space_expansion() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:presentation><draw:page><draw:frame><draw:text-box><text:p>x<text:s text:c="1000001"/></text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

        let error = OdpParser::parse_slides(xml).unwrap_err();
        assert!(error.to_string().contains("safety limit"));
    }
}
