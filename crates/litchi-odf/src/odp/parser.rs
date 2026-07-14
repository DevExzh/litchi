//! ODP-specific parsing utilities.

use super::{
    Shape, Slide, SlideTransition, TransitionDirection, TransitionSound, TransitionSoundShow,
    TransitionSpeed, TransitionStyle, TransitionType,
};
use litchi_core::{Error, Result, ShapeType};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const PRESENTATION_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const SMIL_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";
const SVG_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";

#[derive(Clone, Default)]
struct TransitionStyleDefinition {
    parent: Option<String>,
    transition: SlideTransition,
}

#[derive(Default)]
struct TransitionStyles {
    named: HashMap<String, TransitionStyleDefinition>,
    default: SlideTransition,
}

#[derive(Clone, Copy)]
enum ShapeElement {
    Frame,
    Rect,
    Ellipse,
    Line,
    CustomShape,
    Circle,
    Path,
    Polygon,
    Polyline,
    Connector,
    Group,
}

#[derive(Clone, Copy)]
enum OdpElement {
    Page,
    Notes,
    Shape(ShapeElement),
    Image,
    Table,
    Object,
    TextParagraph,
    TextSpace,
    TextTab,
    TextLineBreak,
    Other,
}

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
    fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
        matches!(namespace, ResolveResult::Bound(Namespace(value)) if *value == expected)
    }

    fn classify(namespace: &ResolveResult<'_>, local_name: &[u8]) -> OdpElement {
        if Self::is_namespace(namespace, DRAW_NAMESPACE) {
            match local_name {
                b"page" => OdpElement::Page,
                b"frame" => OdpElement::Shape(ShapeElement::Frame),
                b"rect" => OdpElement::Shape(ShapeElement::Rect),
                b"ellipse" => OdpElement::Shape(ShapeElement::Ellipse),
                b"line" => OdpElement::Shape(ShapeElement::Line),
                b"custom-shape" => OdpElement::Shape(ShapeElement::CustomShape),
                b"circle" => OdpElement::Shape(ShapeElement::Circle),
                b"path" => OdpElement::Shape(ShapeElement::Path),
                b"polygon" => OdpElement::Shape(ShapeElement::Polygon),
                b"polyline" => OdpElement::Shape(ShapeElement::Polyline),
                b"connector" => OdpElement::Shape(ShapeElement::Connector),
                b"g" => OdpElement::Shape(ShapeElement::Group),
                b"image" => OdpElement::Image,
                b"object" | b"object-ole" | b"plugin" => OdpElement::Object,
                _ => OdpElement::Other,
            }
        } else if Self::is_namespace(namespace, PRESENTATION_NAMESPACE) && local_name == b"notes" {
            OdpElement::Notes
        } else if Self::is_namespace(namespace, TABLE_NAMESPACE) && local_name == b"table" {
            OdpElement::Table
        } else if Self::is_namespace(namespace, TEXT_NAMESPACE) {
            match local_name {
                b"p" | b"h" => OdpElement::TextParagraph,
                b"s" => OdpElement::TextSpace,
                b"tab" => OdpElement::TextTab,
                b"line-break" => OdpElement::TextLineBreak,
                _ => OdpElement::Other,
            }
        } else {
            OdpElement::Other
        }
    }

    fn shape_builder(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        shape_element: ShapeElement,
    ) -> Result<ShapeBuilder> {
        let mut builder = ShapeBuilder::new();
        let presentation_class = Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"class")?;
        builder.is_title = presentation_class.as_deref() == Some("title");
        builder.shape_type = match shape_element {
            ShapeElement::Frame => match presentation_class.as_deref() {
                Some(_) => ShapeType::Placeholder,
                _ => ShapeType::TextBox,
            },
            ShapeElement::Line => ShapeType::Line,
            ShapeElement::Connector => ShapeType::Connector,
            ShapeElement::Group => ShapeType::Group,
            _ => ShapeType::AutoShape,
        };
        builder.name = Self::get_attr(reader, element, DRAW_NAMESPACE, b"name")?;
        if matches!(shape_element, ShapeElement::Line | ShapeElement::Connector) {
            builder.x = Self::get_attr(reader, element, SVG_NAMESPACE, b"x1")?;
            builder.y = Self::get_attr(reader, element, SVG_NAMESPACE, b"y1")?;
            builder.width = Self::get_attr(reader, element, SVG_NAMESPACE, b"x2")?;
            builder.height = Self::get_attr(reader, element, SVG_NAMESPACE, b"y2")?;
        } else {
            builder.x = Self::get_attr(reader, element, SVG_NAMESPACE, b"x")?;
            builder.y = Self::get_attr(reader, element, SVG_NAMESPACE, b"y")?;
            builder.width = Self::get_attr(reader, element, SVG_NAMESPACE, b"width")?;
            builder.height = Self::get_attr(reader, element, SVG_NAMESPACE, b"height")?;
        }
        builder.style_name = Self::get_attr(reader, element, DRAW_NAMESPACE, b"style-name")?.or(
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"style-name")?,
        );
        Ok(builder)
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

    fn push_text_control(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        element_type: OdpElement,
        paragraph: &mut ParagraphText,
    ) -> Result<()> {
        match element_type {
            OdpElement::TextLineBreak => paragraph.push_explicit('\n', 1),
            OdpElement::TextTab => paragraph.push_explicit('\t', 1),
            OdpElement::TextSpace => {
                let count = Self::get_attr(reader, element, TEXT_NAMESPACE, b"c")?
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

    fn parse_optional_bool(value: Option<String>, attribute: &str) -> Result<Option<bool>> {
        value
            .map(|value| match value.as_str() {
                "true" | "1" => Ok(true),
                "false" | "0" => Ok(false),
                _ => Err(Error::InvalidFormat(format!(
                    "invalid {attribute} value '{value}'"
                ))),
            })
            .transpose()
    }

    fn parse_transition_properties(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        transition: &mut SlideTransition,
    ) -> Result<()> {
        transition.transition_type =
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"transition-type")?
                .map(|value| TransitionType::parse(&value))
                .transpose()?;
        transition.style =
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"transition-style")?
                .map(TransitionStyle::new)
                .transpose()?;
        transition.speed =
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"transition-speed")?
                .map(|value| TransitionSpeed::parse(&value))
                .transpose()?;
        transition.smil_type = Self::get_attr(reader, element, SMIL_NAMESPACE, b"type")?;
        transition.smil_subtype = Self::get_attr(reader, element, SMIL_NAMESPACE, b"subtype")?;
        transition.direction = Self::get_attr(reader, element, SMIL_NAMESPACE, b"direction")?
            .map(|value| TransitionDirection::parse(&value))
            .transpose()?;
        transition.set_fade_color(Self::get_attr(
            reader,
            element,
            SMIL_NAMESPACE,
            b"fadeColor",
        )?)?;
        transition.set_duration(Self::get_attr(
            reader,
            element,
            PRESENTATION_NAMESPACE,
            b"duration",
        )?)?;
        Ok(())
    }

    fn parse_transition_sound(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<TransitionSound> {
        let href = Self::get_attr(reader, element, XLINK_NAMESPACE, b"href")?.ok_or_else(|| {
            Error::InvalidFormat("presentation:sound is missing xlink:href".to_string())
        })?;
        if let Some(link_type) = Self::get_attr(reader, element, XLINK_NAMESPACE, b"type")?
            && link_type != "simple"
        {
            return Err(Error::InvalidFormat(format!(
                "invalid presentation:sound xlink:type '{link_type}'"
            )));
        }
        let actuate = Self::get_attr(reader, element, XLINK_NAMESPACE, b"actuate")?;
        if actuate.as_deref().is_some_and(|value| value != "onRequest") {
            return Err(Error::InvalidFormat(format!(
                "invalid presentation:sound xlink:actuate '{}'",
                actuate.as_deref().expect("actuate checked above")
            )));
        }
        let show = Self::get_attr(reader, element, XLINK_NAMESPACE, b"show")?
            .map(|value| TransitionSoundShow::parse(&value))
            .transpose()?;
        let play_full = Self::parse_optional_bool(
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"play-full")?,
            "presentation:play-full",
        )?;
        Ok(TransitionSound {
            href,
            play_full,
            actuate_on_request: actuate.is_some(),
            show,
            xml_id: Self::get_attr(reader, element, XML_NAMESPACE, b"id")?,
        })
    }

    fn parse_transition_style_definitions(xml: &str) -> Result<TransitionStyles> {
        let mut reader = NsReader::from_str(xml);
        let mut buf = Vec::new();
        let mut result = TransitionStyles::default();
        let mut current: Option<(Option<String>, bool, TransitionStyleDefinition)> = None;
        let mut in_properties = false;

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buf)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref element)
                    if Self::is_namespace(&namespace, STYLE_NAMESPACE)
                        && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
                {
                    let family = Self::get_attr(&reader, element, STYLE_NAMESPACE, b"family")?;
                    let is_drawing_page = family.as_deref() == Some("drawing-page");
                    let name = Self::get_attr(&reader, element, STYLE_NAMESPACE, b"name")?;
                    let parent =
                        Self::get_attr(&reader, element, STYLE_NAMESPACE, b"parent-style-name")?;
                    current = Some((
                        name,
                        is_drawing_page,
                        TransitionStyleDefinition {
                            parent,
                            transition: SlideTransition::new(),
                        },
                    ));
                },
                Event::Empty(ref element)
                    if Self::is_namespace(&namespace, STYLE_NAMESPACE)
                        && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
                {
                    let family = Self::get_attr(&reader, element, STYLE_NAMESPACE, b"family")?;
                    if family.as_deref() == Some("drawing-page") {
                        let name = Self::get_attr(&reader, element, STYLE_NAMESPACE, b"name")?;
                        let definition = TransitionStyleDefinition {
                            parent: Self::get_attr(
                                &reader,
                                element,
                                STYLE_NAMESPACE,
                                b"parent-style-name",
                            )?,
                            transition: SlideTransition::new(),
                        };
                        if let Some(name) = name {
                            result.named.insert(name, definition);
                        } else {
                            result.default = definition.transition;
                        }
                    }
                },
                Event::Start(ref element) | Event::Empty(ref element)
                    if current.as_ref().is_some_and(|(_, family, _)| *family)
                        && Self::is_namespace(&namespace, STYLE_NAMESPACE)
                        && element.local_name().as_ref() == b"drawing-page-properties" =>
                {
                    let (_, _, definition) = current.as_mut().expect("style checked above");
                    Self::parse_transition_properties(
                        &reader,
                        element,
                        &mut definition.transition,
                    )?;
                    in_properties = matches!(event, Event::Start(_));
                },
                Event::Start(ref element) | Event::Empty(ref element)
                    if in_properties
                        && Self::is_namespace(&namespace, PRESENTATION_NAMESPACE)
                        && element.local_name().as_ref() == b"sound" =>
                {
                    let (_, _, definition) = current.as_mut().expect("properties require style");
                    definition.transition.sound =
                        Some(Self::parse_transition_sound(&reader, element)?);
                },
                Event::End(ref element)
                    if Self::is_namespace(&namespace, STYLE_NAMESPACE)
                        && element.local_name().as_ref() == b"drawing-page-properties" =>
                {
                    in_properties = false;
                },
                Event::End(ref element)
                    if Self::is_namespace(&namespace, STYLE_NAMESPACE)
                        && matches!(element.local_name().as_ref(), b"style" | b"default-style") =>
                {
                    if let Some((name, is_drawing_page, definition)) = current.take()
                        && is_drawing_page
                    {
                        if let Some(name) = name {
                            result.named.insert(name, definition);
                        } else {
                            result.default = definition.transition;
                        }
                    }
                    in_properties = false;
                },
                Event::Eof => break,
                _ => {},
            }
            buf.clear();
        }
        Ok(result)
    }

    fn resolved_transition_styles(
        content: &str,
        styles: Option<&str>,
    ) -> Result<(HashMap<String, SlideTransition>, SlideTransition)> {
        let mut definitions = TransitionStyles::default();
        if let Some(styles) = styles {
            definitions = Self::parse_transition_style_definitions(styles)?;
        }
        let content_definitions = Self::parse_transition_style_definitions(content)?;
        definitions.named.extend(content_definitions.named);
        if !content_definitions.default.is_empty() {
            definitions.default = content_definitions.default;
        }

        fn resolve(
            name: &str,
            definitions: &HashMap<String, TransitionStyleDefinition>,
            default: &SlideTransition,
            cache: &mut HashMap<String, SlideTransition>,
            visiting: &mut HashSet<String>,
            depth: usize,
        ) -> Result<SlideTransition> {
            if let Some(value) = cache.get(name) {
                return Ok(value.clone());
            }
            if depth > 128 || !visiting.insert(name.to_string()) {
                return Err(Error::InvalidFormat(format!(
                    "cyclic or excessively deep drawing-page style inheritance at '{name}'"
                )));
            }
            let definition = definitions.get(name).cloned().unwrap_or_default();
            let mut value = definition.transition;
            let parent = if let Some(parent) = definition.parent {
                resolve(&parent, definitions, default, cache, visiting, depth + 1)?
            } else {
                default.clone()
            };
            value.inherit_from(&parent);
            visiting.remove(name);
            cache.insert(name.to_string(), value.clone());
            Ok(value)
        }

        let mut resolved = HashMap::with_capacity(definitions.named.len());
        let names: Vec<String> = definitions.named.keys().cloned().collect();
        for name in names {
            resolve(
                &name,
                &definitions.named,
                &definitions.default,
                &mut resolved,
                &mut HashSet::new(),
                0,
            )?;
        }
        Ok((resolved, definitions.default))
    }

    /// Parse all slides from ODP content.xml
    #[cfg(test)]
    pub fn parse_slides(xml_content: &str) -> Result<Vec<Slide>> {
        Self::parse_slides_with_styles(xml_content, None)
    }

    /// Parse slides and resolve drawing-page transition styles.
    pub fn parse_slides_with_styles(
        xml_content: &str,
        styles_xml: Option<&str>,
    ) -> Result<Vec<Slide>> {
        let (transition_styles, default_transition) =
            Self::resolved_transition_styles(xml_content, styles_xml)?;
        let mut reader = NsReader::from_str(xml_content);
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
        let mut current_transition: Option<SlideTransition> = None;

        // Shape parsing state
        let mut current_shape: Option<ShapeBuilder> = None;
        let mut shape_depth = 0;
        let mut current_paragraph: Option<ParagraphText> = None;

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buf)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref element) => {
                    let element_type = Self::classify(&namespace, element.local_name().as_ref());
                    match element_type {
                        OdpElement::Page => {
                            if in_slide {
                                slides.push(Slide {
                                    title: current_slide_title.take(),
                                    text: std::mem::take(&mut current_slide_text),
                                    index: slide_index,
                                    notes: (!current_notes_text.is_empty())
                                        .then(|| std::mem::take(&mut current_notes_text)),
                                    transition: current_transition.take(),
                                    shapes: std::mem::take(&mut current_shapes),
                                });
                                slide_index += 1;
                            }
                            current_slide_title = None;
                            current_slide_has_segment = false;
                            current_notes_has_paragraph = false;
                            let style_name =
                                Self::get_attr(&reader, element, DRAW_NAMESPACE, b"style-name")?;
                            let transition = style_name
                                .as_deref()
                                .and_then(|name| transition_styles.get(name))
                                .unwrap_or(&default_transition)
                                .clone();
                            current_transition = (!transition.is_empty()).then_some(transition);
                            in_slide = true;
                        },
                        OdpElement::Notes if in_slide => in_notes = true,
                        OdpElement::TextParagraph if in_slide => {
                            if current_paragraph.is_some() {
                                return Err(Error::InvalidFormat(
                                    "nested ODP text paragraphs are not supported".to_string(),
                                ));
                            }
                            current_paragraph = Some(ParagraphText::default());
                        },
                        OdpElement::TextSpace | OdpElement::TextTab | OdpElement::TextLineBreak
                            if current_paragraph.is_some() =>
                        {
                            Self::push_text_control(
                                &reader,
                                element,
                                element_type,
                                current_paragraph.as_mut().expect("paragraph checked above"),
                            )?;
                        },
                        _ if in_notes => {},
                        OdpElement::Shape(shape_element) => {
                            if in_slide && current_shape.is_none() {
                                current_shape =
                                    Some(Self::shape_builder(&reader, element, shape_element)?);
                                shape_depth = 0;
                            } else if current_shape.is_some() {
                                shape_depth += 1;
                            }
                        },
                        OdpElement::Image if current_shape.is_some() => {
                            let builder = current_shape.as_mut().expect("shape checked above");
                            builder.shape_type = ShapeType::Picture;
                            builder.image_href =
                                Self::get_attr(&reader, element, XLINK_NAMESPACE, b"href")?;
                        },
                        OdpElement::Table if current_shape.is_some() => {
                            current_shape
                                .as_mut()
                                .expect("shape checked above")
                                .shape_type = ShapeType::Table;
                        },
                        OdpElement::Object if current_shape.is_some() => {
                            current_shape
                                .as_mut()
                                .expect("shape checked above")
                                .shape_type = ShapeType::GraphicFrame;
                        },
                        _ => {},
                    }
                },
                Event::Text(ref text) if current_paragraph.is_some() => {
                    let text = Self::decode_text(text)?;
                    current_paragraph
                        .as_mut()
                        .expect("paragraph checked above")
                        .push_text(&text);
                },
                Event::CData(ref text) if current_paragraph.is_some() => {
                    let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid presentation CDATA: {error}"))
                    })?;
                    current_paragraph
                        .as_mut()
                        .expect("paragraph checked above")
                        .push_text(&decoded);
                },
                Event::GeneralRef(ref reference) if current_paragraph.is_some() => {
                    let text = Self::decode_reference(reference)?;
                    current_paragraph
                        .as_mut()
                        .expect("paragraph checked above")
                        .push_text(&text);
                },
                Event::Empty(ref element) => {
                    let element_type = Self::classify(&namespace, element.local_name().as_ref());
                    match element_type {
                        OdpElement::Page if !in_slide => {
                            let style_name =
                                Self::get_attr(&reader, element, DRAW_NAMESPACE, b"style-name")?;
                            let transition = style_name
                                .as_deref()
                                .and_then(|name| transition_styles.get(name))
                                .unwrap_or(&default_transition)
                                .clone();
                            slides.push(Slide {
                                title: None,
                                text: String::new(),
                                index: slide_index,
                                notes: None,
                                transition: (!transition.is_empty()).then_some(transition),
                                shapes: Vec::new(),
                            });
                            slide_index += 1;
                        },
                        OdpElement::TextParagraph if in_slide => {
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
                        OdpElement::TextSpace | OdpElement::TextTab | OdpElement::TextLineBreak
                            if current_paragraph.is_some() =>
                        {
                            Self::push_text_control(
                                &reader,
                                element,
                                element_type,
                                current_paragraph.as_mut().expect("paragraph checked above"),
                            )?;
                        },
                        _ if in_notes => {},
                        OdpElement::Image => {
                            if let Some(builder) = current_shape.as_mut() {
                                builder.shape_type = ShapeType::Picture;
                                builder.image_href =
                                    Self::get_attr(&reader, element, XLINK_NAMESPACE, b"href")?;
                            }
                        },
                        OdpElement::Table => {
                            if let Some(builder) = current_shape.as_mut() {
                                builder.shape_type = ShapeType::Table;
                            }
                        },
                        OdpElement::Object => {
                            if let Some(builder) = current_shape.as_mut() {
                                builder.shape_type = ShapeType::GraphicFrame;
                            }
                        },
                        OdpElement::Shape(shape_element) if in_slide && current_shape.is_none() => {
                            Self::finish_shape(
                                Self::shape_builder(&reader, element, shape_element)?,
                                &mut current_slide_title,
                                &mut current_slide_text,
                                &mut current_slide_has_segment,
                                &mut current_shapes,
                            );
                        },
                        _ => {},
                    }
                },
                Event::End(ref element) => {
                    let element_type = Self::classify(&namespace, element.local_name().as_ref());
                    if matches!(element_type, OdpElement::TextParagraph)
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
                    if matches!(element_type, OdpElement::Notes) {
                        in_notes = false;
                        buf.clear();
                        continue;
                    }
                    if in_notes {
                        buf.clear();
                        continue;
                    }
                    match element_type {
                        OdpElement::Page => {
                            if in_slide {
                                slides.push(Slide {
                                    title: current_slide_title.take(),
                                    text: std::mem::take(&mut current_slide_text),
                                    index: slide_index,
                                    notes: (!current_notes_text.is_empty())
                                        .then(|| std::mem::take(&mut current_notes_text)),
                                    transition: current_transition.take(),
                                    shapes: std::mem::take(&mut current_shapes),
                                });
                                slide_index += 1;
                            }
                            current_slide_has_segment = false;
                            current_notes_has_paragraph = false;
                            in_slide = false;
                        },
                        OdpElement::Shape(_) => {
                            if shape_depth > 0 {
                                shape_depth -= 1;
                            } else if let Some(builder) = current_shape.take() {
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
                Event::Eof => break,
                _ => {},
            }
            buf.clear();
        }

        Ok(slides)
    }

    /// Helper to extract attribute values
    fn get_attr(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        namespace_uri: &[u8],
        local_name: &[u8],
    ) -> Result<Option<String>> {
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
            if Self::is_namespace(&namespace, namespace_uri) && local.as_ref() == local_name {
                return attribute
                    .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                    .map(|value| Some(value.into_owned()))
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid XML attribute value: {error}"))
                    });
            }
        }
        Ok(None)
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
            transition: None,
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
            transition: None,
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

    #[test]
    fn parses_arbitrary_odf_namespace_prefixes() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:l="http://www.w3.org/1999/xlink" xmlns:tb="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:f="urn:example:wrong"><o:body><o:presentation><f:page><t:p>ignored</t:p></f:page><d:page><d:frame d:name="Aliased Title" p:class="title" s:x="1cm"><d:text-box><t:p>Aliased<t:s/>title</t:p></d:text-box></d:frame><d:frame d:name="Picture"><d:image l:href="Pictures/a&amp;b.png"/></d:frame><d:connector d:name="Link" s:x1="1cm" s:y1="2cm" s:x2="3cm" s:y2="4cm"/><d:frame d:name="Table"><tb:table/></d:frame><p:notes><d:frame><d:text-box><t:p>Aliased note</t:p></d:text-box></d:frame></p:notes></d:page></o:presentation></o:body></o:document-content>"#;

        let slides = OdpParser::parse_slides(xml).unwrap();
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].title.as_deref(), Some("Aliased title"));
        assert_eq!(slides[0].notes.as_deref(), Some("Aliased note"));
        let picture = &slides[0].shapes[0];
        assert_eq!(picture.name(), Some("Picture"));
        assert_eq!(picture.image_href(), Some("Pictures/a&b.png"));
        let connector = &slides[0].shapes[1];
        assert_eq!(connector.shape_type, ShapeType::Connector);
        assert_eq!(connector.position(), (Some("1cm"), Some("2cm")));
        assert_eq!(connector.dimensions(), (Some("3cm"), Some("4cm")));
        assert_eq!(slides[0].shapes[2].shape_type, ShapeType::Table);
    }

    #[test]
    fn resolves_transition_styles_across_package_parts_and_inheritance() {
        let styles = r##"<o:document-styles xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:l="http://www.w3.org/1999/xlink"><o:styles><s:default-style s:family="drawing-page"><s:drawing-page-properties p:transition-speed="slow"/></s:default-style><s:style s:name="Base" s:family="drawing-page"><s:drawing-page-properties p:transition-type="automatic" p:duration="PT8S"><p:sound l:type="simple" l:href="Sounds/a&amp;b.ogg" l:actuate="onRequest" l:show="replace" p:play-full="true"/></s:drawing-page-properties></s:style></o:styles></o:document-styles>"##;
        let content = r##"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0"><o:automatic-styles><s:style s:name="Child" s:family="drawing-page" s:parent-style-name="Base"><s:drawing-page-properties p:transition-style="fade-from-left" p:transition-speed="fast" m:type="fade" m:subtype="crossfade" m:direction="reverse" m:fadeColor="#aB09fF"/></s:style></o:automatic-styles><o:body><o:presentation><d:page d:style-name="Child"/></o:presentation></o:body></o:document-content>"##;

        let slides = OdpParser::parse_slides_with_styles(content, Some(styles)).unwrap();
        let transition = slides[0].transition().unwrap();
        assert_eq!(
            transition.transition_type(),
            Some(TransitionType::Automatic)
        );
        assert_eq!(transition.style().unwrap().as_str(), "fade-from-left");
        assert_eq!(transition.speed(), Some(TransitionSpeed::Fast));
        assert_eq!(transition.smil_type(), Some("fade"));
        assert_eq!(transition.smil_subtype(), Some("crossfade"));
        assert_eq!(transition.direction(), Some(TransitionDirection::Reverse));
        assert_eq!(transition.fade_color(), Some("#aB09fF"));
        assert_eq!(transition.duration(), Some("PT8S"));
        let sound = transition.sound().unwrap();
        assert_eq!(sound.href, "Sounds/a&b.ogg");
        assert_eq!(sound.play_full, Some(true));
        assert!(sound.actuate_on_request);
        assert_eq!(sound.show, Some(TransitionSoundShow::Replace));
    }

    #[test]
    fn rejects_cyclic_transition_style_inheritance() {
        let content = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:automatic-styles><s:style s:name="A" s:family="drawing-page" s:parent-style-name="B"/><s:style s:name="B" s:family="drawing-page" s:parent-style-name="A"/></o:automatic-styles><o:body><o:presentation><d:page d:style-name="A"/></o:presentation></o:body></o:document-content>"#;
        let error = OdpParser::parse_slides_with_styles(content, None).unwrap_err();
        assert!(error.to_string().contains("cyclic"));
    }
}
