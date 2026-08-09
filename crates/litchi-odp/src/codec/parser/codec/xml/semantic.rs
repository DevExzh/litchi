//! Semantic ODP model assembly from XML elements.

use super::super::{
    ANIMATION_NAMESPACE_BYTES, Actuate, AnimationKind, BytesStart, DR3D_NAMESPACE, DRAW_NAMESPACE,
    Direction, DrawingAttribute, DrawingAttributeNamespace, DrawingHyperlink, DrawingShapeKind,
    Element, Error, HashMap, HashSet, HyperlinkShow, Kind, NsReader, OFFICE_NAMESPACE,
    PRESENTATION_NAMESPACE, ParagraphText, Parameter, Parser, Reference, ResolveResult, Result,
    SCRIPT_NAMESPACE, SMIL_NAMESPACE, SVG_NAMESPACE, Shape, ShapeBuilder, ShapeElement, ShapeType,
    Show, Sound, SoundShow, Speed, Style, TABLE_NAMESPACE, TEXT_NAMESPACE, Transition,
    TransitionStyleDefinition, TransitionStyles, Type, XLINK_NAMESPACE, XML_NAMESPACE, XmlVersion,
};

impl Parser {
    pub(super) fn classify(namespace: &ResolveResult<'_>, local_name: &[u8]) -> Element {
        if Self::is_namespace(namespace, ANIMATION_NAMESPACE_BYTES) {
            Kind::from_local_name(local_name).map_or(Element::UnknownAnimation, Element::Animation)
        } else if Self::is_namespace(namespace, DRAW_NAMESPACE) {
            match local_name {
                b"page" => Element::Page,
                b"frame" => Element::Shape(ShapeElement::Frame),
                b"rect" => Element::Shape(ShapeElement::Rect),
                b"ellipse" => Element::Shape(ShapeElement::Ellipse),
                b"line" => Element::Shape(ShapeElement::Line),
                b"custom-shape" => Element::Shape(ShapeElement::CustomShape),
                b"circle" => Element::Shape(ShapeElement::Circle),
                b"path" => Element::Shape(ShapeElement::Path),
                b"polygon" => Element::Shape(ShapeElement::Polygon),
                b"polyline" => Element::Shape(ShapeElement::Polyline),
                b"regular-polygon" => Element::Shape(ShapeElement::RegularPolygon),
                b"page-thumbnail" => Element::Shape(ShapeElement::PageThumbnail),
                b"measure" => Element::Shape(ShapeElement::Measure),
                b"caption" => Element::Shape(ShapeElement::Caption),
                b"connector" => Element::Shape(ShapeElement::Connector),
                b"control" => Element::Shape(ShapeElement::Control),
                b"g" => Element::Shape(ShapeElement::Group),
                b"image" => Element::Image,
                b"object" | b"object-ole" => Element::Object,
                b"plugin" => Element::Plugin,
                b"param" => Element::PluginParameter,
                b"a" => Element::DrawingHyperlink,
                b"enhanced-geometry" => Element::EnhancedGeometry,
                b"equation" => Element::EnhancedEquation,
                b"handle" => Element::EnhancedHandle,
                _ => Element::Other,
            }
        } else if Self::is_namespace(namespace, DR3D_NAMESPACE) {
            match local_name {
                b"scene" => Element::Shape(ShapeElement::ThreeDimensionalScene),
                b"light" => Element::Shape(ShapeElement::ThreeDimensionalLight),
                b"cube" => Element::Shape(ShapeElement::ThreeDimensionalCube),
                b"sphere" => Element::Shape(ShapeElement::ThreeDimensionalSphere),
                b"extrude" => Element::Shape(ShapeElement::ThreeDimensionalExtrude),
                b"rotate" => Element::Shape(ShapeElement::ThreeDimensionalRotate),
                _ => Element::Other,
            }
        } else if Self::is_namespace(namespace, OFFICE_NAMESPACE) {
            match local_name {
                b"event-listeners" => Element::EventListeners,
                b"spreadsheet" => Element::SpreadsheetRoot,
                _ => Element::Other,
            }
        } else if Self::is_namespace(namespace, PRESENTATION_NAMESPACE) {
            if local_name == b"notes" {
                Element::Notes
            } else if local_name == b"event-listener" {
                Element::EventListener
            } else if local_name == b"sound" {
                Element::Sound
            } else {
                AnimationKind::from_local_name(local_name)
                    .map_or(Element::Other, Element::LegacyAnimation)
            }
        } else if Self::is_namespace(namespace, SCRIPT_NAMESPACE) && local_name == b"event-listener"
        {
            Element::ScriptEventListener
        } else if Self::is_namespace(namespace, TABLE_NAMESPACE) {
            match local_name {
                b"table" => Element::Table,
                b"shapes" => Element::SheetShapes,
                _ => Element::Other,
            }
        } else if Self::is_namespace(namespace, TEXT_NAMESPACE) {
            match local_name {
                b"p" | b"h" => Element::TextParagraph,
                b"s" => Element::TextSpace,
                b"tab" => Element::TextTab,
                b"line-break" => Element::TextLineBreak,
                _ => Element::Other,
            }
        } else {
            Element::Other
        }
    }

    pub(super) fn drawing_kind(shape_element: ShapeElement) -> DrawingShapeKind {
        match shape_element {
            ShapeElement::Frame => DrawingShapeKind::Frame,
            ShapeElement::Rect => DrawingShapeKind::Rectangle,
            ShapeElement::Ellipse => DrawingShapeKind::Ellipse,
            ShapeElement::Line => DrawingShapeKind::Line,
            ShapeElement::CustomShape => DrawingShapeKind::CustomShape,
            ShapeElement::Circle => DrawingShapeKind::Circle,
            ShapeElement::Path => DrawingShapeKind::Path,
            ShapeElement::Polygon => DrawingShapeKind::Polygon,
            ShapeElement::Polyline => DrawingShapeKind::Polyline,
            ShapeElement::RegularPolygon => DrawingShapeKind::RegularPolygon,
            ShapeElement::PageThumbnail => DrawingShapeKind::PageThumbnail,
            ShapeElement::Measure => DrawingShapeKind::Measure,
            ShapeElement::Caption => DrawingShapeKind::Caption,
            ShapeElement::Connector => DrawingShapeKind::Connector,
            ShapeElement::Control => DrawingShapeKind::Control,
            ShapeElement::Group => DrawingShapeKind::Group,
            ShapeElement::ThreeDimensionalScene => DrawingShapeKind::ThreeDimensionalScene,
            ShapeElement::ThreeDimensionalLight => DrawingShapeKind::ThreeDimensionalLight,
            ShapeElement::ThreeDimensionalCube => DrawingShapeKind::ThreeDimensionalCube,
            ShapeElement::ThreeDimensionalSphere => DrawingShapeKind::ThreeDimensionalSphere,
            ShapeElement::ThreeDimensionalExtrude => DrawingShapeKind::ThreeDimensionalExtrude,
            ShapeElement::ThreeDimensionalRotate => DrawingShapeKind::ThreeDimensionalRotate,
        }
    }

    pub(super) fn shape_builder(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        shape_element: ShapeElement,
    ) -> Result<ShapeBuilder> {
        let mut builder = ShapeBuilder::new();
        let presentation_class = Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"class")?;
        let drawing_kind = Self::drawing_kind(shape_element);
        builder.is_frame = matches!(shape_element, ShapeElement::Frame);
        builder.drawing_kind = Some(drawing_kind);
        builder.is_title = presentation_class.as_deref() == Some("title");
        builder.shape_type = match shape_element {
            ShapeElement::Frame => match presentation_class.as_deref() {
                Some(_) => ShapeType::Placeholder,
                _ => ShapeType::TextBox,
            },
            ShapeElement::Line | ShapeElement::Measure => ShapeType::Line,
            ShapeElement::Connector => ShapeType::Connector,
            ShapeElement::Group | ShapeElement::ThreeDimensionalScene => ShapeType::Group,
            ShapeElement::Rect
            | ShapeElement::Ellipse
            | ShapeElement::CustomShape
            | ShapeElement::Circle
            | ShapeElement::Path
            | ShapeElement::Polygon
            | ShapeElement::Polyline
            | ShapeElement::RegularPolygon
            | ShapeElement::PageThumbnail
            | ShapeElement::Caption
            | ShapeElement::Control
            | ShapeElement::ThreeDimensionalLight
            | ShapeElement::ThreeDimensionalCube
            | ShapeElement::ThreeDimensionalSphere
            | ShapeElement::ThreeDimensionalExtrude
            | ShapeElement::ThreeDimensionalRotate => ShapeType::AutoShape,
        };
        builder.name = Self::get_attr(reader, element, DRAW_NAMESPACE, b"name")?;
        if matches!(
            shape_element,
            ShapeElement::Line | ShapeElement::Connector | ShapeElement::Measure
        ) {
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
        builder.layer = Self::get_attr(reader, element, DRAW_NAMESPACE, b"layer")?;
        builder.z_index = Self::get_attr(reader, element, DRAW_NAMESPACE, b"z-index")?;
        if let Some(z_index) = &builder.z_index {
            crate::model::slide::validate_z_index(z_index)?;
        }
        builder.transform = Self::get_attr(reader, element, DRAW_NAMESPACE, b"transform")?;
        builder.presentation_class = presentation_class;
        builder.presentation_placeholder = Self::parse_optional_bool(
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"placeholder")?,
            "presentation:placeholder",
        )?;
        builder.presentation_user_transformed = Self::parse_optional_bool(
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"user-transformed")?,
            "presentation:user-transformed",
        )?;
        builder.drawing_attributes = Self::drawing_attributes(reader, element)?;
        Self::validate_required_three_dimensional_attributes(
            drawing_kind,
            &builder.drawing_attributes,
        )?;
        Ok(builder)
    }

    pub(super) fn drawing_attributes(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Vec<DrawingAttribute>> {
        let mut attributes = Vec::new();
        for attribute_result in element.attributes() {
            let attribute = attribute_result.map_err(|error| {
                Error::InvalidFormat(format!("invalid ODP shape attribute: {error}"))
            })?;
            let (namespace_uri, local_name_ref) =
                reader.resolver().resolve_attribute(attribute.key);
            let namespace = if Self::is_namespace(&namespace_uri, DRAW_NAMESPACE) {
                DrawingAttributeNamespace::Drawing
            } else if Self::is_namespace(&namespace_uri, SVG_NAMESPACE) {
                DrawingAttributeNamespace::Svg
            } else if Self::is_namespace(&namespace_uri, DR3D_NAMESPACE) {
                DrawingAttributeNamespace::Dr3d
            } else if Self::is_namespace(&namespace_uri, TABLE_NAMESPACE) {
                DrawingAttributeNamespace::Table
            } else {
                continue;
            };
            let local_name = local_name_ref.as_ref();
            let modeled = match namespace {
                DrawingAttributeNamespace::Drawing => matches!(
                    local_name,
                    b"name" | b"style-name" | b"layer" | b"z-index" | b"transform"
                ),
                DrawingAttributeNamespace::Svg => matches!(
                    local_name,
                    b"x" | b"y" | b"width" | b"height" | b"x1" | b"y1" | b"x2" | b"y2"
                ),
                DrawingAttributeNamespace::Dr3d | DrawingAttributeNamespace::Table => false,
            };
            if modeled {
                continue;
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid ODP shape attribute value: {error}"))
                })?
                .into_owned();
            attributes.push(DrawingAttribute::new(
                namespace,
                String::from_utf8(local_name.to_vec()).map_err(|_err| {
                    Error::InvalidFormat("non-UTF-8 ODP shape attribute name".to_string())
                })?,
                value,
            )?);
        }
        Ok(attributes)
    }

    pub(super) fn media_reference(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Reference> {
        let href = Self::get_attr(reader, element, XLINK_NAMESPACE, b"href")?.ok_or_else(|| {
            Error::InvalidFormat("draw:plugin is missing required xlink:href".to_string())
        })?;
        let link_type =
            Self::get_attr(reader, element, XLINK_NAMESPACE, b"type")?.ok_or_else(|| {
                Error::InvalidFormat("draw:plugin is missing required xlink:type".to_string())
            })?;
        if link_type != "simple" {
            return Err(Error::InvalidFormat(format!(
                "draw:plugin xlink:type must be 'simple', found '{link_type}'"
            )));
        }
        let mut media = Reference::new(href)?;
        if let Some(mime_type) = Self::get_attr(reader, element, DRAW_NAMESPACE, b"mime-type")? {
            media.set_mime_type(mime_type)?;
        }
        if let Some(show) = Self::get_attr(reader, element, XLINK_NAMESPACE, b"show")? {
            media.set_show(Some(Show::parse(&show)?));
        }
        if let Some(actuate) = Self::get_attr(reader, element, XLINK_NAMESPACE, b"actuate")? {
            media.set_actuate(Some(Actuate::parse(&actuate)?));
        }
        if let Some(xml_id) = Self::get_attr(reader, element, XML_NAMESPACE, b"id")? {
            media.set_xml_id(xml_id)?;
        }
        Ok(media)
    }

    pub(super) fn media_parameter(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Parameter> {
        let name = Self::get_attr(reader, element, DRAW_NAMESPACE, b"name")?.ok_or_else(|| {
            Error::InvalidFormat("draw:param is missing required draw:name".to_string())
        })?;
        let value =
            Self::get_attr(reader, element, DRAW_NAMESPACE, b"value")?.ok_or_else(|| {
                Error::InvalidFormat("draw:param is missing required draw:value".to_string())
            })?;
        Parameter::new(name, value)
    }

    pub(super) fn parse_on_request(value: Option<&str>, description: &str) -> Result<bool> {
        match value {
            None => Ok(false),
            Some("onRequest") => Ok(true),
            Some(actuate) => Err(Error::InvalidFormat(format!(
                "invalid {description} xlink:actuate '{actuate}'"
            ))),
        }
    }

    pub(super) fn drawing_hyperlink(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<DrawingHyperlink> {
        Self::require_simple_xlink(reader, element, "draw:a")?;
        let href = Self::required_attr(reader, element, XLINK_NAMESPACE, b"href", "xlink:href")?;
        let mut hyperlink = DrawingHyperlink::new(href)?;
        hyperlink.set_actuate_on_request(Self::parse_on_request(
            Self::get_attr(reader, element, XLINK_NAMESPACE, b"actuate")?.as_deref(),
            "draw:a",
        )?);
        hyperlink.set_show(
            Self::get_attr(reader, element, XLINK_NAMESPACE, b"show")?
                .map(|value| HyperlinkShow::parse(&value))
                .transpose()?,
        );
        hyperlink.set_target_frame_name(Self::get_attr(
            reader,
            element,
            OFFICE_NAMESPACE,
            b"target-frame-name",
        )?)?;
        hyperlink.set_name(Self::get_attr(reader, element, OFFICE_NAMESPACE, b"name")?)?;
        hyperlink.set_title(Self::get_attr(reader, element, OFFICE_NAMESPACE, b"title")?)?;
        hyperlink.set_server_map(Self::parse_optional_bool(
            Self::get_attr(reader, element, OFFICE_NAMESPACE, b"server-map")?,
            "office:server-map",
        )?);
        hyperlink.set_xml_id(Self::get_attr(reader, element, XML_NAMESPACE, b"id")?)?;
        Ok(hyperlink)
    }

    pub(super) fn append_segment(target: &mut String, has_segment: &mut bool, text: &str) {
        if *has_segment {
            target.push('\n');
        }
        target.push_str(text);
        *has_segment = true;
    }

    pub(super) fn finish_shape(
        builder: ShapeBuilder,
        slide_title: &mut Option<String>,
        slide_text: &mut String,
        slide_has_segment: &mut bool,
        shapes: &mut Vec<Shape>,
        retain_text_shapes: bool,
    ) {
        let is_title = builder.is_title;
        let shape = builder.build();
        if retain_text_shapes {
            shapes.push(shape);
        } else if is_title {
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

    pub(super) fn push_parsed_paragraph(
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
        } else if let Some(builder) = shape {
            builder.push_paragraph(text);
        } else {
            Self::append_segment(slide_text, slide_has_segment, text);
        }
    }

    pub(super) fn push_text_control(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        element_type: Element,
        paragraph: &mut ParagraphText,
    ) -> Result<()> {
        match element_type {
            Element::TextLineBreak => paragraph.push_explicit('\n', 1),
            Element::TextTab => paragraph.push_explicit('\t', 1),
            Element::TextSpace => {
                let count = Self::get_attr(reader, element, TEXT_NAMESPACE, b"c")?
                    .map(|value| {
                        value.parse::<usize>().map_err(|_err| {
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
            Element::Page
            | Element::Notes
            | Element::SheetShapes
            | Element::SpreadsheetRoot
            | Element::Shape(_)
            | Element::Image
            | Element::Table
            | Element::Object
            | Element::Plugin
            | Element::PluginParameter
            | Element::DrawingHyperlink
            | Element::EnhancedGeometry
            | Element::EnhancedEquation
            | Element::EnhancedHandle
            | Element::EventListeners
            | Element::EventListener
            | Element::ScriptEventListener
            | Element::Sound
            | Element::TextParagraph
            | Element::Animation(_)
            | Element::UnknownAnimation
            | Element::LegacyAnimation(_)
            | Element::Other => {},
        }
        Ok(())
    }

    pub(super) fn parse_transition_properties(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        transition: &mut Transition,
    ) -> Result<()> {
        transition.transition_type =
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"transition-type")?
                .map(|value| Type::parse(&value))
                .transpose()?;
        transition.style =
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"transition-style")?
                .map(Style::new)
                .transpose()?;
        transition.speed =
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"transition-speed")?
                .map(|value| Speed::parse(&value))
                .transpose()?;
        transition.smil_type = Self::get_attr(reader, element, SMIL_NAMESPACE, b"type")?;
        transition.smil_subtype = Self::get_attr(reader, element, SMIL_NAMESPACE, b"subtype")?;
        transition.direction = Self::get_attr(reader, element, SMIL_NAMESPACE, b"direction")?
            .map(|value| Direction::parse(&value))
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

    pub(super) fn resolved_transition_styles(
        content: &str,
        styles: Option<&str>,
    ) -> Result<(HashMap<String, Transition>, Transition)> {
        fn resolve(
            name: &str,
            definitions: &HashMap<String, TransitionStyleDefinition>,
            default: &Transition,
            cache: &mut HashMap<String, Transition>,
            visiting: &mut HashSet<String>,
            depth: usize,
        ) -> Result<Transition> {
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

        let mut definitions = TransitionStyles::default();
        if let Some(styles_source) = styles {
            definitions = Self::parse_transition_style_definitions(styles_source)?;
        }
        let content_definitions = Self::parse_transition_style_definitions(content)?;
        definitions.named.extend(content_definitions.named);
        if !content_definitions.default.is_empty() {
            definitions.default = content_definitions.default;
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

    pub(super) fn parse_transition_sound(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Sound> {
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
        if let Some(actuate_value) = actuate.as_deref()
            && actuate_value != "onRequest"
        {
            return Err(Error::InvalidFormat(format!(
                "invalid presentation:sound xlink:actuate '{actuate_value}'"
            )));
        }
        let show = Self::get_attr(reader, element, XLINK_NAMESPACE, b"show")?
            .map(|value| SoundShow::parse(&value))
            .transpose()?;
        let play_full = Self::parse_optional_bool(
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"play-full")?,
            "presentation:play-full",
        )?;
        Ok(Sound {
            href,
            play_full,
            actuate_on_request: actuate.is_some(),
            show,
            xml_id: Self::get_attr(reader, element, XML_NAMESPACE, b"id")?,
        })
    }
}
