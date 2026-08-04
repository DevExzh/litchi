//! ODP-specific parsing utilities.

use crate::model::animation::ANIMATION_NAMESPACE;
use crate::model::legacy_animation::validate_legacy_animation_root;
use crate::model::{
    Action, Actuate, Attribute, DrawingAttribute, DrawingAttributeNamespace, DrawingHyperlink,
    DrawingShapeKind, Effect, EffectDirection, EnhancedGeometry, EnhancedGeometryChild,
    EnhancedGeometryChildKind, EventListener, HyperlinkShow, Kind, LegacyAnimationKind,
    LegacyAnimationNode, Namespace, Node, Parameter, Reference, ScriptEventListener, Shape,
    ShapeEventListener, Show, Slide, Transition, TransitionDirection, TransitionSound,
    TransitionSoundShow, TransitionSpeed, TransitionStyle, TransitionType,
};
use litchi_core::{Error, Result, ShapeType};
use quick_xml::XmlVersion;
use quick_xml::events::{BytesRef, BytesStart, Event};
use quick_xml::name::{Namespace as XmlNamespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{HashMap, HashSet};

const DRAW_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:drawing:1.0";
const DR3D_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0";
const OFFICE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:office:1.0";
const PRESENTATION_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:presentation:1.0";
const SCRIPT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:script:1.0";
const STYLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:style:1.0";
const SMIL_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0";
const SVG_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0";
const TABLE_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:table:1.0";
const TEXT_NAMESPACE: &[u8] = b"urn:oasis:names:tc:opendocument:xmlns:text:1.0";
const XLINK_NAMESPACE: &[u8] = b"http://www.w3.org/1999/xlink";
const XML_NAMESPACE: &[u8] = b"http://www.w3.org/XML/1998/namespace";
const ANIMATION_NAMESPACE_BYTES: &[u8] = ANIMATION_NAMESPACE.as_bytes();

#[derive(Clone, Default)]
struct TransitionStyleDefinition {
    parent: Option<String>,
    transition: Transition,
}

#[derive(Default)]
struct TransitionStyles {
    named: HashMap<String, TransitionStyleDefinition>,
    default: Transition,
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
    RegularPolygon,
    PageThumbnail,
    Measure,
    Caption,
    Connector,
    Control,
    Group,
    ThreeDimensionalScene,
    ThreeDimensionalLight,
    ThreeDimensionalCube,
    ThreeDimensionalSphere,
    ThreeDimensionalExtrude,
    ThreeDimensionalRotate,
}

/// Container scope that supplies top-level drawing shapes.
#[derive(Clone, Copy, PartialEq, Eq)]
enum ShapeContainerScope {
    /// `draw:page` elements in presentations and drawings.
    DrawPages,
    /// `table:shapes` children of top-level spreadsheet tables.
    SpreadsheetTables,
}

#[derive(Clone, Copy)]
enum Element {
    Page,
    Notes,
    SheetShapes,
    SpreadsheetRoot,
    Shape(ShapeElement),
    Image,
    Table,
    Object,
    Plugin,
    PluginParameter,
    DrawingHyperlink,
    EnhancedGeometry,
    EnhancedEquation,
    EnhancedHandle,
    EventListeners,
    EventListener,
    ScriptEventListener,
    Sound,
    TextParagraph,
    TextSpace,
    TextTab,
    TextLineBreak,
    Animation(Kind),
    UnknownAnimation,
    LegacyAnimation(LegacyAnimationKind),
    Other,
}

/// Parser for ODP-specific structures.
///
/// This provides parsing logic specific to presentations,
/// including slide and shape parsing.
pub(crate) struct Parser;

/// Internal structure for building shapes during parsing
#[allow(dead_code)]
struct ShapeBuilder {
    shape_type: ShapeType,
    drawing_kind: Option<DrawingShapeKind>,
    drawing_attributes: Vec<DrawingAttribute>,
    children: Vec<Shape>,
    enhanced_geometry: Option<EnhancedGeometry>,
    text: String,
    name: Option<String>,
    x: Option<String>,
    y: Option<String>,
    width: Option<String>,
    height: Option<String>,
    style_name: Option<String>,
    layer: Option<String>,
    z_index: Option<String>,
    transform: Option<String>,
    presentation_class: Option<String>,
    presentation_placeholder: Option<bool>,
    presentation_user_transformed: Option<bool>,
    image_href: Option<String>,
    media: Option<Reference>,
    hyperlink: Option<DrawingHyperlink>,
    event_listeners: Vec<ShapeEventListener>,
    event_listeners_seen: bool,
    is_frame: bool,
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
            drawing_kind: None,
            drawing_attributes: Vec::new(),
            children: Vec::new(),
            enhanced_geometry: None,
            text: String::new(),
            name: None,
            x: None,
            y: None,
            width: None,
            height: None,
            style_name: None,
            layer: None,
            z_index: None,
            transform: None,
            presentation_class: None,
            presentation_placeholder: None,
            presentation_user_transformed: None,
            image_href: None,
            media: None,
            hyperlink: None,
            event_listeners: Vec::new(),
            event_listeners_seen: false,
            is_frame: false,
            is_title: false,
            has_paragraph: false,
        }
    }

    fn build(self) -> Shape {
        Shape {
            shape_type: self.shape_type,
            drawing_kind: self.drawing_kind,
            drawing_attributes: self.drawing_attributes,
            children: self.children,
            enhanced_geometry: self.enhanced_geometry,
            text: self.text,
            name: self.name,
            x: self.x,
            y: self.y,
            width: self.width,
            height: self.height,
            style_name: self.style_name,
            layer: self.layer,
            z_index: self.z_index,
            transform: self.transform,
            presentation_class: self.presentation_class,
            presentation_placeholder: self.presentation_placeholder,
            presentation_user_transformed: self.presentation_user_transformed,
            image_href: self.image_href,
            media: self.media,
            hyperlink: self.hyperlink,
            event_listeners: self.event_listeners,
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

impl Parser {
    fn is_namespace(namespace: &ResolveResult<'_>, expected: &[u8]) -> bool {
        matches!(namespace, ResolveResult::Bound(XmlNamespace(value)) if *value == expected)
    }

    /// Rewinds the running element depth for a subtree consumed in place.
    ///
    /// Nested parsers such as [`Self::parse_enhanced_geometry`] read through
    /// their own element's `Event::End`, so the main event loop never observes
    /// it and would otherwise keep counting that element as open. Spreadsheet
    /// `table:shapes` container tracking compares depths exactly, so the
    /// increment taken on `Event::Start` has to be given back here.
    const fn rewind_consumed_subtree(element_depth: usize) -> usize {
        element_depth.saturating_sub(1)
    }

    fn classify(namespace: &ResolveResult<'_>, local_name: &[u8]) -> Element {
        if Self::is_namespace(namespace, ANIMATION_NAMESPACE_BYTES) {
            Kind::from_local_name(local_name)
                .map(Element::Animation)
                .unwrap_or(Element::UnknownAnimation)
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
                LegacyAnimationKind::from_local_name(local_name)
                    .map(Element::LegacyAnimation)
                    .unwrap_or(Element::Other)
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

    fn drawing_kind(shape_element: ShapeElement) -> DrawingShapeKind {
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

    fn validate_shape_parent(parent: &ShapeBuilder, child: DrawingShapeKind) -> Result<()> {
        match parent.drawing_kind {
            Some(DrawingShapeKind::Group) => {
                if child.is_three_dimensional() && child != DrawingShapeKind::ThreeDimensionalScene
                {
                    return Err(Error::InvalidFormat(
                        "3D drawing objects require a dr3d:scene parent".to_string(),
                    ));
                }
            },
            Some(DrawingShapeKind::ThreeDimensionalScene) => {
                if !child.is_three_dimensional() {
                    return Err(Error::InvalidFormat(
                        "dr3d:scene can only contain 3D lights and objects".to_string(),
                    ));
                }
                if child == DrawingShapeKind::ThreeDimensionalLight
                    && parent.children.iter().any(|existing| {
                        existing.drawing_kind() != Some(DrawingShapeKind::ThreeDimensionalLight)
                    })
                {
                    return Err(Error::InvalidFormat(
                        "dr3d:light elements must precede 3D objects".to_string(),
                    ));
                }
            },
            _ => {
                return Err(Error::InvalidFormat(
                    "nested drawing shapes require a draw:g or dr3d:scene parent".to_string(),
                ));
            },
        }
        Ok(())
    }

    fn validate_three_dimensional_child_element(
        parent: Option<&ShapeBuilder>,
        child: Element,
    ) -> Result<()> {
        let Some(parent_kind) = parent.and_then(|builder| builder.drawing_kind) else {
            return Ok(());
        };
        if !parent_kind.is_three_dimensional() {
            return Ok(());
        }
        if parent_kind != DrawingShapeKind::ThreeDimensionalScene {
            return Err(Error::InvalidFormat(
                "3D light and object elements cannot contain child elements".to_string(),
            ));
        }
        match child {
            Element::Shape(shape) if Self::drawing_kind(shape).is_three_dimensional() => Ok(()),
            // `svg:title`, `svg:desc`, `draw:glue-point`, and foreign
            // extension elements are intentionally handled as opaque content.
            Element::Other => Ok(()),
            _ => Err(Error::InvalidFormat(
                "dr3d:scene can only contain 3D content".to_string(),
            )),
        }
    }

    fn validate_required_three_dimensional_attributes(
        kind: DrawingShapeKind,
        attributes: &[DrawingAttribute],
    ) -> Result<()> {
        let has = |namespace, local_name| {
            attributes.iter().any(|attribute| {
                attribute.namespace() == namespace && attribute.local_name() == local_name
            })
        };
        if kind == DrawingShapeKind::ThreeDimensionalLight
            && !has(DrawingAttributeNamespace::Dr3d, "direction")
        {
            return Err(Error::InvalidFormat(
                "dr3d:light requires dr3d:direction".to_string(),
            ));
        }
        if matches!(
            kind,
            DrawingShapeKind::ThreeDimensionalExtrude | DrawingShapeKind::ThreeDimensionalRotate
        ) {
            for local_name in ["viewBox", "d"] {
                if !has(DrawingAttributeNamespace::Svg, local_name) {
                    return Err(Error::InvalidFormat(format!(
                        "{} requires svg:{local_name}",
                        kind.element_name()
                    )));
                }
            }
        }
        Ok(())
    }

    fn animation_attributes(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Vec<Attribute>> {
        if element.attributes().count() > 256 {
            return Err(Error::InvalidFormat(
                "ODP animation node exceeds 256 attributes".to_string(),
            ));
        }
        let mut attributes = Vec::with_capacity(element.attributes().count());
        let mut expanded_names = HashSet::new();
        for attribute in element.attributes() {
            let attribute = attribute
                .map_err(|error| Error::InvalidFormat(format!("invalid XML attribute: {error}")))?;
            let qualified_name = attribute.key.as_ref();
            if qualified_name == b"xmlns" || qualified_name.starts_with(b"xmlns:") {
                continue;
            }
            let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
            let namespace_uri = match namespace {
                ResolveResult::Unbound => None,
                ResolveResult::Bound(XmlNamespace(uri)) => {
                    Some(std::str::from_utf8(uri).map_err(|_| {
                        Error::InvalidFormat("non-UTF-8 animation namespace URI".to_string())
                    })?)
                },
                ResolveResult::Unknown(prefix) => {
                    return Err(Error::InvalidFormat(format!(
                        "unknown animation attribute namespace prefix '{}'",
                        String::from_utf8_lossy(&prefix)
                    )));
                },
            };
            let local_name = std::str::from_utf8(local_name.as_ref())
                .map_err(|_| {
                    Error::InvalidFormat("non-UTF-8 animation attribute name".to_string())
                })?
                .to_string();
            let namespace = Namespace::from_uri(namespace_uri);
            if !expanded_names.insert((namespace.clone(), local_name.clone())) {
                return Err(Error::InvalidFormat(format!(
                    "duplicate animation attribute '{local_name}'"
                )));
            }
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!("invalid XML attribute value: {error}"))
                })?
                .into_owned();
            if value.len() > 1_048_576 {
                return Err(Error::InvalidFormat(
                    "ODP animation attribute exceeds 1 MiB".to_string(),
                ));
            }
            attributes.push(Attribute::from_parsed(namespace, local_name, value)?);
        }
        Ok(attributes)
    }

    fn parse_animation_node(
        reader: &mut NsReader<&[u8]>,
        start: &BytesStart<'_>,
        kind: Kind,
        depth: usize,
        node_count: &mut usize,
    ) -> Result<Node> {
        if depth > 128 {
            return Err(Error::InvalidFormat(
                "ODP animation nesting exceeds 128 levels".to_string(),
            ));
        }
        *node_count = node_count
            .checked_add(1)
            .ok_or_else(|| Error::InvalidFormat("ODP animation node count overflow".to_string()))?;
        if *node_count > 65_536 {
            return Err(Error::InvalidFormat(
                "ODP animation tree exceeds 65536 nodes".to_string(),
            ));
        }
        let attributes = Self::animation_attributes(reader, start)?;
        let mut children = Vec::new();
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref child) | Event::Empty(ref child) => {
                    if !Self::is_namespace(&namespace, ANIMATION_NAMESPACE_BYTES) {
                        return Err(Error::InvalidFormat(format!(
                            "anim:{} contains a non-animation element",
                            kind.local_name()
                        )));
                    }
                    let Some(child_kind) = Kind::from_local_name(child.local_name().as_ref())
                    else {
                        return Err(Error::InvalidFormat(format!(
                            "unknown ODF animation element '{}'",
                            String::from_utf8_lossy(child.local_name().as_ref())
                        )));
                    };
                    if !kind.allows_child(child_kind) {
                        return Err(Error::InvalidFormat(format!(
                            "anim:{} cannot contain anim:{}",
                            kind.local_name(),
                            child_kind.local_name()
                        )));
                    }
                    let node = if matches!(event, Event::Empty(_)) {
                        *node_count = node_count.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat("ODP animation node count overflow".to_string())
                        })?;
                        if *node_count > 65_536 {
                            return Err(Error::InvalidFormat(
                                "ODP animation tree exceeds 65536 nodes".to_string(),
                            ));
                        }
                        Node::from_parsed(
                            child_kind,
                            Self::animation_attributes(reader, child)?,
                            Vec::new(),
                        )
                    } else {
                        Self::parse_animation_node(
                            reader,
                            child,
                            child_kind,
                            depth + 1,
                            node_count,
                        )?
                    };
                    children.push(node);
                },
                Event::End(ref end) => {
                    if !Self::is_namespace(&namespace, ANIMATION_NAMESPACE_BYTES)
                        || end.local_name().as_ref() != kind.local_name().as_bytes()
                    {
                        return Err(Error::InvalidFormat(format!(
                            "unexpected closing element in anim:{}",
                            kind.local_name()
                        )));
                    }
                    return Ok(Node::from_parsed(kind, attributes, children));
                },
                Event::Text(ref text) => {
                    let text = Self::decode_text(text)?;
                    if !text.trim().is_empty() {
                        return Err(Error::InvalidFormat(format!(
                            "anim:{} cannot contain text",
                            kind.local_name()
                        )));
                    }
                },
                Event::CData(ref text) => {
                    let text = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid animation CDATA: {error}"))
                    })?;
                    if !text.trim().is_empty() {
                        return Err(Error::InvalidFormat(format!(
                            "anim:{} cannot contain text",
                            kind.local_name()
                        )));
                    }
                },
                Event::Eof => {
                    return Err(Error::InvalidFormat(format!(
                        "unterminated anim:{} element",
                        kind.local_name()
                    )));
                },
                Event::GeneralRef(_) => {
                    return Err(Error::InvalidFormat(format!(
                        "anim:{} cannot contain character references",
                        kind.local_name()
                    )));
                },
                _ => {},
            }
            buffer.clear();
        }
    }

    fn parse_legacy_animation_node(
        reader: &mut NsReader<&[u8]>,
        start: &BytesStart<'_>,
        kind: LegacyAnimationKind,
        depth: usize,
        node_count: &mut usize,
    ) -> Result<LegacyAnimationNode> {
        if depth > 128 {
            return Err(Error::InvalidFormat(
                "legacy ODP animation nesting exceeds 128 levels".to_string(),
            ));
        }
        *node_count = node_count.checked_add(1).ok_or_else(|| {
            Error::InvalidFormat("legacy ODP animation node count overflow".to_string())
        })?;
        if *node_count > 65_536 {
            return Err(Error::InvalidFormat(
                "legacy ODP animation tree exceeds 65536 nodes".to_string(),
            ));
        }
        let attributes = Self::animation_attributes(reader, start)?;
        let mut children = Vec::new();
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref child) | Event::Empty(ref child) => {
                    if !Self::is_namespace(&namespace, PRESENTATION_NAMESPACE) {
                        return Err(Error::InvalidFormat(format!(
                            "presentation:{} contains a foreign element",
                            kind.local_name()
                        )));
                    }
                    let child_kind =
                        LegacyAnimationKind::from_local_name(child.local_name().as_ref())
                            .ok_or_else(|| {
                                Error::InvalidFormat(format!(
                                    "unknown legacy presentation animation element '{}'",
                                    String::from_utf8_lossy(child.local_name().as_ref())
                                ))
                            })?;
                    if !kind.allows_child(child_kind) {
                        return Err(Error::InvalidFormat(format!(
                            "presentation:{} cannot contain presentation:{}",
                            kind.local_name(),
                            child_kind.local_name()
                        )));
                    }
                    let node = if matches!(event, Event::Empty(_)) {
                        *node_count = node_count.checked_add(1).ok_or_else(|| {
                            Error::InvalidFormat(
                                "legacy ODP animation node count overflow".to_string(),
                            )
                        })?;
                        if *node_count > 65_536 {
                            return Err(Error::InvalidFormat(
                                "legacy ODP animation tree exceeds 65536 nodes".to_string(),
                            ));
                        }
                        LegacyAnimationNode::from_parsed(
                            child_kind,
                            Self::animation_attributes(reader, child)?,
                            Vec::new(),
                        )
                    } else {
                        Self::parse_legacy_animation_node(
                            reader,
                            child,
                            child_kind,
                            depth + 1,
                            node_count,
                        )?
                    };
                    children.push(node);
                },
                Event::End(ref end)
                    if Self::is_namespace(&namespace, PRESENTATION_NAMESPACE)
                        && end.local_name().as_ref() == kind.local_name().as_bytes() =>
                {
                    return Ok(LegacyAnimationNode::from_parsed(kind, attributes, children));
                },
                Event::Text(ref text) if !Self::decode_text(text)?.trim().is_empty() => {
                    return Err(Error::InvalidFormat(
                        "legacy presentation animations cannot contain text".to_string(),
                    ));
                },
                Event::CData(ref text)
                    if !text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?
                        .trim()
                        .is_empty() =>
                {
                    return Err(Error::InvalidFormat(
                        "legacy presentation animations cannot contain text".to_string(),
                    ));
                },
                Event::Eof => {
                    return Err(Error::InvalidFormat(
                        "unterminated legacy presentation animation tree".to_string(),
                    ));
                },
                Event::End(_) | Event::GeneralRef(_) => {
                    return Err(Error::InvalidFormat(
                        "invalid content in legacy presentation animation tree".to_string(),
                    ));
                },
                _ => {},
            }
            buffer.clear();
        }
    }

    fn shape_builder(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        shape_element: ShapeElement,
    ) -> Result<ShapeBuilder> {
        let mut builder = ShapeBuilder::new();
        let presentation_class = Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"class")?;
        builder.is_frame = matches!(shape_element, ShapeElement::Frame);
        builder.drawing_kind = Some(Self::drawing_kind(shape_element));
        builder.is_title = presentation_class.as_deref() == Some("title");
        builder.shape_type = match shape_element {
            ShapeElement::Frame => match presentation_class.as_deref() {
                Some(_) => ShapeType::Placeholder,
                _ => ShapeType::TextBox,
            },
            ShapeElement::Line | ShapeElement::Measure => ShapeType::Line,
            ShapeElement::Connector => ShapeType::Connector,
            ShapeElement::Group | ShapeElement::ThreeDimensionalScene => ShapeType::Group,
            _ => ShapeType::AutoShape,
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
            builder.drawing_kind.expect("shape kind initialized"),
            &builder.drawing_attributes,
        )?;
        Ok(builder)
    }

    fn drawing_attributes(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Vec<DrawingAttribute>> {
        let mut attributes = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| {
                Error::InvalidFormat(format!("invalid ODP shape attribute: {error}"))
            })?;
            let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
            let namespace = if Self::is_namespace(&namespace, DRAW_NAMESPACE) {
                DrawingAttributeNamespace::Drawing
            } else if Self::is_namespace(&namespace, SVG_NAMESPACE) {
                DrawingAttributeNamespace::Svg
            } else if Self::is_namespace(&namespace, DR3D_NAMESPACE) {
                DrawingAttributeNamespace::Dr3d
            } else if Self::is_namespace(&namespace, TABLE_NAMESPACE) {
                DrawingAttributeNamespace::Table
            } else {
                continue;
            };
            let local_name = local_name.as_ref();
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
                String::from_utf8(local_name.to_vec()).map_err(|_| {
                    Error::InvalidFormat("non-UTF-8 ODP shape attribute name".to_string())
                })?,
                value,
            )?);
        }
        Ok(attributes)
    }

    fn exact_geometry_attributes(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<Vec<DrawingAttribute>> {
        let mut attributes = Vec::new();
        for attribute in element.attributes() {
            let attribute = attribute.map_err(|error| {
                Error::InvalidFormat(format!("invalid enhanced-geometry attribute: {error}"))
            })?;
            let (namespace, local_name) = reader.resolver().resolve_attribute(attribute.key);
            let namespace = if Self::is_namespace(&namespace, DRAW_NAMESPACE) {
                DrawingAttributeNamespace::Drawing
            } else if Self::is_namespace(&namespace, SVG_NAMESPACE) {
                DrawingAttributeNamespace::Svg
            } else if Self::is_namespace(&namespace, DR3D_NAMESPACE) {
                DrawingAttributeNamespace::Dr3d
            } else {
                continue;
            };
            let value = attribute
                .decoded_and_normalized_value(XmlVersion::Implicit1_0, reader.decoder())
                .map_err(|error| {
                    Error::InvalidFormat(format!(
                        "invalid enhanced-geometry attribute value: {error}"
                    ))
                })?
                .into_owned();
            attributes.push(DrawingAttribute::new(
                namespace,
                String::from_utf8(local_name.as_ref().to_vec()).map_err(|_| {
                    Error::InvalidFormat("non-UTF-8 enhanced-geometry attribute name".to_string())
                })?,
                value,
            )?);
        }
        Ok(attributes)
    }

    fn parse_enhanced_geometry(
        reader: &mut NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<EnhancedGeometry> {
        let attributes = Self::exact_geometry_attributes(reader, element)?;
        let mut children = Vec::new();
        let mut handle_seen = false;
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) =
                reader
                    .read_resolved_event_into(&mut buffer)
                    .map_err(|error| {
                        Error::InvalidFormat(format!("invalid enhanced geometry XML: {error}"))
                    })?;
            match event {
                Event::Start(ref child) | Event::Empty(ref child)
                    if Self::is_namespace(&namespace, DRAW_NAMESPACE)
                        && matches!(child.local_name().as_ref(), b"equation" | b"handle") =>
                {
                    if children.len() >= 65_536 {
                        return Err(Error::InvalidFormat(
                            "enhanced geometry exceeds 65536 equations and handles".to_string(),
                        ));
                    }
                    let kind = if child.local_name().as_ref() == b"equation" {
                        if handle_seen {
                            return Err(Error::InvalidFormat(
                                "draw:equation cannot follow draw:handle".to_string(),
                            ));
                        }
                        EnhancedGeometryChildKind::Equation
                    } else {
                        handle_seen = true;
                        EnhancedGeometryChildKind::Handle
                    };
                    children.push(EnhancedGeometryChild {
                        kind,
                        attributes: Self::exact_geometry_attributes(reader, child)?,
                    });
                    if matches!(event, Event::Start(_)) {
                        Self::consume_empty_content(
                            reader,
                            DRAW_NAMESPACE,
                            child.local_name().as_ref(),
                            kind.element_name(),
                        )?;
                    }
                },
                Event::End(ref end)
                    if Self::is_namespace(&namespace, DRAW_NAMESPACE)
                        && end.local_name().as_ref() == b"enhanced-geometry" =>
                {
                    return Ok(EnhancedGeometry {
                        attributes,
                        children,
                    });
                },
                Event::Text(ref text) if Self::decode_text(text)?.trim().is_empty() => {},
                Event::CData(ref text)
                    if text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?
                        .trim()
                        .is_empty() => {},
                Event::Comment(_) | Event::PI(_) => {},
                Event::Eof => {
                    return Err(Error::InvalidFormat(
                        "unterminated draw:enhanced-geometry".to_string(),
                    ));
                },
                _ => {
                    return Err(Error::InvalidFormat(
                        "draw:enhanced-geometry may only contain equations and handles".to_string(),
                    ));
                },
            }
            buffer.clear();
        }
    }

    fn media_reference(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Reference> {
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

    fn media_parameter(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<Parameter> {
        let name = Self::get_attr(reader, element, DRAW_NAMESPACE, b"name")?.ok_or_else(|| {
            Error::InvalidFormat("draw:param is missing required draw:name".to_string())
        })?;
        let value =
            Self::get_attr(reader, element, DRAW_NAMESPACE, b"value")?.ok_or_else(|| {
                Error::InvalidFormat("draw:param is missing required draw:value".to_string())
            })?;
        Parameter::new(name, value)
    }

    fn required_attr(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        namespace: &[u8],
        local_name: &[u8],
        qualified_name: &str,
    ) -> Result<String> {
        Self::get_attr(reader, element, namespace, local_name)?.ok_or_else(|| {
            Error::InvalidFormat(format!("element is missing required {qualified_name}"))
        })
    }

    fn require_simple_xlink(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
        description: &str,
    ) -> Result<()> {
        let link_type =
            Self::required_attr(reader, element, XLINK_NAMESPACE, b"type", "xlink:type")?;
        if link_type != "simple" {
            return Err(Error::InvalidFormat(format!(
                "{description} xlink:type must be 'simple', found '{link_type}'"
            )));
        }
        Ok(())
    }

    fn parse_on_request(value: Option<String>, description: &str) -> Result<bool> {
        match value.as_deref() {
            None => Ok(false),
            Some("onRequest") => Ok(true),
            Some(value) => Err(Error::InvalidFormat(format!(
                "invalid {description} xlink:actuate '{value}'"
            ))),
        }
    }

    fn drawing_hyperlink(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<DrawingHyperlink> {
        Self::require_simple_xlink(reader, element, "draw:a")?;
        let href = Self::required_attr(reader, element, XLINK_NAMESPACE, b"href", "xlink:href")?;
        let mut hyperlink = DrawingHyperlink::new(href)?;
        hyperlink.set_actuate_on_request(Self::parse_on_request(
            Self::get_attr(reader, element, XLINK_NAMESPACE, b"actuate")?,
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

    fn script_event_listener(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<ScriptEventListener> {
        let event_name = Self::required_attr(
            reader,
            element,
            SCRIPT_NAMESPACE,
            b"event-name",
            "script:event-name",
        )?;
        let language = Self::required_attr(
            reader,
            element,
            SCRIPT_NAMESPACE,
            b"language",
            "script:language",
        )?;
        let macro_name = Self::get_attr(reader, element, SCRIPT_NAMESPACE, b"macro-name")?;
        let href = Self::get_attr(reader, element, XLINK_NAMESPACE, b"href")?;
        let link_type = Self::get_attr(reader, element, XLINK_NAMESPACE, b"type")?;
        if href.is_some() {
            Self::require_simple_xlink(reader, element, "script:event-listener")?;
        } else if link_type.is_some() {
            return Err(Error::InvalidFormat(
                "script:event-listener xlink:type requires xlink:href".to_string(),
            ));
        }
        let listener = ScriptEventListener {
            event_name,
            language,
            macro_name,
            href,
            actuate_on_request: Self::parse_on_request(
                Self::get_attr(reader, element, XLINK_NAMESPACE, b"actuate")?,
                "script:event-listener",
            )?,
        };
        listener.validate()?;
        Ok(listener)
    }

    fn presentation_event_listener(
        reader: &NsReader<&[u8]>,
        element: &BytesStart<'_>,
    ) -> Result<EventListener> {
        let event_name = Self::required_attr(
            reader,
            element,
            SCRIPT_NAMESPACE,
            b"event-name",
            "script:event-name",
        )?;
        let action = Action::parse(&Self::required_attr(
            reader,
            element,
            PRESENTATION_NAMESPACE,
            b"action",
            "presentation:action",
        )?)?;
        let mut listener = EventListener::new(event_name, action)?;
        listener.effect = Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"effect")?
            .map(Effect::new)
            .transpose()?;
        listener.direction = Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"direction")?
            .map(EffectDirection::new)
            .transpose()?;
        listener.speed = Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"speed")?
            .map(|value| TransitionSpeed::parse(&value))
            .transpose()?;
        listener.start_scale =
            Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"start-scale")?;
        listener.href = Self::get_attr(reader, element, XLINK_NAMESPACE, b"href")?;
        let link_type = Self::get_attr(reader, element, XLINK_NAMESPACE, b"type")?;
        if listener.href.is_some() {
            Self::require_simple_xlink(reader, element, "presentation:event-listener")?;
        } else if link_type.is_some() {
            return Err(Error::InvalidFormat(
                "presentation:event-listener xlink:type requires xlink:href".to_string(),
            ));
        }
        listener.show_embed =
            match Self::get_attr(reader, element, XLINK_NAMESPACE, b"show")?.as_deref() {
                None => false,
                Some("embed") => true,
                Some(value) => {
                    return Err(Error::InvalidFormat(format!(
                        "invalid presentation:event-listener xlink:show '{value}'"
                    )));
                },
            };
        listener.actuate_on_request = Self::parse_on_request(
            Self::get_attr(reader, element, XLINK_NAMESPACE, b"actuate")?,
            "presentation:event-listener",
        )?;
        listener.verb = Self::get_attr(reader, element, PRESENTATION_NAMESPACE, b"verb")?
            .map(|value| {
                value.parse::<u64>().map_err(|_| {
                    Error::InvalidFormat(format!("invalid presentation:verb '{value}'"))
                })
            })
            .transpose()?;
        listener.validate()?;
        Ok(listener)
    }

    fn consume_empty_content(
        reader: &mut NsReader<&[u8]>,
        namespace_uri: &[u8],
        local_name: &[u8],
        description: &str,
    ) -> Result<()> {
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::End(ref end)
                    if Self::is_namespace(&namespace, namespace_uri)
                        && end.local_name().as_ref() == local_name =>
                {
                    return Ok(());
                },
                Event::Text(ref text) if Self::decode_text(text)?.trim().is_empty() => {},
                Event::CData(ref text)
                    if text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?
                        .trim()
                        .is_empty() => {},
                Event::Eof => {
                    return Err(Error::InvalidFormat(format!(
                        "unterminated {description} element"
                    )));
                },
                _ => {
                    return Err(Error::InvalidFormat(format!(
                        "{description} must not contain content"
                    )));
                },
            }
            buffer.clear();
        }
    }

    fn parse_listener_body(
        reader: &mut NsReader<&[u8]>,
        mut listener: EventListener,
    ) -> Result<EventListener> {
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element)
                    if Self::is_namespace(&namespace, PRESENTATION_NAMESPACE)
                        && element.local_name().as_ref() == b"sound" =>
                {
                    if listener.sound.is_some() {
                        return Err(Error::InvalidFormat(
                            "presentation event listener contains multiple sounds".to_string(),
                        ));
                    }
                    Self::require_simple_xlink(reader, element, "presentation:sound")?;
                    listener.sound = Some(Self::parse_transition_sound(reader, element)?);
                    if matches!(event, Event::Start(_)) {
                        Self::consume_empty_content(
                            reader,
                            PRESENTATION_NAMESPACE,
                            b"sound",
                            "presentation:sound",
                        )?;
                    }
                },
                Event::End(ref end)
                    if Self::is_namespace(&namespace, PRESENTATION_NAMESPACE)
                        && end.local_name().as_ref() == b"event-listener" =>
                {
                    listener.validate()?;
                    return Ok(listener);
                },
                Event::Text(ref text) if Self::decode_text(text)?.trim().is_empty() => {},
                Event::CData(ref text)
                    if text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?
                        .trim()
                        .is_empty() => {},
                Event::Eof => {
                    return Err(Error::InvalidFormat(
                        "unterminated presentation:event-listener".to_string(),
                    ));
                },
                _ => {
                    return Err(Error::InvalidFormat(
                        "presentation:event-listener may only contain presentation:sound"
                            .to_string(),
                    ));
                },
            }
            buffer.clear();
        }
    }

    fn parse_event_listeners(reader: &mut NsReader<&[u8]>) -> Result<Vec<ShapeEventListener>> {
        let mut listeners = Vec::new();
        let mut buffer = Vec::new();
        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buffer)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref element) | Event::Empty(ref element)
                    if Self::is_namespace(&namespace, SCRIPT_NAMESPACE)
                        && element.local_name().as_ref() == b"event-listener" =>
                {
                    if listeners.len() >= 4096 {
                        return Err(Error::InvalidFormat(
                            "ODP shape exceeds 4096 event listeners".to_string(),
                        ));
                    }
                    let listener = Self::script_event_listener(reader, element)?;
                    if matches!(event, Event::Start(_)) {
                        Self::consume_empty_content(
                            reader,
                            SCRIPT_NAMESPACE,
                            b"event-listener",
                            "script:event-listener",
                        )?;
                    }
                    listeners.push(ShapeEventListener::Script(listener));
                },
                Event::Start(ref element) | Event::Empty(ref element)
                    if Self::is_namespace(&namespace, PRESENTATION_NAMESPACE)
                        && element.local_name().as_ref() == b"event-listener" =>
                {
                    if listeners.len() >= 4096 {
                        return Err(Error::InvalidFormat(
                            "ODP shape exceeds 4096 event listeners".to_string(),
                        ));
                    }
                    let listener = Self::presentation_event_listener(reader, element)?;
                    let listener = if matches!(event, Event::Start(_)) {
                        Self::parse_listener_body(reader, listener)?
                    } else {
                        listener
                    };
                    listeners.push(ShapeEventListener::Action(Box::new(listener)));
                },
                Event::End(ref end)
                    if Self::is_namespace(&namespace, OFFICE_NAMESPACE)
                        && end.local_name().as_ref() == b"event-listeners" =>
                {
                    return Ok(listeners);
                },
                Event::Text(ref text) if Self::decode_text(text)?.trim().is_empty() => {},
                Event::CData(ref text)
                    if text
                        .xml_content(XmlVersion::Explicit1_0)
                        .map_err(|error| Error::InvalidFormat(error.to_string()))?
                        .trim()
                        .is_empty() => {},
                Event::Eof => {
                    return Err(Error::InvalidFormat(
                        "unterminated office:event-listeners".to_string(),
                    ));
                },
                _ => {
                    return Err(Error::InvalidFormat(
                        "office:event-listeners may only contain script or presentation listeners"
                            .to_string(),
                    ));
                },
            }
            buffer.clear();
        }
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
        element_type: Element,
        paragraph: &mut ParagraphText,
    ) -> Result<()> {
        match element_type {
            Element::TextLineBreak => paragraph.push_explicit('\n', 1),
            Element::TextTab => paragraph.push_explicit('\t', 1),
            Element::TextSpace => {
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
        transition: &mut Transition,
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
                            transition: Transition::new(),
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
                            transition: Transition::new(),
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
    ) -> Result<(HashMap<String, Transition>, Transition)> {
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
    pub(crate) fn parse_slides(xml_content: &str) -> Result<Vec<Slide>> {
        Self::parse_slides_with_styles(xml_content, None)
    }

    /// Parse slides and resolve drawing-page transition styles.
    pub(crate) fn parse_slides_with_styles(
        xml_content: &str,
        styles_xml: Option<&str>,
    ) -> Result<Vec<Slide>> {
        Self::parse_pages_with_styles(
            xml_content,
            styles_xml,
            false,
            ShapeContainerScope::DrawPages,
        )
    }

    /// Parse drawing pages while retaining title and text-box frames as shapes.
    #[allow(dead_code, reason = "reserved for the dedicated ODG facade")]
    pub(crate) fn parse_drawing_pages(
        xml_content: &str,
        styles_xml: Option<&str>,
    ) -> Result<Vec<Slide>> {
        Self::parse_pages_with_styles(
            xml_content,
            styles_xml,
            true,
            ShapeContainerScope::DrawPages,
        )
    }

    /// Parse `table:shapes` drawing shapes from spreadsheet content.
    ///
    /// Returns one shape list per top-level `table:table` element, in
    /// document order, retaining text-box frames as shapes. Shapes anchored
    /// inside individual table cells are not collected.
    #[allow(dead_code, reason = "reserved for the dedicated ODS facade")]
    pub(crate) fn parse_sheet_shape_tables(xml_content: &str) -> Result<Vec<Vec<Shape>>> {
        let tables = Self::parse_pages_with_styles(
            xml_content,
            None,
            true,
            ShapeContainerScope::SpreadsheetTables,
        )?;
        Ok(tables.into_iter().map(|table| table.shapes).collect())
    }

    fn parse_pages_with_styles(
        xml_content: &str,
        styles_xml: Option<&str>,
        retain_text_shapes: bool,
        container_scope: ShapeContainerScope,
    ) -> Result<Vec<Slide>> {
        let sheet_scope = container_scope == ShapeContainerScope::SpreadsheetTables;
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
        let mut current_transition: Option<Transition> = None;
        let mut current_animations = Vec::new();
        let mut animation_node_count = 0;
        let mut current_legacy_animation = None;
        let mut legacy_animation_node_count = 0;
        let mut shape_node_count = 0usize;

        // Shape parsing state
        let mut shape_stack: Vec<ShapeBuilder> = Vec::new();
        let mut current_paragraph: Option<ParagraphText> = None;
        let mut in_media_plugin = false;
        let mut in_media_parameter = false;
        let mut current_hyperlink: Option<DrawingHyperlink> = None;
        let mut hyperlink_parent_depth = None;
        let mut hyperlink_shape_seen = false;

        // Spreadsheet `table:shapes` container state
        let mut element_depth = 0usize;
        let mut spreadsheet_depth: Option<usize> = None;
        let mut sheet_table_depth: Option<usize> = None;
        let mut sheet_shapes_depth: Option<usize> = None;
        let mut sheet_table_has_shapes = false;

        loop {
            let (namespace, event) = reader
                .read_resolved_event_into(&mut buf)
                .map_err(|error| Error::InvalidFormat(format!("XML parsing error: {error}")))?;
            match event {
                Event::Start(ref element) => {
                    element_depth = element_depth.checked_add(1).ok_or_else(|| {
                        Error::InvalidFormat("XML element depth overflow".to_string())
                    })?;
                    let element_type = Self::classify(&namespace, element.local_name().as_ref());
                    Self::validate_three_dimensional_child_element(
                        shape_stack.last(),
                        element_type,
                    )?;
                    if in_media_parameter {
                        return Err(Error::InvalidFormat(
                            "draw:param cannot contain child elements".to_string(),
                        ));
                    }
                    if in_media_plugin && !matches!(element_type, Element::PluginParameter) {
                        return Err(Error::InvalidFormat(
                            "draw:plugin can only contain draw:param elements".to_string(),
                        ));
                    }
                    match element_type {
                        Element::Page if !sheet_scope => {
                            if in_slide {
                                slides.push(Slide {
                                    title: current_slide_title.take(),
                                    text: std::mem::take(&mut current_slide_text),
                                    index: slide_index,
                                    notes: (!current_notes_text.is_empty())
                                        .then(|| std::mem::take(&mut current_notes_text)),
                                    transition: current_transition.take(),
                                    animations: std::mem::take(&mut current_animations),
                                    legacy_animation: current_legacy_animation.take(),
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
                        Element::Notes if in_slide => in_notes = true,
                        Element::EnhancedGeometry if !shape_stack.is_empty() => {
                            let builder = shape_stack.last().expect("shape checked above");
                            if builder.drawing_kind != Some(DrawingShapeKind::CustomShape) {
                                return Err(Error::InvalidFormat(
                                    "draw:enhanced-geometry requires draw:custom-shape".to_string(),
                                ));
                            }
                            if builder.enhanced_geometry.is_some() {
                                return Err(Error::InvalidFormat(
                                    "draw:custom-shape contains multiple enhanced geometries"
                                        .to_string(),
                                ));
                            }
                            let geometry = Self::parse_enhanced_geometry(&mut reader, element)?;
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                            shape_stack
                                .last_mut()
                                .expect("shape checked above")
                                .enhanced_geometry = Some(geometry);
                        },
                        Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "misplaced custom-shape enhanced geometry".to_string(),
                            ));
                        },
                        Element::LegacyAnimation(kind)
                            if in_slide
                                && !in_notes
                                && shape_stack.is_empty()
                                && current_paragraph.is_none() =>
                        {
                            if kind != LegacyAnimationKind::Animations {
                                return Err(Error::InvalidFormat(
                                    "legacy presentation effects require a presentation:animations root"
                                        .to_string(),
                                ));
                            }
                            if current_legacy_animation.is_some() {
                                return Err(Error::InvalidFormat(
                                    "ODP slide contains multiple presentation:animations roots"
                                        .to_string(),
                                ));
                            }
                            let root = Self::parse_legacy_animation_node(
                                &mut reader,
                                element,
                                kind,
                                1,
                                &mut legacy_animation_node_count,
                            )?;
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                            validate_legacy_animation_root(&root)?;
                            current_legacy_animation = Some(root);
                        },
                        Element::Plugin if !shape_stack.is_empty() => {
                            let builder = shape_stack.last_mut().expect("shape checked above");
                            if !builder.is_frame {
                                return Err(Error::InvalidFormat(
                                    "draw:plugin must be contained directly by draw:frame"
                                        .to_string(),
                                ));
                            }
                            if builder.media.is_some() {
                                return Err(Error::InvalidFormat(
                                    "ODP frame contains multiple draw:plugin elements".to_string(),
                                ));
                            }
                            builder.shape_type = ShapeType::GraphicFrame;
                            builder.media = Some(Self::media_reference(&reader, element)?);
                            in_media_plugin = true;
                        },
                        Element::Plugin if in_slide => {
                            return Err(Error::InvalidFormat(
                                "draw:plugin must be contained by a drawing shape".to_string(),
                            ));
                        },
                        Element::PluginParameter
                            if in_media_plugin
                                && !in_media_parameter
                                && !shape_stack.is_empty() =>
                        {
                            shape_stack
                                .last_mut()
                                .and_then(|builder| builder.media.as_mut())
                                .expect("media plugin state checked above")
                                .add_parameter(Self::media_parameter(&reader, element)?)?;
                            in_media_parameter = true;
                        },
                        Element::DrawingHyperlink
                            if in_slide && !in_notes && current_hyperlink.is_none() =>
                        {
                            current_hyperlink = Some(Self::drawing_hyperlink(&reader, element)?);
                            hyperlink_parent_depth = Some(shape_stack.len());
                            hyperlink_shape_seen = false;
                        },
                        Element::DrawingHyperlink if in_slide => {
                            return Err(Error::InvalidFormat(
                                "nested or misplaced draw:a presentation hyperlink".to_string(),
                            ));
                        },
                        Element::EventListeners if !shape_stack.is_empty() => {
                            let builder = shape_stack.last_mut().expect("shape checked above");
                            if builder
                                .drawing_kind
                                .is_some_and(|kind| kind.is_three_dimensional())
                            {
                                return Err(Error::InvalidFormat(
                                    "3D shapes cannot contain presentation event listeners"
                                        .to_string(),
                                ));
                            }
                            if builder.event_listeners_seen {
                                return Err(Error::InvalidFormat(
                                    "ODP shape contains multiple office:event-listeners elements"
                                        .to_string(),
                                ));
                            }
                            builder.event_listeners = Self::parse_event_listeners(&mut reader)?;
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                            builder.event_listeners_seen = true;
                        },
                        Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "presentation event metadata must be contained by a shape's office:event-listeners"
                                    .to_string(),
                            ));
                        },
                        _ if in_media_parameter => {
                            return Err(Error::InvalidFormat(
                                "draw:param cannot contain child elements".to_string(),
                            ));
                        },
                        _ if in_media_plugin => {
                            return Err(Error::InvalidFormat(
                                "draw:plugin can only contain draw:param elements".to_string(),
                            ));
                        },
                        Element::TextParagraph if in_slide => {
                            if current_paragraph.is_some() {
                                return Err(Error::InvalidFormat(
                                    "nested ODP text paragraphs are not supported".to_string(),
                                ));
                            }
                            current_paragraph = Some(ParagraphText::default());
                        },
                        Element::TextSpace | Element::TextTab | Element::TextLineBreak
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
                        Element::UnknownAnimation if in_slide => {
                            return Err(Error::InvalidFormat(format!(
                                "unknown ODF animation element '{}'",
                                String::from_utf8_lossy(element.local_name().as_ref()),
                            )));
                        },
                        Element::Animation(kind)
                            if in_slide
                                && shape_stack.is_empty()
                                && current_paragraph.is_none() =>
                        {
                            if !kind.allowed_at_page_root() {
                                return Err(Error::InvalidFormat(
                                    "anim:param is only valid below anim:command".to_string(),
                                ));
                            }
                            current_animations.push(Self::parse_animation_node(
                                &mut reader,
                                element,
                                kind,
                                1,
                                &mut animation_node_count,
                            )?);
                            element_depth = Self::rewind_consumed_subtree(element_depth);
                        },
                        Element::SpreadsheetRoot if sheet_scope => {
                            spreadsheet_depth = Some(element_depth);
                        },
                        Element::Table
                            if sheet_scope
                                && shape_stack.is_empty()
                                && spreadsheet_depth
                                    .is_some_and(|depth| element_depth == depth + 1) =>
                        {
                            sheet_table_depth = Some(element_depth);
                            sheet_table_has_shapes = false;
                        },
                        Element::SheetShapes
                            if sheet_scope
                                && sheet_table_depth
                                    .is_some_and(|depth| element_depth == depth + 1) =>
                        {
                            if sheet_table_has_shapes {
                                return Err(Error::InvalidFormat(
                                    "table:table contains multiple table:shapes containers"
                                        .to_string(),
                                ));
                            }
                            sheet_table_has_shapes = true;
                            sheet_shapes_depth = Some(element_depth);
                            in_slide = true;
                        },
                        Element::Shape(shape_element) => {
                            let drawing_kind = Self::drawing_kind(shape_element);
                            shape_node_count =
                                shape_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat("ODP shape count overflow".to_string())
                                })?;
                            if shape_node_count > 65_536 {
                                return Err(Error::InvalidFormat(
                                    "ODP document exceeds 65536 shapes".to_string(),
                                ));
                            }
                            if shape_stack.len() >= 64 {
                                return Err(Error::InvalidFormat(
                                    "ODP shape groups exceed 64 levels".to_string(),
                                ));
                            }
                            let hyperlink_applies = current_hyperlink.is_some()
                                && hyperlink_parent_depth == Some(shape_stack.len());
                            if hyperlink_applies && hyperlink_shape_seen {
                                return Err(Error::InvalidFormat(
                                    "draw:a must wrap exactly one drawing shape".to_string(),
                                ));
                            }
                            if in_slide && shape_stack.is_empty() {
                                if drawing_kind.is_three_dimensional()
                                    && drawing_kind != DrawingShapeKind::ThreeDimensionalScene
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D drawing objects require a dr3d:scene parent"
                                            .to_string(),
                                    ));
                                }
                                if current_hyperlink.is_some() && !hyperlink_applies {
                                    return Err(Error::InvalidFormat(
                                        "misplaced draw:a presentation hyperlink".to_string(),
                                    ));
                                }
                                let mut builder =
                                    Self::shape_builder(&reader, element, shape_element)?;
                                if hyperlink_applies && let Some(hyperlink) = &current_hyperlink {
                                    builder.hyperlink = Some(hyperlink.clone());
                                    hyperlink_shape_seen = true;
                                }
                                shape_stack.push(builder);
                            } else if let Some(parent) = shape_stack.last() {
                                Self::validate_shape_parent(parent, drawing_kind)?;
                                if hyperlink_applies
                                    && parent.drawing_kind
                                        == Some(DrawingShapeKind::ThreeDimensionalScene)
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D scene children cannot be wrapped in draw:a".to_string(),
                                    ));
                                }
                                let mut builder =
                                    Self::shape_builder(&reader, element, shape_element)?;
                                if hyperlink_applies && let Some(hyperlink) = &current_hyperlink {
                                    builder.hyperlink = Some(hyperlink.clone());
                                    hyperlink_shape_seen = true;
                                }
                                shape_stack.push(builder);
                            }
                        },
                        Element::Image if !shape_stack.is_empty() => {
                            let builder = shape_stack.last_mut().expect("shape checked above");
                            builder.shape_type = ShapeType::Picture;
                            builder.image_href =
                                Self::get_attr(&reader, element, XLINK_NAMESPACE, b"href")?;
                        },
                        Element::Table if !shape_stack.is_empty() => {
                            shape_stack
                                .last_mut()
                                .expect("shape checked above")
                                .shape_type = ShapeType::Table;
                        },
                        Element::Object if !shape_stack.is_empty() => {
                            shape_stack
                                .last_mut()
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
                Event::Text(ref text) if in_media_plugin => {
                    let text = Self::decode_text(text)?;
                    if !text.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "draw:plugin cannot contain text".to_string(),
                        ));
                    }
                },
                Event::Text(ref text)
                    if shape_stack.last().is_some_and(|builder| {
                        builder.drawing_kind.is_some_and(|kind| {
                            kind.is_three_dimensional()
                                && kind != DrawingShapeKind::ThreeDimensionalScene
                        })
                    }) =>
                {
                    let text = Self::decode_text(text)?;
                    if !text.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "3D drawing elements cannot contain text".to_string(),
                        ));
                    }
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
                Event::CData(ref text) if in_media_plugin => {
                    let decoded = text.xml_content(XmlVersion::Explicit1_0).map_err(|error| {
                        Error::InvalidFormat(format!("invalid media plugin CDATA: {error}"))
                    })?;
                    if !decoded.trim().is_empty() {
                        return Err(Error::InvalidFormat(
                            "draw:plugin cannot contain text".to_string(),
                        ));
                    }
                },
                Event::GeneralRef(ref reference) if current_paragraph.is_some() => {
                    let text = Self::decode_reference(reference)?;
                    current_paragraph
                        .as_mut()
                        .expect("paragraph checked above")
                        .push_text(&text);
                },
                Event::GeneralRef(_) if in_media_plugin => {
                    return Err(Error::InvalidFormat(
                        "draw:plugin cannot contain character references".to_string(),
                    ));
                },
                Event::GeneralRef(_)
                    if shape_stack.last().is_some_and(|builder| {
                        builder.drawing_kind.is_some_and(|kind| {
                            kind.is_three_dimensional()
                                && kind != DrawingShapeKind::ThreeDimensionalScene
                        })
                    }) =>
                {
                    return Err(Error::InvalidFormat(
                        "3D drawing elements cannot contain character references".to_string(),
                    ));
                },
                Event::CData(ref data)
                    if shape_stack.last().is_some_and(|builder| {
                        builder.drawing_kind.is_some_and(|kind| {
                            kind.is_three_dimensional()
                                && kind != DrawingShapeKind::ThreeDimensionalScene
                        })
                    }) && !data.iter().all(u8::is_ascii_whitespace) =>
                {
                    return Err(Error::InvalidFormat(
                        "3D drawing elements cannot contain CDATA text".to_string(),
                    ));
                },
                Event::Empty(ref element) => {
                    let element_type = Self::classify(&namespace, element.local_name().as_ref());
                    Self::validate_three_dimensional_child_element(
                        shape_stack.last(),
                        element_type,
                    )?;
                    if in_media_parameter {
                        return Err(Error::InvalidFormat(
                            "draw:param cannot contain child elements".to_string(),
                        ));
                    }
                    if in_media_plugin && !matches!(element_type, Element::PluginParameter) {
                        return Err(Error::InvalidFormat(
                            "draw:plugin can only contain draw:param elements".to_string(),
                        ));
                    }
                    match element_type {
                        Element::Page if !sheet_scope && !in_slide => {
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
                                animations: Vec::new(),
                                legacy_animation: None,
                                shapes: Vec::new(),
                            });
                            slide_index += 1;
                        },
                        Element::EnhancedGeometry if !shape_stack.is_empty() => {
                            let builder = shape_stack.last_mut().expect("shape checked above");
                            if builder.drawing_kind != Some(DrawingShapeKind::CustomShape) {
                                return Err(Error::InvalidFormat(
                                    "draw:enhanced-geometry requires draw:custom-shape".to_string(),
                                ));
                            }
                            if builder.enhanced_geometry.is_some() {
                                return Err(Error::InvalidFormat(
                                    "draw:custom-shape contains multiple enhanced geometries"
                                        .to_string(),
                                ));
                            }
                            builder.enhanced_geometry = Some(EnhancedGeometry {
                                attributes: Self::exact_geometry_attributes(&reader, element)?,
                                children: Vec::new(),
                            });
                        },
                        Element::EnhancedGeometry
                        | Element::EnhancedEquation
                        | Element::EnhancedHandle
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "misplaced custom-shape enhanced geometry".to_string(),
                            ));
                        },
                        Element::Plugin => {
                            if let Some(builder) = shape_stack.last_mut() {
                                if !builder.is_frame {
                                    return Err(Error::InvalidFormat(
                                        "draw:plugin must be contained directly by draw:frame"
                                            .to_string(),
                                    ));
                                }
                                if builder.media.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "ODP frame contains multiple draw:plugin elements"
                                            .to_string(),
                                    ));
                                }
                                builder.shape_type = ShapeType::GraphicFrame;
                                builder.media = Some(Self::media_reference(&reader, element)?);
                            } else if in_slide {
                                return Err(Error::InvalidFormat(
                                    "draw:plugin must be contained by a drawing shape".to_string(),
                                ));
                            }
                        },
                        Element::DrawingHyperlink if in_slide => {
                            return Err(Error::InvalidFormat(
                                "draw:a must wrap exactly one non-empty drawing shape".to_string(),
                            ));
                        },
                        Element::EventListeners if !shape_stack.is_empty() => {
                            let builder = shape_stack.last_mut().expect("shape checked above");
                            if builder
                                .drawing_kind
                                .is_some_and(|kind| kind.is_three_dimensional())
                            {
                                return Err(Error::InvalidFormat(
                                    "3D shapes cannot contain presentation event listeners"
                                        .to_string(),
                                ));
                            }
                            if builder.event_listeners_seen {
                                return Err(Error::InvalidFormat(
                                    "ODP shape contains multiple office:event-listeners elements"
                                        .to_string(),
                                ));
                            }
                            builder.event_listeners_seen = true;
                        },
                        Element::EventListeners
                        | Element::EventListener
                        | Element::ScriptEventListener
                        | Element::Sound
                            if in_slide =>
                        {
                            return Err(Error::InvalidFormat(
                                "presentation event metadata must be contained by a shape's office:event-listeners"
                                    .to_string(),
                            ));
                        },
                        Element::PluginParameter
                            if in_media_plugin
                                && !in_media_parameter
                                && !shape_stack.is_empty() =>
                        {
                            shape_stack
                                .last_mut()
                                .and_then(|builder| builder.media.as_mut())
                                .expect("media plugin state checked above")
                                .add_parameter(Self::media_parameter(&reader, element)?)?;
                        },
                        _ if in_media_parameter => {
                            return Err(Error::InvalidFormat(
                                "draw:param cannot contain child elements".to_string(),
                            ));
                        },
                        _ if in_media_plugin => {
                            return Err(Error::InvalidFormat(
                                "draw:plugin can only contain draw:param elements".to_string(),
                            ));
                        },
                        Element::TextParagraph if in_slide => {
                            Self::push_parsed_paragraph(
                                "",
                                in_notes,
                                &mut current_notes_text,
                                &mut current_notes_has_paragraph,
                                shape_stack.last_mut(),
                                &mut current_slide_text,
                                &mut current_slide_has_segment,
                            );
                        },
                        Element::TextSpace | Element::TextTab | Element::TextLineBreak
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
                        Element::LegacyAnimation(kind) if in_slide => {
                            if kind != LegacyAnimationKind::Animations {
                                return Err(Error::InvalidFormat(
                                    "legacy presentation effects require a presentation:animations root"
                                        .to_string(),
                                ));
                            }
                            if current_legacy_animation.is_some() {
                                return Err(Error::InvalidFormat(
                                    "ODP slide contains multiple presentation:animations roots"
                                        .to_string(),
                                ));
                            }
                            legacy_animation_node_count =
                                legacy_animation_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "legacy ODP animation node count overflow".to_string(),
                                    )
                                })?;
                            let root = LegacyAnimationNode::from_parsed(
                                kind,
                                Self::animation_attributes(&reader, element)?,
                                Vec::new(),
                            );
                            validate_legacy_animation_root(&root)?;
                            current_legacy_animation = Some(root);
                        },
                        Element::UnknownAnimation if in_slide => {
                            return Err(Error::InvalidFormat(format!(
                                "unknown ODF animation element '{}'",
                                String::from_utf8_lossy(element.local_name().as_ref()),
                            )));
                        },
                        Element::Animation(kind)
                            if in_slide
                                && shape_stack.is_empty()
                                && current_paragraph.is_none() =>
                        {
                            if !kind.allowed_at_page_root() {
                                return Err(Error::InvalidFormat(
                                    "anim:param is only valid below anim:command".to_string(),
                                ));
                            }
                            animation_node_count =
                                animation_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat(
                                        "ODP animation node count overflow".to_string(),
                                    )
                                })?;
                            if animation_node_count > 65_536 {
                                return Err(Error::InvalidFormat(
                                    "ODP animation tree exceeds 65536 nodes".to_string(),
                                ));
                            }
                            current_animations.push(Node::from_parsed(
                                kind,
                                Self::animation_attributes(&reader, element)?,
                                Vec::new(),
                            ));
                        },
                        Element::Table
                            if sheet_scope
                                && shape_stack.is_empty()
                                && spreadsheet_depth
                                    .is_some_and(|depth| element_depth == depth) =>
                        {
                            slides.push(Slide {
                                title: None,
                                text: String::new(),
                                index: slide_index,
                                notes: None,
                                transition: None,
                                animations: Vec::new(),
                                legacy_animation: None,
                                shapes: Vec::new(),
                            });
                            slide_index += 1;
                        },
                        Element::SheetShapes
                            if sheet_scope
                                && sheet_table_depth
                                    .is_some_and(|depth| element_depth == depth) =>
                        {
                            if sheet_table_has_shapes {
                                return Err(Error::InvalidFormat(
                                    "table:table contains multiple table:shapes containers"
                                        .to_string(),
                                ));
                            }
                            sheet_table_has_shapes = true;
                        },
                        Element::Image => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::Picture;
                                builder.image_href =
                                    Self::get_attr(&reader, element, XLINK_NAMESPACE, b"href")?;
                            }
                        },
                        Element::Table => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::Table;
                            }
                        },
                        Element::Object => {
                            if let Some(builder) = shape_stack.last_mut() {
                                builder.shape_type = ShapeType::GraphicFrame;
                            }
                        },
                        Element::Shape(shape_element) if in_slide => {
                            let drawing_kind = Self::drawing_kind(shape_element);
                            shape_node_count =
                                shape_node_count.checked_add(1).ok_or_else(|| {
                                    Error::InvalidFormat("ODP shape count overflow".to_string())
                                })?;
                            if shape_node_count > 65_536 {
                                return Err(Error::InvalidFormat(
                                    "ODP document exceeds 65536 shapes".to_string(),
                                ));
                            }
                            let hyperlink_applies = current_hyperlink.is_some()
                                && hyperlink_parent_depth == Some(shape_stack.len());
                            if hyperlink_applies && hyperlink_shape_seen {
                                return Err(Error::InvalidFormat(
                                    "draw:a must wrap exactly one drawing shape".to_string(),
                                ));
                            }
                            let mut builder = Self::shape_builder(&reader, element, shape_element)?;
                            if hyperlink_applies && let Some(hyperlink) = &current_hyperlink {
                                builder.hyperlink = Some(hyperlink.clone());
                                hyperlink_shape_seen = true;
                            }
                            if let Some(parent) = shape_stack.last_mut() {
                                Self::validate_shape_parent(parent, drawing_kind)?;
                                if hyperlink_applies
                                    && parent.drawing_kind
                                        == Some(DrawingShapeKind::ThreeDimensionalScene)
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D scene children cannot be wrapped in draw:a".to_string(),
                                    ));
                                }
                                parent.children.push(builder.build());
                            } else {
                                if drawing_kind.is_three_dimensional()
                                    && drawing_kind != DrawingShapeKind::ThreeDimensionalScene
                                {
                                    return Err(Error::InvalidFormat(
                                        "3D drawing objects require a dr3d:scene parent"
                                            .to_string(),
                                    ));
                                }
                                Self::finish_shape(
                                    builder,
                                    &mut current_slide_title,
                                    &mut current_slide_text,
                                    &mut current_slide_has_segment,
                                    &mut current_shapes,
                                    retain_text_shapes,
                                );
                            }
                        },
                        _ => {},
                    }
                },
                Event::End(ref element) => {
                    element_depth = element_depth.saturating_sub(1);
                    let element_type = Self::classify(&namespace, element.local_name().as_ref());
                    if matches!(element_type, Element::TextParagraph) && current_paragraph.is_some()
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
                            shape_stack.last_mut(),
                            &mut current_slide_text,
                            &mut current_slide_has_segment,
                        );
                        buf.clear();
                        continue;
                    }
                    if matches!(element_type, Element::Notes) {
                        in_notes = false;
                        buf.clear();
                        continue;
                    }
                    if matches!(element_type, Element::Plugin) {
                        in_media_plugin = false;
                        buf.clear();
                        continue;
                    }
                    if matches!(element_type, Element::PluginParameter) && in_media_parameter {
                        in_media_parameter = false;
                        buf.clear();
                        continue;
                    }
                    if in_notes {
                        buf.clear();
                        continue;
                    }
                    match element_type {
                        Element::DrawingHyperlink if current_hyperlink.is_some() => {
                            if hyperlink_parent_depth != Some(shape_stack.len())
                                || !hyperlink_shape_seen
                            {
                                return Err(Error::InvalidFormat(
                                    "draw:a must wrap exactly one complete drawing shape"
                                        .to_string(),
                                ));
                            }
                            current_hyperlink = None;
                            hyperlink_parent_depth = None;
                            hyperlink_shape_seen = false;
                        },
                        Element::Page if !sheet_scope => {
                            if in_slide {
                                if current_hyperlink.is_some() {
                                    return Err(Error::InvalidFormat(
                                        "unterminated draw:a presentation hyperlink".to_string(),
                                    ));
                                }
                                slides.push(Slide {
                                    title: current_slide_title.take(),
                                    text: std::mem::take(&mut current_slide_text),
                                    index: slide_index,
                                    notes: (!current_notes_text.is_empty())
                                        .then(|| std::mem::take(&mut current_notes_text)),
                                    transition: current_transition.take(),
                                    animations: std::mem::take(&mut current_animations),
                                    legacy_animation: current_legacy_animation.take(),
                                    shapes: std::mem::take(&mut current_shapes),
                                });
                                slide_index += 1;
                            }
                            current_slide_has_segment = false;
                            current_notes_has_paragraph = false;
                            in_slide = false;
                        },
                        Element::SpreadsheetRoot
                            if sheet_scope
                                && spreadsheet_depth
                                    .is_some_and(|depth| element_depth + 1 == depth) =>
                        {
                            spreadsheet_depth = None;
                        },
                        Element::SheetShapes
                            if sheet_scope
                                && sheet_shapes_depth
                                    .is_some_and(|depth| element_depth + 1 == depth) =>
                        {
                            if current_hyperlink.is_some() {
                                return Err(Error::InvalidFormat(
                                    "unterminated draw:a drawing hyperlink".to_string(),
                                ));
                            }
                            sheet_shapes_depth = None;
                            in_slide = false;
                        },
                        Element::Table
                            if sheet_scope
                                && sheet_table_depth
                                    .is_some_and(|depth| element_depth + 1 == depth) =>
                        {
                            slides.push(Slide {
                                title: None,
                                text: std::mem::take(&mut current_slide_text),
                                index: slide_index,
                                notes: None,
                                transition: None,
                                animations: Vec::new(),
                                legacy_animation: None,
                                shapes: std::mem::take(&mut current_shapes),
                            });
                            slide_index += 1;
                            sheet_table_depth = None;
                            current_slide_has_segment = false;
                        },
                        Element::Shape(_) => {
                            if let Some(builder) = shape_stack.pop() {
                                if let Some(parent) = shape_stack.last_mut() {
                                    parent.children.push(builder.build());
                                    buf.clear();
                                    continue;
                                }
                                Self::finish_shape(
                                    builder,
                                    &mut current_slide_title,
                                    &mut current_slide_text,
                                    &mut current_slide_has_segment,
                                    &mut current_shapes,
                                    retain_text_shapes,
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

    #[test]
    fn preserves_drawing_element_kinds_and_unmodeled_geometry_attributes() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
            xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
            xmlns:r="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0">
            <o:body><o:presentation><d:page>
              <d:rect d:name="rect" d:corner-radius="2mm"/>
              <d:ellipse d:name="ellipse" d:kind="section" d:start-angle="30" s:rx="2cm" s:ry="1cm"/>
              <d:circle d:name="circle" s:cx="3cm" s:cy="4cm" s:r="2cm"/>
              <d:path d:name="path" s:viewBox="0 0 100 100" s:d="M 0 0 L 100 100"/>
              <d:polygon d:name="polygon" s:viewBox="0 0 10 10" d:points="0,0 10,0 5,10"/>
              <d:polyline d:name="polyline" d:points="0,0 5,5 10,0"/>
              <d:regular-polygon d:name="regular" d:corners="7" d:concave="true" d:sharpness="25%"/>
              <d:page-thumbnail d:name="thumb" d:page-number="2"/>
              <d:measure d:name="measure" s:x1="1cm" s:y1="2cm" s:x2="3cm" s:y2="4cm"/>
              <d:caption d:name="caption" d:caption-point-x="1cm" d:caption-point-y="2cm"/>
              <d:connector d:name="connector" d:type="curve" d:start-shape="a" d:end-shape="b" s:x1="0cm" s:y1="0cm" s:x2="1cm" s:y2="1cm"/>
              <d:control d:name="control" d:control="control1"/>
              <d:custom-shape d:name="custom" d:engine="vendor" d:data="opaque" r:transform="rotatex(0.5)">
                <d:enhanced-geometry d:type="non-primitive" s:viewBox="0 0 21600 21600"
                  d:modifiers="10800" d:enhanced-path="M 0 0 L ?f0 21600 Z" r:projection="perspective">
                  <d:equation d:name="f0" d:formula="$0 * 2 &amp; 21600"/>
                  <d:handle d:handle-position="$0 10800" d:handle-range-x-minimum="0" d:handle-range-x-maximum="21600"/>
                </d:enhanced-geometry>
              </d:custom-shape>
            </d:page></o:presentation></o:body>
        </o:document-content>"#;
        let slides = Parser::parse_slides(xml).unwrap();
        let shapes = &slides[0].shapes;
        let expected = [
            DrawingShapeKind::Rectangle,
            DrawingShapeKind::Ellipse,
            DrawingShapeKind::Circle,
            DrawingShapeKind::Path,
            DrawingShapeKind::Polygon,
            DrawingShapeKind::Polyline,
            DrawingShapeKind::RegularPolygon,
            DrawingShapeKind::PageThumbnail,
            DrawingShapeKind::Measure,
            DrawingShapeKind::Caption,
            DrawingShapeKind::Connector,
            DrawingShapeKind::Control,
            DrawingShapeKind::CustomShape,
        ];
        assert_eq!(shapes.len(), expected.len());
        for (index, (shape, expected_kind)) in shapes.iter().zip(expected).enumerate() {
            assert_eq!(shape.drawing_kind(), Some(expected_kind));
            let regenerated = crate::Builder::generate_shape_xml(shape, index).unwrap();
            assert!(regenerated.starts_with(&format!("<{}", expected_kind.element_name())));
            assert!(!regenerated.contains("draw:layer="));
        }
        let ellipse = crate::Builder::generate_shape_xml(&shapes[1], 1).unwrap();
        assert!(ellipse.contains(r#"draw:kind="section""#));
        assert!(ellipse.contains(r#"draw:start-angle="30""#));
        assert!(ellipse.contains(r#"svg:rx="2cm""#));
        let path = crate::Builder::generate_shape_xml(&shapes[3], 3).unwrap();
        assert!(path.contains(r#"svg:viewBox="0 0 100 100""#));
        assert!(path.contains(r#"svg:d="M 0 0 L 100 100""#));
        let connector = crate::Builder::generate_shape_xml(&shapes[10], 10).unwrap();
        assert!(connector.contains(r#"draw:type="curve""#));
        assert!(connector.contains(r#"draw:start-shape="a""#));
        let custom = &shapes[12];
        assert!(custom.drawing_attributes().iter().any(|attribute| {
            attribute.namespace() == DrawingAttributeNamespace::Dr3d
                && attribute.local_name() == "transform"
                && attribute.value() == "rotatex(0.5)"
        }));
        let geometry = custom.enhanced_geometry().unwrap();
        assert_eq!(geometry.children().len(), 2);
        assert_eq!(
            geometry.children()[0].kind(),
            EnhancedGeometryChildKind::Equation
        );
        assert_eq!(
            geometry.children()[1].kind(),
            EnhancedGeometryChildKind::Handle
        );
        let regenerated = crate::Builder::generate_shape_xml(custom, 12).unwrap();
        assert!(regenerated.contains("<draw:enhanced-geometry"));
        assert!(regenerated.contains(r#"dr3d:projection="perspective""#));
        assert!(regenerated.contains(r#"draw:formula="$0 * 2 &amp; 21600""#));
        assert!(regenerated.contains(r#"draw:handle-position="$0 10800""#));
    }

    #[test]
    fn preserves_recursive_inert_three_dimensional_scenes() {
        let xml = r##"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
            xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
            xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
            xmlns:r="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0">
            <o:body><o:presentation><d:page>
              <r:scene s:x="1cm" s:y="2cm" s:width="8cm" s:height="6cm"
                r:vrp="(0 0 10)" r:projection="perspective" r:ambient-color="#112233">
                <r:light r:direction="(0 0 -1)" r:diffuse-color="#ffffff" r:enabled="true"/>
                <r:cube r:min-edge="(-1 -1 -1)" r:max-edge="(1 1 1)" r:transform="rotatex(0.5)"/>
                <r:scene r:shade-mode="phong">
                  <r:sphere r:center="(0 0 0)" r:size="(2 2 2)"/>
                  <r:extrude s:viewBox="0 0 10 10" s:d="M 0 0 L 10 10"/>
                  <r:rotate s:viewBox="0 0 20 20" s:d="M 1 1 L 19 19"/>
                </r:scene>
              </r:scene>
            </d:page></o:presentation></o:body>
        </o:document-content>"##;
        let slides = Parser::parse_slides(xml).unwrap();
        let scene = &slides[0].shapes[0];
        assert_eq!(
            scene.drawing_kind(),
            Some(DrawingShapeKind::ThreeDimensionalScene)
        );
        assert_eq!(scene.x.as_deref(), Some("1cm"));
        assert_eq!(scene.children.len(), 3);
        assert_eq!(
            scene.children[0].drawing_kind(),
            Some(DrawingShapeKind::ThreeDimensionalLight)
        );
        assert_eq!(
            scene.children[1].drawing_kind(),
            Some(DrawingShapeKind::ThreeDimensionalCube)
        );
        let nested = &scene.children[2];
        assert_eq!(
            nested.drawing_kind(),
            Some(DrawingShapeKind::ThreeDimensionalScene)
        );
        assert_eq!(nested.children.len(), 3);
        assert_eq!(
            nested.children[0].drawing_kind(),
            Some(DrawingShapeKind::ThreeDimensionalSphere)
        );
        assert_eq!(
            nested.children[1].drawing_kind(),
            Some(DrawingShapeKind::ThreeDimensionalExtrude)
        );
        assert_eq!(
            nested.children[2].drawing_kind(),
            Some(DrawingShapeKind::ThreeDimensionalRotate)
        );
        let regenerated = crate::Builder::generate_shape_xml(scene, 0).unwrap();
        assert!(regenerated.starts_with("<dr3d:scene"));
        assert!(regenerated.contains(r#"dr3d:projection="perspective""#));
        assert!(regenerated.contains(r#"dr3d:direction="(0 0 -1)""#));
        assert!(regenerated.contains(r#"<dr3d:cube"#));
        assert!(regenerated.contains(r#"<dr3d:sphere"#));
        assert!(regenerated.contains(r#"svg:d="M 0 0 L 10 10""#));
    }

    #[test]
    fn rejects_invalid_three_dimensional_shape_hierarchies() {
        for body in [
            r#"<r:cube/>"#,
            r#"<d:g><r:sphere/></d:g>"#,
            r#"<r:scene><d:rect/></r:scene>"#,
            r#"<r:scene><r:cube/><r:light r:direction="(0 0 -1)"/></r:scene>"#,
            r#"<r:scene><r:cube>not empty</r:cube></r:scene>"#,
            r#"<r:scene><r:cube><d:glue-point/></r:cube></r:scene>"#,
            r#"<r:scene><r:light/></r:scene>"#,
            r#"<r:scene><r:extrude s:d="M 0 0"/></r:scene>"#,
            r#"<r:scene><r:rotate s:viewBox="0 0 10 10"/></r:scene>"#,
        ] {
            let xml = format!(
                r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:r="urn:oasis:names:tc:opendocument:xmlns:dr3d:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><o:body><o:presentation><d:page>{body}</d:page></o:presentation></o:body></o:document-content>"#
            );
            assert!(Parser::parse_slides(&xml).is_err(), "accepted {body}");
        }
    }

    #[test]
    fn rejects_misplaced_or_invalid_enhanced_geometry() {
        for shape in [
            "<d:rect><d:enhanced-geometry/></d:rect>",
            "<d:custom-shape><d:enhanced-geometry/><d:enhanced-geometry/></d:custom-shape>",
            "<d:custom-shape><d:enhanced-geometry><d:handle/><d:equation/></d:enhanced-geometry></d:custom-shape>",
            "<d:custom-shape><d:equation/></d:custom-shape>",
        ] {
            let xml = format!(
                r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:presentation><d:page>{shape}</d:page></o:presentation></o:body></o:document-content>"#
            );
            assert!(Parser::parse_slides(&xml).is_err(), "accepted {shape}");
        }
    }

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
        let slides = Parser::parse_slides(TEST_PRESENTATION_XML).unwrap();
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
        let slides = Parser::parse_slides(TEST_EMPTY_PRESENTATION).unwrap();
        assert!(slides.is_empty());
    }

    #[test]
    fn test_parse_shapes() {
        let slides = Parser::parse_slides(TEST_SHAPES_XML).unwrap();
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
            animations: vec![],
            legacy_animation: None,
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
            animations: vec![],
            legacy_animation: None,
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
            ..Shape::new()
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
            ..Shape::new()
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
                ..Shape::new()
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

        let slides = Parser::parse_slides(xml).unwrap();
        let shape = &slides[0].shapes[0];
        assert_eq!(shape.shape_type, ShapeType::Picture);
        assert_eq!(shape.image_href(), Some("Pictures/a&b.png"));
    }

    #[test]
    fn preserves_shape_stacking_transform_and_presentation_role() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><office:body><office:presentation><draw:page><draw:frame draw:name="Chart" draw:layer="controls" draw:z-index="184467440737095516160" draw:transform="rotate (0.5) translate (1cm 2cm)" presentation:class="chart" presentation:placeholder="true" presentation:user-transformed="false"/></draw:page></office:presentation></office:body></office:document-content>"#;

        let slides = Parser::parse_slides(xml).unwrap();
        let shape = &slides[0].shapes[0];
        assert_eq!(shape.shape_type, ShapeType::Placeholder);
        assert_eq!(shape.layer(), Some("controls"));
        assert_eq!(shape.z_index(), Some("184467440737095516160"));
        assert_eq!(shape.transform(), Some("rotate (0.5) translate (1cm 2cm)"));
        assert_eq!(shape.presentation_class(), Some("chart"));
        assert_eq!(shape.presentation_placeholder, Some(true));
        assert_eq!(shape.presentation_user_transformed, Some(false));
    }

    #[test]
    fn rejects_invalid_shape_stacking_and_boolean_values() {
        for attribute in [
            r#"draw:z-index="-1""#,
            r#"draw:z-index="1.5""#,
            r#"presentation:placeholder="yes""#,
            r#"presentation:user-transformed="no""#,
        ] {
            let xml = format!(
                r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"><office:body><office:presentation><draw:page><draw:frame {attribute}/></draw:page></office:presentation></office:body></office:document-content>"#
            );
            assert!(Parser::parse_slides(&xml).is_err(), "accepted {attribute}");
        }
    }

    #[test]
    fn preserves_shape_groups_and_identifies_opaque_frames() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:xlink="http://www.w3.org/1999/xlink"><office:body><office:presentation><draw:page><draw:g draw:name="Group"><draw:a xlink:href="https://example.test/group-child" xlink:type="simple"><draw:rect/></draw:a></draw:g><draw:frame draw:name="Table"><table:table/></draw:frame><draw:frame draw:name="Object"><draw:object/></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

        let slides = Parser::parse_slides(xml).unwrap();
        let types: Vec<_> = slides[0]
            .shapes
            .iter()
            .map(|shape| shape.shape_type)
            .collect();
        assert_eq!(
            types,
            [ShapeType::Group, ShapeType::Table, ShapeType::GraphicFrame]
        );
        let group = &slides[0].shapes[0];
        assert_eq!(group.children().len(), 1);
        assert_eq!(
            group.children()[0].drawing_kind(),
            Some(DrawingShapeKind::Rectangle)
        );
        assert_eq!(
            group.children()[0].hyperlink().map(DrawingHyperlink::href),
            Some("https://example.test/group-child")
        );
        let regenerated = crate::Builder::generate_shape_xml(group, 0).unwrap();
        assert!(regenerated.starts_with(r#"<draw:g draw:name="Group">"#));
        assert!(regenerated.contains("<draw:rect"));
        assert!(regenerated.contains("<draw:a"));
    }

    #[test]
    fn bounds_nested_shape_group_depth() {
        let mut xml = String::from(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"><o:body><o:presentation><d:page>"#,
        );
        for _ in 0..65 {
            xml.push_str("<d:g>");
        }
        for _ in 0..65 {
            xml.push_str("</d:g>");
        }
        xml.push_str("</d:page></o:presentation></o:body></o:document-content>");
        let error = Parser::parse_slides(&xml).unwrap_err();
        assert!(error.to_string().contains("64 levels"));
    }

    #[test]
    fn preserves_text_across_spans_and_odf_whitespace_elements() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:presentation><draw:page><draw:frame presentation:class="object"><draw:text-box><text:p><text:s/>Hel<text:span>lo</text:span> <text:span>world</text:span><text:s text:c="2"/>again<text:tab/>tab<text:line-break/>line &amp; more</text:p><text:p/><text:p>second paragraph<text:s/></text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

        let slides = Parser::parse_slides(xml).unwrap();
        assert_eq!(
            slides[0].text,
            " Hello world  again\ttab\nline & more\n\nsecond paragraph "
        );
    }

    #[test]
    fn rejects_excessive_explicit_space_expansion() {
        let xml = r#"<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"><office:body><office:presentation><draw:page><draw:frame><draw:text-box><text:p>x<text:s text:c="1000001"/></text:p></draw:text-box></draw:frame></draw:page></office:presentation></office:body></office:document-content>"#;

        let error = Parser::parse_slides(xml).unwrap_err();
        assert!(error.to_string().contains("safety limit"));
    }

    #[test]
    fn parses_arbitrary_odf_namespace_prefixes() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:t="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:l="http://www.w3.org/1999/xlink" xmlns:tb="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:f="urn:example:wrong"><o:body><o:presentation><f:page><t:p>ignored</t:p></f:page><d:page><d:frame d:name="Aliased Title" p:class="title" s:x="1cm"><d:text-box><t:p>Aliased<t:s/>title</t:p></d:text-box></d:frame><d:frame d:name="Picture"><d:image l:href="Pictures/a&amp;b.png"/></d:frame><d:connector d:name="Link" s:x1="1cm" s:y1="2cm" s:x2="3cm" s:y2="4cm"/><d:frame d:name="Table"><tb:table/></d:frame><p:notes><d:frame><d:text-box><t:p>Aliased note</t:p></d:text-box></d:frame></p:notes></d:page></o:presentation></o:body></o:document-content>"#;

        let slides = Parser::parse_slides(xml).unwrap();
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

        let slides = Parser::parse_slides_with_styles(content, Some(styles)).unwrap();
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
        let error = Parser::parse_slides_with_styles(content, None).unwrap_err();
        assert!(error.to_string().contains("cyclic"));
    }

    #[test]
    fn parses_complete_namespace_aware_animation_trees() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:a="urn:oasis:names:tc:opendocument:xmlns:animation:1.0" xmlns:m="urn:oasis:names:tc:opendocument:xmlns:smil-compatible:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" xmlns:l="http://www.w3.org/1999/xlink" xmlns:e="urn:example:animation-extension" xmlns:f="urn:example:not-animation"><o:body><o:presentation><d:page><f:par/><a:par m:begin="slide.begin+1s" p:node-type="timing-root" e:flag="keep &amp; roundtrip"><a:animate a:formula="x+1" m:targetElement="shape1"/><a:animateColor a:color-interpolation="rgb"/><a:animateMotion s:path="M 0 0 L 1 1"/><a:animateTransform s:type="rotate"/><a:audio l:href="Sounds/chime.ogg" xml:id="audio1"/><a:command a:command="show"><a:param a:name="page" a:value="2"/></a:command><a:iterate a:iterate-type="by-paragraph"><a:set m:to="visible"/></a:iterate><a:par/><a:seq><a:transitionFilter m:type="fade"/></a:seq><a:set m:attributeName="visibility"/><a:transitionFilter m:subtype="crossfade"/></a:par></d:page></o:presentation></o:body></o:document-content>"#;

        let slides = Parser::parse_slides(xml).unwrap();
        assert_eq!(slides.len(), 1);
        assert_eq!(slides[0].animations.len(), 1);
        let root = &slides[0].animations[0];
        assert_eq!(root.kind(), Kind::Parallel);
        assert_eq!(root.children().len(), 11);
        assert_eq!(
            root.attribute(&Namespace::Smil, "begin"),
            Some("slide.begin+1s")
        );
        assert_eq!(
            root.attribute(
                &Namespace::Other("urn:example:animation-extension".to_string()),
                "flag"
            ),
            Some("keep & roundtrip")
        );
        assert_eq!(root.children()[0].kind(), Kind::Animate);
        assert_eq!(root.children()[4].kind(), Kind::Audio);
        let command = &root.children()[5];
        assert_eq!(command.kind(), Kind::Command);
        assert_eq!(command.children()[0].kind(), Kind::Parameter);
        assert_eq!(
            command.children()[0].attribute(&Namespace::Animation, "value"),
            Some("2")
        );
        assert_eq!(root.children()[6].children()[0].kind(), Kind::Set);
        assert_eq!(
            root.children()[8].children()[0].kind(),
            Kind::TransitionFilter
        );
    }

    #[test]
    fn rejects_invalid_animation_structure() {
        let invalid_trees = [
            "<a:animate><a:set/></a:animate>",
            "<a:command><a:animate/></a:command>",
            "<a:param a:name=\"orphan\"/>",
            "<a:notInOdf/>",
            "<a:par>not whitespace</a:par>",
        ];
        for tree in invalid_trees {
            let xml = format!(
                r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:a="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"><o:body><o:presentation><d:page>{tree}</d:page></o:presentation></o:body></o:document-content>"#
            );
            assert!(Parser::parse_slides(&xml).is_err(), "accepted {tree}");
        }
    }

    #[test]
    fn bounds_animation_nesting() {
        let mut tree = "<a:par>".repeat(129);
        tree.push_str("<a:set/>");
        tree.push_str(&"</a:par>".repeat(129));
        let xml = format!(
            r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:a="urn:oasis:names:tc:opendocument:xmlns:animation:1.0"><o:body><o:presentation><d:page>{tree}</d:page></o:presentation></o:body></o:document-content>"#
        );

        let error = Parser::parse_slides(&xml).unwrap_err();
        assert!(error.to_string().contains("128 levels"));
    }

    #[test]
    fn parses_inert_media_plugins_and_parameters() {
        let xml = r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><o:body><o:presentation><d:page><d:frame d:name="Movie" s:x="1cm" s:y="2cm"><d:plugin x:href="Media/a&amp;b.mp4" x:type="simple" d:mime-type="video/mp4" x:show="embed" x:actuate="onRequest" xml:id="movie1"><d:param d:name="autoplay" d:value="false"/><d:param d:name="caption" d:value="A &amp; B"> </d:param></d:plugin></d:frame></d:page></o:presentation></o:body></o:document-content>"#;

        let slides = Parser::parse_slides(xml).unwrap();
        let shape = &slides[0].shapes[0];
        assert_eq!(shape.shape_type, ShapeType::GraphicFrame);
        let media = shape.media().unwrap();
        assert_eq!(media.href(), "Media/a&b.mp4");
        assert_eq!(media.mime_type(), Some("video/mp4"));
        assert_eq!(media.show(), Some(Show::Embed));
        assert_eq!(media.actuate(), Some(Actuate::OnRequest));
        assert_eq!(media.xml_id(), Some("movie1"));
        assert_eq!(media.parameters().len(), 2);
        assert_eq!(media.parameters()[0].name(), "autoplay");
        assert_eq!(media.parameters()[1].value(), "A & B");
    }

    #[test]
    fn rejects_schema_invalid_media_plugins() {
        let invalid_plugins = [
            r#"<d:frame><d:plugin x:type="simple"/></d:frame>"#,
            r#"<d:frame><d:plugin x:href="a.mp4"/></d:frame>"#,
            r#"<d:frame><d:plugin x:href="a.mp4" x:type="extended"/></d:frame>"#,
            r#"<d:frame><d:plugin x:href="a.mp4" x:type="simple" x:show="invalid"/></d:frame>"#,
            r#"<d:plugin x:href="a.mp4" x:type="simple"/>"#,
            r#"<d:rect><d:plugin x:href="a.mp4" x:type="simple"/></d:rect>"#,
            r#"<d:frame><d:plugin x:href="a.mp4" x:type="simple">text</d:plugin></d:frame>"#,
            r#"<d:frame><d:plugin x:href="a.mp4" x:type="simple"><d:param d:name="x"/></d:plugin></d:frame>"#,
            r#"<d:frame><d:plugin x:href="a.mp4" x:type="simple"><d:param d:name="x" d:value="y"><d:param d:name="nested" d:value="z"/></d:param></d:plugin></d:frame>"#,
        ];
        for plugin in invalid_plugins {
            let xml = format!(
                r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:presentation><d:page>{plugin}</d:page></o:presentation></o:body></o:document-content>"#
            );
            assert!(Parser::parse_slides(&xml).is_err(), "accepted {plugin}");
        }
    }

    #[test]
    fn parses_shape_hyperlinks_and_inert_event_bindings() {
        let xml = r##"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:sc="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:s="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"><o:body><o:presentation><d:page><d:a x:type="simple" x:href="#page2" x:actuate="onRequest" x:show="replace" o:target-frame-name="_self" o:name="jump" o:title="Jump &amp; return" o:server-map="true" xml:id="link1"><d:rect d:name="Action" s:x="1cm"><o:event-listeners><sc:event-listener sc:event-name="dom:click" sc:language="ooo:script" sc:macro-name="Standard.Module1.Main"/><sc:event-listener sc:event-name="dom:mouseover" sc:language="javascript" x:type="simple" x:href="Scripts/hover.js" x:actuate="onRequest"/><p:event-listener sc:event-name="dom:click" p:action="show" p:effect="fade" p:direction="from-left" p:speed="fast" p:start-scale="50%" x:type="simple" x:href="#page3" x:show="embed" x:actuate="onRequest" p:verb="2"><p:sound x:type="simple" x:href="Sounds/click.ogg" x:actuate="onRequest" x:show="replace" p:play-full="true" xml:id="sound1"/></p:event-listener></o:event-listeners></d:rect></d:a></d:page></o:presentation></o:body></o:document-content>"##;

        let slides = Parser::parse_slides(xml).unwrap();
        let shape = &slides[0].shapes[0];
        let hyperlink = shape.hyperlink().unwrap();
        assert_eq!(hyperlink.href(), "#page2");
        assert!(hyperlink.actuate_on_request());
        assert_eq!(hyperlink.show(), Some(HyperlinkShow::Replace));
        assert_eq!(hyperlink.target_frame_name(), Some("_self"));
        assert_eq!(hyperlink.title(), Some("Jump & return"));
        assert_eq!(hyperlink.server_map(), Some(true));
        assert_eq!(hyperlink.xml_id(), Some("link1"));

        assert_eq!(shape.event_listeners().len(), 3);
        let ShapeEventListener::Script(macro_listener) = &shape.event_listeners()[0] else {
            panic!("expected script listener");
        };
        assert_eq!(
            macro_listener.macro_name.as_deref(),
            Some("Standard.Module1.Main")
        );
        let ShapeEventListener::Action(action) = &shape.event_listeners()[2] else {
            panic!("expected presentation listener");
        };
        assert_eq!(action.action, Action::Show);
        assert_eq!(action.effect.as_ref().unwrap().as_str(), "fade");
        assert_eq!(action.direction.as_ref().unwrap().as_str(), "from-left");
        assert_eq!(action.speed, Some(TransitionSpeed::Fast));
        assert_eq!(action.start_scale.as_deref(), Some("50%"));
        assert_eq!(action.verb, Some(2));
        assert_eq!(action.sound.as_ref().unwrap().href, "Sounds/click.ogg");
    }

    #[test]
    fn rejects_invalid_shape_hyperlinks_and_event_bindings() {
        let invalid = [
            r##"<d:a x:href="#p"><d:rect/></d:a>"##,
            r##"<d:a x:type="simple" x:href="#p"/>"##,
            r##"<d:a x:type="simple" x:href="#p"><d:rect/><d:rect/></d:a>"##,
            r#"<p:event-listener sc:event-name="dom:click" p:action="next-page"/>"#,
            r#"<d:rect><o:event-listeners><sc:event-listener sc:event-name="dom:click" sc:language="ooo:script" sc:macro-name="M" x:type="simple" x:href="S"/></o:event-listeners></d:rect>"#,
            r#"<d:rect><o:event-listeners><p:event-listener sc:event-name="dom:click" p:action="invalid"/></o:event-listeners></d:rect>"#,
            r#"<d:rect><o:event-listeners><p:event-listener sc:event-name="dom:click" p:action="sound"><p:sound x:href="a" x:type="extended"/></p:event-listener></o:event-listeners></d:rect>"#,
            r#"<d:rect><o:event-listeners/><o:event-listeners/></d:rect>"#,
        ];
        for fragment in invalid {
            let xml = format!(
                r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:sc="urn:oasis:names:tc:opendocument:xmlns:script:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:presentation><d:page>{fragment}</d:page></o:presentation></o:body></o:document-content>"#
            );
            assert!(Parser::parse_slides(&xml).is_err(), "accepted {fragment}");
        }
    }

    #[test]
    fn parses_legacy_presentation_effect_trees() {
        let xml = r##"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:x="http://www.w3.org/1999/xlink" xmlns:e="urn:example:effects"><o:body><o:presentation><d:page><p:animations e:mode="legacy"><p:animation-group><p:show-shape d:shape-id="shape1" p:effect="fade" p:speed="fast"><p:sound x:href="Sounds/a&amp;b.ogg" x:type="simple" p:play-full="true"/></p:show-shape><p:dim d:shape-id="shape1" d:color="#808080"/><p:hide-text d:shape-id="shape2"/><p:play d:shape-id="movie1"/></p:animation-group></p:animations></d:page></o:presentation></o:body></o:document-content>"##;

        let slides = Parser::parse_slides(xml).unwrap();
        let root = slides[0].legacy_animation().unwrap();
        assert_eq!(root.kind(), LegacyAnimationKind::Animations);
        assert_eq!(
            root.attribute(&Namespace::Other("urn:example:effects".to_string()), "mode"),
            Some("legacy")
        );
        let group = &root.children()[0];
        assert_eq!(group.kind(), LegacyAnimationKind::Group);
        assert_eq!(group.children().len(), 4);
        let show = &group.children()[0];
        assert_eq!(show.kind(), LegacyAnimationKind::ShowShape);
        assert_eq!(show.attribute(&Namespace::Draw, "shape-id"), Some("shape1"));
        assert_eq!(show.children()[0].kind(), LegacyAnimationKind::Sound);
        assert_eq!(
            show.children()[0].attribute(&Namespace::Xlink, "href"),
            Some("Sounds/a&b.ogg")
        );
    }

    #[test]
    fn rejects_invalid_legacy_presentation_effects() {
        let invalid = [
            r#"<p:show-shape d:shape-id="orphan"/>"#,
            r#"<p:animations><p:show-shape/></p:animations>"#,
            r#"<p:animations><p:dim d:shape-id="s"/></p:animations>"#,
            r#"<p:animations><p:play d:shape-id="s"><p:sound x:href="a" x:type="simple"/></p:play></p:animations>"#,
            r#"<p:animations><p:show-shape d:shape-id="s"><p:sound x:href="a" x:type="extended"/></p:show-shape></p:animations>"#,
            r#"<p:animations>text</p:animations>"#,
        ];
        for effects in invalid {
            let xml = format!(
                r#"<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:d="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:p="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0" xmlns:x="http://www.w3.org/1999/xlink"><o:body><o:presentation><d:page>{effects}</d:page></o:presentation></o:body></o:document-content>"#
            );
            assert!(Parser::parse_slides(&xml).is_err(), "accepted {effects}");
        }
    }
}
