//! Bounded, inert discovery of PowerPoint OLE object shapes.
//!
//! This module returns only stored PresentationML and OPC metadata. It never
//! parses, opens, activates, renders, executes, or otherwise inspects embedded
//! object or package payload bytes.

use crate::common::xml::{is_drawingml_name, unqualified_attribute_value};
use crate::common::{MceCapabilities, MceLimits, process_markup_compatibility};
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::{is_presentationml_name, relationship_attribute_value};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Part};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_TOTAL_SLIDE_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_OLE_OBJECTS: usize = 4_096;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;
const MAX_XML_ATTRIBUTES: usize = 64;
const MAX_ATTRIBUTE_BYTES: usize = 4_096;
const OLE_GRAPHIC_DATA_URI: &str = "http://schemas.openxmlformats.org/presentationml/2006/ole";

/// Whether an OLE object shape stores an embedded object or a declared link.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PptxOleObjectMode {
    /// The shape contains a p:embed element.
    Embedded,
    /// The shape contains a p:link element.
    Linked,
}

/// The declared OPC payload family for an OLE object shape.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum PptxOlePayloadKind {
    /// An OOXML Embedded Object part.
    OleObject,
    /// An OOXML Embedded Package part.
    Package,
}

/// An inert target declared by an OLE object shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub enum PptxOleObjectTarget {
    /// An internal OPC payload part.
    Internal {
        /// Absolute package part name.
        part_name: PackURI,
        /// Declared OPC content type.
        content_type: String,
        /// Declared relationship type URI.
        relationship_type: String,
    },
    /// An external target retained as stored metadata.
    External {
        /// Declared target URI or path.
        target: String,
        /// Declared relationship type URI.
        relationship_type: String,
    },
}

impl PptxOleObjectTarget {
    /// Return the declared relationship type URI.
    #[inline]
    pub fn relationship_type(&self) -> &str {
        match self {
            Self::Internal {
                relationship_type, ..
            }
            | Self::External {
                relationship_type, ..
            } => relationship_type,
        }
    }

    /// Return the target part name for an internal relationship.
    #[inline]
    pub fn part_name(&self) -> Option<&PackURI> {
        match self {
            Self::Internal { part_name, .. } => Some(part_name),
            Self::External { .. } => None,
        }
    }

    /// Return the declared content type for an internal relationship.
    #[inline]
    pub fn content_type(&self) -> Option<&str> {
        match self {
            Self::Internal { content_type, .. } => Some(content_type),
            Self::External { .. } => None,
        }
    }

    /// Return the stored target string for an external relationship.
    #[inline]
    pub fn external_target(&self) -> Option<&str> {
        match self {
            Self::Internal { .. } => None,
            Self::External { target, .. } => Some(target),
        }
    }
}

/// A bounded, inert inventory record for one PowerPoint OLE object shape.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxOleObject {
    slide_index: usize,
    object_index: usize,
    shape_id: Option<u32>,
    shape_name: Option<String>,
    legacy_shape_id: Option<String>,
    name: Option<String>,
    program_id: Option<String>,
    show_as_icon: Option<bool>,
    preview_width: Option<u32>,
    preview_height: Option<u32>,
    mode: PptxOleObjectMode,
    relationship_id: Option<String>,
    payload_kind: Option<PptxOlePayloadKind>,
    target: Option<PptxOleObjectTarget>,
    preview_relationship_id: Option<String>,
}

impl PptxOleObject {
    /// Return the zero-based index of the slide that owns this object.
    #[inline]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Return this object's zero-based source-order index on its slide.
    #[inline]
    pub fn object_index(&self) -> usize {
        self.object_index
    }

    /// Return the graphic-frame shape ID, when stored.
    #[inline]
    pub fn shape_id(&self) -> Option<u32> {
        self.shape_id
    }

    /// Return the graphic-frame shape name, when stored.
    #[inline]
    pub fn shape_name(&self) -> Option<&str> {
        self.shape_name.as_deref()
    }

    /// Return the legacy VML-associated shape ID from the p:oleObj spid attribute.
    #[inline]
    pub fn legacy_shape_id(&self) -> Option<&str> {
        self.legacy_shape_id.as_deref()
    }

    /// Return the OLE object name, when stored.
    #[inline]
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Return the OLE ProgID, when stored.
    #[inline]
    pub fn program_id(&self) -> Option<&str> {
        self.program_id.as_deref()
    }

    /// Return whether the object is stored as an icon, when specified.
    #[inline]
    pub fn show_as_icon(&self) -> Option<bool> {
        self.show_as_icon
    }

    /// Return the stored preview-image width, when specified.
    #[inline]
    pub fn preview_width(&self) -> Option<u32> {
        self.preview_width
    }

    /// Return the stored preview-image height, when specified.
    #[inline]
    pub fn preview_height(&self) -> Option<u32> {
        self.preview_height
    }

    /// Return whether the shape uses p:embed or p:link.
    #[inline]
    pub fn mode(&self) -> PptxOleObjectMode {
        self.mode
    }

    /// Return the optional payload relationship ID from the owning slide.
    #[inline]
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }

    /// Return the payload family inferred from the declared relationship type.
    #[inline]
    pub fn payload_kind(&self) -> Option<PptxOlePayloadKind> {
        self.payload_kind
    }

    /// Return the declared payload target, when a relationship ID is present.
    #[inline]
    pub fn target(&self) -> Option<&PptxOleObjectTarget> {
        self.target.as_ref()
    }

    /// Return the optional relationship ID for the icon or preview image.
    ///
    /// Preview-image relationships are retained only as inert identifiers and
    /// are never loaded or rendered.
    #[inline]
    pub fn preview_relationship_id(&self) -> Option<&str> {
        self.preview_relationship_id.as_deref()
    }
}

#[derive(Default)]
pub(crate) struct OleLoadLimits {
    total_slide_xml_bytes: usize,
    object_count: usize,
}

#[derive(Default)]
struct GraphicFrame {
    depth: usize,
    non_visual_properties_depth: Option<usize>,
    graphic_depth: Option<usize>,
    graphic_data_depth: Option<usize>,
    graphic_data_is_ole: bool,
    has_shape_properties: bool,
    shape_id: Option<u32>,
    shape_name: Option<String>,
    open_object: Option<OpenOleObject>,
    object: Option<ParsedOleObject>,
}

struct OpenOleObject {
    depth: usize,
    legacy_shape_id: Option<String>,
    name: Option<String>,
    program_id: Option<String>,
    show_as_icon: Option<bool>,
    preview_width: Option<u32>,
    preview_height: Option<u32>,
    relationship_id: Option<String>,
    mode: Option<PptxOleObjectMode>,
    pic_depth: Option<usize>,
    preview_relationship_id: Option<String>,
}

struct ParsedOleObject {
    shape_id: Option<u32>,
    shape_name: Option<String>,
    legacy_shape_id: Option<String>,
    name: Option<String>,
    program_id: Option<String>,
    show_as_icon: Option<bool>,
    preview_width: Option<u32>,
    preview_height: Option<u32>,
    mode: PptxOleObjectMode,
    relationship_id: Option<String>,
    preview_relationship_id: Option<String>,
}

/// Load bounded, inert OLE-object metadata from one PresentationML slide.
pub(crate) fn load_slide_ole_objects(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut OleLoadLimits,
) -> Result<Vec<PptxOleObject>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "OLE-object discovery requires a PresentationML slide part",
        ));
    }
    limits.add_slide_xml(slide.blob().len())?;

    scan_ole_objects(slide.blob(), limits)?
        .into_iter()
        .enumerate()
        .map(|(object_index, parsed)| {
            let (payload_kind, target) = match parsed.relationship_id.as_deref() {
                Some(relationship_id) => {
                    let (kind, target) =
                        resolve_target(package, slide_index, slide, relationship_id)?;
                    (Some(kind), Some(target))
                },
                None => (None, None),
            };
            Ok(PptxOleObject {
                slide_index,
                object_index,
                shape_id: parsed.shape_id,
                shape_name: parsed.shape_name,
                legacy_shape_id: parsed.legacy_shape_id,
                name: parsed.name,
                program_id: parsed.program_id,
                show_as_icon: parsed.show_as_icon,
                preview_width: parsed.preview_width,
                preview_height: parsed.preview_height,
                mode: parsed.mode,
                relationship_id: parsed.relationship_id,
                payload_kind,
                target,
                preview_relationship_id: parsed.preview_relationship_id,
            })
        })
        .collect()
}

impl OleLoadLimits {
    fn add_slide_xml(&mut self, bytes: usize) -> Result<()> {
        if bytes > MAX_SLIDE_XML_BYTES {
            return Err(limit("slide XML bytes"));
        }
        self.total_slide_xml_bytes = self
            .total_slide_xml_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("total slide XML bytes"))?;
        if self.total_slide_xml_bytes > MAX_TOTAL_SLIDE_XML_BYTES {
            return Err(limit("total slide XML bytes"));
        }
        Ok(())
    }

    fn add_object(&mut self) -> Result<()> {
        self.object_count = self
            .object_count
            .checked_add(1)
            .ok_or_else(|| limit("OLE-object count"))?;
        if self.object_count > MAX_OLE_OBJECTS {
            return Err(limit("OLE-object count"));
        }
        Ok(())
    }
}

fn scan_ole_objects(xml_bytes: &[u8], limits: &mut OleLoadLimits) -> Result<Vec<ParsedOleObject>> {
    if xml_bytes.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes"));
    }

    let capabilities = MceCapabilities::ooxml_baseline();
    let mce_limits = MceLimits {
        max_input_bytes: MAX_SLIDE_XML_BYTES,
        max_output_bytes: MAX_SLIDE_XML_BYTES,
        max_depth: MAX_XML_DEPTH,
        max_namespace_bindings: 4_096,
        max_directive_tokens: 4_096,
        max_choices_per_alternate: 1_024,
    };
    let xml = process_markup_compatibility(xml_bytes, &capabilities, &mce_limits)?.xml;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut objects = Vec::new();
    let mut frames = Vec::new();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;

    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);

        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth"))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth"));
                }
                if depth == 1 {
                    validate_slide_root(&namespace, element.name(), saw_root)?;
                    saw_root = true;
                }
                inspect_start(
                    &mut frames,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    depth,
                    limits,
                )?;
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth
                    .checked_add(1)
                    .ok_or_else(|| limit("slide XML depth"))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("slide XML depth"));
                }
                if child_depth == 1 {
                    validate_slide_root(&namespace, element.name(), saw_root)?;
                    saw_root = true;
                    closed_root = true;
                }
                inspect_start(
                    &mut frames,
                    &namespace,
                    &element,
                    decoder,
                    &resolver,
                    child_depth,
                    limits,
                )?;
                inspect_end(
                    &mut frames,
                    &mut objects,
                    &namespace,
                    element.name(),
                    child_depth,
                )?;
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid slide XML nesting"));
                }
                if depth == 1 {
                    if !is_presentationml_name(&namespace, element.name(), b"sld") {
                        return Err(invalid(
                            "slide XML must close with a PresentationML sld element",
                        ));
                    }
                    closed_root = true;
                }
                inspect_end(&mut frames, &mut objects, &namespace, element.name(), depth)?;
                depth -= 1;
            },
            Event::DocType(_) => return Err(invalid("slide XML must not contain a DTD")),
            Event::Eof => {
                if !saw_root || !closed_root || depth != 0 || !frames.is_empty() {
                    return Err(invalid("unterminated or missing PresentationML slide root"));
                }
                break;
            },
            _ => {},
        }
    }

    Ok(objects)
}

fn inspect_start(
    frames: &mut Vec<GraphicFrame>,
    namespace: &ResolveResult<'_>,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
    limits: &mut OleLoadLimits,
) -> Result<()> {
    if element.attributes().with_checks(true).count() > MAX_XML_ATTRIBUTES {
        return Err(limit("slide XML attribute count"));
    }
    if is_presentationml_name(namespace, element.name(), b"graphicFrame") {
        frames.push(GraphicFrame {
            depth,
            ..GraphicFrame::default()
        });
        return Ok(());
    }

    let Some(frame) = frames.last_mut() else {
        return Ok(());
    };
    if depth == frame.depth + 1
        && is_presentationml_name(namespace, element.name(), b"nvGraphicFramePr")
    {
        frame.non_visual_properties_depth = Some(depth);
        return Ok(());
    }
    if frame
        .non_visual_properties_depth
        .is_some_and(|value| depth == value + 1)
        && is_presentationml_name(namespace, element.name(), b"cNvPr")
    {
        if frame.has_shape_properties {
            return Err(invalid(
                "graphic frame has multiple direct non-visual shape properties",
            ));
        }
        frame.has_shape_properties = true;
        frame.shape_id = optional_u32(
            unqualified_attribute_value(element, b"id", decoder)?,
            "graphic-frame shape ID",
        )?;
        frame.shape_name = bounded_optional(
            unqualified_attribute_value(element, b"name", decoder)?,
            "graphic-frame shape name",
        )?;
        return Ok(());
    }
    if depth == frame.depth + 1 && is_drawingml_name(namespace, element.name(), b"graphic") {
        frame.graphic_depth = Some(depth);
        return Ok(());
    }
    if frame.graphic_depth.is_some_and(|value| depth == value + 1)
        && is_drawingml_name(namespace, element.name(), b"graphicData")
    {
        frame.graphic_data_depth = Some(depth);
        frame.graphic_data_is_ole = bounded_optional(
            unqualified_attribute_value(element, b"uri", decoder)?,
            "graphic-data URI",
        )?
        .as_deref()
            == Some(OLE_GRAPHIC_DATA_URI);
        return Ok(());
    }
    if frame
        .graphic_data_depth
        .is_some_and(|value| depth == value + 1)
        && frame.graphic_data_is_ole
        && is_presentationml_name(namespace, element.name(), b"oleObj")
    {
        if frame.open_object.is_some() || frame.object.is_some() {
            return Err(invalid("graphic frame has multiple OLE object elements"));
        }
        limits.add_object()?;
        frame.open_object = Some(parse_open_ole_object(element, decoder, resolver, depth)?);
        return Ok(());
    }

    let Some(object) = frame.open_object.as_mut() else {
        return Ok(());
    };
    if depth == object.depth + 1 {
        if is_presentationml_name(namespace, element.name(), b"embed") {
            set_mode(object, PptxOleObjectMode::Embedded)?;
        } else if is_presentationml_name(namespace, element.name(), b"link") {
            set_mode(object, PptxOleObjectMode::Linked)?;
        } else if is_presentationml_name(namespace, element.name(), b"pic") {
            if object.pic_depth.replace(depth).is_some() {
                return Err(invalid("OLE object has multiple preview pictures"));
            }
        }
        return Ok(());
    }
    if object.pic_depth.is_some_and(|value| depth > value)
        && is_drawingml_name(namespace, element.name(), b"blip")
    {
        set_preview_relationship(object, element, decoder, resolver)?;
    }
    Ok(())
}

fn inspect_end(
    frames: &mut Vec<GraphicFrame>,
    objects: &mut Vec<ParsedOleObject>,
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    depth: usize,
) -> Result<()> {
    if let Some(frame) = frames.last_mut() {
        if let Some(object) = frame.open_object.as_mut() {
            if object.pic_depth == Some(depth) && is_presentationml_name(namespace, name, b"pic") {
                object.pic_depth = None;
            }
        }

        let closes_object = frame
            .open_object
            .as_ref()
            .is_some_and(|object| object.depth == depth)
            && is_presentationml_name(namespace, name, b"oleObj");
        if closes_object {
            let object = frame
                .open_object
                .take()
                .ok_or_else(|| invalid("missing open OLE object"))?;
            let object = finish_open_ole_object(object)?;
            if frame.object.replace(object).is_some() {
                return Err(invalid("graphic frame has multiple OLE object elements"));
            }
        }

        if frame.graphic_data_depth == Some(depth)
            && is_drawingml_name(namespace, name, b"graphicData")
        {
            frame.graphic_data_depth = None;
            frame.graphic_data_is_ole = false;
        } else if frame.graphic_depth == Some(depth)
            && is_drawingml_name(namespace, name, b"graphic")
        {
            frame.graphic_depth = None;
        } else if frame.non_visual_properties_depth == Some(depth)
            && is_presentationml_name(namespace, name, b"nvGraphicFramePr")
        {
            frame.non_visual_properties_depth = None;
        }
    }

    if is_presentationml_name(namespace, name, b"graphicFrame")
        && frames.last().is_some_and(|frame| frame.depth == depth)
    {
        let frame = frames
            .pop()
            .ok_or_else(|| invalid("missing open graphic frame"))?;
        if frame.open_object.is_some() {
            return Err(invalid("unterminated OLE object in graphic frame"));
        }
        if let Some(mut object) = frame.object {
            object.shape_id = frame.shape_id;
            object.shape_name = frame.shape_name;
            objects.push(object);
        }
    }
    Ok(())
}

fn parse_open_ole_object(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
    depth: usize,
) -> Result<OpenOleObject> {
    let legacy_shape_id = bounded_optional(
        unqualified_attribute_value(element, b"spid", decoder)?,
        "OLE legacy shape ID",
    )?;
    let name = bounded_optional(
        unqualified_attribute_value(element, b"name", decoder)?,
        "OLE object name",
    )?;
    let program_id = bounded_optional(
        unqualified_attribute_value(element, b"progId", decoder)?,
        "OLE program ID",
    )?;
    let show_as_icon = optional_bool(
        unqualified_attribute_value(element, b"showAsIcon", decoder)?,
        "OLE show-as-icon flag",
    )?;
    let preview_width = optional_u32(
        unqualified_attribute_value(element, b"imgW", decoder)?,
        "OLE preview width",
    )?;
    let preview_height = optional_u32(
        unqualified_attribute_value(element, b"imgH", decoder)?,
        "OLE preview height",
    )?;
    let relationship_id = bounded_optional(
        relationship_attribute_value(element, b"id", decoder, resolver)?,
        "OLE relationship ID",
    )?
    .filter(|value| !value.is_empty());

    Ok(OpenOleObject {
        depth,
        legacy_shape_id,
        name,
        program_id,
        show_as_icon,
        preview_width,
        preview_height,
        relationship_id,
        mode: None,
        pic_depth: None,
        preview_relationship_id: None,
    })
}

fn set_mode(object: &mut OpenOleObject, mode: PptxOleObjectMode) -> Result<()> {
    if object.mode.replace(mode).is_some() {
        return Err(invalid("OLE object has multiple embed or link elements"));
    }
    Ok(())
}

fn set_preview_relationship(
    object: &mut OpenOleObject,
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<()> {
    let embedded = relationship_attribute_value(element, b"embed", decoder, resolver)?
        .filter(|value| !value.is_empty());
    let linked = relationship_attribute_value(element, b"link", decoder, resolver)?
        .filter(|value| !value.is_empty());
    if embedded.is_some() && linked.is_some() {
        return Err(invalid(
            "OLE preview blip cannot have both embed and link relationships",
        ));
    }
    let relationship_id = bounded_optional(embedded.or(linked), "OLE preview relationship ID")?;
    let Some(relationship_id) = relationship_id else {
        return Ok(());
    };
    if object
        .preview_relationship_id
        .replace(relationship_id)
        .is_some()
    {
        return Err(invalid(
            "OLE object has multiple preview blip relationships",
        ));
    }
    Ok(())
}

fn finish_open_ole_object(object: OpenOleObject) -> Result<ParsedOleObject> {
    let mode = object
        .mode
        .ok_or_else(|| invalid("OLE object requires an embed or link element"))?;
    Ok(ParsedOleObject {
        shape_id: None,
        shape_name: None,
        legacy_shape_id: object.legacy_shape_id,
        name: object.name,
        program_id: object.program_id,
        show_as_icon: object.show_as_icon,
        preview_width: object.preview_width,
        preview_height: object.preview_height,
        mode,
        relationship_id: object.relationship_id,
        preview_relationship_id: object.preview_relationship_id,
    })
}

fn resolve_target(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    relationship_id: &str,
) -> Result<(PptxOlePayloadKind, PptxOleObjectTarget)> {
    let relationship = slide.rels().get(relationship_id).ok_or_else(|| {
        OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} OLE object references missing relationship '{relationship_id}'"
        ))
    })?;
    let relationship_type = relationship.reltype().to_owned();
    let payload_kind = payload_kind(&relationship_type).ok_or_else(|| {
        OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} OLE relationship '{relationship_id}' has unsupported type '{relationship_type}'"
        ))
    })?;

    if relationship.is_external() {
        let target = relationship.target_ref().to_owned();
        if target.is_empty() {
            return Err(OoxmlError::InvalidRelationship(format!(
                "slide {slide_index} OLE relationship '{relationship_id}' has an empty external target"
            )));
        }
        bounded(&target, "external OLE target")?;
        return Ok((
            payload_kind,
            PptxOleObjectTarget::External {
                target,
                relationship_type,
            },
        ));
    }

    let part_name = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} OLE relationship '{relationship_id}' has an invalid target: {error}"
        ))
    })?;
    let part = package.get_part(&part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!(
            "slide {slide_index} OLE relationship '{relationship_id}' targets missing part '{}': {error}",
            part_name.as_str()
        ))
    })?;
    let expected = expected_content_type(payload_kind);
    if part.content_type() != expected {
        return Err(OoxmlError::InvalidContentType {
            expected: expected.to_string(),
            got: part.content_type().to_string(),
        });
    }
    Ok((
        payload_kind,
        PptxOleObjectTarget::Internal {
            part_name,
            content_type: part.content_type().to_string(),
            relationship_type,
        },
    ))
}

fn payload_kind(relationship_type: &str) -> Option<PptxOlePayloadKind> {
    match relationship_type {
        rt::OLE_OBJECT | rt::STRICT_OLE_OBJECT => Some(PptxOlePayloadKind::OleObject),
        rt::PACKAGE | rt::STRICT_PACKAGE => Some(PptxOlePayloadKind::Package),
        _ => None,
    }
}

fn expected_content_type(kind: PptxOlePayloadKind) -> &'static str {
    match kind {
        PptxOlePayloadKind::OleObject => ct::OFC_OLE_OBJECT,
        PptxOlePayloadKind::Package => ct::OFC_PACKAGE,
    }
}

fn validate_slide_root(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    root_seen: bool,
) -> Result<()> {
    if root_seen || !is_presentationml_name(namespace, name, b"sld") {
        return Err(invalid(
            "slide XML must have one PresentationML sld root element",
        ));
    }
    Ok(())
}

fn bounded_optional(value: Option<String>, what: &str) -> Result<Option<String>> {
    if let Some(value) = &value {
        bounded(value, what)?;
    }
    Ok(value)
}

fn optional_u32(value: Option<String>, what: &str) -> Result<Option<u32>> {
    value
        .map(|value| {
            bounded(&value, what)?;
            value
                .parse()
                .map_err(|_| invalid(format!("invalid {what} '{value}'")))
        })
        .transpose()
}

fn optional_bool(value: Option<String>, what: &str) -> Result<Option<bool>> {
    value
        .map(|value| {
            bounded(&value, what)?;
            match value.as_str() {
                "true" | "1" => Ok(true),
                "false" | "0" => Ok(false),
                _ => Err(invalid(format!("invalid {what} '{value}'"))),
            }
        })
        .transpose()
}

fn bounded(value: &str, what: &str) -> Result<()> {
    if value.len() > MAX_ATTRIBUTE_BYTES {
        return Err(limit(what));
    }
    Ok(())
}

fn increment_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("slide XML node count"))?;
    if *nodes > MAX_XML_NODES {
        return Err(limit("slide XML node count"));
    }
    Ok(())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(what: &str) -> OoxmlError {
    invalid(format!("{what} exceeds the supported safety limit"))
}
