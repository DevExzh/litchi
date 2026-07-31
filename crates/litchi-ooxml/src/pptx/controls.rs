//! Bounded, inert discovery of PowerPoint slide controls (ActiveX/OCX).
//!
//! A slide may declare embedded controls through the `p:controls` container
//! (ECMA-376 Part 1, 19.3.1.6 `control` and 15.2.13 Controls Part). Each
//! `p:control` element carries presentation metadata (`name`, `showAsIcon`,
//! `imgW`, `imgH`, legacy `spid`) and an `r:id` relationship to a controls
//! part whose `ax:ocx` descriptor names the control class and persistence
//! model and relates to an opaque binary state part.
//!
//! This module returns only stored PresentationML and OPC metadata. It never
//! instantiates a control, resolves a CLSID, decodes MS-OFORMS/CFB state,
//! executes a macro, or follows an external relationship.

use crate::common::xml::unqualified_attribute_value;
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::{is_presentationml_name, relationship_attribute_value};
use crate::xlsx::active_x::ActiveXDescriptor;
use litchi_ooxml_common::{MceCapabilities, MceLimits, process_markup_compatibility};
use litchi_opc::constants::content_type as ct;
use litchi_opc::{OpcPackage, PackURI, Part};
use quick_xml::encoding::Decoder;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{NamespaceResolver, QName, ResolveResult};
use quick_xml::reader::NsReader;

pub use crate::xlsx::active_x::Persistence;

const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_TOTAL_SLIDE_XML_BYTES: usize = 256 * 1024 * 1024;
const MAX_CONTROLS: usize = 4_096;
const MAX_BINARY_BYTES: usize = 64 * 1024 * 1024;
const MAX_TOTAL_BINARY_BYTES: usize = 256 * 1024 * 1024;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;
const MAX_XML_ATTRIBUTES: usize = 64;
const MAX_ATTRIBUTE_BYTES: usize = 4_096;

/// Relationship type from a slide to a controls part (transitional).
const CONTROL_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control";
/// Relationship type from a slide to a controls part (strict).
const STRICT_CONTROL_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/control";
/// Relationship type from a controls part to its binary state part.
const BINARY_RELATIONSHIP: &str =
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary";
/// Content type of a controls part (`ax:ocx` descriptor).
const DESCRIPTOR_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX+xml";
/// Content type of an ActiveX binary state part.
const BINARY_CONTENT_TYPE: &str = "application/vnd.ms-office.activeX";

/// An inert slide control reference (`p:control`) and its resolved descriptor.
#[derive(Debug, Clone)]
pub struct PptxSlideControl {
    slide_index: usize,
    control_index: usize,
    shape_id: Option<String>,
    name: Option<String>,
    show_as_icon: Option<bool>,
    image_width: Option<u32>,
    image_height: Option<u32>,
    relationship_id: Option<String>,
    descriptor: Option<PptxControlDescriptor>,
}

impl PptxSlideControl {
    /// Zero-based slide index within the presentation.
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Zero-based control index within the slide's `p:controls` container.
    pub fn control_index(&self) -> usize {
        self.control_index
    }

    /// The legacy VML shape identifier (`spid`), as stored.
    pub fn shape_id(&self) -> Option<&str> {
        self.shape_id.as_deref()
    }

    /// The control name, when declared.
    pub fn name(&self) -> Option<&str> {
        self.name.as_deref()
    }

    /// Whether the control is displayed as an icon, when declared.
    pub fn show_as_icon(&self) -> Option<bool> {
        self.show_as_icon
    }

    /// The width of the displayed control image in EMUs (`imgW`).
    pub fn image_width(&self) -> Option<u32> {
        self.image_width
    }

    /// The height of the displayed control image in EMUs (`imgH`).
    pub fn image_height(&self) -> Option<u32> {
        self.image_height
    }

    /// The relationship ID from the slide to the controls part (`r:id`).
    pub fn relationship_id(&self) -> Option<&str> {
        self.relationship_id.as_deref()
    }

    /// The resolved controls-part descriptor, when the control declares a
    /// resolvable `r:id` relationship.
    pub fn descriptor(&self) -> Option<&PptxControlDescriptor> {
        self.descriptor.as_ref()
    }
}

/// Inert metadata of a resolved controls part (`ax:ocx` descriptor).
#[derive(Debug, Clone)]
pub struct PptxControlDescriptor {
    part_name: PackURI,
    class_id: String,
    license: Option<String>,
    persistence: Persistence,
    binary: Option<PptxControlBinary>,
}

impl PptxControlDescriptor {
    /// Absolute package part name of the controls part.
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    /// The control class identifier (`ax:classid`), as stored.
    pub fn class_id(&self) -> &str {
        &self.class_id
    }

    /// The license key (`ax:license`), when declared.
    pub fn license(&self) -> Option<&str> {
        self.license.as_deref()
    }

    /// The declared persistence model (`ax:persistence`).
    pub fn persistence(&self) -> Persistence {
        self.persistence
    }

    /// Inert metadata of the binary state part, when the descriptor relates
    /// to one. The binary payload itself is never read or interpreted.
    pub fn binary(&self) -> Option<&PptxControlBinary> {
        self.binary.as_ref()
    }
}

/// Inert OPC metadata of an ActiveX binary state part.
#[derive(Debug, Clone)]
pub struct PptxControlBinary {
    relationship_id: String,
    part_name: PackURI,
    byte_length: usize,
}

impl PptxControlBinary {
    /// The relationship ID from the controls part to the binary part.
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Absolute package part name of the binary part.
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    /// Stored size of the binary payload in bytes. The payload bytes are
    /// never copied, decoded, or executed.
    pub fn byte_length(&self) -> usize {
        self.byte_length
    }
}

/// Shared bounding state for control discovery across a presentation.
#[derive(Default)]
pub(crate) struct ControlLoadLimits {
    total_slide_xml_bytes: usize,
    control_count: usize,
    total_binary_bytes: usize,
}

impl ControlLoadLimits {
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

    fn add_control(&mut self) -> Result<()> {
        self.control_count = self
            .control_count
            .checked_add(1)
            .ok_or_else(|| limit("slide control count"))?;
        if self.control_count > MAX_CONTROLS {
            return Err(limit("slide control count"));
        }
        Ok(())
    }

    fn add_binary(&mut self, bytes: usize) -> Result<()> {
        if bytes > MAX_BINARY_BYTES {
            return Err(limit("control binary bytes"));
        }
        self.total_binary_bytes = self
            .total_binary_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("total control binary bytes"))?;
        if self.total_binary_bytes > MAX_TOTAL_BINARY_BYTES {
            return Err(limit("total control binary bytes"));
        }
        Ok(())
    }
}

#[derive(Default)]
struct ParsedControl {
    shape_id: Option<String>,
    name: Option<String>,
    show_as_icon: Option<bool>,
    image_width: Option<u32>,
    image_height: Option<u32>,
    relationship_id: Option<String>,
}

/// Load bounded, inert control metadata from one PresentationML slide.
pub(crate) fn load_slide_controls(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut ControlLoadLimits,
) -> Result<Vec<PptxSlideControl>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "control discovery requires a PresentationML slide part",
        ));
    }
    limits.add_slide_xml(slide.blob().len())?;

    scan_controls(slide.blob(), limits)?
        .into_iter()
        .enumerate()
        .map(|(control_index, parsed)| {
            let descriptor = match parsed.relationship_id.as_deref() {
                Some(relationship_id) => Some(resolve_descriptor(
                    package,
                    slide_index,
                    slide,
                    relationship_id,
                    limits,
                )?),
                None => None,
            };
            Ok(PptxSlideControl {
                slide_index,
                control_index,
                shape_id: parsed.shape_id,
                name: parsed.name,
                show_as_icon: parsed.show_as_icon,
                image_width: parsed.image_width,
                image_height: parsed.image_height,
                relationship_id: parsed.relationship_id,
                descriptor,
            })
        })
        .collect()
}

fn scan_controls(xml_bytes: &[u8], limits: &mut ControlLoadLimits) -> Result<Vec<ParsedControl>> {
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
    let mut controls = Vec::new();
    // `p:controls` is a child of `p:cSld` (CT_CommonSlideData), so it sits one
    // level below the slide's common-slide-data element.
    let mut common_slide_data_depth: Option<usize> = None;
    let mut container_depth: Option<usize> = None;
    let mut saw_container = false;
    let mut open_control_depth: Option<usize> = None;
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
                } else if depth == 2 && is_presentationml_name(&namespace, element.name(), b"cSld")
                {
                    common_slide_data_depth = Some(depth);
                } else if common_slide_data_depth == Some(depth - 1)
                    && is_presentationml_name(&namespace, element.name(), b"controls")
                {
                    if saw_container {
                        return Err(invalid(
                            "slide has multiple PresentationML controls containers",
                        ));
                    }
                    saw_container = true;
                    container_depth = Some(depth);
                } else if container_depth == Some(depth - 1)
                    && open_control_depth.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"control")
                {
                    limits.add_control()?;
                    controls.push(parse_control(&element, decoder, &resolver)?);
                    open_control_depth = Some(depth);
                }
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
                } else if common_slide_data_depth == Some(child_depth - 1)
                    && is_presentationml_name(&namespace, element.name(), b"controls")
                {
                    // An empty p:controls container declares no controls.
                    if saw_container {
                        return Err(invalid(
                            "slide has multiple PresentationML controls containers",
                        ));
                    }
                    saw_container = true;
                } else if container_depth == Some(child_depth - 1)
                    && open_control_depth.is_none()
                    && is_presentationml_name(&namespace, element.name(), b"control")
                {
                    limits.add_control()?;
                    controls.push(parse_control(&element, decoder, &resolver)?);
                }
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
                if open_control_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"control")
                {
                    open_control_depth = None;
                }
                if container_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"controls")
                {
                    container_depth = None;
                }
                if common_slide_data_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"cSld")
                {
                    common_slide_data_depth = None;
                }
                depth -= 1;
            },
            Event::DocType(_) => return Err(invalid("slide XML must not contain a DTD")),
            Event::Eof => {
                if !saw_root
                    || !closed_root
                    || depth != 0
                    || container_depth.is_some()
                    || open_control_depth.is_some()
                {
                    return Err(invalid("unterminated or missing PresentationML slide root"));
                }
                break;
            },
            _ => {},
        }
    }

    Ok(controls)
}

fn parse_control(
    element: &BytesStart<'_>,
    decoder: Decoder,
    resolver: &NamespaceResolver,
) -> Result<ParsedControl> {
    if element.attributes().with_checks(true).count() > MAX_XML_ATTRIBUTES {
        return Err(limit("slide XML attribute count"));
    }
    Ok(ParsedControl {
        shape_id: bounded_optional(
            unqualified_attribute_value(element, b"spid", decoder)?,
            "control legacy shape ID",
        )?,
        name: bounded_optional(
            unqualified_attribute_value(element, b"name", decoder)?,
            "control name",
        )?,
        show_as_icon: optional_bool(
            unqualified_attribute_value(element, b"showAsIcon", decoder)?,
            "control show-as-icon flag",
        )?,
        image_width: optional_u32(
            unqualified_attribute_value(element, b"imgW", decoder)?,
            "control image width",
        )?,
        image_height: optional_u32(
            unqualified_attribute_value(element, b"imgH", decoder)?,
            "control image height",
        )?,
        relationship_id: bounded_optional(
            relationship_attribute_value(element, b"id", decoder, resolver)?,
            "control relationship ID",
        )?
        .filter(|value| !value.is_empty()),
    })
}

fn resolve_descriptor(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    relationship_id: &str,
    limits: &mut ControlLoadLimits,
) -> Result<PptxControlDescriptor> {
    let relationship = slide.rels().get(relationship_id).ok_or_else(|| {
        OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} control references missing relationship '{relationship_id}'"
        ))
    })?;
    let relationship_type = relationship.reltype();
    if !matches!(
        relationship_type,
        CONTROL_RELATIONSHIP | STRICT_CONTROL_RELATIONSHIP
    ) {
        return Err(OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} control relationship '{relationship_id}' has unsupported type '{relationship_type}'"
        )));
    }
    if relationship.is_external() {
        return Err(OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} control relationship '{relationship_id}' cannot be external"
        )));
    }
    let part_name = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} control relationship '{relationship_id}' has an invalid target: {error}"
        ))
    })?;
    let part = package.get_part(&part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!(
            "slide {slide_index} control relationship '{relationship_id}' targets missing part '{}': {error}",
            part_name.as_str()
        ))
    })?;
    if part.content_type() != DESCRIPTOR_CONTENT_TYPE {
        return Err(OoxmlError::InvalidContentType {
            expected: DESCRIPTOR_CONTENT_TYPE.to_string(),
            got: part.content_type().to_string(),
        });
    }
    let descriptor = ActiveXDescriptor::parse(part.blob())?;

    let binary = match descriptor.relationship_id.as_deref() {
        Some(binary_id) => Some(resolve_binary(
            package,
            slide_index,
            part,
            binary_id,
            limits,
        )?),
        None => None,
    };

    Ok(PptxControlDescriptor {
        part_name,
        class_id: descriptor.class_id,
        license: descriptor.license,
        persistence: descriptor.persistence,
        binary,
    })
}

fn resolve_binary(
    package: &OpcPackage,
    slide_index: usize,
    descriptor_part: &dyn Part,
    relationship_id: &str,
    limits: &mut ControlLoadLimits,
) -> Result<PptxControlBinary> {
    let relationship = descriptor_part
        .rels()
        .get(relationship_id)
        .ok_or_else(|| {
            OoxmlError::InvalidRelationship(format!(
                "slide {slide_index} control descriptor references missing relationship '{relationship_id}'"
            ))
        })?;
    if relationship.reltype() != BINARY_RELATIONSHIP {
        return Err(OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} control binary relationship '{relationship_id}' has unsupported type '{}'",
            relationship.reltype()
        )));
    }
    if relationship.is_external() {
        return Err(OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} control binary relationship '{relationship_id}' cannot be external"
        )));
    }
    let part_name = relationship.target_partname().map_err(|error| {
        OoxmlError::InvalidRelationship(format!(
            "slide {slide_index} control binary relationship '{relationship_id}' has an invalid target: {error}"
        ))
    })?;
    let part = package.get_part(&part_name).map_err(|error| {
        OoxmlError::PartNotFound(format!(
            "slide {slide_index} control binary relationship '{relationship_id}' targets missing part '{}': {error}",
            part_name.as_str()
        ))
    })?;
    if part.content_type() != BINARY_CONTENT_TYPE {
        return Err(OoxmlError::InvalidContentType {
            expected: BINARY_CONTENT_TYPE.to_string(),
            got: part.content_type().to_string(),
        });
    }
    limits.add_binary(part.blob().len())?;
    Ok(PptxControlBinary {
        relationship_id: relationship_id.to_string(),
        part_name,
        byte_length: part.blob().len(),
    })
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

#[cfg(test)]
mod tests {
    use super::*;

    const P_NS: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
    const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    fn scan(xml: &str) -> Result<Vec<ParsedControl>> {
        scan_controls(xml.as_bytes(), &mut ControlLoadLimits::default())
    }

    #[test]
    fn parses_direct_control_elements() {
        let xml = format!(
            r#"<p:sld xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:cSld><p:spTree/><p:controls><p:control spid="1031" name="CheckBox1" showAsIcon="1" r:id="rId2" imgW="2685960" imgH="923760"/><p:control name="CheckBox2" r:id="rId3"/></p:controls></p:cSld><p:clrMapOvr/></p:sld>"#
        );
        let controls = scan(&xml).unwrap();
        assert_eq!(controls.len(), 2);
        assert_eq!(controls[0].shape_id.as_deref(), Some("1031"));
        assert_eq!(controls[0].name.as_deref(), Some("CheckBox1"));
        assert_eq!(controls[0].show_as_icon, Some(true));
        assert_eq!(controls[0].image_width, Some(2_685_960));
        assert_eq!(controls[0].image_height, Some(923_760));
        assert_eq!(controls[0].relationship_id.as_deref(), Some("rId2"));
        assert_eq!(controls[1].name.as_deref(), Some("CheckBox2"));
        assert_eq!(controls[1].show_as_icon, None);
        assert_eq!(controls[1].image_width, None);
        assert_eq!(controls[1].relationship_id.as_deref(), Some("rId3"));
    }

    #[test]
    fn selects_the_markup_compatibility_choice_branch() {
        let xml = format!(
            r#"<p:sld xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:cSld><p:controls><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:v="urn:schemas-microsoft-com:vml" Requires="v"><p:control spid="1031" name="CheckBox1" r:id="rId2" imgW="2685960" imgH="923760"/></mc:Choice><mc:Fallback><p:control name="CheckBox1" r:id="rId2" imgW="2685960" imgH="923760"><p:pic/></p:control></mc:Fallback></mc:AlternateContent></p:controls></p:cSld></p:sld>"#
        );
        let controls = scan(&xml).unwrap();
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].shape_id.as_deref(), Some("1031"));
        assert_eq!(controls[0].relationship_id.as_deref(), Some("rId2"));
    }

    #[test]
    fn falls_back_when_the_choice_namespace_is_not_understood() {
        let xml = format!(
            r#"<p:sld xmlns:p="{P_NS}" xmlns:r="{R_NS}"><p:cSld><p:controls><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:x="urn:unknown-dialect" Requires="x"><p:control spid="1" name="New" r:id="rId9"/></mc:Choice><mc:Fallback><p:control name="Old" r:id="rId2"/></mc:Fallback></mc:AlternateContent></p:controls></p:cSld></p:sld>"#
        );
        let controls = scan(&xml).unwrap();
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].name.as_deref(), Some("Old"));
    }

    #[test]
    fn slide_without_controls_yields_nothing() {
        let xml = format!(r#"<p:sld xmlns:p="{P_NS}"><p:cSld/></p:sld>"#);
        assert!(scan(&xml).unwrap().is_empty());
        let xml = format!(r#"<p:sld xmlns:p="{P_NS}"><p:cSld><p:controls/></p:cSld></p:sld>"#);
        assert!(scan(&xml).unwrap().is_empty());
    }

    #[test]
    fn accepts_the_strict_presentationml_dialect() {
        let xml = r#"<p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships"><p:cSld><p:controls><p:control name="C" r:id="rId1"/></p:controls></p:cSld></p:sld>"#;
        let controls = scan(xml).unwrap();
        assert_eq!(controls.len(), 1);
        assert_eq!(controls[0].relationship_id.as_deref(), Some("rId1"));
    }

    #[test]
    fn rejects_multiple_containers_and_invalid_flags() {
        let xml = format!(
            r#"<p:sld xmlns:p="{P_NS}"><p:cSld><p:controls/><p:controls/></p:cSld></p:sld>"#
        );
        assert!(scan(&xml).is_err());
        let xml = format!(
            r#"<p:sld xmlns:p="{P_NS}"><p:cSld><p:controls><p:control showAsIcon="maybe"/></p:controls></p:cSld></p:sld>"#
        );
        assert!(scan(&xml).is_err());
        let xml = format!(
            r#"<p:sld xmlns:p="{P_NS}"><p:cSld><p:controls><p:control imgW="-1"/></p:controls></p:cSld></p:sld>"#
        );
        assert!(scan(&xml).is_err());
    }

    #[test]
    fn rejects_foreign_roots_and_doctype() {
        assert!(scan(r#"<p:sld xmlns:p="urn:wrong"/> "#).is_err());
        let xml = format!(r#"<!DOCTYPE p:sld><p:sld xmlns:p="{P_NS}"/>"#);
        assert!(scan(&xml).is_err());
    }
}
