//! Bounded, inert InkML content-part discovery for PowerPoint slides.
//!
//! Ink payloads are never rendered, interpreted as handwriting, recognized,
//! modified, or executed. This module only validates the declared OPC graph
//! and exposes small structural counts from the persisted InkML XML.

use crate::common::mce::process_ooxml;
use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::{is_presentationml_name, relationship_attribute_value};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcPackage, PackURI, Part};
use quick_xml::events::Event;
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;

/// The OPC content type for an InkML content part.
pub const INK_CONTENT_TYPE: &str = "application/inkml+xml";

const INKML_NAMESPACE: &[u8] = b"http://www.w3.org/2003/InkML";
const MAX_SLIDE_XML_BYTES: usize = 32 * 1024 * 1024;
const MAX_INK_PART_BYTES: usize = 16 * 1024 * 1024;
const MAX_TOTAL_INK_BYTES: usize = 256 * 1024 * 1024;
const MAX_INK_ANNOTATIONS: usize = 4_096;
const MAX_CONTENT_PARTS_PER_SLIDE: usize = 4_096;
const MAX_XML_NODES: usize = 250_000;
const MAX_XML_DEPTH: usize = 128;
const MAX_INK_TRACES: usize = 65_536;
const MAX_INK_TRACE_GROUPS: usize = 65_536;
const MAX_RELATIONSHIP_ID_BYTES: usize = 1_024;

/// An inert InkML content part anchored on a presentation slide.
///
/// This exposes only package identity and structural counts. It does not
/// expose trace geometry or attempt handwriting recognition or rendering.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PptxInkAnnotation {
    slide_index: usize,
    content_part_index: usize,
    relationship_id: String,
    part_name: PackURI,
    trace_count: usize,
    trace_group_count: usize,
}

impl PptxInkAnnotation {
    /// Return the zero-based index of the slide that owns this annotation.
    #[inline]
    pub fn slide_index(&self) -> usize {
        self.slide_index
    }

    /// Return the zero-based source-order index of the slide content part.
    #[inline]
    pub fn content_part_index(&self) -> usize {
        self.content_part_index
    }

    /// Return the relationship ID from the slide to the InkML part.
    #[inline]
    pub fn relationship_id(&self) -> &str {
        &self.relationship_id
    }

    /// Return the absolute OPC part name of the InkML payload.
    #[inline]
    pub fn part_name(&self) -> &PackURI {
        &self.part_name
    }

    /// Return the number of persisted InkML trace elements.
    #[inline]
    pub fn trace_count(&self) -> usize {
        self.trace_count
    }

    /// Return the number of persisted InkML trace-group elements.
    #[inline]
    pub fn trace_group_count(&self) -> usize {
        self.trace_group_count
    }
}

#[derive(Default)]
pub(crate) struct InkLoadLimits {
    annotation_count: usize,
    total_bytes: usize,
}

#[derive(Default)]
struct InkSummary {
    trace_count: usize,
    trace_group_count: usize,
}

pub(crate) fn load_slide_ink_annotations(
    package: &OpcPackage,
    slide_index: usize,
    slide: &dyn Part,
    limits: &mut InkLoadLimits,
) -> Result<Vec<PptxInkAnnotation>> {
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "Ink content-part discovery requires a PresentationML slide part",
        ));
    }

    let relationship_ids = scan_content_part_relationship_ids(slide.blob())?;
    let mut annotations = Vec::new();
    for (content_part_index, relationship_id) in relationship_ids.into_iter().enumerate() {
        let relationship = slide.rels().get(&relationship_id).ok_or_else(|| {
            OoxmlError::InvalidRelationship(format!(
                "slide {slide_index} Ink content part references missing relationship '{relationship_id}'"
            ))
        })?;
        if relationship.is_external() || relationship.reltype() != rt::CUSTOM_XML {
            return Err(OoxmlError::InvalidRelationship(format!(
                "slide {slide_index} Ink content-part relationship '{relationship_id}' must be an internal customXml relationship"
            )));
        }

        let part_name = relationship.target_partname().map_err(|error| {
            OoxmlError::InvalidRelationship(format!(
                "slide {slide_index} Ink content-part relationship '{relationship_id}' has an invalid target: {error}"
            ))
        })?;
        let part = package.get_part(&part_name).map_err(|error| {
            OoxmlError::PartNotFound(format!(
                "slide {slide_index} Ink content-part relationship '{relationship_id}' targets missing part '{}': {error}",
                part_name.as_str()
            ))
        })?;
        if part.content_type() != INK_CONTENT_TYPE {
            continue;
        }

        limits.add_annotation(part.blob().len())?;
        let summary = inspect_inkml(part.blob())?;
        annotations.push(PptxInkAnnotation {
            slide_index,
            content_part_index,
            relationship_id,
            part_name,
            trace_count: summary.trace_count,
            trace_group_count: summary.trace_group_count,
        });
    }

    Ok(annotations)
}

impl InkLoadLimits {
    fn add_annotation(&mut self, bytes: usize) -> Result<()> {
        self.annotation_count = self
            .annotation_count
            .checked_add(1)
            .ok_or_else(|| limit("Ink annotation count"))?;
        if self.annotation_count > MAX_INK_ANNOTATIONS {
            return Err(limit("Ink annotation count"));
        }
        self.total_bytes = self
            .total_bytes
            .checked_add(bytes)
            .ok_or_else(|| limit("total InkML bytes"))?;
        if self.total_bytes > MAX_TOTAL_INK_BYTES {
            return Err(limit("total InkML bytes"));
        }
        Ok(())
    }
}

fn scan_content_part_relationship_ids(xml_bytes: &[u8]) -> Result<Vec<String>> {
    if xml_bytes.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes"));
    }

    let xml = process_ooxml(xml_bytes)?;
    let mut reader = NsReader::from_reader(xml.as_ref());
    let mut relationship_ids = Vec::new();
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
                if is_presentationml_name(&namespace, element.name(), b"contentPart") {
                    push_content_part_relationship(
                        &mut relationship_ids,
                        &element,
                        decoder,
                        &resolver,
                    )?;
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
                }
                if is_presentationml_name(&namespace, element.name(), b"contentPart") {
                    push_content_part_relationship(
                        &mut relationship_ids,
                        &element,
                        decoder,
                        &resolver,
                    )?;
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
                depth -= 1;
            },
            Event::DocType(_) => return Err(invalid("slide XML must not contain a DTD")),
            Event::Eof => {
                if !saw_root || !closed_root || depth != 0 {
                    return Err(invalid("unterminated or missing PresentationML slide root"));
                }
                break;
            },
            _ => {},
        }
    }

    Ok(relationship_ids)
}

fn push_content_part_relationship(
    relationship_ids: &mut Vec<String>,
    element: &quick_xml::events::BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    resolver: &quick_xml::name::NamespaceResolver,
) -> Result<()> {
    let relationship_id = relationship_attribute_value(element, b"id", decoder, resolver)?
        .ok_or_else(|| invalid("PresentationML contentPart is missing r:id"))?;
    if relationship_id.is_empty() || relationship_id.len() > MAX_RELATIONSHIP_ID_BYTES {
        return Err(invalid(
            "PresentationML contentPart has an invalid relationship ID",
        ));
    }
    if relationship_ids.len() >= MAX_CONTENT_PARTS_PER_SLIDE {
        return Err(limit("slide content-part count"));
    }
    relationship_ids.push(relationship_id);
    Ok(())
}

fn inspect_inkml(xml_bytes: &[u8]) -> Result<InkSummary> {
    if xml_bytes.len() > MAX_INK_PART_BYTES {
        return Err(limit("InkML part bytes"));
    }

    let mut reader = NsReader::from_reader(xml_bytes);
    let mut summary = InkSummary::default();
    let mut nodes = 0usize;
    let mut depth = 0usize;
    let mut saw_root = false;
    let mut closed_root = false;

    loop {
        let event = reader
            .read_resolved_event()
            .map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let (namespace, event) = event;

        match event {
            Event::Start(element) => {
                increment_nodes(&mut nodes)?;
                depth = depth.checked_add(1).ok_or_else(|| limit("InkML depth"))?;
                if depth > MAX_XML_DEPTH {
                    return Err(limit("InkML depth"));
                }
                if depth == 1 {
                    validate_ink_root(&namespace, element.name(), saw_root)?;
                    saw_root = true;
                }
                observe_ink_element(&mut summary, &namespace, element.name())?;
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                let child_depth = depth.checked_add(1).ok_or_else(|| limit("InkML depth"))?;
                if child_depth > MAX_XML_DEPTH {
                    return Err(limit("InkML depth"));
                }
                if child_depth == 1 {
                    validate_ink_root(&namespace, element.name(), saw_root)?;
                    saw_root = true;
                    closed_root = true;
                }
                observe_ink_element(&mut summary, &namespace, element.name())?;
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid InkML nesting"));
                }
                if depth == 1 {
                    if !is_inkml_name(&namespace, element.name(), b"ink") {
                        return Err(invalid("InkML must close with an ink element"));
                    }
                    closed_root = true;
                }
                depth -= 1;
            },
            Event::DocType(_) => return Err(invalid("InkML must not contain a DTD")),
            Event::Eof => {
                if !saw_root || !closed_root || depth != 0 {
                    return Err(invalid("unterminated or missing InkML root"));
                }
                break;
            },
            _ => {},
        }
    }

    Ok(summary)
}

fn observe_ink_element(
    summary: &mut InkSummary,
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
) -> Result<()> {
    if is_inkml_name(namespace, name, b"trace") {
        summary.trace_count = summary
            .trace_count
            .checked_add(1)
            .ok_or_else(|| limit("InkML trace count"))?;
        if summary.trace_count > MAX_INK_TRACES {
            return Err(limit("InkML trace count"));
        }
    } else if is_inkml_name(namespace, name, b"traceGroup") {
        summary.trace_group_count = summary
            .trace_group_count
            .checked_add(1)
            .ok_or_else(|| limit("InkML trace-group count"))?;
        if summary.trace_group_count > MAX_INK_TRACE_GROUPS {
            return Err(limit("InkML trace-group count"));
        }
    }
    Ok(())
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

fn validate_ink_root(
    namespace: &ResolveResult<'_>,
    name: QName<'_>,
    root_seen: bool,
) -> Result<()> {
    if root_seen || !is_inkml_name(namespace, name, b"ink") {
        return Err(invalid("InkML must have one InkML ink root element"));
    }
    Ok(())
}

fn is_inkml_name(namespace: &ResolveResult<'_>, name: QName<'_>, local_name: &[u8]) -> bool {
    name.local_name().as_ref() == local_name
        && matches!(
            namespace,
            ResolveResult::Bound(Namespace(value)) if *value == INKML_NAMESPACE
        )
}

fn increment_nodes(nodes: &mut usize) -> Result<()> {
    *nodes = nodes
        .checked_add(1)
        .ok_or_else(|| limit("XML node count"))?;
    if *nodes > MAX_XML_NODES {
        return Err(limit("XML node count"));
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

    #[test]
    fn scans_strict_slide_content_part_relationships() {
        let xml = br#"<p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main"
            xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships">
            <p:cSld><p:spTree><p:contentPart r:id="rIdInk"/></p:spTree></p:cSld>
        </p:sld>"#;

        assert_eq!(
            scan_content_part_relationship_ids(xml).unwrap(),
            vec!["rIdInk"]
        );
    }

    #[test]
    fn inkml_summary_counts_only_inkml_trace_elements() {
        let xml = br#"<i:ink xmlns:i="http://www.w3.org/2003/InkML" xmlns:f="urn:foreign">
            <i:traceGroup><i:trace/><f:trace/></i:traceGroup>
        </i:ink>"#;

        let summary = inspect_inkml(xml).unwrap();
        assert_eq!(summary.trace_count, 1);
        assert_eq!(summary.trace_group_count, 1);
    }
}
