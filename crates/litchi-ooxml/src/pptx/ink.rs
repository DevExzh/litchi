//! Bounded, inert InkML content-part discovery for PowerPoint slides.
//!
//! Ink payloads are never rendered, interpreted as handwriting, recognized,
//! modified, or executed. This module only validates the declared OPC graph
//! and exposes small structural counts from the persisted InkML XML.

use crate::error::{OoxmlError, Result};
use crate::pptx::namespace::{is_presentationml_name, relationship_attribute_value};
use litchi_ooxml_common::mce::process_ooxml;
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
pub struct Annotation {
    slide_index: usize,
    content_part_index: usize,
    relationship_id: String,
    part_name: PackURI,
    trace_count: usize,
    trace_group_count: usize,
}

impl Annotation {
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
) -> Result<Vec<Annotation>> {
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
        annotations.push(Annotation {
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

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

fn limit(what: &str) -> OoxmlError {
    invalid(format!("{what} exceeds the supported safety limit"))
}

/// The package outcome of storing an InkML annotation onto a slide.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoredInkAnnotation {
    /// The relationship ID from the slide to the new InkML part.
    pub relationship_id: String,
    /// The absolute OPC part name of the stored InkML payload.
    pub part_name: PackURI,
}

/// Store an InkML annotation onto a slide as a `p:contentPart` reference.
///
/// The payload is validated as InkML (single `inkml:ink` root, bounded size
/// and structure) and stored verbatim under `/ppt/ink/`; the slide gains a
/// `customXml` relationship and a `p:contentPart` reference at the end of
/// its shape tree. The slide's namespace dialect (transitional or Strict)
/// is preserved. The ink is never rendered, recognized, or executed.
pub fn store_slide_ink_annotation(
    package: &mut OpcPackage,
    slide_name: &PackURI,
    inkml: &[u8],
) -> Result<StoredInkAnnotation> {
    if inkml.len() > MAX_INK_PART_BYTES {
        return Err(limit("InkML part bytes"));
    }
    // Validate root, structure, and trace bounds before mutating anything.
    inspect_inkml(inkml)?;
    let slide = package.get_part(slide_name)?;
    if slide.content_type() != ct::PML_SLIDE {
        return Err(invalid(
            "Ink annotation storage requires a PresentationML slide part",
        ));
    }

    let part_name = allocate_ink_part_name(package)?;
    let relationship_id = allocate_relationship_id(slide)?;
    let fragment = format!(
        "<p:contentPart xmlns:p=\"{}\" xmlns:r=\"{}\" r:id=\"{relationship_id}\"/>",
        String::from_utf8_lossy(slide_presentationml_namespace(slide.blob())?),
        String::from_utf8_lossy(slide_relationships_namespace(slide.blob())?),
    );
    let updated = insert_content_part(slide.blob(), fragment.as_bytes())?;
    let target = part_name.relative_ref(slide_name.base_uri());

    package.get_part_mut(slide_name)?.set_blob(updated);
    package.add_part(Box::new(litchi_opc::BlobPart::new(
        part_name.clone(),
        INK_CONTENT_TYPE.into(),
        inkml.to_vec(),
    )));
    package
        .get_part_mut(slide_name)?
        .rels_mut()
        .add_relationship(
            rt::CUSTOM_XML.into(),
            target,
            relationship_id.clone(),
            false,
        );
    Ok(StoredInkAnnotation {
        relationship_id,
        part_name,
    })
}

/// Allocate the first free `/ppt/ink/inkN.xml` part name.
fn allocate_ink_part_name(package: &OpcPackage) -> Result<PackURI> {
    const MAX_INK_PART_INDEX: u32 = 1_000_000;
    for index in 1..MAX_INK_PART_INDEX {
        let uri =
            PackURI::new(format!("/ppt/ink/ink{index}.xml")).map_err(OoxmlError::InvalidUri)?;
        if package.get_part(&uri).is_err() {
            return Ok(uri);
        }
    }
    Err(limit("InkML part namespace"))
}

/// Allocate the first free `rIdN` relationship ID on the slide.
fn allocate_relationship_id(slide: &dyn Part) -> Result<String> {
    const MAX_RELATIONSHIP_INDEX: u32 = 1_000_000;
    for index in 1..MAX_RELATIONSHIP_INDEX {
        let id = format!("rId{index}");
        if slide.rels().get(&id).is_none() {
            return Ok(id);
        }
    }
    Err(limit("slide relationship ID namespace"))
}

/// The slide's PresentationML namespace (transitional or Strict).
fn slide_presentationml_namespace(xml: &[u8]) -> Result<&'static [u8]> {
    Ok(
        if contains_strict_root(xml, crate::pptx::namespace::STRICT_PRESENTATIONML_NAMESPACE)? {
            crate::pptx::namespace::STRICT_PRESENTATIONML_NAMESPACE
        } else {
            crate::pptx::namespace::PRESENTATIONML_NAMESPACE
        },
    )
}

/// The relationships namespace matching the slide's dialect.
fn slide_relationships_namespace(xml: &[u8]) -> Result<&'static [u8]> {
    Ok(
        if contains_strict_root(xml, crate::pptx::namespace::STRICT_PRESENTATIONML_NAMESPACE)? {
            crate::pptx::namespace::STRICT_RELATIONSHIPS_NAMESPACE
        } else {
            crate::pptx::namespace::RELATIONSHIPS_NAMESPACE
        },
    )
}

/// Whether the slide root binds the Strict PresentationML namespace.
fn contains_strict_root(xml: &[u8], strict: &[u8]) -> Result<bool> {
    let mut reader = NsReader::from_reader(xml);
    loop {
        match reader.read_event().map_err(xml_error)? {
            Event::Start(element) | Event::Empty(element) => {
                let (namespace, _) = reader.resolver().resolve_element(element.name());
                return Ok(
                    matches!(namespace, ResolveResult::Bound(Namespace(value)) if value == strict),
                );
            },
            Event::Eof => return Err(invalid("slide XML has no root element")),
            _ => {},
        }
    }
}

/// Insert a `p:contentPart` fragment at the end of the slide's shape tree.
fn insert_content_part(xml: &[u8], fragment: &[u8]) -> Result<Vec<u8>> {
    if xml.len() > MAX_SLIDE_XML_BYTES {
        return Err(limit("slide XML bytes"));
    }
    let mut reader = NsReader::from_reader(xml);
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root_seen = false;
    let mut sp_tree_depth = None;
    let mut position = None;
    loop {
        let start = usize::try_from(reader.buffer_position())
            .map_err(|_| invalid("slide XML offset overflow"))?;
        let (namespace, event) = reader.read_resolved_event().map_err(xml_error)?;
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
                    validate_slide_root(&namespace, element.name(), root_seen)?;
                    root_seen = true;
                }
                if is_presentationml_name(&namespace, element.name(), b"spTree")
                    && sp_tree_depth.replace(depth).is_some()
                {
                    return Err(invalid("slide has multiple shape trees"));
                }
            },
            Event::Empty(element) => {
                increment_nodes(&mut nodes)?;
                if !root_seen {
                    validate_slide_root(&namespace, element.name(), false)?;
                    root_seen = true;
                }
                if is_presentationml_name(&namespace, element.name(), b"spTree") {
                    return Err(invalid("cannot insert into an empty shape tree"));
                }
            },
            Event::End(element) => {
                if depth == 0 {
                    return Err(invalid("invalid slide XML nesting"));
                }
                if sp_tree_depth == Some(depth)
                    && is_presentationml_name(&namespace, element.name(), b"spTree")
                {
                    position = Some(start);
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if depth != 0 || !root_seen {
        return Err(invalid("unterminated or missing PresentationML slide root"));
    }
    let position = position.ok_or_else(|| invalid("slide is missing a shape tree"))?;
    let size = xml
        .len()
        .checked_add(fragment.len())
        .ok_or_else(|| limit("updated slide XML bytes"))?;
    if size > MAX_SLIDE_XML_BYTES {
        return Err(limit("updated slide XML bytes"));
    }
    let mut output = Vec::with_capacity(size);
    output.extend_from_slice(&xml[..position]);
    output.extend_from_slice(fragment);
    output.extend_from_slice(&xml[position..]);
    Ok(output)
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

    fn slide_package(conformance_pml: &str, conformance_rel: &str) -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let name = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let xml = format!(
            "<p:sld xmlns:p=\"{conformance_pml}\" xmlns:r=\"{conformance_rel}\">\
             <p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/></p:spTree></p:cSld></p:sld>"
        );
        package.add_part(Box::new(litchi_opc::BlobPart::new(
            name.clone(),
            ct::PML_SLIDE.into(),
            xml.into_bytes(),
        )));
        (package, name)
    }

    const INKML: &[u8] = br#"<ink xmlns="http://www.w3.org/2003/InkML"><traceGroup><trace>1 2 3 4</trace></traceGroup></ink>"#;

    #[test]
    fn stores_and_discovers_ink_annotation_round_trip() {
        let (mut package, slide_name) = slide_package(
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        );
        let stored = store_slide_ink_annotation(&mut package, &slide_name, INKML).unwrap();
        assert_eq!(stored.relationship_id, "rId1");
        assert_eq!(stored.part_name.as_str(), "/ppt/ink/ink1.xml");

        let part = package.get_part(&stored.part_name).unwrap();
        assert_eq!(part.content_type(), INK_CONTENT_TYPE);
        assert_eq!(part.blob(), INKML);

        let slide = package.get_part(&slide_name).unwrap();
        let relationship = slide.rels().get("rId1").unwrap();
        assert_eq!(relationship.reltype(), rt::CUSTOM_XML);
        assert!(!relationship.is_external());
        let xml = String::from_utf8(slide.blob().to_vec()).unwrap();
        assert!(xml.contains("<p:contentPart"));
        assert!(xml.contains("r:id=\"rId1\""));

        // The read-side inventory discovers the stored annotation.
        let mut limits = InkLoadLimits::default();
        let annotations = load_slide_ink_annotations(&package, 0, slide, &mut limits).unwrap();
        assert_eq!(annotations.len(), 1);
        assert_eq!(annotations[0].relationship_id(), "rId1");
        assert_eq!(annotations[0].part_name(), &stored.part_name);
        assert_eq!(annotations[0].trace_count(), 1);
        assert_eq!(annotations[0].trace_group_count(), 1);

        // A second annotation gets distinct ids.
        let stored2 = store_slide_ink_annotation(&mut package, &slide_name, INKML).unwrap();
        assert_eq!(stored2.relationship_id, "rId2");
        assert_eq!(stored2.part_name.as_str(), "/ppt/ink/ink2.xml");
        let slide = package.get_part(&slide_name).unwrap();
        let mut limits = InkLoadLimits::default();
        let annotations = load_slide_ink_annotations(&package, 0, slide, &mut limits).unwrap();
        assert_eq!(annotations.len(), 2);
    }

    #[test]
    fn stores_ink_annotation_in_strict_dialect() {
        let (mut package, slide_name) = slide_package(
            "http://purl.oclc.org/ooxml/presentationml/main",
            "http://purl.oclc.org/ooxml/officeDocument/relationships",
        );
        let stored = store_slide_ink_annotation(&mut package, &slide_name, INKML).unwrap();
        let slide = package.get_part(&slide_name).unwrap();
        let xml = String::from_utf8(slide.blob().to_vec()).unwrap();
        assert!(xml.contains("xmlns:p=\"http://purl.oclc.org/ooxml/presentationml/main\""));
        assert!(
            xml.contains("xmlns:r=\"http://purl.oclc.org/ooxml/officeDocument/relationships\"")
        );
        let mut limits = InkLoadLimits::default();
        let annotations = load_slide_ink_annotations(&package, 0, slide, &mut limits).unwrap();
        assert_eq!(annotations.len(), 1);
        assert_eq!(annotations[0].part_name(), &stored.part_name);
    }

    #[test]
    fn rejects_invalid_ink_and_missing_slide() {
        let (mut package, slide_name) = slide_package(
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
        );
        // Not InkML.
        assert!(store_slide_ink_annotation(&mut package, &slide_name, b"<notInk/>").is_err());
        // Non-slide part.
        let wrong = PackURI::new("/ppt/presentation.xml").unwrap();
        package.add_part(Box::new(litchi_opc::BlobPart::new(
            wrong.clone(),
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
                .into(),
            b"<p:presentation/>".to_vec(),
        )));
        assert!(store_slide_ink_annotation(&mut package, &wrong, INKML).is_err());
        // Empty shape tree cannot be patched.
        let empty_name = PackURI::new("/ppt/slides/slide2.xml").unwrap();
        package.add_part(Box::new(litchi_opc::BlobPart::new(
            empty_name.clone(),
            ct::PML_SLIDE.into(),
            br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree/></p:cSld></p:sld>"#.to_vec(),
        )));
        assert!(store_slide_ink_annotation(&mut package, &empty_name, INKML).is_err());
        // Rejection leaves the slide untouched.
        let slide = package.get_part(&slide_name).unwrap();
        assert!(slide.rels().get("rId1").is_none());
    }
}
