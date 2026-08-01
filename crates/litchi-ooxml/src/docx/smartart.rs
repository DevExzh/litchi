//! Bounded, inert SmartArt (DrawingML diagram) inventory owned by a DOCX main document.
//!
//! A Word SmartArt graphic is anchored in the document body as a
//! `w:drawing`/`a:graphicData` element in the `drawingml/diagram` namespace
//! holding a `dgm:relIds` reference with four relationship IDs (data, layout,
//! quick style, colors). The relationships on the main document part point at
//! the diagram parts under `word/diagrams/`; an optional Microsoft-extension
//! `diagramDrawing` relationship locates the pre-rendered drawing.
//!
//! [`load_smart_arts`] resolves that graph and parses the data-model node
//! trees and the layout/quick-style/colors part headers into a typed
//! inventory. Both the transitional and the ISO Strict namespace dialects are
//! supported. Everything is treated as inert metadata: layout algorithms and
//! style bodies are never interpreted.

use crate::diagrams::{
    DGM_NAMESPACE, DGM_NAMESPACE_STRICT, DIAGRAM_COLORS_REL, DIAGRAM_DATA_REL, DIAGRAM_LAYOUT_REL,
    DIAGRAM_QUICK_STYLE_REL, DiagramDataModel, DiagramDefinition, DiagramNode, DiagramType,
    MS_DIAGRAM_DRAWING_REL, STRICT_DIAGRAM_COLORS_REL, STRICT_DIAGRAM_DATA_REL,
    STRICT_DIAGRAM_LAYOUT_REL, STRICT_DIAGRAM_QUICK_STYLE_REL,
};
use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::constants::content_type as ct;
use litchi_opc::part::Part;
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WS: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const DOCUMENT_CT: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const DIAGRAM_PART_PREFIX: &str = "/word/diagrams/";
const MAX_DOCUMENT_XML: usize = 32 * 1024 * 1024;
const MAX_SMART_ARTS: usize = 64;
const MAX_PART_XML: usize = 16 * 1024 * 1024;
const MAX_TOTAL_BYTES: usize = 256 * 1024 * 1024;
const MAX_NODES: usize = 200_000;
const MAX_DEPTH: usize = 128;

/// Namespace dialect of a DOCX SmartArt graph.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum DocxDiagramConformance {
    /// Transitional (`schemas.openxmlformats.org`) namespaces.
    Transitional,
    /// ISO/IEC 29500 Strict (`purl.oclc.org`) namespaces.
    Strict,
}

impl DocxDiagramConformance {
    fn w(self) -> &'static str {
        if self == Self::Strict { WS } else { W }
    }
    fn a(self) -> &'static str {
        if self == Self::Strict { AS } else { A }
    }
    fn r(self) -> &'static str {
        if self == Self::Strict { RS } else { R }
    }
    fn dgm(self) -> &'static str {
        if self == Self::Strict {
            DGM_NAMESPACE_STRICT
        } else {
            DGM_NAMESPACE
        }
    }
    fn data_rel(self) -> &'static str {
        if self == Self::Strict {
            STRICT_DIAGRAM_DATA_REL
        } else {
            DIAGRAM_DATA_REL
        }
    }
    fn layout_rel(self) -> &'static str {
        if self == Self::Strict {
            STRICT_DIAGRAM_LAYOUT_REL
        } else {
            DIAGRAM_LAYOUT_REL
        }
    }
    fn quick_style_rel(self) -> &'static str {
        if self == Self::Strict {
            STRICT_DIAGRAM_QUICK_STYLE_REL
        } else {
            DIAGRAM_QUICK_STYLE_REL
        }
    }
    fn colors_rel(self) -> &'static str {
        if self == Self::Strict {
            STRICT_DIAGRAM_COLORS_REL
        } else {
            DIAGRAM_COLORS_REL
        }
    }
}

/// A typed, inert SmartArt diagram anchored in a Word document.
#[derive(Clone, Debug)]
pub struct DocxSmartArt {
    /// Namespace dialect of the owning document.
    pub conformance: DocxDiagramConformance,
    /// Relationship ID of the data-model part (`r:dm`).
    pub data_relationship_id: String,
    /// Relationship ID of the layout part (`r:lo`).
    pub layout_relationship_id: Option<String>,
    /// Relationship ID of the quick-style part (`r:qs`).
    pub quick_style_relationship_id: Option<String>,
    /// Relationship ID of the colors part (`r:cs`).
    pub colors_relationship_id: Option<String>,
    /// Data-model part name (e.g. `/word/diagrams/data1.xml`).
    pub data_part_name: String,
    /// Layout part name, when referenced.
    pub layout_part_name: Option<String>,
    /// Quick-style part name, when referenced.
    pub quick_style_part_name: Option<String>,
    /// Colors part name, when referenced.
    pub colors_part_name: Option<String>,
    /// Pre-rendered drawing part name, when discoverable.
    pub drawing_part_name: Option<String>,
    /// Parsed data-model point and connection graph.
    pub data: DiagramDataModel,
    /// Layout part header metadata, when referenced.
    pub layout: Option<DiagramDefinition>,
    /// Quick-style part header metadata, when referenced.
    pub quick_style: Option<DiagramDefinition>,
    /// Colors part header metadata, when referenced.
    pub colors: Option<DiagramDefinition>,
}

impl DocxSmartArt {
    /// Best-effort diagram type, inferred from the data model's layout type
    /// identifier, falling back to the layout part's `uniqueId`.
    pub fn diagram_type(&self) -> DiagramType {
        if let Some(uri) = self
            .data
            .document_point()
            .and_then(|point| point.layout_type_id.as_deref())
        {
            return DiagramType::from_layout_uri(uri);
        }
        if let Some(uri) = self
            .layout
            .as_ref()
            .and_then(|layout| layout.unique_id.as_deref())
        {
            return DiagramType::from_layout_uri(uri);
        }
        DiagramType::Unknown
    }

    /// Content-node hierarchy implied by the data model's connection graph.
    pub fn node_tree(&self) -> Vec<DiagramNode> {
        self.data.node_tree()
    }

    /// All text content of the diagram, one line per content node.
    pub fn text(&self) -> String {
        self.data.text()
    }
}

/// Relationship references of one `dgm:relIds` anchor.
#[derive(Default)]
struct DiagramAnchor {
    data_id: Option<String>,
    layout_id: Option<String>,
    quick_style_id: Option<String>,
    colors_id: Option<String>,
}

/// Load the typed SmartArt inventory anchored in the given main document part.
///
/// Diagrams are returned in document order. The diagram parts are validated
/// (relationship types, `word/diagrams/` containment, content types, size
/// caps) and parsed as inert metadata.
pub fn load_smart_arts(package: &OpcPackage, document_name: &PackURI) -> Result<Vec<DocxSmartArt>> {
    let document = package.get_part(document_name)?;
    if document.content_type() != DOCUMENT_CT {
        return Err(invalid(
            "SmartArt inventory requires a macro-free DOCX main part",
        ));
    }
    let (conformance, anchors) = document_anchors(document.blob())?;
    if anchors.len() > MAX_SMART_ARTS {
        return Err(limit("SmartArt count"));
    }

    // Document-level Microsoft-extension diagramDrawing relationships, in
    // relationship order. These are associated with anchors positionally when
    // unambiguous (see below); Word/LibreOffice also attach the drawing to the
    // data part itself, which is preferred.
    let mut document_drawings = Vec::new();
    for relationship in document.rels().iter() {
        if relationship.reltype() == MS_DIAGRAM_DRAWING_REL {
            let target = relationship_target(document, relationship)?;
            validate_part_path(&target, "diagram drawing")?;
            document_drawings.push(target);
        }
    }

    let mut smart_arts = Vec::with_capacity(anchors.len());
    let mut total = 0usize;
    for (index, anchor) in anchors.iter().enumerate() {
        let data_id = anchor
            .data_id
            .as_deref()
            .ok_or_else(|| invalid("SmartArt anchor lacks a data relationship"))?;
        validate_id(data_id)?;
        let (data_name, data_xml) = resolve_part(
            package,
            document,
            data_id,
            conformance.data_rel(),
            ct::DML_DIAGRAM_DATA,
            "diagram data",
        )?;
        add_total(&mut total, data_xml.len(), "aggregate bytes")?;
        let data = DiagramDataModel::parse(&data_xml)?;

        let (layout_relationship_id, layout_part_name, layout) = resolve_definition(
            package,
            document,
            anchor.layout_id.as_deref(),
            conformance.layout_rel(),
            ct::DML_DIAGRAM_LAYOUT,
            "diagram layout",
            DiagramDefinition::parse_layout,
            &mut total,
        )?;
        let (quick_style_relationship_id, quick_style_part_name, quick_style) = resolve_definition(
            package,
            document,
            anchor.quick_style_id.as_deref(),
            conformance.quick_style_rel(),
            ct::DML_DIAGRAM_STYLE,
            "diagram quick style",
            DiagramDefinition::parse_quick_style,
            &mut total,
        )?;
        let (colors_relationship_id, colors_part_name, colors) = resolve_definition(
            package,
            document,
            anchor.colors_id.as_deref(),
            conformance.colors_rel(),
            ct::DML_DIAGRAM_COLORS,
            "diagram colors",
            DiagramDefinition::parse_colors,
            &mut total,
        )?;

        // Pre-rendered drawing: prefer a relationship owned by the data part
        // (PowerPoint/Word style with a `dsp:dataModelExt` extension); fall
        // back to document-level relationships when the association is
        // unambiguous (LibreOffice style).
        let data_part = package.get_part(&data_name)?;
        let mut drawing_part_name = None;
        for relationship in data_part.rels().iter() {
            if relationship.reltype() == MS_DIAGRAM_DRAWING_REL {
                let target = relationship_target(data_part, relationship)?;
                validate_part_path(&target, "diagram drawing")?;
                let part = package.get_part(&target)?;
                if part.content_type() != ct::DML_DIAGRAM_DRAWING {
                    return Err(invalid("diagram drawing has invalid content type"));
                }
                add_total(&mut total, part.blob().len(), "aggregate bytes")?;
                drawing_part_name = Some(target.as_str().to_owned());
                break;
            }
        }
        if drawing_part_name.is_none() {
            let fallback = if document_drawings.len() == anchors.len() {
                document_drawings.get(index)
            } else if anchors.len() == 1 {
                document_drawings.first()
            } else {
                None
            };
            if let Some(target) = fallback {
                let part = package.get_part(target)?;
                if part.content_type() != ct::DML_DIAGRAM_DRAWING {
                    return Err(invalid("diagram drawing has invalid content type"));
                }
                add_total(&mut total, part.blob().len(), "aggregate bytes")?;
                drawing_part_name = Some(target.as_str().to_owned());
            }
        }

        smart_arts.push(DocxSmartArt {
            conformance,
            data_relationship_id: data_id.to_owned(),
            layout_relationship_id,
            quick_style_relationship_id,
            colors_relationship_id,
            data_part_name: data_name.as_str().to_owned(),
            layout_part_name,
            quick_style_part_name,
            colors_part_name,
            drawing_part_name,
            data,
            layout,
            quick_style,
            colors,
        });
    }
    Ok(smart_arts)
}

/// Resolve and parse an optional definition part (layout, quick style, colors).
#[allow(clippy::too_many_arguments)]
fn resolve_definition(
    package: &OpcPackage,
    document: &dyn Part,
    relationship_id: Option<&str>,
    relationship_type: &str,
    content_type: &str,
    label: &str,
    parse: fn(&str) -> Result<DiagramDefinition>,
    total: &mut usize,
) -> Result<(Option<String>, Option<String>, Option<DiagramDefinition>)> {
    let Some(relationship_id) = relationship_id else {
        return Ok((None, None, None));
    };
    validate_id(relationship_id)?;
    let (name, xml) = resolve_part(
        package,
        document,
        relationship_id,
        relationship_type,
        content_type,
        label,
    )?;
    add_total(total, xml.len(), "aggregate bytes")?;
    let definition = parse(&xml)?;
    Ok((
        Some(relationship_id.to_owned()),
        Some(name.as_str().to_owned()),
        Some(definition),
    ))
}

/// Resolve a document relationship to a bounded, content-type-checked part.
fn resolve_part(
    package: &OpcPackage,
    document: &dyn Part,
    relationship_id: &str,
    relationship_type: &str,
    content_type: &str,
    label: &str,
) -> Result<(PackURI, String)> {
    let relationship = document
        .rels()
        .get(relationship_id)
        .ok_or_else(|| invalid(format!("{label} relationship is missing")))?;
    if relationship.reltype() != relationship_type || relationship.is_external() {
        return Err(invalid(format!(
            "{label} relationship has wrong type or target mode"
        )));
    }
    let name = relationship_target(document, relationship)?;
    validate_part_path(&name, label)?;
    let part = package.get_part(&name)?;
    if part.content_type() != content_type {
        return Err(invalid(format!("{label} has invalid content type")));
    }
    if part.blob().len() > MAX_PART_XML {
        return Err(limit("diagram part XML bytes"));
    }
    let xml = std::str::from_utf8(part.blob())
        .map_err(xml_error)?
        .to_owned();
    Ok((name, xml))
}

/// Scan the document body for `dgm:relIds` SmartArt anchors, detecting the
/// namespace dialect from the document root.
fn document_anchors(xml: &[u8]) -> Result<(DocxDiagramConformance, Vec<DiagramAnchor>)> {
    for conformance in [
        DocxDiagramConformance::Transitional,
        DocxDiagramConformance::Strict,
    ] {
        if let Ok(anchors) = scan_document_xml(xml, conformance) {
            return Ok((conformance, anchors));
        }
    }
    Err(invalid("invalid DOCX document root or SmartArt anchors"))
}

fn scan_document_xml(
    xml: &[u8],
    conformance: DocxDiagramConformance,
) -> Result<Vec<DiagramAnchor>> {
    if xml.len() > MAX_DOCUMENT_XML {
        return Err(limit("document XML bytes"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > MAX_DOCUMENT_XML {
        return Err(limit("processed document XML bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut root = false;
    // Stack of (graphicData depth, parsed anchor) frames for diagram anchors.
    let mut frames: Vec<(usize, Option<DiagramAnchor>)> = Vec::new();
    let mut anchors = Vec::new();

    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                depth += 1;
                nodes += 1;
                if nodes > MAX_NODES || depth > MAX_DEPTH {
                    return Err(limit("document XML structure"));
                }
                let namespace = element_namespace(&reader, &element)?;
                let local = local_name(&element)?;
                if !root {
                    if namespace != conformance.w() || local != "document" {
                        return Err(invalid("invalid document root or namespace"));
                    }
                    root = true;
                } else if namespace == conformance.a() && local == "graphicData" {
                    let is_diagram =
                        attribute(&element, "uri")?.as_deref() == Some(conformance.dgm());
                    if is_diagram {
                        frames.push((depth, None));
                    }
                } else if namespace == conformance.dgm() && local == "relIds" {
                    let Some(frame) = frames.last_mut() else {
                        return Err(invalid("SmartArt relIds is outside diagram graphicData"));
                    };
                    if frame.1.is_some() {
                        return Err(invalid("SmartArt graphicData has multiple relIds children"));
                    }
                    frame.1 = Some(rel_ids(&reader, &element, conformance)?);
                }
            },
            Event::Empty(element) => {
                nodes += 1;
                if nodes > MAX_NODES || depth + 1 > MAX_DEPTH {
                    return Err(limit("document XML structure"));
                }
                let namespace = element_namespace(&reader, &element)?;
                let local = local_name(&element)?;
                if !root {
                    if namespace != conformance.w() || local != "document" {
                        return Err(invalid("invalid document root or namespace"));
                    }
                    root = true;
                } else if namespace == conformance.a()
                    && local == "graphicData"
                    && attribute(&element, "uri")?.as_deref() == Some(conformance.dgm())
                {
                    return Err(invalid("SmartArt graphicData lacks relIds child"));
                } else if namespace == conformance.dgm() && local == "relIds" {
                    // `dgm:relIds` is an empty element in practice; mirror the
                    // `Event::Start` handling so both forms resolve.
                    let Some(frame) = frames.last_mut() else {
                        return Err(invalid("SmartArt relIds is outside diagram graphicData"));
                    };
                    if frame.1.is_some() {
                        return Err(invalid("SmartArt graphicData has multiple relIds children"));
                    }
                    frame.1 = Some(rel_ids(&reader, &element, conformance)?);
                }
            },
            Event::End(_) => {
                if let Some((_, anchor)) = frames.pop_if(|frame| frame.0 == depth) {
                    anchors.push(
                        anchor.ok_or_else(|| invalid("SmartArt graphicData lacks relIds child"))?,
                    );
                }
                if depth == 0 {
                    return Err(invalid("unexpected document XML closing element"));
                }
                depth -= 1;
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::CData(_) => return Err(invalid("CDATA is rejected")),
            Event::Eof => break,
            _ => {},
        }
        buffer.clear();
    }
    if !root || depth != 0 || !frames.is_empty() {
        return Err(invalid("missing or unterminated document root"));
    }
    Ok(anchors)
}

/// Parse the relationship IDs of a `dgm:relIds` element.
fn rel_ids(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    conformance: DocxDiagramConformance,
) -> Result<DiagramAnchor> {
    let mut anchor = DiagramAnchor::default();
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        let (namespace, name) = reader.resolver().resolve_attribute(item.key);
        if resolved(namespace)? != conformance.r() {
            continue;
        }
        let value = std::str::from_utf8(item.value.as_ref()).map_err(xml_error)?;
        let value = quick_xml::escape::unescape(value)
            .map_err(xml_error)?
            .into_owned();
        match name.as_ref() {
            b"dm" => anchor.data_id = Some(value),
            b"lo" => anchor.layout_id = Some(value),
            b"qs" => anchor.quick_style_id = Some(value),
            b"cs" => anchor.colors_id = Some(value),
            _ => {},
        }
    }
    if anchor.data_id.is_none() {
        return Err(invalid("SmartArt relIds lacks a data relationship"));
    }
    Ok(anchor)
}

fn attribute(element: &BytesStart<'_>, name: &str) -> Result<Option<String>> {
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        if item.key.local_name().as_ref() == name.as_bytes() {
            let value = std::str::from_utf8(item.value.as_ref()).map_err(xml_error)?;
            return Ok(Some(
                quick_xml::escape::unescape(value)
                    .map_err(xml_error)?
                    .into_owned(),
            ));
        }
    }
    Ok(None)
}

fn relationship_target(
    part: &dyn Part,
    relationship: &litchi_opc::Relationship,
) -> Result<PackURI> {
    if relationship.is_external() {
        return Err(invalid("external relationship is rejected"));
    }
    PackURI::from_rel_ref(part.partname().base_uri(), relationship.target_ref())
        .map_err(OoxmlError::InvalidFormat)
}

fn validate_part_path(uri: &PackURI, label: &str) -> Result<()> {
    let Some(rest) = uri.as_str().strip_prefix(DIAGRAM_PART_PREFIX) else {
        return Err(invalid(format!("{label} is outside {DIAGRAM_PART_PREFIX}")));
    };
    if rest.is_empty() || rest.contains('/') || !rest.to_ascii_lowercase().ends_with(".xml") {
        return Err(invalid(format!("invalid {label} path or suffix")));
    }
    Ok(())
}

fn validate_id(value: &str) -> Result<()> {
    let mut bytes = value.bytes();
    let Some(first) = bytes.next() else {
        return Err(invalid("relationship ID is empty"));
    };
    if !(first.is_ascii_alphabetic() || first == b'_')
        || !bytes.all(|byte| byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b'-' | b'.'))
    {
        Err(invalid("invalid relationship ID"))
    } else {
        Ok(())
    }
}

fn add_total(total: &mut usize, size: usize, label: &str) -> Result<()> {
    *total = total.checked_add(size).ok_or_else(|| limit(label))?;
    if *total > MAX_TOTAL_BYTES {
        return Err(limit(label));
    }
    Ok(())
}

fn element_namespace(reader: &NsReader<&[u8]>, element: &BytesStart<'_>) -> Result<String> {
    resolved(reader.resolver().resolve_element(element.name()).0)
}

fn local_name(element: &BytesStart<'_>) -> Result<String> {
    Ok(std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned())
}

fn resolved(value: ResolveResult<'_>) -> Result<String> {
    match value {
        ResolveResult::Bound(Namespace(value)) => {
            Ok(std::str::from_utf8(value).map_err(xml_error)?.to_owned())
        },
        ResolveResult::Unbound => Ok(String::new()),
        ResolveResult::Unknown(prefix) => Err(invalid(format!(
            "unbound XML prefix '{}'",
            String::from_utf8_lossy(prefix.as_ref())
        ))),
    }
}

fn xml_error(error: impl std::fmt::Display) -> OoxmlError {
    OoxmlError::Xml(error.to_string())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

fn limit(label: &str) -> OoxmlError {
    invalid(format!("DOCX SmartArt {label} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::part::BlobPart;

    const STRICT_DOCX: &[u8] = include_bytes!(
        "../../../../test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/strict.docx"
    );

    fn document_name() -> PackURI {
        PackURI::new("/word/document.xml").unwrap()
    }

    fn data_xml(conformance: DocxDiagramConformance) -> String {
        format!(
            "<dgm:dataModel xmlns:dgm=\"{}\" xmlns:a=\"{}\">\
             <dgm:ptLst>\
             <dgm:pt modelId=\"0\" type=\"doc\"><dgm:prSet loTypeId=\"urn:microsoft.com/office/officeart/2005/8/layout/cycle2\"/></dgm:pt>\
             <dgm:pt modelId=\"1\"><dgm:t><a:p><a:r><a:t>a</a:t></a:r></a:p></dgm:t></dgm:pt>\
             <dgm:pt modelId=\"2\"><dgm:t><a:p><a:r><a:t>b</a:t></a:r></a:p></dgm:t></dgm:pt>\
             </dgm:ptLst>\
             <dgm:cxnLst>\
             <dgm:cxn modelId=\"10\" srcId=\"0\" destId=\"1\" srcOrd=\"0\" destOrd=\"0\"/>\
             <dgm:cxn modelId=\"11\" srcId=\"0\" destId=\"2\" srcOrd=\"1\" destOrd=\"0\"/>\
             </dgm:cxnLst><dgm:bg/><dgm:whole/></dgm:dataModel>",
            conformance.dgm(),
            conformance.a()
        )
    }

    /// Build a synthetic single-diagram package in the given dialect.
    fn synthetic(conformance: DocxDiagramConformance) -> (OpcPackage, PackURI) {
        let mut package = OpcPackage::new();
        let name = document_name();
        let document_xml = format!(
            "<w:document xmlns:w=\"{}\" xmlns:a=\"{}\" xmlns:dgm=\"{}\" xmlns:r=\"{}\">\
             <w:body><w:p><w:r><w:drawing><a:graphic>\
             <a:graphicData uri=\"{}\">\
             <dgm:relIds r:dm=\"rIdDm\" r:lo=\"rIdLo\" r:qs=\"rIdQs\" r:cs=\"rIdCs\"/>\
             </a:graphicData></a:graphic></w:drawing></w:r></w:p></w:body></w:document>",
            conformance.w(),
            conformance.a(),
            conformance.dgm(),
            conformance.r(),
            conformance.dgm()
        );
        let mut document =
            BlobPart::new(name.clone(), DOCUMENT_CT.into(), document_xml.into_bytes());
        for (id, reltype, target) in [
            ("rIdDm", conformance.data_rel(), "diagrams/data1.xml"),
            ("rIdLo", conformance.layout_rel(), "diagrams/layout1.xml"),
            (
                "rIdQs",
                conformance.quick_style_rel(),
                "diagrams/quickStyle1.xml",
            ),
            ("rIdCs", conformance.colors_rel(), "diagrams/colors1.xml"),
            (
                "rIdDrawing",
                MS_DIAGRAM_DRAWING_REL,
                "diagrams/drawing1.xml",
            ),
        ] {
            document
                .rels_mut()
                .add_relationship(reltype.into(), target.into(), id.into(), false);
        }
        package.add_part(Box::new(document));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/diagrams/data1.xml").unwrap(),
            ct::DML_DIAGRAM_DATA.into(),
            data_xml(conformance).into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/diagrams/layout1.xml").unwrap(),
            ct::DML_DIAGRAM_LAYOUT.into(),
            format!(
                "<dgm:layoutDef xmlns:dgm=\"{}\" uniqueId=\"urn:microsoft.com/office/officeart/2005/8/layout/cycle2\">\
                 <dgm:catLst><dgm:cat type=\"cycle\" pri=\"1000\"/></dgm:catLst></dgm:layoutDef>",
                conformance.dgm()
            )
            .into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/diagrams/quickStyle1.xml").unwrap(),
            ct::DML_DIAGRAM_STYLE.into(),
            format!(
                "<dgm:styleDef xmlns:dgm=\"{}\" uniqueId=\"urn:test/quickstyle/simple1\"/>",
                conformance.dgm()
            )
            .into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/diagrams/colors1.xml").unwrap(),
            ct::DML_DIAGRAM_COLORS.into(),
            format!(
                "<dgm:colorsDef xmlns:dgm=\"{}\" uniqueId=\"urn:test/colors/accent1_2\"/>",
                conformance.dgm()
            )
            .into_bytes(),
        )));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/diagrams/drawing1.xml").unwrap(),
            ct::DML_DIAGRAM_DRAWING.into(),
            b"<dsp:drawing xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\"/>".to_vec(),
        )));
        (package, name)
    }

    #[test]
    fn loads_synthetic_inventory_in_both_dialects() {
        for conformance in [
            DocxDiagramConformance::Transitional,
            DocxDiagramConformance::Strict,
        ] {
            let (package, name) = synthetic(conformance);
            let smart_arts = load_smart_arts(&package, &name).unwrap();
            assert_eq!(smart_arts.len(), 1);
            let smart_art = &smart_arts[0];
            assert_eq!(smart_art.conformance, conformance);
            assert_eq!(smart_art.diagram_type(), DiagramType::Cycle);
            assert_eq!(smart_art.data_relationship_id, "rIdDm");
            assert_eq!(smart_art.data_part_name, "/word/diagrams/data1.xml");
            assert_eq!(
                smart_art.drawing_part_name.as_deref(),
                Some("/word/diagrams/drawing1.xml")
            );
            assert_eq!(smart_art.text(), "a\nb");
            assert_eq!(
                smart_art
                    .layout
                    .as_ref()
                    .unwrap()
                    .categories
                    .first()
                    .unwrap()
                    .category_type,
                "cycle"
            );
        }
    }

    #[test]
    fn reads_libreoffice_strict_fixture() {
        let package = OpcPackage::from_bytes(STRICT_DOCX).unwrap();
        let name = document_name();
        let smart_arts = load_smart_arts(&package, &name).unwrap();
        assert_eq!(smart_arts.len(), 1);
        let smart_art = &smart_arts[0];
        assert_eq!(smart_art.conformance, DocxDiagramConformance::Strict);
        assert_eq!(smart_art.diagram_type(), DiagramType::Cycle);
        assert_eq!(smart_art.text(), "a\nb\nc");
        assert_eq!(smart_art.data.points.len(), 20);
        assert_eq!(smart_art.data.connections.len(), 22);
        assert_eq!(
            smart_art.drawing_part_name.as_deref(),
            Some("/word/diagrams/drawing1.xml")
        );
        assert_eq!(
            smart_art.layout.as_ref().unwrap().unique_id.as_deref(),
            Some("urn:microsoft.com/office/officeart/2005/8/layout/cycle2")
        );
    }

    #[test]
    fn rejects_missing_or_mistyped_relationships() {
        let (package, name) = synthetic(DocxDiagramConformance::Transitional);
        // Wrong content type on the data part.
        let mut package = package;
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/word/diagrams/data1.xml").unwrap(),
            ct::DML_DIAGRAM_LAYOUT.into(),
            data_xml(DocxDiagramConformance::Transitional).into_bytes(),
        )));
        assert!(load_smart_arts(&package, &name).is_err());

        // Anchor references a missing relationship.
        let (package, name) = synthetic(DocxDiagramConformance::Transitional);
        let mut package = package;
        let document = package.get_part_mut(&name).unwrap();
        document.rels_mut().remove("rIdDm");
        assert!(load_smart_arts(&package, &name).is_err());
    }
}
