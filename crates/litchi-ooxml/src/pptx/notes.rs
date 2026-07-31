//! Bounded, inert PresentationML notes-slide and notes-master package graphs.

use crate::error::{OoxmlError, Result};
use litchi_ooxml_common::mce::process_ooxml;
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::part::{BlobPart, Part};
use litchi_opc::{OpcPackage, PackURI};
use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, ResolveResult};
use quick_xml::reader::NsReader;
use std::collections::{BTreeMap, BTreeSet, HashSet};

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const A: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const AS: &str = "http://purl.oclc.org/ooxml/drawingml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const SLIDE_CT: &str = "application/vnd.openxmlformats-officedocument.presentationml.slide+xml";
const THEME_CT: &str = "application/vnd.openxmlformats-officedocument.theme+xml";
const MAX_PRESENTATION_XML: usize = 32 * 1024 * 1024;
const MAX_SLIDE_XML: usize = 16 * 1024 * 1024;
const MAX_NOTES_XML: usize = 8 * 1024 * 1024;
const MAX_MASTER_XML: usize = 16 * 1024 * 1024;
const MAX_THEME_XML: usize = 16 * 1024 * 1024;
const MAX_TOTAL_BYTES: usize = 64 * 1024 * 1024;
const MAX_NOTES_SLIDES: usize = 4096;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 128;
const MAX_ATTRIBUTES: usize = 500_000;
const MAX_ATTRIBUTE_BYTES: usize = 8 * 1024 * 1024;

/// PresentationML namespace and relationship conformance used by a notes graph.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum PptxNotesConformance {
    Transitional,
    Strict,
}

impl PptxNotesConformance {
    fn p(self) -> &'static str {
        if self == Self::Strict { PS } else { P }
    }
    fn a(self) -> &'static str {
        if self == Self::Strict { AS } else { A }
    }
    fn r(self) -> &'static str {
        if self == Self::Strict { RS } else { R }
    }
    fn notes_slide_rel(self) -> &'static str {
        if self == Self::Strict {
            rt::STRICT_NOTES_SLIDE
        } else {
            rt::NOTES_SLIDE
        }
    }
    fn notes_master_rel(self) -> &'static str {
        if self == Self::Strict {
            rt::STRICT_NOTES_MASTER
        } else {
            rt::NOTES_MASTER
        }
    }
    fn slide_rel(self) -> &'static str {
        if self == Self::Strict {
            "http://purl.oclc.org/ooxml/officeDocument/relationships/slide"
        } else {
            rt::SLIDE
        }
    }
    fn theme_rel(self) -> &'static str {
        if self == Self::Strict {
            "http://purl.oclc.org/ooxml/officeDocument/relationships/theme"
        } else {
            rt::THEME
        }
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PptxNotesThemeResource {
    pub relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    pub data: Vec<u8>,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PptxNotesMasterResource {
    pub presentation_relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    pub data: Vec<u8>,
    pub theme: PptxNotesThemeResource,
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PptxNotesSlideResource {
    pub slide_part_name: String,
    pub slide_relationship_id: String,
    pub part_name: String,
    pub content_type: String,
    pub data: Vec<u8>,
    pub backlink_relationship_id: String,
    pub notes_master_relationship_id: String,
}

impl PptxNotesSlideResource {
    /// Flatten the inert notes XML to its DrawingML text runs.
    pub fn text(&self) -> Result<Option<String>> {
        let processed = process_ooxml(&self.data)?;
        let mut reader = Reader::from_reader(processed.as_ref());
        reader.config_mut().trim_text(true);
        let mut in_text = false;
        let mut value = String::new();
        loop {
            match reader.read_event() {
                Ok(Event::Start(element)) if element.local_name().as_ref() == b"t" => {
                    in_text = true
                },
                Ok(Event::Text(text)) if in_text => {
                    let decoded = text.decode().map_err(xml_error)?;
                    let decoded = quick_xml::escape::unescape(&decoded).map_err(xml_error)?;
                    if !value.is_empty() {
                        value.push('\n');
                    }
                    value.push_str(&decoded);
                },
                Ok(Event::End(element)) if element.local_name().as_ref() == b"t" => in_text = false,
                Ok(Event::Eof) => break,
                Err(error) => return Err(xml_error(error)),
                _ => {},
            }
        }
        Ok((!value.is_empty()).then_some(value))
    }
}

#[derive(Clone, Debug, PartialEq, Eq)]
pub struct PptxNotesGraph {
    pub conformance: PptxNotesConformance,
    pub master: PptxNotesMasterResource,
    pub slides: Vec<PptxNotesSlideResource>,
}

#[derive(Default)]
struct XmlScan {
    relationship_attributes: Vec<String>,
    notes_master_ids: Vec<String>,
    slide_ids: Vec<String>,
}

/// Load and validate the complete bounded notes graph for a presentation part.
pub fn load_notes_graph(
    package: &OpcPackage,
    presentation_name: &PackURI,
) -> Result<Option<PptxNotesGraph>> {
    let presentation = package.get_part(presentation_name)?;
    if presentation.content_type() != ct::PML_PRESENTATION_MAIN {
        return Err(invalid(
            "notes graph requires a macro-free PresentationML main part",
        ));
    }
    let conformance = root_conformance(presentation.blob(), MAX_PRESENTATION_XML, "presentation")?;
    let presentation_scan = scan_xml(
        presentation.blob(),
        MAX_PRESENTATION_XML,
        conformance,
        "presentation",
    )?;
    if presentation_scan.notes_master_ids.len() > 1 {
        return Err(invalid(
            "presentation has multiple notesMasterId references",
        ));
    }
    if presentation_scan.slide_ids.len() > MAX_NOTES_SLIDES {
        return Err(limit("slide count"));
    }
    let mut slide_sources = Vec::with_capacity(presentation_scan.slide_ids.len());
    for id in &presentation_scan.slide_ids {
        validate_id(id)?;
        let relationship = presentation
            .rels()
            .get(id)
            .ok_or_else(|| invalid("presentation slide reference is missing its relationship"))?;
        if relationship.reltype() != conformance.slide_rel() || relationship.is_external() {
            return Err(invalid(
                "presentation slide relationship has wrong type or target mode",
            ));
        }
        let target = relationship_target(presentation, relationship)?;
        let slide = package.get_part(&target)?;
        if slide.content_type() != SLIDE_CT {
            return Err(invalid("slide has invalid content type"));
        }
        root_conformance(slide.blob(), MAX_SLIDE_XML, "sld").and_then(|actual| {
            if actual == conformance {
                Ok(actual)
            } else {
                Err(invalid("slide conformance differs from presentation"))
            }
        })?;
        slide_sources.push((id.clone(), target));
    }
    let master_relationships: Vec<_> = presentation
        .rels()
        .iter()
        .filter(|relationship| is_notes_master_rel(relationship.reltype()))
        .collect();
    let Some(master_id) = presentation_scan.notes_master_ids.first() else {
        if !master_relationships.is_empty()
            || package.iter_parts().any(|part| {
                matches!(
                    part.content_type(),
                    ct::PML_NOTES_MASTER | ct::PML_NOTES_SLIDE
                )
            })
            || slide_sources.iter().any(|(_, name)| {
                package.get_part(name).is_ok_and(|part| {
                    part.rels()
                        .iter()
                        .any(|relationship| is_notes_slide_rel(relationship.reltype()))
                })
            })
        {
            return Err(invalid("orphan notes graph exists without notesMasterId"));
        }
        return Ok(None);
    };
    validate_id(master_id)?;
    if master_relationships.len() != 1 {
        return Err(invalid(
            "notesMasterId and presentation notes-master relationships differ",
        ));
    }
    let master_relationship = presentation
        .rels()
        .get(master_id)
        .ok_or_else(|| invalid("notesMasterId relationship is missing"))?;
    if master_relationship.reltype() != conformance.notes_master_rel()
        || master_relationship.is_external()
    {
        return Err(invalid(
            "notes-master relationship has wrong type or target mode",
        ));
    }
    let master_name = relationship_target(presentation, master_relationship)?;
    validate_leaf_path(&master_name, "/ppt/notesMasters/", "notes master")?;
    let master_part = package.get_part(&master_name)?;
    require_content_type(master_part, ct::PML_NOTES_MASTER, "notes master")?;
    validate_resource_xml(
        master_part.blob(),
        MAX_MASTER_XML,
        conformance,
        "notesMaster",
        "notes master",
    )?;
    if master_part.rels().iter().count() != 1 {
        return Err(invalid(
            "notes master must have exactly one bounded theme relationship",
        ));
    }
    let theme_relationship = master_part
        .rels()
        .iter()
        .next()
        .ok_or_else(|| invalid("notes master lacks theme relationship"))?;
    validate_id(theme_relationship.r_id())?;
    if theme_relationship.reltype() != conformance.theme_rel() || theme_relationship.is_external() {
        return Err(invalid(
            "notes-master theme relationship has wrong type or target mode",
        ));
    }
    let theme_name = relationship_target(master_part, theme_relationship)?;
    validate_leaf_path(&theme_name, "/ppt/theme/", "notes-master theme")?;
    let theme_part = package.get_part(&theme_name)?;
    require_content_type(theme_part, THEME_CT, "notes-master theme")?;
    validate_resource_xml(
        theme_part.blob(),
        MAX_THEME_XML,
        conformance,
        "theme",
        "notes-master theme",
    )?;
    if theme_part.rels().iter().next().is_some() {
        return Err(invalid(
            "notes-master theme has unsupported outbound relationships",
        ));
    }
    let mut total = checked_add(
        master_part.blob().len(),
        theme_part.blob().len(),
        "aggregate bytes",
    )?;
    let mut slides = Vec::new();
    let mut discovered = BTreeSet::new();
    for (_, slide_name) in &slide_sources {
        let slide_part = package.get_part(slide_name)?;
        let relationships: Vec<_> = slide_part
            .rels()
            .iter()
            .filter(|relationship| is_notes_slide_rel(relationship.reltype()))
            .collect();
        if relationships.len() > 1 {
            return Err(invalid("slide has multiple notes-slide relationships"));
        }
        let Some(relationship) = relationships.first() else {
            continue;
        };
        validate_id(relationship.r_id())?;
        if relationship.reltype() != conformance.notes_slide_rel() || relationship.is_external() {
            return Err(invalid(
                "notes-slide relationship has wrong type or target mode",
            ));
        }
        let notes_name = relationship_target(slide_part, relationship)?;
        validate_leaf_path(&notes_name, "/ppt/notesSlides/", "notes slide")?;
        if !discovered.insert(notes_name.as_str().to_owned()) {
            return Err(invalid("multiple slides reference the same notes slide"));
        }
        let notes_part = package.get_part(&notes_name)?;
        require_content_type(notes_part, ct::PML_NOTES_SLIDE, "notes slide")?;
        validate_resource_xml(
            notes_part.blob(),
            MAX_NOTES_XML,
            conformance,
            "notes",
            "notes slide",
        )?;
        if notes_part.rels().iter().count() != 2 {
            return Err(invalid(
                "notes slide must have exactly slide and notes-master relationships",
            ));
        }
        let mut backlink = None;
        let mut notes_master = None;
        for child in notes_part.rels().iter() {
            validate_id(child.r_id())?;
            if child.is_external() {
                return Err(invalid("notes slide has an external relationship"));
            }
            if child.reltype() == conformance.slide_rel() {
                if backlink.replace(child).is_some() {
                    return Err(invalid("notes slide has multiple slide backlinks"));
                }
            } else if child.reltype() == conformance.notes_master_rel() {
                if notes_master.replace(child).is_some() {
                    return Err(invalid(
                        "notes slide has multiple notes-master relationships",
                    ));
                }
            } else {
                return Err(invalid(
                    "notes slide has an unsupported outbound relationship",
                ));
            }
        }
        let backlink = backlink.ok_or_else(|| invalid("notes slide lacks slide backlink"))?;
        let notes_master =
            notes_master.ok_or_else(|| invalid("notes slide lacks notes-master relationship"))?;
        if relationship_target(notes_part, backlink)? != *slide_name {
            return Err(invalid("notes slide backlink targets the wrong slide"));
        }
        if relationship_target(notes_part, notes_master)? != master_name {
            return Err(invalid("notes slide targets the wrong notes master"));
        }
        total = checked_add(total, notes_part.blob().len(), "aggregate bytes")?;
        if total > MAX_TOTAL_BYTES {
            return Err(limit("aggregate bytes"));
        }
        slides.push(PptxNotesSlideResource {
            slide_part_name: slide_name.as_str().to_owned(),
            slide_relationship_id: relationship.r_id().to_owned(),
            part_name: notes_name.as_str().to_owned(),
            content_type: notes_part.content_type().to_owned(),
            data: notes_part.blob().to_vec(),
            backlink_relationship_id: backlink.r_id().to_owned(),
            notes_master_relationship_id: notes_master.r_id().to_owned(),
        });
    }
    if package
        .iter_parts()
        .filter(|part| part.content_type() == ct::PML_NOTES_MASTER)
        .count()
        != 1
        || package
            .iter_parts()
            .filter(|part| part.content_type() == ct::PML_NOTES_SLIDE)
            .any(|part| !discovered.contains(part.partname().as_str()))
    {
        return Err(invalid("package contains orphan notes parts"));
    }
    Ok(Some(PptxNotesGraph {
        conformance,
        master: PptxNotesMasterResource {
            presentation_relationship_id: master_id.clone(),
            part_name: master_name.as_str().to_owned(),
            content_type: master_part.content_type().to_owned(),
            data: master_part.blob().to_vec(),
            theme: PptxNotesThemeResource {
                relationship_id: theme_relationship.r_id().to_owned(),
                part_name: theme_name.as_str().to_owned(),
                content_type: theme_part.content_type().to_owned(),
                data: theme_part.blob().to_vec(),
            },
        },
        slides,
    }))
}

/// Deterministically replace the resources of an already coherent notes graph.
/// Validation completes before the first package mutation.
pub fn store_notes_graph(
    package: &mut OpcPackage,
    presentation_name: &PackURI,
    graph: &PptxNotesGraph,
) -> Result<()> {
    let current = load_notes_graph(package, presentation_name)?
        .ok_or_else(|| invalid("store requires an existing coherent notes graph"))?;
    validate_graph_value(graph)?;
    if ownership(&current) != ownership(graph) {
        return Err(invalid(
            "store cannot retarget or orphan existing notes parts",
        ));
    }
    let presentation = package.get_part(presentation_name)?;
    let presentation_scan = scan_xml(
        presentation.blob(),
        MAX_PRESENTATION_XML,
        graph.conformance,
        "presentation",
    )?;
    if presentation_scan.notes_master_ids.as_slice()
        != [graph.master.presentation_relationship_id.as_str()]
    {
        return Err(invalid(
            "presentation notesMasterId does not match graph metadata",
        ));
    }
    let mut slide_names = BTreeSet::new();
    for id in presentation_scan.slide_ids {
        let relationship = presentation
            .rels()
            .get(&id)
            .ok_or_else(|| invalid("presentation slide reference is missing"))?;
        slide_names.insert(
            relationship_target(presentation, relationship)?
                .as_str()
                .to_owned(),
        );
    }
    if graph
        .slides
        .iter()
        .any(|slide| !slide_names.contains(&slide.slide_part_name))
    {
        return Err(invalid(
            "notes graph references a slide outside the presentation",
        ));
    }
    let theme_uri = PackURI::new(&graph.master.theme.part_name).map_err(OoxmlError::InvalidUri)?;
    let master_uri = PackURI::new(&graph.master.part_name).map_err(OoxmlError::InvalidUri)?;
    let theme_part = BlobPart::new(
        theme_uri.clone(),
        graph.master.theme.content_type.clone(),
        graph.master.theme.data.clone(),
    );
    debug_assert!(theme_part.rels().iter().next().is_none());
    package.add_part(Box::new(theme_part));
    let mut master_part = BlobPart::new(
        master_uri.clone(),
        graph.master.content_type.clone(),
        graph.master.data.clone(),
    );
    master_part.rels_mut().add_relationship(
        graph.conformance.theme_rel().into(),
        theme_uri.relative_ref(master_uri.base_uri()),
        graph.master.theme.relationship_id.clone(),
        false,
    );
    package.add_part(Box::new(master_part));
    for slide in &graph.slides {
        let notes_uri = PackURI::new(&slide.part_name).map_err(OoxmlError::InvalidUri)?;
        let slide_uri = PackURI::new(&slide.slide_part_name).map_err(OoxmlError::InvalidUri)?;
        let mut notes_part = BlobPart::new(
            notes_uri.clone(),
            slide.content_type.clone(),
            slide.data.clone(),
        );
        notes_part.rels_mut().add_relationship(
            graph.conformance.slide_rel().into(),
            slide_uri.relative_ref(notes_uri.base_uri()),
            slide.backlink_relationship_id.clone(),
            false,
        );
        notes_part.rels_mut().add_relationship(
            graph.conformance.notes_master_rel().into(),
            master_uri.relative_ref(notes_uri.base_uri()),
            slide.notes_master_relationship_id.clone(),
            false,
        );
        package.add_part(Box::new(notes_part));
    }
    {
        let presentation = package.get_part_mut(presentation_name)?;
        let ids: Vec<_> = presentation
            .rels()
            .iter()
            .filter(|relationship| is_notes_master_rel(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        for id in ids {
            presentation.rels_mut().remove(&id);
        }
        presentation.rels_mut().add_relationship(
            graph.conformance.notes_master_rel().into(),
            master_uri.relative_ref(presentation_name.base_uri()),
            graph.master.presentation_relationship_id.clone(),
            false,
        );
    }
    let by_slide: BTreeMap<_, _> = graph
        .slides
        .iter()
        .map(|slide| (slide.slide_part_name.as_str(), slide))
        .collect();
    for slide_name in slide_names {
        let uri = PackURI::new(&slide_name).map_err(OoxmlError::InvalidUri)?;
        let part = package.get_part_mut(&uri)?;
        let ids: Vec<_> = part
            .rels()
            .iter()
            .filter(|relationship| is_notes_slide_rel(relationship.reltype()))
            .map(|relationship| relationship.r_id().to_owned())
            .collect();
        for id in ids {
            part.rels_mut().remove(&id);
        }
        if let Some(slide) = by_slide.get(slide_name.as_str()) {
            let notes_uri = PackURI::new(&slide.part_name).map_err(OoxmlError::InvalidUri)?;
            part.rels_mut().add_relationship(
                graph.conformance.notes_slide_rel().into(),
                notes_uri.relative_ref(uri.base_uri()),
                slide.slide_relationship_id.clone(),
                false,
            );
        }
    }
    Ok(())
}

pub(crate) fn load_slide_notes_resource(
    package: &OpcPackage,
    slide_name: &PackURI,
) -> Result<Option<PptxNotesSlideResource>> {
    let presentation_name = package.main_document_part()?.partname().clone();
    Ok(
        load_notes_graph(package, &presentation_name)?.and_then(|graph| {
            graph
                .slides
                .into_iter()
                .find(|slide| slide.slide_part_name == slide_name.as_str())
        }),
    )
}

fn validate_graph_value(graph: &PptxNotesGraph) -> Result<()> {
    if graph.slides.len() > MAX_NOTES_SLIDES {
        return Err(limit("notes-slide count"));
    }
    validate_id(&graph.master.presentation_relationship_id)?;
    validate_id(&graph.master.theme.relationship_id)?;
    if graph.master.content_type != ct::PML_NOTES_MASTER
        || graph.master.theme.content_type != THEME_CT
    {
        return Err(invalid("notes master or theme has invalid content type"));
    }
    let master_uri = PackURI::new(&graph.master.part_name).map_err(OoxmlError::InvalidUri)?;
    validate_leaf_path(&master_uri, "/ppt/notesMasters/", "notes master")?;
    let theme_uri = PackURI::new(&graph.master.theme.part_name).map_err(OoxmlError::InvalidUri)?;
    validate_leaf_path(&theme_uri, "/ppt/theme/", "notes-master theme")?;
    validate_resource_xml(
        &graph.master.data,
        MAX_MASTER_XML,
        graph.conformance,
        "notesMaster",
        "notes master",
    )?;
    validate_resource_xml(
        &graph.master.theme.data,
        MAX_THEME_XML,
        graph.conformance,
        "theme",
        "notes-master theme",
    )?;
    let mut total = checked_add(
        graph.master.data.len(),
        graph.master.theme.data.len(),
        "aggregate bytes",
    )?;
    let mut sources = HashSet::new();
    let mut parts = HashSet::new();
    parts.insert(graph.master.part_name.as_str());
    parts.insert(graph.master.theme.part_name.as_str());
    for slide in &graph.slides {
        validate_id(&slide.slide_relationship_id)?;
        validate_id(&slide.backlink_relationship_id)?;
        validate_id(&slide.notes_master_relationship_id)?;
        if slide.backlink_relationship_id == slide.notes_master_relationship_id {
            return Err(invalid("notes-slide relationship IDs collide"));
        }
        if slide.content_type != ct::PML_NOTES_SLIDE {
            return Err(invalid("notes slide has invalid content type"));
        }
        let source = PackURI::new(&slide.slide_part_name).map_err(OoxmlError::InvalidUri)?;
        validate_leaf_path(&source, "/ppt/slides/", "slide")?;
        let uri = PackURI::new(&slide.part_name).map_err(OoxmlError::InvalidUri)?;
        validate_leaf_path(&uri, "/ppt/notesSlides/", "notes slide")?;
        if !sources.insert(slide.slide_part_name.as_str())
            || !parts.insert(slide.part_name.as_str())
        {
            return Err(invalid(
                "notes graph has duplicate source or resource part names",
            ));
        }
        validate_resource_xml(
            &slide.data,
            MAX_NOTES_XML,
            graph.conformance,
            "notes",
            "notes slide",
        )?;
        total = checked_add(total, slide.data.len(), "aggregate bytes")?;
        if total > MAX_TOTAL_BYTES {
            return Err(limit("aggregate bytes"));
        }
    }
    Ok(())
}

fn validate_resource_xml(
    xml: &[u8],
    max: usize,
    conformance: PptxNotesConformance,
    root: &str,
    label: &str,
) -> Result<()> {
    let scan = scan_xml(xml, max, conformance, root)?;
    if !scan.relationship_attributes.is_empty() {
        return Err(invalid(format!(
            "{label} contains unsupported outbound relationship references"
        )));
    }
    Ok(())
}

fn root_conformance(xml: &[u8], max: usize, root: &str) -> Result<PptxNotesConformance> {
    for conformance in [
        PptxNotesConformance::Transitional,
        PptxNotesConformance::Strict,
    ] {
        if scan_xml(xml, max, conformance, root).is_ok() {
            return Ok(conformance);
        }
    }
    Err(invalid(format!("invalid {root} root or namespace")))
}

fn scan_xml(
    xml: &[u8],
    max: usize,
    conformance: PptxNotesConformance,
    expected_root: &str,
) -> Result<XmlScan> {
    if xml.len() > max {
        return Err(limit("XML bytes"));
    }
    let processed = process_ooxml(xml)?;
    if processed.len() > max {
        return Err(limit("processed XML bytes"));
    }
    let mut reader = NsReader::from_reader(processed.as_ref());
    reader.config_mut().trim_text(false);
    let mut buffer = Vec::new();
    let mut depth = 0usize;
    let mut nodes = 0usize;
    let mut attributes = 0usize;
    let mut attribute_bytes = 0usize;
    let mut root_seen = false;
    let mut scan = XmlScan::default();
    loop {
        match reader.read_event_into(&mut buffer).map_err(xml_error)? {
            Event::Start(element) => {
                nodes += 1;
                depth += 1;
                if depth > MAX_DEPTH || nodes > MAX_NODES {
                    return Err(limit("XML structure"));
                }
                inspect_element(
                    &reader,
                    &element,
                    conformance,
                    expected_root,
                    !root_seen,
                    &mut attributes,
                    &mut attribute_bytes,
                    &mut scan,
                )?;
                root_seen = true;
            },
            Event::Empty(element) => {
                nodes += 1;
                if nodes > MAX_NODES || depth >= MAX_DEPTH {
                    return Err(limit("XML structure"));
                }
                inspect_element(
                    &reader,
                    &element,
                    conformance,
                    expected_root,
                    !root_seen,
                    &mut attributes,
                    &mut attribute_bytes,
                    &mut scan,
                )?;
                root_seen = true;
            },
            Event::End(_) => {
                if depth == 0 {
                    return Err(invalid("unexpected XML closing element"));
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
    if !root_seen || depth != 0 {
        return Err(invalid("missing or unterminated XML root"));
    }
    Ok(scan)
}

#[allow(clippy::too_many_arguments)]
fn inspect_element(
    reader: &NsReader<&[u8]>,
    element: &BytesStart<'_>,
    conformance: PptxNotesConformance,
    expected_root: &str,
    is_root: bool,
    attributes: &mut usize,
    attribute_bytes: &mut usize,
    scan: &mut XmlScan,
) -> Result<()> {
    let namespace = resolved(reader.resolver().resolve_element(element.name()).0)?;
    let local = std::str::from_utf8(element.local_name().as_ref())
        .map_err(xml_error)?
        .to_owned();
    if is_root
        && (namespace
            != if expected_root == "theme" {
                conformance.a()
            } else {
                conformance.p()
            }
            || local != expected_root)
    {
        return Err(invalid(format!(
            "invalid {expected_root} root or namespace"
        )));
    }
    for item in element.attributes().with_checks(true) {
        let item = item.map_err(xml_error)?;
        let raw = item.key.as_ref();
        if raw == b"xmlns" || raw.starts_with(b"xmlns:") {
            continue;
        }
        *attributes += 1;
        if *attributes > MAX_ATTRIBUTES {
            return Err(limit("XML attribute count"));
        }
        let (namespace, attr_local) = reader.resolver().resolve_attribute(item.key);
        let namespace = resolved(namespace)?;
        let attr_local = std::str::from_utf8(attr_local.as_ref()).map_err(xml_error)?;
        let raw_value = std::str::from_utf8(item.value.as_ref()).map_err(xml_error)?;
        let value = quick_xml::escape::unescape(raw_value)
            .map_err(xml_error)?
            .into_owned();
        *attribute_bytes = attribute_bytes
            .checked_add(namespace.len() + attr_local.len() + value.len())
            .ok_or_else(|| limit("XML attribute bytes"))?;
        if *attribute_bytes > MAX_ATTRIBUTE_BYTES {
            return Err(limit("XML attribute bytes"));
        }
        if namespace == conformance.r() {
            scan.relationship_attributes.push(value.clone());
            if attr_local == "id" && namespace == conformance.r() {
                if local == "notesMasterId" {
                    scan.notes_master_ids.push(value.clone());
                } else if local == "sldId" {
                    scan.slide_ids.push(value.clone());
                }
            }
        }
    }
    Ok(())
}

fn ownership(graph: &PptxNotesGraph) -> BTreeSet<String> {
    std::iter::once(graph.master.part_name.clone())
        .chain(std::iter::once(graph.master.theme.part_name.clone()))
        .chain(graph.slides.iter().map(|slide| slide.part_name.clone()))
        .collect()
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
fn validate_leaf_path(uri: &PackURI, prefix: &str, label: &str) -> Result<()> {
    let Some(rest) = uri.as_str().strip_prefix(prefix) else {
        return Err(invalid(format!("{label} is outside {prefix}")));
    };
    if rest.is_empty() || rest.contains('/') || !rest.ends_with(".xml") {
        return Err(invalid(format!("invalid {label} part path")));
    }
    Ok(())
}
fn require_content_type(part: &dyn Part, expected: &str, label: &str) -> Result<()> {
    if part.content_type() == expected {
        Ok(())
    } else {
        Err(invalid(format!("{label} has invalid content type")))
    }
}
fn is_notes_slide_rel(value: &str) -> bool {
    matches!(value, rt::NOTES_SLIDE | rt::STRICT_NOTES_SLIDE)
}
fn is_notes_master_rel(value: &str) -> bool {
    matches!(value, rt::NOTES_MASTER | rt::STRICT_NOTES_MASTER)
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
fn checked_add(left: usize, right: usize, label: &str) -> Result<usize> {
    left.checked_add(right).ok_or_else(|| limit(label))
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
    invalid(format!("PPTX notes {label} limit exceeded"))
}

#[cfg(test)]
mod tests {
    use super::*;
    const POI: &[u8] = include_bytes!("../../../../test-data/poi/test-data/slideshow/prProps.pptx");
    const LO: &[u8] =
        include_bytes!("../../../../test-data/libreoffice-core/oox/qa/unit/data/tdf131082.pptx");
    fn presentation() -> PackURI {
        PackURI::new("/ppt/presentation.xml").unwrap()
    }
    #[test]
    fn poi_and_libreoffice_notes_graphs_load_and_store_deterministically() {
        for bytes in [POI, LO] {
            let mut package = OpcPackage::from_bytes(bytes).unwrap();
            let name = presentation();
            let graph = load_notes_graph(&package, &name).unwrap().unwrap();
            assert_eq!(graph.slides.len(), 1);
            assert_eq!(graph.master.content_type, ct::PML_NOTES_MASTER);
            store_notes_graph(&mut package, &name, &graph).unwrap();
            assert_eq!(load_notes_graph(&package, &name).unwrap().unwrap(), graph);
            store_notes_graph(&mut package, &name, &graph).unwrap();
            assert_eq!(load_notes_graph(&package, &name).unwrap().unwrap(), graph);
        }
    }
    fn synthetic(conformance: PptxNotesConformance) -> (OpcPackage, PackURI) {
        let p = conformance.p();
        let a = conformance.a();
        let r = conformance.r();
        let mut package = OpcPackage::new();
        let presentation = presentation();
        let mut pres=BlobPart::new(presentation.clone(),ct::PML_PRESENTATION_MAIN.into(),format!("<p:presentation xmlns:p=\"{p}\" xmlns:r=\"{r}\"><p:notesMasterIdLst><p:notesMasterId r:id=\"rIdMaster\"/></p:notesMasterIdLst><p:sldIdLst><p:sldId id=\"256\" r:id=\"rIdSlide\"/></p:sldIdLst><p:notesSz cx=\"1\" cy=\"1\"/></p:presentation>").into_bytes());
        pres.rels_mut().add_relationship(
            conformance.notes_master_rel().into(),
            "notesMasters/notesMaster1.xml".into(),
            "rIdMaster".into(),
            false,
        );
        pres.rels_mut().add_relationship(
            conformance.slide_rel().into(),
            "slides/slide1.xml".into(),
            "rIdSlide".into(),
            false,
        );
        package.add_part(Box::new(pres));
        let slide_uri = PackURI::new("/ppt/slides/slide1.xml").unwrap();
        let mut slide = BlobPart::new(
            slide_uri,
            SLIDE_CT.into(),
            format!("<p:sld xmlns:p=\"{p}\"><p:cSld/></p:sld>").into_bytes(),
        );
        slide.rels_mut().add_relationship(
            conformance.notes_slide_rel().into(),
            "../notesSlides/notesSlide1.xml".into(),
            "rIdNotes".into(),
            false,
        );
        package.add_part(Box::new(slide));
        let master_uri = PackURI::new("/ppt/notesMasters/notesMaster1.xml").unwrap();
        let mut master=BlobPart::new(master_uri,ct::PML_NOTES_MASTER.into(),format!("<p:notesMaster xmlns:p=\"{p}\" xmlns:a=\"{a}\"><p:cSld/><p:clrMap/></p:notesMaster>").into_bytes());
        master.rels_mut().add_relationship(
            conformance.theme_rel().into(),
            "../theme/theme2.xml".into(),
            "rIdTheme".into(),
            false,
        );
        package.add_part(Box::new(master));
        let notes_uri = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
        let mut notes=BlobPart::new(notes_uri,ct::PML_NOTES_SLIDE.into(),format!("<p:notes xmlns:p=\"{p}\" xmlns:a=\"{a}\" xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:u=\"urn:unsupported\"><mc:AlternateContent><mc:Choice Requires=\"u\"><u:active/></mc:Choice><mc:Fallback><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>Strict note</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></mc:Fallback></mc:AlternateContent><p:clrMapOvr/></p:notes>").into_bytes());
        notes.rels_mut().add_relationship(
            conformance.slide_rel().into(),
            "../slides/slide1.xml".into(),
            "rIdBack".into(),
            false,
        );
        notes.rels_mut().add_relationship(
            conformance.notes_master_rel().into(),
            "../notesMasters/notesMaster1.xml".into(),
            "rIdMaster".into(),
            false,
        );
        package.add_part(Box::new(notes));
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/theme/theme2.xml").unwrap(),
            THEME_CT.into(),
            format!("<a:theme xmlns:a=\"{a}\" name=\"Notes\"/>").into_bytes(),
        )));
        (package, presentation)
    }
    #[test]
    fn strict_mce_graph_round_trips_and_projects_text() {
        let (mut package, name) = synthetic(PptxNotesConformance::Strict);
        let graph = load_notes_graph(&package, &name).unwrap().unwrap();
        assert_eq!(
            graph.slides[0].text().unwrap().as_deref(),
            Some("Strict note")
        );
        store_notes_graph(&mut package, &name, &graph).unwrap();
        assert_eq!(load_notes_graph(&package, &name).unwrap().unwrap(), graph);
    }
    #[test]
    fn rejects_external_wrong_root_outbound_orphan_and_caps_before_mutation() {
        let (mut package, name) = synthetic(PptxNotesConformance::Transitional);
        let notes = PackURI::new("/ppt/notesSlides/notesSlide1.xml").unwrap();
        {
            let part = package.get_part_mut(&notes).unwrap();
            part.rels_mut().remove("rIdBack");
            part.rels_mut().add_relationship(
                rt::SLIDE.into(),
                "https://example.invalid/slide".into(),
                "rIdBack".into(),
                true,
            );
        }
        assert!(load_notes_graph(&package, &name).is_err());
        let (mut package, name) = synthetic(PptxNotesConformance::Transitional);
        package
            .get_part_mut(&PackURI::new("/ppt/notesMasters/notesMaster1.xml").unwrap())
            .unwrap()
            .set_blob(format!("<p:wrong xmlns:p=\"{P}\"/>").into_bytes());
        assert!(load_notes_graph(&package, &name).is_err());
        let (mut package, name) = synthetic(PptxNotesConformance::Transitional);
        package
            .get_part_mut(&notes)
            .unwrap()
            .rels_mut()
            .add_relationship(
                rt::IMAGE.into(),
                "../media/image1.png".into(),
                "rIdImage".into(),
                false,
            );
        assert!(load_notes_graph(&package, &name).is_err());
        let (mut package, name) = synthetic(PptxNotesConformance::Transitional);
        package.add_part(Box::new(BlobPart::new(
            PackURI::new("/ppt/notesSlides/orphan.xml").unwrap(),
            ct::PML_NOTES_SLIDE.into(),
            format!("<p:notes xmlns:p=\"{P}\"><p:cSld/></p:notes>").into_bytes(),
        )));
        assert!(load_notes_graph(&package, &name).is_err());
        let (mut package, name) = synthetic(PptxNotesConformance::Transitional);
        let mut graph = load_notes_graph(&package, &name).unwrap().unwrap();
        graph.slides[0].data = vec![b' '; MAX_NOTES_XML + 1];
        let before = package.get_part(&name).unwrap().blob().to_vec();
        assert!(store_notes_graph(&mut package, &name, &graph).is_err());
        assert_eq!(package.get_part(&name).unwrap().blob(), before);
    }
}
