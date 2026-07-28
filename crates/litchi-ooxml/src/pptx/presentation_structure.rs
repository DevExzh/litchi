//! Package-aware custom-show and section mutation for PowerPoint presentations.
//!
//! Custom shows reference slide relationships while PowerPoint 2010 sections
//! reference stable numeric presentation slide IDs. This module resolves both
//! representations into one validated graph and patches only the corresponding
//! children of `ppt/presentation.xml`.

use crate::error::{OoxmlError, Result};
use crate::pptx::customshow::{CustomShow, CustomShowList};
use crate::pptx::sections::{Section, SectionList};
use litchi_core::xml::escape_xml;
use litchi_opc::OpcPackage;
use quick_xml::events::{BytesStart, Event};
use quick_xml::{Reader, XmlVersion};
use std::collections::{HashMap, HashSet};

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const SLIDE_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/slide";
const SECTION_URI: &str = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 128;

/// One entry in `p:sldIdLst`, resolved through the presentation relationship set.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationSlideReference {
    pub slide_id: u32,
    pub relationship_id: String,
    pub part_name: String,
}

/// Validated presentation ordering, custom shows, and modern sections.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PresentationStructure {
    pub slides: Vec<PresentationSlideReference>,
    pub custom_shows: CustomShowList,
    pub sections: SectionList,
}

/// Load and validate the presentation structure graph.
pub fn load_presentation_structure(package: &OpcPackage) -> Result<PresentationStructure> {
    let presentation = package.main_document_part()?;
    require_presentation(presentation.content_type())?;
    parse_structure_blob(package, presentation.blob(), true)
}

/// Atomically replace custom shows and sections while preserving unrelated XML.
pub fn store_presentation_structure(
    package: &mut OpcPackage,
    value: &PresentationStructure,
) -> Result<()> {
    validate_graph(package, value)?;
    let presentation_name = package.main_document_part()?.partname().clone();
    let original = package.get_part(&presentation_name)?.blob().to_vec();
    let (p_namespace, r_namespace) = document_namespaces(&original);
    let custom_xml = write_custom_shows(&value.custom_shows, &value.slides, p_namespace, r_namespace)?;
    let mut staged = patch_custom_shows(&original, custom_xml.as_bytes())?;
    let section_xml = write_section_extension(&value.sections, p_namespace)?;
    staged = patch_sections(&staged, section_xml.as_bytes(), p_namespace)?;

    let reparsed = parse_structure_blob(package, &staged, true)?;
    if reparsed.slides != value.slides
        || reparsed.custom_shows.shows != value.custom_shows.shows
        || reparsed.sections != value.sections
    {
        return Err(invalid("staged presentation structure did not round-trip"));
    }
    package
        .clear_digital_signatures()
        .map_err(|error| OoxmlError::Other(format!("cannot invalidate package signatures: {error}")))?;
    package.get_part_mut(&presentation_name)?.set_blob(staged);
    Ok(())
}

pub fn find_custom_show(package: &OpcPackage, id: u32) -> Result<Option<CustomShow>> {
    Ok(load_presentation_structure(package)?.custom_shows.get_by_id(id).cloned())
}

pub fn add_custom_show(package: &mut OpcPackage, show: CustomShow) -> Result<()> {
    mutate(package, |graph| {
        if graph.custom_shows.get_by_id(show.id).is_some()
            || graph.custom_shows.get_by_name(&show.name).is_some()
        {
            return Err(invalid("duplicate custom-show ID or name"));
        }
        graph.custom_shows.add(show);
        Ok(())
    })
}

pub fn update_custom_show(
    package: &mut OpcPackage,
    id: u32,
    replacement: CustomShow,
) -> Result<()> {
    mutate(package, |graph| graph.custom_shows.replace_by_id(id, replacement))
}

pub fn replace_custom_show(
    package: &mut OpcPackage,
    id: u32,
    replacement: CustomShow,
) -> Result<()> {
    update_custom_show(package, id, replacement)
}

pub fn remove_custom_show(package: &mut OpcPackage, id: u32) -> Result<bool> {
    let mut graph = load_presentation_structure(package)?;
    let removed = graph.custom_shows.remove_by_id(id).is_some();
    if removed {
        store_presentation_structure(package, &graph)?;
    }
    Ok(removed)
}

pub fn reorder_custom_shows(package: &mut OpcPackage, ordered_ids: &[u32]) -> Result<()> {
    mutate(package, |graph| graph.custom_shows.reorder(ordered_ids))
}

pub fn add_custom_show_slide(package: &mut OpcPackage, show_id: u32, slide_id: u32) -> Result<()> {
    mutate(package, |graph| {
        require_slide(graph, slide_id)?;
        let show = graph
            .custom_shows
            .get_by_id_mut(show_id)
            .ok_or_else(|| invalid(format!("custom show {show_id} was not found")))?;
        if show.slide_ids.contains(&slide_id) {
            return Err(invalid("duplicate custom-show slide membership"));
        }
        show.slide_ids.push(slide_id);
        Ok(())
    })
}

pub fn remove_custom_show_slide(
    package: &mut OpcPackage,
    show_id: u32,
    slide_id: u32,
) -> Result<bool> {
    let mut graph = load_presentation_structure(package)?;
    let show = graph
        .custom_shows
        .get_by_id_mut(show_id)
        .ok_or_else(|| invalid(format!("custom show {show_id} was not found")))?;
    let Some(offset) = show.slide_ids.iter().position(|id| *id == slide_id) else {
        return Ok(false);
    };
    show.slide_ids.remove(offset);
    store_presentation_structure(package, &graph)?;
    Ok(true)
}

pub fn reorder_custom_show_slides(
    package: &mut OpcPackage,
    show_id: u32,
    ordered_slide_ids: &[u32],
) -> Result<()> {
    mutate(package, |graph| {
        let show = graph
            .custom_shows
            .get_by_id_mut(show_id)
            .ok_or_else(|| invalid(format!("custom show {show_id} was not found")))?;
        require_permutation(&show.slide_ids, ordered_slide_ids, "custom-show slide")?;
        show.slide_ids = ordered_slide_ids.to_vec();
        Ok(())
    })
}

pub fn find_section(package: &OpcPackage, id: &str) -> Result<Option<Section>> {
    Ok(load_presentation_structure(package)?.sections.get_by_id(id).cloned())
}

pub fn add_section(package: &mut OpcPackage, mut section: Section) -> Result<String> {
    if section.id.is_none() {
        section.id = Some(litchi_core::id::generate_guid_braced());
    }
    let id = section.id.clone().expect("section ID was assigned");
    let retained = id.clone();
    mutate(package, move |graph| {
        if graph.sections.get_by_id(&id).is_some() {
            return Err(invalid(format!("section {id} already exists")));
        }
        graph.sections.add_section(section);
        graph.sections.sort_slide_membership(
            &graph.slides.iter().map(|slide| slide.slide_id).collect::<Vec<_>>(),
        );
        Ok(())
    })?;
    Ok(retained)
}

pub fn update_section(package: &mut OpcPackage, id: &str, replacement: Section) -> Result<()> {
    let id = id.to_owned();
    mutate(package, move |graph| {
        graph.sections.replace_by_id(&id, replacement)?;
        graph.sections.sort_slide_membership(
            &graph.slides.iter().map(|slide| slide.slide_id).collect::<Vec<_>>(),
        );
        Ok(())
    })
}

pub fn replace_section(package: &mut OpcPackage, id: &str, replacement: Section) -> Result<()> {
    update_section(package, id, replacement)
}

pub fn remove_section(package: &mut OpcPackage, id: &str) -> Result<bool> {
    let mut graph = load_presentation_structure(package)?;
    let removed = graph.sections.remove_by_id(id).is_some();
    if removed {
        store_presentation_structure(package, &graph)?;
    }
    Ok(removed)
}

pub fn reorder_sections(package: &mut OpcPackage, ordered_ids: &[String]) -> Result<()> {
    mutate(package, |graph| graph.sections.reorder(ordered_ids))
}

pub fn add_section_slide(package: &mut OpcPackage, section_id: &str, slide_id: u32) -> Result<()> {
    let section_id = section_id.to_owned();
    mutate(package, move |graph| {
        require_slide(graph, slide_id)?;
        if graph
            .sections
            .sections()
            .iter()
            .any(|section| section.slide_ids.contains(&slide_id))
        {
            return Err(invalid("slide already belongs to a section"));
        }
        graph
            .sections
            .get_by_id_mut(&section_id)
            .ok_or_else(|| invalid(format!("section {section_id} was not found")))?
            .slide_ids
            .push(slide_id);
        graph.sections.sort_slide_membership(
            &graph.slides.iter().map(|slide| slide.slide_id).collect::<Vec<_>>(),
        );
        Ok(())
    })
}

pub fn remove_section_slide(
    package: &mut OpcPackage,
    section_id: &str,
    slide_id: u32,
) -> Result<bool> {
    let mut graph = load_presentation_structure(package)?;
    let section = graph
        .sections
        .get_by_id_mut(section_id)
        .ok_or_else(|| invalid(format!("section {section_id} was not found")))?;
    let Some(offset) = section.slide_ids.iter().position(|id| *id == slide_id) else {
        return Ok(false);
    };
    section.slide_ids.remove(offset);
    store_presentation_structure(package, &graph)?;
    Ok(true)
}

pub fn reorder_section_slides(
    package: &mut OpcPackage,
    section_id: &str,
    ordered_slide_ids: &[u32],
) -> Result<()> {
    let section_id = section_id.to_owned();
    mutate(package, move |graph| {
        let section = graph
            .sections
            .get_by_id_mut(&section_id)
            .ok_or_else(|| invalid(format!("section {section_id} was not found")))?;
        require_permutation(&section.slide_ids, ordered_slide_ids, "section slide")?;
        section.slide_ids = ordered_slide_ids.to_vec();
        Ok(())
    })
}

/// Reconcile memberships after another API has inserted, removed, or reordered slides.
///
/// Deleted slide references are removed, custom-show order is retained, and section
/// memberships are sorted to the current `p:sldIdLst` order.
pub fn synchronize_presentation_structure_after_slide_mutation(
    package: &mut OpcPackage,
) -> Result<()> {
    let presentation = package.main_document_part()?;
    let mut graph = parse_structure_blob(package, presentation.blob(), false)?;
    let live = graph.slides.iter().map(|slide| slide.slide_id).collect::<HashSet<_>>();
    for show in &mut graph.custom_shows.shows {
        show.slide_ids.retain(|id| live.contains(id));
    }
    for section in graph.sections.sections_mut() {
        section.slide_ids.retain(|id| live.contains(id));
    }
    graph.sections.sort_slide_membership(
        &graph.slides.iter().map(|slide| slide.slide_id).collect::<Vec<_>>(),
    );
    store_presentation_structure(package, &graph)
}

fn mutate<F>(package: &mut OpcPackage, operation: F) -> Result<()>
where
    F: FnOnce(&mut PresentationStructure) -> Result<()>,
{
    let mut graph = load_presentation_structure(package)?;
    operation(&mut graph)?;
    store_presentation_structure(package, &graph)
}

fn parse_structure_blob(
    package: &OpcPackage,
    xml: &[u8],
    strict_references: bool,
) -> Result<PresentationStructure> {
    if xml.len() > MAX_BYTES {
        return Err(invalid("presentation structure exceeds 8 MiB"));
    }
    let processed = crate::common::mce::process_ooxml(xml)?;
    let (raw_slides, raw_shows) = parse_core(processed.as_ref())?;
    let presentation = package.main_document_part()?;
    let mut slides = Vec::with_capacity(raw_slides.len());
    for (slide_id, relationship_id) in raw_slides {
        let relationship = presentation.rels().get(&relationship_id).ok_or_else(|| {
            invalid(format!("orphan presentation slide relationship {relationship_id}"))
        })?;
        if relationship.is_external()
            || !matches!(relationship.reltype(), SLIDE_REL | SLIDE_REL_STRICT)
        {
            return Err(invalid(format!(
                "relationship {relationship_id} is not an internal slide relationship"
            )));
        }
        let target = relationship.target_partname()?;
        let part = package.get_part(&target)?;
        if part.content_type()
            != "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
        {
            return Err(invalid(format!("relationship {relationship_id} targets a non-slide part")));
        }
        slides.push(PresentationSlideReference {
            slide_id,
            relationship_id,
            part_name: target.to_string(),
        });
    }
    let rel_to_id = slides
        .iter()
        .map(|slide| (slide.relationship_id.as_str(), slide.slide_id))
        .collect::<HashMap<_, _>>();
    let mut custom_shows = CustomShowList::new();
    for raw in raw_shows {
        let mut show = CustomShow::new(raw.id, raw.name);
        for relationship_id in raw.relationship_ids {
            if let Some(slide_id) = rel_to_id.get(relationship_id.as_str()) {
                show.slide_ids.push(*slide_id);
            } else if strict_references {
                return Err(invalid(format!(
                    "custom show {} has orphan slide relationship {relationship_id}",
                    show.id
                )));
            }
        }
        custom_shows.add(show);
    }
    let graph = PresentationStructure {
        slides,
        custom_shows,
        sections: SectionList::from_xml(xml)?,
    };
    if strict_references {
        validate_graph(package, &graph)?;
    }
    Ok(graph)
}

#[derive(Default)]
struct RawShow {
    id: u32,
    name: String,
    relationship_ids: Vec<String>,
}

#[allow(clippy::type_complexity)]
fn parse_core(xml: &[u8]) -> Result<(Vec<(u32, String)>, Vec<RawShow>)> {
    let mut reader = Reader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut ancestors = Vec::<String>::new();
    let mut slides = Vec::new();
    let mut shows = Vec::new();
    let mut current_show: Option<RawShow> = None;
    let mut nodes = 0usize;
    loop {
        let decoder = reader.decoder();
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes += 1;
                resource(nodes, ancestors.len())?;
                let local = local_name(element.name().as_ref())?;
                if local == "custShow" {
                    if current_show.is_some() {
                        return Err(invalid("nested custom shows are rejected"));
                    }
                    current_show = Some(parse_show(&element, decoder)?);
                }
                ancestors.push(local);
            },
            Ok(Event::Empty(element)) => {
                nodes += 1;
                resource(nodes, ancestors.len())?;
                let local = local_name(element.name().as_ref())?;
                if local == "sldId"
                    && ancestors.len() == 2
                    && ancestors.last().map(String::as_str) == Some("sldIdLst")
                {
                    slides.push(parse_slide_reference(&element, decoder)?);
                } else if local == "sld"
                    && ancestors.last().map(String::as_str) == Some("sldLst")
                {
                    let relationship_id = relationship_id(&element, decoder)?;
                    current_show
                        .as_mut()
                        .ok_or_else(|| invalid("custom-show slide appears outside a custom show"))?
                        .relationship_ids
                        .push(relationship_id);
                }
            },
            Ok(Event::End(element)) => {
                let local = local_name(element.name().as_ref())?;
                let open = ancestors.pop().ok_or_else(|| invalid("unexpected closing element"))?;
                if open != local {
                    return Err(invalid("mismatched presentation element"));
                }
                if local == "custShow" {
                    shows.push(current_show.take().ok_or_else(|| invalid("missing custom show"))?);
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
    if !ancestors.is_empty() || current_show.is_some() {
        return Err(invalid("unterminated presentation structure"));
    }
    Ok((slides, shows))
}

fn parse_show(element: &BytesStart<'_>, decoder: quick_xml::encoding::Decoder) -> Result<RawShow> {
    let attributes = attributes(element, decoder)?;
    let id = required(&attributes, "id")?
        .parse::<u32>()
        .map_err(|_| invalid("invalid custom-show ID"))?;
    let name = required(&attributes, "name")?.to_owned();
    if name.is_empty() {
        return Err(invalid("custom-show name cannot be empty"));
    }
    Ok(RawShow { id, name, relationship_ids: Vec::new() })
}

fn parse_slide_reference(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<(u32, String)> {
    let attributes = attributes(element, decoder)?;
    let id = required_unqualified(&attributes, "id")?
        .parse::<u32>()
        .map_err(|_| invalid("invalid presentation slide ID"))?;
    if id < 256 {
        return Err(invalid("presentation slide ID is below 256"));
    }
    Ok((id, required_qualified(&attributes, "id")?.to_owned()))
}

fn relationship_id(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<String> {
    let attributes = attributes(element, decoder)?;
    Ok(required_qualified(&attributes, "id")?.to_owned())
}

fn attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<(String, String)>> {
    let mut values = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| OoxmlError::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .to_owned();
        if !seen.insert(name.clone()) {
            return Err(invalid("duplicate presentation XML attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| OoxmlError::Xml(error.to_string()))?
            .into_owned();
        values.push((name, value));
    }
    Ok(values)
}

fn required<'a>(values: &'a [(String, String)], local: &str) -> Result<&'a str> {
    values
        .iter()
        .find(|(name, _)| name == local)
        .map(|(_, value)| value.as_str())
        .ok_or_else(|| invalid(format!("missing attribute '{local}'")))
}

fn required_unqualified<'a>(values: &'a [(String, String)], local: &str) -> Result<&'a str> {
    required(values, local)
}

fn required_qualified<'a>(values: &'a [(String, String)], local: &str) -> Result<&'a str> {
    values
        .iter()
        .find(|(name, _)| name.rsplit_once(':').map(|(_, item)| item) == Some(local))
        .map(|(_, value)| value.as_str())
        .ok_or_else(|| invalid(format!("missing relationship attribute '{local}'")))
}

fn validate_graph(package: &OpcPackage, graph: &PresentationStructure) -> Result<()> {
    let mut slide_ids = HashSet::new();
    let mut rel_ids = HashSet::new();
    let mut part_names = HashSet::new();
    for slide in &graph.slides {
        if slide.slide_id < 256 || !slide_ids.insert(slide.slide_id) {
            return Err(invalid("invalid or duplicate presentation slide ID"));
        }
        if !rel_ids.insert(slide.relationship_id.as_str())
            || !part_names.insert(slide.part_name.as_str())
        {
            return Err(invalid("duplicate presentation slide relationship or target"));
        }
    }
    let presentation = package.main_document_part()?;
    for slide in &graph.slides {
        let relationship = presentation
            .rels()
            .get(&slide.relationship_id)
            .ok_or_else(|| invalid(format!("orphan slide relationship {}", slide.relationship_id)))?;
        if relationship.is_external()
            || !matches!(relationship.reltype(), SLIDE_REL | SLIDE_REL_STRICT)
            || relationship.target_partname()?.as_str() != slide.part_name
        {
            return Err(invalid("presentation slide relationship mismatch"));
        }
    }
    let mut show_ids = HashSet::new();
    let mut show_names = HashSet::new();
    for show in &graph.custom_shows.shows {
        if show.name.is_empty()
            || !show_ids.insert(show.id)
            || !show_names.insert(show.name.as_str())
        {
            return Err(invalid("duplicate or empty custom-show identity"));
        }
        let mut membership = HashSet::new();
        for slide_id in &show.slide_ids {
            if !slide_ids.contains(slide_id) || !membership.insert(*slide_id) {
                return Err(invalid("duplicate or orphan custom-show slide reference"));
            }
        }
    }
    let positions = graph
        .slides
        .iter()
        .enumerate()
        .map(|(offset, slide)| (slide.slide_id, offset))
        .collect::<HashMap<_, _>>();
    let mut section_ids = HashSet::new();
    let mut section_membership = HashSet::new();
    for section in graph.sections.sections() {
        let id = section
            .id
            .as_deref()
            .ok_or_else(|| invalid("mutable presentation section requires a stable GUID"))?;
        if !section_ids.insert(id) {
            return Err(invalid("duplicate section GUID"));
        }
        let mut previous = None;
        for slide_id in &section.slide_ids {
            let position = positions
                .get(slide_id)
                .copied()
                .ok_or_else(|| invalid("orphan section slide reference"))?;
            if !section_membership.insert(*slide_id) {
                return Err(invalid("slide belongs to more than one section"));
            }
            if previous.is_some_and(|value| value >= position) {
                return Err(invalid("section membership is not in presentation order"));
            }
            previous = Some(position);
        }
    }
    graph.sections.to_xml()?;
    Ok(())
}

fn write_custom_shows(
    shows: &CustomShowList,
    slides: &[PresentationSlideReference],
    p_namespace: &str,
    r_namespace: &str,
) -> Result<String> {
    if shows.is_empty() {
        return Ok(String::new());
    }
    let relationships = slides
        .iter()
        .map(|slide| (slide.slide_id, slide.relationship_id.as_str()))
        .collect::<HashMap<_, _>>();
    let mut xml = format!(
        "<p:custShowLst xmlns:p=\"{}\" xmlns:r=\"{}\">",
        p_namespace, r_namespace
    );
    for show in &shows.shows {
        xml.push_str("<p:custShow name=\"");
        xml.push_str(&escape_xml(&show.name));
        xml.push_str("\" id=\"");
        xml.push_str(&show.id.to_string());
        xml.push_str("\"><p:sldLst>");
        for slide_id in &show.slide_ids {
            let relationship = relationships
                .get(slide_id)
                .ok_or_else(|| invalid("custom show references an undeclared slide"))?;
            xml.push_str("<p:sld r:id=\"");
            xml.push_str(&escape_xml(relationship));
            xml.push_str("\"/>");
        }
        xml.push_str("</p:sldLst></p:custShow>");
    }
    xml.push_str("</p:custShowLst>");
    if xml.len() > MAX_BYTES {
        return Err(invalid("serialized custom shows exceed 8 MiB"));
    }
    Ok(xml)
}

fn write_section_extension(sections: &SectionList, p_namespace: &str) -> Result<String> {
    let full = sections.to_xml()?;
    if full.is_empty() {
        return Ok(String::new());
    }
    let inner = full
        .strip_prefix("<p:extLst>")
        .and_then(|value| value.strip_suffix("</p:extLst>"))
        .ok_or_else(|| invalid("invalid serialized section extension"))?;
    Ok(inner.replacen(
        "<p:ext ",
        &format!("<p:ext xmlns:p=\"{p_namespace}\" "),
        1,
    ))
}

#[derive(Clone)]
struct Span {
    start: usize,
    end: usize,
    close_start: usize,
    local: String,
}

struct Frame {
    start: usize,
    local: String,
    direct: bool,
    target_custom: bool,
    target_section: bool,
}

struct XmlLayout {
    root_close: usize,
    direct: Vec<Span>,
    custom: Vec<Span>,
    section_extensions: Vec<Span>,
}

fn scan_layout(xml: &[u8]) -> Result<XmlLayout> {
    let mut reader = Reader::from_reader(xml);
    let mut stack = Vec::<Frame>::new();
    let mut direct = Vec::new();
    let mut custom = Vec::new();
    let mut section_extensions = Vec::new();
    let mut root_close = None;
    let mut nodes = 0usize;
    loop {
        let before = reader.buffer_position() as usize;
        let decoder = reader.decoder();
        match reader.read_event() {
            Ok(Event::Start(element)) => {
                nodes += 1;
                resource(nodes, stack.len())?;
                let local = local_name(element.name().as_ref())?;
                let target_section = local == "ext"
                    && attributes(&element, decoder)?
                        .iter()
                        .any(|(name, value)| name == "uri" && value == SECTION_URI);
                stack.push(Frame {
                    start: before,
                    target_custom: local == "custShowLst",
                    target_section,
                    direct: stack.len() == 1,
                    local,
                });
            },
            Ok(Event::Empty(element)) => {
                nodes += 1;
                resource(nodes, stack.len())?;
                let local = local_name(element.name().as_ref())?;
                let span = Span {
                    start: before,
                    end: reader.buffer_position() as usize,
                    close_start: before,
                    local: local.clone(),
                };
                if stack.len() == 1 {
                    direct.push(span.clone());
                }
                if local == "custShowLst" {
                    custom.push(span.clone());
                }
                if local == "ext"
                    && attributes(&element, decoder)?
                        .iter()
                        .any(|(name, value)| name == "uri" && value == SECTION_URI)
                {
                    section_extensions.push(span);
                }
            },
            Ok(Event::End(_)) => {
                let frame = stack.pop().ok_or_else(|| invalid("unexpected closing element"))?;
                let span = Span {
                    start: frame.start,
                    end: reader.buffer_position() as usize,
                    close_start: before,
                    local: frame.local,
                };
                if frame.direct {
                    direct.push(span.clone());
                }
                if frame.target_custom {
                    custom.push(span.clone());
                }
                if frame.target_section {
                    section_extensions.push(span.clone());
                }
                if stack.is_empty() {
                    root_close = Some(before);
                }
            },
            Ok(Event::DocType(_) | Event::PI(_)) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Ok(Event::Eof) => break,
            Err(error) => return Err(OoxmlError::Xml(error.to_string())),
            _ => {},
        }
    }
    if !stack.is_empty() {
        return Err(invalid("unterminated presentation XML"));
    }
    direct.sort_by_key(|span| span.start);
    Ok(XmlLayout {
        root_close: root_close.ok_or_else(|| invalid("missing presentation root close"))?,
        direct,
        custom,
        section_extensions,
    })
}

fn patch_custom_shows(xml: &[u8], replacement: &[u8]) -> Result<Vec<u8>> {
    let layout = scan_layout(xml)?;
    if !layout.custom.is_empty() {
        return replace_spans(xml, &layout.custom, replacement);
    }
    if replacement.is_empty() {
        return Ok(xml.to_vec());
    }
    let insert = layout
        .direct
        .iter()
        .find(|span| schema_rank(&span.local).is_some_and(|rank| rank > 8))
        .map(|span| span.start)
        .unwrap_or(layout.root_close);
    insert_bytes(xml, insert, replacement)
}

fn patch_sections(xml: &[u8], replacement: &[u8], p_namespace: &str) -> Result<Vec<u8>> {
    let layout = scan_layout(xml)?;
    if !layout.section_extensions.is_empty() {
        return replace_spans(xml, &layout.section_extensions, replacement);
    }
    if replacement.is_empty() {
        return Ok(xml.to_vec());
    }
    if let Some(ext_list) = layout.direct.iter().find(|span| span.local == "extLst") {
        if ext_list.close_start == ext_list.start {
            let wrapped = format!(
                "<p:extLst xmlns:p=\"{p_namespace}\">{}</p:extLst>",
                std::str::from_utf8(replacement)
                    .map_err(|error| OoxmlError::Xml(error.to_string()))?
            );
            return replace_spans(xml, std::slice::from_ref(ext_list), wrapped.as_bytes());
        }
        return insert_bytes(xml, ext_list.close_start, replacement);
    }
    let wrapped = format!(
        "<p:extLst xmlns:p=\"{p_namespace}\">{}</p:extLst>",
        std::str::from_utf8(replacement).map_err(|error| OoxmlError::Xml(error.to_string()))?
    );
    insert_bytes(xml, layout.root_close, wrapped.as_bytes())
}

fn replace_spans(xml: &[u8], spans: &[Span], replacement: &[u8]) -> Result<Vec<u8>> {
    let mut ordered = spans.to_vec();
    ordered.sort_by_key(|span| span.start);
    for pair in ordered.windows(2) {
        if pair[0].end > pair[1].start {
            return Err(invalid("overlapping presentation XML patch ranges"));
        }
    }
    let mut output = xml.to_vec();
    for span in ordered.iter().rev() {
        output.splice(span.start..span.end, replacement.iter().copied());
    }
    if output.len() > MAX_BYTES {
        return Err(invalid("patched presentation structure exceeds 8 MiB"));
    }
    Ok(output)
}

fn insert_bytes(xml: &[u8], offset: usize, value: &[u8]) -> Result<Vec<u8>> {
    let mut output = Vec::with_capacity(xml.len().saturating_add(value.len()));
    output.extend_from_slice(&xml[..offset]);
    output.extend_from_slice(value);
    output.extend_from_slice(&xml[offset..]);
    if output.len() > MAX_BYTES {
        return Err(invalid("patched presentation structure exceeds 8 MiB"));
    }
    Ok(output)
}

fn schema_rank(local: &str) -> Option<usize> {
    [
        "sldMasterIdLst", "notesMasterIdLst", "handoutMasterIdLst", "sldIdLst", "sldSz",
        "notesSz", "smartTags", "embeddedFontLst", "custShowLst", "photoAlbum", "custDataLst",
        "kinsoku", "defaultTextStyle", "modifyVerifier", "extLst",
    ]
    .iter()
    .position(|name| *name == local)
}

fn document_namespaces(xml: &[u8]) -> (&'static str, &'static str) {
    if xml.windows(PS.len()).any(|window| window == PS.as_bytes()) {
        (PS, RS)
    } else {
        (P, R)
    }
}

fn require_slide(graph: &PresentationStructure, slide_id: u32) -> Result<()> {
    if graph.slides.iter().any(|slide| slide.slide_id == slide_id) {
        Ok(())
    } else {
        Err(invalid(format!("slide {slide_id} was not found")))
    }
}

fn require_permutation(expected: &[u32], actual: &[u32], what: &str) -> Result<()> {
    let expected_set = expected.iter().copied().collect::<HashSet<_>>();
    let actual_set = actual.iter().copied().collect::<HashSet<_>>();
    if expected_set != actual_set || expected.len() != actual.len() {
        Err(invalid(format!("{what} reorder is not a permutation")))
    } else {
        Ok(())
    }
}

fn require_presentation(content_type: &str) -> Result<()> {
    if matches!(
        content_type,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
            | "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"
    ) {
        Ok(())
    } else {
        Err(invalid("main document is not a PowerPoint presentation"))
    }
}

fn resource(nodes: usize, depth: usize) -> Result<()> {
    if nodes > MAX_NODES || depth >= MAX_DEPTH {
        Err(invalid("presentation structure XML resource limit exceeded"))
    } else {
        Ok(())
    }
}

fn local_name(name: &[u8]) -> Result<String> {
    let name = std::str::from_utf8(name).map_err(|error| OoxmlError::Xml(error.to_string()))?;
    Ok(name.rsplit_once(':').map_or(name, |(_, local)| local).to_owned())
}

fn invalid(message: impl Into<String>) -> OoxmlError {
    OoxmlError::InvalidFormat(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;
    use litchi_opc::constants::relationship_type;
    use litchi_opc::packuri::PackURI;
    use litchi_opc::part::BlobPart;

    fn package_with_slides() -> crate::pptx::Package {
        let mut package = crate::pptx::Package::new().unwrap();
        let opc = package.opc_package_mut();
        let presentation_name = opc.main_document_part().unwrap().partname().clone();
        let mut xml = std::str::from_utf8(opc.get_part(&presentation_name).unwrap().blob())
            .unwrap()
            .to_owned();
        let marker = "<p:sldSz";
        let offset = xml.find(marker).unwrap();
        xml.insert_str(
            offset,
            "<!--preserve--><p:sldIdLst><p:sldId id=\"256\" r:id=\"rIdSlideA\"/><p:sldId id=\"300\" r:id=\"slide-beta\"/></p:sldIdLst>",
        );
        opc.get_part_mut(&presentation_name)
            .unwrap()
            .set_blob(xml.into_bytes());
        opc.get_part_mut(&presentation_name)
            .unwrap()
            .rels_mut()
            .add_relationship(
                relationship_type::SLIDE.into(),
                "slides/slide1.xml".into(),
                "rIdSlideA".into(),
                false,
            );
        opc.get_part_mut(&presentation_name)
            .unwrap()
            .rels_mut()
            .add_relationship(
                relationship_type::SLIDE.into(),
                "slides/slide2.xml".into(),
                "slide-beta".into(),
                false,
            );
        for index in 1..=2 {
            let uri = PackURI::new(&format!("/ppt/slides/slide{index}.xml")).unwrap();
            opc.add_part(Box::new(BlobPart::new(
                uri,
                "application/vnd.openxmlformats-officedocument.presentationml.slide+xml".into(),
                b"<p:sld xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:cSld/><p:clrMapOvr><a:masterClrMapping xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"/></p:clrMapOvr></p:sld>".to_vec(),
            )));
        }
        package
    }

    #[test]
    fn generated_package_crud_preserves_relationship_ids_and_unknown_xml() {
        let mut package = package_with_slides();
        add_custom_show(
            package.opc_package_mut(),
            CustomShow::new(7, "Roadshow").with_slides(vec![300, 256]),
        )
        .unwrap();
        let section_id = add_section(
            package.opc_package_mut(),
            Section::new("Opening", "{11111111-1111-1111-1111-111111111111}")
                .with_slides([256, 300]),
        )
        .unwrap();
        assert_eq!(section_id, "{11111111-1111-1111-1111-111111111111}");
        let graph = load_presentation_structure(package.opc_package()).unwrap();
        assert_eq!(graph.custom_shows.get_by_id(7).unwrap().slide_ids, vec![300, 256]);
        assert_eq!(graph.sections.get_by_id(&section_id).unwrap().slide_ids, vec![256, 300]);
        let xml = std::str::from_utf8(package.opc_package().main_document_part().unwrap().blob())
            .unwrap();
        assert!(xml.contains("r:id=\"slide-beta\""));
        assert!(xml.contains("<!--preserve-->"));

        reorder_custom_show_slides(package.opc_package_mut(), 7, &[256, 300]).unwrap();
        remove_custom_show_slide(package.opc_package_mut(), 7, 300).unwrap();
        remove_section_slide(package.opc_package_mut(), &section_id, 300).unwrap();
        assert_eq!(find_custom_show(package.opc_package(), 7).unwrap().unwrap().slide_ids, vec![256]);
        assert_eq!(find_section(package.opc_package(), &section_id).unwrap().unwrap().slide_ids, vec![256]);
        assert!(remove_custom_show(package.opc_package_mut(), 7).unwrap());
        assert!(remove_section(package.opc_package_mut(), &section_id).unwrap());
    }

    #[test]
    fn orphan_and_non_permutation_mutations_are_atomic() {
        let mut package = package_with_slides();
        add_custom_show(
            package.opc_package_mut(),
            CustomShow::new(9, "Demo").with_slides(vec![256, 300]),
        )
        .unwrap();
        let before = package.opc_package().main_document_part().unwrap().blob().to_vec();
        assert!(reorder_custom_show_slides(package.opc_package_mut(), 9, &[256, 256]).is_err());
        assert_eq!(package.opc_package().main_document_part().unwrap().blob(), before);
        assert!(add_custom_show(
            package.opc_package_mut(),
            CustomShow::new(10, "Broken").with_slides(vec![999]),
        )
        .is_err());
        assert_eq!(package.opc_package().main_document_part().unwrap().blob(), before);
    }

    #[test]
    fn mutable_slide_delete_and_move_keep_memberships_coherent() {
        let mut presentation = crate::pptx::MutablePresentation::new();
        presentation.add_slide().unwrap();
        presentation.add_slide().unwrap();
        presentation.add_slide().unwrap();
        presentation.add_section("All", vec![256, 257, 258]);
        presentation.create_custom_show("All", vec![256, 257, 258]);

        presentation.delete_slide(1).unwrap();
        assert_eq!(presentation.sections().sections()[0].slide_ids, vec![256, 258]);
        assert_eq!(presentation.custom_shows().shows[0].slide_ids, vec![256, 258]);

        presentation.move_slide(1, 0).unwrap();
        assert_eq!(presentation.sections().sections()[0].slide_ids, vec![258, 256]);
        assert_eq!(presentation.custom_shows().shows[0].slide_ids, vec![256, 258]);
        assert_eq!(presentation.add_slide().unwrap().slide_id(), 259);
    }

    #[test]
    fn malformed_xml_is_rejected_and_mce_branches_patch_without_overlap() {
        assert!(parse_core(b"<!DOCTYPE x><p:presentation/>").is_err());
        assert!(parse_core(b"<p:presentation><p:sldIdLst>").is_err());

        let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Choice Requires="p"><p:custShowLst><p:custShow name="old" id="1"><p:sldLst/></p:custShow></p:custShowLst></mc:Choice><mc:Fallback><p:custShowLst/></mc:Fallback></mc:AlternateContent><!--keep--></p:presentation>"#;
        let replacement = br#"<p:custShowLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>"#;
        let patched = patch_custom_shows(xml, replacement).unwrap();
        assert_eq!(
            patched
                .windows(replacement.len())
                .filter(|window| *window == replacement)
                .count(),
            2
        );
        assert!(patched.windows(b"<!--keep-->".len()).any(|window| window == b"<!--keep-->"));
    }
}
