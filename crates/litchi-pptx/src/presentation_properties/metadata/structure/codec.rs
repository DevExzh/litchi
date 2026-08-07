//! Package-aware custom-show and section mutation for PowerPoint presentations.
//!
//! Custom shows reference slide relationships while PowerPoint 2010 sections
//! reference stable numeric presentation slide IDs. This module resolves both
//! representations into one validated graph and patches only the corresponding
//! children of `ppt/presentation.xml`.

use super::super::custom_show::{List as ShowList, Show};
use super::super::sections::{List as SectionList, Section};
use super::model::*;
use crate::presentation_properties::metadata::{escape_xml, new_guid};
use crate::{Error, Result};
use litchi_opc::OpcPackage;
use quick_xml::events::{BytesStart, Event};
use quick_xml::name::{Namespace, QName, ResolveResult};
use quick_xml::reader::NsReader;
use quick_xml::{Reader, XmlVersion};
use std::collections::{HashMap, HashSet};

const P: &str = "http://schemas.openxmlformats.org/presentationml/2006/main";
const PS: &str = "http://purl.oclc.org/ooxml/presentationml/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
const SLIDE_REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";
const SLIDE_REL_STRICT: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships/slide";
const SECTION_URI: &str = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}";
const MAX_BYTES: usize = 8 * 1024 * 1024;
const MAX_NODES: usize = 100_000;
const MAX_DEPTH: usize = 128;
const MAX_SLIDE_ID: u32 = 2_147_483_647;

/// Load and validate the presentation structure graph.
pub fn load(package: &OpcPackage) -> Result<Graph> {
    let presentation = package.main_document_part()?;
    require_presentation(presentation.content_type())?;
    parse_structure_blob(package, presentation.blob(), true)
}

/// Atomically replace custom shows and sections while preserving unrelated XML.
pub fn store(package: &mut OpcPackage, value: &Graph) -> Result<()> {
    validate_graph(package, value)?;
    let presentation_name = package.main_document_part()?.partname().clone();
    let original = package.get_part(&presentation_name)?.blob().to_vec();
    let original_graph = parse_structure_blob(package, &original, true)?;
    let staged = rewrite(&original, &original_graph, value)?;

    if staged == original {
        return Ok(());
    }

    let reparsed = parse_structure_blob(package, &staged, true)?;
    if reparsed.slides != value.slides
        || reparsed.custom_shows.shows != value.custom_shows.shows
        || reparsed.sections != value.sections
    {
        return Err(invalid("staged presentation structure did not round-trip"));
    }
    package.unsign();
    package.get_part_mut(&presentation_name)?.set_blob(staged);
    Ok(())
}

/// Rewrite only the changed structure owners in one validated presentation
/// document. The slide relationship graph is intentionally immutable here;
/// slide insertion and removal belong to the presentation package facade.
pub(crate) fn rewrite(source: &[u8], original: &Graph, value: &Graph) -> Result<Vec<u8>> {
    validate_detached_graph(original, value)?;
    if equivalent_graph(original, value) {
        return Ok(source.to_vec());
    }

    let (p_namespace, r_namespace) = document_namespaces(source);
    let mut staged = source.to_vec();
    if original.custom_shows.shows != value.custom_shows.shows {
        let custom_xml =
            write_custom_shows(&value.custom_shows, &value.slides, p_namespace, r_namespace)?;
        staged = patch_custom_shows(&staged, custom_xml.as_bytes())?;
    }
    if original.sections != value.sections {
        let section_xml = write_section_extension(&value.sections, p_namespace)?;
        staged = patch_sections(&staged, section_xml.as_bytes(), p_namespace)?;
    }

    let reparsed = parse_detached(&staged, &value.slides)?;
    if !equivalent_graph(&reparsed, value) {
        return Err(invalid("staged presentation structure did not round-trip"));
    }
    Ok(staged)
}

/// Parse a rewritten presentation against its already validated relationship
/// topology without needing to reopen any slide part.
pub(crate) fn parse_detached(xml: &[u8], slides: &[Reference]) -> Result<Graph> {
    if xml.len() > MAX_BYTES {
        return Err(invalid("presentation structure exceeds 8 MiB"));
    }
    let processed = litchi_ooxml_common::mce::process_ooxml(xml)?;
    let (raw_slides, raw_shows) = parse_core(processed.as_ref())?;
    let expected_slides = slides
        .iter()
        .map(|slide| (slide.slide_id, slide.relationship_id.clone()))
        .collect::<Vec<_>>();
    if raw_slides != expected_slides {
        return Err(invalid(
            "presentation slide relationship topology changed in a structure transaction",
        ));
    }
    let custom_shows = resolve_custom_shows(raw_shows, slides, true)?;
    let graph = Graph {
        slides: slides.to_vec(),
        custom_shows,
        sections: SectionList::from_xml(xml)?,
    };
    validate_graph_shape(&graph)?;
    Ok(graph)
}

/// Validate detached edits against the source's immutable slide graph.
pub(crate) fn validate_detached_graph(original: &Graph, value: &Graph) -> Result<()> {
    if value.slides != original.slides {
        return Err(invalid(
            "presentation slide ordering and relationship topology are immutable in a structure transaction",
        ));
    }
    validate_graph_shape(value)
}

pub(crate) fn equivalent_graph(left: &Graph, right: &Graph) -> bool {
    left.slides == right.slides
        && left.custom_shows.shows == right.custom_shows.shows
        && left.sections == right.sections
}

pub fn find_custom_show(package: &OpcPackage, id: u32) -> Result<Option<Show>> {
    Ok(load(package)?.custom_shows.get_by_id(id).cloned())
}

pub fn add_custom_show(package: &mut OpcPackage, show: Show) -> Result<()> {
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

pub fn update_custom_show(package: &mut OpcPackage, id: u32, replacement: Show) -> Result<()> {
    mutate(package, |graph| {
        graph.custom_shows.replace_by_id(id, replacement)
    })
}

pub fn replace_custom_show(package: &mut OpcPackage, id: u32, replacement: Show) -> Result<()> {
    update_custom_show(package, id, replacement)
}

pub fn remove_custom_show(package: &mut OpcPackage, id: u32) -> Result<bool> {
    let mut graph = load(package)?;
    let removed = graph.custom_shows.remove_by_id(id).is_some();
    if removed {
        store(package, &graph)?;
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
    let mut graph = load(package)?;
    let show = graph
        .custom_shows
        .get_by_id_mut(show_id)
        .ok_or_else(|| invalid(format!("custom show {show_id} was not found")))?;
    let Some(offset) = show.slide_ids.iter().position(|id| *id == slide_id) else {
        return Ok(false);
    };
    show.slide_ids.remove(offset);
    store(package, &graph)?;
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
    Ok(load(package)?.sections.get_by_id(id).cloned())
}

pub fn add_section(package: &mut OpcPackage, mut section: Section) -> Result<String> {
    if section.id.is_none() {
        section.id = Some(new_guid());
    }
    let id = section.id.clone().expect("section ID was assigned");
    let retained = id.clone();
    mutate(package, move |graph| {
        if graph.sections.get_by_id(&id).is_some() {
            return Err(invalid(format!("section {id} already exists")));
        }
        graph.sections.add_section(section);
        graph.sections.sort_slide_membership(
            &graph
                .slides
                .iter()
                .map(|slide| slide.slide_id)
                .collect::<Vec<_>>(),
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
            &graph
                .slides
                .iter()
                .map(|slide| slide.slide_id)
                .collect::<Vec<_>>(),
        );
        Ok(())
    })
}

pub fn replace_section(package: &mut OpcPackage, id: &str, replacement: Section) -> Result<()> {
    update_section(package, id, replacement)
}

pub fn remove_section(package: &mut OpcPackage, id: &str) -> Result<bool> {
    let mut graph = load(package)?;
    let removed = graph.sections.remove_by_id(id).is_some();
    if removed {
        store(package, &graph)?;
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
            &graph
                .slides
                .iter()
                .map(|slide| slide.slide_id)
                .collect::<Vec<_>>(),
        );
        Ok(())
    })
}

pub fn remove_section_slide(
    package: &mut OpcPackage,
    section_id: &str,
    slide_id: u32,
) -> Result<bool> {
    let mut graph = load(package)?;
    let section = graph
        .sections
        .get_by_id_mut(section_id)
        .ok_or_else(|| invalid(format!("section {section_id} was not found")))?;
    let Some(offset) = section.slide_ids.iter().position(|id| *id == slide_id) else {
        return Ok(false);
    };
    section.slide_ids.remove(offset);
    store(package, &graph)?;
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
pub fn synchronize_after_slide_mutation(package: &mut OpcPackage) -> Result<()> {
    let presentation = package.main_document_part()?;
    let mut graph = parse_structure_blob(package, presentation.blob(), false)?;
    let live = graph
        .slides
        .iter()
        .map(|slide| slide.slide_id)
        .collect::<HashSet<_>>();
    for show in &mut graph.custom_shows.shows {
        show.slide_ids.retain(|id| live.contains(id));
    }
    for section in graph.sections.sections_mut() {
        section.slide_ids.retain(|id| live.contains(id));
    }
    graph.sections.sort_slide_membership(
        &graph
            .slides
            .iter()
            .map(|slide| slide.slide_id)
            .collect::<Vec<_>>(),
    );
    store(package, &graph)
}

fn mutate<F>(package: &mut OpcPackage, operation: F) -> Result<()>
where
    F: FnOnce(&mut Graph) -> Result<()>,
{
    let mut graph = load(package)?;
    operation(&mut graph)?;
    store(package, &graph)
}

fn parse_structure_blob(
    package: &OpcPackage,
    xml: &[u8],
    strict_references: bool,
) -> Result<Graph> {
    if xml.len() > MAX_BYTES {
        return Err(invalid("presentation structure exceeds 8 MiB"));
    }
    let processed = litchi_ooxml_common::mce::process_ooxml(xml)?;
    let (raw_slides, raw_shows) = parse_core(processed.as_ref())?;
    let presentation = package.main_document_part()?;
    let mut slides = Vec::with_capacity(raw_slides.len());
    for (slide_id, relationship_id) in raw_slides {
        let relationship = presentation.rels().get(&relationship_id).ok_or_else(|| {
            invalid(format!(
                "orphan presentation slide relationship {relationship_id}"
            ))
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
            return Err(invalid(format!(
                "relationship {relationship_id} targets a non-slide part"
            )));
        }
        slides.push(Reference {
            slide_id,
            relationship_id,
            part_name: target.to_string(),
        });
    }
    let custom_shows = resolve_custom_shows(raw_shows, &slides, strict_references)?;
    let graph = Graph {
        slides,
        custom_shows,
        sections: SectionList::from_xml(xml)?,
    };
    if strict_references {
        validate_graph(package, &graph)?;
    }
    Ok(graph)
}

fn resolve_custom_shows(
    raw_shows: Vec<RawShow>,
    slides: &[Reference],
    strict_references: bool,
) -> Result<ShowList> {
    let rel_to_id = slides
        .iter()
        .map(|slide| (slide.relationship_id.as_str(), slide.slide_id))
        .collect::<HashMap<_, _>>();
    let mut custom_shows = ShowList::new();
    for raw in raw_shows {
        let mut show = Show::new(raw.id, raw.name);
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
    Ok(custom_shows)
}

#[derive(Default)]
pub(super) struct RawShow {
    pub(super) id: u32,
    pub(super) name: String,
    pub(super) relationship_ids: Vec<String>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
enum CoreElement {
    Presentation,
    SlideIdList,
    SlideId,
    CustomShowList,
    CustomShow,
    SlideList,
    Slide,
    Other,
}

#[derive(Debug)]
struct Attribute {
    namespace: Option<Vec<u8>>,
    local: Vec<u8>,
    value: String,
}

#[allow(clippy::type_complexity)]
pub(super) fn parse_core(xml: &[u8]) -> Result<(Vec<(u32, String)>, Vec<RawShow>)> {
    let mut reader = NsReader::from_reader(xml);
    reader.config_mut().trim_text(true);
    let mut ancestors = Vec::<CoreElement>::new();
    let mut slides = Vec::new();
    let mut shows = Vec::new();
    let mut current_show: Option<RawShow> = None;
    let mut nodes = 0usize;
    loop {
        let decoder = reader.decoder();
        let event = reader
            .read_event()
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        let resolver = reader.resolver().clone();
        let (namespace, event) = resolver.resolve_event(event);
        match event {
            Event::Start(element) => {
                nodes += 1;
                resource(nodes, ancestors.len())?;
                let kind = core_element(&namespace, element.name())?;
                if is_direct_slide_reference(&ancestors, kind) {
                    slides.push(parse_slide_reference(&element, decoder, &reader)?);
                } else if kind == CoreElement::CustomShow {
                    if ancestors.as_slice()
                        != [CoreElement::Presentation, CoreElement::CustomShowList]
                    {
                        return Err(invalid("custom show has invalid presentation ancestry"));
                    }
                    if current_show.is_some() {
                        return Err(invalid("nested custom shows are rejected"));
                    }
                    current_show = Some(parse_show(&element, decoder, &reader)?);
                }
                ancestors.push(kind);
            },
            Event::Empty(element) => {
                nodes += 1;
                resource(nodes, ancestors.len())?;
                let kind = core_element(&namespace, element.name())?;
                if is_direct_slide_reference(&ancestors, kind) {
                    slides.push(parse_slide_reference(&element, decoder, &reader)?);
                } else if kind == CoreElement::Slide
                    && ancestors.as_slice()
                        == [
                            CoreElement::Presentation,
                            CoreElement::CustomShowList,
                            CoreElement::CustomShow,
                            CoreElement::SlideList,
                        ]
                {
                    let relationship_id = relationship_id(&element, decoder, &reader)?;
                    current_show
                        .as_mut()
                        .ok_or_else(|| invalid("custom-show slide appears outside a custom show"))?
                        .relationship_ids
                        .push(relationship_id);
                }
            },
            Event::End(element) => {
                let kind = core_element(&namespace, element.name())?;
                let open = ancestors
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
                if open != kind {
                    return Err(invalid("mismatched presentation element"));
                }
                if kind == CoreElement::CustomShow {
                    shows.push(
                        current_show
                            .take()
                            .ok_or_else(|| invalid("missing custom show"))?,
                    );
                }
            },
            Event::DocType(_) | Event::PI(_) => {
                return Err(invalid("DTDs and processing instructions are rejected"));
            },
            Event::Eof => break,
            _ => {},
        }
    }
    if !ancestors.is_empty() || current_show.is_some() {
        return Err(invalid("unterminated presentation structure"));
    }
    Ok((slides, shows))
}

fn core_element(namespace: &ResolveResult<'_>, name: QName<'_>) -> Result<CoreElement> {
    let presentationml = match namespace {
        ResolveResult::Bound(Namespace(value)) => {
            matches!(*value, value if value == P.as_bytes() || value == PS.as_bytes())
        },
        ResolveResult::Unbound => false,
        ResolveResult::Unknown(prefix) => {
            return Err(invalid(format!(
                "unresolved presentation XML namespace prefix '{}'",
                String::from_utf8_lossy(prefix.as_ref())
            )));
        },
    };
    if !presentationml {
        return Ok(CoreElement::Other);
    }
    Ok(match name.local_name().as_ref() {
        b"presentation" => CoreElement::Presentation,
        b"sldIdLst" => CoreElement::SlideIdList,
        b"sldId" => CoreElement::SlideId,
        b"custShowLst" => CoreElement::CustomShowList,
        b"custShow" => CoreElement::CustomShow,
        b"sldLst" => CoreElement::SlideList,
        b"sld" => CoreElement::Slide,
        _ => CoreElement::Other,
    })
}

fn is_direct_slide_reference(ancestors: &[CoreElement], kind: CoreElement) -> bool {
    kind == CoreElement::SlideId
        && ancestors == [CoreElement::Presentation, CoreElement::SlideIdList]
}

fn parse_show(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    reader: &NsReader<&[u8]>,
) -> Result<RawShow> {
    let attributes = attributes_ns(element, decoder, reader)?;
    let id = required(&attributes, "id")?
        .parse::<u32>()
        .map_err(|_| invalid("invalid custom-show ID"))?;
    let name = required(&attributes, "name")?.to_owned();
    if name.is_empty() {
        return Err(invalid("custom-show name cannot be empty"));
    }
    Ok(RawShow {
        id,
        name,
        relationship_ids: Vec::new(),
    })
}

fn parse_slide_reference(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    reader: &NsReader<&[u8]>,
) -> Result<(u32, String)> {
    let attributes = attributes_ns(element, decoder, reader)?;
    let id = required_unqualified(&attributes, "id")?
        .parse::<u32>()
        .map_err(|_| invalid("invalid presentation slide ID"))?;
    if !(256..=MAX_SLIDE_ID).contains(&id) {
        return Err(invalid("presentation slide ID is outside 256..=2147483647"));
    }
    Ok((id, required_qualified(&attributes, "id")?.to_owned()))
}

fn relationship_id(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    reader: &NsReader<&[u8]>,
) -> Result<String> {
    let attributes = attributes_ns(element, decoder, reader)?;
    Ok(required_qualified(&attributes, "id")?.to_owned())
}

fn attributes_ns(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
    reader: &NsReader<&[u8]>,
) -> Result<Vec<Attribute>> {
    let mut values = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let (namespace, local) = reader.resolver().resolve_attribute(attribute.key);
        let namespace = match namespace {
            ResolveResult::Bound(Namespace(value)) => Some(value.to_vec()),
            ResolveResult::Unbound => None,
            ResolveResult::Unknown(prefix) => {
                return Err(invalid(format!(
                    "unresolved presentation XML attribute prefix '{}'",
                    String::from_utf8_lossy(prefix.as_ref())
                )));
            },
        };
        let local = local.as_ref().to_vec();
        if !seen.insert((namespace.clone(), local.clone())) {
            return Err(invalid("duplicate presentation XML attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        values.push(Attribute {
            namespace,
            local,
            value,
        });
    }
    Ok(values)
}

fn attributes(
    element: &BytesStart<'_>,
    decoder: quick_xml::encoding::Decoder,
) -> Result<Vec<(String, String)>> {
    let mut values = Vec::new();
    let mut seen = HashSet::new();
    for attribute in element.attributes().with_checks(true) {
        let attribute = attribute.map_err(|error| Error::Xml(error.to_string()))?;
        let name = std::str::from_utf8(attribute.key.as_ref())
            .map_err(|error| Error::Xml(error.to_string()))?
            .to_owned();
        if !seen.insert(name.clone()) {
            return Err(invalid("duplicate presentation XML attribute"));
        }
        let value = attribute
            .decoded_and_normalized_value(XmlVersion::Implicit1_0, decoder)
            .map_err(|error| Error::Xml(error.to_string()))?
            .into_owned();
        values.push((name, value));
    }
    Ok(values)
}

fn required<'a>(values: &'a [Attribute], local: &str) -> Result<&'a str> {
    values
        .iter()
        .find(|attribute| attribute.namespace.is_none() && attribute.local == local.as_bytes())
        .map(|attribute| attribute.value.as_str())
        .ok_or_else(|| invalid(format!("missing attribute '{local}'")))
}

fn required_unqualified<'a>(values: &'a [Attribute], local: &str) -> Result<&'a str> {
    required(values, local)
}

fn required_qualified<'a>(values: &'a [Attribute], local: &str) -> Result<&'a str> {
    values
        .iter()
        .find(|attribute| {
            attribute.local == local.as_bytes()
                && attribute.namespace.as_deref().is_some_and(|namespace| {
                    namespace == R.as_bytes() || namespace == RS.as_bytes()
                })
        })
        .map(|attribute| attribute.value.as_str())
        .ok_or_else(|| invalid(format!("missing relationship attribute '{local}'")))
}

fn validate_graph(package: &OpcPackage, graph: &Graph) -> Result<()> {
    validate_graph_shape(graph)?;

    let presentation = package.main_document_part()?;
    for slide in &graph.slides {
        let relationship = presentation
            .rels()
            .get(&slide.relationship_id)
            .ok_or_else(|| {
                invalid(format!(
                    "orphan slide relationship {}",
                    slide.relationship_id
                ))
            })?;
        if relationship.is_external()
            || !matches!(relationship.reltype(), SLIDE_REL | SLIDE_REL_STRICT)
            || relationship.target_partname()?.as_str() != slide.part_name
        {
            return Err(invalid("presentation slide relationship mismatch"));
        }
    }
    Ok(())
}

fn validate_graph_shape(graph: &Graph) -> Result<()> {
    let mut slide_ids = HashSet::new();
    let mut rel_ids = HashSet::new();
    let mut part_names = HashSet::new();
    for slide in &graph.slides {
        if !(256..=MAX_SLIDE_ID).contains(&slide.slide_id) || !slide_ids.insert(slide.slide_id) {
            return Err(invalid("invalid or duplicate presentation slide ID"));
        }
        if !rel_ids.insert(slide.relationship_id.as_str())
            || !part_names.insert(slide.part_name.as_str())
        {
            return Err(invalid(
                "duplicate presentation slide relationship or target",
            ));
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
    shows: &ShowList,
    slides: &[Reference],
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
    Ok(inner.replacen("<p:ext ", &format!("<p:ext xmlns:p=\"{p_namespace}\" "), 1))
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
                let frame = stack
                    .pop()
                    .ok_or_else(|| invalid("unexpected closing element"))?;
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
            Err(error) => return Err(Error::Xml(error.to_string())),
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
                std::str::from_utf8(replacement).map_err(|error| Error::Xml(error.to_string()))?
            );
            return replace_spans(xml, std::slice::from_ref(ext_list), wrapped.as_bytes());
        }
        return insert_bytes(xml, ext_list.close_start, replacement);
    }
    let wrapped = format!(
        "<p:extLst xmlns:p=\"{p_namespace}\">{}</p:extLst>",
        std::str::from_utf8(replacement).map_err(|error| Error::Xml(error.to_string()))?
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
        "sldMasterIdLst",
        "notesMasterIdLst",
        "handoutMasterIdLst",
        "sldIdLst",
        "sldSz",
        "notesSz",
        "smartTags",
        "embeddedFontLst",
        "custShowLst",
        "photoAlbum",
        "custDataLst",
        "kinsoku",
        "defaultTextStyle",
        "modifyVerifier",
        "extLst",
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

fn require_slide(graph: &Graph, slide_id: u32) -> Result<()> {
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
        Err(invalid(
            "presentation structure XML resource limit exceeded",
        ))
    } else {
        Ok(())
    }
}

fn local_name(name: &[u8]) -> Result<String> {
    let name = std::str::from_utf8(name).map_err(|error| Error::Xml(error.to_string()))?;
    Ok(name
        .rsplit_once(':')
        .map_or(name, |(_, local)| local)
        .to_owned())
}

fn invalid(message: impl Into<String>) -> Error {
    Error::Invalid(message.into())
}

#[cfg(test)]
mod tests {
    use super::*;

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
        assert!(
            patched
                .windows(b"<!--keep-->".len())
                .any(|window| window == b"<!--keep-->")
        );
    }

    #[test]
    fn non_empty_slide_ids_are_recorded_once_and_indirect_lookalikes_are_ignored() {
        let xml = br#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:x="urn:future"><p:sldIdLst><p:sldId id="256" r:id="rIdOne"><p:extLst><p:ext uri="{opaque}"><x:future><p:sldId id="999" r:id="rIdNested"/></x:future></p:ext></p:extLst></p:sldId><p:extLst><p:sldId id="998" r:id="rIdIndirect"/></p:extLst><p:sldId id="257" r:id="rIdTwo"/></p:sldIdLst><p:extLst><p:sldId id="997" r:id="rIdOutside"/></p:extLst></p:presentation>"#;

        let (slides, shows) = parse_core(xml).unwrap();

        assert_eq!(
            slides,
            vec![(256, "rIdOne".to_owned()), (257, "rIdTwo".to_owned())]
        );
        assert!(shows.is_empty());
    }
}
